// Command onedrivedl downloads a file from OneDrive given its shared URL.
//
// Usage:
//
//	onedrivedl <share-url> [-o output-file]
package main

import (
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"log"
	"mime"
	"net/http"
	"net/url"
	"os"
	"path"
	"path/filepath"
	"strings"
	"time"
)

const version = "0.2.0"

const userAgent = "onedrivedl/" + version

// Endpoints for link resolution. Variables so tests can point them at a
// local server.
var (
	badgerTokenURL  = "https://api-badgerp.svc.ms/v1.0/token"
	legacyAPIBase   = "https://api.onedrive.com/v1.0"
	personalAPIBase = "https://my.microsoftpersonalcontent.com/_api/v2.0"
)

// App identity OneDrive's web client uses for anonymous browser sessions.
const (
	badgerAppID   = "1141147648"
	badgerAppUUID = "5cbed6ac-a083-4e14-b191-b4ba07653de2"
)

const usageText = `usage: onedrivedl <share-url> [-o output-file]

Downloads a shared OneDrive / SharePoint file, streaming it to disk.

  -o, --output  output file name
                (default: derived from the response)
`

func usage(w io.Writer) { fmt.Fprint(w, usageText) }

func main() {
	log.SetFlags(0)

	shareURL, out, err := parseArgs(os.Args[1:])
	if err != nil {
		fmt.Fprintf(os.Stderr, "onedrivedl: %v\n\n", err)
		usage(os.Stderr)
		os.Exit(2)
	}
	if shareURL == "" {
		usage(os.Stderr)
		os.Exit(2)
	}

	if err := run(shareURL, out); err != nil {
		log.Fatalf("onedrivedl: %v", err)
	}
}

// parseArgs accepts -o/--output (flag-style or attached) before or after
// the positional share URL, so both "<url> -o file" and "-o file <url>" work.
func parseArgs(argv []string) (shareURL, out string, err error) {
	var positional []string
	for i := 0; i < len(argv); i++ {
		a := argv[i]
		switch {
		case a == "--":
			positional = append(positional, argv[i+1:]...)
			i = len(argv)
		case a == "-h" || a == "--help":
			usage(os.Stdout)
			os.Exit(0)
		case a == "-o" || a == "--output":
			if i++; i >= len(argv) {
				return "", "", fmt.Errorf("flag %q needs a file name", "-o")
			}
			out = argv[i]
		case strings.HasPrefix(a, "-o="):
			out = strings.TrimPrefix(a, "-o=")
		case strings.HasPrefix(a, "--output="):
			out = strings.TrimPrefix(a, "--output=")
		case strings.HasPrefix(a, "-o"):
			out = strings.TrimPrefix(a, "-o")
		case strings.HasPrefix(a, "-"):
			return "", "", fmt.Errorf("unknown flag %q", a)
		default:
			positional = append(positional, a)
		}
	}
	if len(positional) > 1 {
		return "", "", fmt.Errorf("unexpected argument %q", positional[1])
	}
	if len(positional) == 1 {
		shareURL = positional[0]
	}
	return shareURL, out, nil
}

func run(shareURL, out string) error {
	client := &http.Client{Timeout: 5 * time.Minute}

	dl, hint, err := downloadURL(client, shareURL)
	if err != nil {
		return err
	}

	req, err := http.NewRequest(http.MethodGet, dl, nil)
	if err != nil {
		return fmt.Errorf("bad URL: %w", err)
	}
	req.Header.Set("User-Agent", userAgent)

	resp, err := client.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(io.LimitReader(resp.Body, 8<<10))
		if resp.StatusCode == http.StatusNotFound {
			return errors.New("link not found or no longer shared (404)")
		}
		if resp.StatusCode == http.StatusTooManyRequests {
			return errors.New("rate limited; try again later (429)")
		}
		if resp.StatusCode == http.StatusUnauthorized {
			return fmt.Errorf("download failed: %s: %s (the share may require sign-in)",
				resp.Status, strings.TrimSpace(string(body)))
		}
		return fmt.Errorf("download failed: %s: %s", resp.Status, strings.TrimSpace(string(body)))
	}

	name := out
	if name == "" {
		name = hint // file name reported by the resolve step, if any
	}
	if name == "" {
		name = filenameFrom(resp) // Content-Disposition, then final URL path
	}

	if err := save(resp.Body, name, resp.ContentLength); err != nil {
		os.Remove(name)
		return err
	}

	size := resp.ContentLength
	if size < 0 {
		size = fileBytes(name)
	}
	fmt.Fprintf(os.Stderr, "Downloaded %s (%s)\n", name, humanBytes(size))
	return nil
}

// downloadURL converts a OneDrive / SharePoint share URL into a direct
// download URL, returning the URL and, if known, its file name.
//
// Links short-links (1drv.ms) are expanded first. Accounts whose OneDrive was
// migrated to SharePoint ("redeem" in the expanded link) no longer allow
// anonymous access via the classic shares API; those are resolved with an
// anonymous browser "badger" token against the personal-content API.
func downloadURL(client *http.Client, shareURL string) (dl, name string, err error) {
	trimmed := strings.TrimSpace(shareURL)
	u, err := url.Parse(trimmed)
	if err != nil {
		return "", "", fmt.Errorf("invalid URL: %w", err)
	}
	switch u.Scheme {
	case "http", "https":
	default:
		return "", "", fmt.Errorf("invalid URL scheme %q", u.Scheme)
	}

	host := strings.ToLower(u.Hostname())
	switch {
	case strings.HasSuffix(host, "sharepoint.com"):
		// SharePoint share links become direct downloads with ?download=1.
		return appendQuery(u, "download=1"), "", nil
	case host == "api.onedrive.com" && strings.Contains(u.Path, "/root/content"):
		return shareURL, "", nil // already a direct download URL
	}

	// Short links (1drv.ms, ...) redirect to onedrive.live.com with the
	// item's resid, cid, authkey or redeem parameters. onedrive.live.com
	// links already carry them.
	if host != "onedrive.live.com" {
		expanded, err := expand(client, shareURL)
		if err != nil {
			return "", "", fmt.Errorf("expanding share link: %w", err)
		}
		u = expanded
	}

	q := u.Query()
	redeem := q.Get("redeem")
	resid := q.Get("resid")
	if resid == "" {
		// Older pages carry the id in "id" (e.g. ABCXXX!12345).
		if id := q.Get("id"); strings.Contains(id, "!") {
			resid = id
		}
	}
	authkey := q.Get("authkey")
	cid := q.Get("cid")
	if cid == "" {
		if _, suffix, ok := strings.Cut(resid, "!"); ok && suffix != "" {
			cid, _, _ = strings.Cut(resid, "!")
		}
	}

	// Links whose owner's account was migrated to SharePoint carry a
	// "redeem" code instead of an authkey and need anonymous "badger"
	// credentials.
	if redeem != "" {
		return redeemedDownloadURL(client, redeem)
	}

	// Classic flow: anonymous drive-item lookup by cid/resid, with optional
	// shared-authkey.
	if cid != "" && resid != "" {
		ep := fmt.Sprintf("%s/drives/%s/items/%s", legacyAPIBase,
			url.PathEscape(cid), url.PathEscape(resid))
		if authkey != "" {
			ep += "?authkey=" + url.QueryEscape(authkey)
		}
		name, err := fetchItem(client, ep, nil)
		if err != nil {
			return "", "", err
		}
		if name.DownloadURL == "" {
			return "", "", noDownloadURL
		}
		return name.DownloadURL, sanitize(name.Name), nil
	}

	// Fallback: the anonymous shares API for base64-encoded share URLs.
	b64 := base64.RawURLEncoding.EncodeToString([]byte(trimmed))
	meta, err := fetchItem(client, fmt.Sprintf("%s/shares/u!%s/root", legacyAPIBase, b64), nil)
	if err != nil {
		return "", "", err
	}
	if meta.DownloadURL == "" {
		return "", "", noDownloadURL
	}
	return meta.DownloadURL, sanitize(meta.Name), nil
}

var noDownloadURL = errors.New("share resolved but no download URL returned; it may require sign-in")

// expand follows a share link's redirects and returns the final URL.
func expand(client *http.Client, link string) (*url.URL, error) {
	resp, err := client.Get(link)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	// Redirect pages are tiny; drain up to 64 KiB so the connection can be
	// reused, and keep only the final URL.
	_, _ = io.Copy(io.Discard, io.LimitReader(resp.Body, 64<<10))
	return resp.Request.URL, nil
}

// redeemedDownloadURL resolves a "redeem" code to a direct download URL using
// the anonymous badger token flow.
func redeemedDownloadURL(client *http.Client, redeem string) (string, string, error) {
	token, err := badgerToken(client)
	if err != nil {
		return "", "", err
	}
	ep := fmt.Sprintf("%s/shares/u!%s/driveitem", personalAPIBase, url.PathEscape(redeem))
	meta, err := fetchItem(client, ep, map[string]string{
		"Authorization": "Badger " + token,
		"Prefer":        "autoredeem",
	})
	if err != nil {
		return "", "", err
	}
	if meta.DownloadURL == "" {
		return "", "", noDownloadURL
	}
	return meta.DownloadURL, sanitize(meta.Name), nil
}

// badgerToken fetches an anonymous browser session token used to authorize
// calls to the personal-content API without signing in.
func badgerToken(client *http.Client) (string, error) {
	body := `{"appId":"` + badgerAppUUID + `"}`
	req, err := http.NewRequest(http.MethodPost, badgerTokenURL, strings.NewReader(body))
	if err != nil {
		return "", err
	}
	req.Header.Set("Content-Type", "application/json")
	req.Header.Set("AppId", badgerAppID)
	req.Header.Set("User-Agent", userAgent)

	resp, err := client.Do(req)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()
	if resp.StatusCode != http.StatusOK {
		b, _ := io.ReadAll(io.LimitReader(resp.Body, 4<<10))
		return "", fmt.Errorf("anonymous session token request failed: %s: %s",
			resp.Status, strings.TrimSpace(string(b)))
	}
	var out struct {
		Token string `json:"token"`
	}
	if err := json.NewDecoder(io.LimitReader(resp.Body, 1<<16)).Decode(&out); err != nil {
		return "", fmt.Errorf("anonymous session token: %w", err)
	}
	if out.Token == "" {
		return "", errors.New("anonymous session token: empty token in response")
	}
	return out.Token, nil
}

// itemMeta carries the parts of a drive-item response we need: the file name
// and its time-limited direct download URL.
type itemMeta struct {
	Name        string `json:"name"`
	DownloadURL string `json:"@content.downloadUrl"`
}

// fetchItem GETs an API endpoint and parses the drive-item metadata.
func fetchItem(client *http.Client, ep string, headers map[string]string) (itemMeta, error) {
	req, err := http.NewRequest(http.MethodGet, ep, nil)
	if err != nil {
		return itemMeta{}, err
	}
	req.Header.Set("User-Agent", userAgent)
	for k, v := range headers {
		req.Header.Set(k, v)
	}

	resp, err := client.Do(req)
	if err != nil {
		return itemMeta{}, err
	}
	defer resp.Body.Close()
	if resp.StatusCode != http.StatusOK {
		b, _ := io.ReadAll(io.LimitReader(resp.Body, 4<<10))
		return itemMeta{}, fmt.Errorf("resolving share failed: %s: %s",
			resp.Status, strings.TrimSpace(string(b)))
	}
	var meta itemMeta
	if err := json.NewDecoder(io.LimitReader(resp.Body, 1<<20)).Decode(&meta); err != nil {
		return itemMeta{}, fmt.Errorf("resolving share: %w", err)
	}
	return meta, nil
}

func appendQuery(u *url.URL, kv string) string {
	sep := "?"
	if u.RawQuery != "" {
		sep = "&"
	}
	return shareURLFor(u) + sep + kv
}

func shareURLFor(u *url.URL) string {
	c := u.Clone() // Go 1.27: deep-copy instead of manual copy
	c.Fragment = ""
	return c.String()
}

// filenameFrom picks an output file name from the HTTP response:
// first the Content-Disposition header, then the final URL's path.
func filenameFrom(resp *http.Response) string {
	if cd := resp.Header.Get("Content-Disposition"); cd != "" {
		if _, params, err := mime.ParseMediaType(cd); err == nil {
			for _, k := range []string{"filename*", "filename"} {
				if v, ok := params[k]; ok && strings.TrimSpace(v) != "" {
					if k == "filename*" {
						if name, ok := decodeExtValue(v); ok && name != "" {
							return sanitize(name)
						}
						continue
					}
					return sanitize(v)
				}
			}
		}
	}
	if u := resp.Request.URL; u != nil {
		if base := path.Base(u.Path); base != "" && base != "/" && base != "." &&
			!strings.HasSuffix(u.Path, "/root/content") {
			return sanitize(base)
		}
	}
	return "download"
}

// decodeExtValue parses RFC 5987 filename*=UTF-8”Name.ext values.
func decodeExtValue(v string) (string, bool) {
	// charset'language'value
	if i := strings.IndexByte(v, '\''); i > 0 {
		if j := strings.IndexByte(v[i+1:], '\''); j >= 0 {
			charset := strings.ToLower(v[:i])
			raw := v[i+j+2:]
			decoded, err := url.PathUnescape(raw)
			if err != nil {
				return "", false
			}
			if charset != "" && charset != "utf-8" && charset != "us-ascii" {
				// Unknown charset: let the raw value through.
				return decoded, true
			}
			return decoded, true
		}
	}
	return "", false
}

func sanitize(name string) string {
	name = strings.TrimSpace(strings.Map(func(r rune) rune {
		switch r {
		case '/', '\\', ':', '\x00':
			return '_'
		}
		return r
	}, name))
	name = strings.Trim(name, "\"")
	if name == "" || name == "." || name == ".." {
		return "download"
	}
	return name
}

// save streams the body to name via a temp file, rendering progress on stderr.
func save(r io.Reader, name string, total int64) error {
	if name == "" {
		return errors.New("empty output file name")
	}
	// Stage the temp file next to the output so the final rename cannot fail
	// with a cross-device link error.
	dir, _ := filepath.Split(name)
	if dir == "" {
		dir = "."
	}
	tmp, err := os.CreateTemp(dir, ".onedrivedl-*")
	if err != nil {
		return err
	}
	tmpName := tmp.Name()
	defer os.Remove(tmpName) // no-op after successful rename

	pw := &progress{total: total, start: time.Now()}
	if _, err := io.Copy(io.MultiWriter(tmp, pw), r); err != nil {
		tmp.Close()
		return err
	}
	if err := tmp.Close(); err != nil {
		return err
	}
	fmt.Fprint(os.Stderr, "\r"+strings.Repeat(" ", 60)+"\r")

	if err := os.Rename(tmpName, name); err != nil {
		return err
	}
	return nil
}

type progress struct {
	total   int64
	written int64
	start   time.Time
	last    time.Time
}

func (p *progress) Write(b []byte) (int, error) {
	p.written += int64(len(b))
	if time.Since(p.last) >= 80*time.Millisecond {
		p.last = time.Now()
		p.render()
	}
	return len(b), nil
}

func (p *progress) render() {
	elapsed := time.Since(p.start).Seconds()
	speed := ""
	if elapsed > 0.5 {
		speed = fmt.Sprintf(" at %s/s", humanBytes(int64(float64(p.written)/elapsed)))
	}
	if p.total > 0 {
		fmt.Fprintf(os.Stderr, "\rDownloading %v of %v (%.0f%%)%s ",
			humanBytes(p.written), humanBytes(p.total), 100*float64(p.written)/float64(p.total), speed)
	} else {
		fmt.Fprintf(os.Stderr, "\rDownloading %s%s", humanBytes(p.written), speed)
	}
}

func humanBytes(n int64) string {
	switch {
	case n >= 1<<40:
		return fmt.Sprintf("%.2f TiB", float64(n)/(1<<40))
	case n >= 1<<30:
		return fmt.Sprintf("%.2f GiB", float64(n)/(1<<30))
	case n >= 1<<20:
		return fmt.Sprintf("%.2f MiB", float64(n)/(1<<20))
	case n >= 1<<10:
		return fmt.Sprintf("%.1f KiB", float64(n)/(1<<10))
	default:
		return fmt.Sprintf("%d B", n)
	}
}

func fileBytes(name string) int64 {
	fi, err := os.Stat(name)
	if err != nil {
		return 0
	}
	return fi.Size()
}
