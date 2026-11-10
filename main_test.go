package main

import (
	"encoding/base64"
	"io"
	"net/http"
	"net/http/httptest"
	"net/url"
	"os"
	"strconv"
	"strings"
	"testing"
)

func TestDownloadURLLocal(t *testing.T) {
	tests := []struct {
		in, want string
		wantErr  bool
	}{
		// SharePoint link -> same URL with download=1 appended
		{
			in:   "https://contoso.sharepoint.com/:u:/g/personal/user/EWEXAMPLE?e=4:abc123",
			want: "https://contoso.sharepoint.com/:u:/g/personal/user/EWEXAMPLE?e=4:abc123&download=1",
		},
		// SharePoint link without a query -> ?download=1
		{
			in:   "https://contoso.sharepoint.com/:u:/g/personal/user/EWEXAMPLE",
			want: "https://contoso.sharepoint.com/:u:/g/personal/user/EWEXAMPLE?download=1",
		},
		// Already a direct download URL -> unchanged
		{
			in:   "https://api.onedrive.com/v1.0/shares/u!QUJD/root/content",
			want: "https://api.onedrive.com/v1.0/shares/u!QUJD/root/content",
		},
		{in: "ftp://example.com/x", wantErr: true},
		{in: "://nope", wantErr: true},
	}
	for _, tt := range tests {
		got, _, err := downloadURL(&http.Client{}, tt.in)
		if tt.wantErr {
			if err == nil {
				t.Errorf("downloadURL(%q) = %q, want error", tt.in, got)
			}
			continue
		}
		if err != nil {
			t.Errorf("downloadURL(%q) error: %v", tt.in, err)
			continue
		}
		if got != tt.want {
			t.Errorf("downloadURL(%q)\n got  %s\n want %s", tt.in, got, tt.want)
		}
	}
}

func TestDownloadURLRedeemed(t *testing.T) {
	// Links whose owner account was migrated to SharePoint carry a "redeem"
	// code: an anonymous badger token unlocks the personal-content API.
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.URL.Path == "/token":
			if r.Header.Get("AppId") != badgerAppID {
				t.Errorf("token request AppId header = %q", r.Header.Get("AppId"))
			}
			body, _ := io.ReadAll(r.Body)
			if !strings.Contains(string(body), badgerAppUUID) {
				t.Errorf("token request body = %q, want it to contain the app UUID", body)
			}
			io.WriteString(w, `{"token":"TOK"}`)
		case strings.HasPrefix(r.URL.Path, "/v2.0/shares/u!"):
			if auth := r.Header.Get("Authorization"); auth != "Badger TOK" {
				t.Errorf("Authorization header = %q, want Badger TOK", auth)
			}
			if p := r.Header.Get("Prefer"); p != "autoredeem" {
				t.Errorf("Prefer header = %q, want autoredeem", p)
			}
			io.WriteString(w, `{"name":"report.pdf","@content.downloadUrl":"https://dl.example.com/tempauth?e=1"}`)
		default:
			t.Errorf("unexpected request path %q", r.URL.Path)
			w.WriteHeader(http.StatusNotFound)
		}
	}))
	defer srv.Close()
	overrideAPIs(t, srv)

	in := "https://onedrive.live.com/?redeem=" + url.QueryEscape(b64("https://1drv.ms/u/c/abc/IQY?e=z")) + "&resid=ABC%2142&cid=ABC"
	dl, name, err := downloadURL(srv.Client(), in)
	if err != nil {
		t.Fatalf("downloadURL: %v", err)
	}
	if dl != "https://dl.example.com/tempauth?e=1" {
		t.Errorf("downloadURL dl = %q", dl)
	}
	if name != "report.pdf" {
		t.Errorf("downloadURL name = %q, want report.pdf", name)
	}
}

func TestDownloadURLLegacyAuthkey(t *testing.T) {
	// Classic flow: expanded link with cid/resid/authkey resolves through the
	// anonymous drive-item API.
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/v1.0/drives/ABC/items/ABC!42" {
			t.Errorf("request path = %q", r.URL.Path)
		}
		if got := r.URL.Query().Get("authkey"); got != "!AKE" {
			t.Errorf("authkey query = %q, want !AKE", got)
		}
		io.WriteString(w, `{"name":"pic.jpg","@content.downloadUrl":"https://dl.example.com/x?e=2"}`)
	}))
	defer srv.Close()
	overrideAPIs(t, srv)

	dl, name, err := downloadURL(srv.Client(), "https://onedrive.live.com/?resid=ABC!42&cid=ABC&authkey=!AKE")
	if err != nil {
		t.Fatalf("downloadURL: %v", err)
	}
	if dl != "https://dl.example.com/x?e=2" || name != "pic.jpg" {
		t.Errorf("downloadURL = (%q, %q)", dl, name)
	}
}

func TestDownloadURLFallbackSharesAPI(t *testing.T) {
	// Expanded link without resid/redeem falls back to the base64-encoded
	// anonymous shares API.
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if strings.HasPrefix(r.URL.Path, "/v1.0/shares/u!") {
			io.WriteString(w, `{"name":"doc.txt","@content.downloadUrl":"https://dl.example.com/y?e=3"}`)
			return
		}
		// Anything else is the link being expanded.
		if r.URL.Path != "/share/UQEXAMPLE" {
			t.Errorf("unexpected request path %q", r.URL.Path)
		}
	}))
	defer srv.Close()
	overrideAPIs(t, srv)

	in := srv.URL + "/share/UQEXAMPLE?e=abc"
	dl, name, err := downloadURL(srv.Client(), in)
	if err != nil {
		t.Fatalf("downloadURL: %v", err)
	}
	if dl != "https://dl.example.com/y?e=3" || name != "doc.txt" {
		t.Errorf("downloadURL = (%q, %q)", dl, name)
	}
}

func overrideAPIs(t *testing.T, srv *httptest.Server) {
	oldTokenURL, oldLegacy, oldPersonal := badgerTokenURL, legacyAPIBase, personalAPIBase
	badgerTokenURL = srv.URL + "/token"
	legacyAPIBase = srv.URL + "/v1.0"
	personalAPIBase = srv.URL + "/v2.0"
	t.Cleanup(func() {
		badgerTokenURL, legacyAPIBase, personalAPIBase = oldTokenURL, oldLegacy, oldPersonal
	})
}

func TestExpand(t *testing.T) {
	final := ""
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch r.URL.Path {
		case "/first":
			http.Redirect(w, r, "/second?resid=ABC!1", http.StatusMovedPermanently)
		case "/second":
			final = r.URL.Query().Get("resid")
		default:
			t.Errorf("unexpected request path %q", r.URL.Path)
			w.WriteHeader(http.StatusNotFound)
		}
	}))
	defer srv.Close()

	u, err := expand(srv.Client(), srv.URL+"/first")
	if err != nil {
		t.Fatalf("expand: %v", err)
	}
	if final != "ABC!1" {
		t.Errorf("expand followed redirects to %q (wanted resid %q)", u, final)
	}
}

func b64(s string) string {
	return base64.RawURLEncoding.EncodeToString([]byte(s))
}

func TestSave(t *testing.T) {
	const data = "hello one drive" // 15 bytes
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Disposition", `attachment; filename="greeting.txt"`)
		w.Header().Set("Content-Length", strconv.Itoa(len(data)))
		io.WriteString(w, data)
	}))
	defer srv.Close()

	tmp := t.TempDir()
	t.Chdir(tmp)
	if err := save(strings.NewReader(data), "greeting.txt", int64(len(data))); err != nil {
		t.Fatalf("save: %v", err)
	}
	got, err := os.ReadFile("greeting.txt")
	if err != nil {
		t.Fatalf("read: %v", err)
	}
	if string(got) != data {
		t.Errorf("file contents = %q, want %q", got, data)
	}
	// Temp files must not linger next to the download.
	entries, _ := os.ReadDir(".")
	for _, e := range entries {
		if strings.HasPrefix(e.Name(), ".onedrivedl-") {
			t.Errorf("temp file %q left behind", e.Name())
		}
	}
}

func TestFilenameFrom(t *testing.T) {
	mkURL := func(s string) *http.Request {
		u, _ := url.Parse(s)
		r, _ := http.NewRequest("GET", u.String(), nil)
		r.URL = u
		return r
	}

	tests := []struct {
		name string
		resp *http.Response
		want string
	}{
		{
			name: "content-disposition simple",
			resp: &http.Response{Header: http.Header{"Content-Disposition": {`attachment; filename="report.pdf"`}}},
			want: "report.pdf",
		},
		{
			name: "content-disposition rfc5987",
			resp: &http.Response{Header: http.Header{"Content-Disposition": {`attachment; filename*=UTF-8''My%20Report%20%E2%98%95%20v2.pdf`}}},
			want: "My Report ☕ v2.pdf",
		},
		{
			name: "final URL path",
			resp: &http.Response{Header: http.Header{}, Request: mkURL("https://example.com/files/holiday.jpg")},
			want: "holiday.jpg",
		},
		{
			name: "badger content path falls back to download",
			resp: &http.Response{Header: http.Header{}, Request: mkURL("https://api.onedrive.com/v1.0/shares/u!x/root/content")},
			want: "download",
		},
		{
			name: "nothing to go on",
			resp: &http.Response{Header: http.Header{}, Request: mkURL("https://example.com/")},
			want: "download",
		},
		{
			name: "path traversal sanitized",
			resp: &http.Response{Header: http.Header{"Content-Disposition": {`attachment; filename="../../evil"`}}},
			want: ".._.._evil",
		},
	}
	for _, tt := range tests {
		if got := filenameFrom(tt.resp); got != tt.want {
			t.Errorf("%s: filenameFrom() = %q, want %q", tt.name, got, tt.want)
		}
	}
}
