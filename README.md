# onedrivedl

An utility to download a file from OneDrive using its shared URL.

```text
onedrivedl <share-url> [-o output-file]
```

## How it works

Share links are expanded and converted to direct download URLs:

| Link type | Strategy |
| --- | --- |
| SharePoint (`*.sharepoint.com`) | `download=1` appended to the share URL |
| Direct `api.onedrive.com .../root/content` URLs | Used as-is |
| Migrated accounts (expanded link carries `redeem=`) | Anonymous browser "badger" token (`api-badgerp.svc.ms/v1.0/token`), then `https://my.microsoftpersonalcontent.com/_api/v2.0/shares/u!<redeem>/driveitem` → its `@content.downloadUrl` |
| Classic accounts (expanded link carries `resid=`/`authkey=`) | `https://api.onedrive.com/v1.0/drives/<cid>/items/<resid>?authkey=...` → its `@content.downloadUrl` |
| Fallback (no `resid`/`redeem` in the expanded link) | Resolved via the classic shares API: `https://api.onedrive.com/v1.0/shares/u!<base64url>/root` |

Since around February 2025, personal OneDrive accounts migrated to a
SharePoint backend no longer authorize anonymous access through the classic
shares API (it fails with `UnauthenticatedVroomException`); such links are
detected by the `redeem` parameter and use the badger-token flow instead.
No sign-in is needed in either case: everything stays anonymous, though the
resolved download URL is time-limited (about an hour).

If `-o` is not given, the file name is taken from the resolve step's
metadata, falling back to the response's `Content-Disposition` header, the
final URL's path, or `download`.

## Example

```shell
$ onedrivedl 'https://1drv.ms/u/s!AbCdEfGhIjKlMnOp' -o report.pdf
Downloaded report.pdf (2.52 MiB)
```

`-o` can appear before or after the URL; `--output` and `-o=name` also work.

## Build

Requires Go 1.27+.

```shell
go build
go test ./...
```

## Notes

- The link must be shared with "anyone with the link" (no sign-in required).
- Sharing history/privacy links with an expired or revoked status return HTTP errors, which are reported as-is.
