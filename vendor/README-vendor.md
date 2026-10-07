# vendor/

Local copies of the two third-party libraries. The viewers load them from this
folder instead of a CDN, so opening a file makes no network request at all and
the exact code that runs is part of the repository.

| File | Library | Version | License | Official source |
|---|---|---|---|---|
| `papaparse.min.js` | [PapaParse](https://github.com/mholt/PapaParse) — CSV parser | 5.4.1 | MIT | npm package `papaparse@5.4.1` |
| `xlsx.full.min.js` | [SheetJS Community Edition](https://sheetjs.com/) — Excel reader/writer | 0.20.3 | Apache-2.0 | `https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js` |

## PapaParse

Taken from the npm package `papaparse@5.4.1` (published 2023-03-23), whose
tarball npm checks against the registry hash on download.

```
sha256  b8e870c5d2b29772f10c9fa9a693c8b896aac8540ed6701e3cc6304c683febdb  papaparse.min.js
```

Check on Linux/macOS with `sha256sum papaparse.min.js` (macOS: `shasum -a 256`),
on Windows with `certutil -hashfile papaparse.min.js SHA256`.

## SheetJS

Download it **only** from SheetJS's own CDN:

The minified file cannot be reviewed by reading it, so check where it comes
from instead: download the official **package**, compare its checksum with the
published one, and take the file out of the verified package.

**1. Download the package** `xlsx-0.20.3.tgz`:

```
curl -O https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz
```

**2. Compare the checksum.** It must match exactly:

Linux / macOS:

```
openssl dgst -sha512 -binary xlsx-0.20.3.tgz | openssl base64 -A
```

expected:

```
oLDq3jw7AcLqKWH2AhCpVTZl8mf6X2YReP+Neh0SJUzV/BdZYjth94tG5toiMB1PPrYtxOCfaoUCkvtuH+3AJA==
```

Windows:

```
certutil -hashfile xlsx-0.20.3.tgz SHA512
```

expected:

```
a0b0eade3c3b01c2ea2961f60210a9553665f267fa5f661178ff8d7a1d12254cd5fc1759623b61f78b46e6da22301d4f3eb62dc4e09f6a850292fb6e1fedc024
```

This is the `integrity` value npm records in `package-lock.json` when a project
installs SheetJS 0.20.3 from its CDN. The same value appears in the lockfiles
of many unrelated projects, for example
[pdobarco/bpo](https://github.com/pdobarco/bpo/blob/main/server/package-lock.json),
[pipeshub-ai](https://github.com/pipeshub-ai/pipeshub-ai/blob/main/frontend/package-lock.json) and
[civicpulse](https://github.com/datarhan/civicpulse/blob/main/package-lock.json),
so a tampered package would stand out.

**3. Take the one file you need out of the package.** Only
`package/dist/xlsx.full.min.js` is needed; everything else in the package
stays out of the repository.

Linux / macOS:

```
tar -xzf xlsx-0.20.3.tgz package/dist/xlsx.full.min.js
mv package/dist/xlsx.full.min.js vendor/
rm -r package xlsx-0.20.3.tgz
```

Windows:

```
tar -xzf xlsx-0.20.3.tgz package/dist/xlsx.full.min.js
move package\dist\xlsx.full.min.js vendor\
rmdir /s /q package
del xlsx-0.20.3.tgz
```

Afterwards this folder contains exactly three files:

```
vendor/
├── README-vendor.md
├── papaparse.min.js
└── xlsx.full.min.js
```

Do not use the `xlsx` package from npm, cdnjs, jsDelivr or unpkg: those stop
at 0.18.5, which is vulnerable to prototype pollution
([CVE-2023-30533](https://github.com/advisories/GHSA-4r6h-8v6p-xvw6), fixed in 0.19.3)
and ReDoS ([CVE-2024-22363](https://github.com/advisories/GHSA-5pgg-2g8v-p4x9),
fixed in 0.20.2) when reading crafted files. The Excel viewer shows a warning
on its start screen if it finds a version older than 0.20.2.

Security scanners often still flag 0.20.3 for these two advisories, because
the fixed versions were never published to npm. That is a false positive.
