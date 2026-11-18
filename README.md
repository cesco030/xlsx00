# XLSX tools for IBM i - packages

Each folder holds one version, built on PUB400.COM (a public IBM i) from the commit named in
`COMMIT` of branch `main`:

- `xlsx-<version>-objects.savf` - the 6 objects (XLSX00, XLSX01 and their programs and help), saved from library `CESCO2`
- `xlsx-<version>-sources.savf` - the sources as a directory save
- `README.md` - install, update, removal, build from the sources
- `XLSX00-help.html`, `XLSX01-help.html` - the command help
- `LICENSE`, `THIRD-PARTY-NOTICES.txt`, `build.log`, `COMMIT`, `SHA256SUMS`

Download the save file in binary (the raw file of GitHub, or `git clone -b packages`), then follow the README of its folder.
