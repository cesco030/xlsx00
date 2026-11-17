# XLSX tools 1.4.1 for IBM i - Excel .xlsx from and to Db2, in pure ILE RPG

Two commands, no Java, no PASE, no PC, no ODBC:

| Command | Program | What it does |
|---|---|---|
| `XLSX00` | `XLSX00R` | Writes the result of an SQL statement or a database file to an Excel `.xlsx` stream file: new workbook, add a sheet, append rows, replace one sheet of a template; headings from column names, labels or text; filter, frozen header, column notes, Excel table, named range, hidden sheet; own deflate (levels 0-9). 200,000 rows x 10 columns in about 6 seconds on one Power9 core. |
| `XLSX01` | `XLSX10R` | Reads one sheet of an `.xlsx` stream file into a Db2 table: creates the table (names and types from the data) or fills an existing one (columns matched by heading or position); finds the data by itself (Excel table, filter range, title rows), reads only visible rows if asked, checks every value before writing (`ERRORS(*STOP)`), report file with the decisions. 200,000 rows x 10 columns in about 14 seconds on one Power9 core. |

Both commands have full F1 help (every parameter, examples, messages); `XLSX00-help.html` and `XLSX01-help.html` are the same help as web pages.

## Package

| File | Contents |
|---|---|
| `xlsx-1.4.1-objects.savf` | save file with 6 objects, saved from library `CESCO2`, target release V7R5M0: `XLSX00R` and `XLSX10R` (*PGM), `XLSX00` and `XLSX01` (*CMD), `XLSX00H` and `XLSX01H` (*PNLGRP) |
| `xlsx-1.4.1-sources.savf` | save file (SAV) of the source directory `/tmp/xlsx-1.4.1`: RPG, include files, command and help sources, the build program `XLSXDIST.clle`, this README, LICENSE, THIRD-PARTY-NOTICES.txt |
| `XLSX00-help.html`, `XLSX01-help.html` | the command help |
| `SHA256SUMS` | checksums |

Requirements: IBM i 7.5 or later. The programs run in activation group *NEW, adopt no authority, and use no other product (only the C runtime of the system). The commands find their program and help panel group through the library list (`*LIBL`), so the objects work in any library.

## Install

1. Copy the save file to the IBM i in binary, for example to `/home/me/xlsx-1.4.1-objects.savf` (FTP `bin`, SCP, or a mapped IFS share).
2. Create a save file, copy the stream file into it, and restore into a library of your choice (here `XLSX`):
   ```
   CRTLIB     LIB(XLSX) TEXT('XLSX tools')
   CRTSAVF    FILE(QGPL/XLSXOBJ)
   CPYFRMSTMF FROMSTMF('/home/me/xlsx-1.4.1-objects.savf') TOMBR('/QSYS.LIB/QGPL.LIB/XLSXOBJ.FILE') MBROPT(*REPLACE)
   RSTOBJ     OBJ(*ALL) SAVLIB(CESCO2) DEV(*SAVF) SAVF(QGPL/XLSXOBJ) RSTLIB(XLSX)
   DLTF       FILE(QGPL/XLSXOBJ)
   ```
   `SAVLIB(CESCO2)` is the library the objects were saved from on the build system; `RSTLIB` is yours.
3. Put the library in the library list of the users (`ADDLIBLE XLSX`, or the job description / system library list). Optional, to use the commands without the library in the list: point them to it
   ```
   CHGCMD CMD(XLSX/XLSX00) PGM(XLSX/XLSX00R) HLPPNLGRP(XLSX/XLSX00H)
   CHGCMD CMD(XLSX/XLSX01) PGM(XLSX/XLSX10R) HLPPNLGRP(XLSX/XLSX01H)
   ```
   and call them qualified (`XLSX/XLSX00 ...`). Repeat this after every update.
4. Owner: the objects belong to a user profile of the build system; on your system that profile does not exist and the restore gives them to `QDFTOWN`. Give them to a profile of yours if you like: `CHGOBJOWN OBJ(XLSX/XLSX00R) OBJTYPE(*PGM) NEWOWN(QPGMR)` and the same for the other five. Public authority is `*USE`.

## Try it

```
XLSX00 SRC('select * from qiws/qcustcdt') TOSTMF('/home/me/cust.xlsx')
XLSX01 FROMSTMF('/home/me/cust.xlsx') TOFILE(QTEMP/CUST)
XLSX01 FROMSTMF('/home/me/any.xlsx') TOFILE(*NONE) REPORT('/home/me/any.txt')
```
The last one only analyses the workbook and writes what it would do (sheet, data area, header, column types, the CREATE TABLE statement) to the report file. Prompt with F4, help with F1.

From RPG: copy the prototypes `XLSX00PR.rpgleinc` / `XLSX10PR.rpgleinc` from the sources and call the programs; with the message parameter passed, errors come back in it instead of an escape message.

## Update

Restore the new package over the old objects. Two ways:
- with *ALLOBJ special authority:
  ```
  RSTOBJ OBJ(*ALL) SAVLIB(CESCO2) DEV(*SAVF) SAVF(QGPL/XLSXOBJ) RSTLIB(XLSX) ALWOBJDIF(*OWNER)
  ```
  (`ALWOBJDIF(*OWNER)` is needed when the installed objects have another owner than the profile of the build system, which is the normal case; without it the restore leaves them out with CPF3766);
- without it: delete the six objects first, then restore as in "Install":
  ```
  DLTPGM PGM(XLSX/XLSX00R)   DLTPGM PGM(XLSX/XLSX10R)
  DLTCMD CMD(XLSX/XLSX00)    DLTCMD CMD(XLSX/XLSX01)
  DLTPNLGRP PNLGRP(XLSX/XLSX00H)   DLTPNLGRP PNLGRP(XLSX/XLSX01H)
  ```
Then repeat steps 3 and 4 of "Install" where you used them. The commands keep no data of their own: nothing else to migrate.

## Remove

Delete the six objects (above), or the library if it holds nothing else: `DLTLIB LIB(XLSX)`.

## Build from the sources

```
CRTSAVF    FILE(QGPL/XLSXSRC)
CPYFRMSTMF FROMSTMF('/home/me/xlsx-1.4.1-sources.savf') TOMBR('/QSYS.LIB/QGPL.LIB/XLSXSRC.FILE') MBROPT(*REPLACE)
RST        DEV('/QSYS.LIB/QGPL.LIB/XLSXSRC.FILE') OBJ(('/tmp/xlsx-1.4.1' *INCLUDE '/home/me/xlsx'))
CRTBNDCL   PGM(QTEMP/XLSXDIST) SRCSTMF('/home/me/xlsx/XLSXDIST.clle')
CALL       PGM(QTEMP/XLSXDIST) PARM(XLSX '/home/me/xlsx/')
```
`XLSXDIST` compiles the modules into QTEMP and creates the six objects in the library given (it must exist). A path longer than 32 characters must be passed padded with blanks to 128 characters (CALL passes shorter literals padded to 32 only). The modules take the CCSID of the job (`TGTCCSID(*JOB)`); the package was built under CCSID 1144.

## Notes

- Dates before 1900-01-01 are written by XLSX00 as text (Excel has no serial number for them).
- XLSX01 writes without commitment control: with `ERRORS(*STOP)` a value that does not fit stops the run before anything is written; a database error during the load (duplicate key, constraint) stops it there, with the rows before it already in the table and the Excel row named in the message.
- Work objects: XLSX00 may create an alias `QTEMP/XLSX00A`, XLSX01 a work table `QTEMP/XLSX10S` and a DSPFFD outfile `QTEMP/XLSX10F` (for tables in QTEMP). Nothing is left outside QTEMP.

## Licence

MIT, see `LICENSE`. Algorithms of zlib and puff are acknowledged in `THIRD-PARTY-NOTICES.txt`.
