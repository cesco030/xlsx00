/* XLSX01 - read a worksheet of an Excel .xlsx stream file into a table   */
/* CPP XLSX10R. Build: XLSXDIST (CPYFRMSTMF to QTEMP/QCMDSRC + CRTCMD).    */
/* Version 1.4.1 (2026-11-17)                                             */
/* XLSX01 FROMSTMF('/home/me/orders.xlsx') TOFILE(MYLIB/ORDERS)           */
             CMD        PROMPT('Import Excel XLSX to table')
             PARM       KWD(FROMSTMF) TYPE(*CHAR) LEN(1024) MIN(1) +
                          EXPR(*YES) CASE(*MIXED) +
                          PROMPT('From stream file (.xlsx)')
             PARM       KWD(TOFILE) TYPE(QUALTAB) MIN(1) +
                          PROMPT('To table')
/* program parameters 3-4 are outputs: constants here, '*ESCAPE' = send   */
/* an escape message for errors                                           */
             PARM       KWD(ROWS) TYPE(*DEC) LEN(15 0) CONSTANT(0)
             PARM       KWD(MSG) TYPE(*CHAR) LEN(256) CONSTANT('*ESCAPE')
             PARM       KWD(SHEET) TYPE(*CHAR) LEN(64) DFT(*FIRST) +
                          SPCVAL((*FIRST)) EXPR(*YES) CASE(*MIXED) +
                          PROMPT('Sheet name or number')
             PARM       KWD(RANGE) TYPE(*CHAR) LEN(256) DFT(*AUTO) +
                          SPCVAL((*AUTO)) EXPR(*YES) CASE(*MIXED) +
                          PROMPT('Cell range, table or name')
             PARM       KWD(HEADER) TYPE(*CHAR) LEN(10) RSTD(*YES) +
                          DFT(*AUTO) VALUES(*AUTO *YES *NO *SKIP) +
                          PROMPT('Header row')
             PARM       KWD(STARTROW) TYPE(*DEC) LEN(7 0) DFT(*AUTO) +
                          RANGE(1 1048576) SPCVAL((*AUTO 0)) +
                          PROMPT('First row (header or data)')
             PARM       KWD(ROWSEL) TYPE(*CHAR) LEN(10) RSTD(*YES) +
                          DFT(*ALL) VALUES(*ALL *VISIBLE) +
                          PROMPT('Rows to read')
             PARM       KWD(MAXROWS) TYPE(*DEC) LEN(9 0) DFT(*ALL) +
                          RANGE(1 1048576) SPCVAL((*ALL 0)) +
                          PROMPT('Maximum rows')
             PARM       KWD(MBROPT) TYPE(*CHAR) LEN(10) RSTD(*YES) +
                          DFT(*ADD) VALUES(*ADD *REPLACE) +
                          PROMPT('Replace or add rows')
             PARM       KWD(CRTFILE) TYPE(*CHAR) LEN(10) RSTD(*YES) +
                          DFT(*YES) VALUES(*YES *NO *REPLACE) +
                          PROMPT('Create table')
             PARM       KWD(MAP) TYPE(*CHAR) LEN(10) RSTD(*YES) +
                          DFT(*AUTO) VALUES(*AUTO *NAME *POS) +
                          PROMPT('Column mapping')
             PARM       KWD(COLMAP) TYPE(*CHAR) LEN(2000) DFT(*NONE) +
                          SPCVAL((*NONE)) EXPR(*YES) CASE(*MIXED) +
                          PROMPT('Column map (Heading=COL;..)')
             PARM       KWD(TYPES) TYPE(*CHAR) LEN(2000) DFT(*AUTO) +
                          SPCVAL((*AUTO '*NONE')) EXPR(*YES) +
                          CASE(*MIXED) +
                          PROMPT('Column types (COL=type;..)')
             PARM       KWD(NULLS) TYPE(*CHAR) LEN(4) RSTD(*YES) +
                          DFT(*NO) VALUES(*NO *YES) +
                          PROMPT('Nullable new columns')
             PARM       KWD(ERRORS) TYPE(*CHAR) LEN(10) RSTD(*YES) +
                          DFT(*STOP) VALUES(*STOP *SKIP *NULL) +
                          PROMPT('Values that do not fit')
             PARM       KWD(STRCCSID) TYPE(*CHAR) LEN(10) DFT(*AUTO) +
                          SPCVAL((*AUTO) (*JOB) (*UTF8) (*UTF16)) +
                          CHOICE('*AUTO, *JOB, *UTF8, CCSID...') +
                          PROMPT('CCSID of new text columns')
             PARM       KWD(TRIM) TYPE(*CHAR) LEN(4) RSTD(*YES) +
                          DFT(*YES) VALUES(*YES *NO) +
                          PROMPT('Trim blanks of text')
             PARM       KWD(DATFMT) TYPE(*CHAR) LEN(10) RSTD(*YES) +
                          DFT(*DMY) VALUES(*DMY *MDY *YMD) +
                          PROMPT('Order of text dates')
             PARM       KWD(REPORT) TYPE(*CHAR) LEN(1024) DFT(*NONE) +
                          SPCVAL((*NONE)) EXPR(*YES) CASE(*MIXED) +
                          PROMPT('Report stream file')
 QUALTAB:    QUAL       TYPE(*NAME) LEN(10) SPCVAL((*NONE)) MIN(1)
             QUAL       TYPE(*NAME) LEN(10) DFT(*LIBL) +
                          SPCVAL((*LIBL) (*CURLIB)) PROMPT('Library')
