/* XLSX00 - SQL statement or database file -> Excel .xlsx stream file    */
/* CPP XLSX00R. Build: XLSXDIST (CPYFRMSTMF to QTEMP/QCMDSRC + CRTCMD).    */
/* Version 1.4.1 (2026-11-17)                                                */
/* XLSX00 SRC('select * from mylib/orders') TOSTMF('/home/me/orders.xlsx') */
             CMD        PROMPT('Export SQL/file to Excel XLSX')
             PARM       KWD(SRC) TYPE(*CHAR) LEN(5000) MIN(1) +
                          EXPR(*YES) CASE(*MIXED) +
                          PROMPT('SQL statement or LIB/FILE(MBR)')
             PARM       KWD(TOSTMF) TYPE(*CHAR) LEN(1024) MIN(1) +
                          EXPR(*YES) CASE(*MIXED) +
                          PROMPT('To stream file (.xlsx)')
             PARM       KWD(MODE) TYPE(*CHAR) LEN(10) RSTD(*YES) +
                          DFT(*REPLACE) VALUES(*REPLACE *ADDSHEET +
                          *APPEND *REPLSHEET) PROMPT('Mode')
             PARM       KWD(SHEET) TYPE(*CHAR) LEN(31) DFT(*DFT) +
                          SPCVAL((*DFT ' ')) EXPR(*YES) +
                          CASE(*MIXED) PROMPT('Sheet name')
             PARM       KWD(HEADER) TYPE(*CHAR) LEN(10) RSTD(*YES) +
                          DFT(*LABELS) VALUES(*LABELS *TEXT +
                          *NAMES *NONE) PROMPT('Header row')
             PARM       KWD(HDRTEXT) TYPE(*CHAR) LEN(5000) +
                          DFT(*NONE) SPCVAL((*NONE ' ')) EXPR(*YES) +
                          CASE(*MIXED) +
                          PROMPT('Column headings (a;b;c)')
             PARM       KWD(FILTER) TYPE(*CHAR) LEN(4) RSTD(*YES) +
                          DFT(*YES) VALUES(*YES *NO) +
                          PROMPT('Autofilter on header row')
             PARM       KWD(FREEZE) TYPE(*CHAR) LEN(4) RSTD(*YES) +
                          DFT(*YES) VALUES(*YES *NO) +
                          PROMPT('Freeze header row')
             PARM       KWD(LEVEL) TYPE(*DEC) LEN(1 0) DFT(2) +
                          RANGE(0 9) PROMPT('Compression level 0-9')
             PARM       KWD(COLDESC) TYPE(*CHAR) LEN(10) RSTD(*YES) +
                          DFT(*NONE) VALUES(*NONE *ROW *NOTE *BOTH) +
                          PROMPT('Column text: row / notes')
/* program parameters 11-12 are outputs: constants here, '*ESCAPE' = keep  */
/* the escape messages for errors                                         */
             PARM       KWD(ROWS) TYPE(*DEC) LEN(15 0) CONSTANT(0)
             PARM       KWD(MSG) TYPE(*CHAR) LEN(256) CONSTANT('*ESCAPE')
             PARM       KWD(RANGENAME) TYPE(*CHAR) LEN(255) DFT(*NONE) +
                          SPCVAL((*NONE ' ')) EXPR(*YES) CASE(*MIXED) +
                          PROMPT('Defined name over the data')
             PARM       KWD(TABLE) TYPE(*CHAR) LEN(255) DFT(*NONE) +
                          SPCVAL((*NONE ' ') (*YES '*YES')) EXPR(*YES) +
                          CASE(*MIXED) PROMPT('Excel table (name or *YES)')
             PARM       KWD(HIDE) TYPE(*CHAR) LEN(10) RSTD(*YES) +
                          DFT(*DFT) VALUES(*DFT *NO *YES *VERY) +
                          PROMPT('Hide the sheet')
