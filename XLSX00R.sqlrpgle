**free
// =====================================================================================
// XLSX00R - SQL statement or database file -> Excel workbook (.xlsx, OOXML/ECMA-376)
//
// Pure ILE RPG, no dependencies (Db2 for i, IFS and C runtime only): the ZIP container,
// deflate and inflate are implemented in XLSX01R / XLSX02R.
//
// Parameters (see XLSX00PR.rpgleinc, command XLSX00):
//   1 SRC     SQL (SELECT / WITH / VALUES) or a file: LIB/FILE, LIB.FILE, FILE, FILE(MBR)
//   2 TOSTMF  target .xlsx
//   3 MODE    *REPLACE  new workbook (default)
//             *ADDSHEET add a sheet to an existing workbook (created if missing)
//             *APPEND   add rows at the bottom of an existing sheet (SHEET, blank = last
//                       sheet; a sheet that does not exist is added as with *ADDSHEET)
//             *REPLSHEET rewrite the content of the sheet named SHEET (blank = default
//                       name), keeping its name and position and all other parts; Excel
//                       recalculates formulas on open (calcChain dropped, fullCalcOnLoad).
//                       A sheet that does not exist is added as with *ADDSHEET.
//   4 SHEET   sheet name (blank: file name, or Sheet1); made unique on *ADDSHEET
//   5 HEADER  *LABELS (column headings, else names) | *TEXT (column text, else
//             heading, else name) | *NAMES | *NONE
//   6 HDRLIST custom headings 'A;B;;D' (an empty item keeps the default heading)
//   7 FILTER  *YES/*NO autofilter on the header row     (default *YES)
//   8 FREEZE  *YES/*NO frozen header row                (default *YES)
//   9 LEVEL   0 = store, 1-4 greedy, 5-9 lazy deflate      (default 2)
//  10 COLDESC column text (LABEL ON ... TEXT IS) of the base columns: *NONE | *ROW
//             (2nd header row) | *NOTE (Excel note on each heading: table.column,
//             type, text, heading) | *BOTH. New sheets only for notes.
//  11 ROWS    out: data rows written
//  12 MSG     out: error text (blank = ok). Not passed, or '*ESCAPE' on entry (this is
//             what command XLSX00 passes): errors -> escape CPF9898.
//  13 RANGENAME workbook defined name over header + data ('' / *NONE = none); replaced
//             when it exists, on *APPEND only its last row moves
//  14 TABLE   *NONE | *YES (name from the sheet name) | table name: Excel table over
//             header + data (own filter buttons). *APPEND extends existing tables that
//             end at the last row; *REPLSHEET drops the old sheet's tables first, so the
//             same name can be created again (template pivots/charts keep working).
//  15 HIDE    *DFT (new sheets visible, existing sheets unchanged) | *NO | *YES | *VERY
//
// Cells: numbers as numbers (DECIMAL scale 1-6 with a #,##0.0.. format), dates/times/
// timestamps as Excel serial values with a date format, text as inline strings
// (UTF-8, Db2 converts from the column CCSID; CCSID 65535 columns are read as job CCSID
// text), binary as hex, LOB text up to 32 KB, NULL = empty cell. Max 1048576 rows.
// The file is written as <TOSTMF>.tmp and renamed at the end (no half-written workbooks).
//
// Build: XLSXDIST.clle (CRTSQLRPGI OBJTYPE(*MODULE) + CRTPGM with XLSX01R, XLSX02R).
//
// Version 1.4.1 (2026-11-17) - XLSX_VERSION in XLSXINT.rpgleinc; bump it on every change
// (minor = new function, patch = fix / small change) and add a line here.
//   1.0.0 2026-10-28 first release: *REPLACE, *ADDSHEET, *APPEND, HEADER, COLDESC, F1 help
//   1.1.0 2026-10-28 deflate: lazy matching at levels 5-9 (zlib-style), level 4 = old 6
//   1.2.0 2026-10-28 MODE(*REPLSHEET)
//   1.3.0 2026-10-28 RANGENAME, TABLE, HIDE
//   1.3.1 2026-10-29 version in messages, object texts, help, docProps/app.xml
//   1.4.0 2026-11-12 reader: command XLSX01 / program XLSX10R (.xlsx -> Db2 table)
//         writer: CR / control characters as _x000D_, literal _xHHHH_ as _x005F_xHHHH_
//   1.4.1 2026-11-17 relocatable sources (/copy relative + INCDIR), XLSXDIST build,
//         XLSX10R with explicit SQL options (*PERIOD, *ISO), package built on PUB400
//         XLSX10R fixes: table text for a short path or sheet name, duplicate headings
//         (column NAME_2), blank date-like text (all were RNX0100)
// =====================================================================================
ctl-opt main(xlsxMain) alloc(*teraspace) bnddir('QC2LE')
        option(*srcstmt : *nodebugio : *nounref);

exec sql set option commit = *none, naming = *sys, datfmt = *iso, timfmt = *iso,
                    decmpt = *period, closqlcsr = *endmod, alwblk = *allread;

/copy 'XLSXINT.rpgleinc'

dcl-c SQL_NUM 1000;
exec sql include SQLDA;

dcl-pr sndPgmMsg extpgm('QMHSNDPM');
  msgId   char(7) const;
  msgFile char(20) const;
  msgData char(512) const;
  msgDLen int(10) const;
  msgType char(10) const;
  callStk char(10) const;
  callCnt int(10) const;
  msgKey  char(4);
  errCode char(8);
end-pr;

dcl-pr getenv pointer extproc('getenv');
  name pointer value options(*string);
end-pr;

dcl-pr fmtLabel varchar(62);
  l varchar(60) const;
end-pr;
dcl-pr fmtFloat varchar(40);
  f float(8) value;
end-pr;

dcl-c XMAXROW   1048576;
dcl-c XBIGMAX   999000;
dcl-c NS_MAIN   'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
dcl-c NS_REL    'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
dcl-c NS_PKG    'http://schemas.openxmlformats.org/package/2006/relationships';
dcl-c NS_CT     'http://schemas.openxmlformats.org/package/2006/content-types';
dcl-c CT_OXML   'application/vnd.openxmlformats-officedocument.';
dcl-c XMLDECL   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

dcl-s TOK_T  varchar(64) ccsid(*utf8) template;
dcl-s XMID_T varchar(32000 : 4) ccsid(*utf8) template;
dcl-s XBIG_T varchar(1000000 : 4) ccsid(*utf8) template;
dcl-pr writePart ind;
  name like(XMID_T) const;
  s    like(XBIG_T);
end-pr;
dcl-pr loadPart int(10);
  name like(XMID_T) const;
  s    like(XBIG_T);
end-pr;
dcl-pr uniqueHeaders;
end-pr;
dcl-pr writeSheetRels ind;
  part like(XMID_T) const;
end-pr;
dcl-pr writeTable ind;
  lastRow int(20) value;
end-pr;

// ---- options -------------------------------------------------------------------------
dcl-s gSrc     varchar(5000);
dcl-s gStmf    varchar(1024);
dcl-s gTmpF    varchar(1024);
dcl-s gMode    char(1);                 // R replace, S add sheet, A append, P repl. sheet
dcl-s gSheet   varchar(31);
dcl-s gSheetGiven ind;
dcl-s gHdrMode char(1);                 // L labels, T text, N names, X none
dcl-s gDesc    char(1);                 // column descriptions: N none, R row, C note, B both
dcl-s gHdrRows int(10);                 // header rows written (0, 1, 2)
dcl-s gNotes   ind;                     // notes on the header of the new sheet
dcl-s gNoteN   int(10);                 // number of the comments / vmlDrawing parts
dcl-s gHdr     ind;
dcl-s gHdrList varchar(5000);
dcl-s gFilter  ind;
dcl-s gFreeze  ind;
dcl-s gLevel   int(10);
dcl-s gRName   varchar(255);            // RANGENAME: defined name over the data ('' = none)
dcl-s gTbl     ind;                     // TABLE: Excel table on the new sheet
dcl-s gTName   varchar(255);            // table name ('' = derived from the sheet name)
dcl-s gHide    char(1);                 // HIDE: ' ' default, N visible, H hidden, V very
dcl-s gTblN    int(10);                 // number of the xl/tables/tableN.xml part
dcl-s gTblId   int(10);                 // table id (unique in the workbook)
dcl-s gApLast  int(20);                 // *APPEND: last row before the new rows
dcl-s gRow0    int(20);                 // first row of the data block (header row)
dcl-s nNames   int(10);                 // table + defined names already in the workbook
dcl-s nTblNm   int(10);                 // the first nTblNm of gNames are table names
dcl-s gNames   varchar(255) ccsid(*utf8) dim(500);  // upper case
dcl-s nDrop    int(10);                 // parts left out of the package (*REPLSHEET)
dcl-s gDrop    varchar(512) ccsid(*utf8) dim(100);
dcl-s gZrN     int(10);                 // entries of the existing package
dcl-s nApTbl   int(10);                 // tables of the appended sheet
dcl-s gApTbl   varchar(512) ccsid(*utf8) dim(50);

// ---- state ---------------------------------------------------------------------------
dcl-s gFail    ind;
dcl-s gMsg     varchar(512);
dcl-s gRows    int(20);
dcl-s gTrunc   ind;
dcl-s gUtf8Ok  ind;
dcl-s gIsFile  ind;
dcl-s gDftSheet varchar(31);
dcl-s gCurOpen ind;
dcl-s gZrOpen  ind;
dcl-s gStyAdd  ind;
dcl-s gNewWB   ind;
dcl-s gOutName varchar(31);              // final sheet name (job CCSID, for messages)
dcl-s gStats   ind;                    // env XLSX00_STATS=Y: timing in the message
dcl-s gFetchUs int(20);
dcl-s gStmt    varchar(32700);
dcl-s gDdl     varchar(2000);

// ---- columns -------------------------------------------------------------------------
dcl-ds col qualified dim(1000);
  kind  char(1);                        // S utf8 text, E job-CCSID text, I int, P packed,
  len   int(10);                        // F float, D date, T time, Z timestamp, H hex,
  dlen  int(10);                        // X skipped
  prec  int(10);
  scale int(10);
  off   int(10);
  ref   varchar(64) ccsid(*utf8);
  styS  varchar(64) ccsid(*utf8);
  cPre  varchar(64) ccsid(*utf8);           // cell start with the column letter
  cPost varchar(64) ccsid(*utf8);           // style + value start of number cells
  hdr   varchar(1600) ccsid(*utf8);
  dsc   varchar(400) ccsid(*utf8);            // escaped description (text) for row 2
  tdesc varchar(40);                          // SQL type, e.g. DECIMAL(11,2)
  cnm   varchar(30);                          // result column name
  hlen  int(10);
  width packed(5 : 1);
end-ds;
dcl-s nCol  int(10);
dcl-s gMulti   ind;                     // block fetch (no LOB result columns)
dcl-s gBlkRows int(10);                 // rows per FETCH
dcl-s gRowLen  int(10);                 // bytes per row in the block (no padding)
dcl-s pBlk  pointer inz(*null);         // row storage area
dcl-s pBInd pointer inz(*null);         // indicator area: 1 SMALLINT per column per row
dcl-s gRowArea char(16000000) based(pBlk);
dcl-s gIndArea char(2000000) based(pBInd);
dcl-s gLbl  varchar(60) dim(1000);         // column heading (DB2_LABEL)
dcl-s gTxt  varchar(50) dim(1000);         // column text (SYSCOLUMNS.COLUMN_TEXT)
dcl-s gBase varchar(400) dim(1000);        // SCHEMA.TABLE.COLUMN of the base column
dcl-s gCus  varchar(256) dim(1000);
dcl-s ST    int(10) dim(11);            // xf: hdr, date, ts, time, dec1..dec6, desc row

// SQLVAR overlay (an element of SQL_VAR in the SQLDA)
dcl-s pv pointer;
dcl-ds sv qualified based(pv) align;
  vType    int(5)  pos(1);
  vLen     int(5)  pos(3);
  vP       uns(3)  pos(3);
  vS       uns(3)  pos(4);
  vData    pointer pos(17);
  vInd     pointer pos(33);
  vNameLen int(5)  pos(49);
  vName    char(30) pos(51);
end-ds;

// ---- UTF-8 tokens --------------------------------------------------------------------
dcl-s tRowA like(TOK_T);
dcl-s tRowB like(TOK_T);
dcl-s tRowE like(TOK_T);
dcl-s tCA   like(TOK_T);
dcl-s tCS   like(TOK_T);
dcl-s tCSP  like(TOK_T);
dcl-s tCSE  like(TOK_T);
dcl-s tCS2  like(TOK_T);
dcl-s tCV   like(TOK_T);
dcl-s tCVE  like(TOK_T);
dcl-s tCV0  like(TOK_T);
dcl-s tAmp  like(TOK_T);
dcl-s tLt   like(TOK_T);
dcl-s tGt   like(TOK_T);
dcl-s tSpc  like(TOK_T);
dcl-s cLt   like(TOK_T);
dcl-s cGt   like(TOK_T);
dcl-s cSl   like(TOK_T);
dcl-s cSp   like(TOK_T);
dcl-s cQuot like(TOK_T);
dcl-s cColon like(TOK_T);
dcl-s cDol  like(TOK_T);
dcl-s cLf   like(TOK_T);
dcl-s cCr   like(TOK_T);
dcl-s cTab  like(TOK_T);
dcl-s uLower like(TOK_T);
dcl-s uUpper like(TOK_T);
dcl-s CLS   uns(3) dim(256);           // 0 = byte goes into a text node unchanged
dcl-ds bx;
  bxn uns(3);
  bxc char(1) ccsid(*hex) pos(1);
end-ds;

// ---- parts of the package ------------------------------------------------------------
dcl-s gCT  like(XBIG_T);
dcl-s gWB  like(XBIG_T);
dcl-s gRel like(XBIG_T);
dcl-s gSty like(XBIG_T);
dcl-s gTmp like(XBIG_T);
dcl-s gU   varchar(100000 : 4) ccsid(*utf8);

// ---- writer --------------------------------------------------------------------------
dcl-s wp pointer inz(*null);
dcl-s we pointer inz(*null);
dcl-s wb uns(3) based(wp);
dcl-s rowNum int(20);
dcl-s rowDig like(TOK_T);
dcl-s excelBase int(10);

// =====================================================================================
// small helpers
// =====================================================================================
dcl-proc fail;
  dcl-pi *n;
    msg varchar(512) const;
  end-pi;
  if not gFail;
    gFail = *on;
    gMsg = msg;
  endif;
end-proc;

// why a defined / table name is not valid for Excel ('' = valid). ASCII letters, digits,
// _ . \ ; first a letter, _ or \ ; not a cell reference (A1 style) or R/C (R1C1 style)
dcl-proc nameErr;
  dcl-pi *n varchar(80);
    n varchar(255) const;
  end-pi;
  dcl-c LETT 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
  dcl-c DIG '0123456789';
  dcl-s up varchar(255);
  dcl-s k int(10);
  if %len(n) = 0 or %len(n) > 255;
    return 'length 1-255';
  endif;
  if %check(LETT + '_\' : %subst(n : 1 : 1)) > 0;
    return 'must start with a letter, _ or \';
  endif;
  if %check(LETT + DIG + '_.\' : n) > 0;
    return 'only letters A-Z, digits, _ . \ allowed';
  endif;
  up = %xlate('abcdefghijklmnopqrstuvwxyz' : 'ABCDEFGHIJKLMNOPQRSTUVWXYZ' : n);
  k = %check('ABCDEFGHIJKLMNOPQRSTUVWXYZ' : up);
  if k >= 2 and k <= 4 and %check(DIG : %subst(up : k)) = 0;
    return 'looks like a cell reference';
  endif;
  if up = 'R' or up = 'C';
    return 'R and C are reserved';
  endif;
  if %subst(up : 1 : 1) = 'R' or %subst(up : 1 : 1) = 'C';
    if %check('RC' + DIG : up) = 0;
      return 'looks like an R1C1 reference';
    endif;
  endif;
  return '';
end-proc;

dcl-proc u;                              // job CCSID -> UTF-8
  dcl-pi *n like(XMID_T);
    s varchar(32000 : 4) const;
  end-pi;
  dcl-s r like(XMID_T);
  r = s;
  return r;
end-proc;

dcl-proc eb;                             // UTF-8 -> job CCSID (messages)
  dcl-pi *n varchar(1000);
    s varchar(1000) ccsid(*utf8) const;
  end-pi;
  dcl-s r varchar(1000);
  r = s;
  return r;
end-proc;

dcl-proc numU;
  dcl-pi *n like(TOK_T);
    n int(20) value;
  end-pi;
  dcl-s r like(TOK_T);
  r = %char(n);
  return r;
end-proc;

dcl-proc xesc;                           // escape for XML text and attributes
  dcl-pi *n like(XMID_T);
    s like(XMID_T) const;
  end-pi;
  dcl-s r like(XMID_T);
  r = %scanrpl(u('&') : u('&amp;') : s);
  r = %scanrpl(u('<') : u('&lt;') : r);
  r = %scanrpl(u('>') : u('&gt;') : r);
  r = %scanrpl(cQuot : u('&quot;') : r);
  return r;
end-proc;

dcl-proc upU;                            // ASCII upper case of a UTF-8 string
  dcl-pi *n like(XMID_T);
    s like(XMID_T) const;
  end-pi;
  return %xlate(uLower : uUpper : s);
end-proc;

dcl-proc digits;                         // leading decimal digits of a UTF-8 string
  dcl-pi *n int(20);
    sc like(XMID_T) const;
  end-pi;
  dcl-s s like(XMID_T);
  dcl-s p pointer;
  dcl-s b uns(3) based(p);
  dcl-s i int(10);
  dcl-s v int(20);
  v = 0;
  s = sc;
  p = %addr(s : *data);
  for i = 1 to %len(s);
    if b < 48 or b > 57 or i > 18;
      leave;
    endif;
    v = v * 10 + b - 48;
    p += 1;
  endfor;
  return v;
end-proc;

dcl-proc noTrailDigits;                  // 'AB123' -> 'AB'
  dcl-pi *n like(XMID_T);
    sc like(XMID_T) const;
  end-pi;
  dcl-s s like(XMID_T);
  dcl-s p pointer;
  dcl-s b uns(3) based(p);
  dcl-s j int(10);
  s = sc;
  j = %len(s);
  dow j > 0;
    p = %addr(s : *data) + j - 1;
    if b < 48 or b > 57;
      leave;
    endif;
    j -= 1;
  enddo;
  if j = 0;
    return u('');
  endif;
  return %subst(s : 1 : j);
end-proc;

dcl-proc colRef;                         // 1 -> A, 27 -> AA
  dcl-pi *n like(TOK_T);
    n int(10) value;
  end-pi;
  dcl-s r varchar(8);
  dcl-s k int(10);
  r = '';
  dow n > 0;
    k = %rem(n - 1 : 26);
    r = %subst('ABCDEFGHIJKLMNOPQRSTUVWXYZ' : k + 1 : 1) + r;
    n = %div(n - 1 : 26);
  enddo;
  return u(r);
end-proc;

// days from civil (proleptic Gregorian), H. Hinnant
dcl-proc dfc;
  dcl-pi *n int(10);
    y int(10) value;
    m int(10) value;
    d int(10) value;
  end-pi;
  dcl-s era int(10);
  dcl-s yoe int(10);
  dcl-s doy int(10);
  dcl-s doe int(10);
  dcl-s mp int(10);
  if m <= 2;
    y -= 1;
  endif;
  era = %div(y : 400);
  yoe = y - era * 400;
  if m > 2;
    mp = m - 3;
  else;
    mp = m + 9;
  endif;
  doy = %div(153 * mp + 2 : 5) + d - 1;
  doe = yoe * 365 + %div(yoe : 4) - %div(yoe : 100) + doy;
  return era * 146097 + doe - 719468;
end-proc;

dcl-proc initTokens;
  dcl-s b int(10);
  tRowA = u('<row r="');
  tRowB = u('">');
  tRowE = u('</row>');
  tCA   = u('<c r="');
  tCS   = u('" t="inlineStr"><is><t>');
  tCSP  = u('" t="inlineStr"><is><t xml:space="preserve">');
  tCSE  = u('</t></is></c>');
  tCS2  = u('" s="');
  tCV   = u('"><v>');
  tCVE  = u('</v></c>');
  tCV0  = u('"/>');
  tAmp  = u('&amp;');
  tLt   = u('&lt;');
  tGt   = u('&gt;');
  tSpc  = u(' ');
  cLt   = u('<');
  cGt   = u('>');
  cSl   = u('/');
  cSp   = u(' ');
  cQuot = u('"');
  cColon = u(':');
  cDol  = u('$');
  bxn = 10;
  cLf = bxc;
  bxn = 13;
  cCr = bxc;
  bxn = 9;
  cTab = bxc;
  uLower = u('abcdefghijklmnopqrstuvwxyz');
  uUpper = u('ABCDEFGHIJKLMNOPQRSTUVWXYZ');
  // bytes that go into a text node unchanged: class 0
  for b = 0 to 255;
    CLS(b + 1) = 1;
    if b = 9 or b = 10 or (b >= 32 and b <> 38 and b <> 60 and b <> 62 and b <> 95);
      CLS(b + 1) = 0;
    endif;
  endfor;
  excelBase = dfc(1899 : 12 : 30);
end-proc;


// =====================================================================================
// XML surgery on package parts (UTF-8 strings)
// =====================================================================================
dcl-proc findOpen;                       // position of '<tag' + (blank > /), 0 = none
  dcl-pi *n int(10);
    s    like(XBIG_T);
    tag  like(TOK_T) const;
    from int(10) value;
  end-pi;
  dcl-s nd like(TOK_T);
  dcl-s p int(10);
  dcl-s q int(10);
  dcl-s c like(TOK_T);
  nd = cLt + tag;
  p = from;
  if p < 1;
    p = 1;
  endif;
  dow p <= %len(s);
    p = %scan(nd : s : p);
    if p = 0;
      leave;
    endif;
    q = p + %len(nd);
    if q <= %len(s);
      c = %subst(s : q : 1);
      if c = cSp or c = cGt or c = cSl or c = cLf or c = cCr or c = cTab;
        return p;
      endif;
    endif;
    p += 1;
  enddo;
  return 0;
end-proc;

dcl-proc tagEnd;                         // position after the '>' of the tag at p
  dcl-pi *n int(10);
    s like(XBIG_T);
    p int(10) value;
  end-pi;
  dcl-s q int(10);
  if p < 1 or p > %len(s);
    return 0;
  endif;
  q = %scan(cGt : s : p);
  if q = 0;
    return 0;
  endif;
  return q + 1;
end-proc;

// 0 = missing, 1 = <tag ...> at p, open tag ends at e, </tag> at c; 2 = <tag .../>
dcl-proc section;
  dcl-pi *n int(10);
    s   like(XBIG_T);
    tag like(TOK_T) const;
    p   int(10);
    e   int(10);
    c   int(10);
  end-pi;
  p = findOpen(s : tag : 1);
  if p = 0;
    return 0;
  endif;
  e = tagEnd(s : p);
  if e = 0;
    return 0;
  endif;
  if %subst(s : e - 2 : 1) = cSl;
    return 2;
  endif;
  c = %scan(cLt + cSl + tag + cGt : s : e);
  if c = 0;
    return 0;
  endif;
  return 1;
end-proc;

dcl-proc countEl;                        // elements <child in [a, b)
  dcl-pi *n int(10);
    s     like(XBIG_T);
    child like(TOK_T) const;
    a     int(10) value;
    b     int(10) value;
  end-pi;
  dcl-s n int(10);
  dcl-s p int(10);
  n = 0;
  p = a;
  dow 1 = 1;
    p = findOpen(s : child : p);
    if p = 0 or p >= b;
      leave;
    endif;
    n += 1;
    p += 1;
  enddo;
  return n;
end-proc;

dcl-proc setCount;                       // count="n" on the open tag of <tag
  dcl-pi *n;
    s   like(XBIG_T);
    tag like(TOK_T) const;
    n   int(10) value;
  end-pi;
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s m int(10);
  dcl-s q int(10);
  p = findOpen(s : tag : 1);
  if p = 0;
    return;
  endif;
  e = tagEnd(s : p);
  m = %scan(u(' count="') : s : p);
  if m > 0 and m < e;
    q = %scan(cQuot : s : m + 8);
    s = %replace(numU(n) : s : m + 8 : q - m - 8);
  else;
    s = %replace(u(' count="') + numU(n) + cQuot : s : p + 1 + %len(tag) : 0);
  endif;
end-proc;

// append items to the collection <tag> (children <child>), returns the index of the
// first new item; a missing collection is created after the open tag of <afterTag>
dcl-proc appendTo;
  dcl-pi *n int(10);
    s        like(XBIG_T);
    tag      like(TOK_T) const;
    child    like(TOK_T) const;
    items    like(XMID_T) const;
    nItems   int(10) value;
    afterTag like(TOK_T) const;
  end-pi;
  dcl-s k int(10);
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s c int(10);
  dcl-s n int(10);
  dcl-s wrap like(XMID_T);
  if %len(s) + %len(items) + 300 > XBIGMAX;
    fail('styles.xml too large');
    return -1;
  endif;
  k = section(s : tag : p : e : c);
  wrap = cLt + tag + u(' count="') + numU(nItems) + u('">') + items
       + cLt + cSl + tag + cGt;
  if k = 0;
    if %len(afterTag) = 0;
      fail('styles.xml without <' + eb(tag) + '>: not supported');
      return -1;
    endif;
    p = findOpen(s : afterTag : 1);
    e = tagEnd(s : p);
    if p = 0 or e = 0;
      fail('styles.xml without <' + eb(afterTag) + '>: not supported');
      return -1;
    endif;
    s = %replace(wrap : s : e : 0);
    return 0;
  endif;
  if k = 2;
    s = %replace(wrap : s : p : e - p);
    return 0;
  endif;
  n = countEl(s : child : e : c);
  s = %replace(items : s : c : 0);
  setCount(s : tag : n + nItems);
  return n;
end-proc;

dcl-proc attr;                           // value of attribute nm in a tag text
  dcl-pi *n like(XMID_T);
    tag like(XMID_T) const;
    nm  like(TOK_T) const;
  end-pi;
  dcl-s m int(10);
  dcl-s sv0 int(10);
  dcl-s q int(10);
  dcl-s r like(XMID_T);
  %len(r) = 0;
  m = %scan(cSp + nm + u('="') : tag);
  if m = 0;
    return r;
  endif;
  sv0 = m + %len(nm) + 3;
  if sv0 > %len(tag);
    return r;
  endif;
  q = %scan(cQuot : tag : sv0);
  if q = 0;
    return r;
  endif;
  if q > sv0;
    r = %subst(tag : sv0 : q - sv0);
  endif;
  return r;
end-proc;

dcl-proc tagText;                        // text of the tag starting at p
  dcl-pi *n like(XMID_T);
    s like(XBIG_T);
    p int(10) value;
  end-pi;
  dcl-s e int(10);
  e = tagEnd(s : p);
  if e = 0 or e - p > 32000;
    return u('');
  endif;
  return %subst(s : p : e - p);
end-proc;

dcl-proc relsOf;                         // xl/worksheets/a.xml -> xl/worksheets/_rels/a.xml.rels
  dcl-pi *n like(XMID_T);
    part like(XMID_T) const;
  end-pi;
  dcl-s k int(10);
  dcl-s j int(10);
  k = 0;
  j = %scan(cSl : part);
  dow j > 0;
    k = j;
    if j >= %len(part);
      leave;
    endif;
    j = %scan(cSl : part : j + 1);
  enddo;
  return %subst(part : 1 : k) + u('_rels/') + %subst(part : k + 1) + u('.rels');
end-proc;

// remove the empty elements <tag .../> whose attribute an ends with suf; returns count
dcl-proc dropEl;
  dcl-pi *n int(10);
    s   like(XBIG_T);
    tag like(TOK_T) const;
    an  like(TOK_T) const;
    suf like(XMID_T) const;
  end-pi;
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s n int(10);
  dcl-s t like(XMID_T);
  dcl-s v like(XMID_T);
  n = 0;
  p = 1;
  dow 1 = 1;
    p = findOpen(s : tag : p);
    if p = 0;
      leave;
    endif;
    e = tagEnd(s : p);
    if e = 0;
      leave;
    endif;
    t = %subst(s : p : e - p);
    v = attr(t : an);
    if %len(v) >= %len(suf) and %subst(t : %len(t) - 1 : 1) = cSl;
      if %subst(v : %len(v) - %len(suf) + 1) = suf;
        s = %replace(u('') : s : p : e - p);
        n += 1;
        iter;
      endif;
    endif;
    p = e;
  enddo;
  return n;
end-proc;

// =====================================================================================
// styles: append our fonts / fills / borders / number formats / cell formats
// =====================================================================================
dcl-proc mergeStyles;
  dcl-pi *n ind;
  end-pi;
  dcl-s k int(10);
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s c int(10);
  dcl-s q int(10);
  dcl-s mx int(10);
  dcl-s v int(10);
  dcl-s j int(10);
  dcl-s fb int(10);
  dcl-s fl int(10);
  dcl-s bd int(10);
  dcl-s x0 int(10);
  dcl-s items like(XMID_T);
  dcl-s nf like(TOK_T);

  mx = 163;
  k = section(gSty : u('numFmts') : p : e : c);
  if k = 1;
    q = e;
    nf = u('numFmtId="');
    dow q < c;
      q = %scan(nf : gSty : q);
      if q = 0 or q >= c;
        leave;
      endif;
      v = 0;
      if q + 10 <= %len(gSty);
        v = digits(%subst(gSty : q + 10 : 1));
        if q + 19 <= %len(gSty);
          v = digits(%subst(gSty : q + 10 : 10));
        endif;
      endif;
      if v > mx;
        mx = v;
      endif;
      q += 10;
    enddo;
  endif;
  %len(items) = 0;
  for j = 1 to 6;
    items += u('<numFmt numFmtId="') + numU(mx + j) + u('" formatCode="#,##0.')
           + u(%subst('000000' : 1 : j)) + u('"/>');
  endfor;
  if appendTo(gSty : u('numFmts') : u('numFmt') : items : 6 : u('styleSheet')) < 0;
    return *off;
  endif;
  fb = appendTo(gSty : u('fonts') : u('font')
       : u('<font><b/><sz val="11"/><color rgb="FF1F3864"/><name val="Calibri"/>'
       + '<family val="2"/></font><font><i/><sz val="10"/><color rgb="FF404040"/>'
       + '<name val="Calibri"/><family val="2"/></font>') : 2 : u(''));
  fl = appendTo(gSty : u('fills') : u('fill')
       : u('<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/>'
       + '<bgColor indexed="64"/></patternFill></fill>') : 1 : u(''));
  bd = appendTo(gSty : u('borders') : u('border')
       : u('<border><left/><right/><top/><bottom style="thin"><color rgb="FF8EA9DB"/>'
       + '</bottom><diagonal/></border>') : 1 : u(''));
  if fb < 0 or fl < 0 or bd < 0;
    return *off;
  endif;
  items = u('<xf numFmtId="0" fontId="') + numU(fb) + u('" fillId="') + numU(fl)
        + u('" borderId="') + numU(bd) + u('" xfId="0" applyFont="1" applyFill="1"'
        + ' applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>');
  items += u('<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0"'
         + ' applyNumberFormat="1"/>');
  items += u('<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0"'
         + ' applyNumberFormat="1"/>');
  items += u('<xf numFmtId="21" fontId="0" fillId="0" borderId="0" xfId="0"'
         + ' applyNumberFormat="1"/>');
  for j = 1 to 6;
    items += u('<xf numFmtId="') + numU(mx + j) + u('" fontId="0" fillId="0"'
           + ' borderId="0" xfId="0" applyNumberFormat="1"/>');
  endfor;
  items += u('<xf numFmtId="0" fontId="') + numU(fb + 1) + u('" fillId="') + numU(fl)
         + u('" borderId="') + numU(bd) + u('" xfId="0" applyFont="1" applyFill="1"'
         + ' applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/>'
         + '</xf>');
  x0 = appendTo(gSty : u('cellXfs') : u('xf') : items : 11 : u(''));
  if x0 < 0;
    return *off;
  endif;
  for j = 1 to 11;
    ST(j) = x0 + j - 1;
  endfor;
  return *on;
end-proc;

// =====================================================================================
// buffered writer on the current zip entry
// =====================================================================================
dcl-proc need;                           // make room for n bytes at wp
  dcl-pi *n ind;
    n int(10) value;
  end-pi;
  if wp <> *null;
    if wp + n <= we;
      return *on;
    endif;
  endif;
  wp = zwReserve(wp : n : we);
  if wp = *null;
    fail(xlErr());
    return *off;
  endif;
  return *on;
end-proc;

dcl-proc putT;                           // token (caller did need())
  dcl-pi *n;
    t like(TOK_T);
  end-pi;
  if %len(t) > 0;
    memcpy(wp : %addr(t : *data) : %len(t));
    wp += %len(t);
  endif;
end-proc;

dcl-proc putBig;                         // a whole part string
  dcl-pi *n ind;
    s like(XBIG_T);
  end-pi;
  zwCommit(wp);
  wp = *null;
  we = *null;
  if zwWrite(%addr(s : *data) : %len(s)) < 0;
    fail(xlErr());
    return *off;
  endif;
  return *on;
end-proc;

dcl-proc putMid;
  dcl-pi *n ind;
    sc like(XMID_T) const;
  end-pi;
  dcl-s s like(XMID_T) static;
  s = sc;
  zwCommit(wp);
  wp = *null;
  we = *null;
  if zwWrite(%addr(s : *data) : %len(s)) < 0;
    fail(xlErr());
    return *off;
  endif;
  return *on;
end-proc;

dcl-proc putRaw;                         // bytes from memory
  dcl-pi *n ind;
    p pointer value;
    n int(20) value;
  end-pi;
  zwCommit(wp);
  wp = *null;
  we = *null;
  if n > 0;
    if zwWrite(p : n) < 0;
      fail(xlErr());
      return *off;
    endif;
  endif;
  return *on;
end-proc;

dcl-proc putUns;                         // decimal digits
  dcl-pi *n;
    v uns(20) value;
  end-pi;
  dcl-s dg uns(3) dim(24);
  dcl-s n int(10);
  dcl-s k int(10);
  n = 0;
  dou v = 0;
    n += 1;
    dg(n) = %rem(v : 10);
    v = %div(v : 10);
  enddo;
  for k = n downto 1;
    wb = 48 + dg(k);
    wp += 1;
  endfor;
end-proc;

dcl-proc putInt;
  dcl-pi *n;
    v int(20) value;
  end-pi;
  dcl-s uv uns(20);
  if v < 0;
    wb = 45;
    wp += 1;
    if v = -9223372036854775807 - 1;
      uv = 9223372036854775807;
      uv += 1;
    else;
      uv = -v;
    endif;
    putUns(uv);
  else;
    putUns(v);
  endif;
end-proc;

// integer part + num/den as up to 10 decimals
dcl-proc putFrac;
  dcl-pi *n;
    ipart int(20) value;
    num   int(20) value;
    den   int(20) value;
  end-pi;
  dcl-s f int(20);
  dcl-s nd int(10);
  dcl-s k int(10);
  dcl-s dg uns(3) dim(10);
  f = %div(num * 10000000000 + %div(den : 2) : den);
  if f >= 10000000000;
    ipart += 1;
    f -= 10000000000;
  endif;
  putUns(ipart);
  if f = 0;
    return;
  endif;
  for k = 10 downto 1;
    dg(k) = %rem(f : 10);
    f = %div(f : 10);
  endfor;
  nd = 10;
  dow nd > 0;
    if dg(nd) <> 0;
      leave;
    endif;
    nd -= 1;
  enddo;
  wb = 46;
  wp += 1;
  for k = 1 to nd;
    wb = 48 + dg(k);
    wp += 1;
  endfor;
end-proc;

// packed decimal at p -> '-123.45' (no leading zeros, no trailing fraction zeros)
dcl-proc putPacked;
  dcl-pi *n;
    p     pointer value;
    prec  int(10) value;
    scale int(10) value;
  end-pi;
  dcl-s pb pointer;
  dcl-s b uns(3) based(pb);
  dcl-s dg uns(3) dim(64);
  dcl-s nb int(10);
  dcl-s nd int(10);
  dcl-s k int(10);
  dcl-s intD int(10);
  dcl-s lastD int(10);
  dcl-s sg int(10);
  dcl-s zero ind;
  dcl-s started ind;
  nb = %div(prec : 2) + 1;
  nd = nb * 2 - 1;
  zero = *on;
  for k = 1 to nd;
    pb = p + %div(k - 1 : 2);
    if %rem(k - 1 : 2) = 0;
      dg(k) = %div(b : 16);
    else;
      dg(k) = %rem(b : 16);
    endif;
    if dg(k) > 9;
      dg(k) = 0;
    endif;
    if dg(k) <> 0;
      zero = *off;
    endif;
  endfor;
  pb = p + nb - 1;
  sg = %rem(b : 16);
  if (sg = 13 or sg = 11) and not zero;
    wb = 45;
    wp += 1;
  endif;
  intD = nd - scale;
  started = *off;
  if intD <= 0;
    wb = 48;
    wp += 1;
  endif;
  for k = 1 to intD;
    if dg(k) <> 0 or started or k = intD;
      started = *on;
      wb = 48 + dg(k);
      wp += 1;
    endif;
  endfor;
  lastD = nd;
  dow lastD > intD;
    if dg(lastD) <> 0;
      leave;
    endif;
    lastD -= 1;
  enddo;
  if lastD > intD;
    wb = 46;
    wp += 1;
    for k = intD + 1 to lastD;
      if k >= 1;
        wb = 48 + dg(k);
        wp += 1;
      endif;
    endfor;
  endif;
end-proc;

// UTF-8 text with XML escaping (control characters -> blank, CR dropped)
// '_' starts _xHHHH_ (would be read back as an escape)?
dcl-proc escUnd;
  dcl-pi *n ind;
    p pointer value;
    n int(10) value;
  end-pi;
  dcl-s c uns(3) based(pc);
  dcl-s pc pointer;
  dcl-s k int(10);
  if n < 7;
    return *off;
  endif;
  pc = p + 1;
  if c <> 120;                          // x
    return *off;
  endif;
  for k = 2 to 5;
    pc = p + k;
    if not ((c >= 48 and c <= 57) or (c >= 65 and c <= 70) or (c >= 97 and c <= 102));
      return *off;
    endif;
  endfor;
  pc = p + 6;
  return (c = 95);
end-proc;

dcl-proc putHexEsc;                     // _x00HH_
  dcl-pi *n;
    c uns(3) value;
  end-pi;
  dcl-c HX '0123456789ABCDEF';
  dcl-s t like(TOK_T);                  // assignment converts to UTF-8
  t = '_x00' + %subst(HX : %div(c : 16) + 1 : 1) + %subst(HX : %rem(c : 16) + 1 : 1) + '_';
  putT(t);
end-proc;

dcl-proc putEsc;
  dcl-pi *n;
    p pointer value;
    n int(10) value;
  end-pi;
  dcl-s pb pointer;
  dcl-s c uns(3) based(pb);
  dcl-s i int(10);
  dcl-s st int(10);
  pb = p;
  st = 0;
  for i = 0 to n - 1;
    if CLS(c + 1) <> 0;
      if i > st;
        memcpy(wp : p + st : i - st);
        wp += i - st;
      endif;
      select;
      when c = 38;
        putT(tAmp);
      when c = 60;
        putT(tLt);
      when c = 62;
        putT(tGt);
      when c = 95;                     // '_': escape only a literal _xHHHH_
        if escUnd(p + i : n - i);
          putHexEsc(95);
        else;
          wb = 95;
          wp += 1;
        endif;
      other;                           // control character: _x000D_ like Excel
        putHexEsc(c);
      endsl;
      st = i + 1;
    endif;
    pb += 1;
  endfor;
  if n > st;
    memcpy(wp : p + st : n - st);
    wp += n - st;
  endif;
end-proc;

dcl-proc setRowDig;                      // rowNum -> rowDig (UTF-8 digits)
  dcl-s v int(20);
  dcl-s dg uns(3) dim(20);
  dcl-s n int(10);
  dcl-s k int(10);
  dcl-s p pointer;
  dcl-s b uns(3) based(p);
  v = rowNum;
  n = 0;
  dou v = 0;
    n += 1;
    dg(n) = 48 + %rem(v : 10);
    v = %div(v : 10);
  enddo;
  %len(rowDig) = n;
  p = %addr(rowDig : *data);
  for k = n downto 1;
    b = dg(k);
    p += 1;
  endfor;
end-proc;

// =====================================================================================
// SQL
// =====================================================================================
dcl-proc sqlFail;
  dcl-pi *n;
    what varchar(50) const;
  end-pi;
  dcl-s sst char(5);
  dcl-s mt varchar(512);
  sst = sqlstate;
  exec sql get diagnostics condition 1 :mt = MESSAGE_TEXT;
  mt = %trimr(mt);
  if %len(mt) > 1;
    if %subst(mt : %len(mt) : 1) = '.';
      mt = %subst(mt : 1 : %len(mt) - 1);
    endif;
  endif;
  fail(what + ' SQLSTATE ' + sst + ': ' + mt);
end-proc;

// does Db2 honour a CCSID 1208 override in SQLNAME on FETCH? (then strings come as UTF-8)
dcl-proc probeUtf8;
  dcl-pi *n ind;
  end-pi;
  dcl-s buf char(16) ccsid(*hex);
  dcl-s nb int(5);
  dcl-s ok ind;
  ok = *off;
  gStmt = 'SELECT CAST(''A'' AS VARCHAR(4) CCSID 37) FROM SYSIBM.SYSDUMMY1';
  exec sql prepare XLS0 from :gStmt;
  if sqlcode < 0;
    return *off;
  endif;
  SQLN = SQL_NUM;
  exec sql describe XLS0 into :SQLDA;
  if sqlcode < 0 or SQLD <> 1;
    return *off;
  endif;
  pv = %addr(SQL_VAR(1));
  sv.vType = 449;
  sv.vLen = 8;
  sv.vData = %addr(buf);
  sv.vInd = %addr(nb);
  sv.vNameLen = 8;
  sv.vName = x'000004B800000000';
  exec sql declare XLC0 cursor for XLS0;
  exec sql open XLC0;
  if sqlcode >= 0;
    exec sql fetch next from XLC0 using descriptor :SQLDA;
    ok = (sqlcode = 0 and %subst(buf : 3 : 1) = x'41');
    exec sql close XLC0;
  endif;
  return ok;
end-proc;

dcl-proc fmtLabel;                       // 3 x 20 column heading -> one line
  dcl-pi *n varchar(62);
    l varchar(60) const;
  end-pi;
  dcl-s r varchar(62);
  dcl-s f char(60);
  dcl-s k int(10);
  f = l;
  r = '';
  for k = 0 to 2;
    if %subst(f : k * 20 + 1 : 20) <> '';
      if r <> '';
        r += ' ';
      endif;
      r += %trim(%subst(f : k * 20 + 1 : 20));
    endif;
  endfor;
  return r;
end-proc;

// heading, text and base column of every result column (SQL descriptor + SYSCOLUMNS)
dcl-proc colMeta;
  dcl-pi *n;
    n int(10) value;
  end-pi;
  dcl-s i int(10);
  dcl-s j int(10);
  dcl-s lbl varchar(60);
  dcl-s bs varchar(128);
  dcl-s bt varchar(128);
  dcl-s bc varchar(128);
  dcl-s cn varchar(128);
  dcl-s ctx varchar(50);
  dcl-s chd varchar(60);
  dcl-s tS varchar(128) dim(1000);
  dcl-s tT varchar(128) dim(1000);
  dcl-s tC varchar(128) dim(1000);
  dcl-s done ind dim(1000);
  gLbl = '';
  gTxt = '';
  gBase = '';
  exec sql deallocate descriptor 'XLSXD';
  exec sql allocate descriptor 'XLSXD' with max 1000;
  exec sql describe XLS1 using sql descriptor 'XLSXD';
  if sqlcode < 0;
    return;
  endif;
  for i = 1 to n;
    exec sql get descriptor 'XLSXD' value :i :lbl = DB2_LABEL,
             :bs = DB2_BASE_SCHEMA_NAME, :bt = DB2_BASE_TABLE_NAME,
             :bc = DB2_BASE_COLUMN_NAME;
    if sqlcode < 0;
      leave;
    endif;
    gLbl(i) = lbl;
    tS(i) = %trim(bs);
    tT(i) = %trim(bt);
    tC(i) = %trim(bc);
    if tT(i) <> '' and tC(i) <> '';
      gBase(i) = tS(i) + '.' + tT(i) + '.' + tC(i);
    endif;
  endfor;
  exec sql deallocate descriptor 'XLSXD';

  // one catalog query per distinct base table
  done = *off;
  for i = 1 to n;
    if done(i) or tT(i) = '' or tC(i) = '';
      iter;
    endif;
    bs = tS(i);
    bt = tT(i);
    exec sql declare XLC2 cursor for
             select column_name, coalesce(column_text, ''), coalesce(column_heading, '')
               from qsys2.syscolumns
              where table_schema = :bs and table_name = :bt;
    exec sql open XLC2;
    dow sqlcode >= 0;
      exec sql fetch next from XLC2 into :cn, :ctx, :chd;
      if sqlcode <> 0;
        leave;
      endif;
      for j = i to n;
        if not done(j) and tS(j) = bs and tT(j) = bt and tC(j) = %trim(cn);
          gTxt(j) = %trim(ctx);
        endif;
      endfor;
    enddo;
    exec sql close XLC2;
    for j = i to n;
      if tS(j) = bs and tT(j) = bt;
        done(j) = *on;
      endif;
    endfor;
  endfor;
end-proc;

dcl-proc typeDesc;                       // SQL type text for the notes
  dcl-pi *n varchar(40);
    bt   int(10) value;
    len  int(10) value;
    prec int(10) value;
    scl  int(10) value;
    cc   int(10) value;
  end-pi;
  dcl-s r varchar(40);
  select;
  when bt = 452;
    r = 'CHAR(' + %char(len) + ')';
  when bt = 448 or bt = 456;
    r = 'VARCHAR(' + %char(len) + ')';
  when bt = 468;
    r = 'GRAPHIC(' + %char(len) + ')';
  when bt = 464 or bt = 472;
    r = 'VARGRAPHIC(' + %char(len) + ')';
  when bt = 408;
    r = 'CLOB';
  when bt = 412;
    r = 'DBCLOB';
  when bt = 988;
    r = 'XML';
  when bt = 500;
    r = 'SMALLINT';
  when bt = 496;
    r = 'INTEGER';
  when bt = 492;
    r = 'BIGINT';
  when bt = 484;
    r = 'DECIMAL(' + %char(prec) + ',' + %char(scl) + ')';
  when bt = 488;
    r = 'NUMERIC(' + %char(prec) + ',' + %char(scl) + ')';
  when bt = 480 and len = 4;
    r = 'REAL';
  when bt = 480;
    r = 'DOUBLE';
  when bt = 996;
    r = 'DECFLOAT';
  when bt = 384;
    r = 'DATE';
  when bt = 388;
    r = 'TIME';
  when bt = 392;
    r = 'TIMESTAMP';
  when bt = 912;
    r = 'BINARY(' + %char(len) + ')';
  when bt = 908;
    r = 'VARBINARY(' + %char(len) + ')';
  when bt = 404;
    r = 'BLOB';
  other;
    r = 'SQLTYPE ' + %char(bt);
  endsl;
  if cc = 65535 and (bt = 452 or bt = 448 or bt = 456);
    r += ' FOR BIT DATA';
  endif;
  return r;
end-proc;

// prepare + describe + bind the SQLDA to the row buffer + open the cursor
dcl-proc openSource;
  dcl-pi *n ind;
  end-pi;
  dcl-s i int(10);
  dcl-s off int(10);
  dcl-s sz int(10);
  dcl-s bt int(10);
  dcl-s cc uns(5);
  dcl-s pcc pointer;
  dcl-s ccv uns(5) based(pcc);
  dcl-s nm varchar(30);
  dcl-s h varchar(256);
  dcl-s w int(10);
  dcl-s hw int(10);
  dcl-s flen int(10);
  dcl-s up varchar(20);
  dcl-s obj varchar(5000);
  dcl-s mbr varchar(10);
  dcl-s p int(10);
  dcl-s q int(10);
  dcl-s lst varchar(5000);
  dcl-s k int(10);

  // ---- statement
  up = %xlate('abcdefghijklmnopqrstuvwxyz' : 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
              : %subst(gSrc + '                    ' : 1 : 7));
  if %subst(up : 1 : 7) = 'SELECT ' or %subst(up : 1 : 5) = 'WITH '
     or %subst(up : 1 : 7) = 'VALUES ' or %subst(up : 1 : 1) = '('
     or %scan(' ' : %trim(gSrc)) > 0 and %scan('(' : %trim(gSrc)) = 0;
    gIsFile = *off;
    gStmt = gSrc;
    gDftSheet = 'Sheet1';
  else;
    gIsFile = *on;
    obj = %trim(gSrc);
    mbr = '';
    p = %scan('(' : obj);
    if p > 0;
      q = %scan(')' : obj : p);
      if q = 0;
        q = %len(obj) + 1;
      endif;
      if q - p - 1 > 0;
        mbr = %trim(%subst(obj : p + 1 : q - p - 1));
      endif;
      obj = %trim(%subst(obj : 1 : p - 1));
    endif;
    p = %scan('/' : obj);
    if p = 0;
      p = %scan('.' : obj);
    endif;
    if p > 0 and p < %len(obj);
      gDftSheet = %subst(obj : p + 1);
    else;
      gDftSheet = obj;
    endif;
    if mbr <> '' and mbr <> '*FIRST';
      gDdl = 'DROP ALIAS QTEMP/XLSX00A';
      exec sql execute immediate :gDdl;
      gDdl = 'CREATE ALIAS QTEMP/XLSX00A FOR ' + obj + ' (' + mbr + ')';
      exec sql execute immediate :gDdl;
      if sqlcode < 0;
        sqlFail('CREATE ALIAS for member ' + mbr);
        return *off;
      endif;
      gStmt = 'SELECT * FROM QTEMP/XLSX00A';
      gDftSheet = mbr;
    else;
      gStmt = 'SELECT * FROM ' + obj;
    endif;
  endif;

  // ---- custom headings
  gCus = '';
  lst = gHdrList;
  k = 0;
  dow %len(lst) > 0 and k < 1000;
    k += 1;
    p = %scan(';' : lst);
    if p = 0;
      gCus(k) = %trim(lst);
      lst = '';
    else;
      if p > 1;
        gCus(k) = %trim(%subst(lst : 1 : p - 1));
      endif;
      if p >= %len(lst);
        lst = '';
      else;
        lst = %subst(lst : p + 1);
      endif;
    endif;
  enddo;

  // ---- prepare / describe
  exec sql prepare XLS1 from :gStmt;
  if sqlcode < 0;
    sqlFail('PREPARE');
    return *off;
  endif;
  SQLN = SQL_NUM;
  exec sql describe XLS1 into :SQLDA using names;
  if sqlcode < 0;
    sqlFail('DESCRIBE');
    return *off;
  endif;
  if SQLD = 0;
    fail('the statement returns no columns (SELECT, WITH or VALUES expected)');
    return *off;
  endif;
  if SQLD > SQL_NUM;
    fail('too many columns: ' + %char(SQLD) + ' (max ' + %char(SQL_NUM) + ')');
    return *off;
  endif;
  nCol = SQLD;
  gLbl = '';
  gTxt = '';
  gBase = '';
  if gHdr and (gHdrMode = 'L' or gHdrMode = 'T' or gDesc <> 'N');
    colMeta(nCol);
  endif;

  // ---- columns
  off = 0;
  gMulti = *on;
  for i = 1 to nCol;
    pv = %addr(SQL_VAR(i));
    pcc = pv + 18;                       // CCSID: bytes 3-4 of SQLDATA after DESCRIBE
    cc = ccv;
    bt = sv.vType - %rem(sv.vType : 2);
    nm = '';
    if sv.vNameLen > 0 and sv.vNameLen <= 30;
      nm = %trimr(%subst(sv.vName : 1 : sv.vNameLen));
    endif;
    col(i).prec = 0;
    col(i).scale = 0;
    col(i).dlen = sv.vLen;
    col(i).tdesc = typeDesc(bt : sv.vLen : sv.vP : sv.vS : cc);
    col(i).cnm = nm;
    select;
    when bt = 452 or bt = 448 or bt = 456;            // CHAR VARCHAR LONG VARCHAR
      if cc = 65535 or not gUtf8Ok;
        col(i).kind = 'E';
        flen = sv.vLen;
      else;
        col(i).kind = 'S';
        flen = sv.vLen * 3;
      endif;
    when bt = 468 or bt = 464 or bt = 472;            // GRAPHIC VARGRAPHIC
      if gUtf8Ok;
        col(i).kind = 'S';
        flen = sv.vLen * 3;
      else;
        col(i).kind = 'E';
        flen = sv.vLen * 2;
      endif;
    when bt = 408 or bt = 412 or bt = 988;            // CLOB DBCLOB XML
      gMulti = *off;                                  // no block fetch with LOBs
      col(i).dlen = 50;
      flen = 32740;
      if gUtf8Ok;
        col(i).kind = 'S';
      else;
        col(i).kind = 'E';
      endif;
    when bt = 500 or bt = 496 or bt = 492;            // SMALLINT INTEGER BIGINT
      col(i).kind = 'I';
      flen = 8;
      if bt = 500;
        col(i).dlen = 6;
      elseif bt = 496;
        col(i).dlen = 11;
      else;
        col(i).dlen = 15;
      endif;
    when bt = 484 or bt = 488;                        // DECIMAL NUMERIC
      col(i).kind = 'P';
      col(i).prec = sv.vP;
      col(i).scale = sv.vS;
      flen = %div(col(i).prec : 2) + 1;
    when bt = 480 or bt = 996;                        // FLOAT DECFLOAT
      col(i).kind = 'F';
      flen = 8;
    when bt = 384;
      col(i).kind = 'D';
      flen = 10;
    when bt = 388;
      col(i).kind = 'T';
      flen = 8;
    when bt = 392;
      col(i).kind = 'Z';
      flen = sv.vLen;
      if flen < 19;
        flen = 26;
      endif;
    when bt = 912 or bt = 908;                        // BINARY VARBINARY
      col(i).kind = 'H';
      flen = sv.vLen;
      if flen > 16000;
        flen = 16000;
      endif;
    other;                                            // BLOB ROWID DATALINK ...
      if bt = 404 or bt = 396;
        gMulti = *off;
      endif;
      col(i).kind = 'X';
      flen = 1;
    endsl;
    if flen > 32740;
      flen = 32740;
    endif;
    if flen < 1;
      flen = 1;
    endif;
    col(i).len = flen;
    select;
    when col(i).kind = 'S' or col(i).kind = 'E';
      sv.vType = 449;
      sv.vLen = flen;
      sz = flen + 2;
    when col(i).kind = 'H' or col(i).kind = 'X';
      sv.vType = 909;
      sv.vLen = flen;
      sz = flen + 2;
    when col(i).kind = 'I';
      sv.vType = 493;
      sv.vLen = 8;
      sz = 8;
    when col(i).kind = 'P';
      sv.vType = 485;
      sz = flen;
    when col(i).kind = 'F';
      sv.vType = 481;
      sv.vLen = 8;
      sz = 8;
    when col(i).kind = 'D';
      sv.vType = 385;
      sv.vLen = 10;
      sz = 10;
    when col(i).kind = 'T';
      sv.vType = 389;
      sv.vLen = 8;
      sz = 8;
    other;
      sv.vType = 393;
      sv.vLen = flen;
      sz = flen;
    endsl;
    col(i).off = off;                      // rows are packed, no alignment
    off += sz;

    // heading and width
    h = '';
    if gCus(i) <> '';
      h = gCus(i);
    elseif gHdrMode = 'T' and gTxt(i) <> '';
      h = gTxt(i);
    elseif (gHdrMode = 'L' or gHdrMode = 'T') and gLbl(i) <> '';
      h = fmtLabel(gLbl(i));
    endif;
    if h = '';
      h = nm;
    endif;
    if h = '';
      h = 'Col' + %char(i);
    endif;
    col(i).hdr = xesc(u(h));
    col(i).hlen = %len(h);
    %len(col(i).dsc) = 0;                // 2nd header row: the text, or the
    if gTxt(i) <> '' and gTxt(i) <> h;   // name when the heading is the text
      col(i).dsc = xesc(u(gTxt(i)));
    elseif nm <> '' and nm <> h;
      col(i).dsc = xesc(u(nm));
    endif;
    col(i).ref = colRef(i);
    select;
    when col(i).kind = 'S' or col(i).kind = 'E';
      w = col(i).dlen;
    when col(i).kind = 'H';
      w = col(i).dlen * 2;
    when col(i).kind = 'I';
      w = col(i).dlen;
    when col(i).kind = 'P';
      w = col(i).prec + 1 + %div(col(i).prec - col(i).scale : 3);
      if col(i).scale > 0;
        w += 1;
      endif;
    when col(i).kind = 'F';
      w = 14;
    when col(i).kind = 'D';
      w = 11;
    when col(i).kind = 'T';
      w = 9;
    when col(i).kind = 'Z';
      w = 19;
    other;
      w = 8;
    endsl;
    hw = col(i).hlen + 2;
    if gFilter;
      hw += 2;
    endif;
    if gHdr and hw > w;
      w = hw;
    endif;
    if gHdr and (gDesc = 'R' or gDesc = 'B') and %len(gTxt(i)) > w;
      w = %len(gTxt(i));
      if w > 30;
        w = 30;
      endif;
    endif;
    if w < 6;
      w = 6;
    endif;
    if w > 60;
      w = 60;
    endif;
    col(i).width = w + 0.7;
  endfor;

  // ---- row storage area (block of rows), pointers to row 1, CCSID overrides
  gRowLen = off;
  if gMulti;
    gBlkRows = %div(15000000 : gRowLen);
    if gBlkRows > 1000;
      gBlkRows = 1000;
    endif;
    if gBlkRows < 1;
      gBlkRows = 1;
    endif;
  else;
    gBlkRows = 1;
  endif;
  pBlk = %alloc(gBlkRows * gRowLen + 16);
  pBInd = %alloc(gBlkRows * nCol * 2 + 16);
  for i = 1 to nCol;
    pv = %addr(SQL_VAR(i));
    sv.vData = pBlk + col(i).off;
    sv.vInd = pBInd + (i - 1) * 2;
    if col(i).kind = 'S';
      sv.vNameLen = 8;
      sv.vName = x'000004B800000000';
    else;
      sv.vNameLen = 0;
    endif;
  endfor;

  exec sql declare XLC1 cursor for XLS1;
  exec sql open XLC1;
  if sqlcode < 0;
    sqlFail('OPEN');
    return *off;
  endif;
  gCurOpen = *on;
  return *on;
end-proc;

dcl-proc closeCursor;
  if gCurOpen;
    exec sql close XLC1;
    gCurOpen = *off;
  endif;
end-proc;

// =====================================================================================
// rows
// =====================================================================================
dcl-proc dg2;                            // 2 EBCDIC digits at p
  dcl-pi *n int(10);
    p pointer value;
  end-pi;
  dcl-s b uns(3) based(p);
  dcl-s v int(10);
  v = (b - 240) * 10;
  p += 1;
  return v + b - 240;
end-proc;

dcl-proc writeHeader;
  dcl-pi *n ind;
  end-pi;
  dcl-s i int(10);
  dcl-s hs like(TOK_T);
  gHdrRows = 0;
  rowNum += 1;
  gHdrRows += 1;
  setRowDig();
  hs = numU(ST(1));
  if not need(64);
    return *off;
  endif;
  putT(tRowA);
  putT(rowDig);
  putT(tRowB);
  for i = 1 to nCol;
    if not need(%len(col(i).hdr) + 120);
      return *off;
    endif;
    putT(tCA);
    putT(col(i).ref);
    putT(rowDig);
    putT(tCS2);
    putT(hs);
    putT(tCS);
    memcpy(wp : %addr(col(i).hdr : *data) : %len(col(i).hdr));
    wp += %len(col(i).hdr);
    putT(tCSE);
  endfor;
  if not need(16);
    return *off;
  endif;
  putT(tRowE);
  if gDesc <> 'R' and gDesc <> 'B';
    return *on;
  endif;

  // second header row: column text (cells styled even when empty, for the band)
  rowNum += 1;
  gHdrRows += 1;
  setRowDig();
  hs = numU(ST(11));
  if not need(64);
    return *off;
  endif;
  putT(tRowA);
  putT(rowDig);
  putT(tRowB);
  for i = 1 to nCol;
    if not need(%len(col(i).dsc) + 120);
      return *off;
    endif;
    putT(tCA);
    putT(col(i).ref);
    putT(rowDig);
    putT(tCS2);
    putT(hs);
    if %len(col(i).dsc) = 0;
      putT(tCV0);
    else;
      putT(tCS);
      memcpy(wp : %addr(col(i).dsc : *data) : %len(col(i).dsc));
      wp += %len(col(i).dsc);
      putT(tCSE);
    endif;
  endfor;
  if not need(16);
    return *off;
  endif;
  putT(tRowE);
  return *on;
end-proc;

// fetch all rows and write them from sheet row (rowNum + 1) on
dcl-proc fetchRows;
  dcl-pi *n ind;
  end-pi;
  dcl-s i int(10);
  dcl-s pc pointer;
  dcl-s hv int(5) based(pc);
  dcl-s iv int(20) based(pc);
  dcl-s fv float(8) based(pc);
  dcl-s pd pointer;
  dcl-s bd uns(3) based(pd);
  dcl-s ebs char(32740) based(pd);
  dcl-s vl int(10);
  dcl-s k int(10);
  dcl-s y int(10);
  dcl-s m int(10);
  dcl-s d int(10);
  dcl-s ser int(10);
  dcl-s secs int(20);
  dcl-s ms int(20);
  dcl-s tok like(TOK_T);
  dcl-s isStr ind;
  dcl-s pre ind;
  dcl-s hx int(10);
  dcl-s t0 timestamp;
  dcl-s nb int(10);
  dcl-s r int(10);
  dcl-s lastBlk ind;
  dcl-s rowBase pointer;
  dcl-s indBase pointer;
  dcl-s pIi pointer;
  dcl-s indv int(5) based(pIi);

  lastBlk = *off;
  dow not lastBlk;
    if gStats;
      t0 = %timestamp();
    endif;
    if gMulti;
      exec sql fetch next from XLC1 for :gBlkRows rows using descriptor :SQLDA
               into :gRowArea indicator :gIndArea;
    else;
      exec sql fetch next from XLC1 using descriptor :SQLDA;
    endif;
    if gStats;
      gFetchUs += %diff(%timestamp() : t0 : *mseconds);
    endif;
    if sqlcode = 100;
      leave;
    endif;
    if sqlcode < 0;
      sqlFail('FETCH');
      return *off;
    endif;
    nb = 1;
    if gMulti;
      nb = SQLERRD(3);
      if SQLERRD(5) = 100;
        lastBlk = *on;
      endif;
      if nb > 0 and SQLERRD(4) <> gRowLen;
        fail('row layout mismatch: Db2 row length ' + %char(SQLERRD(4))
             + ', computed ' + %char(gRowLen));
        return *off;
      endif;
    endif;
  for r = 1 to nb;
    if rowNum >= XMAXROW;
      gTrunc = *on;
      lastBlk = *on;
      leave;
    endif;
    rowBase = pBlk + (r - 1) * gRowLen;
    indBase = pBInd + (r - 1) * nCol * 2;
    rowNum += 1;
    gRows += 1;
    setRowDig();
    if not need(64);
      return *off;
    endif;
    putT(tRowA);
    putT(rowDig);
    putT(tRowB);
    for i = 1 to nCol;
      pIi = indBase + (i - 1) * 2;
      if indv < 0;
        iter;
      endif;
      pc = rowBase + col(i).off;
      select;

      // ---- text
      when col(i).kind = 'S' or col(i).kind = 'E';
        vl = hv;
        pd = pc + 2;
        if col(i).kind = 'S';
          dow vl > 0;
            pd = pc + 1 + vl;
            if bd <> 32;
              leave;
            endif;
            vl -= 1;
          enddo;
          pd = pc + 2;
        else;
          dow vl > 0;
            pd = pc + 1 + vl;
            if bd <> 64;
              leave;
            endif;
            vl -= 1;
          enddo;
          pd = pc + 2;
          if vl > 0;
            gU = %subst(ebs : 1 : vl);
            vl = %len(gU);
            if vl > 32767;
              vl = 32767;
              pd = %addr(gU : *data) + vl;
              dow vl > 0 and bd >= 128 and bd < 192;
                vl -= 1;
                pd -= 1;
              enddo;
            endif;
            pd = %addr(gU : *data);
          endif;
        endif;
        if vl = 0;
          iter;
        endif;
        if not need(vl * 7 + 128);         // _x000D_ = 7 bytes
          return *off;
        endif;
        pre = *off;
        if bd = 32 or bd = 9 or bd = 10;
          pre = *on;
        endif;
        memcpy(wp : %addr(col(i).cPre : *data) : %len(col(i).cPre));
        wp += %len(col(i).cPre);
        memcpy(wp : %addr(rowDig : *data) : %len(rowDig));
        wp += %len(rowDig);
        if pre;
          putT(tCSP);
        else;
          putT(tCS);
        endif;
        putEsc(pd : vl);
        putT(tCSE);
        iter;

      // ---- hex
      when col(i).kind = 'H';
        vl = hv;
        if vl = 0;
          iter;
        endif;
        if not need(vl * 2 + 128);
          return *off;
        endif;
        putT(tCA);
        putT(col(i).ref);
        putT(rowDig);
        putT(tCS);
        pd = pc + 2;
        for k = 1 to vl;
          hx = %div(bd : 16);
          if hx < 10;
            wb = 48 + hx;
          else;
            wb = 55 + hx;
          endif;
          wp += 1;
          hx = %rem(bd : 16);
          if hx < 10;
            wb = 48 + hx;
          else;
            wb = 55 + hx;
          endif;
          wp += 1;
          pd += 1;
        endfor;
        putT(tCSE);
        iter;

      when col(i).kind = 'X';
        iter;
      endsl;

      // ---- numbers, dates
      if not need(160);
        return *off;
      endif;
      isStr = *off;
      if col(i).kind = 'D' or col(i).kind = 'Z';
        pd = pc;
        y = dg2(pc) * 100 + dg2(pc + 2);
        if y < 1900;                     // not representable: keep the text
          isStr = *on;
        endif;
      endif;
      memcpy(wp : %addr(col(i).cPre : *data) : %len(col(i).cPre));
      wp += %len(col(i).cPre);
      memcpy(wp : %addr(rowDig : *data) : %len(rowDig));
      wp += %len(rowDig);
      if isStr;
        putT(tCS);
        if col(i).kind = 'D';
          tok = %subst(ebs : 1 : 10);
        else;
          tok = %subst(ebs : 1 : 19);
        endif;
        putT(tok);
        putT(tCSE);
        iter;
      endif;
      memcpy(wp : %addr(col(i).cPost : *data) : %len(col(i).cPost));
      wp += %len(col(i).cPost);
      select;
      when col(i).kind = 'I';
        putInt(iv);
      when col(i).kind = 'P';
        putPacked(pc : col(i).prec : col(i).scale);
      when col(i).kind = 'F';
        tok = fmtFloat(fv);
        putT(tok);
      when col(i).kind = 'D';
        m = dg2(pc + 5);
        d = dg2(pc + 8);
        ser = dfc(y : m : d) - excelBase;
        if ser < 61;
          ser -= 1;
        endif;
        putUns(ser);
      when col(i).kind = 'T';
        secs = dg2(pc) * 3600 + dg2(pc + 3) * 60 + dg2(pc + 6);
        putFrac(0 : secs : 86400);
      when col(i).kind = 'Z';
        m = dg2(pc + 5);
        d = dg2(pc + 8);
        ser = dfc(y : m : d) - excelBase;
        if ser < 61;
          ser -= 1;
        endif;
        secs = dg2(pc + 11) * 3600 + dg2(pc + 14) * 60 + dg2(pc + 17);
        ms = 0;
        if col(i).len >= 23;
          ms = dg2(pc + 20) * 10 + %rem(dg2(pc + 21) : 10);
        endif;
        putFrac(ser : secs * 1000 + ms : 86400000);
      endsl;
      putT(tCVE);
    endfor;
    if not need(16);
      return *off;
    endif;
    putT(tRowE);
  endfor;
  enddo;
  return *on;
end-proc;

// float -> '-1.25E+2' (as Excel writes it): no '+' sign, no trailing zeros
dcl-proc fmtFloat;
  dcl-pi *n varchar(40);
    f float(8) value;
  end-pi;
  dcl-s s varchar(40);
  dcl-s man varchar(40);
  dcl-s sg varchar(1);
  dcl-s e int(10);
  dcl-s x int(10);
  dcl-s k int(10);
  s = %trim(%char(f));
  sg = '';
  if %subst(s : 1 : 1) = '-';
    sg = '-';
  endif;
  if %subst(s : 1 : 1) = '-' or %subst(s : 1 : 1) = '+';
    s = %subst(s : 2);
  endif;
  e = %scan('E' : s);
  if e = 0;
    return sg + s;
  endif;
  man = %subst(s : 1 : e - 1);
  x = %int(%subst(s : e + 1));
  if %scan('.' : man) > 0;
    k = %len(man);
    dow k > 1 and %subst(man : k : 1) = '0';
      k -= 1;
    enddo;
    if %subst(man : k : 1) = '.';
      k -= 1;
    endif;
    man = %subst(man : 1 : k);
  endif;
  if x = 0 or man = '0';
    return sg + man;
  endif;
  if x > 0;
    return sg + man + 'E+' + %char(x);
  endif;
  return sg + man + 'E' + %char(x);
end-proc;

dcl-proc setColStyles;
  dcl-s i int(10);
  for i = 1 to nCol;
    %len(col(i).styS) = 0;
    select;
    when col(i).kind = 'P' and col(i).scale >= 1 and col(i).scale <= 6;
      col(i).styS = numU(ST(4 + col(i).scale));
    when col(i).kind = 'D';
      col(i).styS = numU(ST(2));
    when col(i).kind = 'Z';
      col(i).styS = numU(ST(3));
    when col(i).kind = 'T';
      col(i).styS = numU(ST(4));
    endsl;
    col(i).cPre = tCA + col(i).ref;
    if %len(col(i).styS) > 0;
      col(i).cPost = tCS2 + col(i).styS + tCV;
    else;
      col(i).cPost = tCV;
    endif;
  endfor;
end-proc;

// =====================================================================================
// sheets
// =====================================================================================
// header notes of the new sheet: sheet relationships, comments part, legacy VML part
dcl-proc writeNotes;
  dcl-pi *n ind;
    part like(XMID_T) const;
  end-pi;
  dcl-s i int(10);
  dcl-s bold varchar(500);
  dcl-s rest varchar(500);
  dcl-c LF x'25';
  // the sheet relationships (comments, vmlDrawing) are written by writeSheetRels
  gTmp = u(XMLDECL + '<comments xmlns="' + NS_MAIN + '"><authors><author>XLSX00R'
       + '</author></authors><commentList>');
  for i = 1 to nCol;
    if gBase(i) <> '';
      bold = gBase(i);
    else;
      bold = col(i).cnm + ' (expression)';
    endif;
    rest = LF + col(i).tdesc;
    if gTxt(i) <> '';
      rest += LF + gTxt(i);
    endif;
    if gLbl(i) <> '' and fmtLabel(gLbl(i)) <> col(i).cnm;
      rest += LF + 'Heading: ' + fmtLabel(gLbl(i));
    endif;
    gTmp += u('<comment ref="') + col(i).ref + u('1" authorId="0"><text><r><rPr><b/>'
          + '<sz val="9"/><rFont val="Tahoma"/><family val="2"/></rPr><t>')
          + xesc(u(bold)) + u('</t></r><r><rPr><sz val="9"/><rFont val="Tahoma"/>'
          + '<family val="2"/></rPr><t xml:space="preserve">') + xesc(u(rest))
          + u('</t></r></text></comment>');
  endfor;
  gTmp += u('</commentList></comments>');
  if not writePart(u('xl/comments') + numU(gNoteN) + u('.xml') : gTmp);
    return *off;
  endif;

  gTmp = u('<xml xmlns:v="urn:schemas-microsoft-com:vml"'
       + ' xmlns:o="urn:schemas-microsoft-com:office:office"'
       + ' xmlns:x="urn:schemas-microsoft-com:office:excel">'
       + '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="')
       + numU(50 + gNoteN) + u('"/></o:shapelayout>'
       + '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"'
       + ' path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
       + '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>');
  for i = 1 to nCol;
    gTmp += u('<v:shape id="_x0000_s') + numU((50 + gNoteN) * 1024 + i)
          + u('" type="#_x0000_t202" style="position:absolute;margin-left:80pt;'
          + 'margin-top:2pt;width:180pt;height:70pt;z-index:') + numU(i)
          + u(';visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">'
          + '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>'
          + '<v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto">'
          + '<div style="text-align:left"></div></v:textbox><x:ClientData'
          + ' ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>')
          + numU(i) + u(', 15, 0, 2, ') + numU(i + 3) + u(', 15, 5, 2</x:Anchor>'
          + '<x:AutoFill>False</x:AutoFill><x:Row>0</x:Row><x:Column>')
          + numU(i - 1) + u('</x:Column></x:ClientData></v:shape>');
  endfor;
  gTmp += u('</xml>');
  return writePart(u('xl/drawings/vmlDrawing') + numU(gNoteN) + u('.vml') : gTmp);
end-proc;

dcl-proc writeNewSheet;
  dcl-pi *n ind;
    part like(XMID_T) const;
  end-pi;
  dcl-s i int(10);
  dcl-s hr int(10);
  dcl-s tail like(XMID_T);
  if zwBegin(part) < 0;
    fail(xlErr());
    return *off;
  endif;
  wp = *null;
  we = *null;
  hr = 0;
  if gHdr;
    hr = 1;
    if gDesc = 'R' or gDesc = 'B';
      hr = 2;
    endif;
  endif;
  gTmp = u(XMLDECL + '<worksheet xmlns="' + NS_MAIN + '" xmlns:r="' + NS_REL + '">');
  if hr > 0 and gFreeze;
    gTmp += u('<sheetViews><sheetView workbookViewId="0"><pane ySplit="') + numU(hr)
          + u('" topLeftCell="A') + numU(hr + 1) + u('" activePane="bottomLeft"'
          + ' state="frozen"/><selection pane="bottomLeft" activeCell="A') + numU(hr + 1)
          + u('" sqref="A') + numU(hr + 1) + u('"/></sheetView></sheetViews>');
  endif;
  gTmp += u('<sheetFormatPr defaultRowHeight="15"/><cols>');
  for i = 1 to nCol;
    gTmp += u('<col min="' + %char(i) + '" max="' + %char(i) + '" width="'
          + %trim(%char(col(i).width)) + '" customWidth="1"/>');
  endfor;
  gTmp += u('</cols><sheetData>');
  if not putBig(gTmp);
    return *off;
  endif;
  rowNum = 0;
  gRow0 = 1;
  if hr > 0;
    gRow0 = hr;                          // name / table: from the (last) header row
  endif;
  if gHdr;
    if gTbl;
      uniqueHeaders();
    endif;
    if not writeHeader();
      return *off;
    endif;
  endif;
  if not fetchRows();
    return *off;
  endif;
  tail = u('</sheetData>');
  if hr > 0 and gFilter and not gTbl;   // a table has its own autoFilter
    i = rowNum;
    if i < hr;
      i = hr;
    endif;
    tail += u('<autoFilter ref="A') + numU(hr) + cColon + col(nCol).ref + numU(i)
          + u('"/>');
  endif;
  tail += u('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75"'
        + ' header="0.3" footer="0.3"/>');
  if gNotes;
    tail += u('<legacyDrawing r:id="rIdXV"/>');
  endif;
  if gTbl;
    tail += u('<tableParts count="1"><tablePart r:id="rIdXT"/></tableParts>');
  endif;
  tail += u('</worksheet>');
  if not putMid(tail);
    return *off;
  endif;
  if zwEnd() < 0;
    fail(xlErr());
    return *off;
  endif;
  if not writeSheetRels(part);
    return *off;
  endif;
  if gNotes;
    if not writeNotes(part);
      return *off;
    endif;
  endif;
  if gTbl;
    return writeTable(rowNum);
  endif;
  return *on;
end-proc;

// byte search in memory, -1 = not found
dcl-proc memFind;
  dcl-pi *n int(20);
    p    pointer value;
    n    int(20) value;
    from int(20) value;
    ndc  like(TOK_T) const;
  end-pi;
  dcl-s pa pointer;
  dcl-s a char(64) ccsid(*hex) based(pa);
  dcl-s b uns(3) based(pa);
  dcl-s pn pointer;
  dcl-s z char(64) ccsid(*hex) based(pn);
  dcl-s f uns(3) based(pn);
  dcl-s i int(20);
  dcl-s m int(10);
  dcl-s nd like(TOK_T);
  nd = ndc;
  m = %len(nd);
  pn = %addr(nd : *data);
  for i = from to n - m;
    pa = p + i;
    if b = f;
      if %subst(a : 1 : m) = %subst(z : 1 : m);
        return i;
      endif;
    endif;
  endfor;
  return -1;
end-proc;

dcl-proc memRFind;                       // last occurrence starting before n - m
  dcl-pi *n int(20);
    p  pointer value;
    n  int(20) value;
    ndc like(TOK_T) const;
  end-pi;
  dcl-s pa pointer;
  dcl-s a char(64) ccsid(*hex) based(pa);
  dcl-s b uns(3) based(pa);
  dcl-s pn pointer;
  dcl-s z char(64) ccsid(*hex) based(pn);
  dcl-s f uns(3) based(pn);
  dcl-s i int(20);
  dcl-s m int(10);
  dcl-s nd like(TOK_T);
  nd = ndc;
  m = %len(nd);
  pn = %addr(nd : *data);
  for i = n - m downto 0;
    pa = p + i;
    if b = f;
      if %subst(a : 1 : m) = %subst(z : 1 : m);
        return i;
      endif;
    endif;
  endfor;
  return -1;
end-proc;

dcl-proc isRowTag;                       // '<row' at offset i followed by blank or '>'
  dcl-pi *n ind;
    p pointer value;
    i int(20) value;
  end-pi;
  dcl-s pc pointer;
  dcl-s c uns(3) based(pc);
  pc = p + i + 4;
  return (c = 32 or c = 62 or c = 10 or c = 13 or c = 9);
end-proc;

dcl-proc patchAutoFilter;                // gTmp holds the sheet tail
  dcl-pi *n;
    newLast int(20) value;
  end-pi;
  dcl-s a int(10);
  dcl-s m int(10);
  dcl-s k int(10);
  dcl-s tag like(XMID_T);
  dcl-s ref like(XMID_T);
  dcl-s nr like(XMID_T);
  a = findOpen(gTmp : u('autoFilter') : 1);
  if a = 0;
    return;
  endif;
  tag = tagText(gTmp : a);
  ref = attr(tag : u('ref'));
  k = %scan(cColon : ref);
  if k = 0 or k = %len(ref);
    return;
  endif;
  nr = %subst(ref : 1 : k) + noTrailDigits(%subst(ref : k + 1)) + numU(newLast);
  m = %scan(u(' ref="') : gTmp : a);
  if m > 0;
    gTmp = %replace(nr : gTmp : m + 6 : %len(ref));
  endif;
end-proc;

// the _xlnm._FilterDatabase name of sheet index sIdx (0-based) follows the new last row
dcl-proc patchFilterName;
  dcl-pi *n;
    sIdx    int(10) value;
    newLast int(20) value;
  end-pi;
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s c int(10);
  dcl-s k int(10);
  dcl-s j int(10);
  dcl-s tag like(XMID_T);
  dcl-s body like(XMID_T);
  p = 1;
  dow 1 = 1;
    p = findOpen(gWB : u('definedName') : p);
    if p = 0;
      leave;
    endif;
    e = tagEnd(gWB : p);
    if e = 0;
      leave;
    endif;
    tag = tagText(gWB : p);
    if attr(tag : u('name')) = u('_xlnm._FilterDatabase')
       and attr(tag : u('localSheetId')) = numU(sIdx);
      c = %scan(u('</definedName>') : gWB : e);
      if c > e and c - e < 30000;
        body = %subst(gWB : e : c - e);
        k = 0;
        j = %scan(cDol : body);
        dow j > 0;
          k = j;
          if j >= %len(body);
            leave;
          endif;
          j = %scan(cDol : body : j + 1);
        enddo;
        if k > 0 and k < %len(body);
          if digits(%subst(body : k + 1)) > 0;
            body = %subst(body : 1 : k) + numU(newLast);
            gWB = %replace(body : gWB : e : c - e);
            gNewWB = *on;
          endif;
        endif;
      endif;
      leave;
    endif;
    p = e;
  enddo;
end-proc;

// rows appended to an existing sheet (pX/nX = its XML)
dcl-proc writeAppendSheet;
  dcl-pi *n ind;
    part like(XMID_T) const;
    pX   pointer value;
    nX   int(20) value;
    sIdx int(10) value;
  end-pi;
  dcl-s sdEnd int(20);
  dcl-s selfC ind;
  dcl-s pr int(20);
  dcl-s lim int(20);
  dcl-s dm int(20);
  dcl-s de int(20);
  dcl-s lastR int(20);
  dcl-s sufS int(20);
  dcl-s sufL int(20);
  dcl-s k int(20);
  dcl-s tagU like(XMID_T);
  dcl-s r like(XMID_T);

  selfC = *off;
  sdEnd = memRFind(pX : nX : u('</sheetData>'));
  if sdEnd < 0;
    sdEnd = memFind(pX : nX : 0 : u('<sheetData/>'));
    selfC = *on;
    if sdEnd < 0;
      fail('sheet ' + eb(part) + ': <sheetData> not found');
      return *off;
    endif;
  endif;

  // last row number
  lastR = 0;
  lim = sdEnd;
  pr = -1;
  dow lim > 4;
    pr = memRFind(pX : lim : u('<row'));
    if pr < 0;
      leave;
    endif;
    if isRowTag(pX : pr);
      leave;
    endif;
    lim = pr + 3;
    pr = -1;
  enddo;
  if pr >= 0;
    k = memFind(pX : sdEnd : pr : u('>'));
    if k > pr and k - pr < 30000;
      %len(tagU) = k - pr + 1;
      memcpy(%addr(tagU : *data) : pX + pr : k - pr + 1);
      r = attr(tagU : u('r'));
      if %len(r) > 0;
        lastR = digits(r);
      endif;
    endif;
    if lastR = 0;                        // rows without r="": count them
      k = 0;
      dow 1 = 1;
        k = memFind(pX : sdEnd : k : u('<row'));
        if k < 0;
          leave;
        endif;
        if isRowTag(pX : k);
          lastR += 1;
        endif;
        k += 4;
      enddo;
    endif;
  endif;

  // <dimension .../> is optional and would be stale: drop it
  k = sdEnd;
  if k > 65536;
    k = 65536;
  endif;
  dm = memFind(pX : k : 0 : u('<dimension'));
  de = -1;
  if dm >= 0;
    de = memFind(pX : sdEnd : dm : u('>'));
  endif;

  if zwBegin(part) < 0;
    fail(xlErr());
    return *off;
  endif;
  wp = *null;
  we = *null;
  if dm >= 0 and de > dm;
    if not putRaw(pX : dm) or not putRaw(pX + de + 1 : sdEnd - de - 1);
      return *off;
    endif;
  else;
    if not putRaw(pX : sdEnd);
      return *off;
    endif;
  endif;
  if selfC;
    if not putMid(u('<sheetData>'));
      return *off;
    endif;
  endif;
  rowNum = lastR;
  gApLast = lastR;
  gRow0 = 1;
  if gHdr and lastR = 0;
    if not writeHeader();
      return *off;
    endif;
  endif;
  if not fetchRows();
    return *off;
  endif;
  if selfC;
    if not putMid(u('</sheetData>'));
      return *off;
    endif;
    sufS = sdEnd + 12;
  else;
    sufS = sdEnd;
  endif;
  sufL = nX - sufS;
  if sufL <= XBIGMAX;
    %len(gTmp) = sufL;
    if sufL > 0;
      memcpy(%addr(gTmp : *data) : pX + sufS : sufL);
    endif;
    patchAutoFilter(rowNum);
    if not putBig(gTmp);
      return *off;
    endif;
  else;
    if not putRaw(pX + sufS : sufL);
      return *off;
    endif;
  endif;
  if zwEnd() < 0;
    fail(xlErr());
    return *off;
  endif;
  patchFilterName(sIdx : rowNum);
  return *on;
end-proc;

// =====================================================================================
// workbook parts
// =====================================================================================
dcl-proc styBase;
  dcl-pi *n like(XMID_T);
  end-pi;
  return u(XMLDECL + '<styleSheet xmlns="' + NS_MAIN + '">'
    + '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/>'
    + '</font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill>'
    + '<fill><patternFill patternType="gray125"/></fill></fills>'
    + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border>'
    + '</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0"'
    + ' borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0"'
    + ' fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1">'
    + '<cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
end-proc;

dcl-proc baseParts;                      // a new, empty package
  gCT = u(XMLDECL + '<Types xmlns="' + NS_CT + '">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.'
      + 'relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="' + CT_OXML
      + 'spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/styles.xml" ContentType="' + CT_OXML
      + 'spreadsheetml.styles+xml"/>'
      + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.'
      + 'openxmlformats-package.core-properties+xml"/>'
      + '<Override PartName="/docProps/app.xml" ContentType="' + CT_OXML
      + 'extended-properties+xml"/></Types>');
  gWB = u(XMLDECL + '<workbook xmlns="' + NS_MAIN + '" xmlns:r="' + NS_REL + '">'
      + '<bookViews><workbookView/></bookViews><sheets></sheets></workbook>');
  gRel = u(XMLDECL + '<Relationships xmlns="' + NS_PKG + '">'
       + '<Relationship Id="rIdXSty" Type="' + NS_REL + '/styles" Target="styles.xml"/>'
       + '</Relationships>');
  gSty = styBase();
end-proc;


dcl-proc sheetNameUsed;                  // escaped name already in workbook.xml?
  dcl-pi *n ind;
    esc like(XMID_T) const;
  end-pi;
  dcl-s p int(10);
  dcl-s tag like(XMID_T);
  dcl-s want like(XMID_T);
  want = upU(esc);
  p = findOpen(gWB : u('sheets') : 1);
  if p = 0;
    return *off;
  endif;
  dow 1 = 1;
    p = findOpen(gWB : u('sheet') : p + 1);
    if p = 0;
      leave;
    endif;
    tag = tagText(gWB : p);
    if upU(attr(tag : u('name'))) = want;
      return *on;
    endif;
  enddo;
  return *off;
end-proc;

// header notes: comments + legacy VML part names not used yet in the package (from n)
dcl-proc notePartNames;
  dcl-pi *n;
    n int(10) value;
  end-pi;
  dcl-s c int(10);
  gNoteN = n;
  dow 1 = 1;
    if %scan(u('/xl/comments') + numU(gNoteN) + u('.xml') : gCT) = 0
       and %scan(u('vmlDrawing') + numU(gNoteN) + u('.vml') : gCT) = 0;
      if not gZrOpen;
        leave;
      endif;
      if zrFind(u('xl/comments') + numU(gNoteN) + u('.xml')) = 0 and
         zrFind(u('xl/drawings/vmlDrawing') + numU(gNoteN) + u('.vml')) = 0;
        leave;
      endif;
    endif;
    gNoteN += 1;
  enddo;
  c = %scan(u('</Types>') : gCT);
  if %scan(u('Extension="vml"') : gCT) = 0;
    gCT = %replace(u('<Default Extension="vml" ContentType="' + CT_OXML
        + 'vmlDrawing"/>') : gCT : c : 0);
    c = %scan(u('</Types>') : gCT);
  endif;
  gCT = %replace(u('<Override PartName="/xl/comments') + numU(gNoteN) + u('.xml"'
      + ' ContentType="' + CT_OXML + 'spreadsheetml.comments+xml"/>') : gCT : c : 0);
end-proc;

// insert a new sheet in workbook.xml, its relationship and content type
dcl-proc addSheetParts;
  dcl-pi *n ind;
    part like(XMID_T);
  end-pi;
  dcl-s n int(10);
  dcl-s k int(10);
  dcl-s q int(10);
  dcl-s mx int(20);
  dcl-s v int(20);
  dcl-s j int(10);
  dcl-s c int(10);
  dcl-s cand varchar(40);
  dcl-s suf varchar(8);
  dcl-s esc like(XMID_T);
  dcl-s tag like(XMID_T);
  dcl-s nm like(XMID_T);
  dcl-s nf like(TOK_T);

  n = 1;
  dow 1 = 1;
    nm = u('xl/worksheets/sheet') + numU(n) + u('.xml');
    if %scan(u('worksheets/sheet') + numU(n) + u('.xml') : gRel) = 0;
      if not gZrOpen;
        leave;
      endif;
      if zrFind(nm) = 0;
        leave;
      endif;
    endif;
    n += 1;
  enddo;
  part = nm;
  k = 1;
  dow %scan(u('Id="rIdXS') + numU(k) + cQuot : gRel) > 0;
    k += 1;
  enddo;
  mx = 0;
  q = 1;
  nf = u('sheetId="');
  dow q <= %len(gWB);
    q = %scan(nf : gWB : q);
    if q = 0;
      leave;
    endif;
    if q + 9 <= %len(gWB);
      v = digits(%subst(gWB : q + 9 : 1));
      if q + 18 <= %len(gWB);
        v = digits(%subst(gWB : q + 9 : 10));
      endif;
      if v > mx;
        mx = v;
      endif;
    endif;
    q += 9;
  enddo;

  // unique name
  j = 1;
  dow 1 = 1;
    if j = 1;
      cand = gSheet;
    else;
      suf = ' (' + %char(j) + ')';
      c = 31 - %len(suf);
      if %len(gSheet) < c;
        c = %len(gSheet);
      endif;
      cand = %subst(gSheet : 1 : c) + suf;
    endif;
    esc = xesc(u(cand));
    if not sheetNameUsed(esc);
      leave;
    endif;
    j += 1;
  enddo;
  gOutName = cand;

  tag = u('<sheet xmlns:r="' + NS_REL + '" name="') + esc + u('" sheetId="')
      + numU(mx + 1) + u('" r:id="rIdXS') + numU(k) + u('"/>');
  c = %scan(u('</sheets>') : gWB);
  if c > 0;
    gWB = %replace(tag : gWB : c : 0);
  else;
    c = %scan(u('<sheets/>') : gWB);
    if c = 0;
      fail('workbook.xml: <sheets> not found (unsupported layout)');
      return *off;
    endif;
    gWB = %replace(u('<sheets>') + tag + u('</sheets>') : gWB : c : 9);
  endif;
  c = %scan(u('</Relationships>') : gRel);
  if c = 0;
    fail('workbook.xml.rels: unsupported layout');
    return *off;
  endif;
  gRel = %replace(u('<Relationship Id="rIdXS') + numU(k) + u('" Type="' + NS_REL
       + '/worksheet" Target="worksheets/sheet') + numU(n) + u('.xml"/>') : gRel : c : 0);
  c = %scan(u('</Types>') : gCT);
  if c = 0;
    fail('[Content_Types].xml: unsupported layout');
    return *off;
  endif;
  gCT = %replace(u('<Override PartName="/') + part + u('" ContentType="' + CT_OXML
      + 'spreadsheetml.worksheet+xml"/>') : gCT : c : 0);

  if gNotes;
    notePartNames(n);
  endif;
  return *on;
end-proc;

// sheet by name (blank = last): part name and 0-based index
dcl-proc findSheet;
  dcl-pi *n ind;
    part like(XMID_T);
    sIdx int(10);
  end-pi;
  dcl-s want like(XMID_T);
  dcl-s tag like(XMID_T);
  dcl-s rid like(XMID_T);
  dcl-s trg like(XMID_T);
  dcl-s p int(10);
  dcl-s q int(10);
  dcl-s s0 int(10);
  dcl-s t int(10);
  dcl-s k int(10);
  %len(want) = 0;
  if gSheetGiven;
    want = upU(xesc(u(gSheet)));
  endif;
  %len(rid) = 0;
  p = findOpen(gWB : u('sheets') : 1);
  if p = 0;
    return *off;
  endif;
  k = -1;
  dow 1 = 1;
    p = findOpen(gWB : u('sheet') : p + 1);
    if p = 0;
      leave;
    endif;
    tag = tagText(gWB : p);
    k += 1;
    if %len(want) = 0 or upU(attr(tag : u('name'))) = want;
      rid = attr(tag : u('r:id'));
      sIdx = k;
      gOutName = eb(attr(tag : u('name')));
      if %len(want) > 0;
        leave;
      endif;
    endif;
  enddo;
  if %len(rid) = 0;
    return *off;
  endif;
  q = %scan(u('Id="') + rid + cQuot : gRel);
  if q = 0;
    return *off;
  endif;
  s0 = 0;
  t = 1;
  dow 1 = 1;
    t = findOpen(gRel : u('Relationship') : t);
    if t = 0 or t > q;
      leave;
    endif;
    s0 = t;
    t += 1;
  enddo;
  if s0 = 0;
    return *off;
  endif;
  trg = attr(tagText(gRel : s0) : u('Target'));
  if %len(trg) = 0;
    return *off;
  endif;
  if %subst(trg : 1 : 1) = cSl;
    part = %subst(trg : 2);
  else;
    part = u('xl/') + trg;
  endif;
  return *on;
end-proc;

// *REPLSHEET: workbook changes. Drop the _xlnm._FilterDatabase name of the sheet (Excel
// recreates it from the new autoFilter), drop calcChain.xml (its cell list would no
// longer match: Excel would "repair" the file) and ask for a full recalculation on open.
// Returns *on when calcChain.xml has to be left out of the package.
dcl-proc prepReplace;
  dcl-pi *n ind;
    sIdx int(10) value;
  end-pi;
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s c int(10);
  dcl-s k int(10);
  dcl-s m int(10);
  dcl-s q int(10);
  dcl-s tag like(XMID_T);
  dcl-s dropCC ind;
  dcl-s nxt like(TOK_T) dim(9);

  p = 1;
  dow 1 = 1;
    p = findOpen(gWB : u('definedName') : p);
    if p = 0;
      leave;
    endif;
    e = tagEnd(gWB : p);
    if e = 0;
      leave;
    endif;
    tag = tagText(gWB : p);
    if attr(tag : u('name')) = u('_xlnm._FilterDatabase')
       and attr(tag : u('localSheetId')) = numU(sIdx);
      if %subst(tag : %len(tag) - 1 : 1) = cSl;
        gWB = %replace(u('') : gWB : p : e - p);
      else;
        c = %scan(u('</definedName>') : gWB : e);
        if c > 0;
          gWB = %replace(u('') : gWB : p : c + 14 - p);
        endif;
      endif;
      leave;
    endif;
    p = e;
  enddo;

  dropCC = (dropEl(gRel : u('Relationship') : u('Type') : u('/calcChain')) > 0);
  dropEl(gCT : u('Override') : u('PartName') : u('/calcChain.xml'));

  p = findOpen(gWB : u('calcPr') : 1);
  if p > 0;
    e = tagEnd(gWB : p);
    m = %scan(u(' fullCalcOnLoad="') : gWB : p);
    if m > 0 and m < e;
      q = %scan(cQuot : gWB : m + 17);
      gWB = %replace(u('1') : gWB : m + 17 : q - m - 17);
    else;
      gWB = %replace(u(' fullCalcOnLoad="1"') : gWB : p + 7 : 0);
    endif;
  else;
    // CT_Workbook order: calcPr comes before these elements
    nxt(1) = u('oleSize');
    nxt(2) = u('customWorkbookViews');
    nxt(3) = u('pivotCaches');
    nxt(4) = u('smartTagPr');
    nxt(5) = u('smartTagTypes');
    nxt(6) = u('webPublishing');
    nxt(7) = u('fileRecoveryPr');
    nxt(8) = u('webPublishObjects');
    nxt(9) = u('extLst');
    p = 0;
    for k = 1 to 9;
      q = findOpen(gWB : nxt(k) : 1);
      if q > 0 and (p = 0 or q < p);
        p = q;
      endif;
    endfor;
    if p = 0;
      p = %scan(u('</workbook>') : gWB);
    endif;
    if p > 0;
      gWB = %replace(u('<calcPr fullCalcOnLoad="1"/>') : gWB : p : 0);
    endif;
  endif;
  return dropCC;
end-proc;

// =====================================================================================
// defined name, table, sheet state
// =====================================================================================
// part name of a relationship target: '../tables/table1.xml' from xl/worksheets/a.xml
dcl-proc resolveTgt;
  dcl-pi *n like(XMID_T);
    part like(XMID_T) const;
    tgt  like(XMID_T) const;
  end-pi;
  dcl-s dir like(XMID_T);
  dcl-s t like(XMID_T);
  dcl-s k int(10);
  dcl-s j int(10);
  t = tgt;
  if %len(t) > 0 and %subst(t : 1 : 1) = cSl;
    return %subst(t : 2);
  endif;
  k = 0;                                 // directory of part, with the final '/'
  j = %scan(cSl : part);
  dow j > 0;
    k = j;
    if j >= %len(part);
      leave;
    endif;
    j = %scan(cSl : part : j + 1);
  enddo;
  if k > 0;
    dir = %subst(part : 1 : k);
  else;
    %len(dir) = 0;
  endif;
  dow %len(t) > 3 and %subst(t : 1 : 3) = u('../');
    t = %subst(t : 4);
    if %len(dir) > 0;
      dir = %subst(dir : 1 : %len(dir) - 1);       // drop the final '/'
      k = 0;
      j = %scan(cSl : dir);
      dow j > 0;
        k = j;
        if j >= %len(dir);
          leave;
        endif;
        j = %scan(cSl : dir : j + 1);
      enddo;
      if k > 0;
        dir = %subst(dir : 1 : k);
      else;
        %len(dir) = 0;
      endif;
    endif;
  enddo;
  return dir + t;
end-proc;

// parts that the relationships of sheet part point to, by relationship type suffix
// ('/table', '/comments', '/vmlDrawing'); appended to arr (max mx), returns the count
dcl-proc relTargets;
  dcl-pi *n int(10);
    part like(XMID_T) const;
    suf  like(TOK_T) const;
    arr  varchar(512) ccsid(*utf8) dim(100) options(*varsize);
    n    int(10);
    mx   int(10) value;
  end-pi;
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s c int(10);
  dcl-s t like(XMID_T);
  dcl-s ty like(XMID_T);
  c = 0;
  if loadPart(relsOf(part) : gTmp) <> 1;
    return 0;
  endif;
  p = 1;
  dow 1 = 1;
    p = findOpen(gTmp : u('Relationship') : p);
    if p = 0;
      leave;
    endif;
    e = tagEnd(gTmp : p);
    if e = 0;
      leave;
    endif;
    t = %subst(gTmp : p : e - p);
    ty = attr(t : u('Type'));
    if %len(ty) >= %len(suf) and n < mx
       and attr(t : u('TargetMode')) <> u('External');
      if %subst(ty : %len(ty) - %len(suf) + 1) = suf;
        n += 1;
        arr(n) = resolveTgt(part : attr(t : u('Target')));
        c += 1;
      endif;
    endif;
    p = e;
  enddo;
  return c;
end-proc;

dcl-proc isDropped;
  dcl-pi *n ind;
    nm like(XMID_T) const;
  end-pi;
  dcl-s i int(10);
  for i = 1 to nDrop;
    if gDrop(i) = nm;
      return *on;
    endif;
  endfor;
  return *off;
end-proc;

dcl-proc nameUsed;                       // upper-case name among tables / defined names
  dcl-pi *n ind;
    up like(XMID_T) const;
  end-pi;
  dcl-s i int(10);
  for i = 1 to nNames;
    if gNames(i) = up;
      return *on;
    endif;
  endfor;
  return *off;
end-proc;

// names of the tables in the package (not dropped) and of the workbook defined names;
// gTblId = highest table id + 1
dcl-proc collectNames;
  dcl-s i int(10);
  dcl-s n int(10);
  dcl-s p int(10);
  dcl-s t like(XMID_T);
  dcl-s v like(XMID_T);
  dcl-ds e likeds(zrEnt_t);
  nNames = 0;
  nTblNm = 0;
  gTblId = 1;
  if not gZrOpen;
    return;
  endif;
  n = gZrN;
  for i = 1 to n;
    zrEntry(i : e);
    if %len(e.name) > 15 and %subst(e.name : 1 : 10) = u('xl/tables/')
       and %subst(e.name : %len(e.name) - 3) = u('.xml') and not isDropped(e.name);
      if loadPart(e.name : gTmp) = 1;
        p = findOpen(gTmp : u('table') : 1);
        if p > 0;
          t = tagText(gTmp : p);
          v = attr(t : u('id'));
          if %len(v) > 0 and digits(v) >= gTblId;
            gTblId = digits(v) + 1;
          endif;
          if nNames < %elem(gNames);
            nNames += 1;
            gNames(nNames) = upU(attr(t : u('displayName')));
          endif;
        endif;
      endif;
    endif;
  endfor;
  nTblNm = nNames;
  p = 1;
  dow nNames < %elem(gNames);
    p = findOpen(gWB : u('definedName') : p);
    if p = 0;
      leave;
    endif;
    t = tagText(gWB : p);
    nNames += 1;
    gNames(nNames) = upU(attr(t : u('name')));
    p += 1;
  enddo;
end-proc;

// table name: given, or derived from the sheet name; unique in the workbook
dcl-proc pickTableName;
  dcl-pi *n ind;
  end-pi;
  dcl-c OK 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_';
  dcl-s b varchar(255);
  dcl-s c varchar(255);
  dcl-s i int(10);
  dcl-s k int(10);
  if gTName <> '';
    if nameUsed(upU(u(gTName)));
      fail('TABLE ' + gTName + ': a table or defined name with this name exists');
      return *off;
    endif;
    return *on;
  endif;
  b = '';
  for i = 1 to %len(gOutName);
    if %check(OK : %subst(gOutName : i : 1)) = 0;
      b += %subst(gOutName : i : 1);
    else;
      b += '_';
    endif;
  endfor;
  if b = '' or nameErr(b) <> '';
    b = 'T_' + b;
  endif;
  if nameErr(b) <> '';
    b = 'Table';
  endif;
  c = b;
  k = 1;
  dow nameUsed(upU(u(c)));
    k += 1;
    c = b + '_' + %char(k);
  enddo;
  gTName = c;
  return *on;
end-proc;

// xl/tables/tableN.xml not used yet + its content type
dcl-proc tablePartName;
  dcl-s c int(10);
  gTblN = 1;
  dow %scan(u('/xl/tables/table') + numU(gTblN) + u('.xml"') : gCT) > 0
      or (gZrOpen and zrFind(u('xl/tables/table') + numU(gTblN) + u('.xml')) > 0);
    gTblN += 1;
  enddo;
  c = %scan(u('</Types>') : gCT);
  gCT = %replace(u('<Override PartName="/xl/tables/table') + numU(gTblN) + u('.xml"'
      + ' ContentType="' + CT_OXML + 'spreadsheetml.table+xml"/>') : gCT : c : 0);
end-proc;

// table name + part for the new sheet; RANGENAME must not collide with a table
dcl-proc tableSetup;
  dcl-pi *n ind;
  end-pi;
  dcl-s i int(10);
  dcl-s up like(XMID_T);
  if gTbl;
    if not pickTableName();
      return *off;
    endif;
    tablePartName();
  endif;
  if gRName <> '';
    up = upU(u(gRName));
    if gTbl and up = upU(u(gTName));
      fail('RANGENAME and TABLE need different names (' + gRName + ')');
      return *off;
    endif;
    for i = 1 to nTblNm;
      if gNames(i) = up;
        fail('RANGENAME ' + gRName + ': a table with this name exists');
        return *off;
      endif;
    endfor;
  endif;
  return *on;
end-proc;

dcl-proc isApTable;
  dcl-pi *n ind;
    nm like(XMID_T) const;
  end-pi;
  dcl-s i int(10);
  for i = 1 to nApTbl;
    if gApTbl(i) = nm;
      return *on;
    endif;
  endfor;
  return *off;
end-proc;

// position in workbook.xml for an element that comes after <sheets> (CT_Workbook order):
// before the first following element that exists, else before </workbook>
dcl-proc wbPos;
  dcl-pi *n int(10);
    beforeCalc ind value;                // *on: definedNames (before calcPr too)
  end-pi;
  dcl-s nxt like(TOK_T) dim(10);
  dcl-s k int(10);
  dcl-s p int(10);
  dcl-s q int(10);
  nxt(1) = u('oleSize');
  nxt(2) = u('customWorkbookViews');
  nxt(3) = u('pivotCaches');
  nxt(4) = u('smartTagPr');
  nxt(5) = u('smartTagTypes');
  nxt(6) = u('webPublishing');
  nxt(7) = u('fileRecoveryPr');
  nxt(8) = u('webPublishObjects');
  nxt(9) = u('extLst');
  nxt(10) = u('calcPr');
  p = 0;
  for k = 1 to 10;
    if k = 10 and not beforeCalc;
      leave;
    endif;
    q = findOpen(gWB : nxt(k) : 1);
    if q > 0 and (p = 0 or q < p);
      p = q;
    endif;
  endfor;
  if p = 0;
    p = %scan(u('</workbook>') : gWB);
  endif;
  return p;
end-proc;

// 'Sheet name'!$A$r0:$X$r1 (escaped for XML)
dcl-proc rangeRef;
  dcl-pi *n like(XMID_T);
    r0 int(20) value;
    r1 int(20) value;
  end-pi;
  dcl-s q like(XMID_T);
  q = %scanrpl(u('''') : u('''''') : u(gOutName));
  return xesc(u('''') + q + u('''!$A$') + numU(r0) + u(':$')
              + col(nCol).ref + cDol + numU(r1));
end-proc;

// RANGENAME in workbook.xml: replaced (new / replaced sheet) or last row moved (*APPEND)
dcl-proc setDefName;
  dcl-pi *n;
    lastRow int(20) value;
  end-pi;
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s c int(10);
  dcl-s k int(10);
  dcl-s j int(10);
  dcl-s want like(XMID_T);
  dcl-s tag like(XMID_T);
  dcl-s body like(XMID_T);
  dcl-s el like(XMID_T);
  if gRName = '';
    return;
  endif;
  if lastRow < gRow0;
    lastRow = gRow0;
  endif;
  want = upU(u(gRName));
  p = 1;
  dow 1 = 1;                             // existing workbook-level name?
    p = findOpen(gWB : u('definedName') : p);
    if p = 0;
      leave;
    endif;
    e = tagEnd(gWB : p);
    if e = 0;
      p = 0;
      leave;
    endif;
    tag = tagText(gWB : p);
    if upU(attr(tag : u('name'))) = want and %len(attr(tag : u('localSheetId'))) = 0
       and %subst(tag : %len(tag) - 1 : 1) <> cSl;
      leave;
    endif;
    p = e;
  enddo;
  if p > 0;
    c = %scan(u('</definedName>') : gWB : e);
    if c = 0;
      return;
    endif;
    if gMode = 'A';                      // keep the start, move the last row
      body = %subst(gWB : e : c - e);
      k = 0;
      j = %scan(cDol : body);
      dow j > 0;
        k = j;
        if j >= %len(body);
          leave;
        endif;
        j = %scan(cDol : body : j + 1);
      enddo;
      if k > 0 and k < %len(body) and digits(%subst(body : k + 1)) > 0;
        body = %subst(body : 1 : k) + numU(lastRow);
      else;
        body = rangeRef(gRow0 : lastRow);
      endif;
    else;
      body = rangeRef(gRow0 : lastRow);
    endif;
    gWB = %replace(body : gWB : e : c - e);
    gNewWB = *on;
    return;
  endif;
  el = u('<definedName name="') + xesc(u(gRName)) + u('">') + rangeRef(gRow0 : lastRow)
     + u('</definedName>');
  p = %scan(u('</definedNames>') : gWB);
  if p > 0;
    gWB = %replace(el : gWB : p : 0);
  else;
    p = %scan(u('<definedNames/>') : gWB);
    if p > 0;
      gWB = %replace(u('<definedNames>') + el + u('</definedNames>') : gWB : p : 15);
    else;
      p = wbPos(*on);
      if p = 0;
        fail('workbook.xml: unsupported layout (defined name)');
        return;
      endif;
      gWB = %replace(u('<definedNames>') + el + u('</definedNames>') : gWB : p : 0);
    endif;
  endif;
  gNewWB = *on;
end-proc;

// state of the sheet with 0-based index sIdx (gHide N/H/V), then: at least one visible
// sheet, and the active / first tab of the window on a visible sheet
dcl-proc setSheetState;
  dcl-pi *n ind;
    sIdx int(10) value;
  end-pi;
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s k int(10);
  dcl-s m int(10);
  dcl-s q int(10);
  dcl-s firstVis int(10);
  dcl-s act int(10);
  dcl-s tag like(XMID_T);
  dcl-s st like(XMID_T);
  dcl-s nv like(TOK_T);
  dcl-s vis ind dim(1000);
  dcl-s nSh int(10);

  if sIdx = -2;                          // the last sheet (just added)
    sIdx = -1;
    p = findOpen(gWB : u('sheets') : 1);
    dow p > 0;
      p = findOpen(gWB : u('sheet') : p + 1);
      if p > 0;
        sIdx += 1;
      endif;
    enddo;
  endif;
  if gHide <> ' ' and sIdx >= 0;
    select;
    when gHide = 'H';
      nv = u(' state="hidden"');
    when gHide = 'V';
      nv = u(' state="veryHidden"');
    other;
      nv = u(' state="visible"');
    endsl;
    p = findOpen(gWB : u('sheets') : 1);
    k = -1;
    dow p > 0;
      p = findOpen(gWB : u('sheet') : p + 1);
      if p = 0;
        leave;
      endif;
      k += 1;
      if k = sIdx;
        e = tagEnd(gWB : p);
        m = %scan(u(' state="') : gWB : p);
        if m > 0 and m < e;
          q = %scan(cQuot : gWB : m + 8);
          gWB = %replace(nv : gWB : m : q - m + 1);
        else;
          gWB = %replace(nv : gWB : p + 6 : 0);
        endif;
        gNewWB = *on;
        leave;
      endif;
    enddo;
  endif;

  // visible sheets
  nSh = 0;
  firstVis = -1;
  p = findOpen(gWB : u('sheets') : 1);
  dow p > 0 and nSh < %elem(vis);
    p = findOpen(gWB : u('sheet') : p + 1);
    if p = 0;
      leave;
    endif;
    tag = tagText(gWB : p);
    st = attr(tag : u('state'));
    nSh += 1;
    vis(nSh) = (%len(st) = 0 or st = u('visible'));
    if vis(nSh) and firstVis < 0;
      firstVis = nSh - 1;
    endif;
  enddo;
  if firstVis < 0;
    fail('HIDE: at least one sheet of the workbook must stay visible');
    return *off;
  endif;
  // <workbookView activeTab="n" firstSheet="n">: move off a hidden sheet
  p = findOpen(gWB : u('workbookView') : 1);
  if p > 0;
    tag = tagText(gWB : p);
    st = attr(tag : u('activeTab'));
    act = 0;
    if %len(st) > 0;
      act = digits(st);
    endif;
    if act >= nSh or not vis(act + 1);
      e = tagEnd(gWB : p);
      m = %scan(u(' activeTab="') : gWB : p);
      if m > 0 and m < e;
        q = %scan(cQuot : gWB : m + 12);
        gWB = %replace(numU(firstVis) : gWB : m + 12 : q - m - 12);
      else;
        gWB = %replace(u(' activeTab="') + numU(firstVis) + cQuot : gWB : p + 13 : 0);
      endif;
      gNewWB = *on;
    endif;
    tag = tagText(gWB : p);
    st = attr(tag : u('firstSheet'));
    if %len(st) > 0;
      act = digits(st);
      if act >= nSh or not vis(act + 1);
        e = tagEnd(gWB : p);
        m = %scan(u(' firstSheet="') : gWB : p);
        q = %scan(cQuot : gWB : m + 13);
        gWB = %replace(numU(firstVis) : gWB : m + 13 : q - m - 13);
        gNewWB = *on;
      endif;
    endif;
  endif;
  return *on;
end-proc;

// header texts unique (case-insensitive), as Excel requires for table columns
dcl-proc uniqueHeaders;
  dcl-s i int(10);
  dcl-s j int(10);
  dcl-s k int(10);
  dcl-s base like(XMID_T);
  dcl-s cur like(XMID_T);
  dcl-s dup ind;
  for i = 2 to nCol;
    base = col(i).hdr;
    cur = base;
    k = 1;
    dow 1 = 1;
      dup = *off;
      for j = 1 to i - 1;
        if upU(col(j).hdr) = upU(cur);
          dup = *on;
          leave;
        endif;
      endfor;
      if not dup;
        leave;
      endif;
      k += 1;
      cur = base + numU(k);
    enddo;
    col(i).hdr = cur;
  endfor;
end-proc;

// relationships of a sheet created by the run: notes and / or table
dcl-proc writeSheetRels;
  dcl-pi *n ind;
    part like(XMID_T) const;
  end-pi;
  if not gNotes and not gTbl;
    return *on;
  endif;
  gTmp = u(XMLDECL + '<Relationships xmlns="' + NS_PKG + '">');
  if gNotes;
    gTmp += u('<Relationship Id="rIdXC" Type="' + NS_REL + '/comments"'
          + ' Target="../comments') + numU(gNoteN) + u('.xml"/><Relationship Id="rIdXV"'
          + ' Type="' + NS_REL + '/vmlDrawing" Target="../drawings/vmlDrawing')
          + numU(gNoteN) + u('.vml"/>');
  endif;
  if gTbl;
    gTmp += u('<Relationship Id="rIdXT" Type="' + NS_REL + '/table"'
          + ' Target="../tables/table') + numU(gTblN) + u('.xml"/>');
  endif;
  gTmp += u('</Relationships>');
  return writePart(relsOf(part) : gTmp);
end-proc;

// xl/tables/tableN.xml over header + data rows (at least one data row)
dcl-proc writeTable;
  dcl-pi *n ind;
    lastRow int(20) value;
  end-pi;
  dcl-s i int(10);
  dcl-s ref like(XMID_T);
  dcl-s nm like(XMID_T);
  if lastRow <= gRow0;
    lastRow = gRow0 + 1;
  endif;
  ref = u('A') + numU(gRow0) + cColon + col(nCol).ref + numU(lastRow);
  nm = xesc(u(gTName));
  gTmp = u(XMLDECL + '<table xmlns="' + NS_MAIN + '" id="') + numU(gTblId)
       + u('" name="') + nm + u('" displayName="') + nm + u('" ref="') + ref
       + u('" totalsRowShown="0">');
  if gFilter;
    gTmp += u('<autoFilter ref="') + ref + u('"/>');
  endif;
  gTmp += u('<tableColumns count="') + numU(nCol) + u('">');
  for i = 1 to nCol;
    gTmp += u('<tableColumn id="') + numU(i) + u('" name="') + col(i).hdr + u('"/>');
  endfor;
  gTmp += u('</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0"'
        + ' showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>');
  return writePart(u('xl/tables/table') + numU(gTblN) + u('.xml') : gTmp);
end-proc;

// *APPEND: a table of the sheet that ended at the old last row now ends at newLast
dcl-proc patchTablePart;
  dcl-pi *n ind;
    nm like(XMID_T) const;
    newLast int(20) value;
  end-pi;
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s k int(10);
  dcl-s m int(10);
  dcl-s tag like(XMID_T);
  dcl-s ref like(XMID_T);
  dcl-s nr like(XMID_T);
  dcl-s el like(TOK_T);
  if loadPart(nm : gTmp) <> 1;
    return *off;
  endif;
  p = findOpen(gTmp : u('table') : 1);
  if p > 0;
    tag = tagText(gTmp : p);
    ref = attr(tag : u('ref'));
    k = %scan(cColon : ref);
    m = digits(attr(tag : u('totalsRowCount')));
    if k > 0 and k < %len(ref) and m = 0
       and digits(%subst(ref : k + 1 + %len(noTrailDigits(%subst(ref : k + 1)))))
           = gApLast;
      nr = %subst(ref : 1 : k) + noTrailDigits(%subst(ref : k + 1)) + numU(newLast);
      gTmp = %scanrpl(u(' ref="') + ref + cQuot : u(' ref="') + nr + cQuot : gTmp);
    endif;
  endif;
  return writePart(nm : gTmp);
end-proc;

dcl-proc writePart;
  dcl-pi *n ind;
    name like(XMID_T) const;
    s    like(XBIG_T);
  end-pi;
  if zwBegin(name) < 0;
    fail(xlErr());
    return *off;
  endif;
  wp = *null;
  we = *null;
  if not putBig(s);
    return *off;
  endif;
  if zwEnd() < 0;
    fail(xlErr());
    return *off;
  endif;
  return *on;
end-proc;

dcl-proc loadPart;                       // 1 loaded, 0 missing, -1 error
  dcl-pi *n int(10);
    name like(XMID_T) const;
    s    like(XBIG_T);
  end-pi;
  dcl-s i int(10);
  dcl-s p pointer;
  dcl-s n int(10);
  i = zrFind(name);
  if i = 0;
    return 0;
  endif;
  p = zrGet(i : n);
  if p = *null;
    fail(xlErr());
    return -1;
  endif;
  if n > XBIGMAX;
    dealloc(n) p;
    fail(eb(name) + ' too large to update');
    return -1;
  endif;
  %len(s) = n;
  if n > 0;
    memcpy(%addr(s : *data) : p : n);
  endif;
  dealloc(n) p;
  return 1;
end-proc;

dcl-proc miscParts;                      // _rels/.rels, docProps
  dcl-pi *n ind;
  end-pi;
  dcl-s usr char(10) inz(*user);
  dcl-s tsU timestamp;
  dcl-s ts varchar(30);
  exec sql set :tsU = current timestamp - current timezone;
  ts = %char(%date(tsU) : *iso) + 'T' + %char(%time(tsU) : *hms) + 'Z';
  gTmp = u(XMLDECL + '<Relationships xmlns="' + NS_PKG + '">'
       + '<Relationship Id="rId1" Type="' + NS_REL + '/officeDocument"'
       + ' Target="xl/workbook.xml"/>'
       + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/'
       + 'relationships/metadata/core-properties" Target="docProps/core.xml"/>'
       + '<Relationship Id="rId3" Type="' + NS_REL + '/extended-properties"'
       + ' Target="docProps/app.xml"/></Relationships>');
  if not writePart(u('_rels/.rels') : gTmp);
    return *off;
  endif;
  gTmp = u(XMLDECL + '<Properties xmlns="http://schemas.openxmlformats.org/'
       + 'officeDocument/2006/extended-properties"><Application>XLSX00R ' + XLSX_VERSION
       + ' IBM i'
       + '</Application></Properties>');
  if not writePart(u('docProps/app.xml') : gTmp);
    return *off;
  endif;
  gTmp = u(XMLDECL + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/'
       + 'package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/'
       + '1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/'
       + '2001/XMLSchema-instance"><dc:creator>') + xesc(u(%trim(usr)))
       + u('</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + ts
       + '</dcterms:created></cp:coreProperties>');
  return writePart(u('docProps/core.xml') : gTmp);
end-proc;

// *REPLACE: a new workbook
dcl-proc buildNew;
  dcl-pi *n ind;
  end-pi;
  dcl-s part like(XMID_T);
  baseParts();
  if not mergeStyles();
    return *off;
  endif;
  setColStyles();
  if not addSheetParts(part);
    return *off;
  endif;
  gZrN = 0;
  collectNames();
  if not tableSetup();
    return *off;
  endif;
  if zwOpen(gTmpF : gLevel) < 0;
    fail(xlErr());
    return *off;
  endif;
  if not writePart(u('[Content_Types].xml') : gCT);
    return *off;
  endif;
  if not miscParts();
    return *off;
  endif;
  if not writePart(u('xl/_rels/workbook.xml.rels') : gRel)
     or not writePart(u('xl/styles.xml') : gSty);
    return *off;
  endif;
  if not writeNewSheet(part);            // before workbook.xml: the name needs the rows
    return *off;
  endif;
  setDefName(rowNum);
  if not writePart(u('xl/workbook.xml') : gWB);
    return *off;
  endif;
  if zwClose() < 0;
    fail(xlErr());
    return *off;
  endif;
  return *on;
end-proc;

// *ADDSHEET / *APPEND on an existing workbook
dcl-proc buildUpdate;
  dcl-pi *n ind;
  end-pi;
  dcl-s n int(10);
  dcl-s i int(10);
  dcl-s r int(10);
  dcl-s sIdx int(10);
  dcl-s part like(XMID_T);
  dcl-s nmCT like(XMID_T);
  dcl-s nmWB like(XMID_T);
  dcl-s nmRel like(XMID_T);
  dcl-s nmSty like(XMID_T);
  dcl-ds e likeds(zrEnt_t);
  dcl-s pX pointer;
  dcl-s nX int(10);
  dcl-s ok ind;
  dcl-s dropCC ind;
  dcl-s oldRels like(XMID_T);
  dcl-s nmCC like(XMID_T);

  nmCC = u('xl/calcChain.xml');
  nmCT = u('[Content_Types].xml');
  nmWB = u('xl/workbook.xml');
  nmRel = u('xl/_rels/workbook.xml.rels');
  nmSty = u('xl/styles.xml');
  n = zrOpen(gStmf);
  if n < 0;
    fail(xlErr());
    return *off;
  endif;
  gZrOpen = *on;
  gZrN = n;
  nDrop = 0;
  nApTbl = 0;
  if loadPart(nmCT : gCT) <> 1 or loadPart(nmWB : gWB) <> 1
     or loadPart(nmRel : gRel) <> 1;
    fail(gStmf + ': not a spreadsheet package (workbook parts missing)');
    return *off;
  endif;
  r = loadPart(nmSty : gSty);
  if r < 0;
    return *off;
  endif;
  gStyAdd = *off;
  if r = 0;                              // no styles.xml: add ours
    gStyAdd = *on;
    gSty = styBase();
    i = %scan(u('</Types>') : gCT);
    gCT = %replace(u('<Override PartName="/xl/styles.xml" ContentType="' + CT_OXML
        + 'spreadsheetml.styles+xml"/>') : gCT : i : 0);
    i = %scan(u('</Relationships>') : gRel);
    gRel = %replace(u('<Relationship Id="rIdXSty" Type="' + NS_REL
         + '/styles" Target="styles.xml"/>') : gRel : i : 0);
  endif;
  if not mergeStyles();
    return *off;
  endif;
  setColStyles();
  gNewWB = *off;
  dropCC = *off;
  %len(oldRels) = 0;
  if gMode = 'A';
    if not findSheet(part : sIdx);
      gMode = 'S';
    else;
      gTbl = *off;                       // existing sheet: its tables are extended
      relTargets(part : u('/table') : gApTbl : nApTbl : %elem(gApTbl));
    endif;
  endif;
  if gMode = 'P';
    gSheetGiven = *on;                   // replace the sheet with the (default) name
    if not findSheet(part : sIdx);
      gMode = 'S';
    elseif %scan(u('worksheets/') : part) = 0;
      fail('sheet "' + gOutName + '" is not a worksheet (chart or dialog sheet)');
      return *off;
    else;
      oldRels = relsOf(part);            // notes, drawings, tables of the old content
      // the old sheet's tables, notes and VML go: their names / numbers become free
      relTargets(part : u('/table') : gDrop : nDrop : %elem(gDrop));
      relTargets(part : u('/comments') : gDrop : nDrop : %elem(gDrop));
      relTargets(part : u('/vmlDrawing') : gDrop : nDrop : %elem(gDrop));
      for i = 1 to nDrop;
        dropEl(gCT : u('Override') : u('PartName') : cSl + gDrop(i));
      endfor;
      if gNotes;
        notePartNames(1);
      endif;
      dropCC = prepReplace(sIdx);
      gNewWB = *on;
    endif;
  endif;
  if gMode = 'S';
    if not addSheetParts(part);
      return *off;
    endif;
    gNewWB = *on;
    sIdx = -2;
  endif;
  collectNames();
  if not tableSetup();
    return *off;
  endif;
  if gHide <> ' ';
    if not setSheetState(sIdx);
      return *off;
    endif;
  endif;

  if zwOpen(gTmpF : gLevel) < 0;
    fail(xlErr());
    return *off;
  endif;
  // [Content_Types].xml first, then the sheet, then everything else
  if gMode = 'S' or gMode = 'P' or gStyAdd;
    ok = writePart(nmCT : gCT);
  else;
    zrEntry(zrFind(nmCT) : e);
    ok = (zwRaw(e.name : e.method : e.flags : e.dtime : e.ddate : e.crc : e.csize
                : e.usize : e.data) = 0);
    if not ok;
      fail(xlErr());
    endif;
  endif;
  if not ok;
    return *off;
  endif;
  if gMode = 'S' or gMode = 'P';
    if not writeNewSheet(part);
      return *off;
    endif;
    setDefName(rowNum);
  else;
    i = zrFind(part);
    if i = 0;
      fail('sheet part ' + eb(part) + ' not found in the package');
      return *off;
    endif;
    pX = zrGet(i : nX);
    if pX = *null;
      fail(xlErr());
      return *off;
    endif;
    ok = writeAppendSheet(part : pX : nX : sIdx);
    dealloc(n) pX;
    if not ok;
      return *off;
    endif;
    setDefName(rowNum);
  endif;
  for i = 1 to n;
    zrEntry(i : e);
    if e.name = nmCT or e.name = part;
      iter;
    endif;
    if gMode = 'P' and (e.name = oldRels or (dropCC and e.name = nmCC)
                        or isDropped(e.name));
      iter;                              // replaced sheet: old relationships, calcChain,
    endif;                               // its tables, notes, VML
    if gMode = 'A' and isApTable(e.name);
      if not patchTablePart(e.name : rowNum);
        return *off;
      endif;
      iter;
    endif;
    if e.name = nmWB and gNewWB;
      ok = writePart(nmWB : gWB);
    elseif e.name = nmRel and (gMode = 'S' or gMode = 'P' or gStyAdd);
      ok = writePart(nmRel : gRel);
    elseif e.name = nmSty;
      ok = writePart(nmSty : gSty);
    else;
      ok = (zwRaw(e.name : e.method : e.flags : e.dtime : e.ddate : e.crc : e.csize
                  : e.usize : e.data) = 0);
      if not ok;
        fail(xlErr());
      endif;
    endif;
    if not ok;
      return *off;
    endif;
  endfor;
  if gStyAdd;                            // the package had no styles.xml
    if not writePart(nmSty : gSty);
      return *off;
    endif;
  endif;
  if zwClose() < 0;
    fail(xlErr());
    return *off;
  endif;
  zrClose();
  gZrOpen = *off;
  return *on;
end-proc;

// =====================================================================================
// messages
// =====================================================================================
dcl-proc sendMsg;
  dcl-pi *n;
    typ char(10) const;
    txt varchar(512) const;
  end-pi;
  dcl-s key char(4);
  dcl-ds err;
    eProv int(10) inz(0);
    eAvl  int(10) inz(0);
  end-ds;
  dcl-s id char(7);
  if typ = '*ESCAPE';
    id = 'CPF9898';
  else;
    id = 'CPF9897';
  endif;
  sndPgmMsg(id : 'QCPFMSG   *LIBL' : txt : %len(txt) : typ : '*PGMBDY' : 1 : key : err);
end-proc;

// =====================================================================================
// main
// =====================================================================================
dcl-proc xlsxMain;
  dcl-pi *n extpgm('XLSX00R');
    pSrc     char(5000) const;
    pStmf    char(1024) const;
    pMode    char(10) const options(*nopass : *omit);
    pSheet   char(31) const options(*nopass : *omit);
    pHeader  char(10) const options(*nopass : *omit);
    pHdrList char(5000) const options(*nopass : *omit);
    pFilter  char(4) const options(*nopass : *omit);
    pFreeze  char(4) const options(*nopass : *omit);
    pLevel   packed(1 : 0) const options(*nopass : *omit);
    pColDesc char(10) const options(*nopass : *omit);
    pRows    packed(15 : 0) options(*nopass : *omit);
    pMsg     char(256) options(*nopass : *omit);
    pRName   char(255) const options(*nopass : *omit);
    pTable   char(255) const options(*nopass : *omit);
    pHide    char(10) const options(*nopass : *omit);
  end-pi;
  dcl-c LO 'abcdefghijklmnopqrstuvwxyz';
  dcl-c UP 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  dcl-s t0 timestamp;
  dcl-s msgOut ind;
  dcl-s ok ind;
  dcl-s x varchar(10);
  dcl-s p int(10);
  dcl-s secs packed(15 : 3);
  dcl-s txt varchar(512);
  dcl-ds st likeds(xlStat_t);

  t0 = %timestamp();
  gFetchUs = 0;
  gStats = (getenv('XLSX00_STATS') <> *null);
  if gStats;
    gStats = (%str(getenv('XLSX00_STATS')) = 'Y');
  endif;
  gFail = *off;
  gMsg = '';
  gRows = 0;
  gTrunc = *off;
  gCurOpen = *off;
  gZrOpen = *off;
  msgOut = *off;
  if %parms() >= 12;
    if %addr(pMsg) <> *null;
      if %subst(pMsg : 1 : 7) <> '*ESCAPE';
        msgOut = *on;
        pMsg = *blanks;
      endif;
    endif;
  endif;

  monitor;
    // ---- parameters
    gSrc = %trim(pSrc);
    p = %scan(x'00' : gSrc);
    if p > 0;
      gSrc = %trim(%subst(gSrc + ' ' : 1 : p - 1));
    endif;
    gStmf = %trim(pStmf);
    p = %scan(x'00' : gStmf);
    if p > 0;
      gStmf = %trim(%subst(gStmf + ' ' : 1 : p - 1));
    endif;
    gMode = 'R';
    if %parms() >= 3;
      if %addr(pMode) <> *null;
        x = %xlate(LO : UP : %trim(pMode));
        if x = '*ADDSHEET' or x = 'ADDSHEET';
          gMode = 'S';
        elseif x = '*APPEND' or x = 'APPEND';
          gMode = 'A';
        elseif x = '*REPLSHEET' or x = 'REPLSHEET';
          gMode = 'P';
        elseif x <> '' and x <> '*REPLACE' and x <> 'REPLACE';
          fail('MODE ' + x + ' not valid (*REPLACE, *ADDSHEET, *APPEND, *REPLSHEET)');
        endif;
      endif;
    endif;
    gSheet = '';
    if %parms() >= 4;
      if %addr(pSheet) <> *null;
        gSheet = %trim(%xlate('[]:*?/\' : '_______' : pSheet));
      endif;
    endif;
    dow %len(gSheet) > 0;
      if %subst(gSheet : 1 : 1) <> '''';
        leave;
      endif;
      gSheet = %trim(%subst(gSheet + ' ' : 2));
    enddo;
    gHdrMode = 'L';
    if %parms() >= 5;
      if %addr(pHeader) <> *null;
        x = %xlate(LO : UP : %trim(pHeader));
        if x = '*NAMES' or x = 'NAMES';
          gHdrMode = 'N';
        elseif x = '*TEXT' or x = 'TEXT';
          gHdrMode = 'T';
        elseif x = '*NONE' or x = 'NONE';
          gHdrMode = 'X';
        endif;
      endif;
    endif;
    gHdr = (gHdrMode <> 'X');
    gHdrList = '';
    if %parms() >= 6;
      if %addr(pHdrList) <> *null;
        gHdrList = %trimr(pHdrList);
        p = %scan(x'00' : gHdrList);
        if p > 0;
          gHdrList = %trimr(%subst(gHdrList + ' ' : 1 : p - 1));
        endif;
      endif;
    endif;
    gFilter = *on;
    if %parms() >= 7;
      if %addr(pFilter) <> *null;
        x = %xlate(LO : UP : %trim(pFilter));
        gFilter = not (x = '*NO' or x = 'NO' or x = 'N');
      endif;
    endif;
    gFreeze = *on;
    if %parms() >= 8;
      if %addr(pFreeze) <> *null;
        x = %xlate(LO : UP : %trim(pFreeze));
        gFreeze = not (x = '*NO' or x = 'NO' or x = 'N');
      endif;
    endif;
    gLevel = 2;
    if %parms() >= 9;
      if %addr(pLevel) <> *null;
        gLevel = pLevel;
      endif;
    endif;
    gDesc = 'N';
    if %parms() >= 10;
      if %addr(pColDesc) <> *null;
        x = %xlate(LO : UP : %trim(pColDesc));
        if x = '*ROW' or x = 'ROW';
          gDesc = 'R';
        elseif x = '*NOTE' or x = 'NOTE';
          gDesc = 'C';
        elseif x = '*BOTH' or x = 'BOTH';
          gDesc = 'B';
        elseif x <> '' and x <> '*NONE' and x <> 'NONE';
          fail('COLDESC ' + x + ' not valid (*NONE, *ROW, *NOTE, *BOTH)');
        endif;
      endif;
    endif;
    gNotes = gHdr and (gDesc = 'C' or gDesc = 'B');
    gRName = '';
    if %parms() >= 13;
      if %addr(pRName) <> *null;
        gRName = %trim(pRName);
        x = %xlate(LO : UP : %subst(gRName + '     ' : 1 : 5));
        if gRName = '*NONE' or x = '*NONE';
          gRName = '';
        endif;
      endif;
    endif;
    if gRName <> '';
      txt = nameErr(gRName);
      if txt <> '';
        fail('RANGENAME ' + gRName + ': ' + txt);
      endif;
    endif;
    gTbl = *off;
    gTName = '';
    if %parms() >= 14;
      if %addr(pTable) <> *null;
        gTName = %trim(pTable);
        x = %xlate(LO : UP : %subst(gTName + '     ' : 1 : 5));
        if gTName = '' or x = '*NONE' or x = 'NONE';
          gTName = '';
        elseif x = '*YES' or x = 'YES';
          gTbl = *on;
          gTName = '';
        else;
          gTbl = *on;
          txt = nameErr(gTName);
          if txt <> '';
            fail('TABLE ' + gTName + ': ' + txt);
          endif;
        endif;
      endif;
    endif;
    if gTbl and not gHdr;
      fail('TABLE needs a header row (HEADER(*NONE) given)');
    endif;
    if gTbl and (gDesc = 'R' or gDesc = 'B');
      fail('TABLE needs one header row: use COLDESC(*NOTE) instead of *ROW / *BOTH');
    endif;
    gHide = ' ';
    if %parms() >= 15;
      if %addr(pHide) <> *null;
        x = %xlate(LO : UP : %trim(pHide));
        if x = '*NO' or x = 'NO';
          gHide = 'N';
        elseif x = '*YES' or x = 'YES';
          gHide = 'H';
        elseif x = '*VERY' or x = 'VERY';
          gHide = 'V';
        elseif x <> '' and x <> '*DFT' and x <> 'DFT';
          fail('HIDE ' + x + ' not valid (*DFT, *NO, *YES, *VERY)');
        endif;
      endif;
    endif;
    if gSrc = '';
      fail('SRC is blank: an SQL statement or a file name is required');
    endif;
    if gStmf = '';
      fail('TOSTMF is blank');
    endif;
    gTmpF = gStmf + '.tmp';

    // ---- run
    if not gFail;
      initTokens();
      gUtf8Ok = probeUtf8();
      ok = openSource();
      if ok;
        gSheetGiven = (gSheet <> '');
        if not gSheetGiven;
          gSheet = %trim(%xlate('[]:*?/\' : '_______' : gDftSheet));
          if gSheet = '';
            gSheet = 'Sheet1';
          endif;
          if %len(gSheet) > 31;
            gSheet = %subst(gSheet : 1 : 31);
          endif;
        endif;
        gOutName = gSheet;
        if gMode <> 'R' and not xlExists(gStmf);
          gMode = 'R';
        endif;
        if gMode = 'R' and (gHide = 'H' or gHide = 'V');
          fail('HIDE(*YES / *VERY) needs an existing workbook with a visible sheet');
          ok = *off;
        elseif gMode = 'R';
          ok = buildNew();
        else;
          ok = buildUpdate();
        endif;
      endif;
      closeCursor();
      if ok and not gFail;
        if xlRename(gTmpF : gStmf) < 0;
          fail(xlErr());
        endif;
      endif;
    endif;
  on-error;
    fail('XLSX00R ended abnormally, status ' + %char(%status()));
  endmon;

  if gFail;
    monitor;
      closeCursor();
    on-error;
    endmon;
    zwAbort();
    zrClose();
    if %parms() >= 11;
      if %addr(pRows) <> *null;
        pRows = 0;
      endif;
    endif;
    if msgOut;
      pMsg = gMsg;
    else;
      sendMsg('*ESCAPE' : 'XLSX00R ' + XLSX_VERSION + ': ' + gMsg);
    endif;
    return;
  endif;

  if %parms() >= 11;
    if %addr(pRows) <> *null;
      pRows = gRows;
    endif;
  endif;
  if gStats;
    xlStats(st);
    txt = 'STATS total ' + %char(%diff(%timestamp() : t0 : *mseconds) / 1000)
        + ' ms, fetch ' + %char(gFetchUs / 1000) + ', crc ' + %char(st.crcUs / 1000)
        + ', deflate ' + %char(st.defUs / 1000) + ', write ' + %char(st.wrtUs / 1000)
        + ' ms; xml ' + %char(st.inB) + ' B -> ' + %char(st.outB) + ' B';
    if msgOut;
      pMsg = txt;
    else;
      sendMsg('*INFO' : txt);
    endif;
  endif;
  if not msgOut;
    secs = %diff(%timestamp() : t0 : *mseconds) / 1000000;
    txt = 'XLSX00R ' + XLSX_VERSION + ': ' + %char(gRows) + ' rows -> sheet "'
        + gOutName + '" of '
        + gStmf + ' in ' + %char(secs) + ' s';
    if gTrunc;
      txt += ' (truncated at 1048576 rows)';
    endif;
    sendMsg('*COMP' : txt);
  endif;
end-proc;
