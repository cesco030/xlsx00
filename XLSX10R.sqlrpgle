**free
// =====================================================================================
// XLSX10R - read a worksheet of an Excel .xlsx stream file into a Db2 table (pure RPG)
//
// CPP of command XLSX01. The reverse of XLSX00R: the .xlsx is unzipped / inflated by
// XLSX02R (own inflate, no Java, no PASE), the sheet XML is scanned with pointers, the
// values go to Db2 with multi-row INSERT statements fed by an SQLDA of UTF-8 values.
//
// Passes over the sheet XML (kept in teraspace):
//   0  sample the first rows        -> data area, header row (unless given)
//   1  profile every column         -> types for a new table, lengths, CCSID needs
//   2  convert + check every value  -> only for an existing table / TYPES override and
//                                      ERRORS(*STOP): nothing is written when a value
//                                      does not fit
//   3  convert + insert             -> blocks of up to 1000 rows per INSERT; a failing
//                                      row is located with SQLERRD(3), the rest of the
//                                      block goes row by row
//
// Parameters (all after the 4th optional; *omit = default):
//   1 STMF      .xlsx path
//   2 TOFILE    char(20): table (10) + library (10); library *LIBL / *CURLIB; table
//               *NONE = only analyse (report, no database change)
//   3 ROWS      out packed(15:0): rows inserted
//   4 MSG       out char(256): error text (blank = ok). '*ESCAPE' on entry or not
//               passed: errors are sent as escape message CPF9898
//   5 SHEET     *FIRST | sheet name | sheet number
//   6 RANGE     *AUTO | A1:F20 | B3 | A:F | Excel table name | workbook defined name
//   7 HEADER    *AUTO | *YES | *NO | *SKIP (row present, names not used)
//   8 STARTROW  0 = *AUTO, else the first row (the header row when there is one)
//   9 ROWSEL    *ALL | *VISIBLE (skip rows hidden by a filter / hidden rows)
//  10 MAXROWS   0 = all
//  11 MBROPT    *ADD | *REPLACE (delete the rows of the existing table first)
//  12 CRTFILE   *YES (create when missing) | *NO | *REPLACE (drop and create)
//  13 MAP       *AUTO | *NAME | *POS  (Excel columns -> existing table columns)
//  14 COLMAP    'Header or letter=COLUMN;...'  (*SKIP as column = ignore)
//  15 TYPES     'COLUMN=DECIMAL(9,2);...'  types of created columns
//  16 NULLS     *NO (NOT NULL WITH DEFAULT) | *YES  (created columns)
//  17 ERRORS    *STOP | *SKIP (row) | *NULL (value -> null / default, text truncated)
//  18 STRCCSID  *AUTO (job CCSID, 1208 when needed) | *JOB | *UTF8 | *UTF16 | number
//  19 TRIM      *YES | *NO  (leading / trailing blanks of text)
//  20 DATFMT    *DMY | *MDY | *YMD  (text dates like 31/12/2026; ISO always works)
//  21 REPORT    *NONE | stream file for a UTF-8 text report
//
// Build: XLSXDIST.clle (CRTSQLRPGI OBJTYPE(*MODULE) + CRTPGM with XLSX01R, XLSX02R).
// Version 1.4.1 (2026-11-17) - XLSX_VERSION in XLSXINT.rpgleinc; history in XLSX00R.
// =====================================================================================
ctl-opt main(xlsxRead) alloc(*teraspace) bnddir('QC2LE')
        option(*srcstmt : *nodebugio : *nounref) datfmt(*iso) timfmt(*iso);

exec sql set option commit = *none, naming = *sys, datfmt = *iso, timfmt = *iso,
                    decmpt = *period, closqlcsr = *endactgrp, alwblk = *allread;

/copy 'XLSXINT.rpgleinc'

dcl-c SQL_NUM 8000;
exec sql include SQLDA;

dcl-c PGMID   'XLSX10R';
dcl-c MAXC    2000;                     // highest Excel column read
dcl-c MAXA    1000;                     // columns in the data area
dcl-c MAXS    40;                       // sampled rows (pass 0)
dcl-c MAXERR  20;                       // errors kept for messages
dcl-c LOWER   'abcdefghijklmnopqrstuvwxyz';
dcl-c UPPER   'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

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
dcl-pr rtvJobI extpgm('QUSRJOBI');
  rcv char(600) options(*varsize);
  rcvLen int(10) const;
  fmt char(8) const;
  job char(26) const;
  intId char(16) const;
end-pr;

dcl-pr getenv pointer extproc('getenv');
  nm pointer value options(*string);
end-pr;

dcl-pr memset pointer extproc('memset');
  p pointer value;
  c int(10) value;
  n uns(10) value;
end-pr;

dcl-pr memchr pointer extproc('memchr');
  s pointer value;
  c int(10) value;
  n uns(10) value;
end-pr;
dcl-pr memcmp int(10) extproc('memcmp');
  a pointer value;
  b pointer value;
  n uns(10) value;
end-pr;
dcl-pr ifsOpen int(10) extproc('open');
  path  pointer value options(*string);
  oflag int(10) value;
  mode  uns(10) value options(*nopass);
  ccsid uns(10) value options(*nopass);
end-pr;
dcl-pr ifsWrite int(10) extproc('write');
  fd  int(10) value;
  buf pointer value;
  n   uns(10) value;
end-pr;
dcl-pr ifsClose int(10) extproc('close');
  fd int(10) value;
end-pr;
dcl-c O_WRONLY 2;
dcl-c O_CREAT  8;
dcl-c O_CCSID  32;
dcl-c O_TRUNC  64;

dcl-s U_T  varchar(32740) ccsid(*utf8) template;     // a text value
dcl-s N_T  varchar(512) ccsid(*utf8) template;       // names, short texts
dcl-s T_T  varchar(64) ccsid(*utf8) template;        // tokens

// ---- options ----------------------------------------------------------------------
dcl-s gStmf    varchar(1024);
dcl-s gLib     varchar(10);
dcl-s gTab     varchar(10);
dcl-s gNone    ind;                     // TOFILE(*NONE): analysis only
dcl-s gSheetSel like(N_T);              // '' = *FIRST
dcl-s gRange   varchar(256);
dcl-s gHdrOpt  char(1);                 // A auto, Y yes, N no, S skip
dcl-s gStart   int(10);
dcl-s gVis     ind;
dcl-s gMaxRows int(20);
dcl-s gMbrRpl  ind;
dcl-s gCrt     char(1);                 // Y create when missing, N no, R drop+create
dcl-s gMapOpt  char(1);                 // A auto, N name, P position
dcl-s gColMap  varchar(2000);
dcl-s gTypes   varchar(2000);
dcl-s gNulls   ind;
dcl-s gErrMode char(1);                 // S stop, K skip row, N null value
dcl-s gCcsOpt  int(10);                 // 0 auto, -1 job, 1208, 1200, other number
dcl-s gTrim    ind;
dcl-s gDatFmt  char(3);                 // DMY MDY YMD
dcl-s gRepPath varchar(1024);

// ---- state ------------------------------------------------------------------------
dcl-s gFail    ind;
dcl-s gMsg     varchar(512);
dcl-s gZrOpen  ind;
dcl-s gStmt    varchar(32740);
dcl-s gRep     varchar(4000000 : 4) ccsid(*utf8);   // report text
dcl-s gDate1904 ind;
// XML tokens in UTF-8, converted once (u() in the scanner loops costs a conversion
// and a 32 KB copy per cell)
dcl-s kC      like(T_T) inz('c');
dcl-s kRow    like(T_T) inz('row');
dcl-s kR      like(T_T) inz('r');
dcl-s kHidden like(T_T) inz('hidden');
dcl-s kOne    like(T_T) inz('1');
dcl-s kTrue   like(T_T) inz('true');
dcl-s kIs     like(T_T) inz('is');
dcl-s kRph    like(T_T) inz('rPh');
dcl-s kT      like(T_T) inz('t');
dcl-s kF      like(T_T) inz('f');
dcl-s kSi     like(T_T) inz('si');
dcl-s gGrpCh char(1);                   // job thousands separator (. or ,)
dcl-s gA2 uns(3) dim(256);                // EBCDIC -> ASCII (putE), 0 = no
dcl-s gA2Ok ind;
dcl-s gRawE ind;                         // putE: keep job CCSID bytes
dcl-s gTm int(20) dim(6);                 // ms: load, pass 0..3, Db2 INSERT
dcl-s gTq timestamp;
dcl-s gTi timestamp;
dcl-s gTu int(20);                       // microseconds in Db2 INSERT
dcl-s dcKey int(10) dim(4096);            // serialDT cache: serial + 1
dcl-s dcIso char(10) dim(4096);
dcl-s gCreated ind;                     // the table was created by this run
dcl-s gExists  ind;
dcl-s gToday   date;
dcl-s gNowTs   timestamp;

// ---- workbook -----------------------------------------------------------------------
dcl-s nSh int(10);
dcl-ds sh qualified dim(256);
  name like(N_T);
  rid  like(T_T);
  part like(N_T);
  state char(1);                        // V visible, H hidden, X very hidden
end-ds;
dcl-s gShIdx int(10);
dcl-s nDn int(10);
dcl-ds dn qualified dim(500);
  name like(N_T);
  lsid int(10);
  ref  like(N_T);
end-ds;
dcl-s nRel int(10);
dcl-ds rl qualified dim(500);
  id   like(T_T);
  typ  like(N_T);
  tgt  like(N_T);
end-ds;
dcl-s gWbPart like(N_T);
dcl-s gStyPart like(N_T);
dcl-s gSstPart like(N_T);

// ---- styles: cell format kind per xf ---------------------------------------------
dcl-s nXf int(10);
dcl-s xfKind char(1) dim(30000);        // N number, D date, Z date+time, T time
dcl-s xfPad  int(5) dim(30000);         // "00000" -> 5 (zero padded codes)
dcl-s nNf int(10);
dcl-s nfId   int(10) dim(2000);
dcl-s nfKind char(1) dim(2000);
dcl-s nfPad  int(5) dim(2000);

// ---- shared strings (decoded UTF-8 in one buffer) --------------------------------
dcl-s nSst int(10);
dcl-s pSstBuf pointer;
dcl-s sstCap int(10);
dcl-s sstUse int(10);
dcl-s pSstOff pointer;
dcl-s pSstLen pointer;
dcl-s pSstChr pointer;
dcl-s pSstU16 pointer;
dcl-s pSstNL  pointer;
dcl-s sstOff int(10) dim(4000000) based(pSstOff);
dcl-s sstLen int(10) dim(4000000) based(pSstLen);
dcl-s sstChr int(10) dim(4000000) based(pSstChr);
dcl-s sstU16 int(10) dim(4000000) based(pSstU16);
dcl-s sstNL  char(1) dim(4000000) based(pSstNL);   // N needs Unicode, L Latin-1

// ---- the sheet XML ------------------------------------------------------------------
dcl-s pSh pointer;
dcl-s nShLen int(10);
dcl-s gSdFrom int(10);                  // first byte inside <sheetData>
dcl-s gSdTo int(10);                    // position of </sheetData>
dcl-s gAfRef like(T_T);                 // <autoFilter ref>
dcl-s gStop ind;                        // scanning ends early
dcl-s gPass int(10);

// ---- current row (cells by Excel column) -------------------------------------------
dcl-s gGen int(10);
dcl-s gRowNo int(20);
dcl-s gRowHid ind;
dcl-s gRowMax int(10);                  // highest column in the row
dcl-s gBeyond int(20);                  // cells right of MAXC
dcl-s rcGen  int(10) dim(2000);
dcl-s rcKind char(1) dim(2000);         // N S B X E d(ISO date text) s(shared)
dcl-s rcSty  int(10) dim(2000);
dcl-s rcP    pointer dim(2000);         // raw <v> bytes in the sheet XML
dcl-s rcL    int(10) dim(2000);
dcl-s rcSst  int(10) dim(2000);         // shared string index (0-based)
dcl-s rcTo   int(10) dim(2000);         // decoded text in the row text buffer
dcl-s rcTl   int(10) dim(2000);
dcl-s pRowTxt pointer;
dcl-s rowTxtCap int(10);
dcl-s rowTxtUse int(10);

// ---- pass 0 sample --------------------------------------------------------------
dcl-s nSm int(10);
dcl-s gSmFrom int(20);                  // sampling starts at this row
dcl-ds sm qualified dim(40);
  row int(20);
  cnt int(10);
  c1  int(10);
  c2  int(10);
end-ds;
dcl-s pSmK pointer;                     // kind per (sample, column)
dcl-s pSmO pointer;                     // text offset / length
dcl-s pSmL pointer;
dcl-s pSmBuf pointer;
dcl-s smUse int(10);
dcl-s smKind char(1) dim(80000) based(pSmK);
dcl-s smOff int(10) dim(80000) based(pSmO);
dcl-s smLen int(10) dim(80000) based(pSmL);

// ---- data area ---------------------------------------------------------------------
dcl-s aHdr  int(20);                    // header row, 0 = none
dcl-s aFirst int(20);                   // first data row
dcl-s aLast int(20);                    // last row of the area, 0 = to the end
dcl-s aC1 int(10);
dcl-s aC2 int(10);
dcl-s nAc int(10);
dcl-s gAreaHow varchar(80);             // how the area was found (report)
dcl-s gHdrHow varchar(80);
dcl-s gLastSeen int(20);                // last row with data seen in pass 1
dcl-s gAfExt ind;                       // autoFilter area: extend over contiguous rows

dcl-ds ac qualified dim(1000);          // area columns
  xc    int(10);                        // Excel column number
  let   varchar(3);
  hdr   like(N_T);
  nm    varchar(30);                    // SQL name for a new table
  sys   varchar(10);
  cN    int(20);                        // numbers (not date formatted)
  cInt  int(20);
  cD    int(20);                        // date formatted, whole days
  cZ    int(20);                        // date + time
  cT    int(20);                        // time only
  cB    int(20);
  cS    int(20);                        // text
  cSD   int(20);                        // text that is a date
  cSZ   int(20);                        // text that is a date + time
  cX    int(20);                        // error values (#N/A ...)
  cE    int(20);                        // empty
  iDig  int(10);                        // max integer digits
  sc    int(10);                        // max scale
  mxU8  int(10);                        // max rendered length: UTF-8 bytes
  mxCh  int(10);                        //   characters
  mxU16 int(10);                        //   UTF-16 units
  nonLat ind;                           // text needs Unicode
  pad   int(10);                        // zero padded number format
  ty    varchar(80);                    // SQL type for CREATE
  ti    int(10);                        // target column, 0 = not loaded
end-ds;

// ---- target table ------------------------------------------------------------------
dcl-s nTc int(10);
dcl-ds tc qualified dim(1000);
  name  varchar(128);
  sys   varchar(10);
  dtype varchar(10);
  len   int(10);
  prec  int(10);
  scale int(10);
  nul   ind;
  ccsid int(10);
  txt   varchar(50);
  hdg   varchar(60);
  skip  ind;                            // generated always / not loadable
  dft   varchar(256);                   // value used for an empty cell (not nullable)
  kind  char(1);                        // C G N I F D T Z B X
  unit  char(1);                        // C: S single byte, U UTF-8 bytes, G UTF-16
  ebc   ind;                            // CCSID 65535: value passed in job CCSID
  imax  varchar(20);                    // integer limit (digits)
  ai    int(10);                        // area column loaded into it
  hasDf char(1);                        // HAS_DEFAULT
end-ds;

// ---- load -------------------------------------------------------------------------
dcl-s nMp int(10);
dcl-s mpA int(10) dim(1000);            // mapped: area column
dcl-s mpT int(10) dim(1000);            //         target column
dcl-s mpCap int(10) dim(1000);          // slot capacity (bytes)
dcl-s mpOff int(10) dim(1000);          // offset in a row slot
dcl-s gRowBytes int(10);
dcl-s gBlk int(10);                     // rows per INSERT
dcl-s pPar pointer;                     // value slots
dcl-s pInd pointer;                     // indicators
dcl-s nBuf int(10);                     // rows in the current block
dcl-s bRow int(20) dim(1000);           // Excel row of each block row
dcl-s gPrepB int(10);                   // rows of the prepared S1
// staging load: each block goes with one blocked INSERT (n ROWS, a plain copy) into
// QTEMP/XLSX10S, one row = one UTF-8 string: Excel row (10 digits), then fixed slots
// [flag N/0][5 digits length][value]. Every chunk (up to 64 MB) one INSERT ... SELECT
// with casts moves the rows into the table (set-based, ~5x faster than multi-row
// VALUES). A Db2 error stops it at row k+1 (SQLERRD(3) = k): reported, then the rest is
// retried with WHERE RRN > k+1. Not used for rows over 8000 bytes, CCSID 65535 columns
// or a statement over 32 KB: then the VALUES insert is used.
dcl-ds stg qualified dim(400);          // Db2 copies whole elements: 3 widths
  r varchar(8000) ccsid(*utf8);
end-ds;
dcl-ds stgM qualified dim(1000);
  r varchar(4000) ccsid(*utf8);
end-ds;
dcl-ds stgS qualified dim(1000);
  r varchar(1000) ccsid(*utf8);
end-ds;
dcl-s gStgK char(1);                    // S / M / L
dcl-s gStgOn ind;
dcl-s gStgMade ind;                      // QTEMP/XLSX10S created by this run
dcl-s gStgW int(10);                    // staging row width
dcl-s gStgN int(20);                    // rows in the staging table
dcl-s gChunk int(20);                   // staging rows per INSERT ... SELECT
dcl-s stOff int(10) dim(1000);          // slot position (1-based) in the staging row
dcl-s gStgWhy varchar(120);
dcl-s gDry ind;                         // pass 2: check only

// ---- counters -----------------------------------------------------------------------
dcl-s gRead int(20);                    // data rows read
dcl-s gIns int(20);
dcl-s gRej int(20);                     // rows not loaded (errors)
dcl-s gNulled int(20);                  // values set to null / default by ERRORS(*NULL)
dcl-s gSubst int(20);                   // values with characters not in an EBCDIC CCSID
dcl-s gErrVal int(20);                  // #N/A etc. loaded as empty
dcl-s gErrN int(20);                    // value errors
dcl-s gErrTxt varchar(300) dim(20);
dcl-s gRowBad ind;

// ---- value conversion work areas ------------------------------------------------------
dcl-ds num_t qualified template;
  neg ind;
  d   char(40);                         // digits, no leading / trailing zeros
  nd  int(10);
  ex  int(10);                          // value = d * 10**ex
end-ds;
dcl-s gVal char(32766) ccsid(*utf8);    // rendered value
dcl-s gVlen int(10);
dcl-s gVnull ind;
dcl-s gDec char(65535) ccsid(*utf8);    // decode work area

dcl-s rR1 int(20);                      // RANGE: rows / columns, 0 = open
dcl-s rR2 int(20);
dcl-s rC1 int(10);
dcl-s rC2 int(10);
dcl-s rHdr int(10);                     // Excel table: header rows (-1 unknown)
dcl-s rTot int(10);                     //              totals rows
dcl-s rHow varchar(80);
dcl-s nTy int(10);
dcl-s tyKey varchar(128) dim(500);
dcl-s tyVal varchar(80) dim(500);
dcl-s gMapHow varchar(80);

dcl-ds dt_t qualified template;
  dd  date;
  iso char(10);                         // dd as ISO text when known (serial cache)
  sec int(10);                          // seconds of the day
  us  int(10);                          // microseconds
  hasD ind;
  hasT ind;
end-ds;

dcl-s gRead2 int(20);                   // data rows of passes 2 / 3

// procedures called before their definition
dcl-pr rowEnd;
end-pr;
dcl-pr sampleRow;
end-pr;
dcl-pr statRow;
end-pr;
dcl-pr loadRow;
end-pr;
dcl-pr flush;
end-pr;

// SQLVAR overlay
dcl-s pv pointer;
dcl-ds sv qualified based(pv) align;
  vType    int(5)  pos(1);
  vLen     int(5)  pos(3);
  vData    pointer pos(17);
  vInd     pointer pos(33);
  vNameLen int(5)  pos(49);
  vName    char(30) pos(51);
end-ds;

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

dcl-proc grpCh;                         // job DECFMT: blank -> 1,000.5  I/J -> 1.000,5
  dcl-pi *n char(1);
  end-pi;
  dcl-s r char(600);
  if gGrpCh = ' ';
    gGrpCh = '.';
    monitor;
      rtvJobI(r : %size(r) : 'JOBI0400' : '*' : ' ');
      if %subst(r : 457 : 1) = ' ';
        gGrpCh = ',';
      endif;
    on-error;
    endmon;
  endif;
  return gGrpCh;
end-proc;

dcl-proc left;                          // first n characters (safe for short strings)
  dcl-pi *n varchar(32740) rtnparm;
    s varchar(32740) const;
    n int(10) value;
  end-pi;
  if %len(s) <= n;
    return s;
  endif;
  return %subst(s : 1 : n);
end-proc;

dcl-proc u;                             // job CCSID -> UTF-8
  dcl-pi *n like(U_T) rtnparm;
    s varchar(32740) const;
  end-pi;
  dcl-s r like(U_T);
  r = s;
  return r;
end-proc;

dcl-proc ebc;                           // UTF-8 -> job CCSID
  dcl-pi *n varchar(32740) rtnparm;
    s like(U_T) const;
  end-pi;
  dcl-s r varchar(32740);
  r = s;
  return r;
end-proc;

dcl-proc rep;                           // report line
  dcl-pi *n;
    s varchar(32740) const;
  end-pi;
  if %len(gRep) + %len(s) * 3 + 2 < 3990000;
    gRep += u(s) + x'0A';
  endif;
end-proc;

dcl-proc repU;                          // report line already UTF-8
  dcl-pi *n;
    s like(U_T) const;
  end-pi;
  if %len(gRep) + %len(s) + 2 < 3990000;
    gRep += s + x'0A';
  endif;
end-proc;

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

dcl-proc colLetters;                    // 28 -> AB
  dcl-pi *n varchar(3);
    c int(10) value;
  end-pi;
  dcl-s r varchar(3);
  dcl-s k int(10);
  r = '';
  dow c > 0;
    k = %rem(c - 1 : 26);
    r = %subst(UPPER : k + 1 : 1) + r;
    c = %div(c - 1 : 26);
  enddo;
  return r;
end-proc;

dcl-proc cellRef;                       // B12 -> col 2, row 12 (0 when not a ref)
  dcl-pi *n ind;
    s varchar(64) const;
    col int(10);
    row int(20);
  end-pi;
  dcl-s i int(10);
  dcl-s ch char(1);
  dcl-s k int(10);
  col = 0;
  row = 0;
  for i = 1 to %len(s);
    ch = %subst(s : i : 1);
    if ch = '$';
      iter;
    endif;
    k = %scan(%xlate(LOWER : UPPER : ch) : UPPER);
    if k > 0;
      if row > 0;
        return *off;
      endif;
      col = col * 26 + k;
    elseif ch >= '0' and ch <= '9';
      row = row * 10 + (%scan(ch : '0123456789') - 1);
    else;
      return *off;
    endif;
  endfor;
  return (col > 0 or row > 0);
end-proc;

dcl-proc onlyDigits;
  dcl-pi *n ind;
    s varchar(64) const;
  end-pi;
  return %len(s) > 0 and %check('0123456789' : s) = 0;
end-proc;

dcl-proc toIntU;                        // UTF-8 digits -> integer (0 = not digits)
  dcl-pi *n int(20);
    v like(N_T) const;
  end-pi;
  dcl-s b uns(3) based(pb);
  dcl-s pb pointer;
  dcl-s i int(10);
  dcl-s r int(20);
  dcl-s w varchar(32) ccsid(*utf8);
  if %len(v) = 0 or %len(v) > 18;
    return 0;
  endif;
  w = v;
  pb = %addr(w : *data);
  r = 0;
  for i = 1 to %len(v);
    if b < 48 or b > 57;
      return 0;
    endif;
    r = r * 10 + b - 48;
    pb += 1;
  endfor;
  return r;
end-proc;

dcl-proc toInt;
  dcl-pi *n int(20);
    s varchar(64) const;
  end-pi;
  if not onlyDigits(s) or %len(s) > 18;
    return 0;
  endif;
  return %int(s);
end-proc;

// bytes of a UTF-8 text: characters, UTF-16 units, needs more than Latin-1 + euro
dcl-proc utf8Info;
  dcl-pi *n;
    p pointer value;
    n int(10) value;
    nch int(10);
    nu16 int(10);
    nonLat ind;
  end-pi;
  dcl-s b uns(3) based(pb);
  dcl-s pb pointer;
  dcl-s b2 uns(3) based(pb2);
  dcl-s b3 uns(3) based(pb3);
  dcl-s pb2 pointer;
  dcl-s pb3 pointer;
  dcl-s i int(10);
  nch = 0;
  nu16 = 0;
  nonLat = *off;
  i = 0;
  dow i < n;
    pb = p + i;
    if b < 128;
      i += 1;
    elseif b < 224;                     // 2 bytes: U+0080..U+07FF
      if b > 195;                       // above U+00FF
        nonLat = *on;
      endif;
      i += 2;
    elseif b < 240;                     // 3 bytes
      if not nonLat;
        pb2 = p + i + 1;
        pb3 = p + i + 2;
        if not (b = 226 and b2 = 130 and b3 = 172);       // euro sign
          nonLat = *on;
        endif;
      endif;
      i += 3;
    else;                               // 4 bytes: surrogate pair
      nonLat = *on;
      nu16 += 1;
      i += 4;
    endif;
    nch += 1;
    nu16 += 1;
  enddo;
end-proc;

// code point -> UTF-8 bytes at p, returns the number of bytes
dcl-proc putCp;
  dcl-pi *n int(10);
    p pointer value;
    cp int(10) value;
  end-pi;
  dcl-s b uns(3) based(pb);
  dcl-s pb pointer;
  pb = p;
  if cp < 128;
    b = cp;
    return 1;
  elseif cp < 2048;
    b = 192 + %div(cp : 64);
    pb += 1;
    b = 128 + %rem(cp : 64);
    return 2;
  elseif cp < 65536;
    b = 224 + %div(cp : 4096);
    pb += 1;
    b = 128 + %rem(%div(cp : 64) : 64);
    pb += 1;
    b = 128 + %rem(cp : 64);
    return 3;
  endif;
  b = 240 + %div(cp : 262144);
  pb += 1;
  b = 128 + %rem(%div(cp : 4096) : 64);
  pb += 1;
  b = 128 + %rem(%div(cp : 64) : 64);
  pb += 1;
  b = 128 + %rem(cp : 64);
  return 4;
end-proc;

dcl-proc hexVal;                        // ASCII hex digit -> 0..15, -1
  dcl-pi *n int(10);
    b uns(3) value;
  end-pi;
  if b >= 48 and b <= 57;
    return b - 48;
  elseif b >= 65 and b <= 70;
    return b - 55;
  elseif b >= 97 and b <= 102;
    return b - 87;
  endif;
  return -1;
end-proc;

// named XML entity (amp lt gt quot apos) -> code point, -1 unknown
dcl-proc entCp;
  dcl-pi *n int(10);
    p pointer value;
    n int(10) value;
  end-pi;
  dcl-s c1 uns(3) based(p1);
  dcl-s c2 uns(3) based(p2);
  dcl-s c3 uns(3) based(p3);
  dcl-s c4 uns(3) based(p4);
  dcl-s p1 pointer;
  dcl-s p2 pointer;
  dcl-s p3 pointer;
  dcl-s p4 pointer;
  p1 = p;
  p2 = p + 1;
  p3 = p + 2;
  p4 = p + 3;
  if n = 2 and c2 = 116;                         // lt gt
    if c1 = 108;
      return 60;
    elseif c1 = 103;
      return 62;
    endif;
  elseif n = 3 and c1 = 97 and c2 = 109 and c3 = 112;        // amp
    return 38;
  elseif n = 4 and c1 = 113 and c2 = 117 and c3 = 111 and c4 = 116;   // quot
    return 34;
  elseif n = 4 and c1 = 97 and c2 = 112 and c3 = 111 and c4 = 115;    // apos
    return 39;
  endif;
  return -1;
end-proc;

// XML text (entities) and Excel _xHHHH_ escapes -> UTF-8 at dst; returns length
dcl-proc xdecode;
  dcl-pi *n int(10);
    src pointer value;
    n int(10) value;
    dst pointer value;
    cap int(10) value;
  end-pi;
  dcl-s b uns(3) based(pb);
  dcl-s pb pointer;
  dcl-s o uns(3) based(po);
  dcl-s po pointer;
  dcl-s i int(10);
  dcl-s k int(10);
  dcl-s j int(10);
  dcl-s cp int(10);
  dcl-s h int(10);
  dcl-s ok ind;
  k = 0;
  i = 0;
  if memchr(src : 38 : n) = *null and memchr(src : 95 : n) = *null;
    if n > cap;
      n = cap;
    endif;
    memcpy(dst : src : n);
    return n;
  endif;
  dow i < n and k < cap - 4;
    pb = src + i;
    if b = 38;                          // &
      j = i + 1;
      dow j < n and j - i < 12;
        pb = src + j;
        if b = 59;                      // ;
          leave;
        endif;
        j += 1;
      enddo;
      pb = src + j;
      if j < n and b = 59 and j - i >= 3;
        cp = -1;
        pb = src + i + 1;
        if b = 35;                      // &#
          cp = 0;
          pb = src + i + 2;
          if b = 120 or b = 88;         // &#x
            for h = i + 3 to j - 1;
              pb = src + h;
              if hexVal(b) < 0;
                cp = -1;
                leave;
              endif;
              cp = cp * 16 + hexVal(b);
            endfor;
          else;
            for h = i + 2 to j - 1;
              pb = src + h;
              if b < 48 or b > 57;
                cp = -1;
                leave;
              endif;
              cp = cp * 10 + b - 48;
            endfor;
          endif;
        else;
          cp = entCp(src + i + 1 : j - i - 1);
        endif;
        if cp >= 0;
          k += putCp(dst + k : cp);
          i = j + 1;
          iter;
        endif;
      endif;
      pb = src + i;
    elseif b = 95 and i + 6 < n;        // _xHHHH_
      ok = *off;
      pb = src + i + 1;
      if b = 120;
        pb = src + i + 6;
        if b = 95;
          cp = 0;
          ok = *on;
          for h = i + 2 to i + 5;
            pb = src + h;
            if hexVal(b) < 0;
              ok = *off;
              leave;
            endif;
            cp = cp * 16 + hexVal(b);
          endfor;
        endif;
      endif;
      if ok;
        k += putCp(dst + k : cp);
        i += 7;
        iter;
      endif;
      pb = src + i;
    endif;
    po = dst + k;
    o = b;
    k += 1;
    i += 1;
  enddo;
  return k;
end-proc;

// =====================================================================================
// XML scanning on a buffer (b = base pointer, n = length, positions are 0-based)
// =====================================================================================
// local name of the element at p ('<' at p): start offset and length (prefix skipped)
dcl-proc locName;
  dcl-pi *n int(10);                    // length of the local name
    b pointer value;
    n int(10) value;
    p int(10) value;
    st int(10);                         // out: start of the local name
  end-pi;
  dcl-s c uns(3) based(pc);
  dcl-s pc pointer;
  dcl-s q int(10);
  q = p + 1;
  st = q;
  dow q < n;
    pc = b + q;
    if c = 32 or c = 62 or c = 47 or c = 9 or c = 10 or c = 13;
      leave;
    endif;
    if c = 58;                          // ':' -> prefix
      st = q + 1;
    endif;
    q += 1;
  enddo;
  return q - st;
end-proc;

dcl-proc isName;                        // local name at p equals nm
  dcl-pi *n ind;
    b pointer value;
    n int(10) value;
    p int(10) value;
    nm like(T_T) const;
  end-pi;
  dcl-s t like(T_T);
  dcl-s st int(10);
  dcl-s l int(10);
  l = locName(b : n : p : st);
  if l <> %len(nm);
    return *off;
  endif;
  t = nm;
  return memcmp(b + st : %addr(t : *data) : l) = 0;
end-proc;

// next start tag with local name nm at or after 'from' (position of '<'), -1
dcl-proc findTag;
  dcl-pi *n int(10);
    b pointer value;
    n int(10) value;
    from int(10) value;
    nm like(T_T) const;
  end-pi;
  dcl-s c uns(3) based(pc);
  dcl-s pc pointer;
  dcl-s pp pointer;
  dcl-s p int(10);
  p = from;
  dow p < n - 1;
    pp = memchr(b + p : 60 : n - p);
    if pp = *null;
      return -1;
    endif;
    p = pp - b;
    pc = pp + 1;
    if c <> 47 and c <> 63 and c <> 33;  // not </ <? <!
      if isName(b : n : p : nm);
        return p;
      endif;
    endif;
    p += 1;
  enddo;
  return -1;
end-proc;

// next end tag </nm> at or after from, -1
dcl-proc findEnd;
  dcl-pi *n int(10);
    b pointer value;
    n int(10) value;
    from int(10) value;
    nm like(T_T) const;
  end-pi;
  dcl-s c uns(3) based(pc);
  dcl-s pc pointer;
  dcl-s pp pointer;
  dcl-s p int(10);
  p = from;
  dow p < n - 1;
    pp = memchr(b + p : 60 : n - p);
    if pp = *null;
      return -1;
    endif;
    p = pp - b;
    pc = pp + 1;
    if c = 47 and isName(b : n : p + 1 : nm);   // '/' then name: shift by one
      return p;
    endif;
    p += 1;
  enddo;
  return -1;
end-proc;

// position after the '>' of the tag at p; selfc = tag ends with '/>'
dcl-proc tagEnd;
  dcl-pi *n int(10);
    b pointer value;
    n int(10) value;
    p int(10) value;
    selfc ind;
  end-pi;
  dcl-s c uns(3) based(pc);
  dcl-s pc pointer;
  dcl-s pp pointer;
  dcl-s q int(10);
  selfc = *off;
  pp = memchr(b + p : 62 : n - p);
  if pp = *null;
    return -1;
  endif;
  q = pp - b;
  pc = pp - 1;
  selfc = (c = 47);
  return q + 1;
end-proc;

// attribute value of the tag [p, e); name compared on its local part when anyPfx
dcl-proc attr;
  dcl-pi *n like(N_T);
    b pointer value;
    p int(10) value;
    e int(10) value;
    nm like(T_T) const;
    anyPfx ind value options(*nopass);
  end-pi;
  dcl-s c uns(3) based(pc);
  dcl-s pc pointer;
  dcl-s q int(10);
  dcl-s s int(10);
  dcl-s ls int(10);
  dcl-s v0 int(10);
  dcl-s quo uns(3);
  dcl-s t like(T_T);
  dcl-s r like(N_T);
  dcl-s anyP ind;
  dcl-s l int(10);
  anyP = *off;
  if %parms() >= 5;
    anyP = anyPfx;
  endif;
  t = nm;
  q = p + 1;
  dow q < e;                            // skip the element name
    pc = b + q;
    if c = 32 or c = 9 or c = 10 or c = 13 or c = 62 or c = 47;
      leave;
    endif;
    q += 1;
  enddo;
  dow q < e;
    pc = b + q;
    if c = 32 or c = 9 or c = 10 or c = 13;
      q += 1;
      iter;
    endif;
    if c = 62 or c = 47;
      leave;
    endif;
    s = q;
    ls = q;
    dow q < e;
      pc = b + q;
      if c = 61 or c = 32 or c = 9 or c = 10 or c = 13;
        leave;
      endif;
      if c = 58;
        ls = q + 1;
      endif;
      q += 1;
    enddo;
    l = q - s;
    dow q < e;                          // to the '='
      pc = b + q;
      if c = 61;
        leave;
      endif;
      q += 1;
    enddo;
    q += 1;
    dow q < e;                          // to the quote
      pc = b + q;
      if c = 34 or c = 39;
        leave;
      endif;
      q += 1;
    enddo;
    if q >= e;
      leave;
    endif;
    quo = c;
    q += 1;
    v0 = q;
    dow q < e;
      pc = b + q;
      if c = quo;
        leave;
      endif;
      q += 1;
    enddo;
    if (l = %len(t) and memcmp(b + s : %addr(t : *data) : l) = 0)
       or (anyP and q - ls > 0 and ls > s and %len(t) = s + l - ls
           and memcmp(b + ls : %addr(t : *data) : %len(t)) = 0);
      l = xdecode(b + v0 : q - v0 : %addr(gDec) : 512);
      r = %subst(gDec : 1 : l);
      return r;
    endif;
    q += 1;
  enddo;
  return r;
end-proc;

// decoded text from position p (after a '>') up to the next '<'
dcl-proc textAt;
  dcl-pi *n like(U_T);
    b pointer value;
    n int(10) value;
    p int(10) value;
  end-pi;
  dcl-s pp pointer;
  dcl-s q int(10);
  dcl-s l int(10);
  dcl-s r like(U_T);
  pp = memchr(b + p : 60 : n - p);
  if pp = *null;
    q = n;
  else;
    q = pp - b;
  endif;
  l = xdecode(b + p : q - p : %addr(gDec) : 32740);
  r = %subst(gDec : 1 : l);
  return r;
end-proc;

// =====================================================================================
// package parts
// =====================================================================================
dcl-proc partGet;                       // inflated part, *null when missing
  dcl-pi *n pointer;
    name like(N_T) const;
    len int(10);
  end-pi;
  dcl-s i int(10);
  dcl-s nm varchar(512) ccsid(*utf8);
  len = 0;
  nm = name;
  if %len(nm) > 0 and %subst(nm : 1 : 1) = u('/');
    nm = %subst(nm : 2);
  endif;
  i = zrFind(nm);
  if i <= 0;
    return *null;
  endif;
  return zrGet(i : len);
end-proc;

// 'xl/worksheets/' + '../tables/t1.xml' -> 'xl/tables/t1.xml'
dcl-proc resolve;
  dcl-pi *n like(N_T);
    base like(N_T) const;               // part whose relationships these are
    tgt like(N_T) const;
  end-pi;
  dcl-s d like(N_T);
  dcl-s t like(N_T);
  dcl-s k int(10);
  dcl-s j int(10);
  dcl-s sl like(T_T);
  sl = u('/');
  t = tgt;
  if %len(t) > 0 and %subst(t : 1 : 1) = sl;
    return %subst(t : 2);
  endif;
  k = 0;
  j = %scan(sl : base);
  dow j > 0;
    k = j;
    if j >= %len(base);
      leave;
    endif;
    j = %scan(sl : base : j + 1);
  enddo;
  d = '';
  if k > 0;
    d = %subst(base : 1 : k);
  endif;
  dow %len(t) > 3 and %subst(t : 1 : 3) = u('../');
    t = %subst(t : 4);
    if %len(d) > 0;
      d = %subst(d : 1 : %len(d) - 1);
      k = 0;
      j = %scan(sl : d);
      dow j > 0;
        k = j;
        if j >= %len(d);
          leave;
        endif;
        j = %scan(sl : d : j + 1);
      enddo;
      if k > 0;
        d = %subst(d : 1 : k);
      else;
        d = '';
      endif;
    endif;
  enddo;
  return d + t;
end-proc;

// relationships of part -> rl()
dcl-proc relsLoad;
  dcl-pi *n int(10);
    part like(N_T) const;
  end-pi;
  dcl-s nm like(N_T);
  dcl-s k int(10);
  dcl-s j int(10);
  dcl-s b pointer;
  dcl-s n int(10);
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s sc ind;
  dcl-s sl like(T_T);
  nRel = 0;
  sl = u('/');
  k = 0;
  j = %scan(sl : part);
  dow j > 0;
    k = j;
    if j >= %len(part);
      leave;
    endif;
    j = %scan(sl : part : j + 1);
  enddo;
  if k > 0;
    nm = %subst(part : 1 : k) + u('_rels/') + %subst(part : k + 1) + u('.rels');
  else;
    nm = u('_rels/') + part + u('.rels');
  endif;
  b = partGet(nm : n);
  if b = *null;
    return 0;
  endif;
  p = 0;
  dow nRel < %elem(rl);
    p = findTag(b : n : p : u('Relationship'));
    if p < 0;
      leave;
    endif;
    e = tagEnd(b : n : p : sc);
    if e < 0;
      leave;
    endif;
    nRel += 1;
    rl(nRel).id = attr(b : p : e : u('Id'));
    rl(nRel).typ = attr(b : p : e : u('Type'));
    if attr(b : p : e : u('TargetMode')) = u('External');
      rl(nRel).tgt = '';
    else;
      rl(nRel).tgt = resolve(part : attr(b : p : e : u('Target')));
    endif;
    p = e;
  enddo;
  dealloc(n) b;
  return nRel;
end-proc;

dcl-proc relByType;                     // first target whose type ends with suf
  dcl-pi *n like(N_T);
    suf like(T_T) const;
  end-pi;
  dcl-s i int(10);
  dcl-s t like(N_T);
  for i = 1 to nRel;
    t = rl(i).typ;
    if %len(t) >= %len(suf);
      if %subst(t : %len(t) - %len(suf) + 1) = suf;
        return rl(i).tgt;
      endif;
    endif;
  endfor;
  return '';
end-proc;

// =====================================================================================
// workbook.xml: sheets, date system, defined names
// =====================================================================================
dcl-proc loadWorkbook;
  dcl-pi *n ind;
  end-pi;
  dcl-s b pointer;
  dcl-s n int(10);
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s q int(10);
  dcl-s i int(10);
  dcl-s sc ind;
  dcl-s v like(N_T);

  gWbPart = u('xl/workbook.xml');
  if relsLoad(u('')) > 0;               // _rels/.rels -> office document
    v = relByType(u('/officeDocument'));
    if v <> '';
      gWbPart = v;
    endif;
  endif;
  b = partGet(gWbPart : n);
  if b = *null;
    fail('not an Excel workbook (no ' + ebc(gWbPart) + ')');
    return *off;
  endif;
  gDate1904 = *off;
  p = findTag(b : n : 0 : u('workbookPr'));
  if p >= 0;
    e = tagEnd(b : n : p : sc);
    v = attr(b : p : e : u('date1904'));
    gDate1904 = (v = kOne or v = kTrue);
  endif;
  nSh = 0;
  p = findTag(b : n : 0 : u('sheets'));
  dow p >= 0 and nSh < %elem(sh);
    p = findTag(b : n : p + 1 : u('sheet'));
    if p < 0;
      leave;
    endif;
    e = tagEnd(b : n : p : sc);
    nSh += 1;
    sh(nSh).name = attr(b : p : e : u('name'));
    sh(nSh).rid = attr(b : p : e : u('id') : *on);
    v = attr(b : p : e : u('state'));
    select;
    when v = kHidden;
      sh(nSh).state = 'H';
    when v = u('veryHidden');
      sh(nSh).state = 'X';
    other;
      sh(nSh).state = 'V';
    endsl;
    p = e - 1;
  enddo;
  nDn = 0;
  p = 0;
  dow nDn < %elem(dn);
    p = findTag(b : n : p : u('definedName'));
    if p < 0;
      leave;
    endif;
    e = tagEnd(b : n : p : sc);
    nDn += 1;
    dn(nDn).name = attr(b : p : e : u('name'));
    v = attr(b : p : e : u('localSheetId'));
    dn(nDn).lsid = -1;
    if onlyDigits(ebc(v));
      dn(nDn).lsid = toInt(ebc(v));
    endif;
    dn(nDn).ref = '';
    if not sc;
      dn(nDn).ref = textAt(b : n : e);
    endif;
    p = e;
  enddo;
  dealloc(n) b;

  // workbook relationships: sheet parts, styles, shared strings
  relsLoad(gWbPart);
  for i = 1 to nSh;
    sh(i).part = '';
    for q = 1 to nRel;
      if rl(q).id = sh(i).rid;
        sh(i).part = rl(q).tgt;
        leave;
      endif;
    endfor;
  endfor;
  gStyPart = relByType(u('/styles'));
  gSstPart = relByType(u('/sharedStrings'));
  if nSh = 0;
    fail('the workbook has no sheets');
    return *off;
  endif;
  return *on;
end-proc;

// =====================================================================================
// styles.xml: which cell formats are dates / times
// =====================================================================================
// N number, D date, Z date + time, T time; elapsed formats like [h]:mm are numbers
dcl-proc fmtKind;
  dcl-pi *n char(1);
    code like(N_T) const;
    pad int(5);
  end-pi;
  dcl-s s varchar(512);
  dcl-s c varchar(512);
  dcl-s i int(10);
  dcl-s ch char(1);
  dcl-s inQ ind;
  dcl-s hasD ind;
  dcl-s hasT ind;
  dcl-s hasM ind;
  dcl-s k int(10);
  pad = 0;
  s = ebc(code);
  k = %scan(';' : s);
  if k > 0;
    s = %subst(s : 1 : k - 1);
  endif;
  // remove "quoted", \x, _x, *x, [...] (but [h] [m] [s] = elapsed -> number)
  c = '';
  inQ = *off;
  i = 1;
  dow i <= %len(s);
    ch = %subst(s : i : 1);
    if inQ;
      if ch = '"';
        inQ = *off;
      endif;
      i += 1;
      iter;
    endif;
    select;
    when ch = '"';
      inQ = *on;
    when ch = '\' or ch = '_' or ch = '*';
      i += 1;
    when ch = '[';
      k = %scan(']' : s : i);
      if k = 0;
        k = %len(s);
      endif;
      if %check('hHmMsS' : %subst(s : i + 1 : k - i - 1)) = 0 and k > i + 1;
        return 'N';                     // elapsed time: a duration, keep the number
      endif;
      i = k;
    other;
      c += %xlate(UPPER : LOWER : ch);
    endsl;
    i += 1;
  enddo;
  if %len(c) >= 2 and %check('0' : c) = 0;
    pad = %len(c);
    return 'N';
  endif;
  if c = 'general' or c = '@';
    return 'N';
  endif;
  hasD = (%scan('y' : c) > 0 or %scan('d' : c) > 0 or %scan('e' : c) > 0);
  hasT = (%scan('h' : c) > 0 or %scan('s' : c) > 0);
  hasM = (%scan('m' : c) > 0);
  if %scan('e+' : c) > 0 or %scan('e-' : c) > 0;
    return 'N';                         // scientific
  endif;
  if hasD and hasT;
    return 'Z';
  elseif hasD;
    return 'D';
  elseif hasT;
    return 'T';
  elseif hasM and %check('m/-. ,' : c) = 0;
    return 'D';                         // mmm / mmmm: month
  endif;
  return 'N';
end-proc;

dcl-proc builtinKind;                   // built-in number format ids
  dcl-pi *n char(1);
    id int(10) value;
  end-pi;
  select;
  when id >= 14 and id <= 17;
    return 'D';
  when id >= 18 and id <= 21;
    return 'T';
  when id = 22;
    return 'Z';
  when id = 45 or id = 47;
    return 'T';
  when (id >= 27 and id <= 36) or (id >= 50 and id <= 58);
    return 'D';
  endsl;
  return 'N';
end-proc;

dcl-proc loadStyles;
  dcl-s b pointer;
  dcl-s n int(10);
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s x int(10);
  dcl-s i int(10);
  dcl-s id int(10);
  dcl-s sc ind;
  dcl-s v like(N_T);
  nXf = 0;
  nNf = 0;
  if gStyPart = '';
    return;
  endif;
  b = partGet(gStyPart : n);
  if b = *null;
    return;
  endif;
  p = findTag(b : n : 0 : u('numFmts'));
  x = -1;
  if p >= 0;
    x = findEnd(b : n : p : u('numFmts'));
  endif;
  dow p >= 0 and nNf < %elem(nfId);
    p = findTag(b : n : p + 1 : u('numFmt'));
    if p < 0 or (x > 0 and p > x);
      leave;
    endif;
    e = tagEnd(b : n : p : sc);
    nNf += 1;
    nfId(nNf) = toInt(ebc(attr(b : p : e : u('numFmtId'))));
    nfKind(nNf) = fmtKind(attr(b : p : e : u('formatCode')) : nfPad(nNf));
    p = e - 1;
  enddo;
  p = findTag(b : n : 0 : u('cellXfs'));
  if p >= 0;
    x = findEnd(b : n : p : u('cellXfs'));
    dow nXf < %elem(xfKind);
      p = findTag(b : n : p + 1 : u('xf'));
      if p < 0 or (x > 0 and p > x);
        leave;
      endif;
      e = tagEnd(b : n : p : sc);
      nXf += 1;
      id = toInt(ebc(attr(b : p : e : u('numFmtId'))));
      xfKind(nXf) = builtinKind(id);
      xfPad(nXf) = 0;
      if id >= 164 or xfKind(nXf) = 'N';
        for i = 1 to nNf;
          if nfId(i) = id;
            xfKind(nXf) = nfKind(i);
            xfPad(nXf) = nfPad(i);
            leave;
          endif;
        endfor;
      endif;
      p = e - 1;
    enddo;
  endif;
  dealloc(n) b;
end-proc;

// =====================================================================================
// sharedStrings.xml -> one UTF-8 buffer; texts of <t> outside <rPh> (phonetic) joined
// =====================================================================================
dcl-proc sstAdd;                        // append bytes to the string buffer
  dcl-pi *n;
    src pointer value;
    n int(10) value;
  end-pi;
  dcl-s nc int(10);
  if n <= 0;
    return;
  endif;
  if sstUse + n + 8 > sstCap;
    nc = sstCap * 2;
    if nc < sstUse + n + 65536;
      nc = sstUse + n + 65536;
    endif;
    pSstBuf = %realloc(pSstBuf : nc);
    sstCap = nc;
  endif;
  memcpy(pSstBuf + sstUse : src : n);
  sstUse += n;
end-proc;

dcl-proc loadSst;
  dcl-pi *n ind;
  end-pi;
  dcl-s b pointer;
  dcl-s n int(10);
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s x int(10);
  dcl-s q int(10);
  dcl-s t0 int(10);
  dcl-s cnt int(10);
  dcl-s cap int(10);
  dcl-s st int(10);
  dcl-s l int(10);
  dcl-s sc ind;
  dcl-s pp pointer;
  dcl-s v like(N_T);
  dcl-s c uns(3) based(pc);
  dcl-s pc pointer;
  dcl-s nch int(10);
  dcl-s nu int(10);
  dcl-s nl ind;
  dcl-s a int(10);
  dcl-s z int(10);
  nSst = 0;
  if gSstPart = '';
    return *on;
  endif;
  b = partGet(gSstPart : n);
  if b = *null;
    return *on;
  endif;
  p = findTag(b : n : 0 : u('sst'));
  cap = 0;
  if p >= 0;
    e = tagEnd(b : n : p : sc);
    cap = toInt(ebc(attr(b : p : e : u('uniqueCount'))));
  endif;
  if cap <= 0;                          // count the <si> elements
    q = 0;
    dow 1 = 1;
      q = findTag(b : n : q : kSi);
      if q < 0;
        leave;
      endif;
      cap += 1;
      q += 3;
    enddo;
  endif;
  if cap > %elem(sstOff);
    fail('too many shared strings: ' + %char(cap));
    dealloc(n) b;
    return *off;
  endif;
  if cap < 1;
    cap = 1;
  endif;
  pSstOff = %alloc(cap * 4 + 16);
  pSstLen = %alloc(cap * 4 + 16);
  pSstChr = %alloc(cap * 4 + 16);
  pSstU16 = %alloc(cap * 4 + 16);
  pSstNL = %alloc(cap + 16);
  sstCap = n + 65536;
  pSstBuf = %alloc(sstCap);
  sstUse = 0;
  p = 0;
  dow nSst < cap;
    p = findTag(b : n : p : kSi);
    if p < 0;
      leave;
    endif;
    e = tagEnd(b : n : p : sc);
    nSst += 1;
    sstOff(nSst) = sstUse;
    if not sc;
      x = findEnd(b : n : e : kSi);
      if x < 0;
        x = n;
      endif;
      q = e;
      dow q < x;                        // <t> parts, <rPh> skipped
        pp = memchr(b + q : 60 : x - q);
        if pp = *null;
          leave;
        endif;
        q = pp - b;
        pc = pp + 1;
        if c = 47;
          q += 1;
          iter;
        endif;
        l = locName(b : n : q : st);
        if l = 3 and isName(b : n : q : kRph);
          q = findEnd(b : n : q : kRph);
          if q < 0;
            q = x;
          endif;
          q += 1;
          iter;
        endif;
        if l = 1 and isName(b : n : q : kT);
          t0 = tagEnd(b : n : q : sc);
          if not sc and t0 > 0;
            pp = memchr(b + t0 : 60 : x - t0);
            if pp = *null;
              q = x;
            else;
              q = pp - b;
            endif;
            l = q - t0;
            if l > 0;
              if sstUse + l + 16 > sstCap;
                pSstBuf = %realloc(pSstBuf : sstCap * 2 + l + 65536);
                sstCap = sstCap * 2 + l + 65536;
              endif;
              sstUse += xdecode(b + t0 : l : pSstBuf + sstUse : l + 8);
            endif;
            iter;
          endif;
        endif;
        q += 1;
      enddo;
      p = x;
    else;
      p = e;
    endif;
    sstLen(nSst) = sstUse - sstOff(nSst);
    // trimmed view + metrics
    a = sstOff(nSst);
    z = a + sstLen(nSst);
    if gTrim;
      dow a < z;
        pc = pSstBuf + a;
        if c <> 32 and c <> 9 and c <> 10 and c <> 13;
          leave;
        endif;
        a += 1;
      enddo;
      dow z > a;
        pc = pSstBuf + z - 1;
        if c <> 32 and c <> 9 and c <> 10 and c <> 13;
          leave;
        endif;
        z -= 1;
      enddo;
    endif;
    sstOff(nSst) = a;
    sstLen(nSst) = z - a;
    utf8Info(pSstBuf + a : z - a : nch : nu : nl);
    sstChr(nSst) = nch;
    sstU16(nSst) = nu;
    if nl;
      sstNL(nSst) = 'N';
    else;
      sstNL(nSst) = 'L';
    endif;
  enddo;
  dealloc(n) b;
  return *on;
end-proc;

// =====================================================================================
// sheet selection, RANGE (cells / Excel table / defined name), sheet XML
// =====================================================================================

// "A1:F20" "$A$1" "B:F" "3:20" -> rR1 rC1 rR2 rC2; *off when not a cell range
dcl-proc parseRef;
  dcl-pi *n ind;
    s0 varchar(256) const;
  end-pi;
  dcl-s s varchar(256);
  dcl-s a varchar(64);
  dcl-s b varchar(64);
  dcl-s k int(10);
  dcl-s c int(10);
  dcl-s r int(20);
  s = %trim(s0);
  k = %scan('!' : s);
  if k > 0;
    s = %subst(s : k + 1);
  endif;
  k = %scan(':' : s);
  if k > 0;
    a = %subst(s : 1 : k - 1);
    b = %subst(s : k + 1);
  else;
    a = s;
    b = '';
  endif;
  if not cellRef(a : c : r);
    return *off;
  endif;
  if c > 16384 or (b = '' and r = 0);   // 'Orders' is a name, not column ORDERS
    return *off;
  endif;
  rC1 = c;
  rR1 = r;
  rC2 = 0;
  rR2 = 0;
  if b <> '';
    if not cellRef(b : c : r) or c > 16384;
      return *off;
    endif;
    rC2 = c;
    rR2 = r;
  elseif c > 0 and r > 0;
    rC2 = 0;                            // a single cell: start of the area
  endif;
  return *on;
end-proc;

dcl-proc sheetPrefix;                   // 'My sheet'!A1 -> My sheet
  dcl-pi *n like(N_T);
    s like(N_T) const;
  end-pi;
  dcl-s k int(10);
  dcl-s r like(N_T);
  k = %scan(u('!') : s);
  if k <= 1;
    return '';
  endif;
  r = %subst(s : 1 : k - 1);
  if %len(r) >= 2 and %subst(r : 1 : 1) = u('''');
    r = %subst(r : 2 : %len(r) - 2);
    r = %scanrpl(u('''''') : u('''') : r);
  endif;
  return r;
end-proc;

dcl-proc sheetByName;
  dcl-pi *n int(10);
    nm like(N_T) const;
  end-pi;
  dcl-s i int(10);
  dcl-s w varchar(512);
  w = %xlate(LOWER : UPPER : ebc(nm));
  for i = 1 to nSh;
    if %xlate(LOWER : UPPER : ebc(sh(i).name)) = w;
      return i;
    endif;
  endfor;
  return 0;
end-proc;

dcl-proc isWorksheet;
  dcl-pi *n ind;
    i int(10) value;
  end-pi;
  return sh(i).part <> '' and %scan(u('worksheets/') : sh(i).part) > 0;
end-proc;

// Excel table named nm anywhere in the workbook -> sheet, ref, header / totals rows
dcl-proc findTable;
  dcl-pi *n int(10);                    // sheet index, 0 = not found
    nm varchar(256) const;
    first ind value;                    // *on: the first table of sheet 'only'
    only int(10) value;
  end-pi;
  dcl-s i int(10);
  dcl-s j int(10);
  dcl-s nr int(10);
  dcl-s b pointer;
  dcl-s n int(10);
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s sc ind;
  dcl-s want varchar(256);
  dcl-s tg like(N_T) dim(50);
  dcl-s nt int(10);
  dcl-s v like(N_T);
  want = %xlate(LOWER : UPPER : nm);
  for i = 1 to nSh;
    if not isWorksheet(i) or (only > 0 and i <> only);
      iter;
    endif;
    nr = relsLoad(sh(i).part);
    nt = 0;
    for j = 1 to nr;
      if %len(rl(j).typ) > 6 and nt < 50;
        if %subst(rl(j).typ : %len(rl(j).typ) - 5) = u('/table');
          nt += 1;
          tg(nt) = rl(j).tgt;
        endif;
      endif;
    endfor;
    for j = 1 to nt;
      b = partGet(tg(j) : n);
      if b = *null;
        iter;
      endif;
      p = findTag(b : n : 0 : u('table'));
      if p >= 0;
        e = tagEnd(b : n : p : sc);
        if first or %xlate(LOWER : UPPER : ebc(attr(b : p : e : u('displayName')))) = want
           or %xlate(LOWER : UPPER : ebc(attr(b : p : e : u('name')))) = want;
          v = attr(b : p : e : u('ref'));
          rHdr = 1;
          rTot = 0;
          if attr(b : p : e : u('headerRowCount')) <> '';
            rHdr = toInt(ebc(attr(b : p : e : u('headerRowCount'))));
          endif;
          if attr(b : p : e : u('totalsRowCount')) <> '';
            rTot = toInt(ebc(attr(b : p : e : u('totalsRowCount'))));
          endif;
          rHow = 'Excel table ' + ebc(attr(b : p : e : u('displayName')))
               + ' (' + ebc(v) + ')';
          dealloc(n) b;
          if parseRef(ebc(v));
            return i;
          endif;
          return 0;
        endif;
      endif;
      dealloc(n) b;
    endfor;
  endfor;
  return 0;
end-proc;

// SHEET + RANGE -> gShIdx and rR1/rC1/rR2/rC2
dcl-proc pickSheet;
  dcl-pi *n ind;
  end-pi;
  dcl-s i int(10);
  dcl-s k int(10);
  dcl-s lst varchar(512);
  dcl-s pfx like(N_T);
  dcl-s want varchar(256);
  gShIdx = 0;
  rR1 = 0;
  rR2 = 0;
  rC1 = 0;
  rC2 = 0;
  rHdr = -1;
  rTot = 0;
  rHow = '';
  if gSheetSel <> '';
    gShIdx = sheetByName(gSheetSel);
    if gShIdx = 0 and onlyDigits(ebc(gSheetSel));
      k = toInt(ebc(gSheetSel));
      if k >= 1 and k <= nSh;
        gShIdx = k;
      endif;
    endif;
    if gShIdx = 0;
      lst = '';
      for i = 1 to nSh;
        if lst <> '';
          lst += ', ';
        endif;
        lst += ebc(sh(i).name);
      endfor;
      fail('sheet ' + ebc(gSheetSel) + ' not found (sheets: ' + lst + ')');
      return *off;
    endif;
  endif;

  if gRange <> '';
    pfx = sheetPrefix(u(gRange));
    if parseRef(gRange);
      rHow = 'RANGE ' + gRange;
      if pfx <> '';
        gShIdx = sheetByName(pfx);
        if gShIdx = 0;
          fail('RANGE ' + gRange + ': sheet ' + ebc(pfx) + ' not found');
          return *off;
        endif;
      endif;
    else;
      k = findTable(gRange : *off : gShIdx);
      if k > 0;
        gShIdx = k;
      else;
        want = %xlate(LOWER : UPPER : gRange);
        for i = 1 to nDn;
          if %xlate(LOWER : UPPER : ebc(dn(i).name)) = want;
            pfx = sheetPrefix(dn(i).ref);
            if pfx <> '' and parseRef(ebc(dn(i).ref));
              k = sheetByName(pfx);
            endif;
            if k > 0;
              gShIdx = k;
              rHow = 'defined name ' + gRange + ' (' + ebc(dn(i).ref) + ')';
              leave;
            endif;
          endif;
        endfor;
        if k = 0;
          fail('RANGE ' + gRange + ' is not a cell range, Excel table or workbook name');
          return *off;
        endif;
      endif;
    endif;
  endif;

  if gShIdx = 0;                        // *FIRST: first visible worksheet
    for i = 1 to nSh;
      if isWorksheet(i) and sh(i).state = 'V';
        gShIdx = i;
        leave;
      endif;
    endfor;
    if gShIdx = 0;
      for i = 1 to nSh;
        if isWorksheet(i);
          gShIdx = i;
          leave;
        endif;
      endfor;
    endif;
  endif;
  if gShIdx = 0 or not isWorksheet(gShIdx);
    fail('sheet ' + ebc(sh(gShIdx).name) + ' is not a worksheet');
    return *off;
  endif;
  return *on;
end-proc;

dcl-proc loadSheet;
  dcl-pi *n ind;
  end-pi;
  dcl-s p int(10);
  dcl-s e int(10);
  dcl-s q int(10);
  dcl-s sc ind;
  dcl-s c uns(3) based(pc);
  dcl-s pc pointer;
  pSh = partGet(sh(gShIdx).part : nShLen);
  if pSh = *null;
    fail('sheet ' + ebc(sh(gShIdx).name) + ': ' + xlErr());
    return *off;
  endif;
  p = findTag(pSh : nShLen : 0 : u('sheetData'));
  gSdFrom = 0;
  gSdTo = 0;
  if p < 0;
    return *on;                         // no data
  endif;
  e = tagEnd(pSh : nShLen : p : sc);
  if sc;
    gSdFrom = e;
    gSdTo = e;
  else;
    gSdFrom = e;
    q = nShLen - 12;                    // </sheetData> is near the end: search back
    dow q > e;
      pc = pSh + q;
      if c = 60;
        pc = pSh + q + 1;
        if c = 47 and isName(pSh : nShLen : q + 1 : u('sheetData'));
          leave;
        endif;
      endif;
      q -= 1;
    enddo;
    gSdTo = q;
    if q <= e;
      gSdTo = nShLen;
    endif;
  endif;
  gAfRef = '';
  p = findTag(pSh : nShLen : gSdTo : u('autoFilter'));
  if p >= 0;
    e = tagEnd(pSh : nShLen : p : sc);
    gAfRef = attr(pSh : p : e : u('ref'));
  endif;
  return *on;
end-proc;

// =====================================================================================
// the sheet scanner: rows and cells of <sheetData>, then rowEnd() per row
// =====================================================================================
dcl-proc rowTxt;                        // decode bytes into the row text buffer
  dcl-pi *n int(10);                    // offset of the text
    src pointer value;
    n int(10) value;
  end-pi;
  dcl-s o int(10);
  if rowTxtUse + n + 16 > rowTxtCap;
    rowTxtCap = (rowTxtUse + n) * 2 + 65536;
    pRowTxt = %realloc(pRowTxt : rowTxtCap);
  endif;
  o = rowTxtUse;
  rowTxtUse += xdecode(src : n : pRowTxt + o : n + 8);
  return o;
end-proc;

dcl-proc scanSheet;
  dcl-s b pointer;
  dcl-s n int(10);
  dcl-s p int(10);
  dcl-s q int(10);
  dcl-s e int(10);
  dcl-s st int(10);
  dcl-s l int(10);
  dcl-s pp pointer;
  dcl-s c uns(3) based(pc);
  dcl-s pc pointer;
  dcl-s sc ind;
  dcl-s inRow ind;
  dcl-s lastRow int(20);
  dcl-s lastCol int(10);
  dcl-s col int(10);
  dcl-s sty int(10);
  dcl-s tk char(1);
  dcl-s a int(10);
  dcl-s an int(10);
  dcl-s v0 int(10);
  dcl-s quo uns(3);
  dcl-s vP pointer;
  dcl-s vL int(10);
  dcl-s tO int(10);
  dcl-s tL int(10);
  dcl-s x int(10);
  dcl-s t0 int(10);
  dcl-s ok ind;
  dcl-s v like(N_T);
  dcl-s cr int(10);
  dcl-s rr int(20);

  b = pSh;
  n = gSdTo;
  p = gSdFrom;
  gStop = *off;
  inRow = *off;
  lastRow = 0;
  dow p < n and not gStop;
    pp = memchr(b + p : 60 : n - p);
    if pp = *null;
      leave;
    endif;
    p = pp - b;
    pc = pp + 1;
    if c = 47;                          // end tag
      if inRow and isName(b : n : p + 1 : kRow);
        rowEnd();
        inRow = *off;
      endif;
      p += 2;
      iter;
    endif;
    l = locName(b : n : p : st);
    pc = b + st;
    if l = 3 and isName(b : n : p : kRow);
      if inRow;
        rowEnd();
        if gStop;
          leave;
        endif;
      endif;
      e = tagEnd(b : n : p : sc);
      v = attr(b : p : e : kR);
      if v <> '';
        gRowNo = toIntU(v);
      else;
        gRowNo = lastRow + 1;
      endif;
      lastRow = gRowNo;
      v = attr(b : p : e : kHidden);
      gRowHid = (v = kOne or v = kTrue);
      gGen += 1;
      gRowMax = 0;
      rowTxtUse = 0;
      lastCol = 0;
      inRow = *on;
      if sc;
        rowEnd();
        inRow = *off;
      endif;
      p = e;
      iter;
    endif;
    if not (l = 1 and c = 99) or not inRow;   // not a <c>
      p += 1;
      iter;
    endif;

    // ---- a cell: attributes r s t
    e = tagEnd(b : n : p : sc);
    if e < 0;
      leave;
    endif;
    col = 0;
    sty = 0;
    tk = 'N';
    a = st + 1;
    dow a < e - 1;
      pc = b + a;
      if c = 32 or c = 9 or c = 10 or c = 13;
        a += 1;
        iter;
      endif;
      if c = 62 or c = 47;
        leave;
      endif;
      an = a;                           // attribute name (local part)
      dow a < e;
        pc = b + a;
        if c = 61 or c = 32;
          leave;
        endif;
        if c = 58;
          an = a + 1;
        endif;
        a += 1;
      enddo;
      x = a - an;
      dow a < e;
        pc = b + a;
        if c = 34 or c = 39;
          leave;
        endif;
        a += 1;
      enddo;
      quo = c;
      a += 1;
      v0 = a;
      dow a < e;
        pc = b + a;
        if c = quo;
          leave;
        endif;
        a += 1;
      enddo;
      if x = 1;
        pc = b + an;
        select;
        when c = 114;                   // r: letters + digits
          col = 0;
          for t0 = v0 to a - 1;
            pc = b + t0;
            if c >= 65 and c <= 90;
              col = col * 26 + c - 64;
            elseif c >= 97 and c <= 122;
              col = col * 26 + c - 96;
            else;
              leave;
            endif;
          endfor;
        when c = 115;                   // s: style index
          sty = 0;
          for t0 = v0 to a - 1;
            pc = b + t0;
            sty = sty * 10 + c - 48;
          endfor;
          sty += 1;
        when c = 116;                   // t: type
          pc = b + v0;
          if a - v0 = 1;
            select;
            when c = 115;
              tk = 's';
            when c = 98;
              tk = 'B';
            when c = 101;
              tk = 'X';
            when c = 100;
              tk = 'd';
            when c = 110;
              tk = 'N';
            endsl;
          elseif a - v0 = 3 and c = 115;          // str
            tk = 'S';
          elseif a - v0 = 9 and c = 105;          // inlineStr
            tk = 'I';
          endif;
        endsl;
      endif;
      a += 1;
    enddo;
    if col = 0;
      col = lastCol + 1;
    endif;
    lastCol = col;

    // ---- content: <v>, <is>, <f>
    vP = *null;
    vL = -1;
    tO = 0;
    tL = -1;
    if not sc;
      q = e;
      dow q < n;
        pp = memchr(b + q : 60 : n - q);
        if pp = *null;
          q = n;
          leave;
        endif;
        q = pp - b;
        pc = pp + 1;
        if c = 47;                      // </...>
          if isName(b : n : q + 1 : kC);
            pp = memchr(b + q : 62 : n - q);
            q = pp - b + 1;
            leave;
          endif;
          q += 2;
          iter;
        endif;
        x = locName(b : n : q : t0);
        pc = b + t0;
        if x = 1 and c = 118;           // <v>
          t0 = tagEnd(b : n : q : ok);
          if ok;
            q = t0;
            iter;
          endif;
          pp = memchr(b + t0 : 60 : n - t0);
          if pp = *null;
            q = n;
            leave;
          endif;
          vP = b + t0;
          vL = pp - vP;
          q = pp - b;
          iter;
        endif;
        if x = 2 and c = 105;           // <is>: <t> parts, <rPh> skipped
          tO = rowTxtUse;
          tL = 0;
          q = tagEnd(b : n : q : ok);
          dow q < n;
            pp = memchr(b + q : 60 : n - q);
            if pp = *null;
              q = n;
              leave;
            endif;
            q = pp - b;
            pc = pp + 1;
            if c = 47 and isName(b : n : q + 1 : kIs);
              leave;
            endif;
            if c <> 47 and isName(b : n : q : kRph);
              q = findEnd(b : n : q : kRph);
              if q < 0;
                q = n;
              endif;
              q += 1;
              iter;
            endif;
            if c <> 47 and isName(b : n : q : kT);
              t0 = tagEnd(b : n : q : ok);
              if not ok;
                pp = memchr(b + t0 : 60 : n - t0);
                if pp <> *null;
                  x = rowTxt(b + t0 : pp - b - t0);
                  q = pp - b;
                  iter;
                endif;
              endif;
            endif;
            q += 1;
          enddo;
          tL = rowTxtUse - tO;
          q += 1;
          iter;
        endif;
        if x = 1 and c = 102;           // <f>: skip the formula
          t0 = tagEnd(b : n : q : ok);
          if ok;
            q = t0;
          else;
            q = findEnd(b : n : t0 : kF);
            if q < 0;
              q = n;
            endif;
            q += 2;
          endif;
          iter;
        endif;
        q += 1;
      enddo;
      p = q;
    else;
      p = e;
    endif;

    // ---- store
    if col > MAXC;
      gBeyond += 1;
      iter;
    endif;
    select;
    when tk = 'I';
      if tL < 0;
        iter;
      endif;
      rcKind(col) = 'S';
      rcTo(col) = tO;
      rcTl(col) = tL;
    when vL < 0;
      iter;                             // no value: empty cell
    when tk = 'S' or tk = 'd' or tk = 'X';
      rcKind(col) = tk;
      rcTo(col) = rowTxt(vP : vL);
      rcTl(col) = rowTxtUse - rcTo(col);
    when tk = 's';
      rcKind(col) = 's';
      rcSst(col) = 0;
      for t0 = 0 to vL - 1;
        pc = vP + t0;
        if c >= 48 and c <= 57;
          rcSst(col) = rcSst(col) * 10 + c - 48;
        endif;
      endfor;
      if rcSst(col) >= nSst;
        iter;                           // bad index: empty
      endif;
    other;
      rcKind(col) = tk;                 // N B
      rcP(col) = vP;
      rcL(col) = vL;
    endsl;
    rcSty(col) = sty;
    rcGen(col) = gGen;
    if col > gRowMax;
      gRowMax = col;
    endif;
  enddo;
  if inRow and not gStop;
    rowEnd();
  endif;
end-proc;

dcl-proc rowEnd;
  select;
  when gPass = 0;
    sampleRow();
  when gPass = 1;
    statRow();
  other;
    loadRow();
  endsl;
end-proc;

// kind of the cell in column c of the current row: E empty, N number, S text,
// B boolean, X error, d ISO date text (s and str count as S)
dcl-proc ck;
  dcl-pi *n char(1);
    c int(10) value;
  end-pi;
  if c < 1 or c > MAXC or rcGen(c) <> gGen;
    return 'E';
  endif;
  if rcKind(c) = 's';
    return 'S';
  endif;
  return rcKind(c);
end-proc;

// text of a text cell (trimmed when TRIM(*YES)): pointer + length
dcl-proc cTxt;
  dcl-pi *n pointer;
    c int(10) value;
    len int(10);
  end-pi;
  dcl-s p pointer;
  dcl-s ch uns(3) based(pc);
  dcl-s pc pointer;
  if rcKind(c) = 's';
    len = sstLen(rcSst(c) + 1);
    return pSstBuf + sstOff(rcSst(c) + 1);
  endif;
  p = pRowTxt + rcTo(c);
  len = rcTl(c);
  if gTrim;
    dow len > 0;
      pc = p;
      if ch <> 32 and ch <> 9 and ch <> 10 and ch <> 13;
        leave;
      endif;
      p += 1;
      len -= 1;
    enddo;
    dow len > 0;
      pc = p + len - 1;
      if ch <> 32 and ch <> 9 and ch <> 10 and ch <> 13;
        leave;
      endif;
      len -= 1;
    enddo;
  endif;
  return p;
end-proc;

dcl-proc cText;                         // the text as a value (max 32740 bytes)
  dcl-pi *n like(U_T) rtnparm;
    c int(10) value;
  end-pi;
  dcl-s p pointer;
  dcl-s l int(10);
  dcl-s r like(U_T);
  p = cTxt(c : l);
  if l > 32740;
    l = 32740;
  endif;
  if l > 0;
    memcpy(%addr(gDec) : p : l);
    r = %subst(gDec : 1 : l);
  endif;
  return r;
end-proc;

dcl-proc fmtOf;                         // number format kind of the cell's style
  dcl-pi *n char(1);
    c int(10) value;
  end-pi;
  if rcSty(c) >= 1 and rcSty(c) <= nXf;
    return xfKind(rcSty(c));
  endif;
  return 'N';
end-proc;

dcl-proc padOf;
  dcl-pi *n int(10);
    c int(10) value;
  end-pi;
  if rcSty(c) >= 1 and rcSty(c) <= nXf;
    return xfPad(rcSty(c));
  endif;
  return 0;
end-proc;

// =====================================================================================
// pass 0: sample the first non-empty rows
// =====================================================================================
dcl-proc sampleRow;
  dcl-s c int(10);
  dcl-s cnt int(10);
  dcl-s c1 int(10);
  dcl-s c2 int(10);
  dcl-s k char(1);
  dcl-s ix int(10);
  dcl-s p pointer;
  dcl-s l int(10);
  dcl-s lo int(10);
  dcl-s hi int(10);
  if gRowNo < gSmFrom;
    return;
  endif;
  if rR2 > 0 and gRowNo > rR2;
    gStop = *on;
    return;
  endif;
  lo = 1;
  hi = gRowMax;
  if rC1 > 0;
    lo = rC1;
  endif;
  if rC2 > 0 and rC2 < hi;
    hi = rC2;
  endif;
  cnt = 0;
  c1 = 0;
  c2 = 0;
  for c = lo to hi;
    if ck(c) <> 'E';
      cnt += 1;
      if c1 = 0;
        c1 = c;
      endif;
      c2 = c;
    endif;
  endfor;
  if cnt = 0;
    return;
  endif;
  nSm += 1;
  sm(nSm).row = gRowNo;
  sm(nSm).cnt = cnt;
  sm(nSm).c1 = c1;
  sm(nSm).c2 = c2;
  for c = c1 to c2;
    ix = (nSm - 1) * MAXC + c;
    k = ck(c);
    smKind(ix) = k;
    smLen(ix) = 0;
    if k = 'E';
      iter;
    endif;
    if k = 'N' or k = 'B';
      p = rcP(c);
      l = rcL(c);
    else;
      p = cTxt(c : l);
    endif;
    if l > 512;
      l = 512;
    endif;
    pSmBuf = %realloc(pSmBuf : smUse + l + 16);
    if l > 0;
      memcpy(pSmBuf + smUse : p : l);
    endif;
    smOff(ix) = smUse;
    smLen(ix) = l;
    smUse += l;
  endfor;
  if nSm >= MAXS;
    gStop = *on;
  endif;
end-proc;

dcl-proc smText;                        // sampled text of (sample k, column c)
  dcl-pi *n like(N_T);
    k int(10) value;
    c int(10) value;
  end-pi;
  dcl-s ix int(10);
  dcl-s r like(N_T);
  if c < sm(k).c1 or c > sm(k).c2;
    return '';
  endif;
  ix = (k - 1) * MAXC + c;
  if smLen(ix) > 0;
    memcpy(%addr(gDec) : pSmBuf + smOff(ix) : smLen(ix));
    r = %subst(gDec : 1 : smLen(ix));
  endif;
  return r;
end-proc;

dcl-proc smK;
  dcl-pi *n char(1);
    k int(10) value;
    c int(10) value;
  end-pi;
  if c < sm(k).c1 or c > sm(k).c2;
    return 'E';
  endif;
  return smKind((k - 1) * MAXC + c);
end-proc;

// =====================================================================================
// numbers: exact decimal digits, value = d * 10**ex (no binary floating point)
// =====================================================================================
dcl-proc numNorm;
  dcl-pi *n;
    x likeds(num_t);
  end-pi;
  dcl-s i int(10);
  i = 1;                                // leading zeros
  dow i <= x.nd and %subst(x.d : i : 1) = '0';
    i += 1;
  enddo;
  if i > 1;
    if i > x.nd;
      x.nd = 0;
    else;
      x.d = %subst(x.d : i : x.nd - i + 1);
      x.nd -= i - 1;
    endif;
  endif;
  dow x.nd > 0 and %subst(x.d : x.nd : 1) = '0';   // trailing zeros
    x.nd -= 1;
    x.ex += 1;
  enddo;
  if x.nd = 0;
    x.neg = *off;
    x.ex = 0;
  endif;
end-proc;

dcl-proc numInc;                        // add 1 to the last kept digit
  dcl-pi *n;
    x likeds(num_t);
  end-pi;
  dcl-s i int(10);
  dcl-s k int(10);
  i = x.nd;
  dow i >= 1;
    k = %scan(%subst(x.d : i : 1) : '0123456789') - 1;
    if k < 9;
      %subst(x.d : i : 1) = %subst('0123456789' : k + 2 : 1);
      return;
    endif;
    %subst(x.d : i : 1) = '0';
    i -= 1;
  enddo;
  x.d = '1' + %subst(x.d : 1 : 39);     // 999 -> 1000
  x.nd += 1;
end-proc;

dcl-proc numKeep;                       // keep nk significant digits, half up
  dcl-pi *n;
    x likeds(num_t);
    nk int(10) value;
  end-pi;
  dcl-s up ind;
  if x.nd <= nk;
    return;
  endif;
  up = (%subst(x.d : nk + 1 : 1) >= '5');
  x.ex += x.nd - nk;
  x.nd = nk;
  if up;
    numInc(x);
  endif;
  numNorm(x);
end-proc;

dcl-proc numScale;                      // round to s decimals, half up
  dcl-pi *n;
    x likeds(num_t);
    s int(10) value;
  end-pi;
  dcl-s k int(10);
  if x.nd = 0 or -x.ex <= s;
    return;
  endif;
  k = -x.ex - s;                        // digits to drop
  if k > x.nd;
    x.nd = 0;
    numNorm(x);
    return;
  endif;
  if k = x.nd;
    if %subst(x.d : 1 : 1) >= '5';
      x.d = '1';
      x.nd = 1;
      x.ex = -s;
    else;
      x.nd = 0;
    endif;
    numNorm(x);
    return;
  endif;
  numKeep(x : x.nd - k);
end-proc;

dcl-proc numIntDig;
  dcl-pi *n int(10);
    x likeds(num_t) const;
  end-pi;
  if x.nd + x.ex > 0;
    return x.nd + x.ex;
  endif;
  return 0;
end-proc;

dcl-proc numSc;
  dcl-pi *n int(10);
    x likeds(num_t) const;
  end-pi;
  if x.ex < 0;
    return -x.ex;
  endif;
  return 0;
end-proc;

dcl-proc numSci;                        // 12345E-2 (for float columns / huge values)
  dcl-pi *n varchar(400);
    x likeds(num_t) const;
  end-pi;
  dcl-s r varchar(400);
  if x.nd = 0;
    return '0';
  endif;
  r = %subst(x.d : 1 : x.nd) + 'E' + %char(x.ex);
  if x.neg;
    r = '-' + r;
  endif;
  return r;
end-proc;

dcl-proc numPlain;                      // 1234.5 / -0.005 / 120000
  dcl-pi *n varchar(400);
    x likeds(num_t) const;
  end-pi;
  dcl-s r varchar(400);
  dcl-s k int(10);
  if x.nd = 0;
    return '0';
  endif;
  if x.ex >= 0;
    if x.ex > 300;
      return numSci(x);
    endif;
    r = %subst(x.d : 1 : x.nd);
    for k = 1 to x.ex;
      r += '0';
    endfor;
  else;
    k = x.nd + x.ex;                    // digits before the point
    if -x.ex > 300;
      return numSci(x);
    endif;
    if k > 0;
      r = %subst(x.d : 1 : k) + '.' + %subst(x.d : k + 1 : x.nd - k);
    else;
      r = '0.';
      dow k < 0;
        r += '0';
        k += 1;
      enddo;
      r += %subst(x.d : 1 : x.nd);
    endif;
  endif;
  if x.neg;
    r = '-' + r;
  endif;
  return r;
end-proc;

dcl-proc numPlainLen;                   // %len(numPlain(x)) without building it
  dcl-pi *n int(10);
    x likeds(num_t) const;
  end-pi;
  dcl-s k int(10);
  dcl-s l int(10);
  if x.nd = 0;
    return 1;
  endif;
  if x.ex > 300 or x.ex < -300;
    return %len(numPlain(x));
  endif;
  if x.ex >= 0;
    l = x.nd + x.ex;
  else;
    k = x.nd + x.ex;
    if k > 0;
      l = x.nd + 1;
    else;
      l = 2 - k + x.nd;
    endif;
  endif;
  if x.neg;
    l += 1;
  endif;
  return l;
end-proc;

dcl-proc numDec;                        // value as packed(31:10) (for date serials)
  dcl-pi *n packed(31 : 10);
    x likeds(num_t) const;
  end-pi;
  dcl-ds y likeds(num_t);
  if numIntDig(x) > 20;
    return 0;
  endif;
  y = x;
  numScale(y : 10);
  return %dec(numPlain(y) : 31 : 10);
end-proc;

// strict number of a <v> element (ASCII): [-]d[.d][E[+-]d]
dcl-proc numRaw;
  dcl-pi *n ind;
    p pointer value;
    n int(10) value;
    x likeds(num_t);
  end-pi;
  dcl-s c uns(3) based(pc);
  dcl-s pc pointer;
  dcl-s i int(10);
  dcl-s dot ind;
  dcl-s any ind;
  dcl-s ev int(10);
  dcl-s eneg ind;
  dcl-s dg uns(3) based(pd);
  dcl-s pd pointer;
  x.neg = *off;
  x.nd = 0;
  x.ex = 0;
  dot = *off;
  any = *off;
  i = 0;
  pc = p;
  if n > 0 and (c = 45 or c = 43);
    x.neg = (c = 45);
    i = 1;
  endif;
  dow i < n;
    pc = p + i;
    if c >= 48 and c <= 57;
      any = *on;
      if x.nd = 0 and c = 48;
        if dot;
          x.ex -= 1;
        endif;
      elseif x.nd < 40;
        x.nd += 1;
        pd = %addr(x.d) + x.nd - 1;     // job CCSID digit X'F0'..X'F9'
        dg = 192 + c;
        if dot;
          x.ex -= 1;
        endif;
      elseif not dot;
        x.ex += 1;
      endif;
    elseif c = 46 and not dot;
      dot = *on;
    elseif (c = 69 or c = 101) and any;
      ev = 0;
      eneg = *off;
      i += 1;
      pc = p + i;
      if i < n and (c = 45 or c = 43);
        eneg = (c = 45);
        i += 1;
      endif;
      if i >= n;
        return *off;
      endif;
      dow i < n;
        pc = p + i;
        if c < 48 or c > 57;
          return *off;
        endif;
        ev = ev * 10 + c - 48;
        if ev > 9999;
          return *off;
        endif;
        i += 1;
      enddo;
      if eneg;
        x.ex -= ev;
      else;
        x.ex += ev;
      endif;
      leave;
    else;
      return *off;
    endif;
    i += 1;
  enddo;
  if not any;
    return *off;
  endif;
  numNorm(x);
  return *on;
end-proc;

dcl-proc numParseE;                     // EBCDIC text -> number (via ASCII bytes)
  dcl-pi *n ind;
    s varchar(200) const;
    x likeds(num_t);
  end-pi;
  dcl-s a varchar(200) ccsid(*utf8);
  a = s;
  return numRaw(%addr(a : *data) : %len(a) : x);
end-proc;

// lenient number in a text: blanks, thousands separators, decimal comma, %,
// currency signs, (123) and 123- negatives
dcl-proc numText;
  dcl-pi *n ind;
    t0 like(U_T) const;
    x likeds(num_t);
  end-pi;
  dcl-s s varchar(200);
  dcl-s w varchar(200);
  dcl-s i int(10);
  dcl-s ch char(1);
  dcl-s neg ind;
  dcl-s pct ind;
  dcl-s nDot int(10);
  dcl-s nCom int(10);
  dcl-s lDot int(10);
  dcl-s lCom int(10);
  dcl-s decCh char(1);
  dcl-s hasE ind;
  dcl-s r varchar(200);
  dcl-s cur varchar(8);
  if %len(t0) = 0 or %len(t0) > 200;
    return *off;
  endif;
  cur = ebc(u'20AC00A30024');           // euro pound dollar in the job CCSID
  s = ebc(t0);
  s = %trim(%xlate(x'41' : ' ' : s));   // NBSP -> blank
  neg = *off;
  pct = *off;
  if %len(s) >= 2 and %subst(s : 1 : 1) = '(' and %subst(s : %len(s) : 1) = ')';
    neg = *on;
    s = %trim(%subst(s : 2 : %len(s) - 2));
  endif;
  dow %len(s) > 0 and %scan(%subst(s : 1 : 1) : cur) > 0;
    s = %trim(%subst(s : 2));
  enddo;
  dow %len(s) > 0 and %scan(%subst(s : %len(s) : 1) : cur + '%') > 0;
    if %subst(s : %len(s) : 1) = '%';
      pct = *on;
    endif;
    s = %trimr(%subst(s : 1 : %len(s) - 1));
  enddo;
  if %len(s) >= 2 and %subst(s : %len(s) : 1) = '-';
    neg = not neg;
    s = %subst(s : 1 : %len(s) - 1);
  endif;
  if %len(s) >= 1 and (%subst(s : 1 : 1) = '-' or %subst(s : 1 : 1) = '+');
    if %subst(s : 1 : 1) = '-';
      neg = not neg;
    endif;
    s = %trim(%subst(s : 2));
  endif;
  if %len(s) = 0 or %check('0123456789.,'' Ee+-' : s) > 0;
    return *off;
  endif;
  if %check('0123456789' : %subst(s : 1 : 1)) > 0 and %subst(s : 1 : 1) <> '.'
     and %subst(s : 1 : 1) <> ',';
    return *off;
  endif;
  // separators
  nDot = 0;
  nCom = 0;
  lDot = 0;
  lCom = 0;
  hasE = *off;
  for i = 1 to %len(s);
    ch = %subst(s : i : 1);
    if ch = '.';
      nDot += 1;
      lDot = i;
    elseif ch = ',';
      nCom += 1;
      lCom = i;
    elseif ch = 'E' or ch = 'e';
      hasE = *on;
    endif;
  endfor;
  decCh = ' ';
  if nDot > 0 and nCom > 0;
    if lDot > lCom;
      decCh = '.';
    else;
      decCh = ',';
    endif;
  elseif nCom = 1 or nDot = 1;
    decCh = ',';
    i = lCom;
    if nDot = 1;
      decCh = '.';
      i = lDot;
    endif;
    // 1.000 / 12,500: one separator + exactly 3 digits after a 1-3 digit integer part
    // that does not start with 0 -> read as the job's thousands separator
    if not hasE and decCh = grpCh() and i >= 2 and i <= 4 and %len(s) = i + 3
       and %subst(s : 1 : 1) <> '0';
      decCh = ' ';
    endif;
  endif;
  w = '';
  for i = 1 to %len(s);
    ch = %subst(s : i : 1);
    if ch = ' ' or ch = '''';
      iter;
    endif;
    if ch = '.' or ch = ',';
      if ch = decCh and (i = lDot or i = lCom);
        w += '.';
      endif;
      iter;
    endif;
    w += ch;
  endfor;
  if hasE and (nCom > 0);
    return *off;
  endif;
  r = w;
  if neg;
    r = '-' + w;
  endif;
  if not numParseE(r : x);
    return *off;
  endif;
  if pct;
    x.ex -= 2;
  endif;
  return *on;
end-proc;

// =====================================================================================
// dates: Excel serials (1900 / 1904 systems) and texts
// =====================================================================================

dcl-proc serialDT;                      // serial -> date + time (ms rounded)
  dcl-pi *n ind;
    v packed(31 : 10) value;
    r likeds(dt_t);
  end-pi;
  dcl-s days int(10);
  dcl-s ms int(20);
  dcl-s k int(10);
  if v < 0 or v > 2958466;
    return *off;
  endif;
  days = %int(v);
  ms = %inth((v - days) * 86400000);
  if ms >= 86400000;
    days += 1;
    ms -= 86400000;
  endif;
  k = %rem(days : 4096) + 1;            // date arithmetic is slow: cache by serial
  if dcKey(k) = days + 1;               // r.dd is only read through dtIsoR()
    r.iso = dcIso(k);
  else;
    monitor;
      if gDate1904;
        r.dd = d'1904-01-01' + %days(days);
      elseif days >= 61;
        r.dd = d'1899-12-30' + %days(days);
      else;
        r.dd = d'1899-12-31' + %days(days);
      endif;
    on-error;
      return *off;
    endmon;
    r.iso = %char(r.dd : *iso);
    dcKey(k) = days + 1;
    dcIso(k) = r.iso;
  endif;
  r.sec = %div(ms : 1000);
  r.us = %rem(ms : 1000) * 1000;
  r.hasD = (days > 0);
  r.hasT = (ms > 0);
  return *on;
end-proc;

dcl-proc dtIso;                         // 2026-10-31
  dcl-pi *n varchar(10);
    d date value;
  end-pi;
  return %char(d : *iso);
end-proc;

dcl-proc dtIsoR;                        // ISO text of r.dd (cached text when there)
  dcl-pi *n varchar(10);
    r likeds(dt_t) const;
  end-pi;
  if r.iso <> '';
    return r.iso;
  endif;
  return %char(r.dd : *iso);
end-proc;

dcl-proc tmText;                        // 13:45:10 (sep ':' or '.')
  dcl-pi *n varchar(8);
    sec int(10) value;
    sep char(1) const;
  end-pi;
  return %editc(%dec(%div(sec : 3600) : 2 : 0) : 'X') + sep
       + %editc(%dec(%rem(%div(sec : 60) : 60) : 2 : 0) : 'X') + sep
       + %editc(%dec(%rem(sec : 60) : 2 : 0) : 'X');
end-proc;

dcl-proc tsDb2;                         // 2026-10-31-13.45.10.123000
  dcl-pi *n varchar(26);
    r likeds(dt_t) const;
  end-pi;
  return dtIsoR(r) + '-' + tmText(r.sec : '.') + '.'
       + %editc(%dec(r.us : 6 : 0) : 'X');
end-proc;

dcl-proc tsText;                        // 2026-10-31 13:45:10 (for text columns)
  dcl-pi *n varchar(23);
    r likeds(dt_t) const;
  end-pi;
  dcl-s s varchar(23);
  s = dtIsoR(r) + ' ' + tmText(r.sec : ':');
  if r.us > 0;
    s += '.' + %subst(%editc(%dec(r.us : 6 : 0) : 'X') : 1 : 3);
  endif;
  return s;
end-proc;

dcl-proc mkDate;
  dcl-pi *n ind;
    y int(10) value;
    m int(10) value;
    d int(10) value;
    r date;
  end-pi;
  if y < 1 or y > 9999 or m < 1 or m > 12 or d < 1 or d > 31;
    return *off;
  endif;
  monitor;
    r = %date(%editc(%dec(y : 4 : 0) : 'X') + '-' + %editc(%dec(m : 2 : 0) : 'X') + '-'
            + %editc(%dec(d : 2 : 0) : 'X') : *iso);
  on-error;
    return *off;
  endmon;
  return *on;
end-proc;

dcl-proc year2;                         // 2-digit year window 1940-2039
  dcl-pi *n int(10);
    y int(10) value;
    nd int(10) value;
  end-pi;
  if nd <= 2;
    if y < 40;
      return 2000 + y;
    endif;
    return 1900 + y;
  endif;
  return y;
end-proc;

// time part "13:45[:10[.123]] [PM]" -> seconds + us
dcl-proc parseTime;
  dcl-pi *n ind;
    s0 varchar(64) const;
    sec int(10);
    us int(10);
  end-pi;
  dcl-s s varchar(64);
  dcl-s g int(10) dim(3);
  dcl-s ng int(10);
  dcl-s i int(10);
  dcl-s ch char(1);
  dcl-s cur int(10);
  dcl-s nd int(10);
  dcl-s fr varchar(9);
  dcl-s inFr ind;
  dcl-s pm char(1);
  s = %xlate(LOWER : UPPER : %trim(s0));
  pm = ' ';
  if %len(s) > 2 and (%subst(s : %len(s) - 1) = 'AM' or %subst(s : %len(s) - 1) = 'PM');
    pm = %subst(s : %len(s) - 1 : 1);
    s = %trimr(%subst(s : 1 : %len(s) - 2));
  endif;
  if %len(s) > 0 and %subst(s : %len(s) : 1) = 'Z';
    s = %subst(s : 1 : %len(s) - 1);
  endif;
  ng = 0;
  cur = 0;
  nd = 0;
  fr = '';
  inFr = *off;
  for i = 1 to %len(s);
    ch = %subst(s : i : 1);
    if ch >= '0' and ch <= '9';
      if inFr;
        if %len(fr) < 6;
          fr += ch;
        endif;
      else;
        cur = cur * 10 + %scan(ch : '0123456789') - 1;
        nd += 1;
        if nd > 2;
          return *off;
        endif;
      endif;
    elseif (ch = ':' or ch = '.') and not inFr and ng < 2 and nd > 0;
      ng += 1;
      g(ng) = cur;
      cur = 0;
      nd = 0;
    elseif (ch = '.' or ch = ',') and ng = 2 and nd > 0 and not inFr;
      inFr = *on;
    else;
      return *off;
    endif;
  endfor;
  if nd = 0 and not inFr;
    return *off;
  endif;
  if not inFr or ng = 2;
    ng += 1;
    g(ng) = cur;
  endif;
  if ng < 2;
    return *off;
  endif;
  if ng = 2;
    g(3) = 0;
  endif;
  if pm = 'P' and g(1) < 12;
    g(1) += 12;
  elseif pm = 'A' and g(1) = 12;
    g(1) = 0;
  endif;
  if g(1) > 24 or g(2) > 59 or g(3) > 59;
    return *off;
  endif;
  sec = g(1) * 3600 + g(2) * 60 + g(3);
  if sec > 86400;
    return *off;
  endif;
  us = 0;
  if fr <> '';
    fr = %subst(fr + '000000' : 1 : 6);
    us = %int(fr);
  endif;
  return *on;
end-proc;

// text -> date / date + time / time; ISO, Db2, d/m/y (DATFMT), yyyymmdd
dcl-proc parseDT;
  dcl-pi *n char(1);                    // ' ' no, D date, Z date + time, T time
    t0 like(U_T) const;
    r likeds(dt_t);
  end-pi;
  dcl-s s varchar(64);
  dcl-s dp varchar(64);
  dcl-s tp varchar(64);
  dcl-s g int(10) dim(3);
  dcl-s gl int(10) dim(3);
  dcl-s ng int(10);
  dcl-s i int(10);
  dcl-s k int(10);
  dcl-s ch char(1);
  dcl-s y int(10);
  dcl-s m int(10);
  dcl-s d int(10);
  dcl-s sep char(1);
  r.iso = '';
  r.sec = 0;
  r.us = 0;
  r.hasD = *off;
  r.hasT = *off;
  if %len(t0) < 4 or %len(t0) > 40;
    return ' ';
  endif;
  s = %trim(ebc(t0));
  if s = '';
    return ' ';
  endif;
  ch = %subst(s : 1 : 1);
  if ch < '0' or ch > '9';
    return ' ';
  endif;
  // time only?
  if %scan(':' : s) > 0 and %scan('-' : s) = 0 and %scan('/' : s) = 0
     and %scan(' ' : s) = 0 and %len(s) <= 15;
    if parseTime(s : r.sec : r.us);
      r.hasT = *on;
      return 'T';
    endif;
    return ' ';
  endif;
  if %len(s) = 8 and onlyDigits(s);
    if mkDate(%int(%subst(s : 1 : 4)) : %int(%subst(s : 5 : 2)) : %int(%subst(s : 7 : 2))
              : r.dd);
      r.hasD = *on;
      return 'D';
    endif;
    return ' ';
  endif;
  // split date / time
  ng = 0;
  k = 0;
  g = 0;
  gl = 0;
  sep = ' ';
  for i = 1 to %len(s);
    ch = %subst(s : i : 1);
    if ch >= '0' and ch <= '9';
      if ng = 0;
        ng = 1;
      endif;
      g(ng) = g(ng) * 10 + %scan(ch : '0123456789') - 1;
      gl(ng) += 1;
      if gl(ng) > 4;
        return ' ';
      endif;
    elseif (ch = '-' or ch = '/' or ch = '.') and ng < 3 and gl(ng) > 0;
      if sep = ' ';
        sep = ch;
      elseif ch <> sep;
        return ' ';
      endif;
      ng += 1;
    elseif ng = 3 and gl(3) > 0;
      k = i;                            // the time starts here
      leave;
    else;
      return ' ';
    endif;
  endfor;
  if ng <> 3 or gl(3) = 0;
    return ' ';
  endif;
  if gl(1) = 4;                         // y m d
    y = g(1);
    m = g(2);
    d = g(3);
  elseif gDatFmt = 'MDY';
    m = g(1);
    d = g(2);
    y = year2(g(3) : gl(3));
  elseif gDatFmt = 'YMD';
    y = year2(g(1) : gl(1));
    m = g(2);
    d = g(3);
  else;
    d = g(1);
    m = g(2);
    y = year2(g(3) : gl(3));
  endif;
  if gl(3) = 3 or (gl(1) <> 4 and gl(3) <> 2 and gl(3) <> 4);
    return ' ';
  endif;
  if not mkDate(y : m : d : r.dd);
    return ' ';
  endif;
  r.hasD = *on;
  if k = 0;
    return 'D';
  endif;
  tp = %subst(s : k);
  ch = %subst(tp : 1 : 1);
  if ch = 'T' or ch = ' ' or ch = '-';
    tp = %trim(%subst(tp : 2));
  else;
    return ' ';
  endif;
  if tp = '';
    return 'D';
  endif;
  if not parseTime(tp : r.sec : r.us);
    return ' ';
  endif;
  r.hasT = *on;
  return 'Z';
end-proc;

// =====================================================================================
// names
// =====================================================================================
dcl-proc fold;                          // accents off (job CCSID text)
  dcl-pi *n varchar(512);
    s varchar(512) const;
  end-pi;
  dcl-s f varchar(60);
  f = ebc(u'00E000E100E200E300E400E500E800E900EA00EB00EC00ED00EE00EF00F200F3'
        + u'00F400F500F600F900FA00FB00FC00FD00FF00E700F100C000C100C200C300C4'
        + u'00C500C800C900CA00CB00CC00CD00CE00CF00D200D300D400D500D600D900DA'
        + u'00DB00DC00DD00C700D1');
  return %xlate(f : 'aaaaaaeeeeiiiiooooouuuuyycnAAAAAAEEEEIIIIOOOOOUUUUYCN' : s);
end-proc;

dcl-proc nkey;                          // compare key: A-Z 0-9 only
  dcl-pi *n varchar(512);
    s varchar(512) const;
  end-pi;
  dcl-s w varchar(512);
  dcl-s r varchar(512);
  dcl-s i int(10);
  dcl-s ch char(1);
  w = %xlate(LOWER : UPPER : fold(s));
  r = '';
  for i = 1 to %len(w);
    ch = %subst(w : i : 1);
    if %scan(ch : UPPER + '0123456789') > 0;
      r += ch;
    endif;
  endfor;
  return r;
end-proc;

dcl-proc colName;                       // header -> SQL column name (max 30)
  dcl-pi *n varchar(30);
    s varchar(512) const;
  end-pi;
  dcl-s w varchar(512);
  dcl-s r varchar(512);
  dcl-s i int(10);
  dcl-s ch char(1);
  dcl-s us ind;
  w = %xlate(LOWER : UPPER : fold(s));
  r = '';
  us = *off;
  for i = 1 to %len(w);
    ch = %subst(w : i : 1);
    if %scan(ch : UPPER + '0123456789') > 0;
      if us and r <> '';
        r += '_';
      endif;
      us = *off;
      r += ch;
    else;
      us = *on;
    endif;
  endfor;
  if %len(r) > 30;
    r = %trimr(%subst(r : 1 : 30) : '_');
  endif;
  return r;
end-proc;

dcl-proc reserved;
  dcl-pi *n ind;
    s varchar(30) const;
  end-pi;
  dcl-s w1 varchar(600);
  w1 = ' ADD ALL ALTER AND ANY AS ASC BETWEEN BY CASE CAST CHECK COLUMN CONSTRAINT'
     + ' CREATE CURRENT CURSOR DEFAULT DELETE DESC DISTINCT DROP ELSE END ESCAPE'
     + ' EXCEPT EXISTS FETCH FOR FOREIGN FROM FULL GRANT GROUP HAVING IN INDEX INNER'
     + ' INSERT INTERSECT INTO IS JOIN KEY LEFT LIKE NOT NULL OF ON OR ORDER OUTER'
     + ' PRIMARY REFERENCES RIGHT ROW ROWS SELECT SET TABLE THEN TO UNION UNIQUE'
     + ' UPDATE USER USING VALUES VIEW WHEN WHERE WITH ';
  return %scan(' ' + s + ' ' : w1) > 0;
end-proc;

// SQL + system names of the area columns for a new table
dcl-proc makeNames;
  dcl-s a int(10);
  dcl-s j int(10);
  dcl-s n int(10);
  dcl-s b varchar(30);
  dcl-s c varchar(30);
  dcl-s dup ind;
  for a = 1 to nAc;
    b = colName(ebc(ac(a).hdr));
    if b = '';
      b = 'COL_' + ac(a).let;
    endif;
    if %scan(%subst(b : 1 : 1) : '0123456789') > 0;
      b = left('C_' + b : 30);
    endif;
    if reserved(b);
      b = left(b : 29) + '_';
    endif;
    c = b;
    n = 1;
    dow 1 = 1;
      dup = *off;
      for j = 1 to a - 1;
        if ac(j).nm = c;
          dup = *on;
          leave;
        endif;
      endfor;
      if not dup;
        leave;
      endif;
      n += 1;
      c = %trimr(left(b : 27)) + '_' + %char(n);
    enddo;
    ac(a).nm = c;
  endfor;
  // system names (10)
  for a = 1 to nAc;
    if %len(ac(a).nm) <= 10;
      c = ac(a).nm;
    else;
      c = %trimr(%subst(ac(a).nm : 1 : 10) : '_');
    endif;
    b = c;
    n = 0;
    dow 1 = 1;
      dup = *off;
      for j = 1 to nAc;
        if j <> a and ((j < a and ac(j).sys = c)
                       or (j > a and %len(ac(j).nm) <= 10 and ac(j).nm = c));
          dup = *on;
          leave;
        endif;
      endfor;
      if not dup;
        leave;
      endif;
      n += 1;
      c = %trimr(%subst(b + '__________' : 1 : 8) : '_')
        + %editc(%dec(n : 2 : 0) : 'X');
    enddo;
    ac(a).sys = c;
  endfor;
end-proc;

// =====================================================================================
// data area and header (after pass 0)
// =====================================================================================
dcl-proc headerTest;                    // sample k looks like a header row
  dcl-pi *n ind;
    k int(10) value;
  end-pi;
  dcl-s c int(10);
  dcl-s j int(10);
  dcl-s i int(10);
  dcl-s n int(10);
  dcl-s kk char(1);
  dcl-s t varchar(512);
  dcl-s hs varchar(512) dim(1000);
  n = 0;
  for c = aC1 to aC2;
    kk = smK(k : c);
    if kk = 'E';
      iter;
    endif;
    if kk <> 'S';
      gHdrHow = 'first row has non-text cells: no header';
      return *off;
    endif;
    n += 1;
  endfor;
  if n = 0;
    return *off;
  endif;
  if k = nSm;
    gHdrHow = 'only one row: taken as header';
    return *on;
  endif;
  for j = k + 1 to nSm;                 // text heading over numbers / dates
    for c = aC1 to aC2;
      if smK(k : c) = 'S' and %scan(smK(j : c) : 'NBd') > 0;
        gHdrHow = 'text over number / date / boolean values: header';
        return *on;
      endif;
    endfor;
  endfor;
  for c = aC1 to aC2;                   // distinct, and not repeated below
    if smK(k : c) <> 'S';
      iter;
    endif;
    t = %xlate(LOWER : UPPER : ebc(smText(k : c)));
    for i = aC1 to c - 1;
      if smK(k : i) = 'S' and hs(i - aC1 + 1) = t;
        gHdrHow = 'repeated texts in the first row: no header';
        return *off;
      endif;
    endfor;
    hs(c - aC1 + 1) = t;
    for j = k + 1 to nSm;
      if %xlate(LOWER : UPPER : ebc(smText(j : c))) = t;
        gHdrHow = 'first row texts repeat in the data: no header';
        return *off;
      endif;
    endfor;
  endfor;
  gHdrHow = 'distinct texts not repeated in the data: header';
  return *on;
end-proc;

// where the table is, before pass 0: RANGE / Excel table / autoFilter
dcl-proc decideSource;
  dcl-s k int(10);
  gAreaHow = '';
  if rHow = '' and gRange = '';
    k = findTable('' : *on : gShIdx);   // the first Excel table of the sheet
    if k = 0 and gAfRef <> '' and parseRef(ebc(gAfRef));
      rHdr = 1;
      rTot = 0;
      gAfExt = *on;                     // a stale filter range: contiguous rows below
      rHow = 'autoFilter ' + ebc(gAfRef);     // count too, a gap ends the table
    elseif k = 0;
      rR1 = 0;
      rR2 = 0;
      rC1 = 0;
      rC2 = 0;
    endif;
  endif;
  gAreaHow = rHow;
  gSmFrom = 1;
  if gStart > 0;
    gSmFrom = gStart;
  elseif rR1 > 0;
    gSmFrom = rR1;
  endif;
end-proc;

dcl-proc decideArea;
  dcl-pi *n ind;
  end-pi;
  dcl-s k int(10);
  dcl-s j int(10);
  dcl-s mx int(10);
  dcl-s thr int(10);
  dcl-s top int(20);
  dcl-s hdr ind;
  dcl-s a int(10);
  dcl-s c int(10);
  dcl-s v like(N_T);
  if nSm = 0;
    fail('no data found in sheet ' + ebc(sh(gShIdx).name));
    return *off;
  endif;
  // first row of the table
  k = 1;
  if gStart = 0 and rR1 = 0;            // heuristic: skip title rows
    mx = 0;
    for j = 1 to nSm;
      if sm(j).cnt > mx;
        mx = sm(j).cnt;
      endif;
    endfor;
    thr = %div(mx + 1 : 2);
    if thr < 2;
      thr = 2;
    endif;
    if mx = 1;
      thr = 1;
    endif;
    for j = 1 to nSm;
      if sm(j).cnt >= thr;
        k = j;
        leave;
      endif;
    endfor;
    if gAreaHow = '';
      gAreaHow = 'first rows of the sheet';
      if k > 1;
        gAreaHow = 'skipped ' + %char(k - 1) + ' title row(s) above row '
                 + %char(sm(k).row);
      endif;
    endif;
  endif;
  top = sm(k).row;
  if gStart > 0 and top <> gStart;      // STARTROW on a blank row
    top = gStart;
    k = 0;
  endif;
  if rR1 > 0 and gStart = 0 and top <> rR1;
    top = rR1;
    k = 0;
  endif;
  // columns
  aC1 = rC1;
  aC2 = rC2;
  if aC1 = 0 or aC2 = 0;
    for j = 1 to nSm;
      if sm(j).row < top;
        iter;
      endif;
      if aC1 = 0 or sm(j).c1 < aC1;
        if rC1 = 0;
          aC1 = sm(j).c1;
        endif;
      endif;
      if sm(j).c2 > aC2 and rC2 = 0;
        aC2 = sm(j).c2;
      endif;
    endfor;
  endif;
  if aC1 = 0 or aC2 < aC1;
    fail('no data found in sheet ' + ebc(sh(gShIdx).name));
    return *off;
  endif;
  nAc = aC2 - aC1 + 1;
  if nAc > MAXA;
    fail('the data area has ' + %char(nAc) + ' columns (max ' + %char(MAXA)
         + '): use RANGE');
    return *off;
  endif;
  // header
  select;
  when gHdrOpt = 'Y' or gHdrOpt = 'S';
    hdr = *on;
    gHdrHow = 'HEADER(*YES)';
    if gHdrOpt = 'S';
      gHdrHow = 'HEADER(*SKIP): header row not used for names';
    endif;
  when gHdrOpt = 'N';
    hdr = *off;
    gHdrHow = 'HEADER(*NO)';
  when rHdr >= 0;
    hdr = (rHdr > 0);
    gHdrHow = 'from ' + rHow;
  when k = 0;
    hdr = *off;
    gHdrHow = 'first row is empty: no header';
  other;
    hdr = headerTest(k);
  endsl;
  if hdr;
    aHdr = top;
    aFirst = top + 1;
  else;
    aHdr = 0;
    aFirst = top;
  endif;
  aLast = rR2;
  if rR2 > 0 and rTot > 0;
    aLast = rR2 - rTot;
  endif;
  // columns: letters, headings
  for a = 1 to nAc;
    c = aC1 + a - 1;
    clear ac(a);
    ac(a).xc = c;
    ac(a).let = colLetters(c);
    ac(a).hdr = '';
    if hdr and gHdrOpt <> 'S' and k > 0;
      v = smText(k : c);
      ac(a).hdr = v;
    endif;
  endfor;
  return *on;
end-proc;

// =====================================================================================
// pass 1: profile the columns
// =====================================================================================
dcl-proc rowHasData;
  dcl-pi *n ind;
  end-pi;
  dcl-s a int(10);
  for a = 1 to nAc;
    if ck(ac(a).xc) <> 'E';
      return *on;
    endif;
  endfor;
  return *off;
end-proc;

dcl-proc rowWanted;                     // the current row is a data row to read
  dcl-pi *n ind;
  end-pi;
  dcl-s a int(10);
  if gRowNo < aFirst;
    return *off;
  endif;
  if aLast > 0 and gRowNo > aLast;
    if gAfExt and gRowNo = aLast + 1 and rowHasData();
      aLast = gRowNo;
    else;
      gStop = *on;
      return *off;
    endif;
  endif;
  if gVis and gRowHid;
    return *off;
  endif;
  for a = 1 to nAc;
    if ck(ac(a).xc) <> 'E';
      return *on;
    endif;
  endfor;
  return *off;                          // blank row
end-proc;

dcl-proc mxLen;
  dcl-pi *n;
    a int(10) value;
    u8 int(10) value;
    ch int(10) value;
    u16 int(10) value;
  end-pi;
  if u8 > ac(a).mxU8;
    ac(a).mxU8 = u8;
  endif;
  if ch > ac(a).mxCh;
    ac(a).mxCh = ch;
  endif;
  if u16 > ac(a).mxU16;
    ac(a).mxU16 = u16;
  endif;
end-proc;

// quick UTF-8 test before parseDT: starts with a digit, only date / time characters
dcl-proc dtLike;
  dcl-pi *n ind;
    p pointer value;
    l int(10) value;
  end-pi;
  dcl-s b uns(3) based(pb);
  dcl-s pb pointer;
  dcl-s i int(10);
  dcl-s f ind;
  pb = p;
  f = *off;
  for i = 1 to l;
    select;
    when b >= 48 and b <= 57;
      f = *on;
    when b = 32 and not f;
    when not f;
      return *off;                      // first non-blank is not a digit
    when b = 32 or b = 45 or b = 47 or b = 46 or b = 58 or b = 44 or b = 43
         or b = 84 or b = 116 or b = 90 or b = 122 or b = 65 or b = 97
         or b = 80 or b = 112 or b = 77 or b = 109;
    other;
      return *off;
    endsl;
    pb += 1;
  endfor;
  return f;
end-proc;

dcl-proc statRow;
  dcl-s a int(10);
  dcl-s c int(10);
  dcl-s k char(1);
  dcl-s fk char(1);
  dcl-ds x likeds(num_t);
  dcl-s v packed(31 : 10);
  dcl-s p pointer;
  dcl-s l int(10);
  dcl-s nch int(10);
  dcl-s nu int(10);
  dcl-s nl ind;
  dcl-ds r likeds(dt_t);
  dcl-s dk char(1);
  dcl-s ln int(10);
  if not rowWanted();
    return;
  endif;
  if gMaxRows > 0 and gRead >= gMaxRows;
    gStop = *on;
    return;
  endif;
  gRead += 1;
  gLastSeen = gRowNo;
  for a = 1 to nAc;
    c = ac(a).xc;
    k = ck(c);
    select;
    when k = 'E';
      ac(a).cE += 1;
    when k = 'N';
      fk = fmtOf(c);
      if not numRaw(rcP(c) : rcL(c) : x);
        ac(a).cX += 1;
        iter;
      endif;
      numKeep(x : 15);
      if fk <> 'N';
        v = numDec(x);
        if v >= 0 and v < 1;
          ac(a).cT += 1;
          mxLen(a : 8 : 8 : 8);
        elseif x.ex >= 0;
          ac(a).cD += 1;
          mxLen(a : 10 : 10 : 10);
        else;
          ac(a).cZ += 1;
          mxLen(a : 23 : 23 : 23);
        endif;
        iter;
      endif;
      ac(a).cN += 1;
      if x.ex >= 0;
        ac(a).cInt += 1;
      endif;
      if numIntDig(x) > ac(a).iDig;
        ac(a).iDig = numIntDig(x);
      endif;
      if numSc(x) > ac(a).sc;
        ac(a).sc = numSc(x);
      endif;
      ln = numPlainLen(x);
      if padOf(c) > ln;
        ln = padOf(c);
        ac(a).pad = padOf(c);
      endif;
      mxLen(a : ln : ln : ln);
    when k = 'B';
      ac(a).cB += 1;
      mxLen(a : 5 : 5 : 5);
    when k = 'X';
      ac(a).cX += 1;
    other;                              // S d
      ac(a).cS += 1;
      p = cTxt(c : l);
      if rcKind(c) = 's';
        nch = sstChr(rcSst(c) + 1);
        nu = sstU16(rcSst(c) + 1);
        nl = (sstNL(rcSst(c) + 1) = 'N');
      else;
        utf8Info(p : l : nch : nu : nl);
      endif;
      mxLen(a : l : nch : nu);
      if nl;
        ac(a).nonLat = *on;
      endif;
      if l >= 6 and l <= 30 and dtLike(p : l);
        dk = parseDT(cText(c) : r);
        if dk = 'D';
          ac(a).cSD += 1;
        elseif dk = 'Z';
          ac(a).cSZ += 1;
        endif;
      endif;
    endsl;
  endfor;
end-proc;

// =====================================================================================
// types of a new table
// =====================================================================================

dcl-proc parseTypes;
  dcl-s s varchar(2000);
  dcl-s it varchar(300);
  dcl-s k int(10);
  dcl-s q int(10);
  nTy = 0;
  s = %trim(gTypes);
  dow s <> '' and nTy < %elem(tyKey);
    k = %scan(';' : s);
    if k > 0;
      it = %trim(%subst(s : 1 : k - 1));
      s = %trim(%subst(s : k + 1));
    else;
      it = s;
      s = '';
    endif;
    q = %scan('=' : it);
    if q > 1;
      nTy += 1;
      tyKey(nTy) = %trim(%subst(it : 1 : q - 1));
      tyVal(nTy) = %xlate(LOWER : UPPER : %trim(%subst(it : q + 1)));
    endif;
  enddo;
end-proc;

dcl-proc roundLen;                      // VARCHAR length with some room
  dcl-pi *n int(10);
    l int(10) value;
  end-pi;
  if l < 10;
    return 10;
  endif;
  return %div(l + 9 : 10) * 10;
end-proc;

dcl-proc inferTypes;
  dcl-s a int(10);
  dcl-s i int(10);
  dcl-s nv int(20);
  dcl-s dateish int(20);
  dcl-s p int(10);
  dcl-s l int(10);
  dcl-s ccs int(10);
  for a = 1 to nAc;
    nv = ac(a).cN + ac(a).cD + ac(a).cZ + ac(a).cT + ac(a).cB + ac(a).cS;
    dateish = ac(a).cD + ac(a).cZ + ac(a).cT + ac(a).cSD + ac(a).cSZ;
    select;
    when nv = 0;
      ac(a).ty = 'VARCHAR(10)';
    when ac(a).cS > 0 and ac(a).cS = ac(a).cSD + ac(a).cSZ and ac(a).cN = 0
         and ac(a).cB = 0 and ac(a).cT = 0;
      if ac(a).cSZ + ac(a).cZ > 0;
        ac(a).ty = 'TIMESTAMP';
      else;
        ac(a).ty = 'DATE';
      endif;
    when ac(a).cS > 0 or (ac(a).cN > 0 and dateish > 0);
      ac(a).ty = 'TEXT';
    when ac(a).cN > 0;
      if ac(a).cInt = ac(a).cN;
        if ac(a).iDig <= 9;
          ac(a).ty = 'INTEGER';
        elseif ac(a).iDig <= 18;
          ac(a).ty = 'BIGINT';
        elseif ac(a).iDig <= 31;
          ac(a).ty = 'DECIMAL(' + %char(ac(a).iDig) + ', 0)';
        else;
          ac(a).ty = 'DOUBLE';
        endif;
      elseif ac(a).sc <= 9 and ac(a).iDig + ac(a).sc <= 29;
        p = ac(a).iDig + ac(a).sc + 2;
        if p < ac(a).sc + 3;
          p = ac(a).sc + 3;
        endif;
        if p > 31;
          p = 31;
        endif;
        ac(a).ty = 'DECIMAL(' + %char(p) + ', ' + %char(ac(a).sc) + ')';
      else;
        ac(a).ty = 'DOUBLE';
      endif;
    when ac(a).cB > 0;
      ac(a).ty = 'SMALLINT';
    when ac(a).cZ > 0 or (ac(a).cT > 0 and ac(a).cD > 0);
      ac(a).ty = 'TIMESTAMP';
    when ac(a).cT > 0;
      ac(a).ty = 'TIME';
    other;
      ac(a).ty = 'DATE';
    endsl;
    if ac(a).ty = 'TEXT';
      ccs = gCcsOpt;
      if ccs = 0;                       // *AUTO
        ccs = -1;
        if ac(a).nonLat;
          ccs = 1208;
        endif;
      endif;
      select;
      when ccs = 1200;
        l = ac(a).mxU16;
        ac(a).ty = 'VARGRAPHIC(' + %char(roundLen(l)) + ') CCSID 1200';
        if roundLen(l) > 16370;
          ac(a).ty = 'DBCLOB(' + %char(l) + ') CCSID 1200';
        endif;
      when ccs = 1208;
        l = ac(a).mxU8;
        ac(a).ty = 'VARCHAR(' + %char(roundLen(l)) + ') CCSID 1208';
        if roundLen(l) > 32700;
          ac(a).ty = 'CLOB(' + %char(l) + ') CCSID 1208';
        endif;
      when ccs > 0;
        l = ac(a).mxU8;
        ac(a).ty = 'VARCHAR(' + %char(roundLen(l)) + ') CCSID ' + %char(ccs);
      other;
        l = ac(a).mxCh;
        ac(a).ty = 'VARCHAR(' + %char(roundLen(l)) + ')';
        if roundLen(l) > 32700;
          ac(a).ty = 'CLOB(' + %char(l) + ')';
        endif;
      endsl;
    endif;
    for i = 1 to nTy;                   // TYPES override
      if nkey(tyKey(i)) = nkey(ac(a).nm)
         or (nkey(tyKey(i)) = nkey(ebc(ac(a).hdr)) and ac(a).hdr <> '')
         or %xlate(LOWER : UPPER : tyKey(i)) = ac(a).let;
        ac(a).ty = tyVal(i);
        leave;
      endif;
    endfor;
  endfor;
end-proc;

dcl-proc q1;                            // 'text' for SQL (quotes doubled)
  dcl-pi *n varchar(200);
    s varchar(200) const;
  end-pi;
  return '''' + %scanrpl('''' : '''''' : s) + '''';
end-proc;

dcl-proc ddl;                           // CREATE TABLE statement
  dcl-pi *n varchar(32740);
  end-pi;
  dcl-s s varchar(32740);
  dcl-s a int(10);
  dcl-s f varchar(10);
  s = 'CREATE TABLE ' + gLib + '/' + gTab + ' (';
  for a = 1 to nAc;
    if a > 1;
      s += ', ';
    endif;
    s += ac(a).nm;
    if ac(a).sys <> ac(a).nm;
      s += ' FOR COLUMN ' + ac(a).sys;
    endif;
    s += ' ' + ac(a).ty;
    if not gNulls;
      s += ' NOT NULL WITH DEFAULT';
    endif;
  endfor;
  f = left(gTab : 9) + 'R';
  s += ') RCDFMT ' + f;
  return s;
end-proc;

dcl-proc execSql;                       // EXECUTE IMMEDIATE, error -> message
  dcl-pi *n ind;
    s varchar(32740) const;
  end-pi;
  dcl-s mt varchar(512);
  gStmt = s;
  exec sql execute immediate :gStmt;
  if sqlcode < 0;
    exec sql get diagnostics condition 1 :mt = MESSAGE_TEXT;
    fail('SQL ' + sqlstate + ': ' + %trim(mt) + ' (' + left(s : 120)
         + ')');
    return *off;
  endif;
  return *on;
end-proc;

dcl-proc createTable;
  dcl-pi *n ind;
  end-pi;
  dcl-s a int(10);
  dcl-s s varchar(32740);
  dcl-s t varchar(512);
  if not execSql(ddl());
    return *off;
  endif;
  gCreated = *on;
  t = left('From ' + ebc(sh(gShIdx).name) + ' of ' + gStmf : 50);
  execSql('LABEL ON TABLE ' + gLib + '/' + gTab + ' IS ' + q1(t));
  s = '';
  for a = 1 to nAc;
    if ac(a).hdr = '';
      iter;
    endif;
    if s <> '';
      s += ', ';
    endif;
    t = left(ebc(ac(a).hdr) : 50);
    s += ac(a).nm + ' TEXT IS ' + q1(t);
  endfor;
  if s <> '';
    execSql('LABEL ON COLUMN ' + gLib + '/' + gTab + ' (' + s + ')');
  endif;
  s = '';
  for a = 1 to nAc;
    if ac(a).hdr = '';
      iter;
    endif;
    if s <> '';
      s += ', ';
    endif;
    t = left(ebc(ac(a).hdr) : 60);
    s += ac(a).nm + ' IS ' + q1(t);
  endfor;
  if s <> '';
    execSql('LABEL ON COLUMN ' + gLib + '/' + gTab + ' (' + s + ')');
  endif;
  gFail = *off;                         // labels are optional
  gMsg = '';
  return *on;
end-proc;

// =====================================================================================
// target table: library, existence, columns
// =====================================================================================
dcl-proc resolveLib;
  dcl-pi *n ind;
  end-pi;
  dcl-s l varchar(10);
  dcl-s t varchar(10);
  t = gTab;
  if gLib = '*LIBL' or gLib = '';
    l = '';
    exec sql select OBJLIB into :l from table(QSYS2.OBJECT_STATISTICS('*LIBL', '*FILE',
             :t)) fetch first 1 rows only;
    if sqlcode = 0 and l <> '';
      gLib = %trim(l);
      return *on;
    endif;
    gLib = '*CURLIB';
  endif;
  if gLib = '*CURLIB';
    l = '';
    exec sql select SCHEMA_NAME into :l from QSYS2.LIBRARY_LIST_INFO
             where TYPE = 'CURRENT' fetch first 1 rows only;
    if sqlcode = 0 and l <> '';
      gLib = %trim(l);
    else;
      gLib = 'QGPL';
    endif;
  endif;
  return *on;
end-proc;

dcl-proc tableExists;
  dcl-pi *n ind;
  end-pi;
  dcl-s l varchar(10);
  dcl-s t varchar(10);
  dcl-s ty char(1);
  l = gLib;
  t = gTab;
  if gLib = 'QTEMP';                    // not in the catalog
    exec sql select 'T' into :ty from table(QSYS2.OBJECT_STATISTICS('QTEMP', '*FILE', :t)) x
             fetch first 1 rows only;
    return (sqlcode = 0);
  endif;
  exec sql select TABLE_TYPE into :ty from QSYS2.SYSTABLES
           where SYSTEM_TABLE_SCHEMA = :l and SYSTEM_TABLE_NAME = :t
           fetch first 1 rows only;
  return (sqlcode = 0);
end-proc;

dcl-proc blankOf;                       // value for an empty cell of a NOT NULL column
  dcl-pi *n varchar(256);
    k char(1) const;
    d0 varchar(256) const;
    hd char(1) const;
  end-pi;
  dcl-s d varchar(256);
  d = %trim(d0);
  if hd = 'Y' and d <> '';
    select;
    when %subst(d : 1 : 1) = '''' and %len(d) >= 2;
      return %scanrpl('''''' : '''' : %subst(d : 2 : %len(d) - 2));
    when %xlate(LOWER : UPPER : d) = 'CURRENT_DATE' or %xlate(LOWER : UPPER : d)
         = 'CURRENT DATE';
      return %char(gToday : *iso);
    when %xlate(LOWER : UPPER : d) = 'CURRENT_TIME' or %xlate(LOWER : UPPER : d)
         = 'CURRENT TIME';
      return %char(%time(gNowTs) : *iso);
    when %scan('CURRENT' : %xlate(LOWER : UPPER : d)) = 1;
      return %char(gNowTs);
    when %check('0123456789.-+E' : d) = 0;
      return d;
    endsl;
  endif;
  select;
  when k = 'N' or k = 'I' or k = 'F';
    return '0';
  when k = 'D';
    return '0001-01-01';
  when k = 'T';
    return '00.00.00';
  when k = 'Z';
    return '0001-01-01-00.00.00.000000';
  when k = 'B';
    return 'false';
  endsl;
  return '';
end-proc;

dcl-proc addCol;                       // one table column -> tc()
  dcl-pi *n;
    cn varchar(128) const;
    cs varchar(10) const;
    dt varchar(10) const;
    ln int(10) value;
    pr int(10) value;
    sc int(10) value;
    nu char(1) const;
    cc int(10) value;
    tx varchar(50) const;
    hd varchar(60) const;
    hdf char(1) const;
    gen varchar(10) const;
    df varchar(256) const;
  end-pi;
  nTc += 1;
  clear tc(nTc);
  tc(nTc).name = %trim(cn);
  tc(nTc).sys = %trim(cs);
  tc(nTc).dtype = %trim(dt);
  tc(nTc).len = ln;
  tc(nTc).prec = pr;
  tc(nTc).scale = sc;
  tc(nTc).nul = (nu = 'Y');
  tc(nTc).ccsid = cc;
  tc(nTc).txt = %trim(tx);
  tc(nTc).hdg = %trim(hd);
  tc(nTc).skip = (%trim(gen) = 'ALWAYS' or (hdf <> 'Y' and hdf <> 'N' and hdf <> 'I'));
  tc(nTc).unit = 'S';
  select;
  when dt = 'CHAR' or dt = 'VARCHAR' or dt = 'CLOB';
    tc(nTc).kind = 'C';
    if cc = 1208;
      tc(nTc).unit = 'U';
    elseif cc = 65535;
      tc(nTc).ebc = *on;
    endif;
  when dt = 'GRAPHIC' or dt = 'VARG' or dt = 'DBCLOB';
    tc(nTc).kind = 'C';
    tc(nTc).unit = 'G';
  when dt = 'DECIMAL' or dt = 'NUMERIC';
    tc(nTc).kind = 'N';
  when dt = 'SMALLINT';
    tc(nTc).kind = 'I';
    tc(nTc).imax = '32767';
  when dt = 'INTEGER';
    tc(nTc).kind = 'I';
    tc(nTc).imax = '2147483647';
  when dt = 'BIGINT';
    tc(nTc).kind = 'I';
    tc(nTc).imax = '9223372036854775807';
  when dt = 'FLOAT' or dt = 'DOUBLE' or dt = 'REAL' or dt = 'DECFLOAT';
    tc(nTc).kind = 'F';
  when dt = 'DATE';
    tc(nTc).kind = 'D';
  when dt = 'TIME';
    tc(nTc).kind = 'T';
  when dt = 'TIMESTMP';
    tc(nTc).kind = 'Z';
  when dt = 'BOOLEAN';
    tc(nTc).kind = 'B';
  other;
    tc(nTc).kind = 'X';               // BINARY BLOB ROWID XML DATALINK ...
  endsl;
  tc(nTc).dft = blankOf(tc(nTc).kind : df : hdf);
  tc(nTc).hasDf = hdf;
end-proc;

dcl-proc describeTarget;
  dcl-pi *n ind;
  end-pi;
  dcl-s l varchar(10);
  dcl-s t varchar(10);
  dcl-s cn varchar(128);
  dcl-s cs varchar(10);
  dcl-s dt varchar(10);
  dcl-s ln int(10);
  dcl-s pr int(10);
  dcl-s sc int(10);
  dcl-s nu char(1);
  dcl-s cc int(10);
  dcl-s tx varchar(50);
  dcl-s hd varchar(60);
  dcl-s hdf char(1);
  dcl-s gen varchar(10);
  dcl-s df varchar(256);
  dcl-s q varchar(2000);
  dcl-s cmd varchar(200);
  l = gLib;
  t = gTab;
  nTc = 0;
  // QTEMP is not in the QSYS2 catalog: same columns from a DSPFFD outfile
  if gLib = 'QTEMP';
    cmd = 'DSPFFD FILE(QTEMP/' + gTab + ') OUTPUT(*OUTFILE) OUTFILE(QTEMP/XLSX10F)';
    exec sql call QSYS2.QCMDEXC(:cmd);
    if sqlcode < 0;
      fail('table ' + gLib + '/' + gTab + ' not found');
      return *off;
    endif;
    q = 'select coalesce(nullif(trim(WHALI2), ''''), nullif(trim(WHALIS), ''''), '
      + 'trim(WHFLDE)), trim(WHFLDE), case WHFLDT '
      + 'when ''A'' then case WHVARL when ''Y'' then ''VARCHAR'' else ''CHAR'' end '
      + 'when ''H'' then case WHVARL when ''Y'' then ''VARCHAR'' else ''CHAR'' end '
      + 'when ''G'' then case WHVARL when ''Y'' then ''VARG'' else ''GRAPHIC'' end '
      + 'when ''O'' then ''CHAR'' when ''J'' then ''CHAR'' when ''E'' then ''CHAR'' '
      + 'when ''P'' then ''DECIMAL'' when ''S'' then ''NUMERIC'' '
      + 'when ''B'' then case when WHFLDP > 0 then ''DECIMAL'' when WHFLDB = 2 '
      + 'then ''SMALLINT'' when WHFLDB = 4 then ''INTEGER'' else ''BIGINT'' end '
      + 'when ''F'' then ''DOUBLE'' when ''L'' then ''DATE'' when ''T'' then ''TIME'' '
      + 'when ''Z'' then ''TIMESTMP'' when ''2'' then ''CLOB'' '
      + 'when ''3'' then ''DBCLOB'' else ''OTHER'' end, '
      + 'int(case when WHFLDT in (''G'', ''3'') then WHFLDG when WHFLDT = ''2'' '
      + 'then WHDBLL when WHVARL = ''Y'' then WHFLDB - 2 else WHFLDB end), '
      + 'int(WHFLDD), int(WHFLDP), WHNULL, int(WHCSID), trim(WHFTXT), '
      + 'varchar(trim(trim(WHCHD1) concat '' '' concat trim(WHCHD2) concat '' '' '
      + 'concat trim(WHCHD3)), 60), '
      + 'case when WHIDC = ''Y'' then ''I'' else ''Y'' end, '
      + 'case when WHIDC = ''Y'' then ''ALWAYS'' else '''' end, '
      + 'case when WHDFTL > 0 then varchar(trim(WHDFT), 256) else '''' end '
      + 'from QTEMP/XLSX10F where WHFLDT <> '' '' order by rrn(XLSX10F)';
    exec sql prepare SQ from :q;
    exec sql declare TQ cursor for SQ;
    exec sql open TQ;
    if sqlcode < 0;
      fail('table ' + gLib + '/' + gTab + ': DSPFFD outfile not readable, SQL '
           + %char(sqlcode));
      return *off;
    endif;
    dow nTc < %elem(tc);
      exec sql fetch TQ into :cn, :cs, :dt, :ln, :pr, :sc, :nu, :cc, :tx, :hd, :hdf,
                             :gen, :df;
      if sqlcode <> 0;
        leave;
      endif;
      addCol(cn : cs : dt : ln : pr : sc : nu : cc : tx : hd : hdf : gen : df);
    enddo;
    exec sql close TQ;
    exec sql drop table QTEMP/XLSX10F;
    if nTc = 0;
      fail('table ' + gLib + '/' + gTab + ' not found');
      return *off;
    endif;
    return *on;
  endif;
  exec sql declare TC cursor for
    select COLUMN_NAME, SYSTEM_COLUMN_NAME, DATA_TYPE, LENGTH,
           coalesce(NUMERIC_PRECISION, 0), coalesce(NUMERIC_SCALE, 0), IS_NULLABLE,
           coalesce(CCSID, 0), coalesce(COLUMN_TEXT, ''),
           coalesce(varchar(COLUMN_HEADING, 60), ''), coalesce(HAS_DEFAULT, 'N'),
           coalesce(IDENTITY_GENERATION, ''),
           coalesce(varchar(COLUMN_DEFAULT, 256), '')
      from QSYS2.SYSCOLUMNS
     where SYSTEM_TABLE_SCHEMA = :l and SYSTEM_TABLE_NAME = :t
     order by ORDINAL_POSITION;
  exec sql open TC;
  dow nTc < %elem(tc);
    exec sql fetch TC into :cn, :cs, :dt, :ln, :pr, :sc, :nu, :cc, :tx, :hd, :hdf,
                           :gen, :df;
    if sqlcode <> 0;
      leave;
    endif;
    addCol(cn : cs : dt : ln : pr : sc : nu : cc : tx : hd : hdf : gen : df);
  enddo;
  exec sql close TC;
  if nTc = 0;
    fail('table ' + gLib + '/' + gTab + ' not found');
    return *off;
  endif;
  return *on;
end-proc;

dcl-proc typeText;                      // DECIMAL(9,2) ... for the report
  dcl-pi *n varchar(40);
    t int(10) value;
  end-pi;
  dcl-s s varchar(40);
  s = tc(t).dtype;
  select;
  when tc(t).kind = 'N';
    s += '(' + %char(tc(t).prec) + ',' + %char(tc(t).scale) + ')';
  when tc(t).kind = 'C';
    s += '(' + %char(tc(t).len) + ')';
    if tc(t).ccsid > 0;
      s += ' CCSID ' + %char(tc(t).ccsid);
    endif;
  endsl;
  if not tc(t).nul;
    s += ' NOT NULL';
  endif;
  return s;
end-proc;

// =====================================================================================
// Excel columns -> table columns
// =====================================================================================
dcl-proc findCol;                       // table column by SQL or system name
  dcl-pi *n int(10);
    s varchar(128) const;
  end-pi;
  dcl-s t int(10);
  dcl-s w varchar(128);
  w = %xlate(LOWER : UPPER : %trim(s));
  if %len(w) >= 2 and %subst(w : 1 : 1) = '"';
    w = %subst(%trim(s) : 2 : %len(%trim(s)) - 2);
  endif;
  for t = 1 to nTc;
    if %xlate(LOWER : UPPER : tc(t).name) = w or tc(t).sys = w;
      return t;
    endif;
  endfor;
  return 0;
end-proc;

dcl-proc mapCols;
  dcl-pi *n ind;
  end-pi;
  dcl-s a int(10);
  dcl-s t int(10);
  dcl-s k int(10);
  dcl-s q int(10);
  dcl-s s varchar(2000);
  dcl-s it varchar(300);
  dcl-s ky varchar(300);
  dcl-s cv varchar(300);
  dcl-s hk varchar(512);
  dcl-s mode char(1);
  dcl-s nmatch int(10);
  dcl-s found ind;
  for a = 1 to nAc;
    ac(a).ti = 0;
  endfor;
  for t = 1 to nTc;
    tc(t).ai = 0;
  endfor;
  // COLMAP 'Header or letter=COLUMN;...'
  s = %trim(gColMap);
  dow s <> '';
    k = %scan(';' : s);
    if k > 0;
      it = %trim(%subst(s : 1 : k - 1));
      s = %trim(%subst(s : k + 1));
    else;
      it = s;
      s = '';
    endif;
    q = %scan('=' : it);
    if q <= 1;
      iter;
    endif;
    ky = %trim(%subst(it : 1 : q - 1));
    cv = %trim(%subst(it : q + 1));
    found = *off;
    for a = 1 to nAc;
      if (ac(a).hdr <> '' and nkey(ebc(ac(a).hdr)) = nkey(ky))
         or ac(a).let = %xlate(LOWER : UPPER : ky);
        found = *on;
        if %xlate(LOWER : UPPER : cv) = '*SKIP' or %xlate(LOWER : UPPER : cv) = '*OMIT';
          ac(a).ti = -1;
        else;
          t = findCol(cv);
          if t = 0;
            fail('COLMAP ' + it + ': column ' + cv + ' not in ' + gLib + '/' + gTab);
            return *off;
          endif;
          ac(a).ti = t;
          tc(t).ai = a;
        endif;
        leave;
      endif;
    endfor;
    if not found;
      fail('COLMAP ' + it + ': no Excel column with heading or letter ' + ky);
      return *off;
    endif;
  enddo;

  mode = gMapOpt;
  if gCreated;
    mode = 'P';
  elseif mode = 'A';
    mode = 'P';
    if aHdr > 0 and gHdrOpt <> 'S';
      nmatch = 0;
      for a = 1 to nAc;
        hk = nkey(ebc(ac(a).hdr));
        if hk = '';
          iter;
        endif;
        for t = 1 to nTc;
          if hk = nkey(tc(t).name) or hk = nkey(tc(t).sys) or hk = nkey(tc(t).txt)
             or hk = nkey(tc(t).hdg);
            nmatch += 1;
            leave;
          endif;
        endfor;
      endfor;
      if nmatch > 0;
        mode = 'N';
      endif;
    endif;
  endif;
  if mode = 'N';
    for a = 1 to nAc;
      if ac(a).ti <> 0;
        iter;
      endif;
      hk = nkey(ebc(ac(a).hdr));
      if hk = '';
        iter;
      endif;
      for k = 1 to 4;                   // name, system name, text, heading
        for t = 1 to nTc;
          if tc(t).ai > 0 or tc(t).skip;
            iter;
          endif;
          if (k = 1 and hk = nkey(tc(t).name)) or (k = 2 and hk = nkey(tc(t).sys))
             or (k = 3 and hk = nkey(tc(t).txt)) or (k = 4 and hk = nkey(tc(t).hdg));
            ac(a).ti = t;
            tc(t).ai = a;
            leave;
          endif;
        endfor;
        if ac(a).ti > 0;
          leave;
        endif;
      endfor;
    endfor;
  else;
    t = 0;
    for a = 1 to nAc;
      if ac(a).ti <> 0;
        iter;
      endif;
      dow 1 = 1;
        t += 1;
        if t > nTc or (not tc(t).skip and tc(t).ai = 0);
          leave;
        endif;
      enddo;
      if t > nTc;
        leave;
      endif;
      ac(a).ti = t;
      tc(t).ai = a;
    endfor;
  endif;
  gMapHow = 'by position';
  if mode = 'N';
    gMapHow = 'by name (column name, system name, text or heading)';
  endif;
  if gCreated;
    gMapHow = 'new table';
  endif;

  nMp = 0;
  for a = 1 to nAc;
    t = ac(a).ti;
    if t <= 0;
      iter;
    endif;
    if tc(t).skip;
      rep('warning: column ' + tc(t).name + ' is generated: Excel column ' + ac(a).let
          + ' not loaded');
      ac(a).ti = 0;
      iter;
    endif;
    if tc(t).kind = 'X';
      rep('warning: column ' + tc(t).name + ' (' + tc(t).dtype + ') not supported: '
          + 'Excel column ' + ac(a).let + ' not loaded');
      ac(a).ti = 0;
      iter;
    endif;
    nMp += 1;
    mpA(nMp) = a;
    mpT(nMp) = t;
  endfor;
  if nMp = 0;
    fail('no Excel column matches a column of ' + gLib + '/' + gTab
         + ' (use MAP(*POS) or COLMAP)');
    return *off;
  endif;
  for t = 1 to nTc;                     // NOT NULL without a value and without default
    if tc(t).ai = 0 and not tc(t).nul and not tc(t).skip and tc(t).hasDf = 'N';
      fail('column ' + tc(t).name + ' of ' + gLib + '/' + gTab + ' needs a value '
           + '(NOT NULL, no default) but no Excel column is mapped to it (COLMAP)');
      return *off;
    endif;
  endfor;
  return *on;
end-proc;

// =====================================================================================
// one value: cell -> text for the target column (UTF-8 in gVal), '' or the reason
// =====================================================================================
dcl-proc putE;                          // job CCSID text -> gVal
  dcl-pi *n;
    s varchar(400) const;
  end-pi;
  dcl-s w varchar(400) ccsid(*utf8);
  dcl-s b uns(3) based(pb);
  dcl-s pb pointer;
  dcl-s o uns(3) based(po);
  dcl-s po pointer;
  dcl-s i int(10);
  dcl-s k int(10);
  dcl-s set varchar(100);
  dcl-s su varchar(100) ccsid(*utf8);
  dcl-s sl varchar(400);
  dcl-s q uns(3) based(pq);
  dcl-s pq pointer;
  if gRawE;                             // non-text column: Db2 gets job CCSID text
    gVlen = %len(s);
    if gVlen > 0;
      sl = s;
      memcpy(%addr(gVal) : %addr(sl : *data) : gVlen);
    endif;
    return;
  endif;
  if not gA2Ok;                         // EBCDIC -> ASCII map of the characters of
    gA2Ok = *on;                        // numbers, dates and times (one conversion)
    set = '0123456789+-.,:/ ETZetz()%ABCDFabcdf';
    su = set;
    if %len(su) = %len(set);
      pb = %addr(set : *data);
      pq = %addr(su : *data);
      for k = 1 to %len(set);
        gA2(b + 1) = q;
        pb += 1;
        pq += 1;
      endfor;
    endif;
  endif;
  sl = s;
  pb = %addr(sl : *data);               // fast path: every byte in the map
  po = %addr(gVal);
  for i = 1 to %len(sl);
    k = gA2(b + 1);
    if k = 0;
      leave;
    endif;
    o = k;
    pb += 1;
    po += 1;
  endfor;
  if i > %len(sl);
    gVlen = %len(sl);
    return;
  endif;
  w = s;
  gVlen = %len(w);
  if gVlen > 0;
    memcpy(%addr(gVal) : %addr(w : *data) : gVlen);
  endif;
end-proc;

dcl-proc cutUtf8;                       // longest prefix with at most mx units
  dcl-pi *n int(10);                    // bytes
    p pointer value;
    n int(10) value;
    mx int(10) value;
    unit char(1) const;
  end-pi;
  dcl-s b uns(3) based(pb);
  dcl-s pb pointer;
  dcl-s i int(10);
  dcl-s cnt int(10);
  dcl-s w int(10);
  dcl-s uw int(10);
  i = 0;
  cnt = 0;
  dow i < n;
    pb = p + i;
    if b < 128;
      w = 1;
    elseif b < 224;
      w = 2;
    elseif b < 240;
      w = 3;
    else;
      w = 4;
    endif;
    select;
    when unit = 'U';
      uw = w;
    when unit = 'G' and w = 4;
      uw = 2;
    other;
      uw = 1;
    endsl;
    if cnt + uw > mx;
      leave;
    endif;
    cnt += uw;
    i += w;
  enddo;
  return i;
end-proc;

dcl-proc boolVal;                       // <v>1</v>
  dcl-pi *n ind;
    c int(10) value;
  end-pi;
  dcl-s b uns(3) based(pb);
  dcl-s pb pointer;
  if rcL(c) < 1;
    return *off;
  endif;
  pb = rcP(c);
  return (b = 49);
end-proc;

dcl-proc boolText;
  dcl-pi *n varchar(5);
    c int(10) value;
  end-pi;
  if boolVal(c);
    return 'TRUE';
  endif;
  return 'FALSE';
end-proc;

dcl-proc render;
  dcl-pi *n varchar(120);
    a int(10) value;
    t int(10) value;
  end-pi;
  dcl-s c int(10);
  dcl-s k char(1);
  dcl-s fk char(1);
  dcl-ds x likeds(num_t);
  dcl-s v packed(31 : 10);
  dcl-ds r likeds(dt_t);
  dcl-s dk char(1);
  dcl-s p pointer;
  dcl-s l int(10);
  dcl-s nch int(10);
  dcl-s nu int(10);
  dcl-s nl ind;
  dcl-s s varchar(400);
  dcl-s tx like(U_T);
  dcl-s n int(10);
  dcl-s mx int(10);
  dcl-s y int(10);
  dcl-s okN ind;
  dcl-s w varchar(20);

  c = ac(a).xc;
  k = ck(c);
  gVnull = *off;
  gVlen = 0;
  if k = 'X';
    gErrVal += 1;
    k = 'E';
  endif;
  if k = 'E';
    if tc(t).nul;
      gVnull = *on;
    else;
      putE(tc(t).dft);
    endif;
    return '';
  endif;
  okN = *off;
  if k = 'N';
    okN = numRaw(rcP(c) : rcL(c) : x);
    if not okN;
      return 'not a number: ' + ebc(cText(c));
    endif;
    numKeep(x : 15);
  endif;

  select;
  // ---------------------------------------------------------------- text
  when tc(t).kind = 'C';
    select;
    when k = 'N';
      fk = fmtOf(c);
      if fk <> 'N' and serialDT(numDec(x) : r);
        v = numDec(x);
        if v < 1;
          putE(tmText(r.sec : ':'));
        elseif x.ex >= 0;
          putE(dtIsoR(r));
        else;
          putE(tsText(r));
        endif;
      else;
        s = numPlain(x);
        if padOf(c) > %len(s) and not x.neg and x.ex >= 0;
          s = %subst('000000000000000000000000000000' : 1 : padOf(c) - %len(s)) + s;
        endif;
        putE(s);
      endif;
    when k = 'B';
      putE(boolText(c));
    other;
      p = cTxt(c : l);
      if l > 32740;
        l = 32740;
      endif;
      if l > 0;
        memcpy(%addr(gVal) : p : l);
      endif;
      gVlen = l;
    endsl;
    // length in the column's unit
    utf8Info(%addr(gVal) : gVlen : nch : nu : nl);
    select;
    when tc(t).unit = 'U';
      n = gVlen;
    when tc(t).unit = 'G';
      n = nu;
    other;
      n = nch;
      if nl and not tc(t).ebc and tc(t).ccsid <> 1208;
        gSubst += 1;
      endif;
    endsl;
    mx = tc(t).len;
    if tc(t).unit = 'G' and (tc(t).dtype = 'GRAPHIC' or tc(t).dtype = 'VARG');
      mx = tc(t).len;
    endif;
    if n > mx;
      if gErrMode = 'N';
        gVlen = cutUtf8(%addr(gVal) : gVlen : mx : tc(t).unit);
        gNulled += 1;
        return '';
      endif;
      return 'too long: ' + %char(n) + ' > ' + %char(mx);
    endif;
  // ---------------------------------------------------------------- numbers
  when tc(t).kind = 'N' or tc(t).kind = 'I' or tc(t).kind = 'F';
    select;
    when k = 'N';
    when k = 'B';
      if boolVal(c);
        numParseE('1' : x);
      else;
        numParseE('0' : x);
      endif;
    other;
      if not numText(cText(c) : x);
        return 'not a number: ' + left(ebc(cText(c)) : 40);
      endif;
    endsl;
    select;
    when tc(t).kind = 'N';
      numScale(x : tc(t).scale);
      if numIntDig(x) > tc(t).prec - tc(t).scale;
        return 'too large for DECIMAL(' + %char(tc(t).prec) + ',' + %char(tc(t).scale)
               + '): ' + numPlain(x);
      endif;
      putE(numPlain(x));
    when tc(t).kind = 'I';
      if x.ex < 0;
        return 'not an integer: ' + numPlain(x);
      endif;
      s = numPlain(x);
      w = tc(t).imax;
      n = numIntDig(x);
      if n > %len(w) or (n = %len(w) and %subst(s : %len(s) - n + 1) > w);
        return 'out of range for ' + tc(t).dtype + ': ' + s;
      endif;
      putE(s);
    other;
      putE(numSci(x));
    endsl;
  // ---------------------------------------------------------------- dates
  when tc(t).kind = 'D' or tc(t).kind = 'T' or tc(t).kind = 'Z';
    select;
    when k = 'N';
      v = numDec(x);
      if x.ex >= 0 and v >= 19000101 and v <= 99991231 and tc(t).kind <> 'T'
         and mkDate(%int(v / 10000) : %int(%rem(%int(v) : 10000) / 100)
                    : %rem(%int(v) : 100) : r.dd);
        r.sec = 0;                      // yyyymmdd as a number
        r.us = 0;
        dk = 'D';
      elseif serialDT(v : r);
        dk = 'Z';
        if v < 1;
          dk = 'T';
        endif;
      else;
        return 'not a date: ' + numPlain(x);
      endif;
    when k = 'B';
      return 'not a date: boolean';
    other;
      dk = parseDT(cText(c) : r);
      if dk = ' ';
        return 'not a date / time: ' + left(ebc(cText(c)) : 40);
      endif;
    endsl;
    select;
    when tc(t).kind = 'D';
      if dk = 'T';
        return 'time without a date';
      endif;
      putE(dtIsoR(r));
    when tc(t).kind = 'T';
      if dk = 'D';
        return 'date without a time';
      endif;
      putE(tmText(r.sec : '.'));
    other;
      if dk = 'T';
        return 'time without a date';
      endif;
      putE(tsDb2(r));
    endsl;
  // ---------------------------------------------------------------- boolean
  when tc(t).kind = 'B';
    select;
    when k = 'N';
      if x.nd > 0;
        putE('true');
      else;
        putE('false');
      endif;
    when k = 'B';
      if boolVal(c);
        putE('true');
      else;
        putE('false');
      endif;
    other;
      s = %xlate(LOWER : UPPER : %trim(ebc(cText(c))));
      if s = 'TRUE' or s = 'YES' or s = 'Y' or s = 'SI' or s = 'S' or s = '1' or s = 'X'
         or s = 'VERO';
        putE('true');
      elseif s = 'FALSE' or s = 'NO' or s = 'N' or s = '0' or s = 'FALSO' or s = '';
        putE('false');
      else;
        return 'not a boolean: ' + left(s : 40);
      endif;
    endsl;
  endsl;
  return '';
end-proc;


// =====================================================================================
// passes 2 and 3: convert, check, insert
// =====================================================================================

dcl-proc addErr;
  dcl-pi *n;
    row int(20) value;
    a int(10) value;
    why varchar(512) const;
  end-pi;
  dcl-s s varchar(600);
  gErrN += 1;
  s = 'row ' + %char(row);
  if a > 0;
    s += ' column ' + ac(a).let;
    if ac(a).hdr <> '';
      s += ' (' + left(ebc(ac(a).hdr) : 30) + ')';
    endif;
  endif;
  s += ': ' + why;
  if gErrN <= MAXERR;
    gErrTxt(gErrN) = left(s : 300);
  endif;
  if gErrN <= 1000;
    rep('error: ' + s);
  endif;
end-proc;

dcl-proc loadRow;
  dcl-s m int(10);
  dcl-s a int(10);
  dcl-s t int(10);
  dcl-s why varchar(120);
  dcl-s bad ind;
  dcl-s slot int(10);
  dcl-s o pointer;
  dcl-s ln int(5) based(o);
  dcl-s ip pointer;
  dcl-s iv int(5) based(ip);
  dcl-s ew varchar(32740);
  if not rowWanted();
    return;
  endif;
  if gMaxRows > 0 and gRead2 >= gMaxRows;
    gStop = *on;
    return;
  endif;
  gRead2 += 1;
  bad = *off;
  slot = nBuf + 1;
  for m = 1 to nMp;
    a = mpA(m);
    t = mpT(m);
    why = render(a : t);
    if why <> '';
      if gErrMode = 'N';
        gNulled += 1;
        gVnull = tc(t).nul;
        gVlen = 0;
        if not gVnull;
          putE(tc(t).dft);
        endif;
      else;
        addErr(gRowNo : a : why);
        bad = *on;
        if gErrMode = 'K' and not gDry;
          leave;
        endif;
        iter;
      endif;
    endif;
    if gDry or bad;
      iter;
    endif;
    o = pPar + (slot - 1) * gRowBytes + mpOff(m);
    ip = pInd + ((slot - 1) * nMp + m - 1) * 2;
    if gVnull;
      iv = -1;
      ln = 0;
      iter;
    endif;
    iv = 0;
    if tc(t).ebc;                       // CCSID 65535 column: pass job CCSID bytes
      ew = %subst(gVal : 1 : gVlen);
      gVlen = %len(ew);
      if gVlen > 0;
        memcpy(%addr(gVal) : %addr(ew : *data) : gVlen);
      endif;
    endif;
    if gVlen > mpCap(m);
      addErr(gRowNo : a : 'value longer than expected (' + %char(gVlen) + ')');
      bad = *on;
      iter;
    endif;
    ln = gVlen;
    if gVlen > 0;
      memcpy(o + 2 : %addr(gVal) : gVlen);
    endif;
  endfor;
  if bad;
    gRej += 1;
    if gDry and gErrN >= 1000;
      gStop = *on;
    endif;
    return;
  endif;
  if gDry;
    return;
  endif;
  bRow(slot) = gRowNo;
  nBuf += 1;
  if nBuf >= gBlk;
    flush();
  endif;
end-proc;

dcl-proc insText;                       // INSERT for n rows
  dcl-pi *n varchar(32740);
    n int(10) value;
  end-pi;
  dcl-s s varchar(32740);
  dcl-s g varchar(8000);
  dcl-s m int(10);
  dcl-s r int(10);
  s = 'INSERT INTO ' + gLib + '/' + gTab + ' (';
  g = '(';
  for m = 1 to nMp;
    if m > 1;
      s += ', ';
      g += ', ';
    endif;
    s += '"' + %scanrpl('"' : '""' : tc(mpT(m)).name) + '"';
    g += '?';
  endfor;
  s += ') VALUES ';
  g += ')';
  for r = 1 to n;
    if r > 1;
      s += ', ';
    endif;
    s += g;
  endfor;
  return s;
end-proc;

dcl-proc setupInsert;
  dcl-pi *n ind;
  end-pi;
  dcl-s m int(10);
  dcl-s r int(10);
  dcl-s j int(10);
  dcl-s t int(10);
  dcl-s a int(10);
  dcl-s cap int(10);
  dcl-s mt varchar(512);
  gRowBytes = 0;
  for m = 1 to nMp;
    t = mpT(m);
    a = mpA(m);
    select;
    when tc(t).kind = 'C';
      cap = ac(a).mxU8;
      if cap < 26;
        cap = 26;
      endif;
      if cap < %len(tc(t).dft) * 3;
        cap = %len(tc(t).dft) * 3;
      endif;
      if cap > 32740;
        cap = 32740;
      endif;
    when tc(t).kind = 'D';
      cap = 10;
    when tc(t).kind = 'T';
      cap = 8;
    when tc(t).kind = 'Z';
      cap = 26;
    when tc(t).kind = 'B';
      cap = 5;
    other;
      cap = 64;
    endsl;
    mpCap(m) = cap;
    mpOff(m) = gRowBytes;
    gRowBytes += cap + 2;
  endfor;
  gBlk = 1000;
  if gBlk * nMp > SQL_NUM;
    gBlk = %div(SQL_NUM : nMp);
  endif;
  j = %len(insText(1));                 // statement text max 32740
  r = %len(insText(2)) - j;             // bytes per further row
  if r > 0 and j + (gBlk - 1) * r > 32700;
    gBlk = %div(32700 - j : r) + 1;
  endif;
  if gBlk * gRowBytes > 67108864;
    gBlk = %div(67108864 : gRowBytes);
  endif;
  if getenv('XLSX10_BLK') <> *null;      // tuning: rows per INSERT
    monitor;
      j = %int(%str(getenv('XLSX10_BLK')));
      if j >= 1 and j < gBlk;
        gBlk = j;
      endif;
    on-error;
    endmon;
  endif;
  if gBlk < 1;
    gBlk = 1;
  endif;
  pPar = %alloc(gBlk * gRowBytes + 16);
  pInd = %alloc(gBlk * nMp * 2 + 16);
  SQLN = SQL_NUM;
  SQLD = gBlk * nMp;
  for r = 1 to gBlk;
    for m = 1 to nMp;
      j = (r - 1) * nMp + m;
      pv = %addr(SQL_VAR(j));
      sv.vType = 449;
      sv.vLen = mpCap(m);
      sv.vData = pPar + (r - 1) * gRowBytes + mpOff(m);
      sv.vInd = pInd + (j - 1) * 2;
      if tc(mpT(m)).ebc;
        sv.vNameLen = 0;
        sv.vName = *blanks;
      else;
        sv.vNameLen = 8;
        sv.vName = x'000004B800000000';
      endif;
    endfor;
  endfor;
  gStmt = insText(gBlk);
  exec sql prepare S1 from :gStmt;
  if sqlcode < 0;
    exec sql get diagnostics condition 1 :mt = MESSAGE_TEXT;
    fail('prepare INSERT: ' + %trim(mt));
    return *off;
  endif;
  gStmt = insText(1);
  exec sql prepare S2 from :gStmt;
  gPrepB = gBlk;
  nBuf = 0;
  setupStage();
  return *on;
end-proc;

dcl-proc setupStage;
  dcl-s m int(10);
  dcl-s t int(10);
  dcl-s o int(10);
  dcl-s v varchar(200);
  dcl-s e varchar(400);
  dcl-s s varchar(32740);
  dcl-s c varchar(32740);
  dcl-s mt varchar(512);
  gStgOn = *off;
  gStgWhy = '';
  if getenv('XLSX10_STAGE') <> *null;
    if %str(getenv('XLSX10_STAGE')) = 'N';
      gStgWhy = 'XLSX10_STAGE=N';
      return;
    endif;
  endif;
  o = 11;                               // 1-10: Excel row number
  for m = 1 to nMp;
    if tc(mpT(m)).ebc;
      gStgWhy = 'CCSID 65535 column';
      return;
    endif;
    stOff(m) = o;
    o += 6 + mpCap(m);
  endfor;
  gStgW = o - 1;
  if gStgW > 8000;
    gStgWhy = 'row of ' + %char(gStgW) + ' bytes';
    return;
  endif;
  gStgK = 'L';
  if gStgW <= 1000;
    gStgK = 'S';
  elseif gStgW <= 4000;
    gStgK = 'M';
  endif;
  s = 'INSERT INTO ' + gLib + '/' + gTab + ' (';
  c = '';
  for m = 1 to nMp;
    t = mpT(m);
    o = stOff(m);
    if tc(t).kind = 'C';
      v = 'SUBSTR(R, ' + %char(o + 6) + ', INT(SUBSTR(R, ' + %char(o + 1) + ', 5)))';
    else;
      v = 'RTRIM(SUBSTR(R, ' + %char(o + 6) + ', ' + %char(mpCap(m)) + '))';
    endif;
    select;
    when tc(t).kind = 'N';
      e = 'DEC(' + v + ', ' + %char(tc(t).prec) + ', ' + %char(tc(t).scale) + ', ''.'')';
    when tc(t).kind = 'I';
      e = 'BIGINT(' + v + ')';
    when tc(t).kind = 'F';
      e = 'DECFLOAT(' + v + ', 34, ''.'')';
    when tc(t).kind = 'D';
      e = 'DATE(' + v + ')';
    when tc(t).kind = 'T';
      e = 'TIME(' + v + ')';
    when tc(t).kind = 'Z';
      e = 'TIMESTAMP(' + v + ')';
    other;
      e = v;
    endsl;
    if tc(t).nul;
      e = 'CASE WHEN SUBSTR(R, ' + %char(o) + ', 1) = ''N'' THEN NULL ELSE ' + e + ' END';
    endif;
    if m > 1;
      s += ', ';
      c += ', ';
    endif;
    s += '"' + %scanrpl('"' : '""' : tc(t).name) + '"';
    c += e;
    if %len(s) + %len(c) > 32000;
      gStgWhy = 'statement over 32 KB';
      return;
    endif;
  endfor;
  s += ') SELECT ' + c + ' FROM QTEMP/XLSX10S S';
  exec sql drop table QTEMP/XLSX10S;
  gStmt = 'CREATE TABLE QTEMP/XLSX10S (R VARCHAR(' + %char(gStgW) + ') CCSID 1208 NOT NULL)';
  exec sql execute immediate :gStmt;
  if sqlcode < 0;
    exec sql get diagnostics condition 1 :mt = MESSAGE_TEXT;
    gStgWhy = 'staging table: ' + left(%trim(mt) : 100);
    return;
  endif;
  gStmt = s;
  exec sql prepare SI from :gStmt;
  if sqlcode < 0;
    exec sql get diagnostics condition 1 :mt = MESSAGE_TEXT;
    gStgWhy = 'INSERT ... SELECT: ' + left(%trim(mt) : 100);
    exec sql drop table QTEMP/XLSX10S;
    return;
  endif;
  gStmt = s + ' WHERE RRN(S) > ?';
  exec sql prepare SJ from :gStmt;
  gStmt = 'SELECT COUNT(*) FROM ' + gLib + '/' + gTab;
  exec sql prepare SC from :gStmt;
  exec sql declare SCN cursor for SC;
  gStgOn = *on;
  gStgMade = *on;
  gStgN = 0;
  gChunk = %div(67108864 : gStgW);
  if gChunk > 1000000;
    gChunk = 1000000;
  endif;
end-proc;

dcl-proc stageRows;                     // block rows r0.. (nr) -> staging strings 1..nr
  dcl-pi *n;
    r0 int(10) value;
    nr int(10) value;
  end-pi;
  dcl-s r int(10);
  dcl-s e int(10);
  dcl-s br int(10);
  dcl-s m int(10);
  dcl-s l int(10);
  dcl-s d int(10);
  dcl-s k int(10);
  dcl-s po pointer;
  dcl-s ob uns(3) based(pob);
  dcl-s pob pointer;
  dcl-s o pointer;
  dcl-s ln int(5) based(o);
  dcl-s ip pointer;
  dcl-s iv int(5) based(ip);
  for e = 1 to nr;
    r = e;
    br = r0 + e - 1;
    select;
    when gStgK = 'S';
      %len(stgS(e).r) = gStgW;
      po = %addr(stgS(e).r : *data);
    when gStgK = 'M';
      %len(stgM(e).r) = gStgW;
      po = %addr(stgM(e).r : *data);
    other;
      %len(stg(e).r) = gStgW;
      po = %addr(stg(e).r : *data);
    endsl;
    memset(po : 32 : gStgW);
    d = bRow(br);
    for k = 10 downto 1;
      pob = po + k - 1;
      ob = 48 + %rem(d : 10);
      d = %div(d : 10);
    endfor;
    for m = 1 to nMp;
      o = pPar + (br - 1) * gRowBytes + mpOff(m);
      ip = pInd + ((br - 1) * nMp + m - 1) * 2;
      pob = po + stOff(m) - 1;
      if iv < 0;
        ob = 78;                        // N
        iter;
      endif;
      ob = 48;
      l = ln;
      d = l;
      for k = 5 downto 1;               // 5 ASCII digits
        pob = po + stOff(m) - 1 + k;
        ob = 48 + %rem(d : 10);
        d = %div(d : 10);
      endfor;
      if l > 0;
        memcpy(po + stOff(m) + 5 : o + 2 : l);
      endif;
    endfor;
  endfor;
end-proc;

dcl-proc dbErrT;                        // Db2 error with its text for row
  dcl-pi *n;
    row int(20) value;
    code int(10) value;
    mt varchar(512) const;
  end-pi;
  addErr(row : 0 : 'Db2 SQL' + %char(code) + ': ' + %trim(mt));
  gRej += 1;
  if gErrMode = 'S';
    gStop = *on;
  endif;
end-proc;

dcl-proc dbErr;                         // error of the last statement for row
  dcl-pi *n;
    row int(20) value;
  end-pi;
  dcl-s mt varchar(512);
  exec sql get diagnostics condition 1 :mt = MESSAGE_TEXT;
  addErr(row : 0 : 'Db2 SQL' + %char(sqlcode) + ': ' + %trim(mt));
  gRej += 1;
  if gErrMode = 'S';
    gStop = *on;
  endif;
end-proc;

dcl-proc insOne;                        // block row r alone (moved to slot 1)
  dcl-pi *n;
    r int(10) value;
  end-pi;
  if r > 1;
    memcpy(pPar : pPar + (r - 1) * gRowBytes : gRowBytes);
    memcpy(pInd : pInd + (r - 1) * nMp * 2 : nMp * 2);
  endif;
  SQLD = nMp;
  exec sql execute S2 using descriptor :SQLDA;
  if sqlcode < 0;
    dbErr(bRow(r));
  else;
    gIns += 1;
  endif;
end-proc;

dcl-proc tabCount;                      // rows in the target table
  dcl-pi *n int(20);
  end-pi;
  dcl-s n int(20);
  n = -1;
  exec sql open SCN;
  if sqlcode = 0;
    exec sql fetch SCN into :n;
    exec sql close SCN;
  endif;
  return n;
end-proc;

dcl-proc stgFlush;                      // staging rows -> table, Db2 errors located
  dcl-s base int(20);
  dcl-s k int(20);
  dcl-s rr int(20);
  dcl-s code int(10);
  dcl-s mt varchar(512);
  dcl-s xr varchar(10);
  dcl-s row int(20);
  dcl-s n0 int(20);
  if gStgN = 0;
    return;
  endif;
  gTi = %timestamp();
  base = 0;
  dow base < gStgN;
    n0 = tabCount();                    // SQLERRD(3) is 0 when INSERT ... SELECT fails
    if base = 0;
      exec sql execute SI;
    else;
      exec sql execute SJ using :base;
    endif;
    code = sqlcode;
    if code >= 0;
      gIns += gStgN - base;
      leave;
    endif;
    exec sql get diagnostics condition 1 :mt = MESSAGE_TEXT;
    k = tabCount() - n0;                // rows in before the failing one (no other
    if k < 0 or k >= gStgN - base;      // job inserting into the table meanwhile)
      k = 0;
    endif;
    gIns += k;
    rr = base + k + 1;
    row = 0;
    exec sql select substr(R, 1, 10) into :xr from QTEMP/XLSX10S S where RRN(S) = :rr;
    if sqlcode = 0;
      row = toInt(xr);
    endif;
    dbErrT(row : code : mt);
    if gStop;
      leave;
    endif;
    base = rr;
  enddo;
  gStgN = 0;
  exec sql truncate table QTEMP/XLSX10S immediate;   // RRNs start again at 1
  if sqlcode < 0;
    exec sql drop table QTEMP/XLSX10S;
    gStmt = 'CREATE TABLE QTEMP/XLSX10S (R VARCHAR(' + %char(gStgW)
          + ') CCSID 1208 NOT NULL)';
    exec sql execute immediate :gStmt;
  endif;
  gTu += %diff(%timestamp() : gTi : *mseconds);
end-proc;

dcl-proc flush;
  dcl-s k int(10);
  dcl-s r int(10);
  dcl-s n int(10);
  dcl-s code int(10);
  dcl-s mt varchar(512);
  if nBuf = 0;
    return;
  endif;
  if gErrMode = 'S' and gErrN > 0;      // ERRORS(*STOP) after a Db2 error: no more rows
    nBuf = 0;
    return;
  endif;
  SQLD = nBuf * nMp;
  code = 0;
  k = 0;
  mt = '';
  if gStgOn;                            // blocked INSERT into the staging table
    r = 1;
    dou r > nBuf or sqlcode < 0;
      n = nBuf - r + 1;
      if gStgK = 'L' and n > %elem(stg);
        n = %elem(stg);
      endif;
      stageRows(r : n);
      gTi = %timestamp();
      select;
      when gStgK = 'S';
        exec sql insert into QTEMP/XLSX10S :n rows values(:stgS);
      when gStgK = 'M';
        exec sql insert into QTEMP/XLSX10S :n rows values(:stgM);
      other;
        exec sql insert into QTEMP/XLSX10S :n rows values(:stg);
      endsl;
      gTu += %diff(%timestamp() : gTi : *mseconds);
      r += n;
    enddo;
    if sqlcode >= 0;
      gStgN += nBuf;
      nBuf = 0;
      if gStgN >= gChunk;
        stgFlush();
      endif;
      return;
    endif;
    exec sql get diagnostics condition 1 :mt = MESSAGE_TEXT;   // not expected:
    gStgWhy = 'blocked INSERT: ' + left(%trim(mt) : 100);      // VALUES from now on
    stgFlush();
    gStgOn = *off;
    mt = '';
    if gStop;
      nBuf = 0;
      return;
    endif;
  endif;
  if code = 0;                          // multi-row VALUES insert
    gTi = %timestamp();
    if nBuf = gPrepB;
      exec sql execute S1 using descriptor :SQLDA;
    else;
      gStmt = insText(nBuf);
      exec sql prepare S3 from :gStmt;
      exec sql execute S3 using descriptor :SQLDA;
    endif;
    code = sqlcode;
    k = sqlerrd(3);
    if code < 0;
      exec sql get diagnostics condition 1 :mt = MESSAGE_TEXT;
    endif;
    gTu += %diff(%timestamp() : gTi : *mseconds);
  endif;
  if code < 0;                          // rows 1..k are in, k+1 failed, rest one by one
    if k < 0 or k >= nBuf;
      k = 0;
    endif;
    gIns += k;
    dbErrT(bRow(k + 1) : code : mt);
    if not gStop;
      for r = k + 2 to nBuf;
        insOne(r);
        if gStop;
          leave;
        endif;
      endfor;
    endif;
  else;
    gIns += nBuf;
  endif;
  nBuf = 0;
  SQLD = gPrepB * nMp;
end-proc;

// =====================================================================================
// report file
// =====================================================================================
dcl-proc writeReport;
  dcl-s fd int(10);
  if gRepPath = '';
    return;
  endif;
  fd = ifsOpen(gRepPath : O_WRONLY + O_CREAT + O_TRUNC + O_CCSID : 436 : 1208);
  if fd < 0;
    return;
  endif;
  ifsWrite(fd : %addr(gRep : *data) : %len(gRep));
  ifsClose(fd);
end-proc;

dcl-proc repColumns;
  dcl-s a int(10);
  dcl-s s varchar(1000);
  dcl-s t int(10);
  rep('');
  rep('Columns (letter | heading | non-empty: numbers / integers / dates / times / '
      + 'date-times / booleans / texts / errors | max length | type):');
  for a = 1 to nAc;
    s = '  ' + ac(a).let + ' | ' + left(ebc(ac(a).hdr) : 40) + ' | '
        + %char(ac(a).cN) + '/' + %char(ac(a).cInt) + '/'
        + %char(ac(a).cD + ac(a).cSD) + '/' + %char(ac(a).cT) + '/'
        + %char(ac(a).cZ + ac(a).cSZ) + '/' + %char(ac(a).cB) + '/'
        + %char(ac(a).cS - ac(a).cSD - ac(a).cSZ) + '/' + %char(ac(a).cX)
        + ' | ' + %char(ac(a).mxCh);
    if ac(a).nonLat;
      s += ' (Unicode)';
    endif;
    t = ac(a).ti;
    if t > 0;
      s += ' | -> ' + tc(t).name + ' ' + typeText(t);
    elseif ac(a).ty <> '';
      s += ' | ' + ac(a).nm + ' ' + ac(a).ty;
    else;
      s += ' | not loaded';
    endif;
    rep(s);
  endfor;
end-proc;

// =====================================================================================
// main
// =====================================================================================
dcl-proc xlsxRead;
  dcl-pi *n extpgm('XLSX10R');
    pStmf    char(1024) const;
    pToFile  char(20) const;
    pRows    packed(15 : 0) options(*nopass : *omit);
    pMsg     char(256) options(*nopass : *omit);
    pSheet   char(64) const options(*nopass : *omit);
    pRange   char(256) const options(*nopass : *omit);
    pHeader  char(10) const options(*nopass : *omit);
    pStart   packed(7 : 0) const options(*nopass : *omit);
    pRowSel  char(10) const options(*nopass : *omit);
    pMaxRows packed(9 : 0) const options(*nopass : *omit);
    pMbrOpt  char(10) const options(*nopass : *omit);
    pCrtFile char(10) const options(*nopass : *omit);
    pMap     char(10) const options(*nopass : *omit);
    pColMap  char(2000) const options(*nopass : *omit);
    pTypes   char(2000) const options(*nopass : *omit);
    pNulls   char(4) const options(*nopass : *omit);
    pErrors  char(10) const options(*nopass : *omit);
    pStrCcs  char(10) const options(*nopass : *omit);
    pTrim    char(4) const options(*nopass : *omit);
    pDatFmt  char(10) const options(*nopass : *omit);
    pReport  char(1024) const options(*nopass : *omit);
  end-pi;
  dcl-s msgOut ind;
  dcl-s x varchar(64);
  dcl-s t0 timestamp;
  dcl-s secs packed(11 : 3);
  dcl-s txt varchar(512);
  dcl-s i int(10);
  dcl-s needCheck ind;
  dcl-s crt ind;
  dcl-s n int(10);
  dcl-s p pointer;

  t0 = %timestamp();
  gNowTs = t0;
  gToday = %date(t0);
  gFail = *off;
  gMsg = '';
  gRep = '';
  msgOut = *off;
  if %parms() >= 4;
    if %addr(pMsg) <> *null;
      if %subst(pMsg : 1 : 7) <> '*ESCAPE';
        msgOut = *on;
        pMsg = *blanks;
      endif;
    endif;
  endif;
  if %parms() >= 3;
    if %addr(pRows) <> *null;
      pRows = 0;
    endif;
  endif;

  // ---- options
  gStmf = %trim(pStmf);
  gTab = %xlate(LOWER : UPPER : %trim(%subst(pToFile : 1 : 10)));
  gLib = %xlate(LOWER : UPPER : %trim(%subst(pToFile : 11 : 10)));
  gNone = (gTab = '*NONE' or gTab = '');
  gSheetSel = '';
  gRange = '';
  gHdrOpt = 'A';
  gStart = 0;
  gVis = *off;
  gMaxRows = 0;
  gMbrRpl = *off;
  gCrt = 'Y';
  gMapOpt = 'A';
  gColMap = '';
  gTypes = '';
  gNulls = *off;
  gErrMode = 'S';
  gCcsOpt = 0;
  gTrim = *on;
  gDatFmt = 'DMY';
  gRepPath = '';
  if %parms() >= 5 and %addr(pSheet) <> *null;
    x = %trim(pSheet);
    if x <> '*FIRST' and x <> '';
      gSheetSel = u(x);
    endif;
  endif;
  if %parms() >= 6 and %addr(pRange) <> *null;
    gRange = %trim(pRange);
    if %xlate(LOWER : UPPER : gRange) = '*AUTO';
      gRange = '';
    endif;
  endif;
  if %parms() >= 7 and %addr(pHeader) <> *null;
    x = %xlate(LOWER : UPPER : %trim(pHeader));
    select;
    when x = '*YES';
      gHdrOpt = 'Y';
    when x = '*NO';
      gHdrOpt = 'N';
    when x = '*SKIP';
      gHdrOpt = 'S';
    when x = '*AUTO' or x = '';
    other;
      fail('HEADER ' + x + ' not valid (*AUTO, *YES, *NO, *SKIP)');
    endsl;
  endif;
  if %parms() >= 8 and %addr(pStart) <> *null;
    gStart = pStart;
  endif;
  if %parms() >= 9 and %addr(pRowSel) <> *null;
    gVis = (%xlate(LOWER : UPPER : %trim(pRowSel)) = '*VISIBLE');
  endif;
  if %parms() >= 10 and %addr(pMaxRows) <> *null;
    gMaxRows = pMaxRows;
  endif;
  if %parms() >= 11 and %addr(pMbrOpt) <> *null;
    gMbrRpl = (%xlate(LOWER : UPPER : %trim(pMbrOpt)) = '*REPLACE');
  endif;
  if %parms() >= 12 and %addr(pCrtFile) <> *null;
    x = %xlate(LOWER : UPPER : %trim(pCrtFile));
    select;
    when x = '*NO';
      gCrt = 'N';
    when x = '*REPLACE';
      gCrt = 'R';
    endsl;
  endif;
  if %parms() >= 13 and %addr(pMap) <> *null;
    x = %xlate(LOWER : UPPER : %trim(pMap));
    select;
    when x = '*NAME';
      gMapOpt = 'N';
    when x = '*POS';
      gMapOpt = 'P';
    endsl;
  endif;
  if %parms() >= 14 and %addr(pColMap) <> *null;
    gColMap = %trim(pColMap);
    if %xlate(LOWER : UPPER : gColMap) = '*NONE';
      gColMap = '';
    endif;
  endif;
  if %parms() >= 15 and %addr(pTypes) <> *null;
    gTypes = %trim(pTypes);
    if %xlate(LOWER : UPPER : gTypes) = '*NONE';
      gTypes = '';
    endif;
  endif;
  if %parms() >= 16 and %addr(pNulls) <> *null;
    gNulls = (%xlate(LOWER : UPPER : %trim(pNulls)) = '*YES');
  endif;
  if %parms() >= 17 and %addr(pErrors) <> *null;
    x = %xlate(LOWER : UPPER : %trim(pErrors));
    select;
    when x = '*SKIP';
      gErrMode = 'K';
    when x = '*NULL';
      gErrMode = 'N';
    endsl;
  endif;
  if %parms() >= 18 and %addr(pStrCcs) <> *null;
    x = %xlate(LOWER : UPPER : %trim(pStrCcs));
    select;
    when x = '*JOB';
      gCcsOpt = -1;
    when x = '*UTF8';
      gCcsOpt = 1208;
    when x = '*UTF16';
      gCcsOpt = 1200;
    when onlyDigits(x);
      gCcsOpt = toInt(x);
    endsl;
  endif;
  if %parms() >= 19 and %addr(pTrim) <> *null;
    gTrim = (%xlate(LOWER : UPPER : %trim(pTrim)) <> '*NO');
  endif;
  if %parms() >= 20 and %addr(pDatFmt) <> *null;
    x = %xlate(LOWER : UPPER : %trim(pDatFmt));
    if x = '*MDY' or x = '*YMD';
      gDatFmt = %subst(x : 2 : 3);
    endif;
  endif;
  if %parms() >= 21 and %addr(pReport) <> *null;
    gRepPath = %trim(pReport);
    if %xlate(LOWER : UPPER : gRepPath) = '*NONE';
      gRepPath = '';
    endif;
  endif;

  rowTxtCap = 1048576;
  pRowTxt = %alloc(rowTxtCap);
  pSmBuf = %alloc(65536);
  pSmK = %alloc(MAXS * MAXC + 16);
  pSmO = %alloc(MAXS * MAXC * 4 + 16);
  pSmL = %alloc(MAXS * MAXC * 4 + 16);

  rep('XLSX10R ' + XLSX_VERSION + ' ' + %char(t0));
  rep('file: ' + gStmf);

  // ---- read the workbook
  if not gFail;
    if zrOpen(gStmf) < 0;
      fail(xlErr());
    else;
      gZrOpen = *on;
    endif;
  endif;
  if not gFail and loadWorkbook();
    loadStyles();
    gTq = %timestamp();
    if loadSst() and pickSheet() and loadSheet();
      gTm(1) = %diff(%timestamp() : t0 : *mseconds) / 1000;
      decideSource();
      gPass = 0;
      nSm = 0;
      gTq = %timestamp();
      smUse = 0;
      scanSheet();
      gTm(2) += %diff(%timestamp() : gTq : *mseconds) / 1000;
      if decideArea();
        gPass = 1;
        gTq = %timestamp();
        gRead = 0;
        scanSheet();
        gTm(3) += %diff(%timestamp() : gTq : *mseconds) / 1000;
        gAfExt = *off;                  // the area end is fixed now
        rep('sheet: ' + ebc(sh(gShIdx).name) + ' (' + ebc(sh(gShIdx).part) + ')');
        rep('area: columns ' + colLetters(aC1) + '-' + colLetters(aC2) + ', first data row '
            + %char(aFirst) + ', last ' + %char(gLastSeen) + ', ' + %char(gRead)
            + ' rows; ' + gAreaHow);
        if aHdr > 0;
          rep('header: row ' + %char(aHdr) + ' (' + gHdrHow + ')');
        else;
          rep('header: none (' + gHdrHow + ')');
        endif;
        if gDate1904;
          rep('date system: 1904');
        endif;
      endif;
    endif;
  endif;

  // ---- target
  if not gFail;
    if gNone;                           // analysis only
      makeNames();
      parseTypes();
      inferTypes();
      repColumns();
      rep('');
      gLib = 'LIB';
      gTab = 'TABLE';
      rep('CREATE statement: ' + ddl());
    else;
      resolveLib();
      gExists = tableExists();
      crt = (not gExists and gCrt <> 'N') or (gExists and gCrt = 'R');
      if not gExists and gCrt = 'N';
        fail('table ' + gLib + '/' + gTab + ' not found and CRTFILE(*NO)');
      endif;
      if not gFail and gExists and gCrt = 'R';
        execSql('DROP TABLE ' + gLib + '/' + gTab);
      endif;
      if not gFail and crt;
        makeNames();
        parseTypes();
        inferTypes();
        rep('');
        rep('CREATE statement: ' + ddl());
        createTable();
      endif;
      if not gFail;
        describeTarget();
      endif;
      if not gFail;
        mapCols();
      endif;
      if not gFail;
        rep('table: ' + gLib + '/' + gTab + ', columns mapped ' + gMapHow);
        repColumns();
        rep('');
      endif;
      if not gFail;
        setupInsert();
        if gStgOn;
          rep('load: blocked INSERT into QTEMP/XLSX10S (' + %char(gStgW)
              + ' bytes per row) + INSERT ... SELECT, ' + %char(gBlk) + ' rows per block');
        elseif not gFail;
          rep('load: multi-row INSERT ... VALUES, ' + %char(gBlk) + ' rows per statement ('
              + gStgWhy + ')');
        endif;
      endif;
      // pass 2: check everything first (existing table or TYPES given)
      needCheck = (not gCreated or gTypes <> '') and gErrMode = 'S';
      if not gFail and needCheck;
        gPass = 2;
        gDry = *on;
        gTq = %timestamp();
        gRead2 = 0;
        scanSheet();
        gDry = *off;
        gTm(4) += %diff(%timestamp() : gTq : *mseconds) / 1000;
        if gErrN > 0;
          fail(%char(gErrN) + ' value(s) do not fit ' + gLib + '/' + gTab
               + ', nothing loaded. First: ' + gErrTxt(1));
          if gCreated;
            execSql('DROP TABLE ' + gLib + '/' + gTab);
            gFail = *on;
          endif;
        endif;
        gRej = 0;
      endif;
      if not gFail and gMbrRpl and not gCreated;
        gStmt = 'TRUNCATE TABLE ' + gLib + '/' + gTab + ' IMMEDIATE';
        exec sql execute immediate :gStmt;
        if sqlcode < 0;
          execSql('DELETE FROM ' + gLib + '/' + gTab);
        endif;
      endif;
      if not gFail;
        gTq = %timestamp();
        gPass = 3;
        gRead2 = 0;
        scanSheet();
        flush();
        if gStgOn and not (gErrMode = 'S' and gErrN > 0);
          stgFlush();
        endif;
        gTm(5) += %diff(%timestamp() : gTq : *mseconds) / 1000;
        if gErrN > 0 and gErrMode = 'S';
          fail('stopped at a Db2 error after ' + %char(gIns) + ' rows: ' + gErrTxt(1));
        endif;
      endif;
    endif;
  endif;

  // ---- end
  if gStgMade;
    exec sql drop table QTEMP/XLSX10S;
    if not gStgOn;
      rep('load switched to INSERT ... VALUES: ' + gStgWhy);
    endif;
  endif;
  secs = %diff(%timestamp() : t0 : *mseconds) / 1000000;
  rep('');
  rep('timing (ms): open + inflate + strings ' + %char(gTm(1)) + ', pass 0 sample '
      + %char(gTm(2)) + ', pass 1 profile ' + %char(gTm(3)) + ', pass 2 check '
      + %char(gTm(4)) + ', pass 3 load ' + %char(gTm(5)) + ' (Db2 INSERT '
      + %char(%div(gTu : 1000)) + ', ' + %char(gBlk) + ' rows per block), total '
      + %char(%int(%diff(%timestamp() : t0 : *mseconds) / 1000)));
  if gZrOpen;
    zrClose();
  endif;
  if pSh <> *null;
    dealloc(n) pSh;
  endif;
  if gFail;
    if %parms() >= 3;                   // rows already in (ERRORS(*STOP) at a Db2 error)
      if %addr(pRows) <> *null;
        pRows = gIns;
      endif;
    endif;
    rep('');
    rep('FAILED: ' + gMsg);
    writeReport();
    for i = 1 to %min(gErrN : MAXERR);
      if not msgOut;
        sendMsg('*DIAG' : gErrTxt(i));
      endif;
    endfor;
    if msgOut;
      pMsg = gMsg;
    else;
      sendMsg('*ESCAPE' : PGMID + ' ' + XLSX_VERSION + ': ' + gMsg);
    endif;
    return;
  endif;
  if gNone;
    txt = PGMID + ' ' + XLSX_VERSION + ': analysed ' + %char(gRead) + ' rows x '
        + %char(nAc) + ' columns of sheet "' + ebc(sh(gShIdx).name) + '" in '
        + %char(secs) + ' s (TOFILE(*NONE): nothing written)';
  else;
    txt = PGMID + ' ' + XLSX_VERSION + ': ' + %char(gIns) + ' rows from sheet "'
        + ebc(sh(gShIdx).name) + '" -> ' + gLib + '/' + gTab;
    if gCreated;
      txt += ' (created)';
    endif;
    if gRej > 0;
      txt += ', ' + %char(gRej) + ' rejected';
    endif;
    if gNulled > 0;
      txt += ', ' + %char(gNulled) + ' values nulled / cut';
    endif;
    if gSubst > 0;
      txt += ', ' + %char(gSubst) + ' texts with characters not in the CCSID';
    endif;
    txt += ' in ' + %char(secs) + ' s';
  endif;
  rep('');
  rep(txt);
  writeReport();
  if not msgOut;
    for i = 1 to %min(gErrN : MAXERR);
      sendMsg('*DIAG' : gErrTxt(i));
    endfor;
    sendMsg('*COMP' : txt);
  endif;
  if %parms() >= 3;
    if %addr(pRows) <> *null;
      pRows = gIns;
    endif;
  endif;
end-proc;
