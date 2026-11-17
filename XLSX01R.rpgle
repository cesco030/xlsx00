**free
// =====================================================================================
// XLSX01R - ZIP writer with streaming deflate (RFC 1951), CRC-32, IFS helpers
//
// Part of the XLSX writer (XLSX00R). No dependencies: IFS and C runtime only (QC2LE).
// One archive at a time (module state), one entry at a time, entries streamed:
//   zwOpen -> { zwBegin -> zwWrite / zwReserve+zwCommit -> zwEnd | zwRaw }* -> zwClose
//
// Deflate: LZ77 with a 32 KB window (hash heads + chains; greedy matching at levels 1-4,
// zlib-style lazy matching at levels 5-9, effort set by the level) and per-block choice
// of fixed or dynamic Huffman codes (16K symbols per block). Validated against zlib by
// the Python mirrors proto/defl.py (greedy) and proto/lazy.py (lazy), same algorithm.
// Level 0 stores (method 0). No ZIP64: entries and archive must stay below 4 GB.
//
// Build: CRTRPGMOD MODULE(QTEMP/XLSX01R) SRCSTMF('<dir>/XLSX01R.rpgle') INCDIR('<dir>')
//          DBGVIEW(*SOURCE) OPTIMIZE(*FULL) TGTCCSID(*JOB)       (see XLSXDIST.clle)
// Version 1.4.1 (2026-11-17): product version XLSX_VERSION in XLSXINT.rpgleinc
// =====================================================================================
ctl-opt nomain alloc(*teraspace) bnddir('QC2LE')
        option(*srcstmt : *nodebugio : *nounref);

/copy 'XLSXINT.rpgleinc'

// ---- IFS / C prototypes --------------------------------------------------------------
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
dcl-pr ifsSeek int(20) extproc('lseek64');
  fd     int(10) value;
  offset int(20) value;
  whence int(10) value;
end-pr;
dcl-pr ifsUnlink int(10) extproc('unlink');
  path pointer value options(*string);
end-pr;
dcl-pr ifsAccess int(10) extproc('access');
  path pointer value options(*string);
  amode int(10) value;
end-pr;
dcl-pr renameUnlink int(10) extproc('Qp0lRenameUnlink');
  oldPath pointer value options(*string);
  newPath pointer value options(*string);
end-pr;
dcl-pr errnoPtr pointer extproc('__errno');
end-pr;
dcl-pr strerror pointer extproc('strerror');
  errnum int(10) value;
end-pr;

dcl-c O_WRONLY     2;
dcl-c O_CREAT      8;
dcl-c O_CCSID      32;
dcl-c O_TRUNC      64;
dcl-c O_LARGEFILE  536870912;
dcl-c MODE_664     436;             // rw-rw-r--
dcl-c SEEK_SET     0;
dcl-c SEEK_END     2;

// ---- buffers / sizes -----------------------------------------------------------------
dcl-c HIST     32768;               // deflate window kept in front of the input
dcl-c IBCHUNK  1048576;             // input bytes compressed per pass
dcl-c IBMAX    1605632;             // HIST + IBCHUNK + XL_RESMAX
dcl-c OBMAX    262144;              // output buffer
dcl-c OBFLUSH  261000;
dcl-c BLK      16384;               // symbols per deflate block
dcl-c MINLOOK  262;                 // lookahead kept when not final (258 + 4)

dcl-s pIB pointer inz(*null);       // input (uncompressed) buffer
dcl-s IB  uns(3) dim(1605632) based(pIB);
dcl-s pOB pointer inz(*null);       // output (to file) buffer
dcl-s OB  uns(3) dim(262144) based(pOB);

// ---- archive state -------------------------------------------------------------------
dcl-s gErr    varchar(512);
dcl-s gFd     int(10) inz(-1);
dcl-s gPath   varchar(1024);
dcl-s gLevel  int(10);
dcl-s filePos int(20);              // bytes already written to the file
dcl-s obLen   int(10);
dcl-s gDTime  int(10);
dcl-s gDDate  int(10);
dcl-s inEntry ind inz(*off);

dcl-ds cd qualified dim(4000);      // central directory (XL_MAXENT)
  name   varchar(512) ccsid(*utf8);
  method int(10);
  flags  int(10);
  dtime  int(10);
  ddate  int(10);
  crc    uns(10);
  csize  uns(10);
  usize  uns(10);
  offset uns(10);
end-ds;
dcl-s nCd int(10);

// ---- current entry -------------------------------------------------------------------
dcl-s eMethod int(10);
dcl-s eHdrOff int(20);              // offset of the local header
dcl-s eDatOff int(20);              // offset of the first data byte
dcl-s eCrc    uns(10);
dcl-s eUsize  int(20);
dcl-s ibLen   int(10);              // bytes in IB
dcl-s ibPos   int(10);              // next byte to compress
dcl-s crcPos  int(10);              // next byte to add to the CRC
dcl-s gBase   int(20);              // stream position of IB(1)

// ---- deflate state -------------------------------------------------------------------
dcl-s HEAD  int(20) dim(32768);     // hash -> last position + 1 (0 = empty)
dcl-s PREV  int(20) dim(32768);     // position -> previous position + 1
dcl-s maxChain int(10);
dcl-s niceLen  int(10);
dcl-s maxIns   int(10);
dcl-s useLazy  ind inz(*off);       // levels 5-9: lazy matching (zlib deflate_slow)
dcl-s goodLen  int(10);             // previous match >= goodLen: search 1/4 of the chain
dcl-s lazyLen  int(10);             // previous match >= lazyLen: no search, take it
dcl-s pvLen    int(10);             // lazy state: match found at the previous position
dcl-s pvDist   int(10);
dcl-s pvAvail  ind;                 // previous position not emitted yet
dcl-s bitBuf   uns(20);
dcl-ds bwd;                         // overlay: low 4 bytes of the bit buffer
  bw  uns(20) pos(1);
  bb5 uns(3) pos(5);
  bb6 uns(3) pos(6);
  bb7 uns(3) pos(7);
  bb8 uns(3) pos(8);
end-ds;
dcl-s bitCnt   int(10);
dcl-s SLIT  uns(5) dim(16384);      // literal byte, or match length
dcl-s SDIST uns(5) dim(16384);      // 0 = literal, else match distance
dcl-s nSym  int(10);
dcl-s FL    int(10) dim(288);       // literal/length frequencies (286 used)
dcl-s FD    int(10) dim(288);       // distance frequencies (30 used)

// ---- static tables (zfInit) ----------------------------------------------------------
dcl-s tabsDone ind inz(*off);
dcl-s CRCT  uns(10) dim(256);
dcl-s CB0   uns(10) dim(256);       // big-endian slicing-by-4 tables (zlib crc32_big)
dcl-s CB1   uns(10) dim(256);
dcl-s CB2   uns(10) dim(256);
dcl-s CB3   uns(10) dim(256);
dcl-ds gSt likeds(xlStat_t);
dcl-s P2    uns(20) dim(64);        // P2(k+1) = 2**k
dcl-s LSYM  int(10) dim(259);       // match length -> symbol (257..285)
dcl-s LEX   int(10) dim(259);       //               extra bits
dcl-s LEV   int(10) dim(259);       //               extra value
dcl-s DC1   int(10) dim(256);       // distance 1..256 -> code
dcl-s DC2   int(10) dim(256);       // (distance-1)/128 -> code (distance > 256)
dcl-s DBASE int(10) dim(30);
dcl-s DEXT  int(10) dim(30);
dcl-s LEXT  int(10) dim(29);
dcl-s FIXL  int(10) dim(288);       // fixed literal/length code lengths
dcl-s FIXC  uns(10) dim(288);       // fixed codes, bit reversed
dcl-s FIXDL int(10) dim(288);
dcl-s FIXDC uns(10) dim(288);
dcl-s CLORD int(10) dim(19);


// ---- local prototypes (procedures used before their definition) ----------------------
dcl-pr zfInit;
end-pr;
dcl-pr hufCodes;
  lens  int(10) dim(288) const;
  codes uns(10) dim(288);
  nsym  int(10) value;
end-pr;
dcl-pr compress;
  final ind value;
end-pr;
dcl-pr flushBlock;
  last ind value;
end-pr;
dcl-pr compLazy;
  final ind value;
end-pr;

// =====================================================================================
// errors
// =====================================================================================
dcl-proc xlSetErr export;
  dcl-pi *n;
    msg varchar(512) const;
  end-pi;
  gErr = msg;
end-proc;

dcl-proc xlErr export;
  dcl-pi *n varchar(512);
  end-pi;
  return gErr;
end-proc;

dcl-proc sysErr;                     // text of the current errno
  dcl-pi *n varchar(256);
  end-pi;
  dcl-s pe pointer;
  dcl-s e int(10) based(pe);
  pe = errnoPtr();
  return %str(strerror(e));
end-proc;

// =====================================================================================
// CRC-32
// =====================================================================================
dcl-proc bswap;                      // byte swap of a 32-bit value
  dcl-pi *n uns(10);
    v uns(10) value;
  end-pi;
  dcl-s r uns(20);
  dcl-s t uns(20);
  t = %bitand(v : 255);
  r = t * 16777216;
  t = %bitand(%div(v : 256) : 255);
  r += t * 65536;
  t = %bitand(%div(v : 65536) : 255);
  r += t * 256;
  r += %div(v : 16777216);
  return r;
end-proc;

// slicing by 4 on big-endian words (method of zlib crc32_big)
dcl-proc xlCrc32 export;
  dcl-pi *n uns(10);
    crc  uns(10) value;
    data pointer value;
    len  int(20) value;
  end-pi;
  dcl-s c uns(10);
  dcl-s p pointer;
  dcl-s b uns(3) based(p);
  dcl-s w uns(10) based(p);
  dcl-s n int(20);
  dcl-s sh uns(20);

  if not tabsDone;
    zfInit();
  endif;
  c = 4294967295 - bswap(crc);
  p = data;
  n = len;
  dow n >= 4;
    c = %bitxor(c : w);
    c = %bitxor(%bitxor(CB0(%bitand(c : 255) + 1)
                       : CB1(%bitand(%div(c : 256) : 255) + 1))
              : %bitxor(CB2(%bitand(%div(c : 65536) : 255) + 1)
                       : CB3(%div(c : 16777216) + 1)));
    p += 4;
    n -= 4;
  enddo;
  dow n > 0;
    sh = %bitand(c : 16777215);
    sh *= 256;
    c = %bitxor(CB0(%bitxor(%div(c : 16777216) : b) + 1) : sh);
    p += 1;
    n -= 1;
  enddo;
  return bswap(4294967295 - c);
end-proc;

dcl-proc xlStats export;
  dcl-pi *n;
    st likeds(xlStat_t);
  end-pi;
  st = gSt;
end-proc;

// =====================================================================================
// IFS helpers
// =====================================================================================
dcl-proc xlExists export;
  dcl-pi *n ind;
    path varchar(1024) const;
  end-pi;
  return (ifsAccess(%trimr(path) : 0) = 0);
end-proc;

dcl-proc xlRename export;
  dcl-pi *n int(10);
    fromPath varchar(1024) const;
    toPath   varchar(1024) const;
  end-pi;
  if renameUnlink(%trimr(fromPath) : %trimr(toPath)) <> 0;
    xlSetErr('rename ' + fromPath + ' -> ' + toPath + ': ' + sysErr());
    return -1;
  endif;
  return 0;
end-proc;

dcl-proc xlUnlink export;
  dcl-pi *n;
    path varchar(1024) const;
  end-pi;
  ifsUnlink(%trimr(path));
end-proc;

// =====================================================================================
// tables
// =====================================================================================
dcl-proc zfInit;
  dcl-s n int(10);
  dcl-s k int(10);
  dcl-s c uns(10);
  dcl-s code int(10);
  dcl-s e int(10);
  dcl-s l int(10);
  dcl-s d int(10);
  dcl-s poly uns(10) inz(3988292384);
  dcl-s LB int(10) dim(29);

  // CRC table, polynomial 0xEDB88320
  for n = 0 to 255;
    c = n;
    for k = 1 to 8;
      if %bitand(c : 1) <> 0;
        c = %bitxor(poly : %div(c : 2));
      else;
        c = %div(c : 2);
      endif;
    endfor;
    CRCT(n + 1) = c;
  endfor;
  for n = 0 to 255;
    c = CRCT(n + 1);
    CB0(n + 1) = bswap(c);
    c = %bitxor(CRCT(%bitand(c : 255) + 1) : %div(c : 256));
    CB1(n + 1) = bswap(c);
    c = %bitxor(CRCT(%bitand(c : 255) + 1) : %div(c : 256));
    CB2(n + 1) = bswap(c);
    c = %bitxor(CRCT(%bitand(c : 255) + 1) : %div(c : 256));
    CB3(n + 1) = bswap(c);
  endfor;

  P2(1) = 1;
  for k = 2 to 64;
    if k <= 63;
      P2(k) = P2(k - 1) * 2;
    endif;
  endfor;

  // length codes 257..285: base and extra bits
  LB(1) = 3;
  LB(2) = 4;
  LB(3) = 5;
  LB(4) = 6;
  LB(5) = 7;
  LB(6) = 8;
  LB(7) = 9;
  LB(8) = 10;
  LB(9) = 11;
  LB(10) = 13;
  LB(11) = 15;
  LB(12) = 17;
  LB(13) = 19;
  LB(14) = 23;
  LB(15) = 27;
  LB(16) = 31;
  LB(17) = 35;
  LB(18) = 43;
  LB(19) = 51;
  LB(20) = 59;
  LB(21) = 67;
  LB(22) = 83;
  LB(23) = 99;
  LB(24) = 115;
  LB(25) = 131;
  LB(26) = 163;
  LB(27) = 195;
  LB(28) = 227;
  LB(29) = 258;
  for code = 1 to 29;
    if code <= 8 or code = 29;
      LEXT(code) = 0;
    else;
      LEXT(code) = %div(code - 9 : 4) + 1;
    endif;
  endfor;
  for code = 1 to 29;
    for e = 0 to P2(LEXT(code) + 1) - 1;
      l = LB(code) + e;
      if l <= 258 and (l < 258 or code = 29);
        LSYM(l + 1) = 256 + code;
        LEX(l + 1) = LEXT(code);
        LEV(l + 1) = e;
      endif;
    endfor;
  endfor;

  // distance codes 0..29
  DBASE(1) = 1;
  for code = 1 to 30;
    if code <= 4;
      DEXT(code) = 0;
    else;
      DEXT(code) = %div(code - 3 : 2);
    endif;
    if code > 1;
      DBASE(code) = DBASE(code - 1) + P2(DEXT(code - 1) + 1);
    endif;
  endfor;
  for code = 1 to 30;
    for d = DBASE(code) to DBASE(code) + P2(DEXT(code) + 1) - 1;
      if d <= 256;
        DC1(d) = code - 1;
      else;
        DC2(%div(d - 1 : 128) + 1) = code - 1;
      endif;
    endfor;
  endfor;

  // fixed Huffman code
  for n = 0 to 287;
    if n <= 143;
      FIXL(n + 1) = 8;
    elseif n <= 255;
      FIXL(n + 1) = 9;
    elseif n <= 279;
      FIXL(n + 1) = 7;
    else;
      FIXL(n + 1) = 8;
    endif;
  endfor;
  hufCodes(FIXL : FIXC : 288);
  FIXDL = 0;
  for n = 1 to 30;
    FIXDL(n) = 5;
  endfor;
  hufCodes(FIXDL : FIXDC : 30);

  CLORD(1) = 16;
  CLORD(2) = 17;
  CLORD(3) = 18;
  CLORD(4) = 0;
  CLORD(5) = 8;
  CLORD(6) = 7;
  CLORD(7) = 9;
  CLORD(8) = 6;
  CLORD(9) = 10;
  CLORD(10) = 5;
  CLORD(11) = 11;
  CLORD(12) = 4;
  CLORD(13) = 12;
  CLORD(14) = 3;
  CLORD(15) = 13;
  CLORD(16) = 2;
  CLORD(17) = 14;
  CLORD(18) = 1;
  CLORD(19) = 15;

  tabsDone = *on;
end-proc;

dcl-proc bitRev;                     // reverse the low n bits of v
  dcl-pi *n uns(10);
    v uns(10) value;
    n int(10) value;
  end-pi;
  dcl-s r uns(10);
  dcl-s k int(10);
  r = 0;
  for k = 1 to n;
    r = r * 2 + %rem(v : 2);
    v = %div(v : 2);
  endfor;
  return r;
end-proc;

// canonical Huffman codes from code lengths, bit reversed (deflate sends LSB first)
dcl-proc hufCodes;
  dcl-pi *n;
    lens  int(10) dim(288) const;
    codes uns(10) dim(288);
    nsym  int(10) value;
  end-pi;
  dcl-s blc int(10) dim(16);
  dcl-s nxt int(10) dim(16);
  dcl-s s int(10);
  dcl-s b int(10);
  dcl-s c int(10);
  dcl-s l int(10);

  blc = 0;
  for s = 1 to nsym;
    blc(lens(s) + 1) += 1;
  endfor;
  blc(1) = 0;
  c = 0;
  nxt = 0;
  for b = 1 to 15;
    c = (c + blc(b)) * 2;
    nxt(b + 1) = c;
  endfor;
  codes = 0;
  for s = 1 to nsym;
    l = lens(s);
    if l > 0;
      codes(s) = bitRev(nxt(l + 1) : l);
      nxt(l + 1) += 1;
    endif;
  endfor;
end-proc;

// Huffman code lengths (two-queue method), halving the frequencies until max <= limit
dcl-proc hufLens;
  dcl-pi *n;
    frq   int(10) dim(288) const;
    lens  int(10) dim(288);
    nsym  int(10) value;
    limit int(10) value;
  end-pi;
  dcl-s f    int(10) dim(288);
  dcl-s lv   int(10) dim(288);       // leaves (symbols) sorted by frequency
  dcl-s iw   int(10) dim(288);       // weights of internal nodes
  dcl-s parL int(10) dim(288);
  dcl-s parI int(10) dim(288);
  dcl-s idep int(10) dim(288);
  dcl-s n  int(10);
  dcl-s s  int(10);
  dcl-s j  int(10);
  dcl-s k  int(10);
  dcl-s t  int(10);
  dcl-s li int(10);
  dcl-s ii int(10);
  dcl-s w  int(10);
  dcl-s d  int(10);
  dcl-s mx int(10);
  dcl-s takeLeaf ind;

  f = frq;
  dou mx <= limit;
    lens = 0;
    lv = 0;
    n = 0;
    for s = 0 to nsym - 1;
      if f(s + 1) > 0;
        j = n;
        n += 1;
        dow j > 0;
          if f(lv(j) + 1) <= f(s + 1);
            leave;
          endif;
          lv(j + 1) = lv(j);
          j -= 1;
        enddo;
        lv(j + 1) = s;
      endif;
    endfor;
    if n = 0;
      return;
    endif;
    if n = 1;
      lens(lv(1) + 1) = 1;
      return;
    endif;
    li = 1;
    ii = 1;
    for k = 1 to n - 1;
      w = 0;
      for t = 1 to 2;
        takeLeaf = *off;
        if li <= n;
          if ii >= k;
            takeLeaf = *on;
          elseif f(lv(li) + 1) <= iw(ii);
            takeLeaf = *on;
          endif;
        endif;
        if takeLeaf;
          w += f(lv(li) + 1);
          parL(li) = k;
          li += 1;
        else;
          w += iw(ii);
          parI(ii) = k;
          ii += 1;
        endif;
      endfor;
      iw(k) = w;
    endfor;
    idep(n - 1) = 0;                 // root
    for k = n - 2 downto 1;
      idep(k) = idep(parI(k)) + 1;
    endfor;
    mx = 0;
    for j = 1 to n;
      d = idep(parL(j)) + 1;
      lens(lv(j) + 1) = d;
      if d > mx;
        mx = d;
      endif;
    endfor;
    if mx > limit;
      for s = 1 to nsym;
        if f(s) > 0;
          f(s) = %div(f(s) + 1 : 2);
        endif;
      endfor;
    endif;
  enddo;
end-proc;

// =====================================================================================
// output
// =====================================================================================
dcl-proc flushOut;
  dcl-pi *n ind;
  end-pi;
  dcl-s t0 timestamp;
  if obLen > 0;
    t0 = %timestamp();
    gSt.outB += obLen;
    if ifsWrite(gFd : pOB : obLen) <> obLen;
      xlSetErr('write ' + gPath + ': ' + sysErr());
      return *off;
    endif;
    filePos += obLen;
    obLen = 0;
    gSt.wrtUs += %diff(%timestamp() : t0 : *mseconds);
  endif;
  return *on;
end-proc;

dcl-proc putByte;
  dcl-pi *n;
    v int(10) value;
  end-pi;
  obLen += 1;
  OB(obLen) = v;
end-proc;

dcl-proc put16;
  dcl-pi *n;
    v int(10) value;
  end-pi;
  putByte(%rem(v : 256));
  putByte(%rem(%div(v : 256) : 256));
end-proc;

dcl-proc put32;
  dcl-pi *n;
    v uns(10) value;
  end-pi;
  putByte(%rem(v : 256));
  putByte(%rem(%div(v : 256) : 256));
  putByte(%rem(%div(v : 65536) : 256));
  putByte(%div(v : 16777216));
end-proc;

dcl-proc putName;
  dcl-pi *n;
    name varchar(512) ccsid(*utf8) const;
  end-pi;
  dcl-s nm varchar(512) ccsid(*utf8);
  nm = name;
  if %len(nm) > 0;
    memcpy(pOB + obLen : %addr(nm : *data) : %len(nm));
    obLen += %len(nm);
  endif;
end-proc;

dcl-proc putBits;
  dcl-pi *n;
    v uns(10) value;
    n int(10) value;
  end-pi;
  bitBuf += v * P2(bitCnt + 1);
  bitCnt += n;
  dow bitCnt >= 8;
    obLen += 1;
    OB(obLen) = %rem(bitBuf : 256);
    bitBuf = %div(bitBuf : 256);
    bitCnt -= 8;
  enddo;
end-proc;

// =====================================================================================
// archive
// =====================================================================================
dcl-proc zwOpen export;
  dcl-pi *n int(10);
    path  varchar(1024) const;
    level int(10) value;
  end-pi;
  dcl-s ts timestamp;

  if not tabsDone;
    zfInit();
  endif;
  if gFd >= 0;
    zwAbort();
  endif;
  if pIB = *null;
    pIB = %alloc(IBMAX + 16);
    pOB = %alloc(OBMAX + 16);
  endif;
  gPath = %trimr(path);
  gFd = ifsOpen(gPath : O_WRONLY + O_CREAT + O_TRUNC + O_CCSID + O_LARGEFILE
                : MODE_664 : 819);
  if gFd < 0;
    xlSetErr('open ' + gPath + ': ' + sysErr());
    return -1;
  endif;
  gLevel = level;
  if gLevel < 0 or gLevel > 9;
    gLevel = 2;
  endif;
  // levels 1-4 greedy, 5-9 lazy; lazy settings tuned with proto/lazy.py + grid2.py
  // (sheet XML of 200k x 10 ERP-like rows, rust_xlsxwriter test, source text)
  useLazy = (gLevel >= 5);
  goodLen = 8;
  lazyLen = 32;
  maxIns = 258;
  select;
  when gLevel <= 1;
    maxChain = 1;
    niceLen = 32;
    maxIns = 4;
  when gLevel = 2;
    maxChain = 4;
    niceLen = 32;
    maxIns = 4;
  when gLevel = 3;
    maxChain = 6;
    niceLen = 32;
    maxIns = 4;
  when gLevel = 4;
    maxChain = 16;
    niceLen = 128;
  when gLevel = 5;
    maxChain = 32;
    niceLen = 128;
    lazyLen = 16;
  when gLevel = 6;
    maxChain = 64;
    niceLen = 128;
  when gLevel = 7;
    maxChain = 256;
    niceLen = 128;
  when gLevel = 8;
    maxChain = 512;
    niceLen = 258;
  other;
    maxChain = 1024;
    niceLen = 258;
  endsl;
  ts = %timestamp();
  gDTime = %subdt(ts : *hours) * 2048 + %subdt(ts : *minutes) * 32
         + %div(%subdt(ts : *seconds) : 2);
  gDDate = (%subdt(ts : *years) - 1980) * 512 + %subdt(ts : *months) * 32
         + %subdt(ts : *days);
  filePos = 0;
  obLen = 0;
  nCd = 0;
  inEntry = *off;
  clear gSt;
  return 0;
end-proc;

dcl-proc localHeader;               // writes a local header, returns its offset
  dcl-pi *n int(20);
    name   varchar(512) ccsid(*utf8) const;
    method int(10) value;
    flags  int(10) value;
    dtime  int(10) value;
    ddate  int(10) value;
    crc    uns(10) value;
    csize  uns(10) value;
    usize  uns(10) value;
  end-pi;
  dcl-s off int(20);
  if obLen + 600 > OBMAX;
    if not flushOut();
      return -1;
    endif;
  endif;
  off = filePos + obLen;
  put32(67324752);                  // PK\3\4
  put16(20);
  put16(flags);
  put16(method);
  put16(dtime);
  put16(ddate);
  put32(crc);
  put32(csize);
  put32(usize);
  put16(%len(name));
  put16(0);
  putName(name);
  return off;
end-proc;

dcl-proc zwBegin export;
  dcl-pi *n int(10);
    name varchar(512) ccsid(*utf8) const;
  end-pi;
  if gFd < 0;
    xlSetErr('zwBegin: archive not open');
    return -1;
  endif;
  if inEntry;
    if zwEnd() < 0;
      return -1;
    endif;
  endif;
  if nCd >= XL_MAXENT;
    xlSetErr('too many zip entries');
    return -1;
  endif;
  if filePos + obLen > 4294000000;
    xlSetErr('archive larger than 4 GB (no ZIP64)');
    return -1;
  endif;
  if gLevel = 0;
    eMethod = 0;
  else;
    eMethod = 8;
  endif;
  eHdrOff = localHeader(name : eMethod : 2048 : gDTime : gDDate : 0 : 0 : 0);
  if eHdrOff < 0;
    return -1;
  endif;
  eDatOff = filePos + obLen;
  nCd += 1;
  cd(nCd).name = name;
  cd(nCd).method = eMethod;
  cd(nCd).flags = 2048;             // bit 11: names are UTF-8
  cd(nCd).dtime = gDTime;
  cd(nCd).ddate = gDDate;
  cd(nCd).offset = eHdrOff;
  eCrc = 0;
  eUsize = 0;
  ibLen = 0;
  ibPos = 0;
  crcPos = 0;
  gBase = 0;
  bitBuf = 0;
  bitCnt = 0;
  nSym = 0;
  FL = 0;
  FD = 0;
  HEAD = 0;
  PREV = 0;
  pvLen = 0;
  pvDist = 0;
  pvAvail = *off;
  inEntry = *on;
  return 0;
end-proc;

// add the not yet counted bytes of IB to CRC and size
dcl-proc crcNew;
  dcl-s t0 timestamp;
  if ibLen > crcPos;
    t0 = %timestamp();
    eCrc = xlCrc32(eCrc : pIB + crcPos : ibLen - crcPos);
    gSt.crcUs += %diff(%timestamp() : t0 : *mseconds);
    gSt.inB += ibLen - crcPos;
    eUsize += ibLen - crcPos;
    crcPos = ibLen;
  endif;
end-proc;

// make room in IB: compress (or store) what is there, keep the deflate window
dcl-proc drain;
  dcl-pi *n ind;
  end-pi;
  dcl-s sh int(10);
  dcl-s t0 timestamp;
  crcNew();
  if eMethod = 0;
    if not flushOut();
      return *off;
    endif;
    if ibLen > 0;
      if ifsWrite(gFd : pIB : ibLen) <> ibLen;
        xlSetErr('write ' + gPath + ': ' + sysErr());
        return *off;
      endif;
      filePos += ibLen;
    endif;
    ibLen = 0;
    ibPos = 0;
    crcPos = 0;
    return *on;
  endif;
  t0 = %timestamp();
  compress(*off);
  gSt.defUs += %diff(%timestamp() : t0 : *mseconds);
  if ibPos > HIST;
    sh = ibPos - HIST;
    memmove(pIB : pIB + sh : ibLen - sh);
    ibPos -= sh;
    ibLen -= sh;
    crcPos -= sh;
    gBase += sh;
  endif;
  return *on;
end-proc;

dcl-proc zwWrite export;
  dcl-pi *n int(10);
    data pointer value;
    len  int(10) value;
  end-pi;
  dcl-s p pointer;
  dcl-s k int(10);
  dcl-s room int(10);
  if not inEntry;
    xlSetErr('zwWrite: no entry open');
    return -1;
  endif;
  p = data;
  dow len > 0;
    room = HIST + IBCHUNK - ibLen;
    if room <= 0;
      if not drain();
        return -1;
      endif;
      iter;
    endif;
    k = len;
    if k > room;
      k = room;
    endif;
    memcpy(pIB + ibLen : p : k);
    ibLen += k;
    p += k;
    len -= k;
  enddo;
  return 0;
end-proc;

dcl-proc zwReserve export;
  dcl-pi *n pointer;
    wp   pointer value;
    need int(10) value;
    wEnd pointer;
  end-pi;
  if wp <> *null;
    ibLen = wp - pIB;
  endif;
  if ibLen >= HIST + IBCHUNK or ibLen + need > IBMAX;
    if not drain();
      wEnd = *null;
      return *null;
    endif;
  endif;
  wEnd = pIB + IBMAX;
  return pIB + ibLen;
end-proc;

dcl-proc zwCommit export;
  dcl-pi *n;
    wp pointer value;
  end-pi;
  if wp <> *null;
    ibLen = wp - pIB;
  endif;
end-proc;

dcl-proc zwEnd export;
  dcl-pi *n int(10);
  end-pi;
  dcl-s csize int(20);
  dcl-s t0 timestamp;
  if not inEntry;
    return 0;
  endif;
  inEntry = *off;
  crcNew();
  if eMethod = 0;
    if not drain();
      return -1;
    endif;
  else;
    t0 = %timestamp();
    compress(*on);
    flushBlock(*on);
    gSt.defUs += %diff(%timestamp() : t0 : *mseconds);
    if bitCnt > 0;
      putBits(0 : 8 - bitCnt);
    endif;
  endif;
  if not flushOut();
    return -1;
  endif;
  csize = filePos - eDatOff;
  if eUsize > 4294967295 or csize > 4294967295;
    xlSetErr('zip entry larger than 4 GB (no ZIP64)');
    return -1;
  endif;
  cd(nCd).crc = eCrc;
  cd(nCd).csize = csize;
  cd(nCd).usize = eUsize;
  // patch crc and sizes in the local header
  put32(eCrc);
  put32(csize);
  put32(eUsize);
  if ifsSeek(gFd : eHdrOff + 14 : SEEK_SET) < 0
     or ifsWrite(gFd : pOB : 12) <> 12
     or ifsSeek(gFd : 0 : SEEK_END) < 0;
    xlSetErr('patch local header ' + gPath + ': ' + sysErr());
    return -1;
  endif;
  obLen = 0;
  return 0;
end-proc;

dcl-proc zwRaw export;
  dcl-pi *n int(10);
    name   varchar(512) ccsid(*utf8) const;
    method int(10) value;
    flags  int(10) value;
    dtime  int(10) value;
    ddate  int(10) value;
    crc    uns(10) value;
    csize  uns(10) value;
    usize  uns(10) value;
    data   pointer value;
  end-pi;
  dcl-s off int(20);
  dcl-s p pointer;
  dcl-s left int(20);
  dcl-s k int(10);
  dcl-s fl int(10);
  if inEntry;
    if zwEnd() < 0;
      return -1;
    endif;
  endif;
  if nCd >= XL_MAXENT;
    xlSetErr('too many zip entries');
    return -1;
  endif;
  fl = flags;
  if %bitand(fl : 8) <> 0;           // no data descriptor in the copy
    fl -= 8;
  endif;
  off = localHeader(name : method : fl : dtime : ddate : crc : csize : usize);
  if off < 0 or not flushOut();
    return -1;
  endif;
  p = data;
  left = csize;
  dow left > 0;
    k = 16000000;
    if left < k;
      k = left;
    endif;
    if ifsWrite(gFd : p : k) <> k;
      xlSetErr('write ' + gPath + ': ' + sysErr());
      return -1;
    endif;
    filePos += k;
    p += k;
    left -= k;
  enddo;
  nCd += 1;
  cd(nCd).name = name;
  cd(nCd).method = method;
  cd(nCd).flags = fl;
  cd(nCd).dtime = dtime;
  cd(nCd).ddate = ddate;
  cd(nCd).crc = crc;
  cd(nCd).csize = csize;
  cd(nCd).usize = usize;
  cd(nCd).offset = off;
  return 0;
end-proc;

dcl-proc zwClose export;
  dcl-pi *n int(10);
  end-pi;
  dcl-s i int(10);
  dcl-s cdOff int(20);
  dcl-s cdLen int(20);
  if gFd < 0;
    return 0;
  endif;
  if inEntry;
    if zwEnd() < 0;
      return -1;
    endif;
  endif;
  if not flushOut();
    return -1;
  endif;
  cdOff = filePos;
  for i = 1 to nCd;
    if obLen + 600 > OBMAX;
      if not flushOut();
        return -1;
      endif;
    endif;
    put32(33639248);                // PK\1\2
    put16(20);                      // made by
    put16(20);                      // needed
    put16(cd(i).flags);
    put16(cd(i).method);
    put16(cd(i).dtime);
    put16(cd(i).ddate);
    put32(cd(i).crc);
    put32(cd(i).csize);
    put32(cd(i).usize);
    put16(%len(cd(i).name));
    put16(0);                       // extra
    put16(0);                       // comment
    put16(0);                       // disk
    put16(0);                       // internal attributes
    put32(0);                       // external attributes
    put32(cd(i).offset);
    putName(cd(i).name);
  endfor;
  cdLen = filePos + obLen - cdOff;
  put32(101010256);                 // PK\5\6
  put16(0);
  put16(0);
  put16(nCd);
  put16(nCd);
  put32(cdLen);
  put32(cdOff);
  put16(0);
  if not flushOut();
    return -1;
  endif;
  if ifsClose(gFd) <> 0;
    gFd = -1;
    xlSetErr('close ' + gPath + ': ' + sysErr());
    return -1;
  endif;
  gFd = -1;
  return 0;
end-proc;

dcl-proc zwAbort export;
  if gFd >= 0;
    ifsClose(gFd);
    gFd = -1;
    ifsUnlink(gPath);
  endif;
  inEntry = *off;
  nCd = 0;
end-proc;

// =====================================================================================
// deflate: LZ77 into the symbol buffer
// =====================================================================================
dcl-proc compress;
  dcl-pi *n;
    final ind value;
  end-pi;
  dcl-s i     int(10);
  dcl-s lim   int(10);
  dcl-s h     int(10);
  dcl-s cand  int(20);
  dcl-s gp    int(20);
  dcl-s cpos  int(20);
  dcl-s dist  int(20);
  dcl-s ci    int(10);
  dcl-s best  int(10);
  dcl-s bdist int(10);
  dcl-s maxLen int(10);
  dcl-s depth int(10);
  dcl-s l     int(10);
  dcl-s k     int(10);
  dcl-s j     int(10);
  dcl-s gj    int(20);
  dcl-s pa    pointer;
  dcl-s pb    pointer;
  dcl-s c8a   char(8) ccsid(*hex) based(pa);
  dcl-s c8b   char(8) ccsid(*hex) based(pb);

  if useLazy;
    compLazy(final);
    return;
  endif;
  if final;
    lim = ibLen;
  else;
    lim = ibLen - MINLOOK;
  endif;
  i = ibPos;
  dow i < lim;
    best = 0;
    bdist = 0;
    if i + 2 < ibLen;
      h = %bitand(IB(i + 1) * 1024 + IB(i + 2) * 32 + IB(i + 3) : 32767) + 1;
      gp = gBase + i;
      cand = HEAD(h);
      HEAD(h) = gp + 1;
      PREV(%bitand(gp : 32767) + 1) = cand;
      maxLen = ibLen - i;
      if maxLen > 258;
        maxLen = 258;
      endif;
      depth = maxChain;
      dow cand > 0 and depth > 0;
        cpos = cand - 1;
        dist = gp - cpos;
        if dist > 32768;
          leave;
        endif;
        ci = cpos - gBase;
        if IB(ci + best + 1) = IB(i + best + 1) and IB(ci + 1) = IB(i + 1)
           and IB(ci + 2) = IB(i + 2);
          l = 2;
          pa = pIB + ci + 2;
          pb = pIB + i + 2;
          dow l + 8 <= maxLen;
            if c8a <> c8b;
              leave;
            endif;
            l += 8;
            pa += 8;
            pb += 8;
          enddo;
          dow l < maxLen;
            if IB(ci + l + 1) <> IB(i + l + 1);
              leave;
            endif;
            l += 1;
          enddo;
          if l > best;
            best = l;
            bdist = dist;
            if l >= niceLen or l >= maxLen;
              leave;
            endif;
          endif;
        endif;
        cand = PREV(%bitand(cpos : 32767) + 1);
        depth -= 1;
      enddo;
      if best = 3 and bdist > 4096;
        best = 0;
      endif;
    endif;
    if best >= 3;
      nSym += 1;
      SLIT(nSym) = best;
      SDIST(nSym) = bdist;
      FL(LSYM(best + 1) + 1) += 1;
      if bdist <= 256;
        FD(DC1(bdist) + 1) += 1;
      else;
        FD(DC2(%div(bdist - 1 : 128) + 1) + 1) += 1;
      endif;
      if best <= maxIns;
        for k = 1 to best - 1;
          j = i + k;
          if j + 2 < ibLen;
            h = %bitand(IB(j + 1) * 1024 + IB(j + 2) * 32 + IB(j + 3) : 32767) + 1;
            gj = gBase + j;
            PREV(%bitand(gj : 32767) + 1) = HEAD(h);
            HEAD(h) = gj + 1;
          endif;
        endfor;
      endif;
      i += best;
    else;
      nSym += 1;
      SLIT(nSym) = IB(i + 1);
      SDIST(nSym) = 0;
      FL(IB(i + 1) + 1) += 1;
      i += 1;
    endif;
    if nSym >= BLK;
      flushBlock(*off);
    endif;
  enddo;
  ibPos = i;
end-proc;

// =====================================================================================
// deflate: LZ77 with lazy matching (levels 5-9, same decisions as zlib deflate_slow).
// A match found at position i is kept pending; at i+1 the chain is searched again for
// a longer one. If none is found the pending match is emitted, otherwise the byte at i
// becomes a literal and the new match is pending. State (pvLen, pvDist, pvAvail) lives
// across calls, so input chunk boundaries do not change the result.
// =====================================================================================
dcl-proc compLazy;
  dcl-pi *n;
    final ind value;
  end-pi;
  dcl-s i     int(10);
  dcl-s lim   int(10);
  dcl-s h     int(10);
  dcl-s cand  int(20);
  dcl-s gp    int(20);
  dcl-s cpos  int(20);
  dcl-s dist  int(20);
  dcl-s ci    int(10);
  dcl-s best  int(10);
  dcl-s bdist int(10);
  dcl-s maxLen int(10);
  dcl-s depth int(10);
  dcl-s l     int(10);
  dcl-s j     int(10);
  dcl-s e     int(10);
  dcl-s gj    int(20);
  dcl-s found ind;
  dcl-s pa    pointer;
  dcl-s pb    pointer;
  dcl-s c8a   char(8) ccsid(*hex) based(pa);
  dcl-s c8b   char(8) ccsid(*hex) based(pb);

  if final;
    lim = ibLen;
  else;
    lim = ibLen - MINLOOK;
  endif;
  i = ibPos;
  dow i < lim;
    best = 0;
    bdist = 0;
    if i + 2 < ibLen;
      h = %bitand(IB(i + 1) * 1024 + IB(i + 2) * 32 + IB(i + 3) : 32767) + 1;
      gp = gBase + i;
      cand = HEAD(h);
      HEAD(h) = gp + 1;
      PREV(%bitand(gp : 32767) + 1) = cand;
      maxLen = ibLen - i;
      if maxLen > 258;
        maxLen = 258;
      endif;
      if cand > 0 and pvLen < lazyLen and pvLen < maxLen;
        depth = maxChain;
        if pvLen >= goodLen;
          depth = %div(depth : 4);
        endif;
        if pvLen >= 3;
          best = pvLen;                 // only a longer match is of interest
        else;
          best = 2;
        endif;
        found = *off;
        dow cand > 0 and depth > 0;
          cpos = cand - 1;
          dist = gp - cpos;
          if dist > 32768;
            leave;
          endif;
          ci = cpos - gBase;
          if IB(ci + best + 1) = IB(i + best + 1) and IB(ci + 1) = IB(i + 1)
             and IB(ci + 2) = IB(i + 2);
            l = 2;
            pa = pIB + ci + 2;
            pb = pIB + i + 2;
            dow l + 8 <= maxLen;
              if c8a <> c8b;
                leave;
              endif;
              l += 8;
              pa += 8;
              pb += 8;
            enddo;
            dow l < maxLen;
              if IB(ci + l + 1) <> IB(i + l + 1);
                leave;
              endif;
              l += 1;
            enddo;
            if l > best;
              best = l;
              bdist = dist;
              found = *on;
              if l >= niceLen or l >= maxLen;
                leave;
              endif;
            endif;
          endif;
          cand = PREV(%bitand(cpos : 32767) + 1);
          depth -= 1;
        enddo;
        if not found or (best = 3 and bdist > 4096);
          best = 0;
          bdist = 0;
        endif;
      endif;
    endif;

    if pvLen >= 3 and best <= pvLen;
      // emit the pending match (it starts at i - 1), hash the positions it covers
      nSym += 1;
      SLIT(nSym) = pvLen;
      SDIST(nSym) = pvDist;
      FL(LSYM(pvLen + 1) + 1) += 1;
      if pvDist <= 256;
        FD(DC1(pvDist) + 1) += 1;
      else;
        FD(DC2(%div(pvDist - 1 : 128) + 1) + 1) += 1;
      endif;
      e = i - 1 + pvLen;
      for j = i + 1 to e - 1;
        if j + 2 < ibLen;
          h = %bitand(IB(j + 1) * 1024 + IB(j + 2) * 32 + IB(j + 3) : 32767) + 1;
          gj = gBase + j;
          PREV(%bitand(gj : 32767) + 1) = HEAD(h);
          HEAD(h) = gj + 1;
        endif;
      endfor;
      i = e;
      pvAvail = *off;
      pvLen = 0;
      pvDist = 0;
    else;
      if pvAvail;                       // the byte at i - 1 becomes a literal
        nSym += 1;
        SLIT(nSym) = IB(i);
        SDIST(nSym) = 0;
        FL(IB(i) + 1) += 1;
      endif;
      pvAvail = *on;
      pvLen = best;
      pvDist = bdist;
      i += 1;
    endif;
    if nSym >= BLK;
      flushBlock(*off);
    endif;
  enddo;

  if final and pvAvail;               // last pending position
    nSym += 1;
    if pvLen >= 3;
      SLIT(nSym) = pvLen;
      SDIST(nSym) = pvDist;
      FL(LSYM(pvLen + 1) + 1) += 1;
      if pvDist <= 256;
        FD(DC1(pvDist) + 1) += 1;
      else;
        FD(DC2(%div(pvDist - 1 : 128) + 1) + 1) += 1;
      endif;
    else;
      SLIT(nSym) = IB(i);
      SDIST(nSym) = 0;
      FL(IB(i) + 1) += 1;
    endif;
    pvAvail = *off;
    pvLen = 0;
    pvDist = 0;
  endif;
  ibPos = i;
end-proc;

// =====================================================================================
// deflate: emit one block (fixed or dynamic Huffman, whichever is smaller)
// =====================================================================================
dcl-proc flushBlock;
  dcl-pi *n;
    last ind value;
  end-pi;
  dcl-s lL  int(10) dim(288);
  dcl-s lD  int(10) dim(288);
  dcl-s lC  int(10) dim(288);
  dcl-s fC  int(10) dim(288);
  dcl-s cL  uns(10) dim(288);
  dcl-s cD  uns(10) dim(288);
  dcl-s cC  uns(10) dim(288);
  dcl-s sq  int(10) dim(320);        // code lengths to transmit
  dcl-s rl  int(10) dim(320);        // run-length encoded: symbol
  dcl-s rx  int(10) dim(320);        //                     extra value
  dcl-s nr  int(10);
  dcl-s ns  int(10);
  dcl-s nl  int(10);
  dcl-s nd  int(10);
  dcl-s hlit int(10);
  dcl-s hdist int(10);
  dcl-s hclen int(10);
  dcl-s i int(10);
  dcl-s k int(10);
  dcl-s r int(10);
  dcl-s v int(10);
  dcl-s sy int(10);
  dcl-s c int(10);
  dcl-s d int(10);
  dcl-s l int(10);
  dcl-s ex int(10);
  dcl-s dyn int(20);
  dcl-s fix int(20);
  dcl-s useDyn ind;

  FL(257) = 1;                       // end of block
  nl = 0;
  nd = 0;
  for i = 1 to 286;
    if FL(i) > 0;
      nl += 1;
    endif;
  endfor;
  for i = 1 to 30;
    if FD(i) > 0;
      nd += 1;
    endif;
  endfor;
  if nl < 2;
    FL(1) = 1;
  endif;
  if FD(1) = 0 and nd < 2;
    FD(1) = 1;
    nd += 1;
  endif;
  if nd < 2;
    FD(2) = 1;
  endif;

  hufLens(FL : lL : 286 : 15);
  hufLens(FD : lD : 30 : 15);
  hlit = 286;
  dow hlit > 257;
    if lL(hlit) <> 0;
      leave;
    endif;
    hlit -= 1;
  enddo;
  hdist = 30;
  dow hdist > 1;
    if lD(hdist) <> 0;
      leave;
    endif;
    hdist -= 1;
  enddo;

  // run-length encode the code lengths
  ns = 0;
  for i = 1 to hlit;
    ns += 1;
    sq(ns) = lL(i);
  endfor;
  for i = 1 to hdist;
    ns += 1;
    sq(ns) = lD(i);
  endfor;
  fC = 0;
  nr = 0;
  i = 1;
  dow i <= ns;
    v = sq(i);
    r = 1;
    dow i + r <= ns;
      if sq(i + r) <> v;
        leave;
      endif;
      r += 1;
    enddo;
    if v = 0 and r >= 3;
      if r > 138;
        r = 138;
      endif;
      nr += 1;
      if r >= 11;
        rl(nr) = 18;
        rx(nr) = r - 11;
        fC(19) += 1;
      else;
        rl(nr) = 17;
        rx(nr) = r - 3;
        fC(18) += 1;
      endif;
      i += r;
      iter;
    endif;
    nr += 1;
    rl(nr) = v;
    rx(nr) = 0;
    fC(v + 1) += 1;
    i += 1;
    r -= 1;
    if v <> 0;
      dow r >= 3;
        k = r;
        if k > 6;
          k = 6;
        endif;
        nr += 1;
        rl(nr) = 16;
        rx(nr) = k - 3;
        fC(17) += 1;
        i += k;
        r -= k;
      enddo;
    endif;
  enddo;
  hufLens(fC : lC : 19 : 7);
  hclen = 19;
  dow hclen > 4;
    if lC(CLORD(hclen) + 1) <> 0;
      leave;
    endif;
    hclen -= 1;
  enddo;

  // cost of both encodings, in bits
  dyn = 17 + 3 * hclen;
  for k = 1 to nr;
    sy = rl(k);
    dyn += lC(sy + 1);
    if sy = 16;
      dyn += 2;
    elseif sy = 17;
      dyn += 3;
    elseif sy = 18;
      dyn += 7;
    endif;
  endfor;
  fix = 3;
  for sy = 0 to 285;
    if FL(sy + 1) > 0;
      ex = 0;
      if sy >= 257;
        ex = LEXT(sy - 256);
      endif;
      dyn += FL(sy + 1) * (lL(sy + 1) + ex);
      fix += FL(sy + 1) * (FIXL(sy + 1) + ex);
    endif;
  endfor;
  for c = 1 to 30;
    if FD(c) > 0;
      dyn += FD(c) * (lD(c) + DEXT(c));
      fix += FD(c) * (5 + DEXT(c));
    endif;
  endfor;
  useDyn = (dyn < fix);

  if useDyn;
    hufCodes(lL : cL : 286);
    hufCodes(lD : cD : 30);
    hufCodes(lC : cC : 19);
    if last;
      putBits(1 : 1);
    else;
      putBits(0 : 1);
    endif;
    putBits(2 : 2);
    putBits(hlit - 257 : 5);
    putBits(hdist - 1 : 5);
    putBits(hclen - 4 : 4);
    for k = 1 to hclen;
      putBits(lC(CLORD(k) + 1) : 3);
    endfor;
    for k = 1 to nr;
      sy = rl(k);
      putBits(cC(sy + 1) : lC(sy + 1));
      if sy = 16;
        putBits(rx(k) : 2);
      elseif sy = 17;
        putBits(rx(k) : 3);
      elseif sy = 18;
        putBits(rx(k) : 7);
      endif;
    endfor;
  else;
    lL = FIXL;
    cL = FIXC;
    lD = FIXDL;
    cD = FIXDC;
    if last;
      putBits(1 : 1);
    else;
      putBits(0 : 1);
    endif;
    putBits(1 : 2);
  endif;
  if obLen > OBFLUSH;
    flushOut();
  endif;

  // the symbols (bit writer inlined, 32 bits flushed at a time: this is the hot loop)
  for k = 1 to nSym;
    d = SDIST(k);
    if d = 0;
      sy = SLIT(k) + 1;
      bitBuf += cL(sy) * P2(bitCnt + 1);
      bitCnt += lL(sy);
    else;
      l = SLIT(k) + 1;
      sy = LSYM(l) + 1;
      bitBuf += (cL(sy) + LEV(l) * P2(lL(sy) + 1)) * P2(bitCnt + 1);
      bitCnt += lL(sy) + LEX(l);
      if bitCnt >= 32;
        bw = bitBuf;
        OB(obLen + 1) = bb8;
        OB(obLen + 2) = bb7;
        OB(obLen + 3) = bb6;
        OB(obLen + 4) = bb5;
        obLen += 4;
        bitBuf = %div(bitBuf : 4294967296);
        bitCnt -= 32;
      endif;
      if d <= 256;
        c = DC1(d) + 1;
      else;
        c = DC2(%div(d - 1 : 128) + 1) + 1;
      endif;
      bitBuf += (cD(c) + (d - DBASE(c)) * P2(lD(c) + 1)) * P2(bitCnt + 1);
      bitCnt += lD(c) + DEXT(c);
    endif;
    if bitCnt >= 32;
      bw = bitBuf;
      OB(obLen + 1) = bb8;
      OB(obLen + 2) = bb7;
      OB(obLen + 3) = bb6;
      OB(obLen + 4) = bb5;
      obLen += 4;
      bitBuf = %div(bitBuf : 4294967296);
      bitCnt -= 32;
      if obLen > OBFLUSH;
        flushOut();
      endif;
    endif;
  endfor;
  putBits(cL(257) : lL(257));        // end of block

  nSym = 0;
  FL = 0;
  FD = 0;
end-proc;
