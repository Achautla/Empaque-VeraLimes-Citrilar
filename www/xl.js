/* Generadores de Excel (ExcelJS). Reciben datos ya calculados y escriben hojas con formato, logotipo y fórmulas vivas.
   Los tres formatos: registro (Formato 1), trazabilidad (Formato 2) y base de datos decorada. */
(function (root) {
  'use strict';
  var XL = {};
  var FN = 'Calibri';
  var NUM = '#,##0;-#,##0;"-"';
  var NUM1 = '#,##0.0;-#,##0.0;"-"';
  var thin = { style: 'thin', color: { argb: 'FF000000' } };
  var hair = { style: 'thin', color: { argb: 'FFB7B7B7' } };
  var med = { style: 'medium', color: { argb: 'FF000000' } };
  var BOX = { top: thin, left: thin, bottom: thin, right: thin };
  var HBOX = { top: hair, left: hair, bottom: hair, right: hair };
  var CAL = [
    { k: 'ctn', l: '230 CTN' }, { k: 'c250', l: '250' }, { k: 'c230', l: '230' }, { k: 'c200', l: '200' },
    { k: 'c175', l: '175' }, { k: 'c150', l: '150' }, { k: 'c110', l: '110' }
  ];
  function fill(rgb) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + rgb } }; }
  function font(o) { return Object.assign({ name: FN, size: 10 }, o || {}); }
  function colN(L) { var n = 0; for (var i = 0; i < L.length; i++) n = n * 26 + (L.charCodeAt(i) - 64); return n; }
  /* v8.28: recorre S columnas las referencias de esta hoja de la columna C en adelante; las de otras hojas (Hoja!A1:B2) no se tocan */
  function shiftF(v, S) {
    if (typeof v !== 'string' || v.charAt(0) !== '=') return v;
    var keep = [];
    var f = v.replace(/('(?:[^']|'')+'|[A-Za-z0-9_]+)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/g, function (m) { keep.push(m); return '\u0001' + (keep.length - 1) + '\u0002'; });
    f = f.replace(/(^|[^A-Za-z0-9_$\u0001])(\$?)([A-Z]{1,3})(\$?)(\d+)/g, function (m, pre, d1, L, d2, num) { var n = colN(L); return pre + d1 + colL(n >= 3 ? n + S : n) + d2 + num; });
    return f.replace(/\u0001(\d+)\u0002/g, function (m, i) { return keep[+i]; });
  }
  function colL(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function A(c, r, abs) { return (abs ? '$' : '') + colL(c) + (abs ? '$' : '') + r; }
  function put(ws, r, c, v, st) {
    var cell = ws.getCell(r, c);
    if (v !== undefined && v !== null && v !== '') cell.value = v;
    if (st) {
      if (st.font) cell.font = st.font;
      if (st.fill) cell.fill = st.fill;
      if (st.border) cell.border = st.border;
      if (st.al) cell.alignment = st.al;
      if (st.fmt) cell.numFmt = st.fmt;
    }
    return cell;
  }
  function fx(f, res) { return { formula: f, result: res === undefined ? 0 : res }; }
  function merge(ws, r1, c1, r2, c2) { if (r1 !== r2 || c1 !== c2) ws.mergeCells(r1, c1, r2, c2); }
  function border(ws, r1, c1, r2, c2, b) {
    for (var r = r1; r <= r2; r++) for (var c = c1; c <= c2; c++) ws.getCell(r, c).border = b || BOX;
  }
  function fillRange(ws, r1, c1, r2, c2, f) {
    for (var r = r1; r <= r2; r++) for (var c = c1; c <= c2; c++) ws.getCell(r, c).fill = f;
  }
  var CEN = { horizontal: 'center', vertical: 'middle', wrapText: true };
  var LEFT = { horizontal: 'left', vertical: 'middle' };
  /* posición exacta en puntos: suma de altos de fila (Excel guarda el ancla en EMU, 12700 por punto) */
  function rowTop(ws, r) { var t = 0; for (var i = 1; i < r; i++) t += ws.getRow(i).height || 15; return t; }
  function imgAt(wb, ws, pic, col, topPt, wpx, hpx) {
    var id = wb.addImage({ base64: pic.b64, extension: 'jpeg' }), r = 1, acc = 0;
    topPt = Math.max(0, topPt);
    while (acc + (ws.getRow(r).height || 15) <= topPt) { acc += ws.getRow(r).height || 15; r++; }
    ws.addImage(id, { tl: { nativeCol: col, nativeColOff: 0, nativeRow: r - 1, nativeRowOff: Math.round((topPt - acc) * 12700) }, ext: { width: Math.round(wpx), height: Math.round(hpx) }, editAs: 'oneCell' });
  }
  function img(wb, ws, logo, col, row, wpx, hmax) {
    if (!logo || !logo.b64) return;
    var id = wb.addImage({ base64: logo.b64, extension: 'jpeg' });
    var h = Math.round(wpx * (logo.h / logo.w));
    if (hmax && h > hmax) { wpx = Math.round(wpx * hmax / h); h = hmax; }
    ws.addImage(id, { tl: { col: col, row: row }, ext: { width: wpx, height: h } });
  }
  function safeName(wb, n) {
    n = String(n || 'Hoja').replace(/[\\\/\?\*\[\]:]/g, '-').slice(0, 31);
    var base = n, i = 2;
    while (wb.getWorksheet(n)) { n = base.slice(0, 27) + ' (' + i + ')'; i++; }
    return n;
  }
  function dmy(s) { if (!s) return ''; var p = String(s).split('-'); return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : s; }
  function usd(s) { if (!s) return ''; var p = String(s).split('-'); return p.length === 3 ? p[1] + '/' + p[2] + '/' + p[0] : s; }
  function numOrStr(v) { var n = parseFloat(v); return (isFinite(n) && String(n) === String(v).trim()) ? n : (v == null ? '' : v); }

  /* ============ FORMATO 1: REGISTRO DE EMPAQUE ============ */
  /* d: {o, M (resultado de compute), logo:{b64,w,h}, sheetName, estado:'proceso'|'completo', tituloExtra, sobrante} */
  XL.addRegistro = function (wb, d) {
    var o = d.o, M = d.M, n = M.rows.length, kg = M.kg;
    /* calibres de la hoja: sin la fila 230 CTN salvo en pedidos viejos que la tienen capturada */
    var CALX = d.cal || CAL, nC = CALX.length, hasCtn = CALX.some(function (c) { return c.k === 'ctn'; });
    var cT = 8 + nC, cSg = cT + 1, cTr = cT + 2, cDs = cT + 3, cJg = cT + 4, cRn = cT + 5, cDf = cT + 6, cTp = cT + 7;
    /* v8.14: peso por caja de cada tipo de empaque. Columnas al final: cajas de cada tipo que lleva la hoja y KILOS EMPACADOS;
       el rendimiento es kilos empacados / kilos recibidos y la diferencia resta los kilos empacados convertidos a cajas campo */
    var WK = d.pesosK || ['RPC', 'CTN', 'CTN10', 'PLT'], WD = { RPC: 17.8, CTN: 16, CTN10: 4.53, PLT: 17.8 }, W = d.pesos || M.W || WD;
    var WN = d.tipoCh || d.tipoNames || {};
    var WU = WK.filter(function (t) { return M.rows.some(function (x) { return x.byW && x.byW[t] > 0; }); });
    if (!WU.length) WU = [(M.tipos && M.tipos[0]) || 'RPC'];
    var cW0 = cTp + 1, cKE = cW0 + WU.length, cEnd = cKE;
    var ws = wb.addWorksheet(safeName(wb, d.sheetName || o.cod), {
      views: [{ showGridLines: false }],
      pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 1, margins: { left: .3, right: .3, top: .4, bottom: .4, header: .2, footer: .2 } }
    });
    ws.properties.tabColor = { argb: d.estado === 'completo' ? 'FF00B050' : 'FFF08030' };
    var lastCol = Math.max(cEnd, 5 + n);
    var widths = { 1: 3, 2: 27, 3: 12 };
    for (var c = 1; c <= lastCol + 1; c++) ws.getColumn(c).width = widths[c] || 11;
    ws.getColumn(2).width = 27;
    [26, 22, 21, 21, 21, 8, 24, 34].forEach(function (h, i) { ws.getRow(i + 1).height = h; }); /* v8.18.4: el renglón de DATE ya cabe la fecha (antes 15) */

    /* encabezado */
    img(wb, ws, d.logo, 1, 0, 150, 100);
    put(ws, 2, 16, 'DATE', { font: font({ size: 13, bold: true }), al: { horizontal: 'right', vertical: 'middle' } });
    merge(ws, 2, 17, 2, 18);
    put(ws, 2, 17, d.fechaDate || usd(o.fecha), { font: font({ size: 14, bold: true }), fill: fill('FFEB9C'), al: CEN, border: BOX });
    ws.getCell(2, 18).border = BOX; ws.getCell(2, 18).fill = fill('FFEB9C');
    put(ws, 3, 4, 'CUSTOMER:', { font: font({ size: 12, bold: true }) });
    merge(ws, 3, 6, 3, 11); put(ws, 3, 6, o.customer || '', { font: font({ size: 12, bold: true }), al: LEFT });
    put(ws, 4, 4, 'Domestic Customer:', { font: font({ size: 12, bold: true }) });
    merge(ws, 4, 6, 4, 11); put(ws, 4, 6, o.domestic || '', { font: font({ size: 12, bold: true }), al: LEFT });
    put(ws, 3, 14, 'Factura', { font: font({ size: 12, bold: true }) });
    merge(ws, 3, 15, 3, 16); put(ws, 3, 15, o.factura || '', { font: font({ size: 12, bold: true }), al: LEFT });
    put(ws, 4, 14, 'PO', { font: font({ size: 12, bold: true }) });
    merge(ws, 4, 15, 4, 16); put(ws, 4, 15, numOrStr(o.po || ''), { font: font({ size: 12, bold: true }), al: LEFT });
    /* v8.8: el kg por caja sigue en la celda O5 para las fórmulas, pero no se ve en la hoja ni al imprimir */
    merge(ws, 5, 15, 5, 16);
    put(ws, 5, 15, kg, { font: font({ size: 11, color: { argb: 'FFFFFFFF' } }), al: CEN, fmt: ';;;' });
    var KG = A(15, 5, true);
    put(ws, 3, 18, 'SEMANA', { font: font({ size: 13, bold: true }), al: { horizontal: 'right', vertical: 'middle' } });
    put(ws, 3, 19, numOrStr(o.semana), { font: font({ size: 13, bold: true }), fill: fill('F4B183'), al: CEN, border: BOX });
    put(ws, 4, 18, o.cod || '', { font: font({ size: 13, bold: true }), al: { horizontal: 'right', vertical: 'middle' } });
    put(ws, 4, 19, o.emb || '', { font: font({ size: 11, bold: true }), al: LEFT });
    var enProceso = d.estado !== 'completo';
    merge(ws, 5, 18, 5, 20);
    put(ws, 5, 18, enProceso ? 'EN PROCESO' : 'COMPLETO', { font: font({ size: 11, bold: true, color: { argb: enProceso ? 'FF9C5700' : 'FF006100' } }), fill: fill(enProceso ? 'FFEB9C' : 'C6EFCE'), al: CEN, border: BOX });

    put(ws, 5, 4, 'TIPO DE EMPAQUE:', { font: font({ size: 12, bold: true }) });
    merge(ws, 5, 6, 5, 11); put(ws, 5, 6, d.tipoTxt || 'RPC', { font: font({ size: 12, bold: true }), al: LEFT });

    /* tabla superior: acumulado por calibre */
    var R_T = 7, R_H = 8, R0 = 9, R_TOT = R0 + nC, R_INC = R0 + nC + 1;
    var cP = 3, cF = 3 + n, cM = cF + 1, cS = cF + 2; /* si n=0 las columnas de resumen empiezan en 3 */
    var title = (d.titulo || o.cod) + (o.variante === '1.5' ? '  ·  1.5 (segunda)' : '');
    /* v6: manifiesto por tipo de empaque. Se coloca debajo; aquí se calcula su posición para las fórmulas */
    var LN = d.lines || [], R_MT = 0, R_M0 = 0;
    /* renglón 023 CITRILAR SEGUNDAS: debajo van los productores que aportan sus segundas, un renglón cada uno */
    var APS = M.rows.map(function (x) { return (x.p.aportes || []).filter(function (a) { return +a.n > 0; }); });
    /* v8.9.1: renglón SOBRANTES: en sus notas va, en texto, cada productor con código, nombre, cajas, calibres y tarima de origen */
    var NS = M.rows.map(function (x) { return (d.notasP && d.notasP[x.p.id]) || []; });
    var RS = [], nRows = 0; M.rows.forEach(function (x, i) { RS.push(nRows); nRows += 2 + APS[i].length + NS[i].length; });
    var anyAp = APS.some(function (a) { return a.length > 0; });
    var rowsN = M.rows.length, R_RT = (R0 + nC + 1 + 3 + 1) + nRows;
    var R_RL = R_RT + 5;
    if (LN.length) { R_MT = R_RL + (d.sobrante && d.sobrante.total > 0 ? 4 : 3); R_M0 = R_MT + 3; }
    function manCells(ci) { var a = []; LN.forEach(function (x, li) { a.push(A(3 + li * 3, R_M0 + ci)); }); return a; }
    var tcEnd = n > 0 ? 2 + n : 3;
    /* el título abarca toda la tabla de calibres (con un solo productor no cabía en una columna) */
    var tEnd = Math.max(tcEnd, cS);
    merge(ws, R_T, 3, R_T, tEnd);
    put(ws, R_T, 3, title, { font: font({ size: 20, bold: true }), fill: fill('D5FC79'), al: CEN, border: BOX });
    fillRange(ws, R_T, 3, R_T, tEnd, fill('D5FC79'));
    put(ws, R_T, 2, '', { fill: fill('D5FC79'), border: BOX });
    ws.getRow(R_T).height = 30;
    put(ws, R_H, 2, 'CALIBRE', { font: font({ size: 10, bold: true }), fill: fill('D5FC79'), al: CEN, border: BOX });
    M.rows.forEach(function (r, i) {
      put(ws, R_H, cP + i, r.p.nombre || ('Productor ' + (i + 1)), { font: font({ size: 8, bold: true }), fill: fill('D5FC79'), al: CEN, border: BOX });
    });
    put(ws, R_H, cF, 'ACUMULADO FINAL', { font: font({ size: 8, bold: true }), fill: fill('D5FC79'), al: CEN, border: BOX });
    put(ws, R_H, cM, 'MANIFIESTO', { font: font({ size: 9, bold: true }), fill: fill('65FB27'), al: CEN, border: BOX });
    put(ws, R_H, cS, 'SOBRANTES', { font: font({ size: 9, bold: true }), fill: fill('F4B183'), al: CEN, border: BOX });
    if (n === 0) { /* sin productores: el acumulado final queda al lado */ }
    ws.getRow(R_H).height = 36;
    CALX.forEach(function (c, ci) {
      var r = R0 + ci;
      put(ws, r, 2, c.l, { font: font({ size: 11, bold: true }), fill: c.k === 'ctn' ? fill('FFFF00') : undefined, al: CEN, border: BOX });
      M.rows.forEach(function (x, i) {
        put(ws, r, cP + i, x.eff[c.k], { font: font({ size: 11 }), fill: fill('D9D9D9'), al: CEN, border: BOX, fmt: NUM });
      });
      put(ws, r, cF, n > 0 ? fx(A(cP + n - 1, r), M.final[c.k]) : 0, { font: font({ size: 11, bold: true }), fill: fill('EDEDED'), al: CEN, border: BOX, fmt: NUM });
      if (LN.length) {
        var mc = manCells(ci).join(',');
        put(ws, r, cM, fx('IF(COUNT(' + mc + ')=0,"",SUM(' + mc + '))', M.man[c.k] == null ? '' : M.man[c.k]), { font: font({ size: 11, bold: true, color: { argb: 'FF0000FF' } }), fill: fill('CCECFF'), al: CEN, border: BOX, fmt: NUM });
      } else put(ws, r, cM, M.man[c.k] == null ? null : M.man[c.k], { font: font({ size: 11, bold: true, color: { argb: 'FF0000FF' } }), fill: fill('CCECFF'), al: CEN, border: BOX, fmt: NUM });
      var sob = M.man[c.k] == null ? '' : M.final[c.k] - M.man[c.k];
      put(ws, r, cS, fx('IF(ISNUMBER(' + A(cM, r) + '),' + A(cF, r) + '-' + A(cM, r) + ',"")', sob), { font: font({ size: 11, bold: true }), fill: fill('B4E5A2'), al: CEN, border: BOX, fmt: '#,##0;-#,##0;"-"' });
    });
    put(ws, R_TOT, 2, 'TOTAL', { font: font({ size: 11, bold: true }), fill: fill('E27EDB'), al: CEN, border: BOX });
    var span = n + 3;
    for (var j = 0; j < span; j++) {
      var cc = cP + j, colSum = 0;
      if (j < n) { CALX.forEach(function (c) { colSum += M.rows[j].eff[c.k]; }); }
      else if (j === n) colSum = M.finT;
      else if (j === n + 1) colSum = M.manT;
      else colSum = M.finT - M.manT;
      var isS = j === n + 2;
      var f = isS ? A(cF, R_TOT) + '-' + A(cM, R_TOT) : 'SUM(' + A(cc, R0) + ':' + A(cc, R0 + nC - 1) + ')';
      put(ws, R_TOT, cc, fx(f, colSum), { font: font({ size: 11, bold: true }), fill: fill(j === n + 1 ? 'F4B183' : 'E27EDB'), al: CEN, border: BOX, fmt: NUM });
    }
    put(ws, R_INC, 2, 'Aporte del productor', { font: font({ size: 8, italic: true, color: { argb: 'FF555555' } }), al: { horizontal: 'right', vertical: 'middle' } });
    M.rows.forEach(function (x, i) {
      var f = i === 0 ? A(cP, R_TOT) : A(cP + i, R_TOT) + '-' + A(cP + i - 1, R_TOT);
      put(ws, R_INC, cP + i, fx(f, x.t), { font: font({ size: 9, italic: true, color: { argb: 'FF555555' } }), fill: fill('DDEBF7'), al: CEN, fmt: NUM });
    });

    /* tabla inferior */
    var R_BH = R_INC + 3, B0 = R_BH + 1;
    ['PRODUCTOR', 'CODIGO DE TRAZABILIDAD', 'CAJAS GRANDES', 'CAJAS CHICAS', 'KILOS', 'CAJAS CAMPO'].concat(CALX.map(function (c) { return c.l; }), ['TOTAL LIMON EXPORTACION', 'SEGUNDAS LIMON NACIONAL', 'TORREON / 275s', 'DESECHO', 'JUGUERA', 'RENDIMIENTO 1RAS', 'DIFERENCIA', 'TIPO DE EMPAQUE'], WU.map(function (t) { return 'CAJAS ' + String(WN[t] || t).toUpperCase(); }), ['KILOS EMPACADOS']).forEach(function (h, i) {
      put(ws, R_BH, 2 + i, h, { font: font({ size: 9, bold: true, color: { argb: 'FFFFFFFF' } }), fill: fill('000000'), al: CEN, border: BOX });
    });
    ws.getRow(R_BH).height = 42;
    var cCamp = 7;
    /* tabla de pesos por caja (editable) debajo de los totales, columnas B y C: las fórmulas de cada renglón la usan */
    var RT_ = n > 0 ? B0 + nRows : B0, R_PW = RT_ + 2, PW = {};
    WK.forEach(function (t, i) { PW[t] = A(3, R_PW + 1 + i, true); });
    function keF(r) { return WU.map(function (t, j) { return A(cW0 + j, r) + '*' + PW[t]; }).join('+'); }
    M.rows.forEach(function (x, i) {
      var r = B0 + RS[i], kil = (x.p.kilos == null || x.p.kilos === '') ? null : +x.p.kilos, ap = APS[i], rn = r + 1 + ap.length;
      ws.getRow(r).height = 22; ws.getRow(rn).height = 15;
      ap.forEach(function (a, j) {
        var ra = r + 1 + j, st = { font: font({ size: 9, italic: true }), fill: fill('F2F2F2'), border: BOX };
        ws.getRow(ra).height = 16;
        put(ws, ra, 2, '      ' + (a.nombre || ''), Object.assign({}, st, { al: { horizontal: 'left', vertical: 'middle' } }));
        put(ws, ra, 3, a.codigo ? (o.emp + '-' + String(a.codigo).padStart(3, '0')) : '', Object.assign({}, st, { al: CEN }));
        put(ws, ra, 4, '', st);
        put(ws, ra, 5, +a.n, Object.assign({}, st, { font: font({ size: 10, bold: true }), al: CEN, fmt: NUM }));
        put(ws, ra, 6, +a.n * 18, Object.assign({}, st, { al: CEN, fmt: NUM1 }));
        put(ws, ra, 7, '', st);
        /* la procedencia va al final (fuera de las columnas de calibre, para no afectar las fórmulas) */
        merge(ws, ra, cRn, ra, cEnd);
        put(ws, ra, cRn, [dmy(a.fecha), a.de || ''].filter(Boolean).join(' · '), Object.assign({}, st, { al: { horizontal: 'left', vertical: 'middle' } }));
      });
      put(ws, r, 2, x.p.nombre || ('Productor ' + (i + 1)), { font: font({ size: 10, bold: true }), fill: fill('D9D9D9'), al: { horizontal: 'left', vertical: 'middle', wrapText: true }, border: BOX });
      put(ws, r, 3, x.p.codigo ? (o.emp + '-' + String(x.p.codigo).padStart(3, '0')) : 0, { font: font({ size: 10, bold: true, color: { argb: 'FFFF0000' } }), al: CEN, border: BOX, fmt: NUM });
      put(ws, r, 4, x.p.k24 == null ? null : +x.p.k24, { font: font({ size: 11, bold: true }), al: CEN, border: BOX, fmt: NUM });
      put(ws, r, 5, x.p.k18 == null ? null : +x.p.k18, { font: font({ size: 11, bold: true }), al: CEN, border: BOX, fmt: NUM });
      put(ws, r, 6, kil, { font: font({ size: 11, bold: true }), al: CEN, border: BOX, fmt: NUM1 });
      put(ws, r, 7, fx('IF(' + A(6, r) + '>0,' + A(6, r) + '/' + KG + ',0)', kil ? kil / kg : 0), { font: font({ size: 11, bold: true }), al: CEN, border: BOX, fmt: NUM1 });
      CALX.forEach(function (c, ci) {
        var tr = R0 + ci, f = i === 0 ? A(cP, tr) : A(cP + i, tr) + '-' + A(cP + i - 1, tr);
        put(ws, r, 8 + ci, fx(f, x.delta[c.k]), { font: font({ size: 11 }), al: CEN, border: BOX, fmt: NUM });
      });
      put(ws, r, cT, fx('SUM(' + A(8, r) + ':' + A(7 + nC, r) + ')', x.t), { font: font({ size: 11, bold: true }), al: CEN, border: BOX, fmt: NUM });
      put(ws, r, cSg, x.p.seg == null && x.p.cera == null ? null : (+x.p.seg || 0) + (+x.p.cera || 0), /* v8.13: incluye las segundas con cera */ { font: font({ size: 11 }), fill: fill('D9D9D9'), al: CEN, border: BOX, fmt: NUM });
      put(ws, r, cTr, x.p.tor == null ? null : +x.p.tor, { font: font({ size: 11 }), fill: fill('D9D9D9'), al: CEN, border: BOX, fmt: NUM });
      put(ws, r, cDs, x.p.des == null ? null : +x.p.des, { font: font({ size: 11 }), fill: fill('D9D9D9'), al: CEN, border: BOX, fmt: NUM });
      put(ws, r, cJg, x.p.jug == null ? null : +x.p.jug, { font: font({ size: 11 }), fill: fill('D9D9D9'), al: CEN, border: BOX, fmt: NUM });
      var rr = r;
      /* cajas por tipo de empaque (con un solo tipo en la hoja, es el total con fórmula) y kilos empacados = cajas × peso por caja */
      WU.forEach(function (t, j) {
        var bv = (x.byW && x.byW[t]) || 0;
        put(ws, r, cW0 + j, WU.length === 1 ? fx(A(cT, rr), bv) : (bv || null), { font: font({ size: 11 }), fill: fill('EAF3E0'), al: CEN, border: BOX, fmt: NUM });
      });
      put(ws, r, cKE, fx(keF(rr), x.kgE || 0), { font: font({ size: 11, bold: true }), fill: fill('EAF3E0'), al: CEN, border: BOX, fmt: NUM1 });
      put(ws, r, cRn, fx('IF(' + A(7, rr) + '>0,' + A(cKE, rr) + '/' + A(6, rr) + ',"")', x.rend == null ? '' : x.rend), { font: font({ size: 11, bold: true, color: { argb: 'FFFF0000' } }), al: CEN, border: BOX, fmt: '0.0%' });
      put(ws, r, cDf, fx('IF(' + A(7, rr) + '>0,' + A(7, rr) + '-' + A(cKE, rr) + '/' + KG + '-' + A(cSg, rr) + '-' + A(cTr, rr) + '-' + A(cDs, rr) + '-' + A(cJg, rr) + ',"")', x.dif == null ? '' : x.dif), { font: font({ size: 11, bold: true, color: { argb: 'FFFF0000' } }), al: CEN, border: BOX, fmt: '#,##0.0;-#,##0.0;"-"' });
      put(ws, r, cTp, (d.tipoNames && d.tipoNames[x.tipo]) || x.tipo || 'RPC', { font: font({ size: 11, bold: true }), al: CEN, border: BOX });
      merge(ws, rn, 2, rn, cEnd);
      put(ws, rn, 2, NS[i].length ? ('Notas: sobrantes del viaje anterior, por productor (código · nombre · cajas · calibres · tipo · tarima de origen)' + (x.p.nota ? '. ' + x.p.nota : '')) : ([(d.entTxt && d.entTxt[x.p.id]) || '', x.p.nota ? 'Notas: ' + x.p.nota : ''].filter(Boolean).join(' · ')), { font: font({ size: 9, italic: true, bold: !!NS[i].length, color: { argb: 'FF444444' } }), al: { horizontal: 'left', vertical: 'middle', wrapText: true } });
      border(ws, rn, 2, rn, cEnd, HBOX);
      NS[i].forEach(function (t, j) {
        var rj = rn + 1 + j; ws.getRow(rj).height = 15; merge(ws, rj, 2, rj, cEnd);
        put(ws, rj, 2, '      ' + t, { font: font({ size: 9, italic: true, color: { argb: 'FF444444' } }), al: { horizontal: 'left', vertical: 'middle' } });
        border(ws, rj, 2, rj, cEnd, HBOX);
      });
    });
    var rowsEnd = B0 + nRows - 1;      /* última fila de captura */
    var dataRows = M.rows.map(function (x, i) { return B0 + RS[i]; });
    var RT = n > 0 ? rowsEnd + 1 : B0;   /* fila de totales */
    if (RT !== RT_) throw new Error('posición de la fila de totales');
    var a = B0, b = Math.max(rowsEnd, B0);
    put(ws, RT, 2, 'TOTAL', { font: font({ size: 11, bold: true }), al: CEN, fill: fill('EDEDED'), border: BOX });
    put(ws, RT, 3, '', { fill: fill('EDEDED'), border: BOX });
    var tk = M.tot;
    var sums = { 4: tk.k24, 5: tk.k18, 6: tk.kilos };
    CALX.forEach(function (c, ci) { sums[8 + ci] = M.final[c.k]; });
    WU.forEach(function (t, j) { var sv = 0; M.rows.forEach(function (x) { sv += (x.byW && x.byW[t]) || 0; }); sums[cW0 + j] = sv; });
    sums[cKE] = tk.kgE || 0;
    sums[cT] = tk.exp; sums[cSg] = tk.seg + (tk.cera || 0); sums[cTr] = tk.tor; sums[cDs] = tk.des; sums[cJg] = tk.jug || 0;
    var campoTot = 0; M.rows.forEach(function (x) { campoTot += x.campo || 0; });
    sums[7] = campoTot;
    Object.keys(sums).forEach(function (k) {
      var c = +k;
      /* con productores que aportan segundas, la suma toma solo los renglones de captura para no contar doble */
      put(ws, RT, c, fx(n > 0 ? (anyAp ? 'SUM(' + dataRows.map(function (rr) { return A(c, rr); }).join(',') + ')' : 'SUM(' + A(c, a) + ':' + A(c, b) + ')') : '0', n > 0 ? sums[c] : 0), { font: font({ size: 11, bold: true }), al: CEN, fill: fill('EDEDED'), border: BOX, fmt: c === 6 || c === 7 || c === cKE ? NUM1 : NUM });
    });
    var G = A(7, a) + ':' + A(7, b);
    var totEqv = 0, totCampoR = 0; M.rows.forEach(function (x) { if (x.rend != null) { totEqv += x.eqv; totCampoR += x.campo; } });
    /* v8.14: kilos empacados de los productores con kilos, entre el kg por caja campo, entre sus cajas campo */
    var numer = 'SUMIF(' + G + ',">0",' + A(cKE, a) + ':' + A(cKE, b) + ')/' + KG;
    put(ws, RT, cRn, fx(n > 0 ? 'IF(SUMIF(' + G + ',">0")>0,(' + numer + ')/SUMIF(' + G + ',">0"),"")' : '""', totCampoR > 0 ? totEqv / totCampoR : ''), { font: font({ size: 11, bold: true, color: { argb: 'FFFF0000' } }), al: CEN, fill: fill('EDEDED'), border: BOX, fmt: '0.0%' });
    var difTot = 0; M.rows.forEach(function (x) { if (x.dif != null) difTot += x.dif; });
    put(ws, RT, cDf, fx(n > 0 ? 'SUM(' + A(cDf, a) + ':' + A(cDf, b) + ')' : '0', difTot), { font: font({ size: 11, bold: true, color: { argb: 'FFFF0000' } }), al: CEN, fill: fill('EDEDED'), border: BOX, fmt: '#,##0.0;-#,##0.0;"-"' });
    ws.getRow(RT).height = 22;
    /* kilos / boxes / total export lemon, como en el formato original */
    var rk = RT + 2, rb = RT + 3, rl = RT + 5;
    put(ws, rk, cT - 1, 'KILOS', { font: font({ size: 10, bold: true, color: { argb: 'FFFFFFFF' } }), fill: fill('000000'), al: CEN });
    put(ws, rb, cT - 1, 'BOXES', { font: font({ size: 12, bold: true, color: { argb: 'FFFFFFFF' } }), fill: fill('00B050'), al: CEN });
    [[cT, tk.exp], [cSg, tk.seg + (tk.cera || 0)], [cTr, tk.tor], [cDs, tk.des], [cJg, tk.jug || 0]].forEach(function (q) {
      var c = q[0], ex = c === cT, KS = +d.kgSeg > 0 ? +d.kgSeg : 18;
      /* v8.14: los kilos de exportación son los kilos empacados (cajas × peso de su tipo de empaque) */
      put(ws, rk, c, ex ? fx(A(cKE, RT), tk.kgE || 0) : fx(A(c, RT) + '*' + KS, q[1] * KS), { font: font({ size: 10, bold: true, color: { argb: 'FFFFFFFF' } }), fill: fill('000000'), al: CEN, fmt: NUM });
      put(ws, rb, c, ex ? fx(A(cT, RT), q[1]) : fx(A(c, rk) + '/' + KS, q[1]), { font: font({ size: 12, bold: true, color: { argb: 'FFFFFFFF' } }), fill: fill('00B050'), al: CEN, fmt: NUM });
    });
    /* tabla de pesos por caja: se puede corregir aquí y se recalculan kilos empacados, rendimiento y diferencia */
    merge(ws, R_PW, 2, R_PW, 3);
    put(ws, R_PW, 2, 'PESO POR CAJA (kg)', { font: font({ size: 10, bold: true, color: { argb: 'FFFFFFFF' } }), fill: fill('375623'), al: CEN, border: BOX });
    ws.getCell(R_PW, 3).border = BOX;
    WK.forEach(function (t, i) {
      var r = R_PW + 1 + i, wv = +W[t] > 0 ? +W[t] : WD[t];
      put(ws, r, 2, String(WN[t] || t), { font: font({ size: 10, bold: WU.indexOf(t) >= 0 }), al: LEFT, border: BOX });
      put(ws, r, 3, wv, { font: font({ size: 11, bold: true, color: { argb: 'FF0000FF' } }), fill: fill('FFF2CC'), al: CEN, border: BOX, fmt: '0.00' });
    });
    put(ws, rl - 1, cT, 'EXPORTADO', { font: font({ size: 8, bold: true, color: { argb: 'FF555555' } }), al: CEN });
    put(ws, rl - 1, cSg, 'MANIFIESTO', { font: font({ size: 8, bold: true, color: { argb: 'FF555555' } }), al: CEN });
    put(ws, rl - 1, cTr, 'DIFERENCIA', { font: font({ size: 8, bold: true, color: { argb: 'FF555555' } }), al: CEN });
    merge(ws, rl, cT - 3, rl, cT - 1);
    put(ws, rl, cT - 3, 'TOTAL EXPORT LEMON', { font: font({ size: 12, bold: true }), fill: fill('92D050'), al: CEN, border: BOX });
    fillRange(ws, rl, cT - 3, rl, cT - 1, fill('92D050'));
    put(ws, rl, cT, fx(A(cT, RT), tk.exp), { font: font({ size: 12, bold: true }), fill: fill('D5FC79'), al: CEN, border: BOX, fmt: NUM });
    put(ws, rl, cSg, fx(A(cM, R_TOT), M.manT), { font: font({ size: 12, bold: true }), al: CEN, border: BOX, fmt: NUM });
    put(ws, rl, cTr, fx(A(cT, rl) + '-' + A(cSg, rl), tk.exp - M.manT), { font: font({ size: 12, bold: true }), fill: fill('FFFF00'), al: CEN, border: BOX, fmt: '#,##0;-#,##0;"-"' });
    ws.getRow(rl).height = 24;
    if (d.sobrante && d.sobrante.total > 0) {
      put(ws, rl + 2, 2, 'Sobrante para el siguiente pedido: ' + d.sobrante.text, { font: font({ size: 10, bold: true, color: { argb: 'FF9C5700' } }) });
    }
    var endR = Math.max(rl + 3, R_PW + WK.length + 1), endC = Math.max(cEnd, cS);
    /* manifiesto por tipo de empaque: tres columnas por tipo (manifiesto capturable, empacado y diferencia con fórmula) */
    if (LN.length) {
      if (R_MT !== rl + (d.sobrante && d.sobrante.total > 0 ? 4 : 3)) throw new Error('posición del manifiesto por tipo');
      var lastMC = 2 + LN.length * 3;
      merge(ws, R_MT, 2, R_MT, lastMC);
      put(ws, R_MT, 2, 'MANIFIESTO POR TIPO DE EMPAQUE · ' + (d.calidad || '').toUpperCase(), { font: font({ size: 12, bold: true }), fill: fill('65FB27'), al: CEN, border: BOX });
      fillRange(ws, R_MT, 2, R_MT, lastMC, fill('65FB27'));
      merge(ws, R_MT + 1, 2, R_MT + 2, 2);
      put(ws, R_MT + 1, 2, 'CALIBRE', { font: font({ size: 10, bold: true }), fill: fill('D5FC79'), al: CEN, border: BOX });
      LN.forEach(function (x, li) {
        var c0 = 3 + li * 3;
        merge(ws, R_MT + 1, c0, R_MT + 1, c0 + 2);
        put(ws, R_MT + 1, c0, x.label, { font: font({ size: 10, bold: true }), fill: fill('D5FC79'), al: CEN, border: BOX });
        fillRange(ws, R_MT + 1, c0, R_MT + 1, c0 + 2, fill('D5FC79')); border(ws, R_MT + 1, c0, R_MT + 1, c0 + 2, BOX);
        ['MANIF.', 'EMPAC.', 'DIF.'].forEach(function (t, j) { put(ws, R_MT + 2, c0 + j, t, { font: font({ size: 8, bold: true }), fill: fill(j === 0 ? 'CCECFF' : 'EDEDED'), al: CEN, border: BOX }); });
      });
      CALX.forEach(function (c, ci) {
        var r = R_M0 + ci;
        put(ws, r, 2, c.l, { font: font({ size: 11, bold: true }), al: CEN, border: BOX });
        LN.forEach(function (x, li) {
          var c0 = 3 + li * 3, mv = x.man[c.k] == null ? null : x.man[c.k], fv = x.fin[c.k] || 0;
          put(ws, r, c0, mv, { font: font({ size: 11, bold: true, color: { argb: 'FF0000FF' } }), fill: fill('CCECFF'), al: CEN, border: BOX, fmt: NUM });
          put(ws, r, c0 + 1, fv, { font: font({ size: 11 }), al: CEN, border: BOX, fmt: NUM });
          put(ws, r, c0 + 2, fx('IF(ISNUMBER(' + A(c0, r) + '),' + A(c0 + 1, r) + '-' + A(c0, r) + ',"")', mv == null ? '' : fv - mv), { font: font({ size: 11, bold: true }), fill: fill('B4E5A2'), al: CEN, border: BOX, fmt: '#,##0;-#,##0;"-"' });
        });
      });
      var rT = R_M0 + nC;
      put(ws, rT, 2, 'TOTAL', { font: font({ size: 11, bold: true }), fill: fill('E27EDB'), al: CEN, border: BOX });
      LN.forEach(function (x, li) {
        var c0 = 3 + li * 3, tm = 0, tf = 0;
        CALX.forEach(function (c) { tm += x.man[c.k] || 0; tf += x.fin[c.k] || 0; });
        [0, 1, 2].forEach(function (j) {
          var cc = c0 + j, f = 'SUM(' + A(cc, R_M0) + ':' + A(cc, rT - 1) + ')';
          put(ws, rT, cc, fx(f, j === 0 ? tm : j === 1 ? tf : tf - tm), { font: font({ size: 11, bold: true }), fill: fill('E27EDB'), al: CEN, border: BOX, fmt: '#,##0;-#,##0;"-"' });
        });
      });
      put(ws, rT + 1, 2, 'El MANIFIESTO de la tabla de arriba es la suma de estos tipos de empaque. Captura aquí el manifiesto de cada tipo.', { font: font({ size: 8, italic: true, color: { argb: 'FF555555' } }) });
      endR = rT + 2; endC = Math.max(endC, lastMC);
    }
    ws.pageSetup.printArea = 'A1:' + colL(endC) + endR;
    return ws;
  };

  /* ============ FORMATO 2: CÓDIGO DE TRAZABILIDAD ============ */
  /* d: {o, hdr, pallets:[{n,marca,lines:[{cal,cant,prod}]}], borrador, motivos[], logo, trailer:{b64,w,h}, sheetName, full} */
  XL.addTraz = function (wb, d) {
    var o = d.o, h = d.hdr || {};
    /* v8.19: orientación elegida en la app (vertical u horizontal), siempre con las tarimas en dos tablas.
       Las dos salen en una sola hoja y el área de impresión siempre incluye completo el dibujo del tráiler. */
    var H = d.orient === 'h', G0 = [1, 8], /* v8.19.1: las dos orientaciones conservan las dos tablas de tarimas */ lastG = G0[G0.length - 1] + 5, TC = lastG + 1; /* TC: columna (base 0) donde va el tráiler */
    var ws = wb.addWorksheet(safeName(wb, d.sheetName || o.emb || 'TRAZABILIDAD'), {
      views: [{ showGridLines: false }],
      pageSetup: { orientation: H ? 'landscape' : 'portrait', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 1, horizontalCentered: true, margins: { left: .3, right: .3, top: .3, bottom: .3, header: .1, footer: .1 } }
    });
    ws.properties.tabColor = { argb: d.borrador ? 'FFC00000' : 'FF00B050' };
    var GW = [7, 17, 9, 10, 9, 13], WS = [];
    G0.forEach(function (g, gi) { WS = WS.concat(GW); WS.push(2.5); });
    WS = WS.concat([12, 2, 2]);
    WS.forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    var LB = font({ size: 10, bold: true }), VAL = font({ size: 10 });
    merge(ws, 1, 1, 1, 6);
    put(ws, 1, 1, 'CODIGO DE TRAZABILIDAD', { font: font({ size: 15, bold: true }), al: CEN });
    ws.getRow(1).height = 24;
    /* v8.47: sin celda de estatus (BORRADOR / PEDIDO FULL / PARCIAL) en el documento, a pedido de la empresa */
    var L = [['Factura#:', h.factura || o.factura], ['Proveedor:', h.proveedor], ['Vendedor:', h.vendedor], ['Tel:', h.tel]];
    L.forEach(function (p, i) {
      var r = 2 + i;
      merge(ws, r, 1, r, 2); put(ws, r, 1, p[0], { font: LB, al: LEFT, border: BOX });
      merge(ws, r, 3, r, 6); put(ws, r, 3, numOrStr(p[1] || ''), { font: VAL, al: LEFT, border: BOX });
      ws.getRow(r).height = 15;
    });
    var Rr = [['PO# CLIENTE:', h.poCliente || o.po], ['Cliente:', h.cliente || o.customer], ['P. Order', h.pOrder], ['SELLO', h.sello]];
    Rr.forEach(function (p, i) {
      var r = 2 + i;
      merge(ws, r, 8, r, 9); put(ws, r, 8, p[0], { font: LB, al: LEFT, border: BOX });
      merge(ws, r, 10, r, 13); put(ws, r, 10, numOrStr(p[1] || ''), { font: font({ size: 10, bold: true }), al: LEFT, border: BOX });
    });
    put(ws, 6, 1, '', {}); ws.getRow(6).height = 6;
    img(wb, ws, d.logo, 0, 6, 128, 82);
    var T = [['Transporte:', h.transporte], ['Teléfono:', h.telTransporte], ['Conductor:', h.conductor], ['PLACAS:', h.placas], ['#CAJA:', h.caja]];
    T.forEach(function (p, i) {
      var r = 7 + i;
      merge(ws, r, 5, r, 6); put(ws, r, 5, p[0], { font: LB, al: { horizontal: 'right', vertical: 'middle' } });
      merge(ws, r, 8, r, 13); put(ws, r, 8, numOrStr(p[1] || ''), { font: font({ size: 10 }), al: LEFT, border: { bottom: hair } });
      ws.getRow(r).height = 15;
    });
    ws.getRow(12).height = 8;
    /* cabecera de la tabla de tarimas (dos grupos de columnas) */
    var HR = 13, HC = [], HC2 = [];
    G0.forEach(function (g) { HC.push([g, 'PALLET'], [g + 1, 'MARCA'], [g + 2, 'CALIBRE'], [g + 3, 'CANTIDAD']); HC2.push([g + 4, 'CODIGO DE TRAZABILIDAD']); });
    HC.forEach(function (p) {
      merge(ws, HR, p[0], HR + 1, p[0]);
      put(ws, HR, p[0], p[1], { font: font({ size: 8, bold: true }), fill: fill('D9D9D9'), al: CEN, border: BOX });
      ws.getCell(HR + 1, p[0]).border = BOX;
    });
    HC2.forEach(function (p) {
      merge(ws, HR, p[0], HR, p[0] + 1);
      put(ws, HR, p[0], p[1], { font: font({ size: 8, bold: true }), fill: fill('D9D9D9'), al: CEN, border: BOX });
      ws.getCell(HR, p[0] + 1).border = BOX;
      put(ws, HR + 1, p[0], 'CAJAS', { font: font({ size: 8, bold: true }), fill: fill('D9D9D9'), al: CEN, border: BOX });
      put(ws, HR + 1, p[0] + 1, 'CODIGO', { font: font({ size: 8, bold: true }), fill: fill('D9D9D9'), al: CEN, border: BOX });
    });
    ws.getRow(HR).height = 14; ws.getRow(HR + 1).height = 14;
    var pal = (d.pallets || []).slice(), total = 0, nReal = pal.filter(function (p) { return !p.other; }).length;
    /* con la foto del tráiler se dibujan los 22 lugares, aunque falten tarimas, para que la tabla tenga el largo de la caja */
    if (d.pad) for (var pi = pal.length; pi < d.pad; pi++) pal.push({ n: pi + 1, marca: '', lines: [], empty: true });
    pal.forEach(function (p) { p.lines = (p.lines && p.lines.length) ? p.lines : [{ cal: '', cant: null, prod: '' }]; total += p.lines.length; });
    /* reparte las tarimas en grupos con filas parecidas, en orden (2 en vertical, 3 en horizontal) */
    var NG = G0.length, groups = G0.map(function () { return []; }), acc = 0, gi0 = 0;
    pal.forEach(function (p) { while (gi0 < NG - 1 && acc >= total * (gi0 + 1) / NG) gi0++; groups[gi0].push(p); acc += p.lines.length; });
    var firstRow = HR + 2, sumCells = {}; G0.forEach(function (g, i) { sumCells[i + 1] = []; });
    function draw(list, c0, gi) {
      var r = firstRow;
      list.forEach(function (p) {
        var n = p.lines.length, r2 = r + n - 1, tot = 0;
        if (p.other) { /* v8.36.2: tarima de otro manifiesto del viaje: lugar en gris y sin datos, como en el packing list */
          merge(ws, r, c0, r2, c0); merge(ws, r, c0 + 1, r2, c0 + 1); merge(ws, r, c0 + 3, r2, c0 + 3);
          for (var oc = c0; oc <= c0 + 5; oc++) for (var orr = r; orr <= r2; orr++) { var ce = ws.getCell(orr, oc); ce.fill = fill('D9D9D9'); ce.border = BOX; }
          for (var orr2 = r; orr2 <= r2; orr2++) ws.getRow(orr2).height = 13.5;
          for (var c1 = c0; c1 <= c0 + 5; c1++) { var cl = ws.getCell(r2, c1), b1 = Object.assign({}, cl.border || BOX); b1.bottom = med; cl.border = b1; }
          r = r2 + 1; return;
        }
        p.lines.forEach(function (l) { tot += +l.cant || 0; });
        merge(ws, r, c0, r2, c0); merge(ws, r, c0 + 1, r2, c0 + 1); merge(ws, r, c0 + 3, r2, c0 + 3);
        put(ws, r, c0, p.n, { font: font({ size: 10, bold: true }), al: CEN, border: BOX });
        put(ws, r, c0 + 1, (p.marca || '') + (p.tag ? '\n' + p.tag : ''), { font: font({ size: 8, bold: true }), al: CEN, border: BOX });
        if (p.tag && n < 2) ws.getRow(r).height = 24;
        var cap = +p.cap || 50, ok = tot === cap || (cap === 50 && (tot === 54 || tot === 55 || tot === 60)); /* v8.19: estándar según presentación */
        var dFill = (tot > 0 && !ok) ? fill('FCE4D6') : undefined;
        put(ws, r, c0 + 3, fx('SUM(' + A(c0 + 4, r) + ':' + A(c0 + 4, r2) + ')', tot), { font: font({ size: 10, bold: true }), fill: dFill, al: CEN, border: BOX, fmt: NUM });
        sumCells[gi].push(A(c0 + 3, r));
        p.lines.forEach(function (l, i) {
          var rr = r + i;
          ws.getRow(rr).height = (p.tag && n < 2) ? 24 : 13.5;
          put(ws, rr, c0 + 2, numOrStr(l.cal || ''), { font: font({ size: 9, bold: true }), al: CEN, border: BOX });
          put(ws, rr, c0 + 4, l.cant == null || l.cant === '' ? null : +l.cant, { font: font({ size: 9, bold: true }), al: CEN, border: BOX, fmt: NUM });
          put(ws, rr, c0 + 5, l.prod || '', { font: font({ size: 9, bold: true }), al: CEN, border: BOX });
        });
        for (var c = c0; c <= c0 + 5; c++) { var cell = ws.getCell(r2, c), b = Object.assign({}, cell.border || BOX); b.bottom = med; cell.border = b; }
        r = r2 + 1;
      });
      return r;
    }
    var ends = groups.map(function (L, i) { return draw(L, G0[i], i + 1); });
    var end = Math.max.apply(null, ends.concat([firstRow + 1]));
    /* pie: totales */
    var tr = end + 1, all = []; G0.forEach(function (g, i) { all = all.concat(sumCells[i + 1]); });
    var cant = 0; pal.forEach(function (p) { p.lines.forEach(function (l) { cant += +l.cant || 0; }); });
    merge(ws, tr, 1, tr, 3); put(ws, tr, 1, 'TOTAL TARIMAS', { font: LB, fill: fill('D9D9D9'), al: CEN, border: BOX });
    put(ws, tr, 4, nReal, { font: LB, al: CEN, border: BOX });
    merge(ws, tr + 1, 1, tr + 1, 3); put(ws, tr + 1, 1, 'TOTAL CAJAS', { font: LB, fill: fill('D9D9D9'), al: CEN, border: BOX });
    put(ws, tr + 1, 4, fx(all.length ? all.join('+') : '0', cant), { font: LB, al: CEN, border: BOX, fmt: NUM });
    if (d.motivos && d.motivos.length) {
      merge(ws, tr + 3, 1, tr + 3, lastG);
      put(ws, tr + 3, 1, 'Pendiente para completar: ' + d.motivos.join(' · '), { font: font({ size: 9, italic: true, color: { argb: 'FFC00000' } }), al: { horizontal: 'left', vertical: 'middle', wrapText: true } });
      ws.getRow(tr + 3).height = 28;
    }
    /* foto del tráiler: la caja (sin la cabina) mide lo mismo que la tabla de tarimas, del encabezado a la última fila */
    var imgBot = 0; /* pt: borde de abajo del dibujo del tráiler, para que el área de impresión no lo corte */
    if (d.trailer && d.trailer.b64 && d.trailer.photo) {
      var T0 = rowTop(ws, HR), T1 = rowTop(ws, end), P = d.trailer, k = (T1 - T0) / (P.boxBot - P.boxTop);
      /* v8.7: la caja se dibuja con un lugar por tarima; con pocas o muchas tarimas el dibujo no queda ni diminuto ni enorme */
      k = Math.max(95 / P.w, Math.min(k, 150 / P.w));
      var wpx = P.w * k * 4 / 3, hpx = P.h * k * 4 / 3, top = T0 - P.boxTop * k;
      if (top < 0) { k = T1 / (P.h - (P.h - P.boxBot)); wpx = P.w * k * 4 / 3; hpx = P.h * k * 4 / 3; top = T1 - P.boxBot * k; }
      ws.getColumn(TC + 1).width = Math.max(12, Math.ceil(wpx / 7) + 2); ws.getColumn(TC + 2).width = 2; ws.getColumn(TC + 3).width = 2;
      imgAt(wb, ws, P, TC, top, wpx, hpx);
      imgBot = Math.max(0, top) + hpx * 0.75;
    } else if (d.trailer && d.trailer.b64) {
      var id = wb.addImage({ base64: d.trailer.b64, extension: 'png' });
      ws.addImage(id, { tl: { col: TC, row: 6 }, ext: { width: d.trailer.w, height: d.trailer.h } });
      ws.getColumn(TC + 1).width = Math.max(12, Math.ceil(d.trailer.w / 7) + 2);
      imgBot = rowTop(ws, 7) + d.trailer.h * 0.75;
    }
    var lastR = tr + 4; while (imgBot > 0 && rowTop(ws, lastR + 1) < imgBot + 4) lastR++;
    ws.pageSetup.printArea = 'A1:' + colL(d.trailer ? TC + 3 : lastG) + lastR;
    d._lastRow = end;
    return ws;
  };

  /* Resumen: tarimas contra lo empacado, por productor (fórmulas SUMIFS sobre la hoja de trazabilidad) */
  XL.addResumen = function (wb, d) {
    var ws = wb.addWorksheet(safeName(wb, 'Resumen ' + (d.o.cod || '')), { views: [{ showGridLines: false, state: 'frozen', ySplit: 4 }], pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    ws.getColumn(1).width = 30; ws.getColumn(2).width = 13;
    for (var c = 3; c <= 14; c++) ws.getColumn(c).width = 11;
    merge(ws, 1, 1, 1, 12);
    put(ws, 1, 1, 'Resumen de tarimas · ' + (d.o.cod || '') + ' · ' + (d.o.emb || ''), { font: font({ size: 15, bold: true }), al: LEFT });
    put(ws, 2, 1, d.borrador ? 'BORRADOR: todavía falta cargar o revisar tarimas.' : 'Tarimas completas.', { font: font({ size: 10, bold: true, color: { argb: d.borrador ? 'FFC00000' : 'FF006100' } }) });
    /* calibres de la tabla: los de tamaño y, si hay cartón de 10 lb, sus medidas */
    var CL = d.cal || CAL, nL = CL.length, cTot = 3 + nL;
    for (var cw = 3; cw <= cTot + 2; cw++) ws.getColumn(cw).width = 11;
    var heads = ['PRODUCTOR', 'CODIGO'].concat(CL.map(function (c) { return c.l; }), ['EN TARIMAS', 'EMPACADO', 'DIFERENCIA']);
    heads.forEach(function (t, i) { put(ws, 4, 1 + i, t, { font: font({ size: 9, bold: true, color: { argb: 'FFFFFFFF' } }), fill: fill('1B2C46'), al: CEN, border: BOX }); });
    ws.getRow(4).height = 28;
    var sn = "'" + d.trazSheet.replace(/'/g, "''") + "'!";
    function rng(col) { return sn + '$' + col + '$15:$' + col + '$400'; }
    d.productores.forEach(function (p, i) {
      var r = 5 + i;
      put(ws, r, 1, p.nombre, { font: font({ size: 10 }), border: BOX });
      put(ws, r, 2, p.key, { font: font({ size: 10, bold: true }), al: CEN, border: BOX });
      CL.forEach(function (c, ci) {
        var crit = /^\d+$/.test(c.l) ? c.l : '"' + c.l + '"';
        var f = 'SUMIFS(' + rng('E') + ',' + rng('F') + ',$B' + r + ',' + rng('C') + ',' + crit + ')+SUMIFS(' + rng('L') + ',' + rng('M') + ',$B' + r + ',' + rng('J') + ',' + crit + ')';
        put(ws, r, 3 + ci, fx(f, p.pal[c.k] || 0), { font: font({ size: 10 }), al: CEN, border: BOX, fmt: NUM });
      });
      put(ws, r, cTot, fx('SUM(' + A(3, r) + ':' + A(cTot - 1, r) + ')', p.palTot), { font: font({ size: 10, bold: true }), al: CEN, border: BOX, fmt: NUM });
      put(ws, r, cTot + 1, p.emp, { font: font({ size: 10 }), al: CEN, border: BOX, fmt: NUM });
      put(ws, r, cTot + 2, fx(A(cTot, r) + '-' + A(cTot + 1, r), p.palTot - p.emp), { font: font({ size: 10, bold: true, color: { argb: 'FFC00000' } }), al: CEN, border: BOX, fmt: '#,##0;-#,##0;"-"' });
    });
    var last = 4 + d.productores.length, tr = last + 1;
    put(ws, tr, 1, 'TOTAL', { font: font({ size: 10, bold: true }), fill: fill('EDEDED'), border: BOX });
    put(ws, tr, 2, '', { fill: fill('EDEDED'), border: BOX });
    for (var c2 = 3; c2 <= cTot + 2; c2++) {
      var tot = 0;
      d.productores.forEach(function (p) { if (c2 < cTot) tot += p.pal[CL[c2 - 3].k] || 0; else if (c2 === cTot) tot += p.palTot; else if (c2 === cTot + 1) tot += p.emp; else tot += p.palTot - p.emp; });
      put(ws, tr, c2, fx(d.productores.length ? 'SUM(' + A(c2, 5) + ':' + A(c2, last) + ')' : '0', tot), { font: font({ size: 10, bold: true }), fill: fill('EDEDED'), al: CEN, border: BOX, fmt: NUM });
    }
    put(ws, tr + 1, 1, 'El sobrante empacado por encima del manifiesto no se tarima: pasa al siguiente pedido.', { font: font({ size: 8, italic: true, color: { argb: 'FF555555' } }) });
    var r0 = tr + 4;
    put(ws, r0 - 1, 1, 'POR CALIBRE Y CALIDAD', { font: font({ size: 11, bold: true }) });
    ['CALIBRE', 'POR TARIMAR 1RA', 'TARIMAS 1RA', 'POR TARIMAR 1.5', 'TARIMAS 1.5', 'DIFERENCIA'].forEach(function (t, i) { put(ws, r0, 1 + i, t, { font: font({ size: 9, bold: true, color: { argb: 'FFFFFFFF' } }), fill: fill('1B2C46'), al: CEN, border: BOX }); });
    d.calibres.forEach(function (x, i) {
      var r = r0 + 1 + i;
      put(ws, r, 1, x.l, { font: font({ size: 10, bold: true }), al: CEN, border: BOX });
      put(ws, r, 2, x.e1, { al: CEN, border: BOX, fmt: NUM });
      put(ws, r, 3, x.p1, { al: CEN, border: BOX, fmt: NUM });
      put(ws, r, 4, x.e15, { al: CEN, border: BOX, fmt: NUM });
      put(ws, r, 5, x.p15, { al: CEN, border: BOX, fmt: NUM });
      put(ws, r, 6, fx('(C' + r + '+E' + r + ')-(B' + r + '+D' + r + ')', x.p1 + x.p15 - x.e1 - x.e15), { font: font({ bold: true, color: { argb: 'FFC00000' } }), al: CEN, border: BOX, fmt: '#,##0;-#,##0;"-"' });
    });
    /* v6: cada calidad y tipo de empaque por separado (las tarimas de cartón solo cubren lo empacado en cartón) */
    if (d.lineas && d.lineas.length) {
      var q0 = r0 + d.calibres.length + 4;
      put(ws, q0 - 1, 1, 'POR CALIDAD Y TIPO DE EMPAQUE', { font: font({ size: 11, bold: true }) });
      ['CALIDAD', 'TIPO DE EMPAQUE', 'CALIBRE', 'POR TARIMAR', 'EN TARIMAS', 'DIFERENCIA'].forEach(function (t, i) { put(ws, q0, 1 + i, t, { font: font({ size: 9, bold: true, color: { argb: 'FFFFFFFF' } }), fill: fill('1B2C46'), al: CEN, border: BOX }); });
      var prev = '';
      d.lineas.forEach(function (x, i) {
        var r = q0 + 1 + i, k = x.q + '|' + x.t, first = k !== prev; prev = k;
        put(ws, r, 1, first ? x.q : '', { font: font({ size: 10, bold: first }), fill: first ? fill('EDEDED') : undefined, border: BOX });
        put(ws, r, 2, first ? x.t : '', { font: font({ size: 10, bold: first }), fill: first ? fill('EDEDED') : undefined, al: CEN, border: BOX });
        put(ws, r, 3, x.l, { font: font({ size: 10, bold: true }), al: CEN, border: BOX });
        put(ws, r, 4, x.e, { al: CEN, border: BOX, fmt: NUM });
        put(ws, r, 5, x.p, { al: CEN, border: BOX, fmt: NUM });
        put(ws, r, 6, fx('E' + r + '-D' + r, x.p - x.e), { font: font({ bold: true, color: { argb: 'FFC00000' } }), al: CEN, border: BOX, fmt: '#,##0;-#,##0;"-"' });
      });
    }
    return ws;
  };

  /* ============ BASE DE DATOS DECORADA ============ */
  /* d: {cols:[{h,w,k,fmt}], rows:[{...}], pr:[...resumen productores], cal:[...resumen calibres], titulo, logo} */
  XL.addBase = function (wb, d) {
    var ws = wb.addWorksheet('Base de datos', { views: [{ state: 'frozen', xSplit: 3, ySplit: 4, showGridLines: false }] });
    merge(ws, 1, 1, 1, 8);
    put(ws, 1, 1, d.titulo || 'Base de datos de empaque', { font: font({ size: 16, bold: true, color: { argb: 'FF1B2C46' } }), al: LEFT });
    ws.getRow(1).height = 26;
    put(ws, 2, 1, d.sub || '', { font: font({ size: 9, italic: true, color: { argb: 'FF555555' } }) });
    d.cols.forEach(function (c, i) {
      ws.getColumn(i + 1).width = c.w || 11;
      put(ws, 4, i + 1, c.h, { font: font({ size: 9, bold: true, color: { argb: 'FFFFFFFF' } }), fill: fill(c.grp === 'cal' ? '4F6B0A' : c.grp === 'out' ? '9C5700' : '1B2C46'), al: CEN, border: BOX });
    });
    ws.getRow(4).height = 34;
    d.rows.forEach(function (row, ri) {
      var r = 5 + ri, zebra = ri % 2 ? fill('F2F5F0') : undefined;
      d.cols.forEach(function (c, i) {
        var v = row[c.k]; if (v === undefined) v = null;
        put(ws, r, i + 1, v, { font: font({ size: 10 }), fill: zebra, al: c.al ? { horizontal: c.al, vertical: 'middle' } : { vertical: 'middle' }, border: HBOX, fmt: c.fmt });
      });
    });
    var last = 4 + Math.max(d.rows.length, 1);
    ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: last, column: d.cols.length } };
    if (d.pr) {
      var w2 = wb.addWorksheet('Por productor', { views: [{ state: 'frozen', ySplit: 4, showGridLines: false }] });
      merge(w2, 1, 1, 1, 8); put(w2, 1, 1, 'Resumen por productor (rendimiento ponderado)', { font: font({ size: 16, bold: true, color: { argb: 'FF1B2C46' } }) });
      var hs = ['PRODUCTOR', 'CODIGO', 'PEDIDOS', 'KILOS', 'CAJAS EXPORTACION', 'SEGUNDAS', 'TORREON', 'DESECHO', 'RENDIMIENTO 1RAS'];
      var ws2 = [32, 10, 10, 12, 14, 11, 11, 11, 14];
      hs.forEach(function (t, i) { w2.getColumn(i + 1).width = ws2[i]; put(w2, 4, i + 1, t, { font: font({ size: 9, bold: true, color: { argb: 'FFFFFFFF' } }), fill: fill('1B2C46'), al: CEN, border: BOX }); });
      w2.getRow(4).height = 30;
      d.pr.forEach(function (p, i) {
        var r = 5 + i, z = i % 2 ? fill('F2F5F0') : undefined;
        [p.nombre, p.codigo, p.pedidos, p.kilos, p.exp, p.seg, p.tor, p.des, p.rend].forEach(function (v, k) {
          put(w2, r, k + 1, v == null ? null : v, { font: font({ size: 10, bold: k === 8 }), fill: z, border: HBOX, al: k === 0 ? LEFT : { horizontal: 'center', vertical: 'middle' }, fmt: k === 8 ? '0.0%' : (k >= 3 ? NUM : undefined) });
        });
      });
      w2.autoFilter = { from: { row: 4, column: 1 }, to: { row: 4 + Math.max(d.pr.length, 1), column: 9 } };
    }
    if (d.cal) {
      var w3 = wb.addWorksheet('Por calibre', { views: [{ showGridLines: false }] });
      merge(w3, 1, 1, 1, 6); put(w3, 1, 1, 'Distribución por calibre', { font: font({ size: 16, bold: true, color: { argb: 'FF1B2C46' } }) });
      ['CALIBRE', 'CAJAS', '% DE EXPORTACION'].forEach(function (t, i) { w3.getColumn(i + 1).width = [14, 14, 18][i]; put(w3, 3, i + 1, t, { font: font({ size: 9, bold: true, color: { argb: 'FFFFFFFF' } }), fill: fill('4F6B0A'), al: CEN, border: BOX }); });
      d.cal.forEach(function (c, i) {
        var r = 4 + i;
        put(w3, r, 1, c.l, { font: font({ bold: true }), al: CEN, border: HBOX });
        put(w3, r, 2, c.n, { al: CEN, border: HBOX, fmt: NUM });
        put(w3, r, 3, fx(A(2, r) + '/SUM($B$4:$B$' + (3 + d.cal.length) + ')', c.pct), { al: CEN, border: HBOX, fmt: '0.0%' });
      });
    }
    (d.extra || []).forEach(function (x) {
      var w = wb.addWorksheet(x.name, { views: [{ state: 'frozen', ySplit: 3, showGridLines: false }] });
      merge(w, 1, 1, 1, 6); put(w, 1, 1, x.name, { font: font({ size: 16, bold: true, color: { argb: 'FF1B2C46' } }) });
      x.cols.forEach(function (c, i) { w.getColumn(i + 1).width = c.w || 11; put(w, 3, i + 1, c.h, { font: font({ size: 9, bold: true, color: { argb: 'FFFFFFFF' } }), fill: fill('1B2C46'), al: CEN, border: BOX }); });
      x.rows.forEach(function (row, ri) {
        x.cols.forEach(function (c, i) { var v = row[c.k]; put(w, 4 + ri, i + 1, v === undefined ? null : v, { font: font({ size: 10 }), fill: ri % 2 ? fill('F2F5F0') : undefined, border: HBOX, al: c.al ? { horizontal: c.al, vertical: 'middle' } : { vertical: 'middle' }, fmt: c.fmt }); });
      });
      w.autoFilter = { from: { row: 3, column: 1 }, to: { row: 3 + Math.max(x.rows.length, 1), column: x.cols.length } };
    });
    return ws;
  };


  /* Packing List (una hoja por viaje): encabezado del envío, una fila por tarima y total de cajas.
     d: {sheetName, logo:{b64,w,h}, hdr:{salida,hora,llegada,proveedor,cliente,vendedor,tel,poCliente,pOrder,sello}, rows:[{n,tam,tipo,cal,cajas}]} */
  function isoDate(s) { var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ''); return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null; }
  function dayFrac(s) { var m = /^(\d{1,2}):(\d{2})/.exec(s || ''); return m ? (+m[1] * 60 + +m[2]) / 1440 : null; }
  XL.addPacking = function (wb, d) {
    var ws = wb.addWorksheet(safeName(wb, d.sheetName || 'Packing List'), {
      views: [{ showGridLines: false }],
      pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 1, horizontalCentered: true,
        margins: { left: 0.4, right: 0.4, top: 0.4, bottom: 0.5, header: 0.2, footer: 0.2 } }
    });
    [4.9, 15.6, 13.4, 28.9, 25.9, 11.4, 10.6, 15.4].forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    var LBL = fill('D6DCE4'), lab = font({ size: 12, bold: true }), val = font({ size: 14, bold: true });
    var h = d.hdr || {};
    /* filas 1 a 9: logotipo */
    var hts = { 3: 39, 4: 21, 5: 21, 6: 21, 7: 21, 8: 21, 9: 19.7 };
    Object.keys(hts).forEach(function (r) { ws.getRow(+r).height = hts[r]; });
    img(wb, ws, d.logo, 1, 0, 330, 238);
    /* fila 10: título */
    merge(ws, 10, 1, 10, 8);
    put(ws, 10, 1, 'PACKING LIST', { font: font({ size: 14, bold: true }), al: CEN, border: { bottom: thin } });
    ws.getRow(10).height = 25.7;
    /* filas 11 a 15: datos del envío */
    function label(r, c1, c2, t, al) {
      merge(ws, r, c1, r, c2);
      fillRange(ws, r, c1, r, c2, LBL); border(ws, r, c1, r, c2, BOX);
      put(ws, r, c1, t, { font: lab, fill: LBL, al: al || LEFT, border: BOX });
    }
    function value(r, c1, c2, v, fmt) {
      merge(ws, r, c1, r, c2); border(ws, r, c1, r, c2, BOX);
      put(ws, r, c1, v, { font: val, al: CEN, border: BOX, fmt: fmt });
    }
    var rows = [
      [11, 'Fecha de salida:', isoDate(h.salida), 'm/d/yyyy', 'Hora de salida', dayFrac(h.hora), 'h:mm AM/PM'],
      [12, 'Fecha de llegada:', isoDate(h.llegada), 'm/d/yyyy', 'PO# CLIENTE:', h.poCliente || '', null],
      [13, 'Proveedor:', h.proveedor || '', null, 'Cliente:', h.cliente || '', null],
      [14, 'Vendedor:', h.vendedor || '', null, 'P. Order:', numOrStr(h.pOrder || ''), null],
      [15, 'Tel:', numOrStr(h.tel || ''), null, 'Num. de sello:', numOrStr(h.sello || ''), null]
    ];
    var rh = { 11: 21, 12: 21, 13: 36.6, 14: 24.75, 15: 26.25 };
    rows.forEach(function (x) {
      var r = x[0];
      label(r, 2, 2, x[1]);
      value(r, 3, 4, x[2], x[3]);
      label(r, 5, 6, x[4]);
      value(r, 7, 8, x[5], x[6]);
      ws.getCell(r, 7).alignment = CEN;
      ws.getRow(r).height = rh[r];
    });
    ws.getRow(16).height = 18.6; ws.getRow(17).height = 25.35;
    /* filas 18 y 19: encabezado de la tabla */
    label(18, 2, 4, '', CEN); label(18, 5, 5, '', CEN); label(18, 6, 6, 'Ryan #:', LEFT);
    merge(ws, 18, 7, 19, 8); fillRange(ws, 18, 7, 19, 8, LBL); border(ws, 18, 7, 19, 8, BOX);
    put(ws, 18, 7, 'No.CAJAS', { font: lab, fill: LBL, al: CEN, border: BOX });
    ['No. Tarima', 'Producto', 'Tamaño', 'RPC/CTN', 'Calidad'].forEach(function (t, i) {
      put(ws, 19, 2 + i, t, { font: lab, fill: LBL, al: CEN, border: BOX });
    });
    ws.getRow(18).height = 24; ws.getRow(19).height = 26.25;
    /* tarimas */
    var list = (d.rows || []).slice(), r0 = 20, r = r0, total = 0, TP = d.trailer && d.trailer.b64 ? d.trailer : null, RHp = 24;
    if (TP) for (var pi = list.length; pi < TP.n; pi++) list.push({ n: pi + 1, empty: true });
    list.forEach(function (x) {
      if (x.empty) {
        var OF = x.other ? fill('D5D9DF') : undefined; /* v8.26: lugar de una tarima de otro manifiesto (otra venta) */
        put(ws, r, 2, x.n, { font: font({ size: 14, bold: true, color: { argb: 'FF888888' } }), al: { vertical: 'middle' }, border: BOX, fill: OF });
        [3, 4, 5, 6].forEach(function (c) { put(ws, r, c, '', { border: BOX, fill: OF }); });
        merge(ws, r, 7, r, 8); border(ws, r, 7, r, 8, BOX); if (OF) put(ws, r, 7, '', { border: BOX, fill: OF });
        ws.getRow(r).height = RHp; r++; return;
      }
      var tam = x.tam == null ? '' : (typeof x.tam === 'string' ? x.tam.replace(') Y (', ') Y\n(') : x.tam), tamTxt = String(tam), tipo = String(x.tipo || 'RPC'), cal = String(x.cal || 'PRIMERA');
      put(ws, r, 2, x.n, { font: font({ size: 16, bold: true }), al: { vertical: 'middle' }, border: BOX });
      put(ws, r, 3, 'LIMES', { font: font({ size: 16, bold: true }), al: CEN, border: BOX });
      put(ws, r, 4, tam, { font: font({ size: tamTxt.length > 24 ? 10 : (tamTxt.length > 12 ? 11 : (tamTxt.length > 7 ? 13 : 16)), bold: true }), al: CEN, border: BOX });
      put(ws, r, 5, tipo, { font: font({ size: tipo.length > 8 ? 12 : 14, bold: true }), al: CEN, border: BOX });
      put(ws, r, 6, cal, { font: font({ size: cal.length > 8 ? 10 : 14, bold: true }), al: CEN, border: BOX });
      merge(ws, r, 7, r, 8); border(ws, r, 7, r, 8, BOX);
      put(ws, r, 7, x.cajas || 0, { font: font({ size: 16, bold: true }), al: CEN, border: BOX, fmt: '0' });
      total += x.cajas || 0;
      ws.getRow(r).height = TP ? RHp : (tamTxt.length > 40 ? 44 : (tamTxt.length > 12 ? 36 : 30.95));
      if (TP && tamTxt.length > 12) ws.getCell(r, 4).font = font({ size: 9, bold: true });
      if (TP && cal.length > 12) ws.getCell(r, 6).font = font({ size: 8, bold: true }); /* tarima mixta: «PRIMERA / UNO Y MEDIO» en dos líneas */
      r++;
    });
    /* total */
    merge(ws, r, 2, r, 6); fillRange(ws, r, 2, r, 6, LBL); border(ws, r, 2, r, 6, BOX);
    put(ws, r, 2, 'TOTAL DE CAJAS', { font: font({ size: 14, bold: true }), fill: LBL, al: { horizontal: 'right', vertical: 'middle', indent: 1 }, border: BOX });
    merge(ws, r, 7, r, 8); border(ws, r, 7, r, 8, BOX);
    put(ws, r, 7, list.length ? fx('SUM(G' + r0 + ':G' + (r - 1) + ')', total) : 0, { font: font({ size: 16, bold: true }), al: CEN, border: BOX, fmt: '#,##0' });
    ws.getRow(r).height = 30;
    var endP = r;
    /* resumen al final: una tabla por calidad (TOTAL | calibre | cajas) y a la derecha total de tarimas, cajas y pesos */
    var totRow = r, totCells = [];
    if (d.calRes && d.calRes.length) {
      var y = r + 2, y0 = y, thin1 = { style: 'thin', color: { argb: 'FF000000' } }, B1 = { top: thin1, left: thin1, bottom: thin1, right: thin1 };
      d.calRes.forEach(function (g) {
        merge(ws, y, 2, y, 4);
        put(ws, y, 2, g.label, { font: font({ size: 9, bold: true }), al: CEN });
        ws.getRow(y).height = 16; y++;
        var first = y;
        g.rows.forEach(function (x, k) {
          var last = k === g.rows.length - 1;
          put(ws, y, 2, k === 0 ? 'TOTAL' : '', { font: font({ size: 11, bold: true }), al: LEFT, border: B1, fill: last ? fill('8497B0') : undefined });
          put(ws, y, 3, numOrStr(x.cal), { font: font({ size: 11 }), al: CEN, border: B1 });
          put(ws, y, 4, x.n, { font: font({ size: 11 }), al: { horizontal: 'right', vertical: 'middle' }, border: B1, fmt: '#,##0' });
          ws.getRow(y).height = 18; y++;
        });
        var lr = y - 1;
        ws.getCell(lr, 2).value = fx('SUM(D' + first + ':D' + lr + ')', g.tot);
        ws.getCell(lr, 2).numFmt = '#,##0'; ws.getCell(lr, 2).alignment = { horizontal: 'right', vertical: 'middle' };
        totCells.push('B' + lr);
        y += 2;
      });
      var under = font({ size: 11, bold: true, underline: true });
      [['Total tarimas:', (d.rows || []).filter(function (x) { return !x.empty; }).length], ['Total cajas:', fx(totCells.join('+') || '0', total)], ['Peso Neto', null], ['Peso Bruto:', null]].forEach(function (a, k) {
        var ry = y0 + 1 + k;
        merge(ws, ry, 5, ry, 6); border(ws, ry, 5, ry, 6, B1);
        put(ws, ry, 5, a[0], { font: under, al: CEN, border: B1 });
        merge(ws, ry, 7, ry, 8); border(ws, ry, 7, ry, 8, B1);
        put(ws, ry, 7, a[1], { font: font({ size: 11 }), al: LEFT, border: B1, fmt: '#,##0' });
      });
      endP = y - 2;
    }
    /* foto del tráiler a la derecha: cada fila de la foto (2 lugares) mide lo mismo que 2 filas de tarimas */
    if (TP) {
      var kk = 2 * RHp / TP.p, wp = TP.w * kk * 4 / 3, hp = TP.h * kk * 4 / 3, tp = rowTop(ws, r0) + RHp - TP.c0 * kk;
      ws.getColumn(9).width = 2; ws.getColumn(10).width = Math.ceil(wp / 7) + 2;
      imgAt(wb, ws, TP, 9, tp, wp, hp);
    }
    ws.pageSetup.printArea = 'A1:' + (TP ? 'J' : 'H') + endP;
    ws.pageSetup.printTitlesRow = '18:19';
    return ws;
  };

  /* ---------- v8.7 · Exportación de resultados financieros (Finanzas) ----------
     Cuatro hojas con fórmulas vivas, a partir de lo mismo que calcula la app:
     Resumen (indicadores), Productores (una fila por productor), Tamaños y calidades, y Viajes (un bloque por pedido,
     como FORMATO VIAJES: tamaños, segundas, Torreón, desecho, Juguera, importes y productores del viaje).
     Los importes que no se pueden calcular salen como «No disponible», nunca como cero. */
  var ND = 'No disponible';
  var MXN = '"$"#,##0.00;-"$"#,##0.00;"-"', USDF = '"US$"#,##0.00;-"US$"#,##0.00;"-"', PCT = '0.0%;-0.0%;"-"', KG = '#,##0.0;-#,##0.0;"-"';
  function rexHead(ws, d, title, ncol) {
    merge(ws, 1, 1, 1, ncol);
    put(ws, 1, 1, title, { font: font({ size: 15, bold: true }), al: LEFT });
    ws.getRow(1).height = 22;
    var info = [d.empresa + ' · ' + d.periodo, 'Generado: ' + d.generado + ' · Filtros: ' + d.filtros.join(' · '), d.unidades];
    info.forEach(function (t, i) { merge(ws, 2 + i, 1, 2 + i, ncol); put(ws, 2 + i, 1, t, { font: font({ size: 9, color: { argb: 'FF555555' } }), al: { horizontal: 'left', vertical: 'middle', wrapText: true } }); ws.getRow(2 + i).height = i === 1 ? 26 : 14; });
    return 6;
  }
  function headRow(ws, r, cols) {
    cols.forEach(function (c, i) { put(ws, r, 1 + i, c.h, { font: font({ size: 9, bold: true, color: { argb: 'FFFFFFFF' } }), fill: fill('1B2C46'), al: CEN, border: BOX }); });
    ws.getRow(r).height = 30;
  }
  /* escribe una tabla: cada columna da su valor o fórmula con c.v(fila, r, L) donde L(clave) es la letra de otra columna */
  function table(ws, r0, cols, rows, tot) {
    var idx = {}; cols.forEach(function (c, i) { idx[c.k] = i + 1; });
    var L = function (k) { return colL(idx[k]); };
    headRow(ws, r0, cols);
    var acc = cols.map(function () { return 0; });
    rows.forEach(function (row, i) {
      var r = r0 + 1 + i;
      cols.forEach(function (c, j) {
        var v = c.v(row, r, L);
        var nv = v && typeof v === 'object' ? v.result : v; if (typeof nv === 'number' && isFinite(nv)) acc[j] += nv;
        put(ws, r, 1 + j, v === undefined ? null : v, { font: font({ size: 9, bold: !!c.b }), al: c.txt ? LEFT : { horizontal: v === ND ? 'center' : 'right', vertical: 'middle' }, border: HBOX, fmt: c.fmt });
      });
    });
    var last = r0 + rows.length, tr = last + 1;
    if (tot) {
      cols.forEach(function (c, j) {
        var v = j === 0 ? 'TOTAL' : c.t ? c.t(r0 + 1, last, tr, L, acc[j]) : (c.sum ? fx('SUM(' + colL(1 + j) + (r0 + 1) + ':' + colL(1 + j) + last + ')', acc[j]) : null);
        put(ws, tr, 1 + j, v, { font: font({ size: 9, bold: true }), fill: fill('EDEDED'), al: j === 0 ? LEFT : { horizontal: v === ND ? 'center' : 'right', vertical: 'middle' }, border: BOX, fmt: c.fmt });
      });
    }
    return { idx: idx, L: L, first: r0 + 1, last: last, tot: tot ? tr : null };
  }
  function sumOf(rows, f) { var s = 0; rows.forEach(function (x) { s += +f(x) || 0; }); return s; }
  /* columnas de productor, según lo elegido (clasificación y tipo de resultado) */
  function prodCols(d, P) {
    var C = [], cl = d.cls, rs = d.res;
    var TT = P.T || d.T;
    if (P.viaje) C.push({ k: 'nom', h: 'PRODUCTOR', w: 40, txt: true, b: true, v: function (p) { return p.nombre + (p.codigo ? ' · ' + p.codigo : ''); } });
    else {
      C.push({ k: 'cod', h: 'CÓDIGO', w: 9, txt: true, v: function (p) { return p.codigo || ''; } });
      C.push({ k: 'nom', h: 'PRODUCTOR', w: 30, txt: true, b: true, v: function (p) { return p.nombre; } });
    }
    if (!P.viaje) C.push({ k: 'ped', h: 'PEDIDOS (VIAJES)', w: 9, fmt: NUM, sum: 1, v: function (p) { return p.ped; }, t: function () { return d.nViajes; } });
    if (rs.kg) {
      C.push({ k: 'kil', h: 'KILOS RECIBIDOS (kg)', w: 12, fmt: KG, sum: 1, v: function (p) { return p.kilos || null; } });
      C.push({ k: 'cam', h: 'CAJAS CAMPO (kilos ÷ ' + d.kgCampo + ')', w: 11, fmt: KG, sum: 1, v: function (p, r, L) { return p.kilos ? fx(L('kil') + r + '/' + P.kgCampo, p.kilos / d.kgCampo) : null; } });
    }
    if (rs.caj && cl.exp) {
      d.cal.forEach(function (c) { C.push({ k: 'c_' + c.k, h: c.l, w: 7.5, fmt: NUM, sum: 1, v: function (p) { return p.cal[c.k] || null; } }); });
      C.push({ k: 'pri', h: 'PRIMERA (cajas)', w: 10, fmt: NUM, b: true, sum: 1, v: function (p, r, L) { var a = L('c_' + d.cal[0].k), b = L('c_' + d.cal[d.cal.length - 1].k); return fx('SUM(' + a + r + ':' + b + r + ')', p.pri); } });
    }
    if (rs.caj) {
      [['seg', 'SEGUNDAS (cajas)'], ['tor', 'TORREÓN (cajas)'], ['des', 'DESECHO (cajas)'], ['jug', 'JUGUERA (cajas)']].forEach(function (q) {
        if (cl[q[0]]) C.push({ k: q[0], h: q[1], w: 9, fmt: NUM, sum: 1, v: function (p) { return p[q[0]] || null; } });
      });
    }
    if (rs.kg && rs.caj) {
      C.push({ k: 'kcl', h: 'KILOS CLASIFICADOS (kg)', w: 12, fmt: KG, sum: 1, v: function (p, r, L) {
        var parts = [];
        if (cl.exp) parts.push(L('pri') + r + '*' + P.kgExp);
        var o = ['seg', 'tor', 'des', 'jug'].filter(function (k) { return cl[k]; }).map(function (k) { return L(k) + r; });
        if (o.length) parts.push('(' + o.join('+') + ')*' + P.kgSeg);
        return parts.length ? fx(parts.join('+'), p.kcl) : null;
      } });
      if (cl.exp && cl.seg && cl.tor && cl.des && cl.jug && d.cal.length === d.calAll)
        C.push({ k: 'dif', h: 'DIFERENCIA (kg) clasificados − recibidos', w: 12, fmt: KG, sum: 1, v: function (p, r, L) { return p.kilos ? fx(L('kcl') + r + '-' + L('kil') + r, p.kcl - p.kilos) : null; } });
    }
    if (rs.rend && cl.exp) {
      C.push({ k: 'ren', h: 'RENDIMIENTO DE PRIMERAS', w: 10, fmt: PCT, v: function (p, r, L) {
        if (p.rend == null) return null;
        return (!p.ctn && rs.caj && rs.kg && d.cal.length === d.calAll) ? fx('IFERROR(' + L('pri') + r + '/' + L('cam') + r + ',"")', p.rend) : p.rend;
      }, t: function (a, b, tr, L) { return (!TT.ctn && rs.caj && rs.kg && d.cal.length === d.calAll) ? fx('IFERROR(' + L('pri') + tr + '/' + L('cam') + tr + ',"")', TT.rend || 0) : TT.rend; } });
    }
    if (rs.imp) {
      var impV = function (k) { return function (p) { return d.impOK ? (p[k] == null ? ND : p[k]) : ND; }; };
      C.push({ k: 'pag', h: 'PAGO AL PRODUCTOR (MXN)', w: 13, fmt: MXN, b: true, v: impV('pago') });
      C.push({ k: 'ing', h: 'INGRESO ATRIBUIDO (MXN)', w: 13, fmt: MXN, v: impV('ing') });
      ['pag', 'pj', 'ptot', 'ing'].forEach(function (k) { var c = C.filter(function (x) { return x.k === k; })[0]; if (c) c.imp = true; });
    }
    return C;
  }
  function prodTable(ws, r0, d, P) {
    var cols = prodCols(d, P);
    cols.forEach(function (c, i) { if (c.imp) c.t = function (a, b, tr, L, sm) { return d.impOK ? fx('SUM(' + colL(1 + i) + a + ':' + colL(1 + i) + b + ')', sm || 0) : ND; }; });
    return { cols: cols, t: table(ws, r0, cols, P.rows, true) };
  }
  /* parámetros editables arriba de cada hoja que los usa */
  function params(ws, r, d) {
    var rows = [['Kilos por caja campo (para cajas campo y rendimiento)', d.kgCampo], ['Kilos por caja de exportación (primera y segundas empacadas)', d.kgExp], ['Kilos por caja de segunda, Torreón, desecho y Juguera', d.kgSeg]];
    rows.forEach(function (x, i) {
      merge(ws, r + i, 1, r + i, 3); put(ws, r + i, 1, x[0], { font: font({ size: 9 }), al: LEFT, border: HBOX });
      put(ws, r + i, 4, x[1], { font: font({ size: 9, bold: true, color: { argb: 'FF1F4E79' } }), fill: fill('FFF2CC'), al: CEN, border: BOX, fmt: '0.0#' });
    });
    return { kgCampo: '$D$' + r, kgExp: '$D$' + (r + 1), kgSeg: '$D$' + (r + 2), next: r + 4 };
  }
  XL.addRex = function (wb, d) {
    var sheets = [];
    /* 1. Productores */
    var wp = wb.addWorksheet('Productores', { views: [{ showGridLines: false, state: 'frozen', ySplit: 10, xSplit: 2 }], pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: .3, right: .3, top: .4, bottom: .4, header: .1, footer: .1 } } });
    var nc = prodCols(d, { rows: d.prods }).length;
    var r = rexHead(wp, d, 'Resultados por productor', Math.max(nc, 6));
    var P = params(wp, r, d); P.rows = d.prods;
    var T = prodTable(wp, P.next, d, P);
    T.cols.forEach(function (c, i) { wp.getColumn(i + 1).width = c.w; });
    wp.autoFilter = { from: { row: P.next, column: 1 }, to: { row: T.t.last, column: T.cols.length } };
    wp.pageSetup.printArea = 'A1:' + colL(T.cols.length) + (T.t.tot + 1);
    wp.pageSetup.printTitlesRow = P.next + ':' + P.next;
    var sn = "'Productores'!";
    /* 2. Resumen (al frente): fórmulas sobre los totales de Productores */
    var wr = wb.addWorksheet('Resumen', { views: [{ showGridLines: false }], pageSetup: { orientation: 'portrait', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: .4, right: .4, top: .4, bottom: .4, header: .1, footer: .1 } } });
    [44, 16, 12].forEach(function (w, i) { wr.getColumn(i + 1).width = w; });
    r = rexHead(wr, d, d.titulo, 3);
    headRow(wr, r, [{ h: 'INDICADOR' }, { h: 'VALOR' }, { h: 'UNIDAD' }]);
    var tr = T.t.tot, Lp = T.t.L, R = [];
    var ref = function (k) { return T.t.idx[k] ? sn + Lp(k) + tr : null; };
    var add = function (l, k, val, u, f2) { var a = ref(k); if (!a && f2 == null) return; R.push([l, a ? fx(a, val) : f2, u, k]); };
    R.push(['Pedidos (viajes) incluidos', d.nViajes, 'pedidos']);
    R.push(['Productores incluidos', d.prods.length, 'productores']);
    add('Kilos recibidos', 'kil', d.T.kilos, 'kg');
    add('Cajas campo', 'cam', d.T.kilos / d.kgCampo, 'cajas campo');
    if (d.res.caj && d.cls.exp) d.cal.forEach(function (c) { add('Cajas de primera · ' + c.l, 'c_' + c.k, d.T.cal[c.k], 'cajas'); });
    add('Cajas de primera (total)', 'pri', d.T.pri, 'cajas');
    add('Segundas', 'seg', d.T.seg, 'cajas');
    add('Torreón', 'tor', d.T.tor, 'cajas');
    add('Desecho', 'des', d.T.des, 'cajas');
    add('Juguera', 'jug', d.T.jug, 'cajas');
    add('Kilos clasificados', 'kcl', d.T.kcl, 'kg');
    add('Diferencia: clasificados − recibidos', 'dif', d.T.kcl - d.T.kilos, 'kg');
    add('Rendimiento de primeras', 'ren', d.T.rend, '%');
    if (d.res.imp) {
      add('Pago de fruta a productores', 'pag', d.T.pago, 'MXN');
      R.push(['Venta de exportación', d.impOK ? d.T.usd : ND, 'USD']);
      R.push(['Ingresos (exportación en pesos + nacional)', d.impOK ? d.T.ing : ND, 'MXN']);
      R.push(['Costos (fruta, material, mano de obra, flete)', d.impOK ? d.T.cost : ND, 'MXN']);
      R.push(['Utilidad bruta', d.impOK ? fx('B' + (r + R.length - 1) + '-B' + (r + R.length), d.T.ing - d.T.cost) : ND, 'MXN']);
    }
    R.forEach(function (x, i) {
      var rr = r + 1 + i, u = x[2];
      put(wr, rr, 1, x[0], { font: font({ size: 10 }), al: LEFT, border: HBOX });
      put(wr, rr, 2, x[1], { font: font({ size: 10, bold: true }), al: { horizontal: x[1] === ND ? 'center' : 'right', vertical: 'middle' }, border: HBOX, fmt: u === 'MXN' ? MXN : u === 'USD' ? USDF : u === '%' ? PCT : u === 'kg' ? KG : NUM });
      put(wr, rr, 3, u, { font: font({ size: 9, color: { argb: 'FF555555' } }), al: CEN, border: HBOX });
    });
    var rEnd = r + R.length;
    if (d.notas && d.notas.length) d.notas.forEach(function (t, i) { merge(wr, rEnd + 2 + i, 1, rEnd + 2 + i, 3); put(wr, rEnd + 2 + i, 1, t, { font: font({ size: 9, italic: true, color: { argb: 'FF555555' } }), al: { horizontal: 'left', vertical: 'middle', wrapText: true } }); wr.getRow(rEnd + 2 + i).height = 26; });
    wr.pageSetup.printArea = 'A1:C' + (rEnd + 2 + (d.notas ? d.notas.length : 0));
    wr.orderNo = -1;
    /* 3. Tamaños y calidades */
    var wt = wb.addWorksheet('Tamaños y calidades', { views: [{ showGridLines: false }], pageSetup: { orientation: 'portrait', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: .4, right: .4, top: .4, bottom: .4, header: .1, footer: .1 } } });
    [26, 12, 12, 12, 11, 11, 13].forEach(function (w, i) { wt.getColumn(i + 1).width = w; });
    r = rexHead(wt, d, 'Tamaños y calidades', 7);
    var PT = params(wt, r, d); r = PT.next;
    if (d.cls.exp || d.cls.seg) {
      var TTam = sumOf(d.tam, function (x) { return (d.cls.exp ? x.a : 0) + (d.cls.seg ? x.b : 0); });
      var tc = [{ k: 'l', h: 'CALIBRE', txt: true, b: true, v: function (x) { return x.l; } }, { k: 'a', h: 'PRIMERA (cajas)', fmt: NUM, sum: 1, v: function (x) { return d.cls.exp ? x.a || null : null; } }, { k: 'b', h: 'SEGUNDAS EMPACADAS (cajas)', fmt: NUM, sum: 1, v: function (x) { return d.cls.seg ? x.b || null : null; } },
        { k: 't', h: 'TOTAL (cajas)', fmt: NUM, b: true, sum: 1, v: function (x, rr, L) { return fx(L('a') + rr + '+' + L('b') + rr, (d.cls.exp ? x.a : 0) + (d.cls.seg ? x.b : 0)); } },
        { k: 'p', h: '% DEL TOTAL', fmt: PCT, v: function (x, rr, L) { return fx('IFERROR(' + L('t') + rr + '/' + L('t') + '$' + (r + 1 + d.tam.length) + ',0)', TTam ? ((d.cls.exp ? x.a : 0) + (d.cls.seg ? x.b : 0)) / TTam : 0); }, t: function () { return 1; } },
        { k: 'kg', h: 'KG POR CAJA', fmt: '0.0#', v: function () { return fx(PT.kgExp, d.kgExp); } },
        { k: 'kil', h: 'KILOS (kg)', fmt: KG, sum: 1, v: function (x, rr, L) { return fx(L('t') + rr + '*' + L('kg') + rr, ((d.cls.exp ? x.a : 0) + (d.cls.seg ? x.b : 0)) * d.kgExp); } }];
      var t1 = table(wt, r, tc, d.tam, true); r = t1.tot + 2;
    }
    var nac = [['seg', 'Segundas'], ['tor', 'Torreón'], ['des', 'Desecho'], ['jug', 'Juguera']].filter(function (q) { return d.cls[q[0]]; }).map(function (q) { return { l: q[1], n: d.T[q[0]] }; });
    if (nac.length) {
      var tn = [{ k: 'l', h: 'CLASIFICACIÓN', txt: true, b: true, v: function (x) { return x.l; } }, { k: 'n', h: 'CAJAS', fmt: NUM, sum: 1, v: function (x) { return x.n || null; } }, { k: 'kg', h: 'KG POR CAJA', fmt: '0.0#', v: function () { return fx(PT.kgSeg, d.kgSeg); } }, { k: 'kil', h: 'KILOS (kg)', fmt: KG, sum: 1, v: function (x, rr, L) { return fx(L('n') + rr + '*' + L('kg') + rr, (x.n || 0) * d.kgSeg); } }];
      var t2 = table(wt, r, tn, nac, true); r = t2.tot + 2;
    }
    put(wt, r, 1, 'Las cajas de exportación y las de segunda, Torreón, desecho y Juguera no se suman entre sí: son cajas de distinto tamaño. Los kilos sí se pueden sumar.', { font: font({ size: 9, italic: true, color: { argb: 'FF555555' } }), al: { horizontal: 'left', vertical: 'middle', wrapText: true } });
    merge(wt, r, 1, r, 7); wt.getRow(r).height = 26;
    wt.pageSetup.printArea = 'A1:G' + r;
    /* 4. Viajes: un bloque por pedido, como FORMATO VIAJES */
    var wv = wb.addWorksheet('Viajes', { views: [{ showGridLines: false }], pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: .3, right: .3, top: .4, bottom: .4, header: .1, footer: .1 } } });
    var vcols = prodCols(d, { rows: [], viaje: true }), wmax = Math.max(vcols.length, 8);
    r = rexHead(wv, d, 'Detalle por viaje (pedido)', wmax);
    var PV = params(wv, r, d); r = PV.next; PV.viaje = true;
    d.viajes.forEach(function (v, vi) {
      merge(wv, r, 1, r, wmax);
      put(wv, r, 1, 'VIAJE ' + (vi + 1) + ' · ' + v.cod + ' · ' + v.emb + (v.po ? ' · PO ' + v.po : '') + ' · ' + v.fechaTxt + (v.cliente ? ' · ' + v.cliente : '') + ' · ' + v.empresa, { font: font({ size: 11, bold: true, color: { argb: 'FFFFFFFF' } }), fill: fill('2F5233'), al: LEFT });
      wv.getRow(r).height = 20; r += 1;
      var tv = [{ k: 'l', h: 'TAMAÑO', txt: true, b: true, v: function (x) { return x.l; } }, { k: 'a', h: 'PRIMERA (cajas)', fmt: NUM, sum: 1, v: function (x) { return x.a || null; } }, { k: 'b', h: 'SEGUNDAS EMPACADAS (cajas)', fmt: NUM, sum: 1, v: function (x) { return x.b || null; } },
        { k: 't', h: 'TOTAL (cajas)', fmt: NUM, b: true, sum: 1, v: function (x, rr, L) { return fx(L('a') + rr + '+' + L('b') + rr, x.a + x.b); } }, { k: 'kg', h: 'KG POR CAJA', fmt: '0.0#', v: function () { return v.kg; } }, { k: 'kil', h: 'KILOS (kg)', fmt: KG, sum: 1, v: function (x, rr, L) { return fx(L('t') + rr + '*' + L('kg') + rr, (x.a + x.b) * v.kg); } }];
      var tt = table(wv, r, tv, v.tam, true); r = tt.tot + 1;
      var other = [['SEGUNDAS', v.seg, 'seg'], ['TORREÓN', v.tor, 'tor'], ['DESECHO', v.des, 'des'], ['JUGUERA', v.jug, 'jug']].filter(function (q) { return d.cls[q[2]]; });
      other.forEach(function (q) {
        put(wv, r, 1, q[0], { font: font({ size: 9, bold: true }), border: HBOX }); put(wv, r, 4, q[1] || null, { font: font({ size: 9 }), al: { horizontal: 'right' }, border: HBOX, fmt: NUM });
        put(wv, r, 5, fx(PV.kgSeg, d.kgSeg), { font: font({ size: 9 }), border: HBOX, fmt: '0.0#' }); put(wv, r, 6, fx('D' + r + '*E' + r, (q[1] || 0) * d.kgSeg), { font: font({ size: 9 }), border: HBOX, fmt: KG }); r++;
      });
      if (d.res.imp) {
        var e0 = r + 1, E = [['Tipo de cambio (MXN por 1 USD)', v.tc, '0.00##'], ['Venta de exportación (USD)', v.usd, USDF], ['Exportación en pesos (MXN)', null, MXN, 'B' + (e0 + 1) + '*B' + e0, v.usd * (v.tc || 0)], ['Venta nacional: segundas, Torreón y Juguera (MXN)', v.nac, MXN],
          ['Ingresos (MXN)', null, MXN, 'B' + (e0 + 2) + '+B' + (e0 + 3), v.ing], ['Compra de fruta a productores (MXN)', v.fruta, MXN], ['Material de empaque (MXN)', v.mat, MXN], ['Mano de obra (MXN)', v.mo, MXN], ['Flete y cruce (MXN)', v.flete, MXN],
          ['Costos (MXN)', null, MXN, 'SUM(B' + (e0 + 5) + ':B' + (e0 + 8) + ')', v.cost], ['Utilidad bruta (MXN)', null, MXN, 'B' + (e0 + 4) + '-B' + (e0 + 9), v.ing - v.cost], ['Margen', null, PCT, 'IFERROR(B' + (e0 + 10) + '/B' + (e0 + 4) + ',0)', v.ing ? (v.ing - v.cost) / v.ing : 0]];
        E.forEach(function (x, i) {
          var rr = e0 + i, bold = /Ingresos|Costos|Utilidad/.test(x[0]);
          put(wv, rr, 1, x[0], { font: font({ size: 9, bold: bold }), border: HBOX });
          put(wv, rr, 2, d.impOK ? (x[3] ? fx(x[3], x[4]) : x[1]) : ND, { font: font({ size: 9, bold: bold }), al: { horizontal: d.impOK ? 'right' : 'center' }, border: HBOX, fmt: x[2] });
        });
        r = e0 + E.length + 1;
      }
      if (v.prods.length) { var pv = prodTable(wv, r, d, { rows: v.prods, viaje: true, T: v.T, kgCampo: PV.kgCampo, kgExp: PV.kgExp, kgSeg: PV.kgSeg }); r = pv.t.tot + 2; }
      else r += 1;
    });
    vcols.forEach(function (c, i) { wv.getColumn(i + 1).width = Math.max(c.w, i < 6 ? 12 : 0); });
    wv.getColumn(1).width = 42;
    wv.pageSetup.printArea = 'A1:' + colL(wmax) + r;
    return [wr, wp, wt, wv];
  };

  /* ============ CIERRE SEMANAL · calca de FORMATO VIAJES ============
     Una hoja por semana y empresa; cada viaje es un bloque completo de FORMATO VIAJES (bloque del viaje a la izquierda y un bloque por productor
     a la derecha, con el color de su clasificación). Los precios por clasificación y los datos fijos están en la hoja PRECIOS de cada semana.
     Todo son fórmulas vivas; las casillas amarillas se pueden cambiar. */
  function calcWb(wb) {
    var memo = {}, busy = {};
    function colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
    function num(v) { return typeof v === 'number' ? v : (typeof v === 'boolean' ? (v ? 1 : 0) : (parseFloat(v) || 0)); }
    function cellVal(sn, ref) {
      var key = sn + '!' + ref, out;
      if (key in memo) { out = memo[key]; if (out instanceof Error) throw out; return out; }
      if (busy[key]) throw new Error('#CIRC');
      var ws = wb.getWorksheet(sn); if (!ws) throw new Error('#REF!');
      var c = ws.getCell(ref), v = c.value;
      if (v && typeof v === 'object' && v.formula !== undefined) {
        busy[key] = 1;
        try { out = ev(parse(v.formula), sn); if (typeof out === 'number' && !isFinite(out)) out = new Error('#NUM!'); } catch (e) { out = e instanceof Error ? e : new Error('#VALUE!'); }
        busy[key] = 0; c.value = { formula: v.formula, result: out instanceof Error ? 0 : out };
      } else out = (v == null ? 0 : (v instanceof Date ? Math.round(v.getTime() / 864e5 + 25569) : v));
      memo[key] = out; if (out instanceof Error) throw out; return out;
    }
    function parse(s) {
      var i = 0;
      function sp() { while (s.charAt(i) === ' ') i++; }
      function m(re) { re.lastIndex = i; var r = re.exec(s); if (r && r.index === i) { i += r[0].length; return r; } return null; }
      function primary() {
        sp(); var r, c = s.charAt(i);
        if (c === '(') { i++; var e = expr(); sp(); i++; return e; }
        if (c === '-') { i++; return { t: 'u', e: primary() }; }
        if (c === '"') { r = m(/"([^"]*)"/y); return { t: 's', v: r[1] }; }
        if ((r = m(/\d+(\.\d+)?/y))) return { t: 'n', v: parseFloat(r[0]) };
        if ((r = m(/([A-Z]+)\(/y))) {
          var args = []; sp();
          if (s.charAt(i) === ')') i++; else { for (;;) { args.push(expr()); sp(); if (s.charAt(i) === ',') { i++; continue; } i++; break; } }
          return { t: 'f', n: r[1], a: args };
        }
        var sn = null;
        if ((r = m(/'((?:[^']|'')+)'!/y))) sn = r[1].replace(/''/g, "'");
        if ((r = m(/\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/y))) {
          if (r[3]) return { t: 'g', sn: sn, c1: colNum(r[1]), r1: +r[2], c2: colNum(r[3]), r2: +r[4] };
          return { t: 'r', sn: sn, ref: r[1] + r[2] };
        }
        throw new Error('#NAME?');
      }
      function term() { var v = primary(); for (;;) { sp(); var c = s.charAt(i); if (c === '*' || c === '/') { i++; v = { t: 'b', o: c, l: v, r: primary() }; } else return v; } }
      function expr() { var v = term(); for (;;) { sp(); var c = s.charAt(i); if (c === '+' || c === '-') { i++; v = { t: 'b', o: c, l: v, r: term() }; } else return v; } }
      return expr();
    }
    function flat(n, cur) {
      if (n.t === 'g') { var o = [], sn = n.sn || cur; for (var r = n.r1; r <= n.r2; r++) for (var c = n.c1; c <= n.c2; c++) { try { o.push(cellVal(sn, colL(c) + r)); } catch (e) { throw e; } } return o; }
      return [ev(n, cur)];
    }
    function ev(n, cur) {
      switch (n.t) {
        case 'n': case 's': return n.v;
        case 'u': return -num(ev(n.e, cur));
        case 'r': return cellVal(n.sn || cur, n.ref);
        case 'b': { var l = num(ev(n.l, cur)), r = num(ev(n.r, cur)); if (n.o === '+') return l + r; if (n.o === '-') return l - r; if (n.o === '*') return l * r; if (r === 0) throw new Error('#DIV/0!'); return l / r; }
        case 'f': {
          if (n.n === 'IFERROR') { try { var v = ev(n.a[0], cur); if (typeof v === 'number' && !isFinite(v)) throw 0; return v; } catch (e) { return ev(n.a[1], cur); } }
          if (n.n === 'SUM') { var t = 0; n.a.forEach(function (x) { flat(x, cur).forEach(function (v) { if (typeof v === 'number') t += v; }); }); return t; }
          /* v8.39: ROUND como Excel (la mitad se redondea alejándose de cero) */
          if (n.n === 'ROUND') { var x0 = num(ev(n.a[0], cur)), d0 = n.a[1] ? num(ev(n.a[1], cur)) : 0, p0 = Math.pow(10, d0); return (x0 < 0 ? -1 : 1) * Math.round(Math.abs(x0) * p0 + 1e-9) / p0; }
          throw new Error('#NAME?');
        }
      }
      throw new Error('#VALUE!');
    }
    wb.worksheets.forEach(function (ws) { ws.eachRow(function (row) { row.eachCell(function (c) { var v = c.value; if (v && typeof v === 'object' && v.formula !== undefined) { try { cellVal(ws.name, c.address); } catch (e) { } } }); }); });
  }

  XL.addCierre = function (wb, C) {
    var FNT = 'Arial Nova Cond Light';
    var USD = '"$"#,##0.00', INT = '#,##0', D2 = '0.00', PC = '0%';
    var COL = { chico: '8ED973', grande: '00FFFF', interno: 'FF66FF' };
    var CLN = { chico: 'CHICOS', grande: 'GRANDES', interno: 'INTERNOS' };
    var CK = ['c250', 'c230', 'c200', 'c175', 'c150', 'c110'], CKL = [250, 230, 200, 175, 150, 110];
    var LB = 'C1E5F5', LP = 'F2CEEF', LG = 'D9F2D0', DG = '3A7D22', YE = 'FFFF00', CR = 'FFF2CC';
    function q(n) { return "'" + String(n).replace(/'/g, "''") + "'"; }
    function T(v) { return typeof v === 'string' ? { __t: v } : v; }
    function W(ws, r, c, v, o) {
      o = o || {};
      var cell = ws.getCell(r, c);
      /* v8.30.1: lo que viene de la captura (nombres, cliente, códigos) va marcado con T() y se escribe siempre como texto, aunque empiece con = */
      if (v && typeof v === 'object' && v.__t !== undefined) { if (v.__t !== '') cell.value = v.__t; }
      else if (v !== undefined && v !== null && v !== '') cell.value = (typeof v === 'string' && v.charAt(0) === '=') ? { formula: v.slice(1), result: 0 } : v;
      var f = { name: FNT, size: o.sz || 14, bold: o.b !== false };
      if (o.color) f.color = { argb: 'FF' + o.color };
      if (o.i) f.italic = true;
      cell.font = f;
      if (o.fill) cell.fill = fill(o.fill);
      if (o.fmt) cell.numFmt = o.fmt;
      cell.alignment = { horizontal: o.al, vertical: 'middle', wrapText: !!o.wrap, shrinkToFit: !!o.shrink };
      return cell;
    }
    function edge(ws, r1, c1, r2, c2, st) {
      var s = { style: st || 'thin', color: { argb: 'FF000000' } };
      for (var r = r1; r <= r2; r++) for (var c = c1; c <= c2; c++) {
        var cell = ws.getCell(r, c), b = Object.assign({}, cell.border || {});
        if (r === r1) b.top = s; if (r === r2) b.bottom = s; if (c === c1) b.left = s; if (c === c2) b.right = s;
        cell.border = b;
      }
    }
    var wsRes = wb.addWorksheet('RESUMEN SEMANAL', { views: [{ showGridLines: false }] });
    var wsPag = wb.addWorksheet('PAGOS PRODUCTORES', { views: [{ showGridLines: false, state: 'frozen', ySplit: 4 }] });
    var tripRefs = [], blocksAll = [];

    C.groups.forEach(function (G) {
      var pn = 'PRECIOS ' + G.tag, P = G.P;
      /* ---- hoja PRECIOS ---- */
      var wp = wb.addWorksheet(pn, { views: [{ showGridLines: false }] });
      [38, 16, 16, 16].forEach(function (w, i) { wp.getColumn(i + 1).width = w; });
      W(wp, 1, 1, 'PRECIOS Y DATOS FIJOS · ' + G.empName + ' · ' + G.label + (G.span ? ' (' + G.span + ')' : ''), { sz: 16 });
      W(wp, 2, 1, 'Los viajes de la hoja ' + G.name + ' leen estas casillas. Cambia una casilla amarilla y todo el cierre se recalcula. Para cambiar un solo viaje, escribe encima de la casilla de ese viaje.', { sz: 10, b: false, i: true, color: '555555' });
      W(wp, 4, 1, 'PAGO A PRODUCTORES POR CLASIFICACIÓN', { sz: 12, fill: 'EDEDED' });
      ['chico', 'grande', 'interno'].forEach(function (k, i) { W(wp, 4, 2 + i, CLN[k], { sz: 12, fill: COL[k], al: 'center' }); });
      var CF = [['p250', 'Calibre 250 (MXN por caja)'], ['p230', 'Calibre 230 (MXN por caja)'], ['p200', 'Calibre 200 (MXN por caja)'], ['p175', 'Calibre 175 (MXN por caja)'], ['p150', 'Calibre 150 (MXN por caja)'], ['p110', 'Calibre 110 (MXN por caja)'], ['seg', 'Segundas (MXN por caja)'], ['tor', "Torreón 275' (MXN por caja)"], ['des', 'Desecho (MXN por caja)'], ['pkg', 'Pago por kilo de báscula (MXN por kg)']];
      var cfRow = {}, colOf = { chico: 2, grande: 3, interno: 4 };
      CF.forEach(function (f, i) {
        var r = 5 + i; cfRow[f[0]] = r;
        W(wp, r, 1, f[1], { sz: 12, b: false, fill: 'FFFFFF' });
        ['chico', 'grande', 'interno'].forEach(function (k) {
          var v = P.clas[k][f[0]];
          if (v === undefined) W(wp, r, colOf[k], '—', { sz: 12, b: false, al: 'center', color: '999999' });
          else W(wp, r, colOf[k], v, { sz: 12, fill: YE, fmt: USD, al: 'center' });
        });
      });
      edge(wp, 4, 1, 4 + CF.length, 4, 'thin');
      var FX = [['tc', 'Tipo de cambio (MXN por USD)', P.tc, '0.00##'], ['kgC', 'Kilos por caja de campo', P.kgC, '0.0#'], ['kgE', 'Kilos por caja de exportación', P.kgE, '0.0#'], ['kgS', 'Kilos por caja de segunda, Torreón, desecho y Juguera', P.kgS, '0.0#']]
        .concat([['aduAm', 'Aduana americana (USD por viaje)'], ['aduMx', 'Aduana mexicana (MXN por viaje)'], ['aduCaja', 'Aduana por caja (USD)'], ['factPri', 'Factura primeras (MXN por caja)'], ['fact15', 'Factura segundas (MXN por caja)'], ['retorno', 'Plástico / retorno (MXN por caja)'], ['gastos', 'Gastos: maquila y seguro (MXN por caja)'], ['seguro', 'Seguro de carga (MXN por caja)'], ['dscFlete', 'Descuento fletes (USD por caja)']].map(function (x) { return [x[0], x[1], P.cierre[x[0]], USD]; }));
      var fx0 = 5 + CF.length + 2, fxRow = {};
      W(wp, fx0 - 1, 1, 'DATOS FIJOS DEL VIAJE (los fija el contador)', { sz: 12, fill: 'EDEDED' }); W(wp, fx0 - 1, 2, '', { fill: 'EDEDED' });
      FX.forEach(function (f, i) { var r = fx0 + i; fxRow[f[0]] = r; W(wp, r, 1, f[1], { sz: 12, b: false }); W(wp, r, 2, f[2], { sz: 12, fill: YE, fmt: f[3], al: 'center' }); });
      edge(wp, fx0 - 1, 1, fx0 + FX.length - 1, 2, 'thin');
      var PR = {
        c: function (k, f) { return q(pn) + '!$' + colL(colOf[k]) + '$' + cfRow[f]; },
        f: function (f) { return q(pn) + '!$B$' + fxRow[f]; }
      };
      /* v8.11: pago a productores grandes por entrega (kilos de báscula × precio de la entrega); el bloque del productor lee de aquí */
      var grRows = {}, GE = [];
      G.trips.forEach(function (t) { t.prods.forEach(function (p) { if (p.cls === 'grande' && p.ents && p.ents.length) GE.push({ t: t, p: p }); }); });
      if (GE.length) {
        var g0 = fx0 + FX.length + 2, gr = g0;
        W(wp, g0 - 1, 1, 'PAGO A PRODUCTORES GRANDES POR ENTREGA', { sz: 12, fill: 'EDEDED' });
        ['Fecha de entrega', 'Kilos de báscula', 'Precio (MXN por kg)', 'Pago (MXN)', 'Anticipo (MXN)', 'Precio de'].forEach(function (x, i) { W(wp, g0 - 1, 2 + i, x, { sz: 10, fill: 'EDEDED', al: 'center', wrap: true }); });
        wp.getRow(g0 - 1).height = 30;
        [16, 16, 16, 16, 16, 26].forEach(function (w0, i) { if ((wp.getColumn(2 + i).width || 0) < w0) wp.getColumn(2 + i).width = w0; });
        GE.forEach(function (x) {
          var r1 = gr;
          x.p.ents.forEach(function (e) {
            W(wp, gr, 1, 'Viaje ' + x.t.n + ' · ' + x.t.cod + ' · ' + x.p.nombre, { sz: 11, b: false });
            W(wp, gr, 2, e.fecha ? new Date(e.fecha + 'T12:00:00Z') : '', { sz: 11, b: false, fmt: 'dd/mm/yyyy', al: 'center' });
            W(wp, gr, 3, e.kilos, { sz: 11, b: false, fmt: '#,##0.##', al: 'center' });
            var lnk = e.same && (e.src === 'clas' || e.src === 'general' || e.src === 'sug');
            W(wp, gr, 4, lnk ? '=' + PR.c('grande', 'pkg') : e.pkg, { sz: 11, fill: YE, fmt: USD, al: 'center' });
            W(wp, gr, 5, '=C' + gr + '*D' + gr, { sz: 11, fmt: USD, al: 'center' });
            W(wp, gr, 6, e.ant || 0, { sz: 11, fill: YE, fmt: USD, al: 'center' });
            W(wp, gr, 7, e.src === 'entrega' ? 'esta entrega' : e.src === 'acuerdo' || e.src === 'hoja' ? 'acuerdo de compra' : lnk ? 'casilla GRANDES de arriba' : 'semana ' + (/S(\d+)$/.exec(e.wk || '') || ['', '?'])[1], { sz: 10, b: false, color: '555555' });
            gr++;
          });
          grRows[x.t.k + '|' + x.p.key] = { a: r1, b: gr - 1 };
        });
        W(wp, gr, 1, 'TOTAL', { sz: 12 }); W(wp, gr, 3, '=SUM(C' + (g0) + ':C' + (gr - 1) + ')', { sz: 12, fmt: '#,##0.##', al: 'center' });
        W(wp, gr, 5, '=SUM(E' + g0 + ':E' + (gr - 1) + ')', { sz: 12, fmt: USD, al: 'center' }); W(wp, gr, 6, '=SUM(F' + g0 + ':F' + (gr - 1) + ')', { sz: 12, fmt: USD, al: 'center' });
        edge(wp, g0 - 1, 1, gr, 7, 'thin');
      }
      var grRng = function (t, p, col) { var z = grRows[t.k + '|' + p.key]; return z ? q(pn) + '!$' + col + '$' + z.a + ':$' + col + '$' + z.b : null; };

      /* ---- hoja SEMANA ---- */
      var ws = wb.addWorksheet(G.name, { views: [{ showGridLines: false, zoomScale: 70 }], pageSetup: { orientation: 'landscape', paperSize: 8, fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
      var maxNb = 1; G.trips.forEach(function (t) { maxNb = Math.max(maxNb, t.prods.length); });
      /* v8.28: columnas por tipo de empaque del manifiesto (solo si el viaje trae más de uno); se insertan antes del total */
      var nTp = function (t) { var n = (t.tiposMan || []).length; return n > 1 ? n : 0; };
      var Smax = 0; G.trips.forEach(function (t) { Smax = Math.max(Smax, nTp(t)); });
      ws.getColumn(1).width = 3; ws.getColumn(2).width = 33.6;
      for (var sc = 0; sc < Smax; sc++) ws.getColumn(3 + sc).width = 13;
      [19, 15, 19, 20, 17, 15].forEach(function (w, i) { ws.getColumn(3 + Smax + i).width = w; });
      for (var b = 0; b < maxNb; b++) { [15, 16, 15, 20, 17, 17, 21].forEach(function (w, i) { ws.getColumn(9 + Smax + 7 * b + i).width = w; }); }
      var PERIOD = 38, RTT = 27, RKR = 28, RMQ = 33; /* v8.10: filas del bloque de productor (total pagado, kilos, cajas maquiladas) */
      G.trips.forEach(function (t, ti) {
        var o = ti * PERIOD, nb = t.prods.length, lastC = 8 + 7 * Math.max(nb, 1), S = nTp(t);
        var R = function (r) { return r + o; };
        var a = function (c, r) { return colL(c) + R(r); };
        /* v8.28: todo se escribe como si no hubiera columnas por tipo; shC/shF recorren S columnas lo que va de la C en adelante */
        var shC = function (c) { return c >= 3 ? c + S : c; };
        var shF = function (v) { return S ? shiftF(v, S) : v; };
        var w = function (r, c, v, op) { return W(ws, R(r), shC(c), shF(v), op); };
        var edgeS = function (r1, c1, r2, c2, st) { edge(ws, r1, shC(c1), r2, shC(c2), st); };
        var h = function (r, ht) { ws.getRow(R(r)).height = ht; };
        for (var r0 = 2; r0 <= 35; r0++) h(r0, 20);
        h(2, 60.6); h(5, 23.4); h(31, 35.4);
        var blocks = t.prods.map(function (p, k) { return { p: p, k: k, c0: 9 + 7 * k, mode: p.cls === 'grande' ? 'kg' : 'box' }; });
        var c15r = function (bl) { return bl.mode === 'kg' ? 12 : 10; };
        var sumB = function (fn) { return blocks.length ? 'SUM(' + blocks.map(fn).join(',') + ')' : '0'; };
        /* título */
        for (var c = 2; c <= lastC + S; c++) ws.getCell(R(2), c).fill = fill('33CC33');
        w(2, 2, ' VIAJE ', { sz: 48, fill: '33CC33' }); w(2, 3, t.n, { sz: 48, fill: '33CC33', al: 'center' }); w(2, 5, '=B' + R(5), { sz: 48, fill: '33CC33' });
        /* encabezado */
        var dt = t.fecha ? new Date(t.fecha + 'T12:00:00Z') : null;
        w(4, 2, dt, { sz: 12, fmt: 'd" de "mmmm" de "yyyy', al: 'left' });
        w(4, 6, t.po ? 'P.O. ' + t.po : '', { sz: 12 });
        w(4, 8, '=B' + R(5), { sz: 12, fill: 'FF33CC', al: 'center' });
        w(5, 2, T(t.cod), { sz: 16 }); w(5, 3, T(t.emb), { sz: 16, fill: 'FF2F92', al: 'center' }); w(5, 5, T(t.cliente), { sz: 16 }); w(5, 8, T(t.pack), { sz: 14, fill: YE, al: 'center' });
        /* tamaños */
        CK.forEach(function (ck, i) {
          var r = 6 + i;
          w(r, 2, CKL[i], { al: 'right' });
          w(r, 3, t.sizes[ck].n || 0, { al: 'center' }); /* v8.39: lo manifestado (todas las ventas del viaje), no lo registrado */
          w(r, 4, 0, { b: false, al: 'center' });
          w(r, 5, '=C' + R(r) + '-D' + R(r), { al: 'center', fmt: '0' });
          w(r, 6, Math.round(t.sizes[ck].p * 10000) / 10000, { al: 'center', fill: CR, fmt: USD });
          w(r, 7, '=E' + R(r) + '*F' + R(r), { b: false, al: 'center', fmt: USD });
        });
        w(12, 3, '=SUM(C' + R(6) + ':C' + R(11) + ')', { al: 'center', color: 'FF0000' }); w(12, 4, '=SUM(D' + R(6) + ':D' + R(11) + ')', { al: 'center' });
        w(12, 5, '=SUM(E' + R(6) + ':E' + R(11) + ')', { al: 'center', fmt: '0' }); w(12, 7, '=SUM(G' + R(6) + ':G' + R(11) + ')', { al: 'center', fmt: USD });
        edgeS(R(6), 2, R(12), 7, 'thin');
        /* v8.28: manifiesto por tipo de empaque (RPC, CTN, CTN 10 lb, PLT) */
        var TM = t.tiposMan || [];
        if (S) {
          W(ws, R(4), 3, 'POR TIPO DE EMPAQUE', { sz: 11, al: 'left' }); W(ws, R(4), 3 + S, 'TOTAL', { sz: 11, al: 'center' });
          TM.forEach(function (tp, j) {
            var c = 3 + j, L = colL(c);
            W(ws, R(5), c, tp.l, { sz: 14, fill: 'FFD966', al: 'center' });
            CK.forEach(function (ck, i) { W(ws, R(6 + i), c, tp.cal[ck] || 0, { al: 'center' }); });
            W(ws, R(12), c, '=SUM(' + L + R(6) + ':' + L + R(11) + ')', { al: 'center', color: 'FF0000' });
            W(ws, R(13), c, tp.c15 || 0, { al: 'center', color: '00B0F0' });
          });
          edge(ws, R(5), 3, R(13), 2 + S, 'thin');
        } else if (TM.length === 1) W(ws, R(4), 3, TM[0].l, { sz: 12, al: 'center', fill: 'FFD966' });
        w(13, 2, 'SEGUNDAS'); w(13, 3, t.c15.n || 0, { al: 'center', color: '00B0F0' });
        w(13, 5, '=H' + R(10) + '*G' + R(17), { al: 'center', fmt: USD }); w(13, 6, Math.round(t.c15.p * 10000) / 10000, { al: 'center', fill: CR, fmt: USD }); w(13, 7, '=C' + R(12), { al: 'center' });
        /* ventas en dólares (H8:H12) */
        w(8, 8, '=G' + R(12), { fill: YE, fmt: USD, al: 'center' }); w(9, 8, '=C' + R(13) + '*F' + R(13), { b: false, fmt: USD, al: 'center' });
        w(10, 8, '=SUM(H' + R(8) + ':H' + R(9) + ')', { fmt: USD, al: 'center', color: 'FF0000' }); w(12, 8, '=IFERROR(H' + R(10) + '/C' + R(14) + ',0)', { fmt: USD, al: 'center', color: 'FF0000' });
        w(7, 8, 'VENTA USD', { sz: 11, al: 'center' });
        /* cajas / flete / aduanas */
        w(14, 2, 'TOTAL CAJAS', { fill: LP }); w(14, 3, '=C' + R(12) + '+C' + R(13) + '+D' + R(13), { fill: LP, color: 'FF0000', al: 'center', fmt: '0.#' }); w(14, 4, '', { fill: LP });
        w(14, 5, '=C' + R(20) + '*G' + R(17), { al: 'center', fmt: USD, color: 'FF0000' }); w(14, 6, 'Aduanas', { sz: 12 }); w(14, 7, '=IFERROR(G' + R(12) + '/G' + R(13) + ',0)', { b: false, fmt: USD, al: 'center' });
        w(15, 2, 'FLETE', { fill: LP }); w(15, 3, t.flete, { fill: LP, fmt: USD, color: '00B0F0', al: 'center' });
        w(15, 5, '=C' + R(14) + '*G' + R(20), { fmt: USD, al: 'center' }); w(15, 6, 'fletes', { sz: 12 }); w(15, 7, '=' + PR.f('dscFlete'), { b: false, fmt: USD, al: 'center' }); w(15, 8, 'ADUANA', { sz: 12 });
        w(16, 2, '', { fill: LP }); w(16, 3, '=C' + R(14), { fill: LP, al: 'center', fmt: '0' });
        w(16, 5, '=C' + R(14) + '*G' + R(19), { fmt: USD, al: 'center', color: 'FF0000' }); w(16, 6, 'carton', { sz: 12 }); w(16, 7, '=G' + R(14) + '-G' + R(15), { fmt: USD, al: 'center' });
        w(17, 2, '', { fill: LP }); w(17, 3, '=IFERROR(C' + R(15) + '/C' + R(16) + ',0)', { fill: LP, fmt: USD, al: 'center' });
        w(17, 5, '=C' + R(14) + '*G' + R(22), { fmt: USD, al: 'center' }); w(17, 6, 'MAQ. Y SEGURO', { sz: 12, fill: 'F837F4' }); w(17, 7, '=' + PR.f('tc'), { fill: '92D050', fmt: USD, al: 'center' }); w(17, 8, 'T.CAMBIO', { sz: 12, fill: '92D050' });
        w(18, 2, 'ADUANA AMERICANA'); w(18, 3, '=' + PR.f('aduAm'), { fmt: USD, al: 'center' }); w(18, 5, '=E' + R(13) + '-E' + R(14) + '-E' + R(15) + '-E' + R(16) + '-E' + R(17), { fmt: USD, al: 'center', color: 'FF0000' }); w(18, 7, '=G' + R(16) + '*G' + R(17), { fmt: USD, al: 'center' });
        w(19, 2, 'A. MEXICANA #1'); w(19, 3, '=IFERROR(' + PR.f('aduMx') + '/G' + R(17) + ',0)', { fmt: USD, al: 'center', color: 'FF0000' });
        for (var c2 = 4; c2 <= 6; c2++) ws.getCell(R(19), shC(c2)).fill = fill('99FF66');
        w(19, 4, 0, { fill: '99FF66', al: 'center' }); w(19, 7, '=' + PR.f('retorno'), { fmt: USD, al: 'center' }); w(19, 8, 'Retorno', { sz: 12 });
        w(20, 3, '=SUM(C' + R(18) + ':C' + R(19) + ')', { fmt: D2, al: 'center' });
        w(20, 5, '=' + sumB(function (bl) { return a(bl.c0 + 4, RTT); }), { fmt: USD, al: 'center' }); w(20, 7, '=C' + R(17), { b: false, fmt: USD, al: 'center' }); w(20, 8, 'FLETE', { sz: 12 });
        w(21, 3, '=IFERROR(C' + R(20) + '/C' + R(12) + ',0)', { fmt: USD, al: 'center' }); w(21, 7, '=' + PR.f('seguro'), { b: false, fmt: USD, al: 'center' }); w(21, 8, 'SEGURO', { sz: 12 });
        w(22, 2, 'Fac. Primeras', { fill: YE }); w(22, 3, '=C' + R(12) + '*C' + R(24), { fmt: USD, al: 'center' }); w(22, 7, '=' + PR.f('gastos'), { color: 'FF0000', fmt: '0.##', al: 'center' }); w(22, 8, 'GASTOS', { sz: 12, color: 'FF0000' });
        w(23, 2, 'Importe en Dolares'); w(23, 3, '=IFERROR(C' + R(22) + '/G' + R(17) + ',0)', { fmt: USD, al: 'center' }); w(23, 7, '=G' + R(18) + '-G' + R(19) + '-G' + R(20) + '-G' + R(21) + '-G' + R(22), { fmt: USD, al: 'center' });
        w(24, 2, 'PRECIO POR CAJA', { fill: '00B0F0' }); w(24, 3, '=' + PR.f('factPri'), { fill: YE, fmt: USD, al: 'center', b: false });
        w(25, 2, 'Fac. Segundas', { fill: '00B050' }); w(25, 3, '=C' + R(13) + '*C' + R(26), { fmt: USD, al: 'center' }); w(25, 7, '=E' + R(20), { fill: '00FFFF', fmt: USD, al: 'center' }); w(25, 8, 'PAGO TOTAL', { fill: '00FFFF', sz: 12, al: 'center' });
        w(26, 2, 'PRECIO POR CAJA'); w(26, 3, '=' + PR.f('fact15'), { fill: LB, fmt: USD, al: 'center', b: false });
        w(27, 2, '', { fill: LG }); w(27, 3, '=C' + R(22) + '+C' + R(25), { fill: LG, fmt: USD, al: 'center' });
        w(27, 5, '=E' + R(18) + '-E' + R(19) + '-E' + R(20) + '-E' + R(21) + '-E' + R(22) + '-E' + R(23) + '-E' + R(24) + '-E' + R(25) + '-E' + R(26), { fill: DG, color: 'FFFFFF', fmt: USD, al: 'center' });
        w(28, 2, '', { fill: LG }); w(28, 3, '', { fill: LG }); w(28, 4, 'VTA. NACIONAL', { sz: 11, al: 'right' }); w(28, 5, t.nac, { fill: CR, fmt: USD, al: 'center' });
        w(28, 6, '2DAS.', { al: 'center' }); w(28, 7, 'CAJAS', { al: 'center' }); w(28, 8, 'KILOS', { al: 'center' });
        w(29, 4, 'RESULTADO', { sz: 11, al: 'right' }); w(29, 5, '=E' + R(27) + '+E' + R(28), { fmt: USD, al: 'center' });
        w(30, 2, 'PAGOS A PRODUCTORES', { sz: 11 }); w(30, 4, '=E' + R(19) + '+E' + R(20) + '+E' + R(21) + '+E' + R(22) + '+E' + R(23) + '+E' + R(24) + '+E' + R(25), { fill: DG, color: 'FFFFFF', fmt: USD, al: 'center' });
        [['PRIMERAS', RKR], ['2DAS.', RKR + 1], [" 275'", RKR + 2], ['Desecho', RKR + 3]].forEach(function (x, i) {
          var r = 29 + i;
          w(r, 6, x[0], { fill: LB, al: 'center' });
          w(r, 7, '=' + sumB(function (bl) { return a(bl.c0 + 2, x[1]); }), { fill: LB, al: 'center', fmt: '0.00' });
          w(r, 8, '=' + sumB(function (bl) { return a(bl.c0 + 4, x[1]); }), { fill: LB, al: 'center', fmt: '0.00' });
        });
        w(31, 2, 'TOTAL DE CAJAS MAQUILADAS', { fill: LB, al: 'center', wrap: true }); w(31, 3, 'CAJAS VIAJE ANTERIOR', { fill: LB, al: 'center', wrap: true, sz: 12 }); w(31, 4, 'SOBRARON', { fill: LB, al: 'center', sz: 12 }); w(31, 5, 'TOTAL', { fill: '00B0F0', al: 'center' });
        /* v8.39 · puente: maquiladas (primera y 1½ de este viaje, sin segundas ni desecho) + llegaron de otro viaje − sobraron = manifestadas */
        var PU = t.puente || { maq: 0, rec: 0, sob: 0, man: 0 };
        w(32, 2, PU.maq, { fill: LB, al: 'center', fmt: '0.#' }); w(32, 3, PU.rec, { fill: '7030A0', color: 'FFFFFF', al: 'center', fmt: '0.#' });
        w(32, 4, PU.sob, { fill: '124F19', color: 'FFFFFF', al: 'center', fmt: '0.#' }); w(32, 5, '=B' + R(32) + '+C' + R(32) + '-D' + R(32), { fill: '00B0F0', al: 'center', fmt: '0.#' });
        w(33, 2, (PU.xde ? 'Manifiesto de ' + (t.empName || '') + ' que salió en el camión de ' + PU.xde + ': sus productores se pagan en ' + PU.xde + ' y el flete lo carga ese viaje · ' : '') + (t.xvN ? 'Incluye en «sobraron» ' + t.xvN + ' cajas del manifiesto de la otra empresa · ' : '') + (Math.abs(PU.dif || 0) > 0.01 ? '⚠ ' + (PU.dif > 0 ? PU.dif + ' cajas quedaron sin salir (en piso o almacén)' : (-PU.dif) + ' cajas manifestadas de más') + ' · ' : '✓ Cuadra con TOTAL CAJAS · ') + 'Primera y 1½ en cajas estándar; sobraron = enviadas a otro viaje o reempacadas.', { sz: 9, b: Math.abs(PU.dif || 0) > 0.01, color: Math.abs(PU.dif || 0) > 0.01 ? 'C00000' : '555555', wrap: false, al: 'left' });
        w(33, 6, '=IFERROR(H' + R(33) + '/' + PR.f('kgS') + ',0)', { fill: LB, al: 'center', fmt: D2 });
        w(33, 7, '=IFERROR(C' + R(12) + '/(' + sumB(function (bl) { return a(bl.c0 + 1, 7); }) + '/' + PR.f('kgC') + '),0)', { fill: LB, al: 'center', fmt: PC, color: 'FF0000' });
        w(33, 8, '=SUM(H' + R(29) + ':H' + R(32) + ')', { fill: LB, al: 'center', fmt: '0.00' });
        edgeS(R(14), 2, R(26), 3, 'medium'); edgeS(R(27), 2, R(28), 3, 'thin'); edgeS(R(31), 2, R(32), 5, 'medium'); edgeS(R(28), 6, R(33), 8, 'thin');

        /* ---- un bloque por productor ---- */
        blocks.forEach(function (bl) {
          var p = bl.p, c0 = bl.c0, mode = bl.mode, k = bl.p.cls;
          var X = function (dc, r) { return a(c0 + dc, r); };
          var xw = function (r, dc, v, op) { return w(r, c0 + dc, v, op); };
          var c15 = c15r(bl);
          var pc = function (f) { return '=' + PR.c(k, f); };
          xw(4, 0, '=$B$' + R(4), { sz: 12, fmt: 'd-mmm-yy', al: 'left' });
          xw(5, 0, '=$B$' + R(5), { sz: 16 }); xw(5, 2, p.cod ? 'Cód. ' + p.cod : '', { sz: 12, fill: YE, al: 'center' });
          for (var cc = 3; cc <= 5; cc++) ws.getCell(R(5), shC(c0 + cc)).fill = fill(COL[k]);
          ws.mergeCells(R(5), shC(c0 + 3), R(5), shC(c0 + 5));
          xw(5, 3, T(p.nombre), { sz: p.nombre.length > 20 ? 12 : 16, fill: COL[k], al: 'center', shrink: true });
          xw(5, 6, '% tamaños', { sz: 11, al: 'center' }); xw(4, 6, CLN[k].charAt(0) + CLN[k].slice(1).toLowerCase(), { sz: 11, al: 'center', color: '555555' });
          CK.forEach(function (ck, i) {
            var r = 6 + i;
            xw(r, 2, '=$B' + R(r), { al: 'right' });
            xw(r, 3, p.cal[ck] || 0, { al: 'center' });
            xw(r, 4, '=$F' + R(r), { fill: CR, fmt: USD, al: 'center' });
            xw(r, 5, '=' + X(3, r) + '*' + X(4, r), { b: false, fmt: USD, al: 'center' });
            xw(r, 6, '=IFERROR(' + X(3, r) + '/' + a(c0 + 3, 12).replace(/(\D+)(\d+)/, '$$$1$$$2') + ',0)', { b: false, fmt: PC, color: 'FF0000', al: 'center' });
          });
          xw(12, 3, '=SUM(' + X(3, 6) + ':' + X(3, 11) + ')', { color: 'FF0000', al: 'center' }); xw(12, 5, '=SUM(' + X(5, 6) + ':' + X(5, 11) + ')', { fmt: USD, al: 'center' });
          edgeS(R(5), c0 + 2, R(12), c0 + 6, 'thin');
          /* barra lateral */
          /* v8.28: cajas grandes y chicas que entregó el productor, como dato informativo */
          var rcb = (p.k24 || p.k18) ? (p.k24 || 0) + ' G Y ' + (p.k18 || 0) + ' CH' : '—';
          if (mode === 'box') {
            xw(6, 0, 'RECIBIDAS'); xw(6, 1, rcb, { al: 'center' });
            xw(7, 0, 'KILO'); xw(7, 1, p.kilos || 0, { color: 'FF0000', al: 'center', fmt: '#,##0.##' });
            xw(8, 0, 'CAJAS'); xw(8, 1, '=' + PR.f('kgC'), { b: false, al: 'center', fmt: '0.0#' });
            xw(9, 1, '=IFERROR(' + X(1, 7) + '/' + X(1, 8) + ',0)', { fill: 'CCFFFF', al: 'center', fmt: D2 });
          } else {
            xw(6, 0, 'RECIBIDAS'); xw(6, 1, rcb, { al: 'center' });
            xw(7, 0, 'KILOS'); xw(7, 1, p.kilos || 0, { al: 'center', fmt: '#,##0.##' });
            var gE = grRng(t, p, 'E'), gA = grRng(t, p, 'F');
            xw(8, 0, 'Bascula (MXN/kg)', { sz: 12, b: false }); xw(8, 1, gE ? '=IFERROR(SUM(' + gE + ')/' + X(1, 7) + ',' + PR.c(k, 'pkg') + ')' : '=' + PR.c(k, 'pkg'), { color: 'FF0000', fmt: USD, al: 'center' });
            xw(9, 1, '=' + X(1, 7) + '*' + X(1, 8), { fmt: USD, al: 'center' });
            xw(10, 0, 'ANTICIPO', { b: false }); xw(10, 1, gA ? '=SUM(' + gA + ')' : 0, { fill: CR, fmt: USD, al: 'center' });
            xw(11, 1, '=' + X(1, 9) + '-' + X(1, 10), { fmt: USD, al: 'center' });
          }
          xw(13, 0, 'CAMPO'); xw(13, 1, mode === 'box' ? '=' + X(1, 9) : '=IFERROR(' + X(1, 7) + '/' + PR.f('kgC') + ',0)', { al: 'center', fmt: D2 });
          xw(13, 2, mode === 'box' ? 'Total de Cajas' : 'Total de Cajas', { sz: 11, color: '3A7D22' }); xw(13, 3, 'Rendimiento', { al: 'center' }); xw(13, 5, '=' + X(3, 12), { al: 'center' });
          xw(14, 0, 'CAJAS', { color: '00B0F0' }); xw(14, 1, '=' + X(3, 12), { color: '00B0F0', al: 'center', fmt: D2 });
          xw(14, 2, mode === 'box' ? '=' + X(2, 23) : '=(' + X(3, 12) + '*' + PR.f('kgE') + ')/' + PR.f('kgS'), { color: '00B0F0', al: 'center', fmt: mode === 'box' ? '0.#' : D2 });
          xw(14, 3, '=IFERROR(' + X(1, 14) + '/' + X(1, 13) + ',0)', { color: 'FF0000', al: 'center', fmt: PC });
          xw(14, 5, '=IFERROR(' + X(5, 12) + '/' + X(5, 13) + ',0)', { b: false, fmt: USD, al: 'center' });
          xw(15, 1, '=' + (mode === 'box' ? X(4, 16) : X(4, 15)), { fmt: USD, al: 'center' }); xw(15, 2, '=$' + colL(c0 + 3) + '$' + R(5), { sz: 11 }); xw(15, 3, 'SUMA DE PRIMERAS', { color: 'FF0000', sz: 12, al: 'center' });
          xw(15, 5, '=' + PR.f('aduCaja'), { b: false, fmt: USD, al: 'center' }); xw(15, 6, 'ADUANA', { sz: 11 });
          xw(16, 0, 'POR CAJA', { color: '1F5C7A', sz: 12 }); xw(16, 1, '=' + X(1, 14) + '*' + X(1, 15), { fmt: USD, al: 'center' }); xw(16, 3, 'LIMON CORRIDO', { al: 'center', sz: 12 });
          xw(16, 5, '=' + X(5, 14) + '-' + X(5, 15), { fmt: USD, al: 'center' });
          xw(17, 5, '=$G$' + R(17), { fill: '92D050', fmt: USD, al: 'center' }); xw(17, 6, 'T.CAMBIO', { fill: '92D050', sz: 11 });
          xw(18, 0, '=IFERROR(' + X(1, 18) + '/' + X(3, 12) + ',0)', { fill: YE, fmt: USD, al: 'center' }); xw(18, 1, '=SUM(' + X(1, 16) + ':' + X(1, 17) + ')', { fill: DG, color: 'FFFFFF', fmt: USD, al: 'center' });
          xw(18, 5, '=' + X(5, 16) + '*' + X(5, 17), { fmt: USD, al: 'center' });
          xw(19, 0, '2 das', { b: false, sz: 12 }); xw(19, 5, '=$G$' + R(19), { b: false, fmt: USD, al: 'center' }); xw(19, 6, 'PLASTICO/RETORNO', { sz: 11 });
          xw(20, 0, 'CAJAS', { color: '00B0F0' }); xw(20, 1, p.seg || 0, { color: '00B0F0', al: 'center' }); xw(20, 5, '=$G$' + R(20), { b: false, fmt: USD, al: 'center' }); xw(20, 6, 'FLETE', { sz: 11 });
          xw(21, 0, 'JUGUERA', { sz: 12 }); xw(21, 1, p.jug || 0, { color: '00B0F0', al: 'center' }); xw(21, 5, '=$G$' + R(21), { b: false, fmt: USD, al: 'center' }); xw(21, 6, 'SEGURO CARGA', { sz: 11 });
          xw(22, 5, '=$G$' + R(22), { color: 'FF0000', fmt: '0.##', al: 'center' }); xw(22, 6, 'GASTOS', { color: 'FF0000', sz: 11 });
          xw(23, 0, "275'", { b: false, sz: 12 }); xw(23, 1, p.tor || 0, { color: '00B0F0', al: 'center' });
          xw(23, 5, '=' + X(5, 18) + '-' + X(5, 19) + '-' + X(5, 20) + '-' + X(5, 22) + '-' + X(5, 21), { fmt: USD, al: 'center' });
          xw(24, 5, '=' + X(1, 15), { b: false, fmt: USD, al: 'center' }); xw(24, 6, 'V.PAGADO', { sz: 11 });
          xw(26, 0, 'DESECHO', { b: false, sz: 12 }); xw(26, 1, p.des || 0, { color: '00B0F0', al: 'center' });
          /* tabla de pago: v8.10 una fila por calibre (250, 230, 200, 175, 150 y 110), luego segundas, Torreón y desecho */
          var P0 = 17, PS = P0 + 6, SG = PS + 1, TO = PS + 2, DE = PS + 3, TT = PS + 4, KR = TT + 1, KT = KR + 4;
          var pf = ['p250', 'p230', 'p200', 'p175', 'p150', 'p110'];
          pf.forEach(function (f, i) {
            var r = P0 + i;
            xw(r, 2, mode === 'box' ? '=ROUND(' + X(3, 6 + i) + '*' + PR.f('kgE') + '/' + PR.f('kgS') + '*2,0)/2' : '=' + X(3, 6 + i), { al: 'center', fmt: '0.#' }); /* v8.39: cajas de 18 kg, al 0.5 más cercano */
            xw(r, 3, mode === 'box' ? pc(f) : '=' + X(4, 15), { fmt: USD, al: 'center' });
            xw(r, 4, '=' + X(2, r) + '*' + X(3, r), { fmt: USD, al: 'center' });
          });
          xw(PS, 2, '=SUM(' + X(2, P0) + ':' + X(2, PS - 1) + ')', { al: 'center' }); xw(PS, 4, '=SUM(' + X(4, P0) + ':' + X(4, PS - 1) + ')', { fmt: USD, al: 'center' });
          if (mode === 'box') {
            xw(16, 4, '=IFERROR(' + X(4, PS) + '/' + X(3, 12) + ',0)', { fmt: USD, al: 'center' });
          } else {
            xw(14, 4, '=' + X(1, 11) + '-' + X(4, SG) + '-' + X(4, TO) + '-' + X(4, DE), { fmt: USD, al: 'center' });
            xw(15, 4, '=IFERROR(' + X(4, 14) + '/' + X(2, PS) + ',0)', { fmt: USD, al: 'center' });
          }
          xw(17, 1, '=' + X(4, SG) + '+' + X(4, TO) + '+' + X(4, DE), { fmt: USD, al: 'center' });
          xw(SG, 2, '=' + X(1, 20), { al: 'center' }); xw(SG, 3, pc('seg'), { fmt: USD, al: 'center' }); xw(SG, 4, '=' + X(2, SG) + '*' + X(3, SG), { fmt: USD, al: 'center' });
          xw(TO, 2, '=' + X(1, 23), { al: 'center' }); xw(TO, 3, pc('tor'), { fmt: USD, al: 'center' }); xw(TO, 4, '=' + X(2, TO) + '*' + X(3, TO), { fmt: USD, al: 'center' });
          xw(DE, 2, '=' + X(1, 26), { al: 'center' }); xw(DE, 3, pc('des'), { fmt: USD, al: 'center' }); xw(DE, 4, '=' + X(2, DE) + '*' + X(3, DE), { fmt: USD, al: 'center' });
          xw(TT, 2, '=' + X(2, PS) + '+' + X(2, SG) + '+' + X(2, TO) + '+' + X(2, DE), { al: 'center', fmt: '0' });
          xw(TT, 3, '=IFERROR(' + X(4, TT) + '/' + X(1, 13) + ',0)', { fill: YE, fmt: USD, al: 'center', color: 'FF0000' });
          xw(TT, 4, '=' + X(4, PS) + '+' + X(4, SG) + '+' + X(4, TO) + '+' + X(4, DE), { fill: YE, fmt: USD, al: 'center' });
          xw(25, 5, '=' + X(5, 23) + '-' + X(5, 24), { fill: YE, fmt: USD, al: 'center' });
          /* kilos por clasificación */
          var kr = [[KR, '=' + X(3, 12), '=' + PR.f('kgE')], [KR + 1, '=' + X(1, 20), '=' + PR.f('kgS')], [KR + 2, '=' + X(1, 23), '=' + PR.f('kgS')], [KR + 3, '=' + X(1, 26), '=' + PR.f('kgS')]];
          kr.forEach(function (x) { xw(x[0], 2, x[1], { color: 'FF0000', al: 'center' }); xw(x[0], 3, x[2], { b: false, al: 'center', fmt: '0.0#' }); xw(x[0], 4, '=' + X(2, x[0]) + '*' + X(3, x[0]), { fmt: '0.0', al: 'center' }); });
          xw(KT, 0, 'TOTAL DE CAJAS', { fill: YE, sz: 12 }); xw(KT, 1, '', { fill: YE });
          xw(KT, 2, '=SUM(' + X(2, KR) + ':' + X(2, KR + 3) + ')', { al: 'center' }); xw(KT, 4, '=SUM(' + X(4, KR) + ':' + X(4, KR + 3) + ')', { fmt: '0.0', al: 'center' });
          xw(KT + 1, 0, 'CAJAS MAQUILADAS', { fill: LG, color: 'FF0000', sz: 12 }); xw(KT + 1, 1, '=' + X(1, 14) + '+' + X(1, 20) + '+' + X(1, 23) + '+' + X(1, 26) + '+' + X(1, 21), { fill: LG, fmt: D2, al: 'center' });
          xw(KT + 1, 2, 'PESO CORRIDA', { color: 'FF0000', sz: 11, al: 'center' }); xw(KT + 1, 3, 'PESO BASCULA', { color: 'FF0000', sz: 11, al: 'center' }); xw(KT + 1, 4, 'DIFERENCIA', { fill: YE, sz: 11, al: 'center' });
          xw(KT + 2, 2, '=' + X(4, KT) + '+' + X(1, 21) + '*' + PR.f('kgS'), { al: 'center', fmt: '0.0' }); xw(KT + 2, 3, '=' + X(1, 7), { al: 'center', fmt: '0.0' }); xw(KT + 2, 4, '=' + X(2, KT + 2) + '-' + X(3, KT + 2), { al: 'center', fmt: '0.0' });
          xw(KT + 3, 2, '=IFERROR(' + X(2, KT + 2) + '/' + PR.f('kgS') + ',0)', { b: false, al: 'center', fmt: D2 }); xw(KT + 3, 3, '=IFERROR(' + X(3, KT + 2) + '/' + PR.f('kgS') + ',0)', { b: false, al: 'center', fmt: D2 }); xw(KT + 3, 4, '=' + X(2, KT + 3) + '-' + X(3, KT + 3), { b: false, al: 'center', fmt: D2 });
          edgeS(R(4), c0, R(KT + 3), c0 + 6, 'medium');
          blocksAll.push({ key: p.key, cod: p.cod, nombre: p.nombre, cls: k, sheet: G.name, o: o, c0: shC(c0), c15: c15, mode: mode, kilos: p.kilos || 0 });
        });
        tripRefs.push({ G: G, t: t, sheet: G.name, o: o, blocks: blocks, S: S });
      });
    });

    /* ---- RESUMEN SEMANAL ---- */
    var ref = function (sn, c, r) { return q(sn) + '!' + c + r; };
    [9, 24, 12, 16, 16, 22, 13, 11, 13, 12, 13, 14, 16, 18, 16, 16].forEach(function (wd, i) { wsRes.getColumn(1 + i).width = wd; });
    W(wsRes, 1, 1, 'CIERRE SEMANAL · ' + C.per.t, { sz: 18 });
    W(wsRes, 2, 1, 'Generado ' + C.gen + '. Los números son fórmulas que leen cada viaje; cambia una casilla amarilla en su hoja y aquí se actualiza.', { sz: 10, b: false, i: true, color: '555555' });
    var HD = ['VIAJE', 'SEMANA', 'EMPRESA', 'PEDIDO', 'EMBARQUE', 'CLIENTE', 'FECHA', 'PRODUCT.', 'CAJAS 1RA', 'CAJAS 2DAS.', 'TOTAL CAJAS', 'VENTA USD', 'VENTA MXN', 'PAGO PRODUCT.', 'VTA. NACIONAL', 'RESULTADO'];
    HD.forEach(function (t, i) { W(wsRes, 4, 1 + i, t, { sz: 11, fill: '1B2C46', color: 'FFFFFF', al: 'center', wrap: true }); });
    wsRes.getRow(4).height = 30;
    var rr = 5, first = rr;
    tripRefs.forEach(function (x) {
      var t = x.t, s = x.sheet, o = x.o;
      var cr = function (c, r) { return '=' + ref(s, colL(colN(c) + (colN(c) >= 3 ? x.S || 0 : 0)), r + o); };
      var vals = [t.n, x.G.label, x.G.empName, T(t.cod), T(t.emb), T(t.cliente), t.fecha ? new Date(t.fecha + 'T12:00:00Z') : '', t.prods.length, cr('C', 12), cr('C', 13), cr('C', 14), cr('H', 10), cr('E', 13), cr('E', 20), cr('E', 28), cr('E', 29)];
      var fm = [null, null, null, null, null, null, 'dd/mm/yyyy', INT, INT, INT, INT, USD, USD, USD, USD, USD];
      vals.forEach(function (v, i) { W(wsRes, rr, 1 + i, v, { sz: 11, b: i === 15, al: i >= 7 ? 'right' : (i === 0 ? 'center' : 'left'), fmt: fm[i] || undefined }); wsRes.getCell(rr, 1 + i).border = { bottom: { style: 'hair', color: { argb: 'FFB7B7B7' } } }; });
      rr++;
    });
    W(wsRes, rr, 1, 'TOTAL', { sz: 11, fill: 'EDEDED' }); for (var cz = 2; cz <= 16; cz++) W(wsRes, rr, cz, '', { fill: 'EDEDED' });
    [8, 9, 10, 11, 12, 13, 14, 15, 16].forEach(function (ci) { var L = colL(ci); W(wsRes, rr, ci, '=SUM(' + L + first + ':' + L + (rr - 1) + ')', { sz: 11, fill: 'EDEDED', al: 'right', fmt: ci <= 11 ? INT : USD }); });
    var totRow = rr; rr += 3;
    W(wsRes, rr, 1, 'PAGO POR CLASIFICACIÓN DE PRODUCTORES', { sz: 12 }); rr++;
    ['CLASIFICACIÓN', '', 'PRODUCTORES', 'CAJAS 1RA', 'CAJAS 2DAS.', 'KILOS', 'PAGO TOTAL'].forEach(function (t, i) { W(wsRes, rr, 1 + i, t, { sz: 11, fill: '1B2C46', color: 'FFFFFF', al: 'center' }); });
    rr++;
    ['chico', 'grande', 'interno'].forEach(function (k) {
      var bs = blocksAll.filter(function (b) { return b.cls === k; });
      var sm = function (fn) { return bs.length ? '=SUM(' + bs.map(fn).join(',') + ')' : 0; };
      var nk = {}; bs.forEach(function (b) { nk[b.key] = 1; });
      W(wsRes, rr, 1, CLN[k], { sz: 11, fill: COL[k] }); W(wsRes, rr, 2, '', { fill: COL[k] });
      W(wsRes, rr, 3, Object.keys(nk).length, { sz: 11, al: 'right', fmt: INT });
      W(wsRes, rr, 4, sm(function (b) { return ref(b.sheet, colL(b.c0 + 3), 12 + b.o); }), { sz: 11, al: 'right', fmt: INT });
      W(wsRes, rr, 5, sm(function (b) { return ref(b.sheet, colL(b.c0 + 1), 20 + b.o); }), { sz: 11, al: 'right', fmt: INT });
      W(wsRes, rr, 6, sm(function (b) { return ref(b.sheet, colL(b.c0 + 1), 7 + b.o); }), { sz: 11, al: 'right', fmt: '#,##0' });
      W(wsRes, rr, 7, sm(function (b) { return ref(b.sheet, colL(b.c0 + 4), 27 + b.o); }), { sz: 11, al: 'right', fmt: USD });
      rr++;
    });
    rr += 1;
    var notas = ['Hojas: PRECIOS (lo que se paga a cada clasificación y los datos fijos del viaje), SEMANA (un bloque por viaje, igual a FORMATO VIAJES) y PAGOS PRODUCTORES (lo que corresponde a cada productor en el periodo).',
      'Colores de los productores: verde = Chicos, cian = Grandes, rosa = Internos. Los Grandes se pagan por kilo de báscula menos el anticipo; Chicos e Internos, por caja y calibre.',
      'Rendimiento = cajas de primera ÷ cajas campo (kilos ÷ kilos por caja de campo de PRECIOS), igual que el Excel de resultados. Las segundas se pagan en «2 das» de cada productor y la Juguera va en su barra lateral (JUGUERA). La fila SEGUNDAS del viaje son las cajas de segunda empacadas para exportación.',
      'La venta nacional (segundas, Torreón y Juguera) sale de Finanzas y se suma al resultado en la casilla VTA. NACIONAL de cada viaje.'];
    if (C.sinCls && C.sinCls.length) notas.push('Productores sin clasificación guardada (se tomaron como Chicos; los de Jesús Lagunes como Internos): ' + C.sinCls.join(', ') + '.');
    if (C.sinPrecio && C.sinPrecio.length) notas.push('Entregas de productores grandes sin precio por kilo (su pago sale en 0 hasta capturarlo en la app): ' + C.sinPrecio.join(', ') + '.');
    if (C.mm > 0) notas.push('No entraron ' + C.mm + ' cajas de calibres en milímetros (CTN 10 lb): FORMATO VIAJES no tiene fila para ellas.');
    notas.forEach(function (t, i) { wsRes.mergeCells(rr + i, 1, rr + i, 16); W(wsRes, rr + i, 1, t, { sz: 10, b: false, i: true, color: '555555', wrap: true, al: 'left' }); wsRes.getRow(rr + i).height = 28; });

    /* ---- PAGOS PRODUCTORES ---- */
    [10, 38, 16, 9, 12, 11, 11, 11, 11, 11, 16].forEach(function (wd, i) { wsPag.getColumn(1 + i).width = wd; });
    W(wsPag, 1, 1, 'PAGOS A PRODUCTORES · ' + C.per.t, { sz: 16 });
    W(wsPag, 2, 1, 'Suma de lo que cada productor tiene en los viajes del periodo (fórmulas sobre las hojas SEMANA).', { sz: 10, b: false, i: true, color: '555555' });
    ['CÓDIGO', 'PRODUCTOR', 'CLASIFICACIÓN', 'VIAJES', 'KILOS', 'PRIMERAS', 'SEGUNDAS', "TORREÓN", 'DESECHO', 'JUGUERA', 'PAGO TOTAL'].forEach(function (t, i) { W(wsPag, 4, 1 + i, t, { sz: 11, fill: '1B2C46', color: 'FFFFFF', al: 'center', wrap: true }); });
    var byKey = {}, order = [];
    blocksAll.forEach(function (b) { if (!byKey[b.key]) { byKey[b.key] = { b: b, list: [] }; order.push(byKey[b.key]); } byKey[b.key].list.push(b); });
    var co = { chico: 0, grande: 1, interno: 2 };
    order.sort(function (x, y) { return co[x.b.cls] - co[y.b.cls] || x.b.nombre.localeCompare(y.b.nombre, 'es'); });
    var pr0 = 5, pr = pr0;
    order.forEach(function (g) {
      var L = g.list, b0 = g.b, sm = function (fn) { return '=SUM(' + L.map(fn).join(',') + ')'; };
      W(wsPag, pr, 1, T(b0.cod ? b0.cod : ''), { sz: 11, al: 'center' }); W(wsPag, pr, 2, T(b0.nombre), { sz: 11, al: 'left' }); W(wsPag, pr, 3, CLN[b0.cls].charAt(0) + CLN[b0.cls].slice(1).toLowerCase(), { sz: 11, fill: COL[b0.cls], al: 'center' });
      W(wsPag, pr, 4, L.length, { sz: 11, al: 'center' });
      var cols = [[5, function (b) { return ref(b.sheet, colL(b.c0 + 1), 7 + b.o); }, '#,##0'], [6, function (b) { return ref(b.sheet, colL(b.c0 + 3), 12 + b.o); }, INT],
        [7, function (b) { return ref(b.sheet, colL(b.c0 + 1), 20 + b.o); }, INT], [8, function (b) { return ref(b.sheet, colL(b.c0 + 1), 23 + b.o); }, INT], [9, function (b) { return ref(b.sheet, colL(b.c0 + 1), 26 + b.o); }, INT], [10, function (b) { return ref(b.sheet, colL(b.c0 + 1), 21 + b.o); }, INT], [11, function (b) { return ref(b.sheet, colL(b.c0 + 4), 27 + b.o); }, USD]];
      cols.forEach(function (x) { W(wsPag, pr, x[0], sm(x[1]), { sz: 11, b: x[0] === 11, al: 'right', fmt: x[2] }); });
      for (var cx = 1; cx <= 11; cx++) wsPag.getCell(pr, cx).border = { bottom: { style: 'hair', color: { argb: 'FFB7B7B7' } } };
      pr++;
    });
    W(wsPag, pr, 1, 'TOTAL', { sz: 11, fill: 'EDEDED' }); for (var cy = 2; cy <= 11; cy++) W(wsPag, pr, cy, '', { fill: 'EDEDED' });
    for (var cs = 5; cs <= 11; cs++) { var Ls = colL(cs); W(wsPag, pr, cs, pr > pr0 ? '=SUM(' + Ls + pr0 + ':' + Ls + (pr - 1) + ')' : 0, { sz: 11, fill: 'EDEDED', al: 'right', fmt: cs === 11 ? USD : INT }); }
    wsPag.autoFilter = { from: { row: 4, column: 1 }, to: { row: Math.max(pr - 1, 5), column: 11 } };
    wb.calcProperties = { fullCalcOnLoad: true };
    calcWb(wb);
    return wb.worksheets;
  };

  /* ---------- v8.12 · hoja semanal de Entrada de fruta (calca de «Entrada de limón», hoja de la semana 42) ----------
     Un bloque por día (lunes a domingo) con Productor, Kilos, Cajas chicas, Cajas grandes y Total (= chicas + grandes × 1.5).
     Debajo: lo que ya está en piso, los estimados de productores locales, Oaxaca y pendientes (para capturar a mano),
     las entradas reales por día, lo que falta por entrar y las cajas de cada viaje de la semana. */
  XL.addRecSem = function (wb, W) {
    var FT = 'Aptos Display', GL = 'C1F0C8', ST = 'F2F2F2', YE = 'FFFF00', RED = 'FF0000';
    var ws = wb.addWorksheet(safeName(wb, 'SEMANA ' + W.num), { views: [{ state: 'frozen', ySplit: 5, zoomScale: 80, showGridLines: true }], pageSetup: { orientation: 'landscape', paperSize: 1, fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    var widths = { A: 6.6, B: 17.6, C: 37.4, D: 23, E: 20.6, F: 20, G: 11, H: 18.3, I: 6.3, J: 41.6, K: 18.3, L: 20.6, M: 20.1, N: 11.6, O: 18.4, P: 7.3, Q: 44.1, R: 19.6, S: 18, T: 20.1, U: 11, V: 11.4, W: 44.3, X: 16, Y: 18.1, Z: 20.3, AA: 11.1, AB: 11.4, AC: 38.9, AD: 26.4, AE: 18.1, AF: 20.3, AG: 11.1, AH: 11.4, AI: 30.6, AJ: 26.4, AK: 18.3, AL: 20.9, AM: 13.3, AN: 30.6, AO: 20, AP: 18, AQ: 20.3, AR: 12 };
    Object.keys(widths).forEach(function (k) { ws.getColumn(k).width = widths[k]; });
    function Wc(r, c, v, o) {
      o = o || {}; var cell = ws.getCell(r, c);
      if (v !== undefined && v !== null && v !== '') cell.value = v;
      cell.font = { name: FT, size: o.sz || 12, bold: o.b !== false, color: o.color ? { argb: 'FF' + o.color } : undefined };
      if (o.fill) cell.fill = fill(o.fill);
      if (o.fmt) cell.numFmt = o.fmt;
      cell.alignment = { horizontal: o.al, vertical: 'middle', wrapText: !!o.wrap };
      if (o.box) { var s = { style: 'thin', color: { argb: 'FF000000' } }; cell.border = { top: s, left: s, bottom: s, right: s }; }
      if (o.dbl) cell.border = Object.assign({}, cell.border || {}, { top: { style: 'double', color: { argb: 'FF000000' } } });
      if (o.bot) cell.border = Object.assign({}, cell.border || {}, { bottom: { style: 'thin', color: { argb: 'FF000000' } } });
      return cell;
    }
    var F = function (f, res) { return { formula: f, result: res == null ? 0 : res }; };
    var C0 = [3, 10, 17, 23, 29, 35, 40];  /* C, J, Q, W, AC, AI, AN */
    var maxN = 0; W.days.forEach(function (d) { maxN = Math.max(maxN, d.rows.length); });
    var N = Math.max(36, maxN), R1 = 7, RN = R1 + N - 1, T = RN + 1, dT = T - 43;
    for (var r = 2; r <= 9; r++) ws.getRow(r).height = r === 6 ? 20.3 : 19.5;
    var dayTot = [];
    W.days.forEach(function (d, i) {
      var c = C0[i], L = function (k) { return colL(c + k); };
      ws.mergeCells(2, c, 3, c + 4); Wc(2, c, d.lab, { sz: 24, al: 'center' });
      ws.mergeCells(4, c, 5, c + 4); Wc(4, c, 'REPORTE DE ENTRADA DE FRUTA', { sz: 18, fill: GL, al: 'center', dbl: true });
      ['Productor', 'Kilos', 'Cajas chicas', 'Cajas Grandes', 'Total'].forEach(function (h, k) { Wc(6, c + k, h, { sz: 16, al: 'center', bot: true }); });
      var sk = 0, sc = 0, sg = 0, st = 0;
      for (var j = 0; j < N; j++) {
        var rr = R1 + j, x = d.rows[j], sf = j % 2 === 0 ? ST : null;
        Wc(rr, c, x ? x.n : null, { box: true, fill: sf });
        Wc(rr, c + 1, x && x.kilos != null ? x.kilos : null, { box: true, fill: sf, fmt: '#,##0.##', al: 'center' });
        Wc(rr, c + 2, x && x.ch != null ? x.ch : null, { b: false, box: true, fill: sf, al: 'center' });
        Wc(rr, c + 3, x && x.gr != null ? x.gr : null, { b: false, box: true, fill: sf, al: 'center' });
        var tv = x ? (+x.ch || 0) + (+x.gr || 0) * 1.5 : 0;
        Wc(rr, c + 4, F('+' + L(2) + rr + '+' + L(3) + rr + '*1.5', tv), { b: false, box: true, al: 'center', fmt: '#,##0.##' });
        if (x) { sk += +x.kilos || 0; sc += +x.ch || 0; sg += +x.gr || 0; st += tv; }
      }
      Wc(T, c + 1, F('SUM(' + L(1) + R1 + ':' + L(1) + RN + ')', sk), { sz: 14, b: false, fmt: '#,##0.##', al: 'center' });
      Wc(T, c + 2, F('SUM(' + L(2) + R1 + ':' + L(2) + RN + ')', sc), { sz: 14, b: false, al: 'center' });
      Wc(T, c + 3, F('SUM(' + L(3) + R1 + ':' + L(3) + RN + ')', sg), { sz: 14, b: false, al: 'center' });
      Wc(T, c + 4, F('SUM(' + L(4) + R1 + ':' + L(4) + RN + ')', st), { sz: 14, al: 'center', fmt: '#,##0.##' });
      Wc(T + 1, c + 4, null, { sz: 14, color: RED, fmt: '0%', al: 'center', dbl: true, fill: YE });
      Wc(T + 2, c + 4, F(L(4) + T + '*' + L(4) + (T + 1), 0), { sz: 14, al: 'center', fmt: '#,##0.##' });
      dayTot.push({ c: c, k: sk, ch: sc, gr: sg, t: st, lab: d.lab, int: d.rows.map(function (x, j) { return x.int ? R1 + j : 0; }).filter(Boolean) });
    });
    /* lo que ya está en piso */
    var sum = function (k) { return dayTot.reduce(function (s, x) { return s + x[k]; }, 0); };
    var ref = function (k) { return '+' + dayTot.map(function (x) { return colL(x.c + k) + T; }).join('+'); };
    Wc(46 + dT, 3, 'PENDIENTE POR TRAER ', { sz: 14, b: false }); Wc(46 + dT, 11, null, { sz: 14, b: false, fill: YE, al: 'center' });
    Wc(47 + dT, 3, 'YA EN PISO O RECORRIDO ', { sz: 14, b: false }); Wc(47 + dT, 11, F(ref(1), sum('k')), { sz: 14, b: false, fmt: '#,##0.##', al: 'center' }); Wc(47 + dT, 12, 'Kilos', { sz: 14, b: false });
    Wc(48 + dT, 3, 'CANCELO SU CORTE ', { sz: 14, b: false }); Wc(48 + dT, 11, F(ref(2), sum('ch')), { sz: 14, b: false, al: 'center' }); Wc(48 + dT, 12, 'Chicas', { sz: 14, b: false });
    Wc(49 + dT, 11, F(ref(3), sum('gr')), { sz: 14, b: false, al: 'center' }); Wc(49 + dT, 12, 'Grandes', { sz: 14, b: false });
    /* estimados (se capturan a mano en el Excel): productores locales, Oaxaca y pendientes */
    var E0 = 55 + dT, M = Math.max(36, W.prods.length), EN = E0 + M - 1, ET = EN + 1, dE = dT + (M - 36);
    var est = [[3, 'REPORTE DE ENTRADA DE FRUTA PRODUCTORES LOCALES ', W.prods], [10, 'ESTIMADO DE ENTRADA DE FRUTA  OAXACA', []], [17, 'PENDIENTE DE ENTREGAR', []]];
    var estTot = {};
    est.forEach(function (b) {
      var c = b[0], L = function (k) { return colL(c + k); };
      ws.mergeCells(52 + dT, c, 53 + dT, c + 5); Wc(52 + dT, c, b[1], { sz: 18, fill: GL, al: 'center' });
      ['Productor', 'Cajas chicas', 'Cajas Grandes', 'Total', '% ', 'Rendimiento '].forEach(function (h, k) { Wc(54 + dT, c + k, h, { sz: 16, al: 'center', bot: true, fmt: k >= 4 ? '0%' : undefined }); });
      for (var j = 0; j < M; j++) {
        var rr = E0 + j, x = b[2][j];
        Wc(rr, c, x ? x.n : null, { box: true });
        Wc(rr, c + 1, null, { b: false, box: true, al: 'center', fill: YE });
        Wc(rr, c + 2, null, { b: false, box: true, al: 'center', fill: YE });
        Wc(rr, c + 3, F('+' + L(1) + rr + '+' + L(2) + rr + '*1.5', 0), { b: false, box: true, al: 'center', fmt: '#,##0.##' });
        Wc(rr, c + 4, 0.7, { box: true, color: RED, fmt: '0%', al: 'center' });
        Wc(rr, c + 5, F('+' + L(3) + rr + '*' + L(4) + rr, 0), { b: false, box: true, fmt: '0.00', al: 'center' });
      }
      Wc(ET, c + 3, F('SUM(' + L(3) + E0 + ':' + L(3) + EN + ')', 0), { sz: 14, al: 'center', fmt: '#,##0.##' });
      Wc(ET, c + 5, F('SUM(' + L(5) + E0 + ':' + L(5) + EN + ')', 0), { sz: 14, al: 'center', fmt: '0.00' });
      estTot[c] = { t: L(3) + ET, r: L(5) + ET, first: L(3) + E0 };
    });
    var r93 = 93 + dE;
    Wc(r93, 3, F('+' + estTot[10].r, 0), { sz: 14, fmt: '0.00', al: 'center' }); Wc(r93, 4, 'OAXACA', { sz: 14, b: false });
    Wc(r93 + 1, 3, F('+' + estTot[3].r, 0), { sz: 14, fmt: '0.00', al: 'center' }); Wc(r93 + 1, 4, 'PRODUCTORES LOCALES', { sz: 14, b: false });
    Wc(r93 + 2, 3, F('SUM(C' + r93 + ':C' + (r93 + 1) + ')', 0), { sz: 14, fmt: '0.00', al: 'center' });
    /* entradas reales por día */
    var r99 = 99 + dE;
    ws.mergeCells(r99, 3, r99, 4); Wc(r99, 3, 'Entradas Reales ', { sz: 18, fill: GL, al: 'center', dbl: true });
    dayTot.forEach(function (x, i) { Wc(r99 + 1 + i, 3, F('+' + colL(x.c) + '2', x.lab), { sz: 14, b: false }); Wc(r99 + 1 + i, 4, F(colL(x.c + 4) + T, x.t), { sz: 14, b: false, al: 'center', fmt: '#,##0.##' }); });
    var rs = r99 + 8; Wc(rs, 4, F('SUM(D' + (r99 + 1) + ':D' + (r99 + 7) + ')', sum('t')), { sz: 14, al: 'center', fmt: '#,##0.##' });
    var r109 = rs + 3;
    Wc(r109, 3, 'FALTAN POR ENTRAR', { sz: 14, fill: YE }); Wc(r109, 4, F(estTot[3].r + '-D' + rs, -sum('t')), { sz: 14, color: RED, fill: YE, fmt: '0.00', al: 'center' });
    var intRefs = []; dayTot.forEach(function (x) { x.int.forEach(function (rr) { intRefs.push(colL(x.c + 4) + rr); }); });
    var nInt = W.prods.filter(function (p) { return p.int; }).length;
    var intEst = nInt ? 'SUM(F' + E0 + ':F' + (E0 + nInt - 1) + ')' : '0';
    Wc(r109 + 1, 3, 'JESUS LAGUNES ', { sz: 14, b: false }); Wc(r109 + 1, 4, F(intEst + (intRefs.length ? '-' + intRefs.join('-') : ''), 0), { sz: 14, b: false, fmt: '0.00', al: 'center' });
    Wc(r109 + 2, 3, 'PRODUCTORES', { sz: 14, b: false }); Wc(r109 + 2, 4, F('D' + r109 + '-D' + (r109 + 1), 0), { sz: 14, b: false, fmt: '0.00', al: 'center' });
    /* cajas por viaje de la semana */
    var r115 = r109 + 6, nV = Math.max(9, W.trips.length);
    Wc(r115, 3, 'VIAJE #', { sz: 14, fill: GL, dbl: true }); Wc(r115, 4, 'CAJAS POR VIAJE', { sz: 14, fill: GL, dbl: true });
    var tv = 0;
    for (var v = 0; v < nV; v++) { var t = W.trips[v]; Wc(r115 + 1 + v, 3, 'VIAJE ' + (v + 1) + (t ? ' · ' + t.cod : ''), { sz: 14, b: false, box: true }); Wc(r115 + 1 + v, 4, t ? Math.round(t.cajas * 100) / 100 : 0, { sz: 14, b: false, box: true, al: 'center', fmt: '#,##0.##' }); tv += t ? t.cajas : 0; }
    var rvt = r115 + 1 + nV; Wc(rvt, 4, F('SUM(D' + (r115 + 1) + ':D' + (rvt - 1) + ')', tv), { sz: 14, al: 'center', fmt: '#,##0.##' });
    Wc(rvt + 3, 3, F('C' + (r93 + 2) + '-D' + rvt, -tv), { sz: 14, color: RED, fmt: '0.00' }); Wc(rvt + 3, 4, 'CAJAS FALTANTES', { sz: 14 });
    return ws;
  };
  /* versión para PDF: un día debajo de otro, en hoja vertical, solo lo capturado */
  XL.addRecSemPdf = function (wb, W, empName) {
    var ws = wb.addWorksheet(safeName(wb, 'PDF ' + W.num), { views: [{ showGridLines: false }], pageSetup: { orientation: 'portrait', paperSize: 1, fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    [3, 44, 16, 14, 14, 14].forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    var GL = 'C1F0C8', ST = 'F2F2F2', DK = '196B24';
    function Wc(r, c, v, o) { o = o || {}; var cell = ws.getCell(r, c); if (v !== undefined && v !== null && v !== '') cell.value = v; cell.font = { name: 'Aptos Display', size: o.sz || 11, bold: !!o.b, color: o.color ? { argb: 'FF' + o.color } : undefined }; if (o.fill) cell.fill = fill(o.fill); if (o.fmt) cell.numFmt = o.fmt; cell.alignment = { horizontal: o.al, vertical: 'middle', wrapText: !!o.wrap }; if (o.box) { var s = { style: 'thin', color: { argb: 'FFBFBFBF' } }; cell.border = { top: s, left: s, bottom: s, right: s }; } return cell; }
    var r = 1;
    ws.mergeCells(r, 2, r, 6); Wc(r, 2, 'ENTRADA DE FRUTA · SEMANA ' + W.num + ' · ' + W.anio, { sz: 18, b: true, color: DK }); ws.getRow(r).height = 28; r++;
    ws.mergeCells(r, 2, r, 6); Wc(r, 2, (empName ? empName + ' · ' : '') + W.span + ' · Total = cajas chicas + cajas grandes × 1.5', { sz: 10, color: '555555' }); r += 2;
    var tk = 0, tc = 0, tg = 0, tt = 0;
    W.days.forEach(function (d) {
      if (!d.rows.length) return;
      ws.mergeCells(r, 2, r, 6); Wc(r, 2, d.lab, { sz: 13, b: true, fill: GL }); ws.getRow(r).height = 22; r++;
      ['Productor', 'Kilos', 'Cajas chicas', 'Cajas grandes', 'Total'].forEach(function (h, k) { Wc(r, 2 + k, h, { b: true, al: k ? 'right' : 'left', box: true }); }); r++;
      var sk = 0, sc = 0, sg = 0, st = 0;
      d.rows.forEach(function (x, j) {
        var t = (+x.ch || 0) + (+x.gr || 0) * 1.5, f = j % 2 ? ST : null;
        Wc(r, 2, x.n, { box: true, fill: f }); Wc(r, 3, x.kilos, { box: true, fill: f, fmt: '#,##0.##', al: 'right' }); Wc(r, 4, x.ch, { box: true, fill: f, al: 'right' }); Wc(r, 5, x.gr, { box: true, fill: f, al: 'right' }); Wc(r, 6, t, { box: true, fill: f, fmt: '#,##0.##', al: 'right', b: true });
        sk += +x.kilos || 0; sc += +x.ch || 0; sg += +x.gr || 0; st += t; r++;
      });
      Wc(r, 2, 'Total del día', { b: true, box: true }); Wc(r, 3, sk, { b: true, box: true, fmt: '#,##0.##', al: 'right' }); Wc(r, 4, sc, { b: true, box: true, al: 'right' }); Wc(r, 5, sg, { b: true, box: true, al: 'right' }); Wc(r, 6, st, { b: true, box: true, fmt: '#,##0.##', al: 'right' });
      tk += sk; tc += sc; tg += sg; tt += st; r += 2;
    });
    if (!W.n) { Wc(r, 2, 'Sin entradas de fruta en esta semana.', { color: '555555' }); r += 2; }
    ws.mergeCells(r, 2, r, 6); Wc(r, 2, 'RESUMEN DE LA SEMANA', { sz: 13, b: true, fill: GL }); r++;
    ['Día', 'Kilos', 'Cajas chicas', 'Cajas grandes', 'Total'].forEach(function (h, k) { Wc(r, 2 + k, h, { b: true, al: k ? 'right' : 'left', box: true }); }); r++;
    W.days.forEach(function (d) { var k = 0, c = 0, g = 0, t = 0; d.rows.forEach(function (x) { k += +x.kilos || 0; c += +x.ch || 0; g += +x.gr || 0; t += (+x.ch || 0) + (+x.gr || 0) * 1.5; }); Wc(r, 2, d.lab, { box: true }); Wc(r, 3, k, { box: true, fmt: '#,##0.##', al: 'right' }); Wc(r, 4, c, { box: true, al: 'right' }); Wc(r, 5, g, { box: true, al: 'right' }); Wc(r, 6, t, { box: true, fmt: '#,##0.##', al: 'right' }); r++; });
    Wc(r, 2, 'YA EN PISO O RECORRIDO', { b: true, box: true }); Wc(r, 3, tk, { b: true, box: true, fmt: '#,##0.##', al: 'right' }); Wc(r, 4, tc, { b: true, box: true, al: 'right' }); Wc(r, 5, tg, { b: true, box: true, al: 'right' }); Wc(r, 6, tt, { b: true, box: true, fmt: '#,##0.##', al: 'right' }); r += 2;
    if (W.trips.length) {
      ws.mergeCells(r, 2, r, 6); Wc(r, 2, 'CAJAS POR VIAJE', { sz: 13, b: true, fill: GL }); r++;
      var sv = 0; W.trips.forEach(function (t, i) { Wc(r, 2, 'VIAJE ' + (i + 1) + ' · ' + t.cod, { box: true }); ws.mergeCells(r, 3, r, 5); Wc(r, 3, t.fecha ? dmy(t.fecha) : '', { box: true, al: 'center', color: '555555' }); Wc(r, 6, Math.round(t.cajas * 100) / 100, { box: true, fmt: '#,##0.##', al: 'right' }); sv += t.cajas; r++; });
      Wc(r, 2, 'Total de los viajes', { b: true, box: true }); Wc(r, 6, Math.round(sv * 100) / 100, { b: true, box: true, fmt: '#,##0.##', al: 'right' }); r++;
    }
    return ws;
  };

  /* ============ v8.13 · SEGUNDAS DE LA SEMANA (calca de la hoja SEMANA42 de «2 das 2026.xlsx») ============
     D: {num, yy, first:[{po,emb,cod,seg,cera,tor,des,c15}], maq:[{nombre,po,emb,cod,ent,c15,tor,des,jug}], noCorS, noCorC, closed}
     Mismo acomodo, títulos y fórmulas que la hoja original; cada bloque crece si la semana trae más renglones.
     Casillas amarillas: se llenan a mano (tickets de Juguera, plástico y Torreón no corrido). */
  XL.addSegSem = function (wb, D) {
    var GL = 'C1F0C8', YE = 'FFFF00', FA = 'Aptos Narrow', FD = 'Aptos Display';
    var NF = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', N0 = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)', MX = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';
    var ws = wb.addWorksheet(safeName(wb, 'SEMANA' + D.num), { views: [{ zoomScale: 90, showGridLines: true }], pageSetup: { orientation: 'portrait', paperSize: 1, fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    var widths = { A: 11.7, B: 17.1, C: 19.9, D: 13.7, E: 13, F: 13, G: 37.7, H: 13.7, I: 13, J: 13, K: 11, L: 9, M: 9 };
    Object.keys(widths).forEach(function (k) { ws.getColumn(k).width = widths[k]; });
    var s = { style: 'thin', color: { argb: 'FF000000' } }, dbl = { style: 'double', color: { argb: 'FF000000' } };
    function C(r, c, v, o) {
      o = o || {}; var cell = ws.getCell(r, c);
      if (v !== undefined && v !== null && v !== '') cell.value = v;
      cell.font = { name: o.fd ? FD : FA, size: o.sz || 11, bold: !!o.b };
      if (o.fill) cell.fill = fill(o.fill);
      if (o.fmt) cell.numFmt = o.fmt;
      cell.alignment = { horizontal: o.al || (typeof v === 'number' || (v && v.formula) ? 'center' : undefined), vertical: 'middle', wrapText: !!o.wrap };
      if (o.box) cell.border = { top: s, left: s, bottom: s, right: s };
      if (o.tot) cell.border = { top: s, bottom: dbl };
      return cell;
    }
    var F = function (f, r) { return { formula: f, result: r == null ? 0 : r }; };
    var R2 = function (x) { return Math.round((+x || 0) * 100) / 100; };
    var sum = function (a, k) { return a.reduce(function (t, x) { return t + (+x[k] || 0); }, 0); };
    var wkTxt = 'Semana ' + D.yy + (D.num < 10 ? '0' : '') + D.num;
    function head(r, cols, labels) { labels.forEach(function (h, i) { C(r, cols + i, h, { b: true, fd: true, sz: 14, fill: GL, box: true, al: 'center', wrap: true }); }); ws.getRow(r).height = 34.8; }
    function title(r, c1, c2, t) { ws.mergeCells(r, c1, r, c2); C(r, c1, t, { b: true, sz: 20, box: true, al: 'center' }); for (var c = c1 + 1; c <= c2; c++) ws.getCell(r, c).border = { top: s, bottom: s, right: c === c2 ? s : undefined }; ws.getRow(r).height = 48; }

    /* --- 1) Concentrado de segundas (primer corrida) --- */
    var A1 = D.first, nF = Math.max(12, A1.length), F0 = 4, FT = F0 + nF, FK = FT + 1, FG = FT + 2;
    title(2, 2, 9, 'Concetrado De Segundas (Primer Corrida) ' + wkTxt);
    head(3, 2, ['PO# Villamex', 'PO# Citricos JC', 'Viaje ', '2Das', '2Das Con Cera', 'Torreon', 'Desecho', 'Calidad 1.5']);
    var fk = ['seg', 'cera', 'tor', 'des', 'c15'];
    for (var i = 0; i < nF; i++) {
      var r = F0 + i, x = A1[i];
      C(r, 2, x ? x.po : null, { box: true, al: 'center' }); C(r, 3, x ? x.emb : null, { box: true, al: 'center' }); C(r, 4, x ? x.cod : null, { box: true, al: 'center' });
      fk.forEach(function (k, j) { C(r, 5 + j, x ? (+x[k] || 0) : (k === 'c15' ? 0 : null), { box: true, fmt: NF, al: 'center' }); });
    }
    var fs = fk.map(function (k) { return sum(A1, k); });
    C(FT, 4, 'Total Cajas', { b: true, al: 'right', tot: true });
    C(FK, 4, 'Kilos Por Caja'); C(FG, 4, 'Total Kilos', { b: true });
    fk.forEach(function (k, j) {
      var L = colL(5 + j);
      C(FT, 5 + j, F('SUM(' + L + F0 + ':' + L + (FT - 1) + ')', fs[j]), { b: true, fmt: NF, tot: true });
      C(FK, 5 + j, 18, { fmt: NF });
      C(FG, 5 + j, F('+' + L + FT + '*' + L + FK, fs[j] * 18), { b: true, fmt: NF });
    });

    /* --- 2) Segundas maquiladas --- */
    var A2 = D.maq, MTt = FG + 3, MH = MTt + 1, M0 = MH + 1, nM = Math.max(12, A2.length), MT = M0 + nM, MK = MT + 1, MG = MT + 2;
    title(MTt, 2, 9, 'Segundas Maquiladas ' + wkTxt);
    head(MH, 2, ['PO# Villamex', 'PO# Citricos JC', 'Viaje ', '2Das Maquiladas', 'Calidad 1.5', 'Torreon', 'Desecho', 'Juguera']);
    var mk = ['ent', 'c15', 'tor', 'des', 'jug'];
    for (i = 0; i < nM; i++) {
      r = M0 + i; x = A2[i];
      C(r, 1, x ? x.nombre : null);
      C(r, 2, x ? x.po : null, { box: true, al: 'center' }); C(r, 3, x ? x.emb : null, { box: true, al: 'center' }); C(r, 4, x ? x.cod : null, { box: true, al: 'center' });
      mk.forEach(function (k, j) { C(r, 5 + j, x ? (+x[k] || 0) : null, { box: true, fmt: NF, al: 'center' }); });
      C(r, 10, F('+E' + r + '-F' + r + '-G' + r + '-H' + r + '-I' + r, x ? R2(x.ent - x.c15 - x.tor - x.des - x.jug) : 0), { fmt: NF });
    }
    var ms = mk.map(function (k) { return sum(A2, k); }), mkg = (D.mkg && D.mkg.length === 5) ? D.mkg : [18, 17.8, 18, 18, 18], TARA = +D.tara > 0 ? +D.tara : 1.64; /* v8.16: pesos del catálogo «Pesos y equivalencias» */
    C(MT, 4, 'Total Cajas', { b: true, al: 'right', tot: true });
    C(MK, 4, 'Kilos Por Caja'); C(MG, 4, 'Total Kilos', { b: true });
    mk.forEach(function (k, j) {
      var L = colL(5 + j);
      C(MT, 5 + j, F('SUM(' + L + M0 + ':' + L + (MT - 1) + ')', ms[j]), { b: true, fmt: NF, tot: true });
      C(MK, 5 + j, mkg[j], { fmt: NF });
      C(MG, 5 + j, F('+' + L + MT + '*' + L + MK, R2(ms[j] * mkg[j])), { b: true, fmt: NF });
    });
    /* 2das de la primer corrida que no se maquilaron (como C36 / E38 de la hoja original) */
    C(MK, 3, F('+E' + FT + '-E' + MT, fs[0] - ms[0]), { fmt: NF });
    C(MG + 1, 5, F('+E' + FT + '-E' + MT, fs[0] - ms[0]), { fmt: NF });

    /* --- 3) Tickets de venta a Juguera (a mano) y peso a justificar --- */
    var TT = MG + 3, TH = TT + 1, T0 = TH + 1, TS = T0 + 5, TK = TS + 1, TG = TS + 2;
    ws.mergeCells(TT, 2, TT, 5); C(TT, 2, 'Tickets Venta Juguera ' + wkTxt, { b: true, sz: 20, box: true, al: 'center' }); ws.getRow(TT).height = 48;
    ws.mergeCells(TT, 7, TT, 10); C(TT, 7, 'Peso A Justificar', { b: true, sz: 20, box: true, al: 'center' });
    head(TH, 2, ['Ticket #', 'Kilos Ticket', 'Precio Por Kilo', 'Importe ']);
    head(TH, 7, ['Concepto / Productor', 'Cajas Juguera Maquila 18Kl', 'Cajas Juguera Maquila 27Kl', 'Total Kilos ']);
    for (i = 0; i < 5; i++) {
      r = T0 + i;
      C(r, 2, null, { box: true, fill: YE, al: 'center' }); C(r, 3, null, { box: true, fill: YE, fmt: NF, al: 'center' }); C(r, 4, null, { box: true, fill: YE, fmt: MX, al: 'center' });
      C(r, 5, F('+C' + r + '*D' + r, 0), { box: true, fmt: MX });
    }
    C(TS, 2, 'Total', { b: true, box: true });
    C(TS, 3, F('SUM(C' + T0 + ':C' + (TS - 1) + ')', 0), { b: true, box: true, fmt: NF });
    C(TS, 5, F('SUM(E' + T0 + ':E' + (TS - 1) + ')', 0), { b: true, box: true, fmt: MX, fill: '00B050' });
    C(TK, 2, 'DESCONTAR PLASTICO A EL TIKE', { fill: YE, al: 'center' });
    /* tara de plástico: cajas que se mandaron a Juguera (Justificación Juguera) × 1.64 kg; se corrige a mano si el ticket trae otra */
    var JG = TS + 5, JJ = JG + 11, JT = JJ + 10;
    var noS = +D.noCorS || 0, noC = +D.noCorC || 0, jugBox = ms[4] + fs[2] + ms[2] + noC + noS;
    C(TG, 2, F('+(H' + JT + '+I' + JT + ')*' + TARA, R2(jugBox * TARA)), { fill: YE, fmt: NF, al: 'center' });
    var pj = [['E', fs[0], 'E'], ['F', fs[1], 'F'], ['G', fs[2], 'G']];
    pj.forEach(function (q, j) {
      r = T0 + j;
      C(r, 7, F('+' + q[0] + '3', ['2Das', '2Das Con Cera', 'Torreon'][j]), { b: true, box: true, al: 'center' });
      C(r, 8, F(q[0] + FT, q[1]), { box: true, fmt: N0 }); C(r, 9, null, { box: true, fmt: NF });
      C(r, 10, F(q[0] + FG, q[1] * 18), { box: true, fmt: NF });
    });
    r = T0 + 3;
    C(r, 7, '2das no corridas (ya van en 2Das)', { box: true, al: 'center' }); C(r, 8, noS + noC, { box: true, fmt: NF }); C(r, 9, null, { box: true, fmt: NF });
    C(r, 10, F('+(H' + r + '*18)+(I' + r + '*27)', (noS + noC) * 18), { box: true, fmt: NF });
    r = T0 + 4;
    C(r, 7, 'Torreon no corridas', { box: true, al: 'center' }); C(r, 8, null, { box: true, fill: YE, fmt: NF }); C(r, 9, null, { box: true, fill: YE, fmt: NF });
    C(r, 10, F('+(H' + r + '*18)+(I' + r + '*27)', 0), { box: true, fmt: NF });
    /* el total no vuelve a sumar las 2das no corridas: ya están dentro de 2Das (en la hoja original se contaban dos veces) */
    C(TS, 7, 'Total Cajas', { b: true, box: true });
    ['H', 'I', 'J'].forEach(function (L, j) {
      var res = j === 0 ? fs[0] + fs[1] + fs[2] : j === 2 ? (fs[0] + fs[1] + fs[2]) * 18 : 0;
      C(TS, 8 + j, F('SUM(' + L + T0 + ':' + L + (T0 + 2) + ')+' + L + (T0 + 4), res), { b: true, box: true, fmt: NF });
    });
    C(TK, 7, 'Kilos Por Caja'); C(TK, 8, 18, { fmt: NF }); C(TK, 9, 27, { fmt: NF });
    C(TG, 7, 'Total Kilos', { b: true }); C(TG, 8, F('+H' + TS + '*H' + TK, (fs[0] + fs[1] + fs[2]) * 18), { b: true, fmt: NF }); C(TG, 9, F('+I' + TS + '*I' + TK, 0), { b: true, fmt: NF });

    /* --- 4) Justificación general: a dónde fue cada caja de segunda --- */
    ws.mergeCells(JG, 7, JG, 10); C(JG, 7, 'Justificacion General', { b: true, sz: 20, box: true, al: 'center' }); ws.getRow(JG).height = 48;
    head(JG + 1, 7, ['Concepto / Productor', 'Cajas Juguera Maquila 18Kl', 'Cajas Juguera Maquila 27Kl', 'Total Kilos ']);
    var jgl = [['F', 1], ['H', 3], ['G', 2], ['I', 4]];
    jgl.forEach(function (q, j) {
      r = JG + 2 + j;
      C(r, 7, F('+' + q[0] + MH, ['Calidad 1.5', 'Desecho', 'Torreon', 'Juguera'][j]), { b: true, box: true, al: 'center' });
      C(r, 8, F('+' + q[0] + MT, ms[q[1]]), { box: true, fmt: N0 }); C(r, 9, null, { box: true, fmt: NF });
      C(r, 10, F('+' + q[0] + MG, R2(ms[q[1]] * mkg[q[1]])), { box: true, fmt: NF });
    });
    r = JG + 6; C(r, 7, '2das no corridas', { box: true, al: 'center' }); C(r, 8, F('+H' + (T0 + 3), noS + noC), { box: true, fmt: NF }); C(r, 9, null, { box: true, fmt: NF }); C(r, 10, F('+(H' + r + '*18)+(I' + r + '*27)', (noS + noC) * 18), { box: true, fmt: NF });
    r = JG + 7; C(r, 7, 'Torriom no corridas', { box: true, al: 'center' }); C(r, 8, F('+H' + (T0 + 4), 0), { box: true, fmt: NF }); C(r, 9, F('+I' + (T0 + 4), 0), { box: true, fmt: NF }); C(r, 10, F('+(H' + r + '*18)+(I' + r + '*27)', 0), { box: true, fmt: NF });
    var gS = ms[1] + ms[3] + ms[2] + ms[4] + noS + noC, gK = R2(ms[1] * 17.8 + (ms[3] + ms[2] + ms[4]) * 18 + (noS + noC) * 18);
    C(JG + 8, 7, 'Total Cajas', { b: true, box: true });
    C(JG + 8, 8, F('SUM(H' + (JG + 2) + ':H' + (JG + 7) + ')', gS), { b: true, box: true, fmt: NF });
    C(JG + 8, 9, F('SUM(I' + (JG + 2) + ':I' + (JG + 7) + ')', 0), { b: true, box: true, fmt: NF });
    C(JG + 8, 10, F('SUM(J' + (JG + 2) + ':J' + (JG + 7) + ')', gK), { b: true, box: true, fmt: NF });
    C(JG + 9, 7, 'Kilos Por Caja'); C(JG + 9, 8, 18, { fmt: NF }); C(JG + 9, 9, 27, { fmt: NF });

    /* --- 5) Justificación Juguera: lo que se vendió a Juguera, en cajas y kilos --- */
    ws.mergeCells(JJ, 7, JJ, 10); C(JJ, 7, 'Justificacion Juguera', { b: true, sz: 20, box: true, al: 'center' }); ws.getRow(JJ).height = 48;
    head(JJ + 1, 7, ['Concepto / Productor', 'Cajas Juguera Maquila 18Kl', 'Cajas Juguera Maquila 27Kl', 'Total Kilos ']);
    r = JJ + 2; C(r, 7, F('+G' + (JG + 5), 'Juguera'), { b: true, box: true, al: 'center' }); C(r, 8, F('+H' + (JG + 5), ms[4]), { box: true, fmt: N0 }); C(r, 9, null, { box: true, fmt: NF }); C(r, 10, F('+J' + (JG + 5), ms[4] * 18), { box: true, fmt: NF });
    r = JJ + 3; C(r, 7, F('+G' + MH, 'Torreon'), { b: true, box: true, al: 'center' }); C(r, 8, F('+G' + FT + '+G' + MT, fs[2] + ms[2]), { box: true, fmt: N0 }); C(r, 9, null, { box: true, fmt: NF }); C(r, 10, F('+G' + FG + '+G' + MG, (fs[2] + ms[2]) * 18), { box: true, fmt: NF });
    /* 2das con cera que se fueron a Juguera sin correr (las que se corrieron ya van en la maquila) */
    r = JJ + 4; C(r, 7, F('+F3', '2Das Con Cera'), { b: true, box: true, al: 'center' }); C(r, 8, noC, { box: true, fmt: N0 }); C(r, 9, null, { box: true, fmt: NF }); C(r, 10, F('18*H' + r, noC * 18), { box: true, fmt: NF });
    r = JJ + 5; C(r, 7, '2DAS NO CORRIDAS ', { b: true, box: true, al: 'center' }); C(r, 8, noS, { box: true, fmt: N0 }); C(r, 9, null, { box: true, fmt: NF }); C(r, 10, F('18*H' + r, noS * 18), { box: true, fmt: NF });
    for (r = JJ + 6; r < JT; r++) { C(r, 7, null, { box: true }); C(r, 8, null, { box: true, fmt: N0 }); C(r, 9, null, { box: true, fmt: NF }); C(r, 10, null, { box: true, fmt: NF }); }
    var jK = (ms[4] + fs[2] + ms[2] + noC + noS) * 18;
    C(JT, 7, 'Total Cajas', { b: true, box: true });
    C(JT, 8, F('SUM(H' + (JJ + 2) + ':H' + (JT - 1) + ')', jugBox), { b: true, box: true, fmt: NF });
    C(JT, 9, F('SUM(I' + (JJ + 2) + ':I' + (JT - 1) + ')', 0), { b: true, box: true, fmt: NF });
    C(JT, 10, F('SUM(J' + (JJ + 2) + ':J' + (JT - 1) + ')', jK), { b: true, box: true, fmt: NF });
    C(JT + 1, 7, 'Kilos Por Caja'); C(JT + 1, 8, 18, { fmt: NF }); C(JT + 1, 9, 27, { fmt: NF });

    /* --- 6) Diferencia: kilos del ticket contra kilos que salieron a Juguera --- */
    var DF = JT + 3;
    ws.mergeCells(DF, 7, DF + 1, 7); C(DF, 7, 'Diferencia', { b: true, fd: true, sz: 14, box: true, al: 'center' });
    ['Kilos Venta', 'Kilos Maquila', 'Diferencia ', 'Cajas 18KL'].forEach(function (h, j) { C(DF, 8 + j, h, { b: true, box: true, al: 'center' }); });
    C(DF + 1, 8, F('+C' + TS, 0), { box: true, fmt: NF, fill: YE });
    C(DF + 1, 9, F('+J' + JT, jK), { box: true, fmt: NF });
    C(DF + 1, 10, F('+H' + (DF + 1) + '-I' + (DF + 1), -jK), { box: true, fmt: NF });
    C(DF + 1, 11, F('+J' + (DF + 1) + '/18', -jK / 18), { b: true, box: true, fmt: NF });
    if (!D.closed && (D.pend || 0) > 0) C(T0 + 3, 11, 'Incluye ' + D.pend + ' que se irán solas al cerrar la semana', { sz: 9 });
    return ws;
  };

  /* ============ v8.38 · INFORME SEMANAL DE SEGUNDAS (por viaje, con trazabilidad) ============
     R viene de segRep(wk) en app.js. Dos hojas:
     1) Concentrado: primera corrida, recepción, maquila (hoja 1.5), destino de la 1½, Juguera/Torreón/desecho, tickets y cuadre.
     2) Trazabilidad: qué pasó con las segundas de cada viaje, de dónde salió lo maquilado y cada entrega por productor.
     Las cifras salen de la app; las celdas amarillas se llenan a mano. Los totales son fórmulas vivas. */
  XL.addSegRep = function (wb, R) {
    var BR = R.emp === 'EC' ? '38582A' : '1B2C46', AC = R.emp === 'EC' ? 'F08030' : '98C858', HD = R.emp === 'EC' ? 'E3EBDD' : 'DFE5EF';
    var GR = 'F2F2F2', YE = 'FFF2CC', WT = 'FFFFFF', MU = '595959', OKC = '1E7B34', BAD = 'C00000';
    var N0 = '#,##0;-#,##0;"–"', N1 = '#,##0.0;-#,##0.0;"–"', MX = '"$"#,##0.00;-"$"#,##0.00;"–"';
    var ln = { style: 'thin', color: { argb: 'FFBFBFBF' } }, dk = { style: 'thin', color: { argb: 'FF' + BR } }, db = { style: 'double', color: { argb: 'FF' + BR } };
    var KS = +R.kg.seg || 18, K15 = +R.kg.c15 || 17.8, TARA = +R.kg.tara || 1.64;
    var wkTxt = 'Semana ' + R.num + ' · ' + R.anio + (R.span ? ' · ' + R.span : '');
    var estado = R.closed ? 'Semana cerrada' : 'Semana abierta: cifras preliminares hasta el cierre';
    var sumK = function (a, k) { return a.reduce(function (t, x) { return t + (+x[k] || 0); }, 0); };
    function sheet(name, land) {
      var ws = wb.addWorksheet(safeName(wb, name), { views: [{ showGridLines: false, zoomScale: 90 }], pageSetup: { orientation: land ? 'landscape' : 'portrait', paperSize: 1, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } } });
      ws.headerFooter.oddFooter = '&L&8' + (R.empFull || '') + ' · Informe de segundas · ' + wkTxt + '&R&8Página &P de &N';
      return ws;
    }
    function mk(ws) {
      function C(r, c, v, o) {
        o = o || {}; var cell = ws.getCell(r, c);
        if (v !== undefined && v !== null && v !== '') cell.value = v;
        cell.font = { name: FN, size: o.sz || 10, bold: !!o.b, italic: !!o.i, color: { argb: 'FF' + (o.col || '000000') } };
        if (o.fill) cell.fill = fill(o.fill);
        if (o.fmt) cell.numFmt = o.fmt;
        var num = typeof v === 'number' || (v && v.formula);
        cell.alignment = { horizontal: o.al || (num ? 'right' : 'left'), vertical: o.va || 'middle', wrapText: o.wrap !== false, indent: o.ind || 0 };
        if (o.bd) cell.border = o.bd === 'tot' ? { top: dk, bottom: db, left: ln, right: ln } : { top: ln, left: ln, bottom: ln, right: ln };
        return cell;
      }
      function M(r, c1, c2, v, o) { if (c2 > c1) ws.mergeCells(r, c1, r, c2); var x = C(r, c1, v, o); if (o && (o.bd || o.fill)) for (var c = c1 + 1; c <= c2; c++) { var y = ws.getCell(r, c); if (o.fill) y.fill = fill(o.fill); if (o.bd) y.border = o.bd === 'tot' ? { top: dk, bottom: db } : { top: ln, bottom: ln, right: c === c2 ? ln : undefined }; } return x; }
      return { C: C, M: M };
    }
    var F = function (f, r) { return { formula: f, result: r == null ? 0 : r }; };
    var R2 = function (x) { return Math.round((+x || 0) * 100) / 100; };
    var LAST = 12; /* columna L */
    function top(ws, H, titulo) {
      ws.getRow(1).height = 8;
      H.M(2, 2, LAST, (R.empFull || '').toUpperCase(), { b: true, sz: 11, col: BR });
      H.M(3, 2, LAST, titulo, { b: true, sz: 18, col: BR }); ws.getRow(3).height = 28;
      H.M(4, 2, LAST, wkTxt + '   ·   ' + estado, { sz: 10, col: MU });
      H.M(5, 2, LAST, 'Generado el ' + (R.genTxt || '') + ' · Caja de segunda = ' + KS + ' kg · Calidad 1½: RPC ' + K15 + ' kg, cartón ' + (+R.kg.ctn || 16) + ' kg · Celdas amarillas: se llenan a mano.', { sz: 8, i: true, col: MU });
      for (var c = 2; c <= LAST; c++) ws.getCell(5, c).border = { bottom: { style: 'medium', color: { argb: 'FF' + AC } } };
      return 7;
    }
    function sec(ws, H, r, t, sub) {
      H.M(r, 2, LAST, t, { b: true, sz: 12, col: WT, fill: BR }); ws.getRow(r).height = 22;
      if (sub) { H.M(r + 1, 2, LAST, sub, { sz: 9, i: true, col: MU }); ws.getRow(r + 1).height = Math.max(15, 13 * Math.ceil(sub.length / 150)); return r + 2; }
      return r + 1;
    }
    function head(ws, H, r, cols) { cols.forEach(function (h) { if (h.c2) H.M(r, h.c, h.c2, h.t, { b: true, sz: 9, fill: HD, bd: 1, al: 'center', col: BR }); else H.C(r, h.c, h.t, { b: true, sz: 9, fill: HD, bd: 1, al: 'center', col: BR }); }); ws.getRow(r).height = 30; }
    function note(ws, H, r, t, o) { H.M(r, 2, LAST, t, Object.assign({ sz: 9, i: true, col: MU }, o || {})); ws.getRow(r).height = Math.max(15, 13 * Math.ceil(String(t).length / 140)); return r + 1; }

    /* ================= HOJA 1 · CONCENTRADO ================= */
    var ws = sheet('SEMANA' + R.num, true), H = mk(ws);
    [2.5, 15, 15, 15, 12, 12, 12, 12, 12, 12, 12, 13].forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    var r = top(ws, H, 'Concentrado semanal de segundas');

    /* --- 1. primera corrida --- */
    r = sec(ws, H, r, '1.  Concentrado de segundas · primera corrida', 'Segundas que dejó cada viaje al empacar primera, por viaje. Torreón, Juguera y desecho son las casillas capturadas en el registro de empaque; «Se llevó» son las que el productor se llevó y no entran a segundas.');
    head(ws, H, r, [{ c: 2, t: 'PO# Villamex' }, { c: 3, t: 'PO# Cítricos JC' }, { c: 4, t: 'Viaje' }, { c: 5, t: 'Segundas' }, { c: 6, t: 'Con cera' }, { c: 7, t: 'Torreón' }, { c: 8, t: 'Juguera' }, { c: 9, t: 'Desecho' }, { c: 10, t: 'Calidad 1½ directa' }, { c: 11, t: 'Se llevó el productor' }, { c: 12, t: 'Total cajas' }]);
    var A1 = R.first, f0 = r + 1, fk = ['seg', 'cera', 'tor', 'jug', 'des', 'c15', 'llevo'];
    if (!A1.length) { r = note(ws, H, r + 1, 'Sin empaques con segundas en esta semana.'); }
    A1.forEach(function (x, i) {
      var rr = f0 + i, z = i % 2 ? GR : null;
      H.C(rr, 2, x.po || '—', { bd: 1, al: 'center', fill: z }); H.C(rr, 3, x.emb || '—', { bd: 1, al: 'center', fill: z }); H.C(rr, 4, x.cod, { bd: 1, al: 'center', b: true, fill: z });
      fk.forEach(function (k, j) { H.C(rr, 5 + j, +x[k] || 0, { bd: 1, fmt: N0, fill: z }); });
      H.C(rr, 12, F('SUM(E' + rr + ':K' + rr + ')', fk.reduce(function (t, k) { return t + (+x[k] || 0); }, 0)), { bd: 1, fmt: N0, b: true, fill: z });
    });
    var fT = f0 + A1.length, fKg = fT + 1, fKt = fT + 2;
    if (A1.length) {
      H.M(fT, 2, 4, 'Total cajas', { b: true, bd: 'tot', al: 'right' });
      for (var c = 5; c <= 12; c++) H.C(fT, c, F('SUM(' + colL(c) + f0 + ':' + colL(c) + (fT - 1) + ')', c === 12 ? fk.reduce(function (t, k) { return t + sumK(A1, k); }, 0) : sumK(A1, fk[c - 5])), { b: true, bd: 'tot', fmt: N0 });
      H.M(fKg, 2, 4, 'Kilos por caja', { al: 'right', col: MU, sz: 9 });
      fk.forEach(function (k, j) { H.C(fKg, 5 + j, k === 'c15' ? K15 : KS, { fmt: N1, col: MU, sz: 9 }); });
      H.M(fKt, 2, 4, 'Total kilos', { b: true, al: 'right' });
      fk.forEach(function (k, j) { var L = colL(5 + j); H.C(fKt, 5 + j, F(L + fT + '*' + L + fKg, R2(sumK(A1, k) * (k === 'c15' ? K15 : KS))), { b: true, fmt: N1 }); });
      H.C(fKt, 12, F('SUM(E' + fKt + ':K' + fKt + ')', R2(fk.reduce(function (t, k) { return t + sumK(A1, k) * (k === 'c15' ? K15 : KS); }, 0))), { b: true, fmt: N1 });
      r = fKt + 2;
    } else r += 1;

    /* --- 2. recepción --- */
    r = sec(ws, H, r, '2.  Llegó así a recepción · Entrada de Fruta', 'Fruta que se recibió ya como segunda, Torreón, Juguera o desecho, sin pasar por la línea. El desecho al recibir se captura en kilos; las cajas equivalentes son kilos ÷ ' + KS + '.');
    head(ws, H, r, [{ c: 2, t: 'Fecha' }, { c: 3, c2: 4, t: 'Productor' }, { c: 5, t: 'Segundas' }, { c: 6, t: 'Torreón' }, { c: 7, t: 'Juguera' }, { c: 8, t: 'Desecho (kg)' }, { c: 9, t: 'Desecho (cajas eq.)' }, { c: 10, c2: 12, t: 'Nota' }]);
    var e0 = r + 1, E = R.recep;
    if (!E.length) { H.M(e0, 2, LAST, 'Sin fruta recibida como segunda, Torreón, Juguera o desecho en esta semana.', { i: true, col: MU, bd: 1, al: 'center' }); r = e0 + 2; }
    else {
      E.forEach(function (x, i) {
        var rr = e0 + i, z = i % 2 ? GR : null;
        H.C(rr, 2, dmy(x.fecha), { bd: 1, al: 'center', fill: z }); H.M(rr, 3, 4, (x.codigo ? pad3(x.codigo) + ' · ' : '') + x.nombre, { bd: 1, fill: z });
        H.C(rr, 5, x.seg, { bd: 1, fmt: N0, fill: z }); H.C(rr, 6, x.tor, { bd: 1, fmt: N0, fill: z }); H.C(rr, 7, x.jug, { bd: 1, fmt: N0, fill: z });
        H.C(rr, 8, x.desKg, { bd: 1, fmt: N1, fill: z }); H.C(rr, 9, F('H' + rr + '/' + KS, x.desCj), { bd: 1, fmt: N1, fill: z });
        H.M(rr, 10, 12, x.todo ? 'Toda la entrega fue desecho' : (x.desKg ? 'Parte de la entrega fue desecho' : ''), { bd: 1, sz: 9, fill: z });
      });
      var eT = e0 + E.length; H.M(eT, 2, 4, 'Total', { b: true, bd: 'tot', al: 'right' });
      var ek = { 5: 'seg', 6: 'tor', 7: 'jug', 8: 'desKg', 9: 'desCj' };[5, 6, 7, 8, 9].forEach(function (c) { H.C(eT, c, F('SUM(' + colL(c) + e0 + ':' + colL(c) + (eT - 1) + ')', R2(sumK(E, ek[c]))), { b: true, bd: 'tot', fmt: c >= 8 ? N1 : N0 }); });
      r = eT + 2;
    }

    /* --- 3. maquila --- */
    r = sec(ws, H, r, '3.  Segundas maquiladas · hoja 1.5', 'Cada renglón de segunda corrido en una hoja 1.5. Remanente = lo que no se empacó y regresó a disponibles. Diferencia = entraron − (calidad 1½ + Torreón + desecho + Juguera + remanente): es la merma o el redondeo entre cajas de campo y cajas empacadas.');
    head(ws, H, r, [{ c: 2, t: 'PO# Villamex' }, { c: 3, t: 'PO# Cítricos JC' }, { c: 4, t: 'Viaje' }, { c: 5, t: 'Entraron' }, { c: 6, t: 'Calidad 1½' }, { c: 7, t: 'Torreón' }, { c: 8, t: 'Desecho' }, { c: 9, t: 'Juguera' }, { c: 10, t: 'Remanente' }, { c: 11, t: 'Diferencia (cajas)' }, { c: 12, t: 'Diferencia (kg)' }]);
    var A3 = R.maq2, m0 = r + 1, mk3 = ['ent', 'c15', 'tor', 'des', 'jug', 'rem'], warn = [];
    if (!A3.length) { H.M(m0, 2, LAST, 'Todavía no se corren segundas en una hoja 1.5 esta semana.', { i: true, col: MU, bd: 1, al: 'center' }); r = m0 + 2; }
    else {
      A3.forEach(function (x, i) {
        var rr = m0 + i, z = i % 2 ? GR : null, sin = !x.src.length;
        if (sin) { z = 'FCE4E4'; warn.push(x.cod + ': ' + x.ent + ' cajas capturadas a mano en la hoja 1.5, sin salir de ninguna segunda disponible. Revisar si están repetidas de otro viaje.'); }
        H.C(rr, 2, x.po || '—', { bd: 1, al: 'center', fill: z }); H.C(rr, 3, x.emb || '—', { bd: 1, al: 'center', fill: z }); H.C(rr, 4, x.cod + (sin ? ' ⚠' : ''), { bd: 1, al: 'center', b: true, fill: z });
        mk3.forEach(function (k, j) { H.C(rr, 5 + j, +x[k] || 0, { bd: 1, fmt: N0, fill: z }); });
        H.C(rr, 11, F('E' + rr + '-SUM(F' + rr + ':J' + rr + ')', x.dif), { bd: 1, fmt: N0, fill: z, col: Math.abs(x.dif) > 0.01 ? BAD : '000000' });
        H.C(rr, 12, x.kgDif, { bd: 1, fmt: N1, fill: z });
      });
      var mT = m0 + A3.length, mKg = mT + 1;
      H.M(mT, 2, 4, 'Total cajas', { b: true, bd: 'tot', al: 'right' });
      for (c = 5; c <= 12; c++) H.C(mT, c, F('SUM(' + colL(c) + m0 + ':' + colL(c) + (mT - 1) + ')', c <= 10 ? sumK(A3, mk3[c - 5]) : c === 11 ? sumK(A3, 'dif') : sumK(A3, 'kgDif')), { b: true, bd: 'tot', fmt: c === 12 ? N1 : N0 });
      H.M(mKg, 2, 4, 'Total kilos', { b: true, al: 'right' });
      H.C(mKg, 5, F('E' + mT + '*' + KS, sumK(A3, 'ent') * KS), { b: true, fmt: N1 });
      H.C(mKg, 6, R2(sumK(A3, 'kgOut')), { b: true, fmt: N1 });
      [7, 8, 9, 10].forEach(function (c) { H.C(mKg, c, F(colL(c) + mT + '*' + KS, sumK(A3, mk3[c - 5]) * KS), { b: true, fmt: N1 }); });
      H.C(mKg, 12, F('E' + mKg + '-SUM(F' + mKg + ':J' + mKg + ')', sumK(A3, 'kgDif')), { b: true, fmt: N1 });
      r = note(ws, H, mKg + 1, 'Kilos de calidad 1½ = cajas RPC × ' + K15 + ' kg + cartón × ' + (+R.kg.ctn || 16) + ' kg, según lo empacado en cada hoja 1.5.');
      warn.forEach(function (w) { r = note(ws, H, r, '⚠ ' + w, { col: BAD, i: false, b: true }); });
      r += 1;
    }

    /* --- 4. destino de la 1½ --- */
    var T4 = R.trips.filter(function (t) { return t.dest15; });
    ws.getRow(r - 1).addPageBreak();
    r = sec(ws, H, r, '4.  Destino de la calidad 1½ empacada', 'A dónde fueron las cajas 1½ que empacó la hoja 1.5 de cada viaje: en sus propias tarimas, enviadas a otro viaje, en tarimas armadas que siguen en el almacén o sueltas en piso.');
    head(ws, H, r, [{ c: 2, c2: 3, t: 'Viaje (hoja 1.5)' }, { c: 4, t: 'Empacadas' }, { c: 5, t: 'En sus tarimas' }, { c: 6, t: 'A otro viaje' }, { c: 7, c2: 8, t: 'Viaje que las recibió' }, { c: 9, t: 'Almacén (sin salir)' }, { c: 10, t: 'En piso' }, { c: 11, c2: 12, t: 'Recibidas de otros viajes' }]);
    var d0 = r + 1;
    if (!T4.length) { H.M(d0, 2, LAST, 'Sin calidad 1½ empacada esta semana.', { i: true, col: MU, bd: 1, al: 'center' }); r = d0 + 2; }
    else {
      T4.forEach(function (t, i) {
        var rr = d0 + i, d = t.dest15, z = i % 2 ? GR : null;
        H.M(rr, 2, 3, t.cod, { bd: 1, b: true, al: 'center', fill: z }); H.C(rr, 4, d.pk, { bd: 1, fmt: N0, fill: z }); H.C(rr, 5, d.own, { bd: 1, fmt: N0, fill: z }); H.C(rr, 6, d.sn, { bd: 1, fmt: N0, fill: z });
        H.M(rr, 7, 8, Object.keys(d.sent).map(function (k) { return k + ' (' + d.sent[k] + ')'; }).join(', ') || '—', { bd: 1, al: 'center', fill: z, sz: 9 });
        H.C(rr, 9, d.alm, { bd: 1, fmt: N0, fill: z }); H.C(rr, 10, F('D' + rr + '-E' + rr + '-F' + rr + '-I' + rr, d.piso), { bd: 1, fmt: N0, fill: z, col: d.piso > 0 ? BAD : '000000' }); H.M(rr, 11, 12, d.rec, { bd: 1, fmt: N0, fill: z });
      });
      var dT = d0 + T4.length; H.M(dT, 2, 3, 'Total', { b: true, bd: 'tot', al: 'right' });
      var dk4 = { 4: 'pk', 5: 'own', 6: 'sn', 9: 'alm', 10: 'piso', 11: 'rec' };[4, 5, 6, 9, 10, 11].forEach(function (c) { H.C(dT, c, F('SUM(' + colL(c) + d0 + ':' + colL(c) + (dT - 1) + ')', T4.reduce(function (t, x) { return t + (+x.dest15[dk4[c]] || 0); }, 0)), { b: true, bd: 'tot', fmt: N0 }); });
      r = dT + 2;
    }

    /* --- 5. Juguera, Torreón y desecho --- */
    r = sec(ws, H, r, '5.  Destino final: Juguera, Torreón y desecho', 'Cada origen por separado: casillas del registro de empaque (primera corrida), casillas de la hoja 1.5 (maquila), recepción, botón «Enviar a Juguera» y cierre de semana. Las cajas de 27 kg y el Torreón no corrido se capturan a mano.');
    head(ws, H, r, [{ c: 2, c2: 5, t: 'Concepto' }, { c: 6, t: 'Cajas 18 kg' }, { c: 7, t: 'Cajas 27 kg' }, { c: 8, t: 'Total kilos' }, { c: 9, c2: 12, t: 'De dónde sale' }]);
    r += 1;
    var J = R.J, TO = R.TO, DS = R.DS, Q = R.Q, gT = {};
    function grp(key, title, items, info) {
      H.M(r, 2, LAST, title, { b: true, fill: HD, col: BR, bd: 1 }); r++;
      var g0 = r, gs = 0;
      items.forEach(function (it) { if (!it[3]) gs += +it[1] || 0;
        H.M(r, 2, 5, it[0], { bd: 1, ind: 1 });
        H.C(r, 6, it[3] ? null : it[1], { bd: 1, fmt: it[4] || N0, fill: it[3] ? YE : null }); if (it[3]) ws.getCell(r, 6).value = null;
        H.C(r, 7, null, { bd: 1, fmt: N0, fill: it[3] ? YE : null });
        H.C(r, 8, F('F' + r + '*' + KS + '+G' + r + '*27', R2((it[3] ? 0 : +it[1] || 0) * KS)), { bd: 1, fmt: N1 });
        H.M(r, 9, 12, it[2] || '', { bd: 1, sz: 8, col: MU });
        r++;
      });
      H.M(r, 2, 5, 'Total ' + title.toLowerCase() + (info ? ' (ya enviado)' : ''), { b: true, bd: 'tot', al: 'right' });
      ['F', 'G', 'H'].forEach(function (L, j) { H.C(r, 6 + j, F('SUM(' + L + g0 + ':' + L + (r - 1) + ')', j === 0 ? R2(gs) : j === 2 ? R2(gs * KS) : 0), { b: true, bd: 'tot', fmt: j === 2 || (j === 0 && key === 'des') ? N1 : N0 }); });
      gT[key] = r; r++;
      (info || []).forEach(function (it) {
        H.M(r, 2, 5, it[0], { ind: 2, i: true, col: MU, sz: 9 }); H.C(r, 6, it[1], { fmt: N0, i: true, col: MU, sz: 9 }); H.C(r, 8, R2((+it[1] || 0) * KS), { fmt: N1, i: true, col: MU, sz: 9 }); H.M(r, 9, 12, it[2] || '', { sz: 8, col: MU, i: true }); r++;
      });
      r++;
    }
    grp('jug', 'Juguera', [
      ['Casillas del registro de empaque (primera corrida)', J.casEmp, 'Casilla «Juguera» de cada productor'],
      ['Casillas de la hoja 1.5 (maquila)', J.casMaq, 'Casilla «Juguera» del renglón de segundas'],
      ['Llegó como Juguera a recepción', J.recep, 'Entrada de Fruta'],
      ['Enviadas desde el almacén («Enviar a Juguera»)', J.boton, 'Segundas sin maquilar'],
      ['Al cerrar la semana', J.cierre, R.closed ? 'Lo que el cierre mandó a Juguera' : 'La semana sigue abierta'],
      ['Cajas de 27 kg (a mano)', 0, 'Si se mandan cajas grandes', true]
    ], R.closed ? null : [
      ['Pendiente: segundas disponibles de esta semana', Q.av, 'No suman: al cerrar la semana se elige su destino'],
      ['Pendiente: disponibles de semanas anteriores', Q.prev || 0, 'No suman: siguen sin destino; se resuelven al cerrar su semana']
    ]);
    grp('tor', 'Torreón', [
      ['Casillas del registro de empaque (primera corrida)', TO.casEmp, 'Casilla «Torreón» de cada productor'],
      ['Casillas de la hoja 1.5 (maquila)', TO.casMaq, 'Casilla «Torreón» del renglón de segundas'],
      ['Llegó como Torreón a recepción', TO.recep, 'Entrada de Fruta que no se cargó al registro'],
      ['Vendidas a Torreón desde el almacén', TO.alm || 0, 'Botón «Venta Torreón»'],
      ['Torreón no corrido (a mano)', 0, 'Si se mandó sin pasar por la línea', true]
    ]);
    grp('des', 'Desecho', [
      ['Casillas del registro de empaque (primera corrida)', DS.casEmp, 'Casilla «Desecho» de cada productor'],
      ['Casillas de la hoja 1.5 (maquila)', DS.casMaq, 'Casilla «Desecho» del renglón de segundas'],
      ['Desecho al recibir (cajas equivalentes)', R2(DS.recepCj), fmtN(DS.recepKg) + ' kg en Entrada de Fruta', false, N1]
    ]);
    r = note(ws, H, r - 1, 'Desecho vaciado del empaque: ' + fmtN(DS.vac) + ' cajas' + (DS.pend ? ' · por vaciar: ' + fmtN(DS.pend) + (R.closed ? '' : ' (se vacían solas al cerrar la semana)') : '') + '.') + 1;

    /* --- 6. tickets de Juguera y diferencia --- */
    ws.getRow(r - 1).addPageBreak();
    r = sec(ws, H, r, '6.  Tickets de venta a Juguera', 'Se llenan a mano con los tickets de báscula de Juguera. El plástico se descuenta con ' + TARA + ' kg por caja enviada.');
    head(ws, H, r, [{ c: 2, t: 'Ticket #' }, { c: 3, t: 'Fecha' }, { c: 4, t: 'Kilos ticket' }, { c: 5, t: 'Precio por kilo' }, { c: 6, t: 'Importe' }, { c: 7, c2: 12, t: 'Observaciones' }]);
    var t0 = r + 1;
    for (var i = 0; i < 6; i++) { var rr = t0 + i; [2, 3, 4, 5].forEach(function (c) { H.C(rr, c, null, { bd: 1, fill: YE, fmt: c === 4 ? N1 : c === 5 ? MX : undefined, al: 'center' }); }); H.C(rr, 6, F('D' + rr + '*E' + rr, 0), { bd: 1, fmt: MX }); H.M(rr, 7, 12, null, { bd: 1, fill: YE }); }
    var tT = t0 + 6; H.M(tT, 2, 3, 'Total', { b: true, bd: 'tot', al: 'right' }); H.C(tT, 4, F('SUM(D' + t0 + ':D' + (tT - 1) + ')', 0), { b: true, bd: 'tot', fmt: N1 }); H.C(tT, 5, null, { bd: 'tot' }); H.C(tT, 6, F('SUM(F' + t0 + ':F' + (tT - 1) + ')', 0), { b: true, bd: 'tot', fmt: MX });
    r = tT + 2;
    head(ws, H, r, [{ c: 2, c2: 3, t: 'Diferencia contra el ticket' }, { c: 4, t: 'Kilos ticket' }, { c: 5, t: 'Plástico (kg)' }, { c: 6, t: 'Kilos netos' }, { c: 7, t: 'Kilos enviados' }, { c: 8, t: 'Diferencia (kg)' }, { c: 9, t: 'Equivale a cajas' }, { c: 10, c2: 12, t: 'Resultado' }]);
    var rd = r + 1, jR = gT.jug;
    H.M(rd, 2, 3, 'Juguera', { b: true, bd: 1, al: 'center' });
    var jS = J.casEmp + J.casMaq + J.recep + J.boton + J.cierre, pl = R2(jS * TARA), jK = R2(jS * KS);
    H.C(rd, 4, F('D' + tT, 0), { bd: 1, fmt: N1 }); H.C(rd, 5, F('(F' + jR + '+G' + jR + ')*' + TARA, pl), { bd: 1, fmt: N1, fill: YE }); H.C(rd, 6, F('D' + rd + '-E' + rd, -pl), { bd: 1, fmt: N1 });
    H.C(rd, 7, F('H' + jR, jK), { bd: 1, fmt: N1 }); H.C(rd, 8, F('F' + rd + '-G' + rd, R2(-pl - jK)), { bd: 1, fmt: N1, b: true }); H.C(rd, 9, F('H' + rd + '/' + KS, R2((-pl - jK) / KS)), { bd: 1, fmt: N1 });
    H.M(rd, 10, 12, F('IF(D' + rd + '=0,"Falta capturar el ticket",IF(ABS(H' + rd + ')<=' + KS + ',"Cuadra",IF(H' + rd + '>0,"El ticket pesa más","El ticket pesa menos")))', 'Falta capturar el ticket'), { bd: 1, b: true, al: 'center' });
    r = rd + 2;

    /* --- 7. cuadre --- */
    r = sec(ws, H, r, '7.  Cuadre de la semana', 'Todo lo que entró a segundas tiene que tener un destino. El remanente de una hoja 1.5 vuelve a entrar como disponible, por eso aparece de los dos lados.');
    head(ws, H, r, [{ c: 2, c2: 5, t: 'Entró a segundas' }, { c: 6, t: 'Cajas' }, { c: 7, c2: 10, t: 'Salió de segundas' }, { c: 11, t: 'Cajas' }, { c: 12, t: '' }]);
    var q0 = r + 1, INr = [['Primera corrida (incluye SOBRANTES y REEMPAQUE)', Q.gen], ['Con cera', Q.cera], ['Llegaron como segunda a recepción', Q.recep], ['Remanente de hojas 1.5', Q.rem], ['Bajaron a segunda en el almacén', Q.alm || 0], ['Desecho reclasificado a segunda', Q.rc || 0]],
      OUTr = [['Maquiladas en hoja 1.5', Q.maq], ['A Juguera desde el almacén', Q.jugB], ['A Juguera al cerrar la semana', Q.jugC], ['Venta nacional', Q.nac || 0], ['Venta Torreón (almacén)', Q.tor || 0], ['Pasaron a desecho', Q.des || 0], ['Bajas en almacén (merma o ajuste)', Q.mer], ['Vendidas al cierre (versiones anteriores)', Q.vend], ['Disponibles', Q.av]];
    var nq = Math.max(INr.length, OUTr.length);
    for (i = 0; i < nq; i++) {
      rr = q0 + i; var a = INr[i], b = OUTr[i];
      H.M(rr, 2, 5, a ? a[0] : null, { bd: 1, ind: 1 }); H.C(rr, 6, a ? a[1] : null, { bd: 1, fmt: N0 });
      H.M(rr, 7, 10, b ? b[0] : null, { bd: 1, ind: 1 }); H.C(rr, 11, b ? b[1] : null, { bd: 1, fmt: N0 });
    }
    var qT = q0 + nq; H.M(qT, 2, 5, 'Total entró', { b: true, bd: 'tot', al: 'right' }); H.C(qT, 6, F('SUM(F' + q0 + ':F' + (qT - 1) + ')', Q.ent), { b: true, bd: 'tot', fmt: N0 });
    H.M(qT, 7, 10, 'Total salió', { b: true, bd: 'tot', al: 'right' }); H.C(qT, 11, F('SUM(K' + q0 + ':K' + (qT - 1) + ')', Q.sal), { b: true, bd: 'tot', fmt: N0 });
    H.C(qT, 12, F('IF(ABS(F' + qT + '-K' + qT + ')<0.01,"Cuadra","Revisar: "&TEXT(F' + qT + '-K' + qT + ',"#,##0"))', Math.abs(Q.otros) < 0.01 ? 'Cuadra' : 'Revisar'), { b: true, al: 'center', col: Math.abs(Q.otros) < 0.01 ? OKC : BAD });
    r = qT + 1;
    if (Q.llevo) r = note(ws, H, r, 'Aparte, los productores se llevaron ' + fmtN(Q.llevo) + ' cajas de segunda (no entran a segundas del empaque).');
    if (Q.sinSrc) r = note(ws, H, r, '⚠ ' + fmtN(Q.sinSrc) + ' cajas maquiladas en hojas 1.5 no salieron de ninguna segunda disponible (capturadas a mano). No entran en este cuadre: revisar la sección 3.', { col: BAD, i: false, b: true });

    /* --- 8. ventas de segunda (v8.42) --- */
    var VT = R.VT;
    if (VT) {
      r += 1;
      r = sec(ws, H, r, '8.  Ventas de segunda', 'Lo que se vendió de segunda en la semana, con el precio de la semana (engrane «Precios de venta» del Almacén). La Juguera paga por kilo; si hay liquidación registrada en Por cobrar, manda la liquidación.');
      head(ws, H, r, [{ c: 2, c2: 5, t: 'Destino' }, { c: 6, t: 'Cajas' }, { c: 7, t: 'Kilos' }, { c: 8, t: 'Importe (MXN)' }, { c: 9, c2: 12, t: 'Nota' }]);
      var v0 = r + 1, VR = [['Venta nacional', VT.nac.n, VT.nac.kg, VT.nac.imp, VT.nac.sin ? 'Falta precio de venta nacional' : ''], ['Torreón · almacén', VT.tor.n, VT.tor.kg, VT.tor.imp, VT.tor.sin ? 'Falta precio de Torreón' : ''], ['Torreón · casillas', VT.torCas.n, null, VT.torCas.imp, VT.torCas.sin ? 'Falta precio de Torreón' : ''], ['Juguera', VT.jug.n, VT.jug.kg, VT.jug.imp, VT.jug.liq ? 'Con liquidación registrada' : VT.jug.sin ? 'Falta precio por kilo de Juguera' : '']];
      VR.forEach(function (x, i) { var rr = v0 + i; H.M(rr, 2, 5, x[0], { bd: 1, ind: 1 }); H.C(rr, 6, x[1], { bd: 1, fmt: N0 }); H.C(rr, 7, x[2] == null ? null : R2(x[2]), { bd: 1, fmt: N1 }); H.C(rr, 8, R2(x[3]), { bd: 1, fmt: MX }); H.M(rr, 9, 12, x[4], { bd: 1, sz: 9, col: x[4] && /Falta/.test(x[4]) ? BAD : MU }); });
      var vT = v0 + VR.length; H.M(vT, 2, 5, 'Total vendido', { b: true, bd: 'tot', al: 'right' }); H.C(vT, 6, F('SUM(F' + v0 + ':F' + (vT - 1) + ')', VR.reduce(function (t, x) { return t + (+x[1] || 0); }, 0)), { b: true, bd: 'tot', fmt: N0 }); H.C(vT, 7, null, { bd: 'tot' }); H.C(vT, 8, F('SUM(H' + v0 + ':H' + (vT - 1) + ')', R2(VT.imp)), { b: true, bd: 'tot', fmt: MX }); H.M(vT, 9, 12, null, { bd: 'tot' });
      r = vT + 1;
      Object.keys(VT.pr || {}).forEach(function (e) { var P = VT.pr[e]; r = note(ws, H, r, 'Precios ' + e + ': venta nacional ' + (P.nac != null ? '$' + R2(P.nac) : '—') + ' por caja · Torreón ' + (P.tor != null ? '$' + R2(P.tor) : '—') + ' por caja · Juguera ' + (P.jugKg != null ? '$' + R2(P.jugKg) + ' por kg' : '—') + '.'); });
      if (Q.des || Q.rc) r = note(ws, H, r, 'Pasaron a desecho ' + fmtN(Q.des || 0) + ' cajas de segunda; se reclasificaron de desecho a segunda ' + fmtN(Q.rc || 0) + ' cajas.');
    }

    /* ================= HOJA 2 · TRAZABILIDAD ================= */
    var w2 = sheet('TRAZABILIDAD S' + R.num, false), H2 = mk(w2);
    [2.5, 16, 14, 14, 12, 12, 12, 12, 12, 12, 12, 12].forEach(function (w, i) { w2.getColumn(i + 1).width = w; });
    r = top(w2, H2, 'Trazabilidad de segundas por viaje');
    r = sec(w2, H2, r, 'A.  Qué pasó con las segundas de cada viaje', 'Resumen en palabras de cómo se movieron las segundas de cada viaje, desde la línea hasta su destino.');
    R.trips.forEach(function (t) {
      H2.M(r, 2, LAST, t.cod + (t.emb ? '   ·   PO# Cítricos JC ' + t.emb : '') + (t.po ? '   ·   PO# Villamex ' + t.po : '') + (t.fecha ? '   ·   ' + dmy(t.fecha) : ''), { b: true, col: BR, fill: HD }); w2.getRow(r).height = 18; r++;
      t.notes.forEach(function (n) { H2.M(r, 2, LAST, '•  ' + n, { sz: 9, col: /⚠/.test(n) ? BAD : '000000', b: /⚠/.test(n), ind: 1 }); w2.getRow(r).height = Math.max(15, 13 * Math.ceil(n.length / 120)); r++; });
      r++;
    });
    r = sec(w2, H2, r, 'B.  Procedencia de lo maquilado en cada hoja 1.5', 'De qué productor, entrega y viaje salió cada caja de segunda que se corrió en una hoja 1.5.');
    head(w2, H2, r, [{ c: 2, t: 'Hoja 1.5 de' }, { c: 3, c2: 5, t: 'Productor' }, { c: 6, t: 'Entrega' }, { c: 7, t: 'Viaje de origen' }, { c: 8, t: 'Tipo' }, { c: 9, t: 'Cajas' }, { c: 10, t: 'Kilos' }, { c: 11, c2: 12, t: 'Observación' }]);
    var p0 = r + 1, P = R.proc.slice().sort(function (a, b) { return String(a.a).localeCompare(String(b.a)) || String(a.de).localeCompare(String(b.de)) || String(a.nombre).localeCompare(String(b.nombre)); });
    if (!P.length) { H2.M(p0, 2, LAST, 'Sin segundas maquiladas esta semana.', { i: true, col: MU, bd: 1, al: 'center' }); r = p0 + 2; }
    else {
      P.forEach(function (x, i) { var rr = p0 + i, z = i % 2 ? GR : null;
        H2.C(rr, 2, x.a, { bd: 1, b: true, al: 'center', fill: z }); H2.M(rr, 3, 5, x.nombre, { bd: 1, fill: z }); H2.C(rr, 6, dmy(x.fecha), { bd: 1, al: 'center', fill: z });
        H2.C(rr, 7, x.de, { bd: 1, al: 'center', fill: z }); H2.C(rr, 8, x.cera ? 'Con cera' : 'Segunda', { bd: 1, al: 'center', fill: z, sz: 9 }); H2.C(rr, 9, x.n, { bd: 1, fmt: N0, fill: z });
        H2.C(rr, 10, F('I' + rr + '*' + KS, x.n * KS), { bd: 1, fmt: N1, fill: z }); H2.M(rr, 11, 12, x.de === x.a ? 'Del mismo viaje' : x.de === 'Recepción' ? 'Llegó como segunda' : 'Vino de otro viaje', { bd: 1, sz: 8, col: MU, fill: z }); });
      var pT = p0 + P.length; H2.M(pT, 2, 8, 'Total maquilado', { b: true, bd: 'tot', al: 'right' }); H2.C(pT, 9, F('SUM(I' + p0 + ':I' + (pT - 1) + ')', sumK(P, 'n')), { b: true, bd: 'tot', fmt: N0 }); H2.C(pT, 10, F('SUM(J' + p0 + ':J' + (pT - 1) + ')', sumK(P, 'n') * KS), { b: true, bd: 'tot', fmt: N1 });
      r = pT + 2;
    }
    w2.getRow(r - 1).addPageBreak();
    r = sec(w2, H2, r, 'C.  Por productor y entrega', 'Cada entrega de segunda de la semana: cuántas salieron, cuántas se maquilaron, cuántas fueron a Juguera y cuántas siguen disponibles.');
    head(w2, H2, r, [{ c: 2, c2: 3, t: 'Productor' }, { c: 4, t: 'Entrega' }, { c: 5, t: 'Viaje de origen' }, { c: 6, t: 'Salieron' }, { c: 7, t: 'Maquiladas' }, { c: 8, t: 'Juguera' }, { c: 9, t: 'Bajas' }, { c: 10, t: 'Disponibles' }, { c: 11, c2: 12, t: 'Origen' }]);
    var l0 = r + 1, LL = R.lots2.slice().sort(function (a, b) { return String(a.nombre).localeCompare(String(b.nombre)) || String(a.fecha).localeCompare(String(b.fecha)); });
    LL.forEach(function (l, i) { var rr = l0 + i, z = i % 2 ? GR : null;
      H2.M(rr, 2, 3, (l.codigo ? pad3(l.codigo) + ' · ' : '') + l.nombre, { bd: 1, fill: z }); H2.C(rr, 4, dmy(l.fecha), { bd: 1, al: 'center', fill: z }); H2.C(rr, 5, l.cod || 'Recepción', { bd: 1, al: 'center', fill: z });
      H2.C(rr, 6, l.tot, { bd: 1, fmt: N0, fill: z }); H2.C(rr, 7, l.proc, { bd: 1, fmt: N0, fill: z }); H2.C(rr, 8, l.jug, { bd: 1, fmt: N0, fill: z }); H2.C(rr, 9, l.mer, { bd: 1, fmt: N0, fill: z });
      H2.C(rr, 10, F('F' + rr + '-G' + rr + '-H' + rr + '-I' + rr, l.av), { bd: 1, fmt: N0, fill: z, b: l.av > 0 });
      H2.M(rr, 11, 12, (l.cera ? 'Con cera · ' : '') + (l.rem ? 'Remanente de hoja 1.5' : l.alm ? 'Bajó a segunda en almacén' : l.src === 'e' ? 'Entrada de Fruta' : 'Primera corrida'), { bd: 1, sz: 8, col: MU, fill: z }); });
    if (LL.length) { var lT = l0 + LL.length; H2.M(lT, 2, 5, 'Total', { b: true, bd: 'tot', al: 'right' }); var lk = { 6: 'tot', 7: 'proc', 8: 'jug', 9: 'mer', 10: 'av' };[6, 7, 8, 9, 10].forEach(function (c) { H2.C(lT, c, F('SUM(' + colL(c) + l0 + ':' + colL(c) + (lT - 1) + ')', sumK(LL, lk[c])), { b: true, bd: 'tot', fmt: N0 }); }); r = lT + 2; }
    else r = l0 + 1;
    note(w2, H2, r, 'Disponibles = salieron − maquiladas − Juguera − bajas. Lo disponible al cerrar la semana se va solo a Juguera y queda en la sección 5 del concentrado.');
    /* v8.43 · Evidencias de cargas: cada foto de fruta y ticket de las salidas de la semana (venta nacional, Torreón, Juguera, desecho) */
    if (R.evid && R.evid.length) {
      var w3 = sheet('Evidencias de cargas', true), H3 = mk(w3), TH = R.evTh || { get: function () { return null; } };
      [2, 17, 15, 14, 16, 28, 9, 12, 16, 30].forEach(function (wd, i) { w3.getColumn(i + 1).width = wd; });
      H3.M(2, 2, 10, 'Evidencias de cargas · ' + wkTxt, { b: true, sz: 14, col: BR, wrap: false });
      H3.M(3, 2, 10, 'Fotos de la fruta y tickets que respaldan cada salida de segunda (venta nacional, Torreón, Juguera y desecho). Las fotos originales (WebP) se ven y descargan en la app, en Recepción › Segundas.', { i: true, col: MU, sz: 9 });
      w3.getRow(3).height = 28;
      var hd = ['Fecha y hora', 'Destino', 'Folio / ticket', 'Viajes', 'Productores', 'Cajas', 'Foto', 'Registró', 'Imagen'];
      hd.forEach(function (t, i) { H3.C(5, i + 2, t, { b: true, col: 'FFFFFF', fill: BR, bd: 1, al: 'center', sz: 9 }); });
      var rr = 6, DT = { nac: 'Venta nacional', tor: 'Venta Torreón', jug: 'Juguera', des: 'Desecho' };
      R.evid.forEach(function (g) {
        g.fotos.forEach(function (f, j) {
          var z = g.anul ? 'F8D7DA' : (j ? null : HD), d = new Date(f.at);
          var ts = String(d.getDate()).padStart(2, '0') + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + d.getFullYear() + ' ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
          H3.C(rr, 2, ts, { bd: 1, fill: z, sz: 9, al: 'center' });
          H3.C(rr, 3, (DT[g.t] || g.t) + (g.anul ? ' (anulada)' : ''), { bd: 1, fill: z, b: !j, sz: 9 });
          H3.C(rr, 4, g.tk || '—', { bd: 1, fill: z, sz: 9, al: 'center' });
          H3.C(rr, 5, (g.cods || []).join(', '), { bd: 1, fill: z, sz: 9 });
          H3.C(rr, 6, (g.prods || []).join(', '), { bd: 1, fill: z, sz: 8 });
          H3.C(rr, 7, j ? null : g.n, { bd: 1, fill: z, fmt: N0 });
          H3.C(rr, 8, f.k === 'tk' ? 'Ticket' : 'Fruta', { bd: 1, fill: z, sz: 9, al: 'center' });
          H3.C(rr, 9, f.by || '', { bd: 1, fill: z, sz: 8 });
          H3.C(rr, 10, null, { bd: 1, fill: z });
          var t = TH.get(f.id);
          if (t && t.b64) {
            var mh = 150, mw = 200, k = Math.min(mw / t.w, mh / t.h), iw = Math.round(t.w * k), ih = Math.round(t.h * k);
            var id = wb.addImage({ base64: t.b64, extension: 'jpeg' });
            w3.addImage(id, { tl: { col: 9.06, row: rr - 1 + 0.06 }, ext: { width: iw, height: ih }, editAs: 'oneCell' });
            w3.getRow(rr).height = Math.max(30, ih * 0.75 + 8);
          } else { H3.C(rr, 10, 'Sin imagen (no se pudo descargar)', { bd: 1, fill: z, i: true, col: MU, sz: 8 }); w3.getRow(rr).height = 30; }
          rr++;
        });
      });
    }
    return ws;
  };
  function pad3(n) { n = String(n == null ? '' : n); while (n.length < 3) n = '0' + n; return n; }
  function fmtN(n) { n = Math.round((+n || 0) * 100) / 100; return n.toLocaleString('es-MX'); }

  XL.CAL = CAL;
  root.XL = XL;
  if (typeof module !== 'undefined' && module.exports) module.exports = XL;
})(typeof window !== 'undefined' ? window : globalThis);
