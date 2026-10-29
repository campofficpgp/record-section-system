/* =====================================================================
   PRINT / PDF / WORD / EXCEL / CSV on every register (2026-10-28)
   The same export engine as the Bill Register: it reads the table that is on screen (filters respected) and builds the file.
   Nothing here changes how data is saved. The "Prepared by" stamp is optional (tick box, off by default).
   ===================================================================== */
(function(){
  "use strict";
  function $(id){ return document.getElementById(id); }
  function bi(en, ur){ return '<span class="en">'+en+'</span><span class="ur">'+ur+'</span>'; }
  function say(en, ur){ try{ toast(document.body.classList.contains("lang-ur") ? ur : en); }catch(e){} }
  function closeMenus(except){
    document.querySelectorAll(".br-menu.open").forEach(function(m){ if(m !== except) m.classList.remove("open"); });
  }
  document.addEventListener("click", function(e){
    var t = e.target, btn = t.closest ? t.closest(".br-menu-btn") : null;
    if(btn){
      var m = btn.closest(".br-menu"), was = m.classList.contains("open");
      closeMenus();
      if(!was) m.classList.add("open");
      return;
    }
    if(t.closest && t.closest(".br-menu-list")){ setTimeout(closeMenus, 0); return; }
    closeMenus();
  });
  document.addEventListener("keydown", function(e){ if(e.key === "Escape") closeMenus(); });
  var TABLE_LIMIT = {};
  var EXPORT_PANELS = ["register","filemovement","letters","custom","judicial","dutyroster","judgesmanage","custommanage"];
  var LIBS = {
    pdflib:["/bill-register-system/lib/pdf-lib.min.js","https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js","https://unpkg.com/pdf-lib@1.17.1/dist/pdf-lib.min.js"],
    jszip:["/bill-register-system/lib/jszip.min.js","https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js","https://unpkg.com/jszip@3.10.1/dist/jszip.min.js"],
    docx:["/bill-register-system/lib/docx.iife.js","https://cdn.jsdelivr.net/npm/docx@9.7.1/dist/index.iife.js","https://unpkg.com/docx@9.7.1/dist/index.iife.js"]
  };
  var LIBCHK = {pdflib:function(){ return window.PDFLib; }, jszip:function(){ return window.JSZip; }, docx:function(){ return window.docx; }};
  function addScript(src){
    return new Promise(function(res, rej){
      var s = document.createElement("script"); s.src = src; s.onload = res;
      s.onerror = function(){ rej(new Error("could not load " + src)); }; document.head.appendChild(s);
    });
  }
  async function loadLib(name){
    if(LIBCHK[name]()) return;
    var err = null;
    for(var i = 0; i < LIBS[name].length; i++){
      try{ await addScript(LIBS[name][i]); if(LIBCHK[name]()) return; }catch(e){ err = e; }
    }
    throw err || new Error("library " + name);
  }
  function clean(s){ return String(s == null ? "" : s).replace(/\s+/g, " ").trim(); }
  function visible(el){ return !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length); }
  function pad2(n){ return (n < 10 ? "0" : "") + n; }
  var MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  function todayText(){ var d = new Date(); return d.getDate() + " " + MONTHS[d.getMonth()] + " " + d.getFullYear(); }
  function todayIso(){ var d = new Date(); return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate()); }
  function download(blob, name){
    var a = document.createElement("a"), url = URL.createObjectURL(blob);
    a.href = url; a.download = name; document.body.appendChild(a); a.click();
    setTimeout(function(){ URL.revokeObjectURL(url); a.remove(); }, 1500);
  }

  /* read the tables that are on screen in one tab */
  function extractTables(panel){
    var out = [], multi = panel.querySelectorAll("table").length > 1;
    panel.querySelectorAll("table").forEach(function(tb){
      if(tb.closest(".entry-form,.search-bar,#payslipOverlay") || !visible(tb)) return;
      var hr = tb.tHead && tb.tHead.rows.length ? tb.tHead.rows[tb.tHead.rows.length - 1] : null;
      if(!hr) return;
      var ths = Array.prototype.slice.call(hr.cells), keep = [], heads = [];
      ths.forEach(function(th, i){
        var txt = clean(th.textContent);
        if(th.classList.contains("row-actions-th") || /^(actions?|عمل)$/i.test(txt)) return;
        if(th.offsetWidth === 0 && th.getClientRects().length === 0) return;
        keep.push(i); heads.push(txt);
      });
      if(!heads.length) return;
      function expand(tr){
        var cells = [], merged = false;
        Array.prototype.forEach.call(tr.cells, function(td){
          var cl = td.cloneNode(true);
          Array.prototype.forEach.call(cl.querySelectorAll("button,script,style"), function(x){ x.remove(); });
          var v = (td.querySelector("input,select,textarea") && !clean(cl.textContent)) ? clean(td.querySelector("input,select,textarea").value) : clean(cl.textContent);
          cells.push({v:v, num:td.classList.contains("num"), skip:td.classList.contains("row-actions") || (td.offsetWidth === 0 && td.getClientRects().length === 0)});
          for(var k = 1; k < (td.colSpan || 1); k++) cells.push({v:"", num:false, skip:false});
          if((td.colSpan || 1) > 1) merged = true;
        });
        return {cells:cells, merged:merged};
      }
      function pick(r){
        if(r.merged && r.cells.length >= ths.length){
          var nonEmpty = r.cells.filter(function(c){ return c.v; }).length;
          if(nonEmpty === 1) return {banner:r.cells.filter(function(c){ return c.v; })[0].v};
        }
        return {cells:keep.map(function(i){ return r.cells[i] ? r.cells[i].v : ""; }), nums:keep.map(function(i){ return !!(r.cells[i] && r.cells[i].num); })};
      }
      var rows = [], foot = [];
      Array.prototype.forEach.call(tb.tBodies, function(tbody){
        Array.prototype.forEach.call(tbody.rows, function(tr){
          if(tr.classList.contains("empty-row") || (tr.offsetWidth === 0 && tr.getClientRects().length === 0)) return;
          var p = pick(expand(tr));
          if(p.banner){ rows.push(p); return; }
          if(!p.cells.some(function(c){ return c; })) return;
          rows.push(p);
        });
      });
      if(tb.tFoot) Array.prototype.forEach.call(tb.tFoot.rows, function(tr){
        var p = pick(expand(tr)); if(p.cells && p.cells.some(function(c){ return c; })) foot.push(p);
      });
      if(!rows.length && !foot.length) return;
      /* leave out columns that are empty on every row (e.g. unused custom fields) so the file stays readable */
      var live = heads.map(function(h, i){
        return i === 0 || rows.concat(foot).some(function(r){ return r.cells && r.cells[i] !== "" && r.cells[i] != null; });
      });
      if(live.some(function(x){ return !x; })){
        heads = heads.filter(function(_, i){ return live[i]; });
        keep = keep.filter(function(_, i){ return live[i]; });
        rows.forEach(function(r){ if(r.cells){ r.cells = r.cells.filter(function(_, i){ return live[i]; }); r.nums = r.nums.filter(function(_, i){ return live[i]; }); } });
        foot.forEach(function(r){ r.cells = r.cells.filter(function(_, i){ return live[i]; }); r.nums = r.nums.filter(function(_, i){ return live[i]; }); });
      }
      var num = heads.map(function(h, i){
        var th = ths[keep[i]], hits = 0, tot = 0;
        if(th.classList.contains("num")) return true;
        rows.forEach(function(r){ if(r.cells){ tot++; if(r.nums[i]) hits++; } });
        return tot > 0 && hits * 2 > tot;
      });
      var card = tb.closest(".card"), h3 = card ? card.querySelector("h3") : null, title = "";
      if(multi) title = h3 ? clean(h3.textContent) : "";
      out.push({title:title, heads:heads, rows:rows, foot:foot, num:num});
    });
    return TABLE_LIMIT[panel.id.replace("panel-", "")] ? out.slice(0, TABLE_LIMIT[panel.id.replace("panel-", "")]) : out;
  }

  /* a totals row such as "TOTAL (14 bills)" in column 1 with empty cells after it: let the label run across those empty cells */
  function footSpan(cells){
    var k = 1; while(k < cells.length && !cells[k]) k++;
    return (cells[0] && k > 1 && k < cells.length) ? k : 1;
  }
  function signoff(){
    var l = {name:"", designation:""}, r = {name:"", designation:""};
    try{
      var v = JSON.parse(localStorage.getItem("campoffice_preparedBy_v1") || "null");
      if(v && (v.name || v.designation)) l = {name:v.name || "", designation:v.designation || ""};
    }catch(e){}
    if(!l.name) l = {name:(STATE.preparedByName || ""), designation:(STATE.preparedByDesignation || "")};
    return {left:l, right:r};
  }
  function stampOnNow(){ try{ return localStorage.getItem("campoffice_includeStamp_v1") === "1"; }catch(e){ return false; } }
  function paperSize(){
    var v = "A4";
    if(v === "Legal") return {w:612, h:1008, tw:12240, th:20160, code:5};
    if(v === "Foolscap") return {w:612, h:936, tw:12240, th:18720, code:14};
    return {w:595.28, h:841.89, tw:11906, th:16838, code:9};
  }
  function panelTitle(panel){
    var h = panel.querySelector(".panel-head h2");
    return (h && clean(h.textContent)) || "Record Section";
  }
  function meta(key, panel, tables){
    var office = "";
    var title = panelTitle(panel), sub = [];
    var chip = panel.querySelector(".results-strip .count-chip, .panel-head .count-chip");
    if(chip && clean(chip.textContent)) sub.push(clean(chip.textContent));
    sub.push("Generated " + todayText());
    var maxCols = 0; tables.forEach(function(t){ maxCols = Math.max(maxCols, t.heads.length); });
    var base = "RecordSection_" + (title.replace(/[^A-Za-z0-9]+/g, "") || key) + "_" + todayIso();
    return {office:office, title:title, sub:sub.join("   |   "), landscape:maxCols > 6, base:base, sign:signoff(), stamp:stampOnNow(), paper:paperSize()};
  }

  /* ---- PDF (vector, pdf-lib) ---- */
  function san(s){
    return String(s == null ? "" : s).replace(/[‘’‛]/g, "'").replace(/[“”]/g, '"').replace(/[–—−]/g, "-")
      .replace(/ /g, " ").replace(/[​-‏‪-‮]/g, "").replace(/[^\x20-\x7E¡-ÿ]/g, "?");
  }
  async function logoBytes(){
    if(window.CampLogos && !CampLogos.show("govt")) return null;   /* switched off in the Logos control */
    try{
      var img = $("brandLogo"), src = (img && img.getAttribute("src")) || "logo-punjab-govt.png";
      var r = await fetch(src); if(!r.ok) return null;
      return new Uint8Array(await r.arrayBuffer());
    }catch(e){ return null; }
  }
  /* the emblem plus its pixel size (read from the PNG header) — used by the Word and Excel letterheads */
  function pngInfo(b){
    if(!b || b.length < 24 || b[1] !== 80 || b[2] !== 78 || b[3] !== 71) return null;
    var dv = new DataView(b.buffer, b.byteOffset, b.byteLength), w = dv.getUint32(16), h = dv.getUint32(20);
    return (w && h) ? {bytes:b, w:w, h:h} : null;
  }
  async function logoInfo(){ return pngInfo(await logoBytes()); }
  /* Public Prosecution Punjab logo — printed (right side of the heading) only when switched on in the Logos control */
  async function prosBytes(){
    if(!(window.CampLogos && CampLogos.show("pros"))) return null;
    try{ var r = await fetch("logo-public-prosecution.png"); if(!r.ok) return null; return new Uint8Array(await r.arrayBuffer()); }catch(e){ return null; }
  }
  async function prosInfo(){ return pngInfo(await prosBytes()); }
  async function buildPdf(tables, m){
    var L = PDFLib, doc = await L.PDFDocument.create();
    var F = await doc.embedFont(L.StandardFonts.Helvetica), FB = await doc.embedFont(L.StandardFonts.HelveticaBold);
    var PW = m.landscape ? m.paper.h : m.paper.w, PH = m.landscape ? m.paper.w : m.paper.h, M = 28, FOOT = 34;
    var navy = L.rgb(.106,.247,.388), goldL = L.rgb(.722,.525,.043), grey = L.rgb(.42,.42,.42), line = L.rgb(.78,.78,.78),
        black = L.rgb(.1,.1,.1), white = L.rgb(1,1,1), zebra = L.rgb(.955,.968,.98), tot = L.rgb(.89,.925,.955), gold = L.rgb(.97,.92,.77);
    var page = null, y = 0, usable = PW - 2 * M;
    function newPage(){ page = doc.addPage([PW, PH]); y = PH - M; }
    function wrap(text, font, size, width){
      var out = [];
      san(text).split("\n").forEach(function(p){
        var words = p.split(/\s+/).filter(Boolean), cur = "";
        if(!words.length){ out.push(""); return; }
        words.forEach(function(w){
          while(font.widthOfTextAtSize(w, size) > width){
            var k = w.length; while(k > 1 && font.widthOfTextAtSize(w.slice(0, k), size) > width) k--;
            if(cur){ out.push(cur); cur = ""; }
            out.push(w.slice(0, k)); w = w.slice(k);
          }
          var t = cur ? cur + " " + w : w;
          if(font.widthOfTextAtSize(t, size) <= width) cur = t; else { out.push(cur); cur = w; }
        });
        if(cur) out.push(cur);
      });
      return out;
    }
    function heading(t, i){
      var h = san(t.heads[i]);
      if((h.match(/\?/g) || []).length * 2 > h.length) return "Col " + (i + 1);   /* Urdu headings cannot be drawn in the standard PDF font */
      return h;
    }
    function layout(t){
      var n = t.heads.length, fs = t.dense ? 7 : (n > 12 ? 8 : (n > 8 ? 8.5 : 9)), w, FLOOR = t.dense ? 4.4 : 6.2;
      for(;;){
        var nat = [], mn = [], padd = t.dense ? 3.5 : 8;
        for(var i = 0; i < n; i++){
          var hd = heading(t, i), maxw = FB.widthOfTextAtSize(hd, fs) + padd, lw = 0;
          hd.split(/\s+/).forEach(function(x){ lw = Math.max(lw, FB.widthOfTextAtSize(x, fs)); });
          var minw = (t.dense ? Math.min(lw, 29) : lw) + padd;
          var rowsAll = t.rows.concat(t.foot);
          for(var r = 0; r < rowsAll.length; r++){
            var c = rowsAll[r].cells; if(!c) continue;
            var tx = san(c[i]); var ww = F.widthOfTextAtSize(tx, fs) + padd;
            maxw = Math.max(maxw, Math.min(ww, 230));
            tx.split(/\s+/).forEach(function(x){ minw = Math.max(minw, Math.min(F.widthOfTextAtSize(x, fs) + padd, 90)); });
            if(t.num[i]) minw = Math.max(minw, Math.min(ww, 120));
          }
          nat.push(Math.max(maxw, minw)); mn.push(Math.min(minw, Math.max(maxw, minw)));
        }
        var sumN = nat.reduce(function(a, b){ return a + b; }, 0);
        if(sumN <= usable){ var k = usable / sumN; w = nat.map(function(x){ return x * k; }); break; }
        var flex = nat.map(function(x, i){ return x - mn[i]; }), sumF = flex.reduce(function(a, b){ return a + b; }, 0), need = sumN - usable;
        if(sumF >= need){ w = nat.map(function(x, i){ return x - flex[i] * need / sumF; }); break; }
        if(fs > FLOOR){ fs -= .4; continue; }
        var k2 = usable / mn.reduce(function(a, b){ return a + b; }, 0); w = mn.map(function(x){ return x * k2; }); break;
      }
      return {w:w, fs:fs};
    }
    function drawHeader(t, lay){
      var hl = 1, x = M, hs = t.heads.map(function(_, i){ var ls = wrap(heading(t, i), FB, lay.fs, lay.w[i] - (t.dense ? 3 : 6)).slice(0, t.dense ? 6 : 2); hl = Math.max(hl, ls.length); return ls; });
      var h = hl * (lay.fs + 1.5) + 9;
      if(y - h - 30 < M + FOOT){ newPage(); }
      page.drawRectangle({x:M, y:y - h, width:usable, height:h, color:navy});
      page.drawRectangle({x:M, y:y - h, width:usable, height:1.6, color:goldL});
      hs.forEach(function(lines, i){
        var yy = y - 5 - lay.fs;
        lines.forEach(function(l){
          var tw = FB.widthOfTextAtSize(l, lay.fs), xx = t.num[i] ? x + lay.w[i] - 3 - tw : x + 3;
          page.drawText(l, {x:xx, y:yy + 1, size:lay.fs, font:FB, color:white}); yy -= lay.fs + 1.5;
        });
        x += lay.w[i];
      });
      y -= h;
    }
    function drawRow(t, lay, r, idx, kind){
      var fs = lay.fs, font = kind === "foot" ? FB : F, lh = fs + 2.2, cols = [], hh = 0, n = t.heads.length;
      var sp = kind === "foot" ? footSpan(r.cells) : 1, spW = 0;
      for(var q = 0; q < sp; q++) spW += lay.w[q];
      for(var i = 0; i < n; i++){
        if(i > 0 && i < sp){ cols.push([]); continue; }
        var ls = wrap(r.cells[i], font, fs, (i === 0 ? spW : lay.w[i]) - (t.dense ? 3 : 6)); cols.push(ls); hh = Math.max(hh, ls.length * lh);
      }
      var rh = t.dense ? Math.max(hh + 4, fs + 6) : Math.max(hh + 6, fs + 10);
      if(y - rh < M + FOOT){ newPage(); drawHeader(t, lay); }
      if(kind === "foot") page.drawRectangle({x:M, y:y - rh, width:usable, height:rh, color:tot});
      else if(idx % 2) page.drawRectangle({x:M, y:y - rh, width:usable, height:rh, color:zebra});
      var x = M;
      for(var j = 0; j < n; j++){
        var yy = y - 3;
        cols[j].forEach(function(l){
          yy -= fs; var tw = font.widthOfTextAtSize(l, fs), xx = t.num[j] ? x + lay.w[j] - 3 - tw : x + 3;
          page.drawText(l, {x:xx, y:yy + 1, size:fs, font:font, color:black}); yy -= 2.2;
        });
        x += lay.w[j];
      }
      page.drawLine({start:{x:M, y:y - rh}, end:{x:M + usable, y:y - rh}, thickness:.4, color:line});
      y -= rh;
    }
    function banner(text, fill, fg, size){
      var lines = wrap(text, FB, size, usable - 12), h = lines.length * (size + 3) + 8;
      if(y - h - 60 < M + FOOT) newPage();
      page.drawRectangle({x:M, y:y - h, width:usable, height:h, color:fill});
      var yy = y - 4 - size;
      lines.forEach(function(l){ page.drawText(l, {x:M + 6, y:yy, size:size, font:FB, color:fg}); yy -= size + 3; });
      y -= h;
    }
    newPage();
    var top = y, emb = null, eb = await logoBytes();
    if(eb){ try{ emb = await doc.embedPng(eb); }catch(e){ emb = null; } }
    if(emb){ var eh = 40, ew = eh * emb.width / emb.height; page.drawImage(emb, {x:M, y:y - eh, width:ew, height:eh}); }
    var pbts = await prosBytes(), pemb = null;
    if(pbts){ try{ pemb = await doc.embedPng(pbts); }catch(e){ pemb = null; } }
    if(pemb){ var ph2 = 40, pw2 = ph2 * pemb.width / pemb.height; page.drawImage(pemb, {x:PW - M - pw2, y:y - ph2, width:pw2, height:ph2}); }
    var h1 = san(m.office), h2 = "";
    page.drawText(h1, {x:(PW - FB.widthOfTextAtSize(h1, 15)) / 2, y:y - 15, size:15, font:FB, color:navy});
    if(h1 !== "")
      page.drawText(san(h2), {x:(PW - F.widthOfTextAtSize(san(h2), 8.5)) / 2, y:y - 29, size:8.5, font:F, color:grey});
    y = top - 46;
    page.drawLine({start:{x:M, y:y}, end:{x:PW - M, y:y}, thickness:2.2, color:goldL});
    page.drawLine({start:{x:M, y:y - 3.6}, end:{x:PW - M, y:y - 3.6}, thickness:.7, color:navy});
    y -= 18;
    var tt = san(m.title);
    page.drawText(tt, {x:(PW - FB.widthOfTextAtSize(tt, 14)) / 2, y:y - 14, size:14, font:FB, color:navy}); y -= 21;
    var st = san(m.sub);
    page.drawText(st, {x:(PW - F.widthOfTextAtSize(st, 8.5)) / 2, y:y - 9, size:8.5, font:F, color:grey}); y -= 20;
    tables.forEach(function(t, ti){
      var lay = layout(t);
      if(t.title) banner(t.title, gold, black, 10.5);
      if(y - 60 < M + FOOT) newPage();
      drawHeader(t, lay);
      var k = 0;
      t.rows.forEach(function(r){
        if(r.banner){
          var hh = lay.fs + 9; if(y - hh < M + FOOT){ newPage(); drawHeader(t, lay); }
          page.drawRectangle({x:M, y:y - hh, width:usable, height:hh, color:gold});
          page.drawText(san(r.banner), {x:M + 4, y:y - hh + 5, size:lay.fs, font:FB, color:black}); y -= hh; k = 0; return;
        }
        drawRow(t, lay, r, k++, "row");
      });
      t.foot.forEach(function(r){ drawRow(t, lay, r, 0, "foot"); });
      y -= 14;
    });
    /* prepared / approved block (optional: only when the stamp box is ticked) */
    if(m.stamp){
    if(y - 70 < M + FOOT) newPage();
    y -= 26;
    var colW = (usable - 40) / 2;
    [["Prepared by", m.sign.left, M], ["Approved by", m.sign.right, M + colW + 40]].forEach(function(b){
      var x0 = b[2];
      page.drawLine({start:{x:x0, y:y}, end:{x:x0 + colW, y:y}, thickness:.7, color:black});
      page.drawText(san(b[0]), {x:x0, y:y - 11, size:8, font:FB, color:navy});
      var nm = san(b[1].name || ""), ds = wrap(b[1].designation || "", F, 8, colW);
      var yy = y - 22;
      if(nm){ page.drawText(nm, {x:x0, y:yy, size:9, font:FB, color:black}); yy -= 11; }
      ds.slice(0, 3).forEach(function(l){ page.drawText(l, {x:x0, y:yy, size:8, font:F, color:grey}); yy -= 10; });
    });
    }
    var pages = doc.getPages();
    pages.forEach(function(p, i){
      p.drawLine({start:{x:M, y:26}, end:{x:PW - M, y:26}, thickness:.6, color:line});
      p.drawText("", {x:M, y:15, size:7, font:F, color:grey});
      var mid = "Generated " + todayText();
      p.drawText(mid, {x:(PW - F.widthOfTextAtSize(mid, 7)) / 2, y:15, size:7, font:F, color:grey});
      var s = "Page " + (i + 1) + " of " + pages.length;
      p.drawText(s, {x:PW - M - FB.widthOfTextAtSize(s, 7.4), y:15, size:7.4, font:FB, color:navy});
    });
    doc.setTitle(san(m.title)); doc.setCreator("");
    return await doc.save();
  }

  /* ---- Word (.docx) ---- */
  async function buildDocx(tables, m){
    var D = window.docx, FONT = "Calibri";
    var pw = m.landscape ? m.paper.th : m.paper.tw, ph = m.landscape ? m.paper.tw : m.paper.th, MG = 567, usable = pw - 2 * MG;
    function run(t, o){ o = o || {}; return new D.TextRun({text:String(t == null ? "" : t), font:FONT, size:o.size || 17, bold:!!o.bold, color:o.color}); }
    function para(parts, o){ o = o || {}; return new D.Paragraph({children:parts, alignment:o.align, spacing:{before:o.before || 0, after:o.after || 0}, keepNext:!!o.keepNext}); }
    var bd = {style:D.BorderStyle.SINGLE, size:4, color:"BBBBBB"}, none = {style:D.BorderStyle.NONE, size:0, color:"FFFFFF"};
    var children = [], lg = await logoInfo(), pg = await prosInfo();
    var headParas = [
      para([run(m.office, {size:30, bold:true, color:"1B3F63"})], {align:D.AlignmentType.CENTER, after:20}),
      para([run("", {size:16, color:"666666"})], {align:D.AlignmentType.CENTER, after:80})
    ];
    if(lg || pg){
      /* emblem on the LEFT, heading centred — same letterhead as the Weekly Court Roster */
      var lh = 54, lw = lg ? Math.round(lh * lg.w / lg.h) : 0, pw3 = pg ? Math.round(lh * pg.w / pg.h) : 0, sideW = 1500, midW = usable - 2 * sideW, nob = {top:none, bottom:none, left:none, right:none};
      children.push(new D.Table({
        rows:[new D.TableRow({cantSplit:true, children:[
          new D.TableCell({children:[new D.Paragraph({children:lg ? [new D.ImageRun({type:"png", data:lg.bytes, transformation:{width:lw, height:lh}})] : []})],
            width:{size:sideW, type:D.WidthType.DXA}, borders:nob, verticalAlign:D.VerticalAlign.CENTER}),
          new D.TableCell({children:headParas, width:{size:midW, type:D.WidthType.DXA}, borders:nob, verticalAlign:D.VerticalAlign.CENTER}),
          new D.TableCell({children:[new D.Paragraph({alignment:D.AlignmentType.RIGHT, children:pg ? [new D.ImageRun({type:"png", data:pg.bytes, transformation:{width:pw3, height:lh}})] : []})], width:{size:sideW, type:D.WidthType.DXA}, borders:nob})
        ]})],
        width:{size:usable, type:D.WidthType.DXA}, columnWidths:[sideW, midW, sideW], layout:D.TableLayoutType.FIXED,
        borders:{top:none, bottom:none, left:none, right:none, insideHorizontal:none, insideVertical:none}}));
    } else { headParas.forEach(function(x){ children.push(x); }); }
    children.push(new D.Paragraph({children:[run(m.title, {size:28, bold:true, color:"1B3F63"})], alignment:D.AlignmentType.CENTER, spacing:{before:80, after:40},
      border:{top:{style:D.BorderStyle.SINGLE, size:14, color:"B8860B", space:6}}}));
    children.push(para([run(m.sub, {size:16, color:"666666"})], {align:D.AlignmentType.CENTER, after:160}));
    tables.forEach(function(t){
      var n = t.heads.length, wt = t.heads.map(function(h, i){
        var mx = t.dense ? Math.min(Math.max.apply(null, String(h).split(/\s+/).map(function(x){ return x.length; })), 11) : Math.max(h.length * 1.05, 5);
        t.rows.concat(t.foot).forEach(function(r){ if(r.cells) mx = Math.max(mx, Math.min(String(r.cells[i] || "").length, t.dense ? (i < 3 ? 16 : 11) : 36)); });
        return t.dense ? Math.max(mx + 1.5, 8.5) : Math.max(mx, t.num[i] ? 8 : 5);
      }), sum = wt.reduce(function(a, b){ return a + b; }, 0);
      var W = wt.map(function(x){ return Math.floor(usable * x / sum); });
      var fs = t.dense ? 9 : n > 12 ? 13 : n > 9 ? 14 : n > 6 ? 16 : 18, cm = t.dense ? 25 : 70;
      function cell(text, i, o){
        o = o || {};
        return new D.TableCell({children:[para([run(text, {size:fs, bold:!!o.bold, color:o.color})], {align:(t.num[i] && !o.left) ? D.AlignmentType.RIGHT : undefined})],
          width:{size:W[i], type:D.WidthType.DXA}, verticalAlign:D.VerticalAlign.CENTER,
          shading:o.fill ? {type:D.ShadingType.CLEAR, color:"auto", fill:o.fill} : undefined, margins:{top:t.dense ? 25 : 45, bottom:t.dense ? 25 : 45, left:cm, right:cm}});
      }
      if(t.title) children.push(new D.Paragraph({children:[run(t.title, {size:20, bold:true, color:"1B3F63"})], shading:{type:D.ShadingType.CLEAR, color:"auto", fill:"F8EBC4"}, spacing:{before:160, after:60}, keepNext:true}));
      var trs = [new D.TableRow({tableHeader:true, cantSplit:true, children:t.heads.map(function(h, i){ return cell(h, i, {bold:true, color:"FFFFFF", fill:"1B3F63"}); })})], k = 0;
      t.rows.forEach(function(r){
        if(r.banner){
          trs.push(new D.TableRow({cantSplit:true, children:[new D.TableCell({columnSpan:n, children:[para([run(r.banner, {size:fs, bold:true})])],
            shading:{type:D.ShadingType.CLEAR, color:"auto", fill:"F8EBC4"}, margins:{top:45, bottom:45, left:70, right:70}})]}));
          k = 0; return;
        }
        var fill = (k++ % 2) ? "F3F6FA" : undefined;
        trs.push(new D.TableRow({cantSplit:true, children:r.cells.map(function(c, i){ return cell(c, i, {fill:fill}); })}));
      });
      t.foot.forEach(function(r){
        var sp = footSpan(r.cells), kids = [];
        if(sp > 1){
          var wsum = 0; for(var q = 0; q < sp; q++) wsum += W[q];
          kids.push(new D.TableCell({columnSpan:sp, children:[para([run(r.cells[0], {size:fs, bold:true})])], width:{size:wsum, type:D.WidthType.DXA}, verticalAlign:D.VerticalAlign.CENTER,
            shading:{type:D.ShadingType.CLEAR, color:"auto", fill:"E3EBF3"}, margins:{top:45, bottom:45, left:70, right:70}}));
        }
        r.cells.forEach(function(c, i){ if(sp > 1 && i < sp) return; kids.push(cell(c, i, {bold:true, fill:"E3EBF3"})); });
        trs.push(new D.TableRow({cantSplit:true, children:kids}));
      });
      children.push(new D.Table({rows:trs, width:{size:usable, type:D.WidthType.DXA}, columnWidths:W, layout:D.TableLayoutType.FIXED,
        borders:{top:bd, bottom:bd, left:bd, right:bd, insideHorizontal:bd, insideVertical:bd}}));
    });
    if(m.stamp) children.push(para([run("")], {before:300}));
    function sg(label, p, w){
      var ps = [para([run("______________________________", {size:16})]), para([run(label, {size:15, bold:true, color:"1B3F63"})])];
      if(p.name) ps.push(para([run(p.name, {size:17, bold:true})]));
      if(p.designation) ps.push(para([run(p.designation, {size:15, color:"666666"})]));
      return new D.TableCell({children:ps, width:{size:w, type:D.WidthType.DXA}, borders:{top:none, bottom:none, left:none, right:none}});
    }
    var half = Math.floor(usable / 2);
    if(m.stamp) children.push(new D.Table({rows:[new D.TableRow({cantSplit:true, children:[sg("Prepared by", m.sign.left, half), sg("Approved by", m.sign.right, half)]})],
      width:{size:usable, type:D.WidthType.DXA}, columnWidths:[half, half], layout:D.TableLayoutType.FIXED,
      borders:{top:none, bottom:none, left:none, right:none, insideHorizontal:none, insideVertical:none}}));
    var doc = new D.Document({
      creator:"", title:m.title,
      sections:[{
        properties:{page:{size:{width:m.paper.tw, height:m.paper.th, orientation:m.landscape ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT}, margin:{top:567, bottom:680, left:MG, right:MG, footer:300}}},
        footers:{default:new D.Footer({children:[new D.Paragraph({alignment:D.AlignmentType.CENTER, children:[
          new D.TextRun({children:["Page ", D.PageNumber.CURRENT, " of ", D.PageNumber.TOTAL_PAGES], font:FONT, size:15, color:"777777"})]})]})},
        children:children
      }]
    });
    return await D.Packer.toBlob(doc);
  }

  /* ---- Excel (.xlsx written straight with JSZip) ---- */
  function xe(s){ return String(s == null ? "" : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
  function colName(i){ var s = ""; i++; while(i > 0){ var r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); } return s; }
  function numOf(s){
    var x = String(s).replace(/[,\s]/g, "").replace(/^\((.*)\)$/, "-$1");
    if(!/^-?\d+(\.\d+)?$/.test(x)) return null;
    if(/^-?0\d/.test(x)) return null;
    return parseFloat(x);
  }
  async function buildXlsx(tables, m){
    /* style ids: 0 def, 1 title, 2 sub, 3 head, 4 head-right, 5 text, 6 text zebra, 7 int, 8 int zebra, 9 dec, 10 dec zebra,
                  11 total text, 12 total int, 13 total dec, 14 banner, 15 sign label, 16 sign text, 17 office */
    var styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="8"><font><sz val="10"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="15"/><color rgb="FF1B3F63"/><name val="Calibri"/></font>' +
      '<font><sz val="10"/><color rgb="FF666666"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="10"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="10.5"/><color rgb="FF1B3F63"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="12"/><color rgb="FF1B3F63"/><name val="Calibri"/></font>' +
      '<font><sz val="9"/><color rgb="FF666666"/><name val="Calibri"/></font></fonts>' +
      '<fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FF1B3F63"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFF3F6FA"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFE3EBF3"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFF8EBC4"/><bgColor indexed="64"/></patternFill></fill></fills>' +
      '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
      '<border><left style="thin"><color rgb="FFC8CED6"/></left><right style="thin"><color rgb="FFC8CED6"/></right><top style="thin"><color rgb="FFC8CED6"/></top><bottom style="thin"><color rgb="FFC8CED6"/></bottom><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="18">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment horizontal="left" vertical="center"/></xf>' +
      '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment horizontal="left" vertical="center"/></xf>' +
      '<xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment vertical="top" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"><alignment vertical="top" wrapText="1"/></xf>' +
      '<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"><alignment horizontal="right" vertical="top"/></xf>' +
      '<xf numFmtId="3" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"><alignment horizontal="right" vertical="top"/></xf>' +
      '<xf numFmtId="4" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"><alignment horizontal="right" vertical="top"/></xf>' +
      '<xf numFmtId="4" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"><alignment horizontal="right" vertical="top"/></xf>' +
      '<xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment vertical="top" wrapText="1"/></xf>' +
      '<xf numFmtId="3" fontId="4" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right" vertical="top"/></xf>' +
      '<xf numFmtId="4" fontId="4" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right" vertical="top"/></xf>' +
      '<xf numFmtId="0" fontId="5" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="left" vertical="center"/></xf>' +
      '<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment horizontal="left" vertical="center"/></xf>' +
      '<xf numFmtId="0" fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment horizontal="left" vertical="center"/></xf>' +
      '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    var lg = await logoInfo(), pg = await prosInfo();
    if(lg){   /* leave room on the left of the three heading rows for the emblem */
      ['<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment horizontal="left" vertical="center"/>',
       '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment horizontal="left" vertical="center"/>',
       '<xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment horizontal="left" vertical="center"/>'].forEach(function(x){
        styles = styles.replace(x, x.replace('<alignment horizontal="left"', '<alignment horizontal="left" indent="8"'));
      });
    }
    function c(ref, text, s){ return '<c r="' + ref + '" s="' + s + '" t="inlineStr"><is><t xml:space="preserve">' + xe(text) + '</t></is></c>'; }
    function cn(ref, v, s){ return '<c r="' + ref + '" s="' + s + '"><v>' + v + '</v></c>'; }
    function lines(text, w){ var n = 0; String(text == null ? "" : text).split("\n").forEach(function(p){ n += Math.max(1, Math.ceil(p.length * 1.1 / Math.max(4, w - 1))); }); return n; }
    var used = {};
    function sheetName(t, i){
      var base = (t.title || m.title).replace(/[\[\]:*?\/\\]/g, " ").trim().slice(0, 28) || ("Sheet" + (i + 1)), nm = base, k = 2;
      while(used[nm.toLowerCase()]) nm = base.slice(0, 26) + " " + (k++);
      used[nm.toLowerCase()] = 1; return nm;
    }
    var sheets = [];
    tables.forEach(function(t, ti){
      var n = t.heads.length, last = colName(n - 1), rows = "", merges = [], r = 1;
      var dec = t.heads.map(function(h, i){
        if(!t.num[i]) return false;
        return t.rows.concat(t.foot).some(function(x){ if(!x.cells) return false; var v = numOf(x.cells[i]); return v !== null && Math.abs(v - Math.round(v)) > 1e-9; });
      });
      var widths = t.heads.map(function(h, i){
        var mx = Math.min(h.length + 2, 24);
        t.rows.concat(t.foot).forEach(function(x){ if(x.cells) mx = Math.max(mx, Math.min(String(x.cells[i] || "").length + 2, 46)); });
        return Math.max(8, mx);
      });
      rows += '<row r="1" ht="22" customHeight="1">' + c("A1", m.office, 17) + "</row>"; merges.push("A1:" + last + "1");
      rows += '<row r="2" ht="26" customHeight="1">' + c("A2", m.title + (t.title ? " — " + t.title : ""), 1) + "</row>"; merges.push("A2:" + last + "2");
      rows += '<row r="3" ht="16" customHeight="1">' + c("A3", m.sub, 2) + "</row>"; merges.push("A3:" + last + "3");
      r = 4;
      rows += '<row r="4" ht="30" customHeight="1">' + t.heads.map(function(h, i){ return c(colName(i) + "4", h, t.num[i] ? 4 : 3); }).join("") + "</row>";
      var k = 0, first = 5;
      function dataRow(x, foot){
        r++;
        var z = (!foot && (k++ % 2)) ? 1 : 0, mx = 1;
        var cells = x.cells.map(function(v, i){
          mx = Math.max(mx, lines(v, widths[i]));
          var nv = t.num[i] ? numOf(v) : null;
          if(nv !== null){ var sid = foot ? (dec[i] ? 13 : 12) : (dec[i] ? 9 + z : 7 + z); return cn(colName(i) + r, nv, sid); }
          return c(colName(i) + r, v, foot ? 11 : (5 + z));
        }).join("");
        rows += '<row r="' + r + '" ht="' + Math.max(17, mx * 13 + 4) + '" customHeight="1">' + cells + "</row>";
      }
      t.rows.forEach(function(x){
        if(x.banner){
          r++; rows += '<row r="' + r + '" ht="20" customHeight="1">' + c("A" + r, x.banner, 14) + t.heads.slice(1).map(function(h, i){ return c(colName(i + 1) + r, "", 14); }).join("") + "</row>";
          merges.push("A" + r + ":" + last + r); k = 0; return;
        }
        dataRow(x, false);
      });
      var lastData = r;
      t.foot.forEach(function(x){
        var sp = footSpan(x.cells); dataRow(x, true);
        if(sp > 1) merges.push("A" + r + ":" + colName(sp - 1) + r);
      });
      r += 2;
      var half = Math.max(1, Math.floor(n / 2));
      /* two sign-off blocks must not share a row element twice, so build them cell by cell */
      var sr = r + 1, lp = m.sign.left, rp = m.sign.right, lc1 = Math.max(0, half - 1), rc0 = half, rc1 = n - 1;
      var labels = [["_________________________", "_________________________", 15], ["Prepared by", "Approved by", 15], [lp.name || "", rp.name || "", 15], [lp.designation || "", rp.designation || "", 16]];
      (m.stamp ? labels : []).forEach(function(L, q){
        var row = sr + q;
        rows += '<row r="' + row + '" ht="' + (q === 3 ? 26 : 17) + '" customHeight="1">' + c("A" + row, L[0], L[2]) + (n > 1 ? c(colName(rc0) + row, L[1], L[2]) : "") + "</row>";
        if(lc1 > 0) merges.push("A" + row + ":" + colName(lc1) + row);
        if(n > 1 && rc1 > rc0) merges.push(colName(rc0) + row + ":" + colName(rc1) + row);
      });
      var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView workbookViewId="0" showGridLines="0"' + (ti === 0 ? ' tabSelected="1"' : '') + '><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
        '<sheetFormatPr defaultRowHeight="15"/><cols>' + widths.map(function(w, i){ return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>'; }).join("") + "</cols>" +
        "<sheetData>" + rows + '</sheetData><autoFilter ref="A4:' + last + (lastData < 5 ? 5 : lastData) + '"/><mergeCells count="' + merges.length + '">' + merges.map(function(x){ return '<mergeCell ref="' + x + '"/>'; }).join("") + "</mergeCells>" +
        '<printOptions horizontalCentered="1"/><pageMargins left="0.4" right="0.4" top="0.5" bottom="0.6" header="0.25" footer="0.3"/>' +
        '<pageSetup paperSize="' + m.paper.code + '" orientation="' + (m.landscape ? "landscape" : "portrait") + '" fitToWidth="1" fitToHeight="0"/>' +
        '<headerFooter><oddFooter>&amp;R&amp;8Page &amp;P of &amp;N</oddFooter></headerFooter>' + ((lg || pg) ? '<drawing r:id="rId1"/>' : '') + '</worksheet>';
      sheets.push({lastW:widths[n - 1], nCols:n, name:sheetName(t, ti), xml:xml, filterRef:"A4:" + last + (lastData < 5 ? 5 : lastData)});
    });
    var wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
      sheets.map(function(s, i){ return '<sheet name="' + xe(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>'; }).join("") + "</sheets><definedNames>" +
      sheets.map(function(s, i){ return '<definedName name="_xlnm._FilterDatabase" localSheetId="' + i + '" hidden="1">\'' + xe(s.name).replace(/'/g, "''") + "'!" + s.filterRef.replace(/([A-Z]+)(\d+)/g, "$$$1$$$2") + "</definedName>"; }).join("") + "</definedNames></workbook>";
    var zip = new JSZip();
    zip.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + ((lg || pg) ? '<Default Extension="png" ContentType="image/png"/>' : '') + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      sheets.map(function(s, i){ return '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' + ((lg || pg) ? '<Override PartName="/xl/drawings/drawing' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' : ''); }).join("") + "</Types>");
    zip.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    zip.file("xl/workbook.xml", wb);
    zip.file("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map(function(s, i){ return '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>'; }).join("") +
      '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
    zip.file("xl/styles.xml", styles);
    if(lg || pg){
      var LH = 44, LW = lg ? Math.round(LH * lg.w / lg.h) : 0, PW2 = pg ? Math.round(LH * pg.w / pg.h) : 0;
      if(lg) zip.file("xl/media/logo.png", lg.bytes);
      if(pg) zip.file("xl/media/pros.png", pg.bytes);
      var RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
      function pic(id, name, rid, cx, cy, col, colOff){
        return '<xdr:oneCellAnchor><xdr:from><xdr:col>' + col + '</xdr:col><xdr:colOff>' + colOff + '</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>38100</xdr:rowOff></xdr:from><xdr:ext cx="' + cx + '" cy="' + cy + '"/>' +
          '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + id + '" name="' + name + '" descr="' + name + '"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
          '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
          '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>';
      }
      sheets.forEach(function(s, i){
        zip.file("xl/worksheets/_rels/sheet" + (i + 1) + ".xml.rels", RELS + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' + (i + 1) + '.xml"/></Relationships>');
        zip.file("xl/drawings/_rels/drawing" + (i + 1) + ".xml.rels", RELS + (lg ? '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/logo.png"/>' : '') +
          (pg ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/pros.png"/>' : '') + '</Relationships>');
        var colPx = Math.round((s.lastW || 10) * 7 + 5), lastCol = Math.max(0, (s.nCols || 1) - 1);
        zip.file("xl/drawings/drawing" + (i + 1) + ".xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
          (lg ? pic(2, "Government of the Punjab", "rId1", LW * 9525, LH * 9525, 0, 38100) : '') +
          (pg ? pic(3, "Public Prosecution Punjab", "rId2", PW2 * 9525, LH * 9525, lastCol, Math.max(0, (colPx - PW2 - 4) * 9525)) : '') + '</xdr:wsDr>');
      });
    }
    sheets.forEach(function(s, i){ zip.file("xl/worksheets/sheet" + (i + 1) + ".xml", s.xml); });
    return await zip.generateAsync({type:"blob", mimeType:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  }

  function buildCsv(tables){
    function q(v){ v = String(v == null ? "" : v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }
    var out = [];
    tables.forEach(function(t, i){
      if(i) out.push("");
      if(t.title) out.push(q(t.title));
      out.push(t.heads.map(q).join(","));
      t.rows.forEach(function(r){ out.push(r.banner ? q(r.banner) : r.cells.map(q).join(",")); });
      t.foot.forEach(function(r){ out.push(r.cells.map(q).join(",")); });
    });
    return new Blob(["﻿" + out.join("\r\n")], {type:"text/csv;charset=utf-8"});
  }

  async function runExport(key, kind, bar){
    var panel = $("panel-" + key);
    if(!panel) return;
    if(kind === "print"){ var pb = $("globalPrintBtn"); if(pb) pb.click(); else window.print(); return; }
    if(typeof canTab === "function" && !canTab("x")){ if(typeof rsDenied === "function") rsDenied(); return; }
    var tables = extractTables(panel);
    if(!tables.length){
      say("There are no records to export on this tab.", "اس ٹیب میں ایکسپورٹ کے لیے کوئی ریکارڈ نہیں ہے۔"); return;
    }
    var m = meta(key, panel, tables);
    bar.classList.add("busy");
    say("Preparing the file…", "فائل تیار ہو رہی ہے…");
    try{
      if(kind === "pdf"){ await loadLib("pdflib"); download(new Blob([await buildPdf(tables, m)], {type:"application/pdf"}), m.base + ".pdf"); }
      else if(kind === "word"){ await loadLib("docx"); download(await buildDocx(tables, m), m.base + ".docx"); }
      else if(kind === "excel"){ await loadLib("jszip"); download(await buildXlsx(tables, m), m.base + ".xlsx"); }
      else if(kind === "csv"){ download(buildCsv(tables), m.base + ".csv"); }
      say("File ready — check your Downloads folder.", "فائل تیار ہو گئی — ڈاؤن لوڈ فولڈر دیکھیں۔");
    }catch(err){
      console.error(err);
      say("Could not build the file: " + (err && err.message ? err.message : err), "فائل نہیں بن سکی: " + (err && err.message ? err.message : err));
    }finally{ bar.classList.remove("busy"); }
  }

  function injectBars(){
    EXPORT_PANELS.forEach(function(key){
      var p = $("panel-" + key), head = p && p.querySelector(".panel-head");
      if(!head || head.querySelector(".br-exportbar")) return;
      var wrap = document.createElement("div"); wrap.className = "br-head-actions";
      Array.prototype.slice.call(head.children, 1).forEach(function(ch){ wrap.appendChild(ch); });
      var bar = document.createElement("div"); bar.className = "br-exportbar br-menu"; bar.setAttribute("data-panel", key);
      bar.innerHTML =
        '<button type="button" class="br-menu-btn" aria-haspopup="true" data-perm="x">⬇ ' + bi("Print / Download", "پرنٹ / ڈاؤن لوڈ") + ' <span class="caret">▾</span></button>' +
        '<div class="br-menu-list" role="menu">' +
        '<button type="button" data-act="print"><span class="mi">🖨️</span><span>' + bi("Print", "پرنٹ") + '<span class="sub">' + bi("stamp is optional", "اسٹامپ اختیاری") + '</span></span></button>' +
        '<label class="br-stamp-row" style="display:flex;align-items:center;gap:8px;padding:8px 12px;font-size:12.5px;cursor:pointer;"><input type="checkbox" class="br-stamp-chk"> <span>' + bi("Add Prepared by stamp (files)", "فائلوں میں تیار کردہ اسٹامپ شامل کریں") + '</span></label>' +
        '<hr>' +
        '<button type="button" data-act="pdf"><span class="mi">📄</span><span>' + bi("PDF file", "پی ڈی ایف فائل") + '</span></button>' +
        '<button type="button" data-act="word"><span class="mi">📝</span><span>' + bi("Word file (.docx)", "ورڈ فائل") + '</span></button>' +
        '<button type="button" data-act="excel"><span class="mi">📊</span><span>' + bi("Excel file (.xlsx)", "ایکسل فائل") + '</span></button>' +
        '<button type="button" data-act="csv"><span class="mi">🧾</span><span>' + bi("CSV file", "سی ایس وی فائل") + '</span></button>' +
        '</div>';
      wrap.appendChild(bar); head.appendChild(wrap);
      var stChk = bar.querySelector(".br-stamp-chk"), stRow = bar.querySelector(".br-stamp-row");
      function stSync(){ stChk.checked = stampOnNow(); }
      stSync();
      bar.querySelector(".br-menu-btn").addEventListener("click", stSync);
      stRow.addEventListener("click", function(e){ e.stopPropagation(); });
      stChk.addEventListener("change", function(){
        try{ localStorage.setItem("campoffice_includeStamp_v1", stChk.checked ? "1" : "0"); }catch(er){}
        document.querySelectorAll(".br-stamp-chk").forEach(function(c){ c.checked = stChk.checked; });
      });
      bar.querySelector(".br-menu-list").addEventListener("click", function(e){
        var b = e.target.closest("button[data-act]"); if(!b) return;
        runExport(key, b.getAttribute("data-act"), bar);
      });
    });
  }
  injectBars();
  try{ rsMarkPermButtons(document); }catch(e){}
})();
