/* Извлечение таблиц учебных планов из PDF приказов (pdf.js уже загружен страницей).
   Колонки классов определяются по координатам заголовка (VII…XI или 10 11), поэтому
   пустые ячейки не сдвигают значения. Результат — массив таблиц с проверкой сумм. */
(function(){
  var ROMAN = { "V":5, "VI":6, "VII":7, "VIII":8, "IX":9, "X":10, "XI":11 };
  var NUM = /^\d+(?:[,.]\d+)?$/;
  function num(s){ return parseFloat(String(s).replace(",", ".")); }
  function lines(items){
    var rows = [];
    items.forEach(function(it){
      var s = it.str; if(!s || !s.trim()) return;
      var x = it.transform[4], y = it.transform[5];
      var row = rows.filter(function(r){ return Math.abs(r.y - y) <= 3; })[0];
      if(!row){ row = { y: y, items: [] }; rows.push(row); }
      row.items.push({ x: x, w: it.width || 0, s: s });
    });
    rows.sort(function(a,b){ return b.y - a.y; });
    rows.forEach(function(r){
      r.items.sort(function(a,b){ return a.x - b.x; });
      // склеиваем фрагменты, стоящие вплотную («1,» + «5», «V» + «II»), в один токен
      var merged = [];
      r.items.forEach(function(i){
        var last = merged[merged.length-1];
        if(last && i.x - (last.x + last.w) < 1.2 && !/\s$/.test(last.s) && !/^\s/.test(i.s)){ last.s += i.s; last.w = i.x + i.w - last.x; }
        else merged.push({ x: i.x, w: i.w, s: i.s });
      });
      r.items = merged;
      r.text = r.items.map(function(i){ return i.s; }).join(" ").replace(/\s+/g," ").trim();
    });
    return rows;
  }
  // заголовок таблицы: строка, где есть ≥2 номера классов (римских или 10/11)
  function headerCols(row){
    // строка заголовка состоит только из номеров классов (римских или 5…11), по возрастанию
    var cols = [], ok = true;
    row.items.forEach(function(i){
      i.s.trim().split(/\s+/).forEach(function(s){
        if(!s) return;
        if(ROMAN[s]) cols.push({ grade: ROMAN[s], x: i.x + i.w/2 });
        else if(/^(5|6|7|8|9|10|11)$/.test(s)) cols.push({ grade: parseInt(s,10), x: i.x + i.w/2 });
        else if(!/^(soat|umumiy|haftalik|jami|T\/r)$/i.test(s)) ok = false;
      });
    });
    if(!ok || cols.length < 2) return null;
    for(var k=1; k<cols.length; k++){ if(cols[k].grade <= cols[k-1].grade) return null; }
    return cols;
  }
  window.extractTables = async function(url, docId){
    var pdf = await pdfjsLib.getDocument(url).promise;
    var tables = [], cur = null, ctx = { annex: null, titleLines: [] };
    for(var p = 1; p <= pdf.numPages; p++){
      var page = await pdf.getPage(p);
      var tc = await page.getTextContent();
      var rows = lines(tc.items);
      for(var ri = 0; ri < rows.length; ri++){
        var r = rows[ri];
        var t = r.text;
        var am = t.match(/(\d+(?:\.\d+)?)\s*-\s*ILOVA/i);
        if(am){ ctx.annex = am[1]; ctx.titleLines = []; continue; }
        var hc = headerCols(r);
        if(hc && (!cur || cur.closed)){
          cur = { doc: docId, annex: ctx.annex, page: p, title: ctx.titleLines.join(" ").replace(/\s+/g," ").trim(),
                  cols: hc, totalX: null, section: "mandatory", rows: [], totals: [], pendingText: [], closed: false };
          // колонка «всего»: правее последнего класса
          tables.push(cur);
          continue;
        }
        if(!cur || cur.closed){ ctx.titleLines.push(t); if(ctx.titleLines.length > 8) ctx.titleLines.shift(); continue; }
        // строка таблицы
        var firstColX = cur.cols[0].x, lastColX = cur.cols[cur.cols.length-1].x;
        var vals = {}, total = null, labelParts = [], idx = null;
        r.items.forEach(function(i){
          var s = i.s.trim(); if(!s) return;
          var cx = i.x + i.w/2;
          if(NUM.test(s) && cx > firstColX - 25){
            if(cx > lastColX + 22){ total = num(s); return; }
            var best = null, bd = 1e9;
            cur.cols.forEach(function(c){ var d = Math.abs(c.x - cx); if(d < bd){ bd = d; best = c; } });
            if(best && bd < 25){ vals[best.grade] = num(s); return; }
          }
          if(/^\d+$/.test(s) && labelParts.length === 0 && idx === null && cx < firstColX - 60){ idx = s; return; }
          labelParts.push(i.s);
        });
        var label = labelParts.join(" ").replace(/\s+/g," ").trim();
        var hasNums = Object.keys(vals).length > 0 || total !== null;
        if(!hasNums){
          if(/^majburiy fan/i.test(label)){ cur.section = "mandatory"; continue; }
          if(/^tanlov fan/i.test(label)){ cur.section = "elective"; continue; }
          if(/^(Haftalik|umumiy|soat|Sinflar kesimida|\(sinflar)/i.test(label)) continue;
          // перенос названия: строка без чисел — продолжение соседней строки
          var prev = cur.rows[cur.rows.length-1];
          if(prev && /^[a-zʻʼ‘’(]/.test(label) && prev.y - r.y < 14 && !prev.contDone){ prev.name += " " + label; prev.contDone = true; }
          else cur.pendingText.push({ text: label, y: r.y });
          continue;
        }
        if(/^(Jami|UMUMIY|Umumiy)/i.test(label) || (!label && idx === null && cur.pendingText.length && /^umumiy/i.test(cur.pendingText[cur.pendingText.length-1].text))){
          var kind = /^jami/i.test(label) ? "jami" : "umumiy";
          cur.totals.push({ kind: kind, section: cur.section, vals: vals, total: total });
          // после итога обязательной части идёт вторая часть (в бизнес-планах №183 без заголовка «Tanlov fanlari»)
          if(kind === "jami") cur.section = "elective";
          if(kind === "umumiy" || (cur.cols.length && /UMUMIY/.test(label))) cur.closed = true;
          continue;
        }
        if(/^[IVX]+\.\s?[A-ZА-Я]/.test(label)) continue; // подытог группы предметов (приказ №183)
        if(!label){
          // название на строке выше (и, возможно, ниже)
          var pt = cur.pendingText.filter(function(x){ return x.y - r.y < 16 && x.y > r.y; }).pop();
          if(pt){ label = pt.text; cur.pendingText.splice(cur.pendingText.indexOf(pt), 1); }
        }
        cur.rows.push({ section: cur.section, name: label, vals: vals, total: total, y: r.y });
      }
      // «Umumiy» без чисел на строке могло прийти последней строкой страницы
      if(cur && !cur.closed && cur.totals.some(function(x){ return x.kind === "umumiy"; })) cur.closed = true;
    }
    // проверка сумм
    tables.forEach(function(tb){
      tb.rows.forEach(function(rw){ delete rw.y; delete rw.contDone; rw.name = rw.name.replace(/\s+/g," ").replace(/ ’/g,"’").trim(); });
      delete tb.pendingText;
      tb.issues = [];
      ["mandatory","elective"].forEach(function(sec){
        var jam = tb.totals.filter(function(x){ return x.kind === "jami" && x.section === sec; })[0];
        if(!jam) return;
        tb.cols.forEach(function(c){
          var sum = tb.rows.filter(function(rw){ return rw.section === sec; }).reduce(function(a,rw){ return a + (rw.vals[c.grade] || 0); }, 0);
          var exp = jam.vals[c.grade];
          if(exp !== undefined && Math.abs(sum - exp) > 0.01) tb.issues.push(sec+" "+c.grade+": сумма "+sum+" ≠ Jami "+exp);
        });
      });
      tb.rows.forEach(function(rw){
        if(rw.total === null) return;
        var s = Object.keys(rw.vals).reduce(function(a,k){ return a + rw.vals[k]; }, 0);
        if(Math.abs(s - rw.total) > 0.01) tb.issues.push("строка «"+rw.name+"»: "+s+" ≠ итог "+rw.total);
      });
      tb.grades = tb.cols.map(function(c){ return c.grade; }); delete tb.cols; delete tb.closed; delete tb.section; delete tb.totalX;
    });
    return tables;
  };
  window.saveJson = async function(name, data){
    var r = await fetch("/save?name=" + encodeURIComponent(name), { method: "POST", body: JSON.stringify(data, null, 1) });
    return r.status;
  };
})();
