/* Разбор выгрузок eMaktab для импорта: сетка по учителям (.xls BIFF), учебная нагрузка (HTML .xls),
   схема класса из генератора уроков (.xls BIFF). Черновик будущего importEmaktab() в index.html. */
window.An = (function(){
  var ta = document.createElement("textarea");
  function dec(s){ ta.innerHTML = String(s == null ? "" : s); return ta.value.replace(/ /g, " ").trim(); }
  async function wb(url){ return XLSX.read(await (await fetch(url + "?t=" + Date.now())).arrayBuffer(), { type: "array" }); }
  function aoa(w){ return XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]], { header: 1, defval: "", raw: false }); }
  // ФИО к ключу: «Абдерахманова Н. М.» = «Абдерахманова Н.М.»; обрезанное «Шайловбаева Л.Х...» — префикс
  function tkey(n){ return dec(n).replace(/\.\.\.$|…$/, "").replace(/[\s.]/g, "").toLowerCase(); }
  function truncated(s){ return /(\.\.\.|…)$/.test(s); }
  function skey(s){ return dec(s).replace(/(\.\.\.|…)$/, "").toLowerCase().replace(/\s+/g, " ").trim(); }

  /* Сетка по учителям: строка заголовка с «Учитель», в ней дни (объединённые ячейки), ниже номера уроков. */
  function parseGrid(w){
    var A = aoa(w), hr = -1;
    for(var r=0; r<Math.min(A.length, 15); r++){ if(/^учитель|^o.?qituvchi/i.test(dec(A[r][0]))){ hr = r; break; } }
    if(hr < 0) throw new Error("нет строки «Учитель»");
    var colMap = {}, d = -1, dayNames = [];
    for(var c=1; c<A[hr].length; c++){
      var dl = dec(A[hr][c]); if(dl){ d++; dayNames.push(dl); }
      var pn = parseInt(dec(A[hr+1][c]), 10);
      if(d >= 0 && pn > 0) colMap[c] = { d: d, p: pn };
    }
    var lessons = [], teachers = [], odd = [];
    for(var r2=hr+2; r2<A.length; r2++){
      var t = dec(A[r2][0]); if(!t) continue;
      teachers.push(t);
      Object.keys(colMap).forEach(function(c){
        var v = dec(A[r2][c]); if(!v) return;
        var lines = v.split(/\n/).map(dec).filter(Boolean);
        for(var i=0; i+1<lines.length; i+=3){
          var m = lines[i+1].match(/^(\S+)(?:\s*\((.*)\))?$/);
          if(!m){ odd.push(v); return; }
          lessons.push({ t: t, subj: lines[i], cls: m[1], grp: m[2] || "", room: lines[i+2] || "", d: colMap[c].d, p: colMap[c].p });
        }
      });
    }
    return { teachers: teachers, lessons: lessons, days: dayNames, odd: odd };
  }
  /* Нагрузка: Учитель | Предмет | Количество часов | классы… */
  function parseLoad(w){
    var A = aoa(w), hr = -1;
    for(var r=0; r<Math.min(A.length, 15); r++){ if(A[r].some(function(x){ return /^предмет$|^fan$/i.test(dec(x)); })){ hr = r; break; } }
    var head = A[hr].map(dec), ci = head.findIndex(function(x){ return /^учитель|^o.?qituvchi/i.test(x); }), si = head.findIndex(function(x){ return /^предмет$|^fan$/i.test(x); });
    var hi = head.findIndex(function(x){ return /количеств|soat/i.test(x); });
    var classes = head.slice(hi+1), rows = [], t = null;
    for(var r2=hr+1; r2<A.length; r2++){
      var row = A[r2].map(dec);
      if(row[ci]) t = row[ci];
      if(!row[si] || !t) continue;
      classes.forEach(function(cl, k){ var n = parseFloat(row[hi+1+k]) || 0; if(n > 0 && cl) rows.push({ t: t, subj: row[si], cls: cl, n: n }); });
    }
    return { classes: classes.filter(Boolean), rows: rows };
  }
  /* Схема класса: пн…вс × номер урока; в ячейке блоки «Предмет / Весь класс|Подгруппа X / Учитель / Кабинет», разделены пустой строкой. */
  function parseScheme(w){
    var A = aoa(w), cells = [];
    var days = A[0].slice(1).map(dec);
    for(var r=1; r<A.length; r++){
      var p = parseInt(dec(A[r][0]), 10); if(!(p > 0)) continue;
      for(var c=1; c<A[r].length; c++){
        var v = String(A[r][c] || ""); if(!dec(v)) continue;
        v.split(/\n\s*\n/).map(function(b){ return b.split(/\n/).map(dec).filter(Boolean); }).filter(function(b){ return b.length >= 2; }).forEach(function(b){
          cells.push({ d: c-1, p: p, subj: b[0], grp: /^весь класс|^butun sinf/i.test(b[1]) ? "" : b[1].replace(/^Подгруппа\s*|^Guruh\s*/i, ""), t: b[2] || "", room: b[3] || "" });
        });
      }
    }
    return { days: days, cells: cells };
  }
  /* Полное название предмета из нагрузки: у того же учителя и класса, по началу обрезанного названия. */
  function fullNames(grid, load){
    var byTC = {};
    load.rows.forEach(function(x){ var k = tkey(x.t) + "|" + x.cls; (byTC[k] = byTC[k] || []).push(x.subj); });
    var res = { ok: 0, exact: 0, ambiguous: [], missing: [] }, map = {};
    grid.lessons.forEach(function(l){
      var cand = (byTC[tkey(l.t) + "|" + l.cls] || []).filter(function(s){ return skey(s).indexOf(skey(l.subj)) === 0; });
      cand = cand.filter(function(s, i){ return cand.indexOf(s) === i; });
      if(cand.length === 1){ l.full = cand[0]; res.ok++; if(!truncated(l.subj)) res.exact++; }
      else if(cand.length > 1){ var ex = cand.filter(function(s){ return skey(s) === skey(l.subj); }); if(ex.length === 1){ l.full = ex[0]; res.ok++; } else res.ambiguous.push(l.subj + " → " + cand.join(" / ")); }
      else res.missing.push(l.t.split(" ")[0].slice(0,3) + "… " + l.cls + " " + l.subj);
    });
    res.ambiguous = dedupe(res.ambiguous); res.missing = dedupe(res.missing);
    return res;
  }
  function dedupe(a){ return a.filter(function(x, i){ return a.indexOf(x) === i; }); }
  /* Сводка по сетке: конфликты, часы класса, группы, номера уроков. */
  function stats(grid){
    var L = grid.lessons, tSlot = {}, cSlot = {}, dbl = 0, clash = 0;
    L.forEach(function(l){
      var k = tkey(l.t) + "|" + l.d + "|" + l.p; (tSlot[k] = tSlot[k] || []).push(l.cls);
      var ck = l.cls + "|" + l.d + "|" + l.p; (cSlot[ck] = cSlot[ck] || []).push(l);
    });
    Object.keys(tSlot).forEach(function(k){ if(dedupe(tSlot[k]).length > 1) dbl++; });
    // в одном слоте класса несколько уроков без групп — столкновение (или объединение, которое надо понять)
    Object.keys(cSlot).forEach(function(k){ var a = cSlot[k]; if(a.length > 1 && a.some(function(l){ return !l.grp; })) clash++; });
    var cls = {}; L.forEach(function(l){ var c = cls[l.cls] = cls[l.cls] || { slots: {}, ps: {}, days: {} }; c.slots[l.d + "|" + l.p] = 1; c.ps[l.p] = 1; c.days[l.d] = 1; });
    var hours = Object.keys(cls).sort(function(a,b){ return parseInt(a) - parseInt(b) || a.localeCompare(b); }).map(function(k){ return k + ":" + Object.keys(cls[k].slots).length + "ч/" + Object.keys(cls[k].days).length + "дн/ур" + Object.keys(cls[k].ps).sort(function(a,b){ return a-b; }).join(""); });
    return { teachers: grid.teachers.length, lessons: L.length, classes: Object.keys(cls).length, teacherDouble: dbl, classClash: clash, groups: dedupe(L.map(function(l){ return l.grp; }).filter(Boolean)), rooms: dedupe(L.map(function(l){ return l.room; })).length, maxP: Math.max.apply(null, L.map(function(l){ return l.p; })), days: grid.days, hours: hours, odd: grid.odd.length };
  }
  /* Схема ↔ сетка: какой класс, и насколько совпадает (предмет по началу, учитель, группа). */
  function matchScheme(scheme, grid){
    var byCls = {};
    grid.lessons.forEach(function(l){ (byCls[l.cls] = byCls[l.cls] || []).push(l); });
    function same(sc, l){ return sc.d === l.d && sc.p === l.p && tkey(l.t).indexOf(tkey(sc.t).slice(0, 6)) === 0 && (skey(sc.subj).indexOf(skey(l.subj)) === 0 || skey(l.full || "").indexOf(skey(sc.subj)) === 0 || skey(sc.subj).indexOf(skey(l.full || "~")) === 0); }
    var best = null;
    Object.keys(byCls).forEach(function(k){
      var ls = byCls[k], hit = scheme.cells.filter(function(sc){ return ls.some(function(l){ return same(sc, l); }); }).length;
      var score = hit / Math.max(scheme.cells.length, ls.length);
      if(!best || score > best.score) best = { cls: k, score: score, hit: hit, scheme: scheme.cells.length, grid: ls.length };
    });
    return best;
  }
  /* Расписание класса за неделю (вкладка «Классы» → экспорт): «Класс: 6-B», «1 четверть, 21 — 27 сентября 2026»,
     строка дней (пн…сб), строка дат; далее строки уроков: колонка B — номер или «#» (вторая группа в том же слоте).
     Ячейка: «Предмет (группа)⏎Учитель⏎08:00 - 08:45⏎Кабинет». */
  function parseClassWeek(w){
    var A = aoa(w);
    var cm = dec(A[0][0]).match(/^(?:Класс|Sinf)\s*:\s*(.+)$/i); if(!cm) throw new Error("нет «Класс:»");
    var dr = A.findIndex(function(r){ return r.some(function(c){ return /^пн$|^du$/i.test(dec(c)); }); });
    var dayCols = []; A[dr].forEach(function(c, i){ if(/^(пн|вт|ср|чт|пт|сб|вс|du|se|ch|pa|ju|sh|ya)$/i.test(dec(c))) dayCols.push(i); });
    var out = [], p = null;
    for(var r=dr+2; r<A.length; r++){
      var num = dec(A[r][1]);
      if(/^\d+$/.test(num)) p = +num; else if(num !== "#") continue;
      dayCols.forEach(function(ci, d){
        var v = String(A[r][ci] || ""); if(!dec(v)) return;
        var L = v.split("\n").map(dec).filter(Boolean), m = L[0].match(/^(.*?)\s*\((.+)\)$/);
        out.push({ cls: dec(cm[1]), d: d, p: p, subj: m ? m[1] : L[0], grp: m ? m[2] : "", t: L[1] || "", time: L[2] || "", room: L[3] || "" });
      });
    }
    return { cls: dec(cm[1]), period: dec(A[1][0]), lessons: out };
  }
  return { dec: dec, wb: wb, parseGrid: parseGrid, parseClassWeek: parseClassWeek, parseLoad: parseLoad, parseScheme: parseScheme, fullNames: fullNames, stats: stats, matchScheme: matchScheme, tkey: tkey, skey: skey };
})();
