/* Сквозные проверки index.html в браузере (без Node): страница открывается во фрейме,
   сценарии нажимают кнопки как пользователь и сверяют данные в localStorage.
   ID сценариев — из docs/ПМИ_конструктор_расписания.md. Промежуточный вариант до Playwright/Vitest
   (ARCHITECTURE, ADR-11): набор переносится туда без изменения сценариев.
   Внимание: тесты очищают localStorage этого адреса (localhost). */
(function(){
  var frame, W, D, results = [], events = [], names = {};
  function sleep(ms){ return new Promise(function(r){ setTimeout(r, ms); }); }
  // события страницы копятся между перезагрузками фрейма — для проверок каталога и персональных данных (P0-5)
  function harvest(){
    try{
      (W && W.__ttEvents || []).forEach(function(e){ if(events.indexOf(e) < 0) events.push(e); });
      var s = st();
      s.classes.forEach(function(c){ names[c.name] = 1; });
      s.teachers.forEach(function(t){ names[t.name] = 1; });
      s.subjects.forEach(function(x){ if(x.name.length > 3) names[x.name] = 1; if(x.nameUz && x.nameUz.length > 3) names[x.nameUz] = 1; });
    }catch(e){}
  }
  async function load(){
    frame.src = "../../index.html?t=" + Date.now();
    await new Promise(function(r){ frame.onload = r; });
    W = frame.contentWindow; D = W.document;
    await sleep(100);
  }
  async function fresh(){
    harvest();
    frame.contentWindow.localStorage.clear();
    await load();
  }
  async function reload(){ harvest(); await load(); }
  // правка сохранённого состояния и перезагрузка страницы — быстрый способ собрать школу для сценария
  async function setState(mutate){
    var body = JSON.parse(W.localStorage.getItem("timetable_builder_v4"));
    mutate(body.state, body);
    W.localStorage.setItem("timetable_builder_v4", JSON.stringify(body));
    await reload();
  }
  // скачивание в тестах не выполняется: перехватываем клик по ссылке и имена листов книги Excel
  function stubDownloads(){
    var log = { files: [], sheets: null };
    W.HTMLAnchorElement.prototype.click = function(){ log.files.push(this.download); };
    if(W.XLSX){
      var orig = W.XLSX.write;
      W.XLSX.write = function(wb, o){ log.sheets = wb.SheetNames.slice(); return orig.call(this, wb, o); };
    }
    return log;
  }
  function toast(){ return D.getElementById("toast").textContent; }
  function cells(name){ return st().schedule.byClass[cls(name).id]; }
  function filled(grid){ var n = 0; grid.forEach(function(r){ r.forEach(function(c){ if(c) n++; }); }); return n; }
  function windowsOf(byClass){
    var n = 0, late = 0;
    Object.keys(byClass).forEach(function(k){ byClass[k].forEach(function(row){
      var last = -1; row.forEach(function(c, i){ if(c) last = i; });
      for(var q = 0; q < last; q++) if(!row[q]) n++;
      if(last >= 0 && !row[0]) late++;
    }); });
    return { windows: n, late: late };
  }
  // учителя: один на предмет на каждые perGrades параллелей (тесная, но выполнимая нагрузка)
  function addTeachers(s, perGrades, maxPerDay){
    var byKey = {}, seq = 0;
    s.classes.forEach(function(c){
      var plan = s.curriculum[c.id] || {};
      Object.keys(plan).forEach(function(sid){
        var key = sid + "|" + Math.floor((c.grade || 0) / perGrades);
        if(!byKey[key]){ seq++; byKey[key] = "teachT" + seq; s.teachers.push({ id: byKey[key], name: "Учитель Тестовый " + seq, subjects: [sid], maxPerDay: maxPerDay, daysOff: [false,false,false,false,false,false] }); }
        plan[sid].teacherId = byKey[key];
      });
    });
  }
  function botSchool(perGrade){
    ans("start"); ans("criterial");
    for(var g = 1; g <= 11; g++) ans(g >= 5 ? String(perGrade) : "1");
    ans("5"); ans("6"); ans("1"); ans("1"); ans("5"); ans("6"); ans("skip"); ans("yes");
  }
  function st(){ return JSON.parse(W.localStorage.getItem("timetable_builder_v4")).state; }
  function cls(name){ return st().classes.filter(function(c){ return c.name === name; })[0]; }
  function $(sel){ var el = D.querySelector(sel); if(!el) throw new Error("нет элемента " + sel); return el; }
  function click(sel){ $(sel).click(); }
  function change(el, v){ el.value = v; el.dispatchEvent(new W.Event("change", { bubbles: true })); }
  function tab(i){ D.querySelectorAll(".step-btn")[i].click(); }
  function ans(v){ click('[data-bot-ans="' + v + '"]'); }
  function botLast(n){ return Array.prototype.slice.call(D.querySelectorAll(".bubble")).slice(-n).map(function(b){ return b.textContent; }); }
  function botOpts(){ return Array.prototype.map.call(D.querySelectorAll("[data-bot-ans]"), function(b){ return b.getAttribute("data-bot-ans"); }); }
  function classSel(act, name){ return $('[data-act="' + act + '"][data-id="' + cls(name).id + '"]'); }
  function setClass(act, name, v){ tab(0); change(classSel(act, name), v); }
  function autofill(name){ tab(4); click('[data-act="cur-select-class"][data-id="' + cls(name).id + '"]'); click('[data-act="cur-autofill"]'); }
  // часы плана класса; групповая пара (технология мальчики/девочки) — одним слотом
  function planHours(name){
    var s = st(), c = cls(name), pl = s.curriculum[c.id] || {}, sum = 0;
    Object.keys(pl).forEach(function(k){ var e = pl[k]; if(e.syncWith && pl[e.syncWith] && e.syncWith < k) return; sum += e.hours; });
    return sum;
  }
  function planNames(name){ var s = st(); return Object.keys(s.curriculum[cls(name).id] || {}).map(function(k){ return s.subjects.filter(function(x){ return x.id === k; })[0].name; }); }
  function cyr(text, allow){ return (text.match(/\S*[А-Яа-яЁё]\S*/g) || []).filter(function(w){ return !allow.test(w); }); }
  function check(id, name, ok, detail){ results.push({ id: id, name: name, ok: !!ok, detail: detail || "" }); }
  // бот до шага 22: тип школы, по одному классу в параллели, план да/нет
  function botToGen(type, plan){
    ans("start"); ans(type);
    for(var g = 1; g <= 11; g++) ans("1");
    ans("5"); ans("6"); ans("1"); ans("1"); ans("5"); ans("6"); ans("skip"); ans(plan);
  }

  var scenarios = [
    ["T-11-01", "Бот: критериальная, без учителей — расписание составлено, шаг «Что дальше?»", async function(){
      await fresh(); botToGen("criterial", "yes"); ans("gen");
      var s = st();
      check("T-11-01", "расписание построено", !!s.schedule);
      check("T-8-17", "0 неразмещённых (групповая пара без учителя не распадается)", s.schedule.unplacedCount === 0, "неразмещённых: " + s.schedule.unplacedCount);
      check("T-11-01", "бот на шаге «Что дальше?»", /Что дальше/.test(botLast(1)[0]), botLast(1)[0]);
      check("T-11-01", "кнопки разбора", ["regen", "go:schedule", "finish"].every(function(v){ return botOpts().indexOf(v) >= 0; }), botOpts().join(","));
      check("T-7-01", "без учителей — предупреждение, не ошибка", !D.querySelector(".warn-box") && /Учителей нет/.test($(".note-box").textContent));
    }],
    ["T-11-11", "Бот: план не заполнен — ошибка, переход к месту, «Продолжить»", async function(){
      await fresh(); botToGen("criterial", "no"); ans("gen");
      check("T-11-11", "бот сообщил об ошибке", /не могу составить/i.test(botLast(1)[0]), botLast(1)[0]);
      check("T-11-11", "бот остался на шаге 22", botOpts().indexOf("continue") >= 0 && botOpts().indexOf("later") >= 0, botOpts().join(","));
      check("T-11-11", "открыт «Учебный план»", /Учебный план/.test($(".page-title").textContent));
      var sel = $("#cur-add-subj-select");
      sel.value = Array.prototype.filter.call(sel.options, function(o){ return /^Математика/.test(o.text); })[0].value;
      click('[data-act="cur-add-subject"]');
      change($('[data-act="cur-hours"]'), "5");
      ans("continue");
      check("T-11-11", "после «Продолжить» расписание составлено", !!st().schedule && /Что дальше/.test(botLast(1)[0]));
      D.getElementById("langToggle").click();
      var bad = cyr(D.getElementById("bot").innerText, /^\d+-[A-Z]/);
      check("T-11-13", "история бота на узбекском без кириллицы", bad.length === 0, bad.slice(0, 5).join(" "));
      D.getElementById("langToggle").click();
    }],
    ["T-9-07", "«По учителям» открывает вид по учителям", async function(){
      await fresh(); botToGen("criterial", "yes"); ans("gen");
      tab(2); $("#new-teach-name").value = "Иванова"; click('[data-act="teach-add"]');
      tab(5); click('[data-act="generate"]'); click('[data-id="__teachers__"]');
      check("T-9-07", "активна кнопка «По учителям»", /По учителям/.test($(".pill-btn.active").textContent));
      check("T-9-07", "карточка учителя", /Иванова/.test($("#main").textContent));
    }],
    ["T-9-05", "Закрепление урока: значок, закрепление, повторная генерация, открепление", async function(){
      await fresh(); botToGen("criterial", "yes"); ans("gen");
      click('[data-act="sched-select-class"][data-id="' + cls("5-A").id + '"]');
      var btn = $('.lock-btn[data-act="cell-lock"]');
      check("T-9-05", "значок — SVG, без эмодзи", !!btn.querySelector("svg") && !/[\u{1F512}\u{1F513}]/u.test($("#main").innerHTML));
      var d = btn.getAttribute("data-day"), p = btn.getAttribute("data-period");
      var before = JSON.stringify(st().schedule.byClass[cls("5-A").id][d][p].lessons);
      // клик по внутреннему элементу значка, как это делает мышь
      btn.querySelector("path").dispatchEvent(new W.MouseEvent("click", { bubbles: true }));
      var lockedTd = $('.lock-btn.is-locked[data-day="' + d + '"][data-period="' + p + '"]');
      check("T-9-05", "ячейка закреплена (клик по значку)", !!lockedTd && lockedTd.closest("td").classList.contains("locked-cell") && st().schedule.byClass[cls("5-A").id][d][p].locked === true);
      click('[data-act="generate"]');
      var after = st().schedule.byClass[cls("5-A").id][d][p];
      check("T-8-11", "после повторной генерации урок на месте и закреплён", after && after.locked && JSON.stringify(after.lessons) === before);
      click('.lock-btn.is-locked[data-day="' + d + '"][data-period="' + p + '"]');
      check("T-9-05", "открепление возвращает выбор урока", !st().schedule.byClass[cls("5-A").id][d][p].locked && !!D.querySelector('select[data-act="cell-change"][data-day="' + d + '"][data-period="' + p + '"]'));
    }],
    ["T-9-11", "Цвета групп предметов", async function(){
      await fresh(); botToGen("criterial", "yes"); ans("gen");
      click('[data-act="sched-select-class"][data-id="' + cls("7-A").id + '"]');
      var filled = Array.prototype.filter.call(D.querySelectorAll("td select[data-act=cell-change]"), function(s){ return s.value; });
      var noColor = filled.filter(function(s){ return !/cat-/.test(s.closest("td").className); });
      check("T-9-11", "у каждого урока есть цвет группы", filled.length > 0 && noColor.length === 0, noColor.length + " без цвета");
      check("T-9-11", "легенда из 7 групп", D.querySelectorAll(".subj-legend > span").length === 7);
      check("T-9-11", "без учителя нет «— —»", !/— —/.test($("#main").textContent));
    }],
    ["T-6-19..22", "Планы приказов №227 и №183", async function(){
      await fresh(); botToGen("criterial", "yes");
      setClass("class-direction", "7-A", "v:philology"); autofill("7-A");
      check("T-6-19", "7-A Филология: 35 ч", planHours("7-A") === 35, planHours("7-A"));
      check("T-6-19", "7-A: предметы по выбору — «вариатив»", ["Родной язык вариатив", "Английский язык вариатив"].every(function(n){ return planNames("7-A").indexOf(n) >= 0; }));
      check("ADR-4", "класс запомнил источник плана", (cls("7-A").appliedPlan || {}).ref === "order227@2026-06-26", JSON.stringify(cls("7-A").appliedPlan));
      setClass("class-direction", "11-A", "t:ix"); setClass("class-secondlang", "11-A", "korean"); autofill("11-A");
      check("T-6-20", "11-A Иностранные языки: 31 ч, один второй язык", planHours("11-A") === 31 && planNames("11-A").filter(function(n){ return /второй/.test(n); }).join() === "Корейский язык как второй иностранный язык", planHours("11-A"));
      setClass("class-lang", "8-A", "kz"); setClass("class-qsample", "8-A", "2"); setClass("class-direction", "8-A", "v:law"); autofill("8-A");
      check("T-6-21", "8-A казахский, образец 2, Юриспруденция: 33 ч, без русского", planHours("8-A") === 33 && planNames("8-A").indexOf("Русский язык") < 0, planHours("8-A"));
      setClass("class-lang", "9-A", "qr");
      var dir9 = classSel("class-direction", "9-A");
      check("T-6-22", "каракалпакский: направлений нет", dir9.querySelectorAll("optgroup").length === 0);
      tab(4); click('[data-act="cur-select-class"][data-id="' + cls("9-A").id + '"]');
      check("T-6-22", "подсказка про Каракалпакстан", /Каракалпакстан/.test($("#main").textContent));
    }],
    ["T-2-02", "Классическая школа: направления №183 недоступны", async function(){
      await fresh(); botToGen("criterial", "yes");
      setClass("class-direction", "10-A", "t:mf");
      tab(0); change($("#cfg-schooltype"), "classical");
      var sel = classSel("class-direction", "10-A");
      check("T-2-02", "нет группы №183", !Array.prototype.some.call(sel.querySelectorAll("optgroup"), function(g){ return /183/.test(g.label); }));
      check("T-2-02", "выбранное помечено «только для критериальной»", /критериальной/.test(sel.options[sel.selectedIndex].text));
      autofill("10-A");
      check("T-2-02", "подставлен базовый план", /Базовый учебный план/.test($("#main").textContent) && (cls("10-A").appliedPlan || {}).source === "base");
      tab(0); change($("#cfg-schooltype"), "criterial"); autofill("10-A");
      check("T-2-02", "обратно: план №183, 31 ч", planHours("10-A") === 31, planHours("10-A"));
    }],
    ["T-1-04", "Узбекский интерфейс без кириллицы", async function(){
      await fresh(); botToGen("criterial", "yes");
      setClass("class-direction", "10-A", "t:nd1");
      D.getElementById("langToggle").click();
      var allow = /^\d+-[A-Z]/, bad = [];
      [0, 4, 5, 6].forEach(function(i){ tab(i); bad = bad.concat(cyr($("#main").innerText, allow)); });
      // предметы — из справочника ЕСП, у них свои узбекские названия; пользовательских данных тут нет
      check("T-1-04", "вкладки «Классы», «План», «Расписание», «Экспорт»", bad.length === 0, bad.slice(0, 8).join(" "));
      D.getElementById("langToggle").click();
    }],
    ["T-6-24", "Данные приказов: итог часов каждой таблицы", async function(){
      var P = W.ORDER_PLANS || null;
      // ORDER_PLANS внутри замыкания страницы — читаем из исходника
      if(!P){
        var src = await (await fetch("../../index.html")).text();
        var m = src.match(/var ORDER_PLANS = (\{[\s\S]*?\});\n\/\* ORDER_PLANS:END/);
        P = JSON.parse(m[1]);
      }
      var TOTAL = { "227": { 7: 35, 8: 33, 9: 34, 10: 31, 11: 31 }, "183": { 10: 31, 11: 31 } };
      var bad = [], n = 0;
      ["227", "183"].forEach(function(o){
        Object.keys(P[o]).forEach(function(k){ Object.keys(P[o][k]).forEach(function(l){
          var tb = P[o][k][l]; n++;
          tb.g.forEach(function(g, gi){
            var sum = 0, alt = 0;
            tb.r.forEach(function(r){
              var name = P.n[r[0]], h = r[2 + gi];
              if(/^(Ingliz|Koreys|Nemis|Fransuz|Xitoy|Yapon) tili$/.test(name)){ alt = Math.max(alt, h); return; } // альтернативы: один из шести
              sum += h;
            });
            sum += alt;
            if(Math.abs(sum - TOTAL[o][g]) > 0.01) bad.push(o + "/" + k + "/" + l + " кл." + g + ": " + sum);
          });
        }); });
      });
      check("T-6-24", "88 таблиц, итог по классу = «Umumiy soat»", n === 88 && bad.length === 0, "таблиц " + n + (bad.length ? "; " + bad.slice(0, 5).join("; ") : ""));
    }],

    /* ---------------- P0 ---------------- */
    ["T-8-26", "P0-1: уроки класса подряд с 1-го урока, без окон", async function(){
      await fresh(); botSchool(2);
      await setState(function(s){ addTeachers(s, 4, 6); s.schedule = null; });
      tab(5); click('[data-act="generate"]');
      var s = st(), w = windowsOf(s.schedule.byClass);
      check("T-8-26", "18 классов, 44 учителя: всё размещено", s.schedule.unplacedCount === 0, "неразмещённых: " + s.schedule.unplacedCount);
      check("T-8-26", "нет окон между уроками", w.windows === 0, "окон: " + w.windows);
      check("T-8-26", "каждый день начинается с 1-го урока", w.late === 0, "дней не с 1-го: " + w.late);
      check("T-8-26", "число окон в итогах составления", /окон у классов: 0/.test($(".gen-summary").textContent), $(".gen-summary").textContent);
      var gr = events.concat(W.__ttEvents).filter(function(e){ return e.goal === "tt_generate_result"; }).pop();
      check("T-13-01", "generate_result: окна, конфликты, длительность", gr && gr.params.windows === 0 && gr.params.conflicts === 0 && gr.params.duration_ms >= 0, gr && JSON.stringify(gr.params));
    }],
    ["T-8-22", "P0-2: правки данных не стирают расписание, закреплённые и ручные правки остаются", async function(){
      await fresh(); botToGen("criterial", "yes"); ans("gen");
      click('[data-act="sched-select-class"][data-id="' + cls("7-A").id + '"]');
      var lockBtn = $('.lock-btn[data-act="cell-lock"]'), ld = +lockBtn.getAttribute("data-day"), lp = +lockBtn.getAttribute("data-period");
      lockBtn.click();
      var before = JSON.stringify(cells("7-A")[ld][lp].lessons);
      // ручная правка: другой урок в незакреплённую ячейку
      var sel = Array.prototype.filter.call(D.querySelectorAll('select[data-act="cell-change"]'), function(x){ return x.value; })[0];
      var md = +sel.getAttribute("data-day"), mp = +sel.getAttribute("data-period");
      var other = Array.prototype.filter.call(sel.options, function(o){ return o.value && o.value !== sel.value; })[0].value;
      change(sel, other);
      check("T-8-22", "ручная правка помечена", cells("7-A")[md][mp].manual === true);
      // правка часов в плане другого класса
      tab(4); click('[data-act="cur-select-class"][data-id="' + cls("8-A").id + '"]');
      var h = $('[data-act="cur-hours"]'); change(h, String(Math.max(1, +h.value - 1)));
      var s = st();
      check("T-8-22", "после правки часов расписание не стёрто и помечено «устарело»", !!s.schedule && s.schedule.stale === true);
      check("T-8-22", "закреплённый урок на месте", s.schedule.byClass[cls("7-A").id][ld][lp].locked && JSON.stringify(s.schedule.byClass[cls("7-A").id][ld][lp].lessons) === before);
      check("T-8-22", "ручная правка на месте", s.schedule.byClass[cls("7-A").id][md][mp].manual === true);
      tab(5);
      check("T-8-22", "плашка «расписание устарело»", /устарело/.test($(".stale-box").textContent));
      // предмет убран из плана — его уроки убираются с сообщением
      tab(4); click('[data-act="cur-select-class"][data-id="' + cls("9-A").id + '"]');
      var rm = $('[data-act="cur-remove-subject"]'), rmSubj = rm.getAttribute("data-subj");
      var had = 0; cells("9-A").forEach(function(r){ r.forEach(function(c){ if(c && c.lessons.some(function(l){ return l.subjectId === rmSubj; })) had++; }); });
      rm.click();
      var left = 0; cells("9-A").forEach(function(r){ r.forEach(function(c){ if(c && c.lessons.some(function(l){ return l.subjectId === rmSubj; })) left++; }); });
      check("T-8-23", "уроки убранного предмета удалены, остальные на месте", had > 0 && left === 0 && filled(cells("9-A")) > 0, "было " + had + ", осталось " + left);
      check("T-8-23", "сообщение, сколько убрано", new RegExp("невозможными: " + had).test(toast()), toast());
      // «Перестроить»: вопрос о ручной правке; «Отмена» ничего не меняет
      tab(5); click('[data-act="generate"]');
      check("T-8-27", "перед перестроением — вопрос о ручных правках", !D.getElementById("dialog").hidden && /без замка: 1/.test(D.getElementById("dialog").textContent), D.getElementById("dialog").textContent);
      click('[data-dialog="cancel"]');
      check("T-8-27", "«Отмена» — расписание не перестроено", st().schedule.stale === true);
      click('[data-act="generate"]'); click('[data-dialog="lock"]');
      s = st();
      var mc = s.schedule.byClass[cls("7-A").id][md][mp];
      check("T-8-27", "«Закрепить и перестроить»: ручная правка закреплена и на месте", !s.schedule.stale && mc && mc.locked && mc.lessons[0].subjectId === other.split("+")[0]);
      check("T-8-11", "закреплённый урок пережил перестроение", JSON.stringify(s.schedule.byClass[cls("7-A").id][ld][lp].lessons) === before);
      var inv = events.concat(W.__ttEvents).filter(function(e){ return e.goal === "tt_schedule_invalidated"; });
      check("T-13-01", "schedule_invalidated: причина и сохранённые правки", inv.length >= 2 && inv[0].params.reason === "plan" && inv[0].params.locked === 1, inv.map(function(e){ return JSON.stringify(e.params); }).join(" "));
    }],
    ["T-10-10", "P0-3: выгрузка Excel при одинаковых и длинных названиях, классе после составления", async function(){
      await fresh(); botToGen("criterial", "yes"); ans("gen");
      await setState(function(s){
        s.classes[1].name = s.classes[0].name; // одинаковые названия
        var subj = Object.keys(s.curriculum[s.classes[5].id])[0];
        s.teachers.push({ id: "teachL1", name: "Абдурахмонова-Каримова Гулнора Шавкатовна", subjects: [], maxPerDay: 6, daysOff: [false,false,false,false,false,false] });
        s.teachers.push({ id: "teachL2", name: "Абдурахмонова-Каримова Гулнора Шухратовна", subjects: [], maxPerDay: 6, daysOff: [false,false,false,false,false,false] });
        s.curriculum[s.classes[5].id][subj].teacherId = "teachL1";
      });
      // класс, добавленный после составления (сетки у него ещё нет)
      click('[data-act="parallel-add"][data-grade="7"]');
      tab(6);
      var log = stubDownloads();
      click('[data-act="export-xlsx"]');
      await sleep(50);
      var sh = log.sheets || [];
      var uniq = sh.filter(function(n, i){ return sh.map(function(x){ return x.toLowerCase(); }).indexOf(n.toLowerCase()) === i; });
      check("T-10-10", "файл сформирован и скачан", log.files.indexOf("raspisanie.xlsx") >= 0, log.files.join());
      check("T-10-10", "имена листов уникальны и не длиннее 31", sh.length > 0 && uniq.length === sh.length && sh.every(function(n){ return n.length <= 31; }), sh.filter(function(n){ return / \(\d\)$/.test(n); }).join(" | "));
      var ex = W.__ttEvents.filter(function(e){ return e.goal === "tt_export_xlsx_result"; }).pop();
      check("T-13-01", "export_xlsx_result — по результату", ex && ex.params.ok === true && ex.params.sheets === sh.length, ex && JSON.stringify(ex.params));
      // пустое название класса не сохраняется
      tab(0);
      var nameInput = classSel("class-name", "5-A"); change(nameInput, "   ");
      check("T-2-20", "пустое название класса не сохраняется", !!cls("5-A") && /не может быть пустым/.test(toast()));
      // сбой сборки — сообщение, а не тишина
      tab(6); W.XLSX.utils.book_new = function(){ throw new Error("x"); };
      click('[data-act="export-xlsx"]');
      check("T-10-10", "при сбое — сообщение", /Не удалось сформировать Excel/.test(toast()), toast());
    }],
    ["T-12-11", "P0-4: запись только при изменениях, ошибка записи, повреждённое сохранение, вторая вкладка", async function(){
      await fresh(); botToGen("criterial", "yes");
      // перерисовка без изменений не пишет в хранилище
      var writes = 0, origSet = W.Storage.prototype.setItem;
      W.Storage.prototype.setItem = function(k, v){ if(k === "timetable_builder_v4") writes++; return origSet.call(this, k, v); };
      tab(1); tab(2); tab(3); tab(0);
      check("T-12-11", "переходы по шагам без правок — 0 записей", writes === 0, "записей: " + writes);
      // ошибка записи — плашка
      W.Storage.prototype.setItem = function(k){ if(k === "timetable_builder_v4"){ var e = new W.DOMException("full", "QuotaExceededError"); throw e; } };
      click('[data-act="parallel-add"][data-grade="5"]');
      check("T-12-12", "при переполнении и правке — «Браузер не сохраняет изменения»", /не сохраняет изменения/.test(D.getElementById("storageBar").textContent));
      W.Storage.prototype.setItem = origSet;
      click('[data-act="parallel-add"][data-grade="6"]');
      check("T-12-12", "после восстановления плашка уходит", !/не сохраняет/.test(D.getElementById("storageBar").textContent));
      // хранилище запрещено: пока правок нет — плашки нет (переходы по шагам ничего не пишут)
      await fresh();
      var gs = W.Storage.prototype.getItem, ss = W.Storage.prototype.setItem;
      W.Storage.prototype.getItem = W.Storage.prototype.setItem = function(){ throw new W.DOMException("denied", "SecurityError"); };
      tab(1); tab(4); tab(0);
      check("T-12-15", "запрет хранилища без правок — плашки нет", !/не сохраняет/.test(D.getElementById("storageBar").textContent));
      click('[data-act="parallel-add"][data-grade="5"]');
      check("T-12-15", "после первой правки — спокойная плашка «сохраните копию»", /не сохраняет изменения/.test(D.getElementById("storageBar").textContent) && !!D.querySelector("#storageBar .note-box"));
      W.Storage.prototype.getItem = gs; W.Storage.prototype.setItem = ss;
      // повреждённое сохранение откладывается, а не затирается
      W.localStorage.setItem("timetable_builder_v4", "{повреждено");
      await reload();
      check("T-12-13", "повреждённое сохранение отложено в копию", W.localStorage.getItem("timetable_builder_v4_corrupt") === "{повреждено");
      check("T-12-13", "пользователь видит, что копия есть", /не удалось прочитать/.test(D.getElementById("storageBar").textContent));
      // вторая вкладка: эта становится «только для чтения»
      await fresh(); botToGen("criterial", "yes");
      var saved = W.localStorage.getItem("timetable_builder_v4");
      localStorage.setItem("timetable_builder_tab", JSON.stringify({ id: "other-tab", ts: Date.now() }));
      await sleep(50);
      check("T-12-14", "другая вкладка → «только для чтения»", D.body.classList.contains("is-readonly") && /другой вкладке/.test(D.getElementById("storageBar").textContent));
      click('[data-act="parallel-add"][data-grade="5"]');
      check("T-12-14", "в режиме чтения работа другой вкладки не затирается", W.localStorage.getItem("timetable_builder_v4") === saved);
    }],
    ["T-6-25", "P0-6: групповые уроки — разные кабинеты, вместимость группы, один учитель, общий блок", async function(){
      await fresh(); botToGen("criterial", "yes");
      await setState(function(s){
        var c = s.classes.filter(function(x){ return x.name === "7-A"; })[0];
        c.studentCount = 30;
        var inf = s.subjects.filter(function(x){ return x.name === "Информатика и информационные технологии"; })[0];
        inf.roomType = "комп";
        s.rooms = [{ id: "roomA", name: "Комп-1", capacity: 15, type: "комп" }, { id: "roomB", name: "Комп-2", capacity: 16, type: "комп" }, { id: "roomC", name: "Актовый", capacity: 40, type: "зал" }];
        s.subjects.push({ id: "subjG2", name: "Информатика и информационные технологии (группа 2)", roomType: "комп", heavy: false });
        var plan = s.curriculum[c.id];
        plan[inf.id] = { hours: 2, teacherId: null, blockSize: 1, syncWith: "subjG2", groupSize: 15 };
        plan.subjG2 = { hours: 2, teacherId: null, blockSize: 1, syncWith: inf.id, groupSize: 15 };
        s.teachers.push({ id: "teachI1", name: "Информатик Первый", subjects: [], maxPerDay: 6, daysOff: [false,false,false,false,false,false] });
        s.teachers.push({ id: "teachI2", name: "Информатик Второй", subjects: [], maxPerDay: 6, daysOff: [false,false,false,false,false,false] });
        plan[inf.id].teacherId = "teachI1"; plan.subjG2.teacherId = "teachI1";
      });
      tab(5);
      check("T-6-25", "один учитель в двух группах — ошибка до составления", /один учитель/.test($(".warn-box").textContent) && $('[data-act="generate"]').disabled);
      await setState(function(s){ s.curriculum[s.classes.filter(function(x){ return x.name === "7-A"; })[0].id].subjG2.teacherId = "teachI2"; });
      tab(5); click('[data-act="generate"]');
      var bad = [], pairs = 0, infId0 = st().subjects.filter(function(x){ return x.name === "Информатика и информационные технологии"; })[0].id;
      cells("7-A").forEach(function(r){ r.forEach(function(c){
        if(!c || c.lessons.length !== 2 || c.lessons.every(function(l){ return l.subjectId !== infId0; })) return;
        pairs++;
        var r1 = c.lessons[0].roomId, r2 = c.lessons[1].roomId;
        if(!r1 || !r2 || r1 === r2 || r1 === "roomC" || r2 === "roomC") bad.push(r1 + "/" + r2);
      }); });
      check("T-6-26", "группы в один слот — в разных компьютерных классах на 15 и 16 мест", pairs === 2 && bad.length === 0, "пар " + pairs + (bad.length ? ", плохо: " + bad.join(" ") : ""));
      // блок пары синхронизирован
      tab(4); click('[data-act="cur-select-class"][data-id="' + cls("7-A").id + '"]');
      var infId = st().subjects.filter(function(x){ return x.name === "Информатика и информационные технологии"; })[0].id;
      change($('[data-act="cur-blocksize"][data-subj="' + infId + '"]'), "2");
      check("T-6-27", "блок второй группы следует за первой", st().curriculum[cls("7-A").id].subjG2.blockSize === 2);
    }],
    ["T-2-21", "P0-7: параллель меняется всегда; 6 → 5 дней убирает субботу с сообщением", async function(){
      await fresh(); botToGen("criterial", "yes"); ans("gen");
      tab(0);
      var g = classSel("class-grade", "9-A");
      check("T-2-21", "у класса с параллелью её можно изменить", g && g.value === "9");
      // закрепляем субботний урок у 8-A, затем делаем класс 5-дневным
      tab(5); click('[data-act="sched-select-class"][data-id="' + cls("8-A").id + '"]');
      $('.lock-btn[data-act="cell-lock"][data-day="5"]').click();
      var sat = filled([cells("8-A")[5]]);
      setClass("class-days", "8-A", "5");
      check("T-2-22", "суббота убрана, остальное на месте", cells("8-A").length === 5 && filled(cells("8-A")) > 0);
      check("T-2-22", "сообщение, сколько убрано", new RegExp("невозможными: " + sat).test(toast()), toast());
      setClass("class-grade", "9-A", "10");
      check("T-2-21", "параллель 9-A → 10 изменена без удаления класса", cls("9-A").grade === 10 && !!st().schedule);
      tab(5); click('[data-act="generate"]');
      check("T-2-22", "после смены дней и параллели составление работает", !st().schedule.stale && cells("8-A").length === 5);
      tab(6); var log = stubDownloads(); click('[data-act="export-xlsx"]'); await sleep(30);
      check("T-2-22", "и выгрузка работает", log.files.indexOf("raspisanie.xlsx") >= 0);
    }],
    ["T-13-06", "P0-5: каталог событий, воронка, без персональных данных", async function(){
      await fresh();
      // воронка одной сессией: бот → план → учителя → составить → выгрузить
      botToGen("criterial", "yes");
      await setState(function(s){ addTeachers(s, 12, 7); });
      tab(4); click('[data-act="cur-select-class"][data-id="' + cls("7-A").id + '"]'); change($('[data-act="cur-hours"]'), $('[data-act="cur-hours"]').value);
      tab(0); change(classSel("class-students", "5-A"), "25");
      tab(5); click('[data-act="generate"]');
      tab(6); var log = stubDownloads(); click('[data-act="export-xlsx"]'); await sleep(30);
      var ms = W.__ttEvents.filter(function(e){ return e.goal === "tt_milestone"; }).map(function(e){ return e.params.name; });
      var funnel = ["classes_configured", "plan_ready", "teachers_assigned", "schedule_generated", "schedule_clean", "exported"];
      check("T-13-06", "воронка «открыл → выгрузил» собирается", funnel.every(function(m){ return ms.indexOf(m) >= 0; }), ms.join(" → "));
      // остальные события каталога
      click('[data-act="export-json"]');
      var inp = $("#import-json-input"), dt = new W.DataTransfer();
      dt.items.add(new W.File([W.localStorage.getItem("timetable_builder_v4")], "n.json", { type: "application/json" }));
      inp.files = dt.files; inp.dispatchEvent(new W.Event("change", { bubbles: true })); await sleep(50);
      tab(2);
      var staff = "<table><tr><td>№</td><td>ФИО</td><td>Классы</td></tr><tr><td>1</td><td>Каримова Дилноза</td><td>5-B</td></tr></table>";
      var si = $("#import-staff-input"), dt2 = new W.DataTransfer();
      dt2.items.add(new W.File([staff], "s.xls", { type: "text/html" }));
      si.files = dt2.files; si.dispatchEvent(new W.Event("change", { bubbles: true })); await sleep(50);
      tab(0); click('[data-act="preset-create"]'); click('[data-act="parallel-add"][data-grade="3"]');
      await setState(function(s){ s.curriculum[s.classes[0].id] = {}; Object.keys(s.curriculum).forEach(function(k){ s.curriculum[k] = {}; }); });
      tab(5); click('[data-act="goto-issue"]');
      D.getElementById("langToggle").click(); D.getElementById("langToggle").click();
      W.dispatchEvent(new W.ErrorEvent("error", { message: "тест", lineno: 1 }));
      W.confirm = function(){ return true; }; tab(6); click('[data-act="reset-all"]');
      // бот: переход к месту ошибки
      await fresh(); botToGen("criterial", "no"); ans("gen"); var go = botOpts().filter(function(v){ return v.indexOf("go:") === 0; })[0]; ans(go);
      // хранилище: ошибка записи и вторая вкладка
      var os = W.Storage.prototype.setItem;
      W.Storage.prototype.setItem = function(k){ if(k === "timetable_builder_v4") throw new W.DOMException("full", "QuotaExceededError"); return os.apply(this, arguments); };
      tab(0); click('[data-act="parallel-add"][data-grade="5"]');
      W.Storage.prototype.setItem = os;
      localStorage.setItem("timetable_builder_tab", JSON.stringify({ id: "other", ts: Date.now() })); await sleep(30);
      harvest();
      var cat = W.__ttCatalog, fired = {};
      events.forEach(function(e){ fired[e.goal.slice(3)] = true; });
      var missing = Object.keys(cat).filter(function(k){ return !fired[k]; });
      check("T-13-04", "каждое событие каталога вызывается", missing.length === 0, "не вызваны: " + missing.join(", "));
      check("T-13-04", "событий вне каталога нет", !(W.__ttUnknownEvents || []).length, (W.__ttUnknownEvents || []).join());
      // ФТ-13.3: в параметрах нет ФИО, названий классов и предметов; у каждого события — install_id
      var nameList = Object.keys(names).filter(function(n){ return n.length > 2; }), leaks = [];
      events.forEach(function(e){ Object.keys(e.params).forEach(function(k){
        var v = String(e.params[k]);
        if(/[А-Яа-яЁё]/.test(v) || nameList.some(function(n){ return v.indexOf(n) >= 0 && !/^\d+$/.test(n); })) leaks.push(e.goal + "." + k + "=" + v);
      }); });
      check("T-13-02", "в событиях нет ФИО, названий классов и предметов", leaks.length === 0 && events.length > 50, leaks.slice(0, 5).join("; ") + " (событий " + events.length + ")");
      check("T-13-02", "у событий есть анонимный install_id", events.every(function(e){ return /^[0-9a-f-]{36}$/.test(e.params.install_id); }));
    }],
    ["T-9-12", "Порядок классов, «Открепить все», бот без эмодзи", async function(){
      await fresh(); botSchool(2); ans("gen");
      var pills = Array.prototype.map.call(D.querySelectorAll('[data-act="sched-select-class"]'), function(b){ return b.textContent; }).filter(function(n){ return n !== "По учителям"; });
      var expected = ["1-A","2-A","3-A","4-A","5-A","5-B","6-A","6-B","7-A","7-B","8-A","8-B","9-A","9-B","10-A","10-B","11-A","11-B"];
      check("T-9-12", "классы по порядку: 1-е, 2-е … внутри — по букве", pills.join() === expected.join(), pills.join(" "));
      click('[data-act="sched-select-class"][data-id="' + cls("5-A").id + '"]');
      $('.lock-btn[data-act="cell-lock"]').click();
      var td = $(".locked-cell");
      check("T-9-05", "закреплённая ячейка без синей рамки", W.getComputedStyle(td).boxShadow === "none", W.getComputedStyle(td).boxShadow);
      click('[data-act="unlock-all"]');
      check("T-9-12", "«Открепить все» снимает все замки", !D.querySelector(".locked-cell") && !D.querySelector('[data-act="unlock-all"]'));
      check("T-11-15", "в боте нет эмодзи", !/[\u{1F300}-\u{1FAFF}]/u.test(D.getElementById("bot").innerHTML));
    }],
    ["T-13-03", "P0-5: код школы из ссылки", async function(){
      await fresh();
      frame.src = "../../index.html?school=sch_123&t=" + Date.now();
      await new Promise(function(r){ frame.onload = r; }); W = frame.contentWindow; D = W.document; await sleep(80);
      var open = W.__ttEvents.filter(function(e){ return e.goal === "tt_app_open"; })[0];
      check("T-13-03", "school в параметрах события", open && open.params.school === "sch_123", open && JSON.stringify(open.params));
      W.localStorage.removeItem("tt_school");
    }]
  ];

  window.runSmoke = async function(){
    frame = document.getElementById("app");
    results = [];
    for(var i = 0; i < scenarios.length; i++){
      try{ await scenarios[i][2](); }
      catch(e){ check(scenarios[i][0], scenarios[i][1] + " — ошибка сценария", false, e.message); }
    }
    var out = document.getElementById("out"), fail = results.filter(function(r){ return !r.ok; }).length;
    out.innerHTML = "<p><strong>" + (fail ? "FAIL: " + fail : "PASS") + "</strong> — проверок " + results.length + "</p><table>" +
      results.map(function(r){ return "<tr class='" + (r.ok ? "ok" : "bad") + "'><td>" + (r.ok ? "PASS" : "FAIL") + "</td><td>" + r.id + "</td><td>" + r.name + "</td><td>" + String(r.detail).replace(/</g, "&lt;") + "</td></tr>"; }).join("") + "</table>";
    window.__smoke = { fail: fail, results: results };
    return window.__smoke;
  };
})();
