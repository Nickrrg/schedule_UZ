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
  // keepChoice — оставить первый вопрос бота («Расписание уже есть в eMaktab?»); иначе выбирается путь «с нуля»
  async function fresh(keepChoice){
    harvest();
    frame.contentWindow.localStorage.clear();
    await load();
    var scratch = D.querySelector('#bot [data-bot-act="path"][data-v="scratch"]');
    if(!keepChoice && scratch) scratch.click();
  }
  // обезличенные выгрузки eMaktab (tests/fixtures/emaktab) — файлом во фрейм
  async function fixture(name){
    var b = await (await fetch("../fixtures/emaktab/" + name)).arrayBuffer();
    return new W.File([b], name);
  }
  // текст группы «Сверка с учебным планом» на экране «Проверка»
  function planText(){
    var d = D.querySelector('[data-test="check-group-plan"]');
    return d ? d.textContent : "";
  }
  // выбрать файлы в поле загрузки eMaktab и дождаться разбора
  async function emkUpload(files, scope){
    var inp = $((scope || "") + "[data-emk-file]"), dt = new W.DataTransfer();
    files.forEach(function(f){ dt.items.add(f); });
    inp.files = dt.files; inp.dispatchEvent(new W.Event("change", { bubbles: true }));
    for(var i = 0; i < 300; i++){
      await sleep(50);
      var dz = D.querySelector((scope || "") + ".emk .drop-title");
      if(dz && !/Читаю/.test(dz.textContent)) break;
    }
    await sleep(30);
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
      W.XLSX.write = function(wb, o){ log.sheets = wb.SheetNames.slice(); var r = orig.call(this, wb, o); log.bytes = r; return r; };
    }
    return log;
  }
  function toast(){ return T("toast").textContent; }
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
  // помощник «с нуля» (0.12, Э10): по perGrade классов в 5–11, 1–4 — по 5 уроков; botToAbout останавливается на «О школе»,
  // botSchool — на шаге «План» с кнопкой «Составить расписание»
  async function botToAbout(perGrade){
    for(var g = 5; g <= 11; g++){ for(var k = 1; k < perGrade; k++) bot('[data-bot-act="cnt"][data-g="' + g + '"][data-v="1"]'); }
    change($('#bot [data-bot-act="mode-per"][data-r="j"]'), "5");
    T("bot-next").click();
  }
  async function botSchool(perGrade){
    await botToAbout(perGrade);
    T("bot-next").click(); T("bot-next").click();
  }
  function st(){ return JSON.parse(W.localStorage.getItem("timetable_builder_v4")).state; }
  function st_safe(){ try{ return D.querySelectorAll('[data-act="parallel-add"]').length; }catch(e){ return "?"; } }
  function cls(name){ return st().classes.filter(function(c){ return c.name === name; })[0]; }
  function $(sel){ var el = D.querySelector(sel); if(!el) throw new Error("нет элемента " + sel); return el; }
  // 0.12: элементы, которые проверяют тесты, ищутся по метке data-test (Процесс_разработки.md, п. 3) — смена вида не ломает тесты
  function T(name){ var el = D.querySelector('[data-test="' + name + '"]'); if(!el) throw new Error("нет метки " + name); return el; }
  function click(sel){ $(sel).click(); }
  function change(el, v){ el.value = v; el.dispatchEvent(new W.Event("change", { bubbles: true })); }
  // 0.12 (Э1): пункт меню конструктора по метке; nav() без аргумента — открытый пункт
  function nav(id){ if(id){ T("nav-" + id).click(); return id; } var el = D.querySelector('#side [aria-current="page"]'); return el ? el.getAttribute("data-go") : ""; }
  // старые номера вкладок → экраны 0.12: 0 классы, 1 предметы, 2 учителя, 3 кабинеты, 4 план класса, 5 сетка, 6 печать и Excel
  function tab(i){
    if(i === 4){ nav("load"); click('[data-act="load-mode"][data-id="class"]'); }
    else nav(["classes", "subjects", "teachers", "rooms", "", "grid", "print"][i]);
  }
  function bot(sel){ var el = T("bot").querySelector(sel); if(!el) throw new Error("нет элемента помощника " + sel); el.click(); }
  // 0.12 (Э9): загрузка из eMaktab — пункт меню «Загрузить из eMaktab»; уже загружено — «Загрузить заново»
  function openEmk(){ nav("upload"); var a = D.querySelector('[data-test="upl-again"]'); if(a) a.click(); }
  function botGen(){ bot('[data-bot-act="gen"]'); if(!T("dialog").hidden) click('[data-dialog="all"]'); }
  function botQ(){ var q = D.querySelector('#bot [data-test="bot-q"]'); return q ? q.textContent : ""; }
  function botTrack(){ return D.querySelectorAll('#bot [data-test="bot-track"] .tstep').length; }
  // 0.12 (Э7): «Сетка». gen — составить кнопкой экрана: вопрос о ручных правках → «Перестроить всё», итог → «Смотреть сетку»
  function closeDlg(){ if(!T("dialog").hidden && D.querySelector('[data-dialog="ok"]')) click('[data-dialog="ok"]'); }
  function gen(){
    nav("grid"); click('[data-act="generate"]');
    if(!T("dialog").hidden && D.querySelector('[data-dialog="all"]')) click('[data-dialog="all"]');
    closeDlg();
  }
  // класс в сетке — через «Другой класс ▾» (любая параллель)
  function gridClass(name){ if(nav() !== "grid") nav("grid"); change(T("grid-other"), cls(name).id); }
  function cellTd(d, p){ return T("cell-" + d + "-" + p); }
  function tileV(name){ return T("tile-" + name).querySelector('[data-test="tile-value"]').textContent; }
  // группа экрана «Проверка»: число в заголовке и тексты пунктов
  function grpN(g){ var el = D.querySelector('[data-test="check-group-' + g + '"]'); return el ? el.querySelector('[data-test="check-count"]').textContent : "—"; }
  function openCell(d, p){ cellTd(d, p).querySelector('[data-act="cell-open"]').click(); return T("grid-pop"); }
  // n-я ячейка с уроком (unlocked — только незакреплённые): { d, p, td }
  function lessonCell(n, unlocked){
    var a = D.querySelectorAll('[data-test="grid-table"] td[data-kind="lesson"]' + (unlocked ? ':not(.locked-cell)' : ''))[n || 0];
    var m = a.getAttribute("data-test").split("-");
    return { d: +m[1], p: +m[2], td: a };
  }
  function tileVal(name){ return T("tile-" + name).querySelector('[data-test="tile-value"]').textContent; }
  // 0.12 (Э3): правка класса — в панели справа; openClass открывает «Классы», параллель и панель класса
  function openClass(name){
    nav("classes");
    var c = cls(name), p = D.querySelector('[data-test="panel-class"]');
    if(p && p.getAttribute("data-id") === c.id) return;
    T("cls-row-" + name).click();
  }
  function openPar(g){ nav("classes"); T("cls-par-edit-" + g).click(); }
  function classSel(act, name){ openClass(name); return $('[data-act="' + act + '"][data-id="' + cls(name).id + '"]'); }
  function setClass(act, name, v){ change(classSel(act, name), v); }
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
  // бот до шага «Составить»: тип школы, по одному классу в параллели; plan "no" — планы стёрты
  async function botToGen(type, plan){
    if(type === "classical") bot('[data-bot-act="type"][data-v="classical"]');
    await botSchool(1);
    if(plan === "no"){
      await setState(function(s){ Object.keys(s.curriculum).forEach(function(k){ s.curriculum[k] = {}; }); });
      T("mode-bot").click();
    }
  }

  // ячейки класса с предметом (по имени) в сохранённом расписании: [{d, p}]
  function subjCells(name, subjRe){
    var s = st(), c = cls(name), out = [], g = s.schedule.byClass[c.id] || [];
    g.forEach(function(row, d){ row.forEach(function(cell, p){ if(cell && cell.lessons.some(function(l){ var x = s.subjects.filter(function(y){ return y.id === l.subjectId; })[0]; return x && subjRe.test(x.name); })) out.push({ d: d, p: p }); }); });
    return out;
  }
  var scenarios = [
    /* ---------- 0.12 ---------- */
    ["T-U-01", "0.12 Э0: стиль Б, шапка, общие детали интерфейса", async function(){
      await fresh();
      var top = W.getComputedStyle(T("topbar"));
      check("T-U-01", "шапка тёмно-синяя #1B2B4B (П-8), страница #F4F7FB", top.backgroundColor === "rgb(27, 43, 75)" && W.getComputedStyle(D.body).backgroundColor === "rgb(244, 247, 251)", top.backgroundColor);
      check("T-U-01", "в шапке нет плашки с названием школы (П-9)", !/Школа\s*№/.test(T("topbar").textContent), T("topbar").textContent.replace(/\s+/g, " "));
      check("T-U-01", "метки шапки на месте: помощник, конструктор, язык", !!T("mode-bot") && !!T("mode-work") && !!T("lang") && !!T("brand"));
      T("mode-work").click();
      var m = W.getComputedStyle(T("main"));
      check("T-U-01", "рабочая область — белая карточка с тенью", m.backgroundColor === "rgb(255, 255, 255)" && m.boxShadow !== "none", m.boxShadow);
      var U = W.__ttUi, box = D.createElement("div");
      check("T-U-02", "детали интерфейса доступны: page, table, panel, tiles, card, empty, crumb, footNav", !!U && ["page", "table", "panel", "tiles", "card", "empty", "crumb", "footNav"].every(function(k){ return typeof U[k] === "function"; }));
      box.innerHTML = U.page("Классы <b>", "Зачем", "", U.table(["Класс", { title: "Дни" }], [{ cells: ["5-A", "6"], act: "panel", id: "c1", sel: true }], { test: "classes" }), { test: "classes" });
      var row = box.querySelector('[data-test="tbl-classes"] tbody tr');
      check("T-U-02", "uiPage: заголовок экранирован, метка экрана; uiTable: строка с data-act/data-id и выделением",
        !!box.querySelector('[data-test="page-classes"]') && box.querySelector('[data-test="page-title"]').textContent === "Классы <b>" && !box.querySelector("h3 b")
        && row.getAttribute("data-act") === "panel" && row.getAttribute("data-id") === "c1" && row.classList.contains("sel") && row.cells.length === 2);
      box.innerHTML = U.table(["A", "B"], []);
      check("T-U-02", "uiTable: пустая таблица — «Пока пусто» на всю ширину", /Пока пусто/.test(box.textContent) && box.querySelector("td").colSpan === 2);
      box.innerHTML = U.panel("5-A", "<p>поля</p>", '<button data-act="x">Удалить</button>', { hint: "Подсказка" }) + U.tiles([{ kind: "w", icon: "clock-pause", value: 49, label: "окон", test: "gaps" }]) + U.empty("users", "Добавьте классы", "Текст", "<button>Создать</button>");
      check("T-U-02", "uiPanel (крестик закрывает), uiTiles (значение), uiEmpty (кнопки)",
        box.querySelector('[data-test="panel-close"]').getAttribute("data-act") === "panel-close" && !!box.querySelector(".pfoot button")
        && box.querySelector('[data-test="tile-gaps"] [data-test="tile-value"]').textContent === "49" && box.querySelector('[data-test="empty"] .empty-acts button').textContent === "Создать");
      var steps = [{ id: "school", label: "Школа", st: "done" }, { id: "classes", label: "Классы", st: "warn" }, { id: "teachers", label: "Учителя", st: "todo" }];
      box.innerHTML = U.crumb(steps, "classes") + U.footNav({ id: "school", label: "Школа" }, { id: "teachers", label: "Учителя" }, "go", true);
      check("T-U-03", "uiCrumb: «Шаг 2 из 3», отрезки: готово — зелёный, текущий — синий",
        box.querySelector('[data-test="crumb-text"]').textContent === "Шаг 2 из 3" && box.querySelector('[data-test="crumb-school"]').className === "d" && box.querySelector('[data-test="crumb-classes"]').className === "c" && box.querySelector('[data-test="crumb-teachers"]').className === "");
      check("T-U-03", "uiFootNav: «← Школа», «Дальше: Учителя →», при замечаниях — «Можно идти дальше…»",
        box.querySelector('[data-test="foot-prev"]').getAttribute("data-id") === "school" && /Дальше: Учителя/.test(box.querySelector('[data-test="foot-next"]').textContent) && !!box.querySelector('[data-test="foot-warn"]'));
      check("T-U-03", "uiCrumb вне шагов и uiFootNav без соседей — пусто", U.crumb(steps, "print") === "" && U.footNav(null, null, "go") === "");
    }],
    /* ---------- 0.11 ---------- */
    ["T-11-21", "0.11: помощник и конструктор — на разных экранах", async function(){
      await fresh(true);
      check("T-11-21", "новый пользователь — только помощник: конструктор скрыт", !T("bot").hidden && T("kon").hidden && T("mode-bot").classList.contains("on"));
      bot('[data-bot-act="path"][data-v="scratch"]');
      check("T-11-21", "после ответа — всё ещё помощник, шаг 1", !T("bot").hidden && botQ() === "Школа и классы", botQ());
      T("mode-work").click();
      check("T-11-21", "«Конструктор» — шаги и рабочая область, помощник скрыт, полоска «шаг 1 из 5»", T("bot").hidden && !T("kon").hidden && /Помощник: шаг 1 из 5/.test(T("mode-strip").textContent) && /^#\/konstruktor\//.test(W.location.hash), W.location.hash);
      click('#modeStrip [data-mode="bot"]');
      check("T-11-21", "«Вернуться в помощник»", !T("bot").hidden && T("kon").hidden && W.location.hash === "#/pomoshnik", W.location.hash);
      T("bot-to-work").click();
      check("T-11-21", "«Открыть конструктор →» из помощника", T("bot").hidden && !T("kon").hidden, nav());
      W.history.back(); await sleep(120);
      check("T-11-21", "«Назад» браузера возвращает в помощник", !T("bot").hidden);
      T("mode-work").click(); await reload();
      check("T-11-21", "после перезагрузки — тот же режим", !T("kon").hidden && T("bot").hidden);
      check("T-11-21", "событие screen_open", events.concat(W.__ttEvents || []).some(function(e){ return e.goal === "tt_screen_open"; }));
    }],
    ["T-6-30..33", "0.11 Б1: приказ №133 — план 1–4 и родственные языки", async function(){
      await fresh(); await botToGen("criterial", "yes");
      var p1 = planNames("1-A");
      check("T-6-30", "бот заполнил 1-A по приказу №133: 21 ч, «Чтение», «Наука (Science)», общая «Технология»", planHours("1-A") === 21 && p1.indexOf("Чтение") >= 0 && p1.indexOf("Наука (Science)") >= 0 && p1.indexOf("Технология") >= 0 && p1.indexOf("Технология (мальчики)") < 0, planHours("1-A") + " " + p1.join(", "));
      check("T-6-30", "2-A и 4-A — 24 ч; источник запомнен", planHours("2-A") === 24 && planHours("4-A") === 24 && (cls("1-A").appliedPlan || {}).ref === "order133@2026-04-10", JSON.stringify(cls("1-A").appliedPlan));
      setClass("class-lang", "3-A", "ru"); autofill("3-A");
      check("T-6-31", "3-A русский: узбекский 2 ч вместо русского, 24 ч", planHours("3-A") === 24 && planNames("3-A").indexOf("Государственный язык (узбекский)") >= 0 && planNames("3-A").indexOf("Русский язык") < 0, planNames("3-A").join(", "));
      check("T-6-31", "источник в плане класса — приказ №133", /Приказ №133 от 10\.04\.2026/.test(T("main").textContent));
      setClass("class-lang", "2-A", "kz"); setClass("class-qsample", "2-A", "1"); autofill("2-A");
      var pl2 = st().curriculum[cls("2-A").id], rc = st().subjects.filter(function(x){ return x.name === "Родной язык и чтение"; })[0];
      check("T-6-32", "2-A казахский, образец 1: «Родной язык и чтение» 6 ч, узбекский 2, русский 1 — 24 ч", !!rc && pl2[rc.id] && pl2[rc.id].hours === 6 && planHours("2-A") === 24 && rc.nameUz === "Ona tili va oʻqish savodxonligi", planNames("2-A").join(", "));
      setClass("class-lang", "8-A", "kz"); setClass("class-qsample", "8-A", "2"); autofill("8-A");
      check("T-6-32", "8-A казахский без направления — по приказу №133: 33 ч, есть родной язык, нет русского", planHours("8-A") === 33 && planNames("8-A").indexOf("Родной язык") >= 0 && planNames("8-A").indexOf("Русский язык") < 0, planHours("8-A") + " " + planNames("8-A").join(", "));
      setClass("class-lang", "4-A", "qr"); tab(4); click('[data-act="cur-select-class"][data-id="' + cls("4-A").id + '"]');
      check("T-6-33", "4-A каракалпакский — официального плана нет, подсказка", /каракалпакским языком обучения официального плана нет/.test(T("main").textContent));
    }],
    ["T-8-30..33", "0.11 А2, А7: правила школы — «Kelajak soati» Пн-1, математика 1–4, проверка", async function(){
      await fresh(); await botToAbout(1);
      check("T-8-30", "в помощнике на экране «О школе» — вопросы «Правила школы»", botQ() === "О школе" && /Правила школы/i.test(T("bot").textContent) && !!D.querySelector('[data-test="bot-rule-kelajak-yes"]') && !!D.querySelector('[data-test="bot-rule-pd-yes"]'));
      T("bot-rule-kelajak-yes").click();
      T("bot-rule-primaryMath-yes").click();
      T("bot-next").click(); T("bot-next").click();
      var r = st().config.rules || {};
      check("T-8-30", "ответы записаны", r.kelajak === true && r.primaryMath === true, JSON.stringify(r));
      nav("school");
      check("T-8-30", "«Школа и правила»: ответы помощника — те же переключатели; в планах нет «Kelajak soati» — кнопка добавить", T("rule-sw-kelajak").getAttribute("aria-checked") === "true" && T("rule-sw-primaryMath").getAttribute("aria-checked") === "true" && !!D.querySelector('[data-test="rules-add-kelajak"]'));
      T("rules-add-kelajak").click();
      nav("grid");
      var kh = st().subjects.filter(function(x){ return x.name === "Час будущего"; })[0];
      check("T-8-30", "«Час будущего» 1 ч добавлен во все классы", !!kh && st().classes.every(function(c){ var e = st().curriculum[c.id][kh.id]; return e && e.hours === 1; }));
      gen();
      var badK = st().classes.filter(function(c){ var a = subjCells(c.name, /^Час будущего$/); return !(a.length === 1 && a[0].d === 0 && a[0].p === 0); }).map(function(c){ return c.name; });
      check("T-8-31", "у всех классов «Kelajak soati» — Пн, 1-й урок", !badK.length, badK.join(", "));
      var badM = [];
      ["1-A", "2-A", "3-A", "4-A"].forEach(function(n){ subjCells(n, /^Математика$/).forEach(function(x){ if(x.p !== (x.d === 0 ? 1 : 0)) badM.push(n + ":" + x.d + "/" + x.p); }); });
      check("T-8-32", "математика 1–4: 1-й урок, в понедельник — 2-й", !badM.length && subjCells("1-A", /^Математика$/).length === 5, badM.join(", "));
      check("T-8-32", "5-A: математика не привязана к 1-му уроку (правило только для 1–4)", subjCells("5-A", /^Математика$/).length === 5);
      nav("check"); check("T-8-33", "проверка: нарушений правил школы нет", T("check-group-rules").querySelector('[data-test="check-count"]').textContent === "нет" && !T("check-group-rules").querySelector('[data-test="check-item"]'));
      // нарушение: «Kelajak soati» 5-A переставлен со вторника — проверка показывает рекомендацию
      await setState(function(s){ var g = s.schedule.byClass[cls("5-A").id], x = g[0][0]; g[0][0] = g[1][0]; g[1][0] = x; });
      nav("check");
      check("T-8-33", "перестановка «Kelajak soati» видна в группе «Правила школы»", /не первым уроком понедельника/.test(T("check-group-rules").textContent));
      T("check-group-rules").querySelector('[data-test="check-show"]').click();
      check("T-U-13", "правило → сетка 5-A, ячейка «Вт, 1-й урок» подсвечена", nav() === "grid" && /5-A/.test(T("grid-cls-info").textContent) && cellTd(1, 0).classList.contains("focus-cell") && /Kelajak soati/.test(T("check-strip-text").textContent));
      nav("school"); T("rule-sw-kelajak").click(); nav("check");
      check("T-8-30", "ответ «Нет» — правило снято, рекомендации нет", st().config.rules.kelajak === false && !/не первым уроком понедельника/.test(T("main").textContent));
    }],
    ["T-4-13..15", "0.11 А4, А8: дни профразвития, свой день учителя, закрытые уроки", async function(){
      await fresh(); await botToGen("criterial", "yes");
      await setState(function(s){ addTeachers(s, 1, null); });
      nav("school");
      check("T-4-13", "правило выключено — раскладки дней нет", !!D.querySelector('[data-test="pd-off"]') && !D.querySelector('[data-act="pd-day"]'));
      T("rule-sw-pd").click();
      check("T-4-13", "«Да» — пресет дней по группам предметов", st().config.rules.pd.on === true && D.querySelectorAll('[data-act="pd-day"]').length === 7 && T("pd-day-natural").value === "2");
      nav("grid");
      gen();
      var wedPhys = 0, thuMath = 0;
      ["7-A", "8-A", "9-A", "10-A", "11-A"].forEach(function(n){ wedPhys += subjCells(n, /^(Физика|Химия|Биология)/).filter(function(x){ return x.d === 2; }).length; });
      ["5-A", "6-A", "7-A"].forEach(function(n){ thuMath += subjCells(n, /^Математика$/).filter(function(x){ return x.d === 3; }).length; });
      check("T-4-13", "физика, химия, биология — не в среду; математика 5–7 — не в четверг", wedPhys === 0 && thuMath === 0, "ср: " + wedPhys + ", чт: " + thuMath);
      check("T-4-13", "всё размещено", (st().schedule.unplacedCount || 0) === 0, st().schedule.unplacedCount);
      nav("school"); change(T("pd-day-natural"), "0");
      check("T-4-13", "пресет можно поменять: естественные — понедельник", st().config.rules.pd.days.natural === 0);
      // свой день учителя и закрытые уроки — в панели учителя (0.12, Э4)
      var phys = st().subjects.filter(function(x){ return x.name === "Физика"; })[0].id, tid = st().curriculum[cls("8-A").id][phys].teacherId;
      nav("teachers"); T("teach-row-" + tid).click();
      check("T-4-14", "панель учителя — сетка день × урок и методический день «По предметам — Пн»", !!T("teach-avail") && /По предметам — Пн/.test(T("teach-pd-auto-day").textContent) && T("teach-pd-auto").getAttribute("aria-pressed") === "true");
      T("teach-pd--1").click();
      check("T-4-14", "свой день: «нет»", st().teachers.filter(function(x){ return x.id === tid; })[0].pdDay === -1 && T("teach-pd--1").classList.contains("on"));
      T("avail-slot-1-0").click();
      T("avail-day-4").click();
      var tt = st().teachers.filter(function(x){ return x.id === tid; })[0];
      check("T-4-15", "закрыт урок «Вт, 1-й» и выходной в пятницу", tt.offSlots.indexOf("1|0") >= 0 && tt.daysOff[4] === true && T("avail-slot-1-0").classList.contains("off") && T("avail-slot-4-2").disabled);
      check("T-4-15", "в строке учителя «Не может»: Пт · 1 урок закрыт", /Пт · 1 урок закрыт/.test(T("teach-row-" + tid).textContent), T("teach-row-" + tid).textContent);
      nav("grid");
      gen();
      var s = st(), bad = 0;
      s.classes.forEach(function(c){ var off = 0; (s.schedule.byClass[c.id] || []).forEach(function(row, d){ row.forEach(function(cell, p){ if(cell && cell.lessons.some(function(l){ return l.teacherId === tid; }) && (d === 4 || (d === 1 && off + p === 0))) bad++; }); }); });
      check("T-4-15", "после составления у учителя нет уроков в пятницу и во вторник 1-м уроком", bad === 0, bad);
      // проверка видит урок в закрытом уроке
      await setState(function(s2){ var t2 = s2.teachers.filter(function(x){ return x.id === tid; })[0]; t2.daysOff[4] = false; t2.offSlots = []; var g = s2.schedule.byClass[cls("8-A").id]; for(var d=0; d<g.length; d++) for(var p=0; p<g[d].length; p++){ var cell = g[d][p]; if(cell && cell.lessons.some(function(l){ return l.teacherId === tid; })){ t2.offSlots = [d + "|" + p]; return; } } });
      nav("check");
      check("T-4-15", "проверка: урок, который у учителя закрыт", /урок закрыт/.test(T("check-group-avail").textContent));
      check("T-4-13", "событие rule_answer", events.concat(W.__ttEvents || []).some(function(e){ return e.goal === "tt_rule_answer"; }));
    }],
    ["T-E-10..12", "0.11 В1: перенос в eMaktab — список изменений, замена педагога, отметка «перенесено»", async function(){
      await fresh();
      openEmk();
      await emkUpload([await fixture("school1_grid.xls"), await fixture("school1_load.xls")]);
      click('[data-emk-act="apply"]');
      nav("transfer");
      check("T-E-10", "экран «Перенос в eMaktab» — изменений нет", !!T("page-transfer") && !!T("empty-transfer-none") && tileV("tr-classes") === "0");
      // правки: у 5-A убран урок Пн 1; у 6-A у английского (весь предмет) — другой учитель
      var info = {};
      await setState(function(s){
        var c5 = s.classes.filter(function(c){ return c.name === "5-A"; })[0], g5 = s.schedule.byClass[c5.id];
        for(var p=0; p<g5[0].length; p++){ if(g5[0][p]){ info.del = p; g5[0][p] = null; break; } }
        var c6 = s.classes.filter(function(c){ return c.name === "6-A"; })[0], g6 = s.schedule.byClass[c6.id], plan6 = s.curriculum[c6.id];
        var sid = Object.keys(plan6).filter(function(k){ var x = s.subjects.filter(function(y){ return y.id === k; })[0]; return x && x.name === "Математика"; })[0];
        var other = s.teachers.filter(function(t){ return t.id !== plan6[sid].teacherId; })[0];
        info.to = other.name;
        g6.forEach(function(row){ row.forEach(function(cell){ if(cell) cell.lessons.forEach(function(l){ if(l.subjectId === sid) l.teacherId = other.id; }); }); });
        plan6[sid].teacherId = other.id;
      });
      nav("transfer");
      function trRow(name){ return D.querySelector('[data-test="tr-row"][data-id="' + cls(name).id + '"]'); }
      trRow("5-A").click();
      var tr = T("card-tr-class").textContent, ch = T("card-tr-class").querySelector('[data-test="tr-change"]');
      check("T-E-10", "5-A — «убрать», Пн", /5-A/.test(tr) && !!ch && ch.getAttribute("data-kind") === "del" && /^Пн/.test(ch.textContent) && /убрать/.test(ch.textContent), tr.slice(0, 300));
      var rp = T("card-tr-repl").textContent;
      check("T-E-11", "6-A — одна «Замена педагога» по математике, без правки ячеек", D.querySelectorAll('[data-test="tr-repl-item"]').length === 1 && /Замена педагога: [^]*Математика \(Весь класс\), 6-A/.test(rp), rp.slice(0, 400));
      trRow("6-A").click();
      var ch6 = Array.prototype.map.call(T("card-tr-class").querySelectorAll('[data-test="tr-change"]'), function(x){ return x.getAttribute("data-kind"); }).join();
      check("T-E-11", "у 6-A в списке класса только замена педагога", ch6 === "repl", ch6);
      var tiles = ["tr-classes", "tr-cells", "tr-repl", "tr-done"].map(tileV).join(" | ");
      check("T-E-10", "плитки: 2 класса с изменениями, 1 ячейка, 1 замена педагога, 0 из 2", tiles === "2 | 1 | 1 | 0 из 2", tiles);
      trRow("5-A").click(); T("tr-scheme-btn").click();
      check("T-E-12", "схема класса в формате eMaktab, изменённая ячейка выделена", !!T("tr-scheme").querySelector("td.is-changed") && /Весь класс/.test(T("tr-scheme").textContent));
      T("tr-mark").click();
      check("T-E-12", "«перенесено» запомнено, отметка в списке", !!(st().emaktab.transferred || {})[cls("5-A").id] && tileV("tr-done") === "1 из 2" && trRow("5-A").querySelector('[data-test="tr-check"]').classList.contains("on") && /Перенесено/.test(T("tr-mark").textContent));
      var printed = null; W.print = function(){ printed = D.getElementById("print").innerHTML; };
      T("tr-print").click();
      check("T-E-12", "печать списка изменений", !!printed && /список изменений/.test(printed) && /☑ 5-A/.test(printed));
      // новая правка в 5-A — отметка снимается сама
      await setState(function(s){ var c5 = s.classes.filter(function(c){ return c.name === "5-A"; })[0], g5 = s.schedule.byClass[c5.id]; for(var p=g5[1].length-1; p>=0; p--){ if(g5[1][p]){ g5[1][p] = null; break; } } });
      nav("transfer");
      check("T-E-12", "после новой правки отметка снята", tileV("tr-done") === "0 из 2");
      var ev = events.concat(W.__ttEvents || []);
      check("T-E-10", "события transfer_list_opened и transfer_marked — только числа", ev.some(function(e){ return e.goal === "tt_transfer_list_opened" && e.params.cells >= 1; }) && ev.some(function(e){ return e.goal === "tt_transfer_marked" && e.params.done === 1; }));
    }],
    ["T-11-01", "Бот: 5 экранов — расписание без учителей, итог, переход в мастер", async function(){
      await fresh();
      check("T-11-01", "помощник открыт на шаге 1 из 5 «Школа и классы»", botQ() === "Школа и классы" && botTrack() === 5 && T("bot-step-school").classList.contains("on"), botQ());
      await botToGen("criterial", "yes");
      check("T-11-01", "шаг «План»: 4 цифры и список готовности", /Учебный план готов/.test(botQ()) && /Учителей нет/.test(T("bot-ready").textContent) && T("bot-kpi").children.length === 4 && T("bot-step-about").classList.contains("done"), botQ());
      botGen();
      var s = st();
      check("T-11-01", "расписание построено", !!s.schedule);
      check("T-8-17", "0 неразмещённых (групповая пара без учителя не распадается)", s.schedule.unplacedCount === 0, "неразмещённых: " + s.schedule.unplacedCount);
      check("T-11-01", "итог: 4 цифры, «Открыть расписание», «Скачать Excel»; в конструкторе — «Сетка»", /Расписание готово/.test(botQ()) && T("bot-kpi").children.length === 4 && !!T("bot-open-grid") && !!T("bot-print") && nav() === "grid", botQ());
      check("T-7-01", "без учителей — предупреждение, не ошибка", !D.querySelector('[data-test="grid-errors"]') && /Учителей нет/.test(T("grid-warnings").textContent));
      await reload();
      check("T-11-02", "после перезагрузки помощник помнит шаг (итог)", !T("bot").hidden && /Расписание готово/.test(botQ()));
    }],
    ["T-11-26", "0.12 Э10: «О школе» в пути «с нуля» — «Жасорат», направления, ответы в конструкторе", async function(){
      await fresh(); await botToAbout(1);
      check("T-11-26", "«О школе» — шаг 2, первый отмечен галочкой", botQ() === "О школе" && T("bot-step-school").classList.contains("done") && T("bot-step-about").classList.contains("on"));
      T("bot-jas-yes").click(); T("bot-jas-5-A").click();
      check("T-11-26", "«Жасорат» «Да» — выбор классов 5–6, 5-A отмечен", cls("5-A").jasorat === true && !cls("6-A").jasorat && T("bot-jas-5-A").classList.contains("on"));
      T("bot-diryes-yes").click();
      var sel = T("bot-dir-sel-10-A"), it = Array.prototype.filter.call(sel.options, function(o){ return /Информационные/.test(o.text); })[0];
      change(sel, it.value);
      check("T-11-26", "направление 10-A — «Информационные технологии»", !!cls("10-A").variativ && T("bot-dir-sel-10-A").selectedOptions[0].text === it.text, cls("10-A").variativ);
      T("bot-next").click(); T("bot-next").click();
      check("T-11-26", "дальше — «Учителя» и «План»", /Учебный план готов/.test(botQ()));
      nav("school");
      check("T-11-26", "«Школа и правила»: 5-A — «Жасорат»", T("jasorat-5-A").checked && !T("jasorat-6-A").checked);
      T("mode-bot").click(); T("bot-back").click(); T("bot-back").click();
      T("bot-jas-no").click(); T("bot-diryes-no").click();
      check("T-11-26", "«Нет» снимает «Жасорат» и направления", !cls("5-A").jasorat && !cls("10-A").variativ && !T("bot-jas-no").parentNode.parentNode.querySelector(".sub"));
    }],
    ["T-11-27", "0.12 Э10: «О школе» в пути «проверить» — направление по догадке, ответ сохраняется", async function(){
      await fresh();
      setClass("class-direction", "10-A", "t:nd1");
      autofill("10-A");
      await setState(function(s){
        var c = s.classes.filter(function(x){ return x.name === "10-A"; })[0]; c.track = "none"; delete c.directionAsked;
        s.schedule = { byClass: {}, unplaced: [], unplacedCount: 0, source: "emaktab" };
        W.localStorage.setItem("tt_bot", JSON.stringify({ v: 2, step: 1, open: true, path: "import" }));
        W.localStorage.setItem("tt_screen", "bot");
      });
      var q = D.querySelector('[data-test="bot-q-dir-10-A"]');
      check("T-11-27", "вопрос «10-A — направление «Бизнес и профессии»?» с предметами из нагрузки", botQ() === "О школе" && !!q && /Бизнес и профессии/.test(q.textContent) && /Алгебра Экономика/.test(q.textContent), q ? q.textContent.slice(0, 150) : botQ());
      T("bot-dir-10-A-yes").click();
      check("T-11-27", "«Да» — направление записано в класс, как «Поставить» на «Проверке»", cls("10-A").track === "nd1" && cls("10-A").directionAsked === true);
      await reload();
      check("T-11-27", "после перезагрузки вопрос на месте, ответ «Да» отмечен", T("bot-dir-10-A-yes").classList.contains("on"));
      T("bot-dir-10-A-no").click();
      check("T-11-27", "«Нет» — без направления, больше не спрашивается на «Проверке»", cls("10-A").track === "none" && cls("10-A").directionAsked === true && T("bot-dir-10-A-no").classList.contains("on"));
    }],
    ["T-11-11", "Бот: план не заполнен — ошибка на шаге «Составить», переход к месту, составление", async function(){
      await fresh(); await botToGen("criterial", "no");
      check("T-11-11", "бот показывает ошибку, «Составить» недоступна", /исправьте/.test(botQ()) && $('#bot [data-bot-act="gen"]').disabled, botQ());
      bot('[data-bot-act="issue"]');
      check("T-11-11", "открыт план класса", /Нагрузка/.test(T("page-title").textContent) && !!D.querySelector("#cur-add-subj-select"));
      var sel = $("#cur-add-subj-select");
      sel.value = Array.prototype.filter.call(sel.options, function(o){ return /^Математика/.test(o.text); })[0].value;
      click('[data-act="cur-add-subject"]');
      change($('[data-act="cur-hours"]'), "5");
      check("T-11-11", "после исправления «Составить» доступна", !$('#bot [data-bot-act="gen"]').disabled);
      botGen();
      check("T-11-11", "расписание составлено", !!st().schedule && /Расписание готово/.test(botQ()));
      T("lang").click();
      var bad = cyr(T("bot").innerText, /^\d+-[A-Z]/);
      check("T-11-13", "бот на узбекском без кириллицы", bad.length === 0, bad.slice(0, 5).join(" "));
      T("lang").click();
    }],
    ["T-E-01", "Пилот: загрузка из eMaktab — сетка и нагрузка (обезличенная школа 1), предпросмотр, сопоставление предметов", async function(){
      await fresh();
      openEmk();
      check("T-E-01", "экран «Загрузить из eMaktab» из меню — сразу выбор файлов", T("page-title").textContent === "Загрузить из eMaktab" && !!D.querySelector("[data-emk-file]") && !D.querySelector('[data-test="upl-done"]'));
      await emkUpload([await fixture("school1_load.xls")]);
      check("T-E-01", "только нагрузка — просьба загрузить сетку", /Теперь нужна сетка расписания/.test(T("main").textContent) && !D.querySelector('[data-emk-act="apply"]'));
      await emkUpload([await fixture("school1_grid.xls")]);
      var tiles = Array.prototype.map.call(D.querySelectorAll(".emk-tiles .imp-tile"), function(x){ return x.textContent; }).join(" | ");
      check("T-E-01", "предпросмотр: 85 учителей, 50 классов, 57 кабинетов, 1731 урок", /85учителей/.test(tiles) && /50классов/.test(tiles) && /57кабинетов/.test(tiles) && /1731уроков в неделю/.test(tiles), tiles);
      var map = Array.prototype.map.call(D.querySelectorAll(".emk-map tbody tr"), function(tr){ var s = tr.querySelector("select"); return tr.cells[0].textContent + "→" + s.options[s.selectedIndex].text + "[" + tr.cells[3].textContent + "]"; });
      check("T-E-01", "предметы: полные названия из нагрузки, дословно ✓, синоним — «проверьте», нет в справочнике — новый",
        map.indexOf("Математика→Математика[✓]") >= 0 && map.indexOf("Физкультура→Физическое воспитание[проверьте]") >= 0 && map.indexOf("Кл. час→Час будущего[проверьте]") >= 0 && map.indexOf("Письмо→Родной язык[проверьте]") >= 0 && map.indexOf("Букварь→Чтение[проверьте]") >= 0 && map.indexOf("Английский язык→Английский язык[✓]") >= 0, map.slice(0, 12).join(" | "));
      var notesT = $(".emk-notes").textContent;
      check("T-E-01", "замечания простыми словами: урока нет на неделе — только классные часы; замены и разовые уроки скрыты", /Что стоит проверить в eMaktab/.test(notesT) && /загрузке это не мешает/.test(notesT) && /возможно, его забыли поставить: 4-D — Kelajak soati \(в eMaktab: «Ma'naviyat soati»\)/.test(notesT) && /правильное название — «Kelajak soati»/.test(notesT) && !/Английский язык \(/.test(notesT) && /уроки раз в две недели \(0,5 ч\) из учебной нагрузки не показаны: \d+/.test(notesT) && !/Экономика/.test(notesT.split("не показаны")[0]), notesT.slice(0, 300));
      check("T-E-01", "кнопка «Загрузить в конструктор» сразу под итогом", !!D.querySelector('.emk-go [data-emk-act="apply"]') && /Проверьте, правильно ли мы распознали предметы/.test(T("main").textContent));
      click('[data-emk-act="apply"]');
      var s = st(), n = 0;
      Object.keys(s.schedule.byClass).forEach(function(k){ s.schedule.byClass[k].forEach(function(r){ r.forEach(function(c){ if(c) n += c.lessons.length; }); }); });
      check("T-E-01", "созданы учителя, классы, кабинеты; в сетке все 1731 урок; источник — eMaktab", s.teachers.length === 85 && s.classes.length === 50 && s.rooms.length === 57 && n === 1731 && s.schedule.source === "emaktab" && s.emaktab.base.length === 1731, s.teachers.length + "/" + s.classes.length + "/" + s.rooms.length + "/" + n);
      check("T-E-01", "открыта «Сетка» с итогами плитками", nav() === "grid" && !!T("tiles-grid"));
      var pairs = 0, en = s.subjects.filter(function(x){ return x.name === "Английский язык (группа 2)"; })[0];
      Object.keys(s.curriculum).forEach(function(k){ var p = s.curriculum[k]; Object.keys(p).forEach(function(sid){ if(p[sid].syncWith && sid === en.id) pairs++; }); });
      check("T-E-01", "группы — связанные пары «предмет» + «предмет (группа 2)» с ярлыком eMaktab", pairs > 10, "пар английского: " + pairs);
      var t1 = s.teachers.filter(function(x){ return x.name === "Учитель001 А.Б."; })[0];
      check("T-E-01", "ФИО и имя eMaktab у учителя; у учителя нет лимита уроков в день", !!t1 && t1.ext.emaktab === "Учитель001 А.Б." && !t1.maxPerDay);
      nav("check");
      var tl = ["conflicts", "windows", "rooms", "rules"].map(function(k){ return k + ":" + tileV(k); }).join(" | ");
      check("T-E-03", "проверка: 0 накладок учителей; спортзал не считается конфликтом кабинета", tileV("conflicts") === "0" && tileV("rooms") === "0", tl);
      await reload();
      var n2 = 0; Object.keys(st().schedule.byClass).forEach(function(k){ st().schedule.byClass[k].forEach(function(r){ r.forEach(function(c){ if(c) n2 += c.lessons.length; }); }); });
      check("T-E-01", "после перезагрузки сетка на месте, ничего не убрано", n2 === 1731 && !st().schedule.stale, n2);
      var ev = (W.__ttEvents || []).concat(events).filter(function(e){ return e.goal === "tt_import_emaktab_result" && e.params.ok; })[0];
      check("T-E-01", "событие import_emaktab_result: только количества", !!ev && ev.params.teachers === 85 && ev.params.lessons === 1731 && ev.params.files === "grid,load");
    }],
    ["T-E-02", "Пилот: повторная загрузка — id и настройки сохраняются, сопоставление предметов запомнено, урок в выходной", async function(){
      await fresh(); openEmk();
      await emkUpload([await fixture("school1_grid.xls"), await fixture("school1_load.xls")]);
      // «Кл. час» сопоставляем с «Воспитанием» — должно запомниться
      var sel = Array.prototype.filter.call(D.querySelectorAll(".emk-map tbody tr"), function(tr){ return tr.cells[0].textContent === "Кл. час"; })[0].querySelector("select");
      change(sel, Array.prototype.filter.call(sel.options, function(o){ return o.text === "Воспитание"; })[0].value);
      click('[data-emk-act="apply"]');
      var s = st(), t1 = s.teachers.filter(function(x){ return x.name === "Учитель001 А.Б."; })[0], c5 = cls("5-A");
      await setState(function(x){ x.teachers.filter(function(t){ return t.id === t1.id; })[0].daysOff[0] = true; x.classes.filter(function(c){ return c.id === c5.id; })[0].studentCount = 25; });
      openEmk();
      await emkUpload([await fixture("school1_grid.xls"), await fixture("school1_load.xls")]);
      var row = Array.prototype.filter.call(D.querySelectorAll(".emk-map tbody tr"), function(tr){ return tr.cells[0].textContent === "Кл. час"; })[0];
      check("T-E-02", "сопоставление «Кл. час → Воспитание» запомнено", row && row.querySelector("select").options[row.querySelector("select").selectedIndex].text === "Воспитание" && /✓/.test(row.cells[3].textContent));
      check("T-E-02", "предупреждение о повторной загрузке", /Повторная загрузка/.test(T("main").textContent));
      click('[data-emk-act="apply"]');
      check("T-E-02", "подтверждение замены данных", !T("dialog").hidden && /Заменить данные школы/.test(T("dialog").textContent));
      click('[data-dialog="go"]');
      var s2 = st(), t2 = s2.teachers.filter(function(x){ return x.id === t1.id; })[0], c52 = cls("5-A");
      check("T-E-02", "без дублей; id учителя и класса прежние; выходной и число учеников сохранены", s2.teachers.length === 85 && s2.classes.length === 50 && !!t2 && t2.daysOff[0] === true && c52.id === c5.id && c52.studentCount === 25);
      nav("check");
      check("T-E-03", "проверка: урок в выходной учителя", /— выходной/.test(T("check-group-avail").textContent));
    }],
    ["T-E-04", "Пилот: неверные файлы — понятные сообщения, событие с причиной", async function(){
      await fresh(); openEmk();
      var staff = "<table><tr><td>№</td><td>ФИО</td><td>Должность</td></tr><tr><td>1</td><td>Каримова Дилноза</td><td>Учитель</td></tr></table>";
      await emkUpload([new W.File([staff], "sotrudniki.xls", { type: "text/html" })]);
      check("T-E-04", "отчёт «Сотрудники» — подсказка", /это отчёт «Сотрудники»/.test($("#main .warn-box").textContent), $("#main .warn-box").textContent);
      var wb = W.XLSX.utils.book_new(); W.XLSX.utils.book_append_sheet(wb, W.XLSX.utils.aoa_to_sheet([["Класс: 6-B"], ["1 четверть"], ["", "", "пн", "вт"]]), "S");
      await emkUpload([new W.File([W.XLSX.write(wb, { bookType: "biff8", type: "array" })], "6b.xls")]);
      check("T-E-04", "расписание одного класса — подсказка", /расписание одного класса/.test($("#main .warn-box").textContent));
      await emkUpload([new W.File(["просто текст"], "x.xls")]);
      check("T-E-04", "посторонний файл — подсказка", /не выгрузка расписания eMaktab/.test($("#main .warn-box").textContent));
      var bad = (W.__ttEvents || []).filter(function(e){ return e.goal === "tt_import_emaktab_result" && e.params.ok === false; }).map(function(e){ return e.params.reason; });
      check("T-E-04", "import_emaktab_result ok:false с причиной", bad.indexOf("staff") >= 0 && bad.indexOf("classweek") >= 0, bad.join());
    }],
    ["T-E-03", "Пилот: проверка — конфликт учителя, объединённый урок, кабинет, окно, переход к ячейке", async function(){
      await fresh();
      await setState(function(s){
        var math = s.subjects.filter(function(x){ return x.name === "Математика"; })[0].id, phys = s.subjects.filter(function(x){ return x.name === "Физика"; })[0].id;
        var a = s.classes.filter(function(c){ return c.name === "6-A"; })[0], b = JSON.parse(JSON.stringify(a)); b.id = "class6b"; b.name = "6-B"; s.classes.push(b);
        s.teachers = [{ id: "tT1", name: "Учитель Один", subjects: [], maxPerDay: 2, daysOff: [false,false,false,false,false,false] }, { id: "tT2", name: "Учитель Два", subjects: [], maxPerDay: null, daysOff: [false,false,false,false,false,false] }];
        s.rooms = [{ id: "rR1", name: "101", capacity: 30, type: null }, { id: "rR2", name: "102", capacity: 30, type: null }];
        s.curriculum[a.id] = {}; s.curriculum[a.id][math] = { hours: 5, teacherId: "tT1", blockSize: 1, syncWith: null };
        s.curriculum[b.id] = {}; s.curriculum[b.id][math] = { hours: 3, teacherId: "tT1", blockSize: 1, syncWith: null }; s.curriculum[b.id][phys] = { hours: 1, teacherId: "tT2", blockSize: 1, syncWith: null };
        function row(){ return [null,null,null,null,null,null]; }
        var A = [row(),row(),row(),row(),row(),row()], B = [row(),row(),row(),row(),row(),row()];
        function L(sid, t, r){ return { lessons: [{ subjectId: sid, teacherId: t, roomId: r }] }; }
        A[0][0] = L(math, "tT1", "rR1"); B[0][0] = L(math, "tT1", "rR1");     // объединённый урок: один учитель, один кабинет
        A[1][0] = L(math, "tT1", "rR1"); B[1][0] = L(phys, "tT2", "rR1");     // кабинет у двух классов
        A[2][1] = L(math, "tT1", "rR1"); B[2][1] = L(math, "tT1", "rR2");     // конфликт учителя
        A[3][0] = L(math, "tT1", "rR1"); A[3][2] = L(math, "tT1", "rR1");     // окно у 6-A во Чт (и 3 урока при максимуме 2 — нет, 2)
        s.schedule = { byClass: {}, unplaced: [], unplacedCount: 0 }; s.schedule.byClass[a.id] = A; s.schedule.byClass[b.id] = B;
      });
      nav("check");
      check("T-E-03", "плитки: 1 накладка учителя, 1 кабинет, окно", tileV("conflicts") === "1" && tileV("rooms") === "1" && +tileV("windows") > 0, ["conflicts", "windows", "rooms"].map(tileV).join("/"));
      check("T-E-03", "группы по видам проблем: накладки, объединённый урок, окна, кабинеты", grpN("teacher") === "1" && grpN("combined") === "1" && grpN("rooms") === "1" && +grpN("windows") > 0, ["teacher", "combined", "windows", "rooms"].map(grpN).join("/"));
      function show(g, re){ var it = Array.prototype.filter.call(T("check-group-" + g).querySelectorAll('[data-test="check-item"]'), function(x){ return !re || re.test(x.textContent); })[0]; it.querySelector('[data-test="check-show"]').click(); }
      // Э8, П-4: «Показать» ведёт к месту проблемы
      show("teacher");
      check("T-U-13", "накладка учителя → «По учителям», строка учителя подсвечена", nav() === "grid" && !!T("tbl-grid-teachers") && T("grid-trow-tT1").classList.contains("hl"));
      // M — все проблемы, а не видимые пункты (в группе видно не больше 10)
      var total = +(/^Проблема 1 из (\d+)$/.exec(T("check-strip-n").textContent) || [])[1];
      check("T-U-13", "полоска «Проблема 1 из M» с текстом проблемы", total > 4 && /Учитель Один — Ср, 2-й урок: 6-A, 6-B/.test(T("check-strip-text").textContent) && !D.querySelector('[data-test="check-prev"]'), T("check-strip").textContent);
      T("check-next").click(); T("check-next").click();
      check("T-U-13", "«Следующая» дважды → окно 6-A: сетка класса, окно подсвечено", T("check-strip-n").textContent === "Проблема 3 из " + total && /6-A/.test(T("grid-cls-info").textContent) && cellTd(3, 1).classList.contains("focus-cell") && cellTd(3, 1).getAttribute("data-kind") === "gap");
      T("check-prev").click();
      check("T-U-13", "«Предыдущая» → объединённый урок, снова «По учителям»", T("check-strip-n").textContent === "Проблема 2 из " + total && !!T("tbl-grid-teachers"));
      T("check-back").click();
      check("T-U-13", "«Вернуться к проверке» — экран «Проверка», полоски нет", nav() === "check" && !D.querySelector('[data-test="check-strip"]'));
      var planItems = function(){ return T("check-group-plan").querySelectorAll('[data-test="check-item"]').length; }, nPlan = +grpN("plan"), shown = planItems();
      T("check-group-plan").querySelector('[data-test="check-more"]').click();
      check("T-U-13", "в группе видно 10 пунктов, «Показать все» открывает остальные", nPlan > 10 && shown === 10 && planItems() === nPlan, shown + " → " + planItems() + " из " + nPlan);
      show("rooms");
      var pop = T("grid-pop");
      check("T-U-13", "кабинет → ячейка 6-A Вт 1-й с открытой карточкой: кабинет занят, свободные кабинеты", /6-A/.test(T("grid-cls-info").textContent) && cellTd(1, 0).classList.contains("focus-cell") && /Занят: 6-B/.test(pop.textContent) && /102/.test(T("pop-room-free").textContent), pop.textContent);
      T("pop-clear").click();
      check("T-U-13", "урок убран — «Эта проблема решена», «Следующая» ведёт дальше", /Эта проблема решена/.test(T("check-strip").textContent) && !!T("check-next") && tileV("conflicts") === "1");
      nav("grid");
      check("T-U-13", "переход из меню убирает полоску", !D.querySelector('[data-test="check-strip"]'));
      nav("check"); show("plan", /^6-B/);
      var mth = st().subjects.filter(function(x){ return x.name === "Математика"; })[0].id;
      check("T-U-13", "сверка с планом → «Нагрузка», 6 класс, ячейка «6-B · Математика» открыта", nav() === "load" && T("load-par-6").getAttribute("aria-pressed") === "true" && T("load-cell-class6b-" + mth).classList.contains("on") && !!T("check-strip"));
    }],
    ["T-E-06", "Пилот: направление класса по предметам ЕСП — вопрос, «Поставить», сверка часов с планом", async function(){
      await fresh();
      setClass("class-direction", "10-A", "t:nd1");
      autofill("10-A");
      var names = planNames("10-A");
      check("T-E-06", "план направления «Бизнес и профессии» — с предметами «… Экономика»", names.indexOf("Алгебра Экономика") >= 0, names.join(", "));
      await setState(function(s){ var c = s.classes.filter(function(x){ return x.name === "10-A"; })[0]; c.track = "none"; delete c.directionAsked; s.schedule = { byClass: {}, unplaced: [], unplacedCount: 0 }; });
      nav("check");
      var dir = D.querySelector('[data-test="check-dir"]') ? T("check-dir").textContent : "";
      check("T-E-06", "вопрос: у 10-A похоже направление «Бизнес и профессии» — по предметам «… Экономика»", /10-A/.test(dir) && /Бизнес и профессии/.test(dir) && /Алгебра Экономика/.test(dir), dir.slice(0, 200));
      check("T-E-06", "пока направление не поставлено — расхождения с базовым планом", /10-A —/.test(planText()));
      click('[data-act="dir-set"][data-id="' + cls("10-A").id + '"]');
      check("T-E-06", "«Поставить» — направление записано в настройку класса, вопрос снят, план сходится", cls("10-A").track === "nd1" && !D.querySelector('[data-test="check-dir"]') && !/10-A —/.test(planText()), planText().slice(0, 200));
      await setState(function(s){
        var c = s.classes.filter(function(x){ return x.name === "10-A"; })[0], pl = s.curriculum[c.id];
        var ae = s.subjects.filter(function(x){ return x.name === "Алгебра Экономика"; })[0].id, a = s.subjects.filter(function(x){ return x.name === "Алгебра"; })[0].id;
        pl[a] = pl[ae]; delete pl[ae];
      });
      nav("check");
      check("T-E-06", "стоит обычная «Алгебра» — заметка «по направлению нужен «Алгебра Экономика»»", /10-A — стоит «Алгебра», а по направлению нужен «Алгебра Экономика»/.test(planText()), planText().slice(0, 300));
      await setState(function(s){ var c = s.classes.filter(function(x){ return x.name === "10-A"; })[0]; c.track = "none"; delete c.directionAsked; });
      nav("check"); click('[data-act="dir-no"][data-id="' + cls("10-A").id + '"]');
      check("T-E-06", "«Нет, без направления» — вопрос больше не задаётся", !D.querySelector('[data-test="check-dir"]') && cls("10-A").directionAsked === true && cls("10-A").track === "none");
    }],
    ["T-2-26", "Пилот: сквозная нумерация — старое сохранение со сменами, звонки школы, латиница в названиях", async function(){
      await fresh(); await botSchool(1);
      await setState(function(s){
        s.schemaVersion = 2;
        for(var g = 1; g <= 11; g++){ var P = s.config.parallels[g]; delete P.start; P.shift = g <= 4 ? 2 : 1; P.periods = g <= 4 ? 5 : 6; P.periodsByDay = null; }
        var c7 = s.classes.filter(function(c){ return c.name === "7-A"; })[0]; c7.shiftOverride = 2;
        s.classes.filter(function(c){ return c.name === "9-A"; })[0].name = "9-Б";
        s.config.bells = { 1: { start: "08:30", len: 45, brk: 10, bigAfter: 3, bigLen: 20 }, 2: { start: "13:30", len: 45, brk: 10, bigAfter: 3, bigLen: 20 } };
        s.schedule = null;
      });
      var s = st(), P = s.config.parallels, c7 = cls("7-A");
      check("T-2-26", "схема данных 3; смена 2 → с 7-го урока (после 6 уроков 1-й смены)", s.schemaVersion === 3 && P[1].start === 7 && P[5].start === 1 && c7.startOverride === 7 && P[1].shift === undefined && c7.shiftOverride === undefined, JSON.stringify([s.schemaVersion, P[1].start, P[5].start, c7.startOverride]));
      check("T-2-26", "звонки 1-й смены стали звонками школы", s.config.schoolBells && s.config.schoolBells.start === "08:30" && !s.config.bells);
      check("T-2-24", "«9-Б» кириллицей стал «9-B»", !!cls("9-B") && !cls("9-Б"));
      T("mode-bot").click(); botGen();
      gridClass("1-A");
      var th1 = T("grid-table").querySelector("tbody th").textContent;
      check("T-2-26", "сетка 1-A: уроки школы 7–11 и время по звонкам школы", /уроки 7–11/.test(T("grid-cls-info").textContent) && /^7/.test(th1) && /14:10–14:55/.test(th1), th1);
      check("T-2-26", "после миграции составление без конфликтов", st().schedule && !D.querySelector('[data-test="grid-table"] td[data-kind="cf"]'));
      openPar(6); change(T("par-start"), "4");
      check("T-2-26", "параллель 6 с 4-го урока — смены перекрываются", st().config.parallels[6].start === 4 && /Уроки школы:\s*1–12/.test(T("main").textContent.replace(/\s+/g, " ")), T("cls-lessons").textContent);
      gridClass("6-A");
      check("T-2-26", "сетка 6-A: уроки 4–9", /уроки 4–9/.test(T("grid-cls-info").textContent));
    }],
    ["T-11-20", "Пилот: бот — «Расписание уже есть в eMaktab?», загрузка, итог проверки, узбекский", async function(){
      await fresh(true);
      check("T-11-20", "первый вопрос помощника", /Расписание уже есть в eMaktab\?/.test(T("bot-card-start").textContent) && !!T("bot-path-import") && !!T("bot-path-scratch"));
      T("lang").click();
      var bad = cyr(T("bot").textContent, /^eMaktab/);
      check("T-11-13", "первый вопрос бота на узбекском без кириллицы", bad.length === 0, bad.slice(0, 5).join(" "));
      T("lang").click();
      bot('[data-bot-act="path"][data-v="import"]');
      check("T-11-20", "путь «проверить»: 3 шага — Загрузка, О школе, Результат", botTrack() === 3 && !!T("bot-step-emk") && !!T("bot-step-about") && !!T("bot-step-result") && /Загрузите данные из eMaktab/.test(botQ()));
      await emkUpload([await fixture("school1_grid.xls"), await fixture("school1_load.xls")], "#bot ");
      check("T-11-25", "П-1: в помощнике нет блока «Что стоит проверить в eMaktab»", !!D.querySelector('#bot [data-emk-act="apply"]') && !D.querySelector("#bot .emk-notes") && !/Что стоит проверить в eMaktab/.test(T("bot").textContent));
      click('#bot [data-emk-act="apply"]');
      check("T-11-25", "после загрузки — «О школе» (П-2): правила, направления по догадке", botQ() === "О школе" && !!T("bot-q-kelajak") && !!T("bot-q-pd") && !!T("bot-dir-note") && nav() === "grid", botQ());
      var dq = D.querySelector('#bot [data-test^="bot-q-dir-"]');
      if(dq){
        var dName = dq.getAttribute("data-test").slice(9);
        dq.querySelector('[data-test$="-yes"]').click();
        var dc = cls(dName);
        check("T-11-25", "«Да» на догадку — у класса направление (как в «Классах»)", !!(dc.variativ || (dc.track && dc.track !== "none")) && dc.directionAsked === true, dName + " " + dc.variativ + " " + dc.track);
        check("T-11-25", "ответ остаётся на экране", T("bot-dir-" + dName + "-yes").classList.contains("on"));
      } else check("T-11-25", "догадок о направлениях нет — пояснение", /направлений не видно/.test(T("bot-dir-note").textContent));
      T("bot-rule-pd-yes").click();
      check("T-11-25", "методические дни «Да» — раскладка по предметам и «Изменить»", /Ср — Физика, химия/.test(T("bot-pd-days").textContent) && !!T("bot-pd-edit"), T("bot-pd-days").textContent.slice(0, 120));
      T("bot-next").click();
      check("T-11-20", "итог проверки: 4 цифры, главные проблемы с «Показать», «Открыть расписание», «Перенести в eMaktab»", /Проверила ваше расписание/.test(botQ()) && T("bot-kpi").children.length === 4 && !!T("bot-open-grid") && !!T("bot-transfer"), botQ());
      var probs = D.querySelectorAll('#bot [data-test="bot-problem"]').length, kc = +T("bot-kpi-conflicts").querySelector("b").textContent;
      check("T-11-25", "не больше 3 главных проблем; накладки есть — первая про накладку", probs >= 1 && probs <= 3 && (!kc || T("bot-problems").querySelector(".ti-alert-octagon")), probs + " / " + kc);
      await reload();
      check("T-11-20", "после перезагрузки помощник на итоге проверки", !T("bot").hidden && /Проверила ваше расписание/.test(botQ()));
      T("bot-show").click();
      check("T-11-25", "«Показать →» ведёт как на «Проверке»: конструктор и полоска «Проблема 1 из N»", T("bot").hidden && !T("kon").hidden && /Проблема 1 из \d+/.test(T("check-strip-n").textContent), T("check-strip").textContent.slice(0, 80));
      T("check-back").click();
      check("T-11-25", "ответы помощника видны в «Школа и правила»: методические дни включены", (nav("school"), T("rule-sw-pd").getAttribute("aria-checked") === "true"));
    }],
    ["T-9-07", "«По учителям» открывает вид по учителям", async function(){
      await fresh(); await botToGen("criterial", "yes"); botGen();
      tab(2); T("teach-new").click(); T("teach-new-name").value = "Иванова"; T("teach-add").click();
      gen(); T("grid-by").click();
      check("T-9-07", "вид по учителям: кнопка «По классам», таблица учителей", /По классам/.test(T("grid-by").textContent) && !!T("tbl-grid-teachers"));
      check("T-9-07", "строка учителя", /Иванова/.test(T("tbl-grid-teachers").textContent));
      var srch = T("grid-t-search"); srch.value = "иванов"; srch.dispatchEvent(new W.Event("input", { bubbles: true }));
      var vis = Array.prototype.filter.call(T("tbl-grid-teachers").querySelectorAll("tbody tr"), function(r){ return !r.hidden; });
      check("T-9-07", "поиск учителя — одна строка", vis.length === 1 && /Иванова/.test(vis[0].textContent), vis.length);
      T("grid-by").click();
      check("T-9-07", "«По классам» возвращает сетку класса", !!T("grid-table") && /По учителям/.test(T("grid-by").textContent));
    }],
    ["T-9-05", "Закрепление урока: значок, закрепление, повторная генерация, открепление", async function(){
      await fresh(); await botToGen("criterial", "yes"); botGen();
      gridClass("5-A");
      var c0 = lessonCell(0), d = c0.d, p = c0.p;
      openCell(d, p);
      var btn = T("pop-lock");
      check("T-9-05", "значок — SVG, без эмодзи", !!btn.querySelector("svg") && !/[\u{1F512}\u{1F513}]/u.test(T("main").innerHTML));
      var before = JSON.stringify(st().schedule.byClass[cls("5-A").id][d][p].lessons);
      // клик по внутреннему элементу значка, как это делает мышь
      btn.querySelector("path").dispatchEvent(new W.MouseEvent("click", { bubbles: true }));
      check("T-9-05", "урок закреплён (клик по значку): замок в ячейке, урок в карточке не меняется", !!cellTd(d, p).querySelector('[data-test="cell-lock"]') && cellTd(d, p).classList.contains("locked-cell") && T("pop-lesson").disabled && st().schedule.byClass[cls("5-A").id][d][p].locked === true);
      gen();
      var after = st().schedule.byClass[cls("5-A").id][d][p];
      check("T-8-11", "после повторной генерации урок на месте и закреплён", after && after.locked && JSON.stringify(after.lessons) === before);
      openCell(d, p); T("pop-lock").click();
      check("T-9-05", "открепление возвращает выбор урока", !st().schedule.byClass[cls("5-A").id][d][p].locked && !T("pop-lesson").disabled && !cellTd(d, p).querySelector('[data-test="cell-lock"]'));
    }],
    ["T-9-11", "Цвета групп предметов", async function(){
      await fresh(); await botToGen("criterial", "yes"); botGen();
      gridClass("7-A");
      var les = Array.prototype.slice.call(D.querySelectorAll('[data-test="grid-table"] td[data-kind="lesson"] .cell'));
      var noColor = les.filter(function(c){ return !/cat-/.test(c.className); });
      check("T-9-11", "у каждого урока есть цвет группы", les.length > 0 && noColor.length === 0, noColor.length + " без цвета");
      check("T-9-11", "легенда из 7 групп", T("grid-legend-groups").querySelectorAll(".subj-legend > span").length === 7);
      check("T-9-11", "без учителя нет «— —»", !/— —/.test(T("main").textContent));
    }],
    ["T-6-19..22", "Планы приказов №227 и №183", async function(){
      await fresh(); await botToGen("criterial", "yes");
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
      check("T-6-22", "подсказка про Каракалпакстан", /Каракалпакстан/.test(T("main").textContent));
    }],
    ["T-2-02", "Классическая школа: направления №183 недоступны", async function(){
      await fresh(); await botToGen("criterial", "yes");
      setClass("class-direction", "10-A", "t:mf");
      nav("school"); T("school-type-classical").click(); tab(0);
      var sel = classSel("class-direction", "10-A");
      check("T-2-02", "нет группы №183", !Array.prototype.some.call(sel.querySelectorAll("optgroup"), function(g){ return /183/.test(g.label); }));
      check("T-2-02", "выбранное помечено «только для критериальной»", /критериальной/.test(sel.options[sel.selectedIndex].text));
      autofill("10-A");
      check("T-2-02", "подставлен базовый план", /Базовый учебный план/.test(T("main").textContent) && (cls("10-A").appliedPlan || {}).source === "base");
      nav("school"); T("school-type-criterial").click(); autofill("10-A");
      check("T-2-02", "обратно: план №183, 31 ч", planHours("10-A") === 31, planHours("10-A"));
    }],
    ["T-1-04", "Узбекский интерфейс без кириллицы", async function(){
      await fresh(); await botToGen("criterial", "yes");
      setClass("class-direction", "10-A", "t:nd1");
      T("lang").click();
      var allow = /^\d+-[A-Z]/, bad = [];
      [0, 4, 5, 6, 2, 1].forEach(function(i){ tab(i); bad = bad.concat(cyr(T("main").innerText, allow)); });
      nav("load"); click('[data-act="load-mode"][data-id="parallel"]');
      bad = bad.concat(cyr(T("main").innerText, allow), cyr(T("side").innerText, allow), cyr(T("bot").innerText, allow));
      // предметы — из справочника ЕСП, у них свои узбекские названия; пользовательских данных тут нет
      check("T-1-04", "шаги мастера, «Нагрузка», «Справочники», бот", bad.length === 0, bad.slice(0, 8).join(" "));
      T("lang").click();
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
      await fresh(); await botSchool(2);
      await setState(function(s){ addTeachers(s, 4, 6); s.schedule = null; });
      gen();
      var s = st(), w = windowsOf(s.schedule.byClass);
      check("T-8-26", "18 классов, 44 учителя: всё размещено", s.schedule.unplacedCount === 0, "неразмещённых: " + s.schedule.unplacedCount);
      check("T-8-26", "нет окон между уроками", w.windows === 0, "окон: " + w.windows);
      check("T-8-26", "каждый день начинается с 1-го урока", w.late === 0, "дней не с 1-го: " + w.late);
      check("T-8-26", "число окон в итогах составления", tileVal("windows") === "0", tileVal("windows"));
      var gr = events.concat(W.__ttEvents).filter(function(e){ return e.goal === "tt_generate_result"; }).pop();
      check("T-13-01", "generate_result: окна, конфликты, длительность", gr && gr.params.windows === 0 && gr.params.conflicts === 0 && gr.params.duration_ms >= 0, gr && JSON.stringify(gr.params));
    }],
    ["T-8-22", "P0-2: правки данных не стирают расписание, закреплённые и ручные правки остаются", async function(){
      await fresh(); await botToGen("criterial", "yes"); botGen();
      gridClass("7-A");
      var L1 = lessonCell(0), ld = L1.d, lp = L1.p;
      openCell(ld, lp); T("pop-lock").click();
      var before = JSON.stringify(cells("7-A")[ld][lp].lessons);
      // ручная правка: другой урок в незакреплённую ячейку — в карточке урока
      var L2 = lessonCell(0, true), md = L2.d, mp = L2.p;
      openCell(md, mp);
      var sel = T("pop-lesson");
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
      nav("grid");
      check("T-8-22", "плашка «расписание устарело»", /устарело/.test(T("grid-stale").textContent));
      // предмет убран из плана — его уроки убираются с сообщением
      tab(4); click('[data-act="cur-select-class"][data-id="' + cls("9-A").id + '"]');
      var rm = $('[data-act="cur-remove-subject"]'), rmSubj = rm.getAttribute("data-subj");
      var had = 0; cells("9-A").forEach(function(r){ r.forEach(function(c){ if(c && c.lessons.some(function(l){ return l.subjectId === rmSubj; })) had++; }); });
      rm.click();
      var left = 0; cells("9-A").forEach(function(r){ r.forEach(function(c){ if(c && c.lessons.some(function(l){ return l.subjectId === rmSubj; })) left++; }); });
      check("T-8-23", "уроки убранного предмета удалены, остальные на месте", had > 0 && left === 0 && filled(cells("9-A")) > 0, "было " + had + ", осталось " + left);
      check("T-8-23", "сообщение, сколько убрано", new RegExp("невозможными: " + had).test(toast()), toast());
      // «Перестроить»: вопрос о ручной правке; «Отмена» ничего не меняет
      nav("grid"); click('[data-act="generate"]');
      check("T-8-27", "перед перестроением — вопрос о ручных правках", !T("dialog").hidden && /без замка: 1/.test(T("dialog").textContent), T("dialog").textContent);
      click('[data-dialog="cancel"]');
      check("T-8-27", "«Отмена» — расписание не перестроено", st().schedule.stale === true);
      click('[data-act="generate"]'); click('[data-dialog="lock"]');
      s = st();
      var mc = s.schedule.byClass[cls("7-A").id][md][mp];
      check("T-8-27", "«Закрепить и перестроить»: ручная правка закреплена и на месте", !s.schedule.stale && mc && mc.locked && mc.lessons[0].subjectId === other.split("+")[0]);
      check("T-8-11", "закреплённый урок пережил перестроение", JSON.stringify(s.schedule.byClass[cls("7-A").id][ld][lp].lessons) === before);
      var inv = events.concat(W.__ttEvents).filter(function(e){ return e.goal === "tt_schedule_invalidated"; });
      check("T-13-01", "schedule_invalidated: причина и сохранённые правки", inv.length >= 2 && inv.some(function(e){ return e.params.reason === "plan" && e.params.locked === 1; }), inv.map(function(e){ return JSON.stringify(e.params); }).join(" "));
    }],
    ["T-10-10", "P0-3: выгрузка Excel при одинаковых и длинных названиях, классе после составления", async function(){
      await fresh(); await botToGen("criterial", "yes"); botGen();
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
      await fresh(); await botToGen("criterial", "yes");
      // перерисовка без изменений не пишет в хранилище
      var writes = 0, origSet = W.Storage.prototype.setItem;
      W.Storage.prototype.setItem = function(k, v){ if(k === "timetable_builder_v4") writes++; return origSet.call(this, k, v); };
      tab(1); tab(2); tab(3); tab(0);
      check("T-12-11", "переходы по шагам без правок — 0 записей", writes === 0, "записей: " + writes);
      // ошибка записи — плашка
      W.Storage.prototype.setItem = function(k){ if(k === "timetable_builder_v4"){ var e = new W.DOMException("full", "QuotaExceededError"); throw e; } };
      click('[data-act="parallel-add"][data-grade="5"]');
      check("T-12-12", "при переполнении и правке — «Браузер не сохраняет изменения»", /не сохраняет изменения/.test(T("storage-bar").textContent));
      W.Storage.prototype.setItem = origSet;
      click('[data-act="parallel-add"][data-grade="6"]');
      check("T-12-12", "после восстановления плашка уходит", !/не сохраняет/.test(T("storage-bar").textContent));
      // хранилище запрещено: пока правок нет — плашки нет (переходы по шагам ничего не пишут)
      await fresh();
      var gs = W.Storage.prototype.getItem, ss = W.Storage.prototype.setItem;
      W.Storage.prototype.getItem = W.Storage.prototype.setItem = function(){ throw new W.DOMException("denied", "SecurityError"); };
      // план класса не открываем: пустой план при показе заполняется по документу — это правка (с 0.11 и у 1-A)
      tab(1); tab(5); tab(0);
      check("T-12-15", "запрет хранилища без правок — плашки нет", !/не сохраняет/.test(T("storage-bar").textContent));
      click('[data-act="parallel-add"][data-grade="5"]');
      check("T-12-15", "после первой правки — спокойная плашка «сохраните копию»", /не сохраняет изменения/.test(T("storage-bar").textContent) && !!T("storage-bar").querySelector(".note-box"), T("storage-bar").textContent.slice(0, 120) + " | hash " + W.location.hash + " | 5-кл: " + st_safe());
      W.Storage.prototype.getItem = gs; W.Storage.prototype.setItem = ss;
      // повреждённое сохранение откладывается, а не затирается
      W.localStorage.setItem("timetable_builder_v4", "{повреждено");
      await reload();
      check("T-12-13", "повреждённое сохранение отложено в копию", W.localStorage.getItem("timetable_builder_v4_corrupt") === "{повреждено");
      check("T-12-13", "пользователь видит, что копия есть", /не удалось прочитать/.test(T("storage-bar").textContent));
      // вторая вкладка: эта становится «только для чтения»
      await fresh(); await botToGen("criterial", "yes");
      var saved = W.localStorage.getItem("timetable_builder_v4");
      localStorage.setItem("timetable_builder_tab", JSON.stringify({ id: "other-tab", ts: Date.now() }));
      await sleep(50);
      check("T-12-14", "другая вкладка → «только для чтения»", D.body.classList.contains("is-readonly") && /другой вкладке/.test(T("storage-bar").textContent));
      click('[data-act="parallel-add"][data-grade="5"]');
      check("T-12-14", "в режиме чтения работа другой вкладки не затирается", W.localStorage.getItem("timetable_builder_v4") === saved);
    }],
    ["T-6-25", "P0-6: групповые уроки — разные кабинеты, вместимость группы, один учитель, общий блок", async function(){
      await fresh(); await botToGen("criterial", "yes");
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
      check("T-6-25", "один учитель в двух группах — ошибка до составления", /один учитель/.test(T("page-grid").textContent) && $('[data-act="generate"]').disabled);
      await setState(function(s){ s.curriculum[s.classes.filter(function(x){ return x.name === "7-A"; })[0].id].subjG2.teacherId = "teachI2"; });
      gen();
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
      await fresh(); await botToGen("criterial", "yes"); botGen();
      tab(0);
      var g = classSel("class-grade", "9-A");
      check("T-2-21", "у класса с параллелью её можно изменить", g && g.value === "9");
      // закрепляем субботний урок у 8-A, затем делаем класс 5-дневным
      gridClass("8-A");
      var satTd = $('[data-test="grid-table"] td[data-kind="lesson"][data-day="5"]');
      satTd.querySelector('[data-act="cell-open"]').click(); T("pop-lock").click();
      var sat = filled([cells("8-A")[5]]);
      setClass("class-days", "8-A", "5");
      check("T-2-22", "суббота убрана, остальное на месте", cells("8-A").length === 5 && filled(cells("8-A")) > 0);
      check("T-2-22", "сообщение, сколько убрано", new RegExp("невозможными: " + sat).test(toast()), toast());
      setClass("class-grade", "9-A", "10");
      check("T-2-21", "параллель 9-A → 10 изменена без удаления класса", cls("9-A").grade === 10 && !!st().schedule);
      gen();
      check("T-2-22", "после смены дней и параллели составление работает", !st().schedule.stale && cells("8-A").length === 5);
      tab(6); var log = stubDownloads(); click('[data-act="export-xlsx"]'); await sleep(30);
      check("T-2-22", "и выгрузка работает", log.files.indexOf("raspisanie.xlsx") >= 0);
    }],
    ["T-13-06", "P0-5: каталог событий, воронка, без персональных данных", async function(){
      await fresh();
      // воронка одной сессией: бот → план → учителя → составить → выгрузить
      await botToGen("criterial", "yes");
      // с 0.11 у 1–4 классов тоже есть план (приказ №133) — учитель на предмет в каждой параллели, чтобы нагрузка была выполнима
      await setState(function(s){ addTeachers(s, 1, 7); });
      tab(4); click('[data-act="cur-select-class"][data-id="' + cls("7-A").id + '"]'); change($('[data-act="cur-hours"]'), $('[data-act="cur-hours"]').value);
      tab(0); change(classSel("class-students", "5-A"), "25");
      gen();
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
      si.files = dt2.files; si.dispatchEvent(new W.Event("change", { bubbles: true }));
      for(var wi = 0; wi < 40 && !D.querySelector('[data-imp-act="apply"]'); wi++) await sleep(25); // файл читается асинхронно
      click('[data-imp-act="apply"]');
      tab(0); click('[data-act="preset-create"]'); click('[data-act="parallel-add"][data-grade="3"]');
      await setState(function(s){ s.curriculum[s.classes[0].id] = {}; Object.keys(s.curriculum).forEach(function(k){ s.curriculum[k] = {}; }); });
      tab(5); click('[data-act="goto-issue"]');
      T("lang").click(); T("lang").click();
      W.dispatchEvent(new W.ErrorEvent("error", { message: "тест", lineno: 1 }));
      W.confirm = function(){ return true; }; T("menu-wipe").click();
      // бот: переход к месту ошибки
      await fresh(); await botToGen("criterial", "no"); bot('[data-bot-act="issue"]');
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
      await fresh(); await botSchool(2); botGen();
      nav("grid");
      var pills = Array.prototype.filter.call(T("grid-other").options, function(o){ return o.value; }).map(function(o){ return o.textContent; });
      var expected = ["1-A","2-A","3-A","4-A","5-A","5-B","6-A","6-B","7-A","7-B","8-A","8-B","9-A","9-B","10-A","10-B","11-A","11-B"];
      check("T-9-12", "классы по порядку: 1-е, 2-е … внутри — по букве", pills.join() === expected.join(), pills.join(" "));
      gridClass("5-A");
      var lc = lessonCell(0); openCell(lc.d, lc.p); T("pop-lock").click(); T("pop-close").click();
      var td = cellTd(lc.d, lc.p).querySelector(".cell");
      check("T-9-05", "закреплённая ячейка без синей рамки (карточка закрыта)", W.getComputedStyle(td).boxShadow === "none", W.getComputedStyle(td).boxShadow);
      T("grid-rebuild").click(); T("grid-unlock-all").click();
      check("T-9-12", "«Перестроить ▾» → «Открепить все» снимает все замки", !D.querySelector(".locked-cell") && !D.querySelector('[data-test="grid-unlock-all"]'));
      check("T-11-15", "в боте нет эмодзи", !/[\u{1F300}-\u{1FAFF}]/u.test(T("bot").innerHTML));
    }],
    ["T-13-03", "P0-5: код школы из ссылки", async function(){
      await fresh();
      frame.src = "../../index.html?school=sch_123&t=" + Date.now();
      await new Promise(function(r){ frame.onload = r; }); W = frame.contentWindow; D = W.document; await sleep(80);
      var open = W.__ttEvents.filter(function(e){ return e.goal === "tt_app_open"; })[0];
      check("T-13-03", "school в параметрах события", open && open.params.school === "sch_123", open && JSON.stringify(open.params));
      W.localStorage.removeItem("tt_school");
    }],

    /* ---------------- P1 ---------------- */
    ["T-U-07", "0.12 Э2: «Школа и правила» — одно место для типа школы, звонков, правил, методических дней, «Жасорат»", async function(){
      await fresh(); await botSchool(1);
      T("mode-work").click(); nav("school");
      check("T-U-07", "экран: карточки «Школа», «Звонки», «Правила школы», «Методические дни»", !!D.querySelector('[data-test="page-school"]') && !!D.querySelector('[data-test="card-school"]') && !!D.querySelector('[data-test="card-bells"]') && !!D.querySelector('[data-test="card-rules"]') && !!D.querySelector('[data-test="card-pd"]'));
      check("T-U-07", "вопросы ещё не заданы — все три помечены", D.querySelectorAll('[data-test^="rule-"].is-new').length === 3);
      T("school-type-classical").click();
      check("T-U-07", "тип школы: «Классическая» сохранён", st().config.schoolType === "classical" && T("school-type-classical").getAttribute("aria-pressed") === "true");
      T("school-type-criterial").click();
      check("T-U-07", "уроков в школе — по параллелям", /^1–\d+$/.test(T("school-slots").textContent), T("school-slots").textContent);
      var j = T("jasorat-5-A"); j.checked = true; j.dispatchEvent(new W.Event("change", { bubbles: true }));
      check("T-U-07", "«Жасорат» у 5-A — отметка на экране школы пишется в класс", cls("5-A").jasorat === true && T("jasorat-5-A").checked);
      T("rule-sw-kelajak").click();
      check("T-U-07", "переключатель правила: включено — ответ «да»", st().config.rules.kelajak === true && T("rule-sw-kelajak").getAttribute("aria-checked") === "true" && !T("rule-kelajak").classList.contains("is-new"));
      T("rule-sw-kelajak").click();
      check("T-U-07", "повторное нажатие — «нет»", st().config.rules.kelajak === false);
      T("bells-init").click();
      check("T-U-07", "звонки заданы на экране школы", !!st().config.schoolBells && !!D.querySelector('[data-test="bell-row-1"]'));
      T("bells-clear").click();
      check("T-U-07", "звонки убраны", !st().config.schoolBells && !!D.querySelector('[data-test="bells-init"]'));
      nav("classes");
      check("T-U-07", "на «Классах» нет типа школы и звонков", !D.querySelector('[data-act="school-type"]') && !D.querySelector('[data-act="bells-init"]') && !D.querySelector('[data-act="bell"]'));
      nav("grid");
      check("T-U-07", "на «Сетке» нет правил школы и закрытых уроков учителей (они — в «Учителях», Э4)", !D.querySelector('[data-act="rule-q"]') && !D.querySelector('[data-act="pd-day"]') && !D.querySelector('[data-test="avail-card"]') && !D.querySelector('[data-act="avail-slot"]'));
      // ответ в помощнике виден на экране школы и наоборот — одни и те же данные
      await setState(function(s){ s.config.rules = { kelajak: true, primaryMath: false, pd: { on: true, days: null } }; });
      nav("school");
      check("T-U-07", "ответы (как из помощника) — на переключателях и раскладка дней", T("rule-sw-kelajak").getAttribute("aria-checked") === "true" && T("rule-sw-primaryMath").getAttribute("aria-checked") === "false" && D.querySelectorAll('[data-act="pd-day"]').length === 7);
    }],
    ["T-U-08", "0.12 Э3: «Классы» — таблица по параллелям, настоящие значения (П-3), панель класса и параллели", async function(){
      await fresh(); await botSchool(1);
      // школа на 50 классов: классы 1–11 добавляются копиями «<параллель>-A»
      await setState(function(s){
        var L = "BCDEFGHIJ", k = 0;
        while(s.classes.length < 50){
          var g = (k % 11) + 1, base = s.classes.filter(function(c){ return c.grade === g; })[0], n = s.classes.filter(function(c){ return c.grade === g; }).length;
          var c = JSON.parse(JSON.stringify(base)); c.id = "clsT" + k; c.name = g + "-" + L[n - 1]; s.classes.push(c); s.curriculum[c.id] = {}; k++;
        }
      });
      T("mode-work").click(); nav("classes");
      var parRows = D.querySelectorAll('#main tr[data-par]'), shown = Array.prototype.filter.call(D.querySelectorAll('#main tr[data-name]'), function(r){ return !r.hidden; });
      check("T-U-08", "50 классов: 11 строк параллелей, классы свёрнуты", st().classes.length === 50 && parRows.length === 11 && shown.length === 0, parRows.length + " / " + shown.length);
      check("T-U-08", "страница не длиннее 2 экранов (≤ 1536 px)", T("main").scrollHeight <= 1536, T("main").scrollHeight + " px");
      check("T-U-08", "подзаголовок: «50 классов · 11 параллелей»", /50 классов · 11 параллелей/.test(T("page-sub").textContent), T("page-sub").textContent);
      T("cls-par-6").click();
      check("T-U-08", "нажатие на параллель раскрывает её классы", !T("cls-row-6-A").hidden && !T("cls-row-6-B").hidden && T("cls-row-5-A").hidden);
      openClass("6-B");
      check("T-U-08", "панель класса: название, параллель, язык, дни, с урока, уроков, учеников, на дому, «Нагрузка класса →», «Удалить»",
        ["cls-name", "cls-grade", "cls-lang", "cls-days", "cls-start", "cls-periods", "cls-students", "cls-homeschool", "cls-load", "cls-del"].every(function(k){ return !!D.querySelector('#kpanel [data-test="' + k + '"]'); })
        && T("panel-class").getAttribute("data-id") === cls("6-B").id && T("cls-row-6-B").classList.contains("sel") && T("kon").classList.contains("withpanel"));
      check("T-U-08", "как в параллели — пунктир, «как в параллели — 6»", T("cls-periods").classList.contains("inh") && /как в параллели — 6/.test(T("cls-periods").selectedOptions[0].text), T("cls-periods").selectedOptions[0].text);
      change(T("cls-periods"), "5");
      check("T-U-08", "своё значение: у класса — оранжевый тег «5 · своё», у параллели — «6-B: 5»",
        cls("6-B").periodsOverride === 5 && /5 · своё/.test(T("cls-row-6-B").textContent) && !!T("cls-row-6-B").querySelector(".tag.diff")
        && /6-B: 5/.test(T("cls-par-6").textContent) && !/своё/.test(T("cls-row-6-A").textContent) && !!T("cls-row-6-A").querySelector(".inh"));
      check("T-U-08", "в панели — «своё» и ссылка «как в параллели»", !T("cls-periods").classList.contains("inh") && !!T("cls-inherit-periods"));
      T("cls-inherit-periods").click();
      check("T-U-08", "«как в параллели» возвращает значение параллели", cls("6-B").periodsOverride === null && !/6-B: 5/.test(T("cls-par-6").textContent));
      setClass("class-days", "6-B", "5");
      check("T-U-08", "дни класса — своё значение в строке параллели", /6-B: 5/.test(T("cls-par-6").textContent) && cls("6-B").daysOverride === 5);
      openClass("10-A");
      check("T-U-08", "у 10-A в панели — направление", !!D.querySelector('#kpanel [data-act="class-direction"]'));
      T("panel-close").click();
      check("T-U-08", "крестик закрывает панель", T("kpanel").hidden && !T("kon").classList.contains("withpanel"));
      openPar(7); change(T("par-days"), "5");
      check("T-U-08", "панель параллели: дни 7 параллели → 5 у всех её классов", st().config.parallels[7].days === 5 && st().classes.filter(function(c){ return c.grade === 7; }).every(function(c){ return !c.daysOverride; }) && /^5/.test(T("cls-par-7").cells[1].textContent));
      T("cls-add").click(); T("add-grade-11").click();
      var c11 = st().classes.filter(function(c){ return c.grade === 11; });
      check("T-U-08", "«Добавить класс» → 11 класс: новый класс открыт в панели", st().classes.length === 51 && T("panel-class").getAttribute("data-id") === c11[c11.length - 1].id, T("panel-title").textContent);
      var srch = T("cls-search"); srch.value = "10-"; srch.dispatchEvent(new W.Event("input", { bubbles: true }));
      var vis = Array.prototype.filter.call(D.querySelectorAll('#main tr[data-name]'), function(r){ return !r.hidden; }).map(function(r){ return r.getAttribute("data-name"); });
      check("T-U-08", "поиск «10-»: видны только классы 10 параллели", vis.length > 0 && vis.every(function(n){ return /^10-/.test(n); }) && T("cls-par-5").hidden && !T("cls-par-10").hidden, vis.join(","));
      srch.value = ""; srch.dispatchEvent(new W.Event("input", { bubbles: true }));
      T("cls-collapse").click();
      openClass("9-A"); T("cls-load").click();
      check("T-U-08", "«Нагрузка класса →» — «Нагрузка», план этого класса", nav() === "load" && /9-A/.test(T("main").textContent) && T("kpanel").hidden);
      var n0 = st().classes.length; openClass("9-B"); T("cls-del").click(); click('[data-dialog="del"]');
      check("T-U-08", "«Удалить» — класс удалён, панель закрыта", st().classes.length === n0 - 1 && !cls("9-B") && T("kpanel").hidden);
      await setState(function(s){ s.classes = []; s.curriculum = {}; });
      nav("classes");
      check("T-U-08", "пустая школа: «Загрузить из eMaktab» и «Создать 1-A … 11-A»", !!D.querySelector('[data-test="empty-classes"] [data-act="kon-go"][data-id="upload"]') && !!T("cls-preset"));
      T("cls-preset").click();
      check("T-U-08", "«Создать 1-A … 11-A» — 11 классов", st().classes.length === 11 && !!D.querySelector('[data-test="tbl-classes"]'));
    }],
    ["T-U-09", "0.12 Э4: «Учителя» — таблица, панель учителя (методический день, когда не может, максимум в день)", async function(){
      await fresh(); await botToGen("criterial", "yes");
      // первый учитель ведёт весь план 5-A (≈ 30 ч); методические дни выключены
      await setState(function(s){
        addTeachers(s, 1, null);
        var c5 = s.classes.filter(function(c){ return c.name === "5-A"; })[0], pl = s.curriculum[c5.id];
        Object.keys(pl).forEach(function(k){ pl[k].teacherId = s.teachers[0].id; });
        s.config.rules = s.config.rules || {}; s.config.rules.pd = { on: false, days: null };
      });
      T("mode-work").click(); nav("teachers");
      var s0 = st(), t1 = s0.teachers[0], n0 = s0.teachers.length;
      var head = T("tbl-teachers").tHead.textContent;
      check("T-U-09", "таблица: строка на каждого учителя; колонки «Учитель, Предметы, Часов, Метод. день, Не может»",
        D.querySelectorAll('#main tr[data-teach-name]').length === n0 && ["Учитель", "Предметы", "Часов", "Метод. день", "Не может"].every(function(x){ return head.indexOf(x) >= 0; }) && /Учителя/.test(T("page-title").textContent), head);
      check("T-U-09", "методические дни выключены — в колонке «—»", T("teach-row-" + t1.id).cells[3].textContent === "—");
      T("teach-row-" + t1.id).click();
      check("T-U-09", "панель учителя: ФИО, предметы и часы, методический день, сетка, максимум, «Нагрузка», «Удалить»",
        ["teach-name", "teach-what", "teach-pd-off", "teach-avail", "teach-max", "teach-load", "teach-del"].every(function(k){ return !!D.querySelector('#kpanel [data-test="' + k + '"]'); })
        && T("panel-teacher").getAttribute("data-id") === t1.id && T("teach-row-" + t1.id).classList.contains("sel") && T("kon").classList.contains("withpanel"));
      // перегруз: максимум 1 урок в день — назначено больше, чем успеет
      change(T("teach-max"), "1");
      var t1s = st().teachers.filter(function(x){ return x.id === t1.id; })[0], ld = 0;
      Object.keys(st().curriculum).forEach(function(cid){ var pl = st().curriculum[cid]; Object.keys(pl).forEach(function(k){ if(pl[k].teacherId === t1.id && pl[k].hours > 0) ld += pl[k].hours; }); });
      check("T-U-09", "максимум 1 в день → часы оранжевым, «1 учитель перегружен»", t1s.maxPerDay === 1 && ld > 6 && !!T("teach-row-" + t1.id).querySelector('[data-test="teach-over"]') && /1 учитель перегружен/.test(T("teach-over-count").textContent), ld);
      change(T("teach-max"), "");
      check("T-U-09", "пустой максимум — без ограничения", st().teachers.filter(function(x){ return x.id === t1.id; })[0].maxPerDay === null && T("teach-max").placeholder === "без ограничения");
      // методические дни включены — день «по предметам», свой день, «по предметам» снова
      await setState(function(s){ s.config.rules = s.config.rules || {}; s.config.rules.pd = { on: true, days: null }; });
      nav("teachers"); T("teach-row-" + t1.id).click();
      check("T-U-09", "методические дни включены — кнопки дня, «по предметам» выбрано", T("teach-pd-auto").classList.contains("on") && D.querySelectorAll('#kpanel [data-act="pd-seg"]').length >= 7);
      T("teach-pd-3").click();
      check("T-U-09", "свой день «Чт» — в данных, в строке и в сетке (столбец Чт закрыт)", st().teachers.filter(function(x){ return x.id === t1.id; })[0].pdDay === 3 && T("teach-row-" + t1.id).cells[3].textContent === "Чт" && T("avail-slot-3-0").disabled && T("avail-day-3").classList.contains("pd"));
      T("teach-pd-auto").click();
      check("T-U-09", "«по предметам» — свой день снят", st().teachers.filter(function(x){ return x.id === t1.id; })[0].pdDay === undefined);
      // закрыть урок и открыть снова
      T("avail-slot-0-1").click(); T("avail-slot-0-1").click();
      check("T-U-09", "повторное нажатие открывает урок", (st().teachers.filter(function(x){ return x.id === t1.id; })[0].offSlots || []).length === 0 && !T("avail-slot-0-1").classList.contains("off"));
      change(T("teach-name"), "Алиева Нигора");
      check("T-U-09", "ФИО меняется в панели и в таблице", st().teachers.filter(function(x){ return x.id === t1.id; })[0].name === "Алиева Нигора" && /Алиева Нигора/.test(T("teach-row-" + t1.id).textContent) && T("panel-title").textContent === "Алиева Нигора");
      T("teach-load").click();
      check("T-U-09", "«Уроки учителя в «Нагрузке» →» — экран «Нагрузка»", nav() === "load" && T("kpanel").hidden);
      nav("teachers");
      var srch = T("teach-search"); srch.value = "алиева"; srch.dispatchEvent(new W.Event("input", { bubbles: true }));
      var vis = Array.prototype.filter.call(D.querySelectorAll('#main tr[data-teach-name]'), function(r){ return !r.hidden; });
      check("T-U-09", "поиск «алиева» — одна строка", vis.length === 1 && vis[0].getAttribute("data-id") === t1.id, vis.length);
      srch.value = ""; srch.dispatchEvent(new W.Event("input", { bubbles: true }));
      T("teach-new").click(); T("teach-new-name").value = "Новиков П."; T("teach-add").click();
      var tn = st().teachers.filter(function(x){ return x.name === "Новиков П."; })[0];
      check("T-U-09", "«Добавить учителя» — новый учитель открыт в панели", !!tn && st().teachers.length === n0 + 1 && T("panel-teacher").getAttribute("data-id") === tn.id);
      T("teach-row-" + t1.id).click(); T("teach-del").click(); click('[data-dialog="del"]');
      var hasA = Object.keys(st().curriculum).some(function(cid){ var pl = st().curriculum[cid]; return Object.keys(pl).some(function(k){ return pl[k].teacherId === t1.id; }); });
      check("T-U-09", "«Удалить» — учитель удалён, его уроки без учителя, панель закрыта", !st().teachers.some(function(x){ return x.id === t1.id; }) && !hasA && T("kpanel").hidden);
      nav("grid");
      check("T-U-09", "на «Сетке» нет блока закрытых уроков", !D.querySelector('[data-act="avail-slot"]') && !D.querySelector('[data-act="avail-day"]'));
      await setState(function(s){ s.teachers = []; Object.keys(s.curriculum).forEach(function(cid){ var pl = s.curriculum[cid]; Object.keys(pl).forEach(function(k){ pl[k].teacherId = null; }); }); });
      nav("teachers");
      check("T-U-09", "нет учителей — пустой экран: «Загрузить из eMaktab», «Добавить учителя», список сотрудников", !!D.querySelector('[data-test="empty-teachers"] [data-act="kon-go"][data-id="upload"]') && !!T("teach-new") && !!T("staff-input"));
    }],
    ["T-U-10", "0.12 Э5: «Нагрузка» — карточка ячейки (П-10), «в пустые классы» без замены назначенных (П-11)", async function(){
      await fresh(); await botSchool(3);
      var math;
      await setState(function(s){
        [["teachM", "Абдуллаева Нодира"], ["teachK", "Каримова Дилноза"], ["teachN", "Назарова Гульнора"], ["teachX1", "Холматов А."], ["teachX2", "Юсупов Б."], ["teachX3", "Эргашев В."], ["teachX4", "Шарипов Г."]].forEach(function(x){
          s.teachers.push({ id: x[0], name: x[1], subjects: [], maxPerDay: null, daysOff: [false,false,false,false,false,false] });
        });
        math = s.subjects.filter(function(x){ return x.name === "Математика"; })[0].id;
        var c = function(n){ return s.classes.filter(function(k){ return k.name === n; })[0].id; };
        s.curriculum[c("5-A")][math].teacherId = "teachM";
        s.curriculum[c("5-B")][math].teacherId = "teachK"; // сознательно назначенный — не должен замениться молча
      });
      T("mode-work").click(); nav("load"); T("load-par-5").click();
      var id5 = { A: cls("5-A").id, B: cls("5-B").id, C: cls("5-C").id }, cm = function(k){ return st().curriculum[id5[k]][math]; };
      check("T-U-10", "экран: заголовок «Нагрузка», параллели кнопками 5…11, таблица параллели, «Часы учителей»",
        /Нагрузка/.test(T("page-title").textContent) && !!T("load-par-11") && T("load-par-5").classList.contains("on") && !!T("load-table") && !!T("load-hours") && T("load-table").tHead.textContent.indexOf("5-C") >= 0);
      var fill = T("load-fill-row-" + math);
      check("T-U-10", "П-11: кнопка «Абдуллаева Нодира — в пустые классы» в строке с пустой ячейкой", /Абдуллаева Нодира — в пустые классы/.test(fill.textContent), fill.textContent);
      fill.click();
      check("T-U-10", "П-11: пустой 5-C получил Абдуллаеву, назначенная в 5-B Каримова не заменена", cm("C").teacherId === "teachM" && cm("B").teacherId === "teachK" && cm("A").teacherId === "teachM");
      check("T-U-10", "П-11: в строке без пустых ячеек кнопки нет", !D.querySelector('[data-test="load-fill-row-' + math + '"]'));
      T("load-cell-" + id5.B + "-" + math).click();
      var items = T("load-teachers").querySelectorAll(".lt-item[data-lt-name]");
      check("T-U-10", "карточка 5-B · Математика: сверху ведущие предмет, у каждого «назначено / можно ч»",
        /5-B · Математика/.test(T("panel-title").textContent) && ["teachK", "teachM"].indexOf(items[0].getAttribute("data-id")) >= 0 && ["teachK", "teachM"].indexOf(items[1].getAttribute("data-id")) >= 0
        && /\d+ \/ \d+ ч/.test(T("load-pick-teachK").textContent) && T("load-pick-teachK").classList.contains("on"), items[0].textContent + " | " + items[1].textContent);
      var srch = T("load-t-search"); srch.value = "кари"; srch.dispatchEvent(new W.Event("input", { bubbles: true }));
      var vis = Array.prototype.filter.call(T("load-teachers").querySelectorAll("[data-lt-name]"), function(b){ return !b.hidden; });
      check("T-U-10", "поиск «кари» — один учитель", vis.length === 1 && vis[0].getAttribute("data-id") === "teachK", vis.length);
      W.confirm = function(){ return true; };
      T("load-replace-all").click();
      var dlg = T("dialog").textContent;
      check("T-U-10", "П-11: замена во всех классах — с подтверждением и списком, кого заменит", !T("dialog").hidden && /5-A — Абдуллаева Нодира → Каримова Дилноза/.test(dlg) && /5-C/.test(dlg), dlg.slice(0, 200));
      click('[data-dialog="cancel"]');
      check("T-U-10", "«Отмена» — ничего не заменено", cm("A").teacherId === "teachM" && cm("C").teacherId === "teachM");
      T("load-replace-all").click(); click('[data-dialog="yes"]');
      check("T-U-10", "«Заменить» — Каримова во всех классах параллели", cm("A").teacherId === "teachK" && cm("B").teacherId === "teachK" && cm("C").teacherId === "teachK");
      // П-10: часы «− / +» и расхождение с приказом
      var h0 = cm("B").hours;
      T("load-h-plus").click();
      var cellB = T("load-cell-" + id5.B + "-" + math);
      check("T-U-10", "П-10: «+» — часы +1; расхождение с приказом — в ячейке «по приказу» и в карточке «Поставить / Оставить»",
        cm("B").hours === h0 + 1 && cellB.classList.contains("w") && /по приказу/.test(cellB.textContent) && /По приказу — /.test(T("load-order").textContent) && !!T("load-order-set") && !!T("load-order-keep"), cm("B").hours + " / " + h0);
      T("load-order-keep").click();
      check("T-U-10", "«Оставить» — часы не меняются, подсветка пропадает", cm("B").hours === h0 + 1 && !T("load-cell-" + id5.B + "-" + math).classList.contains("w") && !!T("load-order-kept"));
      T("load-h-plus").click();
      check("T-U-10", "снова изменили часы — расхождение снова видно", !!T("load-order"));
      T("load-order-set").click();
      check("T-U-10", "«Поставить N ч» — часы как в приказе", cm("B").hours === h0 && !D.querySelector('[data-test="load-order"]'), cm("B").hours);
      // «N предметов без учителя» — переход по пустым ячейкам
      var nEmpty = parseInt(T("load-empty").textContent, 10);
      T("load-empty").click();
      var p1 = T("panel-title").textContent, on1 = D.querySelector("#main .x.on");
      check("T-U-10", "«N предметов без учителя» — открывает пустую ячейку", nEmpty > 0 && !!on1 && on1.classList.contains("no") && !D.querySelector('[data-test="load-unpick"]'), nEmpty + " " + p1);
      T("load-empty").click();
      check("T-U-10", "второе нажатие — следующая пустая ячейка", T("panel-title").textContent !== p1, p1 + " → " + T("panel-title").textContent);
      T("load-pick-teachN").click();
      check("T-U-10", "выбор учителя в карточке назначает его; пустых меньше на 1", parseInt(T("load-empty").textContent, 10) === nEmpty - 1 && !!T("load-unpick"));
      T("load-unpick").click();
      check("T-U-10", "«Снять учителя» — снова без учителя", parseInt(T("load-empty").textContent, 10) === nEmpty);
      T("panel-close").click();
      check("T-U-10", "крестик закрывает карточку", T("kpanel").hidden);
      T("load-cell-" + id5.A + "-" + math).click(); T("load-plan").click();
      check("T-U-10", "«План класса 5-A →» — план класса, карточка закрыта", /Официальный план — 5-A/.test(T("main").textContent) && T("kpanel").hidden);
    }],
    ["T-U-11", "0.12 Э6: «Предметы» и «Кабинеты» — таблицы и правка в панели (всё, что было в «Справочниках»)", async function(){
      await fresh(); await botSchool(1);
      T("mode-work").click(); nav("subjects");
      var s0 = st(), n0 = s0.subjects.length, pe = s0.subjects.filter(function(x){ return x.name === "Физическое воспитание"; })[0];
      var head = T("tbl-subjects").tHead.textContent;
      check("T-U-11", "таблица предметов: строка на предмет; «Предмет, Группа, В eMaktab, Правила»; легенда групп",
        D.querySelectorAll('#main tr[data-subj-name]').length === n0 && ["Предмет", "Группа", "В eMaktab", "Правила"].every(function(x){ return head.indexOf(x) >= 0; })
        && /Предметы/.test(T("page-title").textContent) && D.querySelectorAll('#main .subj-legend > span').length === 7, head);
      var row = T("subj-row-" + pe.id);
      check("T-U-11", "цвет и группа предмета в строке", !!row.querySelector(".subj-dot.cat-pe") && /Физкультура/.test(row.cells[1].textContent), row.cells[1].textContent);
      row.click();
      check("T-U-11", "панель предмета: название, группа, тип кабинета, «не ставить», «всегда», «трудный», «Удалить»",
        ["subj-name", "subj-group", "subj-roomtype", "subj-np", "subj-fixed-day", "subj-heavy", "subj-del"].every(function(k){ return !!D.querySelector('#kpanel [data-test="' + k + '"]'); })
        && T("panel-subject").getAttribute("data-id") === pe.id && T("subj-row-" + pe.id).classList.contains("sel"));
      T("subj-np-1").click(); T("subj-np-7").click(); T("subj-np-7").click();
      function pS(){ return st().subjects.filter(function(x){ return x.id === pe.id; })[0]; }
      check("T-U-11", "«не ставить на уроки»: 1 — включён, 7 — нажат дважды и снят", JSON.stringify(pS().notPeriods) === "[1]" && T("subj-np-1").classList.contains("on") && !T("subj-np-7").classList.contains("on"), JSON.stringify(pS().notPeriods));
      change(T("subj-fixed-day"), "2"); change(T("subj-fixed-period"), "3");
      check("T-U-11", "«всегда Ср, 3-й урок» — в данных и в строке", pS().fixedSlot && pS().fixedSlot.day === 2 && pS().fixedSlot.period === 3 && /всегда Ср, 3-й урок/.test(T("subj-row-" + pe.id).textContent));
      change(T("subj-fixed-day"), "");
      check("T-U-11", "день «нет» — правило снято, урок недоступен", pS().fixedSlot === null && T("subj-fixed-period").disabled);
      change(T("subj-roomtype"), "спортзал"); T("subj-heavy").click();
      check("T-U-11", "тип кабинета и «трудный» сохраняются", pS().roomType === "спортзал" && pS().heavy === true);
      change(T("subj-name"), "  ");
      check("T-U-11", "пустое название не сохраняется", pS().name === "Физическое воспитание" && T("subj-name").value === "Физическое воспитание");
      change(T("subj-name"), "Физкультура");
      check("T-U-11", "название меняется в панели и в таблице", pS().name === "Физкультура" && /Физкультура/.test(T("subj-row-" + pe.id).cells[0].textContent) && T("panel-title").textContent === "Физкультура");
      var srch = T("subj-search"); srch.value = "физкул"; srch.dispatchEvent(new W.Event("input", { bubbles: true }));
      var vis = Array.prototype.filter.call(D.querySelectorAll('#main tr[data-subj-name]'), function(r){ return !r.hidden; });
      check("T-U-11", "поиск «физкул» — одна строка", vis.length === 1 && vis[0].getAttribute("data-id") === pe.id, vis.length);
      srch.value = ""; srch.dispatchEvent(new W.Event("input", { bubbles: true }));
      T("subj-new").click(); T("subj-new-name").value = "Шахматы"; T("subj-add").click();
      var sn = st().subjects.filter(function(x){ return x.name === "Шахматы"; })[0];
      check("T-U-11", "«Добавить предмет» — новый предмет открыт в панели", !!sn && st().subjects.length === n0 + 1 && T("panel-subject").getAttribute("data-id") === sn.id);
      T("subj-del").click(); click('[data-dialog="del"]');
      check("T-U-11", "«Удалить» — предмет удалён, панель закрыта", !st().subjects.some(function(x){ return x.name === "Шахматы"; }) && T("kpanel").hidden);
      check("T-U-11", "карточки «Правила предметов» над списком больше нет", !D.querySelector('[data-act="rule-add"]') && !D.querySelector("#rule-subj"));

      // «Кабинеты»: пустой экран, добавление, правка в панели, удаление
      nav("rooms");
      check("T-U-11", "нет кабинетов — пустой экран: «Загрузить из eMaktab», «Добавить кабинет»", !!D.querySelector('[data-test="empty-rooms"] [data-act="kon-go"][data-id="upload"]') && !!T("room-new"));
      T("room-new").click(); T("room-new-name").value = "214"; T("room-new-type").value = "спортзал"; T("room-new-capacity").value = "60"; T("room-add").click();
      var r1 = st().rooms[0];
      check("T-U-11", "«Добавить кабинет» — кабинет в данных, в таблице и открыт в панели", st().rooms.length === 1 && r1.name === "214" && r1.type === "спортзал" && r1.capacity === 60
        && !!T("room-row-" + r1.id) && T("panel-room").getAttribute("data-id") === r1.id);
      var rh = T("tbl-rooms").tHead.textContent;
      check("T-U-11", "таблица кабинетов: «Кабинет, Тип, Мест»; тип — тегом", ["Кабинет", "Тип", "Мест"].every(function(x){ return rh.indexOf(x) >= 0; }) && !!T("room-row-" + r1.id).querySelector(".tag"), rh);
      check("T-U-11", "панель: для каких предметов этот тип", /Физкультура/.test(T("room-subj").textContent), T("room-subj").textContent);
      change(T("room-capacity"), "35"); change(T("room-type"), ""); change(T("room-name"), "Спортзал");
      function rS(){ return st().rooms[0]; }
      check("T-U-11", "мест, тип, название меняются (пустой тип — обычный)", rS().capacity === 35 && rS().type === null && rS().name === "Спортзал" && /обычный/.test(T("room-row-" + r1.id).textContent));
      change(T("room-name"), "");
      check("T-U-11", "пустое название кабинета не сохраняется", rS().name === "Спортзал");
      T("room-del").click(); click('[data-dialog="del"]');
      check("T-U-11", "«Удалить» — кабинета нет, снова пустой экран", !st().rooms.length && !!D.querySelector('[data-test="empty-rooms"]') && T("kpanel").hidden);
    }],
    ["T-U-12", "0.12 Э7: «Сетка» — готовность (П-5), итог окном, плитки, ячейки Б1, карточка урока, «Перестроить ▾», 188 классов", async function(){
      await fresh(); await botSchool(1);
      T("mode-work").click(); nav("grid");
      check("T-U-12", "расписания нет — «Всё готово к составлению?», одна кнопка, «Перестроить» нет",
        !!T("card-grid-ready") && !!T("grid-build") && !T("grid-build").disabled && !D.querySelector('[data-test="grid-rebuild"]') && !D.querySelector('[data-test="grid-table"]'));
      check("T-U-12", "готовность: классы — 11 ✓, учителей нет, нагрузка — с переходом", /Классы — 11/.test(T("ready-classes").textContent) && T("ready-classes").classList.contains("r-ok") && T("ready-teachers").classList.contains("r-w") && !!T("ready-go-load"));
      T("ready-go-load").click();
      check("T-U-12", "«Назначить →» открывает «Нагрузку»", nav() === "load");
      await setState(function(s){ addTeachers(s, 1, null); });
      nav("grid");
      check("T-U-12", "учителя назначены — строки «Учителя» и «Нагрузка» зелёные", T("ready-load").classList.contains("r-ok") && T("ready-teachers").classList.contains("r-ok"));
      T("grid-build").click();
      check("T-U-12", "итог составления — окно с цифрами", !T("dialog").hidden && /Поставлено уроков: \d+/.test(T("dialog").textContent) && /Не размещено: 0/.test(T("dialog").textContent), T("dialog").textContent);
      click('[data-dialog="ok"]');
      check("T-U-12", "сетка: 4 плитки, «По учителям», «Перестроить ▾» закрыто, пункта этапа 2 нет",
        ["conflicts", "windows", "rules", "unplaced"].every(function(k){ return !!T("tile-" + k); }) && /По учителям/.test(T("grid-by").textContent) && T("grid-dd-menu").hidden && !/Улучшить/.test(T("grid-dd-menu").textContent));
      T("grid-rebuild").click();
      check("T-U-12", "«Перестроить ▾»: «Составить заново», «Очистить расписание»", !T("grid-dd-menu").hidden && !!T("grid-regen") && !!T("grid-clear"));
      T("page-title").click();
      check("T-U-12", "нажатие мимо меню закрывает его", T("grid-dd-menu").hidden);
      check("T-U-12", "проверки и сверки с планом на «Сетке» нет — ссылка на «Проверку»", !D.querySelector('[data-test^="check-group"]') && !D.querySelector('[data-test="grid-ready"]') && !!T("grid-to-check"));
      T("grid-to-check").click();
      check("T-U-12", "ссылка ведёт на «Проверку»", nav() === "check" && !!T("page-check"));
      gridClass("7-A");
      check("T-U-12", "класс 7-A: кнопка параллели нажата, в «Другой класс» — все 11 классов", T("grid-cls-7-A").getAttribute("aria-pressed") === "true" && Array.prototype.filter.call(T("grid-other").options, function(o){ return o.value; }).length === 11);
      // окно и накладка (Б1)
      await setState(function(s){
        var a = s.classes.filter(function(c){ return c.name === "7-A"; })[0], b = s.classes.filter(function(c){ return c.name === "8-A"; })[0];
        var ga = s.schedule.byClass[a.id], gb = s.schedule.byClass[b.id];
        ga[1][1] = null;                                                  // окно у 7-A: Вт, 2-й урок
        // накладка: учитель 8-A ведёт и урок 7-A в то же время (назначен и в плане 7-A)
        var la = ga[0][0].lessons[0], tB = gb[0][0].lessons[0].teacherId;
        la.teacherId = tB; s.curriculum[a.id][la.subjectId].teacherId = tB;
      });
      gridClass("7-A");
      var gap = cellTd(1, 1), cf = cellTd(0, 0), gs = W.getComputedStyle(gap.querySelector(".cell")), cs = W.getComputedStyle(cf.querySelector(".cell"));
      check("T-U-12", "окно — белое с оранжевой рамкой", gap.getAttribute("data-kind") === "gap" && gs.backgroundColor === "rgb(255, 255, 255)" && gs.borderTopColor === "rgb(232, 89, 12)", gs.backgroundColor + " " + gs.borderTopColor);
      check("T-U-12", "накладка — красная", cf.getAttribute("data-kind") === "cf" && cs.borderTopColor === "rgb(229, 72, 77)" && /накладка/.test(cf.textContent), cs.borderTopColor);
      check("T-U-12", "плитки: накладок и окон не меньше 1", +tileVal("conflicts") >= 1 && +tileVal("windows") >= 1, tileVal("conflicts") + "/" + tileVal("windows"));
      // карточка урока
      var L = lessonCell(0, true); openCell(L.d, L.p);
      check("T-U-12", "карточка урока: класс, урок, учитель, замок, «Убрать урок»", /7-A/.test(T("grid-pop").textContent) && !!T("pop-lesson") && !!T("pop-teacher") && !!T("pop-lock") && !!T("pop-clear"));
      T("pop-clear").click();
      check("T-U-12", "«Убрать урок» — ячейка пустая, карточка закрыта", !st().schedule.byClass[cls("7-A").id][L.d][L.p] && !D.querySelector('[data-test="grid-pop"]'));
      openCell(1, 1);
      var opt = Array.prototype.filter.call(T("pop-lesson").options, function(o){ return o.value; })[0].value;
      change(T("pop-lesson"), opt);
      check("T-U-12", "в окно поставлен урок из карточки — окна там больше нет", !!st().schedule.byClass[cls("7-A").id][1][1] && cellTd(1, 1).getAttribute("data-kind") !== "gap");
      // кабинеты в карточке
      await setState(function(s){
        s.rooms = [{ id: "rR1", name: "101", capacity: 30, type: null }, { id: "rR2", name: "102", capacity: 30, type: null }];
        var a = s.classes.filter(function(c){ return c.name === "7-A"; })[0], b = s.classes.filter(function(c){ return c.name === "8-A"; })[0];
        s.schedule.byClass[a.id][2][2].lessons.forEach(function(l){ l.roomId = "rR1"; });
        s.schedule.byClass[b.id][2][2].lessons.forEach(function(l){ l.roomId = "rR1"; });
      });
      gridClass("7-A"); openCell(2, 2);
      check("T-U-12", "кабинет занят другим классом — «Занят: 8-A», свободные списком", /8-A/.test(T("pop-room-busy").textContent) && /102/.test(T("pop-room-free").textContent));
      change(T("pop-room-0"), "rR2");
      var rc = st().schedule.byClass[cls("7-A").id][2][2];
      check("T-U-12", "кабинет сменён в карточке — правка помечена как ручная", rc.lessons[0].roomId === "rR2" && rc.manual === true && !D.querySelector('[data-test="pop-room-busy"]'));
      T("main").dispatchEvent(new W.KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
      check("T-U-12", "Esc закрывает карточку", !D.querySelector('[data-test="grid-pop"]'));
      // 188 классов: переключение без задержек
      await setState(function(s){
        var base = s.classes.slice(), k = 0;
        while(s.classes.length < 188){
          var b = base[k % base.length], c = JSON.parse(JSON.stringify(b));
          c.id = "cBig" + k; c.name = b.grade + "-Z" + k; s.classes.push(c);
          s.curriculum[c.id] = JSON.parse(JSON.stringify(s.curriculum[b.id] || {}));
          s.schedule.byClass[c.id] = JSON.parse(JSON.stringify(s.schedule.byClass[b.id]));
          k++;
        }
      });
      gridClass("5-A");
      var t0 = W.performance.now(); change(T("grid-other"), cls("9-A").id); var dt1 = W.performance.now() - t0;
      t0 = W.performance.now(); D.querySelector('[data-test^="grid-cls-9-"][aria-pressed="false"]').click(); var dt2 = W.performance.now() - t0;
      check("T-U-12", "188 классов: переключение класса быстрее 0,5 с", st().classes.length === 188 && dt1 < 500 && dt2 < 500, Math.round(dt1) + " / " + Math.round(dt2) + " мс");
    }],
    ["T-U-14", "0.12 Э9: «Перенос», «Печать и Excel», «Загрузить» — новый вид, одно место загрузки, вопрос о правках", async function(){
      await fresh();
      nav("print");
      check("T-U-14", "«Печать и Excel» без расписания — пустой экран с переходом к сетке", !!T("empty-print") && !!T("empty-print").querySelector('[data-act="kon-go"][data-id="grid"]'));
      nav("upload");
      check("T-U-14", "«Загрузить из eMaktab»: ничего не загружено — сразу выбор файлов, итога нет", !!T("page-upload") && !!T("upl-form") && !D.querySelector('[data-test="upl-done"]') && !D.querySelector('[data-test="card-upl-note"]'));
      await emkUpload([await fixture("school1_grid.xls"), await fixture("school1_load.xls")]);
      click('[data-emk-act="apply"]');
      nav("classes");
      check("T-U-14", "на «Классах» карточки загрузки нет — одно место в меню", !T("main").querySelector("[data-emk-act]") && !/Загружено из eMaktab|Расписание уже есть в eMaktab/.test(T("main").textContent));
      nav("upload");
      check("T-U-14", "загружено: итог (дата, неделя, 1731 урок), «Загрузить заново», выбор файлов скрыт", /1731 урок/.test(T("upl-done").textContent) && !!T("upl-again") && !D.querySelector('[data-test="upl-form"]'), T("upl-done").textContent);
      check("T-U-14", "правок нет — пояснение: настройки сохранятся, спрошу перед заменой", !T("card-upl-note").classList.contains("is-warn") && /сохранятся/.test(T("card-upl-note").textContent));
      await setState(function(s){ var c5 = s.classes.filter(function(c){ return c.name === "5-A"; })[0], g5 = s.schedule.byClass[c5.id]; for(var p=0; p<g5[0].length; p++){ if(g5[0][p]){ g5[0][p] = null; break; } } });
      nav("upload");
      check("T-U-14", "есть неперенесённая правка — предупреждение «1 изменение в 1 классе … пропадут»", T("card-upl-note").classList.contains("is-warn") && /1 изменение в 1 классе/.test(T("card-upl-note").textContent) && /пропадут/.test(T("card-upl-note").textContent), T("card-upl-note").textContent);
      T("upl-again").click();
      check("T-U-14", "«Загрузить заново» открывает выбор файлов", !!T("upl-form") && !D.querySelector('[data-test="upl-again"]'));
      await emkUpload([await fixture("school1_grid.xls"), await fixture("school1_load.xls")]);
      click('[data-emk-act="apply"]');
      check("T-U-14", "повторная загрузка спрашивает про правки: «пропадут: 1 изменение»", !T("dialog").hidden && /Заменить данные школы/.test(T("dialog").textContent) && /пропадут: 1 изменение/.test(T("dialog").textContent), T("dialog").textContent);
      click('[data-dialog="cancel"]');
      click('[data-emk-act="clear"]');
      check("T-U-14", "после отмены — снова итог и «Загрузить заново»", !D.querySelector('[data-test="upl-form"]') && !!T("upl-again"));
      nav("transfer");
      check("T-U-14", "«Перенос»: плитки, список классов, изменения выбранного класса, «Печать списка»", !!T("tiles-transfer") && D.querySelectorAll('[data-test="tr-row"]').length === 1 && T("tr-row").classList.contains("sel") && !!T("card-tr-class") && !!T("tr-print") && tileV("tr-cells") === "1");
      nav("print");
      check("T-U-14", "«Печать и Excel» — четыре карточки: Excel, класс, учитель, шахматка", ["excel", "class", "teacher", "all"].every(function(k){ return !!D.querySelector('[data-test="card-print-' + k + '"]'); }) && T("print-class-pick").options.length === st().classes.length);
      var printed = null; W.print = function(){ printed = D.getElementById("print").innerHTML; };
      change(T("print-class-pick"), cls("7-A").id); T("print-class").click();
      check("T-U-14", "печать выбранного класса", !!printed && /<h1>7-A<\/h1>/.test(printed));
      printed = null; T("print-teacher").click();
      check("T-U-14", "печать учителя", !!printed && /<h1>/.test(printed));    }],
    ["T-U-17", "0.12 Э12: вопрос перед удалением, кабинет класса, сверка дробных часов", async function(){
      await fresh(); await botSchool(1);
      T("mode-work").click();
      var c5 = cls("5-A"), c6 = cls("6-A");
      await setState(function(s){
        s.rooms = [{ id: "r207", name: "207", capacity: 30, type: null, classId: c5.id }, { id: "r208", name: "208", capacity: 30, type: null }];
      });
      gen();
      var sc = st().schedule, in5 = 0, other5 = 0, foreign = 0, why = [];
      function subjName(id){ var x = st().subjects.filter(function(q){ return q.id === id; })[0]; return x ? x.name : id; }
      Object.keys(sc.byClass).forEach(function(cid){ sc.byClass[cid].forEach(function(day){ day.forEach(function(cell){ (cell && cell.lessons || []).forEach(function(l){
        // две группы одновременно: кабинет класса — у одной, вторая — как раньше, без кабинета
        if(cid === c5.id){ if(l.roomId === "r207") in5++; else if(cell.lessons.length < 2 || !cell.lessons.some(function(x){ return x.roomId === "r207"; })){ other5++; why.push(subjName(l.subjectId)); } }
        else if(l.roomId === "r207") foreign++;
      }); }); }); });
      check("T-U-17", "кабинет класса: уроки 5-A без специального кабинета — в 207 (у двух групп — одна), другие классы его не занимают", in5 > 10 && !other5 && !foreign, in5 + " / " + other5 + " / " + foreign + " " + why.join(", "));
      nav("rooms");
      check("T-U-17", "«Кабинеты»: колонка «Закреплён за классом», у 207 — 5-A", /Закреплён за классом/.test(T("tbl-rooms").tHead.textContent) && /5-A/.test(T("room-row-r207").textContent) && /—/.test(T("room-row-r208").textContent));
      T("room-row-r208").click();
      var opt = Array.prototype.filter.call(T("room-class").options, function(o){ return o.value === c5.id; })[0];
      check("T-U-17", "в панели кабинета — выбор класса; у занятого класса видно «сейчас кабинет 207»", !!opt && /сейчас кабинет 207/.test(opt.textContent) && !!T("room-class-hint"));
      change(T("room-class"), c5.id);
      var rs = st().rooms;
      check("T-U-17", "у класса один кабинет: 208 → 5-A снимает 207", rs.filter(function(r){ return r.id === "r208"; })[0].classId === c5.id && !rs.filter(function(r){ return r.id === "r207"; })[0].classId);
      // удаление: вопрос, «Отмена» ничего не удаляет
      openClass("6-A"); T("cls-del").click();
      var dt = T("dialog").textContent;
      check("T-U-17", "«Удалить» класс — вопрос: что пропадёт (нагрузка, уроки в сетке), «Вернуть нельзя», «Отмена» первой",
        !T("dialog").hidden && /Удалить класс «6-A»\?/.test(dt) && /нагрузка класса — [\d,]+ ч в неделю/.test(dt) && /уроки класса в сетке — \d+/.test(dt) && /Вернуть нельзя/.test(dt)
        && D.querySelector("#dialog [data-dialog]").getAttribute("data-dialog") === "cancel", dt.slice(0, 160));
      click('[data-dialog="cancel"]');
      check("T-U-17", "«Отмена» — класс на месте", !!cls("6-A") && T("dialog").hidden);
      openClass("5-A"); T("cls-del").click(); click('[data-dialog="del"]');
      check("T-U-17", "«Удалить» — класс удалён, кабинет 208 больше ни за кем не закреплён", !cls("5-A") && !st().rooms.filter(function(r){ return r.id === "r208"; })[0].classId);
      nav("subjects");
      var sm = st().subjects.filter(function(x){ return x.name === "Математика"; })[0];
      T("subj-row-" + sm.id).click(); T("subj-del").click();
      var ds = T("dialog").textContent;
      check("T-U-17", "«Удалить» предмет — вопрос с часами во всех классах", /Удалить предмет/.test(ds) && /часы предмета в нагрузке \d+ класс/.test(ds), ds.slice(0, 160));
      click('[data-dialog="cancel"]');
      check("T-U-17", "«Отмена» — предмет на месте", st().subjects.some(function(x){ return x.id === sm.id; }));
      nav("rooms"); T("room-row-r207").click(); T("room-del").click();
      check("T-U-17", "«Удалить» кабинет — тоже с вопросом", !T("dialog").hidden && /Удалить кабинет «207»\?/.test(T("dialog").textContent));
      click('[data-dialog="cancel"]');
      check("T-U-17", "«Отмена» — кабинет на месте", st().rooms.length === 2);
      // сверка с планом: ровные дробные часы — не расхождение (было «География: 1,5 ч вместо 1,5 ч»)
      nav("check");
      var same = (planText().match(/([\d,]+) ч вместо ([\d,]+) ч/g) || []).filter(function(x){ var m = x.match(/([\d,]+) ч вместо ([\d,]+) ч/); return m[1] === m[2]; });
      var frac = Object.keys(st().curriculum).some(function(cid){ var pl = st().curriculum[cid]; return Object.keys(pl).some(function(k){ return pl[k].hours % 1; }); });
      check("T-U-17", "сверка с планом: нет «N ч вместо N ч» при дробных часах в плане", frac && !same.length, same.join(" | "));
    }],
    ["T-U-15", "0.12 Э11: у каждой строки интерфейса uzt(\"…\") есть узбекский перевод", async function(){
      var src = await (await fetch("../../index.html?t=" + Date.now())).text();
      var a = src.indexOf("var UZ_TR = {"), b = src.indexOf("function uzt(s)"), n = src.indexOf("function uzNorm(s)");
      var dict = (0, eval)("(function(){ " + src.slice(a, b) + src.slice(n, src.indexOf("\n", n)) + "\n return UZ_TR; })()");
      var re = /uzt\(\s*"((?:[^"\\]|\\.)*)"\s*\)/g, m, all = 0, miss = [];
      while((m = re.exec(src))){ var s; try{ s = JSON.parse('"' + m[1] + '"'); }catch(e){ continue; } all++; if(dict[s] === undefined && miss.indexOf(s) < 0) miss.push(s); }
      check("T-U-15", "все строки uzt(\"…\") переведены (" + all + ")", all > 500 && !miss.length, miss.slice(0, 5).join(" | "));
    }],
    ["T-U-04", "0.12 Э1: каркас конструктора — меню, адреса, шаги, меню «⋯»", async function(){
      await fresh();
      T("mode-work").click();
      var KEYS = ["school", "classes", "teachers", "subjects", "rooms", "load", "grid", "check", "transfer", "print", "upload"];
      var items = Array.prototype.map.call(D.querySelectorAll("#side [data-go]"), function(b){ return b.getAttribute("data-go"); });
      check("T-U-04", "меню: «Данные школы» (6), «Расписание» (4), внизу «Загрузить из eMaktab»", items.join(",") === KEYS.join(",") && /Данные школы[^]*Расписание/.test(T("side").textContent), items.join(","));
      check("T-U-04", "степпера и кнопки «Справочники» нет", !D.querySelector(".step-btn") && !D.querySelector('[data-test="refs"]'));
      var bad = [];
      KEYS.forEach(function(k){ nav(k); if(nav() !== k || W.location.hash !== "#/konstruktor/" + k || !T("main").textContent.trim()) bad.push(k + ":" + W.location.hash); });
      check("T-U-04", "каждый экран открывается из меню, адрес #/konstruktor/<экран>", !bad.length, bad.join(" "));
      W.location.hash = "#/konstruktor/rooms"; await sleep(120);
      check("T-U-04", "экран открывается по адресу", nav() === "rooms" && /Кабинеты/.test(T("page-title").textContent), nav());
      W.history.back(); await sleep(120);
      check("T-U-04", "«Назад» браузера — предыдущий экран", nav() === "upload", nav());
      var old = { classes: "classes", load: "load", schedule: "grid", "export": "print", refs: "teachers", "import": "upload" }, badOld = [];
      for(var k in old){ W.location.hash = "#/konstruktor/" + k; await sleep(80); if(nav() !== old[k] || W.location.hash !== "#/konstruktor/" + old[k]) badOld.push(k + "→" + nav()); }
      check("T-U-04", "старые адреса 0.11 перенаправляются (schedule→grid, export→print, refs→teachers, import→upload)", !badOld.length, badOld.join(" "));
      nav("check");
      check("T-U-04", "«Проверка» без расписания — пустой экран со ссылкой на сетку", !!D.querySelector('[data-test="empty-check"]') && !!D.querySelector('[data-test="empty-check"] [data-act="kon-go"][data-id="grid"]'));
      nav("transfer");
      check("T-U-04", "«Перенос в eMaktab» без загрузки — пустой экран", !!D.querySelector('[data-test="empty-transfer"]'));

      // «Шаг N из 7» и «Дальше» (П-7)
      nav("school");
      var path = [], texts = [];
      for(var i = 0; i < 7; i++){
        path.push(nav()); texts.push(T("crumb-text").textContent);
        var nx = D.querySelector('[data-test="foot-next"]'); if(!nx) break; nx.click();
      }
      check("T-U-05", "«Дальше» проходит 7 шагов: школа → классы → учителя → нагрузка → сетка → проверка → перенос", path.join(",") === "school,classes,teachers,load,grid,check,transfer" && texts[0] === "Шаг 1 из 7" && texts[6] === "Шаг 7 из 7", path.join(",") + " | " + texts.join(","));
      check("T-U-05", "на последнем шаге нет «Дальше», на первом — «Назад»", !D.querySelector('[data-test="foot-next"]') && /Нагрузка|Проверка/.test(T("foot-prev").textContent), T("foot-prev").textContent);
      T("foot-prev").click();
      check("T-U-05", "«← Проверка» возвращает на шаг назад", nav() === "check");
      nav("school");
      check("T-U-05", "на первом шаге нет «Назад»", !D.querySelector('[data-test="foot-prev"]'));
      nav("subjects");
      check("T-U-05", "необязательные экраны (предметы) — без «Шаг N из 7» и «Дальше»", !D.querySelector('[data-test="crumb"]') && !D.querySelector('[data-test="foot-nav"]'));

      // состояния шагов и бейджи после составления
      T("mode-bot").click(); await botSchool(1); botGen();
      nav("load");
      check("T-U-05", "отрезки: классы — готово, нагрузка без учителей — текущий, «Можно идти дальше и с замечаниями»",
        T("crumb-classes").className === "d" && T("crumb-load").className === "c" && !!D.querySelector('[data-test="foot-warn"]'), T("crumb-classes").className + "/" + T("crumb-load").className);
      nav("grid");
      check("T-U-05", "нагрузка без учителей — оранжевый отрезок; сетка составлена", T("crumb-load").className === "w" && T("crumb-grid").className === "c", T("crumb-load").className);
      var bClasses = T("nav-classes").querySelector('[data-test="badge"]'), bLoad = T("nav-load").querySelector('[data-test="badge"]'), bSchool = T("nav-school").querySelector('[data-test="badge"]');
      check("T-U-05", "бейджи меню: школа ✓, классов 11, нагрузка — оранжевое число предметов без учителя",
        bSchool && bSchool.textContent === "✓" && bClasses && bClasses.textContent === "11" && bLoad && bLoad.classList.contains("w") && +bLoad.textContent > 0, (bSchool && bSchool.textContent) + "/" + (bClasses && bClasses.textContent) + "/" + (bLoad && bLoad.textContent));

      // меню «⋯»: копия, восстановление, «Стереть всё» (переехали из «Выгрузки»)
      check("T-U-06", "меню «⋯» закрыто", T("more-menu").hidden);
      T("more").click();
      check("T-U-06", "«⋯» открывает: сохранить копию, восстановить, стереть всё", !T("more-menu").hidden && !!T("menu-save") && !!T("menu-restore") && !!T("menu-wipe"));
      T("main").click();
      check("T-U-06", "нажатие мимо меню закрывает его", T("more-menu").hidden);
      nav("print");
      check("T-U-06", "на «Печать и Excel» нет копии и сброса", !T("main").querySelector('[data-act="export-json"]') && !T("main").querySelector('[data-act="reset-all"]') && !D.querySelector('[data-test="page-transfer"]'));
      var log = stubDownloads(); T("more").click(); T("menu-save").click();
      check("T-U-06", "«Сохранить копию на компьютер» скачивает файл и закрывает меню", log.files.length === 1 && /\.json$/.test(log.files[0]) && T("more-menu").hidden, log.files.join(","));
      await setState(function(s){ s.teachers.push({ id: "teachW", name: "Учитель Стереть", subjects: [], maxPerDay: null, daysOff: [false,false,false,false,false,false] }); });
      W.confirm = function(){ return true; }; T("more").click(); T("menu-wipe").click();
      check("T-U-06", "«Стереть всё» — учителей и расписания нет, школа как новая", st().teachers.length === 0 && !st().schedule, st().teachers.length + "/" + !!st().schedule);
    }],
    ["T-6-26a", "P1-1: нагрузка параллели — назначение, «на всю параллель», «подставить известных», часы учителя", async function(){
      await fresh(); await botSchool(2);
      await setState(function(s){
        s.teachers.push({ id: "teachM", name: "Абдуллаева Нодира", subjects: [], maxPerDay: null, daysOff: [false,false,false,false,false,true] });
        s.teachers.push({ id: "teachR", name: "Каримова Дилноза", subjects: [], maxPerDay: 6, daysOff: [false,false,false,false,false,false] });
        // Каримова уже ведёт родной язык в 6-A — «известный» учитель предмета
        var c6 = s.classes.filter(function(c){ return c.name === "6-A"; })[0];
        var rod = s.subjects.filter(function(x){ return x.name === "Родной язык"; })[0].id;
        s.curriculum[c6.id][rod].teacherId = "teachR";
      });
      nav("load");
      // 0.12 (Э5): параллель — кнопкой, ячейка «часы + учитель» открывает карточку справа
      T("load-par-5").click();
      var math = st().subjects.filter(function(x){ return x.name === "Математика"; })[0].id;
      var cell = T("load-cell-" + cls("5-A").id + "-" + math);
      check("T-6-26a", "пустая ячейка подсвечена", cell.classList.contains("no") && /нет учителя/.test(cell.textContent), cell.textContent);
      cell.click(); T("load-pick-teachM").click();
      T("load-fill-row-" + math).click();
      var pl = st().curriculum;
      check("T-6-26a", "«Абдуллаева — в пустые классы»: 5-B тоже Абдуллаева", pl[cls("5-A").id][math].teacherId === "teachM" && pl[cls("5-B").id][math].teacherId === "teachM");
      var box = T("load-hours").textContent.replace(/\s+/g, " ");
      check("T-4-10", "часы учителя с учётом выходного: Абдуллаева 10 / 30 (без лимита × 5 рабочих дней)", /Абдуллаева Нодира ?10 \/ 30/.test(box), box.slice(0, 120));
      T("load-known").click();
      var rod2 = st().subjects.filter(function(x){ return x.name === "Родной язык"; })[0].id;
      check("T-6-26a", "«Подставить известных»: родной язык 5-A — Каримова", st().curriculum[cls("5-A").id][rod2].teacherId === "teachR", toast());
      tab(2);
      T("teach-row-teachM").click();
      check("T-4-11", "предметы учителя — из нагрузки: в строке «Математика», в панели — классы", /Математика/.test(T("teach-row-teachM").cells[1].textContent) && /Математика — [^;]*5-A, 5-B/.test(T("teach-what").textContent), T("teach-what").textContent.slice(0, 200));
      T("teach-row-teachR").click();
      var mx = T("teach-max"); change(mx, "");
      check("T-4-12", "пустой «Макс. в день» — без ограничения (не 1)", st().teachers.filter(function(x){ return x.id === "teachR"; })[0].maxPerDay === null);
    }],
    ["T-11-16", "P1-2: повторный проход — классы с правками не удаляются без вопроса; возврат к шагу", async function(){
      await fresh(); await botSchool(2); botGen();
      // правка у 5-B: назначенный учитель
      await setState(function(s){
        s.teachers.push({ id: "teachX", name: "Учитель Тестовый", subjects: [], maxPerDay: null, daysOff: [false,false,false,false,false,false] });
        var c = s.classes.filter(function(x){ return x.name === "5-B"; })[0], k = Object.keys(s.curriculum[c.id])[0];
        s.curriculum[c.id][k].teacherId = "teachX";
      });
      T("mode-bot").click(); bot('[data-bot-act="restart"]');
      check("T-11-16", "«Начать заново» — снова первый вопрос о eMaktab", /Расписание уже есть в eMaktab/.test(T("bot").textContent) && !!T("bot-path-import"));
      bot('[data-bot-act="path"][data-v="scratch"]');
      check("T-11-16", "повторный проход: число классов предложено текущее (5 класс — 2)", /2/.test($('#bot [data-bot-act="cnt"][data-g="5"]').parentNode.querySelector("strong").textContent));
      bot('[data-bot-act="cnt"][data-g="5"][data-v="-1"]'); bot('[data-bot-act="next"]');
      check("T-11-16", "вопрос перед удалением класса с правками", !T("dialog").hidden && /5-B/.test(T("dialog").textContent));
      click('[data-dialog="keep"]');
      check("T-11-16", "«Оставить» — 5-B на месте", !!cls("5-B"));
      await reload();
      check("T-11-16", "после перезагрузки — помощник на шаге 2 «О школе»", !T("bot").hidden && botQ() === "О школе");
      T("mode-work").click();
      check("T-11-16", "в конструкторе полоска «Помощник: шаг 2 из 5»", /Помощник: шаг 2 из 5/.test(T("mode-strip").textContent));
      click('#modeStrip [data-mode="bot"]');
      check("T-11-16", "«Вернуться в помощник» — тот же шаг", !T("bot").hidden && botQ() === "О школе");
    }],
    ["T-2-23", "P1-4: разное число уроков по дням", async function(){
      await fresh(); await botSchool(1);
      openPar(7);
      change(T("par-periods"), "byday");
      change($('[data-act="parallel-pday"][data-grade="7"][data-day="0"]'), "7");
      change($('[data-act="parallel-pday"][data-grade="7"][data-day="5"]'), "4");
      var lbl = T("par-periods").selectedOptions[0].text;
      check("T-2-23", "подпись профиля «Пн 7, Вт–Пт 6, Сб 4»", lbl === "Пн 7, Вт–Пт 6, Сб 4", lbl);
      gen();
      var g = cells("7-A");
      check("T-2-23", "сетка 7-A по дням: 7, 6, 6, 6, 6, 4", g.map(function(r){ return r.length; }).join() === "7,6,6,6,6,4", g.map(function(r){ return r.length; }).join());
      gridClass("7-A");
      check("T-2-23", "лишние уроки субботы показаны как «нет урока»", D.querySelectorAll('[data-test="grid-table"] td[data-kind="off"]').length >= 3);
      check("T-2-23", "0 окон и 0 неразмещённых", st().schedule.unplacedCount === 0 && windowsOf(st().schedule.byClass).windows === 0);
    }],
    ["T-8-28", "P1-5: правила предметов соблюдаются", async function(){
      await fresh(); await botSchool(1);
      // 0.12 (Э6): правила предмета — в панели предмета на экране «Предметы»
      function sid(n){ return st().subjects.filter(function(x){ return x.name === n; })[0].id; }
      var vosId = sid("Воспитание"), peId = sid("Физическое воспитание");
      nav("subjects"); T("subj-row-" + vosId).click();
      change(T("subj-fixed-day"), "0"); change(T("subj-fixed-period"), "1");
      T("subj-row-" + peId).click(); T("subj-np-1").click();
      check("T-8-28", "правила в таблице предметов", /всегда Пн, 1-й урок/.test(T("subj-row-" + vosId).textContent) && /не на уроках: 1/.test(T("subj-row-" + peId).textContent));
      gen();
      var s = st(), vos = s.subjects.filter(function(x){ return x.name === "Воспитание"; })[0].id, pe = s.subjects.filter(function(x){ return x.name === "Физическое воспитание"; })[0].id;
      var bad = [];
      s.classes.forEach(function(c){
        var g = s.schedule.byClass[c.id], plan = s.curriculum[c.id] || {};
        if(plan[vos] && plan[vos].hours > 0 && !(g[0][0] && g[0][0].lessons.some(function(l){ return l.subjectId === vos; }))) bad.push(c.name + ": воспитание не в Пн 1");
        g.forEach(function(row){ if(row[0] && row[0].lessons.some(function(l){ return l.subjectId === pe; })) bad.push(c.name + ": физкультура 1-м"); });
      });
      check("T-8-28", "«Воспитание» — Пн 1-й урок, физкультура не 1-м", bad.length === 0 && s.schedule.unplacedCount === 0, bad.slice(0, 4).join("; "));
    }],
    ["T-8-20", "P1-6: причина неразмещения, «Поставить», обмен перетаскиванием, кабинет при ручной правке", async function(){
      await fresh(); await botSchool(1);
      await setState(function(s){
        s.teachers.push({ id: "teachZ", name: "Тестова Анна", subjects: [], maxPerDay: null, daysOff: [true,true,true,true,true,true] });
        var c = s.classes.filter(function(x){ return x.name === "5-A"; })[0];
        s.curriculum[c.id][s.subjects.filter(function(x){ return x.name === "Математика"; })[0].id].teacherId = "teachZ";
      });
      gen();
      var li = T("unplaced-item").textContent;
      check("T-8-20", "причина: у учителя выходной", /выходной/.test(li), li);
      // учитель выходит на работу — уроки можно поставить вручную
      await setState(function(s){ s.teachers.filter(function(x){ return x.id === "teachZ"; })[0].daysOff = [false,false,false,false,false,false]; });
      nav("grid");
      var before = st().schedule.unplacedCount;
      T("unplaced-place").click();
      var can = D.querySelectorAll('[data-test="cell-place"]').length;
      check("T-9-10", "«Поставить» подсвечивает допустимые ячейки", can > 0, "ячеек: " + can);
      T("cell-place").click();
      check("T-9-10", "урок поставлен, неразмещённых меньше", st().schedule.unplacedCount === before - 1);
      // перетаскивание: уроки двух ячеек меняются местами (сверка по сохранённому расписанию)
      function tdLesson(td){ var c = st().schedule.byClass[td.getAttribute("data-class")][+td.getAttribute("data-day")][+td.getAttribute("data-period")]; return c ? c.lessons.map(function(l){ return l.subjectId; }).join("+") : ""; }
      var tds = Array.prototype.slice.call(D.querySelectorAll('[data-test="grid-table"] td[draggable="true"]'));
      var a = tds[0], b = tds[2], ta = a.getAttribute("data-test"), tb = b.getAttribute("data-test"), va = tdLesson(a), vb = tdLesson(b);
      var dt = new W.DataTransfer();
      a.dispatchEvent(new W.DragEvent("dragstart", { bubbles: true, dataTransfer: dt }));
      b.dispatchEvent(new W.DragEvent("dragover", { bubbles: true, cancelable: true, dataTransfer: dt }));
      b.dispatchEvent(new W.DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: dt }));
      check("T-9-11a", "перетаскивание меняет уроки местами", tdLesson(T(ta)) === vb && tdLesson(T(tb)) === va, va + " / " + vb);
      // ручная правка урока с типом кабинета получает кабинет
      await setState(function(s){
        var inf = s.subjects.filter(function(x){ return x.name === "Информатика и информационные технологии"; })[0]; inf.roomType = "комп";
        s.rooms = [{ id: "roomK", name: "Комп-1", capacity: 30, type: "комп" }];
      });
      gridClass("5-A");
      var infId = st().subjects.filter(function(x){ return x.name === "Информатика и информационные технологии"; })[0].id;
      var cand = Array.prototype.filter.call(D.querySelectorAll('[data-test="grid-table"] td[data-kind="lesson"]'), function(td){ return tdLesson(td) !== infId; })[0];
      cand.querySelector('[data-act="cell-open"]').click();
      var sel = T("pop-lesson"), sd = cand.getAttribute("data-day"), sp = cand.getAttribute("data-period");
      change(sel, infId);
      check("T-9-04", "ручная правка в карточке урока: кабинет подобран", st().schedule.byClass[cls("5-A").id][sd][sp].lessons[0].roomId === "roomK");
    }],
    ["T-6-28", "P1-7: пустые планы, копирование на параллель, план не по источнику", async function(){
      await fresh(); await botSchool(2);
      await setState(function(s){ var c = s.classes.filter(function(x){ return x.name === "6-B"; })[0]; s.curriculum[c.id] = {}; });
      nav("load"); click('[data-act="load-fill-empty"]');
      check("T-6-28", "«Заполнить пустые планы» заполнил 6-B", planHours("6-B") > 0 && /Заполнено планов: 1/.test(toast()), toast());
      tab(4); click('[data-act="cur-select-class"][data-id="' + cls("7-A").id + '"]');
      await setState(function(s){ var c = s.classes.filter(function(x){ return x.name === "7-A"; })[0]; var k = Object.keys(s.curriculum[c.id])[0]; s.curriculum[c.id][k].hours = 9; s.teachers.push({ id: "teachC", name: "Копия Учитель", subjects: [], maxPerDay: null, daysOff: [false,false,false,false,false,false] }); s.curriculum[c.id][k].teacherId = "teachC"; });
      tab(4); click('[data-act="cur-select-class"][data-id="' + cls("7-A").id + '"]');
      click('[data-act="cur-copy"]'); click('[data-dialog="without"]');
      var pa = st().curriculum[cls("7-A").id], pb = st().curriculum[cls("7-B").id], k0 = Object.keys(pa)[0];
      check("T-6-29", "план 7-A скопирован в 7-B без учителей", pb[k0] && pb[k0].hours === 9 && !pb[k0].teacherId && Object.keys(pb).length === Object.keys(pa).length);
      tab(0); change(classSel("class-direction", "8-A"), "v:law");
      tab(4); click('[data-act="cur-select-class"][data-id="' + cls("8-A").id + '"]');
      check("T-2-11", "подсказка: план по другому документу и кнопка обновления", /План заполнен по другому документу/.test(T("main").textContent));
      click('.warn-box [data-act="cur-autofill"]');
      check("T-2-11", "после обновления подсказки нет", !/План заполнен по другому документу/.test(T("main").textContent));
    }],
    ["T-4-08", "P1-8/P1-9: импорт с предпросмотром, должности, однофамильцы, буквы классов", async function(){
      await fresh(); await botSchool(1);
      tab(2);
      var staff = "<table><tr><td colspan='5'>Сотрудники школы №1</td></tr><tr><td>№</td><td>F.I.Sh.</td><td>Lavozim</td><td>Sinflar</td></tr>" +
        "<tr><td>1</td><td>Каримова Дилноза</td><td>Учитель</td><td>5-А</td></tr>" +
        "<tr><td>2</td><td>Турсунов Бахтиёр</td><td>Охранник</td><td></td></tr>" +
        "<tr><td>3</td><td>Рахимов Жасур</td><td>Заместитель директора</td><td></td></tr>" +
        "<tr><td>4</td><td>Каримова Дилноза</td><td>Учитель</td><td>12-Z</td></tr></table>";
      var si = $("#import-staff-input"), dt = new W.DataTransfer();
      dt.items.add(new W.File([staff], "s.xls", { type: "text/html" }));
      si.files = dt.files; si.dispatchEvent(new W.Event("change", { bubbles: true })); await sleep(60);
      var rows = Array.prototype.map.call(D.querySelectorAll(".imp-table tbody tr"), function(r){ return (r.querySelector("input").checked ? "+" : "-") + r.textContent; });
      check("T-4-08", "заголовок найден во 2-й строке, узбекские заголовки", rows.length === 4, rows.join(" | "));
      check("T-4-08", "охранник и зам. директора не отмечены, однофамилец помечен", /^-.*Турсунов/.test(rows[1]) && /^-.*Рахимов.*ведёт уроки/.test(rows[2]) && /^-.*совпадает ФИО/.test(rows[3]), rows.join(" | "));
      click('[data-imp-act="apply"]');
      var s = st();
      check("T-4-08", "добавлен 1 учитель; «5-А» кириллицей не создал дубль «5-A»", s.teachers.length === 1 && s.classes.filter(function(c){ return c.grade === 5; }).length === 1, s.classes.map(function(c){ return c.name; }).join(" "));
      tab(0); change(classSel("class-lang", "6-A"), "ru");
      check("T-2-24", "русский класс остаётся 6-A: буквы — латиница для всех языков", !!cls("6-A") && !cls("6-А") && cls("6-A").instructionLang === "ru");
      click('[data-act="parallel-add"][data-grade="6"]');
      check("T-2-24", "следующий класс — 6-B (латиница по порядку A–Z)", !!cls("6-B"));
    }],
    ["T-10-11", "P1-11: звонки, время в сетке, печать, лист «Все классы»", async function(){
      await fresh(); await botSchool(1); botGen();
      nav("school"); T("bells-init").click();
      var bt = T("bell-row-1").textContent + " | " + T("bell-row-4").textContent;
      check("T-2-25", "звонки 1-й смены: 1. 08:00–08:45, 4-й после большой перемены", /^1\s*08:00–08:45$/.test(T("bell-row-1").textContent) && /^4\s*10:55–11:40$/.test(T("bell-row-4").textContent), bt);
      gridClass("5-A");
      check("T-2-25", "время урока в сетке класса", /08:00–08:45/.test(T("grid-table").querySelector("tbody th").textContent));
      tab(6);
      var printed = null; W.print = function(){ printed = D.getElementById("print").innerHTML; };
      click('[data-act="print-all"]');
      check("T-10-07", "шахматка всех классов собрана для печати", !!printed && /Расписание всех классов/.test(printed) && /5-A/.test(printed));
      var log = stubDownloads(); click('[data-act="export-xlsx"]'); await sleep(30);
      check("T-10-09", "первый лист Excel — «Все классы»", log.sheets && log.sheets[0] === "Все классы", log.sheets && log.sheets.slice(0, 3).join());
      var z = W.XLSX.CFB.read(new Uint8Array(log.bytes), { type: "array" });
      function part(re){ var i = z.FullPaths.findIndex(function(p){ return re.test(p); }); return i < 0 ? "" : new TextDecoder().decode(z.FileIndex[i].content); }
      var styles = part(/styles\.xml$/), sh = part(/worksheets\/sheet2\.xml$/);
      check("T-10-12", "Excel: перенос строк, выравнивание по верху, ширина колонок и высота строк заданы", /wrapText="(1|true)"/.test(styles) && /<cols>/.test(sh) && /customHeight="1"/.test(sh), sh.slice(0, 120));
    }]
  ];

  // only — регулярное выражение по ID сценария (для отладки одного сценария): runSmoke(/T-4-/)
  window.runSmoke = async function(only){
    frame = document.getElementById("app");
    results = [];
    for(var i = 0; i < scenarios.length; i++){
      if(only && !only.test(scenarios[i][0])) continue;
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
