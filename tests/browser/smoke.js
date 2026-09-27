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
  // текст раздела «Расхождения с учебным планом» карточки проверки
  function planText(){
    var d = Array.prototype.filter.call(D.querySelectorAll(".check-card details"), function(x){ return /учебным планом/.test(x.textContent); })[0];
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
  // бот (P1-2): по perGrade классов в 5–11, остальное — по умолчанию; останавливается на шаге «Составить»
  async function botSchool(perGrade){
    for(var g = 5; g <= 11; g++){ for(var k = 1; k < perGrade; k++) bot('[data-bot-act="cnt"][data-g="' + g + '"][data-v="1"]'); }
    bot('[data-bot-act="next"]');
    change($('#bot [data-bot-act="mode-per"][data-r="j"]'), "5");
    bot('[data-bot-act="next"]'); bot('[data-bot-act="next"]'); bot('[data-bot-act="next"]');
  }
  function st(){ return JSON.parse(W.localStorage.getItem("timetable_builder_v4")).state; }
  function cls(name){ return st().classes.filter(function(c){ return c.name === name; })[0]; }
  function $(sel){ var el = D.querySelector(sel); if(!el) throw new Error("нет элемента " + sel); return el; }
  function click(sel){ $(sel).click(); }
  function change(el, v){ el.value = v; el.dispatchEvent(new W.Event("change", { bubbles: true })); }
  // старые номера вкладок → новый мастер (P1-3): 0 классы, 1 предметы, 2 учителя, 3 кабинеты, 4 план класса, 5 расписание, 6 выгрузка
  function tab(i){
    var steps = D.querySelectorAll(".step-btn");
    if(i === 0) steps[0].click();
    else if(i === 1 || i === 2 || i === 3){ if(!D.getElementById("refsBtn").classList.contains("is-on")) D.getElementById("refsBtn").click(); click('[data-act="refs-tab"][data-id="' + ["", "subjects", "teachers", "rooms"][i] + '"]'); }
    else if(i === 4){ steps[1].click(); click('[data-act="load-mode"][data-id="class"]'); }
    else if(i === 5) steps[2].click();
    else if(i === 6) steps[3].click();
  }
  function bot(sel){ click("#bot " + sel); }
  function botGen(){ bot('[data-bot-act="gen"]'); if(!D.getElementById("dialog").hidden) click('[data-dialog="all"]'); }
  function botQ(){ var q = D.querySelector("#bot .bot-q"); return q ? q.textContent : ""; }
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
  // бот до шага «Составить»: тип школы, по одному классу в параллели; plan "no" — планы стёрты
  async function botToGen(type, plan){
    if(type === "classical") bot('[data-bot-act="type"][data-v="classical"]');
    await botSchool(1);
    if(plan === "no"){
      await setState(function(s){ Object.keys(s.curriculum).forEach(function(k){ s.curriculum[k] = {}; }); });
      bot('[data-bot-act="open"]');
    }
  }

  var scenarios = [
    ["T-11-01", "Бот: 5 экранов — расписание без учителей, итог, переход в мастер", async function(){
      await fresh();
      check("T-11-01", "бот открыт на шаге 1 из 5", /сколько классов/.test(botQ()) && D.querySelectorAll("#bot .bot-steps li").length === 5, botQ());
      await botToGen("criterial", "yes");
      check("T-11-01", "шаг «Составить»: список готовности", /Можно составлять/.test(botQ()) && /Учителей нет/.test($("#bot").textContent));
      botGen();
      var s = st();
      check("T-11-01", "расписание построено", !!s.schedule);
      check("T-8-17", "0 неразмещённых (групповая пара без учителя не распадается)", s.schedule.unplacedCount === 0, "неразмещённых: " + s.schedule.unplacedCount);
      check("T-11-01", "итог: плитки и открыт шаг «Расписание»", /Расписание составлено/.test(botQ()) && D.querySelectorAll("#bot .bot-tile").length === 4 && $(".step-btn.active").textContent.indexOf("Расписание") >= 0);
      check("T-7-01", "без учителей — предупреждение, не ошибка", !D.querySelector("#main .warn-box") && /Учителей нет/.test($("#main .note-box").textContent));
      await reload();
      check("T-11-02", "после перезагрузки бот помнит шаг (итог)", /Первое расписание составлено/.test($("#bot").textContent));
    }],
    ["T-11-11", "Бот: план не заполнен — ошибка на шаге «Составить», переход к месту, составление", async function(){
      await fresh(); await botToGen("criterial", "no");
      check("T-11-11", "бот показывает ошибку, «Составить» недоступна", /исправьте/.test(botQ()) && $('#bot [data-bot-act="gen"]').disabled, botQ());
      bot('[data-bot-act="issue"]');
      check("T-11-11", "открыт план класса", /Нагрузка/.test($(".page-title").textContent) && !!D.querySelector("#cur-add-subj-select"));
      var sel = $("#cur-add-subj-select");
      sel.value = Array.prototype.filter.call(sel.options, function(o){ return /^Математика/.test(o.text); })[0].value;
      click('[data-act="cur-add-subject"]');
      change($('[data-act="cur-hours"]'), "5");
      check("T-11-11", "после исправления «Составить» доступна", !$('#bot [data-bot-act="gen"]').disabled);
      botGen();
      check("T-11-11", "расписание составлено", !!st().schedule && /Расписание составлено/.test(botQ()));
      D.getElementById("langToggle").click();
      var bad = cyr(D.getElementById("bot").innerText, /^\d+-[A-Z]/);
      check("T-11-13", "бот на узбекском без кириллицы", bad.length === 0, bad.slice(0, 5).join(" "));
      D.getElementById("langToggle").click();
    }],
    ["T-E-01", "Пилот: загрузка из eMaktab — сетка и нагрузка (обезличенная школа 1), предпросмотр, сопоставление предметов", async function(){
      await fresh();
      click("#emkBtn");
      check("T-E-01", "страница «Загрузка из eMaktab» из шапки", /Загрузка из eMaktab/.test($(".page-title").textContent) && !!D.querySelector("[data-emk-file]"));
      await emkUpload([await fixture("school1_load.xls")]);
      check("T-E-01", "только нагрузка — просьба загрузить сетку", /Теперь нужна сетка расписания/.test($("#main").textContent) && !D.querySelector('[data-emk-act="apply"]'));
      await emkUpload([await fixture("school1_grid.xls")]);
      var tiles = Array.prototype.map.call(D.querySelectorAll(".emk-tiles .imp-tile"), function(x){ return x.textContent; }).join(" | ");
      check("T-E-01", "предпросмотр: 85 учителей, 50 классов, 57 кабинетов, 1731 урок", /85учителей/.test(tiles) && /50классов/.test(tiles) && /57кабинетов/.test(tiles) && /1731уроков в неделю/.test(tiles), tiles);
      var map = Array.prototype.map.call(D.querySelectorAll(".emk-map tbody tr"), function(tr){ var s = tr.querySelector("select"); return tr.cells[0].textContent + "→" + s.options[s.selectedIndex].text + "[" + tr.cells[3].textContent + "]"; });
      check("T-E-01", "предметы: полные названия из нагрузки, дословно ✓, синоним — «проверьте», нет в справочнике — новый",
        map.indexOf("Математика→Математика[✓]") >= 0 && map.indexOf("Физкультура→Физическое воспитание[проверьте]") >= 0 && map.indexOf("Кл. час→＋ новый предмет «Кл. час»[новый]") >= 0 && map.indexOf("Английский язык→Английский язык[✓]") >= 0, map.slice(0, 12).join(" | "));
      check("T-E-01", "сверка сетки и нагрузки в замечаниях", /Есть в нагрузке, но нет в сетке недели/.test($(".emk-notes").textContent));
      click('[data-emk-act="apply"]');
      var s = st(), n = 0;
      Object.keys(s.schedule.byClass).forEach(function(k){ s.schedule.byClass[k].forEach(function(r){ r.forEach(function(c){ if(c) n += c.lessons.length; }); }); });
      check("T-E-01", "созданы учителя, классы, кабинеты; в сетке все 1731 урок; источник — eMaktab", s.teachers.length === 85 && s.classes.length === 50 && s.rooms.length === 57 && n === 1731 && s.schedule.source === "emaktab" && s.emaktab.base.length === 1731, s.teachers.length + "/" + s.classes.length + "/" + s.rooms.length + "/" + n);
      check("T-E-01", "открыт шаг «Расписание» с проверкой", $(".step-btn.active").textContent.indexOf("Расписание") >= 0 && !!D.querySelector(".check-card"));
      var pairs = 0, en = s.subjects.filter(function(x){ return x.name === "Английский язык (группа 2)"; })[0];
      Object.keys(s.curriculum).forEach(function(k){ var p = s.curriculum[k]; Object.keys(p).forEach(function(sid){ if(p[sid].syncWith && sid === en.id) pairs++; }); });
      check("T-E-01", "группы — связанные пары «предмет» + «предмет (группа 2)» с ярлыком eMaktab", pairs > 10, "пар английского: " + pairs);
      var t1 = s.teachers.filter(function(x){ return x.name === "Учитель001 А.Б."; })[0];
      check("T-E-01", "ФИО и имя eMaktab у учителя; у учителя нет лимита уроков в день", !!t1 && t1.ext.emaktab === "Учитель001 А.Б." && !t1.maxPerDay);
      var tl = Array.prototype.map.call(D.querySelectorAll(".check-tiles .bot-tile"), function(x){ return x.textContent; }).join(" | ");
      check("T-E-03", "проверка: 0 конфликтов учителей; спортзал не считается конфликтом кабинета", /^0конфликтов учителей/.test(tl) && /0кабинетов заняты двумя классами/.test(tl), tl);
      await reload();
      var n2 = 0; Object.keys(st().schedule.byClass).forEach(function(k){ st().schedule.byClass[k].forEach(function(r){ r.forEach(function(c){ if(c) n2 += c.lessons.length; }); }); });
      check("T-E-01", "после перезагрузки сетка на месте, ничего не убрано", n2 === 1731 && !st().schedule.stale, n2);
      var ev = (W.__ttEvents || []).concat(events).filter(function(e){ return e.goal === "tt_import_emaktab_result" && e.params.ok; })[0];
      check("T-E-01", "событие import_emaktab_result: только количества", !!ev && ev.params.teachers === 85 && ev.params.lessons === 1731 && ev.params.files === "grid,load");
    }],
    ["T-E-02", "Пилот: повторная загрузка — id и настройки сохраняются, сопоставление предметов запомнено, урок в выходной", async function(){
      await fresh(); click("#emkBtn");
      await emkUpload([await fixture("school1_grid.xls"), await fixture("school1_load.xls")]);
      // «Кл. час» сопоставляем с «Воспитанием» — должно запомниться
      var sel = Array.prototype.filter.call(D.querySelectorAll(".emk-map tbody tr"), function(tr){ return tr.cells[0].textContent === "Кл. час"; })[0].querySelector("select");
      change(sel, Array.prototype.filter.call(sel.options, function(o){ return o.text === "Воспитание"; })[0].value);
      click('[data-emk-act="apply"]');
      var s = st(), t1 = s.teachers.filter(function(x){ return x.name === "Учитель001 А.Б."; })[0], c5 = cls("5-A");
      await setState(function(x){ x.teachers.filter(function(t){ return t.id === t1.id; })[0].daysOff[0] = true; x.classes.filter(function(c){ return c.id === c5.id; })[0].studentCount = 25; });
      click("#emkBtn");
      await emkUpload([await fixture("school1_grid.xls"), await fixture("school1_load.xls")]);
      var row = Array.prototype.filter.call(D.querySelectorAll(".emk-map tbody tr"), function(tr){ return tr.cells[0].textContent === "Кл. час"; })[0];
      check("T-E-02", "сопоставление «Кл. час → Воспитание» запомнено", row && row.querySelector("select").options[row.querySelector("select").selectedIndex].text === "Воспитание" && /✓/.test(row.cells[3].textContent));
      check("T-E-02", "предупреждение о повторной загрузке", /Повторная загрузка/.test($("#main").textContent));
      click('[data-emk-act="apply"]');
      check("T-E-02", "подтверждение замены данных", !D.getElementById("dialog").hidden && /Заменить данные школы/.test(D.getElementById("dialog").textContent));
      click('[data-dialog="go"]');
      var s2 = st(), t2 = s2.teachers.filter(function(x){ return x.id === t1.id; })[0], c52 = cls("5-A");
      check("T-E-02", "без дублей; id учителя и класса прежние; выходной и число учеников сохранены", s2.teachers.length === 85 && s2.classes.length === 50 && !!t2 && t2.daysOff[0] === true && c52.id === c5.id && c52.studentCount === 25);
      tab(5);
      check("T-E-03", "проверка: урок в выходной учителя", /Урок в выходной учителя/.test($(".check-card").textContent));
    }],
    ["T-E-04", "Пилот: неверные файлы — понятные сообщения, событие с причиной", async function(){
      await fresh(); click("#emkBtn");
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
      tab(5);
      var tl = Array.prototype.map.call(D.querySelectorAll(".check-tiles .bot-tile"), function(x){ return x.textContent; }).join(" | ");
      check("T-E-03", "плитки: 1 конфликт учителя, 1 объединённый урок, 1 кабинет, окно", /^1конфликтов учителей/.test(tl) && /1возможно, объединённых уроков/.test(tl) && /1кабинетов заняты двумя классами/.test(tl) && /[1-9]окон у классов/.test(tl), tl);
      var groups = Array.prototype.map.call(D.querySelectorAll(".check-group summary"), function(x){ return x.textContent; }).join(" | ");
      check("T-E-03", "списки по видам проблем", /Конфликты учителей/.test(groups) && /объединённый урок/.test(groups) && /Кабинет занят/.test(groups) && /Окна у классов/.test(groups), groups);
      click('.check-group [data-act="check-go"]');
      var fc = D.querySelector("td.focus-cell");
      check("T-E-03", "переход к ячейке: открыт класс, ячейка подсвечена", !!fc && fc.getAttribute("data-day") === "2" && fc.getAttribute("data-period") === "1", fc ? fc.outerHTML.slice(0, 120) : "нет");
    }],
    ["T-E-06", "Пилот: направление класса по предметам ЕСП — вопрос, «Поставить», сверка часов с планом", async function(){
      await fresh();
      setClass("class-direction", "10-A", "t:nd1");
      autofill("10-A");
      var names = planNames("10-A");
      check("T-E-06", "план направления «Бизнес и профессии» — с предметами «… Экономика»", names.indexOf("Алгебра Экономика") >= 0, names.join(", "));
      await setState(function(s){ var c = s.classes.filter(function(x){ return x.name === "10-A"; })[0]; c.track = "none"; delete c.directionAsked; s.schedule = { byClass: {}, unplaced: [], unplacedCount: 0 }; });
      tab(5);
      var dir = D.querySelector(".check-dir") ? $(".check-dir").textContent : "";
      check("T-E-06", "вопрос: у 10-A похоже направление «Бизнес и профессии» — по предметам «… Экономика»", /10-A/.test(dir) && /Бизнес и профессии/.test(dir) && /Алгебра Экономика/.test(dir), dir.slice(0, 200));
      check("T-E-06", "пока направление не поставлено — расхождения с базовым планом", /10-A —/.test(planText()));
      click('[data-act="dir-set"][data-id="' + cls("10-A").id + '"]');
      check("T-E-06", "«Поставить» — направление записано в настройку класса, вопрос снят, план сходится", cls("10-A").track === "nd1" && !D.querySelector(".check-dir") && !/10-A —/.test(planText()), planText().slice(0, 200));
      await setState(function(s){
        var c = s.classes.filter(function(x){ return x.name === "10-A"; })[0], pl = s.curriculum[c.id];
        var ae = s.subjects.filter(function(x){ return x.name === "Алгебра Экономика"; })[0].id, a = s.subjects.filter(function(x){ return x.name === "Алгебра"; })[0].id;
        pl[a] = pl[ae]; delete pl[ae];
      });
      tab(5);
      check("T-E-06", "стоит обычная «Алгебра» — заметка «по направлению нужен «Алгебра Экономика»»", /10-A — стоит «Алгебра», а по направлению нужен «Алгебра Экономика»/.test(planText()), planText().slice(0, 300));
      await setState(function(s){ var c = s.classes.filter(function(x){ return x.name === "10-A"; })[0]; c.track = "none"; delete c.directionAsked; });
      tab(5); click('[data-act="dir-no"][data-id="' + cls("10-A").id + '"]');
      check("T-E-06", "«Нет, без направления» — вопрос больше не задаётся", !D.querySelector(".check-dir") && cls("10-A").directionAsked === true && cls("10-A").track === "none");
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
      bot('[data-bot-act="open"]'); botGen(); tab(5);
      click('[data-act="sched-select-class"][data-id="' + cls("1-A").id + '"]');
      check("T-2-26", "сетка 1-A: уроки школы 7–11 и время по звонкам школы", /1-A — расписание \(уроки 7–11\)/.test($("#main").textContent) && /^7/.test($("#main table tbody tr th").textContent) && /14:10–14:55/.test($("#main table tbody tr th").textContent), $("#main table tbody tr th").textContent);
      check("T-2-26", "после миграции составление без конфликтов", st().schedule && !Array.prototype.some.call(D.querySelectorAll("td.conflict"), Boolean));
      tab(0); change($('[data-act="parallel-start"][data-grade="6"]'), "4");
      check("T-2-26", "параллель 6 с 4-го урока — смены перекрываются", st().config.parallels[6].start === 4 && /Уроки школы:\s*1–12/.test($("#main").textContent.replace(/\s+/g, " ")), $("#main .hint").textContent.slice(-80));
      tab(5); click('[data-act="sched-select-class"][data-id="' + cls("6-A").id + '"]');
      check("T-2-26", "сетка 6-A: уроки 4–9", /6-A — расписание \(уроки 4–9\)/.test($("#main").textContent));
    }],
    ["T-11-20", "Пилот: бот — «Расписание уже есть в eMaktab?», загрузка, итог проверки, узбекский", async function(){
      await fresh(true);
      check("T-11-20", "первый вопрос бота", /Расписание уже есть в eMaktab\?/.test(botQ()) && !!D.querySelector('#bot [data-bot-act="path"][data-v="import"]'));
      D.getElementById("langToggle").click();
      var bad = cyr($("#bot").textContent, /^eMaktab/);
      check("T-11-13", "первый вопрос бота на узбекском без кириллицы", bad.length === 0, bad.slice(0, 5).join(" "));
      D.getElementById("langToggle").click();
      bot('[data-bot-act="path"][data-v="import"]');
      check("T-11-20", "ветка загрузки: 2 шага", D.querySelectorAll("#bot .bot-steps li").length === 2 && /Загрузите расписание из eMaktab/.test(botQ()));
      await emkUpload([await fixture("school1_grid.xls"), await fixture("school1_load.xls")], "#bot ");
      click('#bot [data-emk-act="apply"]');
      check("T-11-20", "итог проверки в боте и открыт шаг «Расписание»", /Расписание загружено/.test(botQ()) && D.querySelectorAll("#bot .bot-tile").length === 4 && $(".step-btn.active").textContent.indexOf("Расписание") >= 0, botQ());
      await reload();
      check("T-11-20", "после перезагрузки бот свёрнут: загружено из eMaktab", /Расписание загружено из eMaktab/.test($("#bot").textContent));
    }],
    ["T-9-07", "«По учителям» открывает вид по учителям", async function(){
      await fresh(); await botToGen("criterial", "yes"); botGen();
      tab(2); $("#new-teach-name").value = "Иванова"; click('[data-act="teach-add"]');
      tab(5); click('[data-act="generate"]'); click('[data-id="__teachers__"]');
      check("T-9-07", "активна кнопка «По учителям»", /По учителям/.test($(".pill-btn.active").textContent));
      check("T-9-07", "карточка учителя", /Иванова/.test($("#main").textContent));
    }],
    ["T-9-05", "Закрепление урока: значок, закрепление, повторная генерация, открепление", async function(){
      await fresh(); await botToGen("criterial", "yes"); botGen();
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
      await fresh(); await botToGen("criterial", "yes"); botGen();
      click('[data-act="sched-select-class"][data-id="' + cls("7-A").id + '"]');
      var filled = Array.prototype.filter.call(D.querySelectorAll("td select[data-act=cell-change]"), function(s){ return s.value; });
      var noColor = filled.filter(function(s){ return !/cat-/.test(s.closest("td").className); });
      check("T-9-11", "у каждого урока есть цвет группы", filled.length > 0 && noColor.length === 0, noColor.length + " без цвета");
      check("T-9-11", "легенда из 7 групп", D.querySelectorAll(".subj-legend > span").length === 7);
      check("T-9-11", "без учителя нет «— —»", !/— —/.test($("#main").textContent));
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
      check("T-6-22", "подсказка про Каракалпакстан", /Каракалпакстан/.test($("#main").textContent));
    }],
    ["T-2-02", "Классическая школа: направления №183 недоступны", async function(){
      await fresh(); await botToGen("criterial", "yes");
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
      await fresh(); await botToGen("criterial", "yes");
      setClass("class-direction", "10-A", "t:nd1");
      D.getElementById("langToggle").click();
      var allow = /^\d+-[A-Z]/, bad = [];
      [0, 4, 5, 6, 2, 1].forEach(function(i){ tab(i); bad = bad.concat(cyr($("#main").innerText, allow)); });
      D.querySelectorAll(".step-btn")[1].click(); click('[data-act="load-mode"][data-id="parallel"]');
      bad = bad.concat(cyr($("#main").innerText, allow), cyr($("#steps").innerText, allow), cyr($("#bot").innerText, allow));
      // предметы — из справочника ЕСП, у них свои узбекские названия; пользовательских данных тут нет
      check("T-1-04", "шаги мастера, «Нагрузка», «Справочники», бот", bad.length === 0, bad.slice(0, 8).join(" "));
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
      await fresh(); await botSchool(2);
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
      await fresh(); await botToGen("criterial", "yes"); botGen();
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
      await fresh(); await botToGen("criterial", "yes");
      var saved = W.localStorage.getItem("timetable_builder_v4");
      localStorage.setItem("timetable_builder_tab", JSON.stringify({ id: "other-tab", ts: Date.now() }));
      await sleep(50);
      check("T-12-14", "другая вкладка → «только для чтения»", D.body.classList.contains("is-readonly") && /другой вкладке/.test(D.getElementById("storageBar").textContent));
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
      await fresh(); await botToGen("criterial", "yes"); botGen();
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
      await botToGen("criterial", "yes");
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
      click('[data-imp-act="apply"]');
      tab(0); click('[data-act="preset-create"]'); click('[data-act="parallel-add"][data-grade="3"]');
      await setState(function(s){ s.curriculum[s.classes[0].id] = {}; Object.keys(s.curriculum).forEach(function(k){ s.curriculum[k] = {}; }); });
      tab(5); click('[data-act="goto-issue"]');
      D.getElementById("langToggle").click(); D.getElementById("langToggle").click();
      W.dispatchEvent(new W.ErrorEvent("error", { message: "тест", lineno: 1 }));
      W.confirm = function(){ return true; }; tab(6); click('[data-act="reset-all"]');
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
    }],

    /* ---------------- P1 ---------------- */
    ["T-1-12", "P1-3: мастер из 4 шагов с готовностью, «Справочники»", async function(){
      await fresh();
      var labels = Array.prototype.map.call(D.querySelectorAll(".step-btn .step-label"), function(x){ return x.textContent; });
      check("T-1-12", "4 шага: Классы и режим, Нагрузка, Расписание, Выгрузка", labels.join("|") === "Классы и режим|Нагрузка|Расписание|Выгрузка", labels.join("|"));
      await botSchool(1); botGen();
      var notes = Array.prototype.map.call(D.querySelectorAll(".step-btn"), function(b){ return b.className + ":" + b.querySelector(".step-note").textContent; });
      check("T-1-12", "готовность: классы ✓, нагрузка ! «нет учителей», расписание ✓", /st-done[^:]*:11 классов/.test(notes[0]) && /st-warn[^:]*:нет учителей/.test(notes[1]) && /st-done[^:]*:составлено/.test(notes[2]), notes.join(" | "));
      D.getElementById("refsBtn").click();
      check("T-1-12", "«Справочники»: вкладки учителей, предметов, кабинетов", /Справочники/.test($(".page-title").textContent) && D.querySelectorAll('[data-act="refs-tab"]').length === 3);
      click('[data-act="refs-back"]');
      check("T-1-12", "«← К мастеру» возвращает на прежний шаг", /Расписание/.test($(".page-title").textContent));
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
      D.querySelectorAll(".step-btn")[1].click();
      change($('[data-act="load-grade"]'), "5");
      var math = st().subjects.filter(function(x){ return x.name === "Математика"; })[0].id;
      var cell = $('[data-act="cur-teacher"][data-class="' + cls("5-A").id + '"][data-subj="' + math + '"]');
      check("T-6-26a", "пустая ячейка подсвечена", cell.classList.contains("need"));
      change(cell, "teachM");
      click('[data-act="load-fill-row"][data-subj="' + math + '"]');
      var pl = st().curriculum;
      check("T-6-26a", "«→ на всю параллель»: 5-B тоже Абдуллаева", pl[cls("5-A").id][math].teacherId === "teachM" && pl[cls("5-B").id][math].teacherId === "teachM");
      var box = $(".tload-box").textContent.replace(/\s+/g, " ");
      check("T-4-10", "часы учителя с учётом выходного: Абдуллаева 10 / 30 (без лимита × 5 рабочих дней)", /Абдуллаева Нодира ?10 \/ 30/.test(box), box.slice(0, 120));
      click('[data-act="load-known"]');
      var rod2 = st().subjects.filter(function(x){ return x.name === "Родной язык"; })[0].id;
      check("T-6-26a", "«Подставить известных»: родной язык 5-A — Каримова", st().curriculum[cls("5-A").id][rod2].teacherId === "teachR", toast());
      tab(2);
      check("T-4-11", "предметы учителя — из нагрузки", /Математика — 5-A, 5-B/.test($(".teach-table").textContent));
      var mx = $('[data-act="teach-max"][data-id="teachR"]'); change(mx, "");
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
      bot('[data-bot-act="open"]'); bot('[data-bot-act="restart"]');
      check("T-11-16", "«Начать заново» — снова первый вопрос о eMaktab", /Расписание уже есть в eMaktab/.test(botQ()));
      bot('[data-bot-act="path"][data-v="scratch"]');
      check("T-11-16", "повторный проход: число классов предложено текущее (5 класс — 2)", /2/.test($('#bot [data-bot-act="cnt"][data-g="5"]').parentNode.querySelector("strong").textContent));
      bot('[data-bot-act="cnt"][data-g="5"][data-v="-1"]'); bot('[data-bot-act="next"]');
      check("T-11-16", "вопрос перед удалением класса с правками", !D.getElementById("dialog").hidden && /5-B/.test(D.getElementById("dialog").textContent));
      click('[data-dialog="keep"]');
      check("T-11-16", "«Оставить» — 5-B на месте", !!cls("5-B"));
      await reload();
      check("T-11-16", "после перезагрузки — «продолжим с шага 2»", /шаге 2 из 5/.test($("#bot").textContent));
      bot('[data-bot-act="open"]');
      check("T-11-16", "продолжили с того же шага", /Как учатся классы/.test(botQ()));
    }],
    ["T-2-23", "P1-4: разное число уроков по дням", async function(){
      await fresh(); await botSchool(1);
      tab(0);
      change($('[data-act="parallel-periods"][data-grade="7"]'), "byday");
      change($('[data-act="parallel-pday"][data-grade="7"][data-day="0"]'), "7");
      change($('[data-act="parallel-pday"][data-grade="7"][data-day="5"]'), "4");
      var lbl = $('[data-act="parallel-periods"][data-grade="7"]').selectedOptions[0].text;
      check("T-2-23", "подпись профиля «Пн 7, Вт–Пт 6, Сб 4»", lbl === "Пн 7, Вт–Пт 6, Сб 4", lbl);
      tab(5); click('[data-act="generate"]');
      var g = cells("7-A");
      check("T-2-23", "сетка 7-A по дням: 7, 6, 6, 6, 6, 4", g.map(function(r){ return r.length; }).join() === "7,6,6,6,6,4", g.map(function(r){ return r.length; }).join());
      click('[data-act="sched-select-class"][data-id="' + cls("7-A").id + '"]');
      check("T-2-23", "лишние уроки субботы показаны как «нет урока»", D.querySelectorAll("td.off-cell").length >= 3);
      check("T-2-23", "0 окон и 0 неразмещённых", st().schedule.unplacedCount === 0 && windowsOf(st().schedule.byClass).windows === 0);
    }],
    ["T-8-28", "P1-5: правила предметов соблюдаются", async function(){
      await fresh(); await botSchool(1);
      tab(1);
      function opt(n){ return Array.prototype.filter.call($("#rule-subj").options, function(o){ return o.text === n; })[0].value; }
      $("#rule-subj").value = opt("Воспитание"); $("#rule-kind").value = "fixed"; $("#rule-day").value = "0"; $("#rule-period").value = "1"; click('[data-act="rule-add"]');
      $("#rule-subj").value = opt("Физическое воспитание"); $("#rule-kind").value = "not"; $("#rule-periods").value = "1"; click('[data-act="rule-add"]');
      check("T-8-28", "правила в списке", /всегда Пн, 1-й урок/.test($(".rule-list").textContent) && /не на уроках: 1/.test($(".rule-list").textContent));
      tab(5); click('[data-act="generate"]');
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
      tab(5); click('[data-act="generate"]');
      var li = $(".unplaced-list li").textContent;
      check("T-8-20", "причина: у учителя выходной", /выходной/.test(li), li);
      // учитель выходит на работу — уроки можно поставить вручную
      await setState(function(s){ s.teachers.filter(function(x){ return x.id === "teachZ"; })[0].daysOff = [false,false,false,false,false,false]; });
      tab(5);
      var before = st().schedule.unplacedCount;
      click('[data-act="place-start"]');
      var can = D.querySelectorAll(".can-place").length;
      check("T-9-10", "«Поставить» подсвечивает допустимые ячейки", can > 0, "ячеек: " + can);
      click('[data-act="place-here"]');
      check("T-9-10", "урок поставлен, неразмещённых меньше", st().schedule.unplacedCount === before - 1);
      var tds = Array.prototype.slice.call(D.querySelectorAll('td[draggable="true"]'));
      var a = tds[0], b = tds[2], va = a.querySelector("select").value, vb = b.querySelector("select").value;
      var dt = new W.DataTransfer();
      a.dispatchEvent(new W.DragEvent("dragstart", { bubbles: true, dataTransfer: dt }));
      b.dispatchEvent(new W.DragEvent("dragover", { bubbles: true, cancelable: true, dataTransfer: dt }));
      b.dispatchEvent(new W.DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: dt }));
      tds = Array.prototype.slice.call(D.querySelectorAll('td[draggable="true"]'));
      check("T-9-11a", "перетаскивание меняет уроки местами", tds[0].querySelector("select").value === vb && tds[2].querySelector("select").value === va);
      // ручная правка урока с типом кабинета получает кабинет
      await setState(function(s){
        var inf = s.subjects.filter(function(x){ return x.name === "Информатика и информационные технологии"; })[0]; inf.roomType = "комп";
        s.rooms = [{ id: "roomK", name: "Комп-1", capacity: 30, type: "комп" }];
      });
      tab(5); click('[data-act="sched-select-class"][data-id="' + cls("5-A").id + '"]');
      var infId = st().subjects.filter(function(x){ return x.name === "Информатика и информационные технологии"; })[0].id;
      var sel = Array.prototype.filter.call(D.querySelectorAll('select[data-act="cell-change"]'), function(x){ return x.value && x.value !== infId && Array.prototype.some.call(x.options, function(o){ return o.value === infId; }); })[0];
      var sd = sel.getAttribute("data-day"), sp = sel.getAttribute("data-period");
      change(sel, infId);
      check("T-9-04", "ручная правка: кабинет подобран", st().schedule.byClass[cls("5-A").id][sd][sp].lessons[0].roomId === "roomK");
    }],
    ["T-6-28", "P1-7: пустые планы, копирование на параллель, план не по источнику", async function(){
      await fresh(); await botSchool(2);
      await setState(function(s){ var c = s.classes.filter(function(x){ return x.name === "6-B"; })[0]; s.curriculum[c.id] = {}; });
      D.querySelectorAll(".step-btn")[1].click(); click('[data-act="load-fill-empty"]');
      check("T-6-28", "«Заполнить пустые планы» заполнил 6-B", planHours("6-B") > 0 && /Заполнено планов: 1/.test(toast()), toast());
      tab(4); click('[data-act="cur-select-class"][data-id="' + cls("7-A").id + '"]');
      await setState(function(s){ var c = s.classes.filter(function(x){ return x.name === "7-A"; })[0]; var k = Object.keys(s.curriculum[c.id])[0]; s.curriculum[c.id][k].hours = 9; s.teachers.push({ id: "teachC", name: "Копия Учитель", subjects: [], maxPerDay: null, daysOff: [false,false,false,false,false,false] }); s.curriculum[c.id][k].teacherId = "teachC"; });
      tab(4); click('[data-act="cur-select-class"][data-id="' + cls("7-A").id + '"]');
      click('[data-act="cur-copy"]'); click('[data-dialog="without"]');
      var pa = st().curriculum[cls("7-A").id], pb = st().curriculum[cls("7-B").id], k0 = Object.keys(pa)[0];
      check("T-6-29", "план 7-A скопирован в 7-B без учителей", pb[k0] && pb[k0].hours === 9 && !pb[k0].teacherId && Object.keys(pb).length === Object.keys(pa).length);
      tab(0); change(classSel("class-direction", "8-A"), "v:law");
      tab(4); click('[data-act="cur-select-class"][data-id="' + cls("8-A").id + '"]');
      check("T-2-11", "подсказка: план по другому документу и кнопка обновления", /План заполнен по другому документу/.test($("#main").textContent));
      click('.warn-box [data-act="cur-autofill"]');
      check("T-2-11", "после обновления подсказки нет", !/План заполнен по другому документу/.test($("#main").textContent));
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
      tab(0); click('[data-act="bells-init"]');
      var bt = $(".bell-times").textContent;
      check("T-2-25", "звонки 1-й смены: 1. 08:00–08:45, 4-й после большой перемены", /1\. 08:00–08:45/.test(bt) && /4\. 10:55–11:40/.test(bt), bt);
      tab(5); click('[data-act="sched-select-class"][data-id="' + cls("5-A").id + '"]');
      check("T-2-25", "время урока в сетке класса", /08:00–08:45/.test($("#main table tbody tr th").textContent));
      tab(6);
      var printed = null; W.print = function(){ printed = D.getElementById("print").innerHTML; };
      click('[data-act="print-all"]');
      check("T-10-07", "шахматка всех классов собрана для печати", !!printed && /Расписание всех классов/.test(printed) && /5-A/.test(printed));
      var log = stubDownloads(); click('[data-act="export-xlsx"]'); await sleep(30);
      check("T-10-09", "первый лист Excel — «Все классы»", log.sheets && log.sheets[0] === "Все классы", log.sheets && log.sheets.slice(0, 3).join());
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
