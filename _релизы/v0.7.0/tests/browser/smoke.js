/* Сквозные проверки index.html в браузере (без Node): страница открывается во фрейме,
   сценарии нажимают кнопки как пользователь и сверяют данные в localStorage.
   ID сценариев — из docs/ПМИ_конструктор_расписания.md. Промежуточный вариант до Playwright/Vitest
   (ARCHITECTURE, ADR-11): набор переносится туда без изменения сценариев.
   Внимание: тесты очищают localStorage этого адреса (localhost). */
(function(){
  var frame, W, D, results = [];
  function sleep(ms){ return new Promise(function(r){ setTimeout(r, ms); }); }
  async function fresh(){
    frame.contentWindow.localStorage.clear();
    frame.src = "../../index.html?t=" + Date.now();
    await new Promise(function(r){ frame.onload = r; });
    W = frame.contentWindow; D = W.document;
    await sleep(100);
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
