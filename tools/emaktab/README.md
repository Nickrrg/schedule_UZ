# Разбор выгрузок eMaktab (черновик импорта)

`analyze.js` — разборщики четырёх выгрузок eMaktab на SheetJS (та же библиотека, что у страницы для Excel). Это черновик будущего импорта в `index.html`; описание форматов и выводы — `docs/Пилот_импорт_eMaktab.md`.

`conftext.py` — превращает выгрузку Confluence (`.doc` = MIME-письмо с HTML) в текст: `py conftext.py "файл.doc" out.txt`.

## Как прогнать на реальных файлах

Реальные выгрузки лежат в `дано/` (есть ФИО — **не загружать в репозиторий на GitHub**).

1. `py -m http.server 8767 --bind 127.0.0.1` из корня проекта (конфигурация `static` в `.claude/launch.json`).
2. Открыть любую страницу на `http://localhost:8767/`, в консоли подключить SheetJS и `tools/emaktab/analyze.js`:
   ```js
   await new Promise(r=>{ const s=document.createElement('script'); s.src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'; s.onload=r; document.head.appendChild(s); });
   await new Promise(r=>{ const s=document.createElement('script'); s.src='/tools/emaktab/analyze.js'; s.onload=r; document.head.appendChild(s); });
   const grid = An.parseGrid(await An.wb('/дано/<сетка>.xls')), load = An.parseLoad(await An.wb('/дано/<нагрузка>.xls'));
   An.fullNames(grid, load); An.stats(grid);
   ```
3. Архив схем классов (`.rar`) распаковывается встроенным `C:\Windows\System32\tar.exe -xf файл.rar`.
