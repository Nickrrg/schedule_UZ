# Разбор выгрузок eMaktab (черновик импорта)

`analyze.js` — разборщики четырёх выгрузок eMaktab на SheetJS (та же библиотека, что у страницы для Excel) — черновик и инструмент анализа. С 0.10 импорт сетки (A) и нагрузки (B) встроен в `index.html` (блок «ИМПОРТ ИЗ eMaktab», ТЗ ФТ-16); описание форматов и выводы — `docs/Пилот_импорт_eMaktab.md`.

**Внимание:** в `дано/` файлы нагрузки перепутаны: `1TeachingLoad…` — от большой школы (к `Schedule_0-12_21-09-2026_27-09-2026 (1).xls`), `TeachingLoad…` — от школы 1 (к `1Schedule…` / `Schedule_0-12_14-09…`).

**Обезличенные копии для тестов** — `tests/fixtures/emaktab/school1_grid.xls` и `school1_load.xls` (ФИО → «Учитель001 А.Б.», формат и объединённые ячейки сохранены). Сделаны в браузере на SheetJS: сетка переписана в BIFF8 с заменой колонки «Учитель», в нагрузке заменены текстовые узлы с ФИО (в той же кодировке `&#x…;`); файлы сохранены через `POST /save` сервера `tools/npa/devserver.py`.

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
