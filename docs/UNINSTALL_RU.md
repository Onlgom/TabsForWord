# Удаление «Вкладки для Microsoft Word»

(In English: [UNINSTALL_EN.md](UNINSTALL_EN.md).)

## Обычное удаление (1 минута)

1. **Закройте Microsoft Word** (все окна).
2. Дважды щёлкните файл **`Uninstall.cmd`** (лежит рядом с `Install.cmd`
   в папке, из которой вы устанавливали программу).
3. Появится сообщение «УДАЛЕНИЕ ЗАВЕРШЕНО». Готово.

Удаляется всё: файлы программы, записи в реестре, журналы.
Ваши документы Word никак не затрагиваются.

## Если файла Uninstall.cmd больше нет

Удалить можно вручную:

1. Закройте Word.
2. Нажмите Win+R, введите `regedit`, Enter.
3. Удалите разделы (правый клик → Удалить):
   - `HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\TabsForWord.Connect`
   - `HKEY_CURRENT_USER\Software\Classes\TabsForWord.Connect`
   - `HKEY_CURRENT_USER\Software\Classes\TabsForWord.TabStripControl`
   - `HKEY_CURRENT_USER\Software\Classes\CLSID\{A3B7C9D1-5E2F-4A6B-8C0D-1F2E3D4C5B6A}`
   - `HKEY_CURRENT_USER\Software\Classes\CLSID\{D4E5F6A7-8B9C-4D0E-A1B2-C3D4E5F6A7B8}`
   - те же два раздела CLSID в
     `HKEY_CURRENT_USER\Software\Classes\WOW6432Node\CLSID`, если они есть
     (они нужны, чтобы надстройку видел 32-битный Word)
4. Удалите папку `%LOCALAPPDATA%\TabsForWord`
   (вставьте путь в адресную строку Проводника).

## Временно отключить (не удаляя)

Файл → Параметры → Надстройки → «Управление: Надстройки COM» → «Перейти…» →
снимите галочку с «Вкладки для Word» → ОК. Обратно — так же, поставив галочку.
