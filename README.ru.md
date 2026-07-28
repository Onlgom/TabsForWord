# TabsForWord — Вкладки для Microsoft Word

*In English: [README.md](README.md).*

Надстройка для настольного Microsoft Word под Windows, добавляющая панель
вкладок открытых документов: клик по вкладке переключает на соответствующий
документ, как вкладки в браузере.

Начиная с версии 1.6.0 панель по умолчанию встроена прямо в окно Word —
под лентой, без служебной полосы Custom Task Pane (решение [ADR-015](docs/DECISIONS.md);
как это устроено — [docs/NATIVE_TAB_HOST.md](docs/NATIVE_TAB_HOST.md)). Классический
режим CTP остался запасным: включается файлом `Classic-mode.cmd` из папки
установки и включается сам, если native-хост не смог работать.

С версии 1.7.0 язык интерфейса берётся у самого Word (русский Word — русские
вкладки, любой другой — английские), выбрать вручную можно в окне настроек
(решение ADR-016).

## Требования

- Windows 10 версии 1903 (май 2019) или новее, либо Windows 11 —
  на них .NET Framework 4.8 уже встроен в систему.
- Настольный Microsoft Word для Windows (Microsoft 365, Word 2016 и новее).
  Не работает с Word Online, Word для macOS и мобильным Word.
- Права администратора не требуются: установка только для текущего пользователя.

## Для пользователя

- Установка: см. [docs/INSTALL_RU.md](docs/INSTALL_RU.md)
- Удаление: см. [docs/UNINSTALL_RU.md](docs/UNINSTALL_RU.md)
- Готовые файлы для установки: страница [Releases](../../releases)
- Честный список ограничений: [docs/KNOWN_ISSUES.md](docs/KNOWN_ISSUES.md)

## Для разработчика

- Архитектура: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)
- Исследование: [docs/RESEARCH.md](docs/RESEARCH.md)
- Решения: [docs/DECISIONS.md](docs/DECISIONS.md)
- Сборка: `scripts/build-release.ps1`

Документы в `docs/` — на английском (кроме INSTALL_RU / UNINSTALL_RU).

## Лицензия

MIT — файл [LICENSE](LICENSE). Можно пользоваться, изменять и включать в свои
продукты, в том числе платные; нужно лишь сохранить строку об авторстве.
Гарантий никаких: надстройка работает внутри Word рядом с вашими документами,
и вы используете её на свой страх и риск.
