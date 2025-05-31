**Explorer Watcher** is a convenient Windows system utility that runs in the system tray and automatically monitors open Explorer windows, remembering the most recently opened folders and allowing quick restoration when needed.

## Key Features

* **Automatic monitoring** of open Explorer windows every 20 seconds.
* **Saving paths** of the most recently opened folders to a file.
* **Quick access** to recent folders via the tray icon context menu.
* **Restoring Explorer windows** — reopening previously saved folders.
* **Countdown timer** showing time until the next check and update of open windows.
* Ability to **view and edit** the saved list of paths using a standard text editor.
* **Clearing the saved list** with a single click.
* Support for activating and restoring minimized Explorer windows.

## Usage

1. Run the program — the icon will appear in the system tray.
2. Right-click the icon to open the context menu. You will see a list of recently opened folders that you can open or activate with one click.
3. Use the **Restore Windows** option to reopen all saved Explorer windows.
4. Open or clear the saved list of paths via the corresponding menu options.
5. To exit the program, select **Exit**.

## Requirements

* Windows 7 or later.
* .NET Framework (version depends on the build, e.g., .NET Framework 4.7.2 or higher).

## Implementation Details

* Uses the COM interface `SHDocVw.ShellWindows` to get the list of Explorer windows.
* Developed in C# using WinForms for the tray icon and context menu.
* Saved paths file location: `%TEMP%\ExplorerWatcher\last_paths.txt`.

---

**Explorer Watcher** — удобная системная утилита для Windows, которая работает в системном трее и автоматически отслеживает открытые окна проводника (Explorer), запоминая последние открытые папки и позволяя быстро восстанавливать их при необходимости.

## Основные возможности

* **Автоматический мониторинг** открытых окон проводника с периодом 20 секунд.
* **Запоминание путей** последних открытых папок и сохранение их в файл.
* **Быстрый доступ** к последним папкам через контекстное меню иконки в трее.
* **Восстановление окон проводника** — повторное открытие ранее запомненных папок.
* **Отсчет времени** до следующей проверки и обновления списка открытых окон.
* Возможность **просмотра и редактирования** сохраненного списка путей через стандартный текстовый редактор.
* **Очистка сохранённого списка** одним кликом.
* Поддержка активации и восстановления свернутых окон проводника.

## Использование

1. Запустите программу — иконка появится в системном трее.
2. Щелкните правой кнопкой по иконке для вызова контекстного меню. В меню вы увидите список последних открытых папок, которые можно открыть или активировать одним кликом.
3. Используйте пункт **Restore Windows** для восстановления всех сохранённых окон.
4. Откройте или очистите сохранённый список путей через соответствующие пункты меню.
5. Для выхода из программы выберите пункт **Exit**.

## Требования

* Windows 7 и выше.
* .NET Framework (версия в зависимости от сборки, например .NET Framework 4.7.2 или выше).

## Особенности реализации

* Используется COM-интерфейс `SHDocVw.ShellWindows` для получения списка окон проводника.
* Программа реализована на C# с использованием WinForms для иконки в трее и контекстного меню.
* Путь к сохранённому файлу — `%TEMP%\ExplorerWatcher\last_paths.txt`.
