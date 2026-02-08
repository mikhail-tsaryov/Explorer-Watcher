# Explorer Watcher

**Explorer Watcher** is a convenient Windows system utility that runs in the system tray and automatically monitors open Explorer windows, remembering the most recently opened folders and allowing quick restoration when needed. The program now stores up to 10 recent saves and optionally displays a progress bar during monitoring.

## Key Features

* **Automatic monitoring** of open Explorer windows every 60 seconds.  
* **Recent folder history** — stores up to 10 most recent saved lists of paths.  
* **Quick access** to recent folders via the tray icon context menu.  
* **Restoring Explorer windows** — reopening previously saved folders.  
* **Countdown timer** showing time until the next check and update of open windows.  
* **Optional progress bar** to indicate monitoring progress.  
* Ability to **view and edit** saved path lists using a standard text editor.  
* **Clearing history** and current saved paths with a single click.

## Indicators

* **Active (blue icon)** – Explorer windows are open.  
* **Inactive (red icon)** – No Explorer windows detected.  

## Usage

1. Run the program — the icon will appear in the system tray.  
2. Right-click the icon to open the context menu. You will see a list of recently opened folders that you can open or activate with one click.  
3. Use the **Restore Windows** option to reopen all saved Explorer windows.  
4. Switch between or clear recent saved histories via the corresponding menu options.  
5. The optional progress bar shows the current monitoring progress (can be enabled in settings).  
6. To exit the program, select **Exit**.  

## Requirements

* Windows 7 or later.  
* .NET Framework (version depends on the build, e.g., .NET Framework 4.7.2 or higher).  

## Implementation Details

* Uses the COM interface `SHDocVw.ShellWindows` to get the list of Explorer windows.  
* Developed in C# using WinForms for the tray icon and context menu.  
* Paths are saved to `%TEMP%\ExplorerWatcher\EW_paths_#.txt`, where `#` is the save number (1–10).

---

# Explorer Watcher

**Explorer Watcher** — удобная системная утилита для Windows, которая работает в системном трее и автоматически отслеживает открытые окна проводника (Explorer), запоминая последние открытые папки и позволяя быстро восстанавливать их при необходимости. Теперь программа хранит историю до 10 последних сохранений и может отображать опциональный прогресс-бар для мониторинга.

## Основные возможности

* **Автоматический мониторинг** открытых окон проводника с периодом 60 секунд.  
* **История последних папок** — хранение до 10 последних сохранённых списков путей.  
* **Быстрый доступ** к последним папкам через контекстное меню иконки в трее.  
* **Восстановление окон проводника** — повторное открытие ранее сохранённых папок.  
* **Отсчет времени** до следующей проверки и обновления списка открытых окон.  
* **Опциональный прогресс-бар**, показывающий прогресс мониторинга.  
* Возможность **просмотра и редактирования** сохранённых списков путей через стандартный текстовый редактор.  
* **Очистка истории** и текущего списка путей одним кликом.

## Индикаторы

* **Активно (синяя иконка)** — окна проводника открыты.  
* **Неактивно (красная иконка)** — проводник не обнаружен.  

## Использование

1. Запустите программу — иконка появится в системном трее.  
2. Щелкните правой кнопкой по иконке для вызова контекстного меню. В меню вы увидите список последних открытых папок, которые можно открыть или активировать одним кликом.  
3. Используйте пункт **Restore Windows** для восстановления всех сохранённых окон.  
4. Переключайте или очищайте историю последних сохранений через соответствующие пункты меню.  
5. Опциональный прогресс-бар отображает текущий прогресс мониторинга (включается в настройках).  
6. Для выхода из программы выберите пункт **Exit**.  

## Требования

* Windows 7 и выше.  
* .NET Framework (версия в зависимости от сборки, например .NET Framework 4.7.2 или выше).  

## Особенности реализации

* Используется COM-интерфейс `SHDocVw.ShellWindows` для получения списка окон проводника.  
* Программа реализована на C# с использованием WinForms для иконки в трее и контекстного меню.  
* Пути сохраняются в файлы `%TEMP%\ExplorerWatcher\EW_paths_#.txt`, где `#` — номер сохранения (1–10).

