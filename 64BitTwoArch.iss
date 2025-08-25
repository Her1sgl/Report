; -- AutoReport_Setup_Script.iss --
; Скрипт Inno Setup для приложения Auto Report

[Setup]
; Основная информация о приложении
AppName=Auto Report
AppVersion=1.6
; Имя папки в меню Пуск и Program Files
AppPublisher=My Company
AppCopyright=Copyright (C) 2024 My Company

; Имя установочного файла
OutputBaseFilename=AutoReport_Installer_v1.6
; Каталог установки по умолчанию
DefaultDirName={autopf}\AutoReport
; Группа в меню Пуск
DefaultGroupName=Auto Report
; Имя файла лицензии (если есть, удалите комментарий и укажите путь)
; LicenseFile=license.txt
; Иконка для установщика
SetupIconFile=icon.ico
; Имя файла, который будет проверяться для определения, установлено ли приложение
UninstallDisplayName=Auto Report
; Иконка для записи в "Программы и компоненты"
UninstallDisplayIcon={app}\icon.ico

; Запрашивать права администратора (рекомендуется для установки в Program Files)
PrivilegesRequired=admin
; Показывать страницу выбора каталога установки
DisableDirPage=no
; Показывать страницу выбора группы в меню Пуск
DisableProgramGroupPage=no
; Компилировать в 64-битный установщик, если приложение 64-битное
; ArchitecturesInstallIn64BitMode=x64
; Разрешить установку только для текущего пользователя (раскомментируйте, если нужно)
; PrivilegesRequired=lowest
; OutputDir - куда будет сохранен .exe установщика (по умолчанию рядом со скриптом)
OutputDir=.

; Минимальная версия Windows (пример для Windows 10)
MinVersion=10.0

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
; Задача: создать ярлык на рабочем столе
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Источник - путь к вашим файлам, DestDir - куда копировать в установке
; Флаг recursesubdirs копирует подкаталоги, если они есть
; Флаг replacesameversionifnewer заменяет файл, если новая версия
Source: "ReportUpdater.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "config.json"; DestDir: "{app}"; Flags: ignoreversion
; Иконка
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion
; Если у вас есть другие файлы (библиотеки, данные и т.д.), добавьте их аналогично:
; Source: "otherfile.dll"; DestDir: "{app}"; Flags: ignoreversion
; Source: "data\*"; DestDir: "{app}\data"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Создание ярлыков
; Ярлык в меню Пуск
Name: "{group}\Auto Report"; Filename: "{app}\ReportUpdater.exe"; IconFilename: "{app}\icon.ico"
; Ярлык в меню Пуск для удаления программы
Name: "{group}\{cm:UninstallProgram,Auto Report}"; Filename: "{uninstallexe}"
; Ярлык на рабочем столе (создается, если выбрана задача desktopicon)
Name: "{autodesktop}\Auto Report"; Filename: "{app}\ReportUpdater.exe"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon

[Run]
; Опционально: запустить приложение после установки
Filename: "{app}\ReportUpdater.exe"; Description: "{cm:LaunchProgram,Auto Report}"; Flags: nowait postinstall skipifsilent
