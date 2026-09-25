# VBA CLI 1.0.8

CLI для разработки Excel VBA из VS Code с синхронизацией через Excel COM.

## Исправления в 1.0.8

- `vba import` импортирует `.bas` и `.cls` через `VBComponents.Add()` с явным типом: `.bas` → Standard Module, `.cls` → Class Module. Это исключает ситуацию, когда `Example.cls` попадает в Modules.
- Убран batch-импорт через `ConvertFrom-Json`: каждый файл передаётся в PowerShell отдельной строкой. Это устраняет ошибку Windows PowerShell вида `System.Object[]`.
- UTF-8 исходники читаются строго как UTF-8 и передаются в `CodeModule.AddFromString()` как Unicode, поэтому русский текст не должен превращаться в `Ð...`.
- `.frm` импортируется штатным `VBComponents.Import()` с временным ANSI-файлом, потому что UserForm содержит дизайнерские данные.
- При `vba init` PowerShell-консоль показывается на время создания Excel и первичного импорта, а после завершения автоматически закрывается.
- Исправлен старый баг с `param(...)`: временный PowerShell-скрипт не получает никаких строк перед `param`.

## Установка

```powershell
npm install
npm link
```

## Создание проекта

```powershell
vba init test
```

CLI спросит:

```text
Создать стартовые VBA-модули? [Y/n]:
```

После создания `test.xlsm` выполняется первичный импорт, открывается проект в VS Code и запускается watcher.

## Команды

```powershell
vba init [name]
vba dev
vba import
vba export
vba build
vba clean
vba help
```

## Структура

```text
test/
├── test.xlsm
├── vba.json
├── src/
│   ├── Modules/
│   │   └── *.bas
│   ├── Classes/
│   │   └── *.cls
│   └── Forms/
│       └── *.frm
└── build/
```

## Импорт

`vba import` обрабатывает файлы последовательно. Для каждого файла:

```text
src/Modules/modMain.bas -> Standard Module
src/Classes/Example.cls -> Class Module
src/Forms/UserForm1.frm -> UserForm
```

Для `.bas` и `.cls` компонент сначала удаляется (если существует), затем создаётся через `VBComponents.Add()`. Код читается из UTF-8 и вставляется через `CodeModule.AddFromString()`.

## Требования

- Windows
- Microsoft Excel
- Node.js 18+
- Разрешён программный доступ к VBA Project Object Model в Excel

> Если Excel блокирует доступ к `VBProject`, включите в Excel: Trust Center → Macro Settings → Trust access to the VBA project object model.
