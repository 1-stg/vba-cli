#!/usr/bin/env node

"use strict";

const fs = require("fs");
const path = require("path");
const os = require("node:os");
const readline = require("readline");
const { spawn, spawnSync } = require("child_process");
const CLI_VERSION = require("./package.json").version;

let chokidar;
try {
    chokidar = require("chokidar");
} catch {
    chokidar = null;
}

const VBA_EXTENSIONS = [".bas", ".cls", ".frm"];
const COMPONENT_TYPES = {
    ".bas": 1, // vbext_ct_StdModule
    ".cls": 2, // vbext_ct_ClassModule
    ".frm": 3  // vbext_ct_MSForm
};

const COLORS = {
    reset: "\x1b[0m",
    red: "\x1b[31m",
    green: "\x1b[32m",
    yellow: "\x1b[33m",
    cyan: "\x1b[36m",
    gray: "\x1b[90m"
};

function color(text, code) {
    return `${code}${text}${COLORS.reset}`;
}

function info(message) {
    console.log(color(`ℹ ${message}`, COLORS.cyan));
}

function success(message) {
    console.log(color(`✔ ${message}`, COLORS.green));
}

function warning(message) {
    console.log(color(`⚠ ${message}`, COLORS.yellow));
}

function error(message) {
    console.error(color(`✖ ${message}`, COLORS.red));
}

function separator() {
    console.log("");
}

function createDir(dir) {
    fs.mkdirSync(dir, { recursive: true });
}

function writeUtf8(filePath, content) {
    createDir(path.dirname(filePath));
    fs.writeFileSync(filePath, content, "utf8");
}

function readUtf8(filePath) {
    return fs.readFileSync(filePath, "utf8");
}

function isVBAFile(filePath) {
    return VBA_EXTENSIONS.includes(path.extname(filePath).toLowerCase());
}

function normalizePath(filePath) {
    return path.normalize(path.resolve(filePath)).toLowerCase();
}

function quotePowerShellArgument(value) {
    return `'${String(value).replace(/'/g, "''")}'`;
}

function runPowerShell(script, args = [], options = {}) {
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "vba-cli-"));
    const scriptPath = path.join(tempDir, "script.ps1");

    // IMPORTANT:
    // `param(...)` must be the first executable construct in a PowerShell
    // script. Do not prepend any statements here: several CLI scripts use
    // a top-level param block. The previous implementation prepended
    // encoding/error-handling statements and caused PowerShell to interpret
    // `param` as a command, producing: "Имя \"param\" не распознано...".
    // Keep the script byte-for-byte intact and let Node handle the process
    // output as UTF-8.
    const psScript = "\uFEFF" + script.replace(/^\uFEFF/, "");

    fs.writeFileSync(scriptPath, psScript, { encoding: "utf8" });

    const encodedArgs = [
        "-NoLogo",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        scriptPath,
        ...args.map(String)
    ];

    const result = spawnSync("powershell.exe", encodedArgs, {
        windowsHide: options.windowsHide !== undefined ? options.windowsHide : true,
        encoding: "utf8",
        maxBuffer: 20 * 1024 * 1024
    });

    try {
        fs.rmSync(tempDir, { recursive: true, force: true });
    } catch {
        // Ignore temporary directory cleanup errors.
    }

    const stdout = (result.stdout || "").trim();
    const stderr = (result.stderr || "").trim();

    if (result.error) {
        throw new Error(result.error.message);
    }

    if (result.status !== 0) {
        const message = [stdout, stderr].filter(Boolean).join("\n").trim();
        throw new Error(cleanPowerShellError(message || `PowerShell exited with code ${result.status}`));
    }

    return stdout;
}

function cleanPowerShellError(message) {
    if (!message) {
        return "Unknown PowerShell error";
    }

    // Remove CLIXML wrappers if an older PowerShell host still produced them.
    let text = message
        .replace(/^#<\s*CLIXML\s*/i, "")
        .replace(/<Objs[\s\S]*?<\/Objs>/gi, "")
        .trim();

    text = text
        .replace(/_x000D__x000A_/g, "\n")
        .replace(/_x000A_/g, "\n")
        .replace(/_x000D_/g, "\r")
        .replace(/<[^>]+>/g, " ")
        .trim();

    return text || message;
}

function ask(questionText) {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    return new Promise(resolve => {
        rl.question(questionText, answer => {
            rl.close();
            resolve(answer.trim());
        });
    });
}

function getProjectPaths(cwd = process.cwd()) {
    const configPath = path.join(cwd, "vba.json");

    if (!fs.existsSync(configPath)) {
        throw new Error(
            "Не найден vba.json. Запустите команду из корня VBA-проекта."
        );
    }

    let config;

    try {
        config = JSON.parse(readUtf8(configPath));
    } catch (err) {
        throw new Error(`Не удалось прочитать vba.json: ${err.message}`);
    }

    if (!config.workbook) {
        throw new Error("В vba.json не указан параметр workbook.");
    }

    const sourceDir = path.resolve(cwd, config.source || "src");
    const workbookPath = path.resolve(cwd, config.workbook);
    const buildDir = path.resolve(cwd, config.build || "build");

    return {
        root: cwd,
        config,
        configPath,
        sourceDir,
        workbookPath,
        buildDir
    };
}

function getVBAFiles(dir) {
    if (!fs.existsSync(dir)) {
        return [];
    }

    const result = [];

    function walk(currentDir) {
        for (const entry of fs.readdirSync(currentDir, { withFileTypes: true })) {
            const fullPath = path.join(currentDir, entry.name);

            if (entry.isDirectory()) {
                walk(fullPath);
                continue;
            }

            if (entry.isFile() && isVBAFile(fullPath)) {
                result.push(fullPath);
            }
        }
    }

    walk(dir);
    return result.sort();
}

function componentNameFromFile(filePath) {
    return path.basename(filePath, path.extname(filePath));
}

function getCategoryByExtension(extension) {
    switch (extension.toLowerCase()) {
        case ".bas":
            return "Modules";
        case ".cls":
            return "Classes";
        case ".frm":
            return "Forms";
        default:
            return "Other";
    }
}

function createExcelFile(workbookPath) {
    const script = `
param(
    [string]$WorkbookPath
)

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    $workbook.SaveAs($WorkbookPath, 52)
    $workbook.Close($true)

    $excel.Quit()

    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    Write-Output "OK"
}
catch {
    [Console]::Error.WriteLine(("CREATE_WORKBOOK_ERROR: " + $_.Exception.Message))
    exit 1
}
`;

    runPowerShell(script, [workbookPath], { windowsHide: false });
}

function buildStarterFiles(projectPath) {
    const src = path.join(projectPath, "src");
    const modules = path.join(src, "Modules");
    const classes = path.join(src, "Classes");
    const forms = path.join(src, "Forms");

    createDir(modules);
    createDir(classes);
    createDir(forms);

    const modMain = `Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()

    MsgBox "VBA Привет123142!", vbInformation, "VBA CLI"

End Sub
`;

    const modUtils = `Attribute VB_Name = "modUtils"
Option Explicit

Public Function AddNumbers(ByVal a As Double, ByVal b As Double) As Double

    AddNumbers = a + b

End Function
`;

    const exampleClass = `VERSION 1.0 CLASS
BEGIN
    MultiUse = -1
END

Attribute VB_Name = "Example"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private mName As String

Public Property Get Name() As String

    Name = mName

End Property

Public Property Let Name(ByVal value As String)

    mName = value

End Property

Public Function SayHello() As String

    SayHello = "Привет, " & mName & "!"

End Function
`;

    writeUtf8(path.join(modules, "modMain.bas"), modMain);
    writeUtf8(path.join(modules, "modUtils.bas"), modUtils);
    writeUtf8(path.join(classes, "Example.cls"), exampleClass);
}

function createProjectFiles(projectPath, projectName) {
    const config = {
        name: projectName,
        workbook: `${projectName}.xlsm`,
        source: "src",
        build: "build"
    };

    writeUtf8(
        path.join(projectPath, "vba.json"),
        JSON.stringify(config, null, 2) + "\n"
    );

    writeUtf8(
        path.join(projectPath, ".gitignore"),
        [
            "build/",
            ".DS_Store",
            "Thumbs.db",
            "*.tmp"
        ].join("\n") + "\n"
    );

    writeUtf8(
        path.join(projectPath, "README.md"),
        `# ${projectName}

VBA project managed by VBA CLI.

## Commands

\`\`\`powershell
vba dev
vba import
vba export
vba build
vba clean
\`\`\`

Source code is stored in \`src/\`.
`
    );
}

function openInVSCode(projectPath) {
    info("Открытие проекта в VS Code...");

    const absolutePath = path.resolve(projectPath);

    if (process.platform !== "win32") {
        try {
            const child = spawn("code", ["--reuse-window", absolutePath], {
                cwd: absolutePath,
                detached: true,
                stdio: "ignore",
                shell: true,
                windowsHide: true
            });
            child.unref();
            success(`VS Code открыт: ${absolutePath}`);
            return true;
        } catch (err) {
            warning(`Не удалось автоматически открыть VS Code: ${err.message}`);
            warning(`Откройте вручную: ${absolutePath}`);
            return false;
        }
    }

    // Do NOT use spawnSync here. VS Code may keep the command process alive,
    // which made `vba init` appear to hang forever. `start` launches VS Code
    // independently and returns immediately.
    try {
        const child = spawn(
            "cmd.exe",
            ["/d", "/c", "start", "", "code.cmd", "--reuse-window", absolutePath],
            {
                cwd: absolutePath,
                detached: true,
                stdio: "ignore",
                windowsHide: true
            }
        );
        child.unref();
        success(`VS Code открыт: ${absolutePath}`);
        return true;
    } catch (err) {
        warning(`Не удалось автоматически открыть VS Code: ${err.message}`);
        warning(`Откройте вручную: ${absolutePath}`);
        return false;
    }
}
function buildPowerShellImportScript() {
    return String.raw`
param(
    [string]$WorkbookPath,
    [string]$SourcePath,
    [string]$ComponentName,
    [int]$ExpectedType
)

function Get-ExcelApplication {
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        return $excel
    }
}

function Get-WorkbookByPath($excel, $targetPath) {
    $targetFull = [IO.Path]::GetFullPath($targetPath)

    foreach ($wb in $excel.Workbooks) {
        try {
            if ([string]::Equals(
                [IO.Path]::GetFullPath($wb.FullName),
                $targetFull,
                [StringComparison]::OrdinalIgnoreCase
            )) {
                return $wb
            }
        }
        catch {
            # Ignore workbooks that cannot expose FullName.
        }
    }

    return $null
}

function Get-ComponentByName($vbProject, $name) {
    foreach ($component in $vbProject.VBComponents) {
        if ([string]::Equals(
            $component.Name,
            $name,
            [StringComparison]::OrdinalIgnoreCase
        )) {
            return $component
        }
    }

    return $null
}

function Remove-ExistingComponent($vbProject, $name) {
    $existing = Get-ComponentByName $vbProject $name

    if ($null -eq $existing) {
        return
    }

    if ($existing.Type -eq 100) {
        throw "Cannot replace document module '$name'."
    }

    $vbProject.VBComponents.Remove($existing)
}

function Get-Utf8Text($filePath) {
    $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
    return [IO.File]::ReadAllText($filePath, $utf8)
}

function Convert-SourceToExcelAnsi($sourcePath, $tempPath) {
    # Excel/VBE component import is based on the legacy Windows text format.
    # The source files in the project are UTF-8.
    #
    # Use the current Windows ANSI code page rather than hard-coding 1251.
    $text = Get-Utf8Text $sourcePath
    $ansi = [Text.Encoding]::Default
    [IO.File]::WriteAllBytes($tempPath, $ansi.GetBytes($text))
}

function Strip-VbaExportHeader($text) {
    # Remove the CLASS export header.
    $text = [Regex]::Replace(
        $text,
        '(?ms)^\s*VERSION[^\r\n]*\r?\nBEGIN.*?^\s*END\s*\r?\n',
        ''
    )

    # Remove VBIDE Attribute lines. The actual VBA code remains intact.
    $text = [Regex]::Replace(
        $text,
        '(?m)^\s*Attribute\s+[^\r\n]*\r?\n?',
        ''
    )

    return $text.TrimStart()
}

function Create-ComponentFromUnicodeSource($vbProject, $sourcePath, $name, $type) {
    $component = $vbProject.VBComponents.Add($type)

    if ($null -eq $component) {
        throw "VBComponents.Add returned NULL for '$name'."
    }

    $component.Name = $name

    $code = Get-Utf8Text $sourcePath
    $code = Strip-VbaExportHeader $code

    if (-not [string]::IsNullOrWhiteSpace($code)) {
        $component.CodeModule.AddFromString($code)
    }

    return $component
}

try {
    if (-not (Test-Path -LiteralPath $WorkbookPath)) {
        throw "Workbook does not exist: $WorkbookPath"
    }

    if (-not (Test-Path -LiteralPath $SourcePath)) {
        throw "Source file does not exist: $SourcePath"
    }

    $excel = Get-ExcelApplication
    $excel.Visible = $true

    # Force VBE initialization before accessing VBProject.
    $null = $excel.VBE

    $workbook = Get-WorkbookByPath $excel $WorkbookPath

    if ($null -eq $workbook) {
        $workbook = $excel.Workbooks.Open($WorkbookPath)
    }

    if ($workbook.ReadOnly) {
        throw "Workbook is read-only: $WorkbookPath"
    }

    $vbProject = $workbook.VBProject

    $extension = [IO.Path]::GetExtension($SourcePath).ToLowerInvariant()
    $tempImport = $null
    $imported = $null

    try {
        if ($extension -eq ".frm") {
            $tempImport = Join-Path $env:TEMP ("vba-cli-" + [Guid]::NewGuid().ToString() + ".frm")
            Convert-SourceToExcelAnsi $SourcePath $tempImport

            Remove-ExistingComponent $vbProject $ComponentName

            $imported = $vbProject.VBComponents.Import($tempImport)

            if ($null -eq $imported) {
                throw "VBComponents.Import returned NULL for '$ComponentName'."
            }

            if ($imported.Type -ne $ExpectedType) {
                $actualType = $imported.Type
                $vbProject.VBComponents.Remove($imported)
                throw "Imported component '$ComponentName' has type $actualType, expected $ExpectedType."
            }
        }
        else {
            # For .bas/.cls we deliberately use VBComponents.Add().
            # This guarantees:
            #   .bas -> Standard Modules
            #   .cls -> Class Modules
            #
            # It also lets us pass Unicode text directly as a .NET string,
            # avoiding UTF-8 -> ANSI corruption.
            Remove-ExistingComponent $vbProject $ComponentName

            $imported = Create-ComponentFromUnicodeSource $vbProject $SourcePath $ComponentName $ExpectedType

            if ($null -eq $imported) {
                throw "Failed to create component '$ComponentName'."
            }

            if ($imported.Type -ne $ExpectedType) {
                $actualType = $imported.Type
                $vbProject.VBComponents.Remove($imported)
                throw "Created component '$ComponentName' has type $actualType, expected $ExpectedType."
            }
        }

        $workbook.Save()

        Write-Output ("OK|" + $ComponentName + "|" + $ExpectedType)
    }
    finally {
        if ($null -ne $tempImport) {
            Remove-Item -LiteralPath $tempImport -Force -ErrorAction SilentlyContinue
        }
    }
}
catch {
    [Console]::Error.WriteLine(("VBA_IMPORT_ERROR: " + $_.Exception.Message))
    exit 1
}
`;
}

function buildPowerShellBatchImportScript() {
    return String.raw`
param(
    [string]$WorkbookPath,
    [string]$ItemsFile
)

function Get-ExcelApplication {
    try { return [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") }
    catch {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        return $excel
    }
}

function Get-WorkbookByPath($excel, $targetPath) {
    $targetFull = [IO.Path]::GetFullPath($targetPath)
    foreach ($wb in $excel.Workbooks) {
        try {
            if ([string]::Equals([IO.Path]::GetFullPath($wb.FullName), $targetFull, [StringComparison]::OrdinalIgnoreCase)) { return $wb }
        } catch {}
    }
    return $null
}

function Get-ComponentByName($vbProject, $name) {
    foreach ($component in $vbProject.VBComponents) {
        if ([string]::Equals($component.Name, $name, [StringComparison]::OrdinalIgnoreCase)) { return $component }
    }
    return $null
}

function Remove-ExistingComponent($vbProject, $name) {
    $existing = Get-ComponentByName $vbProject $name
    if ($null -eq $existing) { return }
    if ($existing.Type -eq 100) { throw "Cannot replace document module '$name'." }
    $vbProject.VBComponents.Remove($existing)
}

function Get-Utf8Text($filePath) {
    $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
    return [IO.File]::ReadAllText($filePath, $utf8)
}

function Convert-SourceToExcelAnsi($sourcePath, $tempPath) {
    $text = Get-Utf8Text $sourcePath
    $ansi = [Text.Encoding]::Default
    [IO.File]::WriteAllBytes($tempPath, $ansi.GetBytes($text))
}

function Strip-VbaExportHeader($text) {
    $text = [Regex]::Replace($text, '(?ms)^\s*VERSION[^\r\n]*\r?\nBEGIN.*?^\s*END\s*\r?\n', '')
    $text = [Regex]::Replace($text, '(?m)^\s*Attribute\s+[^\r\n]*\r?\n?', '')
    return $text.TrimStart()
}

function Create-ComponentFromUnicodeSource($vbProject, $sourcePath, $name, $type) {
    $component = $vbProject.VBComponents.Add($type)
    if ($null -eq $component) { throw "VBComponents.Add returned NULL for '$name'." }
    $component.Name = $name
    $code = Strip-VbaExportHeader (Get-Utf8Text $sourcePath)
    if (-not [string]::IsNullOrWhiteSpace($code)) { $component.CodeModule.AddFromString($code) }
    return $component
}

try {
    if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook does not exist: $WorkbookPath" }
    if (-not (Test-Path -LiteralPath $ItemsFile)) { throw "Import manifest does not exist: $ItemsFile" }
    $itemsJson = [IO.File]::ReadAllText($ItemsFile, [Text.Encoding]::UTF8)
    $items = @(ConvertFrom-Json -InputObject $itemsJson)
    if ($items.Count -eq 0) { throw "No VBA files to import." }

    $excel = Get-ExcelApplication
    $excel.Visible = $true
    $null = $excel.VBE
    $workbook = Get-WorkbookByPath $excel $WorkbookPath
    if ($null -eq $workbook) { $workbook = $excel.Workbooks.Open($WorkbookPath) }
    if ($workbook.ReadOnly) { throw "Workbook is read-only: $WorkbookPath" }
    $vbProject = $workbook.VBProject

    foreach ($item in $items) {
        $sourcePath = [string]$item.SourcePath
        $componentName = [string]$item.ComponentName
        if (-not (Test-Path -LiteralPath $sourcePath)) { throw "Source file does not exist: $sourcePath" }

        $extension = [IO.Path]::GetExtension($sourcePath).ToLowerInvariant()
        # Do not trust the JSON numeric property here. Windows PowerShell 5.1
        # can materialize JSON values as System.Object[] in some argument/binding
        # situations. Derive the VBIDE component type from the file extension.
        switch ($extension) {
            ".bas" { $expectedType = 1 }
            ".cls" { $expectedType = 2 }
            ".frm" { $expectedType = 3 }
            default { throw "Unsupported VBA extension: $extension" }
        }
        $tempImport = $null
        try {
            if ($extension -eq ".frm") {
                $tempImport = Join-Path $env:TEMP ("vba-cli-" + [Guid]::NewGuid().ToString() + ".frm")
                Convert-SourceToExcelAnsi $sourcePath $tempImport
                Remove-ExistingComponent $vbProject $componentName
                $imported = $vbProject.VBComponents.Import($tempImport)
                if ($null -eq $imported) { throw "VBComponents.Import returned NULL for '$componentName'." }
            } else {
                Remove-ExistingComponent $vbProject $componentName
                $imported = Create-ComponentFromUnicodeSource -vbProject $vbProject -sourcePath $sourcePath -name $componentName -type $expectedType
            }
            if ($imported.Type -ne $expectedType) {
                $actualType = $imported.Type
                $vbProject.VBComponents.Remove($imported)
                throw "Component '$componentName' has type $actualType, expected $expectedType."
            }
            Write-Output ("OK|" + $componentName + "|" + $expectedType)
        } finally {
            if ($null -ne $tempImport) { Remove-Item -LiteralPath $tempImport -Force -ErrorAction SilentlyContinue }
        }
    }
    $workbook.Save()
} catch {
    [Console]::Error.WriteLine(("VBA_IMPORT_ERROR: " + $_.Exception.Message))
    exit 1
}
`;
}

async function importFile(filePath, options = {}) {
    const paths = getProjectPaths(options.cwd || process.cwd());

    if (!fs.existsSync(paths.workbookPath)) {
        throw new Error(`Excel-книга не найдена: ${paths.workbookPath}`);
    }

    const absoluteFile = path.resolve(filePath);
    const extension = path.extname(absoluteFile).toLowerCase();

    if (!VBA_EXTENSIONS.includes(extension)) {
        throw new Error(`Неподдерживаемое VBA-расширение: ${extension}`);
    }

    const componentName = componentNameFromFile(absoluteFile);
    const expectedType = COMPONENT_TYPES[extension];

    info(`Импорт: ${path.relative(paths.root, absoluteFile)}`);

    try {
        await runPowerShell(
            buildPowerShellImportScript(),
            [
                paths.workbookPath,
                absoluteFile,
                componentName,
                expectedType
            ]
        );

        success(`Импортирован: ${componentName}`);
        return true;
    } catch (err) {
        throw new Error(cleanPowerShellError(err.message));
    }
}

async function importVBA(options = {}) {
    console.log("\nVBA IMPORT\n");

    const paths = getProjectPaths(options.cwd || process.cwd());
    const files = getVBAFiles(paths.sourceDir);

    if (files.length === 0) {
        warning("VBA-файлы в src не найдены.");
        return 0;
    }

    let importedCount = 0;
    let failedCount = 0;

    // Import every file through the single-file PowerShell script.
    // This is deliberately sequential: Windows PowerShell 5.1 has several
    // ConvertFrom-Json / COM binding edge cases with arrays and manifests,
    // which previously caused all SourcePath values to be concatenated into
    // one string such as "file1 file2 file3".
    for (const file of files) {
        try {
            await importFile(file, { cwd: paths.root });
            importedCount++;
        } catch (err) {
            failedCount++;
            error(cleanPowerShellError(err.message));
        }
    }

    if (failedCount > 0) {
        throw new Error(
            `Импорт завершён с ошибками. Успешно: ${importedCount}, ошибок: ${failedCount}.`
        );
    }

    success(`Импорт завершён. Успешно: ${importedCount}, ошибок: 0.`);
    return importedCount;
}

async function build() {
    console.log("\nVBA BUILD\n");

    const paths = getProjectPaths();

    createDir(paths.buildDir);

    info("Импорт исходников в Excel...");
    await importVBA();

    success("Build завершён.");
}

function clean() {
    console.log("\nVBA CLEAN\n");

    const paths = getProjectPaths();

    if (!fs.existsSync(paths.buildDir)) {
        info("build/ уже отсутствует.");
        return;
    }

    fs.rmSync(paths.buildDir, {
        recursive: true,
        force: true
    });

    success("build/ удалён.");
}

function debounce(fn, delay = 250) {
    let timer = null;

    return (...args) => {
        if (timer) {
            clearTimeout(timer);
        }

        timer = setTimeout(() => {
            timer = null;
            fn(...args);
        }, delay);
    };
}

async function dev() {
    console.log("\nVBA DEV MODE\n");

    if (!chokidar) {
        throw new Error(
            "Не установлен chokidar. Выполните: npm install"
        );
    }

    const paths = getProjectPaths();

    info(`Project: ${paths.root}`);
    info(`Workbook: ${paths.workbookPath}`);
    info(`Source: ${paths.sourceDir}`);

    if (!fs.existsSync(paths.workbookPath)) {
        throw new Error(`Excel-книга не найдена: ${paths.workbookPath}`);
    }

    if (!fs.existsSync(paths.sourceDir)) {
        throw new Error(`Каталог src не найден: ${paths.sourceDir}`);
    }

    separator();
    info("Ожидание изменений...");
    separator();

    let queue = Promise.resolve();

    function enqueue(task) {
        queue = queue
            .then(task)
            .catch(err => {
                error(err.message);
            });

        return queue;
    }

    const importChanged = debounce(filePath => {
        if (!fs.existsSync(filePath) || !isVBAFile(filePath)) {
            return;
        }

        info(`Изменён: ${path.relative(paths.root, filePath)}`);

        enqueue(async () => {
            try {
                await importFile(filePath, { cwd: paths.root });
            } catch (err) {
                error(err.message);
            }
        });
    }, 300);

    const watcher = chokidar.watch(
        paths.sourceDir,
        {
            persistent: true,
            ignoreInitial: true,
            awaitWriteFinish: {
                stabilityThreshold: 350,
                pollInterval: 100
            },
            ignored: [
                /(^|[\/\\])\../,
                "**/build/**"
            ]
        }
    );

    watcher.on("add", filePath => {
        if (isVBAFile(filePath)) {
            importChanged(filePath);
        }
    });

    watcher.on("change", filePath => {
        if (isVBAFile(filePath)) {
            importChanged(filePath);
        }
    });

    watcher.on("unlink", filePath => {
        if (!isVBAFile(filePath)) {
            return;
        }

        info(`Удалён файл: ${path.relative(paths.root, filePath)}`);

        enqueue(async () => {
            try {
                await removeComponent(filePath, { cwd: paths.root });
            } catch (err) {
                error(err.message);
            }
        });
    });

    watcher.on("error", err => {
        error(`Watcher: ${err.message}`);
    });

    success("Watcher запущен.");
    info("Изменения .bas / .cls / .frm будут автоматически импортироваться в Excel.");
    separator();
    info("Для остановки: Ctrl + C");

    const shutdown = async () => {
        try {
            await watcher.close();
        } finally {
            process.exit(0);
        }
    };

    process.once("SIGINT", shutdown);
    process.once("SIGTERM", shutdown);

    await new Promise(() => { });
}

async function init(projectName) {
    console.log("\nVBA INIT\n");

    if (!projectName) {
        projectName = "vba-project";
    }

    if (!/^[a-zA-Z0-9._-]+$/.test(projectName)) {
        throw new Error(
            "Имя проекта должно содержать только латинские буквы, цифры, '.', '_' или '-'."
        );
    }

    const projectPath = path.resolve(process.cwd(), projectName);

    if (fs.existsSync(projectPath)) {
        throw new Error(`Каталог уже существует: ${projectPath}`);
    }

    info(`Создание проекта "${projectName}"...`);

    createDir(projectPath);

    const answer = await ask("Создать стартовые VBA-модули? [Y/n]: ");
    const createStarter = answer === "" || /^y(es)?$/i.test(answer);

    if (createStarter) {
        buildStarterFiles(projectPath);
        success("Стартовые VBA-файлы созданы.");
    }

    createProjectFiles(projectPath, projectName);

    const workbookPath = path.join(projectPath, `${projectName}.xlsm`);

    info("Создание Excel-книги...");
    createExcelFile(workbookPath);
    success(`Создано: ${projectName}.xlsm`);

    success(`Проект "${projectName}" создан.`);
    info(`Рабочая директория: ${projectPath}`);

    // All following project operations are executed from the new project root.
    process.chdir(projectPath);

    if (createStarter) {
        separator();
        info("Первичный импорт VBA...");

        try {
            await importVBA({ cwd: projectPath });
        } catch (err) {
            error(err.message);
            warning("Проект создан, но первичный импорт завершился с ошибкой.");
        }
    }

    separator();
    openInVSCode(projectPath);

    separator();
    info("Запуск режима разработки...");

    await dev();
}

function printHelp() {
    console.log(`
VBA CLI v${CLI_VERSION}

Использование:

  vba init [name]       Создать новый VBA-проект
  vba dev               Запустить watcher
  vba import            Импортировать src -> Excel
  vba export            Экспортировать Excel -> src
  vba build             Выполнить сборку
  vba clean             Удалить build/
  vba help              Показать справку

Примеры:

  vba init test
  vba dev
  vba import
  vba export
  vba build
  vba clean

Workflow:

  VS Code
      |
      v
  src/*.bas, *.cls, *.frm
      |
      | Ctrl + S
      v
  VBA CLI watcher
      |
      v
  PowerShell
      |
      v
  Excel COM
      |
      v
  VBA Project
`);
}

async function main() {
    const command = process.argv[2];
    const args = process.argv.slice(3);

    try {
        switch (command) {
            case "init":
                await init(args[0]);
                break;

            case "dev":
                await dev();
                break;

            case "import":
                await importVBA();
                break;

            case "export":
                await exportVBA();
                break;

            case "build":
                await build();
                break;

            case "clean":
                clean();
                break;

            case "help":
            case "--help":
            case "-h":
            case undefined:
                printHelp();
                break;

            case "version":
            case "--version":
            case "-v":
                console.log(`v${CLI_VERSION}`);
                break;

            default:
                error(`Неизвестная команда: ${command}`);
                printHelp();
                process.exitCode = 1;
        }
    } catch (err) {
        error(err.message);
        process.exitCode = 1;
    }
}

main();