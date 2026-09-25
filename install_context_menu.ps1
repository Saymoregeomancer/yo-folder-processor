# Додає / прибирає пункт "Yo Folder Processor" у контекстному меню Провідника.
# Пишеться тільки в HKCU — права адміністратора не потрібні.
param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$key = 'YoFolderProcessor'

# клік по папці -> %1, клік по порожньому місцю всередині папки -> %V
$targets = @{
    'HKCU:\Software\Classes\Directory\shell'            = '%1'
    'HKCU:\Software\Classes\Directory\Background\shell' = '%V'
}

$items = @(
    @{ Id = '1_process'; Title = 'Обробити папки';          Bat = 'process_folders.bat' },
    @{ Id = '2_dryrun';  Title = 'Перевірка (dry-run)';     Bat = 'process_folders_dry_run.bat' }
)

foreach ($base in $targets.Keys) {
    $root = Join-Path $base $key
    if (Test-Path -LiteralPath $root) {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}

if ($Uninstall) {
    Write-Host 'Пункт меню видалено.'
    exit 0
}

foreach ($base in $targets.Keys) {
    $arg = $targets[$base]
    $root = Join-Path $base $key
    New-Item -Path $root -Force | Out-Null
    New-ItemProperty -LiteralPath $root -Name 'MUIVerb' -Value 'Yo Folder Processor' -Force | Out-Null
    New-ItemProperty -LiteralPath $root -Name 'SubCommands' -Value '' -Force | Out-Null
    New-ItemProperty -LiteralPath $root -Name 'Icon' -Value "$env:SystemRoot\System32\imageres.dll,-5302" -Force | Out-Null

    foreach ($item in $items) {
        $bat = Join-Path $here $item.Bat
        $sub = Join-Path $root ('shell\' + $item.Id)
        New-Item -Path "$sub\command" -Force | Out-Null
        New-ItemProperty -LiteralPath $sub -Name 'MUIVerb' -Value $item.Title -Force | Out-Null
        Set-Item -LiteralPath "$sub\command" -Value ('"{0}" "{1}"' -f $bat, $arg)
    }
}

Write-Host 'Готово. Клацніть правою кнопкою по папці -> "Yo Folder Processor".'
Write-Host '(Windows 11: спершу "Показати додаткові параметри" / Shift+F10.)'
Write-Host "Скрипти запускаються з: $here"
