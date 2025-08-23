<#
.SYNOPSIS
    Скрипт для отслеживания изменений всех файлов в указанной директории
.DESCRIPTION
    Мониторит изменения файлов (содержимое, размер, атрибуты) в указанной папке
    и выводит уведомления при обнаружении изменений.
.PARAMETER Path
    Путь к директории для мониторинга
.PARAMETER Interval
    Интервал проверки в секундах (по умолчанию 5 секунд)
.PARAMETER IncludeSubdirectories
    Включить мониторинг подкаталогов (по умолчанию false)
.PARAMETER Filter
    Фильтр файлов (например, "*.log" или "*.txt")
.EXAMPLE
    .\DirectoryChangeMonitor.ps1 -Path "C:\logs" -Interval 10
.EXAMPLE
    .\DirectoryChangeMonitor.ps1 -Path "D:\data" -IncludeSubdirectories -Filter "*.csv"
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$Path,
    [int]$Interval = 5,
    [switch]$IncludeSubdirectories,
    [string]$Filter = "*",
    [Parameter(Mandatory=$true)]
    [string]$Server,
    [Parameter(Mandatory=$true)]
    [string]$Infobase
)

# Проверяем существование директории
if (-not (Test-Path -Path $Path -PathType Container)) {
    Write-Error "Директория '$Path' не найдена!"
    exit 1
}

# Инициализация COM-подключения
$ComConnector = New-Object -ComObject V83.ComConnector
$ComConnection = $ComConnector.Connect("srvr='$Server'; ref='$Infobase';")
$Module1C = [System.__ComObject].InvokeMember("МойОбщийМодуль",[System.Reflection.BindingFlags]::GetProperty,$null,$ComConnection,$null)

# Хэш-таблица для хранения состояний файлов
$fileStates = @{}

# Функция для получения информации о файле
function Get-FileInfo {
    param (
        [string]$FilePath
    )
    
    $file = Get-Item -Path $FilePath
    $hash = Get-FileHash -Path $FilePath -Algorithm MD5 -ErrorAction SilentlyContinue
    
    return @{
        FullName = $file.FullName
        LastWriteTime = $file.LastWriteTime
        Size = $file.Length
        Attributes = $file.Attributes
        Hash = if ($hash) { $hash.Hash } else { $null }
    }
}

# Инициализация - получаем информацию о всех файлах
function Initialize-FileStates {
    $params = @{
        Path = $Path
        Filter = $Filter
        File = $true
    }
    
    if ($IncludeSubdirectories) {
        $params.Recurse = $true
    }
    
    Get-ChildItem @params | ForEach-Object {
        $fileStates[$_.FullName] = Get-FileInfo -FilePath $_.FullName
    }
}

#
function Call-1c-method {
    param(
        $filePath,
        $Content
    )

    $Args = @([String]$filePath, [String]$Content)
    return [System.__ComObject].InvokeMember("МойМетод1",[System.Reflection.BindingFlags]::InvokeMethod,$null,$Module1C, $Args)

}

# Функция для проверки изменений
function Check-ForChanges {
    $params = @{
        Path = $Path
        Filter = $Filter
        File = $true
    }
    
    if ($IncludeSubdirectories) {
        $params.Recurse = $true
    }
    
    $currentFiles = Get-ChildItem @params
    
    # 1. Проверяем измененные и существующие файлы
    foreach ($file in $currentFiles) {
        $filePath = $file.FullName
        $currentState = Get-FileInfo -FilePath $filePath
        
        if ($fileStates.ContainsKey($filePath)) {
            $lastState = $fileStates[$filePath]
            $changes = @()
            
            if ($currentState.LastWriteTime -ne $lastState.LastWriteTime) {
                $changes += "время изменения"
            }
            
            if ($currentState.Size -ne $lastState.Size) {
                $changes += "размер ($($lastState.Size) → $($currentState.Size))"
            }
            
            if ($currentState.Attributes -ne $lastState.Attributes) {
                $changes += "атрибуты"
            }
            
            if ($currentState.Hash -ne $lastState.Hash -and $currentState.Hash -ne $null -and $lastState.Hash -ne $null) {
                $changes += "содержимое"
            }
            
            if ($changes.Count -gt 0) {
                $changeMessage = "[$(Get-Date -Format 'HH:mm:ss')] Файл изменен: $filePath Изменения: " + ($changes -join ", ")
                $content = Get-Content -Path $filePath -Raw -Encoding UTF8
                $result = Call-1c-method $filePath $content
                Write-Host $result -ForegroundColor Gray
                Write-Host $changeMessage -ForegroundColor Green
            }
        } else {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Новый файл: $filePath" -ForegroundColor Blue
        }
        
        # Обновляем состояние файла
        $fileStates[$filePath] = $currentState
    }
    
    # 2. Проверяем удаленные файлы
    $removedFiles = @()
    foreach ($filePath in $fileStates.Keys) {
        if (-not (Test-Path -Path $filePath -PathType Leaf)) {
            $removedFiles += $filePath
        }
    }
    
    foreach ($filePath in $removedFiles) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Файл удален: $filePath" -ForegroundColor Red
        $fileStates.Remove($filePath)
    }
}

# Инициализация
Initialize-FileStates
Write-Host "Начало мониторинга директории: $Path"
Write-Host "Всего файлов: $($fileStates.Count)"
Write-Host "Интервал проверки: $Interval секунд"
Write-Host "Включены подкаталоги: $IncludeSubdirectories"
Write-Host "Фильтр файлов: '$Filter'"
Write-Host "Нажмите Ctrl+C для остановки мониторинга"

# Основной цикл мониторинга
try {
     while ($true) {
        Start-Sleep -Seconds $Interval
        Check-ForChanges
    }
}
catch [System.Management.Automation.CmdletInvocationException] {
    if ($_.FullyQualifiedErrorId -eq "CtrlC") {
        Write-Host "Мониторинг остановлен пользователем" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Ошибка мониторинга: $_" -ForegroundColor Red
}
finally {
    Write-Host "Итоговое количество отслеживаемых файлов: $($fileStates.Count)"
    $content = $null
    $result = $null
    $Module1C = $null
    $ComConnection = $null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ComConnector) | Out-Null

    # Принудительная сборка мусора
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
}