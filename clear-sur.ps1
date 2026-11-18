<#
.SYNOPSIS
    Стирает фамилию в строке «Исследуемое лицо» во всех документах Word в папке (включая подпапки).

.DESCRIPTION
    В каждом документе (.doc, .docx, .docm) ищется подпись «Исследуемое лицо»
    (подходит и написание «Иследуемое лицо») и удаляется первое слово стоящего напротив ФИО,
    то есть фамилия. Имя (и отчество, если есть) остаётся.

    Где ищется ФИО:
      1) в той же строке после подписи:  «Исследуемое лицо: Иванов Иван»  ->  «Исследуемое лицо: Иван»
      2) если там пусто, то в соседней справа ячейке таблицы.
    Фраза вроде «Исследуемое лицо ознакомлено…» не трогается: после подписи должно стоять слово с заглавной буквы.

    Оригиналы НЕ меняются: обработанные копии кладутся в папку «<папка> (без фамилий)»
    с той же структурой подпапок. Там же появляется отчёт «отчёт.csv» (открывается в Excel).

.EXAMPLE
    .\clear-surnames.ps1 "D:\Протоколы"
.EXAMPLE
    .\clear-surnames.ps1 "D:\Протоколы" -Replacement "***"     # заменить фамилию на *** вместо удаления
.EXAMPLE
    .\clear-surnames.ps1 "D:\Протоколы" -InPlace               # менять сами файлы, без копий
#>
param(
    [Parameter(Position = 0)][string]$Path,
    [string]$OutPath,
    [string]$Replacement = '',
    [switch]$InPlace
)

$ErrorActionPreference = 'Stop'

# Константы Word
$wdCollapseEnd  = 0
$wdFindStop     = 0
$wdCharacter    = 1
$wdWithInTable  = 12
$wdNoProtection = -1
$wdDoNotSave    = 0

# Ищем «ледуемое», чтобы найти и «Исследуемое», и «Иследуемое», затем проверяем, что дальше «лицо».
$labelSearch = 'ледуемое'
$labelRx = [regex]'(?i)^ледуемое[\s\u00A0]+лицо'
# Первое слово с заглавной буквы после подписи. Пропускаются двоеточия, тире, подчёркивания,
# номера и пояснения в скобках, например «(Ф.И.О.):».
$nameRx  = [regex]'^(?>(?:\([^)]*\)|[\s\u00A0\p{P}\p{S}\p{N}])*)(?<w>\p{Lu}[\p{L}\-''\u2019]*)'

function Get-CleanText([string]$s) {
    $t = ($s -replace '[\r\a\t]', ' ' -replace '\s+', ' ').Trim()
    if ($t.Length -gt 80) { $t = $t.Substring(0, 80) + '…' }
    return $t
}

# Возвращает диапазон, где стоит ФИО (остаток строки или соседняя ячейка), или $null.
function Find-NameRange($label) {
    $para = $label.Paragraphs.Item(1).Range
    $rest = $label.Duplicate
    $rest.SetRange($label.End, $para.End)
    $core = ($rest.Text -replace '[\r\a]', '') -replace '\([^)]*\)', ''
    if ($core -match '\p{L}') { return ,$rest }

    if ($label.Information($wdWithInTable)) {
        $cell = $label.Cells.Item(1)
        $next = $cell.Next
        if ($next -and $next.RowIndex -eq $cell.RowIndex) {
            $r = $next.Range
            [void]$r.MoveEnd($wdCharacter, -1)   # без маркера конца ячейки
            return ,$r
        }
    }
    return $null
}

# Стирает фамилию (первое слово) в диапазоне. Возвращает фамилию или $null.
function Remove-Surname($nameRange) {
    $m = $nameRx.Match($nameRange.Text)
    if (-not $m.Success) { return $null }
    $surname = $m.Groups['w'].Value

    $r = $nameRange.Duplicate
    $f = $r.Find
    $f.ClearFormatting()
    $f.Text = $surname
    $f.Forward = $true
    $f.Wrap = $wdFindStop
    $f.Format = $false
    $f.MatchCase = $true
    $f.MatchWholeWord = $false
    $f.MatchWildcards = $false
    if (-not $f.Execute() -or $r.Text -cne $surname -or $r.End -gt $nameRange.End) { return $null }

    if ($Replacement) {
        $r.Text = $Replacement
    } else {
        [void]$r.MoveEndWhile(' ' + [char]0xA0)   # вместе с пробелом после фамилии
        $r.Text = ''
    }
    return $surname
}

function Invoke-Story($story) {
    $search = $story.Duplicate
    $f = $search.Find
    $f.ClearFormatting()
    $f.Text = $labelSearch
    $f.Forward = $true
    $f.Wrap = $wdFindStop
    $f.Format = $false
    $f.MatchCase = $false
    $f.MatchWholeWord = $false
    $f.MatchWildcards = $false

    $guard = 0
    while ($guard++ -lt 200 -and $f.Execute()) {
        $para  = $search.Paragraphs.Item(1).Range
        $probe = $search.Duplicate
        $probe.SetRange($search.Start, $para.End)
        $m = $labelRx.Match($probe.Text)
        if ($m.Success) {
            $label = $search.Duplicate
            $label.SetRange($search.Start, $search.Start + $m.Length)
            $nameRange = Find-NameRange $label
            if (-not $nameRange) {
                [pscustomobject]@{ Status = 'ФИО напротив не найдено'; Before = ''; After = ''; Surname = '' }
            } else {
                $before = Get-CleanText $nameRange.Text
                $surname = Remove-Surname $nameRange
                if ($surname) {
                    [pscustomobject]@{ Status = 'стёрто'; Before = $before; After = (Get-CleanText $nameRange.Text); Surname = $surname }
                } else {
                    [pscustomobject]@{ Status = 'пропущено: после подписи не ФИО'; Before = $before; After = $before; Surname = '' }
                }
            }
        }
        $search.Collapse($wdCollapseEnd)
    }
}

function Invoke-Document($doc) {
    $protection = $doc.ProtectionType
    if ($protection -ne $wdNoProtection) { $doc.Unprotect() }   # защищённую паролем форму открыть не получится -> ошибка в отчёте
    $doc.TrackRevisions = $false                                 # иначе фамилия останется в исправлениях

    $results = @()
    foreach ($story in $doc.StoryRanges) {                       # основной текст, колонтитулы, надписи
        $s = $story
        while ($s) {
            $results += @(Invoke-Story $s)
            $s = $s.NextStoryRange
        }
    }

    if ($protection -ne $wdNoProtection) { $doc.Protect($protection, $true) }   # вернуть защиту, не сбрасывая поля формы
    return $results
}

# ---------- Основная часть ----------

if (-not $Path) {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Выберите папку с документами Word'
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $Path = $dlg.SelectedPath
}
$Path = (Resolve-Path -LiteralPath $Path).ProviderPath.TrimEnd('\')
if ($InPlace) {
    $OutPath = $Path
} elseif (-not $OutPath) {
    $OutPath = "$Path (без фамилий)"
}
$OutPath = $OutPath.TrimEnd('\')

$files = @(Get-ChildItem -LiteralPath $Path -Recurse -File |
    Where-Object { $_.Extension -match '^\.(docx?|docm)$' -and $_.Name -notlike '~$*' } |
    Where-Object { $InPlace -or -not $_.FullName.StartsWith($OutPath + '\', [StringComparison]::OrdinalIgnoreCase) })

if ($files.Count -eq 0) { Write-Host "В папке $Path нет документов Word." -ForegroundColor Yellow; return }

Write-Host "Документов: $($files.Count)"
if ($InPlace) { Write-Host 'Режим: изменяются сами файлы.' -ForegroundColor Yellow }
else          { Write-Host "Результат будет в папке: $OutPath" }
Write-Host ''

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$word.AutomationSecurity = 3   # макросы в документах не запускать

$report = New-Object System.Collections.Generic.List[object]
$i = 0
try {
    foreach ($file in $files) {
        $i++
        $rel = $file.FullName.Substring($Path.Length).TrimStart('\')
        $target = Join-Path $OutPath $rel
        Write-Host ("[{0}/{1}] {2} ... " -f $i, $files.Count, $rel) -NoNewline

        $doc = $null
        try {
            if (-not $InPlace) {
                [void](New-Item -ItemType Directory -Force -Path (Split-Path $target -Parent))
                Copy-Item -LiteralPath $file.FullName -Destination $target -Force
            }
            # Пароль-заглушка: незащищённые файлы откроются как обычно, а на запароленных будет ошибка вместо окна с вопросом.
            $doc = $word.Documents.Open($target, $false, $false, $false, ' ', ' ', $false, ' ', ' ')
            $results = @(Invoke-Document $doc)
            $cleared = @($results | Where-Object { $_.Status -eq 'стёрто' })
            if ($cleared.Count -gt 0) { $doc.Save() }

            if ($results.Count -eq 0) {
                $report.Add([pscustomobject]@{ 'Файл' = $rel; 'Статус' = 'строки «Исследуемое лицо» нет'; 'Было' = ''; 'Стало' = ''; 'Фамилия ещё встречается в тексте' = '' })
            } else {
                $allText = $doc.Content.Text
                foreach ($r in $results) {
                    $left = ''
                    if ($r.Surname) {
                        # считаем и падежные формы: Иванов -> Иванова, Иванову, Ивановым…
                        $left = [regex]::Matches($allText, '(?<!\p{L})' + [regex]::Escape($r.Surname) + '\p{L}{0,3}(?!\p{L})').Count
                    }
                    $report.Add([pscustomobject]@{ 'Файл' = $rel; 'Статус' = $r.Status; 'Было' = $r.Before; 'Стало' = $r.After; 'Фамилия ещё встречается в тексте' = $left })
                }
            }

            if ($cleared.Count -gt 0) { Write-Host ('стёрто: ' + (($cleared | ForEach-Object { $_.Before }) -join '; ')) -ForegroundColor Green }
            elseif ($results.Count -eq 0) { Write-Host 'строки нет' -ForegroundColor DarkYellow }
            else { Write-Host $results[0].Status -ForegroundColor DarkYellow }
        } catch {
            Write-Host ('ОШИБКА: ' + $_.Exception.Message) -ForegroundColor Red
            $report.Add([pscustomobject]@{ 'Файл' = $rel; 'Статус' = 'ошибка: ' + $_.Exception.Message; 'Было' = ''; 'Стало' = ''; 'Фамилия ещё встречается в тексте' = '' })
        } finally {
            if ($doc) { $doc.Close($wdDoNotSave) }
        }
    }
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}

[void](New-Item -ItemType Directory -Force -Path $OutPath)
$reportPath = Join-Path $OutPath 'отчёт.csv'
$report | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'

$ok = @($report | Where-Object { $_.'Статус' -eq 'стёрто' } | Select-Object -ExpandProperty 'Файл' -Unique).Count
Write-Host ''
Write-Host "Готово. Фамилии стёрты в $ok из $($files.Count) документов." -ForegroundColor Cyan
Write-Host "Отчёт: $reportPath"
