<#
.SYNOPSIS
Simple GUI to merge multiple Excel/CSV files into a single Excel workbook.

.DESCRIPTION
The Invoke-ExcelMergeGui.ps1 script opens a small WinForms window where you can drag and drop
Excel (.xlsx, .xlsm) or CSV files. Two merge modes are available:
- One worksheet per source file (default). The worksheet is named after the source file.
- All data in a single worksheet. Only allowed if all files share the exact same headers.

An optional checkbox adds a 'SourceFile' column containing the source file name.
A 'Header row' selector indicates on which row the header is located (default: 1),
applied to all files (rows above the header are ignored).

An 'Exclusions' box accepts one rule per line, in the form Column=Value. Any row matching
at least one rule is excluded from the export. Matching is case-insensitive and supports
the * and ? wildcards (for example: computer=PC01, computername=SRV*). Rules referencing
a column absent from a file are ignored for that file.

Requires the ImportExcel module (the script offers to install it if missing).

.NOTES
Author: Bastien Perez
Date: 2026/08/11
Version: 1.0

Drag and drop from Windows Explorer does not work if PowerShell runs elevated
(Windows blocks drag and drop between processes with different integrity levels).
Use the 'Add...' button in that case.

Only the first worksheet of each source Excel file is read.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$supportedExtensions = @('.xlsx', '.xlsm', '.csv')

function Import-DataFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [int]$HeaderRow = 1
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLower()

    if ($extension -eq '.csv') {
        # Skip lines before the header, then detect delimiter (';' or ',') on the header line
        $lines = @(Get-Content -Path $Path | Select-Object -Skip ($HeaderRow - 1))
        $data = @()
        if ($lines.Count -gt 0) {
            $headerLine = $lines[0]
            $delimiter = ','
            if (($headerLine -split ';').Count -gt ($headerLine -split ',').Count) {
                $delimiter = ';'
            }
            $data = @($lines | ConvertFrom-Csv -Delimiter $delimiter)
        }
    }
    else {
        # First worksheet only, header at the given row
        $data = @(Import-Excel -Path $Path -StartRow $HeaderRow)
    }

    return $data
}

function Test-RowExcluded {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Row,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[PSCustomObject]]$ExclusionRules
    )

    foreach ($rule in $ExclusionRules) {
        # Property lookup by name is case-insensitive
        $property = $Row.PSObject.Properties[$rule.Column]
        if ($property) {
            $propertyValue = [string]$property.Value
            if ($propertyValue -like $rule.Value) {
                return $true
            }
        }
    }

    return $false
}

function Get-UniqueSheetName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseName,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]]$UsedNames
    )

    # Excel forbids these characters in worksheet names and limits length to 31
    $cleanName = $BaseName -replace '[\[\]\*\?/\\:]', '_'
    if ($cleanName.Length -gt 31) {
        $cleanName = $cleanName.Substring(0, 31)
    }

    $candidate = $cleanName
    $index = 1
    while ($UsedNames -contains $candidate) {
        $suffix = "_$index"
        $maxLength = 31 - $suffix.Length
        $candidate = $cleanName.Substring(0, [Math]::Min($cleanName.Length, $maxLength)) + $suffix
        $index++
    }

    $UsedNames.Add($candidate)
    return $candidate
}

function Add-FileToList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.ListBox]$ListBox
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLower()
    if (($supportedExtensions -contains $extension) -and (-not $ListBox.Items.Contains($Path))) {
        $null = $ListBox.Items.Add($Path)
    }
}

# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

$form = [System.Windows.Forms.Form]::new()
$form.Text = 'Excel / CSV Merger'
$form.Size = [System.Drawing.Size]::new(680, 700)
$form.MinimumSize = [System.Drawing.Size]::new(560, 600)
$form.StartPosition = 'CenterScreen'

$labelDrop = [System.Windows.Forms.Label]::new()
$labelDrop.Text = 'Drag and drop your Excel/CSV files below (or use the Add... button):'
$labelDrop.Location = [System.Drawing.Point]::new(12, 10)
$labelDrop.AutoSize = $true

$listBox = [System.Windows.Forms.ListBox]::new()
$listBox.Location = [System.Drawing.Point]::new(12, 32)
$listBox.Size = [System.Drawing.Size]::new(640, 300)
$listBox.Anchor = 'Top, Bottom, Left, Right'
$listBox.SelectionMode = 'MultiExtended'
$listBox.HorizontalScrollbar = $true
$listBox.AllowDrop = $true

$listBox.Add_DragEnter({
        param($sender, $e)
        if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        }
    })

$listBox.Add_DragDrop({
        param($sender, $e)
        foreach ($droppedPath in $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)) {
            if (Test-Path -Path $droppedPath -PathType Container) {
                # Folder dropped: add all supported files it contains
                foreach ($childFile in (Get-ChildItem -Path $droppedPath -File)) {
                    Add-FileToList -Path $childFile.FullName -ListBox $listBox
                }
            }
            else {
                Add-FileToList -Path $droppedPath -ListBox $listBox
            }
        }
    })

$buttonAdd = [System.Windows.Forms.Button]::new()
$buttonAdd.Text = 'Add...'
$buttonAdd.Location = [System.Drawing.Point]::new(12, 340)
$buttonAdd.Size = [System.Drawing.Size]::new(90, 28)
$buttonAdd.Anchor = 'Bottom, Left'
$buttonAdd.Add_Click({
        $openDialog = [System.Windows.Forms.OpenFileDialog]::new()
        $openDialog.Filter = 'Excel/CSV files (*.xlsx;*.xlsm;*.csv)|*.xlsx;*.xlsm;*.csv'
        $openDialog.Multiselect = $true
        if ($openDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            foreach ($selectedFile in $openDialog.FileNames) {
                Add-FileToList -Path $selectedFile -ListBox $listBox
            }
        }
    })

$buttonRemove = [System.Windows.Forms.Button]::new()
$buttonRemove.Text = 'Remove'
$buttonRemove.Location = [System.Drawing.Point]::new(108, 340)
$buttonRemove.Size = [System.Drawing.Size]::new(90, 28)
$buttonRemove.Anchor = 'Bottom, Left'
$buttonRemove.Add_Click({
        foreach ($selectedItem in @($listBox.SelectedItems)) {
            $listBox.Items.Remove($selectedItem)
        }
    })

$buttonClear = [System.Windows.Forms.Button]::new()
$buttonClear.Text = 'Clear list'
$buttonClear.Location = [System.Drawing.Point]::new(204, 340)
$buttonClear.Size = [System.Drawing.Size]::new(90, 28)
$buttonClear.Anchor = 'Bottom, Left'
$buttonClear.Add_Click({ $listBox.Items.Clear() })

$groupMode = [System.Windows.Forms.GroupBox]::new()
$groupMode.Text = 'Merge mode'
$groupMode.Location = [System.Drawing.Point]::new(12, 378)
$groupMode.Size = [System.Drawing.Size]::new(640, 78)
$groupMode.Anchor = 'Bottom, Left, Right'

$radioSheetPerFile = [System.Windows.Forms.RadioButton]::new()
$radioSheetPerFile.Text = 'One worksheet per file'
$radioSheetPerFile.Location = [System.Drawing.Point]::new(12, 22)
$radioSheetPerFile.AutoSize = $true
$radioSheetPerFile.Checked = $true

$radioSingleSheet = [System.Windows.Forms.RadioButton]::new()
$radioSingleSheet.Text = 'Everything in a single worksheet (identical headers required)'
$radioSingleSheet.Location = [System.Drawing.Point]::new(12, 46)
$radioSingleSheet.AutoSize = $true

$groupMode.Controls.Add($radioSheetPerFile)
$groupMode.Controls.Add($radioSingleSheet)

$groupExclusions = [System.Windows.Forms.GroupBox]::new()
$groupExclusions.Text = 'Exclusions (one rule per line, Column=Value, * and ? wildcards allowed)'
$groupExclusions.Location = [System.Drawing.Point]::new(12, 462)
$groupExclusions.Size = [System.Drawing.Size]::new(640, 90)
$groupExclusions.Anchor = 'Bottom, Left, Right'

$textBoxExclusions = [System.Windows.Forms.TextBox]::new()
$textBoxExclusions.Multiline = $true
$textBoxExclusions.ScrollBars = 'Vertical'
$textBoxExclusions.Location = [System.Drawing.Point]::new(12, 20)
$textBoxExclusions.Size = [System.Drawing.Size]::new(616, 58)
$textBoxExclusions.Anchor = 'Top, Bottom, Left, Right'

$groupExclusions.Controls.Add($textBoxExclusions)

$checkBoxSourceFile = [System.Windows.Forms.CheckBox]::new()
$checkBoxSourceFile.Text = 'Add a SourceFile column (source file name)'
$checkBoxSourceFile.Location = [System.Drawing.Point]::new(12, 560)
$checkBoxSourceFile.AutoSize = $true
$checkBoxSourceFile.Anchor = 'Bottom, Left'

$labelHeaderRow = [System.Windows.Forms.Label]::new()
$labelHeaderRow.Text = 'Header row:'
$labelHeaderRow.Location = [System.Drawing.Point]::new(430, 562)
$labelHeaderRow.AutoSize = $true
$labelHeaderRow.Anchor = 'Bottom, Left'

$numericHeaderRow = [System.Windows.Forms.NumericUpDown]::new()
$numericHeaderRow.Location = [System.Drawing.Point]::new(505, 558)
$numericHeaderRow.Size = [System.Drawing.Size]::new(60, 24)
$numericHeaderRow.Minimum = 1
$numericHeaderRow.Maximum = 1000
$numericHeaderRow.Value = 1
$numericHeaderRow.Anchor = 'Bottom, Left'

$buttonMerge = [System.Windows.Forms.Button]::new()
$buttonMerge.Text = 'Merge'
$buttonMerge.Location = [System.Drawing.Point]::new(12, 590)
$buttonMerge.Size = [System.Drawing.Size]::new(120, 34)
$buttonMerge.Anchor = 'Bottom, Left'

$labelStatus = [System.Windows.Forms.Label]::new()
$labelStatus.Text = ''
$labelStatus.Location = [System.Drawing.Point]::new(144, 599)
$labelStatus.Size = [System.Drawing.Size]::new(508, 20)
$labelStatus.Anchor = 'Bottom, Left, Right'

$buttonMerge.Add_Click({
        if ($listBox.Items.Count -eq 0) {
            $null = [System.Windows.Forms.MessageBox]::Show('No file in the list.', 'Excel / CSV Merger', 'OK', 'Warning')
            return
        }

        # Ensure the ImportExcel module is available
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            $answer = [System.Windows.Forms.MessageBox]::Show('The ImportExcel module is required. Install it now (CurrentUser)?', 'Missing module', 'YesNo', 'Question')
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
                return
            }
            $labelStatus.Text = '[...] Installing ImportExcel module'
            $form.Refresh()
            Install-Module -Name ImportExcel -Scope CurrentUser -Force
        }
        Import-Module -Name ImportExcel

        # Ask for the output file
        $saveDialog = [System.Windows.Forms.SaveFileDialog]::new()
        $saveDialog.Filter = 'Excel workbook (*.xlsx)|*.xlsx'
        $saveDialog.FileName = 'Merged.xlsx'
        if ($saveDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }
        $outputFile = $saveDialog.FileName

        # Export-Excel appends worksheets to an existing workbook, so remove it first
        if (Test-Path -Path $outputFile) {
            Remove-Item -Path $outputFile -Force -Confirm:$false
        }

        try {
            $addSourceFile = $checkBoxSourceFile.Checked
            $headerRow = [int]$numericHeaderRow.Value

            # Parse exclusion rules (one per line, Column=Value); invalid lines are ignored
            [System.Collections.Generic.List[PSCustomObject]]$exclusionRules = @()
            foreach ($ruleLine in ($textBoxExclusions.Text -split "`r?`n")) {
                $ruleLine = $ruleLine.Trim()
                if ((-not $ruleLine) -or ($ruleLine -notmatch '=')) {
                    continue
                }
                $parts = $ruleLine -split '=', 2
                $ruleColumn = $parts[0].Trim()
                $ruleValue = $parts[1].Trim()
                if ($ruleColumn) {
                    $exclusionRules.Add([PSCustomObject]@{
                            Column = $ruleColumn
                            Value  = $ruleValue
                        })
                }
            }
            $excludedRowCount = 0

            if ($radioSheetPerFile.Checked) {
                # -------- Mode: one worksheet per file --------
                [System.Collections.Generic.List[string]]$usedSheetNames = @()
                $processedCount = 0

                foreach ($filePath in $listBox.Items) {
                    $fileName = [System.IO.Path]::GetFileName($filePath)
                    $labelStatus.Text = "[...] Reading $fileName"
                    $form.Refresh()

                    $data = Import-DataFile -Path $filePath -HeaderRow $headerRow

                    if ($exclusionRules.Count -gt 0) {
                        $beforeCount = $data.Count
                        $data = @($data | Where-Object { -not (Test-RowExcluded -Row $_ -ExclusionRules $exclusionRules) })
                        $excludedRowCount += ($beforeCount - $data.Count)
                    }

                    if (-not ($data -and $data.Count -gt 0)) {
                        Write-Warning "No data found in file: $filePath"
                        continue
                    }

                    if ($addSourceFile) {
                        foreach ($row in $data) {
                            $row | Add-Member -NotePropertyName 'SourceFile' -NotePropertyValue $fileName -Force
                        }
                    }

                    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
                    $sheetName = Get-UniqueSheetName -BaseName $baseName -UsedNames $usedSheetNames

                    $exportParams = @{
                        Path          = $outputFile
                        WorksheetName = $sheetName
                        AutoSize      = $true
                        FreezeTopRow  = $true
                        BoldTopRow    = $true
                    }
                    $data | Export-Excel @exportParams
                    $processedCount++
                }

                $summary = "$processedCount file(s) merged into:`n$outputFile"
            }
            else {
                # -------- Mode: everything in a single worksheet --------
                [System.Collections.Generic.List[PSCustomObject]]$allData = @()
                $referenceHeaders = $null
                $referenceFile = $null

                foreach ($filePath in $listBox.Items) {
                    $fileName = [System.IO.Path]::GetFileName($filePath)
                    $labelStatus.Text = "[...] Reading $fileName"
                    $form.Refresh()

                    $data = Import-DataFile -Path $filePath -HeaderRow $headerRow

                    if ($exclusionRules.Count -gt 0) {
                        $beforeCount = $data.Count
                        $data = @($data | Where-Object { -not (Test-RowExcluded -Row $_ -ExclusionRules $exclusionRules) })
                        $excludedRowCount += ($beforeCount - $data.Count)
                    }

                    if (-not ($data -and $data.Count -gt 0)) {
                        Write-Warning "No data found in file: $filePath"
                        continue
                    }

                    $headers = @($data[0].PSObject.Properties.Name)

                    if ($null -eq $referenceHeaders) {
                        $referenceHeaders = $headers
                        $referenceFile = $fileName
                    }
                    else {
                        # Order-insensitive comparison of the header sets
                        $difference = Compare-Object -ReferenceObject $referenceHeaders -DifferenceObject $headers
                        if ($difference) {
                            $differentHeaders = ($difference | ForEach-Object { $_.InputObject }) -join ', '
                            $message = "Different headers between '$referenceFile' and '$fileName':`n$differentHeaders`n`nMerge cancelled."
                            $null = [System.Windows.Forms.MessageBox]::Show($message, 'Incompatible headers', 'OK', 'Error')
                            $labelStatus.Text = '[X] Merge cancelled (incompatible headers)'
                            return
                        }
                    }

                    foreach ($row in $data) {
                        # Reorder columns to match the reference file
                        $orderedRow = $row | Select-Object -Property $referenceHeaders
                        if ($addSourceFile) {
                            $orderedRow | Add-Member -NotePropertyName 'SourceFile' -NotePropertyValue $fileName -Force
                        }
                        $allData.Add($orderedRow)
                    }
                }

                if ($allData.Count -eq 0) {
                    $null = [System.Windows.Forms.MessageBox]::Show('No data found in the files.', 'Excel / CSV Merger', 'OK', 'Warning')
                    $labelStatus.Text = '[!] No data'
                    return
                }

                $exportParams = @{
                    Path          = $outputFile
                    WorksheetName = 'Merged'
                    AutoSize      = $true
                    FreezeTopRow  = $true
                    BoldTopRow    = $true
                }
                $allData | Export-Excel @exportParams

                $rowCount = $allData.Count
                $summary = "$rowCount row(s) merged into:`n$outputFile"
            }

            if ($excludedRowCount -gt 0) {
                $summary = "$summary`n$excludedRowCount row(s) excluded by filters"
            }

            $labelStatus.Text = '[OK] Merge completed'
            $answer = [System.Windows.Forms.MessageBox]::Show("$summary`n`nOpen the file?", 'Merge completed', 'YesNo', 'Information')
            if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
                Invoke-Item -Path $outputFile
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            $null = [System.Windows.Forms.MessageBox]::Show("Error during merge:`n$errorMessage", 'Error', 'OK', 'Error')
            $labelStatus.Text = '[X] Error during merge'
        }
    })

$form.Controls.Add($labelDrop)
$form.Controls.Add($listBox)
$form.Controls.Add($buttonAdd)
$form.Controls.Add($buttonRemove)
$form.Controls.Add($buttonClear)
$form.Controls.Add($groupMode)
$form.Controls.Add($groupExclusions)
$form.Controls.Add($checkBoxSourceFile)
$form.Controls.Add($labelHeaderRow)
$form.Controls.Add($numericHeaderRow)
$form.Controls.Add($buttonMerge)
$form.Controls.Add($labelStatus)

$null = $form.ShowDialog()
$form.Dispose()

return
