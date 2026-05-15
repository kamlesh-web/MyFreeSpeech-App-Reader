Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Speech
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
try { Add-Type -AssemblyName System.IO.Compression.FileSystem } catch { }

# --- Dynamic Paths & Fallbacks ---
$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = $PWD.Path }

$tesseractDllPath = Join-Path $scriptDir "Tesseract.dll"
$tessDataPath = Join-Path $scriptDir "tessdata"

if (Test-Path $tesseractDllPath) {
    Add-Type -Path $tesseractDllPath
} else {
    Write-Warning "Could not find Tesseract.dll. OCR fallback will fail."
}

# --- Global Hotkey Form via C# (C# 5.0 Compatible for PS 5.1) ---
$csharpCode = @"
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

public class ReadOutHotkeyForm : Form {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);
    
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    
    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);

    public event EventHandler TriggerRead;
    public event EventHandler TriggerPrevious;
    public event EventHandler TriggerNext;
    public event EventHandler TriggerPause;
    public event EventHandler TriggerResume;
    public event EventHandler TriggerStop;

    private const int Modifiers = 0x0007; // Ctrl + Alt + Shift
    private bool[] registeredHotkeys = new bool[7];

    public void RegisterReadOutHotkeys() {
        if (!this.IsHandleCreated) return;
        RegisterHotkey(1, 0x52); // R
        RegisterHotkey(2, 0x42); // B
        RegisterHotkey(3, 0x4E); // N
        RegisterHotkey(4, 0x50); // P
        RegisterHotkey(5, 0x55); // U
        RegisterHotkey(6, 0x53); // S
    }

    private void RegisterHotkey(int id, int key) {
        if (!registeredHotkeys[id]) {
            UnregisterHotKey(this.Handle, id);
            registeredHotkeys[id] = RegisterHotKey(this.Handle, id, Modifiers, key);
        }
    }

    public void UnregisterReadOutHotkeys() {
        for (int i = 1; i <= 6; i++) {
            if (registeredHotkeys[i]) {
                UnregisterHotKey(this.Handle, i);
                registeredHotkeys[i] = false;
            }
        }
    }

    protected override void OnHandleCreated(EventArgs e) {
        base.OnHandleCreated(e);
        RegisterReadOutHotkeys();
    }

    protected override void OnHandleDestroyed(EventArgs e) {
        UnregisterReadOutHotkeys();
        base.OnHandleDestroyed(e);
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == 0x0312) {
            int id = m.WParam.ToInt32();
            switch (id) {
                case 1: if (TriggerRead != null) TriggerRead(this, EventArgs.Empty); break;
                case 2: if (TriggerPrevious != null) TriggerPrevious(this, EventArgs.Empty); break;
                case 3: if (TriggerNext != null) TriggerNext(this, EventArgs.Empty); break;
                case 4: if (TriggerPause != null) TriggerPause(this, EventArgs.Empty); break;
                case 5: if (TriggerResume != null) TriggerResume(this, EventArgs.Empty); break;
                case 6: if (TriggerStop != null) TriggerStop(this, EventArgs.Empty); break;
            }
        }
        base.WndProc(ref m);
    }
}
"@

if (-not ("ReadOutHotkeyForm" -as [type])) {
    Add-Type -TypeDefinition $csharpCode -ReferencedAssemblies "System.Windows.Forms"
}

# --- Speech Engine Setup ---
$synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
$synth.Rate = 0
$synth.Volume = 100

$script:paragraphs = @()
$script:currentParagraphIndex = 0
$script:currentSourceName = "Screen"
$script:isReading = $false
$script:selectedVoiceName = $null
$script:speechRate = 0
$script:useOnlineVoices = $false
$script:form = $null
$script:errorLogPath = Join-Path $scriptDir "ScreenReader3.log"
$script:readTimer = $null
$script:hasActiveSpeech = $false
$script:isShuttingDown = $false
$script:isMiniMode = $false
$script:normalBounds = $null
$script:miniDragActive = $false
$script:miniDragStart = [System.Drawing.Point]::Empty
$script:miniWindowStart = [System.Drawing.Point]::Empty

function Write-AppError([string]$message, $errorObject) {
    try {
        $details = if ($errorObject -is [System.Management.Automation.ErrorRecord]) {
            $errorObject.Exception.ToString()
        } elseif ($errorObject -is [System.Exception]) {
            $errorObject.ToString()
        } elseif ($null -ne $errorObject) {
            $errorObject.ToString()
        } else {
            ""
        }

        $entry = @(
            "[$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))] $message"
            $details
            ""
        ) -join [Environment]::NewLine

        Add-Content -LiteralPath $script:errorLogPath -Value $entry -Encoding UTF8
    } catch { }
}

[System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException)
[System.Windows.Forms.Application]::add_ThreadException({
    param($sender, $eventArgs)
    Write-AppError "Unhandled UI thread exception." $eventArgs.Exception
})
[AppDomain]::CurrentDomain.add_UnhandledException({
    param($sender, $eventArgs)
    Write-AppError "Unhandled application exception." $eventArgs.ExceptionObject
})

function Invoke-Ui([scriptblock]$action) {
    if ($script:isShuttingDown) { return }
    try {
        if ($script:form -and -not $script:form.IsDisposed -and $script:form.IsHandleCreated) {
            if ($script:form.InvokeRequired) {
                # GetNewClosure ensures variable scope is safely passed into the UI thread inside PS 5.1
                $callback = { & $action }.GetNewClosure()
                [void]$script:form.BeginInvoke([Action]$callback)
            } else {
                & $action
            }
        }
    } catch { Write-AppError "UI update failed." $_ }
}

function Clamp-Index([int]$index) {
    if ($script:paragraphs.Count -eq 0) { return 0 }
    if ($index -lt 0) { return 0 }
    if ($index -ge $script:paragraphs.Count) { return ($script:paragraphs.Count - 1) }
    return $index
}

function Normalize-ReadableText([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    $text = [regex]::Replace($text, "[\x00-\x08\x0B\x0C\x0E-\x1F]", " ")
    $text = $text -replace "`r`n", "`n"
    $text = $text -replace "`r", "`n"
    $text = [regex]::Replace($text, "[ `t]+`n", "`n")
    $text = [regex]::Replace($text, "`n{3,}", "`n`n")
    return $text.Trim()
}

function Split-TextIntoParagraphs([string]$text) {
    $normalized = Normalize-ReadableText $text
    if ([string]::IsNullOrWhiteSpace($normalized)) { return @() }

    $blocks = [regex]::Split($normalized, "`n\s*`n+")
    $items = New-Object 'System.Collections.Generic.List[string]'

    foreach ($block in $blocks) {
        $clean = [regex]::Replace($block, "`n+", " ").Trim()
        $clean = [regex]::Replace($clean, "\s{2,}", " ")
        if (-not [string]::IsNullOrWhiteSpace($clean)) {
            $items.Add($clean)
        }
    }

    if ($items.Count -le 1) {
        $lineItems = New-Object 'System.Collections.Generic.List[string]'
        foreach ($line in ($normalized -split "`n")) {
            $cleanLine = [regex]::Replace($line, "\s{2,}", " ").Trim()
            if (-not [string]::IsNullOrWhiteSpace($cleanLine)) {
                $lineItems.Add($cleanLine)
            }
        }
        if ($lineItems.Count -gt 1) { return $lineItems.ToArray() }
    }
    return $items.ToArray()
}

function Update-Status {
    Invoke-Ui {
        if (-not $lblStatus) { return }
        if ($script:paragraphs.Count -gt 0) {
            $current = [Math]::Min($script:currentParagraphIndex + 1, $script:paragraphs.Count)
            $lblStatus.Text = "Source: {0}    Paragraph {1} of {2}" -f $script:currentSourceName, $current, $script:paragraphs.Count
        } else {
            $lblStatus.Text = "No text loaded."
        }
    }
}

function Set-ReadingText([string]$text, [string]$sourceName) {
    $loadedParagraphs = @(Split-TextIntoParagraphs $text)
    if ($loadedParagraphs.Count -eq 0) { throw "No readable text was found." }

    $script:paragraphs = $loadedParagraphs
    $script:currentParagraphIndex = 0
    $script:currentSourceName = $sourceName
    $script:isReading = $false
    Update-Status
}

function Update-SelectedVoice {
    if ($comboVoices -and $comboVoices.SelectedItem) {
        $script:selectedVoiceName = $comboVoices.SelectedItem.ToString()
    }
}

function Apply-SpeechSettings {
    $synth.Rate = [Math]::Max(-10, [Math]::Min(10, [int]$script:speechRate))
    if (-not [string]::IsNullOrWhiteSpace($script:selectedVoiceName)) {
        try { $synth.SelectVoice($script:selectedVoiceName) } catch {}
    }
}

function Stop-ReadTimer {
    if ($script:readTimer) { $script:readTimer.Stop() }
}

function Start-ReadTimer {
    if ($script:readTimer -and -not $script:readTimer.Enabled) {
        $script:readTimer.Start()
    }
}

function Cancel-CurrentSpeech {
    try { $synth.SpeakAsyncCancelAll() } catch {}
}

function Wait-ForModifierKeysReleased([int]$timeoutMilliseconds = 1500) {
    $deadline = [DateTime]::Now.AddMilliseconds($timeoutMilliseconds)
    do {
        $shiftDown = ([ReadOutHotkeyForm]::GetAsyncKeyState(0x10) -band 0x8000) -ne 0
        $ctrlDown  = ([ReadOutHotkeyForm]::GetAsyncKeyState(0x11) -band 0x8000) -ne 0
        $altDown   = ([ReadOutHotkeyForm]::GetAsyncKeyState(0x12) -band 0x8000) -ne 0

        if (-not ($shiftDown -or $ctrlDown -or $altDown)) { return $true }
        Start-Sleep -Milliseconds 20
        [System.Windows.Forms.Application]::DoEvents()
    } while ([DateTime]::Now -lt $deadline)
    return $false
}

function Finish-ReadingQuietly {
    $script:isReading = $false
    $script:hasActiveSpeech = $false
    Stop-ReadTimer
    Update-Status
}

function Get-RemainingReadableText {
    if ($script:paragraphs.Count -eq 0) { return "" }
    $script:currentParagraphIndex = Clamp-Index $script:currentParagraphIndex
    $remaining = New-Object 'System.Collections.Generic.List[string]'

    for ($i = $script:currentParagraphIndex; $i -lt $script:paragraphs.Count; $i++) {
        if (-not [string]::IsNullOrWhiteSpace($script:paragraphs[$i])) {
            $remaining.Add($script:paragraphs[$i])
        }
    }
    return [string]::Join("`r`n`r`n", $remaining)
}

function Speak-OnlineText($text, $voiceName) {
    $apiKey = "YOUR_AZURE_API_KEY_HERE"
    if ($apiKey -eq "YOUR_AZURE_API_KEY_HERE") {
        [System.Windows.Forms.MessageBox]::Show("Azure API key missing in script.", "API Notice", 0, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
}

function Speak-CurrentParagraph([switch]$SkipCancel) {
    if ($script:paragraphs.Count -eq 0) { Update-Status; return }

    $script:currentParagraphIndex = Clamp-Index $script:currentParagraphIndex
    $textToRead = Get-RemainingReadableText
    Update-Status

    if ([string]::IsNullOrWhiteSpace($textToRead)) {
        Finish-ReadingQuietly
        return
    }

    if ($script:useOnlineVoices) {
        $script:isReading = $false
        Speak-OnlineText -text $textToRead -voiceName $script:selectedVoiceName
        return
    }

    try {
        if (-not $SkipCancel) {
            $script:isReading = $false
            $script:hasActiveSpeech = $false
            Stop-ReadTimer
            Cancel-CurrentSpeech
        }

        Apply-SpeechSettings
        $script:isReading = $true
        $script:hasActiveSpeech = $true
        $synth.SpeakAsync($textToRead) | Out-Null
        Start-ReadTimer
    } catch {
        Finish-ReadingQuietly
        Write-AppError "SpeakAsync dispatch failed." $_
    }
}

function Start-Reading {
    if ($script:paragraphs.Count -eq 0) { Update-Status; return }
    $script:currentParagraphIndex = Clamp-Index $script:currentParagraphIndex
    Speak-CurrentParagraph
}

function Pause-Reading {
    try {
        if ($synth.State -eq [System.Speech.Synthesis.SynthesizerState]::Speaking) {
            $synth.Pause()
        }
    } catch {}
}

function Resume-Reading {
    try {
        if ($synth.State -eq [System.Speech.Synthesis.SynthesizerState]::Paused) {
            $synth.Resume()
        }
    } catch {}
}

function Stop-Reading {
    $script:isReading = $false
    $script:hasActiveSpeech = $false
    Stop-ReadTimer
    Cancel-CurrentSpeech
    Update-Status
}

function Next-Paragraph {
    if ($script:paragraphs.Count -eq 0) { Update-Status; return }
    if ($script:currentParagraphIndex -lt ($script:paragraphs.Count - 1)) {
        $script:currentParagraphIndex++
        Speak-CurrentParagraph
    } else {
        Finish-ReadingQuietly
    }
}

function Previous-Paragraph {
    if ($script:paragraphs.Count -eq 0) { Update-Status; return }
    if ($script:currentParagraphIndex -gt 0) {
        $script:currentParagraphIndex--
    }
    Speak-CurrentParagraph
}

# --- Screen Capture Logic ---
function Safe-GetClipboardText {
    $text = $null
    try {
        if ([System.Windows.Forms.Clipboard]::ContainsText()) {
            $text = [System.Windows.Forms.Clipboard]::GetText()
        }
    } catch {}
    return $text
}

function Safe-SetClipboardText([string]$text) {
    try {
        if (-not [string]::IsNullOrEmpty($text)) {
            [System.Windows.Forms.Clipboard]::SetText($text)
        } else {
            [System.Windows.Forms.Clipboard]::Clear()
        }
    } catch {}
}

function Get-TextFromFocusedElement([switch]$RequireHighlightedText) {
    $text = $null

    try {
        $focused = [System.Windows.Automation.AutomationElement]::FocusedElement
        if ($focused) {
            $tp = $focused.GetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern)
            $selections = $tp.GetSelection()
            if ($selections -and $selections.Count -gt 0) {
                $range = $selections[0]
                $selectedText = $range.GetText(-1)
                if (-not [string]::IsNullOrWhiteSpace($selectedText)) {
                    $range.MoveEndpointByRange([System.Windows.Automation.Text.TextPatternRangeEndpoint]::End, $tp.DocumentRange, [System.Windows.Automation.Text.TextPatternRangeEndpoint]::End)
                    $text = $range.GetText(-1)
                }
            }
        }
    } catch { }

    if (-not [string]::IsNullOrWhiteSpace($text)) { return $text }

    [void](Wait-ForModifierKeysReleased)
    $hadClipboardText = $false
    $originalClipboard = $null

    try {
        $originalClipboard = Safe-GetClipboardText
        if ($null -ne $originalClipboard) { $hadClipboardText = $true }

        if ($RequireHighlightedText) {
            Safe-SetClipboardText ""
            [System.Windows.Forms.SendKeys]::SendWait("^c")
            Start-Sleep -Milliseconds 200
            $text = Safe-GetClipboardText
        } else {
            Safe-SetClipboardText ""
            [System.Windows.Forms.SendKeys]::SendWait("^+{END}")
            Start-Sleep -Milliseconds 150
            [System.Windows.Forms.SendKeys]::SendWait("^c")
            Start-Sleep -Milliseconds 200
            $text = Safe-GetClipboardText

            if ([string]::IsNullOrWhiteSpace($text)) {
                [System.Windows.Forms.SendKeys]::SendWait("^c")
                Start-Sleep -Milliseconds 200
                $text = Safe-GetClipboardText
            }
        }
    } catch {
        Write-AppError "SendKeys capture failed." $_
    } finally {
        if ($hadClipboardText) { Safe-SetClipboardText $originalClipboard }
    }

    return $text
}

function Read-Continuous([switch]$StopWhenNoSelection) {
    Invoke-Ui {
        $btnRead.Text = "Capturing..."
        $btnRead.Enabled = $false
    }

    try {
        $text = Get-TextFromFocusedElement -RequireHighlightedText:$StopWhenNoSelection
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            Set-ReadingText -text $text -sourceName "Screen"
            Start-Reading
        } elseif ($StopWhenNoSelection) {
            Stop-Reading
        } else {
            Invoke-Ui { $lblStatus.Text = "Could not detect target text." }
        }
    } catch {
        Write-AppError "Read continuous chain error." $_
    } finally {
        Invoke-Ui {
            $btnRead.Text = "&Read From Screen"
            $btnRead.Enabled = $true
            Update-Status
        }
    }
}

function Invoke-ReadButtonAction {
    if ($script:isMiniMode) {
        try {
            $form.Hide()
            Start-Sleep -Milliseconds 250
            Read-Continuous
        } finally {
            Invoke-Ui {
                if (-not $script:isShuttingDown) {
                    $form.Show()
                    $form.Activate()
                }
            }
        }
    } else {
        $form.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
        Start-Sleep -Milliseconds 300
        Read-Continuous
    }
}

# --- GUI Setup ---
$form = New-Object ReadOutHotkeyForm
$script:form = $form
$form.Text = "Continuous Reader (Ctrl+Alt+Shift+R)"
$form.Size = New-Object System.Drawing.Size(640, 310)
$form.MinimumSize = New-Object System.Drawing.Size(640, 310)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true
$form.KeyPreview = $true
$script:allowExit = $false

$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Icon = [System.Drawing.SystemIcons]::Application
$notifyIcon.Text = "ReadOut running"
$notifyIcon.Visible = $false

$trayMenu = New-Object System.Windows.Forms.ContextMenuStrip
$trayRestore = New-Object System.Windows.Forms.ToolStripMenuItem("Restore")
$trayExit = New-Object System.Windows.Forms.ToolStripMenuItem("Exit")

$restoreWindow = {
    Invoke-Ui {
        $form.Show()
        $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        if (Get-Command Set-MiniMode -ErrorAction SilentlyContinue) {
            Set-MiniMode $false
        }
        $form.Activate()
        $notifyIcon.Visible = $false
        $form.RegisterReadOutHotkeys()
    }
}.GetNewClosure()

$trayRestore.Add_Click($restoreWindow)
$notifyIcon.Add_DoubleClick($restoreWindow)
$trayExit.Add_Click({
    $script:allowExit = $true
    $form.Close()
})

[void]$trayMenu.Items.Add($trayRestore)
[void]$trayMenu.Items.Add($trayExit)
$notifyIcon.ContextMenuStrip = $trayMenu

# Attach Hotkey listeners safely
$form.add_TriggerRead({ Read-Continuous -StopWhenNoSelection })
$form.add_TriggerPrevious({ Previous-Paragraph })
$form.add_TriggerNext({ Next-Paragraph })
$form.add_TriggerPause({ Pause-Reading })
$form.add_TriggerResume({ Resume-Reading })
$form.add_TriggerStop({ Stop-Reading })

$form.Add_FormClosing({
    param($sender, $eventArgs)

    if ($eventArgs.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing -and -not $script:allowExit) {
        $eventArgs.Cancel = $true
        $form.Hide()
        $notifyIcon.Visible = $true
        return
    }

    $script:isShuttingDown = $true
    Stop-ReadTimer
    Cancel-CurrentSpeech
    
    try { $script:readTimer.Dispose() } catch {}
    try { $synth.Dispose() } catch {}
    try { $notifyIcon.Dispose() } catch {}
})

# --- NATIVE UI POLLING TIMER ---
$readTimer = New-Object System.Windows.Forms.Timer
$script:readTimer = $readTimer
$readTimer.Interval = 200
$readTimer.Add_Tick({
    if ($script:isShuttingDown) { return }
    try {
        if ($script:isReading) {
            if ($synth.State -eq [System.Speech.Synthesis.SynthesizerState]::Ready) {
                $script:hasActiveSpeech = $false
                Finish-ReadingQuietly
            }
        } else {
            Stop-ReadTimer
        }
    } catch {
        Finish-ReadingQuietly
    }
})

# Voice Type Toggle
$rbLocal = New-Object System.Windows.Forms.RadioButton
$rbLocal.Text = "Local SAPI5 Voices"
$rbLocal.Location = New-Object System.Drawing.Point(24, 18)
$rbLocal.Size = New-Object System.Drawing.Size(150, 24)
$rbLocal.Checked = $true

$rbOnline = New-Object System.Windows.Forms.RadioButton
$rbOnline.Text = "Online Natural Voices"
$rbOnline.Location = New-Object System.Drawing.Point(190, 18)
$rbOnline.Size = New-Object System.Drawing.Size(170, 24)

# Voice Selection Dropdown
$lblVoice = New-Object System.Windows.Forms.Label
$lblVoice.Text = "Voice:"
$lblVoice.Location = New-Object System.Drawing.Point(24, 52)
$lblVoice.Size = New-Object System.Drawing.Size(50, 22)

$comboVoices = New-Object System.Windows.Forms.ComboBox
$comboVoices.Location = New-Object System.Drawing.Point(80, 49)
$comboVoices.Size = New-Object System.Drawing.Size(520, 24)
$comboVoices.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboVoices.Add_SelectedIndexChanged({ Update-SelectedVoice })

# Speech speed control
$lblSpeed = New-Object System.Windows.Forms.Label
$lblSpeed.Text = "Speed: 0"
$lblSpeed.Location = New-Object System.Drawing.Point(24, 88)
$lblSpeed.Size = New-Object System.Drawing.Size(90, 24)

$trackSpeed = New-Object System.Windows.Forms.TrackBar
$trackSpeed.Location = New-Object System.Drawing.Point(110, 82)
$trackSpeed.Size = New-Object System.Drawing.Size(490, 42)
$trackSpeed.Minimum = -10
$trackSpeed.Maximum = 10
$trackSpeed.TickFrequency = 1
$trackSpeed.Value = 0
$trackSpeed.Add_ValueChanged({
    $script:speechRate = [int]$trackSpeed.Value
    $lblSpeed.Text = "Speed: $script:speechRate"
    try { $synth.Rate = $script:speechRate } catch {}
})

function Update-VoiceList {
    $comboVoices.Items.Clear()
    $script:useOnlineVoices = $rbOnline.Checked

    if ($rbLocal.Checked) {
        $localVoices = $synth.GetInstalledVoices() | Where-Object { $_.Enabled } | ForEach-Object { $_.VoiceInfo.Name }
        if ($localVoices) {
            $comboVoices.Items.AddRange($localVoices)
            $comboVoices.SelectedIndex = 0
        }
    } else {
        $onlineVoices = @("en-US-AriaNeural", "en-US-GuyNeural", "en-US-JennyNeural", "en-GB-SoniaNeural")
        $comboVoices.Items.AddRange($onlineVoices)
        $comboVoices.SelectedIndex = 0
    }
    Update-SelectedVoice
}

$rbLocal.Add_CheckedChanged({ Update-VoiceList })
$rbOnline.Add_CheckedChanged({ Update-VoiceList })

# Buttons
$toolTips = New-Object System.Windows.Forms.ToolTip

function New-MiniIcon([string]$kind) {
    $bmp = New-Object System.Drawing.Bitmap -ArgumentList 24, 24
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.Clear([System.Drawing.Color]::Transparent)

    $iconColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    $brush = New-Object System.Drawing.SolidBrush -ArgumentList $iconColor
    $pen = New-Object System.Drawing.Pen -ArgumentList $iconColor, 2.2

    try {
        switch ($kind) {
            "Drag" {
                foreach ($x in @(8, 16)) {
                    foreach ($y in @(6, 12, 18)) {
                        $graphics.FillEllipse($brush, $x, $y, 3, 3)
                    }
                }
            }
            "Read" {
                $graphics.FillRectangle($brush, 4, 9, 4, 6)
                [System.Drawing.Point[]]$points = @(
                    (New-Object System.Drawing.Point -ArgumentList 8, 9),
                    (New-Object System.Drawing.Point -ArgumentList 14, 5),
                    (New-Object System.Drawing.Point -ArgumentList 14, 19),
                    (New-Object System.Drawing.Point -ArgumentList 8, 15)
                )
                $graphics.FillPolygon($brush, $points)
                $graphics.DrawArc($pen, 14, 8, 5, 8, -45, 90)
                $graphics.DrawArc($pen, 14, 5, 9, 14, -45, 90)
            }
            "Previous" {
                $graphics.FillRectangle($brush, 5, 6, 3, 13)
                [System.Drawing.Point[]]$points = @(
                    (New-Object System.Drawing.Point -ArgumentList 18, 5),
                    (New-Object System.Drawing.Point -ArgumentList 18, 19),
                    (New-Object System.Drawing.Point -ArgumentList 8, 12)
                )
                $graphics.FillPolygon($brush, $points)
            }
            "Next" {
                [System.Drawing.Point[]]$points = @(
                    (New-Object System.Drawing.Point -ArgumentList 6, 5),
                    (New-Object System.Drawing.Point -ArgumentList 6, 19),
                    (New-Object System.Drawing.Point -ArgumentList 16, 12)
                )
                $graphics.FillPolygon($brush, $points)
                $graphics.FillRectangle($brush, 17, 6, 3, 13)
            }
            "Pause" {
                $graphics.FillRectangle($brush, 7, 5, 4, 14)
                $graphics.FillRectangle($brush, 14, 5, 4, 14)
            }
            "Resume" {
                [System.Drawing.Point[]]$points = @(
                    (New-Object System.Drawing.Point -ArgumentList 8, 5),
                    (New-Object System.Drawing.Point -ArgumentList 8, 19),
                    (New-Object System.Drawing.Point -ArgumentList 19, 12)
                )
                $graphics.FillPolygon($brush, $points)
            }
            "Stop" {
                $graphics.FillRectangle($brush, 7, 7, 11, 11)
            }
            "Restore" {
                $graphics.DrawRectangle($pen, 6, 9, 11, 10)
                $graphics.DrawRectangle($pen, 9, 6, 11, 10)
            }
        }
    } finally {
        $pen.Dispose()
        $brush.Dispose()
        $graphics.Dispose()
    }

    return $bmp
}

function New-MiniButton([string]$kind, [string]$toolTip, [scriptblock]$clickAction) {
    $button = New-Object System.Windows.Forms.Button
    $button.Size = New-Object System.Drawing.Size(34, 34)
    $button.Margin = New-Object System.Windows.Forms.Padding(2)
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderSize = 0
    $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(78, 78, 78)
    $button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(43, 121, 210)
    $button.BackColor = [System.Drawing.Color]::FromArgb(56, 56, 56)
    $button.Image = New-MiniIcon $kind
    $button.TabStop = $false
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $button.AccessibleName = $toolTip
    $toolTips.SetToolTip($button, $toolTip)
    if ($clickAction) { $button.Add_Click($clickAction) }
    return $button
}

function Start-MiniDrag($sender, $eventArgs) {
    if ($eventArgs.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
    $script:miniDragActive = $true
    $script:miniDragStart = [System.Windows.Forms.Cursor]::Position
    $script:miniWindowStart = $form.Location
    $sender.Capture = $true
}

function Move-MiniDrag($sender, $eventArgs) {
    if (-not $script:miniDragActive) { return }
    $current = [System.Windows.Forms.Cursor]::Position
    $deltaX = $current.X - $script:miniDragStart.X
    $deltaY = $current.Y - $script:miniDragStart.Y
    $form.Location = New-Object System.Drawing.Point (($script:miniWindowStart.X + $deltaX), ($script:miniWindowStart.Y + $deltaY))
}

function Stop-MiniDrag($sender, $eventArgs) {
    $script:miniDragActive = $false
    if ($sender) { $sender.Capture = $false }
}

function Add-MiniDragHandlers($control) {
    $control.Add_MouseDown({ param($sender, $eventArgs) Start-MiniDrag $sender $eventArgs })
    $control.Add_MouseMove({ param($sender, $eventArgs) Move-MiniDrag $sender $eventArgs })
    $control.Add_MouseUp({ param($sender, $eventArgs) Stop-MiniDrag $sender $eventArgs })
}

function Set-MiniMode([bool]$enabled) {
    if ($enabled -eq $script:isMiniMode) { return }

    if ($enabled) {
        if ($form.WindowState -ne [System.Windows.Forms.FormWindowState]::Normal) {
            $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        }

        $script:normalBounds = $form.Bounds
        $script:isMiniMode = $true

        foreach ($control in $normalControls) { $control.Visible = $false }
        $miniPanel.Visible = $true
        $miniPanel.BringToFront()

        $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        $form.MinimumSize = New-Object System.Drawing.Size(330, 54)
        $form.Size = New-Object System.Drawing.Size(330, 54)
        $form.Text = "ReadOut Mini Mode"

        if ($script:normalBounds) {
            $x = $script:normalBounds.Left + [Math]::Max(0, [int](($script:normalBounds.Width - $form.Width) / 2))
            $y = $script:normalBounds.Top + [Math]::Max(0, [int](($script:normalBounds.Height - $form.Height) / 2))
            $form.Location = New-Object System.Drawing.Point($x, $y)
        }

        $chkMiniMode.Checked = $true
    } else {
        $script:isMiniMode = $false
        $miniPanel.Visible = $false
        foreach ($control in $normalControls) { $control.Visible = $true }

        $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
        $form.MinimumSize = New-Object System.Drawing.Size(640, 310)
        if ($script:normalBounds) {
            $form.Bounds = $script:normalBounds
        } else {
            $form.Size = New-Object System.Drawing.Size(640, 310)
        }
        $form.Text = "Continuous Reader (Ctrl+Alt+Shift+R)"
        $chkMiniMode.Checked = $false
    }
}

$btnRead = New-Object System.Windows.Forms.Button
$btnRead.Text = "&Read From Screen"
$btnRead.Location = New-Object System.Drawing.Point(24, 136)
$btnRead.Size = New-Object System.Drawing.Size(128, 36)
$btnRead.Add_Click({ Invoke-ReadButtonAction })

$btnPrevious = New-Object System.Windows.Forms.Button
$btnPrevious.Text = "Pre&vious"
$btnPrevious.Location = New-Object System.Drawing.Point(162, 136)
$btnPrevious.Size = New-Object System.Drawing.Size(82, 36)
$btnPrevious.Add_Click({ Previous-Paragraph })

$btnNext = New-Object System.Windows.Forms.Button
$btnNext.Text = "&Next"
$btnNext.Location = New-Object System.Drawing.Point(254, 136)
$btnNext.Size = New-Object System.Drawing.Size(82, 36)
$btnNext.Add_Click({ Next-Paragraph })

$btnPause = New-Object System.Windows.Forms.Button
$btnPause.Text = "&Pause"
$btnPause.Location = New-Object System.Drawing.Point(346, 136)
$btnPause.Size = New-Object System.Drawing.Size(70, 36)
$btnPause.Add_Click({ Pause-Reading })

$btnResume = New-Object System.Windows.Forms.Button
$btnResume.Text = "Res&ume"
$btnResume.Location = New-Object System.Drawing.Point(426, 136)
$btnResume.Size = New-Object System.Drawing.Size(82, 36)
$btnResume.Add_Click({ Resume-Reading })

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = "&Stop"
$btnStop.Location = New-Object System.Drawing.Point(518, 136)
$btnStop.Size = New-Object System.Drawing.Size(82, 34)
$btnStop.Add_Click({ Stop-Reading })

$chkMiniMode = New-Object System.Windows.Forms.CheckBox
$chkMiniMode.Text = "Mini Mode"
$chkMiniMode.Appearance = [System.Windows.Forms.Appearance]::Button
$chkMiniMode.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$chkMiniMode.Location = New-Object System.Drawing.Point(488, 14)
$chkMiniMode.Size = New-Object System.Drawing.Size(112, 30)
$chkMiniMode.Add_CheckedChanged({
    if ($chkMiniMode.Checked -ne $script:isMiniMode) {
        Set-MiniMode $chkMiniMode.Checked
    }
})

$toolTips.SetToolTip($btnRead, "Alt+R natively; Ctrl+Alt+Shift+R globally")
$toolTips.SetToolTip($btnPrevious, "Alt+V natively; Ctrl+Alt+Shift+B globally")
$toolTips.SetToolTip($btnNext, "Alt+N natively; Ctrl+Alt+Shift+N globally")
$toolTips.SetToolTip($btnPause, "Alt+P natively; Ctrl+Alt+Shift+P globally")
$toolTips.SetToolTip($btnResume, "Alt+U natively; Ctrl+Alt+Shift+U globally")
$toolTips.SetToolTip($btnStop, "Alt+S natively; Ctrl+Alt+Shift+S globally")
$toolTips.SetToolTip($chkMiniMode, "Switch to the draggable mini widget")

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "No text loaded."
$lblStatus.Location = New-Object System.Drawing.Point(24, 190)
$lblStatus.Size = New-Object System.Drawing.Size(596, 24)

$infoLabel = New-Object System.Windows.Forms.Label
$infoLabel.Text = "Tip: Ctrl+Alt+Shift shortcuts operate globally. Closing sends app directly to the system tray."
$infoLabel.Location = New-Object System.Drawing.Point(24, 228)
$infoLabel.Size = New-Object System.Drawing.Size(596, 22)
$infoLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$miniPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$miniPanel.Visible = $false
$miniPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$miniPanel.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 34)
$miniPanel.Padding = New-Object System.Windows.Forms.Padding(7, 8, 7, 7)
$miniPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$miniPanel.WrapContents = $false
$toolTips.SetToolTip($miniPanel, "Drag the widget from the empty background or grip button")

$btnMiniDrag = New-MiniButton "Drag" "Drag widget" $null
$btnMiniDrag.Cursor = [System.Windows.Forms.Cursors]::SizeAll
$btnMiniRead = New-MiniButton "Read" "Read from screen" { Invoke-ReadButtonAction }
$btnMiniPrevious = New-MiniButton "Previous" "Previous paragraph" { Previous-Paragraph }
$btnMiniPause = New-MiniButton "Pause" "Pause reading" { Pause-Reading }
$btnMiniResume = New-MiniButton "Resume" "Resume reading" { Resume-Reading }
$btnMiniStop = New-MiniButton "Stop" "Stop reading" { Stop-Reading }
$btnMiniNext = New-MiniButton "Next" "Next paragraph" { Next-Paragraph }
$btnMiniRestore = New-MiniButton "Restore" "Exit Mini Mode" { Set-MiniMode $false }

Add-MiniDragHandlers $miniPanel
Add-MiniDragHandlers $btnMiniDrag

$miniPanel.Controls.AddRange(@(
    $btnMiniDrag, $btnMiniRead, $btnMiniPrevious, $btnMiniPause,
    $btnMiniResume, $btnMiniStop, $btnMiniNext, $btnMiniRestore
))

$normalControls = @(
    $rbLocal, $rbOnline, $lblVoice, $comboVoices, $lblSpeed, $trackSpeed,
    $btnRead, $btnPrevious, $btnNext, $btnPause, $btnResume, $btnStop, $chkMiniMode,
    $lblStatus, $infoLabel
)

$form.Controls.AddRange($normalControls + @($miniPanel))

Update-VoiceList
Update-Status

[void]$form.ShowDialog()
