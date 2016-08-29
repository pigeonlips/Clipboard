function Get-ClipboardText {

    [CmdletBinding(ConfirmImpact='None', SupportsShouldProcess=$false)] # to support -OutVariable and -Verbose
    param ()

    Add-Type -AssemblyName System.Windows.Forms
    if ([threading.thread]::CurrentThread.ApartmentState.ToString() -eq 'STA') {
        Write-Verbose "STA mode: Using [Windows.Forms.Clipboard] directly."
        # To be safe, we explicitly specify that Unicode (UTF-16) be used - older platforms may default to ANSI.
        [System.Windows.Forms.Clipboard]::GetText([System.Windows.Forms.TextDataFormat]::UnicodeText)
    } else {
        Write-Verbose "MTA mode: Using a [System.Windows.Forms.TextBox] instance for clipboard access."
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Multiline = $tru
        $tb.Paste()
        $tb.Text
    }
}


function Set-ClipboardText() {

    # !! We do NOT use an advanced function here, because we want to use $Input.
    Param(
        [PSObject] $InputObject
        , [switch] $Verbose
    )

    if ($args.count) { throw "Unrecognized parameter(s) specified." }

    # Out-string invariably adds an extra terminating newline, which we want to strip.
    $stripTrailingNewline = $true
    if ($InputObject) { # Direct argument specified.
        if ($InputObject -is [string]) {
            $stripTrailingNewline = $false
            $text = $InputObject # Already a string, use as is.
        } else {
            $text = $InputObject | Out-String # Convert to string as it would display in the console
        }
    } else { # Use pipeline input, if present.
        $text = $input | Out-String # convert ENTIRE pipeline input to string as it would display in the console
    }
    if ($stripTrailingNewline -and $text.Length -gt 2) {
        $text = $text.Substring(0, $text.Length - 2)
    }
    Add-Type -AssemblyName System.Windows.Forms
    if ([threading.thread]::CurrentThread.ApartmentState.ToString() -eq 'STA') {
        if ($Verbose) { # Simulate verbose output.
            $fgColor = 'Cyan'
            if ($PSVersionTable.PSVersion.major -le 2) { $fgColor = 'Yellow' }
            Write-Host -ForegroundColor $fgColor "STA mode: Using [Windows.Forms.Clipboard] directly." 
        }
        if (-not $text) { $text = "`0" } # Strangely, SetText() breaks with an empty string, claiming $null was passed -> use a null char.
        [System.Windows.Forms.Clipboard]::SetText($text, [System.Windows.Forms.TextDataFormat]::UnicodeText)

    } else {
        if ($Verbose) { # Simulate verbose output. 
            $fgColor = 'Cyan'
            if ($PSVersionTable.PSVersion.major -le 2) { $fgColor = 'Yellow' }
            Write-Host -ForegroundColor $fgColor "MTA mode: Using a [System.Windows.Forms.TextBox] instance for clipboard access."
        }
        if (-not $text) { 
            # !! This approach cannot set the clipboard to an empty string: the text box must
            # !! must be *non-empty* in order to copy something. A null character doesn't work.
            # !! We use the least obtrusive alternative - a newline - and issue a warning.
            $text = "`r`n"
            Write-Warning "Setting clipboard to empty string not supported in MTA mode; using newline instead."
        }
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Multiline = $true
        $tb.Text = $text
        $tb.SelectAll()
        $tb.Copy()
    }
}
