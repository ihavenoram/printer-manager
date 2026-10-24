# Printer Manager 2.2.1 - https://ihavenoram.com - https://github.com/ihavenoram/printer-manager
# MIT licence, (c) 2026 ihavenoram.com. If you publish something built on this, please credit ihavenoram.com.
<#
=========================================================
Script Name: PrinterManager-Gui
Description: SYSTEM-capable WinForms printer console. Built to run in ScreenConnect
             Backstage (Run As: SYSTEM) or an ordinary interactive session.

               - Shows the LOCAL queues merged with the SELECTED USER's mapped
                 printers (read from HKEY_USERS - Get-Printer run as SYSTEM cannot
                 see a user's mapped printers at all).
               - Rename / delete a printer, test whether the port or server is
                 reachable, print a test page.
               - Identify the active user and set their default printer, so it
                 actually applies AND sticks.
               - Add or re-add a printer connection for a user.
               - Spooler control, stuck-job clearing, orphaned spool-file purge.
               - Point and Print policy read-out with the Windows error-code decoder.
               - Print server enumeration, or paste JSON from either collector.
               - One-click printmanagement.msc.

WHY A TARGET-USER PICKER
             Every per-user printer write (SetDefaultPrinter, printui.dll,
             Remove-Printer on a UNC) run from SYSTEM silently targets SYSTEM's own
             profile with no error. This console therefore always acts through the
             chosen user's session - see Set-DefaultPrinterForUser.

Requires:    Windows PowerShell 5.1 (WinForms). Launch via Launch-PrinterManager.cmd
             (forces STA). Self-contained: the printer toolkit module is inlined
             below, so nothing else needs to sit beside it.
=========================================================
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [switch]$NoHideConsole
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

function Get-MsgOwner {
    <#
    .SYNOPSIS
        The window a message box should be parented to.
    .DESCRIPTION
        An UNOWNED MessageBox has no parent to sit in front of. On a normal
        desktop that is merely untidy; inside a ScreenConnect Backstage session it
        fights the host window for z-order and visibly BLINKS. Parenting each
        dialog to the top-most open form stops that, and keeps it modal to the
        right window when a sub-dialog is open.

        Defined THIS early on purpose: the toolkit-missing message box below
        calls it before any form exists, and while it lived further down the
        file that dialog - the only thing that explains a missing module - died
        with "Get-MsgOwner is not recognized" instead of ever appearing.
    #>
    $f = @([System.Windows.Forms.Application]::OpenForms)
    if ($f.Count) { return $f[$f.Count - 1] }
    return $null   # Show($null, ...) behaves as the unowned overload
}

# Win32 DISABLES an owner window for as long as an owned modal dialog is up.
# That native state is the authoritative "a dialog is open" signal, and unlike a
# managed flag it needs no bookkeeping at 40-odd call sites. Control.Enabled does
# NOT reflect it - the managed property is untouched by the native EnableWindow.
Add-Type -Namespace PrtWin32 -Name Modal -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr hWnd);
'@ -ErrorAction SilentlyContinue

function Test-ModalOpen {
    <#
    .SYNOPSIS
        Is a modal dialog currently displayed over the main window?
    .DESCRIPTION
        Used to stop the 200ms poll from repainting the window UNDERNEATH an open
        dialog. Appending to the transcript and rewriting the status bar five
        times a second invalidates the main form, and in a ScreenConnect
        Backstage session that shows up as the dialog's own text blanking out and
        coming back - unreadable while you are trying to read it.

        Fails OPEN (returns false) if anything goes wrong: a missed suppression
        is a cosmetic flicker, whereas a stuck 'true' would freeze the transcript
        and the watchdog for the rest of the session.
    #>
    try {
        if (-not $form) { return $false }
        if (-not $form.IsHandleCreated) { return $false }
        return (-not [PrtWin32.Modal]::IsWindowEnabled($form.Handle))
    } catch { return $false }
}

# An unhandled error in ANY click handler otherwise pops the raw .NET
# "Unhandled exception has occurred in your application" dialog, complete with a
# stack trace and a Continue/Quit choice - alarming in front of a client, and it
# tells the technician nothing useful. Route those to the log and a plain message
# box instead. Belt and braces: individual handlers still catch their own errors.
try {
    [System.Windows.Forms.Application]::SetUnhandledExceptionMode(
        [System.Windows.Forms.UnhandledExceptionMode]::CatchException)
    [System.Windows.Forms.Application]::add_ThreadException({
        param($sender, $e)
        $msg = $e.Exception.Message
        try { Write-Log ("Unhandled error: " + $msg) 'ERROR' } catch {}
        try { if ($script:Busy) { Set-Busy $false } } catch {}
        [void](Show-Msg -Text "Something went wrong in the console:`n`n$msg`n`nThe window is still usable - the action was abandoned." -Title 'Printer Manager' -Icon 'Warning')
    })
} catch {}


# --------------------------------------------------------------------------
# Hide the console IN-PROCESS.
# NOT -WindowStyle Hidden on the launcher: that switch hides the FORM as well,
# not just the console host, and the GUI never appears. Hiding the console
# window by handle leaves the WinForms window untouched.
# --------------------------------------------------------------------------
if (-not $NoHideConsole) {
    try {
        Add-Type -Namespace PrtWin32 -Name Native -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@ -ErrorAction Stop
        $h = [PrtWin32.Native]::GetConsoleWindow()
        if ($h -ne [IntPtr]::Zero) { [void][PrtWin32.Native]::ShowWindow($h, 0) }
    } catch { }
}


#region TOOLKIT --- INLINED by Build-PrinterToolkit.ps1 from PrinterToolkit.psm1 v2.2.0 on 2026-10-09 14:31.
# The printer toolkit module, inlined so this one file is all you need.
$script:ToolkitSource = @'
$script:ToolkitVersion = '2.2.0'

# Count of bounded calls that timed out and had to be abandoned (see
# Invoke-WithTimeout). Non-zero means some print data could not be read because a
# call never returned - reported rather than passed off as a clean empty result.
$script:AbandonedRunspaces = 0

# Virtual/software printer detection. Matches on Name OR DriverName so third-party
# PDF writers with generic driver names are caught. Sourced from the Harvester
# implementation in PC_TRANSFER\Migration - the best of the four filters that
# existed across the tree.
$script:VirtualNameRx   = 'Print to PDF|XPS|OneNote|Fax|FlexiPDF|CutePDF|PDF24|PDFCreator|Bullzip|doPDF|novaPDF|Foxit.*Printer|PrimoPDF|Nitro.*PDF Creator|Snagit|Webex|Send To|Quicken|Any2PDF|ZAN Image|Microsoft Print'
$script:VirtualDriverRx = 'ActMask|Microsoft Shared Fax|Microsoft Print To PDF|Send to Microsoft OneNote|Microsoft XPS|PDFCreator|Bullzip|novaPDF|doPDF|Foxit Reader PDF Printer|Foxit PhantomPDF|Nitro PDF Creator|PrimoPDF|Snagit|Webex'

# --------------------------------------------------------------------------
# The JUNK QUEUE catalogue - a DELETE list, and nothing like the filter above.
#
# $script:VirtualNameRx is a DISPLAY filter and is deliberately greedy: it
# matches CutePDF, Snagit, Microsoft Print to PDF and anything else with 'XPS'
# or 'OneNote' in the name. Pointing a remove button at it would destroy
# configuration people rely on. So this is a separate, narrow, EXACT-match
# table, and the two must never be conflated.
#
# Each entry matches on an exact queue name OR an exact driver name (localised
# Windows renames the queue but not always the driver, and vice versa).
# --------------------------------------------------------------------------
$script:JunkPrinterCatalogue = @(
    @{ Id      = 'XPS'
       Label   = 'Microsoft XPS Document Writer'
       Names   = @('Microsoft XPS Document Writer')
       Drivers = @('Microsoft XPS Document Writer v4', 'Microsoft XPS Document Writer')
       Feature = 'Printing-XPSServices-Features'
       Why     = 'Prints to an .xps/.oxps file that almost nothing opens. Backed by a Windows optional feature, so deleting the queue alone is temporary.' }

    @{ Id      = 'FAX'
       Label   = 'Fax'
       Names   = @('Fax')
       Drivers = @('Microsoft Shared Fax Driver')
       Feature = 'FaxServicesClientPackage'
       Why     = 'The Windows Fax queue - useless without a fax modem or service. Backed by a Windows optional feature, so deleting the queue alone is temporary.' }

    @{ Id      = 'ONENOTE10'
       Label   = 'OneNote for Windows 10'
       Names   = @('OneNote for Windows 10')
       Drivers = @()
       Feature = $null
       Why     = 'Print target of the retired OneNote for Windows 10 store app. NOT the same queue as OneNote (Desktop) / Send To OneNote 2016, which is never touched.' }
)

# Second gate. Even if a catalogue entry were ever widened by accident, nothing
# on this list can be removed by the junk sweep. The OneNote entries are the
# reason it exists: 'OneNote' as a substring would take the desktop queue with it.
$script:JunkPrinterNeverRemove = @(
    'OneNote (Desktop)', 'Send To OneNote 2016', 'Send to OneNote 2016',
    'Microsoft Print to PDF', 'Microsoft Print To PDF'
)
$script:JunkPrinterNeverRemoveDrivers = @(
    'Send to Microsoft OneNote 16 Driver', 'Microsoft Print To PDF'
)

# Ports probed for "is this printer alive". 9100 RAW, 515 LPR, 631 IPP.
$script:PrinterProbePorts = @(9100, 515, 631)

# Built-in service SIDs that must never be treated as user profiles.
$script:ServiceSids = @('S-1-5-18', 'S-1-5-19', 'S-1-5-20')


# ---------------------------------------------------------------------------
# region  Version + small utilities
# ---------------------------------------------------------------------------

function Get-PrinterToolkitVersion {
    <# .SYNOPSIS Toolkit version string, shown in the GUI banner and report headers. #>
    [CmdletBinding()] param()
    $script:ToolkitVersion
}

function Invoke-WithTimeout {
    <#
    .SYNOPSIS
        Run a self-contained scriptblock with a hard wall-clock cap, in its own
        runspace. Returns $null on timeout.
    .DESCRIPTION
        Win32_Printer blocks 30-45 SECONDS per dead network queue, which is what
        turns an Atera printer run into "runs forever then Failed". Every risky
        enumeration goes through here.

        Uses a runspace rather than Start-Job: Start-Job costs ~1s of process
        spin-up per call, and this is called several times per run.

        The scriptblock MUST be self-contained (cmdlets only, no closures over
        module functions) - it executes in a fresh runspace with no session state.
    .OUTPUTS
        Whatever the scriptblock emits, or $null if it exceeded -TimeoutSec.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [int]$TimeoutSec = 15,
        [hashtable]$Parameters
    )
    $rs = $null; $ps = $null
    try {
        $rs = [runspacefactory]::CreateRunspace()
        $rs.Open()
        $ps = [powershell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript($ScriptBlock.ToString())
        if ($Parameters) { foreach ($k in $Parameters.Keys) { [void]$ps.AddParameter($k, $Parameters[$k]) } }
        $async = $ps.BeginInvoke()

        if ($async.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds($TimeoutSec))) {
            # Completed on time - safe to tear down synchronously.
            $out = $null
            try { $out = $ps.EndInvoke($async) } catch { $out = $null }
            try { $ps.Dispose() } catch {}
            try { $rs.Close(); $rs.Dispose() } catch {}
            return $out
        }

        # ---- TIMED OUT ----------------------------------------------------
        # ABANDON the pipeline. Do NOT call Stop(), Dispose() or Close() here:
        # all three block until the pipeline actually ends, and the thing we are
        # timing out on is precisely a call that will not end - a print RPC to a
        # dead server, or WMI on a dead network queue (30-45s PER QUEUE).
        # Tearing down synchronously made the "timeout" a lie: measured, a 3s
        # timeout over a 30s uninterruptible body still took 30.1s to return,
        # which is what left the Backstage GUI stuck on an hourglass forever.
        #
        # BeginStop is the async counterpart and returns immediately; when the
        # wedged call finally unblocks, the runspace tears itself down. Until
        # then it leaks one thread, which is a far better outcome than hanging
        # the caller. Nothing below waits on it.
        try { [void]$ps.BeginStop($null, $null) } catch {}
        $script:AbandonedRunspaces = [int]$script:AbandonedRunspaces + 1
        return $null
    } catch {
        # Only reached for setup failures (runspace creation etc.), where the
        # objects are not wedged and can be released normally.
        if ($ps) { try { $ps.Dispose() } catch {} }
        if ($rs) { try { $rs.Close(); $rs.Dispose() } catch {} }
        return $null
    }
}

function Get-AbandonedRunspaceCount {
    <#
    .SYNOPSIS
        How many bounded calls timed out and were abandoned this session.
    .DESCRIPTION
        Each one is a print call that never returned (dead server or dead queue).
        Surfaced so a run that looks "fine but slow" still reports that data is
        missing rather than silently showing an empty result as though it were a
        clean read.
    #>
    [CmdletBinding()] param()
    [int]$script:AbandonedRunspaces
}

function Test-IsLocalHost {
    <#
    .SYNOPSIS
        True if $Name refers to THIS machine - by well-known name, resolved IP, or
        FQDN. Catches DNS aliases/CNAMEs (e.g. "prtserver") pointing at the box we
        are running on, so local methods can be used instead of failing on RPC.
    .NOTES
        Carried over unchanged from Export_Print_Server_Printers_To_JSON.ps1.
    #>
    [CmdletBinding()]
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return $true }
    $n = $Name.TrimStart('\').Trim()
    if ($n -in @($env:COMPUTERNAME, '.', 'localhost', '127.0.0.1', '::1')) { return $true }

    try {
        $localIps = @([System.Net.Dns]::GetHostAddresses([System.Net.Dns]::GetHostName()) |
                      ForEach-Object { $_.IPAddressToString })
        $localIps += @('127.0.0.1', '::1')
        $targetIps = @([System.Net.Dns]::GetHostAddresses($n) | ForEach-Object { $_.IPAddressToString })
        foreach ($t in $targetIps) { if ($localIps -contains $t) { return $true } }
    } catch { }

    try {
        $localFqdn  = [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName
        $targetFqdn = [System.Net.Dns]::GetHostEntry($n).HostName
        if ($localFqdn -and $targetFqdn -and ($localFqdn -eq $targetFqdn)) { return $true }
    } catch { }

    return $false
}

function Test-IsVirtualPrinter {
    <# .SYNOPSIS True for PDF writers, XPS, Fax, OneNote and similar software queues. #>
    [CmdletBinding()]
    param([string]$Name, [string]$DriverName)
    if ($Name       -and $Name       -match $script:VirtualNameRx)   { return $true }
    if ($DriverName -and $DriverName -match $script:VirtualDriverRx) { return $true }
    return $false
}

function Test-IsIppClassDriver {
    <#
    .SYNOPSIS
        True for the generic "Microsoft IPP Class Driver" (and vendor variants).
    .DESCRIPTION
        Windows binds this class driver when a printer is auto-added over WSD. It
        prints, but usually lacks trays/duplex/finishing/accounting, AND it
        negotiates its page-description language over the IPP channel - so moved
        onto a bare RAW:9100 socket it often produces BLANK OR GARBLED pages.
        That is why a port fix and a driver fix have to happen together.

        Deliberately a loose substring match: this is the single canonical test
        used throughout the EPSON_PRINT_DRIVERS toolkit, and it catches
        "Microsoft IPP Class Driver" plus any vendor-badged variant.
    #>
    [CmdletBinding()]
    param([string]$DriverName)
    if (-not $DriverName) { return $false }
    return ($DriverName -like '*IPP Class*')
}

function Get-PrinterStateText {
    <# .SYNOPSIS Win32_Printer PrinterState code -> human text. #>
    [CmdletBinding()]
    param($StateCode)
    if ($null -eq $StateCode) { return '' }
    $states = @{
        0='Ready'; 1='Paused'; 2='Error'; 3='Pending Deletion'; 4='Paper Jam'
        5='Paper Out'; 6='Manual Feed'; 7='Paper Problem'; 8='Offline'; 9='IO Active'
        10='Busy'; 11='Printing'; 12='Output Bin Full'; 13='Not Available'; 14='Waiting'
        15='Processing'; 16='Initialization'; 17='Warming Up'; 18='Toner Low'; 19='No Toner'
        20='Page Punt'; 21='User Intervention Required'; 22='Out of Memory'; 23='Door Open'
        24='Server Unknown'; 25='Power Save'
    }
    $c = [int]$StateCode
    if ($states.ContainsKey($c)) { return $states[$c] }
    return "Unknown ($c)"
}

function Get-PrinterConnectionInfo {
    <#
    .SYNOPSIS
        Classify a port into Type (WSD/USB/Local/IPP/Network/Other), Host and Protocol.
    .NOTES
        Carried over from EPSON_PRINT_DRIVERS\Printer_Inventory_Report.ps1 rather
        than writing a fourth classifier.
    #>
    [CmdletBinding()]
    param([string]$PortName, $TcpPort)

    if (-not $PortName) { return [pscustomobject]@{ Type='Unknown'; Host=''; Protocol='' } }

    if ($PortName -like 'WSD*' -or $PortName -like '*WSD-*') {
        return [pscustomobject]@{ Type='WSD'; Host=''; Protocol='WSD' }
    }
    if ($PortName -match '^USB')         { return [pscustomobject]@{ Type='USB';   Host=''; Protocol='USB' } }
    if ($PortName -match '^(LPT|COM)\d') { return [pscustomobject]@{ Type='Local'; Host=''; Protocol=$PortName } }
    if ($PortName -match '^(FILE|PORTPROMPT|nul)' ) { return [pscustomobject]@{ Type='Local'; Host=''; Protocol=$PortName } }
    if ($PortName -match '^IPP' -or $PortName -match '^https?://') {
        $h = if ($PortName -match '://([^:/]+)') { $matches[1] } else { '' }
        return [pscustomobject]@{ Type='IPP'; Host=$h; Protocol='IPP' }
    }

    if ($TcpPort) {
        # Only a port that actually carries a host address is a TCP/IP port. A
        # local/virtual monitor port also comes back from Get-PrinterPort but has
        # no address, and must not be mistaken for a network queue.
        $probeHost = ''
        foreach ($prop in @('HostAddress', 'PrinterHostAddress')) {
            try { if ($TcpPort.$prop) { $probeHost = [string]$TcpPort.$prop; break } } catch {}
        }
        if (-not $probeHost) {
            $m2 = [regex]::Match($PortName, '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}')
            if ($m2.Success) { return [pscustomobject]@{ Type='Network'; Host=$m2.Value; Protocol='TCP/IP' } }
            return [pscustomobject]@{ Type='Local'; Host=''; Protocol=$PortName }
        }
        $proto = 'TCP/IP'
        $pnum  = $null
        try { $pnum = $TcpPort.PortNumber } catch {}
        $rawProto = $null
        try { $rawProto = $TcpPort.Protocol } catch {}
        # Win32_TCPIPPrinterPort.Protocol is 1=RAW 2=LPR; MSFT_TcpIpPrinterPort.Protocol
        # is the string 'RAW'/'LPR'. Handle both shapes.
        if ("$rawProto" -eq '1' -or "$rawProto" -eq 'RAW') { $proto = if ($pnum) { "RAW:$pnum" } else { 'RAW' } }
        elseif ("$rawProto" -eq '2' -or "$rawProto" -eq 'LPR') { $proto = 'LPR' }
        $hostAddr = ''
        foreach ($prop in @('HostAddress', 'PrinterHostAddress')) {
            try { if ($TcpPort.$prop) { $hostAddr = [string]$TcpPort.$prop; break } } catch {}
        }
        return [pscustomobject]@{ Type='Network'; Host=$hostAddr; Protocol=$proto }
    }

    # No port object: pull an IP out of the port name if one is embedded.
    $m = [regex]::Match($PortName, '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}')
    if ($m.Success) { return [pscustomobject]@{ Type='Network'; Host=$m.Value; Protocol='TCP/IP' } }

    return [pscustomobject]@{ Type='Other'; Host=''; Protocol='' }
}

function ConvertFrom-NoisyJson {
    <# .SYNOPSIS Parse JSON that may be surrounded by console noise (first {/[ to last }/]). #>
    [CmdletBinding()]
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    try { return ($Text | ConvertFrom-Json -ErrorAction Stop) } catch {}
    $i = $Text.IndexOfAny([char[]]@('{', '['))
    if ($i -lt 0) { return $null }
    $end = [Math]::Max($Text.LastIndexOf('}'), $Text.LastIndexOf(']'))
    if ($end -le $i) { return $null }
    try { return ($Text.Substring($i, $end - $i + 1) | ConvertFrom-Json -ErrorAction Stop) } catch { return $null }
}

function ConvertTo-PrinterJsonSafe {
    <#
    .SYNOPSIS
        Flatten to primitives before ConvertTo-Json.
    .DESCRIPTION
        ConvertTo-Json reflecting into WMI/CIM/registry ETS note-properties is the
        documented cause of an Atera run that spins forever and then reports
        "Failed". Everything this module returns is already a plain PSCustomObject,
        but the report builder passes user-supplied fragments through here too.
    #>
    [CmdletBinding()]
    param($InputObject, [int]$Depth = 0)

    if ($null -eq $InputObject) { return $null }
    if ($Depth -gt 12) { return "$InputObject" }

    if ($InputObject -is [string] -or $InputObject -is [bool] -or
        $InputObject -is [int] -or $InputObject -is [long] -or
        $InputObject -is [double] -or $InputObject -is [decimal]) { return $InputObject }
    if ($InputObject -is [datetime]) { return $InputObject.ToString('yyyy-MM-ddTHH:mm:ssZ') }
    if ($InputObject -is [enum])     { return "$InputObject" }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $o = [ordered]@{}
        foreach ($k in $InputObject.Keys) { $o["$k"] = ConvertTo-PrinterJsonSafe $InputObject[$k] ($Depth + 1) }
        return [pscustomobject]$o
    }
    if ($InputObject -is [System.Collections.IEnumerable]) {
        $list = [System.Collections.Generic.List[object]]::new()
        foreach ($item in $InputObject) { $list.Add((ConvertTo-PrinterJsonSafe $item ($Depth + 1))) }
        # The unary comma is load-bearing. Returning a bare array lets PowerShell
        # unroll it, so a ONE-element array reaches ConvertTo-Json as a scalar and
        # serialises as {...} instead of [{...}] - which is the single most common
        # case here (one mapped printer, one signed-in user, one print server) and
        # silently breaks any consumer that indexes the property.
        return ,$list.ToArray()
    }
    if ($InputObject -is [pscustomobject]) {
        $o = [ordered]@{}
        foreach ($p in $InputObject.PSObject.Properties) {
            if ($p.MemberType -notin @('NoteProperty', 'Property')) { continue }
            $o[$p.Name] = ConvertTo-PrinterJsonSafe $p.Value ($Depth + 1)
        }
        return [pscustomobject]$o
    }
    return "$InputObject"
}

# endregion


# ---------------------------------------------------------------------------
# region  Context: who is running, who is logged on, which profiles exist
# ---------------------------------------------------------------------------

function Get-InteractiveUsers {
    <#
    .SYNOPSIS
        Every interactive user session on the box, as DOMAIN\user + SessionId.
    .DESCRIPTION
        explorer.exe owner is the PRIMARY source: it handles multiple simultaneous
        sessions (RDP / fast user switching), whereas Win32_ComputerSystem.UserName
        returns only the console user and is $null when the console is locked or
        nobody is signed in. The console value is merged in as a fallback.
    .OUTPUTS
        PSCustomObject[] { User; Sid; SessionId; Source }
    #>
    [CmdletBinding()] param()

    $found = [System.Collections.Generic.List[object]]::new()
    $seen  = @{}

    try {
        $explorers = @(Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" -ErrorAction Stop)
        foreach ($p in $explorers) {
            $o = $null
            try { $o = Invoke-CimMethod -InputObject $p -MethodName GetOwner -ErrorAction Stop } catch { continue }
            if (-not $o.User) { continue }
            $who = if ($o.Domain) { "$($o.Domain)\$($o.User)" } else { [string]$o.User }
            $key = "$who|$($p.SessionId)"
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true
            $sid = $null
            try { $sid = (New-Object Security.Principal.NTAccount $who).Translate([Security.Principal.SecurityIdentifier]).Value } catch {}
            $found.Add([pscustomobject]@{
                User = $who; Sid = $sid; SessionId = [int]$p.SessionId; Source = 'explorer.exe'
            })
        }
    } catch {}

    if ($found.Count -eq 0) {
        try {
            $who = (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).UserName
            if ($who) {
                $sid = $null
                try { $sid = (New-Object Security.Principal.NTAccount $who).Translate([Security.Principal.SecurityIdentifier]).Value } catch {}
                $found.Add([pscustomobject]@{ User = $who; Sid = $sid; SessionId = -1; Source = 'Win32_ComputerSystem' })
            }
        } catch {}
    }

    # .ToArray(), never @($list) - the array subexpression over a populated
    # generic List throws "Argument types do not match" on PS 5.1.
    $found.ToArray()
}

function Get-UserPrinterProfiles {
    <#
    .SYNOPSIS
        User profiles whose registry hive is currently LOADED, i.e. users who are
        signed in (console or RDP, including locked/disconnected).
    .DESCRIPTION
        Read-only by design - this deliberately does NOT reg-load logged-off
        profiles. S-1-5-21-* are local/domain accounts; S-1-12-1-* are Entra (Azure
        AD) accounts, which were previously invisible and made SYSTEM-context runs
        on Entra-joined PCs fall through to SYSTEM's own empty profile.
    .OUTPUTS
        PSCustomObject[] { Sid; User; ProfilePath; IsActive; SessionId }
    #>
    [CmdletBinding()] param()

    $active = @{}
    foreach ($iu in (Get-InteractiveUsers)) { if ($iu.Sid) { $active[$iu.Sid] = $iu } }

    $profilePaths = @{}
    try {
        foreach ($pl in (Get-CimInstance Win32_UserProfile -ErrorAction Stop)) {
            if ($pl.SID) { $profilePaths[[string]$pl.SID] = [string]$pl.LocalPath }
        }
    } catch {}

    $out = [System.Collections.Generic.List[object]]::new()
    # SilentlyContinue, NOT Stop: enumerating HKEY_USERS throws "Requested registry
    # access is not allowed" for a non-admin (one hive is unreadable) even though
    # the readable hives still come back. -ErrorAction Stop threw the whole result
    # away and left the per-user walk - the core of this toolkit - silently empty.
    $hives = @(Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue)

    foreach ($h in $hives) {
        $sid = $h.PSChildName
        if ($sid -like '*_Classes') { continue }
        if ($sid -in $script:ServiceSids) { continue }
        if ($sid -notlike 'S-1-5-21-*' -and $sid -notlike 'S-1-12-1-*') { continue }

        $user = $sid
        try { $user = ([System.Security.Principal.SecurityIdentifier]$sid).Translate([System.Security.Principal.NTAccount]).Value } catch {}

        $out.Add([pscustomobject]@{
            Sid         = $sid
            User        = $user
            ProfilePath = $(if ($profilePaths.ContainsKey($sid)) { $profilePaths[$sid] } else { '' })
            IsActive    = $active.ContainsKey($sid)
            SessionId   = $(if ($active.ContainsKey($sid)) { $active[$sid].SessionId } else { -1 })
        })
    }
    $out.ToArray()
}

function Get-SpoolerState {
    <# .SYNOPSIS Spool service status/start mode. Check this BEFORE any PrintManagement call. #>
    [CmdletBinding()] param()
    try {
        $svc = Get-CimInstance Win32_Service -Filter "Name='Spooler'" -ErrorAction Stop
        return [pscustomobject]@{
            Status    = [string]$svc.State
            StartMode = [string]$svc.StartMode
            Running   = ($svc.State -eq 'Running')
            Account   = [string]$svc.StartName
            ProcessId = [int]$svc.ProcessId
        }
    } catch {
        return [pscustomobject]@{ Status='Unknown'; StartMode='Unknown'; Running=$false; Account=''; ProcessId=0 }
    }
}

function Get-PrinterContext {
    <#
    .SYNOPSIS
        Who we are, what we can do, and what the box is. Drives the GUI banner and
        the ID script header.
    #>
    [CmdletBinding()] param()

    $id      = [Security.Principal.WindowsIdentity]::GetCurrent()
    $isAdmin = $false
    try { $isAdmin = ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } catch {}

    $os = ''; $isServer = $false; $caption = ''
    try {
        $osi = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $caption = [string]$osi.Caption
        $os = "$caption (build $($osi.BuildNumber))"
        $isServer = ([int]$osi.ProductType -ne 1)
    } catch {}

    $domain = ''; $partOfDomain = $false
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $domain = [string]$cs.Domain
        $partOfDomain = [bool]$cs.PartOfDomain
    } catch {}

    $iu = @(Get-InteractiveUsers)

    [pscustomobject]@{
        ToolkitVersion   = $script:ToolkitVersion
        Host             = $env:COMPUTERNAME
        CurrentName      = $id.Name
        CurrentSid       = $id.User.Value
        IsSystem         = ($id.User.Value -eq 'S-1-5-18')
        IsAdmin          = $isAdmin
        OS               = $os
        OSCaption        = $caption
        IsServer         = $isServer
        Domain           = $domain
        PartOfDomain     = $partOfDomain
        Spooler          = (Get-SpoolerState)
        InteractiveUsers = $iu
        PrimaryUser      = $(if ($iu.Count) { $iu[0].User } else { $null })
        PrimaryUserSid   = $(if ($iu.Count) { $iu[0].Sid }  else { $null })
        HavePrintCmdlets = [bool](Get-Command Get-Printer -ErrorAction SilentlyContinue)
        GeneratedUtc     = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }
}

# endregion


# ---------------------------------------------------------------------------
# region  Reachability (raw sockets only - never Test-NetConnection, never ICMP)
# ---------------------------------------------------------------------------

function Test-TcpPortFast {
    <#
    .SYNOPSIS
        Bounded TCP connect. Returns { Open; Ms; Completed; SocketError }.
    .DESCRIPTION
        Test-NetConnection was measured at 8.43 SECONDS for a port this answers in
        17ms, because it runs ICMP + DNS + route lookup regardless. Across 50
        printers that is 7 minutes versus under 2 seconds.

        Completed and SocketError separate the three outcomes that matter, instead of
        guessing from elapsed time:
          Completed=$false                       no answer - dropped, or host down
          SocketError=ConnectionRefused (10061)  nothing listening, or a rejecting firewall
          SocketError=HostNotFound (11001)       the name does not resolve
        Measured on Windows 11: a REFUSED connect is not fast - Windows retries after
        the refusal and reports it at about 2 seconds. A limit below that reports a
        refusal as a timeout, so callers that need the distinction should allow 3s.

        Close() must happen in a finally: EndConnect on a timed-out handle leaks the socket.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][int]$Port,
        [int]$TimeoutMs = 1500
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $client = $null
    $open = $false
    $completed = $false
    $sockErr = ''
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $async  = $client.BeginConnect($HostName, $Port, $null, $null)
        if ($async.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            $completed = $true
            try { $client.EndConnect($async); $open = $client.Connected }
            catch {
                $open = $false
                $se = $_.Exception.InnerException
                if (-not $se) { $se = $_.Exception }
                $sockErr = $(if ($se -is [System.Net.Sockets.SocketException]) { [string]$se.SocketErrorCode } else { 'Error' })
            }
        }
    } catch {
        $open = $false
        $completed = $true
        $se = $_.Exception.InnerException
        if (-not $se) { $se = $_.Exception }
        $sockErr = $(if ($se -is [System.Net.Sockets.SocketException]) { [string]$se.SocketErrorCode } else { 'Error' })
    } finally {
        if ($client) { try { $client.Close() } catch {} }
        $sw.Stop()
    }
    [pscustomobject]@{ Host = $HostName; Port = $Port; Open = $open; Ms = [int]$sw.ElapsedMilliseconds; Completed = $completed; SocketError = $sockErr }
}

function Test-PrinterPortReachable {
    <#
    .SYNOPSIS
        Is the physical printer answering on any print port? Probes 9100/515/631.
    .DESCRIPTION
        Deliberately not ICMP-gated: this machine's printer drops ping while 9100
        is open. "All probe ports closed" is the real offline signal.
    .OUTPUTS
        PSCustomObject { Host; Reachable; OpenPorts; Results; Resolved; Message }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HostName,
        [int[]]$Ports,
        [int]$TimeoutMs = 1500,
        [switch]$NoEscalate,
        [switch]$Extended
    )
    if (-not $Ports -or $Ports.Count -eq 0) {
        $Ports = if ($Extended) { @(Get-PrinterPortCatalogue | ForEach-Object { $_.Port }) } else { $script:PrinterProbePorts }
    }

    $resolved = $false
    try { $null = [System.Net.Dns]::GetHostAddresses($HostName); $resolved = $true } catch {}
    if (-not $resolved) {
        return [pscustomobject]@{
            Host = $HostName; Reachable = $false; OpenPorts = @(); Results = @()
            Resolved = $false; Message = "Name '$HostName' does not resolve."
        }
    }

    # Concurrent: one timeout window for the whole set rather than one per port.
    $arr  = @(Test-TcpPortSetFast -HostName $HostName -Ports $Ports -TimeoutMs $TimeoutMs)
    $open = @($arr | Where-Object { $_.Open } | ForEach-Object { $_.Port })

    # If nothing on the print ports answered, sweep the wider set before giving
    # up. "Dead on 9100 but answering on 443" is a live device with raw printing
    # turned off - a different fault, and a different fix, from a device that is
    # switched off. Skipped when -NoEscalate, and never needed when it answered.
    $escalated = @()
    if ($open.Count -eq 0 -and -not $NoEscalate) {
        $extra = @(Get-PrinterPortCatalogue | Where-Object { -not $_.Standard } | ForEach-Object { $_.Port })
        $extra = @($extra | Where-Object { $_ -notin $Ports })
        if ($extra.Count) {
            $escalated = @(Test-TcpPortSetFast -HostName $HostName -Ports $extra -TimeoutMs $TimeoutMs)
            $arr = $arr + $escalated
        }
    }
    $openAny = @($arr | Where-Object { $_.Open } | ForEach-Object { $_.Port })

    $msg = if ($open.Count) {
        "Open: $($open -join ', ')"
    } elseif ($openAny.Count) {
        $names = @(foreach ($o in $openAny) { $c = @($script:PrinterPortCatalogue | Where-Object { $_.Port -eq $o })[0]; if ($c) { "$o ($($c.Name))" } else { "$o" } })
        "No print port answered, but the device IS alive - responding on $($names -join ', '). Printing is likely disabled or blocked on it rather than the device being off."
    } else {
        "No response on $($Ports -join '/')$(if ($escalated.Count) { ' or any other known printer port' } else { '' })."
    }

    [pscustomobject]@{
        Host       = $HostName
        Reachable  = ($open.Count -gt 0)
        AliveOnly  = ($open.Count -eq 0 -and $openAny.Count -gt 0)
        OpenPorts  = $open
        OpenAny    = $openAny
        Results    = $arr
        Resolved   = $true
        Message    = $msg
    }
}

function Test-PrintServerReachable {
    <#
    .SYNOPSIS
        Is the print server reachable? TCP 445/135 plus a bounded print$ check.
    .DESCRIPTION
        NEVER tests '\\server\printer' with Test-Path - print shares are not
        filesystem objects and that returns False for a share that works perfectly.
        print$ (the driver share) IS a filesystem share and is a valid probe.

        The print$ check is bounded separately. The TCP timeouts do not bound a
        filesystem call, and an unbounded Test-Path against a host that swallows the
        request froze the console.

        From SYSTEM, SMB authenticates as DOMAIN\MACHINE$, not the user, so a print$
        denial is not proof the user cannot print. On a workgroup machine SYSTEM has
        no network identity at all. Reachable says the host answered - nothing about
        whether the spooler's own RPC path works.
    .OUTPUTS
        PSCustomObject { Server; Reachable; Smb445; Rpc135; DriverShare; DriverShareState; Message }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server,
        [int]$TimeoutMs = 1500,
        [int]$PrintShareTimeoutSec = 8
    )
    $s = $Server.TrimStart('\').Trim()

    $t445 = Test-TcpPortFast -HostName $s -Port 445 -TimeoutMs $TimeoutMs
    $t135 = Test-TcpPortFast -HostName $s -Port 135 -TimeoutMs $TimeoutMs

    $driverShare = $false
    $driverShareState = 'Skipped'
    if ($t445.Open) {
        $pr = Invoke-WithTimeout -TimeoutSec $PrintShareTimeoutSec -Parameters @{ P = "\\$s\print$" } -ScriptBlock {
            param([string]$P)
            try { [bool](Test-Path -LiteralPath $P -ErrorAction Stop) } catch { $false }
        }
        if ($null -eq $pr) {
            $driverShareState = 'Timeout'
        } else {
            $driverShare = [bool](@($pr)[0])
            $driverShareState = $(if ($driverShare) { 'Ok' } else { 'Failed' })
        }
    }

    $reachable = ($t445.Open -or $t135.Open)
    $msg = if (-not $reachable) { "No SMB (445) or RPC (135) response from $s." }
           elseif ($driverShareState -eq 'Timeout') { "Reachable; print$ did not answer within $PrintShareTimeoutSec s." }
           elseif ($driverShare) { 'Reachable; print$ readable.' }
           else { 'Reachable; print$ not readable from this context (normal for SYSTEM on some servers).' }

    [pscustomobject]@{
        Server           = $s
        Reachable        = $reachable
        Smb445           = $t445.Open
        Rpc135           = $t135.Open
        DriverShare      = $driverShare
        DriverShareState = $driverShareState
        Message          = $msg
    }
}

# endregion


# ---------------------------------------------------------------------------
# region  Local printer enumeration (cmdlets -> WMI -> spooler registry)
# ---------------------------------------------------------------------------

function Get-PrintersViaRegistry {
    <#
    .SYNOPSIS
        Read printers straight out of the local spooler registry. No cmdlets, no
        RPC, no spooler service required.
    .DESCRIPTION
        Third-tier fallback, and the only thing that works when the Spooler is
        stopped. NOTE: this path can only ever see LOCAL queues - mapped printers
        are not present under HKLM at all (verified). Use Get-UserPrinterMap for those.
    .NOTES
        Carried over from Export_Print_Server_Printers_To_JSON.ps1.
    #>
    [CmdletBinding()]
    param([switch]$SharedOnly)

    $printRoot = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers'
    if (-not (Test-Path $printRoot)) { return @() }

    $portMap  = @{}
    $portRoot = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\Standard TCP/IP Port\Ports'
    if (Test-Path $portRoot) {
        foreach ($pp in (Get-ChildItem $portRoot -ErrorAction SilentlyContinue)) {
            $pi = Get-ItemProperty $pp.PSPath -ErrorAction SilentlyContinue
            if ($pi.IPAddress) { $portMap[$pp.PSChildName] = [string]$pi.IPAddress }
        }
    }

    $items = foreach ($pk in (Get-ChildItem $printRoot -ErrorAction SilentlyContinue)) {
        $p = Get-ItemProperty $pk.PSPath -ErrorAction SilentlyContinue
        $attr   = [int]$p.Attributes
        $shared = [bool]($attr -band 0x8)      # PRINTER_ATTRIBUTE_SHARED
        if ($SharedOnly -and -not $shared) { continue }

        $port = [string]$p.Port
        $addr = ''
        if ($port -and $portMap.ContainsKey($port)) { $addr = $portMap[$port] }
        if (-not $addr -and $port) {
            $m = [regex]::Match($port, '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}')
            if ($m.Success) { $addr = $m.Value }
        }

        $name   = if ($p.Name) { [string]$p.Name } else { [string]$pk.PSChildName }
        $driver = [string]$p.'Printer Driver'
        $ci     = Get-PrinterConnectionInfo -PortName $port -TcpPort $null
        if ($addr -and -not $ci.Host) { $ci = [pscustomobject]@{ Type='Network'; Host=$addr; Protocol=$ci.Protocol } }

        [pscustomobject]@{
            Name          = $name
            Type          = 'Local'
            ServerName    = ''
            ShareName     = [string]$p.'Share Name'
            Shared        = $shared
            Published     = [bool]($attr -band 0x2000)
            WorkOffline   = [bool]($attr -band 0x400)
            DriverName    = $driver
            DriverVersion = ''
            DriverMajor   = $null
            PackageAware  = $null
            PortName      = $port
            PortType      = $ci.Type
            HostAddress   = $ci.Host
            Protocol      = $ci.Protocol
            SNMPEnabled   = $null
            SNMPCommunity = ''
            SNMPIndex     = $null
            LprQueue      = ''
            IsVirtual     = (Test-IsVirtualPrinter -Name $name -DriverName $driver)
            IsIppClass    = (Test-IsIppClassDriver -DriverName $driver)
            IsDefault     = $false
            PrinterStatus = ''
            StateText     = ''
            JobCount      = $null
            Location      = [string]$p.Location
            Comment       = [string]$p.Description
            Attributes    = $attr
            Source        = 'registry'
        }
    }
    @($items | Where-Object { $null -ne $_ })
}

function Get-LocalPrinterInventory {
    <#
    .SYNOPSIS
        Every printer queue visible in THIS process's spooler view, normalised.
    .DESCRIPTION
        Tier 1 PrintManagement cmdlets (richest: DriverVersion, MajorVersion,
        IsPackageAware, Type=Local|Connection), tier 2 WMI, tier 3 spooler registry.

        IMPORTANT: run as SYSTEM this returns LOCAL queues only. A user's mapped
        printers live solely in HKU\<SID>\Printers\Connections - see Get-UserPrinterMap.

        Every risky call is bounded (see Invoke-WithTimeout): a dead network queue
        blocks Win32_Printer for 30-45s per queue.
    .OUTPUTS
        PSCustomObject[] - see the field list in the registry fallback above.
    #>
    [CmdletBinding()]
    param(
        [switch]$IncludeVirtual,
        [int]$TimeoutSec = 25
    )

    $spooler = Get-SpoolerState
    $result  = @()

    # --- Tier 1: PrintManagement cmdlets -----------------------------------
    # Skipped entirely when the Spooler is stopped: these cmdlets THROW rather
    # than return empty, which would burn the timeout for nothing.
    if ($spooler.Running -and (Get-Command Get-Printer -ErrorAction SilentlyContinue)) {
        $raw = Invoke-WithTimeout -TimeoutSec $TimeoutSec -ScriptBlock {
            $printers = @(Get-Printer -ErrorAction SilentlyContinue)
            $ports    = @(Get-PrinterPort -ErrorAction SilentlyContinue)
            # Select the driver fields we need immediately - never pass whole
            # driver objects around, and never touch Get-PrintConfiguration here
            # (it embeds PrintCapabilitiesXML + PrintTicketXML: 80KB for 6 queues).
            $drivers  = @(Get-PrinterDriver -ErrorAction SilentlyContinue |
                          Select-Object Name, DriverVersion, MajorVersion, IsPackageAware, Manufacturer)
            [pscustomobject]@{ Printers = $printers; Ports = $ports; Drivers = $drivers }
        }

        if ($raw -and $raw.Printers) {
            $portByName   = @{}
            foreach ($p in @($raw.Ports))   { if ($p.Name) { $portByName[[string]$p.Name] = $p } }
            $driverByName = @{}
            foreach ($d in @($raw.Drivers)) { if ($d.Name) { $driverByName[[string]$d.Name] = $d } }

            $defaults = @{}
            try {
                foreach ($w in @(Get-CimInstance Win32_Printer -Property Name,Default,WorkOffline,PrinterStatus,PrinterState -ErrorAction SilentlyContinue)) {
                    $defaults[[string]$w.Name] = $w
                }
            } catch {}

            $list = [System.Collections.Generic.List[object]]::new()
            foreach ($p in @($raw.Printers)) {
                $portName = [string]$p.PortName
                $port     = $(if ($portByName.ContainsKey($portName)) { $portByName[$portName] } else { $null })
                $drv      = $(if ($driverByName.ContainsKey([string]$p.DriverName)) { $driverByName[[string]$p.DriverName] } else { $null })
                $wmi      = $(if ($defaults.ContainsKey([string]$p.Name)) { $defaults[[string]$p.Name] } else { $null })
                $ci       = Get-PrinterConnectionInfo -PortName $portName -TcpPort $port

                # A connection printer's PortName is the SERVER's port name, which
                # leaks the printer's real IP to the client - useful, take it.
                if (-not $ci.Host -and $portName -match '^\d{1,3}(\.\d{1,3}){3}$') {
                    $ci = [pscustomobject]@{ Type='Network'; Host=$portName; Protocol=$ci.Protocol }
                }

                $server = ''
                if ([string]$p.Type -eq 'Connection' -and [string]$p.Name -match '^\\\\([^\\]+)\\') { $server = $matches[1] }

                $list.Add([pscustomobject]@{
                    Name          = [string]$p.Name
                    Type          = [string]$p.Type
                    ServerName    = $server
                    ShareName     = [string]$p.ShareName
                    Shared        = [bool]$p.Shared
                    Published     = [bool]$p.Published
                    WorkOffline   = $(if ($wmi) { [bool]$wmi.WorkOffline } else { $false })
                    DriverName    = [string]$p.DriverName
                    DriverVersion = $(if ($drv) { [string]$drv.DriverVersion } else { '' })
                    DriverMajor   = $(if ($drv) { $drv.MajorVersion } else { $null })
                    PackageAware  = $(if ($drv) { $drv.IsPackageAware } else { $null })
                    PortName      = $portName
                    PortType      = $ci.Type
                    HostAddress   = $ci.Host
                    Protocol      = $ci.Protocol
                    SNMPEnabled   = $(if ($port) { $port.SNMPEnabled } else { $null })
                    SNMPCommunity = $(if ($port) { [string]$port.SNMPCommunity } else { '' })
                    SNMPIndex     = $(if ($port) { $port.SNMPIndex } else { $null })
                    LprQueue      = $(if ($port) { [string]$port.LprQueueName } else { '' })
                    IsVirtual     = (Test-IsVirtualPrinter -Name ([string]$p.Name) -DriverName ([string]$p.DriverName))
                    IsIppClass    = (Test-IsIppClassDriver -DriverName ([string]$p.DriverName))
                    IsDefault     = $(if ($wmi) { [bool]$wmi.Default } else { $false })
                    PrinterStatus = [string]$p.PrinterStatus
                    StateText     = $(if ($wmi) { Get-PrinterStateText $wmi.PrinterState } else { '' })
                    JobCount      = $p.JobCount
                    Location      = [string]$p.Location
                    Comment       = [string]$p.Comment
                    Attributes    = $null
                    Source        = 'cmdlet'
                })
            }
            $result = $list.ToArray()
        }
    }

    # --- Tier 2: WMI --------------------------------------------------------
    if ($result.Count -eq 0 -and $spooler.Running) {
        $raw = Invoke-WithTimeout -TimeoutSec $TimeoutSec -ScriptBlock {
            $printers = @(Get-WmiObject Win32_Printer -ErrorAction SilentlyContinue |
                Select-Object Name, Default, PortName, DriverName, Network, Shared, ShareName, Local,
                              ServerName, PrinterStatus, PrinterState, WorkOffline, Published,
                              Location, Comment, Attributes, JobCountSinceLastReset)
            $ports = @(Get-WmiObject Win32_TCPIPPrinterPort -ErrorAction SilentlyContinue |
                Select-Object Name, HostAddress, PortNumber, Protocol, SNMPEnabled, SNMPCommunity, SNMPDevIndex, Queue)
            [pscustomobject]@{ Printers = $printers; Ports = $ports }
        }

        if ($raw -and $raw.Printers) {
            $portByName = @{}
            foreach ($p in @($raw.Ports)) { if ($p.Name) { $portByName[[string]$p.Name] = $p } }

            $list = [System.Collections.Generic.List[object]]::new()
            foreach ($p in @($raw.Printers)) {
                $portName = [string]$p.PortName
                $port     = $(if ($portByName.ContainsKey($portName)) { $portByName[$portName] } else { $null })
                $ci       = Get-PrinterConnectionInfo -PortName $portName -TcpPort $port
                $attr     = [int]$p.Attributes
                $server   = ([string]$p.ServerName).TrimStart('\')

                $list.Add([pscustomobject]@{
                    Name          = [string]$p.Name
                    Type          = $(if ($p.Network) { 'Connection' } else { 'Local' })
                    ServerName    = $server
                    ShareName     = [string]$p.ShareName
                    Shared        = [bool]$p.Shared
                    Published     = [bool]$p.Published
                    WorkOffline   = [bool]$p.WorkOffline
                    DriverName    = [string]$p.DriverName
                    DriverVersion = ''
                    DriverMajor   = $null
                    PackageAware  = $null
                    PortName      = $portName
                    PortType      = $ci.Type
                    HostAddress   = $ci.Host
                    Protocol      = $ci.Protocol
                    SNMPEnabled   = $(if ($port) { $port.SNMPEnabled } else { $null })
                    SNMPCommunity = $(if ($port) { [string]$port.SNMPCommunity } else { '' })
                    SNMPIndex     = $(if ($port) { $port.SNMPDevIndex } else { $null })
                    LprQueue      = $(if ($port) { [string]$port.Queue } else { '' })
                    IsVirtual     = (Test-IsVirtualPrinter -Name ([string]$p.Name) -DriverName ([string]$p.DriverName))
                    IsIppClass    = (Test-IsIppClassDriver -DriverName ([string]$p.DriverName))
                    IsDefault     = [bool]$p.Default
                    PrinterStatus = [string]$p.PrinterStatus
                    StateText     = (Get-PrinterStateText $p.PrinterState)
                    JobCount      = $null
                    Location      = [string]$p.Location
                    Comment       = [string]$p.Comment
                    Attributes    = $attr
                    Source        = 'wmi'
                })
            }
            $result = $list.ToArray()
        }
    }

    # --- Tier 3: spooler registry ------------------------------------------
    if ($result.Count -eq 0) { $result = @(Get-PrintersViaRegistry) }

    if (-not $IncludeVirtual) { $result = @($result | Where-Object { -not $_.IsVirtual }) }
    @($result)
}

# endregion


# ---------------------------------------------------------------------------
# region  Per-user printer state (the part Get-Printer cannot see from SYSTEM)
# ---------------------------------------------------------------------------

function Get-UserPrinterMap {
    <#
    .SYNOPSIS
        One user's mapped printers, per-user port table, and default printer, read
        from their loaded registry hive.
    .DESCRIPTION
        HKU\<SID>\Printers\Connections key names are the UNC path with every '\'
        replaced by ',' - so '\\server\printer' is stored as ',,server,printer'.
        Round-trips with .Replace(',','\') because the leading ',,' becomes '\\'.
        Share names cannot contain commas, so this is safe.

        The default printer lives at
        HKU\<SID>\Software\Microsoft\Windows NT\CurrentVersion\Windows -> Device
        formatted '<printer name>,winspool,<port>'. Split on the FIRST comma only.

        LegacyDefaultPrinterMode = 1 means "Let Windows manage my default printer"
        is OFF. 0 or absent means Windows will silently overwrite Device on the
        next print - which is why setting a default without also setting this
        value does not stick.

        The Ne0x: port token is per-user and inconsistent (observed Ne04: in
        Devices vs Ne02: in Device for the SAME printer). Treat it as opaque.

        Uses the .NET registry API rather than the provider: the provider costs
        ~480ms per Test-Path on some hives, and this walks every SID.
    .OUTPUTS
        PSCustomObject { Sid; User; Connections; Devices; PrinterPorts; DefaultDevice;
                         DefaultPrinter; DefaultPort; LegacyDefaultPrinterMode }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Sid,
        [string]$User
    )

    if (-not $User) {
        $User = $Sid
        try { $User = ([System.Security.Principal.SecurityIdentifier]$Sid).Translate([System.Security.Principal.NTAccount]).Value } catch {}
    }

    $connections = [System.Collections.Generic.List[object]]::new()
    $devices     = [System.Collections.Generic.List[object]]::new()
    $ports       = [System.Collections.Generic.List[object]]::new()
    $defaultDev  = ''
    $legacy      = $null

    $root = $null
    try { $root = [Microsoft.Win32.Registry]::Users.OpenSubKey($Sid) } catch {}
    if (-not $root) {
        return [pscustomobject]@{
            Sid = $Sid; User = $User; HiveLoaded = $false
            Connections = @(); Devices = @(); PrinterPorts = @()
            DefaultDevice = ''; DefaultPrinter = ''; DefaultPort = ''
            LegacyDefaultPrinterMode = $null
        }
    }

    try {
        # --- Printers\Connections -> the mapped printers ---
        $ck = $null
        try { $ck = $root.OpenSubKey('Printers\Connections') } catch {}
        if ($ck) {
            try {
                foreach ($name in $ck.GetSubKeyNames()) {
                    $unc = $name.Replace(',', '\')
                    $server = ''; $printer = ''
                    if ($unc -match '^\\\\([^\\]+)\\(.+)$') { $server = $matches[1]; $printer = $matches[2] }
                    $sub = $null
                    $guid = ''; $provider = ''; $localConn = $null
                    try {
                        $sub = $ck.OpenSubKey($name)
                        if ($sub) {
                            $guid      = [string]$sub.GetValue('GuidPrinter')
                            $provider  = [string]$sub.GetValue('Provider')
                            $localConn = $sub.GetValue('LocalConnection')
                        }
                    } catch {} finally { if ($sub) { $sub.Close() } }

                    $connections.Add([pscustomobject]@{
                        Unc = $unc; Server = $server; Printer = $printer
                        GuidPrinter = $guid; Provider = $provider; LocalConnection = $localConn
                    })
                }
            } catch {} finally { $ck.Close() }
        }

        # --- Devices / PrinterPorts -> per-user port table (includes local queues) ---
        foreach ($pair in @(
            @{ Path = 'Software\Microsoft\Windows NT\CurrentVersion\Devices';      Bag = $devices },
            @{ Path = 'Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts'; Bag = $ports })) {
            $k = $null
            try { $k = $root.OpenSubKey($pair.Path) } catch {}
            if (-not $k) { continue }
            try {
                foreach ($vn in $k.GetValueNames()) {
                    $val = [string]$k.GetValue($vn)
                    # 'winspool,Ne04:' (Devices) or 'winspool,Ne04:,15,45' (PrinterPorts)
                    $parts = $val -split ','
                    $port  = $(if ($parts.Count -ge 2) { $parts[1] } else { '' })
                    $pair.Bag.Add([pscustomobject]@{ Name = $vn; Port = $port; Raw = $val })
                }
            } catch {} finally { $k.Close() }
        }

        # --- Windows -> Device + LegacyDefaultPrinterMode ---
        $wk = $null
        try { $wk = $root.OpenSubKey('Software\Microsoft\Windows NT\CurrentVersion\Windows') } catch {}
        if ($wk) {
            try {
                $defaultDev = [string]$wk.GetValue('Device')
                $legacy     = $wk.GetValue('LegacyDefaultPrinterMode')
            } catch {} finally { $wk.Close() }
        }
    } finally {
        $root.Close()
    }

    # Device is '<printer>,<provider-or-driver>,<port>'. Split on the FIRST comma
    # for the name, and take the LAST field for the port.
    #
    # The middle field is NOT reliably the literal 'winspool' that most references
    # quote. Measured on Windows 11 26200, WScript.Network.SetDefaultPrinter
    # rewrites the value with the DRIVER NAME in the middle:
    #   before  ...\OFFICE-MFP (HP Color LaserJet MFP M480),winspool,Ne02:
    #   after   ...\OFFICE-MFP (HP Color LaserJet MFP M480),HP Color LaserJet MFP M480 PCL-6 (V4),192.0.2.47
    # Indexing field [1] therefore returns the driver name on some machines and
    # the port on others. The port is always last, and a driver name containing a
    # comma cannot break that.
    $defPrinter = ''; $defPort = ''
    if ($defaultDev) {
        $idx = $defaultDev.IndexOf(',')
        if ($idx -gt 0) {
            $defPrinter = $defaultDev.Substring(0, $idx)
            $tail = $defaultDev.Substring($idx + 1) -split ','
            if ($tail.Count -ge 1) { $defPort = $tail[$tail.Count - 1] }
        } else {
            $defPrinter = $defaultDev
        }
    }

    [pscustomobject]@{
        Sid                      = $Sid
        User                     = $User
        HiveLoaded               = $true
        Connections              = $connections.ToArray()
        Devices                  = $devices.ToArray()
        PrinterPorts             = $ports.ToArray()
        DefaultDevice            = $defaultDev
        DefaultPrinter           = $defPrinter
        DefaultPort              = $defPort
        LegacyDefaultPrinterMode = $legacy
    }
}

function Get-AllUserPrinterMaps {
    <# .SYNOPSIS Get-UserPrinterMap for every loaded profile hive. #>
    [CmdletBinding()] param()
    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($p in (Get-UserPrinterProfiles)) {
        $m = Get-UserPrinterMap -Sid $p.Sid -User $p.User
        Add-Member -InputObject $m -NotePropertyName 'IsActive'    -NotePropertyValue $p.IsActive    -Force
        Add-Member -InputObject $m -NotePropertyName 'ProfilePath' -NotePropertyValue $p.ProfilePath -Force
        $out.Add($m)
    }
    $out.ToArray()
}

# endregion


# ---------------------------------------------------------------------------
# region  Print server enumeration + discovery
# ---------------------------------------------------------------------------

function Get-PrintServerInventory {
    <#
    .SYNOPSIS
        Shared print queues on a print server.
    .DESCRIPTION
        Get-Printer -ComputerName is the primary path and works as a PLAIN DOMAIN
        USER - it only needs the share's default Everyone:Print right plus RPC.
        Remote Get-WmiObject needs LOCAL ADMIN on the print server (measured:
        "Access is denied" for a normal domain user) so it is not attempted
        remotely. net view is the unprivileged fallback and cleanly tags print
        shares by type. When the target resolves to this machine, the local
        spooler registry is used - no RPC required at all.
    .OUTPUTS
        PSCustomObject { Server; Method; Reachable; Printers[]; Message }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server,
        [switch]$IncludeUnshared,
        [int]$TimeoutSec = 25,
        # Without these, enumeration runs as whatever identity the console holds,
        # which is why the Credentials dialog could see a host the Shares tab
        # could not: only the dialog was authenticating.
        [string]$User,
        [string]$Password
    )

    $s     = $Server.TrimStart('\').Trim()
    $local = Test-IsLocalHost $s
    $printers = @()
    $method = ''
    $message = ''
    $authNote = ''

    # Authenticate FIRST when credentials were supplied, so every probe below
    # runs as that account rather than as the console's own identity.
    if ($User -and -not $local) {
        $conn = Connect-PrintHostSession -Server $s -User $User -Password $Password
        $authNote = (@($conn.Notes) -join ' ')
    }

    if ($local) {
        $printers = @(Get-PrintersViaRegistry -SharedOnly:(-not $IncludeUnshared))
        $method   = 'local spooler registry'
        return [pscustomobject]@{
            Server = $s; Method = $method; Reachable = $true
            Printers = $printers; Message = 'Target is this machine - read locally.'
        }
    }

    $reach = Test-PrintServerReachable -Server $s
    if (-not $reach.Reachable) {
        return [pscustomobject]@{
            Server = $s; Method = 'none'; Reachable = $false
            Printers = @(); Message = ("$($reach.Message)$(if ($authNote) { "  [$authNote]" })").Trim()
        }
    }

    if (Get-Command Get-Printer -ErrorAction SilentlyContinue) {
        $raw = Invoke-WithTimeout -TimeoutSec $TimeoutSec -Parameters @{ Srv = $s } -ScriptBlock {
            param($Srv)
            $p = @(Get-Printer -ComputerName $Srv -ErrorAction SilentlyContinue)
            $q = @(Get-PrinterPort -ComputerName $Srv -ErrorAction SilentlyContinue)
            [pscustomobject]@{ Printers = $p; Ports = $q }
        }
        if ($raw -and $raw.Printers -and @($raw.Printers).Count) {
            $portByName = @{}
            foreach ($p in @($raw.Ports)) { if ($p.Name) { $portByName[[string]$p.Name] = $p } }

            $list = [System.Collections.Generic.List[object]]::new()
            foreach ($p in @($raw.Printers)) {
                if (-not $IncludeUnshared -and -not $p.Shared) { continue }
                $portName = [string]$p.PortName
                $port = $(if ($portByName.ContainsKey($portName)) { $portByName[$portName] } else { $null })
                $ci   = Get-PrinterConnectionInfo -PortName $portName -TcpPort $port
                if (-not $ci.Host -and $portName -match '^\d{1,3}(\.\d{1,3}){3}$') {
                    $ci = [pscustomobject]@{ Type='Network'; Host=$portName; Protocol=$ci.Protocol }
                }
                $list.Add([pscustomobject]@{
                    Name        = [string]$p.Name
                    ShareName   = [string]$p.ShareName
                    Shared      = [bool]$p.Shared
                    Published   = [bool]$p.Published
                    DriverName  = [string]$p.DriverName
                    PortName    = $portName
                    HostAddress = $ci.Host
                    Protocol    = $ci.Protocol
                    Location    = [string]$p.Location
                    Comment     = [string]$p.Comment
                    # Only an actually-shared queue has a mappable UNC. Emitting
                    # "\\server\" for an unshared queue produces a path that can
                    # never be connected to and looks like a real one in a grid.
                    Unc         = $(if ($p.ShareName) { "\\$s\$([string]$p.ShareName)" } else { '' })
                    Source      = 'cmdlet'
                })
            }
            return [pscustomobject]@{
                Server = $s; Method = 'Get-Printer -ComputerName'; Reachable = $true
                Printers = $list.ToArray(); Message = ("$($reach.Message)$(if ($authNote) { "  [$authNote]" })").Trim()
            }
        }
        $message = 'Get-Printer -ComputerName returned nothing; falling back to net view.'
    }

    # net view: unprivileged, names + comments only.
    $out = Invoke-WithTimeout -TimeoutSec 20 -Parameters @{ Srv = $s } -ScriptBlock {
        param($Srv)
        & net view "\\$Srv" 2>&1 | Out-String
    }
    $list = [System.Collections.Generic.List[object]]::new()
    if ($out) {
        foreach ($line in ("$out" -split "`r?`n")) {
            if ($line -notmatch '\bPrint\b') { continue }
            $share = ($line.Trim() -split '\s{2,}')[0]
            if (-not $share -or $share -match '^-+$') { continue }
            $list.Add([pscustomobject]@{
                Name = $share; ShareName = $share; Shared = $true; Published = $false
                DriverName = ''; PortName = ''; HostAddress = ''; Protocol = ''
                Location = ''; Comment = ''; Unc = "\\$s\$share"; Source = 'net view'
            })
        }
    }

    [pscustomobject]@{
        Server = $s
        Method = 'net view (names only)'
        Reachable = $true
        Printers = $list.ToArray()
        Message = ("$message Only share names could be captured - run the exporter ON the print server for driver/port detail.$(if ($authNote) { "  [$authNote]" })").Trim()
    }
}

function Get-PrintServerCandidates {
    <#
    .SYNOPSIS
        Discover which print server(s) this machine actually uses.
    .DESCRIPTION
        Ranked, cheapest and most reliable first:
          1. Every loaded user hive's Printers\Connections - in a domain where
             printers are NOT published to AD (the common case; measured 0
             printQueue objects here) this is the only thing that finds the server.
          2. Win32_Printer.ServerName for connection printers in this process view.
          3. AD (objectClass=printQueue) via DirectorySearcher.

        DirectorySearcher, not Get-ADObject: RSAT/the ActiveDirectory module is not
        installed on a typical endpoint, so the AD cmdlets are unavailable. ADSI
        needs nothing. PageSize must be set or the query silently caps at 1000.
    .OUTPUTS
        PSCustomObject { Server; Sources[]; UserCount; AdPrinters[] }
    #>
    [CmdletBinding()]
    param([switch]$SkipAd, [int]$TimeoutSec = 20)

    $servers = @{}
    function Add-Candidate([string]$Name, [string]$Source) {
        if ([string]::IsNullOrWhiteSpace($Name)) { return }
        $n = $Name.TrimStart('\').Trim()
        if (-not $n) { return }
        $key = $n.ToLowerInvariant()
        if (-not $servers.ContainsKey($key)) {
            $servers[$key] = [pscustomobject]@{
                Server = $n; Sources = [System.Collections.Generic.List[string]]::new(); UserCount = 0
            }
        }
        if ($servers[$key].Sources -notcontains $Source) { $servers[$key].Sources.Add($Source) }
    }

    # 1. user hives
    foreach ($m in (Get-AllUserPrinterMaps)) {
        $hit = @{}
        foreach ($c in @($m.Connections)) {
            if (-not $c.Server) { continue }
            Add-Candidate $c.Server 'user connections'
            $hit[$c.Server.ToLowerInvariant()] = $true
        }
        foreach ($k in $hit.Keys) { if ($servers.ContainsKey($k)) { $servers[$k].UserCount++ } }
    }

    # 2. this process's connection printers
    try {
        foreach ($p in (Get-LocalPrinterInventory -IncludeVirtual)) {
            if ($p.ServerName) { Add-Candidate $p.ServerName 'local spooler' }
        }
    } catch {}

    # 3. AD printQueue objects
    $adPrinters = @()
    if (-not $SkipAd) {
        $adPrinters = @(Invoke-WithTimeout -TimeoutSec $TimeoutSec -ScriptBlock {
            try {
                $root = [ADSI]'LDAP://RootDSE'
                $nc = $root.defaultNamingContext
                if (-not $nc) { return @() }
                $searcher = New-Object System.DirectoryServices.DirectorySearcher
                $searcher.SearchRoot = [ADSI]"LDAP://$nc"
                $searcher.Filter     = '(objectClass=printQueue)'
                $searcher.PageSize   = 100
                foreach ($pn in @('printerName','serverName','shortServerName','uNCName','driverName','portName','location','printShareName')) {
                    [void]$searcher.PropertiesToLoad.Add($pn)
                }
                $res = $searcher.FindAll()
                $out = @()
                foreach ($r in $res) {
                    $get = { param($n) if ($r.Properties[$n].Count) { [string]$r.Properties[$n][0] } else { '' } }
                    $out += [pscustomobject]@{
                        PrinterName = (& $get 'printername')
                        ServerName  = (& $get 'servername')
                        ShortServer = (& $get 'shortservername')
                        Unc         = (& $get 'uncname')
                        DriverName  = (& $get 'drivername')
                        PortName    = (& $get 'portname')
                        Location    = (& $get 'location')
                        ShareName   = (& $get 'printsharename')
                    }
                }
                $res.Dispose()
                $out
            } catch { @() }
        })
        foreach ($ad in $adPrinters) {
            $srv = $ad.ShortServer
            if (-not $srv -and $ad.ServerName) { $srv = ($ad.ServerName -split '\.')[0] }
            if (-not $srv -and $ad.Unc -match '^\\\\([^\\]+)\\') { $srv = $matches[1] }
            Add-Candidate $srv 'AD printQueue'
        }
    }

    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($v in $servers.Values) {
        $list.Add([pscustomobject]@{
            Server     = $v.Server
            Sources    = $v.Sources.ToArray()
            UserCount  = $v.UserCount
            AdPrinters = @($adPrinters | Where-Object {
                $_.ShortServer -eq $v.Server -or
                ($_.ServerName -and ($_.ServerName -split '\.')[0] -eq $v.Server)
            })
        })
    }
    # Most-used first: the server most users map from is the one that matters.
    @($list.ToArray() | Sort-Object -Property @{ E = 'UserCount'; Descending = $true }, 'Server')
}

# endregion


# ---------------------------------------------------------------------------
# region  Spooler health, jobs, events, Point and Print policy
# ---------------------------------------------------------------------------

function Get-SpoolerHealth {
    <#
    .SYNOPSIS
        Spooler service state plus the spool directory's orphaned-file picture.
    .DESCRIPTION
        DefaultSpoolDirectory is READ, not hardcoded - it is configurable globally
        and overridable per printer. Reading the spool folder silently returns
        nothing for a non-admin; from SYSTEM it works.

        Orphaned .SPL/.SHD pairs and their age are the stuck-job metric that
        actually works, because PrinterStatus/WorkOffline/DetectedErrorState all
        reported healthy on a queue holding a 7-month-old job.
    #>
    [CmdletBinding()] param()

    $state = Get-SpoolerState
    $spoolDir = "$env:SystemRoot\System32\spool\PRINTERS"
    try {
        $v = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers' -Name 'DefaultSpoolDirectory' -ErrorAction Stop).DefaultSpoolDirectory
        if ($v) { $spoolDir = [string]$v }
    } catch {}

    $files = @()
    $readable = $true
    try {
        $files = @(Get-ChildItem -LiteralPath $spoolDir -Force -ErrorAction Stop |
                   Where-Object { $_.Extension -in @('.SPL', '.SHD', '.spl', '.shd') })
    } catch { $readable = $false }

    $oldest = $null
    $oldestAgeHours = $null
    if ($files.Count) {
        $oldest = ($files | Sort-Object LastWriteTime | Select-Object -First 1).LastWriteTime
        $oldestAgeHours = [Math]::Round(((Get-Date) - $oldest).TotalHours, 1)
    }

    [pscustomobject]@{
        Status         = $state.Status
        StartMode      = $state.StartMode
        Running        = $state.Running
        Account        = $state.Account
        ProcessId      = $state.ProcessId
        SpoolDirectory = $spoolDir
        SpoolReadable  = $readable
        SpoolFileCount = $files.Count
        OldestSpoolFile      = $(if ($oldest) { $oldest.ToString('yyyy-MM-dd HH:mm') } else { '' })
        OldestSpoolAgeHours  = $oldestAgeHours
    }
}

function Get-PrinterJobSnapshot {
    <#
    .SYNOPSIS
        Current print jobs across all queues, with age.
    .DESCRIPTION
        For a CONNECTION printer, Get-PrintJob queries the SERVER queue and
        returns ALL USERS' jobs, and PrinterName comes back as the server-side
        name (not the \\server\share name that was passed in) - so the returned
        PrinterName is recorded separately from the queue we asked about.

        A job with JobStatus=Complete and PagesPrinted=0 that is still queued days
        later is the classic stuck-job signature.
    #>
    [CmdletBinding()]
    param([string[]]$PrinterName, [int]$TimeoutSec = 20)

    if (-not (Get-Command Get-PrintJob -ErrorAction SilentlyContinue)) { return @() }
    if (-not (Get-SpoolerState).Running) { return @() }

    if (-not $PrinterName -or $PrinterName.Count -eq 0) {
        $PrinterName = @((Get-LocalPrinterInventory -IncludeVirtual) | ForEach-Object { $_.Name })
    }
    if (-not $PrinterName -or $PrinterName.Count -eq 0) { return @() }

    $raw = Invoke-WithTimeout -TimeoutSec $TimeoutSec -Parameters @{ Names = [string[]]$PrinterName } -ScriptBlock {
        param([string[]]$Names)
        # Paper size is only on Win32_PrintJob. One query for every queue, keyed by
        # "<printer>|<job id>"; a job it does not list keeps a blank paper size.
        $paper = @{}
        try {
            foreach ($w in @(Get-CimInstance Win32_PrintJob -Property Name, JobId, PaperSize, PaperWidth, PaperLength -ErrorAction Stop)) {
                $wn = [string]$w.Name
                $cut = $wn.LastIndexOf(',')
                if ($cut -gt 0) { $paper[($wn.Substring(0, $cut).Trim().ToLowerInvariant() + '|' + [string]$w.JobId)] = $w }
            }
        } catch {}
        $out = @()
        foreach ($n in $Names) {
            try {
                foreach ($j in @(Get-PrintJob -PrinterName $n -ErrorAction Stop)) {
                    $pw = $paper[(([string]$n).ToLowerInvariant() + '|' + [string]$j.Id)]
                    if (-not $pw -and $j.PrinterName) { $pw = $paper[(([string]$j.PrinterName).ToLowerInvariant() + '|' + [string]$j.Id)] }
                    $out += [pscustomobject]@{
                        PaperSize    = $(if ($pw) { [string]$pw.PaperSize } else { '' })
                        PaperWidth   = $(if ($pw) { [int]$pw.PaperWidth } else { 0 })
                        PaperLength  = $(if ($pw) { [int]$pw.PaperLength } else { 0 })
                        Queue        = $n
                        Id           = $j.Id
                        PrinterName  = [string]$j.PrinterName
                        ComputerName = [string]$j.ComputerName
                        Document     = [string]$j.DocumentName
                        UserName     = [string]$j.UserName
                        JobStatus    = [string]$j.JobStatus
                        PagesPrinted = $j.PagesPrinted
                        TotalPages   = $j.TotalPages
                        Size         = $j.Size
                        Submitted    = $j.SubmittedTime
                    }
                }
            } catch {}
        }
        $out
    }

    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($j in @($raw)) {
        if (-not $j) { continue }
        $ageH = $null
        if ($j.Submitted) { try { $ageH = [Math]::Round(((Get-Date) - [datetime]$j.Submitted).TotalHours, 1) } catch {} }
        Add-Member -InputObject $j -NotePropertyName 'AgeHours' -NotePropertyValue $ageH -Force
        $jpc = Get-PaperCodeFromDimensions -Name ([string]$j.PaperSize) -Width ([int]$j.PaperWidth) -Length ([int]$j.PaperLength)
        Add-Member -InputObject $j -NotePropertyName 'PaperCode' -NotePropertyValue $jpc -Force
        if (-not $j.PaperSize -and $jpc) { $j.PaperSize = Get-PaperSizeName -Code $jpc }
        $list.Add($j)
    }
    $list.ToArray()
}

function Get-PrintServiceEvents {
    <#
    .SYNOPSIS
        Recent print-related events, plus whether each channel is even enabled.
    .DESCRIPTION
        PrintService/Admin is on by default; PrintService/Operational historically
        is NOT (though it was found enabled here) - so channel state is READ from
        the WINEVT Channels key and reported, never assumed. Operational is 1MB
        circular, so retention is short.

        Event IDs are the ones empirically confirmed present rather than the
        commonly-recited list. IDs 215 and 4909 are deliberately excluded - they
        could not be verified.
    #>
    [CmdletBinding()]
    param([int]$LookbackDays = 14, [int]$MaxEvents = 60)

    $since = (Get-Date).AddDays(-$LookbackDays)
    $channels = @(
        'Microsoft-Windows-PrintService/Admin',
        'Microsoft-Windows-PrintService/Operational'
    )

    $chanState = [System.Collections.Generic.List[object]]::new()
    foreach ($c in $channels) {
        $enabled = $null
        try {
            $k = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Channels\$c"
            $v = Get-ItemProperty -Path $k -Name 'Enabled' -ErrorAction Stop
            $enabled = [int]$v.Enabled
        } catch {}
        $chanState.Add([pscustomobject]@{ Channel = $c; Enabled = $enabled })
    }

    # 368 = spooler failed to verify a printer driver package (Point and Print)
    # 372 = document failed to print
    # 808/809/810/812 = plug-in load failure / driver-directory cleanup failures
    # 823 = default printer was changed (the "Let Windows manage" churn signal)
    $ids = @(368, 372, 808, 809, 810, 812, 823)

    $events = [System.Collections.Generic.List[object]]::new()
    foreach ($c in $channels) {
        try {
            $evts = Get-WinEvent -FilterHashtable @{ LogName = $c; StartTime = $since; Id = $ids } -MaxEvents $MaxEvents -ErrorAction Stop
            foreach ($e in $evts) {
                $events.Add([pscustomobject]@{
                    Channel = $c; Id = $e.Id; Time = $e.TimeCreated.ToString('yyyy-MM-dd HH:mm')
                    Level = [string]$e.LevelDisplayName
                    Message = (("$($e.Message)" -split "`r?`n")[0])
                })
            }
        } catch {}
    }

    # Spooler CRASHES never appear in PrintService - they land in System (SCM) and
    # Application (faulting spoolsv.exe).
    $crashes = [System.Collections.Generic.List[object]]::new()
    try {
        foreach ($e in @(Get-WinEvent -FilterHashtable @{ LogName='System'; StartTime=$since; Id=@(7031,7034) } -MaxEvents 40 -ErrorAction Stop)) {
            if ("$($e.Message)" -notmatch 'Print Spooler|Spooler') { continue }
            $crashes.Add([pscustomobject]@{
                Channel='System'; Id=$e.Id; Time=$e.TimeCreated.ToString('yyyy-MM-dd HH:mm')
                Level=[string]$e.LevelDisplayName; Message=(("$($e.Message)" -split "`r?`n")[0])
            })
        }
    } catch {}
    try {
        foreach ($e in @(Get-WinEvent -FilterHashtable @{ LogName='Application'; StartTime=$since; Id=1000 } -MaxEvents 40 -ErrorAction Stop)) {
            if ("$($e.Message)" -notmatch 'spoolsv\.exe') { continue }
            $crashes.Add([pscustomobject]@{
                Channel='Application'; Id=$e.Id; Time=$e.TimeCreated.ToString('yyyy-MM-dd HH:mm')
                Level=[string]$e.LevelDisplayName; Message='spoolsv.exe crash (Application 1000)'
            })
        }
    } catch {}

    [pscustomobject]@{
        Channels    = $chanState.ToArray()
        Events      = $events.ToArray()
        SpoolerCrashes = $crashes.ToArray()
        LookbackDays   = $LookbackDays
    }
}

function Get-PointAndPrintPolicy {
    <#
    .SYNOPSIS
        Point and Print / driver-install hardening, reported TRI-STATE.
    .DESCRIPTION
        "Not set" is NOT "not enforced". Post-2021 patches changed the ABSENT
        defaults: RestrictDriverInstallationToAdministrators defaults to 1
        (CVE-2021-34481) and RpcAuthnLevelPrivacyEnabled defaults to 1
        (CVE-2021-1678). Reporting absent as permissive sends techs down the wrong
        path on 0x11b and 0x709.

        The Point and Print GPO also writes a USER-scope copy under
        HKU\<SID>\Software\Policies\... - a SYSTEM script reading only HKLM misses
        it entirely, so every loaded hive is checked too.
    .OUTPUTS
        PSCustomObject { Machine[]; User[] } where each entry is
        { Name; Path; Value; IsSet; EffectiveValue; Meaning }
    #>
    [CmdletBinding()] param()

    $ppPath  = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint'
    $pkgPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PackagePointAndPrint'
    $ctlPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print'
    $drvPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DriverInstall\Restrictions'

    $defs = @(
        @{ Name='RestrictDriverInstallationToAdministrators'; Path=$ppPath;  Default=1; Meaning='Only admins may install printer drivers via Point and Print (CVE-2021-34481).' }
        @{ Name='NoWarningNoElevationOnInstall';              Path=$ppPath;  Default=0; Meaning='0 = warn and elevate on new connection install.' }
        @{ Name='UpdatePromptSettings';                       Path=$ppPath;  Default=0; Meaning='0 = warn and elevate on driver update.' }
        @{ Name='Restricted';                                 Path=$ppPath;  Default=$null; Meaning='Classic Point and Print Restrictions GPO master switch.' }
        @{ Name='TrustedServers';                             Path=$ppPath;  Default=$null; Meaning='Restrict to the ServerList below.' }
        @{ Name='ServerList';                                 Path=$ppPath;  Default=$null; Meaning='Semicolon-separated allowed print servers.' }
        @{ Name='InForest';                                   Path=$ppPath;  Default=$null; Meaning='Restrict to servers in the forest.' }
        @{ Name='PackagePointAndPrintOnly';                   Path=$pkgPath; Default=0; Meaning='1 = only PACKAGE-AWARE drivers may be installed (v3 drivers fail with 0x00000bcb).' }
        @{ Name='PackagePointAndPrintServerList';             Path=$pkgPath; Default=0; Meaning='1 = restrict package Point and Print to ListofServers.' }
        @{ Name='RpcAuthnLevelPrivacyEnabled';                Path=$ctlPath; Default=1; Meaning='RPC packet privacy for the spooler (CVE-2021-1678). Absent = ENFORCED.' }
        @{ Name='RestrictDriverInstallationToAdministrators'; Path=$drvPath; Default=1; Meaning='Driver-install restriction (alternate policy location).' }
    )

    $machine = [System.Collections.Generic.List[object]]::new()
    foreach ($d in $defs) {
        $val = $null; $isSet = $false
        try {
            $item = Get-ItemProperty -Path $d.Path -Name $d.Name -ErrorAction Stop
            $val = $item.($d.Name)
            $isSet = $true
        } catch {}
        $eff = $(if ($isSet) { $val } else { $d.Default })
        $machine.Add([pscustomobject]@{
            Name           = $d.Name
            Path           = $d.Path
            Value          = $val
            IsSet          = $isSet
            EffectiveValue = $eff
            State          = $(if ($isSet) { "Set=$val" } elseif ($null -ne $d.Default) { "NotSet -> effective default $($d.Default)" } else { 'NotSet' })
            Meaning        = $d.Meaning
        })
    }

    # user-scope copy of the Point and Print GPO
    $user = [System.Collections.Generic.List[object]]::new()
    foreach ($p in (Get-UserPrinterProfiles)) {
        $k = $null
        try { $k = [Microsoft.Win32.Registry]::Users.OpenSubKey("$($p.Sid)\Software\Policies\Microsoft\Windows NT\Printers\PointAndPrint") } catch {}
        if (-not $k) { continue }
        try {
            foreach ($vn in $k.GetValueNames()) {
                $user.Add([pscustomobject]@{
                    User = $p.User; Sid = $p.Sid; Name = $vn; Value = $k.GetValue($vn)
                    Path = "HKU\$($p.Sid)\Software\Policies\Microsoft\Windows NT\Printers\PointAndPrint"
                })
            }
        } catch {} finally { $k.Close() }
    }

    [pscustomobject]@{ Machine = $machine.ToArray(); User = $user.ToArray() }
}

function Get-PrinterErrorCodeHelp {
    <# .SYNOPSIS Decode the common "cannot connect to printer" Windows error codes. #>
    [CmdletBinding()]
    param([string]$Code)

    $map = @(
        [pscustomobject]@{ Code='0x0000011b'; Title='Windows cannot connect to the printer'
            Cause='RPC packet-privacy mismatch (CVE-2021-1678). The client demands RPC_C_AUTHN_LEVEL_PKT_PRIVACY and the server will not negotiate it.'
            Fix='PATCH THE PRINT SERVER. The widely-posted client-side RpcAuthnLevelPrivacyEnabled=0 is a security regression - do not apply it as a first resort.' }
        [pscustomobject]@{ Code='0x00000709'; Title='Double check the printer name'
            Cause='Default-printer set failure, or a post-KB5006670 connection failure. Very often a stale HKCU Device value naming a printer that no longer exists.'
            Fix='Clear/repoint the user Device value, confirm the share still exists on the server with its original name, then re-add the connection.' }
        [pscustomobject]@{ Code='0x00000bcb'; Title='The specified print server does not have a Windows NT printer driver installed'
            Cause='Server is missing a driver for the client architecture, OR PackagePointAndPrintOnly=1 and the driver is not package-aware.'
            Fix='Check Get-PrinterDriver IsPackageAware and MajorVersion on the server. v3 drivers are the usual casualty of PackagePointAndPrintOnly.' }
        [pscustomobject]@{ Code='0x000006e4'; Title='RPC_S_CANNOT_SUPPORT'
            Cause='Post-KB5005565. Client-side rendering / driver-install path is blocked.'
            Fix='Patch both ends; check RestrictDriverInstallationToAdministrators and install the driver locally as admin.' }
        [pscustomobject]@{ Code='0x00000bc4'; Title='No printers were found'
            Cause='Enumeration returned empty - RPC blocked, or the share ACL denies this user.'
            Fix='Verify TCP 135/445 to the server and the print share permissions for the user (not the machine account).' }
    )
    if ($Code) { return @($map | Where-Object { $_.Code -eq $Code }) }
    $map
}

# endregion


# ---------------------------------------------------------------------------
# region  THE DIAGNOSTIC - why is this mapped printer broken?
# ---------------------------------------------------------------------------

function New-PrinterFinding {
    <# .SYNOPSIS Build one finding record. Internal. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Code,
        [ValidateSet('ERROR','WARN','INFO')][string]$Severity = 'WARN',
        [string]$Scope = 'User',
        [string]$User = '',
        [string]$Printer = '',
        [string]$Server = '',
        [string]$Detail = '',
        [string]$LikelyErrorCode = '',
        [string]$Fix = ''
    )
    [pscustomobject]@{
        Code = $Code; Severity = $Severity; Scope = $Scope; User = $User
        Printer = $Printer; Server = $Server; Detail = $Detail
        LikelyErrorCode = $LikelyErrorCode; Fix = $Fix
    }
}

function Invoke-MappedPrinterDiagnostics {
    <#
    .SYNOPSIS
        Correlate every data source into named findings explaining WHY printers
        mapped from a print server are misbehaving.
    .DESCRIPTION
        This is the auxiliary diagnostic the toolkit exists for. Rather than
        dumping fields and leaving the tech to correlate them, it emits
        MAP001-MAP021 findings, each with the Windows error code it typically
        manifests as and a concrete fix.

        MAP008 is REPORTED, NEVER AUTO-APPLIED: setting RpcAuthnLevelPrivacyEnabled=0
        is the widely-posted "fix" for 0x11b and is a security regression.
    .PARAMETER StuckJobHours
        A job older than this is treated as stuck. Default 24.
    .OUTPUTS
        PSCustomObject { Findings[]; Summary; ServersChecked[]; UsersChecked[] }
    #>
    [CmdletBinding()]
    param(
        [int]$StuckJobHours = 24,
        [switch]$SkipReachability,
        [int]$ProbeTimeoutMs = 1500
    )

    $findings = [System.Collections.Generic.List[object]]::new()

    $spooler  = Get-SpoolerHealth
    $policy   = Get-PointAndPrintPolicy
    $local    = @(Get-LocalPrinterInventory -IncludeVirtual)
    $maps     = @(Get-AllUserPrinterMaps)
    $events   = Get-PrintServiceEvents
    $ctx      = Get-PrinterContext

    $localByName = @{}
    foreach ($p in $local) { $localByName[([string]$p.Name).ToLowerInvariant()] = $p }

    # ---- MAP015: spooler down / crashing (machine scope, checked first) ----
    if (-not $spooler.Running) {
        $findings.Add((New-PrinterFinding -Code 'MAP015' -Severity 'ERROR' -Scope 'Machine' `
            -Detail "Print Spooler is $($spooler.Status) (start mode $($spooler.StartMode)). Nothing can print and PrintManagement cmdlets will throw." `
            -Fix 'Start the Spooler service. If it is Disabled this is usually deliberate server hardening - confirm before changing.'))
    }
    if (@($events.SpoolerCrashes).Count -gt 0) {
        $last = @($events.SpoolerCrashes | Sort-Object Time -Descending)[0]
        $findings.Add((New-PrinterFinding -Code 'MAP015' -Severity 'ERROR' -Scope 'Machine' `
            -Detail "Spooler crash/unexpected-stop history: $(@($events.SpoolerCrashes).Count) event(s) in the last $($events.LookbackDays) days, most recent $($last.Time)." `
            -Fix 'A repeatedly crashing spooler is nearly always a bad third-party print driver or monitor. Clear the queue, then remove the suspect driver.'))
    }

    # ---- MAP010 (machine part): orphaned spool files ----
    if ($spooler.SpoolReadable -and $spooler.SpoolFileCount -gt 0 -and $spooler.OldestSpoolAgeHours -gt $StuckJobHours) {
        $findings.Add((New-PrinterFinding -Code 'MAP010' -Severity 'WARN' -Scope 'Machine' `
            -Detail "$($spooler.SpoolFileCount) spool file(s) in $($spooler.SpoolDirectory); oldest is $($spooler.OldestSpoolAgeHours)h old ($($spooler.OldestSpoolFile))." `
            -Fix 'Stop the Spooler, delete the .SPL/.SHD pairs, start it again (Clear-SpoolFiles / the GUI Spooler tab).'))
    }
    if (-not $spooler.SpoolReadable) {
        $findings.Add((New-PrinterFinding -Code 'MAP010' -Severity 'INFO' -Scope 'Machine' `
            -Detail "Spool directory $($spooler.SpoolDirectory) is not readable from this context - stuck-file detection was skipped." `
            -Fix 'Re-run as SYSTEM or an administrator to include the spool-file check.'))
    }

    # ---- MAP007 / MAP008 / MAP006 (policy, machine scope) ----
    $polByName = @{}
    foreach ($m in @($policy.Machine)) { if (-not $polByName.ContainsKey($m.Name)) { $polByName[$m.Name] = $m } }

    $restrict = $polByName['RestrictDriverInstallationToAdministrators']
    if ($restrict -and [int]$restrict.EffectiveValue -eq 1) {
        $sev = if ($restrict.IsSet) { 'INFO' } else { 'INFO' }
        $findings.Add((New-PrinterFinding -Code 'MAP007' -Severity $sev -Scope 'Machine' `
            -Detail "RestrictDriverInstallationToAdministrators is $($restrict.State). Non-admin users cannot install a printer driver via Point and Print." `
            -LikelyErrorCode '0x0000011b / 0x00000709' `
            -Fix 'Expected and correct post-CVE-2021-34481. Pre-stage the driver on the client (pnputil / package-aware v4 driver) rather than weakening this.'))
    }
    $rpc = $polByName['RpcAuthnLevelPrivacyEnabled']
    if ($rpc -and $rpc.IsSet -and [int]$rpc.Value -eq 0) {
        $findings.Add((New-PrinterFinding -Code 'MAP008' -Severity 'WARN' -Scope 'Machine' `
            -Detail 'RpcAuthnLevelPrivacyEnabled is explicitly 0. This is the widely-posted client-side workaround for 0x0000011b and it DOWNGRADES spooler RPC security.' `
            -LikelyErrorCode '0x0000011b' `
            -Fix 'Patch the print server and remove this value. Reported only - this toolkit never sets it.'))
    }
    $pkgOnly = $polByName['PackagePointAndPrintOnly']

    # ---- Paper size and duplex (MAP018 / MAP019 / MAP020 / MAP021) ----
    # The region decides what paper is "right": this PC's system region for local
    # queues, each user's own region for them and their connections. Duplex has no
    # right answer, so only a disagreement is reported (MAP021). Registry reads only.
    $machinePaper = $null
    try { $machinePaper = Get-RegionPaperSize -Sid '.DEFAULT' } catch {}
    $machinePaperCode = if ($machinePaper) { [int]$machinePaper.ExpectedPaperCode } else { 0 }

    # MAP021 - a user's own copy of a printer uses a different duplex setting from
    # the default beside it. Their copy replaces the default for them, which is how
    # two users of one printer get different duplexing.
    $duplexFinding = {
        param($Row, [string]$User, [string]$Server)
        $isConn = ($Row.Kind -eq 'Connection')
        $what = if ($isConn) { "the print server's default (as cached on this PC)" } else { 'the queue default' }
        New-PrinterFinding -Code 'MAP021' -Severity 'INFO' -Scope 'User' -User $User -Printer $Row.Queue -Server $Server `
            -Detail "This user's own copy of the printer's settings is $(Get-DuplexLabel $Row.UserDuplex), while $what is $(Get-DuplexLabel $Row.QueueDuplex). Their copy replaces the default for them, so changing the default does not change what they get." `
            -Fix ("If the default is what they should get: as $User, open Printing Preferences for this printer and choose $(Get-DuplexLabel $Row.QueueDuplex). If their setting is right for everyone, change the default instead" +
                  $(if ($isConn) { " on $Server." } else { ' (Printers tab -> Paper and duplex...).' }) + " This console does not reset a user's own printer preferences.")
    }

    # ---- Per-user connection analysis ----
    $serverCache = @{}
    $serverInvCache = @{}
    $usersChecked = [System.Collections.Generic.List[string]]::new()

    foreach ($map in $maps) {
        $usersChecked.Add($map.User)

        # MAP018 - the user's regional paper size disagrees with their own region.
        # Word and most applications take a new document's size from it, so a
        # profile left on English (United States) sends Letter to an A4 office.
        $userPaper = $null
        try { $userPaper = Get-RegionPaperSize -Sid $map.Sid -User $map.User } catch {}
        if ($userPaper -and $userPaper.Match -eq $false) {
            $findings.Add((New-PrinterFinding -Code 'MAP018' -Severity 'WARN' -Scope 'User' -User $map.User `
                -Detail "Regional paper size is $($userPaper.UserPaper), but $($userPaper.ExpectedFrom) implies $($userPaper.ExpectedPaper). New documents default to $($userPaper.UserPaper), so a printer loaded with $($userPaper.ExpectedPaper) stops and asks for paper." `
                -Fix "GUI: Printers tab -> Paper and duplex... -> set this user's regional paper size to $($userPaper.ExpectedPaper). Only that one setting changes; programs already open pick it up when restarted."))
        }
        $upc = if ($userPaper -and $userPaper.ExpectedPaperCode) { [int]$userPaper.ExpectedPaperCode } else { $machinePaperCode }
        $paperRows = @()
        try { $paperRows = @(Get-PrinterPaperSizeState -Sid $map.Sid -MachineExpectedCode $machinePaperCode -UserExpectedCode $upc) } catch {}

        # This user's own saved preferences on LOCAL queues (DevModePerUser). The
        # queue default itself is MAP019, at machine scope, so only the user's copy
        # is reported here.
        foreach ($lr in @($paperRows | Where-Object { $_.Kind -eq 'Local' -and $_.HasUserCopy })) {
            if ($lr.UserCopyState -eq 'Mismatch') {
                $findings.Add((New-PrinterFinding -Code 'MAP020' -Severity 'INFO' -Scope 'User' -User $map.User -Printer $lr.Queue `
                    -Detail "This user's own saved preferences for '$($lr.Queue)' use $($lr.UserCopy), while the queue default is $($lr.QueueDefault) and their region expects $(Get-PaperSizeName -Code $upc). Their copy replaces the queue default for them." `
                    -Fix "As $($map.User), open Printing Preferences for this printer and choose $(Get-PaperSizeName -Code $upc). This console does not reset a user's own printer preferences."))
            }
            if ($lr.DuplexState -eq 'Differs') { $findings.Add((& $duplexFinding $lr $map.User '')) }
        }

        # MAP004: default will be silently overwritten by Windows
        if (@($map.Connections).Count -gt 0 -or $map.DefaultPrinter) {
            $legacy = $map.LegacyDefaultPrinterMode
            if ($null -eq $legacy -or [int]$legacy -ne 1) {
                $findings.Add((New-PrinterFinding -Code 'MAP004' -Severity 'WARN' -Scope 'User' -User $map.User `
                    -Detail "LegacyDefaultPrinterMode is $(if ($null -eq $legacy) { 'not set' } else { $legacy }) - 'Let Windows manage my default printer' is ON, so Windows will overwrite this user's default after the next print." `
                    -Fix 'Set LegacyDefaultPrinterMode=1 for the user (the GUI does this automatically whenever it sets a default).'))
            }
        }

        # MAP003: default printer names something that does not exist
        if ($map.DefaultPrinter) {
            $dp = $map.DefaultPrinter
            $inConnections = @($map.Connections | Where-Object { $_.Unc -ieq $dp -or $_.Printer -ieq $dp })
            $inSpooler     = $localByName.ContainsKey($dp.ToLowerInvariant())
            if ($inConnections.Count -eq 0 -and -not $inSpooler) {
                $findings.Add((New-PrinterFinding -Code 'MAP003' -Severity 'ERROR' -Scope 'User' -User $map.User -Printer $dp `
                    -Detail "The user's default printer is '$dp' but no such queue exists in their connections or in the spooler." `
                    -LikelyErrorCode '0x00000709' `
                    -Fix 'Set a valid default for the user (GUI: Printers tab -> Set Default). This rewrites Device and LegacyDefaultPrinterMode together.'))
            }
        }

        # MAP009: stale per-user entries with no live queue
        foreach ($d in @($map.Devices)) {
            $n = [string]$d.Name
            if (-not $n) { continue }
            if ($n -match '^\\\\') { continue }   # connections handled separately
            if ($localByName.ContainsKey($n.ToLowerInvariant())) { continue }
            $findings.Add((New-PrinterFinding -Code 'MAP009' -Severity 'INFO' -Scope 'User' -User $map.User -Printer $n `
                -Detail "Per-user registry lists printer '$n' but the spooler has no such queue - a stale entry from a removed printer." `
                -Fix 'Harmless, but it can make a phantom printer appear in old applications. Remove the value from the user Devices/PrinterPorts keys if it confuses users.'))
        }

        foreach ($conn in @($map.Connections)) {
            $srv = $conn.Server
            $unc = $conn.Unc
            if (-not $srv) { continue }

            # MAP020 - this user's copy of the connection, or the print server's
            # default as cached on this PC, sends a different paper size from the
            # region. Read from the client-side cache, so it needs no reachability.
            $paperRow = @($paperRows | Where-Object { [string]$_.Queue -ieq [string]$unc })[0]
            if ($paperRow -and $paperRow.State -eq 'Mismatch') {
                $srvPaperBad = ($paperRow.ServerState -eq 'Mismatch')
                $findings.Add((New-PrinterFinding -Code 'MAP020' -Severity $(if ($srvPaperBad) { 'WARN' } else { 'INFO' }) -Scope 'User' -User $map.User `
                    -Printer $unc -Server $srv `
                    -Detail $(if ($srvPaperBad) { "The print server's default for this printer is $($paperRow.QueueDefault) (as cached on this PC), but the region expects $($paperRow.Expected)." } else { "This user's copy of the printer's settings uses $($paperRow.UserCopy), while the print server's default is $(if ($paperRow.QueueDefault) { $paperRow.QueueDefault } else { 'not cached on this PC' })." }) `
                    -Fix $(if ($srvPaperBad) { "On $srv, run this console: Printers tab -> Paper and duplex... -> set the queue default to $($paperRow.Expected)." } else { "As $($map.User), open Printing Preferences for this printer and choose $($paperRow.Expected). This console does not reset a user's own printer preferences." })))
            }
            if ($paperRow -and $paperRow.DuplexState -eq 'Differs') { $findings.Add((& $duplexFinding $paperRow $map.User $srv)) }

            # --- server reachability (cached per server) ---
            if (-not $SkipReachability) {
                if (-not $serverCache.ContainsKey($srv)) {
                    $serverCache[$srv] = Test-PrintServerReachable -Server $srv -TimeoutMs $ProbeTimeoutMs
                }
                $reach = $serverCache[$srv]
                if (-not $reach.Reachable) {
                    $findings.Add((New-PrinterFinding -Code 'MAP001' -Severity 'ERROR' -Scope 'User' -User $map.User `
                        -Printer $unc -Server $srv `
                        -Detail "Print server '$srv' does not answer on SMB 445 or RPC 135. $($reach.Message)" `
                        -LikelyErrorCode '0x00000bcb / hang' `
                        -Fix 'Check DNS for the server name, VPN/routing, and that the server is up. Nothing else here can be diagnosed until it answers.'))
                    continue
                }

                # --- MAP002: queue missing from the server's share list ---
                if (-not $serverInvCache.ContainsKey($srv)) {
                    $serverInvCache[$srv] = Get-PrintServerInventory -Server $srv
                }
                $inv = $serverInvCache[$srv]
                if (@($inv.Printers).Count -gt 0) {
                    $wanted = $conn.Printer
                    $hit = @($inv.Printers | Where-Object {
                        $_.ShareName -ieq $wanted -or $_.Name -ieq $wanted -or $_.Unc -ieq $unc
                    })
                    if ($hit.Count -eq 0) {
                        $findings.Add((New-PrinterFinding -Code 'MAP002' -Severity 'ERROR' -Scope 'User' -User $map.User `
                            -Printer $unc -Server $srv `
                            -Detail "'$srv' is reachable but does not publish a queue matching '$wanted' (enumerated via $($inv.Method))." `
                            -LikelyErrorCode '0x00000709' `
                            -Fix 'The share was renamed or removed on the server. Re-add the connection using the current share name (GUI: Print Server tab).'))
                    } else {
                        # --- MAP005: client/server driver mismatch ---
                        $srvDriver = [string]$hit[0].DriverName
                        $clientQueue = $localByName[$unc.ToLowerInvariant()]
                        if ($srvDriver -and $clientQueue -and $clientQueue.DriverName -and
                            $clientQueue.DriverName -ne $srvDriver) {
                            $findings.Add((New-PrinterFinding -Code 'MAP005' -Severity 'WARN' -Scope 'User' -User $map.User `
                                -Printer $unc -Server $srv `
                                -Detail "Client driver '$($clientQueue.DriverName)' does not match the server's '$srvDriver'." `
                                -LikelyErrorCode '0x00000bcb' `
                                -Fix 'Remove and re-add the connection so the current server driver is pulled down.'))
                        }

                        # --- MAP011/MAP012/MAP013 against the real device ---
                        $devHost = [string]$hit[0].HostAddress
                        if ($devHost) {
                            $probe = Test-PrinterPortReachable -HostName $devHost -TimeoutMs $ProbeTimeoutMs
                            if (-not $probe.Reachable) {
                                $findings.Add((New-PrinterFinding -Code 'MAP011' -Severity 'WARN' -Scope 'Server' -User $map.User `
                                    -Printer $unc -Server $srv `
                                    -Detail "Device $devHost is not answering on 9100/515/631. $($probe.Message)" `
                                    -Fix 'The printer is powered off, asleep, or has changed IP. Wake it and confirm its address matches the port on the server.'))
                            }
                        }
                        $portName = [string]$hit[0].PortName
                        if ($portName -like 'WSD*') {
                            $findings.Add((New-PrinterFinding -Code 'MAP013' -Severity 'INFO' -Scope 'Server' -User $map.User `
                                -Printer $unc -Server $srv `
                                -Detail "The server queue uses a WSD port ($portName). WSD ports drop out intermittently and cannot be repointed." `
                                -Fix 'Rebuild the queue on a Standard TCP/IP RAW 9100 port. See EPSON_PRINT_DRIVERS\Update_WSD_to_RAW_9100_Port.ps1.'))
                        }
                    }
                }
            }

            # --- MAP006: package-only policy versus a v3 / non-package driver ---
            $clientQueue = $localByName[$unc.ToLowerInvariant()]
            if ($pkgOnly -and [int]$pkgOnly.EffectiveValue -eq 1 -and $clientQueue) {
                $notAware = ($null -ne $clientQueue.PackageAware -and -not $clientQueue.PackageAware)
                $isV3     = ($null -ne $clientQueue.DriverMajor -and [int]$clientQueue.DriverMajor -eq 3)
                if ($notAware -or $isV3) {
                    $findings.Add((New-PrinterFinding -Code 'MAP006' -Severity 'WARN' -Scope 'User' -User $map.User `
                        -Printer $unc -Server $srv `
                        -Detail "PackagePointAndPrintOnly is $($pkgOnly.State) but driver '$($clientQueue.DriverName)' is $(if ($isV3) { 'a v3 driver' } else { 'not package-aware' })." `
                        -LikelyErrorCode '0x00000bcb' `
                        -Fix 'Deploy a package-aware (v4) driver on the server, or pre-stage the driver package on the client with pnputil.'))
                }
            }

            # --- MAP014: queue offline / paused on this client ---
            if ($clientQueue) {
                if ($clientQueue.WorkOffline) {
                    $findings.Add((New-PrinterFinding -Code 'MAP014' -Severity 'WARN' -Scope 'User' -User $map.User `
                        -Printer $unc -Server $srv `
                        -Detail 'The client queue is flagged Use Printer Offline.' `
                        -Fix 'Clear "Use Printer Offline" on the queue, then confirm the server and device are reachable.'))
                }
                if ("$($clientQueue.StateText)" -in @('Paused', 'Error', 'Offline')) {
                    $findings.Add((New-PrinterFinding -Code 'MAP014' -Severity 'WARN' -Scope 'User' -User $map.User `
                        -Printer $unc -Server $srv `
                        -Detail "The client queue state is '$($clientQueue.StateText)'." `
                        -Fix 'Resume the queue and clear any error, then retry a test page.'))
                }
            }
        }
    }

    # ---- Machine-scope pass over LOCAL queues (MAP014 / MAP016 / MAP013 / MAP017) ----
    # Enumerated once: PnP is comparatively slow and the answer is the same for
    # every queue.
    $usbDevices = @()
    try { $usbDevices = @(Get-UsbPrinterDevice) } catch {}

    # The per-connection loop above only ever sees MAPPED printers. A local queue
    # left on "Use Printer Offline", or auto-added over WSD with the generic IPP
    # class driver, would otherwise never be reported at all.
    $localPaper = @{}
    try {
        foreach ($lp in @(Get-PrinterPaperSizeState -MachineExpectedCode $machinePaperCode)) { $localPaper[([string]$lp.Queue).ToLowerInvariant()] = $lp }
    } catch {}
    foreach ($p in $local) {
        # MAP019 - the queue's stored default paper size disagrees with this PC's
        # region. Checked before the virtual-queue skip: a PDF printer left on
        # Letter still makes Letter documents. On a print server these are the
        # shared queues.
        if ($p.Type -ne 'Connection') {
            $pq = $localPaper[([string]$p.Name).ToLowerInvariant()]
            if ($pq -and $pq.ServerState -eq 'Mismatch') {
                $findings.Add((New-PrinterFinding -Code 'MAP019' -Severity $(if ($p.IsVirtual) { 'INFO' } else { 'WARN' }) -Scope 'Machine' -Printer $p.Name `
                    -Detail ("Local queue '$($p.Name)' defaults to $($pq.QueueDefault), but this PC's region expects $($pq.Expected)" + $(if ($machinePaper -and $machinePaper.ExpectedFrom) { " ($($machinePaper.ExpectedFrom))." } else { '.' })) `
                    -Fix "GUI: Printers tab -> Paper and duplex... -> set the queue default to $($pq.Expected). On a print server this is the default new connections get; a user who already has their own copy may keep it (MAP020)."))
            }
        }
        if ($p.IsVirtual) { continue }
        if ($p.Type -eq 'Connection') { continue }   # covered per-user above

        # MAP017 - a USB queue whose printer is not physically attached. Checked
        # BEFORE the WorkOffline test because an absent device is the root cause
        # and WorkOffline is merely its symptom; reporting both would be two
        # findings for one fault.
        $usbMissing = $false
        if ($p.PortType -eq 'USB') {
            $usb = Test-UsbPrinterAttached -PortName $p.PortName -Name $p.Name -DriverName $p.DriverName -Devices $usbDevices
            if (-not $usb.Attached) {
                $usbMissing = $true
                $findings.Add((New-PrinterFinding -Code 'MAP017' -Severity 'WARN' -Scope 'Machine' -Printer $p.Name `
                    -Detail "Local queue '$($p.Name)' is on $($p.PortName) but no USB printer is attached. $($usb.Message)" `
                    -Fix 'Check the printer is powered on and its USB cable is connected to this PC, then re-check. Note the queue and the USB port keep existing regardless, and PrinterStatus still reads Normal - neither proves the printer is there.'))
            }
        }

        if ($p.WorkOffline -and -not $usbMissing) {
            $findings.Add((New-PrinterFinding -Code 'MAP014' -Severity 'WARN' -Scope 'Machine' -Printer $p.Name `
                -Detail "Local queue '$($p.Name)' is flagged Use Printer Offline - jobs queue up and never print." `
                -Fix 'Clear "Use Printer Offline" on the queue (right-click the printer, See what is printing, Printer menu).'))
        }
        if ("$($p.StateText)" -in @('Paused', 'Error')) {
            $findings.Add((New-PrinterFinding -Code 'MAP014' -Severity 'WARN' -Scope 'Machine' -Printer $p.Name `
                -Detail "Local queue '$($p.Name)' is in state '$($p.StateText)'." `
                -Fix 'Resume the queue and clear the error, then send a test page.'))
        }

        # MAP016 - generic IPP class driver. Rebuild candidate.
        if ($p.IsIppClass) {
            $findings.Add((New-PrinterFinding -Code 'MAP016' -Severity 'WARN' -Scope 'Machine' -Printer $p.Name `
                -Detail "Local queue '$($p.Name)' uses the generic '$($p.DriverName)'. Windows binds this when a printer is auto-added; it usually loses trays, duplex, finishing and accounting." `
                -Fix 'Install the real vendor driver, then rebuild the queue onto a Standard TCP/IP RAW:9100 port (GUI: Printers tab -> Repoint / Rebuild). Note the class driver often prints BLANK over RAW:9100, so the driver must be replaced at the same time - not just the port.'))
        }

        # MAP013 - WSD port on a LOCAL queue. WSD is local-subnet discovery only,
        # so the binding dies the moment the device moves or changes subnet.
        if ($p.PortType -eq 'WSD') {
            $findings.Add((New-PrinterFinding -Code 'MAP013' -Severity 'WARN' -Scope 'Machine' -Printer $p.Name `
                -Detail "Local queue '$($p.Name)' is on a WSD port ($($p.PortName)). WSD resolves the device by multicast within one subnet and records no fixed IP, so it breaks on any move or subnet change and cannot be repointed." `
                -Fix 'Rebuild the queue onto a Standard TCP/IP RAW:9100 port pinned to the device IP (GUI: Printers tab -> Repoint / Rebuild).'))
        }
    }

    # ---- MAP012: SNMP configured on a port but the device does not answer 161 ----
    if (-not $SkipReachability) {
        foreach ($p in $local) {
            if ($p.IsVirtual) { continue }
            if (-not $p.HostAddress) { continue }
            if ($null -eq $p.SNMPEnabled -or -not $p.SNMPEnabled) { continue }
            $snmp = Test-TcpPortFast -HostName $p.HostAddress -Port 161 -TimeoutMs $ProbeTimeoutMs
            # SNMP is UDP; a closed TCP 161 is only a hint, so this stays INFO and
            # is only raised when the print ports themselves are also unreachable.
            if (-not $snmp.Open) {
                $probe = Test-PrinterPortReachable -HostName $p.HostAddress -TimeoutMs $ProbeTimeoutMs
                if (-not $probe.Reachable) {
                    $findings.Add((New-PrinterFinding -Code 'MAP012' -Severity 'INFO' -Scope 'Machine' -Printer $p.Name `
                        -Detail "Port '$($p.PortName)' has SNMP status enabled but $($p.HostAddress) is not responding. The spooler's SNMP status check times out and flips the queue to Offline." `
                        -Fix 'Either bring the device back online, or clear "SNMP Status Enabled" on the port to stop the false Offline state.'))
                }
            }
        }
    }

    # ---- MAP010 (job part): stuck jobs ----
    foreach ($j in (Get-PrinterJobSnapshot)) {
        if ($null -eq $j.AgeHours -or $j.AgeHours -lt $StuckJobHours) { continue }
        $findings.Add((New-PrinterFinding -Code 'MAP010' -Severity 'WARN' -Scope 'Machine' -Printer $j.Queue -User ([string]$j.UserName) `
            -Detail "Job $($j.Id) '$($j.Document)' has been queued $($j.AgeHours)h (status $($j.JobStatus), $($j.PagesPrinted)/$($j.TotalPages) pages)." `
            -Fix 'Cancel the job (GUI: Spooler and Jobs tab). A Complete job with 0 pages printed that never leaves the queue is the classic stuck-job signature.'))
    }

    # ---- MAP009 (machine part): no user has any mapped printer ----
    if (@($maps).Count -eq 0) {
        $findings.Add((New-PrinterFinding -Code 'MAP009' -Severity 'INFO' -Scope 'Machine' `
            -Detail 'No user profile hives are loaded, so no per-user mapped printers could be read. Nobody is signed in.' `
            -Fix 'Run this while the affected user is signed in - mapped printers exist only inside their own hive.'))
    }

    $arr = $findings.ToArray()
    $errors = @($arr | Where-Object { $_.Severity -eq 'ERROR' }).Count
    $warns  = @($arr | Where-Object { $_.Severity -eq 'WARN'  }).Count
    $infos  = @($arr | Where-Object { $_.Severity -eq 'INFO'  }).Count

    [pscustomobject]@{
        Findings       = $arr
        ErrorCount     = $errors
        WarnCount      = $warns
        InfoCount      = $infos
        Summary        = "$errors error(s), $warns warning(s), $infos note(s)."
        ServersChecked = @($serverCache.Keys)
        UsersChecked   = $usersChecked.ToArray()
        Context        = $ctx
    }
}

# endregion


# ---------------------------------------------------------------------------
# region  In-session execution (the ONLY way a SYSTEM process changes a user's
#         printers - see the header notes)
# ---------------------------------------------------------------------------

function Start-PrinterUserTask {
    <#
    .SYNOPSIS
        Register + start (non-blocking) a one-shot scheduled task that runs a
        command in the interactive user's session. Caller polls for completion.
    .DESCRIPTION
        RunLevel=Limited + LogonType=Interactive is the standard token Explorer
        uses, so the result is exactly what the user would get themselves.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$CmdPath,
        [string]$User,
        [string]$TaskName = 'Printer Toolkit (in-session)'
    )
    if (-not $User) {
        $iu = @(Get-InteractiveUsers)
        if ($iu.Count) { $User = $iu[0].User }
    }
    if (-not $User)                             { return [pscustomobject]@{ Started=$false; User=$null; Message='No interactive user is logged on - there is no session to act in.' } }
    if (-not (Test-Path -LiteralPath $CmdPath)) { return [pscustomobject]@{ Started=$false; User=$User; Message="Command file not found: $CmdPath" } }
    try {
        $action    = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument ('/c "' + $CmdPath + '"')
        $principal = New-ScheduledTaskPrincipal -UserId $User -LogonType Interactive -RunLevel Limited
        $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        Register-ScheduledTask -TaskName $TaskName -Action $action -Principal $principal -Settings $settings -Force -ErrorAction Stop | Out-Null
        Start-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        return [pscustomobject]@{ Started=$true; User=$User; Message="Started in $User's session." }
    } catch {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        return [pscustomobject]@{ Started=$false; User=$User; Message="Could not start in-session task: $($_.Exception.Message)" }
    }
}

function Invoke-CommandInPrinterUserSession {
    <#
    .SYNOPSIS
        Run a command line in the interactive user's session and return its output.
    .DESCRIPTION
        Writes a temp .cmd under Public\Documents (writable by the standard-user
        task, readable by SYSTEM), runs it as the user, reads the redirected output
        back, cleans up.
    .OUTPUTS
        PSCustomObject { Success; Output; Message; User }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$CommandLine,
        [string]$User,
        [int]$TimeoutSec = 30,
        [scriptblock]$OnPoll
    )
    if (-not $User) {
        $iu = @(Get-InteractiveUsers)
        if ($iu.Count) { $User = $iu[0].User }
    }
    if (-not $User) { return [pscustomobject]@{ Success=$false; Output=$null; Message='No interactive user is logged on.'; User=$null } }

    $pubDocs = Join-Path $env:PUBLIC 'Documents'
    if (-not (Test-Path $pubDocs)) { $pubDocs = $env:PUBLIC }
    $stamp   = [guid]::NewGuid().ToString('N').Substring(0, 8)
    $outFile = Join-Path $pubDocs "prt-$stamp.out"
    $batFile = Join-Path $pubDocs "prt-$stamp.cmd"
    Set-Content -LiteralPath $batFile -Value ("@echo off`r`n( $CommandLine ) > `"$outFile`" 2>&1`r`n") -Encoding ASCII

    $taskName = 'Printer Toolkit probe'
    $out = $null; $msg = ''
    try {
        $start = Start-PrinterUserTask -CmdPath $batFile -User $User -TaskName $taskName
        if (-not $start.Started) { return [pscustomobject]@{ Success=$false; Output=$null; Message=$start.Message; User=$User } }

        # WAITING FOR THE TASK - the two result codes that matter:
        #   0x41301 (267009) the task is running NOW
        #   0x41303 (267011) the task HAS NOT YET RUN
        # A freshly registered task reports 0x41303 - measured, not assumed. The
        # original loop only excluded 0x41301, so 0x41303 counted as "finished"
        # and the wait could fall straight through before the task had even
        # started. Launching powershell.exe in another user's session takes well
        # over the 300ms first poll on a busy machine, so this lost the result
        # and reported a failure for an operation that had usually SUCCEEDED.
        $deadline = (Get-Date).AddSeconds($TimeoutSec)
        $done = $false
        $everRan = $false
        do {
            Start-Sleep -Milliseconds 300
            if ($OnPoll) { try { & $OnPoll } catch {} }
            $info = Get-ScheduledTaskInfo -TaskName $taskName -ErrorAction SilentlyContinue
            $lr = $null
            if ($info) { try { $lr = [int64]$info.LastTaskResult } catch {} }
            if ($null -ne $lr -and $lr -eq 0x41301) { $everRan = $true }
            $done = ($null -ne $lr -and $lr -ne 0x41301 -and $lr -ne 0x41303)
        } while (-not $done -and (Get-Date) -lt $deadline)

        # The '>' redirect creates the output file the instant cmd starts, so its
        # EXISTENCE proves nothing. Wait for actual content, then judge on that.
        $flushUntil = (Get-Date).AddSeconds(3)
        do {
            Start-Sleep -Milliseconds 150
            if ($OnPoll) { try { & $OnPoll } catch {} }
            if (Test-Path -LiteralPath $outFile) {
                $out = Get-Content -LiteralPath $outFile -Raw -ErrorAction SilentlyContinue
            }
        } while (-not $out -and (Get-Date) -lt $flushUntil)

        if ($out) {
            $msg = 'Ran in the user session.'
        } elseif ($done) {
            $msg = "The task ran in ${User}'s session but wrote no output, so the result could not be read. The action itself may still have succeeded - verify before assuming it failed."
        } elseif ($everRan) {
            $msg = "The task was still running in ${User}'s session after $TimeoutSec s, so the result could not be read. The action may still have succeeded - verify before assuming it failed."
        } else {
            $msg = "The task never started in ${User}'s session within $TimeoutSec s - the session could not run it."
        }
    } catch {
        $msg = "Could not run in the user's session: $($_.Exception.Message)"
    } finally {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $batFile -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue
    }
    [pscustomobject]@{ Success = ($null -ne $out); Output = $out; Message = $msg; User = $User }
}

function Invoke-ScriptInPrinterUserSession {
    <#
    .SYNOPSIS
        Run a scriptblock as the target user and return its parsed JSON output.
    .DESCRIPTION
        In-process when we already ARE that user, via a one-shot task when SYSTEM -
        so callers get ONE code path for both Backstage-as-SYSTEM and an ordinary
        elevated session.
    .PARAMETER Vars
        Simple values (string/bool/int) injected as assignments at the top of the
        script, so the same scriptblock works in-process and injected.
    .OUTPUTS
        PSCustomObject { Success; Data; Raw; Message; User }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [hashtable]$Vars,
        [string]$User,
        [int]$TimeoutSec = 45,
        [scriptblock]$OnPoll
    )
    $prefix = ''
    if ($Vars) {
        foreach ($k in $Vars.Keys) {
            $v = $Vars[$k]
            $lit = if ($v -is [bool]) { if ($v) { '$true' } else { '$false' } }
                   elseif ($v -is [int]) { "$v" }
                   else { "'" + (("$v") -replace "'", "''") + "'" }
            $prefix += "`$$k = $lit`r`n"
        }
    }
    $full = $prefix + $ScriptBlock.ToString()
    $isSystem = ((([Security.Principal.WindowsIdentity]::GetCurrent()).User.Value) -eq 'S-1-5-18')

    $currentUser = "$env:USERDOMAIN\$env:USERNAME"
    if (-not $isSystem -and (-not $User -or $User -ieq $currentUser)) {
        $raw = ''
        try { $raw = (& ([scriptblock]::Create($full)) 2>&1 | Out-String) } catch { $raw = "$($_.Exception.Message)" }
        $data = ConvertFrom-NoisyJson $raw
        return [pscustomobject]@{
            Success = [bool]$data; Data = $data; Raw = "$raw".Trim()
            Message = $(if ($data) { 'ok (in-process)' } else { 'no/invalid JSON returned' }); User = $currentUser
        }
    }

    if (-not $User) {
        $iu = @(Get-InteractiveUsers)
        if ($iu.Count) { $User = $iu[0].User }
    }
    if (-not $User) { return [pscustomobject]@{ Success=$false; Data=$null; Raw=$null; Message='No interactive user is logged on.'; User=$null } }

    $pubDocs = Join-Path $env:PUBLIC 'Documents'
    if (-not (Test-Path $pubDocs)) { $pubDocs = $env:PUBLIC }
    $ps1 = Join-Path $pubDocs ("prt-run-{0}.ps1" -f [guid]::NewGuid().ToString('N').Substring(0, 8))
    $body = "`$ErrorActionPreference='SilentlyContinue'; `$ProgressPreference='SilentlyContinue'`r`n" + $full
    Set-Content -LiteralPath $ps1 -Value $body -Encoding UTF8
    try {
        $r = Invoke-CommandInPrinterUserSession -CommandLine ('powershell -NoProfile -ExecutionPolicy Bypass -File "' + $ps1 + '"') `
                -User $User -TimeoutSec $TimeoutSec -OnPoll $OnPoll
        $raw  = "$($r.Output)"
        $data = ConvertFrom-NoisyJson $raw
        return [pscustomobject]@{
            Success = [bool]$data; Data = $data; Raw = $raw.Trim()
            Message = $(if ($data) { 'ok (user session)' } else { $r.Message }); User = $User
        }
    } finally {
        Remove-Item -LiteralPath $ps1 -Force -ErrorAction SilentlyContinue
    }
}

# endregion


# ---------------------------------------------------------------------------
# region  Actions
# ---------------------------------------------------------------------------

function Rename-PrinterSafe {
    <#
    .SYNOPSIS
        Rename a local print queue.
    .DESCRIPTION
        REFUSES connection printers: a connection printer's name IS its UNC path,
        derived from the server's share name, so renaming it locally is not
        meaningful - the queue must be renamed on the server.

        Rename-Printer has NO -ComputerName parameter (unlike Get-Printer and
        Remove-Printer). Remote renames need -CimSession/WinRM plus admin on the
        server, which is out of scope here.
    .OUTPUTS
        PSCustomObject { Success; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$NewName
    )
    if ($Name -match '^\\\\') {
        return [pscustomobject]@{ Success=$false; Message="'$Name' is a connection to a shared queue. Its name is the server's UNC path - rename the queue on the print server instead, then re-add the connection." }
    }
    if ([string]::IsNullOrWhiteSpace($NewName)) {
        return [pscustomobject]@{ Success=$false; Message='New name is empty.' }
    }
    if ($NewName -match ',') {
        return [pscustomobject]@{ Success=$false; Message='Printer names containing a comma break the per-user Device/Connections registry format. Choose a name without commas.' }
    }
    if (-not $PSCmdlet.ShouldProcess($Name, "Rename to '$NewName'")) {
        return [pscustomobject]@{ Success=$false; Message='Skipped (WhatIf).' }
    }
    try {
        Rename-Printer -Name $Name -NewName $NewName -ErrorAction Stop
        return [pscustomobject]@{ Success=$true; Message="Renamed '$Name' to '$NewName'." }
    } catch {
        $first = $_.Exception.Message
        try {
            $esc = $Name -replace "'", "''"
            $p = Get-CimInstance Win32_Printer -Filter "Name='$esc'" -ErrorAction Stop
            $r = Invoke-CimMethod -InputObject $p -MethodName RenamePrinter -Arguments @{ NewPrinterName = $NewName } -ErrorAction Stop
            if ([int]$r.ReturnValue -eq 0) { return [pscustomobject]@{ Success=$true; Message="Renamed '$Name' to '$NewName' (WMI)." } }
            return [pscustomobject]@{ Success=$false; Message="Rename failed. Rename-Printer: $first / WMI RenamePrinter returned $($r.ReturnValue)." }
        } catch {
            return [pscustomobject]@{ Success=$false; Message="Rename failed: $first" }
        }
    }
}

function Remove-PrinterSafe {
    <#
    .SYNOPSIS
        Delete a local queue, or remove a mapped printer from a specific user.
    .DESCRIPTION
        Remove-Printer on a '\\server\printer' name is PER-USER and runs in the
        CALLER's context - from SYSTEM it silently removes it from SYSTEM's own
        profile and leaves the user's alone. So connection removals are injected
        into the target user's session, with registry cleanup as the fallback
        (which takes effect at next logon).

        Registry cleanup deliberately leaves the driver and any local port in
        place - other users may still be using them.
    .OUTPUTS
        PSCustomObject { Success; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$User,
        [string]$Sid,
        [scriptblock]$OnPoll
    )
    if (-not $PSCmdlet.ShouldProcess($Name, 'Remove printer')) {
        return [pscustomobject]@{ Success=$false; Message='Skipped (WhatIf).' }
    }

    # --- local queue ---
    if ($Name -notmatch '^\\\\') {
        try {
            Remove-Printer -Name $Name -ErrorAction Stop
            return [pscustomobject]@{ Success=$true; Message="Removed local queue '$Name'." }
        } catch {
            return [pscustomobject]@{ Success=$false; Message="Could not remove '$Name': $($_.Exception.Message)" }
        }
    }

    # --- per-user connection ---
    $r = Invoke-ScriptInPrinterUserSession -User $User -OnPoll $OnPoll -Vars @{ Target = $Name } -ScriptBlock {
        $ok = $false; $err = ''
        try { Remove-Printer -Name $Target -ErrorAction Stop; $ok = $true }
        catch {
            $err = $_.Exception.Message
            try { (New-Object -ComObject WScript.Network).RemovePrinterConnection($Target, $true, $true); $ok = $true; $err = '' }
            catch { $err = "$err / $($_.Exception.Message)" }
        }
        [pscustomobject]@{ Removed = $ok; Error = $err } | ConvertTo-Json -Compress
    }

    if ($r.Success -and $r.Data.Removed) {
        return [pscustomobject]@{ Success=$true; Message="Removed '$Name' from $($r.User)'s session." }
    }

    # --- registry fallback ---
    if (-not $Sid -and $User) {
        try { $Sid = (New-Object Security.Principal.NTAccount $User).Translate([Security.Principal.SecurityIdentifier]).Value } catch {}
    }
    if (-not $Sid) {
        return [pscustomobject]@{ Success=$false; Message="Could not remove '$Name' in-session ($($r.Message)) and no SID was available for a registry fallback." }
    }

    $removed = $false
    try {
        $keyName = $Name.Replace('\', ',')
        $conn = [Microsoft.Win32.Registry]::Users.OpenSubKey("$Sid\Printers\Connections", $true)
        if ($conn) {
            try {
                if ($conn.GetSubKeyNames() -contains $keyName) { $conn.DeleteSubKeyTree($keyName); $removed = $true }
            } finally { $conn.Close() }
        }
        foreach ($p in @('Software\Microsoft\Windows NT\CurrentVersion\Devices',
                         'Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts')) {
            $k = [Microsoft.Win32.Registry]::Users.OpenSubKey("$Sid\$p", $true)
            if (-not $k) { continue }
            try { if ($k.GetValueNames() -contains $Name) { $k.DeleteValue($Name, $false) } } finally { $k.Close() }
        }
    } catch {
        return [pscustomobject]@{ Success=$false; Message="Registry fallback failed: $($_.Exception.Message)" }
    }

    if ($removed) {
        return [pscustomobject]@{ Success=$true; Message="Removed '$Name' from the user's registry. It disappears at their next sign-in (the running session keeps it cached)." }
    }
    return [pscustomobject]@{ Success=$false; Message="'$Name' was not found in that user's connections. In-session attempt said: $($r.Message)" }
}

function Add-PrinterConnectionForUser {
    <#
    .SYNOPSIS
        Map '\\server\printer' for a user, injected into their session, with a measured diagnosis on failure.
    .NOTES
        Add-PrinterConnection does NOT exist. The real cmdlet is
        Add-Printer -ConnectionName, and it accepts only -ConnectionName (no
        -ComputerName), running per-user in the caller's context.
    .OUTPUTS
        PSCustomObject { Success; Unc; Diagnosis; Message }
        Diagnosis is $null on success, otherwise Get-PrintConnectionDiagnosis output.
        Remedies must be chosen from Diagnosis.Category, never from Message text.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$ConnectionName,
        [string]$User,
        [scriptblock]$OnPoll
    )
    $ConnectionName = $ConnectionName.Trim()
    if ($ConnectionName -notmatch '^\\\\[^\\]+\\.+') {
        return [pscustomobject]@{ Success = $false; Unc = $ConnectionName; Diagnosis = $null; Message = "'$ConnectionName' is not a \\server\share path." }
    }
    if (-not $PSCmdlet.ShouldProcess($ConnectionName, "Add for $User")) {
        return [pscustomobject]@{ Success = $false; Unc = $ConnectionName; Diagnosis = $null; Message = 'Skipped (WhatIf).' }
    }

    $r = Invoke-ScriptInPrinterUserSession -User $User -OnPoll $OnPoll -TimeoutSec 120 -Vars @{ Target = $ConnectionName } -ScriptBlock {
        $ok = $false; $err = ''
        # ERROR_NETNAME_DELETED (64 / 0x80070040) is SOMETIMES transient - an SMB
        # session backing the connection went away between check and add - so one
        # retry after a short pause is cheap. It is not always transient, and a
        # second failure is reported as a failure rather than retried.
        $attempt = 0
        while ($attempt -lt 2 -and -not $ok) {
            $attempt++
            try { Add-Printer -ConnectionName $Target -ErrorAction Stop; $ok = $true; $err = '' }
            catch {
                $err = $_.Exception.Message
                try { (New-Object -ComObject WScript.Network).AddWindowsPrinterConnection($Target); $ok = $true; $err = '' }
                catch { $err = "$err / $($_.Exception.Message)" }
            }
            if (-not $ok -and $attempt -lt 2 -and ($err -match '0x80070040|no longer available|0x8007003B|network name')) {
                Start-Sleep -Seconds 3
            } elseif (-not $ok) { break }
        }
        # If the add failed, measure whether THIS user can open the exact printer on the
        # LIVE host, in this same session. The add runs as the user, so a console probe
        # would answer a different question. OpenPrinter2 with PRINTER_OPTION_NO_CACHE
        # (= 1), because plain OpenPrinter answered from cache for a host that did not
        # even resolve. A handle shows the connection answered for this identity - not
        # that a driver installed or a job printed.
        #
        # Signature built from an array, never a here-string: this module is inlined
        # into a here-string by the build, and a line starting '@ at column 0 would
        # terminate it.
        $rpcOk = $null; $rpcErr = 0
        if (-not $ok) {
            try {
                if (-not ('PrtNoCache.Spool' -as [type])) {
                    $psig = @(
                        '[StructLayout(LayoutKind.Sequential)] public struct PRINTER_OPTIONS { public uint cbSize; public uint dwFlags; }'
                        '[DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "OpenPrinter2W")]'
                        'public static extern bool OpenPrinter2(string pPrinterName, out System.IntPtr phPrinter, System.IntPtr pDefault, ref PRINTER_OPTIONS pOptions);'
                        '[DllImport("winspool.drv", SetLastError = true)]'
                        'public static extern bool ClosePrinter(System.IntPtr hPrinter);'
                    ) -join [Environment]::NewLine
                    Add-Type -Namespace 'PrtNoCache' -Name 'Spool' -MemberDefinition $psig -ErrorAction Stop
                }
                $opt = New-Object 'PrtNoCache.Spool+PRINTER_OPTIONS'
                $opt.cbSize = [uint32][System.Runtime.InteropServices.Marshal]::SizeOf([type]'PrtNoCache.Spool+PRINTER_OPTIONS')
                $opt.dwFlags = [uint32]1
                $hh = [IntPtr]::Zero
                $rpcOk = [PrtNoCache.Spool]::OpenPrinter2($Target, [ref]$hh, [IntPtr]::Zero, [ref]$opt)
                if (-not $rpcOk) { $rpcErr = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error() }
                if ($hh -ne [IntPtr]::Zero) { [void][PrtNoCache.Spool]::ClosePrinter($hh) }
            } catch { $rpcOk = $null }
        }
        [pscustomobject]@{ Added = $ok; Error = $err; Attempts = $attempt; RpcOk = $rpcOk; RpcErr = $rpcErr; RanAs = "$env:USERDOMAIN\$env:USERNAME" } | ConvertTo-Json -Compress
    }

    if ($r.Success -and $r.Data.Added) {
        $tries = if ($r.Data.Attempts -gt 1) { " (succeeded on attempt $($r.Data.Attempts))" } else { '' }
        return [pscustomobject]@{ Success = $true; Unc = $ConnectionName; Diagnosis = $null; Message = "Added '$ConnectionName' for $($r.User)$tries." }
    }

    # NO PARSEABLE RESULT is not the same as FAILURE, and must never be reported as
    # one. The add runs injected in the user's session; if the result cannot be read
    # back, the printer may well have been mapped anyway. Ask the only source that
    # actually knows - the user's own hive - before saying anything.
    if (-not $r.Data) {
        $mapped = $false
        try {
            $sid = ''
            foreach ($p in @(Get-UserPrinterProfiles)) {
                if ($p.User -ieq $r.User -or $p.User -ieq $User) { $sid = [string]$p.Sid; break }
            }
            if ($sid) {
                $m = Get-UserPrinterMap -Sid $sid
                foreach ($c in @($m.Connections)) {
                    if ([string]$c.Unc -ieq $ConnectionName) { $mapped = $true; break }
                }
            }
        } catch {}

        if ($mapped) {
            return [pscustomobject]@{ Success = $true; Unc = $ConnectionName; Diagnosis = $null
                Message = "Added '$ConnectionName' for $($r.User). (The in-session result could not be read back, so this was confirmed by re-reading the user's mapped printers.)" }
        }
        return [pscustomobject]@{ Success = $false; Unc = $ConnectionName; Diagnosis = $null
            Message = ("Could not confirm whether '$ConnectionName' was added for $($r.User). " +
                       "$($r.Message) Re-reading the user's mapped printers does not show it, so it most likely did not map - " +
                       'but this is an unread result rather than a reported failure. Press Rescan and check the list before retrying.') }
    }

    $detail = if ($r.Data.Error) { [string]$r.Data.Error } else { [string]$r.Message }
    $ranAs  = if ($r.Data.RanAs) { [string]$r.Data.RanAs } elseif ($r.User) { [string]$r.User } else { [string]$User }

    # Evidence, all bound to THIS exact UNC. Nothing below guesses from prose.
    $userProbe = $null
    if ($null -eq $r.Data.RpcOk) {
        $userProbe = New-PrintProbeResult -Probe 'Open as user' -State 'Skipped' -Detail 'The in-session open test did not run.' -Identity $ranAs -Target $ConnectionName
    } elseif ([bool]$r.Data.RpcOk) {
        $userProbe = New-PrintProbeResult -Probe 'Open as user' -State 'Ok' -Code '0' -Identity $ranAs -Target $ConnectionName `
            -Detail "$ranAs opened it on the live host (not from cache), in their own session."
    } else {
        $hn = ConvertTo-PrintHostKey $ConnectionName
        $userProbe = New-PrintProbeResult -Probe 'Open as user' -State 'Failed' -Code "$([int]$r.Data.RpcErr)" -Identity $ranAs -Target $ConnectionName `
            -Detail ("$ranAs could not open it on the live host: " + (Get-SmbErrorExplanation -Code ([int]$r.Data.RpcErr) -HostName $hn))
    }
    $co = Test-PrinterShareOpen -Unc $ConnectionName
    $consoleProbe = New-PrintProbeResult -Probe 'Open as console' -State $co.State -Detail $co.Message -Target $ConnectionName `
        -Code $(if ($null -ne $co.Win32Error) { "$($co.Win32Error)" } else { '' }) -Identity "$env:USERDOMAIN\$env:USERNAME"
    $path = Test-PrintRpcPath -ConnectionName $ConnectionName

    $diag = Get-PrintConnectionDiagnosis -ConnectionName $ConnectionName -AddError $detail -UserOpen $userProbe -ConsoleOpen $consoleProbe -Path $path -UserName $ranAs
    [pscustomobject]@{
        Success   = $false
        Unc       = $ConnectionName
        Diagnosis = $diag
        Message   = ("Could not add '$ConnectionName': $detail" + "`r`n`r`n" + (Format-PrintDiagnosis $diag))
    }
}

function Set-DefaultPrinterForUser {
    <#
    .SYNOPSIS
        Set a user's default printer so it actually takes effect AND sticks.
    .DESCRIPTION
        Two things are required and both are done here:

        1. The change must happen in the USER's session. SetDefaultPrinter() and
           rundll32 printui.dll,PrintUIEntry /y run from SYSTEM silently set
           SYSTEM's default, with no error - and SYSTEM has no Printers\Connections
           so the user's mapped printer is not even enumerable there.
        2. LegacyDefaultPrinterMode must be 1, or "Let Windows manage my default
           printer" silently overwrites the change on the next print.

        Registry-only fallback writes Device as '<name>,winspool,<port>', taking
        <port> from THAT user's own Devices value - never invented, never copied
        from another user (the Ne0x: token differs per user).
    .OUTPUTS
        PSCustomObject { Success; Message; Method }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$PrinterName,
        [string]$User,
        [string]$Sid,
        [scriptblock]$OnPoll
    )
    if (-not $PSCmdlet.ShouldProcess($PrinterName, "Set as default for $User")) {
        return [pscustomobject]@{ Success=$false; Message='Skipped (WhatIf).'; Method='' }
    }

    if (-not $Sid -and $User) {
        try { $Sid = (New-Object Security.Principal.NTAccount $User).Translate([Security.Principal.SecurityIdentifier]).Value } catch {}
    }

    # Always pin LegacyDefaultPrinterMode first, so even a successful in-session
    # set is not undone by Windows a minute later.
    $legacySet = $false
    if ($Sid) {
        try {
            $k = [Microsoft.Win32.Registry]::Users.OpenSubKey("$Sid\Software\Microsoft\Windows NT\CurrentVersion\Windows", $true)
            if ($k) {
                try { $k.SetValue('LegacyDefaultPrinterMode', 1, [Microsoft.Win32.RegistryValueKind]::DWord); $legacySet = $true }
                finally { $k.Close() }
            }
        } catch {}
    }

    $r = Invoke-ScriptInPrinterUserSession -User $User -OnPoll $OnPoll -Vars @{ Target = $PrinterName } -ScriptBlock {
        $ok = $false; $err = ''
        try { (New-Object -ComObject WScript.Network).SetDefaultPrinter($Target); $ok = $true }
        catch {
            $err = $_.Exception.Message
            try {
                $esc = $Target -replace "'", "''"
                $p = Get-CimInstance Win32_Printer -Filter "Name='$esc'" -ErrorAction Stop
                $res = Invoke-CimMethod -InputObject $p -MethodName SetDefaultPrinter -ErrorAction Stop
                if ([int]$res.ReturnValue -eq 0) { $ok = $true; $err = '' } else { $err = "$err / CIM returned $($res.ReturnValue)" }
            } catch { $err = "$err / $($_.Exception.Message)" }
        }
        [pscustomobject]@{ SetDefault = $ok; Error = $err } | ConvertTo-Json -Compress
    }

    if ($r.Success -and $r.Data.SetDefault) {
        $note = if ($legacySet) { ' LegacyDefaultPrinterMode set to 1 so Windows will not override it.' } else { ' NOTE: LegacyDefaultPrinterMode could not be set - Windows may override this.' }
        return [pscustomobject]@{ Success=$true; Method='in-session'; Message="Default printer for $($r.User) is now '$PrinterName'.$note" }
    }

    # --- registry fallback ---
    if (-not $Sid) {
        return [pscustomobject]@{ Success=$false; Method=''; Message="In-session set failed ($($r.Message)) and no SID was available for the registry fallback." }
    }

    $port = ''
    try {
        $k = [Microsoft.Win32.Registry]::Users.OpenSubKey("$Sid\Software\Microsoft\Windows NT\CurrentVersion\Devices")
        if ($k) {
            try {
                $raw = [string]$k.GetValue($PrinterName)
                if ($raw) { $parts = $raw -split ','; if ($parts.Count -ge 2) { $port = $parts[1] } }
            } finally { $k.Close() }
        }
    } catch {}

    if (-not $port) {
        return [pscustomobject]@{ Success=$false; Method=''; Message="'$PrinterName' is not in that user's Devices list, so its port token is unknown. Add the printer for the user first - writing an invented port produces error 0x00000709." }
    }

    try {
        $k = [Microsoft.Win32.Registry]::Users.OpenSubKey("$Sid\Software\Microsoft\Windows NT\CurrentVersion\Windows", $true)
        if (-not $k) { throw 'Could not open the user Windows key for writing.' }
        try {
            $k.SetValue('Device', "$PrinterName,winspool,$port", [Microsoft.Win32.RegistryValueKind]::String)
            $k.SetValue('LegacyDefaultPrinterMode', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
        } finally { $k.Close() }
    } catch {
        return [pscustomobject]@{ Success=$false; Method='registry'; Message="Registry fallback failed: $($_.Exception.Message)" }
    }

    [pscustomobject]@{
        Success = $true; Method = 'registry'
        Message = "Wrote the default for $User via the registry. Running applications cache the default, so this fully applies at their next sign-in."
    }
}

function Invoke-PrinterTestPage {
    <#
    .SYNOPSIS
        Print a Windows test page from the target user's session.
    .DESCRIPTION
        Injected so the job goes through the USER's queue - a SYSTEM-context test
        page would print from SYSTEM's spooler view, which cannot see mapped printers.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$PrinterName,
        [string]$User,
        [scriptblock]$OnPoll
    )
    if (-not $PSCmdlet.ShouldProcess($PrinterName, 'Print test page')) {
        return [pscustomobject]@{ Success=$false; Message='Skipped (WhatIf).' }
    }
    $r = Invoke-ScriptInPrinterUserSession -User $User -OnPoll $OnPoll -Vars @{ Target = $PrinterName } -ScriptBlock {
        $ok = $false; $err = ''
        try {
            $esc = $Target -replace "'", "''"
            $p = Get-CimInstance Win32_Printer -Filter "Name='$esc'" -ErrorAction Stop
            $res = Invoke-CimMethod -InputObject $p -MethodName PrintTestPage -ErrorAction Stop
            if ([int]$res.ReturnValue -eq 0) { $ok = $true } else { $err = "PrintTestPage returned $($res.ReturnValue)" }
        } catch { $err = $_.Exception.Message }
        [pscustomobject]@{ Printed = $ok; Error = $err } | ConvertTo-Json -Compress
    }
    if ($r.Success -and $r.Data.Printed) {
        return [pscustomobject]@{ Success=$true; Message="Test page sent to '$PrinterName' as $($r.User)." }
    }
    $detail = if ($r.Data -and $r.Data.Error) { $r.Data.Error } else { $r.Message }
    [pscustomobject]@{ Success=$false; Message="Test page failed: $detail" }
}

function Clear-PrinterQueue {
    <# .SYNOPSIS Cancel all jobs on a queue (Win32_Printer.CancelAllJobs, falling back to Remove-PrintJob). #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$PrinterName)

    if (-not $PSCmdlet.ShouldProcess($PrinterName, 'Cancel all print jobs')) {
        return [pscustomobject]@{ Success=$false; Message='Skipped (WhatIf).' }
    }
    try {
        $esc = $PrinterName -replace "'", "''"
        $p = Get-CimInstance Win32_Printer -Filter "Name='$esc'" -ErrorAction Stop
        $r = Invoke-CimMethod -InputObject $p -MethodName CancelAllJobs -ErrorAction Stop
        if ([int]$r.ReturnValue -eq 0) { return [pscustomobject]@{ Success=$true; Message="Cancelled all jobs on '$PrinterName'." } }
    } catch {}
    try {
        $n = 0
        foreach ($j in @(Get-PrintJob -PrinterName $PrinterName -ErrorAction Stop)) {
            try { Remove-PrintJob -PrinterName $PrinterName -ID $j.Id -ErrorAction Stop; $n++ } catch {}
        }
        return [pscustomobject]@{ Success=$true; Message="Removed $n job(s) from '$PrinterName'." }
    } catch {
        return [pscustomobject]@{ Success=$false; Message="Could not clear '$PrinterName': $($_.Exception.Message)" }
    }
}

function Restart-PrintSpooler {
    <# .SYNOPSIS Restart the Spooler service. #>
    [CmdletBinding(SupportsShouldProcess)]
    param([int]$TimeoutSec = 30)
    if (-not $PSCmdlet.ShouldProcess('Spooler', 'Restart service')) {
        return [pscustomobject]@{ Success=$false; Message='Skipped (WhatIf).' }
    }
    try {
        Restart-Service -Name Spooler -Force -ErrorAction Stop
        $svc = Get-Service -Name Spooler -ErrorAction Stop
        $svc.WaitForStatus('Running', (New-TimeSpan -Seconds $TimeoutSec))
        return [pscustomobject]@{ Success=$true; Message='Spooler restarted.' }
    } catch {
        return [pscustomobject]@{ Success=$false; Message="Spooler restart failed: $($_.Exception.Message)" }
    }
}

function Clear-SpoolFiles {
    <#
    .SYNOPSIS
        Stop the Spooler, delete orphaned .SPL/.SHD files, start it again.
    .DESCRIPTION
        Destructive: every queued job on the machine is discarded. The GUI gates
        this behind a confirmation naming the file count.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([string]$SpoolDirectory)

    if (-not $SpoolDirectory) { $SpoolDirectory = (Get-SpoolerHealth).SpoolDirectory }
    if (-not (Test-Path -LiteralPath $SpoolDirectory)) {
        return [pscustomobject]@{ Success=$false; Message="Spool directory not found: $SpoolDirectory" }
    }
    if (-not $PSCmdlet.ShouldProcess($SpoolDirectory, 'Stop spooler and delete all spool files')) {
        return [pscustomobject]@{ Success=$false; Message='Skipped (WhatIf).' }
    }

    $deleted = 0
    try {
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        Start-Sleep -Milliseconds 500
        foreach ($f in @(Get-ChildItem -LiteralPath $SpoolDirectory -Force -ErrorAction SilentlyContinue |
                         Where-Object { $_.Extension -in @('.SPL', '.SHD', '.spl', '.shd') })) {
            try { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop; $deleted++ } catch {}
        }
    } catch {
        try { Start-Service -Name Spooler -ErrorAction SilentlyContinue } catch {}
        return [pscustomobject]@{ Success=$false; Message="Could not clear the spool folder: $($_.Exception.Message)" }
    } finally {
        try { Start-Service -Name Spooler -ErrorAction SilentlyContinue } catch {}
    }
    [pscustomobject]@{ Success=$true; Message="Deleted $deleted spool file(s) from $SpoolDirectory and restarted the Spooler." }
}

# endregion


# ---------------------------------------------------------------------------
# region  Port / driver repair - IPP-class and WSD queues to Standard TCP/IP
#
# Ported from EPSON_PRINT_DRIVERS\Repoint-PrinterToIP.ps1 (which describes itself
# as vendor-agnostic) and Deploy-EpsonWF78xx.ps1, leaving the Epson driver
# catalogue behind. Field guide: EPSON_PRINT_DRIVERS\Printer_Port_and_Driver_Guide.md
#
# TWO RULES GOVERN EVERYTHING BELOW:
#
#   Rule 1 - WSD and IPP-class ports CANNOT be repointed. Set-Printer -PortName
#            works ONLY on a plain Standard TCP/IP queue that is keeping its
#            driver. Everything else must be REBUILT: new port -> temp queue ->
#            remove old -> rename back.
#
#   Rule 2 - the Microsoft IPP Class Driver negotiates its page-description
#            language over the IPP channel. Point it at a bare RAW:9100 socket
#            and there is no negotiation, so many devices print BLANK or
#            garbled. The port fix and the driver fix must happen TOGETHER;
#            fixing only the port is a trap.
# ---------------------------------------------------------------------------

function Get-QueueExact {
    <#
    .SYNOPSIS
        Exact-name queue lookup that is immune to [ ] * ? in printer names.
    .DESCRIPTION
        Get-Printer -Name treats square brackets as wildcards, and HP auto-names
        routinely contain them (e.g. 'HP Color LaserJet [E366EA]'). Looking such a
        queue up by -Name silently returns nothing, which during a rebuild would
        read as "the old queue is already gone" - so this filters client-side.
        Ported verbatim from Repoint-PrinterToIP.ps1.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)
    Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $Name }
}

function Test-PrintPortHostSet {
    <#
    .SYNOPSIS
        Probe many hosts for a print port CONCURRENTLY. Returns those that answer.
    .DESCRIPTION
        Used by the active IP discovery sweep. Probing an ARP table serially (up to
        ~64 hosts) would take 20s+; issuing every connect at once and waiting a
        single timeout window costs about as much as one probe.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$HostName,
        [int]$Port = 9100,
        [int]$TimeoutMs = 900
    )
    $clients = @{}
    $asyncs  = @{}
    foreach ($h in ($HostName | Select-Object -Unique)) {
        try {
            $c = New-Object System.Net.Sockets.TcpClient
            $clients[$h] = $c
            $asyncs[$h]  = $c.BeginConnect($h, $Port, $null, $null)
        } catch {
            if ($clients.ContainsKey($h)) { try { $clients[$h].Close() } catch {}; $clients.Remove($h) }
        }
    }
    Start-Sleep -Milliseconds $TimeoutMs
    $open = [System.Collections.Generic.List[string]]::new()
    foreach ($h in @($clients.Keys)) {
        try {
            if ($asyncs[$h].IsCompleted) {
                try { $clients[$h].EndConnect($asyncs[$h]) } catch {}
                if ($clients[$h].Connected) { $open.Add($h) }
            }
        } catch { }
        try { $clients[$h].Close() } catch {}
    }
    $open.ToArray()
}

function Get-WsdDeviceAddress {
    <#
    .SYNOPSIS
        Best-effort: pull IPv4 addresses out of the WSD device records in the registry.
    .DESCRIPTION
        A WSD port records no IP - the device is resolved by multicast UUID, which
        is precisely why a WSD queue dies when the printer moves. Windows does
        cache device metadata under the WSD provider keys though, and the device's
        web-UI URL is often in there. Purely a candidate generator: everything it
        returns is TCP-probed before being offered.
    #>
    [CmdletBinding()] param()

    $roots = @(
        'HKLM:\SYSTEM\CurrentControlSet\Enum\SWD\DAFWSDProvider',
        'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Monitors\WSD Port\Ports'
    )
    $found = [System.Collections.Generic.List[string]]::new()
    foreach ($root in $roots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        $keys = @()
        try   { $keys = @(Get-ChildItem -LiteralPath $root -Recurse -Depth 3 -ErrorAction SilentlyContinue) }
        catch { $keys = @(Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue) }
        $keys += @(Get-Item -LiteralPath $root -ErrorAction SilentlyContinue)
        foreach ($k in $keys) {
            $props = $null
            try { $props = Get-ItemProperty -LiteralPath $k.PSPath -ErrorAction SilentlyContinue } catch { continue }
            if (-not $props) { continue }
            foreach ($p in $props.PSObject.Properties) {
                if ($p.Name -like 'PS*') { continue }
                $v = "$($p.Value)"
                if (-not $v) { continue }
                foreach ($m in [regex]::Matches($v, '\b(\d{1,3}(?:\.\d{1,3}){3})\b')) {
                    $ip = $m.Groups[1].Value
                    if ($ip -eq '0.0.0.0' -or $ip -like '127.*' -or $ip -like '255.*') { continue }
                    if (-not $found.Contains($ip)) { $found.Add($ip) }
                }
            }
        }
    }
    $found.ToArray()
}

function Get-MdnsPrinterAddress {
    <#
    .SYNOPSIS
        Best-effort mDNS sweep for network printers (_pdl-datastream._tcp / _ipp._tcp).
    .DESCRIPTION
        Sends a multicast PTR query to 224.0.0.251:5353 and records who replies.
        Deliberately does NOT decode DNS name compression - it treats the
        responder's own source address as the candidate, which for a printer
        advertising its print service is exactly the address wanted. Everything it
        returns is TCP-probed before being offered, so a spurious responder is
        filtered out and a missed printer just means the tech types the IP.
    #>
    [CmdletBinding()]
    param([int]$WaitMs = 1500)

    $udp = $null
    $found = [System.Collections.Generic.List[string]]::new()
    try {
        $udp = New-Object System.Net.Sockets.UdpClient
        $udp.Client.SetSocketOption([System.Net.Sockets.SocketOptionLevel]::Socket,
                                    [System.Net.Sockets.SocketOptionName]::ReuseAddress, $true)
        $udp.Client.ReceiveTimeout = 300
        $udp.Client.Bind((New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)))
        $ep = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Parse('224.0.0.251'), 5353)

        foreach ($svc in @('_pdl-datastream._tcp.local', '_ipp._tcp.local')) {
            $q = [System.Collections.Generic.List[byte]]::new()
            $q.AddRange([byte[]]@(0,0, 0,0, 0,1, 0,0, 0,0, 0,0))   # one question
            foreach ($label in $svc.Split('.')) {
                $b = [System.Text.Encoding]::ASCII.GetBytes($label)
                $q.Add([byte]$b.Length); $q.AddRange($b)
            }
            $q.Add(0)
            $q.AddRange([byte[]]@(0,12))   # QTYPE PTR
            $q.AddRange([byte[]]@(0,1))    # QCLASS IN
            $bytes = $q.ToArray()
            [void]$udp.Send($bytes, $bytes.Length, $ep)
        }

        $deadline = (Get-Date).AddMilliseconds($WaitMs)
        while ((Get-Date) -lt $deadline) {
            $remote = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
            $data = $null
            try { $data = $udp.Receive([ref]$remote) } catch { continue }
            if (-not $data) { continue }
            $ip = $remote.Address.ToString()
            if ($ip -and $ip -notlike '127.*' -and -not $found.Contains($ip)) { $found.Add($ip) }
        }
    } catch {
    } finally {
        if ($udp) { try { $udp.Close() } catch {} }
    }
    $found.ToArray()
}

function Resolve-PrinterAddress {
    <#
    .SYNOPSIS
        Work out what IP a queue's device is actually at. Returns ranked candidates.
    .DESCRIPTION
        The scripts this engine came from have NO address discovery: -IPAddress is
        mandatory, and a WSD port hides the address completely, so there is nothing
        to drive a one-click repair from. This closes that gap.

        Tier 1 costs nothing and uses only data already in hand:
          - the port's own host address (a mapped queue leaks the real device IP
            through the print server's port name)
          - an IPP port URL
          - an IPv4 embedded in the port name
          - the Location field, very often the device web UI
            (seen on this fleet: 'http://192.0.2.47:3910/')
          - another local queue on the same driver already on a TCP/IP port

        Tier 2 (-Active) goes looking: WSD device records in the registry, an
        ARP-table sweep probed for port 9100, and an mDNS query.

        EVERY candidate is TCP-probed before being returned, and results are ranked
        reachable-first. Nothing is applied automatically - this exists to prefill a
        box the technician confirms, because a wrong IP would rebuild a working
        queue onto a dead port.
    .OUTPUTS
        PSCustomObject[] { Address; Source; Confidence; Reachable; OpenPorts }
    #>
    [CmdletBinding()]
    param(
        $Printer,
        [switch]$Active,
        [int]$ProbeTimeoutMs = 700,
        [int]$MaxActiveHosts = 64
    )

    $cands = [System.Collections.Generic.List[object]]::new()
    $seen  = @{}
    function Add-Cand {
        param([string]$Address, [string]$Source, [int]$Confidence)
        if (-not $Address) { return }
        $a = "$Address".Trim()
        if ($a -notmatch '^\d{1,3}(\.\d{1,3}){3}$' -and $a -notmatch '^[A-Za-z][\w\.\-]*$') { return }
        if ($a -in @('0.0.0.0', '127.0.0.1', 'localhost')) { return }
        $key = $a.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { return }
        $seen[$key] = $true
        $cands.Add([pscustomobject]@{ Address = $a; Source = $Source; Confidence = $Confidence })
    }

    # ---- Tier 1 ------------------------------------------------------------
    if ($Printer) {
        Add-Cand $Printer.HostAddress 'current port host address' 90
        $pn = "$($Printer.PortName)"
        if ($pn -match '://([^:/]+)')                   { Add-Cand $matches[1] 'IPP port URL' 85 }
        if ($pn -match '\b(\d{1,3}(?:\.\d{1,3}){3})\b') { Add-Cand $matches[1] 'port name' 80 }
        if ("$($Printer.Location)" -match '\b(\d{1,3}(?:\.\d{1,3}){3})\b') { Add-Cand $matches[1] 'Location field' 70 }
        if ("$($Printer.Comment)"  -match '\b(\d{1,3}(?:\.\d{1,3}){3})\b') { Add-Cand $matches[1] 'Comment field' 40 }

        # Another queue on the same driver with a working TCP/IP port is very often
        # the same physical device added a second time.
        try {
            foreach ($o in (Get-LocalPrinterInventory -IncludeVirtual)) {
                if ("$($o.Name)" -eq "$($Printer.Name)") { continue }
                if (-not $o.HostAddress) { continue }
                if ("$($o.DriverName)" -eq "$($Printer.DriverName)") {
                    Add-Cand $o.HostAddress "same driver on queue '$($o.Name)'" 60
                }
            }
        } catch {}
    }

    # ---- Tier 2 ------------------------------------------------------------
    if ($Active) {
        foreach ($ip in @(Get-WsdDeviceAddress)) { Add-Cand $ip 'WSD device record' 55 }

        $neighbours = @()
        try {
            $neighbours = @(Get-NetNeighbor -AddressFamily IPv4 -ErrorAction Stop |
                Where-Object { $_.State -in @('Reachable', 'Stale', 'Delay', 'Probe') } |
                Select-Object -ExpandProperty IPAddress)
        } catch {
            try {
                $neighbours = @((& arp -a 2>$null) |
                    ForEach-Object { if ($_ -match '\b(\d{1,3}(?:\.\d{1,3}){3})\b') { $matches[1] } })
            } catch {}
        }
        $neighbours = @($neighbours | Where-Object {
            $_ -and $_ -notlike '127.*' -and $_ -notlike '224.*' -and $_ -notlike '239.*' -and
            $_ -ne '255.255.255.255' -and $_ -notlike '*.255'
        } | Select-Object -Unique -First $MaxActiveHosts)

        if ($neighbours.Count) {
            foreach ($ip in @(Test-PrintPortHostSet -HostName $neighbours -Port 9100 -TimeoutMs 900)) {
                Add-Cand $ip 'ARP sweep, answers on 9100' 65
            }
        }

        foreach ($ip in @(Get-MdnsPrinterAddress)) { Add-Cand $ip 'mDNS printer service' 60 }
    }

    # ---- Probe and rank -----------------------------------------------------
    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($c in $cands) {
        $probe = Test-PrinterPortReachable -HostName $c.Address -TimeoutMs $ProbeTimeoutMs
        $out.Add([pscustomobject]@{
            Address    = $c.Address
            Source     = $c.Source
            Confidence = $c.Confidence
            Reachable  = $probe.Reachable
            OpenPorts  = $probe.OpenPorts
        })
    }
    @($out.ToArray() | Sort-Object -Property @{E='Reachable';Descending=$true}, @{E='Confidence';Descending=$true})
}

function Get-QueuePortKind {
    <#
    .SYNOPSIS
        Port kind for a queue's port, resolved against the real port object.
    .DESCRIPTION
        Get-PrinterConnectionInfo can only fall back to guessing from the port NAME
        when it has no port object, and a Standard TCP/IP port is very often named
        after the device rather than its IP (e.g. 'BRW30C9ABF69B57'). Guessing
        returns 'Other' for those, which would push a perfectly repointable queue
        down the destructive rebuild path. So look the port up first.

        Filters client-side rather than using -Name: port names can contain the
        same [ ] wildcard characters that break Get-Printer -Name.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$PortName)
    $port = $null
    try { $port = Get-PrinterPort -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $PortName } | Select-Object -First 1 } catch {}
    (Get-PrinterConnectionInfo -PortName $PortName -TcpPort $port).Type
}

function Get-QueueStateSnapshot {
    <#
    .SYNOPSIS
        Capture everything about a queue that a rebuild would otherwise destroy.
    .DESCRIPTION
        Taken BEFORE any change so it can be replayed afterwards. Includes the
        per-user default-printer flag, which the source scripts explicitly cannot
        restore ("default is per-user and CANNOT be set from SYSTEM") - this
        toolkit can, because Set-DefaultPrinterForUser injects into the user's
        own session.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    $q = Get-QueueExact $Name
    if (-not $q) { return $null }

    $defaultFor = [System.Collections.Generic.List[object]]::new()
    try {
        foreach ($m in (Get-AllUserPrinterMaps)) {
            if ("$($m.DefaultPrinter)" -ieq $Name) {
                $defaultFor.Add([pscustomobject]@{ User = $m.User; Sid = $m.Sid })
            }
        }
    } catch {}

    [pscustomobject]@{
        Name       = [string]$q.Name
        DriverName = [string]$q.DriverName
        PortName   = [string]$q.PortName
        PortKind   = (Get-QueuePortKind -PortName ([string]$q.PortName))
        Shared     = [bool]$q.Shared
        ShareName  = [string]$q.ShareName
        Comment    = [string]$q.Comment
        Location   = [string]$q.Location
        Published  = [bool]$q.Published
        DefaultFor = $defaultFor.ToArray()
    }
}

function New-PrinterTcpPort {
    <#
    .SYNOPSIS
        Ensure a Standard TCP/IP port exists for an address. Returns its name.
    .DESCRIPTION
        Naming follows the toolkit convention: IP_<address>, or IP_<address>_LPR.
        Collision rule carried over from Repoint-PrinterToIP.ps1: if a port of that
        name already exists but points somewhere ELSE, neither reuse nor overwrite
        it - create a new one with a short GUID suffix. Silently reusing a
        same-named port aimed at a different device is how a queue ends up printing
        on the wrong floor.
    .OUTPUTS
        PSCustomObject { Success; PortName; Created; Message }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$IPAddress,
        [ValidateSet('RAW', 'LPR')][string]$Protocol = 'RAW',
        [int]$PortNumber = 9100,
        [string]$LprQueueName = 'lp',
        [string]$PortName,
        [switch]$Preview
    )

    if (-not $PortName) {
        $PortName = if ($Protocol -eq 'LPR') { "IP_${IPAddress}_LPR" } else { "IP_$IPAddress" }
    }
    $existing = Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue
    if ($existing -and $existing.PrinterHostAddress -ne $IPAddress) {
        $PortName = "${PortName}_$([guid]::NewGuid().ToString('N').Substring(0,4))"
        $existing = $null
    }
    if ($existing) {
        return [pscustomobject]@{ Success=$true; PortName=$PortName; Created=$false; Message="Port '$PortName' already exists -> $($existing.PrinterHostAddress)." }
    }
    if ($Preview) {
        $what = if ($Protocol -eq 'LPR') { "${IPAddress} (LPR queue '$LprQueueName')" } else { "${IPAddress}:$PortNumber" }
        return [pscustomobject]@{ Success=$true; PortName=$PortName; Created=$false; Message="Would create $Protocol port '$PortName' -> $what." }
    }
    try {
        if ($Protocol -eq 'LPR') {
            Add-PrinterPort -Name $PortName -PrinterHostAddress $IPAddress -LprQueueName $LprQueueName -ErrorAction Stop
            return [pscustomobject]@{ Success=$true; PortName=$PortName; Created=$true; Message="Created LPR port '$PortName' -> ${IPAddress} (queue '$LprQueueName')." }
        }
        Add-PrinterPort -Name $PortName -PrinterHostAddress $IPAddress -PortNumber $PortNumber -ErrorAction Stop
        return [pscustomobject]@{ Success=$true; PortName=$PortName; Created=$true; Message="Created RAW port '$PortName' -> ${IPAddress}:$PortNumber." }
    } catch {
        return [pscustomobject]@{ Success=$false; PortName=$PortName; Created=$false; Message="Could not create port '$PortName': $($_.Exception.Message)" }
    }
}

function Get-PrinterInfCandidate {
    <#
    .SYNOPSIS
        Find the printer INFs in a folder and read the DRIVER NAMES each one declares.
    .DESCRIPTION
        Install-PrinterDriverFromInf needs the exact model name the INF declares -
        'Canon G6000 series', not 'Canon G6065' and not the INF's filename - and
        getting it wrong fails with an unhelpful error. Typing it blind from a
        vendor folder full of INFs is guesswork, so this reads them instead.

        How a printer INF actually declares its models:

          [Version]      Class=Printer            <- what makes it a printer INF
          [Manufacturer] %CanonName%=Canon,NTamd64
          [Canon.NTamd64] %G6000%=InstallSec,HWID <- one line per model
          [Strings]      G6000="Canon G6000 series"

        So the model list is [Manufacturer] -> decorated models section -> the
        left-hand side of each line, with %tokens% resolved from [Strings].

        Sections for THIS architecture are preferred; the undecorated section is
        used as a fallback, because plenty of vendor INFs still ship one.
    .PARAMETER Match
        A printer or model name to rank against, e.g. 'Canon G6065'. Candidates
        are scored so the closest is first - exact, then all words present, then
        some. Nothing is filtered out by it; a bad guess must never hide the
        right answer.
    .OUTPUTS
        PSCustomObject[] { DriverName; InfPath; InfName; Manufacturer; Section; Score }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Match,
        [int]$MaxFiles = 400
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return @()
    }

    $files = @()
    try {
        if ((Get-Item -LiteralPath $Path).PSIsContainer) {
            $files = @(Get-ChildItem -LiteralPath $Path -Filter '*.inf' -Recurse -File -ErrorAction SilentlyContinue |
                       Select-Object -First $MaxFiles)
        } else {
            $files = @(Get-Item -LiteralPath $Path)
        }
    } catch {}

    $arch = if ([Environment]::Is64BitOperatingSystem) { 'NTamd64' } else { 'NTx86' }
    $out = [System.Collections.Generic.List[object]]::new()

    foreach ($f in $files) {
        $text = $null
        try { $text = Get-Content -LiteralPath $f.FullName -Raw -ErrorAction Stop } catch { continue }
        if (-not $text) { continue }

        # --- split into sections -------------------------------------------
        $sections = @{}
        $cur = ''
        foreach ($line in ($text -split "`r?`n")) {
            $t = $line.Trim()
            if (-not $t -or $t.StartsWith(';')) { continue }
            if ($t -match '^\[(.+?)\]') {
                $cur = $matches[1].Trim()
                if (-not $sections.ContainsKey($cur)) { $sections[$cur] = [System.Collections.Generic.List[string]]::new() }
                continue
            }
            if ($cur) { $sections[$cur].Add($t) }
        }
        if (-not $sections.Count) { continue }

        # --- printer INFs only ---------------------------------------------
        $isPrinter = $false
        foreach ($k in $sections.Keys) {
            if ($k -ieq 'Version') {
                foreach ($v in $sections[$k]) {
                    if ($v -imatch '^\s*Class\s*=\s*Printer\s*$') { $isPrinter = $true }
                    if ($v -imatch '4d36e979-e325-11ce-bfc1-08002be10318') { $isPrinter = $true }
                }
            }
        }
        if (-not $isPrinter) { continue }

        # --- [Strings] for %token% resolution -------------------------------
        $strings = @{}
        foreach ($k in $sections.Keys) {
            if ($k -imatch '^Strings(\..+)?$') {
                foreach ($v in $sections[$k]) {
                    if ($v -match '^\s*([^=]+?)\s*=\s*(.+?)\s*$') {
                        $strings[$matches[1].Trim()] = $matches[2].Trim().Trim('"')
                    }
                }
            }
        }
        $resolve = {
            param([string]$s)
            $r = $s.Trim().Trim('"')
            if ($r -match '^%(.+)%$') {
                $key = $matches[1]
                foreach ($sk in $strings.Keys) { if ($sk -ieq $key) { return $strings[$sk] } }
                return $key
            }
            return $r
        }

        # --- [Manufacturer] -> model sections -------------------------------
        $mfgSection = $null
        foreach ($k in $sections.Keys) { if ($k -ieq 'Manufacturer') { $mfgSection = $sections[$k] } }
        if (-not $mfgSection) { continue }

        foreach ($line in $mfgSection) {
            if ($line -notmatch '^\s*(.+?)\s*=\s*(.+?)\s*$') { continue }
            # Capture BOTH groups before calling anything. $matches is global and
            # $resolve runs its own -match, which replaces it - so reading
            # $matches[2] after the call returned the token match, not the RHS.
            # That silently emptied every INF whose [Manufacturer] uses %tokens%,
            # which is most of them, including Windows' own ntprint.inf.
            $lhsRaw = $matches[1]
            $rhs    = $matches[2]
            $mfgName = & $resolve $lhsRaw
            $parts = @($rhs -split ',' | ForEach-Object { $_.Trim() })
            $baseSection = $parts[0]
            if (-not $baseSection) { continue }

            # Prefer a section decorated for this architecture, else undecorated.
            $wanted = [System.Collections.Generic.List[string]]::new()
            for ($i = 1; $i -lt $parts.Count; $i++) {
                if ($parts[$i] -imatch [regex]::Escape($arch)) { $wanted.Add("$baseSection.$($parts[$i])") }
            }
            $wanted.Add($baseSection)

            foreach ($secName in $wanted) {
                $sec = $null
                foreach ($k in $sections.Keys) { if ($k -ieq $secName) { $sec = $sections[$k]; break } }
                if (-not $sec) { continue }
                foreach ($ml in $sec) {
                    if ($ml -notmatch '^\s*(.+?)\s*=\s*(.+?)\s*$') { continue }
                    $modelRaw = $matches[1]
                    $model = & $resolve $modelRaw
                    if (-not $model -or $model.Length -lt 2) { continue }
                    # Class GUIDs appear on the left of some entries (ntprint.inf
                    # does this). They are not model names and would only clutter
                    # the picker.
                    if ($model -match '^\{?[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}?$') { continue }
                    $out.Add([pscustomobject]@{
                        DriverName   = $model
                        InfPath      = $f.FullName
                        InfName      = $f.Name
                        Manufacturer = $mfgName
                        Section      = $secName
                        Score        = 0
                    })
                }
                break   # first section that exists wins, do not double-count
            }
        }
    }

    # --- de-duplicate: the same model often appears in several INFs ---------
    $seen = @{}
    $uniq = [System.Collections.Generic.List[object]]::new()
    foreach ($c in $out) {
        $key = ($c.DriverName + '|' + $c.InfPath).ToLower()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $uniq.Add($c)
    }

    # --- rank against the printer we are trying to match -------------------
    if ($Match) {
        $m = $Match.ToLower()
        $words = @($m -split '[^a-z0-9]+' | Where-Object { $_.Length -gt 1 })
        foreach ($c in $uniq) {
            $d = $c.DriverName.ToLower()
            $score = 0
            if ($d -eq $m) { $score = 1000 }
            elseif ($d -like "*$m*" -or $m -like "*$d*") { $score = 500 }
            else {
                foreach ($w in $words) { if ($d -like "*$w*") { $score += 100 } }
            }
            # a shorter name that still matches is usually the family driver
            if ($score -gt 0) { $score += [Math]::Max(0, 60 - $c.DriverName.Length) }
            $c.Score = $score
        }
        return @($uniq | Sort-Object -Property @{ Expression = 'Score'; Descending = $true }, 'DriverName')
    }
    return @($uniq | Sort-Object 'DriverName')
}

function Install-PrinterDriverFromInf {
    <#
    .SYNOPSIS
        Stage a driver package into the driver store and register it with the spooler.
    .DESCRIPTION
        The generic Windows two-step, ported from Install-EpsonDriver (vendor-neutral
        apart from its name):
          1. pnputil /add-driver <inf> /install   - stage into the driver store
          2. Add-PrinterDriver -Name <model>      - register with the spooler,
             retrying with an explicit -InfPath

        Point this at the INF inside its ORIGINAL extracted folder, never at an INF
        copied out on its own: printer INFs carry CopyINF dependencies on sibling
        files and staging fails without them.

        -RestartSpooler first when a package has failed once: vendors register
        language-monitor DLLs that spoolsv.exe holds open, which aborts the install
        part-way through.
    .OUTPUTS
        PSCustomObject { Success; Message; DriverName }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$InfPath,
        [Parameter(Mandatory)][string]$DriverName,
        [switch]$RestartSpooler,
        [switch]$Preview
    )

    if (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue) {
        return [pscustomobject]@{ Success=$true; DriverName=$DriverName; Message="Driver '$DriverName' is already registered with the spooler." }
    }
    if (-not (Test-Path -LiteralPath $InfPath)) {
        return [pscustomobject]@{ Success=$false; DriverName=$DriverName; Message="INF not found: $InfPath" }
    }
    if ($Preview) {
        return [pscustomobject]@{ Success=$true; DriverName=$DriverName; Message="Would stage '$InfPath' with pnputil and register driver '$DriverName'." }
    }

    $log = [System.Collections.Generic.List[string]]::new()
    if ($RestartSpooler) { $log.Add((Restart-PrintSpooler).Message) }

    try {
        $null = & pnputil.exe /add-driver "$InfPath" /install 2>&1
        $code = $LASTEXITCODE
        # 3010 = success, reboot required. Other non-zero codes are often benign
        # (already staged), so continue and let registration be the judge.
        if ($code -eq 0 -or $code -eq 3010) { $log.Add("pnputil staged the package (exit $code).") }
        else { $log.Add("pnputil exit $code - continuing; the package may already be staged.") }
    } catch {
        $log.Add("pnputil failed: $($_.Exception.Message)")
    }

    if (-not (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue)) {
        try {
            Add-PrinterDriver -Name $DriverName -ErrorAction Stop
        } catch {
            try {
                Add-PrinterDriver -Name $DriverName -InfPath $InfPath -ErrorAction Stop
            } catch {
                $log.Add("Add-PrinterDriver failed: $($_.Exception.Message)")
                return [pscustomobject]@{
                    Success=$false; DriverName=$DriverName
                    Message=("Driver '$DriverName' could not be registered. " + ($log -join ' ') + " Check the name matches one the INF actually declares.")
                }
            }
        }
    }
    if (-not (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue)) {
        return [pscustomobject]@{ Success=$false; DriverName=$DriverName; Message=("Driver '$DriverName' is still not present after install. " + ($log -join ' ')) }
    }
    [pscustomobject]@{ Success=$true; DriverName=$DriverName; Message=("Registered driver '$DriverName'. " + ($log -join ' ')).Trim() }
}

function Test-RepointResult {
    <#
    .SYNOPSIS
        PASS/FAIL check table for a completed repoint or rebuild.
    .OUTPUTS
        PSCustomObject { AllPassed; Checks[] }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$PortName,
        [Parameter(Mandatory)][string]$IPAddress,
        [string]$QueueName,
        [ValidateSet('RAW', 'LPR')][string]$Protocol = 'RAW',
        [int]$ProbeTimeoutMs = 1500
    )
    $checks = [System.Collections.Generic.List[object]]::new()
    function Add-Chk([string]$n, [bool]$ok, [string]$d) {
        $checks.Add([pscustomobject]@{ Check=$n; Result=$(if($ok){'PASS'}else{'FAIL'}); Detail=$d })
    }

    $port = Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue
    Add-Chk 'New port exists' ([bool]$port) $PortName
    if ($port) {
        Add-Chk 'Port host address' ($port.PrinterHostAddress -eq $IPAddress) "$($port.PrinterHostAddress) (expected $IPAddress)"
        # MSFT_TcpIpPrinterPort reports Protocol as 1=RAW / 2=LPR, but some builds
        # surface the string - accept either.
        $want = if ($Protocol -eq 'LPR') { 2 } else { 1 }
        Add-Chk "Port protocol is $Protocol" ("$($port.Protocol)" -eq "$want" -or "$($port.Protocol)" -eq $Protocol) "Protocol=$($port.Protocol) (1=RAW,2=LPR)"
    }
    if ($QueueName) {
        $fq = Get-QueueExact $QueueName
        if ($fq) {
            Add-Chk "Queue on the new port" ($fq.PortName -eq $PortName) "PortName=$($fq.PortName)"
            $kind = (Get-QueuePortKind -PortName ([string]$fq.PortName))
            Add-Chk "Queue no longer WSD/IPP" ($kind -notin @('WSD','IPP')) "PortKind=$kind"
            Add-Chk "Driver is not the IPP class driver" (-not (Test-IsIppClassDriver -DriverName ([string]$fq.DriverName))) "Driver=$($fq.DriverName)"
        } else {
            Add-Chk "Queue '$QueueName' present" $false 'not found'
        }
    }
    $probe = Test-PrinterPortReachable -HostName $IPAddress -TimeoutMs $ProbeTimeoutMs
    Add-Chk "Device $IPAddress reachable" $probe.Reachable $probe.Message

    $arr = $checks.ToArray()
    [pscustomobject]@{
        AllPassed = (@($arr | Where-Object { $_.Result -eq 'FAIL' }).Count -eq 0)
        Checks    = $arr
    }
}

function Convert-PrinterQueue {
    <#
    .SYNOPSIS
        Repoint or rebuild a queue onto a Standard TCP/IP port, optionally swapping
        in a real driver. The core of the IPP/WSD fix.
    .DESCRIPTION
        Chooses between a simple repoint and a full rebuild using the STRICTER of
        the two predicates found in the source project:

            rebuild required = the port is WSD or IPP
                               OR the driver is changing
                               OR the current driver is the generic IPP class driver

        Repoint-PrinterToIP.ps1 omits that last clause, so an IPP-class queue that
        already sits on a plain TCP/IP port takes its repoint branch and is merely
        warned about - leaving the blank-page driver in place. Deploy-EpsonWF78xx.ps1
        includes it. This uses the safer one.

        Preserved across a rebuild: name, driver, sharing, comment, location, and -
        unlike the source scripts - the per-user default-printer flag.

        NOT preserved, and said out loud rather than silently dropped: printer ACLs,
        printing defaults (DEVMODE), installed options, port SNMP settings, per-user
        DevModePerUser preferences, and any queued jobs.
    .PARAMETER Preview
        Plan only. Returns the exact steps and changes nothing.
    .OUTPUTS
        PSCustomObject { Success; Preview; Message; FinalName; PortName; Rebuilt;
                         Steps[]; Checks[]; Snapshot }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$IPAddress,
        [ValidateSet('RAW', 'LPR')][string]$Protocol = 'RAW',
        [int]$PortNumber = 9100,
        [string]$LprQueueName = 'lp',
        [string]$DriverName,
        [string]$NewName,
        [string]$PortName,
        [switch]$Force,
        [switch]$Preview,
        [switch]$SkipDefaultRestore,
        [scriptblock]$OnPoll
    )

    $steps = [System.Collections.Generic.List[object]]::new()
    function Step([string]$Tag, [string]$Text) {
        $steps.Add([pscustomobject]@{ Tag=$Tag; Text=$Text })
        if ($OnPoll) { try { & $OnPoll } catch {} }
    }
    function Fail([string]$Text) {
        Step 'FAIL' $Text
        [pscustomobject]@{ Success=$false; Preview=[bool]$Preview; Message=$Text; FinalName=$null
                           PortName=$PortName; Rebuilt=$false; Steps=$steps.ToArray(); Checks=@(); Snapshot=$null }
    }

    # ---- Guards -------------------------------------------------------------
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $isAdmin = $false
    try { $isAdmin = ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } catch {}
    if (-not $isAdmin -and -not $Preview) {
        return Fail 'Creating ports and rebuilding queues needs elevation. Run as Administrator, or as SYSTEM from Backstage / Atera.'
    }

    $q = Get-QueueExact $Name
    if (-not $q) { return Fail "No queue named '$Name' exists on this machine." }

    $snapshot     = Get-QueueStateSnapshot -Name $Name
    $targetName   = if ($NewName) { $NewName } else { $Name }
    $driver       = if ($DriverName) { $DriverName } else { [string]$q.DriverName }
    $kind         = (Get-QueuePortKind -PortName ([string]$q.PortName))
    $driverChange = ($driver -ne [string]$q.DriverName)

    Step 'INFO' "Queue '$Name' is currently $($q.DriverName) on $($q.PortName) [$kind]."

    # ---- Rule 2 ------------------------------------------------------------
    if ((Test-IsIppClassDriver -DriverName $driver) -and $Protocol -eq 'RAW') {
        Step 'WARN' "This keeps the generic '$driver', which negotiates over IPP and often prints BLANK or garbled on a RAW:9100 port. Install the real vendor driver and pass it as -DriverName, or use -Protocol LPR / re-add over IPP instead."
    }

    # ---- The driver has to exist -------------------------------------------
    if (-not (Get-PrinterDriver -Name $driver -ErrorAction SilentlyContinue)) {
        return Fail "Driver '$driver' is not installed, so the queue cannot be built with it. Install it first (Install from INF...), or leave the driver unchanged."
    }

    # ---- Reachability gate, BEFORE any change ------------------------------
    $reachPort = if ($Protocol -eq 'LPR') { 515 } else { $PortNumber }
    $probe = Test-TcpPortFast -HostName $IPAddress -Port $reachPort -TimeoutMs 1500
    if ($probe.Open) {
        Step 'OK' "Device answers at ${IPAddress}:$reachPort ($($probe.Ms) ms)."
    } elseif ($Preview) {
        # A preview changes nothing, so an unreachable device must not stop the tech
        # seeing the plan - that is exactly when they most want to read it.
        Step 'WARN' "No response on ${IPAddress}:$reachPort. Previewing anyway; a real run would stop here unless Force is set."
    } elseif ($Force) {
        Step 'WARN' "No response on ${IPAddress}:$reachPort, continuing anyway because Force was set."
    } else {
        return Fail "No response on ${IPAddress}:$reachPort. Confirm the printer is on and that protocol is enabled - some devices ship with RAW/9100 disabled and only LPD/IPP on, so try LPR - or set Force. Nothing was changed."
    }

    # ---- Port ---------------------------------------------------------------
    $portRes = New-PrinterTcpPort -IPAddress $IPAddress -Protocol $Protocol -PortNumber $PortNumber `
                                  -LprQueueName $LprQueueName -PortName $PortName -Preview:$Preview
    if (-not $portRes.Success) { return Fail $portRes.Message }
    $PortName = $portRes.PortName
    Step $(if ($portRes.Created) { 'OK' } else { 'INFO' }) $portRes.Message

    # ---- Repoint or rebuild? -----------------------------------------------
    # Repoint is a WHITELIST, not a blacklist: Set-Printer -PortName is only safe on
    # a queue that is ALREADY a plain Standard TCP/IP port and is keeping its driver.
    # WSD and IPP ports cannot be repointed at all, and USB / LPT / COM / unknown
    # kinds are rebuilt too - moving a local-bus queue onto a network port by
    # editing PortName is not a supported transition. Stating this as "rebuild if
    # WSD or IPP" would silently send a USB queue down the repoint path.
    # NOTE the vocabulary difference: the source toolkit's Get-PortKind calls a
    # Standard TCP/IP port 'TCPIP', whereas Get-PrinterConnectionInfo in this module
    # calls it 'Network'. Accept both, or the whitelist never matches and every
    # queue takes the destructive rebuild path even when a repoint would do.
    $repointable  = @('Network', 'TCPIP')
    $currentIsIpp = (Test-IsIppClassDriver -DriverName ([string]$q.DriverName))
    $mustRebuild  = ($kind -notin $repointable) -or $driverChange -or $currentIsIpp
    if ($mustRebuild) {
        $why = @()
        if ($kind -in @('WSD','IPP'))            { $why += "$kind ports cannot be repointed" }
        elseif ($kind -notin $repointable)       { $why += "the queue is on a $kind port, which cannot be repointed onto TCP/IP" }
        if ($driverChange)                        { $why += "the driver is changing to '$driver'" }
        if ($currentIsIpp)                        { $why += 'the current driver is the generic IPP class driver' }
        Step 'INFO' ('Rebuild required: ' + ($why -join '; ') + '.')
    } else {
        Step 'INFO' 'Already a plain Standard TCP/IP queue and keeping its driver - a simple repoint is enough.'
    }

    if ($mustRebuild -and $snapshot -and @($snapshot.DefaultFor).Count) {
        Step 'INFO' ('Currently the default printer for: ' + ((@($snapshot.DefaultFor) | ForEach-Object { $_.User }) -join ', ') + '.')
    }

    # ---- Preview stops here --------------------------------------------------
    if ($Preview) {
        if ($mustRebuild) {
            Step 'PLAN' "Create temporary queue '$targetName (repoint-temp)' on '$PortName' with driver '$driver'."
            Step 'PLAN' "Remove the old queue '$Name'."
            Step 'PLAN' "Rename the temporary queue to '$targetName'."
            Step 'PLAN' 'Restore comment, location and sharing.'
            if (-not $SkipDefaultRestore -and $snapshot -and @($snapshot.DefaultFor).Count) {
                Step 'PLAN' ('Restore the default-printer flag for ' + ((@($snapshot.DefaultFor) | ForEach-Object { $_.User }) -join ', ') + '.')
            }
            Step 'NOTE' 'A rebuild does NOT carry over printer permissions, printing defaults (DEVMODE), installed options, port SNMP settings, or queued jobs.'
        } else {
            Step 'PLAN' "Repoint '$Name' onto '$PortName'."
        }
        return [pscustomobject]@{
            Success=$true; Preview=$true; Message='Preview only - nothing was changed.'
            FinalName=$targetName; PortName=$PortName; Rebuilt=$mustRebuild
            Steps=$steps.ToArray(); Checks=@(); Snapshot=$snapshot
        }
    }

    # ---- Execute --------------------------------------------------------------
    $finalName = $null
    try {
        if (-not $mustRebuild) {
            $q | Set-Printer -PortName $PortName -ErrorAction Stop
            $finalName = $Name
            Step 'OK' "Repointed '$Name' onto '$PortName'."
        } else {
            $tmp = "$targetName (repoint-temp)"
            $stale = Get-QueueExact $tmp
            if ($stale) {
                $stale | Remove-Printer -ErrorAction SilentlyContinue
                Step 'INFO' "Cleared a leftover '$tmp' from an earlier attempt."
            }

            Add-Printer -Name $tmp -DriverName $driver -PortName $PortName -ErrorAction Stop
            Step 'OK' "Created replacement queue '$tmp' on '$PortName' ($driver)."

            $q | Remove-Printer -ErrorAction Stop
            Step 'OK' "Removed the old queue '$Name' (was $($snapshot.DriverName) on $($snapshot.PortName))."

            if (Get-QueueExact $targetName) {
                Step 'WARN' "A queue named '$targetName' still exists, so the rebuilt queue keeps the name '$tmp'. Rename it by hand once the clash is resolved."
                $finalName = $tmp
            } else {
                (Get-QueueExact $tmp) | Rename-Printer -NewName $targetName -ErrorAction Stop
                $finalName = $targetName
                Step 'OK' "Renamed the replacement queue to '$targetName'."
            }

            $fin = Get-QueueExact $finalName
            $set = @{}
            if ($snapshot.Comment)  { $set.Comment  = $snapshot.Comment }
            if ($snapshot.Location) { $set.Location = $snapshot.Location }
            if ($fin -and $set.Count) {
                $fin | Set-Printer @set -ErrorAction SilentlyContinue
                Step 'OK' 'Restored comment and location.'
            }
            if ($fin -and $snapshot.Shared -and $snapshot.ShareName) {
                try {
                    $fin | Set-Printer -Shared $true -ShareName $snapshot.ShareName -ErrorAction Stop
                    Step 'OK' "Re-shared as '$($snapshot.ShareName)'."
                } catch {
                    Step 'WARN' "Could not re-share '$finalName': $($_.Exception.Message)"
                }
            }
        }
    } catch {
        Step 'FAIL' "Failed on '$Name': $($_.Exception.Message)"
        # Say precisely what state the machine is in - the dangerous window is
        # between removing the old queue and renaming the replacement.
        $still = Get-QueueExact $Name
        $tmpQ  = Get-QueueExact "$targetName (repoint-temp)"
        if (-not $still -and $tmpQ) {
            Step 'WARN' "The old queue is gone and the replacement is still called '$targetName (repoint-temp)'. Rename it to '$targetName' to finish, or just run this again."
        } elseif ($still) {
            Step 'INFO' "The original queue '$Name' is still in place - nothing was lost."
        }
        return [pscustomobject]@{
            Success=$false; Preview=$false; Message="Failed on '$Name': $($_.Exception.Message)"
            FinalName=$(if ($tmpQ) { $tmpQ.Name } else { $null }); PortName=$PortName; Rebuilt=$mustRebuild
            Steps=$steps.ToArray(); Checks=@(); Snapshot=$snapshot
        }
    }

    # ---- Restore per-user defaults ------------------------------------------
    # The source scripts cannot do this and tell the tech to fix it by hand.
    if (-not $SkipDefaultRestore -and $mustRebuild -and $snapshot -and @($snapshot.DefaultFor).Count) {
        foreach ($u in @($snapshot.DefaultFor)) {
            $r = Set-DefaultPrinterForUser -PrinterName $finalName -User $u.User -Sid $u.Sid -OnPoll $OnPoll
            Step $(if ($r.Success) { 'OK' } else { 'WARN' }) ("Default printer for $($u.User): " + $r.Message)
        }
    }

    # ---- Verify ---------------------------------------------------------------
    $verify = Test-RepointResult -PortName $PortName -IPAddress $IPAddress -QueueName $finalName -Protocol $Protocol
    foreach ($c in @($verify.Checks)) {
        Step $(if ($c.Result -eq 'PASS') { 'OK' } else { 'FAIL' }) ("$($c.Check): $($c.Result) - $($c.Detail)")
    }

    $msg = if ($verify.AllPassed) {
        "'$finalName' is now on $Protocol port '$PortName' -> $IPAddress. All $(@($verify.Checks).Count) checks passed."
    } else {
        "'$finalName' was changed but $(@($verify.Checks | Where-Object { $_.Result -eq 'FAIL' }).Count) verification check(s) FAILED - review the steps."
    }
    Step $(if ($verify.AllPassed) { 'OK' } else { 'WARN' }) $msg

    [pscustomobject]@{
        Success   = $verify.AllPassed
        Preview   = $false
        Message   = $msg
        FinalName = $finalName
        PortName  = $PortName
        Rebuilt   = $mustRebuild
        Steps     = $steps.ToArray()
        Checks    = $verify.Checks
        Snapshot  = $snapshot
    }
}

# endregion


# ---------------------------------------------------------------------------
# region  USB-attached printers
#
# A USB print queue keeps working in the spooler whether or not the printer is
# actually plugged in and switched on, and the status fields lie about it.
# Measured on a real queue with the printer disconnected:
#
#     Get-Printer  -> PrinterStatus : Normal        <- false
#     Win32_Printer-> WorkOffline   : True          <- true
#     PnP          -> no USBPRINT device present    <- true
#     Get-PrinterPort -> USB001 still listed        <- ports outlive devices
#
# 'USB001' is described by the spooler as a "Virtual printer port for USB"
# owned by the Dynamic Print Monitor, so the PORT EXISTING PROVES NOTHING.
# Presence has to come from PnP.
# ---------------------------------------------------------------------------

function Get-UsbPrinterDevice {
    <#
    .SYNOPSIS
        USB print devices currently attached, with the spooler port each owns.
    .DESCRIPTION
        Two sources, best first:

          1. HKLM\SYSTEM\CurrentControlSet\Enum\USBPRINT\...\Device Parameters
             carries PortName (e.g. 'USB001'), which is the authoritative link
             between a physical device and a spooler port. Note the whole
             USBPRINT key is absent on a machine where no USB printer has ever
             been attached, and entries can persist for a device that is no
             longer plugged in - so the registry says "was seen", not "is here".

          2. Get-PnpDevice for USBPRINT\ and DOT4\ (multifunction) instances,
             which reports Present/Status for what is attached RIGHT NOW.

        The two are joined: Present comes from PnP, PortName from the registry.
    .OUTPUTS
        PSCustomObject[] { InstanceId; FriendlyName; PortName; Present; Status }
    #>
    [CmdletBinding()] param()

    # --- registry: instance id (lowercased) -> port name --------------------
    $portByInstance = @{}
    $root = 'HKLM:\SYSTEM\CurrentControlSet\Enum\USBPRINT'
    if (Test-Path -LiteralPath $root) {
        foreach ($k in @(Get-ChildItem -LiteralPath $root -Recurse -Depth 2 -ErrorAction SilentlyContinue)) {
            $dp = Join-Path $k.PSPath 'Device Parameters'
            if (-not (Test-Path -LiteralPath $dp)) { continue }
            $pn = $null
            try { $pn = (Get-ItemProperty -LiteralPath $dp -ErrorAction SilentlyContinue).PortName } catch {}
            if (-not $pn) { continue }
            # PSPath tail after the Enum\ root IS the instance id.
            $inst = ($k.Name -replace '^HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Enum\\', '')
            $portByInstance[$inst.ToLowerInvariant()] = [string]$pn
        }
    }

    $out = [System.Collections.Generic.List[object]]::new()
    $devices = @()
    try {
        $devices = @(Get-PnpDevice -ErrorAction Stop |
            Where-Object { $_.InstanceId -like 'USBPRINT\*' -or $_.InstanceId -like 'DOT4PRT\*' -or $_.InstanceId -like 'DOT4\*' })
    } catch {
        # Get-PnpDevice needs the PnpDevice module; fall back to CIM.
        try {
            $devices = @(Get-CimInstance Win32_PnPEntity -ErrorAction Stop |
                Where-Object { $_.PNPDeviceID -like 'USBPRINT\*' -or $_.PNPDeviceID -like 'DOT4*' } |
                ForEach-Object {
                    [pscustomobject]@{ InstanceId = $_.PNPDeviceID; FriendlyName = $_.Name; Status = $_.Status; Present = $true }
                })
        } catch {}
    }

    foreach ($d in $devices) {
        $inst = "$($d.InstanceId)".ToLowerInvariant()
        $port = ''
        if ($portByInstance.ContainsKey($inst)) { $port = $portByInstance[$inst] }
        $out.Add([pscustomobject]@{
            InstanceId   = [string]$d.InstanceId
            FriendlyName = [string]$d.FriendlyName
            PortName     = $port
            Present      = [bool]$d.Present
            Status       = [string]$d.Status
        })
    }

    # Ports the registry knows about but PnP does not - a device that WAS here.
    foreach ($inst in $portByInstance.Keys) {
        if (@($out | Where-Object { $_.InstanceId.ToLowerInvariant() -eq $inst }).Count) { continue }
        $out.Add([pscustomobject]@{
            InstanceId   = $inst
            FriendlyName = ''
            PortName     = $portByInstance[$inst]
            Present      = $false
            Status       = 'Not present'
        })
    }

    $out.ToArray()
}

function Test-UsbPrinterAttached {
    <#
    .SYNOPSIS
        Is the physical printer behind a USB queue actually plugged in and on?
    .DESCRIPTION
        The USB answer to Test-PrinterPortReachable. Matching order:
          1. a present device whose registry PortName equals the queue's port
          2. a present device whose friendly name matches the queue or driver
             name (covers printers that register no PortName)
          3. otherwise: not attached

        WorkOffline is reported alongside as corroboration. PrinterStatus is
        deliberately ignored - it reads 'Normal' for a printer that is not
        physically connected.
    .OUTPUTS
        PSCustomObject { Attached; PortName; Device; WorkOffline; Message }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$PortName,
        [string]$Name,
        [string]$DriverName,
        $Devices
    )

    if (-not $Devices) { $Devices = @(Get-UsbPrinterDevice) }
    $present = @($Devices | Where-Object { $_.Present })

    $hit = @($present | Where-Object { $_.PortName -and $_.PortName -ieq $PortName }) | Select-Object -First 1
    $how = 'port mapping'
    if (-not $hit -and ($Name -or $DriverName)) {
        $hit = @($present | Where-Object {
            $fn = "$($_.FriendlyName)"
            if (-not $fn) { return $false }
            ($Name       -and ($fn -like "*$Name*"       -or $Name       -like "*$fn*")) -or
            ($DriverName -and ($fn -like "*$DriverName*" -or $DriverName -like "*$fn*"))
        }) | Select-Object -First 1
        if ($hit) { $how = 'device name match' }
    }

    $offline = $null
    try {
        if ($Name) {
            $esc = $Name -replace "'", "''"
            $w = Get-CimInstance Win32_Printer -Filter "Name='$esc'" -ErrorAction Stop
            if ($w) { $offline = [bool]$w.WorkOffline }
        }
    } catch {}

    if ($hit) {
        return [pscustomobject]@{
            Attached = $true; PortName = $PortName; Device = $hit; WorkOffline = $offline
            Message = "USB device attached on $PortName ($how)$(if ($hit.FriendlyName) { ": $($hit.FriendlyName)" } else { '' })."
        }
    }

    $why = if (@($Devices).Count -eq 0) {
        'no USB print device is attached to this machine at all'
    } else {
        'no attached USB print device claims this port'
    }
    [pscustomobject]@{
        Attached = $false; PortName = $PortName; Device = $null; WorkOffline = $offline
        Message = "No USB printer detected on $PortName - $why. The port keeps existing whether or not the printer is plugged in, so this usually means the printer is unplugged or switched off.$(if ($offline) { ' The queue is also flagged Use Printer Offline.' } else { '' })"
    }
}

# endregion


# ---------------------------------------------------------------------------
# region  This PC as a print SHARE HOST
#
# Everything else in this toolkit looks at a share from the CLIENT side: can I
# reach that host, what does it publish. When the "print server" is an ordinary
# workstation with a USB printer hanging off it - which is what peer-to-peer
# printer sharing actually looks like in a small office - the useful question is
# the other way round: is THIS PC sharing correctly?
#
# Read-only. Every finding names the blocker and the fix; nothing is applied.
# Point and Print, guest auth, SMB signing and the network profile are all
# security controls, and quietly weakening one to make a printer work is how
# sites end up insecure by accident.
# ---------------------------------------------------------------------------

function Get-PrinterShareAcl {
    <#
    .SYNOPSIS
        Who can use and manage a printer, decoded from its security descriptor.
    .DESCRIPTION
        Win32_Printer.GetSecurityDescriptor() returns a DACL of Win32_ACE. The
        printer-specific access bits are:
            0x4  PRINTER_ACCESS_ADMINISTER  -> Manage this printer
            0x8  PRINTER_ACCESS_USE         -> Print
            0x10 JOB_ACCESS_ADMINISTER      -> Manage documents
        A "Print" ACE is normally 0x20008 (USE + READ_CONTROL), so the bits are
        tested rather than the whole mask compared.
    .OUTPUTS
        PSCustomObject { Ok; Entries[]; PrintTrustees[]; Message }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    $entries = [System.Collections.Generic.List[object]]::new()
    try {
        $esc = $Name -replace "'", "''"
        $p = Get-CimInstance Win32_Printer -Filter "Name='$esc'" -ErrorAction Stop
        if (-not $p) { return [pscustomobject]@{ Ok=$false; Entries=@(); PrintTrustees=@(); Message="Printer '$Name' not found." } }
        $r = Invoke-CimMethod -InputObject $p -MethodName GetSecurityDescriptor -ErrorAction Stop
        if ([int]$r.ReturnValue -ne 0) {
            return [pscustomobject]@{ Ok=$false; Entries=@(); PrintTrustees=@(); Message="GetSecurityDescriptor returned $($r.ReturnValue) (usually means insufficient rights)." }
        }
        foreach ($ace in @($r.Descriptor.DACL)) {
            $mask  = [int]$ace.AccessMask
            $rights = @()
            if ($mask -band 0x8)  { $rights += 'Print' }
            if ($mask -band 0x4)  { $rights += 'Manage printer' }
            if ($mask -band 0x10) { $rights += 'Manage documents' }
            if (-not $rights.Count) { $rights += ("mask 0x{0:X}" -f $mask) }
            $tr = $ace.Trustee
            $who = if ($tr.Domain -and $tr.Name) { "$($tr.Domain)\$($tr.Name)" }
                   elseif ($tr.Name) { [string]$tr.Name }
                   else { [string]$tr.SIDString }
            $entries.Add([pscustomobject]@{
                Trustee   = $who
                Sid       = [string]$tr.SIDString
                Rights    = ($rights -join ', ')
                AceType   = $(if ([int]$ace.AceType -eq 0) { 'Allow' } else { 'DENY' })
                AccessMask= $mask
            })
        }
    } catch {
        return [pscustomobject]@{ Ok=$false; Entries=@(); PrintTrustees=@(); Message="Could not read the printer's permissions: $($_.Exception.Message)" }
    }

    $arr = $entries.ToArray()
    # Who can actually print? Well-known SIDs are checked as well as names, so a
    # non-English Windows still resolves correctly.
    $printers = @($arr | Where-Object {
        $_.AceType -eq 'Allow' -and ($_.AccessMask -band 0x8) -and
        ($_.Sid -in @('S-1-1-0','S-1-5-11','S-1-5-32-545') -or
         $_.Sid -like 'S-1-5-21-*-513' -or
         $_.Trustee -match 'Everyone|Authenticated Users|Domain Users|Users')
    } | ForEach-Object { $_.Trustee })

    [pscustomobject]@{
        Ok = $true; Entries = $arr; PrintTrustees = @($printers)
        Message = $(if ($printers.Count) { "Print granted to: $($printers -join ', ')" } else { 'No broad Print grant found.' })
    }
}

function Get-SharingAuthState {
    <#
    .SYNOPSIS
        The SMB/auth settings that decide whether a client can reach this PC's shares.
    .DESCRIPTION
        The half a domain-centric runbook skips. On a workgroup PC, or one where
        clients connect with LOCAL accounts, these are what break peer sharing:

          AllowInsecureGuestAuth        - Windows 11 24H2+ refuses guest SMB by
                                          default and no longer silently falls
                                          back, so 'no password' sharing stopped
                                          working on modern clients.
          LocalAccountTokenFilterPolicy - unset, a remote LOCAL admin account is
                                          filtered to a standard token, which
                                          breaks remote management calls.
          everyoneincludesanonymous /
          restrictnullsessaccess        - the "password protected sharing" pair.
          RequireSecuritySignature      - SMB signing, client and server side.

        Reported tri-state (set / not set -> effective default) like
        Get-PointAndPrintPolicy, because "not set" is not "not enforced".
    #>
    [CmdletBinding()] param()

    $defs = @(
        @{ Name='AllowInsecureGuestAuth'; Path='HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters'; Default=0
           Meaning='0 = this PC will not connect to a share as Guest. Windows 11 24H2+ default; "no password" sharing fails against it.' }
        @{ Name='AllowInsecureGuestAuth'; Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\LanmanWorkstation'; Default=$null
           Meaning='Policy copy of the above; overrides the service setting when present.' }
        @{ Name='LocalAccountTokenFilterPolicy'; Path='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'; Default=0
           Meaning='0 = a remote connection using a LOCAL account gets a filtered (standard) token. Blocks remote admin operations in a workgroup.' }
        @{ Name='everyoneincludesanonymous'; Path='HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'; Default=0
           Meaning='Part of the "password protected sharing" behaviour.' }
        @{ Name='restrictnullsessaccess'; Path='HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters'; Default=1
           Meaning='1 = null (unauthenticated) sessions cannot reach shares.' }
        @{ Name='RequireSecuritySignature'; Path='HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters'; Default=1
           Meaning='SMB signing required when this PC SERVES a share (Windows 11 24H2+ default 1).' }
        @{ Name='RequireSecuritySignature'; Path='HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters'; Default=1
           Meaning='SMB signing required when this PC CONNECTS to a share.' }
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($d in $defs) {
        $val = $null; $isSet = $false
        try {
            $item = Get-ItemProperty -Path $d.Path -Name $d.Name -ErrorAction Stop
            $val = $item.($d.Name); $isSet = $true
        } catch {}
        $eff = $(if ($isSet) { $val } else { $d.Default })
        $rows.Add([pscustomobject]@{
            Name = $d.Name; Path = $d.Path; Value = $val; IsSet = $isSet; EffectiveValue = $eff
            State = $(if ($isSet) { "Set=$val" } elseif ($null -ne $d.Default) { "NotSet -> effective default $($d.Default)" } else { 'NotSet' })
            Meaning = $d.Meaning
        })
    }

    $partOfDomain = $false; $domain = ''
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $partOfDomain = [bool]$cs.PartOfDomain; $domain = [string]$cs.Domain
    } catch {}

    $guestEnabled = $null
    try {
        $g = Get-CimInstance Win32_UserAccount -Filter "LocalAccount=True AND Name='Guest'" -ErrorAction Stop
        if ($g) { $guestEnabled = -not [bool]$g.Disabled }
    } catch {}

    [pscustomobject]@{
        PartOfDomain = $partOfDomain
        Domain       = $domain
        GuestEnabled = $guestEnabled
        Values       = $rows.ToArray()
    }
}

function Get-SharingNetworkProfile {
    <#
    .SYNOPSIS
        Network profile per adapter. A Public profile silently blocks inbound sharing.
    .DESCRIPTION
        Bounded: Get-NetConnectionProfile can hang on a half-configured adapter.
        Split out from Get-PrinterShareHealth so the SHR rules can be exercised
        against a synthesised profile without touching the real network.
    #>
    [CmdletBinding()] param([int]$TimeoutSec = 10)
    @(Invoke-WithTimeout -TimeoutSec $TimeoutSec -ScriptBlock {
        try { Get-NetConnectionProfile -ErrorAction Stop | Select-Object InterfaceAlias, Name, NetworkCategory, IPv4Connectivity } catch { @() }
    })
}

function Get-SharingFirewallRule {
    <#
    .SYNOPSIS
        The File and Printer Sharing firewall rules, found by locale-independent Name. Bounded.
    .DESCRIPTION
        Rules are matched by Name (FPS-*), which is identical on every Windows
        language, rather than by the English display group "File and Printer
        Sharing", which returns nothing on a German or French install. Per-path
        judgement lives in Get-PrintRpcFirewallState; this is the raw list.
    #>
    [CmdletBinding()] param([int]$TimeoutSec = 25)
    $raw = Invoke-WithTimeout -TimeoutSec $TimeoutSec -ScriptBlock {
        $found = @()
        try { $found = @(Get-NetFirewallRule -Name 'FPS-*' -ErrorAction Stop) } catch { $found = @() }
        if (-not $found.Count) {
            try { $found = @(Get-NetFirewallRule -ErrorAction Stop | Where-Object { $_.Name -like 'FPS-*' }) } catch { $found = @() }
        }
        $found | Select-Object Name, DisplayName, Enabled, Profile, Direction, Action
    }
    if ($null -eq $raw) { return @() }
    return @($raw | Where-Object { $_ })
}

function Get-PrinterShareHealth {
    <#
    .SYNOPSIS
        Is THIS PC sharing its printers correctly? Read-only, with SHR findings.
    .DESCRIPTION
        Answers the host-side question for a PC that shares printers. Groups:

          * core sharing state - which queues are shared, the spooler, the network
            profile (Public silently blocks inbound sharing) and whether print$ is
            readable
          * the three separate inbound paths a client needs, judged per path for the
            ACTIVE profile: SMB 445, the RPC endpoint mapper 135 and the Print
            Spooler's RPC traffic (Get-PrintRpcFirewallState, by rule Name)
          * what the spooler is actually listening on, and whether a pinned
            RpcTcpPort is in effect (Get-PrintSpoolerListener)
          * workgroup / local-account auth, printer permissions, Point and Print

        Findings use an SHR prefix so they are obviously host-side, as distinct from
        the client-side MAP set. One code per condition.
    .OUTPUTS
        PSCustomObject { IsHost; Shares[]; Spooler; Profiles[]; Firewall; FirewallInboundEnabled;
                         RpcFirewall; Listener; RpcConfig; PinnedPort; PrintShare; Auth;
                         PointAndPrint; Findings[]; Summary }
    #>
    [CmdletBinding()]
    param([int]$TimeoutSec = 25)

    $findings = [System.Collections.Generic.List[object]]::new()
    $spooler  = Get-SpoolerHealth
    $policy   = Get-PointAndPrintPolicy
    $auth     = Get-SharingAuthState
    $usbDev   = @()
    try { $usbDev = @(Get-UsbPrinterDevice) } catch {}

    # ---- which queues are shared -------------------------------------------
    $all    = @(Get-LocalPrinterInventory -IncludeVirtual)
    $shared = @($all | Where-Object { $_.Shared -and $_.Type -ne 'Connection' })

    $shares = [System.Collections.Generic.List[object]]::new()
    foreach ($s in $shared) {
        $acl = Get-PrinterShareAcl -Name $s.Name
        $usb = $null
        if ($s.PortType -eq 'USB') {
            $usb = Test-UsbPrinterAttached -PortName $s.PortName -Name $s.Name -DriverName $s.DriverName -Devices $usbDev
        }
        $shares.Add([pscustomobject]@{
            Name          = $s.Name
            ShareName     = $s.ShareName
            Unc           = "\\$env:COMPUTERNAME\$($s.ShareName)"
            PortName      = $s.PortName
            PortType      = $s.PortType
            DriverName    = $s.DriverName
            DriverMajor   = $s.DriverMajor
            PackageAware  = $s.PackageAware
            IsIppClass    = $s.IsIppClass
            Published     = $s.Published
            UsbAttached   = $(if ($usb) { $usb.Attached } else { $null })
            UsbMessage    = $(if ($usb) { $usb.Message } else { '' })
            AclOk         = $acl.Ok
            PrintTrustees = $acl.PrintTrustees
            Acl           = $acl.Entries
            AclMessage    = $acl.Message
        })
    }
    $isHost = ($shares.Count -gt 0)

    # ---- network profile, firewall per path, listener -----------------------
    $profiles = @(Get-SharingNetworkProfile | Where-Object { $_ })
    $fwRules  = @(Get-SharingFirewallRule -TimeoutSec $TimeoutSec)
    $activeCats = @($profiles | ForEach-Object { "$($_.NetworkCategory)" } | Where-Object { $_ } | Select-Object -Unique)
    $inboundEnabled = @($fwRules | Where-Object { "$($_.Direction)" -eq 'Inbound' -and "$($_.Enabled)" -eq 'True' })

    $rpcCfg = Get-PrintRpcConfig
    $pinnedPort = 0
    $tcpRow = @($rpcCfg.Values | Where-Object { $_.Name -eq 'RpcTcpPort' })[0]
    if ($tcpRow -and $tcpRow.IsSet) { try { $pinnedPort = [int]$tcpRow.Value } catch { $pinnedPort = 0 } }
    $listener = Get-PrintSpoolerListener
    $rpcFw    = Get-PrintRpcFirewallState -TimeoutSec $TimeoutSec -PinnedPort $pinnedPort

    # ---- print$ (bounded: a filesystem call is not bounded by any TCP timeout) --
    $printShare = $false
    $psr = Invoke-WithTimeout -TimeoutSec 8 -Parameters @{ P = "\\$env:COMPUTERNAME\print$" } -ScriptBlock {
        param([string]$P)
        try { [bool](Test-Path -LiteralPath $P -ErrorAction Stop) } catch { $false }
    }
    if ($null -ne $psr) { $printShare = [bool](@($psr)[0]) }

    # =======================================================================
    # Findings - only meaningful if this PC actually hosts a share
    # =======================================================================
    if ($isHost) {

        if (-not $spooler.Running) {
            $findings.Add((New-PrinterFinding -Code 'SHR003' -Severity 'ERROR' -Scope 'Machine' `
                -Detail "This PC shares $($shares.Count) printer(s) but the Print Spooler is $($spooler.Status). Every client of those shares is failing." `
                -Fix 'Start the Spooler and set it to Automatic. A PC that shares printers cannot have the spooler disabled.'))
        }

        if ($activeCats -contains 'Public') {
            $findings.Add((New-PrinterFinding -Code 'SHR001' -Severity 'ERROR' -Scope 'Machine' `
                -Detail "A network adapter is on the Public profile ($($activeCats -join ', ')) while this PC shares printers. Windows blocks the inbound sharing rules on Public, so clients cannot connect however correct the share is." `
                -Fix 'Set the LAN connection to Private (or Domain). Settings > Network > properties > Private network. Do not open firewall ports to work around this.'))
        }

        # Three inbound paths, judged separately. The old check fired only when
        # EVERY sharing rule was off, so "SMB on, spooler off" - net view works,
        # adding the printer fails - was invisible.
        if ($rpcFw.State -ne 'Ok') {
            $findings.Add((New-PrinterFinding -Code 'SHR002' -Severity 'WARN' -Scope 'Machine' `
                -Detail "The inbound firewall rules could not be judged ($($rpcFw.State)): $($rpcFw.Message)" `
                -Fix 'Check the firewall directly: SMB 445, the RPC endpoint mapper 135 and the Print Spooler must all be allowed inbound for the active profile.'))
        } else {
            $profTxt = $(if (@($rpcFw.ActiveProfiles).Count) { @($rpcFw.ActiveProfiles) -join ', ' } else { 'unknown' })
            if (-not $rpcFw.Smb.Allowed) {
                $findings.Add((New-PrinterFinding -Code 'SHR002' -Severity 'ERROR' -Scope 'Machine' `
                    -Detail "No enabled inbound rule allows SMB (445) for the active profile ($profTxt). Clients cannot reach this PC's shares at all." `
                    -Fix 'Share Host -> Apply safe fix, which enables the scoped built-in rules. Prefer that over opening custom ports.'))
            }
            if (-not $rpcFw.Epm.Allowed) {
                $findings.Add((New-PrinterFinding -Code 'SHR013' -Severity 'ERROR' -Scope 'Machine' `
                    -Detail "No enabled inbound rule allows the RPC endpoint mapper (135) for the active profile ($profTxt). A client can reach the share over SMB and still fail to connect the printer." `
                    -Fix 'Share Host -> Fix print RPC -> enable the built-in rules for the active profile.'))
            }
            if (-not $rpcFw.Spooler.Allowed) {
                $findings.Add((New-PrinterFinding -Code 'SHR014' -Severity 'ERROR' -Scope 'Machine' `
                    -Detail "No enabled inbound rule allows the Print Spooler's RPC traffic for the active profile ($profTxt). File sharing can work while printer connections fail: net view succeeds, adding the printer does not." `
                    -Fix 'Share Host -> Fix print RPC -> enable the built-in spooler rules for the active profile.'))
            } elseif ($rpcFw.Spooler.LocalSubnetOnly) {
                $findings.Add((New-PrinterFinding -Code 'SHR012' -Severity 'INFO' -Scope 'Machine' `
                    -Detail 'The rules allowing the Print Spooler are scoped to LocalSubnet. A client on another subnet, VLAN or VPN is refused even though the rule is enabled.' `
                    -Fix 'If clients are on another subnet: pin the spooler port and create a narrow rule scoped to their subnet (Share Host -> Fix print RPC). Do not widen to Any.'))
            }
        }

        if ($pinnedPort -gt 0 -and $listener.State -eq 'Ok' -and -not (@($listener.Ports) -contains $pinnedPort)) {
            $findings.Add((New-PrinterFinding -Code 'SHR015' -Severity 'ERROR' -Scope 'Machine' `
                -Detail ("RpcTcpPort is set to $pinnedPort but the Print Spooler is listening on " +
                         $(if (@($listener.Ports).Count) { @($listener.Ports) -join ', ' } else { 'no TCP port' }) +
                         ". The setting is not in effect.") `
                -Fix 'Restart the Print Spooler. If it still does not bind, the port is taken - choose another (Share Host -> Fix print RPC).'))
        }

        # ---- workgroup / local-account auth --------------------------------
        if (-not $auth.PartOfDomain) {
            $guest = @($auth.Values | Where-Object { $_.Name -eq 'AllowInsecureGuestAuth' -and $_.Path -like '*LanmanWorkstation\Parameters' })[0]
            $findings.Add((New-PrinterFinding -Code 'SHR004' -Severity 'WARN' -Scope 'Machine' `
                -Detail "This PC is in a WORKGROUP ($($auth.Domain)) and shares printers, so every client must authenticate with an account that exists HERE. Guest fallback is $(if ($guest -and [int]$guest.EffectiveValue -eq 1) { 'allowed' } else { 'refused' }) ($($guest.State)) - Windows 11 24H2 and later refuse insecure guest SMB by default and no longer fall back silently." `
                -Fix 'Create a matching local account (same username and password) on this PC for each user who needs to print, or move the printer to a direct network connection so no host authentication is needed. Do not enable insecure guest SMB to work around it.'))

            $latfp = @($auth.Values | Where-Object { $_.Name -eq 'LocalAccountTokenFilterPolicy' })[0]
            if ($latfp -and [int]$latfp.EffectiveValue -ne 1) {
                $findings.Add((New-PrinterFinding -Code 'SHR005' -Severity 'INFO' -Scope 'Machine' `
                    -Detail "LocalAccountTokenFilterPolicy is $($latfp.State). A remote connection using a LOCAL admin account is filtered to a standard token, so remote management calls against this PC can fail even when printing itself works." `
                    -Fix 'Expected and secure by default. Only relevant if you need remote administration of this workgroup PC - printing does not require it.'))
            }
        }

        # ---- per-share findings --------------------------------------------
        foreach ($sh in $shares) {
            if ($sh.AclOk -and @($sh.PrintTrustees).Count -eq 0) {
                $findings.Add((New-PrinterFinding -Code 'SHR006' -Severity 'WARN' -Scope 'Machine' -Printer $sh.Name `
                    -Detail "Shared queue '$($sh.Name)' grants no broad Print right - no Everyone, Authenticated Users, Domain Users or Users entry was found in its permissions." `
                    -Fix 'Printer properties > Security: grant Print to the group that needs it. Without this a client can see the share and still be refused.'))
            }
            if ($sh.PortType -eq 'USB' -and $sh.UsbAttached -eq $false) {
                $findings.Add((New-PrinterFinding -Code 'SHR008' -Severity 'ERROR' -Scope 'Machine' -Printer $sh.Name `
                    -Detail "Shared queue '$($sh.Name)' is on $($sh.PortName) but no USB printer is attached to this PC. Every client of $($sh.Unc) is failing, and the queue still reports a normal status. $($sh.UsbMessage)" `
                    -Fix 'Reconnect and power on the printer. A USB-shared printer makes this PC part of the printing infrastructure - it has to stay on, awake and plugged in for anyone else to print.'))
            }
            if ($sh.IsIppClass) {
                $findings.Add((New-PrinterFinding -Code 'SHR009' -Severity 'WARN' -Scope 'Machine' -Printer $sh.Name `
                    -Detail "Shared queue '$($sh.Name)' uses the generic '$($sh.DriverName)'. Every client that connects inherits that driver and its missing trays, duplex and finishing options." `
                    -Fix 'Install the real vendor driver on this host and rebuild the queue (Printers tab -> Repoint / Rebuild) before clients connect to it.'))
            }
            $pkgOnly = @($policy.Machine | Where-Object { $_.Name -eq 'PackagePointAndPrintOnly' })[0]
            if ($pkgOnly -and [int]$pkgOnly.EffectiveValue -eq 1) {
                $notAware = ($null -ne $sh.PackageAware -and -not $sh.PackageAware)
                $isV3     = ($null -ne $sh.DriverMajor -and [int]$sh.DriverMajor -eq 3)
                if ($notAware -or $isV3) {
                    $findings.Add((New-PrinterFinding -Code 'SHR007' -Severity 'WARN' -Scope 'Machine' -Printer $sh.Name `
                        -Detail "Shared queue '$($sh.Name)' uses $(if ($isV3) { 'a v3 driver' } else { 'a driver that is not package-aware' }) while PackagePointAndPrintOnly is $($pkgOnly.State). Clients will refuse to install it and fail with 0x00000bcb." `
                        -Fix 'Deploy a package-aware (v4) driver on this host, or pre-stage the driver package on each client with pnputil so no Point and Print download is needed.'))
                }
            }
        }

        # SHR011 - enumeration versus connection. Listing shares needs a right that is
        # usually admin-only, so a client can see "net view ... System error 5" against
        # this PC and still print perfectly. Reported once, as a note.
        $findings.Add((New-PrinterFinding -Code 'SHR011' -Severity 'INFO' -Scope 'Machine' `
            -Detail "Clients may be refused when LISTING this PC's shares (net view -> System error 5) while still being able to connect to a named printer. Share enumeration and printer connection are separate rights." `
            -Fix 'If a client reports access denied, test the specific share by name - Shares -> Test connection on the client does exactly that.'))

        $restrict = @($policy.Machine | Where-Object { $_.Name -eq 'RestrictDriverInstallationToAdministrators' })[0]
        if ($restrict -and [int]$restrict.EffectiveValue -eq 1) {
            $findings.Add((New-PrinterFinding -Code 'SHR010' -Severity 'INFO' -Scope 'Machine' `
                -Detail "Clients connecting to these shares must install the driver, and RestrictDriverInstallationToAdministrators is $($restrict.State) on THIS host's policy - the same restriction almost certainly applies on the clients." `
                -Fix 'Pre-stage the matching driver on each client as SYSTEM (Shares -> Install driver...), then add the connection. Do not disable the restriction.'))
        }
    }

    $arr = $findings.ToArray()
    [pscustomobject]@{
        IsHost                 = $isHost
        Shares                 = $shares.ToArray()
        Spooler                = $spooler
        Profiles               = @($profiles)
        Firewall               = @($fwRules)
        FirewallInboundEnabled = $inboundEnabled.Count
        RpcFirewall            = $rpcFw
        Listener               = $listener
        RpcConfig              = $rpcCfg
        PinnedPort             = $pinnedPort
        PrintShare             = $printShare
        Auth                   = $auth
        PointAndPrint          = $policy
        Findings               = $arr
        Summary                = "$(@($arr | Where-Object { $_.Severity -eq 'ERROR' }).Count) error(s), $(@($arr | Where-Object { $_.Severity -eq 'WARN' }).Count) warning(s), $(@($arr | Where-Object { $_.Severity -eq 'INFO' }).Count) note(s) across $($shares.Count) share(s)."
    }
}

# endregion


# ---------------------------------------------------------------------------
# region  Per-MACHINE printer connections
#
# Add-Printer -ConnectionName is per-USER and runs in the caller's context, so
# from SYSTEM it connects the printer for SYSTEM and nobody else. The Windows
# answer for "every user of this PC should get this printer" is a per-COMPUTER
# connection: rundll32 printui.dll,PrintUIEntry /ga. That is the right shape for
# Atera / Backstage, where there may be no user signed in at all.
#
# Two traps this code handles, neither obvious:
#   * rundll32 returns IMMEDIATELY and reports no useful exit code, so "it ran"
#     proves nothing - the result has to be verified out of the registry.
#   * /ge only opens a dialog box, so it is useless for scripted enumeration;
#     the connection list is read from the registry instead.
#   * A /ga connection does not appear for an already-signed-in user until they
#     log on again (or the spooler is restarted).
# ---------------------------------------------------------------------------

# Modern Windows records per-machine connections under SYSTEM; the SOFTWARE path
# is the historical location. Both are read so this keeps working either way.
$script:MachineConnRoots = @(
    'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Providers\LanMan Print Services\Servers'
    'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\LanMan Print Services\Servers'
)

function Get-MachinePrinterConnections {
    <#
    .SYNOPSIS
        Per-computer printer connections on this machine (the /ga list), from the registry.
    .OUTPUTS
        PSCustomObject[] { Server; Printer; Unc; Source }
    #>
    [CmdletBinding()] param()

    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($root in $script:MachineConnRoots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        foreach ($srv in @(Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue)) {
            $printersKey = Join-Path $srv.PSPath 'Printers'
            if (-not (Test-Path -LiteralPath $printersKey)) { continue }
            foreach ($prn in @(Get-ChildItem -LiteralPath $printersKey -ErrorAction SilentlyContinue)) {
                $server = $srv.PSChildName.TrimStart('\')
                $name   = $prn.PSChildName
                $out.Add([pscustomobject]@{
                    Server  = $server
                    Printer = $name
                    Unc     = "\\$server\$name"
                    Source  = $root
                })
            }
        }
    }
    $out.ToArray()
}

function Test-MachinePrinterConnection {
    <# .SYNOPSIS Is this UNC present as a per-computer connection? #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ConnectionName)
    $want = $ConnectionName.TrimEnd('\')
    [bool](@(Get-MachinePrinterConnections | Where-Object { $_.Unc -ieq $want }).Count)
}

function Add-PrinterConnectionForMachine {
    <#
    .SYNOPSIS
        Add a shared printer for EVERY user of this PC (per-computer connection).
    .DESCRIPTION
        Runs rundll32 printui.dll,PrintUIEntry /ga and then VERIFIES the result
        against the registry, because rundll32 exits immediately with no useful
        status. Existing users see the printer at their next sign-in; pass
        -RestartSpooler to make the spooler pick it up sooner.
    .OUTPUTS
        PSCustomObject { Success; Message; Unc; AlreadyPresent }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$ConnectionName,
        [int]$TimeoutSec = 25,
        [switch]$RestartSpooler,
        [scriptblock]$OnPoll
    )

    $unc = $ConnectionName.Trim().TrimEnd('\')
    if ($unc -notmatch '^\\\\[^\\]+\\.+') {
        return [pscustomobject]@{ Success=$false; Unc=$unc; AlreadyPresent=$false; Message="'$unc' is not a \\server\share path." }
    }

    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $isAdmin = $false
    try { $isAdmin = ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } catch {}
    if (-not $isAdmin) {
        return [pscustomobject]@{ Success=$false; Unc=$unc; AlreadyPresent=$false; Message='A per-computer connection needs elevation. Run as Administrator, or as SYSTEM from Backstage / Atera.' }
    }

    if (Test-MachinePrinterConnection -ConnectionName $unc) {
        return [pscustomobject]@{ Success=$true; Unc=$unc; AlreadyPresent=$true; Message="'$unc' is already a per-computer connection on this PC." }
    }
    if (-not $PSCmdlet.ShouldProcess($unc, 'Add per-computer printer connection')) {
        return [pscustomobject]@{ Success=$false; Unc=$unc; AlreadyPresent=$false; Message='Skipped (WhatIf).' }
    }

    try {
        Start-Process -FilePath "$env:SystemRoot\System32\rundll32.exe" `
            -ArgumentList ('printui.dll,PrintUIEntry /ga /n"{0}"' -f $unc) -WindowStyle Hidden -ErrorAction Stop
    } catch {
        return [pscustomobject]@{ Success=$false; Unc=$unc; AlreadyPresent=$false; Message="Could not start printui: $($_.Exception.Message)" }
    }

    # rundll32 has already returned; poll the registry for the real outcome.
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $ok = $false
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        if ($OnPoll) { try { & $OnPoll } catch {} }
        if (Test-MachinePrinterConnection -ConnectionName $unc) { $ok = $true; break }
    }

    if (-not $ok) {
        return [pscustomobject]@{
            Success=$false; Unc=$unc; AlreadyPresent=$false
            Message="printui was launched but '$unc' did not appear as a per-computer connection within $TimeoutSec s. Usual causes: the share name is wrong, the host refused the connection, or the driver could not be installed (Point and Print restricts that to administrators). Check the host is reachable and the driver is staged locally."
        }
    }

    $extra = ''
    if ($RestartSpooler) { $extra = ' ' + (Restart-PrintSpooler).Message }
    [pscustomobject]@{
        Success=$true; Unc=$unc; AlreadyPresent=$false
        Message="Added '$unc' as a per-computer connection - it appears for every user of this PC at their next sign-in.$extra"
    }
}

function Remove-PrinterConnectionForMachine {
    <# .SYNOPSIS Remove a per-computer printer connection (/gd), verified. #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$ConnectionName,
        [int]$TimeoutSec = 25,
        [scriptblock]$OnPoll
    )
    $unc = $ConnectionName.Trim().TrimEnd('\')
    if (-not (Test-MachinePrinterConnection -ConnectionName $unc)) {
        return [pscustomobject]@{ Success=$true; Unc=$unc; Message="'$unc' is not a per-computer connection on this PC - nothing to remove." }
    }
    if (-not $PSCmdlet.ShouldProcess($unc, 'Remove per-computer printer connection')) {
        return [pscustomobject]@{ Success=$false; Unc=$unc; Message='Skipped (WhatIf).' }
    }
    try {
        Start-Process -FilePath "$env:SystemRoot\System32\rundll32.exe" `
            -ArgumentList ('printui.dll,PrintUIEntry /gd /n"{0}"' -f $unc) -WindowStyle Hidden -ErrorAction Stop
    } catch {
        return [pscustomobject]@{ Success=$false; Unc=$unc; Message="Could not start printui: $($_.Exception.Message)" }
    }
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        if ($OnPoll) { try { & $OnPoll } catch {} }
        if (-not (Test-MachinePrinterConnection -ConnectionName $unc)) {
            return [pscustomobject]@{ Success=$true; Unc=$unc; Message="Removed the per-computer connection '$unc'." }
        }
    }
    [pscustomobject]@{ Success=$false; Unc=$unc; Message="'$unc' was still present $TimeoutSec s after the removal was requested." }
}

# endregion


# ---------------------------------------------------------------------------
# region  Extended printer port probing
#
# The standard three (RAW 9100, LPR 515, IPP 631) answer "can I print to it".
# When all three are shut, the useful follow-up is "is the device alive at all",
# and that is what the extended set answers: a printer that responds on 80 or
# 443 but not on 9100 is powered on and reachable with raw printing DISABLED,
# which is a completely different fix from a printer that is off.
# ---------------------------------------------------------------------------

# Curated, TCP-probeable and labelled. UDP-only protocols (SNMP 161, Canon BJNP
# 8611-8613, Epson ENPC 3289, WS-Discovery 3702) are deliberately absent - a TCP
# probe against them proves nothing and would produce false "closed" noise.
$script:PrinterPortCatalogue = @(
    @{ Port = 9100; Name = 'RAW / JetDirect';        Standard = $true;  Note = 'The normal way Windows prints to a network printer.' }
    @{ Port = 515;  Name = 'LPR / LPD';              Standard = $true;  Note = 'Fallback when RAW is disabled; needs a queue name.' }
    @{ Port = 631;  Name = 'IPP';                    Standard = $true;  Note = 'Internet Printing Protocol.' }
    @{ Port = 9101; Name = 'RAW port 2';             Standard = $false; Note = 'Second RAW port on multi-port print servers.' }
    @{ Port = 9102; Name = 'RAW port 3';             Standard = $false; Note = 'Third RAW port on multi-port print servers.' }
    @{ Port = 443;  Name = 'HTTPS / IPPS';           Standard = $false; Note = 'Secure web UI or IPP over TLS - proves the device is alive.' }
    @{ Port = 80;   Name = 'HTTP web UI';            Standard = $false; Note = 'Device web page - proves the device is alive even if printing is off.' }
    @{ Port = 8080; Name = 'HTTP alt web UI';        Standard = $false; Note = 'Alternate admin page used by some vendors.' }
    @{ Port = 5357; Name = 'WSD (HTTP)';             Standard = $false; Note = 'Web Services on Devices - what a WSD queue talks to.' }
    @{ Port = 5358; Name = 'WSD (HTTPS)';            Standard = $false; Note = 'Secure WSD transport.' }
    @{ Port = 445;  Name = 'SMB';                    Standard = $false; Note = 'A shared printer hosted on a PC, not a direct device.' }
    @{ Port = 139;  Name = 'NetBIOS session';        Standard = $false; Note = 'Legacy SMB path to a shared printer.' }
    @{ Port = 3910; Name = 'Vendor discovery/setup'; Standard = $false; Note = 'Used by some HP/Canon devices; seen in Location URLs.' }
    @{ Port = 9220; Name = 'JetDirect scan';         Standard = $false; Note = 'Scanning channel on HP multifunction devices.' }
    @{ Port = 21;   Name = 'FTP print';              Standard = $false; Note = 'Legacy FTP printing, occasionally the only thing left open.' }
    @{ Port = 23;   Name = 'Telnet admin';           Standard = $false; Note = 'Legacy admin channel; usually a sign of old firmware.' }
)

function Get-PrinterPortCatalogue {
    <# .SYNOPSIS The ports this toolkit knows how to probe, with what each one means. #>
    [CmdletBinding()] param([switch]$StandardOnly)
    $rows = foreach ($p in $script:PrinterPortCatalogue) {
        if ($StandardOnly -and -not $p.Standard) { continue }
        [pscustomobject]@{ Port = $p.Port; Name = $p.Name; Standard = $p.Standard; Note = $p.Note }
    }
    @($rows)
}

function Test-TcpPortSetFast {
    <#
    .SYNOPSIS
        Probe MANY ports on ONE host concurrently. Returns a result per port.
    .DESCRIPTION
        Test-TcpPortFast is one port at a time, so a 16-port sweep would cost 16
        timeouts back to back. This issues every connect at once and waits a
        single timeout window, so the extended scan costs about as much as one
        probe. Sockets are closed in a finally - EndConnect on a timed-out handle
        leaks one otherwise.
    .OUTPUTS
        PSCustomObject[] { Port; Open; Ms }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][int[]]$Ports,
        [int]$TimeoutMs = 1500
    )
    $clients = @{}; $asyncs = @{}
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    foreach ($p in ($Ports | Select-Object -Unique)) {
        try {
            $c = New-Object System.Net.Sockets.TcpClient
            $clients[$p] = $c
            $asyncs[$p]  = $c.BeginConnect($HostName, $p, $null, $null)
        } catch {
            if ($clients.ContainsKey($p)) { try { $clients[$p].Close() } catch {}; $clients.Remove($p) }
        }
    }
    # One wait for all of them, polled so an all-open host returns quickly.
    $deadline = (Get-Date).AddMilliseconds($TimeoutMs)
    while ((Get-Date) -lt $deadline) {
        $pending = @($asyncs.Keys | Where-Object { -not $asyncs[$_].IsCompleted })
        if (-not $pending.Count) { break }
        Start-Sleep -Milliseconds 60
    }
    $sw.Stop()

    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($p in ($Ports | Select-Object -Unique)) {
        $open = $false
        if ($clients.ContainsKey($p)) {
            try {
                if ($asyncs[$p].IsCompleted) {
                    try { $clients[$p].EndConnect($asyncs[$p]) } catch {}
                    $open = $clients[$p].Connected
                }
            } catch { $open = $false }
            try { $clients[$p].Close() } catch {}
        }
        $out.Add([pscustomobject]@{ Port = $p; Open = $open; Ms = [int]$sw.ElapsedMilliseconds })
    }
    $out.ToArray()
}

# endregion


# ---------------------------------------------------------------------------
# region  Share access: reachable vs enumerable vs connectable
#
# These are THREE independent questions and conflating them sends techs the
# wrong way. The trigger for this code was:
#
#     net view PRINTSRV01
#     System error 5 has occurred.  Access is denied.
#
# That says share ENUMERATION was refused. Enumeration needs rights that are
# commonly admin-only and are routinely restricted further on a domain
# controller. It says nothing about whether the user can print: the spooler
# reaches a printer share over its OWN RPC interface, which frequently works
# when net view does not. "Cannot list, but can connect" is a normal outcome.
# ---------------------------------------------------------------------------

if (-not ('PrtNative.Spool' -as [type])) {
    # OpenPrinter is exactly what Windows calls when a client connects to a
    # printer share, so it is the definitive "can I use this?" test - and unlike
    # Add-Printer it creates nothing. Test-Path is useless here: a print share is
    # not a filesystem object and Test-Path returns False for shares that work.
    # Built from an array rather than a here-string ON PURPOSE. Build-PrinterToolkit
    # wraps this whole module in a here-string to inline it into the GUI, so any
    # line here starting with '@ at column 0 would terminate that here-string and
    # produce a corrupt build. The build guard catches it, but avoiding it is better.
    $sig = @(
        '[DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "OpenPrinterW")]'
        'public static extern bool OpenPrinter(string pPrinterName, out System.IntPtr phPrinter, System.IntPtr pDefault);'
        '[DllImport("winspool.drv", SetLastError = true)]'
        'public static extern bool ClosePrinter(System.IntPtr hPrinter);'
    ) -join [Environment]::NewLine
    Add-Type -Namespace 'PrtNative' -Name 'Spool' -MemberDefinition $sig -ErrorAction SilentlyContinue
}

function Get-SmbErrorExplanation {
    <#
    .SYNOPSIS
        Turn a Windows network error number into something a technician can act on.
    .DESCRIPTION
        Ported from Test-ShareCredential in ALL_MAPPED_DRIVE_FIXER - the most
        directly reusable thing in that project. These codes come back from
        net use, OpenPrinter and SMB alike.
    #>
    [CmdletBinding()]
    param([int]$Code, [string]$Text, [string]$HostName = 'the host')

    switch ($Code) {
        0     { 'Succeeded.' }
        5     { 'Authenticated, but ACCESS DENIED - the identity in play has no rights to this. For a printer, check the printer''s Security tab; for share enumeration, listing shares is usually restricted to administrators.' }
        53    { 'Network path not found - wrong host or share name, or SMB (445) is blocked.' }
        64    { 'The network name is no longer available - the host went away or the SMB session dropped.' }
        67    { 'That SHARE name does not exist on the server (the host answered, the name did not match).' }
        86    { 'The network password is not correct.' }
        1219  { 'A conflicting connection to this server already exists under different credentials. Clear it first - every connection to one server shares a single authenticated session.' }
        1231  { 'Network location cannot be reached - host down, wrong address, or off-subnet.' }
        1311  { "No logon servers available - Windows attempted a DOMAIN logon. Use a LOCAL account instead: $HostName\user." }
        1326  { 'Bad username or password - or the wrong authority. For a workgroup host use HOST\user, not a bare username.' }
        1327  { 'Account restriction: blank passwords not allowed, logon-hours restriction, or the account is disabled or expired.' }
        1355  { "The specified domain does not exist or cannot be contacted - the old domain is gone. Use a LOCAL account: $HostName\user." }
        1385  { 'The account lacks the "Access this computer from the network" right on that machine.' }
        1722  { 'The RPC server is unavailable - the Print Spooler is stopped on the host, or its RPC traffic (135, or the spooler''s own port) is blocked by a firewall.' }
        1726  { 'The remote procedure call failed - the host''s spooler dropped the call.' }
        1727  { 'The remote procedure call failed and did not execute - the connection was lost before the spooler ran it.' }
        1753  { 'No endpoint is registered with the endpoint mapper - the spooler on the host is not listening where port 135 says it should be.' }
        1801  { 'Invalid printer name - the share name is wrong, or the host does not publish it.' }
        default {
            if ($Code) { "Windows error $Code." } else { ("$Text").Trim() }
        }
    }
}

function Get-PrinterHostForms {
    <#
    .SYNOPSIS
        Both the name and IP forms of a host.
    .DESCRIPTION
        Credentials and sessions get cached under whichever form was used, so a
        teardown that only clears one leaves the old identity alive under the
        other. Ported from Get-BothHostForms in the mapped-drive GUI.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$HostName)

    $h = $HostName.TrimStart('\').Trim()
    $forms = [System.Collections.Generic.List[string]]::new()
    if ($h) { $forms.Add($h) }
    try {
        if ($h -match '^\d{1,3}(\.\d{1,3}){3}$') {
            # literal IP - add the reverse name
            $n = [System.Net.Dns]::GetHostEntry($h).HostName
            if ($n) { $forms.Add(($n -split '\.')[0]) }
        } else {
            foreach ($a in [System.Net.Dns]::GetHostAddresses($h)) {
                if ($a.AddressFamily -eq 'InterNetwork') { $forms.Add($a.IPAddressToString) }
            }
            if ($h -match '\.') { $forms.Add(($h -split '\.')[0]) }
        }
    } catch {}
    @($forms.ToArray() | Where-Object { $_ } | Select-Object -Unique)
}

function Test-PrinterShareOpen {
    <#
    .SYNOPSIS
        Can THIS identity open a printer share on the LIVE host - not from cache? Bounded.
    .DESCRIPTION
        Calls OpenPrinter2 with PRINTER_OPTION_NO_CACHE on \\host\share, then closes the
        handle. Creates nothing.

        Why not plain OpenPrinter: measured on this toolkit's own development machine,
        OpenPrinter returned a handle in 13 ms for a mapped printer whose server did not
        even resolve in DNS - Windows answered from cached printer data. Every earlier
        "opens fine" result could have been that. With NO_CACHE the same call went to the
        server and failed after 16 s with 1722 (RPC server unavailable): the truth, and
        a far more specific error than 64.

        What a success proves, precisely: this identity reached the host's spooler and
        obtained a handle, now. Not that a driver will install, that a job will be
        accepted, or that a page will print.

        Runs in its own runspace with a wall-clock cap, because a no-cache open against a
        host that swallows the connection legitimately takes many seconds.
    .OUTPUTS
        PSCustomObject { Unc; State; Opened; Win32Error; Message }
        State is Ok | Failed | Timeout | Skipped. Opened is true only for Ok.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Unc,
        [int]$TimeoutSec = 30
    )
    $hostName = 'the host'
    if ($Unc -match '^\\\\([^\\]+)\\') { $hostName = $matches[1] }

    $raw = Invoke-WithTimeout -TimeoutSec $TimeoutSec -Parameters @{ Target = $Unc } -ScriptBlock {
        param([string]$Target)
        # PRINTER_OPTION_NO_CACHE = 1, PRINTER_OPTION_CACHE = 2. Built from an array,
        # never a here-string: the build inlines this module inside a here-string.
        $sig = @(
            '[StructLayout(LayoutKind.Sequential)] public struct PRINTER_OPTIONS { public uint cbSize; public uint dwFlags; }'
            '[DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "OpenPrinter2W")]'
            'public static extern bool OpenPrinter2(string pPrinterName, out System.IntPtr phPrinter, System.IntPtr pDefault, ref PRINTER_OPTIONS pOptions);'
            '[DllImport("winspool.drv", SetLastError = true)]'
            'public static extern bool ClosePrinter(System.IntPtr hPrinter);'
        ) -join [Environment]::NewLine
        try {
            if (-not ('PrtNoCache.Spool' -as [type])) { Add-Type -Namespace 'PrtNoCache' -Name 'Spool' -MemberDefinition $sig -ErrorAction Stop }
        } catch {
            return [pscustomobject]@{ Loaded = $false; Opened = $false; Err = 0; Msg = $_.Exception.Message }
        }
        $opt = New-Object 'PrtNoCache.Spool+PRINTER_OPTIONS'
        $opt.cbSize = [uint32][System.Runtime.InteropServices.Marshal]::SizeOf([type]'PrtNoCache.Spool+PRINTER_OPTIONS')
        $opt.dwFlags = [uint32]1
        $h = [IntPtr]::Zero; $ok = $false; $err = 0
        try {
            $ok = [PrtNoCache.Spool]::OpenPrinter2($Target, [ref]$h, [IntPtr]::Zero, [ref]$opt)
            if (-not $ok) { $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error() }
        } finally {
            if ($h -ne [IntPtr]::Zero) { try { [void][PrtNoCache.Spool]::ClosePrinter($h) } catch {} }
        }
        [pscustomobject]@{ Loaded = $true; Opened = [bool]$ok; Err = [int]$err; Msg = '' }
    }

    if ($null -eq $raw) {
        return [pscustomobject]@{ Unc = $Unc; State = 'Timeout'; Opened = $false; Win32Error = $null
            Message = "Opening '$Unc' on the live host did not return within $TimeoutSec seconds. No answer is not the same as a refusal." }
    }
    $o = @($raw)[0]
    if (-not $o -or -not [bool]$o.Loaded) {
        return [pscustomobject]@{ Unc = $Unc; State = 'Skipped'; Opened = $false; Win32Error = $null
            Message = "The open test could not run: $(if ($o) { $o.Msg } else { 'no result was returned' })." }
    }
    if ([bool]$o.Opened) {
        return [pscustomobject]@{ Unc = $Unc; State = 'Ok'; Opened = $true; Win32Error = 0
            Message = "This identity opened '$Unc' on the live host (not from cache). That is not proof that a driver will install or a page will print." }
    }
    [pscustomobject]@{ Unc = $Unc; State = 'Failed'; Opened = $false; Win32Error = [int]$o.Err
        Message = "Could not open '$Unc' on the live host: " + (Get-SmbErrorExplanation -Code ([int]$o.Err) -HostName $hostName) }
}

function Get-NetViewShares {
    <#
    .SYNOPSIS
        Shares listed by net view, WITH their type. Returns the error when refused.
    .DESCRIPTION
        The mapped-drive toolkit parses this with '^(.+?)\s{2,}Disk\b', which drops
        printer shares entirely - printer shares are typed Print. That filter would
        have made this whole feature silently return nothing for printers.
    .OUTPUTS
        PSCustomObject { Ok; Shares[]; Win32Error; Raw; Message }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$HostName, [int]$TimeoutSec = 20)

    $h = $HostName.TrimStart('\').Trim()
    $raw = Invoke-WithTimeout -TimeoutSec $TimeoutSec -Parameters @{ Srv = $h } -ScriptBlock {
        param($Srv)
        $ErrorActionPreference = 'Continue'
        $o = (& net.exe view "\\$Srv" 2>&1 | Out-String)
        [pscustomobject]@{ Out = $o; Code = $LASTEXITCODE }
    }
    if (-not $raw) {
        return [pscustomobject]@{ Ok=$false; Shares=@(); Win32Error=$null; Raw=''; Message="net view against $h did not return within $TimeoutSec s." }
    }

    $text = "$($raw.Out)"
    $err  = $null
    if ($text -match 'System error (\d+)') { $err = [int]$matches[1] }

    $shares = [System.Collections.Generic.List[object]]::new()
    foreach ($line in ($text -split "`r?`n")) {
        # Columns are: Share name <spaces> Type <spaces> Used as  Comment
        if ($line -match '^(.+?)\s{2,}(Disk|Print|IPC)\b\s*(.*)$') {
            $shares.Add([pscustomobject]@{
                Name    = $matches[1].Trim()
                Type    = $matches[2]
                Comment = $matches[3].Trim()
                Unc     = "\\$h\$($matches[1].Trim())"
            })
        }
    }

    $ok = ($shares.Count -gt 0 -and -not $err)
    [pscustomobject]@{
        Ok         = $ok
        Shares     = $shares.ToArray()
        Win32Error = $err
        Raw        = $text.Trim()
        Message    = $(if ($ok) { "Listed $($shares.Count) share(s)." }
                       elseif ($err) { "net view refused: " + (Get-SmbErrorExplanation -Code $err -HostName $h) }
                       else { 'net view returned no shares.' })
    }
}

function Connect-PrintHostSession {
    <#
    .SYNOPSIS
        Authenticate this logon session to a print host, so later calls run as
        that account.
    .DESCRIPTION
        Without this, everything runs as whatever identity the console already
        holds and typed credentials have no effect at all. Establishing an IPC$
        session makes net view, the spooler RPC and OpenPrinter all use the
        supplied account - SMB sessions are per LOGON SESSION, so it also covers
        the bounded runspaces those calls run in.

        WINDOWS ONLY ALLOWS ONE IDENTITY PER SERVER. If a session to the host
        already exists, Windows reuses it and silently ignores the credentials
        you just typed - so that case is REPORTED rather than papered over,
        because otherwise a failed test looks like bad credentials when it is
        really a stale session.

        The session is deliberately NOT torn down afterwards: dropping IPC$ out
        from under the spooler produced 0x80070040 on the very next operation.

        The password reaches net.exe on a command line, so it is briefly visible
        in the process list - the same trade-off already accepted for cmdkey. It
        is never logged, never returned, and never written to the report.
    .OUTPUTS
        PSCustomObject { Success; Created; AlreadyConnected; Notes[] }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server,
        [string]$User,
        [string]$Password
    )
    $h = $Server.TrimStart('\').Trim()
    $notes = [System.Collections.Generic.List[string]]::new()
    if (-not $User) {
        return [pscustomobject]@{ Success=$false; Created=$false; AlreadyConnected=$false; Notes=$notes.ToArray() }
    }

    $pre = @()
    try {
        $nu = (& net.exe use 2>&1 | Out-String)
        foreach ($line in ($nu -split "`r?`n")) {
            if ($line -match "(?i)\\\\$([regex]::Escape($h))\\") { $pre += $line.Trim() }
        }
    } catch {}

    if (@($pre).Count) {
        $notes.Add("A session to $h already exists, so the supplied credentials were NOT used - Windows reuses the existing session for every connection to a server. Use Repair to clear the host first if you need to prove a specific account.")
        return [pscustomobject]@{ Success=$false; Created=$false; AlreadyConnected=$true; Notes=$notes.ToArray() }
    }

    $pw = if ($Password) { $Password } else { '' }
    try {
        & net.exe use "\\$h\IPC$" "/user:$User" $pw *>$null
        if ($LASTEXITCODE -eq 0) {
            $notes.Add("Authenticated to $h as $User.")
            return [pscustomobject]@{ Success=$true; Created=$true; AlreadyConnected=$false; Notes=$notes.ToArray() }
        }
        $notes.Add("Could not authenticate to $h as $User - continuing as the console's own identity.")
    } catch {
        $notes.Add("Could not authenticate to $h as ${User}: $($_.Exception.Message)")
    }
    [pscustomobject]@{ Success=$false; Created=$false; AlreadyConnected=$false; Notes=$notes.ToArray() }
}

function Test-PrintShareAccess {
    <#
    .SYNOPSIS
        Three independent verdicts about a print host: reachable, enumerable, connectable.
    .DESCRIPTION
        Deliberately does NOT collapse these into one answer.

          Reachable   - TCP 445/135 (Test-PrintServerReachable)
          Enumerable  - can shares be LISTED? Tried over the spooler's RPC
                        (Get-Printer -ComputerName) and over SMB (net view),
                        recording which one worked. Enumeration is commonly
                        restricted to admins, and more so on a DC.
          Connectable - can a NAMED share be opened? (OpenPrinter) This is the
                        one that decides whether printing works.

        A host that is reachable, refuses enumeration with error 5, and still
        opens its printer share is a completely normal, working configuration.
    .OUTPUTS
        PSCustomObject { Server; Share; Reachable; CanEnumerate; EnumeratedVia;
                         EnumError; Shares[]; CanConnect; ConnectError;
                         Findings[]; Summary }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server,
        [string]$Share,
        [string]$User,
        [string]$Password,
        [int]$TimeoutSec = 20
    )

    $h = $Server.TrimStart('\').Trim()
    $notes = [System.Collections.Generic.List[string]]::new()
    $createdSession = $false

    # ---- 1. reachable ------------------------------------------------------
    $reach = Test-PrintServerReachable -Server $h
    if (-not $reach.Reachable) {
        return [pscustomobject]@{
            Server=$h; Share=$Share; Reachable=$false; CanEnumerate=$false; EnumeratedVia=''
            EnumError=$null; Shares=@(); CanConnect=$false; ConnectError=$null
            Findings=@("$h did not answer on SMB (445) or RPC (135). Nothing else can be assessed until it does.")
            Summary="$h is unreachable."
        }
    }

    # ---- 1b. authenticate first, if credentials were supplied --------------
    # Enumeration otherwise runs as whatever identity the console already holds,
    # so typing credentials had no effect on it whatsoever. Establishing an IPC$
    # session first makes both net view and the spooler RPC run as that account -
    # SMB sessions are per LOGON SESSION, so this covers the bounded runspace the
    # RPC call runs in as well.
    if ($User) {
        $conn = Connect-PrintHostSession -Server $h -User $User -Password $Password
        foreach ($nn in @($conn.Notes)) { $notes.Add($nn) }
        $createdSession = $conn.Created
    }

    # ---- 2. enumerable - RPC first, then SMB -------------------------------
    $viaRpc = @(Invoke-WithTimeout -TimeoutSec $TimeoutSec -Parameters @{ Srv = $h } -ScriptBlock {
        param($Srv)
        try { @(Get-Printer -ComputerName $Srv -ErrorAction Stop | Select-Object Name, ShareName, DriverName, PortName, Shared) } catch { @() }
    })
    $netView = Get-NetViewShares -HostName $h -TimeoutSec $TimeoutSec

    $canEnum = $false; $via = ''; $enumErr = $null
    $shares = @()
    if (@($viaRpc).Count) {
        $canEnum = $true; $via = 'spooler RPC (Get-Printer -ComputerName)'
        $shares = @($viaRpc | Where-Object { $_.Shared } | ForEach-Object {
            [pscustomobject]@{ Name=$_.ShareName; Type='Print'; Comment=$_.Name; Unc="\\$h\$($_.ShareName)" } })
    }
    if ($netView.Ok) {
        if (-not $canEnum) { $canEnum = $true; $via = 'SMB (net view)' } else { $via += ' and SMB (net view)' }
        # Printer shares only. net view lists disk and admin shares too, and
        # reporting those as printer shares made a PC that shares no printer look
        # like a print host.
        if (-not @($shares).Count) { $shares = @($netView.Shares | Where-Object { $_.Type -eq 'Print' }) }
    } else {
        $enumErr = $netView.Win32Error
    }

    # The headline insight: enumeration refused does not mean printing is broken.
    if (-not $netView.Ok -and $enumErr -eq 5) {
        if ($canEnum) {
            $notes.Add("net view was refused (error 5, access denied) but the shares WERE listed over $via. Listing shares needs rights that are usually admin-only, and is often restricted further on a domain controller. This does not stop anyone printing.")
        } else {
            $notes.Add('net view was refused with error 5 (access denied). That is a SHARE ENUMERATION right, usually admin-only - it does not by itself mean printing is blocked. Test a specific printer share by name to find out.')
        }
    }
    if (-not $canEnum -and $reach.Reachable) {
        $notes.Add("Shares could not be listed by either method. If you know the share name, test it directly - enumeration and connection are separate rights.")
    }

    # ---- 3. connectable - the question that actually matters ---------------
    $canConnect = $null; $connErr = $null
    if ($Share) {
        $unc = if ($Share -match '^\\\\') { $Share } else { "\\$h\$Share" }
        $open = Test-PrinterShareOpen -Unc $unc
        $canConnect = $open.Opened
        $connErr    = $open.Win32Error
        $notes.Add($open.Message)
        if ($open.Opened -and -not $canEnum) {
            $notes.Add('Note the combination: this host will not list its shares, but the printer share opens fine. That is a working configuration, not a fault.')
        }
    }

    # Deliberately NOT dropping the session this check created.
    #
    # Every connection to a server rides ONE SMB session, the spooler's included.
    # Tearing down IPC$ here pulled the connection out from under the very next
    # operation, and adding a printer straight after a check failed with
    # 0x80070040 / error 64, ERROR_NETNAME_DELETED. Windows tidies an idle IPC$
    # session up by itself; a leaked one is far cheaper than breaking the action
    # the technician is about to take.
    if ($createdSession) {
        $notes.Add("Left the authenticated session to $h in place - dropping it here would break the next operation against that host. Windows will retire it when idle, or use Repair to clear it deliberately.")
    }

    $summary = if ($canConnect -eq $true) { "$h reachable; printer share opens." }
               elseif ($canConnect -eq $false) { "$h reachable; the printer share would NOT open." }
               elseif ($canEnum) { "$h reachable and enumerable via $via." }
               else { "$h reachable but shares cannot be listed." }

    [pscustomobject]@{
        Server        = $h
        Share         = $Share
        Reachable     = $true
        CanEnumerate  = $canEnum
        EnumeratedVia = $via
        EnumError     = $enumErr
        Shares        = @($shares)
        CanConnect    = $canConnect
        ConnectError  = $connErr
        Findings      = $notes.ToArray()
        Summary       = $summary
    }
}

# endregion


# ---------------------------------------------------------------------------
# region  Credentials for print hosts
#
# Windows runs EVERY connection to a server over one authenticated SMB session,
# so a stored Credential Manager entry for a host decides the identity for
# everything on it - printers included. A leftover entry under a dead domain, or
# under the wrong account, produces access denied that re-adding the printer will
# never fix. That is what Repair-PrinterHostCredential is for.
#
# PASSWORD HANDLING - read before changing anything here:
#   * cmdkey /add /pass:<value> puts the password on the COMMAND LINE. It is
#     briefly visible in the process list and to process-creation auditing
#     (Sysmon, EDR). That is a deliberate, accepted trade-off.
#   * It is never logged, never written to the report, and never placed in the
#     JSON envelope.
#   * The temp script that carries it into the user's session is written to an
#     ACL'd directory, NOT %PUBLIC%\Documents like the other injectors - that
#     path is world-readable and would be strictly worse than the process list.
# ---------------------------------------------------------------------------

function Invoke-SecureScriptInUserSession {
    <#
    .SYNOPSIS
        Run a scriptblock in a user's session via a temp file only they can read.
    .DESCRIPTION
        Same one-shot-task mechanism as Invoke-ScriptInPrinterUserSession, with one
        difference that matters: the script is written under ProgramData with
        inheritance disabled and rights granted ONLY to SYSTEM, Administrators and
        the target user. The ordinary injector uses %PUBLIC%\Documents, which every
        user on the machine can read - fine for a directory listing, not for
        anything carrying a credential.

        The file is created empty, ACL'd, and only then written to, so the content
        is never briefly exposed under inherited permissions.
    .OUTPUTS
        PSCustomObject { Success; Data; Raw; Message; User }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [hashtable]$Vars,
        [Parameter(Mandatory)][string]$User,
        [int]$TimeoutSec = 45,
        [scriptblock]$OnPoll
    )

    $prefix = ''
    if ($Vars) {
        foreach ($k in $Vars.Keys) {
            $v = $Vars[$k]
            $lit = if ($v -is [bool]) { if ($v) { '$true' } else { '$false' } }
                   elseif ($v -is [int]) { "$v" }
                   else { "'" + (("$v") -replace "'", "''") + "'" }
            $prefix += "`$$k = $lit`r`n"
        }
    }
    $body = "`$ErrorActionPreference='SilentlyContinue'; `$ProgressPreference='SilentlyContinue'`r`n" + $prefix + $ScriptBlock.ToString()

    $dir = Join-Path $env:ProgramData 'PrinterToolkit'
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $ps1 = Join-Path $dir ("cred-{0}.ps1" -f [guid]::NewGuid().ToString('N').Substring(0, 8))

    try {
        New-Item -ItemType File -Path $ps1 -Force | Out-Null

        # Start from the file's OWN descriptor and protect it. Building a fresh
        # FileSecurity and calling SetOwner does not work: setting the owner needs
        # a privilege a standard user does not have, Set-Acl then throws, and
        # NOTHING is applied - leaving the file on inherited permissions with
        # BUILTIN\Users read access. Measured, not assumed.
        $acl = Get-Acl -LiteralPath $ps1
        $acl.SetAccessRuleProtection($true, $false)     # break inheritance, copy nothing
        foreach ($r in @($acl.Access)) { [void]$acl.RemoveAccessRule($r) }
        foreach ($who in @('NT AUTHORITY\SYSTEM', 'BUILTIN\Administrators', $User)) {
            try {
                $rights = if ($who -eq $User) { 'ReadAndExecute' } else { 'FullControl' }
                $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
                    $who, $rights, 'Allow')))
            } catch { }
        }
        Set-Acl -LiteralPath $ps1 -AclObject $acl -ErrorAction Stop

        # FAIL CLOSED. Confirm the file really is restricted before the password
        # touches it - if the ACL did not take, writing anyway would put plaintext
        # somewhere every user on the machine can read.
        $check = Get-Acl -LiteralPath $ps1
        $exposed = @($check.Access | Where-Object {
            "$($_.IdentityReference)" -match '(?i)\\Users$|Everyone|Authenticated Users|INTERACTIVE'
        })
        if (-not $check.AreAccessRulesProtected -or $exposed.Count) {
            throw "Refusing to continue: the temp script could not be restricted (inheritance protected=$($check.AreAccessRulesProtected), open to $(@($exposed | ForEach-Object { $_.IdentityReference }) -join ', ')). No credential was written to disk."
        }

        # Content only AFTER the ACL is verified.
        Set-Content -LiteralPath $ps1 -Value $body -Encoding UTF8

        $r = Invoke-CommandInPrinterUserSession -CommandLine ('powershell -NoProfile -ExecutionPolicy Bypass -File "' + $ps1 + '"') `
                -User $User -TimeoutSec $TimeoutSec -OnPoll $OnPoll
        $raw  = "$($r.Output)"
        $data = ConvertFrom-NoisyJson $raw
        return [pscustomobject]@{
            Success = [bool]$data; Data = $data; Raw = $raw.Trim()
            Message = $(if ($data) { 'ok (user session)' } else { $r.Message }); User = $User
        }
    } catch {
        return [pscustomobject]@{ Success=$false; Data=$null; Raw=$null; Message="Secure in-session run failed: $($_.Exception.Message)"; User=$User }
    } finally {
        # Overwrite before deleting - a plain delete leaves the content on disk.
        try { Set-Content -LiteralPath $ps1 -Value (' ' * 512) -Encoding ASCII -ErrorAction SilentlyContinue } catch {}
        Remove-Item -LiteralPath $ps1 -Force -ErrorAction SilentlyContinue
    }
}

function Test-PrinterHostCredential {
    <#
    .SYNOPSIS
        Test a username and password against a print host. Read-only.
    .DESCRIPTION
        Authenticates against \\host\IPC$ and decodes the result, then removes only
        the session it created.

        IMPORTANT: if a session to that host ALREADY exists, this declines to give
        a verdict. Windows would either reuse the existing session (a false pass)
        or return 1219 (a false fail), so any answer would be meaningless. Clear
        the host first with Repair-PrinterHostCredential.

        Runs in the CALLING context. From SYSTEM that still proves whether the
        username and password are valid for the host, which is what is being asked.
    .OUTPUTS
        PSCustomObject { Server; User; Valid; Win32Error; PreExisting; Message }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$User,
        [string]$Password
    )

    $ErrorActionPreference = 'Continue'
    $h   = $Server.TrimStart('\').Trim()
    $ipc = "\\$h\IPC$"

    # Existing session to this host?
    $existing = @()
    try {
        $nu = (& net.exe use 2>&1 | Out-String)
        foreach ($line in ($nu -split "`r?`n")) {
            if ($line -match "(?i)\\\\$([regex]::Escape($h))\\") { $existing += $line.Trim() }
        }
    } catch {}
    if (@($existing).Count) {
        return [pscustomobject]@{
            Server=$h; User=$User; Valid=$null; Win32Error=$null; PreExisting=$true
            Message="A connection to $h already exists, so an explicit credential test would be meaningless - Windows reuses the existing session (false pass) or returns 1219 (false fail). Use Repair to clear the host first, then test."
        }
    }

    $pw = if ($Password) { $Password } else { '' }
    $out = ''; $code = 0
    try {
        # Argument ARRAY, not an interpolated string - nothing goes through a cmd
        # parser. The password is still on the command line; see the region note.
        $out  = (& net.exe use $ipc "/user:$User" $pw 2>&1 | Out-String)
        $code = $LASTEXITCODE
    } catch {
        $out = "$($_.Exception.Message)"; $code = -1
    }

    $err = 0
    if ($out -match 'System error (\d+)') { $err = [int]$matches[1] }
    $valid = ($code -eq 0)

    # Remove ONLY what this test created.
    if ($valid) { try { & net.exe use $ipc /delete /y *>$null } catch {} }

    [pscustomobject]@{
        Server = $h; User = $User; Valid = $valid; Win32Error = $err; PreExisting = $false
        Message = $(if ($valid) { "Credentials accepted by $h." } else { Get-SmbErrorExplanation -Code $err -Text $out -HostName $h })
    }
}

function Add-PrinterViaLocalPort {
    <#
    .SYNOPSIS
        Add a shared printer as a LOCAL queue on a Local Port - no Point and Print.
    .DESCRIPTION
        A workaround when a printer connection cannot be made. Instead of a printer
        CONNECTION, this creates a Local Port named after the UNC and binds a local
        queue to it with a driver already installed on this PC, so Windows never
        downloads a driver from the host.

        What it does NOT do: sidestep authentication. When the queue prints, the
        Print Spooler writes to the share as this PC's MACHINE identity
        ($env:COMPUTERNAME$). In a workgroup the host generally cannot authenticate
        that identity, whatever the share permissions say, and jobs can then fail
        silently. Creating the queue is not proof that it prints - run the delivery test.

        The driver must already be installed here. The queue is machine-wide.
        A port this function creates is removed again if the queue cannot be made,
        and an existing port of the same name is reused only if it is a Local Port.
    .OUTPUTS
        PSCustomObject { Success; Name; PortName; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$ConnectionName,
        [Parameter(Mandatory)][string]$DriverName,
        [string]$Name
    )

    $unc = $ConnectionName.Trim().TrimEnd('\')
    if ($unc -notmatch '^\\\\[^\\]+\\.+') {
        return [pscustomobject]@{ Success = $false; Name = ''; PortName = ''; Message = "'$unc' is not a \\server\share path." }
    }
    if (-not $Name) { $Name = ($unc -split '\\')[-1] }

    if (-not (Test-PrintToolkitElevated)) {
        return [pscustomobject]@{ Success = $false; Name = $Name; PortName = $unc; Message = 'Creating a local queue needs elevation. Run as Administrator, or as SYSTEM from Backstage.' }
    }
    if (-not (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue)) {
        return [pscustomobject]@{ Success = $false; Name = $Name; PortName = $unc
            Message = "Driver '$DriverName' is not installed on this PC. This method never fetches a driver from the host, so it must be here already. Install it first (Shares -> Install driver...)." }
    }
    if (Get-QueueExact $Name) {
        return [pscustomobject]@{ Success = $false; Name = $Name; PortName = $unc; Message = "A queue named '$Name' already exists - remove it or choose another name." }
    }
    $existingPort = Get-PrinterPort -Name $unc -ErrorAction SilentlyContinue
    if ($existingPort) {
        $mon = "$($existingPort.PortMonitor) $($existingPort.Description)"
        if ($mon -notmatch '(?i)local') {
            return [pscustomobject]@{ Success = $false; Name = $Name; PortName = $unc
                Message = "A port named '$unc' already exists but is not a Local Port ($mon). Refusing to bind a queue to it." }
        }
    }
    if (-not $PSCmdlet.ShouldProcess($unc, "Add local queue '$Name' on a local port")) {
        return [pscustomobject]@{ Success = $false; Name = $Name; PortName = $unc; Message = 'Skipped (WhatIf).' }
    }

    $createdPort = $false
    try {
        if (-not $existingPort) {
            Add-PrinterPort -Name $unc -ErrorAction Stop
            $createdPort = $true
        }
        Add-Printer -Name $Name -DriverName $DriverName -PortName $unc -ErrorAction Stop
    } catch {
        $cleanup = ''
        if ($createdPort) {
            try { Remove-PrinterPort -Name $unc -ErrorAction Stop; $cleanup = ' The port created for it was removed again.' }
            catch { $cleanup = " The port created for it could not be removed: $($_.Exception.Message)" }
        }
        return [pscustomobject]@{ Success = $false; Name = $Name; PortName = $unc; Message = "Could not create the local queue: $($_.Exception.Message)$cleanup" }
    }

    $q = Get-QueueExact $Name
    [pscustomobject]@{
        Success  = [bool]$q
        Name     = $Name
        PortName = $unc
        Message  = $(if ($q) {
                        "Created local queue '$Name' printing to $unc with the local '$DriverName' driver; every user of this PC gets it. " +
                        "CREATED IS NOT PRINTING: jobs reach the share as this PC's machine identity ($env:COMPUTERNAME`$), not the signed-in user. " +
                        'In a workgroup the host generally cannot authenticate that identity, whatever the share permissions say. Run the delivery test before relying on it.'
                     } else { 'The queue was not present after creation - check the spooler.' })
    }
}

function Save-PrinterHostCredential {
    <#
    .SYNOPSIS
        Store a credential for a print host in the TARGET USER's Credential Manager.
    .DESCRIPTION
        Injected into the user's session, because a credential saved by SYSTEM
        lands in SYSTEM's vault and does nothing whatsoever for the user.

        Saves under both the name and IP forms of the host when both are known, so
        a connection made by either route authenticates. Verified afterwards with
        cmdkey /list rather than trusting the exit code.
    .OUTPUTS
        PSCustomObject { Success; Server; User; TargetUser; Saved[]; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$User,
        [string]$Password,
        [Parameter(Mandatory)][string]$TargetUser,
        [switch]$BothForms,
        [scriptblock]$OnPoll
    )

    $h = $Server.TrimStart('\').Trim()
    if (-not $PSCmdlet.ShouldProcess("$h for $TargetUser", 'Save credential')) {
        return [pscustomobject]@{ Success=$false; Server=$h; User=$User; TargetUser=$TargetUser; Saved=@(); Message='Skipped (WhatIf).' }
    }

    $forms = if ($BothForms) { @(Get-PrinterHostForms -HostName $h) } else { @($h) }

    $work = {
        $saved = @()
        foreach ($f in ($FormList -split ';')) {
            if (-not $f) { continue }
            & cmdkey.exe /add:$f /user:$CredUser /pass:$CredPass *>$null
            $saved += $f
        }
        # Verify from the vault rather than trusting the exit code.
        $list = (& cmdkey.exe /list 2>&1 | Out-String)
        $present = @()
        foreach ($f in ($FormList -split ';')) {
            if ($f -and $list -match [regex]::Escape("target=$f")) { $present += $f }
        }
        [pscustomobject]@{ Attempted = $saved; Present = $present; User = $env:USERNAME } | ConvertTo-Json -Compress
    }

    $r = Invoke-SecureScriptInUserSession -ScriptBlock $work -User $TargetUser -OnPoll $OnPoll -TimeoutSec 60 -Vars @{
        FormList = ($forms -join ';')
        CredUser = $User
        CredPass = $Password
    }

    if ($r.Success -and @($r.Data.Present).Count) {
        return [pscustomobject]@{
            Success=$true; Server=$h; User=$User; TargetUser=$TargetUser; Saved=@($r.Data.Present)
            Message="Saved the credential for $User in $TargetUser's Credential Manager, under: $(@($r.Data.Present) -join ', '). It now applies to every connection they make to that host."
        }
    }
    [pscustomobject]@{
        Success=$false; Server=$h; User=$User; TargetUser=$TargetUser; Saved=@()
        Message="The credential could not be confirmed in $TargetUser's vault. $($r.Message)"
    }
}

function Repair-PrinterHostCredential {
    <#
    .SYNOPSIS
        Clear stale saved credentials and sessions for a print host, then verify.
    .DESCRIPTION
        The fix for the commonest access-denied cause: a stored credential for the
        host under a dead domain or the wrong account. Because every connection to
        one server rides a single authenticated session, ONE surviving connection
        keeps the old identity alive for all of them - so this reports survivors
        rather than assuming success.

        Clears both the name and IP forms: a credential cached under the other form
        would otherwise silently take over.

        Runs in the user's session - SYSTEM has its own vault and its own
        connections, and clearing those fixes nothing for the user.
    .OUTPUTS
        PSCustomObject { Success; Server; Forms[]; Removed[]; Survivors[]; Log[]; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$TargetUser,
        [scriptblock]$OnPoll
    )

    $h = $Server.TrimStart('\').Trim()
    if (-not $PSCmdlet.ShouldProcess("$h for $TargetUser", 'Clear saved credentials and sessions')) {
        return [pscustomobject]@{ Success=$false; Server=$h; Forms=@(); Removed=@(); Survivors=@(); Log=@(); Message='Skipped (WhatIf).' }
    }
    $forms = @(Get-PrinterHostForms -HostName $h)

    $work = {
        $log = New-Object System.Collections.Generic.List[string]
        $removed = @()
        $forms = @()
        foreach ($f in ($FormList -split ';')) { if ($f) { $forms += $f } }

        foreach ($f in $forms) {
            $o = (& cmdkey.exe /delete:$f 2>&1 | Out-String)
            # Exit code, not the English string - 'deleted successfully' does not
            # appear on a localised Windows.
            if ($LASTEXITCODE -eq 0) { $removed += $f; $log.Add("Removed stored credential for $f") }
            else { $log.Add("No stored credential for $f") }
        }
        foreach ($f in $forms) {
            & net.exe use "\\$f\*" /delete /y *>$null
            $log.Add("Dropped connections to \\$f")
        }

        $cur = (& net.exe use 2>&1 | Out-String)
        $surv = @()
        foreach ($line in ($cur -split "`r?`n")) {
            foreach ($f in $forms) {
                if ($line -match "(?i)\\\\$([regex]::Escape($f))\\") { $surv += $line.Trim(); break }
            }
        }
        [pscustomobject]@{ Removed = $removed; Survivors = $surv; Log = $log.ToArray(); User = $env:USERNAME } | ConvertTo-Json -Depth 4 -Compress
    }

    $r = Invoke-ScriptInPrinterUserSession -ScriptBlock $work -User $TargetUser -OnPoll $OnPoll -TimeoutSec 60 -Vars @{
        FormList = ($forms -join ';')
    }

    if (-not $r.Success) {
        return [pscustomobject]@{ Success=$false; Server=$h; Forms=$forms; Removed=@(); Survivors=@(); Log=@(); Message="Could not run the repair in $TargetUser's session. $($r.Message)" }
    }

    $surv = @($r.Data.Survivors)
    $msg = if ($surv.Count) {
        "Cleared what it could, but $($surv.Count) connection(s) to $h SURVIVED - the old identity is still live for every share on that host. Sign the user out and back in, or reboot, before retrying."
    } else {
        "Cleared saved credentials and connections for $h (forms: $($forms -join ', ')). The next connection will authenticate cleanly."
    }
    [pscustomobject]@{
        Success   = ($surv.Count -eq 0)
        Server    = $h
        Forms     = $forms
        Removed   = @($r.Data.Removed)
        Survivors = $surv
        Log       = @($r.Data.Log)
        Message   = $msg
    }
}

# endregion


# ---------------------------------------------------------------------------
# region  Safe host-side fixes
#
# Only changes that are NOT security downgrades live here. Granting Print on a
# printer, enabling the File and Printer Sharing firewall rules, and setting the
# LAN to Private are all things a print host is supposed to have.
#
# Deliberately absent, and staying absent: AllowInsecureGuestAuth, SMB signing,
# RestrictDriverInstallationToAdministrators and RpcAuthnLevelPrivacyEnabled.
# Those are the settings whose "fix" is a security regression, and the toolkit's
# job is to name them, not to weaken them.
# ---------------------------------------------------------------------------

function Grant-PrinterPrintAccess {
    <#
    .SYNOPSIS
        Grant Print rights on a shared printer to a group. Adds, never replaces.
    .DESCRIPTION
        Works on the printer's SDDL via MSFT_Printer's PermissionSDDL, which is far
        less error-prone than rebuilding a security descriptor by hand.

        In printer SDDL, Print is (A;;SWRC;;;<sid>) - SW is PRINTER_ACCESS_USE and
        RC is READ_CONTROL. The existing DACL is preserved and the new ACE appended,
        so nothing already granted is lost.

        Default trustee is Authenticated Users (AU) rather than Everyone (WD):
        it covers every account that has actually logged on, without extending to
        anonymous.
    .OUTPUTS
        PSCustomObject { Success; Printer; Trustee; Before; After; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Name,
        [ValidateSet('AuthenticatedUsers', 'Everyone', 'Users')][string]$Trustee = 'AuthenticatedUsers'
    )

    $sid = switch ($Trustee) {
        'Everyone'           { 'WD' }
        'Users'              { 'BU' }
        default              { 'AU' }
    }

    # -Full is required: a plain Get-Printer does NOT populate PermissionSDDL, and
    # reading it from an ordinary Get-Printer object silently yields an empty
    # string. Filter client-side because -Name treats [ ] as wildcards.
    $q = $null
    try { $q = Get-Printer -Full -ErrorAction Stop | Where-Object { $_.Name -eq $Name } | Select-Object -First 1 } catch {}
    if (-not $q) { $q = Get-QueueExact $Name }
    if (-not $q) { return [pscustomobject]@{ Success=$false; Printer=$Name; Trustee=$Trustee; Before=''; After=''; Message="No queue named '$Name'." } }

    $before = ''
    try { $before = [string]$q.PermissionSDDL } catch {}
    if (-not $before) {
        return [pscustomobject]@{ Success=$false; Printer=$Name; Trustee=$Trustee; Before=''; After=''
                                  Message="Could not read the permissions for '$Name' (PermissionSDDL was empty). Set them from the printer's Security tab instead." }
    }
    $ace = "(A;;SWRC;;;$sid)"
    if ($before -like "*$ace*") {
        return [pscustomobject]@{ Success=$true; Printer=$Name; Trustee=$Trustee; Before=$before; After=$before
                                  Message="'$Name' already grants Print to $Trustee - nothing to change." }
    }
    if (-not $PSCmdlet.ShouldProcess($Name, "Grant Print to $Trustee")) {
        return [pscustomobject]@{ Success=$false; Printer=$Name; Trustee=$Trustee; Before=$before; After=''; Message='Skipped (WhatIf).' }
    }

    $after = $before + $ace
    try {
        Set-Printer -Name $Name -PermissionSDDL $after -ErrorAction Stop
        $now = ''
        try { $now = [string](Get-Printer -Full -ErrorAction Stop | Where-Object { $_.Name -eq $Name } | Select-Object -First 1).PermissionSDDL } catch {}
        $ok = ($now -like "*$ace*")
        return [pscustomobject]@{
            Success=$ok; Printer=$Name; Trustee=$Trustee; Before=$before; After=$now
            Message=$(if ($ok) { "Granted Print on '$Name' to $Trustee. Existing permissions were left intact." }
                      else { "Set-Printer accepted the change but the ACE is not present afterwards - check the Security tab." })
        }
    } catch {
        return [pscustomobject]@{ Success=$false; Printer=$Name; Trustee=$Trustee; Before=$before; After=''
                                  Message="Could not set permissions on '$Name': $($_.Exception.Message)" }
    }
}

function Enable-PrintSharingFirewall {
    <#
    .SYNOPSIS
        Enable the SCOPED built-in inbound rules a print host needs, by Name.
    .DESCRIPTION
        Selected by locale-independent rule Name, not the English display group.
        Enables the scoped variants for the active profile only:

          Domain            -> FPS-*-NoScope
          Private / Public  -> FPS-*   (ship scoped to LocalSubnet)

        for SMB 445, NetBIOS session 139, the RPC endpoint mapper 135 and the Print
        Spooler's RPC traffic.

        It deliberately does NOT enable the -V2 variants, which admit any remote
        address on every profile. The previous version enabled the whole display
        group, -V2 included - a broader opening than sharing a printer needs.

        Refuses while an adapter is on the Public profile: fix the profile (SHR001)
        rather than opening rules on a public network.
    .OUTPUTS
        PSCustomObject { Success; Enabled; Rules[]; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([string]$Profile)

    if (-not $PSCmdlet.ShouldProcess('File and Printer Sharing', 'Enable scoped inbound firewall rules')) {
        return [pscustomobject]@{ Success = $false; Enabled = 0; Rules = @(); Message = 'Skipped (WhatIf).' }
    }
    $active = if ($Profile) { @($Profile) } else { @(Get-ActiveFirewallProfile) }
    if (-not $active.Count) {
        return [pscustomobject]@{ Success = $false; Enabled = 0; Rules = @(); Message = 'No active network profile could be read, so the rules to enable could not be chosen.' }
    }
    if ($active -contains 'Public') {
        return [pscustomobject]@{ Success = $false; Enabled = 0; Rules = @()
            Message = 'An adapter is on the Public profile. Set the network to Private first (SHR001) - do not open sharing rules on a public network.' }
    }

    $want = [System.Collections.Generic.List[string]]::new()
    foreach ($base in @('FPS-SMB-In-TCP', 'FPS-NB_Session-In-TCP', 'FPS-RPCSS-In-TCP', 'FPS-SpoolSvc-In-TCP')) {
        if ($active -contains 'Domain') { $want.Add("$base-NoScope") }
        if ($active -contains 'Private') { $want.Add($base) }
    }
    $done = [System.Collections.Generic.List[string]]::new()
    $problems = [System.Collections.Generic.List[string]]::new()
    foreach ($n in $want) {
        try {
            if (-not (Get-NetFirewallRule -Name $n -ErrorAction SilentlyContinue)) { $problems.Add("$n not present"); continue }
            Enable-NetFirewallRule -Name $n -ErrorAction Stop
            $done.Add($n)
        } catch { $problems.Add("${n}: $($_.Exception.Message)") }
    }
    [pscustomobject]@{
        Success = ($done.Count -gt 0 -and $problems.Count -eq 0)
        Enabled = $done.Count
        Rules   = $done.ToArray()
        Message = ("Enabled the scoped File and Printer Sharing rules for the $($active -join ', ') profile: " + $(if ($done.Count) { $done -join ', ' } else { 'none' }) + '.' +
                   $(if ($problems.Count) { ' Problems: ' + ($problems -join '; ') + '.' } else { '' }) +
                   ' The any-address -V2 variants were deliberately left alone.')
    }
}

function Set-PrintSharingNetworkPrivate {
    <#
    .SYNOPSIS
        Move Public network adapters to Private so inbound sharing is permitted.
    .DESCRIPTION
        Windows blocks the inbound sharing rules on a Public profile, so a print
        host on Public is unreachable no matter how the share is configured.
        Only adapters currently on Public are touched; Domain stays Domain.
    .OUTPUTS
        PSCustomObject { Success; Changed[]; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param()

    try {
        $pub = @(Get-NetConnectionProfile -ErrorAction Stop | Where-Object { $_.NetworkCategory -eq 'Public' })
    } catch {
        return [pscustomobject]@{ Success=$false; Changed=@(); Message="Could not read the network profiles: $($_.Exception.Message)" }
    }
    if (-not $pub.Count) {
        return [pscustomobject]@{ Success=$true; Changed=@(); Message='No adapter is on the Public profile - nothing to change.' }
    }
    if (-not $PSCmdlet.ShouldProcess(($pub.InterfaceAlias -join ', '), 'Set network profile to Private')) {
        return [pscustomobject]@{ Success=$false; Changed=@(); Message='Skipped (WhatIf).' }
    }
    $changed = @()
    foreach ($p in $pub) {
        try {
            Set-NetConnectionProfile -InterfaceIndex $p.InterfaceIndex -NetworkCategory Private -ErrorAction Stop
            $changed += $p.InterfaceAlias
        } catch { }
    }
    [pscustomobject]@{
        Success = ($changed.Count -gt 0); Changed = $changed
        Message = $(if ($changed.Count) { "Set to Private: $($changed -join ', '). Inbound sharing rules now apply." }
                    else { 'The network profile could not be changed - this needs elevation.' })
    }
}

# endregion


function Get-RemotePrinterDriverInfo {
    <#
    .SYNOPSIS
        What driver does a print host use, and where do you get it from?
    .DESCRIPTION
        Answers "which INF do I need?" without guessing. The host already knows:
        Get-PrinterDriver -ComputerName returns the driver name, its INF and its
        version, over the same spooler RPC that Get-Printer uses and with no admin
        rights needed.

        Also probes the host's print$ share, which for v3 drivers holds the actual
        driver files - if it is readable, the package can often be copied straight
        from there. v4 drivers live in the driver store instead and have to be
        exported with pnputil, so the exact command to run ON the host is included.

        The InfPath returned is the HOST's path, not something reachable from here;
        it identifies the package (the oem<n>.inf number) so pnputil can export it.
    .OUTPUTS
        PSCustomObject { Server; Share; DriverName; InfPath; InfName; MajorVersion;
                         Manufacturer; PackageAware; PrintShareReadable; DriverPaths[];
                         ExportCommands[]; Message }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server,
        [string]$ShareName,
        [int]$TimeoutSec = 25
    )

    $h = $Server.TrimStart('\').Trim()

    # Which driver does that share use?
    $wantDriver = ''
    $printers = @(Invoke-WithTimeout -TimeoutSec $TimeoutSec -Parameters @{ Srv = $h } -ScriptBlock {
        param($Srv)
        try { @(Get-Printer -ComputerName $Srv -ErrorAction Stop | Select-Object Name, ShareName, DriverName, Shared) } catch { @() }
    })
    if ($ShareName) {
        $hit = @($printers | Where-Object { $_.ShareName -ieq $ShareName -or $_.Name -ieq $ShareName }) | Select-Object -First 1
        if ($hit) { $wantDriver = [string]$hit.DriverName }
    }

    # Driver detail from the host.
    $drivers = @(Invoke-WithTimeout -TimeoutSec $TimeoutSec -Parameters @{ Srv = $h } -ScriptBlock {
        param($Srv)
        try { @(Get-PrinterDriver -ComputerName $Srv -ErrorAction Stop |
                Select-Object Name, InfPath, MajorVersion, Manufacturer, IsPackageAware, PrinterEnvironment) } catch { @() }
    })

    $d = $null
    if ($wantDriver) { $d = @($drivers | Where-Object { $_.Name -ieq $wantDriver }) | Select-Object -First 1 }
    if (-not $d -and @($drivers).Count -eq 1) { $d = @($drivers)[0] }

    $infPath = ''; $infName = ''
    if ($d) {
        $infPath = [string]$d.InfPath
        if ($infPath) { $infName = Split-Path $infPath -Leaf }
    }

    # print$ holds v3 driver files; worth knowing whether it can just be copied.
    $printShare = $false
    $paths = [System.Collections.Generic.List[string]]::new()
    try {
        $printShare = [bool](Test-Path -LiteralPath "\\$h\print$" -ErrorAction SilentlyContinue)
        if ($printShare) {
            foreach ($arch in @('x64', 'amd64', 'w32x86')) {
                $p = "\\$h\print$\$arch\3"
                if (Test-Path -LiteralPath $p -ErrorAction SilentlyContinue) { $paths.Add($p) }
            }
        }
    } catch {}

    # pnputil works on the PUBLISHED name (oem<n>.inf), while the spooler reports
    # the ORIGINAL name from the driver store (e.g. hpcad92a4_x64.inf). Knowing the
    # original lets the tech find the oem number in one grep instead of reading a
    # list of every printer driver on the machine.
    $cmds = [System.Collections.Generic.List[string]]::new()
    if ($infName -match '^oem\d+\.inf$') {
        $cmds.Add("pnputil /export-driver $infName C:\Temp\PrinterDriver")
    } elseif ($infName) {
        $cmds.Add("pnputil /enum-drivers /class Printer | findstr /i `"$infName`"     :: gives its oem<n>.inf")
        $cmds.Add('pnputil /export-driver oem<n>.inf C:\Temp\PrinterDriver')
    } else {
        $cmds.Add('pnputil /enum-drivers /class Printer          :: find the oem<n>.inf for this model')
        $cmds.Add('pnputil /export-driver oem<n>.inf C:\Temp\PrinterDriver')
    }

    $msg = if (-not @($drivers).Count) {
        "Could not read the driver list from $h. The spooler RPC path may be blocked - the same thing that stops a printer connection opening."
    } elseif ($d) {
        "$h uses '$($d.Name)' (driver model v$($d.MajorVersion), package-aware=$($d.IsPackageAware)). Its INF on that machine is $infPath"
    } else {
        "Read $(@($drivers).Count) driver(s) from $h but could not match one to '$ShareName' - pick from the list."
    }

    [pscustomobject]@{
        Server             = $h
        Share              = $ShareName
        DriverName         = $(if ($d) { [string]$d.Name } else { '' })
        InfPath            = $infPath
        InfName            = $infName
        MajorVersion       = $(if ($d) { $d.MajorVersion } else { $null })
        Manufacturer       = $(if ($d) { [string]$d.Manufacturer } else { '' })
        PackageAware       = $(if ($d) { $d.IsPackageAware } else { $null })
        Environment        = $(if ($d) { [string]$d.PrinterEnvironment } else { '' })
        AllDrivers         = @($drivers)
        PrintShareReadable = $printShare
        DriverPaths        = $paths.ToArray()
        ExportCommands     = $cmds.ToArray()
        Message            = $msg
    }
}


function Get-PrintRpcConfig {
    <#
    .SYNOPSIS
        Print RPC policy on this PC. Tri-state and read-only.
    .DESCRIPTION
        The settings Microsoft documents for Windows 11 22H2+ print RPC (Learn,
        "Windows 11 RPC connection updates for print"), plus the separate
        CVE-2021-1678 privacy switch.

        RPC over TCP is the default and the recommended configuration. RpcTcpPort
        pins the spooler's listener to one port so it can be measured and
        firewalled precisely. Named pipes are documented but "not recommended".

        Configuration is not activation: a value here may not be in effect until
        the spooler restarts. Get-PrintSpoolerListener shows what is live.
    .OUTPUTS
        PSCustomObject { Values[] } - each { Name; Path; Role; Value; IsSet; EffectiveValue; State; Meaning }
    #>
    [CmdletBinding()] param()

    $rpcPolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC'
    $defs = @(
        @{ Name = 'RpcTcpPort'; Path = $rpcPolicy; Default = $null; Role = 'Host'
           Meaning = 'Fixed TCP port for the spooler RPC listener. Not set = a dynamic port in 49152-65535, which a client cannot test for.' }
        @{ Name = 'RpcProtocols'; Path = $rpcPolicy; Default = $null; Role = 'Host'
           Meaning = 'Protocols the spooler listens on. Not set = TCP only (Windows 11 22H2+ default). 0x7 = named pipes and TCP (not recommended).' }
        @{ Name = 'RpcUseNamedPipeProtocol'; Path = $rpcPolicy; Default = 0; Role = 'Client'
           Meaning = '0 or not set = connect over TCP (default, recommended). 1 = connect over named pipes (Microsoft: not recommended).' }
        @{ Name = 'ForceKerberosForRpc'; Path = $rpcPolicy; Default = 0; Role = 'Both'
           Meaning = '1 = require Kerberos for print RPC. Domain-only: it breaks workgroup printing. Reported, never written by this toolkit.' }
        @{ Name = 'RpcAuthnLevelPrivacyEnabled'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print'; Default = 1; Role = 'Both'
           Meaning = 'RPC packet privacy (CVE-2021-1678). Absent = ENFORCED. Never turn this off as a first move.' }
    )
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($d in $defs) {
        $val = $null; $isSet = $false
        try { $val = (Get-ItemProperty -Path $d.Path -Name $d.Name -ErrorAction Stop).($d.Name); $isSet = $true } catch {}
        $eff = $(if ($isSet) { $val } else { $d.Default })
        $rows.Add([pscustomobject]@{
            Name           = $d.Name
            Path           = $d.Path
            Role           = $d.Role
            Value          = $val
            IsSet          = $isSet
            EffectiveValue = $eff
            State          = $(if ($isSet) { "Set=$val" } elseif ($null -ne $d.Default) { "NotSet -> effective default $($d.Default)" } else { 'NotSet' })
            Meaning        = $d.Meaning
        })
    }
    [pscustomobject]@{ Values = $rows.ToArray() }
}

function Set-PrintRpcNamedPipeFallback {
    <#
    .SYNOPSIS
        ADVANCED, NOT RECOMMENDED: make this PC connect to print hosts over RPC named pipes.
    .DESCRIPTION
        Microsoft on RPC over named pipes for printing: "This configuration isn't
        recommended. However, it can be used if RPC over TCP isn't an option." The
        same page notes it may also need RpcAuthnLevelPrivacyEnabled=0 on the host and
        insecure guest SMB logons - both security downgrades this toolkit will not make.

        Prefer the recommended path: on the host, pin the spooler port and allow it
        through the firewall (Set-PrintRpcTcpPort, Enable-PrintRpcFirewall), then test
        from here (Test-PrintRpcPath).

        On its own this does nothing unless the host also listens on named pipes
        (RpcProtocols = 7, Set-PrintRpcListenerProtocol).

        -Off REMOVES the value, returning this PC to the Windows default (TCP). It no
        longer writes 0. The prior state is snapshotted, so Restore-PrintConfigSnapshot
        can return it exactly. Success means the value reads back AND the spooler
        restarted - it does not prove any connection now works.
    .OUTPUTS
        PSCustomObject { Success; Value; IsSet; RestartSucceeded; SnapshotPath; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$Off,
        [switch]$SkipSpoolerRestart
    )
    $path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC'
    if (-not (Test-PrintToolkitElevated)) {
        return [pscustomobject]@{ Success = $false; Value = $null; IsSet = $null; RestartSucceeded = $null; SnapshotPath = ''
            Message = 'Changing the print RPC protocol needs elevation. Run as Administrator, or as SYSTEM from Backstage.' }
    }
    if (-not $PSCmdlet.ShouldProcess('Print RPC client', $(if ($Off) { 'Remove RpcUseNamedPipeProtocol (Windows default: TCP)' } else { 'Set RpcUseNamedPipeProtocol = 1 (not recommended)' }))) {
        return [pscustomobject]@{ Success = $false; Value = $null; IsSet = $null; RestartSucceeded = $null; SnapshotPath = ''; Message = 'Skipped (WhatIf).' }
    }

    $snap = Save-PrintConfigSnapshot -Label $(if ($Off) { 'namedpipe-client-off' } else { 'namedpipe-client-on' }) -RegistryName @('RpcUseNamedPipeProtocol')
    try {
        if ($Off) {
            if (Test-Path -LiteralPath $path) { Remove-ItemProperty -Path $path -Name 'RpcUseNamedPipeProtocol' -Force -ErrorAction SilentlyContinue }
        } else {
            if (-not (Test-Path -LiteralPath $path)) { New-Item -Path $path -Force | Out-Null }
            Set-ItemProperty -Path $path -Name 'RpcUseNamedPipeProtocol' -Value 1 -Type DWord -ErrorAction Stop
        }
    } catch {
        return [pscustomobject]@{ Success = $false; Value = $null; IsSet = $null; RestartSucceeded = $null; SnapshotPath = $snap.Path
            Message = "Could not write the RPC setting: $($_.Exception.Message)" }
    }

    $restartOk = $null
    if (-not $SkipSpoolerRestart) { $restartOk = [bool](Restart-PrintSpooler).Success }

    $now = $null; $isSet = $false
    try { $now = (Get-ItemProperty -Path $path -Name 'RpcUseNamedPipeProtocol' -ErrorAction Stop).RpcUseNamedPipeProtocol; $isSet = $true } catch {}
    $valueOk = if ($Off) { -not $isSet } else { "$now" -eq '1' }

    $restartTxt = if ($null -eq $restartOk) { ' The spooler was NOT restarted, so this is not in effect yet.' }
                  elseif ($restartOk) { ' Spooler restarted.' }
                  else { ' The spooler did NOT restart - this is not in effect.' }
    $msg = if (-not $valueOk) {
               $(if ($Off) { "RpcUseNamedPipeProtocol could not be removed - it still reads '$now'. Something is re-applying it, most likely Group Policy." }
                 else { "RpcUseNamedPipeProtocol did not stick - it reads '$now'. Something is re-applying it, most likely Group Policy." })
           } elseif ($Off) {
               "RpcUseNamedPipeProtocol removed - this PC connects over TCP again (the Windows default).$restartTxt"
           } else {
               "RpcUseNamedPipeProtocol = 1 written.$restartTxt Microsoft does not recommend this configuration. It changes nothing unless the host also listens on named pipes, and it proves nothing until a connection is tested (Shares -> Test connection)."
           }
    [pscustomobject]@{
        Success          = ($valueOk -and $restartOk -ne $false -and -not ($null -eq $restartOk -and -not $SkipSpoolerRestart))
        Value            = $now
        IsSet            = $isSet
        RestartSucceeded = $restartOk
        SnapshotPath     = $snap.Path
        Message          = $msg
    }
}

function Set-PrintRpcListenerProtocol {
    <#
    .SYNOPSIS
        ADVANCED, NOT RECOMMENDED: make this host's spooler also listen on RPC named pipes.
    .DESCRIPTION
        Microsoft documents RpcProtocols = 0x7 ("RpcOverNamedPipesAndTcp") but
        describes RPC over named pipes as "not recommended", and notes it may also need
        RpcAuthnLevelPrivacyEnabled=0 and insecure guest SMB logons - security
        downgrades this toolkit will not make.

        Prefer the recommended path on this host: Set-PrintRpcTcpPort to pin the
        spooler, then Enable-PrintRpcFirewall.

        0x7 adds named pipes alongside TCP, so clients already working over TCP keep
        working. -Reset removes the value (Windows default: TCP only). The prior state
        is snapshotted. Success means the value reads back AND the spooler restarted -
        it does not prove any client can now connect.
    .PARAMETER Protocols
        Defaults to 7. Lower values narrow what the listener accepts and can cut off
        clients that currently work over TCP.
    .OUTPUTS
        PSCustomObject { Success; Value; IsSet; RestartSucceeded; SnapshotPath; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [ValidateRange(1, 7)][int]$Protocols = 7,
        [switch]$Reset,
        [switch]$SkipSpoolerRestart
    )
    $path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC'
    if (-not (Test-PrintToolkitElevated)) {
        return [pscustomobject]@{ Success = $false; Value = $null; IsSet = $null; RestartSucceeded = $null; SnapshotPath = ''
            Message = 'Changing the print RPC listener needs elevation. Run as Administrator, or as SYSTEM from Backstage.' }
    }
    $what = if ($Reset) { 'Remove RpcProtocols (Windows default: TCP only)' } else { "Set RpcProtocols = $Protocols (not recommended)" }
    if (-not $PSCmdlet.ShouldProcess('Print RPC listener', $what)) {
        return [pscustomobject]@{ Success = $false; Value = $null; IsSet = $null; RestartSucceeded = $null; SnapshotPath = ''; Message = 'Skipped (WhatIf).' }
    }

    $snap = Save-PrintConfigSnapshot -Label $(if ($Reset) { 'namedpipe-listener-reset' } else { 'namedpipe-listener-on' }) -RegistryName @('RpcProtocols')
    try {
        if ($Reset) {
            if (Test-Path -LiteralPath $path) { Remove-ItemProperty -Path $path -Name 'RpcProtocols' -Force -ErrorAction SilentlyContinue }
        } else {
            if (-not (Test-Path -LiteralPath $path)) { New-Item -Path $path -Force | Out-Null }
            Set-ItemProperty -Path $path -Name 'RpcProtocols' -Value $Protocols -Type DWord -ErrorAction Stop
        }
    } catch {
        return [pscustomobject]@{ Success = $false; Value = $null; IsSet = $null; RestartSucceeded = $null; SnapshotPath = $snap.Path
            Message = "Could not write the RPC listener setting: $($_.Exception.Message)" }
    }

    $restartOk = $null
    if (-not $SkipSpoolerRestart) { $restartOk = [bool](Restart-PrintSpooler).Success }

    $now = $null; $isSet = $false
    try { $now = (Get-ItemProperty -Path $path -Name 'RpcProtocols' -ErrorAction Stop).RpcProtocols; $isSet = $true } catch {}
    $valueOk = if ($Reset) { -not $isSet } else { "$now" -eq "$Protocols" }

    $restartTxt = if ($null -eq $restartOk) { ' The spooler was NOT restarted, so this is not in effect yet.' }
                  elseif ($restartOk) { ' Spooler restarted.' }
                  else { ' The spooler did NOT restart - this is not in effect.' }
    $msg = if (-not $valueOk) {
               $(if ($Reset) { "RpcProtocols could not be removed - it still reads '$now'. Something is re-applying it, most likely Group Policy." }
                 else { "RpcProtocols did not stick - it reads '$now' rather than $Protocols. Something is re-applying it, most likely Group Policy." })
           } elseif ($Reset) {
               "RpcProtocols removed - this spooler listens over TCP only again (the Windows default).$restartTxt"
           } else {
               "RpcProtocols = $Protocols written.$restartTxt Microsoft does not recommend RPC over named pipes. Prefer pinning the spooler port and allowing it through the firewall; this proves nothing until a client tests a connection."
           }
    [pscustomobject]@{
        Success          = ($valueOk -and $restartOk -ne $false -and -not ($null -eq $restartOk -and -not $SkipSpoolerRestart))
        Value            = $now
        IsSet            = $isSet
        RestartSucceeded = $restartOk
        SnapshotPath     = $snap.Path
        Message          = $msg
    }
}

function Test-RemotePrintEnumeration {
    <#
    .SYNOPSIS
        Can this console list a host's printers through the print provider? Bounded.
    .DESCRIPTION
        Uses Get-Printer -ComputerName. The PrintManagement module declares
        ComputerName as a QUERY OPTION passed to the printer provider
        (MSFT_Printer_v1.0.cdxml, under <QueryOptions>), not as a remote CIM
        session. An earlier version of this toolkit said this call travels over
        WinRM; that was wrong and has been withdrawn.

        What success shows: this console, as its own identity, could enumerate that
        host's queues. What it does NOT show: that a particular queue opens, that it
        opens for the target user, that a driver installs, or that a job prints. It
        is context, never proof of a repair - an earlier version used a before/after
        pair of these as "proof" across a spooler restart, which proved nothing.

        Bounded, because a blocked call can sit for a long time.
    .OUTPUTS
        PSCustomObject { Ok; Count; Names; Error; TimedOut }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [int]$TimeoutSec = 45
    )

    $target = ([string]$ComputerName).Trim().Trim('\')
    if (-not $target) {
        return [pscustomobject]@{ Ok = $false; Count = 0; Names = @(); Error = 'No host given.'; TimedOut = $false }
    }

    $raw = Invoke-WithTimeout -TimeoutSec $TimeoutSec -Parameters @{ Srv = $target } -ScriptBlock {
        param($Srv)
        try {
            $p = @(Get-Printer -ComputerName $Srv -ErrorAction Stop)
            [pscustomobject]@{ Ok = $true; Count = $p.Count; Names = @($p | ForEach-Object { [string]$_.Name }); Error = '' }
        } catch {
            [pscustomobject]@{ Ok = $false; Count = 0; Names = @(); Error = $_.Exception.Message }
        }
    }

    $r = $null
    if ($raw) { $r = @($raw)[0] }
    if (-not $r) {
        return [pscustomobject]@{
            Ok = $false; Count = 0; Names = @()
            Error = "Timed out after $TimeoutSec seconds - the call never returned."
            TimedOut = $true
        }
    }
    [pscustomobject]@{
        Ok       = [bool]$r.Ok
        Count    = [int]$r.Count
        Names    = @($r.Names)
        Error    = [string]$r.Error
        TimedOut = $false
    }
}


function Enable-PrinterShareForAll {
    <#
    .SYNOPSIS
        Make a printer available to every user on the network, in one action.
    .DESCRIPTION
        The "share this printer to the office" button. Does the four things a print
        host actually needs, all of which are normal configuration rather than
        security downgrades:

          1. share the queue (with a share name if it has none)
          2. grant Print to Everyone on the printer's permissions, appending to
             the existing DACL rather than replacing it
          3. enable the File and Printer Sharing inbound firewall rules
          4. move any Public network adapter to Private, because Windows blocks
             inbound sharing on Public

        WHAT IT DELIBERATELY DOES NOT DO, and why it cannot be a complete answer
        in a workgroup: "Everyone" is evaluated AFTER authentication. A client
        still has to prove who it is to this machine first. In a domain that is
        every domain user, so this is genuinely enough. On a workgroup host it is
        not - the connecting user needs an account that exists here, or guest SMB,
        and Windows 11 24H2+ refuses insecure guest by default for good reason.
        This function will not enable guest, will not touch SMB signing, and will
        not weaken Point and Print. It reports that gap instead of hiding it.
    .OUTPUTS
        PSCustomObject { Success; Printer; ShareName; Steps[]; Warnings[]; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$ShareName,
        [ValidateSet('Everyone', 'AuthenticatedUsers')][string]$Trustee = 'Everyone',
        [switch]$SkipFirewall,
        [switch]$SkipNetworkProfile
    )

    $steps    = [System.Collections.Generic.List[object]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()
    function Step([string]$Tag, [string]$Text) { $steps.Add([pscustomobject]@{ Tag=$Tag; Text=$Text }) }

    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $isAdmin = $false
    try { $isAdmin = ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } catch {}
    if (-not $isAdmin) {
        return [pscustomobject]@{ Success=$false; Printer=$Name; ShareName=''; Steps=@(); Warnings=@()
            Message='Sharing a printer needs elevation. Run as Administrator, or as SYSTEM from Backstage.' }
    }

    $q = Get-QueueExact $Name
    if (-not $q) {
        return [pscustomobject]@{ Success=$false; Printer=$Name; ShareName=''; Steps=@(); Warnings=@()
            Message="No queue named '$Name' on this machine." }
    }
    if (-not $ShareName) {
        $ShareName = if ($q.ShareName) { [string]$q.ShareName } else { ($Name -replace '[,\\\\/]', ' ').Trim() }
    }
    if (-not $PSCmdlet.ShouldProcess($Name, "Share to $Trustee")) {
        return [pscustomobject]@{ Success=$false; Printer=$Name; ShareName=$ShareName; Steps=@(); Warnings=@(); Message='Skipped (WhatIf).' }
    }

    # 1. share it
    if ($q.Shared -and $q.ShareName) {
        Step 'INFO' "Already shared as '$($q.ShareName)'."
        $ShareName = [string]$q.ShareName
    } else {
        try {
            Set-Printer -Name $Name -Shared $true -ShareName $ShareName -ErrorAction Stop
            Step 'OK' "Shared as '$ShareName'."
        } catch {
            Step 'FAIL' "Could not share the queue: $($_.Exception.Message)"
        }
    }

    # 2. printer permissions
    $acl = Grant-PrinterPrintAccess -Name $Name -Trustee $Trustee
    Step $(if ($acl.Success) { 'OK' } else { 'WARN' }) $acl.Message

    # 3. firewall
    if (-not $SkipFirewall) {
        $fw = Enable-PrintSharingFirewall
        Step $(if ($fw.Success) { 'OK' } else { 'WARN' }) $fw.Message
    }

    # 4. network profile
    if (-not $SkipNetworkProfile) {
        $np = Set-PrintSharingNetworkPrivate
        Step $(if ($np.Success) { 'OK' } else { 'INFO' }) $np.Message
    }

    # ---- the honest part: what this cannot fix -----------------------------
    $auth = Get-SharingAuthState
    if (-not $auth.PartOfDomain) {
        $warnings.Add("This host is in a WORKGROUP. 'Everyone' only applies once a client has authenticated, so a user whose account does not exist HERE still cannot connect - granting Everyone does not change that. Create a matching local account (same username and password) on this PC for each person who needs to print, or move the printer to its own network connection. Enabling guest SMB would also do it and is exactly what this toolkit will not do for you.")
    }
    $usb = $null
    if ($q.PortName -match '^USB') {
        $usb = Test-UsbPrinterAttached -PortName ([string]$q.PortName) -Name $Name -DriverName ([string]$q.DriverName)
        if (-not $usb.Attached) {
            $warnings.Add("The printer itself is not attached ($($usb.Message)) - the share will be reachable but nothing will print until the device is connected and switched on, and this PC has to stay on for anyone else to print.")
        }
    }
    if (Test-IsIppClassDriver -DriverName ([string]$q.DriverName)) {
        $warnings.Add("This queue uses the generic '$($q.DriverName)'. Every client that connects inherits it, along with its missing trays and duplex. Consider installing the vendor driver here first.")
    }

    $failed = @($steps.ToArray() | Where-Object { $_.Tag -eq 'FAIL' })
    [pscustomobject]@{
        Success   = ($failed.Count -eq 0)
        Printer   = $Name
        ShareName = $ShareName
        Steps     = $steps.ToArray()
        Warnings  = $warnings.ToArray()
        Message   = $(if ($failed.Count) { "Shared with $($failed.Count) problem(s) - see the steps." }
                      else { "\\$env:COMPUTERNAME\$ShareName is shared and grants Print to $Trustee." })
    }
}


# --------------------------------------------------------------------------
# Junk queue removal - Fax, XPS, OneNote for Windows 10
# --------------------------------------------------------------------------

function Get-PrintFeatureState {
    <#
    .SYNOPSIS
        Enabled/Disabled state of the optional features that RECREATE a junk queue.
    .DESCRIPTION
        Deleting the XPS or Fax queue is temporary: both are Windows optional
        features, and Windows puts the queue back on servicing, a repair, or a
        feature update. Reporting "removed" without saying so would be a lie with
        a delay on it.

        Get-WindowsOptionalFeature REQUIRES ELEVATION even to read - verified, it
        throws "The requested operation requires elevation" as a standard user -
        so this returns 'Unknown' rather than pretending, and the caller must not
        present Unknown as Disabled.
    .OUTPUTS
        PSCustomObject { Elevated; States (hashtable name -> Enabled|Disabled|Absent|Unknown); Message }
    #>
    [CmdletBinding()]
    param([string[]]$FeatureName)

    if (-not $FeatureName -or -not $FeatureName.Count) {
        $FeatureName = @($script:JunkPrinterCatalogue | Where-Object { $_.Feature } | ForEach-Object { $_.Feature })
    }

    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $isAdmin = $false
    try { $isAdmin = ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } catch {}

    $states = @{}
    foreach ($f in $FeatureName) { $states[$f] = 'Unknown' }

    if (-not $isAdmin) {
        return [pscustomobject]@{ Elevated=$false; States=$states
            Message='Reading Windows optional features needs elevation, so it is not known whether the XPS/Fax features are still enabled.' }
    }

    foreach ($f in $FeatureName) {
        try {
            $wf = Get-WindowsOptionalFeature -Online -FeatureName $f -ErrorAction Stop
            $states[$f] = if ($wf) { [string]$wf.State } else { 'Absent' }
        } catch {
            # A feature that does not exist on this SKU is 'Absent', not an error.
            $states[$f] = 'Absent'
        }
    }
    [pscustomobject]@{ Elevated=$true; States=$states; Message='Read the optional feature states.' }
}

function Get-JunkPrinterCandidate {
    <#
    .SYNOPSIS
        Local queues that match the narrow junk catalogue, annotated with the
        facts that decide whether removing one is safe.
    .DESCRIPTION
        Matches on EXACT queue name or EXACT driver name from
        $script:JunkPrinterCatalogue, then applies the never-remove gate. It does
        NOT use the virtual-printer regex, which is a display filter and would
        sweep up CutePDF, Snagit and Microsoft Print to PDF.

        Each candidate carries:
          IsDefaultFor  - users who currently have it as their DEFAULT printer.
                          Deleting one of those is its own support call: Windows
                          then reassigns arbitrarily, or LegacyDefaultPrinterMode
                          picks last-used (the MAP004 finding).
          Feature       - the optional feature that will put it back, if any
          FeatureState  - Enabled / Disabled / Absent / Unknown (needs elevation)
          Shared        - a shared junk queue may have clients pointed at it
    .OUTPUTS
        PSCustomObject[] - empty array when the machine is clean.
    #>
    [CmdletBinding()]
    param([switch]$SkipFeatureCheck)

    $printers = @()
    try { $printers = @(Get-Printer -ErrorAction SilentlyContinue) } catch {}

    # Who has what as their default, so a candidate can be flagged before it goes.
    $defaults = @{}
    try {
        foreach ($m in @(Get-AllUserPrinterMaps)) {
            $dp = [string]$m.DefaultPrinter
            if (-not $dp) { continue }
            if (-not $defaults.ContainsKey($dp)) { $defaults[$dp] = [System.Collections.Generic.List[string]]::new() }
            $defaults[$dp].Add([string]$m.User)
        }
    } catch {}

    $featStates = $null
    if (-not $SkipFeatureCheck) { $featStates = Get-PrintFeatureState }

    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($p in $printers) {
        $name = [string]$p.Name
        $drv  = [string]$p.DriverName

        $hit = $null
        foreach ($c in $script:JunkPrinterCatalogue) {
            $byName   = @($c.Names)   -contains $name
            $byDriver = @($c.Drivers) -contains $drv
            if ($byName -or $byDriver) { $hit = $c; break }
        }
        if (-not $hit) { continue }

        # Second gate - see $script:JunkPrinterNeverRemove.
        if ($script:JunkPrinterNeverRemove -contains $name) { continue }
        if ($script:JunkPrinterNeverRemoveDrivers -contains $drv) { continue }

        $fState = 'Unknown'
        if ($hit.Feature -and $featStates) { $fState = [string]$featStates.States[$hit.Feature] }
        elseif (-not $hit.Feature) { $fState = 'None' }

        $defFor = @()
        if ($defaults.ContainsKey($name)) { $defFor = $defaults[$name].ToArray() }

        $out.Add([pscustomobject]@{
            Name         = $name
            DriverName   = $drv
            PortName     = [string]$p.PortName
            Id           = [string]$hit.Id
            Label        = [string]$hit.Label
            Why          = [string]$hit.Why
            Feature      = $(if ($hit.Feature) { [string]$hit.Feature } else { '' })
            FeatureState = $fState
            ComesBack    = ($fState -eq 'Enabled')
            Shared       = [bool]$p.Shared
            IsDefaultFor = $defFor
        })
    }
    return $out.ToArray()
}

function Remove-JunkPrinter {
    <#
    .SYNOPSIS
        Remove named junk queues, reporting each one separately.
    .DESCRIPTION
        Per-queue results on purpose: a sweep that removes two of three and
        reports a single "done" hides the one that failed, which is the queue the
        technician actually needed to know about.

        Only removes names that are still in the catalogue at call time, so a
        stale dialog cannot be used to delete an arbitrary printer.
    .OUTPUTS
        PSCustomObject { Success; Removed; Failed; Skipped; Results; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string[]]$Name,
        [scriptblock]$OnPoll
    )

    $allowed = @{}
    foreach ($c in @(Get-JunkPrinterCandidate -SkipFeatureCheck)) { $allowed[$c.Name] = $c }

    $results = [System.Collections.Generic.List[object]]::new()
    $removed = 0; $failed = 0; $skipped = 0

    foreach ($n in @($Name)) {
        if (-not $allowed.ContainsKey($n)) {
            $skipped++
            $results.Add([pscustomobject]@{ Name=$n; Success=$false
                Message='Not a junk queue on this PC any more - skipped rather than removing something unexpected.' })
            continue
        }
        if (-not $PSCmdlet.ShouldProcess($n, 'Remove junk printer queue')) {
            $skipped++
            $results.Add([pscustomobject]@{ Name=$n; Success=$false; Message='Skipped (WhatIf).' })
            continue
        }

        $r = Remove-PrinterSafe -Name $n -OnPoll $OnPoll
        if ($r.Success) { $removed++ } else { $failed++ }
        $results.Add([pscustomobject]@{ Name=$n; Success=[bool]$r.Success; Message=[string]$r.Message })
        if ($OnPoll) { & $OnPoll }
    }

    $parts = [System.Collections.Generic.List[string]]::new()
    if ($removed) { $parts.Add("removed $removed") }
    if ($failed)  { $parts.Add("$failed failed") }
    if ($skipped) { $parts.Add("$skipped skipped") }
    $summary = if ($parts.Count) { ($parts -join ', ') } else { 'nothing to do' }

    [pscustomobject]@{
        Success = ($failed -eq 0 -and $removed -gt 0)
        Removed = $removed
        Failed  = $failed
        Skipped = $skipped
        Results = $results.ToArray()
        Message = "Junk queue sweep: $summary."
    }
}

function Disable-PrintOptionalFeature {
    <#
    .SYNOPSIS
        Turn off the Windows optional feature that keeps recreating a junk queue.
    .DESCRIPTION
        The durable half of removing XPS or Fax. Without this the queue returns.

        -NoRestart is always passed: this console must never reboot a machine a
        technician is working on. RestartNeeded is reported instead, so the
        pending restart is stated rather than hidden.

        DISM work is slow (tens of seconds is normal), so the caller should show
        a busy state.
    .OUTPUTS
        PSCustomObject { Success; Feature; State; RestartNeeded; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$FeatureName)

    # Only features this toolkit actually knows about - never an arbitrary one.
    $known = @($script:JunkPrinterCatalogue | Where-Object { $_.Feature } | ForEach-Object { $_.Feature })
    if ($known -notcontains $FeatureName) {
        return [pscustomobject]@{ Success=$false; Feature=$FeatureName; State=$null; RestartNeeded=$false
            Message="'$FeatureName' is not one of the print features this console manages." }
    }

    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $isAdmin = $false
    try { $isAdmin = ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } catch {}
    if (-not $isAdmin) {
        return [pscustomobject]@{ Success=$false; Feature=$FeatureName; State=$null; RestartNeeded=$false
            Message='Disabling a Windows optional feature needs elevation. Run as Administrator, or as SYSTEM from Backstage.' }
    }
    if (-not $PSCmdlet.ShouldProcess($FeatureName, 'Disable Windows optional feature')) {
        return [pscustomobject]@{ Success=$false; Feature=$FeatureName; State=$null; RestartNeeded=$false; Message='Skipped (WhatIf).' }
    }

    try {
        $r = Disable-WindowsOptionalFeature -Online -FeatureName $FeatureName -NoRestart -ErrorAction Stop
        $restart = $false
        try { $restart = [bool]$r.RestartNeeded } catch {}

        # Re-read rather than trusting the call.
        $state = 'Unknown'
        try { $state = [string](Get-WindowsOptionalFeature -Online -FeatureName $FeatureName -ErrorAction Stop).State } catch {}

        $ok = ($state -eq 'Disabled' -or $state -eq 'DisablePending')
        $msg = if ($ok -and $restart) {
                   "$FeatureName is disabled, pending a restart. The queue will not come back once the machine has restarted."
               } elseif ($ok) {
                   "$FeatureName is disabled. The queue will not come back."
               } else {
                   "Asked Windows to disable $FeatureName but it now reads '$state' - the queue may still return."
               }
        [pscustomobject]@{ Success=$ok; Feature=$FeatureName; State=$state; RestartNeeded=$restart; Message=$msg }
    } catch {
        [pscustomobject]@{ Success=$false; Feature=$FeatureName; State=$null; RestartNeeded=$false
            Message="Could not disable ${FeatureName}: $($_.Exception.Message)" }
    }
}

# ===========================================================================
# region  v2.0 - measured print RPC: listener, firewall, pinned port, probes
#
# Microsoft's guidance for Windows 11 22H2+ print RPC (Learn, "Windows 11 RPC
# connection updates for print", 2026-02-12): RPC over TCP is the default and the
# recommended configuration. The fix for a blocked connection is the firewall -
# allow 135 and the spooler's RPC traffic - optionally with the spooler pinned to
# one port (RpcTcpPort). RPC over named pipes "isn't recommended", and may also
# need RpcAuthnLevelPrivacyEnabled=0 and insecure guest logons, which this
# toolkit refuses to set.
#
# The principle of this region: MEASURE, never infer. A dynamic RPC port cannot
# be tested from a client because the client does not know it; pinning the
# spooler to a fixed port turns that into one TCP connect. Every probe returns an
# explicit state - Ok, Failed, Timeout or Skipped - so missing evidence is never
# mistaken for a measured failure.
# ===========================================================================

$script:PrintRpcRuleNames = @{
    Smb     = @('FPS-SMB-In-TCP', 'FPS-SMB-In-TCP-NoScope', 'FPS-SMB-In-TCP-V2')
    Epm     = @('FPS-RPCSS-In-TCP', 'FPS-RPCSS-In-TCP-NoScope', 'FPS-RPCSS-In-TCP-V2')
    Spooler = @('FPS-SpoolSvc-In-TCP', 'FPS-SpoolSvc-In-TCP-NoScope', 'FPS-SpoolSvc-In-TCP-V2')
}
$script:PrintRpcToolkitRulePrefix = 'PrinterToolkit-SpoolerRpc-TCP-'
$script:PrintRpcPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC'
$script:SnapshotDir = Join-Path $env:ProgramData 'PrinterToolkit\snapshots'
# Only these values may ever be recorded or restored. A tampered snapshot file
# must not become a way to write arbitrary registry values as SYSTEM.
$script:SnapshotRegistryAllow = @('RpcTcpPort', 'RpcProtocols', 'RpcUseNamedPipeProtocol')

function Test-PrintToolkitElevated {
    <# .SYNOPSIS True when running elevated or as SYSTEM. #>
    [CmdletBinding()] param()
    $isAdmin = $false
    try {
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
            [Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {}
    return $isAdmin
}

function ConvertTo-PrintHostKey {
    <#
    .SYNOPSIS
        One canonical key for a host, however it was typed.
    .DESCRIPTION
        '\\PRINTHOST-01', 'printhost-01', 'PRINTHOST-01\' and '\\PRINTHOST-01\Canon G6065'
        all become 'printhost-01'. Credential caches and host comparisons use this
        rather than ad-hoc trimming: a Trim('') that stripped nothing is how cached
        credentials silently stopped reaching Enumerate.
    #>
    [CmdletBinding()]
    param([string]$Value)
    $v = ([string]$Value).Trim().Trim('"').TrimStart('\')
    $cut = $v.IndexOf('\')
    if ($cut -ge 0) { $v = $v.Substring(0, $cut) }
    return $v.Trim().ToLowerInvariant()
}

function New-PrintProbeResult {
    <#
    .SYNOPSIS
        One measured observation, with an explicit state.
    .DESCRIPTION
        State is the whole point. Failed means the probe ran and got a negative
        answer. Timeout means it never answered. Skipped means it did not run.
        Reading Skipped or Timeout as Failed is how a missing measurement became a
        confident wrong diagnosis.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Probe,
        [Parameter(Mandatory)][ValidateSet('Ok', 'Failed', 'Timeout', 'Skipped')][string]$State,
        [string]$Code = '',
        [string]$Detail = '',
        [string]$Identity = '',
        [string]$Target = ''
    )
    [pscustomobject]@{ Probe = $Probe; State = $State; Code = $Code; Detail = $Detail; Identity = $Identity; Target = $Target }
}

function Get-PrintSpoolerListener {
    <#
    .SYNOPSIS
        The TCP ports the Print Spooler is ACTUALLY listening on right now.
    .DESCRIPTION
        Activation evidence, not configuration. A registry value can say
        RpcTcpPort = 50500 while the spooler still listens elsewhere, because it has
        not restarted or the port was taken first. This reads the live socket table
        owned by spoolsv.exe. No elevation needed.
    .OUTPUTS
        PSCustomObject { Running; ProcessId; Ports[]; State; Message }
    #>
    [CmdletBinding()] param()

    $proc = @(Get-Process -Name 'spoolsv' -ErrorAction SilentlyContinue)
    if (-not $proc.Count) {
        return [pscustomobject]@{ Running = $false; ProcessId = $null; Ports = @(); State = 'Skipped'
            Message = 'The Print Spooler process is not running, so it is listening on nothing.' }
    }
    $spoolPid = [int]$proc[0].Id
    $ports = @()
    try {
        $conns = @(Get-NetTCPConnection -OwningProcess $spoolPid -State Listen -ErrorAction SilentlyContinue)
        $ports = @($conns | Where-Object { $_ } | ForEach-Object { [int]$_.LocalPort } | Sort-Object -Unique)
    } catch {
        return [pscustomobject]@{ Running = $true; ProcessId = $spoolPid; Ports = @(); State = 'Skipped'
            Message = "Could not read the spooler's listening ports: $($_.Exception.Message)" }
    }
    [pscustomobject]@{
        Running   = $true
        ProcessId = $spoolPid
        Ports     = @($ports)
        State     = 'Ok'
        Message   = $(if ($ports.Count) { 'The Print Spooler is listening on TCP ' + ($ports -join ', ') + '.' }
                      else { 'The Print Spooler is running but listening on no TCP port, so remote clients cannot reach it over RPC/TCP.' })
    }
}

function Get-ActiveFirewallProfile {
    <# .SYNOPSIS The active network categories, as firewall profile names: Domain, Private, Public. #>
    [CmdletBinding()] param()
    $out = [System.Collections.Generic.List[string]]::new()
    foreach ($netProf in @(Get-SharingNetworkProfile)) {
        if (-not $netProf) { continue }
        $c = [string]$netProf.NetworkCategory
        $name = if ($c -match 'Domain') { 'Domain' } elseif ($c -match 'Private') { 'Private' } elseif ($c -match 'Public') { 'Public' } else { '' }
        if ($name -and -not $out.Contains($name)) { $out.Add($name) }
    }
    return $out.ToArray()
}

function Test-PrintToolkitRuleNameAllowed {
    <# .SYNOPSIS Is this a firewall rule name the toolkit may record, change or restore? #>
    [CmdletBinding()] param([string]$Name)
    $n = [string]$Name
    $known = @($script:PrintRpcRuleNames.Smb) + @($script:PrintRpcRuleNames.Epm) + @($script:PrintRpcRuleNames.Spooler)
    if ($known -contains $n) { return $true }
    return ($n -match ('^' + [regex]::Escape($script:PrintRpcToolkitRulePrefix) + '\d{1,5}$'))
}

function Get-PrintRpcFirewallState {
    <#
    .SYNOPSIS
        Would the firewall on THIS PC let a client reach its spooler? Read by rule Name.
    .DESCRIPTION
        Reads the built-in rules by their locale-independent Name rather than the
        English display group "File and Printer Sharing", which does not exist on a
        German or French Windows. Three paths matter, and they fail separately:

          Smb      445   FPS-SMB-In-TCP*        file access and net view
          Epm      135   FPS-RPCSS-In-TCP*      RPC endpoint mapper
          Spooler  RPC   FPS-SpoolSvc-In-TCP*   the printer connection itself

        "SMB allowed, spooler not" is precisely "net view works but adding the
        printer fails", and it was invisible while the check only asked whether ANY
        sharing rule was enabled.

        Judged against the ACTIVE network profile. The Private/Public variants ship
        scoped to LocalSubnet, so a client on another subnet, VLAN or VPN is refused
        even with the rule enabled; that is reported as LocalSubnetOnly.

        Whether the built-in spooler rule (LocalPort=RPC) also admits a port pinned
        with RpcTcpPort is not documented, so it is reported as unknown ($null)
        unless a toolkit rule for that exact port exists. Measure it from the client.
    .OUTPUTS
        PSCustomObject { State; ActiveProfiles[]; Rules[]; Smb; Epm; Spooler; Message }
    #>
    [CmdletBinding()]
    param([int]$TimeoutSec = 25, [int]$PinnedPort = 0)

    $active = @(Get-ActiveFirewallProfile)
    $names = [string[]](@($script:PrintRpcRuleNames.Smb) + @($script:PrintRpcRuleNames.Epm) + @($script:PrintRpcRuleNames.Spooler))

    $raw = Invoke-WithTimeout -TimeoutSec $TimeoutSec -Parameters @{ Names = $names; Prefix = $script:PrintRpcToolkitRulePrefix } -ScriptBlock {
        param([string[]]$Names, [string]$Prefix)
        $rules = [System.Collections.Generic.List[object]]::new()
        foreach ($n in $Names) {
            $one = Get-NetFirewallRule -Name $n -ErrorAction SilentlyContinue
            if ($one) { $rules.Add($one) }
        }
        foreach ($one in @(Get-NetFirewallRule -Name ($Prefix + '*') -ErrorAction SilentlyContinue)) {
            if ($one) { $rules.Add($one) }
        }
        foreach ($rule in $rules) {
            $pf = $null; $af = $null
            try { $pf = $rule | Get-NetFirewallPortFilter -ErrorAction Stop } catch {}
            try { $af = $rule | Get-NetFirewallAddressFilter -ErrorAction Stop } catch {}
            [pscustomobject]@{
                Name          = [string]$rule.Name
                DisplayName   = [string]$rule.DisplayName
                Enabled       = ([string]$rule.Enabled -eq 'True')
                Action        = [string]$rule.Action
                Direction     = [string]$rule.Direction
                Profile       = [string]$rule.Profile
                LocalPort     = (@($pf.LocalPort) -join ',')
                RemoteAddress = (@($af.RemoteAddress) -join ',')
                PolicyStore   = [string]$rule.PolicyStoreSourceType
            }
        }
    }

    if ($null -eq $raw) {
        return [pscustomobject]@{ State = 'Timeout'; ActiveProfiles = $active; Rules = @(); Smb = $null; Epm = $null; Spooler = $null
            Message = "Reading the firewall rules did not finish within $TimeoutSec seconds." }
    }
    $rules = @($raw | Where-Object { $_ })
    if (-not $rules.Count) {
        return [pscustomobject]@{ State = 'Skipped'; ActiveProfiles = $active; Rules = @(); Smb = $null; Epm = $null; Spooler = $null
            Message = 'None of the built-in File and Printer Sharing rules were found by name. This Windows edition may not ship them, or a third-party firewall manages this PC - check it directly.' }
    }

    $covers = {
        param([string]$RuleProfile)
        if (-not $active.Count) { return $false }
        if ($RuleProfile -eq 'Any') { return $true }
        foreach ($a in $active) { if ($RuleProfile -match $a) { return $true } }
        return $false
    }
    $role = {
        param([string[]]$RoleNames, [string]$Label)
        $mine = @($rules | Where-Object { $RoleNames -contains $_.Name })
        $allowing = @($mine | Where-Object { $_.Enabled -and $_.Direction -eq 'Inbound' -and $_.Action -eq 'Allow' -and (& $covers $_.Profile) })
        $localOnly = ($allowing.Count -gt 0) -and (@($allowing | Where-Object { $_.RemoteAddress -notmatch '^LocalSubnet' }).Count -eq 0)
        [pscustomobject]@{
            Label           = $Label
            Found           = $mine.Count
            Allowed         = ($allowing.Count -gt 0)
            AllowingRules   = @($allowing | ForEach-Object { $_.Name })
            LocalSubnetOnly = $localOnly
        }
    }

    $toolkitNames = @($rules | Where-Object { $_.Name -like ($script:PrintRpcToolkitRulePrefix + '*') } | ForEach-Object { $_.Name })
    $smb   = & $role ([string[]]@($script:PrintRpcRuleNames.Smb)) 'SMB 445'
    $epm   = & $role ([string[]]@($script:PrintRpcRuleNames.Epm)) 'RPC endpoint mapper 135'
    $spool = & $role ([string[]](@($script:PrintRpcRuleNames.Spooler) + $toolkitNames)) 'Print Spooler RPC'

    $pinnedCovered = $null
    if ($PinnedPort -gt 0) {
        $exact = @($rules | Where-Object { $_.Name -eq ($script:PrintRpcToolkitRulePrefix + $PinnedPort) -and $_.Enabled -and (& $covers $_.Profile) })
        if ($exact.Count) { $pinnedCovered = $true }
    }
    $spool | Add-Member -NotePropertyName PinnedPort -NotePropertyValue $PinnedPort
    $spool | Add-Member -NotePropertyName PinnedPortCovered -NotePropertyValue $pinnedCovered

    $blocked = @(@($smb, $epm, $spool) | Where-Object { -not $_.Allowed } | ForEach-Object { $_.Label })
    [pscustomobject]@{
        State          = 'Ok'
        ActiveProfiles = $active
        Rules          = $rules
        Smb            = $smb
        Epm            = $epm
        Spooler        = $spool
        Message        = $(if (-not $active.Count) { 'No active network profile could be read, so the rules could not be judged against one.' }
                           elseif ($blocked.Count) { 'Not allowed inbound for the active profile (' + ($active -join ', ') + '): ' + ($blocked -join '; ') + '.' }
                           else { 'SMB, the endpoint mapper and the spooler are all allowed inbound for the active profile (' + ($active -join ', ') + ').' })
    }
}

function Save-PrintConfigSnapshot {
    <#
    .SYNOPSIS
        Record the exact prior state of the settings this toolkit is about to change.
    .DESCRIPTION
        Rollback that writes "the default" is not rollback. A value that was ABSENT
        must come back absent, a rule that was DISABLED must come back disabled, and
        a rule that did not exist must be deleted. This records presence as well as
        value so Restore-PrintConfigSnapshot can do exactly that.

        Holds no secrets. Written to ProgramData so it outlives the console. Only
        allow-listed registry values and rule names are recorded or restored.
    .OUTPUTS
        PSCustomObject { Path; Created; Computer; Label; Registry[]; Rules[] }
    #>
    [CmdletBinding()]
    param(
        [string]$Label = 'change',
        [string[]]$RegistryName = @(),
        [string[]]$RuleName = @()
    )
    $reg = [System.Collections.Generic.List[object]]::new()
    foreach ($n in @($RegistryName)) {
        if ($script:SnapshotRegistryAllow -notcontains $n) { continue }
        $present = $false; $val = $null
        try { $val = (Get-ItemProperty -Path $script:PrintRpcPolicyPath -Name $n -ErrorAction Stop).$n; $present = $true } catch {}
        $reg.Add([pscustomobject]@{ Name = $n; Present = $present; Value = $val })
    }
    $rul = [System.Collections.Generic.List[object]]::new()
    foreach ($n in @($RuleName)) {
        if (-not (Test-PrintToolkitRuleNameAllowed -Name $n)) { continue }
        $one = $null
        try { $one = Get-NetFirewallRule -Name $n -ErrorAction Stop } catch {}
        $rul.Add([pscustomobject]@{ Name = $n; Existed = [bool]$one; Enabled = $(if ($one) { ([string]$one.Enabled -eq 'True') } else { $false }) })
    }
    $snap = [pscustomobject]@{
        Path     = ''
        Created  = (Get-Date).ToString('o')
        Computer = $env:COMPUTERNAME
        Label    = $Label
        Registry = $reg.ToArray()
        Rules    = $rul.ToArray()
    }
    try {
        if (-not (Test-Path -LiteralPath $script:SnapshotDir)) { New-Item -ItemType Directory -Path $script:SnapshotDir -Force | Out-Null }
        $safeLabel = ($Label -replace '[^A-Za-z0-9-]', '')
        $file = Join-Path $script:SnapshotDir ('rpc-{0}-{1}.json' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $safeLabel)
        $snap.Path = $file
        ($snap | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $file -Encoding UTF8
    } catch {
        $snap.Path = ''
    }
    return $snap
}

function Get-PrintConfigSnapshot {
    <# .SYNOPSIS The saved snapshots on this PC, newest first. #>
    [CmdletBinding()] param()
    if (-not (Test-Path -LiteralPath $script:SnapshotDir)) { return @() }
    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($f in @(Get-ChildItem -LiteralPath $script:SnapshotDir -Filter 'rpc-*.json' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)) {
        $o = $null
        try { $txt = Get-Content -LiteralPath $f.FullName -Raw -ErrorAction Stop; $o = $txt | ConvertFrom-Json } catch { continue }
        if (-not $o) { continue }
        $out.Add([pscustomobject]@{ Path = $f.FullName; Created = [string]$o.Created; Label = [string]$o.Label; Computer = [string]$o.Computer })
    }
    return $out.ToArray()
}

function Restore-PrintConfigSnapshot {
    <#
    .SYNOPSIS
        Put allow-listed settings back exactly as a snapshot recorded them.
    .DESCRIPTION
        Registry values that were absent are removed; values that were present get
        their recorded value. Rules that did not exist are deleted; rules that did
        get their recorded enabled state. The spooler is restarted only if a
        registry value actually changed, and the restart result is reported.
    .OUTPUTS
        PSCustomObject { Success; Path; Items[]; RestartSucceeded; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([string]$Path)

    if (-not (Test-PrintToolkitElevated)) {
        return [pscustomobject]@{ Success = $false; Path = $Path; Items = @(); RestartSucceeded = $null
            Message = 'Restoring settings needs elevation. Run as Administrator, or as SYSTEM from Backstage.' }
    }
    if (-not $Path) {
        $all = @(Get-PrintConfigSnapshot)
        if (-not $all.Count) {
            return [pscustomobject]@{ Success = $false; Path = ''; Items = @(); RestartSucceeded = $null; Message = 'There is no saved snapshot on this PC to restore.' }
        }
        $Path = [string]$all[0].Path
    }
    $snap = $null
    try { $txt = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop; $snap = $txt | ConvertFrom-Json } catch {
        return [pscustomobject]@{ Success = $false; Path = $Path; Items = @(); RestartSucceeded = $null; Message = "Could not read the snapshot: $($_.Exception.Message)" }
    }
    if ([string]$snap.Computer -and ([string]$snap.Computer -ne $env:COMPUTERNAME)) {
        return [pscustomobject]@{ Success = $false; Path = $Path; Items = @(); RestartSucceeded = $null
            Message = "That snapshot was taken on $($snap.Computer), not this PC. Refusing to apply it here." }
    }
    if (-not $PSCmdlet.ShouldProcess($Path, 'Restore print RPC settings')) {
        return [pscustomobject]@{ Success = $false; Path = $Path; Items = @(); RestartSucceeded = $null; Message = 'Skipped (WhatIf).' }
    }

    $items = [System.Collections.Generic.List[object]]::new()
    $regChanged = $false
    foreach ($e in @($snap.Registry)) {
        if (-not $e) { continue }
        $n = [string]$e.Name
        if ($script:SnapshotRegistryAllow -notcontains $n) { continue }
        try {
            if ([bool]$e.Present) {
                if (-not (Test-Path -LiteralPath $script:PrintRpcPolicyPath)) { New-Item -Path $script:PrintRpcPolicyPath -Force | Out-Null }
                Set-ItemProperty -Path $script:PrintRpcPolicyPath -Name $n -Value ([int]$e.Value) -Type DWord -ErrorAction Stop
                $items.Add([pscustomobject]@{ Item = $n; Ok = $true; Text = "$n restored to $($e.Value)." })
            } else {
                Remove-ItemProperty -Path $script:PrintRpcPolicyPath -Name $n -Force -ErrorAction SilentlyContinue
                $still = $true
                try { [void](Get-ItemProperty -Path $script:PrintRpcPolicyPath -Name $n -ErrorAction Stop) } catch { $still = $false }
                $items.Add([pscustomobject]@{ Item = $n; Ok = (-not $still); Text = $(if ($still) { "$n could not be removed - something re-applies it, most likely Group Policy." } else { "$n removed (it was not set before)." }) })
            }
            $regChanged = $true
        } catch {
            $items.Add([pscustomobject]@{ Item = $n; Ok = $false; Text = "Could not restore ${n}: $($_.Exception.Message)" })
        }
    }
    foreach ($e in @($snap.Rules)) {
        if (-not $e) { continue }
        $n = [string]$e.Name
        if (-not (Test-PrintToolkitRuleNameAllowed -Name $n)) { continue }
        try {
            $cur = Get-NetFirewallRule -Name $n -ErrorAction SilentlyContinue
            if (-not [bool]$e.Existed) {
                if ($cur) { Remove-NetFirewallRule -Name $n -ErrorAction Stop; $items.Add([pscustomobject]@{ Item = $n; Ok = $true; Text = "Firewall rule $n removed (it did not exist before)." }) }
            } elseif ($cur) {
                $want = $(if ([bool]$e.Enabled) { 'True' } else { 'False' })
                Set-NetFirewallRule -Name $n -Enabled $want -ErrorAction Stop
                $items.Add([pscustomobject]@{ Item = $n; Ok = $true; Text = "Firewall rule $n set back to Enabled=$want." })
            }
        } catch {
            $items.Add([pscustomobject]@{ Item = $n; Ok = $false; Text = "Could not restore rule ${n}: $($_.Exception.Message)" })
        }
    }

    $restartOk = $null
    if ($regChanged) { $restartOk = [bool](Restart-PrintSpooler).Success }
    $failed = @($items | Where-Object { -not $_.Ok }).Count
    [pscustomobject]@{
        Success          = ($failed -eq 0 -and $restartOk -ne $false)
        Path             = $Path
        Items            = $items.ToArray()
        RestartSucceeded = $restartOk
        Message          = ((@($items | ForEach-Object { $_.Text }) -join ' ') +
                            $(if ($null -eq $restartOk) { '' } elseif ($restartOk) { ' Spooler restarted.' } else { ' The spooler did NOT restart - restored values are not in effect until it does.' }))
    }
}

function Set-PrintRpcTcpPort {
    <#
    .SYNOPSIS
        HOST: pin the Print Spooler's RPC listener to one TCP port, and prove it bound.
    .DESCRIPTION
        Microsoft documents RpcTcpPort ("Use a specific port for RPC over TCP
        communication") as the alternative to leaving the spooler on a dynamic port.
        Run it on the PC that SHARES the printer.

        Why pin at all: a dynamic port cannot be tested from a client - the client
        does not know it - so "is the print RPC path blocked?" can only be guessed.
        Pinned, it becomes one TCP connect from the client, and a firewall rule can
        name exactly one port instead of opening a range.

        Success is defined by ACTIVATION, not by the registry write: the function
        restarts the spooler and waits for spoolsv.exe to be listening on the port.
        A value that reads back correctly while the spooler listens elsewhere is
        reported as a failure, with both facts.

        Refuses a port another process is using or that sits in an excluded range.
        Warns for a port inside the dynamic range, where an ephemeral connection can
        take it before the spooler binds. The prior state is snapshotted first.
    .OUTPUTS
        PSCustomObject { Success; Port; Written; RestartSucceeded; Listening;
                         ListenerPorts[]; InDynamicRange; SnapshotPath; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][ValidateRange(1024, 65535)][int]$Port,
        [int]$ListenWaitSec = 20
    )
    $fail = {
        param([string]$Text)
        [pscustomobject]@{ Success = $false; Port = $Port; Written = $false; RestartSucceeded = $null; Listening = $false
                           ListenerPorts = @(); InDynamicRange = $null; SnapshotPath = ''; Message = $Text }
    }
    if (-not (Test-PrintToolkitElevated)) { return (& $fail 'Pinning the spooler port needs elevation. Run as Administrator, or as SYSTEM from Backstage.') }

    $before = Get-PrintSpoolerListener
    $spoolPid = $before.ProcessId
    $users = @(Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue | Where-Object { $_ -and $_.OwningProcess -ne $spoolPid })
    if ($users.Count) {
        $owners = @($users | ForEach-Object { $p = Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue; if ($p) { "$($p.ProcessName) ($($p.Id))" } else { "pid $($_.OwningProcess)" } } | Sort-Object -Unique)
        return (& $fail "TCP $Port is already in use by $($owners -join ', '). Choose another port.")
    }

    # Excluded ranges: parse the numeric columns, which do not change with the
    # display language, rather than the localised headings.
    try {
        foreach ($line in @(& netsh.exe int ipv4 show excludedportrange protocol=tcp 2>$null)) {
            if ([string]$line -match '^\s*(\d+)\s+(\d+)') {
                $lo = [int]$matches[1]; $hi = [int]$matches[2]
                if ($Port -ge $lo -and $Port -le $hi) { return (& $fail "TCP $Port is inside an excluded port range ($lo-$hi) and cannot be bound. Choose another port.") }
            }
        }
    } catch {}
    $inDynamic = $null
    try {
        $nums = [System.Collections.Generic.List[int]]::new()
        foreach ($line in @(& netsh.exe int ipv4 show dynamicport tcp 2>$null)) {
            if ([string]$line -match ':\s*(\d+)\s*$') { $nums.Add([int]$matches[1]) }
        }
        if ($nums.Count -ge 2) { $inDynamic = ($Port -ge $nums[0] -and $Port -lt ($nums[0] + $nums[1])) }
    } catch {}

    if (-not $PSCmdlet.ShouldProcess('Print Spooler', "Pin RPC over TCP to port $Port")) { return (& $fail 'Skipped (WhatIf).') }

    $snap = Save-PrintConfigSnapshot -Label "rpctcpport-$Port" -RegistryName @('RpcTcpPort')
    try {
        if (-not (Test-Path -LiteralPath $script:PrintRpcPolicyPath)) { New-Item -Path $script:PrintRpcPolicyPath -Force | Out-Null }
        Set-ItemProperty -Path $script:PrintRpcPolicyPath -Name 'RpcTcpPort' -Value $Port -Type DWord -ErrorAction Stop
    } catch {
        $o = & $fail "Could not write RpcTcpPort: $($_.Exception.Message)"
        $o.SnapshotPath = $snap.Path
        return $o
    }
    $readBack = $null
    try { $readBack = (Get-ItemProperty -Path $script:PrintRpcPolicyPath -Name 'RpcTcpPort' -ErrorAction Stop).RpcTcpPort } catch {}
    $written = ("$readBack" -eq "$Port")

    $restart = Restart-PrintSpooler
    $listening = $false; $ports = @()
    $deadline = (Get-Date).AddSeconds($ListenWaitSec)
    do {
        $l = Get-PrintSpoolerListener
        $ports = @($l.Ports)
        if ($ports -contains $Port) { $listening = $true; break }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)

    $dynNote = if ($inDynamic) { " Note: $Port is inside this PC's dynamic port range, so an outgoing connection could take it before the spooler binds after a future restart - a port below that range is safer." } else { '' }
    $msg = if (-not $written) {
               "RpcTcpPort did not stick - it reads '$readBack' rather than $Port. Something is re-applying it, most likely Group Policy."
           } elseif (-not [bool]$restart.Success) {
               "RpcTcpPort is written as $Port, but the Print Spooler did not restart ($($restart.Message)). The setting is NOT in effect."
           } elseif ($listening) {
               "PROVEN LIVE: the Print Spooler is now listening on TCP $Port.$dynNote Next: enable the built-in firewall rules here, then on the client run Shares -> Test connection with port $Port."
           } else {
               "RpcTcpPort is written as $Port and the spooler restarted, but it is listening on $(if ($ports.Count) { $ports -join ', ' } else { 'no TCP port' }) - not $Port. The setting is not in effect; nothing is proven. Restore the previous state from this dialog if needed.$dynNote"
           }
    [pscustomobject]@{
        Success          = ($written -and [bool]$restart.Success -and $listening)
        Port             = $Port
        Written          = $written
        RestartSucceeded = [bool]$restart.Success
        Listening        = $listening
        ListenerPorts    = @($ports)
        InDynamicRange   = $inDynamic
        SnapshotPath     = $snap.Path
        Message          = $msg
    }
}

function Enable-PrintRpcFirewall {
    <#
    .SYNOPSIS
        HOST: allow clients to reach this PC's spooler - built-in rules first, one narrow rule only when asked.
    .DESCRIPTION
        Enables the SCOPED built-in rules for the active profile, by Name:
          Domain           -> FPS-RPCSS-In-TCP-NoScope, FPS-SpoolSvc-In-TCP-NoScope
          Private / Public -> FPS-RPCSS-In-TCP,         FPS-SpoolSvc-In-TCP   (LocalSubnet)
        It never enables the -V2 variants, which admit any remote address on every
        profile.

        With -CreateRuleForPort, adds ONE rule named PrinterToolkit-SpoolerRpc-TCP-<port>
        for spoolsv.exe on that TCP port. Its remote scope defaults to LocalSubnet or
        takes a specific address/subnet - 'Any' is refused. This is for when the
        client's measurement has shown the pinned port unreachable; it is not a first
        move.

        Refuses to act while an adapter is on the Public profile: fix the profile
        instead (SHR001). Every change is snapshotted for exact rollback.
    .OUTPUTS
        PSCustomObject { Success; Enabled[]; CreatedRule; SnapshotPath; State; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [int]$Port = 0,
        [switch]$CreateRuleForPort,
        [string]$RemoteAddress = 'LocalSubnet'
    )
    $fail = {
        param([string]$Text)
        [pscustomobject]@{ Success = $false; Enabled = @(); CreatedRule = ''; SnapshotPath = ''; State = $null; Message = $Text }
    }
    if (-not (Test-PrintToolkitElevated)) { return (& $fail 'Changing firewall rules needs elevation. Run as Administrator, or as SYSTEM from Backstage.') }
    $active = @(Get-ActiveFirewallProfile)
    if (-not $active.Count) { return (& $fail 'No active network profile could be read, so the rules to enable could not be chosen.') }
    if ($active -contains 'Public') {
        return (& $fail "An adapter on this PC is on the Public profile. Windows blocks inbound sharing there by design - set the network to Private first (SHR001) rather than opening rules on a public network.")
    }
    $ra = ([string]$RemoteAddress).Trim()
    if (-not $ra -or $ra -match '^(?i)(any|\*)$') { return (& $fail "Refusing remote scope '$RemoteAddress'. Use LocalSubnet or a specific client address or subnet.") }
    if ($CreateRuleForPort -and ($Port -lt 1024 -or $Port -gt 65535)) { return (& $fail 'A narrow rule needs the pinned port (1024-65535).') }

    $want = [System.Collections.Generic.List[string]]::new()
    foreach ($base in @('FPS-RPCSS-In-TCP', 'FPS-SpoolSvc-In-TCP')) {
        if ($active -contains 'Domain') { $want.Add("$base-NoScope") }
        if ($active -contains 'Private') { $want.Add($base) }
    }
    $newRule = if ($CreateRuleForPort) { $script:PrintRpcToolkitRulePrefix + $Port } else { '' }

    $what = 'Enable ' + ($want -join ', ') + $(if ($newRule) { " and create $newRule ($ra)" } else { '' })
    if (-not $PSCmdlet.ShouldProcess('Windows Firewall', $what)) { return (& $fail 'Skipped (WhatIf).') }

    $snapNames = @($want) + $(if ($newRule) { @($newRule) } else { @() })
    $snap = Save-PrintConfigSnapshot -Label $(if ($newRule) { "firewall-$Port" } else { 'firewall' }) -RuleName $snapNames

    $done = [System.Collections.Generic.List[string]]::new()
    $problems = [System.Collections.Generic.List[string]]::new()
    foreach ($n in $want) {
        try {
            $one = Get-NetFirewallRule -Name $n -ErrorAction SilentlyContinue
            if (-not $one) { $problems.Add("$n does not exist on this PC"); continue }
            Enable-NetFirewallRule -Name $n -ErrorAction Stop
            $done.Add($n)
        } catch { $problems.Add("${n}: $($_.Exception.Message)") }
    }
    $created = ''
    if ($newRule) {
        try {
            if (Get-NetFirewallRule -Name $newRule -ErrorAction SilentlyContinue) {
                Set-NetFirewallRule -Name $newRule -Enabled True -RemoteAddress $ra -ErrorAction Stop
            } else {
                $profiles = ($active | Where-Object { $_ -ne 'Public' }) -join ','
                New-NetFirewallRule -Name $newRule -DisplayName "Printer Toolkit - Print Spooler RPC (TCP $Port)" `
                    -Group 'Printer Toolkit' -Direction Inbound -Action Allow -Protocol TCP -LocalPort $Port `
                    -Program (Join-Path $env:SystemRoot 'System32\spoolsv.exe') -RemoteAddress $ra -Profile $profiles `
                    -Description 'Created by Printer Toolkit after a client measured the pinned spooler port as unreachable. Restore the snapshot to remove it.' `
                    -ErrorAction Stop | Out-Null
            }
            $created = $newRule
        } catch { $problems.Add("Could not create ${newRule}: $($_.Exception.Message)") }
    }

    $state = Get-PrintRpcFirewallState -PinnedPort $Port
    $ok = (-not $problems.Count) -and $state.State -eq 'Ok' -and $state.Spooler.Allowed -and $state.Epm.Allowed
    [pscustomobject]@{
        Success      = $ok
        Enabled      = $done.ToArray()
        CreatedRule  = $created
        SnapshotPath = $snap.Path
        State        = $state
        Message      = (@(
            $(if ($done.Count) { 'Enabled: ' + ($done -join ', ') + '.' } else { '' })
            $(if ($created) { "Created $created for spoolsv.exe on TCP $Port, remote scope $ra." } else { '' })
            $(if ($problems.Count) { 'Problems: ' + ($problems -join '; ') + '.' } else { '' })
            $state.Message
            'A rule being enabled here does not prove a client can reach it - measure from the client with Shares -> Test connection.'
        ) | Where-Object { $_ }) -join ' '
    }
}

function Test-PrintRpcPath {
    <#
    .SYNOPSIS
        CLIENT: measure the network path to the host of an EXACT printer connection.
    .DESCRIPTION
        The host is taken from the UNC being tested, never from any other field.
        Verifying one host and then retrying a connection to another is how an
        unrelated result used to decide a retry.

        Measures SMB 445 and the RPC endpoint mapper 135, and - if the host's spooler
        has been pinned with RpcTcpPort - that exact port. Without a pinned port the
        spooler's endpoint is a dynamic port this client cannot know, so that part is
        reported as Skipped with the reason: never as a pass, never as a failure.

        Outcomes come from the socket, not from elapsed time:
          Timeout  no answer at all - silently dropped, or the host is down
          Failed   refused (nothing listening, or a rejecting firewall), or the name
                   does not resolve
        The default limit is 3 s because a refusal on Windows is reported at about 2 s;
        a shorter limit would report every refusal as a timeout.
    .OUTPUTS
        PSCustomObject { Host; Unc; Probes[]; TransportMeasured; Message }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ConnectionName,
        [int]$PinnedPort = 0,
        [int]$TimeoutMs = 3000
    )
    $unc = ([string]$ConnectionName).Trim()
    $hostName = $unc.TrimStart('\')
    $cut = $hostName.IndexOf('\')
    if ($cut -ge 0) { $hostName = $hostName.Substring(0, $cut) }

    $probes = [System.Collections.Generic.List[object]]::new()
    $one = {
        param([string]$Label, [int]$PortNo)
        $t = Test-TcpPortFast -HostName $hostName -Port $PortNo -TimeoutMs $TimeoutMs
        $st = 'Failed'; $dt = ''
        if ($t.Open) {
            $st = 'Ok'; $dt = "answered in $($t.Ms) ms"
        } elseif (-not $t.Completed) {
            $st = 'Timeout'; $dt = "no answer within $TimeoutMs ms (silently dropped, or the host is down)"
        } elseif ($t.SocketError -eq 'ConnectionRefused') {
            $dt = "refused after $($t.Ms) ms (nothing listening, or a firewall rejecting)"
        } elseif ($t.SocketError -eq 'HostNotFound') {
            $dt = "the name '$hostName' does not resolve"
        } else {
            $dt = "did not connect ($($t.SocketError)) after $($t.Ms) ms"
        }
        New-PrintProbeResult -Probe $Label -State $st -Code "$PortNo" -Detail $dt -Target "${hostName}:$PortNo"
    }
    $probes.Add((& $one 'SMB 445' 445))
    $probes.Add((& $one 'RPC endpoint mapper 135' 135))
    if ($PinnedPort -gt 0) {
        $probes.Add((& $one "Spooler RPC $PinnedPort" $PinnedPort))
    } else {
        $probes.Add((New-PrintProbeResult -Probe 'Spooler RPC (dynamic)' -State 'Skipped' -Target $hostName `
            -Detail 'The host spooler uses a dynamic port, which cannot be tested from here. Pin it on the host (Share Host -> Fix print RPC) to make this measurable.'))
    }
    $arr = $probes.ToArray()
    $bad = @($arr | Where-Object { $_.State -in @('Failed', 'Timeout') })
    [pscustomobject]@{
        Host              = $hostName
        Unc               = $unc
        Probes            = $arr
        TransportMeasured = ($PinnedPort -gt 0)
        Message           = $(if ($bad.Count) { "To ${hostName}: " + (@($bad | ForEach-Object { "$($_.Probe) $($_.State.ToLower())" }) -join ', ') + '.' }
                              elseif ($PinnedPort -gt 0) { "To ${hostName}: SMB 445, endpoint mapper 135 and spooler port $PinnedPort all answered." }
                              else { "To ${hostName}: SMB 445 and endpoint mapper 135 answered. The spooler's own port was not testable (not pinned)." })
    }
}

function Test-PrinterOpenAsUser {
    <#
    .SYNOPSIS
        CLIENT: can the TARGET USER open this exact printer on the live host? Measured in their session.
    .DESCRIPTION
        The add runs as the user, so the evidence has to be gathered as the user. A
        console probe answers a different question - the console may be SYSTEM or an
        admin with different credentials - and checking in the console after saving a
        credential for the user reported on an identity that was never used.

        Uses OpenPrinter2 with PRINTER_OPTION_NO_CACHE, because plain OpenPrinter can
        answer from cached printer data for a host that is not reachable at all.

        A handle shows the connection answered for that identity. It does not show that
        a driver will install or that a page will print.
    .OUTPUTS
        New-PrintProbeResult with Probe 'Open as user'.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ConnectionName,
        [string]$User,
        [int]$TimeoutSec = 60,
        [scriptblock]$OnPoll
    )
    $r = Invoke-ScriptInPrinterUserSession -User $User -TimeoutSec $TimeoutSec -OnPoll $OnPoll -Vars @{ Target = $ConnectionName } -ScriptBlock {
        $opened = $false; $err = 0; $loaded = $true; $msg = ''
        try {
            if (-not ('PrtNoCache.Spool' -as [type])) {
                $psig = @(
                    '[StructLayout(LayoutKind.Sequential)] public struct PRINTER_OPTIONS { public uint cbSize; public uint dwFlags; }'
                    '[DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "OpenPrinter2W")]'
                    'public static extern bool OpenPrinter2(string pPrinterName, out System.IntPtr phPrinter, System.IntPtr pDefault, ref PRINTER_OPTIONS pOptions);'
                    '[DllImport("winspool.drv", SetLastError = true)]'
                    'public static extern bool ClosePrinter(System.IntPtr hPrinter);'
                ) -join [Environment]::NewLine
                Add-Type -Namespace 'PrtNoCache' -Name 'Spool' -MemberDefinition $psig -ErrorAction Stop
            }
            $opt = New-Object 'PrtNoCache.Spool+PRINTER_OPTIONS'
            $opt.cbSize = [uint32][System.Runtime.InteropServices.Marshal]::SizeOf([type]'PrtNoCache.Spool+PRINTER_OPTIONS')
            $opt.dwFlags = [uint32]1
            $hh = [IntPtr]::Zero
            $opened = [PrtNoCache.Spool]::OpenPrinter2($Target, [ref]$hh, [IntPtr]::Zero, [ref]$opt)
            if (-not $opened) { $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error() }
            if ($hh -ne [IntPtr]::Zero) { [void][PrtNoCache.Spool]::ClosePrinter($hh) }
        } catch { $loaded = $false; $msg = $_.Exception.Message }
        [pscustomobject]@{ Loaded = $loaded; Opened = [bool]$opened; Err = [int]$err; Msg = $msg; RanAs = "$env:USERDOMAIN\$env:USERNAME" } | ConvertTo-Json -Compress
    }
    $who = if ($r.User) { [string]$r.User } else { [string]$User }
    if (-not $r.Data) {
        $st = if ("$($r.Message)" -match 'No interactive user') { 'Skipped' } else { 'Timeout' }
        return (New-PrintProbeResult -Probe 'Open as user' -State $st -Detail ([string]$r.Message) -Identity $who -Target $ConnectionName)
    }
    $d = $r.Data
    $ranAs = if ($d.RanAs) { [string]$d.RanAs } else { $who }
    if (-not [bool]$d.Loaded) {
        return (New-PrintProbeResult -Probe 'Open as user' -State 'Skipped' -Detail "The open test could not run in the user's session: $($d.Msg)" -Identity $ranAs -Target $ConnectionName)
    }
    if ([bool]$d.Opened) {
        return (New-PrintProbeResult -Probe 'Open as user' -State 'Ok' -Code '0' -Identity $ranAs -Target $ConnectionName `
            -Detail "$ranAs opened it on the live host (not from cache). That is not proof that a driver will install or a page will print.")
    }
    $hn = ConvertTo-PrintHostKey $ConnectionName
    New-PrintProbeResult -Probe 'Open as user' -State 'Failed' -Code "$([int]$d.Err)" -Identity $ranAs -Target $ConnectionName `
        -Detail ("$ranAs could not open it on the live host: " + (Get-SmbErrorExplanation -Code ([int]$d.Err) -HostName $hn))
}

function Get-PrintConnectionDiagnosis {
    <#
    .SYNOPSIS
        Decide WHY a connection failed - or that nothing measurable is failing - from evidence.
    .DESCRIPTION
        Replaces diagnosis-by-prose. The GUI picks a remedy from Category and never
        from message text: the old approach matched "0x80070040" in a message and
        offered a machine-wide transport change straight after that same message had
        ruled transport out.

        Rules in order. Every conclusion names its evidence and whether it was
        Measured (a probe observed it) or Indicated (an error code suggests it).
        Error 64 on its own concludes nothing - it only says a network name became
        unavailable, at some layer.

          Transport  Measured   445, 135 or the host's pinned spooler port could not be reached
          Transport  Measured   a no-cache open failed at the RPC layer (1722/1753/1727/1726)
                                while SMB on 445 answered and nothing showed the RPC layer reached
          None       Measured   no error was given and every probe that ran succeeded
          Identity   Measured   the console opened the printer and the user did not
          Identity   Indicated  access denied / logon failure codes
          Share      Indicated  the printer or share name does not exist on the host
          Driver     Measured   the user opened the printer, yet the add still failed
          Service    Indicated  0x0000011b RPC privacy mismatch - host patch or policy
          Driver     Indicated  Point and Print / driver codes
          Unknown               none of the above could be established

        Opens are OpenPrinter2 with PRINTER_OPTION_NO_CACHE. A handle obtained while 445
        does not answer is still called out: plain OpenPrinter was measured returning a
        cached handle for a host that did not even resolve.
    .OUTPUTS
        PSCustomObject { Category; Confidence; Why; Evidence[]; Recommended[]; Unc; Host; UserName }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ConnectionName,
        [string]$AddError = '',
        $UserOpen,
        $ConsoleOpen,
        $Path,
        [string]$UserName = ''
    )
    $hostName = ConvertTo-PrintHostKey $ConnectionName
    if ($ConnectionName -match '^\\\\([^\\]+)\\') { $hostName = $matches[1] }

    $ev = [System.Collections.Generic.List[object]]::new()
    if ($Path) { foreach ($p in @($Path.Probes)) { if ($p) { $ev.Add($p) } } }
    if ($ConsoleOpen) { $ev.Add($ConsoleOpen) }
    if ($UserOpen) { $ev.Add($UserOpen) }

    $text        = [string]$AddError
    $userCode    = if ($UserOpen -and $UserOpen.Code) { [string]$UserOpen.Code } else { '' }
    $consoleCode = if ($ConsoleOpen -and $ConsoleOpen.Code) { [string]$ConsoleOpen.Code } else { '' }
    $smb  = @($ev | Where-Object { $_.Probe -eq 'SMB 445' })[0]
    $epm  = @($ev | Where-Object { $_.Probe -eq 'RPC endpoint mapper 135' })[0]
    $pin  = @($ev | Where-Object { $_.Probe -like 'Spooler RPC *' -and $_.Probe -notlike '*dynamic*' })[0]
    $uOk   = ($UserOpen -and $UserOpen.State -eq 'Ok')
    $uFail = ($UserOpen -and $UserOpen.State -eq 'Failed')
    $cOk   = ($ConsoleOpen -and $ConsoleOpen.State -eq 'Ok')
    $who   = if ($UserName) { $UserName } else { 'the user' }
    $hardFail = (@($ev | Where-Object { $_.State -in @('Failed', 'Timeout') }).Count -gt 0)

    # RPC-layer failures versus codes that prove the RPC layer WAS reached.
    $rpcLayerCodes = @('1722', '1753', '1727', '1726')
    $reachedCodes  = @('5', '1801', '1326', '1385')
    $rpcCode = ''
    if ($rpcLayerCodes -contains $userCode) { $rpcCode = $userCode } elseif ($rpcLayerCodes -contains $consoleCode) { $rpcCode = $consoleCode }
    $rpcReached = ($uOk -or $cOk -or ($reachedCodes -contains $userCode) -or ($reachedCodes -contains $consoleCode))

    $identityRx = '(?i)access is denied|0x80070005|logon failure|user name or password is incorrect|0x8007052e'
    $shareRx    = '(?i)printer name is invalid|0x80070709|network name cannot be found|0x80070043'
    $driverRx   = '(?i)0x80070bcb|0x00000bcb|\bdriver\b'
    $privacyRx  = '(?i)0x0000011b|0x8007011b'

    $cat = 'Unknown'; $conf = 'Unknown'; $why = ''
    $rec = [System.Collections.Generic.List[string]]::new()

    if ($smb -and $smb.State -in @('Failed', 'Timeout')) {
        $cat = 'Transport'; $conf = 'Measured'
        $why = "SMB on 445 could not be reached from this PC ($($smb.Detail))."
        if ($cOk -or $uOk) { $why += ' A printer handle was still obtained - treat it as cached printer data, not evidence the host is reachable.' }
        $rec.Add('CheckHost')
    } elseif ($pin -and $pin.State -in @('Failed', 'Timeout') -and $epm -and $epm.State -eq 'Ok') {
        $cat = 'Transport'; $conf = 'Measured'
        $why = "The host's pinned spooler RPC port could not be reached ($($pin.Detail)), although 135 answered."
        $rec.Add('NarrowRule')
    } elseif ($epm -and $epm.State -in @('Failed', 'Timeout') -and $smb -and $smb.State -eq 'Ok') {
        $cat = 'Transport'; $conf = 'Measured'
        $why = "The RPC endpoint mapper on 135 could not be reached ($($epm.Detail)), although SMB on 445 answered."
        $rec.Add('EnableHostRules')
    } elseif ($rpcCode -and -not $rpcReached -and $smb -and $smb.State -eq 'Ok') {
        $cat = 'Transport'; $conf = 'Measured'
        $meaning = switch ($rpcCode) {
            '1722'  { 'the RPC server is unavailable' }
            '1753'  { 'no endpoint is registered with the endpoint mapper' }
            '1727'  { 'the call failed and did not execute' }
            default { 'the remote procedure call failed' }
        }
        $why = "Opening the printer on the live host failed at the RPC layer with $rpcCode ($meaning), although SMB on 445 answered. The spooler's RPC traffic is not getting through - a host firewall rule, or the spooler not listening."
        $rec.Add('EnableHostRules')
        if (-not $pin) { $rec.Add('PinRpcPort') }
    } elseif (-not $text -and -not $hardFail -and ($uOk -or ($null -eq $UserOpen -and $cOk))) {
        $cat = 'None'; $conf = 'Measured'
        $why = $(if ($uOk) { "Every probe that ran succeeded, including opening the printer on the live host as $who." }
                 else { 'Every probe that ran succeeded. The target user was not tested - select a signed-in user to include them.' })
        if (-not $pin) { $why += " The spooler's own RPC port was not tested, because it is not pinned on the host." }
        $rec.Add('DeliveryTest')
    } elseif ($uFail -and $cOk) {
        $cat = 'Identity'; $conf = 'Measured'
        $why = "This console opened the printer on the live host and $who did not - same printer, same moment."
        $rec.Add('Credentials')
    } elseif ($text -match $identityRx -or $userCode -in @('5', '1326', '1385')) {
        $cat = 'Identity'; $conf = 'Indicated'
        $why = "The error is an authentication or authorisation refusal for $who."
        $rec.Add('Credentials')
    } elseif ($text -match $shareRx -or $userCode -in @('1801', '67', '1799')) {
        $cat = 'Share'; $conf = 'Indicated'
        $why = 'The host says that printer or share name does not exist as typed.'
        $rec.Add('CheckHost')
    } elseif ($uOk) {
        $cat = 'Driver'; $conf = 'Measured'
        $why = "$who opened the printer on the live host, so the connection answered; the add failed at a later stage - usually installing the driver from the host."
        $rec.Add('InstallDriverLocalPort')
    } elseif ($text -match $privacyRx) {
        $cat = 'Service'; $conf = 'Indicated'
        $why = '0x0000011b is the CVE-2021-1678 RPC privacy change: the host and this PC are at incompatible patch or policy levels.'
        $rec.Add('PatchHost')
    } elseif ($text -match $driverRx) {
        $cat = 'Driver'; $conf = 'Indicated'
        $why = 'The error points at the driver or Point and Print rather than the connection.'
        $rec.Add('InstallDriverLocalPort')
    } else {
        $why = 'No probe isolated the failing stage. Error 64 alone only says a network name became unavailable - it does not say where.'
        $rec.Add('CheckHost')
        if (-not $pin) { $rec.Add('PinRpcPort') }
    }

    [pscustomobject]@{
        Category    = $cat
        Confidence  = $conf
        Why         = $why
        Evidence    = $ev.ToArray()
        Recommended = $rec.ToArray()
        Unc         = $ConnectionName
        Host        = $hostName
        UserName    = $who
    }
}

function Format-PrintDiagnosis {
    <# .SYNOPSIS Render a diagnosis as readable text. Named pipes are never recommended here. #>
    [CmdletBinding()] param($Diagnosis)
    if (-not $Diagnosis) { return '' }
    $d = $Diagnosis
    $h = [string]$d.Host
    $lines = [System.Collections.Generic.List[string]]::new()
    if ([string]$d.Category -eq 'None') { $lines.Add('RESULT: no fault measured') }
    else { $lines.Add("DIAGNOSIS: $($d.Category)  ($($d.Confidence))") }
    $lines.Add('')
    $lines.Add([string]$d.Why)
    $lines.Add('')
    $lines.Add('Evidence:')
    foreach ($e in @($d.Evidence)) {
        if (-not $e) { continue }
        $idTxt = if ($e.Identity) { " as $($e.Identity)" } else { '' }
        $lines.Add(('  [{0,-7}] {1}{2}: {3}' -f $e.State, $e.Probe, $idTxt, $e.Detail))
    }
    $lines.Add('')
    $lines.Add('What to do:')
    foreach ($k in @($d.Recommended)) {
        switch ($k) {
            'CheckHost'       { $lines.Add("  - On $h, run this console and use Share Host -> Check this PC: the exact share name, the spooler, the USB printer, and the firewall rules for the active profile.") }
            'PinRpcPort'      { $lines.Add("  - To make the print RPC path MEASURABLE, pin $h's spooler to a fixed port (Share Host -> Fix print RPC), then run Shares -> Test connection here with that port.") }
            'NarrowRule'      { $lines.Add("  - On ${h}: Share Host -> Fix print RPC -> create the narrow rule for the pinned port, then test again from here.") }
            'EnableHostRules' { $lines.Add("  - On ${h}: Share Host -> Fix print RPC -> enable the built-in rules for the active profile, then test again from here.") }
            'Credentials'     { $lines.Add("  - Credentials... -> Save for $($d.UserName). Check the printer's Security tab on $h. In a workgroup, $($d.UserName) needs an account that exists on $h.") }
            'InstallDriverLocalPort' { $lines.Add('  - Install the vendor driver here (Install driver...), then Add... with the local-port scope - and run the delivery test, because a created queue is not a printed page.') }
            'PatchHost'       { $lines.Add("  - Bring $h and this PC to the same patch level. Do not turn RpcAuthnLevelPrivacyEnabled off on either PC.") }
            'DeliveryTest'    { $lines.Add('  - Nothing is failing at the connection stage. If printing still does not work, add the printer and run the optional delivery test to see whether a job leaves the queue.') }
        }
    }
    return ($lines -join "`r`n")
}

function Get-PrintConnectionEvidence {
    <#
    .SYNOPSIS
        Gather every measurable probe for one exact printer connection, then diagnose it.
    .DESCRIPTION
        One entry point for "why does this connection fail?", used by Test
        connection and by any failed add that did not already measure in-session.
        Every probe is bound to the same UNC. The user-session probe runs only when a
        user is given.
    .OUTPUTS
        PSCustomObject { Unc; UserOpen; ConsoleOpen; Path; Diagnosis; Text }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ConnectionName,
        [string]$User,
        [string]$AddError = '',
        [int]$PinnedPort = 0,
        [scriptblock]$OnPoll
    )
    $unc = $ConnectionName.Trim()
    $userProbe = $null
    if ($User) { $userProbe = Test-PrinterOpenAsUser -ConnectionName $unc -User $User -OnPoll $OnPoll }
    $co = Test-PrinterShareOpen -Unc $unc
    $consoleProbe = New-PrintProbeResult -Probe 'Open as console' -State $co.State -Detail $co.Message -Target $unc `
        -Code $(if ($null -ne $co.Win32Error) { "$($co.Win32Error)" } else { '' }) -Identity "$env:USERDOMAIN\$env:USERNAME"
    $path = Test-PrintRpcPath -ConnectionName $unc -PinnedPort $PinnedPort
    $who = if ($userProbe -and $userProbe.Identity) { [string]$userProbe.Identity } else { [string]$User }
    $diag = Get-PrintConnectionDiagnosis -ConnectionName $unc -AddError $AddError -UserOpen $userProbe -ConsoleOpen $consoleProbe -Path $path -UserName $who
    [pscustomobject]@{ Unc = $unc; UserOpen = $userProbe; ConsoleOpen = $consoleProbe; Path = $path; Diagnosis = $diag; Text = (Format-PrintDiagnosis $diag) }
}

function Invoke-PrinterDeliveryTest {
    <#
    .SYNOPSIS
        Send a test page and watch the job - evidence that "created" became "handed off".
    .DESCRIPTION
        Optional, after an add. Records the queue's job IDs, sends a Windows test
        page, then watches for a new job to appear and leave the queue.

        What it can show: the job was created and left THIS PC's queue. What it
        cannot show: that the host accepted it or that paper came out - say so, and
        check the printer. A fast job can appear and leave between polls, which is
        reported as "not observed", not as a failure.

        Runs as the user for a per-user connection (only their session can see it),
        or in-process for a machine-wide local queue.
    .OUTPUTS
        PSCustomObject { Success; Sent; JobId; LeftQueue; LastStatus; State; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$PrinterName,
        [string]$User,
        [switch]$InProcess,
        [int]$WaitSec = 20,
        [scriptblock]$OnPoll
    )
    if (-not $PSCmdlet.ShouldProcess($PrinterName, 'Send a test page and watch the job')) {
        return [pscustomobject]@{ Success = $false; Sent = $false; JobId = $null; LeftQueue = $null; LastStatus = ''; State = 'Skipped'; Message = 'Skipped (WhatIf).' }
    }
    $core = {
        $out = [ordered]@{ Sent = $false; JobId = $null; LeftQueue = $null; LastStatus = ''; Error = '' }
        try {
            $before = @(Get-PrintJob -PrinterName $Target -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
            $esc = $Target -replace "'", "''"
            $p = Get-CimInstance Win32_Printer -Filter "Name='$esc'" -ErrorAction Stop
            $res = Invoke-CimMethod -InputObject $p -MethodName PrintTestPage -ErrorAction Stop
            if ([int]$res.ReturnValue -ne 0) {
                $out.Error = "PrintTestPage returned $($res.ReturnValue)"
            } else {
                $out.Sent = $true
                $deadline = (Get-Date).AddSeconds([int]$WaitSec)
                do {
                    Start-Sleep -Milliseconds 500
                    $now = @(Get-PrintJob -PrinterName $Target -ErrorAction SilentlyContinue)
                    if (-not $out.JobId) {
                        $new = @($now | Where-Object { $before -notcontains [int]$_.Id })
                        if ($new.Count) { $out.JobId = [int]$new[0].Id; $out.LastStatus = [string]$new[0].JobStatus }
                    }
                    if ($out.JobId) {
                        $still = @($now | Where-Object { [int]$_.Id -eq $out.JobId })
                        if ($still.Count) { $out.LastStatus = [string]$still[0].JobStatus } else { $out.LeftQueue = $true; break }
                    }
                } while ((Get-Date) -lt $deadline)
                if ($out.JobId -and -not $out.LeftQueue) { $out.LeftQueue = $false }
            }
        } catch { $out.Error = $_.Exception.Message }
        [pscustomobject]$out | ConvertTo-Json -Compress
    }

    $data = $null; $ranAs = "$env:USERDOMAIN\$env:USERNAME"
    if ($InProcess) {
        $Target = $PrinterName
        $raw = & $core
        $data = ConvertFrom-NoisyJson ([string]$raw)
    } else {
        $r = Invoke-ScriptInPrinterUserSession -User $User -TimeoutSec ($WaitSec + 40) -OnPoll $OnPoll -Vars @{ Target = $PrinterName; WaitSec = $WaitSec } -ScriptBlock $core
        $data = $r.Data
        if ($r.User) { $ranAs = [string]$r.User }
        if (-not $data) {
            return [pscustomobject]@{ Success = $false; Sent = $false; JobId = $null; LeftQueue = $null; LastStatus = ''; State = 'Timeout'
                Message = "The delivery test result could not be read back from $ranAs's session: $($r.Message)" }
        }
    }
    if (-not $data) {
        return [pscustomobject]@{ Success = $false; Sent = $false; JobId = $null; LeftQueue = $null; LastStatus = ''; State = 'Timeout'; Message = 'The delivery test returned no result.' }
    }
    $caveat = ' That shows the job was handed off from this PC - it does not prove the host accepted it or that a page came out. Check the printer.'
    $msg = if ($data.Error) { "Test page to '$PrinterName' failed: $($data.Error)" }
           elseif (-not [bool]$data.Sent) { "Test page to '$PrinterName' was not sent." }
           elseif ($data.JobId -and [bool]$data.LeftQueue) { "Test page sent as $ranAs; job $($data.JobId) left the queue on this PC.$caveat" }
           elseif ($data.JobId) { "Test page sent as $ranAs; job $($data.JobId) is still in the queue after $WaitSec s (last status: $($data.LastStatus)). It has not been handed off." }
           else { "Test page sent as $ranAs, but no job was observed in the queue - it may have left faster than it could be seen.$caveat" }
    $state = if ($data.Error -or -not [bool]$data.Sent) { 'Failed' } elseif ($data.JobId -and -not [bool]$data.LeftQueue) { 'Failed' } else { 'Ok' }
    [pscustomobject]@{
        Success    = ($state -eq 'Ok')
        Sent       = [bool]$data.Sent
        JobId      = $data.JobId
        LeftQueue  = $data.LeftQueue
        LastStatus = [string]$data.LastStatus
        State      = $state
        Message    = $msg
    }
}

# endregion


# ---------------------------------------------------------------------------
# region  Paper size (v2.1): what Windows will send, against the region
# ---------------------------------------------------------------------------
# Every read here is the registry: a queue's stored default DEVMODE, the copies the
# client-side print provider caches for each connection, and a profile's regional
# paper size. Get-PrintConfiguration reports the same paper code (measured on every
# queue of the build PC) but costs 40 ms to 1 s per queue.

$script:PaperNames = @{ 1 = 'Letter'; 5 = 'Legal'; 8 = 'A3'; 9 = 'A4'; 11 = 'A5' }
# DEVMODE dmDuplex: DMDUP_SIMPLEX 1, DMDUP_VERTICAL 2 (long edge), DMDUP_HORIZONTAL 3
# (short edge). PrintManagement's DuplexingModeEnum is numbered 0, 1, 2 - NOT the same
# numbers (measured) - so duplex is always converted by NAME, never by value.
$script:DuplexNames = @{ 1 = 'OneSided'; 2 = 'TwoSidedLongEdge'; 3 = 'TwoSidedShortEdge' }
$script:DuplexLabels = @{ OneSided = 'one-sided'; TwoSidedLongEdge = 'two-sided (long edge)'; TwoSidedShortEdge = 'two-sided (short edge)' }
$script:PaperSnapshotDir = Join-Path $env:ProgramData 'PrinterToolkit\snapshots'
$script:PrintersRegPath = 'SYSTEM\CurrentControlSet\Control\Print\Printers'
$script:CsrRegPath = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider'
$script:CountryPaperCache = @{}

function ConvertTo-PaperCode {
    <# .SYNOPSIS A paper code from a registry string or number. 0 when it is not a whole number. #>
    [CmdletBinding()] param($Value)
    $n = 0
    if ([int]::TryParse(([string]$Value).Trim(), [ref]$n)) { return $n }
    return 0
}

function Get-PaperCodeFromDimensions {
    <#
    .SYNOPSIS
        A4, Letter or Legal from paper dimensions (tenths of a mm) or a paper name. 0 if neither matches.
    .DESCRIPTION
        Dimensions win over the name: Win32_PrintJob's PaperSize is the driver's own
        wording, while the dimensions are always the sheet.
    #>
    [CmdletBinding()] param([string]$Name = '', [int]$Width = 0, [int]$Length = 0)
    $short = [Math]::Min($Width, $Length); $long = [Math]::Max($Width, $Length)
    if ($short -gt 0) {
        if ([Math]::Abs($short - 2100) -le 5 -and [Math]::Abs($long - 2970) -le 5) { return 9 }
        if ([Math]::Abs($short - 2159) -le 5 -and [Math]::Abs($long - 2794) -le 5) { return 1 }
        if ([Math]::Abs($short - 2159) -le 5 -and [Math]::Abs($long - 3556) -le 5) { return 5 }
    }
    $t = ([string]$Name).Trim()
    if ($t -match '^A4\b') { return 9 }
    if ($t -match '^Letter\b') { return 1 }
    if ($t -match '^Legal\b') { return 5 }
    return 0
}

function Get-PaperSizeName {
    <# .SYNOPSIS A readable name for a DEVMODE / LOCALE_IPAPERSIZE paper code. #>
    [CmdletBinding()] param([int]$Code)
    if ($script:PaperNames.ContainsKey($Code)) { return $script:PaperNames[$Code] }
    if ($Code -ge 256) { return "Driver form $Code" }
    # PrintManagement's PaperSizeEnum uses the same numbering (measured: Letter 1,
    # Legal 5, A3 8, A4 9, A5 11, A6 70). Its underlying type is UInt32.
    $t = 'Microsoft.PowerShell.Cmdletization.GeneratedTypes.PrinterConfiguration.PaperSizeEnum' -as [type]
    if ($t -and $Code -gt 0) {
        try { if ([Enum]::IsDefined($t, [uint32]$Code)) { return [string][Enum]::ToObject($t, [uint32]$Code) } } catch {}
    }
    if ($Code -le 0) { return 'Custom' }
    return "Paper code $Code"
}

function Get-DuplexCode {
    <# .SYNOPSIS The DEVMODE dmDuplex value for a DuplexingMode name. 0 when the name is not one of the three. #>
    [CmdletBinding()] param([string]$Name)
    foreach ($k in @($script:DuplexNames.Keys)) { if ($script:DuplexNames[$k] -eq $Name) { return [int]$k } }
    return 0
}

function Get-DuplexLabel {
    <# .SYNOPSIS Plain wording for a DuplexingMode name, e.g. 'two-sided (long edge)'. #>
    [CmdletBinding()] param([string]$Name)
    if ($Name -and $script:DuplexLabels.ContainsKey($Name)) { return $script:DuplexLabels[$Name] }
    if ($Name) { return $Name }
    return 'not set'
}

function ConvertFrom-PrinterDevMode {
    <#
    .SYNOPSIS
        The paper and duplex fields of a binary DEVMODEW, as the spooler stores it in the registry.
    .DESCRIPTION
        Public DEVMODEW layout: dmSize @68, dmFields @72, dmPaperSize @78,
        dmPaperLength @80 and dmPaperWidth @82 (tenths of a mm), dmDuplex @94,
        dmFormName @102 (32 WCHAR). Paper and duplex were measured against
        Get-PrintConfiguration on every queue of the build PC, and duplex also on a
        print server connection set to two-sided long edge. A buffer too short to
        hold the form name is reported invalid rather than half-read. A custom size
        (code 0) is named from its dimensions when they are A4, Letter or Legal.

        A code of 256 or more is a driver-defined form, such as 271 on a Brother QL
        label printer - not a Windows paper size, and never compared.

        DuplexDeclared is the DM_DUPLEX bit of dmFields: whether the driver uses the
        duplex field at all. A label printer or a PDF writer leaves it clear, and its
        duplex value then means nothing.
    .OUTPUTS
        PSCustomObject { Valid; PaperCode; PaperName; FormName; LengthTenthMm; WidthTenthMm; Fields;
                         DuplexCode; DuplexName; DuplexDeclared }
    #>
    [CmdletBinding()] param($Bytes)
    $b = $null
    if ($Bytes -is [byte[]]) { $b = $Bytes }
    if (-not $b -or $b.Length -lt 166 -or [BitConverter]::ToUInt16($b, 68) -lt 166) {
        return [pscustomobject]@{
            Valid = $false; PaperCode = 0; PaperName = ''; FormName = ''; LengthTenthMm = 0; WidthTenthMm = 0; Fields = [uint32]0
            DuplexCode = 0; DuplexName = ''; DuplexDeclared = $false
        }
    }
    $code = [int][BitConverter]::ToInt16($b, 78)
    $len  = [int][BitConverter]::ToInt16($b, 80)
    $wid  = [int][BitConverter]::ToInt16($b, 82)
    if ($code -eq 0) {
        $fromDims = Get-PaperCodeFromDimensions -Width $wid -Length $len
        if ($fromDims) { $code = $fromDims }
    }
    $fields = [BitConverter]::ToUInt32($b, 72)
    $dup = [int][BitConverter]::ToInt16($b, 94)
    [pscustomobject]@{
        Valid          = $true
        PaperCode      = $code
        PaperName      = (Get-PaperSizeName -Code $code)
        FormName       = [Text.Encoding]::Unicode.GetString($b, 102, 64).Split([char]0)[0]
        LengthTenthMm  = $len
        WidthTenthMm   = $wid
        Fields         = $fields
        DuplexCode     = $dup
        DuplexName     = $(if ($script:DuplexNames.ContainsKey($dup)) { $script:DuplexNames[$dup] } else { '' })
        # DM_DUPLEX = 0x1000
        DuplexDeclared = (($fields -band 4096) -ne 0)
    }
}

function Initialize-PrtNlsType {
    <# .SYNOPSIS Load the GetLocaleInfoEx P/Invoke once per process. Internal. #>
    if ('PrtNls.Locale' -as [type]) { return $true }
    $sig = @(
        '[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]'
        'public static extern int GetLocaleInfoEx(string lpLocaleName, uint LCType, System.Text.StringBuilder lpLCData, int cchData);'
    ) -join [Environment]::NewLine
    try { Add-Type -Namespace PrtNls -Name Locale -MemberDefinition $sig -ErrorAction Stop } catch {}
    return [bool]('PrtNls.Locale' -as [type])
}

function Get-LocaleDefaultPaperCode {
    <#
    .SYNOPSIS
        A culture's built-in paper size (LOCALE_IPAPERSIZE, ignoring any user override). 0 if unknown.
    #>
    [CmdletBinding()] param([string]$LocaleName)
    if (-not $LocaleName) { return 0 }
    if (-not (Initialize-PrtNlsType)) { return 0 }
    $sb = New-Object System.Text.StringBuilder 16
    $n = 0
    # LOCALE_IPAPERSIZE 0x100A | LOCALE_NOUSEROVERRIDE 0x80000000
    try { $n = [PrtNls.Locale]::GetLocaleInfoEx($LocaleName, [uint32]2147487754, $sb, 16) } catch { return 0 }
    if ($n -le 0) { return 0 }
    return (ConvertTo-PaperCode $sb.ToString())
}

function Get-CountryDefaultPaperCode {
    <#
    .SYNOPSIS
        The paper size a country's cultures use by default, from its two-letter code.
    .DESCRIPTION
        Every culture of a country agrees (measured across several A4 and Letter
        countries), so the first specific culture ending in -XX answers for all of them and
        no country list is kept. Cached per process.
    .OUTPUTS
        PSCustomObject { Code; Culture }
    #>
    [CmdletBinding()] param([string]$CountryCode)
    $cc = ([string]$CountryCode).Trim().ToUpperInvariant()
    if ($cc -notmatch '^[A-Z]{2}$') { return [pscustomobject]@{ Code = 0; Culture = '' } }
    if ($script:CountryPaperCache.ContainsKey($cc)) { return $script:CountryPaperCache[$cc] }
    $hit = [pscustomobject]@{ Code = 0; Culture = '' }
    foreach ($c in [Globalization.CultureInfo]::GetCultures([Globalization.CultureTypes]::SpecificCultures)) {
        if ($c.Name -notlike "*-$cc") { continue }
        $code = Get-LocaleDefaultPaperCode -LocaleName $c.Name
        if ($code -gt 0) { $hit = [pscustomobject]@{ Code = $code; Culture = $c.Name }; break }
    }
    $script:CountryPaperCache[$cc] = $hit
    return $hit
}

function Get-RegionPaperSize {
    <#
    .SYNOPSIS
        The paper size a profile's region implies, and the one it is actually set to.
    .DESCRIPTION
        Expected: the home location (Control Panel\International\Geo -> Name, e.g.
        NL) through that country's cultures, falling back to the default of the
        profile's own LocaleName.

        Actual: iPaperSize in the same key (LOCALE_IPAPERSIZE: 1 US Letter, 5 US
        legal, 8 A3, 9 A4). Word and most applications take a new document's size
        from it, which is why a profile left on English (United States) in an
        A4 country's office sends Letter.

        '.DEFAULT' is this PC's system accounts - SYSTEM, and the region the
        machine was set up with. Read-only; needs the profile hive to be loaded.
    .OUTPUTS
        PSCustomObject { Sid; User; HiveLoaded; LocaleName; HomeCountry; UserPaperCode; UserPaper;
                         ExpectedPaperCode; ExpectedPaper; ExpectedFrom; Match; Message }
    #>
    [CmdletBinding()] param([string]$Sid = '.DEFAULT', [string]$User = '')
    if (-not $User) {
        if ($Sid -eq '.DEFAULT') { $User = 'This PC (system accounts)' }
        else {
            $User = $Sid
            try { $User = ([Security.Principal.SecurityIdentifier]$Sid).Translate([Security.Principal.NTAccount]).Value } catch {}
        }
    }
    $locale = ''; $country = ''; $userCode = 0; $loaded = $false
    $k = $null
    try { $k = [Microsoft.Win32.Registry]::Users.OpenSubKey("$Sid\Control Panel\International") } catch {}
    if ($k) {
        $loaded = $true
        try {
            $locale = [string]$k.GetValue('LocaleName')
            $userCode = ConvertTo-PaperCode $k.GetValue('iPaperSize')
            $g = $k.OpenSubKey('Geo')
            if ($g) { try { $country = [string]$g.GetValue('Name') } finally { $g.Close() } }
        } catch {} finally { $k.Close() }
    }
    # No iPaperSize means the locale's own default is in force.
    if ($loaded -and $userCode -le 0 -and $locale) { $userCode = Get-LocaleDefaultPaperCode -LocaleName $locale }

    $expected = 0; $from = ''
    if ($country) {
        $cd = Get-CountryDefaultPaperCode -CountryCode $country
        if ($cd.Code -gt 0) { $expected = [int]$cd.Code; $from = "home location $country" }
    }
    if (-not $expected -and $locale) {
        $expected = Get-LocaleDefaultPaperCode -LocaleName $locale
        if ($expected) { $from = "region format $locale" }
    }
    $match = $null
    if ($expected -gt 0 -and $userCode -gt 0) { $match = ($expected -eq $userCode) }
    $userName = if ($userCode) { Get-PaperSizeName -Code $userCode } else { '' }
    $expName  = if ($expected) { Get-PaperSizeName -Code $expected } else { '' }
    $msg = if (-not $loaded) { "The profile of $User is not loaded, so its regional paper size cannot be read. Check again while they are signed in." }
           elseif (-not $expected) { "The expected paper size for $User could not be worked out from their region." }
           elseif ($match) { "$User uses $userName, which matches $from." }
           else { "$User uses $userName, but $from implies $expName." }
    [pscustomobject]@{
        Sid = $Sid; User = $User; HiveLoaded = $loaded; LocaleName = $locale; HomeCountry = $country
        UserPaperCode = $userCode; UserPaper = $userName
        ExpectedPaperCode = $expected; ExpectedPaper = $expName; ExpectedFrom = $from
        Match = $match; Message = $msg
    }
}

function Get-PaperMatchState {
    <#
    .SYNOPSIS
        Match, Mismatch, NotComparable or Unknown for one paper code against the expected one.
    .DESCRIPTION
        Only the A4 / US-size split is judged: Letter or Legal where A4 is expected,
        or A4 where Letter is expected. A3, A5, envelope and label queues are
        deliberate and are reported NotComparable, never changed.
    #>
    [CmdletBinding()] param([int]$Code, [int]$Expected)
    if ($Code -le 0 -or $Expected -le 0) { return 'Unknown' }
    if ($Code -eq $Expected) { return 'Match' }
    $usSizes = @(1, 5)
    if ($Expected -eq 9 -and $usSizes -contains $Code) { return 'Mismatch' }
    if ($usSizes -contains $Expected -and $Code -eq 9) { return 'Mismatch' }
    return 'NotComparable'
}

function Get-QueuePaperDevMode {
    <# .SYNOPSIS A local queue's stored default paper size, from HKLM. Adds Exists. #>
    [CmdletBinding()] param([string]$Name)
    $exists = $false; $dm = $null
    if ($Name -and $Name -notmatch '\\') {
        $k = $null
        try { $k = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey("$($script:PrintersRegPath)\$Name") } catch {}
        if ($k) {
            $exists = $true
            try { $dm = $k.GetValue('Default DevMode') } catch {} finally { $k.Close() }
        }
    }
    $d = ConvertFrom-PrinterDevMode -Bytes $dm
    Add-Member -InputObject $d -NotePropertyName 'Exists' -NotePropertyValue $exists -Force
    return $d
}

function Get-PrinterPaperSizeState {
    <#
    .SYNOPSIS
        Which paper size and duplex setting each queue and connection will use, against the region.
    .DESCRIPTION
        Local queues: the queue's stored default DEVMODE. On a print server these are
        the shared queues, and their default is what new connections get.

        Connections, for the profile given by -Sid: two copies the client-side print
        provider caches on this PC -
          the server's default  ...\Servers\<server>\Printers\{GuidPrinter} : Default DevMode
          this user's copy      ...\<SID>\Printers\Connections\,,server,share : DefaultDevMode

        A user's own copy of a LOCAL queue, saved from Printing Preferences, is
        HKU\<SID>\Printers\DevModePerUser : <queue name>. While it exists the queue's
        printing defaults do not apply to that user at all, so changing the queue
        default changes nothing for them.

        All registry reads, so this needs neither the user's session nor a reachable
        server. Whether the cached server copy refreshes after the server default
        changes is not assumed.

        Paper is judged against the region. Duplex has no right answer, so it is only
        compared: DuplexState is 'Differs' when a user's copy uses a different setting
        from the default beside it, 'Same' when they agree, and '' when there is no
        pair to compare or the driver does not use duplex.
    .OUTPUTS
        PSCustomObject[] { Queue; Kind; Server; Driver; Shared; ShareName; QueueDefault; QueueDefaultCode;
                           UserCopy; UserCopyCode; HasUserCopy; Expected; ExpectedCode; State; ServerState;
                           UserCopyState; QueueDuplex; QueueDuplexDeclared; UserDuplex; DuplexState;
                           IsVirtual; Note }
    #>
    [CmdletBinding()]
    param(
        [string]$Sid = '',
        [int]$MachineExpectedCode = 0,
        [int]$UserExpectedCode = 0,
        [string[]]$VirtualNames = @(),
        [switch]$SkipLocal
    )
    if (-not $MachineExpectedCode) { try { $MachineExpectedCode = [int](Get-RegionPaperSize -Sid '.DEFAULT').ExpectedPaperCode } catch {} }
    if ($Sid -and -not $UserExpectedCode) { try { $UserExpectedCode = [int](Get-RegionPaperSize -Sid $Sid).ExpectedPaperCode } catch {} }
    if (-not $UserExpectedCode) { $UserExpectedCode = $MachineExpectedCode }
    $rows = [System.Collections.Generic.List[object]]::new()

    # One row from the default and the user's copy beside it. $QExp judges the
    # default, $UExp the user's copy: a local queue's default follows this PC's
    # region, anything the user owns follows theirs.
    $makeRow = {
        param($Info, $Q, $U, [int]$QExp, [int]$UExp)
        $qState = if ($Q.Valid) { Get-PaperMatchState -Code $Q.PaperCode -Expected $QExp } else { 'Unknown' }
        $uState = if ($U.Valid) { Get-PaperMatchState -Code $U.PaperCode -Expected $UExp } else { 'Unknown' }
        $state = if ($qState -eq 'Mismatch' -or $uState -eq 'Mismatch') { 'Mismatch' }
                 elseif ($uState -eq 'Match' -or ($uState -eq 'Unknown' -and $qState -eq 'Match')) { 'Match' }
                 elseif ($uState -eq 'NotComparable' -or $qState -eq 'NotComparable') { 'NotComparable' }
                 else { 'Unknown' }
        $qDup = if ($Q.Valid -and $Q.DuplexDeclared) { [string]$Q.DuplexName } else { '' }
        $uDup = if ($U.Valid -and $U.DuplexDeclared) { [string]$U.DuplexName } else { '' }
        $dState = if ($qDup -and $uDup) { if ($qDup -eq $uDup) { 'Same' } else { 'Differs' } } else { '' }
        $isConn = ($Info.Kind -eq 'Connection')
        $qWord = if ($isConn) { "the print server's default" } else { 'the queue default' }

        $notes = [System.Collections.Generic.List[string]]::new()
        if ($qState -eq 'Mismatch') {
            if ($isConn) { $notes.Add("The print server's default is $($Q.PaperName), as cached on this PC. Change it on $($Info.Server).") }
            else { $notes.Add("Defaults to $($Q.PaperName); this PC's region expects $(Get-PaperSizeName -Code $QExp).") }
        }
        if ($uState -eq 'Mismatch') {
            $notes.Add("This user's own copy uses $($U.PaperName) while $qWord is $(if ($Q.Valid) { $Q.PaperName } else { 'not cached here' }). Change Printing Preferences as the user.")
        }
        if ($dState -eq 'Differs') {
            $notes.Add("This user's own copy is $(Get-DuplexLabel $uDup) while $qWord is $(Get-DuplexLabel $qDup); their copy replaces the default for them.")
        }
        if (-not $notes.Count) {
            if (-not $Q.Valid -and -not $U.Valid) {
                $notes.Add($(if ($isConn) { 'No copy of this printer''s settings is cached on this PC yet.' } else { 'No stored default was found for this queue.' }))
            } elseif ($state -eq 'NotComparable') {
                $notes.Add("$(if ($U.Valid) { $U.PaperName } else { $Q.PaperName }) is left alone: only A4, Letter and Legal are compared.")
            } elseif ($U.Valid -and -not $isConn) {
                $notes.Add('This user has their own saved preferences for this printer, which replace the queue default for them.')
            }
        }
        [pscustomobject]@{
            Queue = [string]$Info.Queue; Kind = [string]$Info.Kind; Server = [string]$Info.Server; Driver = [string]$Info.Driver
            Shared = [bool]$Info.Shared; ShareName = [string]$Info.ShareName
            QueueDefault = $Q.PaperName; QueueDefaultCode = $Q.PaperCode
            UserCopy = $U.PaperName; UserCopyCode = $U.PaperCode; HasUserCopy = [bool]$U.Valid
            Expected = $(if ($QExp) { Get-PaperSizeName -Code $QExp } else { '' }); ExpectedCode = $QExp
            State = $state; ServerState = $qState; UserCopyState = $uState
            QueueDuplex = $qDup; QueueDuplexDeclared = [bool]($Q.Valid -and $Q.DuplexDeclared); UserDuplex = $uDup; DuplexState = $dState
            IsVirtual = [bool]$Info.IsVirtual; Note = ($notes -join ' ')
        }
    }

    if (-not $SkipLocal) {
        $perUser = $null
        if ($Sid) { try { $perUser = [Microsoft.Win32.Registry]::Users.OpenSubKey("$Sid\Printers\DevModePerUser") } catch {} }
        $pk = $null
        try { $pk = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($script:PrintersRegPath) } catch {}
        if ($pk) {
            try {
                foreach ($qn in @($pk.GetSubKeyNames())) {
                    $qk = $null; $dm = $null; $drv = ''; $share = ''; $shared = $false
                    try {
                        $qk = $pk.OpenSubKey($qn)
                        if ($qk) {
                            $dm = $qk.GetValue('Default DevMode')
                            $drv = [string]$qk.GetValue('Printer Driver')
                            $share = [string]$qk.GetValue('Share Name')
                            $attr = $qk.GetValue('Attributes')
                            # PRINTER_ATTRIBUTE_SHARED = 0x8
                            if ($null -ne $attr) { $shared = ((([int64]$attr) -band 8) -ne 0) }
                        }
                    } catch {} finally { if ($qk) { $qk.Close() } }
                    $ub = $null
                    if ($perUser) { try { $ub = $perUser.GetValue($qn) } catch {} }
                    $info = [pscustomobject]@{ Queue = $qn; Kind = 'Local'; Server = ''; Driver = $drv; Shared = $shared; ShareName = $share; IsVirtual = (@($VirtualNames) -contains $qn) }
                    $rows.Add((& $makeRow $info (ConvertFrom-PrinterDevMode -Bytes $dm) (ConvertFrom-PrinterDevMode -Bytes $ub) $MachineExpectedCode $UserExpectedCode))
                }
            } finally { $pk.Close() }
        }
        if ($perUser) { $perUser.Close() }
    }

    if ($Sid) {
        $map = $null
        try { $map = Get-UserPrinterMap -Sid $Sid } catch {}
        $conns = if ($map) { @($map.Connections) } else { @() }
        foreach ($c in $conns) {
            if (-not $c -or -not $c.Server) { continue }
            $usrBytes = $null; $srvBytes = $null
            $keyName = ([string]$c.Unc).Replace('\', ',')
            try {
                $uk = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey("$($script:CsrRegPath)\$Sid\Printers\Connections\$keyName")
                if ($uk) { try { $usrBytes = $uk.GetValue('DefaultDevMode') } finally { $uk.Close() } }
            } catch {}
            $guid = [string]$c.GuidPrinter
            if ($guid -and $guid -notmatch '^\{') { $guid = '{' + $guid + '}' }
            if ($guid) {
                try {
                    $sk = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey("$($script:CsrRegPath)\Servers\$($c.Server)\Printers\$guid")
                    if ($sk) { try { $srvBytes = $sk.GetValue('Default DevMode') } finally { $sk.Close() } }
                } catch {}
            }
            $info = [pscustomobject]@{ Queue = [string]$c.Unc; Kind = 'Connection'; Server = [string]$c.Server; Driver = ''; Shared = $false; ShareName = [string]$c.Printer; IsVirtual = $false }
            $rows.Add((& $makeRow $info (ConvertFrom-PrinterDevMode -Bytes $srvBytes) (ConvertFrom-PrinterDevMode -Bytes $usrBytes) $UserExpectedCode $UserExpectedCode))
        }
    }
    return $rows.ToArray()
}

function Save-PrinterPaperSnapshot {
    <#
    .SYNOPSIS
        Record the paper size about to be changed, so Restore can put exactly it back.
    .DESCRIPTION
        Kept apart from the print RPC snapshots (rpc-*.json), whose allow-list is not
        widened by this. Holds no secrets.
    .OUTPUTS
        PSCustomObject { Path; Created; Computer; Kind; Queue; Sid; User; Previous; PreviousPresent }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('QueuePaper', 'QueueDuplex', 'UserPaper')][string]$Kind,
        [string]$Queue = '',
        [string]$Sid = '',
        [string]$User = '',
        [string]$Previous = '',
        [bool]$PreviousPresent = $true
    )
    $snap = [pscustomobject]@{
        Path = ''; Created = (Get-Date).ToString('o'); Computer = $env:COMPUTERNAME; Kind = $Kind
        Queue = $Queue; Sid = $Sid; User = $User; Previous = $Previous; PreviousPresent = $PreviousPresent
    }
    try {
        if (-not (Test-Path -LiteralPath $script:PaperSnapshotDir)) { New-Item -ItemType Directory -Path $script:PaperSnapshotDir -Force | Out-Null }
        $file = Join-Path $script:PaperSnapshotDir ('paper-{0}-{1}.json' -f (Get-Date -Format 'yyyyMMdd-HHmmss-fff'), $Kind)
        $snap.Path = $file
        ($snap | ConvertTo-Json -Depth 3) | Set-Content -LiteralPath $file -Encoding UTF8
    } catch {
        $snap.Path = ''
    }
    return $snap
}

function Get-PrinterPaperSnapshot {
    <# .SYNOPSIS The saved paper-size snapshots on this PC, newest first, each with a readable label. #>
    [CmdletBinding()] param()
    if (-not (Test-Path -LiteralPath $script:PaperSnapshotDir)) { return @() }
    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($fi in @(Get-ChildItem -LiteralPath $script:PaperSnapshotDir -Filter 'paper-*.json' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)) {
        $o = $null
        try { $o = (Get-Content -LiteralPath $fi.FullName -Raw -ErrorAction Stop) | ConvertFrom-Json } catch { continue }
        if (-not $o) { continue }
        $label = if ([string]$o.Kind -eq 'QueuePaper') { "Queue '$($o.Queue)' default was $($o.Previous)" }
                 elseif ([string]$o.Kind -eq 'QueueDuplex') { "Queue '$($o.Queue)' duplex was $(Get-DuplexLabel ([string]$o.Previous))" }
                 elseif ([string]$o.Kind -eq 'UserPaper') { "$($o.User) regional paper size was $(if ($o.PreviousPresent) { Get-PaperSizeName -Code (ConvertTo-PaperCode $o.Previous) } else { 'not set' })" }
                 else { "Unrecognised snapshot '$($o.Kind)'" }
        $out.Add([pscustomobject]@{ Path = $fi.FullName; Created = [string]$o.Created; Kind = [string]$o.Kind; Computer = [string]$o.Computer; Label = $label })
    }
    return $out.ToArray()
}

function Restore-PrinterPaperSnapshot {
    <#
    .SYNOPSIS
        Put back exactly the paper size a snapshot recorded.
    .DESCRIPTION
        Refuses a snapshot from another computer, a queue that is no longer a local
        queue here, a SID that is not a user account, and any value outside the few
        this toolkit ever writes - so a hand-edited file cannot write anything else.
    .OUTPUTS
        PSCustomObject { Success; Path; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$Path, [scriptblock]$OnPoll)
    $done = { param([bool]$Ok, [string]$Msg) [pscustomobject]@{ Success = $Ok; Path = $Path; Message = $Msg } }
    $o = $null
    try { $o = (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop) | ConvertFrom-Json } catch { return (& $done $false "The snapshot could not be read: $($_.Exception.Message)") }
    if (-not $o) { return (& $done $false 'The snapshot is empty.') }
    if ([string]$o.Computer -ne $env:COMPUTERNAME) { return (& $done $false "This snapshot was taken on '$($o.Computer)', not on $env:COMPUTERNAME - refused.") }
    $kind = [string]$o.Kind
    if ($kind -eq 'QueuePaper') {
        $prev = [string]$o.Previous
        if (@('A4', 'Letter', 'Legal') -notcontains $prev) { return (& $done $false "The recorded size '$prev' is not one this toolkit changes - refused.") }
        $q = Get-QueuePaperDevMode -Name ([string]$o.Queue)
        if (-not $q.Exists) { return (& $done $false "There is no local queue named '$($o.Queue)' on this PC any more - refused.") }
        if (-not $PSCmdlet.ShouldProcess([string]$o.Queue, "Restore default paper size $prev")) { return (& $done $false 'Skipped (WhatIf).') }
        $r = Set-PrinterQueuePaperSize -Name ([string]$o.Queue) -PaperSize $prev -NoSnapshot
        return (& $done ([bool]$r.Success) ('Restore: ' + $r.Message))
    }
    if ($kind -eq 'QueueDuplex') {
        $prev = [string]$o.Previous
        if (@($script:DuplexNames.Values) -notcontains $prev) { return (& $done $false "The recorded duplex setting '$prev' is not one this toolkit changes - refused.") }
        $q = Get-QueuePaperDevMode -Name ([string]$o.Queue)
        if (-not $q.Exists) { return (& $done $false "There is no local queue named '$($o.Queue)' on this PC any more - refused.") }
        if (-not $PSCmdlet.ShouldProcess([string]$o.Queue, "Restore default duplex $prev")) { return (& $done $false 'Skipped (WhatIf).') }
        $r = Set-PrinterQueueDuplex -Name ([string]$o.Queue) -Mode $prev -NoSnapshot
        return (& $done ([bool]$r.Success) ('Restore: ' + $r.Message))
    }
    if ($kind -eq 'UserPaper') {
        $sid = [string]$o.Sid
        if ($sid -notmatch '^S-1-5-21-[\d-]+$') { return (& $done $false "'$sid' is not a user account SID - refused.") }
        if ($o.PreviousPresent) {
            $code = ConvertTo-PaperCode $o.Previous
            if (@(1, 5, 8, 9) -notcontains $code) { return (& $done $false "The recorded paper size '$($o.Previous)' is not one this toolkit writes - refused.") }
            if (-not $PSCmdlet.ShouldProcess([string]$o.User, "Restore regional paper size $code")) { return (& $done $false 'Skipped (WhatIf).') }
            $r = Set-UserRegionPaperSize -Sid $sid -User ([string]$o.User) -PaperCode $code -NoSnapshot -OnPoll $OnPoll
        } else {
            if (-not $PSCmdlet.ShouldProcess([string]$o.User, 'Remove regional paper size')) { return (& $done $false 'Skipped (WhatIf).') }
            $r = Set-UserRegionPaperSize -Sid $sid -User ([string]$o.User) -ClearValue -NoSnapshot -OnPoll $OnPoll
        }
        return (& $done ([bool]$r.Success) ('Restore: ' + $r.Message))
    }
    return (& $done $false "Unrecognised snapshot kind '$kind' - refused.")
}

function Set-PrinterQueuePaperSize {
    <#
    .SYNOPSIS
        Set a LOCAL queue's default paper size, and prove it took.
    .DESCRIPTION
        Uses Set-PrintConfiguration, which Microsoft documents as needing
        administrator credentials. Refuses a connection (\\server\share - change it
        on the server), a queue that does not exist here, and a queue whose current
        default is not A4, Letter or Legal. The previous size is snapshotted first.

        Success is decided by re-reading the queue's stored default DEVMODE, not by
        the cmdlet returning. On a print server this is the shared queue's default
        for new connections; a user who already has their own copy may keep it.
    .OUTPUTS
        PSCustomObject { Success; Queue; Previous; Now; SnapshotPath; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][ValidateSet('A4', 'Letter', 'Legal')][string]$PaperSize,
        [switch]$NoSnapshot,
        [int]$TimeoutSec = 60
    )
    $out = {
        param([bool]$Ok, [string]$Msg, [string]$Prev = '', [string]$Now = '', [string]$Snap = '')
        [pscustomobject]@{ Success = $Ok; Queue = $Name; Previous = $Prev; Now = $Now; SnapshotPath = $Snap; Message = $Msg }
    }
    if ($Name -match '^\\\\') { return (& $out $false "'$Name' is a connection to a printer on another computer. Change its default on that print server.") }
    if (-not (Test-PrintToolkitElevated)) { return (& $out $false 'Changing a queue default needs elevation (Administrator or SYSTEM).') }
    $before = Get-QueuePaperDevMode -Name $Name
    if (-not $before.Exists) { return (& $out $false "There is no local queue named '$Name' on this PC.") }
    if (-not $before.Valid) { return (& $out $false "'$Name' has no stored default to change.") }
    $wantCode = @{ Letter = 1; Legal = 5; A4 = 9 }[$PaperSize]
    if ($before.PaperCode -eq $wantCode) { return (& $out $true "'$Name' already defaults to $PaperSize." $before.PaperName $before.PaperName) }
    if (@(1, 5, 9) -notcontains $before.PaperCode) { return (& $out $false "'$Name' defaults to $($before.PaperName), which is left alone - only A4, Letter and Legal defaults are changed." $before.PaperName $before.PaperName) }
    if (-not $PSCmdlet.ShouldProcess($Name, "Set default paper size to $PaperSize")) { return (& $out $false 'Skipped (WhatIf).') }

    $snapPath = ''
    if (-not $NoSnapshot) {
        $snap = Save-PrinterPaperSnapshot -Kind 'QueuePaper' -Queue $Name -Previous $before.PaperName
        if (-not $snap.Path) { return (& $out $false 'The previous paper size could not be saved, so nothing was changed.' $before.PaperName $before.PaperName) }
        $snapPath = $snap.Path
    }
    $res = Invoke-WithTimeout -TimeoutSec $TimeoutSec -Parameters @{ QueueName = $Name; Size = $PaperSize } -ScriptBlock {
        param([string]$QueueName, [string]$Size)
        try { Set-PrintConfiguration -PrinterName $QueueName -PaperSize $Size -ErrorAction Stop; 'ok' } catch { 'error: ' + $_.Exception.Message }
    }
    $after = Get-QueuePaperDevMode -Name $Name
    if ($after.Valid -and $after.PaperCode -eq $wantCode) {
        return (& $out $true "'$Name' now defaults to $PaperSize (was $($before.PaperName)), confirmed from its stored default." $before.PaperName $after.PaperName $snapPath)
    }
    $why = if ($null -eq $res) { "Set-PrintConfiguration did not finish within $TimeoutSec s" }
           elseif ("$res" -like 'error:*') { "$res".Substring(6).Trim() }
           else { 'the stored default did not change' }
    return (& $out $false "'$Name' still defaults to $($after.PaperName): $why." $before.PaperName $after.PaperName $snapPath)
}

function Set-PrinterQueueDuplex {
    <#
    .SYNOPSIS
        Set a LOCAL queue's default duplex setting, and prove it took.
    .DESCRIPTION
        The same shape as Set-PrinterQueuePaperSize. It uses Set-PrintConfiguration,
        which Microsoft documents as needing administrator credentials, and refuses a
        connection (\\server\share - change it on the server) and a queue that does
        not exist here. The previous setting is snapshotted first, and success is
        decided by re-reading the queue's stored default DEVMODE.

        It also refuses a queue whose stored default does not declare duplex
        (DM_DUPLEX clear): its driver does not use the setting, as on a label
        printer or a PDF writer.

        A user with their own saved preferences for the queue keeps them - their
        copy replaces the queue default for them (MAP021).

        Duplex names are converted by name, never by number: the DEVMODE counts
        1-3 and PrintManagement's DuplexingModeEnum counts 0-2 (measured).
    .OUTPUTS
        PSCustomObject { Success; Queue; Previous; Now; SnapshotPath; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][ValidateSet('OneSided', 'TwoSidedLongEdge', 'TwoSidedShortEdge')][string]$Mode,
        [switch]$NoSnapshot,
        [int]$TimeoutSec = 60
    )
    $out = {
        param([bool]$Ok, [string]$Msg, [string]$Prev = '', [string]$Now = '', [string]$Snap = '')
        [pscustomobject]@{ Success = $Ok; Queue = $Name; Previous = $Prev; Now = $Now; SnapshotPath = $Snap; Message = $Msg }
    }
    $label = Get-DuplexLabel $Mode
    if ($Name -match '^\\\\') { return (& $out $false "'$Name' is a connection to a printer on another computer. Change its default on that print server.") }
    if (-not (Test-PrintToolkitElevated)) { return (& $out $false 'Changing a queue default needs elevation (Administrator or SYSTEM).') }
    $before = Get-QueuePaperDevMode -Name $Name
    if (-not $before.Exists) { return (& $out $false "There is no local queue named '$Name' on this PC.") }
    if (-not $before.Valid) { return (& $out $false "'$Name' has no stored default to change.") }
    if (-not $before.DuplexDeclared) { return (& $out $false "The driver for '$Name' does not use a duplex setting, so it is left alone.") }
    if (-not $before.DuplexName) { return (& $out $false "'$Name' has an unrecognised duplex value ($($before.DuplexCode)), so it is left alone.") }
    $want = Get-DuplexCode -Name $Mode
    if ($before.DuplexCode -eq $want) { return (& $out $true "'$Name' is already $label by default." $before.DuplexName $before.DuplexName) }
    if (-not $PSCmdlet.ShouldProcess($Name, "Set default duplex to $Mode")) { return (& $out $false 'Skipped (WhatIf).') }

    $snapPath = ''
    if (-not $NoSnapshot) {
        $snap = Save-PrinterPaperSnapshot -Kind 'QueueDuplex' -Queue $Name -Previous $before.DuplexName
        if (-not $snap.Path) { return (& $out $false 'The previous duplex setting could not be saved, so nothing was changed.' $before.DuplexName $before.DuplexName) }
        $snapPath = $snap.Path
    }
    $res = Invoke-WithTimeout -TimeoutSec $TimeoutSec -Parameters @{ QueueName = $Name; Duplex = $Mode } -ScriptBlock {
        param([string]$QueueName, [string]$Duplex)
        try { Set-PrintConfiguration -PrinterName $QueueName -DuplexingMode $Duplex -ErrorAction Stop; 'ok' } catch { 'error: ' + $_.Exception.Message }
    }
    $after = Get-QueuePaperDevMode -Name $Name
    if ($after.Valid -and $after.DuplexCode -eq $want) {
        return (& $out $true "'$Name' is now $label by default (was $(Get-DuplexLabel $before.DuplexName)), confirmed from its stored default. Anyone with their own saved preferences for it keeps theirs." $before.DuplexName $after.DuplexName $snapPath)
    }
    $why = if ($null -eq $res) { "Set-PrintConfiguration did not finish within $TimeoutSec s" }
           elseif ("$res" -like 'error:*') { "$res".Substring(6).Trim() }
           else { 'the stored default did not change' }
    return (& $out $false "'$Name' is still $(Get-DuplexLabel $after.DuplexName) by default: $why." $before.DuplexName $after.DuplexName $snapPath)
}

function Set-UserRegionPaperSize {
    <#
    .SYNOPSIS
        Set one user's regional paper size (LOCALE_IPAPERSIZE), and prove it took.
    .DESCRIPTION
        Changes that single setting - never the user's region format, language or
        home location. In the user's own session it calls SetLocaleInfoW and then
        broadcasts WM_SETTINGCHANGE "intl", as Region settings does; programs
        already open pick it up when restarted. If that did not take and the
        profile hive is loaded, iPaperSize is written directly, which applies to
        programs started after the user next signs in.

        -ClearValue (used by restore) removes iPaperSize so the locale's own default
        applies again. Success is decided by reading iPaperSize back.
    .OUTPUTS
        PSCustomObject { Success; User; Previous; Now; Method; SnapshotPath; Message }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Sid,
        [string]$User = '',
        [ValidateSet(1, 5, 8, 9)][int]$PaperCode = 9,
        [switch]$ClearValue,
        [switch]$NoSnapshot,
        [scriptblock]$OnPoll
    )
    $out = {
        param([bool]$Ok, [string]$Msg, [string]$Prev = '', [string]$Now = '', [string]$How = '', [string]$Snap = '')
        [pscustomobject]@{ Success = $Ok; User = $User; Previous = $Prev; Now = $Now; Method = $How; SnapshotPath = $Snap; Message = $Msg }
    }
    if ($Sid -notmatch '^S-1-5-21-[\d-]+$') { return (& $out $false "'$Sid' is not a user account SID. Only a user's own regional paper size is changed.") }
    if (-not $User) {
        $User = $Sid
        try { $User = ([Security.Principal.SecurityIdentifier]$Sid).Translate([Security.Principal.NTAccount]).Value } catch {}
    }
    $read = {
        $present = $false; $val = ''; $isLoaded = $false
        $rk = $null
        try { $rk = [Microsoft.Win32.Registry]::Users.OpenSubKey("$Sid\Control Panel\International") } catch {}
        if ($rk) {
            $isLoaded = $true
            try { $raw = $rk.GetValue('iPaperSize'); if ($null -ne $raw) { $present = $true; $val = [string]$raw } } catch {} finally { $rk.Close() }
        }
        [pscustomobject]@{ Loaded = $isLoaded; Present = $present; Value = $val }
    }
    $nameOf = { param($State) if ($State.Present) { Get-PaperSizeName -Code (ConvertTo-PaperCode $State.Value) } else { 'not set' } }

    $before = & $read
    if (-not $before.Loaded) { return (& $out $false "$User is not signed in (their profile is not loaded), so their regional paper size cannot be read or changed.") }
    $target = [string]$PaperCode
    if (-not $ClearValue -and $before.Present -and $before.Value -eq $target) {
        return (& $out $true "$User already uses $(Get-PaperSizeName -Code $PaperCode)." $before.Value $before.Value 'none')
    }
    if ($ClearValue -and -not $before.Present) { return (& $out $true "$User has no regional paper size set, so their region's default already applies." '' '' 'none') }
    $what = if ($ClearValue) { 'Remove the regional paper size' } else { "Set regional paper size to $(Get-PaperSizeName -Code $PaperCode)" }
    if (-not $PSCmdlet.ShouldProcess($User, $what)) { return (& $out $false 'Skipped (WhatIf).') }

    $snapPath = ''
    if (-not $NoSnapshot) {
        $snap = Save-PrinterPaperSnapshot -Kind 'UserPaper' -Sid $Sid -User $User -Previous $before.Value -PreviousPresent $before.Present
        if (-not $snap.Path) { return (& $out $false 'The previous setting could not be saved, so nothing was changed.' $before.Value $before.Value) }
        $snapPath = $snap.Path
    }

    $how = ''
    $sessionMsg = ''
    if ($ClearValue) {
        try {
            $wk = [Microsoft.Win32.Registry]::Users.OpenSubKey("$Sid\Control Panel\International", $true)
            if ($wk) { try { $wk.DeleteValue('iPaperSize', $false); $how = 'registry' } finally { $wk.Close() } }
        } catch {}
        $after = & $read
        if (-not $after.Present) { return (& $out $true "$User's regional paper size was removed, so their region's default applies again." $before.Value '' $how $snapPath) }
        return (& $out $false "$User's regional paper size could not be removed; it is still $(& $nameOf $after)." $before.Value $after.Value $how $snapPath)
    }

    $r = Invoke-ScriptInPrinterUserSession -User $User -OnPoll $OnPoll -Vars @{ Code = $target } -ScriptBlock {
        $ok = $false; $err = ''
        try {
            $sig = @(
                '[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]'
                'public static extern bool SetLocaleInfoW(uint Locale, uint LCType, string lpLCData);'
                '[DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]'
                'public static extern System.IntPtr SendMessageTimeoutW(System.IntPtr hWnd, uint Msg, System.UIntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out System.UIntPtr lpdwResult);'
            ) -join [Environment]::NewLine
            if (-not ('PrtNls.Setter' -as [type])) { Add-Type -Namespace PrtNls -Name Setter -MemberDefinition $sig -ErrorAction Stop }
            # LOCALE_USER_DEFAULT 0x400, LOCALE_IPAPERSIZE 0x100A
            $ok = [PrtNls.Setter]::SetLocaleInfoW([uint32]1024, [uint32]4106, $Code)
            if ($ok) {
                $res = [UIntPtr]::Zero
                # HWND_BROADCAST 0xFFFF, WM_SETTINGCHANGE 0x1A, SMTO_ABORTIFHUNG 0x2, 3 s
                [void][PrtNls.Setter]::SendMessageTimeoutW([IntPtr]65535, [uint32]26, [UIntPtr]::Zero, 'intl', [uint32]2, [uint32]3000, [ref]$res)
            } else {
                $err = 'SetLocaleInfoW failed with ' + [Runtime.InteropServices.Marshal]::GetLastWin32Error()
            }
        } catch { $err = $_.Exception.Message }
        [pscustomobject]@{ Set = [bool]$ok; Error = $err } | ConvertTo-Json -Compress
    }
    if ($r.Success -and $r.Data.Set) { $how = 'in-session' }
    elseif ($r.Data -and $r.Data.Error) { $sessionMsg = [string]$r.Data.Error }
    else { $sessionMsg = [string]$r.Message }

    $after = & $read
    if (-not ($after.Present -and $after.Value -eq $target)) {
        try {
            $wk = [Microsoft.Win32.Registry]::Users.OpenSubKey("$Sid\Control Panel\International", $true)
            if ($wk) { try { $wk.SetValue('iPaperSize', $target, [Microsoft.Win32.RegistryValueKind]::String); $how = 'registry' } finally { $wk.Close() } }
        } catch {}
        $after = & $read
    }
    $was = & $nameOf $before
    if ($after.Present -and $after.Value -eq $target) {
        $tail = if ($how -eq 'registry') { ' It was written to their profile directly, so programs started after they next sign in use it.' } else { ' Programs already open pick it up when they are restarted.' }
        return (& $out $true "$User's regional paper size is now $(Get-PaperSizeName -Code $PaperCode) (was $was), confirmed from their profile.$tail" $before.Value $after.Value $how $snapPath)
    }
    return (& $out $false "$User's regional paper size is still $(& $nameOf $after). $sessionMsg" $before.Value $after.Value $how $snapPath)
}

# endregion
'@
#endregion TOOLKIT

. ([scriptblock]::Create($script:ToolkitSource))


# --------------------------------------------------------------------------
# Context
# --------------------------------------------------------------------------
$script:Ctx        = Get-PrinterContext
$script:IsSystem   = $script:Ctx.IsSystem
$script:TargetUser = $null
$script:TargetSid  = $null
$script:Printers   = @()
$script:UserMap    = $null
$script:Findings   = @()
$script:ServerRows = @()
$script:Jobs       = @()
$script:LogIndex   = 0
$script:Busy       = $false
# Hard ceiling on any single background operation. Generous enough for a slow
# domain with several unreachable servers, short enough that a tech is never
# left staring at a frozen console.
$script:WorkerTimeoutSec = 180
$script:WorkerStarted    = Get-Date


# --------------------------------------------------------------------------
# Form scaffold
# --------------------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = "Printer Manager  -  $($env:COMPUTERNAME)"
$form.Size          = New-Object System.Drawing.Size(1120, 840)
# Low enough to fit a modest ScreenConnect viewport; docking keeps every band
# usable at this size.
$form.MinimumSize   = New-Object System.Drawing.Size(820, 560)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)


function New-Btn {
    param([string]$Text, [int]$X, [int]$Y, [int]$W = 120, [int]$H = 27, $Parent)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size($W, $H)
    if ($Parent) { $Parent.Controls.Add($b) }
    $b
}
function New-Lbl {
    param([string]$Text, [int]$X, [int]$Y, $Parent, [bool]$Bold = $false)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    # AutoSize so a longer string at higher DPI can never be clipped or paint
    # over the control beside it.
    $l.AutoSize = $true
    if ($Bold) { $l.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold) }
    if ($Parent) { $Parent.Controls.Add($l) }
    $l
}
function Set-DoubleBuffered {
    <#
    .SYNOPSIS
        Turn on double buffering for a control that does not expose it publicly.
    .DESCRIPTION
        DataGridView repaints row by row and erases its background first, so any
        refresh flickers - badly over a remote session, where every intermediate
        frame is actually transmitted. The DoubleBuffered property is protected,
        hence the reflection; this is the standard remedy.

        Best-effort by design: if reflection is blocked the grid still works, it
        just flickers as before.
    #>
    param($Control)
    try {
        $p = $Control.GetType().GetProperty('DoubleBuffered',
                [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
        if ($p) { $p.SetValue($Control, $true, $null) }
    } catch {}
}
function New-Grid {
    param([int]$X, [int]$Y, [int]$W, [int]$H, $Parent)
    $g = New-Object System.Windows.Forms.DataGridView
    Set-DoubleBuffered $g
    $g.Location = New-Object System.Drawing.Point($X, $Y)
    $g.Size = New-Object System.Drawing.Size($W, $H)
    $g.Anchor = [System.Windows.Forms.AnchorStyles]'Top,Left,Right,Bottom'
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.AllowUserToResizeRows = $false
    $g.ReadOnly = $true
    $g.RowHeadersVisible = $false
    $g.SelectionMode = 'FullRowSelect'
    $g.MultiSelect = $false
    # Fill, not None. With fixed pixel widths the Printers grid alone totalled
    # ~1260px, so on a smaller remote screen (ScreenConnect) the right-hand
    # columns ran off and needed horizontal scrolling. Fill shares the ACTUAL
    # width between columns in proportion to their FillWeight, so every column
    # is always on screen at any window size.
    $g.AutoSizeColumnsMode = 'Fill'
    # Truncated cells then show their full value on hover.
    $g.ShowCellToolTips = $true
    $g.BackgroundColor = [System.Drawing.Color]::White
    if ($Parent) { $Parent.Controls.Add($g) }
    $g
}
function Add-Col {
    param($Grid, [string]$Name, [string]$Header, [int]$Width)
    $i = $Grid.Columns.Add($Name, $Header)
    # In Fill mode Width is ignored and FillWeight sets the proportion, so the
    # designed pixel width becomes the weight - the columns keep their relative
    # sizing but always add up to the width actually available.
    $Grid.Columns[$i].FillWeight = $Width
    $Grid.Columns[$i].MinimumWidth = [Math]::Min(50, $Width)
    $Grid.Columns[$i].DefaultCellStyle.WrapMode = 'False'
    $i
}
function Set-Rows {
    <# Repopulate a grid from objects, mapping property names to columns in order. #>
    param($Grid, $Data, [string[]]$Props)
    $Grid.SuspendLayout()
    $Grid.Rows.Clear()
    foreach ($d in @($Data)) {
        if ($null -eq $d) { continue }
        $vals = New-Object object[] $Props.Count
        for ($i = 0; $i -lt $Props.Count; $i++) {
            $v = $d.($Props[$i])
            $vals[$i] = if ($null -eq $v) { '' } else { "$v" }
        }
        [void]$Grid.Rows.Add($vals)
    }
    $Grid.ResumeLayout()
}


# --------------------------------------------------------------------------
# Banner + target user
# --------------------------------------------------------------------------
# The window is three docked bands - header on top, log + status on the bottom,
# tabs filling the middle. Absolute positions with Left,Right,Bottom anchors put
# the log pane and status bar off the bottom/right edge on a smaller remote
# screen; docking is recomputed by the form at whatever size it ends up.
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = 'Top'
$pnlHeader.Height = 92
$form.Controls.Add($pnlHeader)

$pnlFooter = New-Object System.Windows.Forms.Panel
$pnlFooter.Dock = 'Bottom'
$pnlFooter.Height = 128
$form.Controls.Add($pnlFooter)

$lblBanner = New-Object System.Windows.Forms.Label
$lblBanner.Location = New-Object System.Drawing.Point(12, 8)
$lblBanner.Size = New-Object System.Drawing.Size(1080, 20)
$lblBanner.Anchor = [System.Windows.Forms.AnchorStyles]'Top,Left,Right'
$lblBanner.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$ctxTag = if ($script:IsSystem) { 'SYSTEM' } elseif ($script:Ctx.IsAdmin) { 'elevated user' } else { 'standard user' }
$iuText = if (@($script:Ctx.InteractiveUsers).Count) { (@($script:Ctx.InteractiveUsers) | ForEach-Object { $_.User }) -join ', ' } else { '(nobody signed in)' }
$lblBanner.Text = "Context: $($script:Ctx.CurrentName)  [$ctxTag]      Signed in: $iuText      Host: $($script:Ctx.Host)      Spooler: $($script:Ctx.Spooler.Status)      Toolkit v$($script:Ctx.ToolkitVersion)"
if ($script:IsSystem -and @($script:Ctx.InteractiveUsers).Count -eq 0) {
    $lblBanner.ForeColor = [System.Drawing.Color]::Firebrick
    $lblBanner.Text += '   - no user session to act into!'
} elseif ($script:IsSystem) {
    $lblBanner.ForeColor = [System.Drawing.Color]::DarkGreen
}
$pnlHeader.Controls.Add($lblBanner)

New-Lbl 'Target user:' 12 36 $pnlHeader $true | Out-Null
$cboUser = New-Object System.Windows.Forms.ComboBox
$cboUser.Location = New-Object System.Drawing.Point(92, 33)
$cboUser.Size = New-Object System.Drawing.Size(300, 24)
$cboUser.DropDownStyle = 'DropDownList'
$pnlHeader.Controls.Add($cboUser)

$lblUserNote = New-Lbl '' 400 36 $pnlHeader
$lblUserNote.ForeColor = [System.Drawing.Color]::DimGray

$btnRescan = New-Btn 'Rescan' 12 62 100 27 $pnlHeader
$btnReport = New-Btn 'Report...' 118 62 100 27 $pnlHeader


# --------------------------------------------------------------------------
# Tabs
# --------------------------------------------------------------------------
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$form.Controls.Add($tabs)
# Docking is applied highest z-order index first, and whatever is processed last
# takes the remaining area - so the Fill control must be at index 0.
$tabs.BringToFront()

$tabPrinters = New-Object System.Windows.Forms.TabPage; $tabPrinters.Text = 'Printers'
$tabDiag     = New-Object System.Windows.Forms.TabPage; $tabDiag.Text     = 'Diagnose'
$tabSpool    = New-Object System.Windows.Forms.TabPage; $tabSpool.Text    = 'Spooler and Jobs'
$tabPolicy   = New-Object System.Windows.Forms.TabPage; $tabPolicy.Text   = 'Policy'
$tabServer   = New-Object System.Windows.Forms.TabPage; $tabServer.Text   = 'Shares'
$tabHost     = New-Object System.Windows.Forms.TabPage; $tabHost.Text     = 'Share Host'
foreach ($t in @($tabPrinters, $tabDiag, $tabSpool, $tabPolicy, $tabServer, $tabHost)) {
    $t.BackColor = [System.Drawing.Color]::White
    # AutoScroll per TAB PAGE, not just on the form. Controls inside a tab are
    # placed at fixed coordinates and the widest button row reaches ~914px, so on
    # a small Backstage session (the form clamps to the working area) the
    # right-hand buttons would be clipped with no scrollbar to reach them.
    $t.AutoScroll = $true
    [void]$tabs.TabPages.Add($t)
}

# --------------------------------------------------------------------------
# Tab contents.
#
# LAYOUT RULE: every tab is built as docked bands - an optional Top panel for
# headers, a Bottom panel for footer controls, and the grid filling whatever is
# left. Docking is recomputed by the container at any size, so nothing can end
# up underneath the grid.
#
# Absolute positions with Top+Bottom anchors were tried first and are wrong for
# this: a grid anchored that way keeps its top fixed and follows the container's
# bottom edge, so it grows straight over any fixed-Y control beneath it - on a
# maximised window the whole Printers button area vanished behind the grid.
# Anchoring the buttons to the bottom instead is no good either, because a
# TabPage's children are not laid out when Add_Shown fires, so the offsets are
# computed against a zero-height container and the controls land off-screen.
# --------------------------------------------------------------------------
function New-DockBand {
    param($Parent, [string]$Side, [int]$Height)
    $pnl = New-Object System.Windows.Forms.Panel
    $pnl.Dock = $Side
    $pnl.Height = $Height
    $pnl.BackColor = [System.Drawing.Color]::White
    $Parent.Controls.Add($pnl)
    $pnl
}
function Set-FillGrid {
    param($Grid)
    $Grid.Dock = 'Fill'
    # Docking is applied from the highest z-order index down, and whatever is
    # processed last receives the remaining area. BringToFront puts the grid at
    # index 0 so the docked bands reserve their space first.
    $Grid.BringToFront()
}

# ---- Tab 1: Printers -----------------------------------------------------
$pnlPrintersBtns = New-DockBand $tabPrinters 'Bottom' 108
$dgvPrinters = New-Grid 0 0 10 10 $tabPrinters
Add-Col $dgvPrinters 'pName'    'Printer'      330 | Out-Null
Add-Col $dgvPrinters 'pType'    'Type'         120 | Out-Null
Add-Col $dgvPrinters 'pKind'    'Port kind'     80 | Out-Null
Add-Col $dgvPrinters 'pServer'  'Server'       110 | Out-Null
Add-Col $dgvPrinters 'pDriver'  'Driver'       210 | Out-Null
Add-Col $dgvPrinters 'pPort'    'Port'         130 | Out-Null
Add-Col $dgvPrinters 'pHost'    'Host/IP'      110 | Out-Null
Add-Col $dgvPrinters 'pDefault' 'Default'       60 | Out-Null
Add-Col $dgvPrinters 'pStatus'  'Status'       110 | Out-Null
Set-FillGrid $dgvPrinters

$pRow1 = 5; $pRow2 = 39; $pRow3 = 73
$btnRefreshP  = New-Btn 'Refresh'            8   $pRow1 110 27 $pnlPrintersBtns
$btnRename    = New-Btn 'Rename...'          124 $pRow1 110 27 $pnlPrintersBtns
$btnDelete    = New-Btn 'Delete...'          240 $pRow1 110 27 $pnlPrintersBtns
$btnSetDef    = New-Btn 'Set as Default'     356 $pRow1 130 27 $pnlPrintersBtns
$btnTestPort  = New-Btn 'Test Reachability'  492 $pRow1 140 27 $pnlPrintersBtns
$btnTestPage  = New-Btn 'Print Test Page'    638 $pRow1 130 27 $pnlPrintersBtns

$btnPrintMgmt = New-Btn 'Open printmanagement.msc' 8   $pRow2 190 27 $pnlPrintersBtns
$btnDevPrint  = New-Btn 'Devices and Printers'     204 $pRow2 160 27 $pnlPrintersBtns
$btnRepoint   = New-Btn 'Repoint / Rebuild...'    370 $pRow2 150 27 $pnlPrintersBtns
$lblPrinterHint = New-Lbl '' 530 ($pRow2 + 5) $pnlPrintersBtns

# Row 3. A third row rather than squeezing onto row 2: that row already reaches
# x=686 and the hint label sits immediately after it.
$btnJunk = New-Btn 'Remove junk queues...' 8 $pRow3 170 27 $pnlPrintersBtns
$btnPaper = New-Btn 'Paper and duplex...' 184 $pRow3 140 27 $pnlPrintersBtns
$lblJunk = New-Lbl '' 332 ($pRow3 + 5) $pnlPrintersBtns
$lblJunk.ForeColor = [System.Drawing.Color]::DimGray
$lblPrinterHint.ForeColor = [System.Drawing.Color]::DimGray
$lblPrinterHint.Text = 'Per-user actions apply to the Target user above.'

# ---- Tab 2: Diagnose -----------------------------------------------------
$pnlDiagTop = New-DockBand $tabDiag 'Top' 42
# 140, not 200: at the 820x560 minimum the tab is only ~277px tall, and a
# 200px band left the findings grid at ~35px. The grid grows on a big window.
$pnlDiagBot = New-DockBand $tabDiag 'Bottom' 140
$dgvFindings = New-Grid 0 0 10 10 $tabDiag
Add-Col $dgvFindings 'fCode'  'Code'      70  | Out-Null
Add-Col $dgvFindings 'fSev'   'Severity'  75  | Out-Null
Add-Col $dgvFindings 'fScope' 'Scope'     70  | Out-Null
Add-Col $dgvFindings 'fUser'  'User'      140 | Out-Null
Add-Col $dgvFindings 'fPrn'   'Printer'   300 | Out-Null
Add-Col $dgvFindings 'fDet'   'Detail'    380 | Out-Null
Set-FillGrid $dgvFindings

$btnRunDiag = New-Btn 'Run diagnostics' 8 6 130 27 $pnlDiagTop
$chkSkipNet = New-Object System.Windows.Forms.CheckBox
$chkSkipNet.Text = 'Skip network probes (faster)'
$chkSkipNet.Location = New-Object System.Drawing.Point(148, 10)
$chkSkipNet.AutoSize = $true
$pnlDiagTop.Controls.Add($chkSkipNet)
New-Lbl 'Stuck job after (h):' 350 11 $pnlDiagTop | Out-Null
$numStuck = New-Object System.Windows.Forms.NumericUpDown
$numStuck.Location = New-Object System.Drawing.Point(462, 8)
$numStuck.Size = New-Object System.Drawing.Size(60, 24)
$numStuck.Minimum = 1; $numStuck.Maximum = 9999; $numStuck.Value = 24
$pnlDiagTop.Controls.Add($numStuck)
$btnCopyDiag = New-Btn 'Copy report' 540 6 110 27 $pnlDiagTop

$txtFinding = New-Object System.Windows.Forms.TextBox
$txtFinding.Dock = 'Fill'
$txtFinding.Multiline = $true
$txtFinding.ReadOnly = $true
$txtFinding.ScrollBars = 'Vertical'
$txtFinding.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtFinding.BackColor = [System.Drawing.Color]::White
$pnlDiagBot.Controls.Add($txtFinding)

# ---- Tab 3: Spooler and Jobs --------------------------------------------
$pnlSpoolTop = New-DockBand $tabSpool 'Top' 96
$pnlSpoolBot = New-DockBand $tabSpool 'Bottom' 110
$dgvJobs = New-Grid 0 0 10 10 $tabSpool
Add-Col $dgvJobs 'jQueue' 'Queue'    300 | Out-Null
Add-Col $dgvJobs 'jId'    'Id'       50  | Out-Null
Add-Col $dgvJobs 'jUser'  'User'     120 | Out-Null
Add-Col $dgvJobs 'jDoc'   'Document' 240 | Out-Null
Add-Col $dgvJobs 'jSt'    'Status'   110 | Out-Null
Add-Col $dgvJobs 'jPages' 'Pages'    80  | Out-Null
Add-Col $dgvJobs 'jAge'   'Age (h)'  80  | Out-Null
Add-Col $dgvJobs 'jPaper' 'Paper'    90  | Out-Null
Set-FillGrid $dgvJobs

$lblSpool = New-Lbl 'Spooler: (not read yet)' 8 10 $pnlSpoolTop $true
$btnSpoolRefresh = New-Btn 'Refresh'              8   34 110 27 $pnlSpoolTop
$btnSpoolRestart = New-Btn 'Restart Spooler'      124 34 130 27 $pnlSpoolTop
$btnSpoolClear   = New-Btn 'Clear Spool Files...' 260 34 150 27 $pnlSpoolTop
$lblSpoolDir = New-Lbl '' 8 68 $pnlSpoolTop
$lblSpoolDir.ForeColor = [System.Drawing.Color]::DimGray

$btnCancelJob  = New-Btn 'Cancel selected job' 8   6 150 27 $pnlSpoolBot
$btnCancelAll  = New-Btn 'Cancel all on queue' 164 6 150 27 $pnlSpoolBot
$lblEventsHdr = New-Lbl 'Recent PrintService / spooler events:' 8 42 $pnlSpoolBot $true
$txtEvents = New-Object System.Windows.Forms.TextBox
$txtEvents.Location = New-Object System.Drawing.Point(8, 64)
$txtEvents.Size = New-Object System.Drawing.Size(1054, 40)
$txtEvents.Anchor = [System.Windows.Forms.AnchorStyles]'Top,Left,Right,Bottom'
$txtEvents.Multiline = $true; $txtEvents.ReadOnly = $true; $txtEvents.ScrollBars = 'Both'
$txtEvents.WordWrap = $false
$txtEvents.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtEvents.BackColor = [System.Drawing.Color]::White
$pnlSpoolBot.Controls.Add($txtEvents)

# ---- Tab 4: Policy -------------------------------------------------------
$pnlPolicyTop = New-DockBand $tabPolicy 'Top' 30
$pnlPolicyBot = New-DockBand $tabPolicy 'Bottom' 150
$dgvPolicy = New-Grid 0 0 10 10 $tabPolicy
Add-Col $dgvPolicy 'yName'  'Value name'    290 | Out-Null
Add-Col $dgvPolicy 'yState' 'State'         210 | Out-Null
Add-Col $dgvPolicy 'yPath'  'Registry path' 520 | Out-Null
Set-FillGrid $dgvPolicy

New-Lbl 'Point and Print / driver install hardening. An UNSET value is NOT an unenforced value - the State column shows the effective value.' 8 7 $pnlPolicyTop $true | Out-Null

$lblCodesHdr = New-Lbl 'Common printer connection error codes:' 8 6 $pnlPolicyBot $true
$txtCodes = New-Object System.Windows.Forms.TextBox
$txtCodes.Location = New-Object System.Drawing.Point(8, 28)
$txtCodes.Size = New-Object System.Drawing.Size(1054, 116)
$txtCodes.Anchor = [System.Windows.Forms.AnchorStyles]'Top,Left,Right,Bottom'
$txtCodes.Multiline = $true; $txtCodes.ReadOnly = $true; $txtCodes.ScrollBars = 'Vertical'
$txtCodes.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtCodes.BackColor = [System.Drawing.Color]::White
$pnlPolicyBot.Controls.Add($txtCodes)

# ---- Tab 5: Print Server -------------------------------------------------
$pnlServerTop = New-DockBand $tabServer 'Top' 136
$pnlServerBot = New-DockBand $tabServer 'Bottom' 40
$dgvServer = New-Grid 0 0 10 10 $tabServer
Add-Col $dgvServer 'sShare'  'Share'      220 | Out-Null
Add-Col $dgvServer 'sName'   'Queue name' 300 | Out-Null
Add-Col $dgvServer 'sDriver' 'Driver'     240 | Out-Null
Add-Col $dgvServer 'sPort'   'Port'       130 | Out-Null
Add-Col $dgvServer 'sHost'   'Host/IP'    120 | Out-Null
Add-Col $dgvServer 'sUnc'    'UNC'        300 | Out-Null
Set-FillGrid $dgvServer

# This tab is the TARGET side of a shared printer: everything on it runs on, and
# acts on, the PC that CONNECTS. Work that belongs on the PC sharing the printer
# lives on Share Host, and the banner says so because in Backstage it is easy to be
# on the wrong machine.
$lblServerRole = New-Lbl 'TARGET PC  -  run this tab on the PC that CONNECTS to the shared printer.' 8 6 $pnlServerTop $true
$lblServerRole.ForeColor = [System.Drawing.Color]::FromArgb(0, 84, 147)

$sRow1 = 30; $sRow2 = 66; $sRow3 = 102
New-Lbl 'Print server:' 8 ($sRow1 + 5) $pnlServerTop $true | Out-Null
$txtServer = New-Object System.Windows.Forms.TextBox
$txtServer.Location = New-Object System.Drawing.Point(92, ($sRow1 + 2))
$txtServer.Size = New-Object System.Drawing.Size(220, 24)
$pnlServerTop.Controls.Add($txtServer)
$btnDiscover   = New-Btn 'Discover'       320 $sRow1 100 27 $pnlServerTop
$btnEnumerate  = New-Btn 'Enumerate'      426 $sRow1 100 27 $pnlServerTop
$btnCreds      = New-Btn 'Credentials...' 532 $sRow1 120 27 $pnlServerTop
$btnPasteJson  = New-Btn 'Paste JSON...'  658 $sRow1 120 27 $pnlServerTop

New-Lbl 'Scope:' 8 ($sRow2 + 5) $pnlServerTop $true | Out-Null
$cboAddScope = New-Object System.Windows.Forms.ComboBox
$cboAddScope.Location = New-Object System.Drawing.Point(58, ($sRow2 + 2))
$cboAddScope.Size = New-Object System.Drawing.Size(290, 24)
$cboAddScope.DropDownStyle = 'DropDownList'
# ONE Add button, and it always honours this scope. There used to be two, and
# 'Add for user' silently ignored the dropdown.
[void]$cboAddScope.Items.Add('All users on this PC (per-computer)')
[void]$cboAddScope.Items.Add('Only the target user (applies immediately)')
[void]$cboAddScope.Items.Add('All users, via local port (no Point and Print)')
$cboAddScope.SelectedIndex = 0
$pnlServerTop.Controls.Add($cboAddScope)
$btnAddToPc          = New-Btn 'Add...'             356 $sRow2 90  27 $pnlServerTop
$btnSrvTest          = New-Btn 'Test connection...' 452 $sRow2 140 27 $pnlServerTop
$btnSrvRemoveMachine = New-Btn 'Remove per-PC...'   598 $sRow2 130 27 $pnlServerTop

$btnInstallDrv = New-Btn 'Install driver...'        8   $sRow3 130 27 $pnlServerTop
$btnDrvInfo    = New-Btn 'Host driver info...'      144 $sRow3 140 27 $pnlServerTop
$btnSrvRpc     = New-Btn 'Advanced: named pipes...' 290 $sRow3 170 27 $pnlServerTop
$lblSrvRpc = New-Lbl '' 468 ($sRow3 + 5) $pnlServerTop
$lblSrvRpc.ForeColor = [System.Drawing.Color]::DimGray

$lblServerNote = New-Lbl '' 8 10 $pnlServerBot
$lblServerNote.ForeColor = [System.Drawing.Color]::DimGray
$lblServerNote.MaximumSize = New-Object System.Drawing.Size(1050, 0)

# ---- Tab 6: Share Host (this PC sharing its own printers) ------------------
# A separate tab rather than a second panel on Shares: that tab already owns a
# grid, and two grids plus two button bands do not fit the ~1024x700 window a
# ScreenConnect session gives you.
$pnlHostTop = New-DockBand $tabHost 'Top' 134
$pnlHostBot = New-DockBand $tabHost 'Bottom' 120
$dgvHost = New-Grid 0 0 10 10 $tabHost
Add-Col $dgvHost 'hCode' 'Code'      70  | Out-Null
Add-Col $dgvHost 'hSev'  'Severity'  75  | Out-Null
Add-Col $dgvHost 'hPrn'  'Printer'   260 | Out-Null
Add-Col $dgvHost 'hDet'  'Finding'   620 | Out-Null
Set-FillGrid $dgvHost

# This tab is the SOURCE side of a shared printer: everything on it runs on, and
# changes, the PC that SHARES the printer. 'Share to everyone' moved here from the
# Printers tab because it only ever acts on this PC.
$lblHostRole = New-Lbl 'SOURCE PC  -  run this tab on the PC that SHARES the printer.' 8 6 $pnlHostTop $true
$lblHostRole.ForeColor = [System.Drawing.Color]::FromArgb(0, 110, 60)

$hRow1 = 30; $hRow3 = 100
$btnHostCheck = New-Btn 'Check this PC'        8   $hRow1 120 27 $pnlHostTop
$btnHostCopy  = New-Btn 'Copy UNC'             134 $hRow1 90  27 $pnlHostTop
$btnHostFix   = New-Btn 'Apply safe fix...'    230 $hRow1 130 27 $pnlHostTop
$btnShareAll  = New-Btn 'Share to everyone...' 366 $hRow1 150 27 $pnlHostTop
$lblHostState = New-Lbl 'Not checked yet.' 524 ($hRow1 + 5) $pnlHostTop $true
$lblHostShares = New-Lbl '' 8 ($hRow1 + 34) $pnlHostTop
$lblHostShares.ForeColor = [System.Drawing.Color]::DimGray
$lblHostShares.MaximumSize = New-Object System.Drawing.Size(1040, 32)

$btnHostRpcPort = New-Btn 'Fix print RPC (recommended)...' 8   $hRow3 210 27 $pnlHostTop
$btnHostRpc     = New-Btn 'Advanced: named pipes...'       224 $hRow3 170 27 $pnlHostTop
$lblHostRpc = New-Lbl '' 402 ($hRow3 + 5) $pnlHostTop
$lblHostRpc.ForeColor = [System.Drawing.Color]::DimGray

$txtHost = New-Object System.Windows.Forms.TextBox
$txtHost.Dock = 'Fill'
$txtHost.Multiline = $true; $txtHost.ReadOnly = $true; $txtHost.ScrollBars = 'Both'
$txtHost.WordWrap = $false
$txtHost.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtHost.BackColor = [System.Drawing.Color]::White
$pnlHostBot.Controls.Add($txtHost)


# --------------------------------------------------------------------------
# Log pane + status
# --------------------------------------------------------------------------
$status = New-Object System.Windows.Forms.Label
$status.Dock = 'Bottom'
$status.Height = 20
$status.Text = 'Ready.'
$pnlFooter.Controls.Add($status)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Dock = 'Fill'
$txtLog.Multiline = $true; $txtLog.ReadOnly = $true
# Both scrollbars: a long path in the transcript was running off the right edge
# with no way to reach it.
$txtLog.ScrollBars = 'Both'
$txtLog.WordWrap = $false
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtLog.BackColor = [System.Drawing.Color]::White
$pnlFooter.Controls.Add($txtLog)
$txtLog.BringToFront()

function Write-Log {
    param([string]$Message, [string]$Tag = 'INFO')
    $line = "{0}  [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Tag, $Message
    $txtLog.AppendText($line + [Environment]::NewLine)
    $status.Text = $Message
}
function Set-Busy {
    param([bool]$On, [string]$What = '')
    $script:Busy = $On
    $form.Cursor = if ($On) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default }
    # Skip anything that is not there, so moving or removing a button can never
    # turn every action into an error.
    foreach ($b in @($btnRescan, $btnRefreshP, $btnRename, $btnDelete, $btnSetDef, $btnTestPort,
                     $btnTestPage, $btnRunDiag, $btnSpoolRefresh, $btnSpoolRestart,
                     $btnSpoolClear, $btnCancelJob, $btnCancelAll, $btnEnumerate, $btnDiscover,
                     $btnRepoint, $btnReport, $btnHostCheck, $btnAddToPc, $btnCreds,
                     $btnHostFix, $btnInstallDrv, $btnDrvInfo, $btnShareAll,
                     $btnSrvTest, $btnSrvRemoveMachine, $btnHostRpcPort, $btnHostRpc, $btnSrvRpc, $btnJunk, $btnPaper)) {
        if ($b) { $b.Enabled = -not $On }
    }
    if ($On -and $What) { $status.Text = $What }
}
# Short in-session operations pump messages through this so the window never
# greys out while a one-shot task runs in the user's session.
$script:Pump = { [System.Windows.Forms.Application]::DoEvents() }


# --------------------------------------------------------------------------
# Background worker: keeps the UI alive during the slow reads (a dead network
# queue blocks WMI for 30-45s, and a reachability sweep is seconds per host).
# The runspace starts with NO session state, so it rebuilds the toolkit from
# $script:ToolkitSource - the same text the UI thread dot-sourced.
# --------------------------------------------------------------------------
$script:Worker = {
    param($Shared, $ToolkitSource, $Mode, $TargetSid, $StuckHours, $SkipNet, $ServerName, $Payload, $CredUser, $CredPass)

    function Emit([string]$m) { [void]$Shared.Log.Add(("{0}  [INFO] {1}" -f (Get-Date -Format 'HH:mm:ss'), $m)) }
    try {
        . ([scriptblock]::Create($ToolkitSource))

        if ($Mode -eq 'scan') {
            $Shared.Status = 'Reading printers...'; $Shared.Progress = 10
            Emit 'Enumerating local print queues...'
            $Shared.Printers = @(Get-LocalPrinterInventory -IncludeVirtual)

            $Shared.Status = 'Reading user hives...'; $Shared.Progress = 40
            Emit 'Reading per-user mapped printers from HKEY_USERS...'
            $Shared.Profiles = @(Get-UserPrinterProfiles)
            $Shared.Maps     = @(Get-AllUserPrinterMaps)

            $Shared.Status = 'Reading spooler + policy...'; $Shared.Progress = 65
            $Shared.SpoolerHealth = Get-SpoolerHealth
            $Shared.Policy        = Get-PointAndPrintPolicy

            $Shared.Status = 'Reading jobs + events...'; $Shared.Progress = 85
            $Shared.Jobs   = @(Get-PrinterJobSnapshot)
            $Shared.Events = Get-PrintServiceEvents -LookbackDays 14

            $Shared.Progress = 100
            Emit ("Scan complete: {0} queue(s), {1} profile(s), {2} job(s)." -f @($Shared.Printers).Count, @($Shared.Profiles).Count, @($Shared.Jobs).Count)
            $Shared.ScanVersion = [int]$Shared.ScanVersion + 1
        }
        elseif ($Mode -eq 'diag') {
            $Shared.Status = 'Running diagnostics...'; $Shared.Progress = 20
            Emit 'Correlating printers, servers, policy, jobs and events...'
            $d = Invoke-MappedPrinterDiagnostics -StuckJobHours $StuckHours -SkipReachability:$SkipNet
            $Shared.Diag = $d
            $Shared.Progress = 100
            Emit ('Diagnostics complete: ' + $d.Summary)
            $Shared.DiagVersion = [int]$Shared.DiagVersion + 1
        }
        elseif ($Mode -eq 'server') {
            $Shared.Status = "Enumerating $ServerName..."; $Shared.Progress = 30
            Emit "Enumerating shared queues on \\$ServerName ..."
            $inv = Get-PrintServerInventory -Server $ServerName -User $CredUser -Password $CredPass
            $Shared.ServerInv = $inv
            $Shared.Progress = 100
            Emit ("$ServerName : " + @($inv.Printers).Count + " queue(s) via " + $inv.Method)
            $Shared.ServerVersion = [int]$Shared.ServerVersion + 1
        }
        elseif ($Mode -eq 'repoint') {
            $Shared.Status = 'Repointing...'; $Shared.Progress = 20
            Emit "Repointing '$($Payload.Name)' onto $($Payload.Protocol) $($Payload.IPAddress)..."
            $args = @{
                Name         = $Payload.Name
                IPAddress    = $Payload.IPAddress
                Protocol     = $Payload.Protocol
                PortNumber   = $Payload.PortNumber
                LprQueueName = $Payload.LprQueueName
            }
            if ($Payload.DriverName) { $args.DriverName = $Payload.DriverName }
            if ($Payload.Force)      { $args.Force = $true }
            if (-not $Payload.RestoreDefaults) { $args.SkipDefaultRestore = $true }
            $res = Convert-PrinterQueue @args
            foreach ($s in @($res.Steps)) { [void]$Shared.Log.Add(("{0}  [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $s.Tag, $s.Text)) }
            $Shared.Repoint = $res
            $Shared.Progress = 100
            $Shared.RepointVersion = [int]$Shared.RepointVersion + 1
        }
        elseif ($Mode -eq 'addconn') {
            # The whole add runs here, off the UI thread: the add itself and, on
            # failure, every probe that explains it. Against a host that swallows
            # connections those probes can take tens of seconds, and on the UI thread
            # that froze the window while it was explaining the failure.
            $Shared.Status = "Adding $($Payload.Unc)..."; $Shared.Progress = 20
            Emit "Adding $($Payload.Unc) for $($Payload.ScopeLabel)..."
            $res = $null
            if ($Payload.Scope -eq 'User') {
                $res = Add-PrinterConnectionForUser -ConnectionName $Payload.Unc -User $Payload.User
            } elseif ($Payload.Scope -eq 'Machine') {
                $res = Add-PrinterConnectionForMachine -ConnectionName $Payload.Unc
                if (-not $res.Success) {
                    Emit 'Per-computer add failed - measuring why...'
                    $ev = Get-PrintConnectionEvidence -ConnectionName $Payload.Unc -User $Payload.User -AddError ([string]$res.Message)
                    $res | Add-Member -NotePropertyName Diagnosis -NotePropertyValue $ev.Diagnosis -Force
                    $res | Add-Member -NotePropertyName Message -NotePropertyValue ([string]$res.Message + "`r`n`r`n" + $ev.Text) -Force
                }
            } elseif ($Payload.Scope -eq 'LocalPort') {
                $res = Add-PrinterViaLocalPort -ConnectionName $Payload.Unc -DriverName $Payload.DriverName -Name $Payload.QueueName
            }
            $Shared.AddConn = [pscustomobject]@{ Payload = $Payload; Result = $res }
            $Shared.Progress = 100
            $Shared.AddConnVersion = [int]$Shared.AddConnVersion + 1
        }
        elseif ($Mode -eq 'evidence') {
            $Shared.Status = "Testing $($Payload.Unc)..."; $Shared.Progress = 20
            Emit ("Testing the connection to $($Payload.Unc)" + $(if ([int]$Payload.PinnedPort) { " including pinned port $($Payload.PinnedPort)" } else { '' }) + '...')
            $Shared.Evidence = Get-PrintConnectionEvidence -ConnectionName $Payload.Unc -User $Payload.User -PinnedPort ([int]$Payload.PinnedPort)
            $Shared.Progress = 100
            Emit ('Connection test result: ' + $Shared.Evidence.Diagnosis.Category)
            $Shared.EvidenceVersion = [int]$Shared.EvidenceVersion + 1
        }
        elseif ($Mode -eq 'delivery') {
            $Shared.Status = "Sending a test page to $($Payload.PrinterName)..."; $Shared.Progress = 20
            Emit "Delivery test: $($Payload.PrinterName)"
            if ($Payload.InProcess) {
                $Shared.Delivery = Invoke-PrinterDeliveryTest -PrinterName $Payload.PrinterName -InProcess
            } else {
                $Shared.Delivery = Invoke-PrinterDeliveryTest -PrinterName $Payload.PrinterName -User $Payload.User
            }
            $Shared.Progress = 100
            $Shared.DeliveryVersion = [int]$Shared.DeliveryVersion + 1
        }
        elseif ($Mode -eq 'sharehost') {
            $Shared.Status = 'Checking this PC as a share host...'; $Shared.Progress = 25
            Emit 'Reading shares, permissions, firewall, network profile and SMB auth...'
            $Shared.HostHealth = Get-PrinterShareHealth
            $Shared.Progress = 100
            Emit ('Share host check complete: ' + $Shared.HostHealth.Summary)
            $Shared.HostVersion = [int]$Shared.HostVersion + 1
        }
        elseif ($Mode -eq 'discover') {
            $Shared.Status = 'Discovering print servers...'; $Shared.Progress = 30
            Emit 'Looking for print servers in user hives, the spooler and AD...'
            $Shared.Candidates = @(Get-PrintServerCandidates)
            $Shared.Progress = 100
            Emit ('Found ' + @($Shared.Candidates).Count + ' candidate server(s).')
            $Shared.DiscoverVersion = [int]$Shared.DiscoverVersion + 1
        }
    } catch {
        $Shared.Error = $_.Exception.Message
        [void]$Shared.Log.Add(("{0}  [ERROR] {1}" -f (Get-Date -Format 'HH:mm:ss'), $_.Exception.Message))
    } finally {
        # How many print calls timed out and were abandoned. The UI reports this
        # so a partial read is never presented as a clean one.
        try { $Shared.Abandoned = Get-AbandonedRunspaceCount } catch { $Shared.Abandoned = 0 }
        $Shared.Done = $true
    }
}

function Start-Worker {
    param([string]$Mode, [string]$ServerName = '', $Payload = $null)
    if ($script:Busy) { return }

    $sh = [hashtable]::Synchronized(@{})
    $sh.Log = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
    $sh.Status = ''; $sh.Progress = 0; $sh.Done = $false; $sh.Error = $null
    $sh.ScanVersion = 0; $sh.DiagVersion = 0; $sh.ServerVersion = 0; $sh.DiscoverVersion = 0
    $script:Shared = $sh
    $script:LogIndex = 0
    $script:Mode = $Mode

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($script:Worker.ToString())
    [void]$ps.AddParameter('Shared', $sh)
    [void]$ps.AddParameter('ToolkitSource', $script:ToolkitSource)
    [void]$ps.AddParameter('Mode', $Mode)
    [void]$ps.AddParameter('TargetSid', $script:TargetSid)
    [void]$ps.AddParameter('StuckHours', [int]$numStuck.Value)
    [void]$ps.AddParameter('SkipNet', [bool]$chkSkipNet.Checked)
    [void]$ps.AddParameter('ServerName', $ServerName)
    # Credentials the technician supplied in the Credentials dialog, held for this
    # session only. Never logged, never written to the report or the JSON.
    $cu = ''; $cp = ''
    if ($ServerName -and $script:HostCred) {
        $k = ConvertTo-PrintHostKey $ServerName
        if ($script:HostCred.ContainsKey($k)) { $cu = $script:HostCred[$k].User; $cp = $script:HostCred[$k].Password }
    }
    [void]$ps.AddParameter('CredUser', $cu)
    [void]$ps.AddParameter('CredPass', $cp)
    [void]$ps.AddParameter('Payload', $Payload)

    $script:Ps = $ps
    $script:Rs = $rs
    $script:Async = $ps.BeginInvoke()
    $script:WorkerStarted = Get-Date
    # Staging a driver package and Add-Printer are legitimately slow, so the
    # repoint path gets a much longer ceiling than a read-only scan.
    $script:WorkerTimeoutSec = if ($Mode -eq 'repoint') { 600 } elseif (@('addconn', 'evidence', 'delivery') -contains $Mode) { 300 } else { 180 }
    Set-Busy $true "Working..."
    $script:Timer.Start()
}

$script:Timer = New-Object System.Windows.Forms.Timer
$script:Timer.Interval = 200
$script:Timer.Add_Tick({
    $sh = $script:Shared
    if (-not $sh) { return }

    # A dialog is open: touch NOTHING on the main window. Every repaint under an
    # owned modal dialog makes that dialog's text blank and redraw in Backstage.
    # Nothing is lost by waiting - the log is a high-water mark, so the pending
    # lines are appended on the first tick after the dialog closes, and the
    # watchdog simply defers rather than tearing down a worker while the
    # technician is mid-read.
    if (Test-ModalOpen) { return }

    # Log high-water mark so each line is appended exactly once. Batched into a
    # SINGLE AppendText: one repaint per tick instead of one per line, which
    # matters most during a noisy scan.
    $count = $sh.Log.Count
    if ($script:LogIndex -lt $count) {
        $pending = New-Object System.Text.StringBuilder
        while ($script:LogIndex -lt $count) {
            [void]$pending.AppendLine([string]$sh.Log[$script:LogIndex])
            $script:LogIndex++
        }
        $txtLog.AppendText($pending.ToString())
    }
    # Only when it actually changed - an identical assignment still repaints the
    # status strip, five times a second, for nothing.
    if ($sh.Status -and "$($sh.Status)" -ne $status.Text) { $status.Text = "$($sh.Status)" }

    if (-not $sh.Done) {
        # WATCHDOG. A worker that never finishes must never strand the console.
        # Even with Invoke-WithTimeout fixed, a print call can wedge somewhere
        # unbounded, and an unrecoverable hourglass in Backstage is the worst
        # possible failure - the tech cannot even read what did come back.
        $elapsed = ((Get-Date) - $script:WorkerStarted).TotalSeconds
        if ($elapsed -gt $script:WorkerTimeoutSec) {
            $script:Timer.Stop()
            # BeginStop, not Stop/Dispose: those block until the wedged pipeline ends.
            try { [void]$script:Ps.BeginStop($null, $null) } catch {}
            Write-Log ("Abandoned the background scan after {0:N0}s - a print call did not return. Partial data is kept; press Rescan to retry." -f $elapsed) 'ERROR'
            Set-Busy $false
            $status.Text = 'Scan abandoned (a print call never returned).'
            return
        }
        return
    }

    $script:Timer.Stop()
    # Safe to tear down synchronously: the worker set Done in its own finally,
    # so the pipeline has already run to completion.
    try { $script:Ps.EndInvoke($script:Async) } catch {}
    try { $script:Ps.Dispose() } catch {}
    try { $script:Rs.Close(); $script:Rs.Dispose() } catch {}

    # try/finally is load-bearing. $ErrorActionPreference is 'Stop' at script
    # scope, so ANY error while rendering results would otherwise skip
    # Set-Busy $false and leave the window disabled behind an hourglass with no
    # message explaining why.
    try {
        switch ($script:Mode) {
            'scan'     { Apply-ScanResults $sh }
            'diag'     { Apply-DiagResults $sh }
            'server'   { Apply-ServerResults $sh }
            'discover' { Apply-DiscoverResults $sh }
        'repoint'  { Apply-RepointResults $sh }
        'sharehost'{ Apply-HostResults $sh }
        'addconn'  { Apply-AddConnResults $sh }
        'evidence' { Apply-EvidenceResults $sh }
        'delivery' { Apply-DeliveryResults $sh }
        }
    } catch {
        Write-Log ("Could not render the results: " + $_.Exception.Message) 'ERROR'
    } finally {
        Set-Busy $false
    }

    if ($sh.Error) {
        $status.Text = "Error: $($sh.Error)"
    } else {
        $status.Text = 'Ready.'
        # A silent partial read is worse than a slow one - say so explicitly.
        $abandoned = 0
        try { $abandoned = [int]$sh.Abandoned } catch {}
        if ($abandoned -gt 0) {
            Write-Log ("$abandoned print call(s) timed out and were abandoned - some data below is incomplete (usually an unreachable print server or a dead queue).") 'WARN'
            $status.Text = "Ready - $abandoned call(s) timed out, data is partial."
        }
    }

    # Follow-ups an Apply-* step asked for. They run here, after Busy has been
    # released, because Start-Worker refuses to start while a worker is finishing -
    # calling it from inside Apply-* silently did nothing.
    if ($script:PendingRemedy) {
        $pr = $script:PendingRemedy; $script:PendingRemedy = $null
        Invoke-DiagnosisRemedy -Diagnosis $pr.Diagnosis -Unc $pr.Unc
    }
    if ($script:PendingDelivery -and -not $script:Busy) {
        $pd = $script:PendingDelivery; $script:PendingDelivery = $null; $script:PendingScan = $false
        Start-Worker 'delivery' '' $pd
        return
    }
    if ($script:PendingScan) {
        $script:PendingScan = $false
        if (-not $script:Busy) { Start-Worker 'scan' }
    }
})


# --------------------------------------------------------------------------
# Rendering
# --------------------------------------------------------------------------
function Get-SelectedUserMap {
    if (-not $script:TargetSid) { return $null }
    @($script:Maps) | Where-Object { $_.Sid -eq $script:TargetSid } | Select-Object -First 1
}

function Build-PrinterRows {
    <#
      Merge what the SPOOLER can see with what the SELECTED USER actually has.
      Run as SYSTEM those are different sets: the user's mapped printers appear
      only in their hive, so they are added here as 'Connection (user)' rows.
      Default is taken from the USER's hive, never from the spooler's Default
      flag (which as SYSTEM is SYSTEM's own default).
    #>
    $map = Get-SelectedUserMap
    $userDefault = if ($map) { "$($map.DefaultPrinter)" } else { '' }

    $rows = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    foreach ($p in @($script:Printers)) {
        $key = "$($p.Name)".ToLowerInvariant()
        $seen[$key] = $true
        $rows.Add([pscustomobject]@{
            Name    = $p.Name
            Type    = $p.Type
            PortKind= $p.PortType
            Server  = $p.ServerName
            Driver  = $p.DriverName
            Port    = $p.PortName
            HostIp  = $p.HostAddress
            Default = $(if ($userDefault -and $userDefault -ieq "$($p.Name)") { 'YES' } else { '' })
            Status  = ("{0}{1}" -f $p.PrinterStatus, $(if ($p.WorkOffline) { ' (OFFLINE FLAG)' } elseif ($p.StateText -and $p.StateText -ne 'Ready') { " / $($p.StateText)" } else { '' }))
            Raw     = $p
            IsUserOnly = $false
        })
    }
    if ($map) {
        foreach ($c in @($map.Connections)) {
            $key = "$($c.Unc)".ToLowerInvariant()
            if ($seen.ContainsKey($key)) { continue }
            $rows.Add([pscustomobject]@{
                Name    = $c.Unc
                Type    = 'Connection (user)'
                PortKind= ''
                Server  = $c.Server
                Driver  = '(on server)'
                Port    = ''
                HostIp  = ''
                Default = $(if ($userDefault -and $userDefault -ieq "$($c.Unc)") { 'YES' } else { '' })
                Status  = 'from user hive'
                Raw     = $c
                IsUserOnly = $true
            })
        }
    }
    $rows.ToArray()
}

function Refresh-PrinterGrid {
    # Remember WHAT was selected, never WHERE. Row indices do not survive a
    # rebuild: Get-SelectedPrinterRow maps the selected INDEX into
    # $script:PrinterRows, so if a refresh lands while a printer is selected, the
    # same index can now be a different printer and every button on this tab
    # silently retargets. That is how a label printer got deleted - the grid
    # refreshed under the selection and the confirm named a different queue.
    $keepName = $null
    try {
        $sel = Get-SelectedPrinterRow
        if ($sel) { $keepName = [string]$sel.Name }
    } catch {}

    $script:PrinterRows = @(Build-PrinterRows)
    Set-Rows $dgvPrinters $script:PrinterRows @('Name', 'Type', 'PortKind', 'Server', 'Driver', 'Port', 'HostIp', 'Default', 'Status')

    # Adding rows makes a DataGridView select the first one, which would silently
    # aim every button at whatever now sits at the top. Clear it, then restore the
    # selection ONLY if the same printer is still present. If it is gone, nothing
    # is selected and the buttons refuse - which is the safe answer.
    $dgvPrinters.ClearSelection()
    if ($keepName) {
        $rows = @($script:PrinterRows)
        for ($i = 0; $i -lt $rows.Count -and $i -lt $dgvPrinters.Rows.Count; $i++) {
            if ([string]$rows[$i].Name -eq $keepName) {
                $dgvPrinters.Rows[$i].Selected = $true
                try { $dgvPrinters.CurrentCell = $dgvPrinters.Rows[$i].Cells[0] } catch {}
                break
            }
        }
    }
    # Amber the port kinds that cannot survive a device move - they are the
    # rebuild candidates the Repoint button exists for.
    for ($i = 0; $i -lt $dgvPrinters.Rows.Count; $i++) {
        if ("$($dgvPrinters.Rows[$i].Cells[2].Value)" -in @('WSD', 'IPP')) {
            $dgvPrinters.Rows[$i].Cells[2].Style.ForeColor = [System.Drawing.Color]::FromArgb(180, 95, 0)
            $dgvPrinters.Rows[$i].Cells[2].Style.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
        }
    }
    $map = Get-SelectedUserMap
    if ($map) {
        $legacy = $map.LegacyDefaultPrinterMode
        $warn = if ($null -eq $legacy -or [int]$legacy -ne 1) { '  WARNING: Windows manages this user''s default and will overwrite it.' } else { '' }
        $lblUserNote.Text = "Default: $($map.DefaultPrinter)$warn"
        $lblUserNote.ForeColor = if ($warn) { [System.Drawing.Color]::Firebrick } else { [System.Drawing.Color]::DimGray }
    } else {
        $lblUserNote.Text = 'No user selected - per-user actions are unavailable.'
        $lblUserNote.ForeColor = [System.Drawing.Color]::Firebrick
    }
}

function Apply-ScanResults {
    param($sh)
    $script:Printers = @($sh.Printers)
    $script:Maps     = @($sh.Maps)
    $script:Profiles = @($sh.Profiles)
    $script:Jobs     = @($sh.Jobs)

    # Keep the junk-queue line honest after any change, including a plain Delete.
    Update-JunkLabel

    # Rebuild the user list, preserving the current selection where possible.
    $prev = $script:TargetSid
    $cboUser.Items.Clear()
    foreach ($p in $script:Profiles) {
        [void]$cboUser.Items.Add(("{0}{1}" -f $p.User, $(if ($p.IsActive) { '   [signed in]' } else { '' })))
    }
    if ($cboUser.Items.Count) {
        $idx = 0
        if ($prev) {
            for ($i = 0; $i -lt $script:Profiles.Count; $i++) { if ($script:Profiles[$i].Sid -eq $prev) { $idx = $i; break } }
        } else {
            # Default to the active session - that is who the tech is helping.
            for ($i = 0; $i -lt $script:Profiles.Count; $i++) { if ($script:Profiles[$i].IsActive) { $idx = $i; break } }
        }
        $cboUser.SelectedIndex = $idx
    } else {
        $script:TargetUser = $null; $script:TargetSid = $null
        Refresh-PrinterGrid
    }

    # Spooler tab
    $h = $sh.SpoolerHealth
    if ($h) {
        $lblSpool.Text = "Spooler: $($h.Status) ($($h.StartMode)), account $($h.Account)"
        $lblSpool.ForeColor = if ($h.Running) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::Firebrick }
        $spoolTxt = if ($h.SpoolReadable) {
            "Spool dir: $($h.SpoolDirectory)   files: $($h.SpoolFileCount)" + $(if ($h.OldestSpoolAgeHours) { "   oldest: $($h.OldestSpoolAgeHours)h ($($h.OldestSpoolFile))" } else { '' })
        } else {
            "Spool dir: $($h.SpoolDirectory)   (not readable from this context - run as SYSTEM/admin)"
        }
        $lblSpoolDir.Text = $spoolTxt
        $script:SpoolerHealth = $h
    }
    Set-Rows $dgvJobs $script:Jobs @('Queue', 'Id', 'UserName', 'Document', 'JobStatus', 'PagesPrinted', 'AgeHours', 'PaperSize')
    # Amber: a job asking for a different paper size from this PC's region - the kind
    # that makes an A4 printer stop and ask for Letter. Pause the queue to catch one.
    $jobExpect = 0
    try { $jobExpect = [int](Get-RegionPaperSize -Sid '.DEFAULT').ExpectedPaperCode } catch {}
    if ($jobExpect) {
        $jobArr = @($script:Jobs)
        for ($i = 0; $i -lt $dgvJobs.Rows.Count -and $i -lt $jobArr.Count; $i++) {
            if ((Get-PaperMatchState -Code ([int]$jobArr[$i].PaperCode) -Expected $jobExpect) -eq 'Mismatch') {
                $dgvJobs.Rows[$i].Cells[7].Style.ForeColor = [System.Drawing.Color]::FromArgb(180, 95, 0)
                $dgvJobs.Rows[$i].Cells[7].Style.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
            }
        }
    }

    # Events
    $ev = $sh.Events
    if ($ev) {
        $sb = New-Object System.Text.StringBuilder
        foreach ($c in @($ev.Channels)) {
            $st = if ($null -eq $c.Enabled) { 'unknown' } elseif ([int]$c.Enabled -eq 1) { 'enabled' } else { 'DISABLED - nothing is being logged' }
            [void]$sb.AppendLine(("{0,-52} {1}" -f $c.Channel, $st))
        }
        [void]$sb.AppendLine('')
        if (@($ev.SpoolerCrashes).Count) {
            [void]$sb.AppendLine("*** $(@($ev.SpoolerCrashes).Count) spooler crash/unexpected-stop event(s) - usually a bad third-party driver or port monitor ***")
            foreach ($e in @($ev.SpoolerCrashes)) { [void]$sb.AppendLine(("{0}  {1,-12} {2,-5} {3}" -f $e.Time, $e.Channel, $e.Id, $e.Message)) }
            [void]$sb.AppendLine('')
        }
        if (@($ev.Events).Count) {
            foreach ($e in @($ev.Events)) { [void]$sb.AppendLine(("{0}  {1,-5} {2}" -f $e.Time, $e.Id, $e.Message)) }
        } else {
            [void]$sb.AppendLine("(no matching events in the last $($ev.LookbackDays) days)")
        }
        $txtEvents.Text = $sb.ToString()
    }

    # Policy tab
    if ($sh.Policy) {
        Set-Rows $dgvPolicy @($sh.Policy.Machine) @('Name', 'State', 'Path')
        for ($i = 0; $i -lt $dgvPolicy.Rows.Count; $i++) {
            $st = "$($dgvPolicy.Rows[$i].Cells[1].Value)"
            if ($st -like 'Set=*') { $dgvPolicy.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkBlue }
        }
    }
}

function Apply-DiagResults {
    param($sh)
    $d = $sh.Diag
    if (-not $d) { return }
    $script:Findings = @($d.Findings)
    Set-Rows $dgvFindings $script:Findings @('Code', 'Severity', 'Scope', 'User', 'Printer', 'Detail')
    for ($i = 0; $i -lt $dgvFindings.Rows.Count; $i++) {
        switch ("$($dgvFindings.Rows[$i].Cells[1].Value)") {
            'ERROR' { $dgvFindings.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick }
            'WARN'  { $dgvFindings.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 95, 0) }
            default { $dgvFindings.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::DimGray }
        }
    }
    Write-Log $d.Summary $(if ($d.ErrorCount) { 'ERROR' } elseif ($d.WarnCount) { 'WARN' } else { 'OK' })
    if ($script:Findings.Count -eq 0) { $txtFinding.Text = 'No printer problems detected.' }
}

function Show-CredInUse {
    <# Say plainly when Enumerate is authenticating, and as whom. Never the password. #>
    $srv = ConvertTo-PrintHostKey $txtServer.Text
    if ($srv -and $script:HostCred -and $script:HostCred.ContainsKey($srv)) {
        $lblServerNote.Text = "Using the credentials supplied for $srv ($($script:HostCred[$srv].User)) - held for this session only."
        $lblServerNote.ForeColor = [System.Drawing.Color]::DarkGreen
    }
}
function Apply-ServerResults {
    param($sh)
    Show-CredInUse
    $inv = $sh.ServerInv
    if (-not $inv) { return }
    $script:ServerRows = @($inv.Printers)
    Set-Rows $dgvServer $script:ServerRows @('ShareName', 'Name', 'DriverName', 'PortName', 'HostAddress', 'Unc')
    # Nothing preselected: Add and Test must act on a row the technician chose, not on
    # whatever the grid selected by itself when the rows were added.
    $dgvServer.ClearSelection()
    $lblServerNote.Text = "$($inv.Server): reachable=$($inv.Reachable), method=$($inv.Method). $($inv.Message)"
    # Access denied is the case a credential can actually fix - offer it rather
    # than leaving the tech to find the button.
    if ($inv.Reachable -and @($inv.Printers).Count -eq 0) {
        if ((Show-Msg -Text "No shares could be listed from '$($inv.Server)', though the host is reachable.`n`nThis is often an access-denied on share ENUMERATION, which is usually an admin-only right and does not necessarily stop printing.`n`nOpen the credentials dialog to test access and, if needed, supply credentials?" -Title 'No shares listed' -Buttons 'YesNo' -Icon 'Question') -eq 'Yes') {
            Show-CredentialsDialog -Server $inv.Server
        }
    }
}

function Apply-DiscoverResults {
    param($sh)
    $c = @($sh.Candidates)
    if ($c.Count -eq 0) { Write-Log 'No print servers discovered.' 'WARN'; return }
    $txtServer.Text = $c[0].Server
    $lblServerNote.Text = 'Discovered: ' + (($c | ForEach-Object { "$($_.Server) (users=$($_.UserCount); $($_.Sources -join ', '))" }) -join '   |   ')
    Write-Log ("Discovered " + $c.Count + " print server(s); pre-filled " + $c[0].Server) 'OK'
}


# --------------------------------------------------------------------------
# Selection helpers
# --------------------------------------------------------------------------
function Get-SelectedPrinterRow {
    if ($dgvPrinters.SelectedRows.Count -eq 0) { return $null }
    $i = $dgvPrinters.SelectedRows[0].Index
    if ($i -lt 0 -or $i -ge @($script:PrinterRows).Count) { return $null }
    @($script:PrinterRows)[$i]
}
function Require-Printer {
    $r = Get-SelectedPrinterRow
    if (-not $r) { [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 'Select a printer first.', 'Printer Manager', 'OK', 'Information') }
    $r
}
function Require-User {
    if (-not $script:TargetUser) {
        [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 
            "No target user is selected.`n`nPer-user printer actions must run inside a user's session - a SYSTEM process changing printers only ever changes SYSTEM's own profile.",
            'Printer Manager', 'OK', 'Warning')
        return $false
    }
    $true
}
function Show-TextDialog {
    <#
    .SYNOPSIS
        A long message in a proper scrollable window instead of a MessageBox.
    .DESCRIPTION
        A native MessageBox has no scrollbar and no width of its own: it grows
        WIDE to fit the text before it wraps, up to a screen-relative cap. The
        diagnosis strings in this toolkit run to 1100 characters with no line
        breaks at all - measured, that wants a 6120px-wide box, over three times
        a 1920 screen and far worse on the 1024-wide screen a Backstage session
        gives you. The result is an enormous dialog that flickers unreadably in
        ScreenConnect, while short confirmations on the same machine are fine.

        So anything long goes here instead: a real owned form, word-wrapped,
        scrollable, resizable, and sized to the screen it is actually on. Short
        confirmations keep using MessageBox, which handles them perfectly well.
    .OUTPUTS
        'OK' | 'Yes' | 'No'
    #>
    param(
        [string]$Text,
        [string]$Title = 'Printer Manager',
        [ValidateSet('OK','YesNo')][string]$Buttons = 'OK',
        [ValidateSet('Information','Warning','Question','Error')][string]$Icon = 'Information',
        [ValidateSet('Yes','No')][string]$Default = 'Yes'
    )

    $owner = Get-MsgOwner

    # Size to the screen the console is actually on, never to the text.
    $scr = $null
    try { $scr = if ($owner) { [System.Windows.Forms.Screen]::FromControl($owner).WorkingArea }
                 else { [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea } } catch {}
    $maxW = 660; $maxH = 460
    if ($scr) {
        $maxW = [Math]::Max(420, [Math]::Min(660, [int]($scr.Width  * 0.7)))
        $maxH = [Math]::Max(320, [Math]::Min(460, [int]($scr.Height * 0.7)))
    }

    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Title
    $f.Size = New-Object System.Drawing.Size($maxW, $maxH)
    $f.MinimumSize = New-Object System.Drawing.Size(400, 260)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'Sizable'
    $f.MinimizeBox = $false; $f.MaximizeBox = $true
    $f.ShowInTaskbar = $false
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $bot = New-DockBand $f 'Bottom' 46

    $box = New-Object System.Windows.Forms.TextBox
    $box.Multiline = $true
    $box.ReadOnly = $true
    $box.WordWrap = $true                     # the whole point - MessageBox will not
    $box.ScrollBars = 'Vertical'
    $box.Dock = 'Fill'
    $box.BackColor = [System.Drawing.Color]::White
    $box.Text = $Text
    $f.Controls.Add($box)
    $box.BringToFront()

    $result = $(if ($Buttons -eq 'YesNo') { 'No' } else { 'OK' })

    if ($Buttons -eq 'YesNo') {
        $bYes = New-Btn 'Yes' 10 9 100 27 $bot
        $bNo  = New-Btn 'No'  116 9 100 27 $bot
        $bYes.Add_Click({ $script:TextDialogResult = 'Yes'; $f.Close() })
        $bNo.Add_Click({  $script:TextDialogResult = 'No';  $f.Close() })
        # A destructive action must not sit under the Enter key, and must not be
        # where a fast second click lands.
        if ($Default -eq 'No') { $f.AcceptButton = $bNo } else { $f.AcceptButton = $bYes }
        $f.CancelButton = $bNo
    } else {
        $bOk = New-Btn 'OK' 10 9 100 27 $bot
        $bOk.Add_Click({ $script:TextDialogResult = 'OK'; $f.Close() })
        $f.AcceptButton = $bOk
        $f.CancelButton = $bOk
    }

    $bCopy = New-Btn 'Copy' 232 9 90 27 $bot
    $bCopy.Add_Click({
        try { [System.Windows.Forms.Clipboard]::SetText($box.Text) } catch {}
    })

    $script:TextDialogResult = $result
    $f.Add_Shown({
        $box.SelectionStart = 0
        $box.SelectionLength = 0
        if ($f.AcceptButton) { $f.AcceptButton.Focus() | Out-Null }
    })
    [void]$f.ShowDialog($owner)
    $f.Dispose()
    return $script:TextDialogResult
}

function Show-Msg {
    <#
    .SYNOPSIS
        Show a message, picking the presenter that suits its LENGTH.
    .DESCRIPTION
        Short text -> native MessageBox, which is compact and familiar and which
        is confirmed fine in Backstage. Long text -> Show-TextDialog, because a
        MessageBox cannot wrap or scroll and turns a long diagnosis into an
        oversized flickering window.

        One entry point so no call site has to make that judgement.
    .OUTPUTS
        'OK' | 'Yes' | 'No'
    #>
    param(
        [string]$Text,
        [string]$Title = 'Printer Manager',
        [ValidateSet('OK','YesNo')][string]$Buttons = 'OK',
        [ValidateSet('Information','Warning','Question','Error')][string]$Icon = 'Information',
        [ValidateSet('Yes','No')][string]$Default = 'Yes'
    )
    $t = "$Text"
    $lines = ([regex]::Matches($t, "`n")).Count + 1
    # 220 chars is about three lines in a default MessageBox - past that it starts
    # growing wide rather than tall, which is the behaviour being avoided.
    if ($t.Length -gt 220 -or $lines -gt 6) {
        return (Show-TextDialog -Text $t -Title $Title -Buttons $Buttons -Icon $Icon -Default $Default)
    }
    # Button2 = the SECOND button (No). Without this the default is Yes, so Enter
    # or a stray fast click confirms a destructive action.
    $db = if ($Default -eq 'No') { [System.Windows.Forms.MessageBoxDefaultButton]::Button2 }
          else { [System.Windows.Forms.MessageBoxDefaultButton]::Button1 }
    $r = [System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), $t, $Title, $Buttons, $Icon, $db)
    return "$r"
}

function Show-Result {
    param($Result, [string]$OkTitle = 'Done')
    # Via Show-Msg: the diagnosis strings run past 1000 characters with no line
    # breaks, and a MessageBox grows WIDE rather than wrapping them.
    if ($Result.Success) {
        Write-Log $Result.Message 'OK'
        [void](Show-Msg -Text $Result.Message -Title $OkTitle -Icon 'Information')
    } else {
        Write-Log $Result.Message 'WARN'
        [void](Show-Msg -Text $Result.Message -Title 'Not completed' -Icon 'Warning')
    }
}
function Show-InputBox {
    <# Minimal modal prompt - Read-Host is useless in a GUI and VB InputBox is not always present. #>
    param([string]$Prompt, [string]$Title, [string]$Default = '')
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Title; $f.Size = New-Object System.Drawing.Size(520, 175)
    $f.StartPosition = 'CenterParent'; $f.FormBorderStyle = 'FixedDialog'
    $f.MinimizeBox = $false; $f.MaximizeBox = $false
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Prompt; $l.Location = New-Object System.Drawing.Point(12, 12)
    $l.Size = New-Object System.Drawing.Size(480, 36); $f.Controls.Add($l)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Text = $Default; $t.Location = New-Object System.Drawing.Point(12, 54)
    $t.Size = New-Object System.Drawing.Size(480, 24); $f.Controls.Add($t)
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'OK'; $ok.Location = New-Object System.Drawing.Point(316, 92)
    $ok.Size = New-Object System.Drawing.Size(84, 28); $ok.DialogResult = 'OK'; $f.Controls.Add($ok)
    $ca = New-Object System.Windows.Forms.Button
    $ca.Text = 'Cancel'; $ca.Location = New-Object System.Drawing.Point(408, 92)
    $ca.Size = New-Object System.Drawing.Size(84, 28); $ca.DialogResult = 'Cancel'; $f.Controls.Add($ca)
    $f.AcceptButton = $ok; $f.CancelButton = $ca
    $r = $f.ShowDialog()
    $v = $t.Text
    $f.Dispose()
    if ($r -eq 'OK') { return $v }
    return $null
}


# --------------------------------------------------------------------------
# Handlers
# --------------------------------------------------------------------------
$cboUser.Add_SelectedIndexChanged({
    $i = $cboUser.SelectedIndex
    if ($i -lt 0 -or $i -ge @($script:Profiles).Count) { return }
    $p = @($script:Profiles)[$i]
    $script:TargetUser = $p.User
    $script:TargetSid  = $p.Sid
    Refresh-PrinterGrid
})

$btnRescan.Add_Click({ Start-Worker 'scan' })
$btnRefreshP.Add_Click({ Start-Worker 'scan' })

$btnRename.Add_Click({
    $r = Require-Printer; if (-not $r) { return }
    if ($r.Name -match '^\\\\') {
        [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 
            "'$($r.Name)' is a connection to a shared queue.`n`nIts name IS the server's UNC path, so it cannot be renamed on this machine. Rename the queue on the print server, then re-add the connection.",
            'Cannot rename a connection', 'OK', 'Information')
        return
    }
    $new = Show-InputBox -Prompt "New name for '$($r.Name)':" -Title 'Rename printer' -Default $r.Name
    if ([string]::IsNullOrWhiteSpace($new) -or $new -eq $r.Name) { return }
    Set-Busy $true 'Renaming...'
    try { Show-Result (Rename-PrinterSafe -Name $r.Name -NewName $new) 'Renamed' } finally { Set-Busy $false }
    Start-Worker 'scan'
})

$btnDelete.Add_Click({
    $r = Require-Printer; if (-not $r) { return }
    $isConn = ($r.Name -match '^\\\\')
    $who = if ($isConn) { " from $($script:TargetUser)" } else { ' from this machine' }
    if ($isConn -and -not (Require-User)) { return }
    $msg = "Delete '$($r.Name)'$who?" + $(if ($isConn) {
        "`n`nThis removes the mapping for that user only. The driver and port stay in place for other users."
    } else {
        "`n`nThis removes the queue for ALL users of this machine."
    })
    if ((Show-Msg -Text $msg -Title "Confirm delete - $($r.Name)" -Buttons 'YesNo' -Icon 'Warning' -Default 'No') -ne 'Yes') { return }
    # Breadcrumb BEFORE it goes, so the queue can be rebuilt from the transcript
    # if this turns out to be the wrong one.
    Write-Log ("Deleting '{0}'  [driver: {1}] [port: {2}] [type: {3}] - note these if it needs rebuilding." -f `
               $r.Name, $r.Driver, $r.Port, $r.Type) 'WARN'
    Set-Busy $true 'Deleting...'
    try {
        $res = if ($isConn) { Remove-PrinterSafe -Name $r.Name -User $script:TargetUser -Sid $script:TargetSid -OnPoll $script:Pump }
               else         { Remove-PrinterSafe -Name $r.Name }
        Show-Result $res 'Deleted'
    } finally { Set-Busy $false }
    Start-Worker 'scan'
})

$btnSetDef.Add_Click({
    $r = Require-Printer; if (-not $r) { return }
    if (-not (Require-User)) { return }
    if ([System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 
            "Set '$($r.Name)' as the default printer for $($script:TargetUser)?`n`nLegacyDefaultPrinterMode will also be set to 1 so Windows does not overwrite it after the next print.",
            'Confirm default printer', 'YesNo', 'Question') -ne 'Yes') { return }
    Set-Busy $true 'Setting default printer...'
    try {
        $res = Set-DefaultPrinterForUser -PrinterName $r.Name -User $script:TargetUser -Sid $script:TargetSid -OnPoll $script:Pump
        Show-Result $res 'Default printer set'
    } finally { Set-Busy $false }
    Start-Worker 'scan'
})

$btnTestPort.Add_Click({
    $r = Require-Printer; if (-not $r) { return }
    Set-Busy $true 'Testing reachability...'
    try {
        $lines = New-Object System.Collections.Generic.List[string]

        # Server side first - for a mapped printer this is what usually breaks.
        if ($r.Server) {
            $s = Test-PrintServerReachable -Server $r.Server
            $lines.Add("Print server $($r.Server)")
            $lines.Add("   SMB 445 : $(if ($s.Smb445) { 'open' } else { 'no response' })")
            $lines.Add("   RPC 135 : $(if ($s.Rpc135) { 'open' } else { 'no response' })")
            $lines.Add("   print$  : $(if ($s.DriverShare) { 'readable' } else { 'not readable from this context' })")
            $lines.Add("   => $($s.Message)")
            $lines.Add('')
            $lines.Add('Note: a print share is not a filesystem object, so testing \\server\printer')
            $lines.Add('with Test-Path always fails even when printing works. These are the valid probes.')
            $lines.Add('')
        }

        # USB first. A USB queue has no network host, but "is the printer actually
        # plugged in and switched on" is the same question the tech is asking, and
        # Windows will not answer it - PrinterStatus reads Normal for a printer
        # that is not there.
        $isUsb = ($r.Raw -and "$($r.Raw.PortType)" -eq 'USB')
        if ($isUsb) {
            $u = Test-UsbPrinterAttached -PortName $r.Raw.PortName -Name $r.Name -DriverName $r.Raw.DriverName
            $lines.Add("USB port $($r.Raw.PortName)")
            $lines.Add("   device : $(if ($u.Attached) { 'ATTACHED' } else { 'NOT ATTACHED' })")
            if ($u.Device -and $u.Device.FriendlyName) { $lines.Add("   name   : $($u.Device.FriendlyName)") }
            if ($null -ne $u.WorkOffline) { $lines.Add("   queue  : $(if ($u.WorkOffline) { 'flagged Use Printer Offline' } else { 'online' })") }
            $lines.Add("   => $($u.Message)")
            $lines.Add('')
        }

        $target = $r.HostIp
        if (-not $target -and $r.Raw -and $r.Raw.HostAddress) { $target = $r.Raw.HostAddress }

        # A WSD port records no address at all, which is exactly the queue a tech
        # most wants to test. Fall back to the same free (no-network) discovery the
        # Repoint dialog uses, and say where the address came from so nobody
        # mistakes a guess for fact.
        $guessedFrom = ''
        if (-not $target -and $r.Raw) {
            try {
                $cand = @(Resolve-PrinterAddress -Printer $r.Raw) | Select-Object -First 1
                if ($cand) { $target = $cand.Address; $guessedFrom = $cand.Source }
            } catch {}
        }
        if ($guessedFrom) {
            $lines.Add("This port records no address (kind: $(if ($r.Raw) { $r.Raw.PortType } else { '?' })).")
            $lines.Add("Testing the best guess instead - from $guessedFrom.")
            $lines.Add('')
        }

        if ($target) {
            $p = Test-PrinterPortReachable -HostName $target
            $lines.Add("Device $target")
            foreach ($x in @($p.Results)) {
                # The outer parentheses are REQUIRED. Inside a method call the commas
                # are parsed as method arguments, so .Add("{0}{1}{2}" -f $a, $b, $c)
                # becomes Add(("...".-f $a), $b, $c) - the format string then gets one
                # argument and {1} throws "Index must be ... less than the size of the
                # argument list". Always wrap a -f expression used as an argument.
                $cat   = @(Get-PrinterPortCatalogue | Where-Object { $_.Port -eq $x.Port })[0]
                $label = if ($cat) { $cat.Name } else { '' }
                $state = if ($x.Open) { 'OPEN' } else { 'closed' }
                $lines.Add(("   {0,-5} {1,-22} {2}" -f $x.Port, $label, $state))
            }
            $lines.Add('')
            $lines.Add("   => $($p.Message)")
            if ($p.AliveOnly) {
                $lines.Add('')
                $lines.Add('The device answered on a non-printing port, so it is powered on and on')
                $lines.Add('the network. Check that RAW/9100 (or LPD) is enabled in its web interface.')
            }
        } elseif (-not $r.Server -and -not $isUsb) {
            $lines.Add("'$($r.Name)' has no network host and no USB device to test (port type: $(if ($r.Raw) { $r.Raw.PortType } else { 'unknown' })).")
        }

        $text = ($lines -join [Environment]::NewLine)
        Write-Log ("Reachability tested for " + $r.Name) 'INFO'
        [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), $text, "Reachability - $($r.Name)", 'OK', 'Information')
    } catch {
        Write-Log ("Reachability test failed: " + $_.Exception.Message) 'ERROR'
        [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 
            "Could not complete the reachability test:`n`n$($_.Exception.Message)",
            'Reachability test failed', 'OK', 'Warning')
    } finally { Set-Busy $false }
})

$btnTestPage.Add_Click({
    $r = Require-Printer; if (-not $r) { return }
    if (-not (Require-User)) { return }
    if ([System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), "Print a Windows test page to '$($r.Name)' as $($script:TargetUser)?", 'Confirm test page', 'YesNo', 'Question') -ne 'Yes') { return }
    Set-Busy $true 'Printing test page...'
    try { Show-Result (Invoke-PrinterTestPage -PrinterName $r.Name -User $script:TargetUser -OnPoll $script:Pump) 'Test page' }
    finally { Set-Busy $false }
})

$btnPrintMgmt.Add_Click({
    # printmanagement.msc works in Backstage; launch it through mmc.exe explicitly
    # so it does not depend on the .msc file association.
    try {
        Start-Process -FilePath 'mmc.exe' -ArgumentList 'printmanagement.msc' -ErrorAction Stop
        Write-Log 'Opened printmanagement.msc' 'OK'
    } catch {
        Write-Log "Could not open printmanagement.msc: $($_.Exception.Message)" 'WARN'
        [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 
            "Could not open printmanagement.msc:`n$($_.Exception.Message)`n`nOn Windows client SKUs it may need the 'Print Management' optional feature (RSAT).",
            'Print Management', 'OK', 'Warning')
    }
})

$btnDevPrint.Add_Click({
    # control.exe printers, not ms-settings:printers - the ms-settings URI needs a
    # user session and does nothing from SYSTEM in session 0.
    try { Start-Process -FilePath 'control.exe' -ArgumentList 'printers' -ErrorAction Stop; Write-Log 'Opened Devices and Printers' 'OK' }
    catch { Write-Log "Could not open Devices and Printers: $($_.Exception.Message)" 'WARN' }
})

# ---- Diagnose ----
$btnRunDiag.Add_Click({ $txtFinding.Text = ''; Start-Worker 'diag' })

$dgvFindings.Add_SelectionChanged({
    if ($dgvFindings.SelectedRows.Count -eq 0) { return }
    $i = $dgvFindings.SelectedRows[0].Index
    if ($i -lt 0 -or $i -ge @($script:Findings).Count) { return }
    $f = @($script:Findings)[$i]
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("$($f.Code)  [$($f.Severity)]  scope: $($f.Scope)")
    [void]$sb.AppendLine('')
    if ($f.User)    { [void]$sb.AppendLine("User    : $($f.User)") }
    if ($f.Printer) { [void]$sb.AppendLine("Printer : $($f.Printer)") }
    if ($f.Server)  { [void]$sb.AppendLine("Server  : $($f.Server)") }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('WHAT WAS FOUND')
    [void]$sb.AppendLine("  $($f.Detail)")
    if ($f.LikelyErrorCode) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('USUALLY SEEN AS')
        [void]$sb.AppendLine("  $($f.LikelyErrorCode)")
        foreach ($code in ($f.LikelyErrorCode -split '\s*/\s*')) {
            foreach ($h in @(Get-PrinterErrorCodeHelp -Code $code.Trim())) {
                [void]$sb.AppendLine("    $($h.Code)  $($h.Title)")
                [void]$sb.AppendLine("      cause: $($h.Cause)")
                [void]$sb.AppendLine("      fix  : $($h.Fix)")
            }
        }
    }
    if ($f.Fix) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('FIX')
        [void]$sb.AppendLine("  $($f.Fix)")
    }
    $txtFinding.Text = $sb.ToString()
})

$btnCopyDiag.Add_Click({
    if (@($script:Findings).Count -eq 0) { Write-Log 'Nothing to copy - run diagnostics first.' 'WARN'; return }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Printer diagnostics - $($script:Ctx.Host) - $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
    [void]$sb.AppendLine('')
    foreach ($f in @($script:Findings)) {
        [void]$sb.AppendLine("[$($f.Severity)] $($f.Code)  $($f.User) $($f.Printer)")
        [void]$sb.AppendLine("    $($f.Detail)")
        if ($f.LikelyErrorCode) { [void]$sb.AppendLine("    seen as: $($f.LikelyErrorCode)") }
        if ($f.Fix)             { [void]$sb.AppendLine("    fix    : $($f.Fix)") }
        [void]$sb.AppendLine('')
    }
    try { Set-Clipboard -Value $sb.ToString(); Write-Log 'Report copied to the clipboard.' 'OK' }
    catch { Write-Log "Clipboard unavailable: $($_.Exception.Message)" 'WARN' }
})

# ---- Spooler ----
$btnSpoolRefresh.Add_Click({ Start-Worker 'scan' })

$btnSpoolRestart.Add_Click({
    if ((Show-Msg -Text 'Restart the Print Spooler? Any in-flight jobs will be lost.' -Title 'Confirm restart' -Buttons 'YesNo' -Icon 'Warning' -Default 'No') -ne 'Yes') { return }
    Set-Busy $true 'Restarting spooler...'
    try { Show-Result (Restart-PrintSpooler) 'Spooler' } finally { Set-Busy $false }
    Start-Worker 'scan'
})

$btnSpoolClear.Add_Click({
    $h = $script:SpoolerHealth
    $n = if ($h) { $h.SpoolFileCount } else { '?' }
    $dir = if ($h) { $h.SpoolDirectory } else { '(unknown)' }
    if ((Show-Msg -Text "Stop the Spooler, delete ALL $n spool file(s) in`n$dir`nand start it again?`n`nEvery queued job on this machine will be discarded." -Title 'Confirm spool purge' -Buttons 'YesNo' -Icon 'Warning' -Default 'No') -ne 'Yes') { return }
    Set-Busy $true 'Clearing spool files...'
    try { Show-Result (Clear-SpoolFiles) 'Spool files' } finally { Set-Busy $false }
    Start-Worker 'scan'
})

$btnCancelJob.Add_Click({
    if ($dgvJobs.SelectedRows.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 'Select a job first.', 'Printer Manager', 'OK', 'Information'); return }
    $i = $dgvJobs.SelectedRows[0].Index
    if ($i -lt 0 -or $i -ge @($script:Jobs).Count) { return }
    $j = @($script:Jobs)[$i]
    if ((Show-Msg -Text "Cancel job $($j.Id) '$($j.Document)' on $($j.Queue)?" -Title 'Confirm cancel' -Buttons 'YesNo' -Icon 'Warning' -Default 'No') -ne 'Yes') { return }
    Set-Busy $true 'Cancelling job...'
    try {
        Remove-PrintJob -PrinterName $j.Queue -ID $j.Id -ErrorAction Stop
        Write-Log "Cancelled job $($j.Id) on $($j.Queue)." 'OK'
    } catch {
        Write-Log "Could not cancel job $($j.Id): $($_.Exception.Message)" 'WARN'
    } finally { Set-Busy $false }
    Start-Worker 'scan'
})

$btnCancelAll.Add_Click({
    if ($dgvJobs.SelectedRows.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 'Select a job on the queue you want to clear.', 'Printer Manager', 'OK', 'Information'); return }
    $i = $dgvJobs.SelectedRows[0].Index
    if ($i -lt 0 -or $i -ge @($script:Jobs).Count) { return }
    $q = @($script:Jobs)[$i].Queue
    if ((Show-Msg -Text "Cancel ALL jobs on '$q'?" -Title 'Confirm cancel all' -Buttons 'YesNo' -Icon 'Warning' -Default 'No') -ne 'Yes') { return }
    Set-Busy $true 'Cancelling jobs...'
    try { Show-Result (Clear-PrinterQueue -PrinterName $q) 'Queue cleared' } finally { Set-Busy $false }
    Start-Worker 'scan'
})

# ---- Print server ----
$btnDiscover.Add_Click({ Start-Worker 'discover' })

$btnEnumerate.Add_Click({
    $s = $txtServer.Text.Trim()
    if (-not $s) { [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 'Enter a print server name, or press Discover.', 'Printer Manager', 'OK', 'Information'); return }
    Start-Worker 'server' $s
})

$btnPasteJson.Add_Click({
    $json = Show-InputBox -Prompt 'Paste the JSON from Get-PrinterDiagnostics or Export_Print_Server_Printers_To_JSON (markers optional):' -Title 'Paste collector JSON'
    if ([string]::IsNullOrWhiteSpace($json)) { return }
    # Assign FIRST, then wrap: @($raw | ConvertFrom-Json) collapses a JSON array
    # into a single object on PS 5.1.
    $data = ConvertFrom-NoisyJson $json
    if (-not $data) { [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 'That did not parse as JSON.', 'Paste JSON', 'OK', 'Warning'); return }

    $rows = @()
    if ($data.Printers) {
        # Export_Print_Server_Printers_To_JSON envelope
        $src = "$($data.Source)"
        $rows = @($data.Printers | ForEach-Object {
            [pscustomobject]@{
                ShareName = $_.ShareName; Name = $_.Name; DriverName = $_.DriverName
                PortName = $_.PortName; HostAddress = $_.PortAddress
                # Unshared queues have no mappable UNC - leave it blank rather than
                # showing "\\server\", which looks valid but can never be connected.
                Unc = $(if ($_.ShareName) { "\\$src\$($_.ShareName)" } else { '' })
            }
        })
        $lblServerNote.Text = "Loaded $($rows.Count) queue(s) from $src (exported $($data.GeneratedUtc), method $($data.Method))."
    } elseif ($data.PrintServers) {
        # Get-PrinterDiagnostics envelope
        $rows = @($data.PrintServers | ForEach-Object { $_.Printers } | Where-Object { $_ })
        $lblServerNote.Text = "Loaded $($rows.Count) queue(s) from the $($data.Source) diagnostics export ($($data.GeneratedUtc))."
    } else {
        [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 'That JSON has no Printers or PrintServers section.', 'Paste JSON', 'OK', 'Warning')
        return
    }
    $script:ServerRows = $rows
    Set-Rows $dgvServer $script:ServerRows @('ShareName', 'Name', 'DriverName', 'PortName', 'HostAddress', 'Unc')
    # Nothing preselected: Add and Test must act on a row the technician chose, not on
    # whatever the grid selected by itself when the rows were added.
    $dgvServer.ClearSelection()
    Write-Log "Loaded $($rows.Count) queue(s) from pasted JSON." 'OK'
})



# --------------------------------------------------------------------------
# Repoint / Rebuild dialog
#
# Fixes the two independent defects behind most "printer stopped working"
# calls: a WSD/IPP port that cannot survive a move, and the generic Microsoft
# IPP Class Driver. See Convert-PrinterQueue for the two rules that govern it.
#
# Preview runs INLINE (read-only, bounded probes, sub-second). Apply is handed
# to the background worker, because staging a driver and Add-Printer are slow
# and a wedged spooler call must never freeze the console.
# --------------------------------------------------------------------------
function Show-RepointDialog {
    param($Row)

    $raw      = $Row.Raw
    $curPort  = if ($raw) { [string]$raw.PortName }   else { [string]$Row.Port }
    $curDrv   = if ($raw) { [string]$raw.DriverName } else { [string]$Row.Driver }
    $kind     = try { Get-QueuePortKind -PortName $curPort } catch { 'Unknown' }
    $isIpp    = try { Test-IsIppClassDriver -DriverName $curDrv } catch { $false }

    $f = New-Object System.Windows.Forms.Form
    $f.Text = "Repoint / Rebuild  -  $($Row.Name)"
    $f.Size = New-Object System.Drawing.Size(780, 640)
    $f.MinimumSize = New-Object System.Drawing.Size(700, 560)
    $f.StartPosition = 'CenterParent'
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    function DL([string]$t, [int]$x, [int]$y, [bool]$bold = $false) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $t; $l.Location = New-Object System.Drawing.Point($x, $y); $l.AutoSize = $true
        if ($bold) { $l.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold) }
        $f.Controls.Add($l); $l
    }

    DL 'Current:' 12 14 $true | Out-Null
    DL ("{0}  on  {1}   [{2}]" -f $curDrv, $curPort, $kind) 80 14 | Out-Null

    # ---- Rule 2 banner ----------------------------------------------------
    $lblRule2 = New-Object System.Windows.Forms.Label
    $lblRule2.Location = New-Object System.Drawing.Point(12, 38)
    $lblRule2.Size = New-Object System.Drawing.Size(740, 34)
    $lblRule2.Anchor = [System.Windows.Forms.AnchorStyles]'Top,Left,Right'
    $lblRule2.ForeColor = [System.Drawing.Color]::Firebrick
    $lblRule2.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
    if ($isIpp) {
        $lblRule2.Text = "This queue uses the generic Microsoft IPP Class Driver. It negotiates its page language over IPP, so on a bare RAW:9100 port it often prints BLANK or garbled. Pick the real vendor driver below, or install one, rather than only changing the port."
    } else {
        $lblRule2.Text = ''
    }
    $f.Controls.Add($lblRule2)

    # ---- Address -----------------------------------------------------------
    DL 'Device IP:' 12 84 $true | Out-Null
    $txtIp = New-Object System.Windows.Forms.TextBox
    $txtIp.Location = New-Object System.Drawing.Point(90, 81)
    $txtIp.Size = New-Object System.Drawing.Size(160, 24)
    $f.Controls.Add($txtIp)

    $btnFind = New-Object System.Windows.Forms.Button
    $btnFind.Text = 'Find it'; $btnFind.Location = New-Object System.Drawing.Point(258, 80)
    $btnFind.Size = New-Object System.Drawing.Size(80, 26)
    $f.Controls.Add($btnFind)

    $cboCand = New-Object System.Windows.Forms.ComboBox
    $cboCand.Location = New-Object System.Drawing.Point(346, 81)
    $cboCand.Size = New-Object System.Drawing.Size(406, 24)
    $cboCand.DropDownStyle = 'DropDownList'
    $cboCand.Anchor = [System.Windows.Forms.AnchorStyles]'Top,Left,Right'
    $f.Controls.Add($cboCand)

    $lblIpNote = DL '' 90 110
    $lblIpNote.ForeColor = [System.Drawing.Color]::DimGray

    $script:RepointCands = @()
    $fillCands = {
        param($cands)
        $script:RepointCands = @($cands)
        $cboCand.Items.Clear()
        foreach ($c in $script:RepointCands) {
            [void]$cboCand.Items.Add(("{0}  -  {1}{2}" -f $c.Address, $c.Source, $(if ($c.Reachable) { "  (answers on $($c.OpenPorts -join '/'))" } else { '  (no answer)' })))
        }
        if ($cboCand.Items.Count) {
            $cboCand.SelectedIndex = 0
            if (-not $txtIp.Text) { $txtIp.Text = $script:RepointCands[0].Address }
            $lblIpNote.Text = "Best guess from: $($script:RepointCands[0].Source). Confirm before applying."
        } else {
            $lblIpNote.Text = 'No address could be worked out - type the printer IP from its control panel.'
        }
    }

    $cboCand.Add_SelectedIndexChanged({
        $i = $cboCand.SelectedIndex
        if ($i -ge 0 -and $i -lt @($script:RepointCands).Count) {
            $txtIp.Text = @($script:RepointCands)[$i].Address
            $lblIpNote.Text = "From: $(@($script:RepointCands)[$i].Source). Confirm before applying."
        }
    })

    $btnFind.Add_Click({
        $f.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $btnFind.Enabled = $false
        $lblIpNote.Text = 'Searching: WSD records, ARP sweep on 9100, mDNS...'
        [System.Windows.Forms.Application]::DoEvents()
        try { & $fillCands (Resolve-PrinterAddress -Printer $raw -Active) }
        catch { $lblIpNote.Text = "Search failed: $($_.Exception.Message)" }
        finally { $btnFind.Enabled = $true; $f.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    # ---- Protocol ----------------------------------------------------------
    DL 'Protocol:' 12 142 $true | Out-Null
    $rbRaw = New-Object System.Windows.Forms.RadioButton
    $rbRaw.Text = 'RAW (port'; $rbRaw.Location = New-Object System.Drawing.Point(90, 140)
    $rbRaw.Size = New-Object System.Drawing.Size(78, 24); $rbRaw.Checked = $true
    $f.Controls.Add($rbRaw)
    $numPort = New-Object System.Windows.Forms.NumericUpDown
    $numPort.Location = New-Object System.Drawing.Point(168, 141)
    $numPort.Size = New-Object System.Drawing.Size(66, 24)
    $numPort.Minimum = 1; $numPort.Maximum = 65535; $numPort.Value = 9100
    $f.Controls.Add($numPort)
    DL ')' 238 143 | Out-Null

    $rbLpr = New-Object System.Windows.Forms.RadioButton
    $rbLpr.Text = 'LPR (515), queue'; $rbLpr.Location = New-Object System.Drawing.Point(266, 140)
    $rbLpr.Size = New-Object System.Drawing.Size(120, 24)
    $f.Controls.Add($rbLpr)
    $txtLpr = New-Object System.Windows.Forms.TextBox
    $txtLpr.Text = 'lp'; $txtLpr.Location = New-Object System.Drawing.Point(388, 141)
    $txtLpr.Size = New-Object System.Drawing.Size(90, 24)
    $f.Controls.Add($txtLpr)

    $lblProto = DL 'Some devices ship with RAW/9100 disabled and only LPD/IPP enabled - try LPR if RAW does not answer.' 90 168
    $lblProto.ForeColor = [System.Drawing.Color]::DimGray

    # ---- Driver ------------------------------------------------------------
    DL 'Driver:' 12 200 $true | Out-Null
    $cboDrv = New-Object System.Windows.Forms.ComboBox
    $cboDrv.Location = New-Object System.Drawing.Point(90, 197)
    $cboDrv.Size = New-Object System.Drawing.Size(520, 24)
    $cboDrv.DropDownStyle = 'DropDownList'
    $cboDrv.Anchor = [System.Windows.Forms.AnchorStyles]'Top,Left,Right'
    $f.Controls.Add($cboDrv)

    $btnInf = New-Object System.Windows.Forms.Button
    $btnInf.Text = 'Install from INF...'; $btnInf.Location = New-Object System.Drawing.Point(618, 196)
    $btnInf.Size = New-Object System.Drawing.Size(134, 26)
    $btnInf.Anchor = [System.Windows.Forms.AnchorStyles]'Top,Right'
    $f.Controls.Add($btnInf)

    $loadDrivers = {
        $cboDrv.Items.Clear()
        [void]$cboDrv.Items.Add("(keep current)  $curDrv")
        try {
            foreach ($d in (Get-PrinterDriver -ErrorAction SilentlyContinue | Sort-Object Name)) {
                [void]$cboDrv.Items.Add([string]$d.Name)
            }
        } catch {}
        $cboDrv.SelectedIndex = 0
    }
    & $loadDrivers

    $btnInf.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = 'Printer driver INF (*.inf)|*.inf'
        $ofd.Title  = 'Pick the INF inside its original extracted folder'
        if ($ofd.ShowDialog() -ne 'OK') { return }
        $inf = $ofd.FileName
        $name = Show-InputBox -Prompt "Exact driver name to register (must match a model the INF declares, e.g. 'EPSON WF-7830 Series'):" -Title 'Driver name'
        if ([string]::IsNullOrWhiteSpace($name)) { return }
        $f.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $btnInf.Enabled = $false
        try {
            $r = Install-PrinterDriverFromInf -InfPath $inf -DriverName $name.Trim()
            [void](Show-Msg -Text $r.Message -Title $(if ($r.Success) { 'Driver installed' } else { 'Driver not installed' }) -Icon $(if ($r.Success) { 'Information' } else { 'Warning' }))
            if ($r.Success) {
                & $loadDrivers
                $idx = $cboDrv.Items.IndexOf($name.Trim())
                if ($idx -ge 0) { $cboDrv.SelectedIndex = $idx }
            }
        } finally { $btnInf.Enabled = $true; $f.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    # ---- Options -----------------------------------------------------------
    $chkForce = New-Object System.Windows.Forms.CheckBox
    $chkForce.Text = 'Apply even if the device does not answer'
    $chkForce.Location = New-Object System.Drawing.Point(90, 230)
    $chkForce.AutoSize = $true
    $f.Controls.Add($chkForce)

    $chkDef = New-Object System.Windows.Forms.CheckBox
    $chkDef.Text = 'Restore the default-printer flag for affected users'
    $chkDef.Location = New-Object System.Drawing.Point(90, 254)
    $chkDef.AutoSize = $true; $chkDef.Checked = $true
    $f.Controls.Add($chkDef)

    # ---- Steps -------------------------------------------------------------
    DL 'Plan / result:' 12 286 $true | Out-Null
    $txtSteps = New-Object System.Windows.Forms.TextBox
    $txtSteps.Location = New-Object System.Drawing.Point(12, 308)
    $txtSteps.Size = New-Object System.Drawing.Size(740, 232)
    $txtSteps.Anchor = [System.Windows.Forms.AnchorStyles]'Top,Left,Right,Bottom'
    $txtSteps.Multiline = $true; $txtSteps.ReadOnly = $true; $txtSteps.ScrollBars = 'Both'
    $txtSteps.WordWrap = $false
    $txtSteps.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtSteps.BackColor = [System.Drawing.Color]::White
    $f.Controls.Add($txtSteps)

    # ---- Buttons -----------------------------------------------------------
    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Text = 'Preview'; $btnPreview.Location = New-Object System.Drawing.Point(12, 550)
    $btnPreview.Size = New-Object System.Drawing.Size(110, 30)
    $btnPreview.Anchor = [System.Windows.Forms.AnchorStyles]'Left,Bottom'
    $f.Controls.Add($btnPreview)

    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = 'Apply'; $btnApply.Location = New-Object System.Drawing.Point(556, 550)
    $btnApply.Size = New-Object System.Drawing.Size(96, 30)
    $btnApply.Anchor = [System.Windows.Forms.AnchorStyles]'Right,Bottom'
    $f.Controls.Add($btnApply)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'; $btnCancel.Location = New-Object System.Drawing.Point(658, 550)
    $btnCancel.Size = New-Object System.Drawing.Size(94, 30)
    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]'Right,Bottom'
    $btnCancel.DialogResult = 'Cancel'
    $f.Controls.Add($btnCancel)
    $f.CancelButton = $btnCancel

    $getCfg = {
        $drvSel = "$($cboDrv.SelectedItem)"
        $driver = if ($cboDrv.SelectedIndex -le 0) { $null } else { $drvSel }
        @{
            Name         = $Row.Name
            IPAddress    = $txtIp.Text.Trim()
            Protocol     = $(if ($rbLpr.Checked) { 'LPR' } else { 'RAW' })
            PortNumber   = [int]$numPort.Value
            LprQueueName = $txtLpr.Text.Trim()
            DriverName   = $driver
            Force        = [bool]$chkForce.Checked
            RestoreDefaults = [bool]$chkDef.Checked
        }
    }

    $renderSteps = {
        param($res)
        $sb = New-Object System.Text.StringBuilder
        foreach ($s in @($res.Steps)) { [void]$sb.AppendLine(("[{0,-4}] {1}" -f $s.Tag, $s.Text)) }
        $txtSteps.Text = $sb.ToString()
    }

    $btnPreview.Add_Click({
        $cfg = & $getCfg
        if (-not $cfg.IPAddress) { $txtSteps.Text = 'Enter the device IP first (or press Find it).'; return }
        $f.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $btnPreview.Enabled = $false
        try {
            $p = @{ Name=$cfg.Name; IPAddress=$cfg.IPAddress; Protocol=$cfg.Protocol
                    PortNumber=$cfg.PortNumber; LprQueueName=$cfg.LprQueueName
                    Preview=$true; OnPoll=$script:Pump }
            if ($cfg.DriverName) { $p.DriverName = $cfg.DriverName }
            if (-not $cfg.RestoreDefaults) { $p.SkipDefaultRestore = $true }
            & $renderSteps (Convert-PrinterQueue @p)
        } catch {
            $txtSteps.Text = "Preview failed: $($_.Exception.Message)"
        } finally { $btnPreview.Enabled = $true; $f.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $script:RepointConfig = $null
    $btnApply.Add_Click({
        $cfg = & $getCfg
        if (-not $cfg.IPAddress) { $txtSteps.Text = 'Enter the device IP first (or press Find it).'; return }

        $rebuildWarn = if ($kind -in @('Network','TCPIP') -and -not $cfg.DriverName) {
            "The queue will be REPOINTED onto the new port (nothing is destroyed)."
        } else {
            "The queue will be REBUILT: it is removed and recreated.`n`nKept: name, driver, sharing, comment, location$(if ($cfg.RestoreDefaults) { ', and the default-printer flag for affected users' } else { '' }).`nLOST: printer permissions, printing defaults (duplex/colour/tray), installed options, port SNMP settings, and any queued jobs."
        }
        $msg = "Apply to '$($Row.Name)'?`n`nNew port : $($cfg.Protocol) -> $($cfg.IPAddress)$(if ($cfg.Protocol -eq 'RAW') { ':' + $cfg.PortNumber } else { " (queue '" + $cfg.LprQueueName + "')" })`nDriver   : $(if ($cfg.DriverName) { $cfg.DriverName } else { "$curDrv (unchanged)" })`n`n$rebuildWarn`n`nThis affects every user of this machine."
        if ((Show-Msg -Text $msg -Title 'Confirm repoint / rebuild' -Buttons 'YesNo' -Icon 'Warning' -Default 'No') -ne 'Yes') { return }

        $script:RepointConfig = $cfg
        $f.DialogResult = 'OK'
        $f.Close()
    })

    # Prefill from the free (tier 1) sources as the dialog opens.
    $f.Add_Shown({
        try { & $fillCands (Resolve-PrinterAddress -Printer $raw) } catch {}
        $txtSteps.Text = 'Press Preview to see exactly what would change. Nothing happens until you press Apply.'
    })

    $result = $f.ShowDialog()
    $f.Dispose()
    if ($result -eq 'OK') { return $script:RepointConfig }
    return $null
}

function Apply-RepointResults {
    param($sh)
    $res = $sh.Repoint
    if (-not $res) { return }
    $sb = New-Object System.Text.StringBuilder
    foreach ($s in @($res.Steps)) { [void]$sb.AppendLine(("[{0,-4}] {1}" -f $s.Tag, $s.Text)) }
    foreach ($line in ($sb.ToString() -split "`r?`n")) {
        if ("$line".Trim()) { Write-Log $line $(if ($line -like '`[FAIL*') { 'ERROR' } elseif ($line -like '`[WARN*') { 'WARN' } else { 'INFO' }) }
    }
    [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 
        $res.Message + $(if (@($res.Checks).Count) { "`n`nVerification:`n" + ((@($res.Checks) | ForEach-Object { "  $($_.Result)  $($_.Check) - $($_.Detail)" }) -join "`n") } else { '' }),
        $(if ($res.Success) { 'Repoint complete' } else { 'Repoint finished with problems' }),
        'OK', $(if ($res.Success) { 'Information' } else { 'Warning' }))
}

$btnJunk.Add_Click({
    Set-Busy $true 'Looking for junk queues...'
    $cands = @()
    try {
        # Feature state needs elevation to READ, so unelevated the dialog says
        # "unknown" rather than implying the queue will stay gone.
        $cands = @(Get-JunkPrinterCandidate)
    } catch {
        Write-Log ("Could not check for junk queues: " + $_.Exception.Message) 'WARN'
    } finally { Set-Busy $false }

    $did = Show-JunkQueueDialog -Candidates $cands
    Update-JunkLabel
    if ($did) { Start-Worker 'scan' }
})

$btnShareAll.Add_Click({
    # SOURCE-side action: shares a queue that lives on THIS PC. It used to depend on
    # the Printers tab selection; on Share Host the technician picks the queue here.
    $local = @(@($script:PrinterRows) | Where-Object { $_ -and -not $_.IsUserOnly -and "$($_.Name)" -notmatch '^\\\\' })
    if (-not $local.Count) {
        [void](Show-Msg -Text 'No local printer queues were found on this PC to share. Press Rescan first.' -Title 'Share to everyone' -Icon 'Information')
        return
    }
    $r = Show-PickList -Title 'Share to everyone' -Prompt 'Pick the printer on THIS PC to share to the network:' -Items $local -Label { param($x) '{0}   [{1}]' -f $x.Name, $x.Port }
    if (-not $r) { return }
    $share = if ($r.Raw -and $r.Raw.ShareName) { [string]$r.Raw.ShareName } else { "$($r.Name)" }
    $share = Show-InputBox -Title 'Share name' -Default $share -Prompt (
        "Share '$($r.Name)' to everyone on the network.`n`nThis will: share the queue, grant Print to Everyone, " +
        "enable the scoped File and Printer Sharing firewall rules, and move a Public network to Private.`n`n" +
        "It will NOT enable guest access, weaken SMB signing, or touch Point and Print.`n`nShare name:")
    if ([string]::IsNullOrWhiteSpace($share)) { return }

    Set-Busy $true 'Sharing to everyone...'
    try {
        $res = Enable-PrinterShareForAll -Name $r.Name -ShareName $share.Trim() -Trustee Everyone
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine($res.Message)
        [void]$sb.AppendLine('')
        foreach ($s in @($res.Steps)) {
            [void]$sb.AppendLine(("[{0,-4}] {1}" -f $s.Tag, $s.Text))
            Write-Log ("Share to everyone: " + $s.Text) $(if ($s.Tag -eq 'FAIL') { 'ERROR' } elseif ($s.Tag -eq 'WARN') { 'WARN' } else { 'INFO' })
        }
        if (@($res.Warnings).Count) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('STILL IN THE WAY:')
            foreach ($w in @($res.Warnings)) { [void]$sb.AppendLine("  * $w"); Write-Log $w 'WARN' }
        }
        [void](Show-Msg -Text $sb.ToString() -Title $(if ($res.Success) { 'Shared' } else { 'Shared with problems' }) -Icon $(if ($res.Success) { 'Information' } else { 'Warning' }))
    } finally { Set-Busy $false }
    Start-Worker 'sharehost'
})

$btnRepoint.Add_Click({
    $r = Require-Printer; if (-not $r) { return }
    if ($r.IsUserOnly -or "$($r.Name)" -match '^\\') {
        [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 
            "'$($r.Name)' is a connection to a queue shared from a print server.`n`nIts port and driver live on the SERVER, so repoint it there - changing it here would only affect this one machine.",
            'Repoint on the server instead', 'OK', 'Information')
        return
    }
    $cfg = Show-RepointDialog -Row $r
    if (-not $cfg) { return }
    Start-Worker 'repoint' '' $cfg
})



# --------------------------------------------------------------------------
# Print RPC transport
#
# Windows 11 sends print RPC over TCP: port 135 (endpoint mapper), then the spooler's
# own port, dynamic 49152-65535 by default. A client cannot test a dynamic port it
# does not know, so the RECOMMENDED repair (Microsoft) is on the host: pin the spooler
# with RpcTcpPort and allow that port - Share Host -> Fix print RPC (recommended).
#
# Named pipes (host RpcProtocols = 7, client RpcUseNamedPipeProtocol = 1) remain as
# an ADVANCED option only. Microsoft: "This configuration isn't recommended", and it
# may also need RpcAuthnLevelPrivacyEnabled = 0 and insecure guest logons - which
# this console reports but will never write.
# --------------------------------------------------------------------------
function Get-RpcStateLine {
    <#
    .SYNOPSIS
        One short line of LIVE print RPC state for a button band.
    .DESCRIPTION
        Host side shows activation next to configuration: what the spooler is really
        listening on, and whether a pinned RpcTcpPort is in effect. Client side shows
        how this PC connects. Kept short to fit a 1024-wide ScreenConnect window.
    #>
    param([ValidateSet('Host', 'Client')][string]$Side = 'Host')
    $by = @{}
    try { foreach ($v in @((Get-PrintRpcConfig).Values)) { $by[[string]$v.Name] = $v } } catch {}
    if ($Side -eq 'Client') {
        $cp = $by['RpcUseNamedPipeProtocol']
        # Kept short: this sits at x=468 and must fit the 820px minimum window.
        if ($cp -and $cp.IsSet -and "$($cp.Value)" -eq '1') { return 'Named pipes in use (not recommended).' }
        return 'Connects over TCP (Windows default).'
    }
    $tp = $by['RpcTcpPort']
    $ports = @()
    try { $ports = @((Get-PrintSpoolerListener).Ports) } catch {}
    $pinTxt = if ($tp -and $tp.IsSet) { "pinned to $($tp.Value)" } else { 'not pinned' }
    $live = if ($ports.Count) { 'listening on ' + ($ports -join ', ') } else { 'not listening on TCP' }
    $warn = ''
    if ($tp -and $tp.IsSet) {
        $want = 0
        try { $want = [int]$tp.Value } catch {}
        if (-not ($ports -contains $want)) { $warn = ' - NOT IN EFFECT' }
    }
    "Spooler $live; RpcTcpPort $pinTxt$warn"
}
function Update-JunkLabel {
    <# One line on the Printers band saying whether there is anything to clean up. #>
    try {
        # -SkipFeatureCheck: the label must never pay for a DISM query, and the
        # feature state is only needed once the dialog is actually open.
        $n = @(Get-JunkPrinterCandidate -SkipFeatureCheck).Count
        $lblJunk.Text = if ($n) { "$n junk queue(s) on this PC - Fax / XPS / OneNote for Windows 10." }
                        else    { 'No Fax, XPS or OneNote-for-Windows-10 queue on this PC.' }
    } catch { $lblJunk.Text = '' }
}
function Update-RpcStateLabels {
    $lblHostRpc.Text = Get-RpcStateLine -Side 'Host'
    $lblSrvRpc.Text  = Get-RpcStateLine -Side 'Client'
}
function Get-RpcListenerIsOn {
    $by = @{}
    try { foreach ($v in @((Get-PrintRpcConfig).Values)) { $by[[string]$v.Name] = $v } } catch {}
    $lp = $by['RpcProtocols']
    ($lp -and $lp.IsSet -and "$($lp.Value)" -eq '7')
}
function Get-RpcClientIsOn {
    $by = @{}
    try { foreach ($v in @((Get-PrintRpcConfig).Values)) { $by[[string]$v.Name] = $v } } catch {}
    $cp = $by['RpcUseNamedPipeProtocol']
    ($cp -and $cp.IsSet -and "$($cp.Value)" -eq '1')
}

function Invoke-RpcProbeResponsive {
    <#
    .SYNOPSIS
        Get-Printer -ComputerName, bounded, WITHOUT freezing the window.
    .DESCRIPTION
        Test-RemotePrintEnumeration blocks its caller while it waits, which is
        fine in a script and wrong in a dialog - up to 45s of dead window is the
        Backstage hourglass all over again. This runs the same call in its own
        runspace and pumps messages while it waits.

        On timeout the pipeline is ABANDONED with BeginStop, never Stop/Dispose:
        those block until the wedged call ends, which is exactly what will not
        happen when a print RPC is being swallowed by a firewall.
    #>
    param([string]$ComputerName, [int]$TimeoutSec = 20)

    $target = ([string]$ComputerName).Trim().Trim('\')
    if (-not $target) { return [pscustomobject]@{ Ok=$false; Count=0; Names=@(); Error='No host given.'; TimedOut=$false } }

    $rs = [runspacefactory]::CreateRunspace(); $rs.Open()
    $ps = [powershell]::Create(); $ps.Runspace = $rs
    [void]$ps.AddScript({
        param($Srv)
        try {
            $p = @(Get-Printer -ComputerName $Srv -ErrorAction Stop)
            [pscustomobject]@{ Ok=$true; Count=$p.Count; Names=@($p | ForEach-Object { [string]$_.Name }); Error='' }
        } catch {
            [pscustomobject]@{ Ok=$false; Count=0; Names=@(); Error=$_.Exception.Message }
        }
    })
    [void]$ps.AddParameter('Srv', $target)

    $async = $ps.BeginInvoke()
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while (-not $async.IsCompleted -and (Get-Date) -lt $deadline) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    }
    if (-not $async.IsCompleted) {
        try { [void]$ps.BeginStop($null, $null) } catch {}
        return [pscustomobject]@{ Ok=$false; Count=0; Names=@()
            Error="No answer within $TimeoutSec seconds - the RPC call never returned, which is itself the symptom."
            TimedOut=$true }
    }
    $out = $null
    try { $out = $ps.EndInvoke($async) } catch {}
    try { $ps.Dispose(); $rs.Close(); $rs.Dispose() } catch {}
    $r = $null
    if ($out) { $r = @($out)[0] }
    if (-not $r) { return [pscustomobject]@{ Ok=$false; Count=0; Names=@(); Error='No result returned.'; TimedOut=$false } }
    [pscustomobject]@{ Ok=[bool]$r.Ok; Count=[int]$r.Count; Names=@($r.Names); Error=[string]$r.Error; TimedOut=$false }
}

function Show-RpcFixDialog {
    <#
    .SYNOPSIS
        Apply, re-apply, TEST or remove one half of the print RPC fix.
    .DESCRIPTION
        Replaces the old toggle buttons, which were a dead end: once the value
        was set the only thing on offer was to revert it, so the moment you most
        needed to diagnose was the moment the button stopped diagnosing.

        Three things this makes possible that the toggle did not:

          * TEST without changing anything. Listing the host's printers is the
            whole diagnosis, and it must not require reverting a working setting
            first.
          * RE-APPLY. A registry value can read "set" while the fix is NOT in
            effect, because the spooler has not restarted since. Re-applying
            rewrites the value AND restarts the spooler, which is a real repair
            rather than a no-op.
          * See BOTH halves at once, because the commonest field mistake is
            applying the host half on the client.
    #>
    param([ValidateSet('Host','Client')][string]$Role)

    $isHost = ($Role -eq 'Host')

    $f = New-Object System.Windows.Forms.Form
    $f.Text = $(if ($isHost) { 'ADVANCED, not recommended - named-pipe listener (host)' } else { 'ADVANCED, not recommended - named pipes (client)' })
    $f.Size = New-Object System.Drawing.Size(720, 560)
    $f.MinimumSize = New-Object System.Drawing.Size(660, 500)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'Sizable'
    $f.MinimizeBox = $false; $f.MaximizeBox = $false
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $top = New-DockBand $f 'Top' 190
    $bot = New-DockBand $f 'Bottom' 52
    $mid = New-Object System.Windows.Forms.Panel
    $mid.Dock = 'Fill'; $f.Controls.Add($mid); $mid.BringToFront()

    $lblWhich = New-Lbl '' 10 8 $top $true
    # Microsoft: RPC over named pipes "isn't recommended", and may also need
    # RpcAuthnLevelPrivacyEnabled=0 and insecure guest logons - downgrades this console
    # will not make. The recommended path is Share Host -> Fix print RPC (recommended).
    $lblWhich.Text = $(if ($isHost) {
        'NOT RECOMMENDED by Microsoft - named-pipe listener (RpcProtocols = 7) on the PC that SHARES the printer.'
    } else {
        'NOT RECOMMENDED by Microsoft - named pipes (RpcUseNamedPipeProtocol = 1) on the PC that connects.'
    })

    $lblState = New-Lbl '' 10 32 $top
    $lblState.MaximumSize = New-Object System.Drawing.Size(680, 0)
    $lblWarn = New-Lbl '' 10 92 $top
    $lblWarn.MaximumSize = New-Object System.Drawing.Size(680, 0)
    $lblWarn.ForeColor = [System.Drawing.Color]::Firebrick

    New-Lbl 'Print host to test against:' 10 146 $top | Out-Null
    $txtRpcHost = New-Object System.Windows.Forms.TextBox
    $txtRpcHost.Location = New-Object System.Drawing.Point(170, 143)
    $txtRpcHost.Size = New-Object System.Drawing.Size(220, 24)
    $top.Controls.Add($txtRpcHost)
    $seed = ([string]$txtServer.Text).Trim().Trim('\')
    $txtRpcHost.Text = $seed

    $btnTest = New-Btn 'Test now' 400 142 100 27 $top
    $lblTestNote = New-Lbl 'changes nothing' 508 147 $top
    $lblTestNote.ForeColor = [System.Drawing.Color]::DimGray

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Multiline = $true; $txtOut.ReadOnly = $true
    $txtOut.ScrollBars = 'Both'; $txtOut.WordWrap = $false
    $txtOut.Dock = 'Fill'
    $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtOut.BackColor = [System.Drawing.Color]::White
    $mid.Controls.Add($txtOut)

    $btnApply  = New-Btn 'Apply / re-apply' 10 12 140 27 $bot
    $btnRemove = New-Btn 'Remove setting' 156 12 130 27 $bot
    $btnClose  = New-Btn 'Close' 292 12 90 27 $bot
    $lblBusy   = New-Lbl '' 392 18 $bot
    $lblBusy.ForeColor = [System.Drawing.Color]::DimGray

    $script:RpcDialogChanged = $false

    $say = {
        param([string]$Text)
        $txtOut.AppendText($Text + [Environment]::NewLine)
        $txtOut.SelectionStart = $txtOut.TextLength
        $txtOut.ScrollToCaret()
    }

    $refresh = {
        $by = @{}
        try { foreach ($v in @((Get-PrintRpcConfig).Values)) { $by[[string]$v.Name] = $v } } catch {}
        $lp = $by['RpcProtocols']; $cp = $by['RpcUseNamedPipeProtocol']
        $lpSet = ($lp -and $lp.IsSet); $cpSet = ($cp -and $cp.IsSet)

        $lines = [System.Collections.Generic.List[string]]::new()
        $lines.Add("This PC is $env:COMPUTERNAME.")
        $lines.Add(("   listener  RpcProtocols            = {0}" -f $(if ($lpSet) { "$($lp.Value)" } else { 'not set (Windows default)' })))
        $lines.Add(("   client    RpcUseNamedPipeProtocol = {0}" -f $(if ($cpSet) { "$($cp.Value)" } else { 'not set (0 = over TCP)' })))
        $lblState.Text = ($lines -join [Environment]::NewLine)

        # The commonest mistake in the field, and invisible unless it is said.
        if ($lpSet -and $cpSet) {
            $lblWarn.Text = 'BOTH halves are set on THIS machine. The listener half belongs on the PC that SHARES the printer. If that is a different PC, this one carries a setting it does not need and the real host may still be missing it - check there.'
        } else { $lblWarn.Text = '' }

        $mySet = if ($isHost) { $lpSet } else { $cpSet }
        $btnRemove.Enabled = $mySet
        $btnApply.Text = if ($mySet) { 'Re-apply + restart' } else { 'Apply' }
        Update-RpcStateLabels
    }

    $runTest = {
        $h = ([string]$txtRpcHost.Text).Trim().Trim('\')
        if (-not $h) { & $say 'Enter a print host first - listing its printers IS the test.'; return }
        $btnTest.Enabled = $false; $btnApply.Enabled = $false; $btnRemove.Enabled = $false
        $lblBusy.Text = "Listing printers on $h ..."
        $f.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            & $say ''
            & $say ("--- {0}  test against {1} ---" -f (Get-Date -Format 'HH:mm:ss'), $h)
            $r = Invoke-RpcProbeResponsive -ComputerName $h -TimeoutSec 20
            if ($r.Ok) {
                & $say ("REMOTE ENUMERATION WORKS - {0} printer(s) listed on {1}." -f $r.Count, $h)
                foreach ($n in @($r.Names)) { & $say ("     " + $n) }
                & $say ''
                & $say 'This console could list that host''s printers. That is context only: it does not'
                & $say 'show that a particular printer opens for the target user, and it is not a test of'
                & $say 'the print RPC path. Use Shares -> Test connection for that.'
            } else {
                & $say ("REMOTE ENUMERATION FAILS: " + $r.Error)
                & $say ''
                & $say 'Listing failed. On its own that does not locate the fault.'
                & $say 'Prefer the recommended path: on the host, Share Host -> Fix print RPC'
                & $say '(pin the spooler port, enable the built-in rules), then from the client'
                & $say 'Shares -> Test connection with that port.'
            }
        } catch {
            & $say ("Test failed: " + $_.Exception.Message)
        } finally {
            $f.Cursor = [System.Windows.Forms.Cursors]::Default
            $lblBusy.Text = ''
            $btnTest.Enabled = $true; $btnApply.Enabled = $true
            & $refresh
        }
    }

    $btnTest.Add_Click($runTest)

    $btnApply.Add_Click({
        $btnTest.Enabled = $false; $btnApply.Enabled = $false; $btnRemove.Enabled = $false
        $lblBusy.Text = 'Applying and restarting the spooler...'
        $f.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            & $say ''
            & $say ("--- {0}  apply ---" -f (Get-Date -Format 'HH:mm:ss'))
            $r = if ($isHost) { Set-PrintRpcListenerProtocol } else { Set-PrintRpcNamedPipeFallback }
            & $say $r.Message
            Write-Log $r.Message $(if ($r.Success) { 'OK' } else { 'WARN' })
            if ($r.Success) { $script:RpcDialogChanged = $true }
        } catch {
            & $say ("Apply failed: " + $_.Exception.Message)
        } finally {
            $f.Cursor = [System.Windows.Forms.Cursors]::Default
            $lblBusy.Text = ''
            $btnTest.Enabled = $true; $btnApply.Enabled = $true
            & $refresh
        }
        # The point of applying is to change the answer, so re-test straight away.
        if (([string]$txtRpcHost.Text).Trim()) { & $runTest }
    })

    $btnRemove.Add_Click({
        $what = if ($isHost) { 'RpcProtocols (the listener half)' } else { 'RpcUseNamedPipeProtocol (the client half)' }
        $q = "Remove $what from this PC and return it to the Windows default?" + [Environment]::NewLine + [Environment]::NewLine + 'The spooler will be restarted.'
        if ((Show-Msg -Text $q -Title 'Remove the setting?' -Buttons 'YesNo' -Icon 'Question' -Default 'No') -ne 'Yes') { return }
        $btnTest.Enabled = $false; $btnApply.Enabled = $false; $btnRemove.Enabled = $false
        $lblBusy.Text = 'Removing and restarting the spooler...'
        $f.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            & $say ''
            & $say ("--- {0}  remove ---" -f (Get-Date -Format 'HH:mm:ss'))
            $r = if ($isHost) { Set-PrintRpcListenerProtocol -Reset } else { Set-PrintRpcNamedPipeFallback -Off }
            & $say $r.Message
            Write-Log $r.Message $(if ($r.Success) { 'OK' } else { 'WARN' })
            if ($r.Success) { $script:RpcDialogChanged = $true }
        } catch {
            & $say ("Remove failed: " + $_.Exception.Message)
        } finally {
            $f.Cursor = [System.Windows.Forms.Cursors]::Default
            $lblBusy.Text = ''
            $btnTest.Enabled = $true; $btnApply.Enabled = $true
            & $refresh
        }
    })

    $btnClose.Add_Click({ $f.Close() })

    $f.Add_Shown({
        & $refresh
        & $say 'Named pipes only work when BOTH sides are set:'
        & $say '   HOST   (shares the printer) : RpcProtocols            = 7'
        & $say '   CLIENT (connects)           : RpcUseNamedPipeProtocol = 1'
        & $say ''
        & $say 'Microsoft does not recommend RPC over named pipes. Prefer Share Host -> Fix print'
        & $say 'RPC (recommended) on the host, then Shares -> Test connection on the client.'
        & $say 'Test now only lists printers from this console - context, not proof.'
    })
    [void]$f.ShowDialog((Get-MsgOwner))
    $f.Dispose()
    return $script:RpcDialogChanged
}

$btnHostRpc.Add_Click({ [void](Show-RpcFixDialog -Role 'Host');   Update-RpcStateLabels })
$btnSrvRpc.Add_Click({  [void](Show-RpcFixDialog -Role 'Client'); Update-RpcStateLabels })


# --------------------------------------------------------------------------
# Share Host tab + "add a share to this PC"
# --------------------------------------------------------------------------
# --------------------------------------------------------------------------
# v2.0 - shared printers, source vs target
#
# Shares (TARGET) and Share Host (SOURCE) each act only on the PC they run on.
# Anything slow runs in the background worker so the window stays usable; inside a
# modal dialog, where the worker's results cannot be applied, Invoke-ToolkitResponsive
# does the same job. Remedies are chosen from a measured diagnosis category, never
# from message text, and named pipes are never offered automatically.
# --------------------------------------------------------------------------
function Show-PickList {
    <# .SYNOPSIS Pick one item from a list. Returns the chosen item, or $null. #>
    param([string]$Title, [string]$Prompt, [object[]]$Items, [scriptblock]$Label)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Title
    $f.Size = New-Object System.Drawing.Size(640, 420)
    $f.MinimumSize = New-Object System.Drawing.Size(460, 300)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'Sizable'
    $f.MinimizeBox = $false; $f.MaximizeBox = $false; $f.ShowInTaskbar = $false
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $top = New-DockBand $f 'Top' 46
    $bot = New-DockBand $f 'Bottom' 46
    $lp = New-Lbl $Prompt 10 8 $top
    $lp.MaximumSize = New-Object System.Drawing.Size(600, 36)
    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Dock = 'Fill'
    $lst.IntegralHeight = $false
    $f.Controls.Add($lst)
    $lst.BringToFront()
    $pickItems = @($Items)
    foreach ($it in $pickItems) { [void]$lst.Items.Add($(if ($Label) { [string](& $Label $it) } else { [string]$it })) }
    $bOk = New-Btn 'OK' 10 9 100 27 $bot
    $bCancel = New-Btn 'Cancel' 116 9 100 27 $bot
    $script:PickListResult = $null
    $bOk.Add_Click({ if ($lst.SelectedIndex -ge 0) { $script:PickListResult = $pickItems[$lst.SelectedIndex]; $f.Close() } })
    $lst.Add_DoubleClick({ if ($lst.SelectedIndex -ge 0) { $script:PickListResult = $pickItems[$lst.SelectedIndex]; $f.Close() } })
    $bCancel.Add_Click({ $f.Close() })
    $f.AcceptButton = $bOk
    $f.CancelButton = $bCancel
    $f.Add_Shown({ if ($lst.Items.Count) { $lst.SelectedIndex = 0 } })
    [void]$f.ShowDialog((Get-MsgOwner))
    $f.Dispose()
    return $script:PickListResult
}

function Invoke-ToolkitResponsive {
    <#
    .SYNOPSIS
        Run one toolkit command in a private runspace while a dialog stays responsive.
    .DESCRIPTION
        Inside a modal dialog the background worker cannot be used: its results are only
        applied while no dialog is open. This dot-sources the same toolkit source into a
        private runspace, pumps messages while it waits, and on timeout abandons the
        pipeline with BeginStop rather than blocking on it.
    .OUTPUTS
        PSCustomObject { TimedOut; Result; Error }
    #>
    param(
        [Parameter(Mandatory)][string]$Command,
        [hashtable]$Arguments = @{},
        [int]$TimeoutSec = 90
    )
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        param($ToolkitSource, $Command, $Arguments)
        . ([scriptblock]::Create($ToolkitSource))
        & $Command @Arguments
    }.ToString())
    [void]$ps.AddParameter('ToolkitSource', $script:ToolkitSource)
    [void]$ps.AddParameter('Command', $Command)
    [void]$ps.AddParameter('Arguments', $Arguments)
    $async = $ps.BeginInvoke()
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while (-not $async.IsCompleted -and (Get-Date) -lt $deadline) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    }
    if (-not $async.IsCompleted) {
        try { [void]$ps.BeginStop($null, $null) } catch {}
        return [pscustomobject]@{ TimedOut = $true; Result = $null; Error = "Did not finish within $TimeoutSec seconds." }
    }
    $out = $null; $err = ''
    try { $out = $ps.EndInvoke($async) } catch { $err = $_.Exception.Message }
    if (-not $err -and $ps.Streams.Error.Count) { $err = [string]$ps.Streams.Error[0] }
    try { $ps.Dispose(); $rs.Close(); $rs.Dispose() } catch {}
    $res = $null
    if ($out -and $out.Count) { $res = $out[$out.Count - 1] }
    [pscustomobject]@{ TimedOut = $false; Result = $res; Error = $err }
}

function Get-SelectedServerUnc {
    <# .SYNOPSIS The UNC of the selected Shares row, or one typed in. '' if cancelled. #>
    param([string]$Title = 'Shared printer')
    if ($dgvServer.SelectedRows.Count -gt 0) {
        $i = $dgvServer.SelectedRows[0].Index
        if ($i -ge 0 -and $i -lt @($script:ServerRows).Count) {
            $u = [string]@($script:ServerRows)[$i].Unc
            if ($u) { return $u.Trim() }
        }
    }
    $typed = Show-InputBox -Prompt 'Shared printer, as \\host\sharename :' -Title $Title -Default '\\'
    if ([string]::IsNullOrWhiteSpace($typed)) { return '' }
    $typed = $typed.Trim()
    if ($typed -eq '\\') { return '' }
    return $typed
}

function Start-EvidenceTest {
    <#
    .SYNOPSIS
        Test the EXACT connection, in the background: path, live open as console, live open as the target user.
    #>
    param([string]$Unc)
    $u = ([string]$Unc).Trim()
    if ($u -notmatch '^\\\\[^\\]+\\.+') { Write-Log "'$u' is not a \\host\sharename path." 'WARN'; return }
    $portTxt = Show-InputBox -Title 'Pinned spooler port (optional)' -Default '' -Prompt (
        "If the host's spooler has been pinned to a TCP port (Share Host -> Fix print RPC on that PC), enter it and that exact port is tested as well.`n`nLeave blank if it has not been pinned.")
    if ($null -eq $portTxt) { return }
    $port = 0
    if ($portTxt.Trim()) {
        if (-not [int]::TryParse($portTxt.Trim(), [ref]$port) -or $port -lt 1 -or $port -gt 65535) {
            Write-Log "'$portTxt' is not a port number." 'WARN'
            return
        }
    }
    Start-Worker 'evidence' '' ([pscustomobject]@{ Unc = $u; User = [string]$script:TargetUser; PinnedPort = $port })
}

function Invoke-DiagnosisRemedy {
    <#
    .SYNOPSIS
        Offer the ONE next step that fits a measured diagnosis category.
    .DESCRIPTION
        Chosen from Diagnosis.Category, never from message text. Transport and Unknown
        get no change on this PC at all: the recommended repair is on the host, and the
        right step here is to measure. Named pipes are never offered from here.
    #>
    param($Diagnosis, [string]$Unc)
    if (-not $Diagnosis) { return }
    $hostName = [string]$Diagnosis.Host
    switch ([string]$Diagnosis.Category) {
        'None' { Write-Log "No fault was measured for $Unc." 'OK' }
        'Identity' {
            if ((Show-Msg -Text "The cause is IDENTITY ($($Diagnosis.Confidence)): $($Diagnosis.Why)`r`n`r`nOpen the Credentials dialog for $hostName now?" -Title 'Fix credentials?' -Buttons 'YesNo' -Icon 'Question') -eq 'Yes') {
                Show-CredentialsDialog -Server $hostName
            }
        }
        'Driver' {
            if ((Show-Msg -Text "The cause is the DRIVER ($($Diagnosis.Confidence)): $($Diagnosis.Why)`r`n`r`nInstall the vendor driver on this PC now? Then Add... again with the local-port scope, and run the delivery test - a created queue is not a printed page." -Title 'Install a driver?' -Buttons 'YesNo' -Icon 'Question') -eq 'Yes') {
                $installed = Show-InstallDriverDialog
                if ($installed) {
                    $cboAddScope.SelectedIndex = 2
                    Write-Log "Driver '$installed' installed. Scope switched to local port - select the queue and press Add... again." 'OK'
                }
            }
        }
        default {
            $ask = "Nothing on this PC should be changed - the diagnosis is $($Diagnosis.Category) ($($Diagnosis.Confidence)).`r`n`r`n" +
                   "Run Test connection against $Unc now? If $hostName's spooler has been pinned to a port you will be asked for it, which makes the print RPC path measurable."
            if ((Show-Msg -Text $ask -Title 'Test the connection?' -Buttons 'YesNo' -Icon 'Question') -eq 'Yes') {
                Start-EvidenceTest -Unc $Unc
            }
        }
    }
}

function Apply-AddConnResults {
    param($sh)
    $a = $sh.AddConn
    if (-not $a) { return }
    $p = $a.Payload
    $res = $a.Result
    if (-not $res) { Write-Log 'The add did not return a result.' 'ERROR'; return }
    Show-Result $res $(if ($p.Scope -eq 'LocalPort') { 'Local queue created' } else { 'Printer added' })
    if ($res.Success) {
        if ($p.Scope -eq 'Machine') {
            Write-Log 'Per-computer connection registered. Users see it at their next sign-in, so a delivery test from here is not possible yet.' 'INFO'
        } else {
            $isLocal = ($p.Scope -eq 'LocalPort')
            $pname = if ($isLocal) { [string]$res.Name } else { [string]$p.Unc }
            $who = if ($isLocal) { 'from this PC' } else { "as $($p.User)" }
            $ask = "Optional: send a test page to '$pname' $who and watch the job leave the queue?`r`n`r`nCreated is not the same as printing. This shows whether the job is handed off from this PC; it cannot prove a page came out."
            if ((Show-Msg -Text $ask -Title 'Delivery test?' -Buttons 'YesNo' -Icon 'Question' -Default 'No') -eq 'Yes') {
                $script:PendingDelivery = [pscustomobject]@{ PrinterName = $pname; User = [string]$p.User; InProcess = $isLocal }
            }
        }
    } elseif ($res.PSObject.Properties.Name -contains 'Diagnosis' -and $res.Diagnosis) {
        $script:PendingRemedy = [pscustomobject]@{ Diagnosis = $res.Diagnosis; Unc = [string]$p.Unc }
    }
    $script:PendingScan = $true
}

function Apply-EvidenceResults {
    param($sh)
    $e = $sh.Evidence
    if (-not $e) { return }
    $d = $e.Diagnosis
    $lvl = if ($d.Category -eq 'None') { 'OK' } else { 'WARN' }
    Write-Log ("Connection test for $($e.Unc): $($d.Category) ($($d.Confidence)) - $($d.Why)") $lvl
    [void](Show-Msg -Text $e.Text -Title "Connection test - $($e.Unc)" -Icon $(if ($d.Category -eq 'None') { 'Information' } else { 'Warning' }))
}

function Apply-DeliveryResults {
    param($sh)
    $r = $sh.Delivery
    if (-not $r) { return }
    Write-Log $r.Message $(if ($r.Success) { 'OK' } else { 'WARN' })
    [void](Show-Msg -Text $r.Message -Title $(if ($r.Success) { 'Delivery test' } else { 'Delivery test - not handed off' }) -Icon $(if ($r.Success) { 'Information' } else { 'Warning' }))
    $script:PendingScan = $true
}

function Show-RpcPortDialog {
    <#
    .SYNOPSIS
        HOST: the recommended print RPC repair - pin the spooler port, allow it, prove it.
    .DESCRIPTION
        Run on the PC that SHARES the printer. Microsoft's recommended configuration keeps
        RPC over TCP:
          1. pin the spooler to a chosen port, and prove it is listening there
          2. enable the scoped built-in firewall rules
          then, on the client, Shares -> Test connection with that port
          3. only if that measured the port as closed, create one narrow rule for it
        Restore puts back exactly what a snapshot recorded. Every change runs off the UI
        thread, so the dialog stays usable.
    #>
    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Fix print RPC - host (recommended)'
    $f.Size = New-Object System.Drawing.Size(790, 600)
    $f.MinimumSize = New-Object System.Drawing.Size(700, 520)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'Sizable'
    $f.MinimizeBox = $false; $f.MaximizeBox = $true; $f.ShowInTaskbar = $false
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $top = New-DockBand $f 'Top' 142
    $bot = New-DockBand $f 'Bottom' 44
    $mid = New-Object System.Windows.Forms.Panel
    $mid.Dock = 'Fill'; $f.Controls.Add($mid); $mid.BringToFront()

    New-Lbl 'Run this on the PC that SHARES the printer.' 10 8 $top $true | Out-Null
    $lblWhy = New-Lbl "Microsoft's recommended fix keeps print RPC over TCP: pin the spooler to one port, allow it, then test that port from the client." 10 28 $top
    $lblWhy.MaximumSize = New-Object System.Drawing.Size(750, 0)
    $lblWhy.ForeColor = [System.Drawing.Color]::DimGray

    New-Lbl 'Port:' 10 66 $top | Out-Null
    $txtPort = New-Object System.Windows.Forms.TextBox
    $txtPort.Location = New-Object System.Drawing.Point(48, 63)
    $txtPort.Size = New-Object System.Drawing.Size(70, 24)
    $top.Controls.Add($txtPort)
    $btnPin     = New-Btn '1. Pin spooler to port'    126 61 170 27 $top
    $btnRules   = New-Btn '2. Enable built-in rules'  302 61 170 27 $top
    $btnRestore = New-Btn 'Restore previous...'       478 61 140 27 $top
    $btnRefresh = New-Btn 'Refresh'                   624 61 90  27 $top

    New-Lbl 'Narrow rule scope:' 10 106 $top | Out-Null
    $txtScope = New-Object System.Windows.Forms.TextBox
    $txtScope.Location = New-Object System.Drawing.Point(126, 103)
    $txtScope.Size = New-Object System.Drawing.Size(170, 24)
    $txtScope.Text = 'LocalSubnet'
    $top.Controls.Add($txtScope)
    $btnNarrow = New-Btn '3. Create narrow rule for port' 302 101 210 27 $top
    $lblNarrow = New-Lbl 'Only after the client test shows the port closed.' 520 106 $top
    $lblNarrow.ForeColor = [System.Drawing.Color]::DimGray

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Multiline = $true; $txtOut.ReadOnly = $true
    $txtOut.ScrollBars = 'Both'; $txtOut.WordWrap = $false
    $txtOut.Dock = 'Fill'
    $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtOut.BackColor = [System.Drawing.Color]::White
    $mid.Controls.Add($txtOut)

    $btnClose = New-Btn 'Close' 10 9 90 27 $bot
    $lblWait = New-Lbl '' 110 14 $bot
    $lblWait.ForeColor = [System.Drawing.Color]::DimGray

    $say = {
        param([string]$Text)
        $txtOut.AppendText($Text + [Environment]::NewLine)
        $txtOut.SelectionStart = $txtOut.TextLength
        $txtOut.ScrollToCaret()
    }
    $dlgButtons = @($btnPin, $btnRules, $btnRestore, $btnRefresh, $btnNarrow, $btnClose)
    $dlgBusy = {
        param([bool]$On, [string]$What)
        foreach ($b in $dlgButtons) { $b.Enabled = -not $On }
        $lblWait.Text = $(if ($On) { $What } else { '' })
        $f.Cursor = $(if ($On) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default })
        if (-not $On) { $btnNarrow.Enabled = ([string]$txtPort.Text -match '^\s*\d+\s*$') }
    }
    $readPort = {
        $n = 0
        if ([int]::TryParse(([string]$txtPort.Text).Trim(), [ref]$n) -and $n -ge 1024 -and $n -le 65535) { return $n }
        return 0
    }

    $showState = {
        & $dlgBusy $true 'Reading the listener, RPC policy and firewall...'
        try {
            $pinned = 0
            $cfg = $null; $lis = $null
            try { $cfg = Get-PrintRpcConfig } catch {}
            try { $lis = Get-PrintSpoolerListener } catch {}
            if ($cfg) {
                $row = @($cfg.Values | Where-Object { $_.Name -eq 'RpcTcpPort' })[0]
                if ($row -and $row.IsSet) { try { $pinned = [int]$row.Value } catch {} }
            }
            $fwR = Invoke-ToolkitResponsive -Command 'Get-PrintRpcFirewallState' -Arguments @{ PinnedPort = $pinned } -TimeoutSec 60
            & $say ''
            & $say ("--- {0}  state of {1} ---" -f (Get-Date -Format 'HH:mm:ss'), $env:COMPUTERNAME)
            $ports = if ($lis) { @($lis.Ports) } else { @() }
            & $say ('  Spooler listening on : ' + $(if ($ports.Count) { $ports -join ', ' } else { 'no TCP port' }))
            & $say ('  RpcTcpPort           : ' + $(if ($pinned) { "$pinned" + $(if ($ports -contains $pinned) { ' (in effect)' } else { ' (NOT in effect)' }) } else { 'not set - dynamic port, which a client cannot test' }))
            $fw = $fwR.Result
            if ($fw -and $fw.State -eq 'Ok') {
                & $say ('  Active profile       : ' + (@($fw.ActiveProfiles) -join ', '))
                foreach ($roleObj in @($fw.Smb, $fw.Epm, $fw.Spooler)) {
                    & $say ("  {0,-21}: {1}{2}" -f $roleObj.Label, $(if ($roleObj.Allowed) { 'allowed (' + (@($roleObj.AllowingRules) -join ', ') + ')' } else { 'NOT allowed' }), $(if ($roleObj.LocalSubnetOnly) { '  [LocalSubnet only]' } else { '' }))
                }
                if ($pinned) {
                    & $say ('  Rule for port ' + $pinned + $(if ($fw.Spooler.PinnedPortCovered) { '     : a toolkit rule exists for it' } else { '     : unknown whether the built-in rule admits it - measure from the client' }))
                }
                if (@($fw.ActiveProfiles) -contains 'Public') { & $say '  !! An adapter is on the PUBLIC profile. Set the network to Private before changing any rule.' }
            } else {
                & $say ('  Firewall             : could not be judged - ' + $(if ($fw) { $fw.Message } else { $fwR.Error }))
            }
            $snaps = @()
            try { $snaps = @(Get-PrintConfigSnapshot) } catch {}
            & $say ('  Saved snapshots      : ' + $snaps.Count)
            if ($pinned -and -not ([string]$txtPort.Text).Trim()) { $txtPort.Text = "$pinned" }
        } finally { & $dlgBusy $false '' }
    }

    $runChange = {
        param([string]$Label, [string]$Command, [hashtable]$Arguments, [string]$LogPrefix)
        & $dlgBusy $true "$Label..."
        try {
            $r = Invoke-ToolkitResponsive -Command $Command -Arguments $Arguments -TimeoutSec 120
            $msg = if ($r.Result) { [string]$r.Result.Message } elseif ($r.TimedOut) { 'Did not finish in time - check the Print Spooler service.' } else { "Failed: $($r.Error)" }
            & $say ''
            & $say ("--- {0}  {1} ---" -f (Get-Date -Format 'HH:mm:ss'), $Label)
            & $say $msg
            Write-Log ("${LogPrefix}: " + $msg) $(if ($r.Result -and $r.Result.Success) { 'OK' } else { 'WARN' })
        } finally { & $dlgBusy $false '' }
        & $showState
    }

    $btnPin.Add_Click({
        $port = & $readPort
        if (-not $port) { & $say 'Enter a port between 1024 and 65535 first.'; return }
        $q = "Pin this PC's Print Spooler RPC listener to TCP $port?`r`n`r`nThe spooler restarts, so jobs printing on this PC are interrupted. The previous state is saved so it can be restored exactly.`r`n`r`nSuccess is only reported if the spooler is then really listening on $port."
        if ((Show-Msg -Text $q -Title 'Pin the spooler port?' -Buttons 'YesNo' -Icon 'Question' -Default 'No') -ne 'Yes') { return }
        & $runChange "pin spooler to TCP $port" 'Set-PrintRpcTcpPort' @{ Port = $port } "Pin spooler RPC port $port"
    })

    $btnRules.Add_Click({
        $q = "Enable the scoped built-in inbound rules for the active profile - the RPC endpoint mapper (135) and the Print Spooler?`r`n`r`nThe any-address -V2 rules are never touched. The previous rule state is saved for exact restore."
        if ((Show-Msg -Text $q -Title 'Enable built-in rules?' -Buttons 'YesNo' -Icon 'Question' -Default 'No') -ne 'Yes') { return }
        & $runChange 'enable built-in spooler rules' 'Enable-PrintRpcFirewall' @{} 'Enable built-in spooler rules'
    })

    $btnNarrow.Add_Click({
        $port = & $readPort
        if (-not $port) { & $say 'Enter the pinned port first.'; return }
        $scope = ([string]$txtScope.Text).Trim()
        $q = "Create ONE inbound rule allowing spoolsv.exe on TCP $port from '$scope'?`r`n`r`nDo this only after Shares -> Test connection on the client measured port $port as closed. The rule is recorded in the snapshot and removed again by Restore. 'Any' is refused."
        if ((Show-Msg -Text $q -Title 'Create a narrow rule?' -Buttons 'YesNo' -Icon 'Warning' -Default 'No') -ne 'Yes') { return }
        & $runChange "narrow rule TCP $port ($scope)" 'Enable-PrintRpcFirewall' @{ Port = $port; CreateRuleForPort = $true; RemoteAddress = $scope } "Narrow spooler rule TCP $port"
    })

    $btnRestore.Add_Click({
        $snaps = @()
        try { $snaps = @(Get-PrintConfigSnapshot) } catch {}
        if (-not $snaps.Count) { & $say 'There is no saved snapshot on this PC to restore.'; return }
        $pick = Show-PickList -Title 'Restore a snapshot' -Prompt 'Newest first. Values that were absent are removed again; rules that did not exist are deleted.' -Items $snaps -Label { param($s) '{0}   {1}' -f $s.Created, $s.Label }
        if (-not $pick) { return }
        if ((Show-Msg -Text "Restore '$($pick.Label)' from $($pick.Created)?`r`n`r`nThe spooler restarts if a registry value changes." -Title 'Restore?' -Buttons 'YesNo' -Icon 'Warning' -Default 'No') -ne 'Yes') { return }
        & $runChange "restore $($pick.Label)" 'Restore-PrintConfigSnapshot' @{ Path = [string]$pick.Path } 'Restore print RPC snapshot'
    })

    $btnRefresh.Add_Click({ & $showState })
    $btnClose.Add_Click({ $f.Close() })
    $txtPort.Add_TextChanged({ $btnNarrow.Enabled = ([string]$txtPort.Text -match '^\s*\d+\s*$') })

    $f.Add_Shown({
        & $say 'Recommended order on this PC:  1. Pin the spooler to a port   2. Enable the built-in rules.'
        & $say 'Then on the CLIENT: Shares -> Test connection, entering the same port.'
        & $say 'Only if that shows the port closed, come back here for 3. Create narrow rule.'
        & $showState
    })
    [void]$f.ShowDialog((Get-MsgOwner))
    $f.Dispose()
}


function Apply-HostResults {
    param($sh)
    $h = $sh.HostHealth
    if (-not $h) { return }
    $script:HostHealth   = $h
    $script:HostFindings = @($h.Findings)

    Set-Rows $dgvHost $script:HostFindings @('Code', 'Severity', 'Printer', 'Detail')
    for ($i = 0; $i -lt $dgvHost.Rows.Count; $i++) {
        switch ("$($dgvHost.Rows[$i].Cells[1].Value)") {
            'ERROR' { $dgvHost.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick }
            'WARN'  { $dgvHost.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 95, 0) }
            default { $dgvHost.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::DimGray }
        }
    }

    if (-not $h.IsHost) {
        $lblHostState.Text = 'This PC does not share any printer.'
        $lblHostState.ForeColor = [System.Drawing.Color]::DimGray
        $lblHostShares.Text = 'Nothing to check. Share a printer here, or use the Shares tab to connect to a printer shared from another PC.'
        $txtHost.Text = ''
        return
    }

    $lblHostState.Text = "Sharing $(@($h.Shares).Count) printer(s) - $($h.Summary)"
    $lblHostState.ForeColor = if ($h.Findings | Where-Object { $_.Severity -eq 'ERROR' }) { [System.Drawing.Color]::Firebrick }
                              elseif ($h.Findings | Where-Object { $_.Severity -eq 'WARN' }) { [System.Drawing.Color]::FromArgb(180, 95, 0) }
                              else { [System.Drawing.Color]::DarkGreen }
    $lblHostShares.Text = 'Shares: ' + ((@($h.Shares) | ForEach-Object { $_.Unc }) -join '   |   ')

    # Context panel: the environment facts that decide whether clients can connect.
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Network profile : " + ((@($h.Profiles) | ForEach-Object { "$($_.InterfaceAlias)=$($_.NetworkCategory)" }) -join ', '))
    [void]$sb.AppendLine("File+Print FW   : $($h.FirewallInboundEnabled) of $(@($h.Firewall | Where-Object { "$($_.Direction)" -eq 'Inbound' }).Count) inbound FPS rule(s) enabled")
    [void]$sb.AppendLine("Spooler         : $($h.Spooler.Status) ($($h.Spooler.StartMode))")
    [void]$sb.AppendLine("print`$ readable : $($h.PrintShare)")
    [void]$sb.AppendLine("Membership      : " + $(if ($h.Auth.PartOfDomain) { "domain $($h.Auth.Domain)" } else { "WORKGROUP $($h.Auth.Domain) - clients must authenticate with an account that exists on THIS PC" }))
    [void]$sb.AppendLine("Guest account   : " + $(if ($null -eq $h.Auth.GuestEnabled) { 'unknown' } elseif ($h.Auth.GuestEnabled) { 'enabled' } else { 'disabled' }))
    [void]$sb.AppendLine('')
    # What a connecting client needs from this host, per path, and what the spooler is
    # really listening on. Configuration and activation are shown separately on purpose.
    [void]$sb.AppendLine('Print RPC on this PC (what a connecting client needs):')
    [void]$sb.AppendLine('   Spooler listening on : ' + $(if (@($h.Listener.Ports).Count) { @($h.Listener.Ports) -join ', ' } else { 'no TCP port' }))
    [void]$sb.AppendLine('   RpcTcpPort           : ' + $(if ([int]$h.PinnedPort) { "$($h.PinnedPort)" + $(if (@($h.Listener.Ports) -contains [int]$h.PinnedPort) { ' (in effect)' } else { ' (NOT in effect)' }) } else { 'not set - dynamic port, which a client cannot test' }))
    if ($h.RpcFirewall -and $h.RpcFirewall.State -eq 'Ok') {
        foreach ($roleObj in @($h.RpcFirewall.Smb, $h.RpcFirewall.Epm, $h.RpcFirewall.Spooler)) {
            [void]$sb.AppendLine(("   {0,-21}: {1}{2}" -f $roleObj.Label, $(if ($roleObj.Allowed) { 'allowed' } else { 'NOT allowed' }), $(if ($roleObj.LocalSubnetOnly) { ' (LocalSubnet only)' } else { '' })))
        }
    } elseif ($h.RpcFirewall) {
        [void]$sb.AppendLine('   Firewall             : ' + [string]$h.RpcFirewall.Message)
    }
    foreach ($v in @($h.RpcConfig.Values)) { [void]$sb.AppendLine(("   {0,-32} {1}" -f $v.Name, $v.State)) }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('SMB / auth settings that gate access to these shares:')
    foreach ($v in @($h.Auth.Values)) { [void]$sb.AppendLine(("   {0,-32} {1}" -f $v.Name, $v.State)) }
    [void]$sb.AppendLine('')
    foreach ($s in @($h.Shares)) {
        [void]$sb.AppendLine("SHARE  $($s.Unc)")
        [void]$sb.AppendLine("   queue      : $($s.Name)")
        [void]$sb.AppendLine("   driver     : $($s.DriverName)$(if ($null -ne $s.DriverMajor) { "  (v$($s.DriverMajor)$(if ($null -ne $s.PackageAware) { ", package-aware=$($s.PackageAware)" } else { '' }))" } else { '' })")
        [void]$sb.AppendLine("   port       : $($s.PortName) [$($s.PortType)]")
        if ($null -ne $s.UsbAttached) { [void]$sb.AppendLine("   usb device : $(if ($s.UsbAttached) { 'attached' } else { 'NOT ATTACHED' })") }
        [void]$sb.AppendLine("   permissions: $($s.AclMessage)")
        [void]$sb.AppendLine('')
    }
    $txtHost.Text = $sb.ToString()
    Update-RpcStateLabels
    Write-Log $h.Summary $(if ($h.Findings | Where-Object { $_.Severity -eq 'ERROR' }) { 'ERROR' } elseif ($h.Findings | Where-Object { $_.Severity -eq 'WARN' }) { 'WARN' } else { 'OK' })
}

$btnHostCheck.Add_Click({ $txtHost.Text = ''; Start-Worker 'sharehost' })

$btnHostFix.Add_Click({
    $h = $script:HostHealth
    if (-not $h) { [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 'Run Check this PC first.', 'Printer Manager', 'OK', 'Information'); return }
    $codes = @($h.Findings | ForEach-Object { $_.Code })

    # Only offer what this machine actually needs, and only the changes that are
    # not security downgrades. Guest auth, SMB signing and Point-and-Print are
    # never offered here - naming them is the job, weakening them is not.
    $offers = @()
    if ($codes -contains 'SHR001') { $offers += 'profile' }
    if ($codes -contains 'SHR002') { $offers += 'firewall' }
    if ($codes -contains 'SHR006') { $offers += 'acl' }
    if (-not $offers.Count) {
        [void](Show-Msg -Text "Nothing here can be safely auto-fixed.`n`nThe remaining findings are either informational, or they are security settings this console deliberately will not change (guest authentication, SMB signing, Point and Print, RPC privacy)." -Title 'No safe fix available' -Icon 'Information')
        return
    }

    Set-Busy $true 'Applying safe fixes...'
    try {
        if ($offers -contains 'profile') {
            if ([System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 'Set Public network adapters to Private? Windows blocks inbound sharing on Public.', 'Network profile', 'YesNo', 'Question') -eq 'Yes') {
                $r = Set-PrintSharingNetworkPrivate; Write-Log $r.Message $(if ($r.Success) { 'OK' } else { 'WARN' })
            }
        }
        if ($offers -contains 'firewall') {
            if ([System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 'Enable the File and Printer Sharing inbound firewall rules? Scope and profile restrictions are left as shipped.', 'Firewall', 'YesNo', 'Question') -eq 'Yes') {
                $r = Enable-PrintSharingFirewall; Write-Log $r.Message $(if ($r.Success) { 'OK' } else { 'WARN' })
            }
        }
        if ($offers -contains 'acl') {
            foreach ($f in @($h.Findings | Where-Object { $_.Code -eq 'SHR006' })) {
                if ([System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), "Grant Print on '$($f.Printer)' to Authenticated Users? Existing permissions are kept.", 'Printer permissions', 'YesNo', 'Question') -eq 'Yes') {
                    $r = Grant-PrinterPrintAccess -Name $f.Printer -Trustee AuthenticatedUsers
                    Write-Log $r.Message $(if ($r.Success) { 'OK' } else { 'WARN' })
                }
            }
        }
    } finally { Set-Busy $false }
    Start-Worker 'sharehost'   # re-check so the findings reflect what changed
})

$dgvHost.Add_SelectionChanged({
    if ($dgvHost.SelectedRows.Count -eq 0) { return }
    $i = $dgvHost.SelectedRows[0].Index
    if ($i -lt 0 -or $i -ge @($script:HostFindings).Count) { return }
    $f = @($script:HostFindings)[$i]
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("$($f.Code)  [$($f.Severity)]")
    if ($f.Printer) { [void]$sb.AppendLine("Printer : $($f.Printer)") }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('WHAT WAS FOUND')
    [void]$sb.AppendLine("  $($f.Detail)")
    if ($f.Fix) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('FIX')
        [void]$sb.AppendLine("  $($f.Fix)")
    }
    $txtHost.Text = $sb.ToString()
})

$btnHostCopy.Add_Click({
    $uncs = @($script:HostHealth.Shares | ForEach-Object { $_.Unc })
    if (-not $uncs.Count) { Write-Log 'No shares on this PC to copy.' 'WARN'; return }
    try { Set-Clipboard -Value ($uncs -join [Environment]::NewLine); Write-Log "Copied $($uncs.Count) share path(s) to the clipboard." 'OK' }
    catch { Write-Log "Clipboard unavailable: $($_.Exception.Message)" 'WARN' }
})

$btnAddToPc.Add_Click({
    # ONE Add for a shared printer, honouring the scope dropdown. It collects what
    # it needs here, then runs in the background worker; on failure the worker
    # measures why, and the next step is chosen from that diagnosis category.
    $unc = Get-SelectedServerUnc -Title 'Add a shared printer'
    if (-not $unc) { return }
    if ($unc -notmatch '^\\\\[^\\]+\\.+') { [void](Show-Msg -Text "'$unc' is not a \\host\sharename path." -Title 'Add' -Icon 'Warning'); return }

    switch ($cboAddScope.SelectedIndex) {
        2 {
            $guess = ''
            if ($dgvServer.SelectedRows.Count -gt 0) {
                $gi = $dgvServer.SelectedRows[0].Index
                if ($gi -ge 0 -and $gi -lt @($script:ServerRows).Count) { $guess = "$(@($script:ServerRows)[$gi].DriverName)" }
            }
            $drv = Show-InputBox -Title 'Driver for the local queue' -Default $guess -Prompt (
                "This creates a LOCAL queue printing to $unc, so Windows never fetches a driver from the host.`n`n" +
                'The driver must ALREADY be installed on this PC. Enter its exact name (Install driver... can install one first):')
            if ([string]::IsNullOrWhiteSpace($drv)) { return }
            $drv = $drv.Trim()
            if (-not (Get-PrinterDriver -Name $drv -ErrorAction SilentlyContinue)) {
                if ((Show-Msg -Text "The driver '$drv' is not installed on this PC, and this method never fetches one from the host.`n`nInstall it from an INF now?" -Title 'Driver not installed' -Buttons 'YesNo' -Icon 'Question') -ne 'Yes') { return }
                $installed = Show-InstallDriverDialog -SuggestName $drv
                if (-not $installed) { return }
                $drv = $installed
            }
            $qname = Show-InputBox -Title 'Queue name' -Default (($unc -split '\\')[-1]) -Prompt 'Name for the new local queue:'
            if ([string]::IsNullOrWhiteSpace($qname)) { return }
            $ask = "Create local queue '$($qname.Trim())' printing to $unc with the local driver '$drv'?`n`n" +
                   "Every user of this PC gets it. Jobs reach the share as this PC's machine identity, which a workgroup host may refuse - run the delivery test afterwards."
            if ((Show-Msg -Text $ask -Title 'Confirm local-port printer' -Buttons 'YesNo' -Icon 'Question') -ne 'Yes') { return }
            Start-Worker 'addconn' '' ([pscustomobject]@{ Scope = 'LocalPort'; ScopeLabel = 'every user, via local port'; Unc = $unc; DriverName = $drv; QueueName = $qname.Trim(); User = [string]$script:TargetUser })
        }
        1 {
            if (-not (Require-User)) { return }
            if ((Show-Msg -Text "Add '$unc' for $($script:TargetUser) only? It applies immediately, in their session." -Title 'Confirm' -Buttons 'YesNo' -Icon 'Question') -ne 'Yes') { return }
            Start-Worker 'addconn' '' ([pscustomobject]@{ Scope = 'User'; ScopeLabel = "just $($script:TargetUser)"; Unc = $unc; User = [string]$script:TargetUser })
        }
        default {
            if ($script:Ctx.IsAdmin -eq $false -and -not $script:IsSystem) {
                [void](Show-Msg -Text "A per-computer connection needs elevation.`n`nRun this console as Administrator, or as SYSTEM from Backstage - or choose the per-user scope." -Title 'Elevation required' -Icon 'Warning')
                return
            }
            $ask = "Add '$unc' for EVERY user of this PC (per-computer)?`n`n" +
                   'Each user sees it at their next sign-in and connects in their OWN context, so this does not fix a user who cannot authenticate to the host.'
            if ((Show-Msg -Text $ask -Title 'Confirm' -Buttons 'YesNo' -Icon 'Question') -ne 'Yes') { return }
            Start-Worker 'addconn' '' ([pscustomobject]@{ Scope = 'Machine'; ScopeLabel = 'every user (per-computer)'; Unc = $unc; User = [string]$script:TargetUser })
        }
    }
})

$btnSrvTest.Add_Click({
    # Measures the EXACT connection - path, live open as this console, live open as the
    # target user - without changing anything on either PC.
    $unc = Get-SelectedServerUnc -Title 'Test a shared printer'
    if (-not $unc) { return }
    Start-EvidenceTest -Unc $unc
})

$btnSrvRemoveMachine.Add_Click({
    if ($script:Ctx.IsAdmin -eq $false -and -not $script:IsSystem) {
        [void](Show-Msg -Text 'Removing a per-computer connection needs elevation. Run as Administrator, or as SYSTEM from Backstage.' -Title 'Elevation required' -Icon 'Warning')
        return
    }
    $conns = @()
    try { $conns = @(Get-MachinePrinterConnections | Where-Object { $_ }) } catch {}
    if (-not $conns.Count) { [void](Show-Msg -Text 'This PC has no per-computer printer connections.' -Title 'Remove per-computer connection' -Icon 'Information'); return }
    $pick = Show-PickList -Title 'Remove a per-computer connection' -Prompt 'Every user of this PC loses the chosen printer at their next sign-in.' -Items $conns -Label { param($c) [string]$c.Unc }
    if (-not $pick) { return }
    $target = [string]$pick.Unc
    if ((Show-Msg -Text "Remove the per-computer connection '$target'?`n`nEvery user of this PC loses it at their next sign-in. The host and its driver are not touched." -Title "Confirm remove - $target" -Buttons 'YesNo' -Icon 'Warning' -Default 'No') -ne 'Yes') { return }
    Write-Log "Removing per-computer connection '$target' - note it if it needs re-adding." 'WARN'
    Set-Busy $true 'Removing per-computer connection...'
    try {
        $r = Remove-PrinterConnectionForMachine -ConnectionName $target -OnPoll $script:Pump
        # Verify rather than trust: the removal is requested through printui, whose
        # own exit status says nothing.
        $still = $true
        try { $still = [bool](Test-MachinePrinterConnection -ConnectionName $target) } catch {}
        if ($r.Success -and $still) {
            $r = [pscustomobject]@{ Success = $false; Message = "$($r.Message) But it is still registered when checked again - it has not gone." }
        } elseif (-not $r.Success -and -not $still) {
            $r = [pscustomobject]@{ Success = $true; Message = "'$target' is no longer registered (confirmed by re-reading the registry), although the removal reported: $($r.Message)" }
        }
        Show-Result $r 'Per-computer connection removed'
    } finally { Set-Busy $false }
    Start-Worker 'scan'
})

$btnHostRpcPort.Add_Click({ Show-RpcPortDialog; Update-RpcStateLabels })


function Show-PaperSizeDialog {
    <#
    .SYNOPSIS
        Printers tab: the paper size and duplex setting each queue and connection uses, with the supported fixes.
    .DESCRIPTION
        Paper is judged against the region: this PC's system region for its local
        queues, the target user's region for them and their connections. Duplex has
        no right answer, so it is only compared: a user whose own saved preferences
        use a different setting from the default is shown in amber, because their
        copy replaces the default for them.

        The reads are registry-only and take well under a second, so they run here.
        The changes run off the UI thread through Invoke-ToolkitResponsive:
          * a local queue's default paper size (needs elevation)
          * a local queue's default duplex setting (needs elevation)
          * the target user's regional paper size (in their session)
        All confirm with No as the default, are snapshotted, and are verified by
        re-reading. A user's own saved printer preferences are never reset here.
    #>
    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Paper and duplex'
    $f.Size = New-Object System.Drawing.Size(1000, 600)
    $f.MinimumSize = New-Object System.Drawing.Size(790, 480)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'Sizable'
    $f.MinimizeBox = $false; $f.MaximizeBox = $true; $f.ShowInTaskbar = $false
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $top = New-DockBand $f 'Top' 58
    $bot = New-DockBand $f 'Bottom' 104
    $lblPc = New-Lbl '' 10 8 $top $true
    $lblPc.MaximumSize = New-Object System.Drawing.Size(960, 0)
    $lblUsr = New-Lbl '' 10 32 $top
    $lblUsr.MaximumSize = New-Object System.Drawing.Size(960, 0)

    $grid = New-Grid 0 0 10 10 $f
    Add-Col $grid 'ppQueue' 'Printer'        250 | Out-Null
    Add-Col $grid 'ppKind'  'Kind'            75 | Out-Null
    Add-Col $grid 'ppDef'   'Default paper'  110 | Out-Null
    Add-Col $grid 'ppUser'  "User's paper"    95 | Out-Null
    Add-Col $grid 'ppExp'   'Expected'        75 | Out-Null
    Add-Col $grid 'ppDup'   'Default duplex' 130 | Out-Null
    Add-Col $grid 'ppUDup'  "User's duplex"  130 | Out-Null
    Add-Col $grid 'ppState' 'Paper check'     95 | Out-Null
    Add-Col $grid 'ppNote'  'Note'           330 | Out-Null
    Set-FillGrid $grid

    # Queue actions and the user action on the first row; housekeeping below, so the
    # band fits the dialog's minimum width.
    $btnQueue   = New-Btn 'Set queue paper...'       10  8  190 27 $bot
    $btnDuplex  = New-Btn 'Set queue duplex...'      206 8  150 27 $bot
    $btnUserPap = New-Btn "Set user's paper size..." 362 8  220 27 $bot
    $btnRestore = New-Btn 'Restore previous...'      10  42 140 27 $bot
    $btnRefresh = New-Btn 'Refresh'                  156 42 90  27 $bot
    $btnClose   = New-Btn 'Close'                    252 42 90  27 $bot
    $lblWait = New-Lbl '' 10 78 $bot
    $lblWait.ForeColor = [System.Drawing.Color]::DimGray
    $lblWait.MaximumSize = New-Object System.Drawing.Size(960, 0)

    $amber = [System.Drawing.Color]::FromArgb(180, 95, 0)
    $boldFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $state = @{ Machine = $null; User = $null; Rows = @() }

    $selectedRow = {
        if ($grid.SelectedRows.Count -eq 0) { return $null }
        $ix = $grid.SelectedRows[0].Index
        $all = @($state.Rows)
        if ($ix -lt 0 -or $ix -ge $all.Count) { return $null }
        return $all[$ix]
    }
    $updateButtons = {
        $r = & $selectedRow
        $isLocal = [bool]($r -and $r.Kind -eq 'Local')
        $btnQueue.Enabled = [bool]($isLocal -and $r.ServerState -eq 'Mismatch' -and $r.Expected)
        $btnQueue.Text = if ($isLocal -and $r.Expected) { "Set queue paper to $($r.Expected)..." } else { 'Set queue paper...' }
        $btnDuplex.Enabled = [bool]($isLocal -and $r.QueueDuplexDeclared -and $r.QueueDuplex)
        $u = $state.User
        $btnUserPap.Enabled = [bool]($u -and $u.HiveLoaded -and $u.Match -eq $false -and $u.ExpectedPaperCode)
        $btnUserPap.Text = if ($u -and $u.ExpectedPaper) { "Set user's paper size to $($u.ExpectedPaper)..." } else { "Set user's paper size..." }
    }
    $refresh = {
        $state.Machine = $null; $state.User = $null
        try { $state.Machine = Get-RegionPaperSize -Sid '.DEFAULT' } catch {}
        if ($script:TargetSid) { try { $state.User = Get-RegionPaperSize -Sid ([string]$script:TargetSid) -User ([string]$script:TargetUser) } catch {} }
        $m = $state.Machine
        $u = $state.User
        $mc = if ($m) { [int]$m.ExpectedPaperCode } else { 0 }
        $uc = if ($u -and $u.ExpectedPaperCode) { [int]$u.ExpectedPaperCode } else { $mc }
        $virt = @(@($script:Printers) | Where-Object { $_ -and $_.IsVirtual } | ForEach-Object { [string]$_.Name })
        $rows = @()
        try { $rows = @(Get-PrinterPaperSizeState -Sid ([string]$script:TargetSid) -MachineExpectedCode $mc -UserExpectedCode $uc -VirtualNames $virt) }
        catch { $lblWait.Text = 'Could not read paper and duplex settings: ' + $_.Exception.Message }
        $state.Rows = $rows

        $lblPc.Text = if ($m -and $m.ExpectedPaperCode) { "This PC: $($m.ExpectedFrom) -> $($m.ExpectedPaper) expected for its queues." }
                      else { 'This PC: the expected paper size could not be worked out from its region.' }
        $lblUsr.Text = if (-not $script:TargetSid) { 'No target user is selected, so no user region, saved preferences or printer connections are shown.' }
                       elseif (-not $u -or -not $u.HiveLoaded) { "$($script:TargetUser): profile not loaded - check again while they are signed in." }
                       elseif ($u.Match -eq $false) { "$($u.User): regional paper size is $($u.UserPaper), but $($u.ExpectedFrom) implies $($u.ExpectedPaper)." }
                       else { "$($u.User): regional paper size is $($u.UserPaper), matching $($u.ExpectedFrom)." }
        $lblUsr.ForeColor = if ($u -and $u.Match -eq $false) { $amber } else { [System.Drawing.Color]::Black }

        # Display copy: a driver that does not use duplex says so rather than
        # showing a one-sided value it ignores.
        $view = @(foreach ($rw in $rows) {
            [pscustomobject]@{
                Queue = $rw.Queue; Kind = $rw.Kind; QueueDefault = $rw.QueueDefault; UserCopy = $rw.UserCopy; Expected = $rw.Expected
                QueueDuplex = $(if ($rw.QueueDuplex) { $rw.QueueDuplex } elseif ($rw.QueueDefault) { 'not used' } else { '' })
                UserDuplex = $rw.UserDuplex; State = $rw.State; Note = $rw.Note
            }
        })
        Set-Rows $grid $view @('Queue', 'Kind', 'QueueDefault', 'UserCopy', 'Expected', 'QueueDuplex', 'UserDuplex', 'State', 'Note')
        for ($i = 0; $i -lt $grid.Rows.Count -and $i -lt $rows.Count; $i++) {
            if ($rows[$i].State -eq 'Mismatch') { $grid.Rows[$i].DefaultCellStyle.ForeColor = $amber }
            elseif ($rows[$i].State -in @('NotComparable', 'Unknown')) { $grid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::DimGray }
            if ($rows[$i].DuplexState -eq 'Differs') {
                foreach ($ci in 5, 6) { $grid.Rows[$i].Cells[$ci].Style.ForeColor = $amber; $grid.Rows[$i].Cells[$ci].Style.Font = $boldFont }
            }
        }
        $grid.ClearSelection()
        & $updateButtons
    }

    $dlgButtons = @($btnQueue, $btnDuplex, $btnUserPap, $btnRestore, $btnRefresh, $btnClose)
    $dlgBusy = {
        param([bool]$On, [string]$What)
        foreach ($b in $dlgButtons) { $b.Enabled = -not $On }
        if ($On) { $lblWait.Text = $What }
        $f.Cursor = if ($On) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default }
        if (-not $On) { & $updateButtons }
    }
    $runChange = {
        param([string]$Label, [string]$Command, [hashtable]$Arguments)
        & $dlgBusy $true "$Label..."
        $msg = ''
        try {
            $res = Invoke-ToolkitResponsive -Command $Command -Arguments $Arguments -TimeoutSec 150
            $msg = if ($res.Result) { [string]$res.Result.Message } elseif ($res.TimedOut) { 'Did not finish in time - nothing was confirmed.' } else { "Failed: $($res.Error)" }
            Write-Log ('Paper and duplex: ' + $msg) $(if ($res.Result -and $res.Result.Success) { 'OK' } else { 'WARN' })
        } finally { & $dlgBusy $false '' }
        & $refresh
        $lblWait.Text = $msg
    }
    $needsElevation = {
        if ($script:Ctx.IsAdmin -eq $false -and -not $script:IsSystem) {
            [void](Show-Msg -Text 'Changing a queue default needs elevation. Run this console as Administrator, or as SYSTEM from Backstage.' -Title 'Elevation required' -Icon 'Warning')
            return $true
        }
        return $false
    }

    $grid.Add_SelectionChanged({ & $updateButtons })

    $btnQueue.Add_Click({
        $r = & $selectedRow
        if (-not $r -or $r.Kind -ne 'Local' -or $r.ServerState -ne 'Mismatch' -or -not $r.Expected) { return }
        if (& $needsElevation) { return }
        $q = "Set the default paper size of '$($r.Queue)' to $($r.Expected)?`r`n`r`nIt is $($r.QueueDefault) now. " +
             $(if ($r.Shared) { 'This queue is shared, so new connections get the new default. ' } else { '' }) +
             'Anyone with their own saved Printing Preferences for it keeps theirs. The previous size is saved so Restore previous can put it back.'
        if ((Show-Msg -Text $q -Title 'Set queue paper?' -Buttons 'YesNo' -Icon 'Question' -Default 'No') -ne 'Yes') { return }
        & $runChange "Setting '$($r.Queue)' to $($r.Expected)" 'Set-PrinterQueuePaperSize' @{ Name = [string]$r.Queue; PaperSize = [string]$r.Expected }
    })

    $btnDuplex.Add_Click({
        $r = & $selectedRow
        if (-not $r -or $r.Kind -ne 'Local' -or -not $r.QueueDuplexDeclared -or -not $r.QueueDuplex) { return }
        if (& $needsElevation) { return }
        $modes = @(@(
            [pscustomobject]@{ Mode = 'OneSided';          Label = 'One-sided' }
            [pscustomobject]@{ Mode = 'TwoSidedLongEdge';  Label = 'Two-sided, long edge (flip like a book)' }
            [pscustomobject]@{ Mode = 'TwoSidedShortEdge'; Label = 'Two-sided, short edge (flip like a notepad)' }
        ) | Where-Object { $_.Mode -ne $r.QueueDuplex })
        $pick = Show-PickList -Title 'Set queue duplex' -Prompt "'$($r.Queue)' is $(Get-DuplexLabel $r.QueueDuplex) by default. Choose its new default:" -Items $modes -Label { param($x) $x.Label }
        if (-not $pick) { return }
        $q = "Set the default duplex of '$($r.Queue)' to $(Get-DuplexLabel $pick.Mode)?`r`n`r`n" +
             $(if ($r.Shared) { 'This queue is shared, so new connections get the new default. ' } else { '' }) +
             'Anyone with their own saved Printing Preferences for it keeps theirs - their copy replaces the default for them. The previous setting is saved so Restore previous can put it back.'
        if ((Show-Msg -Text $q -Title 'Set queue duplex?' -Buttons 'YesNo' -Icon 'Question' -Default 'No') -ne 'Yes') { return }
        & $runChange "Setting '$($r.Queue)' to $(Get-DuplexLabel $pick.Mode)" 'Set-PrinterQueueDuplex' @{ Name = [string]$r.Queue; Mode = [string]$pick.Mode }
    })

    $btnUserPap.Add_Click({
        $u = $state.User
        if (-not $u -or $u.Match -ne $false -or -not $u.ExpectedPaperCode) { return }
        $q = "Set $($u.User)'s regional paper size to $($u.ExpectedPaper)?`r`n`r`nIt is $($u.UserPaper) now, but $($u.ExpectedFrom) implies $($u.ExpectedPaper). Word and most programs use it for new documents.`r`n`r`n" +
             'Only this one setting changes - not their region format, language or home location. Programs already open pick it up when restarted. The previous value is saved so Restore previous can put it back.'
        if ((Show-Msg -Text $q -Title 'Set the user''s paper size?' -Buttons 'YesNo' -Icon 'Question' -Default 'No') -ne 'Yes') { return }
        & $runChange "Setting $($u.User)'s regional paper size to $($u.ExpectedPaper)" 'Set-UserRegionPaperSize' @{ Sid = [string]$u.Sid; User = [string]$u.User; PaperCode = [int]$u.ExpectedPaperCode }
    })

    $btnRestore.Add_Click({
        $snaps = @()
        try { $snaps = @(Get-PrinterPaperSnapshot) } catch {}
        if (-not $snaps.Count) { $lblWait.Text = 'There is no saved paper or duplex snapshot on this PC.'; return }
        $pick = Show-PickList -Title 'Restore a setting' -Prompt 'Newest first. The recorded setting is put back exactly.' -Items $snaps -Label { param($s) '{0}   {1}' -f $s.Created, $s.Label }
        if (-not $pick) { return }
        if ((Show-Msg -Text "Restore this?`r`n`r`n$($pick.Label)" -Title 'Restore?' -Buttons 'YesNo' -Icon 'Warning' -Default 'No') -ne 'Yes') { return }
        & $runChange "Restoring: $($pick.Label)" 'Restore-PrinterPaperSnapshot' @{ Path = [string]$pick.Path }
    })

    $btnRefresh.Add_Click({ $lblWait.Text = ''; & $refresh })
    $btnClose.Add_Click({ $f.Close() })
    $f.Add_Shown({ & $refresh })
    [void]$f.ShowDialog((Get-MsgOwner))
    $f.Dispose()
}

$btnPaper.Add_Click({ Show-PaperSizeDialog })


# --------------------------------------------------------------------------
# Credentials dialog - test / save / repair access to a print host
#
# The password is masked, never logged, never written to the report and never
# placed in the JSON envelope. It does reach cmdkey on a command line, which is
# briefly visible in the process list; that trade-off was chosen deliberately.
# --------------------------------------------------------------------------
function Show-JunkQueueDialog {
    <#
    .SYNOPSIS
        Review and remove the junk queues - Fax, XPS, OneNote for Windows 10.
    .DESCRIPTION
        A REVIEW step rather than a one-click purge, for three reasons that each
        cost a callback if ignored:

          * XPS and Fax are Windows OPTIONAL FEATURES. Deleting the queue is
            temporary - Windows puts it back on servicing or a repair - so the
            dialog says so per queue and offers to disable the feature too.
          * Deleting somebody's DEFAULT printer is its own support call. Windows
            then reassigns arbitrarily, or LegacyDefaultPrinterMode picks
            last-used, which is the MAP004 finding.
          * A shared junk queue may have clients pointed at it.

        The candidate list comes from the narrow exact-match catalogue, never
        from the virtual-printer regex - that one is a display filter and would
        offer up CutePDF and Microsoft Print to PDF.
    #>
    param($Candidates)

    $cands = @($Candidates)
    if (-not $cands.Count) {
        [void](Show-Msg -Text "Nothing to clean up - this PC has no Fax, Microsoft XPS Document Writer or OneNote for Windows 10 queue.`n`nOnly those three are ever offered. PDF writers you actually use (Microsoft Print to PDF, CutePDF and the like) are deliberately left alone." -Title 'No junk queues' -Icon 'Information')
        return $false
    }

    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Remove junk print queues'
    $f.Size = New-Object System.Drawing.Size(760, 560)
    $f.MinimumSize = New-Object System.Drawing.Size(660, 480)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'Sizable'
    $f.MinimizeBox = $false; $f.MaximizeBox = $false
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    # Heights measured, not guessed: at 46 the second label's descender was
    # clipped, and at 128 the button row sat 3px below the band.
    $top = New-DockBand $f 'Top' 54
    $bot = New-DockBand $f 'Bottom' 142
    $mid = New-Object System.Windows.Forms.Panel
    $mid.Dock = 'Fill'
    $f.Controls.Add($mid)
    $mid.BringToFront()

    $lblTop = New-Lbl 'Tick what to remove. Nothing is deleted until you press Remove selected.' 10 8 $top $true
    $lblTop2 = New-Lbl 'Only Fax, Microsoft XPS Document Writer and OneNote for Windows 10 are ever listed here.' 10 26 $top
    $lblTop2.ForeColor = [System.Drawing.Color]::DimGray

    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Dock = 'Fill'
    $clb.CheckOnClick = $true
    $clb.IntegralHeight = $false
    $mid.Controls.Add($clb)

    $txtWhy = New-Object System.Windows.Forms.TextBox
    $txtWhy.Multiline = $true; $txtWhy.ReadOnly = $true
    $txtWhy.ScrollBars = 'Vertical'
    $txtWhy.Dock = 'Top'
    $txtWhy.Height = 74
    $txtWhy.BackColor = [System.Drawing.Color]::White
    $bot.Controls.Add($txtWhy)

    $chkFeature = New-Object System.Windows.Forms.CheckBox
    $chkFeature.Text = 'Also disable the Windows feature, so XPS/Fax cannot come back (may need a restart)'
    $chkFeature.Location = New-Object System.Drawing.Point(10, 80)
    $chkFeature.Size = New-Object System.Drawing.Size(560, 22)
    $bot.Controls.Add($chkFeature)

    $btnGo     = New-Btn 'Remove selected' 10  104 140 27 $bot
    $btnCancel = New-Btn 'Cancel'          156 104 100 27 $bot
    $lblCount  = New-Lbl '' 266 110 $bot
    $lblCount.ForeColor = [System.Drawing.Color]::DimGray

    # --- populate ---------------------------------------------------------
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($c in $cands) {
        $tags = [System.Collections.Generic.List[string]]::new()
        if ($c.ComesBack) { $tags.Add('comes back unless the feature is disabled') }
        if (@($c.IsDefaultFor).Count) { $tags.Add('DEFAULT for ' + (@($c.IsDefaultFor) -join ', ')) }
        if ($c.Shared) { $tags.Add('SHARED') }
        $suffix = if ($tags.Count) { '   -   ' + ($tags -join '; ') } else { '' }
        $rows.Add([pscustomobject]@{ Cand = $c; Text = ($c.Name + $suffix) })
    }
    foreach ($r in $rows) {
        # Pre-ticked ONLY when it is nobody's default and not shared. Anything
        # with a consequence starts unticked so it takes a deliberate click.
        $safe = ((@($r.Cand.IsDefaultFor).Count -eq 0) -and (-not $r.Cand.Shared))
        [void]$clb.Items.Add($r.Text, $safe)
    }

    $anyFeature = @($cands | Where-Object { $_.Feature }).Count -gt 0
    $chkFeature.Enabled = $anyFeature
    if (-not $anyFeature) { $chkFeature.Text = 'No Windows feature backs these queues - nothing to disable.' }

    $updateCount = {
        $n = $clb.CheckedIndices.Count
        $lblCount.Text = "$n of $($clb.Items.Count) selected"
        $btnGo.Enabled = ($n -gt 0)
    }
    $showWhy = {
        $i = $clb.SelectedIndex
        if ($i -lt 0 -or $i -ge $rows.Count) { $txtWhy.Text = ''; return }
        $c = $rows[$i].Cand
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("$($c.Name)")
        [void]$sb.AppendLine("   driver : $($c.DriverName)")
        [void]$sb.AppendLine("   port   : $($c.PortName)")
        [void]$sb.AppendLine("   why    : $($c.Why)")
        if ($c.Feature) {
            $fs = switch ("$($c.FeatureState)") {
                'Enabled'  { 'ENABLED - the queue WILL come back unless this is disabled too' }
                'Disabled' { 'already disabled - the queue should not return' }
                'Absent'   { 'not present on this Windows edition' }
                default    { 'unknown (reading it needs elevation)' }
            }
            [void]$sb.AppendLine("   feature: $($c.Feature) - $fs")
        }
        if (@($c.IsDefaultFor).Count) {
            [void]$sb.AppendLine("   WARNING: this is the DEFAULT printer for $(@($c.IsDefaultFor) -join ', '). Removing it leaves Windows to pick a new default on its own.")
        }
        if ($c.Shared) { [void]$sb.AppendLine('   WARNING: this queue is SHARED - other machines may be pointed at it.') }
        $txtWhy.Text = $sb.ToString()
    }
    $clb.Add_SelectedIndexChanged($showWhy)
    $clb.Add_ItemCheck({ $f.BeginInvoke($updateCount) | Out-Null })

    $script:JunkDialogDone = $false
    $btnCancel.Add_Click({ $f.Close() })
    $btnGo.Add_Click({
        $picked = [System.Collections.Generic.List[object]]::new()
        foreach ($i in @($clb.CheckedIndices)) { $picked.Add($rows[$i].Cand) }
        if (-not $picked.Count) { return }

        $names = @($picked | ForEach-Object { $_.Name })
        $warn  = [System.Collections.Generic.List[string]]::new()
        foreach ($c in $picked) {
            if (@($c.IsDefaultFor).Count) { $warn.Add("  $($c.Name) is the DEFAULT for $(@($c.IsDefaultFor) -join ', ')") }
            if ($c.Shared) { $warn.Add("  $($c.Name) is SHARED") }
        }
        $ask = [System.Collections.Generic.List[string]]::new()
        $ask.Add('Remove these queues?')
        $ask.Add('')
        foreach ($n in $names) { $ask.Add("  $n") }
        if ($warn.Count) {
            $ask.Add('')
            $ask.Add('Note:')
            foreach ($w in $warn) { $ask.Add($w) }
        }
        if ($chkFeature.Checked) {
            $feats = @($picked | Where-Object { $_.Feature } | ForEach-Object { $_.Feature } | Select-Object -Unique)
            if ($feats.Count) {
                $ask.Add('')
                $ask.Add('The Windows feature(s) will also be disabled so they cannot return:')
                foreach ($ft in $feats) { $ask.Add("  $ft") }
                $ask.Add('That can take a minute and may report a pending restart.')
            }
        }
        if ((Show-Msg -Text ($ask -join "`r`n") -Title 'Confirm removal' -Buttons 'YesNo' -Icon 'Warning' -Default 'No') -ne 'Yes') { return }

        $btnGo.Enabled = $false; $btnCancel.Enabled = $false
        $f.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $res = Remove-JunkPrinter -Name $names -OnPoll $script:Pump
            foreach ($r in @($res.Results)) { Write-Log ("$($r.Name): $($r.Message)") $(if ($r.Success) { 'OK' } else { 'WARN' }) }

            $tail = [System.Collections.Generic.List[string]]::new()
            $tail.Add($res.Message)
            if ($chkFeature.Checked) {
                $feats = @($picked | Where-Object { $_.Feature } | ForEach-Object { $_.Feature } | Select-Object -Unique)
                foreach ($ft in $feats) {
                    $fr = Disable-PrintOptionalFeature -FeatureName $ft
                    Write-Log $fr.Message $(if ($fr.Success) { 'OK' } else { 'WARN' })
                    $tail.Add($fr.Message)
                }
            } elseif (@($picked | Where-Object { $_.ComesBack }).Count) {
                $tail.Add('')
                $tail.Add('NOTE: the XPS/Fax feature is still enabled, so Windows will recreate that queue on a servicing update or a repair. Run this again with the feature box ticked to make it stick.')
            }
            $script:JunkDialogDone = $true
            [void][System.Windows.Forms.MessageBox]::Show($f, ($tail -join "`r`n"),
                $(if ($res.Failed) { 'Finished with problems' } else { 'Junk queues removed' }), 'OK',
                $(if ($res.Failed) { 'Warning' } else { 'Information' }))
            $f.Close()
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show($f, "Could not finish: $($_.Exception.Message)", 'Not completed', 'OK', 'Warning')
        } finally {
            $f.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnGo.Enabled = $true; $btnCancel.Enabled = $true
        }
    })

    # Anchors and the first selection go in Add_Shown - set earlier they inflate
    # the controls (see the Repoint dialog for the same trap).
    $f.Add_Shown({
        if ($clb.Items.Count) { $clb.SelectedIndex = 0 }
        & $showWhy
        & $updateCount
    })
    [void]$f.ShowDialog((Get-MsgOwner))
    $f.Dispose()
    return $script:JunkDialogDone
}

function Show-InstallDriverDialog {
    <#
    .SYNOPSIS
        Install a printer driver from a folder of INFs, picking the model for you.
    .DESCRIPTION
        The driver has to exist BEFORE a local-port queue can be created, so this
        stands on its own rather than living inside the Repoint dialog.

        Point it at the folder the vendor package was extracted to and it reads
        every printer INF underneath, resolves the %tokens% and lists the model
        names those INFs actually declare - ranked against the printer you are
        chasing. That name has to match EXACTLY for the install to work, and it
        is rarely the filename or the marketing name: a Canon G6065 is installed
        by a driver called 'Canon G6000 series'.
    .OUTPUTS
        [string] driver name, or $null if nothing was installed.
    #>
    param([string]$SuggestName = '')

    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Install a printer driver'
    $f.Size = New-Object System.Drawing.Size(760, 520)
    $f.MinimumSize = New-Object System.Drawing.Size(660, 460)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'Sizable'
    $f.MinimizeBox = $false; $f.MaximizeBox = $false
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $top = New-DockBand $f 'Top' 104
    $bot = New-DockBand $f 'Bottom' 78
    $mid = New-Object System.Windows.Forms.Panel
    $mid.Dock = 'Fill'; $f.Controls.Add($mid); $mid.BringToFront()

    New-Lbl 'Point this at the folder the vendor driver was extracted to.' 10 8 $top $true | Out-Null
    $lblHint = New-Lbl 'Every printer INF underneath is read, and the model names they declare are listed below.' 10 27 $top
    $lblHint.ForeColor = [System.Drawing.Color]::DimGray

    New-Lbl 'Folder:' 10 60 $top | Out-Null
    $txtFolder = New-Object System.Windows.Forms.TextBox
    $txtFolder.Location = New-Object System.Drawing.Point(60, 57)
    $txtFolder.Size = New-Object System.Drawing.Size(430, 24)
    $top.Controls.Add($txtFolder)
    $btnBrowse = New-Btn 'Browse...' 498 56 90 27 $top
    $btnScan   = New-Btn 'Scan' 594 56 80 27 $top

    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Dock = 'Fill'
    $lst.IntegralHeight = $false
    $lst.Font = New-Object System.Drawing.Font('Consolas', 9)
    $mid.Controls.Add($lst)

    $lblPick = New-Lbl 'No folder scanned yet.' 10 8 $bot
    $lblPick.ForeColor = [System.Drawing.Color]::DimGray
    $lblPick.MaximumSize = New-Object System.Drawing.Size(720, 0)
    $btnInstall = New-Btn 'Install selected' 10 40 140 27 $bot
    $btnManual  = New-Btn 'Pick an INF myself...' 156 40 160 27 $bot
    $btnCancel  = New-Btn 'Cancel' 322 40 90 27 $bot

    $script:InfDlgResult = $null
    $script:InfCands = @()
    $btnInstall.Enabled = $false

    $showPick = {
        $i = $lst.SelectedIndex
        if ($i -lt 0 -or $i -ge @($script:InfCands).Count) { $btnInstall.Enabled = $false; return }
        $c = @($script:InfCands)[$i]
        $lblPick.Text = "Will register '$($c.DriverName)' from $($c.InfName)`r`n$($c.InfPath)"
        $btnInstall.Enabled = $true
    }
    $lst.Add_SelectedIndexChanged($showPick)

    $doScan = {
        $p = ([string]$txtFolder.Text).Trim().Trim('"')
        if (-not $p) { $lblPick.Text = 'Choose a folder first.'; return }
        if (-not (Test-Path -LiteralPath $p)) { $lblPick.Text = "Not found: $p"; return }
        $btnScan.Enabled = $false; $btnBrowse.Enabled = $false
        $f.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $lst.Items.Clear(); $btnInstall.Enabled = $false
        $lblPick.Text = 'Reading INFs...'
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $script:InfCands = @(Get-PrinterInfCandidate -Path $p -Match $SuggestName)
            foreach ($c in $script:InfCands) {
                [void]$lst.Items.Add(("{0,-46} {1}" -f $c.DriverName, $c.InfName))
            }
            if (@($script:InfCands).Count) {
                # Best match first, and pre-selected - but the whole list stays
                # visible, because a wrong guess must never hide the right answer.
                $lst.SelectedIndex = 0
                & $showPick
                if ($SuggestName -and @($script:InfCands)[0].Score -le 0) {
                    $lblPick.Text = "$(@($script:InfCands).Count) driver(s) found, but none look like '$SuggestName' - pick the right one yourself."
                }
            } else {
                $lblPick.Text = "No printer INFs found under that folder. Make sure the package is EXTRACTED, not still a .exe or .zip."
            }
        } catch {
            $lblPick.Text = "Scan failed: $($_.Exception.Message)"
        } finally {
            $f.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnScan.Enabled = $true; $btnBrowse.Enabled = $true
        }
    }
    $btnScan.Add_Click($doScan)

    $btnBrowse.Add_Click({
        $fb = New-Object System.Windows.Forms.FolderBrowserDialog
        $fb.Description = 'Pick the folder the printer driver was extracted to'
        if ($fb.ShowDialog($f) -eq 'OK') {
            $txtFolder.Text = $fb.SelectedPath
            & $doScan
        }
    })

    $install = {
        param([string]$InfPath, [string]$DriverName)
        $f.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $btnInstall.Enabled = $false; $btnManual.Enabled = $false
        try {
            $r = Install-PrinterDriverFromInf -InfPath $InfPath -DriverName $DriverName
            Write-Log $r.Message $(if ($r.Success) { 'OK' } else { 'WARN' })
            [void](Show-Msg -Text $r.Message -Title $(if ($r.Success) { 'Driver installed' } else { 'Driver not installed' }) -Icon $(if ($r.Success) { 'Information' } else { 'Warning' }))
            if ($r.Success) { $script:InfDlgResult = $DriverName; $f.Close() }
        } catch {
            [void](Show-Msg -Text "Install failed: $($_.Exception.Message)" -Title 'Driver not installed' -Icon 'Warning')
        } finally {
            $f.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnInstall.Enabled = $true; $btnManual.Enabled = $true
        }
    }

    $btnInstall.Add_Click({
        $i = $lst.SelectedIndex
        if ($i -lt 0 -or $i -ge @($script:InfCands).Count) { return }
        $c = @($script:InfCands)[$i]
        & $install $c.InfPath $c.DriverName
    })

    # The old behaviour, kept for an INF this parser cannot read.
    $btnManual.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = 'Printer driver INF (*.inf)|*.inf'
        $ofd.Title  = 'Pick the INF inside its ORIGINAL extracted folder'
        if ($ofd.ShowDialog($f) -ne 'OK') { return }
        $inf = $ofd.FileName
        $nm = Show-InputBox -Title 'Driver name' -Default $SuggestName -Prompt (
            "Exact driver name to register - it must match a model this INF declares.`n`nINF: $inf")
        if ([string]::IsNullOrWhiteSpace($nm)) { return }
        & $install $inf $nm.Trim()
    })

    $btnCancel.Add_Click({ $f.Close() })

    $f.Add_Shown({
        if ($SuggestName) { $lblHint.Text = "Looking for a driver that matches '$SuggestName'. The closest match is selected first." }
        $txtFolder.Focus() | Out-Null
    })
    [void]$f.ShowDialog((Get-MsgOwner))
    $f.Dispose()
    return $script:InfDlgResult
}
function Show-CredentialsDialog {
    param([string]$Server = '')

    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Print host credentials'
    $f.Size = New-Object System.Drawing.Size(720, 600)
    $f.MinimumSize = New-Object System.Drawing.Size(660, 520)
    $f.StartPosition = 'CenterParent'
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    function DL([string]$t, [int]$x, [int]$y, [bool]$bold = $false) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $t; $l.Location = New-Object System.Drawing.Point($x, $y); $l.AutoSize = $true
        if ($bold) { $l.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold) }
        $f.Controls.Add($l); $l
    }

    DL 'Host:' 12 16 $true | Out-Null
    $txtSrv = New-Object System.Windows.Forms.TextBox
    $txtSrv.Location = New-Object System.Drawing.Point(96, 13)
    $txtSrv.Size = New-Object System.Drawing.Size(220, 24)
    $txtSrv.Text = $Server
    $f.Controls.Add($txtSrv)

    DL 'Share (optional):' 330 16 | Out-Null
    $txtShare = New-Object System.Windows.Forms.TextBox
    $txtShare.Location = New-Object System.Drawing.Point(440, 13)
    $txtShare.Size = New-Object System.Drawing.Size(240, 24)
    $txtShare.Anchor = [System.Windows.Forms.AnchorStyles]'Top,Left,Right'
    $f.Controls.Add($txtShare)

    DL 'Username:' 12 50 $true | Out-Null
    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location = New-Object System.Drawing.Point(96, 47)
    $txtUser.Size = New-Object System.Drawing.Size(220, 24)
    $f.Controls.Add($txtUser)
    $lblUserHint = DL 'For a workgroup host use HOST\user - a bare name makes Windows try a domain logon.' 330 51
    $lblUserHint.ForeColor = [System.Drawing.Color]::DimGray

    DL 'Password:' 12 84 $true | Out-Null
    $txtPass = New-Object System.Windows.Forms.TextBox
    $txtPass.Location = New-Object System.Drawing.Point(96, 81)
    $txtPass.Size = New-Object System.Drawing.Size(220, 24)
    $txtPass.UseSystemPasswordChar = $true
    $f.Controls.Add($txtPass)
    $chkShow = New-Object System.Windows.Forms.CheckBox
    $chkShow.Text = 'Show'; $chkShow.Location = New-Object System.Drawing.Point(326, 83); $chkShow.AutoSize = $true
    $chkShow.Add_CheckedChanged({ $txtPass.UseSystemPasswordChar = (-not $chkShow.Checked) })
    $f.Controls.Add($chkShow)

    DL 'Save for user:' 12 118 $true | Out-Null
    $cboFor = New-Object System.Windows.Forms.ComboBox
    $cboFor.Location = New-Object System.Drawing.Point(96, 115)
    $cboFor.Size = New-Object System.Drawing.Size(300, 24)
    $cboFor.DropDownStyle = 'DropDownList'
    foreach ($p in @($script:Profiles)) { [void]$cboFor.Items.Add($p.User) }
    if ($cboFor.Items.Count) {
        $i = 0
        for ($x = 0; $x -lt @($script:Profiles).Count; $x++) { if (@($script:Profiles)[$x].Sid -eq $script:TargetSid) { $i = $x; break } }
        $cboFor.SelectedIndex = $i
    }
    $f.Controls.Add($cboFor)
    $lblForHint = DL 'A credential saved by SYSTEM lands in SYSTEM''s vault and does nothing for the user.' 404 119
    $lblForHint.ForeColor = [System.Drawing.Color]::DimGray

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = New-Object System.Drawing.Point(12, 152)
    $txtOut.Size = New-Object System.Drawing.Size(680, 350)
    $txtOut.Anchor = [System.Windows.Forms.AnchorStyles]'Top,Left,Right,Bottom'
    $txtOut.Multiline = $true; $txtOut.ReadOnly = $true; $txtOut.ScrollBars = 'Both'
    $txtOut.WordWrap = $false
    $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtOut.BackColor = [System.Drawing.Color]::White
    $f.Controls.Add($txtOut)

    $btnCheck  = New-Object System.Windows.Forms.Button
    $btnCheck.Text = 'Check access'; $btnCheck.Location = New-Object System.Drawing.Point(12, 516)
    $btnCheck.Size = New-Object System.Drawing.Size(120, 30)
    $btnCheck.Anchor = [System.Windows.Forms.AnchorStyles]'Left,Bottom'; $f.Controls.Add($btnCheck)

    $btnTest   = New-Object System.Windows.Forms.Button
    $btnTest.Text = 'Test credentials'; $btnTest.Location = New-Object System.Drawing.Point(138, 516)
    $btnTest.Size = New-Object System.Drawing.Size(130, 30)
    $btnTest.Anchor = [System.Windows.Forms.AnchorStyles]'Left,Bottom'; $f.Controls.Add($btnTest)

    $btnSave   = New-Object System.Windows.Forms.Button
    $btnSave.Text = 'Save'; $btnSave.Location = New-Object System.Drawing.Point(274, 516)
    $btnSave.Size = New-Object System.Drawing.Size(90, 30)
    $btnSave.Anchor = [System.Windows.Forms.AnchorStyles]'Left,Bottom'; $f.Controls.Add($btnSave)

    $btnRepair = New-Object System.Windows.Forms.Button
    $btnRepair.Text = 'Repair...'; $btnRepair.Location = New-Object System.Drawing.Point(370, 516)
    $btnRepair.Size = New-Object System.Drawing.Size(100, 30)
    $btnRepair.Anchor = [System.Windows.Forms.AnchorStyles]'Left,Bottom'; $f.Controls.Add($btnRepair)

    $btnClose  = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(598, 516)
    $btnClose.Size = New-Object System.Drawing.Size(94, 30)
    $btnClose.Anchor = [System.Windows.Forms.AnchorStyles]'Right,Bottom'
    $btnClose.DialogResult = 'Cancel'; $f.Controls.Add($btnClose)
    $f.CancelButton = $btnClose

    function Out-Line([string]$t) { $txtOut.AppendText($t + [Environment]::NewLine) }
    function Busy([bool]$on) {
        $f.Cursor = if ($on) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default }
        foreach ($b in @($btnCheck, $btnTest, $btnSave, $btnRepair)) { $b.Enabled = -not $on }
        [System.Windows.Forms.Application]::DoEvents()
    }

    $btnCheck.Add_Click({
        $s = $txtSrv.Text.Trim(); if (-not $s) { Out-Line 'Enter a host first.'; return }
        Busy $true
        try {
            # Pass whatever was typed. Without this the check ran as the console's
            # own identity and the credential fields had no effect on enumeration.
            $ta = @{ Server = $s; Share = ($txtShare.Text.Trim()) }
            if ($txtUser.Text.Trim()) {
                $ta.User = $txtUser.Text.Trim(); $ta.Password = $txtPass.Text
                # Hold them for THIS session so Enumerate on the Shares tab runs as
                # the same account. Previously only this dialog authenticated, which
                # is why it could see a host the Shares tab could not. In memory
                # only: never logged, never saved to disk, never in the report.
                if (-not $script:HostCred) { $script:HostCred = @{} }
                $script:HostCred[(ConvertTo-PrintHostKey $s)] = @{ User = $ta.User; Password = $ta.Password }
            }
            # Off the UI thread: the no-cache open inside this check legitimately takes many
            # seconds against a host that swallows connections, and a modal dialog cannot
            # use the background worker.
            $ar = Invoke-ToolkitResponsive -Command 'Test-PrintShareAccess' -Arguments $ta -TimeoutSec 150
            $a = $ar.Result
            if (-not $a) { Out-Line ('The check did not complete: ' + $(if ($ar.TimedOut) { 'it timed out.' } else { [string]$ar.Error })); return }
            Out-Line "=== $s ==="
            Out-Line ("  reachable    : {0}" -f $a.Reachable)
            Out-Line ("  can enumerate: {0}{1}" -f $a.CanEnumerate, $(if ($a.EnumeratedVia) { "  (via $($a.EnumeratedVia))" } else { '' }))
            if ($null -ne $a.CanConnect) { Out-Line ("  can connect  : {0}" -f $a.CanConnect) }
            if (@($a.Shares).Count) {
                Out-Line '  printer shares:'
                foreach ($sh in @($a.Shares)) { Out-Line ("     {0}" -f $sh.Unc) }
            }
            foreach ($n in @($a.Findings)) { Out-Line ("  * " + $n) }
            Out-Line ''
        } catch { Out-Line "Check failed: $($_.Exception.Message)" }
        finally { Busy $false }
    })

    $btnTest.Add_Click({
        $s = $txtSrv.Text.Trim(); $u = $txtUser.Text.Trim()
        if (-not $s -or -not $u) { Out-Line 'Enter a host and a username first.'; return }
        Busy $true
        try {
            $r = Test-PrinterHostCredential -Server $s -User $u -Password $txtPass.Text
            $verdict = if ($null -eq $r.Valid) { 'no verdict' } elseif ($r.Valid) { 'ACCEPTED' } else { 'REJECTED' }
            Out-Line "=== credential test: $u @ $s -> $verdict ==="
            Out-Line ("  " + $r.Message)
            if ($r.Win32Error) { Out-Line ("  windows error: {0}" -f $r.Win32Error) }
            Out-Line ''
        } catch { Out-Line "Test failed: $($_.Exception.Message)" }
        finally { Busy $false }
    })

    $btnSave.Add_Click({
        $s = $txtSrv.Text.Trim(); $u = $txtUser.Text.Trim(); $who = "$($cboFor.SelectedItem)"
        if (-not $s -or -not $u) { Out-Line 'Enter a host and a username first.'; return }
        if (-not $who) { Out-Line 'No user selected to save the credential for.'; return }
        if ([System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 
                "Save the credential for '$u' against '$s' in $who's Credential Manager?`n`nIt will apply to every connection they make to that host.",
                'Confirm save', 'YesNo', 'Question') -ne 'Yes') { return }
        Busy $true
        try {
            $r = Save-PrinterHostCredential -Server $s -User $u -Password $txtPass.Text -TargetUser $who -BothForms -OnPoll $script:Pump
            Out-Line "=== save -> $(if ($r.Success) { 'OK' } else { 'FAILED' }) ==="
            Out-Line ("  " + $r.Message)
            # Log the action, never the secret.
            Write-Log ("Credential for '$u' saved against '$s' for $who : " + $(if ($r.Success) { 'ok' } else { 'failed' })) $(if ($r.Success) { 'OK' } else { 'WARN' })

            if ($r.Success) {
                # A freshly stored credential is not picked up instantly, and it only
                # applies to a NEW session - so an immediate re-check can still show
                # the old failure and look like the save did nothing. Wait, then
                # re-check automatically rather than leaving that to be misread.
                Out-Line ''
                Out-Line '  Waiting for the credential to register, then re-checking...'
                Out-Line "  NOTE: this re-check runs as THIS CONSOLE, not as $who. To test $who,"
                Out-Line '  use Shares -> Test connection on the exact printer - that runs in their session.'
                for ($i = 0; $i -lt 12; $i++) { Start-Sleep -Milliseconds 250; [System.Windows.Forms.Application]::DoEvents() }
                $a2 = Test-PrintShareAccess -Server $s
                Out-Line ("  re-check: reachable={0}  can enumerate={1}{2}" -f $a2.Reachable, $a2.CanEnumerate, $(if ($a2.EnumeratedVia) { " (via $($a2.EnumeratedVia))" } else { '' }))
                if (@($a2.Shares).Count) {
                    foreach ($sh in @($a2.Shares)) { Out-Line ("     " + $sh.Unc) }
                } elseif (-not $a2.CanEnumerate) {
                    Out-Line '  Still cannot list shares. The saved credential applies to NEW connections,'
                    Out-Line '  so an existing session to the host can keep the old identity alive - use'
                    Out-Line '  Repair to clear it, then check again. Listing may also simply be denied to'
                    Out-Line '  this account, which does not stop printing: test a share by name instead.'
                }
            }
            Out-Line ''
        } catch { Out-Line "Save failed: $($_.Exception.Message)" }
        finally { Busy $false }
    })

    $btnRepair.Add_Click({
        $s = $txtSrv.Text.Trim(); $who = "$($cboFor.SelectedItem)"
        if (-not $s) { Out-Line 'Enter a host first.'; return }
        if (-not $who) { Out-Line 'No user selected to repair.'; return }
        if ([System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 
                "Clear saved credentials AND existing connections to '$s' for $who ?`n`nEvery connection to that host shares one authenticated session, so any open file or drive on it will be dropped too.",
                'Confirm repair', 'YesNo', 'Warning') -ne 'Yes') { return }
        Busy $true
        try {
            $r = Repair-PrinterHostCredential -Server $s -TargetUser $who -OnPoll $script:Pump
            Out-Line "=== repair -> $(if ($r.Success) { 'clean' } else { 'incomplete' }) ==="
            foreach ($l in @($r.Log)) { Out-Line ("  " + $l) }
            if (@($r.Survivors).Count) {
                Out-Line '  SURVIVING CONNECTIONS:'
                foreach ($sv in @($r.Survivors)) { Out-Line ("     " + $sv) }
            }
            Out-Line ("  " + $r.Message)
            Out-Line ''
            Write-Log ("Credential repair for '$s' ($who): " + $r.Message) $(if ($r.Success) { 'OK' } else { 'WARN' })
        } catch { Out-Line "Repair failed: $($_.Exception.Message)" }
        finally { Busy $false }
    })

    $f.Add_Shown({
        Out-Line 'Check access asks three separate questions: is the host reachable, can its'
        Out-Line 'shares be LISTED, and can a named printer share be OPENED.'
        Out-Line ''
        Out-Line 'They are independent. "net view ... System error 5" means listing was refused -'
        Out-Line 'that right is usually admin-only, and is commonly restricted further on a domain'
        Out-Line 'controller. It does not by itself stop anyone printing.'
        Out-Line ''
    })

    [void]$f.ShowDialog()
    $f.Dispose()
}

$btnCreds.Add_Click({
    Show-CredentialsDialog -Server ($txtServer.Text.Trim())
})

$btnDrvInfo.Add_Click({
    $s = $txtServer.Text.Trim()
    if (-not $s) { [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 'Enter a print host first.', 'Printer Manager', 'OK', 'Information'); return }
    $share = ''
    if ($dgvServer.SelectedRows.Count -gt 0) {
        $i = $dgvServer.SelectedRows[0].Index
        if ($i -ge 0 -and $i -lt @($script:ServerRows).Count) { $share = "$(@($script:ServerRows)[$i].ShareName)" }
    }
    Set-Busy $true "Reading driver details from $s..."
    try {
        $d = Get-RemotePrinterDriverInfo -Server $s -ShareName $share
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine($d.Message)
        [void]$sb.AppendLine('')
        if ($d.DriverName) {
            [void]$sb.AppendLine("Driver name : $($d.DriverName)")
            [void]$sb.AppendLine("  use this EXACT name when installing it here")
            [void]$sb.AppendLine("INF on host : $($d.InfPath)")
            [void]$sb.AppendLine("Model       : v$($d.MajorVersion), package-aware=$($d.PackageAware), $($d.Environment)")
            if ($d.Manufacturer) { [void]$sb.AppendLine("Manufacturer: $($d.Manufacturer)") }
            [void]$sb.AppendLine('')
        }
        [void]$sb.AppendLine("print`$ readable from here : $($d.PrintShareReadable)")
        foreach ($p in @($d.DriverPaths)) { [void]$sb.AppendLine("   driver files: $p") }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('To get the package off the host, run THERE as administrator:')
        foreach ($c in @($d.ExportCommands)) { [void]$sb.AppendLine("   $c") }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Then copy that folder here and use Install driver... on the INF inside it.')
        [void]$sb.AppendLine('Keep the whole folder - printer INFs depend on sibling files and staging')
        [void]$sb.AppendLine('fails if the INF is copied out on its own.')
        if (@($d.AllDrivers).Count -and -not $d.DriverName) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('Drivers on that host:')
            foreach ($x in @($d.AllDrivers)) { [void]$sb.AppendLine("   $($x.Name)   [$($x.InfPath)]") }
        }
        $text = $sb.ToString()
        try { Set-Clipboard -Value $text } catch {}
        Write-Log "Read driver details from $s$(if ($d.DriverName) { " : $($d.DriverName)" } else { '' })" 'INFO'
        [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), $text + "`n(copied to the clipboard)", "Driver on $s", 'OK', 'Information')
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), "Could not read driver details: $($_.Exception.Message)", 'Driver info', 'OK', 'Warning')
    } finally { Set-Busy $false }
})

$btnInstallDrv.Add_Click({
    # Pre-fill from the selected remote queue when there is one - that name is
    # usually exactly what the INF declares.
    $guess = ''
    if ($dgvServer.SelectedRows.Count -gt 0) {
        $i = $dgvServer.SelectedRows[0].Index
        if ($i -ge 0 -and $i -lt @($script:ServerRows).Count) { $guess = "$(@($script:ServerRows)[$i].DriverName)" }
    }
    [void](Show-InstallDriverDialog -SuggestName $guess)
})

# --------------------------------------------------------------------------
# Report - dump everything currently on screen to a text file and open Notepad.
#
# Renders straight from the DataGridViews rather than re-querying, so the report
# is exactly what the technician is looking at - same rows, same order, same
# values - and it costs nothing extra to produce.
# --------------------------------------------------------------------------
function Format-GridAsText {
    param($Grid, [int]$MaxCell = 60)

    if (-not $Grid -or $Grid.Rows.Count -eq 0) { return @('  (none)') }

    $cols = @($Grid.Columns | Where-Object { $_.Visible })
    $head = @($cols | ForEach-Object { [string]$_.HeaderText })
    $rows = @()
    foreach ($r in $Grid.Rows) {
        if ($r.IsNewRow) { continue }
        $rows += ,@($cols | ForEach-Object {
            $v = "$($r.Cells[$_.Index].Value)"
            if ($v.Length -gt $MaxCell) { $v = $v.Substring(0, $MaxCell - 1) + '~' }
            $v
        })
    }

    # Width each column to its widest value so the columns line up in Notepad.
    $w = @()
    for ($c = 0; $c -lt $head.Count; $c++) {
        $max = $head[$c].Length
        foreach ($row in $rows) { if ($row[$c].Length -gt $max) { $max = $row[$c].Length } }
        $w += $max
    }

    $out = New-Object System.Collections.Generic.List[string]
    $line = ''
    for ($c = 0; $c -lt $head.Count; $c++) { $line += $head[$c].PadRight($w[$c]) + '  ' }
    $out.Add('  ' + $line.TrimEnd())
    $line = ''
    for ($c = 0; $c -lt $head.Count; $c++) { $line += ('-' * $w[$c]) + '  ' }
    $out.Add('  ' + $line.TrimEnd())
    foreach ($row in $rows) {
        $line = ''
        for ($c = 0; $c -lt $head.Count; $c++) { $line += $row[$c].PadRight($w[$c]) + '  ' }
        $out.Add('  ' + $line.TrimEnd())
    }
    $out.ToArray()
}

function New-PrinterReportText {
    $sb = New-Object System.Text.StringBuilder
    # NOTE: these helpers must not be single letters. PowerShell resolves aliases
    # BEFORE functions, so a function named H was shadowed by the built-in 'h'
    # alias for Get-History and the report died with
    # "Cannot bind parameter 'Id' ... convert TARGET USER to System.Int64".
    function Add-Head([string]$t) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('=' * 78)
        [void]$sb.AppendLine('  ' + $t)
        [void]$sb.AppendLine('=' * 78)
    }
    function Add-Line([string]$t = '') { [void]$sb.AppendLine($t) }
    function Add-Grid($Grid) { foreach ($x in (Format-GridAsText $Grid)) { Add-Line $x } }

    [void]$sb.AppendLine('#' * 78)
    [void]$sb.AppendLine('  PRINTER REPORT')
    [void]$sb.AppendLine("  Host      : $($script:Ctx.Host)")
    [void]$sb.AppendLine("  Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  by Printer Manager v$($script:Ctx.ToolkitVersion)")
    [void]$sb.AppendLine("  Context   : $($script:Ctx.CurrentName)  [$(if ($script:IsSystem) { 'SYSTEM' } elseif ($script:Ctx.IsAdmin) { 'elevated user' } else { 'standard user' })]")
    [void]$sb.AppendLine("  OS        : $($script:Ctx.OS)")
    [void]$sb.AppendLine("  Domain    : $(if ($script:Ctx.PartOfDomain) { $script:Ctx.Domain } else { "$($script:Ctx.Domain) (workgroup)" })")
    [void]$sb.AppendLine('#' * 78)

    Add-Head 'TARGET USER'
    if ($script:TargetUser) {
        $map = Get-SelectedUserMap
        Add-Line "  User            : $($script:TargetUser)"
        Add-Line "  SID             : $($script:TargetSid)"
        if ($map) {
            Add-Line "  Default printer : $($map.DefaultPrinter)"
            $legacy = $map.LegacyDefaultPrinterMode
            Add-Line "  Windows manages the default : $(if ($null -eq $legacy -or [int]$legacy -ne 1) { "YES - it will overwrite this default after the next print (LegacyDefaultPrinterMode=$legacy)" } else { 'no (sticky)' })"
            Add-Line "  Mapped printers : $(@($map.Connections).Count)"
        }
    } else {
        Add-Line '  (no user selected - nobody signed in)'
    }

    Add-Head 'PRINTERS'
    Add-Grid $dgvPrinters

    Add-Head 'DIAGNOSTIC FINDINGS'
    if (@($script:Findings).Count -eq 0) {
        Add-Line '  (diagnostics not run - use the Diagnose tab)'
    } else {
        Add-Grid $dgvFindings
        Add-Line ''
        Add-Line '  DETAIL'
        foreach ($f in @($script:Findings)) {
            Add-Line ''
            Add-Line "  $($f.Code)  [$($f.Severity)]  scope: $($f.Scope)"
            if ($f.User)    { Add-Line "    User    : $($f.User)" }
            if ($f.Printer) { Add-Line "    Printer : $($f.Printer)" }
            if ($f.Server)  { Add-Line "    Server  : $($f.Server)" }
            Add-Line "    Found   : $($f.Detail)"
            if ($f.LikelyErrorCode) { Add-Line "    Seen as : $($f.LikelyErrorCode)" }
            if ($f.Fix)             { Add-Line "    Fix     : $($f.Fix)" }
        }
    }

    Add-Head 'SPOOLER AND JOBS'
    $h = $script:SpoolerHealth
    if ($h) {
        Add-Line "  Spooler        : $($h.Status) ($($h.StartMode)), account $($h.Account)"
        Add-Line "  Spool folder   : $($h.SpoolDirectory)"
        Add-Line "  Spool files    : $(if ($h.SpoolReadable) { "$($h.SpoolFileCount)$(if ($h.OldestSpoolAgeHours) { "  (oldest $($h.OldestSpoolAgeHours)h)" } else { '' })" } else { 'not readable from this context' })"
    }
    Add-Line ''
    Add-Line '  Queued jobs:'
    Add-Grid $dgvJobs
    if ("$($txtEvents.Text)".Trim()) {
        Add-Line ''
        Add-Line '  Recent PrintService / spooler events:'
        foreach ($ln in ("$($txtEvents.Text)" -split "`r?`n")) { if ($ln.Trim()) { Add-Line "    $ln" } }
    }

    Add-Head 'POINT AND PRINT POLICY'
    Add-Line '  An UNSET value is NOT an unenforced value - State shows the effective value.'
    Add-Line ''
    Add-Grid $dgvPolicy

    Add-Head 'SHARE HOST (this PC)'
    if (-not $script:HostHealth) {
        Add-Line '  (not checked - use the Share Host tab)'
    } elseif (-not $script:HostHealth.IsHost) {
        Add-Line '  This PC does not share any printer.'
    } else {
        Add-Line "  $($script:HostHealth.Summary)"
        Add-Line ''
        Add-Grid $dgvHost
        Add-Line ''
        foreach ($ln in ("$($txtHost.Text)" -split "`r?`n")) { if ($ln.Trim()) { Add-Line "  $ln" } }
    }

    Add-Head 'SHARES (remote)'
    Add-Grid $dgvServer
    if ("$($lblServerNote.Text)".Trim()) { Add-Line ''; Add-Line "  $($lblServerNote.Text)" }

    Add-Head 'SESSION TRANSCRIPT'
    foreach ($ln in ("$($txtLog.Text)" -split "`r?`n")) { if ($ln.Trim()) { Add-Line "  $ln" } }

    Add-Line ''
    Add-Line '--- end of report ---'
    $sb.ToString()
}

$btnReport.Add_Click({
    Set-Busy $true 'Building report...'
    try {
        $text = New-PrinterReportText
        $name = "PrinterReport_$($script:Ctx.Host)_$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
        # Public\Documents rather than the SYSTEM profile's TEMP: from Backstage the
        # tech needs to be able to find and copy the file afterwards.
        $dir = Join-Path $env:PUBLIC 'Documents'
        if (-not (Test-Path -LiteralPath $dir)) { $dir = $env:TEMP }
        $path = Join-Path $dir $name
        Set-Content -LiteralPath $path -Value $text -Encoding UTF8
        Write-Log "Report written to $path" 'OK'
        try {
            Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$path`"" -ErrorAction Stop
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), 
                "The report was saved but Notepad could not be started:`n$($_.Exception.Message)`n`nFile:`n$path",
                'Report saved', 'OK', 'Warning')
        }
    } catch {
        Write-Log ("Report failed: " + $_.Exception.Message) 'ERROR'
        [void][System.Windows.Forms.MessageBox]::Show((Get-MsgOwner), "Could not build the report:`n`n$($_.Exception.Message)", 'Report failed', 'OK', 'Warning')
    } finally { Set-Busy $false }
})

# --------------------------------------------------------------------------
# Layout robustness for small Backstage sessions
# --------------------------------------------------------------------------
foreach ($c in @($form.Controls)) {
    if ($c -is [System.Windows.Forms.Button] -or $c -is [System.Windows.Forms.TextBox] -or
        $c -is [System.Windows.Forms.ComboBox] -or $c -is [System.Windows.Forms.TabControl]) { $c.BringToFront() }
}
$status.BringToFront()
# Not needed now the form is fully docked, and it produced stray scrollbars.
$form.AutoScroll = $false
try {
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    if ($form.Height -gt $wa.Height) { $form.Height = [Math]::Max(560, $wa.Height - 20) }
    if ($form.Width  -gt $wa.Width)  { $form.Width  = [Math]::Max(820, $wa.Width  - 20) }
    if ($form.MinimumSize.Height -gt $form.Height -or $form.MinimumSize.Width -gt $form.Width) {
        $form.MinimumSize = New-Object System.Drawing.Size([Math]::Min(820, $form.Width), [Math]::Min(560, $form.Height))
    }
} catch {}

$form.Add_FormClosing({
    try { if ($script:Timer) { $script:Timer.Stop() } } catch {}
    try { if ($script:Ps) { $script:Ps.Dispose() } } catch {}
    try { if ($script:Rs) { $script:Rs.Close(); $script:Rs.Dispose() } } catch {}
})

# Seed the static tabs, then kick off the first scan.
$txtCodes.Text = (@(Get-PrinterErrorCodeHelp) | ForEach-Object {
    "$($_.Code)  $($_.Title)" + [Environment]::NewLine +
    "    cause : $($_.Cause)" + [Environment]::NewLine +
    "    fix   : $($_.Fix)" + [Environment]::NewLine
}) -join [Environment]::NewLine

Write-Log "Printer Manager v$($script:Ctx.ToolkitVersion) on $($script:Ctx.Host) - running as $($script:Ctx.CurrentName)." 'INFO'
if ($script:IsSystem) {
    Write-Log 'Running as SYSTEM: per-user actions are injected into the target user''s session.' 'INFO'
}
# Credentials supplied during THIS session, keyed by host. Memory only - the
# window clears them on close so nothing outlives the console.
$script:HostCred = @{}
$form.Add_FormClosed({ try { $script:HostCred.Clear() } catch {} })
# Follow-up work an Apply-* step queues, for the timer to start once Busy is released.
$script:PendingRemedy = $null
$script:PendingDelivery = $null
$script:PendingScan = $false

Update-RpcStateLabels
Update-JunkLabel
$form.Add_Shown({ Start-Worker 'scan' })

[void]$form.ShowDialog()
$form.Dispose()
