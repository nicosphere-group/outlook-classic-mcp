# Runs the Outlook helper against a stand-in for Outlook, under Windows PowerShell 5.1.
# Usage (Windows):  powershell -NoProfile -ExecutionPolicy Bypass -File tests\helper.tests.ps1
# Checks the helper's syntax, its request/answer format and its logic. It cannot check
# real Outlook; that needs a PC with Outlook.

$ErrorActionPreference = 'Stop'
$helperPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'src\outlook-helper.ps1'

# ---- 1. the helper must parse cleanly on this PowerShell ----
$parseErrors = $null
$tokens = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($helperPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors -and $parseErrors.Count -gt 0) {
    $parseErrors | ForEach-Object { Write-Host ("PARSE ERROR line {0}: {1}" -f $_.Extent.StartLineNumber, $_.Message) }
    exit 1
}
Write-Host ("Parsed OK on PowerShell {0}" -f $PSVersionTable.PSVersion)

# ---- 2. a small stand-in for the Outlook object model ----

class FakeList : System.Collections.IEnumerable {
    [object[]]$All = @()
    [int]$Count = 0
    [void] Push([object]$x) { $this.All += $x; $this.Count = $this.All.Length }
    [object] Item([object]$key) {
        if ($key -is [int]) { return $this.All[$key - 1] }   # 1-based, like Outlook
        foreach ($x in $this.All) { if ($x.Name -eq $key) { return $x } }
        throw "no such item: $key"
    }
    [System.Collections.IEnumerator] GetEnumerator() { return $this.All.GetEnumerator() }
    [bool] ResolveAll() { return $true }
    [object] Add([string]$path) {
        $this.Push([pscustomobject]@{ FileName = [System.IO.Path]::GetFileName($path); DisplayName = ''; Size = 10 })
        return 'leaked-return-value'
    }
}

class FakeItems {
    [object[]]$All = @()
    [int]$Count = 0
    [int]$Position = 0
    FakeItems([object[]]$items) { $this.All = $items; $this.Count = $items.Length }
    [void] Sort([string]$field, [bool]$descending) {
        $this.All = @($this.All | Sort-Object -Property ReceivedTime -Descending)
    }
    [object] GetFirst() { $this.Position = 0; return $this.GetNext() }
    [object] GetNext() {
        if ($this.Position -ge $this.All.Length) { return $null }
        $item = $this.All[$this.Position]
        $this.Position++
        return $item
    }
    [object] Restrict([string]$filter) {
        if ($filter -eq '[UnRead] = True') { return [FakeItems]::new(@($this.All | Where-Object { $_.UnRead })) }
        $m = [regex]::Match($filter, "like '%(.+?)%'")
        if ($m.Success) {
            $needle = $m.Groups[1].Value
            return [FakeItems]::new(@($this.All | Where-Object { $_.Subject -like "*$needle*" }))
        }
        return [FakeItems]::new($this.All)
    }
}

$script:registry = @{}
$script:nextEntry = 0
$script:storeId = 'STORE-0001'

function New-FakeRecipients($mail) {
    $list = [FakeList]::new()
    $type = 0
    foreach ($field in @($mail.To, $mail.CC, $mail.BCC)) {
        $type++
        foreach ($address in ([string]$field).Split(';')) {
            $a = $address.Trim()
            if ($a.Length -gt 0) { $list.Push([pscustomobject]@{ Name = $a; Address = $a; Type = $type }) }
        }
    }
    return , $list
}

function New-FakeMail([string]$subject, [string]$body, [datetime]$received, [bool]$unread, [bool]$sent) {
    $mail = [pscustomobject]@{
        Class = 43; Sent = $sent; Subject = $subject; Body = $body; HTMLBody = ''; BodyFormat = 1
        SenderName = 'Tanaka'; SenderEmailAddress = 'tanaka@example.co.jp'; SenderEmailType = 'SMTP'
        UnRead = $unread; ReceivedTime = $received; LastModificationTime = $received
        EntryID = ''; Parent = $null; To = ''; CC = ''; BCC = ''; SendUsingAccount = $null
        Attachments = [FakeList]::new(); Displayed = $false
    }
    $mail | Add-Member -MemberType ScriptProperty -Name Recipients -Value { New-FakeRecipients $this }
    $mail | Add-Member -MemberType ScriptMethod -Name Save -Value {
        if ($this.EntryID -eq '') {
            $script:nextEntry++
            $this.EntryID = 'ENTRY{0:D4}' -f $script:nextEntry
            $this.Parent = $script:draftsFolder
            $script:registry[$this.EntryID] = $this
        }
    }
    $mail | Add-Member -MemberType ScriptMethod -Name Send -Value { $this.Sent = $true; return $true }
    $mail | Add-Member -MemberType ScriptMethod -Name Display -Value { param($modal) $this.Displayed = $true }
    $mail | Add-Member -MemberType ScriptMethod -Name Reply -Value {
        $r = New-FakeMail ('RE: ' + $this.Subject) ('> ' + $this.Body) (Get-Date) $false $false
        $r.To = $this.SenderEmailAddress
        $r.BodyFormat = 2
        $r.HTMLBody = '<html><body lang="ja"><p>original</p></body></html>'
        return $r
    }
    $mail | Add-Member -MemberType ScriptMethod -Name ReplyAll -Value { return $this.Reply() }
    $mail | Add-Member -MemberType ScriptMethod -Name Forward -Value {
        return (New-FakeMail ('FW: ' + $this.Subject) ('> ' + $this.Body) (Get-Date) $false $false)
    }
    return $mail
}

function New-FakeFolder([string]$name, [object[]]$mails) {
    $folder = [pscustomobject]@{
        Name = $name; FolderPath = ('\\me@example.com\' + $name); DefaultItemType = 0
        StoreID = $script:storeId; Folders = [FakeList]::new(); Mails = $mails
        UnReadItemCount = @($mails | Where-Object { $_.UnRead }).Count
    }
    $folder | Add-Member -MemberType ScriptProperty -Name Items -Value { [FakeItems]::new($this.Mails) }
    foreach ($m in $mails) {
        $script:nextEntry++
        $m.EntryID = 'ENTRY{0:D4}' -f $script:nextEntry
        $m.Parent = $folder
        $script:registry[$m.EntryID] = $m
    }
    return $folder
}

$jp = [string][char]0x4EF6 + [char]0x540D   # two Japanese characters, built without non-ASCII source
$inboxMails = @(
    (New-FakeMail ('old ' + $jp) 'old body' ([datetime]'2026-01-05T08:00:00') $false $true),
    (New-FakeMail ('new ' + $jp) 'new body' ([datetime]'2026-01-20T09:30:00') $true $true),
    (New-FakeMail 'middle' 'middle body' ([datetime]'2026-01-10T12:00:00') $true $true)
)
$inboxFolder = New-FakeFolder 'Inbox' $inboxMails
$script:draftsFolder = New-FakeFolder 'Drafts' @()
$root = [pscustomobject]@{ Folders = [FakeList]::new() }
$root.Folders.Push($inboxFolder)
$root.Folders.Push($script:draftsFolder)

$store = [pscustomobject]@{ StoreID = $script:storeId; DisplayName = 'me@example.com'; Root = $root }
$store | Add-Member -MemberType ScriptMethod -Name GetRootFolder -Value { return $this.Root }
$store | Add-Member -MemberType ScriptMethod -Name GetDefaultFolder -Value {
    param($type)
    if ($type -eq 6) { return $inboxFolder }
    if ($type -eq 16) { return $script:draftsFolder }
    throw 'no such folder'
}
$account = [pscustomobject]@{ DisplayName = 'Me'; SmtpAddress = 'me@example.com'; AccountType = 2; DeliveryStore = $store }

$ns = [pscustomobject]@{ Stores = [FakeList]::new(); Accounts = [FakeList]::new(); DefaultStore = $store }
$ns.Stores.Push($store)
$ns.Accounts.Push($account)
$ns | Add-Member -MemberType ScriptMethod -Name GetItemFromID -Value {
    param($entry, $storeId)
    if ($script:registry.ContainsKey($entry)) { return $script:registry[$entry] }
    throw 'not found'
}
$fakeApp = [pscustomobject]@{ Ns = $ns }
$fakeApp | Add-Member -MemberType ScriptMethod -Name GetNamespace -Value { param($n) return $this.Ns }
$fakeApp | Add-Member -MemberType ScriptMethod -Name CreateItem -Value {
    param($t) return (New-FakeMail '' '' (Get-Date) $false $false)
}

# ---- 3. load the helper (without its read loop) and point it at the stand-in ----
. $helperPath -NoLoop
$script:app = $fakeApp

$script:failures = 0
function Check([string]$name, $actual, $expected) {
    if ("$actual" -ceq "$expected") {
        Write-Host "ok   $name"
    } else {
        Write-Host "FAIL $name : expected [$expected] but got [$actual]"
        $script:failures++
    }
}
function Ask([string]$op, $arguments) {
    $line = ConvertTo-Json -InputObject @{ id = 7; op = $op; args = $arguments } -Depth 8 -Compress
    $raw = Invoke-Request $line
    if ($raw -isnot [string] -or $raw.Contains("`n")) { Write-Host "FAIL $op : answer is not one line"; $script:failures++ }
    return (ConvertFrom-Json -InputObject $raw)
}

# request / answer format
$r = Ask 'ping' @{ echo = $jp }
Check 'ping ok' $r.ok $true
Check 'ping id' $r.id 7
Check 'ping echo keeps Japanese' $r.result.echo $jp
$r = Ask 'no_such_op' @{}
Check 'unknown op is an error' $r.ok $false
Check 'unknown op kind' $r.kind 'invalid'
$bad = ConvertFrom-Json -InputObject (Invoke-Request 'this is not json')
Check 'garbage input is an error, not a crash' $bad.ok $false

# reading
$r = Ask 'list_accounts' @{}
Check 'accounts count' @($r.result.accounts).Count 1
Check 'account type' $r.result.accounts[0].type 'pop3'
Check 'account default' $r.result.accounts[0].is_default $true
Check 'data files' @($r.result.data_files)[0] 'me@example.com'

$r = Ask 'list_folders' @{ account = $null }
Check 'folders count' @($r.result.folders).Count 2
Check 'folder path has no leading backslashes' $r.result.folders[0].path 'me@example.com\Inbox'
Check 'folder unread' $r.result.folders[0].unread 2

$r = Ask 'list_emails' @{ folder = 'inbox'; account = $null; since = $null; until = $null; unread_only = $false; limit = 2; offset = 0 }
Check 'list ok' $r.ok $true
Check 'list newest first' $r.result.emails[0].subject ('new ' + $jp)
Check 'list limit' @($r.result.emails).Count 2
Check 'list has_more' $r.result.has_more $true
Check 'list received has an offset' ($r.result.emails[0].received -match '^2026-01-20T09:30:00[+-]\d\d:\d\d$') $true
$firstId = [string]$r.result.emails[0].id
Check 'id has a store key and an entry id' ($firstId -match '^[0-9a-f]{10}\.ENTRY\d{4}$') $true

$offsetText = (Get-Date).ToString('zzz')
$r = Ask 'list_emails' @{ folder = 'inbox'; account = $null; since = ('2026-01-08T00:00:00' + $offsetText); until = ('2026-01-15T23:59:59' + $offsetText); unread_only = $false; limit = 20; offset = 0 }
Check 'date window finds exactly the middle mail' @($r.result.emails).Count 1
Check 'date window subject' $r.result.emails[0].subject 'middle'

$r = Ask 'list_emails' @{ folder = 'inbox'; account = $null; since = $null; until = $null; unread_only = $true; limit = 20; offset = 1 }
Check 'unread + offset' @($r.result.emails).Count 1

$r = Ask 'search_emails' @{ folder = 'inbox'; account = $null; filter = "@SQL=""urn:schemas:httpmail:subject"" like '%middle%'"; since = $null; until = $null; limit = 20 }
Check 'search' @($r.result.emails).Count 1

$r = Ask 'list_emails' @{ folder = 'me@example.com\Inbox'; account = $null; since = $null; until = $null; unread_only = $false; limit = 20; offset = 0 }
Check 'folder by path' @($r.result.emails).Count 3
$r = Ask 'list_emails' @{ folder = 'me@example.com\Nowhere'; account = $null; since = $null; until = $null; unread_only = $false; limit = 20; offset = 0 }
Check 'missing folder kind' $r.kind 'not_found'

$r = Ask 'read_email' @{ id = $firstId }
Check 'read subject' $r.result.subject ('new ' + $jp)
Check 'read body' $r.result.body 'new body'
Check 'read is_draft' $r.result.is_draft $false
Check 'read from' $r.result.from.address 'tanaka@example.co.jp'
$r = Ask 'read_email' @{ id = 'nodot' }
Check 'bad id kind' $r.kind 'invalid'
$r = Ask 'read_email' @{ id = 'ffffffffff.ENTRY0001' }
Check 'unknown store kind' $r.kind 'not_found'

# drafting
$r = Ask 'create_draft' @{ to = @('a@example.com', 'b@example.com'); cc = @(); bcc = @('c@example.org'); subject = ('draft ' + $jp); body = 'hello'; account = $null; attachments = @('C:\docs\report.pdf') }
Check 'create ok' $r.ok $true
$draftId = [string]$r.result.id
Check 'create returns an id even though Attachments.Add leaked a value' ($draftId -match '^[0-9a-f]{10}\.ENTRY\d{4}$') $true

$r = Ask 'get_draft' @{ id = $draftId }
Check 'draft to count' @($r.result.to).Count 2
Check 'draft cc is an empty list' @($r.result.cc).Count 0
Check 'draft bcc' $r.result.bcc[0].address 'c@example.org'
Check 'draft subject' $r.result.subject ('draft ' + $jp)
Check 'draft sender' $r.result.sender_address 'me@example.com'
Check 'draft account type' $r.result.account_type 'pop3'
Check 'draft attachment' $r.result.attachments[0].filename 'report.pdf'
$raw = Invoke-Request (ConvertTo-Json -InputObject @{ id = 1; op = 'get_draft'; args = @{ id = $draftId } } -Compress)
Check 'a one-item list is still written as a list' ($raw -match '"bcc":\[\{') $true
Check 'an empty list is written as a list' ($raw -match '"cc":\[\]') $true

$r = Ask 'update_draft' @{ id = $draftId; set_to = $true; to = @('only@example.com'); set_cc = $false; cc = @(); set_bcc = $false; bcc = @(); set_subject = $false; subject = ''; set_body = $true; body = 'changed'; add_attachments = @() }
Check 'update ok' $r.ok $true
$r = Ask 'get_draft' @{ id = $draftId }
Check 'update replaced to' @($r.result.to).Count 1
Check 'update kept bcc' @($r.result.bcc).Count 1
Check 'update kept subject' $r.result.subject ('draft ' + $jp)
Check 'update changed body' $r.result.body 'changed'

$r = Ask 'get_draft' @{ id = $firstId }
Check 'a received mail is not a draft' $r.kind 'not_draft'
$r = Ask 'send' @{ id = $firstId }
Check 'a received mail cannot be sent' $r.kind 'not_draft'
$r = Ask 'update_draft' @{ id = $firstId; set_to = $false; to = @(); set_cc = $false; cc = @(); set_bcc = $false; bcc = @(); set_subject = $true; subject = 'x'; set_body = $false; body = ''; add_attachments = @() }
Check 'a received mail cannot be edited' $r.kind 'not_draft'

$r = Ask 'create_reply_draft' @{ id = $firstId; text = 'thanks'; html_block = '<div>thanks</div><br>'; reply_all = $false }
Check 'reply ok' $r.ok $true
$entry = ([string]$r.result.id).Split('.')[1]
Check 'reply text goes above the quote, inside body' $script:registry[$entry].HTMLBody '<html><body lang="ja"><div>thanks</div><br><p>original</p></body></html>'
$r = Ask 'create_forward_draft' @{ id = $firstId; to = @('f@example.com'); text = 'fyi'; html_block = '<div>fyi</div><br>' }
Check 'forward ok' $r.ok $true
$entry = ([string]$r.result.id).Split('.')[1]
Check 'forward text on top of a plain-text mail' ($script:registry[$entry].Body.StartsWith('fyi')) $true

$r = Ask 'display' @{ id = $draftId }
Check 'display ok' $r.ok $true
Check 'display did not send' $script:registry[$draftId.Split('.')[1]].Sent $false
$r = Ask 'send' @{ id = $draftId }
Check 'send ok' $r.ok $true
Check 'send marked sent' $script:registry[$draftId.Split('.')[1]].Sent $true
$r = Ask 'send' @{ id = $draftId }
Check 'cannot send twice' $r.kind 'not_draft'

# the confirmation window (needs a desktop session; skipped where there is none)
if ([Environment]::UserInteractive) {
    try {
        $view = @{
            title = 'Send this email?'; heading = 'Do you want to send this email?'; subheading = 'Check the details.'
            warning_lines = @('There are 12 recipients.')
            rows = @(
                @{ label = 'To'; value = 'a@example.com'; empty = $false; emphasis = $true },
                @{ label = 'CC'; value = 'None'; empty = $true; emphasis = $false },
                @{ label = 'Subject'; value = $jp; empty = $false; emphasis = $true }
            )
            body_label = 'Message'; body = "line 1`nline 2"
            send_text = 'Send'; open_text = 'Open in Outlook to check'; cancel_text = "Don't send"
            countdown_format = 'Closes automatically in {0} seconds'
        }
        $r = Ask 'confirm' @{ view = $view; timeout = 2 }
        Check 'window ok' $r.ok $true
        Check 'an unanswered window means timeout, never send' $r.result.decision 'timeout'

        # What each button answers, and that nothing answers "send" by default.
        $jsonView = ConvertFrom-Json -InputObject (ConvertTo-Json -InputObject $view -Depth 6)
        foreach ($case in @(@('Send', 'send'), @('Open', 'open'), @('Cancel', 'cancel'))) {
            $w = New-ConfirmWindow $jsonView 30
            $button = $w.FindName($case[0])
            $button.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
            Check ("button " + $case[0] + " answers " + $case[1]) $script:dialogDecision $case[1]
        }
        $w = New-ConfirmWindow $jsonView 30
        Check 'a freshly built window answers cancel until a button is pressed' $script:dialogDecision 'cancel'
        Check 'the send button is not the default button' $w.FindName('Send').IsDefault $false
        Check 'the cancel button is the Esc button' $w.FindName('Cancel').IsCancel $true
        Check 'button order: cancel, open on the left; send on the right' (
            [string][System.Windows.Controls.DockPanel]::GetDock($w.FindName('Cancel')) + ',' +
            [string][System.Windows.Controls.DockPanel]::GetDock($w.FindName('Open')) + ',' +
            [string][System.Windows.Controls.DockPanel]::GetDock($w.FindName('Send'))) 'Left,Left,Right'
        $w.Close()
    } catch {
        Write-Host ("FAIL window test: " + $_.Exception.Message)
        $script:failures++
    }

    # Render the real window to pictures, for people to look at.
    $dist = Join-Path (Split-Path -Parent $PSScriptRoot) 'dist'
    foreach ($lang in @('ja', 'en')) {
        $samplePath = Join-Path $dist "sample-view-$lang.json"
        if (-not (Test-Path $samplePath)) { continue }
        try {
            $sample = ConvertFrom-Json -InputObject ([System.IO.File]::ReadAllText($samplePath, [System.Text.Encoding]::UTF8))
            $w = New-ConfirmWindow $sample 174
            $w.Show()
            $w.Dispatcher.Invoke([Action] { }, [System.Windows.Threading.DispatcherPriority]::ContextIdle)
            $root = $w.Content
            $scale = 2
            $bitmap = New-Object System.Windows.Media.Imaging.RenderTargetBitmap(
                [int]($root.ActualWidth * $scale), [int]($root.ActualHeight * $scale), (96 * $scale), (96 * $scale),
                [System.Windows.Media.PixelFormats]::Pbgra32)
            $bitmap.Render($root)
            $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
            $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
            $png = Join-Path $dist "confirm-window-$lang.png"
            $stream = [System.IO.File]::Create($png)
            try { $encoder.Save($stream) } finally { $stream.Close() }
            $w.Close()
            Write-Host ("picture written: {0} ({1}x{2})" -f $png, $bitmap.PixelWidth, $bitmap.PixelHeight)
        } catch {
            Write-Host ("skip picture ($lang): " + $_.Exception.Message)
        }
    }
} else {
    Write-Host 'skip window test: no desktop session'
}

if ($script:failures -gt 0) { Write-Host "$($script:failures) check(s) failed"; exit 1 }
Write-Host 'All helper checks passed'
