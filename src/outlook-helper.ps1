# Outlook helper for outlook-classic-mcp. Runs under Windows PowerShell 5.1.
#
# Reads one JSON request per line from stdin and writes one JSON answer per line to stdout.
# It only does what it is told: the rules (what may be sent, where files may be saved,
# when to ask the human) live in the Node side. Nothing here sends mail on its own.
#
# Written for PowerShell 5.1: no ternary, no ??, no ?. operators.
param([switch]$NoLoop)

$ErrorActionPreference = 'Stop'

$OL_MAIL = 43
$OL_FORMAT_HTML = 2
$OL_MAIL_ITEM_TYPE = 0
$MAX_SCAN = 5000
$MAX_FOLDERS = 500
$MAX_FOLDER_DEPTH = 6
$SCAN_BUDGET_SECONDS = 20
$PR_SMTP_ADDRESS = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001F'
$WELL_KNOWN = @{ inbox = 6; sent = 5; drafts = 16; deleted = 3; outbox = 4; junk = 23 }

$script:app = $null

# ---- errors -------------------------------------------------------------------

function Fail([string]$kind, [string]$message) {
    $e = New-Object System.Exception($message)
    $e.Data['kind'] = $kind
    throw $e
}

function Find-ComException($exception) {
    $e = $exception
    while ($null -ne $e) {
        if ($e -is [System.Runtime.InteropServices.COMException]) { return $e }
        $e = $e.InnerException
    }
    return $null
}

# Turns a failure into a message a non-technical person can act on.
function Convert-Failure($exception) {
    $e = $exception
    while ($null -ne $e) {
        if ($e.Data -and $e.Data.Contains('kind')) {
            return @{ kind = [string]$e.Data['kind']; message = [string]$e.Message }
        }
        $e = $e.InnerException
    }
    $com = Find-ComException $exception
    if ($null -ne $com) {
        $script:app = $null   # reconnect on the next call
        $code = '{0:X8}' -f $com.ErrorCode
        if ($code -eq '800401F3' -or $code -eq '80040154') {
            return @{ kind = 'unavailable'; message = 'Classic Outlook was not found on this PC. The new Outlook app is not supported; classic Outlook (OUTLOOK.EXE) is required.' }
        }
        if ($code -eq '80080005') {
            return @{ kind = 'unavailable'; message = "Outlook did not respond. Start Outlook normally (not 'Run as administrator'), make sure Claude is not running as administrator either, and try again." }
        }
        if ($code -eq '800706BA' -or $code -eq '800706BE') {
            return @{ kind = 'unavailable'; message = 'The connection to Outlook was lost. Make sure Outlook is open and try again.' }
        }
        return @{ kind = 'error'; message = ('Outlook reported a problem: ' + [string]$com.Message) }
    }
    return @{ kind = 'error'; message = ('Outlook reported a problem: ' + [string]$exception.Message) }
}

# ---- connection and lookup ------------------------------------------------------

function Get-App {
    if ($null -eq $script:app) {
        $script:app = New-Object -ComObject Outlook.Application
    }
    return $script:app
}

function Get-Ns {
    $app = Get-App
    return $app.GetNamespace('MAPI')
}

function Get-StoreKey([string]$storeId) {
    $sha = [System.Security.Cryptography.SHA1]::Create()
    try {
        $hash = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($storeId))
    } finally {
        $sha.Dispose()
    }
    $hex = ''
    foreach ($b in $hash) { $hex += $b.ToString('x2') }
    return $hex.Substring(0, 10)
}

# An email is identified by its EntryID together with its StoreID (carried as a short key).
function Get-ItemId($item) {
    return (Get-StoreKey ([string]$item.Parent.StoreID)) + '.' + [string]$item.EntryID
}

function Find-Item([string]$id) {
    $dot = $id.IndexOf('.')
    if ($dot -le 0 -or $dot -eq ($id.Length - 1)) {
        Fail 'invalid' 'That email reference is not valid. List the emails again.'
    }
    $key = $id.Substring(0, $dot)
    $entry = $id.Substring($dot + 1)
    $ns = Get-Ns
    $storeId = $null
    foreach ($store in $ns.Stores) {
        if ((Get-StoreKey ([string]$store.StoreID)) -eq $key) { $storeId = [string]$store.StoreID; break }
    }
    if ($null -eq $storeId) {
        Fail 'not_found' "That email's mailbox is no longer available. List the emails again."
    }
    $item = $null
    try {
        $item = $ns.GetItemFromID($entry, $storeId)
    } catch {
        Fail 'not_found' 'That email was not found. It may have been moved or deleted; list the emails again.'
    }
    if ($null -eq $item) {
        Fail 'not_found' 'That email was not found. It may have been moved or deleted; list the emails again.'
    }
    return $item
}

function Find-Draft([string]$id) {
    $item = Find-Item $id
    if ([int]$item.Class -ne $OL_MAIL -or [bool]$item.Sent) {
        Fail 'not_draft' 'That item is not an unsent draft, so it cannot be changed or sent.'
    }
    return $item
}

function Find-Account($ns, [string]$name) {
    $wanted = $name.Trim()
    foreach ($account in $ns.Accounts) {
        if (([string]$account.DisplayName -ieq $wanted) -or ([string]$account.SmtpAddress -ieq $wanted)) {
            return $account
        }
    }
    return $null
}

function Find-Store($ns, $account) {
    if ([string]::IsNullOrWhiteSpace([string]$account)) { return $ns.DefaultStore }
    $found = Find-Account $ns ([string]$account)
    if ($null -ne $found) { return $found.DeliveryStore }
    foreach ($store in $ns.Stores) {
        if ([string]$store.DisplayName -ieq ([string]$account).Trim()) { return $store }
    }
    Fail 'not_found' "No account or data file named '$account' was found."
}

function Resolve-MailFolder($ns, $folder, $account) {
    $name = ([string]$folder).Trim()
    if ($name.Length -eq 0) { $name = 'inbox' }
    if ($WELL_KNOWN.ContainsKey($name.ToLowerInvariant())) {
        $store = Find-Store $ns $account
        try {
            return $store.GetDefaultFolder($WELL_KNOWN[$name.ToLowerInvariant()])
        } catch {
            Fail 'not_found' "This account has no '$name' folder."
        }
    }
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($p in $name.Split('\')) { if ($p.Length -gt 0) { $parts.Add($p) } }
    if ($parts.Count -eq 0) { Fail 'invalid' 'The folder name is empty.' }
    $current = $null
    foreach ($store in $ns.Stores) {
        if ([string]$store.DisplayName -ieq $parts[0]) {
            $current = $store.GetRootFolder()
            $parts.RemoveAt(0)
            break
        }
    }
    if ($null -eq $current) {
        $store = Find-Store $ns $account
        $current = $store.GetRootFolder()
    }
    foreach ($part in $parts) {
        try {
            $current = $current.Folders.Item($part)
        } catch {
            Fail 'not_found' "The folder '$folder' was not found. Use a path from list_folders."
        }
    }
    return $current
}

# ---- converting Outlook items ---------------------------------------------------

function Get-AccountType($account) {
    $type = -1
    try { $type = [int]$account.AccountType } catch { }
    if ($type -eq 0) { return 'exchange' }
    if ($type -eq 1) { return 'imap' }
    if ($type -eq 2) { return 'pop3' }
    return 'other'
}

function Get-SenderInfo($item) {
    $name = [string]$item.SenderName
    $address = [string]$item.SenderEmailAddress
    if ([string]$item.SenderEmailType -eq 'EX') {
        # Exchange gives an internal directory name; ask for the normal email address.
        try {
            $user = $item.Sender.GetExchangeUser()
            if ($null -ne $user) {
                $smtp = [string]$user.PrimarySmtpAddress
                if ($smtp.Length -gt 0) { $address = $smtp }
            }
        } catch { }
    }
    return @{ name = $name; address = $address }
}

function Get-RecipientInfo($item) {
    $to = New-Object System.Collections.Generic.List[object]
    $cc = New-Object System.Collections.Generic.List[object]
    $bcc = New-Object System.Collections.Generic.List[object]
    foreach ($r in $item.Recipients) {
        $address = [string]$r.Address
        if (-not $address.Contains('@')) {
            $smtp = ''
            try { $smtp = [string]$r.PropertyAccessor.GetProperty($PR_SMTP_ADDRESS) } catch { }
            if ($smtp.Length -eq 0) {
                try {
                    $user = $r.AddressEntry.GetExchangeUser()
                    if ($null -ne $user) { $smtp = [string]$user.PrimarySmtpAddress }
                } catch { }
            }
            if ($smtp.Length -gt 0) { $address = $smtp }
        }
        $person = @{ name = [string]$r.Name; address = $address }
        $type = 1
        try { $type = [int]$r.Type } catch { }
        if ($type -eq 2) { $cc.Add($person) }
        elseif ($type -eq 3) { $bcc.Add($person) }
        else { $to.Add($person) }
    }
    return @{ to = $to; cc = $cc; bcc = $bcc }
}

function Format-Person($p) {
    if ($p.name.Length -gt 0 -and $p.address.Length -gt 0 -and $p.name -ne $p.address) {
        return $p.name + ' <' + $p.address + '>'
    }
    if ($p.address.Length -gt 0) { return $p.address }
    return $p.name
}

function Format-People($people) {
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($p in $people) { $result.Add((Format-Person $p)) }
    return @{ list = $result }
}

# Outlook hands back local wall-clock time without a zone. Returns $null for "no date".
function Get-ItemTime($item) {
    foreach ($property in @('ReceivedTime', 'LastModificationTime')) {
        try {
            $d = [datetime]$item.$property
            if ($d.Year -gt 1900 -and $d.Year -lt 4000) {
                $local = [datetime]::SpecifyKind($d, [System.DateTimeKind]::Unspecified)
                return New-Object System.DateTimeOffset($local, [System.TimeZoneInfo]::Local.GetUtcOffset($local))
            }
        } catch { }
    }
    return $null
}

function Format-Iso($when) {
    if ($null -eq $when) { return $null }
    return $when.ToString('yyyy-MM-ddTHH:mm:sszzz', [System.Globalization.CultureInfo]::InvariantCulture)
}

function ConvertTo-Instant($text) {
    if ([string]::IsNullOrWhiteSpace([string]$text)) { return $null }
    return [System.DateTimeOffset]::Parse([string]$text, [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-Summary($item, $when) {
    return @{
        id               = Get-ItemId $item
        subject          = [string]$item.Subject
        from             = Get-SenderInfo $item
        received         = Format-Iso $when
        unread           = [bool]$item.UnRead
        attachment_count = [int]$item.Attachments.Count
    }
}

function Get-AttachmentInfo($item) {
    $result = New-Object System.Collections.Generic.List[object]
    $attachments = $item.Attachments
    $count = [int]$attachments.Count
    for ($i = 1; $i -le $count; $i++) {
        $a = $attachments.Item($i)
        $name = ''
        try { $name = [string]$a.FileName } catch { }
        if ($name.Length -eq 0) { try { $name = [string]$a.DisplayName } catch { } }
        if ($name.Length -eq 0) { $name = "attachment-$i" }
        $size = 0
        try { $size = [long]$a.Size } catch { }
        $result.Add(@{ index = $i; filename = $name; size = $size })
    }
    return @{ list = $result }
}

# Walks items newest-first and filters dates here, so the result does not depend on the
# PC's regional date format. Stops at the item limit or the time budget.
function Get-Page($items, $since, $until, [int]$limit, [int]$offset) {
    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    $null = $items.Sort('[ReceivedTime]', $true)
    $emails = New-Object System.Collections.Generic.List[object]
    $skipped = 0
    $scanned = 0
    $hasMore = $false
    $timedOut = $false
    $item = $items.GetFirst()
    while ($null -ne $item) {
        $scanned++
        if ($scanned -gt $MAX_SCAN) { $hasMore = $true; break }
        if ($clock.Elapsed.TotalSeconds -gt $SCAN_BUDGET_SECONDS) { $hasMore = $true; $timedOut = $true; break }
        $class = 0
        try { $class = [int]$item.Class } catch { }
        if ($class -eq $OL_MAIL) {
            $when = Get-ItemTime $item
            if ($null -ne $when -and $null -ne $since -and $when -lt $since) { break }   # sorted newest first
            $tooNew = ($null -ne $when -and $null -ne $until -and $when -gt $until)
            if (-not $tooNew) {
                if ($skipped -lt $offset) { $skipped++ }
                elseif ($emails.Count -ge $limit) { $hasMore = $true; break }
                else { $emails.Add((Get-Summary $item $when)) }
            }
        }
        $item = $items.GetNext()
    }
    return @{ emails = $emails; has_more = $hasMore; timed_out = $timedOut }
}

# ---- reading --------------------------------------------------------------------

function Invoke-ListAccounts($p) {
    $ns = Get-Ns
    $accounts = New-Object System.Collections.Generic.List[object]
    $first = $true
    foreach ($a in $ns.Accounts) {
        $accounts.Add(@{
                name       = [string]$a.DisplayName
                address    = [string]$a.SmtpAddress
                type       = Get-AccountType $a
                is_default = $first
            })
        $first = $false
    }
    $files = New-Object System.Collections.Generic.List[object]
    foreach ($store in $ns.Stores) { $files.Add([string]$store.DisplayName) }
    return @{ accounts = $accounts; data_files = $files }
}

function Add-Folders($parent, [int]$depth, $result) {
    if ($depth -gt $MAX_FOLDER_DEPTH) { return }
    foreach ($sub in $parent.Folders) {
        if ($result.Count -ge $MAX_FOLDERS) { return }
        try {
            if ([int]$sub.DefaultItemType -eq $OL_MAIL_ITEM_TYPE) {
                $result.Add(@{
                        path   = ([string]$sub.FolderPath).TrimStart('\')
                        name   = [string]$sub.Name
                        total  = [int]$sub.Items.Count
                        unread = [int]$sub.UnReadItemCount
                    })
            }
            Add-Folders $sub ($depth + 1) $result
        } catch {
            # A folder that cannot be read (offline, no permission) is skipped, not fatal.
        }
    }
}

function Invoke-ListFolders($p) {
    $ns = Get-Ns
    $result = New-Object System.Collections.Generic.List[object]
    if (-not [string]::IsNullOrWhiteSpace([string]$p.account)) {
        $store = Find-Store $ns $p.account
        Add-Folders ($store.GetRootFolder()) 1 $result
    } else {
        foreach ($store in $ns.Stores) { Add-Folders ($store.GetRootFolder()) 1 $result }
    }
    return @{ folders = $result }
}

function Invoke-ListEmails($p) {
    $ns = Get-Ns
    $mailFolder = Resolve-MailFolder $ns $p.folder $p.account
    $items = $mailFolder.Items
    if ([bool]$p.unread_only) { $items = $items.Restrict('[UnRead] = True') }
    return Get-Page $items (ConvertTo-Instant $p.since) (ConvertTo-Instant $p.until) ([int]$p.limit) ([int]$p.offset)
}

function Invoke-SearchEmails($p) {
    $ns = Get-Ns
    $mailFolder = Resolve-MailFolder $ns $p.folder $p.account
    $items = $mailFolder.Items
    if (-not [string]::IsNullOrWhiteSpace([string]$p.filter)) { $items = $items.Restrict([string]$p.filter) }
    return Get-Page $items (ConvertTo-Instant $p.since) (ConvertTo-Instant $p.until) ([int]$p.limit) 0
}

function Invoke-ReadEmail($p) {
    $item = Find-Item ([string]$p.id)
    if ([int]$item.Class -ne $OL_MAIL) {
        Fail 'not_found' 'That item is not an email (for example, a meeting request).'
    }
    $people = Get-RecipientInfo $item
    $result = Get-Summary $item (Get-ItemTime $item)
    $result['to'] = (Format-People $people.to).list
    $result['cc'] = (Format-People $people.cc).list
    $result['bcc'] = (Format-People $people.bcc).list
    $result['is_draft'] = (-not [bool]$item.Sent)
    $result['body'] = [string]$item.Body
    $result['attachments'] = (Get-AttachmentInfo $item).list
    return $result
}

function Invoke-ListAttachments($p) {
    $item = Find-Item ([string]$p.id)
    return @{ attachments = (Get-AttachmentInfo $item).list }
}

function Invoke-SaveAttachment($p) {
    $item = Find-Item ([string]$p.id)
    $attachments = $item.Attachments
    $index = [int]$p.index
    if ($index -lt 1 -or $index -gt [int]$attachments.Count) {
        Fail 'not_found' 'That attachment number does not exist on this email.'
    }
    $attachment = $attachments.Item($index)
    $null = $attachment.SaveAsFile([string]$p.dest)
    return @{ saved = $true }
}

# ---- drafting -------------------------------------------------------------------

function Join-Addresses($values) {
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($v in @($values)) { if ($null -ne $v -and ([string]$v).Length -gt 0) { $list.Add([string]$v) } }
    return [string]::Join('; ', $list.ToArray())
}

function Add-Files($mail, $paths) {
    foreach ($path in @($paths)) {
        if ($null -ne $path -and ([string]$path).Length -gt 0) { $null = $mail.Attachments.Add([string]$path) }
    }
}

# Puts the new text above the quoted original without destroying its formatting.
function Add-TextOnTop($mail, [string]$text, [string]$htmlBlock) {
    if ($text.Length -eq 0) { return }
    if ([int]$mail.BodyFormat -eq $OL_FORMAT_HTML) {
        $html = [string]$mail.HTMLBody
        $match = [regex]::Match($html, '<body[^>]*>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($match.Success) {
            $cut = $match.Index + $match.Length
            $mail.HTMLBody = $html.Substring(0, $cut) + $htmlBlock + $html.Substring($cut)
        } else {
            $mail.HTMLBody = $htmlBlock + $html
        }
    } else {
        $mail.Body = $text + "`r`n`r`n" + [string]$mail.Body
    }
}

function Invoke-CreateDraft($p) {
    $app = Get-App
    $ns = Get-Ns
    $mail = $app.CreateItem(0)
    if (-not [string]::IsNullOrWhiteSpace([string]$p.account)) {
        $account = Find-Account $ns ([string]$p.account)
        if ($null -eq $account) { Fail 'not_found' "No account named '$($p.account)' was found." }
        # A plain assignment does not stick for this property from PowerShell.
        $null = $mail.GetType().InvokeMember('SendUsingAccount', [System.Reflection.BindingFlags]::SetProperty, $null, $mail, @($account))
    }
    $mail.To = Join-Addresses $p.to
    $mail.CC = Join-Addresses $p.cc
    $mail.BCC = Join-Addresses $p.bcc
    $mail.Subject = [string]$p.subject
    $mail.Body = [string]$p.body
    $null = $mail.Recipients.ResolveAll()
    Add-Files $mail $p.attachments
    $null = $mail.Save()
    return @{ id = Get-ItemId $mail }
}

function Invoke-CreateReplyDraft($p) {
    $original = Find-Item ([string]$p.id)
    if ([bool]$p.reply_all) { $reply = $original.ReplyAll() } else { $reply = $original.Reply() }
    Add-TextOnTop $reply ([string]$p.text) ([string]$p.html_block)
    $null = $reply.Save()
    return @{ id = Get-ItemId $reply }
}

function Invoke-CreateForwardDraft($p) {
    $original = Find-Item ([string]$p.id)
    $forward = $original.Forward()
    $forward.To = Join-Addresses $p.to
    $null = $forward.Recipients.ResolveAll()
    Add-TextOnTop $forward ([string]$p.text) ([string]$p.html_block)
    $null = $forward.Save()
    return @{ id = Get-ItemId $forward }
}

function Invoke-UpdateDraft($p) {
    $mail = Find-Draft ([string]$p.id)
    $recipientsChanged = $false
    if ([bool]$p.set_to) { $mail.To = Join-Addresses $p.to; $recipientsChanged = $true }
    if ([bool]$p.set_cc) { $mail.CC = Join-Addresses $p.cc; $recipientsChanged = $true }
    if ([bool]$p.set_bcc) { $mail.BCC = Join-Addresses $p.bcc; $recipientsChanged = $true }
    if ([bool]$p.set_subject) { $mail.Subject = [string]$p.subject }
    if ([bool]$p.set_body) { $mail.Body = [string]$p.body }
    if ($recipientsChanged) { $null = $mail.Recipients.ResolveAll() }
    Add-Files $mail $p.add_attachments
    $null = $mail.Save()
    return @{ updated = $true }
}

function Invoke-GetDraft($p) {
    $mail = Find-Draft ([string]$p.id)
    $account = $null
    try { $account = $mail.SendUsingAccount } catch { }
    if ($null -eq $account) {
        $accounts = (Get-Ns).Accounts
        if ([int]$accounts.Count -gt 0) { $account = $accounts.Item(1) }
    }
    $senderName = ''
    $senderAddress = ''
    $accountType = 'other'
    if ($null -ne $account) {
        $senderName = [string]$account.DisplayName
        $senderAddress = [string]$account.SmtpAddress
        $accountType = Get-AccountType $account
    }
    $people = Get-RecipientInfo $mail
    return @{
        sender_name    = $senderName
        sender_address = $senderAddress
        account_type   = $accountType
        to             = $people.to
        cc             = $people.cc
        bcc            = $people.bcc
        subject        = [string]$mail.Subject
        body           = [string]$mail.Body
        attachments    = (Get-AttachmentInfo $mail).list
    }
}

# Called only after the Node side has received "send" from the confirmation window.
function Invoke-Send($p) {
    $mail = Find-Draft ([string]$p.id)
    $null = $mail.Send()
    return @{ sent = $true }
}

function Invoke-Display($p) {
    $mail = Find-Draft ([string]$p.id)
    $null = $mail.Display($false)
    return @{ displayed = $true }
}

# ---- the confirmation window ----------------------------------------------------
# Sending is never the default: closing the window, pressing Enter or Esc, or waiting
# all mean "do not send".

function Add-RichText($box, [string]$text, $font, [int]$indent) {
    $box.SelectionStart = $box.TextLength
    $box.SelectionLength = 0
    $box.SelectionFont = $font
    $box.SelectionIndent = $indent
    $box.AppendText($text)
}

function New-DialogButton([string]$text, [string]$decision) {
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $text
    $button.Tag = $decision
    $button.AutoSize = $true
    $button.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $button.Padding = New-Object System.Windows.Forms.Padding(14, 6, 14, 6)
    $button.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $button.Add_Click({
            $script:dialogDecision = [string]$this.Tag
            $this.FindForm().Close()
        })
    return $button
}

function Invoke-Confirm($p) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    # This helper is started hidden (so no console window appears), and Windows applies that
    # "hidden" state to the first window the process shows. Spend it on a throwaway form,
    # otherwise the confirmation window itself would be open but invisible.
    if (-not $script:windowPrimed) {
        $dummy = New-Object System.Windows.Forms.Form
        $dummy.ShowInTaskbar = $false
        $dummy.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        $dummy.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
        $dummy.Location = New-Object System.Drawing.Point(-32000, -32000)
        $dummy.Size = New-Object System.Drawing.Size(1, 1)
        $dummy.Opacity = 0
        $dummy.Show()
        $dummy.Close()
        $dummy.Dispose()
        $script:windowPrimed = $true
    }

    $view = $p.view
    $script:dialogDecision = 'cancel'
    $script:dialogRemaining = [int]$p.timeout
    $script:dialogFormat = [string]$view.countdown_format

    $baseFont = [System.Drawing.SystemFonts]::MessageBoxFont
    $font = New-Object System.Drawing.Font($baseFont.FontFamily, 10)
    $bold = New-Object System.Drawing.Font($font, [System.Drawing.FontStyle]::Bold)
    $headingFont = New-Object System.Drawing.Font($baseFont.FontFamily, 13, [System.Drawing.FontStyle]::Bold)

    $form = New-Object System.Windows.Forms.Form
    $script:dialogForm = $form
    $form.Text = [string]$view.title
    $form.Font = $font
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.ClientSize = New-Object System.Drawing.Size(720, 560)
    $form.MinimumSize = New-Object System.Drawing.Size(560, 420)
    $form.MinimizeBox = $false
    $form.ShowIcon = $false
    $form.ShowInTaskbar = $true
    $form.TopMost = $true
    $form.Padding = New-Object System.Windows.Forms.Padding(14)

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $layout.ColumnCount = 1
    $layout.RowCount = 5
    $null = $layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $null = $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $null = $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $null = $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $null = $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $null = $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))

    $heading = New-Object System.Windows.Forms.Label
    $heading.Text = [string]$view.heading
    $heading.Font = $headingFont
    $heading.AutoSize = $true
    $heading.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
    $layout.Controls.Add($heading, 0, 0)

    $warningLines = New-Object System.Collections.Generic.List[string]
    foreach ($line in @($view.warning_lines)) {
        if ($null -ne $line -and ([string]$line).Length -gt 0) { $warningLines.Add(([string][char]0x26A0) + ' ' + [string]$line) }
    }
    if ($warningLines.Count -gt 0) {
        $warning = New-Object System.Windows.Forms.Label
        $warning.Text = [string]::Join([Environment]::NewLine, $warningLines.ToArray())
        $warning.Font = $bold
        $warning.BackColor = [System.Drawing.Color]::FromArgb(255, 244, 206)
        $warning.ForeColor = [System.Drawing.Color]::FromArgb(138, 31, 17)
        $warning.AutoSize = $true
        $warning.Dock = [System.Windows.Forms.DockStyle]::Fill
        $warning.Padding = New-Object System.Windows.Forms.Padding(10, 8, 10, 8)
        $warning.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
        $layout.Controls.Add($warning, 0, 1)
    }

    $text = New-Object System.Windows.Forms.RichTextBox
    $text.ReadOnly = $true
    $text.Dock = [System.Windows.Forms.DockStyle]::Fill
    $text.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $text.BackColor = [System.Drawing.Color]::White
    $text.DetectUrls = $false   # links in an email must not be clickable from here
    $text.TabStop = $false
    $text.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $text.WordWrap = $true
    foreach ($row in @($view.rows)) {
        Add-RichText $text ([string]$row.label + "`n") $bold 0
        Add-RichText $text ([string]$row.value + "`n") $font 24
    }
    Add-RichText $text ("`n" + [string]$view.body_label + "`n") $bold 0
    Add-RichText $text ([string]$view.body + "`n") $font 0
    $text.Select(0, 0)
    $layout.Controls.Add($text, 0, 2)

    $send = New-DialogButton ([string]$view.send_text) 'send'
    $open = New-DialogButton ([string]$view.open_text) 'open'
    $cancel = New-DialogButton ([string]$view.cancel_text) 'cancel'

    $buttons = New-Object System.Windows.Forms.TableLayoutPanel
    $buttons.Dock = [System.Windows.Forms.DockStyle]::Fill
    $buttons.AutoSize = $true
    $buttons.ColumnCount = 4
    $buttons.RowCount = 1
    $buttons.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 0)
    $null = $buttons.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $null = $buttons.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $null = $buttons.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $null = $buttons.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $buttons.Controls.Add($send, 0, 0)
    $buttons.Controls.Add($open, 1, 0)
    $buttons.Controls.Add($cancel, 3, 0)
    $layout.Controls.Add($buttons, 0, 3)

    $countdown = New-Object System.Windows.Forms.Label
    $script:dialogCountdown = $countdown
    $countdown.AutoSize = $true
    $countdown.ForeColor = [System.Drawing.Color]::DimGray
    $countdown.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    $countdown.Text = [string]::Format($script:dialogFormat, $script:dialogRemaining)
    $layout.Controls.Add($countdown, 0, 4)
    $form.Controls.Add($layout)

    # No AcceptButton: Enter never means "send". Focus starts on "Don't send", so Enter or
    # Space there, Esc, and closing the window all decline.
    $form.CancelButton = $cancel
    $form.ActiveControl = $cancel

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({
            $script:dialogRemaining = $script:dialogRemaining - 1
            if ($script:dialogRemaining -le 0) {
                $script:dialogDecision = 'timeout'
                $script:dialogForm.Close()
            } else {
                $script:dialogCountdown.Text = [string]::Format($script:dialogFormat, $script:dialogRemaining)
            }
        })
    $form.Add_Shown({
            # Second line of defence against the window being open but hidden (see above):
            # changing the window state makes Windows show it for real.
            $this.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
            $this.WindowState = [System.Windows.Forms.FormWindowState]::Normal
            $this.Activate()
        })

    try {
        $timer.Start()
        $null = $form.ShowDialog()
    } finally {
        $timer.Stop()
        $timer.Dispose()
        $form.Dispose()
    }
    return @{ decision = [string]$script:dialogDecision }
}

# ---- request loop ---------------------------------------------------------------

function Invoke-Operation([string]$op, $p) {
    switch ($op) {
        'ping' { $r = @{ pong = $true; version = [string]$PSVersionTable.PSVersion; echo = $p.echo } }
        'list_accounts' { $r = Invoke-ListAccounts $p }
        'list_folders' { $r = Invoke-ListFolders $p }
        'list_emails' { $r = Invoke-ListEmails $p }
        'search_emails' { $r = Invoke-SearchEmails $p }
        'read_email' { $r = Invoke-ReadEmail $p }
        'list_attachments' { $r = Invoke-ListAttachments $p }
        'save_attachment' { $r = Invoke-SaveAttachment $p }
        'create_draft' { $r = Invoke-CreateDraft $p }
        'create_reply_draft' { $r = Invoke-CreateReplyDraft $p }
        'create_forward_draft' { $r = Invoke-CreateForwardDraft $p }
        'update_draft' { $r = Invoke-UpdateDraft $p }
        'get_draft' { $r = Invoke-GetDraft $p }
        'send' { $r = Invoke-Send $p }
        'display' { $r = Invoke-Display $p }
        'confirm' { $r = Invoke-Confirm $p }
        default { Fail 'invalid' "Unknown operation '$op'." }
    }
    # If a call leaked extra output, the real answer is the last thing returned.
    if ($r -is [array]) { $r = $r[$r.Length - 1] }
    return $r
}

function Invoke-Request([string]$line) {
    $id = $null
    try {
        $request = ConvertFrom-Json -InputObject $line
        $id = $request.id
        $result = Invoke-Operation ([string]$request.op) $request.args
        $response = @{ id = $id; ok = $true; result = $result }
    } catch {
        $failure = Convert-Failure $_.Exception
        $response = @{ id = $id; ok = $false; kind = $failure.kind; message = $failure.message }
    }
    return (ConvertTo-Json -InputObject $response -Depth 12 -Compress)
}

if (-not $NoLoop) {
    # Raw UTF-8 streams: the console code page (cp932 on Japanese Windows) is never involved.
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $reader = New-Object System.IO.StreamReader([Console]::OpenStandardInput(), $utf8)
    $writer = New-Object System.IO.StreamWriter([Console]::OpenStandardOutput(), $utf8)
    $writer.AutoFlush = $true
    while ($true) {
        $line = $reader.ReadLine()
        if ($null -eq $line) { break }
        if ($line.Trim().Length -eq 0) { continue }
        $writer.WriteLine((Invoke-Request $line))
    }
}
