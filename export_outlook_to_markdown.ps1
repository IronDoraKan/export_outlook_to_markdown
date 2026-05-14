param(
    [string]$MailboxName,
    [string]$FolderPath = "Inbox",
    [string]$OutputDir = ".\outlook_markdown",
    [int]$MaxMessages = 50,
    [switch]$SingleFileOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-SafeFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $sanitized = $Name

    foreach ($char in $invalidChars) {
        $escaped = [Regex]::Escape([string]$char)
        $sanitized = $sanitized -replace $escaped, "_"
    }

    $sanitized = $sanitized.Trim()
    if ([string]::IsNullOrWhiteSpace($sanitized)) {
        return "untitled"
    }

    return $sanitized
}

function Convert-ToMarkdownText {
    param(
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return ""
    }

    $normalized = $Text -replace "`r`n", "`n"
    $normalized = $normalized -replace "`r", "`n"

    return $normalized.Trim()
}

function Get-OutlookFolderByPath {
    param(
        [Parameter(Mandatory = $true)]
        $Namespace,
        [string]$MailboxName,
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )

    if ([string]::IsNullOrWhiteSpace($MailboxName)) {
        $currentFolder = $Namespace.GetDefaultFolder(6)
        $segments = @(($FolderPath -split "[\\/]") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

        if ($segments.Count -gt 0 -and $segments[0].ToLowerInvariant() -eq "inbox") {
            if ($segments.Count -eq 1) {
                $segments = @()
            }
            else {
                $segments = $segments[1..($segments.Count - 1)]
            }
        }
    }
    else {
        $currentFolder = $null
        foreach ($folder in $Namespace.Folders) {
            if ($folder.Name -eq $MailboxName) {
                $currentFolder = $folder
                break
            }
        }

        if ($null -eq $currentFolder) {
            throw "Mailbox '$MailboxName' was not found."
        }

        $segments = @(($FolderPath -split "[\\/]") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    foreach ($segment in $segments) {
        $nextFolder = $null
        foreach ($subFolder in $currentFolder.Folders) {
            if ($subFolder.Name -eq $segment) {
                $nextFolder = $subFolder
                break
            }
        }

        if ($null -eq $nextFolder) {
            throw "Folder path '$FolderPath' was not found."
        }

        $currentFolder = $nextFolder
    }

    return $currentFolder
}

function New-MailMarkdown {
    param(
        [Parameter(Mandatory = $true)]
        $MailItem
    )

    $subject = if ([string]::IsNullOrWhiteSpace($MailItem.Subject)) { "(No subject)" } else { $MailItem.Subject.Trim() }
    $sender = if ([string]::IsNullOrWhiteSpace($MailItem.SenderName)) { "(Unknown sender)" } else { $MailItem.SenderName.Trim() }
    $senderEmail = ""

    try {
        if ($null -ne $MailItem.SenderEmailAddress) {
            $senderEmail = $MailItem.SenderEmailAddress.Trim()
        }
    }
    catch {
        $senderEmail = ""
    }

    $receivedTime = ""
    try {
        if ($null -ne $MailItem.ReceivedTime) {
            $receivedTime = ([datetime]$MailItem.ReceivedTime).ToString("yyyy-MM-dd HH:mm:ss")
        }
    }
    catch {
        $receivedTime = ""
    }

    $toLine = ""
    try {
        if ($null -ne $MailItem.To) {
            $toLine = $MailItem.To.Trim()
        }
    }
    catch {
        $toLine = ""
    }

    $ccLine = ""
    try {
        if ($null -ne $MailItem.CC) {
            $ccLine = $MailItem.CC.Trim()
        }
    }
    catch {
        $ccLine = ""
    }

    $body = Convert-ToMarkdownText -Text $MailItem.Body

    $lines = @()
    $lines += "# $subject"
    $lines += ""
    $lines += "- Subject: $subject"
    $lines += "- From: $sender" + $(if ($senderEmail) { " <$senderEmail>" } else { "" })
    $lines += "- To: $toLine"
    $lines += "- Cc: $ccLine"
    $lines += "- Received: $receivedTime"
    $lines += ""
    $lines += "## Body"
    $lines += ""

    if ([string]::IsNullOrWhiteSpace($body)) {
        $lines += "_No body content_"
    }
    else {
        $lines += $body
    }

    return ($lines -join "`r`n")
}

Write-Host "正在連線到 Outlook..."
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$targetFolder = Get-OutlookFolderByPath -Namespace $namespace -MailboxName $MailboxName -FolderPath $FolderPath

if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$items = @()
foreach ($item in $targetFolder.Items) {
    if ($item.Class -eq 43) {
        $items += $item
    }
}

$items = $items | Sort-Object ReceivedTime -Descending
if ($MaxMessages -gt 0) {
    $items = $items | Select-Object -First $MaxMessages
}

$combinedLines = @()
$combinedLines += "# Outlook Export"
$combinedLines += ""
$combinedLines += "- Folder: $($targetFolder.FolderPath)"
$combinedLines += "- Exported At: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
$combinedLines += "- Message Count: $($items.Count)"
$combinedLines += ""

$index = 1
foreach ($mail in $items) {
    $subject = if ([string]::IsNullOrWhiteSpace($mail.Subject)) { "(No subject)" } else { $mail.Subject.Trim() }
    $receivedTime = ""

    try {
        if ($null -ne $mail.ReceivedTime) {
            $receivedTime = ([datetime]$mail.ReceivedTime).ToString("yyyy-MM-dd_HHmmss")
        }
    }
    catch {
        $receivedTime = "unknown_time"
    }

    $fileName = "{0:D3}_{1}_{2}.md" -f $index, $receivedTime, (Get-SafeFileName -Name $subject)
    $filePath = Join-Path $OutputDir $fileName
    $markdown = New-MailMarkdown -MailItem $mail

    if (-not $SingleFileOnly) {
        [System.IO.File]::WriteAllText($filePath, $markdown, [System.Text.Encoding]::UTF8)
    }

    $combinedLines += "---"
    $combinedLines += ""
    $combinedLines += $markdown
    $combinedLines += ""

    $index += 1
}

$combinedFile = Join-Path $OutputDir "emails.md"
[System.IO.File]::WriteAllText($combinedFile, ($combinedLines -join "`r`n"), [System.Text.Encoding]::UTF8)

Write-Host "完成。"
Write-Host "總共匯出 $($items.Count) 封郵件。"
Write-Host "合併檔案: $combinedFile"

if (-not $SingleFileOnly) {
    Write-Host "單封郵件檔案資料夾: $(Resolve-Path -LiteralPath $OutputDir)"
}
