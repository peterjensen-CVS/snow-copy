#### Hightlight and copy snow case number
#### Use context menu of windows desktop to open the cases rapidly
####  It is presumed that script will be used with the context menu
#### the Context menu presumes this script file lives in c:\Snow-Copy

#Written by Peter Jensen

$clipboardContent = $null
$content = $null


# Read the clipboard content
$clipboardContent = Get-Clipboard

# ServiceNow instance URL
$serviceNowInstance = "https://aetnaprod1.service-now.com/"


if (-not $clipboardContent) {
    Write-Error "Clipboard is empty."
    exit
}

# Determine the type of record
function DetermineRecordType {
    param (
        [string]$content
    )
    if ($content -match "^INC\d+$") {
        return "incident"
    } elseif ($content -match "^PRB\d+$") {
        return "problem"
    } elseif ($content -match "^CTASK\d+$") {
        return "ctask"
    } elseif ($content -match "^CHG\d+$") {
        return "change"
    } elseif ($content -match "^RITM\d+$") {
        return "ritm"       
    } elseif ($content -match "^TASK\d+$") {
        return "task"       
    } else {
        return $null
    }
}

$recordType = DetermineRecordType -content $clipboardContent

if (-not $recordType) {
    Write-Error "Could not determine the record type from the clipboard content."
    exit
}

# Construct the URL based on the record type
switch ($recordType) {
    "incident" { $url = "$serviceNowInstance/nav_to.do?uri=incident.do?sys_id=$clipboardContent" }
    "problem" { $url = "$serviceNowInstance/nav_to.do?uri=problem.do?sys_id=$clipboardContent" }
    "ctask" { $url = "$serviceNowInstance/nav_to.do?uri=sc_task.do?sys_id=$clipboardContent" }
    "change" { $url = "$serviceNowInstance/nav_to.do?uri=change_request.do?sys_id=$clipboardContent" }
    "ritm" { $url = "$serviceNowInstance/nav_to.do?uri=sc_req_item.do?sys_id=$clipboardContent" }
    "task" { $url = "$serviceNowInstance/nav_to.do?uri=sc_task.do?sys_id=$clipboardContent" }
    default { Write-Error "Unknown record type."; exit }
}

# Open the URL in Edge browser
Start-Process "msedge" -ArgumentList $url


