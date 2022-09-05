function GetMSALToken {
    $authResult = Get-MsalToken -ClientId 'd1ddf0e4-d672-4dae-b554-9d5bdfd93547' -Scope @("https://graph.microsoft.com/Calendars.Read","https://graph.microsoft.com/Calendars.ReadWrite") -RedirectUri 'urn:ietf:wg:oauth:2.0:oob' -ForceRefresh
    $authHeader = @{
        'Content-Type'  = 'application/json'
        'Authorization' = "Bearer " + $authResult.AccessToken
        'ExpiresOn'     = $authResult.ExpiresOn
    }
    return $authHeader
}

function GetGraphToken {
    $global:authToken = GetMSALToken
    return $global:authToken
}

$null = GetGraphToken

$date1 = Read-Host "Enter the desired Date in locale format"

$startdate = Get-date $date1 -Format "yyyy-MM-dd"
$filter = '$select=id&$top=1000000&$filter=start/DateTime'

$response = (Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/me/calendar/events?$filter lt '$($startdate)'" -Method Get -Headers $authToken).Content


$json = $response | ConvertFrom-Json

$events = $json.Value.id

foreach ($e in $events) {
    $bod = @{}
    $bod.sensitivity = "private"
    
    $body = $bod | ConvertTo-Json
    $null = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/me/calendar/events/$e" -Method Patch -Body $body -Headers $authToken
}
