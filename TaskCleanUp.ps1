param(
    [Parameter (Mandatory=$true, Position=1)]
    [string]$client_id 
)

Add-Type -AssemblyName System.Web

Function Show-OAuthWindow
{
    param(
        [System.Uri]$Url
    )


    Add-Type -AssemblyName System.Windows.Forms
 
    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
    $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url ) }
    $DocComp  = {
        $Global:uri = $web.Url.AbsoluteUri
        if ($Global:Uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
    }
    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($DocComp)
    $form.Controls.Add($web)
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() | Out-Null

    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
    $output = @{}
    foreach($key in $queryOutput.Keys){
        $output["$key"] = $queryOutput[$key]
    }
    
    $output
}



$redirectUrl = "https://login.microsoftonline.com/common/oauth2/nativeclient"
$loginUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=code&redirect_uri=" + 
            [System.Web.HttpUtility]::UrlEncode($redirectUrl) + 
            "&client_id=$client_id" +
            "&scope=" + [System.Web.HttpUtility]::UrlEncode("https://outlook.office.com/tasks.readwrite")

$queryOutput = Show-OAuthWindow -Url $loginUrl

$AuthorizationPostRequest = 
    "grant_type=authorization_code" + "&" +
    "redirect_uri=" + [System.Web.HttpUtility]::UrlEncode($redirectUrl) + "&" +
    "client_id=$client_id" + "&" +
    "code=" + $queryOutput["code"]

$Authorization = Invoke-RestMethod   -Method Post `
                        -ContentType application/x-www-form-urlencoded `
                        -Uri https://login.microsoftonline.com/common/oauth2/v2.0/token `
                        -Body $AuthorizationPostRequest


$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add('Accept','Application/Json')
$headers.Add("Authorization", "Bearer " + $Authorization.access_token)

$Tasks = (ConvertFrom-Json ((Invoke-WebRequest -Uri "https://outlook.office.com/api/v2.0/me/tasks?`$top=1000" -Headers $headers).ToString())).value
$CompletedTasks = ( $Tasks | Where-Object {$_.Status -eq "Completed"} | Where-Object { (NEW-TIMESPAN -Start $_.LastModifiedDateTime -End (Get-Date)).TotalDays -ge 1} )

$CompletedTasks |% { Invoke-WebRequest -Method Delete -Uri  "https://outlook.office365.com/api/v2.0/me/tasks('$($_.Id)')" -Headers $headers }
