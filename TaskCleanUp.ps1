param(
    [Parameter (Mandatory=$true, Position=1)][string]$client_id 
)

Add-Type -AssemblyName System.Web

Function Authorize
{
    Function AuthorizeWithCode
    {
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

        Write-Host "Get tokens using code flow...`n"

        $redirectUrl = "https://login.microsoftonline.com/common/oauth2/nativeclient"
        $loginUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=code&prompt=login&redirect_uri=" + 
                [System.Web.HttpUtility]::UrlEncode($redirectUrl) + 
                "&client_id=$client_id" +
                "&scope=offline_access%20" + [System.Web.HttpUtility]::UrlEncode("https://outlook.office.com/tasks.readwrite")

        $queryOutput = Show-OAuthWindow -Url $loginUrl

        $authorizationPostRequest = 
            "grant_type=authorization_code" + "&" +
            "redirect_uri=" + [System.Web.HttpUtility]::UrlEncode($redirectUrl) + "&" +
            "client_id=$client_id" + "&" +
            "code=" + $queryOutput["code"]

        $authorization = Invoke-RestMethod  -Method Post `
                            -ContentType application/x-www-form-urlencoded `
                            -Uri https://login.microsoftonline.com/common/oauth2/v2.0/token `
                            -Body $authorizationPostRequest
        $authorization | ConvertTo-Json | Out-File -FilePath "tokens"    

        $authorization.access_token                
    }

    Function AuthorizeWithRefreshToken
    {
        if([System.IO.File]::Exists("tokens"))
        {
            Write-Host "Refresh tokens...`n"

            $authorization = Get-Content 'tokens' | Out-String | ConvertFrom-Json

            $authorizationPostRequest = 
                "grant_type=refresh_token" + "&" +
                "scope=offline_access%20" + [System.Web.HttpUtility]::UrlEncode("https://outlook.office.com/tasks.readwrite") + "&" +
                "client_id=$client_id" + "&" +
                "refresh_token=" + $authorization.refresh_token

            $authorization = Invoke-RestMethod  -Method Post `
                            -ContentType application/x-www-form-urlencoded `
                            -Uri https://login.microsoftonline.com/common/oauth2/v2.0/token `
                            -Body $authorizationPostRequest
                
            $authorization | ConvertTo-Json | Out-File -FilePath "tokens"    

            $authorization.access_token
        }
    }

    $result = AuthorizeWithRefreshToken 

    if(!$result) 
    {
        $result = AuthorizeWithCode
    }

    $result
}

$access_token = Authorize

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add('Accept','Application/Json')
$headers.Add("Authorization", "Bearer " + $access_token)

Write-Host "Get tasks...`n"
$Tasks = (ConvertFrom-Json ((Invoke-WebRequest -Uri "https://outlook.office.com/api/v2.0/me/tasks?`$top=1000" -Headers $headers -UseBasicParsing).ToString())).value
$CompletedTasks = ( $Tasks | Where-Object {$_.Status -eq "Completed"} | Where-Object { (NEW-TIMESPAN -Start $_.LastModifiedDateTime -End (Get-Date)).TotalDays -ge 1} )

Write-Host "Dropping completed tasks`n"
$CompletedTasks |% { 
    Write-Host "."
    Invoke-WebRequest -Method Delete -Uri  "https://outlook.office365.com/api/v2.0/me/tasks('$($_.Id)')" -Headers $headers -UseBasicParsing
}
