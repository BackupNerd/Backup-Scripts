# ----- About: ----
    # N-able Cove Data Protection Authenticate
    # Revision v24.09.26
    # Author: Eric Harless, Head Backup Nerd - N-able
    # Twitter @Backup_Nerd  Email:eric.harless@n-able.com
    # Script repository @ https://github.com/backupnerd
    # Schedule a meeting @ https://calendly.com/backup_nerd/
# -----------------------------------------------------------#>  ## About

<# ----- Legal: ----
    # Sample scripts are not supported under any N-able support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # N-able expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose.
    # In no event shall N-able or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
# -----------------------------------------------------------#>  ## Legal

<# ----- Compatibility: ----
    # For use with N-able | Cove Data Protection
    # Sample scripts may contain non-public API calls which are subject to change without notification
# -----------------------------------------------------------#>  ## Compatibility

<# ----- Behavior: ----
    # Authenticate to https://backup.management console
    #
    # https://documentation.n-able.com/covedataprotection/USERGUIDE/documentation/Content/service-management/json-api/home.htm
# -----------------------------------------------------------#>  ## Behavior

Function CDP-Authenticate {
                                        <#> Comment out params section and add as inputs for use with N-able Automation Manager 
    param (
        [string]$Username,              ## Cove login email address or API username used to access https://backup.management
        [string]$Password,              ## Cove API password used to access https://backup.management
        )                               #># Comment out params section and add as inputs for use with N-able Automation Manager 
        
    if ($Username -and -not $Password) {
        $Credential = Get-Credential -Message "Enter your API user credentials for Cove Data Protection" -username $username 
    }
    elseif (-not $Username -or -not $Password) {
        $Credential = Get-Credential -Message "Enter your API user credentials for Cove Data Protection"
    }
    
    if ($Credential) {
        $Username = $Credential.UserName
        $SecurePassword = $Credential.Password
        $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword))
    }

    if ($Username -and $Password) {
        $url = "https://api.backup.management/jsonapi"
        $data = @{}
        $data.jsonrpc = '2.0'
        $data.id = '2'
        $data.method = 'Login'
        $data.params = @{}
        $data.params.username = $username
        $data.params.password = $password

        $webrequest = Invoke-WebRequest -Method POST `
            -ContentType 'application/json' `
            -Body (ConvertTo-Json $data) `
            -Uri $url `
            -SessionVariable Script:websession `
            -UseBasicParsing
            $Script:cookies = $websession.Cookies.GetCookies($url)
            $Script:websession = $websession
            $Script:Authenticate = $webrequest | convertfrom-json

        #Debug Write-Output "$($Script:cookies[0].name) = $($cookies[0].value)"

        if ($authenticate.visa) { 
            $Script:visa = $authenticate.visa
            Write-Output "  Authentication Success: User $($Authenticate.result.result.emailaddress)"
        }else{
            Write-Output $Script:strLineSeparator 
            Write-Output "  Authentication Failed: Please confirm your Backup.Management API User Credentials"
            Write-Output "  Please Note: Multiple failed authentication attempts could temporarily lockout your user account"
            Write-Output $Script:strLineSeparator 
            Write-Warning $authenticate.error.message
        }
    }

} ## Use Backup.Management credentials to Authenticate

CDP-Authenticate