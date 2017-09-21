    <#
    .SYNOPSIS
        Get an Intune application deployment overview through the Microsoft Graph API
        NOTE: This function requires that AzureAD (Preview) module is installed. Use 'Install-Module -Name AzureAD' or 'Install-Module -Name AzureADPreview' to install it.

    .PARAMETER TenantName
        A tenant name should be provided in the following format: tenantname.onmicrosoft.com.

    .PARAMETER ClientID
        Application ID for an Azure AD application.

    .PARAMETER ExportPath
        An path should be provided to export the files to
        
    .PARAMETER Runtype
        Runtype will define which information needs to be exported. Possible options are: user, device and useranddevice.

    .EXAMPLE
        Get-ApplicationDeploymentStatus.ps1 -TenantName domain.onmicrsoft.com -ClientID "<GUID>" -ExportPath "c:\temp" -runtype useranddevice

    .NOTES
    Author:      Arjan Vroege
    Contact:     @ArjanVroege
    Created:     2017-09-17

    Version history:
    1.0.0 - (2017-09-17) Script created
    1.1.0 - (2017-09-20) Added User Deployment information
#>
[CmdletBinding()]
 param(
        [parameter(Mandatory=$true, HelpMessage="A tenant name should be provided in the following format: tenantname.onmicrosoft.com.")]
        [ValidateNotNullOrEmpty()]
        [string]$TenantName,

        [parameter(Mandatory=$true, HelpMessage="Application ID for an Azure AD application.")]
        [ValidateNotNullOrEmpty()]
        [string]$ClientID,

        [parameter(Mandatory=$true, HelpMessage="Defines which type of data need to be exported. Possible options are: 'user', 'device' and 'useranddevice'.")]
        [ValidateNotNullOrEmpty()]
        [string]$runtype,

        [parameter(Mandatory=$true, HelpMessage="The location where the exported files need to be saved")]
        [ValidateNotNullOrEmpty()]
        [string]$exportpath = "c:\temp"
 )

. "< Location of the Get-MSGraphAuthenticationToken.ps1 file, example: c:\temp\Get-MSGraphAuthenticationToken.ps1 >"

function get_grp_member($GroupId) {
    [int]$AssignedUsers = 0;

    try {
        $Resource    = "groups/$GroupId/members/"
        $uri         = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"
        $grp_members = (Invoke-RestMethod -Uri $uri –Headers $AuthenticationHeader –Method Get).Value
    }

    catch {
        $ex = $_.Exception
        Write-Host "Request to $Uri failed with HTTP Status $([int]$ex.Response.StatusCode) $($ex.Response.StatusDescription)" -f Red
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        write-host
        break
    }

    foreach( $member in $grp_members) {
        if ( $member.'@odata.type' -eq '#microsoft.graph.user') {
            $AssignedUsers++
        } elseif($member.'@odata.type' -eq '#microsoft.graph.group') {
            $groupmemberid = $member.id

            try {
                $Resource    = "groups/$groupmemberid/members/"
                $uri         = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"
                $grp_members1= (Invoke-RestMethod -Uri $uri –Headers $AuthenticationHeader –Method Get).Value
            }

            catch {
                $ex = $_.Exception
                Write-Host "Request to $Uri failed with HTTP Status $([int]$ex.Response.StatusCode) $($ex.Response.StatusDescription)" -f Red
                $errorResponse = $ex.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($errorResponse)
                $reader.BaseStream.Position = 0
                $reader.DiscardBufferedData()
                $responseBody = $reader.ReadToEnd();
                Write-Host "Response content:`n$responseBody" -f Red
                Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
                write-host
                break
            }

            foreach( $member1 in $grp_members1) {
                if ( $member1.'@odata.type' -eq '#microsoft.graph.user') {
                    $AssignedUsers++
                }
            }
        }
    }

    return $AssignedUsers
}

$AuthenticationHeader = Get-MSGraphAuthenticationToken -TenantName $TenantName -ClientID $ClientID
$graphApiVersion      = "beta"
$Resource             = "deviceAppManagement/mobileApps"
$AppStatusDProcessed  = @()
$AppStatusUProcessed  = @()
$AppDeplDStatistics   = @()
$AppDeplUStatistics   = @()
$AppDepDStatus_csv    = $exportpath + "\App_Depl_Dev_Stat_Export_" + $(get-date -f dd-MM-yyyy-H-mm-ss) + ".csv"
$AppDepUStatus_csv    = $exportpath + "\App_Depl_Usr_Stat_Export_" + $(get-date -f dd-MM-yyyy-H-mm-ss) + ".csv"
$AppDepStatus_html    = $exportpath + "\App_Depl_Stat_Export_" + $(get-date -f dd-MM-yyyy-H-mm-ss) + ".html"

try {
    $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"
    $apps = (Invoke-RestMethod -Uri $uri –Headers $AuthenticationHeader –Method Get).Value | Where-Object { (!($_.'@odata.type').Contains("managed")) -and (!($_.'@odata.type').Contains("#microsoft.graph.iosVppApp")) }
}

catch {
    $ex = $_.Exception
    Write-Host "Request to $Uri failed with HTTP Status $([int]$ex.Response.StatusCode) $($ex.Response.StatusDescription)" -f Red
    $errorResponse = $ex.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();
    Write-Host "Response content:`n$responseBody" -f Red
    Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    write-host
    break
}

foreach( $app in $apps) {
    $AppDisplayName    = $app.displayName
    $AppId             = $app.id
    $AppType           = $app.'@odata.type'
    $AppDeployedTotal  = @()
    $AppDeployedNotIns = @()
    $AppDeployedFailed = @()
    $Users              = 0;
    
    if( $app.'@odata.type' -eq '#microsoft.graph.officeSuiteApp') {
        $AppType       = "Office365 App"
    } elseif ( $app.'@odata.type' -eq '#microsoft.graph.windowsMobileMSI') {
        $AppType       = "MDM MSI"
    } elseif ( $app.'@odata.type' -eq '#microsoft.graph.windowsStoreForBusinessApp') {
        $AppType       = "WSfB App"
    } elseif ( $app.'@odata.type' -eq '#microsoft.graph.windowsUniversalAppX') {
        $AppType       = "Universal App"
    } elseif ( $app.'@odata.type' -eq '#microsoft.graph.webApp') {
        $AppType       = "Web App"
    } elseif ( $app.'@odata.type' -eq '#microsoft.graph.windowsStoreApp') {
        $AppType       = "Store App"
    } else {
        $AppType       = "Unknown"
    }

    Write-Host "Retrieving data for application: $AppDisplayName ($appid)"
    
    if(($runtype -eq 'device') -or ($runtype -eq 'useranddevice')) { 
        Write-Host "Retrieving device deployment data for application: $AppDisplayName ($appid)"   
        try {
            $Resource    = "deviceAppManagement/mobileApps/$AppId/deviceStatuses/"
            $uri         = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"
            $appdstatus  = (Invoke-RestMethod -Uri $uri –Headers $AuthenticationHeader –Method Get).Value
        }

        catch {
            $ex = $_.Exception
            Write-Host "Request to $Uri failed with HTTP Status $([int]$ex.Response.StatusCode) $($ex.Response.StatusDescription)" -f Red
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
            Write-Host "Response content:`n$responseBody" -f Red
            Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
            write-host
            break
        }
    
        foreach($statusd in $appdstatus) {
            $statusd | Add-Member -MemberType NoteProperty -Name 'AppName' -Value $AppDisplayName
            $statusd | Add-Member -MemberType NoteProperty -Name 'AppType' -Value $AppType
            $AppStatusDProcessed += $statusd
        }

        Write-Host "Exporting detailed data for application: $AppDisplayName ($appid) to $AppDepDStatus_csv"
        $AppStatusDProcessed | Where-Object { ($_.AppName -eq $AppDisplayName) -and ($_.mobileAppInstallStatusValue -ne 'notApplicable')} | select AppName,AppType,userPrincipalName, UserName, OSDescription, OSVersion, mobileAppInstallStatusValue, errorcode, lastsyncdatatime, deviceName, deviceid | Export-Csv -Path $AppDepDStatus_csv -Delimiter "," -NoTypeInformation -Append
    
        Write-Host "Generating Device Deployment statistics for application: $AppDisplayName ($appid)"
        $AppDeployedTotal           = (($AppStatusDProcessed | Where-Object { ($_.AppName -eq $AppDisplayName) -and ($_.mobileAppInstallStatusValue -ne 'notApplicable')}) | Measure-Object).Count
        $AppDeployedSucces          = (($AppStatusDProcessed | Where-Object { ($_.AppName -eq $AppDisplayName) -and ($_.mobileAppInstallStatusValue -eq 'installed')}) | Measure-Object).Count
        $AppDeployedFailed          = (($AppStatusDProcessed | Where-Object { ($_.AppName -eq $AppDisplayName) -and ($_.mobileAppInstallStatusValue -eq 'failed')}) | Measure-Object).Count
        $AppDeployedNotIns          = (($AppStatusDProcessed | Where-Object { ($_.AppName -eq $AppDisplayName) -and ($_.mobileAppInstallStatusValue -eq "notInstalled")}) | Measure-Object).Count
    
        if($AppDeployedTotal -gt 0) { 
            $AppDeployedSuccesRate  = ($AppDeployedSucces / $AppDeployedTotal).ToString("P")
            $AppDeployedFailedRate  = ($AppDeployedFailed / $AppDeployedTotal).ToString("P")
        } else {
            $AppDeployedSuccesRate  = '0,00 %'
            $AppDeployedFailedRate  = '0,00 %'
        }

        $props = @{
            AppName                 = $AppDisplayName
            AppType                 = $AppType
            AppDeployedTotal        = $AppDeployedTotal
            AppDeployedSucces       = $AppDeployedSucces
            AppDeployedFailed       = $AppDeployedFailed
            AppDeployedNotIns       = $AppDeployedNotIns
            AppDeployedSuccesRate   = $AppDeployedSuccesRate
            AppDeployedFailedRate   = $AppDeployedFailedRate
        }
        $ServiceObject = New-Object -TypeName PSObject -Property $props
        $AppDeplDStatistics += $ServiceObject
    }


    if (($runtype -eq 'user') -or ($runtype -eq 'useranddevice')) {
        Write-Host "Retrieving user deployment data for application: $AppDisplayName ($appid)"   
        try {
            $Resource    = "deviceAppManagement/mobileApps/$AppId/userStatuses/"
            $uri         = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"
            $appustatus = (Invoke-RestMethod -Uri $uri –Headers $AuthenticationHeader –Method Get).Value
        }

        catch {
            $ex = $_.Exception
            Write-Host "Request to $Uri failed with HTTP Status $([int]$ex.Response.StatusCode) $($ex.Response.StatusDescription)" -f Red
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
            Write-Host "Response content:`n$responseBody" -f Red
            Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
            write-host
            break
        }

        foreach($statusU in $appustatus) {
            $statusU | Add-Member -MemberType NoteProperty -Name 'AppName' -Value $AppDisplayName
            $statusU | Add-Member -MemberType NoteProperty -Name 'AppType' -Value $AppType
            $AppStatusUProcessed += $statusU
        }

        Write-Host "Exporting detailed data for application: $AppDisplayName ($appid) to $AppDepUStatus_csv"
        $AppStatusUProcessed | Where-Object { ($_.AppName -eq $AppDisplayName)} | select AppName, AppType, userPrincipalName, UserName, installedDeviceCount, failedDeviceCount, notInstalledDeviceCount | Export-Csv -Path $AppDepUStatus_csv -Delimiter "," -NoTypeInformation -Append

        Write-Host "Retrieving user group assignments for application: $AppDisplayName ($appid)"   
        try {
            $Resource    = "deviceAppManagement/mobileApps/$AppId/groupAssignments/"
            $uri         = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"
            $appassing   = (Invoke-RestMethod -Uri $uri –Headers $AuthenticationHeader –Method Get).Value
        }

        catch {
            $ex = $_.Exception
            Write-Host "Request to $Uri failed with HTTP Status $([int]$ex.Response.StatusCode) $($ex.Response.StatusDescription)" -f Red
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
            Write-Host "Response content:`n$responseBody" -f Red
            Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
            write-host
            break
        }

        foreach($group in $appassing) {
            $GroupId   = $group.targetGroupId
            Write-Host "Retrieving Group Members of Group $GroupId" 
            $users += get_grp_member $GroupId
            
        }
        
        Write-Host "Generating Device Deployment statistics for application: $AppDisplayName ($appid)"
        [int]$AppAssignedTotal           = (($AppStatusUProcessed | Where-Object { ($_.AppName -eq $AppDisplayName)}) | Measure-Object).Count

        if($AppAssignedTotal -gt 0) {
            $AppAssignedInstalled       = (($AppStatusUProcessed | Where-Object { ($_.AppName -eq $AppDisplayName)}) | Measure-Object -Property installedDeviceCount -Sum).Sum
            $AppAssignedFailed          = (($AppStatusUProcessed | Where-Object { ($_.AppName -eq $AppDisplayName)}) | Measure-Object -Property failedDeviceCount -Sum).Sum
            $AppAssignedNotIns          = (($AppStatusUProcessed | Where-Object { ($_.AppName -eq $AppDisplayName)}) | Measure-Object -Property notInstalledDeviceCount -Sum).Sum
        } else {
            $AppAssignedInstalled       = '-'
            $AppAssignedFailed          = '-'
            $AppAssignedNotIns          = '-'
        }
        
        $props = @{
            AppName                 = $AppDisplayName
            AppType                 = $AppType
            UserAssignments         = $users
            UserTotalDeployed       = $AppAssignedTotal
            UserTotalInstalled      = $AppAssignedInstalled
            UserTotalFailed         = $AppAssignedFailed
            UserTotalNotIns         = $AppAssignedNotIns
        }
        $ServiceObject = New-Object -TypeName PSObject -Property $props
        $AppDeplUStatistics += $ServiceObject
    }
    Write-Host ""
}

$html = '<style>'
$html = $html + 'BODY{background-color:#FAFAFA;font-family:"Trebuchet MS", Helvetica, sans-serif}'
$html = $html + 'TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}'
$html = $html + 'TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:#6E6E6E}'
$html = $html + 'TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:#D8D8D8;text-align:center;}'
$html = $html + '</style>'
$html = $html + '<head><title>Intune Application Deployment Summary</title></head>'

$AppDeplDStatistics | select AppName,AppType,AppDeployedSucces,AppDeployedFailed,AppDeployedNotIns,AppDeployedTotal,AppDeployedSuccesRate,AppDeployedFailedRate | Sort-Object AppName | ConvertTo-HTML -head $html -body "<H2>Intune Application Device Deployment Status Overview</H2>" | Out-File $AppDepStatus_html

$AppDeplUStatistics | select AppName,AppType,UserAssignments,UserTotalDeployed,UserTotalInstalled,UserTotalFailed,UserTotalNotIns | Sort-Object AppName | ConvertTo-HTML -body "<H2>Intune Application User Deployment Status Overview</H2>" | Out-File $AppDepStatus_html -Append    