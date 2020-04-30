<#
v 1.0.5
#>

[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null

Add-Type -AssemblyName PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing, System.Windows.Forms.DataVisualization

$script:hash = @{}

#region functions

function Get-AuthToken {


[cmdletbinding()]

param
(
    [Parameter(Mandatory=$true)]
    $User
)

$userUpn = New-Object "System.Net.Mail.MailAddress" -ArgumentList $User

$tenant = $userUpn.Host

    $AadModule = Get-Module -Name "AzureAD" -ListAvailable

    if ($null -eq $AadModule) {

        $AadModule = Get-Module -Name "AzureADPreview" -ListAvailable

    }

    if ($null -eq $AadModule) {
        [System.Windows.MessageBox]::Show("Install by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt. Script can't continue..."  , 'AzureAD Powershell module not installed...','OK','Warning')
    }


    if($AadModule.count -gt 1){

        $Latest_Version = ($AadModule | Select-Object version | Sort-Object)[-1]

        $aadModule = $AadModule | Where-Object { $_.version -eq $Latest_Version.version }

            # Checking if there are multiple versions of the same module found

            if($AadModule.count -gt 1){

            $aadModule = $AadModule | Select-Object -Unique

            }

        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"

    }

    else {

        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"

    }

[System.Reflection.Assembly]::LoadFrom($adal) | Out-Null

[System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null

$clientId = "d1ddf0e4-d672-4dae-b554-9d5bdfd93547"

$redirectUri = "urn:ietf:wg:oauth:2.0:oob"

$resourceAppIdURI = "https://graph.microsoft.com"

$authority = "https://login.microsoftonline.com/$Tenant"

    try {

    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority

    # https://msdn.microsoft.com/en-us/library/azure/microsoft.identitymodel.clients.activedirectory.promptbehavior.aspx
    # Change the prompt behaviour to force credentials each time: Auto, Always, Never, RefreshSession

    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Auto"

    $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")

    $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI,$clientId,$redirectUri,$platformParameters,$userId).Result

        # If the accesstoken is valid then create the authentication header

        if($authResult.AccessToken){

        # Creating header for Authorization token

        $authHeader = @{
            'Content-Type'='application/json'
            'Authorization'="Bearer " + $authResult.AccessToken
            'ExpiresOn'=$authResult.ExpiresOn
            }

        return $authHeader

        }

        else {
             [System.Windows.MessageBox]::Show("Authorization Access Token is null, please re-run authentication..."  , 'Authorization Access Token is null','OK','Warning')

        }

    }

    catch {

       [System.Windows.MessageBox]::Show($_.Exception.ItemName , 'Get-AuthToken','OK','Warning')
    

    }

}

function Get-AADUser(){

[cmdletbinding()]

param
(
    $userPrincipalName,
    $Property
)

$graphApiVersion = "v1.0"
$User_resource = "users"

    try {

        if($userPrincipalName -eq "" -or $null -eq $userPrincipalName){

        $uri = "https://graph.microsoft.com/$graphApiVersion/$($User_resource)"
        (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).Value

        }

        else {

            if($Property -eq "" -or $null -eq $Property){

            $uri = "https://graph.microsoft.com/$graphApiVersion/$($User_resource)/$userPrincipalName"
            Write-Verbose $uri
            Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get

            }

            else {

            $uri = "https://graph.microsoft.com/$graphApiVersion/$($User_resource)/$userPrincipalName/$Property"
            Write-Verbose $uri
            (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).Value

            }

        }

    }

    catch {

    $ex = $_.Exception
    $errorResponse = $ex.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();
    Write-Host "Response content:`n$responseBody" -f Red
    Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    [System.Windows.MessageBox]::Show('Get-AADUser  ' + $responseBody , 'Request failed','OK','Error')

    }

}

function Get-AADGroup(){

[cmdletbinding()]

param
(
    $GroupName,
    $id,
    [switch]$Members
)

$graphApiVersion = "v1.0"
$Group_resource = "groups"
    
    try {

        if($id){

        $uri = "https://graph.microsoft.com/$graphApiVersion/$($Group_resource)?`$filter=id eq '$id'"
        (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).Value

        }
        
        elseif($GroupName -eq "" -or $null -eq $GroupName){
        
        $uri = "https://graph.microsoft.com/$graphApiVersion/$($Group_resource)"
        (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).Value
        
        }

        else {
            
            if(!$Members){

            $uri = "https://graph.microsoft.com/$graphApiVersion/$($Group_resource)?`$filter=displayname eq '$GroupName'"
            (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).Value
            
            }
            
            elseif($Members){
            
            $uri = "https://graph.microsoft.com/$graphApiVersion/$($Group_resource)?`$filter=displayname eq '$GroupName'"
            $Group = (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).Value
            
                if($Group){

                $GID = $Group.id

                $Group.displayName
                write-host

                $uri = "https://graph.microsoft.com/$graphApiVersion/$($Group_resource)/$GID/Members"
                (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).Value

                }

            }
        
        }

    }

    catch {

    $ex = $_.Exception
    $errorResponse = $ex.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();
    Write-Host "Response content:`n$responseBody" -f Red
    Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    [System.Windows.MessageBox]::Show('Get-AADGroup  ' + $responseBody , 'Request failed','OK','Error')

    }

}

function Add-AADGroupMember(){


[cmdletbinding()]

param
(
    $GroupId,
    $AADMemberId
)

$graphApiVersion = "v1.0"
$Resource = "groups"
    
    try {

    $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$GroupId/members/`$ref"

$JSON = @"
{
    "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/$AADMemberId"
}
"@

    Invoke-RestMethod -Uri $uri -Headers $authToken -Method Post -Body $Json -ContentType "application/json"

    }

    catch {

    $ex = $_.Exception
    $errorResponse = $ex.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();
    Write-Host "Response content:`n$responseBody" -f Red
    Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    [System.Windows.MessageBox]::Show('Add-AADGroupMember  ' + $responseBody , 'Request failed','OK','Error')

    }

}

function Remove-AADGroupMember(){


[cmdletbinding()]

param
(
    $GroupId,
    $AADMemberId
)

$graphApiVersion = "v1.0"
$Resource = "groups"
    
    try {

    $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$GroupId/members/$AADMemberId/`$ref"


    Invoke-RestMethod -Uri $uri -Headers $authToken -Method Delete -ContentType "application/json"

    }

    catch {

    $ex = $_.Exception
    $errorResponse = $ex.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();
    #Write-Host "Response content:`n$responseBody" -f Red
    #SWrite-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    [System.Windows.MessageBox]::Show('Remove-AADGroupMember  ' + $responseBody , 'Request failed','OK','Error')

    }

}

function New-AADGroup(){


    [cmdletbinding()]
    
    param
    (
       [Parameter(Mandatory = $true)]
       [ValidateNotNullOrEmpty()]
       [ValidateSet('Office', 'OfficeDynamic','Security')]
       [String]$GroupType,
       [string]$description,
       [string]$displayName,
       [string]$mailEnabled = $false,
       [string]$mailNickname = $displayName,
       [string]$securityEnabled = $false,
       [string]$membershipRule
    )

         try {
    
          $uri = "https://graph.microsoft.com/v1.0/groups"

          if($GroupType -eq "Office" -or $GroupType -eq "OfficeDynamic"){ 

            if ($GroupType -eq "Office"){

              $flag = "Office"

              $Json = @{
                  "description"     = $description
                  "displayName"     = $displayName
                  "groupTypes"      = @("Unified")
                  "mailEnabled"     = $mailEnabled
                  "mailNickname"    = $mailNickname
                  "securityEnabled" = $securityEnabled
              } | ConvertTo-Json
              
            }
            elseif($GroupType -eq "OfficeDynamic"){
             
              $flag = "OfficeDynamic"

              $Json = @{
                  "description"     = $description
                  "displayName"     = $displayName
                  "groupTypes"      = @("Unified","DynamicMembership")
                  "mailEnabled"     = $mailEnabled
                  "mailNickname"    = $mailNickname
                  "membershipRule"  = $membershipRule 
                  "membershipRuleProcessingState" = "On"
                  "securityEnabled" = $securityEnabled
              } | ConvertTo-Json
 
            }
        
          }
          elseif($GroupType -eq "Security" -and $membershipRule -ne ""){  
            
            $flag = "Security with Rule"

            $JSON = @{
                "description"    = $description
                "displayName"    = $displayName
                "groupTypes"     = @("DynamicMembership")
                "mailEnabled"    = "false"
                "mailNickname"   = $displayName
                "membershipRule" = $membershipRule 
                "membershipRuleProcessingState" = "On"
                "securityEnabled" = $securityEnabled
            } | ConvertTo-Json

          }
          elseif($GroupType -eq "Security" -and $membershipRule -eq ""){

            $flag = "Security without Rule"

            $JSON = @{
                "description"     = $description
                "displayName"     = $displayName
                "groupTypes"      = @("DynamicMembership")
                "mailEnabled"     = "false"
                "mailNickname"    = $displayName
                "securityEnabled" = $securityEnabled
            } | ConvertTo-Json

          }
          elseif($GroupType -eq "Security" -and $global:securityEnabled.IsChecked -eq $false){

            $flag = "Security without Security"

            $JSON = @{
                "description"     = $description
                "displayName"     = $displayName
                "groupTypes"      = @("DynamicMembership")
                "securityEnabled" = "false"
                "mailEnabled"     = "false"
            } | ConvertTo-Json

          }          
          

                Invoke-RestMethod -Uri $uri -Headers $authToken -Method Post -Body $Json -ContentType "application/json"
                $global:Create_Group_Result.Visibility = "Visible"
                $global:Create_Group_Result.Content="The group was successfully created"
                $global:Create_Group_Result.Foreground="Green"
        }
    
        catch {
    
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        [System.Windows.MessageBox]::Show("New-AADGroup ($($flag)) " + $responseBody , 'Request failed','OK','Error')
    
        }
    
}

function Get-AADDevice(){

     
    [cmdletbinding()]
    
    param
    (
        $DeviceID
    )
    
    # Defining Variables
    $graphApiVersion = "v1.0"
    $Resource = "devices"
        
        try {

        if ($DeviceID -eq "" -or $null -eq $DeviceID){
            $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"
    
            (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).value 
        }
        else{
            $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)?`$filter=deviceId eq '$DeviceID'"
    
            (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).value 
        }

        
    
        }
    
        catch {
    
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        [System.Windows.MessageBox]::Show('Get-AADDevice  ' + $responseBody , 'Request failed','OK','Error')
    
        }
}

function Get-TennatInfo(){
  
        try {
        
            $uri = "https://graph.microsoft.com/v1.0/organization"
    
            (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).value 
 
        }
    
        catch {
    
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        write-host
        [System.Windows.MessageBox]::Show('Get-TennatInfo  ' + $responseBody , 'Request failed','OK','Error')
    
        }
}

function Get-Registration(){
  
    try {
    
        $uri = "https://graph.microsoft.com/beta/reports/getCredentialUserRegistrationCount"

        (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).value 

    }

    catch {

    $ex = $_.Exception
    $errorResponse = $ex.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();
    Write-Host "Response content:`n$responseBody" -f Red
    Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    write-host
    [System.Windows.MessageBox]::Show('Get-RegistrationInfo  ' + $responseBody , 'Request failed','OK','Error')
   

    }
}

function Get-Subscription(){
  
    try {
    
        $uri = "https://graph.microsoft.com/beta/subscribedSkus"

        (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).value 

    }

    catch {

    $ex = $_.Exception
    $errorResponse = $ex.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();
    Write-Host "Response content:`n$responseBody" -f Red
    Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    [System.Windows.MessageBox]::Show('Get-Subscription  ' + $responseBody , 'Request failed','OK','Error')

    }
}

function Get-ManagedDevices(){

       
    [cmdletbinding()]
    
    param
    (
        [switch]$IncludeEAS,
        [switch]$ExcludeMDM
    )
    
   
    $graphApiVersion = "beta"
    $Resource = "deviceManagement/managedDevices"
    
    try {
    
        $Count_Params = 0
    
        if($IncludeEAS.IsPresent){ $Count_Params++ }
        if($ExcludeMDM.IsPresent){ $Count_Params++ }
            
            if($Count_Params -gt 1){
    
            write-warning "Multiple parameters set, specify a single parameter -IncludeEAS, -ExcludeMDM or no parameter against the function"
            Write-Host
            break
    
            }
            
            elseif($IncludeEAS){
    
            $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource"
    
            }
    
            elseif($ExcludeMDM){
    
            $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource`?`$filter=managementAgent eq 'eas'"
    
            }
            
            else {
        
            $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource`?`$filter=managementAgent eq 'mdm' and managementAgent eq 'easmdm'"
            Write-Warning "EAS Devices are excluded by default, please use -IncludeEAS if you want to include those devices"
            Write-Host
    
            }
    
            (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).Value
        
        }
    
        catch {
    
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        [System.Windows.MessageBox]::Show('Get-ManagedDevices  ' + $responseBody , 'Request failed','OK','Error')
    
        }
}
    
function Show-Toast{

[cmdletbinding()]

param
(
    $Message,
    $Title

)

$AppID   = "Microsoft.AutoGenerated.{923DD477-5846-686B-A659-0FCCD73851A8}" 

$template = @"
<toast launch="action=viewAlarm&amp;alarmId=3" scenario="alarm">

  <visual>
    <binding template="ToastGeneric">
      <text>$Title</text>
      <text>$Message</text>
    </binding>
  </visual>

 </toast>
"@

$xml = New-Object Windows.Data.Xml.Dom.XmlDocument
$xml.LoadXml($template)


$toast = New-Object Windows.UI.Notifications.ToastNotification $xml
  
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppID).Show($toast)

}

function New-Chart() {
    param(
        [hashtable]$Params,
        [string]$ChartTitle,
        [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]$Type,
        $PieLabelStyle = "Outside",
        $Color = 'Black'
    )
 
    
    $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $Chart.Width = 430
    $Chart.Height = 330
    $Chart.Left = 10
    $Chart.Top = 10
 
    
    $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $Chart.ChartAreas.Add($ChartArea)
    $Chart.ChartAreas[0].Area3DStyle.Enable3D = "True"
    [void]$Chart.Series.Add("Data") 
 
    
    $Params.GetEnumerator() | ForEach-Object {
        $datapoint = new-object System.Windows.Forms.DataVisualization.Charting.DataPoint(0, $_.Value.Value)
        $datapoint.AxisLabel = "$($_.Value.Header)" + " (" + $($_.Value.Value) + ")"
        $Chart.Series["Data"].Points.Add($datapoint)
    }
 
    $Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::$Type
    $Chart.Series["Data"]["PieLabelStyle"] = $PieLabelStyle
    $Chart.Series["Data"].Font = "Segoe UI, 9pt"
    $Chart.Series["Data"]["PieLineColor"] = "Black"
    $Chart.Series["Data"].LabelForeColor = $Color
    $Chart.Series["Data"]["PieDrawingStyle"] = "Default" #"Concave"
    ($Chart.Series["Data"].Points.FindMaxByValue())["Exploded"] = $true
 
    $Title = new-object System.Windows.Forms.DataVisualization.Charting.Title
    $Chart.Titles.Add($Title)
    $Chart.Titles[0].Font = "Segoe UI, 14pt"
    $Chart.Titles[0].Text = $ChartTitle
 
    $Stream = New-Object System.IO.MemoryStream
    $Chart.SaveImage($Stream,"png")
    $script:hash.Stream = $Stream.GetBuffer()
    $Stream.Dispose()
}

function ConvertSKUto-FrendlyName{
    param(
        [string]$SKU
    )
    
    switch ($SKU){
     "O365_BUSINESS_ESSENTIALS"     {"Office 365 Business Essentials"}
     "O365_BUSINESS_PREMIUM"        {"Office 365 Business Premium"}
     "DESKLESSPACK"                 {"Office 365 (Plan K1)"}
     "DESKLESSWOFFPACK"             {"Office 365 (Plan K2)"}
     "LITEPACK"                     {"Office 365 (Plan P1)"}
     "EXCHANGESTANDARD"             {"Office 365 Exchange Online Only"}
     "STANDARDPACK"                 {"Enterprise Plan E1"}
     "STANDARDWOFFPACK"             {"Office 365 (Plan E2)"}
     "ENTERPRISEPACK"               {"Enterprise Plan E3"}
     "ENTERPRISEPACKLRG"            {"Enterprise Plan E3"}
     "ENTERPRISEWITHSCAL"           {"Enterprise Plan E4"}
     "STANDARDPACK_STUDENT"         {"Office 365 (Plan A1) for Students"}
     "STANDARDWOFFPACKPACK_STUDENT" {"Office 365 (Plan A2) for Students"}
     "ENTERPRISEPACK_STUDENT"       {"Office 365 (Plan A3) for Students"}
     "ENTERPRISEWITHSCAL_STUDENT"   {"Office 365 (Plan A4) for Students"}
     "STANDARDPACK_FACULTY"         {"Office 365 (Plan A1) for Faculty"}
     "STANDARDWOFFPACKPACK_FACULTY" {"Office 365 (Plan A2) for Faculty"}
     "ENTERPRISEPACK_FACULTY"       {"Office 365 (Plan A3) for Faculty"}
     "ENTERPRISEWITHSCAL_FACULTY"   {"Office 365 (Plan A4) for Faculty"}
     "ENTERPRISEPACK_B_PILOT"       {"Office 365 (Enterprise Preview)"}
     "STANDARD_B_PILOT"             {"Office 365 (Small Business Preview)"}
     "VISIOCLIENT"                  {"Visio Pro Online"}
     "POWER_BI_ADDON"               {"Office 365 Power BI Addon"}
     "POWER_BI_INDIVIDUAL_USE"      {"Power BI Individual User"}
     "POWER_BI_STANDALONE"          {"Power BI Stand Alone"}
     "POWER_BI_STANDARD"            {"Power-BI Standard"}
     "PROJECTESSENTIALS"            {"Project Lite"}
     "PROJECTCLIENT"                {"Project Professional"}
     "PROJECTONLINE_PLAN_1"         {"Project Online"}
     "PROJECTONLINE_PLAN_2"         {"Project Online and PRO"}
     "ProjectPremium"               {"Project Online Premium"}
     "ECAL_SERVICES"                {"ECAL"}
     "EMS"                          {"Enterprise Mobility Suite"}
     "RIGHTSMANAGEMENT_ADHOC"       {"Windows Azure Rights Management"}
     "MCOMEETADV"                   {"PSTN conferencing"}
     "SHAREPOINTSTORAGE"            {"SharePoint storage"}
     "PLANNERSTANDALONE"            {"Planner Standalone"}
     "CRMIUR"                       {"CMRIUR"}
     "BI_AZURE_P1"                  {"Power BI Reporting and Analytics"}
     "INTUNE_A"                     {"Windows Intune Plan A"}
     "PROJECTWORKMANAGEMENT"        {"Office 365 Planner Preview"}
     "ATP_ENTERPRISE"               {"Exchange Online Advanced Threat Protection"}
     "EQUIVIO_ANALYTICS"            {"Office 365 Advanced eDiscovery"}
     "AAD_BASIC"                    {"Azure Active Directory Basic"}
     "RMS_S_ENTERPRISE"             {"Azure Active Directory Rights Management"}
     "AAD_PREMIUM"                  {"Azure Active Directory Premium"}
     "MFA_PREMIUM"                  {"Azure Multi-Factor Authentication"}
     "STANDARDPACK_GOV"             {"Microsoft Office 365 (Plan G1) for Government"}
     "STANDARDWOFFPACK_GOV"         {"Microsoft Office 365 (Plan G2) for Government"}
     "ENTERPRISEPACK_GOV"           {"Microsoft Office 365 (Plan G3) for Government"}
     "ENTERPRISEWITHSCAL_GOV"       {"Microsoft Office 365 (Plan G4) for Government"}
     "DESKLESSPACK_GOV"             {"Microsoft Office 365 (Plan K1) for Government"}
     "ESKLESSWOFFPACK_GOV"          {"Microsoft Office 365 (Plan K2) for Government"}
     "EXCHANGESTANDARD_GOV"         {"Microsoft Office 365 Exchange Online (Plan 1) only for Government"}
     "EXCHANGEENTERPRISE_GOV"       {"Microsoft Office 365 Exchange Online (Plan 2) only for Government"}
     "SHAREPOINTDESKLESS_GOV"       {"SharePoint Online Kiosk"}
     "EXCHANGE_S_DESKLESS_GOV"      {"Exchange Kiosk"}
     "RMS_S_ENTERPRISE_GOV"         {"Windows Azure Active Directory Rights Management"}
     "OFFICESUBSCRIPTION_GOV"       {"Office ProPlus"}
     "MCOSTANDARD_GOV"              {"Lync Plan 2G"}
     "SHAREPOINTWAC_GOV"            {"Office Online for Government"}
     "SHAREPOINTENTERPRISE_GOV"     {"SharePoint Plan 2G"}
     "EXCHANGE_S_ENTERPRISE_GOV"    {"Exchange Plan 2G"}
     "EXCHANGE_S_ARCHIVE_ADDON_GOV" {"Exchange Online Archiving"}
     "EXCHANGE_S_DESKLESS"          {"Exchange Online Kiosk"}
     "SHAREPOINTDESKLESS"           {"SharePoint Online Kiosk"}
     "SHAREPOINTWAC"                {"Office Online"}
     "YAMMER_ENTERPRISE"            {"Yammer Enterprise"}
     "EXCHANGE_L_STANDARD"          {"Exchange Online (Plan 1)"}
     "MCOLITE"                      {"Lync Online (Plan 1)"}
     "SHAREPOINTLITE"               {"SharePoint Online (Plan 1)"}
     "OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ" {"Office ProPlus"}
     "EXCHANGE_S_STANDARD_MIDMARKET"      {"Exchange Online (Plan 1)"}
     "MCOSTANDARD_MIDMARKET"        {"Lync Online (Plan 1)"}
     "SHAREPOINTENTERPRISE_MIDMARKET" {"SharePoint Online (Plan 1)"}
     "OFFICESUBSCRIPTION"           {"Office ProPlus"}
     "YAMMER_MIDSIZE"               {"Yammer"}
     "DYN365_ENTERPRISE_PLAN1"      {"Dynamics 365 Customer Engagement Plan Enterprise Edition"}
     "ENTERPRISEPREMIUM_NOPSTNCONF" {"Enterprise E5 (without Audio Conferencing)"}
     "ENTERPRISEPREMIUM"            {"Enterprise E5 (with Audio Conferencing)"}
     "MCOSTANDARD"                  {"Skype for Business Online Standalone Plan 2"}
     "PROJECT_MADEIRA_PREVIEW_IW_SKU" {"Dynamics 365 for Financials for IWs"}
     "STANDARDWOFFPACK_IW_STUDENT"  {"Office 365 Education for Students"}
     "STANDARDWOFFPACK_IW_FACULTY"  {"Office 365 Education for Faculty"}
     "EOP_ENTERPRISE_FACULTY"       {"Exchange Online Protection for Faculty"}
     "EXCHANGESTANDARD_STUDENT"     {"Exchange Online (Plan 1) for Students"}
     "OFFICESUBSCRIPTION_STUDENT"   {"Office ProPlus Student Benefit"}
     "STANDARDWOFFPACK_FACULTY"     {"Office 365 Education E1 for Faculty"}
     "STANDARDWOFFPACK_STUDENT"     {"Microsoft Office 365 (Plan A2) for Students"}
     "DYN365_FINANCIALS_BUSINESS_SKU" {"Dynamics 365 for Financials Business Edition"}
     "DYN365_FINANCIALS_TEAM_MEMBERS_SKU" {"Dynamics 365 for Team Members Business Edition"}
     "FLOW_FREE"                    {"Microsoft Flow Free"}
     "POWER_BI_PRO"                 {"Power BI Pro"}
     "O365_BUSINESS"                {"Office 365 Business"}
     "DYN365_ENTERPRISE_SALES"      {"Dynamics Office 365 Enterprise Sales"}
     "RIGHTSMANAGEMENT"             {"Rights Management"}
     "PROJECTPROFESSIONAL"          {"Project Professional"}
     "VISIOONLINE_PLAN1"            {"Visio Online Plan 1"}
     "EXCHANGEENTERPRISE"           {"Exchange Online Plan 2"}
     "DYN365_ENTERPRISE_P1_IW"      {"Dynamics 365 P1 Trial for Information Workers"}
     "DYN365_ENTERPRISE_TEAM_MEMBERS" {"Dynamics 365 For Team Members Enterprise Edition"}
     "CRMSTANDARD"      {"Microsoft Dynamics CRM Online Professional"}
     "EXCHANGEARCHIVE_ADDON"        {"Exchange Online Archiving For Exchange Online"}
     "EXCHANGEDESKLESS"             {"Exchange Online Kiosk"}
     "SPZA_IW"                      {"App Connect"}
     "WINDOWS_STORE"                {"Windows Store for Business"}
     "MCOEV"                        {"Microsoft Phone System"}
     "VIDEO_INTEROP"                {"Polycom Skype Meeting Video Interop for Skype for Business"}
     "SPE_E5"                       {"Microsoft 365 E5"}
     "SPE_E3"                       {"Microsoft 365 E3"}
     "ATA"                          {"Advanced Threat Analytics"}
     "MCOPSTN2"                     {"Domestic and International Calling Plan"}
     "FLOW_P1"                      {"Microsoft Flow Plan 1"}
     "FLOW_P2"                      {"Microsoft Flow Plan 2"}
     "CRMSTORAGE"                   {"Microsoft Dynamics CRM Online Additional Storage"}
     "SMB_APPS"                     {"Microsoft Business Apps"}
     "MICROSOFT_BUSINESS_CENTER"    {"Microsoft Business Center"}
     "DYN365_TEAM_MEMBERS"          {"Dynamics 365 Team Members"}
     "STREAM"                       {"Microsoft Stream Trial"}
     "EMSPREMIUM"                   {"ENTERPRISE MOBILITY + SECURITY E5"}
    default { $SKU }
    }
    
 }

#endregion functions

# XAML Main window
[xml]$xaml = @"
 <Window
	 xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
	 xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
	 Name="window" 
	 WindowStyle="None"
     Title="AAD Object Management"
	 WindowState="Normal" 
     ResizeMode="CanResizeWithGrip" 
	 ShowInTaskbar="False"
     WindowStartupLocation="CenterScreen"
     BorderBrush="DarkCyan"
     BorderThickness="1.5"
     AllowsTransparency="True"
     >

     <WindowChrome.WindowChrome>
         <WindowChrome     
             CaptionHeight="1"  
             CornerRadius ="0"
             ResizeBorderThickness="14"         
             GlassFrameThickness="14">
         </WindowChrome>
     </WindowChrome.WindowChrome>

     <Window.Background>
         <LinearGradientBrush EndPoint="0.5,1" MappingMode="RelativeToBoundingBox" StartPoint="0.5,0">
             <GradientStop Color="White" Offset="1"/>
             <GradientStop Color="Azure" Offset="10"/>
    	 </LinearGradientBrush>
     </Window.Background>

<!-- ******************** Style Section  ********************-->

 <Window.Resources>

     <Style TargetType="{x:Type Window}"> 
         <!-- <Setter Property="FontFamily" Value="Jokerman Regular" />  -->   
         <!-- <Setter Property="FontFamily" Value="Castellar Regular" /> --> 
         <!-- <Setter Property="FontFamily" Value="Baskerville Old Face Regular" />  -->    
         <!-- <Setter Property="FontFamily" Value="Microsoft Sans Serif Regular" /> --> 
         <Setter Property="FontFamily" Value="Segoe UI" />      
     </Style> 

     <Style x:Key="ListViewItemStretchStyle" TargetType="ListViewItem">
         <Setter Property="HorizontalContentAlignment" Value="Stretch" />
     </Style>
 
     <Style x:Key="ListBoxItemStretchStyle" TargetType="ListBoxItem">
         <Setter Property="HorizontalContentAlignment" Value="Stretch" />
     </Style>    

<!-- ******************** Window Border Style ************************ --> 

     <Style TargetType="{x:Type Border}">
         <Setter Property="CornerRadius" Value="5"/>
         <Setter Property="Padding" Value="5"/>
     </Style>

<!-- ******************** End Window Border Style ******************** --> 

<!-- ************************** Buttons Style  *********************** --> 

     <Style x:Key="ButtonTemplate" TargetType="Button" >
         <Setter Property="OverridesDefaultStyle" Value="True" />
         <Setter Property="Cursor" Value="Hand" />
         <Setter Property="Template">
             <Setter.Value>
                 <ControlTemplate TargetType="Button">
                     <Border Name="border" BorderThickness="2" BorderBrush="DarkCyan" Background="{TemplateBinding Background}" CornerRadius="3" >
                         <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" />
                     </Border>
                     <ControlTemplate.Triggers>
                         <Trigger Property="IsMouseOver" Value="True">
                             <Setter Property="Opacity" Value="0.5" />
                         </Trigger>
                     </ControlTemplate.Triggers>
                 </ControlTemplate>
             </Setter.Value>
         </Setter>
     </Style>

<!-- ******************** End Buttons Style ********************--> 

<!-- ********************* TextBoxes Style  ********************--> 

     <Style x:Key="TextBoxTemplate" TargetType="TextBox">
         <Setter Property="Template">
             <Setter.Value>
                 <ControlTemplate TargetType="{x:Type TextBox}">
                     <Border x:Name="border" BorderBrush="DarkCyan" BorderThickness="1.5" Background="{TemplateBinding Background}" SnapsToDevicePixels="True" CornerRadius="3">
                         <ScrollViewer x:Name="PART_ContentHost" Focusable="false" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden"/>
                     </Border>
                     <ControlTemplate.Triggers>
                         <Trigger Property="IsEnabled" Value="false">
                         </Trigger>
                         <Trigger Property="IsMouseOver" Value="true">
                             <Setter Property="BorderBrush" TargetName="border" Value="DarkCyan"/>
                             <Setter Property="Opacity" TargetName="border" Value="0.5"/>
                         </Trigger>
                         <Trigger Property="IsFocused" Value="true">
                             <Setter Property="BorderBrush" TargetName="border" Value="DarkCyan"/>
                             <Setter Property="Opacity" TargetName="border" Value="0.5"/>
                         </Trigger>
                     </ControlTemplate.Triggers>
                 </ControlTemplate>
             </Setter.Value>
         </Setter>
     </Style>

<!-- ******************** End TextBoxes Style  ********************--> 

<!-- ********************* RadioButtons Style  ********************-->

     <Style x:Key="RadioTemplate" TargetType="RadioButton" >
         <Setter Property="Template">
             <Setter.Value>
                 <ControlTemplate TargetType="{x:Type RadioButton}">
                     <BulletDecorator Background="White" Cursor="Hand">
                         <BulletDecorator.Bullet>
                             <Grid Height="16" Width="16"> <!--Define size of the Bullet-->
                                 <Border 
                                     Name="RadioOuter" 
                                     Background="Transparent" 
                                     BorderBrush="DarkCyan" 
                                     BorderThickness="1.5" 
                                     CornerRadius="2">
                                     <Border.Effect>
                                         <DropShadowEffect BlurRadius="1" ShadowDepth="0.5" />
                                     </Border.Effect>
                                 </Border>
                                 <Border CornerRadius="0" Margin="4" Name="RadioMark" Background="DarkCyan" Visibility="Hidden" />
                             </Grid>
                         </BulletDecorator.Bullet>
                         <TextBlock Margin="5,1,0,0" Foreground="Black" FontSize="15">
                             <ContentPresenter />
                             </TextBlock>
                     </BulletDecorator>
                     <ControlTemplate.Triggers>
                         <Trigger Property="IsChecked" Value="true">
                             <Setter TargetName="RadioMark" Property="Visibility" Value="Visible"/>
                             <Setter TargetName="RadioOuter" Property="BorderBrush" Value="DarkCyan" />
                         </Trigger>
                         <Trigger Property="IsMouseOver" Value="true">
                             <Setter TargetName="RadioOuter" Property="Background" Value="Cyan" />
                             <Setter TargetName="RadioOuter" Property="BorderBrush" Value="{DynamicResource ApplicationAccentBrush}" />
                         </Trigger>
                     </ControlTemplate.Triggers>
                 </ControlTemplate>
             </Setter.Value>
         </Setter>
     </Style>

 <!-- ******************** End RadioButtons Style  ********************-->

 <!-- ************************ CheckBox Style  ************************-->

     <Style x:Key="CheckBoxFocusVisual">
         <Setter Property="Control.Template">
             <Setter.Value>
                 <ControlTemplate>
                     <Border>
                         <Rectangle Margin="15,0,0,0" StrokeThickness="1" Stroke="#60000000" StrokeDashArray="1 2"/>
                     </Border>
                 </ControlTemplate>
             </Setter.Value>
         </Setter>
     </Style>
     <Style x:Key="CheckBoxTemplate" TargetType="CheckBox">
         <Setter Property="SnapsToDevicePixels" Value="true"/>
         <Setter Property="OverridesDefaultStyle" Value="true"/>
         <Setter Property="FontFamily" Value="{DynamicResource MetroFontRegular}"/>
         <Setter Property="FocusVisualStyle" Value="{StaticResource CheckBoxFocusVisual}"/>
         <Setter Property="Foreground" Value="Black"/>
         <Setter Property="Background" Value="#3f3f3f"/>
         <Setter Property="FontSize" Value="12"/>
         <Setter Property="Margin" Value="0,5,0,0"/>
         <Setter Property="Template">
             <Setter.Value>
                 <ControlTemplate TargetType="CheckBox">
                     <BulletDecorator Background="Transparent">
                         <BulletDecorator.Bullet>
                             <Border x:Name="Border"  
                                 Width="16" 
                                 Height="16" 
                                 CornerRadius="1" 
                                 BorderBrush="DarkCyan"
                                 BorderThickness="1.5">
                                 <Border.Effect>
                                     <DropShadowEffect BlurRadius="1" ShadowDepth="0.5" />
                                 </Border.Effect>
                                     <Path 
                                        Width="7" Height="7" 
                                        x:Name="CheckMark"
                                        SnapsToDevicePixels="False" 
                                        Stroke="DarkCyan"
                                        StrokeThickness="1.5"
                                        Stretch="Fill"
                                        StrokeEndLineCap="Round"
                                        StrokeStartLineCap="Round"
                                        Data="M 0 0 L 7 7 M 0 7 L 7 0" />
                             </Border>
                         </BulletDecorator.Bullet>
                         <TextBlock Margin="5,0,0,0" Foreground="Black" FontSize="15">
                             <ContentPresenter />
                         </TextBlock>
                     </BulletDecorator>
                         <ControlTemplate.Triggers>
                             <Trigger Property="IsChecked" Value="false">
                                 <Setter TargetName="CheckMark" Property="Visibility" Value="Collapsed"/>
                             </Trigger>
                             <Trigger Property="IsChecked" Value="{x:Null}">
                                 <Setter TargetName="CheckMark" Property="Data" Value="M 0 7 L 7 0" />
                             </Trigger>
                             <Trigger Property="IsMouseOver" Value="true">
                                 <Setter TargetName="Border" Property="Background" Value="Cyan" />
                                 <Setter TargetName="Border" Property="BorderBrush" Value="{DynamicResource ApplicationAccentBrush}" />
                             </Trigger>
                             <Trigger Property="IsEnabled" Value="false">
                                 <Setter Property="Foreground" Value="#c1c1c1"/>
                             </Trigger>
                         </ControlTemplate.Triggers>
                 </ControlTemplate>
             </Setter.Value>
         </Setter>
     </Style>

 <!-- ******************** End ChekBox Style  ********************-->

 </Window.Resources>    

 <!-- ******************** End Style Section  ********************-->

 <!-- ******************** First Grid  ********************-->

 <Grid  >
     <Grid.RowDefinitions>
         <RowDefinition />
     </Grid.RowDefinitions>
     <Grid.ColumnDefinitions>
         <ColumnDefinition Width="0.15*"/>
         <ColumnDefinition />
     </Grid.ColumnDefinitions>

 <!-- ******************** Grid Splitter ********************-->

     <GridSplitter Grid.Row="1" Grid.Column="0" ShowsPreview="True" Width="2"  VerticalAlignment="Stretch">
         <GridSplitter.Template>
             <ControlTemplate TargetType="{x:Type GridSplitter}">
                 <Grid>
                     <Button Content="⁞" />
                     <Rectangle Fill="DarkCyan" />
                 </Grid>
             </ControlTemplate>
         </GridSplitter.Template>
     </GridSplitter>

 <!-- ******************** End of Grid Splitter ********************-->

 <!-- ********************** Left Pane ********************-->

     <Grid Grid.Row="0" Grid.Column="0">
         <Grid.RowDefinitions>
             <RowDefinition />
             <RowDefinition />
             <RowDefinition />
             <RowDefinition />
             <RowDefinition />
             <RowDefinition />
             <RowDefinition />
             <RowDefinition />
             <RowDefinition Height="0.6*"/>
         </Grid.RowDefinitions>
         <Button x:Name = "Main_Btn" Grid.Row="0" FontSize="20" Background="SteelBlue" Foreground="White" Style="{StaticResource ButtonTemplate}" Margin ="0 0 5 5" >
            <Button.Content>
                <Viewbox StretchDirection="DownOnly" Stretch="Uniform">
                     <ContentControl Content="Home"/>
               </Viewbox>
            </Button.Content>
         </Button>
         <Button x:Name = "Users_Btn" Grid.Row="1" FontSize="20" Background="SteelBlue" Foreground="White" Style="{StaticResource ButtonTemplate}" Margin ="0 0 5 5">
             <Button.Content>
                 <Viewbox StretchDirection="DownOnly" Stretch="Uniform">
                     <ContentControl Content="AAD Users"/>
                 </Viewbox>
             </Button.Content>
         </Button>
         <Button x:Name = "Devices_Btn" Grid.Row="2" FontSize="20" Background="SteelBlue" Foreground="White" Style="{StaticResource ButtonTemplate}" Margin ="0 0 5 5">
             <Button.Content>
                 <Viewbox StretchDirection="DownOnly" Stretch="Uniform">
                     <ContentControl Content="AAD Devices"/>
                 </Viewbox>
             </Button.Content>
         </Button>
         <Grid  Grid.Row="8"  Margin ="-5 0 0 0">
             <Grid.ColumnDefinitions>
                 <ColumnDefinition />
                 <ColumnDefinition />
                 <ColumnDefinition />
             </Grid.ColumnDefinitions>
             <Button x:Name = "Min_btn" Grid.Column="0" FontWeight="Bold" Style="{StaticResource ButtonTemplate}" FontSize="20" Background="SteelBlue" Foreground="White"  Margin ="5 5 5 5">
                 <Button.Content>
                     <Viewbox StretchDirection="DownOnly" Stretch="Uniform">
                         <Line X1="3" Y1="30" X2="28" Y2="30" Stroke="White" StrokeThickness="4" Margin ="0 0 0 5"/>
                     </Viewbox>
                 </Button.Content>
             </Button>
             <Button x:Name = "Max_btn" Grid.Column="1" Background="SteelBlue" Style="{StaticResource ButtonTemplate}" Margin ="5 5 5 5">
                 <Button.Content>
                     <Viewbox StretchDirection="DownOnly" Stretch="Uniform">
                         <Rectangle Fill="SteelBlue" HorizontalAlignment="Left"
                            Height="80" Margin="10,10,10,10"
                            Stroke="White" StrokeThickness="10"
                            VerticalAlignment="Top" Width="100"
                            RadiusY="13.5" RadiusX="13.5" /> 
                     </Viewbox>
                 </Button.Content>
             </Button>
             <Button x:Name = "Exit_btn" Grid.Column="2" FontWeight="Bold" FontSize="25" Background="SteelBlue" Style="{StaticResource ButtonTemplate}" Foreground="White"  Margin ="5 5 5 5">
                 <Button.Content>
                     <Viewbox StretchDirection="DownOnly" Stretch="Uniform">
                         <ContentControl Content="X"/>
                     </Viewbox>
                 </Button.Content>
             </Button>
         </Grid>
     </Grid>

 <!-- ********************** End of Left Pane ********************-->

 <!-- ********************** Right Pane ********************-->

 <!-- ********************** _Main grid  ShowGridLines="True"********************-->
    
     <Grid x:Name = "Main_grd" Grid.Row="0" Grid.Column="1" >
         <Grid.RowDefinitions>
             <RowDefinition Height="0.15*"/>
             <RowDefinition Height="0.02*"/>
             <RowDefinition Height="0.1*"/>
             <RowDefinition Height="0.02*"/>
             <RowDefinition />
             <RowDefinition Height="0.1*"/>
             <RowDefinition Height="0.05*"/>
         </Grid.RowDefinitions>
         <Grid.ColumnDefinitions>
             <ColumnDefinition Width="0.02*"/>
             <ColumnDefinition />
             <ColumnDefinition Width="0.1*"/>
             <ColumnDefinition Width="0.1*"/>
             <ColumnDefinition Width="0.02*"/>
         </Grid.ColumnDefinitions>

 <!-- ********************** _Main grid Header  ********************-->

         <Grid Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="5">
             <Rectangle Fill="SteelBlue" Margin ="5,0,0,0" RadiusY="3" RadiusX="3"/> 
             <Viewbox StretchDirection="DownOnly" Stretch="Uniform" HorizontalAlignment="Right">             
                  <TextBlock Text="Azure AD Objects" Margin="0,0,35,0" Foreground="White" FontSize="30" FontFamily="Segoe UI" />
             </Viewbox>
         </Grid>

 <!-- ********************** _Main grid End of Header  ********************-->     

 <!-- ********************** _Main grid Login part  ********************-->

         <Viewbox StretchDirection="Both" Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2" Stretch="Uniform" HorizontalAlignment="Right" >
             <TextBox x:Name = "Login_tb" 
                 Style="{StaticResource TextBoxTemplate}" 
                 Padding="2,2,2,2" 
                 Margin="5,5,5,5" 
                 FontSize="15" 
                 HorizontalAlignment="Left" 
                 MinWidth="400" 
                 MinHeight="15" 
                 FontFamily="Segoe UI" />
         </Viewbox>
         <Button x:Name = "Login_Btn" Grid.Row="2" Grid.Column="3" FontSize="15" Background="White" Foreground="Black" Style="{StaticResource ButtonTemplate}" Margin ="5,5,5,5" >
             <Button.Content>
                 <Viewbox StretchDirection="DownOnly" Stretch="Uniform">
                     <ContentControl Content="Login"/>
                 </Viewbox>
             </Button.Content>
         </Button>

 <!-- ********************** _Main grid End of Login part ********************-->

 <!-- ********************** _Main grid Details ShowGridLines="True"********************-->

 <Grid x:Name="Dash_grd" Grid.Row="4" Grid.Column="1" Grid.ColumnSpan="3"  >
     <Grid.RowDefinitions>
         <RowDefinition />
         <RowDefinition Height="0.3*"/>
         <RowDefinition />
     </Grid.RowDefinitions>
     <Grid.ColumnDefinitions>
         <ColumnDefinition />
         <ColumnDefinition />
         <ColumnDefinition />
     </Grid.ColumnDefinitions>

     
     <Grid x:Name="Overview_grd" Grid.Row="0" Grid.Column="0" >
         <Grid.RowDefinitions>
             <RowDefinition />
             <RowDefinition />
             <RowDefinition />
             <RowDefinition />
             <RowDefinition />
          </Grid.RowDefinitions>
         <Grid.ColumnDefinitions>
             <ColumnDefinition Width="0.35*"/>
             <ColumnDefinition />
         </Grid.ColumnDefinitions>  
        
         <Label Grid.Row="0" Grid.Column="0" Content="Overview" FontSize="17" HorizontalAlignment="Left"  VerticalAlignment="Center" />
         <Label x:Name="Company_lb" Grid.Row="1" Grid.Column="0" FontSize="20" HorizontalAlignment="Left" VerticalAlignment="Center" FontWeight="Bold"/>
         <Label Grid.Row="2" Grid.Column="0" Content="Tennant:" FontSize="15" HorizontalAlignment="Left"  VerticalAlignment="Center"/>
         <Label Grid.Row="3" Grid.Column="0" Content="Tennant ID:" FontSize="15" HorizontalAlignment="Left"  VerticalAlignment="Center"/>
         <Label Grid.Row="4" Grid.Column="0" Content="Capabilities:" FontSize="15" HorizontalAlignment="Left"  VerticalAlignment="Center"/>
         <Label x:Name="Tennant_lb" Grid.Row="2" Grid.Column="1"  FontSize="15" HorizontalAlignment="Left"  VerticalAlignment="Center"/>
         <Label x:Name="Tennantid_lb" Grid.Row="3" Grid.Column="1"  FontSize="15" HorizontalAlignment="Left"  VerticalAlignment="Center"/>
         <Label x:Name="Capabilities_lb" Grid.Row="4" Grid.Column="1"  FontSize="15" HorizontalAlignment="Left"  VerticalAlignment="Center"/>

     </Grid>
     <Label  Grid.Row="0" Grid.Column="1" Content="Subscribed SKUs" FontSize="17" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="15,3,0,0"/>
     <Grid Grid.Row="0" Grid.Column="1" ShowGridLines="True" Margin="15,45,5,0">
     
        <ListView x:Name="SKU_lv" Grid.Column="0" Margin="0,0,5,0" FontSize="15" Foreground="Black"  ItemsSource="{Binding}" IsSynchronizedWithCurrentItem="True" Panel.ZIndex="1" BorderThickness="0">
            <ListView.Resources>
                <Style TargetType="{x:Type GridViewColumnHeader}">
                     <Setter Property="FontSize" Value="15"/>
                </Style>
  
            </ListView.Resources>

      

             <ListView.View>
                     <GridView>
                         <GridView.Columns>
                             <GridViewColumn DisplayMemberBinding="{Binding SKU}" Header="SKU"  Width="240"/>
                             <GridViewColumn DisplayMemberBinding="{Binding Enabled}" Header="Enabled" />
                          </GridView.Columns>
                     </GridView>
             </ListView.View>
         </ListView>       
     </Grid>

     <Image Grid.Row="2" Grid.Column="0" x:Name="Chart_usr" /> 
     <Image Grid.Row="2" Grid.Column="1" x:Name="Chart_grp" />
     <Image Grid.Row="2" Grid.Column="2" x:Name="Chart_dev" />
     








 </Grid>

 <!-- ********************** _Main grid End of Details ********************-->

 <!-- *************************** _Main grid Footer  *************************-->

    <Grid Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="5" Margin="0,0,-5,-5">
         <Rectangle Fill="LightGray" />  
         <Viewbox StretchDirection="DownOnly" Stretch="Uniform">         
             <TextBlock x:Name="Footer_tb" Margin="3,3,0,0" Foreground="White" FontSize="15" FontFamily="Segoe UI"/>
         </Viewbox>
    </Grid>   

 <!-- ********************** _Main grid End of Footer  ********************-->

 </Grid>

 <!-- ********************** _Main grid End of Main grid  ********************-->

 <!-- ************************* _Users grid  ***********************-->

 <Grid x:Name = "Users_grd" Grid.Row="0" Grid.Column="1" >
     <Grid.RowDefinitions>
         <RowDefinition Height="0.15*"/>
         <RowDefinition Height="0.02*"/>
         <RowDefinition Height="0.1*"/>
         <RowDefinition Height="0.02*"/>
         <RowDefinition />
         <RowDefinition Height="0.1*"/>
         <RowDefinition Height="0.05*"/>
     </Grid.RowDefinitions>
     <Grid.ColumnDefinitions>
         <ColumnDefinition Width="0.02*"/>
         <ColumnDefinition />
         <ColumnDefinition Width="0.2*"/>
         <ColumnDefinition Width="0.2*"/>
         <ColumnDefinition Width="0.02*"/>
     </Grid.ColumnDefinitions>


 <!-- ********************** _Users Header  ********************-->

 <Grid Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="5">
     <Rectangle Fill="SteelBlue" Margin ="5,0,0,0" RadiusY="3" RadiusX="3"/> 
     <Viewbox StretchDirection="DownOnly" Stretch="Uniform" HorizontalAlignment="Right">             
         <TextBlock Text="Azure AD Users" Margin="0,0,35,0" Foreground="White" FontSize="30" FontFamily="Segoe UI" />
     </Viewbox>
     </Grid>

 <!-- ********************** _Users End of Header  ********************-->

     <Button x:Name = "Add_grp_Btn" Grid.Row="2" Grid.Column="2" FontSize="15" Background="White" Foreground="Black" Style="{StaticResource ButtonTemplate}" Margin ="5,5,5,5" >
     <Button.Content>
         <Viewbox StretchDirection="DownOnly" Stretch="Uniform">
             <ContentControl Content="Create Group"/>
         </Viewbox>
     </Button.Content>
     </Button>

     <Button x:Name = "Del_grp_Btn" Grid.Row="2" Grid.Column="3" FontSize="15" Background="White" Foreground="Black" Style="{StaticResource ButtonTemplate}" Margin ="5,5,5,5" >
     <Button.Content>
         <Viewbox StretchDirection="DownOnly" Stretch="Uniform">
             <ContentControl Content="Delete Group(s)"/>
         </Viewbox>
     </Button.Content>
     </Button>

 <!-- ********************** _Users Tables  ********************-->

 <Grid Grid.Row="4" Grid.Column="1" Grid.ColumnSpan="4" >
     <Grid.ColumnDefinitions>
         <ColumnDefinition/>
        <ColumnDefinition />
    </Grid.ColumnDefinitions> 

     <GridSplitter Grid.Column="0" ShowsPreview="True" Width="2"  VerticalAlignment="Stretch">
         <GridSplitter.Template>
             <ControlTemplate TargetType="{x:Type GridSplitter}">
                 <Grid>
                     <Button Content="⁞" />
                     <Rectangle Fill="DarkCyan" />
                 </Grid>
             </ControlTemplate>
         </GridSplitter.Template>
     </GridSplitter>

 <!-- -->

     <ListView x:Name="Users_lv" Grid.Column="0" Margin="0,0,5,0" FontSize="12" ItemsSource="{Binding}" IsSynchronizedWithCurrentItem="True" BorderThickness="0">
         <ListView.Resources>
             <Style TargetType="{x:Type GridViewColumnHeader}">
                 <Setter Property="HorizontalContentAlignment" Value="Left"/>
                 <Setter Property="Background" Value="SteelBlue"/>
                 <Setter Property="Foreground" Value="White"/>
                 <Setter Property="Padding" Value="5,5,5,5"/>
                 <Setter Property="FontSize" Value="14"/>
             </Style>
         </ListView.Resources>

         <ListView.View>
             <GridView AllowsColumnReorder="true">
                 <GridView.ColumnHeaderContextMenu>
                     <ContextMenu >
                         <MenuItem x:Name="Users_lv_Asc_sort" Header="Sore - Ascending" />
                         <MenuItem x:Name="Users_lv_Desc_sort" Header="Sort - Descending" />
                         <Separator />
                         <MenuItem x:Name="Users_lv_Refresh" Header="Refresh" />                                     
                     </ContextMenu>
                 </GridView.ColumnHeaderContextMenu>                 
                     <GridView.Columns>
                         <GridViewColumn Width="30">
                             <GridViewColumn.Header>
                                 <CheckBox x:Name="Users_lv_SelectAll" HorizontalAlignment="Center" VerticalAlignment="Center" IsThreeState="False" />
                             </GridViewColumn.Header>
                             <GridViewColumn.CellTemplate>
                                 <DataTemplate>
                                     <CheckBox Tag="{Binding ID}" IsChecked="{Binding RelativeSource={RelativeSource AncestorType={x:Type ListViewItem}}, Path=IsSelected}" Style="{StaticResource CheckBoxTemplate}"/>  
                                 </DataTemplate>
                             </GridViewColumn.CellTemplate>
                         </GridViewColumn>
                         <GridViewColumn DisplayMemberBinding="{Binding UserName}" Header="User Name" Width="170"/>
                         <GridViewColumn DisplayMemberBinding="{Binding UPN}" Header="UPN" Width="250"/>
                     </GridView.Columns>
                 </GridView>
         </ListView.View>
     </ListView>

 <!-- -->

     <ListView x:Name="Groups_lv" Grid.Column="1"  FontSize="12" ItemsSource="{Binding}" IsSynchronizedWithCurrentItem="True" ScrollViewer.CanContentScroll="True" BorderThickness="0" >
         <ListView.Resources>
                 <Style TargetType="{x:Type GridViewColumnHeader}">
                 <Setter Property="HorizontalContentAlignment" Value="Left"/>
                 <Setter Property="Background" Value="SteelBlue"/>
                 <Setter Property="Foreground" Value="White"/>
                 <Setter Property="Padding" Value="5,5,5,5"/>
                 <Setter Property="FontSize" Value="14"/>
             </Style>
         </ListView.Resources>

         <ListView.View>
             <GridView AllowsColumnReorder="true" ColumnHeaderToolTip="Authors">

                 <GridView.ColumnHeaderContextMenu>
                     <ContextMenu >
                         <MenuItem x:Name="Groups_lv_Asc_sort" Header="Sort - Ascending" />
                         <MenuItem x:Name="Groups_lv_Desc_sort" Header="Sort - Descending" />
                         <Separator />
                         <MenuItem x:Name="Groups_lv_Refresh" Header="Refresh" />
                     </ContextMenu>
                 </GridView.ColumnHeaderContextMenu>

                 <GridView.Columns>
                     <GridViewColumn Width="30">
                         <GridViewColumn.Header>
                             <CheckBox x:Name="Groups_lv_SelectAll" HorizontalAlignment="Center" VerticalAlignment="Center" IsThreeState="False" />
                         </GridViewColumn.Header>                     
                         <GridViewColumn.CellTemplate>
                             <DataTemplate>
                                 <CheckBox Tag="{Binding ID}" IsChecked="{Binding RelativeSource={RelativeSource AncestorType={x:Type ListViewItem}}, Path=IsSelected}" Style="{StaticResource CheckBoxTemplate}"/>  
                             </DataTemplate>
                         </GridViewColumn.CellTemplate>
                     </GridViewColumn>
                         <GridViewColumn DisplayMemberBinding="{Binding GroupName}" Header="Group Name" Width="170"/>
                         <GridViewColumn DisplayMemberBinding="{Binding isOnPrem}"  Header="isOnPrem" Width="70"/>
                         <GridViewColumn Header="Mail" Width="250">
                         <GridViewColumn.CellTemplate>
                             <DataTemplate>
                                 <TextBlock Text="{Binding Mail}" TextDecorations="Underline" Foreground="Blue" Cursor="Hand" />
                             </DataTemplate>
                         </GridViewColumn.CellTemplate>                     
                     </GridViewColumn>
                 </GridView.Columns>
             </GridView>
         </ListView.View>
     </ListView>

 </Grid>

 <!-- ********************** _Users End of Tables  ********************-->

 <!-- ********************** _Users Add Button and Text  ********************-->

     <Viewbox StretchDirection="DownOnly" Stretch="Uniform" Grid.Row="5" Grid.Column="1">    
         <Label x:Name="Notif_lb" FontFamily="Segoe UI" FontSize="20" FontStyle="Italic" Foreground="LightGray">
             Select users and groups and click &quot;Add&quot; or &quot;Remove&quot; button 
         </Label>
     </Viewbox>
     
     <Button x:Name = "Add_Btn" Grid.Row="5" Grid.Column="2" FontSize="15" Background="White" Foreground="Black" Style="{StaticResource ButtonTemplate}" Margin ="5,5,5,5" >
             <Button.Content>
                 <Viewbox StretchDirection="DownOnly" Stretch="Uniform">
                     <ContentControl Content="Add"/>
                 </Viewbox>
             </Button.Content>
     </Button>

     <Button x:Name = "Remove_Btn" Grid.Row="5" Grid.Column="3" FontSize="15" Background="White" Foreground="Black" Style="{StaticResource ButtonTemplate}" Margin ="5,5,5,5" >
             <Button.Content>
                 <Viewbox StretchDirection="DownOnly" Stretch="Uniform">
                     <ContentControl Content="Remove"/>
                 </Viewbox>
             </Button.Content>
     </Button>

 <!-- ********************** _Users End of Add Button and Text  ********************-->

 <!-- ****************************** _Users Footer  *********************************-->

    <Grid Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="5" Margin="0,0,-5,-5">
         <Rectangle Fill="LightGray" />  
         <Viewbox StretchDirection="DownOnly" Stretch="Uniform">         
             <TextBlock x:Name="FooterU_tb" Margin="3,3,0,0" Foreground="White" FontSize="15" FontFamily="Segoe UI"/>
         </Viewbox>
    </Grid>   

 <!-- **************************** _Users End of Footer  *****************************-->

 </Grid>

 <!-- ************************* _Users End of Users grid  ***********************-->

 <!-- ************************* _Devices grid  ***********************-->
  
 <Grid x:Name = "Devices_grd" Grid.Row="0" Grid.Column="1">
     <Grid.RowDefinitions>
         <RowDefinition Height="0.15*"/>
         <RowDefinition Height="0.02*"/>
         <RowDefinition Height="0.1*"/>
         <RowDefinition Height="0.02*"/>
         <RowDefinition />
         <RowDefinition Height="0.1*"/>
         <RowDefinition Height="0.05*"/>
     </Grid.RowDefinitions>
     <Grid.ColumnDefinitions>
         <ColumnDefinition Width="0.02*"/>
         <ColumnDefinition />
         <ColumnDefinition Width="0.1*"/>
         <ColumnDefinition Width="0.1*"/>
         <ColumnDefinition Width="0.02*"/>
     </Grid.ColumnDefinitions>


 <!-- ********************** _Devices Header  ********************-->

 <Grid Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="5">
     <Rectangle Fill="SteelBlue" Margin ="5,0,0,0" RadiusY="3" RadiusX="3"/> 
     <Viewbox StretchDirection="DownOnly" Stretch="Uniform" HorizontalAlignment="Right">             
         <TextBlock Text="Azure AD Devices" Margin="0,0,35,0" Foreground="White" FontSize="30" FontFamily="Segoe UI" />
     </Viewbox>
     </Grid>

 <!-- ********************** _Devices End of Header  ********************-->

 <!-- ********************** _Devices Tables  ********************-->

 <Grid Grid.Row="4" Grid.Column="1" Grid.ColumnSpan="4" >
     <Grid.ColumnDefinitions>
         <ColumnDefinition/>
        <ColumnDefinition />
    </Grid.ColumnDefinitions> 

     <GridSplitter Grid.Column="0" ShowsPreview="True" Width="2"  VerticalAlignment="Stretch">
         <GridSplitter.Template>
             <ControlTemplate TargetType="{x:Type GridSplitter}">
                 <Grid>
                     <Button Content="⁞" />
                     <Rectangle Fill="DarkCyan" />
                 </Grid>
             </ControlTemplate>
         </GridSplitter.Template>
     </GridSplitter>

 <!-- -->

     <ListView x:Name="Devices_lv" Grid.Column="0" Margin="0,0,5,0" FontSize="12" ItemsSource="{Binding}" IsSynchronizedWithCurrentItem="True" BorderThickness="0">
         <ListView.Resources>
             <ContextMenu x:Key="ItemContextMenu_Devices_lv">
                 <MenuItem Header="Properties"></MenuItem>
             </ContextMenu>         
             <Style TargetType="{x:Type GridViewColumnHeader}">
                 <Setter Property="HorizontalContentAlignment" Value="Left"/>
                 <Setter Property="Background" Value="SteelBlue"/>
                 <Setter Property="Foreground" Value="White"/>
                 <Setter Property="Padding" Value="5,5,5,5"/>
                 <Setter Property="FontSize" Value="14"/>
              </Style>
         </ListView.Resources>
         <ListView.ItemContainerStyle>
                 <Style TargetType="{x:Type ListViewItem}">
                 <Setter Property="ContextMenu" Value="{StaticResource ItemContextMenu_Devices_lv}"/>
             </Style>
         </ListView.ItemContainerStyle>
         <ListView.View>
                 <GridView>
                    <GridView.ColumnHeaderContextMenu>
                        <ContextMenu >
                            <MenuItem x:Name="Devices_lv_Asc_sort" Header="Sort - Ascending" />
                            <MenuItem x:Name="Devices_lv_Desc_sort" Header="Sort - Descending" />
                            <Separator />
                            <MenuItem x:Name="Devices_lv_Refresh" Header="Refresh" />                                        
                         </ContextMenu>
                    </GridView.ColumnHeaderContextMenu>                 
                     <GridView.Columns>
                         <GridViewColumn Width="30">
                             <GridViewColumn.Header>
                                 <CheckBox x:Name="Devices_lv_SelectAll" HorizontalAlignment="Center" VerticalAlignment="Center" IsThreeState="False" />
                             </GridViewColumn.Header>   
                             <GridViewColumn.CellTemplate>
                                 <DataTemplate>
                                     <CheckBox Tag="{Binding ID}" IsChecked="{Binding RelativeSource={RelativeSource AncestorType={x:Type ListViewItem}}, Path=IsSelected}" Style="{StaticResource CheckBoxTemplate}"/>  
                                 </DataTemplate>
                             </GridViewColumn.CellTemplate>
                         </GridViewColumn>
                         <GridViewColumn DisplayMemberBinding="{Binding DeviceName}" Header="Device Name" Width="110"/>
                         <GridViewColumn DisplayMemberBinding="{Binding isMDM}" Header="IsMDM" />
                         <GridViewColumn DisplayMemberBinding="{Binding OS}" Header="OS" Width="70"/>
                         <GridViewColumn DisplayMemberBinding="{Binding Version}" Header="Version" Width="100" />
                         <GridViewColumn DisplayMemberBinding="{Binding LastSign}" Header="Lastsign" Width="130"/>
                     </GridView.Columns>
                 </GridView>
         </ListView.View>
     </ListView>

 <!-- -->

     <ListView x:Name="Groups_lv_d" Grid.Column="1"  FontSize="12"  ItemsSource="{Binding}" IsSynchronizedWithCurrentItem="True" ScrollViewer.CanContentScroll="True"  HorizontalAlignment="Stretch" VerticalAlignment="Top" BorderThickness="0" >
         <ListView.Resources>
             <ContextMenu x:Key="ItemContextMenu_Group_lv_d">
                 <MenuItem Header="Properties"></MenuItem>
             </ContextMenu>
             <Style TargetType="{x:Type GridViewColumnHeader}">
                <!-- <Setter Property="HorizontalContentAlignment" Value="Left"/> -->
                 <Setter Property="Background" Value="SteelBlue"/>
                 <Setter Property="Foreground" Value="White"/>
                 <Setter Property="Padding" Value="5,5,5,5"/>
                 <Setter Property="FontSize" Value="14"/>
                 <Setter Property="HorizontalContentAlignment" Value="Stretch" />
                 <Setter Property="HorizontalContentAlignment" Value="Stretch" />
             </Style>
          </ListView.Resources>
          <ListView.ItemContainerStyle>
             <Style TargetType="{x:Type ListViewItem}">
                 <Setter Property="ContextMenu" Value="{StaticResource ItemContextMenu_Group_lv_d}"/>
              </Style>
         </ListView.ItemContainerStyle>
         <ListView.View>
              <GridView AllowsColumnReorder="true" ColumnHeaderToolTip="Authors">
                 <GridView.ColumnHeaderContextMenu>
                     <ContextMenu >
                         <MenuItem x:Name="Groups_lv_d_Asc_sort" Header="Sort - Ascending" />
                         <MenuItem x:Name="Groups_lv_d_Desc_sort" Header="Sort - Descending" />
                         <Separator />
                         <MenuItem x:Name="Groups_lv_d_Refresh" Header="Refresh" />                                 
                       </ContextMenu>
                 </GridView.ColumnHeaderContextMenu>
                 
                 <GridView.Columns >
                     <GridViewColumn Width="30">
                         <GridViewColumn.Header>
                             <CheckBox x:Name="Groups_lv_d_SelectAll" HorizontalAlignment="Center" VerticalAlignment="Center" IsThreeState="False" />
                         </GridViewColumn.Header>   
                         <GridViewColumn.CellTemplate>
                             <DataTemplate>
                                 <CheckBox Tag="{Binding ID}" IsChecked="{Binding RelativeSource={RelativeSource AncestorType={x:Type ListViewItem}}, Path=IsSelected}" Style="{StaticResource CheckBoxTemplate}"/>  
                             </DataTemplate>
                         </GridViewColumn.CellTemplate>
                     </GridViewColumn>
                         <GridViewColumn DisplayMemberBinding="{Binding GroupName}" Header="Group Name" Width="170"/>
                         <GridViewColumn DisplayMemberBinding="{Binding isOnPrem}"  Header="isOnPrem" Width="70"/>
                         <GridViewColumn Header="Mail" Width="250">
                         <GridViewColumn.CellTemplate>
                             <DataTemplate>
                                 <TextBlock Text="{Binding Mail}" TextDecorations="Underline" Foreground="Blue" Cursor="Hand" />
                             </DataTemplate>
                         </GridViewColumn.CellTemplate>                     
                     </GridViewColumn>
                 </GridView.Columns>
             </GridView>
         </ListView.View>
     </ListView>

 </Grid>

 <!-- ********************** _Devices End of Tables  ********************-->

 <!-- ********************** _Devices Add Button and Text  ********************-->

 <Viewbox StretchDirection="DownOnly" Stretch="Uniform" Grid.Row="5" Grid.Column="1">    
     <Label x:Name="Notif_lb_d" FontFamily="Segoe UI" FontSize="20" FontStyle="Italic" Foreground="LightGray">
         Select devices and groups and click &quot;Add&quot; or &quot;Remove&quot; button 
     </Label>
 </Viewbox>
 
 <Button x:Name = "Add_Btn_d" Grid.Row="5" Grid.Column="2" FontSize="15" Background="White" Foreground="Black" Style="{StaticResource ButtonTemplate}" Margin ="5,5,5,5" >
         <Button.Content>
             <Viewbox StretchDirection="DownOnly" Stretch="Uniform">
                 <ContentControl Content="Add"/>
             </Viewbox>
         </Button.Content>
 </Button>

 <Button x:Name = "Remove_Btn_d" Grid.Row="5" Grid.Column="3" FontSize="15" Background="White" Foreground="Black" Style="{StaticResource ButtonTemplate}" Margin ="5,5,5,5" >
         <Button.Content>
             <Viewbox StretchDirection="DownOnly" Stretch="Uniform">
                 <ContentControl Content="Remove"/>
             </Viewbox>
         </Button.Content>
 </Button>

 <!-- ********************** _Devices End of Add Button and Text  ********************-->

 <!-- ********************** _Devices Footer  ********************-->

 <Grid Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="5" Margin="0,0,-5,-5">
     <Rectangle Fill="LightGray" />  
     <Viewbox StretchDirection="DownOnly" Stretch="Uniform">         
         <TextBlock x:Name="Footer_tb_d" Margin="3,3,0,0" Foreground="White" FontSize="15" FontFamily="Segoe UI"/>
     </Viewbox>
 </Grid>   

 <!-- ********************** _Devices End of Footer  ********************-->

 </Grid>

 <!-- ************************* _Devices End of Users grid  ***********************-->



 </Grid>
 </Window>
"@

# XAML Child window
[xml]$grp_xaml = @"
<Window x:Name="grp_Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Create New Group"
        Height="500" 
        Width="600"
        ResizeMode="CanResizeWithGrip" 
        ShowInTaskbar="False"
        WindowStartupLocation="CenterScreen"
        BorderBrush="DarkCyan"
        BorderThickness="2"
        WindowStyle="ToolWindow"
        FontFamily="Segoe UI"
        >
         <Window.Resources>

         <!-- ************************** Buttons Style  *********************** --> 

         <Style x:Key="ButtonTemplate1" TargetType="Button" >
             <Setter Property="OverridesDefaultStyle" Value="True" />
             <Setter Property="Cursor" Value="Hand" />
             <Setter Property="Template">
                 <Setter.Value>
                     <ControlTemplate TargetType="Button">
                         <Border Name="border" BorderThickness="2" BorderBrush="DarkCyan" Background="{TemplateBinding Background}" CornerRadius="3" >
                             <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" />
                         </Border>
                         <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                  <Setter Property="Opacity" Value="0.5" />
                             </Trigger>
                         </ControlTemplate.Triggers>
                     </ControlTemplate>
                 </Setter.Value>
             </Setter>
         </Style>
   
         <!-- ************************ End Buttons Style ********************--> 

         <!-- ************************ CheckBox Style  ************************-->

         <Style x:Key="CheckBoxFocusVisual">
             <Setter Property="Control.Template">
                 <Setter.Value>
                     <ControlTemplate>
                         <Border>
                             <Rectangle Margin="15,0,0,0" StrokeThickness="1" Stroke="#60000000" StrokeDashArray="1 2"/>
                         </Border>
                     </ControlTemplate>
                 </Setter.Value>
             </Setter>
         </Style>
         <Style x:Key="CheckBoxTemplate" TargetType="CheckBox">
             <Setter Property="SnapsToDevicePixels" Value="true"/>
             <Setter Property="OverridesDefaultStyle" Value="true"/>
             <Setter Property="FontFamily" Value="{DynamicResource MetroFontRegular}"/>
             <Setter Property="FocusVisualStyle" Value="{StaticResource CheckBoxFocusVisual}"/>
             <Setter Property="Foreground" Value="Black"/>
             <Setter Property="Background" Value="#3f3f3f"/>
             <Setter Property="FontSize" Value="12"/>
             <Setter Property="Margin" Value="0,5,0,0"/>
             <Setter Property="Template">
                 <Setter.Value>
                     <ControlTemplate TargetType="CheckBox">
                         <BulletDecorator Background="Transparent">
                             <BulletDecorator.Bullet>
                                 <Border x:Name="Border"  
                                     Width="16" 
                                     Height="16" 
                                     CornerRadius="1" 
                                     BorderBrush="DarkCyan"
                                     BorderThickness="1.5">
                                     <Border.Effect>
                                         <DropShadowEffect BlurRadius="1" ShadowDepth="0.5" />
                                     </Border.Effect>
                                         <Path 
                                            Width="7" Height="7" 
                                            x:Name="CheckMark"
                                            SnapsToDevicePixels="False" 
                                            Stroke="DarkCyan"
                                            StrokeThickness="1.5"
                                            Stretch="Fill"
                                            StrokeEndLineCap="Round"
                                            StrokeStartLineCap="Round"
                                            Data="M 0 0 L 7 7 M 0 7 L 7 0" />
                                 </Border>
                             </BulletDecorator.Bullet>
                             <TextBlock Margin="5,0,0,0" Foreground="Black" FontSize="15">
                                 <ContentPresenter />
                             </TextBlock>
                         </BulletDecorator>
                             <ControlTemplate.Triggers>
                                 <Trigger Property="IsChecked" Value="false">
                                     <Setter TargetName="CheckMark" Property="Visibility" Value="Collapsed"/>
                                 </Trigger>
                                 <Trigger Property="IsChecked" Value="{x:Null}">
                                     <Setter TargetName="CheckMark" Property="Data" Value="M 0 7 L 7 0" />
                                 </Trigger>
                                 <Trigger Property="IsMouseOver" Value="true">
                                     <Setter TargetName="Border" Property="Background" Value="Cyan" />
                                     <Setter TargetName="Border" Property="BorderBrush" Value="{DynamicResource ApplicationAccentBrush}" />
                                 </Trigger>
                                 <Trigger Property="IsEnabled" Value="false">
                                     <Setter Property="Foreground" Value="#c1c1c1"/>
                                 </Trigger>
                             </ControlTemplate.Triggers>
                     </ControlTemplate>
                 </Setter.Value>
             </Setter>
         </Style>
    
        <!-- ******************** End ChekBox Style  ********************-->  
        
 
        </Window.Resources>

    <Grid >
         <Grid.RowDefinitions>
             <RowDefinition Height="0.8*"/>
             <RowDefinition Height="0.7*"/>
             <RowDefinition Height="0.7*"/>
             <RowDefinition Height="0.7*"/>
             <RowDefinition />
             <RowDefinition />
         </Grid.RowDefinitions>
         <Grid.ColumnDefinitions>
             <ColumnDefinition Width="0.1*"/>
             <ColumnDefinition />
             <ColumnDefinition />
             <ColumnDefinition Width="0.1*"/>
         </Grid.ColumnDefinitions>

          <Grid Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="4" Margin="0,2,5,2">
             <Rectangle Fill="SteelBlue" Margin ="5,0,0,0" RadiusY="3" RadiusX="3"/> 
             <Viewbox StretchDirection="DownOnly" Stretch="Uniform" HorizontalAlignment="Right">             
                 <TextBlock Text="Create New Group" Margin="0,0,35,0" Foreground="White" FontSize="30" FontFamily="Segoe UI" />
             </Viewbox>
            </Grid>

        <CheckBox x:Name="G_Unified" Content="Unified (Office 365)" Grid.Row="1" Grid.Column="2" FontSize="13" Margin="10,5,5,5" Style="{StaticResource CheckBoxTemplate}" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <CheckBox x:Name="G_DynamicMembership" Content="DynamicMembership (Security)" Grid.Row="1" Grid.Column="2" FontSize="13" Margin="10,25,5,5" Style="{StaticResource CheckBoxTemplate}" HorizontalAlignment="Left" VerticalAlignment="Center"/>

        <CheckBox x:Name="G_mailEnabled" Content="Mail Enabled" Grid.Row="3" Grid.Column="2" FontSize="13" Margin="10,5,5,5" Style="{StaticResource CheckBoxTemplate}" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <CheckBox x:Name="G_securityEnabled" Content="Security Enabled" Grid.Row="3" Grid.Column="2" FontSize="13" Margin="10,25,5,5" Style="{StaticResource CheckBoxTemplate}" HorizontalAlignment="Left" VerticalAlignment="Center"/>
        
       <!-- ******************** Text box  ********************--> 

        <Viewbox Grid.Row="1" Grid.Column="1" HorizontalAlignment="Left" Stretch="Uniform">
          <Grid >
               <Grid.Resources>
                   <VisualBrush x:Key="LoginHint" Stretch="None" AlignmentX="left" AlignmentY="Top" >
                        <VisualBrush.Transform>
                            <TranslateTransform X="5" Y="3" />
                        </VisualBrush.Transform>
                        <VisualBrush.Visual>
                            <Grid HorizontalAlignment="Left">
                              <TextBlock FontFamily="SEGOEWP"  FontSize="10"   FontWeight="Normal" 
                                   HorizontalAlignment="Left" VerticalAlignment="Center" Foreground="Gray" FontStyle="Italic" Opacity="1"  
                                   Text="Enter Group Name"/>
                            </Grid>
                        </VisualBrush.Visual>
                    </VisualBrush>
                </Grid.Resources>
                <TextBox x:Name="G_displayName" MinWidth="150" Margin="2,2,2,2" >
                  <TextBox.Style>
                     <Style TargetType="{x:Type TextBox}">
                       <Setter Property="Background" Value="Transparent"/>
                         <Style.Triggers>
                           <DataTrigger Binding="{Binding ElementName=G_displayName,Path=Text}" Value="" >
                              <Setter Property="Background" Value="{StaticResource LoginHint}"/>
                           </DataTrigger>
                         </Style.Triggers>
                        </Style>
                    </TextBox.Style>
                 </TextBox>
             </Grid>
         </Viewbox>

       <!-- ******************** Text box  ********************--> 

        <Viewbox Grid.Row="2" Grid.Column="1" HorizontalAlignment="Left" Stretch="Uniform"  Grid.ColumnSpan="2">
          <Grid >
               <Grid.Resources>
                   <VisualBrush x:Key="LoginHint" Stretch="None" AlignmentX="left" AlignmentY="Top" >
                        <VisualBrush.Transform>
                            <TranslateTransform X="5" Y="3" />
                        </VisualBrush.Transform>
                        <VisualBrush.Visual>
                            <Grid HorizontalAlignment="Left">
                              <TextBlock FontFamily="SEGOEWP"  FontSize="10"   FontWeight="Normal" 
                                   HorizontalAlignment="Left" VerticalAlignment="Center" Foreground="Gray" FontStyle="Italic" Opacity="1"  
                                   Text="Enter Group Description"/>
                            </Grid>
                        </VisualBrush.Visual>
                    </VisualBrush>
                </Grid.Resources>
                <TextBox x:Name="G_Description" MinWidth="300" Margin="2,2,2,2" >
                  <TextBox.Style>
                     <Style TargetType="{x:Type TextBox}">
                       <Setter Property="Background" Value="Transparent"/>
                         <Style.Triggers>
                           <DataTrigger Binding="{Binding ElementName=G_Description,Path=Text}" Value="" >
                              <Setter Property="Background" Value="{StaticResource LoginHint}"/>
                           </DataTrigger>
                         </Style.Triggers>
                        </Style>
                    </TextBox.Style>
                 </TextBox>
             </Grid>
         </Viewbox>

      <!-- ******************** Text box  ********************--> 

        <Viewbox Grid.Row="3" Grid.Column="1" HorizontalAlignment="Left" Stretch="Uniform">
          <Grid >
               <Grid.Resources>
                   <VisualBrush x:Key="LoginHint" Stretch="None" AlignmentX="left" AlignmentY="Top" >
                        <VisualBrush.Transform>
                            <TranslateTransform X="5" Y="3" />
                        </VisualBrush.Transform>
                        <VisualBrush.Visual>
                            <Grid HorizontalAlignment="Left">
                              <TextBlock FontFamily="SEGOEWP"  FontSize="10"   FontWeight="Normal" 
                                   HorizontalAlignment="Left" VerticalAlignment="Center" Foreground="Gray" FontStyle="Italic" Opacity="1"  
                                   Text="Enter Mail Nickname"/>
                            </Grid>
                        </VisualBrush.Visual>
                    </VisualBrush>
                </Grid.Resources>
                <TextBox x:Name="G_MailName" MinWidth="150" Margin="2,2,2,2" >
                  <TextBox.Style>
                     <Style TargetType="{x:Type TextBox}">
                       <Setter Property="Background" Value="Transparent"/>
                         <Style.Triggers>
                           <DataTrigger Binding="{Binding ElementName=G_MailName,Path=Text}" Value="" >
                              <Setter Property="Background" Value="{StaticResource LoginHint}"/>
                           </DataTrigger>
                         </Style.Triggers>
                        </Style>
                    </TextBox.Style>
                 </TextBox>
             </Grid>
         </Viewbox>

       <!-- ******************** Text box  ********************--> 

       <!-- ******************** Text box  ********************-->

       <Viewbox Grid.Row="4" Grid.Column="1" HorizontalAlignment="Left" VerticalAlignment="Top" Stretch="Uniform"  Grid.ColumnSpan="2" >
          <Grid >
               <Grid.Resources>
                   <VisualBrush x:Key="LoginHint" Stretch="None" AlignmentX="left" AlignmentY="Top" >
                        <VisualBrush.Transform>
                            <TranslateTransform X="5" Y="3" />
                        </VisualBrush.Transform>
                        <VisualBrush.Visual>
                            <Grid HorizontalAlignment="Left">
                              <TextBlock FontFamily="SEGOEWP"  FontSize="10"   FontWeight="Normal" 
                                   HorizontalAlignment="Left" VerticalAlignment="Center" Foreground="Gray" FontStyle="Italic" Opacity="1"  
                                   Text="Enter Membership Rule"/>
                            </Grid>
                        </VisualBrush.Visual>
                    </VisualBrush>
                </Grid.Resources>
                <TextBox x:Name="G_MembershipRule" MinWidth="300" Margin="2,2,2,2"  MinHeight="50" TextWrapping="Wrap" AcceptsReturn="True">
                  <TextBox.Style>
                     <Style TargetType="{x:Type TextBox}">
                       <Setter Property="Background" Value="Transparent"/>
                         <Style.Triggers>
                           <DataTrigger Binding="{Binding ElementName=G_MembershipRule,Path=Text}" Value="" >
                              <Setter Property="Background" Value="{StaticResource LoginHint}"/>
                           </DataTrigger>
                         </Style.Triggers>
                        </Style>
                    </TextBox.Style>
                 </TextBox>
             </Grid>
         </Viewbox>

       <!-- ******************** Text box  ********************-->

       <Label x:Name="Create_Group_Result" Grid.Row="5" Grid.Column="1" FontSize="16"  VerticalAlignment="Center" HorizontalAlignment="Left" />

         <Button x:Name="Create_Group" Grid.Row="5" Grid.Column="2" 
                 Style="{StaticResource ButtonTemplate1}" 
                 Content="Create Group" 
                 HorizontalAlignment="Right" 
                 Margin="5,5,5,5" 
                 VerticalAlignment="Center" 
                 Width="130"
                 Height="40"
                 FontSize="17" 
                 />
    </Grid>
</Window>
"@

# Add assembly
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 
$Reader=(New-Object System.Xml.XmlNodeReader $xaml)
$Window=[Windows.Markup.XamlReader]::Load( $Reader )

$xaml.SelectNodes("//*[@*[contains(translate(name(.), 'n', 'N'), 'Name')]]") | ForEach-Object {
     New-Variable  -Name $_.Name -Value $Window.FindName($_.Name) -Force -ErrorAction SilentlyContinue
}

# Main Window's functions
$window.Add_MouseDoubleClick({
     $window.set_windowstate("Normal")
}) 

$Window.Add_MouseLeftButtonDown({
     $window.DragMove()
})

$Window.Add_Loaded({
     $Main_grd.Visibility     = "Visible"
     $Users_grd.Visibility    = "Hidden"
     $Devices_grd.Visibility  = "Hidden"
     $Groups_lv.Visibility    = "Hidden"
     $USers_lv.Visibility     = "Hidden"
     $Notif_lb.Visibility     = "Hidden"
     $Add_btn.Visibility      = "Hidden"
     $Remove_btn.Visibility   = "Hidden"
     $Devices_Btn.IsEnabled   = $False
     $Users_Btn.IsEnabled     = $False
     $Devices_Btn.Opacity     = "0.2"
     $Users_Btn.Opacity       = "0.2"
     $Overview_grd.Visibility = "Hidden"
     $SKU_lv.Visibility       = "Hidden"
     $Dash_grd.Visibility     = "Hidden"
     $Login_tb.Text = "admin@M365x898520.onmicrosoft.com"
}) 

$Exit_btn.Add_Click({
     $window.Close()
     #Remove-Variable * -ErrorAction SilentlyContinue
     #Remove-Module * 
     $error.Clear()
})

$Min_btn.Add_Click({
     $window.set_windowstate("Minimized")
})

$Max_btn.Add_Click({
     $window.set_windowstate("Maximized")
})

$Main_Btn.Add_Click({
     $Main_grd.Visibility    = "Visible"
     $Users_grd.Visibility   = "Hidden"
     $Devices_grd.Visibility = "Hidden"
})

$Users_Btn.Add_Click({

     $Main_grd.Visibility    = "Hidden"
     $Users_grd.Visibility   = "Visible"
     $Devices_grd.Visibility = "Hidden"
     $Groups_lv.Visibility   = "Visible"
     $USers_lv.Visibility    = "Visible"
     $Notif_lb.Visibility    = "Visible"
     $Add_btn.Visibility     = "Visible"
     $Remove_btn.Visibility  = "Visible"
      
})

$Devices_Btn.Add_Click({
     $Main_grd.Visibility    = "Hidden"
     $Users_grd.Visibility   = "Hidden"
     $Devices_grd.Visibility = "Visible"
})

$Login_Btn.Add_Click({


    if($global:authToken){

         $DateTime = (Get-Date).ToUniversalTime()
         $TokenExpires = ($authToken.ExpiresOn.datetime - $DateTime).Minutes

         if($TokenExpires -le 0){

             [System.Windows.MessageBox]::Show('Authentication Token expired ' +  $TokenExpires + 'minutes ago' , 'Authentication Token expired','OK','Information')

             if($Login_tb.Text -eq $null -or $Login_tb.Text -eq ""){

                 [System.Windows.MessageBox]::Show('Please specify your user principal name for Azure Authentication' , 'Wrong UPN','OK','Information')
             }
             else{

                 $global:authToken = Get-AuthToken -User $(($Login_tb.Text).Trim())
             }
         } 
     }
     else{
         if($Login_tb.Text -eq $null -or $Login_tb.Text -eq ""){

              [System.Windows.MessageBox]::Show('Please specify your user principal name for Azure Authentication' , 'Wrong UPN','OK','Information')
         }
         else{

             $global:authToken = Get-AuthToken -User $(($Login_tb.Text).Trim())
         }
     }

     if ( $global:authToken){

         $Props        = (Get-TennatInfo).verifiedDomains | Where-Object {$_.isDefault -eq "True"}
         $TennantiD    = (Get-TennatInfo).id
         $CompanyName  = (Get-TennatInfo).displayname
         $Capabilities = $Props.capabilities
         $Tennant      = $Props.name

         $Company_lb.Content = $CompanyName
         $Tennant_lb.Content = $Tennant
         $Tennantid_lb.Content = $TennantiD
         $Capabilities_lb.Content = $Capabilities
         
         $Overview_grd.Visibility = "Visible"
         #Get-Registration

         $Users = Get-AADUser
         $Groups = Get-AADGroup
         #$MDevices = Get-ManagedDevices -IncludeEAS
         $Devices = Get-AADDevice
       
         $Users | Select-Object @{name="Id";ex={$_.id}},@{ name="UserName";ex={$_.Displayname}},@{ name="UPN";ex={$_.userprincipalname}} | ForEach-Object {$Users_lv.addchild($_)}

         $Users_lv.Items.SortDescriptions.Clear()
         $Users_lv.Items.SortDescriptions.Add([System.ComponentModel.SortDescription]::new("UserName", [System.ComponentModel.ListSortDirection]::Ascending))
       
         $Groups | Select-Object @{name="Id";ex={$_.id}},@{name="GroupName";ex={$_.Displayname}}, @{ name="isOnPrem";ex={if($_.onPremisesDomainName){"True"}else{"False"}}} ,@{ name="Mail";ex={$_.mail}}| ForEach-Object {
             $Groups_lv.addchild($_)
             $Groups_lv_d.addchild($_)
         }

         $Groups_lv.Items.SortDescriptions.Clear()
         $Groups_lv.Items.SortDescriptions.Add([System.ComponentModel.SortDescription]::new("GroupName", [System.ComponentModel.ListSortDirection]::Ascending))

         $Groups_lv_d.Items.SortDescriptions.Clear()
         $Groups_lv_d.Items.SortDescriptions.Add([System.ComponentModel.SortDescription]::new("GroupName", [System.ComponentModel.ListSortDirection]::Ascending))         
        
         <#
         $MDevices | Select-Object @{name="Id";ex={$_.id}},@{name="DeviceName";ex={$_.deviceName}}, @{ name="Vendor";ex={$_.manufacturer}}, 
         @{ name="OS";ex={$_.operatingSystem}}, @{ name="Version";ex={$_.osVersion}} | ForEach-Object {
             $Devices_lv.addchild($_)  
         }    
         #>
         $Devices | Select-Object @{name="Id";ex={$_.id}},@{name="DeviceName";ex={$_.displayName}}, @{ name="isMDM";ex={if($_.mdmAppId){"True"}else{"False"}}}, 
         @{ name="OS";ex={$_.operatingSystem}}, @{ name="Version";ex={$_.operatingSystemVersion}}, @{ name="LastSign";ex={(get-date -Date $_.approximateLastSignInDateTime -Format 'dd.MM.yyyy HH:mm:ss').ToString()}} | ForEach-Object {
             $Devices_lv.addchild($_)  
         }    

         $Devices_lv.Items.SortDescriptions.Clear()
         $Devices_lv.Items.SortDescriptions.Add([System.ComponentModel.SortDescription]::new("DeviceName", [System.ComponentModel.ListSortDirection]::Ascending)) 
         
         $Subscription = Get-Subscription
         
         $Subscription | Select-Object @{name="SKU";ex={ ConvertSKUto-FrendlyName -SKU $_.skuPartNumber }},@{name="Enabled";ex={$_.prepaidUnits.enabled}} | ForEach-Object{
            $SKU_lv.addchild($_) 
            #$_.skuPartNumber; $_.prepaidUnits.enabled; $_.prepaidUnits.suspended;$_.prepaidUnits.warning
        
        }
         
         $SKU_lv.Visibility        = "Visible"

         $Devices_Btn.IsEnabled   = $True
         $Users_Btn.IsEnabled     = $True
         $Devices_Btn.Opacity     = "1"
         $Users_Btn.Opacity       = "1"

         
         $Params_dev = @{}
         $Params_usr = @{}
         $Params_grp = @{}
     
         $Params_dev.MDMDevices= @{}
         $Params_dev.MDMDevices.Header = "Intune managed"
         $Params_dev.MDMDevices.Value = ($Devices | Where-Object {$null -ne $_.mdmAppId}).Count
         $Params_dev.nMDMDevices = @{}
         $Params_dev.nMDMDevices.Header = "AAD joined"
         $Params_dev.nMDMDevices.Value = ($Devices | Where-Object {$null -eq $_.mdmAppId}).Count

         $Params_usr.User = @{}
         $Params_usr.User.Header = "AAD Users" 
         $Params_usr.User.Value =  $Users.Count

         $Params_grp.AADGroup = @{}
         $Params_grp.AADGroup.Header = "AAD Groups" 
         $Params_grp.AADGroup.Value =  ($Groups | Where-Object {$null -eq $_.onPremisesDomainName}).Count
         $Params_grp.ADGroup = @{}
         $Params_grp.ADGroup.Header = "AD Groups" 
         $Params_grp.ADGroup.Value =  ($Groups | Where-Object {$null -ne $_.onPremisesDomainName}).Count


         $script:hash = @{}
         New-Chart $Params_dev -ChartTitle "Azure AD devices ($($Devices.Count)) - total " -Type Pie -PieLabelStyle "Inside"
         $Chart_dev.Source = $script:hash.Stream
         $script:hash = @{}
         New-Chart $Params_usr -ChartTitle "AAD user count" -Type Pie -PieLabelStyle "Inside" -Color "White"
         $Chart_usr.Source = $script:hash.Stream
         $script:hash = @{}
         New-Chart $Params_grp -ChartTitle "AAD groups ($($Groups.Count)) - total " -Type Pie -PieLabelStyle "Inside"
         $Chart_grp.Source = $script:hash.Stream

     }

     $Dash_grd.Visibility        = "Visible"
})

$Add_btn.Add_Click({

     if(!($Users_lv.SelectedItems)){

         [System.Windows.MessageBox]::Show('No user selected' , 'No user selected','OK','Information')
     }
     else{
         if(!($Groups_lv.SelectedItems)){

             [System.Windows.MessageBox]::Show('No group selected' , 'No group selected','OK','Information')

         }
         else{
           
             foreach($User in $Users_lv.SelectedItems){
            
                 foreach($Group in $Groups_lv.SelectedItems){

                     $isMember = Get-AADGroup -GroupName $Group.GroupName -Members | Where-Object {$_.id -eq $User.id}

                     if (!($isMember)){
                         Add-AADGroupMember -GroupId $Group.id -AADMemberId $User.id
                         Show-Toast -Title "The AAD object has been changed" -Message "$($User.UserName) was added into '$($Group.GroupName)' AAD group."
                     }
                     else{
                         Show-Toast -Title "The AAD object hasn't been changed" -Message "$($User.UserName) is already in the '$($Group.GroupName)' AAD group."
                     }
                 }
             } 
         }
     }
})

$Remove_btn.Add_Click({

     if(!($Users_lv.SelectedItems)){
         [System.Windows.MessageBox]::Show('No user selected' , 'No user selected','OK','Information')
     }
     else{

         if(!($Groups_lv.SelectedItems)){
             [System.Windows.MessageBox]::Show('No group selected' , 'No group selected','OK','Information')
         }
         else{
            
             foreach($User in $Users_lv.SelectedItems){
            
                 foreach($Group in $Groups_lv.SelectedItems){

                    $isMember = Get-AADGroup -GroupName $Group.GroupName -Members | Where-Object {$_.id -eq $User.id}

                     if ($isMember){

                         Remove-AADGroupMember -GroupId $Group.id -AADMemberId $User.id
                         Show-Toast -Title "The AAD object has been changed" -Message "$($User.UserName) was removed from the '$($Group.GroupName)' AAD group"
                     }
                     else{
                          Show-Toast -Title "The AAD object hasn't been changed" -Message "$($User.UserName) hasn't found in the '$($Group.GroupName)' AAD group."
                     }
                 }
             } 
         }
     }
})

$Add_btn_d.Add_Click({

    if(!($Devices_lv.SelectedItems)){

        [System.Windows.MessageBox]::Show('No device selected' , 'No device selected','OK','Information')
    }
    else{
        if(!($Groups_lv_d.SelectedItems)){

            [System.Windows.MessageBox]::Show('No group selected' , 'No group selected','OK','Information')

        }
        else{
          
            foreach($Device in $Devices_lv.SelectedItems){
           
                foreach($Group in $Groups_lv_d.SelectedItems){

                    $isMember = Get-AADGroup -GroupName $Group.GroupName -Members | Where-Object {$_.id -eq $Device.id}

                    if (!($isMember)){
                        Add-AADGroupMember -GroupId $Group.id -AADMemberId $Device.id
                        Show-Toast -Title "The AAD object has been changed" -Message "$($Device.DeviceName) was added into '$($Group.GroupName)' AAD group."
                    }
                    else{
                        Show-Toast -Title "The AAD object hasn't been changed" -Message "$($Device.DeviceName) is already in the '$($Group.GroupName)' AAD group."
                    }
                }
            } 
        }
    }
})    

$Remove_btn_d.Add_Click({

    if(!($Devices_lv.SelectedItems)){
        [System.Windows.MessageBox]::Show('No device selected' , 'No device selected','OK','Information')
    }
    else{

        if(!($Groups_lv_d.SelectedItems)){
            [System.Windows.MessageBox]::Show('No group selected' , 'No group selected','OK','Information')
        }
        else{
           
            foreach($Device in $Devices_lv.SelectedItems){
           
                foreach($Group in $Groups_lv_d.SelectedItems){

                   $isMember = Get-AADGroup -GroupName $Group.GroupName -Members | Where-Object {$_.id -eq $Device.id}

                    if ($isMember){

                        Remove-AADGroupMember -GroupId $Group.id -AADMemberId $Device.id
                        Show-Toast -Title "The AAD object has been changed" -Message "$($Device.DeviceName) was removed from the '$($Group.GroupName)' AAD group"
                    }
                    else{
                         Show-Toast -Title "The AAD object hasn't been changed" -Message "$($Device.DeviceName) hasn't found in the '$($Group.GroupName)' AAD group."
                    }
                }
            } 
        }
    }
})    

$Users_lv_SelectAll.Add_Checked({
    foreach($U_item in $Users_lv.Items){
                   
        $Users_lv.SelectedItems.Add($U_item);          
    } 

})

$Users_lv_SelectAll.Add_UnChecked({
    foreach($U_item in $Users_lv.Items){
                   
        $Users_lv.SelectedItems.Remove($U_item);          
    } 

})

$Groups_lv_SelectAll.Add_Checked({
    foreach($G_item in $Groups_lv.Items){
                   
        $Groups_lv.SelectedItems.Add($G_item);          
    } 

})

$Groups_lv_SelectAll.Add_UnChecked({
    foreach($G_item in $Groups_lv.Items){
                   
        $Groups_lv.SelectedItems.Remove($G_item);          
    } 

})

$Groups_lv_d_SelectAll.Add_Checked({
    foreach($G_d_item in $Groups_lv_d.Items){
                   
        $Groups_lv_d.SelectedItems.Add($G_d_item);          
    } 

})

$Groups_lv_d_SelectAll.Add_UnChecked({
    foreach($G_d_item in  $Groups_lv_d.Items){
                   
        $Groups_lv_d.SelectedItems.Remove($G_d_item);          
    } 

})

$Devices_lv_SelectAll.Add_Checked({
    foreach($D_item in $Devices_lv.Items){
                   
        $Devices_lv.SelectedItems.Add($D_item);          
    } 

})

$Devices_lv_SelectAll.Add_UnChecked({
    foreach($D_item in $Devices_lv.Items){
                   
        $Devices_lv.SelectedItems.Remove($D_item);          
    } 

})

$Groups_lv_d_Asc_sort.Add_Click({
    $Groups_lv_d.Items.SortDescriptions.Clear()
    $Groups_lv_d.Items.SortDescriptions.Add([System.ComponentModel.SortDescription]::new("GroupName", [System.ComponentModel.ListSortDirection]::Ascending))
 })
 
 $Groups_lv_d_Desc_sort.Add_Click({
    $Groups_lv_d.Items.SortDescriptions.Clear()
    $Groups_lv_d.Items.SortDescriptions.Add([System.ComponentModel.SortDescription]::new("GroupName", [System.ComponentModel.ListSortDirection]::Descending))
 })

 $Groups_lv_Asc_sort.Add_Click({
    $Groups_lv.Items.SortDescriptions.Clear()
    $Groups_lv.Items.SortDescriptions.Add([System.ComponentModel.SortDescription]::new("GroupName", [System.ComponentModel.ListSortDirection]::Ascending))
 })
 
 $Groups_lv_Desc_sort.Add_Click({
    $Groups_lv.Items.SortDescriptions.Clear()
    $Groups_lv.Items.SortDescriptions.Add([System.ComponentModel.SortDescription]::new("GroupName", [System.ComponentModel.ListSortDirection]::Descending))
 })

 $Users_lv_Asc_sort.Add_Click({
    $Users_lv.Items.SortDescriptions.Clear()
    $Users_lv.Items.SortDescriptions.Add([System.ComponentModel.SortDescription]::new("UserName", [System.ComponentModel.ListSortDirection]::Ascending))
 })
 
 $Users_lv_Desc_sort.Add_Click({
    $Users_lv.Items.SortDescriptions.Clear()
    $Users_lv.Items.SortDescriptions.Add([System.ComponentModel.SortDescription]::new("UserName", [System.ComponentModel.ListSortDirection]::Descending))
 })

 $Devices_lv_Asc_sort.Add_Click({
    $Devices_lv.Items.SortDescriptions.Clear()
    $Devices_lv.Items.SortDescriptions.Add([System.ComponentModel.SortDescription]::new("DeviceName", [System.ComponentModel.ListSortDirection]::Ascending))
 })
 
 $Devices_lv_Desc_sort.Add_Click({
    $Devices_lv.Items.SortDescriptions.Clear()
    $Devices_lv.Items.SortDescriptions.Add([System.ComponentModel.SortDescription]::new("DeviceName", [System.ComponentModel.ListSortDirection]::Descending))
 })

$Groups_lv_Refresh.Add_Click({

    $Groups_lv.Items.Clear()
    $Groups_lv_d.Items.Clear()
    
    $Groups = Get-AADGroup

     $Groups | Select-Object @{name="Id";ex={$_.id}},@{name="GroupName";ex={$_.Displayname}}, @{ name="isOnPrem";ex={if($_.onPremisesDomainName){"True"}else{"False"}}} ,@{ name="Mail";ex={$_.mail}}| ForEach-Object {
        $Groups_lv.addchild($_)
        $Groups_lv_d.addchild($_)
    }

    $Groups_lv.Items.SortDescriptions.Clear()
    $Groups_lv.Items.SortDescriptions.Add([System.ComponentModel.SortDescription]::new("GroupName", [System.ComponentModel.ListSortDirection]::Ascending))

    $Groups_lv_d.Items.SortDescriptions.Clear()
    $Groups_lv_d.Items.SortDescriptions.Add([System.ComponentModel.SortDescription]::new("GroupName", [System.ComponentModel.ListSortDirection]::Ascending)) 
})

# Child Window's functions
 $Add_grp_btn.Add_Click({

     $Reader_g = (New-Object System.Xml.XmlNodeReader $grp_xaml)
     $grp_Window   = [Windows.Markup.XamlReader]::Load( $Reader_g )
     $grp_xaml.SelectNodes("//*[@*[contains(translate(name(.), 'n', 'N'), 'Name')]]") | ForEach-Object {
        New-Variable  -Name $_.Name -Value $grp_Window.FindName($_.Name) -Force -ErrorAction SilentlyContinue
     }

     $Create_Group.Add_Click({
        
        if($G_DynamicMembership.IsChecked -and $G_Unified.IsChecked){

            New-AADGroup -GroupType OfficeDynamic `
                         -displayName ($G_displayName.Text).Trim() `
                         -mailEnabled $G_mailEnabled.IsChecked `
                         -mailNickname ($G_displayName.Text).Trim()`
                         -membershipRule ($G_MembershipRule.Text).Trim() `
                         -description ($G_Description.Text).Trim() `
                         -securityEnabled $G_securityEnabled.IsChecked
        }
        elseif($G_DynamicMembership.IsChecked){
                   
            New-AADGroup -GroupType Security `
                         -displayName ($G_displayName.Text).Trim() `
                         -mailNickname ($G_displayName.Text).Trim()`
                         -membershipRule ($G_MembershipRule.Text).Trim() `
                         -description ($G_Description.Text).Trim() `
                         -securityEnabled $G_securityEnabled.IsChecked
        }
        elseif($G_Unified.IsChecked){

            New-AADGroup -GroupType Office `
                         -displayName ($G_displayName.Text).Trim() `
                         -mailEnabled $G_mailEnabled.IsChecked `
                         -description ($G_Description.Text).Trim() `
                         -mailNickname ($G_MailName.Text).Trim()
        }
        else{
            [System.Windows.MessageBox]::Show('Choose Group Type' , 'Choose Group Type','OK','Warning')
        }
     })   
    
    $G_DynamicMembership.add_Checked({
     
         if(!($G_Unified.IsChecked)){
        
             $G_mailEnabled.IsEnabled  = $false
             $G_MailName.IsEnabled     = $false

         }
     })

    $G_DynamicMembership.add_UnChecked({
     
        $G_mailEnabled.IsEnabled  = $true
        $G_MailName.IsEnabled     = $true

     })

    $G_Unified.add_Checked({
     
         if($G_DynamicMembership.IsChecked){
        
             $G_mailEnabled.IsEnabled  = $true
             $G_MailName.IsEnabled     = $true

         }
     })

     $G_Unified.add_UnChecked({
     
        if($G_DynamicMembership.IsChecked){
       
            $G_mailEnabled.IsEnabled  = $false
            $G_MailName.IsEnabled     = $false

        }
        else{
            $G_mailEnabled.IsEnabled  = $true
            $G_MailName.IsEnabled     = $true           
        }
    })     
     
    $G_displayName.Add_TextChanged({
        $Create_Group_Result.Visibility = "Hidden"
    })

     $base64 = "iVBORw0KGgoAAAANSUhEUgAAABgAAAAVCAIAAADTi7lxAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAFiUAABYlAUlSJPAAAAAoSURBVDhPY/z//z8DNQATlKYYjBpEGIwaRBiMGkQYjBpEGAxbgxgY
     AGPbAydP2uQjAAAAAElFTkSuQmCC"
     $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
     $bitmap.BeginInit()
     $bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($base64)
     $bitmap.EndInit()
     $bitmap.Freeze()
	 $grp_Window.Icon = $bitmap
     $grp_Window.ShowDialog()
 })


$async = $window.Dispatcher.InvokeAsync({$window.ShowDialog() | Out-Null})
$async.Wait() | Out-Null

#Remove-Variable * -ErrorAction SilentlyContinue
#Remove-Module * 
$error.Clear()