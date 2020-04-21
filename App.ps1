[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null

Add-Type -AssemblyName PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing 

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
        write-host
        write-host "AzureAD Powershell module not installed..." -f Red
        write-host "Install by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt" -f Yellow
        write-host "Script can't continue..." -f Red
        write-host
        exit
    }

# Getting path to ActiveDirectory Assemblies
# If the module count is greater than 1 find the latest version

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

        Write-Host
        Write-Host "Authorization Access Token is null, please re-run authentication..." -ForegroundColor Red
        Write-Host
        break

        }

    }

    catch {

    write-host $_.Exception.Message -f Red
    write-host $_.Exception.ItemName -f Red
    write-host
    break

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
    write-host
    break

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
    write-host
    break

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
    write-host
    break

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
    Write-Host "Response content:`n$responseBody" -f Red
    Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    write-host
    break

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

#endregion functions

# XAML
[xml]$xaml = @"
<Window
	xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
	xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
	Name="window" 
	WindowStyle="None"
    Title="AAD Objects Management"
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

        <!-- ******************** Window Border Style ****************************** --> 

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
                            <Rectangle 
                                Margin="15,0,0,0"
                                StrokeThickness="1"
                                Stroke="#60000000"
                                StrokeDashArray="1 2"/>
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

        <Button x:Name = "Groups_Btn" Grid.Row="2" FontSize="20" Background="SteelBlue" Foreground="White" Style="{StaticResource ButtonTemplate}" Margin ="0 0 5 5">
           <Button.Content>
              <Viewbox StretchDirection="DownOnly" Stretch="Uniform">
                  <ContentControl Content="AAD Groups"/>
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

<!-- ********************** Main grid  ShowGridLines="True"********************-->
    
    <Grid x:Name = "Main_grd" Grid.Row="0" Grid.Column="1">
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

<!-- ********************** Header  ********************-->

      <Grid Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="5">
          <Rectangle Fill="SteelBlue" Margin ="5,0,0,0" RadiusY="3" RadiusX="3"/> 
          <Viewbox StretchDirection="DownOnly" Stretch="Uniform" HorizontalAlignment="Right">             
                <TextBlock Text="Azure AD Objects Management" Margin="0,0,35,0" Foreground="White" FontSize="30" FontFamily="Segoe UI" />
          </Viewbox>
      </Grid>

<!-- ********************** End of Header  ********************-->     

<!-- ********************** Login part  ********************-->

      <Viewbox StretchDirection="Both" Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2" Stretch="Uniform" HorizontalAlignment="Right" >
          <TextBox x:Name = "Login_tb" 
                Style="{StaticResource TextBoxTemplate}" 
                Padding="2,2,2,2" 
                Margin="5,5,5,5" 
                FontSize="15" 
                HorizontalAlignment="Left" 
                MinWidth="400" 
                MinHeight="15" 
                FontFamily="Segoe UI"
                
                />
   
      </Viewbox>

      <Button x:Name = "Login_Btn" Grid.Row="2" Grid.Column="3" FontSize="15" Background="White" Foreground="Black" Style="{StaticResource ButtonTemplate}" Margin ="5,5,5,5" >
           <Button.Content>
              <Viewbox StretchDirection="DownOnly" Stretch="Uniform">
                  <ContentControl Content="Login"/>
              </Viewbox>
           </Button.Content>
      </Button>

 <!-- ********************** End of Login part ********************-->

<!-- ********************** Tables  ********************-->

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
           
         <ListView x:Name="Users_lv" Grid.Column="0" Margin="0,0,5,0" FontSize="12" ItemsSource="{Binding}" IsSynchronizedWithCurrentItem="True">
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
                <GridView>
                    <GridView.Columns>
                        <GridViewColumn>
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                   <CheckBox Tag="{Binding ID}" IsChecked="{Binding RelativeSource={RelativeSource AncestorType={x:Type ListViewItem}}, Path=IsSelected}" Style="{StaticResource CheckBoxTemplate}"/>  
                               </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn DisplayMemberBinding="{Binding UserName}" Header="User Name" />
                        <GridViewColumn DisplayMemberBinding="{Binding UPN}" Header="UPN" />
                    </GridView.Columns>
                </GridView>
            </ListView.View>
        </ListView>
   

      <ListView x:Name="Groups_lv" Grid.Column="1"  FontSize="12" ItemsSource="{Binding}" IsSynchronizedWithCurrentItem="True" ScrollViewer.CanContentScroll="True"  HorizontalAlignment="Stretch" VerticalAlignment="Top">
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
                            <MenuItem Header="Ascending" />
                            <MenuItem Header="Descending" />
                        </ContextMenu>
                  </GridView.ColumnHeaderContextMenu>
                
                <GridView.Columns>
                    <GridViewColumn>
                        <GridViewColumn.CellTemplate>
                            <DataTemplate>
                               <CheckBox Tag="{Binding ID}" IsChecked="{Binding RelativeSource={RelativeSource AncestorType={x:Type ListViewItem}}, Path=IsSelected}" Style="{StaticResource CheckBoxTemplate}"/>  
                           </DataTemplate>
                        </GridViewColumn.CellTemplate>
                    </GridViewColumn>
                    <GridViewColumn DisplayMemberBinding="{Binding GroupName}" Header="Group Name" />
                    <GridViewColumn DisplayMemberBinding="{Binding isOnPrem}"  Header="isOnPrem" />
                    <GridViewColumn Header="Mail">
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

<!-- ********************** End of Tables  ********************-->

<!-- ********************** Add Button and Text  ********************-->

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

<!-- ********************** End of Add Button and Text  ********************-->

<!-- ********************** Footer  ********************-->

      <Grid Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="5" Margin="0,0,-5,-5">
          <Rectangle Fill="LightGray" />  
          <Viewbox StretchDirection="DownOnly" Stretch="Uniform">         
             <TextBlock x:Name="Footer_tb" Margin="3,3,0,0" Foreground="White" FontSize="15" FontFamily="Segoe UI"/>
          </Viewbox>
      </Grid>   

<!-- ********************** End of Footer  ********************-->

    </Grid>

<!-- ********************** End of Main grid  ********************-->

    <Grid x:Name = "Users_grd" Grid.Row="0" Grid.Column="1" ShowGridLines="True">
       <Grid.RowDefinitions>
          <RowDefinition />
          <RowDefinition />
       </Grid.RowDefinitions>
       <Grid.ColumnDefinitions>
         <ColumnDefinition/>
         <ColumnDefinition />
       </Grid.ColumnDefinitions>
 
    </Grid>

    <Grid x:Name = "Groups_grd" Grid.Row="0" Grid.Column="1" ShowGridLines="True">
       <Grid.RowDefinitions>
          <RowDefinition />
          <RowDefinition />
        </Grid.RowDefinitions>
       <Grid.ColumnDefinitions>
         <ColumnDefinition/>
         <ColumnDefinition />
         <ColumnDefinition />
      </Grid.ColumnDefinitions>
    </Grid>

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

$window.Add_MouseDoubleClick({
    
   $window.set_windowstate("Normal")
}) 

$Window.Add_MouseLeftButtonDown({

    $window.DragMove()

})

$Window.Add_Loaded({
    $Main_grd.Visibility   = "Visible"
    $Users_grd.Visibility  = "Hidden"
    $Groups_grd.Visibility = "Hidden"

    $Groups_lv.Visibility  = "Hidden"
    $USers_lv.Visibility  = "Hidden"

    $Notif_lb.Visibility  = "Hidden"
    $Add_btn.Visibility  = "Hidden"
    $Remove_btn.Visibility  = "Hidden"

}) 

$Exit_btn.Add_Click({
  
   $window.Close()

   Remove-Variable * -ErrorAction SilentlyContinue
   Remove-Module * 
   $error.Clear()
})

$Min_btn.Add_Click({
    $window.set_windowstate("Minimized")
})

$Max_btn.Add_Click({
    $window.set_windowstate("Maximized")
})

$Main_Btn.Add_Click({
    $Main_grd.Visibility   = "Visible"
    $Users_grd.Visibility  = "Hidden"
    $Groups_grd.Visibility = "Hidden"
})

$Users_Btn.Add_Click({
    $Main_grd.Visibility   = "Hidden"
    $Users_grd.Visibility  = "Visible"
    $Groups_grd.Visibility = "Hidden"

})

$Groups_Btn.Add_Click({
    $Main_grd.Visibility   = "Hidden"
    $Users_grd.Visibility  = "Hidden"
    $Groups_grd.Visibility = "Visible"
})

$Login_Btn.Add_Click({

   #[System.Windows.MessageBox]::Show('Computer found: ', 'Find in Active Directory','OK','Information')

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


    if ($global:authToken){

        #$Footer_tb.Text = "Azure Authentication is successful"

    }
    else{

        #$Footer_tb.Text = "Azure Authentication is failed"
    }

  }
    
  $Users = Get-AADUser
  $Groups = Get-AADGroup

  #$Footer_tb.Text = "Data collection..."

  $USers | Select-Object @{name="Id";ex={$_.id}},@{ name="UserName";ex={$_.Displayname}},@{ name="UPN";ex={$_.userprincipalname}} | ForEach-Object {$Users_lv.addchild($_)}

  $Groups | Select-Object @{name="Id";ex={$_.id}},@{name="GroupName";ex={$_.Displayname}}, @{ name="isOnPrem";ex={if($_.onPremisesDomainName){"True"}else{"False"}}} ,@{ name="Mail";ex={$_.mail}}| ForEach-Object {$Groups_lv.addchild($_)}
  
  $Groups_lv.Visibility  = "Visible"
  $USers_lv.Visibility   = "Visible"

  $Notif_lb.Visibility   = "Visible"
  $Add_btn.Visibility    = "Visible"
  $Remove_btn.Visibility = "Visible"
     
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

$window.ShowDialog() | Out-Null

