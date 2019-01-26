#############################################################################################
#      Moves an entire OneDrive from one user to another user's OneDrive subfolder 
#      Example: 
#       http://my.contoso.on.ca/personal/user1/Documents/ 
#                   will be moved to 
#       http://my.contoso.on.ca/personal/user2/Documents/archive_user1
#############################################################################################
$DESTINATION_FOLDER_PREFIX = "archive_"
$OneDriveRootURL = "http://my.contoso.on.ca"

$ErrorActionPreference = "Stop";

Clear-Host
$host.ui.RawUI.WindowTitle = "Copy files from one OneDrive to another"

# Check prerequisites for the script
if (Get-Module -ListAvailable -Name SharePointPnPPowerShell*) {
   # OK
} else {
    Write-Host "Module SharePointPnPPowerShell* does not exist. Please, Run these two commands: " -ForegroundColor Red
    Write-Host "Install-Module SharePointPnPPowerShellOnline" -ForegroundColor Yellow
    Write-Host "Set-ExecutionPolicy Unrestricted -Force" -ForegroundColor Yellow
    PAUSE
    exit
}

function VerifySiteCollectionURL($OneDriveURL){
    try{
        $site = $null
        Connect-PnPOnline �Url $OneDriveURL -CurrentCredentials   
        $site = Get-PnPSite -Includes Usage

        $k = 1/$site.Usage.StoragePercentageUsed
        $storageUsed = [math]::Round($site.Usage.Storage/1MB)
        $MaximumQuota = [math]::Round($storageUsed * $k)
        $percentageUsed = [math]::Round($site.Usage.StoragePercentageUsed, 2)

        Write-Host [OK] $OneDriveURL URL was verified  -ForegroundColor Green 
        Write-host `t[Info] OneDrive size: $("$storageUsed MB out of $MaximumQuota MB. It's $percentageUsed% of the total capacity")

        return @{ StorageUsed=$storageUsed; MaximumQuota=$MaximumQuota };
    }
    catch {
        $site = $null    
        Write-Host -ForegroundColor Yellow [Error] Could not retrieve OneDrive for $OneDriveURL
        $wshell = New-Object -ComObject Wscript.Shell
        $value = $wshell.Popup("$("$OneDriveURL does not exist or cannot be found")",0,"ERROR",0x1)
        Throw
    }
}


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# SEND FROM:
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Copy OneDrive'
$form.Size = New-Object System.Drawing.Size(320,320)
$form.StartPosition = 'CenterScreen'


$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(330,20)
$label.Text = 'From (login name without domain)'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

# SEND TO:
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10,90)
$label2.Size = New-Object System.Drawing.Size(330,20)
$label2.Text = 'To (login name without domain)'
$form.Controls.Add($label2)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(10,110)
$textBox2.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox2)


$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(10,190)
$OKButton.Size = New-Object System.Drawing.Size(260,53)
$OKButton.Text = 'Copy OneDrive'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)


$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $fromUser = $textBox.Text
    $toUser = $textBox2.Text
     
    Write-host -ForegroundColor Blue Pre-copy checks:
    $sourceStorageInfo = VerifySiteCollectionURL $("$OneDriveRootURL/personal/$fromUser")
    $destinationStorageInfo = VerifySiteCollectionURL $("$OneDriveRootURL/personal/$toUser")

    if (($sourceStorageInfo["StorageUsed"] + $destinationStorageInfo["StorageUsed"]) -gt $destinationStorageInfo["MaximumQuota"] ){
        Write-host -ForegroundColor Magenta "[Cancelled!] There is not enough space in the destination OneDrive. You need at least" ($sourceStorageInfo["StorageUsed"] + $destinationStorageInfo["StorageUsed"])MB of space, but you only have ($destinationStorageInfo["MaximumQuota"])MB. Increase the quota for the destination and then restart this app.
        PAUSE
        exit
    } else{
        Write-host -ForegroundColor Green "[OK] Destination has enough space to contain the source OneDrive"
    }
    

    $wshell = New-Object -ComObject Wscript.Shell
    $proceed = $wshell.Popup("$("$OneDriveURL Do you want to proceed?")",0,"Everything looks good?",0x1)
    
    if($proceed -eq 1){
        Connect-PnPOnline -URL $OneDriveRootURL -CurrentCredentials   

        Copy-PnPFile -SourceUrl $("personal/$fromUser/Documents") -TargetUrl $("/personal/$toUser/Documents/$DESTINATION_FOLDER_PREFIX$fromUser") -SkipSourceFolderName -OverwriteIfAlreadyExists -Force

        $IE=new-object -com internetexplorer.application
        $DestinationURL = $("$OneDriveRootURL/personal/$toUser/Documents/")
        $IE.navigate2($DestinationURL)
        $IE.visible=$true

        Write-host -ForegroundColor Yellow "[Success] $($destinationStorageInfo["StorageUsed"])MB of files moved from $($fromUser) to $($toUser)"
        PAUSE
    }

}