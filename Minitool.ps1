if (Get-Module -ListAvailable -Name WriteAscii) {
    #Do Nothing
} else {
    Write-Host "Missing Module: WriteAscii, Installing begins"
    Install-Module WriteAscii
}
if (Get-Module -ListAvailable -Name vmxtoolkit) {
    #Do Nothing
} else {
    Write-Host "Missing Module: vmxtoolkit, Installing begins"
    Install-Module vmxtoolkit
}

Import-Module WriteAscii
Import-Module activedirectory


Write-Ascii -InputObject 'MINITOOL'

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 


Function startmenu{
    Write-Host ""
    Write-Host "1) User Creation"
    Write-Host "2) OU Creation"
    Write-Host "3) Share Creation"
    Write-Host "4) VM Creation"
    Write-Host "5) CSV User Creation"

    do {$answer = Read-Host "Choose between 1 - 5"}
    until ("1","2","3","4","5" -ccontains $answer)

    switch($answer) {
        '1'{
            Bruger
        }
        '2'{
            OU
        }
        '3'{
            Share
        }
        '4'{
            VMOP
        }
        '5'{
            CSVCreation
        }
    }
}

function CSVCreation{

    Write-Host "Choose where you wanna create the csv file"

    $DirecPath = New-Object System.Windows.Forms.FolderBrowserDialog
    $null = $DirecPath.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))

    $csvName = Read-Host "Choose csv name"

    New-Item -Path $DirecPath.SelectedPath -Name "$($csvName).csv"

    Write-Host $DirecPath.SelectedPath


    Add-Content -Path "$($DirecPath.SelectedPath)\$($csvName).csv" -Value 'Firstname,Lastname,SAM,OU,Password,Description'

    Write-Host "$($DirecPath.SelectedPath)\$($csvName).csv"

    $Firstname = Read-Host "Type Firstname"
    $Lastname = Read-Host "Type Lastname"
    $SAM = Read-Host "Type SAM"
    $OU = Get-ADOrganizationalUnit -Filter * | select name,DistinguishedName | Out-GridView -OutputMode Single -Title "Select your OU"
    $Password = Read-Host "Type password. remeber it to meet the password complexity requirements"
    $Description = Read-Host "Type Description"

    $csvfile = @(
    $Firstname + ',' + $Lastname + ',' + $SAM + ',"' + $ou.DistinguishedName.ToString() + '",' + $Password + ',' + $Description
    )

    Write-Host "Firstname,Lastname,SAM,OU,Password,Description"
    $csvfile

    do {$answer = Read-Host "Is this correct yes or no"}
    until ("yes","no" -ccontains $answer)

    if ($answer -eq "yes") {
        $csvfile | foreach { Add-Content -Path "$($DirecPath.SelectedPath)\$($csvName).csv" -Value $_}
        Write-Host "CSV created: $($DirecPath.SelectedPath)\$($csvName).csv"
    }


}

Function Bruger{
    Write-Host "Choose your csv file"
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = 'CSV File (*.csv)|*.csv' }
    $null = $FileBrowser.ShowDialog()

    $Users = Import-Csv -Path $FileBrowser.FileName

    foreach ($User in $Users)
    {
        $Displayname = $User.'Firstname' + " " + $User.'Lastname'
        $UserFirstname = $User.'Firstname'
        $UserLastname = $User.'Lastname'
        $OU = $User.'OU'
        $SAM = $User.'SAM'
        $UPN = $User.'Firstname' + $User.'Lastname' + "@" + $User.'Maildomain'
        $Description = $User.'Description'
        $Password = $User.'Password'

        if (Get-ADUser -Filter "SamAccountName -eq '$SAM'") {
            Write-Host "$Displayname already exists."
        } else {
            Write-Host "Creating user: $Displayname"
            New-ADUser `
                    -Name "$Displayname" `
                    -Displayname "$Displayname" `
                    -SamAccountName $SAM `
                    -UserPrincipalName $UPN `
                    -GivenName "$UserFirstname" `
                    -Surname "$UserLastname" `
                    -Description "$Description" `
                    -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) `
                    -Enabled $true `
                    -Path "$OU" `
                    -ChangePasswordAtlogon $False `
                    -PasswordNeverExpires $true 
        }
    }
}

Function OU{
    Write-Host "Choose your csv file"
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = 'CSV File (*.csv)|*.csv' }
    $null = $FileBrowser.ShowDialog()

    $ADOU = Import-csv -Path $FileBrowser.FileName -Delimiter ","

    foreach ($ou in $ADOU) {
        $name = $ou.name
        $path = $ou.path

        $newOU = "OU=$($name),$($path)"

        if (Get-ADOrganizationalUnit -Filter "distinguishedName -eq '$newOU'") {
            Write-Host "$name already exists."
        } else {
            Write-Host "$name creating user."
            New-ADOrganizationalUnit `
                -Name $name `
                -path $path
        }
    }
}

Function Share{
    Write-Host "Choose your csv file"
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = 'CSV File (*.csv)|*.csv' }
    $null = $FileBrowser.ShowDialog()

    $Path = Read-Host "Type your share path"

    $FileBrowser1 = Import-csv -Path $FileBrowser.FileName -Delimiter ","

    foreach ($FileBrowsers in $FileBrowser1)
    {
        $User = $FileBrowsers.SAM

        New-SmbShare -Name $User -FullAccess "$($env:USERDNSDOMAIN)\$($User)"  -Path $Path  -FolderEnumerationMode Unrestricted
        Grant-FileShareAccess -Name $User -AccessRight "Full" -AccountName $User
    }
}

Function VMOP{
    Import-Module vmxtoolkit
    $vmxname = Read-Host -Prompt "Type the name of VM"
    $vmxDiskSize = Read-Host "Type disc size"
    $ConvertedvmxDiskSize = [int64]$vmxDiskSize.Replace('GB','') * 1GB

    Write-Host "ISO Path"
    $ISOPath = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = 'ISO File (*.iso)|*.iso' }
    $IPath = $ISOPath.ShowDialog()

    Write-Host "Choose where you wanna store your VM"
    $ChangePath = New-Object System.Windows.Forms.FolderBrowserDialog
    $CPath = $ChangePath.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))

    [int]$memmb = Read-Host -Prompt "Choose how much memeory you want to give your VM"

    New-VMX -VMXName $vmxname -Path $ChangePath.SelectedPath -Type Hyper-V -Firmware BIOS | 
    New-VMXScsiDisk -NewDiskSize $ConvertedvmxDiskSize -Path "$($ChangePath.SelectedPath)\$($vmxname)" -NewDiskname SCSI0_0 | 
    Add-VMXScsiDisk -LUN 0 -Controller 0 | Connect-VMXcdromImage -ISOfile $ISOPath.FileName | 
    Set-VMXNetworkAdapter -Adapter 0 -AdapterType e1000e  -ConnectionType nat | 
    Set-VMXmemory -VMXName $vmxname -MemoryMB $memmb
    
    Set-VMXDisplayName -DisplayName $vmxname -config "$($ChangePath.SelectedPath)\$($vmxname)\$($vmxname).vmx"
    
    Start-Sleep -Seconds 2

    Start-VMX -VMXName $vmxname -Path "$($ChangePath.SelectedPath)\$($vmxname)" -config "$($ChangePath.SelectedPath)\$($vmxname)\$($vmxname).vmx"
}

startmenu
