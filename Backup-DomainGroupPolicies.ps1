<#
.SYNOPSIS
Backup-DomainGroupPolicies Allows a user to create a .zip file filled with the current Group Policies in the domain.
These packages can also be restored from this script (It will also link these back to the OUs).

The naming scheme for the zip files is:
GPOBackup-month-day-year-hour-minute

.DESCRIPTION
Facilitates the backing up and restoring of Group Policies.

.EXAMPLE
./Backup-DomainGroupPolicies
Backup all policies in a Domain.

.EXAMPLE
./Backup-DomainGroupPolicies -Path \\file-server\backups -Policies "random policy",laps
Backup all policies in a Domain with a specified output path and only specified policies.

.EXAMPLE
./Backup.DomainGroupPolicies -Path C:\path\GPOBackup-04-16-2018-02-00
Restoring from a package generate by this script.
.LINK
https://github.com/bandit145/Backup-DomainGroupPolicies

#>
param(
    [Switch]$Restore,
    [String]$Path=(Get-Item -Path ./).FullName,
    [Array]$Policies=$null
)

$ErrorActionPreference = "Stop"
Add-Type -Assembly "system.io.compression.filesystem"

try{
    Import-Module -Name "GroupPolicy","ActiveDirectory"
}
catch {
    Write-Error "GroupPolicy/ActiveDirectory module is missing! Please install RSAT tools! https://www.microsoft.com/en-us/download/details.aspx?id=45520"
}

function Get-Links{
    param($links)
    $proper_links = [System.Collections.ArrayList]@()
    foreach ($link in $links){
        [System.Collections.ArrayList]$dn = (Get-ADDomain).DistinguishedName.Split(",")
        #just grabbin it from one place instead of splitting it again
        $dn.Reverse()
        foreach ($path in $link.SOMPath.Split("/")){
            if (($link.SOMPath.Split("/")).IndexOf($path) -gt 0){
                $dn.Add("ou=$path") | Out-Null
            }
        }
        #reverse to put into proper ldap dn order
        $dn.Reverse()
        $proper_links.Add($dn -join ",") | Out-Null
    }
    return $proper_links
}

function Backup-GPOs{
    #month-day-year-hour-minute
    $backup_dir = New-Item -Path (-join($Path,"GPOBackup-",(Get-Date -UFormat "%m-%d-%Y-%H-%M"))) -Type Directory
    if ($Policies){
        foreach ($policy in $Policies){
           Backup-GPO -Name $policy -Path $backup_dir.FullName | Out-Null
        }
    }
    else{
        Backup-GPO -All -Path $backup_dir.FullName | Out-Null
    }
    [io.compression.zipfile]::CreateFromDirectory($backup_dir.FullName,(-join($backup_dir.FullName,".zip")))
    Remove-item -Path $backup_dir.FullName  -Recurse -Force
    Write-Output (-join("INFO: All policies specified have been backed up to ",$backup_dir,".zip"))
}

function Restore-GPOs{
    $temp_folder = -join((Split-Path $SCRIPT:Myinvocation.mycommand.path -Parent),"\temp_gpo")
    New-Item -Path $temp_folder -Type Directory | Out-Null
    [io.compression.zipfile]::ExtractToDirectory($Path,$temp_folder)
    foreach($file in (Get-ChildItem -Path $temp_folder -Exclude "*.xml")){
        [xml]$gpo_info = Get-Content -Path (-join($file.FullName,"\gpreport.xml"))
        try{
            Get-GPO -Name $gpo_info.GPO.Name | Out-Null
            Write-Output (-join("INFO: ",$gpo_info.GPO.Name," exists in AD, importing settings!"))
        }
        catch{
            Write-Output (-join("INFO: ",$gpo_info.GPO.Name," is missing from AD, creating GPO!"))
        }
        Import-GPO -BackupGPOName $gpo_info.GPO.Name -TargetName $gpo_info.GPO.Name -Path $temp_folder -CreateIfNeeded | Out-Null
        foreach($link in (Get-Links $gpo_info.GPO.LinksTo)){
            if (Test-Path "AD:\$link"){
                try{
                    New-GPLink -Name $gpo_info.GPO.Name -Target $link -LinkEnabled yes | Out-Null
                    Write-Output (-join("INFO: Linking ",$gpo_info.GPO.Name," to ",$link))
                }
                catch{
                    Write-Output (-join("WARN: ",$gpo_info.GPO.Name," already seems linked to ",$link,", Continuing..."))
                }
            }
            else{
                Write-Output (-join("INFO: ",$link," does not exist on Domain, Ignoring and continuing for ",$gpo_info.GPO.Name))
            }
        }
    }
    Remove-item -Path $temp_folder -Recurse -Force
}

    
if($Restore){
    Restore-GPOs
}
else{
    Backup-GPOs
}