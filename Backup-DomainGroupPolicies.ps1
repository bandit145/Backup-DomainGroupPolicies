[CmdletBinding(DefaultParameterSetName="none")]
param(
    [Parameter(ParameterSetName="backup",Mandatory=$True)]
    [Int]$RetentionDays,
    [Parameter(ParameterSetName="restore",Mandatory=$True)]
    [String]$RestorePackage,
    [String]$Path=(Get-Item -Path ./).FullName+"\"
)

$ErrorActionPreference = "Stop"

Add-Type -Assembly "system.io.compression.filesystem"

try{
    Import-Module -Name "GroupPolicy","ActiveDirectory"
}
catch {
    Write-Error "GroupPolicy module is missing! Please install RSAT tools!"
}

function Get-GPOManifest{
    param($gpo)
    $manifest = New-Object -TypeName psobject -Property @{Name="";Links=[System.Collections.ArrayList]@();ID=""}
    [xml]$xml_report = Get-GPOReport -Name $gpo.displayname -ReportType xml
    $manifest.Name = $xml_report.GPO.Name
    $manifest.ID = $xml_report.GPO.Identifier.Identifier."#text".Trim("{}")
    if ($xml_report.GPO.LinksTo.length -gt 0){
        foreach ($link in $xml_report.GPO.LinksTo){
            [System.Collections.ArrayList]$dn = (Get-ADDomain).DistinguishedName.Split(",")
            foreach ($path in $link.SOMPath.Split("/")){
                if (($link.SOMPath.Split("/")).IndexOf($path) -gt 0){
                    $dn.Add("ou=$path") | Out-Null
                }
            }
            $manifest.Links.Add($dn -join ",") | Out-Null
        }
    }
    return $manifest
}

function Backup-GPOs{
    $backup_dir = New-Item -Path (-join($Path,"GPOBackup-",(Get-Date -UFormat "%m-%d-%y-%H"))) -Type Directory
    foreach($gpo in Get-GPO -All){
        $gpo_dir = New-Item -Path (-join($backup_dir.FullName,"\",$gpo.displayname)) -Type Directory
        Get-GPOManifest -gpo $gpo | ConvertTo-Json | Out-File -FilePath (-join($gpo_dir.FullName,"\manifest.json"))
        Backup-GPO -Name $gpo.displayname -Path (-join($backup_dir.FullName,"\",$gpo.displayname))
    }
    #[io.compression.zipfile]::CreateFromDirectory($backup_dir.FullName,(-join($backup_dir.FullName,".zip")))
    #Remove-item -Path $backup_dir.FullName  -Recurse
}

function Restore-GPOs{
    Expand-Archive -FilePath $RestorePackage -OutputPath $Path
    $package_name = $RestorePackage.split("\")[$RestorePackage.Split("\").Length-1]
    $backup_path= $Path+"\"+$package_name+"\"+$package_name
    foreach($file in (Get-ChildItem -Path $unzipped_file -Recurse)){
        $gpo_info = Get-Content -Path
        #see if the gpo exists if not do a full restore and relink (only if ou exists)
        #if it does exist just do a restore 
        try{
            Get-GPO -Name $gpo_info.Name
            Restore-GPO -Path (-join($backup_path,"\",$gpo_info.Name))
        }
        catch{
            Import-GPO -BackupGPOName $gpo_info.Name -TargetGPOName $gpo_info.Name -Path $backup_path -CreateIfNeeded
        }
        foreach($link in $gpo_info.Link){
            if (Test-Path "AD:\$link"){
                New-GPLink -Name $gpo_info.Name -LinkEnabled yes
            }
            else{
                Write-Output "INFO: "+$link+" does not exist on Domain, Ignoring and continuing for "+$gpo_info.Name
            }
        }
    }
}

function Main{
    Write-Output $path
    
    if($psCmdlet.ParameterSetName -eq "restore"){
        Restore-GPOs
    }
    else{
        Backup-GPOs
    }

}

Main