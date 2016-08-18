<#
    Custom functions for common operations in SCCM and MDT.
    
    Created 06/20/16
    
    Changelog:
        06/20/16 - v 1.0.0
            Initial build
            Added Split-DriverSource
#>

#region Split-DriverSource 
function Split-DriverSource
{
    param
    (
        [string]$SourcePath,
        [string]$DestPath,
        [switch]$Verbose
    )
    
    Import-Module PSAlphaFS
    
    $DestFileCount = 0
    $SourceFileCount = 0
    
    if (!(Test-Path -Path $DestPath))
    {
        New-Item -Path $DestPath -ItemType Directory | Out-Null
    }
    
    $FullSource = Get-ChildItem -Path $SourcePath -Recurse
    $SourceFileCount = ($FullSource | ?{$_.Attributes -ne 'Directory'}).Count
    
    $DriverFiles = $FullSource | ?{
        ($_.Extension -eq '.bin') -or
        ($_.Extension -eq '.cab') -or
        ($_.Extension -eq '.cat') -or
        ($_.Extension -eq '.dll') -or
        ($_.Extension -eq '.inf') -or
        ($_.Extension -eq '.ini') -or
        ($_.Extension -eq '.oem') -or
        ($_.Extension -eq '.sys')
    }
    $DestFileCount = ($DriverFiles | ?{$_.Attributes -ne 'Directory'}).Count
    
    foreach ($File in $DriverFiles)
    {
        $SourceFile = $File.FullName
        $ReplacePath = $SourcePath -replace '\\','\\'
        $DestDir = $SourceFile -replace ($ReplacePath, $DestPath)
        
        if (!(Test-Path -Path $DestDir))
        {
            New-Item -Path $DestDir -ItemType Directory | Out-Null
        }
        
        Copy-Item -Path $SourceFile -Destination $DestDir | Out-Null
    }
    
    $TotalStats = New-Object -TypeName psobject
    $TotalStats | Add-Member -MemberType NoteProperty -Name FilesKept -Value $DestFileCount
    $TotalStats | Add-Member -MemberType NoteProperty -Name FilesDropped -Value ($SourceFileCount - $DestFileCount)
    $TotalStats | Add-Member -MemberType NoteProperty -Name OriginalFiles -Value $SourceFileCount
    $TotalStats | Add-Member -MemberType NoteProperty -Name SourceSizeGB -Value (Get-FolderSize -Path $SourcePath).SizeinGB
    $TotalStats | Add-Member -MemberType NoteProperty -Name NewSizeGB -Value (Get-FolderSize -Path $DestPath).SizeInGB
    
    return $TotalStats
}
<#
    Example Output:
    
    PS C:\> Split-DriverSource -SourcePath '\\<Driversource>\Windows7x64-old' -DestPath '<Driversource>\Windows7x64'

    FilesKept     : 932
    FilesDropped  : 1526
    OriginalFiles : 2458
    SourceSizeGB  : 1.29
    NewSizeGB     : 0.59
    
    PS C:\>
#>

#endregion
#region Update-CMSiteName
function Update-CMSiteName
{
    param
    (
        [string]$SiteName = 'HOU',
        [string]$SiteServer = 'housccm03.dxpe.com',
        [string]$NewSiteDesc
    )
    
    $FullSite = Get-WmiObject -Class 'SMS_SCI_SiteDefinition' -Namespace "root/SMS/site_$SiteName" -ComputerName $SiteServer
    
    if (!($NewSiteDesc))
    {
        Write-Host "Current site description is - $($FullSite.SiteName)"
        $NewSiteDesc = Read-Host -Prompt "Enter new description: "   
    }
    
    $OldSiteDesc = $FullSite.SiteName
    $FullSite.SiteName = $NewSiteDesc
    $FullSite.Put()
    
    $CurrentSiteDesc = (Get-WmiObject -Class 'SMS_SCI_SiteDefinition' -Namespace "root/SMS/site_$SiteName" -ComputerName $SiteServer).SiteName
    if ($CurrentSiteDesc -ne $NewSiteDesc)
    {
        Write-Host 'There was an error updating the site description.' -ForegroundColor Red
    }
    else
    {
        Write-Host "Site description successfully updated.`r`nOld: $OldSiteDesc`r`nNew: $NewSiteDesc"
    }
}
#endregion