<#
    Custom functions for common operations in SCCM and MDT.
    
    Created 06/20/16
    
    Changelog:
        06/20/16 - v 1.0.0
            Initial build
            Added Split-DriverSource
        03/16/17 - v 1.0.1
            Added Get-CMCollectionMembership
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

#region Get-CMCollectionMembership
Function Get-CMCollectionMembership
{
    <# 
            .SYNOPSIS 
                Determine the SCCM collection membership.
            .DESCRIPTION
                This function allows you to determine the SCCM collection membership of a given user/computer.
            .PARAMETER  Type 
                Specify the type of member you are querying. Possible values : 'User' or 'Computer'
            .PARAMETER  ResourceName 
                Specify the name of your member : username or computername.
            .PARAMETER  SiteServer
                Specify the name of the site server to query.
            .PARAMETER  SiteCode
                Specify the site code on the targeted server.
            .EXAMPLE 
                Get-Collections -Type computer -ResourceName PC001
                Get-Collections -Type user -ResourceName User01
            .Notes 
                Author : Antoine DELRUE 
                WebSite: http://obilan.be 
    #> 

    param(
    [Parameter(Mandatory=$true,Position=1)]
    [ValidateSet("User", "Computer")]
    [string]$Type,

    [Parameter(Mandatory=$true,Position=2)]
    [string]$ResourceName,

    [Parameter(Mandatory=$false,Position=3)]
    [string]$SiteServer = 'housccm03.dxpe.com',

    [Parameter(Mandatory=$false,Position=4)]
    [string]$SiteCode = 'HOU'
    ) #end param

    Switch ($type)
        {
            User {
                Try {
                    $ErrorActionPreference = 'Stop'
                    $resource = Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Class "SMS_R_User" | ? {$_.Name -ilike "*$resourceName*"}                            
                }
                catch {
                    Write-Warning ('Failed to access "{0}" : {1}' -f $SiteServer, $_.Exception.Message)
                }

            }

            Computer {
                Try {
                    $ErrorActionPreference = 'Stop'
                    $resource = Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Class "SMS_R_System" | ? {$_.Name -ilike "$resourceName"}                           
                }
                catch {
                    Write-Warning ('Failed to access "{0}" : {1}' -f $SiteServer, $_.Exception.Message)
                }
            }
        }

    $ids = (Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Class SMS_CollectionMember_a -filter "ResourceID=`"$($Resource.ResourceId)`"").collectionID
    # A little trick to make the function work with SCCM 2012
    if ($ids -eq $null)
    {
            $ids = (Get-WmiObject -ComputerName $SiteServer -Namespace "root\sms\site_$SiteCode" -Class SMS_FullCollectionMembership -filter "ResourceID=`"$($Resource.ResourceId)`"").collectionID
    }

    $array = @()

    foreach ($id in $ids)
    {
        $Collection = get-WMIObject -ComputerName $SiteServer -namespace "root\sms\site_$SiteCode" -class sms_collection -Filter "collectionid=`"$($id)`""
        $Object = New-Object PSObject
        $Object | Add-Member -MemberType NoteProperty -Name "Collection Name" -Value $Collection.Name
        $Object | Add-Member -MemberType NoteProperty -Name "Collection ID" -Value $id
        $Object | Add-Member -MemberType NoteProperty -Name "Comment" -Value $Collection.Comment
        $array += $Object
    }

    return $array
}
#endregion