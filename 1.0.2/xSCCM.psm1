<#
    Custom functions for common operations in SCCM and MDT.
    
    Created 06/20/16
    
    Changelog:
        06/20/16 - v 1.0.0
            Initial build
            Added Split-DriverSource
        03/16/17 - v 1.0.1
            Added Get-CMCollectionMembership
        06/27/17 - v 1.0.2
            Added Clear-CMCache
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

#region Clear-CMCache
function Clear-CMCache
{
    <#
    ******************************************
    Name: Clear-CMCache
    Purpose: Remotely clear c:\windows\ccmcache on target computers
    Author: Andrew Ogden
    Email: andrew.ogden@dxpe.com

    Scriptblock function borrowed from user 0byt3 in this Reddit thread: https://www.reddit.com/r/SCCM/comments/3m8uh9/script_sms_client_to_clear_cache_then_install/
    #>
    param
    (
        [array]$ComputerName,
        [switch]$ResetWUCache
    )

    For ($i = 0; $i -lt $($ComputerName.Count); $i++)
    {
        # Create a session object for easy cleanup so we aren't leaving half-open
        # remote sessions everywhere
        try
        {
            Write-Progress -Activity "Clearing remote caches..." -Status "$($ComputerName[$i]) ($i/$($ComputerName.Count))" -PercentComplete ($($i/$($ComputerName.Count))*100)
            $CacheSession = New-PSSession -ComputerName $ComputerName[$i] -ErrorAction Stop
        }
        catch
        {
            Write-Host "$(Get-Date -UFormat "%m/%d/%y - %H:%M:%S") > ERROR: Failed to create session for $($ComputerName[$i])"
            Write-Host -ForegroundColor Yellow -Object $($error[0].Exception.Message)
            continue
        }

        # How big is the CM Cache?
        # We'll access the remote session a first time here to set up the COM object
        # and gather some preliminay data. We're also saving the cache size into a
        # local variable here for some reporting
        $SpaceSaved = Invoke-Command -Session $CacheSession -ScriptBlock {
            # Create CM object and gather cache info
            $cm = New-Object -ComObject UIResource.UIResourceMgr
            $cmcache = $cm.GetCacheInfo()
            $CacheElements = $cmcache.GetCacheElements()

            # Report space in use back to the local variable in MB
            $(($cmcache.TotalSize - $cmcache.FreeSize))
         }

        # Clear the CM cache
        # Now we're accessing the session a second time to clear the cache (assuming it's not  already empty)
        Invoke-Command -Session $CacheSession -ScriptBlock {
            if ($CacheElements.Count -gt 0)
            {
                # Echo total cache size
                Write-Host "$(($cmcache.TotalSize - $cmcache.FreeSize))" -NoNewline -ForegroundColor Yellow
                Write-Host " MB used by $(($cmcache.GetCacheElements()).Count) cache items on $env:computername"

                # Remove each object
                foreach ($CacheObj in $CacheElements)
                {
                    # Log individual elements
                    $eid = $CacheObj.CacheElementId
                    #Write-Host "Removing content ID $eid with size $(($CacheObj.ContentSize) / 1000)MB from $env:ComputerName"

                    # Delete content object
                    $cmcache.DeleteCacheElement($eid)
                }
            }
            else
            {
                Write-Host "Cache already empty on $env:ComputerName!"
            }
        }

        # Clean the WU cache (if requested)
        if ($ResetWUCache)
        {
            # This time we're going to access the remote session to count the size of the 
            # WU cache and add that to the existing variable
            $SpaceSaved += Invoke-Command -Session $CacheSession -ScriptBlock {
                $SizeCount = 0
                foreach ($f in (Get-childItem -Path "$env:SystemRoot\SoftwareDistribution" -Recurse))
                {
                    $SizeCount += $f.Length
                }

                # Report size in mb
                $SizeCount / 1mb
            }

            # Now we hop back into the remote session again to finish clearing
            # out the WU cache
            Invoke-Command -Session $CacheSession -ScriptBlock {
                Stop-Service wuauserv -Force -WarningAction SilentlyContinue

                Write-Host "Resetting WU Cache on $env:ComputerName..."
                Remove-Item -Path "$env:SystemRoot\SoftwareDistribution" -Force -Recurse

                # Restart WU and wait a few seconds for it to create a new cache folder
                Start-Service wuauserv -WarningAction SilentlyContinue
                Start-Sleep -Seconds 10

                # Verify that a new cache folder was created and throw an error if not
                if (!(Get-Item -Path "$env:SystemRoot\SoftwareDistribution"))
                {
                    Write-Host -Object "Failed to recreate SoftwareDistribution folder!" -ForegroundColor Red
                }
            }

            # We're accessing the session again a final time to determine the new size of the
            # WU cache to subtract from our saved space
            $SpaceSaved -= Invoke-Command -Session $CacheSession -ScriptBlock {
                $SizeCount = 0
                foreach ($f in (Get-childItem -Path "$env:SystemRoot\SoftwareDistribution" -Recurse))
                {
                    $SizeCount += $f.Length
                }

                # Report size in mb
                $SizeCount / 1mb
            }
        }

        # Report the space saved
        Write-Host -Object "Space saved on $($ComputerName[$i]): " -NoNewline
        Write-Host -Object $("{0:N2}" -f $SpaceSaved) -ForegroundColor Green -NoNewline
        Write-Host -Object " MB"

        # Clean up the session when done
        try
        {
            Remove-PSSession -Session $CacheSession -ErrorAction Stop
        }
        catch
        {
            Write-Host "ERROR: Failed to clean up session for $($ComputerName[$i])"
            Write-Host -ForegroundColor Yellow -Object $($error[0].Exception.Message)
            continue
        }

        # Clean up this variable to ensure that it doesn't bleed into subsequent iterations
        Clear-Variable -Name SpaceSaved
    }
}

<#
Example output:

PS C:\temp> Clear-CMCache -ComputerName dxpepc2314 -ResetWUCache
12796 MB used by 1411 cache items on DXPEPC2314
Resetting WU Cache on DXPEPC2314...
Space saved on dxpepc2314: 15,452.77 MB

PS C:\temp>
#>
#endregion