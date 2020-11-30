# This script must be run on a MECM Management Point

param (
    [Parameter(Position = 0,Mandatory=$true)]
    [string]$Folder,
    [Parameter(Position = 1,Mandatory=$true)]
    [string]$SiteCode
)

#Load Configuration Manager PowerShell Module
Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5)+ '\ConfigurationManager.psd1')

#Get SiteCode
$SiteCode = (Get-PSDrive -PSProvider CMSITE | Select-Object -First 1)
Set-Location $SiteCode":"

#Error Handling and output
$FormatEnumerationLimit = -1
#Clear-Host
$ErrorActionPreference= 'Stop'

# From https://gallery.technet.microsoft.com/scriptcenter/ConfigMgr-UpdateRefresh-68041cc7
Function Update-CMDeviceCollection
{
    <#
    .Synopsis
       Update SCCM Device Collection
    .DESCRIPTION
       Update SCCM Device Collection. Use the -Wait switch to wait for the update to complete.
    .EXAMPLE
       Update-CMDeviceCollection -DeviceCollectionName "All Workstations"
    .EXAMPLE
       Update-CMDeviceCollection -DeviceCollectionName "All Workstations" -Wait -Verbose
    #>

    [CmdletBinding()]
    [OutputType([int])]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $DeviceCollectionName
    )

    Begin {
        Write-Verbose "$DeviceCollectionName : Update Started"
    }
    Process {
        $Collection = Get-CMDeviceCollection -Name $DeviceCollectionName
        $null = Invoke-WmiMethod `
        -Path "ROOT\SMS\Site_$($SiteCode.Name):SMS_Collection.CollectionId='$($Collection.CollectionId)'" `
        -Name RequestRefresh `
        -ComputerName $SiteCode.Root
    }
    End { 
    
    }
}


#Import Device Collection files
Write-Host "Loading Configurations..."
$Collections = @()
$DCFiles=( Get-ChildItem $PSScriptRoot/collections/ -Filter '*.json' | Sort-Object -Property Name )
ForEach ( $File in $DCFiles ) {
    $Collections+=( Get-Content $File.FullName | ForEach-Object { $ExecutionContext.InvokeCommand.ExpandString($_) } | ConvertFrom-Json )
    Write-Host -ForegroundColor Green ("Imported Configuration:" + $File.Name )
}

# Create Default Folder
  # Parent Container Node ID is from:
    # (Get-WmiObject -Namespace "root\sms\site_XXX" -Query 'select * from SMS_ObjectContainerNode' -ComputerName <CAS OR Primary FQDN> | ? {$_.Name -eq "Some Folder"}).ParentContainerNodeID

$CollectionFolder = @{Name = $Folder; ObjectType = 5000; ParentContainerNodeId = 16}
Set-WmiInstance -Namespace "root\sms\site_$($SiteCode.Name)" -Class "SMS_ObjectContainerNode" -Arguments $CollectionFolder -ComputerName $SiteCode.Root -ErrorAction Ignore | Out-Null
$FolderPath = ($SiteCode.Name +":\DeviceCollection\" + "Production\" + $CollectionFolder.Name)

#Find Existing Collections
$ExistingCollections = Get-CMDeviceCollection -Name "${Ship} - *"

ForEach ( $Collection In $Collections ) {
    $DCChanged = $False
    # Create DCs if missing
    If ( -Not ( $ExistingCollections.Name -Contains $Collection.Name ) ) {
        Write-Host ( "No collection named `"" + $Collection.Name + "`". Creating a new one.")
        # Create DC if missing
        Try {
            New-CMDeviceCollection -Name $Collection.Name -Comment ($Collection.Comment + "`nGitLab-CI Note: This Device Collection was created using GitLab-CI and will be overwritten next time the pipeline executes.") -LimitingCollectionName $Collection.LimitingCollection -RefreshType 2 | Out-Null
            Write-Host -ForegroundColor Green ( "`tNew Device Collection Created: " + $Collection.Name )
        }
        Catch {
            Write-host "-----------------"
            Write-host -ForegroundColor Red ("There was an error creating the: " + $Collection.Name + " collection.")
            Write-host "-----------------"
            $ErrorCount += 1
            $_.Exception|format-list -force
            Pause
        }
        # Move DC to correct folder
        Try {
            Move-CMObject -FolderPath $FolderPath -InputObject $(Get-CMDeviceCollection -Name $Collection.Name)
            Write-Host -ForegroundColor Green ( "`t`"" + $Collection.Name + "`" moved to `"" + $CollectionFolder.Name + "`"" )
        }
        Catch {
            Write-host "-----------------"
            Write-host -ForegroundColor Red ("There was an error moving the `"" + $Collection.Name +"`" collection to `"" + $CollectionFolder.Name +"`".")
            Write-host "-----------------"
            $_.Exception|format-list -force
            Pause
        }
        $DCChanged = $True
    } Else {
        Write-Host -ForegroundColor Green ( "`"" + $Collection.Name + "`" already exists. Checking rules for compliance:" )
    }    
    
    # Create new or changed queries
    $DCQueries = ($ExistingCollections | Where-Object Name -eq $Collection.Name).CollectionRules | Where-Object ObjectClass -eq SMS_CollectionRuleQuery
    ForEach ( $Query in $Collection.Queries ) {
        if ( -Not ( $DCQueries.RuleName -Contains $Query.Name ) ) {
            Try {
                Add-CMDeviceCollectionQueryMembershipRule -CollectionName $Collection.Name -QueryExpression $Query.Expression -RuleName $Query.Name
            }
            Catch {
                Write-host "-----------------"
                Write-host -ForegroundColor Red ("There was an error creating the query rule `"" + $Query.Name +"`" for `"" + $Collection.Name +"`".")
                Write-host "-----------------"
                $_ | select * | format-list -force
                $_.Exception|format-list -force
                Pause
            }
            Write-Host -ForegroundColor Yellow ( "`tNew query `"" + $Query.Name + "`" created." )
            $DCChanged = $True
        } else {
            $DCMatchingQuery = $DCQueries | Where-Object RuleName -eq $Query.Name
            if ( -Not ( ($DCMatchingQuery.QueryExpression -replace '\s+',' ') -eq ($Query.Expression -replace '\s+',' ') ) ) {
                Try {
                    Remove-CMDeviceCollectionQueryMembershipRule -CollectionName $Collection.Name -RuleName $Query.Name -Force
                    Add-CMDeviceCollectionQueryMembershipRule -CollectionName $Collection.Name -QueryExpression $Query.Expression -RuleName $Query.Name
                }
                Catch {
                    Write-host "-----------------"
                    Write-host -ForegroundColor Red ("There was an error recreating query `"" + $Query.Name +"`" for `"" + $Collection.Name +"`".")
                    Write-host "-----------------"
                    $_.Exception|format-list -force
                    Pause
                }
                Write-Host -ForegroundColor Yellow ( "`tQuery mismatch for: `"" + $Query.Name + "`". Deleted and re-created." )
                $DCChanged = $True
            } else {
                Write-Host -ForegroundColor Green ( "`tQuery `"" + $Query.Name + "`" matches source. No change needed." )
            }
        }
    }

    # Delete removed queries
    ForEach ( $Query in $DCQueries ) {
        if ( -Not ( $Collection.Queries.Name -Contains $Query.RuleName ) ) {
            Try {
                Remove-CMDeviceCollectionQueryMembershipRule -CollectionName $Collection.Name -RuleName $Query.RuleName -Force
            }
            Catch {
                Write-host "-----------------"
                Write-host -ForegroundColor Red ("There was an error removing Query `"" + $Query.RuleName +"`" from `"" + $Collection.Name +"`".")
                Write-host "-----------------"
                $_.Exception|format-list -force
                Pause
            }
            Write-Host -ForegroundColor Yellow ( "`tRemoved Query `"" + $Query.RuleName + "`" from MECM. (Not found in source)" )
            $DCChanged = $True
        }
    }

    # Create new direct members
    $DCDirectMembers = ($ExistingCollections | Where-Object Name -eq $Collection.Name).CollectionRules | Where-Object ObjectClass -eq SMS_CollectionRuleDirect
    ForEach ( $Member in $Collection.DirectMembers ) {
        if ( -Not ( $DCDirectMembers.RuleName -Contains $Member ) ) {
            Try {
                $DeviceResourceID = (Get-CMDevice -name $Member).ResourceID
                Add-CMDeviceCollectionDirectMembershipRule -CollectionName $Collection.Name -ResourceId $DeviceResourceID
            }
            Catch {
                Write-host "-----------------"
                Write-host -ForegroundColor Red ("There was an error adding `"" + $Member +"`" to `"" + $Collection.Name +"`".")
                Write-host "-----------------"
                $_.Exception|format-list -force
                Pause
            }
            Write-Host -ForegroundColor Yellow ( "`tNew direct member `"" + $Member + "`" added." )
            $DCChanged = $True
        } else {
                Write-Host -ForegroundColor Green ( "`tDirect Member `"" + $Member + "`" already exists. No change needed." )
        }
    }

    # Delete removed direct members
    ForEach ( $Member in $DCDirectMembers ) {
        if ( -Not ( $Collection.DirectMembers -Contains $Member.RuleName ) ) {
            Try {
                Remove-CMDeviceCollectionDirectMembershipRule -CollectionName $Collection.Name -ResourceName $Member.RuleName -Force
            }
            Catch {
                Write-host "-----------------"
                Write-host -ForegroundColor Red ("There was an error removing `"" + $Member.RuleName +"`" from `"" + $Collection.Name +"`".")
                Write-host "-----------------"
                $_.Exception|format-list -force
                Pause
            }
            Write-Host -ForegroundColor Yellow ( "`tRemoved `"" + $Member.RuleName + "`" from `"" + $Collection.Name + "`". (Not found in source)" )
            $DCChanged = $True
        }
    }

    # Create new Include Collection rules
    $DCIncludeMembers = ($ExistingCollections | Where-Object Name -eq $Collection.Name).CollectionRules | Where-Object ObjectClass -eq SMS_CollectionRuleIncludeCollection
    ForEach ( $IncludeCollection in $Collection.IncludeCollections ) {
        if ( -Not ( $DCIncludeMembers.RuleName -Contains $IncludeCollection ) ) {
            Try {
                Add-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection.Name -IncludeCollectionName $IncludeCollection
            }
            Catch {
                Write-host "-----------------"
                Write-host -ForegroundColor Red ("There was an error adding Include Collection rule `"" + $IncludeCollection +"`" to `"" + $Collection.Name +"`".")
                Write-host "-----------------"
                $_ | select * | format-list -force
                $_.Exception|format-list -force
                Pause
            }
            Write-Host -ForegroundColor Yellow ( "`tNew Include Collection rule `"" + $IncludeCollection + "`" added." )
            $DCChanged = $True
        } else {
                Write-Host -ForegroundColor Green ( "`tInclude Collection rule `"" + $IncludeCollection + "`" already exists. No change needed." )
        }
    }

    # Delete removed Include Collection rules
    ForEach ( $Member in $DCIncludeMembers ) {
        if ( -Not ( $Collection.IncludeCollections -Contains $Member.RuleName ) ) {
            Try {
                Remove-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection.Name -IncludeCollectionName $Member.RuleName -Force
            }
            Catch {
                Write-host "-----------------"
                Write-host -ForegroundColor Red ("There was an error removing Include Collection rule `"" + $Member.RuleName +"`" from `"" + $Collection.Name +"`".")
                Write-host "-----------------"
                $_.Exception|format-list -force
                Pause
            }
            Write-Host -ForegroundColor Yellow ( "`tRemoved Include Collection rule `"" + $Member.RuleName + "`" from `"" + $Collection.Name + "`". (Not found in source)" )
            $DCChanged = $True
        }
    }

    # Create new Exclude Collection rules
    $DCExcludeMembers = ($ExistingCollections | Where-Object Name -eq $Collection.Name).CollectionRules | Where-Object ObjectClass -eq SMS_CollectionRuleExcludeCollection
    ForEach ( $ExcludeCollection in $Collection.ExcludeCollections ) {
        if ( -Not ( $DCExcludeMembers.RuleName -Contains $ExcludeCollection ) ) {
            Try {
                Add-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection.Name -ExcludeCollectionName $ExcludeCollection
            }
            Catch {
                Write-host "-----------------"
                Write-host -ForegroundColor Red ("There was an error adding Exclude Collection rule `"" + $ExcludeCollection +"`" to `"" + $Collection.Name +"`".")
                Write-host "-----------------"
                $_ | select * | format-list -force
                $_.Exception|format-list -force
                Pause
            }
            Write-Host -ForegroundColor Yellow ( "`tNew Exclude Collection rule `"" + $ExcludeCollection + "`" added." )
            $DCChanged = $True
        } else {
                Write-Host -ForegroundColor Green ( "`tExclude Collection rule `"" + $ExcludeCollection + "`" already exists. No change needed." )
        }
    }

    # Delete removed Exclude rules
    ForEach ( $Member in $DCExcludeMembers ) {
        if ( -Not ( $Collection.ExcludeCollections -Contains $Member.RuleName ) ) {
            Try {
                Remove-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection.Name -ExcludeCollectionName $Member.RuleName -Force
            }
            Catch {
                Write-host "-----------------"
                Write-host -ForegroundColor Red ("There was an error removing Exclude Collection rule `"" + $Member.RuleName +"`" from `"" + $Collection.Name +"`".")
                Write-host "-----------------"
                $_.Exception|format-list -force
                Pause
            }
            Write-Host -ForegroundColor Yellow ( "`tRemoved Exclude Collection rule `"" + $Member.RuleName + "`" from `"" + $Collection.Name + "`". (Not found in source)" )
            $DCChanged = $True
        }
    }

    if ( $DCChanged -eq $True ) {
        Write-Host -ForegroundColor Yellow ( "`tChanges to `"" + $Collection.Name + "`" detected. Refreshing Device Collection." )
        Update-CMDeviceCollection -DeviceCollectionName $Collection.Name
    }
}

#Rename missing DCs
ForEach ( $Collection in $ExistingCollections ) {
    if ( -Not ( $Collections.Name -Contains $Collection.Name ) -And -Not ( $Collection.Name -Like 'NOT IN GIT - *') ) {
        Try {
            Set-CMCollection -Name $Collection.Name -NewName ( "NOT IN GIT - " + $Collection.Name )
        }
        Catch {
            Write-host "-----------------"
            Write-host -ForegroundColor Red ("There was an error moving `"" + $Collection.Name)
            Write-host "-----------------"
            $_.Exception|format-list -force
            Pause
        }
        Write-Host -ForegroundColor Yellow ( "`tMoved `"" + $Collection.Name + "`" to `"" + "NOT IN GIT - " + $Collection.Name + "`"" )
    }
}
