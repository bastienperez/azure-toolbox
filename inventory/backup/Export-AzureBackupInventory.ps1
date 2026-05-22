#Requires -Modules Az.Accounts, Az.RecoveryServices, Az.Compute, Az.Sql, Az.Storage, ImportExcel

<#
.SYNOPSIS
    Generates a global backup inventory and identifies unprotected resources across all accessible Azure subscriptions.

.DESCRIPTION
    Iterates through every Azure subscription, queries all Recovery Services Vaults,
    extracts Protected Items and identifies resources that are not protected by backup.
    Use -ExportToExcel to export results into an Excel report with multiple worksheets;
    otherwise the function returns a hashtable containing the collected data as PowerShell objects.

.PARAMETER ExportToExcel
    Switch to export the inventory to an Excel file.
    When omitted, the function returns a hashtable with keys BackupInventory, BackupPolicies,
    UnprotectedResources, and PolicyDistribution.

.OUTPUTS
    When -ExportToExcel is used: Azure_Backup_Global_Report.xlsx with multiple worksheets.
    Otherwise: hashtable with keys BackupInventory, BackupPolicies, UnprotectedResources, PolicyDistribution.
#>

function Export-AzureBackupInventory {
    [CmdletBinding()]
    param (
        [Parameter()]
        [switch]$ExportToExcel
    )

    $ErrorActionPreference = 'Stop'

    $context = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Write-Host 'No active Azure session, launching login...' -ForegroundColor Yellow
        $null = Connect-AzAccount
    }

    $subscriptions = Get-AzSubscription -ErrorAction Stop | Where-Object { $_.State -eq 'Enabled' }
    Write-Host "$($subscriptions.Count) subscription(s) found." -ForegroundColor Cyan

    [System.Collections.Generic.List[Object]]$backupInventory = @()
    [System.Collections.Generic.List[Object]]$backupPolicies = @()
    [System.Collections.Generic.List[Object]]$unprotectedResources = @()
    $subIndex = 0

    foreach ($sub in $subscriptions) {
        $subIndex++
        Write-Verbose "[LOOP] Starting subscription loop $subIndex of $($subscriptions.Count): $($sub.Name)"
        Write-Host "[$subIndex/$($subscriptions.Count)] $($sub.Name)" -ForegroundColor Cyan

        try {
            $null = Set-AzContext -SubscriptionId $sub.Id -ErrorAction Stop
            $vaults = Get-AzRecoveryServicesVault -ErrorAction Stop

            if ($vaults.Count -eq 0) {
                Write-Verbose "No vaults found in '$($sub.Name)'."
                continue
            }

            Write-Verbose "$($vaults.Count) vault(s) in '$($sub.Name)'"

            $vaultIndex = 0
            foreach ($vault in $vaults) {
                $vaultIndex++
                Write-Verbose "[LOOP] Starting vault loop $vaultIndex of $($vaults.Count): $($vault.Name)"

                $null = Set-AzRecoveryServicesVaultContext -Vault $vault -ErrorAction Stop
                try {
                    $policies = Get-AzRecoveryServicesBackupProtectionPolicy -VaultId $vault.ID -ErrorAction SilentlyContinue
                    if ($policies) {
                        Write-Verbose "Found $($policies.Count) backup policy(ies)"
                        foreach ($policy in $policies) {
                            # Extract schedule information
                            $scheduleFrequency = ''
                            $backupTimes = ''
                            $scheduleDays = ''
                            $windowStartTime = ''
                            $windowDuration = ''
                            $timeZone = ''
                            $interval = ''
                            
                            if ($policy.SchedulePolicy) {
                                $sched = $policy.SchedulePolicy
                                $scheduleFrequency = $sched.ScheduleRunFrequency

                                if ($sched.ScheduleRunTimeZone) {
                                    $timeZone = $sched.ScheduleRunTimeZone
                                }

                                # V1 policies (SimpleSchedulePolicy): schedule data lives at the top level.
                                # V2 policies (SimpleSchedulePolicyV2): schedule data lives in
                                # DailySchedule / WeeklySchedule / HourlySchedule sub-objects, so the
                                # top-level ScheduleRunTimes/ScheduleRunDays are empty.
                                if ($sched.ScheduleRunTimes) {
                                    $backupTimes = ($sched.ScheduleRunTimes | ForEach-Object { $_.ToString('HH:mm') }) -join ', '
                                }
                                if ($sched.ScheduleRunDays) {
                                    $scheduleDays = ($sched.ScheduleRunDays) -join ', '
                                }
                                if ($sched.ScheduleWindowStartTime) {
                                    $windowStartTime = $sched.ScheduleWindowStartTime.ToString('HH:mm')
                                }
                                if ($sched.ScheduleWindowDuration) {
                                    $windowDuration = "$($sched.ScheduleWindowDuration) hours"
                                }
                                if ($sched.ScheduleInterval) {
                                    $interval = "$($sched.ScheduleInterval) hours"
                                }

                                # V2 Daily/Weekly schedule
                                if ($sched.DailySchedule -and $sched.DailySchedule.ScheduleRunTimes) {
                                    $backupTimes = ($sched.DailySchedule.ScheduleRunTimes | ForEach-Object { $_.ToString('HH:mm') }) -join ', '
                                }
                                if ($sched.WeeklySchedule) {
                                    if ($sched.WeeklySchedule.ScheduleRunTimes) {
                                        $backupTimes = ($sched.WeeklySchedule.ScheduleRunTimes | ForEach-Object { $_.ToString('HH:mm') }) -join ', '
                                    }
                                    if ($sched.WeeklySchedule.ScheduleRunDays) {
                                        $scheduleDays = ($sched.WeeklySchedule.ScheduleRunDays) -join ', '
                                    }
                                }

                                # V2 Hourly schedule
                                if ($sched.HourlySchedule) {
                                    if ($sched.HourlySchedule.WindowStartTime) {
                                        $windowStartTime = $sched.HourlySchedule.WindowStartTime.ToString('HH:mm')
                                    }
                                    if ($sched.HourlySchedule.WindowDuration) {
                                        $windowDuration = "$($sched.HourlySchedule.WindowDuration) hours"
                                    }
                                    if ($sched.HourlySchedule.Interval) {
                                        $interval = "$($sched.HourlySchedule.Interval) hours"
                                    }
                                }
                            }

                            # Extract retention information
                            $dailyRetention = ''
                            $weeklyRetention = ''
                            $monthlyRetention = ''
                            $yearlyRetention = ''
                            if ($policy.RetentionPolicy) {
                                if ($policy.RetentionPolicy.DailySchedule) {
                                    $dailyRetention = "$($policy.RetentionPolicy.DailySchedule.DurationCountInDays) days"
                                }
                                if ($policy.RetentionPolicy.WeeklySchedule) {
                                    $weeklyRetention = "$($policy.RetentionPolicy.WeeklySchedule.DurationCountInWeeks) weeks"
                                }
                                if ($policy.RetentionPolicy.MonthlySchedule) {
                                    $monthlyRetention = "$($policy.RetentionPolicy.MonthlySchedule.DurationCountInMonths) months"
                                }
                                if ($policy.RetentionPolicy.YearlySchedule) {
                                    $yearlyRetention = "$($policy.RetentionPolicy.YearlySchedule.DurationCountInYears) years"
                                }
                            }

                            $policyRecord = [PSCustomObject]@{
                                SubscriptionName     = $sub.Name
                                ResourceGroupName    = $vault.ResourceGroupName
                                VaultName            = $vault.Name
                                PolicyName           = $policy.Name
                                WorkloadType         = $policy.WorkloadType
                                BackupManagementType = $policy.BackupManagementType
                                IsEnabled            = $policy.IsEnabled
                                Frequency            = $scheduleFrequency
                                ScheduleDays         = $scheduleDays
                                BackupTimes          = $backupTimes
                                WindowStartTime      = $windowStartTime
                                WindowDuration       = $windowDuration
                                TimeZone             = $timeZone
                                Interval             = $interval
                                DailyRetention       = $dailyRetention
                                WeeklyRetention      = $weeklyRetention
                                MonthlyRetention     = $monthlyRetention
                                YearlyRetention      = $yearlyRetention
                            }
                            $backupPolicies.Add($policyRecord)
                        }
                    }
                }
                catch {
                    Write-Warning "Could not retrieve backup policies: $($_.Exception.Message)"
                }

                $containerTypes = @(
                    'AzureVM'
                    'AzureSQL'
                    'AzureStorage'
                    'AzureVMAppContainer'
                    'Windows'
                )

                [System.Collections.Generic.List[Object]]$allItems = @()

                $containerTypeIndex = 0
                foreach ($containerType in $containerTypes) {
                    $containerTypeIndex++
                    Write-Verbose "[LOOP] Starting container type loop $containerTypeIndex of $($containerTypes.Count): $containerType"
                    try {
                        $containerParams = @{
                            VaultId       = $vault.ID
                            ContainerType = $containerType
                            ErrorAction   = 'SilentlyContinue'
                        }
                        $containers = Get-AzRecoveryServicesBackupContainer @containerParams

                        if (-not $containers) { 
                            Write-Verbose "No containers for '$containerType'"
                            continue 
                        }
                    
                        Write-Verbose "$($containers.Count) container(s) for '$containerType'"

                        # Map container type to workload type
                        $workloadTypeMapping = @{
                            'AzureVM'             = 'AzureVM'
                            'AzureSQL'            = 'AzureSQLDatabase'
                            'AzureStorage'        = 'AzureFiles'
                            'AzureVMAppContainer' = 'MSSQL'
                            'Windows'             = 'AzureVM'
                        }

                        $containerIndex = 0
                        foreach ($container in $containers) {
                            $containerIndex++
                            Write-Verbose "[LOOP] Starting container loop $containerIndex of $($containers.Count): $($container.Name)"
                            $itemParams = @{
                                VaultId      = $vault.ID
                                Container    = $container
                                WorkloadType = $workloadTypeMapping[$containerType]
                                ErrorAction  = 'SilentlyContinue'
                            }
                            $items = Get-AzRecoveryServicesBackupItem @itemParams

                            if ($items) {
                                Write-Verbose "[LOOP] Processing $($items.Count) items from container $($container.Name)"
                                foreach ($item in $items) {
                                    $allItems.Add($item)
                                }
                            }
                        }
                    }
                    catch {
                        Write-Warning "Could not query container type '$containerType': $($_.Exception.Message)"
                    }
                }

                # Also query workload-based items (SQL in VM, SAP HANA) directly
                Write-Verbose 'Scanning workload-based items...'
                $workloadTypes = @('AzureVM', 'AzureSQLDatabase', 'AzureFiles', 'MSSQL', 'SAPHanaDatabase')

                $workloadTypeIndex = 0
                foreach ($wlType in $workloadTypes) {
                    $workloadTypeIndex++
                    Write-Verbose "[LOOP] Starting workload type loop $workloadTypeIndex of $($workloadTypes.Count): $wlType"
                    try {
                        $workloadParams = @{
                            VaultId              = $vault.ID
                            BackupManagementType = 'AzureWorkload'
                            WorkloadType         = $wlType
                            ErrorAction          = 'SilentlyContinue'
                        }
                        $wlItems = Get-AzRecoveryServicesBackupItem @workloadParams

                        if ($wlItems) {
                            Write-Verbose "$($wlItems.Count) workload item(s) for '$wlType'"
                            Write-Verbose "[LOOP] Processing $($wlItems.Count) workload items for type $wlType"
                            foreach ($item in $wlItems) {
                                # Avoid duplicates by checking the item ID
                                if ($allItems.Id -notcontains $item.Id) {
                                    $allItems.Add($item)
                                }
                            }
                        }
                    }
                    catch {
                        # Silently skip unsupported workload/management type combinations
                    }
                }

                if ($allItems.Count -eq 0) {
                    Write-Verbose 'No protected items in this vault.'
                    continue
                }

                Write-Host "  $($allItems.Count) item(s) - $($vault.Name)" -ForegroundColor Green

                $itemIndex = 0
                foreach ($item in $allItems) {
                    $itemIndex++
                
                    # Parse complex names like "VM;iaasvmcontainerv2;resourcegroup;resourcename"
                    $resourceGroup = $item.FriendlyName  # Fallback
                    $resourceName = ''
                
                    if ($item.Name -match ';') {
                        $parts = $item.Name -split ';'
                        if ($parts.Count -ge 4) {
                            $resourceGroup = $parts[2]  # ResourceGroup is the 3rd part
                            $resourceName = $parts[3]   # ResourceName is the 4th part
                        }
                    }
                
                    # Fallback to SourceResourceId if parsing failed
                    if (-not $resourceName -and $item.SourceResourceId) {
                        $resourceName = ($item.SourceResourceId -split '/')[-1]
                    }
                
                    $itemDetails = "Name: $($item.Name)"
                    if ($resourceGroup) { $itemDetails += " | RG: $resourceGroup" }
                    if ($resourceName) { $itemDetails += " | ResourceName: $resourceName" }
                    Write-Verbose "[LOOP] Processing backup item $itemIndex of $($allItems.Count): $itemDetails"
                    $policyName = ''
                    if ($item.PolicyId) {
                        $policyName = ($item.PolicyId -split '/')[-1]
                    }

                    $rpCount = 0
                    try {
                        $rpParams = @{
                            VaultId     = $vault.ID
                            Item        = $item
                            ErrorAction = 'SilentlyContinue'
                        }
                        $recoveryPoints = Get-AzRecoveryServicesBackupRecoveryPoint @rpParams

                        if ($recoveryPoints) {
                            $rpCount = $recoveryPoints.Count
                        }
                    }
                    catch {
                        # Recovery point count unavailable
                    }

                    $record = [PSCustomObject]@{
                        SubscriptionName   = $sub.Name
                        ResourceGroupName  = $vault.ResourceGroupName
                        VaultName          = $vault.Name
                        BackupItemName     = $item.Name
                        ResourceName       = $resourceName
                        ResourceGroup      = $resourceGroup
                        SourceResourceId   = $item.SourceResourceId
                        ItemType           = $item.WorkloadType
                        ProtectionStatus   = $item.ProtectionStatus
                        RecoveryPointCount = $rpCount
                        BackupPolicyName   = $policyName
                        BackupPolicyLink   = if ($policyName) { "=HYPERLINK(`"#'Backup Policies'!A1`",`"🔗 $policyName`")" } else { '' }
                    }

                    $backupInventory.Add($record)
                }
            }
        }
        catch {
            Write-Warning "Failed on subscription '$($sub.Name)': $($_.Exception.Message)"
        }
    }

    Write-Verbose 'Processing policy statistics...'

    $policyStats = @{}
    foreach ($item in $backupInventory) {
        if ($item.BackupPolicyName -and $item.BackupPolicyName -ne '') {
            $key = "$($item.VaultName)|$($item.BackupPolicyName)"
            if (-not $policyStats.ContainsKey($key)) {
                $policyStats[$key] = 0
            }
            $policyStats[$key]++
        }
    }

    foreach ($policy in $backupPolicies) {
        $key = "$($policy.VaultName)|$($policy.PolicyName)"
        $resourceCount = if ($policyStats.ContainsKey($key)) { $policyStats[$key] } else { 0 }

        $hyperlinkFormula = "=HYPERLINK(`"#'Backup Inventory'!A1`",`"🔗 $($policy.PolicyName)`")"

        $policy | Add-Member -NotePropertyName 'ResourceCount' -NotePropertyValue $resourceCount -Force
        $policy | Add-Member -NotePropertyName 'PolicyNameLink' -NotePropertyValue $hyperlinkFormula -Force
    }

    Write-Verbose 'Scanning for unprotected resources...'

    $protectedResourceIds = $backupInventory | Where-Object { $_.SourceResourceId } | Select-Object -ExpandProperty SourceResourceId -Unique

    $subIndex = 0
    foreach ($sub in $subscriptions) {
        $subIndex++
        Write-Verbose "[$subIndex/$($subscriptions.Count)] $($sub.Name) - unprotected scan"

        try {
            $null = Set-AzContext -SubscriptionId $sub.Id -ErrorAction Stop

            Write-Verbose 'Scanning VMs...'
            $vms = Get-AzVM -ErrorAction SilentlyContinue
            foreach ($vm in $vms) {
                if ($vm.Id -notin $protectedResourceIds) {
                    $unprotectedResources.Add([PSCustomObject]@{
                            SubscriptionName  = $sub.Name
                            ResourceGroupName = $vm.ResourceGroupName
                            ResourceName      = $vm.Name
                            ResourceType      = 'Virtual Machine'
                            ResourceId        = $vm.Id
                            Location          = $vm.Location
                            Status            = $vm.PowerState
                        })
                }
            }

            Write-Verbose 'Scanning SQL Databases...'
            $sqlServers = Get-AzSqlServer -ErrorAction SilentlyContinue
            foreach ($server in $sqlServers) {
                $databases = Get-AzSqlDatabase -ServerName $server.ServerName -ResourceGroupName $server.ResourceGroupName -ErrorAction SilentlyContinue | Where-Object { $_.DatabaseName -ne 'master' }
                foreach ($db in $databases) {
                    if ($db.ResourceId -notin $protectedResourceIds) {
                        $unprotectedResources.Add([PSCustomObject]@{
                                SubscriptionName  = $sub.Name
                                ResourceGroupName = $db.ResourceGroupName
                                ResourceName      = "$($server.ServerName)/$($db.DatabaseName)"
                                ResourceType      = 'SQL Database'
                                ResourceId        = $db.ResourceId
                                Location          = $server.Location
                                Status            = $db.Status
                            })
                    }
                }
            }

            Write-Verbose 'Scanning Storage Accounts...'
            $storageAccounts = Get-AzStorageAccount -ErrorAction SilentlyContinue
            foreach ($storage in $storageAccounts) {
                if ($storage.Id -notin $protectedResourceIds) {
                    $unprotectedResources.Add([PSCustomObject]@{
                            SubscriptionName  = $sub.Name
                            ResourceGroupName = $storage.ResourceGroupName
                            ResourceName      = $storage.StorageAccountName
                            ResourceType      = 'Storage Account'
                            ResourceId        = $storage.Id
                            Location          = $storage.Location
                            Status            = $storage.StatusOfPrimary
                        })
                }
            }

            Write-Verbose "$($unprotectedResources.Count) unprotected resources so far"
        }
        catch {
            Write-Warning "Failed scanning unprotected resources in '$($sub.Name)': $($_.Exception.Message)"
        }
    }

    [System.Collections.Generic.List[Object]]$policyDistribution = @()
    foreach ($group in ($backupInventory | Group-Object BackupPolicyName)) {
        $policyDistribution.Add([PSCustomObject]@{
                PolicyName    = if ($group.Name) { $group.Name } else { 'No Policy' }
                ResourceCount = $group.Count
            })
    }
    $policyDistribution = $policyDistribution | Sort-Object ResourceCount -Descending

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $OutputPath = "$($env:USERPROFILE)\$now-Azure_Backup_Global_Report.xlsx"

        if ($backupInventory.Count -eq 0 -and $backupPolicies.Count -eq 0 -and $unprotectedResources.Count -eq 0) {
            Write-Warning 'No data found. The Excel file will not be created.'
        }
        else {
            Write-Host "Exporting to: $OutputPath" -ForegroundColor Cyan

            if ($backupInventory.Count -gt 0) {
                $backupInventory | Export-Excel -Path $OutputPath -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName 'Backup Inventory'
                Write-Verbose "Backup inventory: $($backupInventory.Count) items exported"
            }

            if ($backupPolicies.Count -gt 0) {
                $backupPolicies | Export-Excel -Path $OutputPath -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName 'Backup Policies'
                Write-Verbose "Backup policies: $($backupPolicies.Count) exported"
            }

            if ($unprotectedResources.Count -gt 0) {
                $unprotectedResources | Export-Excel -Path $OutputPath -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName 'Unprotected Resources'
                Write-Verbose "Unprotected resources: $($unprotectedResources.Count) exported"
            }

            if ($policyDistribution) {
                $policyDistribution | Export-Excel -Path $OutputPath -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName 'Policy Distribution'
            }

            Write-Host "Report saved: $OutputPath" -ForegroundColor Green
            Write-Host "$($backupInventory.Count) backed up | $($backupPolicies.Count) policies | $($unprotectedResources.Count) unprotected" -ForegroundColor Cyan
        }
    }
    else {
        return @{
            BackupInventory      = $backupInventory
            BackupPolicies       = $backupPolicies
            UnprotectedResources = $unprotectedResources
            PolicyDistribution   = $policyDistribution
        }
    }
}