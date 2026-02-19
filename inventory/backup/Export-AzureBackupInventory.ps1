#Requires -Modules Az.Accounts, Az.RecoveryServices, Az.Compute, Az.Sql, Az.Storage, ImportExcel

<#
.SYNOPSIS
    Generates a global backup inventory and identifies unprotected resources across all accessible Azure subscriptions.

.DESCRIPTION
    Iterates through every Azure subscription, queries all Recovery Services Vaults,
    extracts Protected Items and identifies resources that are not protected by backup.
    Exports results into an Excel report with multiple worksheets.

.OUTPUTS
    Azure_Backup_Global_Report.xlsx with worksheets for backup inventory, policies, and unprotected resources.
#>

function Export-AzureBackupInventory {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$OutputPath
    )

    $ErrorActionPreference = 'Stop'

    # Generate timestamped filename if not provided
    if (-not $OutputPath) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $OutputPath = "$($env:USERPROFILE)\$now-Azure_Backup_Global_Report.xlsx"
    }

    # ---------------------------------------------------------------------------
    # 1. Ensure we have an active Azure session
    # ---------------------------------------------------------------------------
    Write-Host '=== Azure Backup Global Inventory ===' -ForegroundColor Cyan
    Write-Host ''

    $context = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Write-Host '[INFO] No active Azure session found. Launching interactive login...' -ForegroundColor Yellow
        Connect-AzAccount | Out-Null
    }

    # ---------------------------------------------------------------------------
    # 2. Retrieve all accessible subscriptions
    # ---------------------------------------------------------------------------
    Write-Host '[INFO] Retrieving subscription list...' -ForegroundColor Yellow
    $subscriptions = Get-AzSubscription -ErrorAction Stop | Where-Object { $_.State -eq 'Enabled' }
    Write-Host "[INFO] Found $($subscriptions.Count) enabled subscription(s)." -ForegroundColor Green
    Write-Host ''

    # ---------------------------------------------------------------------------
    # 3. Iterate subscriptions and collect backup items
    # ---------------------------------------------------------------------------
    [System.Collections.Generic.List[PSCustomObject]]$backupInventory = @()
    [System.Collections.Generic.List[PSCustomObject]]$backupPolicies = @()
    [System.Collections.Generic.List[PSCustomObject]]$unprotectedResources = @()
    $subIndex = 0

    foreach ($sub in $subscriptions) {
        $subIndex++
        Write-Verbose "[LOOP] Starting subscription loop $subIndex of $($subscriptions.Count): $($sub.Name)"
        Write-Host "[PROGRESS] Scanning subscription $subIndex/$($subscriptions.Count)" -ForegroundColor Magenta

        try {
            Write-Host "[$subIndex/$($subscriptions.Count)] Processing subscription: $($sub.Name) ($($sub.Id))" -ForegroundColor Cyan

            # Set the working subscription context
            Set-AzContext -SubscriptionId $sub.Id -ErrorAction Stop | Out-Null

            # Discover all Recovery Services Vaults in this subscription
            $vaults = Get-AzRecoveryServicesVault -ErrorAction Stop

            if ($vaults.Count -eq 0) {
                Write-Host '  -> No Recovery Services Vaults found. Skipping.' -ForegroundColor DarkGray
                continue
            }

            Write-Host "  -> Found $($vaults.Count) vault(s)." -ForegroundColor Green

            $vaultIndex = 0
            foreach ($vault in $vaults) {
                $vaultIndex++
                Write-Verbose "[LOOP] Starting vault loop $vaultIndex of $($vaults.Count): $($vault.Name)"
                Write-Host "     Vault: $($vault.Name) (RG: $($vault.ResourceGroupName))" -ForegroundColor White

                # Set vault context for backup cmdlets
                Set-AzRecoveryServicesVaultContext -Vault $vault -ErrorAction Stop
                Write-Host '       Scanning containers...' -ForegroundColor DarkGray

                # Get backup policies for this vault
                Write-Host '       Retrieving backup policies...' -ForegroundColor DarkGray
                try {
                    $policies = Get-AzRecoveryServicesBackupProtectionPolicy -VaultId $vault.ID -ErrorAction SilentlyContinue
                    if ($policies) {
                        Write-Host "         -> Found $($policies.Count) backup policy(ies)" -ForegroundColor DarkGray
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
                                $scheduleFrequency = $policy.SchedulePolicy.ScheduleRunFrequency
                                if ($policy.SchedulePolicy.ScheduleRunTimes) {
                                    $backupTimes = ($policy.SchedulePolicy.ScheduleRunTimes | ForEach-Object { $_.ToString('HH:mm') }) -join ', '
                                }
                                if ($policy.SchedulePolicy.ScheduleRunDays) {
                                    $scheduleDays = ($policy.SchedulePolicy.ScheduleRunDays) -join ', '
                                }
                                if ($policy.SchedulePolicy.ScheduleWindowStartTime) {
                                    $windowStartTime = $policy.SchedulePolicy.ScheduleWindowStartTime.ToString('HH:mm')
                                }
                                if ($policy.SchedulePolicy.ScheduleWindowDuration) {
                                    $windowDuration = "$($policy.SchedulePolicy.ScheduleWindowDuration) hours"
                                }
                                if ($policy.SchedulePolicy.ScheduleRunTimeZone) {
                                    $timeZone = $policy.SchedulePolicy.ScheduleRunTimeZone
                                }
                                if ($policy.SchedulePolicy.ScheduleInterval) {
                                    $interval = "$($policy.SchedulePolicy.ScheduleInterval) hours"
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
                    Write-Host "         [WARN] Could not retrieve backup policies: $($_.Exception.Message)" -ForegroundColor DarkYellow
                }

                # Query containers for all workload types we care about
                $containerTypes = @(
                    'AzureVM'
                    'AzureSQL'
                    'AzureStorage'
                    'AzureVMAppContainer'
                    'Windows'
                )

                [System.Collections.Generic.List[PSCustomObject]]$allItems = @()

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
                            Write-Host "         -> No containers found for type '$containerType'" -ForegroundColor DarkGray
                            continue 
                        }
                    
                        Write-Host "         -> Found $($containers.Count) container(s) for type '$containerType'" -ForegroundColor DarkGray

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
                        Write-Host "       [WARN] Could not query container type '$containerType': $($_.Exception.Message)" -ForegroundColor DarkYellow
                    }
                }

                # Also query workload-based items (SQL in VM, SAP HANA) directly
                Write-Host '       Scanning workload-based items directly...' -ForegroundColor DarkGray
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
                            Write-Host "         -> Found $($wlItems.Count) workload item(s) for type '$wlType'" -ForegroundColor DarkGray
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
                    Write-Host '       No protected items found in this vault.' -ForegroundColor DarkGray
                    continue
                }

                Write-Host "       Found $($allItems.Count) protected item(s)." -ForegroundColor Green
                Write-Host '       Processing backup items details...' -ForegroundColor DarkGray

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
                    # Resolve the backup policy name
                    $policyName = ''
                    if ($item.PolicyId) {
                        $policyName = ($item.PolicyId -split '/')[-1]
                    }

                    # Attempt to get recovery point count
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
            Write-Host "  [ERROR] Failed to process subscription '$($sub.Name)': $($_.Exception.Message)" -ForegroundColor Red
            Write-Host '  Continuing with next subscription...' -ForegroundColor Yellow
        }
    }

    # ---------------------------------------------------------------------------
    # 4. Find unprotected resources
    # ---------------------------------------------------------------------------
    Write-Host ''
    Write-Host '[INFO] Processing policy statistics and creating links...' -ForegroundColor Yellow
    
    # Count resources per policy
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
    
    # Add resource count and hyperlinks to policies
    for ($i = 0; $i -lt $backupPolicies.Count; $i++) {
        $policy = $backupPolicies[$i]
        $key = "$($policy.VaultName)|$($policy.PolicyName)"
        $resourceCount = if ($policyStats.ContainsKey($key)) { $policyStats[$key] } else { 0 }
        
        # Create hyperlink to Backup Inventory sheet with visual indicator
        $hyperlinkFormula = "=HYPERLINK(`"#'Backup Inventory'!A1`",`"🔗 $($policy.PolicyName)`")"
        
        $policy | Add-Member -NotePropertyName 'ResourceCount' -NotePropertyValue $resourceCount -Force
        $policy | Add-Member -NotePropertyName 'PolicyNameLink' -NotePropertyValue $hyperlinkFormula -Force
    }
    
    Write-Host '[INFO] Scanning for unprotected resources...' -ForegroundColor Yellow
    
    # Get list of protected resource IDs for comparison
    $protectedResourceIds = $backupInventory | Where-Object { $_.SourceResourceId } | Select-Object -ExpandProperty SourceResourceId -Unique
    
    $subIndex = 0
    foreach ($sub in $subscriptions) {
        $subIndex++
        Write-Host "[PROGRESS] Scanning unprotected resources in subscription $subIndex/$($subscriptions.Count)" -ForegroundColor Magenta
        
        try {
            Set-AzContext -SubscriptionId $sub.Id -ErrorAction Stop | Out-Null
            
            # Get all VMs
            Write-Host '  -> Scanning VMs...' -ForegroundColor DarkGray
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
            
            # Get all SQL Databases
            Write-Host '  -> Scanning SQL Databases...' -ForegroundColor DarkGray
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
            
            # Get all Storage Accounts
            Write-Host '  -> Scanning Storage Accounts...' -ForegroundColor DarkGray
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
            
            Write-Host "  -> Found $($unprotectedResources.Count) unprotected resources so far" -ForegroundColor DarkGray
        }
        catch {
            Write-Host "  [ERROR] Failed to scan unprotected resources in '$($sub.Name)': $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # ---------------------------------------------------------------------------
    # 5. Export all results
    # ---------------------------------------------------------------------------
    Write-Host ''
    
    # Create policy distribution for export
    $policyDistribution = $backupInventory | Group-Object BackupPolicyName | ForEach-Object {
        [PSCustomObject]@{
            PolicyName    = if ($_.Name) { $_.Name } else { 'No Policy' }
            ResourceCount = $_.Count
        }
    } | Sort-Object ResourceCount -Descending
    
    if ($backupInventory.Count -eq 0 -and $backupPolicies.Count -eq 0 -and $unprotectedResources.Count -eq 0) {
        Write-Host '[WARNING] No data found across any subscription.' -ForegroundColor Yellow
        Write-Host 'The Excel file will not be created.'
    }
    else {
        Write-Host -ForegroundColor Cyan "Exporting backup information to Excel file: $OutputPath"
        
        # Export inventory data
        if ($backupInventory.Count -gt 0) {
            $backupInventory | Export-Excel -Path $OutputPath -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName 'Backup Inventory'
            Write-Host "[SUCCESS] Backup inventory exported: $($backupInventory.Count) items" -ForegroundColor Green
        }

        # Export policies data to a separate worksheet
        if ($backupPolicies.Count -gt 0) {
            $backupPolicies | Export-Excel -Path $OutputPath -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName 'Backup Policies'
            Write-Host "[SUCCESS] Backup policies exported: $($backupPolicies.Count) policies" -ForegroundColor Green
        }

        # Export unprotected resources to a separate worksheet
        if ($unprotectedResources.Count -gt 0) {
            $unprotectedResources | Export-Excel -Path $OutputPath -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName 'Unprotected Resources'
            Write-Host "[SUCCESS] Unprotected resources exported: $($unprotectedResources.Count) resources" -ForegroundColor Green
        }

        # Export policy distribution
        if ($policyDistribution) {
            $policyDistribution | Export-Excel -Path $OutputPath -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -WorksheetName 'Policy Distribution'
            Write-Host '[SUCCESS] Policy distribution exported' -ForegroundColor Green
        }

        Write-Host "[SUCCESS] Report exported to: $OutputPath" -ForegroundColor Green
        Write-Host "[INFO] Total items: $($backupInventory.Count) backed up, $($backupPolicies.Count) policies, $($unprotectedResources.Count) unprotected" -ForegroundColor Cyan
    }

    Write-Host ''
    Write-Host '=== Inventory complete ===' -ForegroundColor Cyan
}