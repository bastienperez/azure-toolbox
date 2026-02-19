<#
    .SYNOPSIS
    Generates Azure topology diagrams for all subscriptions and resource groups.
    
    .PARAMETER Theme
    Visual theme for the diagram output.
    
    .PARAMETER OutputFilePath
    Directory path where diagram files will be saved.
    
    .PARAMETER LabelVerbosity
    Level of detail in resource labels (1=minimal, 2=standard, 3=detailed).
    
    .PARAMETER CategoryDepth
    Depth of resource categorization (1=basic, 2=standard, 3=detailed).
    
    .PARAMETER OutputFormat
    Image format for the generated diagrams.
    
    .PARAMETER Direction
    Layout direction for the diagram visualization.

    .NOTES
    This function requires the AzViz module and GraphViz to be installed.
    First, ensure you have GraphViz installed on your system:
    - Windows: winget install graphviz
    - Mac: brew install graphviz
    - Linux: sudo apt install graphviz / sudo yum install graphviz

    Then, install the AzViz PowerShell module from the PSGallery repository:
    Install-Module -Name AzViz -Scope CurrentUser -Repository PSGallery

    Script inspired by: https://www.reddit.com/r/PowerShell/comments/mmxq6a/powershell_module_to_visualize_and_document_azure/
#>
    
function Get-AzTopology {

    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory,
            HelpMessage = 'Select the visual theme: light, dark, or neon'
        )]
        [ValidateSet('light', 'dark', 'neon')]
        [string]$Theme,

        [Parameter(
            HelpMessage = 'Directory path where diagram files will be saved'
        )]
        [ValidateScript({
                if (Test-Path -Path $_ -IsValid) { $true }
                else { throw 'Invalid path format' }
            })]
        [string]$OutputFilePath = 'C:\temp',

        [Parameter(
            HelpMessage = 'Label verbosity level (1-3, where 3 is most detailed)'
        )]
        [ValidateRange(1, 3)]
        [int]$LabelVerbosity = 3,

        [Parameter(
            HelpMessage = 'Resource categorization depth (1-3, where 3 is most detailed)'
        )]
        [ValidateRange(1, 3)]
        [int]$CategoryDepth = 3,

        [Parameter(
            HelpMessage = 'Output image format for diagrams'
        )]
        [ValidateSet('png', 'svg')]
        [string]$OutputFormat = 'png',

        [Parameter(
            HelpMessage = 'Layout direction for diagram visualization'
        )]
        [ValidateSet('left-to-right', 'top-to-bottom')]
        [string]$Direction = 'left-to-right'
    )

    # Initialize variables
    $dateStamp = Get-Date -Format 'MM-dd-yyyy'
    $extension = if ($OutputFormat -eq 'svg') { '.svg' } else { '.png' }

    try {
        # Get all available subscriptions
        $subscriptions = Get-AzSubscription
        
        if (-not $subscriptions) {
            Write-Warning 'No Azure subscriptions found'
            return
        }

        Write-Output "Found $($subscriptions.Count) subscription(s) to process"

        # Process each subscription
        foreach ($subscription in $subscriptions) {
            try {
                # Set the current Azure context
                $null = Get-AzSubscription -SubscriptionName $subscription.Name | Set-AzContext
                Write-Output "Processing subscription: $($subscription.Name)"
                
                # Create subscription directory
                $subscriptionName = $subscription.Name
                $subscriptionPath = Join-Path -Path $OutputFilePath -ChildPath $dateStamp -AdditionalChildPath $subscriptionName
                $null = New-Item -Path $subscriptionPath -ItemType Directory -Force
                
                # Get all resource groups in the current subscription
                $resourceGroups = Get-AzResourceGroup
                
                if (-not $resourceGroups) {
                    Write-Warning "No resource groups found in subscription: $subscriptionName"
                    continue
                }

                Write-Output "Found $($resourceGroups.Count) resource group(s) in $subscriptionName"

                # Process each resource group
                foreach ($resourceGroup in $resourceGroups) {
                    try {
                        $resourceGroupName = $resourceGroup.ResourceGroupName
                        $fileName = $resourceGroupName + $extension
                        $outputPath = Join-Path -Path $subscriptionPath -ChildPath $fileName

                        Write-Output "Generating diagram for resource group: $resourceGroupName"

                        # Configure parameters for Export-AzViz using splatting
                        $exportParams = @{
                            ResourceGroup  = $resourceGroupName
                            OutputFilePath = $outputPath
                            Theme          = $Theme
                            OutputFormat   = $OutputFormat
                            CategoryDepth  = $CategoryDepth
                            Direction      = $Direction
                            LabelVerbosity = $LabelVerbosity
                        }

                        # Generate the topology diagram
                        Export-AzViz @exportParams
                    }
                    catch {
                        Write-Error "Failed to process resource group '$resourceGroupName': $_"
                        continue
                    }
                }
            }
            catch {
                Write-Error "Failed to process subscription '$($subscription.Name)': $_"
                continue
            }
        }
    }
    catch {
        Write-Error "An error occurred while processing Azure topology: $_"
        return
    }

    Write-Output "Azure topology diagrams generated successfully in: $OutputFilePath\$dateStamp"
}