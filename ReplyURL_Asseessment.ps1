# Set the execution policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Check the preferred Graph API version
$Choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$Choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))
$Choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
$ExploreVersion = $Host.UI.PromptForChoice("Graph API Version", "Do you like to use Beta version of Graph API?", $Choices, 0)

# Install needed modules
if ($ExploreVersion) {
    $ExploreVersion = ".Beta"
} else {
    $ExploreVersion = ""
}

if (Get-Module -ListAvailable -Name Microsoft.Graph$ExploreVersion) {
    Write-Host "Microsoft.Graph${$ExploreVersion} is already installed"
} else {
    Install-Module Microsoft.Graph$ExploreVersion -Scope CurrentUser -AllowClobber -Repository PSGallery -Force
    Import-Module Microsoft.Graph$ExploreVersion
}

#######################################################################
#######################################################################
#######################################################################

# Connect to the Microsoft Graph with the given scope
# $Scopes = "User.Read.All", "Group.ReadWrite.All"
if ($Scope) {
    Connect-MgGraph -Scopes $Scopes -NoWelcome
} else {
    Connect-MgGraph -NoWelcome
}

# Get the current datetime and set the output folder
$Date = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$OutputFolder = "$HOME\Downloads\GraphAPI"


#######################################################################
# Basic functions
#######################################################################

# Get the list Enterprise Apps
$eApps = Get-MgApplication -All | Select-Object Id, DisplayName, @{Name="HomePageUrl"; Expression={$_.Web.HomePageUrl}}, @{Name="LogoutUrl"; Expression={$_.Web.LogoutUrl}}, @{Name="RedirectUris"; Expression={$_.Web.RedirectUris}}


#######################################################################
# Get the list of conditional access policies (Apps)
#######################################################################
$Apps = Get-MgApplication -All | Select-Object Id, DisplayName, @{Name="HomePageUrl"; Expression={$_.Web.HomePageUrl}}, @{Name="LogoutUrl"; Expression={$_.Web.LogoutUrl}}, @{Name="RedirectUris"; Expression={$_.Web.RedirectUris}}

$Apps = $Apps | ForEach-Object {
    $Id = $_.Id
    $DisplayName = $_.DisplayName
    $HomePageUrl = $_.HomePageUrl
    $LogoutUrl = $_.LogoutUrl
    $RedirectUris = $_.RedirectUris

    # If no Redirect URIs, create a new row with a specific message
    if ($RedirectUris -eq $null -or $RedirectUris.Count -eq 0) {
        [PSCustomObject]@{
            Id            = $Id
            DisplayName   = $DisplayName
            HomePageUrl   = $HomePageUrl
            LogoutUrl     = $LogoutUrl
            RedirectUri   = "No Redirect URI"
            Issue         = "N/A"
        }
    } else {

        # For each Redirect URI, create a new row in the output with issue detection
        $RedirectUris | ForEach-Object {
            # Initialize the issue message
            $Issue = ""
            $StatusCode = "N/A"

            # Check for HTTP (non-secure)
            if ($_ -like "http://*") {
                $Issue = "HTTP"
            }

            # Check for simple hostnames/computer names (e.g., http://dummy, http://computername)
            if ($_ -match "[a-zA-Z0-9\-]+$") {
                if ($Issue) { $Issue += "; " }
                $Issue += "Local Hostname"
            }

            # Check for localhost (development only)
            if ($_ -match "(localhost|127\.\d{1,3}\.\d{1,3}\.\d{1,3}|::1)") {
                if ($Issue) { $Issue += "; " }
                $Issue += "Localhost"
            }

            # Check for wildcard (subdomain issues)
            if ($_ -like "`*") {
                if ($Issue) { $Issue += "; " }
                $Issue += "Wildcard"
            }

            # If no issues, set to "None"
            if (-not $Issue) {
                $Issue = "None"
            }

            # Get the HTTP header Status Code
            try {
                $response = Invoke-WebRequest -Uri $_ -Method Head -ErrorAction Stop
                $StatusCode = $response.StatusCode
            } catch {
                $StatusCode = "Error: $($_.Exception.Message)"
            }

            # Output the results with the Issue column
            [PSCustomObject]@{
                Id            = $Id
                DisplayName   = $DisplayName
                HomePageUrl   = $HomePageUrl
                LogoutUrl     = $LogoutUrl
                RedirectUri   = $_
                Issue         = $Issue
                StatusCode    = $StatusCode
            }
        }
    }
}

# Export to CSV
$Apps | Export-Csv -Path "$OutputFolder\output_Apps_$Date.csv" -NoTypeInformation