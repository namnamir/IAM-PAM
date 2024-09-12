# Set the execution policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Install needed modules
# Install-Module Microsoft.Graph -Scope CurrentUser -Repository PSGallery -Force
Install-Module Microsoft.Graph.Beta -Scope CurrentUser -Repository PSGallery -Force

# Connect to the Microsoft Graph with the given scope
$Scopes = "User.Read.All", "Group.ReadWrite.All"
Connect-MgGraph -Scopes $Scopes

# Get the current datetime and set the output folder
$Date = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$OutputFolder = "$HOME\Downloads"

# Get the list of all users with the given filter and objects
$Filter = "startswith(userPrincipalName,'adm.') or startswith(userPrincipalName,'appadm.')"
$Objects = "UserPrincipalName", "id"
$OutputPath = "$OutputFolder\output_$Date.csv"

Get-MgUser -All -Filter $Filter | Select-Object $Objects | Export-Csv -Path $OutputPath -Encoding UTF8

