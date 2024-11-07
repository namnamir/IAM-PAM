# Set the execution policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Check the preferred Graph API version
$Choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$Choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))
$Choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
$ExploreVersion = $Host.UI.PromptForChoice("Graph API Version", "Do you like to use Beta version of Graph API?", $Choices, 1)

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
}

Import-Module Microsoft.Graph$ExploreVersion

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

$builtinNames = @{
    # Applications
    "23523755-3a2b-41ca-9315-f81f3f566a95" = "ACOM Azure Website"
    "74658136-14ec-4630-ad9b-26e160ff0fc6" = "ADIbizaUX"
    "69893ee3-dd10-4b1c-832d-4870354be3d8" = "AEM-DualAuth"
    "7ab7862c-4c57-491e-8a45-d52a7e023983" = "App Service"
    "0cb7b9ec-5336-483b-bc31-b15b5788de71" = "ASM Campaign Servicing"
    "7b7531ad-5926-4f2d-8a1d-38495ad33e17" = "Azure Advanced Threat Protection"
    "e9f49c6b-5ce5-44c8-925d-015017e9f7ad" = "Azure Data Lake"
    "835b2a73-6e10-4aa5-a979-21dfda45231c" = "Azure Lab Services Portal"
    "c44b4083-3bb0-49c1-b47d-974e53cbdf3c" = "Azure Portal"
    "022907d3-0f1b-48f7-badc-1ba6abab6d66" = "Azure SQL Database"
    "37182072-3c9c-4f6a-a4b3-b3f91cacffce" = "AzureSupportCenter"
    "9ea1ad79-fdb6-4f9a-8bc3-2b70f96e34c7" = "Bing"
    "20a11fe0-faa8-4df5-baf2-f965f8f9972e" = "ContactsInferencingEmailProcessor"
    "bb2a2e3a-c5e7-4f0a-88e0-8e01fd3fc1f4" = "CPIM Service"
    "e64aa8bc-8eb4-40e2-898b-cf261a25954f" = "CRM Power BI Integration"
    "00000007-0000-0000-c000-000000000000" = "Dataverse"
    "60c8bde5-3167-4f92-8fdb-059f6176dc0f" = "Enterprise Roaming and Backup"
    "497effe9-df71-4043-a8bb-14cf78c4b63b" = "Exchange Admin Center"
    "f5eaa862-7f08-448c-9c4e-f4047d4d4521" = "FindTime"
    "b669c6ea-1adf-453f-b8bc-6d526592b419" = "Focused Inbox"
    "c35cb2ba-f88b-4d15-aa9d-37bd443522e1" = "GroupsRemoteApiRestClient"
    "d9b8ec3a-1e4e-4e08-b3c2-5baf00c0fcb0" = "HxService"
    "a57aca87-cbc0-4f3c-8b9e-dc095fdc8978" = "IAM Supportability"
    "16aeb910-ce68-41d1-9ac3-9e1673ac9575" = "IrisSelectionFrontDoor"
    "d73f4b35-55c9-48c7-8b10-651f6f2acb2e" = "MCAPI Authorization Prod"
    "944f0bd1-117b-4b1c-af26-804ed95e767e" = "Media Analysis and Transformation Service"
    "0cd196ee-71bf-4fd6-a57c-b491ffd4fb1e" = "Media Analysis and Transformation Service"
    "80ccca67-54bd-44ab-8625-4b79c4dc7775" = "Microsoft 365 Security and Compliance Center"
    "ee272b19-4411-433f-8f28-5c13cb6fd407" = "Microsoft 365 Support Service"
    "0000000c-0000-0000-c000-000000000000" = "Microsoft App Access Panel"
    "65d91a3d-ab74-42e6-8a2f-0add61688c74" = "Microsoft Approval Management"
    "38049638-cc2c-4cde-abe4-4479d721ed44" = "Microsoft Approval Management"
    "29d9ed98-a469-4536-ade2-f981bc1d605e" = "Microsoft Authentication Broker"
    "04b07795-8ddb-461a-bbee-02f9e1bf7b46" = "Microsoft Azure CLI"
    "1950a258-227b-4e31-a9cf-717495945fc2" = "Microsoft Azure PowerShell"
    "0000001a-0000-0000-c000-000000000000" = "MicrosoftAzureActiveAuthn"
    "cf36b471-5b44-428c-9ce7-313bf84528de" = "Microsoft Bing Search"
    "2d7f3606-b07d-41d1-b9d2-0d0c9296a6e8" = "Microsoft Bing Search for Microsoft Edge"
    "1786c5ed-9644-47b2-8aa0-7201292175b6" = "Microsoft Bing Default Search Engine"
    "3090ab82-f1c1-4cdf-af2c-5d7a6f3e2cc7" = "Microsoft Defender for Cloud Apps"
    "60ca1954-583c-4d1f-86de-39d835f3e452" = "Microsoft Defender for Identity (formerly Radius Aad Syncer)"
    "18fbca16-2224-45f6-85b0-f7bf2b39b3f3" = "Microsoft Docs"
    "00000015-0000-0000-c000-000000000000" = "Microsoft Dynamics ERP"
    "6253bca8-faf2-4587-8f2f-b056d80998a7" = "Microsoft Edge Insider Addons Prod"
    "99b904fd-a1fe-455c-b86c-2f9fb1da7687" = "Microsoft Exchange ForwardSync"
    "00000007-0000-0ff1-ce00-000000000000" = "Microsoft Exchange Online Protection"
    "51be292c-a17e-4f17-9a7e-4b661fb16dd2" = "Microsoft Exchange ProtectedServiceHost"
    "fb78d390-0c51-40cd-8e17-fdbfab77341b" = "Microsoft Exchange REST API Based Powershell"
    "47629505-c2b6-4a80-adb1-9b3a3d233b7b" = "Microsoft Exchange Web Services"
    "c9a559d2-7aab-4f13-a6ed-e7e9c52aec87" = "Microsoft Forms"
    "00000003-0000-0000-c000-000000000000" = "Microsoft Graph"
    "74bcdadc-2fdc-4bb3-8459-76d06952a0e9" = "Microsoft Intune Web Company Portal"
    "fc0f3af4-6835-4174-b806-f7db311fd2f3" = "Microsoft Intune Windows Agent"
    "d3590ed6-52b3-4102-aeff-aad2292ab01c" = "Microsoft Office"
    "00000006-0000-0ff1-ce00-000000000000" = "Microsoft Office 365 Portal"
    "67e3df25-268a-4324-a550-0de1c7f97287" = "Microsoft Office Web Apps Service"
    "d176f6e7-38e5-40c9-8a78-3998aab820e7" = "Microsoft Online Syndication Partner Portal"
    "93625bc8-bfe2-437a-97e0-3d0060024faa" = "Microsoft password reset service"
    "871c010f-5e61-4fb1-83ac-98610a7e9110" = "Microsoft Power BI"
    "28b567f6-162c-4f54-99a0-6887f387bbcc" = "Microsoft Storefronts"
    "cf53fce8-def6-4aeb-8d30-b158e7b1cf83" = "Microsoft Stream Portal"
    "98db8bd6-0cc0-4e67-9de5-f187f1cd1b41" = "Microsoft Substrate Management"
    "fdf9885b-dd37-42bf-82e5-c3129ef5a302" = "Microsoft Support"
    "1fec8e78-bce4-4aaf-ab1b-5451cc387264" = "Microsoft Teams"
    "cc15fd57-2c6c-4117-a88c-83b1d56b4bbe" = "Microsoft Teams Services"
    "5e3ce6c0-2b1f-4285-8d4b-75ee78787346" = "Microsoft Teams Web Client"
    "95de633a-083e-42f5-b444-a4295d8e9314" = "Microsoft Whiteboard Services"
    "dfe74da8-9279-44ec-8fb2-2aed9e1c73d0" = "O365 SkypeSpaces Ingestion Service"
    "4345a7b9-9a63-4910-a426-35363201d503" = "O365 Suite UX"
    "00000002-0000-0ff1-ce00-000000000000" = "Office 365 Exchange Online"
    "00b41c95-dab0-4487-9791-b9d2c32c80f2" = "Office 365 Management"
    "66a88757-258c-4c72-893c-3e8bed4d6899" = "Office 365 Search Service"
    "00000003-0000-0ff1-ce00-000000000000" = "Office 365 SharePoint Online"
    "94c63fef-13a3-47bc-8074-75af8c65887a" = "Office Delve"
    "93d53678-613d-4013-afc1-62e9e444a0a5" = "Office Online Add-in SSO"
    "2abdc806-e091-4495-9b10-b04d93c3f040" = "Office Online Client Microsoft Entra ID- Augmentation Loop"
    "b23dd4db-9142-4734-867f-3577f640ad0c" = "Office Online Client Microsoft Entra ID- Loki"
    "17d5e35f-655b-4fb0-8ae6-86356e9a49f5" = "Office Online Client Microsoft Entra ID- Maker"
    "b6e69c34-5f1f-4c34-8cdf-7fea120b8670" = "Office Online Client MSA- Loki"
    "243c63a3-247d-41c5-9d83-7788c43f1c43" = "Office Online Core SSO"
    "a9b49b65-0a12-430b-9540-c80b3332c127" = "Office Online Search"
    "4b233688-031c-404b-9a80-a4f3f2351f90" = "Office.com"
    "89bee1f7-5e6e-4d8a-9f3d-ecd601259da7" = "Office365 Shell WCSS-Client"
    "0f698dd4-f011-4d23-a33e-b36416dcb1e6" = "OfficeClientService"
    "4765445b-32c6-49b0-83e6-1d93765276ca" = "OfficeHome"
    "4d5c2d63-cf83-4365-853c-925fd1a64357" = "OfficeShredderWacClient"
    "62256cef-54c0-4cb4-bcac-4c67989bdc40" = "OMSOctopiPROD"
    "ab9b8c07-8f02-4f72-87fa-80105867a763" = "OneDrive SyncEngine"
    "2d4d3d8e-2be3-4bef-9f87-7875a61c29de" = "OneNote"
    "27922004-5251-4030-b22d-91ecd9a37ea4" = "Outlook Mobile"
    "a3475900-ccec-4a69-98f5-a65cd5dc5306" = "Partner Customer Delegated Admin Offline Processor"
    "bdd48c81-3a58-4ea9-849c-ebea7f6b6360" = "Password Breach Authenticator"
    "35d54a08-36c9-4847-9018-93934c62740c" = "PeoplePredictions"
    "00000009-0000-0000-c000-000000000000" = "Power BI Service"
    "ae8e128e-080f-4086-b0e3-4c19301ada69" = "Scheduling"
    "ffcb16e8-f789-467c-8ce9-f826a080d987" = "SharedWithMe"
    "08e18876-6177-487e-b8b5-cf950c1e598c" = "SharePoint Online Web Client Extensibility"
    "b4bddae8-ab25-483e-8670-df09b9f1d0ea" = "Signup"
    "00000004-0000-0ff1-ce00-000000000000" = "Skype for Business Online"
    "61109738-7d2b-4a0b-9fe3-660b1ff83505" = "SpoolsProvisioning"
    "91ca2ca5-3b3e-41dd-ab65-809fa3dffffa" = "Sticky Notes API"
    "13937bba-652e-4c46-b222-3003f4d1ff97" = "Substrate Context Service"
    "26abc9a8-24f0-4b11-8234-e86ede698878" = "SubstrateDirectoryEventProcessor"
    "a970bac6-63fe-4ec5-8884-8536862c42d4" = "Substrate Search Settings Management Service"
    "905fcf26-4eb7-48a0-9ff0-8dcc7194b5ba" = "Sway"
    "97cb1f73-50df-47d1-8fb0-0271f2728514" = "Transcript Ingestion"
    "268761a2-03f3-40df-8a8b-c3db24145b6b" = "Universal Store Native Client"
    "00000005-0000-0ff1-ce00-000000000000" = "Viva Engage (formerly Yammer)"
    "3c896ded-22c5-450f-91f6-3d1ef0848f6e" = "WeveEngine"
    "00000002-0000-0000-c000-000000000000" = "Windows Azure Active Directory"
    "8edd93e1-2103-40b4-bd70-6e34e586362d" = "Windows Azure Security Resource Provider"
    "797f4846-ba00-4fd7-ba43-dac1f8f63013" = "Windows Azure Service Management API"
    "a3b79187-70b2-4139-83f9-6016c58cd27b" = "WindowsDefenderATP Portal"
    "26a7ee05-5602-4d76-a7ba-eae8b7b67941" = "Windows Search"
    "1b3c667f-cde3-4090-b60b-3d2abd0117f0" = "Windows Spotlight"
    "45a330b1-b1ec-4cc1-9161-9f03992aa49f" = "Windows Store for Business"
    "c1c74fed-04c9-4704-80dc-9f79a2e515cb" = "Yammer Web"
    "e1ef36fd-b883-4dbf-97f0-9ece4b576fc6" = "Yammer Web Embed"
    "de8bc8b5-d9f9-48b1-a8ad-b748da725064" = "Graph Explorer"
    "14d82eec-204b-4c2f-b7e8-296a70dab67e" = "Microsoft Graph Command Line Tools"
    "7ae974c5-1af7-4923-af3a-fb1fd14dcb7e" = "OutlookUserSettingsConsumer"
    "5572c4c0-d078-44ce-b81c-6cbf8d3ed39e" = "Vortex [wsfed enabled]"
    # Enterprise applications
    "39e6ea5b-4aa4-4df2-808b-b6b5fb8ada6f" = "Dynamics Provision"
    "7df0a125-d3be-4c96-aa54-591f83ff541c" = "Microsoft Flow"
    "475226c6-020e-4fb2-8a90-7a972cbfc1d4" = "Microsoft PowerApps"
    "637fcc9f-4a9b-4aaa-8713-a2a3cfda1505" = "Dynamics CRM Online Administration"
    "f53895d3-095d-408f-8e93-8f94b391404e" = "Project Online"
    # Roles
    "9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3" = "Application Administrator"
    "cf1c38e5-3621-4004-a7cb-879624dced7c" = "Application Developer"
    "9c6df0f2-1e7c-4dc3-b195-66dfbd24aa8f" = "Attack Payload Author"
    "c430b396-e693-46cc-96f3-db01bf8bb62a" = "Attack Simulation Administrator"
    "58a13ea3-c632-46ae-9ee0-9c0d43cd7f3d" = "Attribute Assignment Administrator"
    "ffd52fa5-98dc-465c-991d-fc073eb59f8f" = "Attribute Assignment Reader"
    "8424c6f0-a189-499e-bbd0-26c1753c96d4" = "Attribute Definition Administrator"
    "1d336d2c-4ae8-42ef-9711-b3604ce3fc2c" = "Attribute Definition Reader"
    "5b784334-f94b-471a-a387-e7219fc49ca2" = "Attribute Log Administrator"
    "9c99539d-8186-4804-835f-fd51ef9e2dcd" = "Attribute Log Reader"
    "c4e39bd9-1100-46d3-8c65-fb160da0071f" = "Authentication Administrator"
    "25a516ed-2fa0-40ea-a2d0-12923a21473a" = "Authentication Extensibility Administrator"
    "0526716b-113d-4c15-b2c8-68e3c22b9f80" = "Authentication Policy Administrator"
    "e3973bdf-4987-49ae-837a-ba8e231c7286" = "Azure DevOps Administrator"
    "7495fdc4-34c4-4d15-a289-98788ce399fd" = "Azure Information Protection Administrator"
    "aaf43236-0c0d-4d5f-883a-6955382ac081" = "B2C IEF Keyset Administrator"
    "3edaf663-341e-4475-9f94-5c398ef6c070" = "B2C IEF Policy Administrator"
    "b0f54661-2d74-4c50-afa3-1ec803f12efe" = "Billing Administrator"
    "892c5842-a9a6-463a-8041-72aa08ca3cf6" = "Cloud App Security Administrator"
    "158c047a-c907-4556-b7ef-446551a6b5f7" = "Cloud Application Administrator"
    "7698a772-787b-4ac8-901f-60d6b08affd2" = "Cloud Device Administrator"
    "17315797-102d-40b4-93e0-432062caca18" = "Compliance Administrator"
    "e6d1a23a-da11-4be4-9570-befc86d067a7" = "Compliance Data Administrator"
    "b1be1c3e-b65d-4f19-8427-f6fa0d97feb9" = "Conditional Access Administrator"
    "5c4f9dcd-47dc-4cf7-8c9a-9e4207cbfc91" = "Customer LockBox Access Approver"
    "38a96431-2bdf-4b4c-8b6e-5d3d8abac1a4" = "Desktop Analytics Administrator"
    "88d8e3e3-8f55-4a1e-953a-9b9898b8876b" = "Directory Readers"
    "d29b2b05-8046-44ba-8758-1e26182fcf32" = "Directory Synchronization Accounts"
    "9360feb5-f418-4baa-8175-e2a00bac4301" = "Directory Writers"
    "8329153b-31d0-4727-b945-745eb3bc5f31" = "Domain Name Administrator"
    "44367163-eba1-44c3-98af-f5787879f96a" = "Dynamics 365 Administrator"
    "963797fb-eb3b-4cde-8ce3-5878b3f32a3f" = "Dynamics 365 Business Central Administrator"
    "3f1acade-1e04-4fbc-9b69-f0302cd84aef" = "Edge Administrator"
    "29232cdf-9323-42fd-ade2-1d097af3e4de" = "Exchange Administrator"
    "31392ffb-586c-42d1-9346-e59415a2cc4e" = "Exchange Recipient Administrator"
    "6e591065-9bad-43ed-90f3-e9424366d2f0" = "External ID User Flow Administrator"
    "0f971eea-41eb-4569-a71e-57bb8a3eff1e" = "External ID User Flow Attribute Administrator"
    "be2f45a1-457d-42af-a067-6ec1fa63bc45" = "External Identity Provider Administrator"
    "a9ea8996-122f-4c74-9520-8edcd192826c" = "Fabric Administrator"
    "62e90394-69f5-4237-9190-012177145e10" = "Global Administrator"
    "f2ef992c-3afb-46b9-b7cf-a126ee74c451" = "Global Reader"
    "ac434307-12b9-4fa1-a708-88bf58caabc1" = "Global Secure Access Administrator"
    "fdd7a751-b60b-444a-984c-02652fe8fa1c" = "Groups Administrator"
    "95e79109-95c0-4d8e-aee3-d01accf2d47b" = "Guest Inviter"
    "729827e3-9c14-49f7-bb1b-9608f156bbb8" = "Helpdesk Administrator"
    "8ac3fc64-6eca-42ea-9e69-59f4c7b60eb2" = "Hybrid Identity Administrator"
    "45d8d3c5-c802-45c6-b32a-1d70b5e1e86e" = "Identity Governance Administrator"
    "eb1f4a8d-243a-41f0-9fbd-c7cdf6c5ef7c" = "Insights Administrator"
    "25df335f-86eb-4119-b717-0ff02de207e9" = "Insights Analyst"
    "31e939ad-9672-4796-9c2e-873181342d2d" = "Insights Business Leader"
    "3a2c62db-5318-420d-8d74-23affee5d9d5" = "Intune Administrator"
    "74ef975b-6605-40af-a5d2-b9539d836353" = "Kaizala Administrator"
    "b5a8dcf3-09d5-43a9-a639-8e29ef291470" = "Knowledge Administrator"
    "744ec460-397e-42ad-a462-8b3f9747a02c" = "Knowledge Manager"
    "4d6ac14f-3453-41d0-bef9-a3e0c569773a" = "License Administrator"
    "59d46f88-662b-457b-bceb-5c3809e5908f" = "Lifecycle Workflows Administrator"
    "ac16e43d-7b2d-40e0-ac05-243ff356ab5b" = "Message Center Privacy Reader"
    "790c1fb9-7f7d-4f88-86a1-ef1f95c05c1b" = "Message Center Reader"
    "8c8b803f-96e1-4129-9349-20738d9f9652" = "Microsoft 365 Migration Administrator"
    "9f06204d-73c1-4d4c-880a-6edb90606fd8" = "Microsoft Entra Joined Device Local Administrator"
    "1501b917-7653-4ff9-a4b5-203eaf33784f" = "Microsoft Hardware Warranty Administrator"
    "281fe777-fb20-4fbb-b7a3-ccebce5b0d96" = "Microsoft Hardware Warranty Specialist"
    "d24aef57-1500-4070-84db-2666f29cf966" = "Modern Commerce Administrator"
    "d37c8bed-0711-4417-ba38-b4abe66ce4c2" = "Network Administrator"
    "2b745bdf-0803-4d80-aa65-822c4493daac" = "Office Apps Administrator"
    "92ed04bf-c94a-4b82-9729-b799a7a4c178" = "Organizational Branding Administrator"
    "e48398e2-f4bb-4074-8f31-4586725e205b" = "Organizational Messages Approver"
    "507f53e4-4e52-4077-abd3-d2e1558b6ea2" = "Organizational Messages Writer"
    "4ba39ca4-527c-499a-b93d-d9b492c50246" = "Partner Tier1 Support"
    "e00e864a-17c5-4a4b-9c06-f5b95a8d5bd8" = "Partner Tier2 Support"
    "966707d0-3269-4727-9be2-8c3a10f19b9d" = "Password Administrator"
    "af78dc32-cf4d-46f9-ba4e-4428526346b5" = "Permissions Management Administrator"
    "11648597-926c-4cf3-9c36-bcebb0ba8dcc" = "Power Platform Administrator"
    "644ef478-e28f-4e28-b9dc-3fdde9aa0b1f" = "Printer Administrator"
    "e8cef6f1-e4bd-4ea8-bc07-4b8d950f4477" = "Printer Technician"
    "7be44c8a-adaf-4e2a-84d6-ab2649e08a13" = "Privileged Authentication Administrator"
    "e8611ab8-c189-46e8-94e1-60213ab1f814" = "Privileged Role Administrator"
    "4a5d8f65-41da-4de4-8968-e035b65339cf" = "Reports Reader"
    "0964bb5e-9bdb-4d7b-ac29-58e794862a40" = "Search Administrator"
    "8835291a-918c-4fd7-a9ce-faa49f0cf7d9" = "Search Editor"
    "194ae4cb-b126-40b2-bd5b-6091b380977d" = "Security Administrator"
    "5f2222b1-57c3-48ba-8ad5-d4759f1fde6f" = "Security Operator"
    "5d6b6bb7-de71-4623-b4af-96380a352509" = "Security Reader"
    "f023fd81-a637-4b56-95fd-791ac0226033" = "Service Support Administrator"
    "f28a1f50-f6e7-4571-818b-6a12f2af6b6c" = "SharePoint Administrator"
    "1a7d78b6-429f-476b-b8eb-35fb715fffd4" = "SharePoint Embedded Administrator"
    "75941009-915a-4869-abe7-691bff18279e" = "Skype for Business Administrator"
    "69091246-20e8-4a56-aa4d-066075b2a7a8" = "Teams Administrator"
    "baf37b3a-610e-45da-9e62-d9d1e5e8914b" = "Teams Communications Administrator"
    "f70938a0-fc10-4177-9e90-2178f8765737" = "Teams Communications Support Engineer"
    "fcf91098-03e3-41a9-b5ba-6f0ec8188a12" = "Teams Communications Support Specialist"
    "3d762c5a-1b6c-493f-843e-55a3b42923d4" = "Teams Devices Administrator"
    "aa38014f-0993-46e9-9b45-30501a20909d" = "Teams Telephony Administrator"
    "112ca1a2-15ad-4102-995e-45b0bc479a6a" = "Tenant Creator"
    "75934031-6c7e-415a-99d7-48dbd49e875e" = "Usage Summary Reports Reader"
    "fe930be7-5e62-47db-91af-98c3a49a38b1" = "User Administrator"
    "27460883-1df1-4691-b032-3b79643e5e63" = "User Experience Success Manager"
    "e300d9e7-4a2b-4295-9eff-f1c78b36cc98" = "Virtual Visits Administrator"
    "92b086b3-e367-4ef2-b869-1de128fb986e" = "Viva Goals Administrator"
    "87761b17-1ed2-4af3-9acd-92a150038160" = "Viva Pulse Administrator"
    "11451d60-acb2-45eb-a7d6-43d0f0125c13" = "Windows 365 Administrator"
    "32696413-001a-46ae-978c-ce0f6b3620d2" = "Windows Update Deployment Administrator"
    "810a2642-a034-447f-a5e8-41beaa378541" = "Yammer Administrator"
}

#######################################################################
# Basic functions
#######################################################################

# Get the list Enterprise Apps
$eApps = Get-MgServicePrincipal -All

function Get-NameFromGUID {
    param ($guids)
    $name = @()
    
    if (-Not ($guids -is [system.array])) {
        $guids = $guids -split ","
    }
    for ($i = 0; $i -lt $guids.Count; $i++) {
        if ($builtinNames.ContainsKey($guids[$i])) {
            $name += $builtinNames[$guids[$i]]
            continue
        }
        try {
            # Get the AuthenticationStrengthPolicy's name by ID
            $name += (Get-MgPolicyAuthenticationStrengthPolicy -AuthenticationStrengthPolicyId $guids[$i] -ErrorAction Stop).DisplayName
        } catch {
            try {
            # Get the user's name by ID
            $name += (Get-MgUser -UserId $guids[$i] -ErrorAction Stop).DisplayName
            } catch {
                try {
                    # Get the group's name by ID
                    $name += (Get-MgGroup -GroupId $guids[$i] -ErrorAction Stop).DisplayName
                } catch {
                    try {
                        # Get the managed app's name by ID
                        $name += (Get-MgApplication -ApplicationId $guids[$i] -ErrorAction Stop).DisplayName
                    } catch {
                        try {
                            # Get the policy's name by ID
                            $name += (Get-MgPolicy -PolicyId $guids[$i] -ErrorAction Stop).DisplayName
                        } catch {
                            try {
                                # Get the network's name by ID
                                $name += (Get-MgNetwork -NetworkId $guids[$i] -ErrorAction Stop).DisplayName
                            } catch {
                                try {
                                    # Get the CAP's name by ID
                                    $name += (Get-MgIdentityConditionalAccessNamedLocation -NamedLocationId $guids[$i] -ErrorAction Stop).DisplayName
                                } catch {
                                    try {
                                        # Get the enterpris app's name by ID
                                        $output = ($eApps | Where-Object -Property AppID -eq  $guids[$i]).DisplayName
                                        if ($output) {
                                            $name += $output
                                        } else {
                                            $name += $guids[$i]
                                        }
                                    }
                                    catch {
                                        $name += $guids[$i]  # Return the GUID if no match is found
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    
    return $name
}

# Convert values to text
function Convert-ToText {
    param ($value)
    if ($value -is [System.array] -and $value -notlike '*String*') {
        return ($value -join ", ")
    } elseif ($value -is [PSCustomObject]) {
        return ($value | ConvertTo-Json -Compress)
    } else {
        return $value
    }
}


#######################################################################
# Get the list of conditional access policies (CAP)
#######################################################################
$CAPs = Get-MgIdentityConditionalAccessPolicy -All # List of CAPs

# Create a custom object for CSV
$csv = @()
$i = 1
foreach ($CAP in $CAPs) {
    Write-Host "[$i/$($CAPs.count)] $($CAP.DisplayName)"
    $i += 1
    $csv += [PSCustomObject]@{
        "Conditions.Applications.ApplicationFilter.Mode" = Convert-ToText $CAP.Conditions.Applications.ApplicationFilter.Mode
        "Conditions.Applications.ApplicationFilter.Rule" = Convert-ToText $CAP.Conditions.Applications.ApplicationFilter.Rule
        "Conditions.Applications.ExcludeApplications" = Convert-ToText (Get-NameFromGUID $CAP.Conditions.Applications.ExcludeApplications)
        "Conditions.Applications.IncludeApplications" = Convert-ToText (Get-NameFromGUID $CAP.Conditions.Applications.IncludeApplications)
        "Conditions.Applications.IncludeAuthenticationContextClassReferences" = Convert-ToText $CAP.Conditions.Applications.IncludeAuthenticationContextClassReferences
        "Conditions.Applications.IncludeUserActions" = Convert-ToText $CAP.Conditions.Applications.IncludeUserActions
        "Conditions.ClientAppTypes" = Convert-ToText $CAP.Conditions.ClientAppTypes
        "Conditions.ClientApplications.ExcludeServicePrincipals" = Convert-ToText $CAP.Conditions.ClientApplications.ExcludeServicePrincipals
        "Conditions.ClientApplications.IncludeServicePrincipals" = Convert-ToText $CAP.Conditions.ClientApplications.IncludeServicePrincipals
        "Conditions.ClientApplications.ServicePrincipalFilter.Mode" = Convert-ToText $CAP.Conditions.ClientApplications.ServicePrincipalFilter.Mode
        "Conditions.ClientApplications.ServicePrincipalFilter.Rule" = Convert-ToText $CAP.Conditions.ClientApplications.ServicePrincipalFilter.Rule
        "Conditions.Devices.DeviceFilter.Mode" = Convert-ToText $CAP.Conditions.Devices.DeviceFilter.Mode
        "Conditions.Devices.DeviceFilter.Rule" = Convert-ToText $CAP.Conditions.Devices.DeviceFilter.Rule
        "Conditions.InsiderRiskLevels" = Convert-ToText $CAP.Conditions.InsiderRiskLevels
        "Conditions.Locations.ExcludeLocations" = Convert-ToText (Get-NameFromGUID $CAP.Conditions.Locations.ExcludeLocations)
        "Conditions.Locations.IncludeLocations" = Convert-ToText (Get-NameFromGUID $CAP.Conditions.Locations.IncludeLocations)
        "Conditions.Platforms.ExcludePlatforms" = Convert-ToText $CAP.Conditions.Platforms.ExcludePlatforms
        "Conditions.Platforms.IncludePlatforms" = Convert-ToText $CAP.Conditions.Platforms.IncludePlatforms
        "Conditions.ServicePrincipalRiskLevels" = Convert-ToText $CAP.Conditions.ServicePrincipalRiskLevels
        "Conditions.SignInRiskLevels" = Convert-ToText $CAP.Conditions.SignInRiskLevels
        "Conditions.UserRiskLevels" = Convert-ToText $CAP.Conditions.UserRiskLevels
        "Conditions.Users.ExcludeGroups" = Convert-ToText (Get-NameFromGUID $CAP.Conditions.Users.ExcludeGroups)
        "Conditions.Users.ExcludeGuestsOrExternalUsers.GuestOrExternalUserTypes" = Convert-ToText $CAP.Conditions.Users.ExcludeGuestsOrExternalUsers.GuestOrExternalUserTypes
        "Conditions.Users.ExcludeRoles" = Convert-ToText (Get-NameFromGUID $CAP.Conditions.Users.ExcludeRoles)
        "Conditions.Users.ExcludeUsers" = Convert-ToText (Get-NameFromGUID $CAP.Conditions.Users.ExcludeUsers)
        "Conditions.Users.IncludeGroups" = Convert-ToText (Get-NameFromGUID $CAP.Conditions.Users.IncludeGroups)
        "Conditions.Users.IncludeGuestsOrExternalUsers.GuestOrExternalUserTypes" = Convert-ToText $CAP.Conditions.Users.IncludeGuestsOrExternalUsers.GuestOrExternalUserTypes
        "Conditions.Users.IncludeRoles" = Convert-ToText (Get-NameFromGUID $CAP.Conditions.Users.IncludeRoles)
        "Conditions.Users.IncludeUsers" = Convert-ToText (Get-NameFromGUID $CAP.Conditions.Users.IncludeUsers)
        "CreatedDateTime" = Convert-ToText $CAP.CreatedDateTime
        "Description" = Convert-ToText $CAP.Description
        "DisplayName" = Convert-ToText $CAP.DisplayName
        "GrantControls.AuthenticationStrength.AllowedCombinations" = Convert-ToText $CAP.GrantControls.AuthenticationStrength.AllowedCombinations
        "GrantControls.AuthenticationStrength.CombinationConfigurations" = Convert-ToText $CAP.GrantControls.AuthenticationStrength.CombinationConfigurations
        "GrantControls.AuthenticationStrength.CreatedDateTime" = Convert-ToText $CAP.GrantControls.AuthenticationStrength.CreatedDateTime
        "GrantControls.AuthenticationStrength.Description" = Convert-ToText $CAP.GrantControls.AuthenticationStrength.Description
        "GrantControls.AuthenticationStrength.DisplayName" = Convert-ToText $CAP.GrantControls.AuthenticationStrength.DisplayName
        # "GrantControls.AuthenticationStrength.Id" = Convert-ToText $CAP.GrantControls.AuthenticationStrength.Id
        "GrantControls.AuthenticationStrength.ModifiedDateTime" = Convert-ToText $CAP.GrantControls.AuthenticationStrength.ModifiedDateTime
        "GrantControls.AuthenticationStrength.PolicyType" = Convert-ToText $CAP.GrantControls.AuthenticationStrength.PolicyType
        "GrantControls.AuthenticationStrength.RequirementsSatisfied" = Convert-ToText $CAP.GrantControls.AuthenticationStrength.RequirementsSatisfied
        "GrantControls.BuiltInControls" = Convert-ToText $CAP.GrantControls.BuiltInControls
        "GrantControls.CustomAuthenticationFactors" = Convert-ToText $CAP.GrantControls.CustomAuthenticationFactors
        "GrantControls.Operator" = Convert-ToText $CAP.GrantControls.Operator
        "GrantControls.TermsOfUse" = Convert-ToText $CAP.GrantControls.TermsOfUse
        "Id" = Convert-ToText $CAP.Id
        "ModifiedDateTime" = Convert-ToText $CAP.ModifiedDateTime
        "SessionControls.ApplicationEnforcedRestrictions.IsEnabled" = Convert-ToText $CAP.SessionControls.ApplicationEnforcedRestrictions.IsEnabled
        "SessionControls.CloudAppSecurity.CloudAppSecurityType" = Convert-ToText $CAP.SessionControls.CloudAppSecurity.CloudAppSecurityType
        "SessionControls.CloudAppSecurity.IsEnabled" = Convert-ToText $CAP.SessionControls.CloudAppSecurity.IsEnabled
        "SessionControls.DisableResilienceDefaults" = Convert-ToText $CAP.SessionControls.DisableResilienceDefaults
        "SessionControls.PersistentBrowser.IsEnabled" = Convert-ToText $CAP.SessionControls.PersistentBrowser.IsEnabled
        "SessionControls.PersistentBrowser.Mode" = Convert-ToText $CAP.SessionControls.PersistentBrowser.Mode
        "SessionControls.SignInFrequency.AuthenticationType" = Convert-ToText $CAP.SessionControls.SignInFrequency.AuthenticationType
        "SessionControls.SignInFrequency.FrequencyInterval" = Convert-ToText $CAP.SessionControls.SignInFrequency.FrequencyInterval
        "SessionControls.SignInFrequency.IsEnabled" = Convert-ToText $CAP.SessionControls.SignInFrequency.IsEnabled
        "SessionControls.SignInFrequency.Type" = Convert-ToText $CAP.SessionControls.SignInFrequency.Type
        "SessionControls.SignInFrequency.Value" = Convert-ToText $CAP.SessionControls.SignInFrequency.Value
        "State" = Convert-ToText $CAP.State
        "TemplateId" = Convert-ToText $CAP.TemplateId
        "AdditionalProperties" = Convert-ToText ($CAP.AdditionalProperties | Out-String)
    }
}

# Export to CSV
$csv | Export-Csv -Path "$OutputFolder\output_CAP_$Date.csv" -NoTypeInformation
