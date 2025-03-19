Install-Module AzureADPreview
Import-Module Microsoft.Graph.Beta.Identity.DirectoryManagement
Import-Module Microsoft.Graph.Beta.Groups

Connect-MgGraph -Scopes "Directory.ReadWrite.All", "Group.Read.All"

$SecurityGroupName = "Administradores Globais"
$SettingsObjectID = (Get-MgBetaDirectorySetting | Where-Object -Property Displayname -Value "Group.Unified" -EQ).id

If (!$SettingsObjectID) {

	$Params = @{
		templateId = "62375ab9-6b52-47ed-826b-58e47e0e304b"
		values     = @(
			@{
				name  = "EnableMSStandardBlockedWords"
				value = "True"
			}
		)
	}
	
	New-MgBetaDirectorySetting -BodyParameter $Params
	$SettingsObjectID = (Get-MgBetaDirectorySetting | Where-Object -Property Displayname -Value "Group.Unified" -EQ).Id

}
 
$GroupId = (Get-MgBetaGroup | Where-Object { $_.displayname -eq $SecurityGroupName }).Id

$Params = @{
	templateId = "62375ab9-6b52-47ed-826b-58e47e0e304b"
	values     = @(
		@{
			Name  = "EnableMIPLabels"
			Value = "True"
		}
		@{
			name  = "EnableGroupCreation"
			value = "False"
		}
		@{
			name  = "GroupCreationAllowedGroupId"
			value = $GroupId
		}
		@{
			name  = "EnableMSStandardBlockedWords"
			value = "True"
		}
	)
}

Update-MgBetaDirectorySetting -DirectorySettingId $SettingsObjectID -BodyParameter $Params
(Get-MgBetaDirectorySetting -DirectorySettingId $SettingsObjectID).Values

Disconnect-MgGraph