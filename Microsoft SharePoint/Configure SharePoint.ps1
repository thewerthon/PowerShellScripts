# Install: https://pnp.github.io/powershell/articles/installation.html
# Register: https://pnp.github.io/powershell/articles/registerapplication.html
# Uninstall-Module PnP.PowerShell -AllVersions
# Install-Module PnP.PowerShell -Scope AllUsers -AllowPrerelease -SkipPublisherCheck
# Update-Module PnP.PowerShell -Scope AllUsers -AllowPrerelease -Force

# Function to Test Object
Function Test-Object {

    Param(
        [Object]$Object = $Null,
        [String]$Message,
        [Switch]$Silent
    )

    If ($Null -Ne $Object) { Return $True }
    If ($Message -And -Not $Silent) { Write-Host $Message -ForegroundColor Red }
    Return $False

}

# Function to Test Object Single
Function Test-ObjectSingle {

    Param(
        [Object]$Object = $Null,
        [String]$Message,
        [Switch]$Silent
    )

    If ($Object.Count -Eq 1) { Return $True }
    If ($Message -And -Not $Silent) { Write-Host $Message -ForegroundColor Red }
    Return $False

}

# Function to Test Object Collection
Function Test-ObjectCollection {

    Param(
        [Object]$Object = $Null,
        [String]$Message,
        [Switch]$Silent
    )

    If ($Object.Count -Gt 1) { Return $True }
    If ($Message -And -Not $Silent) { Write-Host $Message -ForegroundColor Red }
    Return $False

}

# Function to Invoke Operation
Function Invoke-Operation {

    Param(
        [Parameter(Mandatory = $True)][String]$Message,
        [Parameter(Mandatory = $True)][ScriptBlock]$Operation,
        [Switch]$OutputInfos,
        [Switch]$OutputErrors,
        [Switch]$ReturnValue,
        [Switch]$Silent
    )

    Try {

        If (-Not $Silent) { Write-Host "$($Message)... " -ForegroundColor Cyan -NoNewline }
        
        $Output = & $Operation *>&1

        If (-Not $Silent) { Write-Host "success!" -ForegroundColor Green }

        If ($OutputInfos) { $Output | ForEach-Object { If ($_ -IsNot [System.Management.Automation.ErrorRecord]) { Write-Host $_ -ForegroundColor White } } }

        If ($OutputErrors) { $Output | ForEach-Object { If ($_ -Is [System.Management.Automation.ErrorRecord]) { Write-Host $_ -ForegroundColor Red } } }

        If ($ReturnValue) { Return $Output }

    } Catch {

        If (-Not $Silent) { Write-Host "failed!" -ForegroundColor Magenta }
        
        If ($OutputInfos) { $Output | ForEach-Object { If ($_ -IsNot [System.Management.Automation.ErrorRecord]) { Write-Host $_ -ForegroundColor White } } }

        If ($OutputErrors) { $Output | ForEach-Object { If ($_ -Is [System.Management.Automation.ErrorRecord]) { Write-Host $_ -ForegroundColor Red } } }

        If ($ReturnValue) { Return $Null }

    }

}

# Function to Test Tenant
Function Test-Tenant {

    Param(
        [Switch]$Silent
    )

    Test-Object -Object $Global:CurrentTenant -Message "Not connected to a tenant." -Silent:$Silent

}

# Function to Connect Tenant
Function Connect-Tenant {

    Param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)][Object]$Tenant,
        [Switch]$Silent
    )

    Process {

        If (-Not (Test-ObjectSingle -Object $Tenant -Message "Multiple tenants detected, please provide only one tenant." -Silent:$Silent)) { Return }

        Invoke-Operation -Message "Connecting to tenant: $($Tenant.Name)" -Silent:$Silent -Operation {
            
            Connect-PnPOnline -Url $Tenant.AdminUrl -ClientId $Tenant.ClientID -OSLogin
            Set-Variable -Name "CurrentTenant" -Value $Tenant -Scope Global
            Set-Variable -Name "CurrentSite" -Value $Tenant.AdminUrl -Scope Global

        }

    }

}

# Function to Disconnect Tenant
Function Disconnect-Tenant {

    Param (
        [Switch]$Silent
    )

    If (-Not (Test-Tenant -Silent:$Silent)) { Return }

    Invoke-Operation -Message "Disconnecting from tenant: $($Global:CurrentTenant.Name)" -Silent:$Silent -Operation {

        Disconnect-PnPOnline
        Set-Variable -Name "CurrentTenant" -Value $Null -Scope Global
        Set-Variable -Name "CurrentSite" -Value $Null -Scope Global

    }

}

# Function to Get Tenant
Function Get-Tenant {

    Param(
        [String]$Slug
    )

    #$Tenant = Get-PnPTenant
    $Tenant = $Tenants | Where-Object Slug -EQ $Slug
    Return $Tenant

}

# Function To Get Tenants
Function Get-Tenants {

    $Tenants = @(

        [PSCustomObject]@{
            Slug     = "siwindbr"
            Name     = "SIW Kits Eólicos"
            Theme    = "SIW 2020"
            Domain   = "siw.ind.br"
            BaseUrl  = "https://siwindbr.sharepoint.com"
            AdminUrl = "https://siwindbr-admin.sharepoint.com"
            ClientID = "8b14c0ea-5f50-4c5c-b2f8-a50b5ca20d8b"
        },

        [PSCustomObject]@{
            Slug     = "gcgestao"
            Name     = "GC Gestão"
            Theme    = "GC 2025"
            Domain   = "gcgestao.com.br"
            BaseUrl  = "https://gcgestao.sharepoint.com"
            AdminUrl = "https://gcgestao-admin.sharepoint.com"
            ClientID = "91aac6c3-b063-4175-8073-7e5b5a4ff281"
        },

        [PSCustomObject]@{
            Slug     = "inteceletrica"
            Name     = "Intec Elétrica"
            Theme    = "Intec 2025"
            Domain   = "inteceletrica.com.br"
            BaseUrl  = "https://inteceletrica.sharepoint.com"
            AdminUrl = "https://inteceletrica-admin.sharepoint.com"
            ClientID = "7735abc1-32a8-416b-a7be-3d2496ba4724"
        }

    )

    Return $Tenants

}

# Function to Set Tenant
Function Set-Tenant {

    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [Object]$Tenant
    )

    Begin {
    
        If (-Not (Test-Tenant)) { Return }

    }

    Process {

        $TenantParams = @{
            AllowCommentsTextOnEmailEnabled            = $True
            AllowFilesWithKeepLabelToBeDeletedODB      = $False
            AllowFilesWithKeepLabelToBeDeletedSPO      = $False
            AnyoneLinkTrackUsers                       = $True
            BlockUserInfoVisibilityInOneDrive          = "ApplyToNoUsers "
            BlockUserInfoVisibilityInSharePoint        = "ApplyToNoUsers"
            CommentsOnFilesDisabled                    = $False
            CommentsOnListItemsDisabled                = $False
            ConditionalAccessPolicy                    = "AllowFullAccess"
            CoreDefaultLinkToExistingAccess            = $True
            CoreDefaultShareLinkRole                   = "View"
            CoreDefaultShareLinkScope                  = "SpecificPeople"
            CoreSharingCapability                      = "ExternalUserAndGuestSharing"
            DefaultLinkPermission                      = "View"
            DefaultSharingLinkType                     = "Direct"
            DisableAddToOneDrive                       = $True
            DisableBackToClassic                       = $True
            DisablePersonalListCreation                = $False
            DisplayNamesOfFileViewers                  = $True
            DisplayNamesOfFileViewersInSpo             = $True
            DisplayStartASiteOption                    = $False
            EnableAIPIntegration                       = $True
            EnableAutoExpirationVersionTrim            = $True
            EnableAutoNewsDigest                       = $True
            EnableDiscoverableByOrganizationForVideos  = $True
            EnableSensitivityLabelForPDF               = $True
            ExtendPermissionsToUnprotectedFiles        = $True
            ExternalUserExpirationRequired             = $True
            ExternalUserExpireInDays                   = 90
            FileAnonymousLinkType                      = "View"
            FolderAnonymousLinkType                    = "View"
            HideDefaultThemes                          = $True
            HideSyncButtonOnDocLib                     = $True
            HideSyncButtonOnODB                        = $True
            HideSyncButtonOnTeamSite                   = $True
            IncludeAtAGlanceInShareEmails              = $True
            IsDataAccessInCardDesignerEnabled          = $True
            IsFluidEnabled                             = $True
            IsLoopEnabled                              = $True
            MassDeleteNotificationDisabled             = $False
            ODBAccessRequests                          = "On"
            ODBMembersCanShare                         = "Off"
            OneDriveDefaultLinkToExistingAccess        = $False
            OneDriveDefaultShareLinkRole               = "View"
            OneDriveDefaultShareLinkScope              = "SpecificPeople"
            OneDriveSharingCapability                  = "ExistingExternalUserSharingOnly"
            OrphanedPersonalSitesRetentionPeriod       = 365
            PreventExternalUsersFromReSharing          = $True
            ProvisionSharedWithEveryoneFolder          = $False
            PublicCdnAllowedFileTypes                  = "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF"
            PublicCdnEnabled                           = $True
            RecycleBinRetentionPeriod                  = 93
            RequireAcceptingAccountMatchInvitedAccount = $True
            RequireAnonymousLinksExpireInDays          = 90
            SearchResolveExactEmailOrUPN               = $False
            SelfServiceSiteCreationDisabled            = $True
            SharingCapability                          = "ExternalUserAndGuestSharing"
            ShowAllUsersClaim                          = $False
            ShowEveryoneClaim                          = $False
            ShowEveryoneExceptExternalUsersClaim       = $False
            ShowOpenInDesktopOptionForSyncedFiles      = $True
            SocialBarOnSitePagesDisabled               = $False
            SpecialCharactersStateInFileFolderNames    = "Allowed"
            ViewersCanCommentOnMediaDisabled           = $False
        }

        Invoke-Operation -Message "Setting tenant: $($Tenant.Name)" -Operation {

            Set-PnPTenant @TenantParams -Force

        }

    }

}

# Function to Test Site
Function Test-Site {

    Param(
        [Switch]$Silent
    )

    Test-Object -Object $Global:CurrentSite -Message "Not connected to a site." -Silent:$Silent

}

# Function to Connect Site
Function Connect-Site {

    Param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)][Object]$Site,
        [Switch]$ReturnConnection,
        [Switch]$Silent
    )

    Process {

        If (-Not (Test-Tenant -Silent:$Silent)) { Return }
        If (-Not (Test-ObjectSingle -Object $Site -Message "Multiple sites detected, please provide only one site." -Silent:$Silent)) { Return }

        $SiteUrl = If ($Site -Is [String]) { $Site } Else { $Site.Url }
        $SiteTitle = If ($Site -Is [String]) { $Site } Else { $Site.Title }

        Invoke-Operation "Connecting to site: $($SiteTitle)" -ReturnValue:$ReturnConnection -Silent:$Silent -Operation {

            Connect-PnPOnline -Url $SiteUrl -ClientId $Global:CurrentTenant.ClientID -OSLogin -ReturnConnection:$ReturnConnection
            Set-Variable -Name "CurrentSite" -Value $Site -Scope Global

        }

    }

}

# Function to Disconnect Site
Function Disconnect-Site {

    Param (
        [Switch]$Silent
    )

    If (-Not (Test-Site -Silent:$Silent)) { Return }
    If (-Not (Test-Tenant -Silent:$Silent)) { Return }

    $SiteTitle = If ($Global:CurrentSite -Is [String]) { $Global:CurrentSite } Else { $Global:CurrentSite.Title }

    Invoke-Operation -Message "Disconnecting from site: $($SiteTitle)" -Silent:$Silent -Operation {

        Disconnect-PnPOnline
        Set-Variable -Name "CurrentTenant" -Value $Null -Scope Global
        Set-Variable -Name "CurrentSite" -Value $Null -Scope Global

    }

    Connect-PnPOnline -Url $Global:CurrentTenant.AdminUrl -ClientId $Global:CurrentTenant.ClientID -OSLogin
    Set-Variable -Name "CurrentSite" -Value $Global:CurrentTenant.AdminUrl -Scope Global

}

# Function to Get Site
Function Get-Site {

    Param(
        [Parameter(Mandatory = $True)][String]$Url
    )

    If (-Not (Test-Tenant)) { Return }
    $Site = Get-PnPTenantSite -Identity $Url
    Return $Site

}

# Function to Get Sites
Function Get-Sites {

    Param(
        [Switch]$Filter
    )

    If (-Not (Test-Tenant)) { Return }
    $Sites = Get-PnPTenantSite
    If ($Filter) { $Sites = $Sites | Where-Object { ($_.Template -Match "SitePage" -Or $_.Template -Match "Group") -And $_.Url -NotMatch "/marca" } }
    Return $Sites

}

# Function to Set Site
Function Set-Site {

    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)][Object]$Site,
        [Switch]$SetAll,
        [Switch]$SetAdmins,
        [Switch]$SetVersioning
    )

    Begin {

        If (-Not (Test-Tenant)) { Return }
        $GlobalAdmin = (Get-PnPUser | Where-Object Title -Match "Administradores Globais")[0]
        $OtherAdmins = $Null

    }

    Process {

        # Start Site Params
        If ($Site.Template -Match "SitePage" -And $Site.Url.EndsWith("sharepoint.com/")) {

            $SiteParams = @{
                DefaultLinkPermission                       = "View"
                DefaultLinkToExistingAccess                 = $True
                DefaultShareLinkRole                        = "View"
                DefaultShareLinkScope                       = "SpecificPeople"
                DefaultSharingLinkType                      = "Direct"
                DenyAddAndCustomizePages                    = $True
                DisableSharingForNonOwners                  = $True
                InheritVersionPolicyFromTenant              = $True
                OverrideSharingCapability                   = $False
                OverrideTenantAnonymousLinkExpirationPolicy = $False
                OverrideTenantExternalUserExpirationPolicy  = $False
                SharingCapability                           = "ExternalUserAndGuestSharing"
            }

        }
        
        # Others Sites Params
        If ($Site.Template -Match "SitePage" -And -Not $Site.Url.EndsWith("sharepoint.com/")) {

            $SiteParams = @{
                DefaultLinkPermission                       = "View"
                DefaultLinkToExistingAccess                 = $True
                DefaultShareLinkRole                        = "View"
                DefaultShareLinkScope                       = "SpecificPeople"
                DefaultSharingLinkType                      = "Direct"
                DenyAddAndCustomizePages                    = $True
                DisableSharingForNonOwners                  = $True
                InheritVersionPolicyFromTenant              = $True
                OverrideSharingCapability                   = $False
                OverrideTenantAnonymousLinkExpirationPolicy = $False
                OverrideTenantExternalUserExpirationPolicy  = $False
                SharingCapability                           = "ExistingExternalUserSharingOnly"
            }

        }
        
        # Group Sites Params
        If ($Site.Template -Match "Group") {

            $SiteParams = @{
                DefaultLinkPermission                       = "View"
                DefaultLinkToExistingAccess                 = $True
                DefaultShareLinkRole                        = "View"
                DefaultShareLinkScope                       = "SpecificPeople"
                DefaultSharingLinkType                      = "Direct"
                DenyAddAndCustomizePages                    = $True
                DisableSharingForNonOwners                  = $True
                InheritVersionPolicyFromTenant              = $True
                OverrideSharingCapability                   = $False
                OverrideTenantAnonymousLinkExpirationPolicy = $False
                OverrideTenantExternalUserExpirationPolicy  = $False
                SharingCapability                           = "ExistingExternalUserSharingOnly"
            }
            
        }

        # Connect to Site
        $Connection = Invoke-Operation -Message "Connecting to site: $($Site.Title)" -ReturnValue -Operation {

            Connect-Site -Site $Site -Silent -ReturnConnection

        }

        # Set Site Admin
        If ($SetAll -Or $SetAdmins) {

            Invoke-Operation -Message "Setting site administrators" -Operation {

                $SiteAdmins = Get-PnPSiteCollectionAdmin -Connection $Connection
                Add-PnPSiteCollectionAdmin -Owners $OtherAdmins.LoginName -PrimarySiteCollectionAdmin $GlobalAdmin.LoginName -Connection $Connection
                $SiteAdmins | Where-Object LoginName -NotIn ($GlobalAdmin.LoginName, $OtherAdmins.LoginName) | Remove-PnPSiteCollectionAdmin -Connection $Connection

            }

        }

        # Set Versioning
        If ($SetAll -Or $SetVersioning) {

            Invoke-Operation -Message "Setting site versioning" -Operation {

                Set-PnPSiteVersionPolicy -EnableAutoExpirationVersionTrim $True -ApplyToNewDocumentLibraries -ApplyToExistingDocumentLibraries -Connection $Connection
                Set-PnPSiteVersionPolicy -InheritFromTenant -Connection $Connection

            }

        }

        # Set Site Params
        Invoke-Operation -Message "Setting site parameters" -Operation {

            Set-PnPTenantSite -Identity $Site.Url @SiteParams -Connection $Connection
            Set-PnPWebHeader -HeaderLayout "Standard" -HeaderEmphasis "None" -HideTitleInHeader:$False -HeaderBackgroundImageUrl $Null -LogoAlignment Left -Connection $Connection
            Set-PnPFooter -Enabled:$False -Layout "Simple" -BackgroundTheme "Neutral" -Title $Null -LogoUrl $Null -Connection $Connection
            Disable-PnPSharingForNonOwnersOfSite -Identity $Site.Url -Connection $Connection

        }

    }

}

# Get-PnPSitePolicy
# Get-PnPSiteVersionPolicy
# Get-PnPSiteVersionPolicyStatus
# Get-PnPWeb
# Get-PnPSubWeb
# Get-PnPTenantCdnEnabled -CdnType Public
# Get-PnPTenantId
# Get-PnPUser
# Get-PnPUserProfilePhoto
# Get-PnPUserProfileProperty

# Get-PnPNavigationNode

# Get-PnPPage
# Get-PnPPageComponent
# Get-PnPFeature
# Get-PnPGroup
# Get-PnPGroupMember
# Get-PnPGroupPermissions
# Get-PnPSiteGroup
# Get-PnPFileSharingLink
# Get-PnPFolderSharingLink
# Get-PnPSearchConfiguration
# Get-PnPSearchSettings
# Configure Personal Sites

# Invoke Testing
Function Invoke-Testing {

    # Connect Tenant
    $Tenants = Get-Tenants
    $Tenant = $Tenants[0]
    $Tenant | Connect-Tenant

    # Connect Site
    $Sites = Get-Sites -Filter
    $Site = $Sites[0]
    $Site | Connect-Site

}


# Function to Get Fields
Function Get-Fields {

    Param(
        [Parameter(Mandatory = $True)]
        [Object]$List
    )

    $Fields = Get-PnPField -List $List.Id
    Return $Fields

}

# Function to Get Views
Function Get-Views {

    Param(
        [Parameter(Mandatory = $True)]
        [Object]$List
    )

    $Views = Get-PnPView -List $List.Id | Where-Object { $_.Hidden -Eq $False }
    Return $Views

}

# Function to Get Lists
Function Get-Lists {

    $Lists = Get-PnPList | Where-Object { $_.Hidden -Eq $False -And $_.IsCatalog -Eq $False -And $_.BaseType -In ("DocumentLibrary", "GenericList") }
    Return $Lists

}

# Function to Set Field
Function Set-Field {

    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [Object]$Field
    )

    Process {
        
        # Fields Formats
        $DateFormat = '{"$schema":"https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json","elmType":"div","children":[{"elmType":"span","style":{"padding-right":"10px","padding-bottom":"2px","font-size":"16px"},"attributes":{"iconName":"Calendar"}},{"elmType":"span","txtContent":"@currentField.displayValue","style":{"padding-bottom":"4px"}}]}'
        $PersonFormat = '{"$schema":"https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json","elmType":"div","style":{"display":"flex","flex-wrap":"wrap","overflow":"hidden"},"children":[{"elmType":"div","forEach":"person in @currentField","defaultHoverField":"[$person]","attributes":{"class":"ms-bgColor-neutralLighter ms-fontColor-neutralSecondary"},"style":{"display":"flex","overflow":"hidden","align-items":"center","border-radius":"28px","margin":"4px 8px 4px 0px","min-width":"28px","height":"28px"},"children":[{"elmType":"img","attributes":{"src":"=''/_layouts/15/userphoto.aspx?size=S&accountname='' + [$person.email]","title":"[$person.title]"},"style":{"width":"28px","height":"28px","display":"block","border-radius":"50%"}},{"elmType":"div","style":{"overflow":"hidden","white-space":"nowrap","text-overflow":"ellipsis","padding":"0px 12px 2px 6px","display":"=if(length(@currentField) > 1, ''none'', ''flex'')","flex-direction":"column"},"children":[{"elmType":"span","txtContent":"[$person.title]","style":{"display":"inline","overflow":"hidden","white-space":"nowrap","text-overflow":"ellipsis","font-size":"12px","height":"15px"}},{"elmType":"span","txtContent":"[$person.department]","style":{"display":"inline","overflow":"hidden","white-space":"nowrap","text-overflow":"ellipsis","font-size":"9px"}}]}]}]}'

        Switch ($Field.InternalName) {
            
            "Author" { $Field.Title = "Criado Por"; $Field.CustomFormatter = $PersonFormat }
            "Editor" { $Field.Title = "Modificado Por"; $Field.CustomFormatter = $PersonFormat }
            "Created" { $Field.Title = "Criado Em"; $Field.CustomFormatter = $DateFormat }
            "Modified" { $Field.Title = "Modificado Em"; $Field.CustomFormatter = $DateFormat }
            Default {}

        }

        $Field.Update()
        

    }

}

# Function to Set View
Function Set-View {

    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [Object]$View
    )

    Process {
        
    }

}

# Function to Set List
Function Set-List {

    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [Object]$List
    )

    Process {

        # Configure Columns Formats
        Get-PnPField -List $List.Id | Where-Object { $_.InternalName -In $ColumnsMapping.Keys } | ForEach-Object {

            If ($ColumnsMapping.ContainsKey($_.InternalName)) {

                $_.Title = $ColumnsMapping[$_.InternalName].Title
                $_.CustomFormatter = $ColumnsMapping[$_.InternalName].CustomFormatter

            }

        }

    }
    
}

# Configure Lists
ForEach ($Item In $Lists) {

    # Configure Views
    $Item.Views | Where-Object { $_.Hidden -Eq $False } | ForEach-Object {

        If ($Item.BaseType -Eq "DocumentLibrary") {



        }

    }

    # Configure Views
    ForEach ($View in $Views) {

        If ($List.BaseType -Eq "DocumentLibrary") {
            
            $ViewFields = @("DocIcon", "LinkFilename", "Author", "Created", "Editor", "Modified", "FileSizeDisplay")
            $ColumnWidth = '<FieldRef Name="Nome" width="500" /><FieldRef Name="Modificado Por" width="200" /><FieldRef Name="Modificado Em" width="200" /><FieldRef Name="Criado Por" width="200" /><FieldRef Name="Criado Em" width="200" /><FieldRef Name="Tamanho do Arquivo" width="200" />'
        
        } Else {
            
            $ViewFields = $View.ViewFields
            $ColumnWidth = $View.ColumnWidth
        
        }

        Set-PnPView -Identity $View.Id -Values @{ ColumnWidth = $ColumnWidth; } -List $List.Id -Fields $ViewFields

    }

    # <FieldRef Name="Modificado Por" width="229" /><FieldRef Name="Nome" width="418" /><FieldRef Name="Modificado Em" width="188" />

}