# Uninstall-Module PnP.PowerShell -AllVersions
# Install-Module PnP.PowerShell -Scope AllUsers -AllowPrerelease -SkipPublisherCheck
# Update-Module PnP.PowerShell -Scope AllUsers -AllowPrerelease -Force

# Clear Host
Clear-Host

# Function to Grant App Rights
Function Grant-AppRights {

    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [Object]$Tenant
    )

    If (-Not (Get-Module MSOnline)) { Install-Module MSOnline -Force }

    Connect-MSOLService
    
    $App = Get-MsolServicePrincipal -AppPrincipalId $Tenant.ClientID
    $Role = Get-MsolRole -RoleName "Company Administrator"

    Add-MsolRoleMember -RoleObjectId $Role.ObjectId -RoleMemberType "ServicePrincipal" -RoleMemberObjectId $App.ObjectId
    Get-MsolRoleMember -RoleObjectId $Role.ObjectId

}

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

# Function to Test Tenant
Function Test-Tenant {

    Param(
        [Switch]$Silent
    )

    Test-Object -Object $Global:CurrentTenant -Message "Not connected to a tenant." -Silent:$Silent

}

# Function to Test Site
Function Test-Site {

    Param(
        [Switch]$Silent
    )

    Test-Object -Object $Global:CurrentSite -Message "Not connected to a site." -Silent:$Silent

}

# Function to Invoke Operation
Function Invoke-Operation {

    Param(
        [Parameter(Mandatory = $True)][String]$Message,
        [Parameter(Mandatory = $True)][ScriptBlock]$Operation,
        [Switch]$ShowInfo,
        [Switch]$ShowErrors,
        [Switch]$ReturnValue,
        [Switch]$Silent
    )

    Try {

        If (-Not $Silent) { Write-Host "$($Message)... " -ForegroundColor Cyan -NoNewline }
        
        $Output = & $Operation *>&1

        If (-Not $Silent) { Write-Host "success!" -ForegroundColor Green }

        If ($ShowInfo) { Write-Host $Output -ForegroundColor Gray }

        If ($ReturnValue) { Return $Output }

        $Output

    } Catch {

        If (-Not $Silent) { Write-Host "failed!" -ForegroundColor Magenta }
        
        If ($ShowInfo) { $Output }

        If ($ShowErrors) { Write-Host $_.Exception.Message -ForegroundColor Red }

        If ($ReturnValue) { Return $Null }

        $Output

    }

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
            
            Connect-PnPOnline -Url $Tenant.AdminUrl -ClientId $Tenant.ClientID -Interactive
            Set-Variable -Name "CurrentTenant" -Value $Tenant -Scope Global
            Set-Variable -Name "CurrentSite" -Value $Tenant.AdminUrl -Scope Global

        }

    }

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

            Connect-PnPOnline -Url $SiteUrl -ClientId $Global:CurrentTenant.ClientID -Interactive -ReturnConnection:$ReturnConnection
            Set-Variable -Name "CurrentSite" -Value $Site -Scope Global

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

    Connect-PnPOnline -Url $Global:CurrentTenant.AdminUrl -ClientId $Global:CurrentTenant.ClientID -Interactive
    Set-Variable -Name "CurrentSite" -Value $Global:CurrentTenant.AdminUrl -Scope Global

}

# Function To Get Tenants
Function Get-Tenants {

    $Tenants = @(

        [PSCustomObject]@{
            Slug     = "siwindbr"
            Name     = "SIW Kits Eólicos"
            Domain   = "siw.ind.br"
            BaseUrl  = "https://siwindbr.sharepoint.com"
            AdminUrl = "https://siwindbr-admin.sharepoint.com"
            ClientID = "8b14c0ea-5f50-4c5c-b2f8-a50b5ca20d8b"
            EventsID = "502190fd-356c-434c-a73f-db7146b5c1eb"
            Theme    = '{"name":"SIW Kits Eólicos","isInverted":false,"palette":{"themeDarker":"#835719","themeDark":"#b27622","themeDarkAlt":"#d38c28","themePrimary":"#eb9d2d","themeSecondary":"#eda744","themeTertiary":"#f3c27d","themeLight":"#f9e0bc","themeLighter":"#fceedb","themeLighterAlt":"#fefbf6","black":"#000000","neutralDark":"#201f1e","neutralPrimary":"#323130","neutralPrimaryAlt":"#3b3a39","neutralSecondary":"#605e5c","neutralTertiary":"#a19f9d","neutralTertiaryAlt":"#c8c6c4","neutralLight":"#edebe9","neutralLighter":"#f3f2f1","neutralLighterAlt":"#faf9f8","white":"#ffffff","neutralQuaternaryAlt":"#e1dfdd","neutralQuaternary":"#d0d0d0","accent":"#ffc000"}}'
        },

        [PSCustomObject]@{
            Slug     = "gcgestao"
            Name     = "GC Gestão"
            Domain   = "gcgestao.com.br"
            BaseUrl  = "https://gcgestao.sharepoint.com"
            AdminUrl = "https://gcgestao-admin.sharepoint.com"
            ClientID = "91aac6c3-b063-4175-8073-7e5b5a4ff281"
            EventsID = "2eb9023a-c795-4c5e-b536-2975c670ac40"
            Theme    = '{"name":"GC Gestão","isInverted":false,"palette":{"themeDarker":"#002a61","themeDark":"#003984","themeDarkAlt":"#00449c","themePrimary":"#004aad","themeSecondary":"#165cb7","themeTertiary":"#5288ce","themeLight":"#a1bfe7","themeLighter":"#cbdcf2","themeLighterAlt":"#f2f6fc","black":"#000000","neutralDark":"#201f1e","neutralPrimary":"#323130","neutralPrimaryAlt":"#3b3a39","neutralSecondary":"#605e5c","neutralTertiary":"#a19f9d","neutralTertiaryAlt":"#c8c6c4","neutralLight":"#edebe9","neutralLighter":"#f3f2f1","neutralLighterAlt":"#faf9f8","white":"#ffffff","neutralQuaternaryAlt":"#e1dfdd","neutralQuaternary":"#d0d0d0","accent":"#159aff"}}'
        },

        [PSCustomObject]@{
            Slug     = "inteceletrica"
            Name     = "Intec Elétrica"
            Domain   = "inteceletrica.com.br"
            BaseUrl  = "https://inteceletrica.sharepoint.com"
            AdminUrl = "https://inteceletrica-admin.sharepoint.com"
            ClientID = "7735abc1-32a8-416b-a7be-3d2496ba4724"
            EventsID = "c38fba0b-6ca6-4e1b-9b55-f284a18a6333"
            Theme    = '{"name":"Intec Elétrica","isInverted":false,"palette":{"themeDarker":"#002849","themeDark":"#003663","themeDarkAlt":"#004075","themePrimary":"#004782","themeSecondary":"#115891","themeTertiary":"#4883b4","themeLight":"#98bcda","themeLighter":"#c5daeb","themeLighterAlt":"#f0f6fa","black":"#000000","neutralDark":"#201f1e","neutralPrimary":"#323130","neutralPrimaryAlt":"#3b3a39","neutralSecondary":"#605e5c","neutralTertiary":"#a19f9d","neutralTertiaryAlt":"#c8c6c4","neutralLight":"#edebe9","neutralLighter":"#f3f2f1","neutralLighterAlt":"#faf9f8","white":"#ffffff","neutralQuaternaryAlt":"#e1dfdd","neutralQuaternary":"#d0d0d0","accent":"#159aff"}}'
        }

    )

    Return $Tenants

}

# Function to Get Sites
Function Get-Sites {

    Param(
        [Switch]$SharePoint,
        [Switch]$OneDrive,
        [Switch]$Teams,
        [Switch]$Channels
    )

    If (-Not (Test-Tenant)) { Return }

    $Sites = @()
    $AllSites = Get-PnPTenantSite -IncludeOneDriveSites

    If ($SharePoint) { $Sites = $Sites + ($AllSites | Where-Object Template -Match "SitePage" | Where-Object Url -NotMatch "/marca") }
    If ($OneDrive) { $Sites = $Sites + ($AllSites | Where-Object Template -Match "SpsPers" | Where-Object Url -Match "/personal/") }
    If ($Teams) { $Sites = $Sites + ($AllSites | Where-Object Template -Match "Group") }
    If ($Channels) { $Sites = $Sites + ($AllSites | Where-Object Template -Match "TeamChannel") }

    If (-Not $SharePoint -And -Not $OneDrive -And -Not $Teams -And -Not $Channels) { $Sites = $AllSites }

    Return $Sites

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

# Function to Get Site
Function Get-Site {

    Param(
        [Parameter(Mandatory = $True)][String]$Url
    )

    If (-Not (Test-Tenant)) { Return }
    $Site = Get-PnPTenantSite -Identity $Url
    Return $Site

}

# Function to Set Tenant
Function Set-Tenant {

    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [Object]$Tenant,
        [Switch]$ShowInfo,
        [Switch]$ShowErrors
    )

    Begin {
    
        If (-Not (Test-Tenant)) { Return }

    }

    Process {

        $TenantParams = @{
            AllowCommentsTextOnEmailEnabled            = $True
            AllowFilesWithKeepLabelToBeDeletedODB      = $False
            AllowFilesWithKeepLabelToBeDeletedAppO     = $False
            AnyoneLinkTrackUsers                       = $True
            BlockUserInfoVisibilityInOneDrive          = "ApplyToNoUsers "
            BlockUserInfoVisibilityInSharePoint        = "ApplyToNoUsers"
            CommentsOnFilesDisabled                    = $False
            CommentsOnListItemsDisabled                = $False
            ConditionalAccesAppolicy                   = "AllowFullAccess"
            CoreDefaultLinkToExistingAccess            = $True
            CoreDefaultShareLinkRole                   = "View"
            CoreDefaultShareLinkScope                  = "SpecificPeople"
            CoreSharingCapability                      = "ExternalUserAndGuestSharing"
            DefaultLinkPermission                      = "View"
            DefaultSharingLinkType                     = "Direct"
            DisableAddToOneDrive                       = $True
            DisableBackToClassic                       = $True
            DisablePersonalListCreation                = $False
            DiApplayNamesOfFileViewers                 = $True
            DiApplayNamesOfFileViewersInAppo           = $True
            DiApplayStartASiteOption                   = $False
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
            AppecialCharactersStateInFileFolderNames   = "Allowed"
            ViewersCanCommentOnMediaDisabled           = $False
        }

        Invoke-Operation -Message "Setting tenant: $($Tenant.Name)" -ShowInfo:$ShowInfo -ShowErrors:$ShowErrors -Operation {

            Set-PnPTenant @TenantParams -Force

        }

    }

}

# Function to Set Site
Function Set-Site {

    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)][Object]$Site,
        [Switch]$All,
        [Switch]$Admins,
        [Switch]$Params,
        [Switch]$Appearance,
        [Switch]$Versioning,
        [Switch]$Navigation,
        [Switch]$ShowInfo,
        [Switch]$ShowErrors
    )

    Begin {

        If (-Not (Test-Tenant)) { Return }
        If ($All -Or $Admins) { $GlobalAdmin = "Administradores Globais"; $OtherAdmins = $Null }
        If ($All -Or $Appearance) { $TenantTheme = ConvertFrom-Json $Global:CurrentTenant.Theme -AsHashtable }
        If ($All -Or $Navigation) { $EventsList = $Global:CurrentTenant.EventsID }

    }

    Process {

        # Check Home Site
        $IsHome = $Site.Url.Replace("/", "").EndsWith(".sharepoint.com")

        # Connect to Site
        $Connection = Invoke-Operation -Message "Connecting to site: $($Site.Title)" -ShowInfo:$ShowInfo -ShowErrors:$ShowErrors -ReturnValue -Operation {

            Connect-Site -Site $Site -Silent -ReturnConnection

        }; $Operations = 0

        # Set Site Admin
        If ($All -Or $Admins) {

            Invoke-Operation -Message "Setting site administrators" -ShowInfo:$ShowInfo -ShowErrors:$ShowErrors -Operation {

                $SiteAdmins = Get-PnPSiteCollectionAdmin -Connection $Connection
                Add-PnPSiteCollectionAdmin -Owners $OtherAdmins -PrimarySiteCollectionAdmin $GlobalAdmin -Connection $Connection
                $SiteAdmins | Where-Object Title -NE $GlobalAdmin | Where-Object LoginName -NotIn ($OtherAdmins) | Remove-PnPSiteCollectionAdmin -Connection $Connection

            }; $Operations++

        }

        # Set Versioning
        If ($All -Or $Versioning) {

            Invoke-Operation -Message "Setting site versioning" -ShowInfo:$ShowInfo -ShowErrors:$ShowErrors -Operation {

                $Status = (Get-PnPSiteVersionPolicyStatus).Status

                If ($Status -Ne "New") {

                    Set-PnPSiteVersionPolicy -EnableAutoExpirationVersionTrim $True -ApplyToNewDocumentLibraries -ApplyToExistingDocumentLibraries -Connection $Connection
                    Set-PnPSiteVersionPolicy -InheritFromTenant -Connection $Connection

                }
                
            }; $Operations++

        }

        # Set Site Appearance
        If ($All -Or $Appearance) {

            Invoke-Operation -Message "Setting site appearance" -ShowInfo:$ShowInfo -ShowErrors:$ShowErrors -Operation {

                Remove-PnPTenantTheme -Identity $TenantTheme.name -Connection $Connection
                Add-PnPTenantTheme -Identity $TenantTheme.name -Palette $TenantTheme.palette -IsInverted $TenantTheme.isInverted -Overwrite -Connection $Connection
                Set-PnPWebTheme -Theme $TenantTheme.name -Connection $Connection

                Set-PnPWebHeader -HeaderLayout "Standard" -HeaderEmphasis "None" -HideTitleInHeader:$False -HeaderBackgroundImageUrl $Null -LogoAlignment Left -Connection $Connection
                Set-PnPFooter -Enabled:$False -Layout "Simple" -BackgroundTheme "Neutral" -Title $Null -LogoUrl $Null -Connection $Connection
                
            }; $Operations++

        }

        # SharePoint Home Site Configuration
        If ($Site.Template -Match "SitePage" -And $IsHome) {

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

            If ($All -Or $Navigation) { 

                #Add-PnPNavigationNode -Title "Wiki" -Location "QuickLaunch" -Url "wiki/" ## =========================================== ##
            
            }

        }
        
        # SharePoint Other Sites Configuration
        If ($Site.Template -Match "SitePage" -And -Not $IsHome) {

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
        
        # OneDrive Sites Configuration
        If ($Site.Template -Match "SpsPers" -And $Site.Url -Match "/personal/") {

            $SiteParams = @{
                DefaultLinkPermission                       = "View"
                DefaultLinkToExistingAccess                 = $False
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

        # Teams Sites Configuration
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

        # Channels Sites Configuration
        If ($Site.Template -Match "TeamChannel") {

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

        # Set Site Params
        If ($All -Or $Params -Or $Operations -Eq 0) {

            Invoke-Operation -Message "Setting site parameters" -ShowInfo:$ShowInfo -ShowErrors:$ShowErrors -Operation {

                Set-PnPTenantSite -Identity $Site.Url @SiteParams -Connection $Connection
                Disable-PnPSharingForNonOwnersOfSite -Identity $Site.Url -Connection $Connection

            }; $Operations++

        }
    
        # Line Break
        Write-Host ""

    }

}

# Get-PnPWeb => Site Configs
# Get-PnPSubWeb => Subsite Configs
# Get-PnPTenantCdnEnabled -CdnType Public
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
    $Sites = Get-Sites -SharePoint
    $Sites | Set-Site -All

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
        $DateFormat = '{"$schema":"https://developer.microsoft.com/json-schemas/App/v2/column-formatting.schema.json","elmType":"div","children":[{"elmType":"Appan","style":{"padding-right":"10px","padding-bottom":"2px","font-size":"16px"},"attributes":{"iconName":"Calendar"}},{"elmType":"Appan","txtContent":"@currentField.diApplayValue","style":{"padding-bottom":"4px"}}]}'
        $PersonFormat = '{"$schema":"https://developer.microsoft.com/json-schemas/App/v2/column-formatting.schema.json","elmType":"div","style":{"diApplay":"flex","flex-wrap":"wrap","overflow":"hidden"},"children":[{"elmType":"div","forEach":"person in @currentField","defaultHoverField":"[$person]","attributes":{"class":"ms-bgColor-neutralLighter ms-fontColor-neutralSecondary"},"style":{"diApplay":"flex","overflow":"hidden","align-items":"center","border-radius":"28px","margin":"4px 8px 4px 0px","min-width":"28px","height":"28px"},"children":[{"elmType":"img","attributes":{"src":"=''/_layouts/15/userphoto.aAppx?size=S&accountname='' + [$person.email]","title":"[$person.title]"},"style":{"width":"28px","height":"28px","diApplay":"block","border-radius":"50%"}},{"elmType":"div","style":{"overflow":"hidden","white-Appace":"nowrap","text-overflow":"ellipsis","padding":"0px 12px 2px 6px","diApplay":"=if(length(@currentField) > 1, ''none'', ''flex'')","flex-direction":"column"},"children":[{"elmType":"Appan","txtContent":"[$person.title]","style":{"diApplay":"inline","overflow":"hidden","white-Appace":"nowrap","text-overflow":"ellipsis","font-size":"12px","height":"15px"}},{"elmType":"Appan","txtContent":"[$person.department]","style":{"diApplay":"inline","overflow":"hidden","white-Appace":"nowrap","text-overflow":"ellipsis","font-size":"9px"}}]}]}]}'

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
            
            $ViewFields = @("DocIcon", "LinkFilename", "Author", "Created", "Editor", "Modified", "FileSizeDiApplay")
            $ColumnWidth = '<FieldRef Name="Nome" width="500" /><FieldRef Name="Modificado Por" width="200" /><FieldRef Name="Modificado Em" width="200" /><FieldRef Name="Criado Por" width="200" /><FieldRef Name="Criado Em" width="200" /><FieldRef Name="Tamanho do Arquivo" width="200" />'
        
        } Else {
            
            $ViewFields = $View.ViewFields
            $ColumnWidth = $View.ColumnWidth
        
        }

        Set-PnPView -Identity $View.Id -Values @{ ColumnWidth = $ColumnWidth; } -List $List.Id -Fields $ViewFields

    }

    # <FieldRef Name="Modificado Por" width="229" /><FieldRef Name="Nome" width="418" /><FieldRef Name="Modificado Em" width="188" />

}