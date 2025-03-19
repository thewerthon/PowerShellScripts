# Install: https://pnp.github.io/powershell/articles/installation.html
# Register: https://pnp.github.io/powershell/articles/registerapplication.html

# Tenant Variables
$AdminUrl = "https://siwindbr-admin.sharepoint.com"
$SitesUrls = @("https://siwindbr.sharepoint.com")
$ClientID = "91aac6c3-b063-4175-8073-7e5b5a4ff281" #GC
$ClientID = "7735abc1-32a8-416b-a7be-3d2496ba4724" #IT
$ClientID = "8b14c0ea-5f50-4c5c-b2f8-a50b5ca20d8b" #SIW

# Function to Connect Site
Function Connect-Site {

    Param (
        [Parameter(Mandatory = $True)]
        [String]$SiteUrl
    )

    $ClientID = "8b14c0ea-5f50-4c5c-b2f8-a50b5ca20d8b"
    Connect-PnPOnline $SiteUrl -ClientId $ClientID -OSLogin -PersistLogin -WarningAction Ignore

}

# Function to Disconnect Site
Function Disconnect-Site {

    Disconnect-PnPOnline -ClearPersistedLogin

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

# Function to Set Site
Function Set-Site {

    Param(
        [Parameter(Mandatory = $True)]
        [String]$SiteUrl
    )

    # Connect
    Connect-Site -SiteUrl $SiteUrl
    
    # Set Lists
    Get-Lists | Set-List

    # Disconnect
    Disconnect-Site

}