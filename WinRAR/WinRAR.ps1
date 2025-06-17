# Variables
$WinRar = "C:\Program Files\WinRAR\Rar.exe"

# Test WinRar Path
Function Test-WinRarPath {

    If (-Not (Test-Path $WinRar)) { 
        
        Throw "WinRAR executable not found at '$WinRar'"
    
    }

}

# Test WinRar Origin
Function Test-WinRarOrigin {

    Param(
        [Parameter(Mandatory = $True)][String]$Origin
    )

    If (-Not (Test-Path $Origin)) {
        
        Throw "Origin '$Origin' does not exist."
    
    }

}

# Compress WinRar Archive
Function Compress-WinRarArchive {

    [Alias("rar", "compress")]

    Param(
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)][Alias("o")][String]$Origin,
        [Parameter(Mandatory = $False, ValueFromPipeline = $False)][Alias("d")][String]$Destination,
        [Alias("r")][Switch]$RemoveAfter
    )

    Begin {

        Test-WinRarPath
        $Flags = If ($RemoveAfter) { "-df" }

    }
    
    Process {

        Test-WinRarOrigin $Origin
        $Item = Get-Item $Origin

        If ($Destination) {

            $Dest = $Destination
            
        } Else {

            $Name = $Item.BaseName
            $Path = If ($Item.PSIsContainer) { $Item.Parent.FullName } Else { $Item.DirectoryName }
            $Dest = Join-Path $Path "$Name.rar"

        }

        If ($Item.PSIsContainer) {

            Push-Location $Item.FullName
            & "$WinRar" a -ep1 -r -m5 $Flags "$Dest" *
            Pop-Location

            If ($RemoveAfter) {
                
                Start-Sleep -Milliseconds 250
                Remove-Item $Item.FullName -Recurse -Force
            
            }

        } Else {

            & "$WinRar" a -ep1 -m5 $Flags "$Dest" "$($Item.FullName)"

        }

        Write-Host
    
    }

}
