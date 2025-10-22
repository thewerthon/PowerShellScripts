# Variables
$WinRar = "C:\Program Files\WinRAR\Rar.exe"

# Test WinRar Path
function Test-WinRarPath {

	if (-not (Test-Path $WinRar)) {

		throw "WinRAR executable not found at '$WinRar'"

	}

}

# Test WinRar Origin
function Test-WinRarOrigin {

	param(
		[Parameter(Mandatory = $True)][String]$Origin
	)

	if (-not (Test-Path $Origin)) {

		throw "Origin '$Origin' does not exist."

	}

}

# Compress WinRar Archive
function Compress-WinRarArchive {

	[Alias("rar", "compress")]

	param(
		[Parameter(Mandatory = $True, ValueFromPipeline = $True)][Alias("o")][String]$Origin,
		[Parameter(Mandatory = $False, ValueFromPipeline = $False)][Alias("d")][String]$Destination,
		[Alias("r")][Switch]$RemoveAfter
	)

	begin {

		Test-WinRarPath
		$Flags = if ($RemoveAfter) { "-df" }

	}

	process {

		Test-WinRarOrigin $Origin
		$Item = Get-Item $Origin

		if ($Destination) {

			$Dest = $Destination

		} else {

			$Name = $Item.BaseName
			$Path = if ($Item.PSIsContainer) { $Item.Parent.FullName } else { $Item.DirectoryName }
			$Dest = Join-Path $Path "$Name.rar"

		}

		if ($Item.PSIsContainer) {

			Push-Location $Item.FullName
			& "$WinRar" a -ep1 -r -m5 $Flags "$Dest" *
			Pop-Location

			if ($RemoveAfter) {

				Start-Sleep -Milliseconds 250
				Remove-Item $Item.FullName -Recurse -Force

			}

		} else {

			& "$WinRar" a -ep1 -m5 $Flags "$Dest" "$($Item.FullName)"

		}

		Write-Host

	}

}
