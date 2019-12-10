﻿function Get-Path{ return $(Get-Location).Path }


function Set-Log([string]$Message, [string]$Type, [switch]$NoNewline, [switch]$NoDisplay, [switch]$NoLog, [switch]$Force){
	$ScriptPath = Get-Path
	if(!$ScriptName){ $ScriptName = "tmp.ps1" }
	$LogName = $ScriptName.Replace(".ps1",".log")
	$LogFile = Join-Path $ScriptPath $LogName
	
	if(Test-Path $LogFile){
		$size = $(Get-ChildItem $LogFile).Length
		if($size -ge 1048576){ Remove-Item $LogFile -Force -Confirm:$false }
	}
	
	switch($Type){
		"Info" { $color = "White" }
		"Clean" { $color = "Green" }
		"Warning" { $color = "Yellow" }
		"Error" { $color = "Red" }
		default { $color = "White"; $Type = "Info" }
	}
	
	if(!$NoDisplay -and !$Quiet -and !$Force){
		if(!$NoNewline){ Set-Log $Message -ForegroundColor $Color }
		else{ Set-Log $Message -ForegroundColor $Color -NoNewline }
	}
	
	if(!$NoLog){
		$date = Get-Date -Format "dd/MM/yyyy hh:mm:ss"
		$line = "$date [$Type] $Message"
		Out-File $LogFile -Encoding UTF8 -Append -InputObject $line
	}
}




Import-Module ActiveDirectory

# users à exclure
$UsersToExclude = "campiglios","ccampiglio"
# chemin des profils utilisateurs dans le registre
$Key = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
# récupération liste utilisateurs
$Sub = Get-ChildItem $Key
# Sessions utilisateurs actives
$ActiveUsers = qwinsta.exe | ? { $_ -like "*Actif*" } | % {$_.split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)[1]}
$Global:ScriptPath = Get-Path
$global:ScriptName = $MyInvocation.MyCommand.name








ForEach($S in $Sub){
	$Items = Get-ItemProperty -Path $S.PSPath
	$FullPath = $items.ProfileImagePath
	
    # récupération SID
	$SID = $s.Name.Split("\")[6]
	
    # vérification du compte AD
	$ADUser = Get-ADUser -Identity $SID
	if($ADUser){
		$Name = $ADUser.SamAccountName
		
		# test si session utilisateur actives
		if($ActiveUsers -notcontains $Name){
		
			# vérification si le compte AD n'est pas exclu
			if($UsersToExclude -notcontains $Name){
			
				Set-Log "$ADUser.Name | SID  $ADuser.SID"
				# $Size = $(Get-ChildItem $FullPath -Recurse -Force | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1MB
				$Size = 0
				$SubFolder = Get-ChildItem $FullPath -Force | Where-Object { $_.psiscontainer -eq $true }
				ForEach($Folder in $SubFolder){
					Set-Location $Folder
					$Location = Get-Location
					if($Location.Path -eq $Folder.FullName){
						$TempSize = $(Get-ChildItem $Location -Recurse -Force | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1MB
						$Size += $TempSize
						Set-Log -NoDisplay "Size : $Size | TempSize : $TempSize"
					}
					else{ Set-Error "Impossible de se positionner dans le répertoire ($Folder)" }
				}

				# test de la taille du dossier
				if($Size -ge 200){
					Set-Log "Folder : $FullPath | Size : $Size"
					
					# Suppression clé de registre et dossier utilisateur
					#Remove-Item -Recurse -Force -Path $FullPath
					#Remove-Item $S -Force
				}
				else{ Set-Log "Taille inférieure à 200 Mo : $FullPath | $Size" }
			}
			else{ Set-Log "Compte exclu : $($ADUser.Name) | SID $($ADuser.SID)" }
		}
		else{ Set-Log "Utilisateur connecté : $Name" }
	}
	else{ Set-Log "Compte non trouvé dans l'AD" }
}

#####################################################################

# Pour tester la boucle ligne 9 renseigner la valeur dans $s = $sub[0]
# Décommenter les Remove-Item