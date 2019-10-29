function Get-Count([object]$Object){ return $($Object | Measure-Object).Count }

function Get-Path{ return $(Get-Location).Path }

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
		if(!$NoNewline){ Write-Host $Message -ForegroundColor $Color }
		else{ Write-Host $Message -ForegroundColor $Color -NoNewline }
	}
	
	if(!$NoLog){
		$date = Get-Date -Format "dd/MM/yyyy hh:mm:ss"
		$line = "$date [$Type] $Message"
		Out-File $LogFile -Encoding UTF8 -Append -InputObject $line
	}
}

function Set-Error([string]$Message, [switch]$Exit){
	$global:GlobalError = $Message
	$global:StillActive = $false
	
	if($Message){
		Set-Log "Error" "Error" -NoNewline
		Set-Log " : $Message"
	}
	else{ Set-Log "Error" "Error" }
	
	if($Exit){
		if($WorkBook){ Unload-ExcelFile $WorkBook }
		Set-Log "Arrêt du script"
		exit
	}
}

function Set-Clean([string]$Message){
	if($Message){
		Set-Log "Ok" "Clean" -NoNewline
		Set-Log " : $Message"
	}
	else{ Set-Log "Ok" "Clean" }
}

function Set-Warning([string]$Message){
	if($Message){
		Set-Log "Warning" "Warning" -NoNewline
		Set-Log " : $Message"
	}
	else{ Set-Log "Warning" "Warning" }
}

function Load-File([string]$FilePath){
	Set-Log "Chargement du fichier ... 			" -NoNewline
	if(Test-Path -LiteralPath $FilePath){
		$data = Get-Content -LiteralPath $FilePath -ea $QuietOutput -wa $QuietOutput
		if($data){
			$count = $($data | Measure-Object).Count
			if($count -gt 0){
				Set-Log "Ok" "Clean" -NoNewline
				Set-Log " : $count entrée(s) chargée(s)"
				Return $data
			}
			else{ Set-Error "Aucune entrée chargée" -Exit }
		}
		else{ Set-Error "Problème lors du chargement du fichier" -Exit }
	}
	else{ Set-Error "Fichier introuvable ou inaccessible" -Exit }
}

function Load-ExcelFile([string]$FilePath, [switch]$ReadOnly){
	Set-Log "Chargement du fichier Excel ...			" -NoNewline
	if(Test-Path -LiteralPath $FilePath){
		$global:ExcelApp = New-Object -ComObject Excel.Application
		if($ExcelApp){
            $ExcelApp.Visible = $false
            if($ReadOnly){ $WorkBook = $ExcelApp.Workbooks.open($FilePath, $false, $true) }
			else{ $WorkBook = $ExcelApp.Workbooks.open($FilePath) }
			if($WorkBook){
				Set-Log "Ok" "Clean"
				Return $WorkBook
			}
			else{ Set-Error "Impossible de charger le fichier" -Exit }
		}
		else{ Set-Error "Impossible de charger l'objet Com Excel" -Exit }
	}
	else{ Set-Error "Fichier introuvable ou inaccessible" -Exit }
}

function Get-ExcelSheet([object]$WorkBook,[string]$SheetName){
	$WorkSheet = $WorkBook.WorkSheets | Where-Object { $_.Name -eq $SheetName }
	if($WorkSheet){ Return $WorkSheet }
}

function New-ExcelSheet([object]$WorkBook,[string]$SheetName){
	Set-Log "Création d'une nouvelle feuille Excel ...	" -NoNewline
	$WorkSheet = $Workbook.Worksheets.Add([System.Reflection.Missing]::Value,$WorkBook.Worksheets.Item($WorkBook.Worksheets.count))
	if($WorkSheet){
		$WorkSheet.Name = $SheetName
		Set-Log "Ok" "Clean"
		Return $WorkSheet
	}
	else{ Set-Error "Impossible de créer une nouvelle feuille" }
}

function Remove-ExcelSheet([object]$WorkBook,[string]$SheetName){
	Set-Log "Suppression d'une feuille Excel ...		" -NoNewline
	$WorkSheet = $WorkBook.Worksheets.Item($SheetName)
	if($WorkSheet){
		$WorkSheet.Delete()
		try{ $WorkSheetTest = $WorkBook.Worksheets.Item($SheetName) }
		catch{ }
		if(!$WorkSheetTest){ Set-Log "Ok" "Clean" }
		else{ Set-Error "Impossible de supprimer la feuille" }
	}
	else{ Set-Error "Feuille introuvable" }
}

function Unload-ExcelFile([object]$WorkBook, [switch]$Save){
	Set-Log "Déchargement du fichier Excel ...		" -NoNewline
	$try = 0
	$PidComObject = Get-PIDComObject $ExcelApp
	
	if($Save){ $WorkBook.Save() }
	$WorkBook.Close()
	$ExcelApp.Quit()
	
	while([System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelApp) -gt 0 -and $try -lt 5){ $try++ }
	[System.GC]::Collect()
	
	Start-Sleep 1
	$Process = Get-Process | ? { $_.Id -eq $PidComObject }
	if(!$Process){ Set-Log "Ok" "Clean" }
	else{ Set-Error "Impossible de décharger correctement les objets COM" }
}

function New-OutlookMail([string[]]$To, [string[]]$Cc, [string[]]$Cci, [string]$Subject, [string]$Body, [switch]$Activate){
	$OutlookApp = New-Object -ComObject Outlook.Application
	if($OutlookApp){
		$Mail = $OutlookApp.CreateItem(0)
		$Mail.Subject = $Subject
		$Mail.Body = $Body
		
		if($Activate){
			$Inspector = $Mail.GetInspector
			$Inspector.Activate()		
		}
		
		if($To){ foreach($dst in $To){ $Mail.Recipients.Add($dst) | Out-Null } }
		if($Cc){
			$tmp = $null
			foreach($dst in $Cc){ $tmp += $dst + "; " }
			$Mail.Cc = $tmp
		}
		if($Cci){
			$tmp = $null
			foreach($dst in $Cci){ $tmp += $dst + "; " }
			$Mail.BCC = $tmp
		}
		
		$Mail.Save()
	}
	else{ Set-Error "Impossible de charger l'objet Com Outlook" -Exit }
}

function Enable-PSSnapin([string]$Snapin, [string]$SnapinName){
	Set-Log "Chargement du Snapin $SnapinName ... 		" -NoNewline
	$InitPSSnapin = Get-PSSnapin $Snapin -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
	if(!$InitPSSnapin){
		$RegistPSSnapin = Get-PSSnapin -Registered | Where { $_.Name -eq $Snapin }
		if($RegistPSSnapin){
			$ChkPSSnapin = Add-PSSnapin $Snapin -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -PassThru
			if($ChkPSSnapin){ Set-Log "Ok" "Clean" }
			else{ Set-Error "Impossible de charger le snapin $SnapinName" -Exit }
		}
		else{ Set-Error "Impossible de charger le snapin $SnapinName : Snapin non enregistré sur cette machine" -Exit }
	}
	else{ Set-Log "Ok" "Clean" }
}

function Enable-Module([string]$Module){
	Set-Log "Chargement du Snapin $Module ... 		" -NoNewline
	$InitModule = Get-Module $Module -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
	if(!$InitModule){
		$RegistModule = Get-Module -ListAvailable | Where { $_.Name -eq $Module }
		if($RegistModule){
			$ChkModule = Import-Module $Module -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -PassThru
			if($ChkModule){ Set-Log "Ok" "Clean" }
			else{ Set-Error "Impossible de charger le snapin $Module" -Exit }
		}
		else{ Set-Error "Impossible de charger le snapin $Module : Module non enregistré sur cette machine" -Exit }
	}
	else{ Set-Log "Ok" "Clean" }
}

function Load-PSSnapin([switch]$VMwareCore, [switch]$VMwareVDS){
	if($VMwareCore){ Enable-PSSnapin "VMware.VimAutomation.Core" "VMware Core" }
	if($VMwareVDS){ Enable-PSSnapin "VMware.VimAutomation.Vds" "VMware VDS" }	
}

function Load-Module([switch]$ActiveDirectory){
	if($ActiveDirectory){ Enable-Module "ActiveDirectory" }
}

function Connect-vCenterServer([string]$vCenterServer){
	Set-Log "Connexion au vCenter ... 			" -NoNewline
	if(Test-Connection $vCenterServer -Quiet -Count 1){
		$Connection = Connect-VIServer -Server $vCenterServer -ea $QuietOutput -wa $QuietOutput
		if($Connection){ Set-Log "Ok" "Clean" }
		else{ Set-Error "Connexion impossible au serveur vCenter" -Exit }
	}
	else{ Set-Error "Connexion impossible au serveur vCenter : serveur injoignable" -Exit }
}

function Disconnect-vCenterServer([string]$vCenterServer){
	if($vCenterServer){
		Set-Log "Déconnexion du vCenter ... 			" -NoNewline
		Disconnect-VIServer -Server $vCenterServer -Confirm:$false -Force -ea $QuietOutput -wa $QuietOutput
		if(!($global:DefaultVIServers -match $vCenterServer)){ Set-Log "Ok" "Clean" }
		else{ Set-Error "Impossible de se déconnecter du vCenter" }
	}
	else{
		Set-Log "Déconnexion des vCenters ... 			" -NoNewline
		Disconnect-VIServer -Confirm:$false -Force -ea $QuietOutput -wa $QuietOutput
		if(!($global:DefaultVIServers)){ Set-Log "Ok" "Clean" }
		else{ Set-Error "Impossible de se déconnecter de tous les vCenters" }
	}
}

function Invoke-RemoteCmd([string]$Command, [string]$Argument, [string]$Server){
	$WinFolder = Get-ChildItem "C:\Windows"
	$TestPSExec = $WinFolder | Where-Object { $_.Name -eq "psexec.exe" }
	
	if($TestPSExec){ 
		$Result = psexec \\$Server "$Command" "$Argument"
	}
	else{ $Result = "Error : PSExec introuvable" }
	
	Return $Result
}

function Get-RemoteService([string]$Name, [string]$Server, [System.Management.Automation.PSCredential]$Credential){
	$Service = Get-WmiObject -ComputerName $Server -Class Win32_Service -Filter "Name='$Name'" -Credential $Credential
	Return $Service
}

function Set-RemoteService{
	Param(
		[Parameter(Mandatory=$true,Position=0)]
		[string]$Name,
		[Parameter(Mandatory=$true,Position=1)]
		[string]$StartupType,
		[Parameter(Mandatory=$true,Position=2)]
		[string]$Server,
		[Parameter(Mandatory=$true,Position=3)]		
		[System.Management.Automation.PSCredential]$Credential
	)
	
	$array = "Automatic","Manual","Disabled"
	if($array -match $StartupType){
		Set-Log "Changement du mode de démarrage du service ... 	" -NoNewline
		
		$Service = Get-RemoteService -Name $Name -Server $Server -Credential $Credential
		if($Service){
			$Result = $Service.Change($null,$null,$null,$null,$StartupType)
			$Service = Get-RemoteService -Name $Name -Server $Server -Credential $Credential
			if($Service.StartMode -eq $StartupType){ Set-Log "Ok" "Clean"; $res = $true }
			else{ Set-Error "Erreur lors du changement de mode de démarrage"; $res = $false }
		}
		else{ Set-Error "Service introuvable ou erreur lors de la récupération" -Exit }
	}
	else{ Set-Error "Argument StartupType incorrect (Automatic, Manual, Disabled)" -Exit }
	return $res
}

function Stop-RemoteService([string]$Name, [string]$DisplayName, [string]$Server, [System.Management.Automation.PSCredential]$Credential){
	Set-Log "Arrêt du service $DisplayName ... 		" -NoNewline
	$Service = Get-RemoteService -Name $Name -Server $Server -Credential $Credential
	if($Service){
		if($Service.State -eq "Running"){
		
			$Service.StopService() | Out-Null
			$try = 0
			
			while($try -lt 10){
				$Service = Get-RemoteService -Name $Name -Server $Server -Credential $Credential
				if($Service.State -eq "Stopped"){ Set-Log "Ok" "Clean"; $try = 10; $res = $true }
				else{ $try++ }
				Start-Sleep 1
			}
			
			if(!$res){ Set-Error "Le service ne s'est pas arrêté après 10 secondes"; $res = $false } 
		}
		elseif($Service.State -eq "Stopped"){ Set-Log "Ok" "Clean"; $res = $true }
		else{ Set-Error "Etat du service non pris  en compte : $Service.State"; $res = $false }
	}
	else{ Set-Error "Service introuvable ou erreur lors de la récupération"; $res = $false }
	return $res
}

function Get-OSVersion([string]$Server){
	$Result = $false
	$OS = Get-WmiObject -ComputerName $Server -Class Win32_OperatingSystem
	
	if($OS){
		$MajorVersion = $OS.Version.Substring(0,3)
		
		switch($MajorVersion){
			"5.2" { $OSVersion = "Windows Server 2003" }
			"6.0" { $OSVersion = "Windows Server 2008" }
			"6.1" { $OSVersion = "Windows Server 2008 R2" }
			"6.2" { $OSVersion = "Windows Server 2012" }
			"6.3" { $OSVersion = "Windows Server 2012 R2" }
			default { $OSVersion = "Unknown" }
		}
		
		if($OSVersion -ne "Unknown"){ $Result = $OSVersion }
		else{ Set-Error "OS inconnu" }
	}
	else{ Set-Error "La requête WMI n'a retournée aucun résultat" }
	
	return $Result
}

function Get-PIDComObject([object]$ComObject){
	Add-Type -TypeDefinition @"
		using System;
		using System.Runtime.InteropServices;

		public static class Win32Api{
			[System.Runtime.InteropServices.DllImportAttribute( "User32.dll", EntryPoint =  "GetWindowThreadProcessId" )]
			public static extern int GetWindowThreadProcessId( [System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd, out int lpdwProcessId );

			[DllImport("User32.dll", CharSet = CharSet.Auto)]
			public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
		}
"@

	$HWND = $ComObject.Hwnd
	$PidComObject = [IntPtr]::Zero
	[Win32Api]::GetWindowThreadProcessId($HWND, [ref]$PidComObject) | Out-Null
	Return $PidComObject
}

function Test-ComputerAccount([string]$Server){
	$Test = Get-ADComputer $Server -ea SilentlyContinue
	if($Test -ne $null){ $Result = $true }
	else{ $Result = $false }
	
	return $Result
}

function Test-Credential([System.Management.Automation.PSCredential]$Credential, [string]$Domain){
	$UserName = $Credential.UserName
	$Password = $Credential.GetNetworkCredential().Password

	if(!$Domain){ $Domain = ([ADSI]"").distinguishedName }
	$CurrentDomain = "LDAP://" + $Domain
	$Authent = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain, $UserName, $Password)

	if($Authent.name -eq $null){ Set-Error "Domaine, utilisateur ou mot de passe incorrect" -Exit }
	else{ Set-Log "Ok" "Clean" }
}

function Test-TCPPort([string]$Server, [int]$Port){
	$TimeOut = 1000
	$Result = $false
	$Socket = New-Object System.Net.Sockets.TCPClient
	$Connect = $Socket.BeginConnect($Server,$Port,$null,$null)
	
	if($Connect.IsCompleted){
		$Wait = $Connect.AsyncWaitHandle.WaitOne($TimeOut,$false)			
		if(!$Wait){
			$Socket.Close() 
			$Result = $false 
		}
		else{
			$Socket.EndConnect($Connect)
			$Socket.Close()
			$Result = $true
		}
	}
	else{ $Result = $false }
	
	return $Result
}

function Test-WMIAccess([string]$Server, [System.Management.Automation.PSCredential]$Credential){
	$result = $false
	
	try{
		if($Credential){ $WMI = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Server -Credential $Credential}
		else{ $WMI = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Server }
	}
	catch{ $result = $false }
	
	if($WMI){ $result = $true }
	
	return $result
}

function Out-ClipBoard{
    param(
        [Parameter(ValueFromPipeline = $true)]
        [object[]]$InputObject
    )

	$objectsToProcess += $InputObject

	$objectsToProcess | PowerShell -NoProfile -STA -Command {
		Add-Type -Assembly PresentationCore
		$clipText = ($input | Out-String -Stream) -join "`r`n"
		[Windows.Clipboard]::SetText($clipText)
	}
}

function Invoke-SqlCommand([string]$DBServer, [string]$DBName, [string]$DBUser, [string]$DBPassword, [string]$Query){
	$ConnectionString = "Server=$DBServer;Database=$DBName;uid=$DBUser;pwd=$DBPassword"
	$Connection = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
	$Connection.Open()
	
	$Command = New-Object System.Data.SqlClient.SqlCommand 
	$Command.Connection = $Connection
	$Command.CommandText = $Query
	
	$Adapter = New-Object System.Data.SqlClient.SQLDataAdapter($Command)
	$Dataset = New-Object System.Data.DataSet
	$Adapter.Fill($Dataset) | Out-Null
	
	$Dataset.Tables | Select-Object -Expand Rows
}

function Get-ScheduleTask{
	Param(
		[Parameter(Mandatory=$false,Position=0)]
		[string]$TaskName,
		[Parameter(Mandatory=$false,Position=1)]
		[string]$ComputerName
	)

	$Object = New-Object -ComObject Schedule.Service
	
	if($ComputerName){ $Object.Connect($ComputerName) }
	else{ $Object.Connect($env:COMPUTERNAME) }
	
	$TaskFolder = $Object.GetFolder("\")
	$TaskList = $TaskFolder.GetTasks(1)
	
	if($Name){
		foreach($Task in $TaskList){
			if($Task.Name -eq $Name){
				$Result = $true
				break
			}
		}
	}
	else{
		$Result = @()
		foreach($Task in $TaskList){ $Result += $Task }
	}
	
	return $Result
}

function Create-ScheduleTask{
	Param(
		[Parameter(Mandatory=$true,Position=0)]
		[string]$TaskName,
		[Parameter(Mandatory=$false,Position=1)]
		[string]$Path,
		[Parameter(Mandatory=$false,Position=2)]
		[string]$ArgumentList,
		[Parameter(Mandatory=$false,Position=3)]
		[string]$WorkingDirectory,
		[Parameter(Mandatory=$false,Position=4)]
		[string]$ComputerName
	)

	$Object = New-Object -ComObject Schedule.Service
	
	if($ComputerName){ $Object.Connect($ComputerName) }
	else{ $Object.Connect($env:COMPUTERNAME) }
	
	$RootFolder = $Object.GetFolder("\")
	$TaskDefinition = $Object.NewTask(0)
	
	$Settings = $TaskDefinition.Settings
	$Settings.Enabled = $True
	$Settings.StartWhenAvailable = $True
	$Settings.Hidden = $False
	
	$Triggers = $TaskDefinition.Triggers
	$Trigger = $Triggers.Create(7)

	$Action = $TaskDefinition.Actions.Create(0)
	$Action.Path = $Path
	$Action.Arguments = $ArgumentList
	$Action.WorkingDirectory = $WorkingDirectory
	
	$Task = $RootFolder.RegisterTaskDefinition($TaskName, $TaskDefinition, 2, "System", $null , 5)
	$CheckTask = Get-ScheduleTask -TaskName $TaskName -ComputerName $ComputerName
	
	if($CheckTask){ $CheckTask | Select-Object Name,Enabled,NextRunTime }
	else{ Write-Host "Erreur lors de la création de la tâche" -Foregroundcolor Red }
}

function Remove-ScheduleTask{
	Param(
		[Parameter(Mandatory=$true,Position=0)]
		[string]$TaskName,
		[Parameter(Mandatory=$false,Position=1)]
		[string]$ComputerName
	)

	$Object = New-Object -ComObject Schedule.Service
	
	if($ComputerName){ $Object.Connect($ComputerName) }
	else{ $Object.Connect($env:COMPUTERNAME) }
	
	$TaskFolder = $Object.GetFolder("\")
	$TaskList = $TaskFolder.GetTasks(1)
	
	foreach($Task in $TaskList){
		if($Task.Name -eq $Name){
			$TaskFolder.DeleteTask($Task.Name,0)
		}
	}
}

function Get-RemoteRegistry{
	Param(
		[Parameter(Mandatory=$true,Position=0)]
		[string]$KeyPath,
		[Parameter(Mandatory=$true,Position=1)]
		[string]$Value,
		[Parameter(Mandatory=$false,Position=2)]
		[string]$ComputerName
	)
	
	if(!$ComputerName){ $ComputerName = $env:COMPUTERNAME }
	
	$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
	$SubKey = $Reg.OpenSubKey($KeyPath)
	$Key = $SubKey.GetValue($Value)
	
	return $Key
}