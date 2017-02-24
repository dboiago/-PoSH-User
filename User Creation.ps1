<#
Description:	Script to create AD, Exchange(On-Prem), O365(Cloud/MSOL), Lync(Skype) profiles, assign a Long Distance code, populate their AD groups and create their desktop and home folder
				with some shortcuts. (comma before "and" is an oxford comma..not a typo)
Author:			Devon Boiago
Created:		28/01/14 (guessing since it was created before term)
Last Edit:		7/2/17

Notes:			Form buttons: 0: OK, 1: OK Cancel, 2: Abort Retry Ignore, 3: Yes No Cancel, 4: Yes No, 5: Retry Cancel			
		
Last Edit:		Updated the CreateADEmail0365 function to do a check for [Redacted] users and set their default SMTP address to @[Redacted].com


				Appended the AD user verification section to check if the user need a roaming profile or not ($RoamingProfile)
				Appended CreateADEmail0365 function for laptop(no roaming profile) users to not expire their password
				Appended the User Creation, Exchange and AD section to blank out the ProfilePath for laptop(no roaming profile) users
			
To do:			There is a number of places where something hard coded can/should be replaced with a variable; can create some arrays for those etc.; Possibly some general clean up as well
#>		

#region Initialization
Clear-Host
# Add the AD and Exchange Modules
Import-Module ActiveDirectory
Import-Module MSOnline
if((Get-PSSnapin | Where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.SnapIn"}) -eq $null) {Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;}
#Pop up windows
Add-Type –AssemblyName System.Windows.Forms
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') #This line is needed to run the VisualBasic dialog boxes.
#endregion

#region Variables and Arrays
#Account creation
$NewUser = @()
$LDCode = @()
#Errors
$Error.Clear()
#Parsing with excel
$ExcelPC = "EXCELPCNAME" #PC that has excel for parsing spreadsheets
$ExcelTemp = "\\SERVERNAME\DRIVE\FOLDER\" #Temp folder for parsing excel docs; Do NOT leave out the trailing slash
#region Remote Connectons - Would be great to pull these from an encrypted DB instead of having in here, but for now it's in here
#Remote connections
$Key = (0,0,0,0,00,00,000,000,0,0,0,00,00,00,00,000,0,00,0,0,0,0,00,00) #Random numbers to generate a key with for remoting
$SecurePasswordKey = "BIGSTRINGTHATSENCRYPTED="
$SessionUser = 'ACCOUNTNAME' #Account used to connect to remote PCs/Servers
$SessionPass =  ConvertTo-SecureString -String $SecurePasswordKey -Key $key #Password used to connect to remote PCs/Servers
$SessionCreds = New-Object System.Management.Automation.PSCredential($SessionUser,$SessionPass) #Name and Pass combined for remoting
#Lync remote connections
$LyncServer = "LYNCSERVERNAME"
$LyncUser = 'LYNCUSERNAME'
$LyncPass =  'LYNCPASSWORD' | ConvertTo-SecureString -AsPlainText -Force # This should be encrypted or in another file or a DB or just prompt as well
$LyncCreds = New-Object System.Management.Automation.PSCredential($LyncUser,$LyncPass)
#O365 remote connections
$o365key = (0,0,0,0,00,00,000,000,0,0,0,00,00,00,00,000,0,00,0,0,0,0,00,00)
$o365PasswordKey = "BIGSTRINGTHATSENCRYPTED="
$o365User = 'EMAILNAME@DOMAIN.DOMAINONMICROSOFT.com' #Account used to connect to remote PCs/Servers
$o365Pass =  ConvertTo-SecureString -String $o365PasswordKey -Key $o365key #Password used to connect to remote PCs/Servers
$o365creds = New-Object System.Management.Automation.PSCredential($o365User,$o365Pass) #Name and Pass combined for remoting
$Azurekey = (0,0,0,0,00,00,000,000,0,0,0,00,00,00,00,000,0,00,0,0,0,0,00,00)
$AzurePasswordKey = "BIGSTRINGTHATSENCRYPTED="
$AzureUser = 'AZUREUSERNAME'
$AzurePass =  ConvertTo-SecureString -String $AzurePasswordKey -Key $Azurekey
$AzureCreds = New-Object System.Management.Automation.PSCredential($AzureUser,$AzurePass)
#endregion
#E-mail
$EmailTo = "FIRST LAST <EMAIL@DOMAIN.com>", "FIRST LAST <EMAIL@DOMAIN.com>" #Email(s) to send messages to for completion, etc
$MailServer = "MAILSERVERNAME"
#Logging
$FileName = "User Creation "+(Get-Date -Format D)+".txt"
$LogPath = "\\SERVERNAME\DRIVE\FOLDER\$FileName"
#MSOL License (with O365 you need to explicitly pick what to remove, not what to add)
$Database = "MAILBOX DATABASE 0123456789"
$acctSku = "DOMAINISHNAME:PACKNAME"
$disabledPlans = @("SHAREPOINTENTERPRISE","SHAREPOINTWAC","INTUNE_O365","RMS_S_ENTERPRISE","MCOSTANDARD") #Examples
$LicenseOptions = New-MsolLicenseOptions -AccountSkuId $acctSku -DisabledPlans $disabledPlans
#$RemoteHost = "REMOTEHOSTSERVERNAME" #This was just left in as it is referenced, but not used any longer
#$DeliveryDomain = "DOMAIN.MAIL.DOMAINONMICROSOFT.com" #Same as above
#Script multi-instance protection - This is to prevent issues with Excel portions mainly
$LockName = ($Env:TEMP+"\"+$($Script:MyInvocation.MyCommand.Name)+"ScriptLock.lock")
#endregion

#region Functions
# Function to create a temporary file to (very basically) prevent running multiple instances of the same script
Function ScriptLock ($LockName=$LockName) {
	if (Test-Path ($LockName)) {
		if ([datetime](Get-ChildItem $LockName).creationtime -lt (Get-Date).addminutes(30)) {
			$pause = [Microsoft.VisualBasic.Interaction]::MsgBox("There is an iteration of the script currently running. The script will now exit. `n`n $(Get-Content $LockName)", 0, "Warning")
			exit
		}
		else { rm $LockName }
	}
	else { "$LockName locked down until $((Get-Date).addminutes(30)) or script completes - Safe to manually delete after this time" | Out-File $LockName }
}

# Function for logging
Function Logging([string]$ToLog, [string]$Type, [string]$LogPath=$LogPath) {
	try {
		switch ($Type) {
			"Error" {(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - ERROR - " + $Error[0].Exception.Message + " Line: " + $Error[0].InvocationInfo.ScriptLineNumber + " Char: " + $Error[0].InvocationInfo.OffsetInLine | Out-File -FilePath $LogPath -append }
			"Warning" {(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - WARNING - $ToLog" | Out-File -FilePath $LogPath -append}
			Default {(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - INFO - $ToLog" | Out-File -FilePath $LogPath -append}
		}
	}
	catch { 
		# Not sending to the Error Hanlding function since it write to log, which would write to errhandle, which would then try to log....infinitely
		# Not certain $EmailTo and $MailServer will be passed thru; may need to just create a pop up or something
		$Body = "Unable to write to log at $LogPath`n`n", $Error[0].Exception.Message, $Error[0].InvocationInfo.positionmessage
		Send-MailMessage -To $EmailTo -Subject "User Creation Error for Account: $FullName" -Body "$Body" -From "NOTIFICATIONNAME@DOMAIN.COM" -SmtpServer $MailServer
	}
}

# Function for error handling
Function ErrHandle([string]$MailServer=$MailServer, [string]$EmailTo=$EmailTo, [string]$ErrMsg, [string]$FullName, [string]$ExceptionMsg, [string]$InvocationInfo, [string]$ErrAct="exit", $LockName=$LockName) {
	$NoEcho = [System.Windows.Forms.MessageBox]::Show("$ErrMsg. `n $ExceptionMsg `n $InvocationInfo", "Error")
	$Body = "$($ErrMsg). Please review the log file at `"$LogPath`"`n`n", $ExceptionMsg, $InvocationInfo
	Send-MailMessage -To $EmailTo -Subject "User Creation Error for Account: $FullName" -Body "$Body" -From "NOTIFICATIONNAME@DOMAIN.COM" -SmtpServer $MailServer
	Logging -Type error -LogPath $LogPath
	if ($ErrAct -ne "continue") { rm $LockName; exit }
}

# Function for O365 license checking
Function O365LicCheck {
	try {
		if (($LicenseCheck | Where-Object {$_.SkuPartNumber -eq "ENTERPRISEPACK"}).ConsumedUnits -ge ($LicenseCheck | Where-Object {$_.SkuPartNumber -eq "ENTERPRISEPACK"}).ActiveUnits) {
			return "NoLicAvail"
		}
	}
	catch {  ErrHandle -ErrMsg "Unable to check for O365 licenses." -FullName $FullName -MailServer $MailServer -ExceptionMsg $_.Exception.Message -InvocationInfo $_.InvocationInfo.positionmessage }
}

# Function for user creation
Function CreateADEmail0365([string]$FullName, [string]$sAMName, [string]$FullOU, [string]$UserPrincipal, [string]$FirstName, [string]$LastName, [securestring]$Password, [string]$Database, [string]$ProfPath, [string]$ScriptPath, [string]$HomeDrive, [string]$HomeDir, [string]$Office, [string]$Street, [string]$City, [string]$State, [string]$Postal, [string]$Country, [string]$OU, [string]$Company, [Object]$o365creds, [string]$acctSku=$acctSku, $LockName=$LockName) {
	# Create the mailbox on-prem; will also create ad AD account
	try {
		#region Exchange/O365 mailbox creation
		# Don't need to do this on-prem step anymore since New-RemoteMailbox will create one on-prem as well
		#New-Mailbox -Name $FullName -Alias $sAMName -OrganizationalUnit $FullOU -UserPrincipalName $UserPrincipal -SamAccountName $sAMName -FirstName $FirstName -Initials '' -LastName $LastName -Password $Password -ResetPasswordOnNextLogon $true -Database $Database | Out-Null
		
		# XX (XX not the real Country) office cant reset their password so that part is false, for regular office users it needs to be changed upon next logon from the temp
		if ($Country -eq "XX") { 
			New-RemoteMailbox -Name $FullName -Alias $sAMName -OnPremisesOrganizationalUnit $FullOU -UserPrincipalName $UserPrincipal -SamAccountName $sAMName -FirstName $FirstName -Initials '' -LastName $LastName -Password $Password -ResetPasswordOnNextLogon $false | Out-Null
			sleep 5 #Give it a chance to create
		}
		# Laptop users password should not expire in case it expires while offsite; but they still need to set a proper password upon first login
		elseif ($ProfPath -eq "") {
			New-RemoteMailbox -Name $FullName -Alias $sAMName -OnPremisesOrganizationalUnit $FullOU -UserPrincipalName $UserPrincipal -SamAccountName $sAMName -FirstName $FirstName -Initials '' -LastName $LastName -Password $Password -ResetPasswordOnNextLogon $true | Out-Null
			sleep 5 #Give it a chance to create
		}
		else {
			New-RemoteMailbox -Name $FullName -Alias $sAMName -OnPremisesOrganizationalUnit $FullOU -UserPrincipalName $UserPrincipal -SamAccountName $sAMName -FirstName $FirstName -Initials '' -LastName $LastName -Password $Password -ResetPasswordOnNextLogon $true | Out-Null
			sleep 5 #Give it a chance to create
		}
		# Test to see if the Mailbox and User created
		try { 
			Get-ADUser -Filter "samaccountname -eq '$sAMName'" | Out-Null
			Logging -ToLog "The on-prem mailbox and AD account for $FullName has been created successfully" -LogPath $LogPath -Type info
		}
		catch { 
			ErrHandle -ErrMsg "User account creation was unsuccessful" -FullName $FullName -MailServer $MailServer -ExceptionMsg $Error[0].Exception.Message -InvocationInfo $Error[0].InvocationInfo.positionmessage
		}
		#endregion
		#region Additional AD settings
		# Set additional AD info that doesn't get applied initially
		 #because set-aduser HATES life and refuses a blank string in pretty much every way possible when included as a variable as though it will murder it's family (gives the error that just says 'replace') i'm just adding a check for ease instead of messing around with it as it wasn't working properly trying to assign it as $null, empty, whitespace, etc.
		if ($ProfPath -eq "none") { Set-ADUser -Identity $sAMName -ScriptPath $ScriptPath -HomeDrive $HomeDrive -HomeDirectory $HomeDir -Office $Office -StreetAddress $Street -City $City -State $State -PostalCode $Postal -Country $Country -Department $OU -Company $Company }
		else { Set-ADUser -Identity $sAMName -ProfilePath $ProfPath -ScriptPath $ScriptPath -HomeDrive $HomeDrive -HomeDirectory $HomeDir -Office $Office -StreetAddress $Street -City $City -State $State -PostalCode $Postal -Country $Country -Department $OU -Company $Company } #-Title $Role -Manager $Manager
		Logging -ToLog "Additional AD information for $FullName has been set" -LogPath $LogPath -Type info
		#endregion
		#region Sync AD/Exchange with O365
		try {
			Invoke-Command -ComputerName azu-w12s-dc01 -Authentication credssp -Credential $SessionCreds -ArgumentList $EmailTo, $MailServer, $FullName, $LogPath, $UserPrincipal, $LockName -ScriptBlock { 
			param($MailServer, $EmailTo, $FullName, $LogPath, $UserPrincipal, $LockName)
			#region Functions				
			# Function for error handling
			Function ErrHandle([string]$MailServer=$MailServer, $EmailTo=$EmailTo, [string]$ErrMsg, [string]$FullName, [string]$ExceptionMsg, [string]$InvocationInfo, [string]$ErrAct="exit", $LockName=$LockName) {
				Write-Host ("$ErrMsg. `n $ExceptionMsg `n $InvocationInfo", "Error") -ForegroundColor Magenta
				$Body = "$($ErrMsg). Please review the log file at `"$LogPath`"`n`n", $ExceptionMsg, $InvocationInfo
				Send-MailMessage -To $EmailTo -Subject "User Creation Error for Account: $FullName" -Body "$Body" -From "NOTIFICATIONNAME@DOMAIN.COM" -SmtpServer $MailServer
				Logging -Type error -LogPath $LogPath
				if ($ErrAct -ne "continue") { rm $LockName; exit }
			}
			# Function for logging
			Function Logging([string]$ToLog, [string]$Type, [string]$LogPath=$LogPath, $EmailTo=$EmailTo, $MailServer=$MailServer) {
				try {
					switch ($Type) {
						"Error" {(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - ERROR - " + $Error[0].Exception.Message + " Line: " + $Error[0].InvocationInfo.ScriptLineNumber + " Char: " + $Error[0].InvocationInfo.OffsetInLine | Out-File -FilePath $LogPath -append }
						"Warning" {(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - WARNING - $ToLog" | Out-File -FilePath $LogPath -append}
						Default {(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - INFO - $ToLog" | Out-File -FilePath $LogPath -append}
					}
				}
				catch { 
					# Not sending to the Error Hanlding function since it write to log, which would write to errhandle, which would then try to log....infinitely
					# Not certain $EmailTo and $MailServer will be passed thru; may need to just create a pop up or something
					$Body = "Unable to write to log at $LogPath`n`n", $Error[0].Exception.Message, $Error[0].InvocationInfo.positionmessage
					Send-MailMessage -To $EmailTo -Subject "User Creation Error for Account: $FullName" -Body "$Body" -From "NOTIFICATIONNAME@DOMAIN.COM" -SmtpServer $MailServer
				}
			}
			#endregion
				try {
					# Start the AD Sync with MSOL; Delta to only update changes since last sync (takes aprox 30 seconds as per a couple tests/checking logs)
					Write-Host "Running AD Sync Cycle, this will take a few minutes..." -ForegroundColor Cyan
					sleep 180
					Start-ADSyncSyncCycle -PolicyType Delta
					logging -ToLog "Successfully syncronised the cloud and AD accounts for $($UserPrincipal)" -Type Info -LogPath $LogPath
				}
				catch {
					logging -ToLog "An error occured attempting to syncronise the could and AD account for $($UserPrincipal)" -Type Error -LogPath $LogPath
				}
				sleep 90 #Allowing a minute and a half for the sync to go through; padding the time so can cut back if needed
			}
		}
		catch {
			logging -ToLog "An error occured attempting to syncronise the could and AD account for $($UserPrincipal)" -Type Error -LogPath $LogPath
		}	
		#endregion
		#region O365 license application
		Connect-MsolService -Credential $o365creds
		$LicenseCheck = Get-MsolAccountSku
		if (O365LicCheck -eq "NoLicAvail") {
			[System.Windows.Forms.MessageBox]::Show("There are no more O365 licenses available. Remove un-needed licenses or buy more. This user will be created without an O365 email address.","Error") | Out-Null
			Logging -ToLog "There are no more O365 licenses available. Remove un-needed licenses or buy more. This user will be created without an O365 email address." -Type Warning -LogPath $LogPath
		}
		else {
			# Test to see if the account was synced first
			while (!(Get-MsolUser -UserPrincipalName $UserPrincipal -ErrorAction SilentlyContinue)) { Write-Host "User not yet synced, waiting 30 seconds before retrying..." -ForegroundColor Cyan; sleep 30 }
				try {
					#Can ignore the commented out, leaving for refernece, but didn't work/wasn't needed
					#If you want to license users as they are created them from the shell, or after they’ve been synchronized to the cloud with DirSync, you’ll need to know the AccountSkuId, which is in the format of tenant:SkuPartNumber
					#New-MsolUser -UserPrincipalName $UserPrincipal -DisplayName $FullName -FirstName $FirstName -LastName $Lastname -LicenseAssignment ($LicenseCheck | where {$_.skupartnumber -eq "SKUPARTNUMBER"}) -UsageLocation $Country
					#New-MoveRequest -Identity $UserPrincipal -Remote -RemoteHostName $RemoteHost -TargetDeliveryDomain $DeliveryDomain -RemoteCredential $o365creds -BadItemLimit 1000
					#Enable-RemoteMailbox $sAMName -RemoteRoutingAddress $UserRemoteAddress
					Set-MsolUser -UserPrincipalName $UserPrincipal -UsageLocation $Country -erroraction stop
					Set-MsolUserLicense -UserPrincipalName $UserPrincipal -AddLicenses $acctSku -LicenseOptions $LicenseOptions -erroraction stop #Tried scripting this to choose the correct sku but is horrible for no reason; use Get-MsolAccountSku to see the skus and (Get-MsolAccountSku | Where-Object {$_.SkuPartNumber -eq "ENTERPRISEPACK"}).ServiceStatus to see the packs
					
					#Change the primary SMTP address and disable the email address policy(needs to be done to change them anyway) if the user is in [Redacted] as they are special snowflakes; can't be done during new-remotemailbox
					if ($Office -eq "[Redacted]") {
						Set-RemoteMailbox -Identity $sAMName -EmailAddressPolicyEnabled $false
						Set-RemoteMailbox -Identity $sAMName -EmailAddresses SMTP:$sAMName@OTHERDOMAIN.com,smtp:$sAMName@DOMAIN.com
					}
					
					#Set the password to never expire for laptop users
					if (($Country -eq "XX") -or ($ProfPath -eq "")) { Set-ADUser -Identity $sAMName -PasswordNeverExpires $true }
					
					#>
					Logging -ToLog "The O365 account for $UserPrincipal has been created"  -Type info -LogPath $LogPath
				}
				catch {  ErrHandle -ErrMsg "User O365 account creation for $($FullName) was unsuccessful" -FullName $FullName -MailServer $MailServer -ExceptionMsg $_.Exception.Message -InvocationInfo $_.InvocationInfo.positionmessage }
			}
	}
	catch { ErrHandle -ErrMsg "User account creation for $($FullName) was unsuccessful" -FullName $FullName -MailServer $MailServer -ExceptionMsg $_.Exception.Message -InvocationInfo $_.InvocationInfo.positionmessage }
}

# Function to search excel for a long distance code
Function LongDist-Excel([string]$FilePath, [string]$SheetName = "", [int]$ColValue, [int]$Row, [string]$sAMName, [string]$FullName, [string]$OU, [string]$ExcelPC, [string]$ExcelTemp, [string]$LDURL, [string]$MailServer=$MailServer, [string]$EmailTo=$EmailTo, [string]$LogPath=$LogPath, $LockName=$LockName) {
		Invoke-Command -ComputerName $ExcelPC -Authentication Credssp -Credential $SessionCreds -ArgumentList $FilePath, $SheetName, $ColValue, $Row, $sAMName, $FullName, $OU, $MailServer, $EmailTo, $LogPath, $ExcelTemp, $LDURL, $LockName -ScriptBlock {
		param($FilePath, $SheetName, $ColValue, $Row, $sAMName, $FullName, $OU, $MailServer, $EmailTo, $LogPath, $ExcelTemp, $LDURL, $LockName)
	
		#region Functions				
		# Function for error handling
		Function ErrHandle([string]$MailServer=$MailServer, $EmailTo=$EmailTo, [string]$ErrMsg, [string]$FullName, [string]$ExceptionMsg, [string]$InvocationInfo, [string]$ErrAct="exit",$LockName=$LockName) {
		Write-Host ("$ErrMsg. `n $ExceptionMsg `n $InvocationInfo", "Error") -ForegroundColor Magenta
		$Body = "$($ErrMsg). Please review the log file at `"$LogPath`"`n`n", $ExceptionMsg, $InvocationInfo
		Send-MailMessage -To $EmailTo -Subject "User Creation Error for Account: $FullName" -Body "$Body" -From "NOTIFICATIONNAME@DOMAIN.COM" -SmtpServer $MailServer
		Logging -Type error -LogPath $LogPath
		if ($ErrAct -ne "continue") { rm $LockName; exit }
		}
		# Function for logging
		Function Logging([string]$ToLog, [string]$Type, [string]$LogPath=$LogPath, $EmailTo=$EmailTo, $MailServer=$MailServer) {
		try {
			switch ($Type) {
				"Error" {(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - ERROR - " + $Error[0].Exception.Message + " Line: " + $Error[0].InvocationInfo.ScriptLineNumber + " Char: " + $Error[0].InvocationInfo.OffsetInLine | Out-File -FilePath $LogPath -append }
				"Warning" {(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - WARNING - $ToLog" | Out-File -FilePath $LogPath -append}
				Default {(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - INFO - $ToLog" | Out-File -FilePath $LogPath -append}
			}
		}
		catch { 
			# Not sending to the Error Hanlding function since it write to log, which would write to errhandle, which would then try to log....infinitely
			# Not certain $EmailTo and $MailServer will be passed thru; may need to just create a pop up or something
			$Body = "Unable to write to log at $LogPath`n`n", $Error[0].Exception.Message, $Error[0].InvocationInfo.positionmessage
			Send-MailMessage -To $EmailTo -Subject "User Creation Error for Account: $FullName" -Body "$Body" -From "NOTIFICATIONNAME@DOMAIN.COM" -SmtpServer $MailServer
		}
	}
		#endregion
	
		#region Download the Long Distance spreadsheet
		try {
			# Adding a line to kill running Excel processes as sometimes it hangs causing errors
			Write-Host `n"Ignore any red text about 'excel', it's just checking to make sure it's not already running" -ForegroundColor Cyan
			if (Get-Process -processname excel) {Stop-Process -processname excel -Force | Out-Null}
			# Create the folder if it doesn't exist
			if (!(Test-Path $ExcelTemp)){ New-Item -ItemType directory -Path $ExcelTemp | Out-Null }
			# Download the files
			$wc = new-object system.net.webclient
			$wc.UseDefaultCredentials = $true
			$FullPath = $ExcelTemp+($LDURL).split("/")[-1]	
			# Remove old copies if pres
			if (Test-Path $FullPath) {Remove-Item -Path ($ExcelTemp+($LDURL).split("/")[-1])}
			$wc.DownloadFile($LDURL, $FullPath)
		}
		catch {  ErrHandle -ErrMsg "Error downloading the distance sheet" -FullName $FullName -MailServer $MailServer -ExceptionMsg $_.Exception.Message -InvocationInfo $_.InvocationInfo.positionmessage }
		#endregion
	
		# Trims the OU names for the sake of the worksheet
		switch -wildcard ($OU) {
			"CITY*"			 {$OU = ($OU -replace "CITY").trim()}
			"DEPT1*"			 {$OU = ($OU -replace "and DEPTNAME2").trim()}
			"* DEPT3" 	 {$OU = "DEPT3"}
			"DEPT4"				 {$OU = "DEPT4"}
			"*DEPT5"				 {$OU = "DEPT5"}
			"DEPT6*"			 {$OU = "DEPT6"}
		}
			
		try {
			$RowEntry = @()
			$objExcel = New-Object -ComObject Excel.Application
			# Disable the 'visible' property so the document won't open in excel
			$objExcel.Visible = $false
			# Crush dialog boxes; defaults to save when closing
			$objExcel.DisplayAlerts = $false
			# Open the Excel file and save it in $WorkBook
			$WorkBook = $objExcel.Workbooks.Open($FilePath)
			# Load the WorkSheet
			$WorkSheet = $WorkBook.sheets.item($SheetName)

			#Perform the search
			Start-Sleep -Seconds 1
			$RowMax = ($WorkSheet.usedrange.rows).count
			$ColMax = ($WorkSheet.usedrange.columns).count
			for ($row=$Row; $row -le $RowMax; $row++) {
				[string]$CellValue = $WorkSheet.cells.item($row,$ColValue).formula
				if ($CellValue -match $OU) {
					[int]$RowCount = ([regex]::split(($WorkSheet.cells.item($row,($ColValue-1)).formula), "\\|\/")).count # Count the entires already asigned to this one
					$RowEntry += [pscustomobject]@{ Row = $row; Count = $RowCount }
				}
				elseif (([string]::IsNullOrWhiteSpace($CellValue)) -and (!([string]::IsNullOrWhiteSpace($WorkSheet.cells.item($row,($ColValue-4)))))){
					# Empty row with a LDnumber we could use
					if ([string]::IsNullOrWhiteSpace((([regex]::split(($WorkSheet.cells.item($row,($ColValue-1)).formula), "\\|\/"))))) {[int]$RowCount = 0} else {[int]$RowCount = ([regex]::split(($WorkSheet.cells.item($row,($ColValue-1)).formula), "\\|\/")).count} # Count the entires already asigned to this one; ignoring whitespace
					$RowEntry += [pscustomobject]@{ Row = $row; Count = $RowCount }
				}
			} # Row end
			
			# Error checking
			if ($RowEntry -eq $null) {
				Write-Host "No availble codes found." -ForegroundColor Magenta
				Logging -ToLog "No availble codes found" -LogPath $LogPath -Type War
				$WorkBook.close()
				$objExcel.Quit()
				[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null
				if (Get-Process -processname excel) {Stop-Process -processname excel -Force}
				return
			}
			
			# Get the entry with the least amount of names assocaiated with it (also WTH -ascending not working)
			$Use = $RowEntry | Sort-Object count -Descending  | Select-Object -Last 1
			# Update the spreadsheet
			if ([string]::IsNullOrWhiteSpace($WorkSheet.cells.item($Use.row,($ColValue-1)).formula)) {
				$WorkSheet.cells.item($Use.row,($ColValue-1)).value2 = "$($sAMName) $($FullName)"
				$LDCode = $($WorkSheet.cells.item($Use.row,1).value2)
			}
			else {
				$WorkSheet.cells.item($Use.row,($ColValue-1)).value2 = ($WorkSheet.cells.item($Use.row,($ColValue-1)).formula) + "\ $($sAMName) $($FullName)"
				$LDCode = $($WorkSheet.cells.item($Use.row,1).value2)
			}
			if ([string]::IsNullOrWhiteSpace($WorkSheet.cells.item($Use.row, $ColValue).formula)) {
				$WorkSheet.cells.item($Use.row, $ColValue).value2 = $OU
			}

			# Clean up
			$Workbook.save()
			$WorkBook.close()
			$objExcel.Quit()
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null
			if (Get-Process -processname excel) {Stop-Process -processname excel -Force}

			# Update the log and upload the revised spreadsheet
			$wc = new-object system.net.webclient
			$wc.UseDefaultCredentials = $true
			$FullPath = Get-ChildItem ($ExcelTemp+($LDURL).split("/")[-1])
			$wc.UploadFile($LDURL, "PUT", $FullPath.fullname) | Out-Null
			
			# Give back the number
			Logging -ToLog "The long distance code for $($Fullname) has been set to $($LDCode)" -LogPath $LogPath -Type info
			return $LDCode
		}
		catch { ErrHandle -ErrMsg "Long distance code could not be set. Possible causes are the department name or worksheet was not found in the workbook." -FullName $FullName -MailServer $MailServer -ExceptionMsg $_.Exception.Message -InvocationInfo $_.InvocationInfo.positionmessage }
	}
}

# Function to alert to the successful completion of the script
Function Success-Message {
	[Console]::Beep(987 ,53 )
	sleep -milliseconds 53 
	[Console]::Beep(987 ,53 )
	sleep -milliseconds 53 
	[Console]::Beep(987 ,53 )
	sleep -milliseconds 53 
	[Console]::Beep(987 ,428 )
	[Console]::Beep(784 ,428 )
	[Console]::Beep(880 ,428 )
	[Console]::Beep(987 ,107 )
	sleep -milliseconds 214 
	[Console]::Beep(880 ,107 )
	[Console]::Beep(987 ,857 )
	[Console]::Beep(740 ,428 )
	[Console]::Beep(659 ,428 )
	[Console]::Beep(740 ,428 )
	[Console]::Beep(659 ,107 )
	sleep -milliseconds 107 
	[Console]::Beep(880 ,428 )
	[Console]::Beep(880 ,107 )
	sleep -milliseconds 107 
	[Console]::Beep(830 ,428 )
	[Console]::Beep(880 ,107 )
	sleep -milliseconds 107 
	[Console]::Beep(830 ,428 )
	[Console]::Beep(830 ,107 )
	sleep -milliseconds 107 
	[Console]::Beep(740 ,428 )
	[Console]::Beep(659 ,428 )
	[Console]::Beep(622 ,428 )
	[Console]::Beep(659 ,107 )
	sleep -milliseconds 107 
	[Console]::Beep(554 ,1714 )
	[Console]::Beep(740 ,428 )
	[Console]::Beep(659 ,428 )
	[Console]::Beep(740 ,428 )
	[Console]::Beep(659 ,107 )
	sleep -milliseconds 107 
	[Console]::Beep(880 ,428 )
	[Console]::Beep(880 ,107 )
	sleep -milliseconds 107 
	[Console]::Beep(830 ,428 )
	[Console]::Beep(880 ,107 )
	sleep -milliseconds 107 
	[Console]::Beep(830 ,428 )
	[Console]::Beep(830 ,107 )
	sleep -milliseconds 107 
	[Console]::Beep(740 ,428 )
	[Console]::Beep(659 ,428 )
	[Console]::Beep(740 ,428 )
	[Console]::Beep(880 ,107 )
	sleep -milliseconds 107 
	[Console]::Beep(987 ,1714)
}
#endregion

#region Start logging and lock the script
ScriptLock
"" | Out-File -FilePath $LogPath -append
(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - INFO - Start Time" | Out-File -FilePath $LogPath -append
$StartFlag = (Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - INFO - Start Time"
#endregion

#region AD user verification
# Get the users First and Last Name and check to make sure it's not a duplicate
do {
	$FirstName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the new users First Name.","Users Name")
	if ([string]::IsNullOrWhiteSpace($FirstName)) { logging -tolog "The User creation script has been cancelled by user" -type warning; (Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - INFO - End Time" | Out-File -FilePath $LogPath -append; rm $LockName; exit }
	$Lastname = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the new users Last Name.","Users Name")
	if ([string]::IsNullOrWhiteSpace($LastName)) { logging -tolog "The User creation script has been cancelled by user" -type warning; (Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - INFO - End Time" | Out-File -FilePath $LogPath -append; rm $LockName; exit }

	if (Get-ADUser -Filter "surname -eq '$LastName' -and givenname -eq '$FirstName'"){
		$NameExists = [Microsoft.VisualBasic.Interaction]::MsgBox("A user with named $FirstName $LastName already exists, continue anyway?","YesNo","Warning")
		if ($NameExists -eq "Yes") { $FLAvailable = $true }
		else { $FLAvailable = $false }
	}
	else {
	    $FLAvailable = $true
	}
}
until ($FLAvailable)

# Set other information based on their name
$FullName = "$FirstName $LastName"
$sAMName = $FirstName + $Lastname.Substring(0,1)

# Set UserPrincipalName and check to see if it exists; if it does try a variation; todo: make the domain addresses a variable instead
$UserPrincipal = $sAMName+"@DOMAIN.com"
$UserRemoteAddress = $sAMName+"@DOMAIN.MAIL.ONMICROSOFT.COM"
$i = 0
$j = 0
$t = 0

do {
	do {
		# Check to see if the name is already in use
		if (Get-ADUser -Filter "sAMAccountName -eq '$sAMName'"){
			$UPNAvailable = $false
			$j++
			$UserPrincipal = $FirstName + $LastName.substring(0,$j)+ "@DOMAIN.COM"
			$sAMName = $FirstName + $Lastname.substring(0,$j)
			$UserRemoteAddress = $sAMName+"@DOMAIN.MAIL.ONMICROSOFT.COM"
		}	
		else{
		    $UPNAvailable = $true
		    }
	}
	until ($UPNAvailable)
	do {
		# Check to see if the name is too long
		if ($UserPrincipal.Length -gt 25){
			$UPNAvailable2 = $false
			$i++
			$UserPrincipal = $FirstName + $LastName.substring($i,1) + "@DOMAIN.COM"
			$sAMName = $FirstName + $LastName.substring($i,1)
			$UserRemoteAddress = $sAMName+"@DOMAIN.MAIL.ONMICROSOFT.COM"
			}
		else{
		$UPNAvailable2 = $true
			}
		# Check to see if the first name alone is too long; if it is drop a letter
		if (($FirstName -ge 10) -and ($UserPrincipal.Length -gt 25)) {
			$t++
			$UPNAvailable2 = $false
			$UserPrincipal = $FirstName.Substring(0, ($FirstName.Length-$t))+ $LastName.Substring(($i-1),1) + "@DOMAIN.COM"
			$sAMName = $FirstName.Substring(0, ($FirstName.Length-$t))+ $LastName.Substring(($i-1),1)
			$UserRemoteAddress = $sAMName+"@DOMAIN.MAIL.ONMICROSOFT.COM"
			}
		else{
			$UPNAvailable2 = $true
			}
	}
	until ($UPNAvailable2)
	$UPNAvailable3 = $true
}
until ($UPNAvailable3)


do {
	$Location = [Microsoft.VisualBasic.Interaction]::InputBox("Select OU Location (For XX employees enter XX)", "Location", "DEFAULT LOCATION")
	if ([string]::IsNullOrWhiteSpace($Location)) { logging -tolog "The User creation script has been cancelled by user" -type warning; (Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - INFO - End Time" | Out-File -FilePath $LogPath -append; rm $LockName; exit }
	$OU = [Microsoft.VisualBasic.Interaction]::InputBox("Enter an OU Department (eg: Finance, for XX employees enter the City)", "DEFAULT DEPARTMENT")
	if ([string]::IsNullOrWhiteSpace($OU)) { logging -tolog "The User creation script has been cancelled by user" -type warning; (Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - INFO - End Time" | Out-File -FilePath $LogPath -append; rm $LockName; exit }
	#region Some quality of life checks to update the OUs to match the location if something similar but not exactly what is in AD is put in
	if (($Location -match "XY" -and $OU -notmatch "XY") -or ($Location -match "CITY1" -and $OU -notmatch "CITY1")) {
		$OU = $Location + " " + $OU	
	}
	elseif ($Location -match "CITY2" -and $OU -notmatch "CITY2-1") {
		$OU = "CISTY2-1 " + $OU
	}
	if ($OU -like "DEPART1") {
		if ($Location -notmatch "CITY3") {
			$OU = $Location + " DEAPRT and DEPART"
		}
		else {
			$OU = "DEPART and DEPART"
		}
	}
	#endregion
	$OUPath = "AD:\OU=$($OU),OU=Users,OU=$($Location),DC=DOMAIN,DC=com"
	$FullOU = "DOMAIN.com/$($Location)/users/$($OU)"
	if (Test-Path $OUPath) {
		$oufound = $true
	}
	else {
		$pause = [Microsoft.VisualBasic.Interaction]::MsgBox("No such OU found. Pres OK to try again", 0, "Warning")
		$OUFound = $false
	}
}
until ($OUFound)

do {
	$Password = ConvertTo-SecureString ([Microsoft.VisualBasic.Interaction]::InputBox("Enter a temporary password", "Password", "DEFAULT PASSWORD")) -AsPlainText -Force
	if(($Password -eq "") -or ($Password -eq $null)){
		$pause = [Microsoft.VisualBasic.Interaction]::MsgBox("Password cannot be blank. Pres OK to try again", 0, "Warning")
		$PassOK = $false
	}
	else{
	$PassOK = $true
	}
}
until ($PassOK)

do {
# Get the like user and make sure they exist
	$LikeUser = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the logon name of the like user", "Like User")
	if ([string]::IsNullOrWhiteSpace($LikeUser)) { logging -tolog "The User creation script has been cancelled by user" -type warning; (Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - INFO - End Time" | Out-File -FilePath $LogPath -append; rm $LockName; exit }
	if (Get-ADUser -Filter "SAMaccountname -eq '$LikeUser'" ){
		$LUFullName = (Get-ADUser -Identity $LikeUser -Properties GivenName).GivenName + " " + (Get-ADUser -Identity $LikeUser -Properties SurName).SurName
		$RLUser = [Microsoft.VisualBasic.Interaction]::MsgBox("Use $LUFullName as the like user?", 4, "Warning")
		if(($RLUser -eq "") -or ($RLUser -eq $null) -or ($RLUser -eq "Yes")){
			$LUCheck = $true
		}
		else{
			$LUCheck = $false
		}
	}
	else{
		$pause = [Microsoft.VisualBasic.Interaction]::MsgBox("The Like User cannot be found. Press enter to try again.", 0, "Warning")
		$LUCheck = $false
	}
}
until ($LUCheck)

# Laptop and some remote desktop only users should not have one
$RoamingProfile = [Microsoft.VisualBasic.Interaction]::MsgBox("Create a roaming profile for the user?", 4, "Roaming Profile Setup")

#endregion

#region User creation

#region Exchange and AD


# Faux-Create the Mailbox/AD Account
$NoEcho = [System.Windows.Forms.MessageBox]::Show("Creating Account:`nName: $FullName  `nAlias: $sAMName  `nOrganizationalUnit: $OU  `nUserPrincipalName: $UserPrincipal `nThis will take a minute. Press OK to continue.","Information")
if ($NoEcho -eq "Cancel") { logging -tolog "The User creation script has been cancelled by user" -type warning; (Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - INFO - End Time" | Out-File -FilePath $LogPath -append; rm $LockName; exit }

# Setup is slightly different based on user location
switch ($Location) {
	"CITY1" {
				if ($RoamingProfile -eq "No") { $ProfPath = "none" } else { $ProfPath = "\\SERVER\FOLDER\$sAMName" }
				$Office = "MapleGrove"; $HomeDir = "\\SERVER\FOLDER\$sAMName"; $HomeDrive = "F:"
				$ScriptPath = "LOGINSCRIPT"; $Street = "ADDRESS"; $City = "CITY1"; $State = "STATE"; $Postal = "POSTAL"; $Country = "ZZ"; $Company = "COMPANY"
				CreateADEmail0365 -FullName $FullName -FirstName $FirstName -LastName $LastName -sAMName $sAMName -UserPrincipal $UserPrincipal -Password $Password -Database $Database -ProfPath $ProfPath -ScriptPath $ScriptPath -HomeDrive $HomeDrive -HomeDir $HomeDir -Office $Office -Street $Street -City $City -State $State -Postal $Postal -Country $Country -OU $OU -FullOu $FullOU -Company $Company -O365Creds $o365creds
	}
	"CITY2" {	
				if ($RoamingProfile -eq "No") { $ProfPath = "none" } else { $ProfPath = "\\SERVER\FOLDER\$sAMName" }
				$Office = "CITY2"; $HomeDir = "\\SERVER\FOLDER\$sAMName"; $HomeDrive = "F:"
				$ScriptPath = "LOGINSCRIPT"; $Street = "ADDRESS"; $City = "CITY2"; $State = "STATE"; $Postal = "POSTAL"; $Country = "ZZ"; $Company = "COMPANY"
				CreateADEmail0365 -FullName $FullName -FirstName $FirstName -LastName $LastName -sAMName $sAMName -UserPrincipal $UserPrincipal -Password $Password -Database $Database -ProfPath $ProfPath -ScriptPath $ScriptPath -HomeDrive $HomeDrive -HomeDir $HomeDir -Office $Office -Street $Street -City $City -State $State -Postal $Postal -Country $Country -OU $OU -FullOu $FullOU -Company $Company -O365Creds $o365creds
	}
	"CITY3" {
				if ($RoamingProfile -eq "No") { $ProfPath = "none" } else { $ProfPath = "\\SERVER\FOLDER\$sAMName" }
				$Office = "CITY3"; $HomeDir = "\\SERVER\FOLDER\$sAMName"; $HomeDrive = "F:"
				$ScriptPath = "LOGINSCRIPT"; $Street = "ADDRESS"; $City = "CITY3"; $State = "STATE"; $Postal = "POSTAL"; $Country = "ZZ"; $Company = "COMPANY"
				CreateADEmail0365 -FullName $FullName -FirstName $FirstName -LastName $LastName -sAMName $sAMName -UserPrincipal $UserPrincipal -Password $Password -Database $Database -ProfPath $ProfPath -ScriptPath $ScriptPath -HomeDrive $HomeDrive -HomeDir $HomeDir -Office $Office -Street $Street -City $City -State $State -Postal $Postal -Country $Country -OU $OU -FullOu $FullOU -Company $Company -O365Creds $o365creds
	}
	"CITY4" {
				if ($RoamingProfile -eq "No") { $ProfPath = "none" } else { $ProfPath = "\\SERVER\FOLDER\$sAMName" }
				$Office = "CITY4"; $HomeDir = "\\SERVER\FOLDER\$sAMName"; $HomeDrive = "F:"
				$ScriptPath = "LOGINSCRIPT"; $Street = "ADDRESS"; $City = "CITY4"; $State = "STATE"; $Postal = "POSTAL"; $Country = "ZZ"; $Company = "COMPANY"
				CreateADEmail0365 -FullName $FullName -FirstName $FirstName -LastName $LastName -sAMName $sAMName -UserPrincipal $UserPrincipal -Password $Password -Database $Database -ProfPath $ProfPath -ScriptPath $ScriptPath -HomeDrive $HomeDrive -HomeDir $HomeDir -Office $Office -Street $Street -City $City -State $State -Postal $Postal -Country $Country -OU $OU -FullOu $FullOU -Company $Company -O365Creds $o365creds
					
				# ADSI needs to be used to set Remote Desktop aka Terminal Services profile settings
				$User = [ADSI] "LDAP://CN=$($FullName),OU=$($OU),OU=Users,OU=$($Location),DC=DOMAIN,DC=com"
				$User.psbase.Invokeset("terminalservicesprofilepath",$ProfPath)
				$User.psbase.invokeSet("TerminalServicesHomeDirectory",$HomeDir)
				$User.psbase.invokeSet("TerminalServicesHomeDrive","F:")
				$User.setinfo()
	}
	"CITY5" {
				if ($RoamingProfile -eq "No") { $ProfPath = "none" } else { $ProfPath = "\\SERVER\FOLDER\$sAMName" }
				# Even thought some settings are different they are still considered MapleGrove users for some settings
				$Office = "CITY5"; $Location = "CITY5"; $HomeDir = "\\cambmain\users1\$sAMName"; $HomeDrive = "F:"
				$ScriptPath = "LOGINSCRIPT"; $Street = "ADDRESS"; $City = "CITY5"; $State = "STATE"; $Postal = "POSTAL"; $Country = "ZZ"; $Company = "COMPANY"
				CreateADEmail0365 -FullName $FullName -FirstName $FirstName -LastName $LastName -sAMName $sAMName -UserPrincipal $UserPrincipal -Password $Password -Database $Database -ProfPath $ProfPath -ScriptPath $ScriptPath -HomeDrive $HomeDrive -HomeDir $HomeDir -Office $Office -Street $Street -City $City -State $State -Postal $Postal -Country $Country -OU $OU -FullOu $FullOU -Company $Company -O365Creds $o365creds
					
				<# They shouldn't need this anymore since the fiber was installed, leaving in case
				# ADSI needs to be used to set Remote Desktop aka Terminal Services profile settings
				$User = [ADSI] "LDAP://CN=$($FullName),OU=$($OU),OU=Users,OU=$($Location),DC=DOMAIN,DC=com"
				$User.psbase.Invokeset("terminalservicesprofilepath",$ProfPath)
				$User.psbase.invokeSet("TerminalServicesHomeDirectory",$HomeDir)
				$User.psbase.invokeSet("TerminalServicesHomeDrive","F:")
				$User.setinfo() #>	
	}
	"COUNTRY1" {			
		switch ($OU) {
			"CITY6" {
						if ($RoamingProfile -eq "No") { $ProfPath = "none" } else { $ProfPath = "\\SERVER\FOLDER\$sAMName" }
						$Office = "CITY6"; $HomeDir = "\\SERVER\FOLDER\$($sAMName)"; $HomeDrive = "F:"
						$ScriptPath = "LOGINSCRIPT"; $Street = "ADDRESS"; $City = "CITY6"; $State = "STATE"; $Postal = "POSTAL"; $Country = "XX"; $Company = "COMPANY"
						CreateADEmail0365 -FullName $FullName -FirstName $FirstName -LastName $LastName -sAMName $sAMName -UserPrincipal $UserPrincipal -Password $Password -Database $Database -ProfPath $ProfPath -ScriptPath $ScriptPath -HomeDrive $HomeDrive -HomeDir $HomeDir -Office $Office -Street $Street -City $City -State $State -Postal $Postal -Country $Country -OU $OU -FullOu $FullOU -Company $Company -O365Creds $o365creds
			}
			"CITY7" {
						if ($RoamingProfile -eq "No") { $ProfPath = "none" } else { $ProfPath = "\\SERVER\FOLDER\$sAMName" }
						$Office = "CITY7"; $HomeDir = "\\SERVER\FOLDER\$($sAMName)"; $HomeDrive = "F:"
						$ScriptPath = "LOGINSCRIPT"; $Street = "ADDRESS"; $City = "CITY7"; $State = "STATE"; $Postal = "POSTAL"; $Country = "XX"; $Company = "COMPANY"
						CreateADEmail0365 -FullName $FullName -FirstName $FirstName -LastName $LastName -sAMName $sAMName -UserPrincipal $UserPrincipal -Password $Password -Database $Database -ProfPath $ProfPath -ScriptPath $ScriptPath -HomeDrive $HomeDrive -HomeDir $HomeDir -Office $Office -Street $Street -City $City -State $State -Postal $Postal -Country $Country -OU $OU -FullOu $FullOU -Company $Company -O365Creds $o365creds
			}
		}			
	}
}
$NoEcho = [System.Windows.Forms.MessageBox]::Show("Account created successfully for:`nName: $FullName `nAlias: $sAMName `nOrganizationalUnit: $OU `nUserPrincipalName: $UserPrincipal `nSAMName: $sAMName","Information")
if ($NoEcho -eq "Cancel") { logging -tolog "The User creation script has been cancelled by user" -type warning; (Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - INFO - End Time" | Out-File -FilePath $LogPath -append; rm $LockName; exit }
#endregion

#region User profile
# Create the users profile folder and copy over any specific programs they may need to their desktop

#Enumerate and set the like user's memberof groups to the new user
$LUGroups = Get-ADUser $LikeUser -property memberof | select memberof
foreach($group in $LUGroups.memberof){
	Add-ADGroupMember -Identity $group -Members $sAMName
}

# XX user accounts are all local
if ($Location -ne "XX") {
	try {
		# Create the folder
		New-Item -ItemType directory -Path ($ProfPath+".v2") -ErrorAction Stop | Out-Null

		if (Test-Path $ProfPath".v2") {
			# Permissions Section
			$ACL = Get-Acl ($ProfPath+".v2")
			$Permission = New-Object System.Security.AccessControl.FileSystemAccessRule($UserPrincipal,"fullcontrol", "containerinherit, objectinherit", "none", "allow")
			$AdminPerm = New-Object System.Security.AccessControl.FileSystemAccessRule("administrators", "fullcontrol", "containerinherit, objectinherit", "none", "allow")
			$SystemPerm = New-Object System.Security.AccessControl.FileSystemAccessRule("system","fullcontrol", "containerinherit, objectinherit", "none", "allow")

			# Remove all permissions
			$ACL.Access | %{$ACL.RemoveAccessRule($_)}  | Out-Null
				
			# Set permissions
			$ACL.SetAccessRule($Permission)
			$ACL.SetAccessRule($AdminPerm)
			$ACL.SetAccessRule($SystemPerm)
			Set-ACL ($ProfPath+".v2") $ACL
					
			# Inheritance removal
			$ACL.SetAccessRuleProtection($true,$false)
			$ACL | Set-ACL

			# Create the desktop folder and copy over desktop items if there are any
			New-Item -ItemType directory -Path ($ProfPath+".v2\desktop") -ErrorAction Stop | Out-Null
			if (Test-Path "\\SERVER\FOLDER\FOLDER\$($OU)"){
				Copy-Item -Path "\\SERVER\FOLDER\FOLDER\$($OU)\*.*" -Destination ($ProfPath+".v2\desktop") | Out-Null
			}
		}
	}
	catch { ErrHandle -ErrMsg "Error with profile path $($ProfPath).v2" -FullName $FullName -MailServer $MailServer -ExceptionMsg $_.Exception.Message -InvocationInfo $_.InvocationInfo.positionmessage -ErrAct Continue }

	try {
		# Create the Home Directory folder; Doing this becasue although AD shows it properly it doesn't seem to auto create
		New-Item -ItemType directory -Path ($HomeDir) | Out-Null

		if (Test-Path $HomeDir) {
			# Permissions Section
			$ACL = Get-Acl ($HomeDir)
			$Permission = New-Object System.Security.AccessControl.FileSystemAccessRule($UserPrincipal,"fullcontrol", "containerinherit, objectinherit", "none", "allow")
			$AdminPerm = New-Object System.Security.AccessControl.FileSystemAccessRule("administrators", "fullcontrol", "containerinherit, objectinherit", "none", "allow")
			$SystemPerm = New-Object System.Security.AccessControl.FileSystemAccessRule("system","fullcontrol", "containerinherit, objectinherit", "none", "allow")

			# Remove all permissions
			$ACL.Access | %{$ACL.RemoveAccessRule($_)}  | Out-Null
				
			# Set permissions
			$ACL.SetAccessRule($Permission)
			$ACL.SetAccessRule($AdminPerm)
			$ACL.SetAccessRule($SystemPerm)
			Set-ACL ($HomeDir) $ACL
					
			# Inheritance removal
			$ACL.SetAccessRuleProtection($true,$false)
			$ACL | Set-ACL

			Copy-Item -Path "\\SERVER\FOLDER\FOLDER\FOLDER\FILE" -Destination $HomeDir | Out-Null
		}
	}
	catch { ErrHandle -ErrMsg "Error with directory $($HomeDir)" -FullName $FullName -MailServer $MailServer -ExceptionMsg $_.Exception.Message -InvocationInfo $_.InvocationInfo.positionmessage -ErrAct Continue }
}
#endregion

#endregion #User creation end

#region Lync
# Connect the the Lync server and enable the account
$NoEcho = [Microsoft.VisualBasic.Interaction]::MsgBox("Connecting to Lync server and configuring profile. `nPress ok to continue.", 1, "Lync Setup")
if ($NoEcho -eq "Cancel") { logging -tolog "The Lync setup for $($FullName) has been cancelled by user" -type warning }
else {
	Invoke-Command -ComputerName $LyncServer -Authentication Credssp -Credential $LyncCreds -ArgumentList $FullName, $Location, $MailServer, $EmailTo, $LogPath, $LockName -ScriptBlock {
		param($FullName, $Location, $MailServer, $EmailTo, $LogPath, $LockName)
		#region Functions
		Import-Module 'C:\Program Files\Common Files\Microsoft Lync Server 2010\Modules\Lync\Lync.psd1'
		# It was just easier to create the function than to pass a funtion into an invoke that invokes to use it.. could have called a file (like a .dll) but don't want to
		Function ErrHandle([string]$MailServer=$MailServer, $EmailTo=$EmailTo, [string]$ErrMsg, [string]$FullName, [string]$ExceptionMsg, [string]$InvocationInfo, [string]$ErrAct="exit", $LockName=$LockName) {
			Write-Host ("$ErrMsg. `n $ExceptionMsg `n $InvocationInfo", "Error") -ForegroundColor Magenta
			$Body = "$($ErrMsg). Please review the log file at `"$LogPath`"`n`n", $ExceptionMsg, $InvocationInfo
			Send-MailMessage -To $EmailTo -Subject "User Creation Error for Account: $FullName" -Body "$Body" -From "NOTIFICATIONNAME@DOMAIN.COM" -SmtpServer $MailServer
			Logging -Type error -LogPath $LogPath
			if ($ErrAct -ne "continue") { rm $LockName; exit }
		}
		# Function for logging
		Function Logging([string]$ToLog, [string]$Type, [string]$LogPath=$LogPath, $EmailTo=$EmailTo, $MailServer=$MailServer) {
			try {
				switch ($Type) {
					"Error" {(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - ERROR - " + $Error[0].Exception.Message + " Line: " + $Error[0].InvocationInfo.ScriptLineNumber + " Char: " + $Error[0].InvocationInfo.OffsetInLine | Out-File -FilePath $LogPath -append }
					"Warning" {(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - WARNING - $ToLog" | Out-File -FilePath $LogPath -append}
					Default {(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - INFO - $ToLog" | Out-File -FilePath $LogPath -append}
				}
			}
			catch { 
				# Not sending to the Error Hanlding function since it write to log, which would write to errhandle, which would then try to log....infinitely
				# Not certain $EmailTo and $MailServer will be passed thru; may need to just create a pop up or something
				$Body = "Unable to write to log at $LogPath`n`n", $Error[0].Exception.Message, $Error[0].InvocationInfo.positionmessage
				Send-MailMessage -To $EmailTo -Subject "User Creation Error for Account: $FullName" -Body "$Body" -From "NOTIFICATIONNAME@DOMAIN.COM" -SmtpServer $MailServer
			}
		}
		#endregion
		try {
			Enable-CSUser -identity $($FullName) -RegistrarPool SERVERNAME.DOMAIN.COM -sipaddresstype emailaddress -sipdomain DOMAIN.COM -erroraction stop #todo: Should make those into variables
			sleep 30 #give it a moment to enable, otherwise it doesn't find the username when setting policies; try/catch in case it still doesn't
			logging -tolog "The Lync account for $($FullName) has been enabled" -type info
			Write-Host "`nThe Lync account for $($FullName) has been enabled" -ForegroundColor Cyan
		}
		catch {
			errhandle -ErrMsg "The Lync account wast not enabled, please manually enable the account." -FullName $FullName -MailServer $MailServer -ExceptionMsg $_.Exception.Message -InvocationInfo $_.InvocationInfo.positionmessage
			Write-Host ("`n$ErrMsg. `n $ExceptionMsg `n $InvocationInfo", "Error") -ForegroundColor Magenta
		}
		try {
			Grant-CSConferencingPolicy -identity $($FullName) -policyname "Policy 1 (High)" -erroraction stop #todo: make the policy name into a variable
			logging -tolog "The Lync conferencing policy for $($FullName) has been set to Policy 1 (High)" -type info
			Write-Host "`nThe Lync conferencing policy for $($FullName) has been set to Policy 1 (High)" -ForegroundColor cyan
		}
		catch {
			errhandle -ErrMsg "The conferencing Policy was not set, please set manually to Policy 1 (High)" -FullName $FullName -MailServer $MailServer -ExceptionMsg $_.Exception.Message -InvocationInfo $_.InvocationInfo.positionmessage
			Write-Host ("`n$ErrMsg. `n $ExceptionMsg `n $InvocationInfo", "Error") -ForegroundColor Magenta
		}
		if ($Location -eq "US") {
			try {
				Grant-CSExternalAccessPolicy -identity $($FullName) -policyname "Allow Outside Access" -erroraction stop
				logging -tolog "The Lync External Access Policy for $($FullName) has been set to Allow Outside Access" -type info
				Write-Host "`nThe Lync External Access Policy for $($FullName) has been set to Allow Outside Access" -ForegroundColor cyan
			}
			catch {
				errhandle -errmsg "The external access policy was not set, please set manually to Allow Outside Access" -FullName $FullName -MailServer $MailServer -ExceptionMsg $_.Exception.Message -InvocationInfo $_.InvocationInfo.positionmessage
				Write-Host ("`n$ErrMsg. `n $ExceptionMsg `n $InvocationInfo", "Error") -ForegroundColor Magenta
			}
		}
	}
}
#endregion

#region Long Distance Code
$NoEcho = [Microsoft.VisualBasic.Interaction]::MsgBox("Press OK to set a long distance code for the user. `nCancel to skip.", 1, "Long Distance Code Setup")
if ($NoEcho -eq "Cancel") { logging -tolog "The Long Distance code setup for $($FullName) has been cancelled by user" -type warning }
else {
	# Check which spreadsheet to load; both spreadsheets are formatted differently
	if ($Location -match "LOCATIONNAME1") {
		$LDURL = "http://WEBSITE/SITES/DEPT/FOLDERNAME/FOLDER/FILENAME1.xlsx"
		$LDCode = LongDist-Excel -FilePath $ExcelTemp"FILENAME1.xlsx" -SheetName "Sheet1" -ColValue 5 -Row 4 -SAMName $sAMName -FullName $FullName -OU $OU -ExcelPC $ExcelPC -ExcelTemp $ExcelTemp -LDURL $LDURL
	}
	elseif (($Location -match "LOCATIONNAME2") -or ($Location -match "LOCATIONNAME3") -or ($Location -match "LOCATIONNAME4")) {
		$LDURL = "http://WEBSITE/SITES/DEPT/FOLDERNAME/FOLDER/FILENAME2.XLS"
		$LDCode = LongDist-Excel -FilePath $ExcelTemp"FILENAME2.XLS" -SheetName "Long Distance" -ColValue 5 -Row 2 -SAMName $sAMName -FullName $FullName -OU $OU -ExcelPC $ExcelPC -ExcelTemp $ExcelTemp -LDURL $LDURL
	}

	# Here's the code; skip if it wasn't needed to be set
	if (!([string]::IsNullOrWhiteSpace($LDCode))) {
		[System.Windows.Forms.MessageBox]::Show("Long distance code for $($FullName): $($LDCode)","Info") | Out-Null
	}

	# Remove the temporary directory used for Excel
	if (Test-Path ($ExcelTemp)) {
	rmdir -Path $ExcelTemp -Recurse -Force
	}
}
#endregion

#region Time stamp the end and play alert
(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - INFO - End Time" | Out-File -FilePath $LogPath -append
$EndFlag = (Get-Date -Format "yyyy-MM-dd hh:mm:ss tt").tostring() + " - INFO - End Time"
Success-Message
rm $LockName
#endregion