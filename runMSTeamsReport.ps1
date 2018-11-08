$strReportFile = ".\TeamsGroupReport.csv"
$strCredentialUIDFile = "..\win-uid.txt"
$strCredentialPwdFile = "..\win-pass.txt"

$strCredentialUID= get-content -path $strCredentialUIDFile 
$strPassword = get-content -path $strCredentialPwdFile | convertto-securestring
$userCredential = new-object -typename PSCredential -argumentlist $strCredentialUID,$strPassword

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

#$Groups = Get-UnifiedGroup -ResultSize 10 -Filter {name -eq "ITCentralSupport_fbee2f8e95"} |Where-Object {$_.ProvisioningOption -eq 'ExchangeProvisioningFlags:481' -or $_.ProvisioningOption -eq 'ExchangeProvisioningFlags:3552'}
$Groups = Get-UnifiedGroup -ResultSize Unlimited |Where-Object {$_.ProvisioningOption -eq 'ExchangeProvisioningFlags:481' -or $_.ProvisioningOption -eq 'ExchangeProvisioningFlags:3552'}

write-host "DisplayName`tDescription`tMemberCount`tLastPost`tOwner`tDepartment"
set-content -Path $strReportFile -Value "`"DisplayName`",`"Description`",`"MemberCount`",`"WhenCreated`",`"LastPost`",`"Owner`",`"Department`""

ForEach ($G in $Groups) {
	$strNewestItemTimestamp = (Get-MailboxFolderStatistics -Identity $G.Alias -IncludeOldestAndNewestItems -FolderScope ConversationHistory).NewestItemReceivedDate

	write-host -NoNewline ($G).DisplayName "`t" ($G).MemberJoinRestriction "`t" ($G).GroupMemberCount "`t" ($G).WhenCreatedUTC  "`t" $strNewestItemTimestamp
	add-content -NoNewLine -Path $strReportFile -Value "`"$(($G).DisplayName)`",`"$(($G).Notes)`",`"$(($G).GroupMemberCount)`",`"$(($G).WhenCreatedUTC)`",`"$strNewestItemTimestamp `""
	$strOwners = ($G).ManagedBy
	ForEach($strOwner in $strOwners){
		$objOwner = get-aduser -Filter {name -eq $strOwner} -Properties department
		write-host -NoNewLine  "`t" $strOwner "`t" $objOwner.department
		add-content -NoNewLine -Path $strReportFile -Value ",`"$strOwner `",`" $($objOwner.department)`""
	}
	write-host ""
	add-content -Path $strReportFile ""
}

Remove-PSSession $Session

