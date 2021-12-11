<#
    Export-ADUsers.ps1 - Exports users from Active Directory to Excel.
    Copyright (C) 2020  Ville Ojamo

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
#>

# Requires https://github.com/dfinke/ImportExcel

## Settings that must be changed according to environment
# LDAP search base
$searchbase = "ou=People,dc=corp,dc=fabrikam,dc=com"
# Where to save the report
$reportfile = "D:\adusers-$(Get-Date -Format yyyy-MM-dd).xlsx"
# Mail domain for checking Mail attribute correctness
$maildomain = "fabrikam.com"

if (!(Get-Module -ListAvailable | Where-Object {$_.Name -eq "ImportExcel"})) {
 Write-Host "The ImportExcel module needs to be installed. See https://github.com/dfinke/ImportExcel"
 exit 1
}

$uacflags = @("SCRIPT","ACCOUNTDISABLE","","HOMEDIR_REQUIRED","LOCKOUT","PASSWD_NOTREQD","PASSWD_CANT_CHANGE","ENCRYPTED_TEXT_PWD_ALLOWED",
 "TEMP_DUPLICATE_ACCOUNT","NORMAL_ACCOUNT","","INTERDOMAIN_TRUST_ACCOUNT","WORKSTATION_TRUST_ACCOUNT","SERVER_TRUST_ACCOUNT","","","DONT_EXPIRE_PASSWORD",
 "MNS_LOGON_ACCOUNT","SMARTCARD_REQUIRED","TRUSTED_FOR_DELEGATION","NOT_DELEGATED","USE_DES_KEY_ONLY","DONT_REQ_PREAUTH","PASSWORD_EXPIRED",
 "TRUSTED_TO_AUTH_FOR_DELEGATION","","PARTIAL_SECRETS_ACCOUNT")

$userdata = Get-ADUser -SearchBase $searchbase -Filter * -Properties * |
 Select-Object -Property sAMAccountName,@{ Name = "PwdLastSet"; Expression = { [datetime]::FromFileTime($_.PwdLastSet) } },
 @{ Name = 'userAccountControl'; Expression = { $value = $_.userAccountControl; $flags = $uacflags; 1..($flags.length) | ? {$value -band [math]::Pow(2,$_)} | % { $flags[$_] + " " } | Out-String }; },
 Displayname,GivenName,SurName,Title,Mail,@{ Name = 'Aliasmail'; Expression = { $_.url -join ';'; }; },Company,Department,
 Office,EmployeeNumber,MobilePhone,primaryGroupID,EmployeeID,msNPAllowDialin,generationQualifier,EmployeeType,userWorkStations,Manager
$excel = $userdata | Export-Excel -Path $reportfile -AutoSize -FreezeTopRow -AutoFilter -BoldTopRow -ClearSheet -WorksheetName Users -AutoNameRange -NoNumberConversion MobilePhone,PwdLastSet -ExcludeProperty Aliasmail -PassThru

$sheet = $excel.Workbook.Worksheets["Users"]

## conditional formatting for common errors
# if sAMAccountName is not "familyname.givenname"
Add-ConditionalFormatting -Address $sheet.Cells["sAMAccountName"] -RuleType NotEqual -ConditionValue '=IF(LEN(G2)>0,LOWER(CONCATENATE(E2, ".", F2)),A2)' -ForegroundColor Red
# if Mail is not "familyname.givenname@maildomain"
Add-ConditionalFormatting -Address $sheet.Cells["Mail"] -RuleType NotEqual -ConditionValue "=CONCATENATE(A2, $($maildomain))" -ForegroundColor Red
# if Displayname is not "Familyname Givenname"
Add-ConditionalFormatting -Address $sheet.Cells["Displayname"] -RuleType NotEqual -ConditionValue '=IF(LEN(G2)>0,CONCATENATE(E2, " ", F2),D2)' -ForegroundColor Red
# if primaryGroupID is Domain Users and Title doesn't exist
Add-ConditionalFormatting -Address $sheet.Cells["primaryGroupID"] -RuleType Equal -ConditionValue "=IF(LEN(G2)>0,M2,513)" -ForegroundColor Red

## conditional foratting for account flags
# blue strikethru text if account is disabled
Add-ConditionalFormatting -Address $sheet.Cells["userAccountControl"] -RuleType ContainsText -ConditionValue "ACCOUNTDISABLE" -ForegroundColor Blue -StrikeThru
# blue italic text if account has expired password
Add-ConditionalFormatting -Address $sheet.Cells["userAccountControl"] -RuleType ContainsText -ConditionValue "PASSWORD_EXPIRED" -ForegroundColor Blue -Italic
# blue underlined text if account is locked
Add-ConditionalFormatting -Address $sheet.Cells["userAccountControl"] -RuleType ContainsText -ConditionValue "LOCKOUT" -ForegroundColor Blue -Underline
# yellow background if flags affecting security
Add-ConditionalFormatting -Address $sheet.Cells["userAccountControl"] -RuleType ContainsText -ConditionValue "PASSWD_NOTREQD" -BackgroundColor Yellow
Add-ConditionalFormatting -Address $sheet.Cells["userAccountControl"] -RuleType ContainsText -ConditionValue "USE_DES_KEY_ONLY" -BackgroundColor Yellow
Add-ConditionalFormatting -Address $sheet.Cells["userAccountControl"] -RuleType ContainsText -ConditionValue "DONT_REQ_PREAUTH" -BackgroundColor Yellow

for ($i = 0; $i -lt $userdata.Length; $i++) {
 if ($userdata[$i].Aliasmail -ne "") {
  $null = $sheet.Cells["H$($i+2)"].AddComment($userdata[$i].Aliasmail, $userdata[$i].Mail)
 }
}

Close-ExcelPackage -Show $excel
