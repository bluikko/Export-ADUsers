# Export-ADUsers
PowerShell script to export users from Active Directory to Excel and format it nicely.

Everyone has their own version and this is my attempt at it. It probably needs modifications for other environments.

## Requirements
Excel functionality is provided by the excellent ImportExcel module. It must be installed first. See https://github.com/dfinke/ImportExcel

## Configuration
Following variables **must** be changed according to the environment:
* `$searchbase`: LDAP search base.
* `$reportfile`: Where to save the file.
* `$maildomain`: Mail domain for checking Mail attribute correctness.

## Output
Following attributes will be exported:
* sAMAccountName
* PwdLastSet
* UserAccountControl decoded to flags
* Displayname
* GivenName
* SurName
* Title
* Mail (Aliasmail attribute, if not null, added as comment to the cell; each Aliasmail element separated with ';')
* Company
* Department
* Office
* EmployeeNumber
* MobilePhone
* primaryGroupID
* EmployeeID
* msNPAllowDialin
* generationQualifier
* EmployeeType
* userWorkStations
* Manager

The above attributes are important for me. Many of them might not be important for someone else so feel free to delete what you don't need. Formatting hardcodes column IDs (does not use named ranges) so those might need to be changed if the columns to be exported are changed.

Formatting will be applied to the exported columns:
* If sAMAccountName is not GivenName.SurName, use red text.
* If Mail is not sAMAccountName@$maildomain, use red text.
* If Displayname is not GivenName SurName, use red text.
* if primaryGroupID is Domain Users and Title does not exist, use red text.
* UserAccountControl flags - formattings do stack if many conditions apply:
** Blue strikethru text if account is disabled.
** Blue italic text if account has expired password.
** Blue underlined text if account is locked.
** Yellow background if flags affecting security: PASSWD_NOTREQD / USE_DES_KEY_ONLY / DONT_REQ_PREAUTH.

## TODO
* Make `$maildomain` empty by default and do Mail attribute checks only if it is set.
* Mail attribute check should probably check GivenName.FamilyName instead of sAMAccountName.
* Privileged/special primaryGroupIDs might need some highlighting.
* Group memberships should be exported somehow. Perhaps to another worksheet as a matrix with sAMAccountName at Y and group name (printed vertically) at X.
