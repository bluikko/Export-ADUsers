# Export-ADUsers
PowerShell script that exports users from Active Directory to an Excel file and formats it nicely. Discrepancies and other potential issues are highlighted.

Everyone has their own version and this is my attempt at it. It probably needs modifications for other environments.

## Requirements
Excel functionality is provided by the excellent ImportExcel module. It must be installed first. See https://github.com/dfinke/ImportExcel

## Configuration
Following variables **must** be changed according to the environment:
* `$searchbase`: LDAP search base.
* `$reportfile`: Where to save the resulting `.xlsx` report file.
* `$maildomain`: Mail domain for checking _Mail_ attribute correctness.

## Output
Following LDAP attributes will be exported:
* _sAMAccountName_
* _PwdLastSet_
* _UserAccountControl_ (decoded to flags; each flag separated with a space + newline)
* _Displayname_
* _GivenName_
* _SurName_
* _Title_
* _Mail_ (_Aliasmail_ attribute, if not null, added as comment to the cell; each _Aliasmail_ element separated with ';')
* _Company_
* _Department_
* _Office_
* _EmployeeNumber_
* _MobilePhone_
* _primaryGroupID_
* _EmployeeID_
* _msNPAllowDialin_
* _generationQualifier_
* _EmployeeType_
* _userWorkStations_
* _Manager_

The above attributes are important to me. Many of them might not be important for someone else so `Select-Object -Property` should be changed. Formatting hardcodes column IDs (does not use named ranges) so those might need to be changed if the columns to be exported are changed.

Formatting will be applied to the exported columns:
* If _sAMAccountName_ is not _GivenName_._SurName_, use red text.
* If _Mail_ is not _sAMAccountName_@`$maildomain`, use red text.
* If _Displayname_ is not _GivenName_ _SurName_, use red text.
* If _primaryGroupID_ is Domain Users and _Title_ does not exist, use red text.

_UserAccountControl_ attribute is decoded to flag names and select flags invoke formatting; they stack if multiple conditions apply:
* Blue strikethru text if account is disabled.
* Blue italic text if account has an expired password.
* Blue underlined text if account is locked.
* Yellow background if flags affecting security: `PASSWD_NOTREQD` / `USE_DES_KEY_ONLY` / `DONT_REQ_PREAUTH`.

## TODO
* Fix _PwdLastSet_ set to 0 to display something sensible.
* Make `$maildomain` empty by default and check _Mail_ attribute validity only if it is set.
* _Mail_ attribute check should probably use _GivenName_._SurName_ instead of _sAMAccountName_.
* Privileged/special _primaryGroupID_ might need highlighting.
* Group memberships should be exported somehow. Perhaps to another worksheet as a matrix with _sAMAccountName_ at Y and group name (printed vertically) at X.
