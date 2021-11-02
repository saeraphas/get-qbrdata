# get-qbrdata
formerly DQDQBRDC (Douglas' Quick and Dirty QBR Data Collector)

### INTRODUCTION
This script collects data from a Windows AD domain controller and builds a number of reports intended for customer review as part of periodic housekeeping, license audits, true-ups, etc.

### PREREQUISITES:
This script depends on WMF 4 and the ActiveDirectory Powershell module,  and is intended to run from a domain controller, though any domain member server or workstation with the RSAT-AD-PowerShell module installed should work.
```
Install-WindowsFeature RSAT-AD-PowerShell
```

I have personally tested this script on Windows Server 2008R2 with WMF 4 and higher though older platforms may also work. Your mileage may vary.

### USAGE:

```
.\Get-QBRData.ps1
```

There are currently no parameters or switches, but the reports available are dependant on the user account that executes the script.

When run with a local SYSTEM account, this script produces
- AD users report with last login date and last password date
- AD custom groups membership report
- inactive users report
- inactive computers report
- workstation OS end-of-life inventory
- domain admin group membership report

When run with a domain admin account, this script additionally produces
- server storage space report 
- service accounts report
- server interface nameserver report

Reports are generated in both HTML and CSV format.
- HTML reports are generally suitable for review by non-technical customer contacts and can be used in a print>strikethrough>scanback flow.
- CSV reports are generally suitable for review by customer contacts who have some familiarity with Excel, and can be used for automation if you want.

### LICENSE:
This script is distributed under "THE BEER-WARE LICENSE" (Revision 42):
As long as you retain this notice you can do whatever you want with this stuff.
If we meet some day, and you think this stuff is worth it, you can buy me a beer in return.

### CONTACT:
Douglas Hammond (douglas@douglashammond.com)

### CHANGELOG:
| Date | Note |
| --- | --- |
| 2020-08-28 | Added script header for first github version.|
| 2020-09-02 | Added check for NT AUTHORITY\SYSTEM to prevent ugly errors on storage report.|
| 2020-09-02 | Added check to remove old output ZIP before creating a new one. |
| 2020-09-14 | Added services report, Expanded customization of HTML header |
| 2020-09-16 | Fixed services report, output works as intended now. |
| 2020-10-13 | Added passwordlastset to domain admins report. |
| 2020-10-13 | Added usersaudit report by request. |
| 2020-10-19 | Added EOL OS report, moved and renamed per-report variables for legibility. |
| 2020-10-20 | Added LastLogonDate to EOL OS report, only collect enabled accounts in usersaudit. |
| 2020-12-02 | The McRib update: more meat! (moved report building to a function that always generates HTML and CSV). Also fixed Domain Admins report, output now correctly includes nested groups. |
| 2020-12-10 | Added check for empty report data to prevent console error and 0-byte output files. |
| 2021-01-19 | Added nameserver report.|
| 2021-01-21 | Moved changelog and some remarks from header to README.md |
| 2021-06-04 | Added share and cert reports |
| 2021-10-29 | Revised AD group report to support 5000+ users |
| 2021-11-01 | Revised server reports to only attempt data collection from online servers |
| 2021-11-02 | Replaced CSV output type with Excel |

### ROADMAP:
- 365 account assigned licenses
- 365 account MFA status
- 365 account onprem sync status
- 365 account litigation hold status
