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

The reports available are dependant on the user account used to run the script.

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

Reports are generated in both HTML and Excel format.
- HTML reports are generally suitable for review by non-technical customer contacts and can be used in a print>strikethrough>scanback flow.
- Excel reports are consolidated into a single tabbed workbook.

### LICENSE:
This script is distributed under "THE BEER-WARE LICENSE" (Revision 42):
As long as you retain this notice you can do whatever you want with this stuff.
If we meet some day, and you think this stuff is worth it, you can buy me a beer in return.

### CONTACT:
Douglas Hammond (douglas@douglashammond.com)
