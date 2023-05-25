# Template to get teamcity information in excel

This template can be used to run a powershell script which creates an excel that will list all builds information of teamcity pipeline with Projects from a paticular data.

## Prerequisite

* Administrator Access
* Powershell version >= 5.1
* **Cross check on Powershell**
    $PSVersion
* start the PS script:

```pwsh
    .\build_data.ps1 -TeamcityToken ihd98d12y6316155#2297 -Server_Url https://teamcity.com -startdate "20230101"
```
