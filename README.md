# Get-ExpiringCertsForPRTG

Monitor expiring certificates on your Windows ADCS PKI using PRTG (https://www.paessler.com/prtg). The script can get multiple expiring certificates and shows the common name, template name and the days until expiration in the status field of the PRTG sensor:

![Sensor Overview](https://github.com/michaelmiklis/Get-ExpiringCertsForPRTG/raw/assets/server-overview.PNG)

For checking the next three certificates that will expire, you'll need to create three individual sensors. Use individual parameters for each of the three sensors:

|  Sensor name  | Parameters |
| --------------| -----------|
| Ceritiface Expiration #1 |  `-ResultSize 3` and `-ReturnIndex 0`
| Ceritiface Expiration #2 |  `-ResultSize 3` and `-ReturnIndex 1`
| Ceritiface Expiration #3 |  `-ResultSize 3` and `-ReturnIndex 2`

More detailled parameter help can be found below.

## Creating a new sensor

### Step 1

![Add Sensor]( https://github.com/michaelmiklis/Get-ExpiringCertsForPRTG/raw/assets/add-sensor.PNG)

### Step 2

Choose the Get-ExpiringCertsForPRTG.ps1 from the drop down list and specify the parameters:
![Add Sensor Detail](https://github.com/michaelmiklis/Get-ExpiringCertsForPRTG/raw/assets/add-sensor-detail.PNG)

## Parameters

|  Parameter                          | Description |
| ----------------------------------- | ----------- |
`-CAComputername`                     | FQDN to the CA (Use Get-CertificationAuthority to get the name)
`-StartDate`                          | DateTime for the beginning of the timespan.
`-EndDate`                            | DateTime for the end of the timespan.
`-ExcludeTemplateList`                |  Certificate template names to exclude from result
`-ExcludeAutoEnrollEnabledTemplates`  | Switch to exclude Auto-Enroll enabled certificate templates from result
`-LimitMinWarning`                    | Warning threshold in days (changing the warning threshold requires the PRTG sensor to be re-created)
`-LimitMinError`                      | Error threshold in days (changing the warning threshold requires the PRTG sensor to be re-created)
`-ResultSize`                         | Number of certificates returned in the result
`-ReturnIndex`                        | Index of certificate to be returned for sensor data (value must not be greater than ResultSize)

## Example
The following command line show an example to return the first expiring certificate between now (omitted, as no -StartDate parameter specified) and the next 12 months (omitted, as no -EndDate parameter specified), the certificates generated from the template CAExchange and MDM_Exchange_auth as well as auto-enrollment enabled certificates will be excludeded from the result:

`Get-ExpiringCertsForPRTG.ps1 -CAComputername "My-CA.domain.local" -ExcludeTemplateList "MDM_Exchange_auth;CAExchange" -ExcludeAutoEnrollEnabledTemplates -LimitMinWarning 15 -LimitMinError 30 -ResultSize 10 -ReturnIndex 0`

The following command line show an example to return the second expiring certificate between 01/01/2020 and 12/31/2030, the certificates generated from the template CAExchange and MDM_Exchange_auth as well as auto-enrollment enabled certificates will be excludeded from the result:

`Get-ExpiringCertsForPRTG.ps1 -CAComputername "My-CA.domain.local" -StartDate 01/01/2020 $EndDate 12/31/2030 -ExcludeTemplateList "MDM_Exchange_auth;CAExchange" -ExcludeAutoEnrollEnabledTemplates -LimitMinWarning 15 -LimitMinError 30 -ResultSize 10 -ReturnIndex 2`


## Requirements

### PSPKI PowerShell Module

All interactions with the PKI is done via the PS Module PSPKI. This must be installed on the monitorining server / node, where the script ist being executed.

`PS> Install-Module PSPKI`

More information on PSPKI can be found in their Github repo:
<https://github.com/PKISolutions/PSPKI>

## Official PRTG documentation

The following links refer to the original PRTG documentation.

### PRTG Manual: EXE/Script Advanced Sensor

<https://www.paessler.com/manuals/prtg/exe_script_advanced_sensor>

### Advanced Script, HTTP Data, and REST Custom Sensors

<https://www.paessler.com/manuals/prtg/custom_sensors#advanced_sensors>