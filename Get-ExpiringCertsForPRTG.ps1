######################################################################
## (C) 2020 Michael Miklis (michaelmiklis.de)
##
##
## Filename:      Get-ExpiringCertsForPRTG.ps1
##
## Version:       1.0
##
## Release:       Final
##
## Requirements:  PSPKI PowerShell Module (https://github.com/PKISolutions/PSPKI)
##
## Description:   Script used as PRTG monitoring sensor for expiring
##                certificates on a Windows ADCS instance
##
## This script is provided 'AS-IS'.  The author does not provide
## any guarantee or warranty, stated or implied.  Use at your own
## risk. You are free to reproduce, copy & modify the code, but
## please give the author credit.
##
####################################################################

<#
    .SYNOPSIS
    Lists all expiring certificates in a PRTG compliant JSON output

    .DESCRIPTION
    The Get-ExpiringCertsForPRTG CMDlet gets all expiring certificates
    within the specified timespan and creates a JSON output for usage
    as a PRTG sensor. Details about the JSON structure can be found in
    the PRTG manuel at https://www.paessler.com/manuals/prtg/custom_sensors.
  
    .PARAMETER CAComputername
    FQDN to the CA (Use Get-CertificationAuthority to get the name)

    .PARAMETER StartDate
    DateTime for the beginning of the timespan.

    .PARAMETER EndDate
    DateTime for the end of the timespan.

    .PARAMETER ExcludeTemplateList
    Certificate template names to exclude from result

    .PARAMETER ExcludeAutoEnrollEnabledTemplates
    Switch to exclude Auto-Enroll enabled certificate templates from result

    .PARAMETER LimitMinWarning
    Warning threshold in days (changing the warning threshold requires the PRTG sensor to be re-created)

    .PARAMETER LimitMinError
    Error threshold in days (changing the warning threshold requires the PRTG sensor to be re-created)
  
    .PARAMETER ResultSize
    Number of certificates returned in the result

      
    .PARAMETER ReturnIndex
    Index of certificate to be returned for sensor data (value must not be greater than ResultSize)
   
    .EXAMPLE
    Get-ExpiringCertsForPRTG.ps1 -CAComputername "My-CA.domain.local" -StartDate $(Get-Date) $EndDate $((Get-Date).AddMonths(3)) -ExcludeTemplateList "MDM_Exchange_auth;CAExchange" -ExcludeAutoEnrollEnabledTemplates -LimitMinWarning 15 -LimitMinError 30 -ResultSize 10 -ReturnIndex 1
#>


[CmdletBinding()] 
param (
    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()][string]$CAComputername,
    [parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()][DateTime]$StartDate = $((Get-Date)),
    [parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()][DateTime]$EndDate = $((Get-Date).AddMonths(12)),
    [parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()][String]$ExcludeTemplateList = "",
    [parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()][Switch]$ExcludeAutoEnrollEnabledTemplates,
    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()][int]$LimitMinWarning,
    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()][int]$LimitMinError,
    [parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()][int]$ResultSize = 10,
    [parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()][int]$ReturnIndex = 0  
)


Set-StrictMode -Version latest


# requires PSPKI PowerShell Module (tested with version 3.5)
# (https://github.com/PKISolutions/PSPKI)
Import-Module PSPKI

# get ca context
$CA = Get-CertificationAuthority -ComputerName $CAComputername

# get all certificates issued in the specified time frame
$IssuedCerts = Get-IssuedRequest -CertificationAuthority $CA -Filter "NotAfter -ge $StartDate", "NotAfter -le $EndDate" | Sort-Object -Property NotAfter

# remove auto enrolled certificates if specified as parameter
if ($ExcludeAutoEnrollEnabledTemplates)
{
    $AutoEnrollEnabledTemplates = Get-CertificateTemplate | Where-Object { $_.AutoenrollmentAllowed -eq $true }
}

else
{
    $AutoEnrollEnabledTemplates = New-Object PSObject -Property @{
			Name = "";
			DisplayName = "";
			OID = New-Object PSObject -Property @{
			            Value = "";
                  }
		}
}


# remove templates that are specified as exclude list parameter
if ($ExcludeTemplateList -ne "")
{
    $ExcludeTemplateArray = $ExcludeTemplateList.Split(";")
}
else
{
    $ExcludeTemplateArray = @() 
}

# initialized output / result arra
[array]$prtgObjectArray = @() 
$i = 0

# loop trogh variables
foreach ($Cert in $IssuedCerts)
{

    # some v1 templates do not have a display / friendly name - using OID for display instead
    if ($null -eq $Cert.CertificateTemplateOid.FriendlyName)
    {
        $Cert.CertificateTemplateOid.FriendlyName = $Cert.CertificateTemplateOid.Value
    }


    # ignore Autoenrolled templates
    if ($($AutoEnrollEnabledTemplates.OID | Select-Object -ExpandProperty Value).Contains($Cert.CertificateTemplateOid.Value) -eq $true)
    {
        continue
    }

    
    # ignore templates from ignore list (friendly name)
    if ($Cert.CertificateTemplateOid.FriendlyName -ne "")
    {
        if ($ExcludeTemplateArray.Contains($Cert.CertificateTemplateOid.FriendlyName) -eq $true)
        {
            continue
        }

    }
      
                
    # build prtg sensor JSON output string (PSObject -> JSON)
    #
    # JSON structure sample:
    #
    # {
    # "prtg":  {
    #        "result":  [
    #                        {
    #                             "Channel":  "Certificate expiration",
    #                            "Value":  6915,
    #                            "Unit":  "Custom",
    #                            "CustomUnit":  "Days",
    #                            "Float":  0,
    #                            "LimitMinWarning":  15,
    #                            "LimitMinError":  30,
    #                            "LimitMode":  1
    #                        },
    #                        {
    #                            "Channel":  "Placeholder",
    #                            "Value":  0
    #                        }
    #                    ],
    #         "text":  "Demo-Certificate expires in 6915 Days"
    #     }
    # }

    $channelObject = New-Object -TypeName PSObject 
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "Channel" -Value "Certificate expiration"
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "Value" -Value $([math]::truncate((New-TimeSpan -Start $(Get-Date) -End $Cert.NotAfter).TotalDays))
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "Unit" -Value "Custom"
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "CustomUnit" -Value "Days"
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "Float" -Value 0
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "LimitMinWarning" -Value $LimitMinWarning
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "LimitMinError" -Value $LimitMinError
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "LimitMode" -Value 1

    # increment
    $i++

    [array]$channelObjectArray = @()
    $channelObjectArray += $channelObject


    $channelObject = New-Object -TypeName PSObject 
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "Channel" -Value "Placeholder"
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "Value" -Value 0

    $channelObjectArray += $channelObject

    $resultObject = New-Object -TypeName PSObject
    Add-Member -InputObject $resultObject -MemberType NoteProperty -Name "result" -Value $($channelObjectArray)
    Add-Member -InputObject $resultObject -MemberType NoteProperty -Name "text" -Value $("{0} ({1}) expires in {2} Days" -f $Cert.CommonName, $Cert.CertificateTemplateOid.FriendlyName, $([math]::truncate((New-TimeSpan -Start $(Get-Date) -End $Cert.NotAfter).TotalDays)))

    $prtgObject = New-Object -TypeName PSObject
    Add-Member -InputObject $prtgObject -MemberType NoteProperty -Name "prtg" -Value $resultObject


    # add PSObject to result array
    $prtgObjectArray += $prtgObject


    # break loop if ResultSize is reached
    if ($i -eq $ResultSize - 1)
    {
        break
    }

}

# return empty prtg structure if ReturnIndex is out of range
if (($ReturnIndex -gt $prtgObjectArray.Count) -or ($prtgObjectArray.Count -eq 0))
{
    $channelObject = New-Object -TypeName PSObject 
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "Channel" -Value "Certificate expiration"
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "Value" -Value 99999
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "Unit" -Value "Custom"
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "CustomUnit" -Value "Days"
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "Float" -Value 0
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "LimitMinWarning" -Value 0
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "LimitMinError" -Value 0
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "LimitMode" -Value 1

    [array]$channelObjectArray = @()
    $channelObjectArray += $channelObject


    $channelObject = New-Object -TypeName PSObject 
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "Channel" -Value "Placeholder"
    Add-Member -InputObject $channelObject -MemberType NoteProperty -Name "Value" -Value 0

    $channelObjectArray += $channelObject

    $resultObject = New-Object -TypeName PSObject
    Add-Member -InputObject $resultObject -MemberType NoteProperty -Name "result" -Value $($channelObjectArray)
    Add-Member -InputObject $resultObject -MemberType NoteProperty -Name "text" -Value $("ReturnIndex {0} out of range - Result only contains {1} elements" -f $ReturnIndex, $prtgObjectArray.Count)

    $prtgObject = New-Object -TypeName PSObject
    Add-Member -InputObject $prtgObject -MemberType NoteProperty -Name "prtg" -Value $resultObject

    $prtgObject | ConvertTo-Json -Depth 10

}

# retrun JSON structure
else
{
    $prtgObjectArray[$ReturnIndex] | ConvertTo-Json -Depth 10
} 
