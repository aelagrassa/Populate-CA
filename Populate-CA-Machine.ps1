################################################################################

###
#    Variables to set environment. You will need to update these.
###

#Which Issuing CA to use
$CertificationAuthority = "YourCA"

#Specify which templates to use. Go by Template Name, NOT Template Display Name
$Templates = @(
    "Template-01",
    "Template-02",
    "Template-03"
)

#How many certificates to generate?
$CertQuota = 10

###
#    End variables to set environment.
###

##############################################################################

Write-Host -ForegroundColor Yellow "###################################################################################"
Write-Host -ForegroundColor Yellow "Settings for certificate generation script:"
Write-Host -ForegroundColor Yellow "Issuing Certification Authority ... $CertificationAuthority"
ForEach ($Template in $Templates)
    {
        Write-Host -Fore Yellow    "Template .......................... $Template"
    }
Write-Host -ForegroundColor Yellow "Certs to Issue .................... $CertQuota"
Write-Host -ForegroundColor Yellow "###################################################################################"



#Import Required Modules  [PSPKI]
$RequiredModules = @(
    "PSPKI"
)

#Gather Start Time
$StartTime = Get-Date

#Announce module import and check
Write-Host -ForegroundColor Cyan "INFO: Performing module import and check..."

ForEach ($RequiredModule in $RequiredModules)
    {
        Write-Host -Fore Cyan "INFO: Importing $RequiredModule Module..."
        Try {
                Import-Module $RequiredModule -ErrorAction Stop
            }
        Catch
            {
                Write-Host -Fore Red "ERROR: $($RequiredModule) failed to import. Press enter to terminate script."
                pause
                exit
            }
    }

$ImportedModule = Get-Module -Name $RequiredModule
        if ($ImportedModule.Name -Match "$RequiredModule")
            {
                Write-Host -ForegroundColor Green "INFO: Found $RequiredModule Module, continuing..."
            }
        else
            {
                Write-Host -ForegroundColor Red "ERROR: $RequiredModule did not throw an import error but could not be found. Press enter to terminate script."
                pause
                exit
            }

Write-Host -Fore Green "INFO: All required modules found, continuing..."

#Set Certs Issued count value to 0
$CertsIssued = 0

Write-Host -Fore Yellow "###################################################################################"

do
{
#Iterate Certs Issued count up
$CertsIssued++

Write-Host -Fore Cyan "Starting issuance of certificate $CertsIssued of $CertQuota"

#Define Random String to Server as Common Name for Certificate
$RandomizedName = ( -join ((0x30..0x39) + ( 0x41..0x5A) + ( 0x61..0x7A) | Get-Random -Count 16  | % {[char]$_}) )

#Define Template
$Template = $Templates | Get-Random

Write-Host -Fore Cyan "Generating CSR: Template = $Template | Common Name = $RandomizedName"

$CertName = "Test-Certificate-$RandomizedName"
$CSRPath = "C:\Scripts\Mass Certificate Generation\Temp\$($CertName).csr"
$INFPath = "C:\Scripts\Mass Certificate Generation\Temp\$($CertName).inf"

$INF =
@"
[Version]
Signature= '$Windows NT$' 

[NewRequest]
Subject = "CN=$RandomizedName"
KeySpec = 1
KeyLength = 2048
Exportable = TRUE
MachineKeySet = TRUE
SMIME = False
PrivateKeyArchive = FALSE
UserProtected = FALSE
UseExistingKeySet = FALSE
ProviderName = "Microsoft RSA SChannel Cryptographic Provider"
ProviderType = 12
RequestType = PKCS10
KeyUsage = 0xa0

[RequestAttributes]
CertificateTemplate=$Template

[EnhancedKeyUsageExtension]

OID=1.3.6.1.5.5.7.3.1 
"@

$INF | out-file -filepath $INFPath -force
certreq -new $INFPath $CSRPath | Out-Null

Write-Host -Fore Cyan "INFO: Submitting CSR for certificate $CertsIssued of $CertQuota"

Submit-CertificateRequest -Path $CSRPath -CertificationAuthority $CertificationAuthority | Out-Null

Write-Host -Fore Cyan "INFO: Cleaning up temp files for certificate $CertsIssued of $CertQuota"
Remove-Item $INFPath
Remove-Item $CSRPath

Write-Host -Fore Green "INFO: Completed issuance of certificate $CertsIssued of $CertQuota"
}
until($CertsIssued -eq $CertQuota)

#Gather End Time
$EndTime = Get-Date
$ExecutionTime = New-TimeSpan -Start $StartTime -End $EndTime

Write-Host -Fore Green "INFO: Script completed"
Write-Host -Fore Green "INFO: Execution Time: $ExecutionTime"