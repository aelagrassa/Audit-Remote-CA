################################################################################

###
#    Variables to set environment. You will need to update these.
###

#Specify which Certificate Authorities to querry
$CertificationAuthorities = @(
    "CA Computer Name 01",
    "CA Computer Name 02"
)

#Specify the folder in which you want reports to be saved
$CSVFolder = "C:\Path\To\Where\Reports\Should\Go"

#Specify Page Size. This is how many requests will be gathered at once from your CA.
#The module creator recommends 50000 requests at a time for best throughput
#Please see https://www.pkisolutions.com/adcs-certification-authority-database-query-numbers/ for details.
$PageSize = 250

###
#    End variables to set environment.
###

##############################################################################

Write-Host -ForegroundColor Yellow "###################################################################################"
Write-Host -ForegroundColor Yellow "Settings for Issuing Certification Authority Auditing Script:"
ForEach ($CertificationAuthority in $CertificationAuthorities)
    {
        Write-Host -Fore Yellow    "Certification Authority ........... $CertificationAuthority"
    }
Write-Host -ForegroundColor Yellow "Report Output Folder .............. $CSVFolder"
Write-Host -ForegroundColor Yellow "Page Size ......................... $PageSize"
Write-Host -ForegroundColor Yellow "###################################################################################"


#Gather Start Time
$StartTime = Get-Date

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

ForEach ($CertificationAuthority in $CertificationAuthorities)
    {
        Write-Host -ForegroundColor Cyan "INFO: Connecting to $CertificationAuthority..."
        Try 
            {
                $CertificationAuthorityObject = Connect-CertificationAuthority -ComputerName $CertificationAuthority
            }
        Catch
            {
                Write-Host -ForegroundColor Red "ERROR: Could not connect to $CertificationAuthority!"
                pause
                exit
            }

        $TemplatesArray = @{}

        Write-Host -Fore Cyan "INFO: Generating list of templates published on $CertificationAuthority..."
        $Templates = $CertificationAuthority | Get-IssuedRequest -Property CertificateTemplate | Select-Object -Property CertificateTemplate -Unique

        ForEach ($Template in $Templates.CertificateTemplate)
            {
                If ($Template -match "1.3.6.1.4.1")
                    {
                        Write-Host -Fore Cyan "INFO: $Template is likely an OID"
                        Try
                            {
                                $TemplateObject = Get-CertificateTemplate -OID $Template -ErrorAction Stop
                                Write-Host -Fore Green "INFO: Found $($TemplateObject.DisplayName) with OID $Template"
                                $TemplatesArray.add($Template,$($TemplateObject.DisplayName))
                            }
                        Catch
                            {
                                Write-Host -Fore Yellow "WARNING: Could not associate OID $Template with an existing template. It may have been deleted or is no longer used."
                                $TemplatesArray.add($Template,"Removed")
                            }                     
                    }
                Else
                    {
                        Write-Host -Fore Cyan "INFO: $Template is likely not an OID, using OID value as Display Name"
                        $TemplatesArray.add($Template,$Template)
                    }
            }

        Write-Host -ForegroundColor Cyan "INFO: Creating CSV report file for $CertificationAuthority..."
        Try 
            {
                $CSV = New-Item -ItemType File -Path "C:\Users\t0.alagrassa\Desktop\$CertificationAuthority Issued Certificates.csv" -Force
            }
        Catch
            {
                Write-Host -ForegroundColor Red "ERROR: Could not create report file for $CertificationAuthority!"
                pause
            }

        #Set Last ID to zero to start from the top
        $LastID = 0

        #Set CA Audit Start Time
        $CAAuditStartTime = Get-Date
        Write-Host -Fore Cyan "INFO: $CertificationAuthority audit started at $CAAuditStartTime"

        do {
            $ReadRows = 0
            Write-Host -fore Cyan "INFO: Gathering next $PageSize requests starting at request ID $LastID..."
            $CertificationAuthorityObject | Get-IssuedRequest -Filter "RequestID -gt $LastID" -Page 1 -PageSize $PageSize -Property * | %{
                
                #Iterate up to next row
                $ReadRows++

                #Define last ID used
                $LastID = $_.Properties["RequestID"]

                #Define variables to match template OID with given name in hash table
                $TemplateOID = $_.CertificateTemplate
                $TemplateClearname = $TemplatesArray.Get_Item($TemplateOID)

                #Create a custom object for output
                $OutputObject = [PSCustomObject]@{
                RequestID =                    $_.RequestID
                CommonName =                   $_.CommonName
                DistinguishedName =            $_.DistinguishedName
                CertificateTemplateOID =       $_.CertificateTemplate
                CertificateTemplateGivenName = $TemplateClearname
                NotBefore =                    $_.NotBefore
                NotAfter =                     $_.NotAfter
                GeneralFlags =                 $_.GeneralFlags
                PrivatekeyFlags =              $_.PrivatekeyFlags
                SerialNumber =                 $_.SerialNumber
                Country =                      $_.Country
                Organization =                 $_.Organization
                OrgUnit =                      $_.OrgUnit
                Locality =                     $_.Locality
                State =                        $_.State
                Title =                        $_.Title 
                GivenName =                    $_.GivenName
                Surname =                      $_.SurName
                DomainComponent =              $_.DomainComponent
                EMail =                        $_.EMail
                StreetAddress =                $_.StreetAddress
                PublicKeyLength =              $_.PublicKeyLength
                PublicKeyAlgorithm =           $_.PublicKeyAlgorithm
                PublishExpiredCertInCRL =      $_.PublishExpiredCertInCRL
                ConfigString =                 $_.ConfigString
                IssuingCA =                    $CertificationAuthority
                }
                $OutputObject | Export-CSV -Path $CSV -Append -Force -NoClobber -NoTypeInformation                 
             }
        } while ($ReadRows -eq $PageSize)
        Write-Host -Fore Green "INFO: All issued requests for $CertificationAuthority have been audited!"
        $CAAuditEndTime = Get-Date
        $CAAuditTimeSpan = New-TimeSpan -Start $CAAuditStartTime -End $CAAuditEndTime
        Write-Host -Fore Green "INFO: $CertificationAuthority Audit Time: $CAAuditTimeSpan"
    }

#Gather End Time
$EndTime = Get-Date
$ExecutionTime = New-TimeSpan -Start $StartTime -End $EndTime

Write-Host -Fore Green "INFO: Script completed"
Write-Host -Fore Green "INFO: Execution Time: $ExecutionTime"