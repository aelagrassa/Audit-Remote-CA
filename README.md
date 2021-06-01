What will this script do?

Connect to specified Issuing Certification Authorities to generate a CSV report that includes all certificates issued by the CA. It will create a seperate CSV for each Issuing CA audited by the script.

What are the requirements?

Network connectivity to the CA(s) you are looking to audit as well as the PSPKI module on the computer performing the audit.

How do I run this?

Use the PS file and update the following variables to match your environment:

$CertificationAuthorities
$CSVFolder
$PageSize
