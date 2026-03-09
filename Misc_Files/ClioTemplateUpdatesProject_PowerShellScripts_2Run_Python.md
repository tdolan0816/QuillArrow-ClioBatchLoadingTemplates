**Python Based Clio Batch Upload/Download Processing and Dynamic RegEx Substitution Dictionary Mapping for Custom Fields**







**Clio Template Batch Download Processing - Python File "clio\_templates\_sync.py"** **- PowerShell Execution Script:**



python "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_templates\_sync.py" `

  --token-file "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_tokens.json" `

  download `

  --source document-templates `

  --output-dir "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\Template\_Download" `

  --manifest "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_templates\_manifest.json"



\*The above PowerShell Script is for Downloading Word Documents (.docx) ***ONLY,*** In order to Download ***All*** files within the folder you need to include this additional flag:



**--include-non-docx**





**Validation Script for Determining All Custom Fields w/in a Template and All Templates w/in - Python File "verify\_template\_updates.py" - PowerShell Execution Script:**



python "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\inventory\_custom\_fields.py" `

  --input-dir "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\Template\_Download" `

  --output "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\custom\_field\_inventory.xlsx" `

  --deep-scan







**Clio Template RegEx Dictionary Mapping for Custom Fields - Python File "mass\_update\_templates.py" - PowerShell Execution Script:**



python "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\mass\_update\_templates.py" `

  --input-dir "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\Template\_Download" `

  --output-dir "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\Template\_Upload" `

  --excel "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\CustomField\_LookupTable.xlsx" `

  --sheet "LookupTable" `

  --old-col "Old\_Value" `

  --new-col "New\_Value" `

  --literal `

  --join-runs `

  --ignore-case





**Validation Script for the Verification of Custom Field Update - Python File "verify\_template\_updates.py" - PowerShell Execution Script:**



python "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\verify\_template\_updates.py" `

&nbsp; --input-dir "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\Template\_Upload" `

&nbsp; --excel "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\CustomField\_LookupTable.xlsx" `

&nbsp; --sheet "LookupTable" `

&nbsp; --old-col "Old\_Value" `

&nbsp; --new-col "New\_Value" `

&nbsp; --literal `

&nbsp; --ignore-case `

&nbsp; --deep-scan





**Clio Template Batch Upload Processing - Python File "clio\_templates\_sync.py"** **- PowerShell Execution Script:**



python "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_templates\_sync.py" `

&nbsp; --token-file "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_tokens.json" `

&nbsp; --verbose `

&nbsp; upload `

&nbsp; --manifest "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_templates\_manifest.json" `

&nbsp; --upload-dir "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\Template\_Upload" `

&nbsp; --template-upload-mode create `

&nbsp; --skip-unchanged `

&nbsp; --skip-invalid `

&nbsp; --delete-old





**Clio Batch Delete Processing - PowerShell Script:**



$token = (Get-Content ".\\clio\_tokens.json" | ConvertFrom-Json).access\_token

$headers = @{ Authorization = "Bearer $token" }



\# Example list of IDs

$ids = @(

"10233800",

"10233815",

"10233830",

"10233845",

"10233860"

)



foreach ($id in $ids) {

&nbsp; Invoke-RestMethod -Method Delete `

&nbsp;   -Headers $headers `

&nbsp;   -Uri "https://app.clio.com/api/v4/document\_templates/$id.json"

}







