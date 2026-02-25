**Python Based Clio Batch Upload/Download Processing and Dynamic RegEx Substitution Dictionary Mapping for Custom Fields**





**Clio Template Batch Download Processing - PowerShell Execution Script:**



python "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_templates\_sync.py" `

&nbsp; --token-file "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_tokens.json" `

&nbsp; download `

&nbsp; --source document-templates `

&nbsp; --output-dir "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\Template\_Download" `

&nbsp; --manifest "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_templates\_manifest.json"





**Clio Template RegEx Dictionary Mapping for Custom Fields - PowerShell Execution Script:**



python "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\mass\_update\_templates.py" `

&nbsp; --input-dir "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\Template\_Download" `

&nbsp; --output-dir "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\Template\_Upload" `

&nbsp; --excel "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\CustomField\_LookupTable.xlsx" `

&nbsp; --sheet "LookupTable" `

&nbsp; --old-col "Old\_Value" `

&nbsp; --new-col "New\_Value" `

&nbsp; --literal `

&nbsp; --join-runs `

&nbsp; --ignore-case



**Clio Template Batch Upload Processing - PowerShell Execution Script:**



python "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_templates\_sync.py" `

&nbsp; --token-file "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_tokens.json" `

&nbsp; --verbose `

&nbsp; upload `

&nbsp; --manifest "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_templates\_manifest.json" `

&nbsp; --upload-dir "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\Template\_Upload"







