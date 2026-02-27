**Python Based Clio Batch Upload/Download Processing and Dynamic RegEx Substitution Dictionary Mapping for Custom Fields**





**Clio Template Batch Download Processing - PowerShell Execution Script:**



python "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_templates\_sync.py" `

  --token-file "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_tokens.json" `

  download `

  --source document-templates `

  --output-dir "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\Template\_Download" `

  --manifest "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_templates\_manifest.json"







**Clio Template RegEx Dictionary Mapping for Custom Fields - PowerShell Execution Script:**



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





**Clio Template Batch Upload Processing - PowerShell Execution Script:**



python "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_templates\_sync.py" `

  --token-file "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_tokens.json" `

  --verbose `

  upload `

  --manifest "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_templates\_manifest.json" `

  --upload-dir "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\Template\_Upload" `

  --template-upload-mode create `

  --name-suffix "\_Updated\_{date}"

  --delete-old





python "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_templates\_sync.py" `

  --token-file "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_tokens.json" `

  --verbose `



  --manifest "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\clio\_templates\_manifest.json" `

  --upload-dir "C:\\Users\\Tim\\OneDrive - quillarrowlaw.com\\Documents\\ClioTemplates\_CustomFields\_MassUpdate\\Template\_Upload" `

  --delete-old

