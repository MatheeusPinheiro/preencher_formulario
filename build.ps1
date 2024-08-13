$exclude = @("venv", "preencher_form.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "preencher_form.zip" -Force