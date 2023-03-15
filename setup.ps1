az login
$appId = az ad app create --display-name "MSGraph console app" --public-client-redirect-uris "http://localhost" --query appId -o tsv
((Get-Content -path Program.cs -Raw) -replace '512f9b03-dfa2-4e01-8391-265b9a8ee064',$appId) | Set-Content -Path Program.cs