$nugetUrl = "https://nexus.tools.cihs.gov.on.ca/repository/nexusraw/Microsoft.PackageManagement.NuGetProvider.2.8.5.208.zip"
$localFolder = "$env:LOCALAPPDATA\PackageManagement\ProviderAssemblies\nuget\2.8.5.208"
if(-not (Test-Path $localFolder)){
    Write-Host "Creating $localFolder"
    New-Item -Type "Directory" -Path $localFolder
}
$localFilePath = Join-Path $localFolder "Microsoft.PackageManagement.NuGetProvider.dll"
if(-not (Test-Path $localFilePath)){
    Write-Host $nuget Provider does not exist. Downloading from Nexus
    $clnt = new-object System.Net.WebClient
    $filePath = Join-Path $env:TEMP "Microsoft.PackageManagement.NuGetProvider.2.8.5.208.zip"
    $clnt.DownloadFile($nugetUrl,$filePath)

    # Unzip the file to specified location
    $shell_app=new-object -com shell.application 
    $zip_file = $shell_app.namespace($filePath) 
    $destination = $shell_app.namespace($localFolder) 
    $destination.Copyhere($zip_file.items())
}
