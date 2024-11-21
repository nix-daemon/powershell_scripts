param (
    [string]$SourceDir = "C:\Source", # Replace with your source directory
    [string]$TargetDir = "C:\Target" # Replace with your target directory
)

# Ensure the target directory exists
if (!(Test-Path -Path $TargetDir)) {
    New-Item -ItemType Directory -Path $TargetDir -Force
}

# Get all .doc and .docx files recursively from the source directory
$files = Get-ChildItem -Path $SourceDir -Recurse -File -Include *.doc, *.docx

foreach ($file in $files) {
    # Construct the target directory path
    $relativePath = $file.DirectoryName.Substring($SourceDir.Length).TrimStart("\")
    $targetPath = Join-Path -Path $TargetDir -ChildPath $relativePath

    # Create the target directory if it doesn't exist
    if (!(Test-Path -Path $targetPath)) {
        New-Item -ItemType Directory -Path $targetPath -Force
    }

    # Copy the file to the target directory
    Copy-Item -Path $file.FullName -Destination $targetPath -Force
    Write-Host "Copied: $($file.FullName) to $targetPath"
}
