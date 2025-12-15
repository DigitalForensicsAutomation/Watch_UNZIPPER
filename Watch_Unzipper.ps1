# === CONFIGURE PATHS ===
$sourceFolder      = "C:\Transfer\Target"        # Folder with .001/.002 split archives
$extractFolder     = "C:\Transfer\TempExtract"   # Temporary extraction folder
$finalFolder       = "C:\Transfer\Completed"    # Final folder for extracted files
$sevenZipPath      = "C:\Program Files\7-Zip\7z.exe"  # Path to 7-Zip executable
$pollInterval      = 30                           # Time in seconds between folder scans

# Ensure required folders exist
$foldersToCheck = @($extractFolder, $finalFolder, "$sourceFolder\Processed", "$sourceFolder\Failed")
foreach ($f in $foldersToCheck) { if (-not (Test-Path $f)) { New-Item -ItemType Directory -Path $f | Out-Null } }

# === FUNCTION: Check if file is locked ===
function Test-FileLocked {
    param([string]$file)
    try {
        $stream = [System.IO.File]::Open($file,'Open','ReadWrite','None')
        if ($stream) { $stream.Close(); return $false }
    } catch {
        return $true
    }
}

# === FUNCTION: Process Archives ===
function Process-Archives {
    $archives = Get-ChildItem -Path $sourceFolder -Filter "*.001" -File | Sort-Object LastWriteTime

    foreach ($archive in $archives) {
        $archivePath = $archive.FullName
        $archiveName = [System.IO.Path]::GetFileNameWithoutExtension($archive.Name)

        if (Test-FileLocked $archivePath) {
            Write-Host "$(Get-Date -Format HH:mm:ss) - Skipping locked file: $archivePath"
            continue
        }

        # Extract to temp folder
        $extractTarget = Join-Path $extractFolder $archiveName
        if (-not (Test-Path $extractTarget)) { New-Item -ItemType Directory -Path $extractTarget | Out-Null }

        Write-Host "$(Get-Date -Format HH:mm:ss) - Extracting: $archivePath"
        & "$sevenZipPath" x "$archivePath" -o"$extractTarget" -y

        if ($LASTEXITCODE -eq 0) {
            Write-Host "$(Get-Date -Format HH:mm:ss) - Extraction successful: $archivePath"

            # Move extracted files to final folder
            $extractedFiles = Get-ChildItem -Path $extractTarget -Recurse -File
            foreach ($file in $extractedFiles) {
                $relativePath = $file.FullName.Substring($extractTarget.Length).TrimStart('\')
                $destination = Join-Path $finalFolder $relativePath

                $destDir = Split-Path $destination
                if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }

                if (Test-Path $destination) {
                    Write-Host "$(Get-Date -Format HH:mm:ss) - Duplicate found, renaming: $($file.Name)"
                    $destination = Join-Path $destDir ("COPY_" + $file.Name)
                }

                Move-Item -Path $file.FullName -Destination $destination
            }

            # Clean up temp folder
            Remove-Item -Path $extractTarget -Recurse -Force

            # Move original archive to Processed
            $processedPath = Join-Path "$sourceFolder\Processed" $archive.Name
            Move-Item -Path $archivePath -Destination $processedPath
        } else {
            Write-Host "$(Get-Date -Format HH:mm:ss) - Extraction failed: $archivePath"
            $failedPath = Join-Path "$sourceFolder\Failed" $archive.Name
            Move-Item -Path $archivePath -Destination $failedPath
        }
    }
}

# === WATCH FOLDER LOOP ===
Write-Host "$(Get-Date -Format HH:mm:ss) - Watch folder started. Monitoring $sourceFolder every $pollInterval seconds."

while ($true) {
    try {
        Process-Archives
    } catch {
        Write-Host "$(Get-Date -Format HH:mm:ss) - ERROR: $_"
    }

    Start-Sleep -Seconds $pollInterval
}
