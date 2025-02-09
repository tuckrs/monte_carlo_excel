# PowerShell script to install Excel add-in
param(
    [string]$AddinPath,
    [string]$CustomUIPath
)

# Import Office COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

try {
    # Get the add-ins folder path
    $addinsFolder = [System.Environment]::GetFolderPath('ApplicationData') + "\Microsoft\AddIns"
    
    # Create the add-ins folder if it doesn't exist
    if (-not (Test-Path $addinsFolder)) {
        New-Item -ItemType Directory -Path $addinsFolder -Force
    }
    
    # Copy the add-in to the AddIns folder
    $targetPath = Join-Path $addinsFolder "monte_carlo.xlam"
    Copy-Item -Path $AddinPath -Destination $targetPath -Force
    
    # Register the add-in
    $addins = $excel.AddIns2
    $addin = $addins.Add($targetPath, $false)
    $addin.Installed = $true
    
    # Add registry entries for the custom UI
    $registryPath = "HKCU:\Software\Microsoft\Office\Excel\AddIns\MonteCarloExcel"
    New-Item -Path $registryPath -Force
    New-ItemProperty -Path $registryPath -Name "Description" -Value "Monte Carlo Simulation Add-in for Excel" -PropertyType String -Force
    New-ItemProperty -Path $registryPath -Name "FriendlyName" -Value "Monte Carlo Excel" -PropertyType String -Force
    New-ItemProperty -Path $registryPath -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force
    
    Write-Host "Excel add-in installed successfully"
} catch {
    Write-Error "Error installing Excel add-in: $_"
    exit 1
} finally {
    # Clean up
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
