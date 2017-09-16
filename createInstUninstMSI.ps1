####################################################################
# Name:         createInstUninstMSI.ps1
# Description:  Script for creating install.cmd and uninstall.cmd 
#               for MSI package.
# Usage:        createInstUninstMSI.ps1 <full path to MSI package>        
####################################################################

param (
    [IO.FileInfo] $PathToMSI
)

$folderPath = $PathToMSI.DirectoryName
$fileName = $PathToMSI.Name

if (!(Test-Path $PathToMSI.FullName)) {
    throw "File '{0}' does not exist" -f $PathToMSI.FullName
}

try {
    $windowsInstaller = New-Object -com WindowsInstaller.Installer
    $database = $windowsInstaller.GetType().InvokeMember( "OpenDatabase", "InvokeMethod", $null,
    $windowsInstaller, @($PathToMSI.FullName, 0))
    $query = "SELECT Value FROM Property WHERE Property = 'ProductCode'"
    $view = $database.GetType().InvokeMember(
        "OpenView", "InvokeMethod", $null, $database, ($query)
    )
    $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null)
    $record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)
    $productCode = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 1)
} catch {
    echo "Exception while works with MSI DataBase: $_.Exception.GetType().FullName : $_.Exception.Message"    
	throw $_.Exception
} finally {
    $view.Close();
    ([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$database) -gt 0) > Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

write-output "msiexec `/i `"%~dp0$fileName`" `/qn" | out-file "$folderPath\install.cmd" -encoding ASCII
write-output "msiexec `/x $productCode /qn" | out-file "$folderPath\uninstall.cmd" -encoding ASCII

if (!(Test-Path $folderPath\install.cmd)) {
    throw "File install.cmd was not created."
}
if (!(Test-Path $folderPath\uninstall.cmd)) {
    throw "File uninstall.cmd was not created."
}
