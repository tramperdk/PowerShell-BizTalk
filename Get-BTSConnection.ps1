<#
.SYNOPSIS
    Connects to Microsoft.BizTalk.ExplorerOM on localhost
.DESCRIPTION
    Connects to the Assembly called Microsoft.BizTalk.ExplorerOM, this assembly is found on BizTalk servers.
.EXAMPLE
    PS C:\>. .\Get-BTSConnection.ps1
    This example returns a [bool] and a message about connectivity.
.INPUTS
    N/A
.OUTPUTS
    $Return
.NOTES
    General notes
#>
function Get-BTSConnection {
    try {
        # Prep hashtable to hold our return values.
        [hashtable]$Return = @{}
    
        # Load ExplorerOM.
        if([bool] [System.reflection.Assembly]::LoadWithPartialName("Microsoft.BizTalk.ExplorerOM")) {
            $Catalog = New-Object Microsoft.BizTalk.ExplorerOM.BtsCatalogExplorer
        }
        else {
            $Return.Connected   = $false
            $Return.Message     = "Unable to load Microsoft.BizTalk.ExplorerOM, are you running script on a BizTalk server?"
        }
        
        # Fetch local BizTalk settings.
        $mgmtDB    = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0\Administration' -ErrorAction SilentlyContinue).MgmtDBName
        $sqlServer = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0\Administration' -ErrorAction SilentlyContinue).MgmtDBServer
    
        if(($mgmtDB -eq $null) -or ($sqlServer -eq $null)) {
            $Return.Connected   = $false
            $Return.Message     = "mgmtDB or sqlServer variable is null, are you running script on a BizTalk server?"
        }
    
        # Use previously gathered information in Catalog connectionstring.
        $Catalog.ConnectionString = "SERVER=$sqlServer;DATABASE=$mgmtDB;Integrated Security=SSPI" 
    
        if ($Catalog) {
            $Return.Connected = $true
            $Return.Message = "Connected"   
        }
    }
    catch {
        $Return.Connected = $false
        $Return.Message = $_.Exception.Message
    }
    
    return $Return    
}