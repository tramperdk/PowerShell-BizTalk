[void] [System.reflection.Assembly]::LoadWithPartialName("Microsoft.BizTalk.ExplorerOM")
$Catalog = New-Object Microsoft.BizTalk.ExplorerOM.BtsCatalogExplorer
$mgmtDB = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0\Administration').MgmtDBName
$sqlServer = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0\Administration').MgmtDBServer
$Catalog.ConnectionString = "SERVER=$sqlServer;DATABASE=$mgmtDB;Integrated Security=SSPI"