# Get message by messageID

$findResults = $view = $service = $Mailbox = $inbox = $mailboxId = $folderId = $unread_count = $result = $null

# Message ID
$messageID = '<48d2c0c371e248abbf008ca4fe80b1a0@bgelov.ru>'
# Username
$u = 'testmailbox1'
# Domain
$domain = '@bgelov.ru'
# Generate mail address
$mailbox = $u + $domain

# EWS
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016_SP1)
$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind
$service.AutodiscoverUrl($aceuser.mail.ToString())

    $view = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
    $view.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.Webservices.Data.BasePropertySet]::FirstClassProperties)
    $view.PropertySet.Add([Microsoft.Exchange.Webservices.Data.FolderSchema]::DisplayName)


    $view.Traversal = [Microsoft.Exchange.Webservices.Data.FolderTraversal]::Deep

    # Search filter will exclude any Search Folders 
    $PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
    $sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")

	    # Check Anchor header for Exchange 2013/Office365
	    if($service.HttpHeaders.ContainsKey("X-AnchorMailbox")){
		    $service.HttpHeaders["X-AnchorMailbox"] = $mailbox
	    }else{
		    $service.HttpHeaders.Add("X-AnchorMailbox", $mailbox);
	    }
	    #"AnchorMailbox : " + $service.HttpHeaders["X-AnchorMailbox"]

    $inbox = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.Webservices.Data.WellKnownFolderName]::Inbox, $mailbox)

$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(10)

$searchFilter =  new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageId,$messageID)

$findResults = $service.FindItems($inbox,$searchFilter,$view)

$findResults
