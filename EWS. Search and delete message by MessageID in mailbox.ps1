# EWS. Search and delete message by MessageID in mailbox
$findResults = $view = $service = $Mailbox = $inbox = $mailboxId = $folderId = $unread_count = $result = $null

# Users accounts
$users = 'test1', 'test2', 'test3'
# Domain name
$domain = '@bgelov.ru'
# Message ID
$searchString = '<9ecd7713913f4489b78136fe93365017@bgelov.ru>'
# Message FROM
$from = 'testmailbox1@bgelov.ru'
# Trigger, Delete message or not
[bool]$delete = $false

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


# Search message filter
$SearchFilterCollection = @()
$SearchFilter1 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageId,$searchString)
$SearchFilter2 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::From,$from)
$SearchFilterCollection = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
$SearchFilterCollection.Add($SearchFilter1)
$SearchFilterCollection.Add($SearchFilter2)
# How many results display
$iv = new-object Microsoft.Exchange.WebServices.Data.ItemView(50)
$propertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 

foreach ($u in $users) {

    $mailbox = $u + $domain

	    #check Anchor header for Exchange 2013/Office365
	    if($service.HttpHeaders.ContainsKey("X-AnchorMailbox")){
		    $service.HttpHeaders["X-AnchorMailbox"] = $mailbox
	    }else{
		    $service.HttpHeaders.Add("X-AnchorMailbox", $mailbox);
	    }
	    #"AnchorMailbox : " + $service.HttpHeaders["X-AnchorMailbox"]

    $mailboxId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.Webservices.Data.WellKnownFolderName]::MsgFolderRoot, $mailbox)

    $findResults = $service.FindFolders($mailboxId, $sfSearchFilter, $view)

    foreach ($f in $findResults) {
    
        $folderName = $f.DisplayName
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId($f.id)
        $inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderId)

        if ($inbox.TotalCount -gt 0) {

            $messages = $service.FindItems($Inbox.Id, $SearchFilterCollection, $iv)

            # Each email message
            if($messages.TotalCount -gt 0){
                foreach($message in $messages){
       
                    Write-Host "Found message! Message ID is $($message.Id.UniqueId)"
                    if($delete -eq $true){
                        $emailMessage = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service,$Message.Id.UniqueId,$propertySet)
                        # We can use HardDelete delete message or SoftDelete
                        $emailMessage.Delete("HardDelete")
                    }
                    Break # Exit
                }

            }else{
                Write-Host "Search returned no results"
            }
        }
        $inbox = $folderId = $unread_count = $folderName = $unread_count = $str = $null
    }
    $findResults = $mailboxId = $Mailbox = $null
} #foreach user
