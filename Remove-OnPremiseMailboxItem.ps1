Function Remove-OnPremiseMailboxItem 
{

  <#
    .SYNOPSIS
    "Remove-OnPremiseMailboxItem" function has been created to ease removal SPAM/unwanted messages from mailboxes.

    .DESCRIPTION 
    Generates a report or does the removal action for the specified server, database, mailbox or list of mailboxes.
    Use only one parameter at a time depending on the scope of your mailbox report / removal.

    .OUTPUTS
    Outputs of the file is a CSV file and HTML email report (if choosen to generate with -EmailReportTo switch)

    .EXAMPLE
    Remove-OnPremiseMailboxItem -ReadOnlyReportMode -SourceIsSingleExchMailbox testmailbox@domain.com -ItemKind meeting -ItemSubject "Monthly meeting" -ItemRecipient -EmailReportTo admin@domain.com
    
    .EXAMPLE
    Remove-OnPremiseMailboxItem -ReadOnlyReportMode -SourceIsExchServer WROEXCHSRV1 -ItemKind email -ItemSubject "Enlarge your P3n1s" -ItemTime "07/02/2017" -ItemSender SomeAddress@tryhere.com -EmailReportTo admin@domain.com

    .EXAMPLE
    Remove-OnPremiseMailboxItem -SourceIsExchDatabase WRO-EXCH-DB1 -ItemKind email -ItemSubject "Try our services 123" -ItemTime 07/02/2017 -ItemRecipient eva@domain.com -ItemSender SPAM@tryhere.com
    
    .EXAMPLE
    Remove-OnPremiseMailboxItem -SourceIsTxtFile "C:\folder\mailbox-list.txt" -ItemKind contacts -EmailReportTo admin@domain.com -WhatIf 

    .LINK
    https://paweljarosz.wordpress.com/2017/02/08/powershell-script-for-exchange-mailbox-item-email-meeting-contacts-etc-removal

    .NOTES
    Written By: Pawel Jarosz
    Website:	http://paweljarosz.wordpress.com
    GitHub:     https://github.com/zaicnupagadi
    Technet:    https://gallery.technet.microsoft.com/scriptcenter/site/mydashboard
    
    Change Log
    V1.00, 07/02/2017 - Initial version

#>


  [CmdletBinding(DefaultParameterSetName = 'Set 2', 
      SupportsShouldProcess = $true, 
      PositionalBinding = $false,
  ConfirmImpact = 'High')]
  [Alias('Remove-MailItem')]
  [OutputType([void])]
  
  Param
  (
    
    # Read only mode (report mode) switch. With this switch no items will be removed
    [Parameter(Mandatory = $false, 
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromRemainingArguments = $false)]
    [Parameter(Mandatory = $false, ParameterSetName = "Set 2")]
    [Parameter(Mandatory = $false, ParameterSetName = "Set 3")]
    [Parameter(Mandatory = $false, ParameterSetName = "Set 4")]
    [Parameter(Mandatory = $false, ParameterSetName = "Set 5")]
    [switch]
    $ReadOnlyReportMode,

    # Process only single mailbox, input - single mailbox
    [Parameter(Mandatory = $true, 
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromRemainingArguments = $false,    
    ParameterSetName = "Set 2")]
    [string]
    $SourceIsSingleExchMailbox,


    # Process mailboxes that are in txt file (smtp addresses, samaccountnames etc.), input - file
    [Parameter(Mandatory = $true, 
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromRemainingArguments = $false, 
    ParameterSetName = "Set 3")]
    [string]
    $SourceIsTxtFile,

    #Process mailboxes in one database, input - database name
    [Parameter(Mandatory = $false, 
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromRemainingArguments = $false,   
    ParameterSetName = "Set 4")]
    [string]
    $SourceIsExchDatabase,

    #Process mailboxes on one server, input - server name
    [Parameter(Mandatory = $false, 
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromRemainingArguments = $false,
    ParameterSetName = "Set 5")]
    [string]
    $SourceIsExchServer,

    #Message type (specify if item is regular email, meeting, contact etc. )
    [Parameter(Mandatory = $false, 
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromRemainingArguments = $false)]
    [Parameter(ParameterSetName = "Set 2")]
    [Parameter(ParameterSetName = "Set 3")]
    [Parameter(ParameterSetName = "Set 4")]
    [Parameter(ParameterSetName = "Set 5")]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("email","meeting","tasks","notes","contacts","im")]
    [string]
    $ItemKind,

    #Item subject - subject of searched item
    [Parameter(Mandatory = $false, 
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromRemainingArguments = $false)]
    [Parameter(ParameterSetName = "Set 2")]
    [Parameter(ParameterSetName = "Set 3")]
    [Parameter(ParameterSetName = "Set 4")]
    [Parameter(ParameterSetName = "Set 5")]
    [string]
    $ItemSubject,

    #Item time - specify when particular item has been received
    [Parameter(Mandatory = $false, 
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromRemainingArguments = $false)]
    [Parameter(ParameterSetName = "Set 2")]
    [Parameter(ParameterSetName = "Set 3")]
    [Parameter(ParameterSetName = "Set 4")]
    [Parameter(ParameterSetName = "Set 5")]
    [string]
    $ItemTime,
    
    #Item recipient - specify who was the recipient of the particular item
    [Parameter(Mandatory = $false, 
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromRemainingArguments = $false)]
    [Parameter(ParameterSetName = "Set 2")]
    [Parameter(ParameterSetName = "Set 3")]
    [Parameter(ParameterSetName = "Set 4")]
    [Parameter(ParameterSetName = "Set 5")]
    [string]
    $ItemRecipient,
        
    #Item sender - specify who was the sender of the particular item
    [Parameter(Mandatory = $false, 
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromRemainingArguments = $false)]
    [Parameter(ParameterSetName = "Set 2")]
    [Parameter(ParameterSetName = "Set 3")]
    [Parameter(ParameterSetName = "Set 4")]
    [Parameter(ParameterSetName = "Set 5")]
    [string]
    $ItemSender,

    #Creates HTML email report and sends it to recipients SMTP address given as parameter value
    [Parameter(Mandatory = $false, 
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromRemainingArguments = $false)]
    [Parameter(ParameterSetName = "Set 2")]
    [Parameter(ParameterSetName = "Set 3")]
    [Parameter(ParameterSetName = "Set 4")]
    [Parameter(ParameterSetName = "Set 5")]
    [string]
    $EmailReportTo

    )

    DynamicParam {
        if (!$ReadOnlyReportMode) {
            #create a new ParameterAttribute Object
            $ConfirmAttribute = New-Object System.Management.Automation.ParameterAttribute
            $ConfirmAttribute.Position = 3
            $ConfirmAttribute.Mandatory = $true
            $ConfirmAttribute.HelpMessage = "Function run with these switches will remove items from choosen mailboxes, to confirm this action please type 'YES'."
 
            #create an attributecollection object for the attribute we just created.
            $attributeCollection = new-object System.Collections.ObjectModel.Collection[System.Attribute]
 
            #add our custom attribute
            $attributeCollection.Add($ConfirmAttribute)
 
            #add our paramater specifying the attribute collection
            $ConfirmParam = New-Object System.Management.Automation.RuntimeDefinedParameter('Confirmation', [String], $attributeCollection)
 
            #expose the name of our parameter
            $paramDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('Confirmation', $ConfirmParam)
            return $paramDictionary
        }
    }

    Process {
    $SearchResultReport = @()
    $SearchRemovalReport = @()
    $Search_Time = Get-Date -Format yyyyMMdd_HHmm
    $TargetFolder = "MBX_Item_removal $Search_Time $env:UserName"
    $TargetMailbox = "spam"

    if ($ItemKind){$QKind = "Kind:$ItemKind"} else {$QKind = $NULL}
    if ($ItemSubject){$QSubject = "Subject:'$ItemSubject'" } else {$QSubject = $NULL}
    if ($ItemRecipient){$QRecipient = "To:'$ItemRecipient'"} else {$QRecipient = $NULL}
    if ($ItemSender){$QSender = "From:'$ItemSender'"} else {$QSender = $NULL}
    if ($ItemTime){$QTime = "Received:'$ItemTime'"} else {$QTime = $NULL}

    if($SourceIsSingleExchMailbox) { $Mailboxes = @(Get-Mailbox "$SourceIsSingleExchMailbox") }
    if($SourceIsExchServer){ $Mailboxes = @(Get-Mailbox -server "$SourceIsExchServer" -resultsize unlimited -IgnoreDefaultScope) } 
    if($SourceIsExchDatabase){ $Mailboxes = @(Get-Mailbox -database "$SourceIsExchDatabase" -resultsize unlimited -IgnoreDefaultScope) }
    if($SourceIsTxtFile) {	$Mailboxes = @(Get-Content "$SourceIsTxtFile" -ReadCount 1 | Get-Mailbox -resultsize unlimited) }
    $mailboxcount = $mailboxes.count

        if (!$ItemKind -and !$ItemSubject -and !$ItemRecipient -and !$ItemSender){
        Write-Output "Please provide at least one criteria for searching the item"
        } else {

            $i = 0
            ForEach ($Mailbox in $Mailboxes){ 
            $i = $i + 1
	        $pct = $i/$mailboxcount * 100

	        Write-Progress -Activity "Collecting mailbox details" -Status "Processing mailbox $i of $mailboxcount - $Mailbox" -PercentComplete $pct -Id 1
               
                $InfoBody = "<p>Script invoked by <b>$env:UserName</b> on <b>$env:ComputerName</b></p>"
                $InfoBody += "<p>Mailbox source: <b>$SourceIsSingleExchMailbox $SourceIsExchServer $SourceIsExchDatabase $SourceIsTxtFile</b></p>"
                $InfoBody += "<p>Item kind: <b>$QKind</b></p>"
                $InfoBody += "<p>Item subject: <b>$QSubject</b></p>"
                $InfoBody += "<p>Item recipient: <b>$QRecipient</b></p>"
                $InfoBody += "<p>Item time: <b>$QTime</b></p>"

                if ($ReadOnlyReportMode){
                $Action = "Report only search"
                
                $SingleSearch = (Search-mailbox $Mailbox.primarysmtpaddress -SearchQuery "$QKind $QSubject $QRecipient $QSender $QTime" -EstimateResultOnly)
                $SearchResult = New-Object PSObject
	            $SearchResult | Add-Member NoteProperty -Name "MailboxName" -Value $Mailbox.Name
	            $SearchResult | Add-Member NoteProperty -Name "NumberOfMessages" -Value $SingleSearch.ResultItemsCount
	            $SearchResult | Add-Member NoteProperty -Name "ItemsSize" -Value $SingleSearch.ResultItemsSize
                $SearchResultReport += $SearchResult

                } elseif (!$ReadOnlyReportMode -and $PSBoundParameters.Confirmation -eq "YES") {
                $Action = "Deletion search"

                $InfoBody += "<p>Destination mailbox for deleted emelements: <b>'$TargetMailbox'</b> folder: <b>'$QTime$TargetFolder'</b></p>"
                
                $SingleSearch = (Search-mailbox $Mailbox.primarysmtpaddress -SearchQuery "$QKind $QSubject $QRecipient $QSender $QTime" -TargetFolder "$TargetFolder" -TargetMailbox $TargetMailbox -DeleteContent -confirm:$false -force) 
               
                $RemovalResult = New-Object PSObject
	            $RemovalResult | Add-Member NoteProperty -Name "MailboxName" -Value $Mailbox.Name
	            $RemovalResult | Add-Member NoteProperty -Name "NumberOfMessages" -Value $SingleSearch.ResultItemsCount
	            $RemovalResult | Add-Member NoteProperty -Name "ItemsSize" -Value $SingleSearch.ResultItemsSize
                $SearchRemovalReport += $RemovalResult

                } else {
                break;
                }
                
           

}
}


$a = "<style>"
$a = $a + "h1, h5, th { text-align: center;font-family: Segoe UI; } "
$a = $a + "table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; } "
$a = $a + "th { background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; } "
$a = $a + "td { font-size: 11px; padding: 5px 20px; color: #000; } "
$a = $a + "tr { background: #b8d1f3; }  "
$a = $a + "tr:nth-child(even) { background: #dae5f4; }  "
$a = $a + "tr:nth-child(odd) { background: #b8d1f3; } "
$a = $a + "p {font-family: Arial, Helvetica, sans-serif;}"
$a = $a + "</style>"


if ($SearchResult){
$SearchResultReport | Export-Csv -Path MailboxSearchReport.csv -NoTypeInformation
$EmailBody = $SearchResultReport | ConvertTo-Html -head $a -body "<H1>Mailbox Search - Report Mode</H1>$InfoBody"
}

if ($RemovalResult) {
$SearchRemovalReport | Export-Csv -Path MailboxRemovalReport.csv -NoTypeInformation
$EmailBody = $SearchRemovalReport | ConvertTo-Html -head $a -body "<H1>Mailbox Search - Deletion Mode</H1>$InfoBody"
}

if ($EmailReportTo -and $Action) {

$EmailSubject = "[REPORT] $Action process has been triggered on mailboxes"
$EmailRcpt = $EmailReportTo
$EmailFrom = "EmailItemRemoval@domain.com"
$EmailServer = "mailserver.domain.com"
Send-MailMessage -From "$EmailFrom" -To "$EmailRcpt" -Subject "$EmailSubject" -Body "$EmailBody" -SmtpServer "$EmailServer" -BodyAsHtml
}


}

}