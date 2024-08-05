#region Functions
Function Connect-ExOnline {

    [CmdletBinding()]
   
    Param(
        # Setting up session requirements
	#Uncomment below two line if you want to use a credetnial key and modify the cdm: Connect-ExchangeOnline accordingly. 
        #$authkey = (Get-content "creds.dat" | ConvertTo-SecureString),
        #$Creds = (New-Object System.Management.Automation.PSCredential("testuser@domain.com",$authkey)),
        )
    Try {
        "Attempting Connection" | Out-File testrun.log -Append
	Connect-ExchangeOnline 
     
}
    Catch {
            Write-Output $_
            "Connection Failed" | Out-File testrun.log -Append
            return "Connection to Exchange Online couldn't be established."
          }
}


##Below Get-MailboxCreatedYesterday fetches a list of the Mailboxes which were created in last 24 hours.
function Get-MailboxCreatedYesterday {
    param (
            $Today = (Get-Date).Date ,
            $Yesterday = $Today.AddDays(-1) ,
            $FileDate = (Get-Date $Yesterday -Format 'yyyy-MM-dd'),
            $FileName = "$FileDate-Mailboxes_Created.csv",
            $MailboxListFile = (Join-Path ".\MailboxList\"  $FileName)
        )
   
    Get-Mailbox -Filter "WhenMailboxCreated -gt '$Yesterday' -and WhenMailboxCreated -lt '$Today'" -RecipientTypeDetails UserMailbox | Sort WhenMailboxCreated |
        Select-Object DisplayName, PrimarySMTPAddress, RecipientTypeDetails, WhenMailboxcreated, RetentionPolicy, LitigationHoldEnabled |
            Export-Csv $MailboxListFile -NoTypeInformation
    return $MailboxListFile
}


##Below function is to apply the custom\new retention Policy to the User Mailboxes. Provide the new\custom retention policy name in the variable: $Policy.
function Update-MRMPolicy {
    param (
        $MailboxList,
        $Policy = "NewPolicyName",
        $Mailboxes = (Import-Csv $MailboxList),
        $LogFileName = ((Get-Date -Format 'yyyy-MM-dd') + "-policy-update.log"),
        $LogFile = (Join-Path ".\Logs\" $LogFileName),
        $ErrorMailboxes = @(),
        $ErrorReportRecipients = ("testuser@domain.com")#, "testuser1@domain.com", "testuser2@domain.com")
       )
   
    ForEach ($Mailbox in $Mailboxes) {
        if ($Mailbox.InScope -eq "Yes") {
            try {
                  Set-Mailbox -Identity $Mailbox.PrimarySmtpAddress -RetentionPolicy $Policy -ErrorAction STOP
                  "Retention Policy updated to $Policy for $($Mailbox.PrimarySmtpAddress)" | Out-File -FilePath $LogFile -Append
                 
                }
            catch {
                    "Error Encountered while applying MRM policy to $($Mailbox.PrimarySmtpAddress)" | Out-File -FilePath $LogFile -Append
                    $Error[0].Exception.Message | Out-File -FilePath $LogFile -Append
                    $ErrorMailboxes += $Mailbox.PrimarySmtpAddress
                   }
            }
       }
   
         
       If ($ErrorMailboxes.length -gt 0) {
        Send-MailMessage -From "RetentionPolicy@domain.com" -To $ErrorReportRecipients -SmtpServer smtp.office365.com -Subject "Error encountered during applying MRM policy" -Attachments $LogFile -Body "$ErrorMailboxes"

        }
}

function Get-MailboxWithDefaultPolicy {
    param (
            $FileDate = (Get-Date -Format 'yyyy-MM-dd'),
            $FileName = "$FileDate-Consolidated-MailboxList.csv",
            $MailboxListFile = (Join-Path ".\MailboxList\"  $FileName),
            $Policy= "Default MRM Policy" )
   
    Get-Mailbox -Filter {RetentionPolicy -eq $Policy} -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Sort WhenMailboxCreated |Select-Object DisplayName, PrimarySMTPAddress, RecipientTypeDetails, WhenMailboxcreated, RetentionPolicy |Export-Csv $MailboxListFile -NoTypeInformation
    return $MailboxListFile
}
function Validate-PostScript {
    param (
        $MailboxList
    )
    $Mailboxes = Import-Csv $MailboxList    
    foreach ( $mailbox in $Mailboxes) {
        If($mailbox.InScope -eq "Yes") {
            $UpdatedRetentionPolicy = Get-Mailbox $mailbox.PrimarySMTPAddress | Select-Object -ExpandProperty RetentionPolicy
            $mailbox | Add-Member -MemberType NoteProperty -Name "Updated_RetentionPolicy" -Value $UpdatedRetentionPolicy -Force
        }
    }
    $Mailboxes | Export-Csv $MailboxList -NoTypeInformation
}

function Format-AlternateRows {
    param(
        # Table as input
        [Parameter(mandatory = $true,
            ValueFromPipeline = $true)]
        [System.Object]
        $InputTable
    )
    $head = @"
    <style>
    table, th, td {border-width: 1px; border-style: solid; border-color: black;}
    th{background-color:LightBlue}
    tr{background-color:Yellow}
    .odd{background-color:#ffffff;}
    .even{background-color:#dddddd;}
    </style>
"@
    $class = "odd"
    $HtmlTable = $InputTable | ConvertTo-Html -Head $head
    foreach ($Line in $HtmlTable) {
        If ($Line -like "<tr><td>*") {
            $NewTable += $Line.replace("<tr>", "<tr class = $class>")
            If ($class -eq "odd") {
                $class = "even"
            }
            else {
                $class = "odd"
            }
        }
        else {
            $NewTable += $Line
        }
    }
    Return $NewTable
}

function Send-Report {
    param (
        $MailboxList
    )
    $Mailboxes = Import-Csv $MailboxList
   
    $Report = @()
    $ReportData = [ordered]@{
        'Description' = "Mailboxes created yesterday with RecipienttypeDetails as UserMailbox"
        'Count'       = ($Mailboxes.RecipientTypeDetails | Where-Object {$_ -eq "UserMailbox"}).count
    }
    $ReportDataobj = New-Object -TypeName PSCustomObject -Property $ReportData
    $Report += $ReportDataobj

    $ReportData = [ordered]@{
        'Description' = "Retention policy updated on mailboxes"
        'Count'       = ($Mailboxes.Updated_RetentionPolicy | Where-Object {$_ -eq "NEW POLICY NAME"}).count
    }
    $ReportDataobj = New-Object -TypeName PSCustomObject -Property $ReportData
    $Report += $ReportDataobj

         
    $ReportHTML = Format-AlternateRows -InputTable $Report
    Send-MailMessage -To testuser@domain.com -Subject "Daily Report for Retention Policy" -From "RetentionPolicy@domain.com" `
        -SmtpServer smtp.office365.com -BodyAsHtml "$ReportHTML"
}

#endregion

       
If(-not (Test-Path .\Logs)) {
    New-Item -ItemType directory -Name "Logs" | Out-Null
    }
If(-not (Test-Path .\MailboxList)) {
    New-Item -ItemType directory -Name "MailboxList" | Out-Null
    }

"Starting Execution" | Out-File testrun.log -Append
Connect-ExOnline
"Now going for pulling data from ExchangeOnline" | Out-File testrun.log -Append
$MBXList = Get-MailboxCreatedYesterday
Update-MRMPolicy -MailboxList $MBXList
Validate-PostScript -MailboxList $MBXList
Send-Report  -MailboxList $MBXList
Get-PSSession | Remove-PSSession
