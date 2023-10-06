<#
Incident States:
New         1
In Progress 2
On Hold     3
Resolved    6
Closed      7
Canceled    8

SCTASK States:
Pending           -5
Open              1
Work in Progress  2
Closed Complete   3
Closed Incomplete 4
Closed Skipped    7
#>

$Global:ServiceNow_Server = "https://*****.service-now.com"

function Add-ServiceNowAttachment{
param(
[Parameter(Mandatory)]
[ValidateSet("sc_task","incident")]
$TicketType,
[Parameter(Mandatory)]
$TicketSysID,
$File,
[switch]$SkipVerification
)
    if (!$ServiceNow_Session){Confirm-ServiceNowSession}

        #Upload Attachment to Ticket in ServiceNow
        if($File -and $File -ne ""){
            $FileOb = Get-Item $File
            $SN_Attachment_File = @{
                'SafeFileName' = $FileOb.FullName.substring($FileOb.FullName.LastIndexOf("\")+1)
                'FileName' = $FileOb.FullName
            }
        }else{
            $SN_Attachment_File = Get-File
        }
        $SN_Attachment_FileName = $SN_Attachment_File.SafeFileName
        $SN_Attachment_Table_Name = $TicketType
        $SN_Attachment_Table_Sys_Id = $TicketSysID
        $SN_Attachment_Content_Type = Get-MimeType $SN_Attachment_File.FileName
        $SN_Attachment_Payload_File = $SN_Attachment_File.FileName
        $SN_Attachment_Payload_File_Bin = [IO.File]::ReadAllBytes($SN_Attachment_Payload_File)
        $SN_Attachment_Encoding = [System.Text.Encoding]::GetEncoding("iso-8859-1")
        $SN_Attachment_Payload_File_Encoding = $SN_Attachment_Encoding.GetString($SN_Attachment_Payload_File_Bin)
        $SN_Attachment_GUID = ((New-Guid).Guid | Out-String).Trim()
        $LF = "`r`n"
        $SN_Attachment_Body = (
            "-----------------------------$SN_Attachment_GUID",
            "Content-Disposition: form-data; name=`"sysparm_ck`"",
            "",
            $SN_User_Token,
            "-----------------------------$SN_Attachment_GUID",
            "Content-Disposition: form-data; name=`"attachments_modified`"",
            "",
            "",
            "-----------------------------$SN_Attachment_GUID",
            "Content-Disposition: form-data; name=`"sysparm_sys_id`"",
            "",
            $SN_Attachment_Table_Sys_Id,
            "-----------------------------$SN_Attachment_GUID",
            "Content-Disposition: form-data; name=`"sysparm_table`"",
            "",
            $SN_Attachment_Table_Name,
            "-----------------------------$SN_Attachment_GUID",
            "Content-Disposition: form-data; name=`"max_size`"",
            "",
            "1024",
            "-----------------------------$SN_Attachment_GUID",
            "Content-Disposition: form-data; name=`"file_types`"",
            "",
            "",
            "-----------------------------$SN_Attachment_GUID",
            "Content-Disposition: form-data; name=`"sysparm_nostack`"",
            "",
            "yes",
            "-----------------------------$SN_Attachment_GUID",
            "Content-Disposition: form-data; name=`"sysparm_redirect`"",
            "",
            "attachment_uploaded.do?sysparm_domain_restore=false&sysparm_nostack=yes",
            "-----------------------------$SN_Attachment_GUID",
            "Content-Disposition: form-data; name=`"sysparm_encryption_context`"",
            "",
            "",
            "-----------------------------$SN_Attachment_GUID",
            "Content-Disposition: form-data; name=`"attachFile`"; filename=`"$SN_Attachment_FileName`"",
            "Content-Type: $SN_Attachment_Content_Type",
            "",
            $SN_Attachment_Payload_File_Encoding,
            "-----------------------------$SN_Attachment_GUID--",
            ""
        ) -join $LF

    $global:SN_Submit_Attachment = Invoke-WebRequest -Uri "https://$ServiceNow_Server/sys_attachment.do?sysparm_record_scope=global" -Method "POST" -ContentType "multipart/form-data; boundary=---------------------------$SN_Attachment_GUID" -Body $SN_Attachment_Body -WebSession $ServiceNow_Session
    if ($SN_Submit_Attachment.StatusCode -eq "200") {
        #$INC_ID = $Submit_INC_2.number
        Write-Host "*** Successfully Submitted Attachment `"$SN_Attachment_FileName`" for Ticket $TicketSysID ***" -ForegroundColor Green
    }else{
        Write-Host "File attachment upload failed!`nStatus: $($SN_Submit_Attachment.StatusCode)`n"
    }
}

function Close-ServiceNowIncident{
param(
$SysID
)

    while($True){
        try{
            $body = @{"state" = 8} | ConvertTo-Json -Compress #8 is for cancel
            $Cancel_Incident = Invoke-RestMethod "https://$ServiceNow_Server/incident_list.do?JSONv2&sysparm_sys_id=$SysID&sysparm_action=update" -WebSession $ServiceNow_Session `
            -Method Post -ContentType "application/json" -Body $body -ErrorVariable Modify_Incident_Error
            break
        }catch{
            if($ReAuthTried){
                Write-Host "Failed to submit ticket after verifying session, skipping!" -ForegroundColor Red
                $ReAuthTried = $False
                return
            }
            Write-Host "Error occured while submitting SC Task request to SNOW! Retrying..." -ForegroundColor Yellow
            if($RetryCount -eq 3){
                if(!$ReAuthTried){
                    Write-Host "Failed to submit ticket 3 times in a row...Verifying Session & Authentication status!" -ForegroundColor Red
                    Confirm-ServiceNowSession
                    $ReAuthTried = $True
                }
            }
            $RetryCount += 1
        }
    }
    return $Cancel_Incident.records
}

function Close-ServiceNowSession{
    Get-EventSubscriber -Force | Unregister-Event -Force
    if($ServiceNow_Session_Timer){$ServiceNow_Session_Timer.enabled = $false}
    Remove-Variable -Name "ServiceNow_*", "SN_*" -ErrorAction SilentlyContinue
}
    
function Confirm-ServiceNowSession{
    if($ServiceNow_Session){
        $SN_User_Profile_Page_Refresh = (Invoke-RestMethod -Uri "https://$ServiceNow_Server/sys_user.do?JSONv2&sysparm_action=get&sysparm_sys_id=$SN_UserID" -WebSession $ServiceNow_Session).records
        $SN_DisplayName_Refresh = $SN_User_Profile_Page_Refresh.name

        if($SN_DisplayName -ne $SN_DisplayName_Refresh){
            Write-Host "Service Now session expired! Refreshing..." -ForegroundColor Yellow
            $ServiceNow_Session_Timer.Enabled = $False
            Unregister-Event -SubscriptionId ($ServiceNow_Session_Timer_Event.Id)
            New-ServiceNowSession
        }
    }else{
        Write-Host "Service Now session not found!" -ForegroundColor Red
        New-ServiceNowSession
    }
}

function Get-AuthCertificate {
    $Certificates = [System.Security.Cryptography.X509Certificates.X509Certificate2[]](Get-ChildItem Cert:\CurrentUser\My | where {$_.NotAfter -gt (Get-Date) -and $_.EnhancedKeyUsageList.FriendlyName -match "Smart Card Logon|Client Authentication"}) | select Thumbprint,FriendlyName,@{l="Issuer";e={$_.Issuer.Split(",")[0]}}
    
    $Certificates | Add-Member -MemberType NoteProperty -Name "Index" -Value 0
    $i=0
    foreach($Cert in $Certificates){$Cert.Index=$i;$i++}
    $Certificates = $Certificates | select Index,FriendlyName,Thumbprint,Issuer
    Write-Host "******Smart Card Certificates******`n" -ForegroundColor Yellow
    write-host "$(($Certificates | Out-String).Trim())"

    Write-Host "`nCertificate #: " -NoNewline -ForegroundColor Yellow
    $i = Read-Host

    return $Certificates[$i]
}

function Get-File {
param(
$File=""
)
    if ($File -ne ""){
        $FileOb = Get-Item $File
        $Directory = $FileOb.Directory.FullName
        $FileName = $FileOb.FullName.substring($FileOb.FullName.LastIndexOf("\")+1)
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            InitialDirectory = $Directory
            FileName = $FileName
            Filter = 'All files (*.*)|*.*|Archive|*.7z;*.cab;*.tar;*.gz;*.zip|CSV|*.csv|Excel (*.xls;*.xlsx)|*.xls*|HTML|*.html|Image|*.bmp;*.gif;*.jpg;*.jpeg|JSON|*.json|Outlook|*.msg|PDF|*.pdf|PowerPoint|*.pptx|PS1|*.ps1|TXT|*.rtf;*.txt|Visio|*.vsdx|Word (*.doc;*.docx)|*.doc*|XML|*.xml'
        }
    }else{
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            InitialDirectory = [Environment]::GetFolderPath('Desktop')
            Filter = 'All files (*.*)|*.*|Archive|*.7z;*.cab;*.tar;*.gz;*.zip|CSV|*.csv|Excel (*.xls;*.xlsx)|*.xls*|HTML|*.html|Image|*.bmp;*.gif;*.jpg;*.jpeg|JSON|*.json|Outlook|*.msg|PDF|*.pdf|PowerPoint|*.pptx|PS1|*.ps1|TXT|*.rtf;*.txt|Visio|*.vsdx|Word (*.doc;*.docx)|*.doc*|XML|*.xml'
        }
    }

    $null = $FileBrowser.ShowDialog()
    return $FileBrowser | Select SafeFileName, FileName
}

function Get-MimeType {
    param( 
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)] 
        [ValidateScript({Test-Path $_})]
        [String]$File
        )
    Add-Type -AssemblyName "System.Web"
    [System.Web.MimeMapping]::GetMimeMapping($File)
}

function Get-ServiceNowCategories {
    $global:ServiceNowCATsFilePath = "$($PSScriptRoot)\ServiceNow_Categories.json"

    if(Test-Path $ServiceNowCATsFilePath){
        $global:ServiceNow_Categories = (Get-Content $ServiceNowCATsFilePath -Raw) | ConvertFrom-Json
        Write-Host "ServiceNow Categories JSON file import successful!" -ForegroundColor Green
    }else{
        Write-Host "ServiceNow Categories JSON file not found!" -ForegroundColor Red
        Write-Host "Download latest ServiceNow Categories JSON file?(y/n): " -ForegroundColor Yellow -NoNewline
        $confirm = Read-Host

        if($confirm.ToLower() -eq "y" -or $confirm.ToLower() -eq "yes"){
            Update-ServiceNowCategories
            $global:ServiceNow_Categories = (Get-Content $ServiceNowCATsFilePath -Raw) | ConvertFrom-Json
            Write-Host "Service Now Categories hash table created successfully!" -ForegroundColor Green
        }else{
            return $null
        }
    }
}

function Get-ServiceNowGroups {
    $global:ServiceNowGroupsFilePath = "$($PSScriptRoot)\ServiceNow_Groups.json"

    if(Test-Path $ServiceNowGroupsFilePath){
        $global:ServiceNow_Groups = (Get-Content $ServiceNowGroupsFilePath -Raw) | ConvertFrom-Json
        Write-Host "ServiceNow Groups JSON file import successful!" -ForegroundColor Green
    }else{
        Write-Host "ServiceNow Groups JSON file not found!" -ForegroundColor Red
        Write-Host "Download latest ServiceNow Groups JSON file?(y/n): " -ForegroundColor Yellow -NoNewline
        $confirm = Read-Host

        if($confirm.ToLower() -eq "y" -or $confirm.ToLower() -eq "yes"){
            Update-ServiceNowGroups
            $global:ServiceNow_Groups = (Get-Content $ServiceNowGroupsFilePath -Raw) | ConvertFrom-Json
            Write-Host "Service Now Groups array created successfully!" -ForegroundColor Green
        }else{
            return $null
        }
    }
}

function Get-ServiceNowRecord{
param(
[Parameter(Mandatory=$true)]
[ValidateSet("ChangeRequest","ChangeTask","Group","Incident","Request","RequestItem","ScheduledTask","User","ConfigurationItem")]
$RecordType,
[Parameter(Mandatory=$false)]
$SysID,
[Parameter(Mandatory=$false)]
$FirstName,
[Parameter(Mandatory=$false)]
$LastName,
[Parameter(Mandatory=$false)]
$GroupName,
[Parameter(Mandatory=$false)]
$ComputerName,
[Parameter(Mandatory=$false)]
$GroupNameSearch,
[Parameter(Mandatory=$false)]
$TicketNumber,
[Parameter(Mandatory=$false)]
$TicketSearch
)
    if (!$ServiceNow_Session){Confirm-ServiceNowSession}

    switch ($RecordType.toLower()){
        "user" {
            $RecordTypeURL = "sys_user_list.do"
            if ($SysID){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_sys_id=$SysID" -WebSession $ServiceNow_Session).records
            }elseif($FirstName -and $LastName){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_query=last_name=$LastName^first_name=$FirstName" -WebSession $ServiceNow_Session).records
            }elseif($FirstName){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_query=last_name=$FirstName" -WebSession $ServiceNow_Session).records
            }elseif($LastName){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_query=last_name=$LastName" -WebSession $ServiceNow_Session).records
            }else{
                Write-Host "A Sys ID or First/Last name is required to run this command." -ForegroundColor Red
                return
            }
            return
        }
        "group" {
            $RecordTypeURL = "sys_user_group_list.do"
            if ($SysID){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_sys_id=$SysID" -WebSession $ServiceNow_Session).records
            }elseif($GroupName){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_query=name=$GroupName" -WebSession $ServiceNow_Session).records
            }elseif($GroupNameSearch){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_query=nameLIKE$GroupNameSearch" -WebSession $ServiceNow_Session).records | select assignment_group,closed_at,closed_by,description,impact,number,opened_by,parent,priority,short_description,state,sys_created_on,sys_id,sys_updated_by,sys_updated_on,urgency
            }else{
                Write-Host "A Sys ID or Group name is required to run this command." -ForegroundColor Red
                return
            }
            return $ServiceNowRecord
        }
        "scheduledtask" {
            $RecordTypeURL = "sc_task_list.do"
            if ($SysID){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_sys_id=$SysID" -WebSession $ServiceNow_Session).records
            }elseif($TicketNumber){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_query=number=$TicketNumber" -WebSession $ServiceNow_Session).records
            }elseif($TicketSearch){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_query=short_descriptionLIKE$TicketSearch^ORdescriptionLIKE$TicketSearch" -WebSession $ServiceNow_Session).records
            }else{
                Write-Host "A SysID or Ticket Number is required to run this command." -ForegroundColor Red
                return
            }
            return $ServiceNowRecord
        }
        "changerequest" {
            $RecordTypeURL = "change_request_list.do"
            if ($SysID){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_sys_id=$SysID" -WebSession $ServiceNow_Session).records
            }elseif($TicketNumber){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_query=number=$TicketNumber" -WebSession $ServiceNow_Session).records
            }else{
                Write-Host "A SysID or Ticket Number is required to run this command." -ForegroundColor Red
                return
            }
            return $ServiceNowRecord
        }
        "changetask" {
            $RecordTypeURL = "change_task_list.do"
            if ($SysID){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_sys_id=$SysID" -WebSession $ServiceNow_Session).records
            }elseif($TicketNumber){
                #https://$ServiceNow_Server/sc_task_list.do?sysparm_query=number%3DSCTASK0345015&sysparm_first_row=1&sysparm_view=
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_query=number=$TicketNumber" -WebSession $ServiceNow_Session).records
            }else{
                Write-Host "A SysID or Ticket Number is required to run this command." -ForegroundColor Red
                return
            }
            return $ServiceNowRecord
        }
        "incident" {
            $RecordTypeURL = "incident_list.do"
            if ($SysID){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_sys_id=$SysID" -WebSession $ServiceNow_Session).records
            }elseif($TicketNumber){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_query=number=$TicketNumber" -WebSession $ServiceNow_Session).records
            }else{
                Write-Host "A SysID or Ticket Number is required to run this command." -ForegroundColor Red
                return
            }
            return $ServiceNowRecord
        }
        "request" {
            $RecordTypeURL = "sc_request_list.do"
            if ($SysID){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_sys_id=$SysID" -WebSession $ServiceNow_Session).records
            }elseif($TicketNumber){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_query=number=$TicketNumber" -WebSession $ServiceNow_Session).records
            }else{
                Write-Host "A SysID or Ticket Number is required to run this command." -ForegroundColor Red
                return
            }
            return $ServiceNowRecord
        }
        "requestitem" {
            $RecordTypeURL = "sc_req_item_list.do"
            if ($SysID){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_sys_id=$SysID" -WebSession $ServiceNow_Session).records
            }elseif($TicketNumber){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_query=number=$TicketNumber" -WebSession $ServiceNow_Session).records
            }else{
                Write-Host "A SysID or Ticket Number is required to run this command." -ForegroundColor Red
                return
            }
            return $ServiceNowRecord
        }
        "configurationitem" {
            $RecordTypeURL = "cmdb_ci_pc_hardware.do"
            if ($SysID){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_sys_id=$SysID" -WebSession $ServiceNow_Session).records
            }elseif($ComputerName){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_query=name=$ComputerName" -WebSession $ServiceNow_Session).records
            }else{
                Write-Host "A SysID or Computer Name is required to run this command." -ForegroundColor Red
                return
            }
            return $ServiceNowRecord
        }
        Default {
            $RecordTypeURL = "sc_task_list.do"
            if ($SysID){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_sys_id=$SysID" -WebSession $ServiceNow_Session).records
            }elseif($TicketNumber){
                $global:ServiceNowRecord = (Invoke-RestMethod -Method Get -Uri "https://$ServiceNow_Server/$($RecordTypeURL)?JSONv2&sysparm_query=number=$TicketNumber" -WebSession $ServiceNow_Session).records
            }else{
                Write-Host "A SysID or Ticket Number is required to run this command." -ForegroundColor Red
                return
            }
            return $ServiceNowRecord
        }
    }
}

function Get-ServiceNowServices {
    $global:ServiceNowServicesFilePath = "$($PSScriptRoot)\ServiceNow_Services.JSON"

    if(Test-Path $ServiceNowServicesFilePath){
        $global:ServiceNow_Services = (Get-Content $ServiceNowServicesFilePath -Raw) | ConvertFrom-Json
        Write-Host "ServiceNow Services JSON file import successful!" -ForegroundColor Green
    }else{
        Write-Host "ServiceNow Services JSON file not found!" -ForegroundColor Red
        Write-Host "Download latest ServiceNow Services JSON file?(y/n): " -ForegroundColor Yellow -NoNewline
        $confirm = Read-Host

        if($confirm.ToLower() -eq "y" -or $confirm.ToLower() -eq "yes"){
            Update-ServiceNowServices
            $global:ServiceNow_Services = (Get-Content $ServiceNowServicesFilePath -Raw) | ConvertFrom-Json
            Write-Host "Service Now Services array created successfully!" -ForegroundColor Green
        }else{
            return $null
        }
    }
}

function New-ServiceNowIncident{
param(
[Parameter(Mandatory)]
$ShortDescription,
[Parameter(Mandatory)]
$Description,
[Parameter(Mandatory)]
$Category,
[Parameter(Mandatory)]
$Subcategory,
$Service,
[Parameter(Mandatory)]
$Group,
$AssignedTo,
$Impact = "3 - Low",
$Urgency = "3 - Low",
$Parent,
$File="",
[switch]$UploadAttachment,
#$ConfigurationItem,
[switch]$SkipVerification
)

    if(-not$ServiceNow_Groups){Get-ServiceNowGroups}
    if($ServiceNow_Groups.name.Contains($Group)){
        $INC_Group_Name_ID = ($ServiceNow_Groups | where {$_.name -eq $Group}).sys_id
    }else{
        Write-Host "No Group ID found for: $Group`r`nShort Description: $ShortDescription`r`nExiting function!" -ForegroundColor Red
        return $null
    }

    #If Parent exists, get SYSID of Parent Task
    $SN_Ticket_Parent = ""
    if($Parent -ne "" -and $Parent -ne $null){
        try{
            $SN_Ticket_Parent = (Get-ServiceNowRecord -RecordType ScheduledTask -TicketNumber $Parent).sys_id
            if(!($SN_Ticket_Parent)){
                $SN_Ticket_Parent = (Get-ServiceNowRecord -RecordType Incident -TicketNumber $Parent).sys_id
            }
        }catch{
            Write-Host "Error finding SYSID for Parent task! Proceeding with no parent!"
            $SN_Ticket_Parent = ""
        }
    }

    #Create Ticket for ServiceNow
    $SN_Ticket_CallerID = $SN_DisplayName
    $SN_Ticket_Location = $SN_Location_Name
    $SN_Ticket_Category = $Category
    $SN_Ticket_SubCategory = $Subcategory
    $SN_Ticket_Service = $Service
    $SN_Ticket_ContactType = "Self-service"
    $SN_Ticket_Impact = $Impact
    $SN_Ticket_Urgency = $Urgency
    #$SN_Ticket_DueDate = $INC_MitigationDate
    #$SN_Ticket_ScheduledDate = $INC_MitigationDate
    $SN_Ticket_AssignmentGroup = $INC_Group_Name_ID
    $SN_Ticket_AssignedTo = $AssignedTo
    $SN_Ticket_ShortDescription = $ShortDescription
    $SN_Ticket_Description = $Description
    $SN_Ticket_Body = @{
        parent                          = $SN_Ticket_Parent
        caller_id                       = $SN_Ticket_CallerID
        location                        = $SN_Ticket_Location
        category                        = $SN_Ticket_Category
        subcategory                     = $SN_Ticket_SubCategory
        business_service                = $SN_Ticket_Service
        contact_type                    = $SN_Ticket_ContactType
        severity                        = $SN_Ticket_Impact
        urgency                         = $SN_Ticket_Urgency
        #due_date                        = $SN_Ticket_DueDate
        assignment_group                = $SN_Ticket_AssignmentGroup
        assigned_to                     = $SN_Ticket_AssignedTo
        short_description               = $SN_Ticket_ShortDescription
        description                     = $SN_Ticket_Description
    }

    $SN_Ticket_Headers = @{
        'Accept' = "application/json"
        'X-UserToken' = $SN_User_Token
    }

    Write-Host "`n*** Incident Details Overview ***" -ForegroundColor Yellow
    Write-Host "`nCustomer:" -ForegroundColor Cyan
    Write-Host $SN_DisplayName -ForegroundColor White
    Write-Host "`nCustomer Location:" -ForegroundColor Cyan
    Write-Host $SN_Location_Name -ForegroundColor White
    Write-Host "`nShort Description:" -ForegroundColor Cyan
    Write-Host $ShortDescription -ForegroundColor White
    Write-Host "`nDescription:" -ForegroundColor Cyan
    Write-Host $Description -ForegroundColor White
    Write-Host "`nAssignment Group:" -ForegroundColor Cyan
    Write-Host $Group -ForegroundColor White
    Write-Host "ID: $INC_Group_Name_ID" -ForegroundColor White
    Write-Host "`nCategory:" -ForegroundColor Cyan
    Write-Host $SN_Ticket_Category -ForegroundColor White
    Write-Host "`nSubcategory:" -ForegroundColor Cyan
    Write-Host $SN_Ticket_SubCategory -ForegroundColor White
    Write-Host "`nService:" -ForegroundColor Cyan
    Write-Host $SN_Ticket_Service -ForegroundColor White
    Write-Host ""

    if (-not $SkipVerification){
        $confirm = Read-Host "Continue(y/n)"
    }else{
        $confirm = "y"
    }

    if($confirm.ToLower() -eq "y"){
        $RetryCount = 0
        $ReAuthTried = $False
        while($True){
            try{
                $Submit_INC = Invoke-WebRequest -Uri "https://$ServiceNow_Server/incident.do?JSONv2&sysparm_action=insert" -Method "POST" -ContentType "application/json" -Headers $SN_Ticket_Headers -Body $($SN_Ticket_Body | ConvertTo-Json) -WebSession $ServiceNow_Session
                break
            }catch{
                if($ReAuthTried){
                    Write-Host "Failed to submit ticket after verifying session, skipping!" -ForegroundColor Red
                    $ReAuthTried = $False
                    return
                }
                Write-Host "Error occured while submitting ticket request to SNOW! Retrying!"
                if($RetryCount -eq 3){
                    if(!$ReAuthTried){
                        Write-Host "Failed to submit ticket 3 times in a row...Verifying Session!" -ForegroundColor Red
                        Confirm-ServiceNowSession
                        $ReAuthTried = $True
                    }
                }
                $RetryCount += 1
            }
        }
        $Submit_INC_2 = ($Submit_INC.Content | ConvertFrom-JSON).records[0]
        if (($Submit_INC.StatusCode -eq "200")) {
            $Global:INC_Number = $Submit_INC_2.number
            $Global:INC_SysID = $Submit_INC_2.sys_id
            Write-Host "*** Successfully Submitted `"$INC_SysID`" to ServiceNow ***`n" -ForegroundColor Green
            if($UploadAttachment.IsPresent){
                if($File -ne "" -and $File -ne $null){
                    Add-ServiceNowAttachment -TicketType 'incident' -TicketSysID $INC_SysID -File $File
                }else{
                    Add-ServiceNowAttachment -TicketType 'incident' -TicketSysID $INC_SysID
                }
            }
            return "$INC_Number,$INC_SysID"
        }
    }else{
        Write-Host "Aborting Ticket Creation!`n" -ForegroundColor Red
    }
}

function New-ServiceNowSCTask{
param(
[Parameter(Mandatory)]
$ShortDescription,
[Parameter(Mandatory)]
$Description,
[Parameter(Mandatory)]
$Group,
$AssignedTo="",
$Service="",
$File="",
[switch]$UploadAttachment,
[switch]$SkipVerification
)

    if (-not($ServiceNow_Groups)){Get-ServiceNowGroups}
    if($ServiceNow_Groups.name.Contains($Group)){
        $Group_ID = ($ServiceNow_Groups | where {$_.name -eq $Group}).sys_id
    }elseif($Group.length -eq 32){
        $Group_ID = $Group
    }else{
        Write-Host "No Group ID found for: $Group`r`nShort Description: $ShortDescription`r`nExiting function!" -ForegroundColor Red
        return $null
    }

    #Create Ticket for ServiceNow
    $SN_Ticket_CallerID = $SN_DisplayName
    $SN_Ticket_Location = $SN_Location_Name
    $SN_Ticket_Type = "catalog_task"
    $SN_Ticket_ContactType = "Self-service"
    $SN_Ticket_ShortDescription = $ShortDescription
    $SN_Ticket_Description = $Description
    $SN_Ticket_AssignmentGroup = $Group_ID
    $SN_Ticket_AssignedTo = $AssignedTo
    $SN_Ticket_Service = $Service
    #$SN_Ticket_Category = ""
    #$SN_Ticket_SubCategory = ""
    #$SN_Ticket_LOS = ""
    #$SN_Ticket_Impact = $Impact
    #$SN_Ticket_Urgency = $Urgency
    #$SN_Ticket_DueDate = $INC_MitigationDate
    #$SN_Ticket_ScheduledDate = $INC_MitigationDate

    $SN_Ticket_Body = @{
        sysparm_quantity="1"
        sysparm_item_guid=((New-Guid).Guid -replace "-","")
        get_portal_messages="true"
        sysparm_no_validation="true"
        engagement_channel="sp"
        referrer=$null
        variables = @{
            short_description=$SN_Ticket_ShortDescription
            results_that_may_help=""
            assignment_group=$SN_Ticket_AssignmentGroup
            sub_category=""
            cmdb_ci=""
            description=$SN_Ticket_Description
            business_service=$SN_Ticket_Service
            requested_for=$SN_UserID
            sc_task_template=""
            ai_search_results=""
            basic_request_variables="true"
            location=$SN_LocationID
            sp_attachments=""
            incident_template=""
            itil_type_of_ticket=$SN_Ticket_Type
            category=""
            assigned_to=$SN_Ticket_AssignedTo
        }
    }

    $headers = @{
        'Accept' = "application/json"
        'X-UserToken' = $SN_User_Token
    }

    Write-Host "`n*** SCTask Details Overview ***" -ForegroundColor Yellow
    Write-Host "`nCustomer:" -ForegroundColor Cyan
    Write-Host $SN_DisplayName -ForegroundColor White
    Write-Host "`nCustomer Location:" -ForegroundColor Cyan
    Write-Host $SN_Location_Name -ForegroundColor White
    Write-Host "`nShort Description:" -ForegroundColor Cyan
    Write-Host $SN_Ticket_ShortDescription -ForegroundColor White
    Write-Host "`nDescription:" -ForegroundColor Cyan
    Write-Host $Description -ForegroundColor White
    Write-Host "`nAssignment Group:" -ForegroundColor Cyan
    Write-Host $Group -ForegroundColor White
    Write-Host "$SN_Ticket_AssignmentGroup" -ForegroundColor White
    Write-Host "`nAssigned to:" -ForegroundColor Cyan
    Write-Host $SN_Ticket_AssignedTo -ForegroundColor White
    Write-Host "`nService:" -ForegroundColor Cyan
    Write-Host $SN_Ticket_Service -ForegroundColor White
    Write-Host ""

    $Continue = "false"
    while ($Continue -eq "false") {
        if(-not$SkipVerification.IsPresent){
            Write-Host "`nWould You Like to Submit This SCTask (y/n)? " -ForegroundColor Green -NoNewline
            $Submit_Confirm = Read-Host
        }else{
            $Submit_Confirm = "y"
        }
        if ($Submit_Confirm -imatch "^y$") {
            $RetryCount = 0
            $ReAuthTried = $False
            while($True){
                try{
                    $Submit_SCTask = Invoke-WebRequest -Uri "https://$ServiceNow_Server/api/sn_sc/v1/servicecatalog/items/5d167fd96ce4fc1004ed764f2fe89f42/order_now" -Method "POST" -ContentType "application/json" -Body $($SN_Ticket_Body | ConvertTo-Json) -WebSession $ServiceNow_Session -Headers $headers -ErrorVariable Submit_SCTask_Error
                    break
                }catch{
                    if($ReAuthTried){
                        Write-Host "Failed to submit ticket after verifying session, skipping!" -ForegroundColor Red
                        $ReAuthTried = $False
                        return
                    }
                    Write-Host "Error occured while submitting SC Task request to SNOW! Retrying..." -ForegroundColor Yellow
                    if($RetryCount -eq 3){
                        if(!$ReAuthTried){
                            Write-Host "Failed to submit ticket 3 times in a row...Verifying Session & Authentication status!" -ForegroundColor Red
                            Confirm-ServiceNowSession
                            $ReAuthTried = $True
                        }
                    }
                    $RetryCount += 1
                }
            }
            $Submit_SCTask_2 = ($Submit_SCTask.Content | ConvertFrom-JSON).result
            write-host $Submit_SCTask_Error -ForegroundColor Red

            if (($Submit_SCTask.StatusCode -eq "200")) {
                $RequestSysID = $Submit_SCTask_2.sys_id

                #-----Query 'SC_Task' for a record that contains a 'Request' with SysID from above-----
                $SCTaskQuery = Invoke-RestMethod -Uri "https://$ServiceNow_Server/sc_task.do?JSONv2&sysparm_query=request=$RequestSysID" -WebSession $ServiceNow_Session
                $global:SCTask_SysID = $SCTaskQuery.records[0].sys_id
                $global:SCTask_Number = $SCTaskQuery.records[0].number

                Write-Host "*** Successfully Submitted `"$SCTask_Number`" to ServiceNow ***" -ForegroundColor Green
                if($UploadAttachment.IsPresent){
                    if($File -ne "" -and $File -ne $null){
                        Add-ServiceNowAttachment -TicketType 'sc_task' -TicketSysID $SCTask_SysID -File $File
                    }else{
                        Add-ServiceNowAttachment -TicketType 'sc_task' -TicketSysID $SCTask_SysID
                    }
                }

                $continue = "true"
                return $SCTask_Number
            }else{
                Write-Host "*** Failed to submit SCTask to ServiceNow ***`nTrying again!" -ForegroundColor Red
            }
        }else{
            $continue = "true"
            Write-Host "Skipping SCTask creation in Service Now!`nExiting!" -ForegroundColor Red
        }
    }
}

function New-ServiceNowSession{
param(
    $Server,
    $Username,
    $Pass,
    [switch]$CertificateAuth,
    [switch]$DoD
)
    if($Global:ServiceNow_Server -match "\*" -and !$Server){
        Write-Host "No server was provided for ServiceNow connection!" -ForegroundColor Red
        return
    }elseif($Global:ServiceNow_Server -match "\*" -and $Server){
        if($Server -match "http|https"){
            $server = ($Server -replace "(https://|http://)","" -replace "/","")
        }
        $Global:ServiceNow_Server = $Server
    }

    Close-ServiceNowSession

    Write-Host "Connecting to Service Now..." -ForegroundColor Yellow
    try{
        $SN_Login_Page = Invoke-WebRequest -Uri "https://$ServiceNow_Server/" -SessionVariable global:ServiceNow_Session -ErrorAction Stop
        if($SN_Login_Page.StatusCode -ne 200){
            Write-Host "Connection to ServiceNow failed!`nStatus Code: $($SN_Login_Page.StatusCode)"
            return
        }
    }catch{
        Write-Host "Connection to ServiceNow failed!`nError: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    if ($SN_Login_Page.Content -match "var g_ck = '(.*)'") {$SN_GCK_Token = $matches[1];write-host "Found G_CK Token: $SN_GCK_Token" -ForegroundColor Green}
    
    try{
        if($Username -and $Pass){
            $SN_Banner_Page = Invoke-WebRequest -Uri "https://$ServiceNow_Server/login.do" -Method "POST" -ContentType "application/x-www-form-urlencoded" -Body @{
                "sysparm_ck" = $SN_GCK_Token
                "user_name" = $Username
                "user_password" = $Pass
                "not_important"=$null
                "ni.nolog.user_password" = $true
                "ni.noecho.user_name" = $true
                "ni.noecho.user_password" = $true
                "sys_action" = "sysverb_login"
                "sysparm_login_url" = "welcome.do"} -WebSession $ServiceNow_Session
        }elseif($DoD -and -not($Username)){
            $global:SN_Cert = Get-AuthCertificate
            $SN_Banner_Page = Invoke-WebRequest -Uri "https://$ServiceNow_Server/my.policy" -Certificate $SN_Cert -Method "POST" -ContentType "application/x-www-form-urlencoded" -Body "choice=1" -WebSession $ServiceNow_Session
        }else{
            $global:SN_Cert = Get-AuthCertificate
            $SN_Banner_Page = Invoke-WebRequest -Uri "https://$ServiceNow_Server/login.do" -Certificate $SN_Cert -Method "POST" -ContentType "application/x-www-form-urlencoded" -WebSession $ServiceNow_Session
        }
        if($SN_Banner_Page.StatusCode -ne 200){
            Write-Host "Authentication to ServiceNow failed!`nStatus Code: $($SN_Banner_Page.StatusCode)"
            return
        }
    }catch{
        Write-Host "Authentication to ServiceNow failed!`nError: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    #Retrieve and Set Current User Settings
    if ($SN_Banner_Page.Content -match "window.NOW.user.userID = '(.*?)'") {$global:SN_UserID = $matches[1];write-host "Found User ID: $SN_UserID" -ForegroundColor Green}
    if ($SN_Banner_Page.Content -match "var g_ck = '(.*)'") {$global:SN_User_Token = $matches[1];write-host "Found User Token: $SN_User_Token" -ForegroundColor Green}

    $global:SN_User_Profile_Page = (Invoke-RestMethod -Uri "https://$ServiceNow_Server/sys_user.do?JSONv2&sysparm_action=get&sysparm_sys_id=$SN_UserID" -WebSession $ServiceNow_Session -Headers @{"X-UserToken"=$SN_User_Token} -ErrorAction Stop).records

    $global:SN_DisplayName = $SN_User_Profile_Page.name
    $global:SN_UserName = $SN_User_Profile_Page.user_name
    $global:SN_LocationID = $SN_User_Profile_Page.location
    #$global:SN_Location_Name = ((Invoke-WebRequest -Uri "https://$ServiceNow_Server/cmn_location.do?JSONv2&sysparm_action=get&sysparm_sys_id=$SN_LocationID" -WebSession $ServiceNow_Session).Content | ConvertFrom-JSON).records.name
    $global:SN_Location_Name = (Invoke-RestMethod -UseBasicParsing -Uri "https://$ServiceNow_Server/xmlhttp.do" -Method "POST" -WebSession $ServiceNow_Session -ContentType "application/x-www-form-urlencoded; charset=UTF-8" -Body "sysparm_processor=AjaxClientHelper&sysparm_scope=global&sysparm_want_session_messages=true&sysparm_name=getDisplay&sysparm_table=cmn_location&sysparm_value=$SN_LocationID&sysparm_synch=true&ni.nolog.x_referer=ignore").xml.answer

    Write-Host "Display Name: $SN_DisplayName`nUsername: $SN_UserName`nLocation: $SN_Location_Name`n" -ForegroundColor Green

    Write-Host "Connected to Service Now!`n" -ForegroundColor Green
    $ServiceNow_Session_Expires = ($ServiceNow_Session.Cookies.GetCookies("https://$ServiceNow_Server") | where {$_.Name -eq "glide_session_store"}).Expires
    $global:ServiceNow_Session_Expires_Minutes = [math]::Floor((New-TimeSpan -Start (Get-Date) -End $ServiceNow_Session_Expires).TotalMinutes)
    write-host "Session Expiry: $ServiceNow_Session_Expires_Minutes minutes"
    New-SNSessionRefresher
}

function New-SNSessionRefresher{
    $global:ServiceNow_Session_Timer = New-Object System.Timers.Timer

    $Action = {
        $ServiceNow_Session_Expires = ($ServiceNow_Session.Cookies.GetCookies("https://$ServiceNow_Server") | where {$_.Name -eq "glide_session_store"}).Expires
        $global:ServiceNow_Session_Expires_Minutes = [math]::Floor((New-TimeSpan -Start (Get-Date) -End $ServiceNow_Session_Expires).TotalMinutes)

        $SN_User_Profile_Page_Refresh = (Invoke-RestMethod -Uri "https://$ServiceNow_Server/sys_user.do?JSONv2&sysparm_action=get&sysparm_sys_id=$SN_UserID" -WebSession $ServiceNow_Session).records
        $SN_DisplayName_Refresh = $SN_User_Profile_Page_Refresh.name

        if($SN_DisplayName -ne $SN_DisplayName_Refresh){
            Write-Host "Service Now session expired! Refreshing..." -ForegroundColor Yellow
            $ServiceNow_Session_Timer.Enabled = $False
            Unregister-Event -SubscriptionId ($ServiceNow_Session_Timer_Event.Id)
            New-ServiceNowSession
        }
    }

    $global:ServiceNow_Session_Timer_Event = Register-ObjectEvent -InputObject $ServiceNow_Session_Timer -EventName Elapsed -Action $Action
    $ServiceNow_Session_Timer.Interval = 570000
    $ServiceNow_Session_Timer.AutoReset = $True
    $ServiceNow_Session_Timer.Enabled = $True
}

function Search-Customer{
param($Name)
    $wr = Invoke-RestMethod -UseBasicParsing -Uri "https://$ServiceNow_Server/xmlhttp.do" `
    -Method "POST" `
    -WebSession $ServiceNow_Session `
    -Headers @{
      "X-UserToken"=$SN_User_Token
    } `
    -ContentType "application/x-www-form-urlencoded; charset=UTF-8" `
    -Body "sysparm_processor=Reference&sysparm_scope=global&sysparm_want_session_messages=true&ni.nolog.x_referer=ignore&sysparm_name=incident.caller_id&sysparm_max=15&sysparm_chars=$Name&sysparm_value=&ac_columns=user_name;u_district;email&ac_order_by=name"
}

function Update-ServiceNowCategories {
    Confirm-ServiceNowSession

    $global:ServiceNowCATsFilePath = "$($PSScriptRoot)\ServiceNow_Categories.json"
    $i = 0
    $CategoryListHash = [ordered]@{}
    $CategoryList = (Invoke-RestMethod -UseBasicParsing -Uri "https://$ServiceNow_Server/xmlhttp.do" `
    -Method "POST" `
    -WebSession $ServiceNow_Session `
    -Headers @{
        "X-UserToken"=$SN_User_Token
    } `
    -ContentType "application/x-www-form-urlencoded; charset=UTF-8" `
    -Body "sysparm_processor=PickList&sysparm_scope=global&sysparm_want_session_messages=true&sysparm_name=incident.category&sysparm_chars=*&sysparm_nomax=true&ni.nolog.x_referer=ignore&x_referer=incident.do").xml[1].ChildNodes.Name

    foreach($Cat in $CategoryList){
        if($i%5 -eq 0){
            Write-Host "$([int](($i/88)*100))%.." -NoNewline
        }
        $wr = Invoke-RestMethod -UseBasicParsing -Uri "https://$ServiceNow_Server/xmlhttp.do" `
        -Method "POST" `
        -WebSession $ServiceNow_Session `
        -Headers @{
            "X-UserToken"=$SN_User_Token
        } `
        -ContentType "application/x-www-form-urlencoded; charset=UTF-8" `
        -Body "sysparm_processor=PickList&sysparm_scope=global&sysparm_want_session_messages=true&sysparm_value=$cat&sysparm_name=incident.subcategory&sysparm_chars=*&sysparm_nomax=true&ni.nolog.x_referer=ignore&x_referer=incident.do"

        $CategoryListHash[$Cat] = $wr.xml[1].ChildNodes.name
        $i++
    }
    Write-Host "100%..Download Complete! Saving to file..."
    $CategoryListHash.GetEnumerator() | ConvertTo-Json | Out-File $ServiceNowCATsFilePath -Verbose
    Write-Host "`nService Now Categories JSON file updated successfully!" -ForegroundColor Green
}

function Update-ServiceNowGroups {
    Confirm-ServiceNowSession

    $global:ServiceNowGroupsFilePath = "$($PSScriptRoot)\ServiceNow_Groups.json"
    (Invoke-RestMethod -Uri "https://$ServiceNow_Server/sys_user_group_list.do?JSONv2" -WebSession $ServiceNow_Session).records | where {$_.name -ne "" -and $_.name -ne $null} | select name,sys_id,manager | sort name | ConvertTo-Json | Out-File $ServiceNowGroupsFilePath -Force
    Write-Host "Service Now Groups JSON file updated successfully!" -ForegroundColor Green
}

function Update-ServiceNowServices {
    Confirm-ServiceNowSession

    $global:ServiceNowServicesFilePath = "$($PSScriptRoot)\ServiceNow_Groups.json"
    $ServiceNow_Services = (Invoke-RestMethod -UseBasicParsing -Uri "https://$ServiceNow_Server/cmdb_ci_service_list.do?JSONv2&sysparm_target=incident.business_service" -WebSession $ServiceNow_Session -Headers @{"X-UserToken"=$SN_User_Token}).records | where {$_.name -ne "" -and $_.name -ne $null} | select name,sys_id | sort name | ConvertTo-Json | Out-File $ServiceNowServicesFilePath -Force
    Write-Host "Service Now Services JSON file updated successfully!" -ForegroundColor Green
}



Export-ModuleMember -Function Add-ServiceNowAttachment
Export-ModuleMember -Function Close-ServiceNowIncident
Export-ModuleMember -Function Close-ServiceNowSession
Export-ModuleMember -Function Confirm-ServiceNowSession
Export-ModuleMember -Function Get-ServiceNowCategories
Export-ModuleMember -Function Get-ServiceNowGroups
Export-ModuleMember -Function Get-ServiceNowRecord
Export-ModuleMember -Function Get-ServiceNowServices
Export-ModuleMember -Function New-ServiceNowIncident
Export-ModuleMember -Function New-ServiceNowSCTask
Export-ModuleMember -Function New-ServiceNowSession
Export-ModuleMember -Function Update-ServiceNowCategories
Export-ModuleMember -Function Update-ServiceNowGroups
Export-ModuleMember -Function Update-ServiceNowServices


<#

Get Incidents - Max Record Count 5
$incidents = irm -Uri "https://$ServiceNow_Server/incident_list.do?JSONv2&sysparm_record_count=5" -WebSession $ServiceNow_Session

Get Incident Categories
$incidentCategories = irm -Uri "https://$ServiceNow_Server/sys_choice.do?JSONv2&sysparm_query=name=incident^element=category^ORDERBYname" -WebSession $ServiceNow_Session
$incidentCategories.records | select label, sequence, sys_id

Get Incident Subcategories
$incidentSubcategories = irm -Uri "https://$ServiceNow_Server/sys_choice.do?JSONv2&sysparm_query=name=incident^element=subcategory^ORDERBYname" -WebSession $ServiceNow_Session
$incidentSubcategories.records | select label, dependent_value, sys_id

Get Incident States
$incidentStates = irm -Uri "https://$ServiceNow_Server/sys_choice.do?JSONv2&sysparm_query=name=incident^element=state^ORDERBYname" -WebSession $ServiceNow_Session
$incidentStates.records | select label, sequence, sys_id

Get Incident Channels
$incidentChannels = irm -Uri "https://$ServiceNow_Server/sys_choice.do?JSONv2&sysparm_query=name=incident^element=contact_type^ORDERBYname" -WebSession $ServiceNow_Session
$incidentChannels.records | select label, sequence, sys_id

Get Incident Impact
$incidentImpacts = irm -Uri "https://$ServiceNow_Server/sys_choice.do?JSONv2&sysparm_query=name=incident^element=severity^ORDERBYname" -WebSession $ServiceNow_Session
$incidentImpacts.records | select label, sequence, sys_id

Get Incident Urgency
$incidentUrgencies = irm -Uri "https://$ServiceNow_Server/sys_choice.do?JSONv2&sysparm_query=name=incident^element=severity^ORDERBYname" -WebSession $ServiceNow_Session
$incidentUrgencies.records | select label, sequence, sys_id






Get Services
$incidentServices = irm -Uri "https://$ServiceNow_Server/cmdb_ci_service_list.do?JSONv2" -WebSession $ServiceNow_Session
$incidentServices.records

Get Service Offerings
$incidentServiceOfferings = irm -Uri "https://$ServiceNow_Server/service_offering_list.do?JSONv2" -WebSession $ServiceNow_Session
$incidentServicesOfferings.records

Get Configuration Items
$incidentConfigItems = irm -Uri "https://$ServiceNow_Server/cmdb_ci_list.do?JSONv2" -WebSession $ServiceNow_Session
$incidentConfigItems.records

Get Assignment Groups
$incidentGroups = irm -Uri "https://$ServiceNow_Server/sys_user_group_list.do?JSONv2" -WebSession $ServiceNow_Session
$incidentGroups.records

Get Users
$incidentUsers = irm -Uri "https://$ServiceNow_Server/sys_user_list.do?JSONv2" -WebSession $ServiceNow_Session
$incidentUsers.records

#>
