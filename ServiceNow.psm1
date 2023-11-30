Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
$Global:ServiceNow_Server = "https://*****.service-now.com"
$Global:ServiceNow_Lists = @{}

$Global:ServiceNow_REST_Status_Codes = @{
    200 = @{"Message" = "Success";"Details" = "Success with response body."}
    201 = @{"Message" = "Created";"Details" = "Success with response body."}
    204 = @{"Message" = "Success";"Details" = "Success with no response body."}
    400 = @{"Message" = "Bad Request";"Details" = "The request URI does not match the APIs in the system, or the operation failed for unknown reasons. Invalid headers can also cause this error."}
    401 = @{"Message" = "Unauthorized";"Details" = "The user is not authorized to use the API."}
    403 = @{"Message" = "Forbidden";"Details" = "The requested operation is not permitted for the user. This error can also be caused by ACL failures, or business rule or data policy constraints."}
    404 = @{"Message" = "Not Found";"Details" = "The requested resource was not found. This can be caused by an ACL constraint or if the resource does not exist."}
    405 = @{"Message" = "Method not allowed";"Details" = "The HTTP action is not allowed for the requested REST API, or it is not supported by any API."}
    406 = @{"Message" = "Not acceptable";"Details" = "The endpoint does not support the response format specified in the request Accept header."}
    415 = @{"Message" = "Unsupported media type";"Details" = "The endpoint does not support the format of the request body."}
}

function Parse-String ($String,$StartStr,$EndStr){
    if($String.IndexOf($StartStr) -eq -1){return ""}
    $StartStrPos = $String.IndexOf($StartStr)+$StartStr.Length

    if($String.IndexOf($EndStr,$StartStrPos) -eq -1){return ""}
    $EndStrPos = $String.IndexOf($EndStr,$StartStrPos)

    $ParsedStr = $String.Substring($StartStrPos,$EndStrPos-$StartStrPos)
    if($ParsedStr -ne "" -and $ParsedStr -ne $null){return $ParsedStr}

    return ""
}

function Add-ServiceNowAttachment{
<#
.SYNOPSIS
Uploads an attachment to a ServiceNow ticket of the specified type.

.DESCRIPTION
This function allows you to add an attachment to a specified ServiceNow ticket, such as an incident or service task (sc_task).

.PARAMETER TicketType
Specifies the type of the ServiceNow ticket. Use "incident" for incidents and "sc_task" for service tasks.

.PARAMETER TicketNumber
Specifies the unique ticket number (INC******* or SCTASK*******) of the ServiceNow ticket to which you want to attach a file.

.PARAMETER TicketSysID
Specifies the unique system ID (SysID) of the ServiceNow ticket to which you want to attach a file.

.PARAMETER File
Specifies the path to the file that you want to attach to the ServiceNow ticket.

.EXAMPLE
# Example 1: Upload an attachment to an incident using TicketNumber
Add-ServiceNowAttachment -TicketType "incident" -TicketNumber "INC0123456" -File "C:\Documents\attachment.pdf"

This example uploads the "attachment.pdf" file to the incident with Ticket Number "INC0123456" in ServiceNow.

.EXAMPLE
# Example 2: Upload an attachment to a service task using the SysID
Add-ServiceNowAttachment -TicketType "sc_task" -TicketSysID "57af7aec73d423002728660c4cf6a71c" -File "D:\Files\document.docx"

This example uploads the "document.docx" file to the service task with Sys ID "57af7aec73d423002728660c4cf6a71c" in ServiceNow.

#>
param(
[Parameter(Mandatory)]
[ValidateSet("sc_task","incident")]
$TicketType,
$TicketNumber,
$TicketSysID,
$File
)
    if($TicketNumber){
        switch($TicketType){
            "sc_task" {
                $TicketSysID = (Get-ServiceNowRecord -RecordType ScheduledTask -TicketNumber $TicketNumber).sys_id
            }
            "incident" {
                $TicketSysID = (Get-ServiceNowRecord -RecordType Incident -TicketNumber $TicketNumber).sys_id
            }
        }
    }

    if(-not $TicketSysID){
        Write-Host "Ticket SysID or Number is required for this function!" -ForegroundColor Red
        return
    }

    #Upload Attachment to Ticket in ServiceNow
    if($File -and (Test-Path $File -PathType Leaf)){
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

    try{
        #$global:SN_Submit_Attachment = Invoke-WebRequest -Uri "https://$ServiceNow_Server/sys_attachment.do?sysparm_record_scope=global" -Method "POST" -ContentType "multipart/form-data; boundary=---------------------------$SN_Attachment_GUID" -Body $SN_Attachment_Body -WebSession $ServiceNow_Session
        $global:SN_Submit_Attachment = New-ServiceNowWebRequest -Endpoint "/sys_attachment.do?sysparm_record_scope=global" -Method Post -ContentType "multipart/form-data; boundary=---------------------------$SN_Attachment_GUID" -Body $SN_Attachment_Body
        if ($SN_Submit_Attachment.StatusCode -eq "200") {
            Write-Host "*** Successfully Submitted Attachment `"$SN_Attachment_FileName`" for Ticket $TicketNumber ***" -ForegroundColor Green
        }else{
            Write-Host "File attachment upload failed!`nStatus: $($SN_Submit_Attachment.StatusCode)`n"
        }
    }catch{
        Write-Host "File attachment upload failed!`nError: $($_.Exception.Message)`n"
    }
}

function Close-ServiceNowIncident{
param(
$SysID,$TicketNum,$State,$CloseCode,$CloseNotes
)
    #If TicketNum provided, pull SysID
    if($TicketNum -ne "" -and $TicketNum -ne $null){$SysID = (Get-ServiceNowRecord -RecordType Incident -TicketNumber $TicketNum).sys_id}
    #If SysID doesn't exist, exit function
    if($SysID -eq "" -or $SysID -eq $null){Write-Host "Missing Incident SysID! Please provide and try again!" -ForegroundColor Red;return}
    
    #If State is not integer, convert it
    try{[int]$State}catch{try{$State = (Get-ServiceNowList -Name 'incident.state' | where {$_.name -eq $State}).value}catch{Write-Host "Error converting Incident State to an integer value!" -ForegroundColor Red;return}}

    #If State is Resolved, verify CloseCode and CloseNotes were provided and not Null or Blank
    if($State -eq (Get-ServiceNowList -Name 'incident.state' | where {$_.name -eq "Resolved"}).value){
        if ($CloseNotes -eq "" -or $CloseNotes -eq $null){Write-Host "Please provide a Close Note for a Resolution." -ForegroundColor Red;return}
        if ($CloseCode -eq "" -or $CloseCode -eq $null){Write-Host "Please provide a Close Code for a Resolution." -ForegroundColor Red;return}
        if( -not ((Get-ServiceNowList -Name 'incident.close_code').value.Contains($CloseCode)) ){Write-Host "Incident Close Code is not valid. Please try again!" -ForegroundColor Red;return}
        $body = @{"state" = $State;"close_code"=$CloseCode;"close_notes"=$CloseNotes} | ConvertTo-Json -Compress
    }else{
        $body = @{"state" = $State} | ConvertTo-Json -Compress
    }

    return (New-ServiceNowWebRequest -Endpoint "/incident_list.do?JSONv2&sysparm_sys_id=$SysID&sysparm_action=update" -Method Post -ContentType "application/json" -Body $body -REST).records
}

function Close-ServiceNowSession{
    Get-EventSubscriber -Force | Unregister-Event -Force
    if($ServiceNow_Session_Timer){$ServiceNow_Session_Timer.enabled = $false}
    Remove-Variable -Name "ServiceNow_*", "SN_*" -Scope Global -ErrorAction SilentlyContinue
}

function Confirm-ServiceNowSession{
    if($ServiceNow_Session){
        #$SN_User_Profile_Page_Refresh = (Invoke-RestMethod -Uri "https://$ServiceNow_Server/sys_user.do?JSONv2&sysparm_action=get&sysparm_sys_id=$SN_UserID" -WebSession $ServiceNow_Session).records
        $SN_User_Profile_Page_Refresh = (New-ServiceNowWebRequest -Endpoint "/sys_user.do?JSONv2&sysparm_action=get&sysparm_sys_id=$SN_UserID" -REST).records
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
    $Certificates = [System.Security.Cryptography.X509Certificates.X509Certificate2[]](Get-ChildItem Cert:\CurrentUser\My | where {$_.NotAfter -gt (Get-Date) -and $_.EnhancedKeyUsageList.FriendlyName -match "Smart Card Logon|Client Authentication"})
    $Certificates2 = $Certificates.psobject.Copy()

    $Certificates2 | Add-Member -MemberType NoteProperty -Name "Index" -Value 0
    $i=0
    foreach($Cert in $Certificates2){$Cert.Index=$i;$i++}
    Write-Host "`n******Smart Card Certificates******`n" -ForegroundColor Yellow
    write-host $(($Certificates2 | select Index,Thumbprint,FriendlyName,@{l="Issuer";e={$_.Issuer.Split(",")[0]}} |  Out-String).Trim())

    Write-Host "`nCertificate #: " -NoNewline -ForegroundColor Yellow
    $i = Read-Host

    return $Certificates[$i]
}

function Get-File {
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
        Filter = 'All files (*.*)|*.*|Archive|*.7z;*.cab;*.tar;*.gz;*.zip|CSV|*.csv|Excel (*.xls;*.xlsx)|*.xls*|HTML|*.html|Image|*.bmp;*.gif;*.jpg;*.jpeg|JSON|*.json|Outlook|*.msg|PDF|*.pdf|PowerPoint|*.pptx|PS1|*.ps1|TXT|*.rtf;*.txt|Visio|*.vsdx|Word (*.doc;*.docx)|*.doc*|XML|*.xml'
    }

    $null = $FileBrowser.ShowDialog()
    return $FileBrowser | Select SafeFileName, FileName
}

function Get-MimeType {
    param( 
        [Parameter(Mandatory, ValueFromPipeline=$true, Position=0)] 
        [ValidateScript({Test-Path $_})]
        [String]$File
        )
    Add-Type -AssemblyName "System.Web"
    [System.Web.MimeMapping]::GetMimeMapping($File)
}

function Get-ServiceNowCategories {
    $global:SN_CATsFilePath = "$($PSScriptRoot)\ServiceNow_Categories.json"

    if(Test-Path $SN_CATsFilePath){
        $global:ServiceNow_Categories = (Get-Content $SN_CATsFilePath -Raw) | ConvertFrom-Json
        Write-Host "ServiceNow Categories JSON file import successful!" -ForegroundColor Green
    }else{
        Write-Host "ServiceNow Categories JSON file not found!" -ForegroundColor Red
        Write-Host "Download latest ServiceNow Categories JSON file?(y/n): " -ForegroundColor Yellow -NoNewline
        $confirm = Read-Host

        if($confirm.ToLower() -eq "y" -or $confirm.ToLower() -eq "yes"){
            Update-ServiceNowCategories
            $global:ServiceNow_Categories = (Get-Content $SN_CATsFilePath -Raw) | ConvertFrom-Json
            Write-Host "Service Now Categories hash table created successfully!" -ForegroundColor Green
        }else{
            return $null
        }
    }
}

function Get-ServiceNowGroups {
    $global:SN_GroupsFilePath = "$($PSScriptRoot)\ServiceNow_Groups.json"

    if(Test-Path $SN_GroupsFilePath){
        $global:ServiceNow_Groups = (Get-Content $SN_GroupsFilePath -Raw) | ConvertFrom-Json
        Write-Host "ServiceNow Groups JSON file import successful!" -ForegroundColor Green
    }else{
        Write-Host "ServiceNow Groups JSON file not found!" -ForegroundColor Red
        Write-Host "Download latest ServiceNow Groups JSON file?(y/n): " -ForegroundColor Yellow -NoNewline
        $confirm = Read-Host

        if($confirm.ToLower() -eq "y" -or $confirm.ToLower() -eq "yes"){
            Update-ServiceNowGroups
            $global:ServiceNow_Groups = (Get-Content $SN_GroupsFilePath -Raw) | ConvertFrom-Json
            Write-Host "Service Now Groups array created successfully!" -ForegroundColor Green
        }else{
            return $null
        }
    }
}

function Get-ServiceNowList{
<#
.SYNOPSIS
Retrieves a Choice/Pick list's labels and values in ServiceNow.

.EXAMPLE
$ServiceNow_Incident_States = Get-ServiceNowList -Name "incident.state"
#>
param(
$Name
)
    if ($ServiceNow_Lists.Contains($Name)){
        return $ServiceNow_Lists.$($Name)
    }else{
        #$List = (Invoke-RestMethod -UseBasicParsing -Uri "https://$ServiceNow_Server/xmlhttp.do" -Method "POST" -WebSession $ServiceNow_Session -ContentType "application/x-www-form-urlencoded; charset=UTF-8" -Body "sysparm_processor=PickList&sysparm_scope=global&sysparm_want_session_messages=true&sysparm_name=$Name&sysparm_chars=*&sysparm_nomax=true").xml[1].ChildNodes
        $List = (New-ServiceNowWebRequest -Endpoint "/xmlhttp.do" -Method Post -ContentType "application/x-www-form-urlencoded; charset=UTF-8" -Body "sysparm_processor=PickList&sysparm_scope=global&sysparm_want_session_messages=true&sysparm_name=$Name&sysparm_chars=*&sysparm_nomax=true" -REST).xml[1].ChildNodes
        $ServiceNow_Lists.Add($Name,$List)
        return $List
    }
}

function Get-ServiceNowRecord{
param(
[Parameter(Mandatory)]
[ValidateSet("ChangeRequest","ChangeTask","CustomerServiceCase","Group","Incident","Request","RequestItem","ScheduledTask","User","ConfigurationItem")]
$RecordType,
$SysID,
$FirstName,
$LastName,
$GroupName,
$ComputerName,
$GroupNameSearch,
$TicketNumber,
$TicketSearch,
[switch]$GetHistory
)
    switch ($RecordType.toLower()){
        "user" {
            $RecordTypeURL = "sys_user_list.do"
            if ($SysID){
                $SN_Query = "sys_id=$SysID"
            }elseif($FirstName -and $LastName){
                $SN_Query = "last_name=$LastName^first_name=$FirstName"
            }elseif($FirstName){
                $SN_Query = "first_name=$FirstName"
            }elseif($LastName){
                $SN_Query = "last_name=$LastName"
            }else{
                Write-Host "A Sys ID or First/Last name is required to run this command." -ForegroundColor Red
                return
            }
        }
        "group" {
            $RecordTypeURL = "sys_user_group_list.do"
            if ($SysID){
                $SN_Query = "sys_id=$SysID"
            }elseif($GroupName){
                $SN_Query = "name=$GroupName"
            }elseif($GroupNameSearch){
                $SN_Query = "nameLIKE$GroupNameSearch" #-WebSession $ServiceNow_Session).records   #| select assignment_group,closed_at,closed_by,description,impact,number,opened_by,parent,priority,short_description,state,sys_created_on,sys_id,sys_updated_by,sys_updated_on,urgency
            }else{
                Write-Host "A Sys ID or Group name is required to run this command." -ForegroundColor Red
                return
            }
        }
        "scheduledtask" {
            $RecordTypeURL = "sc_task_list.do"
            if ($SysID){
                $SN_Query = "sys_id=$SysID"
            }elseif($TicketNumber){
                $SN_Query = "number=$TicketNumber"
            }elseif($TicketSearch){
                $SN_Query = "short_descriptionLIKE$TicketSearch^ORdescriptionLIKE$TicketSearch"
            }else{
                Write-Host "A SysID or Ticket Number is required to run this command." -ForegroundColor Red
                return
            }
        }
        "changerequest" {
            $RecordTypeURL = "change_request_list.do"
            if ($SysID){
                $SN_Query = "sys_id=$SysID"
            }elseif($TicketNumber){
                $SN_Query = "number=$TicketNumber"
            }else{
                Write-Host "A SysID or Ticket Number is required to run this command." -ForegroundColor Red
                return
            }
        }
        "changetask" {
            $RecordTypeURL = "change_task_list.do"
            if ($SysID){
                $SN_Query = "sys_id=$SysID"
            }elseif($TicketNumber){
                #https://$ServiceNow_Server/sc_task_list.do?sysparm_query=number%3DSCTASK0345015&sysparm_first_row=1&sysparm_view=
                $SN_Query = "number=$TicketNumber"
            }else{
                Write-Host "A SysID or Ticket Number is required to run this command." -ForegroundColor Red
                return
            }
        }
        "incident" {
            $RecordTypeURL = "incident_list.do"
            if ($SysID){
                $SN_Query = "sys_id=$SysID"
            }elseif($TicketNumber){
                $SN_Query = "number=$TicketNumber"
            }else{
                Write-Host "A SysID or Ticket Number is required to run this command." -ForegroundColor Red
                return
            }
        }
        "request" {
            $RecordTypeURL = "sc_request_list.do"
            if ($SysID){
                $SN_Query = "sys_id=$SysID"
            }elseif($TicketNumber){
                $SN_Query = "number=$TicketNumber"
            }else{
                Write-Host "A SysID or Ticket Number is required to run this command." -ForegroundColor Red
                return
            }
        }
        "requestitem" {
            $RecordTypeURL = "sc_req_item_list.do"
            if ($SysID){
                $SN_Query = "sys_id=$SysID"
            }elseif($TicketNumber){
                $SN_Query = "number=$TicketNumber"
            }else{
                Write-Host "A SysID or Ticket Number is required to run this command." -ForegroundColor Red
                return
            }
        }
        "configurationitem" {
            $RecordTypeURL = "cmdb_ci_pc_hardware.do"
            if ($SysID){
                $SN_Query = "sys_id=$SysID"
            }elseif($ComputerName){
                $SN_Query = "name=$ComputerName"
            }else{
                Write-Host "A SysID or Computer Name is required to run this command." -ForegroundColor Red
                return
            }
        }
        "customerservicecase" {
            $RecordTypeURL = "sn_customerservice_case_list.do"
            if ($SysID){
                $SN_Query = "sys_id=$SysID"
            }elseif($TicketNumber){
                $SN_Query = "number=$TicketNumber"
            }else{
                Write-Host "A SysID or Ticket Number is required to run this command." -ForegroundColor Red
                return
            }
        }
        Default {
            $RecordTypeURL = "incident_list.do"
            if ($SysID){
                $SN_Query = "sys_id=$SysID"
            }elseif($TicketNumber){
                $SN_Query = "number=$TicketNumber"
            }else{
                Write-Host "A SysID or Ticket Number is required to run this command." -ForegroundColor Red
                return
            }
        }
    }

    if($GetHistory.IsPresent){
        $RecordType2 = $RecordTypeURL -replace "_list\.do",""
        if($SysID -eq $null -and $TicketNumber){
            $SysID = (Get-ServiceNowRecord -RecordType $RecordType -TicketNumber $TicketNumber).sys_id
        }elseif($SysID -eq $null){
            Write-Host "A ticket number or Sys ID is required to retrieve ticket history." -ForegroundColor Red
            return
        }

        New-ServiceNowWebRequest -Endpoint "/angular.do?sysparm_type=user_preference&sysparm_pref_name=$RecordType2.activity.filter&sysparm_action=set&sysparm_pref_value=assigned_to%2Ctrue%3Bcmdb_ci%2Ctrue%3Bstate%2Ctrue%3Bimpact%2Ctrue%3Bpriority%2Ctrue%3Bopened_by%2Ctrue%3Bwork_notes%2Ctrue%3Bcomments%2Ctrue%3B*Attachments*%2Ctrue%3Bshort_description%2Ctrue%3Bassignment_group%2Ctrue%3B*EmailAutogenerated*%2Ctrue%3B*EmailCorrespondence*%2Ctrue" | Out-Null
        return (New-ServiceNowWebRequest -Endpoint "/angular.do?sysparm_type=list_history&table=$RecordType2&action=get_new_entries&sysparm_silent_request=true&sysparm_auto_request=true&sysparm_timestamp=&include_attachments=&sys_id=$SysID" -REST).entries
    }

    return (New-ServiceNowWebRequest -Endpoint "/$($RecordTypeURL)?JSONv2&sysparm_query=$SN_Query" -REST).records
}

function Get-ServiceNowServices {
    $global:ServiceNowServicesFilePath = "$($PSScriptRoot)\ServiceNow_Services.json"

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

function Get-ServiceNowStats {
    $SN_Stats_MasterStr = "*****Instance Information*****`n"
    $SN_Stats = New-ServiceNowWebRequest -Endpoint "/stats.do" -REST
    $StartPos = $SN_Stats.IndexOf("<br/>")
    $LastPos = $SN_Stats.LastIndexOf("<br/>")
    while($StartPos -lt $LastPos){
        $ParseStr = (Parse-String $SN_Stats.Substring($StartPos) "<br/>" "<")

        if($ParseStr -match ":"){
            $SN_Stats_MasterStr += $ParseStr + "`n"
        }

        $NextStrong = $SN_Stats.IndexOf("<strong>",$StartPos+5)
        $StartPos = $SN_Stats.IndexOf("<br/>",$StartPos+5)
        if($NextStrong -lt $StartPos -and $NextStrong -ne -1){
            $SN_Stats_MasterStr += "`n"
            $SN_Stats_MasterStr += "*****" + (Parse-String $SN_Stats.Substring($NextStrong) "<strong>" "</strong>") + "*****" + "`n"
        }
    }
    return $SN_Stats_MasterStr
}

function Get-ServiceNowUserUnique{
param($FirstName,$LastName)
    $UserSearch = Get-ServiceNowRecord -RecordType User -FirstName $FirstName -LastName $LastName
    if($UserSearch.Count -ne 1){
        while($True){
            Write-Host "`nMultiple ServiceNow users found:" -ForegroundColor Cyan
            $i=0
            foreach($User in $UserSearch){
                Write-Host "$i - $($User.name)"
                $i++
            }
            Write-Host "`nPlease select a user: " -NoNewline -ForegroundColor Cyan
            try{
                $resp = [int](Read-Host)
                return $UserSearch[$resp]
            }catch{
                Write-Host "Invalid response. Try again!`n" -ForegroundColor Red
                continue
            }
        }
    }else{
        return $UserSearch[0]
    }
}

#Needs cleaned up
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
        assignment_group                = $SN_Ticket_AssignmentGroup
        assigned_to                     = $SN_Ticket_AssignedTo
        short_description               = $SN_Ticket_ShortDescription
        description                     = $SN_Ticket_Description
    }
     $SN_Ticket_Body = $SN_Ticket_Body | ConvertTo-Json

    #$SN_Ticket_Headers = @{
    #    'Accept' = "application/json"
    #    'X-UserToken' = $SN_User_Token
    #}

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
        <#
        while($True){
            try{#STOPPED RIGHT HERE
                #$Submit_INC = Invoke-WebRequest -Uri "https://$ServiceNow_Server/incident.do?JSONv2&sysparm_action=insert" -Method "POST" -ContentType "application/json" -Headers $SN_Ticket_Headers -Body $($SN_Ticket_Body | ConvertTo-Json) -WebSession $ServiceNow_Session
                $Submit_INC = New-ServiceNowWebRequest -Endpoint "/incident.do?JSONv2&sysparm_action=insert" -Method Post -ContentType "application/json" -Body $SN_Ticket_Body -REST
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
        #>
        $Submit_INC = New-ServiceNowWebRequest -Endpoint "/incident.do?JSONv2&sysparm_action=insert" -Method Post -ContentType "application/json" -Body $SN_Ticket_Body
        $Submit_INC_2 = ($Submit_INC.Content | ConvertFrom-JSON).records[0]
        if (($Submit_INC.StatusCode -eq "200")) {
            $Global:INC_Number = $Submit_INC_2.number
            $Global:INC_SysID = $Submit_INC_2.sys_id
            Write-Host "*** Successfully Submitted `"$INC_SysID`" to ServiceNow ***`n" -ForegroundColor Green
            if($File -ne ""){
                Add-ServiceNowAttachment -TicketType 'incident' -TicketSysID $INC_SysID -File $File
            }
            return "$INC_Number,$INC_SysID"
        }
    }else{
        Write-Host "Aborting Ticket Creation!`n" -ForegroundColor Red
    }
}

#Needs cleaned up
function New-ServiceNowIncidentAdvanced{
<#
.SYNOPSIS
Creates a new Incident in ServiceNow using a custom ticket body.

.DESCRIPTION
This function allows you to submit an Incident ticket to ServiceNow using a custom ticket body.

.PARAMETER SN_Ticket_Body
Specifies the details of the ticket body which is submitted directly to ServiceNow.

.PARAMETER File
Specifies the path to the file that you want to attach to the ServiceNow ticket.

.PARAMETER SkipVerification
Skips the manual ticket details verification process before ticket is submitted.

.EXAMPLE

# Example 1: Creates a new Incident in ServiceNow using a custom ticket body.

$INC_Body = [ordered]@{
    caller_id                       = "Tuter, Abel"
    business_service                = "IT Services"
    category                        = "Software"
    subcategory                     = "Email"
    contact_type                    = "Phone"
    severity                        = "4 - Low"
    urgency                         = "4 - Low"
    assignment_group                = "Change Management"
    assigned_to                     = "Smith, David"
    short_description               = "Short Desc Here"
    description                     = "Long Desc Here`nLine 2`nLine3"
}

New-ServiceNowIncidentAdvanced -SN_Ticket_Body $INC_Body
#>
param(
$SN_Ticket_Body,
[String]$File="",
[switch]$SkipVerification
)
    Write-Host "`n*** Incident Details Overview ***" -ForegroundColor Yellow
    Write-Host "$(($SN_Ticket_Body | ft -HideTableHeaders -AutoSize -Wrap | Out-String).Trim())`n" -ForegroundColor Cyan

    if (-not $SkipVerification){
        $confirm = Read-Host "Continue(y/n)"
    }else{
        $confirm = "y"
    }

    if($confirm.ToLower() -eq "y"){
        $SN_Ticket_Body = ($SN_Ticket_Body | ConvertTo-Json -Compress)
        $Global:INC_Submit = New-ServiceNowWebRequest -Endpoint "/incident.do?JSONv2&sysparm_action=insert" -Method Post -ContentType "application/json" -Body $SN_Ticket_Body -REST
        if(-not ($INC_Submit | Get-Member -Name records)){
            Write-Host "Error occured during web request! Exiting function!" -ForegroundColor Red
            return $INC_Submit
        }
        $Global:INC_Number = $INC_Submit.records[0].number
        $Global:INC_SysID = $INC_Submit.records[0].sys_id
        Write-Host "*** Successfully Submitted `"$INC_Number`" to ServiceNow ***`n" -ForegroundColor Green
        if($File -ne ""){
                Write-Host "Uploading file ($File) to $INC_Number..." -ForegroundColor Yellow
                Add-ServiceNowAttachment -TicketType 'incident' -TicketSysID $INC_SysID -File $File
            }
        return $INC_Submit.records
    }else{
        Write-Host "Aborting Ticket Creation!`n" -ForegroundColor Red
    }
}

#SKIPPING New-ServiceNowWebRequest Conversion for now...
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
                    $Submit_SCTask = Invoke-WebRequest -UseBasicParsing -Uri "https://$ServiceNow_Server/api/sn_sc/v1/servicecatalog/items/5d167fd96ce4fc1004ed764f2fe89f42/order_now" -Method "POST" -ContentType "application/json" -Body $($SN_Ticket_Body | ConvertTo-Json) -WebSession $ServiceNow_Session -Headers $headers -ErrorVariable Submit_SCTask_Error
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
                $SCTaskQuery = Invoke-RestMethod -UseBasicParsing -Uri "https://$ServiceNow_Server/sc_task.do?JSONv2&sysparm_query=request=$RequestSysID" -WebSession $ServiceNow_Session
                $global:SCTask_SysID = $SCTaskQuery.records[0].sys_id
                $global:SCTask_Number = $SCTaskQuery.records[0].number

                Write-Host "*** Successfully Submitted `"$SCTask_Number`" to ServiceNow ***" -ForegroundColor Green
                if($File -ne ""){
                    Write-Host "Uploading file ($File) to $SCTask_Number..." -ForegroundColor Yellow
                    Add-ServiceNowAttachment -TicketType 'sc_task' -TicketSysID $SCTask_SysID -File $File
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
    [switch]$CertificateAuth
)
    if($Global:ServiceNow_Server -match "\*" -and !$Server){
        Write-Host "No server was provided for ServiceNow connection!" -ForegroundColor Red
        return
    }elseif($Global:ServiceNow_Server -match "\*" -and $Server){
        if($Server -match "http|https"){
            $Server = ($Server -replace "(https://|http://)","" -replace "/","")
        }
        $Global:ServiceNow_Server = $Server
    }

    if($ServiceNow_Server -match "http|https"){
        $ServiceNow_Server = ($ServiceNow_Server -replace "(https://|http://)","" -replace "/","")
    }

    Write-Host "Connecting to $ServiceNow_Server..." -ForegroundColor Yellow
    try{
        $SN_Login_Page = Invoke-WebRequest -UseBasicParsing -Uri "https://$ServiceNow_Server" -SessionVariable global:ServiceNow_Session -ErrorAction Stop
        if($SN_Login_Page.StatusCode -ne 200){
            Write-Host "Connection to ServiceNow failed!`nStatus Code: $($SN_Login_Page.StatusCode)"
            return
        }
    }catch{
        Write-Host "Connection to ServiceNow failed!`nError: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    if ($SN_Login_Page.Content -match "g_ck = '(.*)'") {$SN_GCK_Token = $matches[1];write-host "Found G_CK Token: $($SN_GCK_Token.Substring(0,10))...." -ForegroundColor Green}
    
    try{
        if($ServiceNow_Server -match "aesmp\.army\.mil"){
            #Create AESMP web session
            $AESMP_MainPage = Invoke-RestMethod -UseBasicParsing -Uri "https://$ServiceNow_Server" -SessionVariable global:ServiceNow_Session -Verbose
            $Portal_ID = Parse-String -String $AESMP_MainPage -StartStr "ng-init=`"portal_id = '" -EndStr "'"
            $UnixEpochTime = ((Get-Date -UFormat %s) -replace "\.","").Substring(0,13)

            #Retrieve Glide SSO ID from AESMP
            Invoke-RestMethod -UseBasicParsing -Uri "https://$ServiceNow_Server/csm?id=landing" -WebSession $ServiceNow_Session | Out-Null
            $AESMP_LandingPage = (Invoke-WebRequest -UseBasicParsing -Uri "https://$ServiceNow_Server/api/now/sp/page?id=landing&time=$UnixEpochTime&portal_id=$Portal_ID&request_uri=%2Fcsm%3Fid%3Dlanding" -WebSession $ServiceNow_Session).Content
            $Glide_SSO_ID = Parse-String -String $AESMP_LandingPage -StartStr '"href":"/login_with_sso.do?glide_sso_id=' -EndStr "`""
            ##$AESMP_SSO_Endpoint = Parse-String -String $AESMP_LandingPage -StartStr '"href":"' -EndStr '"'

            #Retrieve HTTP Redirect for EAMS Authentication
            $AESMP_Login_SSO = Invoke-WebRequest -UseBasicParsing -Uri "https://$ServiceNow_Server/login_with_sso.do?glide_sso_id=$Glide_SSO_ID" -WebSession $ServiceNow_Session
            ##$AESMP_Login_SSO = Invoke-WebRequest -UseBasicParsing -Uri "https://$ServiceNow_Server$AESMP_SSO_Endpoint" -WebSession $ServiceNow_Session
            $EAMS_Redirect_URL = $AESMP_Login_SSO.BaseResponse.ResponseUri.AbsoluteUri
            $EAMS_Redirect = Invoke-RestMethod -UseBasicParsing -Uri $EAMS_Redirect_URL -WebSession $ServiceNow_Session
            $EAMS_URL = Parse-String -String $EAMS_Redirect -StartStr "top.location.href = '" -EndStr "'"

            #Retrieve required tokens from EAMS Main Page
            $EAMS_MainPage = Invoke-RestMethod -UseBasicParsing -Uri $EAMS_URL -WebSession $ServiceNow_Session

            #Create hashtable to convert login data to correct key values for EAMS login POST request
            $EAMS_Auth_Request_Elements_Map = [ordered]@{
                "authenticity_token" = "authenticity_token"
                "sso_session_orig_url" = "sso_session[orig_url]"
                "sso_session_orig_method" = "sso_session[orig_method]"
                "sso_session_renewed_session" = "sso_session[renewed_session]"
                "sso_session_pki_upgrade" = "sso_session[pki_upgrade]"
                "SAMLRequest" = "SAMLRequest"
                "RelayState" = "RelayState"
            }
            $EAMS_Auth_Request_Elements = ($EAMS_Auth_Request_Elements_Map.keys | Out-String -Stream)[1..$EAMS_Auth_Request_Elements_Map.Count]

            #Create POST request body required for EAMS login
            $EAMS_Auth_Request = [ordered]@{}
            $EAMS_Auth_Request.Add("authenticity_token",(Parse-String $EAMS_MainPage -StartStr 'name="authenticity_token" value="' -EndStr '"'))
            foreach($Element in $EAMS_Auth_Request_Elements){
                $EAMS_Auth_Request.Add($EAMS_Auth_Request_Elements_Map[$Element],(Parse-String $EAMS_MainPage -StartStr "id=`"$Element`" value=`"" -EndStr '"'))
            }

            #Retrieve Smart Card certificate and login to EAMS
            $global:SN_Cert = Get-AuthCertificate
            $EAMS_Login = Invoke-WebRequest -UseBasicParsing -Uri "https://federation.eams.army.mil/pool/sso/saml/authenticate?request_client_cert=true" -WebSession $ServiceNow_Session -Certificate $SN_Cert -Method Post -ContentType "application/x-www-form-urlencoded" -Body $EAMS_Auth_Request -Verbose
            $Global:EAMS_Login_Redirect_URL = $EAMS_Login.BaseResponse.ResponseUri.AbsoluteUri
            #Write-Host "EAMS Redirect URL: $EAMS_Login_Redirect_URL`n"

            #Return EAMS login status message
            if($EAMS_Login.Content -match "Your account has been successfully authenticated"){
                Write-Host "`nEAMS authentication successful!" -ForegroundColor Green
            }else{
                Write-Host "`nEAMS authentication failed!" -ForegroundColor Red
            }

            #Retrieve login tokens from EAMS website to forward to AESMP
            $EAMS_Login_Redirect = Invoke-WebRequest -UseBasicParsing -Uri $EAMS_Login_Redirect_URL -WebSession $ServiceNow_Session
            $AESMP_Auth_Request = @{
                "authenticity_token" = (Parse-String $EAMS_Login_Redirect.Content -StartStr 'name="authenticity_token" value="' -EndStr '"')
                "SAMLResponse" = (Parse-String $EAMS_Login_Redirect.Content -StartStr "id=`"SAMLResponse`" value=`"" -EndStr '"')
                "RelayState" = (Parse-String $EAMS_Login_Redirect.Content -StartStr "id=`"RelayState`" value=`"" -EndStr '"')
            }

            #Login to AESMP using previous SAML login tokens pulled from EAMS
            $AESMP_Login = Invoke-WebRequest -UseBasicParsing -Uri "https://$ServiceNow_Server/navpage.do" -WebSession $ServiceNow_Session -Method Post -ContentType "application/x-www-form-urlencoded" -Body $AESMP_Auth_Request
            $AESMP_Login_Redirect_URL = $AESMP_Login.BaseResponse.ResponseUri.AbsoluteUri
            $SN_Banner_Page = Invoke-WebRequest -UseBasicParsing -Uri $AESMP_Login_Redirect_URL -WebSession $ServiceNow_Session
        }elseif($Username -and $Pass){
            $SN_Banner_Page = Invoke-WebRequest -UseBasicParsing -Uri "https://$ServiceNow_Server/login.do" -Method "POST" -ContentType "application/x-www-form-urlencoded" -Body @{
                "sysparm_ck" = $SN_GCK_Token
                "user_name" = $Username
                "user_password" = $Pass
                "not_important"=$null
                "ni.nolog.user_password" = $true
                "ni.noecho.user_name" = $true
                "ni.noecho.user_password" = $true
                "sys_action" = "sysverb_login"
                "sysparm_login_url" = "welcome.do"} -WebSession $ServiceNow_Session
        }elseif($CertificateAuth.IsPresent){
            $global:SN_Cert = Get-AuthCertificate
            $SN_Banner_Page = Invoke-WebRequest -UseBasicParsing -Uri "https://$ServiceNow_Server/my.policy" -Certificate $SN_Cert -Method "POST" -ContentType "application/x-www-form-urlencoded" -Body "choice=1" -WebSession $ServiceNow_Session
        }else{
            Write-Host "ServiceNow session was not created. Session type was not specified." -ForegroundColor Red
            return
        }

        if($SN_Banner_Page.StatusCode -ne 200){
            Write-Host "Authentication to ServiceNow failed!`nStatus Code: $($SN_Banner_Page.StatusCode)"
            return
        }
    }catch{
        Write-Host "Authentication to ServiceNow failed!`nError: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
    
    if($SN_Banner_Page.Content -match "Session Expired|logged_in = false"){
        Write-Host "Authentication to ServiceNow failed!`n" -ForegroundColor Red
        return
    }else{
        Write-Host "Authenticated to ServiceNow successfully!`n" -ForegroundColor Green
    }

    #Retrieve and Set Current User Settings
    if ($SN_Banner_Page.Content -match "window.NOW.user.userID = '(.*?)'") {$global:SN_UserID = $matches[1];write-host "User ID: $SN_UserID" -ForegroundColor Green}
    if ($SN_Banner_Page.Content -match "window.NOW.user_id = '(.*?)'") {$global:SN_UserID = $matches[1];write-host "User ID: $SN_UserID" -ForegroundColor Green}
    if ($SN_Banner_Page.Content -match "`"userID`" : `"(.*?)`",") {$global:SN_UserID = $matches[1];write-host "User ID: $SN_UserID" -ForegroundColor Green} #For Admins I believe?
    if ($SN_Banner_Page.Content -match "g_ck = '(.*)'") {$global:SN_User_Token = $matches[1];write-host "Found User Token: $($SN_User_Token.Substring(0,10))....`n" -ForegroundColor Green}

    $ServiceNow_Session.Headers.Add("X-UserToken",$SN_User_Token)
    $global:SN_User_Profile_Page = (Invoke-RestMethod -UseBasicParsing -Uri "https://$ServiceNow_Server/sys_user.do?JSONv2&sysparm_action=get&sysparm_sys_id=$SN_UserID" -WebSession $ServiceNow_Session -ErrorAction Stop).records

    $global:SN_DisplayName = $SN_User_Profile_Page.name
    $global:SN_UserName = $SN_User_Profile_Page.user_name
    $global:SN_LocationID = $SN_User_Profile_Page.location
    $global:SN_Location_Name = (Invoke-RestMethod -UseBasicParsing -Uri "https://$ServiceNow_Server/xmlhttp.do" -Method "POST" -WebSession $ServiceNow_Session -ContentType "application/x-www-form-urlencoded; charset=UTF-8" -Body "sysparm_processor=AjaxClientHelper&sysparm_scope=global&sysparm_want_session_messages=true&sysparm_name=getDisplay&sysparm_table=cmn_location&sysparm_value=$SN_LocationID&sysparm_synch=true&ni.nolog.x_referer=ignore").xml.answer

    Write-Host "Display Name: $SN_DisplayName`nUsername: $SN_UserName`nLocation: $SN_Location_Name" -ForegroundColor Green

    $ServiceNow_Session_Expires = ($ServiceNow_Session.Cookies.GetCookies("https://$ServiceNow_Server") | where {$_.Name -eq "glide_session_store"}).Expires
    $global:ServiceNow_Session_Expires_Minutes = [math]::Floor((New-TimeSpan -Start (Get-Date) -End $ServiceNow_Session_Expires).TotalMinutes)
    Write-Host "Session Expiry: $ServiceNow_Session_Expires_Minutes minutes (Refreshes every 9.5 minutes)`n`n" -ForegroundColor Yellow
    New-SNSessionRefresher
}

function New-ServiceNowWebRequest{
param(
    $Endpoint,
    [ValidateSet("Get","Post")]$Method="Get",
    $ContentType="",
    $Body,
    [switch]$REST
)
    if (!$ServiceNow_Session){Confirm-ServiceNowSession}

    for($Retry=1;$Retry -le 3;$Retry++){
        try{
            if($REST.IsPresent){
                if($Body -ne $null -and $Body -ne ""){
                    $ServiceNow_WR = Invoke-RestMethod -UseBasicParsing "https://$ServiceNow_Server$Endpoint" -WebSession $ServiceNow_Session `
                    -Method $Method -ContentType $ContentType -Body $Body
                }else{
                    $ServiceNow_WR = Invoke-RestMethod -UseBasicParsing "https://$ServiceNow_Server$Endpoint" -WebSession $ServiceNow_Session `
                    -Method $Method
                }
            }else{
                if($Body -ne $null -and $Body -ne ""){
                    $ServiceNow_WR = Invoke-WebRequest -UseBasicParsing "https://$ServiceNow_Server$Endpoint" -WebSession $ServiceNow_Session `
                    -Method $Method -ContentType $ContentType -Body $Body
                }else{
                    $ServiceNow_WR = Invoke-WebRequest -UseBasicParsing "https://$ServiceNow_Server$Endpoint" -WebSession $ServiceNow_Session `
                    -Method $Method
                }
            }
            return $ServiceNow_WR
        }catch{
            if($Retry -eq 3){
                Write-Host "Failed to submit web request 3 times in a row..." -ForegroundColor Red
                Write-Host  "Try again?(y/n): " -ForegroundColor Yellow -NoNewline
                $resp = Read-Host
                if($resp.ToLower() -match "y|yes"){
                    $Retry=0
                }else{
                    break
                }
            }else{
                Write-Host "Error occured while submitting web request to SNOW! Retrying..." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
            }
        }
    }
}

function New-SNSessionRefresher{
    $global:ServiceNow_Session_Timer = New-Object System.Timers.Timer

    $Action = {
        $ServiceNow_Session_Expires = ($ServiceNow_Session.Cookies.GetCookies("https://$ServiceNow_Server") | where {$_.Name -eq "glide_session_store"}).Expires
        $global:ServiceNow_Session_Expires_Minutes = [math]::Floor((New-TimeSpan -Start (Get-Date) -End $ServiceNow_Session_Expires).TotalMinutes)

        #$SN_User_Profile_Page_Refresh = (Invoke-RestMethod -Uri "https://$ServiceNow_Server/sys_user.do?JSONv2&sysparm_action=get&sysparm_sys_id=$SN_UserID" -WebSession $ServiceNow_Session).records
        $SN_User_Profile_Page_Refresh = (New-ServiceNowWebRequest -Endpoint "/sys_user.do?JSONv2&sysparm_action=get&sysparm_sys_id=$SN_UserID" -REST).records
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

function Search-ServiceNowCustomer{
param($Name,$Fields)
    if($Fields -ne "" -and $Fields -ne $null){
        $ReturnFields = $Fields
    }else{
        $ReturnFields = "first_name;last_name;user_name;email"    
    }

    return (New-ServiceNowWebRequest -Endpoint "/xmlhttp.do" -Method Post -ContentType "application/x-www-form-urlencoded; charset=UTF-8" -Body "sysparm_processor=Reference&sysparm_scope=global&sysparm_want_session_messages=true&sysparm_name=incident.caller_id&sysparm_max=15&sysparm_chars=$Name&ac_columns=$Fields&ac_order_by=name" -REST).xml.ChildNodes
}

function Search-ServiceNowRecord{
<#
.SYNOPSIS
Executes a query for records in ServiceNow.

.EXAMPLE
$Search_SN_Record = Search-ServiceNowRecords -RecordType "User" -Query "first_name=Abel^last_name=Smith"
#>
param(
[Parameter(Mandatory)]
[ValidateSet("ChangeRequest","ChangeTask","CustomerServiceCase","Group","Incident","Request","RequestItem","ScheduledTask","User","ConfigurationItem")]
$RecordType,
[Parameter(Mandatory)]
$Query
)
    switch ($RecordType.toLower()){
        "user" {
            $RecordTypeURL = "sys_user_list.do"
        }
        "group" {
            $RecordTypeURL = "sys_user_group_list.do"
        }
        "scheduledtask" {
            $RecordTypeURL = "sc_task_list.do"
        }
        "changerequest" {
            $RecordTypeURL = "change_request_list.do"
        }
        "changetask" {
            $RecordTypeURL = "change_task_list.do"
        }
        "incident" {
            $RecordTypeURL = "incident_list.do"
        }
        "request" {
            $RecordTypeURL = "sc_request_list.do"
        }
        "requestitem" {
            $RecordTypeURL = "sc_req_item_list.do"
        }
        "configurationitem" {
            $RecordTypeURL = "cmdb_ci_pc_hardware_list.do"
        }
        "customerservicecase" {
            $RecordTypeURL = "sn_customerservice_case_list.do"
        }
    }
    if($RecordTypeURL -ne "" -and $RecordTypeURL -ne $null){
        $SN_WR = New-ServiceNowWebRequest -Endpoint "/$($RecordTypeURL)?JSONv2&sysparm_query=$Query" -REST
        if(-not ($SN_WR | Get-Member -Name records)){
            Write-Host "Error occured during web request! Exiting function!" -ForegroundColor Red
            return $INC_Submit
        }else{
            return $SN_WR.records
        }
    }else{
        Write-Host "Record Type was not found in switch statement. Exiting function..." -ForegroundColor Red
        return
    }
}

function Update-ServiceNowCategories {
    $global:SN_CATsFilePath = "$($PSScriptRoot)\ServiceNow_Categories.json"
    $i = 0
    $SN_CategoryListHash = [ordered]@{}
    
    $SN_CategoryList = Get-ServiceNowList -Name "incident.category"
    <#
    $CategoryList = (Invoke-RestMethod -UseBasicParsing -Uri "https://$ServiceNow_Server/xmlhttp.do" `
    -Method "POST" `
    -WebSession $ServiceNow_Session `
    -Headers @{
        "X-UserToken"=$SN_User_Token
    } `
    -ContentType "application/x-www-form-urlencoded; charset=UTF-8" `
    -Body "sysparm_processor=PickList&sysparm_scope=global&sysparm_want_session_messages=true&sysparm_name=incident.category&sysparm_chars=*&sysparm_nomax=true").xml[1].ChildNodes.Name
    #>

    foreach($Cat in $SN_CategoryList){
        if($i%5 -eq 0){
            Write-Host "$([int](($i/88)*100))%.." -NoNewline
        }

        $wr = New-ServiceNowWebRequest -Endpoint "/xmlhttp.do" -Method Post -ContentType "application/x-www-form-urlencoded; charset=UTF-8" -Body "sysparm_processor=PickList&sysparm_scope=global&sysparm_want_session_messages=true&sysparm_value=$cat&sysparm_name=incident.subcategory&sysparm_chars=*&sysparm_nomax=true" -REST
        <#
        $wr = Invoke-RestMethod -UseBasicParsing -Uri "https://$ServiceNow_Server/xmlhttp.do" `
        -Method "POST" `
        -WebSession $ServiceNow_Session `
        -Headers @{
            "X-UserToken"=$SN_User_Token
        } `
        -ContentType "application/x-www-form-urlencoded; charset=UTF-8" `
        -Body "sysparm_processor=PickList&sysparm_scope=global&sysparm_want_session_messages=true&sysparm_value=$cat&sysparm_name=incident.subcategory&sysparm_chars=*&sysparm_nomax=true"
        #>

        $SN_CategoryListHash[$Cat] = $wr.xml[1].ChildNodes.name
        $i++
    }
    Write-Host "100%..Download Complete! Saving to file..."
    $SN_CategoryListHash.GetEnumerator() | ConvertTo-Json | Out-File $SN_CATsFilePath -Verbose
    Write-Host "`nService Now Categories JSON file updated successfully!" -ForegroundColor Green
}

function Update-ServiceNowGroups {
    $global:SN_GroupsFilePath = "$($PSScriptRoot)\ServiceNow_Groups.json"
    #(Invoke-RestMethod -Uri "https://$ServiceNow_Server/sys_user_group_list.do?JSONv2" -WebSession $ServiceNow_Session).records | where {$_.name -ne "" -and $_.name -ne $null} | select name,sys_id,manager | sort name | ConvertTo-Json | Out-File $SN_GroupsFilePath -Force
    (New-ServiceNowWebRequest -Endpoint "/sys_user_group_list.do?JSONv2" -REST).records  | where {$_.name -ne "" -and $_.name -ne $null} | select name,sys_id,manager | sort name | ConvertTo-Json | Out-File $SN_GroupsFilePath -Force
    Write-Host "Service Now Groups JSON file updated successfully!" -ForegroundColor Green
}

function Update-ServiceNowRecord{
<#
.SYNOPSIS
Updates a single or multiple fields for a record in ServiceNow.

.EXAMPLE
#Example 1: Create body paramters to update and pass to Update-ServiceNowRecord command.
$BodyParams = @{
"state" = "Resolved"
"close_code"="Duplicate"
"close_notes"="This incident is a duplicate of INC1234567"
}

$Update_SN_record = Update-ServiceNowRecord -RecordType "Incident" -TicketNum "INC8675309" -BodyParams $BodyParams
#>
param(
[Parameter(Mandatory)]
[ValidateSet("ChangeRequest","ChangeTask","CustomerServiceCase","Group","Incident","Request","RequestItem","ScheduledTask","User","ConfigurationItem")]
$RecordType,
$SysID,
$TicketNum,
[Parameter(Mandatory)]
$BodyParams
)
    if($TicketNum -ne "" -and $TicketNum -ne $null){$SysID = (Get-ServiceNowRecord -RecordType Incident -TicketNumber $TicketNum).sys_id}

    if($SysID -eq "" -or $SysID -eq $null){Write-Host "Missing record SysID! Please provide and try again!" -ForegroundColor Red;return}

    if($BodyParams.GetType().Name -eq "Hashtable"){$BodyParams = $BodyParams | ConvertTo-Json -Compress}

    switch ($RecordType.toLower()){
        "user" {
            $RecordTypeURL = "sys_user_list.do"
        }
        "group" {
            $RecordTypeURL = "sys_user_group_list.do"
        }
        "scheduledtask" {
            $RecordTypeURL = "sc_task_list.do"
        }
        "changerequest" {
            $RecordTypeURL = "change_request_list.do"
        }
        "changetask" {
            $RecordTypeURL = "change_task_list.do"
        }
        "incident" {
            $RecordTypeURL = "incident_list.do"
        }
        "request" {
            $RecordTypeURL = "sc_request_list.do"
        }
        "requestitem" {
            $RecordTypeURL = "sc_req_item_list.do"
        }
        "configurationitem" {
            $RecordTypeURL = "cmdb_ci_pc_hardware.do"
        }
        "customerservicecase" {
            $RecordTypeURL = "sn_customerservice_case_list.do"
        }
    }
    if($RecordTypeURL -ne "" -and $RecordTypeURL -ne $null){
        return (New-ServiceNowWebRequest -Endpoint "/$($RecordTypeURL)?JSONv2&sysparm_sys_id=$($SysID)&sysparm_action=update" -Method Post -ContentType "application/json" -Body $BodyParams -REST).records
    }else{
        Write-Host "Record Type was not found in switch statement. Exiting function..." -ForegroundColor Red
        return
    }
}

function Update-ServiceNowServices {
    $global:ServiceNowServicesFilePath = "$($PSScriptRoot)\ServiceNow_Services.json"
    $ServiceNow_Services = (New-ServiceNowWebRequest -Endpoint "/cmdb_ci_service_list.do?JSONv2&sysparm_target=incident.business_service" -REST).records | where {$_.name -ne "" -and $_.name -ne $null} | select name,sys_id | sort name | ConvertTo-Json | Out-File $ServiceNowServicesFilePath -Force
    #$ServiceNow_Services = (Invoke-RestMethod -UseBasicParsing -Uri "https://$ServiceNow_Server/cmdb_ci_service_list.do?JSONv2&sysparm_target=incident.business_service" -WebSession $ServiceNow_Session -Headers @{"X-UserToken"=$SN_User_Token}).records | where {$_.name -ne "" -and $_.name -ne $null} | select name,sys_id | sort name | ConvertTo-Json | Out-File $ServiceNowServicesFilePath -Force
    Write-Host "Service Now Services JSON file updated successfully!" -ForegroundColor Green
}



Export-ModuleMember -Function Add-ServiceNowAttachment
Export-ModuleMember -Function Close-ServiceNowIncident
Export-ModuleMember -Function Close-ServiceNowSession
Export-ModuleMember -Function Confirm-ServiceNowSession
Export-ModuleMember -Function Get-AuthCertificate
Export-ModuleMember -Function Get-ServiceNowCategories
Export-ModuleMember -Function Get-ServiceNowGroups
Export-ModuleMember -Function Get-ServiceNowRecord
Export-ModuleMember -Function Get-ServiceNowServices
Export-ModuleMember -Function Get-ServiceNowStats
Export-ModuleMember -Function Get-ServiceNowUserUnique
Export-ModuleMember -Function New-ServiceNowIncident
Export-ModuleMember -Function New-ServiceNowIncidentAdvanced
Export-ModuleMember -Function Get-ServiceNowList
#Export-ModuleMember -Function New-ServiceNowSCTask           #This functions needs additional review/recoding
Export-ModuleMember -Function New-ServiceNowSession
Export-ModuleMember -Function New-ServiceNowWebRequest
Export-ModuleMember -Function Search-ServiceNowCustomer
Export-ModuleMember -Function Search-ServiceNowRecord
Export-ModuleMember -Function Update-ServiceNowCategories
Export-ModuleMember -Function Update-ServiceNowGroups
Export-ModuleMember -Function Update-ServiceNowRecord
Export-ModuleMember -Function Update-ServiceNowServices
