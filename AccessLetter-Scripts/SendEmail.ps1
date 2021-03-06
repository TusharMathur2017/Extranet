Clear-Host
if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null)
{ Add-PSSnapin "Microsoft.SharePoint.PowerShell" }
$ErrorActionPreference = "Stop"
cd "E:\Install\AccessLetter"
$siteURL = "http://connect.bain.com/al"
$siteURLCom = "http://community.bain.com/sites/Legal/professionalstandards/acessletter/"
$alSite = Get-SPWeb $siteURL

function SendEmail($Subject, $Message, $SentBy, $SentTo) {
    try {
        $smtp = New-object Net.Mail.SmtpClient("mail.bain.com");
        $from = New-Object Net.Mail.MailAddress($SentBy);
        $to = New-Object Net.Mail.MailAddress($SentTo);
        $msg = new-object Net.Mail.MailMessage($from, $to)
        $msg.subject = $Subject
        $msg.IsBodyHtml = $true;
        $msg.body = $Message;
        $smtp.Send($msg);
        $msg.Dispose();
        return $true
    }
    catch {
        write-host "P2->: " $_.Exception.Message
    }
}
function SetFormDigest() {
    $response = PostRequest ("_api/contextinfo", $null)
    $formDigest = $response.d.GetContextWebInformation.FormDigestValue
    $headers.Add("X-RequestDigest", $formDigest);
}
function PutRequest ($endpoint, $body) {
    if ($headers["IF-MATCH"].length -eq 0) {
        $headers.Add("IF-MATCH", "*");
        $headers.Add("X-HTTP-Method", "MERGE");
    }
    # $headers.Add("X-HTTP-Method", "MERGE");
    # $headers.Add("IF-MATCH", "*");
    return Request $endpoint $body ([Microsoft.PowerShell.Commands.WebRequestMethod]::Post)
}
function GetRequest ($endpoint, $body) {
    return Request $endpoint $body ([Microsoft.PowerShell.Commands.WebRequestMethod]::Get)
}
function PostRequest ($endpoint, $body) {
    return Request $endpoint $body ([Microsoft.PowerShell.Commands.WebRequestMethod]::Post)
}
function Request ($endpoint, $body, $method) {
    return Invoke-RestMethod -Uri ($url + $endpoint) -Headers $headers -Method $method -Body $body -Credential $cred -ContentType "application/json; odata=verbose"
}
function EmailManager() {
    $spQuery = New-Object Microsoft.SharePoint.SPQuery;
    $spQuery.Query = "<Where><And><Eq><FieldRef Name='ConsentOn' /><Value IncludeTimeValue='FALSE' Type='DateTime'><Today /></Value></Eq><And><Eq><FieldRef Name='Consent' /><Value Type='Boolean'>1</Value></Eq><Eq><FieldRef Name='ManagerInformed' /><Value Type='Boolean'>0</Value></Eq></And></And></Where>"
    $qualifiedItems = $alSite.Lists["Consent"].GetItems($spQuery);
    write-host ""
    write-host "Notifying $($qualifiedItems.Count) Managers"

    if ($qualifiedItems.Count -gt 0) {
        $headers = @{accept = "application/json; odata=verbose"}
        $url = $siteURLCom
        $cred = New-Object System.Management.Automation.PSCredential "Intranet_prdinstall", (Get-Content "E:\Install\AccessLetter\passkey.txt" | ConvertTo-SecureString)
        try {
            SetFormDigest
        }
        catch {
            write-host "Unable to establish connection function()storeToMainList`n$($_.Exception)"
            continue
        }
    }
    foreach ($eachItem in $qualifiedItems) {
        $userfield = New-Object Microsoft.SharePoint.SPFieldUserValue($alSite, $eachItem['Author'].ToString());
        $res = GetRequest("_api/web/lists/getbytitle('AccessletterDatabase')/GetItemById($($eachItem['Title']))", $null)
        if ($res.Length -gt 0) {
            $datacorrected = $res -creplace '"Id":', '"Idnew":'
            $resultsT = ConvertFrom-Json -InputObject $datacorrected
            $editThis = $resultsT.d.results.Idnew
            $CaseCode = $resultsT.d.CaseCode
            $prjName = $resultsT.d.ProjectName
            $vName = Remove-Diacritics($eachItem['SName'])
			$vEmail = Remove-Diacritics($eachItem['SEmail'])
			$vCompany = Remove-Diacritics($eachItem['SCompany'])
            $emailSubject = "$($prjName) $($CaseCode) $($vName) - Access to Bain Materials"
        }
        else {
            $emailSubject = "$($vName) Access to Bain Materials"
        }
        $emailBody = Get-Content -Path 'E:\Install\AccessLetter\EmailTemplate-Manager.html'
        $emailBody = $emailBody.Replace("[BAINCONTACT]", $userfield.User.Name.Split(",")[1].trim())
        $emailBody = $emailBody.Replace("[INDIVIDUALNAME]", $vName)
        $emailBody = $emailBody.Replace("[EMAILADDRESS]", $vEmail)
        $emailBody = $emailBody.Replace("[COMPANYNAME]", $vCompany)
        $emailBody = $emailBody.Replace("[ACCESSLINK]", "$($siteURL)/res/Consent.aspx?identity=$($eachItem['ConsentURL'])")
        write-host "Emailing $($userfield.User.Email)"
        $emailSent = SendEmail -Subject $emailSubject -Message $emailBody -SentBy "AccessLetter@bain.com" -SentTo $userfield.User.Email

        SendEmail `
            -Subject "Thank You. Request for Project $($prjName) registered" `
            -Message "Thank you for accepting Bain & Company's terms of access. Your information is being processed and you should receive the report from the project manager shortly.Please remember that the terms of access only allow you to share the report internally within your organization and its affiliates so if you want to pass the Materials to someone in another organization, they will also need to accept Bain & Company's terms of access before they can access the Materials." `
            -SentBy "AccessLetter@bain.com" -SentTo $vEmail

        if ($emailSent -eq $true) {
            $eachItem['ManagerInformed'] = $true
            $eachItem.Update()
        }
        $listMetadata = @{
            __metadata = @{'type' = 'SP.Data.AccessletterDatabaseListItem' };
            Consent    = $eachItem['Consent']
            ConsentOn  = $eachItem['ConsentOn']
            SName      = $vName
            SEmail     = $vEmail
            SCompany   = $vCompany
        } | ConvertTo-Json
        try {
            $response = PutRequest -endpoint "_api/web/lists/getbytitle('AccessletterDatabase')/GetItemById($($eachItem['Title']))" -body $listMetadata
            $alSite.Lists["Consent"].GetItemById($eachItem["ID"]).Delete()
            write-host "AccessletterDatabase list updated and Original Entry deleted"
        }
        catch {
            write-host "$($_.Exception)"
            continue
        }
    }
}
function EmailVendor() {

    $spQuery = New-Object Microsoft.SharePoint.SPQuery;
    $spQuery.Query = "<Where><Eq><FieldRef Name='EmailSent' /><Value Type='Boolean'>No</Value></Eq></Where>";
    $qualifiedItems = $alSite.Lists["Consent"].GetItems($spQuery);
    write-host "Notifying $($qualifiedItems.Count) Third Party Vendors"
    foreach ($eachItem in $qualifiedItems) {
        $userfield = New-Object Microsoft.SharePoint.SPFieldUserValue($alSite, $eachItem['Author'].ToString());
        if ($eachItem['ProjectName'] -ne $null) {
            $emailSubject = "$($eachItem['ProjectName']) - Access to Bain Materials"
        }
        else {
            $emailSubject = "Access to Bain Materials"
        }
        $emailBody = Get-Content -Path 'E:\Install\AccessLetter\EmailTemplate-Party.html'
        $emailBody = $emailBody.Replace("[BAINCONTACT]", $userfield.LookupValue.Split(".")[0].trim())
        $emailBody = $emailBody.Replace("[LINK]", "$($siteURL)/res/Consent.aspx?identity=$($eachItem['ConsentURL'])")
        write-host "-Emailing $($eachItem['Title'])"
        $emailSent = SendEmail -Subject $emailSubject -Message $emailBody -SentBy $userfield.LookupValue -SentTo $eachItem['Title']
        if ($emailSent -eq $false) {
            return
        }
        $ItemID = $eachItem['ID'].ToString()
        $Title = $eachItem['Title']
        $CaseCode = $eachItem['CaseCode']
        $ClientName = $eachItem['ClientName']
        $ProjectName = $eachItem['ProjectName']
        $RCompany = $eachItem['RCompany']
        $RName = $eachItem['RName']
        $VendorDiligence = $eachItem['VendorDiligence']
        $Consent = $eachItem['Consent']
        $ConsentOn = $eachItem['ConsentOn']
        $SName = $eachItem['SName']
        $SCompany = $eachItem['SCompany']
        $SEmail = $eachItem['SEmail']
        $EmailSent = $eachItem['EmailSent']
        $ConsentURL = $eachItem['ConsentURL']
        $managerEmail = $eachItem["Author"].split("#")[1].ToLower()

        $archiveID = CreateArchiveEntry -ItemID $ItemID -Title $Title -CaseCode $CaseCode -ClientName $ClientName -ProjectName $ProjectName -RCompany $RCompany `
            -VendorDiligence $VendorDiligence -Consent $Consent -ConsentOn $ConsentOn -SName $SName -SCompany $SCompany `
            -SEmail $SEmail -RName $RName -managerEmail $managerEmail

        $eachItem['Title'] = $archiveID
        $eachItem['CaseCode'] = ""
        $eachItem['ClientName'] = ""
        $eachItem['ProjectName'] = ""
        $eachItem['RCompany'] = ""
        $eachItem['RName'] = ""
        $eachItem['VendorDiligence'] = ""
        $eachItem['EmailSent'] = $true
        $eachItem.Update()
    }
}
function Remove-Diacritics {
    param ([String]$src = [String]::Empty)
    $normalized = $src.Normalize( [Text.NormalizationForm]::FormD )
    $sb = new-object Text.StringBuilder
    $normalized.ToCharArray() | % {
        if ( [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($_)
        }
    }
    $sb.ToString()
}
function CreateArchiveEntry($ItemID, $Title, $CaseCode, $ClientName, $ProjectName, $RCompany, $VendorDiligence, $Consent, $ConsentOn, $SName, $SCompany, $SEmail, $RName, $managerEmail) {
    $headers = @{accept = "application/json; odata=verbose"}
    $url = $siteURLCom
    $cred = New-Object System.Management.Automation.PSCredential "Intranet_prdinstall", (Get-Content "E:\Install\AccessLetter\passkey.txt" | ConvertTo-SecureString)
    try {
        SetFormDigest
    }
    catch {
        write-host "Unable to establish connection function()storeToMainList $($_.Exception)"
        continue
    }
    $res = GetRequest("_api/web/lists/getbytitle('AccessletterDatabase')/Items?`$select=ID&`$filter=Title eq '$($Title)'", $null)
    if ($res.d.results.Count -eq 0) {
        $listMetadata = @{
            __metadata      = @{'type' = 'SP.Data.AccessletterDatabaseListItem' }
            Title           = Remove-Diacritics($Title)
            CaseCode        = $CaseCode
            ClientName      = $ClientName
            ProjectName     = $ProjectName
            RCompany        = Remove-Diacritics($RCompany)
            VendorDiligence = $VendorDiligence
            Consent         = $Consent
            ConsentOn       = $ConsentOn
            SName           = Remove-Diacritics($SName)
            SCompany        = Remove-Diacritics($SCompany)
            SEmail          = Remove-Diacritics($SEmail)
            RName           = Remove-Diacritics($RName)
            ManagerEmail    = $managerEmail
            PreviousID      = $ItemID
        } | ConvertTo-Json
        try {
            $response = PostRequest -endpoint "_api/web/lists/getbytitle('AccessletterDatabase')/Items" -body $listMetadata
            $datacorrected = $response -creplace '"Id":', '"Idnew":'
            $resultsT = ConvertFrom-Json -InputObject $datacorrected
            write-host "AccessletterDatabase new entry created"
            return $resultsT.d.Idnew
        }
        catch {
            write-host "$($_.Exception.Message) - function CreateArchiveEntry(1)"
            continue
        }
    }
}
function UpdateArchiveEntry {
    $VendorComebackRequestsItems = $alSite.Lists["VendorComebackRequests"].Items
    if ($VendorComebackRequestsItems.Count -gt 0) {
        $headers = @{accept = "application/json; odata=verbose"}
        $url = $siteURLCom
        $cred = New-Object System.Management.Automation.PSCredential "Intranet_prdinstall", (Get-Content "E:\Install\AccessLetter\passkey.txt" | ConvertTo-SecureString)
        try {
            SetFormDigest
        }
        catch {
            write-host "Unable to establish connection function()storeToMainList`n$($_.Exception)"
            continue
        }
    }
    foreach ($eachItem in $VendorComebackRequestsItems) {
        $current_PreviousID = $eachItem["PreviousID"]
        $current_SName = Remove-Diacritics($eachItem["SName"])
        $current_SEmail = Remove-Diacritics($eachItem["SEmail"])
        $current_SCompany = Remove-Diacritics($eachItem["SCompany"])
        $current_ConsentOn = $eachItem["ConsentOn"]

        $res = GetRequest("_api/web/lists/getbytitle('AccessletterDatabase')/Items?`$select=ID,ManagerEmail,SName,ProjectName&`$filter=PreviousID eq '$($current_PreviousID)'", $null)
        if ($res.Length -gt 0) {
            $datacorrected = $res -creplace '"Id":', '"Idnew":'
            $resultsT = ConvertFrom-Json -InputObject $datacorrected
            $editThis = $resultsT.d.results.Idnew
            $managerEmail = $resultsT.d.results.ManagerEmail
            $prjName = $resultsT.d.results.ProjectName
            $vName = $resultsT.d.results.SName
            if ($prjName -ne $null) {
                $emailSubject = "$($resultsT.d.results.ProjectName) $($vName)- Access to Bain Materials"
            }
            else {
                $emailSubject = "$($vName) Access to Bain Materials"
            }
            $emailBody = Get-Content -Path 'E:\Install\AccessLetter\EmailTemplate-Manager.html'
            $emailBody = $emailBody.Replace("[BAINCONTACT]", $managerEmail.split(".")[0].trim())
            $emailBody = $emailBody.Replace("[INDIVIDUALNAME]", $current_SName)
            $emailBody = $emailBody.Replace("[EMAILADDRESS]", $current_SEmail)
            $emailBody = $emailBody.Replace("[COMPANYNAME]", $current_SCompany)
            $emailBody = $emailBody.Replace("[ACCESSLINK]", "$($siteURL)/res/Consent.aspx")
            write-host "Emailing $($managerEmail)"
            $emailSent = SendEmail -Subject $emailSubject -Message $emailBody -SentBy "AccessLetter@bain.com" -SentTo $managerEmail

            SendEmail `
                -Subject "Thank You. Request for Project $($prjName) registered" `
                -Message "Thank you for accepting Bain & Company's terms of access. Your information is being processed and you should receive the report from the project manager shortly.Please remember that the terms of access only allow you to share the report internally within your organization and its affiliates so if you want to pass the Materials to someone in another organization, they will also need to accept Bain & Company's terms of access before they can access the Materials." `
                -SentBy "AccessLetter@bain.com" -SentTo $current_SEmail

            $listMetadata = @{
                __metadata = @{'type' = 'SP.Data.AccessletterDatabaseListItem' };
                ConsentOn  = $current_ConsentOn
                SName      = $current_SName
                SEmail     = $current_SEmail
                SCompany   = $current_SCompany
            } | ConvertTo-Json
            try {
                $response = PutRequest -endpoint "_api/web/lists/getbytitle('AccessletterDatabase')/GetItemById($($editThis))" -body $listMetadata
                $alSite.Lists["VendorComebackRequests"].GetItemById($eachItem["ID"]).Delete()
                write-host "AccessletterDatabase list updated and Original Entry deleted"
            }
            catch {
                write-host "$($_.Exception)"
                continue
            }
        }
    }
}
EmailVendor
EmailManager
UpdateArchiveEntry