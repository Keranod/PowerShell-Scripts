# Create new -com object to control Outlook if opened
$outlook = new-object -com Outlook.Application -ea 1
$contactsObject = $outlook.session.GetDefaultFolder(10)

# Login, password and path to access csv with contacts
$user = ''
$password = ''
$csvDirectory = ''

# Get all properties from new item in contacts to have a list of them
function Get-ContactProperties 
{
    $newcontact = $contactsObject.Items.Add()
    $Props = $newcontact | gm -MemberType property | ?{$_.definition -like 'string*{set}*'}
    $newcontact.Delete()
    $Props | ForEach-Object {$_.Name}
}
$properties = Get-ContactProperties

#Variable to tell if net used
$NetUsed = $false

#net uses network location $csvDirectory if not already accessible
if([System.IO.Directory]::Exists($csvDirectory) -eq $false){
    $Dummy = net use $csvDirectory /user:$User $Password
    $NetUsed = $true
}

# Import latest contact list from H:\mail
$latestContactsList = Import-Csv -Path "$csvDirectory\ContactList.csv"

# Get modify time for contactlist
$contactsListModifyTime = (Get-ItemProperty -Path "$csvDirectory\ContactList.csv").LastWriteTime

# Check if 'R & D' contact exists, if not prompt for contacts update, if yes check if CustomCreationTime is different than contactslistmodifytime, is yes prompt for update
$rndContact = $contactsObject.Items.Find("[Email1Address] = 'rnd@'")
if (!$rndContact) {
    $rndContact = $contactsObject.Items.Find("[FullName] = 'R & D'")
}

if ($rndContact) {
    $rndCustomContactCreationTime = $rndContact.UserProperties['CustomCreationTime'].Value
}

if ($rndCustomContactCreationTime -ge $contactsListModifyTime) {exit}

$wshell = New-Object -ComObject Wscript.Shell
$answer = $wshell.Popup("Do you want to update your Outlook contact list?",0,"Question",32+4)

if ($answer -ne 6) {exit}


# For each contact from contact list check if exists in outlook contacts by checking for email address, if not add it
foreach ($contact in $latestContactsList) {
    $fullName = $contact.name
    $email = $contact.'E-mail Address'
    $jobTitle = $contact.'Job Title'
    $mobile = $contact.'Mobile Phone'

    # Check if contact with specific emailaddress exists and if yes skip it
    $contactWithEmailAlreadyExists = $contactsObject.Items.Find("[Email1Address] = '$email'")
    if ($contactWithEmailAlreadyExists) {continue}

    # Check if contact with specific Full Name exists and if yes skip it
    $contactWithFullNameAlreadyExists = $contactsObject.Items.Find("[FullName] = ""$fullName""")
    if ($contactWithFullNameAlreadyExists) {continue}

    # Create new contactObject, modify properties and save to current outlook session contacts list
    $newContactObject = $contactsObject.Items.Add()
    $newContactObject.TaskSubject = $fullName
    $newContactObject.MobileTelephoneNumber = [string]$mobile
    $newContactObject.JobTitle = $jobTitle
    $newContactObject.FullName = $fullName
    $newContactObject.Email1Address = $email
    $newContactObject.Email1DisplayName = $email
    $newContactObject.Subject = $fullName

    $newContactObject.Save()
}

# Get 'R & D' contact and change/create CustomCreationTime to contacts modify date so it is a reference is contact list needs updating
$rndContact = $contactsObject.Items.Find("[Email1Address] = 'rnd@'")
if (!$rndContact) {
    $rndContact = $contactsObject.Items.Find("[FullName] = 'R & D'")
}

if (!$rndContact) {
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Updating contact list failed",0,"Information",48+4)
    exit
}

$rndContact.UserProperties.Add('CustomCreationTime', [Microsoft.Office.Interop.Outlook.OlUserPropertyType]::olDateTime, $true, $false) | Out-Null
$rndContact.UserProperties['CustomCreationTime'].Value = $contactsListModifyTime
$rndContact.Save()

#Deletes net use
    if($NetUsed -eq $true){
        $Dummy = net use $csvDirectory /delete
}

#Clears variables and modules
Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *
