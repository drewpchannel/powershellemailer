Function Select-FolderDialog
{
  param([string]$Description="Select Folder",[string]$RootFolder="Desktop")

  [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
     Out-Null     

  $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
  $objForm.Rootfolder = $RootFolder
  $objForm.Description = $Description
  $Show = $objForm.ShowDialog()
  If ($Show -eq "OK")
    {
      Return $objForm.SelectedPath
    }
  Else {
      Write-Error "Operation cancelled by user."
    }
}
$fileListPath = Select-FolderDialog
$fileList = Get-ChildItem $fileListPath
$usersNotFound = New-Object System.Collections.Generic.List[System.Object]
$i = 0

function getUserEmail 
{
  param ($loginName)
  $checkUser = Get-ADUser -f "sAMAccountName -eq '$loginName'"
  if ($checkUser -ne $null)
  {
    $userToEmail = Get-ADUser -Identity $loginName -Properties *
    return $userToEmail.EmailAddress.ToLower()
  } else {
    $usersNotFound.Add("$loginName $i `n")
  }
}

function emailUsers
{
  param
  (
    [string]$emailAddressToSend,
    [string]$attachmentPath
  )
  $outlook = new-object -comobject outlook.application
  $email = $outlook.CreateItem(0)
  $email.To = $emailAddressToSend
  $email.Subject = "Phone and headset sign out forms"
  $email.Body = "
  Hello, `n I'm sending out this email to get all the headsets and phones signed out.  If you see If you did not receive this equipment please contact GSSSI.Support@gsssi.org.  If everything next to your item is blank.  Example, YeaLink Phone is listed but everything else is blank.  This shows you did not receive a phone and can be ignored if you did not request one.  If everything next to headset is blank please fill in the approximate day received if you got a headset.    `n If everything looks correct, please sign in the Employee Signature field and send the signed form back to GSSSI.Support@gsssi.org.  Instructions on how to digitally sign Adobe documents have also been attached to this email.Thanks, Drew Poulin
" 
  $email.Attachments.add($attachmentPath)
  $email.Attachments.add('C:\How to Do a Digital Signature in Adobe Acrobat Reader DC (Updated).docx')
  $email.Send()
  $outlook.Quit()
}

foreach ($file in $fileList) 
{
  $i += 1
  $indUserString = $file.ToString()
  $parseIndUser = $indUserString -split ' '
  $firstInit = $parseIndUser[1].Substring(0,1).ToLower()
  $lastName = $parseIndUser[0].ToLower()
  $loginName = $firstInit+$lastName
  $filename = $fileListPath+'\'+$indUserString
  $usersEmailAddress = getUserEmail($loginName)

  <# 
  send out email here 
  replace if statement with emailUsers -emailAddressToSend $usersEmailAddress -attachmentPath $filename 
  #>
  if ($usersEmailAddress -eq 'john.denoto-roy@gsssi.org') {
    Write-Host 'Sending email----------------------'
    Write-Host $filename
    Write-Host 'Trying to attach' $indUserString
    emailUsers -emailAddressToSend $usersEmailAddress -attachmentPath $filename 
  }
}

Write-Host "Users that were not found (login name) (file number): `n" $usersNotFound