#VSTS Access Info
#Insert PAT
$pat = Get-Content ""
$encodedPat = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$pat"))

#Smtp Info
#Insert Server Info
$smtpServer = ""
#Insert Domain Info
$domain = ""
#Insert Username
$smtpUsername = ""
#Insert Password
$smtpPassword = ""

#Access Json Settings
#Insert Json File Path
$path = ""
#Insert Json File Name
$file = ""

try{
    $settingsFile = $path + "\" + $file;
    $settings = Get-Content -Raw -Path $settingsFile | Convertfrom-Json;
}catch{
    log $_ red
    exit;
}

function GetAdmins($adminInfo)
{
    while($adminInfo -ne $null)
    {
        if($adminInfo.ProjectName -notin $allProjects)
        {
            $allAdmins = @()

            for($i = 1; $i -le $adminInfo.Count; $i++)
            {
                if($adminInfo[$i - 1].ProjectName -eq $adminInfo[$i].ProjectName)
                {
                    $properties = @{
                        'ProjectName'= $adminInfo[$i - 1].ProjectName
                        'Email'= $adminInfo[$i - 1].Email}
                    $projectAdmin = New-Object –TypeName PSObject -Prop $properties
                    $allAdmins += $projectAdmin
                }
                else
                {
                    $properties = @{
                        'ProjectName'= $adminInfo[$i - 1].ProjectName
                        'Email'= $adminInfo[$i - 1].Email}
                    $projectAdmin = New-Object –TypeName PSObject -Prop $properties
                    $allAdmins += $projectAdmin
                    break
                }
            }

            $global:allProjects += $allAdmins[0].ProjectName
            $adminInfo = $adminInfo | Where-Object { $allProjects -notcontains $_.ProjectName }
            SendMail $allAdmins
        }
    }
}

function SendMail($adminList)
{    
   $msgPath = $settings.msgPath
   $subject = $settings.msgSubject
   $msgSubject = $subject.Replace('[projectname]',$adminList[0].ProjectName)
   $msgBody = (Get-Content $msgPath).Replace('[projectname]',$adminList[0].ProjectName)
    
   # Create e-mail message
   $msg = new-object Net.Mail.MailMessage

   #Set properties
   $msg.Subject = $msgSubject.ToString()

   # Set e-mail body
   $msg.IsBodyHtml = $True
   $msg.body = $msgBody

   $attachment = New-Object System.Net.Mail.Attachment –ArgumentList "C:\Users\ledennin\Pictures\Repojsonexample.PNG"
   $attachment.ContentDisposition.Inline = $True
   $attachment.ContentDisposition.DispositionType = "Inline"
   $attachment.ContentType.MediaType = "image/png"    
   $attachment.ContentId = 'image1.png'
   $msg.Attachments.Add($attachment)

   # SMTP server setup
   $smtpClient = New-Object Net.Mail.SmtpClient($smtpServer, 25)  
   $smtpClient.UseDefaultCredentials = "false" 
   $smtpClient.Credentials = New-Object System.Net.NetworkCredential($smtpUsername,$smtpPassword,"$domain")
   $smtpClient.EnableSsl = "true"
   $smtpClient.Timeout = 300000

   # Email structure 
   $msg.From = $settings.msgFrom
   $msg.ReplyTo = $settings.msgFrom

   foreach($email in $adminList.Email)
   {
        $msg.To.Add($email)
   }
    
   $msgCC = @($settings.msgCC)
   foreach($email in $msgCC)
   {
       $msg.CC.Add($email)
   }

   # Sending email 
   $smtpClient.Send($msg) 
}

#Import data from Excel
$fullDataSet = Import-Excel $settings.prFile
$allProjects = @()
GetAdmins $fullDataSet