    #Custom variables For Signature Management
    $SignatureName = 'New Mails Signature New Year'
    $SignatureNameReply = 'Reply Signature'
    $SignatureVer = '1.0'
    $UseSignOnNew = '1'        #If set to '0', the signature will be added as signature for new mails.
    $UseSignOnReply = '1'      #If set to '0', the signature will be added as signature for reply mails.
    $ForceSignatureNew = '1'   #If set to '0', the signature will be editable in Outlook and if set to '1' will be non-editable and forced - forced also as reply.
    $ForceSignatureReply = '1' #If set to '0', the signature will be editable in Outlook and if set to '1' will be non-editable and forced.




    #################################################################################################################################
    #################################################################################################################################
    ###############################################   Music On Hold Check    ########################################################
    #################################################################################################################################
    #################################################################################################################################
    #################################################################################################################################
    if (Test-Path 'c:\Bumper_Tag.wma'){}
    else{
        Invoke-WebRequest -Uri 'http://airbelgium.com/MusicOnHold/Bumper_Tag.wma' -OutFile 'c:\Bumper_Tag.wma'
    }


    #################################################################################################################################
    #################################################################################################################################
    ###############################################   Signature configuration   #####################################################
    #################################################################################################################################
    #################################################################################################################################
    #################################################################################################################################


    #Environment variables
    $wdTypes = Add-Type -AssemblyName 'Microsoft.Office.Interop.Word' -Passthru
    $wdSaveFormat = $wdTypes | Where {$_.Name -eq "wdSaveFormat"}
    $NeedIt = 0

    $AppData=(Get-Item env:appdata).value
    $SigPath = '\Microsoft\Signatures'
    $LocalSignaturePath = $AppData+$SigPath
    $RemoteSignaturePathFull = $SigSource
    $UserName = $env:username

    #Retriving implementations of Signatures from registry and check if we just need to change signature Name
    $SignatureInstVer = Get-ItemProperty -Name 'VersionSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue

    If (Get-ItemProperty -Name 'VersionSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue)
    {
        $SignatureInstVer = Get-ItemProperty -Name 'VersionSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue
        $SignatureInstVer = $SignatureInstVer.VersionSignature.ToString()
         If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue)
        {
            $SignatureInstName = Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue
            $SignatureInstName = $SignatureInstName.NewSignature.ToString()
                If ($SignatureInstName -ne $SignatureName)
                {
                    New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force 
                }
        }
         If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue)
        {
            $SignatureInstName = Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue
            $SignatureInstName = $SignatureInstName.NewSignature.ToString()
                If ($SignatureInstName -ne $SignatureNameReply)
                {
                    New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureNameReply -PropertyType 'String' -Force 
                }
        }
    }
    Else { 
        $SignatureInstVer = '0'
    } 

    #Using data to determinate if change of signature is needed
    If ($SignatureVer -ne $SignatureInstVer)
    {
        $NeedIt = 1
    }

    #Implementation Of Signature
    If ($NeedIt -gt 0)
    {
        #Check signature path (needs to be created if a signature has never been created for the profile
        if (!(Test-Path -path $LocalSignaturePath)) {
               New-Item $LocalSignaturePath -Type Directory
        }
        $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Hello, Your email signature is going to be updated. Please answer the next 4 questions. Please be careful, if you do a mistake, you will need to call support!",0,"Warning",0x1)

        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $strName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your full name (ex: Albert Einstein)","Enter Full Name")
        $strTitle = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your Job Title :", "Enter your Job Title")
        $strPhone = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your mobile phone number (ex: +32 472 11 11 11) :", "Enter mobile Number")
        $strPhone2 = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your business phone number (ex: +32 10 23 45 XX) :", "Enter Business Phone number")
    
        $stream = [System.IO.StreamWriter] "$LocalSignaturePath\\New Mails Signature New Year.htm"
        $stream.WriteLine("<!DOCTYPE HTML PUBLIC `"-//W3C//DTD HTML 4.0 Transitional//EN`">")
        $stream.WriteLine("<HTML><HEAD><TITLE>Signature</TITLE>")
        $stream.WriteLine("<META http-equiv=Content-Type content=`"text/html; charset=windows-1252`">")
        $stream.WriteLine("<BODY>")
        $stream.WriteLine('<table style="font-family:Arial;border-spacing:0px;color:#000000;font-size:12px; border-collapse:collapse" cellspacing="0" cellpadding="0" border="0"><tr><td style=" padding:0cm 0cm 0cm 0cm;" width="631"><b>')
        $stream.WriteLine($strName + "</b><br />" + $strTitle)
        $stream.WriteLine('</td></tr><tr><td style="padding:7pt 5pt 0cm 0cm;"><a href="http://www.airbelgium.com/" target="_blank"><img src="http://airbelgium.com/signatures/signatureLogo.png" id="x_Picture_x0020_53" style="width: 2.5052in; height: 0.7864in;" width="241" border="0" height="76"></a></td></tr><tr><td style="padding:7pt 5pt 0cm 0cm;"><b>M.</b>&nbsp;')
        $stream.WriteLine($strPhone + '&nbsp;|&nbsp;<b>T.</b>&nbsp;' + $strPhone2)
        $stream.WriteLine('</td></tr><tr><td style="padding:7pt 5pt 0cm 0cm;"><b><span style="line-height:115%; color:#404040" lang="EN-US">AIR BELGIUM S.A.</span></b><br />Rue Emile Francqui, 7 | 1435 Mont-Saint-Guibert<br />Belgium</td></tr><tr style="padding-bottom: 5cm;"><td style="padding:7pt 5pt 0cm 0cm;">')
        $stream.WriteLine('<a href="http://www.facebook.com/airbelgium/" target="_blank"><img src="http://airbelgium.com/signatures/signatureFacebook.png" alt="Facebook" style="width: 0.1979in; height: 0.1979in;" width="19" border="0" height="19"></a>&nbsp;<a href="http://twitter.com/airbelgium_off" target="_blank"><img src="http://airbelgium.com/signatures/signatureTwitter.png" alt="Twitter" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a>&nbsp;<a href="https://www.linkedin.com/company/10647555?trk=tyah&amp;trkInfo=tarId:1472236574956,tas:air%20belgium,idx:2-1-2" target="_blank"><img src="http://airbelgium.com/signatures/signatureLinkedIn.png" alt="LinkedIn" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a>&nbsp;<a href="http://www.youtube.com/" target="_blank"><img src="http://airbelgium.com/signatures/signatureYoutube.png" alt="YouTube" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a>&nbsp;<a href="https://www.instagram.com/airbelgium_official/" target="_blank"><img src="http://airbelgium.com/signatures/ABInstagramSignature.png" alt="Instagram" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a><a href="http://www.airbelgium.com/" target="_blank"><img src="http://airbelgium.com/signatures/signatureWebsite.png" alt="Website" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a></p></td></tr>')
        $stream.WriteLine('</td></tr><tr><td><img src="http://airbelgium.com/signatures/SignatureNoel.jpg" alt="Merry XMas !" border="0" width="550" height="175" /></td></tr>')
        $stream.WriteLine('<tr style="height:17.75pt;"><td style="padding:7pt 5pt 0cm 0cm; height:17.75pt" width="631" valign="bottom"><i>This message (including any attachments) contains confidential information intended for a specific individual and purpose, and is protected by law. If you are not the intended recipient, you should delete this message and are hereby notified that any disclosure, duplication, or distribution of this message, or the taking of any action based on it, is strictly prohibited.</i></td></tr></table>')
        $stream.WriteLine("</BODY>")
        $stream.WriteLine("</HTML>")
        $stream.close()

        $stream = [System.IO.StreamWriter] "$LocalSignaturePath\\New Mails Signature.htm"
        $stream.WriteLine("<!DOCTYPE HTML PUBLIC `"-//W3C//DTD HTML 4.0 Transitional//EN`">")
        $stream.WriteLine("<HTML><HEAD><TITLE>Signature</TITLE>")
        $stream.WriteLine("<META http-equiv=Content-Type content=`"text/html; charset=windows-1252`">")
        $stream.WriteLine("<BODY>")
        $stream.WriteLine('<table style="font-family:Arial;border-spacing:0px;color:#000000;font-size:12px; border-collapse:collapse" cellspacing="0" cellpadding="0" border="0"><tr><td style=" padding:0cm 0cm 0cm 0cm;" width="631"><b>')
        $stream.WriteLine($strName + "</b><br />" + $strTitle)
        $stream.WriteLine('</td></tr><tr><td style="padding:7pt 5pt 0cm 0cm;"><a href="http://www.airbelgium.com/" target="_blank"><img src="http://airbelgium.com/signatures/signatureLogo.png" id="x_Picture_x0020_53" style="width: 2.5052in; height: 0.7864in;" width="241" border="0" height="76"></a></td></tr><tr><td style="padding:7pt 5pt 0cm 0cm;"><b>M.</b>&nbsp;')
        $stream.WriteLine($strPhone + '&nbsp;|&nbsp;<b>T.</b>&nbsp;' + $strPhone2)
        $stream.WriteLine('</td></tr><tr><td style="padding:7pt 5pt 0cm 0cm;"><b><span style="line-height:115%; color:#404040" lang="EN-US">AIR BELGIUM S.A.</span></b><br />Rue Emile Francqui, 7 | 1435 Mont-Saint-Guibert<br />Belgium</td></tr><tr style="padding-bottom: 5cm;"><td style="padding:7pt 5pt 0cm 0cm;">')
        $stream.WriteLine('<a href="http://www.facebook.com/airbelgium/" target="_blank"><img src="http://airbelgium.com/signatures/signatureFacebook.png" alt="Facebook" style="width: 0.1979in; height: 0.1979in;" width="19" border="0" height="19"></a>&nbsp;<a href="http://twitter.com/airbelgium_off" target="_blank"><img src="http://airbelgium.com/signatures/signatureTwitter.png" alt="Twitter" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a>&nbsp;<a href="https://www.linkedin.com/company/10647555?trk=tyah&amp;trkInfo=tarId:1472236574956,tas:air%20belgium,idx:2-1-2" target="_blank"><img src="http://airbelgium.com/signatures/signatureLinkedIn.png" alt="LinkedIn" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a>&nbsp;<a href="http://www.youtube.com/" target="_blank"><img src="http://airbelgium.com/signatures/signatureYoutube.png" alt="YouTube" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a>&nbsp;<a href="https://www.instagram.com/airbelgium_official/" target="_blank"><img src="http://airbelgium.com/signatures/ABInstagramSignature.png" alt="Instagram" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a><a href="http://www.airbelgium.com/" target="_blank"><img src="http://airbelgium.com/signatures/signatureWebsite.png" alt="Website" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a></p></td></tr>')
        $stream.WriteLine('<tr style="height:17.75pt;"><td style="padding:7pt 5pt 0cm 0cm; height:17.75pt" width="631" valign="bottom"><i>This message (including any attachments) contains confidential information intended for a specific individual and purpose, and is protected by law. If you are not the intended recipient, you should delete this message and are hereby notified that any disclosure, duplication, or distribution of this message, or the taking of any action based on it, is strictly prohibited.</i></td></tr></table>')
        $stream.WriteLine("</BODY>")
        $stream.WriteLine("</HTML>")
        $stream.close()

        $stream = [System.IO.StreamWriter] "$LocalSignaturePath\\Reply Signature.htm"
        $stream.WriteLine("<!DOCTYPE HTML PUBLIC `"-//W3C//DTD HTML 4.0 Transitional//EN`">")
        $stream.WriteLine("<HTML><HEAD><TITLE>Signature</TITLE>")
        $stream.WriteLine("<META http-equiv=Content-Type content=`"text/html; charset=windows-1252`">")
        $stream.WriteLine("<BODY>")
        $stream.WriteLine('<table style="font-family:Arial;border-spacing:0px;color:#000000;font-size:12px; border-collapse:collapse" cellspacing="0" cellpadding="0" border="0"><tr><td style=" padding:0cm 0cm 0cm 0cm;" width="631"><b>')
        $stream.WriteLine($strName + "</b><br />" + $strTitle)
        $stream.WriteLine('</td></tr><tr><td style="padding:7pt 5pt 0cm 0cm;"><a href="http://www.airbelgium.com/" target="_blank"><img src="http://airbelgium.com/signatures/signatureLogo.png" id="x_Picture_x0020_53" style="width: 2.5052in; height: 0.7864in;" width="241" border="0" height="76"></a></td></tr>')
        $stream.WriteLine('<tr style="height:17.75pt;"><td style="padding:7pt 5pt 0cm 0cm; height:17.75pt" width="631" valign="bottom"><i>This message (including any attachments) contains confidential information intended for a specific individual and purpose, and is protected by law. If you are not the intended recipient, you should delete this message and are hereby notified that any disclosure, duplication, or distribution of this message, or the taking of any action based on it, is strictly prohibited.</i></td></tr></table>')
        $stream.WriteLine("</BODY>")
        $stream.WriteLine("</HTML>")
        $stream.close()
 
        If (Test-Path HKCU:'\Software\Microsoft\Office\16.0')
 
        {
            $Outlook = 'Outlook'
            if ($Outlook -ne $null)
            {
                Stop-Process -Name $Outlook -Force
            }
 
            #$MSWord = New-Object -comobject word.application
            #$EmailOptions = $MSWord.EmailOptions
            #$EmailSignature = $EmailOptions.EmailSignature
            #$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
            #If ($UseSignOnNew -eq '1')
            #{
            #    $EmailSignature.NewMessageSignature="$SignatureName"
            #}
            #If ($UseSignOnReply -eq '1')
            #{
            #    $EmailSignature.ReplyMessageSignature="$SignatureNameReply"
            #}
            Stop-Process -Name $Outlook

        If ($ForceSignatureNew -eq '1')
        {
            New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force 
            #If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
            #Else { 
            #    New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force 
            #} 
        }

        If ($ForceSignatureReply -eq '1')
        {
            #If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
            #Else { 
            New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureNameReply -PropertyType 'String' -Force
            #    } 
        }
        }

        #Write Signature specified Registry Values
        If (Get-ItemProperty -Name 'VersionSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue)
        {
            Set-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'VersionSignature' -Value $SignatureVer
        }
        Else { 
            New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'VersionSignature' -Value $SignatureVer -PropertyType 'String' 
        } 


        $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup('Script is completed! Mail signatures in Outlook have been updated',0,'All Done',0x0)
    }


















    #kill window
    Get-Process PowerShell | stop-process