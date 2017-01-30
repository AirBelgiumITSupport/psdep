    #general Script Version
    $generalScriptVersion = '1.1'

    #Specific Users Script Version
    $SpecificUserScriptVersion = '1.0'

    #Custom variables For Signature Management
    $SignatureName = 'AB New Mails Signature'
    $SignatureNameReply = 'AB Reply Signature'
    $SignatureVer = '1.7'
    $UseSignOnNew = '1'        #If set to '0', the signature will be added as signature for new mails.
    $UseSignOnReply = '1'      #If set to '0', the signature will be added as signature for reply mails.
    $ForceSignatureNew = '0'   #If set to '0', the signature will be editable in Outlook and if set to '1' will be non-editable and forced - forced also as reply.
    $ForceSignatureReply = '0' #If set to '0', the signature will be editable in Outlook and if set to '1' will be non-editable and forced.

    #Sites to be added as trusted
    $SubDomain = "lync.com", "outlook.com", "microsoftonline.com", "sharepoint.com", "airbusworld.com", "airbusdoc.com", "airbus.com", "airbus2.com"  #This will add: https://*.lync.com/ etc.
    

    #Environnment Variables
    $AppData=(Get-Item env:appdata).value
    $SigPath = '\Microsoft\Signatures'
    $LocalSignaturePath = $AppData+$SigPath
    $RemoteSignaturePathFull = $SigSource
    $UserName = $env:username

    #Empty microsoft Upload center cache
    #C:\Users\FabienDelhaye\AppData\Local\Microsoft\Office\Spw    
    del $AppData'\..\Local\Microsoft\Office\16.0\OfficeFileCache1\*' 2>null


    #################################################################################################################################
    #################################################################################################################################
    ###############################################   Specific Users    #############################################################
    #################################################################################################################################
    #################################################################################################################################
    #################################################################################################################################
    $needTOExecuteSpecificUserScript = 0
    ###############################################
    ######### Copy this bloc once per user#########
    ###############################################
    #######  Do not forget to update UserID #######
    ###############################################
    if ($UserName -eq 'fabien.delhaye'){
        #New-ItemProperty HKCU:'\Software\AB\ITScript' -Name 'ABScriptVersion' -Value $SpecificUserScriptVersion -PropertyType 'String' -Force 
        If (Get-ItemProperty -Name 'SpecificUserScriptVersion' -Path HKCU:'\Software\AB\ITScript' -ErrorAction SilentlyContinue) { 
            $SpecificUserScriptInstalledVersion = Get-ItemProperty -Name 'SpecificUserScriptVersion' -Path HKCU:'\Software\AB\ITScript'
            $SpecificUserScriptInstalledVersion = $SpecificUserScriptInstalledVersion.SpecificUserScriptVersion.ToString()

            if ($SpecificUserScriptInstalledVersion -eq $SpecificUserScriptVersion){}else{
                $needTOExecuteSpecificUserScript = 1
                New-ItemProperty HKCU:'\Software\AB\ITScript' -Name 'SpecificUserScriptVersion' -Value $SpecificUserScriptVersion -PropertyType 'String' -Force 
            }
        } 
        Else {
            New-Item -Path HKCU:'\Software\AB\ITScript' 
		    New-ItemProperty HKCU:'\Software\AB\ITScript' -Name 'SpecificUserScriptVersion' -Value $SpecificUserScriptVersion -PropertyType 'String' -Force 
            $needTOExecuteSpecificUserScript = 1
        }

        if($needTOExecuteSpecificUserScript -eq 1){
            #Do your admin stuff !
            #Do your admin stuff !
            #Do your admin stuff !
            #Do your admin stuff !


        }
    }
    ###############################################
    ################# END OF BLOC #################
    ###############################################
    

    #################################################################################################################################
    #################################################################################################################################
    ###############################################   Music On Hold Check    ########################################################
    #################################################################################################################################
    #################################################################################################################################
    #################################################################################################################################
    if (Test-Path 'c:\MusicOnHold\Bumper_Tag.wma'){}
    else{
        if (Test-Path 'c:\MusicOnHold\Bumper_Tag.wma'){}else{ New-Item 'c:\MusicOnHold' -ItemType Directory }
        Invoke-WebRequest -Uri 'http://airbelgium.com/MusicOnHold/Bumper_Tag.wma' -OutFile 'c:\MusicOnHold\Bumper_Tag.wma'
    }
	

    #################################################################################################################################
    #################################################################################################################################
    #############################################   Update default home Page    #####################################################
    #########################################  Pin and Upin application to TaskBar  #################################################
    ################################################  Adding Trusted Sites  #########################################################
    #################################################################################################################################
    #################################################################################################################################
	#New-ItemProperty HKCU:'\Software\AB\ITScript' -Name 'ABScriptVersion' -Value $SpecificUserScriptVersion -PropertyType 'String' -Force 
	If (Get-ItemProperty -Name 'generalScriptVersion' -Path HKCU:'\Software\AB\ITScript' -ErrorAction SilentlyContinue) { 
		$scriptInstalledVersion = Get-ItemProperty -Name 'generalScriptVersion' -Path HKCU:'\Software\AB\ITScript'
		$scriptInstalledVersion = $scriptInstalledVersion.generalScriptVersion.ToString()

		if ($scriptInstalledVersion -eq $generalScriptVersion){
            $needTOExecuteScript = 0
        }else{
			$needTOExecuteScript = 1
			New-ItemProperty HKCU:'\Software\AB\ITScript' -Name 'generalScriptVersion' -Value $generalScriptVersion -PropertyType 'String' -Force 
		}
	} 
	Else {
		New-Item -Path HKCU:'\Software\AB\ITScript'
		New-ItemProperty HKCU:'\Software\AB\ITScript' -Name 'generalScriptVersion' -Value $generalScriptVersion -PropertyType 'String' -Force 
		$needTOExecuteScript = 1
	}


	if($needTOExecuteScript -eq 1){	

        function Pin-App ([string]$appname, [switch]$unpin, [switch]$start, [switch]$taskbar, [string]$path) {
            if ($unpin.IsPresent) {
                $action = "Unpin"
            } else {
                $action = "Pin"
            }
    
            if (-not $taskbar.IsPresent -and -not $start.IsPresent) {
                Write-Error "Specify -taskbar and/or -start!"
            }
    
            if ($taskbar.IsPresent) {
                try {
                    $exec = $false
                    if ($action -eq "Unpin") {
                        ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'Unpin from taskbar'} | %{$_.DoIt(); $exec = $true}
                        if ($exec) {
                            Write "App '$appname' unpinned from Taskbar"
                        } else {
                            if (-not $path -eq "") {
                                Pin-App-by-Path $path -Action $action
                            } else {
                                Write "'$appname' not found or 'Unpin from taskbar' not found on item!"
                            }
                        }
                    } else {
                        ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'Pin to taskbar'} | %{$_.DoIt(); $exec = $true}
                
                        if ($exec) {
                            Write "App '$appname' pinned to Taskbar"
                        } else {
                            if (-not $path -eq "") {
                                Pin-App-by-Path $path -Action $action
                            } else {
                                Write "'$appname' not found or 'Pin to taskbar' not found on item!"
                            }
                        }
                    }
                } catch {
                    Write-Error "Error Pinning/Unpinning $appname to/from taskbar!"
                }
            }
    
            if ($start.IsPresent) {
                try {
                    $exec = $false
                    if ($action -eq "Unpin") {
                        ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'Unpin from Start'} | %{$_.DoIt(); $exec = $true}
                
                        if ($exec) {
                            Write "App '$appname' unpinned from Start"
                        } else {
                            if (-not $path -eq "") {
                                Pin-App-by-Path $path -Action $action -start
                            } else {
                                Write "'$appname' not found or 'Unpin from Start' not found on item!"
                            }
                        }
                    } else {
                        ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'Pin to Start'} | %{$_.DoIt(); $exec = $true}
                
                        if ($exec) {
                            Write "App '$appname' pinned to Start"
                        } else {
                            if (-not $path -eq "") {
                                Pin-App-by-Path $path -Action $action -start
                            } else {
                                Write "'$appname' not found or 'Pin to Start' not found on item!"
                            }
                        }
                    }
                } catch {
                    Write-Error "Error Pinning/Unpinning $appname to/from Start!"
                }
            }
        }

        function Pin-App-by-Path([string]$Path, [string]$Action, [switch]$start) {
            if ($Path -eq "") {
                Write-Error -Message "You need to specify a Path" -ErrorAction Stop
            }
            if ($Action -eq "") {
                Write-Error -Message "You need to specify an action: Pin or Unpin" -ErrorAction Stop
            }
            if ((Get-Item -Path $Path -ErrorAction SilentlyContinue) -eq $null){
                Write-Error -Message "$Path not found" -ErrorAction Stop
            }
            $Shell = New-Object -ComObject "Shell.Application"
            $ItemParent = Split-Path -Path $Path -Parent
            $ItemLeaf = Split-Path -Path $Path -Leaf
            $Folder = $Shell.NameSpace($ItemParent)
            $ItemObject = $Folder.ParseName($ItemLeaf)
            $Verbs = $ItemObject.Verbs()
    
            if ($start.IsPresent) {
                switch($Action){
                    "Pin"   {$Verb = $Verbs | Where-Object -Property Name -EQ "&Pin to Start"}
                    "Unpin" {$Verb = $Verbs | Where-Object -Property Name -EQ "Un&pin from Start"}
                    default {Write-Error -Message "Invalid action, should be Pin or Unpin" -ErrorAction Stop}
                }
            } else {
                switch($Action){
                    "Pin"   {$Verb = $Verbs | Where-Object -Property Name -EQ "Pin to Tas&kbar"}
                    "Unpin" {$Verb = $Verbs | Where-Object -Property Name -EQ "Unpin from Tas&kbar"}
                    default {Write-Error -Message "Invalid action, should be Pin or Unpin" -ErrorAction Stop}
                }
            }
    
            if($Verb -eq $null){
                Write-Error -Message "That action is not currently available on this Path" -ErrorAction Stop
            } else {
                $Result = $Verb.DoIt()
            }
        }


        Set-ItemProperty -Path HKCU:'\Software\Microsoft\Internet Explorer\Main\' -Name 'start page' -Value 'https://airbelgium.sharepoint.com/SitePages/Accueil.aspx' -Force
        Pin-App "Microsoft Edge" -unpin -taskbar  
        Pin-App "Microsoft Edge" -unpin -start
        Pin-App "Skype Preview" -unpin -start
        Pin-App "Internet Explorer" -pin -taskbar  
        Pin-App "Internet Explorer" -pin -start                 
        Pin-App "Outlook 2016" -pin -start                 
        Pin-App "OneNote 2016" -pin -start   
        Pin-App "PowerPoint 2016" -pin -start
        Pin-App "Word 2016" -pin -start   
        Pin-App "Excel 2016" -pin -start                      
        Pin-App "Skype for Business 2016" -pin -start   
        Pin-App "Skype for Business 2016" -pin -taskbar   
        Pin-App "Store" -unpin -start
        Pin-App "Store" -unpin -taskbar



        #Initialize key variables
        $UserRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains"
        $DWord = 2
        #Main function
        If($TrustedSites)
        {
            #Adding trusted sites in the registry
            Foreach($TruestedSite in $TrustedSites)
            {
                #If user does not specify the user type. By default,the script will add the trusted sites for the current user.
                    New-Item -Path $UserRegPath\$TruestedSite -Force
                    New-Item -Path $UserRegPath\$TruestedSite\* -Force
                    New-ItemProperty $UserRegPath\$TruestedSite\* -Name 'https' -Value 2 -PropertyType 'DWORD' -Force
                    Write-Host "Successfully added '$TruestedSite' domain to trusted Sites in Internet Explorer."
            }
        }
        
        #dism.exe /Online /Export-DefaultAppAssociations:C:\AppAssoc.xml
        #dism.exe /online /Import-DefaultAppAssociations:C:\AppAssoc.xml

        #Set office Document Cache to max 1 day
        
        #Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\ftp\UserChoice' -name ProgId IE.FTP
        #Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice' -name ProgId IE.HTTP
        #Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice' -name ProgId IE.HTTPS

        #Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\ftp\UserChoice' -name ProgId FirefoxURL
        #Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice' -name ProgId FirefoxURL
        #Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice' -name ProgId FirefoxURL

        <#
            (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\ftp\UserChoice').ProgId
            (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice').ProgId
            (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice').ProgId
        #>



        <#
        #powershell
            $AppsList = "Microsoft.BingFinance","Microsoft.BingNews","Microsoft.BingWeather","Microsoft.XboxApp","Microsoft.MicrosoftSolitaireCollection","Microsoft.BingSports","Microsoft.ZuneMusic","Microsoft.ZuneVideo","Microsoft.Windows.Photos","Microsoft.People","Microsoft.MicrosoftOfficeHub","Microsoft.WindowsMaps","microsoft.windowscommunicationsapps","Microsoft.Getstarted","Microsoft.3DBuilder","Microsoft.Office.Sway"

            ForEach ($App in $AppsList) 
            { 
                $PackageFullName = (Get-AppxPackage $App).PackageFullName
                $ProPackageFullName = (Get-AppxProvisionedPackage -online | where {$_.Displayname -eq $App}).PackageName
                    write-host $PackageFullName
                    Write-Host $ProPackageFullName 
                if ($PackageFullName) 
                { 
                    Write-Host "Removing Package: $App"
                    remove-AppxPackage -package $PackageFullName 
                } 
                else 
                { 
                    Write-Host "Unable to find package: $App" 
                } 
                    if ($ProPackageFullName) 
                { 
                    Write-Host "Removing Provisioned Package: $ProPackageFullName"
                    Remove-AppxProvisionedPackage -online -packagename $ProPackageFullName 
                } 
                else 
                { 
                    Write-Host "Unable to find provisioned package: $App" 
                } 
            }

        #>


	 }





    #################################################################################################################################
    #################################################################################################################################
    ###############################################   Signature configuration   #####################################################
    #################################################################################################################################
    #################################################################################################################################
    #################################################################################################################################


    #Environment variables
    $NeedIt = 0


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
        $UserNameSignature = $userName.ToLower()
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/AirBelgiumITSupport/psdep/master/signatures/AB-Signature_$UserNameSignature.html" -OutFile "$LocalSignaturePath\\AB New Mails Signature.htm"
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/AirBelgiumITSupport/psdep/master/signatures/AB-SignatureReply_$UserNameSignature.html" -OutFile "$LocalSignaturePath\\AB Reply Signature.htm"

        #$wshell = New-Object -ComObject Wscript.Shell
        #$wshell.Popup("Hello, Your email signature is going to be updated. Please answer the next 4 questions. Please be careful, if you do a mistake, you will need to call support!",0,"Warning",0x1)

        #[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        #$strName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your full name, Please do not use special chars (with accent). ex: Albert Einstein","Enter Full Name")
        ##$strTitle = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your Job Title :", "Enter your Job Title")
        #$strPhone = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your mobile phone number (ex: +32 472 11 11 11) :", "Enter mobile Number")
        #$strPhone2 = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your business phone number (ex: +32 10 23 45 XX) :", "Enter Business Phone number")
   # 
   #     $stream = [System.IO.StreamWriter] "$LocalSignaturePath\\New Mails Signature New Year.htm"
        #$stream.WriteLine("<!DOCTYPE HTML PUBLIC `"-//W3C//DTD HTML 4.0 Transitional//EN`">")
        #$stream.WriteLine("<HTML><HEAD><TITLE>Signature</TITLE>")
        #$stream.WriteLine("<META http-equiv=Content-Type content=`"text/html; charset=windows-1252`">")
        #$stream.WriteLine("<BODY>")
        #$stream.WriteLine('<table style="font-family:Arial;border-spacing:0px;color:#000000;font-size:12px; border-collapse:collapse" cellspacing="0" cellpadding="0" border="0"><tr><td style=" padding:0cm 0cm 0cm 0cm;" width="631"><b>')
        #$stream.WriteLine($strName + "</b><br />" + $strTitle)
        #$stream.WriteLine('</td></tr><tr><td style="padding:7pt 5pt 0cm 0cm;"><a href="http://www.airbelgium.com/" target="_blank"><img src="http://airbelgium.com/signatures/signatureLogo.png" id="x_Picture_x0020_53" style="width: 2.5052in; height: 0.7864in;" width="241" border="0" height="76"></a></td></tr><tr><td style="padding:7pt 5pt 0cm 0cm;"><b>M.</b>&nbsp;')
        #$stream.WriteLine($strPhone + '&nbsp;|&nbsp;<b>T.</b>&nbsp;' + $strPhone2)
        #$stream.WriteLine('</td></tr><tr><td style="padding:7pt 5pt 0cm 0cm;"><b><span style="line-height:115%; color:#404040" lang="EN-US">AIR BELGIUM S.A.</span></b><br />Rue Emile Francqui, 7 | 1435 Mont-Saint-Guibert<br />Belgium</td></tr><tr style="padding-bottom: 5cm;"><td style="padding:7pt 5pt 0cm 0cm;">')
        #$stream.WriteLine('<a href="http://www.facebook.com/airbelgium/" target="_blank"><img src="http://airbelgium.com/signatures/signatureFacebook.png" alt="Facebook" style="width: 0.1979in; height: 0.1979in;" width="19" border="0" height="19"></a>&nbsp;<a href="http://twitter.com/airbelgium_off" target="_blank"><img src="http://airbelgium.com/signatures/signatureTwitter.png" alt="Twitter" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a>&nbsp;<a href="https://www.linkedin.com/company/10647555?trk=tyah&amp;trkInfo=tarId:1472236574956,tas:air%20belgium,idx:2-1-2" target="_blank"><img src="http://airbelgium.com/signatures/signatureLinkedIn.png" alt="LinkedIn" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a>&nbsp;<a href="http://www.youtube.com/" target="_blank"><img src="http://airbelgium.com/signatures/signatureYoutube.png" alt="YouTube" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a>&nbsp;<a href="https://www.instagram.com/airbelgium_official/" target="_blank"><img src="http://airbelgium.com/signatures/ABInstagramSignature.png" alt="Instagram" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a><a href="http://www.airbelgium.com/" target="_blank"><img src="http://airbelgium.com/signatures/signatureWebsite.png" alt="Website" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a></p></td></tr>')
        #$stream.WriteLine('</td></tr><tr><td><img src="http://airbelgium.com/signatures/SignatureNoel.jpg" alt="Merry XMas !" border="0" width="550" height="175" /></td></tr>')
        ##$stream.WriteLine('<tr style="height:17.75pt;"><td style="padding:7pt 5pt 0cm 0cm; height:17.75pt" width="631" valign="bottom"><i>This message (including any attachments) contains confidential information intended for a specific individual and purpose, and is protected by law. If you are not the intended recipient, you should delete this message and are hereby notified that any disclosure, duplication, or distribution of this message, or the taking of any action based on it, is strictly prohibited.</i></td></tr></table>')
        #$stream.WriteLine("</BODY>")
        #$stream.WriteLine("</HTML>")
        #$stream.close()

        #$stream = [System.IO.StreamWriter] "$LocalSignaturePath\\New Mails Signature.htm"
        ##$stream.WriteLine("<!DOCTYPE HTML PUBLIC `"-//W3C//DTD HTML 4.0 Transitional//EN`">")
        #$stream.WriteLine("<HTML><HEAD><TITLE>Signature</TITLE>")
        #$stream.WriteLine("<META http-equiv=Content-Type content=`"text/html; charset=windows-1252`">")
        #$stream.WriteLine("<BODY>")
        #$stream.WriteLine('<table style="font-family:Arial;border-spacing:0px;color:#000000;font-size:12px; border-collapse:collapse" cellspacing="0" cellpadding="0" border="0"><tr><td style=" padding:0cm 0cm 0cm 0cm;" width="631"><b>')
        #$stream.WriteLine($strName + "</b><br />" + $strTitle)
        #$stream.WriteLine('</td></tr><tr><td style="padding:7pt 5pt 0cm 0cm;"><a href="http://www.airbelgium.com/" target="_blank"><img src="http://airbelgium.com/signatures/signatureLogo.png" id="x_Picture_x0020_53" style="width: 2.5052in; height: 0.7864in;" width="241" border="0" height="76"></a></td></tr><tr><td style="padding:7pt 5pt 0cm 0cm;"><b>M.</b>&nbsp;')
        #$stream.WriteLine($strPhone + '&nbsp;|&nbsp;<b>T.</b>&nbsp;' + $strPhone2)
        #$stream.WriteLine('</td></tr><tr><td style="padding:7pt 5pt 0cm 0cm;"><b><span style="line-height:115%; color:#404040" lang="EN-US">AIR BELGIUM S.A.</span></b><br />Rue Emile Francqui, 7 | 1435 Mont-Saint-Guibert<br />Belgium</td></tr><tr style="padding-bottom: 5cm;"><td style="padding:7pt 5pt 0cm 0cm;">')
        #$stream.WriteLine('<a href="http://www.facebook.com/airbelgium/" target="_blank"><img src="http://airbelgium.com/signatures/signatureFacebook.png" alt="Facebook" style="width: 0.1979in; height: 0.1979in;" width="19" border="0" height="19"></a>&nbsp;<a href="http://twitter.com/airbelgium_off" target="_blank"><img src="http://airbelgium.com/signatures/signatureTwitter.png" alt="Twitter" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a>&nbsp;<a href="https://www.linkedin.com/company/10647555?trk=tyah&amp;trkInfo=tarId:1472236574956,tas:air%20belgium,idx:2-1-2" target="_blank"><img src="http://airbelgium.com/signatures/signatureLinkedIn.png" alt="LinkedIn" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a>&nbsp;<a href="http://www.youtube.com/" target="_blank"><img src="http://airbelgium.com/signatures/signatureYoutube.png" alt="YouTube" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a>&nbsp;<a href="https://www.instagram.com/airbelgium_official/" target="_blank"><img src="http://airbelgium.com/signatures/ABInstagramSignature.png" alt="Instagram" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a><a href="http://www.airbelgium.com/" target="_blank"><img src="http://airbelgium.com/signatures/signatureWebsite.png" alt="Website" style="width: 0.1927in; height: 0.1927in;" width="19" border="0" height="19"></a></p></td></tr>')
        #$stream.WriteLine('<tr style="height:17.75pt;"><td style="padding:7pt 5pt 0cm 0cm; height:17.75pt" width="631" valign="bottom"><i>This message (including any attachments) contains confidential information intended for a specific individual and purpose, and is protected by law. If you are not the intended recipient, you should delete this message and are hereby notified that any disclosure, duplication, or distribution of this message, or the taking of any action based on it, is strictly prohibited.</i></td></tr></table>')
        #$stream.WriteLine("</BODY>")
        #$stream.WriteLine("</HTML>")
        #$stream.close()

        #$stream = [System.IO.StreamWriter] "$LocalSignaturePath\\Reply Signature.htm"
        #$stream.WriteLine("<!DOCTYPE HTML PUBLIC `"-//W3C//DTD HTML 4.0 Transitional//EN`">")
        #$stream.WriteLine("<HTML><HEAD><TITLE>Signature</TITLE>")
        #$stream.WriteLine("<META http-equiv=Content-Type content=`"text/html; charset=windows-1252`">")
        ##$stream.WriteLine("<BODY>")
        ##$stream.WriteLine('<table style="font-family:Arial;border-spacing:0px;color:#000000;font-size:12px; border-collapse:collapse" cellspacing="0" cellpadding="0" border="0"><tr><td style=" padding:0cm 0cm 0cm 0cm;" width="631"><b>')
        #$stream.WriteLine($strName + "</b><br />" + $strTitle)
        #$stream.WriteLine('</td></tr><tr><td style="padding:7pt 5pt 0cm 0cm;"><a href="http://www.airbelgium.com/" target="_blank"><img src="http://airbelgium.com/signatures/signatureLogo.png" id="x_Picture_x0020_53" style="width: 2.5052in; height: 0.7864in;" width="241" border="0" height="76"></a></td></tr>')
        #$stream.WriteLine('<tr style="height:17.75pt;"><td style="padding:7pt 5pt 0cm 0cm; height:17.75pt" width="631" valign="bottom"><i>This message (including any attachments) contains confidential information intended for a specific individual and purpose, and is protected by law. If you are not the intended recipient, you should delete this message and are hereby notified that any disclosure, duplication, or distribution of this message, or the taking of any action based on it, is strictly prohibited.</i></td></tr></table>')
        #$stream.WriteLine("</BODY>")
        #$stream.WriteLine("</HTML>")
        #$stream.close()
 
        If (Test-Path HKCU:'\Software\Microsoft\Office\16.0')
 
        {
            $Outlook = 'Outlook'
            if ($Outlook -ne $null)
            {
                Stop-Process -Name $Outlook -Force
            }
 
            $MSWord = New-Object -comobject word.application
            $EmailOptions = $MSWord.EmailOptions
            $EmailSignature = $EmailOptions.EmailSignature
            $EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
            If ($UseSignOnNew -eq '1')
            {
                $EmailSignature.NewMessageSignature="$SignatureName"
            }
            If ($UseSignOnReply -eq '1')
            {
                $EmailSignature.ReplyMessageSignature="$SignatureNameReply"
            }
            Stop-Process -Name $Outlook

        If ($ForceSignatureNew -eq '1')
        {
            New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force 
            If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
            Else { 
                New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force 
            } 
        }else{
            If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue) { 
                Get-Item -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' | Remove-ItemProperty -Name NewSignature            
            } 
        }

        If ($ForceSignatureReply -eq '1')
        {
            If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
            Else { 
            New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureNameReply -PropertyType 'String' -Force
                } 
        }else{
            If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue) { 
                Get-Item -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' | Remove-ItemProperty -Name ReplySignature            
            } 
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


        #$wshell = New-Object -ComObject Wscript.Shell
        #$wshell.Popup('Script is completed! Mail signatures in Outlook have been updated',0,'All Done',0x0)
    }


















    #kill window
    Get-Process PowerShell | stop-process