    #general Script Version
    $generalScriptVersion = '1.0'

    #Specific Users Script Version
    $SpecificUserScriptVersion = '1.0'

    #Custom variables For Signature Management
    $SignatureName = 'New Mails Signature'
    $SignatureNameReply = 'Reply Signature'
    $SignatureVer = '1.1'
    $UseSignOnNew = '1'        #If set to '0', the signature will be added as signature for new mails.
    $UseSignOnReply = '1'      #If set to '0', the signature will be added as signature for reply mails.
    $ForceSignatureNew = '0'   #If set to '0', the signature will be editable in Outlook and if set to '1' will be non-editable and forced - forced also as reply.
    $ForceSignatureReply = '0' #If set to '0', the signature will be editable in Outlook and if set to '1' will be non-editable and forced.







    #Environnment Variables
    $AppData=(Get-Item env:appdata).value
    $SigPath = '\Microsoft\Signatures'
    $LocalSignaturePath = $AppData+$SigPath
    $RemoteSignaturePathFull = $SigSource
    $UserName = $env:username


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
    if ($UserName -eq 'thierry.naert'){
        #New-ItemProperty HKCU:'\Software\AB\ITScript' -Name 'ABScriptVersion' -Value $SpecificUserScriptVersion -PropertyType 'String' -Force 
        If (Get-ItemProperty -Name 'ABScriptVersion' -Path HKCU:'\Software\AB\ITScript' -ErrorAction SilentlyContinue) { 
            $SpecificUserScriptInstalledVersion = Get-ItemProperty -Name 'ABScriptVersion' -Path HKCU:'\Software\AB\ITScript'
            $SpecificUserScriptInstalledVersion = $SpecificUserScriptInstalledVersion.ABScriptVersion.ToString()

            if ($SpecificUserScriptInstalledVersion -eq $SpecificUserScriptVersion){}else{
                $needTOExecuteSpecificUserScript = 1
                New-ItemProperty HKCU:'\Software\AB\ITScript' -Name 'ABScriptVersion' -Value $SpecificUserScriptVersion -PropertyType 'String' -Force 
            }
        } 
        Else {
            New-Item -Path HKCU:'\Software\AB\ITScript' -Force 
            New-ItemProperty HKCU:'\Software\AB\ITScript' -Name 'ABScriptVersion' -Value $SpecificUserScriptVersion -PropertyType 'String' -Force 
            $needTOExecuteSpecificUserScript = 1
        }

        if($needTOExecuteSpecificUserScript -eq 1){
            #Do your admin stuff !
            #Do your admin stuff !
            #Do your admin stuff !
            #Do your admin stuff !

            $wshell = New-Object -ComObject Wscript.Shell
            $wshell.Popup(".... Banana? ",0,"Warning",0x1)
            $wshell.Popup("Hello Thierry! ",0,"Warning",0x1)
            $wshell.Popup("That's just a test ",0,"Warning",0x1)
            Start-Sleep -s 4
            $wshell.Popup("No more popups!",0,"Warning",0x1)
            $wshell.Popup("Promise !",0,"Warning",0x1)
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
		New-Item -Path HKCU:'\Software\AB\ITScript' -Force 
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
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/AirBelgiumITSupport/psdep/master/signatures/AB-Signature_$UserNameSignature.html" -OutFile "$LocalSignaturePath\\New Mails Signature.htm"
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/AirBelgiumITSupport/psdep/master/signatures/AB-SignatureReply_$UserNameSignature.html" -OutFile "$LocalSignaturePath\\Reply Signature.htm"
 
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