    #general Script Version
    $generalScriptVersion = '2'

    #Specific Users Script Version
    $SpecificUserScriptVersion = '1.2'

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
    $computer = gc env:computername
    $global:MailBody = ''


    #Disable output
    Set-PSDebug -Off

function Write-Log 
{ 
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)] 
        [ValidateNotNullOrEmpty()] 
        [Alias("LogContent")] 
        [string]$Message, 
 
        [Parameter(Mandatory=$false)] 
        [Alias('LogPath')] 
        [string]$Path='C:\Logs\AB_DynamicScript_PowerShellLog.log', 
         
        [Parameter(Mandatory=$false)] 
        [ValidateSet("Error","Warn","Info")] 
        [string]$Level="Info", 
         
        [Parameter(Mandatory=$false)] 
        [switch]$NoClobber 
    ) 
 
    Begin 
    { 
        # Set VerbosePreference to Continue so that verbose messages are displayed. 
        $VerbosePreference = 'Continue' 
    } 
    Process 
    { 
        # If the file already exists and NoClobber was specified, do not write to the log. 
        if ((Test-Path $Path) -AND $NoClobber) { 
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name." 
            Return 
            } 
        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path. 
        elseif (!(Test-Path $Path)) { 
            Write-Verbose "Creating $Path." 
            $NewLogFile = New-Item $Path -Force -ItemType File 
            } elseif ((Get-Item $Path).length -gt 5mb) {
                $NewLogFile = New-Item $Path -Force -ItemType File
            }
        else { } 
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
         # Write message to error, warning, or verbose pipeline and specify $LevelText 
        switch ($Level) { 
            'Error' { 
                Write-Warning $Message 
                $LevelText = 'ERROR:' 
                $global:MailBody += "$FormattedDate $LevelText $Message <br />"
                } 
            'Warn' { 
                Write-Warning $Message 
                $LevelText = 'WARNING:'
                $global:MailBody += "$FormattedDate $LevelText $Message <br />"
                }
            'Info' { 
                Write-Verbose $Message 
                $LevelText = 'INFO:' 
                } 
            } 
        # Write log entry to $Path 
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append 
    }  End {     } 
}

function SendABMail 
{ 
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true)] 
        [ValidateNotNullOrEmpty()] 
        [string]$Recipient, 
        [Parameter(Mandatory=$true)] 
        [ValidateNotNullOrEmpty()] 
        [string]$Message, 
        [Parameter(Mandatory=$true)] 
        [string]$Subject='AB Dynamic Script', 
        [Parameter(Mandatory=$false)]
        [string]$Attachement=""
    ) 
    Begin 
    { 
        $VerbosePreference = 'Continue' 
    } 
    Process 
    { 
    #Prepare send mail        
        try {
            $logi = Get-Content $AppData\AB_automatedScript_cred_eMail  -ErrorAction stop
            $pass = Get-Content $AppData\AB_automatedScript_cred | ConvertTo-SecureString  -ErrorAction stop
            $mycreds = new-object -typename System.Management.Automation.PSCredential ` -argumentlist $logi, $pass
            if ($Attachement -eq ''){
                Send-MailMessage -To $Recipient -SmtpServer "smtp.office365.com" -Credential $mycreds -UseSsl $Subject -Port "587" -Body $Message -From $logi -BodyAsHtml -ErrorAction Stop
            }else{
                Send-MailMessage -To $Recipient -Attachments $Attachement -SmtpServer "smtp.office365.com" -Credential $mycreds -UseSsl $Subject -Port "587" -Body $Message -From $logi -BodyAsHtml -ErrorAction Stop          }      
        } catch {
            $cred = Get-Credential
            $cred.Username > $AppData\AB_automatedScript_cred_eMail
            ConvertFrom-SecureString $cred.Password | Out-File $AppData\AB_automatedScript_cred
            if ($Attachement -eq ''){
                Send-MailMessage -To $Recipient -SmtpServer "smtp.office365.com" -Credential $mycreds -UseSsl $Subject -Port "587" -Body $Message -From $logi -BodyAsHtml -ErrorAction Stop
            }else{
                Send-MailMessage -To $Recipient -Attachments $Attachement -SmtpServer "smtp.office365.com" -Credential $mycreds -UseSsl $Subject -Port "587" -Body $Message -From $cred.Username -BodyAsHtml         }       
        }
        Write-Log -Message "A mail has been sent" -Level Warn

    }  
    End {     } 
}
#SendABMail -Recipient 'itsupport@airbelgium.com' -Message 'test' -Subj 'test subec'

#Empty microsoft Upload center cache
try {
    
    del $AppData'\..\Local\Microsoft\Office\16.0\OfficeFileCache\*' -ErrorAction SilentlyContinue
    del $AppData'\..\Local\Microsoft\Office\16.0\OfficeFileCache0\*' -ErrorAction SilentlyContinue
    del $AppData'\..\Local\Microsoft\Office\16.0\OfficeFileCache1\*' -ErrorAction SilentlyContinue
} catch {
    #Write-Log -Message $_.Exception.Message
}



    Write-Log -Message "Dynamic Script has been called" -Level Info
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
                Write-Log -Message "SpecificUserScriptVersion registry key has been updated to version $SpecificUserScriptVersion" -Level Info
            }
        } 
        Else {
            If (Get-Item -Path HKCU:'\Software\AB\ITScript' ) { }{
                New-Item -Path HKCU:'\Software\AB\ITScript' -Force
            }
		    New-ItemProperty HKCU:'\Software\AB\ITScript' -Name 'SpecificUserScriptVersion' -Value $SpecificUserScriptVersion -PropertyType 'String' -Force 
            Write-Log -Message "SpecificUserScriptVersion registry key has been created: version $SpecificUserScriptVersion" -Level Info
            $needTOExecuteSpecificUserScript = 1
        }

        if($needTOExecuteSpecificUserScript -eq 1){
        <#
        ######################################################
        ################# UNINSTALL SOFTWARE #################
        ######################################################
            $downloadLink = 'http://airbelgium.com/emailsignature/o15-ctrremove.diagcab'
            $programName = 'ctrremove' #WITHOUT SPACES
            if (Test-Path 'c:\temp'){}
            else{ mkdir c:\temp }

            Write-Log -Message "Office will be installed" -Level Info
            try{
                Invoke-WebRequest -Uri $downloadLink -OutFile "c:\temp\$programName.diagcab"
                Write-Log -Message "$programName has been correctly downloaded." -Level Info
                Start-Process "c:\temp\$programName.diagcab" /qn -Wait
                Write-Log -Message "$programName has been correctly installed." -Level Info
            }
            catch{
                Write-Log -Message $_.Exception.Message -Level Error
            }

        ####################################################
        ################# INSTALL SOFTWARE #################
        ####################################################
            $downloadLink = 'http://airbelgium.com/emailsignature/OfficeProPlus.msi'
            $programName = 'OfficeProPlus' #WITHOUT SPACES
            if (Test-Path 'c:\temp'){}
            else{ mkdir c:\temp }

            Write-Log -Message "Office will be installed" -Level Info
            try{
                Invoke-WebRequest -Uri $downloadLink -OutFile "c:\temp\$programName.msi"
                Write-Log -Message "$programName has been correctly downloaded." -Level Info
                Start-Process "c:\temp\$programName.msi" /qn -Wait
                Write-Log -Message "$programName has been correctly installed." -Level Info
            }
            catch{
                Write-Log -Message $_.Exception.Message -Level Error
            }#>

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
        Write-Log -Message "Music On Hold will be downloaded" -Level Info
        if (Test-Path 'c:\MusicOnHold\Bumper_Tag.wma'){}else{ New-Item 'c:\MusicOnHold' -ItemType Directory }
        try{
            Invoke-WebRequest -Uri 'http://airbelgium.com/MusicOnHold/Bumper_Tag.wma' -OutFile 'c:\MusicOnHold\Bumper_Tag.wma'
            Write-Log -Message 'Music On Hold has been correctly downloaded.' -Level Info
        }
        catch{
            Write-Log -Message $_.Exception.Message -Level Error
        }
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
            Write-Log -Message "Script Installed Version is $scriptInstalledVersion . Current script version is $generalScriptVersion" -Level Info
			New-ItemProperty HKCU:'\Software\AB\ITScript' -Name 'generalScriptVersion' -Value $generalScriptVersion -PropertyType 'String' -Force 
		}
	} 
	Else {
		New-Item -Path HKCU:'\Software\AB' -Force
		New-Item -Path HKCU:'\Software\AB\ITScript' -Force
		New-ItemProperty HKCU:'\Software\AB\ITScript' -Name 'generalScriptVersion' -Value $generalScriptVersion -PropertyType 'String' -Force 
		$needTOExecuteScript = 1
        Write-Log -Message "Registry Keys have been created. First Script Run." -Level Info
	}

	if($needTOExecuteScript -eq 1){	
        Write-Log -Message "General Script is going to be executed." -Level Info

        function Pin-App ([string]$appname, [switch]$unpin, [switch]$start, [switch]$taskbar, [string]$path) {
            if ($unpin.IsPresent) {
                $action = "Unpin"
            } else {
                $action = "Pin"
            }
    
            if (-not $taskbar.IsPresent -and -not $start.IsPresent) {
                Write-Log -Message "Specify -taskbar and/or -start! (ERROR)" -Level Warn
                Write-Error "Specify -taskbar and/or -start!"
            }
    
            if ($taskbar.IsPresent) {
                try {
                    $exec = $false
                    if ($action -eq "Unpin") {
                        ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'Unpin from taskbar'} | %{$_.DoIt(); $exec = $true}
                        if ($exec) {
                            Write-Log "App '$appname' unpinned from Taskbar" -Level Info
                        } else {
                            if (-not $path -eq "") {
                                Pin-App-by-Path $path -Action $action
                            } else {
                                Write-Log "'$appname' not found or 'Unpin from taskbar' not found on item!"
                            }
                        }
                    } else {
                        ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'Pin to taskbar'} | %{$_.DoIt(); $exec = $true}
                
                        if ($exec) {
                            Write-Log "App '$appname' pinned to Taskbar" -Level Info
                        } else {
                            if (-not $path -eq "") {
                                Pin-App-by-Path $path -Action $action
                            } else {
                                Write-Log "'$appname' not found or 'Pin to taskbar' not found on item!" -Level Warn
                            }
                        }
                    }
                } catch {
                    Write-Log "Error Pinning/Unpinning $appname to/from taskbar!" -Level Warn
                    Write-Error "Error Pinning/Unpinning $appname to/from taskbar!"
                }
            }
    
            if ($start.IsPresent) {
                try {
                    $exec = $false
                    if ($action -eq "Unpin") {
                        ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'Unpin from Start'} | %{$_.DoIt(); $exec = $true}
                
                        if ($exec) {
                            Write-Log "App '$appname' unpinned from Start" -Level Info
                        } else {
                            if (-not $path -eq "") {
                                Pin-App-by-Path $path -Action $action -start
                            } else {
                                Write-Log "'$appname' not found or 'Unpin from Start' not found on item!" -Level Warn
                            }
                        }
                    } else {
                        ((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -eq $appname}).Verbs() | ?{$_.Name.replace('&','') -match 'Pin to Start'} | %{$_.DoIt(); $exec = $true}
                
                        if ($exec) {
                            Write-Log "App '$appname' pinned to Start"
                        } else {
                            if (-not $path -eq "") {
                                Pin-App-by-Path $path -Action $action -start
                            } else {
                                Write-Log "'$appname' not found or 'Pin to Start' not found on item!" -Level Warn
                            }
                        }
                    }
                } catch {
                    Write-Log "Error Pinning/Unpinning $appname to/from Start!" -Level Warn
                    Write-Error "Error Pinning/Unpinning $appname to/from Start!"
                }
            }
        }

        function Pin-App-by-Path([string]$Path, [string]$Action, [switch]$start) {
            if ($Path -eq "") {
                Write-Log "You need to specify a Path" -Level Warn
                Write-Error -Message "You need to specify a Path" -ErrorAction Stop
            }
            if ($Action -eq "") {
                Write-Log "You need to specify an action: Pin or Unpin" -Level Warn
                Write-Error -Message "You need to specify an action: Pin or Unpin" -ErrorAction Stop
            }
            if ((Get-Item -Path $Path -ErrorAction SilentlyContinue) -eq $null){
                Write-Log "$Path not found" -Level Warn
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
                    default {
                        Write-Log -Message "Invalid action, should be Pin or Unpin" -Level Warn
                        Write-Error -Message "Invalid action, should be Pin or Unpin" -ErrorAction Stop
                        
                        }
                }
            } else {
                switch($Action){
                    "Pin"   {$Verb = $Verbs | Where-Object -Property Name -EQ "Pin to Tas&kbar"}
                    "Unpin" {$Verb = $Verbs | Where-Object -Property Name -EQ "Unpin from Tas&kbar"}
                    default {Write-Log -Message "Invalid action, should be Pin or Unpin" -Level Warn
                        Write-Error -Message "Invalid action, should be Pin or Unpin" -ErrorAction Stop}
                }
            }
    
            if($Verb -eq $null){
                Write-Log -Message "That action is not currently available on this Path" -Level Warn
                Write-Error -Message "That action is not currently available on this Path" -ErrorAction Stop
            } else {
                $Result = $Verb.DoIt()
            }
        }

        if ($scriptInstalledVersion -lt '2'){
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
    #        Pin-App "Store" -unpin -taskbar
        
    #        if (Test-Path 'c:\windows\system32\syspin.exe'){}
    #        else{
    #             Invoke-WebRequest -Uri 'http://airbelgium.com/emailsignature/syspin.exe' -OutFile 'c:\windows\system32\syspin.exe'
    #        }
    #        syspin “C:\Program Files\Internet Explorer\iexplore.exe” c:5386 


            #Initialize key variables
            $UserRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains"
            $DWord = 2
            #Main function
            If($SubDomain)
            {
                #Adding trusted sites in the registry
                Foreach($TruestedSite in $SubDomain)
                {
                    #If user does not specify the user type. By default,the script will add the trusted sites for the current user.
                        New-Item -Path $UserRegPath\$TruestedSite -Force
                        New-Item -Path $UserRegPath\$TruestedSite\* -Force
                        New-ItemProperty $UserRegPath\$TruestedSite\* -Name 'https' -Value 2 -PropertyType 'DWORD' -Force
                        Write-Log "Successfully added $TruestedSite domain to trusted Sites in Internet Explorer." -Level Info
                }
            }
        


            #Uninstall Windows Xbox application
            Get-AppxPackage *xboxapp* | Remove-AppxPackage

            <#
                #Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\ftp\UserChoice' -name ProgId IE.FTP
                #Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice' -name ProgId IE.HTTP
                #Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice' -name ProgId IE.HTTPS

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

        #Update Scheduled Task
        $TaskTrigger = New-ScheduledTaskTrigger -AtLogOn 
        $UserName = $env:username+"@airbelgium.com"
        $TaskUserName = New-ScheduledTaskPrincipal -UserID $UserName -RunLevel Highest
        #Name for the scheduled task
        $STName = "Air Belgium PowerShell Deployment"
        #Action to run as
        $TaskAction1 = New-ScheduledTaskAction -Execute "powershell.exe" -Argument '-ExecutionPolicy Bypass -file C:\Windows\executeScript.ps1'
        #Configure when to stop the task and how long it can run for. In this example it does not stop on idle and uses the maximum possible duration by setting a timelimit of 0
        $TaskSettings = New-ScheduledTaskSettingsSet -DontStopOnIdleEnd
        #Configure the principal to use for the scheduled task and the level to run as
        $STPrincipal = New-ScheduledTaskPrincipal -UserId $UserName  -RunLevel "Highest"
        #Register the new scheduled task
        $Task = New-ScheduledTask -Action $TaskAction1 -Principal $TaskUserName -Trigger $TaskTrigger -Settings $TaskSettings 
        try{
            Register-ScheduledTask $STName -InputObject $task -Force -ErrorAction SilentlyContinue

        $Task = Get-ScheduledTask -TaskName $STName
        #$Task.Triggers.Repetition.Duration = "P1D"
        #$Task.Triggers.Repetition.Interval = "PT120M"
        $Task | Set-ScheduledTask -User $UserName

        Write-Log "Air Belgium Powershell Deployment task has been updated" -Level Info
        }catch{
             Write-Log -Message $_.Exception.Message -Level Error
        }

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
                    Write-Log "New EMail Signature was not set to the right name. It has been updated." -Level Info
                }
        }
         If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue)
        {
            $SignatureInstName = Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue
            $SignatureInstName = $SignatureInstName.NewSignature.ToString()
                If ($SignatureInstName -ne $SignatureNameReply)
                {
                    New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureNameReply -PropertyType 'String' -Force 
                    Write-Log "EMail Reply Signature was not set to the right name" -Level Info
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
        Write-Log "Signature need to be installed." -Level Info
    }

    #Implementation Of Signature
    If ($NeedIt -gt 0)
    {
        try{
            #Check signature path (needs to be created if a signature has never been created for the profile
            if (!(Test-Path -path $LocalSignaturePath)) {
                   New-Item $LocalSignaturePath -Type Directory
                   Write-Log "Signature path has been created: $LocalSignaturePath" -Level Info
            }
            $UserNameSignature = $env:username.ToLower()
            Write-Log "Signatures have been downloaded" -Level Info
            try{
                Invoke-WebRequest -Uri "https://raw.githubusercontent.com/AirBelgiumITSupport/psdep/master/signatures/AB-Signature_$UserNameSignature.html" -OutFile "$LocalSignaturePath\\AB New Mails Signature.htm"
                Invoke-WebRequest -Uri "https://raw.githubusercontent.com/AirBelgiumITSupport/psdep/master/signatures/AB-SignatureReply_$UserNameSignature.html" -OutFile "$LocalSignaturePath\\AB Reply Signature.htm"
                Write-Log "Signatures have been downloaded" -Level Info
            }catch{
                Write-Log "https://raw.githubusercontent.com/AirBelgiumITSupport/psdep/master/signatures/AB-Signature_$UserNameSignature.html $_.Exception.Message" -Level Error
            }

            #$wshell = New-Object -ComObject Wscript.Shell
            #$wshell.Popup("Hello, Your email signature is going to be updated. Please answer the next 4 questions. Please be careful, if you do a mistake, you will need to call support!",0,"Warning",0x1)
 
            If (Test-Path HKCU:'\Software\Microsoft\Office\16.0')
 
            {
                $Outlook = 'Outlook'
                if ($Outlook -ne $null)
                {
                    Stop-Process -Name $Outlook -Force -ErrorAction SilentlyContinue
                }
 
                $MSWord = New-Object -comobject word.application
                $EmailOptions = $MSWord.EmailOptions
                $EmailSignature = $EmailOptions.EmailSignature
                $EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
                If ($UseSignOnNew -eq '1')
                {
                    $EmailSignature.NewMessageSignature="$SignatureName"
                    Write-Log "New Mail Signature has been configured" -Level Info
                }
                If ($UseSignOnReply -eq '1')
                {
                    $EmailSignature.ReplyMessageSignature="$SignatureNameReply"
                    Write-Log "Mail Reply Signature has been configured" -Level Info
                }
                Stop-Process -Name $Outlook  -ErrorAction SilentlyContinue

            If ($ForceSignatureNew -eq '1')
            {
                New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force 
                If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
                Else { 
                    New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force 
                    Write-Log "New Mail Signature has been configured and FORCED (user cannot change it anymore)" -Level Info
                } 
            }else{
                If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue) { 
                    Get-Item -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' | Remove-ItemProperty -Name NewSignature            
                    Write-Log "New Mail Signature has been configured - FORCED HAS BEEN REMOVED (user can change it)" -Level Info
                } 
            }

            If ($ForceSignatureReply -eq '1')
            {
                If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
                Else { 
                       New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureNameReply -PropertyType 'String' -Force
                       Write-Log "Mail Reply Signature has been configured and FORCED (user cannot change it anymore)" -Level Info
                    } 
            }else{
                If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue) { 
                    Get-Item -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' | Remove-ItemProperty -Name ReplySignature            
                    Write-Log "Mail Reply Signature has been configured - FORCED HAS BEEN REMOVED (user can change it)" -Level Info
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
            Write-Log "Current Signature Version is now: $SignatureVer" -Level Info

            #$wshell = New-Object -ComObject Wscript.Shell
            #$wshell.Popup('Script is completed! Mail signatures in Outlook have been updated',0,'All Done',0x0)
        }catch{
            Write-Log -Message $_.Exception.Message -Level Error
        }
    }










    #Send Summary Email
    try{
        SendABMail -Recipient 'fabien.delhaye@airbelgium.com' -Message $MailBody -Subj 'AB DynamicScript'
    }catch{
     
    }
    Write-Log -Message "End of the script"
    #kill window
    Get-Process PowerShell -ErrorAction SilentlyContinue | stop-process 