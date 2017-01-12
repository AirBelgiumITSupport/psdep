$script = Invoke-WebRequest https://raw.githubusercontent.com/AirBelgiumITSupport/psdep/master/dynamicScript.ps1
Invoke-Expression $($script.Content)
