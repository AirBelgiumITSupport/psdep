$script = Invoke-WebRequest https://raw.githubusercontent.com/AirBelgiumITSupport/psdep/master/test.ps1
Invoke-Expression $($script.Content)