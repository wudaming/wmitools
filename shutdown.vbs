'uint32 Win32ShutdownTracker(
'  [in] uint32 Timeout,
'  [in] string Comment,
'  [in] uint32 ReasonCode,
'  [in] sint32 Flags
');
' Flag: 0 = Logoff, 4 = Forced Logoff (0+4), 1 = Shutdown, 2 = Reboot, 6 = Forced Reboot (2+4), 8 = Power Off, 12 = Forced Power Off (8+4) - 2 (Reboot) 


strComputer = "." 
Set objWMIService = GetObject("winmgmts:{impersonationlevel=impersonate,(debug, shutdown)}\\" & strComputer & "\root\CIMV2") 

for each OsObject in objWMIService.instancesof("Win32_OperatingSystem")
    result = OsObject.Win32ShutdownTracker(60,"shutdown test",0,1)
Next
