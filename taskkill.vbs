Function KillAll(ProcessName)

    Dim objWMIService, colProcess
    Dim strComputer, strList, p
    Dim i :i= 0
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name like '" & ProcessName & "'")

    For Each p in colProcess
        p.Terminate    
    i = i+1       
    Next

End Function

call KillAll("EXCEL.exe")