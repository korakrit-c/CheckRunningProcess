Function CheckRunningProcess(ProcessName)
    strComputer = "."

    Ret = False
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    Set colProcessList = objWMIService.ExecQuery("Select Name from Win32_Process WHERE Name LIKE '" & ProcessName & "%'")

    For Each Process in colProcessList
        objGetOwn = Process.GetOwner(strNameOfUser)
        If strNameOfUser = Environment("UserName") Then
            Ret = True
            Exit For
        Else
            Ret = False
        End If
    Next

    CheckRunningProcess = Ret

End Function
