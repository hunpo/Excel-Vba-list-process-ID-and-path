# Excel-Vba-list-process-ID-and-path
'Using excel vba list process id and path via VMI 
Private Sub process()
    On Error Resume Next
    Dim counts As Byte, i As Integer, xProcesses As Object
    Set xProcesses = GetObject("WinMgmts:").InstancesOf("Win32_Process")
    counts = xProcesses.Count + 1
    
    i = 2
    Range("A1:G65000").ClearContents
    
    cells(1, 1) = "½ø³Ì"
    cells(1, 2) = "ÓÃ»§"
    cells(1, 3) = "½ø³ÌID"
    cells(1, 4) = "ÄÚ´æ"
    cells(1, 5) = "Â·¾¶"
    For Each xProcess In xProcesses
        With xProcess
            If .GetOwner(user, Domain) = 0 Then
                cells(i, 1) = .Caption: cells(i, 2) = user: cells(i, 3) = .ProcessID: cells(i, 4) = .WorkingSetSize / 1024: cells(i, 5) = .ExecutablePath
            Else
                cells(i, 1) = .Caption: cells(i, 2) = "": cells(i, 3) = .ProcessID: cells(i, 4) = .WorkingSetSize / 1024: cells(i, 5) = .ExecutablePath
            End If
        End With
        i = i + 1
    Next
    Range("A1").Sort (Columns("A")), xlAscending, Header:=xlYes
    'ActiveCell.CurrentRegion.EntireColumn.AutoFit
    'ActiveCell.CurrentRegion.EntireRow.AutoFit
    Range("a:e").EntireColumn.AutoFit
End Sub



