Sub Export_form_2H2024()
    Dim wsData As Worksheet
    Dim wsPSC As Worksheet, wsGoalPSC As Worksheet
    Dim wsOB As Worksheet, wsGoal_OB As Worksheet
    Dim wsOther As Worksheet, wsGoal_Other As Worksheet
    Dim cell As Range
    Dim agent_name_string As String, agent_code_string As String, service_lob As String
    Dim file_name As String, savePath As String
    Dim tempWorkbook As Workbook, trackingWorkbook As Workbook
    Dim trackingSheet As Worksheet
    Dim trackingPath As String
    Dim lastRow As Long

    
    ' Set worksheets
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsPSC = ThisWorkbook.Sheets("PSC")
    Set wsGoalPSC = ThisWorkbook.Sheets("Goal_PSC")
    Set wsOB = ThisWorkbook.Sheets("OB")
    Set wsGoal_OB = ThisWorkbook.Sheets("Goal_OB")
    Set wsOther = ThisWorkbook.Sheets("Other")
    Set wsGoal_Other = ThisWorkbook.Sheets("Goal_Other")
    
    ' Tracking file path
    trackingPath = "D:\1. H-CONFIDENTAL\4.H - Special\2H2024.G1-G3\Tracking_Export_Process.xlsx"
    
    Application.ScreenUpdating = False
    
    For Each cell In wsData.Range("A2:A" & wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row)
        agent_name_string = wsData.Cells(cell.Row, 5).Value
        agent_code_string = wsData.Cells(cell.Row, 6).Value
        service_lob = wsData.Cells(cell.Row, 1).Value
        file_name = service_lob & "_" & agent_code_string & "_" & agent_name_string
        
        ' Define save path based on service_lob
        Select Case service_lob
            Case "PSC-CSC"
                savePath = "D:\1. H-CONFIDENTAL\4.H - Special\2H2024.G1-G3\PSC-CSC\" & file_name & ".xlsx"
                Set tempWorkbook = Workbooks.Add
                wsPSC.Copy Before:=tempWorkbook.Sheets(1)
                tempWorkbook.Sheets(1).Name = "G1-G3 " & agent_name_string
                tempWorkbook.Sheets(1).Range("M6").Value = agent_code_string
                'custom 01
                tempWorkbook.Sheets(1).Range("B6:O9").Value = tempWorkbook.Sheets(1).Range("B6:O9").Value
                tempWorkbook.Sheets(1).Range("L20:L24").Value = tempWorkbook.Sheets(1).Range("L20:L24").Value
                tempWorkbook.Sheets(1).Range("N20:N24").Value = tempWorkbook.Sheets(1).Range("N20:N24").Value
                tempWorkbook.Sheets(1).Range("L39:L44").Value = tempWorkbook.Sheets(1).Range("L39:L44").Value
                tempWorkbook.Sheets(1).Range("N39:N44").Value = tempWorkbook.Sheets(1).Range("N39:N44").Value
                tempWorkbook.Sheets(1).Range("B50:O52").Value = tempWorkbook.Sheets(1).Range("B50:O52").Value
                tempWorkbook.Sheets(1).Range("C59:O64").Value = tempWorkbook.Sheets(1).Range("C59:O64").Value
                
                wsGoalPSC.Copy After:=tempWorkbook.Sheets(tempWorkbook.Sheets.Count)
                tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Name = "Goal " & agent_name_string
                tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Range("M6").Value = agent_code_string
                
                'custom 02
                tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Range("B6:O9").Value = tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Range("B6:O9").Value
                tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Range("C59:O64").Value = tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Range("C59:O64").Value
                
            Case "Outbound"
                savePath = "D:\1. H-CONFIDENTAL\4.H - Special\2H2024.G1-G3\Outbound\" & file_name & ".xlsx"
                Set tempWorkbook = Workbooks.Add
                wsOB.Copy Before:=tempWorkbook.Sheets(1)
                tempWorkbook.Sheets(1).Name = "G1-G3 " & agent_name_string
                tempWorkbook.Sheets(1).Range("M6").Value = agent_code_string
                
                'custom 01
                tempWorkbook.Sheets(1).Range("B6:O9").Value = tempWorkbook.Sheets(1).Range("B6:O9").Value
                tempWorkbook.Sheets(1).Range("L20:L24").Value = tempWorkbook.Sheets(1).Range("L20:L24").Value
                tempWorkbook.Sheets(1).Range("N20:N24").Value = tempWorkbook.Sheets(1).Range("N20:N24").Value
                tempWorkbook.Sheets(1).Range("L39:L44").Value = tempWorkbook.Sheets(1).Range("L39:L44").Value
                tempWorkbook.Sheets(1).Range("N39:N44").Value = tempWorkbook.Sheets(1).Range("N39:N44").Value
                tempWorkbook.Sheets(1).Range("B50:O52").Value = tempWorkbook.Sheets(1).Range("B50:O52").Value
                tempWorkbook.Sheets(1).Range("C59:O64").Value = tempWorkbook.Sheets(1).Range("C59:O64").Value
                
                wsGoal_OB.Copy After:=tempWorkbook.Sheets(tempWorkbook.Sheets.Count)
                tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Name = "Goal " & agent_name_string
                tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Range("M6").Value = agent_code_string
                
                'custom 02
                tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Range("B6:O9").Value = tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Range("B6:O9").Value
                tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Range("C59:O64").Value = tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Range("C59:O64").Value
                
            Case Else
                savePath = "D:\1. H-CONFIDENTAL\4.H - Special\2H2024.G1-G3\Other\" & file_name & ".xlsx"
                Set tempWorkbook = Workbooks.Add
                wsOther.Copy Before:=tempWorkbook.Sheets(1)
                tempWorkbook.Sheets(1).Name = "G1-G3 " & agent_name_string
                tempWorkbook.Sheets(1).Range("M6").Value = agent_code_string
                
                'custom 01
                tempWorkbook.Sheets(1).Range("B6:O9").Value = tempWorkbook.Sheets(1).Range("B6:O9").Value
                tempWorkbook.Sheets(1).Range("L20:L24").Value = tempWorkbook.Sheets(1).Range("L20:L24").Value
                tempWorkbook.Sheets(1).Range("N20:N24").Value = tempWorkbook.Sheets(1).Range("N20:N24").Value
                tempWorkbook.Sheets(1).Range("L39:L44").Value = tempWorkbook.Sheets(1).Range("L39:L44").Value
                tempWorkbook.Sheets(1).Range("N39:N44").Value = tempWorkbook.Sheets(1).Range("N39:N44").Value
                tempWorkbook.Sheets(1).Range("B50:O52").Value = tempWorkbook.Sheets(1).Range("B50:O52").Value
                tempWorkbook.Sheets(1).Range("C59:O64").Value = tempWorkbook.Sheets(1).Range("C59:O64").Value
                
                wsGoal_Other.Copy After:=tempWorkbook.Sheets(tempWorkbook.Sheets.Count)
                tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Name = "Goal " & agent_name_string
                tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Range("M6").Value = agent_code_string
                
                'custom 02
                tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Range("B6:O9").Value = tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Range("B6:O9").Value
                tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Range("C59:O64").Value = tempWorkbook.Sheets(tempWorkbook.Sheets.Count).Range("C59:O64").Value
                
        End Select
        
         ' Delete default Sheet1 if it exists
        Application.DisplayAlerts = False
        Dim ws As Worksheet
        For Each ws In tempWorkbook.Sheets
            If ws.Name = "Sheet1" Then ws.Delete
        Next ws
        Application.DisplayAlerts = True
            
        ' Save new workbook with overwrite enabled
        Application.DisplayAlerts = False
        If Dir(savePath) <> "" Then
        On Error Resume Next
            Kill savePath
            If Err.Number <> 0 Then
                MsgBox "Không th? xóa file: " & savePath, vbExclamation
                Err.Clear
            End If
            On Error GoTo 0
        End If
        tempWorkbook.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        tempWorkbook.Close SaveChanges:=False
        
' Record tracking log by appending new data
        On Error Resume Next
        Set trackingWorkbook = Workbooks.Open(trackingPath)
        If trackingWorkbook Is Nothing Then
            Set trackingWorkbook = Workbooks.Add
            Set trackingSheet = trackingWorkbook.Sheets(1)
            trackingSheet.Name = "Tracking"
            trackingSheet.Range("A1:G1").Value = Array("Timestamp", "LOB", "Emp_code", "Full Name", "File Name", "Save Path", "User")
        Else
            Set trackingSheet = trackingWorkbook.Sheets(1)
        End If
        On Error GoTo 0
        
        ' Find the last row and append new record instead of overwriting
        lastRow = trackingSheet.Cells(trackingSheet.Rows.Count, 1).End(xlUp).Row + 1
        trackingSheet.Cells(lastRow, 1).Value = Now()
        trackingSheet.Cells(lastRow, 2).Value = service_lob
        trackingSheet.Cells(lastRow, 3).Value = agent_code_string
        trackingSheet.Cells(lastRow, 4).Value = agent_name_string
        trackingSheet.Cells(lastRow, 5).Value = file_name & ".xlsx"
        trackingSheet.Cells(lastRow, 6).Value = savePath
        trackingSheet.Cells(lastRow, 7).Value = Application.UserName
        
        ' Save and close tracking workbook
        Application.DisplayAlerts = False
        trackingWorkbook.SaveAs Filename:=trackingPath, FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        trackingWorkbook.Close SaveChanges:=True
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "Export completed!", vbInformation
End Sub


