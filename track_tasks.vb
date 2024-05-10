Sub AddTask()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TaskList")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Prompt user for task details
    Dim taskName As String
    taskName = InputBox("Enter task name:")
    
    ' Add task details to the appropriate cells
    ws.Cells(lastRow + 1, 1).Value = taskName
    ' Add other task details similarly
End Sub

Sub UpdateProgress()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TaskList")
    
    ' Prompt user to select a task
    Dim taskRow As Long
    taskRow = Application.Match(InputBox("Enter task name:"), ws.Columns(1), 0)
    
    If Not IsError(taskRow) Then
        ' Prompt user to enter the progress percentage
        Dim progressPercent As Double
        progressPercent = InputBox("Enter progress percentage:")
        
        ' Update the progress cell for the selected task
        ws.Cells(taskRow, 8).Value = progressPercent
    Else
        MsgBox "Task not found."
    End If
End Sub

Sub HighlightTasks()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TaskList")
    
    ' Apply conditional formatting to highlight overdue tasks
    ' Example:
    ' ws.Range("E2:E" & ws.Cells(ws.Rows.Count, "E").End(xlUp).Row).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=TODAY()"
    ' ws.Range("E2:E" & ws.Cells(ws.Rows.Count, "E").End(xlUp).Row).FormatConditions(1).Interior.Color = RGB(255, 0, 0)
    
    ' Apply conditional formatting to highlight high-priority tasks
    ' Example:
    ' ws.Range("F2:F" & ws.Cells(ws.Rows.Count, "F").End(xlUp).Row).FormatConditions.Add Type:=xlTextString, String:="High", TextOperator:=xlContains
    ' ws.Range("F2:F" & ws.Cells(ws.Rows.Count, "F").End(xlUp).Row).FormatConditions(2).Interior.Color = RGB(255, 255, 0)
    
    ' Apply conditional formatting to highlight completed tasks
    ' Example:
    ' ws.Range("H2:H" & ws.Cells(ws.Rows.Count, "H").End(xlUp).Row).FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="100"
    ' ws.Range("H2:H" & ws.Cells(ws.Rows.Count, "H").End(xlUp).Row).FormatConditions(3).Interior.Color = RGB(0, 255, 0)
End Sub

Sub FilterAndSort()
    ' Enable filtering and sorting options for the task list based on different criteria
    ' Example:
    ' ThisWorkbook.Sheets("TaskList").AutoFilterMode = False
    ' ThisWorkbook.Sheets("TaskList").Range("A1:H" & ThisWorkbook.Sheets("TaskList").Cells(ThisWorkbook.Sheets("TaskList").Rows.Count, "A").End(xlUp).Row).AutoFilter
End Sub

Sub GenerateCharts()
    ' Generate charts or graphs to visualize project progress, task distribution, and status
    ' Example:
    ' ThisWorkbook.Sheets.Add
    ' ThisWorkbook.ActiveSheet.Shapes.AddChart2.Select
    ' With Selection
    '     .ChartType = xlColumnClustered
    '     .SetSourceData Source:=ThisWorkbook.Sheets("TaskList").Range("H2:H" & ThisWorkbook.Sheets("TaskList").Cells(ThisWorkbook.Sheets("TaskList").Rows.Count, "H").End(xlUp).Row)
    ' End With
End Sub

Sub CommunicationPlatform()
    ' Set up a separate sheet or section for team members to input updates, issues, and comments related to tasks
End Sub

Sub SetAlerts()
    ' Use Excel's conditional formatting or add-ins to create alerts for approaching or overdue deadlines
    ' Example:
    ' Use Excel's built-in conditional formatting rules for dates
End Sub

Sub UpdateTracker()
    ' Ensure that the tracker is updated regularly with the latest task status and progress
End Sub

Sub ReviewTracker()
    ' Periodically review the tracker to identify bottlenecks and adjust priorities accordingly
End Sub
