Attribute VB_Name = "modTasks"
Option Compare Database

' Show a popup of today's incomplete tasks
'----------------------------------------
Public Sub ShowDailyTasks()
    Dim rs As DAO.Recordset
    Dim strMessage As String

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT TaskName, TaskDate FROM Tasks " & _
        "WHERE TaskDate = Date() AND IsCompleted = False")

    If Not rs.EOF Then
        strMessage = "Today's Tasks:" & vbCrLf & vbCrLf
        Do While Not rs.EOF
            strMessage = strMessage & "- " & rs!TaskName & _
                         " (" & Format(rs!TaskDate, "dd-mmm-yyyy") & ")" & vbCrLf
            rs.MoveNext
        Loop
        MsgBox strMessage, vbInformation, "Daily Task Reminder"
    Else
        MsgBox "No tasks due today!", vbInformation, "Daily Task Reminder"
    End If

    rs.Close
    Set rs = Nothing
End Sub

