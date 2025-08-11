VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'----------------------------------------
' Show a popup of today's incomplete tasks
'----------------------------------------
Private Sub ShowDailyTasks()
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

'----------------------------------------
' Add New Task button
'----------------------------------------
Private Sub cmdAddNew_Click()
    ' Save current record if there are changes
    If Me.Dirty Then Me.Dirty = False
    
    ' Move to a blank record for adding a new task
    DoCmd.GoToRecord , , acNewRec
    
    ' Optional defaults for new task
    Me.IsCompleted = False
    Me.TaskDate = Date
End Sub

'----------------------------------------
' When the form opens
'----------------------------------------
Private Sub Form_Load()
    ShowDailyTasks
End Sub

'----------------------------------------
' When Submit button is clicked
'----------------------------------------
Private Sub cmdSubmit_Click()
    ' Save the record
    If Me.Dirty Then Me.Dirty = False

    ' Tell the user what happened
    If Nz(Me.IsCompleted, False) Then
        MsgBox "Task marked as complete.", vbInformation, "Submitted"
    Else
        MsgBox "Task marked as incomplete.", vbInformation, "Submitted"
    End If

    ' Show updated popup
    ShowDailyTasks
End Sub

'----------------------------------------
' After a new task is inserted
'----




