Attribute VB_Name = "modAgenda"
Option Explicit
Public AdressRegisterPath As String 'Tells the path to the choosen Adressregister
Public objCon As ADODB.Connection
Public objRs As ADODB.Recordset     'The Contactsbook recordset object
Public objCalRS As ADODB.Recordset  'The Calendar recordset object

Public Type Alert
    Description As String
    Memo As String
    SoundPath As String
    Time As Date
    Index As Integer
End Type
Public Warning() As Alert
Public NoOfAlerts As Integer

Sub Main()
Load frmMenu
End Sub

Public Function GetCon() As ADODB.Connection
If objCon.State <> adStateOpen Then
    objCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AdressRegisterPath & _
                ";Persist Security Info=False"
    objCon.Open
End If
Set GetCon = objCon
End Function

Public Sub GetTodaysAlerts()
Dim i As Integer
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
NoOfAlerts = 0
ReDim Warning(NoOfAlerts)
With RS
If .State = adStateOpen Then .Close

.ActiveConnection = GetCon
.CursorType = adOpenKeyset
.Source = "Select * from tblCal where datum =#" & Date & "# order by tid"
.Open

If .RecordCount > 0 Then
    .MoveFirst
While Not .EOF
    NoOfAlerts = NoOfAlerts + 1
    ReDim Preserve Warning(NoOfAlerts)
    Warning(NoOfAlerts).Time = Format(.Fields(2), "Hh:Nn")
    Warning(NoOfAlerts).SoundPath = .Fields(6) & ""
    Warning(NoOfAlerts).Description = .Fields(3)
    Warning(NoOfAlerts).Memo = .Fields(4)
    Warning(NoOfAlerts).Index = i
    .MoveNext
    i = i + 1
Wend
End If
RS.Close
Set RS = Nothing
End With

End Sub

