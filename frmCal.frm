VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New/Edit Calendar events"
   ClientHeight    =   4185
   ClientLeft      =   7230
   ClientTop       =   525
   ClientWidth     =   7200
   Icon            =   "frmCal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkEditmode 
      Caption         =   "Editmode"
      Height          =   255
      Left            =   4020
      TabIndex        =   24
      ToolTipText     =   "Unheck me if you want to browse through already set events"
      Top             =   1140
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Alert"
      Height          =   555
      Left            =   60
      TabIndex        =   21
      ToolTipText     =   "Tells if to use sound alert or not"
      Top             =   2280
      Width           =   1395
      Begin VB.OptionButton optAlert 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optAlert 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComCtl2.MonthView Cal 
      Height          =   2370
      Left            =   4020
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1380
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   19202050
      TitleBackColor  =   12632256
      TrailingForeColor=   -2147483633
      CurrentDate     =   37217
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   3810
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5786
            MinWidth        =   5786
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "Edit mode"
            TextSave        =   "Edit mode"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2187
            MinWidth        =   2187
            TextSave        =   "12/1/2001"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2010
            MinWidth        =   2010
            TextSave        =   "2:14 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Get"
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Browse for the soundfile"
      Top             =   2220
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3000
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Sound (*.mp3;*.wav;*.mid)|*.mp3;*.wav;*.mid"
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Add new"
      Height          =   375
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Add new event"
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Update"
      Height          =   375
      Index           =   1
      Left            =   900
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Update current event"
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Delete"
      Height          =   375
      Index           =   2
      Left            =   1740
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Delete Current event"
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Clear"
      Height          =   375
      Index           =   3
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Clears everything"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtRepeatTimes 
      Height          =   285
      Left            =   2940
      TabIndex        =   10
      Text            =   "1"
      ToolTipText     =   "How many times to repeat the event"
      Top             =   3000
      Width           =   375
   End
   Begin VB.ListBox lstRepeatDays 
      Height          =   255
      Left            =   600
      TabIndex        =   9
      ToolTipText     =   "Mark the interval you want to repeat the event"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   4
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   7
      Top             =   2520
      Width           =   2355
   End
   Begin VB.TextBox txtCal 
      Height          =   1365
      Index           =   3
      Left            =   60
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      ToolTipText     =   "Memo for the event"
      Top             =   780
      Width           =   3855
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   2
      Left            =   900
      MaxLength       =   30
      TabIndex        =   2
      ToolTipText     =   "Short description of the event"
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "The Time the event is or is to be fired"
      Top             =   240
      Width           =   735
   End
   Begin VB.ListBox lstcal 
      Height          =   840
      Left            =   4020
      TabIndex        =   0
      ToolTipText     =   "The events of the selected day"
      Top             =   240
      Width           =   3075
   End
   Begin VB.Label lblcal 
      Caption         =   "Calendar events"
      Height          =   255
      Index           =   3
      Left            =   4020
      TabIndex        =   17
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblRepeatTimes 
      Caption         =   "No of times"
      Height          =   255
      Left            =   2100
      TabIndex        =   12
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lbRepeatdays 
      Caption         =   "Repeat"
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblcal 
      Caption         =   "Path to alert sound"
      Height          =   255
      Index           =   4
      Left            =   1620
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblcal 
      Caption         =   "Memo"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   6
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label lblcal 
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   900
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblcal 
      Caption         =   "Time"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
   Begin VB.Menu menuSystray 
      Caption         =   "SystrayMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuSystrayOpen 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuSystrayClose 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private NoOfAlerts As Integer


Private Sub Cal_DateClick(ByVal DateClicked As Date)
Cal.Value = DateClicked
If chkEditmode.Value = 0 Then GetCalRecords
End Sub

Private Sub cmdEdit_Click(Index As Integer)

Select Case Index
    Case 0
        If CheckData = True Then
            CreateNew
            GetCalRecords
        End If
    Case 1
        If objCalRS.RecordCount > 0 Then
            If CheckData = True Then
                UpdateCurPost
                GetCalRecords
            End If
        End If
    Case 2
        If objCalRS.RecordCount > 0 Then
            DeleteCurPost
            GetCalRecords
        End If
    Case 3
        Clear
    Case 4
        CD.DialogTitle = "Choose Soundfile for the alert"
        CD.ShowOpen
        If CD.FileName <> "" Then txtCal(4).Text = CD.FileName
End Select
End Sub

Private Sub cmdGetPath_Click()
CD.ShowOpen
If CD.FileName <> "" Then txtCal(4).Text = CD.FileName
End Sub

Private Sub Form_Load()
Me.Show
ListDays
Cal.Value = Date
Set objCalRS = New ADODB.Recordset
GetCalRecords
lstRepeatDays.Selected(0) = True
End Sub

Private Function CheckData() As Boolean
Dim i As Integer
Dim bolTemp As Boolean
If IsDate(txtCal(1)) = True Then
    If Trim(txtCal(2).Text) <> "" Then
        bolTemp = True
    Else
        Beep
        StatusBar1.Panels(1).Text = "The description cannot be empty"
        txtCal(2).SetFocus
    End If
Else
    Beep
    StatusBar1.Panels(1).Text = "The time is not valid"
    txtCal(1).SetFocus
End If
CheckData = bolTemp
End Function

Private Sub ListDays()
lstRepeatDays.AddItem "One Time"
lstRepeatDays.AddItem "Everyday"
lstRepeatDays.AddItem "Mon-Fri"
lstRepeatDays.AddItem "Sat-Sun"
lstRepeatDays.AddItem "Monday"
lstRepeatDays.AddItem "Tuesday"
lstRepeatDays.AddItem "Wednesday"
lstRepeatDays.AddItem "Thursday"
lstRepeatDays.AddItem "Friday"
lstRepeatDays.AddItem "Saturday"
lstRepeatDays.AddItem "Sunday"
lstRepeatDays.AddItem "Annual"
End Sub

Private Sub GetCalRecords()
Dim i As Integer
Clear
lstcal.Clear
With objCalRS
If .State = adStateOpen Then .Close

.ActiveConnection = GetCon
.LockType = adLockOptimistic
.CursorType = adOpenKeyset
.Source = "Select * from tblCal where datum =#" & Cal.Value & "# order by tid"
.Open

If .RecordCount > 0 Then
    .MoveFirst
While Not .EOF
    lstcal.AddItem Format(.Fields(2), "Hh:Nn") & " " & .Fields(3)
    .MoveNext
Wend
    .MoveFirst
    showCurrentItem
    StatusBar1.Panels(1).Text = .RecordCount & " Records !"
Else
    Beep
    StatusBar1.Panels(1).Text = "No records !"
End If

End With
End Sub

Private Sub DeleteCurPost()
objCalRS.Delete adAffectCurrent
objCalRS.Update
End Sub

Private Sub UpdateCurPost()
objCalRS.Fields(2) = Trim(txtCal(1))
objCalRS.Fields(3) = Trim(txtCal(2))
objCalRS.Fields(4) = Trim(txtCal(3))
objCalRS.Fields(6) = Trim(txtCal(4))
objCalRS.Fields(1) = Cal.Value
If optAlert(0) Then objCalRS.Fields(5) = True
If optAlert(1) Then objCalRS.Fields(5) = False
objCalRS.Update
End Sub

Private Sub CreateNew()
Dim mydate As Date
Dim i As Integer
mydate = Cal.Value
For i = 1 To Val(txtRepeatTimes)
    objCalRS.AddNew
    Select Case lstRepeatDays.ListIndex
        Case 0
        objCalRS.Fields(1) = mydate
        Case 1
            objCalRS.Fields(1) = mydate
            mydate = DateAdd("d", 1, mydate)
        Case 2
            objCalRS.Fields(1) = mydate
            mydate = mydate + 1
            If Format(mydate, "dddd") = "Saturday" Then mydate = DateAdd("d", 2, mydate)
        Case 3
            objCalRS.Fields(1) = mydate
            mydate = DateAdd("d", 1, mydate)
            If Format(mydate, "dddd") = "Monday" Then mydate = DateAdd("d", 5, mydate)
        Case 11
            objCalRS.Fields(1) = mydate
            mydate = DateAdd("yyyy", 1, mydate)
        Case Else
            While Format(mydate, "dddd") <> lstRepeatDays.List(lstRepeatDays.ListIndex)
                mydate = DateAdd("d", 1, mydate)
            Wend
            objCalRS.Fields(1) = mydate
            mydate = DateAdd("d", 7, mydate)
    End Select
    
    objCalRS.Fields(2) = Trim(txtCal(1))
    objCalRS.Fields(3) = Trim(txtCal(2))
    objCalRS.Fields(4) = Trim(txtCal(3))
    objCalRS.Fields(6) = Trim(txtCal(4))
    If optAlert(0) Then objCalRS.Fields(5) = True
    If optAlert(1) Then objCalRS.Fields(5) = False
    objCalRS.Update
Next
End Sub
Private Sub showCurrentItem()
Dim i As Integer
Cal.Value = objCalRS.Fields(1)
txtCal(1) = Format(objCalRS.Fields(2), "Hh:Nn")
txtCal(2) = objCalRS.Fields(3) & ""
txtCal(3) = objCalRS.Fields(4) & ""
txtCal(4) = objCalRS.Fields(6) & ""
If objCalRS.Fields(5) = True Then optAlert(0).Value = True
If objCalRS.Fields(5) = False Then optAlert(1).Value = True

End Sub

Private Sub Clear()
Dim i As Integer
For i = 1 To 4
txtCal(i).Text = ""
Next
txtCal(1).Text = "HH:MM"
optAlert(1).Value = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
objCalRS.Close
Set objCalRS = Nothing
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Me.Hide
End Sub

Private Sub lstcal_Click()
objCalRS.MoveFirst
objCalRS.Move (lstcal.ListIndex)
showCurrentItem
End Sub

Private Sub optAlert_Click(Index As Integer)
Select Case Index
    Case 0
        cmdEdit(4).Enabled = True
    Case 1
        cmdEdit(4).Enabled = False
End Select
End Sub

Private Sub txtRepeatTimes_Change()
If IsNumeric(txtRepeatTimes) = False Then
    txtRepeatTimes.Text = 1
    txtRepeatTimes.SelStart = Len(txtRepeatTimes)
ElseIf Val(txtRepeatTimes) > 99 Then
    txtRepeatTimes.Text = 99
    txtRepeatTimes.SelLength = Len(txtRepeatTimes)
ElseIf Val(txtRepeatTimes) = 0 Then
    txtRepeatTimes.Text = 1
    txtRepeatTimes.SelStart = Len(txtRepeatTimes)
End If

End Sub
