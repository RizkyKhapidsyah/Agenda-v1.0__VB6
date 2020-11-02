VERSION 5.00
Begin VB.Form frmCalRead 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Todays Calendar events"
   ClientHeight    =   3900
   ClientLeft      =   7230
   ClientTop       =   5295
   ClientWidth     =   4020
   Icon            =   "frmCalRead.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   55000
      Left            =   3480
      Top             =   1440
   End
   Begin VB.TextBox txtMemo 
      Height          =   2085
      Left            =   60
      Locked          =   -1  'True
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      ToolTipText     =   "Memo for the event"
      Top             =   1740
      Width           =   3855
   End
   Begin VB.ListBox lstcal 
      Height          =   1230
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "The events of the selected day"
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Memo"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Calendar Events"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmCalRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListEvents()
Dim i As Integer
lstcal.Clear
txtMemo.Text = ""
If NoOfAlerts > 0 Then
For i = 1 To NoOfAlerts
lstcal.AddItem Format(Warning(i).Time, "Hh:Nn") & " " & Warning(i).Description
Next
lstcal.ListIndex = 0
End If
End Sub
Private Sub Form_Load()
Me.Show
GetTodaysAlerts
ListEvents
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Me.Hide
End Sub

Private Sub lstcal_Click()
txtMemo.Text = Warning(lstcal.ListIndex + 1).Memo
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
GetTodaysAlerts
ListEvents

If Format(Time, "Hh:Nn") = Format(#12:01:00 AM#, "Hh:Nn") Then
    GetTodaysAlerts
    ListEvents
    Exit Sub
End If
If NoOfAlerts > 0 Then
    For i = 1 To NoOfAlerts
        If Format(Warning(i).Time, "Hh:Nn") = Format(Time, "Hh:Nn") Then
            Me.WindowState = 0
            Me.Show
            lstcal.ListIndex = (Warning(i).Index)
            If Warning(i).SoundPath <> "" Then
                Call RunFile(Warning(i).SoundPath, Me)
                Warning(i).SoundPath = ""
            End If
        End If
    Next
End If
End Sub
