VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Knoton´s Agenda"
   ClientHeight    =   7695
   ClientLeft      =   6345
   ClientTop       =   1590
   ClientWidth     =   12660
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   12660
   Begin MSComDlg.CommonDialog CDCreateOpen 
      Left            =   300
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Adress Register (*.adr)|*.adr"
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   2700
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu SysTrayMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuMenu 
         Caption         =   "Show Menu"
      End
      Begin VB.Menu mnuOpenAgenda 
         Caption         =   "Open Agenda"
      End
      Begin VB.Menu mnuCloseAgenda 
         Caption         =   "Close Agenda"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCreateAgenda 
         Caption         =   "Create Agenda"
      End
      Begin VB.Menu mnuBackupAgenda 
         Caption         =   "Backup Agenda"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRestoreAgenda 
         Caption         =   "Restore Agenda"
      End
      Begin VB.Menu mnuOpenContactsbook 
         Caption         =   "Show Contactsbook"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCloseContactsbook 
         Caption         =   "Hide Contactsbook"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpenCalendar 
         Caption         =   "Show Calendar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCloseCalendar 
         Caption         =   "Hide Calendar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditCalendar 
         Caption         =   "Show Edit Calendar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHideEditCalendar 
         Caption         =   "Hide Edit Calendar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuDevWeb 
         Caption         =   "Developers Website"
      End
      Begin VB.Menu mnuDevMail 
         Caption         =   "Developers mail"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If Command <> "" Then
        AdressRegisterPath = Command
        LoadAgenda
Else
Associate "Agenda", ".adr", "Agenda File", App.path & "\BOOK02.ICO"
    End If
Call InitSystray(Me)
    Me.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim IconEvent As Long
    IconEvent = X / Screen.TwipsPerPixelX

Select Case IconEvent
  Case WM_LBUTTONUP
    SetForegroundWindow Me.hwnd
    Me.PopupMenu Me.SysTrayMenu
  Case WM_RBUTTONUP
    SetForegroundWindow Me.hwnd
    Me.PopupMenu Me.SysTrayMenu
End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
UnloadForms
Call CloseSystray

End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Me.Hide
End Sub

Private Sub mnuBackupAgenda_Click()
Dim strTemp As String
Dim i As Integer

mnuCloseAgenda_Click
CDCreateOpen.DialogTitle = "Where do you want to put your backup ?"
CDCreateOpen.FileName = ""
    For i = 1 To Len(AdressRegisterPath) - 1
        If Mid(AdressRegisterPath, i, 1) = "\" Then
        strTemp = Mid(AdressRegisterPath, 1, i)
        End If
    Next
CDCreateOpen.FileName = Mid(AdressRegisterPath, Len(strTemp) + 1)
CDCreateOpen.ShowSave

If CDCreateOpen.FileName <> "" Then FileCopy AdressRegisterPath, CDCreateOpen.FileName
CDCreateOpen.FileName = ""
LoadAgenda
End Sub

Private Sub mnuCloseAgenda_Click()
UnloadForms
CloseAllConRS
HideHiddenMenues
End Sub

Private Sub mnuCloseCalendar_Click()
frmCalRead.Hide
End Sub

Private Sub mnuCloseContactsbook_Click()
Contactsbook.Hide
End Sub

Private Sub mnuCreateAgenda_Click()
CDCreateOpen.InitDir = App.path
CDCreateOpen.DialogTitle = "Create Agenda as"
CDCreateOpen.ShowSave
If CDCreateOpen.FileName <> "" Then
FileCopy App.path & "\TEMPLATE.bak", CDCreateOpen.FileName
AdressRegisterPath = CDCreateOpen.FileName
End If
End Sub

Private Sub mnuDevMail_Click()
WebEmailOpen "mailto:knoton@hotmail.com"
End Sub

Private Sub mnuDevWeb_Click()
WebEmailOpen ("http://www.knoton.dns2go.com")
End Sub

Private Sub mnuEditCalendar_Click()
frmCal.WindowState = 0
frmCal.Show
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHideEditCalendar_Click()
frmCal.Hide
End Sub

Private Sub mnuMenu_Click()
Me.WindowState = 0
Me.Show
End Sub

Private Sub mnuOpenAgenda_Click()
CDCreateOpen.InitDir = App.path
CDCreateOpen.DialogTitle = "Open agenda"
CDCreateOpen.FileName = ""
CDCreateOpen.ShowOpen
If CDCreateOpen.FileName <> "" Then
    AdressRegisterPath = CDCreateOpen.FileName
    CDCreateOpen.FileName = ""
    LoadAgenda
End If
End Sub

Private Sub mnuOpenCalendar_Click()
frmCalRead.WindowState = 0
frmCalRead.Show
End Sub

Private Sub mnuOpenContactsbook_Click()
Contactsbook.WindowState = 0
Contactsbook.Show
End Sub

Private Sub mnuRestoreAgenda_Click()
Dim strTemp As String
Dim i As Integer

CDCreateOpen.DialogTitle = "Select Agenda to restore"
CDCreateOpen.ShowOpen
CDCreateOpen.FileName = ""
If CDCreateOpen.FileName <> "" Then
AdressRegisterPath = CDCreateOpen.FileName

    For i = 1 To Len(AdressRegisterPath) - 1
        If Mid(AdressRegisterPath, i, 1) = "\" Then
        strTemp = Mid(AdressRegisterPath, 1, i)
        End If
    Next
strTemp = "\" & Mid(AdressRegisterPath, Len(strTemp) + 1)
FileCopy CDCreateOpen.FileName, App.path & strTemp
CDCreateOpen.FileName = ""
End If

End Sub

Private Sub ShowHiddenMenues()
mnuOpenCalendar.Visible = True
mnuEditCalendar.Visible = True
mnuOpenContactsbook.Visible = True
mnuCloseContactsbook.Visible = True
mnuCloseCalendar.Visible = True
mnuCloseAgenda.Visible = True
mnuHideEditCalendar.Visible = True
mnuBackupAgenda.Visible = True
mnuCreateAgenda.Visible = False
mnuOpenAgenda.Visible = False
mnuRestoreAgenda.Visible = False
End Sub

Private Sub HideHiddenMenues()
mnuOpenCalendar.Visible = False
mnuEditCalendar.Visible = False
mnuOpenContactsbook.Visible = False
mnuCloseContactsbook.Visible = False
mnuCloseCalendar.Visible = False
mnuHideEditCalendar.Visible = False
mnuBackupAgenda.Visible = False
mnuCloseAgenda.Visible = False
mnuCreateAgenda.Visible = True
mnuOpenAgenda.Visible = True
mnuRestoreAgenda.Visible = True
End Sub
Private Sub LoadAgenda()
Set objCon = New ADODB.Connection
Load frmCalRead
Load Contactsbook
Load frmCal
frmCal.Hide
Contactsbook.Hide
ShowHiddenMenues

End Sub
Private Sub UnloadForms()
Unload frmCalRead
Unload Contactsbook
Unload frmCal
End Sub
Private Sub CloseAllConRS()
Set objCalRS = Nothing
Set objRs = Nothing
Set objCon = Nothing

End Sub
