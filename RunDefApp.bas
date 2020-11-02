Attribute VB_Name = "RunDefApp"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10


Public Sub RunFile(path As String, SenderForm As Form)
If path <> "" Then
    ShellExecute SenderForm.hwnd, vbNullString, path, vbNullString, "C:\", SW_SHOWMINIMIZED
End If
End Sub

'***Help about Function RunFile***'
' If Windows has a default program associated with the file it will run in that program
' The Parameter Path is the full Path to the file you want to run
' The Parameter SenderForm is the Calling Container for ex Form1 or ME

