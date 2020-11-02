Attribute VB_Name = "WebEmail"
Option Explicit

Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

Public Function WebEmailOpen(UrlMailto As String) As Boolean
    WebEmailOpen = ShellExecute(&O0, "Open", UrlMailto, vbNullString, vbNullString, 4)
End Function

