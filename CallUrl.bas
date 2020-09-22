Attribute VB_Name = "CallUrl"
Option Explicit

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub Call_Url(obj As Object, url As String)
    'url="http://www..."
    'url="mailto:a@b.net"
    Call ShellExecute(obj.hwnd, "Open", url, "", "", 1)
End Sub
