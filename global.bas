Attribute VB_Name = "global"
Option Explicit

Public strLastPath As String

Public Sub LoadSettings()

  strLastPath = GetSetting(App.Title, "Settings", "Path", App.Path)
  strLastPath = getPath

End Sub

Public Sub SaveSettings()

  Call SaveSetting(App.Title, "Settings", "Path", getPath)

End Sub

Private Function getPath() As String

'--> Get Path information out of full filename
Dim intPos As Integer

  intPos = InStrRev(strLastPath, "\")
  getPath = Mid(strLastPath, 1, intPos)

End Function
