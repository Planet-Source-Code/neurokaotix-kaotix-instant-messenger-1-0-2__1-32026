Attribute VB_Name = "dModule"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public YourSN As String
Public FinalClose As Boolean
'Public Function FileExists(sFullPath As String) As Boolean
'    Dim oFile As New Scripting.FileSystemObject
'    FileExists = oFile.FileExists(sFullPath)
'End Function
