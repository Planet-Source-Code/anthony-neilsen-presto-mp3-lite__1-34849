Attribute VB_Name = "Module1"
Option Explicit
Global DoingScan As Boolean
Global ScanDir As Integer
Global ScanAccel As Single
Global ScanINK As Single
Global CurPosition As Double
Global MaxBal As Integer
Global MaxVol As Integer
Global CurHover As Integer

Global CurDuration As Double
Global CurTime As Double
Global CurSongTitle As String

Global MPvisible As Boolean
Global MPwindowON As Boolean

Global IsAMovie As Boolean
Global MovieIsVisible As Boolean





Global CurFile As String
Global LastFolder As String
Global CommandString As String
Global Secret As String

Global KillString As String * 300

Global IsRunning As Boolean
Global FormTOP As Integer
Global FormLEFT As Integer
Global FormWIDTH As Integer
Global FormHEIGHT As Integer

Global MovieTOP As Integer
Global MovieLEFT As Integer


Global GetOut As Boolean
Global Sval As Integer
Global IsOnTop As Integer
Global IsLoopMode As Integer

Global CurVolume As Integer
Global CurBalance As Integer

Global tWIN As New tpnWIN


Sub Main()
CommandString = Trim(Command)

Dim ff As Integer
ff = FreeFile
If App.PrevInstance Then
   Open App.Path & "\killprev.fle" For Random As ff Len = 300
   LSet KillString = CommandString
   Put #ff, 1, KillString
   Close ff
   End
   
End If
   
   


''''MsgBox (CommandString)


'CommandString = "D:\Mp3s\Soul Funk\Eternal - I Wanna Be The Only One.mp3"

Secret = Trim(CommandString)



frmMain.Show

End Sub




Function GetSongTitle(str7) As String
Dim i7 As Integer
For i7 = Len(str7) To 1 Step -1
   If Mid(str7, i7, 1) = "\" Then
      GetSongTitle = Mid(str7, i7 + 1, Len(str7) - i7 - 4)
      Exit Function
   End If
Next i7

End Function



Function GetFolder(str7) As String
Dim i7 As Integer
i7 = InStrRev(str7, "\")
If i7 = 0 Then
   GetFolder = ""
Else
   GetFolder = Left(str7, i7)
End If




End Function















Sub GetSettings()

Open App.Path & "\PrestoMP3lite.ini" For Random As #1
If LOF(1) = 0 Then
   Call MakeNewIni
End If
Close #1: Open App.Path & "\PrestoMP3lite.ini" For Input As #1
Input #1, FormTOP
Input #1, FormLEFT
Input #1, FormWIDTH
Input #1, FormHEIGHT
Input #1, LastFolder
Input #1, IsOnTop
Input #1, IsLoopMode
Input #1, CurBalance
Input #1, CurVolume
Input #1, MovieTOP
Input #1, MovieLEFT

Close #1


End Sub







Sub MakeNewIni()
LastFolder = "c:\"
FormTOP = 0
FormLEFT = 0
FormWIDTH = 3900
FormHEIGHT = 1065
MovieTOP = 750
MovieLEFT = 0

IsOnTop = 0
IsLoopMode = 0
CurBalance = 0
CurVolume = 0

Call SaveSettings

End Sub

Sub SaveSettings()

Close #1: Open App.Path & "\PrestoMP3lite.ini" For Output As #1
Print #1, FormTOP
Print #1, FormLEFT
Print #1, FormWIDTH
Print #1, FormHEIGHT
Print #1, LastFolder
Print #1, IsOnTop
Print #1, IsLoopMode
Print #1, CurBalance
Print #1, CurVolume
Print #1, MovieTOP
Print #1, MovieLEFT
Close #1



End Sub




