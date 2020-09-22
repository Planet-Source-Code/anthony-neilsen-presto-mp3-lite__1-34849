VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   8790
   ClientLeft      =   3705
   ClientTop       =   1290
   ClientWidth     =   8685
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":1CFA
   ScaleHeight     =   8790
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   Begin PicClip.PictureClip BVclip 
      Left            =   2670
      Top             =   4785
      _ExtentX        =   900
      _ExtentY        =   582
      _Version        =   393216
      Picture         =   "frmMain.frx":A466
   End
   Begin VB.PictureBox VolSizer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   480
      ScaleHeight     =   360
      ScaleWidth      =   660
      TabIndex        =   9
      Top             =   1575
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox BalSizer 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   135
      ScaleHeight     =   615
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   1995
      Visible         =   0   'False
      Width           =   480
   End
   Begin PicClip.PictureClip OTLOOPclip 
      Left            =   6090
      Top             =   4620
      _ExtentX        =   1217
      _ExtentY        =   1746
      _Version        =   393216
      Picture         =   "frmMain.frx":ADA8
   End
   Begin PicClip.PictureClip ButtClip 
      Left            =   3510
      Top             =   4545
      _ExtentX        =   3810
      _ExtentY        =   1746
      _Version        =   393216
      Picture         =   "frmMain.frx":D212
   End
   Begin VB.Timer ScanTimer 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   4260
      Top             =   2355
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   210
      Index           =   1
      Left            =   1320
      TabIndex        =   7
      Top             =   2310
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   210
      Index           =   0
      Left            =   750
      TabIndex        =   6
      Top             =   2325
      Width           =   435
   End
   Begin VB.Timer KillTimer 
      Interval        =   500
      Left            =   5445
      Top             =   2385
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4860
      Top             =   2340
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3135
      TabIndex        =   2
      Top             =   1695
      Width           =   2295
      Begin VB.Label labCurTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   345
         TabIndex        =   4
         Top             =   150
         Width           =   360
      End
      Begin VB.Label labDuration 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   345
         TabIndex        =   3
         Top             =   -15
         Width           =   360
      End
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4395
      Width           =   1470
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   1185
      Top             =   4950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Image VolPic 
      Height          =   165
      Left            =   3060
      Picture         =   "frmMain.frx":141C4
      Top             =   135
      Width           =   255
   End
   Begin VB.Image BalPic 
      Height          =   165
      Left            =   150
      Picture         =   "frmMain.frx":14442
      Top             =   135
      Width           =   255
   End
   Begin VB.Image Butt 
      Height          =   330
      Index           =   9
      Left            =   2985
      Picture         =   "frmMain.frx":146C0
      Top             =   375
      Width           =   345
   End
   Begin VB.Image Butt 
      Height          =   330
      Index           =   8
      Left            =   2595
      Picture         =   "frmMain.frx":14D32
      Top             =   375
      Width           =   345
   End
   Begin VB.Image Butt 
      Height          =   330
      Index           =   7
      Left            =   2325
      Picture         =   "frmMain.frx":153A4
      Top             =   375
      Width           =   270
   End
   Begin VB.Image Butt 
      Height          =   330
      Index           =   6
      Left            =   2055
      Picture         =   "frmMain.frx":158B6
      Top             =   375
      Width           =   270
   End
   Begin VB.Image Butt 
      Height          =   330
      Index           =   5
      Left            =   1785
      Picture         =   "frmMain.frx":15DC8
      Top             =   375
      Width           =   270
   End
   Begin VB.Image Butt 
      Height          =   330
      Index           =   4
      Left            =   1515
      Picture         =   "frmMain.frx":162DA
      Top             =   375
      Width           =   270
   End
   Begin VB.Image Butt 
      Height          =   330
      Index           =   3
      Left            =   1140
      Picture         =   "frmMain.frx":167EC
      Top             =   375
      Width           =   270
   End
   Begin VB.Image Butt 
      Height          =   330
      Index           =   2
      Left            =   705
      Picture         =   "frmMain.frx":16CFE
      Top             =   375
      Width           =   270
   End
   Begin VB.Image Butt 
      Height          =   330
      Index           =   1
      Left            =   435
      Picture         =   "frmMain.frx":17210
      Top             =   375
      Width           =   270
   End
   Begin VB.Image Butt 
      Height          =   330
      Index           =   0
      Left            =   165
      Picture         =   "frmMain.frx":17722
      Top             =   375
      Width           =   270
   End
   Begin VB.Label labCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Presto Mp3 Lite"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   465
      TabIndex        =   5
      Top             =   120
      Width           =   2550
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer 
      CausesValidation=   0   'False
      Height          =   660
      Left            =   345
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3135
      Width           =   4575
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -80
      WindowlessVideo =   0   'False
   End
   Begin VB.Label labScan 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   390
      TabIndex        =   10
      Top             =   150
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub BalPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
   tWIN.SizeOBJ BalSizer, "R"
Else
   BalSizer.Width = MaxBal
End If
End Sub



Private Sub BalSizer_Resize()
Dim w As Integer
If GetOut Then Exit Sub
'Exit Sub
w = BalSizer.Width
If w > MaxBal * 2 Then
   w = MaxBal * 2
   GetOut = True
   BalSizer.Width = w
   GetOut = False
End If
CurBalance = w

Call SetCurBalance

End Sub
Sub SetCurBalance()
Dim b As Integer
If CurBalance = MaxBal Then
   BalPic.Picture = BVclip.GraphicCell(2)
Else
   BalPic.Picture = BVclip.GraphicCell(0)
End If

If frmMovie!MediaPlayer.ReadyState = 4 Or frmMovie!MediaPlayer.ReadyState = 3 Then
b = (CurBalance - MaxBal) * 20
frmMovie!MediaPlayer.Balance = b
End If

End Sub

Private Sub Butt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Pst As Integer

If Button = 1 Then

   Select Case Index
   Case 3  'open
      If Button = 1 Then
         Butt(Index).Picture = ButtClip.GraphicCell(Index + 16)
         Call QuickLoad("")
         Butt(Index).Picture = ButtClip.GraphicCell(Index)
      End If
   Case 8  'LOOPMODE
      Call ToggleLoopMode
      Exit Sub
   
   Case 9  'ONTOP
      Call ToggleOnTop
      Exit Sub
   
   Case Else
   End Select
   
   
   
   If frmMovie!MediaPlayer.ReadyState <> 4 Then Exit Sub
   Pst = frmMovie!MediaPlayer.PlayState
   
   
   Select Case Index
   Case 5, 6 'rewind ffwd
      Butt(Index).Picture = ButtClip.GraphicCell(Index + 16)
      Butt(Index).Refresh
      ScanDir = Index - 5
      ScanAccel = 1
      ScanINK = 0
      If frmMovie!MediaPlayer.PlayState = 2 Then
         frmMovie!MediaPlayer.Pause
      End If
      Call DoScan
      ScanTimer.Enabled = True
   
   Case Else
   End Select
   
   
   
   Select Case Pst
   Case 0 ' STOPPED
      Select Case Index
      Case 0  'play
         frmMovie!MediaPlayer.Play
   
      Case 1  'pause
         'frmMovie!MediaPlayer.Pause
      Case 2  'stop
         'frmMovie!MediaPlayer.Stop
         'frmMovie!MediaPlayer.CurrentPosition = 0
      Case Else
      End Select
   
   Case 1 ' PAUSED
      Select Case Index
      Case 0  'play
         frmMovie!MediaPlayer.Play
      Case 1  'pause
         'frmMovie!MediaPlayer.Pause
      Case 2  'stop
         frmMovie!MediaPlayer.Stop
         frmMovie!MediaPlayer.CurrentPosition = 0
      Case Else
      End Select
   
   Case 2 ' PLAYING
      Select Case Index
      Case 0  'play
         If IsAMovie Then
            If MPwindowON Then
               MPwindowON = False
               frmMovie.Hide
            Else
               MPwindowON = True
               frmMovie.Show
            End If
         End If
         
         'frmMovie!MediaPlayer.Play
      Case 1  'pause
         frmMovie!MediaPlayer.Pause
      Case 2  'stop
         frmMovie!MediaPlayer.Stop
         frmMovie!MediaPlayer.CurrentPosition = 0
      Case Else
      End Select
   
   
   
   
   Case Else
   
   End Select
   
   Text1.SetFocus

End If


End Sub

Private Sub Butt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Index <> CurHover Then
   CurHover = Index
   Call ShowCurHover
End If

End Sub

Sub ShowCurHover()
Dim NGstr As String
Dim i7 As Integer
Dim cTIME As String
Dim cDUR As String
Select Case CurHover
Case 4, 5, 6, 7
   If frmMovie!MediaPlayer.ReadyState = 4 Or frmMovie!MediaPlayer.ReadyState = 3 Then
      Timer1.Enabled = True
   End If
Case Else
   Timer1.Enabled = False
   If frmMovie!MediaPlayer.ReadyState = 4 Or frmMovie!MediaPlayer.ReadyState = 3 Then
      labCaption = CurSongTitle
   Else
      labCaption = "Presto Mp3 Lite"
   End If
End Select


If frmMovie!MediaPlayer.ReadyState <> 4 Then
   NGstr = "00010000"
Else
   Select Case frmMovie!MediaPlayer.PlayState
   Case 0: NGstr = "10011001"
   Case 1: NGstr = "10111001"
   Case 2: NGstr = "01111111"
   Case Else
      NGstr = "00010000"
   End Select
End If

For i7 = 0 To 7
   If Mid(NGstr, i7 + 1, 1) = "1" Then
      If i7 = CurHover Then
         Butt(i7).Picture = ButtClip.GraphicCell(i7 + 8)
      Else
         Butt(i7).Picture = ButtClip.GraphicCell(i7)
      End If
   End If
Next i7
NGstr = Mid(Str(IsLoopMode), 2) & Mid(Str(IsOnTop), 2)
For i7 = 8 To 9
   If Mid(NGstr, i7 - 7, 1) = "0" Then
      If i7 = CurHover Then
         Butt(i7).Picture = OTLOOPclip.GraphicCell(i7 - 6)
      Else
         Butt(i7).Picture = OTLOOPclip.GraphicCell(i7 - 8)
      End If
   End If
Next i7





End Sub





Private Sub Butt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
labScan.Visible = False

If Button = 2 Then
   Unload frmMain
   End
Else
   If Index <> 3 And frmMovie!MediaPlayer.ReadyState <> 4 Then Exit Sub

   Select Case Index
   Case 5, 6 'rewind ffwd
      ScanTimer.Enabled = False
      frmMovie!MediaPlayer.Play
   Case Else
   End Select

   
   If Index > 2 And Index < 8 Then
      Butt(Index).Picture = ButtClip.GraphicCell(Index + 8)
   Else
      Call ShowPlayStateButts
   End If
End If

End Sub
Sub ShowPlayStateButts()
Dim i7 As Integer
Dim k As Integer

Select Case frmMovie!MediaPlayer.PlayState
Case 0: k = 2
Case 1: k = 1
Case 2: k = 0
Case Else
k = -1
End Select
For i7 = 0 To 2
   If i7 = k Then
      Butt(i7).Picture = ButtClip.GraphicCell(i7 + 16)
   Else
      Butt(i7).Picture = ButtClip.GraphicCell(i7)
   End If
   
Next i7


End Sub




Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)


ScanDir = Index
ScanAccel = 1
ScanINK = 1
frmMovie!MediaPlayer.Pause
Call DoScan
ScanTimer.Enabled = True


End Sub

Private Sub Command1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
ScanTimer.Enabled = False
frmMovie!MediaPlayer.Play

End Sub

Sub QuickLoad(gSTR)
Dim isCMD As Boolean
Dim Tstr As String
isCMD = False
If gSTR <> "" Then
   isCMD = True
   Tstr = gSTR
   GoTo skipCDL
End If
On Error GoTo OpenErr 'catches error when user hits cancel
dlgCommon.Filter = "All Files (*.*)|*.*" 'sets the file type
dlgCommon.InitDir = LastFolder
dlgCommon.FileName = LastFolder & "*.*"
dlgCommon.ShowOpen


Tstr = dlgCommon.FileName
skipCDL:
LastFolder = GetFolder(Tstr)
CurFile = Tstr

If isCMD Then frmMovie!MediaPlayer.Stop
frmMovie!MediaPlayer.Open (Tstr)


Exit Sub


OpenErr:
If Err.Number = 32755 Then Exit Sub

dlgCommon.FileName = "D:\Mp3s\*.*"
MsgBox Err.Description & Err.Number
'MsgBox (LastFolder)
Resume

End Sub















Private Sub Form_Load()
Dim i7 As Integer

ButtClip.Rows = 3
ButtClip.Cols = 8

BVclip.Rows = 2
BVclip.Cols = 2
OTLOOPclip.Rows = 3
OTLOOPclip.Cols = 2


frmMain.Height = 51 * 15
frmMain.Width = 230 * 15
tWIN.SetRegion frmMain, 0


CurHover = -1
Call ShowCurHover

Call ShowPlayStateButts

Call GetSettings
frmMovie.Top = MovieTOP
frmMovie.Left = MovieLEFT


MaxVol = 450
MaxBal = 300

GetOut = True
BalSizer.Width = CurBalance
VolSizer.Width = MaxVol - CurVolume
GetOut = False

Call SetCurBalance
Call SetCurVolume

frmMain.Top = FormTOP
frmMain.Left = FormLEFT

If Secret <> "" Then
   Call QuickLoad(Secret)
End If

Call ShowLoopMode
Call ShowOnTop

End Sub
Sub ToggleOnTop()
If IsOnTop = 0 Then
   IsOnTop = 1
Else
   IsOnTop = 0
End If
Call ShowOnTop
End Sub

Sub ShowOnTop()
If MovieIsVisible Then
   tWIN.FormOnTop frmMovie, IsOnTop
End If
tWIN.FormOnTop frmMain, IsOnTop
If CurHover = 9 And IsOnTop = 0 Then
   Butt(9).Picture = OTLOOPclip.GraphicCell(3)
Else
   Butt(9).Picture = OTLOOPclip.GraphicCell(1 + IsOnTop * 4)
End If


End Sub


Sub ToggleLoopMode()

If IsLoopMode = 0 Then
   IsLoopMode = 1
Else
   IsLoopMode = 0
End If
Call ShowLoopMode

End Sub

Sub ShowLoopMode()
If CurHover = 8 And IsLoopMode = 0 Then
   Butt(8).Picture = OTLOOPclip.GraphicCell(2)
Else
   Butt(8).Picture = OTLOOPclip.GraphicCell(IsLoopMode * 4)
End If


End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
   tWIN.DragOBJ frmMain
End If

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)


If CurHover <> -1 Then
CurHover = -1
Call ShowCurHover
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
   Unload frmMain
   End
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

frmMovie!MediaPlayer.Stop

If frmMain.WindowState = 0 Then
   FormTOP = frmMain.Top
   FormLEFT = frmMain.Left
   FormWIDTH = frmMain.Width
   FormHEIGHT = frmMain.Height
   
   MovieTOP = frmMovie.Top
   MovieLEFT = frmMovie.Left
   
End If
Unload frmMovie

Call SaveSettings
Set frmMain = Nothing

End Sub


Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
   tWIN.DragOBJ frmMain
End If
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
   Unload frmMain
   End
End If

End Sub






Private Sub KillTimer_Timer()
Dim NewSong As String
Dim ff As Integer

KillTimer.Enabled = False


ff = FreeFile
Open App.Path & "\killprev.fle" For Random As ff Len = 300
LSet KillString = CommandString
Get #ff, 1, KillString
NewSong = Trim(KillString)
LSet KillString = ""
Put #ff, 1, KillString
Close ff

If NewSong <> "" Then
   frmMovie!MediaPlayer.Visible = False
   MPvisible = False
   DoEvents
  
   LastFolder = GetFolder(NewSong)
   CurFile = NewSong
   
   frmMovie!MediaPlayer.Stop
   frmMovie!MediaPlayer.Open (CurFile)
   Text1.SetFocus
End If

tWIN.FormOnTop frmMain, IsOnTop
KillTimer.Enabled = True



End Sub


Private Sub labCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
   tWIN.DragOBJ frmMain
End If
End Sub

Private Sub labCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
   Unload frmMain
   End
End If

End Sub







Private Sub ScanTimer_Timer()
Call DoScan

End Sub
Sub DoScan()
Dim amt As Integer

labScan.Width = 2655 * (CurPosition / CurDuration)
labScan.Visible = True
Select Case ScanDir
Case 0: amt = -1 * ScanAccel
Case 1: amt = 1 * ScanAccel
Case Else
End Select
CurPosition = CurPosition + amt
If CurPosition < 0 Then CurPosition = 0
If CurPosition > frmMovie!MediaPlayer.Duration - 0.5 Then
   CurPosition = frmMovie!MediaPlayer.Duration - 0.5
End If

frmMovie!MediaPlayer.CurrentPosition = CurPosition
'DoEvents
Select Case ScanINK
Case 0 To 5
   ScanAccel = ScanAccel + 1
Case 6 To 15
   ScanAccel = ScanAccel + 2
Case 16 To 20
   ScanAccel = ScanAccel + 5
Case Else
   ScanAccel = ScanAccel + 10
End Select

If ScanINK < 30 Then ScanINK = ScanINK + 1

 

End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If DoingScan Then Exit Sub
If KeyCode = 27 Then
   KeyCode = 0
   frmMovie!MediaPlayer.Stop
ElseIf KeyCode = 37 Then
   CurHover = 5
   Call ShowCurHover
   DoingScan = True
   Call Butt_MouseDown(5, 1, 0, 0, 0)
ElseIf KeyCode = 39 Then
   CurHover = 6
   Call ShowCurHover
   DoingScan = True
   Call Butt_MouseDown(6, 1, 0, 0, 0)
   
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim str7 As String
str7 = UCase(Chr(KeyAscii))
KeyAscii = 0

Select Case str7
Case "+"
   VolSizer.Width = VolSizer.Width + 15

Case "-"
   If VolSizer.Width > 15 Then
      VolSizer.Width = VolSizer.Width - 15
   End If



Case "O"    'open new song
   Butt(3).Picture = ButtClip.GraphicCell(3 + 16)
   Call QuickLoad("")
   Butt(3).Picture = ButtClip.GraphicCell(3)

Case " "     'toggle pause/play
   If frmMovie!MediaPlayer.PlayState = 0 Or frmMovie!MediaPlayer.PlayState = 1 Then
      frmMovie!MediaPlayer.Play
   ElseIf frmMovie!MediaPlayer.PlayState = 2 Then
      frmMovie!MediaPlayer.Pause
   End If
   
Case "S": frmMovie!MediaPlayer.Stop

Case "P": frmMovie!MediaPlayer.Play

Case Chr(24), Chr(3)
   Unload frmMain
   End

Case Else
End Select

End Sub



Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
DoingScan = False
If KeyCode = 37 Then
   Call Butt_MouseUp(5, 1, 0, 0, 0)
ElseIf KeyCode = 39 Then
   Call Butt_MouseUp(6, 1, 0, 0, 0)
   
End If
CurHover = -1
Call ShowCurHover

End Sub

Private Sub Timer1_Timer()
Dim cTIME As String
Dim cDUR As String

CurTime = frmMovie!MediaPlayer.CurrentPosition
cDUR = GetHMS(CurDuration)
cTIME = GetHMS(CurTime)
labCaption = cDUR & "   >   " & cTIME




End Sub









Function GetHMS(n) As String
Dim h As Integer
Dim m As Integer
Dim s As Integer
Dim s2 As Single
h = Int(n / 3600)
m = Int(n / 60)
s = Int(n - (h * 3600) - (m * 60))
s2 = Int((n - s) * 10)
If n < 3600 Then
   GetHMS = Right("00" & Mid(Str(m), 2), 2) & ":" & Right("00" & Mid(Str(s), 2), 2) ' & ":" & Mid(Str(s2), 2)
Else
   GetHMS = Mid(Str(h), 2) & ":" & Right("00" & Mid(Str(m), 2), 2) & ":" & Right("00" & Mid(Str(s), 2), 2) ' & ":" & Mid(Str(s2), 2)
End If



End Function

Private Sub VolPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
   tWIN.SizeOBJ VolSizer, "R"
Else
   VolSizer.Width = MaxVol
   
End If
End Sub



Private Sub VolSizer_Resize()
Dim w As Integer
If GetOut Then Exit Sub
'Exit Sub
w = VolSizer.Width
If w > MaxVol Then
   w = MaxVol
   GetOut = True
   VolSizer.Width = w
   GetOut = False
End If
CurVolume = MaxVol - w

Call SetCurVolume


End Sub


Sub SetCurVolume()
Dim v As Integer
If CurVolume = 0 Then
   VolPic.Picture = BVclip.GraphicCell(3)
Else
   VolPic.Picture = BVclip.GraphicCell(1)
End If
If frmMovie!MediaPlayer.ReadyState = 4 Or frmMovie!MediaPlayer.ReadyState = 3 Then
v = CurVolume * -18
frmMovie!MediaPlayer.Volume = v
End If

End Sub
