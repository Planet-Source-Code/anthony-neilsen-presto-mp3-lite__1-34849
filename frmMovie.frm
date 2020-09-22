VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmMovie 
   Appearance      =   0  'Flat
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   ClientHeight    =   525
   ClientLeft      =   1830
   ClientTop       =   1455
   ClientWidth     =   2415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMovie.frx":0000
   ScaleHeight     =   525
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer 
      CausesValidation=   0   'False
      Height          =   225
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   3090
      AudioStream     =   -1
      AutoSize        =   -1  'True
      AutoStart       =   -1  'True
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
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
      DisplaySize     =   0
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
      TransparentAtStart=   -1  'True
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -80
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frmMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_GotFocus()
frmMain!Text1.SetFocus
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
   tWIN.DragOBJ frmMovie
Else
'Stop
   If MediaPlayer.DisplaySize = 2 Then
      MediaPlayer.DisplaySize = 0
   Else
      MediaPlayer.DisplaySize = 2
   End If
   Call SyncWindowSize
   'frmMovie.Hide
   'MovieIsVisible = False
End If
frmMain!Text1.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmMovie = Nothing

End Sub

Private Sub MediaPlayer_DisplayModeChange()
Call SyncWindowSize

End Sub

Private Sub MediaPlayer_GotFocus()
frmMain!Text1.SetFocus

End Sub

Private Sub MediaPlayer_MouseDown(Button As Integer, ShiftState As Integer, x As Single, y As Single)

frmMain!Text1.SetFocus

End Sub




Private Sub MediaPlayer_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
Call frmMain.ShowPlayStateButts
If OldState = 2 And NewState = 0 Then
   If Not IsAMovie Then
      If IsLoopMode = 1 Then
         If MediaPlayer.CurrentPosition = MediaPlayer.Duration Then
         MediaPlayer.Play
         End If
      End If
   End If

End If

   
End Sub

Private Sub MediaPlayer_PositionChange(ByVal oldPosition As Double, ByVal newPosition As Double)
CurPosition = newPosition

End Sub

Private Sub MediaPlayer_ReadyStateChange(ReadyState As MediaPlayerCtl.ReadyStateConstants)
If ReadyState = 4 Then
   If Not MPvisible Then
      If CurDuration > 0.5 Then
         frmMovie!MediaPlayer.Visible = True
         MPvisible = True
      End If
      Call SyncWindowSize
   End If
   CurDuration = MediaPlayer.Duration
   CurPosition = 0
   CurSongTitle = GetSongTitle(CurFile)
   frmMain!labCaption = CurSongTitle
   
   Call frmMain.SetCurBalance
   Call frmMain.SetCurVolume
   
End If


End Sub

Sub SyncWindowSize()

frmMovie.Width = MediaPlayer.Width
If MediaPlayer.Height > 825 Then
   IsAMovie = True
   frmMovie.Height = MediaPlayer.Height - 555
   frmMovie.Show
   tWIN.FormOnTop frmMovie, IsOnTop
   MovieIsVisible = True
   MPwindowON = True
Else
   IsAMovie = False
   tWIN.FormOnTop frmMovie, 0
   frmMovie.Hide
   MovieIsVisible = False
   MPwindowON = False
End If

End Sub
