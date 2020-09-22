VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tpnWIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub DragOBJ(frm)

ReleaseCapture
SendMessage frm.hwnd, &HA1, 2, 0&

End Sub
Sub SizeOBJ(f, act)
Dim act2 As Integer
Select Case UCase(act)
Case "T": act2 = 12
Case "B": act2 = 15
Case "L": act2 = 10
Case "R": act2 = 11
Case "TL": act2 = 13
Case "TR": act2 = 14
Case "BL": act2 = 16
Case "BR": act2 = 17
Case Else
Exit Sub
End Select

Call ReleaseCapture
Call SendMessage(f.hwnd, &HA1, act2, 0)

End Sub


Sub FormOnTop(f, act)
Select Case act

Case 1
   SetWindowPos f.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE Or SWP_NOMOVE
Case 0
   SetWindowPos f.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE Or SWP_NOMOVE
Case Else

End Select

End Sub


Sub SetRegion(ff, cc As Long)
Dim hRgn As Long
hRgn = GetBitmapRegion(ff.Picture, cc)
SetWindowRgn ff.hwnd, hRgn, True

End Sub

