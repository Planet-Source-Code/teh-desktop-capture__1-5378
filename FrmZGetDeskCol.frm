VERSION 5.00
Begin VB.Form FrmZGetDeskCol 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Desktop"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3225
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   77
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.PictureBox PicColor 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   2760
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox PicDesk 
      AutoRedraw      =   -1  'True
      Height          =   735
      Left            =   1200
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton CmdTimer 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   3240
      Top             =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Press Space bar  or <enter> to triggle"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2625
   End
End
Attribute VB_Name = "FrmZGetDeskCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Author  : TEH   mailto: tehmax@cyberdude.com                            '
'                                                                         '
' Project :  Capture desktop color and picture from mouse cursor position '
'            and save picture to App.Path & "\TempImg" & i% & ".bmp" .    '
'                                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim j%


Private Sub CmdSave_Click()
On Error Resume Next
Static i%
Dim FileSelect$
i% = i% + 1
j% = i%
FileSelect$ = App.Path & "\TempImg" & i% & ".bmp"
PicDesk.Picture = PicDesk.Image
SavePicture PicDesk.Picture, FileSelect$
End Sub

Private Sub CmdTimer_Click()
Timer1.Enabled = Not Timer1.Enabled
If Timer1.Enabled Then
    CmdTimer.Caption = "&Stop"
Else
    CmdTimer.Caption = "&Start"
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim ret%, k%
On Error Resume Next
If j% Then ret = MsgBox("Kill save picture ?", vbCritical + vbYesNo)
If ret = vbYes Then
    For k% = 1 To j%
        Kill App.Path & "\TempImg" & k% & ".bmp"
    Next k%
End If


End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then
    PicDesk.Height = Me.ScaleHeight - PicDesk.Top - 10
    PicDesk.Width = Me.ScaleWidth - PicDesk.Left - 10
End If
End Sub

Private Sub Timer1_Timer()
PicColor.BackColor = GetDcColor
GetDcPic
End Sub

Public Function GetDcPic() As Long
Dim DeskHdc&, ret&
Dim Pxy As POINTAPI
    ' Get Desktop DC
    DeskHdc = GetDC(0)
    'Get mouse position
    GetCursorPos Pxy
    GetDcPic = BitBlt(PicDesk.hdc, 0, 0, PicDesk.Width, PicDesk.Height, DeskHdc, Pxy.X, Pxy.Y, vbSrcCopy)  'GetCursorPos(Pxy.X), GetCursorPos(Pxy.Y))
    ret = ReleaseDC(0&, DeskHdc)
    PicDesk.Refresh
End Function

Public Function GetDcColor() As Double
Dim DeskHdc&, ret&
Dim Pxy As POINTAPI
    ' Get Desktop DC
    DeskHdc = GetDC(0)
    'Get mouse position
    GetCursorPos Pxy
    GetDcColor = GetPixel(DeskHdc, Pxy.X, Pxy.Y)  'GetCursorPos(Pxy.X), GetCursorPos(Pxy.Y))
    ret& = ReleaseDC(0&, DeskHdc)
End Function

