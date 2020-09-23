VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   630
   ClientLeft      =   1665
   ClientTop       =   3465
   ClientWidth     =   1560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   42
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   104
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Click()

CreateGame 1     'Initialize game
frmMain.Show    'Show window
Unload Me

End Sub

Private Sub tmrEffect_Timer()
Dim nX As Long, nY As Long, Cnt As Long
Dim Delay As Long, T1 As Long, T2 As Long
Dim Start As Byte, nSteps As Byte
Dim Ret As POINTAPI

Delay = 0: nSteps = 3

For Start = 0 To nSteps - 1
    T1 = GetTickCount
    T2 = GetTickCount
    For Cnt = Start To frmSplash.ScaleWidth Step nSteps
        While T2 - T1 < Delay
            T2 = GetTickCount
        Wend
        nX = Cnt
        nY = Me.ScaleWidth
        frmSplash.Line (Cnt, 0)-(Cnt, Me.ScaleHeight), 0
        frmSplash.Line (0, Cnt)-(Me.ScaleWidth, Cnt), 0
        T1 = GetTickCount
    Next Cnt
Next Start

CreateGame 1     'Initialize game
frmMain.Show    'Show window
Unload Me

End Sub


