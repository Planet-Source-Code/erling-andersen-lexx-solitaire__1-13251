VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lexx Solitaire - By R.Ling"
   ClientHeight    =   1110
   ClientLeft      =   1695
   ClientTop       =   1845
   ClientWidth     =   4755
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   74
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   Visible         =   0   'False
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   855
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Panel 1"
            TextSave        =   "Panel 1"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Text            =   "Panel 2"
            TextSave        =   "Panel 2"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrUpdateTime 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer tmr790Blink 
      Enabled         =   0   'False
      Interval        =   1750
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox imgLoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Menu mnuTopFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewgame 
         Caption         =   "&New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuFile_Sep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTopOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsLayout 
         Caption         =   "&Layout"
         Begin VB.Menu mnuOptionsLayoutBg 
            Caption         =   "&Background"
            Begin VB.Menu mnuOptionsLayoutBgBg 
               Caption         =   "Background 1"
               Index           =   0
            End
         End
         Begin VB.Menu mnuOptionsLayout_Sep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsLayoutFront 
            Caption         =   "&Card Front"
            Begin VB.Menu mnuOptionsLayoutFrontFront 
               Caption         =   "Front 1"
               Index           =   0
            End
         End
         Begin VB.Menu mnuOptionsLayoutBack 
            Caption         =   "&Card Back"
            Begin VB.Menu mnuOptionsLayoutBackBack 
               Caption         =   "Back 1"
               Index           =   0
            End
         End
      End
      Begin VB.Menu mnuOptions_Sep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsLanguage 
         Caption         =   "&Language"
         Begin VB.Menu mnuOptionsLanguageEnglish 
            Caption         =   "&English"
         End
         Begin VB.Menu mnuLanguage_Sep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsLanguageNorwegian 
            Caption         =   "&Norwegian"
         End
         Begin VB.Menu mnuOptionsLanguageGerman 
            Caption         =   "&German"
         End
      End
      Begin VB.Menu mnuOptionsSounds 
         Caption         =   "&Sounds"
         Begin VB.Menu mnuOptionsSoundsDefsnd 
            Caption         =   "Ordinary Move"
            Begin VB.Menu mnuOptionsSoundsDefsndSnd 
               Caption         =   "Sound 1"
               Index           =   0
            End
         End
         Begin VB.Menu mnuOptionsSoundsGoalsnd 
            Caption         =   "Move To Goal"
            Begin VB.Menu mnuOptionsSoundsGoalsndSnd 
               Caption         =   "Sound 1"
               Index           =   0
            End
         End
         Begin VB.Menu mnuOptionsSounds_Sep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsSoundsPlaysounds 
            Caption         =   "&Play Sounds"
         End
      End
      Begin VB.Menu mnuOptionsRules 
         Caption         =   "&Rules"
         Begin VB.Menu mnuOptionsRulesDraw 
            Caption         =   "Draw.."
            Begin VB.Menu mnuOptionsRulesDrawOne 
               Caption         =   "&One Card"
            End
            Begin VB.Menu mnuOptionsRulesDrawThree 
               Caption         =   "&Three Cards"
            End
         End
      End
      Begin VB.Menu mnuOptions_Sep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions790 
         Caption         =   "Show &790"
      End
   End
   Begin VB.Menu mnuTopDebug 
      Caption         =   "&Debug"
      Begin VB.Menu mnuDebugWin 
         Caption         =   "Win Game"
      End
      Begin VB.Menu mnuDebug_Sep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDebugNoshuffle 
         Caption         =   "No Shuffle Game"
      End
      Begin VB.Menu mnuDebugResetsettings 
         Caption         =   "Reset Settings"
      End
      Begin VB.Menu mnuDebug_Sep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDebugDisabledebug 
         Caption         =   "Disable Debugging"
      End
   End
   Begin VB.Menu mnuTopHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Index"
      End
      Begin VB.Menu mnuHelp_Sep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About.."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub TriggerClick(X As Long, Y As Long)

Form_MouseDown 1, 0, CSng(X), CSng(Y)

End Sub


Private Sub Form_DblClick()
If Info.MouseButton <> 1 Then Exit Sub

CheckDoubleClick Info.ACPos.X, Info.ACPos.Y

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static Buffer As String

Info.Interrupt = True

Buffer = Buffer & Chr$(KeyCode)

If InStr(1, Buffer, "enter debug mode", vbTextCompare) > 0 Then
    'MsgBox Buffer
    Buffer = vbNullString
    mnuTopDebug.Visible = True
    Info.Debugging = True
End If

If InStr(1, Buffer, "exit debug mode", vbTextCompare) > 0 Then
    'MsgBox Buffer
    Buffer = vbNullString
    mnuTopDebug.Visible = False
    Info.Debugging = False
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Info.MouseButton = Button
If Button = 2 And Not Info.Moving Then
    Info.DblClickPos.X = X
    Info.DblClickPos.Y = Y
End If
If Button <> 1 Or Info.Moving Then Exit Sub

If Not Info.Interrupt Then
    Info.Interrupt = True
    Exit Sub
End If

Info.ACPos.X = X
Info.ACPos.Y = Y
If CheckMouseDown_Deck(X, Y) Then Exit Sub
If CheckMouseDown_Placeholders(X, Y) Then Exit Sub
If CheckMouseDown_GoalCells(X, Y) Then Exit Sub

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

If Not Info.Moving Then Exit Sub
DrawActiveCard X, Y

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If Not CheckDoubleClick(Info.DblClickPos.X, Info.DblClickPos.Y) Then
        'If the user didn't click on a valid card, then make the options menu pop up:
        PopupMenu mnuTopOptions
    End If
End If

If Button <> 1 Then Exit Sub

If CheckMouseUp_Placeholders(X, Y) Then GoTo WasPlaced
If CheckMouseUp_GoalCells(X, Y) Then GoTo WasPlaced

'If the card wasn't placed anywhere, put it back where it came from:
If Info.Moving Then MoveBack

Exit Sub
WasPlaced:

If Info.srcType = 3 Then
    If Info.nRemoved = 3 Then Info.nRemoved = 0
    'Info.nActive = 1
    'DrawDeck GameField.BB.hDC
    'DrawDeck frmMain.hDC
    'Info.nActive = 0
End If

End Sub

Private Sub Form_Paint()

'DrawGameField
FlipScreen

End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1

'Free Resources:
CleanUp

'Terminate App:
End

End Sub





Private Sub mnuDebugDisabledebug_Click()

Info.Debugging = False
mnuTopDebug.Visible = False
frmMain.Caption = cWINDOWCAPTION
SaveSettings

End Sub

Private Sub mnuDebugDoeffect_Click()

DoEffect

End Sub



Private Sub mnuDebugNoshuffle_Click()

Init_Cards
Init_Placeholders
CreateGame 0

End Sub





Private Sub mnuDebugReadsettings_Click()

ReadSettings

GameField.TmpBG.fPath = Info.AppPath & "BG\BG" & Trim(Str(Info.nBG + 1)) & ".BMP"
FreeRes GameField.TmpBG
LoadRes GameField.TmpBG
MakeBG

frmMain.Caption = cWINDOWCAPTION
DrawGameField

End Sub

Private Sub mnuDebugRefresh_Click()

DrawGameField

End Sub

Private Sub mnuDebugResetsettings_Click()

Info.nCardFront = 0
Info.nCardBack = 0

Info.nBG = 0
GameField.TmpBG.fPath = Info.AppPath & "BG\BG1.BMP"
FreeRes GameField.TmpBG
LoadRes GameField.TmpBG
MakeBG

'Refresh Gamefield:
DrawGameField

End Sub

Private Sub mnuDebugSavesettings_Click()

SaveSettings

End Sub

Private Sub mnuDebugWin_Click()
Dim Cnt As Byte

For Cnt = 1 To 4
    GameField.GoalCell(Cnt).nCards = 13
Next

CheckWinner

End Sub

Private Sub mnuFileExit_Click()

CleanUp
End

End Sub

Private Sub mnuFileNewgame_Click()

If MsgBox(cNEWGAMEMSG, vbQuestion Or vbYesNo, cNEWGAMECAPTION) = 6 Then
    CreateGame 1
End If

End Sub

Private Sub mnuFileUndo_Click()

Undo

End Sub

Private Sub mnuHelpAbout_Click()
Dim Logo As tResource, X As Single, Y As Single
Dim XSpeed As Long, YSpeed As Long
Dim nTilesX As Integer, OffSetX As Long, Cnt As Byte
Dim nTilesY As Integer, OffSetY As Long, Cnt2 As Byte
Dim Stars() As POINTAPI
Dim StarCols() As Long
Dim StarXSpeed As Long, StarYSpeed As Long
Dim XDir As Integer, YDir As Integer

Info.Interrupt = False
Logo.fPath = Info.AppPath & "LOGO.BMP"
LoadRes Logo


nTilesX = (GameField.nW \ Logo.nW) + 1
nTilesY = (GameField.nH \ Logo.nH) + 1
If (nTilesX - 1) * Logo.nW < GameField.nW Then nTilesX = nTilesX + 1
If (nTilesY - 1) * Logo.nH < GameField.nH Then nTilesY = nTilesY + 1
OffSetX = 0
OffSetY = 0

ReDim Stars(1 To 200): ReDim StarCols(1 To 200)

StarXSpeed = -3
StarYSpeed = 3
XDir = -1
YDir = 1

For Cnt = 1 To UBound(Stars)
    Stars(Cnt).X = (Rnd * GameField.nW) - (GameField.nW * XDir)
    Stars(Cnt).Y = (Rnd * GameField.nH) - (GameField.nW * YDir)
    StarCols(Cnt) = Rnd * 16777215
Next Cnt

While Not Info.Interrupt
    For Cnt = 0 To nTilesY
        For Cnt2 = 0 To nTilesX
            BitBlt GameField.BB.hDC, OffSetX + ((Cnt - 1) * Logo.nW), OffSetY + ((Cnt2 - 1) * Logo.nH), Logo.nW, Logo.nH, Logo.hDC, 0, 0, SRCCOPY
        Next Cnt2
    Next Cnt
    'For Cnt = 1 To UBound(Stars)
    '    Stars(Cnt).X = Stars(Cnt).X + StarXSpeed
    '    Stars(Cnt).Y = Stars(Cnt).Y + StarYSpeed
    '    If StarXSpeed > 0 And Stars(Cnt).X > GameField.nW Then
    '        Stars(Cnt).X = (Rnd * GameField.nW)
    '        Stars(Cnt).Y = (Rnd * GameField.nH)
    '    ElseIf StarXSpeed < 0 And Stars(Cnt).X < 0 Then
    '        Stars(Cnt).X = (Rnd * GameField.nW)
    '        Stars(Cnt).Y = (Rnd * GameField.nH)
    '    End If
    '    If StarYSpeed > 0 And Stars(Cnt).Y > GameField.nH Then
    '        Stars(Cnt).X = (Rnd * GameField.nW)
    '        Stars(Cnt).Y = (Rnd * GameField.nH)
    '    ElseIf StarYSpeed < 0 And Stars(Cnt).Y < 0 Then
    '        Stars(Cnt).X = (Rnd * GameField.nW)
    '        Stars(Cnt).Y = (Rnd * GameField.nH)
    '    End If
    '    SetPixel GameField.BB.hDC, Stars(Cnt).X + 1, Stars(Cnt).Y, StarCols(Cnt)
        'SetPixel GameField.BB.hDC, Stars(Cnt).X + 1, Stars(Cnt).Y + 1, StarCols(Cnt)
        'SetPixel GameField.BB.hDC, Stars(Cnt).X + 1, Stars(Cnt).Y + 2, StarCols(Cnt)
        'SetPixel GameField.BB.hDC, Stars(Cnt).X, Stars(Cnt).Y + 1, StarCols(Cnt)
        'SetPixel GameField.BB.hDC, Stars(Cnt).X + 2, Stars(Cnt).Y + 1, StarCols(Cnt)
    'Next Cnt
FlipScreen
OffSetX = OffSetX + 2
OffSetY = OffSetY - 1
If OffSetX > Logo.nW Then OffSetX = 1
If OffSetX < -Logo.nW Then OffSetX = -1
If OffSetY > Logo.nH Then OffSetY = 1
If OffSetY < -Logo.nH Then OffSetY = -1
DoEvents
Wend

FreeRes Logo
DrawGameField

End Sub

Private Sub mnuOptions790_Click()

Info.Show790 = Not Info.Show790

If Info.Show790 Then
    DrawGameField
    tmr790Blink.Enabled = True
Else
    tmr790Blink.Enabled = False
    DrawGameField
End If

End Sub

Private Sub mnuOptionsLanguageEnglish_Click()

Info.Language = "EN"
Init_Language
Init_Menus

End Sub

Private Sub mnuOptionsLanguageGerman_Click()

Info.Language = "DE"
Init_Language
Init_Menus

End Sub

Private Sub mnuOptionsLanguageNorwegian_Click()

Info.Language = "NO"
Init_Language
Init_Menus

End Sub

Private Sub mnuOptionsLayoutBackBack_Click(Index As Integer)

Info.nCardBack = Index
DrawGameField

End Sub

Private Sub mnuOptionsLayoutBgBg_Click(Index As Integer)

Info.nBG = Index
GameField.TmpBG.fPath = Info.AppPath & "BG\BG" & Trim(Str(Index + 1)) & ".BMP"
LoadRes GameField.TmpBG
MakeBG
DrawGameField

End Sub

Private Sub mnuOptionsLayoutFrontFront_Click(Index As Integer)

Info.nCardFront = Index
DrawGameField

End Sub

Private Sub mnuOptionsRulesDrawOne_Click()
Info.nDrawCards = 1
DrawGameField
End Sub

Private Sub mnuOptionsRulesDrawThree_Click()
Info.nDrawCards = 3
DrawGameField
End Sub

Private Sub mnuOptionsSoundsDefsndSnd_Click(Index As Integer)

DoEvents
Info.nDefSnd = Index
LoadSnd Info.AppPath & "Snd" & Trim(Str(Info.nDefSnd + 1)) & ".wav", GameField.DefaultSnd

End Sub

Private Sub mnuOptionsSoundsGoalsndSnd_Click(Index As Integer)

DoEvents
Info.nGoalSnd = Index
LoadSnd Info.AppPath & "Snd" & Trim(Str(Info.nGoalSnd + 1)) & ".wav", GameField.GoalSnd

End Sub

Private Sub mnuOptionsSoundsPlaysounds_Click()

Info.PlaySounds = Not Info.PlaySounds

End Sub

Private Sub mnuTopFile_Click()

Info.Interrupt = True

If Info.Undo.Available Then
    mnuFileUndo.Enabled = True
Else
    mnuFileUndo.Enabled = False
End If

End Sub

Private Sub mnuTopHelp_Click()

Info.Interrupt = True

End Sub

Private Sub mnuTopOptions_Click()
Dim Cnt As Byte

Info.Interrupt = True

For Cnt = 1 To Info.nCardfronts
    mnuOptionsLayoutFrontFront(Cnt - 1).Checked = False
Next
mnuOptionsLayoutFrontFront(Info.nCardFront).Checked = True

For Cnt = 1 To Info.nCardBacks
    mnuOptionsLayoutBackBack(Cnt - 1).Checked = False
Next
mnuOptionsLayoutBackBack(Info.nCardBack).Checked = True

For Cnt = 1 To Info.nBGs
    mnuOptionsLayoutBgBg(Cnt - 1).Checked = False
Next
mnuOptionsLayoutBgBg(Info.nBG).Checked = True

For Cnt = 1 To Info.nSnds
    mnuOptionsSoundsDefsndSnd(Cnt - 1).Checked = False
    mnuOptionsSoundsGoalsndSnd(Cnt - 1).Checked = False
Next
mnuOptionsSoundsDefsndSnd(Info.nDefSnd).Checked = True
mnuOptionsSoundsGoalsndSnd(Info.nGoalSnd).Checked = True

If Info.Show790 Then
    mnuOptions790.Checked = True
Else
    mnuOptions790.Checked = False
End If

If Info.nDrawCards = 1 Then
    mnuOptionsRulesDrawOne.Checked = True
    mnuOptionsRulesDrawThree.Checked = False
Else
    mnuOptionsRulesDrawOne.Checked = False
    mnuOptionsRulesDrawThree.Checked = True
End If

If Info.PlaySounds Then
    mnuOptionsSoundsPlaysounds.Checked = True
Else
    mnuOptionsSoundsPlaysounds.Checked = False
End If

mnuOptionsLanguageEnglish.Checked = False
mnuOptionsLanguageGerman.Checked = False
mnuOptionsLanguageNorwegian.Checked = False
Select Case Info.Language
Case "EN"
    mnuOptionsLanguageEnglish.Checked = True
Case "DE"
    mnuOptionsLanguageGerman.Checked = True
Case "NO"
    mnuOptionsLanguageNorwegian.Checked = True
End Select

End Sub

Private Sub tmr790Blink_Timer()
Dim nFrames As Integer
Dim nFrame As Integer
Dim T1 As Long, T2 As Long
Dim FrameDelay As Integer

With GameField
    If Not Info.Interrupt Or Not Info.Show790 Then Exit Sub
    If Info.Moving And Info.ACPos.X > (.Pos790.X - 64) And Info.ACPos.X < (.Pos790.X + .Img790.nW) And Info.ACPos.Y < (.Pos790.Y + (.Img790.nH \ 2)) Then Exit Sub
    
    FrameDelay = 30
    nFrames = .Img790ani.nW \ 35
    T1 = GetTickCount
    For nFrame = 0 To nFrames - 1
        While (T2 - T1 < FrameDelay) Or Info.Moving And Info.ACPos.X > (.Pos790.X - 64) And Info.ACPos.X < (.Pos790.X + .Img790.nW) And Info.ACPos.Y < (.Pos790.Y + (.Img790.nH \ 2))
            T2 = GetTickCount
            DoEvents
        Wend
        'If Not Info.Moving Then
            If Info.Moving And Info.ACPos.X > (.Pos790.X - 64) And Info.ACPos.X < (.Pos790.X + .Img790.nW) And Info.ACPos.Y < (.Pos790.Y + (.Img790.nH \ 2)) Or Not Info.Show790 Then Exit Sub
            BitBlt .BB.hDC, .Pos790.X + 22, .Pos790.Y + 37, 35, 33, .Img790ani.hDC, nFrame * 35, 0, SRCCOPY
            BitBlt .BB.hDC, .Pos790.X + 68, .Pos790.Y + 36, 35, 33, .Img790ani.hDC, nFrame * 35, 33, SRCCOPY
            BitBlt .BB.hDC, .Pos790.X + 41, .Pos790.Y + 86, 35, 33, .Img790ani.hDC, nFrame * 35, 66, SRCCOPY
            BitBlt frmMain.hDC, .Pos790.X, .Pos790.Y, .Img790.nW, .Img790.nH \ 2, .BB.hDC, .Pos790.X, .Pos790.Y, SRCCOPY
        'End If
        T1 = GetTickCount
    Next nFrame
    T1 = GetTickCount
    For nFrame = nFrames - 1 To 0 Step -1
        While (T2 - T1 < FrameDelay) Or Info.Moving And Info.ACPos.X > (.Pos790.X - 64) And Info.ACPos.X < (.Pos790.X + .Img790.nW) And Info.ACPos.Y < (.Pos790.Y + (.Img790.nH \ 2))
            T2 = GetTickCount
            DoEvents
        Wend
        'If Not Info.Moving Then
            If Info.Moving And Info.ACPos.X > (.Pos790.X - 64) And Info.ACPos.X < (.Pos790.X + .Img790.nW) And Info.ACPos.Y < (.Pos790.Y + (.Img790.nH \ 2)) Or Not Info.Show790 Then Exit Sub
            BitBlt .BB.hDC, .Pos790.X + 22, .Pos790.Y + 37, 35, 33, .Img790ani.hDC, nFrame * 35, 0, SRCCOPY
            BitBlt .BB.hDC, .Pos790.X + 68, .Pos790.Y + 36, 35, 33, .Img790ani.hDC, nFrame * 35, 33, SRCCOPY
            BitBlt .BB.hDC, .Pos790.X + 41, .Pos790.Y + 86, 35, 33, .Img790ani.hDC, nFrame * 35, 66, SRCCOPY
            BitBlt frmMain.hDC, .Pos790.X, .Pos790.Y, .Img790.nW, .Img790.nH \ 2, .BB.hDC, .Pos790.X, .Pos790.Y, SRCCOPY
        'End If
    
        T1 = GetTickCount
    Next nFrame
End With

End Sub

Private Sub tmrUpdateTime_Timer()
Dim Elapsed As Long
Dim nHours As Long, nMinutes As Long, nSeconds As Long
Dim strH As String, strMin As String, strSec As String
Dim strTimeStatus As String
Dim MaxLen As Byte

Info.ThisTime = GetTickCount
Elapsed = (Info.ThisTime - Info.StartTime)

nHours = Elapsed \ 3600000
Elapsed = Elapsed - (nHours * 3600000)

nMinutes = Elapsed \ 60000
Elapsed = Elapsed - (nMinutes * 60000)

nSeconds = Elapsed \ 1000

strH = "": strMin = ""

If nHours > 0 Then strH = Trim(Str(nHours)) & " h "
If nMinutes > 0 Or nHours > 0 Then strMin = Trim(Str(nMinutes)) & " m "
strSec = Trim(Str(nSeconds)) & " s "

strH = Space$((3 + 3) - Len(strH)) & strH
strMin = Space$((2 + 3) - Len(strMin)) & strMin
strSec = Space$((2 + 3) - Len(strSec)) & strSec

strTimeStatus = strH & strMin & strSec
MaxLen = (3 + 3) + (2 + 3) + (2 + 3)
strTimeStatus = Space$(MaxLen - Len(strTimeStatus)) & strTimeStatus
frmMain.Status.Panels(2).Text = strTimeStatus

End Sub


