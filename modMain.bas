Attribute VB_Name = "modMain"
'Require explicit variable declaration
Option Explicit

'API STUFF:
'---------------------------------------------
'TYPES:
Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

'DECLARES:
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function MyGetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'CONSTANTS:
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCCOPY = &HCC0020
Public Const SND_MEMORY = &H4
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_LOOP = &H8
Public Const SND_PURGE = &H40


'---------------------------------------------

'Constants:
'-----------------------------------------
Global Const LR_LOADFROMFILE = 10
Global Const LR_CREATEDIBSECTION = 2000
Global Const APPNAME = "LexxSol"
'-----------------------------------------

'Types:
'-----------------------------------------
Type tCardElement
 Pos                    As POINTAPI     'Position inside the card
 Flag                   As Byte         '0 = not in use; 1 = normal; 2 = upside-down
End Type

Type tCard
 Elmt10(1 To 4)         As tCardElement 'The imgs on the card showing the sort. 10x10 pixels.
 Elmt15(1 To 10)        As tCardElement 'The imgs on the card showing the sort & value. 15x15 pixels.
 Value                  As Byte         'The card's value
 Sort                   As Byte         'The card type
 Flag                   As Byte         '0 = facedown; 1 = faceup
End Type

Type tPlaceHolder
 Pos                    As POINTAPI     'Placeholder position in gamefield
 Cards()                As Byte         'The cards in the placeholder. Not all may be in use.
 nCards                 As Byte         'The number of cards in the placeholder.
End Type

Type tResource
 fPath                  As String       'The path of the img file
 hBMP                   As Long         'The handle to the memory BMP
 hOldBMP                As Long         'The handle to the old BMP
 hdc                    As Long         'The Device Context handle
 nW                     As Integer      'The width of the BMP
 nH                     As Integer      'The height of the BMP
End Type

Type tUndo
 nCards As Byte
 nSrc As Byte
 nDest As Byte
 Available As Boolean
End Type

Type tGameField
 Card()                 As tCard        'The cards.
 CardNone               As tCard        'The no-card card
 CardBack               As tCard        'The upside-down card
 PlaceHolder(1 To 7)    As tPlaceHolder 'The placeholders in the middle
 GoalCell(1 To 4)       As tPlaceHolder 'The placeholders at the top
 Deck(1 To 2)           As tPlaceHolder 'The deck
 Img15                  As tResource    '15x15 img resource
 ImgCardBack            As tResource    'Card back images
 ImgCardFront           As tResource    'Card front images
 ImgGoal                As tResource    'Goal Cell Images
 ImgSign                As tResource    'Number/letter resource (10x10 pixels)
 ImgCard                As tResource    'Images on the img cards
 BB                     As tResource    'Back Buffer
 BG                     As tResource    'Background Image(in memory)
 TmpBG                  As tResource    'Temporary Background buffer(for tiling)
 Img790                 As tResource    'Picture of the 790 robot head
 Img790ani              As tResource    'Animation of 790 robot head
 BufTmp                 As tResource    'A temporary buffer
 nW                     As Long         'Width of gamefield
 nH                     As Long         'Height of gamefield
 Pos790                 As POINTAPI     'Position of the 790 robot head
 DefaultSnd             As String       'Buffer for default sound effect
 GoalSnd                As String       'Buffer for "move to goal" sound effect
 WinSnd                 As String       'Buffer for winning sound (drumroll)
 Win2Snd                As String       'Buffer for winning applause
End Type

Type tInfo
 ACPos                  As POINTAPI
 ClickPos               As POINTAPI
 DblClickPos            As POINTAPI
 Undo                   As tUndo
 AppPath                As String
 ActiveCard(1 To 13)    As Byte
 nActive                As Byte
 srcPH                  As Byte
 srcType                As Byte
 StartTime              As Long
 ThisTime               As Long
 nCardFront             As Integer
 nCardfronts            As Integer
 nCardBack              As Integer
 nCardBacks             As Integer
 nBG                    As Integer
 nBGs                   As Integer
 nDefSnd                As Integer
 nGoalSnd               As Integer
 nSnds                  As Integer
 MouseButton            As Byte
 nDrawCards             As Byte
 nRemoved               As Byte
 Moving                 As Boolean
 Interrupt              As Boolean
 CheckingWinner         As Boolean
 Debugging              As Boolean
 Show790                As Boolean
 PlaySounds             As Boolean
 UserName               As String
 Language               As String
End Type

'-----------------------------------------


'Globals:
'-----------------------------------------
Global GameField        As tGameField          'The Gamefield!!
Global Info             As tInfo                    'Some info for the app
Global Cards()          As tCard                   'The predefined cards

'Menu Item Captions:
    'Top Menus:
    '-------------------------------------------
    Global cWINDOWCAPTION As String
    Global cTOPFILE As String
    Global cTOPOPTIONS As String
    Global cTOPHELP As String
    '-------------------------------------------
    'File Sub Menu:
    '-------------------------------------------
    Global cNEWGAME As String
    Global cUNDO As String
    Global cEXIT As String
    '-------------------------------------------
    'Options Sub Menu:
    '-------------------------------------------
    Global cLANG As String
        Global cLANG_EN As String
        Global cLANG_DE As String
        Global cLANG_NO As String
    Global cLAYOUT As String
        Global cCARDFRONT As String
            Global cCARDFRONTFRONT As String
        Global cCARDBACK As String
            Global cCARDBACKBACK As String
        Global cBG As String
            Global cBGBG As String
    Global cSOUNDS As String
        Global cSOUNDSDEFSND As String
            Global cSOUNDSDEFSNDSND As String
        Global cSOUNDSGOALSND As String
            Global cSOUNDSGOALSNDSND As String
        Global cSOUNDSPLAY As String
    Global cRULES As String
        Global cRULESDRAW As String
            Global cRULESDRAWONE As String
            Global cRULESDRAWTHREE As String
    Global cSHOW790 As String
    'Help Sub Menu:
    '-------------------------------------------
    Global cABOUT As String
    Global cINDEX As String
    '-------------------------------------------
    'Misc Items:
    '-------------------------------------------
    Global cNEWGAMECAPTION As String
    Global cNEWGAMEMSG As String
    '-------------------------------------------

'-----------------------------------------

Public Sub Init_Menus()
Dim Cnt As Integer

With frmMain
    'Top Menus:
    '-----------------------------------------
    .mnuTopFile.Caption = cTOPFILE
    .mnuTopOptions.Caption = cTOPOPTIONS
    .mnuTopHelp.Caption = cTOPHELP
    '-----------------------------------------
    'File Sub Menu:
    '-----------------------------------------
    .mnuFileNewgame.Caption = cNEWGAME
    .mnuFileUndo.Caption = cUNDO
    .mnuFileExit.Caption = cEXIT
    '-----------------------------------------
    'Options Sub Menu:
    '-----------------------------------------
    .mnuOptionsLanguage.Caption = cLANG
        .mnuOptionsLanguageEnglish.Caption = cLANG_EN
        .mnuOptionsLanguageGerman.Caption = cLANG_DE
        .mnuOptionsLanguageNorwegian.Caption = cLANG_NO
    .mnuOptionsLayout.Caption = cLAYOUT
        .mnuOptionsLayoutFront.Caption = cCARDFRONT
            For Cnt = 0 To Info.nCardfronts - 1
                .mnuOptionsLayoutFrontFront(Cnt).Caption = cCARDFRONTFRONT & Trim(Str(Cnt + 1))
            Next Cnt
        .mnuOptionsLayoutBack.Caption = cCARDBACK
            For Cnt = 0 To Info.nCardBacks - 1
                .mnuOptionsLayoutBackBack(Cnt).Caption = cCARDBACKBACK & Trim(Str(Cnt + 1))
            Next Cnt
        .mnuOptionsLayoutBg.Caption = cBG
            For Cnt = 0 To Info.nBGs - 1
                .mnuOptionsLayoutBgBg(Cnt).Caption = cBGBG & Trim(Str(Cnt + 1))
            Next Cnt
    .mnuOptionsSounds.Caption = cSOUNDS
        .mnuOptionsSoundsDefsnd.Caption = cSOUNDSDEFSND
            For Cnt = 0 To Info.nSnds - 1
                .mnuOptionsSoundsDefsndSnd(Cnt).Caption = cSOUNDSDEFSNDSND & Trim(Str(Cnt + 1))
            Next Cnt
        .mnuOptionsSoundsGoalsnd.Caption = cSOUNDSGOALSND
            For Cnt = 0 To Info.nSnds - 1
                .mnuOptionsSoundsGoalsndSnd(Cnt).Caption = cSOUNDSGOALSNDSND & Trim(Str(Cnt + 1))
            Next Cnt
        .mnuOptionsSoundsPlaysounds.Caption = cSOUNDSPLAY
    .mnuOptionsRules.Caption = cRULES
        .mnuOptionsRulesDraw.Caption = cRULESDRAW
            .mnuOptionsRulesDrawOne.Caption = cRULESDRAWONE
            .mnuOptionsRulesDrawThree.Caption = cRULESDRAWTHREE
    .mnuOptions790.Caption = cSHOW790
    '-----------------------------------------
    'Help Sub Menu:
    '-----------------------------------------
    .mnuHelpAbout.Caption = cABOUT
    .mnuHelpHelp.Caption = cINDEX
    '-----------------------------------------
    
    'Window Caption:
    frmMain.Caption = cWINDOWCAPTION
    
End With

End Sub

Public Sub Init_Language()

If Info.Language = "NO" Then
    'Top Menus:
    '-------------------------------------------
    cWINDOWCAPTION = "Lexx Kabal - Laget av R.Ling"
    cTOPFILE = "&Fil"
    cTOPOPTIONS = "&Valg"
    cTOPHELP = "&Hjelp"
    '-------------------------------------------
    'File Sub Menu:
    '-------------------------------------------
    cNEWGAME = "&Nytt Spill"
    cUNDO = "&Angre"
    cEXIT = "Avslutt"
    '-------------------------------------------
    'Options Sub Menu:
    '-------------------------------------------
    cLANG = "Språk"
        cLANG_EN = "Engelsk (English)"
        cLANG_DE = "Tysk (German)"
        cLANG_NO = "Norsk (Norwegian)"
    cLAYOUT = "Utseende"
        cCARDFRONT = "Kort-Forside"
            cCARDFRONTFRONT = "Forside "
        cCARDBACK = "Kort-Bakside"
            cCARDBACKBACK = "Bakside "
        cBG = "Bakgrunn"
            cBGBG = "Bakgrunn "
    cSOUNDS = "&Lyder"
        cSOUNDSDEFSND = "Vanlig Flytt"
            cSOUNDSDEFSNDSND = "Lyd "
        cSOUNDSGOALSND = "Flytt Til Mål"
            cSOUNDSGOALSNDSND = "Lyd "
        cSOUNDSPLAY = "Spill Av Lyder"
    cRULES = "&Regler"
        cRULESDRAW = "Trekk.."
            cRULESDRAWONE = "Ett Kort"
            cRULESDRAWTHREE = "Tre Kort"
    cSHOW790 = "Vis &790"
    'Help Sub Menu:
    '-------------------------------------------
    cABOUT = "&Om Lexx Kabal.."
    cINDEX = "Innhold"
    '-------------------------------------------
    'Misc Items:
    '-------------------------------------------
    cNEWGAMECAPTION = "Nytt Spill?"
    cNEWGAMEMSG = "Dette vil avslutte gjeldende spill. Vil du fortsette?"
    '-------------------------------------------
    
ElseIf Info.Language = "EN" Then

    'Top Menus:
    '-------------------------------------------
    cWINDOWCAPTION = "Lexx Solitaire - By R.Ling"
    cTOPFILE = "&File"
    cTOPOPTIONS = "&Options"
    cTOPHELP = "&Help"
    '-------------------------------------------
    'File Sub Menu:
    '-------------------------------------------
    cNEWGAME = "&New Game"
    cUNDO = "&Undo"
    cEXIT = "Exit"
    '-------------------------------------------
    'Options Sub Menu:
    '-------------------------------------------
    cLANG = "Language"
        cLANG_EN = "English"
        cLANG_DE = "German"
        cLANG_NO = "Norwegian"
    cLAYOUT = "Layout"
        cCARDFRONT = "Card Front"
            cCARDFRONTFRONT = "Front "
        cCARDBACK = "Card Back"
            cCARDBACKBACK = "Back "
        cBG = "Background"
            cBGBG = "Background "
    cSOUNDS = "&Sounds"
        cSOUNDSDEFSND = "Ordinary Move"
            cSOUNDSDEFSNDSND = "Sound "
        cSOUNDSGOALSND = "Move To Goal"
            cSOUNDSGOALSNDSND = "Sound "
        cSOUNDSPLAY = "Play Sounds"
    cRULES = "&Rules"
        cRULESDRAW = "Draw.."
            cRULESDRAWONE = "One Card"
            cRULESDRAWTHREE = "Three Cards"
    cSHOW790 = "Show &790"
    '-------------------------------------------
    'Help Sub Menu:
    '-------------------------------------------
    cABOUT = "&About Lexx Solitaire.."
    cINDEX = "Index"
    '-------------------------------------------
    'Misc Items:
    '-------------------------------------------
    cNEWGAMECAPTION = "Start New Game?"
    cNEWGAMEMSG = "This will end the current game. Do you wish to continue?"
    '-------------------------------------------

ElseIf Info.Language = "DE" Then

    'Top Menus:
    '-------------------------------------------
    cWINDOWCAPTION = "Lexx Patience - von R.Ling"
    cTOPFILE = "&Datei"
    cTOPOPTIONS = "&Wahlen"
    cTOPHELP = "&Hilfe"
    '-------------------------------------------
    'File Sub Menu:
    '-------------------------------------------
    cNEWGAME = "&Neues Spiel"
    cUNDO = "&Bereuen"
    cEXIT = "Abschließ"
    '-------------------------------------------
    'Options Sub Menu:
    '-------------------------------------------
    cLANG = "Sprache"
        cLANG_EN = "Englisch (English)"
        cLANG_DE = "Deutsch (German)"
        cLANG_NO = "Norwegisch (Norwegian)"
    cLAYOUT = "Äußere"
        cCARDFRONT = "Karten-Vorderseite"
            cCARDFRONTFRONT = "Vorderseite "
        cCARDBACK = "Karten-Rückseite"
            cCARDBACKBACK = "Rückseite "
        cBG = "Hintergrund"
            cBGBG = "Hintergrund "
    cSOUNDS = "&Laute"
        cSOUNDSDEFSND = "Normal - Ziehen"
            cSOUNDSDEFSNDSND = "Laut "
        cSOUNDSGOALSND = "Ziel - Ziehen"
            cSOUNDSGOALSNDSND = "Laut "
        cSOUNDSPLAY = "Spiel Laute Aus"
    cRULES = "Regeln"
        cRULESDRAW = "Ziehen.."
            cRULESDRAWONE = "Ein Karte"
            cRULESDRAWTHREE = "Drei Karten"
    cSHOW790 = "Vorführ &790"
    '-------------------------------------------
    'Help Sub Menu:
    '-------------------------------------------
    cABOUT = "&Über Lexx Patience.."
    cINDEX = "Inhalt"
    '-------------------------------------------
    'Misc Items:
    '-------------------------------------------
    cNEWGAMECAPTION = "Starten Neues Spiel?"
    cNEWGAMEMSG = "Dies will das geltende Spiel abschließen. Willst du fortsetzen?"
    '-------------------------------------------
End If


End Sub

Public Function CheckMouseUp_Placeholders(x As Single, y As Single) As Boolean
Dim Xpos        As Integer, Ypos        As Integer
Dim SrcX        As Integer, SrcY        As Integer
Dim DestX       As Integer, DestY       As Integer
Dim FrameCnt    As Integer, nFrames     As Integer
Dim nPlaceH     As Byte, Cnt            As Byte

CheckMouseUp_Placeholders = False
If Info.Moving Then

    'Find out if the card has landed in a placeholder or a goalcell, and if it is a valid move:
    '---------------------------------------------------------
    For Cnt = 1 To 7
        If x - Info.ClickPos.x >= GameField.PlaceHolder(Cnt).Pos.x - 37 And _
            x - Info.ClickPos.x < GameField.PlaceHolder(Cnt).Pos.x + 37 And _
            y - Info.ClickPos.y >= GameField.PlaceHolder(Cnt).Pos.y - 25 And _
            y - Info.ClickPos.y <= GameField.PlaceHolder(Cnt).Pos.y + ((GameField.PlaceHolder(Cnt).nCards + 5) * 14) + 25 _
        Then
            nPlaceH = Cnt
            Info.Moving = True
            Exit For
        End If
    Next Cnt

    If nPlaceH = 0 Then Exit Function
    If GameField.PlaceHolder(nPlaceH).nCards = 0 And GameField.Card(Info.ActiveCard(1)).Value <> 13 Then Exit Function
    If nPlaceH > 0 And GameField.Card(Info.ActiveCard(1)).Flag = 1 And _
    GameField.Card(Info.ActiveCard(1)).Sort \ 2 <> GameField.Card(GameField.PlaceHolder(nPlaceH).Cards(GameField.PlaceHolder(nPlaceH).nCards)).Sort \ 2 And _
    GameField.Card(Info.ActiveCard(1)).Value = GameField.Card(GameField.PlaceHolder(nPlaceH).Cards(GameField.PlaceHolder(nPlaceH).nCards)).Value - 1 Or _
    (GameField.Card(Info.ActiveCard(1)).Value = 13 And GameField.PlaceHolder(nPlaceH).nCards = 0) _
    Then
        'Valid Move: Transfer the cards to the placeholder/goalcell:
        '---------------------------------------------------------
        For Cnt = 1 To Info.nActive
            GameField.PlaceHolder(nPlaceH).nCards = GameField.PlaceHolder(nPlaceH).nCards + 1
            GameField.PlaceHolder(nPlaceH).Cards(GameField.PlaceHolder(nPlaceH).nCards) = Info.ActiveCard(Cnt)
        Next
        Info.Undo.nCards = Info.nActive
        Info.Undo.nSrc = Info.srcPH
        Info.Undo.nDest = nPlaceH
        Info.Undo.Available = True
        
        'Remove the active cards:
        Info.nActive = 0
        Info.Moving = False

        'Set function to true:
        CheckMouseUp_Placeholders = True

        'Draw the placeholder:
        DrawPlaceHolder nPlaceH, GameField.BB.hdc
        FlipScreen

        If Not (Info.srcType = 1 And Info.srcPH = nPlaceH) Then
            'Play sound:
        PlaySnd GameField.DefaultSnd
        End If
    End If
End If

End Function
Public Sub MoveBack()
Dim SrcX As Long, SrcY As Long, DestX As Long, DestY As Long, Xpos As Long, Ypos As Long
Dim nFrames As Byte, FrameCnt As Byte
Dim SrcPlaceHolder As Byte, Cnt As Long

Select Case Info.srcType
Case 1
    With GameField.PlaceHolder(Info.srcPH)
        SrcX = Info.ACPos.x
        DestX = .Pos.x
        SrcY = Info.ACPos.y
        DestY = .Pos.y + (.nCards * 14)
    End With
Case 2
    With GameField.GoalCell(Info.srcPH - 7)
        SrcX = Info.ACPos.x
        DestX = .Pos.x
        SrcY = Info.ACPos.y
        DestY = .Pos.y
    End With
Case 3
    With GameField.Deck(2)
        SrcX = Info.ACPos.x
        If Info.nDrawCards = 3 Then
        'Three Card option.
        '-----------------------
            'If Info.nRemoved > 0 And Info.nRemoved + GameField.Deck(2).nCards >= 3 Then
            '    If GameField.Deck(2).nCards >= 3 Then
            '        DestX = .Pos.X + ((3 - Info.nRemoved) * 15)
            '    Else
            '        DestX = .Pos.X + (((GameField.Deck(2).nCards + 1) - (4 - Info.nRemoved)) * 15)
            '    End If
            'Else
            '    DestX = .Pos.X
            'End If
            Dim XtraPix As Integer
            If GameField.Deck(2).nCards >= 3 Then
                XtraPix = (3 - Info.nRemoved) * 15
            Else
                'XtraPix = (GameField.Deck(2).nCards - 1) * 15
                If Info.nRemoved + GameField.Deck(2).nCards >= 3 Then
                    XtraPix = (3 - Info.nRemoved) * 15
                Else
                XtraPix = (GameField.Deck(2).nCards) * 15
                End If
            End If
            DestX = .Pos.x + XtraPix
        '-----------------------
        Else
            DestX = .Pos.x
        End If
        SrcY = Info.ACPos.y
        DestY = .Pos.y
    End With
End Select

nFrames = 25
Info.ClickPos.x = 0
Info.ClickPos.y = 0
    
For FrameCnt = 1 To nFrames
    Xpos = SrcX + (((DestX - SrcX) * FrameCnt) \ nFrames)
    Ypos = SrcY + (((DestY - SrcY) * FrameCnt) \ nFrames)
    DrawActiveCard CSng(Xpos), CSng(Ypos)
Next

Select Case Info.srcType
Case 1
    For Cnt = 1 To Info.nActive
        GameField.PlaceHolder(Info.srcPH).nCards = GameField.PlaceHolder(Info.srcPH).nCards + 1
        GameField.PlaceHolder(Info.srcPH).Cards(GameField.PlaceHolder(Info.srcPH).nCards) = Info.ActiveCard(Cnt)
    Next Cnt
    DrawPlaceHolder Info.srcPH, GameField.BB.hdc
Case 2
    For Cnt = 1 To Info.nActive
        GameField.GoalCell(Info.srcPH - 7).nCards = GameField.GoalCell(Info.srcPH - 7).nCards + 1
        GameField.GoalCell(Info.srcPH - 7).Cards(GameField.GoalCell(Info.srcPH - 7).nCards) = Info.ActiveCard(Cnt)
    Next Cnt
    DrawGoalCell Info.srcPH - 7, GameField.BB.hdc
Case 3
    For Cnt = 1 To Info.nActive
        GameField.Deck(2).nCards = GameField.Deck(2).nCards + 1
        GameField.Deck(2).Cards(GameField.Deck(2).nCards) = Info.ActiveCard(Cnt)
        If Info.nRemoved > 0 Then
            Info.nRemoved = Info.nRemoved - 1
        Else
            Info.nRemoved = 2
        End If
    Next Cnt
    DrawDeck GameField.BB.hdc
End Select

Info.nActive = 0
Info.Moving = False
FlipScreen
    
End Sub
Public Sub DrawActiveCard(x As Single, y As Single)
Dim cPos As POINTAPI, TmpPos As POINTAPI, TmpPos2 As POINTAPI, sX As Long, sY As Long, MinX As Long, MinY As Long
Dim MaxX As Long, MaxY As Long
Dim Cnt As Integer

TmpPos.x = x - Info.ClickPos.x
TmpPos.y = y - Info.ClickPos.y

If Not Info.nActive > 0 Then Exit Sub

If Abs(TmpPos.x - Info.ACPos.x) < 64 And Abs(TmpPos.y - Info.ACPos.y) < 104 Then
    If TmpPos.x > Info.ACPos.x Then
        sX = 0
        cPos.x = TmpPos.x - Info.ACPos.x
        MinX = Info.ACPos.x
        MaxX = TmpPos.x + 64
    Else
        cPos.x = 0
        sX = Info.ACPos.x - TmpPos.x
        MinX = TmpPos.x
        MaxX = Info.ACPos.x + 64
    End If
    If TmpPos.y > Info.ACPos.y Then
        sY = 0
        cPos.y = TmpPos.y - Info.ACPos.y
        MinY = Info.ACPos.y
        MaxY = TmpPos.y + 104 + (14 * (Info.nActive - 1))
    Else
        cPos.y = 0
        sY = Info.ACPos.y - TmpPos.y
        MinY = TmpPos.y
        MaxY = Info.ACPos.y + 104 + (14 * (Info.nActive - 1))
    End If


    BitBlt GameField.BufTmp.hdc, 0, 0, 128, 376, GameField.BB.hdc, MinX, MinY, SRCCOPY
    Info.ACPos = TmpPos
    
    For Cnt = 1 To Info.nActive
        DrawCard GameField.BufTmp.hdc, cPos, GameField.Card(Info.ActiveCard(Cnt))
        cPos.y = cPos.y + 14
    Next Cnt
    BitBlt frmMain.hdc, MinX, MinY, MaxX - MinX, MaxY - MinY, GameField.BufTmp.hdc, 0, 0, SRCCOPY
    
Else
    BitBlt frmMain.hdc, Info.ACPos.x, Info.ACPos.y, 64, 104 + (14 * Info.nActive), GameField.BB.hdc, Info.ACPos.x, Info.ACPos.y, SRCCOPY
    Info.ACPos = TmpPos
    TmpPos.x = 0
    TmpPos.y = 0
    For Cnt = 0 To Info.nActive - 1
        With GameField.Card(Info.ActiveCard(Cnt + 1))
            If .Flag = 1 Then
                TmpPos2.x = 0
                TmpPos2.y = Cnt * 14
                DrawCard GameField.BufTmp.hdc, TmpPos2, GameField.Card(Info.ActiveCard(Cnt + 1))
            Else
                DrawCard GameField.BufTmp.hdc, TmpPos2, GameField.CardNone
            End If
        End With
    Next Cnt
    TmpPos.x = 0
    TmpPos.y = 14 * (Info.nActive - 1)
    DrawCard GameField.BufTmp.hdc, TmpPos, GameField.Card(Info.ActiveCard(Info.nActive))
    BitBlt frmMain.hdc, Info.ACPos.x, Info.ACPos.y, 64, 104 + (14 * (Info.nActive - 1)), GameField.BufTmp.hdc, 0, 0, SRCCOPY
End If

End Sub


Public Sub DrawCard(hdc As Long, Pos As POINTAPI, cCard As tCard)
Dim nCnt As Byte, SrcX As Long, SrcY As Long

If cCard.Flag = 0 Then

    'Blit the back:
    BitBlt hdc, Pos.x, Pos.y, 64, 104, GameField.ImgCardBack.hdc, 0, Info.nCardBack * 104, SRCCOPY

ElseIf cCard.Flag = 1 Then
    
    'Blit the card front:
    '-------------------------------------
    BitBlt hdc, Pos.x, Pos.y, 64, 104, GameField.ImgCardFront.hdc, 0, Info.nCardFront * 104, SRCCOPY
    '-------------------------------------

    'Card imgs on image cards:
    '-------------------------------------
    If cCard.Value > 10 Then
        BitBlt hdc, Pos.x + 12, Pos.y + 17, 40, 70, GameField.ImgCard.hdc, (cCard.Value - 11) * 40, 0, SRCCOPY
    End If

    '-------------------------------------

    'The 15x15 imgs:
    '-------------------------------------
    For nCnt = 1 To 10
        With cCard.Elmt15(nCnt)
            SrcX = cCard.Sort * 15
            If .Flag = 1 Then
                BitBlt hdc, Pos.x + .Pos.x, Pos.y + .Pos.y, 15, 15, GameField.Img15.hdc, SrcX, 30, SRCAND
                BitBlt hdc, Pos.x + .Pos.x, Pos.y + .Pos.y, 15, 15, GameField.Img15.hdc, SrcX, 0, SRCINVERT
            ElseIf .Flag = 2 Then
                BitBlt hdc, Pos.x + .Pos.x, Pos.y + .Pos.y, 15, 15, GameField.Img15.hdc, SrcX, 45, SRCAND
                BitBlt hdc, Pos.x + .Pos.x, Pos.y + .Pos.y, 15, 15, GameField.Img15.hdc, SrcX, 15, SRCINVERT
            Else
                Exit For
            End If
        End With
    Next
    '-------------------------------------
    
    'The Numbers/Letters:
    SrcX = (cCard.Value - 1) * 10
    SrcY = ((cCard.Sort \ 2) * 40)
    BitBlt hdc, Pos.x + 4, Pos.y + 4, 10, 10, GameField.ImgSign.hdc, SrcX, SrcY + 20, SRCAND
    BitBlt hdc, Pos.x + 4, Pos.y + 4, 10, 10, GameField.ImgSign.hdc, SrcX, SrcY, SRCINVERT
    
    BitBlt hdc, Pos.x + 50, Pos.y + 88, 10, 10, GameField.ImgSign.hdc, SrcX, SrcY + 30, SRCAND
    BitBlt hdc, Pos.x + 50, Pos.y + 88, 10, 10, GameField.ImgSign.hdc, SrcX, SrcY + 10, SRCINVERT

ElseIf cCard.Flag = 2 Then
    
    'Blit the "no cards"(faded) image:
    BitBlt hdc, Pos.x, Pos.y, 64, 104, GameField.ImgCardBack.hdc, 64, Info.nCardBack * 104, SRCCOPY

End If

End Sub


Public Sub Init_Cards()
Dim SortCnt As Byte, CardCnt As Byte, nCard As Byte
ReDim Cards(1 To 13)
ReDim GameField.Card(0 To 52) 'Base 0 to avoid subscript out of range errors

'CARD ELEMENT POSITIONS:
'-----------------------------------------------------------------------------------------------

'Card 1:
Cards(1).Elmt15(1).Pos.x = 24: Cards(1).Elmt15(1).Pos.y = 44: Cards(1).Elmt15(1).Flag = 1

'Card 2:
Cards(2).Elmt15(1).Pos.x = 24: Cards(2).Elmt15(1).Pos.y = 24: Cards(2).Elmt15(1).Flag = 1
Cards(2).Elmt15(2).Pos.x = 24: Cards(2).Elmt15(2).Pos.y = 64: Cards(2).Elmt15(2).Flag = 2

'Card 3:
Cards(3).Elmt15(1).Pos.x = 24: Cards(3).Elmt15(1).Pos.y = 18: Cards(3).Elmt15(1).Flag = 1
Cards(3).Elmt15(2).Pos.x = 24: Cards(3).Elmt15(2).Pos.y = 44: Cards(3).Elmt15(2).Flag = 1
Cards(3).Elmt15(3).Pos.x = 24: Cards(3).Elmt15(3).Pos.y = 70: Cards(3).Elmt15(3).Flag = 2

'Card 4:
Cards(4).Elmt15(1).Pos.x = 14: Cards(4).Elmt15(1).Pos.y = 20: Cards(4).Elmt15(1).Flag = 1
Cards(4).Elmt15(2).Pos.x = 35: Cards(4).Elmt15(2).Pos.y = 20: Cards(4).Elmt15(2).Flag = 1
Cards(4).Elmt15(3).Pos.x = 14: Cards(4).Elmt15(3).Pos.y = 69: Cards(4).Elmt15(3).Flag = 2
Cards(4).Elmt15(4).Pos.x = 35: Cards(4).Elmt15(4).Pos.y = 69: Cards(4).Elmt15(4).Flag = 2

'Card 5:
Cards(5).Elmt15(1).Pos.x = 14: Cards(5).Elmt15(1).Pos.y = 20: Cards(5).Elmt15(1).Flag = 1
Cards(5).Elmt15(2).Pos.x = 35: Cards(5).Elmt15(2).Pos.y = 20: Cards(5).Elmt15(2).Flag = 1
Cards(5).Elmt15(3).Pos.x = 14: Cards(5).Elmt15(3).Pos.y = 69: Cards(5).Elmt15(3).Flag = 2
Cards(5).Elmt15(4).Pos.x = 35: Cards(5).Elmt15(4).Pos.y = 69: Cards(5).Elmt15(4).Flag = 2
Cards(5).Elmt15(5).Pos.x = 24: Cards(5).Elmt15(5).Pos.y = 44: Cards(5).Elmt15(5).Flag = 1

'Card 6:
Cards(6).Elmt15(1).Pos.x = 14: Cards(6).Elmt15(1).Pos.y = 18: Cards(6).Elmt15(1).Flag = 1
Cards(6).Elmt15(2).Pos.x = 35: Cards(6).Elmt15(2).Pos.y = 18: Cards(6).Elmt15(2).Flag = 1
Cards(6).Elmt15(3).Pos.x = 14: Cards(6).Elmt15(3).Pos.y = 44: Cards(6).Elmt15(3).Flag = 1
Cards(6).Elmt15(4).Pos.x = 35: Cards(6).Elmt15(4).Pos.y = 44: Cards(6).Elmt15(4).Flag = 1
Cards(6).Elmt15(5).Pos.x = 14: Cards(6).Elmt15(5).Pos.y = 70: Cards(6).Elmt15(5).Flag = 2
Cards(6).Elmt15(6).Pos.x = 35: Cards(6).Elmt15(6).Pos.y = 70: Cards(6).Elmt15(6).Flag = 2

'Card 7:
Cards(7).Elmt15(1).Pos.x = 14: Cards(7).Elmt15(1).Pos.y = 18: Cards(7).Elmt15(1).Flag = 1
Cards(7).Elmt15(2).Pos.x = 35: Cards(7).Elmt15(2).Pos.y = 18: Cards(7).Elmt15(2).Flag = 1
Cards(7).Elmt15(3).Pos.x = 14: Cards(7).Elmt15(3).Pos.y = 48: Cards(7).Elmt15(3).Flag = 1
Cards(7).Elmt15(4).Pos.x = 35: Cards(7).Elmt15(4).Pos.y = 48: Cards(7).Elmt15(4).Flag = 1
Cards(7).Elmt15(5).Pos.x = 14: Cards(7).Elmt15(5).Pos.y = 70: Cards(7).Elmt15(5).Flag = 2
Cards(7).Elmt15(6).Pos.x = 35: Cards(7).Elmt15(6).Pos.y = 70: Cards(7).Elmt15(6).Flag = 2
Cards(7).Elmt15(7).Pos.x = 24: Cards(7).Elmt15(7).Pos.y = 34: Cards(7).Elmt15(7).Flag = 1

'Card 8:
Cards(8).Elmt15(1).Pos.x = 14: Cards(8).Elmt15(1).Pos.y = 14: Cards(8).Elmt15(1).Flag = 1
Cards(8).Elmt15(2).Pos.x = 35: Cards(8).Elmt15(2).Pos.y = 14: Cards(8).Elmt15(2).Flag = 1
Cards(8).Elmt15(3).Pos.x = 14: Cards(8).Elmt15(3).Pos.y = 34: Cards(8).Elmt15(3).Flag = 1
Cards(8).Elmt15(4).Pos.x = 35: Cards(8).Elmt15(4).Pos.y = 34: Cards(8).Elmt15(4).Flag = 1
Cards(8).Elmt15(5).Pos.x = 14: Cards(8).Elmt15(5).Pos.y = 54: Cards(8).Elmt15(5).Flag = 2
Cards(8).Elmt15(6).Pos.x = 35: Cards(8).Elmt15(6).Pos.y = 54: Cards(8).Elmt15(6).Flag = 2
Cards(8).Elmt15(7).Pos.x = 14: Cards(8).Elmt15(7).Pos.y = 74: Cards(8).Elmt15(7).Flag = 2
Cards(8).Elmt15(8).Pos.x = 35: Cards(8).Elmt15(8).Pos.y = 74: Cards(8).Elmt15(8).Flag = 2

'Card 9:
Cards(9).Elmt15(1).Pos.x = 10: Cards(9).Elmt15(1).Pos.y = 14: Cards(9).Elmt15(1).Flag = 1
Cards(9).Elmt15(2).Pos.x = 39: Cards(9).Elmt15(2).Pos.y = 14: Cards(9).Elmt15(2).Flag = 1
Cards(9).Elmt15(3).Pos.x = 10: Cards(9).Elmt15(3).Pos.y = 34: Cards(9).Elmt15(3).Flag = 1
Cards(9).Elmt15(4).Pos.x = 39: Cards(9).Elmt15(4).Pos.y = 34: Cards(9).Elmt15(4).Flag = 1
Cards(9).Elmt15(5).Pos.x = 10: Cards(9).Elmt15(5).Pos.y = 54: Cards(9).Elmt15(5).Flag = 2
Cards(9).Elmt15(6).Pos.x = 39: Cards(9).Elmt15(6).Pos.y = 54: Cards(9).Elmt15(6).Flag = 2
Cards(9).Elmt15(7).Pos.x = 10: Cards(9).Elmt15(7).Pos.y = 74: Cards(9).Elmt15(7).Flag = 2
Cards(9).Elmt15(8).Pos.x = 39: Cards(9).Elmt15(8).Pos.y = 74: Cards(9).Elmt15(8).Flag = 2
Cards(9).Elmt15(9).Pos.x = 24: Cards(9).Elmt15(9).Pos.y = 44: Cards(9).Elmt15(9).Flag = 1

'Card 10:
Cards(10).Elmt15(1).Pos.x = 10: Cards(10).Elmt15(1).Pos.y = 14: Cards(10).Elmt15(1).Flag = 1
Cards(10).Elmt15(2).Pos.x = 39: Cards(10).Elmt15(2).Pos.y = 14: Cards(10).Elmt15(2).Flag = 1
Cards(10).Elmt15(3).Pos.x = 10: Cards(10).Elmt15(3).Pos.y = 34: Cards(10).Elmt15(3).Flag = 1
Cards(10).Elmt15(4).Pos.x = 39: Cards(10).Elmt15(4).Pos.y = 34: Cards(10).Elmt15(4).Flag = 1
Cards(10).Elmt15(5).Pos.x = 10: Cards(10).Elmt15(5).Pos.y = 54: Cards(10).Elmt15(5).Flag = 2
Cards(10).Elmt15(6).Pos.x = 39: Cards(10).Elmt15(6).Pos.y = 54: Cards(10).Elmt15(6).Flag = 2
Cards(10).Elmt15(7).Pos.x = 10: Cards(10).Elmt15(7).Pos.y = 74: Cards(10).Elmt15(7).Flag = 2
Cards(10).Elmt15(8).Pos.x = 39: Cards(10).Elmt15(8).Pos.y = 74: Cards(10).Elmt15(8).Flag = 2
Cards(10).Elmt15(9).Pos.x = 24: Cards(10).Elmt15(9).Pos.y = 24: Cards(10).Elmt15(9).Flag = 1
Cards(10).Elmt15(10).Pos.x = 24: Cards(10).Elmt15(10).Pos.y = 64: Cards(10).Elmt15(10).Flag = 2

'Card 11:
 Cards(11).Elmt15(1).Pos.x = 14: Cards(11).Elmt15(1).Pos.y = 19: Cards(11).Elmt15(1).Flag = 1
 Cards(11).Elmt15(2).Pos.x = 35: Cards(11).Elmt15(2).Pos.y = 70: Cards(11).Elmt15(2).Flag = 2

'Card 12:
 Cards(12).Elmt15(1).Pos.x = 14: Cards(12).Elmt15(1).Pos.y = 19: Cards(12).Elmt15(1).Flag = 1
 Cards(12).Elmt15(2).Pos.x = 35: Cards(12).Elmt15(2).Pos.y = 70: Cards(12).Elmt15(2).Flag = 2

'Card 13:
 Cards(13).Elmt15(1).Pos.x = 14: Cards(13).Elmt15(1).Pos.y = 19: Cards(13).Elmt15(1).Flag = 1
 Cards(13).Elmt15(2).Pos.x = 35: Cards(13).Elmt15(2).Pos.y = 70: Cards(13).Elmt15(2).Flag = 2

'-----------------------------------------------------------------------------------------------

For SortCnt = 0 To 3: For CardCnt = 1 To 13
    nCard = (SortCnt * 13) + CardCnt
    GameField.Card(nCard) = Cards(CardCnt)
    GameField.Card(nCard).Sort = SortCnt
    GameField.Card(nCard).Value = CardCnt
Next CardCnt: Next SortCnt

'Misc cards:
GameField.CardBack.Flag = 0
GameField.CardNone.Flag = 2

End Sub

Public Sub LoadRes(MyRes As tResource)
Dim MyBMP As BITMAP

If Dir(MyRes.fPath) = "" Then
    Exit Sub
End If

'Free res:
FreeRes MyRes

With MyRes
    .hdc = CreateCompatibleDC(0)
    frmMain.imgLoad.Picture = LoadPicture(.fPath)
    .hBMP = CreateCompatibleBitmap(frmMain.hdc, frmMain.imgLoad.ScaleWidth, frmMain.imgLoad.ScaleHeight)
    '.hBMP = LoadImage(0, .fPath, 0, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
    .hOldBMP = SelectObject(.hdc, .hBMP)
    BitBlt .hdc, 0, 0, frmMain.imgLoad.ScaleWidth, frmMain.imgLoad.ScaleHeight, frmMain.imgLoad.hdc, 0, 0, SRCCOPY
    MyGetObject .hBMP, Len(MyBMP), MyBMP
    .nW = MyBMP.bmWidth
    .nH = MyBMP.bmHeight
    Set frmMain.imgLoad.Picture = Nothing
    'If Info.Debugging Then frmMain.Caption = frmMain.Caption & " " & Trim(Str(MyBMP.bmBitsPixel))
    If .hBMP = 0 Then MsgBox "Error: Unable to load " & .fPath & "!!!"
End With

End Sub

Public Sub FreeRes(MyRes As tResource)
With MyRes
    SelectObject .hdc, .hOldBMP
    DeleteObject .hBMP
    DeleteObject .hOldBMP
    DeleteDC .hdc
End With
End Sub

Public Sub CleanUp()

'Free All Resources:
'--------------------------------
With GameField
    FreeRes .Img15
    FreeRes .ImgCard
    FreeRes .ImgCardFront
    FreeRes .ImgCardBack
    FreeRes .ImgSign
    FreeRes .Img790
    FreeRes .Img790ani
    FreeRes .TmpBG
    FreeRes .BufTmp
    FreeRes .ImgGoal
    FreeRes .BB
    FreeRes .BG
End With

Erase GameField.Card
Erase GameField.PlaceHolder
Erase GameField.GoalCell

'Remove menu items(if debugging is enabled and the user just wishes to reload the GFX)
Dim nCnt As Long
For nCnt = 1 To Info.nCardfronts - 1
    Unload frmMain.mnuOptionsLayoutFrontFront(nCnt)
Next

For nCnt = 1 To Info.nCardBacks - 1
    Unload frmMain.mnuOptionsLayoutBackBack(nCnt)
Next

For nCnt = 1 To Info.nBGs - 1
    Unload frmMain.mnuOptionsLayoutBgBg(nCnt)
Next

frmMain.Caption = cWINDOWCAPTION

'Save the current settings:
SaveSettings
'--------------------------------

End Sub

Public Sub Main()
Dim WrongRes As Boolean, WrongColDepth As Boolean, ErrMsg As String
Dim Ret As Integer: Ret = 0

'Check screen resolution:
'--------------------------------------
WrongRes = False: WrongColDepth = False
If Screen.Width \ Screen.TwipsPerPixelX < 800 Or _
Screen.Height \ Screen.TwipsPerPixelY < 600 Then
    WrongRes = True
    ErrMsg = ErrMsg & "This game was designed for a" & vbLf & "screen resolution of 800x600 pixels" & vbLf & "or higher. You should change the screen" & vbLf & "resolution before playing this game."
End If
If ErrMsg <> "" Then ErrMsg = ErrMsg & vbLf & vbLf & "Do you want to ignore these errors and proceed?"
If ErrMsg <> "" Then Ret = MsgBox(ErrMsg, vbYesNo Or vbExclamation, "Error(s) Reported At Startup:")
If Ret = 7 Then End
'--------------------------------------

'Move Main window into position:
'------------------------
Load frmMain
frmMain.Move 0, 0, (640 + 3 + 3) * Screen.TwipsPerPixelX, (480 + 40 + 4) * Screen.TwipsPerPixelY
frmMain.Left = (Screen.Width \ 2) - (frmMain.Width \ 2)
frmMain.Top = 0 '(Screen.Height \ 2) - (frmMain.Height \ 2)
GameField.nW = frmMain.ScaleWidth
GameField.nH = frmMain.ScaleHeight
DoEvents
'------------------------

'Fix app path:
'---------------
If Len(App.Path) = 3 Then
    Info.AppPath = App.Path
Else
    Info.AppPath = App.Path & "\"
End If
'---------------

'Show Splash Window:
'-------------------------------------
'Load frmSplash
'frmSplash.Width = 472 * Screen.TwipsPerPixelX
'frmSplash.Height = 331 * Screen.TwipsPerPixelY
'frmSplash.Left = (Screen.Width \ 2) - (frmSplash.Width \ 2)
'frmSplash.Top = (Screen.Height \ 2) - (frmSplash.Height \ 2)
'SetWindowPos frmSplash.hWnd, -1, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE

'frmSplash.Picture = LoadPicture(Info.AppPath & "Splash.bmp")
'frmSplash.Show
'DoEvents
'-------------------------------------

'Get the user's logon name:
Info.UserName = Space$(255): GetUserName Info.UserName, 255
Info.UserName = Trim(Info.UserName)
If Info.UserName <> "" Then
    Info.UserName = Left(Info.UserName, Len(Info.UserName) - 1)
Else
    Info.UserName = "$Default$"
End If

'Initialize:
'---------------------------------
ReadSettings
Init_Language
Init_Menus
Init_Res
Init_Cards
Init_Placeholders
Init_Sounds

'Make Undo unavailable:
Info.Undo.Available = False

'Place the 790 robot head:
GameField.Pos790.x = 195 '189
GameField.Pos790.y = 15
'---------------------------------

frmMain.tmr790Blink.Enabled = True
frmSplash.tmrEffect.Enabled = True

End Sub

Public Sub DrawGoalCells(hdc As Long)
Dim Cnt As Byte

For Cnt = 1 To 4
    With GameField.GoalCell(Cnt)
        If .nCards > 0 Then
            DrawCard hdc, .Pos, GameField.Card(.Cards(.nCards))
        Else
            BitBlt hdc, .Pos.x, .Pos.y, 64, 104, GameField.BG.hdc, .Pos.x, .Pos.y, SRCCOPY
            BitBlt hdc, .Pos.x, .Pos.y, 64, 104, GameField.ImgGoal.hdc, (Cnt - 1) * 64, 104, SRCAND
            BitBlt hdc, .Pos.x, .Pos.y, 64, 104, GameField.ImgGoal.hdc, (Cnt - 1) * 64, 0, SRCINVERT
        End If
    End With
Next

End Sub

Public Sub DrawGoalCell(nGoalCell As Byte, hdc As Long)

With GameField.GoalCell(nGoalCell)
    If .nCards > 0 Then
        DrawCard hdc, .Pos, GameField.Card(.Cards(.nCards))
    Else
        BitBlt hdc, .Pos.x, .Pos.y, 64, 104, GameField.BG.hdc, .Pos.x, .Pos.y, SRCCOPY
        BitBlt hdc, .Pos.x, .Pos.y, 64, 104, GameField.ImgGoal.hdc, (nGoalCell - 1) * 64, 104, SRCAND
        BitBlt hdc, .Pos.x, .Pos.y, 64, 104, GameField.ImgGoal.hdc, (nGoalCell - 1) * 64, 0, SRCINVERT
    End If
End With

End Sub

Public Sub Init_Placeholders()
Dim x       As Integer, y       As Integer
Dim Cnt     As Integer, CardCnt As Integer, nCard As Integer

nCard = 1
x = 100
y = 30 + 104 + 30
For Cnt = 1 To 7
    With GameField.PlaceHolder(Cnt)
        ReDim .Cards(0 To 20)
        .Pos.x = x
        .Pos.y = y
        .nCards = Cnt
        x = x + 75
        For CardCnt = 1 To Cnt
            .Cards(CardCnt) = nCard
            nCard = nCard + 1
            GameField.Card(.Cards(CardCnt)).Flag = 2
        Next CardCnt
        GameField.Card(.Cards(.nCards)).Flag = 1
    End With
Next

x = 100 + (3 * 75)
y = 30

For Cnt = 1 To 4
    With GameField.GoalCell(Cnt)
        ReDim .Cards(0 To 13)
        .Pos.x = x
        .Pos.y = y
        .nCards = 0
    End With
    x = x + 75
Next

With GameField.Deck(1)
    ReDim .Cards(0 To 52)
    .Pos.x = 100 - 75
    .Pos.y = 30
    .nCards = 52 - (nCard - 1)
    For CardCnt = nCard To 52
        .Cards(CardCnt - (nCard - 1)) = CardCnt
        GameField.Card(CardCnt).Flag = 1
    Next
End With

With GameField.Deck(2)
    ReDim .Cards(0 To 52)
    .Pos.x = (100 - 75) + 75
    .Pos.y = 30
    .nCards = 0
End With

Info.Interrupt = True

End Sub

Public Sub DrawPlaceHolders(hdc As Long)
Dim PHcnt       As Integer, CardCnt     As Integer
Dim MyPos       As POINTAPI, nY         As Integer

For PHcnt = 1 To 7
    If GameField.PlaceHolder(PHcnt).nCards > 0 Then
        MyPos.x = 0
        MyPos.y = 0
        nY = 0
        BitBlt hdc, GameField.PlaceHolder(PHcnt).Pos.x, GameField.PlaceHolder(PHcnt).Pos.y, 64, (GameField.PlaceHolder(PHcnt).nCards * 14) + 104, GameField.BG.hdc, GameField.PlaceHolder(PHcnt).Pos.x, GameField.PlaceHolder(PHcnt).Pos.y, SRCCOPY
        For CardCnt = 1 To GameField.PlaceHolder(PHcnt).nCards - 1 Step 1
            MyPos.x = GameField.PlaceHolder(PHcnt).Pos.x
            MyPos.y = GameField.PlaceHolder(PHcnt).Pos.y + nY
            If GameField.Card(GameField.PlaceHolder(PHcnt).Cards(CardCnt)).Flag = 1 Then
                DrawCard hdc, MyPos, GameField.Card(GameField.PlaceHolder(PHcnt).Cards(CardCnt))
            Else
                DrawCard hdc, MyPos, GameField.CardBack
            End If
            nY = nY + 14
        Next
        MyPos.x = GameField.PlaceHolder(PHcnt).Pos.x
        MyPos.y = GameField.PlaceHolder(PHcnt).Pos.y + nY
        If GameField.Card(GameField.PlaceHolder(PHcnt).Cards(CardCnt)).Flag = 1 Then
            DrawCard hdc, MyPos, GameField.Card(GameField.PlaceHolder(PHcnt).Cards(CardCnt))
        Else
            DrawCard hdc, MyPos, GameField.CardBack
        End If
        nY = nY + 14
    Else
        With GameField.PlaceHolder(PHcnt)
            BitBlt hdc, .Pos.x, .Pos.y, 64, GameField.nH - .Pos.y, GameField.BG.hdc, .Pos.x, .Pos.y, SRCCOPY
        End With
    End If
Next PHcnt

End Sub

Public Sub DrawPlaceHolder(nPlaceH As Byte, hdc As Long)
Dim CardCnt     As Integer
Dim MyPos       As POINTAPI, nY     As Integer

BitBlt hdc, GameField.PlaceHolder(nPlaceH).Pos.x, GameField.PlaceHolder(nPlaceH).Pos.y, 64, GameField.nH - GameField.PlaceHolder(nPlaceH).Pos.y, GameField.BG.hdc, GameField.PlaceHolder(nPlaceH).Pos.x, GameField.PlaceHolder(nPlaceH).Pos.y, SRCCOPY

If GameField.PlaceHolder(nPlaceH).nCards > 0 Then
    MyPos.x = 0
    MyPos.y = 0
    nY = 0
    For CardCnt = 1 To GameField.PlaceHolder(nPlaceH).nCards - 1 Step 1
        MyPos.x = GameField.PlaceHolder(nPlaceH).Pos.x
        MyPos.y = GameField.PlaceHolder(nPlaceH).Pos.y + nY
        If GameField.Card(GameField.PlaceHolder(nPlaceH).Cards(CardCnt)).Flag = 1 Then
            DrawCard hdc, MyPos, GameField.Card(GameField.PlaceHolder(nPlaceH).Cards(CardCnt))
        Else
            DrawCard hdc, MyPos, GameField.CardBack
        End If
        nY = nY + 14
    Next
    MyPos.x = GameField.PlaceHolder(nPlaceH).Pos.x
    MyPos.y = GameField.PlaceHolder(nPlaceH).Pos.y + nY
    If GameField.Card(GameField.PlaceHolder(nPlaceH).Cards(CardCnt)).Flag = 1 Then
        DrawCard hdc, MyPos, GameField.Card(GameField.PlaceHolder(nPlaceH).Cards(CardCnt))
    Else
        DrawCard hdc, MyPos, GameField.CardBack
    End If
    nY = nY + 14
Else
    With GameField.PlaceHolder(nPlaceH)
        BitBlt hdc, .Pos.x, .Pos.y, 64, GameField.nH - .Pos.y, GameField.BG.hdc, .Pos.x, .Pos.y, SRCCOPY
    End With
End If
    
End Sub
Public Sub DrawDeck(hdc As Long)
Dim nLines As Byte, Cnt As Integer
Dim TmpPos As POINTAPI

'Restore Background:
'---------------------------------------
With GameField.Deck(1)
    BitBlt hdc, .Pos.x, .Pos.y, 64, 104, GameField.BG.hdc, .Pos.x, .Pos.y, SRCCOPY
End With

With GameField.Deck(2)
    BitBlt hdc, .Pos.x - 11, .Pos.y, 11, 104, GameField.BG.hdc, .Pos.x - 11, .Pos.y, SRCCOPY
    BitBlt hdc, .Pos.x, .Pos.y, 64 + (15 * 2), 104, GameField.BG.hdc, .Pos.x, .Pos.y, SRCCOPY
End With
'---------------------------------------

'Draw sidelines:
'---------------------------------------
nLines = GameField.Deck(1).nCards \ 10
For Cnt = 1 To nLines
    BitBlt hdc, GameField.Deck(1).Pos.x + 64 + (3 * (Cnt - 1)), GameField.Deck(1).Pos.y, 3, 104, GameField.ImgCardBack.hdc, 64 - 3, Info.nCardBack * 104, SRCCOPY
Next Cnt
'---------------------------------------

'Draw deck1 card:
'---------------------------------------
If GameField.Deck(1).nCards > 0 Then
    DrawCard hdc, GameField.Deck(1).Pos, GameField.CardBack
Else
    DrawCard hdc, GameField.Deck(1).Pos, GameField.CardNone
End If
'---------------------------------------

'Draw deck2 card(s):
'---------------------------------------
With GameField.Deck(2)
    'Code for drawing three cards:
    '---------------------------------------
    Dim nStart As Integer
    If Info.nDrawCards = 3 Then
        If .nCards > 0 Then
            Dim nCards As Integer
            If .nCards >= 3 Then
                If Info.nActive = 1 And Info.srcType = 3 Then
                    nStart = .nCards - (2 - Info.nRemoved)
                    nCards = .nCards
                Else
                    nStart = .nCards - 2
                    nCards = .nCards
                End If
                For Cnt = nStart To nCards
                    TmpPos = .Pos
                    TmpPos.x = TmpPos.x + ((Cnt - nStart) * 15)
                    DrawCard hdc, TmpPos, GameField.Card(.Cards(.nCards - (nCards - Cnt)))
                Next Cnt
            Else
                If Info.nActive = 1 And Info.srcType = 3 And Info.nRemoved > 0 Then
                    nStart = (.nCards - 2) + Info.nRemoved
                Else
                    nStart = 1
                End If
                nCards = .nCards
                For Cnt = nStart To nCards
                    TmpPos = .Pos
                    TmpPos.x = TmpPos.x + ((Cnt - nStart) * 15)
                    DrawCard hdc, TmpPos, GameField.Card(.Cards(.nCards - (nCards - Cnt)))
                Next Cnt
            End If
            'Draw sidelines:
            '---------------------
            If Info.nDrawCards = 1 Then
                nLines = .nCards \ 10
                For Cnt = 1 To nLines
                    BitBlt hdc, GameField.Deck(2).Pos.x - (3 * Cnt), GameField.Deck(2).Pos.y, 3, 104, GameField.ImgCardFront.hdc, 0, Info.nCardBack * 104, SRCCOPY
                Next Cnt
            End If
            '---------------------
            
        Else
            
            BitBlt hdc, .Pos.x, .Pos.y, 64 + (15 * 2), 104, GameField.BG.hdc, .Pos.x, .Pos.y, SRCCOPY
        
        End If
    '---------------------------------------
    'Code for drawing one card:
    '---------------------------------------
    Else
    
        If .nCards > 0 Then
            nLines = .nCards \ 10
            DrawCard hdc, .Pos, GameField.Card(.Cards(.nCards))
            For Cnt = 1 To nLines
                BitBlt hdc, GameField.Deck(2).Pos.x - (3 * Cnt), GameField.Deck(2).Pos.y, 3, 104, GameField.ImgCardFront.hdc, 0, Info.nCardBack * 104, SRCCOPY
            Next Cnt
        Else
            BitBlt hdc, .Pos.x, .Pos.y, 64, 104, GameField.BG.hdc, .Pos.x, .Pos.y, SRCCOPY
        End If
    End If
    '---------------------------------------
End With
'---------------------------------------

End Sub

Public Sub CreateGame(Flag As Byte)

If Flag = 1 Then
    'Shuffle Cards:
    Dim Cnt As Long, Cnt2 As Long
    Dim A As Single, B As Single
    Dim Tmp As tCard
    Dim Num(1 To 2) As Long

    Randomize 'Start random number generator:
    'Shuffle through all the cards some times:
    For Cnt2 = 1 To 50
        For Cnt = 1 To 52
            A = Cnt
            B = Int(Rnd * 52) + 1
            'Switch:
            Tmp = GameField.Card(B)
            GameField.Card(B) = GameField.Card(A)
            GameField.Card(A) = Tmp
        Next Cnt
    Next Cnt2

    'Pick random cards and switch 'em:
    For Cnt = 1 To 10000
        A = Int(Rnd * 52) + 1
        B = Int(Rnd * 52) + 1
    
        Tmp = GameField.Card(B)               '   \
        GameField.Card(B) = GameField.Card(A) '    }Switch
        GameField.Card(A) = Tmp               '   /
    
    Next
End If

Info.Undo.Available = False

Init_Placeholders

'Restart game timer:
Info.StartTime = GetTickCount
frmMain.tmrUpdateTime.Enabled = True

DrawGameField

End Sub

Public Sub FlipScreen()

'Blit the backcuffer to screen:
BitBlt frmMain.hdc, 0, 0, GameField.nW, GameField.nH, GameField.BB.hdc, 0, 0, SRCCOPY

End Sub
Public Sub Init_Res()

With GameField

    'If Info.Debugging Then frmMain.Caption = frmMain.Caption & "  [BitsPerPixel]:"
    .Img15.fPath = Info.AppPath & "SORT.BMP"            'Sort specific signs
    .ImgCardFront.fPath = Info.AppPath & "FRONT.BMP"    'Card fronts
    .ImgCardBack.fPath = Info.AppPath & "BACK.BMP"      'Card backs
    .ImgSign.fPath = Info.AppPath & "SIGN.BMP"          'Card Signs (number/letters)
    .ImgCard.fPath = Info.AppPath & "IMGCARD.BMP"       'Image cards
    .ImgGoal.fPath = Info.AppPath & "GOAL.BMP"          'Goal Cell Images
    .Img790.fPath = Info.AppPath & "790.BMP"            '790 robot head
    .Img790ani.fPath = Info.AppPath & "790ani.BMP"            '790 robot head animation
    
    LoadRes .Img15
    LoadRes .ImgCardFront
    LoadRes .ImgCardBack
    LoadRes .ImgSign
    LoadRes .ImgCard
    LoadRes .ImgGoal
    LoadRes .Img790
    LoadRes .Img790ani

    .BG = MakeMemBMP(frmMain.ScaleWidth, frmMain.ScaleHeight)
    .BB = MakeMemBMP(frmMain.ScaleWidth, frmMain.ScaleHeight)
    .BufTmp = MakeMemBMP(128, 208 + (14 * 12))
End With

'CardFronts:
'-----------------------------------
Dim nDecks As Byte, Cnt As Byte
nDecks = GameField.ImgCardFront.nH \ 104
Info.nCardfronts = nDecks
If nDecks > 1 Then
    For Cnt = 2 To nDecks
        Load frmMain.mnuOptionsLayoutFrontFront(Cnt - 1)
        frmMain.mnuOptionsLayoutFrontFront(Cnt - 1).Caption = cCARDFRONTFRONT & Trim(Str(Cnt))
    Next Cnt
End If
'-----------------------------------

'CardBacks:
'-----------------------------------
nDecks = GameField.ImgCardBack.nH \ 104
Info.nCardBacks = nDecks
If nDecks > 1 Then
    For Cnt = 2 To nDecks
        Load frmMain.mnuOptionsLayoutBackBack(Cnt - 1)
        frmMain.mnuOptionsLayoutBackBack(Cnt - 1).Caption = cCARDBACKBACK & Trim(Str(Cnt))
    Next Cnt
End If
'-----------------------------------

'Backgrounds:
'-----------------------------------
Dim nBGs As Byte, strPath As String
strPath = Info.AppPath & "BG\BG1.BMP"

While Dir(strPath) <> vbNullString
    nBGs = nBGs + 1
    strPath = Info.AppPath & "BG\BG" & Trim(Str(nBGs)) & ".BMP"
Wend

nBGs = nBGs - 1
Info.nBGs = nBGs

If nBGs > 1 Then
    For Cnt = 2 To nBGs
        Load frmMain.mnuOptionsLayoutBgBg(Cnt - 1)
        frmMain.mnuOptionsLayoutBgBg(Cnt - 1).Caption = cBGBG & Trim(Str(Cnt))
    Next Cnt
End If

'Sounds:
'-----------------------------------
Info.nSnds = 0
strPath = Info.AppPath & "Snd1.wav"

While Dir(strPath) <> vbNullString
    Info.nSnds = Info.nSnds + 1
    strPath = Info.AppPath & "Snd" & Trim(Str(Info.nSnds)) & ".wav"
Wend
Info.nSnds = Info.nSnds - 1

If Info.nSnds > 1 Then
    For Cnt = 2 To Info.nSnds
        Load frmMain.mnuOptionsSoundsDefsndSnd(Cnt - 1)
        Load frmMain.mnuOptionsSoundsGoalsndSnd(Cnt - 1)
        frmMain.mnuOptionsSoundsDefsndSnd(Cnt - 1).Caption = cSOUNDSDEFSNDSND & Trim(Str(Cnt))
        frmMain.mnuOptionsSoundsGoalsndSnd(Cnt - 1).Caption = cSOUNDSGOALSNDSND & Trim(Str(Cnt))
    Next
End If

CheckSettings

'-----------------------------------

GameField.TmpBG.fPath = Info.AppPath & "BG\BG" & Trim(Str(Info.nBG + 1)) & ".BMP"
LoadRes GameField.TmpBG
LoadSnd Info.AppPath & "Win.wav", GameField.WinSnd
LoadSnd Info.AppPath & "Win2.wav", GameField.Win2Snd
MakeBG

'-----------------------------------

End Sub

Public Function MakeMemBMP(nW As Long, nH As Long) As tResource
Dim MyBMP As BITMAP

With MakeMemBMP
    .hdc = 0
    .hdc = CreateCompatibleDC(0)
    .hBMP = 0
    .hBMP = CreateCompatibleBitmap(frmMain.hdc, nW, nH)
    .hOldBMP = SelectObject(.hdc, .hBMP)
    MyGetObject .hBMP, Len(MyBMP), MyBMP
    'If Info.Debugging Then frmMain.Caption = frmMain.Caption & " " & Trim(Str(MyBMP.bmBitsPixel))
    .nW = nW: .nH = nH
    If .hBMP = 0 Then MsgBox "Failed to create memory bitmap!"
End With

End Function

Public Sub DrawGameField()

'Draw Background:
BitBlt GameField.BB.hdc, 0, 0, GameField.BB.nW, GameField.BB.nH, GameField.BG.hdc, 0, 0, SRCCOPY

'Eventually Draw 790 robot head:
If Info.Show790 Then
    With GameField
        BitBlt .BB.hdc, .Pos790.x, .Pos790.y, .Img790.nW, .Img790.nH \ 2, .Img790.hdc, 0, .Img790.nH \ 2, SRCAND
        BitBlt .BB.hdc, .Pos790.x, .Pos790.y, .Img790.nW, .Img790.nH \ 2, .Img790.hdc, 0, 0, SRCINVERT
    End With
End If

'Draw Placeholders w/cards:
DrawPlaceHolders GameField.BB.hdc
DrawGoalCells GameField.BB.hdc
DrawDeck GameField.BB.hdc

'Fix Statusbar:
With frmMain.Status
    .Panels(1).Width = 100
    .Panels(2).Width = GameField.nW - .Panels(1).Width
End With

'Blit backbuffer to screen:
FlipScreen

End Sub
Public Function CheckMouseDown_Placeholders(x As Single, y As Single) As Boolean
Dim Cnt As Byte, nX As Integer, nY As Integer, nCard As Integer, nPlaceH As Byte
Dim CardCnt As Integer

CheckMouseDown_Placeholders = False
'Check if the point is inside one of the placeholders or goalcells:
For Cnt = 1 To 7
    If x >= GameField.PlaceHolder(Cnt).Pos.x And _
       x < (GameField.PlaceHolder(Cnt).Pos.x + 64) And _
       y >= GameField.PlaceHolder(Cnt).Pos.y And _
       y <= (GameField.PlaceHolder(Cnt).Pos.y + 104 + (GameField.PlaceHolder(Cnt).nCards * 14)) _
    Then
        'User clicked inside one of them, set up active card:
        nX = x - GameField.PlaceHolder(Cnt).Pos.x
        nY = y - GameField.PlaceHolder(Cnt).Pos.y
        For CardCnt = 1 To GameField.PlaceHolder(Cnt).nCards - 1
            If nY >= ((CardCnt - 1) * 14) And nY < ((CardCnt * 14)) Then
                nCard = CardCnt
                nPlaceH = Cnt
            End If
        Next
        If nY >= (GameField.PlaceHolder(Cnt).nCards - 1) * 14 And nY < ((GameField.PlaceHolder(Cnt).nCards - 1) * 14) + 104 Then
            nCard = GameField.PlaceHolder(Cnt).nCards
            nPlaceH = Cnt
        End If
        Info.Moving = True
    End If
Next

If nPlaceH < 1 Or nCard < 1 Then Exit Function

If GameField.Card(GameField.PlaceHolder(nPlaceH).Cards(nCard)).Flag = 1 Then
    For Cnt = nCard To GameField.PlaceHolder(nPlaceH).nCards
        Info.ActiveCard(Cnt - (nCard - 1)) = GameField.PlaceHolder(nPlaceH).Cards(Cnt)
    Next
    Info.nActive = GameField.PlaceHolder(nPlaceH).nCards - (nCard - 1)
    GameField.PlaceHolder(nPlaceH).nCards = GameField.PlaceHolder(nPlaceH).nCards - (Info.nActive)
    Info.ClickPos.x = nX
    Info.ClickPos.y = nY - ((nCard - 1) * 14)
    Info.srcPH = nPlaceH
    Info.srcType = 1
    Info.Moving = True
    DrawPlaceHolder nPlaceH, GameField.BB.hdc
    DrawActiveCard x, y
Else
    If nCard = GameField.PlaceHolder(nPlaceH).nCards Then
        'Turn the card:
        GameField.Card(GameField.PlaceHolder(nPlaceH).Cards(nCard)).Flag = 1
        DrawPlaceHolder nPlaceH, GameField.BB.hdc
        FlipScreen
        Info.Moving = False
        Info.nActive = 0
        Info.Undo.Available = False 'If the user has turned a card, things will be messed up after Undo..
    End If
End If

CheckMouseDown_Placeholders = True

End Function



Public Function CheckMouseDown_GoalCells(x As Single, y As Single) As Boolean
Dim Cnt As Byte, nX As Integer, nY As Integer, nCard As Integer, nPlaceH As Byte
Dim CardCnt As Integer

CheckMouseDown_GoalCells = False
'Check if the point is inside one of the goalcells:
For Cnt = 1 To 4
    If x >= GameField.GoalCell(Cnt).Pos.x And _
        x < (GameField.GoalCell(Cnt).Pos.x + 64) And _
        y >= GameField.GoalCell(Cnt).Pos.y And _
        y <= (GameField.GoalCell(Cnt).Pos.y + 104) _
    Then
        'User clicked inside one of them, set up active card:
        nPlaceH = Cnt
        nCard = GameField.GoalCell(Cnt).nCards
        nX = x - GameField.GoalCell(Cnt).Pos.x
        nY = y - GameField.GoalCell(Cnt).Pos.y
    End If
Next

If nPlaceH = 0 Or nCard = 0 Then Exit Function

Info.ActiveCard(1) = GameField.GoalCell(nPlaceH).Cards(nCard)
Info.nActive = 1
GameField.GoalCell(nPlaceH).nCards = GameField.GoalCell(nPlaceH).nCards - 1
Info.ClickPos.x = nX
Info.ClickPos.y = nY
Info.srcPH = nPlaceH + 7
Info.srcType = 2
Info.Moving = True
DrawGoalCell nPlaceH, GameField.BB.hdc
DrawActiveCard x, y

CheckMouseDown_GoalCells = True

End Function
Public Function CheckMouseUp_GoalCells(x As Single, y As Single) As Boolean
Dim Xpos As Integer, Ypos As Integer
Dim SrcX As Integer, SrcY As Integer
Dim DestX As Integer, DestY As Integer
Dim FrameCnt As Integer, nFrames As Integer
Dim nPlaceH As Byte, Cnt As Byte

'Set the function to false temporarily:
CheckMouseUp_GoalCells = False

'Do stuff:
If Info.Moving Then

    'Find out if the card has landed in a goalcell, and if it is a valid move:
    '---------------------------------------------------------
    For Cnt = 1 To 4
        If x - Info.ClickPos.x >= GameField.GoalCell(Cnt).Pos.x - 37 And _
            x - Info.ClickPos.x < (GameField.GoalCell(Cnt).Pos.x + 37) And _
            y - Info.ClickPos.y >= GameField.GoalCell(Cnt).Pos.y - 57 And _
            y - Info.ClickPos.y <= GameField.GoalCell(Cnt).Pos.y + 57 _
        Then
            nPlaceH = Cnt
            Info.Moving = True
            Exit For
        End If
    Next Cnt

    If nPlaceH = 0 Then Exit Function
    If Info.nActive > 1 Then Exit Function
    If GameField.GoalCell(nPlaceH).nCards = 0 And _
       GameField.Card(Info.ActiveCard(1)).Value <> 1 Then Exit Function

    If nPlaceH > 0 And GameField.Card(Info.ActiveCard(1)).Flag = 1 And _
    GameField.Card(Info.ActiveCard(1)).Sort + 1 = nPlaceH And _
    GameField.Card(Info.ActiveCard(1)).Value = GameField.Card(GameField.GoalCell(nPlaceH).Cards(GameField.GoalCell(nPlaceH).nCards)).Value + 1 Or _
    (GameField.Card(Info.ActiveCard(1)).Value = 1 And GameField.GoalCell(nPlaceH).nCards = 0) And _
    GameField.Card(Info.ActiveCard(1)).Sort + 1 = nPlaceH _
    Then

        'Valid Move: Transfer the card to the goalcell:
        '---------------------------------------------------------
        GameField.GoalCell(nPlaceH).nCards = GameField.GoalCell(nPlaceH).nCards + 1
        GameField.GoalCell(nPlaceH).Cards(GameField.GoalCell(nPlaceH).nCards) = Info.ActiveCard(1)
        Info.Undo.nCards = 1
        Info.Undo.nSrc = Info.srcPH
        Info.Undo.nDest = nPlaceH + 7
        Info.Undo.Available = True
        
        'Remove the active card:
        Info.nActive = 0
        Info.Moving = False

        'Draw the placeholder:
        DrawGoalCell nPlaceH, GameField.BB.hdc
        
        If Info.srcType <> 2 And Info.PlaySounds Then
            'Play sound:
            PlaySnd GameField.GoalSnd
        End If
        
        If Info.srcType = 3 And Info.nRemoved = 3 Then
            Info.nRemoved = 0
            DrawDeck GameField.BB.hdc
        End If
        
        FlipScreen
        CheckMouseUp_GoalCells = True
    
    End If
End If

'Check if the user has won:
CheckWinner

End Function
Public Sub CheckWinner()
Dim Win As Boolean, Cnt As Byte

If Info.CheckingWinner Then Exit Sub

Info.CheckingWinner = True

Win = True
For Cnt = 1 To 4
    If GameField.GoalCell(Cnt).nCards <> 13 Then Win = False
Next Cnt

If Win = True Then
    'Play Sound
    sndPlaySound GameField.WinSnd, SND_MEMORY Or SND_SYNC
    DoEffect 'Show winning effect
    
    If MsgBox("Do you want to start a new game?", vbQuestion Or vbYesNo, "New Game?") = 6 Then
        CreateGame 1
        DrawGameField
        FlipScreen
    Else
        CleanUp
        End
    End If
End If

Info.CheckingWinner = False

End Sub
Public Sub DoEffect()

'Do winning effect:
'----------------------------------------------
Dim x As Single, y As Single, MyPos As POINTAPI
Dim XSpeed As Single, YSpeed As Single
'Dim T1 As Long, T2 As Long
Dim nCard As Byte, TmpCard As tCard
Dim MaxSpeed As Single
Dim nDir As Single

'Disable 790:
frmMain.tmr790Blink.Enabled = False

Randomize 'Initialize random number generator
MaxSpeed = 20
Info.Interrupt = False
nCard = 1
x = GameField.GoalCell(1).Pos.x '= Int((GameField.nW \ 2) - 32)
y = GameField.GoalCell(1).Pos.y '= Int((GameField.nH \ 2) - 52)
XSpeed = 4
YSpeed = 3
'T1 = GetTickCount
'While T2 - T1 < 60000 And Not Info.Interrupt

sndPlaySound GameField.Win2Snd, SND_MEMORY Or SND_ASYNC Or SND_LOOP

While Not Info.Interrupt

    MyPos.x = x
    MyPos.y = y
    TmpCard = GameField.Card(nCard)
    TmpCard.Flag = 1

    DrawCard GameField.BB.hdc, MyPos, TmpCard
    BitBlt frmMain.hdc, x, y, 64, 104, GameField.BB.hdc, x, y, SRCCOPY
    
    If x + XSpeed > (GameField.nW - 64) Or x + XSpeed < 0 Then
        XSpeed = (-XSpeed) + ((Int(Rnd * 9) + 1) - 5)
        YSpeed = YSpeed + ((Int(Rnd * 9) + 1) - 5)
    End If
    If y + YSpeed > (GameField.nH - 104) Or y + YSpeed < 0 Then
        YSpeed = (-YSpeed) + ((Int(Rnd * 9) + 1) - 5)
        XSpeed = XSpeed + ((Int(Rnd * 9) + 1) - 5)
    End If
    
    x = x + XSpeed
    y = y + YSpeed
    
    XSpeed = XSpeed + ((Int(Rnd * 3) + 1) - 2)
    YSpeed = YSpeed + ((Int(Rnd * 3) + 1) - 2)
    
    nCard = nCard + 1
    If nCard > 52 Then nCard = 1: DoEvents
    'T2 = GetTickCount
Wend

'Stop sound:
sndPlaySound "", SND_PURGE

'----------------------------------------------
'Update Gamefield:
DrawGameField
'Enable 790:
frmMain.tmr790Blink.Enabled = True

End Sub

Public Function CheckMouseDown_Deck(x As Single, y As Single) As Boolean
Dim Cnt As Integer, nX As Integer, nY As Integer, nCard As Integer, nPlaceH As Byte
Dim CardCnt As Integer
Dim XtraPix As Integer

If Info.nDrawCards = 3 Then
    If GameField.Deck(2).nCards >= 3 Then
        XtraPix = (2 - Info.nRemoved) * 15
    Else
        'XtraPix = (GameField.Deck(2).nCards - 1) * 15
        If Info.nRemoved + GameField.Deck(2).nCards >= 3 Then
            XtraPix = (2 - Info.nRemoved) * 15
        Else
            XtraPix = (GameField.Deck(2).nCards - 1) * 15
        End If
    End If
End If

CheckMouseDown_Deck = False
'Check if the point is inside one of the goalcells:
    If x >= GameField.Deck(1).Pos.x And _
        x < (GameField.Deck(1).Pos.x + 64) And _
        y >= GameField.Deck(1).Pos.y And _
        y <= (GameField.Deck(1).Pos.y + 104) Then
        'User clicked inside deck 1
        nPlaceH = 1
        nCard = GameField.Deck(1).nCards
        nX = x - GameField.Deck(1).Pos.x
        nY = y - GameField.Deck(1).Pos.y
        
    ElseIf x >= GameField.Deck(2).Pos.x + XtraPix And _
        x < (GameField.Deck(2).Pos.x + 64 + XtraPix) And _
        y >= GameField.Deck(2).Pos.y And _
        y <= (GameField.Deck(2).Pos.y + 104) Then
        'User clicked inside deck 1
        nPlaceH = 2
        nCard = GameField.Deck(2).nCards
        nX = x - GameField.Deck(2).Pos.x
        nY = y - GameField.Deck(2).Pos.y
    End If

If nPlaceH = 0 And nCard = 0 Then Exit Function

Select Case nPlaceH
Case 1
    
    If Info.nDrawCards = 3 Then
    'Code for drawing three cards:
    '--------------------------------------------
        If GameField.Deck(1).nCards >= 3 Then
            Info.nRemoved = 0
            For Cnt = 1 To 3
                GameField.Deck(2).nCards = GameField.Deck(2).nCards + 1
                GameField.Deck(2).Cards(GameField.Deck(2).nCards) = GameField.Deck(1).Cards(GameField.Deck(1).nCards)
                GameField.Deck(1).nCards = GameField.Deck(1).nCards - 1
            Next Cnt
            DrawDeck GameField.BB.hdc
            BitBlt frmMain.hdc, GameField.Deck(1).Pos.x, GameField.Deck(1).Pos.y, 75 + 64 + (15 * 2), 104, GameField.BB.hdc, GameField.Deck(1).Pos.x, GameField.Deck(1).Pos.y, SRCCOPY
        ElseIf GameField.Deck(1).nCards > 0 Then
            For Cnt = 1 To GameField.Deck(1).nCards
                GameField.Deck(2).nCards = GameField.Deck(2).nCards + 1
                GameField.Deck(2).Cards(GameField.Deck(2).nCards) = GameField.Deck(1).Cards(Cnt)
            Next Cnt
            GameField.Deck(1).nCards = 0
            DrawDeck GameField.BB.hdc
            BitBlt frmMain.hdc, GameField.Deck(1).Pos.x, GameField.Deck(1).Pos.y, 75 + 64 + (15 * 2), 104, GameField.BB.hdc, GameField.Deck(1).Pos.x, GameField.Deck(1).Pos.y, SRCCOPY
        Else
            If GameField.Deck(2).nCards > 0 Then
                For Cnt = 1 To GameField.Deck(2).nCards
                    GameField.Deck(1).Cards(Cnt) = GameField.Deck(2).Cards(GameField.Deck(2).nCards - (Cnt - 1))
                Next Cnt
                GameField.Deck(1).nCards = GameField.Deck(2).nCards
                GameField.Deck(2).nCards = 0
                DrawDeck GameField.BB.hdc
                BitBlt frmMain.hdc, GameField.Deck(1).Pos.x, GameField.Deck(1).Pos.y, 75 + 64 + (15 * 2), 104, GameField.BB.hdc, GameField.Deck(1).Pos.x, GameField.Deck(1).Pos.y, SRCCOPY
            End If
        End If
    '--------------------------------------------
    Else
    'Code for drawing one card at at time:
    '--------------------------------------------
        If GameField.Deck(1).nCards > 0 Then
            GameField.Deck(2).nCards = GameField.Deck(2).nCards + 1
            GameField.Deck(2).Cards(GameField.Deck(2).nCards) = GameField.Deck(1).Cards(GameField.Deck(1).nCards)
            GameField.Deck(1).nCards = GameField.Deck(1).nCards - 1
            DrawDeck GameField.BB.hdc
            BitBlt frmMain.hdc, GameField.Deck(1).Pos.x, GameField.Deck(1).Pos.y, 75 + 64, 104, GameField.BB.hdc, GameField.Deck(1).Pos.x, GameField.Deck(1).Pos.y, SRCCOPY
        Else
            If GameField.Deck(2).nCards > 0 Then
                For Cnt = 1 To GameField.Deck(2).nCards
                    GameField.Deck(1).Cards(Cnt) = GameField.Deck(2).Cards(GameField.Deck(2).nCards - (Cnt - 1))
                Next Cnt
                GameField.Deck(1).nCards = GameField.Deck(2).nCards
                GameField.Deck(2).nCards = 0
                DrawDeck GameField.BB.hdc
                BitBlt frmMain.hdc, GameField.Deck(1).Pos.x, GameField.Deck(1).Pos.y, 75 + 64, 104, GameField.BB.hdc, GameField.Deck(1).Pos.x, GameField.Deck(1).Pos.y, SRCCOPY
            End If
        End If
    End If
    
    Info.Undo.Available = False
    '--------------------------------------------

Case 2  'User wants to drag a card from the deck:
    'Code for drawing three cards:
    '--------------------------------------------
    If Info.nDrawCards = 3 Then
        If GameField.Deck(2).nCards > 0 Then
            Info.ActiveCard(1) = GameField.Deck(2).Cards(nCard)
            Info.nActive = 1
            GameField.Deck(2).nCards = GameField.Deck(2).nCards - 1
            Info.ClickPos.x = nX - XtraPix
            Info.ClickPos.y = nY
            Info.srcPH = nPlaceH + 11
            Info.srcType = 3
            Info.Moving = True
            Info.nRemoved = Info.nRemoved + 1
            'If Info.nRemoved = 3 Then Info.nRemoved = 0
            DrawDeck GameField.BB.hdc
            With GameField.Deck(2)
                BitBlt frmMain.hdc, .Pos.x, .Pos.y, (3 - Info.nRemoved) * 15, 104, GameField.BB.hdc, .Pos.x, .Pos.y, SRCCOPY
            End With
            DrawActiveCard x, y
        End If
    '--------------------------------------------
    'Code for drawing one card at a time:
    '--------------------------------------------
    Else
        If GameField.Deck(2).nCards > 0 Then
            Info.ActiveCard(1) = GameField.Deck(2).Cards(nCard)
            Info.nActive = 1
            GameField.Deck(2).nCards = GameField.Deck(2).nCards - 1
            Info.ClickPos.x = nX
            Info.ClickPos.y = nY
            Info.srcPH = nPlaceH + 11
            Info.srcType = 3
            Info.Moving = True
            DrawDeck GameField.BB.hdc
            DrawActiveCard x, y
    End If
    End If
    '--------------------------------------------
End Select

'If the function wasn't left previously, it returns true:
CheckMouseDown_Deck = True

End Function
Public Function CheckDoubleClick(x As Long, y As Long) As Boolean
Dim Cnt As Byte, nX As Long, nY As Long, nCard As Byte, nPlaceH As Byte, CardCnt As Integer
Dim TmpCard(1 To 2) As tCard

'Set function temporarily to false:
CheckDoubleClick = False

If Info.Moving Then Exit Function

For Cnt = 1 To 7
    If x >= GameField.PlaceHolder(Cnt).Pos.x And _
        x < (GameField.PlaceHolder(Cnt).Pos.x + 64) And _
        y >= GameField.PlaceHolder(Cnt).Pos.y And _
        y <= (GameField.PlaceHolder(Cnt).Pos.y + 104 + (GameField.PlaceHolder(Cnt).nCards * 14)) _
    Then
        'User clicked inside one of them, set up active card:
        nX = x - GameField.PlaceHolder(Cnt).Pos.x
        nY = y - GameField.PlaceHolder(Cnt).Pos.y
        For CardCnt = 1 To GameField.PlaceHolder(Cnt).nCards - 1
            If nY >= ((CardCnt - 1) * 14) And nY < ((CardCnt * 14)) Then
                nCard = CardCnt
                nPlaceH = Cnt
            End If
        Next
        If nY >= (GameField.PlaceHolder(Cnt).nCards - 1) * 14 And nY < ((GameField.PlaceHolder(Cnt).nCards - 1) * 14) + 104 Then
            nCard = GameField.PlaceHolder(Cnt).nCards
            nPlaceH = Cnt
        End If
    End If
Next

If nPlaceH > 0 Then
    With GameField.PlaceHolder(nPlaceH)
        If Not GameField.Card(.Cards(.nCards)).Flag = 1 Then Exit Function
    End With
End If

If nPlaceH > 0 Then
    If nCard = GameField.PlaceHolder(nPlaceH).nCards Then
        TmpCard(1) = GameField.Card(GameField.PlaceHolder(nPlaceH).Cards(GameField.PlaceHolder(nPlaceH).nCards))
        TmpCard(2) = GameField.Card(GameField.GoalCell(TmpCard(1).Sort + 1).Cards(GameField.GoalCell(TmpCard(1).Sort + 1).nCards))
        If TmpCard(1).Value = TmpCard(2).Value + 1 Then
            GameField.GoalCell(TmpCard(1).Sort + 1).nCards = GameField.GoalCell(TmpCard(1).Sort + 1).nCards + 1
            GameField.GoalCell(TmpCard(1).Sort + 1).Cards(GameField.GoalCell(TmpCard(1).Sort + 1).nCards) = GameField.PlaceHolder(nPlaceH).Cards(GameField.PlaceHolder(nPlaceH).nCards)
            GameField.PlaceHolder(nPlaceH).nCards = GameField.PlaceHolder(nPlaceH).nCards - 1
            DrawGoalCell TmpCard(1).Sort + 1, GameField.BB.hdc
            DrawGoalCell TmpCard(1).Sort + 1, frmMain.hdc
            DrawPlaceHolder nPlaceH, GameField.BB.hdc
            With GameField.PlaceHolder(nPlaceH)
                BitBlt frmMain.hdc, .Pos.x, .Pos.y, 64, GameField.nH - .Pos.y, GameField.BB.hdc, .Pos.x, .Pos.y, SRCCOPY
            End With
            Info.Undo.nCards = 1
            Info.Undo.nSrc = nPlaceH
            Info.Undo.nDest = 7 + TmpCard(1).Sort + 1
            Info.Undo.Available = True
            PlaySnd GameField.GoalSnd
            CheckDoubleClick = True
        End If
    End If
Else
    'Check deck:
    If x >= GameField.Deck(2).Pos.x And x <= GameField.Deck(2).Pos.x + 64 And y >= GameField.Deck(2).Pos.y And y <= GameField.Deck(2).Pos.y + 104 Then
        TmpCard(1) = GameField.Card(GameField.Deck(2).Cards(GameField.Deck(2).nCards))
        TmpCard(2) = GameField.Card(GameField.GoalCell(TmpCard(1).Sort + 1).Cards(GameField.GoalCell(TmpCard(1).Sort + 1).nCards))
        If TmpCard(1).Value = TmpCard(2).Value + 1 Then
            GameField.GoalCell(TmpCard(1).Sort + 1).nCards = GameField.GoalCell(TmpCard(1).Sort + 1).nCards + 1
            GameField.GoalCell(TmpCard(1).Sort + 1).Cards(GameField.GoalCell(TmpCard(1).Sort + 1).nCards) = GameField.Deck(2).Cards(GameField.Deck(2).nCards)
            GameField.Deck(2).nCards = GameField.Deck(2).nCards - 1
            DrawGoalCell TmpCard(1).Sort + 1, GameField.BB.hdc
            DrawGoalCell TmpCard(1).Sort + 1, frmMain.hdc
            Info.nRemoved = Info.nRemoved + 1
            If Info.nRemoved = 3 Then Info.nRemoved = 0
            Info.nActive = 1
            DrawDeck GameField.BB.hdc
            FlipScreen
            'DrawDeck frmMain.hDC
            Info.nActive = 0
            'With GameField.Deck(1)
                'BitBlt frmMain.hDC, .Pos.X, .Pos.Y, 139 + (15 * 2), 104, GameField.BB.hDC, .Pos.X, .Pos.Y, SRCCOPY
            'End With
            Info.Undo.nCards = 1
            Info.Undo.nSrc = 13
            Info.Undo.nDest = 7 + TmpCard(1).Sort + 1
            Info.Undo.Available = True
            PlaySnd GameField.GoalSnd
            CheckDoubleClick = True
        End If
    Else
        'Check other deck:
        If x >= GameField.Deck(1).Pos.x And x <= GameField.Deck(1).Pos.x + 64 And y >= GameField.Deck(1).Pos.y And y <= GameField.Deck(1).Pos.y + 104 Then
            frmMain.TriggerClick x, y
            CheckDoubleClick = True
        End If
    End If
End If
CheckWinner 'Check if the player has won:

End Function
Public Sub SaveSettings()
Dim Section As String
Section = Info.UserName

SaveSetting APPNAME, Section, "Front", Info.nCardFront
SaveSetting APPNAME, Section, "Deck", Info.nCardBack
SaveSetting APPNAME, Section, "BG", Info.nBG
SaveSetting APPNAME, Section, "Draw", Info.nDrawCards
SaveSetting APPNAME, Section, "OrdinaryMove", Info.nDefSnd
SaveSetting APPNAME, Section, "GoalMove", Info.nGoalSnd
SaveSetting APPNAME, Section, "Language", Info.Language

If Info.Show790 Then
    SaveSetting APPNAME, Section, "Show790", 1
Else
    SaveSetting APPNAME, Section, "Show790", 0
End If

If Info.PlaySounds Then
    SaveSetting APPNAME, Section, "PlaySounds", 1
Else
    SaveSetting APPNAME, Section, "PlaySounds", 0
End If

If Info.Debugging Then
    SaveSetting APPNAME, Section, "Version", "Debug"
Else
    SaveSetting APPNAME, Section, "Version", "Release"
End If

End Sub
Public Sub ReadSettings()
Dim Section As String
Section = Info.UserName

Info.nCardFront = GetSetting(APPNAME, Section, "Front", 0)
Info.nCardBack = GetSetting(APPNAME, Section, "Deck", 0)
Info.nBG = GetSetting(APPNAME, Section, "BG", 0)
Info.nDrawCards = GetSetting(APPNAME, Section, "Draw", 1)
Info.nDefSnd = GetSetting(APPNAME, Section, "OrdinaryMove", 0)
Info.nGoalSnd = GetSetting(APPNAME, Section, "GoalMove", 1)

If Info.nDrawCards <> 1 And Info.nDrawCards <> 3 Then Info.nDrawCards = 1

If GetSetting(APPNAME, Section, "Show790", 1) = 0 Then
    Info.Show790 = False
    frmMain.tmr790Blink.Enabled = False
Else
    Info.Show790 = True
    frmMain.tmr790Blink.Enabled = True
End If

If GetSetting(APPNAME, Section, "PlaySounds", 1) = 1 Then
    Info.PlaySounds = True
Else
    Info.PlaySounds = False
End If

If GetSetting(APPNAME, Section, "Version", "Release") = "Release" Then
    Info.Debugging = False
    frmMain.mnuTopDebug.Visible = False
Else
    Info.Debugging = True
    frmMain.mnuTopDebug.Visible = True
End If

Info.Language = GetSetting(APPNAME, Section, "Language", "EN")
If Info.Language <> "EN" And Info.Language <> "NO" And Info.Language <> "DE" Then Info.Language = "EN"

End Sub

Public Sub Init_Sounds()

LoadSnd Info.AppPath & "Snd" & Trim(Str(Info.nDefSnd + 1)) & ".wav", GameField.DefaultSnd
LoadSnd Info.AppPath & "Snd" & Trim(Str(Info.nGoalSnd + 1)) & ".wav", GameField.GoalSnd

End Sub

Public Sub CheckSettings()

If Info.nCardFront + 1 > Info.nCardfronts Or Info.nCardFront + 1 < 0 Then Info.nCardFront = 0
If Info.nCardBack + 1 > Info.nCardBacks Or Info.nCardBack + 1 < 0 Then Info.nCardBack = 0
If Info.nDefSnd + 1 > Info.nSnds Or Info.nDefSnd + 1 < 0 Then Info.nDefSnd = 0
If Info.nGoalSnd + 1 > Info.nSnds Or Info.nGoalSnd + 1 < 0 Then Info.nGoalSnd = 0

If Dir(Info.AppPath & "BG\BG" & Trim(Str(Info.nBG + 1)) & ".BMP") = "" Then
    Info.nBG = 0
End If

End Sub
Public Sub MakeBG()
Dim nTilesX As Long, nTilesY As Long
Dim Xpos As Long, Ypos As Long
Dim ScrW As Long, ScrH As Long, BGW As Long, BGH As Long
Dim SrcDC As Long, DestDC As Long

BGW = GameField.TmpBG.nW
BGH = GameField.TmpBG.nH

ScrW = GameField.nW
ScrH = GameField.nH

nTilesX = ScrW \ BGW
If nTilesX * BGW < ScrW Then nTilesX = nTilesX + 1

nTilesY = ScrH \ BGH
If nTilesY * BGH < ScrH Then nTilesY = nTilesY + 1

SrcDC = GameField.TmpBG.hdc
DestDC = GameField.BG.hdc

Dim nRes As Long
For Ypos = 1 To nTilesY
    For Xpos = 1 To nTilesX
        nRes = BitBlt(DestDC, (Xpos - 1) * BGW, (Ypos - 1) * BGH, BGW, BGH, SrcDC, 0, 0, SRCCOPY)
    Next Xpos
Next Ypos

FreeRes GameField.TmpBG

End Sub
Public Sub LoadSnd(fPath As String, Buffer As String)
Dim fNum As Integer

If Dir(fPath) = vbNullString Then GoTo Error_LoadSnd

fNum = FreeFile
Open fPath For Binary As fNum
    Buffer = Space$(FileLen(fPath))
    Get #fNum, 1, Buffer
Close #fNum

Exit Sub
Error_LoadSnd:
MsgBox "Unable to load file " & fPath
Call CleanUp: End

End Sub
Public Sub PlaySnd(Buffer As String)

If Info.PlaySounds Then sndPlaySound Buffer, SND_MEMORY Or SND_ASYNC

End Sub

Public Sub Undo()
Dim Cnt As Byte
Dim Cards() As Byte

If Not Info.Undo.Available Then Exit Sub

With Info.Undo
    ReDim Cards(1 To .nCards)
    If .nDest < 8 Then
        For Cnt = 1 To .nCards
            Cards(Cnt) = GameField.PlaceHolder(.nDest).Cards(GameField.PlaceHolder(.nDest).nCards)
            GameField.PlaceHolder(.nDest).nCards = GameField.PlaceHolder(.nDest).nCards - 1
        Next Cnt
        DrawPlaceHolder .nDest, GameField.BB.hdc
    ElseIf .nDest < 12 Then
        For Cnt = 1 To .nCards
            Cards(Cnt) = GameField.GoalCell(.nDest - 7).Cards(GameField.GoalCell(.nDest - 7).nCards)
            GameField.GoalCell(.nDest - 7).nCards = GameField.GoalCell(.nDest - 7).nCards - 1
        Next Cnt
        DrawGoalCell (.nDest - 7), GameField.BB.hdc
    ElseIf .nDest < 14 Then
        For Cnt = 1 To .nCards
            Cards(Cnt) = GameField.Deck(2).Cards(GameField.Deck(2).nCards)
            GameField.Deck(2).nCards = GameField.Deck(2).nCards - 1
        Next Cnt
        DrawDeck GameField.BB.hdc
    End If
    
    If .nSrc < 8 Then
        For Cnt = 1 To .nCards
            GameField.PlaceHolder(.nSrc).Cards(GameField.PlaceHolder(.nSrc).nCards + 1) = Cards(Cnt)
            GameField.PlaceHolder(.nSrc).nCards = GameField.PlaceHolder(.nSrc).nCards + 1
        Next Cnt
        DrawPlaceHolder .nSrc, GameField.BB.hdc
    ElseIf .nSrc < 12 Then
        For Cnt = 1 To .nCards
            GameField.GoalCell(.nSrc - 7).Cards(GameField.GoalCell(.nSrc - 7).nCards + 1) = Cards(Cnt)
            GameField.GoalCell(.nSrc - 7).nCards = GameField.GoalCell(.nSrc - 7).nCards + 1
        Next Cnt
        DrawGoalCell .nSrc - 7, GameField.BB.hdc
    ElseIf .nSrc < 14 Then
        For Cnt = 1 To .nCards
            GameField.Deck(2).Cards(GameField.Deck(2).nCards + 1) = Cards(Cnt)
            GameField.Deck(2).nCards = GameField.Deck(2).nCards + 1
        Next Cnt
        DrawDeck GameField.BB.hdc
    End If
    
End With
FlipScreen
Info.Undo.Available = False

End Sub

Public Sub Find_Res()


End Sub
