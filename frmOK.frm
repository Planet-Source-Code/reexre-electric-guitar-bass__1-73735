VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frmOK 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080C0FF&
   Caption         =   "Form1"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   699
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chSaveFrames 
      Caption         =   "SaveFrames at 25FPS"
      Height          =   375
      Left            =   12240
      TabIndex        =   24
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CheckBox chRenderSelection 
      BackColor       =   &H0080C0FF&
      Caption         =   "Render Only Selection"
      Height          =   255
      Left            =   10080
      TabIndex        =   23
      Top             =   5280
      Width           =   2055
   End
   Begin VB.VScrollBar sNoteSampler 
      Height          =   495
      LargeChange     =   2
      Left            =   10080
      Max             =   4
      Min             =   2
      SmallChange     =   2
      TabIndex        =   20
      Top             =   4680
      Value           =   2
      Width           =   255
   End
   Begin VB.HScrollBar sDrawGuitFreq 
      Height          =   255
      Left            =   9840
      Max             =   5000
      Min             =   1
      TabIndex        =   19
      Top             =   2160
      Value           =   5000
      Width           =   3735
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   375
      Left            =   7560
      TabIndex        =   18
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   6000
      TabIndex        =   17
      Top             =   6720
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   9960
      TabIndex        =   16
      Top             =   6000
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Max             =   1000
   End
   Begin VB.CommandButton cmdDeleteSong 
      Caption         =   "Delete this song"
      Height          =   495
      Left            =   13560
      TabIndex        =   15
      ToolTipText     =   "Clean Song Editor"
      Top             =   6600
      Width           =   1215
   End
   Begin VB.PictureBox PicSongTOP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   975
      TabIndex        =   14
      ToolTipText     =   "Left Click-Selection Start     Right click-Selection End"
      Top             =   7200
      Width           =   14655
   End
   Begin VB.HScrollBar sSONGPOS 
      Height          =   375
      Left            =   120
      Max             =   40
      TabIndex        =   13
      Top             =   9900
      Width           =   14655
   End
   Begin VB.TextBox txtMPC 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   10
      Text            =   "txtMPE"
      ToolTipText     =   "Song Speed (in Millisecond X Cell)"
      Top             =   5880
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   7560
      Pattern         =   "*.txt"
      TabIndex        =   9
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdSaveSong 
      Caption         =   "Save Song"
      Height          =   495
      Left            =   9240
      TabIndex        =   8
      Top             =   6600
      Width           =   1575
   End
   Begin VB.PictureBox PicSong 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   975
      TabIndex        =   6
      Top             =   7560
      Width           =   14655
      Begin VB.TextBox FretIN 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   7
         Text            =   "FretIN"
         ToolTipText     =   "Type FRET"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.PictureBox PicG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2310
      Left            =   120
      ScaleHeight     =   152
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   5
      ToolTipText     =   "GUITAR [Button1 & 2 to move pickUps]"
      Top             =   120
      Width           =   9630
   End
   Begin VB.HScrollBar sFRET 
      Height          =   255
      Left            =   11880
      Max             =   12
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtCorda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   2
      Text            =   "1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Commandstring 
      Caption         =   "Render Song"
      Height          =   1215
      Left            =   9960
      TabIndex        =   1
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton buttonTest 
      Height          =   855
      Left            =   9840
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SOUND"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   10080
      TabIndex        =   22
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lSAMPLER 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   21
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Load SONG"
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label 
      BackColor       =   &H0080C0FF&
      Caption         =   $"frmOK.frx":0000
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   6480
      Width           =   4215
   End
   Begin VB.Label lFRET 
      Height          =   375
      Left            =   11880
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Klen As Single


Private pPOS1 As Long
Private pPOS2 As Long

Private Sub cmdCopy_Click()
    SongCopy

End Sub

Private Sub cmdDeleteSong_Click()
    DeleteSong


End Sub

Private Sub cmdPaste_Click()


    SongPaste
    DrawSONG PicSong, sSONGPOS.Value

End Sub

Private Sub cmdSaveSong_Click()
    Dim S          As String
    S = InputBox("Song Name", "Save song", SongName)
    If Len(S) < 4 Then MsgBox "wrong file name": Exit Sub

    If LCase(Right$(S, 4)) <> ".txt" Then S = S & ".txt"

    SaveSong File1.Path & "\" & S

    SongName = S

    File1.Refresh

End Sub

Private Sub Commandstring_Click()
    Dim I          As Long
    Dim CT         As Long
    Dim fr         As Long

    Dim K          As Long
    Dim Y          As Single
    Dim NOTA       As Long

    Dim FN         As String

    Dim BACKW()    As Integer


    Dim Corda      As Long

    Dim V
    Dim Vmax       As Single
    Dim CurEV      As Long

    Dim CS         As Long
    Dim OO         As Long

    Dim SS         As Long

    Dim EVstart    As Long
    Dim EVend      As Long

    Dim StopSlide(1 To 6) As Long
    For I = 1 To 6
        StopSlide(I) = -1
    Next



    Busy Me, True

If Dir(App.Path & "\frame\*.jpg") <> "" Then Kill App.Path & "\frame\*.jpg"


    Corda = txtCorda

    SONGSPEED = 44.1 * Val(frmOK.txtMPC)


    If chRenderSelection Then
        EVstart = SelStart
        EVend = SelEnd

        ReDim WF(SONGSPEED * (EVend - EVstart + 2))
        ReDim BACKW(UBound(WF))

    Else
        EVstart = 0
        ReDim WF(SONGSPEED * (UBound(GuitEventsAT)))
        ReDim BACKW(UBound(WF))

    End If



    GUIT.INIT 440 * Klen, pPOS1 * Klen, pPOS2 * Klen, PicG    '329

    GUIT.Draw

    SongPrePlay


    FN = File1.Path & "\" & SongName & ".wav"

    CT = 0
    I = 0
    OO = 0
    CurEV = EVstart               '0
    Do


If chSaveFrames Then
        If CT Mod Fix(NoteSampler * 44100 / 25) = 0 Then '25Fps
            GUIT.Draw
            SaveJPG PicG.Image, App.Path & "\Frame\" & Format(fr, "00000") & ".jpg", 98
            fr = fr + 1
        End If
End If

        CT = CT + 1
        GUIT.Simulate


        If CT Mod NoteSampler = 0 Then
            WF(I) = Fix(16000 * (GUIT.GetPickUP1 + GUIT.GetPickUP2))
            '        Stop


            If I + SONGSPEED * EVstart = SONGSPEED * (CurEV) Then
                DrawSONG PicSong, OO + EVstart: OO = OO + 1: PicSong.Refresh: PicSongTOP.Refresh

                For CS = 1 To 6
                    If GuitEventsAT(CurEV).Fret(CS) >= 0 Then


                        If StopSlide(CS) <> CurEV Then GUIT.PlayString CS, Klen * (61 - 10 * Cos(0.0625 * CurEV * Pi2)), Klen * GuitEventsAT(CurEV).F(CS), GuitEventsAT(CurEV).Fret(CS)
                        'GUIT.PlayString CS, 50 + Rnd * 10, 100, GuitEventsAT(CurEV).Fret(CS)

                        If GuitEventsAT(CurEV).Slide(CS).StartPos <> 0 Then

                            StopSlide(CS) = GuitEventsAT(CurEV).Slide(CS).EndPos


                            '                             GUIT.SlideString CS, 0.02 * SONGSPEED / 44100 * (GuitEventsAT(CurEV).Slide(CS).EndPos - GuitEventsAT(CurEV).Slide(CS).StartPos) / _
                                                          (GUIT.GetFretPos(GuitEventsAT(CurEV).Slide(CS).EndFret) - GUIT.GetFretPos(GuitEventsAT(CurEV).Slide(CS).StartFret))


                            GUIT.SlideString CS, (1 / NoteSampler) * 1 / SONGSPEED * (GUIT.GetFretPos(GuitEventsAT(CurEV).Slide(CS).EndFret) - GUIT.GetFretPos(GuitEventsAT(CurEV).Slide(CS).StartFret)) / _
                                                 (GuitEventsAT(CurEV).Slide(CS).EndPos - GuitEventsAT(CurEV).Slide(CS).StartPos)



                        End If
                        If StopSlide(CS) = CurEV Then

                            GUIT.SlideString CS, 0

                        End If


                    End If

                Next

                CurEV = CurEV + 1
            End If
            I = I + 1
        End If

        If I Mod sDrawGuitFreq = 0 Then
            'Me.Cls
            'For K = 0 To GUIT.GetLength
            '    Y = 320 + GUIT.getALLStringsAT(K) * 300
            '    Me.Line (K, Y)-((K), Y + 7), vbRed
            'Next
            'For K = 0 To GUIT.GetLength
            'For SS = 1 To 6
            '    Y = 400 - SS * 25 + GUIT.getStringsAT(SS, K) * 50
            '    Me.Line (K, Y)-((K), Y + 2), vbRed
            'Next
            'Next
            GUIT.Draw




            Me.Caption = Int(1000 * (I / UBound(WF))) * 0.1 & "%"    ' UBound(WF) - I
            PB.Value = 1000 * (I / UBound(WF))
            DoEvents
        End If

    Loop While I <> UBound(WF)


    'Cut Hi freq
    Me.Caption = "Cut Hi freq"
    For K = 1 To 25
        BACKW = WF
        For I = 1 To UBound(WF) - 1
            WF(I) = (BACKW(I - 1) + BACKW(I) + BACKW(I + 1)) * 0.33333333
        Next
    Next


    Me.Caption = "saving..."

    CreateWaveFile WF(), FN
    Me.Caption = "play"
    Play FN

    GoTo skip

    FOUR.NumberOfSamples = 2 ^ 15    ' UBound(WF)
    For I = 1 To UBound(WF)
        FOUR.RealIn(I) = WF(I)
        FOUR.ImagIn(I) = 0
    Next

    For I = 1 To UBound(WF)
        V = FOUR.ComplexOut(I)
        If V > Vmax Then Vmax = V
    Next

    For I = 1 To UBound(WF)
        frmOK.PSet (I / 50, 400 + 200 * (FOUR.ComplexOut(I) / Vmax)), vbBlue
    Next

    Me.Caption = Vmax / 44100

skip:

    Busy Me, False
    Me.Caption = "Ready"
    DrawSONG PicSong, sSONGPOS.Value

    'DetectSoundNote

End Sub

Private Sub File1_Click()
    SongName = File1.filename
    LoadSong File1.Path & "\" & File1.filename
    DrawSONG PicSong, sSONGPOS.Value
End Sub

Private Sub Form_Load()
    pPOS1 = 40
    pPOS2 = 80
    
    File1.Path = App.Path & "\Songs"

    Randomize Timer

    INITSound Me.hwnd
    InitFREQ
    SongName = "song.txt"
    LoadSong App.Path & "\songs\song.txt"
    DrawSONG PicSong, sSONGPOS.Value

    Klen = 0.5
    GUIT.INIT 440 * Klen, pPOS1 * Klen, pPOS2 * Klen, PicG    '329
    GUIT.Draw


If Dir(App.Path & "\Frame", vbDirectory) = "" Then MkDir App.Path & "\Frame"


End Sub
Private Sub Form_Unload(Cancel As Integer)
    StopSound
    Cleanup

End Sub

Private Sub buttonTest_Click()
    Dim I          As Long
    Dim K          As Long

    ReDim WF(100000)

    DELAY.SetNSamples 44100 * 2
    Me.Caption = "Load wave"



    WF = LoadWaveAndConvert(App.Path & "\INPUTwave.wav", 44100, 16, 1)

    ReDim Preserve WF(UBound(WF) + 44100 * 5)

    '    For I = 0 To UBound(WF)
    '
    '        'WF(I) = Cos(I * (220 - I / 100000)) * 30000
    '         WF(I) = Cos(I * 220) * 30000
    '    Next
    Me.Caption = "Effect"
    For I = 1 To UBound(WF)
        WF(I) = WF(I) * 0.6 + DELAY.DoSTEP(WF(I)) * 0.1 + _
                DELAY.GetSample(44100 * 2 * 0.25) * 0.1 + _
                DELAY.GetSample(44100 * 2 * 0.5) * 0.1 + _
                DELAY.GetSample(44100 * 2 * 0.75) * 0.1

    Next


    Me.Caption = "save"
    CreateWaveFile WF(), App.Path & "\Test.wav"

    Me.Caption = "play"
    Play App.Path & "\Test.wav"

End Sub

Private Sub FretIN_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        If FretIN = "" Then FretIN = GuitEventsAT(SongCurPOS).Fret(SongCurSTR)
        GuitEventsAT(SongCurPOS).Fret(SongCurSTR) = Val(FretIN)
        GuitEventsAT(SongCurPOS).F(SongCurSTR) = 100
        FretIN.Visible = False
        DrawSONG PicSong, sSONGPOS.Value
    End If

End Sub





Private Sub PicG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim XX         As Long

    If Button = 1 Then
        XX = 440 * ((GUIT.ImageW - X) / GUIT.ImageW)
        If XX > 440 Then XX = 440
        If XX < 0 Then XX = 0
        GUIT.PickPos1 = XX * Klen
        pPOS1 = XX
        GUIT.Draw
    End If
    If Button = 2 Then
        XX = 440 * ((GUIT.ImageW - X) / GUIT.ImageW)
        If XX > 440 Then XX = 440
        If XX < 0 Then XX = 0
        GUIT.PickPos2 = XX * Klen
        pPOS2 = XX
        GUIT.Draw
    End If

End Sub

Private Sub PicG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim XX As Long

        If Button = 1 Then
        XX = 440 * ((GUIT.ImageW - X) / GUIT.ImageW)
        If XX > 440 Then XX = 440
        If XX < 0 Then XX = 0
        GUIT.PickPos1 = XX * Klen
        pPOS1 = XX
        GUIT.Draw
    End If
    If Button = 2 Then
        XX = 440 * ((GUIT.ImageW - X) / GUIT.ImageW)
        If XX > 440 Then XX = 440
        If XX < 0 Then XX = 0
        GUIT.PickPos2 = XX * Klen
        pPOS2 = XX
        GUIT.Draw
    End If


End Sub

Private Sub PicSong_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim UB         As Long
    Dim I          As Long


    If Button = 1 And FretIN.Visible = True Then Call FretIN_KeyPress(13)    ': Exit Sub


    FretIN.Left = songKX * (X \ songKX) + 10
    FretIN.Top = songKY * (Y \ songKY) + 10

    SongCurPOS = (X - 10) / songKX + sSONGPOS.Value

    If UBound(GuitEventsAT) < SongCurPOS Then
        UB = UBound(GuitEventsAT) + 1

        RedimSong (SongCurPOS + 1)
        'ReDim Preserve GuitEventsAT(0 To SongCurPOS)
        'For I = UB To SongCurPOS
        '    GuitEventsAT(I).Fret(1) = -1
        '    GuitEventsAT(I).Fret(2) = -1
        '    GuitEventsAT(I).Fret(3) = -1
        '    GuitEventsAT(I).Fret(4) = -1
        '    GuitEventsAT(I).Fret(5) = -1
        '    GuitEventsAT(I).Fret(6) = -1
        'Next

        sSONGPOS.Max = UBound(GuitEventsAT) - 24
        If sSONGPOS.Max < 0 Then sSONGPOS.Max = 1
    End If

    SongCurSTR = 6 - (Y - songKY * 0.5) / songKY
    If Button = 1 Then

        FretIN = ""
        FretIN.Visible = True
        FretIN.SetFocus

    End If

    If Button = 2 Then

        GuitEventsAT(SongCurPOS).Fret(SongCurSTR) = -1
        FretIN.Visible = False
        DrawSONG PicSong, sSONGPOS.Value
    End If


End Sub

Private Sub PicSongTOP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        SelStart = (X - 10) / songKX + sSONGPOS.Value
        If SelStart > SelEnd Then SelEnd = SelStart
    End If
    If Button = 2 Then
        SelEnd = (X - 10) / songKX + sSONGPOS.Value
        If SelEnd < SelStart Then SelStart = SelEnd
        
        RedimSong SelEnd + 1, False
    End If




    DrawSONG PicSong, sSONGPOS.Value
End Sub

Private Sub sFRET_Change()
    lFRET = "FRET " & sFRET
End Sub

Private Sub sFRET_Scroll()
    lFRET = "FRET " & sFRET
End Sub

Private Sub sNoteSampler_Change()
NoteSampler = sNoteSampler
If NoteSampler = 4 Then lSAMPLER = "GUITAR" Else: lSAMPLER = "BASS"
End Sub

Private Sub sNoteSampler_Scroll()
NoteSampler = sNoteSampler
If NoteSampler = 4 Then lSAMPLER = "GUITAR" Else: lSAMPLER = "BASS"

End Sub

Private Sub sSONGPOS_Change()


    DrawSONG PicSong, sSONGPOS.Value

End Sub

Private Sub sSONGPOS_Scroll()
    DrawSONG PicSong, sSONGPOS.Value
End Sub
