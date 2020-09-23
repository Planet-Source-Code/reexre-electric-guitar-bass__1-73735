VERSION 5.00
Begin VB.Form frmMAIN 
   Caption         =   "Karplus-Strong   Plucked-String"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   539
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   562
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox midiLIST 
      Height          =   1230
      Left            =   6360
      TabIndex        =   49
      Top             =   3240
      Width           =   1815
   End
   Begin VB.ComboBox cmbMAINQ 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6360
      TabIndex        =   47
      Text            =   "Wave out Q"
      ToolTipText     =   "Wave Quality"
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Frame fEFFECTS 
      Caption         =   "Song Effect"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   21
      Top             =   4560
      Width           =   5535
      Begin VB.Frame fFUZZ 
         Caption         =   "FUZZ Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   42
         Top             =   720
         Visible         =   0   'False
         Width           =   5295
         Begin VB.HScrollBar hFUZZs 
            Height          =   255
            Left            =   120
            Max             =   10000
            Min             =   1
            TabIndex        =   43
            Top             =   480
            Value           =   2000
            Width           =   1815
         End
         Begin VB.Label Label17 
            Caption         =   "-"
            Height          =   255
            Left            =   2040
            TabIndex        =   45
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "Strength"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame fFlanger 
         Caption         =   "Flanger Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Visible         =   0   'False
         Width           =   5295
         Begin VB.HScrollBar hFlALen 
            Height          =   255
            Left            =   120
            Max             =   5000
            TabIndex        =   37
            Top             =   480
            Value           =   1000
            Width           =   1815
         End
         Begin VB.HScrollBar hFlaPres 
            Height          =   255
            Left            =   120
            Max             =   1000
            TabIndex        =   36
            Top             =   1080
            Value           =   500
            Width           =   1815
         End
         Begin VB.Label Label13 
            Caption         =   "1 Seconds"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   41
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Length (Speed)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "Presence %"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "50 "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   38
            Top             =   1080
            Width           =   855
         End
      End
      Begin VB.Frame fDelay 
         Caption         =   "Delay Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   5295
         Begin VB.HScrollBar hDelPres 
            Height          =   255
            Left            =   120
            Max             =   1000
            TabIndex        =   31
            Top             =   1080
            Value           =   250
            Width           =   1815
         End
         Begin VB.HScrollBar hDelSec 
            Height          =   255
            Left            =   120
            Max             =   5000
            TabIndex        =   29
            Top             =   480
            Value           =   1000
            Width           =   1815
         End
         Begin VB.Label Label9 
            Caption         =   "25"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   34
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Presence %"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Delay"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "1 Seconds"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   30
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.OptionButton oeREVERSE 
         Caption         =   "Reverse"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton oeHICUT 
         Caption         =   "HiCut"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton oeFUZZ 
         Caption         =   "FUZZ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton oeFLANGER 
         Caption         =   "Flanger"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton oeDELAY 
         Caption         =   "Delay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton oeNONE 
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame fSOUND 
      Caption         =   "Samples Parameters"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   4095
      Begin VB.HScrollBar Hdelo 
         Height          =   255
         Left            =   120
         Max             =   10000
         TabIndex        =   15
         Top             =   1080
         Value           =   5000
         Width           =   3735
      End
      Begin VB.HScrollBar hTIMBRE 
         Height          =   255
         Left            =   120
         Max             =   10000
         TabIndex        =   14
         Top             =   480
         Value           =   5000
         Width           =   3735
      End
      Begin VB.HScrollBar hDeDe 
         Height          =   255
         Left            =   120
         Max             =   10000
         TabIndex        =   13
         Top             =   2280
         Value           =   5000
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.OptionButton oMode1 
         Caption         =   "M1"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1680
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton oMODE2 
         Caption         =   "M2"
         Height          =   255
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton oMODE3 
         Caption         =   "M3"
         Height          =   255
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton oMODE4 
         Caption         =   "M4"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "% Delay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "% Low Pass"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Timbre Seed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "% D2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "%D1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   495
      End
   End
   Begin VB.HScrollBar hSPEED 
      Height          =   255
      Left            =   6360
      Max             =   30000
      Min             =   100
      TabIndex        =   7
      Top             =   6720
      Value           =   15000
      Width           =   1815
   End
   Begin VB.TextBox txtSEC 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6360
      TabIndex        =   6
      Text            =   "20"
      ToolTipText     =   "Seconds to Render"
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CheckBox chDOAll 
      Caption         =   "ALL RANGE SAMPLES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.PictureBox picR 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      FillColor       =   &H00E0E0E0&
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   367
      TabIndex        =   4
      Top             =   1080
      Width           =   5535
   End
   Begin VB.PictureBox picL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      FillColor       =   &H00E0E0E0&
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   367
      TabIndex        =   3
      Top             =   600
      Width           =   5535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Render WAV song"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   2
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LoadMIDI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CREATE SAMPLE(S)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label19 
      Caption         =   "Before render Check ALL RANGE SAMPLES and click ""Create samples"". (Do it at least 1 time)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   51
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "Midi"
      Height          =   255
      Left            =   6360
      TabIndex        =   50
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Virtual Guitar  based on Karplus Strong  Plucked String"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -112
      TabIndex        =   48
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Song Speed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   46
      Top             =   6480
      Width           =   1815
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private BUint()    As Integer
Private DELAY()    As Integer
Private LOWpass()  As Integer

Dim I              As Long


Dim MidiSong       As String
Dim MidiTrack()    As Integer


'57
Const LowN = 33
Const HiN = 69
Const MaxSampleLen = 300000 / 2


Private Sub Form_Load()
    cmbMAINQ.AddItem "11025"
    cmbMAINQ.AddItem "22050"
    cmbMAINQ.AddItem "44100"
    cmbMAINQ.ListIndex = 1



    midiLIST.AddItem "Pujol_Seguidilla.Mid"    '2
    midiLIST.AddItem "Romero_Tango_Angelita.Mid"    '3
    midiLIST.AddItem "Vasco Rossi - Splendida_giornata.mid"    '6
    midiLIST.AddItem "Deep_Purple_Highway_Star1.mid"    '2
    midiLIST.AddItem "Bach_Fugue_BWV542.mid"    '2
    midiLIST.AddItem "Espanoleta.Mid"    '1

    midiLIST.ListIndex = 0

    'wich track will be played
    ReDim MidiTrack(midiLIST.ListCount)
    MidiTrack(0) = 2
    MidiTrack(1) = 3
    MidiTrack(2) = 6
    MidiTrack(3) = 2
    MidiTrack(4) = 2
    MidiTrack(5) = 1

    INITSound Me.hwnd
    InitFREQ

    If Dir(App.Path & "\INSTR", vbDirectory) = "" Then MkDir App.Path & "\INSTR"

End Sub


Private Sub Command1_Click()
    Stop



    Dim DELAY()    As Integer
    Dim Delay2()   As Integer
    Dim LOWpass()  As Integer
    Dim fiLTER()   As Integer

    Dim MaxV       As Long

    Dim C2         As Double
    Dim C          As Double
    Dim CC         As Double

    Dim NDEL       As Double
    Dim NDELdec    As Double

    Dim BufferPtr  As Long
    Dim F          As Integer

    Dim FN         As String

    Dim pDel       As Single
    Dim pLow       As Single
    Dim pDel1      As Single
    Dim pDel2      As Single


    pLow = Hdelo / 10000
    pDel = 1 - pLow


    pDel2 = hDeDe / 10000
    pDel1 = 1 - pDel2



    If chDOAll.Value = Checked Then
        notefrom = LowN
        noteto = HiN
    Else
        notefrom = 57
        noteto = 57
    End If



    For F = notefrom To noteto

        Randomize CLng(hTIMBRE) * 100
        FN = NoteToName(F)


        MaxV = MaxSampleLen

        ReDim BUint(MaxV)
        ReDim DELAY(MaxV)
        ReDim LOWpass(MaxV)
        ReDim Delay2(MaxV)
        ReDim fiLTER(MaxV)

        For I = 1 To MaxV
            BUint(I) = 0
        Next

        For I = 1 To MaxV / 10
            BUint(I) = ((Rnd * 255 - 127) * 255) * 1
        Next

        Stop

        NDEL = Conv(FREQ(F))
        NDELdec = NDEL - Int(NDEL)
        NDEL = Int(NDEL)

        Me.Caption = "Note " & F & " - Frequence " & FREQ(F) & "   KarplusStrongDelay:" & NDEL
        DoEvents

        '*************************

        For C = 1 To MaxV
            If C > NDEL Then
                'C = C '+ NDEL
                DELAY(C) = BUint(C - NDEL)
                If oMode1 Then LOWpass(C) = DELAY(C) * 0.25 + DELAY(C - 1) * 0.25 + DELAY(C - 2) * 0.25 + DELAY(C - 3) * 0.25
                If oMODE2 Then LOWpass(C) = DELAY(C) * 0.333 + DELAY(C - 1) * 0.333 + DELAY(C - 2) * 0.333
                If oMODE3 Then LOWpass(C) = DELAY(C) * 0.5 + DELAY(C - 1) * 0.5
                If oMODE4 Then LOWpass(C) = DELAY(C) * pDel1 + DELAY(C - 1) * pDel2
                BUint(C) = DELAY(C) * pDel + LOWpass(C) * pLow
            End If
        Next C

        '************************
        GoTo SKIP
        '    For C = 1 To MaxV
        '        If C - NDEL > 0 Then
        '        BUint(C) = BUint(C) * 0.5 + fiLTER(C) * 0.5
        '            Delay(C) = BUint(C - NDEL)
        '        fiLTER(C) = Delay(C) * 0.5 + Delay(C - 1) * 0.5
        '        BUint(C) = Delay(C) * 0.5 + fiLTER(C) * 0.5
        '        End If
        '    Next
        '
SKIP:
        Stop

        CreateWaveFile FN, 44100, 1, 16
        BufferPtr = 45
        For I = 0 To MaxV
            Put #1, BufferPtr, BUint(I)
            BufferPtr = BufferPtr + 2
        Next


        CloseWaveFile


        Play FN

    Next F



End Sub

Private Sub Command2_Click()
    Dim T          As String


    FilterCtlChMsg = True
    FilterSysExMsg = True


    'T = readMidiFile("c:\roberto\midi\mozart_turca.mid")
    'T = readMidiFile("c:\roberto\midi\ReelBigFish_Sell_Out.mid")
    'T = readMidiFile("c:\roberto\midi\Rodriguez_La_Cumparsita__Kar_By_Deus.mid")    'track 8
    'T = readMidiFile("c:\roberto\midi\3SUPERTRAMP__Give_A_Little_Bit_from_the_album_Even_In_The_.mid")

    'T = readMidiFile("c:\roberto\midi\Sting__shape_of_my_heart.mid")

    'T = readMidiFile("c:\roberto\midi\Deep_Purple_Highway_Star1.Mid")


    'T = readMidiFile("c:\roberto\midi\Bach_Bach__brano_16.mid")

    'T = readMidiFile("c:\roberto\midi\Vasco Rossi - Incredibile_romantica.Mid")

    'T = readMidiFile("c:\roberto\midi\Mercyful_Fate-Come_to_the_Sabbath.mid")


    'T = readMidiFile("c:\roberto\midi\GrandFunkRailroad_Footstompin_Music.Mid")


    'T = readMidiFile("c:\roberto\midi\mozart_turca.Mid")

    'T = readMidiFile("c:\roberto\midi\Vasco Rossi - Splendida_giornata.Mid")


    'T = readMidiFile(App.Path & "\espanoleta.Mid")

    'T = readMidiFile(App.Path & "\Anonym_La_Cucaracha.Mid")


    'T = readMidiFile(App.Path & "\Anonym_Sakura_Theme_Variations.Mid") '3

    'T = readMidiFile(App.Path & "\Anonym_Six_Lute_Pieces_6.Saltarello.Mid") ' 3

    'T = readMidiFile(App.Path & "\Paganini_Allegretto_MS90.mid") '3

    'T = readMidiFile(App.Path & "\Paganini_SonatineII_Passo_doppio.Mid") '2"

    '*
    'T = readMidiFile(App.Path & "\Pujol_Seguidilla.Mid") '2
    '*
    'T = readMidiFile(App.Path & "\Romero_Tango_Angelita.Mid") '3


    T = readMidiFile(App.Path & "\" & MidiSong)    '2



    '-------------------------------------------------------
    Dim s          As String
    Dim S2()       As String
    Dim N          As Integer
    Dim v          As Integer
    Dim NE         As Integer

    Dim C          As Long
    Dim Human      As Long
    Dim kHUM

    kHUM = 5

    Open App.Path & "\OutMIDI.txt" For Input As 1
    Open App.Path & "\outmidiE.txt" For Output As 2


    '                        2
    'While InStr(1, s, "Track 2") = 0 '6 piano splendida
    While InStr(1, s, "Track " & MidiTrack(midiLIST.ListIndex)) = 0

        Input #1, s
    Wend

Nread:

    Input #1, s

    If InStr(1, s, "seq/tr") Then
        'Stop
        S2 = Split(s, "|")
        Print #2, S2(3)
    End If

    If InStr(1, s, "Note on") <> 0 Then
        S2 = Split(s, "|")
        C = C + Val(S2(1))
        N = Val(S2(4))
        v = Val(S2(5))
        If v > kHUM * 2 Then v = v - kHUM * Rnd * 2
        Human = C + Rnd * kHUM / 2
        Print #2, Human & "|NOTE|" & N & "|" & v
    End If

    If InStr(1, s, "Note off") <> 0 Then
        S2 = Split(s, "|")
        C = C + Val(S2(1))
        N = Val(S2(4))
        v = 0
        Human = C + Rnd * kHUM / 2
        Print #2, Human & "|NOTE|" & N & "|" & v
    End If


    If InStr(1, s, "Pitch bend") <> 0 Then
        S2 = Split(s, "|")
        C = C + Val(S2(1))

        'Print #2, C & "|PITCH|" & HexToDec(S2(4)) & "|" & HexToDec(S2(5))
        Print #2, C & "|PITCH|" & HexToDec(S2(4)) + HexToDec(S2(5)) * 128 - 8192

    End If

    'Stop

    If InStr(1, s, "end of track") = 0 Then GoTo Nread

    Close 1
    Close 2

    Me.Caption = "DONE"



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Open App.Path & "\outmidiE.txt" For Input As 2
    N = 1
    NE = 0

    'Stop
    ReDim inMIDIEvent(0)
    inMIDIEvent(0).CurrTime = 99999999

    While Not (EOF(2))

        Input #2, s

        If InStr(1, s, "NOTE") <> 0 Then
            '    Stop

            ReDim Preserve inMIDInote(N)
            S2 = Split(s, "|")
            inMIDInote(N).CurrTime = S2(0)
            inMIDInote(N).Note = S2(2)    '+ 12 '24
            inMIDInote(N).Vol = S2(3)
            While inMIDInote(N).Note > HiN
                inMIDInote(N).Note = inMIDInote(N).Note - 12
            Wend

            While inMIDInote(N).Note < LowN
                inMIDInote(N).Note = inMIDInote(N).Note + 12
            Wend

            N = N + 1
        End If


        If InStr(1, s, "PITCH") <> 0 Then

            ReDim Preserve inMIDIEvent(NE)

            S2 = Split(s, "|")
            inMIDIEvent(NE).CurrTime = S2(0)


            inMIDIEvent(NE).PITCH = S2(2)    '+ S2(3) * 128
            NE = NE + 1

        End If

    Wend
    Close 2
    'Stop
    Me.Caption = "load Converted OK"



End Sub

Private Sub Command3_Click()


    Dim Anote()    As Integer

    Dim MainQ
    Dim SampleQ
    Dim ConstL     As Long
    Dim Ind        As Long
    Dim DivSpeed   As Double
    Dim SampleQonMainQ As Double

    Dim BOI        As Long
    Dim SAMP()     As Integer
    Dim ReadSamp() As Integer
    Dim SamplePos(0 To HiN) As Double

    Dim MN         As Long

    Dim buOUT()    As Integer

    Dim buOUT2()   As Integer



    picL.Cls
    picR.Cls



    Command2_Click


    SampleQ = 44100

    MainQ = Val(cmbMAINQ)


    SampleQonMainQ = SampleQ / MainQ

    For I = LowN To HiN
        'Stop

        Me.Caption = "Loading Sample Note " & I
        ReadSamp = LoadWaveAndConvert(NoteToName(I), SampleQ, 16, 1)
        ReDim Preserve SAMP(LowN To HiN, 0 To UBound(ReadSamp)) As Integer

        For J = 0 To UBound(ReadSamp)
            SAMP(I, J) = ReadSamp(J)
        Next J

    Next I




    ReDim Anote(0)
    ReDim Vol2(0)

    ''''''''''''''''''''''''''''''''''''''''''''''''
    'CreateWaveFile App.Path & "\SONG.Wav", MainQ, 2, 16
    ''''''''''''''''''''''''''''''''''''''''''''''''''

    ConstL = MainQ * Val(txtSEC)  '30 'Seconds


    For I = 0 To 119
        Vol(I) = 0
    Next



    Ind = 45                      'SuperConstant

    'DivSpeed = (45 / hSPEED) * MainQ
    DivSpeed = (40 / hSPEED) * MainQ



    BOI = -1                      'bufferOutIndex
    For I = 1 To ConstL


        ReDim Anote(0)

        If MN <= UBound(inMIDInote) Then
            If I / DivSpeed > inMIDInote(MN).CurrTime Then

                Vol(inMIDInote(MN).Note) = inMIDInote(MN).Vol

                SamplePos(inMIDInote(MN).Note) = 0

                MN = MN + 1

                PITCH = 1

            End If
        End If

        If MEV <= UBound(inMIDIEvent) Then
            If I / DivSpeed > inMIDIEvent(MEV).CurrTime Then

                PITCH = 1 + 0.5 * (inMIDIEvent(MEV).PITCH / 16384) / 1.12246193181818

                MEV = MEV + 1

            End If
        End If


        nc = 0
        For N = 0 To 119
            If Vol(N) <> 0 Then
                nc = nc + 1
                ReDim Preserve Anote(nc)
                ReDim Preserve Vol2(nc)
                Anote(nc) = N
                Vol2(nc) = Vol(N)
            End If
        Next

        s = ""

        SampleL = 0
        SampleR = 0
        For ia = 1 To UBound(Anote)

            's = s & Anote(ia) & " "

            SamplePos(Anote(ia)) = SamplePos(Anote(ia)) + PITCH * SampleQonMainQ


            ntempo = SamplePos(Anote(ia))


            If ntempo < MaxSampleLen Then

                'Stereo
                'If nTempo Mod 2 <> 0 Then nTempo = nTempo - 1
                'SampleL = SampleL + InpSound(nTempo) * 0.35 '0.25
                'SampleR = SampleR + InpSound(nTempo + 1) * 0.35 '* 0.25

                SampleL = SampleL + SAMP(Anote(ia), ntempo) * 0.25
                SampleR = SampleR + SAMP(Anote(ia), ntempo) * 0.25


            End If
            '        Stop

        Next ia

        '    Stop
        '*************
        ''WriteWave
        'Put #1, Ind, CInt(SampleL)
        'Ind = Ind + 2
        'Put #1, Ind, CInt(SampleR)
        'Ind = Ind + 2
        '********************

        If Abs(SampleL) > 32512 Then SampleL = Sgn(SampleL) * 32512
        If Abs(SampleR) > 32512 Then SampleR = Sgn(SampleR) * 32512

        BOI = BOI + 2
        ReDim Preserve buOUT(BOI)
        buOUT(BOI - 1) = SampleL
        buOUT(BOI) = SampleR


        If I / 10000 = I \ 10000 Then
            Me.Caption = "Rendering WAV ....  " & ConstL - I

            Xl = I * picL.ScaleWidth / ConstL


            Yl = picL.Height / 2 + (SampleL * picL.Height / 2) / Amplitude    '+ picL.Top
            picL.Line (Xl, Yl)-(xl0, yl0), vbGreen


            Yr = picL.Height / 2 + (SampleR * picR.Height / 2) / Amplitude    '+ picL.Top
            picR.Line (Xl, Yr)-(xl0, yr0), vbGreen


            xl0 = Xl
            yl0 = Yl
            yr0 = Yr


            DoEvents


        End If
    Next




    picL.Refresh

    picR.Refresh

    '''''''''''''''''''''''''''''''
    'CloseWaveFile
    'Play App.Path & "\song.wav"
    '''''''''''''''''''''''''''''''''
    Me.Caption = "Rendering Done."

    SaveArrayAsWave App.Path & "\SONG.wav", buOUT, MainQ, 16, 2
    If oeNONE Then Play App.Path & "\SONG.wav"


    '**********************************************************
    '**********************************************************

    If Not (oeNONE) Then
        Me.Caption = "Rendering effect...."
        DoEvents

        ReDim buOUT2(UBound(buOUT))
    End If


    If oeDELAY Then
        '******* DELAY stereo
        Dim LL     As Double
        Dim Pr1    As Double
        Dim Pr2    As Double

        LL = MainQ * hDelSec / 1000
        Pr2 = hDelPres / 1000
        Pr1 = 1 - Pr2

        For I = 0 To UBound(buOUT) - 2
            If I > LL Then
                buOUT2(I) = buOUT(I) * Pr1 + buOUT2(I - LL + 1) * Pr2
            Else
                buOUT2(I) = buOUT(I)
            End If
        Next
        '********
    End If


    If oeFLANGER Then
        '******* Flanger
        Dim Lf     As Double
        Dim Fp1    As Double
        Dim Fp2    As Double

        Lf = MainQ * hFlALen / 1000
        Fp2 = hFlaPres / 1000
        Fp1 = 1 - Fp2

        For I = 0 To UBound(buOUT) - 100

            If I > 100 Then
                buOUT2(I) = buOUT(I) * Fp1 + buOUT2(I - 100 + 100 * Sin(I / Lf)) * Fp2
            Else
                buOUT2(I) = buOUT(I)
            End If
        Next
        '********
    End If


    If oeFUZZ Then
        '******* FUZZ
        For I = 0 To UBound(buOUT)
            If Abs(CLng(buOUT(I)) * hFUZZs) < 32000 Then
                buOUT2(I) = buOUT(I) * hFUZZs
            Else
                buOUT2(I) = 32000 * Sgn(buOUT(I))
            End If

        Next
        '********
    End If

    If oeHICUT Then
        '******* Smooth 'Kinda Hi-Cut
        For I = 20 To UBound(buOUT) - 20
            For J = -20 To 20
                buOUT2(I) = buOUT2(I) + buOUT(I + J) / 40
            Next
        Next
        '********
    End If

    If oeREVERSE Then
        '******* Reverse
        For I = 0 To UBound(buOUT)
            buOUT2(I) = buOUT(-I + UBound(buOUT))
        Next
        '********
    End If

    If Not (oeNONE) Then
        SaveArrayAsWave App.Path & "\SONGEffect.wav", buOUT2, MainQ, 16, 2
        Play App.Path & "\SONGEffect.wav"
        Me.Caption = "Effect Done"
    End If


    '**********************************************************
    '**********************************************************







End Sub




Public Function NoteToName(F) As String

    NoteToName = App.Path & "\INSTR\N" & Format(F, "00") & ".wav"

End Function

Private Sub hDelPres_Change()
    Label9 = hDelPres / 10

End Sub

Private Sub hDelSec_Change()
    Label6 = hDelSec / 1000 & " Seconds"

End Sub

Private Sub hFlALen_Change()
    Label13 = hFlALen / 1000 & " Seconds"

End Sub

Private Sub hFlaPres_Change()
    Label10 = hFlaPres / 10

End Sub

Private Sub midiLIST_Click()
    MidiSong = midiLIST

End Sub

Private Sub oeDELAY_Click()
    fDelay.Visible = oeDELAY
    fFlanger.Visible = oeFLANGER
    fFUZZ.Visible = oeFUZZ
End Sub

Private Sub oeFLANGER_Click()
    fDelay.Visible = oeDELAY
    fFlanger.Visible = oeFLANGER
    fFUZZ.Visible = oeFUZZ
End Sub

Private Sub oeFUZZ_Click()
    fDelay.Visible = oeDELAY
    fFlanger.Visible = oeFLANGER
    fFUZZ.Visible = oeFUZZ

End Sub

Private Sub oeHICUT_Click()
    fDelay.Visible = oeDELAY
    fFlanger.Visible = oeFLANGER
    fFUZZ.Visible = oeFUZZ
End Sub

Private Sub oeNONE_Click()
    fDelay.Visible = oeDELAY
    fFlanger.Visible = oeFLANGER

    fFUZZ.Visible = oeFUZZ
End Sub

Private Sub oeREVERSE_Click()
    fDelay.Visible = oeDELAY
    fFlanger.Visible = oeFLANGER
    fFUZZ.Visible = oeFUZZ
End Sub

Private Sub oMode1_Click()
    hDeDe.Visible = oMODE4
End Sub

Private Sub oMODE2_Click()
    hDeDe.Visible = oMODE4
End Sub

Private Sub oMODE3_Click()
    hDeDe.Visible = oMODE4
End Sub

Private Sub oMODE4_Click()
    hDeDe.Visible = oMODE4


End Sub
