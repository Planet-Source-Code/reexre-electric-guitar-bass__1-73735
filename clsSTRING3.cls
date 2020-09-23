VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSTRING3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public N           As Long
Public Nm1         As Long

Public ThicKness   As Single
Public INVthickness As Single

Private Vy()        As Single

Private Y()        As Single



Private D          As Single

'Private Const FRETK As single = 1 / 1.059467
Private Const FRETK As Single = 1 / (2 ^ (1 / 12))


Public Fret        As Long

Public IsPlaying   As Boolean
Public ShowPickFinger   As Boolean

Private DistFromNut As Single
Public FingerFretPOS As Single


Private CNT        As Long
Private StartCNT   As Long


Public FingerPickPOS As Long

Private NoteLength As Long
Private FingerLength As Long

Public SlideValue As Single


'Const S1 = (12 + 5 + 4 + 5 + 5 + 5) / 12
'Const S2 = (12 + 5 + 4 + 5 + 5) / 12
'Const S3 = (12 + 5 + 4 + 5) / 12
'Const S4 = (12 + 5 + 4) / 12
'Const S5 = (12 + 5) / 12
'Const S6 = (12) / 12


Public Sub SetLength(L, wSTR)
    Dim K
    K = 1.42

    N = L
    Nm1 = N - 1
    ReDim Y(0 To N)
    
    ReDim Vy(0 To N)


    Select Case (wSTR)

        Case 1
            ThicKness = 0.5 / FRETK ^ (12 + 5 + 4 + 5 + 5 + 5)
        Case 2
            ThicKness = 0.5 / FRETK ^ (12 + 5 + 4 + 5 + 5)
        Case 3
            ThicKness = 0.5 / FRETK ^ (12 + 5 + 4 + 5)
        Case 4
            ThicKness = 0.5 / FRETK ^ (12 + 5 + 4)
        Case 5
            ThicKness = 0.5 / FRETK ^ (12 + 5)
        Case 6
            ThicKness = 0.5 / FRETK ^ (12)
    End Select


    INVthickness = 1 / ThicKness

    INVthickness = INVthickness * 10 / 12
    CNT = 0
    StartCNT = 0
    IsPlaying = False
    ShowPickFinger = False
    
    NoteLength = NoteSampler * 1.5 * 44100
    FingerLength = NoteSampler * 0.1 * 44100
    
End Sub


Public Sub PickString(pos, Force, wfret)
    Dim K1         As Single
    Dim K2         As Single
    Dim I          As Long

    If wfret < 0 Then Exit Sub

    StartCNT = CNT
   
    IsPlaying = True
    ShowPickFinger = True

 
 Fret = wfret

    DistFromNut = N * FRETK ^ Fret - 1

    FingerPickPOS = pos
    FingerFretPOS = DistFromNut
    
    Force = Force * 0.1

    K1 = Force / pos
    K2 = Force / (DistFromNut - pos)

    For I = 0 To pos
        Y(I) = I * K1
    Next
    For I = pos + 1 To DistFromNut
        Y(I) = (DistFromNut - I) * K2
    Next

    For I = 0 To DistFromNut
        Vy(I) = 0
    Next

    For I = DistFromNut + 1 To N
        Y(I) = 0
        Vy(I) = 0

    Next


End Sub



Public Sub StepSimulate()
    Dim Vmax       As Single
    Dim I          As Long

    Dim dX         As Single
    Dim dY         As Single

    Dim D          As Single
    Dim DD         As Single
    Dim Im1 As Long
    
    Dim DFNp1 As Long
    
    DFNp1 = DistFromNut + 1

    For I = 1 To DistFromNut      'Nm1
        Y(I) = Y(I) + Vy(I) * INVthickness
    Next
    
    
    
    dY = (Y(1) - Y(0)) * INVthickness
    Vy(1) = Vy(1) - dY
        
    For I = 2 To DistFromNut 'DFNp1  'Nm1
        Im1 = I - 1
        dY = (Y(I) - Y(Im1)) * INVthickness
        Vy(I) = Vy(I) - dY
        Vy(Im1) = Vy(Im1) + dY

    Next
    dY = (Y(DFNp1) - Y(DistFromNut)) * INVthickness
    Vy(DistFromNut) = Vy(DistFromNut) + dY
    
    
    
    'Vy(1) = Vy(1) * 0.25
    'Vy(1) = Vy(1) * 0.75
    Vy(1) = Vy(1) * 0.5
    
    
'       Stop

'If SlideValue <> 0 Then DistFromNut = DistFromNut + SlideValue
DistFromNut = DistFromNut + SlideValue
FingerFretPOS = DistFromNut

    
       
    CNT = CNT + 1
    ''NoteSampler*2*44100
    
    If CNT > StartCNT + NoteLength Then
        IsPlaying = False
        For I = 0 To N
            Y(I) = 0
            Vy(I) = 0
        Next
    End If



    If CNT > StartCNT + FingerLength Then
        ShowPickFinger = False
    End If

End Sub


Public Function GetStringYatPos(pos) As Single
    GetStringYatPos = Y(pos) * 0.025
End Function
Public Function getstringVELatPos(pos) As Single
    getstringVELatPos = Vy(pos)
End Function
Public Function GetPickUP(pos) As Single
    GetPickUP = Vy(pos)
End Function


Public Function GetFretPos(fr)
    GetFretPos = N * FRETK ^ fr
End Function


