VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGuitar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Str(1 To 6) As New clsSTRING3

Private Length     As Long

Public PickPos1    As Long
Public PickPos2    As Long

Private PIC        As PictureBox
Private PicHDC As Long

Public ImageW As Single
Public ImageH As Single


Public Sub INIT(Leng, vPickPos1, vPickPos2, Picture As PictureBox)
    PicHDC = Picture.Hdc
    Set PIC = Picture

    ImageH = PIC.Height
    ImageW = PIC.Width
    
    PickPos1 = vPickPos1
    PickPos2 = vPickPos2

    Length = Leng

    Dim I          As Long

    For I = 1 To 6
        Str(I).SetLength Length, I
    Next


End Sub
Public Property Get GetLength() As Long
    GetLength = Length
End Property

Public Function GetPickUP1() As Single
    Dim I          As Long
    Dim V          As Single

    For I = 1 To 6
        V = V + Str(I).GetPickUP(PickPos1)
    Next

    GetPickUP1 = V
End Function
Public Function GetPickUP2() As Single
    Dim I          As Long
    Dim V          As Single

    For I = 1 To 6
        V = V + Str(I).GetPickUP(PickPos2)
    Next

    GetPickUP2 = V

End Function

Public Function getALLStringsAT(pos) As Single
    Dim I          As Long
    Dim V          As Single

    For I = 1 To 6
        V = V + Str(I).GetStringYatPos(pos)
        'V = V + Str(I).getstringVELatPos(pos)
    Next

    getALLStringsAT = V

End Function
Public Function getStringsAT(wSTR, pos) As Single

        getStringsAT = Str(wSTR).GetStringYatPos(pos)


End Function


Public Sub PlayString(wSTR, Pickpos, Force, Fret)

    If wSTR < 1 Then wSTR = 1
    If wSTR > 6 Then wSTR = 6

    'If Fret > 12 Then Fret = 12

    Str(wSTR).PickString Pickpos, Force, Fret


End Sub
Public Sub SlideString(wSTR, Amount)

Str(wSTR).SlideValue = Amount


End Sub

Public Sub Simulate()
    Dim I          As Long
    For I = 1 To 6
        If Str(I).IsPlaying Then Str(I).StepSimulate
    Next
End Sub


Public Function GetFretPos(fr)
    GetFretPos = Str(1).GetFretPos(fr)
End Function

Public Function GetFingerPickPOS(wSTR)
GetFingerPickPOS = Str(wSTR).FingerPickPOS
End Function

Public Function GetFingerFretPOS(wSTR)
GetFingerFretPOS = Str(wSTR).FingerFretPOS
End Function


Public Sub Draw()
    Dim S          As Long
    Dim F          As Long

    Dim X1         As Long
    Dim Y1         As Long
    Dim X2         As Long
    Dim Y2         As Long
    Dim X          As Long

    Dim INVLength  As Single
    Dim InvW       As Single
    Dim XXX        As Long

    Dim KsD        As Single
    Dim H2         As Single
    Dim V1 As Single
    Dim V2 As Single
    Dim V As Single
    
    

    INVLength = 1 / Length
    InvW = 1 / ImageW

    KsD = ImageH * 0.1
    H2 = ImageH * 0.5

    BitBlt PicHDC, 0, 0, ImageW, ImageH, PicHDC, 0, 0, vbWhite


    'Draw Pickups
    X1 = ImageW - ImageW * (PickPos1 * 0.995) * INVLength
    Y1 = H2 - 3 * KsD
    Y2 = H2 + 3 * KsD
    X2 = X1 + 5
    FastLine PicHDC, X1, Y1, X1, Y2, 5, vbGreen
    
    V1 = GetPickUP1
    FastLine PicHDC, X1, Y2 + 5, X1, Y2 + 5 + V1 * 100, 5, vbRed
   
    
    
    
    X2 = ImageW - ImageW * (PickPos2 * 0.995) * INVLength
    Y1 = H2 - 3 * KsD
    FastLine PicHDC, X2, Y1, X2, Y2, 5, vbGreen
    V2 = GetPickUP2
    
    FastLine PicHDC, X2, Y2 + 5, X2, Y2 + 5 + V2 * 100, 5, vbRed
    FastLine PicHDC, (X1 + X2) * 0.5, Y2 + 5, (X1 + X2) * 0.5, Y2 + 5 + (V1 + V2) * 100, 5, vbRed
 
    
    'Draw strings

    For X = 0 To ImageW Step ImageW * INVLength
        XXX = Length * (ImageW - X) * InvW
        For S = 1 To 6
            Y1 = H2 + (3.5 - S) * KsD
            Y2 = Y1 + getStringsAT(S, XXX) * KsD * 5
            'PIC.Line (X, Y2)-(X, Y2 + 2), vbBlue, BF
            FastLine PicHDC, X, Y2, X, Y2 + 2, 1, vbBlack
        Next
    Next

    'Draw Frets
    Y1 = H2 - 4 * KsD
    Y2 = H2 + 4 * KsD
    For F = 0 To 12
        X1 = ImageW - ImageW * (GetFretPos(F) * 0.995) * INVLength
        X2 = X1
        '  PIC.Line (X1, Y1)-(X2, Y2), IIf(F = 0, vbRed, vbBlue), BF
        FastLine PicHDC, X1, Y1, X2, Y2 + 2, 1, IIf(F = 0, vbRed, vbBlue)
    Next

    For S = 1 To 6
        If Str(S).IsPlaying Then

            XXX = GetFingerFretPOS(S)
            X = ImageW - ImageW * XXX * INVLength
            X = X - 4
            Y1 = H2 + (3.5 - S) * KsD
            MyCircle PicHDC, X, Y1, 4, 4, vbRed
            
            If Str(S).ShowPickFinger Then
            XXX = GetFingerPickPOS(S)
            X = ImageW - ImageW * XXX * INVLength
            MyCircle PicHDC, X, Y1, 4, 4, vbBlue
            End If
        End If
    Next



    PIC.Refresh

End Sub



