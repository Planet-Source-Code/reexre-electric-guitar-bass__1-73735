VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSTRING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Length     As Long

Private S()        As Single
Private delR       As New clsDELAYs
Private delL       As New clsDELAYs

Public Tension     As Single
Public LinearMassDensity As Single
Private K          As Single
Private T          As Long


Public Property Let SetLength(V As Long)
    Length = V
    ReDim S(Length)
    delR.SetNSamples (Length)
    delL.SetNSamples (Length)

End Property

Public Property Get GetLength() As Long
    GetLength = Length
End Property

Public Sub StepSimulate()
    Dim dLout As Single
    Dim dRout As Single
    Dim I          As Long

    C = Sqr(Tension / LinearMassDensity)


    dRout = -delR.GetOutput
    dLout = -delL.GetOutput

    delR.DoSTEP Sgn(dLout) * (Abs(dLout) ^ 0.999)
    delL.DoSTEP Sgn(dRout) * (Abs(dRout) ^ 0.9993)
    
    delR.SMOOTH 0.995
    delL.SMOOTH 0.995


    For I = 0 To Length - 1
        S(I) = delR.GetSample(I) + delL.GetSample(Length - I - 1)
        'S(I) = delR.GetSample(I) + delL.GetSample(I)

        'S(I) = delR.GetSample(T - I / C) + delL.GetSample(T - (Length - I) / C)
    Next
    T = T + 1


End Sub

Public Property Let SetStringAtPoint(wP, Value)
    Dim I          As Long

    I = Nsamples - wP
    delR.SetValue(wP) = Value
    delL.SetValue(I) = Value

End Property

Public Property Get GetStringAtPoint(wP) As Single
    GetStringAtPoint = S(wP)
End Property

Public Sub PickString(pt, Amp)
    Dim I          As Long
    Dim J          As Long
    Dim JJ         As Long
    Dim K1         As Single
    Dim K2         As Single

    Dim AW1        As Single
    AW1 = Amp / (Length - pt)
    AW2 = Amp / pt




    K1 = Amp
    For I = pt To Length - 2
        J = I Mod Length
        JJ = Length - J - 1
        delR.SetValue(J) = K1
        delL.SetValue(JJ) = K1: K1 = K1 - AW1
    Next

    K2 = Amp
    For I = pt - 1 To 1 Step -1
        J = I Mod Length
        If J < 0 Then J = J + Length
        JJ = Length - J - 1
        delR.SetValue(J) = K2
        delL.SetValue(JJ) = K2: K2 = K2 - AW2
    Next





End Sub
