VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDELAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CopyRight 2011 Roberto Mior
'Commercial pourpose forbidden


Option Explicit

Public IDX         As Long

Private Nsamples   As Long

Private Samples()  As Integer
Private Buff()     As Integer

Public Property Let SetValue(wP, V)
    Samples(wP) = V
End Property

Public Sub SetNSamples(NofSamples)

    Nsamples = NofSamples

    ReDim Preserve Samples(Nsamples - 1)
    ReDim Preserve Buff(Nsamples - 1)
End Sub




Public Function DoSTEP(InputSample As Integer) As Integer
    Dim ID2        As Long

    IDX = (IDX + 1) Mod Nsamples
    Samples(IDX) = InputSample

    ID2 = (IDX + 1) Mod Nsamples

    DoSTEP = Samples(ID2)



End Function

Public Function GetSample(Offset As Long) As Integer
    Dim I          As Long
    I = (IDX + Offset) Mod Nsamples
    If I < 0 Then I = I + Nsamples

    GetSample = Samples(I)
End Function

Public Function SetInput(V)
    Samples(IDX) = V
End Function


Public Function GetOutput()
    GetOutput = Samples((IDX + 1) Mod Nsamples)
End Function


Public Sub SMOOTH(P)
    Dim I          As Long
    Dim I1         As Long
    Dim I2         As Long


    Buff = Samples

    For I = 0 To Nsamples - 1
        I1 = (I + 1) Mod Nsamples
        I2 = (I - 1) Mod Nsamples
        If I2 < 0 Then I2 = I2 + Nsamples

        Samples(I) = Samples(I) * P + Samples(I1) * (1 - P) * 0.5 + Samples(I2) * (1 - P) * 0.5

    Next


End Sub

