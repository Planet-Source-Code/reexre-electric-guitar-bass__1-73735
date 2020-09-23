VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFourier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Fast Fourier Transformation - FFT
'Fast Inverse Fourier Transformation - IFFT

'...and it is fast indeed - will transform 2048 samples in under 1.5 mSec now on an Atlon 1800 MHz CPU
'                                               (with a little help from my friends)
Option Explicit

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Const HIGH_PRIORITY_CLASS As Long = &H80
Private PrevPrioCls As Long

Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Const MCHex As String = "31 C0 8B 54 24 08 8B 4C 24 0C 67 E3 06 D1 DA 11 C0 E2 FA 8B 54 24 10 89 02 C2 10 00"
'from ASM
'       xor     eax,eax
'       mov     edx,[esp+08h]
'       mov     ecx,[esp+0Ch]
'       jcxz    GetOut
'Again:
'       rcr     edx,1
'       adc     eax,eax
'       loop    Again
'GetOut:
'       mov     edx,dword [esp+010h]
'       mov     [edx],eax
'       ret     16

Private MCBin()    As Byte
Private Const MirrorOffset As Long = 28 + 4 * 10    '28 is vtable pointer offset, 4 the entry size and then: 10 <= 8 public + 3rd private in alphabetical order - 1 because it's zero based (???)

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private AtStart    As Currency
Private AtEnd      As Currency
Private CPUSpeed   As Currency

Private UBSamples  As Long        'upper bound of samples
Private NumBits    As Long        'number of bits needed to express above
Private StageSize  As Long        'the number of samples in current computation stage
Private NumButter  As Long        'the number of butterflies in current stage
Private I          As Long        'helpers, enumerators and such
Private j          As Long
Private K          As Long
Private L          As Long
Private Pi         As Double      'what it says:  pi
Private tmp        As Double
Private UnknownSize As Boolean
Private NeedsDoing As Boolean
Private myReverse  As Boolean
Private Const Ioor As String = "Index out of range or number of samples unknown"
Private Const Nosm As String = "Number of Samples must be a positive value of the form 2 ^ n"

Private Type Sample               'sample consists of a real and an imaginary value in gaussian complex plane
    Real           As Double
    Imag           As Double
End Type

Private S          As Sample
Private T          As Sample
Private U          As Sample
Private ValuesIn() As Sample
Private ValuesOut() As Sample

Private Sub Class_Initialize()

    Dim hx()       As String
    Dim VTableAddress As Long
    Dim CodeAddress As Long

    UnknownSize = True
    ReDim ValuesIn(0)
    ReDim ValuesOut(0)
    NeedsDoing = True

    QueryPerformanceFrequency CPUSpeed

    'the following is here tnx to Paul Caton
    hx = Split(MCHex, " ")
    ReDim MCBin(0 To UBound(hx))
    For I = 0 To UBound(hx)
        MCBin(I) = Val("&H" & hx(I))
    Next I
    CodeAddress = VarPtr(MCBin(0))

    'patch Mirror Function
    MemCopy VarPtr(VTableAddress), ObjPtr(Me), 4    'get vTable address
    MemCopy VTableAddress + MirrorOffset, VarPtr(CodeAddress), 4    'patch proper entry in vTable

End Sub

Public Property Get ComplexOut(Index As Long) As Double

    With GetIt(Index)
        ComplexOut = Sqr(.Real * .Real + .Imag * .Imag)
    End With                      'GETIT(INDEX)

End Property

Private Function GetIt(Index As Long) As Sample

    If UnknownSize Or Index < 1 Or Index > UBound(ValuesIn) + 1 Then
        err.Raise 381, , Ioor
    Else                          'NOT UNKNOWNSIZE...
        If NeedsDoing Then
            NeedsDoing = False

            '=======================================================================================
            'Begin Fast Fourier Transformation

            QueryPerformanceCounter AtStart
            PrevPrioCls = GetPriorityClass(GetCurrentProcess)
            SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS

            Pi = 4 * Atn(1)
            If myReverse Then     'from fourier back to samples
                Pi = -Pi
            End If

            UBSamples = UBound(ValuesIn)
            NumBits = Log(UBSamples + 1) / Log(2)    'the number of bits needed to express UBSamples

            For I = 0 To UBSamples
                ValuesOut(Mirror(I, NumBits)) = ValuesIn(I)    'copy to mirrored fourier array
            Next I

            StageSize = 1

            Do
                'divide and conquer
                NumButter = StageSize
                StageSize = NumButter * 2

                T.Real = Pi / NumButter
                S.Real = Sin(T.Real / 2)
                S.Real = 2 * S.Real * S.Real
                S.Imag = Sin(T.Real)

                For I = 0 To UBSamples Step StageSize
                    U.Real = 1
                    U.Imag = 0
                    For j = I To I + NumButter - 1
                        'butterfly calculation
                        K = j + NumButter
                        With ValuesOut(K)
                            T.Real = U.Real * .Real - U.Imag * .Imag
                            T.Imag = U.Imag * .Real + U.Real * .Imag
                            .Real = ValuesOut(j).Real - T.Real
                            .Imag = ValuesOut(j).Imag - T.Imag
                        End With  'VALUESOUT(K)
                        With ValuesOut(j)
                            .Real = .Real + T.Real
                            .Imag = .Imag + T.Imag
                        End With  'VALUESOUT(J)
                        tmp = S.Real * U.Real + S.Imag * U.Imag
                        U.Imag = U.Imag - (S.Real * U.Imag - S.Imag * U.Real)
                        U.Real = U.Real - tmp
                    Next j, I
                Loop Until StageSize > UBSamples

                If myReverse Then    'we will have to normalize the complex values
                    tmp = UBSamples + 1
                    For I = 0 To UBSamples
                        With ValuesOut(I)
                            .Real = .Real / tmp
                            .Imag = .Imag / tmp
                        End With  'ValuesOut(I)
                    Next I
                End If

                SetPriorityClass GetCurrentProcess, PrevPrioCls
                QueryPerformanceCounter AtEnd

                'End Fast Fourier Transformation
                '=======================================================================================

            End If
            GetIt = ValuesOut(Index - 1)
        End If

    End Function

Public Property Let ImagIn(Index As Long, nuValueIn As Double)

    If UnknownSize Or Index < 1 Or Index > UBound(ValuesIn) + 1 Then
        err.Raise 381, , Ioor
    Else                          'NOT UNKNOWNSIZE...
        ValuesIn(Index - 1).Imag = nuValueIn
        NeedsDoing = True
    End If

End Property

Public Property Get ImagOut(Index As Long) As Double

    ImagOut = GetIt(Index).Imag

End Property

Private Function Mirror(ByVal Index As Long, ByVal NumBits As Long) As Long

'reverse [numbits] bits; eg bits 6 5 4 3 2 1 0 etc are shifted into posn 0 1 2 3 4 5 6

    Mirror = 0                    'dummy: will be patched

End Function

Public Property Let NumberOfSamples(nuNumSam As Long)

    If nuNumSam > 1 And (nuNumSam - 1 And nuNumSam) = 0 Then
        ReDim ValuesIn(0 To nuNumSam - 1)
        ReDim ValuesOut(0 To nuNumSam - 1)
        UnknownSize = False
        NeedsDoing = True
    Else                          'NOT NUNUMSAM...
        err.Raise 380, , Nosm
    End If

End Property

Public Property Let RealIn(Index As Long, nuValueIn As Double)

    If UnknownSize Or Index < 1 Or Index > UBound(ValuesIn) + 1 Then
        err.Raise 381, , Ioor
    Else                          'NOT UNKNOWNSIZE...
        ValuesIn(Index - 1).Real = nuValueIn
        NeedsDoing = True
    End If

End Property

Public Property Get RealOut(Index As Long) As Double

    RealOut = GetIt(Index).Real

End Property

Public Property Get Timing() As Single

    Timing = (AtEnd - AtStart) / CPUSpeed * 1000

End Property

Public Property Let TransformReverse(nuReverse As Boolean)

    myReverse = CBool(nuReverse)

End Property

':) Ulli's VB Code Formatter V2.21.6 (2006-Apr-06 22:44)  Decl: 64  Code: 188  Total: 252 Lines
':) CommentOnly: 26 (10,3%)  Commented: 23 (9,1%)  Empty: 59 (23,4%)  Max Logic Depth: 7
