Attribute VB_Name = "wConvert"
Type tmidNO
    CurrTime       As Long
    Note           As Long
    Vol            As Long
    Time           As Single
End Type

Type tMidiEvent
    CurrTime       As Long
    PITCH          As Single
End Type


Public freq()      As Double
Public FreqToNote() As String

Public Vol(119)    As Single

Public PITCH       As Single



Public inMIDInote() As tmidNO
Public inMIDIEvent() As tMidiEvent


Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
                              (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)


Public Function ConvertWave(buffer As Variant, FromFormat As WAVEFORMATEX, ToFormat As WAVEFORMATEX) As Variant
    Dim Buffer16() As Integer
    Dim Buffer8()  As Byte

    Dim RetBuff16() As Integer
    Dim RetBuff8() As Byte

    If FromFormat.nBitsPerSample = 16 Then
        Select Case VarType(buffer)
            Case (vbArray Or vbByte)
                Buffer8 = buffer
                Buffer16 = Convert16_8To16(Buffer8)
                Erase Buffer8
            Case (vbArray Or vbInteger)
                Buffer16 = buffer
            Case Else
                ConvertWave = vbEmpty
                Exit Function
        End Select

        If ToFormat.nBitsPerSample = 8 Then
            RetBuff8 = ConvertWave16to8(Buffer16)
            Erase Buffer16

            GoTo To8Bit           ' JUMP TO 8 BIT
        ElseIf ToFormat.nBitsPerSample = 16 Then
            RetBuff16 = Buffer16
            Erase Buffer16
        Else
            ConvertWave = vbEmpty
            Exit Function
        End If

To16Bit:

        If FromFormat.nChannels = 1 And ToFormat.nChannels = 2 Then
            RetBuff16 = ConvertWaveMonoToStereo16(RetBuff16)
        ElseIf FromFormat.nChannels = 2 And ToFormat.nChannels = 1 Then
            RetBuff16 = ConvertWaveStereoToMono16(RetBuff16)
        ElseIf FromFormat.nChannels <> ToFormat.nChannels Then
            ConvertWave = vbEmpty
            Exit Function
        End If

        If FromFormat.lSamplesPerSec <> ToFormat.lSamplesPerSec Then
            Select Case FromFormat.lSamplesPerSec / ToFormat.lSamplesPerSec
                Case 0.25
                    RetBuff16 = ConvertWave16MultiplySamplesBy2(RetBuff16, ToFormat.nChannels = 2)
                    RetBuff16 = ConvertWave16MultiplySamplesBy2(RetBuff16, ToFormat.nChannels = 2)
                Case 0.5
                    RetBuff16 = ConvertWave16MultiplySamplesBy2(RetBuff16, ToFormat.nChannels = 2)
                Case 2
                    RetBuff16 = ConvertWave16DivideSamplesBy2(RetBuff16, ToFormat.nChannels = 2)
                Case 4
                    RetBuff16 = ConvertWave16DivideSamplesBy2(RetBuff16, ToFormat.nChannels = 2)
                    RetBuff16 = ConvertWave16DivideSamplesBy2(RetBuff16, ToFormat.nChannels = 2)
                Case Else
                    RetBuff16 = ConvertWave16ReSample(RetBuff16, FromFormat.lSamplesPerSec, ToFormat.lSamplesPerSec, ToFormat.nChannels = 2)
            End Select
        End If

        ConvertWave = RetBuff16
    ElseIf FromFormat.nBitsPerSample = 8 Then
        Select Case VarType(buffer)
            Case (vbArray Or vbByte)
                Buffer8 = buffer
            Case (vbArray Or vbInteger)
                Buffer16 = buffer

                ReDim Buffer8((UBound(Buffer16) + 1) * 2 - 1)
                CopyMemory Buffer8(0), buffer(16), UBound(Buffer8) + 1

                Erase Buffer16
            Case Else
                ConvertWave = vbEmpty
                Exit Function
        End Select

        If ToFormat.nBitsPerSample = 16 Then
            RetBuff16 = ConvertWave8to16(Buffer8)
            Erase Buffer8

            GoTo To16Bit          ' JUMP TO 16 BIT
        ElseIf ToFormat.nBitsPerSample = 8 Then
            RetBuff8 = Buffer8
            Erase Buffer8
        Else
            ConvertWave = vbEmpty
            Exit Function
        End If
To8Bit:

        If FromFormat.nChannels = 1 And ToFormat.nChannels = 2 Then
            RetBuff8 = ConvertWaveMonoToStereo8(RetBuff8)
        ElseIf FromFormat.nChannels = 2 And ToFormat.nChannels = 1 Then
            RetBuff8 = ConvertWaveStereoToMono8(RetBuff8)
        ElseIf FromFormat.nChannels <> ToFormat.nChannels Then
            ConvertWave = vbEmpty
            Exit Function
        End If

        If FromFormat.lSamplesPerSec <> ToFormat.lSamplesPerSec Then
            Select Case FromFormat.lSamplesPerSec / ToFormat.lSamplesPerSec
                Case 0.25
                    RetBuff8 = ConvertWave8MultiplySamplesBy2(RetBuff8, ToFormat.nChannels = 2)
                    RetBuff8 = ConvertWave8MultiplySamplesBy2(RetBuff8, ToFormat.nChannels = 2)
                Case 0.5
                    RetBuff8 = ConvertWave8MultiplySamplesBy2(RetBuff8, ToFormat.nChannels = 2)
                Case 2
                    RetBuff8 = ConvertWave8DivideSamplesBy2(RetBuff8, ToFormat.nChannels = 2)
                Case 4
                    RetBuff8 = ConvertWave8DivideSamplesBy2(RetBuff8, ToFormat.nChannels = 2)
                    RetBuff8 = ConvertWave8DivideSamplesBy2(RetBuff8, ToFormat.nChannels = 2)
                Case Else
                    RetBuff8 = ConvertWave8ReSample(RetBuff8, FromFormat.lSamplesPerSec, ToFormat.lSamplesPerSec, ToFormat.nChannels = 2)
            End Select
        End If

        ConvertWave = RetBuff8
    Else
        ConvertWave = vbEmpty
        Exit Function
    End If
End Function


Public Function ConvertWave16ReSample(Buff() As Integer, ByVal FromSample As Long, ByVal ToSample As Long, ByVal Stereo As Boolean) As Integer()
    Dim K As Long, Lx As Long, Rx As Long
    Dim ret() As Integer, Per As Double, NewSize As Long

    If Not Stereo Then
        NewSize = Fix((UBound(Buff) + 1) * (ToSample / FromSample) + 0.5)
        ReDim ret(NewSize - 1)

        For K = 0 To UBound(ret) - 1
            Per = K / UBound(ret)

            Lx = Fix(UBound(Buff) * Per)

            ret(K) = FindYForX(UBound(Buff) * Per, Lx, Buff(Lx), Lx + 1, Buff(Lx + 1))
        Next K

        ret(UBound(ret)) = Buff(UBound(Buff))
    Else
        NewSize = Fix((UBound(Buff) + 1) * ToSample / FromSample + 0.5)
        NewSize = NewSize - (NewSize Mod 2)
        ReDim ret(NewSize - 1)

        For K = 0 To UBound(ret) Step 2
            Per = K / (UBound(ret) + 2)

            ' Left channel
            Lx = Fix(UBound(Buff) * Per / 2#) * 2
            ret(K + 0) = FindYForX(UBound(Buff) * Per, Lx, Buff(Lx), Lx + 2, Buff(Lx + 2))

            ' Right channel
            Rx = Lx + 1
            ret(K + 1) = FindYForX(UBound(Buff) * Per + 1, Rx, Buff(Rx), Rx + 2, Buff(Rx + 2))
        Next K

        ret(UBound(ret) - 1) = Buff(UBound(Buff) - 1)
        ret(UBound(ret)) = Buff(UBound(Buff))
    End If

    ConvertWave16ReSample = ret
End Function

Public Function ConvertWave8ReSample(Buff() As Byte, ByVal FromSample As Long, ByVal ToSample As Long, ByVal Stereo As Boolean) As Byte()
    Dim K As Long, Lx As Long, Rx As Long
    Dim ret() As Byte, Per As Double, NewSize As Long

    If Not Stereo Then
        NewSize = Fix((UBound(Buff) + 1) * ToSample / FromSample + 0.5)
        ReDim ret(NewSize - 1)

        For K = 0 To UBound(ret) - 1
            Per = K / UBound(ret)

            Lx = Fix(UBound(Buff) * Per)

            ret(K) = FindYForX(UBound(Buff) * Per, Lx, Buff(Lx), Lx + 1, Buff(Lx + 1))
        Next K

        ret(UBound(ret)) = Buff(UBound(Buff))
    Else
        NewSize = Fix((UBound(Buff) + 1) * ToSample / FromSample + 0.5)
        NewSize = NewSize - (NewSize Mod 2)
        ReDim ret(NewSize - 1)

        For K = 0 To UBound(ret) Step 2
            Per = K / (UBound(ret) + 2)

            ' Left channel
            Lx = Fix(UBound(Buff) * Per / 2#) * 2
            ret(K + 0) = FindYForX(UBound(Buff) * Per, Lx, Buff(Lx), Lx + 2, Buff(Lx + 2))

            ' Right channel
            Rx = Lx + 1
            ret(K + 1) = FindYForX(UBound(Buff) * Per + 1, Rx, Buff(Rx), Rx + 2, Buff(Rx + 2))
        Next K

        ret(UBound(ret) - 1) = Buff(UBound(Buff) - 1)
        ret(UBound(ret)) = Buff(UBound(Buff))
    End If

    ConvertWave8ReSample = ret
End Function

' convert a 16 bit wave, multiply samples by 2
Public Function ConvertWave16MultiplySamplesBy2(buffer() As Integer, ByVal Stereo As Boolean) As Integer()
    Dim K          As Long
    Dim RetBuff()  As Integer

    ReDim RetBuff(UBound(buffer) * 2)

    If Not Stereo Then
        For K = 0 To UBound(buffer) - 1
            RetBuff(K * 2) = buffer(K)
            RetBuff(K * 2 + 1) = (CLng(buffer(K)) + buffer(K + 1)) \ 2
        Next K
    Else
        For K = 0 To UBound(buffer) - 3 Step 2
            RetBuff(K * 2 + 0) = buffer(K + 0)
            RetBuff(K * 2 + 2) = (CLng(buffer(K)) + buffer(K + 2)) \ 2

            RetBuff(K * 2 + 1) = buffer(K + 1)
            RetBuff(K * 2 + 3) = (CLng(buffer(K + 1)) + buffer(K + 3)) \ 2
        Next K
    End If

    ConvertWave16MultiplySamplesBy2 = RetBuff
End Function

' convert a 8 bit wave, multiply samples by 2
Public Function ConvertWave8MultiplySamplesBy2(buffer() As Byte, ByVal Stereo As Boolean) As Byte()
    Dim K          As Long
    Dim RetBuff()  As Byte

    ReDim RetBuff(UBound(buffer) * 2)

    If Not Stereo Then
        For K = 0 To UBound(buffer) - 1
            RetBuff(K * 2) = buffer(K)
            RetBuff(K * 2 + 1) = (CLng(buffer(K)) + buffer(K + 1)) \ 2
        Next K
    Else
        For K = 0 To UBound(buffer) - 3 Step 2
            RetBuff(K * 2 + 0) = buffer(K + 0)
            RetBuff(K * 2 + 2) = (CLng(buffer(K)) + buffer(K + 2)) \ 2

            RetBuff(K * 2 + 1) = buffer(K + 1)
            RetBuff(K * 2 + 3) = (CLng(buffer(K + 1)) + buffer(K + 3)) \ 2
        Next K
    End If

    ConvertWave8MultiplySamplesBy2 = RetBuff
End Function

' convert a 16 bit wave, divide samples by 2
Public Function ConvertWave16DivideSamplesBy2(buffer() As Integer, ByVal Stereo As Boolean) As Integer()
    Dim K          As Long
    Dim RetBuff()  As Integer

    ReDim RetBuff((UBound(buffer) + 1) \ 2 - 1)

    If Not Stereo Then
        For K = 0 To UBound(buffer) Step 2
            RetBuff(K \ 2) = (CLng(buffer(K)) + buffer(K + 1)) \ 2
        Next K
    Else
        For K = 0 To UBound(buffer) - 4 Step 4
            RetBuff(K \ 2 + 0) = (CLng(buffer(K + 0)) + buffer(K + 2)) \ 2
            RetBuff(K \ 2 + 1) = (CLng(buffer(K + 1)) + buffer(K + 3)) \ 2
        Next K
    End If

    ConvertWave16DivideSamplesBy2 = RetBuff
End Function

' convert a 8 bit wave, divide samples by 2
Public Function ConvertWave8DivideSamplesBy2(buffer() As Byte, ByVal Stereo As Boolean) As Byte()
    Dim K          As Long
    Dim RetBuff()  As Byte

    ReDim RetBuff((UBound(buffer) + 1) \ 2 - 1)

    If Not Stereo Then
        For K = 0 To UBound(buffer) Step 2
            RetBuff(K \ 2) = (CLng(buffer(K)) + buffer(K + 1)) \ 2
        Next K
    Else
        For K = 0 To UBound(buffer) - 4 Step 4
            RetBuff(K \ 2 + 0) = (CLng(buffer(K + 0)) + buffer(K + 2)) \ 2
            RetBuff(K \ 2 + 1) = (CLng(buffer(K + 1)) + buffer(K + 3)) \ 2
        Next K
    End If
    ConvertWave8DivideSamplesBy2 = RetBuff
End Function


' convert a 16 bit sound from a Byte array to Integer array
Public Function Convert16_8To16(buffer() As Byte) As Integer()
    Dim Buff()     As Integer

    ReDim Buff((UBound(buffer) + 1) / 2 - 1)

    CopyMemory Buff(0), buffer(0), UBound(buffer) + 1

    Convert16_8To16 = Buff
End Function


' convert 16 bit to 8 bit
Public Function ConvertWave16to8(buffer() As Integer) As Byte()
    Dim K As Long, Val As Long
    Dim RetBuff()  As Byte

    ReDim RetBuff(UBound(buffer))

    For K = 0 To UBound(buffer)
        Val = buffer(K)
        Val = (Val + 32768) \ 256

        RetBuff(K) = Val
    Next K

    ConvertWave16to8 = RetBuff
End Function

' convert 8 bit to 16 bit
Public Function ConvertWave8to16(buffer() As Byte) As Integer()
    Dim K As Long, Val As Long
    Dim RetBuff()  As Integer

    ReDim RetBuff(UBound(buffer))

    For K = 0 To UBound(buffer)
        Val = (buffer(K) - 127) * 256

        RetBuff(K) = Val
    Next K

    ConvertWave8to16 = RetBuff
End Function


' convert mono to stereo for 16 bit buffer
Public Function ConvertWaveMonoToStereo16(buffer() As Integer) As Integer()
    Dim K          As Long
    Dim RetBuff()  As Integer

    ReDim RetBuff(UBound(buffer) * 2 + 1)

    For K = 0 To UBound(buffer)
        RetBuff(K * 2 + 0) = buffer(K)
        RetBuff(K * 2 + 1) = buffer(K)
    Next K

    ConvertWaveMonoToStereo16 = RetBuff
End Function

' convert mono to stereo for 8 bit buffer
Public Function ConvertWaveMonoToStereo8(buffer() As Byte) As Byte()
    Dim K          As Long
    Dim RetBuff()  As Byte

    ReDim RetBuff(UBound(buffer) * 2)

    For K = 0 To UBound(buffer)
        RetBuff(K * 2 + 0) = buffer(K)
        RetBuff(K * 2 + 1) = buffer(K)
    Next K

    ConvertWaveMonoToStereo8 = RetBuff
End Function



' convert stereo to mono for 16 bit buffer
Public Function ConvertWaveStereoToMono16(buffer() As Integer) As Integer()
    Dim K As Long, Val As Long
    Dim RetBuff()  As Integer


    ReDim RetBuff((UBound(buffer) + 1) \ 2 - 1)

    For K = 0 To UBound(RetBuff)
        Val = buffer(K * 2)
        Val = (Val + buffer(K * 2 + 1)) \ 2

        RetBuff(K) = Val
    Next K

    ConvertWaveStereoToMono16 = RetBuff
End Function

' convert stereo to mono for 8 bit buffer
Public Function ConvertWaveStereoToMono8(buffer() As Byte) As Byte()
    Dim K As Long, Val As Long
    Dim RetBuff()  As Byte

    ReDim Buff((UBound(buffer) + 1) \ 2 - 1)

    For K = 0 To UBound(RetBuff)
        Val = buffer(K * 2)
        Val = (Val + buffer(K * 2 + 1)) \ 2

        RetBuff(K) = Val
    Next K

    ConvertWaveStereoToMono8 = RetBuff
End Function


'Here is a sample image on how it looks like when you use line intersection formula:
'http://www.vbforums.com/attachment.php?attachmentid=46316
'
'And here is the code to convert using the line intersection formula:

Public Function FindYForX(ByVal X As Double, ByVal X1 As Double, ByVal Y1 As Double, _
                          ByVal X2 As Double, ByVal Y2 As Double) As Double

    Dim M As Double, B As Double

    M = (Y1 - Y2) / (X1 - X2)
    B = Y1 - M * X1

    FindYForX = M * X + B
End Function




Sub InitFREQ()

    ReDim freq(119)
    ReDim FreqToNote(119)

    freq(0) = 16.351              '   C
    freq(1) = 17.324              '   C# / Db
    freq(2) = 18.354              '   D
    freq(3) = 19.445              '   D# / Eb
    freq(4) = 20.601              '   E
    freq(5) = 21.827              '   F
    freq(6) = 23.124              '   F# / Gb
    freq(7) = 24.499              '   G
    freq(8) = 25.956              '   G# / Ab
    freq(9) = 27.5                '   A
    freq(10) = 29.135             '   A# / Bb
    freq(11) = 30.868             '   B
    '
    freq(12) = 32.703             '   C
    freq(13) = 34.648             '   C# / Db
    freq(14) = 36.708             '   D
    freq(15) = 38.891             '   D# / Eb
    freq(16) = 41.203             '   E
    freq(17) = 43.654             '   F
    freq(18) = 46.249             '   F# / Gb
    freq(19) = 48.999             '   G
    freq(20) = 51.913             '   G# / Ab
    freq(21) = 55                 '   A
    freq(22) = 58.27              '   A# / Bb
    freq(23) = 61.735             '   B
    '
    freq(24) = 65.406             '   C
    freq(25) = 69.296             '   C# / Db
    freq(26) = 73.416             '   D
    freq(27) = 77.782             '   D# / Eb
    freq(28) = 82.407             '   E
    freq(29) = 87.307             '   F
    freq(30) = 92.499             '   F# / Gb
    freq(31) = 97.999             '   G
    freq(32) = 103.826            '   G# / Ab
    freq(33) = 110                '   A
    freq(34) = 116.541            '   A# / Bb
    freq(35) = 123.471            '   B
    '
    freq(36) = 130.813            '   C
    freq(37) = 138.591            '   C# / Db
    freq(38) = 146.832            '   D
    freq(39) = 155.563            '   D# / Eb
    freq(40) = 164.814            '   E
    freq(41) = 174.614            '   F
    freq(42) = 184.997            '   F# / Gb
    freq(43) = 195.998            '   G
    freq(44) = 207.652            '   G# / Ab
    freq(45) = 220                '   A
    freq(46) = 233.082            '   A# / Bb
    freq(47) = 246.942            '   B
    '
    freq(48) = 261.626            '   C
    freq(49) = 277.183            '   C# / Db
    freq(50) = 293.665            '   D
    freq(51) = 311.127            '   D# / Eb
    freq(52) = 329.628            '   E
    freq(53) = 349.228            '   F
    freq(54) = 369.994            '   F# / Gb
    freq(55) = 391.995            '   G
    freq(56) = 415.305            '   G# / Ab
    freq(57) = 440                '   A
    freq(58) = 466.164            '   A# / Bb
    freq(59) = 493.883            '   B
    '
    freq(60) = 523.251            '   C
    freq(61) = 554.365            '   C# / Db
    freq(62) = 587.33             '   D
    freq(63) = 622.254            '   D# / Eb
    freq(64) = 659.255            '   E
    freq(65) = 698.456            '   F
    freq(66) = 739.989            '   F# / Gb
    freq(67) = 783.991            '   G
    freq(68) = 830.609            '   G# / Ab
    freq(69) = 880                '   A
    freq(70) = 932.328            '   A# / Bb
    freq(71) = 987.767            '   B
    '
    freq(72) = 1046.502           '   C
    freq(73) = 1108.731           '   C# / Db
    freq(74) = 1174.659           '   D
    freq(75) = 1244.508           '   D# / Eb
    freq(76) = 1318.51            '   E
    freq(77) = 1396.913           '   F
    freq(78) = 1479.978           '   F# / Gb
    freq(79) = 1567.982           '   G
    freq(80) = 1661.219           '   G# / Ab
    freq(81) = 1760               '   A
    freq(82) = 1864.655           '   A# / Bb
    freq(83) = 1975.533           '   B
    '
    freq(84) = 2093.005           '   C
    freq(85) = 2217.461           '   C# / Db
    freq(86) = 2349.318           '   D
    freq(87) = 2489.016           '   D# / Eb
    freq(88) = 2637.021           '   E
    freq(89) = 2793.826           '   F
    freq(90) = 2959.955           '   F# / Gb
    freq(91) = 3135.964           '   G
    freq(92) = 3322.438           '   G# / Ab
    freq(93) = 3520               '   A
    freq(94) = 3729.31            '   A# / Bb
    freq(95) = 3951.066           '   B
    '
    freq(96) = 4186.009           '   C
    freq(97) = 4434.922           '   C# / Db
    freq(98) = 4698.636           '   D
    freq(99) = 4978.032           '   D# / Eb
    freq(100) = 5274.042          '   E
    freq(101) = 5587.652          '   F
    freq(102) = 5919.91           '   F# / Gb
    freq(103) = 6271.928          '   G
    freq(104) = 6644.876          '   G# / Ab
    freq(105) = 7040              '   A
    freq(106) = 7458.62           '   A# / Bb
    freq(107) = 7902.132          '   B
    '
    freq(108) = 8372.018          '   C
    freq(109) = 8869.844          '   C# / Db
    freq(110) = 9397.272          '   D
    freq(111) = 9956.064          '   D# / Eb
    freq(112) = 10548.084         '   E
    freq(113) = 11175.304         '   F
    freq(114) = 11839.82          '   F# / Gb
    freq(115) = 12543.856         ''  G
    freq(116) = 13289.752         '   G# / Ab
    freq(117) = 14080             '   A
    freq(118) = 14917.24          '   A# / Bb
    freq(119) = 15804.264         '   B



    For i = 0 To 9
            FreqToNote(i * 12) = i & " C"
            FreqToNote(i * 12 + 1) = i & " C / dB"
            FreqToNote(i * 12 + 2) = i & " D"
            FreqToNote(i * 12 + 3) = i & " D# / Eb"
            FreqToNote(i * 12 + 4) = i & " E"
            FreqToNote(i * 12 + 5) = i & " F"
            FreqToNote(i * 12 + 6) = i & " F# / Gb"
            FreqToNote(i * 12 + 7) = i & " G"
            FreqToNote(i * 12 + 8) = i & " G# / Ab"
            FreqToNote(i * 12 + 9) = i & " A"
            FreqToNote(i * 12 + 10) = i & " A# / Bb"
            FreqToNote(i * 12 + 11) = i & " B"
    Next

End Sub




