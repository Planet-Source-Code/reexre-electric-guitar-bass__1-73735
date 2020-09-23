Attribute VB_Name = "modWAVE"
Option Explicit

'Create the DirectSound8 Object
Public dx          As New DirectX8
Public ds          As DirectSound8
Public dsBuffer    As DirectSoundSecondaryBuffer8

Public Const Pi = 3.141592654
Public Const Pi2 = 6.283185308

Public SampleRate  As Long        '11025

Public Sample      As Integer

Public SampleL     As Long        'Integer
Public SampleR     As Long        'Integer


Public Const Amplitude As Single = 30000    '127


Public Channels    As Integer



Public BitsPerSample As Integer


Public WFfrom      As WAVEFORMATEX
Public WFto        As WAVEFORMATEX

Public BU()        As Byte


Public InpSound()  As Integer



Public bSize       As Long
Public CAP         As DSBCAPS



Public Sub CreateWaveFile(Data() As Integer, FileName As String, Optional SRate = 44100, Optional nChan = 1, Optional BitXsample = 16)
    Dim BufferPtr  As Long
    Dim I          As Long
    Dim filesize   As Long
    Dim UB         As Long

    SampleRate = SRate
    Channels = nChan
    BitsPerSample = BitXsample

    'FileName = App.Path & "\" & FileName
    On Error Resume Next
    Kill FileName                 'REM this line if file does not exist

    Open FileName For Binary Access Write As #1
    Put #1, 1, "RIFF"             '"RIFF" header
    Put #1, 5, CInt(0)            'Filesize - 8, will write later
    Put #1, 9, "WAVEfmt "         '"WAVEfmt " header - not space after fmt
    Put #1, 17, CLng(16)          'Lenth of format data
    Put #1, 21, CInt(1)           'Wave type PCM
    Put #1, 23, CInt(Channels)    '1 channel
    Put #1, 25, CLng(SampleRate)  '44.1 kHz SampleRate
    Put #1, 29, CLng((SampleRate * BitsPerSample * Channels) / 8)
    Put #1, 33, CInt((BitsPerSample * Channels) / 8)
    Put #1, 35, CInt(BitsPerSample)
    Put #1, 37, "data"            '"data" Chunkheader
    Put #1, 41, CInt(0)           'Filesize - 44, will write later

    BufferPtr = 45
    UB = UBound(Data)
    For I = 0 To UB
        Put #1, BufferPtr, Data(I)
        BufferPtr = BufferPtr + 2
    Next



    filesize = LOF(1)
    Put #1, 5, CLng(filesize - 8)
    Put #1, 41, CLng(filesize - 44)
    Close #1

End Sub
'Get the file length, write it into the header and close the file.
Public Sub CloseWaveFile()
    Dim filesize   As Long

    filesize = LOF(1)
    Put #1, 5, CLng(filesize - 8)
    Put #1, 41, CLng(filesize - 44)
    Close #1


End Sub


'Define the DirectSound8 buffer, create it and set the play mode
Public Sub Play(FileName As String)

    Dim bufferDesc As DSBUFFERDESC

    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_STATIC Or DSBCAPS_STICKYFOCUS

    ' bufferDesc.fxFormat.nChannels = Channels
    'Stop


    Set dsBuffer = ds.CreateSoundBufferFromFile(FileName, bufferDesc)
    'dsBuffer.Play DSBPLAY_LOOPING


    dsBuffer.Play DSBPLAY_DEFAULT

    'dsBuffer.SetFrequency 800

End Sub


Public Sub INITSound(formhWnd As Long)

    On Local Error Resume Next
    Set ds = dx.DirectSoundCreate("")
    If err.Number <> 0 Then
        MsgBox "Unable to start DirectSound"
        End
    End If
    ds.SetCooperativeLevel formhWnd, DSSCL_PRIORITY
End Sub


'Dispose of the DirectSound Object and its buffer
Public Sub Cleanup()

    If Not (dsBuffer Is Nothing) Then dsBuffer.Stop
    Set dsBuffer = Nothing
    Set ds = Nothing
    Set dx = Nothing

End Sub

Public Sub StopSound()
'sw = 2
    dsBuffer.Stop
    Set dsBuffer = Nothing

End Sub


Public Function LoadWaveAndConvert(inFileName, SampleRate, Bits, Channels) As Integer()



    Dim bufferDesc As DSBUFFERDESC

    bufferDesc.lFlags = DSBCAPS_STATIC Or DSBCAPS_STICKYFOCUS


    Set dsBuffer = ds.CreateSoundBufferFromFile(inFileName, bufferDesc)

    dsBuffer.GetFormat WFfrom

    dsBuffer.GetCaps CAP

    bSize = CAP.lBufferBytes

    ReDim BU(0 To bSize)

    dsBuffer.ReadBuffer 0, bSize, BU(0), DSBLOCK_DEFAULT


    With WFfrom

        '    MsgBox "Input File" & vbCrLf & "Sample Rate " & .lSamplesPerSec & " BitsPerSample " & .nBitsPerSample & " Channels " & .nChannels

    End With


    With WFto

        .lSamplesPerSec = SampleRate    'IIf(WFfrom.lSamplesPerSec > 12000, WFfrom.lSamplesPerSec / 2, WFfrom.lSamplesPerSec)
        .nBitsPerSample = Bits
        .nChannels = Channels

        .nBlockAlign = .nBitsPerSample / 8 * .nChannels
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
        .nFormatTag = WAVE_FORMAT_PCM

    End With

    dsBuffer.SaveToFile App.Path & "\dsBufferOUT.wav"


    'InpSound = ConvertWave(BU, WFfrom, WFto)

    LoadWaveAndConvert = ConvertWave(BU, WFfrom, WFto)

End Function


Public Sub SaveArrayAsWave(FileName As String, Ar() As Integer, SamplesPerSec, Bits, Channels)


    Dim TmpWF      As WAVEFORMATEX
    Dim dsBuffer2  As DirectSoundSecondaryBuffer8


    With TmpWF
        .lSamplesPerSec = SamplesPerSec
        .nBitsPerSample = Bits
        .nChannels = Channels
        .nFormatTag = WAVE_FORMAT_PCM
        .nBlockAlign = .nBitsPerSample / 8 * .nChannels
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign

    End With


    Dim bufferDesc As DSBUFFERDESC

    bufferDesc.lFlags = DSBCAPS_STATIC Or DSBCAPS_STICKYFOCUS

    bufferDesc.lBufferBytes = (UBound(Ar)) * Channels
    bufferDesc.fxFormat = TmpWF

    Set dsBuffer2 = ds.CreateSoundBuffer(bufferDesc)


    dsBuffer2.WriteBuffer 0, UBound(Ar) * 2, Ar(0), DSBLOCK_DEFAULT

    dsBuffer2.SaveToFile FileName

End Sub


Public Function HexToDec(S As String)
    Dim CH1        As Integer
    Dim CH2        As Integer
    Dim S1         As String
    Dim S2         As String

    S1 = Left$(S, 1)
    S2 = Right$(S, 1)

    CH1 = Asc(S1)
    CH2 = Asc(S2)

    If CH1 > 58 Then
        CH1 = CH1 - 55
    Else
        CH1 = CH1 - 48
    End If
    If CH2 > 58 Then
        CH2 = CH2 - 55
    Else
        CH2 = CH2 - 48
    End If

    HexToDec = CH2 + CH1 * 16


End Function

Public Function Conv(fr)
' pow(2,(fr-69)/12.0) * 440.0 )
    Conv = 96800 / fr



End Function
