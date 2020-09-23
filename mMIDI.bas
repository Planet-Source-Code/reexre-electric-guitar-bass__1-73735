Attribute VB_Name = "mMIDI"
' this module containes all midi API function and my own midi function
' non specific to the sound canvas

Option Explicit
Public FilterNoteMsg As Boolean   ' while reading a midi file
Public FilterCtlChMsg As Boolean
Public FilterSysExMsg As Boolean

' midi out
Public hMidiOUT    As Long        ' handle midi out port
Public mMPU401OUT  As Long        ' roland mpu401 out device
Public midiMessageOut As Long     ' short message status byte
Public midiData1   As Long        ' short message data byte
Public midiData2   As Long        ' short message data byte
Public CurChannel  As Integer     ' short msg channel/part sequence 0-15
Public CurDChannel As Integer     ' dump channel/part sequence 9,0-8,10-15

' midi in
Public hMidiIN     As Long        ' only used for midi thru
Public mMPU401IN   As Long        ' roland mpu401 in device

'API - many structures and functions aren't used in this progr.
'Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long
'Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
'Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As Any, ByVal iMaxLength As Long) As Long
Public Const MAXPNAMELEN = 32     '  max product name length (including NULL)
Type MMakePiano
    wType          As Long
    U              As Long
End Type
Type midi
    songptrpos     As Long
End Type
Type MIDIEVENT
    dwDeltaTime    As Long        '  Ticks since last event
    dwStreamID     As Long        '  Reserved; must be zero
    dwEvent        As Long        '  Event type and parameters
    dwParms(1)     As Long        '  Parameters if this is a long event
End Type
Type MIDIHDR
    lpData         As String
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser         As Long
    dwFlags        As Long
    lpNext         As Long
    Reserved       As Long
    dwOffset       As Long
    dwReserved(4)  As Long
End Type

Type MIDIINCAPS
    wMid           As Integer
    wPid           As Integer
    vDriverVersion As Long
    szPname        As String * MAXPNAMELEN
End Type
Type MIDIOUTCAPS
    wMid           As Integer
    wPid           As Integer
    vDriverVersion As Long
    szPname        As String * MAXPNAMELEN
    wTechnology    As Integer
    wVoices        As Integer
    wNotes         As Integer
    wChannelMask   As Integer
    dwSupport      As Long
End Type
Type MIDIPROPTEMPO
    cbStruct       As Long
    dwTempo        As Long
End Type
Type MIDIPROPTIMEDIV
    cbStruct       As Long
    dwTimeDiv      As Long
End Type
Type MIDISTRMBUFFVER
    dwVersion      As Long        '  Stream buffer format version
    dwMid          As Long        '  Manufacturer ID as defined in MMREG.H
    dwOEMVersion   As Long        '  Manufacturer version for custom ext
End Type

' MIDI API Functions for Windows 95
Declare Function midiConnect Lib "winmm.dll" (ByVal hmi As Long, ByVal hmo As Long, pReserved As Any) As Long
Declare Function midiDisconnect Lib "winmm.dll" (ByVal hmi As Long, ByVal hmo As Long, pReserved As Any) As Long
Declare Function midiInAddBuffer Lib "winmm.dll" (ByVal hMidiIN As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiInClose Lib "winmm.dll" (ByVal hMidiIN As Long) As Long
Declare Function midiInGetDevCaps Lib "winmm.dll" Alias "midiInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIINCAPS, ByVal uSize As Long) As Long
Declare Function midiInGetErrorText Lib "winmm.dll" Alias "midiInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function midiInGetID Lib "winmm.dll" (ByVal hMidiIN As Long, lpuDeviceID As Long) As Long
Declare Function midiInGetNumDevs Lib "winmm.dll" () As Long
Declare Function midiInMessage Lib "winmm.dll" (ByVal hMidiIN As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Declare Function midiInOpen Lib "winmm.dll" (lphMidiIn As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiInPrepareHeader Lib "winmm.dll" (ByVal hMidiIN As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiInReset Lib "winmm.dll" (ByVal hMidiIN As Long) As Long
Declare Function midiInStart Lib "winmm.dll" (ByVal hMidiIN As Long) As Long
Declare Function midiInStop Lib "winmm.dll" (ByVal hMidiIN As Long) As Long
Declare Function midiInUnprepareHeader Lib "winmm.dll" (ByVal hMidiIN As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiOutCacheDrumPatches Lib "winmm.dll" (ByVal hMidiOUT As Long, ByVal uPatch As Long, lpKeyArray As Long, ByVal uFlags As Long) As Long
Declare Function midiOutCachePatches Lib "winmm.dll" (ByVal hMidiOUT As Long, ByVal uBank As Long, lpPatchArray As Long, ByVal uFlags As Long) As Long
Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOUT As Long) As Long
Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Declare Function midiOutGetErrorText Lib "winmm.dll" Alias "midiOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function midiOutGetID Lib "winmm.dll" (ByVal hMidiOUT As Long, lpuDeviceID As Long) As Long
Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Declare Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Declare Function midiOutLongMsg Lib "winmm.dll" (ByVal hMidiOUT As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiOutMessage Lib "winmm.dll" (ByVal hMidiOUT As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiOutPrepareHeader Lib "winmm.dll" (ByVal hMidiOUT As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiOutReset Lib "winmm.dll" (ByVal hMidiOUT As Long) As Long
Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOUT As Long, ByVal dwMsg As Long) As Long
Declare Function midiOutUnprepareHeader Lib "winmm.dll" (ByVal hMidiOUT As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Declare Function midiStreamClose Lib "winmm.dll" (ByVal hms As Long) As Long
Declare Function midiStreamOpen Lib "winmm.dll" (phms As Long, puDeviceID As Long, ByVal cMidi As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Declare Function midiStreamOut Lib "winmm.dll" (ByVal hms As Long, pmh As MIDIHDR, ByVal cbmh As Long) As Long
Declare Function midiStreamPause Lib "winmm.dll" (ByVal hms As Long) As Long
Declare Function midiStreamPosition Lib "winmm.dll" (ByVal hms As Long, lpmmt As MMakePiano, ByVal cbmmt As Long) As Long
Declare Function midiStreamProperty Lib "winmm.dll" (ByVal hms As Long, lppropdata As Byte, ByVal dwProperty As Long) As Long
Declare Function midiStreamRestart Lib "winmm.dll" (ByVal hms As Long) As Long
Declare Function midiStreamStop Lib "winmm.dll" (ByVal hms As Long) As Long

' General error return values
Public Const MMSYSERR_BASE = 0
Public Const MMSYSERR_NOERROR = 0    '  no error
Public Const MMSYSERR_ERROR = (MMSYSERR_BASE + 1)    '  unspecified error
Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)    '  device ID out of range
Public Const MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3)    '  driver failed enable
Public Const MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4)    '  device already allocated
Public Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)    '  device handle is invalid
Public Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)    '  no device driver present
Public Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)    '  memory allocation error
Public Const MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8)    '  function isn't supported
Public Const MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9)    '  error value out of range
Public Const MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10)    '  invalid flag passed
Public Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)    '  invalid parameter passed
Public Const MMSYSERR_HANDLEBUSY = (MMSYSERR_BASE + 12)    '  handle being used
'  simultaneously on another
'  thread (eg callback)
Public Const MMSYSERR_INVALIDALIAS = (MMSYSERR_BASE + 13)    '  "Specified alias not found in WIN.INI
Public Const MMSYSERR_LASTERROR = (MMSYSERR_BASE + 13)    '  last error in range

'  flags for dwFlags field of MIDIHDR structure
Public Const MHDR_DONE = &H1      '  done bit
Public Const MHDR_PREPARED = &H2  '  set if header prepared
Public Const MHDR_INQUEUE = &H4   '  reserved for driver
Public Const MHDR_VALID = &H7     '  valid flags / ;Internal /

'  flags used with waveOutOpen(), waveInOpen(), midiInOpen(), and
'  midiOutOpen() to specify the type of the dwCallback parameter.
Public Const CALLBACK_TYPEMASK = &H70000    '  callback type mask
Public Const CALLBACK_NULL = &H0  '  no callback
Public Const CALLBACK_WINDOW = &H10000    '  dwCallback is a HWND
Public Const CALLBACK_TASK = &H20000    '  dwCallback is a HTASK
Public Const CALLBACK_FUNCTION = &H30000    '  dwCallback is a FARPROC

'  manufacturer IDs
Public Const MM_MICROSOFT = 1     '  Microsoft Corp.

'  product IDs
Public Const MM_MIDI_MAPPER = 1   '  MIDI Mapper
Public Const MM_WAVE_MAPPER = 2   '  Wave Mapper

Public Const MM_SNDBLST_MIDIOUT = 3    '  Sound Blaster MIDI output port
Public Const MM_SNDBLST_MIDIIN = 4    '  Sound Blaster MIDI input port
Public Const MM_SNDBLST_SYNTH = 5    '  Sound Blaster internal synthesizer
Public Const MM_SNDBLST_WAVEOUT = 6    '  Sound Blaster waveform output
Public Const MM_SNDBLST_WAVEIN = 7    '  Sound Blaster waveform input

Public Const MM_ADLIB = 9         '  Ad Lib-compatible synthesizer

Public Const MM_MPU401_MIDIOUT = 10    '  MPU401-compatible MIDI output port
Public Const MM_MPU401_MIDIIN = 11    '  MPU401-compatible MIDI input port

Public Const MM_PC_JOYSTICK = 12  '  Joystick adapter

Public Const MM_MIM_OPEN = &H3C1  '  MIDI input
Public Const MM_MIM_CLOSE = &H3C2
Public Const MM_MIM_DATA = &H3C3
Public Const MM_MIM_LONGDATA = &H3C4
Public Const MM_MIM_ERROR = &H3C5
Public Const MM_MIM_LONGERROR = &H3C6

Public Const MM_MOM_OPEN = &H3C7  '  MIDI output
Public Const MM_MOM_CLOSE = &H3C8
Public Const MM_MOM_DONE = &H3C9

'----------------------------------------------------------------
' MIDI status messages
Public Const NOTE_OFF = &H80
Public Const NOTE_ON = &H90
Public Const POLY_KEY_PRESS = &HA0
Public Const CONTROLLER_CHANGE = &HB0
Public Const PROGRAM_CHANGE = &HC0
Public Const CHANNEL_PRESSURE = &HD0
Public Const PITCH_BEND = &HE0

' MIDI Controller Numbers Constants
Public Const MOD_WHEEL = 1
Public Const BREATH_CONTROLLER = 2
Public Const FOOT_CONTROLLER = 4
Public Const PORTAMENTO_TIME = 5
Public Const MAIN_VOLUME = 7
Public Const BALANCE = 8
Public Const PAN = 10
Public Const EXPRESS_CONTROLLER = 11
Public Const DAMPER_PEDAL = 64
Public Const PORTAMENTO = 65
Public Const SOSTENUTO = 66
Public Const SOFT_PEDAL = 67
Public Const HOLD_2 = 69
Public Const EXTERNAL_FX_DEPTH = 91
Public Const TREMELO_DEPTH = 92
Public Const CHORUS_DEPTH = 93
Public Const DETUNE_DEPTH = 94
Public Const PHASER_DEPTH = 95
Public Const DATA_INCREMENT = 96
Public Const DATA_DECREMENT = 97

'MIDI Mapper
Public Const MIDI_MAPPER = -1

'  flags for wTechnology field of MIDIOUTCAPS structure
Public Const MOD_MIDIPORT = 1     '  output port
Public Const MOD_SYNTH = 2        '  generic internal synth
Public Const MOD_SQSYNTH = 3      '  square wave internal synth
Public Const MOD_FMSYNTH = 4      '  FM internal synth
Public Const MOD_MAPPER = 5       '  MIDI mapper

'  flags for dwSupport field of MIDIOUTCAPS
Public Const MIDICAPS_VOLUME = &H1    '  supports volume control
Public Const MIDICAPS_LRVOLUME = &H2    '  separate left-right volume control
Public Const MIDICAPS_CACHE = &H4

' not used in this progr.
Public Sub MidiIN_Port(ByVal OpenClose As String)
    Dim midiError  As Long

    If OpenClose = "open" Then
        midiError = midiInOpen(hMidiIN, 0, AddressOf MidiIN_Proc, 0, CALLBACK_FUNCTION)
        If midiError <> MMSYSERR_NOERROR Then ShowMMErr "midiIN_Open", midiError
    Else
        If hMidiIN <> 0 Then
            midiError = midiInClose(hMidiIN)
            hMidiIN = 0
            If midiError <> MMSYSERR_NOERROR Then ShowMMErr "midiIN_Close", midiError
        End If
    End If
End Sub

' not used in this progr.
Public Sub MidiIN_Proc(ByVal hmIN As Long, ByVal wMsg As Long, ByVal dwInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long)
    Dim txt        As String
    On Error Resume Next
    Select Case wMsg
        Case MM_MIM_OPEN: txt = "open"
        Case MM_MIM_CLOSE: txt = "close"
        Case MM_MIM_DATA:
            If dwParam1 < &HF0 Then
                txt = "data" & " " & Hex(dwParam1) & " " & Hex(dwParam2)
            End If
        Case MM_MIM_LONGDATA: txt = "longdata" & " " & Hex(dwParam1) & " " & Hex(dwParam2)
        Case MM_MIM_ERROR: txt = "error" & " " & Hex(dwParam1) & " " & Hex(dwParam2)
        Case MM_MIM_LONGERROR: txt = "longerror"
        Case Else: txt = "???"
    End Select
    ' Form1.Label1.Caption = txt
End Sub

'Public Sub MidiOUT_Port(ByVal OpenClose As String)
'   Dim midiError As Long
'
'   If OpenClose = "open" Then
'      If mMPU401OUT = 256 Then frmDevCap.Show 1
'      midiError = midiOutOpen(hMidiOUT, mMPU401OUT, vbNull, 0, CALLBACK_NULL)
'      If midiError <> MMSYSERR_NOERROR Then ShowMMErr "midiOUT_Open", midiError
'      Else
'      If hMidiOUT <> 0 Then
'         midiError = midiOutClose(hMidiOUT)
'         hMidiOUT = 0
'         If midiError <> MMSYSERR_NOERROR Then ShowMMErr "midiOUT_Close", midiError
'         End If
'      End If
'End Sub

' this function only works when no midiIn port opened yet
'Public Sub MidiTHRU_Port(ByVal OpenClose As String)
'   Dim midiError As Long
'
'   If OpenClose = "open" Then
'      If hMidiOUT = 0 Then MidiOUT_Port "open"
'      If hMidiOUT = 0 Then Exit Sub
'      midiError = midiInOpen(hMidiIN, mMPU401IN, 0, 0, CALLBACK_NULL)
'      If midiError <> MMSYSERR_NOERROR Then
'         ShowMMErr "midiTHRU_Open", midiError
'         Else
'         midiError = midiConnect(hMidiIN, hMidiOUT, 0)
'         If midiError <> MMSYSERR_NOERROR Then ShowMMErr "midiConnect", midiError
'         midiError = midiInStart(hMidiIN)
'         End If'
'
'      Else
'      If hMidiIN <> 0 Then
'         midiError = midiInStop(hMidiIN)
'         If hMidiOUT <> 0 Then
'            midiError = midiDisconnect(hMidiIN, hMidiOUT, 0)
'            If midiError <> MMSYSERR_NOERROR Then ShowMMErr "midiDisconnect", midiError
'            End If
'         midiError = midiInClose(hMidiIN)
'         If midiError <> MMSYSERR_NOERROR Then ShowMMErr "midiTHRU_Close", midiError
'         End If
'      End If
'End Sub

' read midi FF type properties
Function readMidiFF(ByVal ch As Long, pos As Long, EndOfTrack As Boolean) As String
    Dim I As Long, Bytes As Long
    Dim B As Byte, B2 As Byte, B3 As Byte, B4 As Byte, B5 As Byte
    Dim txt As String, txt2 As String * 13
    Get #ch, pos, B2: pos = pos + 1
    If B2 = 0 Then
        Get #ch, pos, B3: pos = pos + 1
        If B3 = 0 Then
            txt = txt & "seqnr/posfile"
        Else
            Get #ch, pos, B4: pos = pos + 1
            Get #ch, pos, B5: pos = pos + 1
            txt = txt & "seq nr       " & CStr(B5 * 256 + B4)
        End If
    ElseIf B2 >= 1 And B2 <= 7 Then
        txt2 = Choose(B2, "text", "copyright", "seq/tr. name", "instrument", "lyric", "marker", "cue point") & " - "
        txt = txt & txt2
        Bytes = readVarLen(ch, pos)
        For I = 1 To Bytes
            Get #ch, pos, B: pos = pos + 1: txt = txt & Chr(B)
        Next I
    ElseIf B2 = &H20 Then
        txt = txt & "midi chann   "
        Get #ch, pos, B3: pos = pos + 1
        Get #ch, pos, B4: pos = pos + 1
        If B3 <> 0 Then txt = txt & "???len"
        txt = txt & HexByte(B4)
    ElseIf B2 = &H21 Then
        txt = txt & "midi port    "
        Get #ch, pos, B3: pos = pos + 1
        Get #ch, pos, B4: pos = pos + 1
        If B3 <> 0 Then txt = txt & "???len"
        txt = txt & HexByte(B4)
    ElseIf B2 = &H2F Then
        txt = txt & "end of track "
        Get #ch, pos, B3: pos = pos + 1
        EndOfTrack = True
    ElseIf B2 = &H51 Then
        txt = txt & "tempo        "
        Get #ch, pos, B3: pos = pos + 1
        Bytes = B3
        If Bytes <> 3 Then txt = txt & " ???len "
        Get #ch, pos, B3: pos = pos + 1
        Get #ch, pos, B4: pos = pos + 1
        Get #ch, pos, B5: pos = pos + 1
        txt = txt & CStr(CLng(60000000 / CLng(CLng(B3) * 256 * 256 + CLng(B4) * 256 + CLng(B5)))) & " BPM"
    ElseIf B2 = &H54 Then
        txt = txt & "SMPTE Offs   "
        Get #ch, pos, B3: pos = pos + 1
        Bytes = B3
        If Bytes <> 5 Then txt = txt & " ???len"
        For I = 1 To Bytes
            Get #ch, pos, B: pos = pos + 1: txt = txt & HexByte(B)
        Next I
    ElseIf B2 = &H58 Then
        txt = txt & "time sign    "
        Get #ch, pos, B3: pos = pos + 1
        Bytes = B3
        If Bytes <> 4 Then txt = txt & " ???len "
        Get #ch, pos, B4: pos = pos + 1
        Get #ch, pos, B5: pos = pos + 1
        txt = txt & CStr(B4) & "/" & CStr(2 ^ B5) & " - "
        Get #ch, pos, B4: pos = pos + 1: txt = txt & B4 & " clocks/metr.click - "
        Get #ch, pos, B5: pos = pos + 1: txt = txt & B5 & " 32nd/quarter "
    ElseIf B2 = &H59 Then
        txt = txt & "key sign     "
        Get #ch, pos, B3: pos = pos + 1
        Bytes = B3
        If Bytes <> 2 Then txt = txt & " ???len"
        For I = 1 To Bytes
            Get #ch, pos, B: pos = pos + 1: txt = txt & HexByte(B) & " "
        Next I
    ElseIf B2 = &H7F Then
        Bytes = readVarLen(ch, pos)
        txt = txt & "propr.- len  " & CStr(Bytes)
        pos = pos + Bytes
    End If
    readMidiFF = txt
End Function

' read and display a midi file
Public Function readMidiFile(ByVal File As String) As String
    Dim ch As Long, I As Long
    Dim txt As String, reg As String, deltaT As String * 7, Stat As String
    Dim MT         As String * 4
    Dim FormatType As Integer
    Dim NumTracks As Integer, Track As Integer
    Dim Division   As Integer
    Dim NumBytes As Long, Bte As Long
    Dim strBytes   As Long
    Dim Status     As Byte
    Dim pos As Long, pPos As Long, P As Long
    Dim Lng        As Long
    Dim Intg       As Integer
    Dim B1 As Byte, B2 As Byte, B3 As Byte, B4 As Byte, B5 As Byte, B As Byte
    Dim DT         As Long
    Dim EndOfTrack As Boolean

    Open App.Path & "\OutMIDI.txt" For Output As 33

    txt = txt & UCase(GetFileTitle(File)) & vbCrLf

    Print #33, txt




    ch = FreeFile
    Open File For Binary As ch
    Get #ch, 1, MT
    If MT <> "MThd" Then txt = "Geen Midi header! ": GoTo ReadMidiFileEND
    Get #ch, 5, B1
    Get #ch, 6, B2
    Get #ch, 7, B3
    Get #ch, 8, B4
    If Not (B1 = 0 And B2 = 0 And B3 = 0 And B4 = 6) Then
        txt = txt & "Midi Header lengte is fout (moet 00 00 00 06 zijn)": GoTo ReadMidiFileEND
        Print #33, "error"
    End If

    Get #ch, 9, B1
    Get #ch, 10, B2
    FormatType = B1 * 256 + B2
    txt = txt & "Format type = " & CStr(FormatType)
    Print #33, "Format type = " & CStr(FormatType)
    Select Case FormatType
        Case 0: txt = txt & " - single track any channel" & vbCrLf
        Case 1: txt = txt & " - multi tracks sep channels" & vbCrLf
        Case 2: txt = txt & " - multi patterns-songs" & vbCrLf
        Case Else: txt = txt & " - onbekend = fout": GoTo ReadMidiFileEND
    End Select

    Get #ch, 11, B1
    Get #ch, 12, B2
    NumTracks = B1 * 256 + B2
    txt = txt & "NumTracks = " & CStr(NumTracks) & vbCrLf
    Print #33, "NumTracks = " & CStr(NumTracks)

    If FormatType = 0 And NumTracks > 1 Then txt = txt & "Aantal tracks klopt niet met het formaat type = fout.": GoTo ReadMidiFileEND

    Get #ch, 13, B1
    Get #ch, 14, B2
    Division = B1 * 256 + B2
    txt = txt & "Division = " & CStr(Division) & " PPQN" & vbCrLf
    Print #33, "Division = " & CStr(Division) & " PPQN"
    pos = 15
    For Track = 1 To NumTracks
        EndOfTrack = False

        Get #ch, pos, MT
        If MT <> "MTrk" Then txt = txt & "Geen Midi track gevonden op de verwachte plaats! ": GoTo ReadMidiFileEND
        Get #ch, , B1
        Get #ch, , B2
        Get #ch, , B3
        Get #ch, , B4
        NumBytes = CLng(CLng(B1) * 256 ^ 3 + CLng(B2) * 256 ^ 2 + CLng(B3) * 256 + CLng(B4))
        txt = txt & vbCrLf & "Track " & CStr(Track) & "     lengte = " & CStr(NumBytes) & vbCrLf

        Print #33, vbCrLf & "Track " & CStr(Track) & "     lengte = " & CStr(NumBytes)


        pos = pos + 8
        pPos = pos

        Status = 0
        While pos - pPos < NumBytes
            Get #ch, pos, B1
            If B1 = &HFF Then
                pos = pos + 1
                Status = B1
                reg = readMidiFF(ch, pos, EndOfTrack)
                Stat = " " & HexByte(Status) & " "

            Else

                DT = readVarLen(ch, pos)
                deltaT = CStr(DT)
                Get #ch, pos, B1
                If (B1 And &H80) = &H80 Then
                    Status = B1
                    Stat = " " & HexByte(Status) & " "
                    pos = pos + 1
                Else
                    Stat = "r" & HexByte(Status) & " "
                End If

                '               Stop

                Select Case Status And &HF0

                    Case &H80
                        Get #ch, pos, B2: pos = pos + 1
                        Get #ch, pos, B3: pos = pos + 1
                        If FilterNoteMsg = False Then reg = "Note off|" & isNote(B2) & "|" & CStr(B3)
                    Case &H90
                        Get #ch, pos, B2: pos = pos + 1
                        Get #ch, pos, B3: pos = pos + 1
                        If FilterNoteMsg = False Then reg = "Note on |" & isNote(B2) & "|" & CStr(B3)
                    Case &HB0
                        Get #ch, pos, B2: pos = pos + 1
                        Get #ch, pos, B3: pos = pos + 1
                        If FilterCtlChMsg = False Then reg = "Ctl Change.. " & HexByte(B2) & " " & HexByte(B3)
                    Case &HC0
                        Get #ch, pos, B2: pos = pos + 1
                        reg = "Prg Change.. " & HexByte(B2)
                    Case &HD0
                        Get #ch, pos, B2: pos = pos + 1
                        reg = "Chan Press.. " & HexByte(B2)
                    Case &HE0
                        Get #ch, pos, B2: pos = pos + 1
                        Get #ch, pos, B3: pos = pos + 1
                        reg = "Pitch bend..|" & HexByte(B2) & "|" & HexByte(B3)
                    Case &HF0
                        Select Case Status
                            Case &HFE
                            Case &HFF
                                reg = readMidiFF(ch, pos, EndOfTrack)
                            Case &HF0
                                P = pos
                                Lng = readVarLen(ch, pos)
                                If FilterSysExMsg = False Then reg = "SysEx - len: " & CStr(Lng)
                                pos = pos + Lng
                            Case &HF7
                            Case Else
                        End Select
                End Select
            End If

            ''        frmReadMid.SetProgress Pos
            DoEvents
            If Cancel = True Then GoTo ReadMidiFileEND:


            If reg <> "" Then Print #33, "|" & deltaT & "|" & Stat & "|" & reg

            If reg <> "" Then txt = txt & "|" & deltaT & "|" & Stat & "|" & reg & vbCrLf: reg = ""

            ''If Len(txt) > 3200000 Then txt = txt & vbCrLf & "Tekst te lang...." & vbCrLf: GoTo ReadMidiFileEND:
            If Len(txt) > 1000 Then txt = ""
        Wend

ReadMidiFileNEXTTRACK:
        ''   If Len(txt) > 32000 Then txt = txt & vbCrLf & "Tekst te lang...." & vbCrLf: GoTo ReadMidiFileEND:
        ''   frmReadMid.SetProgress Pos



        DoEvents
    Next Track

ReadMidiFileEND:
    Close ch
    readMidiFile = txt

    Close 33


End Function

' function used in break/design time
Function showMIDIHDR(mh As MIDIHDR) As String
    Dim txt        As String
    txt = txt & "lpData " & getComStrHex(mh.lpData) & vbCrLf
    txt = txt & "dwBufferLength " & CStr(mh.dwBufferLength) & vbCrLf
    txt = txt & "dwBytesRecorded " & CStr(mh.dwBytesRecorded) & vbCrLf
    txt = txt & "dwFlags " & CStr(mh.dwFlags) & vbCrLf
    txt = txt & "dwOffset " & CStr(mh.dwOffset) & vbCrLf
    txt = txt & "Reserved " & CStr(mh.Reserved) & vbCrLf
    txt = txt & "dwUser " & CStr(mh.dwUser) & vbCrLf
    showMIDIHDR = txt
End Function

Public Sub SysExDT1(ByVal ComStr As String)
    Dim midiError  As Long
    Dim mh         As MIDIHDR

    ' they say the length should be multiples of 4
    ' but it seems to work without following line too
    While Len(ComStr) < 16: ComStr = ComStr & Chr(0): Wend

    mh.lpData = ComStr
    mh.dwBufferLength = Len(mh.lpData)
    mh.dwFlags = 0
    'MsgBox showMIDIHDR(mh)
    midiError = midiOutPrepareHeader(hMidiOUT, mh, 24 + mh.dwBufferLength)
    If midiError <> MMSYSERR_NOERROR Then
        ShowMMErr "midiOutPrepareHeader", midiError
    Else
        'MsgBox showMIDIHDR(mh)
        midiError = midiOutLongMsg(hMidiOUT, mh, 24 + mh.dwBufferLength)
        If midiError <> MMSYSERR_NOERROR Then
            ShowMMErr "midiOutLongMsg", midiError
        Else
            ' Normaly the function would look like this:
            '
            ' While mh.dwFlags <> MHDR_DONE : DoEvents Wend
            '
            ' But something went wrong, so I had to set the
            ' flag myself, to force my way out
            ' Can somebody help to solve this problem?
            mh.dwFlags = MHDR_DONE
        End If
        midiError = midiOutUnprepareHeader(hMidiOUT, mh, 24 + mh.dwBufferLength)
        If midiError <> MMSYSERR_NOERROR Then ShowMMErr "midiOutUnprepareHeader", midiError
    End If
End Sub

Public Sub SendMidiShortOut()
    Dim midiMessage As Long
    Dim lowint As Long, highint As Long

    'Pack MIDI message data into 4 byte long integer
    lowint = (midiData1 * 256) + midiMessageOut
    highint = (midiData2 * 256) * 256
    midiMessage = lowint + highint
    'Windows MIDI API function
    midiOutShortMsg hMidiOUT, midiMessage
End Sub

Public Sub ShowMMErr(InFunct As String, MMErr)
    Dim msg        As String

    msg = String(255, " ")
    If InStr(1, InFunct, "out", vbTextCompare) = 0 Then
        midiInGetErrorText MMErr, msg, 255
    Else
        midiOutGetErrorText MMErr, msg, 255
    End If
    msg = InFunct & vbCrLf & msg & vbCrLf
    Select Case MMErr
        Case MMSYSERR_NOERROR: msg = msg & "no error"
        Case MMSYSERR_ERROR: msg = msg & "unspecified error"
        Case MMSYSERR_BADDEVICEID: msg = msg & "device ID out of range"
        Case MMSYSERR_NOTENABLED: msg = msg & "driver failed enable"
        Case MMSYSERR_ALLOCATED: msg = msg & "device already allocated"
        Case MMSYSERR_INVALHANDLE: msg = msg & "device handle is invalid"
        Case MMSYSERR_NODRIVER: msg = msg & "no device driver present"
        Case MMSYSERR_NOMEM: msg = msg & "memory allocation error"
        Case MMSYSERR_NOTSUPPORTED: msg = msg & "function isn't supported"
        Case MMSYSERR_BADERRNUM: msg = msg & "error value out of range"
        Case MMSYSERR_INVALFLAG: msg = msg & "invalid flag passed"
        Case MMSYSERR_INVALPARAM: msg = msg & "invalid parameter passed"
        Case MMSYSERR_HANDLEBUSY: msg = msg & "handle being used simultaneously on another thread (eg callback)"
        Case MMSYSERR_INVALIDALIAS: msg = msg & "Specified alias not found in WIN.INI"
        Case MMSYSERR_LASTERROR: msg = msg & "last error in range"
    End Select
    MsgBox msg
End Sub

