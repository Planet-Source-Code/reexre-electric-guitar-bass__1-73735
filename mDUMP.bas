Attribute VB_Name = "mDUMP"
' this module containes routines mostly specific to the sound canvas

Option Explicit
' dump
Public CurDmpMidFile As String    ' current dump file
Public CurDmpMidFileTitle As String
Public dmpB(3680)  As Byte        ' dump data bytes un-nibbled
Public dmpBc       As Long        ' dump dat bytecount
Public dmpM(3680)  As Boolean     ' marked (non-default values)
Public Rx(16, 16)  As Byte        ' Rx. Switches parsed 2bytes=16bit
Public PartParName(112) As String * 20    ' pure sequential order PART param. names
Public AllParName(64) As String * 20    ' pure sequential order COMMON/ALL param. names
Public dmpDivision As Integer     ' Pulses Per Quarter Note
Public dmpTrackname As String
Public NumBytes    As Long        ' track length
Public dmpTempo    As Long        ' BMP
Public dmpTSnn     As Byte        ' Time Sign numerator
Public dmpTSdd     As Byte        '           2 ^ dmpTSdd = denominator
Public dmpTScc     As Byte        ' Time Sign clocks/metronomclick
Public dmpTSbb     As Byte        '           32nd/quarter notes

' when saving altern in MIDfile
Public SaveAlternat As Boolean    ' we are saving --> include varlen length value
Public UseShMsg    As Boolean     ' use short message where possible
Public StartWithGSReset As Boolean    ' include a sound canvas reset
Public PartSwitch(16) As Boolean  ' filter parts (or fade form = opposite)
Public AltNoCommon As Boolean     ' no common section
Public SaveDumpAltFile As String  ' name of the altern. file
Public SaveFadeFile As String

Public mManuID     As Byte        ' &H41 - Roland
Public mDeviceID   As Byte        ' default 17 (only different when using more SC)
Public mModelID    As Byte        ' GS=&H42 SC55=&H45

Type SC55PARAMTYPE
    Bytes          As String * 2
    MinHexStr      As String * 4  ' for re-assembling data file - design time
    MaxHexStr      As String * 4  ' could be left out
    MinDecStr      As String * 6
    MaxDecStr      As String * 6
    MinHex         As Long
    MaxHex         As Long
    MinDec         As Single
    MaxDec         As Single
    name           As String * 15
End Type

Type SC55PARAM
    ShortName      As String * 7
    ByteOffs       As Byte        ' offset into dump bytes mem block
    Address        As String * 6  ' param address
    Type           As Byte        ' param type ID
        Default    As Byte        ' default value
        Altern     As Byte        ' param alt ID
        name       As String * 20
    End Type

    Type SHORTMSGs                ' table with alt
        CommandStr As String * 20
        name       As String * 15
    End Type

Type REVERBMACRO
    name           As String * 15
    rPL            As Byte        ' pre-lpf
    rTM            As Byte        ' time
    rFB            As Byte        ' feedback
End Type

Type CHORUSMACRO
    name           As String * 15
    cFB            As Byte        ' feedback
    cDL            As Byte        ' delay
    cRT            As Byte        ' rate
    cDP            As Byte        ' depth
End Type

Public ChoM(7)     As CHORUSMACRO
Public RevM(7)     As REVERBMACRO
Public ShMsg(25)   As SHORTMSGs
Public ShMsgCount  As Long
Public APar(64)    As SC55PARAM
Public AParCount   As Long
Public PPar(112)   As SC55PARAM
Public PParCount   As Long
Public PType(30)   As SC55PARAMTYPE
Public PTypesCount As Long

' marks all non-default values in dump block
Public Sub CheckNonDefaults()
    Dim I As Long, j As Long      ' counters
    Dim iB         As Byte
    Dim BO         As Long        ' byte offset
    Dim Marked     As Boolean
    Dim ch         As Long        ' channel/part
    Dim rMac       As Byte        ' reverb macro
    Dim cMac       As Byte        ' chorus macro

    ' common
    I = 0                         ' Master tune
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = 4 And dmpB(BO + 1) = 0, False, True)
    I = 1                         ' Master volume
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = APar(I).Default, False, True)
    I = 2                         ' Master key shift
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = APar(I).Default, False, True)
    I = 3                         ' Master panpot
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = APar(I).Default, False, True)

    I = 5                         ' parts partial reserve
    For j = 0 To 15
        BO = j + APar(I).ByteOffs
        iB = dmpB(BO)
        dmpM(BO) = IIf(iB = Val(Mid("2622222222000000", j + 1, 1)), False, True)
    Next j

    I = 6                         ' reverb
    BO = APar(I).ByteOffs: rMac = dmpB(BO)
    dmpM(BO) = IIf(dmpB(BO) = APar(I).Default, False, True)
    I = 7                         ' rev char
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = rMac, False, True)
    I = 8                         ' rev pre-lpf
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = RevM(rMac).rPL, False, True)
    I = 9                         ' rev level
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = &H40, False, True)
    I = 10                        ' rev time
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = RevM(rMac).rTM, False, True)
    I = 11                        ' rev feedb
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = RevM(rMac).rFB, False, True)
    I = 12                        ' rev to cho
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = 0, False, True)

    I = 13                        ' chorus
    BO = APar(I).ByteOffs: cMac = dmpB(BO)
    dmpM(BO) = IIf(dmpB(BO) = APar(I).Default, False, True)
    I = 14                        ' cho pre-lpf
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = 0, False, True)
    I = 15                        ' cho level
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = &H40, False, True)
    I = 16                        ' cho feedb
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = ChoM(cMac).cFB, False, True)
    I = 17                        ' cho delay
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = ChoM(cMac).cDL, False, True)
    I = 18                        ' cho rate
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = ChoM(cMac).cRT, False, True)
    I = 19                        ' cho depth
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = ChoM(cMac).cDP, False, True)
    I = 20                        ' cho to rev
    BO = APar(I).ByteOffs: dmpM(BO) = IIf(dmpB(BO) = 0, False, True)

    ' parts parameters
    For I = 0 To PParCount - 1
        If PPar(I).ShortName = "RX1 0/1" Then Exit For
        For j = 0 To 15
            BO = 72 + j * 112 + PPar(I).ByteOffs
            iB = dmpB(BO)
            If PPar(I).ShortName = "CHANNEL" Then
                Marked = IIf(iB = Val("&H" & Mid("9012345678ABCDEF", j + 1, 1)), False, True)
            Else
                Marked = IIf(iB = PPar(I).Default, False, True)
            End If
            If j = 0 And I = 3 And iB = &HB0 Then Marked = False    ' mono/poly-drum...
            dmpM(BO) = Marked
        Next j
    Next I

    j = 0: While PPar(j).ShortName <> "M PITCH": j = j + 1: Wend
    For I = j To PParCount - 1
        For j = 0 To 15
            BO = 72 + j * 112 + PPar(I).ByteOffs
            dmpM(BO) = IIf(dmpB(BO) = PPar(I).Default, False, True)
        Next j
    Next I

    ' Rx. switches
    For j = 0 To 15
        iB = dmpB(72 + j * 112 + PPar(36).ByteOffs)
        iB = RotateLByte(iB)
        For I = 0 To 7
            Rx(I, j) = IIf((iB And (2 ^ (7 - I))) = 0, 0, 1)
        Next I
        iB = dmpB(72 + j * 112 + PPar(37).ByteOffs)
        iB = RotateRByte(iB)
        For I = 0 To 7
            Rx(I + 8, j) = IIf((iB And (2 ^ (I))) = 0, 0, 1)
        Next I
    Next j
    For I = 0 To 15
        iB = Rx(I, 7): Rx(I, 7) = Rx(I, 15): Rx(I, 15) = iB
    Next I

End Sub

' patchname according to dump
Public Function getPatchName(ByVal CHANNEL As Long) As String
    Dim iB As Byte, tB As Byte, bB As Byte
    Dim txt        As String

    iB = dmpB(72 + CHANNEL * 112 + PPar(1).ByteOffs)    'prg nr
    tB = dmpB(72 + CHANNEL * 112 + PPar(3).ByteOffs)    'M/P A D

    If (tB And &H60) = 0 Then
        bB = dmpB(72 + CHANNEL * 112 + PPar(0).ByteOffs)
        txt = isPatch(bB, iB)
    Else
        txt = isDrumSet(iB)
    End If

    getPatchName = txt
End Function

' read program data in "sc55par.dat"
Public Sub GetSC55Params(ByVal File As String)
    Dim I          As Long
    Dim ch         As Long
    Dim Regel      As String
    Dim modinfo    As Boolean
    Dim shortmsg   As Boolean
    Dim types      As Boolean
    Dim common     As Boolean
    Dim parts      As Boolean

    ch = FreeFile
    Open File For Input As ch
    While Not EOF(ch)
        If I = -1 Then I = I + 1
        Line Input #ch, Regel
        If UCase(Left(Regel, 12)) = "[MODULEINFO]" Then modinfo = True: I = -1
        If UCase(Left(Regel, 10)) = "[SHORTMSG]" Then shortmsg = True: modinfo = False: I = -1
        If UCase(Left(Regel, 10)) = "[PARTYPES]" Then types = True: shortmsg = False: I = -1
        If UCase(Left(Regel, 8)) = "[COMMON]" Then common = True: types = False: I = -1
        If UCase(Left(Regel, 7)) = "[PARTS]" Then parts = True: common = False: I = -1
        If modinfo = True And I > -1 Then
            If LCase(Left(Regel, 6)) = "manuid" Then mManuID = Val("&H" & Mid(Regel, InStr(6, Regel, "=") + 1))
            If LCase(Left(Regel, 8)) = "deviceid" Then mDeviceID = Val("&H" & Mid(Regel, InStr(8, Regel, "=") + 1))
            If LCase(Left(Regel, 7)) = "modelid" Then mModelID = Val("&H" & Mid(Regel, InStr(7, Regel, "=") + 1))
        End If
        If shortmsg = True And I > -1 Then
            If Len(Regel) > 3 Then
                ShMsg(I).CommandStr = Mid(Regel, 4, 20)
                ShMsg(I).name = Mid(Regel, 25)
                I = I + 1: ShMsgCount = I
            End If
        End If
        If types = True And I > -1 Then
            If splitSC55Type(Regel, PType(I)) Then I = I + 1: PTypesCount = I
        End If
        If common = True And I > -1 Then
            If splitSC55Param(Regel, APar(I)) Then I = I + 1: AParCount = I
        End If
        If parts = True And I > -1 Then
            If splitSC55Param(Regel, PPar(I)) Then I = I + 1: PParCount = I
        End If
    Wend
    Close ch
    If mManuID = 0 Then mManuID = &H41
    If mDeviceID = 0 Then mDeviceID = &H10
    If mModelID = 0 Then mModelID = &H42
End Sub

' fill in the macro structures, containing the default data
' for each macro
Public Sub GetMacros()
    RevM(0).name = "Room 1": RevM(0).rPL = 3: RevM(0).rFB = 0: RevM(0).rTM = &H50
    RevM(1).name = "Room 2": RevM(1).rPL = 4: RevM(1).rFB = 0: RevM(1).rTM = &H38
    RevM(2).name = "Room 3": RevM(2).rPL = 0: RevM(2).rFB = 0: RevM(2).rTM = &H40
    RevM(3).name = "Hall 1": RevM(3).rPL = 4: RevM(3).rFB = 0: RevM(3).rTM = &H48
    RevM(4).name = "Hall 2": RevM(4).rPL = 0: RevM(4).rFB = 0: RevM(4).rTM = &H40
    RevM(5).name = "Plate": RevM(5).rPL = 0: RevM(5).rFB = 0: RevM(5).rTM = &H58
    RevM(6).name = "Delay": RevM(6).rPL = 0: RevM(6).rFB = 0: RevM(6).rTM = &H20
    RevM(7).name = "Pan delay": RevM(7).rPL = 0: RevM(7).rFB = &H20: RevM(7).rTM = &H40

    ChoM(0).name = "Chorus 1": ChoM(0).cFB = 0: ChoM(0).cDL = &H70: ChoM(0).cRT = &H3: ChoM(0).cDP = &H5
    ChoM(1).name = "Chorus 2": ChoM(1).cFB = &H5: ChoM(1).cDL = &H50: ChoM(1).cRT = &H9: ChoM(1).cDP = &H13
    ChoM(2).name = "Chorus 3": ChoM(2).cFB = &H8: ChoM(2).cDL = &H50: ChoM(2).cRT = &H3: ChoM(2).cDP = &H13
    ChoM(3).name = "Chorus 4": ChoM(3).cFB = &H10: ChoM(3).cDL = &H40: ChoM(3).cRT = &H9: ChoM(3).cDP = &H10
    ChoM(4).name = "Feedback": ChoM(4).cFB = &H40: ChoM(4).cDL = &H7F: ChoM(4).cRT = &H2: ChoM(4).cDP = &H18
    ChoM(5).name = "Flanger": ChoM(5).cFB = &H70: ChoM(5).cDL = &H7F: ChoM(5).cRT = &H1: ChoM(5).cDP = &H5
    ChoM(6).name = "Delay": ChoM(6).cFB = 0: ChoM(6).cDL = &H7F: ChoM(6).cRT = 0: ChoM(6).cDP = &H7F
    ChoM(7).name = "Delay(FB)": ChoM(7).cFB = &H50: ChoM(7).cDL = &H7F: ChoM(7).cRT = 0: ChoM(7).cDP = &H7F

End Sub

' hard or/and software reset to default values
Public Sub GSResetAll()
    Dim msg        As String

    msg = "Reset all to GS?" & vbCrLf
    If hMidiOUT = 0 Then msg = msg & "(data only)"
    If MsgBox(msg, vbOKCancel, "RESET") = vbCancel Then Exit Sub
    Screen.MousePointer = vbHourglass
    SetDefaultDump
    If hMidiOUT <> 0 Then SysExDT1 makeComStr("40007F", 0, 0, 1, False)
    Screen.MousePointer = vbDefault
End Sub

' real representation of pan. Nr is between 0-127
Public Function isPanPot(ByVal Nr As Integer) As String
    Dim txt        As String
    Nr = Nr - 64
    If Nr = -64 Then
        txt = "Rnd"
    Else
        If Nr < 0 Then
            txt = "L" & Format(Abs(Nr))
        ElseIf Nr = 0 Then
            txt = "0"
        Else
            txt = "R" & Format(Nr)
        End If
    End If
    isPanPot = txt
End Function

' one tricky way of playing a midi file
' creates a html text with the EMBED-tag
Public Function playMidiFile(ByVal File As String) As String
    Dim txt        As String
    txt = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>" & vbCrLf
    txt = txt & "<HTML>" & vbCrLf
    txt = txt & "  <HEAD>" & vbCrLf
    txt = txt & "    <TITLE>Play Midi File</TITLE>" & vbCrLf
    txt = txt & "  </HEAD>" & vbCrLf
    txt = txt & "  <BODY BGCOLOR=#E8E8E8>" & vbCrLf
    txt = txt & "  <CENTER>" & vbCrLf
    txt = txt & "    <H1>Play Midi File</H1>" & vbCrLf
    txt = txt & "    <P>" & File & "<BR><BR>" & vbCrLf
    txt = txt & "    <EMBED SRC='" & File & "' WIDTH='256'></EMBED>" & vbCrLf
    txt = txt & "    </P>" & vbCrLf
    txt = txt & "    <P><BR><SMALL>In case you hear nothing after you push the start button:<BR>" & vbCrLf
    txt = txt & "    <OL>" & vbCrLf
    txt = txt & "    <LI>Wait at least several seconds</LI>" & vbCrLf
    txt = txt & "    <LI>Use the function <I>Close All ports</I> and try again</LI>" & vbCrLf
    txt = txt & "    <LI>Close all active sequencers (Cubase, CakeWalk,...)</LI>" & vbCrLf
    txt = txt & "    </OL></SMALL></P></CENTER>" & vbCrLf
    txt = txt & "  </BODY>" & vbCrLf
    txt = txt & "</HTML>" & vbCrLf
    playMidiFile = txt
End Function

' saves the alternative for the current dump
Public Sub SaveAlternative(ByVal AltShMsg As Boolean)
    Dim iB As Byte, jB As Byte    ' byte buffers
    Dim tB As Byte, B As Byte
    Dim I As Long, j As Long      ' counters
    Dim BO         As Long        ' byteoffset
    Dim txt        As String      ' file text generation
    Dim Marked     As Boolean     ' is byte marked?
    Dim ch         As Long        ' file handle
    Dim Lng        As Long        ' buffer
    Dim DT         As String * 1  ' delta-time
    Dim File       As String      ' filename
    Dim dmpCh      As Byte        ' channel/part

    CheckNonDefaults              ' mark bytes

    DT = Chr(24)                  ' constant delta-time (should be calculated)

    If AltNoCommon = True Then GoTo SaveAlternativePARTS:
    If StartWithGSReset = True Then
        txt = txt & DT & makeComStr("40007F", 0, 0, 1, True)
    End If

    ' common
    I = 1                         'Master volume
    iB = dmpB(APar(I).ByteOffs)
    If iB <> APar(I).Default Then txt = txt & DT & makeComStr(APar(I).Address, 0, iB, PType(APar(I).Type).Bytes, True)
    I = 3                         'Master panpot
    iB = dmpB(APar(I).ByteOffs)
    If iB <> APar(I).Default Then txt = txt & DT & makeComStr(APar(I).Address, 0, iB, PType(APar(I).Type).Bytes, True)
    'reverb chorus
    For I = 6 To AParCount - 1
        BO = APar(I).ByteOffs
        If dmpM(BO) = True Then
            txt = txt & DT & makeComStr(APar(I).Address, 0, dmpB(BO), PType(APar(I).Type).Bytes, True)
        End If
    Next I

SaveAlternativePARTS:
    ' parts parameters
    For j = 0 To 15
        If PartSwitch(j) = True Then GoTo SaveAlternativeNEXTPART:
        dmpCh = Choose(j + 1, 1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 10, 11, 12, 13, 14, 15)
        ' drumpart?
        iB = dmpB(72 + dmpCh * 112 + PPar(3).ByteOffs)    'M/P A D
        If Not ((iB And &H60) = 0) Then
            iB = (iB And &H60) \ 32
            If Not (iB = 1 And dmpCh = 0) Then    ' not default drum part
                txt = txt & DT & makeComStr("401n15", dmpCh, iB, 25, True)
            End If
        End If
        ' bank & progr.no.
        I = 0
        iB = dmpB(72 + dmpCh * 112 + PPar(0).ByteOffs)    'bank
        jB = dmpB(72 + dmpCh * 112 + PPar(1).ByteOffs)    'prg nr
        If iB <> 0 Or jB <> 0 Then    ' one of those not null = non default
            If AltShMsg Then      ' using short msg?
                If iB <> 0 Then   ' bank nr <>0 = non-default
                    txt = txt & DT & Chr(&HB0 Or j) & Chr(0) & Chr(iB)
                    txt = txt & DT & Chr(&H20) & Chr(0)
                End If
                txt = txt & DT & Chr(&HC0 Or j) & Chr(jB)
            Else
                txt = txt & DT & makeComStr(PPar(I).Address, dmpCh, jB * 256 + iB, PPar(I).Type, True)
            End If
        End If

        ' other params
        For I = 2 To PParCount - 1
            If Left(PPar(I).ShortName, 2) <> "RX" And Left(PPar(I).ShortName, 2) <> "M/" Then
                BO = 72 + dmpCh * 112 + PPar(I).ByteOffs
                If dmpM(BO) = True Then
                    If AltShMsg And PPar(I).Altern <> 0 Then
                        txt = txt & DT & makeAltShMsg(PPar(I).Altern, j, dmpB(BO))
                    Else
                        txt = txt & DT & makeComStr(PPar(I).Address, dmpCh, dmpB(BO), PPar(I).Type, True)
                    End If
                End If
            End If
        Next I
SaveAlternativeNEXTPART:
    Next j

    File = SaveDumpAltFile
    If InStr(File, ":\") = 0 Then File = App.Path & "\" & File
    ch = FreeFile
    Open File For Binary As ch
    Put #ch, 1, "MThd"
    B = 0: Put #ch, , B: Put #ch, , B: Put #ch, , B: Put #ch, , CByte(6)
    Put #ch, , B: Put #ch, , B    ' formattype=0
    Put #ch, , B: Put #ch, , CByte(1)    ' numtracks
    Put #ch, , CByte(1): Put #ch, , CByte(128)    ' division
    Put #ch, , "MTrk"
    Lng = 10 + Len(txt) + 3       ' numbytes
    Put #ch, , CByte(Lng \ (256 ^ 3))
    Put #ch, , CByte((Lng \ (256 ^ 2)) And 255)
    Put #ch, , CByte((Lng \ 256) And 255)
    Put #ch, , CByte(Lng And 255)
    B = 0: Put #ch, , B           ' DT=0
    B = &HFF: Put #ch, , B        ' trackname FF 03
    B = 3: Put #ch, , B
    B = 7: Put #ch, , B: Put #ch, , "DumpAlt"    ' strlen+str
    Put #ch, , txt                ' generated alternatives
    Put #ch, , DT
    B = &HFF: Put #ch, , B        ' end of track
    B = &H2F: Put #ch, , B
    B = 0: Put #ch, , B
    Close ch
End Sub

' shortmessage and dump version of channel/part no. or different
' input channel = 0 - 15
Public Sub SetChannel(ByVal CHANNEL As Integer)
    CurChannel = CHANNEL
    CurDChannel = Choose(CurChannel + 1, 1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 10, 11, 12, 13, 14, 15)
End Sub

' makes the alternative for the current dump, only to display them
Public Function strAlternative(ByVal AltShMsg As Boolean) As String
    Dim iB As Byte, jB As Byte
    Dim I As Long, j As Long, BO As Long, dmpCh As Long
    Dim txt        As String

    CheckNonDefaults
    txt = "Alternat - SC55 dump : " & CurDmpMidFileTitle & vbCrLf & vbCrLf

    If AltNoCommon = True Then GoTo strAlternativePARTS:
    If StartWithGSReset = True Then
        txt = txt & "Reset" & Space(29) & getComStrHex(makeComStr("40007F", 0, 0, 1, False)) & vbCrLf
    End If

    txt = txt & vbCrLf & "  *** COMMON ***" & vbCrLf
    ' common
    I = 1                         'Master volume
    iB = dmpB(APar(I).ByteOffs)
    If iB <> APar(I).Default Then
        txt = txt & APar(I).name & "  "
        txt = txt & strSYXDataBVal(Hex(iB), 0, APar(I).Type, 1, 8) & " -> "
        txt = txt & getComStrHex(makeComStr(APar(I).Address, 0, iB, PType(APar(I).Type).Bytes, False)) & vbCrLf
    End If
    I = 3                         'Master panpot
    iB = dmpB(APar(I).ByteOffs)
    If iB <> APar(I).Default Then
        txt = txt & APar(I).name & "  "
        txt = txt & strSYXDataBVal(Hex(iB), 0, APar(I).Type, 1, 8) & " -> "
        txt = txt & getComStrHex(makeComStr(APar(I).Address, 0, iB, PType(APar(I).Type).Bytes, False)) & vbCrLf
    End If
    'reverb chorus
    For I = 6 To AParCount - 1
        iB = dmpB(APar(I).ByteOffs)
        If dmpM(APar(I).ByteOffs) = True Then
            txt = txt & APar(I).name & "  "
            txt = txt & strSYXDataBVal(Hex(iB), 0, APar(I).Type, 1, 8) & " -> "
            txt = txt & getComStrHex(makeComStr(APar(I).Address, 0, iB, PType(APar(I).Type).Bytes, False)) & vbCrLf
        End If
    Next I

strAlternativePARTS:
    ' parts parameters
    For j = 0 To 15
        If PartSwitch(j) = True Then GoTo strAlternativeNEXTPART:
        dmpCh = Choose(j + 1, 1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 10, 11, 12, 13, 14, 15)
        txt = txt & vbCrLf & "  *** PART " & CStr(j + 1) & " ***" & vbCrLf
        ' drumpart?
        iB = dmpB(72 + dmpCh * 112 + PPar(3).ByteOffs)    'M/P A D
        If Not ((iB And &H60) = 0) Then
            iB = (iB And &H60) \ 32
            If Not (iB = 1 And dmpCh = 0) Then    ' not default drum part
                txt = txt & "Part Mode             "
                txt = txt & FixStr("Drum " & CStr(iB), 8, " ") & " -> "
                txt = txt & getComStrHex(makeComStr("401n15", dmpCh, iB, 25, False)) & vbCrLf
            End If
        End If
        ' bank & progr.no.
        I = 0
        iB = dmpB(72 + dmpCh * 112 + PPar(I).ByteOffs)
        jB = dmpB(72 + dmpCh * 112 + PPar(I).ByteOffs + 1)
        If iB <> 0 Or jB <> 0 Then    ' one of those not null = non default
            If AltShMsg Then
                If iB <> 0 Then
                    txt = txt & PPar(I).name & "  "
                    txt = txt & strSYXDataBVal(Hex(iB), 0, PPar(I).Type, 1, 8) & " -> "
                    txt = txt & HexByte(&HB0 Or j) & " " & HexByte(0) & " " & HexByte(iB) & " " & HexByte(&HB0 Or j) & " 20 00" & vbCrLf
                End If
                txt = txt & PPar(1).name & "  "
                txt = txt & strSYXDataBVal(Hex(jB), 0, PPar(1).Type, 1, 8) & " -> "
                txt = txt & HexByte(&HC0 Or j) & " " & HexByte(jB) & vbCrLf
            Else
                txt = txt & PPar(1).name & "  "
                txt = txt & strSYXDataBVal(Hex(jB), 0, PPar(1).Type, 1, 8) & " -> "
                txt = txt & getComStrHex(makeComStr(PPar(1).Address, dmpCh, jB * 256 + iB, PPar(I).Type, False)) & vbCrLf
            End If
        End If

        ' other params
        For I = 2 To PParCount - 1
            If Left(PPar(I).ShortName, 2) <> "RX" And Left(PPar(I).ShortName, 2) <> "M/" Then
                BO = 72 + dmpCh * 112 + PPar(I).ByteOffs
                iB = dmpB(BO)
                If dmpM(BO) = True Then
                    txt = txt & PPar(I).name & "  "
                    txt = txt & strSYXDataBVal(Hex(iB), 0, PPar(I).Type, 1, 8) & " -> "
                    If AltShMsg And PPar(I).Altern <> 0 Then
                        txt = txt & getComStrHex(makeAltShMsg(PPar(I).Altern, j, iB)) & vbCrLf
                    Else
                        txt = txt & getComStrHex(makeComStr(PPar(I).Address, dmpCh, iB, PPar(I).Type, False)) & vbCrLf
                    End If
                End If
            End If
        Next I
strAlternativeNEXTPART:
    Next j
    strAlternative = txt
End Function

' dump channel/part to shortmessage version (0 - 15)
Public Function isChannel(ByVal ch As Long) As Long
    isChannel = Choose(ch + 1, 9, 0, 1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14, 15)
End Function

' returns a panpot string
' returns the proper patchname according to bank nr and instrument nr
Public Function isPatch(ByVal BANK As Integer, ByVal Nr As Integer) As String
    Dim txt As String, K As String
    Dim B As Integer, PB As Integer

    Select Case BANK
        Case 0: B = 1: K = " "
        Case 127: B = 2: K = "#"
        Case Else:
            If Nr < 120 Then
                If BANK = 8 Then
                    B = 3: K = "+"
                ElseIf BANK = 10 Then
                    B = 4: K = "+"
                Else
                    B = 1: K = " "
                End If
            Else
                B = BANK + 2
                K = "+"
            End If

    End Select
    PB = B
    Select Case Nr
        Case 0: B = IIf(B > 2, 1, B): txt = Choose(B, "Piano 1", "Acou Piano 1")
        Case 1: B = IIf(B > 2, 1, B): txt = Choose(B, "Piano 2", "Acou Piano 2")
        Case 2: B = IIf(B > 2, 1, B): txt = Choose(B, "Piano 3", "Acou Piano 3")
        Case 3: B = IIf(B > 2, 1, B): txt = Choose(B, "Honkytonk", "Elec Piano 1")
        Case 4: B = IIf(B > 3, 1, B): txt = Choose(B, "E. Piano 1", "Elec Piano 2", "Detuned EP 1")
        Case 5: B = IIf(B > 3, 1, B): txt = Choose(B, "E. Piano 2", "Elec Piano 3", "Detuned EP 2")
        Case 6: B = IIf(B > 3, 1, B): txt = Choose(B, "Harpsichord", "Elec Piano 4", "Coupled Hps.")
        Case 7: B = IIf(B > 2, 1, B): txt = Choose(B, "Clavinet", "Honkytonk")
        Case 8: B = IIf(B > 2, 1, B): txt = Choose(B, "Celesta", "Elec Org 1")
        Case 9: B = IIf(B > 2, 1, B): txt = Choose(B, "Glockenspiel", "Elec Org 2")
        Case 10: B = IIf(B > 2, 1, B): txt = Choose(B, "Music Box", "Elec Org 3")
        Case 11: B = IIf(B > 2, 1, B): txt = Choose(B, "Vibraphone", "Elec Org 4")
        Case 12: B = IIf(B > 2, 1, B): txt = Choose(B, "Marimba", "Pipe Org 1")
        Case 13: B = IIf(B > 2, 1, B): txt = Choose(B, "Xylophone", "Pipe Org 2")
        Case 14: B = IIf(B > 3, 1, B): txt = Choose(B, "Tubular Bell", "Pipe Org 3", "Church Bell")
        Case 15: B = IIf(B > 2, 1, B): txt = Choose(B, "Santur", "Accordion")
        Case 16: B = IIf(B > 3, 1, B): txt = Choose(B, "Organ 1", "Harpsi 1", "Detuned Or. 1")
        Case 17: B = IIf(B > 3, 1, B): txt = Choose(B, "Organ 2", "Harpsi 2", "Detuned Or. 2")
        Case 18: B = IIf(B > 2, 1, B): txt = Choose(B, "Organ 3", "Harpsi 3")
        Case 19: B = IIf(B > 3, 1, B): txt = Choose(B, "Church Org. 1", "Clavi 1", "Church Org 2")
        Case 20: B = IIf(B > 2, 1, B): txt = Choose(B, "Reed Organ", "Clavi 2")
        Case 21: B = IIf(B > 3, 1, B): txt = Choose(B, "Accordion Fr", "Clavi 3", "Accordion It")
        Case 22: B = IIf(B > 2, 1, B): txt = Choose(B, "Harmonica", "Celesta 1")
        Case 23: B = IIf(B > 2, 1, B): txt = Choose(B, "Bandneon", "Celesta 2")
        Case 24: B = IIf(B > 3, 1, B): txt = Choose(B, "Nylon str. Gt", "Syn Brass 1", "Ukulele")
        Case 25: B = IIf(B > 4, 1, B): txt = Choose(B, "Steel str. Gt", "Syn Brass 2", "12str. Gt", "Mandolin")
        Case 26: B = IIf(B > 3, 1, B): txt = Choose(B, "Jazz Gt", "Syn Brass 3", "Hawaian Gt")
        Case 27: B = IIf(B > 3, 1, B): txt = Choose(B, "Clean Gt", "Syn Brass 4", "Chorus Gt")
        Case 28: B = IIf(B > 3, 1, B): txt = Choose(B, "Muted Gt", "Syn Bass 1", "Funk Gt")
        Case 29: B = IIf(B > 2, 1, B): txt = Choose(B, "Overdrive Gt", "Syn Bass 2")
        Case 30: B = IIf(B > 3, 1, B): txt = Choose(B, "DystortionGt", "Syn Bass 3", "Feedback Gt")
        Case 31: B = IIf(B > 3, 1, B): txt = Choose(B, "Gt Harmonics", "Syn Bass 4", "Gt Feedback")
        Case 32: B = IIf(B > 2, 1, B): txt = Choose(B, "Acoustic Bs.", "Fantasy")
        Case 33: B = IIf(B > 2, 1, B): txt = Choose(B, "Fingered Bs.", "Harmo Pan")
        Case 34: B = IIf(B > 2, 1, B): txt = Choose(B, "Picked Bs.", "Chorale")
        Case 35: B = IIf(B > 2, 1, B): txt = Choose(B, "Fretless Bs.", "Glasses")
        Case 36: B = IIf(B > 2, 1, B): txt = Choose(B, "Slap Bs 1", "Soundtrack")
        Case 37: B = IIf(B > 2, 1, B): txt = Choose(B, "Slap Bs 2", "Atmosphere")
        Case 38: B = IIf(B > 3, 1, B): txt = Choose(B, "Synth Bass 1", "Warm Bell", "Synth Bass 3")
        Case 39: B = IIf(B > 3, 1, B): txt = Choose(B, "Synth Bass 2", "Funny Vox", "Synth Bass 4")
        Case 40: B = IIf(B > 2, 1, B): txt = Choose(B, "Violin", "Echo Bell")
        Case 41: B = IIf(B > 2, 1, B): txt = Choose(B, "Viola", "Ice Rain")
        Case 42: B = IIf(B > 2, 1, B): txt = Choose(B, "Cello", "Oboe 2001")
        Case 43: B = IIf(B > 2, 1, B): txt = Choose(B, "Contrabass", "Echo Pan")
        Case 44: B = IIf(B > 2, 1, B): txt = Choose(B, "Tremolo Str", "Doctor Solo")
        Case 45: B = IIf(B > 2, 1, B): txt = Choose(B, "PizzicatoStr", "School Daze")
        Case 46: B = IIf(B > 2, 1, B): txt = Choose(B, "Harp", "Bellsinger")
        Case 47: B = IIf(B > 2, 1, B): txt = Choose(B, "Timpani", "Square Wave")
        Case 48: B = IIf(B > 3, 1, B): txt = Choose(B, "Strings", "Str Sect 1", "Orchestra")
        Case 49: B = IIf(B > 2, 1, B): txt = Choose(B, "Slow String", "Str Sect 2")
        Case 50: B = IIf(B > 3, 1, B): txt = Choose(B, "Synth Strings1", "Str Sect 3", "Syn. Strings3")
        Case 51: B = IIf(B > 2, 1, B): txt = Choose(B, "Synth Strings1", "Pizzicato")
        Case 52: B = IIf(B > 2, 1, B): txt = Choose(B, "Choir Aahs", "Violin 1")
        Case 53: B = IIf(B > 2, 1, B): txt = Choose(B, "Voice Oohs", "Violin 2")
        Case 54: B = IIf(B > 2, 1, B): txt = Choose(B, "SynVox", "Cello 2")
        Case 55: B = IIf(B > 2, 1, B): txt = Choose(B, "OrchestraHit", "Cello 2")
        Case 56: B = IIf(B > 2, 1, B): txt = Choose(B, "Trumpet", "Contrabass")
        Case 57: B = IIf(B > 2, 1, B): txt = Choose(B, "Trombone", "Harp 1")
        Case 58: B = IIf(B > 2, 1, B): txt = Choose(B, "Tuba", "Harp 2")
        Case 59: B = IIf(B > 2, 1, B): txt = Choose(B, "MutedTrumpet", "Guitar 1")
        Case 60: B = IIf(B > 2, 1, B): txt = Choose(B, "French Horn", "Guitar 2")
        Case 61: B = IIf(B > 3, 1, B): txt = Choose(B, "Brass 1", "Elec Gtr 1", "Brass 2")
        Case 62: B = IIf(B > 3, 1, B): txt = Choose(B, "Synth Brass1", "Elec Gtr 2", "Synth Brass3")
        Case 63: B = IIf(B > 3, 1, B): txt = Choose(B, "Synth Brass2", "Sitar", "Synth Brass4")
        Case 64: B = IIf(B > 2, 1, B): txt = Choose(B, "Soprano Sax", "Acou Bass 1")
        Case 65: B = IIf(B > 2, 1, B): txt = Choose(B, "Alto Sax", "Acou Bass 2")
        Case 66: B = IIf(B > 2, 1, B): txt = Choose(B, "Tenor Sax", "Elec Bass 1")
        Case 67: B = IIf(B > 2, 1, B): txt = Choose(B, "Baritone Sax", "Elec Bass 2")
        Case 68: B = IIf(B > 2, 1, B): txt = Choose(B, "Oboe", "Slap Bass 1")
        Case 69: B = IIf(B > 2, 1, B): txt = Choose(B, "English Horn", "Slap Bass 2")
        Case 70: B = IIf(B > 2, 1, B): txt = Choose(B, "Bassoon", "Fretless 1")
        Case 71: B = IIf(B > 2, 1, B): txt = Choose(B, "Clarinet", "Fretless 2")
        Case 72: B = IIf(B > 2, 1, B): txt = Choose(B, "Piccolo", "Flute 1")
        Case 73: B = IIf(B > 2, 1, B): txt = Choose(B, "Flute", "Flute 2")
        Case 74: B = IIf(B > 2, 1, B): txt = Choose(B, "Recorder", "Piccolo 1")
        Case 75: B = IIf(B > 2, 1, B): txt = Choose(B, "Pan Flute", "Piccolo 2")
        Case 76: B = IIf(B > 2, 1, B): txt = Choose(B, "Bottle Blow", "Recorder")
        Case 77: B = IIf(B > 2, 1, B): txt = Choose(B, "Shakuhachi", "Pan Pipes")
        Case 78: B = IIf(B > 2, 1, B): txt = Choose(B, "Whistle", "Sax 1")
        Case 79: B = IIf(B > 2, 1, B): txt = Choose(B, "Ocarina", "Sax 2")
        Case 80: B = IIf(B > 2, 1, B): txt = Choose(B, "Square Wave", "Sax 3")
        Case 81: B = IIf(B > 2, 1, B): txt = Choose(B, "Saw Wave", "Sax 4")
        Case 82: B = IIf(B > 2, 1, B): txt = Choose(B, "Syn. Calliope", "Clarinet 1")
        Case 83: B = IIf(B > 2, 1, B): txt = Choose(B, "Chiffer Lead", "Clarinet 2")
        Case 84: B = IIf(B > 2, 1, B): txt = Choose(B, "Charang", "Oboe")
        Case 85: B = IIf(B > 2, 1, B): txt = Choose(B, "Solo Vox", "Engl Horn")
        Case 86: B = IIf(B > 2, 1, B): txt = Choose(B, "5th Saw. Wave", "Basson")
        Case 87: B = IIf(B > 2, 1, B): txt = Choose(B, "Bass & Lead", "Harmonica")
        Case 88: B = IIf(B > 2, 1, B): txt = Choose(B, "Fantasia", "Trumpet 1")
        Case 89: B = IIf(B > 2, 1, B): txt = Choose(B, "Warm Pad", "Trumpet 2")
        Case 90: B = IIf(B > 2, 1, B): txt = Choose(B, "Polysynth", "Trombone 1")
        Case 91: B = IIf(B > 2, 1, B): txt = Choose(B, "Space Voice", "Trombone 2")
        Case 92: B = IIf(B > 2, 1, B): txt = Choose(B, "Bowed Glass", "Fr Horn 1")
        Case 93: B = IIf(B > 2, 1, B): txt = Choose(B, "Metal Pad", "Fr Horn 2")
        Case 94: B = IIf(B > 2, 1, B): txt = Choose(B, "Halo Pad", "Tuba")
        Case 95: B = IIf(B > 2, 1, B): txt = Choose(B, "Sweep Pad", "Brs Sect 1")
        Case 96: B = IIf(B > 2, 1, B): txt = Choose(B, "Ice Rain", "Brs Sect 2")
        Case 97: B = IIf(B > 2, 1, B): txt = Choose(B, "Soundtrack", "Vibe 1")
        Case 98: B = IIf(B > 2, 1, B): txt = Choose(B, "Crystal", "Vibe 2")
        Case 99: B = IIf(B > 2, 1, B): txt = Choose(B, "Athmosphere", "Syn Mallet")
        Case 100: B = IIf(B > 2, 1, B): txt = Choose(B, "Brightness", "Windbell")
        Case 101: B = IIf(B > 2, 1, B): txt = Choose(B, "Goblin", "Glock")
        Case 102: B = IIf(B > 2, 1, B): txt = Choose(B, "Echo Drops", "Tube Bell")
        Case 103: B = IIf(B > 2, 1, B): txt = Choose(B, "Start Theme", "Xylophone")
        Case 104: B = IIf(B > 2, 1, B): txt = Choose(B, "Sitar", "Marimba")
        Case 105: B = IIf(B > 2, 1, B): txt = Choose(B, "Banjo", "Koto")
        Case 106: B = IIf(B > 2, 1, B): txt = Choose(B, "Shamisen", "Sho")
        Case 107: B = IIf(B > 3, 1, B): txt = Choose(B, "Koto", "Shakuhachi", "Taisho Koto")
        Case 108: B = IIf(B > 2, 1, B): txt = Choose(B, "Kalimba", "Whistle 1")
        Case 109: B = IIf(B > 2, 1, B): txt = Choose(B, "Bag Pipe", "Whistle 2")
        Case 110: B = IIf(B > 2, 1, B): txt = Choose(B, "Fiddle", "Bottleblow")
        Case 111: B = IIf(B > 2, 1, B): txt = Choose(B, "Shanai", "Breathpipe")
        Case 112: B = IIf(B > 2, 1, B): txt = Choose(B, "Tinkle Bell", "Timpani")
        Case 113: B = IIf(B > 2, 1, B): txt = Choose(B, "Agogo", "Melodic Tom")
        Case 114: B = IIf(B > 2, 1, B): txt = Choose(B, "Steel Drums", "Deep Snare")
        Case 115: B = IIf(B > 3, 1, B): txt = Choose(B, "Woodblock", "Elec Perc 1", "Castanets")
        Case 116: B = IIf(B > 3, 1, B): txt = Choose(B, "Taiko", "Elec Perc 2", "Concert BD")
        Case 117: B = IIf(B > 3, 1, B): txt = Choose(B, "Melo Tom 1", "Taiko", "Melo Tom 2")
        Case 118: B = IIf(B > 3, 1, B): txt = Choose(B, "Synth Drum", "Taiko Rim", "808 Tom")
        Case 119: B = IIf(B > 2, 1, B): txt = Choose(B, "Reverse Cym.", "Cymbal")
        Case 120: B = IIf(B > 4, 1, B): txt = Choose(B, "Gt FretNoise", "Castanets", "Gt Cut Noise", "String Slap")
        Case 121: B = IIf(B > 2, 1, B): txt = Choose(B, "Fl. Keyclick", "Triangle")
        Case 122: B = IIf(B > 7, 1, B): txt = Choose(B, "Seashore", "Orchestra Hit", "Rain", "Thunder", "Wind", "Stream", "Bubble")
        Case 123: B = IIf(B > 4, 1, B): txt = Choose(B, "Bird", "Telephone", "Dog", "Horse")
        Case 124: B = IIf(B > 7, 1, B): txt = Choose(B, "Telephone 1", "Bird Tweet", "Telephone 2", "DoorCreaking", "Door", "Scratch", "Windchime")
        Case 125: B = IIf(B > 11, 1, B): txt = Choose(B, "Helicopter", "One Note Jam", "Car engine", "Car stop", "Car pass", "Car crash", "Siren", "Train", "Jetplane", "Starship", "Burst Noise")
        Case 126: B = IIf(B > 8, 1, B): txt = Choose(B, "Applause", "Water Bell", "Applause", "Laughing", "Screaming", "Punch", "Heart Beat", "Footstep")
        Case 127: B = IIf(B > 5, 1, B): txt = Choose(B, "Gun Shot", "Jungle Tune", "Machinegun", "Lasergun", "Explosion")
    End Select
    If PB <> B Then K = " "
    isPatch = Format(Nr + 1, "000") & K & txt
End Function

' is SC-55 specific
' nr is between 0-127, but not all of them are used
Public Function isDrumSet(ByVal Nr As Integer) As String
    Dim txt        As String
    Select Case Nr
        Case 0: txt = "001*Standard"
        Case 8: txt = "009*Room"
        Case 16: txt = "017*Power"
        Case 24: txt = "025*Electronic"
        Case 25: txt = "026*TR-808"
        Case 32: txt = "033*Jazz"
        Case 40: txt = "041*Brush"
        Case 48: txt = "049*Orchestra"
        Case 56: txt = "057*SFX"
        Case 127: txt = "128*CM-64/32L"
        Case Else: txt = "????"
    End Select
    isDrumSet = txt
End Function

' is in fact not realy SC-55 specific
Public Function isNote(ByVal Nr As Long) As String
    isNote = Nr
    '  Dim octaaf As Long
    '  Dim noot As String
    '  octaaf = (Nr \ 12)
    '  noot = Nr Mod 12
    '  isNote = Choose(noot + 1, "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B") & Format(octaaf - 1)
End Function

' Gives the best, real presentation of a parameter value
' based upon the parameter type.
' the value is passed as (hex)text, so byte or 2byte integer
' values can be passed as well as text itself.
Public Function isValue(ByVal txt As String, pt As Byte) As String
    Dim Value      As Single
    Dim ntxt       As String
    Dim minH As Single, maxH As Single
    Dim minD As Single, maxD As Single

    Value = Val("&H" & txt)
    Select Case pt                'parameter type
        Case 0, 13                ' no type, no changes
            ntxt = txt
        Case 1, 5, 11
            ntxt = Format(Value)
        Case 4                    'reverb macro
            ntxt = Trim(RevM(Value).name)
        Case 6                    'chorus macro
            ntxt = Trim(ChoM(Value).name)
        Case 9                    'master tune
            ntxt = Format((Exp((Value - 1024) / 17300)) * 440, "#0.0")
        Case 12                   'bank/prg
            ntxt = Format(Value + 1)
        Case 14                   'channel
            ntxt = IIf(Value < 16, Format(Value + 1), "Off")
        Case 15                   'mono/poly
            If (Value And &H80) = 0 Then ntxt = "M" Else ntxt = "P"
            ntxt = ntxt & "-" & Format((Value And &H3))
            ntxt = ntxt & "-" & Format((Value And &H60) \ 32)
        Case 21                   'pan
            ntxt = isPanPot(Value)
        Case 22                   'k range
            ntxt = isNote(Value)
        Case Else
            minH = PType(pt).MinHex
            maxH = PType(pt).MaxHex
            minD = PType(pt).MinDec
            maxD = PType(pt).MaxDec
            'Value = minD + (Value - minH) * ((maxD - minD) / (maxH - minH))
            Value = Convert(Value, minH, maxH, minD, maxD)
            ntxt = IIf(CInt(Value) = Value, Format(Value), Format(Value, "#0.0"))
    End Select
    isValue = ntxt
End Function

' makes one alternative short message
' shtmsg=ID, Chann=channel/part
Public Function makeAltShMsg(ByVal shtmsg As Long, ByVal Chan As Byte, ByVal Value As Long) As String
    Dim I          As Long
    Dim txt As String, Part As String
    Dim sht        As String

    sht = Trim(ShMsg(shtmsg).CommandStr)
    For I = 1 To Len(sht) Step 3
        Part = Mid(sht, I, 2)
        If Right(Part, 1) = "n" Then
            txt = txt & Chr(Val("&H" & Left(Part, 1)) * 16 + Chan)
        ElseIf Part = "vv" Then
            txt = txt & Chr(Value)
        Else
            txt = txt & Chr(Val("&H" & Part))
        End If
        If shtmsg >= 13 And shtmsg <= 20 Then
            If I = 7 Or I = 13 Then txt = txt & Chr(96)    ' &H60
        End If
    Next I
    makeAltShMsg = txt
End Function

' makes a SysEx commandstr
' SveAlt= when saving in a midi file, a Big Endian should be
' inserted after the &HF0 with the no. of following sysex bytes
Public Function makeComStr(ByVal Address As String, ByVal Chan As Byte, ByVal Value As Long, ByVal tpe As Long, ByVal SveAlt As Boolean) As String
    Dim I          As Long
    Dim B(20) As Byte, Data As String
    Dim Bytes As Integer, sumID As Integer

    B(0) = &HF0
    B(1) = mManuID
    B(2) = mDeviceID
    B(3) = mModelID
    B(4) = &H12                   ' data set 1
    B(5) = Val("&H" & Left(Address, 2))
    If InStr(1, Address, "n") > 0 Then
        B(6) = Val("&H" & Mid(Address, 3, 1)) * 16 Or Chan
    Else
        B(6) = Val("&H" & Mid(Address, 3, 2))
    End If
    B(7) = Val("&H" & Right(Address, 2))
    Select Case Val(PType(tpe).Bytes)    ' number of bytes
        Case 1
            B(8) = Value
            Bytes = 11
        Case 2
            If tpe = 12 Then
                B(8) = Value And &H7F
                B(9) = (Value \ 256) And &H7F
            Else
                B(8) = Value And &HF    'nibblize
                B(9) = (Value And &HF0) \ 16
            End If
            Bytes = 12
        Case 4
            B(8) = Value And &HF  'nibblize
            B(9) = (Value And &HF0) \ 16
            B(10) = (Value And &HF00) \ 256
            B(11) = (Value And &HF000) \ 4096
            Bytes = 14
    End Select
    sumID = Bytes - 2
    B(sumID) = 0
    B(Bytes - 1) = &HF7
    For I = 5 To sumID - 1: B(sumID) = CByte(CInt(B(sumID) + CInt(B(I))) Mod 255): Next I
    B(sumID) = -B(sumID) And 127
    For I = 0 To Bytes - 1
        If I = 1 And SveAlt = True Then Data = Data & Chr(Bytes - 1)
        Data = Data & Chr(B(I))
    Next I
    makeComStr = Data
End Function

' File = midi file with syx dump at start of file
Public Function SYXdumpOpen(ByVal File As String, ByRef ret As String) As Boolean
    Dim B          As Byte        ' byte buffers
    Dim B1 As Byte, B2 As Byte, B3 As Byte, B4 As Byte
    Dim txt As String, fout As String    ' errors
    Dim ch         As Long        ' file handle
    Dim I As Long, pos As Long    ' counter, position in file
    Dim MT         As String * 4
    Dim Division   As Integer     ' Pulses Per Quarter Note
    Dim NumBytes   As Long        ' track len
    Dim Bytes      As Long        ' var length long
    Dim bBytes     As Byte        ' fixed 1 byte
    Dim DrumPatch  As Long        ' position of drumpatch
    Dim TrackInfo  As Boolean
    Dim Trackname  As String
    Dim Tempo      As Long        ' Beats/Minute
    Dim SMPTEOffs  As Long
    Dim TSnn As Byte, TSdd As Byte    ' time sign
    Dim TScc As Byte, TSbb As Byte    ' time sign cc=clocks/metronomclick bb=32nd/quarter notes
    Dim KSsf       As Byte        ' Key signature -7 for 7 flats, -1 for 1 flat, etc, 0 for key of c, 1 for 1 sharp, etc.
    Dim KSmi       As Byte        ' Key signature 0=major/1=minor
    Dim nB(7360)   As Byte        ' nibblized bytes
    Dim nBc        As Long        ' nibble bytes count
    Dim DT         As Long        ' delta time
    Dim DataAccepted As Boolean   ' new data ok?

    ch = FreeFile
    Open File For Binary As ch

    ' header
    Get #ch, 1, MT
    If MT <> "MThd" Then MsgBox "This is not a midifile.": GoTo SYXdumpOpenEND
    For I = 5 To 12
        Get #ch, I, B
        If Choose(I - 4, 0, 0, 0, 6) <> B Then MsgBox "This midifile is corrupted!.": GoTo SYXdumpOpenEND
    Next I

    ' format type
    For I = 9 To 12
        Get #ch, I, B
        If Choose(I - 8, 0, 0, 0, 1) <> B Then MsgBox "This is not a midifile with ONE dump TRACK.": GoTo SYXdumpOpenEND
    Next I

    ' division
    Get #ch, 13, B1: Get #ch, 14, B2
    Division = B1 * 256 + B2

    ' track 1
    Get #ch, 15, MT
    If MT <> "MTrk" Then MsgBox "Track is displaced or missing!": GoTo SYXdumpOpenEND
    Get #ch, 19, B1: Get #ch, 20, B2: Get #ch, 21, B3: Get #ch, 22, B4
    NumBytes = CLng(B1) * 256 ^ 3 + CLng(B2) * 256 ^ 2 + CLng(B3) * 256 + CLng(B4)
    If NumBytes < 8000 Then fout = fout & "This is not a midifile with a sc55 dump track": GoTo SYXdumpOpenEND

    ' track data
    pos = 23
    DT = readVarLen(ch, pos)      ' Big Endian delta time
    Get #ch, pos, B2: pos = pos + 1
    If B2 = &HFF Then
        While B2 = &HFF
            Get #ch, pos, B3: pos = pos + 1
            If B3 = 3 Then        ' trackname
                Bytes = readVarLen(ch, pos)
                For I = 1 To Bytes
                    Get #ch, pos, B: pos = pos + 1: Trackname = Trackname & Chr(B)
                Next I
            ElseIf B3 = &H51 Then    ' tempo
                Get #ch, pos, bBytes: pos = pos + 1
                For I = 1 To bBytes
                    Get #ch, pos, B: pos = pos + 1
                    Tempo = Tempo + CLng(B) * 256 ^ (bBytes - I)
                Next I
                If bBytes <> 3 Then fout = fout & "Tempo should be stored with 3 bytes" & vbCrLf
            ElseIf B3 = &H54 Then    ' SMPTE Offs
                Get #ch, pos, bBytes: pos = pos + 1
                For I = 1 To bBytes
                    Get #ch, pos, B: pos = pos + 1
                    SMPTEOffs = CLng(B) * 256 ^ (bBytes - I)
                Next I
                If bBytes <> 5 Then fout = fout & "SMPTE Offs should be stored with 5 bytes" & vbCrLf
            ElseIf B3 = &H58 Then    ' time sign
                Get #ch, pos, bBytes: pos = pos + 1
                If bBytes <> 4 Then fout = fout & "Time signature should be stored with 4 bytes" & vbCrLf
                Get #ch, pos, TSnn: pos = pos + 1
                Get #ch, pos, TSdd: pos = pos + 1
                Get #ch, pos, TScc: pos = pos + 1
                Get #ch, pos, TSbb: pos = pos + 1
            ElseIf B3 = &H59 Then    ' key sign
                Get #ch, pos, bBytes: pos = pos + 1
                If bBytes <> 2 Then fout = fout & "Key signature should be stored with 2 bytes" & vbCrLf
                Get #ch, pos, KSsf: pos = pos + 1
                Get #ch, pos, KSmi: pos = pos + 1
            ElseIf B3 = &H7F Then
                Bytes = readVarLen(ch, pos)
                pos = pos + Bytes
            End If
            Get #ch, pos, B1: pos = pos + 1
            Get #ch, pos, B2: pos = pos + 1
        Wend
        TrackInfo = True
        pos = pos - 2
    Else
        pos = 23
        TrackInfo = False
    End If

    DT = readVarLen(ch, pos)
    Get #ch, pos, B2: pos = pos + 1
    If B2 = &HF0 Then
        Bytes = readVarLen(ch, pos)
        If Bytes <> 137 Then fout = fout & "Length of first packet is incorrect"
        ' check sound module type
        For I = 1 To 7
            Get #ch, pos, B: pos = pos + 1
            If Choose(I, &H41, &H10, &H42, &H12, &H48, &H0, &H0) <> B Then MsgBox "This is a dump of another sound module!": GoTo SYXdumpOpenEND
        Next I
        ' read first packet
        For I = 1 To 128
            Get #ch, pos, nB(nBc): pos = pos + 1: nBc = nBc + 1
        Next I
        Get #ch, pos, B1: pos = pos + 1    ' checksum
        Get #ch, pos, B2: pos = pos + 1    ' EOX - F7
        If B2 <> &HF7 Then fout = fout & "Packet EOX error" & vbCrLf: GoTo SYXdumpOpenEND
        ' next packets
        DT = readVarLen(ch, pos)
        Get #ch, pos, B2: pos = pos + 1
        Do While B2 = &HF0
            Bytes = readVarLen(ch, pos)
            For I = 1 To 4: Get #ch, pos, B: pos = pos + 1: Next I
            Get #ch, pos, B3: pos = pos + 1    ' address block &h48/&h49
            If B3 = &H49 And DrumPatch = 0 Then DrumPatch = nBc
            Get #ch, pos, B: pos = pos + 1
            Get #ch, pos, B4: pos = pos + 1    ' packet nr.
            For I = 1 To Bytes - 9    ' nibblized bytes
                Get #ch, pos, nB(nBc): pos = pos + 1: nBc = nBc + 1
            Next I
            Get #ch, pos, B1: pos = pos + 1    'checksum
            Get #ch, pos, B2: pos = pos + 1
            If B2 <> &HF7 Then fout = fout & "EOX error in packet" & CStr(B4) & ", in " & HexByte(B3) & "0000 adres-block" & vbCrLf: GoTo SYXdumpOpenEND
            DT = readVarLen(ch, pos)
            Get #ch, pos, B2: pos = pos + 1
        Loop
    Else
        fout = fout & "No System Exlusive dump found where espected!" & vbCrLf
    End If

    txt = UCase(GetFileTitle(File)) & vbCrLf
    txt = txt & "Division    " & CStr(Division) & " PPQN" & vbCrLf
    txt = txt & "NumBytes    " & CStr(NumBytes) & vbCrLf
    If TrackInfo = True Then
        txt = txt & "Trackname   " & Trackname & vbCrLf
        If Tempo > 0 Then txt = txt & "Tempo       " & CStr(CLng(60000000 / Tempo)) & " BPM" & vbCrLf
        If TSdd > 0 Then
            txt = txt & "Time sign   " & CStr(TSnn) & " / " & CStr(2 ^ TSdd) & " - "
            txt = txt & TScc & " clocks/metr.click - "
            txt = txt & TSbb & " 32nd/quarter " & vbCrLf
        End If
    End If
    If nBc > 0 Then
        DataAccepted = True
        For I = 0 To nBc - 1
            If nB(I) > 16 Then
                fout = fout & "Found bytes using more than 4 bits. All bytes should be nibblized !" & vbCrLf
                DataAccepted = False
                Exit For
            End If
        Next I

        If DataAccepted = True Then
            ' un-nibblize
            For I = 0 To nBc - 1 Step 2
                dmpB(I \ 2) = nB(I) * 16 Or nB(I + 1)
            Next I
            dmpBc = nBc \ 2
            txt = vbCrLf & txt & "Totaal Packet Bytes:" & nBc & " nibblized (" & CStr(dmpBc) & " un-nibblized)" & vbCrLf
            txt = txt & "Only the first " & CStr(DrumPatch \ 2) & " bytes (non-drumpatch values) can be accessed and edited through this program." & vbCrLf
            dmpTrackname = Trackname
            dmpDivision = Division: dmpTempo = Tempo
            dmpTSnn = TSnn: dmpTSdd = TSdd: dmpTScc = TScc: dmpTSbb = TSbb
        Else
            fout = fout & "Data not accepted!" & vbCrLf
        End If
    End If
SYXdumpOpenEND:
    Close ch
    If fout <> "" Then txt = txt & "Errors:" & vbCrLf & fout & vbCrLf
    ret = txt
    SYXdumpOpen = DataAccepted
    Screen.MousePointer = 0
End Function

Public Sub SYXdumpSave(ByVal File As String)
    Dim B          As Byte
    Dim ch         As Long        ' file handle
    Dim I As Long, j As Long      ' counters
    Dim nB(7360)   As Byte        ' nibblized bytes
    Dim nBc        As Long        ' nibble bytes count
    Dim DT         As Byte        ' delta time
    Dim NumBytes   As Long        ' track len
    Dim chksum     As Byte        ' check sum
    Dim dB         As Long        ' dump byte counter

    nBc = 0
    For I = 0 To dmpBc - 1
        nB(nBc) = (dmpB(I) And &HF0) \ 16: nBc = nBc + 1
        nB(nBc) = (dmpB(I) And &HF): nBc = nBc + 1
    Next I

    ch = FreeFile
    Open File For Binary As ch
    Put #ch, 1, "MThd"
    B = 0
    Put #ch, , B: Put #ch, , B: Put #ch, , B: Put #ch, , CByte(6)
    Put #ch, , B: Put #ch, , B    ' formattype=0
    Put #ch, , B: Put #ch, , CByte(1)    ' numtracks
    Put #ch, , CByte((dmpDivision \ 256) And 255)    ' division
    Put #ch, , CByte(dmpDivision And 255)
    Put #ch, , "MTrk"

    ' 29 full part packets + 14 x 2 full drum packets
    ' full packet = 1 DeltaTime byte
    '               1 BOX byte - &HF0
    '               2 following bytes size 137 Big Endian - &H81, &H9
    '               3 modelbytes - &H41, &H10, &H42
    '               1 DT1 (data set 1) byte - &H12
    '               3 address bytes - &H480000 patch all &H490000 drum map
    '             128 packet data bytes
    '               1 checksum byte
    '               1 EOX byte - &HF7
    '           = 141 bytes
    ' 1 non-full part packet of 29 (16 data bytes + 12 other)
    ' 2 non-full drum packet of 36 (24 data bytes + 12 other)*2
    NumBytes = (29 + 28) * (137 + 4) + 29 + 72    ' packet bytes
    If dmpTrackname <> "" Then NumBytes = NumBytes + 4 + Len(dmpTrackname)
    If dmpTempo <> 0 Then NumBytes = NumBytes + 7
    If dmpTSnn <> 0 Then NumBytes = NumBytes + 8
    NumBytes = NumBytes + 5       ' end of track bytes
    Put #ch, , CByte(NumBytes \ (256 ^ 3))
    Put #ch, , CByte((NumBytes \ (256 ^ 2)) And 255)
    Put #ch, , CByte((NumBytes \ 256) And 255)
    Put #ch, , CByte(NumBytes And 255)

    If dmpTrackname <> "" Then
        Put #ch, , CByte(0)
        Put #ch, , CByte(&HFF)
        Put #ch, , CByte(3)
        Put #ch, , CByte(Len(dmpTrackname))
        For I = 1 To Len(dmpTrackname)
            Put #ch, , CByte(Asc(Mid(dmpTrackname, I, 1)))
        Next I
    End If
    If dmpTempo <> 0 Then
        Put #ch, , CByte(0)
        Put #ch, , CByte(&HFF)
        Put #ch, , CByte(&H51)
        Put #ch, , CByte(3)
        Put #ch, , CByte((dmpTempo \ (256 ^ 2)) And 255)
        Put #ch, , CByte((dmpTempo \ 256) And 255)
        Put #ch, , CByte(dmpTempo And 255)
    End If
    If dmpTSnn <> 0 Then
        Put #ch, , CByte(0)
        Put #ch, , CByte(&HFF)
        Put #ch, , CByte(&H58)
        Put #ch, , CByte(4)
        Put #ch, , dmpTSnn
        Put #ch, , dmpTSdd
        Put #ch, , dmpTScc
        Put #ch, , dmpTSbb
    End If

    ' common/parts
    DT = 70
    For I = 0 To &H1C
        Put #ch, , DT
        Put #ch, , CByte(&HF0)
        Put #ch, , CByte(&H81)
        Put #ch, , CByte(&H9)
        For j = 1 To 5
            Put #ch, , CByte(Choose(j, &H41, &H10, &H42, &H12, &H48))
        Next j
        Put #ch, , CByte(I)
        Put #ch, , CByte(0)
        chksum = &H48 + I
        For j = 1 To 128
            Put #ch, , nB(dB)
            chksum = CByte(CInt(chksum + CInt(nB(dB))) Mod 256)
            dB = dB + 1
        Next j
        chksum = -chksum And 127
        Put #ch, , chksum: chksum = 0
        Put #ch, , CByte(&HF7)
    Next I
    Put #ch, , DT
    Put #ch, , CByte(&HF0)
    Put #ch, , CByte(&H19)
    For j = 1 To 5
        Put #ch, , CByte(Choose(j, &H41, &H10, &H42, &H12, &H48))
    Next j
    Put #ch, , CByte(&H1D)
    Put #ch, , CByte(0)
    chksum = &H48 + &H1D
    For j = 1 To 16
        Put #ch, , nB(dB)
        chksum = CByte(CInt(chksum + CInt(nB(dB))) Mod 256)
        dB = dB + 1
    Next j
    chksum = -chksum And 127
    Put #ch, , chksum: chksum = 0
    Put #ch, , CByte(&HF7)

    ' drumpatches
    ' 1st 14 full packets (128 bytes)
    For I = 0 To &HD
        Put #ch, , DT
        Put #ch, , CByte(&HF0)
        Put #ch, , CByte(&H81)
        Put #ch, , CByte(&H9)
        For j = 1 To 5
            Put #ch, , CByte(Choose(j, &H41, &H10, &H42, &H12, &H49))
        Next j
        Put #ch, , CByte(I)
        Put #ch, , CByte(0)
        chksum = &H49 + I
        For j = 1 To 128
            Put #ch, , nB(dB)
            chksum = CByte(CInt(chksum + CInt(nB(dB))) Mod 256)
            dB = dB + 1
        Next j
        chksum = -chksum And 127
        Put #ch, , chksum: chksum = 0
        Put #ch, , CByte(&HF7)
    Next I
    ' a small packets (24bytes)
    Put #ch, , DT
    Put #ch, , CByte(&HF0)
    Put #ch, , CByte(&H21)
    For j = 1 To 5
        Put #ch, , CByte(Choose(j, &H41, &H10, &H42, &H12, &H49))
    Next j
    Put #ch, , CByte(&HE)
    Put #ch, , CByte(0)
    chksum = &H49 + &HE
    For j = 1 To 24
        Put #ch, , nB(dB)
        chksum = CByte(CInt(chksum + CInt(nB(dB))) Mod 256)
        dB = dB + 1
    Next j
    chksum = -chksum And 127
    Put #ch, , chksum: chksum = 0
    Put #ch, , CByte(&HF7)

    ' again 14 full packets
    For I = 0 To &HD
        Put #ch, , DT
        Put #ch, , CByte(&HF0)
        Put #ch, , CByte(&H81)
        Put #ch, , CByte(&H9)
        For j = 1 To 5
            Put #ch, , CByte(Choose(j, &H41, &H10, &H42, &H12, &H49))
        Next j
        Put #ch, , CByte(&H10 + I)
        Put #ch, , CByte(0)
        chksum = &H49 + &H10 + I
        For j = 1 To 128
            Put #ch, , nB(dB)
            chksum = CByte(CInt(chksum + CInt(nB(dB))) Mod 256)
            dB = dB + 1
        Next j
        chksum = -chksum And 127
        Put #ch, , chksum: chksum = 0
        Put #ch, , CByte(&HF7)
    Next I
    ' a small packet
    Put #ch, , DT
    Put #ch, , CByte(&HF0)
    Put #ch, , CByte(&H21)
    For j = 1 To 5
        Put #ch, , CByte(Choose(j, &H41, &H10, &H42, &H12, &H49))
    Next j
    Put #ch, , CByte(&H1E)
    Put #ch, , CByte(0)
    chksum = &H49 + &H1E
    For j = 1 To 24
        Put #ch, , nB(dB)
        chksum = CByte(CInt(chksum + CInt(nB(dB))) Mod 256)
        dB = dB + 1
    Next j
    chksum = -chksum And 127
    Put #ch, , chksum: chksum = 0
    Put #ch, , CByte(&HF7)

    'end track
    Put #ch, , DT
    Put #ch, , CByte(&HFF)
    Put #ch, , CByte(&H2F)
    Put #ch, , CByte(0)
    Put #ch, , CByte(0)
    Close ch
End Sub

' sets the param. we can edit, to their default values ~ soft GS Reset
Public Sub SetDefaultDump()
    Dim iB         As Byte
    Dim I As Long, j As Long
    Dim txt        As String

    For I = 0 To AParCount - 1
        Select Case APar(I).Type
            Case 9                ' master tune
                dmpB(APar(I).ByteOffs) = 4
                dmpB(APar(I).ByteOffs + 1) = 0
            Case 10               ' patchname
                For j = 0 To 15
                    dmpB(j + APar(I).ByteOffs) = Asc(Mid("- Sound Canvas -", j + 1, 1))
                Next j
            Case 11               ' partial reserve
                For j = 0 To 15
                    dmpB(j + APar(I).ByteOffs) = Val(Mid("2622222222000000", j + 1, 1))
                Next j
            Case Else
                dmpB(APar(I).ByteOffs) = APar(I).Default
        End Select
    Next I

    For I = 0 To PParCount - 1
        Select Case PPar(I).ShortName
            Case "CHANNEL"
                For j = 0 To 15
                    iB = Val("&H" & Mid("9012345678ABCDEF", j + 1, 1))
                    dmpB(72 + j * 112 + PPar(I).ByteOffs) = iB
                Next j
            Case "M/P A D"
                dmpB(72 + PPar(I).ByteOffs) = &HB0
                For j = 1 To 15
                    iB = PPar(I).Default
                    dmpB(72 + j * 112 + PPar(I).ByteOffs) = iB
                Next j
            Case Else
                For j = 0 To 15
                    iB = PPar(I).Default
                    dmpB(72 + j * 112 + PPar(I).ByteOffs) = iB
                Next j
        End Select
    Next I

End Sub

' returns the Data Based version of the dump
Public Function strSYXDataB(ByVal Weerg As Long) As String
    Dim iB         As Byte
    Dim BO         As Long
    Dim I As Long, j As Long, K As Long
    Dim txt As String, valHex As String, txt2 As String
    Dim Opschr     As String * 18
    Dim Marked     As Boolean
    Dim ch         As Long
    Dim LnChan As Long, LnAll As Long

    CheckNonDefaults

    'breedte van een kolom
    If Weerg = 0 Then LnChan = 4 Else LnChan = 6
    LnAll = 9

    If CurDmpMidFileTitle <> "" Then txt = UCase(CurDmpMidFileTitle) & vbCrLf & vbCrLf

    ' common
    I = 4: BO = APar(I).ByteOffs
    Opschr = "Patch name ": txt = txt & Opschr
    For j = BO To BO + 15: txt = txt & Chr(dmpB(j)): Next j
    txt = txt & vbCrLf & vbCrLf

    I = 0: BO = APar(I).ByteOffs
    Opschr = "Master tune "
    valHex = HexByte(dmpB(BO)) & " " & HexByte(dmpB(BO + 1))
    txt = txt & Opschr & strSYXDataBVal(valHex, dmpM(BO), APar(I).Type, Weerg, LnAll)
    txt = txt & vbTab & vbTab

    I = 2: BO = APar(I).ByteOffs
    Opschr = "Master key shift "
    valHex = HexByte(dmpB(BO))
    txt = txt & Opschr & strSYXDataBVal(valHex, dmpM(BO), APar(I).Type, Weerg, LnAll)
    txt = txt & vbCrLf

    I = 1: BO = APar(I).ByteOffs
    Opschr = "Master volume"
    valHex = HexByte(dmpB(BO))
    txt = txt & Opschr & strSYXDataBVal(valHex, dmpM(BO), APar(I).Type, Weerg, LnAll)
    txt = txt & vbTab & vbTab

    I = 3: BO = APar(I).ByteOffs
    Opschr = "Master panpot"
    valHex = HexByte(dmpB(BO))
    txt = txt & Opschr & strSYXDataBVal(valHex, dmpM(BO), APar(I).Type, Weerg, LnAll)
    txt = txt & vbCrLf & vbCrLf

    ' reverb chorus
    For I = 13 To AParCount - 1
        BO = APar(I).ByteOffs
        Opschr = APar(I).name
        valHex = HexByte(dmpB(BO))
        txt = txt & Opschr & strSYXDataBVal(valHex, dmpM(BO), APar(I).Type, Weerg, LnAll)
        txt = txt & vbTab & vbTab
        If I < AParCount - 1 Then
            BO = APar(I - 7).ByteOffs
            Opschr = APar(I - 7).name
            valHex = HexByte(dmpB(BO))
            txt = txt & Opschr & strSYXDataBVal(valHex, dmpM(BO), APar(I - 7).Type, Weerg, LnAll)
            txt = txt & vbCrLf
        End If
    Next I
    txt = txt & vbCrLf & vbCrLf

    ' patches
    For j = 0 To 3
        ch = isChannel(j) + 1
        Opschr = getPatchName(j)
        txt2 = txt2 & Format(ch, "00") & " " & Opschr & vbTab
        ch = isChannel(j + 4) + 1
        Opschr = getPatchName(j + 4)
        txt2 = txt2 & Format(ch, "00") & " " & Opschr & vbTab
        ch = isChannel(j + 8) + 1
        Opschr = getPatchName(j + 8)
        txt2 = txt2 & Format(ch, "00") & " " & Opschr & vbTab
        ch = isChannel(j + 12) + 1
        Opschr = getPatchName(j + 12)
        txt2 = txt2 & Format(ch, "00") & " " & Opschr & vbCrLf
    Next j
    txt = txt & txt2 & vbCrLf
    txt2 = ""

    ' parts partial reserve
    I = 5
    txt = txt & IIf(Weerg = 0, APar(I).name, APar(I).ShortName) & " "
    For j = 0 To 15
        BO = j + APar(I).ByteOffs
        iB = dmpB(BO)
        valHex = HexByte(iB)
        txt = txt & strSYXDataBVal(valHex, dmpM(BO), APar(I).Type, Weerg, LnChan)
    Next j
    txt = txt & vbCrLf & vbCrLf

    ' parts headers
    txt = txt & "<B>PART    " & IIf(Weerg = 0, Space(13), "")
    For j = 0 To 15
        txt2 = Choose(j + 1, "10", "1", "2", "3", "4", "5", "6", "7", "8", "9", "11", "12", "13", "14", "15", "16")
        txt = txt & FixStr(txt2, LnChan, " ")
    Next j
    txt = txt & "</B>" & vbCrLf & vbCrLf

    ' parts parameters
    For I = 0 To PParCount - 1
        If Left(PPar(I).ShortName, 2) = "RX" Then
            If Weerg = 0 Then
                txt = txt & IIf(Weerg = 0, PPar(I).name, PPar(I).ShortName) & " "
                For j = 0 To 15
                    iB = dmpB(72 + j * 112 + PPar(I).ByteOffs)
                    valHex = HexByte(iB)
                    Marked = IIf(iB = PPar(I).Default, False, True)
                    txt = txt & strSYXDataBVal(valHex, Marked, PPar(I).Type, Weerg, LnChan)
                Next j
            Else
                For K = 0 To 15
                    txt = txt & "Rx." & Choose(K + 1, "Bend ", "Caf  ", "PrgCh", "CtlCh", "Paf  ", "Note ", "RPN  ", "NRPN ", "Modul", "Volum", "Pan  ", "Expr ", "Hold ", "Port ", "Sost ", "Soft ")
                    For j = 0 To 15
                        valHex = IIf(Rx(K, j) = 0, "Off", "On")
                        txt = txt & strSYXDataBVal(valHex, IIf(Rx(K, j) = 1, False, True), 0, 0, LnChan)
                    Next j
                    txt = txt & vbCrLf
                Next K
                I = I + 1
            End If

        Else
            txt = txt & IIf(Weerg = 0, PPar(I).name, PPar(I).ShortName) & " "
            For j = 0 To 15
                BO = 72 + j * 112 + PPar(I).ByteOffs
                iB = dmpB(BO)
                valHex = HexByte(iB)
                txt = txt & strSYXDataBVal(valHex, dmpM(BO), PPar(I).Type, Weerg, LnChan)
            Next j
        End If
        txt = txt & vbCrLf
    Next I
    strSYXDataB = txt
End Function

' returns the Rx.Switches table in Hex/Dec(Weerg), with fixed length columns (Ln)
Private Function strSYXDataBRx(Weerg As Long, Ln As Long) As String
    Dim I As Long, j As Long
    Dim Rx(16, 16) As Byte
    Dim B          As Byte
    Dim txt As String, valHex As String
    Dim Marked     As Boolean

    If Weerg = 0 Then
        For I = 36 To 37
            txt = txt & IIf(Weerg = 0, PPar(I).name, PPar(I).ShortName) & " "
            For j = 0 To 15
                B = dmpB(72 + j * 112 + PPar(I).ByteOffs)
                valHex = HexByte(B)
                Marked = IIf(B = PPar(I).Default, False, True)
                txt = txt & strSYXDataBVal(valHex, Marked, PPar(I).Type, Weerg, Ln)
            Next j
            txt = txt & vbCrLf
        Next I
        strSYXDataBRx = txt
        Exit Function
    End If
    For j = 0 To 15
        B = dmpB(72 + j * 112 + PPar(36).ByteOffs)
        B = RotateLByte(B)
        For I = 0 To 7
            Rx(I, j) = IIf((B And (2 ^ (7 - I))) = 0, 0, 1)
        Next I
        B = dmpB(72 + j * 112 + PPar(37).ByteOffs)
        B = RotateRByte(B)
        For I = 0 To 7
            Rx(I + 8, j) = IIf((B And (2 ^ (I))) = 0, 0, 1)
        Next I
    Next j
    For I = 0 To 15
        B = Rx(I, 7): Rx(I, 7) = Rx(I, 15): Rx(I, 15) = B
    Next I
    For I = 0 To 15
        txt = txt & "Rx." & Choose(I + 1, "Bend ", "Caf  ", "PrgCh", "CtlCh", "Paf  ", "Note ", "RPN  ", "NRPN ", "Modul", "Volum", "Pan  ", "Expr ", "Hold ", "Port ", "Sost ", "Soft ")
        For j = 0 To 15
            valHex = IIf(Rx(I, j) = 0, "Off", "On")
            txt = txt & strSYXDataBVal(valHex, IIf(Rx(I, j) = 1, False, True), 0, 0, Ln)
        Next j
        txt = txt & vbCrLf
    Next I
    strSYXDataBRx = txt
End Function

' returns a fixed (Ln) string with a Hex or Dec (Weerg) param. value (text)
' according to param type, html-marked or not, right aligned
Private Function strSYXDataBVal(ByVal text As String, Marked As Boolean, ParType As Byte, Weerg As Long, Ln As Long) As String
    Dim txt        As String
    If Weerg = 1 Then txt = isValue(text, ParType) Else txt = text
    Select Case Len(txt)
        Case Is > Ln              ' too long
            txt = Left(txt, Ln)
        Case Ln
        Case Is < Ln              ' too short
            txt = FixStr(txt, Ln, " ")
    End Select
    If Marked Then
        txt = "<SPAN ID='red'>" & txt & "</SPAN>"
    End If
    strSYXDataBVal = txt
End Function

' used in getSC55Param
' splits one row in the param. section
Private Function splitSC55Param(ByVal txt As String, par As SC55PARAM) As Boolean
    If Len(txt) < 3 Then Exit Function
    par.ShortName = Mid(txt, 5, 7)
    par.ByteOffs = Val(Mid(txt, 13, 3))
    par.Address = Mid(txt, 17, 6)
    par.Type = Val(Mid(txt, 23, 3))
    par.Default = Val("&H" & Mid(txt, 27, 2))
    par.Altern = Val(Mid(txt, 30, 3))
    par.name = Mid(txt, 33)
    splitSC55Param = True
End Function

' used in getSC55Param
' splits one row in the param. type section
Private Function splitSC55Type(ByVal txt As String, tpe As SC55PARAMTYPE) As Boolean
    If Len(txt) < 3 Then Exit Function

    tpe.Bytes = Mid(txt, 4, 2)
    tpe.MinHexStr = Mid(txt, 7, 4)
    tpe.MaxHexStr = Mid(txt, 12, 4)
    tpe.MinDecStr = Mid(txt, 17, 6)
    tpe.MaxDecStr = Mid(txt, 24, 6)
    tpe.name = Mid(txt, 31)
    tpe.MinHex = Val("&H" & tpe.MinHexStr)
    tpe.MaxHex = Val("&H" & tpe.MaxHexStr)
    tpe.MinDec = Val(tpe.MinDecStr)
    tpe.MaxDec = Val(tpe.MaxDecStr)
    'Debug.Print tpe.MinHex, tpe.MaxHex, tpe.MinDec, tpe.MaxDec
    splitSC55Type = True
End Function

