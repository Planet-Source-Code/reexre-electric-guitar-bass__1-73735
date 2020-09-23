Attribute VB_Name = "mDRIVERS"
' this module containes general functions

Option Explicit
Public OK          As Boolean
Public Cancel      As Boolean
Public CurHTMLfile As String

' API
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40
Public Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020   ' (DWORD) dest = source


Public Sub BevelPic(PIC As PictureBox, _
                    ByVal X1 As Long, ByVal Y1 As Long, _
                    ByVal X2 As Long, ByVal Y2 As Long, _
                    ByVal inset As Boolean)
    Dim above As Long, under As Long

    If inset = False Then
        above = vb3DHighlight
        under = vbButtonShadow
    Else
        above = vbButtonShadow
        under = vb3DHighlight
    End If
    X2 = X2 - 1
    Y2 = Y2 - 1
    PIC.Line (X1, Y1)-(X2, Y2), vb3DDKShadow, B
    X1 = X1 + 1
    Y1 = Y1 + 1
    X2 = X2 - 1
    Y2 = Y2 - 1
    PIC.Line (X1, Y1)-(X2, Y1), above
    PIC.Line (X1, Y1)-(X1, Y2), above
    PIC.Line (X2, Y1)-(X2, Y2), under
    PIC.Line (X1, Y2)-(X2, Y2), under

End Sub

' returns a binary string of a byte
Public Function BinStr(ByVal B As Byte) As String
    Dim I As Long, txt As String
    For I = 0 To 7
        txt = txt & IIf((B And (2 ^ (7 - I))) = 0, "0", "1")
    Next I
    BinStr = txt
End Function

' returns the value of a binary string
Public Function BinValue(ByVal BinStr As String) As Variant
    Dim V          As Variant
    Dim I As Long, L As Long
    L = Len(BinStr)
    For I = 1 To L
        V = V + 2 ^ (L - I) * Val(Mid(BinStr, I, 1))
    Next I
    BinValue = V
End Function

' Hourglass mousepointer is not visible above a browser window
' this routine makes busy-ness visible
Public Sub Busy(frm As Form, truefalse As Boolean)
    Static BackColor As Long
    If truefalse = True Then
        Screen.MousePointer = vbHourglass
        BackColor = frm.BackColor
        frm.BackColor = RGB(200, 200, 200)
    Else
        Screen.MousePointer = vbDefault
        frm.BackColor = BackColor
    End If
End Sub
Public Function Convert(ByVal Value As Single, _
                        ByVal FromMin As Single, ByVal FromMax As Single, _
                        ByVal ToMin As Single, ByVal ToMax As Single) As Single

    Dim F As Single, ToAdd As Long
    If (ToMax - ToMin) = (FromMax - FromMin) Then
        F = 1
    Else
        If ToMin < 0 Then ToAdd = 1 Else ToAdd = 0
        F = (ToMax - ToMin) / (FromMax - FromMin + ToAdd)
    End If

    Convert = ToMin + (Value - FromMin) * F
End Function


' transforms a txt file into a html file
' isn't realy necessary here
Public Function File2html(ByVal File As String, ByVal Title As String) As String
    Dim ch As Long, I As Long
    Dim Regel      As String
    Dim txt        As String
    Dim ipb        As String
    Dim Kolom() As Variant, KolomH() As Variant, K As Long, aK As Long
    Dim tW         As Long

    txt = txt & GetHeader(File, Title)
    txt = txt & "    " & File & vbCrLf
    ch = FreeFile
    Open File For Input As ch
    Line Input #ch, Regel
    txt = txt & "    <PRE>" & vbCrLf
    txt = txt & "    " & Regel & vbCrLf
    While Not EOF(ch)
        Line Input #ch, Regel
        txt = txt & "    " & Regel & vbCrLf
    Wend
    txt = txt & "    </PRE>" & vbCrLf

    Close ch
    txt = txt & GetFooter()

    File2html = txt
End Function

' sets a value to a fixed length string, right aligned
' used with proportional fonts
Function FixStr(ByVal S As Variant, ByVal L As Long, ByVal F As String) As String
    Dim txt        As String
    txt = CStr(S)
    While Len(txt) < L
        txt = Left(F, 1) & txt
    Wend
    FixStr = txt
End Function

Public Function getBit(ByVal Value As Byte, ByVal BitNo As Long) As Long
    getBit = IIf((Value And (2 ^ BitNo)) = 0, 0, 1)
End Function


' chr(0) is not the end of the string, so generate your own
' string representation of the commandstring for displaying it
Public Function getComStrStr(ByVal CommandStr As String) As String
    Dim txt As String, K As String * 1
    Dim I          As Long
    For I = 1 To Len(CommandStr)
        K = Mid(CommandStr, I, 1)
        If Asc(K) < 32 Then
            txt = txt & Chr(128)
        Else
            txt = txt & K
        End If
    Next I
    getComStrStr = txt
End Function

' decimal representation of the commandstring
Public Function getComStrDec(ByVal CommandStr As String) As String
    Dim txt        As String
    Dim I          As Long
    For I = 1 To Len(CommandStr)
        txt = txt & Format(Asc(Mid(CommandStr, I, 1)), "000") & " "
    Next I
    getComStrDec = txt
End Function


' hexadecimal representation of the commandstring
Public Function getComStrHex(ByVal CommandStr As String) As String
    Dim txt        As String
    Dim I          As Long
    For I = 1 To Len(CommandStr)
        txt = txt & HexByte(Asc(Mid(CommandStr, I, 1))) & " "
    Next I
    getComStrHex = txt
End Function

Public Function setBit(ByVal Value As Byte, ByVal BitNo As Long, ByVal OnOff As Boolean) As Byte
    Value = Value Or (2 ^ BitNo)
    If OnOff = False Then Value = Value Xor (2 ^ BitNo)
    setBit = Value
End Function

Public Function GetFileTitle(ByVal FullFilename As Variant) As Variant
    Dim I          As Integer
    Dim nm         As String
    Dim K          As String * 1

    For I = Len(FullFilename) To 1 Step -1
        K = Mid(FullFilename, I, 1)
        If K <> "\" Then nm = K + nm Else Exit For
    Next I
    GetFileTitle = nm
End Function

' makes a html footer
Public Function GetFooter() As String
    Dim txt        As String
    txt = txt & "  </BODY>" & vbCrLf
    txt = txt & "</HTML>" & vbCrLf
    GetFooter = txt
End Function

' makes a html header
Public Function GetHeader(ByVal File As String, Optional ByVal Title As String) As String
    Dim txt        As String

    If Title = "" Then Title = GetFileTitle(File)
    txt = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>" & vbCrLf
    txt = txt & "<HTML>" & vbCrLf
    txt = txt & "  <HEAD>" & vbCrLf
    txt = txt & "    <TITLE>" & Title & "</TITLE>" & vbCrLf
    txt = txt & "    <STYLE TYPE='text/css'>" & vbCrLf
    txt = txt & "        PRE   {font-size:8pt;}" & vbCrLf
    txt = txt & "        #red    {color:red;}" & vbCrLf
    txt = txt & "        #green   {color:green;}" & vbCrLf
    txt = txt & "        BODY     {background-color:#E8E8E8;}" & vbCrLf
    txt = txt & "    </STYLE>" & vbCrLf
    txt = txt & "  </HEAD>" & vbCrLf
    txt = txt & "  <BODY>" & vbCrLf
    GetHeader = txt
End Function

' hex string of a byte, fixed to 2 digits
Public Function HexByte(ByVal B As Byte) As String
    Dim txt        As String
    txt = Hex(B)
    If Len(txt) = 1 Then txt = "0" & txt
    HexByte = txt
End Function

' duration in 1000/sec
Public Sub Pauze(ByVal Duration As Long)
    Dim MakePiano
    MakePiano = Timer
    While Timer - MakePiano < Duration / 1000: DoEvents: Wend
End Sub

' read Big Endian - variable length variabel
' ch=filehandle, pos=position in file
Public Function readVarLen(ByVal ch As Long, pos As Long) As Long
    Dim Value      As Long
    Dim C          As Byte

    Get #ch, pos, C: pos = pos + 1
    Value = C
    If (Value And &H80) <> 0 Then
        Value = Value And &H7F
        Do
            Value = Value * 128
            Get #ch, pos, C: pos = pos + 1
            C = C And &H7F
            Value = Value + C
        Loop While (C And &H80) <> 0
    End If
    readVarLen = Value
End Function

' binary rotate Left
Public Function RotateLByte(ByVal B As Byte) As Byte
    RotateLByte = ((B * 2) Mod 256) + IIf((B And 128) = 0, 0, 1)
End Function

' binary rotate Right
Public Function RotateRByte(ByVal B As Byte) As Byte
    RotateRByte = (B \ 2) Or IIf((B And 1) = 0, 0, 128)
End Function

' write Big Endian - I don't use it! Can be removed
Public Sub WriteVarLen(ByVal ch As Long, ByVal Value As Long)
    Dim buffer     As Long
    buffer = Value And &H7F
    While Value \ 128 > 0
        Value = Value \ 128
        buffer = buffer * 256
        buffer = buffer Or ((Value And &H7F) Or &H80)
    Wend
    Do
        Put #ch, , CByte(buffer And 255)    ': Pos = Pos + 1
        If (buffer And &H80) Then
            buffer = buffer \ 256
        Else
            Exit Do
        End If
    Loop
End Sub
