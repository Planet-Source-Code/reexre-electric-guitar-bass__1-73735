Attribute VB_Name = "modGoertzel"
'http://www.vbforums.com/showthread.php?t=388815

Private Function Goertzel( _
         sngData() As Integer, _
         ByVal N As Long, _
         ByVal freq As Single, _
         ByVal sampr As Long _
         ) As Single

    Dim Skn        As Single
    Dim Skn1       As Single
    Dim Skn2       As Single
    Dim c          As Single
    Dim c2         As Single
    Dim I          As Long
      
    ''''''''''''''''''''''''
    'sgnData type originally was single
    ''''''''''''''''''''''''
    c = 2 * Pi * freq / sampr
    c2 = Cos(c)

    For I = 0 To N - 1
        Skn2 = Skn1
        Skn1 = Skn
        Skn = 2 * c2 * Skn1 - Skn2 + sngData(I)
    Next

    Goertzel = Skn - Exp(-c) * Skn1
End Function

Private Function power(ByVal value As Single) As Single
    power = 20 * Log(Abs(value)) / Log(10)
End Function


'Usage: (returns dB of 8000 Hz in the signal at 44100 samples/s)
Public Function GetdB(intSignal() As Integer, HzFrequence, SignalFreq)
GetdB = power(Goertzel(intSignal, UBound(intSignal) + 1, HzFrequence, SignalFreq))

End Function


Public Sub DetectSoundNote()
Dim I As Long
Dim Max As Single
Dim db As Single
Dim N As String

For I = 0 To 119

db = GetdB(WF, freq(I), 44100)

If db > Max Then Max = db: N = FreqToNote(I) & "  (" & freq(I) & "Hz)"


Next
MsgBox N

End Sub
