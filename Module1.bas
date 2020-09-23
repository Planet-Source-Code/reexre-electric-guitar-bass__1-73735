Attribute VB_Name = "Module1"
Public Type tGuitEVENT
    Str            As Long
    FRET           As Long
End Type


Public GE()        As tGuitEVENT



Public Sub INITGE()

    ReDim GE(80)

    GE(1).Str = 1
    GE(1).FRET = 0
    GE(2).Str = 1
    GE(2).FRET = 0
    GE(3).Str = 1
    GE(3).FRET = 5
    GE(4).Str = 1
    GE(4).FRET = 4

    GE(5).Str = 1
    GE(5).FRET = 0
    GE(6).Str = 1
    GE(6).FRET = 0
    GE(7).Str = 1
    GE(7).FRET = 5
    GE(8).Str = 1
    GE(8).FRET = 7

    GE(9).Str = 1
    GE(9).FRET = 0
    GE(10).Str = 1
    GE(10).FRET = 0
    GE(11).Str = 1
    GE(11).FRET = 5
    GE(12).Str = 1
    GE(12).FRET = 4

    GE(13).Str = 1
    GE(13).FRET = 5
    GE(14).Str = 1
    GE(14).FRET = 4
    GE(15).Str = 1
    GE(15).FRET = 0
    GE(16).Str = 1
    GE(16).FRET = 5

    GE(17).Str = 6
    GE(17).FRET = 0
    GE(18).Str = 6
    GE(18).FRET = 0



End Sub
