Attribute VB_Name = "modSONG"
Option Explicit
Public Type tSlide
    StartPos As Long
    EndPos As Long
    StartFret As Long
    EndFret As Long
End Type

Public Type tguitNOTE
    Fret(1 To 6)   As Long
    F(1 To 6)      As Single
    Slide(1 To 6) As tSlide
End Type



Public GuitEventsAT() As tguitNOTE
Public Cache()     As tguitNOTE

Public SongTime    As Long

Public SongCurPOS  As Long
Public SongCurSTR  As Long
Public SongCurFRET As Long
Public songKX      As Single
Public songKY      As Single

Public SONGSPEED   As Long

Public SongName    As String

Public SelStart    As Long
Public SelEnd      As Long

Public NoteSampler As Long

Public Sub DrawSONG(P As PictureBox, ByVal Offset)
    Dim I          As Long
    Dim KK         As Long
    Dim I2
    Dim Kstep
    Dim Xto        As Long
    Dim X          As Long
    Dim Y          As Long
    Dim pos        As Long
    Dim S          As Long




    P.Cls
    frmOK.PicSongTOP.Cls

    songKX = 20
    songKY = P.Height / 6


    If Offset + 48 < UBound(GuitEventsAT) Then
        Xto = Offset + 48
    Else
        Xto = UBound(GuitEventsAT)
    End If



    Kstep = 1
    I2 = -2

    For X = -Offset * songKX To P.Width Step songKX * Kstep
        If X + I2 >= 0 Then P.Line (X + I2, 0)-(X + I2, P.Height), vbBlack
        frmOK.PicSongTOP.CurrentX = X
        frmOK.PicSongTOP.CurrentY = 1
        If (X + Offset * songKX) / songKX >= SelStart And _
           (X + Offset * songKX) / songKX <= SelEnd Then

            frmOK.PicSongTOP.Print "|||||||||"
        End If
    Next





    Kstep = 1
    I2 = -2
    Do
        For X = -Offset * songKX To P.Width Step songKX * Kstep
            If X + I2 >= 0 Then P.Line (X + I2, 0)-(X + I2, P.Height), vbBlack
            If Kstep > 4 Then

                frmOK.PicSongTOP.CurrentX = X
                frmOK.PicSongTOP.CurrentY = 1
                frmOK.PicSongTOP.Print (X + Offset * songKX) / songKX & " (" & 0.125 * (X + Offset * songKX) / songKX & ")"
            End If
        Next

        For Y = 0 To P.Height Step songKY * Kstep
            P.Line (0, Y)-(P.Width, Y), vbBlack
        Next
        Kstep = Kstep * 2
        I2 = I2 + 1


    Loop While Kstep < 64


    For pos = 0 + Offset To Xto   ' UBound(GuitEventsAT)
        For S = 1 To 6
            Y = (6 - S) * songKY
            X = (pos - Offset) * songKX
            P.CurrentX = X + 2
            P.CurrentY = Y
            If GuitEventsAT(pos).Fret(S) <> -1 Then P.Print GuitEventsAT(pos).Fret(S)
        Next
    Next




End Sub

Public Sub SaveSong(FN As String)
    Dim S          As String
    Dim St         As Long

    Dim E          As Long
    Open FN For Output As 1




    Print #1, frmOK.txtMPC


    For E = 0 To UBound(GuitEventsAT)
        S = ""
        For St = 1 To 6
            S = S & GuitEventsAT(E).Fret(St) & "|" & GuitEventsAT(E).F(St) & "|"

        Next

        Print #1, S
    Next

    Close 1

End Sub

Public Sub LoadSong(FN As String)
    Dim S          As String
    Dim St         As Long
    Dim SP()       As String
    Dim E          As Long



    ReDim SP(0)

    Open FN For Input As 1

    Input #1, S: frmOK.txtMPC = S

    SONGSPEED = 44.1 * Val(frmOK.txtMPC)


    E = 0
    While Not (EOF(1))
        'S = ""
        'For St = 1 To 6
        '    S = S & GuitEventsAT(E).Fret(St) & "|" & GuitEventsAT(E).F(St) & "|"
        'Next


        Input #1, S
        SP = Split(S, "|")
        ReDim Preserve GuitEventsAT(0 To E)
        GuitEventsAT(E).Fret(1) = SP(0)
        GuitEventsAT(E).F(1) = SP(1)
        GuitEventsAT(E).Fret(2) = SP(2)
        GuitEventsAT(E).F(2) = SP(3)
        GuitEventsAT(E).Fret(3) = SP(4)
        GuitEventsAT(E).F(3) = SP(5)
        GuitEventsAT(E).Fret(4) = SP(6)
        GuitEventsAT(E).F(4) = SP(7)
        GuitEventsAT(E).Fret(5) = SP(8)
        GuitEventsAT(E).F(5) = SP(9)
        GuitEventsAT(E).Fret(6) = SP(10)
        GuitEventsAT(E).F(6) = SP(11)


        E = E + 1
    Wend

    Close 1

    frmOK.sSONGPOS.Max = UBound(GuitEventsAT) - 24
    If frmOK.sSONGPOS.Max < 0 Then frmOK.sSONGPOS.Max = 1

    If NoteSampler = 0 Then NoteSampler = 4
    frmOK.sNoteSampler = NoteSampler
End Sub
Public Sub DeleteSong()


    Dim I          As Long

    ReDim GuitEventsAT(0 To 16)

    For I = 0 To UBound(GuitEventsAT)

        GuitEventsAT(I).Fret(1) = -1
        GuitEventsAT(I).Fret(2) = -1
        GuitEventsAT(I).Fret(3) = -1
        GuitEventsAT(I).Fret(4) = -1
        GuitEventsAT(I).Fret(5) = -1
        GuitEventsAT(I).Fret(6) = -1
    Next
    frmOK.sSONGPOS = 0
    frmOK.sSONGPOS.Max = UBound(GuitEventsAT) - 24
    If frmOK.sSONGPOS.Max < 0 Then frmOK.sSONGPOS.Max = 1

    DrawSONG frmOK.PicSong, frmOK.sSONGPOS.value

End Sub

Public Sub RedimSong(BOUND, Optional Evenlower As Boolean)
    Dim I          As Long
    Dim UB         As Long
    UB = UBound(GuitEventsAT)
    BOUND = BOUND + 4


    If BOUND > UB Then

        ReDim Preserve GuitEventsAT(0 To BOUND)

        For I = UB + 1 To BOUND

            GuitEventsAT(I).Fret(1) = -1
            GuitEventsAT(I).Fret(2) = -1
            GuitEventsAT(I).Fret(3) = -1
            GuitEventsAT(I).Fret(4) = -1
            GuitEventsAT(I).Fret(5) = -1
            GuitEventsAT(I).Fret(6) = -1
        Next
        'sSONGPOS = 0
        frmOK.sSONGPOS.Max = UBound(GuitEventsAT) - 24
        If frmOK.sSONGPOS.Max < 0 Then frmOK.sSONGPOS.Max = 1

        DrawSONG frmOK.PicSong, frmOK.sSONGPOS.value
    Else
        If Evenlower Then ReDim Preserve GuitEventsAT(0 To BOUND)
    End If

End Sub

Public Sub SongCopy()
    Dim I          As Long

    ReDim Cache(1 To SelEnd - SelStart + 1)


    For I = SelStart To SelEnd
        Cache(I - SelStart + 1) = GuitEventsAT(I)
    Next

End Sub

Public Sub SongPaste()


    Dim I          As Long



    I = UBound(Cache)
    If I + SelStart > UBound(GuitEventsAT) Then RedimSong (SelStart + I)


    For I = SelStart To SelStart + UBound(Cache) - 1
        GuitEventsAT(I) = Cache(I - SelStart + 1)
    Next

End Sub


Public Sub SongPrePlay()
    Dim I          As Long
    Dim S          As Long
    Dim lastFret(1 To 6) As Long
    Dim Slide(1 To 6) As Boolean
Dim SlideStartPOS(1 To 6) As Long

    For I = 0 To UBound(GuitEventsAT)
        For S = 1 To 6
            GuitEventsAT(I).Slide(S).StartFret = 0
            GuitEventsAT(I).Slide(S).EndFret = 0
            GuitEventsAT(I).Slide(S).StartPos = 0
            GuitEventsAT(I).Slide(S).EndPos = 0
        Next
    Next

    For I = 0 To UBound(GuitEventsAT)
        For S = 1 To 6
            If Slide(S) = False Then
                If GuitEventsAT(I).Fret(S) >= 0 Then
                                
                    lastFret(S) = GuitEventsAT(I).Fret(S): 'SlideStartPOS(S) = I
                    'GuitEventsAT(I).Slide(S).StartFret = GuitEventsAT(I).Fret(S)
                    'If GuitEventsAT(I + 1).Fret(S) = -2 Then GuitEventsAT(I).Slide(S).StartPos = I

                End If
                
                If GuitEventsAT(I).Fret(S) = -2 Then

                    Slide(S) = True
                    
                    SlideStartPOS(S) = I - 1
                    GuitEventsAT(I - 1).Slide(S).StartPos = I - 1
                    GuitEventsAT(I - 1).Slide(S).StartFret = lastFret(S)

                    
                End If
            Else
                If GuitEventsAT(I).Fret(S) >= 0 Then
                                       
                                       
                    GuitEventsAT(SlideStartPOS(S)).Slide(S).EndFret = GuitEventsAT(I).Fret(S)
                    GuitEventsAT(SlideStartPOS(S)).Slide(S).EndPos = I

                                      
                    Slide(S) = False
              
                End If

            End If

        Next

    Next

End Sub

