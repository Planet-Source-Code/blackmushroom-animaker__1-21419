Attribute VB_Name = "modAnimaker"
Option Explicit

Public msindex As Integer   'for editing sprite-values (passed to frmSprite): not very nice...

Public workpath As String
Public bmpname As String
Public bmpsizeX As Integer
Public bmpsizeY As Integer

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal animX As Long, ByVal animY As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020   'Copies the source over the destination


Public Sub init()
    workpath = App.path + "\"
    
End Sub

Public Sub readd_spritelist()
Dim a As Integer
Dim b As String
    With frmAnimaker
        .lstSprites.Clear
        For a = 0 To spritecount - 1
            b = Str(a) + ":   X" + Trim(Str(sprites(a).X)) + "  Y" + Trim(Str(sprites(a).Y)) + "  W" + Trim(Str(sprites(a).W)) + "  H" + Trim(Str(sprites(a).H))
            .lstSprites.AddItem (b)
        Next a
    
    End With
End Sub

Public Sub readd_anilist()
Dim a As Integer
Dim b As String
    With frmAnimaker
        .lstAni.Clear
        For a = 1 To anicount
            b = Str(a) + ": " + ani(a - 1).name
            .lstAni.AddItem (b)
        Next a
    
    End With
End Sub


Public Sub display_anidesc()
Dim a As Integer
Dim b As String
Dim c As Integer
    
    With frmAnimaker
        c = .lstAni.ListIndex
        .lstAnidesc.Clear
        For a = 1 To ani(c).spritecount
            b = Str(a) + ": Sprite " + Str(ani(c).spritelst(a - 1))
            .lstAnidesc.AddItem (b)
        Next a
    
    End With

End Sub

Public Sub save_ani(path As String)
Dim qnr As Integer
Dim a As Integer
Dim b As Integer
    qnr = FreeFile
    Open path For Output As #qnr
        Print #qnr, "Animationfile generated for " + bmpname + " with Animaker"
        Print #qnr, spritecount
        For a = 0 To spritecount
            Print #qnr, sprites(a).X
            Print #qnr, sprites(a).Y
            Print #qnr, sprites(a).W
            Print #qnr, sprites(a).H
        Next a
        Print #qnr, anicount
        For a = 0 To anicount
            Print #qnr, ani(a).spritecount
            Print #qnr, ani(a).speed
            Write #qnr, ani(a).name
            For b = 0 To ani(a).spritecount
                Print #qnr, ani(a).spritelst(b)
            Next b
        Next a
    Close #qnr

End Sub

