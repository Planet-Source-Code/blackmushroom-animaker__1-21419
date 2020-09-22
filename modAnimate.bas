Attribute VB_Name = "modAnimate"
'   *************************
'   THIS IS THE "ENGINE" !!!!
'   *************************
'
'  - Insert just this module into your game !
'
'  - Call ani_init(path to an ani-file) at the beginning
'  - Call animate(index or animation-name) in the main loop
'  - Then call drawAnimatedSprite(x,y,index or animation-name ) to draw the animated sprite
'
'   NOTE: You have to load the bitmaps by yourself !
'
'   The rest of the project is just to fill the data-structure with information...
'   You don't need it for your games.
'   You don't have to care about the code, except the spots marked
'   for modification (with ########). There you have to add own code for your game.
'   Use bitblt or DirectDraw or anything else in the draw...Sprite-routines
'   I didn't do much commentation in the rest, as it is not needed for games.
'   (The coding of the other project-files isn't very nice, lots of confusing a's and b's...)
'   (c) 2001 by BlackMushroom

Option Explicit

'#############
'Modify these to extend the Animaker
Public Const NUMBER_OF_SPRITES = 1000
Public Const NUMBER_OF_ANIMATIONS = 1000
Public Const NUMBER_OF_SPRITES_IN_ANIMATION = 10
'#############

Type spritedesc
    X As Integer    'x on source-bmp
    Y As Integer    'y on source-bmp
    H As Integer    'height
    W As Integer    'width
End Type

Type animationdesc
    spritelst(NUMBER_OF_SPRITES_IN_ANIMATION) As Integer    'here max 10 frames, add more if needed!
    spritecount As Integer
    speed As Integer    '0 means every mainloop; >0 means every 'speed-1' mainloops
    delay As Integer    ''delay' counts from 'speed' to 0, then pointer is increased (to next sprite)
    pointer As Integer  'points at the active sprite in the list
    name As String      'the alternative animation-name
End Type

Public ani() As animationdesc   'add more definitions for other bitmaps/ani-files !
Public anicount As Integer      '...and more counters
Public sprites() As spritedesc
Public spritecount As Integer


Public Sub load_ani(path As String)
Dim qnr As Integer
Dim a As Integer
Dim b As Integer
Dim dummy As String
    qnr = FreeFile
    Open path For Input As #qnr
        Input #qnr, dummy
        Input #qnr, spritecount
        ReDim sprites(NUMBER_OF_SPRITES) As spritedesc  'replace by spritecount in games!!!
        For a = 0 To spritecount                        'you won't need 1000 empty sprites
            Input #qnr, sprites(a).X
            Input #qnr, sprites(a).Y
            Input #qnr, sprites(a).W
            Input #qnr, sprites(a).H
        Next a
        Input #qnr, anicount
        ReDim ani(NUMBER_OF_ANIMATIONS) As animationdesc    'replace by anicount in games!!!
        For a = 0 To anicount
            Input #qnr, ani(a).spritecount
            Input #qnr, ani(a).speed
            Input #qnr, ani(a).name
            For b = 0 To ani(a).spritecount
                Input #qnr, ani(a).spritelst(b)
            Next b
        Next a
    Close #qnr
End Sub

Public Sub ani_init(Optional path_anifile As String)
'use to load an *.ani-file
    If path_anifile <> "" Then
        Call load_ani(path_anifile)
    Else
        'for this editor: 1000 empty sprites and animations
        ReDim sprites(NUMBER_OF_SPRITES) As spritedesc
        ReDim ani(NUMBER_OF_ANIMATIONS) As animationdesc
    End If
End Sub

Public Sub animate(Optional aniindex As Integer, Optional aniname As String)
'use this to animate a sprite (in your main-loop)
Dim a As Integer
    
    If aniname <> "" Then   'find index to aniname
        For a = 0 To anicount
            If ani(a).name = aniname Then
                aniindex = a
                Exit For
            End If
        Next a
    End If
    'Animate!
    If ani(aniindex).delay > 0 Then
        ani(aniindex).delay = ani(aniindex).delay - 1   'count down delay
    Else
        ani(aniindex).delay = ani(aniindex).speed
        ani(aniindex).pointer = ani(aniindex).pointer + 1
        If ani(aniindex).pointer >= ani(aniindex).spritecount Then ani(aniindex).pointer = 0
    End If
End Sub


Public Sub drawAnimatedSprite(X As Integer, Y As Integer, Optional animation_index As Integer, Optional aniname As String)
'use this to display an animated sprite
Dim a As Integer
    If aniname <> "" Then   'find index to aniname
        For a = 0 To anicount
            If ani(a).name = aniname Then
                animation_index = a
                Exit For
            End If
        Next a
    End If



Dim Sindex As Integer
Dim sourceX As Integer
Dim sourceY As Integer
Dim sourceW As Integer
Dim sourceH As Integer

    Sindex = ani(animation_index).spritelst(ani(animation_index).pointer)   'what animation-frame
    
    
    sourceX = sprites(Sindex).X     'use theese values in your graphics-routine
    sourceY = sprites(Sindex).Y     'to get the position of the sprite on the bitmap
    sourceW = sprites(Sindex).W
    sourceH = sprites(Sindex).H
    
    
    '##########
    '  MODIFY (here's just some lame blitting for instant viewing...)
Dim ret As Long
    ret = BitBlt(frmAnimaker.picAni.hDC, X, Y, sourceW, sourceH, frmAnimaker.picBmp.hDC, sourceX, sourceY, SRCCOPY)
    frmAnimaker.picAni.Refresh
    frmAnimaker.lbsprite.Caption = Str(Sindex)
    '##########
    
End Sub

Public Sub drawSprite(X As Integer, Y As Integer, sprite_index As Integer)
'Use this routine to draw any not animated indexed sprite
Dim sourceX As Integer
Dim sourceY As Integer
Dim sourceW As Integer
Dim sourceH As Integer

    sourceX = sprites(sprite_index).X     'use theese values in your graphics-routine
    sourceY = sprites(sprite_index).Y     'to get the position of the sprite on the bitmap
    sourceW = sprites(sprite_index).W
    sourceH = sprites(sprite_index).H

    '##########
    '  MODIFY (same as above..., not used here)
    'Dim ret As Long
    'ret = BitBlt(somewhere.hDC, X, Y, sourceW, sourceH, somesource.hDC, sourceX, sourceY, SRCCOPY)
    'somewhere.Refresh
    '##########

End Sub

Public Function getSprite(animation_index As Integer) As spritedesc     '!!! Returns a spritedesc-variable!!!
'returns the currently active sprite from a given animation_index
Dim Sindex As Integer
    Sindex = ani(animation_index).spritelst(ani(animation_index).pointer)
    getSprite.X = sprites(Sindex).X     'fill the return-variable with the sprite-values...
    getSprite.Y = sprites(Sindex).Y
    getSprite.W = sprites(Sindex).W
    getSprite.H = sprites(Sindex).H


End Function
