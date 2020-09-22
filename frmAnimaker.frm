VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAnimaker 
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10665
   Icon            =   "frmAnimaker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10665
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add to Ani"
      Height          =   255
      Left            =   9600
      TabIndex        =   12
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton btnChange 
      Caption         =   "Change"
      Height          =   255
      Left            =   9600
      TabIndex        =   33
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton btnDeleteS 
      Caption         =   "Delete"
      Height          =   255
      Left            =   9600
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton btnNewS 
      Caption         =   "Insert"
      Height          =   255
      Left            =   9600
      TabIndex        =   32
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   6960
      TabIndex        =   31
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtPath 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "<none>"
      Top             =   960
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   9720
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   8160
      TabIndex        =   27
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton btnAniDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   9600
      TabIndex        =   21
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load Ani"
      Height          =   375
      Left            =   9360
      TabIndex        =   20
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtspeed 
      Height          =   285
      Left            =   9720
      TabIndex        =   19
      Text            =   "0"
      Top             =   3600
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10200
      Top             =   4560
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "New"
      Height          =   255
      Left            =   9600
      TabIndex        =   17
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save Ani"
      Height          =   375
      Left            =   9360
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   255
      Left            =   8040
      TabIndex        =   15
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   255
      Left            =   8040
      TabIndex        =   14
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton btnDeleteA 
      Caption         =   "Delete"
      Height          =   255
      Left            =   9600
      TabIndex        =   13
      Top             =   4200
      Width           =   975
   End
   Begin VB.ListBox lstAnidesc 
      Height          =   840
      Left            =   6840
      TabIndex        =   11
      Top             =   4200
      Width           =   2775
   End
   Begin VB.PictureBox picAni 
      AutoRedraw      =   -1  'True
      CausesValidation=   0   'False
      FillStyle       =   0  'Ausgefüllt
      Height          =   1215
      Left            =   6840
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   10
      Top             =   5400
      Width           =   1215
   End
   Begin VB.ListBox lstAni 
      Height          =   1035
      Left            =   6840
      TabIndex        =   9
      ToolTipText     =   "Click to select..."
      Top             =   2880
      Width           =   2775
   End
   Begin VB.CommandButton btnGrid 
      Caption         =   "Get Sprites"
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   6360
      Width           =   255
   End
   Begin VB.PictureBox picback 
      BorderStyle     =   0  'Kein
      Height          =   6495
      Left            =   120
      ScaleHeight     =   433
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   441
      TabIndex        =   2
      Top             =   120
      Width           =   6615
      Begin VB.HScrollBar HScroll1 
         CausesValidation=   0   'False
         Height          =   255
         LargeChange     =   100
         Left            =   0
         Max             =   0
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   6240
         Width           =   6375
      End
      Begin VB.VScrollBar VScroll1 
         CausesValidation=   0   'False
         Height          =   6255
         LargeChange     =   100
         Left            =   6360
         Max             =   0
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picBmp 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   6255
         Left            =   0
         MousePointer    =   2  'Kreuz
         ScaleHeight     =   413
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   421
         TabIndex        =   5
         Top             =   0
         Width           =   6375
         Begin VB.Shape shpSprite 
            BorderColor     =   &H00000000&
            BorderWidth     =   3
            DrawMode        =   6  'Stift und inverse Anzeige maskieren
            Height          =   1215
            Left            =   960
            Top             =   1440
            Visible         =   0   'False
            Width           =   2055
         End
      End
   End
   Begin VB.CommandButton btnLoadBmp 
      Caption         =   "Load BMP"
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstSprites 
      Height          =   1035
      Left            =   6840
      TabIndex        =   0
      ToolTipText     =   "Click to select, doubleclick to change..."
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label7 
      Caption         =   "Preview:"
      Height          =   255
      Left            =   6840
      TabIndex        =   30
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Bitmap:"
      Height          =   255
      Left            =   6840
      TabIndex        =   29
      Top             =   960
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   9240
      MousePointer    =   14  'Pfeil und Fragezeichen
      Picture         =   "frmAnimaker.frx":030A
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Zentriert
      Caption         =   "Sprite:"
      Height          =   195
      Left            =   8280
      TabIndex        =   26
      Top             =   6000
      Width           =   570
   End
   Begin VB.Label Label5 
      Caption         =   "Sprites in this Animation:"
      Height          =   255
      Left            =   6840
      TabIndex        =   25
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Speed:"
      Height          =   255
      Left            =   9720
      TabIndex        =   24
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Animations:"
      Height          =   255
      Left            =   6840
      TabIndex        =   23
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Sprites:"
      Height          =   255
      Left            =   6840
      TabIndex        =   22
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lbsprite 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "0"
      Height          =   255
      Left            =   8280
      TabIndex        =   18
      Top             =   6240
      Width           =   615
   End
End
Attribute VB_Name = "frmAnimaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAdd_Click()
Dim a As Integer
Dim c As Integer
    If lstAni.ListIndex <> -1 Then
        c = lstAni.ListIndex
        a = ani(c).spritecount
        If a < NUMBER_OF_SPRITES_IN_ANIMATION Then
            ani(c).spritelst(a) = lstSprites.ListIndex
            ani(c).spritecount = ani(c).spritecount + 1
            Call display_anidesc
        Else
            MsgBox "Max. " + Str(NUMBER_OF_SPRITES_IN_ANIMATION) + " sprites per animation. Change the Type-Def for more !", vbOKOnly, "Error"
        End If
    End If
End Sub

Private Sub btnAniDelete_Click()
Dim a As Integer
    If lstAni.ListIndex <> -1 Then
        For a = lstAni.ListIndex + 1 To anicount + 1
            ani(a - 1) = ani(a)
        Next a
        anicount = anicount - 1
        Call readd_anilist
    End If
End Sub

Private Sub btnChange_Click()
    If lstSprites.ListIndex <> -1 Then
        msindex = lstSprites.ListIndex
        Load frmSprite
        frmSprite.Show 1
    End If
End Sub

Private Sub btnDeleteA_Click()
Dim a As Integer
Dim b As Integer
    If lstAnidesc.ListIndex <> -1 And lstAni.ListIndex <> -1 Then
        a = lstAnidesc.ListIndex
        For b = a + 1 To lstAnidesc.ListCount
            ani(lstAni.ListIndex).spritelst(b - 1) = ani(lstAni.ListIndex).spritelst(b)
        Next b
        ani(lstAni.ListIndex).spritecount = ani(lstAni.ListIndex).spritecount - 1
        display_anidesc
    End If
End Sub

Private Sub btnDeleteS_Click()
Dim a As Integer
    If lstSprites.ListIndex <> -1 Then
        For a = lstSprites.ListIndex To lstSprites.ListCount
            sprites(a) = sprites(a + 1)
        Next a
        spritecount = spritecount - 1
        Call readd_spritelist
    End If
End Sub

Private Sub btnGrid_Click()
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer

Dim frame As String
    frame = InputBox("Frames? (Y/N)", "Frames", "Y")
    a = Val(InputBox("Sizex=", "SizeX", "49"))
    b = Val(InputBox("Sizey=", "SizeY", "49"))
    spritecount = 0
    If a <> 0 And b <> 0 Then
        For c = 0 To Int(bmpsizeY / b) - 1
            For d = 0 To Int(bmpsizeX / a) - 1
                If frame = "Y" Then
                    sprites(spritecount).X = d * a
                    sprites(spritecount).Y = c * b
                    sprites(spritecount).W = a - 1
                    sprites(spritecount).H = b - 1
                Else
                    sprites(spritecount).X = d * a
                    sprites(spritecount).Y = c * b
                    sprites(spritecount).W = a
                    sprites(spritecount).H = b
                End If
                spritecount = spritecount + 1
            
            Next d
        Next c
        readd_spritelist
    
    End If
End Sub

Private Sub btnHelp_Click()
Dim a As String
    a = a + "1. Open a bmp containing your sprites." + vbCrLf
    a = a + "2. Click on 'Get Sprites' and specify" + vbCrLf
    a = a + "   the width and height of one sprite." + vbCrLf
    a = a + "   Or insert and change the sprite-values" + vbCrLf
    a = a + "   manually for different sprite sizes." + vbCrLf
    a = a + "3. Test the frames by clicking on the sprites." + vbCrLf
    a = a + "4. Select 'New' form 'animations' and name" + vbCrLf
    a = a + "   the animation. Select it from the list." + vbCrLf
    a = a + "5. Rightclick on the bmp to add a sprite to" + vbCrLf
    a = a + "   the animation." + vbCrLf
    a = a + "6. Select 'Play' to preview the animation!" + vbCrLf
    a = a + "" + vbCrLf
    MsgBox a, vbOKOnly, "A little help..."
End Sub

Private Sub btnLoad_Click()
Dim name As String
    With comDlg
        .DialogTitle = "Load Animation file"
        .Filter = "Animation file (*.ani)|*.ani"
        .InitDir = workpath
        .FileName = ""
        .ShowOpen
        name = .FileName
    End With
    If name <> "" Then
        Call load_ani(name)
        Call readd_spritelist
        Call readd_anilist
    End If

End Sub

Private Sub btnLoadBmp_Click()
Dim a As Integer
    
    With comDlg
        .DialogTitle = "Load Bitmap"
        .Filter = "All Images (*.bmp;*.jpg;*.jpeg;*.gif)|*.bmp;*.jpg;*.jpeg;*.gif)|All Files (*.*)|*.*"
        .InitDir = workpath
        .ShowOpen
        bmpname = .FileName
    End With
    
    If bmpname <> "" Then
        With frmAnimaker
            .picBmp.Picture = LoadPicture(bmpname)
            .picBmp.Refresh
            a = .picBmp.Width - 425
            If a > 0 Then
                .HScroll1.Max = a
            Else
                .HScroll1.Max = 0
            End If
            a = .picBmp.Height - 425
            If a > 0 Then
                .VScroll1.Max = a
            Else
                .VScroll1.Max = 0
            End If
            bmpsizeX = .picBmp.Width
            bmpsizeY = .picBmp.Height
            txtPath.Text = bmpname
        End With
    End If
End Sub

Private Sub btnNew_Click()
Dim a As String
    a = InputBox("Name:", "New Animation", "unnamed")
    If a <> "" Then
        ani(anicount).name = a
        anicount = anicount + 1
        Call readd_anilist
    End If
End Sub

Private Sub btnNewS_Click()
Dim a As Integer
    If lstSprites.ListIndex <> -1 Then
        For a = lstSprites.ListCount To lstSprites.ListIndex + 1 Step -1
            sprites(a + 1) = sprites(a)
        Next a
    Else
        lstSprites.AddItem ("")
        lstSprites.ListIndex = 0
    End If
    spritecount = spritecount + 1
    sprites(lstSprites.ListIndex + 1).X = 0
    sprites(lstSprites.ListIndex + 1).Y = 0
    sprites(lstSprites.ListIndex + 1).W = 0
    sprites(lstSprites.ListIndex + 1).H = 0
    Call readd_spritelist
End Sub

Private Sub btnPlay_Click()
    Timer1.Enabled = True
    btnPlay.Enabled = False
    btnStop.Enabled = True
End Sub

Private Sub btnQuit_Click()
    End
End Sub

Private Sub btnSave_Click()
Dim name As String
    With comDlg
        .DialogTitle = "Save Animation file"
        .Filter = "Animation file (*.ani)|*.ani"
        .InitDir = workpath
        .FileName = ""
        .ShowSave
        name = .FileName
    End With
    If name <> "" Then
        Call save_ani(name)
        
        
    End If
End Sub

Private Sub btnStop_Click()
    Timer1.Enabled = False
    btnPlay.Enabled = True
    btnStop.Enabled = False

End Sub

Private Sub Form_Load()
    ani_init
    init
End Sub


Private Sub HScroll1_Change()
    picBmp.Left = -HScroll1.Value
    picBmp.Refresh
End Sub


Private Sub Image1_Click()
    Load frmInfo
    frmInfo.Show 1
End Sub

Private Sub lstAni_Click()
    txtspeed.Text = Trim(Str(ani(lstAni.ListIndex).speed))
    btnPlay.Enabled = True
    Call display_anidesc
End Sub

Private Sub lstAnidesc_Click()
    shpSprite.Top = sprites(ani(lstAni.ListIndex).spritelst(lstAnidesc.ListIndex)).Y
    shpSprite.Left = sprites(ani(lstAni.ListIndex).spritelst(lstAnidesc.ListIndex)).X
    shpSprite.Height = sprites(ani(lstAni.ListIndex).spritelst(lstAnidesc.ListIndex)).H
    shpSprite.Width = sprites(ani(lstAni.ListIndex).spritelst(lstAnidesc.ListIndex)).W

    shpSprite.Visible = True

End Sub

Private Sub lstSprites_Click()
    shpSprite.Top = sprites(lstSprites.ListIndex).Y
    shpSprite.Left = sprites(lstSprites.ListIndex).X
    shpSprite.Height = sprites(lstSprites.ListIndex).H
    shpSprite.Width = sprites(lstSprites.ListIndex).W

    shpSprite.Visible = True
End Sub

Private Sub lstSprites_DblClick()
    Call btnChange_Click
End Sub

Private Sub picBmp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a As Integer
Dim b As Integer

    For a = 0 To spritecount
        If X > sprites(a).X And X < sprites(a).X + sprites(a).W And Y > sprites(a).Y And Y < sprites(a).Y + sprites(a).H Then
            lstSprites.ListIndex = a
            lstSprites.Refresh
            If Button = 2 Then
                'add to anilst
                Call btnAdd_Click
            End If
            Exit For
            
        End If
    Next a
End Sub

Private Sub picBmp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmAnimaker.Caption = "X=" + Str(X) + "           Y=" + Str(Y)
End Sub

Private Sub Timer1_Timer()
    'animate!
    If lstAni.ListIndex <> -1 Then
        
        'This is an example of an "engine"-call !!!
    
        Call animate(lstAni.ListIndex)
        picAni.Cls
        Call drawAnimatedSprite(0, 0, lstAni.ListIndex)
    End If
End Sub

Private Sub txtspeed_Change()
    ani(lstAni.ListIndex).speed = Val(txtspeed.Text)
End Sub

Private Sub VScroll1_Change()
    picBmp.Top = -VScroll1.Value
    picBmp.Refresh
End Sub
