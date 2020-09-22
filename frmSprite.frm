VERSION 5.00
Begin VB.Form frmSprite 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Modify sprite definition data"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4410
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtH 
      Height          =   285
      Left            =   3720
      TabIndex        =   10
      Text            =   "Text5"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtW 
      Height          =   285
      Left            =   3720
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   3720
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   3720
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtIndex 
      Height          =   285
      Left            =   3720
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      Height          =   2895
      Left            =   120
      MousePointer    =   2  'Kreuz
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2895
      Begin VB.Shape shpSprite 
         BorderColor     =   &H00000000&
         DrawMode        =   6  'Stift und inverse Anzeige maskieren
         Height          =   975
         Left            =   840
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Height"
      Height          =   195
      Left            =   3120
      TabIndex        =   5
      Top             =   1800
      Width           =   465
   End
   Begin VB.Label Label4 
      Caption         =   "Width"
      Height          =   195
      Left            =   3120
      TabIndex        =   4
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label Label3 
      Caption         =   "Y"
      Height          =   195
      Left            =   3480
      TabIndex        =   3
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label Label2 
      Caption         =   "X"
      Height          =   195
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   105
   End
   Begin VB.Label Label1 
      Caption         =   "Index:"
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnOK_Click()
    Call readd_spritelist
    frmSprite.Hide
    Unload frmSprite
End Sub

Private Sub Form_Load()
    Call redraw
End Sub

Public Sub redraw()
Dim ret As Long
    txtIndex.Text = Trim(Str(msindex))
    txtX.Text = Trim(Str(sprites(msindex).X))
    txtY.Text = Trim(Str(sprites(msindex).Y))
    txtW.Text = Trim(Str(sprites(msindex).W))
    txtH.Text = Trim(Str(sprites(msindex).H))
    picSprite.Cls
    ret = BitBlt(picSprite.hDC, 0, 0, 189, 189, frmAnimaker.picBmp.hDC, sprites(msindex).X + Int(sprites(msindex).W / 2) - 95, sprites(msindex).Y + Int(sprites(msindex).H / 2) - 95, SRCCOPY)
    shpSprite.Left = 95 - Int(sprites(msindex).W / 2)
    shpSprite.Top = 95 - Int(sprites(msindex).H / 2)
    shpSprite.Width = sprites(msindex).W
    shpSprite.Height = sprites(msindex).H
    picSprite.Refresh
End Sub

Private Sub picSprite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmSprite.Caption = "X=" + Str(sprites(msindex).X + Int(sprites(msindex).W / 2) - 95 + X) + "     Y=" + Str(sprites(msindex).Y + Int(sprites(msindex).H / 2) - 95 + Y)
End Sub

Private Sub txtH_Change()
    sprites(msindex).H = Val(txtH.Text)
    Call redraw

End Sub

Private Sub txtIndex_Change()
    msindex = Val(txtIndex.Text)
    Call redraw

End Sub

Private Sub txtW_Change()
    sprites(msindex).W = Val(txtW.Text)
    Call redraw

End Sub

Private Sub txtX_Change()
    sprites(msindex).X = Val(txtX.Text)
    Call redraw
End Sub

Private Sub txtY_Change()
    sprites(msindex).Y = Val(txtY.Text)
    Call redraw

End Sub
