VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://spiele.freepage/blackmushroom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   165
      MousePointer    =   10  'Aufwärtspfeil
      TabIndex        =   6
      Top             =   2160
      Width           =   3315
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   $"frmInfo.frx":0000
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Or you can visit my Homepage at:"
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BlackMushroom@AOL.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   645
      MousePointer    =   10  'Aufwärtspfeil
      TabIndex        =   2
      Top             =   1440
      Width           =   2325
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "(C) 2001 by BlackMushroom Send Comments/Questions to:"
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "AniMaker"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub Label3_Click()
    Call_Url Me, "mailto:BlackMushroom@AOL.com"
End Sub

Private Sub Label6_Click()
    Call_Url Me, "http://spiele.freepage.de/blackmushroom"

End Sub
