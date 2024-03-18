VERSION 5.00
Begin VB.Form frmPanelAccount 
   BorderStyle     =   0  'None
   Caption         =   "Panel de Cuenta"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPanelAccount.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   9
      Left            =   6180
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2220
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   8
      Left            =   4680
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2220
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   7
      Left            =   3180
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2220
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   6
      Left            =   1680
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2220
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   5
      Left            =   180
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2220
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   4
      Left            =   6180
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   420
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   3
      Left            =   4680
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   420
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   2
      Left            =   3180
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   420
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   1
      Left            =   1680
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   0
      Left            =   180
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   420
      Width           =   1140
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   3945
      TabIndex        =   23
      Top             =   4410
      Width           =   45
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   3975
      TabIndex        =   22
      Top             =   4290
      Width           =   675
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   3975
      TabIndex        =   21
      Top             =   4140
      Width           =   345
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   10
      Left            =   6075
      TabIndex        =   20
      Top             =   1860
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   9
      Left            =   4575
      TabIndex        =   19
      Top             =   1860
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   8
      Left            =   3060
      TabIndex        =   18
      Top             =   1860
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   7
      Left            =   1575
      TabIndex        =   17
      Top             =   1860
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   60
      TabIndex        =   16
      Top             =   1860
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   6075
      TabIndex        =   15
      Top             =   60
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   4575
      TabIndex        =   14
      Top             =   60
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   3060
      TabIndex        =   13
      Top             =   60
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   1575
      TabIndex        =   12
      Top             =   60
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   60
      TabIndex        =   11
      Top             =   60
      Width           =   1365
   End
   Begin VB.Image cmdConnect 
      Height          =   615
      Left            =   5790
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   4065
      Width           =   1755
   End
   Begin VB.Image cmdDelete 
      Height          =   615
      Left            =   1920
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   4065
      Width           =   1755
   End
   Begin VB.Image cmdCrear 
      Height          =   615
      Left            =   0
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   4065
      Width           =   1755
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   9
      Left            =   6015
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   8
      Left            =   4515
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   7
      Left            =   3015
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   6
      Left            =   1515
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   5
      Left            =   15
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   4
      Left            =   6015
      Top             =   0
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   3
      Left            =   4515
      Top             =   0
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   2
      Left            =   3015
      Top             =   0
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   1
      Left            =   1515
      Top             =   0
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   0
      Left            =   15
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lblAccData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2250
      TabIndex        =   0
      Top             =   2370
      Width           =   3705
   End
   Begin VB.Image cmdExit 
      Height          =   615
      Left            =   8025
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   2085
      Width           =   1755
   End
   Begin VB.Image cmdChange 
      Height          =   615
      Left            =   6180
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   2085
      Width           =   1755
   End
End
Attribute VB_Name = "frmPanelAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
