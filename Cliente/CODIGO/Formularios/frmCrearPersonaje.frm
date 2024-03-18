VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "ImperiumAO 1.3"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonaje.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox lstFamiliar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      ItemData        =   "frmCrearPersonaje.frx":5D961
      Left            =   8520
      List            =   "frmCrearPersonaje.frx":5D963
      Style           =   2  'Dropdown List
      TabIndex        =   48
      Top             =   1860
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.TextBox txtFamiliar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9330
      MaxLength       =   20
      TabIndex        =   47
      Top             =   990
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.PictureBox picFamiliar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   10380
      ScaleHeight     =   1185
      ScaleWidth      =   840
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1575
      Width           =   870
   End
   Begin VB.PictureBox headview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1695
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   40
      Top             =   4545
      Width           =   375
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":5D965
      Left            =   870
      List            =   "frmCrearPersonaje.frx":5D97E
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3810
      Width           =   2055
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":5D9B1
      Left            =   840
      List            =   "frmCrearPersonaje.frx":5D9BB
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3150
      Width           =   2055
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":5D9D4
      Left            =   870
      List            =   "frmCrearPersonaje.frx":5D9D6
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2490
      Width           =   2055
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2100
      MaxLength       =   25
      TabIndex        =   7
      Top             =   1050
      Width           =   5865
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":5D9D8
      Left            =   8550
      List            =   "frmCrearPersonaje.frx":5D9DA
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3585
      Width           =   2745
   End
   Begin VB.Label lblFamiInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Descropcion del familiar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   8550
      TabIndex        =   49
      Top             =   2295
      Width           =   1635
   End
   Begin VB.Image imgNoDisp 
      Height          =   2145
      Left            =   8415
      Top             =   780
      Width           =   3045
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2520
      TabIndex        =   45
      Top             =   7500
      Width           =   255
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   1
      Left            =   2700
      Top             =   5670
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   1
      Left            =   2700
      Top             =   5820
      Width           =   195
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   2
      Left            =   2700
      Top             =   6030
      Width           =   195
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   3
      Left            =   2700
      Top             =   6390
      Width           =   195
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   4
      Left            =   2700
      Top             =   6750
      Width           =   195
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   5
      Left            =   2700
      Top             =   7080
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   2
      Left            =   2700
      Top             =   6180
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   3
      Left            =   2700
      Top             =   6540
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   4
      Left            =   2700
      Top             =   6900
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   5
      Left            =   2700
      Top             =   7230
      Width           =   195
   End
   Begin VB.Image Head 
      Height          =   600
      Index           =   0
      Left            =   1320
      Top             =   4440
      Width           =   390
   End
   Begin VB.Image Head 
      Height          =   600
      Index           =   1
      Left            =   2160
      Top             =   4440
      Width           =   390
   End
   Begin VB.Label lblModRaza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   44
      Top             =   7140
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   2160
      TabIndex        =   43
      Top             =   6780
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2160
      TabIndex        =   42
      Top             =   6420
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   41
      Top             =   6060
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   39
      Top             =   5700
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgClase 
      Height          =   3570
      Left            =   8490
      Stretch         =   -1  'True
      Top             =   4230
      Width           =   2835
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   5310
      TabIndex        =   38
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   5310
      TabIndex        =   37
      Top             =   2340
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   5310
      TabIndex        =   36
      Top             =   3090
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   5310
      TabIndex        =   35
      Top             =   3450
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   5310
      TabIndex        =   34
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   5310
      TabIndex        =   33
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   5310
      TabIndex        =   32
      Top             =   4590
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   5310
      TabIndex        =   31
      Top             =   4950
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   5310
      TabIndex        =   30
      Top             =   5340
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   5310
      TabIndex        =   29
      Top             =   5700
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   5310
      TabIndex        =   28
      Top             =   6090
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   5310
      TabIndex        =   27
      Top             =   6450
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   5310
      TabIndex        =   26
      Top             =   6840
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   5310
      TabIndex        =   25
      Top             =   7200
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   7365
      TabIndex        =   24
      Top             =   2340
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   7365
      TabIndex        =   23
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   7365
      TabIndex        =   22
      Top             =   3090
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   7365
      TabIndex        =   21
      Top             =   3450
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5D9DC
      Top             =   4680
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5DB2E
      Top             =   4560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   39
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5DC80
      Top             =   4320
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   38
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5DDD2
      Top             =   4170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   37
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5DF24
      Top             =   3930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   36
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E076
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   35
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E1C8
      Top             =   3570
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E31A
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E46C
      Top             =   3180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E5BE
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   31
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E710
      Top             =   2820
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E862
      Top             =   2670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   29
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5E9B4
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   28
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":5EB06
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   26
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5EC58
      Top             =   7170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5EDAA
      Top             =   6810
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5EEFC
      Top             =   6420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F04E
      Top             =   6060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   18
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F1A0
      Top             =   5670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   16
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F2F2
      Top             =   5280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   14
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F444
      Top             =   4920
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F596
      Top             =   4560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F6E8
      Top             =   4170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   8
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F83A
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   6
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5F98C
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5FADE
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   2
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5FC30
      Top             =   2670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5FD82
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   1
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":5FED4
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   27
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60026
      Top             =   7290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   25
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60178
      Top             =   6930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   23
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":602CA
      Top             =   6540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   21
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":6041C
      Top             =   6180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   19
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":6056E
      Top             =   5790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   17
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":606C0
      Top             =   5430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   15
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60812
      Top             =   5040
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   13
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60964
      Top             =   4680
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   11
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60AB6
      Top             =   4290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   9
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60C08
      Top             =   3930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   7
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60D5A
      Top             =   3540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   5
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60EAC
      Top             =   3180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   3
      Left            =   5550
      MouseIcon       =   "frmCrearPersonaje.frx":60FFE
      Top             =   2790
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   7365
      TabIndex        =   20
      Top             =   4590
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   7365
      TabIndex        =   19
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   7365
      TabIndex        =   18
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   7365
      TabIndex        =   17
      Top             =   4950
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   42
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61150
      Top             =   4920
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   43
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":612A2
      Top             =   5040
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   7365
      TabIndex        =   16
      Top             =   5340
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   44
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":613F4
      Top             =   5280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   45
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61546
      Top             =   5430
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   7365
      TabIndex        =   15
      Top             =   5700
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   46
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61698
      Top             =   5670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   47
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":617EA
      Top             =   5790
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   7365
      TabIndex        =   14
      Top             =   6090
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   48
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":6193C
      Top             =   6060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   49
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61A8E
      Top             =   6180
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   7365
      TabIndex        =   13
      Top             =   6450
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   50
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61BE0
      Top             =   6420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   51
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61D32
      Top             =   6540
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   7365
      TabIndex        =   12
      Top             =   6840
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   52
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61E84
      Top             =   6810
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   53
      Left            =   7590
      MouseIcon       =   "frmCrearPersonaje.frx":61FD6
      Top             =   6930
      Width           =   195
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   2610
      TabIndex        =   11
      Top             =   8220
      Width           =   6795
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6795
      TabIndex        =   6
      Top             =   7260
      Width           =   270
   End
   Begin VB.Image boton 
      Height          =   615
      Index           =   1
      Left            =   720
      MouseIcon       =   "frmCrearPersonaje.frx":62128
      MousePointer    =   99  'Custom
      Top             =   8160
      Width           =   1605
   End
   Begin VB.Image boton 
      Height          =   570
      Index           =   0
      Left            =   9600
      MouseIcon       =   "frmCrearPersonaje.frx":6227A
      MousePointer    =   99  'Custom
      Top             =   8160
      Width           =   1680
   End
   Begin VB.Label lblAtributos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   2445
      TabIndex        =   4
      Top             =   6780
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2445
      TabIndex        =   3
      Top             =   6420
      Width           =   210
   End
   Begin VB.Label lblAtributos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   2445
      TabIndex        =   2
      Top             =   7140
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2445
      TabIndex        =   1
      Top             =   6060
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   2445
      TabIndex        =   0
      Top             =   5700
      Width           =   210
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.13.3
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Type tModRaza
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single
End Type

Private Type tModClase
    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    DanoArmas As Double
    DanoProyectiles As Double
    Escudo As Double
    Magia As Double
    Vida As Double
    Hit As Double
End Type

Private ModRaza()  As tModRaza
Private ModClase() As tModClase

Public Actual As Integer
Public SkillPoints As Byte
Private MaxEleccion As Integer, MinEleccion As Integer

Private botonCrear As Boolean

Private Function CheckData() As Boolean
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Comprobacion antes de crear el PJ
'**************************************
    
    '¿Puso un nombre?
    If LenB(txtNombre.Text) = 0 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_NOMBRE_PJ").item("TEXTO")
        txtNombre.SetFocus
        Exit Function
    End If

    '¿Selecciono una raza?
    If UserRaza = 0 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_RAZA").item("TEXTO")
        Exit Function
    End If
    
    '¿Selecciono el Sexo?
    If UserSexo = 0 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_SEXO").item("TEXTO")
        Exit Function
    End If
    
    '¿Seleciono la clase?
    If UserClase = 0 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_CLASE").item("TEXTO")
        Exit Function
    End If

    '¿Estamos intentando crear sin tener el AccountName?
    If Len(AccountName) = 0 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_HASH").item("TEXTO")
        Exit Function
    End If
    
    '¿El nombre de usuario supera los 30 caracteres?
    If LenB(UserName) > 30 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_BAD_NOMBRE_PJ").item("TEXTO").item(1)
        Exit Function
    End If
    
    If UserHogar = 0 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_HOGAR").item("TEXTO")
        Exit Function
    End If
    
    If SkillPoints > 0 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_SKILLS").item("TEXTO")
        Exit Function
    End If
    
    'Toqueteado x Salvito
    Dim i As Integer
    Dim Suma As Byte
    For i = 1 To NUMATRIBUTOS
        If Val(lblAtributos(i).Caption) > 18 Then
            lblInfo.Caption = JsonLanguage.item("VALIDACION_ATRIBUTOS").item("TEXTO")
            Exit Function
        End If
        
        Suma = Suma + lblAtributos(i).Caption
    Next i
    
    If Suma <> 70 Then
        lblInfo.Caption = JsonLanguage.item("VALIDACION_ATRIBUTOS").item("TEXTO")
        Exit Function
    End If
    
    If lstFamiliar.Visible = True Then
    
        If UserPet.tipo = 0 Then
            lblInfo.Caption = "Seleccione su familiar o mascota."
            Exit Function
            
        ElseIf UserPet.Nombre = "" Then
            lblInfo.Caption = "Asigne un nombre a su familiar o mascota."
            Exit Function
            
        ElseIf Len(UserPet.Nombre) > 30 Then
            lblInfo.Caption = ("El nombre de tu familiar o mascota debe tener menos de 30 letras.")
            Exit Function
            
        End If
    
    End If
    
    CheckData = True

End Function

Private Sub Boton_Click(Index As Integer)
    Call Sound.Sound_Play(SND_CLICK)
    
    Select Case Index
    
        Case 0
            
            Dim Count   As Byte
            Dim i       As Integer
            Dim k       As Object
            
            i = 1
            For Each k In Skill
                UserSkills(i) = k.Caption
                i = i + 1
            Next
            
            'Nombre de usuario
            UserName = LTrim(txtNombre.Text)
                    
            '¿El nombre esta vacio y es correcto?
            If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                Call MostrarMensaje(JsonLanguage.item("VALIDACION_BAD_NOMBRE_PJ").item("TEXTO").item(2))
                Exit Sub
            End If
            
            'Solo permitimos 1 espacio en los nombres
            For i = 1 To Len(UserName)
                If mid(UserName, i, 1) = Chr(32) Then Count = Count + 1
            Next i
            
            If Count > 1 Then
                Call MostrarMensaje(JsonLanguage.item("VALIDACION_BAD_NOMBRE_PJ").item("TEXTO").item(3))
                Exit Sub
            End If
            
            UserHogar = lstHogar.ListIndex + 1
            UserPet.tipo = lstFamiliar.ListIndex + 1
            UserPet.Nombre = frmCrearPersonaje.txtFamiliar.Text
            
            'Comprobamos que todo este OK
            If Not CheckData Then Exit Sub
            
            For i = 1 To NUMATRIBUTOS
                UserAtributos(i) = Val(lblAtributos(i).Caption)
            Next i
            
            EstadoLogin = E_MODO.CrearNuevoPJ
            
            'Limpio la lista de hechizos
            frmMain.hlst.Clear
                
            'Conexion!!!
            If Not frmMain.Client.State = sckConnected Then
                Call MostrarMensaje(JsonLanguage.item("ERROR_CONN_LOST").item("TEXTO"))
                Unload Me
            Else
                'Si ya mandamos el paquete, evitamos que se pueda volver a mandar
                botonCrear = True
                Call Login
                botonCrear = False
            End If
            
            'Mandamos el tutorial de inicio
            'bShowTutorial = True
            
        Case 1
            If ClientSetup.bMusic <> CONST_DESHABILITADA Then
                If ClientSetup.bMusic <> CONST_DESHABILITADA Then
                    Sound.NextMusic = MUS_VolverInicio
                    Sound.Fading = 500
                End If
            End If
            
            Unload Me
            
            frmCharList.Visible = True

    End Select
End Sub

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

    Randomize Timer
    
    RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
    If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function

Private Sub Command1_Click(Index As Integer)
    Call Sound.Sound_Play(SND_CLICK)
    
    Dim indice
    If (Index And &H1) = 0 Then
        If SkillPoints > 0 Then
            indice = Index \ 2
            Skill(indice).Caption = Val(Skill(indice).Caption) + 1
            SkillPoints = SkillPoints - 1
        End If
    Else
        If SkillPoints < 10 Then
            
            indice = Index \ 2
            If Val(Skill(indice).Caption) > 0 Then
                Skill(indice).Caption = Val(Skill(indice).Caption) - 1
                SkillPoints = SkillPoints + 1
            End If
        End If
    End If
    
    Puntos.Caption = SkillPoints
End Sub

Private Sub Form_Load()
    Me.Picture = General_Load_Picture_From_Resource("cp-interface.bmp")
    
    Call LoadCharInfo
    
    SkillPoints = 10
    Puntos.Caption = SkillPoints
    
    Dim i As Integer
    
    lstProfesion.Clear
    For i = LBound(ListaClases) To UBound(ListaClases)
        lstProfesion.AddItem ListaClases(i)
    Next i
    
    lstHogar.Clear
    
    For i = LBound(Ciudades()) To UBound(Ciudades())
        lstHogar.AddItem Ciudades(i)
    Next i
    
    lstRaza.Clear
    
    For i = LBound(ListaRazas()) To UBound(ListaRazas())
        lstRaza.AddItem ListaRazas(i)
    Next i
    
    lstProfesion.Clear
    
    For i = LBound(ListaClases()) To UBound(ListaClases())
        lstProfesion.AddItem ListaClases(i)
    Next i
    
    For i = 1 To NUMATRIBUTES
        lblAtributos(i).Caption = 6
    Next i
    
    lstProfesion.ListIndex = 1
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserHead = 0
    lblTotal.Caption = 40
    
End Sub

Private Sub Head_Click(Index As Integer)
    
    Call Sound.Sound_Play(SND_CLICK)
    
    Select Case Index
    
        Case 0
            UserHead = CheckCabeza(UserHead - 1)

        Case 1
            UserHead = CheckCabeza(UserHead + 1)
    
    End Select
    
    If UserHead > 0 Then Call DrawHead(UserHead)

End Sub

Private Sub ImgAtributoMas_Click(Index As Integer)
    Call Sound.Sound_Play(SND_CLICK)
    
    If lblAtributos(Index).Caption < 18 And lblTotal.Caption > 0 Then
        lblAtributos(Index).Caption = lblAtributos(Index).Caption + 1
        lblTotal.Caption = lblTotal.Caption - 1
    End If
    
End Sub

Private Sub ImgAtributoMenos_Click(Index As Integer)
    Call Sound.Sound_Play(SND_CLICK)
    
    If lblTotal.Caption = "40" Then Exit Sub
    If lblAtributos(Index).Caption > 6 Then
        lblAtributos(Index).Caption = lblAtributos(Index).Caption - 1
        lblTotal.Caption = lblTotal.Caption + 1
    End If
    
End Sub

Private Sub lstFamiliar_Click()

    If lstFamiliar.ListIndex > 0 Then
        lblFamiInfo.Caption = ListaFamiliares(lstFamiliar.ListIndex).Desc
        picFamiliar.Picture = General_Load_Picture_From_Resource(ListaFamiliares(lstFamiliar.ListIndex).Imagen)
    Else
        lblFamiInfo.Caption = "Selecciona tu familiar o mascota para saber más de él"
        picFamiliar.Picture = Nothing
    End If

End Sub

Private Sub lstProfesion_Click()
On Error Resume Next
    imgClase.Picture = General_Load_Picture_From_Resource(LCase(lstProfesion.Text & ".bmp"))
    
    UserClase = lstProfesion.ListIndex + 1
    
    If UserClase = eClass.Mage Then
        frmCrearPersonaje.txtFamiliar.Visible = True
        frmCrearPersonaje.lstFamiliar.Visible = True
        imgNoDisp.Picture = Nothing
        lblFamiInfo.Visible = True
        picFamiliar.Visible = True
        Call CambioFamiliar(5)
        
    ElseIf UserClase = eClass.Hunter Or UserClase = eClass.Druid Then
        frmCrearPersonaje.txtFamiliar.Visible = True
        frmCrearPersonaje.lstFamiliar.Visible = True
        imgNoDisp.Picture = Nothing
        lblFamiInfo.Visible = True
        picFamiliar.Visible = True
        Call CambioFamiliar(4)
    
    Else
    
        frmCrearPersonaje.txtFamiliar.Visible = False
        frmCrearPersonaje.lstFamiliar.Visible = False
        imgNoDisp.Picture = General_Load_Picture_From_Resource("mascotanodisp" & ".bmp")
        picFamiliar.Visible = False
        lblFamiInfo.Visible = False
        
    End If
    
    Call UpdateRazaMod
    
End Sub

Private Sub lstGenero_Click()
    UserSexo = lstGenero.ListIndex + 1
    Call DameCabezas
    
End Sub

Private Sub lstRaza_Click()
    UserRaza = lstRaza.ListIndex + 1
    Call DameCabezas
    
    Call UpdateRazaMod
    
End Sub

Sub DameCabezas()
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Asignamos un cuerpo y unac abeza segun la raza y el sexo
'**************************************

    Select Case UserSexo
    
        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano
                    UserHead = eCabezas.HUMANO_H_PRIMER_CABEZA
                    UserBody = eCabezas.HUMANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = eCabezas.ELFO_H_PRIMER_CABEZA
                    UserBody = eCabezas.ELFO_H_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = eCabezas.DROW_H_PRIMER_CABEZA
                    UserBody = eCabezas.DROW_H_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = eCabezas.ENANO_H_PRIMER_CABEZA
                    UserBody = eCabezas.ENANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = eCabezas.GNOMO_H_PRIMER_CABEZA
                    UserBody = eCabezas.GNOMO_H_CUERPO_DESNUDO
                    
                Case eRaza.Orco
                    UserHead = eCabezas.ORCO_H_PRIMER_CABEZA
                    UserBody = eCabezas.ORCO_H_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
            
        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano
                    UserHead = eCabezas.HUMANO_M_PRIMER_CABEZA
                    UserBody = eCabezas.HUMANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = eCabezas.ELFO_M_PRIMER_CABEZA
                    UserBody = eCabezas.ELFO_M_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = eCabezas.DROW_M_PRIMER_CABEZA
                    UserBody = eCabezas.DROW_M_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = eCabezas.ENANO_M_PRIMER_CABEZA
                    UserBody = eCabezas.ENANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = eCabezas.GNOMO_M_PRIMER_CABEZA
                    UserBody = eCabezas.GNOMO_M_CUERPO_DESNUDO
                    
                Case eRaza.Orco
                    UserHead = eCabezas.ORCO_M_PRIMER_CABEZA
                    UserBody = eCabezas.ORCO_M_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
            
        Case Else
            UserHead = 0
            UserBody = 0
            
    End Select
    
    If UserHead > 0 Then Call DrawHead(UserHead)
    
End Sub

Private Function CheckCabeza(ByVal Head As Integer) As Integer

On Error GoTo errhandler

    Select Case UserSexo

        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano

                    If Head > eCabezas.HUMANO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.HUMANO_H_PRIMER_CABEZA + (Head - eCabezas.HUMANO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.HUMANO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.HUMANO_H_ULTIMA_CABEZA - (eCabezas.HUMANO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Elfo

                    If Head > eCabezas.ELFO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ELFO_H_PRIMER_CABEZA + (Head - eCabezas.ELFO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ELFO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ELFO_H_ULTIMA_CABEZA - (eCabezas.ELFO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.ElfoOscuro

                    If Head > eCabezas.DROW_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.DROW_H_PRIMER_CABEZA + (Head - eCabezas.DROW_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.DROW_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.DROW_H_ULTIMA_CABEZA - (eCabezas.DROW_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Enano

                    If Head > eCabezas.ENANO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ENANO_H_PRIMER_CABEZA + (Head - eCabezas.ENANO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ENANO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ENANO_H_ULTIMA_CABEZA - (eCabezas.ENANO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Gnomo

                    If Head > eCabezas.GNOMO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.GNOMO_H_PRIMER_CABEZA + (Head - eCabezas.GNOMO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.GNOMO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.GNOMO_H_ULTIMA_CABEZA - (eCabezas.GNOMO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                    
                Case eRaza.Orco

                    If Head > eCabezas.ORCO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ORCO_H_PRIMER_CABEZA + (Head - eCabezas.ORCO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ORCO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ORCO_H_ULTIMA_CABEZA - (eCabezas.ORCO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case Else
                    CheckCabeza = CheckCabeza(Head)
                    
            End Select
        
        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano

                    If Head > eCabezas.HUMANO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.HUMANO_M_PRIMER_CABEZA + (Head - eCabezas.HUMANO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.HUMANO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.HUMANO_M_ULTIMA_CABEZA - (eCabezas.HUMANO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Elfo

                    If Head > eCabezas.ELFO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ELFO_M_PRIMER_CABEZA + (Head - eCabezas.ELFO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ELFO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ELFO_M_ULTIMA_CABEZA - (eCabezas.ELFO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.ElfoOscuro

                    If Head > eCabezas.DROW_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.DROW_M_PRIMER_CABEZA + (Head - eCabezas.DROW_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.DROW_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.DROW_M_ULTIMA_CABEZA - (eCabezas.DROW_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Enano

                    If Head > eCabezas.ENANO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ENANO_M_PRIMER_CABEZA + (Head - eCabezas.ENANO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ENANO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ENANO_M_ULTIMA_CABEZA - (eCabezas.ENANO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Gnomo

                    If Head > eCabezas.GNOMO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.GNOMO_M_PRIMER_CABEZA + (Head - eCabezas.GNOMO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.GNOMO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.GNOMO_M_ULTIMA_CABEZA - (eCabezas.GNOMO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Orco

                    If Head > eCabezas.ORCO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ORCO_M_PRIMER_CABEZA + (Head - eCabezas.ORCO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ORCO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ORCO_M_ULTIMA_CABEZA - (eCabezas.ORCO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case Else
                    CheckCabeza = Head
                    
            End Select

        Case Else
            CheckCabeza = Head
            
    End Select
    
errhandler:

    If Err.number Then
        Call LogError(Err.number, Err.Description, "frmCrearPersonaje.CheckCabeza")
    End If
    
    Exit Function
    
End Function

Public Sub UpdateRazaMod()
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Actualiza los modificadores de atributos que otorga cada raza
'**************************************

    If lstRaza.ListIndex > -1 Then
        
        UserRaza = lstRaza.ListIndex + 1
        
        With ModRaza(UserRaza)
            lblModRaza(eAtributos.Fuerza) = IIf(.Fuerza >= 0, "+", vbNullString) & .Fuerza
            lblModRaza(eAtributos.Agilidad) = IIf(.Agilidad >= 0, "+", vbNullString) & .Agilidad
            lblModRaza(eAtributos.Inteligencia) = IIf(.Inteligencia >= 0, "+", vbNullString) & .Inteligencia
            lblModRaza(eAtributos.Carisma) = IIf(.Carisma >= 0, "+", "") & .Carisma
            lblModRaza(eAtributos.Constitucion) = IIf(.Constitucion >= 0, "+", vbNullString) & .Constitucion
        End With
        
    End If
    
End Sub

Private Sub LoadCharInfo()
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Carga los modificadores de cada raza
'**************************************

    Dim SearchVar As String
    Dim i         As Integer

    ReDim ModRaza(1 To NUMRAZAS)

    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(Carga.Path(Lenguajes) & "CharInfo_" & Language & ".dat")
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS
    
        With ModRaza(i)
            SearchVar = Replace(ListaRazas(i), " ", vbNullString)
        
            .Fuerza = CSng(Lector.GetValue("MODRAZA", SearchVar + "Fuerza"))
            .Agilidad = CSng(Lector.GetValue("MODRAZA", SearchVar + "Agilidad"))
            .Inteligencia = CSng(Lector.GetValue("MODRAZA", SearchVar + "Inteligencia"))
            .Carisma = CSng(Lector.GetValue("MODRAZA", SearchVar + "Carisma"))
            .Constitucion = CSng(Lector.GetValue("MODRAZA", SearchVar + "Constitucion"))
        End With
        
    Next i

End Sub

Private Sub DrawHead(ByVal Head As Integer)

    Dim DR  As RECT
    Dim Grh As Long

    Grh = HeadData(Head).Head(3).GrhIndex
    
    With headview
        DR.Right = .Width - 5
        DR.Bottom = .Height - 3
        DR.Left = -5
        DR.Top = -3
    End With
        
    Call DrawGrhtoHdc(headview, Grh, DR)

End Sub

Private Sub CambioFamiliar(ByVal NumFamiliares As Integer)

    If NumFamiliares = 5 Then
    
        ReDim ListaFamiliares(1 To NumFamiliares) As tListaFamiliares
        ListaFamiliares(1).name = "Elemental De Fuego"
        ListaFamiliares(1).Desc = "Hecho de puro fuego, lanzará tormentas sobre tus contrincantes."
        ListaFamiliares(1).Imagen = "elefuego.bmp"
        
        ListaFamiliares(2).name = "Elemental De Agua"
        ListaFamiliares(2).Desc = "Con su cuerpo acuoso paralizará a tus enemigos."
        ListaFamiliares(2).Imagen = "eleagua.bmp"
        
        ListaFamiliares(3).name = "Elemental De Tierra"
        ListaFamiliares(3).Desc = "Sus fuertes brazos inmovilizarán cualquier criatura viviente."
        ListaFamiliares(3).Imagen = "eletierra.bmp"
        
        ListaFamiliares(4).name = "Ely"
        ListaFamiliares(4).Desc = "Te protegerá constantemente con sus conjuros defensivos."
        ListaFamiliares(4).Imagen = "ely.bmp"
        
        ListaFamiliares(5).name = "Fuego Fatuo"
        ListaFamiliares(5).Desc = "Débil pero con gran poder mágico, siempre estará a tu lado."
        ListaFamiliares(5).Imagen = "fatuo.bmp"
        
    Else
    
        ReDim ListaFamiliares(1 To NumFamiliares) As tListaFamiliares
        ListaFamiliares(1).name = "Tigre"
        ListaFamiliares(1).Desc = "Poseen grandes y filosas garras para atacar a tus oponentes."
        ListaFamiliares(1).Imagen = "tigre.bmp"
        
        ListaFamiliares(2).name = "Lobo"
        ListaFamiliares(2).Desc = "Astutos y arrogantes, su mordedura causa estragos en sus víctimas."
        ListaFamiliares(2).Imagen = "lobo.bmp"
        
        ListaFamiliares(3).name = "Oso Pardo"
        ListaFamiliares(3).Desc = "Se caracterizan por ser territoriales y muy resistentes."
        ListaFamiliares(3).Imagen = "oso.bmp"
        
        ListaFamiliares(4).name = "Ent"
        ListaFamiliares(4).Desc = "¡Esta robusta criatura te defenderá cual muro de piedra!"
        ListaFamiliares(4).Imagen = "ent.bmp"
    
    End If
    
    Dim i As Integer
    lstFamiliar.Clear
    lstFamiliar.AddItem ""
    For i = 1 To UBound(ListaFamiliares)
        lstFamiliar.AddItem ListaFamiliares(i).name
    Next i
    
    lstFamiliar.ListIndex = 0

End Sub

Private Sub txtfamiliar_GotFocus()
    lblInfo.Caption = "Mucho cuidado al colocarle nombre a su familiar, no puede ponerle el mismo o parecido nombre de su personaje, recuerde que es su companía. En caso de que el familiar o mascota tenga nombre inapropiado, podrá ser retirado."
    
End Sub

Private Sub txtNombre_GotFocus()
    lblInfo.Caption = "Sea cuidadoso al seleccionar el nombre de su personaje, ImperiumClasico es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotación politica los administradores borrarán su personaje y no habrá ninguna posibilidad de recuperarlo."

End Sub
