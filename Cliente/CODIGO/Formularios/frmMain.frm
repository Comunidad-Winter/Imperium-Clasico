VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   8970
   ClientLeft      =   15
   ClientTop       =   -3300
   ClientWidth     =   11970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   DrawStyle       =   6  'Inside Solid
   FillColor       =   &H00008080&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004080&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmMain.frx":0ECA
   ScaleHeight     =   598
   ScaleMode       =   0  'User
   ScaleWidth      =   798
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   0
      Left            =   240
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   1
      Left            =   825
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   3
      Left            =   1995
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   2
      Left            =   1410
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   4
      Left            =   2580
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   5
      Left            =   3165
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   6
      Left            =   3750
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   7
      Left            =   4335
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   8
      Left            =   4920
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   9
      Left            =   5505
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   10
      Left            =   6090
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6225
      Left            =   210
      ScaleHeight     =   415
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   545
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2070
      Width           =   8175
      Begin VB.Timer Second 
         Enabled         =   0   'False
         Interval        =   1050
         Left            =   7560
         Top             =   120
      End
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   210
      MaxLength       =   500
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1755
      Visible         =   0   'False
      Width           =   7470
   End
   Begin VB.PictureBox MiniMapa 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   10200
      MouseIcon       =   "frmMain.frx":37610
      ScaleHeight     =   97
      ScaleMode       =   0  'User
      ScaleWidth      =   97
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7335
      Width           =   1455
      Begin VB.Shape UserM 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         Height          =   45
         Left            =   750
         Top             =   750
         Width           =   45
      End
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   9015
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   1
      Top             =   2220
      Width           =   2415
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   8865
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2085
      Visible         =   0   'False
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1455
      Left            =   240
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   180
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   2566
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":37762
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image cmdCerrar 
      Height          =   225
      Left            =   11580
      Top             =   150
      Width           =   255
   End
   Begin VB.Image cmdMinimizar 
      Height          =   225
      Left            =   11340
      Top             =   150
      Width           =   225
   End
   Begin VB.Image imgMiniCerra 
      Enabled         =   0   'False
      Height          =   315
      Left            =   11325
      Top             =   150
      Width           =   510
   End
   Begin VB.Image nomodocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":377DF
      ToolTipText     =   "Modo Combate"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image nomodoseguro 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":37C1D
      ToolTipText     =   "Seguro"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image nomodorol 
      Height          =   255
      Left            =   9645
      Picture         =   "frmMain.frx":3805B
      ToolTipText     =   "Modo Rol"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Image modocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":38499
      ToolTipText     =   "Modo Combate"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image modoseguro 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":388D7
      ToolTipText     =   "Seguro"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image modorol 
      Height          =   255
      Left            =   9645
      Picture         =   "frmMain.frx":38D15
      ToolTipText     =   "Modo Rol"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Label MapExp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   1
      Left            =   8820
      TabIndex        =   19
      Top             =   870
      Width           =   1815
   End
   Begin VB.Shape ExpShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8820
      Top             =   900
      Width           =   1815
   End
   Begin VB.Label MapExp 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa desconocido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   8640
      TabIndex        =   18
      Top             =   7020
      Width           =   3105
   End
   Begin VB.Label lblInvInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   9000
      TabIndex        =   17
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Image cmdOpciones 
      Height          =   450
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   4935
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Image cmdQuest 
      Height          =   450
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   3765
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Image cmdClanes 
      Height          =   450
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   3180
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Image cmdEstadisticas 
      Height          =   450
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   2595
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Image cmdGrupo 
      Height          =   450
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   2010
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Image cmdTorneos 
      Height          =   450
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   4350
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label lblFU 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9330
      TabIndex        =   16
      Top             =   8340
      Width           =   345
   End
   Begin VB.Label lblAG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9330
      TabIndex        =   15
      Top             =   8550
      Width           =   345
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   135
      Left            =   8745
      TabIndex        =   14
      Top             =   5850
      Width           =   1350
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   135
      Left            =   8745
      TabIndex        =   13
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label lblST 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   135
      Left            =   8745
      TabIndex        =   12
      Top             =   6660
      Width           =   1365
   End
   Begin VB.Label lblHAM 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   135
      Left            =   10320
      TabIndex        =   11
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label lblSED 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Height          =   135
      Left            =   10320
      TabIndex        =   10
      Top             =   6660
      Width           =   1365
   End
   Begin VB.Image cmdDropGold 
      Height          =   300
      Left            =   10260
      MousePointer    =   99  'Custom
      Top             =   5670
      Width           =   300
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      Height          =   135
      Left            =   10320
      Top             =   6690
      Width           =   1365
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   135
      Left            =   10320
      Top             =   6285
      Width           =   1365
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   135
      Left            =   8745
      Top             =   6690
      Width           =   1365
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8745
      Top             =   6285
      Width           =   1365
   End
   Begin VB.Shape Hpshp 
      BackColor       =   &H00000080&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8745
      Top             =   5880
      Width           =   1365
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10950
      TabIndex        =   9
      Top             =   870
      Width           =   435
   End
   Begin VB.Image lblChat 
      Height          =   255
      Left            =   7815
      Top             =   1740
      Width           =   555
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NickDelPersonaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8610
      TabIndex        =   8
      Top             =   180
      Width           =   2625
   End
   Begin VB.Image imgHora 
      Height          =   480
      Left            =   6675
      Top             =   8430
      Width           =   1695
   End
   Begin VB.Image btnSolapa 
      Height          =   510
      Index           =   2
      Left            =   10740
      MouseIcon       =   "frmMain.frx":392AB
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Image btnInfo 
      Height          =   390
      Left            =   10650
      MouseIcon       =   "frmMain.frx":393FD
      MousePointer    =   99  'Custom
      Top             =   4935
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image btnLanzar 
      Height          =   390
      Left            =   8775
      MouseIcon       =   "frmMain.frx":3954F
      MousePointer    =   99  'Custom
      Top             =   4935
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Image btnSolapa 
      Height          =   510
      Index           =   1
      Left            =   9660
      MouseIcon       =   "frmMain.frx":396A1
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Image btnSolapa 
      Height          =   510
      Index           =   0
      Left            =   8580
      MouseIcon       =   "frmMain.frx":397F3
      Top             =   1230
      Width           =   1065
   End
   Begin VB.Label lblMinimizar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   14565
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   420
      Index           =   0
      Left            =   11460
      MouseIcon       =   "frmMain.frx":39945
      MousePointer    =   99  'Custom
      Top             =   3405
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   420
      Index           =   1
      Left            =   11475
      MouseIcon       =   "frmMain.frx":39A97
      MousePointer    =   99  'Custom
      Top             =   2910
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label GldLbl 
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
      Height          =   165
      Left            =   10620
      TabIndex        =   0
      Top             =   5745
      Width           =   1110
   End
   Begin VB.Image InvEqu 
      Height          =   4275
      Left            =   8580
      Picture         =   "frmMain.frx":39BE9
      Top             =   1230
      Width           =   3240
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmMain
'    Project    : ARGENTUM
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Marquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matias Fernando Pequeno
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
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez

Option Explicit

Public tX                  As Byte
Public tY                  As Byte
Public MouseX              As Long
Public MouseY              As Long
Public MouseBoton          As Long
Public MouseShift          As Long
Private clicX              As Long
Private clicY              As Long

Public UltPos As Integer

Private clsFormulario           As clsFormMovementManager
Private cBotonLanzar            As clsGraphicalButton
Private cBotonInfo              As clsGraphicalButton
Private cBotonCerrar            As clsGraphicalButton
Private cBotonMinimizar         As clsGraphicalButton
Private cBotonChat              As clsGraphicalButton
Private cBotonGrupo             As clsGraphicalButton
Private cBotonEstadisticas      As clsGraphicalButton
Private cBotonClanes            As clsGraphicalButton
Private cBotonQuest             As clsGraphicalButton
Private cBotonTorneos           As clsGraphicalButton
Private cBotonOpciones          As clsGraphicalButton

Public LastButtonPressed   As clsGraphicalButton

Public WithEvents Client   As clsSocket
Attribute Client.VB_VarHelpID = -1

Private FirstTimeChat      As Boolean

'Usado para controlar que no se dispare el binding de la tecla CTRL cuando se usa CTRL+Tecla.
Dim CtrlMaskOn             As Boolean

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Public Sub dragInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call Protocol.WriteMoveItem(originalSlot, newSlot, eMoveType.Inventory)
End Sub

Private Sub btnGrupo_Click()
    Call WriteRequestPartyForm
End Sub

Private Sub cmdCerrar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    If UserParalizado Then 'Inmo

        With FontTypes(FontTypeNames.FONTTYPE_WARNING)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_NO_SALIR").item("TEXTO"), .Red, .Green, .Blue, .bold, .italic)
        End With
        
        Exit Sub
        
    End If
    
    ' Nos desconectamos y lo mando al Panel de la Cuenta
    Call WriteQuit
End Sub

Private Sub btnSolapa_Click(Index As Integer)
    Call Sound.Sound_Play(SND_CLICK)

    Select Case Index
    
        Case 0 'Inventario
            InvEqu.Picture = General_Load_Picture_From_Resource("centroinventario.bmp", True)
            
            ' Activo controles de inventario
            PicInv.Visible = True
        
            ' Desactivo controles de hechizo
            hlst.Visible = False
            btnInfo.Visible = False
            btnLanzar.Visible = False
            
            cmdMoverHechi(0).Visible = False
            cmdMoverHechi(1).Visible = False
            
            DoEvents
            Call Inventario.DrawInventory
            
            cmdGrupo.Visible = False
            cmdEstadisticas.Visible = False
            cmdClanes.Visible = False
            cmdQuest.Visible = False
            cmdTorneos.Visible = False
            cmdOpciones.Visible = False
            
            lblInvInfo.Visible = True
        
        Case 1 'Hechizos
            InvEqu.Picture = General_Load_Picture_From_Resource("centrohechizos.bmp", True)
            
            ' Activo controles de hechizos
            hlst.Visible = True
            btnInfo.Visible = True
            btnLanzar.Visible = True
            
            cmdMoverHechi(0).Visible = True
            cmdMoverHechi(1).Visible = True
            
            ' Desactivo controles de inventario
            PicInv.Visible = False
            
            cmdGrupo.Visible = False
            cmdEstadisticas.Visible = False
            cmdClanes.Visible = False
            cmdQuest.Visible = False
            cmdTorneos.Visible = False
            cmdOpciones.Visible = False
            
            lblInvInfo.Visible = False
            
        Case 2 'Menu
            InvEqu.Picture = General_Load_Picture_From_Resource("centromenu.bmp", True)
            
            ' Activo controles de inventario
            PicInv.Visible = False
        
            ' Desactivo controles de hechizo
            hlst.Visible = False
            btnInfo.Visible = False
            btnLanzar.Visible = False
            
            cmdMoverHechi(0).Visible = False
            cmdMoverHechi(1).Visible = False
            
            cmdGrupo.Visible = True
            cmdEstadisticas.Visible = True
            cmdClanes.Visible = True
            cmdQuest.Visible = True
            cmdTorneos.Visible = True
            cmdOpciones.Visible = True
            
            lblInvInfo.Visible = False
            
    End Select
End Sub

Private Sub cmdDropGold_Click()
    Inventario.SelectGold
    If UserGLD > 0 Then
        frmCantidad.Show , frmMain
    End If
End Sub

Private Sub cmdGrupo_Click()
    Call WriteRequestPartyForm
End Sub

Private Sub cmdMinimizar_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub Form_Activate()

    Call Inventario.DrawInventory

End Sub

Private Sub Form_Load()
    ClientSetup.SkinSeleccionado = GetVar(Carga.Path(Init) & CLIENT_FILE, "Parameters", "SkinSelected")
    
    'Me.Picture = General_Load_Picture_From_Resource("1.bmp", True)
    Me.Picture = General_Load_Picture_From_Resource("todo.bmp", True)
    InvEqu.Picture = General_Load_Picture_From_Resource("centroinventario.bmp", True)
    
    cmdMoverHechi(1).Picture = General_Load_Picture_From_Resource("[hechizos]flechaarriba-down.bmp", True)
    cmdMoverHechi(0).Picture = General_Load_Picture_From_Resource("[hechizos]flechaabajo-down.bmp", True)
    
    If Not ResolucionCambiada Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        Call clsFormulario.Initialize(Me, 120)
    End If
        
    Call LoadButtons
    
    ' Seteamos el caption
    Me.Caption = Form_Caption
    
    ' Removemos la barra de titulo pero conservando el caption para la barra de tareas
    Call Form_RemoveTitleBar(Me)

    ' Reseteamos el tamanio de la ventana para que no queden bordes blancos
    Me.Width = 12000
    Me.Height = 9000
        
    ' Detect links in console
    Call EnableURLDetect(RecTxt.hWnd, Me.hWnd)
    
    ' Make the console transparent
    Call SetWindowLong(RecTxt.hWnd, -20, &H20&)
    RecTxt.BackColor = RGB(24, 23, 21)
    
    CtrlMaskOn = False
    
    FirstTimeChat = True
    SendingType = 1
    
    UltPos = -1
    
End Sub

Private Sub LoadButtons()
    
    Dim GrhPath As String
    Dim i As Integer

    Set LastButtonPressed = New clsGraphicalButton
    
    lblMinimizar.MouseIcon = picMouseIcon
    
    Set cBotonLanzar = New clsGraphicalButton
    Set cBotonInfo = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonMinimizar = New clsGraphicalButton
    Set cBotonChat = New clsGraphicalButton
    Set cBotonGrupo = New clsGraphicalButton
    Set cBotonEstadisticas = New clsGraphicalButton
    Set cBotonClanes = New clsGraphicalButton
    Set cBotonQuest = New clsGraphicalButton
    Set cBotonTorneos = New clsGraphicalButton
    Set cBotonOpciones = New clsGraphicalButton
    
    
    Call cBotonLanzar.Initialize(btnLanzar, "", _
                                 "[hechizos]lanzar-over.bmp", _
                                 "[hechizos]lanzar-down.bmp", Me, , , , , True)
                                     
    Call cBotonInfo.Initialize(btnInfo, "", _
                               "[hechizos]info-over.bmp", _
                               "[hechizos]info-down.bmp", Me, , , , , True)
                                     
    Call cBotonCerrar.Initialize(cmdCerrar, "", _
                                "cerrarover.bmp", _
                                "cerrardown.bmp", Me, , , , , True)
                                
    Call cBotonMinimizar.Initialize(cmdMinimizar, "", _
                                "minimizarover.bmp", _
                                "minimizardown.bmp", Me, , , , , True)
                                
    Call cBotonChat.Initialize(lblChat, "", _
                               "modotextoover.bmp", _
                               "modotextodown.bmp", Me, , , , , True)
                                
    Call cBotonGrupo.Initialize(cmdGrupo, "", _
                               "[menu]grupo-over.bmp", _
                               "[menu]grupo-down.bmp", Me, , , , , True)
                               
    Call cBotonEstadisticas.Initialize(cmdEstadisticas, "", _
                               "[menu]estadisticas-over.bmp", _
                               "[menu]estadisticas-down.bmp", Me, , , , , True)
                                                              
    Call cBotonClanes.Initialize(cmdClanes, "", _
                               "[menu]clanes-over.bmp", _
                               "[menu]clanes-down.bmp", Me, , , , , True)
                                                              
    Call cBotonQuest.Initialize(cmdQuest, "", _
                               "[menu]quests-over.bmp", _
                               "[menu]quests-down.bmp", Me, , , , , True)
                                                                                     
    Call cBotonTorneos.Initialize(cmdTorneos, "", _
                               "[menu]torneos-over.bmp", _
                               "[menu]torneos-down.bmp", Me, , , , , True)
                                                              
    Call cBotonOpciones.Initialize(cmdOpciones, "", _
                               "[menu]opciones-over.bmp", _
                               "[menu]opciones-down.bmp", Me, , , , , True)

End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)

    If hlst.Visible = True Then
        If hlst.ListIndex = -1 Then Exit Sub
        Dim sTemp As String
    
        Select Case Index

            Case 1 'subir

                If hlst.ListIndex = 0 Then Exit Sub

            Case 0 'bajar

                If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
        End Select
    
        Call WriteMoveSpell(Index = 1, hlst.ListIndex + 1)
        
        Select Case Index

            Case 1 'subir
                sTemp = hlst.List(hlst.ListIndex - 1)
                hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex - 1

            Case 0 'bajar
                sTemp = hlst.List(hlst.ListIndex + 1)
                hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex + 1
        End Select
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    '***************************************************
    'Autor: Unknown
    'Last Modification: 18/11/2010
    '18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
    '18/11/2010: Amraphen - Agregue el handle correspondiente para las nuevas configuraciones de teclas (CTRL+0..9).
    '***************************************************
    If (Not SendTxt.Visible) Then
        
        If KeyCode = vbKeyControl Then

            'Chequeo que no se haya usado un CTRL + tecla antes de disparar las bindings.
            If CtrlMaskOn Then
                CtrlMaskOn = False
                Exit Sub
            End If
        End If
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then

            Select Case KeyCode

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    If ClientSetup.bMusic = CONST_MP3 Then
                        Sound.Music_Stop
                        ClientSetup.bMusic = CONST_DESHABILITADA
                    Else
                        ClientSetup.bMusic = CONST_MP3
                    End If
                        
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSound)
                    'Audio.SoundActivated = Not Audio.SoundActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
                    ClientSetup.FPSShow = Not ClientSetup.FPSShow
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Domar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Robar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Ocultarse)
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                        
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)

                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    Call WriteSafeToggle

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatSafe)
                    Call WriteCombatToggle
                    
            End Select
            
        End If
        
    
        Select Case KeyCode
            
            Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
                Call Mod_General.Client_Screenshot(frmMain.hDC, 800, 600)
                    
            Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
                Call frmOpciones.Show(vbModeless, frmMain)
                
            Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
    
                Call WriteQuit
                
            Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
    
                If Shift <> 0 Then Exit Sub
                
                If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
                If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                    If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
                Else
    
                    If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub
                End If
                
                If frmCustomKeys.Visible Then Exit Sub 'Chequeo si esta visible la ventana de configuracion de teclas.
                
                Call WriteAttack
                
            Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
                
                If (Not Comerciando) And (Not MirandoAsignarSkills) And (Not frmMSG.Visible) And (Not MirandoForo) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                    Call CompletarEnvioMensajes
                    SendTxt.Visible = True
                    SendTxt.SetFocus
                Else
                    Call Enviar_SendTxt
                End If
        
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionUno)
                Call accionMacrosKey(1)
                
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionDos)
                Call accionMacrosKey(2)
                
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionTres)
                Call accionMacrosKey(3)
            
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionCuatro)
                Call accionMacrosKey(4)
            
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionCinco)
                Call accionMacrosKey(5)
            
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionSeis)
                Call accionMacrosKey(6)
            
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionSiete)
                Call accionMacrosKey(7)
            
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionOcho)
                Call accionMacrosKey(8)
            
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionNueve)
                Call accionMacrosKey(9)
            
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionDiez)
                Call accionMacrosKey(10)
            
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionOnce)
                Call accionMacrosKey(11)
            
        End Select
     End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call DisableURLDetect
    
End Sub

Private Sub btnClanes_Click()
    
    If frmGuildLeader.Visible Then Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
End Sub

Private Sub cmdEstadisticas_Click()

    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    LlegoFamily = False
    Call WriteRequestAtributes
    Call WriteRequestSkills
    Call WriteRequestMiniStats
    Call WriteRequestFame
    Call WriteRequestFamily
    Call FlushBuffer

    Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    
    Alocados = SkillPoints
    frmEstadisticas.lblLibres.Caption = SkillPoints
    
    Call frmEstadisticas.MostrarAsignacion
    
    frmEstadisticas.Iniciar_Labels
    frmEstadisticas.Show , frmMain
    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    
End Sub

Private Sub btnMapa_Click()
    
    Call frmMapa.Show(vbModeless, frmMain)
End Sub

Private Sub InvEqu_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub lblScroll_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub LbLChat_Click()
    frmMensaje.PopupMenuMensaje
End Sub

Private Sub lblMana_Click()

   Call ParseUserCommand("/MEDITAR")
End Sub

Private Sub cmdOpciones_Click()
    Call frmOpciones.Show(vbModeless, frmMain)
End Sub

Private Sub lblMinimizar_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub PicMH_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("COLOR").item(3), _
                        False, False, True)
End Sub

Private Sub MapExp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

     
     
    With charlist(UserCharIndex)
    
        If UltPos <> Index Then
        
            If UltPos >= 0 Then
                If Index = 1 Then
                    MapExp(Index).Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel)) & "%"
                    
                Else
                
                    If ClientSetup.VerLugar = 1 Then
                        MapExp(Index).Caption = mapInfo.name
                        
                    Else
                        MapExp(Index).Caption = "Posición: " & UserMap & ", " & .Pos.X & "  " & .Pos.Y
                    
                    End If

                End If
            End If
            
    
            If Index = 1 Then
                MapExp(Index).Caption = UserExp & "/" & UserPasarNivel
                
            Else

                If ClientSetup.VerLugar = 1 Then
                    MapExp(Index).Caption = mapInfo.name
                        
                Else
                    MapExp(Index).Caption = "Posición: " & UserMap & ", " & .Pos.X & "  " & .Pos.Y
                    
                End If
            End If
            
            If UserPasarNivel = 0 Then
                MapExp(Index).Caption = "¡Nivel máximo!"
            End If
                
            UltPos = Index
        End If
        
    End With
    
End Sub

Private Sub modocombate_Click()
    Call WriteCombatToggle
End Sub

Private Sub modoseguro_Click()
    Call WriteSafeToggle
End Sub

Private Sub modorol_Click()
    'Call ClientTCP.Send_Data(Role_Mode)
    'CurrentUser.Rol = Not CurrentUser.Rol
    'frmMain.modorol.Visible = Not frmMain.modorol.Visible
    'frmMain.nomodorol.Visible = Not frmMain.nomodorol.Visible
End Sub

Private Sub nomodocombate_Click()
    Call WriteCombatToggle
End Sub

Private Sub nomodorol_Click()
    'Call ClientTCP.Send_Data(Role_Mode)
    'CurrentUser.Rol = Not CurrentUser.Rol
    'frmMain.modorol.Visible = Not frmMain.modorol.Visible
    'frmMain.nomodorol.Visible = Not frmMain.nomodorol.Visible
End Sub

Private Sub nomodoseguro_Click()
    Call WriteSafeToggle
End Sub

Private Sub picMacro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim BotonDerecho As Boolean
    
    If Button = vbRightButton Then _
        BotonDerecho = True

    Call accionMacrosKey(Index + 1, BotonDerecho)
    
End Sub

Private Sub RecTxt_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    StartCheckingLinks
End Sub

Private Sub SendTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Para borrar el mensaje de fondo
    If FirstTimeChat Then
        SendTxt.Text = vbNullString
        FirstTimeChat = False
        ' Cambiamos el color de texto al original
        SendTxt.ForeColor = &HE0E0E0
    End If
    
errhandler:
    
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = vbNullString
        SendTxt.Text = vbNullString
        KeyCode = 0
        SendTxt.Visible = False
        
        If PicInv.Visible Then
            PicInv.SetFocus
        ElseIf hlst.Visible Then
            hlst.SetFocus
        End If
    End If
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()

    If UserEstado = 1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
        End With
    Else

        If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
            If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                Call WriteDrop(Inventario.SelectedItem, 1)
            Else

                If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                    If Not Comerciando Then frmCantidad.Show , frmMain
                End If
            End If
        End If
    End If
End Sub

Private Sub AgarrarItem()

    If UserEstado = 1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
        End With
    Else
        Call WritePickUp
    End If
End Sub

Private Sub UsarItem()

    If pausa Then Exit Sub
    
    If Comerciando Then Exit Sub
    
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
        Call WriteUseItem(Inventario.SelectedItem)
    End If
    
End Sub

Private Sub EquiparItem()

    If UserEstado = 1 Then
    
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
        End With
        
    Else
    
        If Comerciando Then Exit Sub
        
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
            Call WriteEquipItem(Inventario.SelectedItem)
        End If
        
    End If
End Sub

Private Sub btnLanzar_Click()
    
    If hlst.List(hlst.ListIndex) <> JsonLanguage.item("NADA").item("TEXTO") And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
            UsaMacro = True
        End If
    End If
    
End Sub

Private Sub btnLanzar_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub btnInfo_Click()
    
    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)
    End If
    
End Sub

Private Sub DespInv_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    MouseBoton = Button
    MouseShift = Shift
    
    '¿Hizo click derecho?
    If Button = 2 Then
        If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
            Call WriteAccionClick(tX, tY)
        End If
    End If
    
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub MainViewPic_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub MainViewPic_DblClick()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 12/27/2007
    '12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
    '**************************************************************
    If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteAccionClick(tX, tY)
    End If
    
End Sub

Private Sub MainViewPic_Click()

    If Cartel Then Cartel = False
    
    Dim MENSAJE_ADVERTENCIA As String
    Dim VAR_LANZANDO        As String
    
    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        
        If Not InGameArea() Then Exit Sub
        
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then

                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1

                    If CnTd = 3 Then
                        Call WriteUseSpellMacro
                        CnTd = 0
                    End If
                    UsaMacro = False
                End If

                'Invitando party
                If InvitandoParty = True Then
                    frmMain.MousePointer = vbDefault
                    Call WriteInvitarPartyClick(tX, tY)
                    InvitandoParty = False
                    Exit Sub
                End If
    
                '[/ybarra]
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        'frmMain.MousePointer = vbDefault
                        'UsingSkill = 0

                        'With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            'VAR_LANZANDO = JsonLanguage.item("PROYECTILES").item("TEXTO")
                            'MENSAJE_ADVERTENCIA = JsonLanguage.item("MENSAJE_MACRO_ADVERTENCIA").item("TEXTO")
                            'MENSAJE_ADVERTENCIA = Replace$(MENSAJE_ADVERTENCIA, "VAR_LANZADO", VAR_LANZANDO)
                            
                            'Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
                        'End With
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            'frmMain.MousePointer = vbDefault
                            'UsingSkill = 0

                            'With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                'VAR_LANZANDO = JsonLanguage.item("PROYECTILES").item("TEXTO")
                                'MENSAJE_ADVERTENCIA = JsonLanguage.item("MENSAJE_MACRO_ADVERTENCIA").item("TEXTO")
                                'MENSAJE_ADVERTENCIA = Replace$(MENSAJE_ADVERTENCIA, "VAR_LANZADO", VAR_LANZANDO)
                                
                                'Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
                            'End With
                            Exit Sub
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                'frmMain.MousePointer = vbDefault
                                'UsingSkill = 0

                                'LwK: ¿Poner aqui el bloqueo del cursor si no paso el intervalo del hechizo?
                                Exit Sub
                            End If
                        Else

                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                'frmMain.MousePointer = vbDefault
                                'UsingSkill = 0

                                'LwK: ¿Poner aqui el bloqueo del cursor si no paso el intervalo del hechizo?
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If
                    
                    If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                    
                    frmMain.MousePointer = vbDefault
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                'Call WriteRightClick(tx, tY) 'Proximamnete lo implementaremos..
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then

            If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                If MouseBoton = vbLeftButton Then
                    Call WriteWarpChar("YO", UserMap, tX, tY)
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_DblClick()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 12/27/2007
    '12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
    '**************************************************************
    If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteAccionClick(tX, tY)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X - MainViewPic.Left
    MouseY = Y - MainViewPic.Top
    
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > MainViewPic.Width Then
        MouseX = MainViewPic.Width
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewPic.Height Then
        MouseY = MainViewPic.Height
    End If
    
    LastButtonPressed.ToggleToNormal
    
    ' Disable links checking (not over consola)
    StopCheckingLinks
    
    If UltPos >= 0 Then
        If UltPos = 1 Then
            MapExp(UltPos).Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel)) & "%"
            
        Else

            If ClientSetup.VerLugar = 1 Then
                MapExp(UltPos).Caption = mapInfo.name
                
            Else
                MapExp(0).Caption = "Posición: " & UserMap & ", " & charlist(UserCharIndex).Pos.X & "  " & charlist(UserCharIndex).Pos.Y
            
            End If

        End If
        
        If UserPasarNivel = 0 Then
            frmMain.MapExp(1).Caption = "¡Nivel máximo!"
            
        End If
        
        UltPos = -1
        
    End If
    
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub lblDropGold_Click()

    Inventario.SelectGold

    If UserGLD > 0 Then
        If Not Comerciando Then frmCantidad.Show , frmMain
    End If
    
End Sub

Private Sub picInv_DblClick()
'**********************************************
'Autor: Lorwik
'Fecha: 14/07/2020
'Descripcion: DobleClick sobre el inventario
'**********************************************
    'Esta validacion es para que el juego no rompa si hacemos doble click
    If MirandoTrabajo > 0 Then Exit Sub
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    '¿Es un slot valido?
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
        Call WriteAccionInventario(Inventario.SelectedItem)
    End If
    
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Sound.Sound_Play(SND_CLICK)
End Sub

Private Sub RecTxt_Change()

    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar

    If Not Application.IsAppActive() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    
    ElseIf (Not Comerciando) And _
           (Not MirandoAsignarSkills) And _
           (Not frmMSG.Visible) And _
           (Not MirandoForo) And _
           (Not frmEstadisticas.Visible) And _
           (Not frmCantidad.Visible) And _
           (Not MirandoParty) Then

        If PicInv.Visible Then
            PicInv.SetFocus
                        
        ElseIf hlst.Visible Then
            hlst.SetFocus

        End If

    End If

End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)

    If PicInv.Visible Then
        PicInv.SetFocus
    Else
        hlst.SetFocus
    End If

End Sub

Private Sub SendTxt_Change()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 3/06/2006
    '3/06/2006: Maraxus - impedi se inserten caracteres no imprimibles
    '**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = JsonLanguage.item("MENSAJE_SOY_CHEATER").item("TEXTO")
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i         As Long
        Dim tempstr   As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))

            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0
End Sub

Private Sub CompletarEnvioMensajes()

    Select Case SendingType
        Case 1
            SendTxt.Text = vbNullString
        Case 2
            SendTxt.Text = "-"
        Case 3
            SendTxt.Text = ("\" & sndPrivateTo & " ")
        Case 4
            SendTxt.Text = "/CMSG "
        Case 5
            SendTxt.Text = "/PMSG "
        Case 6
            SendTxt.Text = "; "
    End Select
    
    stxtbuffer = SendTxt.Text
    SendTxt.SelStart = Len(SendTxt.Text)

End Sub

Private Sub Enviar_SendTxt()
    
    Dim str1 As String
    Dim str2 As String
    
    If Len(stxtbuffer) > 255 Then stxtbuffer = mid$(stxtbuffer, 1, 255)
    
    'Send text
    If Left$(stxtbuffer, 1) = "/" Then
        Call ParseUserCommand(stxtbuffer)

    'Shout
    ElseIf Left$(stxtbuffer, 1) = "-" Then
        If Right$(stxtbuffer, Len(stxtbuffer) - 1) <> vbNullString Then Call ParseUserCommand(stxtbuffer)
        SendingType = 2
        
    'Global
    ElseIf Left$(stxtbuffer, 1) = ";" Then
        If LenB(Right$(stxtbuffer, Len(stxtbuffer) - 1)) > 0 And InStr(stxtbuffer, ">") = 0 Then Call ParseUserCommand(stxtbuffer)
        SendingType = 6

    'Privado
    ElseIf Left$(stxtbuffer, 1) = "\" Then
        str1 = Right$(stxtbuffer, Len(stxtbuffer) - 1)
        str2 = ReadField(1, str1, 32)
        If LenB(str1) > 0 And InStr(str1, ">") = 0 Then Call ParseUserCommand("\" & str1)
        sndPrivateTo = str2
        SendingType = 3
                
    'Say
    Else
        If LenB(stxtbuffer) > 0 Then Call ParseUserCommand(stxtbuffer)
        SendingType = 1
    End If

    stxtbuffer = vbNullString
    SendTxt.Text = vbNullString
    SendTxt.Visible = False
    
End Sub

Private Sub AbrirMenuViewPort()
    'TODO: No usar variable de compilacion y acceder a esto desde el config.ini
    #If (ConMenuseConextuales = 1) Then

        If tX >= MinXBorder And tY >= MinYBorder And tY <= MaxYBorder And tX <= MaxXBorder Then

            If MapData(tX, tY).CharIndex > 0 Then
                If charlist(MapData(tX, tY).CharIndex).invisible = False Then
        
                    Dim m As frmMenuseFashion
                    Set m = New frmMenuseFashion
            
                    Load m
                    m.SetCallback Me
                    m.SetMenuId 1
                    m.ListaInit 2, False
            
                    If LenB(charlist(MapData(tX, tY).CharIndex).Nombre) <> 0 Then
                        m.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).Nombre, True
                    Else
                        m.ListaSetItem 0, "<NPC>", True
                    End If
                    m.ListaSetItem 1, JsonLanguage.item("COMERCIAR").item("TEXTO")
            
                    m.ListaFin
                    m.Show , Me

                End If
            End If
        End If

    #End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)

    Select Case MenuId

        Case 0 'Inventario

            Select Case Sel

                Case 0

                Case 1

                Case 2 'Tirar
                    Call TirarItem

                Case 3 'Usar

                    If MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
                        Call UsarItem
                    End If

                Case 3 'equipar
                    Call EquiparItem
            End Select
    
        Case 1 'Menu del ViewPort del engine

            Select Case Sel

                Case 0 'Nombre
                    Call WriteLeftClick(tX, tY)
        
                Case 1 'Comerciar
                    Call WriteLeftClick(tX, tY)
                    Call WriteCommerceStart
            End Select
    End Select
End Sub
 
''''''''''''''''''''''''''''''''''''''
'     WINDOWS API                    '
''''''''''''''''''''''''''''''''''''''
Private Sub Client_Connect()
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.Length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)
    
    Second.Enabled = True
    
    Select Case EstadoLogin

        Case E_MODO.CrearNuevoPJ, E_MODO.Normal
            Call Login

        Case E_MODO.Dados
            frmCrearPersonaje.Show
        
    End Select
 
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
    Dim RD     As String
    Dim Data() As Byte
    
    Client.GetData RD, vbByte, bytesTotal
    Data = StrConv(RD, vbFromUnicode)
    
    'Set data in the buffer
    Call incomingData.WriteBlock(Data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
    
End Sub

Private Sub Client_CloseSck()
    
    Debug.Print "Cerrando la conexion via API de Windows..."

    If frmMain.Visible = True Then frmMain.Visible = False
    Call ResetAllInfo
    Mod_Declaraciones.Conectando = True
    frmConnect.Show
End Sub

Private Sub Client_Error(ByVal number As Integer, _
                         Description As String, _
                         ByVal sCode As Long, _
                         ByVal Source As String, _
                         ByVal HelpFile As String, _
                         ByVal HelpContext As Long, _
                         CancelDisplay As Boolean)
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    
    frmConnect.MousePointer = 1
    
    Second.Enabled = False
 
    If Client.State <> sckClosed Then Client.CloseSck
    
    Mod_Declaraciones.Conectando = True
    frmConnect.Show
 
End Sub

Private Function InGameArea() As Boolean
'********************************************************************
'Author: NicoNZ
'Last Modification: 29/09/2019
'Checks if last click was performed within or outside the game area.
'********************************************************************
    If clicX < 0 Or clicX > frmMain.MainViewPic.Width Then Exit Function
    If clicY < 0 Or clicY > frmMain.MainViewPic.Height Then Exit Function
    
    InGameArea = True
End Function

Private Sub hlst_Click()
    
    With hlst

        .BackColor = vbBlack

    End With

End Sub

Private Sub Minimapa_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    If Button = vbRightButton Then
        Call WriteWarpChar("YO", UserMap, CByte(X - 1), CByte(Y - 1))
        Call ActualizarMiniMapa
        
    End If
    
End Sub
    'fin Incorporado ReyarB

Public Sub ActualizarMiniMapa()
    '***************************************************
    'Author: Martin Gomez (Samke)
    'Last Modify Date: 21/03/2020 (ReyarB)
    'Integrado por Reyarb
    'Se agrego campo de vision del render (Recox)
    'Ajustadas las coordenadas para centrarlo (WyroX)
    'Ajuste de coordenadas y tamaÃ±o del visor (ReyarB)
    '***************************************************
    Me.UserM.Left = UserPos.X - 2
    Me.UserM.Top = UserPos.Y - 2
    Me.MiniMapa.Refresh
End Sub

Private Sub btnRetos_Click()
    Call FrmRetos.Show(vbModeless, frmMain)
End Sub

Public Sub ControlSM(ByVal Index As Byte, ByVal Mostrar As Boolean)
    
    Select Case Index
    
        Case eSMType.sCombatMode
            If Mostrar Then
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("TEXTO"), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("COLOR").item(3), _
                                     True, False, True)
                                        
                modocombate.ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("TEXTO")
                
                frmMain.modocombate.Visible = True
                frmMain.nomodocombate.Visible = False
                
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SEGURO_COMBAT_OFF").item("TEXTO"), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_OFF").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_OFF").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_OFF").item("COLOR").item(3), _
                                     True, False, True)
                                        
                nomodocombate.ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("TEXTO")
                
                frmMain.modocombate.Visible = False
                frmMain.nomodocombate.Visible = True
                
            End If
            
            
            
        Case eSMType.sSafemode
            
            If Mostrar Then
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("TEXTO"), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("COLOR").item(3), _
                                     True, False, True)
                                        
                nomodoseguro.ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("TEXTO")
                
                frmMain.modoseguro.Visible = True
                frmMain.nomodoseguro.Visible = False
                
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("TEXTO"), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("COLOR").item(3), _
                                     True, False, True)
                                        
                modoseguro.ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("TEXTO")
                
                frmMain.modoseguro.Visible = False
                frmMain.nomodoseguro.Visible = True

            End If
        
    End Select
    
End Sub
