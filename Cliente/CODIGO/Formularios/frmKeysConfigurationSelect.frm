VERSION 5.00
Begin VB.Form frmKeysConfigurationSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuracion Controles / Config Keys"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ImpC_Client.uAOButton btnNormalKeys 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   7560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      TX              =   ""
      ENAB            =   -1  'True
      FCOL            =   16777215
      OCOL            =   16777215
      PICE            =   "frmKeysConfigurationSelect.frx":0000
      PICF            =   "frmKeysConfigurationSelect.frx":001C
      PICH            =   "frmKeysConfigurationSelect.frx":0038
      PICV            =   "frmKeysConfigurationSelect.frx":0054
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ImpC_Client.uAOButton btnAlternativeKeys 
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   7560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      TX              =   ""
      ENAB            =   -1  'True
      FCOL            =   16777215
      OCOL            =   16777215
      PICE            =   "frmKeysConfigurationSelect.frx":0070
      PICF            =   "frmKeysConfigurationSelect.frx":008C
      PICH            =   "frmKeysConfigurationSelect.frx":00A8
      PICV            =   "frmKeysConfigurationSelect.frx":00C4
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblAlternativeTitle 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo Alternativo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   5640
      TabIndex        =   6
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label lblNormalTitle 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo Normal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   4560
      X2              =   4560
      Y1              =   1800
      Y2              =   8040
   End
   Begin VB.Label lblAlternativeText 
      BackStyle       =   0  'Transparent
      Caption         =   "Legacy (Directional Arrows) "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1695
      Left            =   5520
      TabIndex        =   4
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label lblNormalText 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1455
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Image imgAlternativeKeyboard 
      Height          =   1335
      Left            =   5400
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Image imgNormalKeyboard 
      Height          =   1335
      Left            =   240
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmKeysConfigurationSelect.frx":00E0
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1335
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "frmKeysConfigurationSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   Call LoadTextsForm
   imgAlternativeKeyboard.Picture = General_Load_Picture_From_Resource("19.bmp", False)
   imgNormalKeyboard.Picture = General_Load_Picture_From_Resource("20.bmp", False)
End Sub

Private Sub LoadTextsForm()
   lblAlternativeText.Caption = JsonLanguage.item("FRM_KEYS_CONFIGURATION_LBL_ALTERNATIVE_TEXT").item("TEXTO")
   lblAlternativeTitle.Caption = JsonLanguage.item("FRM_KEYS_CONFIGURATION_LBL_ALTERNATIVE_TITLE").item("TEXTO")
   lblNormalText.Caption = JsonLanguage.item("FRM_KEYS_CONFIGURATION_LBL_NORMAL_TEXT").item("TEXTO")
   lblNormalTitle.Caption = JsonLanguage.item("FRM_KEYS_CONFIGURATION_LBL_NORMAL_TITLE").item("TEXTO")
   lblTitle.Caption = JsonLanguage.item("FRM_KEYS_CONFIGURATION_LBL_TITLE").item("TEXTO")
   btnNormalKeys.Caption = JsonLanguage.item("FRM_KEYS_CONFIGURATION_BTN_NORMAL_KEYS").item("TEXTO")
   btnAlternativeKeys.Caption = JsonLanguage.item("FRM_KEYS_CONFIGURATION_BTN_ALTERNATIVE_KEYS").item("TEXTO")
End Sub

Private Sub btnAlternativeKeys_Click()
   Call CustomKeys.LoadDefaults(False)
   SetFalseMostrarBindKeysSelection
   Unload Me
End Sub

Private Sub btnNormalKeys_Click()
   Call CustomKeys.LoadDefaults(True)
   SetFalseMostrarBindKeysSelection
   Unload Me
End Sub

Private Sub SetFalseMostrarBindKeysSelection()
   Call WriteVar(Carga.Path(Init) & CLIENT_FILE, "OTHER", "MOSTRAR_BIND_KEYS_SELECTION", "False")
End Sub

