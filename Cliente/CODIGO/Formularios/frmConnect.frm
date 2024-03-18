VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   5.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ListBox lst_servers 
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
      Height          =   2175
      ItemData        =   "frmConnect.frx":000C
      Left            =   6510
      List            =   "frmConnect.frx":0013
      TabIndex        =   2
      Top             =   1845
      Width           =   3675
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2715
      Width           =   2355
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   0
      Top             =   1830
      Width           =   4215
   End
   Begin SHDocVwCtl.WebBrowser noticias 
      Height          =   3375
      Left            =   1800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4920
      Width           =   8370
      ExtentX         =   14764
      ExtentY         =   5953
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Image imgRecargar 
      Height          =   375
      Left            =   9960
      Top             =   1440
      Width           =   255
   End
   Begin VB.Image imgMinimizar 
      Height          =   300
      Left            =   11325
      Top             =   60
      Width           =   300
   End
   Begin VB.Image imgCerrar 
      Height          =   300
      Left            =   11640
      Top             =   60
      Width           =   300
   End
   Begin VB.Image imgRecuperar 
      Height          =   810
      Left            =   3975
      MousePointer    =   99  'Custom
      Top             =   3300
      Width           =   2070
   End
   Begin VB.Image imgRegistrar 
      Height          =   810
      Left            =   1770
      MousePointer    =   99  'Custom
      Top             =   3300
      Width           =   2070
   End
   Begin VB.Image imgConectar 
      Height          =   615
      Left            =   4335
      MousePointer    =   99  'Custom
      Top             =   2445
      Width           =   1755
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Lector As clsIniManager

Private cBotonConectar     As clsGraphicalButton
Private cBotonRegistrar    As clsGraphicalButton
Private cBotonRecuperar    As clsGraphicalButton
Private cBotonCerrar       As clsGraphicalButton
Private cBotonMinimizar    As clsGraphicalButton

Public LastButtonPressed   As clsGraphicalButton

Private Sub imgCerrar_Click()
    Call CloseClient
End Sub

Private Sub imgConectar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    Call WriteVar(Carga.Path(Init) & CLIENT_FILE, "PARAMETERS", "SERVERSELECT", ServIndSel)
        
    'update user info
    AccountName = frmConnect.txtNombre.Text
    AccountPassword = frmConnect.txtPasswd.Text
        
    'Clear spell list
    frmMain.hlst.Clear
        
    If CheckUserData() = True Then
        Call Protocol.Connect(E_MODO.Normal)
    End If
    
End Sub

Private Sub imgMinimizar_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub imgRecargar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call ListarServidores
End Sub

Private Sub imgRecuperar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call ShellExecute(0, "Open", "http://imperiumclasico.com.ar/", "", App.Path, SW_SHOWNORMAL)

End Sub

Private Sub imgRegistrar_Click()
    Call ShellExecute(0, "Open", "http://imperiumclasico.com.ar/", "", App.Path, SW_SHOWNORMAL)
    Call Sound.Sound_Play(SND_CLICK)
End Sub

Private Sub lst_servers_Click()
    ServIndSel = lst_servers.ListIndex
    CurServerIp = Servidor(ServIndSel).Ip
    CurServerPort = Servidor(ServIndSel).Puerto
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
  '  If KeyAscii = vbKeyReturn Then btnConectarse_Click
End Sub

Private Sub Form_Load()
    ' Seteamos el caption
    Me.Caption = Form_Caption
    
    Me.Picture = General_Load_Picture_From_Resource("conectar.bmp")
    
    Call noticias.Navigate("http://winterao.com.ar/impc/noticias.html")
    
    ServIndSel = GetVar(Carga.Path(Init) & CLIENT_FILE, "PARAMETERS", "SERVERSELECT")
        
    Call LoadButtons
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Call CloseClient
    End If

End Sub

Private Sub LoadButtons()

    Set LastButtonPressed = New clsGraphicalButton
    
    Set cBotonConectar = New clsGraphicalButton
    Set cBotonRegistrar = New clsGraphicalButton
    Set cBotonRecuperar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonMinimizar = New clsGraphicalButton
    
    imgCerrar.MouseIcon = picMouseIcon
    imgMinimizar.MouseIcon = picMouseIcon
    
    Call cBotonConectar.Initialize(imgConectar, "", _
                                 "botconectarover.bmp", _
                                 "botconectardown.bmp", Me)
                                 
    Call cBotonRegistrar.Initialize(imgRegistrar, "", _
                                 "botcrearover.bmp", _
                                 "botcreardown.bmp", Me)
                                 
    Call cBotonRecuperar.Initialize(imgRecuperar, "", _
                                 "botrecuperarover.bmp", _
                                 "botrecuperardown.bmp", Me)
                                 
    Call cBotonCerrar.Initialize(imgCerrar, "", _
                                 "conceover.bmp", _
                                 "concedown.bmp", Me)
                                 
    Call cBotonMinimizar.Initialize(imgMinimizar, "", _
                                 "conminover.bmp", _
                                 "conmindown.bmp", Me)
                                 
                                 
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LastButtonPressed.ToggleToNormal
End Sub
