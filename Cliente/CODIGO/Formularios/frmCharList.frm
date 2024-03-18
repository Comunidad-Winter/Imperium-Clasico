VERSION 5.00
Begin VB.Form frmCharList 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "ImperiumClasico"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCharList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   9
      Left            =   8445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   9
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   8
      Left            =   6945
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   8
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   7
      Left            =   5445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   6
      Left            =   3945
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   5
      Left            =   2445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   4
      Left            =   8445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   3
      Left            =   6945
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   3
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   2
      Left            =   5445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   1
      Left            =   3945
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   0
      Left            =   2445
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   3930
      Width           =   1140
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clase"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   6210
      TabIndex        =   23
      Top             =   7920
      Width           =   390
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   6210
      TabIndex        =   22
      Top             =   7770
      Width           =   675
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   6210
      TabIndex        =   21
      Top             =   7620
      Width           =   1605
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
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
      Left            =   8340
      TabIndex        =   20
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
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
      Left            =   6840
      TabIndex        =   19
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
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
      Left            =   5325
      TabIndex        =   18
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
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
      Left            =   3840
      TabIndex        =   17
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
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
      Left            =   2325
      TabIndex        =   16
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
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
      Left            =   8340
      TabIndex        =   15
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
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
      Left            =   6840
      TabIndex        =   14
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
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
      Left            =   5325
      TabIndex        =   13
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
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
      Left            =   3840
      TabIndex        =   12
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
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
      Left            =   2325
      TabIndex        =   11
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
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
      Left            =   2280
      TabIndex        =   10
      Top             =   2370
      Width           =   3705
   End
   Begin VB.Image imgAccion 
      Height          =   300
      Index           =   6
      Left            =   11640
      Tag             =   "0"
      Top             =   60
      Width           =   300
   End
   Begin VB.Image imgAccion 
      Height          =   300
      Index           =   5
      Left            =   11325
      Tag             =   "0"
      Top             =   60
      Width           =   300
   End
   Begin VB.Image WebLink 
      Height          =   345
      Left            =   8490
      MousePointer    =   99  'Custom
      Top             =   8550
      Width           =   3405
   End
   Begin VB.Image imgConectar 
      Height          =   615
      Left            =   8025
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image imgBorrarPj 
      Height          =   615
      Left            =   4155
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image imgSalir 
      Height          =   615
      Left            =   8025
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   2085
      Width           =   1755
   End
   Begin VB.Image imgCrearPJ 
      Height          =   615
      Left            =   2235
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image imgCambiarPass 
      Height          =   615
      Left            =   6180
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   2085
      Width           =   1755
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   9
      Left            =   8280
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   8
      Left            =   6780
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   7
      Left            =   5280
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   6
      Left            =   3780
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   5
      Left            =   2280
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   4
      Left            =   8280
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   3
      Left            =   6780
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   2
      Left            =   5280
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   1
      Left            =   3780
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   0
      Left            =   2280
      Top             =   3510
      Width           =   1455
   End
End
Attribute VB_Name = "frmCharList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cBotonConectar     As clsGraphicalButton
Private cBotonCrearPJ      As clsGraphicalButton
Private cBotonSalir        As clsGraphicalButton
Private cBotonBorrar       As clsGraphicalButton
Private cBotonCambiarPass  As clsGraphicalButton

Public LastButtonPressed   As clsGraphicalButton

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Call imgSalir_Click
    End If

End Sub

Private Sub Form_Load()

    Dim i As Long

    Unload frmConnect
    
    Me.Picture = General_Load_Picture_From_Resource("cuenta.bmp")
    Me.Icon = frmMain.Icon
    ' Seteamos el caption
    Me.Caption = Form_Caption

    For i = 1 To 10
        lblAccData(i).Caption = vbNullString
    Next i

    Me.lblAccData(0).Caption = AccountName
    
    Call LoadButtons
    
End Sub

Private Sub imgBorrarPj_Click()
    If PJAccSelected < 1 Then
        Call MostrarMensaje(JsonLanguage.item("ERROR_PERSONAJE_NO_SELECCIONADO").item("TEXTO"))
        Exit Sub
    End If
                            
    frmBorrarPJ.Show
End Sub

Private Sub imgCambiarPass_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call ShellExecute(0, "Open", "http://imperiumclasico.com.ar/", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub imgConectar_Click()
    If Not frmMain.Client.State = sckConnected Then
        MsgBox JsonLanguage.item("ERROR_CONN_LOST").item("TEXTO")
        frmConnect.Show
                
    Else
        If Mod_Declaraciones.Conectando Then
            Mod_Declaraciones.Conectando = False
            Call WriteLoginExistingChar
                    
            DoEvents
            
            Call FlushBuffer
        End If
    End If
End Sub

Private Sub imgCrearPJ_Click()
    If NumberOfCharacters > 9 Then
        MsgBox JsonLanguage.item("ERROR_DEMASIADOS_PJS").item("TEXTO")
        Exit Sub
    End If
            
    If ClientSetup.bMusic <> CONST_DESHABILITADA Then
        If ClientSetup.bMusic <> CONST_DESHABILITADA Then
            Sound.NextMusic = MUS_CrearPersonaje
            Sound.Fading = 500
        End If
    End If
            
    Dim LoopC As Long
        
    For LoopC = 1 To 10
        If LenB(lblAccData(LoopC).Caption) = 0 Then
            frmCrearPersonaje.Show
            Exit Sub
        End If
    Next LoopC
End Sub

Private Sub imgSalir_Click()
    frmMain.Client.CloseSck
    Call ResetAllInfoAccounts
    Call ListarServidores
    frmConnect.Visible = True
    
    Unload Me
End Sub

Private Sub picChar_Click(Index As Integer)
    On Error Resume Next
    
    If LenB(cPJ(Index + 1).Nombre) <> 0 Then
        'El PJ seleccionado queda guardado
        UserName = cPJ(Index + 1).Nombre
        PJAccSelected = Index + 1
        
        lblCharData(0).Caption = "Nivel: " & cPJ(Index + 1).Level
        lblCharData(1).Caption = "Mapa: " & cPJ(Index + 1).Map
        lblCharData(2).Caption = "Clase: " & cPJ(Index + 1).Class
    End If

End Sub

Private Sub picChar_DblClick(Index As Integer)

    Call imgConectar_Click

End Sub

Private Sub WebLink_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call ShellExecute(0, "Open", "http://imperiumclasico.com.ar/", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub LoadButtons()

    Set LastButtonPressed = New clsGraphicalButton
    
    Set cBotonConectar = New clsGraphicalButton
    Set cBotonCrearPJ = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    Set cBotonBorrar = New clsGraphicalButton
    Set cBotonCambiarPass = New clsGraphicalButton
    
    Call cBotonCambiarPass.Initialize(imgCambiarPass, "", _
                                 "acccambiarover.bmp", _
                                 "acccambiardown.bmp", Me)
                                 
    Call cBotonCrearPJ.Initialize(imgCrearPJ, "", _
                                 "acccreover.bmp", _
                                 "acccredown.bmp", Me)
                                 
    Call cBotonSalir.Initialize(imgSalir, "", _
                                 "accsaover.bmp", _
                                 "accsaldown.bmp", Me)
                                 
    Call cBotonBorrar.Initialize(imgBorrarPj, "", _
                                 "accborrarover.bmp", _
                                 "accborrardown.bmp", Me)
                                 
    Call cBotonConectar.Initialize(imgConectar, "", _
                                 "accconover.bmp", _
                                 "acccondown.bmp", Me)
                                 
                                 
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LastButtonPressed.ToggleToNormal
End Sub

