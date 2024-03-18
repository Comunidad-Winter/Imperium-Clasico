VERSION 5.00
Begin VB.Form frmConfirmacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   214
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   268
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgCancelar 
      Height          =   525
      Left            =   240
      Top             =   2595
      Width           =   1695
   End
   Begin VB.Image imgAceptar 
      Height          =   525
      Left            =   2040
      Top             =   2595
      Width           =   1695
   End
   Begin VB.Label msg 
      BackStyle       =   0  'Transparent
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
      Height          =   1995
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3495
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmConfirmacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonAceptar As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    ' TODO: Traducir los textos de las imagenes via labels en visual basic, para que en el futuro si se quiere se pueda traducir a mas idiomas
    ' No ando con mas ganas/tiempo para hacer eso asi que se traducen las imagenes asi tenemos el juego en ingles.
    ' Tambien usar los controles uAObuttons para los botones, usar de ejemplo frmCambiaMotd.frm
    Me.Picture = General_Load_Picture_From_Resource("124.bmp", False)
    
    Call LoadButtons
End Sub

Private Sub Form_Deactivate()
    Me.SetFocus
End Sub

Private Sub LoadButtons()
    Dim boton As String
    
   ' GrhPath = Carga.path(Interfaces)

    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    
    If Language = "spanish" Then
        boton = "btnaceptar.bmp"
    Else
        boton = "btnaccept.bmp"
    End If
    
    Call cBotonAceptar.Initialize(imgAceptar, Carga.Path(Interfaces) & boton, _
                                     Carga.Path(Interfaces) & "123.bmp", _
                                     Carga.Path(Interfaces) & "122.bmp", Me)
                                     
                                     
    If Language = "spanish" Then
        boton = "btncancelar.bmp"
    Else
        boton = "btncancel.bmp"
    End If
                                     
    Call cBotonCancelar.Initialize(imgCancelar, Carga.Path(Interfaces) & boton, _
                                     Carga.Path(Interfaces) & "125.bmp", _
                                     Carga.Path(Interfaces) & "126.bmp", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
    msg.Caption = "" 'Limpiamos el caption
    Unload Me
End Sub

Private Sub imgAceptar_Click()
    Unload Me
End Sub

Private Sub imgCancelar_Click()
    Unload Me
End Sub

Private Sub msg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

