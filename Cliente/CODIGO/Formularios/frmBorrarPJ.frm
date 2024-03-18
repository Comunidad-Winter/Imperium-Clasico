VERSION 5.00
Begin VB.Form frmBorrarPJ 
   BorderStyle     =   0  'None
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   110
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtconfirmacion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   780
      Width           =   5295
   End
   Begin VB.Image cmdBorrar 
      Height          =   540
      Left            =   3480
      Top             =   1050
      Width           =   1725
   End
   Begin VB.Image cmdVolver 
      Height          =   540
      Left            =   840
      Top             =   1050
      Width           =   1725
   End
   Begin VB.Label lblATENCIÓNESTAS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¡ATENCIÓN ESTAS A PUNTO DE BORRAR UN PERSONAJE!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   5220
   End
   Begin VB.Label lblEstasSeguro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Escribe ""BORRAR XXXXX"" para eliminarlo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   3960
   End
End
Attribute VB_Name = "frmBorrarPJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonCancelar As clsGraphicalButton
Private cBotonBorrar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub cmdBorrar_Click()
'*************************************
'Autor: Lorwik
'Fecha: 21/05/2020
'Descripcion: Preguntamos si desea eliminar el PJ y en caso de afirmacion mandamos eliminar
'*************************************

    '¿Paso la verificacion?
    If CheckBorrarData = False Then Exit Sub

    Select Case MsgBox(JsonLanguage.item("FRMPANELACCOUNT_CONFIRMAR_BORRAR_PJ").item("TEXTO"), vbYesNo + vbExclamation, JsonLanguage.item("FRMPANELACCOUNT_CONFIRMAR_BORRAR_PJ_TITULO").item("TEXTO"))
    
        Case vbYes
            
            Call WriteDeleteChar
            
            'Salimos del form
            Call CMDVolver_Click
            
        Case vbNo
            CMDVolver_Click
            Exit Sub
            
    End Select
End Sub

Private Sub CMDVolver_Click()
'*************************************
'Autor: Lorwik
'Fecha: 21/05/2020
'Descripcion: Reseteamos y salimos del form
'*************************************

    lblEstasSeguro.Caption = vbNullString
    
    Unload Me
End Sub

Private Function CheckBorrarData() As Boolean
'*************************************
'Autor: Lorwik
'Fecha: 21/05/2020
'Descripcion: Checkeamos antes de borrar
'*************************************

    '¿El indice es correcto?
    If PJAccSelected < 1 Or PJAccSelected > 10 Then
        Call MostrarMensaje("Error al borrar el PJ. Intentelo de nuevo o contacte con un Administrador.")
        CheckBorrarData = False
        Exit Function
    End If
    
    'Escribio la palabra magica?
    If Not txtconfirmacion.Text = "BORRAR " & cPJ(PJAccSelected).Nombre Then
        Call MostrarMensaje("Escribe BORRAR " & cPJ(PJAccSelected).Nombre & " para eliminar el personaje.")
        CheckBorrarData = False
        Exit Function
    End If
    
    CheckBorrarData = True
End Function

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    lblEstasSeguro.Caption = "Escribe 'BORRAR " & cPJ(PJAccSelected).Nombre & "' para eliminarlo"
    
    Me.Picture = General_Load_Picture_From_Resource("215.bmp", False)
    
    Call LoadButtons
    
End Sub

Private Sub LoadButtons()

   ' GrhPath = Carga.path(Interfaces)

    Set cBotonCancelar = New clsGraphicalButton
    Set cBotonBorrar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    If Language = "spanish" Then

        Call cBotonCancelar.Initialize(cmdVolver, "4.bmp", _
                                          "216.bmp", _
                                          "217.bmp", Me)
                                          
        Call cBotonBorrar.Initialize(cmdBorrar, "3.bmp", _
                                          "213.bmp", _
                                          "214.bmp", Me)
    Else
    
        Call cBotonCancelar.Initialize(cmdVolver, "6.bmp", _
                                          "218.bmp", _
                                          "219.bmp", Me)
                                          
        Call cBotonBorrar.Initialize(cmdBorrar, "5.bmp", _
                                          "211.bmp", _
                                          "212.bmp", Me)
        
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub


