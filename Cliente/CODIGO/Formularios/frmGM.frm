VERSION 5.00
Begin VB.Form frmGM 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Formulario de mensaje a administradores"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   ClipControls    =   0   'False
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
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   311
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CMDEnviar 
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   5160
      Width           =   4215
   End
   Begin VB.OptionButton OptDenuncia 
      Caption         =   "Acusacion"
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
      Height          =   195
      Index           =   3
      Left            =   1980
      TabIndex        =   5
      Top             =   1410
      Width           =   1095
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Sugerencia"
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
      Height          =   195
      Index           =   2
      Left            =   1980
      TabIndex        =   3
      Top             =   1650
      Width           =   1335
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Reporte de bug"
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
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1650
      Width           =   1575
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Consulta Regular"
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
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   1410
      Value           =   -1  'True
      Width           =   1755
   End
   Begin VB.TextBox TXTMessage 
      BackColor       =   &H00FFFFFF&
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
      Height          =   2055
      Left            =   225
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2220
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGM.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   4320
      Width           =   4215
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGM.frx":009D
      ForeColor       =   &H00000000&
      Height          =   1245
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'By Lorwik
'Form encargado de enviar el mensaje al GM :)
'************************************************************

Private Sub CMDEnviar_Click()

    Call Sound.Sound_Play(SND_CLICK)
    
    If TXTMessage.Text = "" Then
        MsgBox "Debes de escribir el motivo de tu consulta."
        Exit Sub
    End If
    
    If optConsulta(0).value = True Then
        Call WriteGMRequest(0, TXTMessage.Text)
        Unload Me
        Exit Sub
        
    ElseIf optConsulta(1).value = True Then
        Call WriteGMRequest(1, TXTMessage.Text)
        Unload Me
        Exit Sub
        
    ElseIf optConsulta(2).value = True Then
        Call WriteGMRequest(2, TXTMessage.Text)
        Unload Me
        Exit Sub
        
    ElseIf optConsulta(3).value = True Then
        Call WriteGMRequest(3, TXTMessage.Text)
        Unload Me
        Exit Sub
    End If
    
    
End Sub

Private Sub Form_Load()
    lblInfo.Caption = JsonLanguage.item("VENTANAGM").item("SOPORTE")
End Sub

Private Sub optConsulta_Click(Index As Integer)
    Select Case Index
        Case 0
            lblInfo.Caption = JsonLanguage.item("VENTANAGM").item("SOPORTE")
        Case 1
            lblInfo.Caption = JsonLanguage.item("VENTANAGM").item("BUG")
        Case 2
            lblInfo.Caption = JsonLanguage.item("VENTANAGM").item("SUGERENCIA")
        Case 3
            lblInfo.Caption = JsonLanguage.item("VENTANAGM").item("DENUNCIA")
    End Select

End Sub

Private Sub CMDVolver_Click()
    Unload Me
End Sub

