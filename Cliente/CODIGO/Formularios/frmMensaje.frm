VERSION 5.00
Begin VB.Form frmMensaje 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   3915
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
   ScaleHeight     =   178
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   261
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgCerrar 
      Height          =   465
      Left            =   1365
      Tag             =   "1"
      Top             =   2070
      Width           =   1200
   End
   Begin VB.Label msg 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   300
      TabIndex        =   0
      Top             =   480
      Width           =   3315
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuMensaje 
      Caption         =   "Mensaje"
      Visible         =   0   'False
      Begin VB.Menu mnuNormal 
         Caption         =   "Normal"
      End
      Begin VB.Menu mnuGritar 
         Caption         =   "Gritar"
      End
      Begin VB.Menu mnuPrivado 
         Caption         =   "Privado"
      End
      Begin VB.Menu mnuClan 
         Caption         =   "Clan"
      End
      Begin VB.Menu mnuGrupo 
         Caption         =   "Grupo"
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "Global"
      End
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private clsFormulario As clsFormMovementManager

Private cBotonCerrar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Deactivate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    ' TODO: Traducir los textos de las imagenes via labels en visual basic, para que en el futuro si se quiere se pueda traducir a mas idiomas
    ' No ando con mas ganas/tiempo para hacer eso asi que se traducen las imagenes asi tenemos el juego en ingles.
    ' Tambien usar los controles uAObuttons para los botones, usar de ejemplo frmCambiaMotd.frm
    Me.Picture = General_Load_Picture_From_Resource("info.bmp")
    
    Call LoadButtons
End Sub

Private Sub LoadButtons()

   ' GrhPath = Carga.path(Interfaces)

    Set cBotonCerrar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    If Language = "spanish" Then

        Call cBotonCerrar.Initialize(imgCerrar, "3.bmp", _
                                          "213.bmp", _
                                          "214.bmp", Me)
    Else
    
        Call cBotonCerrar.Initialize(imgCerrar, "5.bmp", _
                                          "211.bmp", _
                                          "212.bmp", Me)
        
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
    msg.Caption = "" 'Limpiamos el caption
    Unload Me
End Sub

Private Sub msg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Public Sub PopupMenuMensaje()
    
    Select Case SendingType
        Case 1
            mnuNormal.Checked = True
            mnuGritar.Checked = False
            mnuPrivado.Checked = False
            mnuClan.Checked = False
            mnuGrupo.Checked = False
            mnuGlobal.Checked = False
            
        Case 2
            mnuNormal.Checked = False
            mnuGritar.Checked = True
            mnuPrivado.Checked = False
            mnuClan.Checked = False
            mnuGrupo.Checked = False
            mnuGlobal.Checked = False
            
        Case 3
            mnuNormal.Checked = False
            mnuGritar.Checked = False
            mnuPrivado.Checked = True
            mnuClan.Checked = False
            mnuGrupo.Checked = False
            mnuGlobal.Checked = False
            
        Case 4
            mnuNormal.Checked = False
            mnuGritar.Checked = False
            mnuPrivado.Checked = False
            mnuClan.Checked = True
            mnuGrupo.Checked = False
            mnuGlobal.Checked = False
            
        Case 5
            mnuNormal.Checked = False
            mnuGritar.Checked = False
            mnuPrivado.Checked = False
            mnuClan.Checked = False
            mnuGrupo.Checked = True
            mnuGlobal.Checked = False
            
        Case 6
            mnuNormal.Checked = False
            mnuGritar.Checked = False
            mnuPrivado.Checked = False
            mnuClan.Checked = False
            mnuGrupo.Checked = False
            mnuGlobal.Checked = True
            
    End Select
    
    PopupMenu mnuMensaje
    
End Sub

'[Lorwik]
'Moví este menú acá para que se pueda ver el caption del
'frmMain sin que se tenga que ver el ControlBox

Private Sub mnuNormal_Click()

    SendingType = 1
    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuGritar_click()

    SendingType = 2
    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuPrivado_click()

    sndPrivateTo = InputBox("Nombre del destinatario:", vbNullString)
    
    If sndPrivateTo <> vbNullString Then
        SendingType = 3
        If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
    Else
        MsgBox "¡Escribe un nombre."
    End If

End Sub

Private Sub mnuClan_click()

    SendingType = 4
    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuGrupo_click()

    SendingType = 5
    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuGlobal_Click()

    SendingType = 6
    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub
