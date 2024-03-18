VERSION 5.00
Begin VB.Form frmBindKey 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignar acción"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBindKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Aceptar"
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
      Left            =   1800
      TabIndex        =   7
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtComandoEnvio 
      Enabled         =   0   'False
      Height          =   285
      Left            =   390
      TabIndex        =   4
      Text            =   "Comando a enviar"
      Top             =   2040
      Width           =   2655
   End
   Begin VB.OptionButton optAccion 
      Caption         =   "Usar ítem elegido"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   3135
   End
   Begin VB.OptionButton optAccion 
      Caption         =   "Equipar ítem elegido"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2720
      Width           =   3135
   End
   Begin VB.OptionButton optAccion 
      Caption         =   "Lanzar hechizo elegido"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2430
      Width           =   3135
   End
   Begin VB.OptionButton optAccion 
      Caption         =   "Enviar comando"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3240
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3240
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBindKey.frx":000C
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label lblTecla 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tecla:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "/"
      Height          =   255
      Left            =   270
      TabIndex        =   5
      Top             =   2070
      Width           =   105
   End
End
Attribute VB_Name = "frmBindKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmBindKey - ImperiumAO - v1.3.0
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Autor: Lorwik
'Fecha: 07/03/2021
'*****************************************************************

Option Explicit

Private Sub cmdAccept_Click()

    Dim SpellSlot As Integer
    Dim InvSlot As Byte
    Dim Comando As String

    If optAccion(0).value Then
    
        MacrosKey(MacroElegido).TipoAccion = 1
        SpellSlot = 0
        InvSlot = 0
        Comando = txtComandoEnvio.Text
    
    ElseIf optAccion(1).value Then
        MacrosKey(MacroElegido).TipoAccion = 2
        SpellSlot = frmMain.hlst.ListIndex + 1
        InvSlot = 0
        Comando = ""
    
    ElseIf optAccion(2).value Then
    
        MacrosKey(MacroElegido).TipoAccion = 3
        SpellSlot = 0
        InvSlot = Inventario.SelectedItem
        Comando = ""
    
    ElseIf optAccion(3).value Then
    
        MacrosKey(MacroElegido).TipoAccion = 4
        SpellSlot = 0
        InvSlot = Inventario.SelectedItem
        Comando = ""
    
    End If

    Call WriteGuardarMacro(MacroElegido, SpellSlot, InvSlot, Comando)
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    lblTecla.Caption = "Tecla: F" & MacroElegido
    
    '¿Estamos editando?
    If MacrosKey(MacroElegido).TipoAccion <> 0 Then
    
        Select Case MacrosKey(MacroElegido).TipoAccion
        
            Case 1 'Comandos
                optAccion(0).value = True
                txtComandoEnvio.Text = MacrosKey(MacroElegido).Comando
                txtComandoEnvio.Enabled = True
            
            Case 2 'Hechizos
                optAccion(1).value = True
            
            Case 3 'Equipar Item
                optAccion(2).value = True
            
            Case 4 'User Item
                optAccion(3).value = True
        
        End Select
    
    End If
    
End Sub

Private Sub optAccion_Click(Index As Integer)
    If Index = 0 Then
        txtComandoEnvio.Enabled = True
    Else
        txtComandoEnvio.Enabled = False
    End If
End Sub
