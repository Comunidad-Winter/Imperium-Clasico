VERSION 5.00
Begin VB.Form frmTutorial 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bienvenido"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton imgAnterior 
      Caption         =   "Anterior"
      Height          =   480
      Left            =   360
      TabIndex        =   6
      Top             =   7080
      Width           =   2970
   End
   Begin VB.CommandButton imgSiguiente 
      Caption         =   "Siguiente"
      Height          =   600
      Left            =   5640
      TabIndex        =   5
      Top             =   6960
      Width           =   2850
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   525
      TabIndex        =   4
      Top             =   435
      Width           =   7725
   End
   Begin VB.Label lblMensaje 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5790
      Left            =   525
      TabIndex        =   3
      Top             =   840
      Width           =   7725
   End
   Begin VB.Label lblPagTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7365
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblPagActual 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6870
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   8430
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   75
      Width           =   255
   End
End
Attribute VB_Name = "frmTutorial"
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

Private Type tTutorial
    sTitle As String
    sPage As String
End Type

Private Tutorial() As tTutorial
Private NumPages As Long
Private CurrentPage As Long

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = General_Load_Picture_From_Resource("187.bmp", False)

    Call LoadTextForms
    Call LoadButtons
    Call LoadTutorial
    
    CurrentPage = 1
    Call SelectPage(CurrentPage)
End Sub

Private Sub LoadTextForms()
    imgSiguiente.Caption = JsonLanguage.item("FRM_TUTORIAL_SIGUIENTE").item("TEXTO")
    imgAnterior.Caption = JsonLanguage.item("FRM_TUTORIAL_ANTERIOR").item("TEXTO")
End Sub

Private Sub LoadButtons()
    imgAnterior.Enabled = False
    lblCerrar.MouseIcon = picMouseIcon
End Sub


Private Sub imgAnterior_Click()

    If Not imgAnterior.Enabled Then Exit Sub
    
    CurrentPage = CurrentPage - 1
    
    If CurrentPage = 1 Then imgAnterior.Enabled = False
    
    If Not imgSiguiente.Enabled Then imgSiguiente.Enabled = True
    
    Call SelectPage(CurrentPage)
End Sub

Private Sub imgSiguiente_Click()
    
    If Not imgSiguiente.Enabled Then Exit Sub
    
    CurrentPage = CurrentPage + 1
    
    ' Si paso de la ultima pagina, cierra
    If CurrentPage > NumPages Then
        imgSiguiente.Caption = "Cerrar"
        bShowTutorial = False 'Mientras no se pueda tildar/destildar para verlo mas tarde, esto queda asi :P
        Unload Me
        
    Else
        imgSiguiente.Caption = "Siguiente"
        
    End If
    
    ' Habilita el boton anterior
    If Not imgAnterior.Enabled Then imgAnterior.Enabled = True
    
    Call SelectPage(CurrentPage)
End Sub

Private Sub LoadTutorial()
    
    Dim TutorialPath As String
    Dim lPage As Long
    Dim NumLines As Long
    Dim lLine As Long
    Dim sLine As String
    
    ' Obtenemos el lenguage en ingles o castellano mediante la variable global del modLenguage
    TutorialPath = Carga.Path(Lenguajes) & "Tutorial_" & Language & ".dat"
    NumPages = Val(GetVar(TutorialPath, "INIT", "NumPags"))
    
    If NumPages > 0 Then
        ReDim Tutorial(1 To NumPages)
        
        ' Cargo paginas
        For lPage = 1 To NumPages
            NumLines = Val(GetVar(TutorialPath, "PAG" & lPage, "NumLines"))
            
            With Tutorial(lPage)
                
                .sTitle = GetVar(TutorialPath, "PAG" & lPage, "Title")
                
                ' Cargo cada linea de la pagina
                For lLine = 1 To NumLines
                    sLine = GetVar(TutorialPath, "PAG" & lPage, "Line" & lLine)
                    .sPage = .sPage & sLine & vbNewLine
                Next lLine
            End With
            
        Next lPage
    End If
    
    lblPagTotal.Caption = NumPages
End Sub

Private Sub SelectPage(ByVal lPage As Long)
    lblTitulo.Caption = Tutorial(lPage).sTitle
    lblMensaje.Caption = Tutorial(lPage).sPage
    lblPagActual.Caption = lPage
End Sub

