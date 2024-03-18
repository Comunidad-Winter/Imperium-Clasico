VERSION 5.00
Begin VB.Form frmParty 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Grupo"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton imgSalirParty 
      Caption         =   "Abandonar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3180
      Width           =   3375
   End
   Begin VB.CommandButton imgAgregar 
      Caption         =   "Invitar"
      Height          =   390
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2730
      Width           =   1650
   End
   Begin VB.CommandButton imgExpulsar 
      Caption         =   "Expulsar"
      Height          =   390
      Left            =   1860
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2730
      Width           =   1635
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   3600
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2730
      Width           =   1230
   End
   Begin VB.ListBox lstMembers 
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
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4710
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmParty.frx":0000
      Height          =   1050
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   4635
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmParty.frm
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
On Error GoTo Form_Load_Err
    
    lstMembers.Clear

    MirandoParty = True
    Exit Sub

Form_Load_Err:
    MsgBox Err.Description & vbNewLine & _
               "in ARGENTUM.frmParty.Form_Load " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MirandoParty = False
End Sub

Private Sub imgAgregar_Click()
    Call WriteRequestPartyForm(True)
    'Call WritePartyAcceptMember
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub imgExpulsar_Click()
  
    If lstMembers.ListIndex < 0 Then Exit Sub

    Dim fName As String
    fName = GetName

    If Len(fName) > 0 Then
        Call WritePartyKick(fName)
        Unload Me
        
        ' Para que no llame al form si disolvio la party
        If UCase$(fName) <> UCase$(UserName) Then Call WriteRequestPartyForm
    End If

End Sub

Private Function GetName() As String
'**************************************************************
'Author: ZaMa
'Last Modify Date: 27/12/2009
'**************************************************************
    Dim sName As String
    
    sName = Trim$(mid$(lstMembers.List(lstMembers.ListIndex), 1, InStr(lstMembers.List(lstMembers.ListIndex), " (")))
    If Len(sName) > 0 Then GetName = sName
        
End Function

Private Sub imgSalirParty_Click()
    Call WritePartyLeave
    Unload Me
End Sub
