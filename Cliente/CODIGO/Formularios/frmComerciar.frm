VERSION 5.00
Begin VB.Form frmComerciar 
   BorderStyle     =   0  'None
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   510
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Cantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Text            =   "1"
      Top             =   6960
      Width           =   510
   End
   Begin VB.PictureBox picInvUser 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3960
      Left            =   3720
      ScaleHeight     =   264
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   166
      TabIndex        =   4
      Top             =   2580
      Width           =   2490
   End
   Begin VB.PictureBox picInvNpc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3960
      Left            =   780
      ScaleHeight     =   264
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   164
      TabIndex        =   3
      Top             =   2580
      Width           =   2460
   End
   Begin VB.Image cmdMasMenos 
      Height          =   420
      Index           =   1
      Left            =   3840
      Tag             =   "1"
      Top             =   6870
      Width           =   195
   End
   Begin VB.Image cmdMasMenos 
      Height          =   420
      Index           =   0
      Left            =   2940
      Tag             =   "1"
      Top             =   6870
      Width           =   195
   End
   Begin VB.Image imgVender 
      Height          =   450
      Left            =   4230
      Top             =   6855
      Width           =   2175
   End
   Begin VB.Image imgComprar 
      Height          =   450
      Left            =   585
      Top             =   6855
      Width           =   2175
   End
   Begin VB.Image imgCross 
      Height          =   405
      Left            =   6480
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   120
      Width           =   345
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   5100
      TabIndex        =   1
      Top             =   1890
      Width           =   1035
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   1530
      Width           =   2985
   End
End
Attribute VB_Name = "frmComerciar"
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

Public LasActionBuy As Boolean
Private ClickNpcInv As Boolean

Private cBotonCruz As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub cantidad_Change()
    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
    
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = MAX_INVENTORY_OBJS
    End If
    
    If ClickNpcInv Then
        If InvComNpc.SelectedItem <> 0 Then
            'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
            Label1(1).Caption = CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).valor, Val(cantidad.Text))   'No mostramos numeros reales
        End If
    Else
        If InvComUsu.SelectedItem <> 0 Then
            Label1(1).Caption = CalculateBuyPrice(Inventario.valor(InvComUsu.SelectedItem), Val(cantidad.Text))  'No mostramos numeros reales
        End If
    End If
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub cmdMasMenos_Click(Index As Integer)

    Call Sound.Sound_Play(SND_CLICK)

    Select Case Index
        Case 0
            cantidad.Text = str((Val(cantidad.Text) - 1))
        
        Case 1
            cantidad.Text = str((Val(cantidad.Text) + 1))
    End Select
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    'Cargamos la interfase
    Me.Picture = General_Load_Picture_From_Resource("comercio.bmp")
    
    Call LoadButtons
End Sub

Private Sub Form_Activate()
On Error Resume Next

    Call InvComUsu.DrawInventory
    Call InvComNpc.DrawInventory

End Sub

Private Sub Form_GotFocus()
On Error Resume Next

    Call InvComUsu.DrawInventory
    Call InvComNpc.DrawInventory

End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    GrhPath = Carga.Path(Interfaces)
    
    'Lo dejamos solo para que no explote, habria que sacar estos LastButtonPressed
    Set LastButtonPressed = New clsGraphicalButton

    Set cBotonCruz = New clsGraphicalButton
    
    'Call cBotonCruz.Initialize(imgCross, "", _
                                    "171.bmp", _
                                    "171.bmp", Me)

End Sub

''
' Calculates the selling price of an item (The price that a merchant will sell you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.

Private Function CalculateSellPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo Error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateSellPrice = CCur(objValue * 1000000) / 1000000 * objAmount + 0.5
    
    Exit Function
Error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.number
End Function
''
' Calculates the buying price of an item (The price that a merchant will buy you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.
Private Function CalculateBuyPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo Error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateBuyPrice = Fix(CCur(objValue * 1000000) / 1000000 * objAmount)
    
    Exit Function
Error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.number
End Function

Private Sub imgComprar_Click()
    ' Debe tener seleccionado un item para comprarlo.
    If InvComNpc.SelectedItem = 0 Then Exit Sub
    
    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
    
    Call Sound.Sound_Play(SND_CLICK)
    
    LasActionBuy = True
    If UserGLD >= CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).valor, Val(cantidad.Text)) Then
        Call WriteCommerceBuy(InvComNpc.SelectedItem, Val(cantidad.Text))
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SIN_ORO_SUFICIENTE").item("TEXTO"), 2, 51, 223, 1, 1)
        Exit Sub
    End If

    Call InvComUsu.DrawInventory
    Call InvComNpc.DrawInventory
    
End Sub

Private Sub imgCross_Click()
    Call WriteCommerceEnd
End Sub

Private Sub imgVender_Click()
    ' Debe tener seleccionado un item para comprarlo.
    If InvComUsu.SelectedItem = 0 Then Exit Sub

    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
    
    Call Sound.Sound_Play(SND_CLICK)
    
    LasActionBuy = False

    Call WriteCommerceSell(InvComUsu.SelectedItem, Val(cantidad.Text))

    Call InvComUsu.DrawInventory
    Call InvComNpc.DrawInventory
    
End Sub

Private Sub picInvNpc_Click()
    Dim ItemSlot As Byte
    
    ItemSlot = InvComNpc.SelectedItem
    If ItemSlot = 0 Then Exit Sub
    
    ClickNpcInv = True
    InvComUsu.DeselectItem
    
    Label1(0).Caption = NPCInventory(ItemSlot).name
    Label1(1).Caption = CalculateSellPrice(NPCInventory(ItemSlot).valor, Val(cantidad.Text))  'No mostramos numeros reales
    
    If NPCInventory(ItemSlot).Amount <> 0 Then
    
        Select Case NPCInventory(ItemSlot).OBJType
            Case eObjType.otWeapon
                Label1(2).Caption = JsonLanguage.item("GOLPE").item("TEXTO") & ":" & NPCInventory(ItemSlot).MaxHit & "/" & JsonLanguage.item("GOLPE").item("TEXTO") & ":" & NPCInventory(ItemSlot).MinHit
                Label1(2).Visible = True
            Case eObjType.otArmadura, eObjType.otcasco, eObjType.otescudo
                Label1(2).Caption = JsonLanguage.item("DEFENSA").item("TEXTO") & ":" & NPCInventory(ItemSlot).MaxDef & "/" & JsonLanguage.item("DEFENSA").item("TEXTO") & ":" & NPCInventory(ItemSlot).MinDef
                Label1(2).Visible = True
            Case Else
                Label1(2).Visible = False
        End Select
    Else
        Label1(2).Visible = False
    End If
End Sub

Private Sub picInvUser_Click()
    Dim ItemSlot As Byte
    
    ItemSlot = InvComUsu.SelectedItem
    
    If ItemSlot = 0 Then Exit Sub
    
    ClickNpcInv = False
    InvComNpc.DeselectItem
    
    Label1(0).Caption = Inventario.ItemName(ItemSlot)
    Label1(1).Caption = "$: " & CalculateBuyPrice(Inventario.valor(ItemSlot), Val(cantidad.Text)) 'No mostramos numeros reales
    
    If Inventario.Amount(ItemSlot) <> 0 Then
    
        Select Case Inventario.OBJType(ItemSlot)
            Case eObjType.otWeapon, eObjType.otFlechas
                Label1(2).Caption = JsonLanguage.item("GOLPE").item("TEXTO") & ":" & Inventario.MaxHit(ItemSlot) & "/" & JsonLanguage.item("GOLPE").item("TEXTO") & ":" & Inventario.MinHit(ItemSlot)
                Label1(2).Visible = True
            Case eObjType.otArmadura, eObjType.otcasco, eObjType.otescudo
                Label1(2).Caption = JsonLanguage.item("DEFENSA").item("TEXTO") & ":" & Inventario.MaxDef(ItemSlot) & "/" & JsonLanguage.item("DEFENSA").item("TEXTO") & ":" & Inventario.MinDef(ItemSlot)
                Label1(2).Visible = True
            Case Else
                Label1(2).Visible = False
        End Select
    Else
        Label1(2).Visible = False
    End If
End Sub

