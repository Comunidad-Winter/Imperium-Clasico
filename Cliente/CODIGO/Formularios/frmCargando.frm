VERSION 5.00
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgProgress 
      Height          =   570
      Left            =   2250
      Top             =   8070
      Width           =   7500
   End
End
Attribute VB_Name = "frmCargando"
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

Public NoInternetConnection As Boolean

Private porcentajeActual As Integer
 
Private Const PROGRESS_DELAY = 10
Private Const PROGRESS_DELAY_BACKWARDS = 4
Private Const DEFAULT_PROGRESS_WIDTH = 500
Private Const DEFAULT_STEP_FORWARD = 1
Private Const DEFAULT_STEP_BACKWARDS = -3

Private Sub Form_Load()
    Me.Analizar
    Me.Picture = General_Load_Picture_From_Resource("cargando.bmp")
    imgProgress.Picture = General_Load_Picture_From_Resource("barra.bmp")
    
    ' Seteamos el caption
    Me.Caption = Form_Caption
End Sub

Private Sub LOGO_KeyPress(KeyAscii As Integer)
    Debug.Print 2
End Sub

Private Sub Status_KeyPress(KeyAscii As Integer)
    Debug.Print 1
End Sub

Function Analizar()
On Error Resume Next
    Dim binaryFileToOpen As String

    If NoInternetConnection Then
        MsgBox "No hay conexion a internet, verificar que tengas internet/No Internet connection, please verify"
        Exit Function
    End If
           
End Function

Private Sub progresoConDelay(ByVal porcentaje As Integer)
 
    If porcentaje = porcentajeActual Then Exit Sub
     
    Dim step As Integer, stepInterval As Integer, Timer As Long, tickCount As Long
     
    If (porcentaje > porcentajeActual) Then
        step = DEFAULT_STEP_FORWARD
        stepInterval = PROGRESS_DELAY
    Else
        step = DEFAULT_STEP_BACKWARDS
        stepInterval = PROGRESS_DELAY_BACKWARDS
    End If
     
    Do Until compararPorcentaje(porcentaje, porcentajeActual, step)
        Do Until (Timer + stepInterval) <= GetTickCount()
            DoEvents
        Loop
        Timer = GetTickCount()
        porcentajeActual = porcentajeActual + step
        Call establecerProgreso(porcentajeActual)
    Loop
 
End Sub
 
Private Sub establecerProgreso(ByVal nuevoPorcentaje As Integer)
 
    If nuevoPorcentaje >= 0 And nuevoPorcentaje <= 100 Then
        imgProgress.Width = DEFAULT_PROGRESS_WIDTH * CLng(nuevoPorcentaje) / 100
    ElseIf nuevoPorcentaje > 100 Then
        imgProgress.Width = DEFAULT_PROGRESS_WIDTH
    Else
        imgProgress.Width = 0
    End If
    porcentajeActual = nuevoPorcentaje
 
End Sub
 
Private Function compararPorcentaje(ByVal porcentajeTarget As Integer, ByVal porcentajeAct As Integer, ByVal step As Integer) As Boolean
 
    If step = DEFAULT_STEP_FORWARD Then
        compararPorcentaje = (porcentajeAct >= porcentajeTarget)
    Else
        compararPorcentaje = (porcentajeAct <= porcentajeTarget)
    End If
 
End Function

Public Sub ActualizarCarga(ByVal Mensaje As String, ByVal Progreso As Byte)
'***********************************************
'Autor: Lorwik
'Fecha: 13/07/2020
'Descripcion: Actualiza el progreso de carga
'***********************************************

    Call LogError(0, Mensaje, "Iniciando")
    Call progresoConDelay(Progreso)
End Sub

