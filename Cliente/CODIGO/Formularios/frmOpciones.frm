VERSION 5.00
Begin VB.Form frmOpciones 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6885
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
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraInformación 
      Caption         =   "Información"
      Height          =   1335
      Left            =   120
      TabIndex        =   23
      Top             =   4290
      Width           =   3255
      Begin VB.CommandButton cmdNecesitásAyuda 
         Caption         =   "¿Necesitás &ayuda?"
         Height          =   345
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton cmdManual 
         Caption         =   "&www.imperiumclasico.com.ar"
         Height          =   345
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.CommandButton imgSalir 
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   120
      TabIndex        =   22
      Top             =   6090
      Width           =   3255
   End
   Begin VB.CommandButton imgConfigTeclas 
      Caption         =   "Configuración de controles"
      Height          =   360
      Left            =   120
      TabIndex        =   21
      Top             =   5700
      Width           =   3255
   End
   Begin VB.Frame FraVideo 
      Caption         =   "Apariencia y performance"
      Height          =   2865
      Left            =   3480
      TabIndex        =   20
      Top             =   120
      Width           =   3285
      Begin VB.CheckBox chkVerNombre 
         Caption         =   "Ver nombre del mapa"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Audio"
      ForeColor       =   &H00000000&
      Height          =   4065
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3255
      Begin VB.HScrollBar scrMusic 
         Height          =   315
         LargeChange     =   15
         Left            =   120
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   16
         Top             =   3570
         Width           =   2895
      End
      Begin VB.HScrollBar scrVolume 
         Height          =   315
         LargeChange     =   15
         Left            =   120
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   15
         Top             =   2400
         Width           =   2895
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Sonido habilitado"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   660
         Width           =   2985
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Música habilitada"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2985
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Sonidos Ambientales habilitado"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2985
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Invertir canales de audio (L / R)"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1530
         Width           =   2985
      End
      Begin VB.HScrollBar scrAmbient 
         Height          =   315
         LargeChange     =   15
         Left            =   120
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   10
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de música"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   19
         Top             =   3360
         Width           =   2865
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de sonidos ambientales"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   18
         Top             =   2790
         Width           =   2865
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de audio"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   2190
         Width           =   2835
      End
   End
   Begin VB.Frame FraSkins 
      Caption         =   "General"
      Height          =   3375
      Left            =   3480
      TabIndex        =   0
      Top             =   3090
      Width           =   3285
      Begin VB.ComboBox cmdLenguajesComboBox 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2640
         Width           =   1815
      End
      Begin VB.ComboBox cmdSkinsComboBox 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmOpciones.frx":0152
         Left            =   240
         List            =   "frmOpciones.frx":0154
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox ComMouseHechizos 
         Height          =   315
         ItemData        =   "frmOpciones.frx":0156
         Left            =   240
         List            =   "frmOpciones.frx":0160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1920
         Width           =   1815
      End
      Begin VB.ComboBox ComMouseGeneral 
         Height          =   315
         ItemData        =   "frmOpciones.frx":0179
         Left            =   240
         List            =   "frmOpciones.frx":0186
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblLenguaje 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lenguaje"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   660
      End
      Begin VB.Label lblSkinDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skin de Interfaces"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label lblMouseHechizos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Grafico de Hechizos"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   1920
      End
      Begin VB.Label lblMouseGrafico 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Grafico General"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmOpciones"
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

Private bMusicActivated As Boolean
Private bSoundActivated As Boolean
Private bSoundEffectsActivated As Boolean

Private loading As Boolean

Private Sub chkop_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Lorwik
'Fecha: 30/05/2020
'Descripcion: Activa/Desactiva los sonidos
'***************************************************

    Call Sound.Sound_Play(SND_CLICK)

    Select Case Index
    
        Case 0 'Musica
                    
            If chkop(Index).value = vbUnchecked Then
                Sound.Music_Stop
                ClientSetup.bMusic = CONST_DESHABILITADA
                scrMusic.Enabled = False
            Else
                ClientSetup.bMusic = CONST_MP3
                scrMusic.Enabled = True
            End If
        
        Case 1 'Sonido
    
            If chkop(Index).value = vbUnchecked Then
                'scrAmbient.Enabled = False
                scrVolume.Enabled = False
                ClientSetup.bSound = 0
            Else
                ClientSetup.bSound = 1
                scrVolume.Enabled = True
            End If
            
        Case 2 'Ambiente
            
            If chkop(Index).value = vbUnchecked Then
                ClientSetup.bAmbient = 0
                Call Sound.Sound_Stop_All
            Else
                ClientSetup.bAmbient = 1
                scrAmbient.Enabled = True
                Call Sound.Ambient_Load(Sound.AmbienteActual, ClientSetup.AmbientVol)
                Call Sound.Ambient_Play
            End If
            
        Case 5 'FPS
            ClientSetup.LimiteFPS = Not ClientSetup.LimiteFPS
            
    End Select
End Sub

Private Sub chkVerNombre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Lorwik
'Fecha: 31/03/2021
'Descripcion: Activa/Desactiva el nombre del mapa
'***************************************************

    If ClientSetup.VerLugar = 0 Then
        ClientSetup.VerLugar = 1
        frmMain.MapExp(0).Caption = mapInfo.name
        
    Else
        ClientSetup.VerLugar = 0
        frmMain.MapExp(0).Caption = "Posición: " & UserMap & ", " & charlist(UserCharIndex).Pos.X & "  " & charlist(UserCharIndex).Pos.Y
        
    End If

    chkVerNombre.value = ClientSetup.VerLugar
End Sub

Private Sub cmdManual_Click()
    If Not loading Then _
        Call Sound.Sound_Play(SND_CLICK)
    Call ShellExecute(0, "Open", "http://imperiumclasico.com.ar/wiki/", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub scrMusic_Change()
'***************************************************
'Author: Lorwik
'Fecha: 30/05/2020
'Descripcion: Setea el volumen de la musica
'***************************************************

    If ClientSetup.bMusic <> CONST_DESHABILITADA Then
        Sound.Music_Volume_Set scrMusic.value
        Sound.VolumenActualMusicMax = scrMusic.value
        ClientSetup.MusicVolume = Sound.VolumenActualMusicMax
    End If

End Sub

Private Sub scrAmbient_Change()
'***************************************************
'Author: Lorwik
'Fecha: 30/05/2020
'Descripcion: Setea el volumen del sonido ambiente
'***************************************************

    If ClientSetup.bAmbient = 1 Then
        Sound.VolumenActualAmbient_set scrAmbient.value
        ClientSetup.AmbientVol = Sound.VolumenActualAmbient
    End If
    
End Sub

Private Sub scrVolume_Change()
'***************************************************
'Author: Lorwik
'Fecha: 30/05/2020
'Descripcion: Setea el volumen de los sonidos
'***************************************************

    If ClientSetup.bSound = 1 Then
        Sound.VolumenActual = scrVolume.value
        ClientSetup.SoundVolume = Sound.VolumenActual
    End If

End Sub

Private Sub cmdLenguajesComboBox_Click()
'***************************************************
'Author: Recox
'Last Modification: 01/04/2019
'10/11/2019: Recox - Seteamos el lenguaje del juego
'***************************************************
    Call WriteVar(Carga.Path(Init) & CLIENT_FILE, "Parameters", "Language", cmdLenguajesComboBox.Text)
    MsgBox ("Debe reiniciar el juego aplicar el cambio de idioma. Idioma Seleccionado: " & cmdLenguajesComboBox.Text)
End Sub

Private Sub cmdSkinsComboBox_Click()
'***************************************************
'Author: Recox
'Last Modification: 01/04/2019
'08/11/2019: Recox - Seteamos el skin
'***************************************************
    Call WriteVar(Carga.Path(Init) & CLIENT_FILE, "Parameters", "SkinSelected", cmdSkinsComboBox.Text)
    MsgBox ("Debe reiniciar el juego aplicar el cambio de skin. Skin Seleccionado: " & cmdSkinsComboBox.Text)
End Sub

Private Sub ComMouseGeneral_Click()
'***************************************************
'Author: Lorwik
'Last Modification: 26/04/2020
'26/04/2020: Lorwik - Seteamos el mouse general
'***************************************************
    Call WriteVar(Carga.Path(Init) & CLIENT_FILE, "Parameters", "MOUSEGENERAL", ComMouseGeneral.ListIndex)
    MsgBox ("Debe reiniciar el juego aplicar el cambio de mouse. Mouse Seleccionado: " & ComMouseGeneral.Text)
End Sub

Private Sub ComMouseHechizos_Click()
'***************************************************
'Author: Lorwik
'Last Modification: 26/04/2020
'26/04/2020: Lorwik - Seteamos el mouse baston
'***************************************************
    Call WriteVar(Carga.Path(Init) & CLIENT_FILE, "Parameters", "MOUSEBASTON", ComMouseHechizos.ListIndex)
    MsgBox ("Debe reiniciar el juego aplicar el cambio de mouse. Mouse Seleccionado: " & ComMouseHechizos.Text)
End Sub

Private Sub imgConfigTeclas_Click()
    If Not loading Then _
        Call Sound.Sound_Play(SND_CLICK)
    Call frmCustomKeys.Show(vbModal, Me)
End Sub

Private Sub imgSalir_Click()

    Call Carga.GuardarConfiguracion
    
    Unload Me
    
    If frmMain.Visible = True Then
        frmMain.SetFocus
    ElseIf frmConnect.Visible = True Then
        frmConnect.SetFocus
    End If
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    LoadSkinsInComboBox
    LoadLenguajesInComboBox
    
    loading = True      'Prevent sounds when setting check's values
    LoadUserConfig
    loading = False     'Enable sounds when setting check's values
End Sub

Private Sub LoadSkinsInComboBox()
    Dim sFileName As String
    sFileName = Dir$(Carga.Path(skins), vbArchive)
    
    Do While sFileName > vbNullString
        Call cmdSkinsComboBox.AddItem(sFileName)
        sFileName = Dir$()
    Loop
    
    'Boorramos los 2 primeros items por que son . y ..
    Call cmdSkinsComboBox.RemoveItem(0)
'    Call cmdSkinsComboBox.RemoveItem(0)
End Sub

Private Sub LoadLenguajesInComboBox()
    Dim sFileName As String
    sFileName = Dir$(App.Path & "\Lenguajes\", vbArchive)
    
    Do While sFileName > vbNullString
        sFileName = Replace(sFileName, ".json", vbNullString)
        Call cmdLenguajesComboBox.AddItem(sFileName)
        sFileName = Dir$()
    Loop

End Sub

Private Sub LoadUserConfig()

    'Musica
    If ClientSetup.bMusic = CONST_DESHABILITADA Then
        chkop(0).value = 0
        scrMusic.value = ClientSetup.MusicVolume
    Else
        chkop(0).value = 1
        scrMusic.value = ClientSetup.MusicVolume
    End If
    
    'Sonidos
    If ClientSetup.bSound = 1 Then
        chkop(1).value = vbChecked
        chkop(3).value = IIf(ClientSetup.Invertido = True, 1, 0)
        scrVolume.value = ClientSetup.SoundVolume
    Else
        chkop(1).value = vbUnchecked
        chkop(3).value = IIf(ClientSetup.Invertido = True, 1, 0)
        chkop(3).Enabled = False
        scrVolume.value = ClientSetup.SoundVolume
        scrVolume.Enabled = False
    End If
    
    'Ambiente
    If ClientSetup.bAmbient = 1 Then
        chkop(3).value = vbChecked
        scrAmbient.value = ClientSetup.AmbientVol
    Else
        chkop(3).value = vbUnchecked
        scrAmbient.value = ClientSetup.AmbientVol
    End If
    
    If ClientSetup.LimiteFPS = True Then
        'chkop(5).value = vbChecked
    Else
        'chkop(5).value = vbUnchecked
    End If

End Sub
