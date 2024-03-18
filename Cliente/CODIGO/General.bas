Attribute VB_Name = "Mod_General"
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
Public bFogata As Boolean

Private lFrameTimer As Long

Private m_Jpeg             As clsJpeg
Private m_FileName         As String

Private keysMovementPressedQueue As clsArrayList

'Remove Title Bar
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Function GetRawName(ByRef sName As String) As String
'***************************************************
'Author: ZaMa
'Last Modify Date: 13/01/2010
'Last Modified By: -
'Returns the char name without the clan name (if it has it).
'***************************************************

    Dim Pos As Integer
    
    Pos = InStr(1, sName, "<")
    
    If Pos > 0 Then
        GetRawName = Trim$(Left$(sName, Pos - 1))
    Else
        GetRawName = sName
    End If

End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, _
                    ByVal Text As String, _
                    Optional ByVal Red As Integer = -1, _
                    Optional ByVal Green As Integer, _
                    Optional ByVal Blue As Integer, _
                    Optional ByVal bold As Boolean = False, _
                    Optional ByVal italic As Boolean = False, _
                    Optional ByVal bCrLf As Boolean = True, _
                    Optional ByVal Alignment As Byte = rtfLeft)
    
'****************************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D apperance!
'****************************************************
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martin Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'Jopi 17/08/2019 : Consola transparente.
'Jopi 17/08/2019 : Ahora podes especificar el alineamiento del texto.
'****************************************************
    With RichTextBox
        
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        ' 0 = Left
        ' 1 = Center
        ' 2 = Right
        .SelAlignment = Alignment

        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        
        .SelText = Text

        ' Esto arregla el bug de las letras superponiendose la consola del frmMain
        If Not RichTextBox = frmMain.RecTxt Then RichTextBox.Refresh

    End With
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    Dim Len_cad As Long
    
    cad = LCase$(cad)
    Len_cad = Len(cad)
    
    For i = 1 To Len_cad
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("Âº")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData() As Boolean
    
    'Validamos los datos del user
    
    Dim LoopC As Long
    Dim CharAscii As Integer
    Dim Len_accountName As Long, Len_accountPassword As Long
    
    If LenB(AccountPassword) = 0 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_PASSWORD").item("TEXTO"))
        Exit Function
    End If
    
    Len_accountPassword = Len(AccountPassword)
    
    For LoopC = 1 To Len_accountPassword
        CharAscii = Asc(mid$(AccountPassword, LoopC, 1))
        If Not LegalCharacter(CharAscii) Then
            Call MostrarMensaje(Replace$(JsonLanguage.item("VALIDACION_BAD_PASSWORD").item("TEXTO").item(2), "VAR_CHAR_INVALIDO", Chr$(CharAscii)))
            Exit Function
        End If
    Next LoopC

    If Not AsciiValidos(AccountName) Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_BAD_ACCOUNTNAME").item("TEXTO").item(1))
        Exit Function
    End If

    If LenB(AccountName) = 0 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_BAD_ACCOUNTNAME").item("TEXTO").item(1))
        Exit Function
    End If

    If Len(AccountName) > 24 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_BAD_ACCOUNTNAME").item("TEXTO").item(2))
        Exit Function
    End If
        
    Len_accountName = Len(AccountName)
    
    For LoopC = 1 To Len_accountName
        CharAscii = Asc(mid$(AccountName, LoopC, 1))
        If Not LegalCharacter(CharAscii) Then
            Call MostrarMensaje(Replace$(JsonLanguage.item("VALIDACION_BAD_PASSWORD").item("TEXTO").item(4), "VAR_CHAR_INVALIDO", Chr$(CharAscii)))
            Exit Function
        End If
    Next LoopC
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True

    'Unload the connect form
    Unload frmCrearPersonaje
    Unload frmConnect
    Unload frmCharList
    
    'Vaciamos la cola de movimiento
    keysMovementPressedQueue.Clear

    frmMain.lblName.Caption = UserName
    
    'Load main form
    frmMain.Visible = True
    
    Call DibujarMenuMacros
    Call Time_Logic(HoraActual)
    
    Mod_Declaraciones.Conectando = True

    FPSFLAG = True

End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
    Call Map_MoveTo(RandomNumber(SOUTH, EAST))
End Sub

Private Sub AddMovementToKeysMovementPressedQueue()
    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyUp)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyUp)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyUp)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyUp)) ' Remueve la tecla que teniamos presionada
    End If

    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyDown)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyDown)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyDown)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyDown)) ' Remueve la tecla que teniamos presionada
    End If

    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyLeft)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyLeft)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyLeft)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyLeft)) ' Remueve la tecla que teniamos presionada
    End If

    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyRight)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyRight)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyRight)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyRight)) ' Remueve la tecla que teniamos presionada
    End If
End Sub

Private Sub CheckKeys()
     '*****************************************************************
    'Checks keys and respond
    '*****************************************************************
    Static lastmovement As Long

    'No input allowed while Argentum is not the active window
    If Not Application.IsAppActive() Then Exit Sub
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking while writting in the forum.
    If MirandoForo Then Exit Sub
    'If game is paused, abort movement.
    If pausa Then Exit Sub

    'Si esta chateando, no mover el pj, tanto para chat de clanes y normal
    If frmMain.SendTxt.Visible And ClientSetup.BloqueoMovimiento Then Exit Sub

    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            Call AddMovementToKeysMovementPressedQueue

            'Move Up
            If keysMovementPressedQueue.GetLastItem() = CustomKeys.BindedKey(eKeyType.mKeyUp) Then
                Call Map_MoveTo(NORTH)
                Call Char_UserPos
                Exit Sub
            End If
            
            'Move Right
            If keysMovementPressedQueue.GetLastItem() = CustomKeys.BindedKey(eKeyType.mKeyRight) Then
                Call Map_MoveTo(EAST)
                Call Char_UserPos
                Exit Sub
            End If
        
            'Move down
            If keysMovementPressedQueue.GetLastItem() = CustomKeys.BindedKey(eKeyType.mKeyDown) Then
                Call Map_MoveTo(SOUTH)
                Call Char_UserPos
                Exit Sub
            End If
        
            'Move left
            If keysMovementPressedQueue.GetLastItem() = CustomKeys.BindedKey(eKeyType.mKeyLeft) Then
                Call Map_MoveTo(WEST)
                Call Char_UserPos
                Exit Sub
            End If
           
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            End If

            Call Char_UserPos
        End If

    End If
    
End Sub

Sub SwitchMap(ByVal Map As Integer)
'**********************************************************************************
'Autor: Lorwik
'Fecha: 11/06/2020
'Descripción: Intentamos descomprimir el mapa, si existe lo cargamos
'**********************************************************************************
    
    Dim bytArr()    As Byte
    Dim InfoHead    As INFOHEADER
    Dim Musica      As String
    
    'Reseteamos el Array antes que nada, o por la velocidad que pueda tardar en comprobar si el mapa existe, se pueden _
    producir errores.
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize)
    
    InfoHead = File_Find(Carga.Path(ePath.recursos) & "\Mapas" & Formato, LCase$("Mapa" & Map & ".csm"))
    
    If InfoHead.lngFileSize <> 0 Then

        'Limpieza adicional del mapa. PARCHE: Solucion a bug de clones. [Gracias Yhunja]
        'EDIT: cambio el rango de valores en x y para solucionar otro bug con respecto al cambio de mapas
        Call Char_CleanAll
        
        'Borramos las particulas de lluvia
        Call mDx8_Clima.RemoveWeatherParticlesAll
        
        'Borramos las particulas activas en el mapa.
        Call Particle_Group_Remove_All
        
        'Borramos todas las luces
        Call LightRemoveAll
        
        Musica = mapInfo.Music
        
        'Cargamos el mapa.
        Call Carga.CargarMapa(Map)
        
        Call DibujarMinimapa
        
        If ClientSetup.VerLugar = 1 Then frmMain.MapExp(0).Caption = mapInfo.name
        
        CurMap = Map
        
        Call Init_Ambient(Map)
        Call InfoMapa
        
        'Si estamos jugando y no en el conectar...
        If frmMain.Visible Then
            'Resetear el mensaje en render con el nombre del mapa.
            renderText = mapInfo.name
            renderFont = 2
            colorRender = 240

            If Val(Sound.MusicActual) <> Val(mapInfo.Music) Then
                'Reproducimos la música del mapa
                If ClientSetup.bMusic <> CONST_DESHABILITADA Then
                    If ClientSetup.bMusic <> CONST_DESHABILITADA Then
                        Sound.NextMusic = Val(mapInfo.Music)
                        Sound.Fading = 500
                    End If
                End If
            End If
        End If
    Else
    
        'no encontramos el mapa en el hd
        Call MsgBox(JsonLanguage.item("ERROR_MAPAS").item("TEXTO"))
        
        Call CloseClient
        
    End If
    
End Sub

Public Sub InfoMapa()
    If InfoMapAct = True Then
        frmMain.MapExp(0).Caption = "Posición: " & UserMap & ", " & UserPos.X & ", " & UserPos.Y
    Else
        If Not MapName = "" Then
            frmMain.MapExp(0).Caption = Trim$(MapName)
        Else
            frmMain.MapExp(0).Caption = "Mapa Desconocido"
        End If
    End If
End Sub
Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(File, FileType) <> "")
End Function

Sub Main()
    Static lastFlush As Long
    ' Detecta el idioma del sistema (TRUE) y carga las traducciones
    Call SetLanguageApplication

    Call GenerateContra
    
    'Load client configurations.
    Call Carga.LeerConfiguracion

    #If Desarrollo = 0 Then
        If GetVar(Carga.Path(Init) & CLIENT_FILE, "PARAMETERS", "LAUCH") <> 1 Then
            Call MsgBox("Para iniciar Imperium Clasico debes hacerlo desde el Launcher.", vbCritical)
            End
        Else
            Call WriteVar(Carga.Path(Init) & CLIENT_FILE, "PARAMETERS", "LAUCH", "0")
        End If
    
        If Application.FindPreviousInstance Then
            Call MsgBox(JsonLanguage.item("OTRO_CLIENTE_ABIERTO").item("TEXTO"), vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
            End
        End If
    #End If
    
    'Read command line. Do it AFTER config file is loaded to prevent this from
    'canceling the effects of "/nores" option.
    Call LeerLineaComandos
    
    'usaremos esto para ayudar en los parches
    Call SaveSetting("ImperiumClasico", "Init", "Path", App.Path & "\")
    
    ChDrive App.Path
    ChDir App.Path
    Windows_Temp_Dir = General_Get_Temp_Dir
    Form_Caption = "ImperiumClasico v" & App.Major & "." & App.Minor & "." & App.Revision
    
    'Set resolution BEFORE the loading form is displayed, therefore it will be centered.
    Call Resolution.SetResolution(800, 600)

    ' Load constants, classes, flags, graphics..
    Call LoadInitialConfig
    
    'Busca y pre-selecciona un server.
    Call ListarServidores
    ServIndSel = 0
    
    '¿Actualizaciones para el Launcher?
    If FileExist(App.Path & "\ImperiumClasicoLauncher.exe.up", vbNormal) Then
    
        If FileExist(App.Path & "\ImperiumClasicoLauncher.exe", vbNormal) Then Kill App.Path & "\ImperiumClasicoLauncher.exe"
        Name "ImperiumClasicoLauncher.exe.up" As "ImperiumClasicoLauncher.exe"
        
        MsgBox "Se ha encontrado una actualización del Launcher. Imperium Clasico se reiniciara."
        Call Shell(App.Path & "\ImperiumClasicoLauncher.exe", vbNormalFocus)
        Call CloseClient
    
    End If
  
    frmConnect.Visible = True
    
    If ClientSetup.bMusic <> CONST_DESHABILITADA Then
        Sound.NextMusic = MUS_Inicio
        Sound.Fading = 350
        Sound.Sound_Render
    End If
    
    'Inicializacion de variables globales
    prgRun = True
    pausa = False
    
    ' Intervals
    LoadTimerIntervals
 
    Dialogos.Font = frmMain.Font
    
    lFrameTimer = GetTickCount

    Do While prgRun

        'Solo dibujamos si la ventana no esta minimizada
        If frmMain.WindowState <> vbMinimized And frmMain.Visible Then
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            
            Call CheckKeys
            
        End If
        
        If (ClientSetup.bSound = 1 Or ClientSetup.bMusic <> CONST_DESHABILITADA) Then Call Sound.Sound_Render
        
        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            
            lFrameTimer = GetTickCount
        End If
        
        If timeGetTime >= lastFlush Then
            ' If there is anything to be sent, we send it
            Call FlushBuffer
            lastFlush = timeGetTime + 10
        End If
        DoEvents
    Loop
    
    Call CloseClient
End Sub

Public Function GetVersionOfTheGame() As String
    GetVersionOfTheGame = GetVar(Carga.Path(Init) & CLIENT_FILE, "Cliente", "VersionTagRelease")
End Function

Private Sub LoadInitialConfig()
'***************************************************
'Author: Recox
'Last Modification: 30/10/2019
'15/03/2011: ZaMa - Initialize classes lazy way.
'30/10/2019: Recox - Initialize Mouse icons
'***************************************************
    
    'Cargamos los graficos de mouse guardados
    ClientSetup.MouseGeneral = Val(GetVar(Carga.Path(Init) & CLIENT_FILE, "PARAMETERS", "MOUSEGENERAL"))
    ClientSetup.MouseBaston = Val(GetVar(Carga.Path(Init) & CLIENT_FILE, "PARAMETERS", "MOUSEBASTON"))
    
    'Si es 0 cargamos el por defecto
    If ClientSetup.MouseBaston > 0 Then
        ' Mouse Pointer and Mouse Icon (Loaded before opening any form with buttons in it)
        Set picMouseIcon = LoadPicture(App.Path & "\Recursos\MouseIcons\Baston" & ClientSetup.MouseBaston & ".ico")
    End If

    ' Mouse Icon to use in the rest of the game this one is animated
    ' We load it in frmMain but for some reason is loaded in the rest of the game
    ' Better for us :)
    Dim CursorAniDir As String
    Dim Cursor As Long
    
    'Si es 0 cargamos el por defecto
    If ClientSetup.MouseGeneral > 0 Then
        CursorAniDir = App.Path & "\Recursos\MouseIcons\MAIN.cur"
        hSwapCursor = SetClassLong(frmMain.hWnd, GLC_HCURSOR, LoadCursorFromFile(CursorAniDir))
        hSwapCursor = SetClassLong(frmMain.MainViewPic.hWnd, GLC_HCURSOR, LoadCursorFromFile(CursorAniDir))
        hSwapCursor = SetClassLong(frmMain.hlst.hWnd, GLC_HCURSOR, LoadCursorFromFile(CursorAniDir))
    End If
   
    frmCargando.Show
    frmCargando.Refresh
    
    '#######
    ' CLASES
    Call frmCargando.ActualizarCarga(JsonLanguage.item("INICIA_CLASES").item("TEXTO"), 10)
                            
    Set Dialogos = New clsDialogs
    Set Sound = New clsSoundEngine
    Set Inventario = New clsGraphicalInventory
    Set CustomKeys = New clsCustomKeys
    Set incomingData = New clsByteQueue
    Set outgoingData = New clsByteQueue
    Set MainTimer = New clsTimer
    Set clsForos = New clsForum
    Set frmMain.Client = New clsSocket

    'Esto es para el movimiento suave de pjs, para que el pj termine de hacer el movimiento antes de empezar otro
    Set keysMovementPressedQueue = New clsArrayList
    Call keysMovementPressedQueue.Initialize(1, 4)

    Call frmCargando.ActualizarCarga(frmCargando.Caption = JsonLanguage.item("HECHO").item("TEXTO"), 20)

    '#############
    ' DIRECT SOUND
    Call frmCargando.ActualizarCarga(JsonLanguage.item("INICIA_SONIDO").item("TEXTO"), 30)
    '
    'Inicializamos el sonido
    If Sound.Initialize_Engine(frmMain.hWnd, Path(ePath.recursos), Path(ePath.recursos), Path(ePath.recursos), False, (ClientSetup.bSound > 0), (ClientSetup.bMusic <> CONST_DESHABILITADA), ClientSetup.SoundVolume, ClientSetup.MusicVolume, ClientSetup.Invertido) Then
        'frmCargando.picLoad.Width = 300
    Else
        MsgBox "¡No se ha logrado iniciar el engine de DirectSound! Reinstale los últimos controladores de DirectX. No habrá soporte de audio en el juego.", vbCritical, "Advertencia"
        frmOpciones.Frame2.Enabled = False
    End If

    If ClientSetup.bMusic <> CONST_DESHABILITADA Then
        Sound.NextMusic = MUS_Carga
        Sound.Fading = 350
        Sound.Sound_Render
    End If
    
    DoEvents

    Call frmCargando.ActualizarCarga(JsonLanguage.item("HECHO").item("TEXTO"), 40)
    
    '###########
    ' CONSTANTES
    Call frmCargando.ActualizarCarga(JsonLanguage.item("INICIA_CONSTANTES").item("TEXTO"), 50)
    
    Call InicializarNombres
    
    ' Initialize FONTTYPES
    Call Protocol.InitFonts
 
    UserMap = 0
    
    Call frmCargando.ActualizarCarga(JsonLanguage.item("HECHO").item("TEXTO"), 60)

    '##############
    ' MOTOR GRAFICO
    Call frmCargando.ActualizarCarga(JsonLanguage.item("INICIA_MOTOR_GRAFICO").item("TEXTO"), 70)
    
    'Iniciamos el Engine de DirectX 8
    Call mDx8_Engine.Engine_DirectX8_Init
          
    'Tile Engine
    Call InitTileEngine(frmMain.hWnd, 32, 32, 8, 8)
    
    Call mDx8_Engine.Engine_DirectX8_Aditional_Init

    Call frmCargando.ActualizarCarga(JsonLanguage.item("HECHO").item("TEXTO"), 80)
    
    '###################
    ' ANIMACIONES EXTRAS
    Call frmCargando.ActualizarCarga(JsonLanguage.item("INICIA_FXS").item("TEXTO"), 90)
    
    Call CargarTips
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    Call CargarPasos
    
    Call frmCargando.ActualizarCarga(JsonLanguage.item("HECHO").item("TEXTO"), 95)
    
    'Inicializamos el inventario grafico
    Call Inventario.Initialize(DirectD3D8, frmMain.PicInv, MAX_INVENTORY_SLOTS, , , , , , , , True)
    
    'Set cKeys = New Collection
    Call frmCargando.ActualizarCarga(JsonLanguage.item("BIENVENIDO").item("TEXTO"), 100)

    Unload frmCargando
    
End Sub

Private Sub LoadTimerIntervals()
    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/03/2011
    'Set the intervals of timers
    '***************************************************
    
    With MainTimer
    
        Call .SetInterval(TimersIndex.Attack, eIntervalos.INT_ATTACK)
        Call .SetInterval(TimersIndex.Work, eIntervalos.INT_WORK)
        Call .SetInterval(TimersIndex.UseItemWithU, eIntervalos.INT_USEITEMU)
        Call .SetInterval(TimersIndex.UseItemWithDblClick, eIntervalos.INT_USEITEMDCK)
        Call .SetInterval(TimersIndex.SendRPU, eIntervalos.INT_SENTRPU)
        Call .SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
        Call .SetInterval(TimersIndex.Arrows, eIntervalos.INT_ARROWS)
        Call .SetInterval(TimersIndex.CastAttack, eIntervalos.INT_CAST_ATTACK)
        Call .SetInterval(TimersIndex.ChangeHeading, eIntervalos.INT_CHANGE_HEADING)
    
        'Init timers
        Call .Start(TimersIndex.Attack)
        Call .Start(TimersIndex.Work)
        Call .Start(TimersIndex.UseItemWithU)
        Call .Start(TimersIndex.UseItemWithDblClick)
        Call .Start(TimersIndex.SendRPU)
        Call .Start(TimersIndex.CastSpell)
        Call .Start(TimersIndex.Arrows)
        Call .Start(TimersIndex.CastAttack)
        Call .Start(TimersIndex.ChangeHeading)
    
    End With

End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, File
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Funcion para chequear el email
'
'  Corregida por Maraxus para que reconozca como validas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim Lx    As Long
    Dim iAsc  As Integer
    Dim Len_sString As Long
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . despues de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        'pre-calculo la cantidad de caracteres para mejorar el rendimiento
        Len_sString = Len(sString) - 1
        
        '3er test: Recorre todos los caracteres y los valida
        For Lx = 0 To Len_sString
            If Not (Lx = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (Lx + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next Lx
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como validas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer aca....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    '*******************************************
    'Author: Unknown
    'Last Modification: -
    '
    '*******************************************

    If X > XMinMapSize And X < XMaxMapSize + 1 And Y > YMinMapSize And Y < YMaxMapSize + 1 Then

        With MapData(X, Y)

            If ((.Graphic(1).GrhIndex >= 1505 And .Graphic(1).GrhIndex <= 1520) Or _
                (.Graphic(1).GrhIndex >= 12439 And .Graphic(1).GrhIndex <= 12454) Or _
                (.Graphic(1).GrhIndex >= 5665 And .Graphic(1).GrhIndex <= 5680) Or _
                (.Graphic(1).GrhIndex >= 13547 And .Graphic(1).GrhIndex <= 13562)) And _
                .Graphic(2).GrhIndex = 0 Then
                
                HayAgua = True
            
            Else
                HayAgua = False

            End If

        End With

    Else
        HayAgua = False

    End If

End Function

''
' Checks the command line parameters, if you are running Ao with /nores command
'
'

Public Sub LeerLineaComandos()
'*************************************************
'Author: Unknown
'Last modified: 25/11/2008 (BrianPr)
'
'*************************************************
    
    Dim i As Long, t() As String, Upper_t As Long, Lower_t As Long
    
    'Parseo los comandos
    t = Split(Command, " ")
    Lower_t = LBound(t)
    Upper_t = UBound(t)
    
    For i = Lower_t To Upper_t
        Select Case UCase$(t(i))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
        End Select
    Next i

End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, clases, skills, atributos, etc.
'**************************************************************
    
    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cLindos) = "Lindos"
    Ciudades(eCiudad.cRinkel) = "Rinkel"
    Ciudades(eCiudad.cArghal) = "Arghal"
    
    ListaRazas(eRaza.Humano) = JsonLanguage.item("RAZAS").item("HUMANO")
    ListaRazas(eRaza.Elfo) = JsonLanguage.item("RAZAS").item("ELFO")
    ListaRazas(eRaza.ElfoOscuro) = JsonLanguage.item("RAZAS").item("ELFO_OSCURO")
    ListaRazas(eRaza.Gnomo) = JsonLanguage.item("RAZAS").item("GNOMO")
    ListaRazas(eRaza.Enano) = JsonLanguage.item("RAZAS").item("ENANO")
    ListaRazas(eRaza.Orco) = JsonLanguage.item("RAZAS").item("ORCO")

    ListaClases(eClass.Mage) = JsonLanguage.item("CLASES").item("MAGO")
    ListaClases(eClass.Cleric) = JsonLanguage.item("CLASES").item("CLERIGO")
    ListaClases(eClass.Warrior) = JsonLanguage.item("CLASES").item("GUERRERO")
    ListaClases(eClass.Assasin) = JsonLanguage.item("CLASES").item("ASESINO")
    ListaClases(eClass.Thief) = JsonLanguage.item("CLASES").item("LADRON")
    ListaClases(eClass.Bard) = JsonLanguage.item("CLASES").item("BARDO")
    ListaClases(eClass.Druid) = JsonLanguage.item("CLASES").item("DRUIDA")
    ListaClases(eClass.Bandit) = JsonLanguage.item("CLASES").item("BANDIDO")
    ListaClases(eClass.Paladin) = JsonLanguage.item("CLASES").item("PALADIN")
    ListaClases(eClass.Hunter) = JsonLanguage.item("CLASES").item("CAZADOR")
    ListaClases(eClass.Nigromante) = JsonLanguage.item("CLASES").item("NIGROMANTE")
    ListaClases(eClass.Mercenario) = JsonLanguage.item("CLASES").item("MERCENARIO")
    ListaClases(eClass.Gladiador) = JsonLanguage.item("CLASES").item("GLADIADOR")
    ListaClases(eClass.Pescador) = JsonLanguage.item("CLASES").item("PESCADOR")
    ListaClases(eClass.Herrero) = JsonLanguage.item("CLASES").item("HERRERO")
    ListaClases(eClass.Lenador) = JsonLanguage.item("CLASES").item("LENADOR")
    ListaClases(eClass.Minero) = JsonLanguage.item("CLASES").item("MINERO")
    ListaClases(eClass.Carpintero) = JsonLanguage.item("CLASES").item("CARPINTERO")
    ListaClases(eClass.Sastre) = JsonLanguage.item("CLASES").item("SASTRE")
   
    SkillsNames(eSkill.Magia) = JsonLanguage.item("HABILIDADES").item("MAGIA").item("TEXTO")
    SkillsNames(eSkill.Robar) = JsonLanguage.item("HABILIDADES").item("ROBAR").item("TEXTO")
    SkillsNames(eSkill.Tacticas) = JsonLanguage.item("HABILIDADES").item("EVASION_EN_COMBATE").item("TEXTO")
    SkillsNames(eSkill.Armas) = JsonLanguage.item("HABILIDADES").item("COMBATE_CON_ARMAS").item("TEXTO")
    SkillsNames(eSkill.Meditar) = JsonLanguage.item("HABILIDADES").item("MEDITAR").item("TEXTO")
    SkillsNames(eSkill.Apunalar) = JsonLanguage.item("HABILIDADES").item("APUNALAR").item("TEXTO")
    SkillsNames(eSkill.Ocultarse) = JsonLanguage.item("HABILIDADES").item("OCULTARSE").item("TEXTO")
    SkillsNames(eSkill.Supervivencia) = JsonLanguage.item("HABILIDADES").item("SUPERVIVENCIA").item("TEXTO")
    SkillsNames(eSkill.Talar) = JsonLanguage.item("HABILIDADES").item("TALAR").item("TEXTO")
    SkillsNames(eSkill.Comerciar) = JsonLanguage.item("HABILIDADES").item("COMERCIO").item("TEXTO")
    SkillsNames(eSkill.Defensa) = JsonLanguage.item("HABILIDADES").item("DEFENSA_CON_ESCUDOS").item("TEXTO")
    SkillsNames(eSkill.pesca) = JsonLanguage.item("HABILIDADES").item("PESCA").item("TEXTO")
    SkillsNames(eSkill.Mineria) = JsonLanguage.item("HABILIDADES").item("MINERIA").item("TEXTO")
    SkillsNames(eSkill.Carpinteria) = JsonLanguage.item("HABILIDADES").item("CARPINTERIA").item("TEXTO")
    SkillsNames(eSkill.Herreria) = JsonLanguage.item("HABILIDADES").item("HERRERIA").item("TEXTO")
    SkillsNames(eSkill.Liderazgo) = JsonLanguage.item("HABILIDADES").item("LIDERAZGO").item("TEXTO")
    SkillsNames(eSkill.Domar) = JsonLanguage.item("HABILIDADES").item("DOMAR_ANIMALES").item("TEXTO")
    SkillsNames(eSkill.Proyectiles) = JsonLanguage.item("HABILIDADES").item("COMBATE_A_DISTANCIA").item("TEXTO")
    SkillsNames(eSkill.Marciales) = JsonLanguage.item("HABILIDADES").item("COMBATE_CUERPO_A_CUERPO").item("TEXTO")
    SkillsNames(eSkill.Navegacion) = JsonLanguage.item("HABILIDADES").item("NAVEGACION").item("TEXTO")
    SkillsNames(eSkill.Botanica) = JsonLanguage.item("HABILIDADES").item("BOTANICA").item("TEXTO")
    SkillsNames(eSkill.Sastreria) = JsonLanguage.item("HABILIDADES").item("SASTRERIA").item("TEXTO")
    SkillsNames(eSkill.Arrojadizas) = JsonLanguage.item("HABILIDADES").item("ARROJADIZAS").item("TEXTO")
    SkillsNames(eSkill.Resistencia) = JsonLanguage.item("HABILIDADES").item("RESISTENCIA").item("TEXTO")
    SkillsNames(eSkill.Musica) = JsonLanguage.item("HABILIDADES").item("MUSICA").item("TEXTO")

    AtributosNames(eAtributos.Fuerza) = JsonLanguage.item("ATRIBUTOS").item("FUERZA")
    AtributosNames(eAtributos.Agilidad) = JsonLanguage.item("ATRIBUTOS").item("AGILIDAD")
    AtributosNames(eAtributos.Inteligencia) = JsonLanguage.item("ATRIBUTOS").item("INTELIGENCIA")
    AtributosNames(eAtributos.Carisma) = JsonLanguage.item("ATRIBUTOS").item("CARISMA")
    AtributosNames(eAtributos.Constitucion) = JsonLanguage.item("ATRIBUTOS").item("CONSTITUCION")
End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
'**************************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    frmMain.RecTxt.Text = vbNullString
    
    Call Dialogos.RemoveAllDialogs
End Sub

Public Sub CloseClient()
    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 8/14/2007
    'Frees all used resources, cleans up and leaves
    '**************************************************************
    
    ' Allow new instances of the client to be opened
    Call Application.ReleaseInstance
    
    EngineRun = False
    
    'WyroX:
    'Guardamos antes de cerrar porque algunas configuraciones
    'no se guardan desde el menu opciones (Por ej: M=Musica)
    'Fix: intentaba guardar cuando el juego cerraba por un error,
    'antes de cargar los recursos. Me aprovecho de prgRun
    'para saber si ya fueron cargados
    If prgRun Then
        Call Carga.GuardarConfiguracion
    End If

    'Cerramos Sockets/Winsocks/WindowsAPI
    frmMain.Client.CloseSck
    
    'Stop tile engine
    Call Engine_DirectX8_End

    'Destruimos los objetos publicos creados
    Call Sound.Engine_DeInitialize
    Set Sound = Nothing
    
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    Set JsonLanguage = Nothing
    Set frmMain.Client = Nothing
    
    Call UnloadAllForms
    
    'Si se cambio la resolucion, la reseteamos.
    If ResolucionCambiada Then Resolution.ResetResolution
    
    End
    
End Sub

Public Function EsGM(ByVal CharIndex As Integer) As Boolean

    If charlist(CharIndex).priv >= 1 And charlist(CharIndex).priv <= 5 Or charlist(CharIndex).priv = 25 Then
        EsGM = True
    End If
    
    EsGM = False

End Function

Public Function EsNPC(ByVal CharIndex As Integer) As Boolean

    If charlist(CharIndex).iHead = 0 Then
        EsNPC = True
    End If
    
    EsNPC = False

End Function

Public Function getTagPosition(ByVal Nick As String) As Integer
    
    Dim buf As Integer
        buf = InStr(Nick, "<")

    If buf > 0 Then
        getTagPosition = buf
        Exit Function
    End If
    
    buf = InStr(Nick, "[")

    If buf > 0 Then
        getTagPosition = buf
        Exit Function
    End If
    
    getTagPosition = Len(Nick) + 2
    
End Function

Public Function getCharIndexByName(ByVal name As String) As Integer
    
    Dim i As Long

    For i = 1 To LastChar

        If charlist(i).Nombre = name Then
            getCharIndexByName = i
            Exit Function
        End If
    Next i

End Function

Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Returns true if the post is sticky.
'***************************************************
    Select Case ForumType
        Case eForumMsgType.ieCAOS_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieGENERAL_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieREAL_STICKY
            EsAnuncio = True
            
    End Select
    
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte
'***************************************************
'Author: ZaMa
'Last Modification: 01/03/2010
'Returns the forum alignment.
'***************************************************
    Select Case yForumType
        Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
            ForumAlignment = eForumType.ieCAOS
            
        Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
            ForumAlignment = eForumType.ieGeneral
            
        Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
            ForumAlignment = eForumType.ieREAL
            
    End Select
    
End Function

Public Sub ResetAllInfo(Optional ByVal UnloadForms As Boolean = True)

    ' Disable timers
    frmMain.Second.Enabled = False
    Connected = False
    Call frmMain.hlst.Clear ' Ponemos esto aca para limpiar la lista de hechizos al desconectarse.
    
    If UnloadForms Then
        'Unload all forms except frmMain, frmConnect
        Dim frm As Form
        For Each frm In Forms
            If frm.name <> frmMain.name And _
               frm.name <> frmConnect.name Then
                
                Call Unload(frm)
            End If
        Next
    End If
    
    On Local Error GoTo 0
    
    If UnloadForms Then
        If Not frmCrearPersonaje.Visible Then frmCharList.Visible = True
        ' Return to connection screen
        frmMain.Visible = False
    End If
    
    'Stop audio
    Sound.Sound_Stop_All
    Sound.Ambient_Stop
    
    ' Reset flags
    pausa = False
    UserMeditar = False
    UserEstupido = False
    UserCiego = False
    UserDescansar = False
    UserParalizado = False
    UserNavegando = False
    UserEvento = False
    bFogata = False
    bFogata = False
    Comerciando = False
    bShowTutorial = False
    
    MirandoAsignarSkills = False
    MirandoEstadisticas = False
    MirandoForo = False
    MirandoTrabajo = 0
    MirandoParty = False
    UserMap = 0
    
    'Delete all kind of dialogs
    Call CleanDialogs

    'Reset some char variables...
    Dim i As Long
    For i = 1 To LastChar
        charlist(i).invisible = False

    Next i

    ' Reset stats
    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserELO = 0
    Alocados = 0
    UserEquitando = 0
    Alocados = 0
    SkillPoints = 0
    
    With UserPet
        .Nombre = vbNullString
        .tipo = 0
        .Habilidad = vbNullString
        .ELU = 0
        .EXP = 0
        .MaxHIT = 0
        .MinHIT = 0
        .MaxHP = 0
        .MinHP = 0
    End With
    
    Call Actualizar_Estado(e_estados.MedioDia)

    Call SetSpeedUsuario(SPEED_NORMAL)

    ' Reset skills
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    ' Reset attributes
    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    ' Clear inventory slots
    Inventario.ClearAllSlots
    
    For i = 1 To NUMMACROS
        MacrosKey(i).TipoAccion = 0
        MacrosKey(i).invName = vbNullString
        MacrosKey(i).InvGrh = 0
        MacrosKey(i).invName = vbNullString
        MacrosKey(i).Comando = vbNullString
        
        frmMain.picMacro(i - 1).Picture = Nothing
    Next i

    Call DibujarMenuMacros
    
    If ClientSetup.bMusic <> CONST_DESHABILITADA Then
        If ClientSetup.bMusic <> CONST_DESHABILITADA Then
            Sound.NextMusic = MUS_VolverInicio
            Sound.Fading = 200
        End If
    End If

End Sub

Public Sub ResetAllInfoAccounts()
'**************************************
'Autor: Lorwik
'Fecha: 21/05/2020
'Descripcion: Borra los datos almacenados de una cuenta
'**************************************

    If NumberOfCharacters > 0 Then
    
        Dim LoopC As Long
        
        For LoopC = 1 To NumberOfCharacters
        
            With cPJ(LoopC)
                .Nombre = ""
                .Body = 0
                .Head = 0
                .weapon = 0
                .shield = 0
                .helmet = 0
                .Class = 0
                .Race = 0
                .Map = 0
                .Level = 0
                .Criminal = False
                .Dead = False
                
                .GameMaster = False
            End With
            
            frmCharList.picChar(LoopC).Cls
            frmCharList.picChar(LoopC).Refresh
            frmCharList.lblAccData(LoopC).Caption = ""
            frmCharList.lblCharData(0).Caption = ""
            frmCharList.lblCharData(1).Caption = ""
            frmCharList.lblCharData(2).Caption = ""
            
        Next LoopC
        
        frmCharList.lblAccData(0).Caption = ""
        
    End If
End Sub

' USO: If ArrayInitialized(Not ArrayName) Then ...
Public Function ArrayInitialized(ByVal TheArray As Long) As Boolean
'***************************************************
'Author: Jopi
'Last Modify Date: 03/01/2020
'Chequea que se haya inicializado el Array.
'***************************************************
    
    ArrayInitialized = Not (TheArray = -1&)

End Function

Public Sub SetSpeedUsuario(ByVal speed As Double)
    Engine_BaseSpeed = speed
End Sub

Public Function CheckIfIpIsNumeric(CurrentIp As String) As String
    If IsNumeric(mid$(CurrentIp, 1, 1)) Then
        CheckIfIpIsNumeric = True
    Else
        CheckIfIpIsNumeric = False
    End If
End Function

Public Sub Client_Screenshot(ByVal hDC As Long, ByVal Width As Long, ByVal Height As Long)
'*******************************
'Autor: ???
'Fecha: ???
'*******************************

On Error GoTo ErrorHandler

    Dim i As Long
    Dim Index As Long
    i = 1
    
    Set m_Jpeg = New clsJpeg
    
    '80 Quality
    m_Jpeg.Quality = 100
    
    'Sample the cImage by hDC
    m_Jpeg.SampleHDC hDC, Width, Height
    
    m_FileName = App.Path & "\Fotos\ImperiumClasico_Foto"
    
    If Dir(App.Path & "\Fotos", vbDirectory) = vbNullString Then
        MkDir (App.Path & "\Fotos")
    End If
    
    Do While Dir(m_FileName & Trim(str(i)) & ".jpg") <> vbNullString
        i = i + 1
        DoEvents
    Loop
    
    Index = i
    
    m_Jpeg.Comment = "Character: " & UserName & " - " & Format(Date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm AM/PM")
    
    'Save the JPG file
    m_Jpeg.SaveFile m_FileName & Trim(str(Index)) & ".jpg"
    
    Call AddtoRichTextBox(frmMain.RecTxt, "¡Captura realizada con exito! Se guardo en " & m_FileName & Trim(str(Index)) & ".jpg", 204, 193, 155, 0, 1)
    
    Set m_Jpeg = Nothing
    
    Exit Sub

ErrorHandler:
    Call AddtoRichTextBox(frmMain.RecTxt, "¡Error en la captura!", 204, 193, 155, 0, 1)

End Sub

Public Sub MostrarMensaje(ByVal Mensaje As String)
'****************************************
'Autor: Lorwik
'Fecha: 20/05/2020
'Descripción: Llama al frmMensaje para mostrar un cartel de mensaje
'****************************************
    'Call Sound.Sound_Play(SND_MSG)
    
    frmMensaje.msg.Caption = Mensaje
    frmMensaje.Show

End Sub

Public Function General_Distance_Get(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer) As Integer
'*************************************
'Author: Lorwik
'Last Modify Date: Unknown
'*************************************

    General_Distance_Get = Abs(x1 - x2) + Abs(y1 - y2)

End Function

Public Sub Form_RemoveTitleBar(F As Form)
    Dim Style As Long
    ' Get window's current style bits.
    Style = GetWindowLong(F.hWnd, GWL_STYLE)
    ' Set the style bit for the title off.
    Style = Style And Not WS_CAPTION

    ' Send the new style to the window.
    SetWindowLong F.hWnd, GWL_STYLE, Style
    ' Repaint the window.
    'SetWindowPos f.hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
End Sub

Public Sub ListarServidores()

    On Error Resume Next
    
    Dim lista() As String
    Dim Elementos As Byte
    
    Dim i As Byte
    Dim responseServer As String
    
    Set Inet = New clsInet
    
    responseServer = Inet.OpenRequest("https://tuurl.com/server-listiac.txt", "GET")
    responseServer = Inet.Execute
    responseServer = Inet.GetResponseAsString
    
    lista = Split(responseServer, ";")
    
    'Limpiamos la info anterior
    frmConnect.lst_servers.Clear
    
    ReDim Servidor(0 To UBound(lista())) As Servidores
    
    For i = 0 To UBound(lista())
        Servidor(i).Ip = ReadField(1, lista(i), Asc("|"))
        Servidor(i).Puerto = ReadField(2, lista(i), Asc("|"))
        Servidor(i).Nombre = ReadField(3, lista(i), Asc("|"))
        
        frmConnect.lst_servers.AddItem Servidor(i).Nombre, i
    Next i
    
    If ServIndSel < 0 Or ServIndSel > frmConnect.lst_servers.ListCount Then _
        ServIndSel = 0
    
    frmConnect.lst_servers.ListIndex = ServIndSel

End Sub

Public Sub DibujarMinimapa()

    Dim map_x, map_y, Capas As Byte
    
    'Primero limpiamos el minimapa anterior
    frmMain.MiniMapa.Cls
    
    For map_y = YMinMapSize To YMaxMapSize
        For map_x = XMinMapSize To XMaxMapSize
        
            For Capas = 1 To 2
                If MapData(map_x, map_y).Graphic(Capas).GrhIndex > 0 Then
                    SetPixel frmMain.MiniMapa.hDC, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(Capas).GrhIndex).mini_map_color
                End If
                
            Next Capas
        Next map_x
    Next map_y
   
   'Refrescamos
    frmMain.MiniMapa.Refresh
End Sub

Public Sub accionMacrosKey(ByVal Index As Byte, Optional ByVal BotonDerecho As Boolean = False)
'*********************************************
'Autor: Lorwik
'Fecha: 07/03/2021
'Descripcion: Si existe el macro manda al server la orden, si no existe se abre la config
'*********************************************

    MacroElegido = Index

    With MacrosKey(Index)
    
        If .TipoAccion = 0 Or BotonDerecho Then
            frmBindKey.Show vbModeless, frmMain
       Else
            Select Case .TipoAccion
                Case 1 '¿Es un comando?
                    Call ParseUserCommand("/" & .Comando)
                
                Case 2 'Hechizos
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                        End With
                    End If
                    
                    If MainTimer.Check(TimersIndex.Work) Then _
                        Call WriteEjecutarMacro(Index)
                        
                Case 3 'Equipar
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                        End With
                    End If
            
                    If MainTimer.Check(TimersIndex.UseItemWithDblClick) Then _
                        Call WriteEjecutarMacro(Index)
                        
                
                Case 4 'Usar
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then _
                        Call WriteEjecutarMacro(Index)
                
            End Select
        End If
    End With
    
End Sub

