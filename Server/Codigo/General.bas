Attribute VB_Name = "General"
'Argentum Online 0.12.2
'Copyright (C) 2002 Marquez Pablo Ignacio
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

#If False Then

    Dim X, Y, Map, K, errHandler, obj, index, n, Email As Variant

#End If

Global LeerNPCs As clsIniManager

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, _
                     Optional ByVal Mimetizado As Boolean = False)
    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/14/07
    'Da cuerpo desnudo a un usuario
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************

    Dim MiCuerpoDesnudo As Integer

    With UserList(UserIndex)

        MiCuerpoDesnudo = CuerpoDesnudo(.Genero, .Raza)
    
        If Mimetizado Then
            .CharMimetizado.body = MiCuerpoDesnudo
        Else
            .Char.body = MiCuerpoDesnudo

        End If
    
        .flags.Desnudo = 1

    End With

End Sub

Public Function CuerpoDesnudo(ByVal Genero As eGenero, ByVal Raza As eRaza)
'************************************************
'Autor: Lorwik
'Fecha: 01/05/2020
'Descripcion: Devuelve el cuerpo desnudo correspondiente a la raza y sexo
'************************************************

    Select Case Genero

            Case eGenero.Hombre

                Select Case Raza

                    Case eRaza.Humano
                        CuerpoDesnudo = 21

                    Case eRaza.Drow
                        CuerpoDesnudo = 32

                    Case eRaza.Elfo
                        CuerpoDesnudo = 21

                    Case eRaza.Gnomo
                        CuerpoDesnudo = 53

                    Case eRaza.Enano
                        CuerpoDesnudo = 53
                        
                    Case eRaza.Orco
                        CuerpoDesnudo = 248

                End Select

            Case eGenero.Mujer

                Select Case Raza

                    Case eRaza.Humano
                        CuerpoDesnudo = 39

                    Case eRaza.Drow
                        CuerpoDesnudo = 40

                    Case eRaza.Elfo
                        CuerpoDesnudo = 39

                    Case eRaza.Gnomo
                        CuerpoDesnudo = 60

                    Case eRaza.Enano
                        CuerpoDesnudo = 60
                        
                    Case eRaza.Orco
                        CuerpoDesnudo = 249

                End Select

        End Select
End Function

Sub Bloquear(ByVal toMap As Boolean, _
             ByVal sndIndex As Integer, _
             ByVal X As Integer, _
             ByVal Y As Integer, _
             ByVal b As Boolean)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'b ahora es boolean,
    'b=true bloquea el tile en (x,y)
    'b=false desbloquea el tile en (x,y)
    'toMap = true -> Envia los datos a todo el mapa
    'toMap = false -> Envia los datos al user
    'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
    'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s
    '***************************************************

    If toMap Then
        Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, b))
    Else
        Call WriteBlockPosition(sndIndex, X, Y, b)

    End If

End Sub

Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    '*******************************************
    'Author: Unknown
    'Last Modification: -
    '
    '*******************************************

    If Map > 0 And Map < NumMaps + 1 And X > XMinMapSize And X < XMaxMapSize + 1 And Y > YMinMapSize And Y < YMaxMapSize + 1 Then

        With MapData(Map, X, Y)

            If ((.Graphic(1) >= 1505 And .Graphic(1) <= 1520) Or _
                (.Graphic(1) >= 12439 And .Graphic(1) <= 12454) Or _
                (.Graphic(1) >= 5665 And .Graphic(1) <= 5680) Or _
                (.Graphic(1) >= 13547 And .Graphic(1) <= 13562)) And _
                .Graphic(2) = 0 Then
                
                HayAgua = True
            
            Else
                HayAgua = False

            End If

        End With

    Else
        HayAgua = False

    End If

End Function

Private Function HayLava(ByVal Map As Integer, _
                         ByVal X As Integer, _
                         ByVal Y As Integer) As Boolean

    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/12/07
    '***************************************************
    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        If MapData(Map, X, Y).Graphic(1) >= 5837 And MapData(Map, X, Y).Graphic(1) <= 5852 Then
            HayLava = True
        Else
            HayLava = False

        End If

    Else
        HayLava = False

    End If

End Function

Function HaySacerdote(ByVal UserIndex As Integer) As Boolean
    '******************************
    'Adaptacion a 13.0: Kaneidra
    'Last Modification: 15/05/2012
    '******************************
 
    Dim X As Integer, Y As Integer
    
    With UserList(UserIndex)
    
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
       
                If MapData(.Pos.Map, X, Y).NPCIndex > 0 Then
                    If Npclist(MapData(.Pos.Map, X, Y).NPCIndex).NPCtype = eNPCType.Revividor Then
                       
                        If Distancia(.Pos, Npclist(MapData(.Pos.Map, X, Y).NPCIndex).Pos) < 5 Then
                            HaySacerdote = True
                            Exit Function
                        End If

                    End If

                End If
           
            Next X
        Next Y
    
    End With
 
    HaySacerdote = False
 
End Function

Sub EnviarSpawnList(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim K          As Long
    Dim npcNames() As String
    
    ReDim npcNames(1 To UBound(SpawnList)) As String
    
    For K = 1 To UBound(SpawnList)
        npcNames(K) = SpawnList(K).NpcName
    Next K
    
    Call WriteSpawnList(UserIndex, npcNames())

End Sub

Public Function GetVersionOfTheServer() As String
    GetVersionOfTheServer = GetVar(App.Path & "\Server.ini", "INIT", "VersionTagRelease")
End Function

Sub Main()
    '***************************************************
    'Author: Unknown
    'Last Modification: 15/03/2011
    '15/03/2011: ZaMa - Modularice todo, para que quede mas claro.
    '***************************************************

    On Error Resume Next
    
    ChDir App.Path
    ChDrive App.Path
    
    'Inicializamos la cabecera
    Call IniciarCabecera
    
    Call BanIpCargar
    
    Call BanGlobalChatCargar
    GlobalChatActive = True
    
    ' Start loading..
    frmCargando.Show
    
    ' Constants & vars
    frmCargando.Label1(2).Caption = "Cargando constantes..."
    Call LoadConstants
    Call InicializarSonidos
    DoEvents
    
    ' Motd
    frmCargando.Label1(2).Caption = "Cargando Motd..."
    Call LoadMotd
    DoEvents
    
    ' Arrays
    frmCargando.Label1(2).Caption = "Iniciando Arrays..."
    Call LoadArrays
    
    ' Server.ini & Apuestas.dat & Ciudades.dat
    frmCargando.Label1(2).Caption = "Cargando Server.ini"
    Call LoadSini 'Configuraci�n general (Server.ini)
    Call Load_Rates 'Rates (Rates.ini)
    Call loadAdministrativeUsers 'Gms (GameMasters.ini)
    Call CargarCiudades
    Call CargaApuestas
    
    'Base de datos MySQL
#If DBConexionUnica = 1 Then
    frmCargando.Label1(2).Caption = "Cargando Base de datos"
    
    Set User_Database = New clsDataBase
    Set Account_Database = New clsDataBase
    
    Call Load_ConfigDatBase
    Call User_Database.Database_Connect
    Call Account_Database.Database_Connect
#End If

    ' Npcs.dat
    frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
    Call CargaNpcsDat

    ' Obj.dat
    frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
    Call LoadOBJData
    Call LoadGlobalDrop
    
    ' Hechizos.dat
    frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
    Call CargarHechizos
    
    frmCargando.Label1(2).Caption = "Cargando Objetos de Profesiones"
    Call LoadArmasHerreria
    Call LoadArmadurasHerreria
    Call LoadObjCarpintero
    Call LoadObjAlquimia
    Call LoadObjSastre
    
    ' Balance.dat
    frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
    Call LoadBalance
    
    ' Armaduras faccionarias
    frmCargando.Label1(2).Caption = "Cargando ArmadurasFaccionarias.dat"
    Call LoadArmadurasFaccion

    
    ' Mapas
    If BootDelBackUp Then
        frmCargando.Label1(2).Caption = "Cargando Backup"
        Call CargarBackUp
    Else
        frmCargando.Label1(2).Caption = "Cargando Mapas"
        Call LoadMapData

    End If
    
    Call InitializeAreas
    
    'Arenas de Retos
    Call LoadArenas
    
    ' Connections
    Call ResetUsersConnections
    
    ' Sockets
    Call SocketConfig
    
    frmCargando.Label1(2).Caption = "Cargando Clima"
    Call SortearHorario 'Lorwik> Lo coloco aqui o no funciona
    
    ' Timers
    Call InitMainTimers
    
    ' End loading..
    Unload frmCargando
    
    'Log start time
    LogServerStartTime
    
    'Ocultar
    If HideMe Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If
    
    tInicioServer = GetTickCount() And &H7FFFFFFF

    NombreServidor = GetVar(ConfigPath & "Server.ini", "INIT", "Nombre")

    frmMain.Caption = NombreServidor & " Server v. " & ULTIMAVERSION & " - Iniciado en el puerto: " & Puerto

    'Este ultimo es para saber siempre los records en el frmMain
    frmMain.txtRecordOnline.Text = RecordUsuariosOnline

    'En caso que la API este activada, la abrimos :)
    'el repositorio para hacer funcionar esto, es este: https://github.com/ao-libre/ao-api-server
    'Si no tienen interes en usarlo pueden desactivarlo en el Server.ini
    If ConexionAPI Then
        ApiNodeJsTaskId = Shell("cmd /c cd " & ApiPath & " && npm start")
    End If

End Sub

Private Sub LoadConstants()

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Loads all constants and general parameters.
    '*****************************************************************
    Dim i As Integer
    
    On Error Resume Next
   
    LastBackup = Format(Now, "Short Time")
    Minutos = Format(Now, "Short Time")
    
    ' Paths
    DatPath = App.Path & "\Dat\"
    ConfigPath = App.Path & "\Configuracion\"
    
    'Lorwik: Nueva subida de Skills, subira de 2 en 2 hasta el lvl max.
    LevelSkill(1).LevelValue = 2
    For i = 2 To 50
        LevelSkill(i).LevelValue = LevelSkill(i - 1).LevelValue + 2
    Next i
    
    ' Races
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.Drow) = "Drow"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    ListaRazas(eRaza.Orco) = "Orco"
    
    ' Classes
    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Nigromante) = "Nigromante"
    ListaClases(eClass.Mercenario) = "Mercenario"
    ListaClases(eClass.Gladiador) = "Gladiador"
    ListaClases(eClass.Pescador) = "Pescador"
    ListaClases(eClass.Herrero) = "Herrero"
    ListaClases(eClass.Lenador) = "Le�ador"
    ListaClases(eClass.Minero) = "Minero"
    ListaClases(eClass.Carpintero) = "Carpintero"
    ListaClases(eClass.Sastre) = "Sastre"
    
    ' Skills
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasion en combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apunalar) = "Apunalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.talar) = "Talar"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Marciales) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"
    SkillsNames(eSkill.Equitacion) = "Equitacion"
    SkillsNames(eSkill.Botanica) = "Botanica"
    SkillsNames(eSkill.Alquimia) = "Alquimia"
    SkillsNames(eSkill.Arrojadizas) = "Armas Arrojadizas"
    SkillsNames(eSkill.Resistencia) = "Resistencia Magica"
    SkillsNames(eSkill.Musica) = "Musica"
    
    ' Attributes
    ListaAtributos(eAtributos.Fuerza) = "Fuerza"
    ListaAtributos(eAtributos.Agilidad) = "Agilidad"
    ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
    ListaAtributos(eAtributos.Carisma) = "Carisma"
    ListaAtributos(eAtributos.Constitucion) = "Constitucion"
    
    ' Fishes
    ListaPeces(1) = PECES_POSIBLES.PESCADO1
    ListaPeces(2) = PECES_POSIBLES.PESCADO2
    ListaPeces(3) = PECES_POSIBLES.PESCADO3
    ListaPeces(4) = PECES_POSIBLES.PESCADO4

    'Bordes del mapa
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)
    
    Set Ayuda = New cCola
    Set Denuncias = New cCola
    Denuncias.MaxLenght = MAX_DENOUNCES

    MaxUsers = 0

    ' Initialize classes
    Set WSAPISock2Usr = New Collection
    Protocol.InitAuxiliarBuffer

    Set aClon = New clsAntiMassClon
    Set TrashCollector = New Collection

End Sub

Private Sub LoadArrays()

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Loads all arrays
    '*****************************************************************
    On Error Resume Next

    ' Load Records
    Call LoadRecords
    
    ' Load guilds info
    Call LoadGuildsDB
    
    ' Load forbidden words
    Call CargarForbidenWords

End Sub

Private Sub ResetUsersConnections()

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Resets Users Connections.
    '*****************************************************************
    On Error Resume Next

    Dim LoopC As Long

    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC
    
End Sub

Private Sub InitMainTimers()

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Initializes Main Timers.
    '*****************************************************************
    On Error Resume Next

    With frmMain
        .AutoSave.Enabled = True

        .GameTimer.Enabled = True
        .PacketResend.Enabled = True
        .TIMER_AI.Enabled = True
        .Auditoria.Enabled = True
    End With
    
End Sub

Private Sub SocketConfig()

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Sets socket config.
    '*****************************************************************
    On Error Resume Next

    Call SecurityIp.InitIpTables(1000)
    
    If LastSockListen >= 0 Then
        Call apiclosesocket(LastSockListen) 'Cierra el socket de escucha
    End If

    Call IniciaWsApi(frmMain.hWnd)
    
    SockListen = ListenForConnect(Puerto, hWndMsg, vbNullString)

    If SockListen <> -1 Then
        ' Guarda el socket escuchando
        Call WriteVar(ConfigPath & "Server.ini", "INIT", "LastSockListen", SockListen)
    Else
        Call MsgBox("Ha ocurrido un error al iniciar el socket del Servidor.", vbCritical + vbOKOnly)
    End If
    
    frmMain.txtStatus.Text = Date & " " & time & " - Escuchando conexiones entrantes ..."
    
End Sub

Function FileExist(ByVal File As String, _
                   Optional FileType As VbFileAttribute = vbNormal) As Boolean
    '*****************************************************************
    'Se fija si existe el archivo
    '*****************************************************************

    FileExist = LenB(Dir$(File, FileType)) <> 0

End Function

Function ReadField(ByVal Pos As Integer, _
                   ByRef Text As String, _
                   ByVal SepASCII As Byte) As String
    '*****************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/15/2004
    'Gets a field from a delimited string
    '*****************************************************************

    Dim i          As Long
    Dim lastPos    As Long
    Dim CurrentPos As Long
    Dim delimiter  As String * 1
    
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

Function MapaValido(ByVal Map As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    MapaValido = Map >= 1 And Map <= NumMaps

End Function

Sub MostrarNumUsers()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    frmMain.txtNumUsers.Text = NumUsers

End Sub

Sub MostrarNumCuentas()
    '***************************************************
    'Author: Lorwik
    'Last Modification: 21/05/2020
    '
    '***************************************************

    frmMain.txtNumCuentas.Text = NumCuentas

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Arg As String

    Dim i   As Integer
    
    For i = 1 To 33
    
        Arg = ReadField(i, cad, 44)
    
        If LenB(Arg) = 0 Then Exit Function
    
    Next i
    
    ValidInputNP = True

End Function

Sub Restart()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'Se asegura de que los sockets estan cerrados e ignora cualquier err
    On Error Resume Next

    If frmMain.Visible Then frmMain.txtStatus.Text = "Reiniciando."
    
    Dim LoopC As Long


    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

    For LoopC = 1 To MaxUsers
        Call CloseSocket(LoopC)
    Next
    
    'Initialize statistics!!
    Call Statistics.Initialize
    
    For LoopC = 1 To UBound(UserList())
        Set UserList(LoopC).incomingData = Nothing
        Set UserList(LoopC).outgoingData = Nothing
    Next LoopC
    
    ReDim UserList(1 To MaxUsers) As User
    
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC
    
    LastUser = 0
    NumUsers = 0
    NumCuentas = 0
    
    Call FreeNPCs
    Call FreeCharIndexes
    
    Call LoadSini
    
    Call ResetForums
    Call LoadOBJData
    Call LoadGlobalDrop
    
    Call LoadMapData
    
    Call CargarHechizos

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " servidor reiniciado correctamente. - Escuchando conexiones entrantes ..."
    
    'Log it
    Dim n As Integer

    n = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #n
    Print #n, Date & " " & time & " servidor reiniciado."
    Close #n
    
    'Ocultar
    
    If HideMe Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)

    End If
  
End Sub

Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 15/11/2009
    '15/11/2009: ZaMa - La lluvia no quita stamina en las arenas.
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '**************************************************************

    With UserList(UserIndex)

        If MapInfo(.Pos.Map).Zona <> "DUNGEON" Then
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger <> eTrigger.BAJOTECHO And _
             MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger <> eTrigger.CASA And _
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger <> eTrigger.ZONASEGURA Then _
                Intemperie = True
        Else
            Intemperie = False

        End If

    End With
    
    'En las arenas no te afecta la lluvia
    If IsArena(UserIndex) Then Intemperie = False

End Function

Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer

    For i = 1 To MAXMASCOTAS

        With UserList(UserIndex)

            If .MascotasIndex(i) > 0 Then
                If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                    Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia - 1

                    If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(.MascotasIndex(i), 0)

                End If

            End If

        End With

    Next i

End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Unkonwn
    'Last Modification: 23/11/2009
    'If user is naked and it's in a cold map, take health points from him
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    Dim modifi As Integer
    
    With UserList(UserIndex)

        If .Counters.Frio < IntervaloFrio Then
            .Counters.Frio = .Counters.Frio + 1
        Else '

            If TerrainStringToByte(MapInfo(.Pos.Map).Terreno) = eTerrain.terrain_nieve Then
                Call WriteConsoleMsg(UserIndex, "Estas muriendo de frio, abrigate o moriras!!", FontTypeNames.FONTTYPE_INFO)
                modifi = Porcentaje(.Stats.MaxHp, 5)
                .Stats.MinHp = .Stats.MinHp - modifi
                
                If .Stats.MinHp < 1 Then
                    Call WriteConsoleMsg(UserIndex, "Has muerto de frio!!", FontTypeNames.FONTTYPE_INFO)
                    .Stats.MinHp = 0
                    Call UserDie(UserIndex)

                End If
                
                Call WriteUpdateHP(UserIndex)
            Else
                modifi = Porcentaje(.Stats.MaxSta, 5)
                Call QuitarSta(UserIndex, modifi)
                Call WriteUpdateSta(UserIndex)

            End If
            
            .Counters.Frio = 0

        End If

    End With

End Sub

Public Sub EfectoLava(ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 23/11/2009
    'If user is standing on lava, take health points from him
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    With UserList(UserIndex)

        If .Counters.Lava < IntervaloFrio Then 'Usamos el mismo intervalo que el del frio
            .Counters.Lava = .Counters.Lava + 1
        Else

            If HayLava(.Pos.Map, .Pos.X, .Pos.Y) Then
                Call WriteConsoleMsg(UserIndex, "Quitate de la lava, te estas quemando!!", FontTypeNames.FONTTYPE_INFO)
                .Stats.MinHp = .Stats.MinHp - Porcentaje(.Stats.MaxHp, 5)
                    
                If .Stats.MinHp < 1 Then
                    Call WriteConsoleMsg(UserIndex, "Has muerto quemado!!", FontTypeNames.FONTTYPE_INFO)
                    .Stats.MinHp = 0
                    Call UserDie(UserIndex)

                End If
                    
                Call WriteUpdateHP(UserIndex)
    
            End If
                
            .Counters.Lava = 0

        End If

    End With

End Sub

''
' Maneja  el efecto del estado atacable
'
' @param UserIndex  El index del usuario a ser afectado por el estado atacable
'

Public Sub EfectoEstadoAtacable(ByVal UserIndex As Integer)
    '******************************************************
    'Author: ZaMa
    'Last Update: 18/09/2010 (ZaMa)
    '18/09/2010: ZaMa - Ahora se activa el seguro cuando dejas de ser atacable.
    '******************************************************

    ' Si ya paso el tiempo de penalizacion
    If Not IntervaloEstadoAtacable(UserIndex) Then
        ' Deja de poder ser atacado
        UserList(UserIndex).flags.AtacablePor = 0
        
        ' Activo el seguro si deja de estar atacable
        If Not UserList(UserIndex).flags.Seguro Then
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn)

        End If
        
        ' Send nick normal
        Call RefreshCharStatus(UserIndex)

    End If
    
End Sub

''
' Maneja el tiempo y el efecto del mimetismo
'
' @param UserIndex  El index del usuario a ser afectado por el mimetismo
'

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)

    '******************************************************
    'Author: Unknown
    'Last Update: 16/09/2010 (ZaMa)
    '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
    '16/09/2010: ZaMa - Se recupera la apariencia de la barca correspondiente despues de terminado el mimetismo.
    '******************************************************
    Dim Barco As ObjData
    
    With UserList(UserIndex)

        If .Counters.Mimetismo < IntervaloInvisible Then
            .Counters.Mimetismo = .Counters.Mimetismo + 1
        Else
            'restore old char
            Call WriteConsoleMsg(UserIndex, "Recuperas tu apariencia normal.", FontTypeNames.FONTTYPE_INFO)
            
            If .flags.Navegando Then
                If .flags.Muerto = 0 Then
                    Call ToggleBoatBody(UserIndex)
                Else
                    .Char.body = iFragataFantasmal
                    .Char.ShieldAnim = NingunEscudo
                    .Char.WeaponAnim = NingunArma
                    .Char.CascoAnim = NingunCasco
                    .Char.AuraAnim = NingunAura
                    .Char.AuraColor = NingunAura

                End If

            Else
                .Char.body = .CharMimetizado.body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                .Char.AuraAnim = .CharMimetizado.AuraAnim
                .Char.AuraColor = .CharMimetizado.AuraColor

            End If
            
            With .Char
                Call ChangeUserChar(UserIndex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraAnim, .AuraColor)

            End With
            
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
            ' Se fue el efecto del mimetismo, puede ser atacado por npcs
            .flags.Ignorado = False

        End If

    End With

End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 16/09/2010 (ZaMa)
    '16/09/2010: ZaMa - Al perder el invi cuando navegas, no se manda el mensaje de sacar invi (ya estas visible).
    '***************************************************

    With UserList(UserIndex)

        If .Counters.Invisibilidad < IntervaloInvisible Then
            .Counters.Invisibilidad = .Counters.Invisibilidad + 1
        Else
            .Counters.Invisibilidad = RandomNumber(-100, 100) ' Invi variable :D
            .flags.invisible = 0

            If .flags.Oculto = 0 Then
                Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                
                ' Si navega ya esta visible..
                If Not .flags.Navegando = 1 Then
                    Call SetInvisible(UserIndex, .Char.CharIndex, False)

                End If
                
            End If

        End If

    End With

End Sub

Public Sub EfectoParalisisNpc(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With Npclist(NPCIndex)

        If .Contadores.Paralisis > 0 Then
            .Contadores.Paralisis = .Contadores.Paralisis - 1
        Else
            .flags.Paralizado = 0
            .flags.Inmovilizado = 0

        End If

    End With

End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With UserList(UserIndex)

        If .Counters.Ceguera > 0 Then
            .Counters.Ceguera = .Counters.Ceguera - 1
        Else

            If .flags.Ceguera = 1 Then
                .flags.Ceguera = 0
                Call WriteBlindNoMore(UserIndex)

            End If

            If .flags.Estupidez = 1 Then
                .flags.Estupidez = 0
                Call WriteDumbNoMore(UserIndex)

            End If
        
        End If

    End With

End Sub

Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/12/2010
    '02/12/2010: ZaMa - Now non-magic clases lose paralisis effect under certain circunstances.
    '***************************************************

    With UserList(UserIndex)
    
        If .Counters.Paralisis > 0 Then
        
            Dim CasterIndex As Integer

            CasterIndex = .flags.ParalizedByIndex
        
            ' Only aplies to non-magic clases
            If .Stats.MaxMAN = 0 Then

                ' Paralized by user?
                If CasterIndex <> 0 Then
                
                    ' Close? => Remove Paralisis
                    If UserList(CasterIndex).Name <> .flags.ParalizedBy Then
                        Call RemoveParalisis(UserIndex)
                        Exit Sub
                        
                        ' Caster dead? => Remove Paralisis
                    ElseIf UserList(CasterIndex).flags.Muerto = 1 Then
                        Call RemoveParalisis(UserIndex)
                        Exit Sub
                    
                    ElseIf .Counters.Paralisis > IntervaloParalizadoReducido Then

                        ' Out of vision range? => Reduce paralisis counter
                        If Not InVisionRangeAndMap(UserIndex, UserList(CasterIndex).Pos) Then
                            ' Aprox. 1500 ms
                            .Counters.Paralisis = IntervaloParalizadoReducido
                            Exit Sub

                        End If

                    End If
                
                    ' Npc?
                Else
                    CasterIndex = .flags.ParalizedByNpcIndex
                    
                    ' Paralized by npc?
                    If CasterIndex <> 0 Then
                    
                        If .Counters.Paralisis > IntervaloParalizadoReducido Then

                            ' Out of vision range? => Reduce paralisis counter
                            If Not InVisionRangeAndMap(UserIndex, Npclist(CasterIndex).Pos) Then
                                ' Aprox. 1500 ms
                                .Counters.Paralisis = IntervaloParalizadoReducido
                                Exit Sub

                            End If

                        End If

                    End If
                    
                End If

            End If
            
            .Counters.Paralisis = .Counters.Paralisis - 1

        Else
            Call RemoveParalisis(UserIndex)

        End If

    End With

End Sub

Public Sub RemoveParalisis(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    'Removes paralisis effect from user.
    '***************************************************
    With UserList(UserIndex)
        .flags.Paralizado = 0
        .flags.Inmovilizado = 0
        .flags.ParalizedBy = vbNullString
        .flags.ParalizedByIndex = 0
        .flags.ParalizedByNpcIndex = 0
        .Counters.Paralisis = 0
        Call WriteParalizeOK(UserIndex)

    End With

End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, _
                      ByRef EnviarStats As Boolean, _
                      ByVal Intervalo As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With UserList(UserIndex)

        Dim massta As Integer

        'Si esta trabajando no recupera energia
        If .flags.MacroTrabajo Then Exit Sub

        If .Stats.MinSta < .Stats.MaxSta Then
            If .Counters.STACounter < Intervalo Then
                .Counters.STACounter = .Counters.STACounter + 1
                
            Else
                EnviarStats = True
                .Counters.STACounter = 0

                If .flags.Desnudo Then Exit Sub 'Desnudo no sube energia. (ToxicWaste)
               
                massta = RandomNumber(1, Porcentaje(.Stats.MaxSta, (.Stats.UserSkills(eSkill.Supervivencia) / 2) + 5))
                .Stats.MinSta = .Stats.MinSta + massta

                If .Stats.MinSta > .Stats.MaxSta Then
                    .Stats.MinSta = .Stats.MaxSta

                End If

            End If

        End If

    End With
    
End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim n As Integer
    
    With UserList(UserIndex)

        If .Counters.Veneno < IntervaloVeneno Then
            .Counters.Veneno = .Counters.Veneno + 1
        Else
            Call WriteConsoleMsg(UserIndex, "Estas envenenado, si no te curas moriras.", FontTypeNames.FONTTYPE_VENENO)
            .Counters.Veneno = 0
            n = RandomNumber(1, 5)
            .Stats.MinHp = .Stats.MinHp - n

            If .Stats.MinHp < 1 Then Call UserDie(UserIndex)
            Call WriteUpdateHP(UserIndex)

        End If

    End With

End Sub

Public Sub EfectoIncinerado(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 05/09/2020
    '
    '***************************************************

    Dim n As Integer
    
    With UserList(UserIndex)

        If .Counters.Quema < IntervaloIncinerado Then
            .Counters.Quema = .Counters.Quema + 1
        Else
            Call WriteConsoleMsg(UserIndex, "Estas ardiendo, si no te apagas moriras.", FontTypeNames.FONTTYPE_WARNING)
            .Counters.Quema = 0
            n = RandomNumber(5, 20)
            .Stats.MinHp = .Stats.MinHp - n

            If .Stats.MinHp < 1 Then Call UserDie(UserIndex)
            Call WriteUpdateHP(UserIndex)

        End If

    End With

End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ??????
    'Last Modification: 11/27/09 (Budi)
    'Cuando se pierde el efecto de la pocion updatea fz y agi (No me gusta que ambos atributos aunque se haya modificado solo uno, pero bueno :p)
    '***************************************************
    With UserList(UserIndex)

        'Controla la duracion de las pociones
        If .flags.DuracionEfecto > 0 Then
            .flags.DuracionEfecto = .flags.DuracionEfecto - 1

            If .flags.DuracionEfecto = 0 Then
                .flags.TomoPocion = False
                .flags.TipoPocion = 0

                'volvemos los atributos al estado normal
                Dim loopX As Integer
                
                For loopX = 1 To NUMATRIBUTOS
                    .Stats.UserAtributos(loopX) = .Stats.UserAtributosBackUP(loopX)
                Next loopX
                
                Call WriteUpdateStrenghtAndDexterity(UserIndex)

            End If

        End If

    End With

End Sub

Public Sub HambreYSed(ByVal UserIndex As Integer, ByRef fenviarAyS As Boolean)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With UserList(UserIndex)

        If Not .flags.Privilegios And PlayerType.User Then Exit Sub
        
        'Sed
        If .Stats.MinAGU > 0 Then
            If .Counters.AGUACounter < IntervaloSed Then
                .Counters.AGUACounter = .Counters.AGUACounter + 1
            Else
                .Counters.AGUACounter = 0
                
                If Lloviendo And TerrainStringToByte(MapInfo(.Pos.Map).Terreno) = eTerrain.terrain_desierto And MapInfo(.Pos.Map).Zona = "BOSQUE" Then
                    .Stats.MinAGU = .Stats.MinAGU - 20
                    Call WriteConsoleMsg(UserIndex, "Estas en una tormenta de arena, sientes el doble de sed.", FontTypeNames.FONTTYPE_INFO)
                Else
                    .Stats.MinAGU = .Stats.MinAGU - 10
                End If
                
                If .Stats.MinAGU <= 0 Then
                    .Stats.MinAGU = 0
                    .flags.Sed = 1

                End If
                
                fenviarAyS = True

            End If

        End If
        
        'hambre
        If .Stats.MinHam > 0 Then
            If .Counters.COMCounter < IntervaloHambre Then
                .Counters.COMCounter = .Counters.COMCounter + 1
            Else
                .Counters.COMCounter = 0
                
                If Lloviendo And TerrainStringToByte(MapInfo(.Pos.Map).Terreno) = eTerrain.terrain_nieve And MapInfo(.Pos.Map).Zona = "BOSQUE" Then
                    .Stats.MinHam = .Stats.MinHam - 20
                    Call WriteConsoleMsg(UserIndex, "Estas en una tormenta de nieve, sientes el doble de hambre.", FontTypeNames.FONTTYPE_INFO)
                Else
                    .Stats.MinHam = .Stats.MinHam - 10
                End If
                
                If .Stats.MinHam <= 0 Then
                    .Stats.MinHam = 0
                    .flags.Hambre = 1

                End If

                fenviarAyS = True

            End If

        End If

    End With

End Sub

Public Sub Sanar(ByVal UserIndex As Integer, _
                 ByRef EnviarStats As Boolean, _
                 ByVal Intervalo As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With UserList(UserIndex)
    
        Dim mashit As Integer

        If .Invent.AnilloEqpObjIndex <> 0 Then
            If ObjData(.Invent.AnilloEqpObjIndex).Efectomagico = eEfectos.AceleraVida Then
                Intervalo = Intervalo - Porcentaje(Intervalo, 40)
            End If
        End If

        'con el paso del tiempo va sanando....pero muy lentamente ;-)
        If .Stats.MinHp < .Stats.MaxHp Then
            If .Counters.HPCounter < Intervalo Then
                .Counters.HPCounter = .Counters.HPCounter + 1
            Else
                mashit = RandomNumber(2, Porcentaje(.Stats.MaxSta, 5))
                
                .Counters.HPCounter = 0
                .Stats.MinHp = .Stats.MinHp + mashit

                If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
                Call WriteConsoleMsg(UserIndex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
                EnviarStats = True

            End If

        End If

    End With

End Sub

Public Sub CargaNpcsDat()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando NPCs.dat."
    
    ' Leemos el NPCs.dat y lo almacenamos en la memoria.
    Set LeerNPCs = New clsIniManager
    Call LeerNPCs.Initialize(DatPath & "NPCs.dat")
    
    'Cargamos el total de NPC
    TotalNPCDat = GetVar(DatPath & "NPCs.dat", "INIT", "NumNPCs")
    
    ' Cargamos la lista de NPC's hostiles disponibles para spawnear.
    Call CargarSpawnList

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo el archivo NPCs.dat."

End Sub


 
Public Function ReiniciarAutoUpdate() As Double
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'WorldSave
    Call ES.DoBackUp

    'commit experiencias
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

End Sub
 
Sub GuardarUsuarios()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    haciendoBK = True
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))
    
    Dim i As Integer

    For i = 1 To LastUser

        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, False)

        End If

    Next i
    
    'se guardan los seguimientos
    Call SaveRecords
    
    Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

    haciendoBK = False

End Sub

Sub SaveUser(ByVal UserIndex As Integer, Optional ByVal SaveTimeOnline As Boolean = True)
    '*************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last modified: 06/12/2018 (CHOTS)
    'Saves the User, in the database or charfile
    '*************************************************

    On Error GoTo ErrorHandler

    With UserList(UserIndex)

        If .clase = 0 Or .Stats.ELV = 0 Then
            Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & .Name)
            Exit Sub

        End If

        If .flags.Mimetizado = 1 Then
            .Char.body = .CharMimetizado.body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            .Char.AuraAnim = .CharMimetizado.AuraAnim
            .Char.AuraColor = .CharMimetizado.AuraColor
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
            ' Se fue el efecto del mimetismo, puede ser atacado por npcs
            .flags.Ignorado = False

        End If

        Dim Prom As Long

        Prom = (-.Reputacion.AsesinoRep) + (-.Reputacion.BandidoRep) + .Reputacion.BurguesRep + (-.Reputacion.LadronesRep) + .Reputacion.NobleRep + .Reputacion.PlebeRep
        Prom = Prom / 6
        .Reputacion.Promedio = Prom
        
        Call SaveUserToDatabase(UserIndex, SaveTimeOnline)

    End With

    Exit Sub

ErrorHandler:
    Call LogError("Error en SaveUser - Userindex: " & UserIndex)

End Sub

Sub LoadUser(ByVal UserIndex As Integer)
    '*************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last modified: 09/10/2018 (CHOTS)
    'Loads the user from the database or charfile
    '*************************************************

    On Error GoTo ErrorHandler

    Call LoadUserFromDatabase(UserIndex)

    With UserList(UserIndex)

        If .flags.Paralizado = 1 Then
            .Counters.Paralisis = IntervaloParalizado

        End If

        'Obtiene el indice-objeto del arma
        If .Invent.WeaponEqpSlot > 0 Then
            .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex

        End If
        
        'Obtiene el indice-objeto del arma
        If .Invent.NudiEqpIndex > 0 Then
            .Invent.NudiEqpIndex = .Invent.Object(.Invent.NudiEqpSlot).ObjIndex

        End If

        'Obtiene el indice-objeto del armadura
        If .Invent.ArmourEqpSlot > 0 Then
            .Invent.ArmourEqpObjIndex = .Invent.Object(.Invent.ArmourEqpSlot).ObjIndex
            .flags.Desnudo = 0
        Else
            .flags.Desnudo = 1

        End If

        'Obtiene el indice-objeto del escudo
        If .Invent.EscudoEqpSlot > 0 Then
            .Invent.EscudoEqpObjIndex = .Invent.Object(.Invent.EscudoEqpSlot).ObjIndex

        End If
        
        'Obtiene el indice-objeto del casco
        If .Invent.CascoEqpSlot > 0 Then
            .Invent.CascoEqpObjIndex = .Invent.Object(.Invent.CascoEqpSlot).ObjIndex

        End If

        'Obtiene el indice-objeto barco
        If .Invent.BarcoSlot > 0 Then
            .Invent.BarcoObjIndex = .Invent.Object(.Invent.BarcoSlot).ObjIndex

        End If

        'Obtiene el indice-objeto municion
        If .Invent.MunicionEqpSlot > 0 Then
            .Invent.MunicionEqpObjIndex = .Invent.Object(.Invent.MunicionEqpSlot).ObjIndex

        End If

        '[Alejo]
        'Obtiene el indice-objeto anilo
        If .Invent.AnilloEqpSlot > 0 Then
            .Invent.AnilloEqpObjIndex = .Invent.Object(.Invent.AnilloEqpSlot).ObjIndex

        End If

        If .Invent.MonturaObjIndex > 0 Then
            .Invent.MonturaObjIndex = .Invent.Object(.Invent.MonturaObjIndex).ObjIndex
        End If

        If .flags.Muerto = 0 Then
            .Char = .OrigChar
        Else
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.WeaponAnim = NingunArma
            .Char.ShieldAnim = NingunEscudo
            .Char.CascoAnim = NingunCasco
            .Char.Heading = eHeading.SOUTH

        End If

    End With

    Exit Sub

ErrorHandler:
    Call LogError("Error en LoadUser: " & UserList(UserIndex).Name & " - " & Err.Number & " - " & Err.description)

End Sub

Public Sub FreeNPCs()

    '***************************************************
    'Autor: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Releases all NPC Indexes
    '***************************************************
    Dim LoopC As Long
    
    ' Free all NPC indexes
    For LoopC = 1 To MAXNPCS
        Npclist(LoopC).flags.NPCActive = False
    Next LoopC

End Sub

Public Sub FreeCharIndexes()
    '***************************************************
    'Autor: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Releases all char indexes
    '***************************************************
    ' Free all char indexes (set them all to 0)
    Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))

End Sub

Public Sub ReproducirSonido(ByVal Destino As SendTarget, _
                            ByVal index As Integer, _
                            ByVal SoundIndex As Integer)
    Call SendData(Destino, index, PrepareMessagePlayWave(SoundIndex, UserList(index).Pos.X, UserList(index).Pos.Y))

End Sub


Public Function Tilde(ByRef data As String) As String

    Dim temp As String

    'Pato
    temp = UCase$(data)
 
    If InStr(1, temp, "Á") Then temp = Replace$(temp, "Á", "A")
   
    If InStr(1, temp, "e") Then temp = Replace$(temp, "e", "E")
   
    If InStr(1, temp, "Í") Then temp = Replace$(temp, "Í", "I")
   
    If InStr(1, temp, "Ó") Then temp = Replace$(temp, "Ó", "O")
   
    If InStr(1, temp, "U") Then temp = Replace$(temp, "U", "U")
   
    Tilde = temp
        
End Function

Public Sub CloseServer()
    
    'Si tenemos la API activada, la matamos.
    If ConexionAPI Then
        Shell ("taskkill /PID " & ApiNodeJsTaskId)
    End If
    
    End
End Sub

Private Sub InicializarSonidos()
'****************************************
'Autor: Lorwik
'Fecha: 01/05/2020
'Descripci�n: Inicializa las variable de los Sonidos
'****************************************

    SND_SWING = 2
    SND_TALAR = 13
    SND_PESCAR = 14
    SND_MINERO = 15
    SND_WARP = 3
    SND_PUERTA = 5
    SND_NIVEL = 128
    SND_USERMUERTE = 11
    SND_IMPACTO = 10
    SND_IMPACTO2 = 12
    SND_LENADOR = 13
    SND_FOGATA = 14
    SND_AVE(1) = 21
    SND_AVE(2) = 22
    SND_AVE(3) = 34
    SND_GRILLO(1) = 28
    SND_GRILLO(2) = 29
    SND_SACARARMA = 25
    SND_ESCUDO(1) = 211
    SND_ESCUDO(2) = 212
    SND_ESCUDO(3) = 213
    SND_ESCUDO(4) = 214
    SND_TRABAJO_HERRERO = 150
    SND_TRABAJO_CARPINTERO = 168
    SND_BEBER = 135
    SND_RESUCITAR_SACERDOTE = 103
    SND_CURAR_SACERDOTE = 104
    SND_DUMMY = 115
    SND_DUMMY2 = 116
    
End Sub

Public Sub LogGlobal(ByVal Str As String)
'***************************************************
'Autor: Lorwik
'Fecha: 09/06/2020
'Descripcion: Guardamos todo lo que se habla por el chat global
'***************************************************

    Dim nfile As Integer
    
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\GlobalChat(" & Month(Date) & "-" & Year(Date) & ").log" For Append Shared As #nfile
    
        Print #nfile, Date & " " & time & " " & Str
        
    Close #nfile

End Sub

Public Sub BanGlobalChatCargar()
'***************************************************
'Autor: Lorwik
'Fecha: 09/06/2020
'Descripcion: Carga la lista de baneados del chat global
'***************************************************
    Dim ArchN As Long
    Dim Tmp As String
    Dim ArchivoLog As String

    ArchivoLog = App.Path & "\Dat\BanGlobalChat.dat"

    Set BanUsersChatGlobal = New Collection

    ArchN = FreeFile()
    Open ArchivoLog For Input As #ArchN

    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        BanUsersChatGlobal.Add Tmp
    Loop

    Close #ArchN
End Sub

Public Sub BanGlobalChatAgregar(ByVal UserName As String)
'***************************************************
'Autor: Lorwik
'Fecha: 09/06/2020
'Descripcion: Agrega un nuevo baneado del chat global
'***************************************************

    BanUsersChatGlobal.Add UserName

    Call BanGlobalChatGuardar
End Sub

Public Function BanGlobalChatBuscar(ByVal UserName As String) As Long
'***************************************************
'Autor: Lorwik
'Fecha: 09/06/2020
'Descripcion: Busca un usuario baneado del chat global de entre la lista
'***************************************************

    Dim Dale As Boolean
    Dim LoopC As Long

    Dale = True
    LoopC = 1
    Do While LoopC <= BanUsersChatGlobal.Count And Dale
        Dale = (BanUsersChatGlobal.Item(LoopC) <> UserName)
        LoopC = LoopC + 1
    Loop

    If Dale Then
        BanGlobalChatBuscar = 0
    Else
        BanGlobalChatBuscar = LoopC - 1
    End If
End Function

Public Function BanGlobalChatQuitar(ByVal UserName As String) As Boolean
'***************************************************
'Autor: Lorwik
'Fecha: 09/06/2020
'Descripcion: Elimina a un usuario baneado del chat global de la lista
'***************************************************
On Error Resume Next

    Dim n As Long

    n = BanGlobalChatBuscar(UserName)
    If n > 0 Then
        BanUsersChatGlobal.Remove n
        BanGlobalChatGuardar
        BanGlobalChatQuitar = True
    Else
        BanGlobalChatQuitar = False
    End If

End Function

Public Sub BanGlobalChatGuardar()
'***************************************************
'Autor: Lorwik
'Fecha: 09/06/2020
'Descripcion: Guarda la lista de usuarios baneados del chat global
'***************************************************

    Dim ArchivoLog As String
    Dim ArchN As Long
    Dim LoopC As Long

    ArchivoLog = App.Path & "\Dat\BanGlobalChat.dat"

    ArchN = FreeFile()
    Open ArchivoLog For Output As #ArchN

    For LoopC = 1 To BanUsersChatGlobal.Count
        Print #ArchN, BanUsersChatGlobal.Item(LoopC)
    Next LoopC

    Close #ArchN
End Sub

Public Function HexToColor(ByRef HexColor As String) As Long
'**********************************
'Autor: Lorwik
'Fecha: 08/03/2021
'Descripcion: Convierte colores Hexadecimales en Long
'**********************************

    ' variable size byte array
    Dim bytHex() As Byte
    ' we only accept one length, 6 characters = 12 bytes
    If LenB(HexColor) = 12 Then
        ' convert string to byte array
        bytHex = HexColor
        ' if a value is now higher than 57, we reduce it by 7
        If bytHex(0) > &H39 Then bytHex(0) = bytHex(0) - 7
        If bytHex(2) > &H39 Then bytHex(2) = bytHex(2) - 7
        If bytHex(4) > &H39 Then bytHex(4) = bytHex(4) - 7
        If bytHex(6) > &H39 Then bytHex(6) = bytHex(6) - 7
        If bytHex(8) > &H39 Then bytHex(8) = bytHex(8) - 7
        If bytHex(10) > &H39 Then bytHex(10) = bytHex(10) - 7
        ' this function is "stupid", it assumes it gets correct data...
        '  makes it faster, but you can give it any string that is 6 characters long, no error, ever!
        '  we take 4 bits for each six characters, and place it in the correct position of a Long,
        '  making up 24 bits that are required to represent a color value
        HexToColor = ((bytHex(0) And &HF&) * &H10&) Or (bytHex(2) And &HF&) _
            Or ((bytHex(4) And &HF&) * &H1000&) Or ((bytHex(6) And &HF&) * &H100&) _
            Or ((bytHex(8) And &HF&) * &H100000) Or ((bytHex(10) And &HF&) * &H10000)
    End If
End Function

