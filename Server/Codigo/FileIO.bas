Attribute VB_Name = "ES"
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

'***************************
'Map format .CSM
'***************************
Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
End Type

Private Type tDatosBloqueados
    X As Integer
    Y As Integer
End Type

Private Type tDatosGrh
    X As Integer
    Y As Integer
    GrhIndex As Long
End Type

Private Type tDatosTrigger
    X As Integer
    Y As Integer
    Trigger As Integer
End Type

Private Type tDatosLuces
    R As Integer
    g As Integer
    b As Integer
    range As Byte
    X As Integer
    Y As Integer
End Type

Private Type tDatosParticulas
    X As Integer
    Y As Integer
    Particula As Long
End Type

Private Type tDatosNPC
    X As Integer
    Y As Integer
    NPCIndex As Integer
End Type

Private Type tDatosObjs
    X As Integer
    Y As Integer
    ObjIndex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTE
    X As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type tMapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type

Private Type tMapDat
    map_name As String
    battle_mode As Boolean
    backup_mode As Boolean
    restrict_mode As String
    music_number As String
    zone As String
    terrain As String
    ambient As String
    lvlMinimo As String
    LuzBase As Long
    version As Long
    NoTirarItems As Boolean
End Type

Public MapSize As tMapSize
Public MapDat As tMapDat
'********************************
'END - Load Map with .CSM format
'********************************

#If False Then

    Dim X, Y, n, Map, Mapa, Email, max, Value As Variant

#End If

Public Sub IniciarCabecera()

    With MiCabecera
        .Desc = "WinterAO Resurrection mod Argentum Online by Noland Studios. http://winterao.com.ar"
        .crc = Rnd * 245
        .MagicWord = Rnd * 92
    End With
    
End Sub

Public Sub CargarSpawnList()
    '****************************************************************************************
    'Author: Unknown
    'Last Modification: 27/03/2020
    'Cargo la lista de NPC's hostiles desde el NPC's.dat
    '****************************************************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Invokar.dat"
    
    ReDim SpawnList(1 To val(LeerNPCs.GetValue("INIT", "NumNPCs"))) As tCriaturasEntrenador
    
    Dim i As Integer: i = 0
    
    Dim LoopC As Long
    For LoopC = 1 To UBound(SpawnList)
        
        If val(LeerNPCs.GetValue("NPC" & LoopC, "Hostile")) = 1 And _
           val(LeerNPCs.GetValue("NPC" & LoopC, "NpcType")) <> 10 Then
            
            i = i + 1
            
            SpawnList(i).NPCIndex = LoopC
            SpawnList(i).NpcName = LeerNPCs.GetValue("NPC" & LoopC, "Name")
            
        End If
        
    Next
    
    ' Hacemos el trim a la lista.
    ReDim Preserve SpawnList(1 To i) As tCriaturasEntrenador
    
    If frmMain.Visible Then frmMain.txtStatus.Text = "Lista de NPC's hostiles se cargo correctamente"
    
End Sub

Function EsAdmin(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsAdmin = (val(Administradores.GetValue("Admin", Name)) = 1)

End Function

Function EsDios(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsDios = (val(Administradores.GetValue("Dios", Name)) = 1)

End Function

Function EsSemiDios(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsSemiDios = (val(Administradores.GetValue("SemiDios", Name)) = 1)

End Function

Function EsGmEspecial(ByRef Name As String) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsGmEspecial = (val(Administradores.GetValue("Especial", Name)) = 1)

End Function

Function EsConsejero(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsConsejero = (val(Administradores.GetValue("Consejero", Name)) = 1)

End Function

Function EsRolesMaster(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsRolesMaster = (val(Administradores.GetValue("RM", Name)) = 1)

End Function

Public Function EsGmChar(ByRef Name As String) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 27/03/2011
    'Returns true if char is administrative user.
    '***************************************************
    
    Dim EsGm As Boolean
    
    ' Admin?
    EsGm = EsAdmin(Name)

    ' Dios?
    If Not EsGm Then EsGm = EsDios(Name)

    ' Semidios?
    If Not EsGm Then EsGm = EsSemiDios(Name)

    ' Consejero?
    If Not EsGm Then EsGm = EsConsejero(Name)

    EsGmChar = EsGm

End Function

Public Sub loadAdministrativeUsers()
    'Admines     => Admin
    'Dioses      => Dios
    'SemiDioses  => SemiDios
    'Especiales  => Especial
    'Consejeros  => Consejero
    'RoleMasters => RM
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Administradores/Dioses/Gms."

    'Si esta mierda tuviese array asociativos el codigo seria tan lindo.
    Dim buf  As Integer

    Dim i    As Long

    Dim Name As String
       
    ' Public container
    Set Administradores = New clsIniManager
    
    ' Server ini info file
    Dim ServerIni As clsIniManager

    Set ServerIni = New clsIniManager
    
    Call ServerIni.Initialize(ConfigPath & "GameMasters.ini")
       
    ' Admines
    buf = val(ServerIni.GetValue("INIT", "Admines"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Admines", "Admin" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Admin", Name, "1")

    Next i
    
    ' Dioses
    buf = val(ServerIni.GetValue("INIT", "Dioses"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Dioses", "Dios" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Dios", Name, "1")
        
    Next i
    
    ' Especiales
    buf = val(ServerIni.GetValue("INIT", "Especiales"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Especiales", "Especial" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Especial", Name, "1")
        
    Next i
    
    ' SemiDioses
    buf = val(ServerIni.GetValue("INIT", "SemiDioses"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("SemiDioses", "SemiDios" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("SemiDios", Name, "1")
        
    Next i
    
    ' Consejeros
    buf = val(ServerIni.GetValue("INIT", "Consejeros"))
        
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Consejeros", "Consejero" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Consejero", Name, "1")
        
    Next i
    
    ' RolesMasters
    buf = val(ServerIni.GetValue("INIT", "RolesMasters"))
        
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("RolesMasters", "RM" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("RM", Name, "1")
    Next i
    
    Set ServerIni = Nothing

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Los Administradores/Dioses/Gms se han cargado correctamente."

End Sub

Public Function GetCharPrivs(ByRef UserName As String) As PlayerType
    '****************************************************
    'Author: ZaMa
    'Last Modification: 18/11/2010
    'Reads the user's charfile and retrieves its privs.
    '***************************************************

    Dim Privs As PlayerType

    If EsAdmin(UserName) Then
        Privs = PlayerType.Admin
        
    ElseIf EsDios(UserName) Then
        Privs = PlayerType.Dios

    ElseIf EsSemiDios(UserName) Then
        Privs = PlayerType.SemiDios
        
    ElseIf EsConsejero(UserName) Then
        Privs = PlayerType.Consejero
    
    Else
        Privs = PlayerType.User

    End If

    GetCharPrivs = Privs

End Function

Public Function TxtDimension(ByVal Name As String) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim n As Integer, cad As String, Tam As Long

    n = FreeFile(1)
    Open Name For Input As #n
    Tam = 0

    Do While Not EOF(n)
        Tam = Tam + 1
        Line Input #n, cad
    Loop
    Close n
    TxtDimension = Tam

End Function

Public Sub CargarForbidenWords()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Nombres prohibidos (NombresInvalidos.txt)."

    ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))

    Dim n As Integer, i As Integer

    n = FreeFile(1)
    Open DatPath & "NombresInvalidos.txt" For Input As #n
    
    For i = 1 To UBound(ForbidenNames)
        Line Input #n, ForbidenNames(i)
    Next i
    
    Close n

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - NombresInvalidos.txt han cargado con exito."

End Sub

Public Sub CargarHechizos()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    '###################################################
    '#               ATENCION PELIGRO                  #
    '###################################################
    '
    '   NO USAR GetVar PARA LEER Hechizos.dat !!!!
    '
    'El que ose desafiar esta LEY, se las tendra que ver
    'con migo. Para leer Hechizos.dat se debera usar
    'la nueva clase clsLeerInis.
    '
    'Alejo
    '
    '###################################################

    On Error GoTo errHandler

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Hechizos."
    
    Dim Hechizo As Integer
    Dim Str     As String
    
    Dim Leer    As clsIniManager

    Set Leer = New clsIniManager
    
    Call Leer.Initialize(DatPath & "Hechizos.dat")
    
    'obtiene el numero de hechizos
    NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))
    
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumeroHechizos
    frmCargando.cargar.Value = 0
    
    'Llena la lista
    For Hechizo = 1 To NumeroHechizos

        With Hechizos(Hechizo)
            .Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
            .Desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
            .PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
            
            .HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
            .TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
            .PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
            
            .Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
            .WAV = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
            .FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
            .Particle = val(Leer.GetValue("Hechizo" & Hechizo, "Particle"))
            
            .loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
            
            '    .Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
            
            .SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
            .MinHp = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
            .MaxHp = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
            
            .SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
            .MiMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
            .MaMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
            
            .SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
            .MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
            .MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
            
            .SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
            .MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
            .MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
            
            .SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
            .MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
            .MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
            
            .SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
            .MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
            .MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
            
            .SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
            .MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
            .MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
            
            .SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
            .MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
            .MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
            
            .Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
            .Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
            .Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
            .RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
            .RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
            .RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
            
            .CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
            .Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
            .CuraQuemaduras = val(Leer.GetValue("Hechizo" & Hechizo, "CuraQuemaduras"))
            .Incinera = val(Leer.GetValue("Hechizo" & Hechizo, "Incinera"))
            .Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
            .RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
            .Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
            .Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
            
            .Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
            .Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
            
            .Warp = val(Leer.GetValue("Hechizo" & Hechizo, "Warp"))
            
            .Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
            .NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
            .cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
            .Mimetiza = val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
            
            '    .Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
            '    .ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
            
            .MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
            .ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
            
            'Barrin 30/9/03
            .StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
            
            .Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
            frmCargando.cargar.Value = frmCargando.cargar.Value + 1
            
            .NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
            .StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
            
            'Portales
            .Portal = val(Leer.GetValue("Hechizo" & Hechizo, "Portal"))
            Str = Leer.GetValue("Hechizo" & Hechizo, "PortalMap")
            
            .PortalPos.Map = val(ReadField(1, Str, 45))
            .PortalPos.X = val(ReadField(2, Str, 45))
            .PortalPos.Y = val(ReadField(3, Str, 45))
            
            .Casteo = val(Leer.GetValue("Hechizo" & Hechizo, "Casteo"))
            .CastFX = val(Leer.GetValue("Hechizo" & Hechizo, "CastFX"))
            
            .RadioArea = val(Leer.GetValue("Hechizo" & Hechizo, "AreaEfecto"))

        End With

    Next Hechizo
    
    Set Leer = Nothing

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Los hechizos se han cargado con exito."
    
    Exit Sub

errHandler:
    MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.description
 
End Sub

Sub LoadMotd()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando archivo MOTD.INI."

    Dim i As Integer

    MaxLines = val(GetVar(ConfigPath & "Motd.ini", "INIT", "NumLines"))
    
    ReDim MOTD(1 To MaxLines)

    For i = 1 To MaxLines
        MOTD(i).texto = GetVar(ConfigPath & "Motd.ini", "Motd", "Line" & i)
        MOTD(i).Formato = vbNullString
    Next i

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - El archivo MOTD.INI fue cargado con exito"

End Sub

Public Sub DoBackUp()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Los hechizos se han cargado con exito."

    haciendoBK = True
    
    ' Lo saco porque elimina elementales y mascotas - Maraxus
    ''''''''''''''lo pongo aca x sugernecia del yind
    'For i = 1 To LastNPC
    '    If Npclist(i).flags.NPCActive Then
    '        If Npclist(i).Contadores.TiempoExistencia > 0 Then
    '            Call MuereNpc(i, 0)
    '        End If
    '    End If
    'Next i
    '''''''''''/'lo pongo aca x sugernecia del yind
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

    Call WorldSave
    Call modGuilds.v_RutinaElecciones
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

    haciendoBK = False
    
    'Log
    On Error Resume Next

    Dim nfile As Integer

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - El WorldSave (backup) se hizo correctamente."

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time
    Close #nfile

End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByRef MAPFILE As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2011
    '10/08/2010 - Pato: Implemento el clsByteBuffer para el grabado de mapas
    '12/01/2011 - Amraphen: Ahora no se hace backup de NPCs prohibidos (Mascotas, Invocados )
    '***************************************************

    On Error Resume Next

    Dim FreeFileMap As Long
    Dim FreeFileInf As Long

    Dim Y           As Long
    Dim X           As Long

    Dim ByFlags     As Byte

    Dim LoopC       As Long

    Dim MapWriter   As clsByteBuffer
    Dim InfWriter   As clsByteBuffer
    Dim IniManager  As clsIniManager

    Dim NpcInvalido As Boolean
    
    Set MapWriter = New clsByteBuffer
    Set InfWriter = New clsByteBuffer
    Set IniManager = New clsIniManager
    
    If FileExist(MAPFILE & ".map", vbNormal) Then
        Call Kill(MAPFILE & ".map")
    End If
    
    If FileExist(MAPFILE & ".inf", vbNormal) Then
        Call Kill(MAPFILE & ".inf")
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    
    Open MAPFILE & ".Map" For Binary As FreeFileMap
    
    Call MapWriter.initializeWriter(FreeFileMap)
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open MAPFILE & ".Inf" For Binary As FreeFileInf
    
    Call InfWriter.initializeWriter(FreeFileInf)
    
    'map Header
    Call MapWriter.putInteger(MapInfo(Map).MapVersion)
        
    Call MapWriter.putString(MiCabecera.Desc, False)
    Call MapWriter.putLong(MiCabecera.crc)
    Call MapWriter.putLong(MiCabecera.MagicWord)
    
    Call MapWriter.putDouble(0)
    
    'inf Header
    Call InfWriter.putDouble(0)
    Call InfWriter.putInteger(0)
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            With MapData(Map, X, Y)
                ByFlags = 0
                
                If .Blocked Then ByFlags = ByFlags Or 1
                If .Graphic(2) Then ByFlags = ByFlags Or 2
                If .Graphic(3) Then ByFlags = ByFlags Or 4
                If .Graphic(4) Then ByFlags = ByFlags Or 8
                If .Trigger Then ByFlags = ByFlags Or 16
                
                Call MapWriter.putByte(ByFlags)
                
                Call MapWriter.putLong(.Graphic(1))
                
                For LoopC = 2 To 4
                    If .Graphic(LoopC) Then Call MapWriter.putLong(.Graphic(LoopC))
                Next LoopC
                
                If .Trigger Then Call MapWriter.putInteger(CInt(.Trigger))
                
                '.inf file
                ByFlags = 0
                
                If .ObjInfo.ObjIndex > 0 Then
                    
                    If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otFogata Then
                        .ObjInfo.ObjIndex = 0
                        .ObjInfo.Amount = 0
                    End If

                End If
    
                If .TileExit.Map Then ByFlags = ByFlags Or 1
                
                ' No hacer backup de los NPCs invalidos ( Mascotas, Invocados )
                If .NPCIndex Then
                    
                    NpcInvalido = (Npclist(.NPCIndex).MaestroUser > 0)
                    
                    If Not NpcInvalido Then ByFlags = ByFlags Or 2

                End If
                
                If .ObjInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
                Call InfWriter.putByte(ByFlags)
                
                If .TileExit.Map Then
                    Call InfWriter.putInteger(.TileExit.Map)
                    Call InfWriter.putInteger(.TileExit.X)
                    Call InfWriter.putInteger(.TileExit.Y)
                End If
                
                If .NPCIndex And Not NpcInvalido Then Call InfWriter.putInteger(Npclist(.NPCIndex).Numero)
                
                If .ObjInfo.ObjIndex Then
                    Call InfWriter.putInteger(.ObjInfo.ObjIndex)
                    Call InfWriter.putInteger(.ObjInfo.Amount)
                End If
                
                NpcInvalido = False

            End With

        Next X
    Next Y
    
    Call MapWriter.saveBuffer
    Call InfWriter.saveBuffer
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf
    
    Set MapWriter = Nothing
    Set InfWriter = Nothing

    With MapInfo(Map)
        'write .dat file
        Call IniManager.ChangeValue("Mapa" & Map, "Name", .Name)
        Call IniManager.ChangeValue("Mapa" & Map, "MusicNum", .music)
        Call IniManager.ChangeValue("Mapa" & Map, "MagiaSinefecto", .MagiaSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "InviSinEfecto", .InviSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "ResuSinEfecto", .ResuSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "StartPos", .StartPos.Map & "-" & .StartPos.X & "-" & .StartPos.Y)
        Call IniManager.ChangeValue("Mapa" & Map, "OnDeathGoTo", .OnDeathGoTo.Map & "-" & .OnDeathGoTo.X & "-" & .OnDeathGoTo.Y)
    
        Call IniManager.ChangeValue("Mapa" & Map, "Terreno", TerrainByteToString(.Terreno))
        Call IniManager.ChangeValue("Mapa" & Map, "Zona", .Zona)
        Call IniManager.ChangeValue("Mapa" & Map, "Restringir", RestrictByteToString(.Restringir))
        Call IniManager.ChangeValue("Mapa" & Map, "BackUp", Str(.BackUp))
    
        If .Pk Then
            Call IniManager.ChangeValue("Mapa" & Map, "Pk", "0")
        Else
            Call IniManager.ChangeValue("Mapa" & Map, "Pk", "1")

        End If
        
        Call IniManager.ChangeValue("Mapa" & Map, "OcultarSinEfecto", .OcultarSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "InvocarSinEfecto", .InvocarSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "RoboNpcsPermitido", .RoboNpcsPermitido)
    
        Call IniManager.DumpFile(MAPFILE & ".dat")

    End With
    
    Set IniManager = Nothing

End Sub

Sub LoadArmasHerreria()

    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))
    
    ReDim Preserve ArmasHerrero(1 To n) As Integer
    
    For lc = 1 To n
        ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
    Next lc

End Sub

Sub LoadArmadurasHerreria()

    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))
    
    ReDim Preserve ArmadurasHerrero(1 To n) As Integer
    
    For lc = 1 To n
        ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
    Next lc

End Sub

Sub LoadObjCarpintero()

    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjCarpintero(1 To n) As Integer
    
    For lc = 1 To n
        ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
    Next lc

End Sub

Sub LoadObjAlquimia()

    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "ObjAlquimia.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjAlquimia(1 To n) As Integer
    
    For lc = 1 To n
        ObjAlquimia(lc) = val(GetVar(DatPath & "ObjAlquimia.dat", "Obj" & lc, "Index"))
    Next lc

End Sub

Sub LoadObjSastre()

    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "ObjSastre.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjSastre(1 To n) As Integer
    
    For lc = 1 To n
        ObjSastre(lc) = val(GetVar(DatPath & "ObjSastre.dat", "Obj" & lc, "Index"))
    Next lc

End Sub

Sub LoadBalance()
    '***************************************************
    'Author: Unknown
    'Last Modification: 15/04/2010
    '15/04/2010: ZaMa - Agrego recompensas faccionarias.
    '***************************************************

    Dim i As Long

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando el archivo Balance.dat"
    
    'Modificadores de Clase
    For i = 1 To NUMCLASES

        With ModClase(i)
            .Evasion = val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
            .AtaqueArmas = val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
            .AtaqueProyectiles = val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
            .AtaqueMarciales = val(GetVar(DatPath & "Balance.dat", "MODATAQUEMARCIALES", ListaClases(i)))
            .DanoArmas = val(GetVar(DatPath & "Balance.dat", "MODDANOARMAS", ListaClases(i)))
            .DanoProyectiles = val(GetVar(DatPath & "Balance.dat", "MODDANOPROYECTILES", ListaClases(i)))
            .DanoMarciales = val(GetVar(DatPath & "Balance.dat", "MODDANOMarciales", ListaClases(i)))
            .Escudo = val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))

        End With

    Next i
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS

        With ModRaza(i)
            .Fuerza = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Fuerza"))
            .Agilidad = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Agilidad"))
            .Inteligencia = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Inteligencia"))
            .Carisma = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Carisma"))
            .Constitucion = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Constitucion"))

        End With

    Next i
    
    'Modificadores de Vida
    For i = 1 To NUMCLASES
        ModVida(i) = val(GetVar(DatPath & "Balance.dat", "MODVIDA", ListaClases(i)))
    Next i
    
    'Distribucion de Vida
    For i = 1 To 5
        DistribucionEnteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "E" + CStr(i)))
    Next i

    For i = 1 To 4
        DistribucionSemienteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "S" + CStr(i)))
    Next i
    
    'Extra
    PorcentajeRecuperoMana = val(GetVar(DatPath & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))

    'Party
    ExponenteNivelParty = val(GetVar(DatPath & "Balance.dat", "PARTY", "ExponenteNivelParty"))
    
    ' Recompensas faccionarias
    For i = 1 To NUM_RANGOS_FACCION
        RecompensaFacciones(i - 1) = val(GetVar(DatPath & "Balance.dat", "RECOMPENSAFACCION", "Rango" & i))
    Next i
    
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo con exito el archivo Balance.dat"

End Sub

Sub LoadOBJData()
    '*****************************************************************************************
    'Author: Unknown
    'Last Modification: 06/02/2020
    '03/02/2020: WyroX - Agrego nivel y skill minimo a ciertos objetos. Nuevas habilidades para anillos
    '06/02/2020: WyroX - MinSkill queda solo para barcos y lingotes (porque tienen una comprobacion especial).
    '                             - Skill requerido modificable para items equipables
    '*****************************************************************************************

    '###################################################
    '#               ATENCION PELIGRO                  #
    '###################################################
    '
    ' NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
    '
    'El que ose desafiar esta LEY, se las tendra que ver
    'con migo. Para leer desde el OBJ.DAT se debera usar
    'la nueva clase clsLeerInis.
    '
    'Alejo
    '
    '###################################################

    'Call LogTarea("Sub LoadOBJData")

    On Error GoTo errHandler

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando base de datos de los objetos."
    
    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim Object      As Integer
    Dim Leer        As clsIniManager
    Dim i           As Integer
    Dim n           As Integer
    Dim S           As String
    Dim Aura        As String
    
    Set Leer = New clsIniManager
    
    Call Leer.Initialize(DatPath & "Obj.dat")
    
    'obtiene el numero de obj
    NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumObjDatas
    frmCargando.cargar.Value = 0
    
    ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    
    'Llena la lista
    For Object = 1 To NumObjDatas

        With ObjData(Object)
            .Name = Leer.GetValue("OBJ" & Object, "Name")
            
            'Pablo (ToxicWaste) Log de Objetos.
            .Log = val(Leer.GetValue("OBJ" & Object, "Log"))
            .NoLog = val(Leer.GetValue("OBJ" & Object, "NoLog"))
            '07/09/07
            
            .GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
            
            .ParticulaIndex = val(Leer.GetValue("OBJ" & Object, "CreaParticulaPiso"))
            
            .OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
            
            .Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
            
            .Subtipo = val(Leer.GetValue("OBJ" & Object, "Subtipo"))
            
            S = Leer.GetValue("OBJ" & Object, "CreaLuz")
            
            .CreaLuz.Rango = val(ReadField(1, S, Asc(":")))
            .CreaLuz.Color = HexToColor(Right(ReadField(2, S, Asc(":")), 6)) 'Hacemos la conversion de colores Hexadecimales a Long
            
            Select Case .OBJType

                Case eOBJType.otArmadura
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .SkSastreria = val(Leer.GetValue("OBJ" & Object, "SkSastreria"))
                    .PielLobo = val(Leer.GetValue("OBJ" & Object, "PielLobo"))
                    .PielOsoPardo = val(Leer.GetValue("OBJ" & Object, "PielOsoPardo"))
                    .PielOsoPolar = val(Leer.GetValue("OBJ" & Object, "PielOsoPolar"))
                    
                    Aura = Leer.GetValue("OBJ" & Object, "CreaGRH")
                    .GrhAura = val(ReadField(1, Aura, Asc(":")))
                    .AuraColor = HexToColor(Right(ReadField(2, Aura, Asc(":")), 6))
                    
                Case eOBJType.otNudillos
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    
                Case eOBJType.otEscudo
                    .ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    
                    Aura = Leer.GetValue("OBJ" & Object, "CreaGRH")
                    .GrhAura = val(ReadField(1, Aura, Asc(":")))
                    .AuraColor = HexToColor(Right(ReadField(2, Aura, Asc(":")), 6))
                
                Case eOBJType.otCasco
                    .CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .SkSastreria = val(Leer.GetValue("OBJ" & Object, "SkSastreria"))
                    .PielLobo = val(Leer.GetValue("OBJ" & Object, "PielLobo"))
                    .PielOsoPardo = val(Leer.GetValue("OBJ" & Object, "PielOsoPardo"))
                    .PielOsoPolar = val(Leer.GetValue("OBJ" & Object, "PielOsoPolar"))
                
                Case eOBJType.otWeapon
                    .WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .Apunala = val(Leer.GetValue("OBJ" & Object, "Apunala"))
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
                    .Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
                    .Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
                    
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    
                    .WeaponRazaEnanaAnim = val(Leer.GetValue("OBJ" & Object, "RazaEnanaAnim"))
                    
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            
                    Aura = Leer.GetValue("OBJ" & Object, "CreaGRH")
                    .GrhAura = val(ReadField(1, Aura, Asc(":")))
                    .AuraColor = HexToColor(Right(ReadField(2, Aura, Asc(":")), 6))
                
                Case eOBJType.otInstrumentos
                    .Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
                    .Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
                    .Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
                    'Pablo (ToxicWaste)
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                    .IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
                    .IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
                    .IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
                
                Case otPociones
                    .TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
                    .MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
                    .MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
                    .DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
                    .SkAlquimia = val(Leer.GetValue("OBJ" & Object, "SkPociones"))
                    .Raices = val(Leer.GetValue("OBJ" & Object, "Raices"))
                
                Case eOBJType.otBarcos
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otFlechas
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))

                Case eOBJType.otMonturas
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))

                Case eOBJType.otAnillo 'Pablo (ToxicWaste)
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))

                    '(WyroX)
                    .ImpideParalizar = val(Leer.GetValue("OBJ" & Object, "ImpideParalizar")) <> 0
                    .ImpideAturdir = val(Leer.GetValue("OBJ" & Object, "ImpideAturdir")) <> 0
                    .ImpideCegar = val(Leer.GetValue("OBJ" & Object, "ImpideCegar")) <> 0
                    '(/WyroX)
                    .Efectomagico = val(Leer.GetValue("OBJ" & Object, "Efectomagico"))
                    .QueAtributo = val(Leer.GetValue("OBJ" & Object, "QueAtributo"))
                    .QueSkill = val(Leer.GetValue("OBJ" & Object, "QueSkill"))
                    
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    
                Case eOBJType.otTeleport
                    .Radio = val(Leer.GetValue("OBJ" & Object, "Radio"))
                    
                Case eOBJType.otPasajes
                    .DesdeMap = val(Leer.GetValue("OBJ" & Object, "Desde"))
                    .HastaMap = val(Leer.GetValue("OBJ" & Object, "Map"))
                    .HastaX = val(Leer.GetValue("OBJ" & Object, "X"))
                    .HastaY = val(Leer.GetValue("OBJ" & Object, "Y"))
                    .NecesitaNave = val(Leer.GetValue("OBJ" & Object, "NecesitaNave"))
                    
                Case eOBJType.otBolsasOro
                    .CuantoAgrega = val(Leer.GetValue("OBJ" & Object, "CuantoAgrega"))
                    
                Case eOBJType.otForos
                    Call AddForum(Leer.GetValue("OBJ" & Object, "ID"))

            End Select
            
            .CuantoAumento = val(Leer.GetValue("OBJ" & Object, "cuantoaumento"))
            
            .Speed = val(Leer.GetValue("OBJ" & Object, "Speed")) 'Cambia la velocidad
            
            .Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
            .HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
            
            .LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
            .MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
            
            .MaxHp = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
            .MinHp = val(Leer.GetValue("OBJ" & Object, "MinHP"))
            
            .Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
            .Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
            
            .MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
            .MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
            
            .MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
            .MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
            .def = (.MinDef + .MaxDef) / 2
            
            .RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
            .RazaDrow = val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
            .RazaElfa = val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
            .RazaGnoma = val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
            .RazaHumana = val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
            
            .valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
            
            .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
            .MinLevel = val(Leer.GetValue("OBJ" & Object, "MinELV"))
            
            .Peso = val(Leer.GetValue("OBJ" & Object, "peso"))
            
            .Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
            
            .Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))

            If .Cerrada = 1 Then
                .Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
                .Clave = val(Leer.GetValue("OBJ" & Object, "Clave"))

            End If
            
            'Puertas y llaves
            .Clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
            
            .texto = Leer.GetValue("OBJ" & Object, "Texto")
            .GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
            
            .Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
            .ForoID = Leer.GetValue("OBJ" & Object, "ID")
            
            .Acuchilla = val(Leer.GetValue("OBJ" & Object, "Acuchilla"))
            
            'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
            i = 1: n = 1
            S = Leer.GetValue("OBJ" & Object, "CP" & i)
            Do While Len(S) > 0
                If ClaseToEnum(S) > 0 Then .ClaseProhibida(n) = ClaseToEnum(S)
                        
                If n = NUMCLASES Then Exit Do
                        
                n = n + 1: i = i + 1
                S = Leer.GetValue("OBJ" & Object, "CP" & i)
            Loop
            
            .DefensaMagicaMax = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
            .DefensaMagicaMin = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
            
            .ResistenciaMagica = val(Leer.GetValue("OBJ" & Object, "ResistenciaMagica"))
            
            .SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
            
            If .SkCarpinteria > 0 Then _
                .Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))

           ' Skill minimo
            S = Leer.GetValue("OBJ" & Object, "SkillRequerido")
            If Len(S) > 0 Then
                .SkillCantidad = val(ReadField(2, S, Asc("-")))
    
                S = Replace(UCase$(ReadField(1, S, Asc("-"))), "+", " ")
                For i = 1 To NUMSKILLS
                    If S = UCase$(SkillsNames(i)) Then
                        .SkillRequerido = i
                        Exit For
                    End If
                Next i
            End If

            'Bebidas
            .MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
            
            .NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
            
            .NoRobable = val(Leer.GetValue("OBJ" & Object, "NoRobable"))
            
            frmCargando.cargar.Value = frmCargando.cargar.Value + 1

        End With

    Next Object
    
    Set Leer = Nothing
    
    ' Inicializo los foros faccionarios
    Call AddForum(FORO_CAOS_ID)
    Call AddForum(FORO_REAL_ID)

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo base de datos de los objetos. Operacion Realizada con exito."
    
    Exit Sub
errHandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.description

End Sub

Sub LoadGlobalDrop()
'**********************************************
'Autor: Lorwik
'Fecha: 01/07/2020
'Descripcion: Carga la lista de drops globales de NPCs
'**********************************************

    On Error GoTo errHandler

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando base de datos de drop globales."
    
    Dim i As Integer
    Dim ln As String

    Dim Leer   As clsIniManager

    Set Leer = New clsIniManager
    
    Call Leer.Initialize(DatPath & "global_drop.dat")
    
    'obtiene el numero de obj
    NUMGLOBALDROPS = val(Leer.GetValue("GLOBAL", "NumDrops"))
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NUMGLOBALDROPS
    frmCargando.cargar.Value = 0
    
    ReDim Preserve GlobalDROPObject(1 To NUMGLOBALDROPS) As GlobalObj
    
    For i = 1 To NUMGLOBALDROPS
    
        GlobalDROPObject(i).ObjIndex = Leer.GetValue("DROP" & i, "ObjIndex")
        
        ln = Leer.GetValue("DROP" & i, "Amount")
        
        GlobalDROPObject(i).MinAmount = val(ReadField(1, ln, Asc("-")))
        GlobalDROPObject(i).MaxAmount = val(ReadField(2, ln, Asc("-")))
        GlobalDROPObject(i).Prob = Leer.GetValue("DROP" & i, "Prob")
    
    Next i
    
    Set Leer = Nothing
    
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo base de datos de los drop globales. Operacion Realizada con exito."
    
    Exit Sub
errHandler:
    MsgBox "error cargando drop globales " & Err.Number & ": " & Err.description

End Sub

Function GetVar(ByVal File As String, _
                ByVal Main As String, _
                ByVal Var As String, _
                Optional EmptySpaces As Long = 1024) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim sSpaces  As String ' This will hold the input that the program will retrieve

    Dim szReturn As String ' This will be the defaul value if the string is not found
      
    szReturn = vbNullString
      
    sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
      
    GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
      
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando backup."
    
    Dim Map       As Integer

    Dim tFileName As String
    
    On Error GoTo man
        
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
        
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.Value = 0
        
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
        
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
        
    For Map = 1 To NumMaps

        If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".Dat", "Mapa" & Map, "BackUp")) <> 0 Then
            tFileName = App.Path & "\WorldBackUp\Mapa" & Map
                
            If Not FileExist(tFileName & ".*") Then 'Miramos que exista al menos uno de los 3 archivos, sino lo cargamos de la carpeta de los mapas
                tFileName = App.Path & MapPath & "Mapa" & Map

            End If

        Else
            tFileName = App.Path & MapPath & "Mapa" & Map

        End If
            
        Call CargarMapa(Map, tFileName)
            
        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        DoEvents
    Next Map
    
    Exit Sub

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se termino de cargar el backup."

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)
 
End Sub

Sub LoadMapData()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando mapas..."
    
    Dim Map       As Integer

    Dim tFileName As String
    
    On Error GoTo man
        
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
        
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.Value = 0
        
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
        
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
          
    For Map = 1 To NumMaps
            
        tFileName = App.Path & MapPath & "Mapa" & Map
        Call CargarMapa(Map, tFileName)
            
        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        DoEvents
    Next Map
    
    Exit Sub

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargaron todos los mapas. Operacion Realizada con exito."

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

End Sub

Public Sub CargarMapa(ByVal Map As Long, ByVal MAPFl As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 10/08/2010
    '***************************************************
    
    On Error GoTo errh
    
    Dim fh              As Integer
    Dim MH              As tMapHeader
    Dim Blqs()          As tDatosBloqueados
    Dim L1()            As Long
    Dim L2()            As tDatosGrh
    Dim L3()            As tDatosGrh
    Dim L4()            As tDatosGrh
    Dim Triggers()      As tDatosTrigger
    Dim Luces()         As tDatosLuces
    Dim Particulas()    As tDatosParticulas
    Dim Objetos()       As tDatosObjs
    Dim NPCs()          As tDatosNPC
    Dim TEs()           As tDatosTE
    Dim MapSize         As tMapSize
    Dim MapDat          As tMapDat
    Dim npcfile         As String
    Dim i               As Long
    Dim j               As Long
    Dim LaCabecera      As tCabecera
    
    fh = FreeFile
    
    Open MAPFl & ".csm" For Binary Access Read As fh
    
        Get #fh, , LaCabecera
    
        Get #fh, , MH
        Get #fh, , MapSize
        Get #fh, , MapDat
        
        ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As Long
        
        Get #fh, , L1
        
        With MH
            If .NumeroBloqueados > 0 Then
                ReDim Blqs(1 To .NumeroBloqueados)
                Get #fh, , Blqs
                For i = 1 To .NumeroBloqueados
                    MapData(Map, Blqs(i).X, Blqs(i).Y).Blocked = 1
                Next i
            End If
            
            If .NumeroLayers(2) > 0 Then
                ReDim L2(1 To .NumeroLayers(2))
                Get #fh, , L2
                For i = 1 To .NumeroLayers(2)
                    MapData(Map, L2(i).X, L2(i).Y).Graphic(2) = L2(i).GrhIndex
                Next i
            End If
            
            If .NumeroLayers(3) > 0 Then
                ReDim L3(1 To .NumeroLayers(3))
                Get #fh, , L3
                For i = 1 To .NumeroLayers(3)
                    MapData(Map, L3(i).X, L3(i).Y).Graphic(3) = L3(i).GrhIndex
                Next i
            End If
            
            If .NumeroLayers(4) > 0 Then
                ReDim L4(1 To .NumeroLayers(4))
                Get #fh, , L4
                For i = 1 To .NumeroLayers(4)
                    MapData(Map, L4(i).X, L4(i).Y).Graphic(4) = L4(i).GrhIndex
                Next i
            End If
            
            If .NumeroTriggers > 0 Then
                ReDim Triggers(1 To .NumeroTriggers)
                Get #fh, , Triggers
                For i = 1 To .NumeroTriggers
                    MapData(Map, Triggers(i).X, Triggers(i).Y).Trigger = Triggers(i).Trigger
                Next i
            End If
            
            If .NumeroParticulas > 0 Then
                ReDim Particulas(1 To .NumeroParticulas)
                Get #fh, , Particulas
            End If
            
            If .NumeroLuces > 0 Then
                ReDim Luces(1 To .NumeroLuces)
                Get #fh, , Luces
            End If
            
            If .NumeroOBJs > 0 Then
                ReDim Objetos(1 To .NumeroOBJs)
                Get #fh, , Objetos
                For i = 1 To .NumeroOBJs
                    MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.ObjIndex = Objetos(i).ObjIndex
                    MapData(Map, Objetos(i).X, Objetos(i).Y).ObjInfo.Amount = Objetos(i).ObjAmmount
                Next i
            End If
                
            If .NumeroNPCs > 0 Then
                ReDim NPCs(1 To .NumeroNPCs)
                Get #fh, , NPCs
                For i = 1 To .NumeroNPCs
                    MapData(Map, NPCs(i).X, NPCs(i).Y).NPCIndex = NPCs(i).NPCIndex
                    If MapData(Map, NPCs(i).X, NPCs(i).Y).NPCIndex > 0 Then
                        
                        npcfile = DatPath & "NPCs.dat"
                        
                        'Si el npc debe hacer respawn en la pos original la guardamos
                        If val(GetVar(npcfile, "NPC" & MapData(Map, NPCs(i).X, NPCs(i).Y).NPCIndex, "PosOrig")) = 1 Then
                            MapData(Map, NPCs(i).X, NPCs(i).Y).NPCIndex = OpenNPC(MapData(Map, NPCs(i).X, NPCs(i).Y).NPCIndex)
                            Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NPCIndex).Orig.Map = Map
                            Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NPCIndex).Orig.X = NPCs(i).X
                            Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NPCIndex).Orig.Y = NPCs(i).Y
                        Else
                            MapData(Map, NPCs(i).X, NPCs(i).Y).NPCIndex = OpenNPC(MapData(Map, NPCs(i).X, NPCs(i).Y).NPCIndex)
                        End If
                        
                        If Not MapData(Map, NPCs(i).X, NPCs(i).Y).NPCIndex = 0 Then
                            Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NPCIndex).Pos.Map = Map
                            Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NPCIndex).Pos.X = NPCs(i).X
                            Npclist(MapData(Map, NPCs(i).X, NPCs(i).Y).NPCIndex).Pos.Y = NPCs(i).Y
       
                            Call MakeNPCChar(True, 0, MapData(Map, NPCs(i).X, NPCs(i).Y).NPCIndex, Map, NPCs(i).X, NPCs(i).Y)
                        End If
                        
                    End If
                Next i
            End If
                
            If .NumeroTE > 0 Then
                ReDim TEs(1 To .NumeroTE)
                Get #fh, , TEs
                For i = 1 To .NumeroTE
                    MapData(Map, TEs(i).X, TEs(i).Y).TileExit.Map = TEs(i).DestM
                    MapData(Map, TEs(i).X, TEs(i).Y).TileExit.X = TEs(i).DestX
                    MapData(Map, TEs(i).X, TEs(i).Y).TileExit.Y = TEs(i).DestY
                Next i
            End If
            
        End With
    
    Close fh
    
        
    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax
            If L1(i, j) > 0 Then
                MapData(Map, j, i).Graphic(1) = L1(j, i)
            End If
        Next i
    Next j
    
    'Cargamos los extras
    With MapInfo(Map)
        .Name = Trim$(MapDat.map_name)
        .music = Trim$(MapDat.music_number)
        .NoTirarItems = MapDat.NoTirarItems

        If MapDat.lvlMinimo = "" Then
            .lvlMinimo = 0
        Else
            .lvlMinimo = Trim$(MapDat.lvlMinimo)
        End If

        .Pk = MapDat.battle_mode
        
        .Terreno = Trim$(MapDat.terrain)
        .Zona = Trim$(MapDat.zone)
        .Restringir = Trim$(RestrictStringToByte(MapDat.restrict_mode))
        .BackUp = MapDat.backup_mode
        
    End With
    
Exit Sub

errh:
    'Call LogError("Error cargando mapa: " & map & " - Pos: " & .X & "," & Y & "." & Err.description)
End Sub


Sub LoadSini()
'***************************************************
'Author: Unknown
'Last Modification: 13/11/2019 (Recox)
'CHOTS: Database params
'Cucsifae: Agregados multiplicadores exp y oro
'CHOTS: Agregado multiplicador oficio
'CHOTS: Agregado min y max Dados
'Jopi: Uso de clsIniManager para cargar los valores.
'Recox: Cargamos si el centinela esta activo o no.
'***************************************************

    Dim Temporal As Long
    
    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    
    If frmMain.Visible Then
        frmMain.txtStatus.Text = "Cargando info de inicio del server."
    End If
    
    Call Lector.Initialize(ConfigPath & "Server.ini")
    
    BootDelBackUp = CBool(val(Lector.GetValue("INIT", "IniciarDesdeBackUp")))
    
    'Misc
    Puerto = val(Lector.GetValue("INIT", "StartPort"))
    LastSockListen = val(Lector.GetValue("INIT", "LastSockListen"))
    HideMe = CBool(Lector.GetValue("INIT", "Hide"))
    AllowMultiLogins = CBool(val(Lector.GetValue("INIT", "AllowMultiLogins")))
    IdleLimit = val(Lector.GetValue("INIT", "IdleLimit"))
    LimiteConexionesPorIp = val(Lector.GetValue("INIT", "LimiteConexionesPorIp"))
    
    'Lee la version correcta del cliente
    ULTIMAVERSION = Lector.GetValue("INIT", "VersionBuildCliente")

    'Esto es para ver si el centinela esta activo o no.
    isCentinelaActivated = CBool(val(Lector.GetValue("INIT", "CentinelaAuditoriaTrabajoActivo")))

    PuedeCrearPersonajes = val(Lector.GetValue("INIT", "PuedeCrearPersonajes"))
    ServerSoloGMs = val(Lector.GetValue("INIT", "ServerSoloGMs"))
    
    ArmaduraImperial1 = val(Lector.GetValue("INIT", "ArmaduraImperial1"))
    ArmaduraImperial2 = val(Lector.GetValue("INIT", "ArmaduraImperial2"))
    ArmaduraImperial3 = val(Lector.GetValue("INIT", "ArmaduraImperial3"))
    TunicaMagoImperial = val(Lector.GetValue("INIT", "TunicaMagoImperial"))
    TunicaMagoImperialEnanos = val(Lector.GetValue("INIT", "TunicaMagoImperialEnanos"))
    ArmaduraCaos1 = val(Lector.GetValue("INIT", "ArmaduraCaos1"))
    ArmaduraCaos2 = val(Lector.GetValue("INIT", "ArmaduraCaos2"))
    ArmaduraCaos3 = val(Lector.GetValue("INIT", "ArmaduraCaos3"))
    TunicaMagoCaos = val(Lector.GetValue("INIT", "TunicaMagoCaos"))
    TunicaMagoCaosEnanos = val(Lector.GetValue("INIT", "TunicaMagoCaosEnanos"))
    
    VestimentaImperialHumano = val(Lector.GetValue("INIT", "VestimentaImperialHumano"))
    VestimentaImperialEnano = val(Lector.GetValue("INIT", "VestimentaImperialEnano"))
    TunicaConspicuaHumano = val(Lector.GetValue("INIT", "TunicaConspicuaHumano"))
    TunicaConspicuaEnano = val(Lector.GetValue("INIT", "TunicaConspicuaEnano"))
    ArmaduraNobilisimaHumano = val(Lector.GetValue("INIT", "ArmaduraNobilisimaHumano"))
    ArmaduraNobilisimaEnano = val(Lector.GetValue("INIT", "ArmaduraNobilisimaEnano"))
    ArmaduraGranSacerdote = val(Lector.GetValue("INIT", "ArmaduraGranSacerdote"))
    
    VestimentaLegionHumano = val(Lector.GetValue("INIT", "VestimentaLegionHumano"))
    VestimentaLegionEnano = val(Lector.GetValue("INIT", "VestimentaLegionEnano"))
    TunicaLobregaHumano = val(Lector.GetValue("INIT", "TunicaLobregaHumano"))
    TunicaLobregaEnano = val(Lector.GetValue("INIT", "TunicaLobregaEnano"))
    TunicaEgregiaHumano = val(Lector.GetValue("INIT", "TunicaEgregiaHumano"))
    TunicaEgregiaEnano = val(Lector.GetValue("INIT", "TunicaEgregiaEnano"))
    SacerdoteDemoniaco = val(Lector.GetValue("INIT", "SacerdoteDemoniaco"))
    
    EnTesting = CBool(Lector.GetValue("INIT", "Testing"))
    
    ContadorAntiPiquete = val(Lector.GetValue("INIT", "ContadorAntiPiquete"))
    MinutosCarcelPiquete = val(Lector.GetValue("INIT", "MinutosCarcelPiquete"))

    'Atributos Iniciales
    EstadisticasInicialesUsarConfiguracionPersonalizada = CBool(val(Lector.GetValue("ESTADISTICASINICIALESPJ", "Activado")))

    'Intervalos
    SanaIntervaloSinDescansar = val(Lector.GetValue("INTERVALOS", "SanaIntervaloSinDescansar"))
    StaminaIntervaloSinDescansar = val(Lector.GetValue("INTERVALOS", "StaminaIntervaloSinDescansar"))
    SanaIntervaloDescansar = val(Lector.GetValue("INTERVALOS", "SanaIntervaloDescansar"))
    StaminaIntervaloDescansar = val(Lector.GetValue("INTERVALOS", "StaminaIntervaloDescansar"))
    StaminaIntervaloLloviendo = val(Lector.GetValue("INTERVALOS", "StaminaIntervaloLloviendo"))
    IntervaloSed = val(Lector.GetValue("INTERVALOS", "IntervaloSed"))
    IntervaloHambre = val(Lector.GetValue("INTERVALOS", "IntervaloHambre"))
    IntervaloVeneno = val(Lector.GetValue("INTERVALOS", "IntervaloVeneno"))
    IntervaloIncinerado = val(Lector.GetValue("INTERVALOS", "IntervaloIncinerado"))
    IntervaloParalizado = val(Lector.GetValue("INTERVALOS", "IntervaloParalizado"))
    IntervaloInvisible = val(Lector.GetValue("INTERVALOS", "IntervaloInvisible"))
    IntervaloFrio = val(Lector.GetValue("INTERVALOS", "IntervaloFrio"))
    IntervaloWavFx = val(Lector.GetValue("INTERVALOS", "IntervaloWAVFX"))
    IntervaloNPCPuedeAtacar = val(Lector.GetValue("INTERVALOS", "IntervaloNpcPuedeAtacar"))
    IntervaloInvocacion = val(Lector.GetValue("INTERVALOS", "IntervaloInvocacion"))
    IntervaloParaConexion = val(Lector.GetValue("INTERVALOS", "IntervaloParaConexion"))
    IntervaloUserPuedeCastear = val(Lector.GetValue("INTERVALOS", "IntervaloLanzaHechizo"))
    IntervaloUserPuedeTrabajar = val(Lector.GetValue("INTERVALOS", "IntervaloTrabajo"))
    IntervaloUserPuedeAtacar = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeAtacar"))
    INTERVALO_GLOBAL = val(Lector.GetValue("INTERVALOS", "IntervaloGlobal"))
    IntervaloPuedeMakrear = val(Lector.GetValue("INTERVALOS", "IntervaloMakreo"))
    
    'TODO : Agregar estos intervalos al form!!!
    IntervaloMagiaGolpe = val(Lector.GetValue("INTERVALOS", "IntervaloMagiaGolpe"))
    IntervaloGolpeMagia = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeMagia"))
    IntervaloGolpeUsar = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeUsar"))
    IntervaloOcultable = val(Lector.GetValue("INTERVALOS", "IntervaloPuedeOcultar"))
    IntervaloTocar = val(Lector.GetValue("INTERVALOS", "IntervaloPuedeTocar"))
    
    '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    IntervaloPuedeSerAtacado = val(Lector.GetValue("TIMERS", "IntervaloPuedeSerAtacado"))
    IntervaloAtacable = val(Lector.GetValue("TIMERS", "IntervaloAtacable"))
    IntervaloOwnedNpc = val(Lector.GetValue("TIMERS", "IntervaloOwnedNpc"))
    

    MinutosWs = val(Lector.GetValue("INTERVALOS", "IntervaloWS"))

    If MinutosWs < 60 Then MinutosWs = 180
    
    MinutosGuardarUsuarios = val(Lector.GetValue("INTERVALOS", "IntervaloGuardarUsuarios"))
    IntervaloCerrarConexion = val(Lector.GetValue("INTERVALOS", "IntervaloCerrarConexion"))
    IntervaloReconexionDB = val(Lector.GetValue("INTERVALOS", "IntervaloReconexionDB"))
    IntervaloUserPuedeUsar = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsar"))
    IntervaloFlechasCazadores = val(Lector.GetValue("INTERVALOS", "IntervaloFlechasCazadores"))
    
    IntervaloOculto = val(Lector.GetValue("INTERVALOS", "IntervaloOculto"))
    
    '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
      
    RecordUsuariosOnline = val(Lector.GetValue("INIT", "Record"))

    ' HappyHour
    Dim lDayNumberTemp As Long
    Dim sDayName As String
    
    iniHappyHourActivado = CBool(val(Lector.GetValue("HAPPYHOUR", "Activado")))
    For lDayNumberTemp = 1 To 7
        sDayName = Lector.GetValue("HAPPYHOUR", "Dia" & lDayNumberTemp)
        HappyHourDays(lDayNumberTemp).Hour = val(ReadField(1, sDayName, 45)) ' GSZAO
        HappyHourDays(lDayNumberTemp).Multi = val(ReadField(2, sDayName, 45)) ' 0.13.5
        
        If HappyHourDays(lDayNumberTemp).Hour < 0 Or HappyHourDays(lDayNumberTemp).Hour > 23 Then
            HappyHourDays(lDayNumberTemp).Hour = 20 ' Hora de 0 a 23.
        End If
        
        If HappyHourDays(lDayNumberTemp).Multi < 0 Then
            HappyHourDays(lDayNumberTemp).Multi = 0
        End If
    Next

    'Conexion con la API hecha en Node.js
    'Mas info aqui: https://github.com/ao-libre/ao-api-server/
    ConexionAPI = CBool(Lector.GetValue("CONEXIONAPI", "Activado"))
    ApiUrlServer = Lector.GetValue("CONEXIONAPI", "UrlServer")
    ApiPath = Lector.GetValue("CONEXIONAPI", "ApiPath")
      
    'Max users
    Temporal = val(Lector.GetValue("INIT", "MaxUsers"))

    If MaxUsers = 0 Then
        MaxUsers = Temporal
        ReDim UserList(1 To MaxUsers) As User

    End If
    
    '&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    'Se agrego en LoadBalance y en el Balance.dat
    'PorcentajeRecuperoMana = val(Lector.GetValue("BALANCE", "PorcentajeRecuperoMana"))
    
    ''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    Call Statistics.Initialize
    
    Set Lector = Nothing
    
    Set ConsultaPopular = New ConsultasPopulares
    Call ConsultaPopular.LoadData

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo la info de inicio del server (Sinfo.ini)"
    
End Sub

Public Sub Load_ConfigDatBase()

    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    
    If frmMain.Visible Then
        frmMain.txtStatus.Text = "Cargando info de inicio del server."
    End If
    
    Call Lector.Initialize(ConfigPath & "DataBase.ini")
    
    Call User_Database.Inicialiar(Lector.GetValue("USER_DATABASE", "DSN"), Lector.GetValue("USER_DATABASE", "Host"), _
                                  Lector.GetValue("USER_DATABASE", "Name"), Lector.GetValue("USER_DATABASE", "Username"), _
                                  Lector.GetValue("USER_DATABASE", "Password"))

    Call Account_Database.Inicialiar(Lector.GetValue("ACC_DATABASE", "DSN"), Lector.GetValue("ACC_DATABASE", "Host"), _
                                     Lector.GetValue("ACC_DATABASE", "Name"), Lector.GetValue("ACC_DATABASE", "Username"), _
                                     Lector.GetValue("ACC_DATABASE", "Password"))
    
    Set Lector = Nothing

End Sub

Public Sub Load_Rates()

    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    
    If frmMain.Visible Then
        frmMain.txtStatus.Text = "Cargando info de inicio del server."
    End If
    
    Call Lector.Initialize(ConfigPath & "Rates.ini")

    STAT_MAXELV = val(Lector.GetValue("INIT", "NivelMaximo"))
    
    ExpMultiplier = val(Lector.GetValue("INIT", "ExpMulti"))
    OroMultiplier = val(Lector.GetValue("INIT", "OroMulti"))
    OficioMultiplier = val(Lector.GetValue("INIT", "OficioMulti"))

    DropItemsAlMorir = CBool(Lector.GetValue("INIT", "DropItemsAlMorir"))
    
    ArtesaniaCosto = val(Lector.GetValue("INIT", "ArtesaniaCosto"))

    DificultadExtraer = val(Lector.GetValue("DIFICULTAD", "DificultadExtraer"))
    
    Set Lector = Nothing
    
End Sub

Sub CargarCiudades()
    
    '***************************************************
    'Author: Jopi
    'Last Modification: 15/05/2019 (Jopi)
    'Jopi: Uso de clsIniManager para cargar los valores.
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Ciudades.dat"
    
    Dim Lector As clsIniManager: Set Lector = New clsIniManager
    
    Call Lector.Initialize(DatPath & "Ciudades.dat")
        
        With Ullathorpe
            .Map = Lector.GetValue("Ullathorpe", "Mapa")
            .X = Lector.GetValue("Ullathorpe", "X")
            .Y = Lector.GetValue("Ullathorpe", "Y")
        End With
        
        With Nix
            .Map = Lector.GetValue("Nix", "Mapa")
            .X = Lector.GetValue("Nix", "X")
            .Y = Lector.GetValue("Nix", "Y")
        End With
        
        With Banderbill
            .Map = Lector.GetValue("Banderbill", "Mapa")
            .X = Lector.GetValue("Banderbill", "X")
            .Y = Lector.GetValue("Banderbill", "Y")
        End With
        
        With Rinkel
            .Map = Lector.GetValue("Rinkel", "Mapa")
            .X = Lector.GetValue("Rinkel", "X")
            .Y = Lector.GetValue("Rinkel", "Y")
        End With
        
        With Lindos
            .Map = Lector.GetValue("Lindos", "Mapa")
            .X = Lector.GetValue("Lindos", "X")
            .Y = Lector.GetValue("Lindos", "Y")
        End With
        
        With Arghal
            .Map = Lector.GetValue("Arghal", "Mapa")
            .X = Lector.GetValue("Arghal", "X")
            .Y = Lector.GetValue("Arghal", "Y")
        End With
        
        With Prision
            .Map = Lector.GetValue("Prision", "Mapa")
            .X = Lector.GetValue("Prision", "X")
            .Y = Lector.GetValue("Prision", "Y")
        End With
        
        With Libertad
            .Map = Lector.GetValue("Prision-Afuera", "Mapa")
            .X = Lector.GetValue("Prision-Afuera", "X")
            .Y = Lector.GetValue("Prision-Afuera", "Y")
        End With
        
        With DungeonNew
            .Map = Lector.GetValue("DungeonNew", "Mapa")
            .X = Lector.GetValue("DungeonNew", "X")
            .Y = Lector.GetValue("DungeonNew", "Y")
        End With

    Set Lector = Nothing
    
    Ciudades(eCiudad.cUllathorpe) = Ullathorpe
    Ciudades(eCiudad.cNix) = Nix
    Ciudades(eCiudad.cBanderbill) = Banderbill
    Ciudades(eCiudad.cRinkel) = Rinkel
    Ciudades(eCiudad.cLindos) = Lindos
    Ciudades(eCiudad.cArghal) = Arghal

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargaron las ciudades.dat"

End Sub

Sub WriteVar(ByVal File As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal Value As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Escribe VAR en un archivo
    '***************************************************

    writeprivateprofilestring Main, Var, Value, File
    
End Sub

Function criminal(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim L As Long
    
    With UserList(UserIndex).Reputacion
        L = (-.AsesinoRep) + (-.BandidoRep) + .BurguesRep + (-.LadronesRep) + .NobleRep + .PlebeRep
        L = L / 6
        criminal = (L < 0)

    End With

End Function

Sub BackUPnPc(ByVal NPCIndex As Integer, ByVal hFile As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 10/09/2010
    '10/09/2010 - Pato: Optimice el BackUp de NPCs
    '***************************************************

    Dim LoopC As Integer
    
    Print #hFile, "[NPC" & Npclist(NPCIndex).Numero & "]"
    
    With Npclist(NPCIndex)
        'General
        Print #hFile, "Name=" & .Name
        Print #hFile, "Desc=" & .Desc
        Print #hFile, "Head=" & val(.Char.Head)
        Print #hFile, "Body=" & val(.Char.body)
        Print #hFile, "ShieldAnim=" & val(.Char.ShieldAnim)
        Print #hFile, "WeaponAnim=" & val(.Char.WeaponAnim)
        Print #hFile, "CascoAnim=" & val(.Char.CascoAnim)
        Print #hFile, "Heading=" & val(.Char.Heading)
        Print #hFile, "Movement=" & val(.Movement)
        Print #hFile, "Attackable=" & val(.Attackable)
        Print #hFile, "Comercia=" & val(.Comercia)
        Print #hFile, "TipoItems=" & val(.TipoItems)
        Print #hFile, "Hostil=" & val(.Hostile)
        Print #hFile, "GiveEXP=" & val(.GiveEXP)
        Print #hFile, "GiveGLD=" & val(.GiveGLD)
        Print #hFile, "InvReSpawn=" & val(.InvReSpawn)
        Print #hFile, "NpcType=" & val(.NPCtype)
        
        'Stats
        Print #hFile, "Alineacion=" & val(.Stats.Alineacion)
        Print #hFile, "DEF=" & val(.Stats.def)
        Print #hFile, "MaxHit=" & val(.Stats.MaxHIT)
        Print #hFile, "MaxHp=" & val(.Stats.MaxHp)
        Print #hFile, "MinHit=" & val(.Stats.MinHIT)
        Print #hFile, "MinHp=" & val(.Stats.MinHp)
        
        'Flags
        Print #hFile, "ReSpawn=" & val(.flags.Respawn)
        Print #hFile, "BackUp=" & val(.flags.BackUp)
        Print #hFile, "Domable=" & val(.flags.Domable)
        Print #hFile, "TiempoRetardoMin= " & val(.flags.TiempoRetardoMin)
        Print #hFile, "TiempoRetardoMax= " & val(.flags.TiempoRetardoMax)
        
        'Inventario
        Print #hFile, "NroItems=" & val(.Invent.NroItems)

        If .Invent.NroItems > 0 Then

            For LoopC = 1 To .Invent.NroItems
                Print #hFile, "Obj" & LoopC & "=" & .Invent.Object(LoopC).ObjIndex & "-" & .Invent.Object(LoopC).Amount & "-" & .Invent.Object(LoopC).RandomDrop
            Next LoopC

        End If
        
        Print #hFile, ""

    End With

End Sub

Sub CargarNpcBackUp(ByVal NPCIndex As Integer, ByVal NpcNumber As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'Status
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando backup Npc"
    
    Dim npcfile As String
    
    'If NpcNumber > 499 Then
    '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
    'Else
    npcfile = DatPath & "bkNPCs.dat"
    'End If
    
    With Npclist(NPCIndex)
    
        .Numero = NpcNumber
        .Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
        .Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
        .Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
        .NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))
        
        .Char.body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
        .Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
        .Char.WeaponAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "WeaponAnim"))
        .Char.CascoAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "CascoAnim"))
        .Char.ShieldAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "ShieldAnim"))
        .Char.Heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))
        
        .Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
        .Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
        .Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
        .GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))
        
        .GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))
        
        .InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))
        
        .Stats.MaxHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
        .Stats.MinHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
        .Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
        .Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
        .Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
        .Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
        
        Dim LoopC As Integer

        Dim ln    As String

        .Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))

        If .Invent.NroItems > 0 Then

            For LoopC = 1 To MAX_INVENTORY_SLOTS
                ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
                .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
                .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
                .Invent.Object(LoopC).RandomDrop = val(ReadField(3, ln, 45))
               
            Next LoopC

        Else

            For LoopC = 1 To MAX_INVENTORY_SLOTS
                .Invent.Object(LoopC).ObjIndex = 0
                .Invent.Object(LoopC).Amount = 0
                .Invent.Object(LoopC).RandomDrop = 0
            Next LoopC

        End If
        
        .flags.NPCActive = True
        .flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
        .flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
        .flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
        .flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))
        
        .flags.TiempoRetardoMax = val(GetVar(npcfile, "NPC" & NpcNumber, "TiempoRetardoMax"))
        .flags.TiempoRetardoMin = val(GetVar(npcfile, "NPC" & NpcNumber, "TiempoRetardoMin"))
        
        'Tipo de items con los que comercia
        .TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

    End With

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo el archivo bkNPCs.dat"

End Sub

Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal Motivo As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call WriteVar(App.Path & "\Dat\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
    Call WriteVar(App.Path & "\Dat\" & "BanDetail.dat", BannedName, "Reason", Motivo)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer

    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub

Public Sub CargaApuestas()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando apuestas.dat"

    Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo el archivo apuestas.dat"

End Sub

Public Function getLimit(ByVal Mapa As Integer, ByVal side As Byte) As Integer

    '***************************************************
    'Author: Budi
    'Last Modification: 31/01/2010
    'Retrieves the limit in the given side in the given map.
    'TODO: This should be set in the .inf map file.
    '***************************************************
    Dim X As Long

    Dim Y As Long

    If Mapa <= 0 Then Exit Function

    For X = 15 To 87
        For Y = 0 To 3

            Select Case side

                Case eHeading.NORTH
                    getLimit = MapData(Mapa, X, 7 + Y).TileExit.Map

                Case eHeading.EAST
                    getLimit = MapData(Mapa, 92 - Y, X).TileExit.Map

                Case eHeading.SOUTH
                    getLimit = MapData(Mapa, X, 94 - Y).TileExit.Map

                Case eHeading.WEST
                    getLimit = MapData(Mapa, 9 + Y, X).TileExit.Map

            End Select

            If getLimit > 0 Then Exit Function
        Next Y
    Next X

End Function

Public Sub LoadArmadurasFaccion()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/04/2010
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando armaduras faccionarias"
    
    Dim ClassIndex    As Long
    
    Dim ArmaduraIndex As Integer
    
    For ClassIndex = 1 To NUMCLASES
    
        ' Defensa minima para armadas altos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinArmyAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        
        ' Defensa minima para armadas bajos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinArmyBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        
        ' Defensa minima para caos altos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinCaosAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        
        ' Defensa minima para caos bajos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinCaosBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
    
        ' Defensa media para armadas altos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedArmyAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        
        ' Defensa media para armadas bajos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedArmyBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        
        ' Defensa media para caos altos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedCaosAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        
        ' Defensa media para caos bajos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedCaosBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
    
        ' Defensa alta para armadas altos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaArmyAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        
        ' Defensa alta para armadas bajos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaArmyBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        
        ' Defensa alta para caos altos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaCaosAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        
        ' Defensa alta para caos bajos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaCaosBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
    
    Next ClassIndex

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo el archivo ArmadurasFaccionarias.dat"

End Sub
