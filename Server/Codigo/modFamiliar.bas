Attribute VB_Name = "modFamiliar"
'MODULO DE FAMILIARES
'******************************
'Autor: Lorwik
'Fecha: 07/04/2021
'******************************

Option Explicit

Public Type tFamily
    Id As Integer
    Nombre As String
    Nivel As Byte
    Exp As Long
    ELU As Long
    Tipo As Integer
    
    MinHp As Long
    MaxHp As Long
    MinHIT As Long
    MaxHit As Long
    
    LanzaSpells As Byte
    Spell(3) As Integer
    
    Paralizado As Byte
    Inmovilizado As Byte
    Muerto As Byte
    
    NpcIndex As Integer
    Invocado As Byte
End Type

Private Enum eTipoFamily
    eFuego = 1
    eAgua
    eTierra
    fFauto
    Ely
    Ent
    Tigre
    Lobo
    Oso
End Enum

Public Enum HabilidadesFamiliar
    HABILIDAD_CURAR = 5
    HABILIDAD_PARA = 9
    HABILIDAD_GOLPE_PARALIZA = 10
    HABILIDAD_GOLPE_ENTORPECE = 11
    HABILIDAD_GOLPE_DESARMA = 12
    HABILIDAD_GOLPE_ENCEGA = 13
    HABILIDAD_GOLPE_ENVENENA = 14
    HABILIDAD_TORMENTA = 15
    HABILIDAD_DESENCANTAR = 22
    HABILIDAD_INMO = 24
    HABILIDAD_DETECTAR = 62
    HABILIDAD_MISIL = 92
    HABILIDAD_DESCARGA = 93
End Enum

Public Function CreateFamiliarNewUser(ByVal UserIndex As Integer, ByVal UserClase As Byte, ByVal PetName As String, ByVal PetTipo As Byte) As Boolean
'************************************
'Autor: Lorwik
'Fecha: 08/04/2021
'Descripción: Crea el familiar para un nuevo usuario
'************************************

    Dim i       As Byte
    Dim Count   As Integer
    Dim nIndex  As Integer

    'Familiares
    If UserClase = eClass.Druid Or UserClase = eClass.Hunter Then
       'Si no llego con alguno de esto familiares o es un error o es un intento de hack
        If PetTipo <> eTipoFamily.Ent And PetTipo <> eTipoFamily.Oso And PetTipo <> eTipoFamily.Tigre And PetTipo <> eTipoFamily.Lobo Then
            CreateFamiliarNewUser = False
            Exit Function
        End If
       
    ElseIf UserClase = eClass.Mage Then
        'Si no llego con alguno de esto familiares o es un error o es un intento de hack
        If PetTipo <> eTipoFamily.Ely And PetTipo <> eTipoFamily.eFuego And PetTipo <> eTipoFamily.eTierra And PetTipo <> eTipoFamily.eAgua And PetTipo <> eTipoFamily.fFauto Then
            CreateFamiliarNewUser = False
            Exit Function
        End If
        
    Else 'Si no es ninguna clase de las anteriores no tiene familiar
        CreateFamiliarNewUser = True
        Exit Function
    
    End If
        
    '¿Nombre invalido?
    If Not AsciiValidos(PetName) Or LenB(PetName) = 0 Then
        Call WriteErrorMsg(UserIndex, "Nombre del familiar invalido.")
        CreateFamiliarNewUser = False
        Exit Function
    
    End If
        
    'Solo permitimos 1 espacio en los nombres
    For i = 1 To Len(PetName)
        If mid(PetName, i, 1) = Chr(32) Then Count = Count + 1
                
    Next i
        
    nIndex = IndexDeFamiliar(PetTipo)
        
    With UserList(UserIndex)
    
        .Familiar.Nombre = PetName
        .Familiar.Tipo = PetTipo
        .Familiar.Nivel = 1
        .Familiar.Exp = 0
        .Familiar.ELU = 300
        .Familiar.MaxHp = val(LeerNPCs.GetValue("NPC" & nIndex, "MaxHP"))
        .Familiar.MinHp = val(LeerNPCs.GetValue("NPC" & nIndex, "MinHP"))
        .Familiar.MaxHit = val(LeerNPCs.GetValue("NPC" & nIndex, "MaxHIT"))
        .Familiar.MinHIT = val(LeerNPCs.GetValue("NPC" & nIndex, "MinHIT"))
        
    End With
    
    CreateFamiliarNewUser = True
    Exit Function
    
End Function

Public Sub ResetFamiliar(ByVal UserIndex As Integer)
'************************************
'Autor: Lorwik
'Fecha: 07/04/2021
'Descripción: Resetea los datos del familiar de un usuario
'************************************

    Dim i As Byte

    With UserList(UserIndex)

        Call RetirarFamiliar(UserIndex, True)
        
        If .Familiar.Invocado = 1 Then _
            Npclist(.Familiar.Id).EsFamiliar = 0
            
        .Familiar.Nombre = vbNullString
        .Familiar.Nivel = 0
        .Familiar.Exp = 0
        .Familiar.ELU = 0
        .Familiar.Tipo = 0
        .Familiar.MinHp = 0
        .Familiar.MaxHp = 0
        .Familiar.MinHIT = 0
        .Familiar.MaxHit = 0
        .Familiar.Paralizado = 0
        .Familiar.Muerto = 0
        .Familiar.NpcIndex = 0
        .Familiar.Invocado = 0
        
        For i = 0 To 3
            .Familiar.Spell(i) = 0
        Next i

    End With
    
End Sub

Public Sub InvocarFamiliar(ByVal UserIndex As Integer, ByVal Cast As Boolean)
'*****************************************
'Autor: Lorwik
'Fecha: 07/04/2021
'Descripción: Invoca al familiar
'*****************************************

    Dim h   As Integer
    
    With UserList(UserIndex)
    
        If MapInfo(UserList(UserIndex).Pos.Map).Pk = False And MapInfo(UserList(UserIndex).Pos.Map).Restringir <> RestrictStringToByte("NEWBIE") Then
            Call WriteConsoleMsg(UserIndex, "No puedes invocar a tu familiar en zonas seguras.", FontTypeNames.FONTTYPE_INFO)
            Cast = False
            Exit Sub

        End If
    
        '¿Familiar muerto?
        If .Familiar.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "Tu familiar esta muerto, acercate al templo mas cercano para que sea resucitado.", FontTypeNames.FONTTYPE_INFO)
            Cast = False
            Exit Sub

        End If

        '¿Podria ser que no tuviera familiar?
        If .Familiar.Tipo < 1 Then
            Call WriteConsoleMsg(UserIndex, "No tienes ningún familiar.", FontTypeNames.FONTTYPE_INFO)
            Call LogError("Error en InvocarFamiliar, el usuario " & .Name & " intento invocar a un familiar que no tiene.")
            Cast = False
            Exit Sub

        End If
        
        h = .flags.Hechizo
        
        If .Familiar.Invocado = 0 Then
        
            '¿Es valido el lugar donde va a invocar el familiar?
            If Not LegalPos(.flags.TargetMap, .flags.TargetX, .flags.TargetY, False, True) Then
                Call WriteConsoleMsg(UserIndex, "No puedes invocar tu familiar ahí.", FontTypeNames.FONTTYPE_INFO)
                Cast = False
                Exit Sub
            End If
        
            If TraerFamiliar(UserIndex, .flags.TargetMap, .flags.TargetX, .flags.TargetY) Then _
                Call InfoHechizo(UserIndex)
        
        Else 'Si ya esta invocamos es que lo quiere retirar
            Call RetirarFamiliar(UserIndex, True)
            
        End If
        
        
        Cast = True
    End With

End Sub

Public Function TraerFamiliar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte) As Boolean
'*****************************************
'Autor: Lorwik
'Fecha: 07/04/2021
'Descripción: Hace aparecer al familiar donde el usuario
'*****************************************

    On Error GoTo TraerFamiliar_Err

    Dim Pos As WorldPos
    
    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y

    With UserList(UserIndex)
    
        .Familiar.NpcIndex = IndexDeFamiliar(.Familiar.Tipo)
        
        'Invocamos el familiar y guardamos el ID
        .Familiar.Id = SpawnNpc(.Familiar.NpcIndex, Pos, False, False)
        
        If .Familiar.Id < 1 Then
            Call WriteConsoleMsg(UserIndex, "No hay espacio aquí para tu mascota.", FontTypeNames.FONTTYPE_INFO)
            Call LogError("Error en InvocarFamiliar, el usuario " & .Name & " intento invocar a un familiar y no puedo hacer spawn.")
            TraerFamiliar = False
            Exit Function

        End If
            
        Call CargarFamiliar(UserIndex)
            
        Npclist(.Familiar.Id).MaestroUser = UserIndex
        Call FollowAmo(.Familiar.Id)
        
    End With

    TraerFamiliar = True
    
    Exit Function
    
TraerFamiliar_Err:
    Call LogError("Error en TraerFamiliar: " & Err.Number & " - " & Err.description)

End Function

Public Sub RetirarFamiliar(ByVal UserIndex As Integer, ByVal DesInvocar As Boolean)
'*****************************************
'Autor: Lorwik
'Fecha: 07/04/2021
'Descripción: Desinvocar a un familiar invocado
'*****************************************

    With UserList(UserIndex)
    
        If .Familiar.Id > 0 Then
        
            .Familiar.MinHp = Npclist(.Familiar.Id).Stats.MinHp
            Call QuitarNPC(.Familiar.Id)
            
        End If
        
        'Quizas solo queremos quitar el familiar y no desinvocarlo
        If DesInvocar Then _
            .Familiar.Invocado = 0
    
    End With
End Sub

Private Sub CargarFamiliar(ByVal UserIndex As Integer)
'*****************************************
'Autor: Lorwik
'Fecha: 07/04/2021
'Descripción: Carga los datos del familiar y los aplica al NPC que se creo previamente
'*****************************************
    Dim i As Byte
    Dim nHabilidades As Byte
    
    With UserList(UserIndex)
    
        Npclist(.Familiar.Id).Name = .Familiar.Nombre
        Npclist(.Familiar.Id).Stats.MinHp = .Familiar.MinHp
        Npclist(.Familiar.Id).Stats.MaxHp = .Familiar.MaxHp
        Npclist(.Familiar.Id).Stats.MinHIT = .Familiar.MinHIT
        Npclist(.Familiar.Id).Stats.MaxHit = .Familiar.MaxHit
        Npclist(.Familiar.Id).EsFamiliar = 1

        Npclist(.Familiar.Id).Movement = TipoAI.SigueAmo
        Npclist(.Familiar.Id).SpeedVar = 50
        Npclist(.Familiar.Id).Target = 0
        Npclist(.Familiar.Id).TargetNPC = 0
        
        
        If FamiliarFisicoMagico(UserIndex) Then '¿Es magico?
            
            nHabilidades = nHabilidadesFamily(UserIndex)
            
            If nHabilidades <> 0 Then
                ReDim Npclist(.Familiar.Id).Spells(1 To nHabilidades)
                
                Npclist(.Familiar.Id).flags.LanzaSpells = nHabilidades
                
                For i = 0 To nHabilidades - 1
                    Npclist(.Familiar.Id).Spells(i + 1) = .Familiar.Spell(i)
                Next i
            
            End If
        Else '¿Es fisico?
        
            If .Familiar.Spell(i) = HabilidadesFamiliar.HABILIDAD_GOLPE_PARALIZA Then
                
                
            ElseIf .Familiar.Spell(i) = HabilidadesFamiliar.HABILIDAD_GOLPE_ENVENENA Then
                Npclist(.Familiar.Id).Veneno = 1
                
            ElseIf .Familiar.Spell(i) = HabilidadesFamiliar.HABILIDAD_GOLPE_DESARMA Then
                Npclist(.Familiar.Id).Desarma = 1
                
            ElseIf .Familiar.Spell(i) = HabilidadesFamiliar.HABILIDAD_GOLPE_ENTORPECE Then
                Npclist(.Familiar.Id).Entorpece = 1
                
            ElseIf .Familiar.Spell(i) = HabilidadesFamiliar.HABILIDAD_GOLPE_ENCEGA Then
                Npclist(.Familiar.Id).Ciega = 1
                
            End If
        End If
        
        .Familiar.Invocado = 1
    
    End With

End Sub

Public Function nHabilidadesFamily(ByVal UserIndex As Integer)
'*****************************************
'Autor: Lorwik
'Fecha: 07/04/2021
'Descripción: Devuelve el numero de habildiades aprendidas por un familiar
'*****************************************
    
    Dim i As Byte
    Dim Count As Byte

    With UserList(UserIndex)
            
        For i = 0 To 3
            If .Familiar.Spell(i) > 0 Then Count = Count + 1
        Next i
    
        nHabilidadesFamily = Count
    
    End With
End Function

Public Sub Familiar_Muerte(ByVal UserIndex As Integer, ByVal Muere As Boolean)
'*****************************************
'Autor: Lorwik
'Fecha: 07/04/2021
'Descripción: Establece el familiar como muerto o vivo
'*****************************************

    With UserList(UserIndex)
    
        If Muere Then
            Call RetirarFamiliar(UserIndex, True)
            .Familiar.Muerto = 1
            .Familiar.MinHp = 0
            Call WriteConsoleMsg(UserIndex, "Tu familiar ha muerto, llevalo a un veterinario para que lo reviva.", FontTypeNames.FONTTYPE_INFO)
            
        Else
            .Familiar.Muerto = 0
            .Familiar.MinHp = .Familiar.MaxHp
            Call WriteConsoleMsg(UserIndex, "¡Tu familiar ha revivido!", FontTypeNames.FONTTYPE_INFO)
            
        End If

    End With
    
End Sub

Public Sub FamiliarAtacaUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
'*****************************************
'Autor: Lorwik
'Fecha: 07/04/2021
'Descripción: Da la orden al familiar de atacar a un usuario
'*****************************************

    With UserList(VictimIndex)
    
        If .Familiar.Invocado = 1 Then
            
            Npclist(.Familiar.Id).flags.AttackedBy = UserList(AttackerIndex).Name
            Npclist(.Familiar.Id).Movement = TipoAI.NPCDEFENSA
            Npclist(.Familiar.Id).Hostile = 1
            
        End If
    
    End With

End Sub

Public Function IndexDeFamiliar(ByVal Tipo As Byte) As Byte
        
On Error GoTo IndexDeFamiliar_Err
        
    Select Case Tipo

        Case eTipoFamily.eFuego
            IndexDeFamiliar = 128 'Elemental de Fuego

        Case eTipoFamily.eAgua
            IndexDeFamiliar = 127 'Elemental de Agua

        Case eTipoFamily.eTierra
            IndexDeFamiliar = 129 'Elemental de Tierra

        Case eTipoFamily.fFauto
            IndexDeFamiliar = 126 'Fuego Fatuo

        Case eTipoFamily.Ely
            IndexDeFamiliar = 132 'Ely

        Case eTipoFamily.Ent
            IndexDeFamiliar = 145 'Ent

        Case eTipoFamily.Tigre
            IndexDeFamiliar = 130 'Tigre

        Case eTipoFamily.Lobo
            IndexDeFamiliar = 133 'Lobo

        Case eTipoFamily.Oso
            IndexDeFamiliar = 131 'Oso Pardo
            
    End Select
 
Exit Function

IndexDeFamiliar_Err:
    Call LogError(Err.Number & " - " & Err.description & " - " & "ModFamiliar.IndexDeFamiliar" & " - " & Erl)
    Resume Next
        
End Function

Public Sub CheckFamilyLevel(ByVal UserIndex As Integer, Optional ByVal PrintInConsole As Boolean = True)
    '*************************************************
    'Author: Lorwik
    'Last modified: 06/09/2019
    'Chequea que el familiar no halla alcanzado el siguiente nivel,
    'de lo contrario le da la vida, y golpe correspodiente.
    '*************************************************
    
    Dim AumentoHIT       As Integer
    Dim AumentoHP        As Integer
    Dim SubiodeLvL       As Boolean
    
    On Error GoTo errHandler
    
    SubiodeLvL = False
    
    With UserList(UserIndex)

        Do While .Familiar.Exp >= .Familiar.ELU
            
            'Checkea si alcanzo el maximo nivel
            If .Familiar.Nivel >= STAT_MAXELV Then
                .Familiar.Exp = 0
                .Familiar.ELU = 0
                Exit Sub

            End If
            
            'Store it!
            Call Statistics.FamilyLevelUp(UserIndex)
            
            If PrintInConsole Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .Pos.X, .Pos.Y))
                Call WriteConsoleMsg(UserIndex, "Tu familiar ha subido de nivel!", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .Familiar.Nivel = .Stats.ELV + 1
            
            .Familiar.Exp = .Familiar.Exp - .Familiar.ELU
                  
            'Tabla de experiencia de la 0.99z
            If .Familiar.Nivel < 11 Then
                .Familiar.ELU = .Familiar.ELU * 1.3
                
            ElseIf .Familiar.Nivel < 25 Then
                .Familiar.ELU = .Familiar.ELU * 1.4
                
            Else
                .Familiar.ELU = .Familiar.ELU * 1.1

            End If
        
            Select Case .Familiar.Tipo

                Case eTipoFamily.Ent
                    AumentoHP = 18
                    AumentoHIT = 5
                    
                Case eTipoFamily.Oso
                    AumentoHP = 35
                    AumentoHIT = 6
                    
                Case eTipoFamily.Tigre
                    AumentoHP = 20
                    AumentoHIT = 6
                    
                Case eTipoFamily.Lobo
                    AumentoHP = 23
                    AumentoHIT = 5
                    
                Case eTipoFamily.Ely
                    AumentoHP = 20
                    AumentoHIT = 3
                    
                Case eTipoFamily.eFuego
                    AumentoHP = 25
                    AumentoHIT = 4
                    
                Case eTipoFamily.eAgua
                    AumentoHP = 25
                    AumentoHIT = 4
                    
                Case eTipoFamily.fFauto
                    AumentoHP = 15
                    AumentoHIT = 2
                
                Case Else
                    AumentoHP = 18
                    AumentoHIT = 5

            End Select
            
            'Actualizamos HitPoints
            .Familiar.MaxHp = .Familiar.MaxHp + AumentoHP

            If .Familiar.MaxHp > STAT_MAXHP Then .Familiar.MaxHp = STAT_MAXHP
            
            'Actualizamos Golpe Maximo
            .Familiar.MaxHit = .Familiar.MaxHit + AumentoHIT
            
            'Actualizamos Golpe Minimo
            .Familiar.MinHIT = .Familiar.MinHIT + AumentoHIT
            
            'Notificamos al user
            If PrintInConsole Then
                If AumentoHP > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Tu familiar ha ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)

                End If

                If AumentoHIT > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El golpe maximo de tu familiar aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteConsoleMsg(UserIndex, "El golpe minimo de tu familiar aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)

                End If
            End If
            
            '¿Aprende alguna habilidad?
            Call FamilirAprendeHabilidad(UserIndex)
            
            'Marcamos que subio del lvl
            SubiodeLvL = True
            
            Call LogDesarrollo("El familiar de " & .Name & " paso a nivel " & .Stats.ELV & " gano HP: " & AumentoHP)
            
            .Familiar.MinHp = .Familiar.MaxHp

        Loop
        
    End With
    
    'Aqui llamar paquete donde se deberia actualizar la info del familiar en el cliente
    
    Exit Sub

errHandler:
    Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.description)

End Sub

Sub WarpFamiliar(ByVal UserIndex As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 06/09/2019
'Teletransporta al familiar a la ubicacion del usuario
'*************************************************
    On Error GoTo WarpFamiliar_Err
        
    With UserList(UserIndex)

        If .Familiar.Invocado = 1 Then
            Call RetirarFamiliar(UserIndex, False)
            
             If MapInfo(UserList(UserIndex).Pos.Map).Pk = False And MapInfo(UserList(UserIndex).Pos.Map).Restringir <> RestrictStringToByte("NEWBIE") Then
                Call WriteConsoleMsg(UserIndex, "No se permiten familiares en zona segura. " & .Familiar.Nombre & " te esperará afuera.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call TraerFamiliar(UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
            
        End If
    
    End With
            
    Exit Sub

WarpFamiliar_Err:
    Call LogError(Err.Number & " - " & Err.description & " - " & "modFamiliar.WarpFamiliar" & " - " & Erl)
    Resume Next
        
End Sub

Public Function FamiliarFisicoMagico(ByVal UserIndex As Integer) As Boolean
'*************************************************
'Author: Lorwik
'Last modified: 08/09/2019
'Determina si el familiar es de tipo fisico o magico
'False = Fisico
'True = Magico
'*************************************************

    Select Case UserList(UserIndex).Familiar.Tipo
    
        Case eTipoFamily.Ent, eTipoFamily.Oso, eTipoFamily.Tigre, eTipoFamily.Lobo
            FamiliarFisicoMagico = False
            Exit Function
            
        Case eTipoFamily.Ely, eTipoFamily.eFuego, eTipoFamily.eTierra, eTipoFamily.eAgua, eTipoFamily.fFauto
            FamiliarFisicoMagico = True
            Exit Function
            
    End Select
    
End Function

Public Sub FamilirAprendeHabilidad(ByVal UserIndex As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 09/09/2019
'El familiar aprende su habilidad correspondiente a su nivel y tipo
'*************************************************

    Dim Aprendio As Byte

    With UserList(UserIndex)
    
        Select Case .Familiar.Nivel
        
            Case 10
                If .Familiar.Tipo = eTipoFamily.Ely Then
                    .Familiar.Spell(0) = HabilidadesFamiliar.HABILIDAD_CURAR
                    Aprendio = 1
                    
                ElseIf .Familiar.Tipo = eTipoFamily.fFauto Then
                    .Familiar.Spell(0) = HabilidadesFamiliar.HABILIDAD_MISIL
                    Aprendio = 1
                    
                End If
            
            Case 15
                If .Familiar.Tipo = eTipoFamily.Ent Then
                    .Familiar.Spell(1) = HabilidadesFamiliar.HABILIDAD_GOLPE_ENVENENA
                    Aprendio = 2
                    
                ElseIf .Familiar.Tipo = eTipoFamily.Tigre Or .Familiar.Tipo = eTipoFamily.Lobo Then
                    .Familiar.Spell(1) = HabilidadesFamiliar.HABILIDAD_GOLPE_ENTORPECE
                    Aprendio = 2
                    
                ElseIf .Familiar.Tipo = eTipoFamily.Ely Or .Familiar.Tipo = eTipoFamily.fFauto Then
                    .Familiar.Spell(1) = HabilidadesFamiliar.HABILIDAD_INMO
                    Aprendio = 2
                    
                End If
            
            Case 20
                If .Familiar.Tipo = eTipoFamily.Ent Or .Familiar.Tipo = eTipoFamily.Tigre Or .Familiar.Tipo = eTipoFamily.Lobo Then
                    .Familiar.Spell(2) = HabilidadesFamiliar.HABILIDAD_GOLPE_PARALIZA
                    Aprendio = 3
                    
                ElseIf .Familiar.Tipo = eTipoFamily.Ely Or .Familiar.Tipo = eTipoFamily.fFauto Then
                    .Familiar.Spell(2) = HabilidadesFamiliar.HABILIDAD_DESCARGA
                    Aprendio = 3
                    
                ElseIf .Familiar.Tipo = eTipoFamily.eFuego Then
                    .Familiar.Spell(2) = HabilidadesFamiliar.HABILIDAD_TORMENTA
                    Aprendio = 3
                    
                ElseIf .Familiar.Tipo = eTipoFamily.eTierra Then
                    .Familiar.Spell(2) = HabilidadesFamiliar.HABILIDAD_INMO
                    Aprendio = 3
                    
                ElseIf .Familiar.Tipo = eTipoFamily.eAgua Then
                    .Familiar.Spell(2) = HabilidadesFamiliar.HABILIDAD_PARA
                    Aprendio = 3
                    
                End If
            
            Case 30
                
                If .Familiar.Tipo = eTipoFamily.Ent Or .Familiar.Tipo = eTipoFamily.Oso Then
                    .Familiar.Spell(3) = HabilidadesFamiliar.HABILIDAD_GOLPE_DESARMA
                    Aprendio = 4
                    
                ElseIf .Familiar.Tipo = eTipoFamily.Tigre Or .Familiar.Tipo = eTipoFamily.Lobo Then
                    .Familiar.Spell(3) = HabilidadesFamiliar.HABILIDAD_GOLPE_ENCEGA
                    Aprendio = 4
                    
                ElseIf .Familiar.Tipo = eTipoFamily.Ely Then
                    .Familiar.Spell(3) = HabilidadesFamiliar.HABILIDAD_DESENCANTAR
                    Aprendio = 4
                    
                ElseIf .Familiar.Tipo = eTipoFamily.fFauto Then
                    .Familiar.Spell(3) = HabilidadesFamiliar.HABILIDAD_DETECTAR
                    Aprendio = 4
                    
                End If
                
        End Select
        
        If Aprendio > 0 Then _
            Call WriteConsoleMsg(UserIndex, "¡" & .Familiar.Nombre & " aprendio " & HabilidadName(.Familiar.Spell(Aprendio - 1)) & "!", FontTypeNames.FONTTYPE_INFO)
    
    End With

End Sub

Public Function HabilidadName(ByVal Habilidad As Integer) As String
'*************************************************
'Author: Lorwik
'Last modified: 09/09/2019
'Devuelve el nombre de la habilidad
'*************************************************

    Select Case Habilidad
        Case HABILIDAD_INMO
            HabilidadName = "Inmoviliza"
            
        Case HABILIDAD_PARA
            HabilidadName = "Paraliza"
            
        Case HABILIDAD_DESCARGA
            HabilidadName = "Lanza descargas"
            
        Case HABILIDAD_TORMENTA
            HabilidadName = "Tormenta de fuego"
            
        Case HABILIDAD_DESENCANTAR
            HabilidadName = "Desencanta al amo"
            
        Case HABILIDAD_CURAR
            HabilidadName = "Cura al amo"
            
        Case HABILIDAD_MISIL
            HabilidadName = "Lanza misiles mágicos"
            
        Case HABILIDAD_DETECTAR
            HabilidadName = "Detecta invisibles"
            
        Case HABILIDAD_GOLPE_PARALIZA
            HabilidadName = "Paraliza con los golpes"
            
        Case HABILIDAD_GOLPE_ENTORPECE
            HabilidadName "Entorpece con los golpes"
            
        Case HABILIDAD_GOLPE_DESARMA
            HabilidadName = "Desarma con los golpes"
            
        Case HABILIDAD_GOLPE_ENCEGA
            HabilidadName = "Encega con los golpes"
            
        Case HABILIDAD_GOLPE_ENVENENA
            HabilidadName = "Envenena con los golpes"
            
        Case Else
            HabilidadName = "Desconocida (" & Habilidad & ")"
    End Select

End Function

Public Function FamiliarPuedeCurar(ByVal UserIndex As Integer) As Integer
'*************************************************
'Author: Lorwik
'Last modified: 09/09/2019
'Devuelve si el familiar tiene la habilidad de curar
'*************************************************

    Dim i As Byte
    Dim nHabilidades As Byte

    With UserList(UserIndex)
    
        nHabilidades = nHabilidadesFamily(UserIndex)
        
        If nHabilidades <> 0 Then
        
            For i = 0 To nHabilidades - 1
            
                If .Familiar.Spell(i) = HabilidadesFamiliar.HABILIDAD_CURAR Then
                    FamiliarPuedeCurar = HabilidadesFamiliar.HABILIDAD_CURAR
                    Exit Function
                End If
            
            Next i
        
        End If
    
    End With
    
    FamiliarPuedeCurar = 0

End Function

Public Sub FamiliarLanzaHechizo(ByVal NpcIndex As Integer, ByVal NpcVictima As Integer)

    Dim FamiliarCura As Integer
    
    Dim Prob         As Byte
                                         
    With Npclist(NpcIndex)
                                         
        Prob = RandomNumber(1, 100)

        If Prob < 30 Then
                                         
            FamiliarCura = FamiliarPuedeCurar(.MaestroUser) '¿Puede curar?
                                            
            'Si puede curar... ¿El maestro esta cascao?
            If FamiliarCura <> 0 And UserList(.MaestroUser).Stats.MinHp <> UserList(.MaestroUser).Stats.MaxHp Then
                                                
                Call NpcLanzaUnSpell(NpcIndex, .MaestroUser, FamiliarCura)
                                                
            ElseIf Npclist(NpcIndex).flags.LanzaSpells <> 0 Then
                                            
                Call NpcLanzaUnSpellSobreNpc(NpcIndex, NpcVictima)
                                                
            End If

        End If

    End With

End Sub
