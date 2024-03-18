Attribute VB_Name = "modDatabase"
'Modulo gestor de la base de datos.
'Adaptado y mejorado por Lorwik

Option Explicit

Sub SaveUserToDatabase(ByVal UserIndex As Integer, _
                       Optional ByVal SaveTimeOnline As Boolean = True)
    '*************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last modified: 14/10/2018
    'Saves the User to the database
    '*************************************************

    On Error GoTo ErrorHandler

    With UserList(UserIndex)

        If .Id > 0 Then
            Call UpdateUserToDatabase(UserIndex, SaveTimeOnline)
        Else
            Call InsertUserToDatabase(UserIndex, SaveTimeOnline)

        End If

    End With
    
    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Unable to save User to Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)

End Sub

Sub InsertUserToDatabase(ByVal UserIndex As Integer, _
                         Optional ByVal SaveTimeOnline As Boolean = True)
    '*************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last modified: 04/10/2018
    'Inserts a new user to the database, then gets its ID and assigns it
    '*************************************************

    On Error GoTo ErrorHandler

    Dim query  As String

    Dim UserID As Integer

    Dim LoopC  As Integer

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    'Basic user data
    query = "INSERT INTO personaje SET "
    query = query & "name = (?), cuenta_id = (?), level = (?), exp = (?), elu = (?), genre_id = (?), race_id = (?), class_id = (?), "
    query = query & "home_id = (?), description = (?), gold = (?), free_skillpoints = (?), assigned_skillpoints = (?), pos_map = (?), pos_x = (?), pos_y = (?), "
    query = query & "body_id = (?), head_id = (?), weapon_id = (?), helmet_id = (?), shield_id = (?), items_amount = (?), slot_armour = (?), slot_weapon = (?), "
    query = query & "slot_nudillos = (?), min_hp = (?), max_hp = (?), min_man = (?), max_man = (?), min_sta = (?), max_sta = (?), min_ham = (?), "
    query = query & "max_ham = (?), min_sed = (?), max_sed = (?), min_hit = (?), max_hit = (?), rep_noble = (?), rep_plebe = (?), Pareja = (?), rep_average = (?)"

    With UserList(UserIndex)
    
        Call User_Database.MakeQuery(query, True, _
                                    .Name, .AccountInfo.Id, .Stats.ELV, .Stats.Exp, .Stats.ELU, .Genero, .Raza, .clase, _
                                    .Hogar, .Desc, .Stats.Gld, .Stats.SkillPts, .Counters.AsignedSkills, .Pos.Map, .Pos.X, .Pos.Y, _
                                    .Char.body, .Char.Head, .Char.WeaponAnim, .Char.CascoAnim, .Char.ShieldAnim, .Invent.NroItems, .Invent.ArmourEqpSlot, .Invent.WeaponEqpSlot, _
                                    .Invent.NudiEqpSlot, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMAN, .Stats.MaxMAN, .Stats.MinSta, .Stats.MaxSta, .Stats.MinHam, _
                                    .Stats.MaxHam, .Stats.MinAGU, .Stats.MaxAGU, .Stats.MinHIT, .Stats.MaxHit, .Reputacion.NobleRep, .Reputacion.PlebeRep, .flags.miPareja, .Reputacion.Promedio)

        'Get the user ID
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute("SELECT LAST_INSERT_ID();")

        If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
            UserID = 1

        End If

        UserID = val(User_Database.Database_RecordSet.Fields(0).Value)
        Set User_Database.Database_RecordSet = Nothing

        .Id = UserID
        
        '*******************************************************************
        'Familiar
        '*******************************************************************
        query = "INSERT INTO familiar (user_id, nombre, level, exp, elu, tipo, min_hp, max_hp, min_hit, max_hit) VALUES (?,?,?,?,?,?,?,?,?,?)"
        
        Call User_Database.MakeQuery(query, True, .Id, .Familiar.Nombre, .Familiar.Nivel, .Familiar.Exp, .Familiar.ELU, .Familiar.Tipo, .Familiar.MinHp, .Familiar.MaxHp, .Familiar.MinHIT, .Familiar.MaxHit)

        '*******************************************************************
        'Atributos
        '*******************************************************************
        query = "INSERT INTO atributos (user_id, "

        For LoopC = 1 To NUMATRIBUTOS
            query = query & " att" & LoopC
            If LoopC < NUMATRIBUTOS Then query = query & ", "
        Next LoopC

        query = query & ") VALUES (" & .Id & ", "

        For LoopC = 1 To NUMATRIBUTOS
            query = query & .Stats.UserAtributos(LoopC)
            If LoopC < NUMATRIBUTOS Then query = query & ", "

        Next LoopC
        
        query = query & ");"

        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Hechizos
        '*******************************************************************
        query = "INSERT INTO spell (user_id, "

        For LoopC = 1 To MAXUSERHECHIZOS
            query = query & " spell_id" & LoopC
            If LoopC < MAXUSERHECHIZOS Then query = query & ", "
        Next LoopC

        query = query & ") VALUES (" & .Id & ", "

        For LoopC = 1 To MAXUSERHECHIZOS
            query = query & .Stats.UserHechizos(LoopC)
            If LoopC < MAXUSERHECHIZOS Then query = query & ", "

        Next LoopC
        
        query = query & ");"

        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Inventario
        '*******************************************************************
        query = "INSERT INTO inventario_items (user_id, "
        
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            query = query & "item_id" & LoopC & ", amount" & LoopC & ", is_equipped" & LoopC
            If LoopC < MAX_INVENTORY_SLOTS Then query = query & ", "
        Next LoopC
        
        query = query & ") VALUES (" & .Id & ", "
        
        For LoopC = 1 To MAX_INVENTORY_SLOTS

            query = query & .Invent.Object(LoopC).ObjIndex & ", "
            query = query & .Invent.Object(LoopC).Amount & ", "
            query = query & .Invent.Object(LoopC).Equipped
            If LoopC < MAX_INVENTORY_SLOTS Then query = query & ", "
            
        Next LoopC
        
        query = query & ");"

        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Boveda
        '*******************************************************************
        query = "INSERT INTO banco_items (user_id, "
        
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            query = query & "item_id" & LoopC & ", amount" & LoopC
            If LoopC < MAX_BANCOINVENTORY_SLOTS Then query = query & ", "
        Next LoopC
        
        query = query & ") VALUES (" & .Id & ", "
        
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

            query = query & .BancoInvent.Object(LoopC).ObjIndex & ", "
            query = query & .BancoInvent.Object(LoopC).Amount
            If LoopC < MAX_BANCOINVENTORY_SLOTS Then query = query & ", "
            
        Next LoopC
        
        query = query & ");"

        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Skills
        '*******************************************************************
        query = "INSERT INTO skillpoint (user_id, "
        
        For LoopC = 1 To NUMSKILLS
            query = query & "sk" & LoopC & ", exp" & LoopC & ", elu" & LoopC
            If LoopC < NUMSKILLS Then query = query & ", "
        Next LoopC
        
        query = query & ") VALUES (" & .Id & ", "

        For LoopC = 1 To NUMSKILLS
            query = query & .Stats.UserSkills(LoopC) & ", "
            query = query & .Stats.ExpSkills(LoopC) & ", "
            query = query & .Stats.EluSkills(LoopC)
            If LoopC < NUMSKILLS Then query = query & ", "

        Next LoopC
        
        query = query & ");"

        Call User_Database.Database_Connection.Execute(query)
        
        '*******************************************************************
        'Mascotas
        '*******************************************************************
        query = "INSERT INTO pet (user_id, "
        
        For LoopC = 1 To MAXMASCOTAS
            query = query & "pet" & LoopC
            If LoopC < MAXMASCOTAS Then query = query & ", "
        Next LoopC

        query = query & ") VALUES (" & .Id & ", "

        For LoopC = 1 To MAXMASCOTAS
            query = query & .MascotasIndex(LoopC)
            If LoopC < MAXMASCOTAS Then query = query & ", "
        Next LoopC

        query = query & ");"

        Call User_Database.Database_Connection.Execute(query)
        
        '*******************************************************************
        'Macros
        '*******************************************************************
        query = "INSERT INTO macros (user_id, "
        
        For LoopC = 1 To NUMMACROS
            query = query & "tipoaccion" & LoopC & ", "
            query = query & "spell" & LoopC & ", "
            query = query & "inv" & LoopC & ", "
            query = query & "command" & LoopC
            If LoopC < NUMMACROS Then query = query & ", "
        Next LoopC

        query = query & ") VALUES (" & .Id & ", "

        For LoopC = 1 To NUMMACROS
            query = query & .MacrosKey(LoopC).TipoAccion & ", "
            query = query & .MacrosKey(LoopC).hList & ", "
            query = query & .MacrosKey(LoopC).InvObj & ", "
            query = query & "'" & .MacrosKey(LoopC).Comando & "'"
            
            If LoopC < NUMMACROS Then query = query & ", "
        Next LoopC

        query = query & ");"

        Call User_Database.Database_Connection.Execute(query)

    End With
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Unable to INSERT User to Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)

End Sub

Sub UpdateUserToDatabase(ByVal UserIndex As Integer, _
                         Optional ByVal SaveTimeOnline As Boolean = True)
    '*************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last modified: 04/10/2018
    'Updates an existing user in the database
    '*************************************************

    On Error GoTo ErrorHandler

    Dim query  As String

    Dim UserID As Integer

    Dim LoopC  As Integer

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    'Basic user data
        
        query = "UPDATE personaje SET "
        query = query & "name = (?), level = (?), exp = (?), elu = (?), genre_id = (?), race_id = (?), class_id = (?), home_id = (?), description = (?), gold = (?), bank_gold = (?), free_skillpoints = (?), "
        query = query & "assigned_skillpoints = (?), pet_amount = (?), pos_map = (?), pos_x = (?), pos_y = (?), last_map = (?), body_id = (?), head_id = (?), weapon_id = (?), helmet_id = (?), shield_id = (?), "
        query = query & "aura_id = (?), aura_color = (?), heading = (?), items_amount = (?), slot_armour = (?), slot_weapon = (?), slot_nudillos = (?), slot_helmet = (?), slot_shield = (?), slot_ammo = (?), "
        query = query & "slot_ship = (?), slot_ring = (?), min_hp = (?), max_hp = (?), min_man = (?), max_man = (?), min_sta = (?), max_sta = (?), min_ham = (?), max_ham = (?), min_sed = (?), max_sed = (?), "
        query = query & "min_hit = (?), max_hit = (?), killed_npcs = (?), killed = (?), rep_asesino = (?), rep_bandido = (?), rep_burgues = (?), rep_ladron = (?), rep_noble = (?), rep_plebe = (?), rep_average = (?), "
        query = query & "is_naked = (?), is_poisoned = (?), is_incinerado = (?), is_hidden = (?), is_hungry = (?), is_thirsty = (?), is_ban = (?), is_dead = (?), is_sailing = (?), is_paralyzed = (?), "
        query = query & "counter_pena = (?), pertenece_consejo_real = (?), pertenece_consejo_caos = (?), pertenece_real = (?), pertenece_caos = (?), ciudadanos_matados = (?), criminales_matados = (?), "
        query = query & "recibio_armadura_real = (?), recibio_armadura_caos = (?), recibio_exp_real = (?), recibio_exp_caos = (?), recompensas_real = (?), recompensas_caos = (?), "
        query = query & "reenlistadas = (?), fecha_ingreso = (?), nivel_ingreso = (?), matados_ingreso = (?), siguiente_recompensa = (?), guild_index = (?), is_global = (?), "
        query = query & "modocombate = (?), seguro = (?), pareja = (?) WHERE id = (?)"

    With UserList(UserIndex)
    
        Call User_Database.MakeQuery(query, True, _
                                    .Name, .Stats.ELV, .Stats.Exp, .Stats.ELU, .Genero, .Raza, .clase, .Hogar, .Desc, .Stats.Gld, .Stats.Banco, .Stats.SkillPts, _
                                    .Counters.AsignedSkills, .NroMascotas, .Pos.Map, .Pos.X, .Pos.Y, .flags.lastMap, .Char.body, .Char.Head, .Char.WeaponAnim, .Char.CascoAnim, .Char.ShieldAnim, _
                                    .Char.AuraAnim, .Char.AuraColor, .Char.Heading, .Invent.NroItems, .Invent.ArmourEqpSlot, .Invent.WeaponEqpSlot, .Invent.AnilloEqpSlot, .Invent.CascoEqpSlot, .Invent.EscudoEqpSlot, .Invent.MunicionEqpSlot, _
                                    .Invent.BarcoSlot, .Invent.AnilloEqpSlot, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMAN, .Stats.MaxMAN, .Stats.MinSta, .Stats.MaxSta, .Stats.MinHam, .Stats.MaxHam, .Stats.MinAGU, .Stats.MaxAGU, _
                                    .Stats.MinHIT, .Stats.MaxHit, .Stats.NPCsMuertos, .Stats.Muertes, .Reputacion.AsesinoRep, .Reputacion.BandidoRep, .Reputacion.BurguesRep, .Reputacion.LadronesRep, .Reputacion.NobleRep, .Reputacion.PlebeRep, .Reputacion.Promedio, _
                                    .flags.Desnudo, .flags.Envenenado, .flags.Incinerado, .flags.Escondido, .flags.Hambre, .flags.Sed, .flags.Ban, .flags.Muerto, .flags.Navegando, .flags.Paralizado, _
                                    .Counters.Pena, (.flags.Privilegios And PlayerType.RoyalCouncil), (.flags.Privilegios And PlayerType.ChaosCouncil), .Faccion.ArmadaReal, .Faccion.FuerzasCaos, .Faccion.CiudadanosMatados, .Faccion.CriminalesMatados, _
                                    .Faccion.RecibioArmaduraReal, .Faccion.RecibioArmaduraCaos, .Faccion.RecibioExpInicialReal, .Faccion.RecibioExpInicialCaos, .Faccion.RecompensasReal, .Faccion.RecompensasCaos, _
                                    .Faccion.Reenlistadas, .Faccion.FechaIngreso, .Faccion.NivelIngreso, .Faccion.MatadosIngreso, .Faccion.NextRecompensa, .GuildIndex, .flags.Global, _
                                    IIf(.flags.ModoCombate = True, "1", "0"), IIf(.flags.Seguro = True, "1", "0"), .flags.miPareja, .Id)
                                        
        '*******************************************************************
        'Familiar
        '*******************************************************************
                                        
        query = "UPDATE familiar SET nombre = (?), level = (?), exp = (?), elu = (?), tipo = (?), min_hp = (?), max_hp = (?), min_hit = (?), max_hit = (?), h_id1 = (?), h_id2 = (?), h_id3 = (?), h_id4 = (?) WHERE user_id = (?)"
                                        
        Call User_Database.MakeQuery(query, True, .Familiar.Nombre, .Familiar.Nivel, .Familiar.Exp, .Familiar.ELU, .Familiar.Tipo, .Familiar.MinHp, .Familiar.MaxHp, .Familiar.MinHIT, .Familiar.MaxHit, _
                                    .Familiar.Spell(0), .Familiar.Spell(1), .Familiar.Spell(2), .Familiar.Spell(3), .Id)
                                        
        '*******************************************************************
        'Hechizos
        '*******************************************************************
        
        query = "UPDATE spell SET "
        
        For LoopC = 1 To MAXUSERHECHIZOS
            
            query = query & "spell_id" & LoopC & " = '" & .Stats.UserHechizos(LoopC) & "' "
            If LoopC < MAXUSERHECHIZOS Then query = query & ", "
            
        Next LoopC
        
        query = query & " WHERE user_id = '" & .Id & "'"

        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Inventario
        '*******************************************************************
        
        query = "UPDATE inventario_items SET "
        
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            
            query = query & "item_id" & LoopC & " = '" & .Invent.Object(LoopC).ObjIndex & "', "
            query = query & "amount" & LoopC & " = '" & .Invent.Object(LoopC).Amount & "', "
            query = query & "is_equipped" & LoopC & " = '" & .Invent.Object(LoopC).Equipped & "'"
            
            If LoopC < MAX_INVENTORY_SLOTS Then query = query & ", "
            
        Next LoopC
        
        query = query & " WHERE user_id = '" & .Id & "'"
        
        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Boveda
        '*******************************************************************
        
        query = "UPDATE banco_items SET "
        
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            
            query = query & "item_id" & LoopC & " = '" & .BancoInvent.Object(LoopC).ObjIndex & "', "
            query = query & "amount" & LoopC & " = '" & .BancoInvent.Object(LoopC).Amount & "'"
            
            If LoopC < MAX_BANCOINVENTORY_SLOTS Then query = query & ", "
            
        Next LoopC
        
        query = query & " WHERE user_id = '" & .Id & "'"
        
        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Skills
        '*******************************************************************
        query = "UPDATE skillpoint SET "
        
        For LoopC = 1 To NUMSKILLS
            
            query = query & "sk" & LoopC & " = '" & .Stats.UserSkills(LoopC) & "', "
            query = query & "exp" & LoopC & " = '" & .Stats.ExpSkills(LoopC) & "', "
            query = query & "elu" & LoopC & " = '" & .Stats.EluSkills(LoopC) & "'"
            If LoopC < NUMSKILLS Then query = query & ", "

        Next LoopC
        
        query = query & " WHERE user_id = '" & .Id & "'"
        
        Call User_Database.Database_Connection.Execute(query)
       
        '*******************************************************************
        'Mascotas
        '*******************************************************************
        Dim petType As Integer
        
        query = "UPDATE pet SET "
        
        For LoopC = 1 To MAXMASCOTAS
            
            'CHOTS | I got this logic from SaveUserToCharfile
            If .MascotasIndex(LoopC) > 0 Then
                If Npclist(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
                    petType = .MascotasType(LoopC)
                Else
                    petType = 0

                End If

            Else
                petType = .MascotasType(LoopC)

            End If

            query = query & "pet" & LoopC & " = '" & petType & "' "
            If LoopC < MAXMASCOTAS Then query = query & ", "
        Next LoopC
        
        query = query & "WHERE user_id = '" & .Id & "'"

        Call User_Database.Database_Connection.Execute(query)
        
        '*******************************************************************
        'Macros
        '*******************************************************************
        query = "UPDATE macros SET "
        
        For LoopC = 1 To NUMMACROS
            
            query = query & "tipoaccion" & LoopC & " = '" & .MacrosKey(LoopC).TipoAccion & "', "
            query = query & "spell" & LoopC & " = '" & .MacrosKey(LoopC).hList & "', "
            query = query & "inv" & LoopC & " = '" & .MacrosKey(LoopC).InvObj & "', "
            query = query & "command" & LoopC & " = '" & .MacrosKey(LoopC).Comando & "'"
            
            If LoopC < NUMMACROS Then query = query & ", "
            
        Next LoopC
        
        query = query & " WHERE user_id = '" & .Id & "'"
        
        Call User_Database.Database_Connection.Execute(query)

    End With

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Unable to UPDATE personaje to Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)

End Sub

Sub LoadUserFromDatabase(ByVal UserIndex As Integer)
    '*************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last modified: 09/10/2018
    'Loads the user from the database
    '*************************************************

    On Error GoTo ErrorHandler

    Dim query As String

    Dim LoopC As Byte

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    'Basic user data
    With UserList(UserIndex)
        query = "SELECT *, DATE_FORMAT(fecha_ingreso, '%Y-%m-%d') as 'fecha_ingreso_format' FROM personaje WHERE UPPER(name) = (?)"

        If Not User_Database.MakeQuery(query, False, UCase$(.Name)) Then Exit Sub

        'Start setting data
        .Id = User_Database.Database_RecordSet!Id
        .Name = User_Database.Database_RecordSet!Name
        .Stats.ELV = User_Database.Database_RecordSet!level
        .Stats.Exp = User_Database.Database_RecordSet!Exp
        .Stats.ELU = User_Database.Database_RecordSet!ELU
        .Genero = User_Database.Database_RecordSet!genre_id
        .Raza = User_Database.Database_RecordSet!race_id
        .clase = User_Database.Database_RecordSet!class_id
        .Hogar = User_Database.Database_RecordSet!home_id
        .Desc = User_Database.Database_RecordSet!description
        .Stats.Gld = User_Database.Database_RecordSet!Gold
        .Stats.Banco = User_Database.Database_RecordSet!bank_gold
        .Stats.SkillPts = User_Database.Database_RecordSet!free_skillpoints
        .Counters.AsignedSkills = User_Database.Database_RecordSet!assigned_skillpoints
        .NroMascotas = User_Database.Database_RecordSet!pet_amount
        .Pos.Map = User_Database.Database_RecordSet!pos_map
        .Pos.X = User_Database.Database_RecordSet!pos_x
        .Pos.Y = User_Database.Database_RecordSet!pos_y
        .flags.lastMap = User_Database.Database_RecordSet!last_map
        .OrigChar.body = User_Database.Database_RecordSet!body_id
        .OrigChar.Head = User_Database.Database_RecordSet!head_id
        .OrigChar.WeaponAnim = User_Database.Database_RecordSet!weapon_id
        .OrigChar.CascoAnim = User_Database.Database_RecordSet!helmet_id
        .OrigChar.ShieldAnim = User_Database.Database_RecordSet!shield_id
        .OrigChar.Heading = User_Database.Database_RecordSet!Heading
        .OrigChar.AuraAnim = User_Database.Database_RecordSet!Aura_id
        .OrigChar.AuraColor = User_Database.Database_RecordSet!Aura_color
        .Invent.NroItems = User_Database.Database_RecordSet!items_amount
        .Invent.ArmourEqpSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_armour, 0)
        .Invent.WeaponEqpSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_weapon, 0)
        .Invent.NudiEqpSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_nudillos, 0)
        .Invent.CascoEqpSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_helmet, 0)
        .Invent.EscudoEqpSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_shield, 0)
        .Invent.MunicionEqpSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_ammo, 0)
        .Invent.BarcoSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_ship, 0)
        .Invent.AnilloEqpSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_ring, 0)
        .Stats.MinHp = User_Database.Database_RecordSet!min_hp
        .Stats.MaxHp = User_Database.Database_RecordSet!max_hp
        .Stats.MinMAN = User_Database.Database_RecordSet!min_man
        .Stats.MaxMAN = User_Database.Database_RecordSet!max_man
        .Stats.MinSta = User_Database.Database_RecordSet!min_sta
        .Stats.MaxSta = User_Database.Database_RecordSet!max_sta
        .Stats.MinHam = User_Database.Database_RecordSet!min_ham
        .Stats.MaxHam = User_Database.Database_RecordSet!max_ham
        .Stats.MinAGU = User_Database.Database_RecordSet!min_sed
        .Stats.MaxAGU = User_Database.Database_RecordSet!max_sed
        .Stats.MinHIT = User_Database.Database_RecordSet!min_hit
        .Stats.MaxHit = User_Database.Database_RecordSet!max_hit
        .Stats.NPCsMuertos = User_Database.Database_RecordSet!killed_npcs
        .Stats.Muertes = User_Database.Database_RecordSet!killed
        .Reputacion.AsesinoRep = User_Database.Database_RecordSet!rep_asesino
        .Reputacion.BandidoRep = User_Database.Database_RecordSet!rep_bandido
        .Reputacion.BurguesRep = User_Database.Database_RecordSet!rep_burgues
        .Reputacion.LadronesRep = User_Database.Database_RecordSet!rep_ladron
        .Reputacion.NobleRep = User_Database.Database_RecordSet!rep_noble
        .Reputacion.PlebeRep = User_Database.Database_RecordSet!rep_plebe
        .Reputacion.Promedio = User_Database.Database_RecordSet!rep_average
        .flags.Desnudo = User_Database.Database_RecordSet!is_naked
        .flags.Envenenado = User_Database.Database_RecordSet!is_poisoned
        .flags.Incinerado = User_Database.Database_RecordSet!is_incinerado
        .flags.Escondido = User_Database.Database_RecordSet!is_hidden
        .flags.Hambre = User_Database.Database_RecordSet!is_hungry
        .flags.Sed = User_Database.Database_RecordSet!is_thirsty
        .flags.Ban = User_Database.Database_RecordSet!is_ban
        .flags.Muerto = User_Database.Database_RecordSet!is_dead
        .flags.Navegando = User_Database.Database_RecordSet!is_sailing
        .flags.Paralizado = User_Database.Database_RecordSet!is_paralyzed
        .Counters.Pena = User_Database.Database_RecordSet!counter_pena
        .flags.Global = User_Database.Database_RecordSet!is_global
        .flags.ModoCombate = User_Database.Database_RecordSet!ModoCombate
        .flags.Seguro = User_Database.Database_RecordSet!Seguro
        .flags.miPareja = User_Database.Database_RecordSet!pareja
    
        'Add Cr3p-1 agregamos esto para que funcione los casamientos. (xD)
        If Len(.flags.miPareja) > 0 Then
            .flags.toyCasado = 1
        End If
        '\Add
    
        If User_Database.Database_RecordSet!pertenece_consejo_real Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.RoyalCouncil

        End If

        If User_Database.Database_RecordSet!pertenece_consejo_caos Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.ChaosCouncil

        End If

        .Faccion.ArmadaReal = User_Database.Database_RecordSet!pertenece_real
        .Faccion.FuerzasCaos = User_Database.Database_RecordSet!pertenece_caos
        .Faccion.CiudadanosMatados = User_Database.Database_RecordSet!ciudadanos_matados
        .Faccion.CriminalesMatados = User_Database.Database_RecordSet!criminales_matados
        .Faccion.RecibioArmaduraReal = User_Database.Database_RecordSet!recibio_armadura_real
        .Faccion.RecibioArmaduraCaos = User_Database.Database_RecordSet!recibio_armadura_caos
        .Faccion.RecibioExpInicialReal = User_Database.Database_RecordSet!recibio_exp_real
        .Faccion.RecibioExpInicialCaos = User_Database.Database_RecordSet!recibio_exp_caos
        .Faccion.RecompensasReal = User_Database.Database_RecordSet!recompensas_real
        .Faccion.RecompensasCaos = User_Database.Database_RecordSet!recompensas_caos
        .Faccion.Reenlistadas = User_Database.Database_RecordSet!Reenlistadas
        .Faccion.FechaIngreso = SanitizeNullValue(User_Database.Database_RecordSet!fecha_ingreso_format, vbNullString)
        .Faccion.NivelIngreso = SanitizeNullValue(User_Database.Database_RecordSet!nivel_ingreso, 0)
        .Faccion.MatadosIngreso = SanitizeNullValue(User_Database.Database_RecordSet!matados_ingreso, 0)
        .Faccion.NextRecompensa = SanitizeNullValue(User_Database.Database_RecordSet!siguiente_recompensa, 0)

        .GuildIndex = SanitizeNullValue(User_Database.Database_RecordSet!Guild_Index, 0)

        Set User_Database.Database_RecordSet = Nothing
        
        '*******************************************************************
        'Familiar
        '*******************************************************************
        
        query = "SELECT * FROM familiar WHERE user_id = " & .Id & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)
        
        .Familiar.Nombre = User_Database.Database_RecordSet!Nombre
        .Familiar.Nivel = User_Database.Database_RecordSet!level
        .Familiar.Exp = User_Database.Database_RecordSet!Exp
        .Familiar.ELU = User_Database.Database_RecordSet!ELU
        .Familiar.Tipo = User_Database.Database_RecordSet!Tipo
        .Familiar.MinHp = User_Database.Database_RecordSet!min_hp
        .Familiar.MaxHp = User_Database.Database_RecordSet!max_hp
        .Familiar.MinHIT = User_Database.Database_RecordSet!min_hit
        .Familiar.MaxHit = User_Database.Database_RecordSet!max_hit
        .Familiar.Spell(0) = User_Database.Database_RecordSet!h_id1
        .Familiar.Spell(1) = User_Database.Database_RecordSet!h_id2
        .Familiar.Spell(2) = User_Database.Database_RecordSet!h_id3
        .Familiar.Spell(3) = User_Database.Database_RecordSet!h_id4
        
        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Atributos
        '*******************************************************************
        query = "SELECT * FROM atributos WHERE user_id = " & .Id & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)
    
        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            
            User_Database.Database_RecordSet.MoveFirst
            
            For LoopC = 1 To NUMATRIBUTOS

                .Stats.UserAtributos(LoopC) = User_Database.Database_RecordSet("att" & LoopC)
                .Stats.UserAtributosBackUP(LoopC) = .Stats.UserAtributos(LoopC)

            Next LoopC

        End If

        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Hechizos
        '*******************************************************************
        query = "SELECT * FROM spell WHERE user_id = " & .Id & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst

            For LoopC = 1 To MAXUSERHECHIZOS
                .Stats.UserHechizos(LoopC) = User_Database.Database_RecordSet("spell_id" & LoopC)
            Next LoopC

        End If

        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Mascotas
        '*******************************************************************
        query = "SELECT * FROM pet WHERE user_id = " & .Id & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst

            For LoopC = 1 To MAXMASCOTAS
                .MascotasType(LoopC) = User_Database.Database_RecordSet("pet" & LoopC)
            Next LoopC

        End If

        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Inventario
        '*******************************************************************
        query = "SELECT * FROM inventario_items WHERE user_id = " & .Id & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst
                
            For LoopC = 1 To MAX_INVENTORY_SLOTS
                .Invent.Object(LoopC).ObjIndex = User_Database.Database_RecordSet("item_id" & LoopC)
                .Invent.Object(LoopC).Amount = User_Database.Database_RecordSet("Amount" & LoopC)
                .Invent.Object(LoopC).Equipped = User_Database.Database_RecordSet("is_equipped" & LoopC)
            Next LoopC
                
        End If

        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Boveda
        '*******************************************************************
        query = "SELECT * FROM banco_items WHERE user_id = " & .Id & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst
                
            For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
                .BancoInvent.Object(LoopC).ObjIndex = User_Database.Database_RecordSet("item_id" & LoopC)
                .BancoInvent.Object(LoopC).Amount = User_Database.Database_RecordSet("Amount" & LoopC)
            Next LoopC
                
        End If

        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Skills
        '*******************************************************************
        query = "SELECT * FROM skillpoint WHERE user_id = " & .Id & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst

            For LoopC = 1 To NUMSKILLS
                .Stats.UserSkills(LoopC) = User_Database.Database_RecordSet("sk" & LoopC)
                .Stats.ExpSkills(LoopC) = User_Database.Database_RecordSet("exp" & LoopC)
                .Stats.EluSkills(LoopC) = User_Database.Database_RecordSet("elu" & LoopC)
                If .Stats.EluSkills(LoopC) < 1 Then .Stats.EluSkills(LoopC) = 200
            Next LoopC

        End If

        Set User_Database.Database_RecordSet = Nothing
        
        '*******************************************************************
        'Macros
        '*******************************************************************
        query = "SELECT * FROM macros WHERE user_id = " & .Id & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst

            For LoopC = 1 To NUMMACROS
                .MacrosKey(LoopC).TipoAccion = User_Database.Database_RecordSet("tipoaccion" & LoopC)
                .MacrosKey(LoopC).hList = User_Database.Database_RecordSet("spell" & LoopC)
                .MacrosKey(LoopC).InvObj = User_Database.Database_RecordSet("inv" & LoopC)
                .MacrosKey(LoopC).Comando = User_Database.Database_RecordSet("command" & LoopC)
            Next LoopC

        End If

        Set User_Database.Database_RecordSet = Nothing

    End With

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Unable to LOAD User from Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function PersonajeExisteDatabase(ByVal UserName As String) As Boolean
    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT id FROM personaje WHERE UPPER(name) = (?) AND deleted = FALSE;"

    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        PersonajeExisteDatabase = False
        Exit Function

    End If

    PersonajeExisteDatabase = (User_Database.Database_RecordSet.RecordCount > 0)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in PersonajeExisteDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function BANCheckDatabase(ByVal UserName As String) As Boolean

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT is_ban FROM personaje WHERE UPPER(name) = (?)"

    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        BANCheckDatabase = False
        Exit Function

    End If

    BANCheckDatabase = CBool(User_Database.Database_RecordSet!is_ban)

    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in BANCheckDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub UnBanDatabase(ByVal UserName As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler
    
    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    Call User_Database.MakeQuery("UPDATE personaje SET is_ban = FALSE WHERE UPPER(name) = (?)", True, UCase$(UserName))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in UnBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserGuildIndexDatabase(ByVal UserName As String) As Integer

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT guild_index FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        GetUserGuildIndexDatabase = 0
        Exit Function

    End If

    GetUserGuildIndexDatabase = SanitizeNullValue(User_Database.Database_RecordSet!Guild_Index, 0)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub CopyUserDatabase(ByVal UserName As String, ByVal newName As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    Call User_Database.MakeQuery("UPDATE personaje SET name = (?) WHERE UPPER(name) = (?)", True, newName, UCase$(UserName))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in CopyUserDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub MarcarPjComoQueYaVotoDatabase(ByVal UserIndex As Integer, _
                                         ByVal NumeroEncuesta As Integer)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    Call User_Database.MakeQuery("UPDATE personaje SET votes_amount = (?) WHERE id = (?)", True, NumeroEncuesta, UserList(UserIndex).Id)
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in MarcarPjComoQueYaVotoDatabase: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function PersonajeCantidadVotosDatabase(ByVal UserName As String) As Integer

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT votes_amount FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        PersonajeCantidadVotosDatabase = 0
        Exit Function

    End If

    PersonajeCantidadVotosDatabase = CInt(User_Database.Database_RecordSet!votes_amount)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in PersonajeCantidadVotosDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveBan(ByVal UserName As String, _
                           ByVal Reason As String, _
                           ByVal BannedBy As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query     As String

    Dim cantPenas As Byte

    cantPenas = GetUserAmountOfPunishments(UserName)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    Call User_Database.MakeQuery("UPDATE personaje SET is_ban = TRUE WHERE UPPER(name) = (?)", True, UCase$(UserName))

    query = "INSERT INTO punishment SET user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?)), number = (?), reason = (?)"
    Call User_Database.MakeQuery(query, True, UCase$(UserName), (cantPenas + 1), BannedBy & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in SaveBan: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserAmountOfPunishments(ByVal UserName As String) As Integer

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT COUNT(1) as punishments FROM punishment WHERE user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?))"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        GetUserAmountOfPunishments = 0
        Exit Function

    End If

    GetUserAmountOfPunishments = CInt(User_Database.Database_RecordSet!punishments)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserAmountOfPunishments: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SendUserPunishments(ByVal UserIndex As Integer, _
                                       ByVal UserName As String, _
                                       ByVal Count As Integer)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT * FROM punishment WHERE user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?))"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        User_Database.Database_RecordSet.MoveFirst

        While Not User_Database.Database_RecordSet.EOF

            Call WriteConsoleMsg(UserIndex, User_Database.Database_RecordSet!Number & " - " & User_Database.Database_RecordSet!Reason, FontTypeNames.FONTTYPE_INFO)

            User_Database.Database_RecordSet.MoveNext
        Wend

    End If

    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendUserPunishments: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserPos(ByVal UserName As String) As String

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT pos_map, pos_x, pos_y FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        GetUserPos = vbNullString
        Exit Function

    End If

    GetUserPos = User_Database.Database_RecordSet!pos_map & "-" & User_Database.Database_RecordSet!pos_x & "-" & User_Database.Database_RecordSet!pos_y
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserPos: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveUserPunishment(ByVal UserName As String, _
                                      ByVal Number As Integer, _
                                      ByVal Reason As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "INSERT INTO punishment SET user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?)), number = (?), reason = (?)"
    Call User_Database.MakeQuery(query, True, UCase$(UserName), Number, Reason)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserPunishment: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub AlterUserPunishment(ByVal UserName As String, _
                                       ByVal Number As Integer, _
                                       ByVal Reason As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE punishment SET reason = (?) WHERE number = (?) AND user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, Reason, Number, UCase$(UserName))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in AlterUserPunishment: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub ResetUserFacciones(ByVal UserName As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET pertenece_real = FALSE, pertenece_caos = FALSE, ciudadanos_matados = 0, criminales_matados = FALSE, "
    query = query & "recibio_armadura_real = FALSE, recibio_armadura_caos = FALSE, recibio_exp_real = FALSE, recibio_exp_caos = FALSE, "
    query = query & "recompensas_real = 0, recompensas_caos = 0, reenlistadas = 0, fecha_ingreso = NULL, nivel_ingreso = NULL, "
    query = query & "matados_ingreso = NULL, siguiente_recompensa = NULL WHERE UPPER(name) = (?)"

    Call User_Database.MakeQuery(query, True, UCase$(UserName))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in ResetUserFacciones: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub KickUserCouncils(ByVal UserName As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET pertenece_consejo_real = FALSE, pertenece_consejo_caos = FALSE WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, UCase$(UserName))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in KickUserCouncils: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub KickUserFacciones(ByVal UserName As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET pertenece_real = FALSE, pertenece_caos = FALSE WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, UCase$(UserName))
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in KickUserFacciones: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub KickUserChaosLegion(ByVal UserName As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET pertenece_caos = FALSE, reenlistadas = 200 WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, UCase$(UserName))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in KickUserChaosLegion: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub KickUserRoyalArmy(ByVal UserName As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET pertenece_real = FALSE, reenlistadas = 200 WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, UCase$(UserName))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in KickUserRoyalArmy: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub UpdateUserLogged(ByVal UserName As String, ByVal Logged As Byte)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET is_logged = " & IIf(Logged = 1, "TRUE", "FALSE") & " WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, UCase$(UserName))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in UpdateUserLogged: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserLastIps(ByVal UserName As String) As String

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT last_ip FROM cuentas WHERE id = (SELECT cuenta_id FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        GetUserLastIps = vbNullString
        Exit Function

    End If

    GetUserLastIps = User_Database.Database_RecordSet!last_ip
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserLastIps: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserSkills(ByVal UserName As String) As String

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    GetUserSkills = vbNullString

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT number, value FROM skillpoint WHERE user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?))"
    
   If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        User_Database.Database_RecordSet.MoveFirst

        While Not User_Database.Database_RecordSet.EOF

            GetUserSkills = GetUserSkills & "CHAR>" & SkillsNames(User_Database.Database_RecordSet!Number) & " = " & User_Database.Database_RecordSet!Value & vbCrLf

            User_Database.Database_RecordSet.MoveNext
        Wend

    End If

    Set User_Database.Database_RecordSet = Nothing

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserSkills: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserFreeSkills(ByVal UserName As String) As Integer

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT free_skillpoints FROM personaje WHERE UPPER(name) = (?)"
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        GetUserFreeSkills = 0
        Exit Function

    End If

    GetUserFreeSkills = CInt(User_Database.Database_RecordSet!free_skillpoints)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserFreeSkills: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveUserTrainingTime(ByVal UserName As String, _
                                        ByVal trainingTime As Long)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET counter_training = (?) WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, trainingTime, UCase$(UserName))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserTrainingTime: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserTrainingTime(ByVal UserName As String) As Long

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT counter_training FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        GetUserTrainingTime = 0
        Exit Function

    End If

    GetUserTrainingTime = CLng(User_Database.Database_RecordSet!counter_training)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserTrainingTime: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function UserBelongsToRoyalArmy(ByVal UserName As String) As Boolean

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT pertenece_real FROM personaje WHERE UPPER(name) = (?) AND deleted = FALSE;"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        UserBelongsToRoyalArmy = False
        Exit Function

    End If

    UserBelongsToRoyalArmy = CBool(User_Database.Database_RecordSet!pertenece_real)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in UserBelongsToRoyalArmy: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function UserBelongsToChaosLegion(ByVal UserName As String) As Boolean

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT pertenece_caos FROM personaje WHERE UPPER(name) = (?) AND deleted = FALSE;"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        UserBelongsToChaosLegion = False
        Exit Function

    End If

    UserBelongsToChaosLegion = CBool(User_Database.Database_RecordSet!pertenece_caos)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in UserBelongsToChaosLegion: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserLevel(ByVal UserName As String) As Byte

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT level FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        GetUserLevel = 0
        Exit Function

    End If

    GetUserLevel = CByte(User_Database.Database_RecordSet!level)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserLevel: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserPromedio(ByVal UserName As String) As Long

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT rep_average FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        GetUserPromedio = 0
        Exit Function

    End If

    GetUserPromedio = CLng(User_Database.Database_RecordSet!rep_average)
    Set User_Database.Database_RecordSet = Nothing

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserPromedio: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserReenlists(ByVal UserName As String) As Byte

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT reenlistadas FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        GetUserReenlists = 0
        Exit Function

    End If

    GetUserReenlists = CByte(User_Database.Database_RecordSet!Reenlistadas)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserReenlists: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveUserReenlists(ByVal UserName As String, ByVal Reenlists As Byte)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET reenlistadas = (?) WHERE UPPER(name) = (?)"

    Call User_Database.MakeQuery(query, True, Reenlists, UCase$(UserName))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserReenlists: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendUserStatsTxtDatabase(ByVal sendIndex As Integer, ByVal UserName As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    If Not PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & UserName, FontTypeNames.FONTTYPE_INFO)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If
    
        query = "SELECT level, exp, elu, min_sta, max_sta, min_hp, max_hp, min_man, max_man, min_hit, max_hit, gold FROM personaje WHERE UPPER(name) = (?)"
        
        If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Call WriteConsoleMsg(sendIndex, "Nivel: " & User_Database.Database_RecordSet!level & "  EXP: " & User_Database.Database_RecordSet!Exp & "/" & User_Database.Database_RecordSet!ELU, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Energia: " & User_Database.Database_RecordSet!min_sta & "/" & User_Database.Database_RecordSet!max_sta, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & User_Database.Database_RecordSet!min_hp & "/" & User_Database.Database_RecordSet!max_hp, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Mana: " & User_Database.Database_RecordSet!min_man & "/" & User_Database.Database_RecordSet!max_man, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Golpe: " & User_Database.Database_RecordSet!min_hit & "/" & User_Database.Database_RecordSet!max_hit, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Oro: " & User_Database.Database_RecordSet!Gold, FontTypeNames.FONTTYPE_INFO)

        Set User_Database.Database_RecordSet = Nothing
        
        #If DBConexionUnica = 0 Then
            Call User_Database.Database_Close
        #End If

    End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendUserStatsTxtDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendUserMiniStatsTxtFromDatabase(ByVal sendIndex As Integer, _
                                            ByVal UserName As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    If Not PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & UserName, FontTypeNames.FONTTYPE_INFO)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If
    
        query = "SELECT killed_npcs, killed_users, ciudadanos_matados, criminales_matados, class_id, genre_id, race_id FROM personaje WHERE UPPER(name) = (?)"
        
        If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Call WriteConsoleMsg(sendIndex, "Pj: " & UserName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "CiudadanosMatados: " & User_Database.Database_RecordSet!ciudadanos_matados & ", CriminalesMatados: " & User_Database.Database_RecordSet!criminales_matados & ", UsuariosMatados: " & User_Database.Database_RecordSet!killed_users, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCs muertos: " & User_Database.Database_RecordSet!killed_npcs, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(User_Database.Database_RecordSet!class_id), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Genero: " & IIf(CByte(User_Database.Database_RecordSet!ciudadanos_matados) = eGenero.Hombre, "Hombre", "Mujer"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Raza: " & ListaRazas(User_Database.Database_RecordSet!race_id), FontTypeNames.FONTTYPE_INFO)

        Set User_Database.Database_RecordSet = Nothing
        
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendUserMiniStatsTxtFromDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendUserOROTxtFromDatabase(ByVal sendIndex As Integer, _
                                      ByVal UserName As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    If Not PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        #If DBConexionUnica = 0 Then
            Call User_Database.Database_Connect
        #Else
            'Si perdimos la conexion reconectamos
            If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
        #End If

        query = "SELECT bank_gold FROM personaje WHERE UPPER(name) = (?)"
        
        If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Call WriteConsoleMsg(sendIndex, "Pj: " & UserName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Oro en banco: " & User_Database.Database_RecordSet!bank_gold, FontTypeNames.FONTTYPE_INFO)

        Set User_Database.Database_RecordSet = Nothing
        
        #If DBConexionUnica = 0 Then
            Call User_Database.Database_Close
        #End If

    End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendUserOROTxtFromDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendUserInvTxtFromDatabase(ByVal sendIndex As Integer, _
                                      ByVal UserName As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query   As String
    Dim LoopC   As Byte

    Dim ObjInd  As Long

    If Not PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        #If DBConexionUnica = 0 Then
            Call User_Database.Database_Connect
        #Else
            'Si perdimos la conexion reconectamos
            If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
        #End If

        query = "SELECT "

        For LoopC = 1 To MAX_INVENTORY_SLOTS
            query = query & "item_id" & LoopC & ", amount" & LoopC
            If LoopC < MAX_INVENTORY_SLOTS Then query = query & ", "
        Next LoopC

        query = query & " FROM inventario_items WHERE user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?))"
        
        If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
            User_Database.Database_RecordSet.MoveFirst

            While Not User_Database.Database_RecordSet.EOF

                ObjInd = val(User_Database.Database_RecordSet!item_id)

                If ObjInd > 0 Then
                    Call WriteConsoleMsg(sendIndex, "Objeto " & User_Database.Database_RecordSet!Number & " " & ObjData(ObjInd).Name & " Cantidad:" & User_Database.Database_RecordSet!Amount, FontTypeNames.FONTTYPE_INFO)

                End If

                User_Database.Database_RecordSet.MoveNext
            Wend
        Else
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)

        End If

        Set User_Database.Database_RecordSet = Nothing
        
        #If DBConexionUnica = 0 Then
            Call User_Database.Database_Close
        #End If

    End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendUserInvTxtFromDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendUserBovedaTxtFromDatabase(ByVal sendIndex As Integer, _
                                         ByVal UserName As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query   As String
    Dim LoopC   As Byte

    Dim ObjInd As Long

    If Not PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        #If DBConexionUnica = 0 Then
            Call User_Database.Database_Connect
        #Else
            'Si perdimos la conexion reconectamos
            If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
        #End If
        
        query = "SELECT "

        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            query = query & "item_id" & LoopC & ", amount" & LoopC
            If LoopC < MAX_BANCOINVENTORY_SLOTS Then query = query & ", "
        Next LoopC
        
        query = query & " FROM banco_items WHERE user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?))"
        
        If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
            User_Database.Database_RecordSet.MoveFirst

            While Not User_Database.Database_RecordSet.EOF

                ObjInd = val(User_Database.Database_RecordSet!item_id)

                If ObjInd > 0 Then
                    Call WriteConsoleMsg(sendIndex, "Objeto " & User_Database.Database_RecordSet!Number & " " & ObjData(ObjInd).Name & " Cantidad:" & User_Database.Database_RecordSet!Amount, FontTypeNames.FONTTYPE_INFO)

                End If

                User_Database.Database_RecordSet.MoveNext
            Wend
        Else
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)

        End If

        Set User_Database.Database_RecordSet = Nothing
        
        #If DBConexionUnica = 0 Then
            Call User_Database.Database_Close
        #End If

    End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendUserBovedaTxtFromDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendCharacterInfoDatabase(ByVal UserIndex As Integer, ByVal UserName As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim gName       As String

    Dim Miembro     As String

    Dim GuildActual As Integer

    Dim query       As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT race_id, class_id, genre_id, level, gold, bank_gold, rep_average, guild_requests_history, guild_index, guild_member_history, pertenece_real, pertenece_caos, ciudadanos_matados, criminales_matados FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        Call WriteConsoleMsg(UserIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    ' Get the character's current guild
    GuildActual = SanitizeNullValue(User_Database.Database_RecordSet!Guild_Index, 0)

    If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
        gName = "<" & GuildName(GuildActual) & ">"
    Else
        gName = "Ninguno"

    End If

    'Get previous guilds
    Miembro = SanitizeNullValue(User_Database.Database_RecordSet!guild_member_history, vbNullString)

    If Len(Miembro) > 400 Then
        Miembro = ".." & Right$(Miembro, 400)

    End If

    Call Protocol.WriteCharacterInfo(UserIndex, UserName, User_Database.Database_RecordSet!race_id, User_Database.Database_RecordSet!class_id, User_Database.Database_RecordSet!genre_id, User_Database.Database_RecordSet!level, User_Database.Database_RecordSet!Gold, User_Database.Database_RecordSet!bank_gold, User_Database.Database_RecordSet!rep_average, SanitizeNullValue(User_Database.Database_RecordSet!guild_requests_history, vbNullString), gName, Miembro, User_Database.Database_RecordSet!pertenece_real, User_Database.Database_RecordSet!pertenece_caos, User_Database.Database_RecordSet!ciudadanos_matados, User_Database.Database_RecordSet!criminales_matados)

#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
#End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendCharacterInfoDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserGuildMemberDatabase(ByVal UserName As String) As String

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT guild_member_history FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        GetUserGuildMemberDatabase = vbNullString
        Exit Function

    End If

    GetUserGuildMemberDatabase = SanitizeNullValue(User_Database.Database_RecordSet!guild_member_history, vbNullString)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildAspirantDatabase(ByVal UserName As String) As Integer

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT guild_aspirant_index FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        GetUserGuildAspirantDatabase = 0
        Exit Function

    End If

    GetUserGuildAspirantDatabase = SanitizeNullValue(User_Database.Database_RecordSet!guild_aspirant_index, 0)
    Set User_Database.Database_RecordSet = Nothing
    
#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
#End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildRejectionReasonDatabase(ByVal UserName As String) As String

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT guild_rejected_because FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        GetUserGuildRejectionReasonDatabase = vbNullString
        Exit Function

    End If

    GetUserGuildRejectionReasonDatabase = SanitizeNullValue(User_Database.Database_RecordSet!guild_rejected_because, vbNullString)
    Set User_Database.Database_RecordSet = Nothing
    
#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
#End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildPedidosDatabase(ByVal UserName As String) As String

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT guild_requests_history FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(UserName)) Then
        GetUserGuildPedidosDatabase = vbNullString
        Exit Function

    End If

    GetUserGuildPedidosDatabase = SanitizeNullValue(User_Database.Database_RecordSet!guild_requests_history, vbNullString)
    Set User_Database.Database_RecordSet = Nothing
    
#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
#End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveUserGuildRejectionReasonDatabase(ByVal UserName As String, _
                                                ByVal Reason As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET guild_rejected_because = (?) WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, Reason, UCase$(UserName))
    
#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
#End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildIndexDatabase(ByVal UserName As String, _
                                      ByVal GuildIndex As Integer)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET guild_index = (?) WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, GuildIndex, UCase$(UserName))
    
#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
#End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildAspirantDatabase(ByVal UserName As String, _
                                         ByVal AspirantIndex As Integer)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET guild_aspirant_index = (?) WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, AspirantIndex, UCase$(UserName))

#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
#End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildMemberDatabase(ByVal UserName As String, ByVal guilds As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET guild_member_history = (?) WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, guilds, UCase$(UserName))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildPedidosDatabase(ByVal UserName As String, ByVal Pedidos As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET guild_requests_history = (?) WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, Pedidos, UCase$(UserName))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function SanitizeNullValue(ByVal Value As Variant, _
                                  ByVal defaultValue As Variant) As Variant
    SanitizeNullValue = IIf(IsNull(Value), defaultValue, Value)

End Function

Public Sub SaveOroBanco(ByVal UserName As String, ByVal Oros As Long)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 31/03/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET bank_gold = bank_gold + '" & Oros & "' WHERE id = " & GetAccountID(UserName)

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveOroBanco: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub
