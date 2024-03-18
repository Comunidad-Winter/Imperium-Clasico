Attribute VB_Name = "NPCs"
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

'????????????????????????????
'????????????????????????????
'????????????????????????????
'                        Modulo NPC
'????????????????????????????
'????????????????????????????
'????????????????????????????
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'????????????????????????????
'????????????????????????????
'????????????????????????????

Option Explicit
#If False Then

    Dim X, Y, n, Map As Variant

#End If

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer
    
    For i = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasIndex(i) = NPCIndex Then
            UserList(UserIndex).MascotasIndex(i) = 0
            UserList(UserIndex).MascotasType(i) = 0
         
            UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1
            Exit For

        End If

    Next i

End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1

End Sub

Public Sub MuereNpc(ByVal NPCIndex As Integer, ByVal UserIndex As Integer)

    '********************************************************
    'Author: Unknown
    'Llamado cuando la vida de un NPC llega a cero.
    'Last Modify Date: 13/07/2010
    '24/01/2007: Pablo (ToxicWaste): Agrego para actualizacion de tag si cambia de status.
    '22/05/2010: ZaMa - Los caos ya no suben nobleza ni plebe al atacar npcs.
    '23/05/2010: ZaMa - El usuario pierde la pertenencia del npc.
    '********************************************************
    On Error GoTo errHandler

    Dim MiNPC As NPC

    MiNPC = Npclist(NPCIndex)

    Dim EraCriminal     As Boolean
   
    'Respawn de NPC con retardo
    If MiNPC.flags.TiempoRetardoMin > 0 Then
        RetardoSpawn(MiNPC.Numero).Tiempo = RandomNumber(MiNPC.flags.TiempoRetardoMin, MiNPC.flags.TiempoRetardoMax)
        RetardoSpawn(MiNPC.Numero).Mapa = MiNPC.Orig.Map
        RetardoSpawn(MiNPC.Numero).X = MiNPC.Orig.X
        RetardoSpawn(MiNPC.Numero).Y = MiNPC.Orig.Y
        RetardoSpawn(MiNPC.Numero).NPCNUM = MiNPC.Numero
    End If
  '/Respawn de NPC con retardo
      
    If MiNPC.EsFamiliar = 1 Then
        Call Familiar_Muerte(MiNPC.EsFamiliar, True)
        Call WriteConsoleMsg(UserIndex, "�Has matado el familiar de " & UserList(MiNPC.MaestroUser).Name & "!", FontTypeNames.FONTTYPE_FIGHT)
        
    End If
      
    'Quitamos el npc
    Call QuitarNPC(NPCIndex) '
    
    If UserIndex > 0 Then ' Lo mato un usuario?

        With UserList(UserIndex)
        
            If MiNPC.flags.Snd3 > 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, MiNPC.Pos.X, MiNPC.Pos.Y))

            End If

            .flags.TargetNPC = 0
            .flags.TargetNpcTipo = eNPCType.Comun
            
            'El user que lo mato tiene mascotas?
            If .NroMascotas > 0 Then

                Dim t As Integer

                For t = 1 To MAXMASCOTAS

                    If .MascotasIndex(t) > 0 Then
                        If Npclist(.MascotasIndex(t)).TargetNPC = NPCIndex Then
                            Call FollowAmo(.MascotasIndex(t))

                        End If

                    End If

                Next t

            End If
            
            Call WriteConsoleMsg(UserIndex, "Has matado a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
            
            If .Stats.NPCsMuertos < 32000 Then .Stats.NPCsMuertos = .Stats.NPCsMuertos + 1
            
            EraCriminal = criminal(UserIndex)
            
            If MiNPC.Stats.Alineacion = 0 Then
            
                If MiNPC.Numero = Guardias Then
                    .Reputacion.NobleRep = 0
                    .Reputacion.PlebeRep = 0
                    .Reputacion.AsesinoRep = .Reputacion.AsesinoRep + 500

                    If .Reputacion.AsesinoRep > MAXREP Then .Reputacion.AsesinoRep = MAXREP

                End If
                
                If MiNPC.MaestroUser = 0 Then
                    .Reputacion.AsesinoRep = .Reputacion.AsesinoRep + vlASESINO

                    If .Reputacion.AsesinoRep > MAXREP Then .Reputacion.AsesinoRep = MAXREP

                End If
                
            ElseIf Not esCaos(UserIndex) Then

                If MiNPC.Stats.Alineacion = 1 Then
                    .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlCAZADOR

                    If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
                        
                ElseIf MiNPC.Stats.Alineacion = 2 Then
                    .Reputacion.NobleRep = .Reputacion.NobleRep + vlASESINO / 2

                    If .Reputacion.NobleRep > MAXREP Then .Reputacion.NobleRep = MAXREP
                        
                ElseIf MiNPC.Stats.Alineacion = 4 Then
                    .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlCAZADOR

                    If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
                        
                End If

            End If
            
            Dim EsCriminal As Boolean

            EsCriminal = criminal(UserIndex)
            
            ' Cambio de alienacion?
            If EraCriminal <> EsCriminal Then
                
                ' Se volvio pk?
                If EsCriminal Then
                    If esArmada(UserIndex) Then Call ExpulsarFaccionReal(UserIndex)
                
                    ' Se volvio ciuda
                Else

                    If esCaos(UserIndex) Then Call ExpulsarFaccionCaos(UserIndex)

                End If
                
                Call RefreshCharStatus(UserIndex)

            End If
                        
            Call CheckUserLevel(UserIndex)
            
            If .Familiar.Invocado = 1 Then
                Call CheckFamilyLevel(UserIndex)
            End If
            
            If NPCIndex = .flags.ParalizedByNpcIndex Then
                Call RemoveParalisis(UserIndex)

            End If
            
        End With
                        
    End If ' Userindex > 0
   
    If MiNPC.MaestroUser = 0 Then
        'Tiramos el inventario
        Call NPC_TIRAR_ITEMS(UserIndex, MiNPC)
        'ReSpawn o no
        If MiNPC.flags.TiempoRetardoMin = 0 Then Call ReSpawnNpc(MiNPC)

    End If
                        
    If UserIndex < 1 Then
        UserIndex = MiNPC.MaestroUser

        If UserIndex = 0 Then Exit Sub

    End If

    Exit Sub

errHandler:
    Call LogError("Error en MuereNpc - Error: " & Err.Number & " - Desc: " & Err.description)

End Sub

Private Sub ResetNpcFlags(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'Clear the npc's flags
    
    With Npclist(NPCIndex).flags
        .AfectaParalisis = 0
        .AguaValida = 0
        .AttackedBy = vbNullString
        .AttackedFirstBy = vbNullString
        .BackUp = 0
        .Bendicion = 0
        .Domable = 0
        .Envenenado = 0
        .Incinerado = 0
        .Faccion = 0
        .Follow = False
        .AtacaDoble = 0
        .LanzaSpells = 0
        .invisible = 0
        .Maldicion = 0
        .SiguiendoGm = False
        .OldHostil = 0
        .OldMovement = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Respawn = 0
        .RespawnOrigPos = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .TierraInvalida = 0

    End With

End Sub

Private Sub ResetNpcCounters(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With Npclist(NPCIndex).Contadores
        .Paralisis = 0
        .TiempoExistencia = 0
        .Ataque = 0

    End With

End Sub

Private Sub ResetNpcCharInfo(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With Npclist(NPCIndex).Char
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Particle = 0
        .Head = 0
        .Heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0

    End With

End Sub

Private Sub ResetNpcCriatures(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim j As Long
    
    With Npclist(NPCIndex)

        For j = 1 To .NroCriaturas
            .Criaturas(j).NPCIndex = 0
            .Criaturas(j).NpcName = vbNullString
        Next j
        
        .NroCriaturas = 0

    End With

End Sub

Sub ResetExpresiones(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim j As Long
    
    With Npclist(NPCIndex)

        For j = 1 To .NroExpresiones
            .Expresiones(j) = vbNullString
        Next j
        
        .NroExpresiones = 0

    End With

End Sub

Private Sub ResetNpcMainInfo(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '22/05/2010: ZaMa - Ahora se resetea el dueno del npc tambien.
    '***************************************************

    Dim j As Long

    With Npclist(NPCIndex)
        .Attackable = 0
        .Comercia = 0
        .GiveEXP = 0
        .GiveGLD = 0
        .Hostile = 0
        .InvReSpawn = 0
        
        If .MaestroUser > 0 Then Call QuitarMascota(.MaestroUser, NPCIndex)
        If .MaestroNpc > 0 Then Call QuitarMascotaNpc(.MaestroNpc)
        If .Owner > 0 Then Call PerdioNpc(.Owner)
        
        .MaestroUser = 0
        .MaestroNpc = 0
        .Owner = 0
        
        .Mascotas = 0
        .Movement = 0
        .Name = vbNullString
        .NPCtype = 0
        .Numero = 0
        .Orig.Map = 0
        .Orig.X = 0
        .Orig.Y = 0
        .PoderAtaque = 0
        .PoderEvasion = 0
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .SkillDomar = 0
        .Target = 0
        .TargetNPC = 0
        .TipoItems = 0
        .Veneno = 0
        .Desc = vbNullString
        
        .ClanIndex = 0
        
        For j = 1 To .NroSpells
            .Spells(j) = 0
        Next j

    End With
    
    Call ResetNpcCharInfo(NPCIndex)
    Call ResetNpcCriatures(NPCIndex)
    Call ResetExpresiones(NPCIndex)

End Sub

Public Sub QuitarNPC(ByVal NPCIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Now npcs lose their owner
    '***************************************************
    On Error GoTo errHandler

    With Npclist(NPCIndex)
        .flags.NPCActive = False
        
        If InMapBounds(.Pos.Map, .Pos.X, .Pos.Y) Then
            Call EraseNPCChar(NPCIndex)

        End If

    End With
          
    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(NPCIndex)
    Call ResetNpcFlags(NPCIndex)
    Call ResetNpcCounters(NPCIndex)
    
    Call ResetNpcMainInfo(NPCIndex)
    
    If NPCIndex = LastNPC Then

        Do Until Npclist(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1

            If LastNPC < 1 Then Exit Do
        Loop

    End If
      
    If NumNPCs <> 0 Then
        NumNPCs = NumNPCs - 1

    End If

    Exit Sub

errHandler:
    Call LogError("Error en QuitarNPC")

End Sub

Public Sub QuitarPet(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 18/11/2009
    'Kills a pet
    '***************************************************
    On Error GoTo errHandler

    Dim i        As Integer

    Dim PetIndex As Integer

    With UserList(UserIndex)
        
        ' Busco el indice de la mascota
        For i = 1 To MAXMASCOTAS

            If .MascotasIndex(i) = NPCIndex Then PetIndex = i
        Next i
        
        ' Poco probable que pase, pero por las dudas..
        If PetIndex = 0 Then Exit Sub
        
        ' Limpio el slot de la mascota
        .NroMascotas = .NroMascotas - 1
        .MascotasIndex(PetIndex) = 0
        .MascotasType(PetIndex) = 0
        
        ' Elimino la mascota
        Call QuitarNPC(NPCIndex)

    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en QuitarPet. Error: " & Err.Number & " Desc: " & Err.description & " NpcIndex: " & NPCIndex & " UserIndex: " & UserIndex & " PetIndex: " & PetIndex)

End Sub

Private Function TestSpawnTrigger(Pos As WorldPos, _
                                  Optional PuedeAgua As Boolean = False) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    If LegalPos(Pos.Map, Pos.X, Pos.Y, PuedeAgua) Then
        TestSpawnTrigger = MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> eTrigger.POSINVALIDA And _
                           MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> eTrigger.CASA And _
                           MapData(Pos.Map, Pos.X, Pos.Y).Trigger <> eTrigger.BAJOTECHO

    End If
    
End Function

Public Function CrearNPC(NroNPC As Integer, _
                         Mapa As Integer, _
                         OrigPos As WorldPos, _
                         Optional ByVal CustomHead As Integer) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: 22/07/2019 - WyroX: Intentamos NO spawnear NPCs de agua en tierra, a menos que se alcance el límite de iteraciones.
    '
    '***************************************************

    'Crea un NPC del tipo NRONPC

    Dim Pos            As WorldPos

    Dim newPos         As WorldPos

    Dim altpos         As WorldPos

    Dim nIndex         As Integer

    Dim PosicionValida As Boolean

    Dim Iteraciones    As Long

    Dim PuedeAgua      As Boolean

    Dim PuedeTierra    As Boolean

    Dim Map            As Integer

    Dim X              As Integer

    Dim Y              As Integer

    nIndex = OpenNPC(NroNPC) 'Conseguimos un indice
    
    If nIndex > MAXNPCS Then Exit Function
    
    ' Cabeza customizada
    If CustomHead <> 0 Then Npclist(nIndex).Char.Head = CustomHead
    
    PuedeAgua = Npclist(nIndex).flags.AguaValida
    PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)
    
    'Necesita ser respawned en un lugar especifico
    If InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
        
        Map = OrigPos.Map
        X = OrigPos.X
        Y = OrigPos.Y
        Npclist(nIndex).Orig = OrigPos
        Npclist(nIndex).Pos = OrigPos
       
    Else
        
        Pos.Map = Mapa 'mapa
        altpos.Map = Mapa
        
        Do While Not PosicionValida
            Pos.X = RandomNumber(MinXBorder, MaxXBorder)    'Obtenemos posicion al azar en x
            Pos.Y = RandomNumber(MinYBorder, MaxYBorder)    'Obtenemos posicion al azar en y
            
            Call ClosestLegalPos(Pos, newPos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana

            If newPos.X <> 0 And newPos.Y <> 0 Then
                altpos.X = newPos.X
                altpos.Y = newPos.Y
            End If

            'Si X e Y son iguales a 0 significa que no se encontro posicion valida
            If LegalPos(newPos.Map, newPos.X, newPos.Y, PuedeAgua, PuedeTierra) And Not HayPCarea(newPos) And TestSpawnTrigger(newPos, PuedeAgua) Then
                'Asignamos las nuevas coordenas solo si son validas
                Npclist(nIndex).Pos.Map = newPos.Map
                Npclist(nIndex).Pos.X = newPos.X
                Npclist(nIndex).Pos.Y = newPos.Y
                PosicionValida = True
            End If
                
            'for debug
            Iteraciones = Iteraciones + 1

            If Iteraciones > MAXSPAWNATTEMPS Then
                If altpos.X <> 0 And altpos.Y <> 0 Then
                    Npclist(nIndex).Pos.Map = altpos.Map
                    Npclist(nIndex).Pos.X = altpos.X
                    Npclist(nIndex).Pos.Y = altpos.Y
                    PosicionValida = True
                Else
                    ' WyroX: Super� la cantidad de intentos sin ninguna posici�n v�lida? Probamos un intento m�s pero sin el flag "PuedeTierra"
                    Call ClosestLegalPos(Pos, newPos, PuedeAgua)

                    If newPos.X <> 0 And newPos.Y <> 0 Then
                        Npclist(nIndex).Pos.Map = newPos.Map
                        Npclist(nIndex).Pos.X = newPos.X
                        Npclist(nIndex).Pos.Y = newPos.Y
                        PosicionValida = True
                    Else
                        altpos.X = 50
                        altpos.Y = 50
                        Call ClosestLegalPos(altpos, newPos)

                        If newPos.X <> 0 And newPos.Y <> 0 Then
                            Npclist(nIndex).Pos.Map = newPos.Map
                            Npclist(nIndex).Pos.X = newPos.X
                            Npclist(nIndex).Pos.Y = newPos.Y
                            PosicionValida = True
                        Else
                            Call QuitarNPC(nIndex)
                            Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & Mapa & " NroNpc:" & NroNPC)
                            Exit Function

                        End If
                    End If
                End If

            End If

        Loop
            
        'asignamos las nuevas coordenas
        Map = Npclist(nIndex).Pos.Map
        X = Npclist(nIndex).Pos.X
        Y = Npclist(nIndex).Pos.Y

    End If
            
    'Crea el NPC
    Call MakeNPCChar(True, Map, nIndex, Map, X, Y)
    
    CrearNPC = nIndex
    
End Function

Public Sub MakeNPCChar(ByVal toMap As Boolean, _
                       sndIndex As Integer, _
                       NPCIndex As Integer, _
                       ByVal Map As Integer, _
                       ByVal X As Integer, _
                       ByVal Y As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '***************************************************
    
    Dim CharIndex As Integer
    Dim Color As Byte
    
    With Npclist(NPCIndex)
    
        If .Char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            .Char.CharIndex = CharIndex
            CharList(CharIndex) = NPCIndex
    
        End If
        
        MapData(Map, X, Y).NPCIndex = NPCIndex

        If Not toMap Then
            Call WriteCharacterCreate(sndIndex, .Char.body, .Char.Head, .Char.Heading, .Char.CharIndex, _
                X, Y, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, "", Color, 0, NingunAura, NingunAura)
    '
        Else
            Call AgregarNpc(NPCIndex)
    
        End If
    
    End With

End Sub

Public Sub ChangeNPCChar(ByVal NPCIndex As Integer, _
                         ByVal body As Integer, _
                         ByVal Head As Integer, _
                         ByVal Heading As eHeading)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If NPCIndex > 0 Then

        With Npclist(NPCIndex).Char
            .body = body
            .Head = Head
            .Heading = Heading
            
            Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessageCharacterChange(body, Head, Heading, .CharIndex, .WeaponAnim, .ShieldAnim, 0, 0, .CascoAnim, NingunAura, NingunAura))

        End With

    End If

End Sub

Private Sub EraseNPCChar(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If Npclist(NPCIndex).Char.CharIndex <> 0 Then CharList(Npclist(NPCIndex).Char.CharIndex) = 0

    If Npclist(NPCIndex).Char.CharIndex = LastChar Then

        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1

            If LastChar <= 1 Then Exit Do
        Loop

    End If

    'Quitamos del mapa
    MapData(Npclist(NPCIndex).Pos.Map, Npclist(NPCIndex).Pos.X, Npclist(NPCIndex).Pos.Y).NPCIndex = 0

    'Actualizamos los clientes
    Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessageCharacterRemove(Npclist(NPCIndex).Char.CharIndex))

    'Update la lista npc
    Npclist(NPCIndex).Char.CharIndex = 0

    'update NumChars
    NumChars = NumChars - 1

End Sub

Public Function MoveNPCChar(ByVal NPCIndex As Integer, ByVal nHeading As Byte) As Boolean
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 06/04/2009
    '06/04/2009: ZaMa - Now npcs can force to change position with dead character
    '01/08/2009: ZaMa - Now npcs can't force to chance position with a dead character if that means to change the terrain the character is in
    '26/09/2010: ZaMa - Turn sub into function to know if npc has moved or not.
    '***************************************************

    On Error GoTo errh

    Dim nPos      As WorldPos

    Dim UserIndex As Integer
    
    With Npclist(NPCIndex)
        nPos = .Pos
        Call HeadtoPos(nHeading, nPos)
        
        ' es una posicion legal
        If LegalPosNPC(nPos.Map, nPos.X, nPos.Y, .flags.AguaValida = 1, .MaestroUser <> 0) Then
            
            If .flags.AguaValida = 0 And HayAgua(.Pos.Map, nPos.X, nPos.Y) Then Exit Function
            If .flags.TierraInvalida = 1 And Not HayAgua(.Pos.Map, nPos.X, nPos.Y) Then Exit Function
            
            UserIndex = MapData(.Pos.Map, nPos.X, nPos.Y).UserIndex

            ' Si hay un usuario a donde se mueve el npc, entonces esta muerto o es un gm invisible
            If UserIndex > 0 Then
                
                ' No se traslada caspers de agua a tierra
                If HayAgua(.Pos.Map, nPos.X, nPos.Y) And Not HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then Exit Function

                ' No se traslada caspers de tierra a agua
                If Not HayAgua(.Pos.Map, nPos.X, nPos.Y) And HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then Exit Function
                
                'Se choca con los gm invisible si es que esta siguiendo a uno por el comando /seguir
                If .flags.SiguiendoGm = True And UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function
                
                With UserList(UserIndex)
                    ' Actualizamos posicion y mapa
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
                    .Pos.X = Npclist(NPCIndex).Pos.X
                    .Pos.Y = Npclist(NPCIndex).Pos.Y
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                        
                    ' Avisamos a los usuarios del area, y al propio usuario lo forzamos a moverse
                    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, .Pos.X, .Pos.Y))
                    Call WriteForceCharMove(UserIndex, InvertHeading(nHeading))

                End With

            End If
            
            Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))

            'Update map and user pos
            MapData(.Pos.Map, .Pos.X, .Pos.Y).NPCIndex = 0
            .Pos = nPos
            .Char.Heading = nHeading
            MapData(.Pos.Map, nPos.X, nPos.Y).NPCIndex = NPCIndex
            Call CheckUpdateNeededNpc(NPCIndex, nHeading)
        
            ' Npc has moved
            MoveNPCChar = True
        
        ElseIf .MaestroUser = 0 Then

            If .Movement = TipoAI.NpcPathfinding Then
                'Someone has blocked the npc's way, we must to seek a new path!
                .PFINFO.PathLenght = 0

            End If

        End If

    End With
    
    Exit Function

errh:
    LogError ("Error en move npc " & NPCIndex & ". Error: " & Err.Number & " - " & Err.description)

End Function

Function NextOpenNPC() As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Dim LoopC As Long
      
    For LoopC = 1 To MAXNPCS + 1

        If LoopC > MAXNPCS Then Exit For
        If Not Npclist(LoopC).flags.NPCActive Then Exit For
    Next LoopC
      
    NextOpenNPC = LoopC
    Exit Function

errHandler:
    Call LogError("Error en NextOpenNPC")

End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 10/07/2010
    '10/07/2010: ZaMa - Now npcs can't poison dead users.
    '***************************************************

    Dim n As Integer
    
    With UserList(UserIndex)

        If .flags.Muerto = 1 Then Exit Sub
        
        n = RandomNumber(1, 100)

        If n < 30 Then
            .flags.Envenenado = 1
            Call WriteConsoleMsg(UserIndex, "La criatura te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)

        End If

    End With
    
End Sub

Sub NpcParalizaUser(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 09/04/2021
    '***************************************************

    Dim n As Integer
    
    With UserList(UserIndex)

        If .flags.Muerto = 1 Then Exit Sub
        
        n = RandomNumber(1, 100)

        If n < 30 Then
            .flags.Paralizado = 1
            .Counters.Paralisis = IntervaloParalizado
            Call WriteParalizeOK(UserIndex)
            Call WriteConsoleMsg(UserIndex, "La criatura te ha paralizado!!", FontTypeNames.FONTTYPE_FIGHT)

        End If

    End With
    
End Sub

Sub NpcEntorpeceUser(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 09/04/2021
    '***************************************************

    Dim n As Integer
    
    With UserList(UserIndex)

        If .flags.Muerto = 1 Then Exit Sub
        
        n = RandomNumber(1, 100)

        If n < 30 Then
            .flags.Estupidez = 1
            .Counters.Ceguera = IntervaloInvisible
            
            Call WriteDumbNoMore(UserIndex)
            Call WriteConsoleMsg(UserIndex, "La criatura te ha estupidizado!!", FontTypeNames.FONTTYPE_FIGHT)

        End If

    End With
    
End Sub

Sub NpcCiegaUser(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 09/04/2021
    '***************************************************

    Dim n As Integer
    
    With UserList(UserIndex)

        If .flags.Muerto = 1 Then Exit Sub
        
        n = RandomNumber(1, 100)

        If n < 30 Then
            .flags.Ceguera = 1
            .Counters.Ceguera = IntervaloInvisible
                          
            Call WriteBlind(UserIndex)
            Call WriteConsoleMsg(UserIndex, "La criatura te ha estupidizado!!", FontTypeNames.FONTTYPE_FIGHT)

        End If

    End With
    
End Sub

Sub NpcDesarmaUser(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 09/04/2021
    '***************************************************

    Dim n As Integer
    
    With UserList(UserIndex)

        If .flags.Muerto = 1 Then Exit Sub
        
        If .Invent.WeaponEqpObjIndex > 0 Then
        
            ' Se lo desequipo
            Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                
            Call WriteConsoleMsg(UserIndex, "La criatura te ha desarmado!!", FontTypeNames.FONTTYPE_FIGHT)
                
            Exit Sub
        
        End If
    End With
    
End Sub

Sub NpcIncineraUser(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 05/09/2020
    'Descripci�n: Un NPC provoca quemaduras a un usuario
    '***************************************************

    Dim n As Integer
    
    With UserList(UserIndex)

        If .flags.Muerto = 1 Then Exit Sub
        
        n = RandomNumber(1, 100)

        If n < 30 Then
            .flags.Incinerado = 1
            Call WriteConsoleMsg(UserIndex, "La criatura te ha incinerado!!", FontTypeNames.FONTTYPE_FIGHT)

        End If

    End With
    
End Sub

Sub NpcParalizaNpc(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 09/04/2021
    '***************************************************

    Dim n As Integer
    
    With Npclist(NPCIndex)
        
        n = RandomNumber(1, 100)

        If n < 30 Then
            .flags.Paralizado = 1
            .Contadores.Paralisis = IntervaloParalizado

        End If

    End With
    
End Sub

Function SpawnNpc(ByVal NPCIndex As Integer, _
                  Pos As WorldPos, _
                  ByVal FX As Boolean, _
                  ByVal Respawn As Boolean, Optional ByVal OrigPos As Boolean = False) As Integer

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 06/15/2008
    '23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
    '06/15/2008 -> Optimize el codigo. (NicoNZ)
    '***************************************************
    Dim newPos         As WorldPos

    Dim altpos         As WorldPos

    Dim nIndex         As Integer

    Dim PosicionValida As Boolean

    Dim PuedeAgua      As Boolean

    Dim PuedeTierra    As Boolean

    Dim Map            As Integer

    Dim X              As Integer

    Dim Y              As Integer

    nIndex = OpenNPC(NPCIndex, Respawn)    'Conseguimos un indice

    If nIndex > MAXNPCS Then
        SpawnNpc = 0
        Exit Function

    End If

    PuedeAgua = Npclist(nIndex).flags.AguaValida
    PuedeTierra = Not Npclist(nIndex).flags.TierraInvalida = 1
        
    Call ClosestLegalPos(Pos, newPos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
    Call ClosestLegalPos(Pos, altpos, PuedeAgua)
    'Si X e Y son iguales a 0 significa que no se encontro posicion valida

    If newPos.X <> 0 And newPos.Y <> 0 Then
        'Asignamos las nuevas coordenas solo si son validas
        Npclist(nIndex).Pos.Map = newPos.Map
        Npclist(nIndex).Pos.X = newPos.X
        Npclist(nIndex).Pos.Y = newPos.Y
        PosicionValida = True
    Else

        If altpos.X <> 0 And altpos.Y <> 0 Then
            Npclist(nIndex).Pos.Map = altpos.Map
            Npclist(nIndex).Pos.X = altpos.X
            Npclist(nIndex).Pos.Y = altpos.Y
            PosicionValida = True
        Else
            PosicionValida = False

        End If

    End If

    If Not PosicionValida Then
        Call QuitarNPC(nIndex)
        SpawnNpc = 0
        Exit Function

    End If

    'asignamos las nuevas coordenas
    Map = newPos.Map
    X = Npclist(nIndex).Pos.X
    Y = Npclist(nIndex).Pos.Y
    
    '30/04/2016 - Lorwik: Se utiliza principalmente para los NPC con retardo de Spawn
    If OrigPos Then
        Npclist(nIndex).Orig.Map = Map
        Npclist(nIndex).Orig.X = X
        Npclist(nIndex).Orig.Y = Y
    End If

    'Crea el NPC
    Call MakeNPCChar(True, Map, nIndex, Map, X, Y)

    If FX Then
        Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
        Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(Npclist(nIndex).Char.CharIndex, FXIDs.FXWARP, 0))

    End If

    SpawnNpc = nIndex

End Function

Sub ReSpawnNpc(MiNPC As NPC)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.Map, MiNPC.Orig)

End Sub

Public Sub NPCTelep(ByVal NPCIndex As Integer, Posicion As WorldPos, ByVal FXTelep As Boolean)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 16/08/2020
    'Teletransporta a un NPC a una posicion
    '***************************************************

    Dim nHeading As eHeading

    With Npclist(NPCIndex)
    
        '�Es una posicion legal?
        If LegalPosNPC(Posicion.Map, Posicion.X, Posicion.Y, .flags.AguaValida = 1) Then
            
            If .flags.AguaValida = 0 And HayAgua(Posicion.Map, Posicion.X, Posicion.Y) Then Exit Sub
            If .flags.TierraInvalida = 1 And Not HayAgua(Posicion.Map, Posicion.X, Posicion.Y) Then Exit Sub
  
            Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessageCharacterMove(.Char.CharIndex, Posicion.X, Posicion.Y))
                
            'Sacamos el NPC de la antigua posicion
            MapData(.Pos.Map, .Pos.X, .Pos.Y).NPCIndex = 0
            
            'Cambiamos el Heading
            Call HeadtoPos(nHeading, Posicion)
            .Char.Heading = nHeading
            
            'Cambiamos la antigua por la nueva
            .Pos = Posicion
            
            'A�adimos el NPC a la nueva posici�n en el mapa
            MapData(Posicion.Map, Posicion.X, Posicion.Y).NPCIndex = NPCIndex
            
            Call CheckUpdateNeededNpc(NPCIndex, nHeading)
            
            '�Mostramos FX?
            If FXTelep Then
                Call SendData(SendTarget.ToPCArea, NPCIndex, PrepareMessagePlayWave(SND_WARP, Posicion.X, Posicion.Y))
                Call SendData(SendTarget.ToPCArea, NPCIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
                
            End If
                
        End If
    End With
    
End Sub

Private Sub NPCTirarOro(ByRef MiNPC As NPC)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'SI EL NPC TIENE ORO LO TIRAMOS
    If MiNPC.GiveGLD > 0 Then

        Dim MiObj As obj

        Dim MiAux As Long

        MiAux = MiNPC.GiveGLD

        Do While MiAux > MAX_INVENTORY_OBJS
            MiObj.Amount = MAX_INVENTORY_OBJS
            MiObj.ObjIndex = iORO
            Call TirarItemAlPiso(MiNPC.Pos, MiObj)
            MiAux = MiAux - MAX_INVENTORY_OBJS
        Loop

        If MiAux > 0 Then
            MiObj.Amount = MiAux
            MiObj.ObjIndex = iORO
            Call TirarItemAlPiso(MiNPC.Pos, MiObj)

        End If

    End If

End Sub

Public Function OpenNPC(ByVal NpcNumber As Integer, _
                        Optional ByVal Respawn = True) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    '###################################################
    '#               ATENCION PELIGRO                  #
    '###################################################
    '
    '     NO USAR GetVar PARA LEER LOS NPCS !!!!
    '
    'El que ose desafiar esta LEY, se las tendra que ver
    'conmigo. Para leer los NPCS se debera usar la
    'nueva clase clsIniManager.
    '
    'Alejo
    '
    '###################################################
    Dim NPCIndex As Integer

    Dim Leer     As clsIniManager

    Dim LoopC    As Long

    Dim ln       As String
    
    Set Leer = LeerNPCs
    
    'If requested index is invalid, abort
    If Not Leer.KeyExists("NPC" & NpcNumber) Then
        OpenNPC = MAXNPCS + 1
        Exit Function

    End If
    
    NPCIndex = NextOpenNPC
    
    If NPCIndex > MAXNPCS Then 'Limite de npcs
        OpenNPC = NPCIndex
        Exit Function

    End If
    
    With Npclist(NPCIndex)
        .Numero = NpcNumber
        .Name = Leer.GetValue("NPC" & NpcNumber, "Name")
        .Desc = Leer.GetValue("NPC" & NpcNumber, "Desc")
        
        .Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
        .flags.OldMovement = .Movement
        
        .flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
        .flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
        .flags.Faccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))
        .flags.AtacaDoble = val(Leer.GetValue("NPC" & NpcNumber, "AtacaDoble"))
        
        .NPCtype = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))
        
        .Char.body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
        .Char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
        .Char.Heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))
        .Char.ShieldAnim = val(Leer.GetValue("NPC" & NpcNumber, "ShieldAnim"))
        .Char.WeaponAnim = val(Leer.GetValue("NPC" & NpcNumber, "WeaponAnim"))
        .Char.CascoAnim = val(Leer.GetValue("NPC" & NpcNumber, "CascoAnim"))
        
        .Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
        .Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
        .Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
        .flags.OldHostil = .Hostile
        
        .GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP")) * ExpMultiplier
        If HappyHourActivated And (HappyHour <> 0) Then .GiveEXP = .GiveEXP * HappyHour
        
        .flags.ExpCount = .GiveEXP
        
        .Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))
        .Quema = val(Leer.GetValue("NPC" & NpcNumber, "Quema"))
        .Paraliza = val(Leer.GetValue("NPC" & NpcNumber, "Paraliza"))
        .Entorpece = val(Leer.GetValue("NPC" & NpcNumber, "Entorpece"))
        .Ciega = val(Leer.GetValue("NPC" & NpcNumber, "Ciega"))
        .Desarma = val(Leer.GetValue("NPC" & NpcNumber, "Desarma"))
        
        .flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))
        
        .GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD")) * OroMultiplier
        
        .PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
        .PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))
        
        .InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))
        
        With .Stats
            .MaxHp = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
            .MinHp = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
            .MaxHit = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
            .MinHIT = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
            .def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
            .defM = val(Leer.GetValue("NPC" & NpcNumber, "DEFm"))
            .Alineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))

        End With
        
        .Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))

        For LoopC = 1 To .Invent.NroItems
            ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
            .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
            .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
            .Invent.Object(LoopC).RandomDrop = val(ReadField(3, ln, 45))
        Next LoopC
        
        .flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))

        If .flags.LanzaSpells > 0 Then ReDim .Spells(1 To .flags.LanzaSpells)

        For LoopC = 1 To .flags.LanzaSpells
            .Spells(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
        Next LoopC
        
        If .NPCtype = eNPCType.Entrenador Then
            .NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
            ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador

            For LoopC = 1 To .NroCriaturas
                .Criaturas(LoopC).NPCIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
                .Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
            Next LoopC

        End If
        
        With .flags
            .NPCActive = True
            
            If Respawn Then
                .Respawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
            Else
                .Respawn = 1

            End If
            
            .BackUp = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
            .RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
            .AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
            
            .Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
            .Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
            .Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))
            
            .TiempoRetardoMax = val(Leer.GetValue("NPC" & NpcNumber, "TiempoRetardoMax"))
            .TiempoRetardoMin = val(Leer.GetValue("NPC" & NpcNumber, "TiempoRetardoMin"))
  
        End With
        
        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        .NroExpresiones = val(Leer.GetValue("NPC" & NpcNumber, "NROEXP"))

        If .NroExpresiones > 0 Then ReDim .Expresiones(1 To .NroExpresiones) As String

        For LoopC = 1 To .NroExpresiones
            .Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
        Next LoopC

        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        
        'Tipo de items con los que comercia
        .TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))
        
        .Ciudad = val(Leer.GetValue("NPC" & NpcNumber, "Ciudad"))
        
        .Instruye = val(Leer.GetValue("NPC" & NpcNumber, "Instruye"))
        
        .SpeedVar = val(Leer.GetValue("NPC" & NpcNumber, "Speed"))

    End With
    
    'Update contadores de NPCs
    If NPCIndex > LastNPC Then LastNPC = NPCIndex
    NumNPCs = NumNPCs + 1
    
    'Devuelve el nuevo Indice
    OpenNPC = NPCIndex

End Function

Public Sub DoFollow(ByVal NPCIndex As Integer, ByVal UserName As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With Npclist(NPCIndex)

        If .flags.Follow Then
            .flags.AttackedBy = vbNullString
            .flags.Follow = False
            .flags.SiguiendoGm = False
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
        Else
            .flags.AttackedBy = UserName
            .flags.Follow = True
            .flags.SiguiendoGm = True
            .Movement = TipoAI.NPCDEFENSA
            .Hostile = 0

        End If

    End With

End Sub

Public Sub FollowAmo(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With Npclist(NPCIndex)
        .flags.Follow = True
        .Movement = TipoAI.SigueAmo
        .Hostile = 0
        .Target = 0
        .TargetNPC = 0

    End With

End Sub

Public Sub ValidarPermanenciaNpc(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Chequea si el npc continua perteneciendo a algun usuario
    '***************************************************

    With Npclist(NPCIndex)

        If IntervaloPerdioNpc(.Owner) Then Call PerdioNpc(.Owner)

    End With

End Sub
