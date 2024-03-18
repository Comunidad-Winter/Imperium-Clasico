Attribute VB_Name = "SistemaCombate"

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
'
'Diseno y correccion del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

'9/01/2008 Pablo (ToxicWaste) - Ahora TODOS los modificadores de Clase se controlan desde Balance.dat

Option Explicit

Public Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer

    If a > b Then
        MinimoInt = b
    Else
        MinimoInt = a

    End If

End Function

Public Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer

    If a > b Then
        MaximoInt = a
    Else
        MaximoInt = b

    End If

End Function

Public Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.Defensa) * ModClase(UserList(UserIndex).clase).Escudo) / 2

End Function

Public Function PoderEvasion(ByVal UserIndex As Integer) As Long

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Dim lTemp As Long

    With UserList(UserIndex)
        lTemp = (.Stats.UserSkills(eSkill.Tacticas) + .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).Evasion
       
        PoderEvasion = (lTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))

    End With

End Function

Private Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)

        If .Stats.UserSkills(eSkill.Armas) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Armas) * ModClase(.clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Armas) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Armas) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueArmas
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueArmas

        End If
        
        PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))

    End With

End Function

Private Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim PoderAtaqueTemp  As Long

    Dim SkillProyectiles As Integer
    
    With UserList(UserIndex)
     
        SkillProyectiles = .Stats.UserSkills(eSkill.Proyectiles)
    
        If SkillProyectiles < 31 Then
            PoderAtaqueTemp = SkillProyectiles * ModClase(.clase).AtaqueProyectiles
        ElseIf SkillProyectiles < 61 Then
            PoderAtaqueTemp = (SkillProyectiles + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueProyectiles
        ElseIf SkillProyectiles < 91 Then
            PoderAtaqueTemp = (SkillProyectiles + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueProyectiles
        Else
            PoderAtaqueTemp = (SkillProyectiles + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueProyectiles

        End If
        
        PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))

    End With

End Function

Private Function PoderAtaqueMarciales(ByVal UserIndex As Integer) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim PoderAtaqueTemp As Long

    Dim MarcialesSkill  As Integer
    
    With UserList(UserIndex)
    
        MarcialesSkill = .Stats.UserSkills(eSkill.Marciales)
    
        If MarcialesSkill < 31 Then
            PoderAtaqueTemp = MarcialesSkill * ModClase(.clase).AtaqueMarciales
        ElseIf MarcialesSkill < 61 Then
            PoderAtaqueTemp = (MarcialesSkill + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueMarciales
        ElseIf MarcialesSkill < 91 Then
            PoderAtaqueTemp = (MarcialesSkill + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueMarciales
        Else
            PoderAtaqueTemp = (MarcialesSkill + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueMarciales

        End If
        
        PoderAtaqueMarciales = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))

    End With

End Function

Public Function UserImpactoNpc(ByVal UserIndex As Integer, _
                               ByVal NPCIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim PoderAtaque As Long

    Dim Arma        As Integer

    Dim Skill       As eSkill

    Dim ProbExito   As Long
    
    Dim MunicionObjIndex    As Integer
    
    Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
    
    If Arma > 0 Then 'Usando un arma
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(UserIndex)
            Skill = eSkill.Proyectiles

        Else
            PoderAtaque = PoderAtaqueArma(UserIndex)
            Skill = eSkill.Armas

        End If

    Else 'Peleando con punos
        PoderAtaque = PoderAtaqueMarciales(UserIndex)
        Skill = eSkill.Marciales

    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - Npclist(NPCIndex).PoderEvasion) * 0.4)))
    
    UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
    
    If UserImpactoNpc Then
        Call SubirSkill(UserIndex, Skill, True)
    Else
        Call SubirSkill(UserIndex, Skill, False)

    End If

End Function

Public Function NpcImpacto(ByVal NPCIndex As Integer, _
                           ByVal UserIndex As Integer) As Boolean

    '*************************************************
    'Author: Unknown
    'Last modified: 03/15/2006
    'Revisa si un NPC logra impactar a un user o no
    '03/15/2006 Maraxus - Evite una division por cero que eliminaba NPCs
    '*************************************************
    Dim Rechazo           As Boolean

    Dim ProbRechazo       As Long

    Dim ProbExito         As Long

    Dim UserEvasion       As Long

    Dim NpcPoderAtaque    As Long

    Dim PoderEvasioEscudo As Long

    Dim SkillTacticas     As Long

    Dim SkillDefensa      As Long
    
    UserEvasion = PoderEvasion(UserIndex)
    NpcPoderAtaque = Npclist(NPCIndex).PoderAtaque
    PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)
    
    SkillTacticas = UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(UserIndex).Stats.UserSkills(eSkill.Defensa)
    
    'Esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
    
    NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    ' el usuario esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        If Not NpcImpacto Then
            If SkillDefensa + SkillTacticas > 0 Then  'Evitamos division por cero
                ' Chances are rounded
                ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
            Else
                ProbRechazo = 10 'Si no tiene skills le dejamos el 10% minimo

            End If
            
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                
            If Rechazo Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO(RandomNumber(1, 4)), UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                Call WriteMultiMessage(UserIndex, eMessages.BlockedWithShieldUser) 'Call WriteBlockedWithShieldUser(UserIndex)
                Call SubirSkill(UserIndex, eSkill.Defensa, True)
            Else
                Call SubirSkill(UserIndex, eSkill.Defensa, False)

            End If

        End If

    End If

End Function

Public Function CalcularDano(ByVal UserIndex As Integer, _
                             Optional ByVal NPCIndex As Integer = 0) As Long

    '***************************************************
    'Author: Unknown
    'Last Modification: 01/04/2010 (ZaMa)
    '01/04/2010: ZaMa - Modifico el dano de Marciales.
    '16/03/2021: Lorwik - Agrego bonificadores de Marciales para los nudillos.
    '***************************************************
    Dim DanoArma    As Long

    Dim DanoUsuario As Long

    Dim Arma        As ObjData

    Dim ModifClase  As Single

    Dim proyectil   As ObjData

    Dim DanoMaxArma As Long

    Dim DanoMinArma As Long
    
    With UserList(UserIndex)

        If .Invent.WeaponEqpObjIndex > 0 Then
            Arma = ObjData(.Invent.WeaponEqpObjIndex)
            
            ' Ataca a un npc?
            If NPCIndex > 0 Then
                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.clase).DanoProyectiles
                    DanoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
                    DanoMaxArma = Arma.MaxHit
                    
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DanoArma = DanoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHit)

                        ' For some reason this isn't done...
                        'DanoMaxArma = DanoMaxArma + proyectil.MaxHIT
                    End If

                Else
                    ModifClase = ModClase(.clase).DanoArmas
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mata Dragones?
                        If Npclist(NPCIndex).NPCtype = DRAGON Then 'Ataca Dragon?
                            DanoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
                            DanoMaxArma = Arma.MaxHit
                            
                        Else ' Sino es Dragon dano es 1
                            DanoArma = 1
                            DanoMaxArma = 1

                        End If

                    Else
                        DanoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
                        DanoMaxArma = Arma.MaxHit

                    End If

                End If

            Else ' Ataca usuario

                If Arma.proyectil = 1 Then
                    ModifClase = ModClase(.clase).DanoProyectiles
                    DanoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
                    DanoMaxArma = Arma.MaxHit
                     
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DanoArma = DanoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHit)

                        ' For some reason this isn't done...
                        'DanoMaxArma = DanoMaxArma + proyectil.MaxHIT
                    End If

                Else
                    ModifClase = ModClase(.clase).DanoArmas
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                        ModifClase = ModClase(.clase).DanoArmas
                        DanoArma = 1 ' Si usa la espada mataDragones dano es 1
                        DanoMaxArma = 1
                    Else
                        DanoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
                        DanoMaxArma = Arma.MaxHit

                    End If

                End If

            End If

        Else
            ModifClase = ModClase(.clase).DanoMarciales

            ' Dano sin nudillos
            DanoMinArma = 1
            DanoMaxArma = 3
            
            If .Invent.NudiEqpIndex > 0 Then
                Arma = ObjData(.Invent.NudiEqpIndex)
                DanoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
                DanoMaxArma = Arma.MaxHit
            End If
            
            DanoArma = DanoArma + RandomNumber(DanoMinArma, DanoMaxArma) 'Hacemos que sea "tipo" una daga el ataque de Wrestling
            DanoMaxArma = DanoMaxArma * 2
        End If
        
        DanoUsuario = RandomNumber(.Stats.MinHIT, .Stats.MaxHit)
        
        CalcularDano = (3 * DanoArma + ((DanoMaxArma / 5) * MaximoInt(0, .Stats.UserAtributos(eAtributos.Fuerza) - 15)) + DanoUsuario) * ModifClase

    End With

End Function

Public Sub UserDanoNpc(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 07/04/2010 (Pato)
    '25/01/2010: ZaMa - Agrego poder acuchillar npcs.
    '07/04/2010: ZaMa - Los asesinos apunalan acorde al dano base sin descontar la defensa del npc.
    '07/04/2010: Pato - Si se mata al dragon en party se loguean los miembros de la misma.
    '11/07/2010: ZaMa - Ahora la defensa es solo ignorada para asesinos.
    '***************************************************

    Dim dano                                 As Long

    Dim DanoBase                             As Long

    Dim PI                                   As Integer

    Dim MembersOnline(1 To PARTY_MAXMEMBERS) As Integer

    Dim Text                                 As String

    Dim i                                    As Integer
    
    Dim BoatIndex                            As Integer
    
    DanoBase = CalcularDano(UserIndex, NPCIndex)
    
    'Al menos que pegue por 1
    If DanoBase < 1 Then DanoBase = 1
    
    '�Tiene brazalete de ogro?
    If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).Efectomagico = eEfectos.AumentaGolpe Then _
            DanoBase = DanoBase + ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).CuantoAumento
    End If
    
    'esta navegando? si es asi le sumamos el dano del barco
    If UserList(UserIndex).flags.Navegando = 1 Then
    
        BoatIndex = UserList(UserIndex).Invent.BarcoObjIndex

        If BoatIndex > 0 Then
            DanoBase = DanoBase + RandomNumber(ObjData(BoatIndex).MinHIT, ObjData(BoatIndex).MaxHit)

        End If

    End If
    
    With Npclist(NPCIndex)
    
        dano = DanoBase - .Stats.def
        
        If dano < 0 Then dano = 0
        
        Call WriteMultiMessage(UserIndex, eMessages.UserHitNPC, dano)
        Call CalcularDarExp(UserIndex, NPCIndex, dano)

        .Stats.MinHp = .Stats.MinHp - dano
        
        'Renderizo dano en render
        Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(dano, UserList(UserIndex).Char.CharIndex, vbBlue, True))
        
        If .Stats.MinHp > 0 Then

            'Trata de apunalar por la espalda al enemigo
            If PuedeApunalar(UserIndex) Then
                
                ' La defensa se ignora solo en asesinos
                If UserList(UserIndex).clase <> eClass.Assasin Then
                    DanoBase = dano

                End If
                
                Call DoApunalar(UserIndex, NPCIndex, 0, DanoBase)
                
            End If
            
            'trata de dar golpe critico
            Call DoGolpeCritico(UserIndex, NPCIndex, 0, dano)
            
            If PuedeAcuchillar(UserIndex) Then
                Call DoAcuchillar(UserIndex, NPCIndex, 0, dano)

            End If

        End If
        
        If .Stats.MinHp <= 0 Then

            ' Si era un Dragon rompemos la espada mataDragones
            If .NPCtype = DRAGON Then

                If .Stats.MaxHp > 100000 Then
                    Text = UserList(UserIndex).Name & " mato un dragon"
                    PI = UserList(UserIndex).PartyIndex
                    
                    If PI > 0 Then
                        Call Parties(PI).ObtenerMiembrosOnline(MembersOnline())
                        Text = Text & " estando en party "
                        
                        For i = 1 To PARTY_MAXMEMBERS

                            If MembersOnline(i) > 0 Then
                                Text = Text & UserList(MembersOnline(i)).Name & ", "

                            End If

                        Next i
                        
                        Text = Left$(Text, Len(Text) - 2) & ")"

                    End If
                    
                    Call LogDesarrollo(Text & ".")

                End If

            End If
            
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            For i = 1 To MAXMASCOTAS

                If UserList(UserIndex).MascotasIndex(i) > 0 Then
                    If Npclist(UserList(UserIndex).MascotasIndex(i)).TargetNPC = NPCIndex Then
                        Npclist(UserList(UserIndex).MascotasIndex(i)).TargetNPC = 0
                        Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = TipoAI.SigueAmo

                    End If

                End If

            Next i
            
            Call MuereNpc(NPCIndex, UserIndex)

        End If

    End With

End Sub

Public Sub NpcDano(ByVal NPCIndex As Integer, ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 18/09/2010 (ZaMa)
    '18/09/2010: ZaMa - Ahora se considera siempre la defensa del barco y el escudo.
    '***************************************************

    Dim dano        As Integer

    Dim Lugar       As Integer

    Dim obj         As ObjData
    
    Dim BoatDefense As Integer

    Dim HeadDefense As Integer

    Dim BodyDefense As Integer
    
    Dim BoatIndex   As Integer

    Dim HelmetIndex As Integer

    Dim ArmourIndex As Integer

    Dim ShieldIndex As Integer
    
    dano = RandomNumber(Npclist(NPCIndex).Stats.MinHIT, Npclist(NPCIndex).Stats.MaxHit)
    
    With UserList(UserIndex)

        ' Navega?
        If .flags.Navegando = 1 Then
            ' En barca suma defensa
            BoatIndex = .Invent.BarcoObjIndex

            If BoatIndex > 0 Then
                obj = ObjData(BoatIndex)
                BoatDefense = RandomNumber(obj.MinDef, obj.MaxDef)

            End If

        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
        
            Case PartesCuerpo.bCabeza
            
                'Si tiene casco absorbe el golpe
                HelmetIndex = .Invent.CascoEqpObjIndex

                If HelmetIndex > 0 Then
                    obj = ObjData(HelmetIndex)
                    HeadDefense = RandomNumber(obj.MinDef, obj.MaxDef)

                End If
                
            Case Else
                
                Dim MinDef As Integer

                Dim MaxDef As Integer
            
                'Si tiene armadura absorbe el golpe
                ArmourIndex = .Invent.ArmourEqpObjIndex

                If ArmourIndex > 0 Then
                    obj = ObjData(ArmourIndex)
                    MinDef = obj.MinDef
                    MaxDef = obj.MaxDef

                End If
                
                ' Si tiene casco absorbe el golpe
                ShieldIndex = .Invent.EscudoEqpObjIndex

                If ShieldIndex > 0 Then
                    obj = ObjData(ShieldIndex)
                    MinDef = MinDef + obj.MinDef
                    MaxDef = MaxDef + obj.MaxDef

                End If
                
                BodyDefense = RandomNumber(MinDef, MaxDef)
        
        End Select
        
        ' Dano final
        dano = dano - HeadDefense - BodyDefense - BoatDefense

        If dano < 1 Then dano = 1
        
        Call WriteMultiMessage(UserIndex, eMessages.NPCHitUser, Lugar, dano)
        
        If .flags.Privilegios And PlayerType.User Then .Stats.MinHp = .Stats.MinHp - dano
        
        'Renderizo el dano en render.
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(dano, Npclist(NPCIndex).Char.CharIndex, vbBlue, True))
        
        If .flags.Meditando Then
            If dano > Fix(.Stats.MinHp / 100 * .Stats.UserAtributos(eAtributos.Inteligencia) * .Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
                .flags.Meditando = False
                Call WriteMeditateToggle(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                .Char.Particle = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticleChar(.Char.CharIndex, .Char.Particle, False, 0))

            End If

        End If
        
        'Muere el usuario
        If .Stats.MinHp <= 0 Then
            Call WriteMultiMessage(UserIndex, eMessages.NPCKillUser)  'Le informamos que ha muerto ;)
            
            'Si lo mato un guardia
            If criminal(UserIndex) Then
                If Npclist(NPCIndex).NPCtype = eNPCType.GuardiaReal Then
                    Call RestarCriminalidad(UserIndex)

                End If

            End If
            
            If Npclist(NPCIndex).MaestroUser > 0 Then
                Call AllFollowAmo(Npclist(NPCIndex).MaestroUser)
            Else

                'Al matarlo no lo sigue mas
                With Npclist(NPCIndex)

                    If .Stats.Alineacion = 0 Then
                        .Movement = .flags.OldMovement
                        .Hostile = .flags.OldHostil
                        .flags.AttackedBy = vbNullString

                    End If

                End With
                
            End If
            
            Call UserDie(UserIndex)

        End If

    End With

End Sub

Public Sub RestarCriminalidad(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim EraCriminal As Boolean

    EraCriminal = criminal(UserIndex)
    
    With UserList(UserIndex).Reputacion

        If .BandidoRep > 0 Then
            .BandidoRep = .BandidoRep - vlASALTO

            If .BandidoRep < 0 Then .BandidoRep = 0
        ElseIf .LadronesRep > 0 Then
            .LadronesRep = .LadronesRep - (vlCAZADOR * 10)

            If .LadronesRep < 0 Then .LadronesRep = 0

        End If
    
        If EraCriminal And Not criminal(UserIndex) Then
        
            If esCaos(UserIndex) Then Call ExpulsarFaccionCaos(UserIndex)
            
            Call RefreshCharStatus(UserIndex)

        End If
    
    End With

End Sub

Public Sub CheckPets(ByVal NPCIndex As Integer, _
                     ByVal UserIndex As Integer, _
                     Optional ByVal CheckElementales As Boolean = True)
    '***************************************************
    'Author: Unknown
    'Last Modification: 15/04/2010
    '15/04/2010: ZaMa - Las mascotas no se apropian de npcs.
    '***************************************************

    Dim j As Integer
    
    ' Si no tengo mascotas, para que cheaquear lo demas?
    If UserList(UserIndex).NroMascotas = 0 Then Exit Sub
    
    If Not PuedeAtacarNPC(UserIndex, NPCIndex, , True) Then Exit Sub
    
    With UserList(UserIndex)

        For j = 1 To MAXMASCOTAS

            If .MascotasIndex(j) > 0 Then
                If .MascotasIndex(j) <> NPCIndex Then
                    If CheckElementales Or (Npclist(.MascotasIndex(j)).Numero <> ELEMENTALFUEGO And Npclist(.MascotasIndex(j)).Numero <> ELEMENTALTIERRA) Then
                    
                        If Npclist(.MascotasIndex(j)).TargetNPC = 0 Then Npclist(.MascotasIndex(j)).TargetNPC = NPCIndex
                        Npclist(.MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc

                    End If

                End If

            End If

        Next j

    End With

End Sub

Public Sub AllFollowAmo(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim j As Integer
    
    For j = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasIndex(j) > 0 Then
            Call FollowAmo(UserList(UserIndex).MascotasIndex(j))

        End If

    Next j

End Sub

Public Function NpcAtacaUser(ByVal NPCIndex As Integer, _
                             ByVal UserIndex As Integer) As Boolean
    '*************************************************
    'Author: Unknown
    'Last Modification: 06/04/2020
    '06/04/2020: FrankoH298 - Si un npc ataca a un usuario lo desmonta.
    '*************************************************

    With UserList(UserIndex)

        If .flags.AdminInvisible = 1 Then Exit Function
        If (Not .flags.Privilegios And PlayerType.User) <> 0 And Not .flags.AdminPerseguible Then Exit Function
        
        '�Esta trabajando?
        If UserList(UserIndex).flags.MacroTrabajo <> 0 Then
            Call DejardeTrabajar(UserIndex)
        End If
        
        '<<<< Equitando >>>
        If .flags.Equitando = 1 Then
            Call UnmountMontura(UserIndex)
            Call WriteEquitandoToggle(UserIndex)
            
        End If
    End With
    
    With Npclist(NPCIndex)

        ' El npc puede atacar ???
        'If .CanAttack = 1 Then
        If IntervaloPermiteAtacarNpc(NPCIndex) Then
            NpcAtacaUser = True
            Call CheckPets(NPCIndex, UserIndex, False)
            
            If .Target = 0 Then .Target = UserIndex
            
            If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then
                UserList(UserIndex).flags.AtacadoPorNpc = NPCIndex

            End If

        Else
            NpcAtacaUser = False
            Exit Function

        End If
        
        '.CanAttack = 0
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))

        End If
        
        If UserList(UserIndex).Familiar.Invocado = 1 And Npclist(UserList(UserIndex).Familiar.Id).TargetNPC = 0 Then
            Npclist(UserList(UserIndex).Familiar.Id).TargetNPC = NPCIndex
            Npclist(UserList(UserIndex).Familiar.Id).Movement = TipoAI.NpcAtacaNpc
            
            Npclist(NPCIndex).TargetNPC = UserList(UserIndex).Familiar.Id
            Npclist(NPCIndex).Movement = TipoAI.NpcAtacaNpc
        End If

    End With
    
    If NpcImpacto(NPCIndex, UserIndex) Then

        With UserList(UserIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            If .flags.Meditando = False Then
                If .flags.Navegando = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0))

                End If

            End If
            
            Call NpcDano(NPCIndex, UserIndex)
            Call WriteUpdateHP(UserIndex)
            
            'Puede envenenar?
            If Npclist(NPCIndex).Veneno = 1 Then Call NpcEnvenenarUser(UserIndex)
            
            'Puede incinerar?
            If Npclist(NPCIndex).Quema = 1 Then Call NpcIncineraUser(UserIndex)
            
            'Puede paralizar?
            If Npclist(NPCIndex).Paraliza = 1 Then Call NpcParalizaUser(UserIndex)
            
            'Puede Entorpecer?
            If Npclist(NPCIndex).Entorpece = 1 Then Call NpcEntorpeceUser(UserIndex)
            
            'Puede cegar?
            If Npclist(NPCIndex).Ciega = 1 Then Call NpcCiegaUser(UserIndex)
            
            'Puede desarmar?
            If Npclist(NPCIndex).Desarma = 1 Then Call NpcDesarmaUser(UserIndex)

        End With
        
        Call SubirSkill(UserIndex, eSkill.Tacticas, False)
        
    Else
        Call WriteMultiMessage(UserIndex, eMessages.NPCSwing)
        
        'Renderizo el dano en render.
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("'FALLA'", Npclist(NPCIndex).Char.CharIndex, vbBlue, True))
        
        Call SubirSkill(UserIndex, eSkill.Tacticas, True)

    End If
    
    'Controla el nivel del usuario
    Call CheckUserLevel(UserIndex)
    
    If UserList(UserIndex).Familiar.Invocado = 1 Then
        Call CheckFamilyLevel(UserIndex)
    End If

End Function

Private Function NpcImpactoNpc(ByVal Atacante As Integer, _
                               ByVal Victima As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim PoderAtt  As Long

    Dim PoderEva  As Long

    Dim ProbExito As Long
    
    PoderAtt = Npclist(Atacante).PoderAtaque
    PoderEva = Npclist(Victima).PoderEvasion
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtt - PoderEva) * 0.4))
    NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

End Function

Public Sub NpcDanoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim dano        As Integer
    
    With Npclist(Atacante)
        dano = RandomNumber(.Stats.MinHIT, .Stats.MaxHit)
        
        If .MaestroUser <> 0 Then
            Call WriteConsoleMsg(.MaestroUser, Npclist(Atacante).Name & " ha golpeado a la criatura por " & dano, FontTypeNames.FONTTYPE_FIGHT)
            Call CalcularDarExp(.MaestroUser, Victima, Abs(dano))
        End If
        
        Npclist(Victima).Stats.MinHp = Npclist(Victima).Stats.MinHp - dano
        
        'Renderizo el dano en render.
        Call SendData(SendTarget.ToPCArea, Atacante, PrepareMessageChatOverHead(dano, .Char.CharIndex, vbBlue, True))
        
        If Npclist(Victima).Stats.MinHp < 1 Then
            .Movement = .flags.OldMovement
            
            If LenB(.flags.AttackedBy) <> 0 Then
                .Hostile = .flags.OldHostil

            End If

            If .MaestroUser > 0 Then
                Call FollowAmo(Atacante)

            End If
            
            Call MuereNpc(Victima, .MaestroUser)

        End If
 
    End With

End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, _
                       ByVal Victima As Integer, _
                       Optional ByVal cambiarMOvimiento As Boolean = True)
    '*************************************************
    'Author: Unknown
    'Last modified: 01/03/2009
    '23/05/2010: ZaMa - Ahora los elementales renuevan el tiempo de pertencia del npc que atacan si pertenece a su amo.
    '*************************************************
    
    Dim MasterIndex As Integer
    With Npclist(Atacante)
        
        ' El npc puede atacar ???
        
        If IntervaloPermiteAtacarNpc(Atacante) Then
            If cambiarMOvimiento Then
                Npclist(Victima).TargetNPC = Atacante
                Npclist(Victima).Movement = TipoAI.NpcAtacaNpc

            End If

        Else
            Exit Sub

        End If
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))

        End If
        
        MasterIndex = .MaestroUser
        
        ' Tiene maestro?
        If MasterIndex > 0 Then

            ' Su maestro es dueno del npc al que ataca?
            If Npclist(Victima).Owner = MasterIndex Then
                ' Renuevo el timer de pertenencia
                Call IntervaloPerdioNpc(MasterIndex, True)

            End If

        End If
        
        If NpcImpactoNpc(Atacante, Victima) Then
            If Npclist(Victima).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))

            End If
        
            If MasterIndex > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))

            End If
            
            Call NpcDanoNpc(Atacante, Victima)
            
            'Puede paralizar?
            If Npclist(Atacante).Paraliza = 1 Then Call NpcParalizaNpc(Victima)
            
        Else

            If MasterIndex > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
                
                'Renderizo el dano en render.
                Call SendData(SendTarget.ToPCArea, Atacante, PrepareMessageChatOverHead("'FALLA'", .Char.CharIndex, vbBlue, True))
                
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_SWING, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
                
                'Renderizo el dano en render.
                Call SendData(SendTarget.ToPCArea, Atacante, PrepareMessageChatOverHead("'FALLA'", .Char.CharIndex, vbBlue, True))
                
            End If

        End If

    End With
End Sub

Public Function UsuarioAtacaNpc(ByVal UserIndex As Integer, _
                                ByVal NPCIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 06/04/2020
    '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados por npcs cuando los atacan.
    '14/01/2010: ZaMa - Lo transformo en funcion, para que no se pierdan municiones al atacar targets invalidos.
    '13/02/2011: Amraphen - Ahora la stamina es quitada cuando efectivamente se ataca al NPC.
    '06/04/2020: FrankoH298 - Si un usuario ataca a un npc lo desmonta.
    '***************************************************

    On Error GoTo errHandler
    
    If Not PuedeAtacarNPC(UserIndex, NPCIndex) Then Exit Function
    
    With UserList(UserIndex)
        '<<<< Equitando >>>
        If .flags.Equitando = 1 Then
            Call UnmountMontura(UserIndex)
            Call WriteEquitandoToggle(UserIndex)
            
        End If
        
    End With
    
    Call NPCAtacado(NPCIndex, UserIndex)
    
    If UserImpactoNpc(UserIndex, NPCIndex) Then
        If Npclist(NPCIndex).flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessagePlayWave(Npclist(NPCIndex).flags.Snd2, Npclist(NPCIndex).Pos.X, Npclist(NPCIndex).Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(NPCIndex).Pos.X, Npclist(NPCIndex).Pos.Y))

        End If
        
        Call UserDanoNpc(UserIndex, NPCIndex)
    Else
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        
        'Renderizo el dano en render.
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("'FALLAS'", UserList(UserIndex).Char.CharIndex, vbBlue, True))
        
        Call WriteMultiMessage(UserIndex, eMessages.UserSwing)

    End If
    
    '�Paraliza?
    Call GolpeParalizaNpc(UserIndex, NPCIndex)
    
    ' Revelo su condicion de usuario al atacar, los npcs lo van a atacar
    UserList(UserIndex).flags.Ignorado = False
    
    UsuarioAtacaNpc = True
    
    Exit Function
    
errHandler:

    Dim UserName As String
    
    If UserIndex > 0 Then UserName = UserList(UserIndex).Name
    
    Call LogError("Error en UsuarioAtacaNpc. Error " & Err.Number & " : " & Err.description & ". User: " & UserIndex & "-> " & UserName & ". NpcIndex: " & NPCIndex & ".")
    
End Function

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 13/02/2011 (Amraphen)
    '13/02/2011: Amraphen - Ahora se quita la stamina en el sub UsuarioAtacaNPC.
    '***************************************************

    Dim index     As Integer

    Dim AttackPos As WorldPos
    Dim bot_Index As Byte
    
    Dim Aumenta     As Integer
    
    If UserList(UserIndex).flags.ModoCombate = False Then
        Call WriteConsoleMsg(UserIndex, "Para atacar debes activar el modo combate", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Check bow's interval
    If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
    
    'Check Spell-Magic interval
    If Not IntervaloPermiteMagiaGolpe(UserIndex) Then

        'Check Attack interval
        If Not IntervaloPermiteAtacar(UserIndex) Then
            Exit Sub

        End If

    End If
    
    With UserList(UserIndex)

        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).Peso > 0 Then _
                Aumenta = ObjData(.Invent.WeaponEqpObjIndex).Peso
        End If

        'Chequeamos que tenga por lo menos 10 de stamina.
        If .Stats.MinSta < 10 + Aumenta Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(UserIndex, "Estas muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Estas muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)

            End If

            Exit Sub

        End If
        
        AttackPos = .Pos
        Call HeadtoPos(.Char.Heading, AttackPos)
        
        'Exit if not legal
        If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Exit Sub

        End If
        
        index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex
        
        'USUARIOS
        '**************
        If index > 0 Then
            Call UsuarioAtacaUsuario(UserIndex, index)
            Call RestarEnergiaGolpe(UserIndex)
            Call WriteUpdateUserStats(UserIndex)
            Call WriteUpdateUserStats(index)
            Exit Sub

        End If
        
        index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NPCIndex
        
        'NPC
        '*************
        If index > 0 Then
            If Npclist(index).Attackable Then
                If Npclist(index).MaestroUser > 0 And MapInfo(Npclist(index).Pos.Map).Pk = False Then
                    Call WriteConsoleMsg(UserIndex, "No puedes atacar mascotas en zona segura.", FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub

                End If
                
                Call UsuarioAtacaNpc(UserIndex, index)
            Else
                Call WriteConsoleMsg(UserIndex, "No puedes atacar a este NPC.", FontTypeNames.FONTTYPE_WARNING)

            End If
            
            Call RestarEnergiaGolpe(UserIndex)
            Call WriteUpdateUserStats(UserIndex)
            
            Exit Sub

        End If
        
        index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).ObjInfo.ObjIndex
        
        'Dummys
        '***********
        If index > 0 Then
        
            If ObjData(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).ObjInfo.ObjIndex).OBJType = eOBJType.otDummy Then
                Call UsuarioAtacaDummy(UserIndex, index, AttackPos)
            
            End If
            
            Call RestarEnergiaGolpe(UserIndex)
            Call WriteUpdateUserStats(UserIndex)
            
            Exit Sub
        
        End If
        
        'Si llego aqui es un golpe al aire
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
        Call RestarEnergiaGolpe(UserIndex)
        Call WriteUpdateUserStats(UserIndex)
        
        If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
            
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1

    End With

End Sub

Private Sub RestarEnergiaGolpe(ByVal UserIndex As Integer)
'****************************************
'Autor: Lorwik
'Fecha: 11/03/2021
'Descripci�n: Resta la energia de un golpe con un arma o sin armas
'****************************************

    Dim Aumenta As Integer

    With UserList(UserIndex)
    
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).Peso > 0 Then _
                Aumenta = ObjData(.Invent.WeaponEqpObjIndex).Peso
        End If

        'Quitamos stamina
        Call QuitarSta(UserIndex, RandomNumber(1, 10) + Aumenta)
    
    End With
End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, _
                               ByVal VictimaIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 21/05/2010
    '21/05/2010: ZaMa - Evito division por cero.
    '***************************************************

    On Error GoTo errHandler

    Dim ProbRechazo            As Long

    Dim Rechazo                As Boolean

    Dim ProbExito              As Long

    Dim PoderAtaque            As Long

    Dim UserPoderEvasion       As Long

    Dim UserPoderEvasionEscudo As Long

    Dim Arma                   As Integer

    Dim SkillTacticas          As Long

    Dim SkillDefensa           As Long

    Dim ProbEvadir             As Long

    Dim Skill                  As eSkill
    
    With UserList(VictimaIndex)
    
        SkillTacticas = .Stats.UserSkills(eSkill.Tacticas)
        SkillDefensa = .Stats.UserSkills(eSkill.Defensa)
        
        Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
        
        'Calculamos el poder de evasion...
        UserPoderEvasion = PoderEvasion(VictimaIndex)
        
        If .Invent.EscudoEqpObjIndex > 0 Then
            UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
            UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo

        Else
            UserPoderEvasionEscudo = 0

        End If
        
        'Esta usando un arma ???
        If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(Arma).proyectil = 1 Then
                PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
                Skill = eSkill.Proyectiles
            Else
                PoderAtaque = PoderAtaqueArma(AtacanteIndex)
                Skill = eSkill.Armas

            End If

        Else
            PoderAtaque = PoderAtaqueMarciales(AtacanteIndex)
            Skill = eSkill.Marciales

        End If
        
        ' Chances are rounded
        ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtaque - UserPoderEvasion) * 0.4))
        
        ' Se reduce la evasion un 25%
        If .flags.Meditando Then
            ProbEvadir = (100 - ProbExito) * 0.75
            ProbExito = MinimoInt(90, 100 - ProbEvadir)

        End If
        
        UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
        
        ' el usuario esta usando un escudo ???
        If .Invent.EscudoEqpObjIndex > 0 Then

            'Fallo ???
            If Not UsuarioImpacto Then
                
                Dim SumaSkills As Integer
                
                ' Para evitar division por 0
                SumaSkills = MaximoInt(1, SkillDefensa + SkillTacticas)
                
                ' Chances are rounded
                ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / SumaSkills))
                Rechazo = (RandomNumber(1, 100) <= ProbRechazo)

                If Rechazo Then
                    'Se rechazo el ataque con el escudo
                    Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO(RandomNumber(1, 4)), .Pos.X, .Pos.Y))
                      
                    Call WriteMultiMessage(AtacanteIndex, eMessages.BlockedWithShieldother)
                    Call WriteMultiMessage(VictimaIndex, eMessages.BlockedWithShieldUser)
                    
                    Call SubirSkill(VictimaIndex, eSkill.Defensa, True)
                Else
                    Call SubirSkill(VictimaIndex, eSkill.Defensa, False)

                End If

            End If

        End If
        
        If Not UsuarioImpacto Then
            Call SubirSkill(AtacanteIndex, Skill, False)

        End If

    End With
    
    Exit Function
    
errHandler:

    Dim AtacanteNick As String

    Dim VictimaNick  As String
    
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
    
    Call LogError("Error en UsuarioImpacto. Error " & Err.Number & " : " & Err.description & " AtacanteIndex: " & AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)

End Function

Public Function UsuarioAtacaDummy(ByVal UserIndex As Integer, _
                                    ByVal DummyIndex As Integer, ByRef AttackPos As WorldPos) As Boolean
    '***************************************************
    'Author: Lorwik
    'Last Modification: 06/03/2021
    '***************************************************

    On Error GoTo errHandler
    
    Dim PoderAtt    As Long
    Dim PoderEva    As Long
    Dim ProbExito   As Long
    Dim Exito       As Boolean
    Dim Experiencia As Long
    Dim dano        As Integer
    Dim ExpaDar     As Long
    
    With UserList(UserIndex)
    
        If ObjData(MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).ObjInfo.ObjIndex).Subtipo = 1 Then
            If Abs(.Pos.X - AttackPos.X) > RANGO_VISION_X Or Abs(.Pos.Y - AttackPos.Y) > RANGO_VISION_Y Then
                Call WriteConsoleMsg(UserIndex, "Estas muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Function
            End If
                    
        Else
            'Comprobacion de distancia
            If Distancia(AttackPos, .Pos) > 2 Then
                Call WriteConsoleMsg(UserIndex, "�Estas demasiado lejos!", FontTypeNames.FONTTYPE_INFO)
                Exit Function
                    
            End If
                    
        End If
    
        PoderAtt = 50
        PoderEva = 50
        
        ' Chances are rounded
        ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtt - PoderEva) * 0.4))
        Exito = (RandomNumber(1, 100) <= ProbExito)

        '�Tuvo exito?
        If Exito Then
            
            Call CheckUserLevel(UserIndex)
            
            If .Familiar.Invocado = 1 Then
                Call CheckFamilyLevel(UserIndex)
            End If
            
            dano = CalcularDano(UserIndex)
            
            'Al menos que pegue por 1
            If dano < 1 Then dano = 1
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_DUMMY, .Pos.X, .Pos.Y))
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(dano, .Char.CharIndex, vbBlue, True))
            
            Call WriteMultiMessage(UserIndex, eMessages.UserHitNPC, dano)
            
            ExpaDar = 50 * ExpMultiplier 'Experiencia por golpe que va a dar un dummy
            
            '�Esta en party?
            If UserList(UserIndex).PartyIndex > 0 Then
                Call mdParty.ObtenerExito(UserIndex, ExpaDar, AttackPos.Map, AttackPos.X, AttackPos.Y)
                
            Else
                UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar
    
                If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP
                Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpaDar & " puntos de experiencia.", FontTypeNames.FONTTYPE_EXP)
    
            End If
        
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("'FALLAS'", .Char.CharIndex, vbBlue, True))
            
            Call WriteMultiMessage(UserIndex, eMessages.UserSwing)
            
        End If

    End With
    
    UsuarioAtacaDummy = True
    
    Exit Function
    
errHandler:
    Call LogError("Error en UsuarioAtacaUsuario. Error " & Err.Number & " : " & Err.description)

End Function

Public Function UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, _
                                    ByVal VictimaIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 06/04/2020
    '14/01/2010: ZaMa - Lo transformo en funcion, para que no se pierdan municiones al atacar targets
    '                    invalidos, y evitar un doble chequeo innecesario
    '06/04/2020: FrankoH298 - Si un usuario ataca a un npc lo desmonta.
    '***************************************************

    On Error GoTo errHandler

    If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Function
    
    With UserList(AtacanteIndex)
        
        If Abs(.Pos.X - UserList(VictimaIndex).Pos.X) > RANGO_VISION_X Or Abs(.Pos.Y - UserList(VictimaIndex).Pos.Y) > RANGO_VISION_Y Then
            Call WriteConsoleMsg(AtacanteIndex, "Estas muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function
        End If
        
        Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)
        
        If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            If UserList(VictimaIndex).flags.Navegando = 0 Then
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, FXSANGRE, 0))

            End If
            
            Call GolpeParaliza(AtacanteIndex, VictimaIndex)
            
            'Pablo (ToxicWaste): Guantes de Hurto del Bandido en accion
            If .clase = eClass.Bandit Then
                Call DoDesequipar(AtacanteIndex, VictimaIndex)

            End If
            
            Call SubirSkill(VictimaIndex, eSkill.Tacticas, False)
            Call UserDanoUser(AtacanteIndex, VictimaIndex)
        Else

            ' Invisible admins doesn't make sound to other clients except itself
            If .flags.AdminInvisible = 1 Then
                Call UserList(AtacanteIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))

            End If
            
            Call WriteMultiMessage(AtacanteIndex, eMessages.UserSwing)
            Call WriteMultiMessage(VictimaIndex, eMessages.UserAttackedSwing, AtacanteIndex)
            
            'Renderizo el dano en render.
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessageChatOverHead("'FALLAS'", .Char.CharIndex, vbBlue, True))
            
            Call SubirSkill(VictimaIndex, eSkill.Tacticas, True)

        End If
        
        If .clase = eClass.Thief Then Call Desarmar(AtacanteIndex, VictimaIndex)

    End With
    
    UsuarioAtacaUsuario = True
    
    Exit Function
    
errHandler:
    Call LogError("Error en UsuarioAtacaUsuario. Error " & Err.Number & " : " & Err.description)

End Function

Public Sub UserDanoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010 (ZaMa)
    '12/01/2010: ZaMa - Implemento armas arrojadizas y probabilidad de acuchillar.
    '11/03/2010: ZaMa - Ahora no cuenta la muerte si estaba en estado atacable, y no se vuelve criminal.
    '18/09/2010: ZaMa - Ahora se cosidera la defensa de los barcos siempre.
    '***************************************************
    
    On Error GoTo errHandler

    Dim dano          As Long

    Dim Lugar         As Byte

    Dim obj           As ObjData
    
    Dim BoatDefense   As Integer

    Dim BodyDefense   As Integer

    Dim HeadDefense   As Integer

    Dim WeaponBoost   As Integer
    
    Dim BoatIndex     As Integer

    Dim WeaponIndex   As Integer

    Dim HelmetIndex   As Integer

    Dim ArmourIndex   As Integer

    Dim ShieldIndex   As Integer
    
    Dim BarcaIndex    As Integer

    Dim ArmaIndex     As Integer

    Dim CascoIndex    As Integer

    Dim ArmaduraIndex As Integer
    
    dano = CalcularDano(AtacanteIndex)
    
    'Al menos que pegue por 1
    If dano < 1 Then dano = 1
    
    Call UserEnvenena(AtacanteIndex, VictimaIndex)
    
    Call UserIncinera(AtacanteIndex, VictimaIndex)
    
    With UserList(AtacanteIndex)
        
        ' Aumento de dano por barca (atacante)
        If .flags.Navegando = 1 Then
            
            BoatIndex = .Invent.BarcoObjIndex
            
            If BoatIndex > 0 Then
                obj = ObjData(BoatIndex)
                dano = dano + RandomNumber(obj.MinHIT, obj.MaxHit)

            End If
            
        End If
        
        ' Aumento de defensa por barca (victima)
        If UserList(VictimaIndex).flags.Navegando = 1 Then
            
            BoatIndex = UserList(VictimaIndex).Invent.BarcoObjIndex
            
            If BoatIndex > 0 Then
                obj = ObjData(BoatIndex)
                BoatDefense = RandomNumber(obj.MinDef, obj.MaxDef)

            End If
            
        End If
        
        ' Refuerzo arma (atacante)
        WeaponIndex = .Invent.WeaponEqpObjIndex

        If WeaponIndex > 0 Then
            WeaponBoost = ObjData(WeaponIndex).Refuerzo

        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
        
            Case PartesCuerpo.bCabeza
            
                'Si tiene casco absorbe el golpe
                HelmetIndex = UserList(VictimaIndex).Invent.CascoEqpObjIndex

                If HelmetIndex > 0 Then
                    obj = ObjData(HelmetIndex)
                    HeadDefense = RandomNumber(obj.MinDef, obj.MaxDef)

                End If
            
            Case Else
                
                Dim MinDef As Integer

                Dim MaxDef As Integer
                
                'Si tiene armadura absorbe el golpe
                ArmourIndex = UserList(VictimaIndex).Invent.ArmourEqpObjIndex

                If ArmourIndex > 0 Then
                    obj = ObjData(ArmourIndex)
                    MinDef = obj.MinDef
                    MaxDef = obj.MaxDef

                End If
                
                ' Si tiene escudo, tambien absorbe el golpe
                ShieldIndex = UserList(VictimaIndex).Invent.EscudoEqpObjIndex

                If ShieldIndex > 0 Then
                    obj = ObjData(ShieldIndex)
                    MinDef = MinDef + obj.MinDef
                    MaxDef = MaxDef + obj.MaxDef

                End If
                
                BodyDefense = RandomNumber(MinDef, MaxDef)
                
        End Select
        
        dano = dano + WeaponBoost - HeadDefense - BodyDefense - BoatDefense

        If dano < 0 Then dano = 1
        
        Call WriteMultiMessage(AtacanteIndex, eMessages.UserHittedUser, UserList(VictimaIndex).Char.CharIndex, Lugar, dano)
        Call WriteMultiMessage(VictimaIndex, eMessages.UserHittedByUser, .Char.CharIndex, Lugar, dano)
        
        UserList(VictimaIndex).Stats.MinHp = UserList(VictimaIndex).Stats.MinHp - dano
        
        If .flags.Hambre = 0 And .flags.Sed = 0 Then

            'Si usa un arma quizas suba "Combate con armas"
            If WeaponIndex > 0 Then
                If ObjData(WeaponIndex).proyectil Then
                    'es un Arco. Sube Armas a Distancia
                    Call SubirSkill(AtacanteIndex, eSkill.Proyectiles, True)
                    
                    ' Si acuchilla
                    If PuedeAcuchillar(AtacanteIndex) Then
                        Call DoAcuchillar(AtacanteIndex, 0, VictimaIndex, dano)

                    End If

                Else
                    'Sube combate con armas.
                    Call SubirSkill(AtacanteIndex, eSkill.Armas, True)

                End If

            Else
                'sino tal vez lucha libre
                Call SubirSkill(AtacanteIndex, eSkill.Marciales, True)

            End If
                    
            'Trata de apunalar por la espalda al enemigo
            If PuedeApunalar(AtacanteIndex) Then
                Call DoApunalar(AtacanteIndex, 0, VictimaIndex, dano)

            End If

            'e intenta dar un golpe critico [Pablo (ToxicWaste)]
            Call DoGolpeCritico(AtacanteIndex, 0, VictimaIndex, dano)

        End If
        
        'Doble chekeo innecesario, pero bueno..
        'Hecho para que no envie apu + golpe normal.
        If Not PuedeApunalar(AtacanteIndex) Then _
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessageChatOverHead(dano, .Char.CharIndex, vbBlue, True))
        
        If UserList(VictimaIndex).Stats.MinHp <= 0 Then
            
            ' No cuenta la muerte si estaba en estado atacable
            If UserList(VictimaIndex).flags.AtacablePor <> AtacanteIndex Then
                'Store it!
                Call Statistics.StoreFrag(AtacanteIndex, VictimaIndex)
                
                Call ContarMuerte(VictimaIndex, AtacanteIndex)

            End If
            
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            Dim j As Integer

            For j = 1 To MAXMASCOTAS

                If .MascotasIndex(j) > 0 Then
                    If Npclist(.MascotasIndex(j)).Target = VictimaIndex Then
                        Npclist(.MascotasIndex(j)).Target = 0
                        Call FollowAmo(.MascotasIndex(j))

                    End If

                End If

            Next j
            
            Call ActStats(VictimaIndex, AtacanteIndex)
            Call UserDie(VictimaIndex, AtacanteIndex)
        Else
            'Esta vivo - Actualizamos el HP
            Call WriteUpdateHP(VictimaIndex)

        End If

        'Controla el nivel del usuario
        Call CheckUserLevel(AtacanteIndex)
        
        If .Familiar.Invocado = 1 Then
            Call CheckFamilyLevel(AtacanteIndex)
        End If
    
    End With
    
    Exit Sub
    
errHandler:

    Dim AtacanteNick As String

    Dim VictimaNick  As String
    
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
    
    Call LogError("Error en UserDanoUser. Error " & Err.Number & " : " & Err.description & " AtacanteIndex: " & AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)

End Sub

Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
    '***************************************************
    'Autor: Unknown
    'Last Modification: 05/05/2010
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    '10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
    '05/05/2010: ZaMa - Ahora no suma puntos de bandido al atacar a alguien en estado atacable.
    '***************************************************

    If TriggerZonaPelea(AttackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
    Dim EraCriminal       As Boolean

    Dim VictimaEsAtacable As Boolean
    
    If Not criminal(AttackerIndex) Then
        If Not criminal(VictimIndex) Then
            ' Si la victima no es atacable por el agresor, entonces se hace pk
            VictimaEsAtacable = UserList(VictimIndex).flags.AtacablePor = AttackerIndex

            If Not VictimaEsAtacable Then Call VolverCriminal(AttackerIndex)

        End If

    End If
    
    With UserList(VictimIndex)

        If .flags.Meditando Then
            .flags.Meditando = False
            Call WriteMeditateToggle(VictimIndex)
            Call WriteConsoleMsg(VictimIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
            .Char.Particle = 0
            Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageCreateParticleChar(.Char.CharIndex, .Char.Particle, False, 0))

        End If

    End With
    
    EraCriminal = criminal(AttackerIndex)
    
    ' Si ataco a un atacable, no suma puntos de bandido
    If Not VictimaEsAtacable Then

        With UserList(AttackerIndex).Reputacion

            If Not criminal(VictimIndex) Then
                .BandidoRep = .BandidoRep + vlASALTO

                If .BandidoRep > MAXREP Then .BandidoRep = MAXREP
                
                .NobleRep = .NobleRep * 0.5

                If .NobleRep < 0 Then .NobleRep = 0
            Else
                .NobleRep = .NobleRep + vlNoble

                If .NobleRep > MAXREP Then .NobleRep = MAXREP

            End If

        End With

    End If
    
    If criminal(AttackerIndex) Then
        If UserList(AttackerIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(AttackerIndex)
        
        If Not EraCriminal Then Call RefreshCharStatus(AttackerIndex)
    ElseIf EraCriminal Then
        Call RefreshCharStatus(AttackerIndex)

    End If
    
    Call FamiliarAtacaUsuario(AttackerIndex, VictimIndex)
    
    Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)
    
    'Si la victima esta saliendo se cancela la salida
    Call CancelExit(VictimIndex)
    
    'Si esta casteando, lo cancelamos
    Call CancelCast(VictimIndex)

End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    'Reaccion de las mascotas
    Dim iCount As Integer
    
    For iCount = 1 To MAXMASCOTAS

        If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(victim).Name
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1

        End If

    Next iCount

End Sub

Public Function PuedeAtacar(ByVal AttackerIndex As Integer, _
                            ByVal VictimIndex As Integer) As Boolean

    '***************************************************
    'Autor: Unknown
    'Last Modification: 02/04/2010
    'Returns true if the AttackerIndex is allowed to attack the VictimIndex.
    '24/01/2007 Pablo (ToxicWaste) - Ordeno todo y agrego situacion de Defensa en ciudad Armada y Caos.
    '24/02/2009: ZaMa - Los usuarios pueden atacarse entre si.
    '02/04/2010: ZaMa - Los armadas no pueden atacar nunca a los ciudas, salvo que esten atacables.
    '***************************************************
    On Error GoTo errHandler

    'MUY importante el orden de estos "IF"...
    
    'Estas muerto no podes atacar
    If UserList(AttackerIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function

    End If
    
    'No podes atacar a alguien muerto
    If UserList(VictimIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a un espiritu.", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function

    End If
    
    '�Est� trabajando?
    If UserList(AttackerIndex).flags.MacroTrabajo <> 0 Then
        Call WriteConsoleMsg(AttackerIndex, "�Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    
    ' No podes atacar si estas en consulta
    If UserList(AttackerIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
        Exit Function

    End If
    
    ' No podes atacar si esta en consulta
    If UserList(VictimIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estan en consulta.", FontTypeNames.FONTTYPE_INFO)
        Exit Function

    End If
    
    '�Estan en party?
    If UserList(AttackerIndex).PartyIndex > 0 Then
        If UserList(AttackerIndex).PartyIndex = UserList(VictimIndex).PartyIndex Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a miembros de tu mismo grupo.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
    
        End If
    End If

    'Estamos en una Arena? o un trigger zona segura?
    Select Case TriggerZonaPelea(AttackerIndex, VictimIndex)

        Case eTrigger6.TRIGGER6_PERMITE
            PuedeAtacar = (UserList(VictimIndex).flags.AdminInvisible = 0)
            Exit Function
        
        Case eTrigger6.TRIGGER6_PROHIBE
            PuedeAtacar = False
            Exit Function
        
        Case eTrigger6.TRIGGER6_AUSENTE

            'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
            If (UserList(VictimIndex).flags.Privilegios And PlayerType.User) = 0 Then
                If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(AttackerIndex, "El ser es demasiado poderoso.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False
                Exit Function

            End If

    End Select
    
    'Ataca un ciudadano?
    If Not criminal(VictimIndex) Then

        ' El atacante es ciuda?
        If Not criminal(AttackerIndex) Then

            ' El atacante es armada?
            If esArmada(AttackerIndex) Then

                ' La victima es armada?
                If esArmada(VictimIndex) Then
                    ' No puede
                    Call WriteConsoleMsg(AttackerIndex, "Los soldados del ejercito real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
                    Exit Function

                End If

            End If
            
            ' Ciuda (o army) atacando a otro ciuda (o army)
            If UserList(VictimIndex).flags.AtacablePor = AttackerIndex Then

                ' Se vuelve atacable.
                If ToogleToAtackable(AttackerIndex, VictimIndex, False) Then
                    PuedeAtacar = True
                    Exit Function

                End If

            End If

        End If

        ' Ataca a un criminal
    Else

        'Sos un Caos atacando otro caos?
        If esCaos(VictimIndex) Then
            If esCaos(AttackerIndex) Then
                Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legion oscura tienen prohibido atacarse entre si.", FontTypeNames.FONTTYPE_WARNING)
                Exit Function

            End If

        End If

    End If
    
    'Tenes puesto el seguro?
    If UserList(AttackerIndex).flags.Seguro Then
        If Not criminal(VictimIndex) Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar ciudadanos, para hacerlo debes desactivar el seguro.", FontTypeNames.FONTTYPE_WARNING)
            PuedeAtacar = False
            Exit Function

        End If

    Else

        ' Un ciuda es atacado
        If Not criminal(VictimIndex) Then

            ' Por un armada sin seguro
            If esArmada(AttackerIndex) Then
                ' No puede
                Call WriteConsoleMsg(AttackerIndex, "Los soldados del ejercito real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False
                Exit Function

            End If

        End If

    End If
    
    'Estas en un Mapa Seguro?
    If MapInfo(UserList(VictimIndex).Pos.Map).Pk = False Then
        If esArmada(AttackerIndex) Then
            If UserList(AttackerIndex).Faccion.RecompensasReal > 11 Then
                If UserList(VictimIndex).Pos.Map = 58 Or UserList(VictimIndex).Pos.Map = 59 Or UserList(VictimIndex).Pos.Map = 60 Then
                    Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! Estas siendo atacado y no podras defenderte.", FontTypeNames.FONTTYPE_WARNING)
                    PuedeAtacar = True 'Beneficio de Armadas que atacan en su ciudad.
                    Exit Function

                End If

            End If

        End If

        If esCaos(AttackerIndex) Then
            If UserList(AttackerIndex).Faccion.RecompensasCaos > 11 Then
                If UserList(VictimIndex).Pos.Map = 151 Or UserList(VictimIndex).Pos.Map = 156 Then
                    Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! Estas siendo atacado y no podras defenderte.", FontTypeNames.FONTTYPE_WARNING)
                    PuedeAtacar = True 'Beneficio de Caos que atacan en su ciudad.
                    Exit Function

                End If

            End If

        End If

        Call WriteConsoleMsg(AttackerIndex, "Esta es una zona segura, aqui no puedes atacar a otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function

    End If
    
    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
    If MapData(UserList(VictimIndex).Pos.Map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).Trigger = eTrigger.ZONASEGURA Or MapData(UserList(AttackerIndex).Pos.Map, UserList(AttackerIndex).Pos.X, UserList(AttackerIndex).Pos.Y).Trigger = eTrigger.ZONASEGURA Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes pelear aqui.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function

    End If
    
    PuedeAtacar = True
    Exit Function

errHandler:
    Call LogError("Error en PuedeAtacar. Error " & Err.Number & " : " & Err.description)

End Function

Public Function PuedeAtacarNPC(ByVal AttackerIndex As Integer, _
                               ByVal NPCIndex As Integer, _
                               Optional ByVal Paraliza As Boolean = False, _
                               Optional ByVal IsPet As Boolean = False) As Boolean
    '***************************************************
    'Autor: Unknown Author (Original version)
    'Returns True if AttackerIndex can attack the NpcIndex
    'Last Modification: 04/07/2010
    '24/01/2007 Pablo (ToxicWaste) - Orden y correccion de ataque sobre una mascota y guardias
    '14/08/2007 Pablo (ToxicWaste) - Reescribo y agrego TODOS los casos posibles cosa de usar
    'esta funcion para todo lo referente a ataque a un NPC. Ya sea Magia, Fisico o a Distancia.
    '16/11/2009: ZaMa - Agrego validacion de pertenencia de npc.
    '02/04/2010: ZaMa - Los armadas ya no peuden atacar npcs no hotiles.
    '23/05/2010: ZaMa - El inmo/para renuevan el timer de pertenencia si el ataque fue a un npc propio.
    '04/07/2010: ZaMa - Ahora no se puede apropiar del dragon de dd.
    '***************************************************

    On Error GoTo errHandler

    With Npclist(NPCIndex)
    
        'Estas muerto?
        If UserList(AttackerIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(AttackerIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Function

        End If
        
        'Sos consejero?
        If UserList(AttackerIndex).flags.Privilegios And PlayerType.Consejero Then
            'No pueden atacar NPC los Consejeros.
            Exit Function

        End If
        
        ' No podes atacar si estas en consulta
        If UserList(AttackerIndex).flags.EnConsulta Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar npcs mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Function

        End If
        
        'Es una criatura atacable?
        If .Attackable = 0 Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar esta criatura.", FontTypeNames.FONTTYPE_INFO)
            Exit Function

        End If
        
        '�Est� trabajando?
        If UserList(AttackerIndex).flags.MacroTrabajo <> 0 Then
            Call WriteConsoleMsg(AttackerIndex, "�Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        'Es valida la distancia a la cual estamos atacando?
        If Abs(.Pos.X - UserList(AttackerIndex).Pos.X) > RANGO_VISION_X Or Abs(.Pos.Y - UserList(AttackerIndex).Pos.Y) > RANGO_VISION_Y Then
            Call WriteConsoleMsg(AttackerIndex, "Estas muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function

        End If

        '�Intenta atacar a su propio familiar?
        If UserList(AttackerIndex).Familiar.Id = NPCIndex Then
            Call WriteConsoleMsg(AttackerIndex, "�No puedes atacar a tu familiar!", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function

        End If
        
        'Es una criatura No-Hostil?
        If .Hostile = 0 Then

            'Es Guardia del Caos?
            If .NPCtype = eNPCType.Guardiascaos Then

                'Lo quiere atacar un caos?
                If esCaos(AttackerIndex) Then
                    Call WriteConsoleMsg(AttackerIndex, "No puedes atacar Guardias del Caos siendo de la legion oscura.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function

                End If

                'Es guardia Real?
            ElseIf .NPCtype = eNPCType.GuardiaReal Then

                'Lo quiere atacar un Armada?
                If esArmada(AttackerIndex) Then
                    Call WriteConsoleMsg(AttackerIndex, "No puedes atacar Guardias Reales siendo del ejercito real.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function

                End If

                'Tienes el seguro puesto?
                If UserList(AttackerIndex).flags.Seguro Then
                    Call WriteConsoleMsg(AttackerIndex, "Para poder atacar Guardias Reales debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                Else
                    Call WriteConsoleMsg(AttackerIndex, "Atacaste un Guardia Real! Eres un criminal.", FontTypeNames.FONTTYPE_INFO)
                    Call VolverCriminal(AttackerIndex)
                    PuedeAtacarNPC = True
                    Exit Function

                End If
        
                'No era un Guardia, asi que es una criatura No-Hostil comun.
                'Para asegurarnos que no sea una Mascota:
            ElseIf .MaestroUser = 0 Then

                'Si sos ciudadano tenes que quitar el seguro para atacarla.
                If Not criminal(AttackerIndex) Then
                    
                    ' Si sos armada no podes atacarlo directamente
                    If esArmada(AttackerIndex) Then
                        Call WriteConsoleMsg(AttackerIndex, "Los miembros del ejercito real no pueden atacar npcs no hostiles.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function

                    End If
                
                    'Sos ciudadano, tenes el seguro puesto?
                    If UserList(AttackerIndex).flags.Seguro Then
                        Call WriteConsoleMsg(AttackerIndex, "Para atacar a este NPC debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    Else
                        'No tiene seguro puesto. Puede atacar pero es penalizado.
                        Call WriteConsoleMsg(AttackerIndex, "Atacaste un NPC no-hostil. Continua haciendolo y te podras convertir en criminal.", FontTypeNames.FONTTYPE_INFO)
                        'NicoNZ: Cambio para que al atacar npcs no hostiles no bajen puntos de nobleza
                        Call DisNobAuBan(AttackerIndex, 0, 1000)
                        PuedeAtacarNPC = True
                        Exit Function

                    End If

                End If

            End If

        End If
    
        Dim MasterIndex As Integer

        MasterIndex = .MaestroUser
        
        'Es el NPC mascota de alguien?
        If MasterIndex > 0 Then
            
            ' Dueno de la mascota ciuda?
            If Not criminal(MasterIndex) Then
                
                ' Atacante ciuda?
                If Not criminal(AttackerIndex) Then
                    
                    ' Si esta en estado atacable puede atacar su mascota sin problemas
                    If UserList(MasterIndex).flags.AtacablePor = AttackerIndex Then
                        ' Toogle to atacable and restart the timer
                        Call ToogleToAtackable(AttackerIndex, MasterIndex)
                        PuedeAtacarNPC = True
                        Exit Function

                    End If
                    
                    'Atacante armada?
                    If esArmada(AttackerIndex) Then
                        'El atacante es Armada y esta intentando atacar mascota de un Ciudadano
                        Call WriteConsoleMsg(AttackerIndex, "Los miembros del ejercito real no pueden atacar mascotas de ciudadanos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function

                    End If
                    
                    'El atacante es Ciudadano y esta intentando atacar mascota de un Ciudadano.
                    If UserList(AttackerIndex).flags.Seguro Then
                        'El atacante tiene el seguro puesto. No puede atacar.
                        Call WriteConsoleMsg(AttackerIndex, "Para atacar mascotas de ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    Else
                        'El atacante no tiene el seguro puesto. Recibe penalizacion.
                        Call WriteConsoleMsg(AttackerIndex, "Has atacado la Mascota de un ciudadano. Eres un criminal.", FontTypeNames.FONTTYPE_INFO)
                        Call VolverCriminal(AttackerIndex)
                        PuedeAtacarNPC = True
                        Exit Function

                    End If

                Else

                    'El atacante es criminal y quiere atacar un elemental ciuda, pero tiene el seguro puesto (NicoNZ)
                    If UserList(AttackerIndex).flags.Seguro Then
                        Call WriteConsoleMsg(AttackerIndex, "Para atacar mascotas de ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function

                    End If

                End If
            
                'Es mascota de un caos?
            ElseIf esCaos(MasterIndex) Then

                'Es Caos el Dueno.
                If esCaos(AttackerIndex) Then
                    'Un Caos intenta atacar una criatura de un Caos. No puede atacar.
                    Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legion oscura no pueden atacar mascotas de otros legionarios. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Function

                End If

            End If
            
            ' No es mascota de nadie, le pertenece a alguien?
        ElseIf .Owner > 0 Then
        
            Dim OwnerUserIndex As Integer

            OwnerUserIndex = .Owner
            
            ' Puede atacar a su propia criatura!
            If OwnerUserIndex = AttackerIndex Then
                PuedeAtacarNPC = True
                Call IntervaloPerdioNpc(OwnerUserIndex, True) ' Renuevo el timer
                Exit Function

            End If
            
            ' Esta compartiendo el npc con el atacante? => Puede atacar!
            If UserList(OwnerUserIndex).flags.ShareNpcWith = AttackerIndex Then
                PuedeAtacarNPC = True
                Exit Function

            End If
            
            ' Si son del mismo clan o party, pueden atacar (No renueva el timer)
            If Not SameClan(OwnerUserIndex, AttackerIndex) And Not SameParty(OwnerUserIndex, AttackerIndex) Then
            
                ' Si se le agoto el tiempo
                If IntervaloPerdioNpc(OwnerUserIndex) Then ' Se lo roba :P
                    Call PerdioNpc(OwnerUserIndex)
                    Call ApropioNpc(AttackerIndex, NPCIndex)
                    PuedeAtacarNPC = True
                    Exit Function
                    
                    ' Si lanzo un hechizo de para o inmo
                ElseIf Paraliza Then
                
                    ' Si ya esta paralizado o inmobilizado, no puedo inmobilizarlo de nuevo
                    If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then
                        
                        'TODO_ZAMA: Si dejo esto asi, los pks con seguro peusto van a poder inmobilizar criaturas con dueno
                        ' Si es pk neutral, puede hacer lo que quiera :P.
                        If Not criminal(AttackerIndex) And Not criminal(OwnerUserIndex) Then
                        
                            'El atacante es Armada
                            If esArmada(AttackerIndex) Then
                                
                                'Intententa paralizar un npc de un armada?
                                If esArmada(OwnerUserIndex) Then
                                    'El atacante es Armada y esta intentando paralizar un npc de un armada: No puede
                                    Call WriteConsoleMsg(AttackerIndex, "Los miembros del Ejercito Real no pueden paralizar criaturas ya paralizadas pertenecientes a otros miembros del Ejercito Real", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function
                                
                                    'El atacante es Armada y esta intentando paralizar un npc de un ciuda
                                Else

                                    ' Si tiene seguro no puede
                                    If UserList(AttackerIndex).flags.Seguro Then
                                        Call WriteConsoleMsg(AttackerIndex, "Para paralizar criaturas ya paralizadas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                                        Exit Function
                                    Else

                                        ' Si ya estaba atacable, no podra atacar a un npc perteneciente a otro ciuda
                                        If ToogleToAtackable(AttackerIndex, OwnerUserIndex) Then
                                            Call WriteConsoleMsg(AttackerIndex, "Has paralizado la criatura de un ciudadano, ahora eres atacable por el.", FontTypeNames.FONTTYPE_INFO)
                                            PuedeAtacarNPC = True

                                        End If
                                        
                                        Exit Function
                                        
                                    End If

                                End If
                                
                                ' El atacante es ciuda
                            Else

                                'El atacante tiene el seguro puesto, no puede paralizar
                                If UserList(AttackerIndex).flags.Seguro Then
                                    Call WriteConsoleMsg(AttackerIndex, "Para paralizar criaturas ya paralizadas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function
                                    
                                    'El atacante no tiene el seguro puesto, ataca.
                                Else

                                    ' Si ya estaba atacable, no podra atacar a un npc perteneciente a otro ciuda
                                    If ToogleToAtackable(AttackerIndex, OwnerUserIndex) Then
                                        Call WriteConsoleMsg(AttackerIndex, "Has paralizado la criatura de un ciudadano, ahora eres atacable por el.", FontTypeNames.FONTTYPE_INFO)
                                        PuedeAtacarNPC = True

                                    End If
                                    
                                    Exit Function

                                End If

                            End If
                            
                            ' Al menos uno de los dos es criminal
                        Else

                            ' Si ambos son caos
                            If esCaos(AttackerIndex) And esCaos(OwnerUserIndex) Then
                                'El atacante es Caos y esta intentando paralizar un npc de un Caos
                                Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legion oscura no pueden paralizar criaturas ya paralizadas por otros legionarios.", FontTypeNames.FONTTYPE_INFO)
                                Exit Function

                            End If

                        End If
                    
                        ' El npc no esta inmobilizado ni paralizado
                    Else

                        ' Si no tiene dueno, puede apropiarselo
                        If OwnerUserIndex = 0 Then
                        
                            ' Siempre que no posea uno ya (el inmo/para no cambia pertenencia de npcs).
                            If UserList(AttackerIndex).flags.OwnedNpc = 0 Then
                                Call ApropioNpc(AttackerIndex, NPCIndex)

                            End If
                            
                            ' Si inmobiliza a su propio npc, renueva el timer
                        ElseIf OwnerUserIndex = AttackerIndex Then
                            Call IntervaloPerdioNpc(OwnerUserIndex, True) ' Renuevo el timer

                        End If
                        
                        ' Siempre se pueden paralizar/inmobilizar npcs con o sin dueno
                        ' que no tengan ese estado
                        PuedeAtacarNPC = True
                        Exit Function

                    End If
                    
                    ' No lanzo hechizos inmobilizantes
                Else
                    
                    ' El npc le pertenece a un ciudadano
                    If Not criminal(OwnerUserIndex) Then
                        
                        'El atacante es Armada y esta intentando atacar un npc de un Ciudadano
                        If esArmada(AttackerIndex) Then
                        
                            'Intententa atacar un npc de un armada?
                            If esArmada(OwnerUserIndex) Then
                                'El atacante es Armada y esta intentando atacar el npc de un armada: No puede
                                Call WriteConsoleMsg(AttackerIndex, "Los miembros del Ejercito Real no pueden atacar criaturas pertenecientes a otros miembros del Ejercito Real", FontTypeNames.FONTTYPE_INFO)
                                Exit Function
                            
                                'El atacante es Armada y esta intentando atacar un npc de un ciuda
                            Else
                                
                                ' Si tiene seguro no puede
                                If UserList(AttackerIndex).flags.Seguro Then
                                    Call WriteConsoleMsg(AttackerIndex, "Para atacar criaturas ya pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function
                                Else

                                    ' Si ya estaba atacable, no podra atacar a un npc perteneciente a otro ciuda
                                    If ToogleToAtackable(AttackerIndex, OwnerUserIndex) Then
                                        Call WriteConsoleMsg(AttackerIndex, "Has atacado a la criatura de un ciudadano, ahora eres atacable por el.", FontTypeNames.FONTTYPE_INFO)
                                        PuedeAtacarNPC = True

                                    End If
                                    
                                    Exit Function

                                End If

                            End If
                            
                            ' No es aramda, puede ser criminal o ciuda
                        Else
                            
                            'El atacante es Ciudadano y esta intentando atacar un npc de un Ciudadano.
                            If Not criminal(AttackerIndex) Then
                                
                                If UserList(AttackerIndex).flags.Seguro Then
                                    'El atacante tiene el seguro puesto. No puede atacar.
                                    Call WriteConsoleMsg(AttackerIndex, "Para atacar criaturas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function
                                
                                    'El atacante no tiene el seguro puesto, ataca.
                                Else

                                    If ToogleToAtackable(AttackerIndex, OwnerUserIndex) Then
                                        Call WriteConsoleMsg(AttackerIndex, "Has atacado a la criatura de un ciudadano, ahora eres atacable por el.", FontTypeNames.FONTTYPE_INFO)
                                        PuedeAtacarNPC = True

                                    End If
                                    
                                    Exit Function

                                End If
                                
                                'El atacante es criminal y esta intentando atacar un npc de un Ciudadano.
                            Else

                                ' Es criminal atacando un npc de un ciuda, con seguro puesto.
                                If UserList(AttackerIndex).flags.Seguro Then
                                    Call WriteConsoleMsg(AttackerIndex, "Para atacar criaturas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Function

                                End If
                                
                                PuedeAtacarNPC = True

                            End If

                        End If
                        
                        ' Es npc de un criminal
                    Else

                        If esCaos(OwnerUserIndex) Then

                            'Es Caos el Dueno.
                            If esCaos(AttackerIndex) Then
                                'Un Caos intenta atacar una npc de un Caos. No puede atacar.
                                Call WriteConsoleMsg(AttackerIndex, "Los miembros de la Legion Oscura no pueden atacar criaturas de otros legionarios. ", FontTypeNames.FONTTYPE_INFO)
                                Exit Function

                            End If

                        End If

                    End If

                End If

            End If
            
            ' Si no tiene dueno el npc, se lo apropia
        Else

            ' Solo pueden apropiarse de npcs los caos, armadas o ciudas.
            If Not criminal(AttackerIndex) Or esCaos(AttackerIndex) Then

                ' No puede apropiarse del dragon de dd!
                If Npclist(NPCIndex).NPCtype <> DRAGON Then

                    ' Si es una mascota atacando, no se apropia del npc
                    If Not IsPet Then

                        ' No es dueno de ningun npc => Se lo apropia.
                        If UserList(AttackerIndex).flags.OwnedNpc = 0 Then
                            Call ApropioNpc(AttackerIndex, NPCIndex)
                            ' Es dueno de un npc, pero no puede ser de este porque no tiene propietario.
                        Else

                            ' Se va a aduenar del npc (y perder el otro) solo si no inmobiliza/paraliza
                            If Not Paraliza Then Call ApropioNpc(AttackerIndex, NPCIndex)

                        End If

                    End If

                End If

            End If

        End If

    End With
    
    PuedeAtacarNPC = True
        
    Exit Function
        
errHandler:
    
    Dim AtckName  As String

    Dim OwnerName As String

    If AttackerIndex > 0 Then AtckName = UserList(AttackerIndex).Name
    If OwnerUserIndex > 0 Then OwnerName = UserList(OwnerUserIndex).Name
    
    Call LogError("Error en PuedeAtacarNpc. Erorr: " & Err.Number & " - " & Err.description & " Atacante: " & AttackerIndex & "-> " & AtckName & ". Owner: " & OwnerUserIndex & "-> " & OwnerName & ". NpcIndex: " & NPCIndex & ".")

End Function

Private Function SameClan(ByVal UserIndex As Integer, _
                          ByVal OtherUserIndex As Integer) As Boolean
    '***************************************************
    'Autor: ZaMa
    'Returns True if both players belong to the same clan.
    'Last Modification: 16/11/2009
    '***************************************************
    SameClan = (UserList(UserIndex).GuildIndex = UserList(OtherUserIndex).GuildIndex) And UserList(UserIndex).GuildIndex <> 0

End Function

Private Function SameParty(ByVal UserIndex As Integer, _
                           ByVal OtherUserIndex As Integer) As Boolean
    '***************************************************
    'Autor: ZaMa
    'Returns True if both players belong to the same party.
    'Last Modification: 16/11/2009
    '***************************************************
    SameParty = UserList(UserIndex).PartyIndex = UserList(OtherUserIndex).PartyIndex And UserList(UserIndex).PartyIndex <> 0

End Function

Sub CalcularDarExp(ByVal UserIndex As Integer, _
                   ByVal NPCIndex As Integer, _
                   ByVal ElDano As Long)

    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/09/06 Nacho
    'Reescribi gran parte del Sub
    'Ahora, da toda la experiencia del npc mientras este vivo.
    '***************************************************
    Dim ExpaDar As Long
Debug.Print "ElDano: " & ElDano
    With UserList(UserIndex)
    
        '[Nacho] Chekeamos que las variables sean validas para las operaciones
        If ElDano <= 0 Then ElDano = 0
        If Npclist(NPCIndex).Stats.MaxHp <= 0 Then Exit Sub
        If ElDano > Npclist(NPCIndex).Stats.MinHp Then ElDano = Npclist(NPCIndex).Stats.MinHp
        
        Debug.Print ElDano & "*(" & Npclist(NPCIndex).GiveEXP & " / " & Npclist(NPCIndex).Stats.MaxHp & ")"
        '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
        ExpaDar = CLng(ElDano * (Npclist(NPCIndex).GiveEXP / Npclist(NPCIndex).Stats.MaxHp))

        If ExpaDar <= 0 Then Exit Sub
        
        '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
        'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
        'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
        If ExpaDar > Npclist(NPCIndex).flags.ExpCount Then
            ExpaDar = Npclist(NPCIndex).flags.ExpCount
            Npclist(NPCIndex).flags.ExpCount = 0
        Else
            Npclist(NPCIndex).flags.ExpCount = Npclist(NPCIndex).flags.ExpCount - ExpaDar
    
        End If
        
        'Si el familiar esta invocado, primero hacemos la reparticion
        If .Familiar.Invocado = 1 Then
        
            'Dividimos
            ExpaDar = ExpaDar / 2
            
            'Si la repartici�n quedo por debajo de 1, le damos 1 puntito a cada uno
            If ExpaDar < 1 Then ExpaDar = 1
            
            .Familiar.Exp = .Familiar.Exp + ExpaDar
            
            If .Familiar.Exp > MAXEXP Then .Familiar.Exp = MAXEXP
            Call WriteConsoleMsg(UserIndex, "Tu familiar ha ganado " & ExpaDar & " puntos de experiencia.", FontTypeNames.FONTTYPE_EXP)
            
            Call CheckFamilyLevel(UserIndex)
        
        End If
        
        '[Nacho] Le damos la exp al user
        If ExpaDar > 0 Then
            If UserList(UserIndex).PartyIndex > 0 Then
                Call mdParty.ObtenerExito(UserIndex, ExpaDar, Npclist(NPCIndex).Pos.Map, Npclist(NPCIndex).Pos.X, Npclist(NPCIndex).Pos.Y)
            Else
            
                UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar
    
                If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP
                Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpaDar & " puntos de experiencia.", FontTypeNames.FONTTYPE_EXP)
    
            End If
            
            Call CheckUserLevel(UserIndex)
    
        End If
    
    End With

End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, _
                                 ByVal Destino As Integer) As eTrigger6
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'TODO: Pero que rebuscado!!
    'Nigo:  Te lo redisene, pero no te borro el TODO para que lo revises.
    On Error GoTo errHandler

    Dim tOrg As eTrigger

    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(Origen).Pos.Map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).Trigger
    tDst = MapData(UserList(Destino).Pos.Map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).Trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE

        End If

    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE

    End If

    Exit Function
errHandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.description)

End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim ObjInd As Integer
    
    ObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    If ObjInd > 0 Then
        If ObjData(ObjInd).proyectil = 1 Then
            ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
            If ObjData(ObjInd).Subtipo = 3 Then
                GoTo Envenena
            End If
        End If
    End If
    
    ObjInd = UserList(AtacanteIndex).Invent.AnilloEqpObjIndex
    If ObjInd > 0 Then
        If ObjData(ObjInd).Efectomagico = eEfectos.Envenena Then
            GoTo Envenena
        End If
    End If
    
    Exit Sub
    
Envenena:
    If RandomNumber(1, 40) < 3 Then
        UserList(VictimaIndex).flags.Envenenado = 3
        Call WriteConsoleMsg(VictimaIndex, UserList(AtacanteIndex).Name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(AtacanteIndex, "Has envenenado a " & UserList(VictimaIndex).Name & "!!", FontTypeNames.FONTTYPE_FIGHT)
    End If
    Call FlushBuffer(VictimaIndex)
    Exit Sub

End Sub

Sub UserIncinera(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 05/09/2020
    '
    '***************************************************

    Dim ObjInd As Integer
    
    ObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    If ObjInd > 0 Then
        If ObjData(ObjInd).proyectil = 1 Then
            ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
            If ObjData(ObjInd).Subtipo = 2 Then
                GoTo Incinera
            End If
        End If
    End If
    
    ObjInd = UserList(AtacanteIndex).Invent.AnilloEqpObjIndex
    If ObjInd > 0 Then
        If ObjData(ObjInd).Efectomagico = eEfectos.Incinera Then
            GoTo Incinera
        End If
    End If
    
    Exit Sub
    
Incinera:
    If RandomNumber(1, 35) < 2 Then
        UserList(VictimaIndex).flags.Incinerado = 1
        Call WriteConsoleMsg(VictimaIndex, UserList(AtacanteIndex).Name & " te ha incinerado!!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(AtacanteIndex, "Has incinerado a " & UserList(VictimaIndex).Name & "!!", FontTypeNames.FONTTYPE_FIGHT)
    End If
    Call FlushBuffer(VictimaIndex)
    Exit Sub

End Sub

Public Sub GolpeParaliza(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
'*********************************************************************
'Autor: Lorwik
'Fecha: 10/03/2021
'Descripci�n: Probabilidad de inmovilizar al enemigo si tiene la orbe acuatica equipada
'*********************************************************************

    Dim res As Byte
    Dim probm As Integer
    Dim UsandoOrbe As Boolean

    Dim TuvoExito As Boolean
    Dim ParalisisSpell As Byte
    
    ParalisisSpell = 9

    With UserList(UserIndex)
    
        If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
        
        'Esta usando un arma con la orbe acuatica?
        If .Invent.WeaponEqpObjIndex > 0 And .Invent.AnilloEqpObjIndex > 0 Then
            If ObjData(.Invent.AnilloEqpObjIndex).Efectomagico = eEfectos.Paraliza Then _
                UsandoOrbe = True
                
        Else
            Exit Sub
            
        End If
        
        'Si no tiene orbe y esta desarmado.
        If UsandoOrbe And Not .Invent.WeaponEqpObjIndex <> 0 Then
            res = RandomNumber(0, ObtenerSuerte(.Stats.UserSkills(eSkill.Marciales)))
            
            'Nudillos sube la probabilidad
            If .Invent.NudiEqpIndex <> 0 Then probm = 10
            'Se gladiador tambien
            If .clase = eClass.Gladiador Then probm = probm + 10
            
            res = res - Porcentaje(res, probm)
            If res < 5 Then
                TuvoExito = True
            End If
            
        ElseIf UsandoOrbe Then 'Pero si la tiene...
            res = RandomNumber(1, 40)
            If res < 3 Then
                TuvoExito = True
            End If
            
        End If
    
        '�Tuvo exito al paralizar?
        If TuvoExito Then
            UserList(VictimaIndex).flags.Paralizado = 1
            UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado * 0.5
            Call WriteParalizeOK(VictimaIndex)

            If Hechizos(ParalisisSpell).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(Hechizos(ParalisisSpell).WAV, UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y))
            If Hechizos(ParalisisSpell).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, Hechizos(ParalisisSpell).FXgrh, Hechizos(ParalisisSpell).loops))
            If Hechizos(ParalisisSpell).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateParticleChar(UserList(VictimaIndex).Char.CharIndex, Hechizos(ParalisisSpell).Particle, True, GENERAL_PARTICLE_LIFE))
                
            Call WriteConsoleMsg(UserIndex, "Tu golpe ha dejado inm�vil a tu oponente", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(VictimaIndex, "�El golpe te ha dejado inm�vil!", FontTypeNames.FONTTYPE_INFO)
        End If
    
    End With
End Sub

Public Sub GolpeParalizaNpc(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)
'*********************************************************************
'Autor: Lorwik
'Fecha: 10/03/2021
'Descripci�n: Probabilidad de inmovilizar al enemigo si tiene la orbe acuatica equipada
'*********************************************************************

    Dim res As Byte
    Dim probm As Integer
    Dim UsandoOrbe As Boolean
    Dim ParalisisSpell As Byte
    Dim TuvoExito As Boolean
    
    ParalisisSpell = 9
    
    With UserList(UserIndex)
    
        If Npclist(NPCIndex).flags.Paralizado = 1 Then Exit Sub
    
        If .Invent.WeaponEqpObjIndex > 0 And .Invent.AnilloEqpObjIndex > 0 Then
            If ObjData(.Invent.AnilloEqpObjIndex).Efectomagico = eEfectos.Paraliza Then _
                UsandoOrbe = True
        Else
            Exit Sub
        End If
        
        If Not UsandoOrbe And .Invent.WeaponEqpObjIndex = 0 Then
            res = RandomNumber(0, ObtenerSuerte(.Stats.UserSkills(eSkill.Marciales)))
            
            If .Invent.NudiEqpIndex <> 0 Then probm = 10
            If .clase = eClass.Gladiador Then probm = probm + 10
            res = res - Porcentaje(res, probm)
            Debug.Print "LAAAAAAA"
            If res < 5 Then
                TuvoExito = True
                
            End If
        ElseIf UsandoOrbe Then
            Debug.Print "LOOOOOOOOO"
            res = RandomNumber(1, 35)
            If res < 5 Then
                TuvoExito = True
                
            End If
        End If
        
        If TuvoExito Then
            Npclist(NPCIndex).flags.Paralizado = 1
            Npclist(NPCIndex).Contadores.Paralisis = IntervaloParalizado
            Call WriteConsoleMsg(UserIndex, "Tu golpe ha dejado inm�vil a la criatura", FontTypeNames.FONTTYPE_INFO)
                
            If Hechizos(ParalisisSpell).WAV <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(ParalisisSpell).WAV, Npclist(NPCIndex).Pos.X, Npclist(NPCIndex).Pos.Y))
            If Hechizos(ParalisisSpell).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(Npclist(NPCIndex).Char.CharIndex, Hechizos(ParalisisSpell).FXgrh, Hechizos(ParalisisSpell).loops))
            If Hechizos(ParalisisSpell).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticleChar(Npclist(NPCIndex).Char.CharIndex, Hechizos(ParalisisSpell).Particle, True, GENERAL_PARTICLE_LIFE))
        End If
    
    End With
End Sub

Public Sub LanzarProyectil(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 10/07/2010
    'Throws an arrow or knive to target user/npc.
    '***************************************************
    On Error GoTo errHandler

    Dim MunicionSlot    As Byte

    Dim MunicionIndex   As Integer

    Dim WeaponSlot      As Byte

    Dim WeaponIndex     As Integer

    Dim TargetUserIndex As Integer

    Dim TargetNpcIndex  As Integer

    Dim DummyINT        As Integer
    
    Dim Threw           As Boolean

    Threw = True
    
    'Make sure the item is valid and there is ammo equipped.
    With UserList(UserIndex)
        
        With .Invent
            MunicionSlot = .MunicionEqpSlot
            MunicionIndex = .MunicionEqpObjIndex
            WeaponSlot = .WeaponEqpSlot
            WeaponIndex = .WeaponEqpObjIndex

        End With
        
        ' Tiene arma equipada?
        If WeaponIndex = 0 Then
            DummyINT = 1
            Call WriteConsoleMsg(UserIndex, "No tienes un arco o cuchilla equipada.", FontTypeNames.FONTTYPE_INFO)
            
            ' En un slot valido?
        ElseIf WeaponSlot < 1 Or WeaponSlot > .CurrentInventorySlots Then
            DummyINT = 1
            Call WriteConsoleMsg(UserIndex, "No tienes un arco o cuchilla equipada.", FontTypeNames.FONTTYPE_INFO)
            
            ' Usa municion? (Si no la usa, puede ser un arma arrojadiza)
        ElseIf ObjData(WeaponIndex).Municion = 1 Then
        
            ' La municion esta equipada en un slot valido?
            If MunicionSlot < 1 Or MunicionSlot > .CurrentInventorySlots Then
                DummyINT = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones equipadas.", FontTypeNames.FONTTYPE_INFO)
                
                ' Tiene municion?
            ElseIf MunicionIndex = 0 Then
                DummyINT = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones equipadas.", FontTypeNames.FONTTYPE_INFO)
                
                ' Son flechas?
            ElseIf ObjData(MunicionIndex).OBJType <> eOBJType.otFlechas Then
                DummyINT = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones.", FontTypeNames.FONTTYPE_INFO)
                
                ' Tiene suficientes?
            ElseIf .Invent.Object(MunicionSlot).Amount < 1 Then
                DummyINT = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones.", FontTypeNames.FONTTYPE_INFO)

            End If
            
            ' Es un arma de proyectiles?
        ElseIf ObjData(WeaponIndex).proyectil <> 1 Then
            DummyINT = 2

        End If
        
        If DummyINT <> 0 Then
            If DummyINT = 1 Then
                Call Desequipar(UserIndex, WeaponSlot)

            End If
            
            Call Desequipar(UserIndex, MunicionSlot)
            Exit Sub

        End If
    
        'Quitamos stamina
        If .Stats.MinSta >= 10 Then
            Call QuitarSta(UserIndex, RandomNumber(1, 10))
        Else

            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(UserIndex, "Estas muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Estas muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)

            End If

            Exit Sub

        End If
        
        Call LookatTile(UserIndex, .Pos.Map, X, Y)
        
        TargetUserIndex = .flags.TargetUser
        TargetNpcIndex = .flags.TargetNPC
        
        'Validate target
        If TargetUserIndex > 0 Then

            'Only allow to atack if the other one can retaliate (can see us)
            If Abs(UserList(TargetUserIndex).Pos.X - .Pos.X) > RANGO_VISION_X Or Abs(UserList(TargetUserIndex).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
            
            'Prevent from hitting self
            If TargetUserIndex = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Attack!
            Threw = UsuarioAtacaUsuario(UserIndex, TargetUserIndex)
            
        ElseIf TargetNpcIndex > 0 Then

            'Only allow to atack if the other one can retaliate (can see us)
            If Abs(Npclist(TargetNpcIndex).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(Npclist(TargetNpcIndex).Pos.X - .Pos.X) > RANGO_VISION_X Then
                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
            
            'Is it attackable???
            If Npclist(TargetNpcIndex).Attackable <> 0 Then
                'Attack!
                Threw = UsuarioAtacaNpc(UserIndex, TargetNpcIndex)

            End If

        End If
        
        ' Solo pierde la municion si pudo atacar al target, o tiro al aire
        If Threw Then
            
            Dim Slot As Byte
            
            ' Tiene equipado arco y flecha?
            If ObjData(WeaponIndex).Municion = 1 Then
                Slot = MunicionSlot
                ' Tiene equipado un arma arrojadiza
            Else
                Slot = WeaponSlot

            End If
            
            'Take 1 knife/arrow away
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            Call UpdateUserInv(False, UserIndex, Slot)
            
        End If
        
    End With
    
    Exit Sub

errHandler:

    Dim UserName As String

    If UserIndex > 0 Then UserName = UserList(UserIndex).Name

    Call LogError("Error en LanzarProyectil " & Err.Number & ": " & Err.description & ". User: " & UserName & "(" & UserIndex & ")")

End Sub

