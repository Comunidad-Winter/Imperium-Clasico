Attribute VB_Name = "Acciones"
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

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal UserIndex As Integer, _
           ByVal Map As Integer, _
           ByVal X As Integer, _
           ByVal Y As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim tempIndex As Integer
    
    On Error Resume Next

    'Rango Vision? (ToxicWaste)
    If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
        Exit Sub

    End If
    
    'Posicion valida?
    If InMapBounds(Map, X, Y) Then

        With UserList(UserIndex)

            If MapData(Map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
                tempIndex = MapData(Map, X, Y).NpcIndex
                
                'Set the target NPC
                .flags.TargetNPC = tempIndex
                
                If Npclist(tempIndex).Comercia = 1 Then

                    'Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                        Exit Sub

                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub

                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    'Iniciamos la rutina pa' comerciar.
                    Call IniciarComercioNPC(UserIndex)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Banquero Then

                    'Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                        Exit Sub

                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub

                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    'A depositar de una
                    Call WriteAbrirGoliath(UserIndex)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Revividor Then

                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                        Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    'Revivimos si es necesario
                    If .flags.Muerto = 1 And (Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex)) Then
                        Call SacerdoteResucitateUser(UserIndex)
                    End If
                    
                    If Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex) Then
                        Call SacerdoteHealUser(UserIndex)
                    End If
                    
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Veterinario Then
                
                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                        Call WriteConsoleMsg(UserIndex, "El veterinario no puede atender a tu familiar debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    If .Familiar.Tipo < 1 Then
                        Call WriteConsoleMsg(UserIndex, "No tienes familiares.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    If .Familiar.Muerto = 0 And .Familiar.MaxHp = .Familiar.MinHp Then
                        Call WriteConsoleMsg(UserIndex, "Tu familiar esta sano.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    Call Familiar_Muerte(UserIndex, False)
                    
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Entrenador Then
                
                    Call AccionParaEntrenador(UserIndex)
                    
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Subastador Then
                
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 1 Then
                        Call WriteConsoleMsg(UserIndex, "El subastador pide que te acerques m�s.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    .flags.Subastando = True
                    Call WriteIniciarSubastasOrConsulta(UserIndex)
                  
                End If

                'Es un obj?
            ElseIf MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
                
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType

                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X, Y, UserIndex)

                    Case eOBJType.otCarteles 'Es un cartel
                        Call AccionParaCartel(Map, X, Y, UserIndex)

                    Case eOBJType.otForos 'Foro
                        Call AccionParaForo(Map, X, Y, UserIndex)

                    Case eOBJType.otLena    'Lena

                        If tempIndex = FOGATA_APAG And .flags.Muerto = 0 Then
                            Call AccionParaRamita(Map, X, Y, UserIndex)

                        End If
                        
                    Case eOBJType.OtPozos 'Pozos
                        Call AccionParaPozos(Map, X, Y, UserIndex)

                End Select

                '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
            ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X + 1, Y).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
                    
                End Select
            
            ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
        
                Select Case ObjData(tempIndex).OBJType

                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)

                End Select
            
            ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType

                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X, Y + 1, UserIndex)

                End Select

            End If

        End With

    End If

End Sub

Public Sub AccionParaForo(ByVal Map As Integer, _
                          ByVal X As Integer, _
                          ByVal Y As Integer, _
                          ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/01/2010
    '02/01/2010: ZaMa - Agrego foros faccionarios
    '***************************************************

    On Error Resume Next

    Dim Pos As WorldPos
    
    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y
    
    If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    If SendPosts(UserIndex, ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).ForoID) Then
        Call WriteShowForumForm(UserIndex)

    End If
    
End Sub

Sub AccionParaPuerta(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
        If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
            If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 Then

                'Abre la puerta
                If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexAbierta
                    
                    Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).GrhIndex, 0, 0, 0, X, Y))
                    
                    'Desbloquea
                    MapData(Map, X, Y).Blocked = 0
                    MapData(Map, X - 1, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(True, Map, X, Y, 0)
                    Call Bloquear(True, Map, X - 1, Y, 0)
                      
                    'Sonido
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                'Cierra puerta
                MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexCerrada
                
                Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).GrhIndex, 0, 0, 0, X, Y))
                                
                MapData(Map, X, Y).Blocked = 1
                MapData(Map, X - 1, Y).Blocked = 1
                
                Call Bloquear(True, Map, X - 1, Y, 1)
                Call Bloquear(True, Map, X, Y, 1)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))

            End If
        
            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
        Else
            Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)

        End If

    Else
        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = 8 Then
  
        If Len(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).texto) > 0 Then
            Call WriteShowSignal(UserIndex, MapData(Map, X, Y).ObjInfo.ObjIndex)

        End If
  
    End If

End Sub

Sub AccionParaRamita(ByVal Map As Integer, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim Suerte             As Byte

    Dim Exito              As Byte

    Dim obj                As obj

    Dim SkillSupervivencia As Byte

    Dim Pos                As WorldPos
    
    With Pos
        .Map = Map
        .X = X
        .Y = Y
    End With
    

    With UserList(UserIndex)

        If Distancia(Pos, .Pos) > 2 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        If MapData(Map, X, Y).Trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
            Call WriteConsoleMsg(UserIndex, "No puedes hacer fogatas en zona segura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        SkillSupervivencia = .Stats.UserSkills(eSkill.Supervivencia)
    
        If SkillSupervivencia < 6 Then
            Suerte = 3
        
        ElseIf SkillSupervivencia <= 10 Then
            Suerte = 2
        
        Else
            Suerte = 1

        End If
    
        Exito = RandomNumber(1, Suerte)
    
        If Exito = 1 Then
            If MapInfo(.Pos.Map).Zona <> Ciudad Then
            
                With obj
                    .ObjIndex = FOGATA
                    .Amount = 1
                End With
            
                Call WriteConsoleMsg(UserIndex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
            
                Call MakeObj(obj, Map, X, Y)
            
                Call mLimpieza.AgregarObjetoLimpieza(Pos)
            
                Call SubirSkill(UserIndex, eSkill.Supervivencia, True)
            Else
                Call WriteConsoleMsg(UserIndex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
            Call SubirSkill(UserIndex, eSkill.Supervivencia, False)

        End If

    End With

End Sub

Public Sub AccionParaSacerdote(ByVal UserIndex As Integer)
'******************************
'Adaptacion a 13.0: Kaneidra
'Last Modification: 07/01/2020
'Refactorizo para que el Sacerdote haga una sola cosa y no 20 diferentes alrededor del codigo dependiendo de como se usa (Recox)
'******************************
    
    With UserList(UserIndex)
        
        ' Si esta muerto...
        If .flags.Muerto = 1 Then
            Call SacerdoteResucitateUser(UserIndex)

        End If
        
        ' Si esta herido... lo curamos.
        If .Stats.MinHp < .Stats.MaxHp Then
            Call SacerdoteHealUser(UserIndex)
        End If
        
    End With
 
End Sub

Public Sub AccionParaEntrenador(ByVal UserIndex As Integer)
'******************************
'Autor: Lorwik
'Last Modification: 18/05/2020
'Refactorizo para que el Entrenador mande la lista para entrenar
'******************************
    With UserList(UserIndex)
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
    
        End If
            
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
    
        End If
            
        'Make sure it's the trainer
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
            
        Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC)
    End With
End Sub

Sub AccionParaPozos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
        
    On Error GoTo AccionParaPozos_Err

    Dim Pos As WorldPos

    Pos.Map = Map
    Pos.X = X
    Pos.Y = Y

    With UserList(UserIndex)
    
        'Comprobacion de distancia
        If Distancia(Pos, .Pos) > 2 Then
            Call WriteConsoleMsg(UserIndex, "�Estas demasiado lejos!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        
        End If
        
        
        If MapData(Map, X, Y).ObjInfo.Amount <= 1 Then
            Call WriteConsoleMsg(UserIndex, "El pozo esta drenado, regresa mas tarde...", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        
        End If
        
        'Tipo de pozo 1
        If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo = 1 Then
            If .Stats.MinMAN = .Stats.MaxMAN Then
                Call WriteConsoleMsg(UserIndex, "No tenes necesidad del pozo...", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
        
                End If
        
            .Stats.MinMAN = .Stats.MaxMAN
            MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount - 1
            Call WriteConsoleMsg(UserIndex, "El pozo sacia tu sed. �Tu man� a sido restaurada!", FontTypeNames.FONTTYPE_EJECUCION)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
            Call WriteUpdateUserStats(UserIndex)
            Exit Sub
        
        End If
    
        'Tipo de pozo 2
        If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo = 2 Then
            If .Stats.MinAGU = .Stats.MaxAGU Then
                Call WriteConsoleMsg(UserIndex, "No tenes necesidad del pozo...", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
        
            End If
        
            .Stats.MinAGU = .Stats.MaxAGU
            .flags.Sed = 0 'Bug reparado 27/01/13
            MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount - 1
            Call WriteConsoleMsg(UserIndex, "Sientes la frescura del pozo. �Ya no sientes sed!", FontTypeNames.FONTTYPE_EJECUCION)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
            Call WriteUpdateHungerAndThirst(UserIndex)
            Exit Sub
        
        End If
    
    End With
    
    Exit Sub

AccionParaPozos_Err:
    Call LogError(Err.Number & " - " & Err.description & "Acciones.AccionParaPozos")
    
End Sub

