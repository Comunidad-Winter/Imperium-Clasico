Attribute VB_Name = "mMainLoop"
Option Explicit

Public prgRun As Boolean

Private Const INTERVALO_AI_GENERAL As Long = 350

Public Sub Auditoria()

    On Error GoTo errhand
    
    Call PasarSegundo 'sistema de desconexion de 10 segs
    
    Static centinelSecs As Byte

    centinelSecs = centinelSecs + 1

    If centinelSecs = 30 Then
        'Every 5 seconds, we try to call the player's attention so it will report the code.
        Call modCentinela.AvisarUsuarios
    
        centinelSecs = 0

    End If

    Exit Sub

errhand:

    Call LogError("Error en Timer Auditoria. Err: " & Err.description & " - " & Err.Number)

End Sub

Public Sub PacketResend()

    '***************************************************
    'Autor: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/01/07
    'Attempts to resend to the user all data that may be enqueued.
    '***************************************************
    On Error GoTo errHandler:

    Dim i As Long
    For i = 1 To LastUser
        If UserList(i).ConnIDValida Then Call FlushBuffer(i)
    Next i

    Exit Sub

errHandler:
    Call LogError("Error en packetResend - Error: " & Err.Number & " - Desc: " & Err.description)

    Resume Next

End Sub

Public Sub TIMER_AI()

    On Error GoTo ErrorHandler

    Dim NPCIndex As Long
    Dim Mapa     As Integer
    Dim e_p      As Integer
    
    'Barrin 29/9/03
    If Not haciendoBK And Not EnPausa Then

        'Update NPCs
        For NPCIndex = 1 To LastNPC
            
            With Npclist(NPCIndex)

                If .flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
                
                    ' Chequea si contiua teniendo dueno
                    If .Owner > 0 Then
                        Call ValidarPermanenciaNpc(NPCIndex)
                    End If
                
                    If .flags.Paralizado = 1 Then
                        Call EfectoParalisisNpc(NPCIndex)
                    Else

                        'Usamos AI si hay algun user en el mapa
                        If .flags.Inmovilizado = 1 Then
                            Call EfectoParalisisNpc(NPCIndex)

                        End If
                            
                        Mapa = .Pos.Map
                            
                        If Mapa > 0 Then
                            'Si no hay usuarios en el mapa no hacemos nada
                            If MapInfo(Mapa).NumUsers > 0 Then
                                '�El NPC tiene movimiento?
                                If .Movement <> TipoAI.ESTATICO Then
                                    Call NPCAI(NPCIndex)

                                End If

                            End If

                        End If

                    End If

                End If

            End With

        Next NPCIndex

    End If
    
    Exit Sub

ErrorHandler:
    Call LogError("Error en TIMER_AI_Timer " & Npclist(NPCIndex).Name & " mapa:" & Npclist(NPCIndex).Pos.Map)
    Call MuereNpc(NPCIndex, 0)

End Sub

Public Sub GameTimer()

    '********************************************************
    'Author: Unknown
    'Last Modify Date: -
    '********************************************************
    Dim iUserIndex   As Long
    Dim bEnviarStats As Boolean
    Dim bEnviarAyS   As Boolean
    Dim i        As Long
    
    On Error GoTo hayerror
    
    '<<<<<< Procesa eventos de los usuarios >>>>>>
    For iUserIndex = 1 To LastUser

        With UserList(iUserIndex)

            'Conexion activa?
            If .ConnID <> -1 Then
                'User valido?
                
                If .ConnIDValida And .flags.UserLogged Then
                    
                    '[Alejo-18-5]
                    bEnviarStats = False
                    bEnviarAyS = False
                    
                    If .flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
                    If .flags.Ceguera = 1 Or .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
                    If .flags.Muerto = 0 Then
                        
                        '[Consejeros]
                        If (.flags.Privilegios And PlayerType.User) Then Call EfectoLava(iUserIndex)
                        
                        If .flags.Desnudo <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoFrio(iUserIndex)
                        
                        If .flags.Meditando Then Call DoMeditar(iUserIndex)
                        
                        If .flags.Envenenado <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoVeneno(iUserIndex)
                        
                        If .flags.Incinerado <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoIncinerado(iUserIndex)
                        
                        If .flags.AdminInvisible <> 1 Then
                            If .flags.invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)

                        End If
                        
                        If .flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
                        
                        'Macro de Trabajo
                        If .flags.MacroTrabajo <> 0 Then
                            .Counters.MacroTrabajo = .Counters.MacroTrabajo + 1
                            If .Counters.MacroTrabajo >= IntervaloPuedeMakrear Then
                                .Counters.MacroTrabajo = 0
                                Call MacroTrabajo(iUserIndex, .flags.MacroTrabajo)
                            End If
                        End If
                        
                        If .flags.AtacablePor <> 0 Then Call EfectoEstadoAtacable(iUserIndex)
                        
                        Call DuracionPociones(iUserIndex)
                        
                        Call HambreYSed(iUserIndex, bEnviarAyS)
                        
                        If .flags.Hambre = 0 And .flags.Sed = 0 Then
                            If Lloviendo Then
                                If Not Intemperie(iUserIndex) Then
                                    If Not .flags.Descansar Then
                                        'No esta descansando
                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)

                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False

                                        End If
                                        
                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)

                                        If bEnviarStats Then
                                            Call WriteUpdateSta(iUserIndex)
                                            bEnviarStats = False

                                        End If

                                    Else
                                        'Esta descansando
                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)

                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False

                                        End If

                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)

                                        If bEnviarStats Then
                                            Call WriteUpdateSta(iUserIndex)
                                            bEnviarStats = False

                                        End If

                                        'termina de descansar automaticamente
                                        If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
                                            Call WriteRestOK(iUserIndex)
                                            Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                            .flags.Descansar = False

                                        End If
                                        
                                    End If

                                Else

                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloLloviendo - UserList(iUserIndex).Stats.UserSkills(eSkill.Supervivencia))
                                    
                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False

                                    End If

                                End If

                            Else

                                If Not .flags.Descansar Then
                                    'No esta descansando
                                    
                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)

                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False

                                    End If

                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)

                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False

                                    End If
                                    
                                Else
                                    'esta descansando
                                    
                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)

                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False

                                    End If

                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)

                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False

                                    End If

                                    'termina de descansar automaticamente
                                    If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
                                        Call WriteRestOK(iUserIndex)
                                        Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                        .flags.Descansar = False

                                    End If
                                    
                                End If

                            End If

                        End If
                        
                        If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                        
                        If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
                        
                    End If 'Muerto
                
                'Inactividad de cuentas
                ElseIf .ConnIDValida And .flags.UserLogged = False And .flags.AccountLogged Then
                    '.Counters.IdleCount = .Counters.IdleCount + 1
                    
                    If .Counters.IdleCount > IntervaloParaConexion Then
                        .Counters.IdleCount = 0
                        Call CloseSocket(iUserIndex)
                    End If
                    
                Else 'no esta logeado?
                    'Inactive players will be removed!
                    .Counters.IdleCount = .Counters.IdleCount + 1

                    If .Counters.IdleCount > IntervaloParaConexion Then
                        .Counters.IdleCount = 0
                        Call CloseUser(iUserIndex)

                    End If

                End If 'UserLogged
                
                'Ya terminamos de procesar el paquete, sigamos recibiendo.
                .Counters.PacketsTick = 0
                
            End If

        End With

    Next iUserIndex

    Exit Sub

hayerror:
    LogError ("Error en GameTimer: " & Err.description & " UserIndex = " & iUserIndex)

End Sub

Public Sub PasarSegundo()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Dim i As Long
    
    'Limpieza del mundo
    If tickLimpieza > 0 Then
        tickLimpieza = tickLimpieza - 1
                
        Select Case tickLimpieza
                                                        
            Case 300
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza del mundo en 5 Minuto. Atentos!!", FontTypeNames.FONTTYPE_SERVER))

            Case 60
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza del mundo en 1 Minuto. Atentos!!", FontTypeNames.FONTTYPE_SERVER))
                
            Case 5 To 1
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza del mundo en " & tickLimpieza & " segundos. Atentos!!", FontTypeNames.FONTTYPE_SERVER))
            
            Case 0
                Call BorrarObjetosLimpieza
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza del mundo finalizada.", FontTypeNames.FONTTYPE_SERVER))
                
        End Select
        
    End If
    
    For i = 1 To LastUser

        With UserList(i)

            If .flags.UserLogged Then
            
                'Portales
                If .PortalTiempo > 0 Then
                    .PortalTiempo = .PortalTiempo - 1
                    If .PortalTiempo < 1 Then Call Borrar_Portal_User(i)
                End If
                
                '�Esta casteando un hechizo?
                If .flags.CasteoSpell.Casteando = True Then
                    .flags.CasteoSpell.TimeCast = .flags.CasteoSpell.TimeCast - 1
                    
                    If .flags.CasteoSpell.TimeCast <= 0 Then
                        Call LanzarHechizo(.flags.CasteoSpell.SpellID, i)
                        Call ResetCasteo(i)
                    End If
                End If
            
                'Cerrar usuario
                If .Counters.Saliendo Then
                    .Counters.Salir = .Counters.Salir - 1
                    Call WriteConsoleMsg(i, "Cerrando en... " & .Counters.Salir, FontTypeNames.FONTTYPE_INFO)
                    
                    If .Counters.Salir <= 0 Then
                        Call WriteConsoleMsg(i, "Gracias por jugar WinterAO", FontTypeNames.FONTTYPE_INFO)
                        Call WriteDisconnect(i)
                        Call FlushBuffer(i)
                        Call CloseUser(i)
                    End If

                End If
                
                ' Conteo de los Retos
                If .Counters.TimeFight > 0 Then
                    .Counters.TimeFight = .Counters.TimeFight - 1
                    
                    ' Cuenta regresiva de retos y eventos
                    If .Counters.TimeFight = 0 Then
                        Call WriteConsoleMsg(i, "Cuenta -> YA!", FontTypeNames.FONTTYPE_FIGHT)
                                             
                        If .flags.SlotReto > 0 Then
                            Call WriteUserInEvent(i)
                        End If
                    
                    Else
                        Call WriteConsoleMsg(i, "Cuenta -> " & .Counters.TimeFight, FontTypeNames.FONTTYPE_GUILD)
                    
                    End If
                
                End If
                
                If .Counters.Pena > 0 Then

                    'Restamos las penas del personaje
                    If .Counters.Pena > 0 Then
                        .Counters.Pena = .Counters.Pena - 1
                 
                        If .Counters.Pena < 1 Then
                            .Counters.Pena = 0
                            Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                            Call WriteConsoleMsg(i, "Has sido liberado!", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If
                    
                End If
                
                If Not .Pos.Map = 0 Then

                    'Counter de piquete
                    If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.ANTIPIQUETE Then
                            If .flags.Muerto = 0 Then
                                .Counters.PiqueteC = .Counters.PiqueteC + 1
                                .Counters.ContadorPiquete = .Counters.ContadorPiquete + 1
                                If .Counters.ContadorPiquete = 6 Then
                                    Call WriteConsoleMsg(i, "Estas obstruyendo la via publica, muevete o seras encarcelado!!!", FontTypeNames.FONTTYPE_INFO)
                                    .Counters.ContadorPiquete = 0
                                End If
                                If .Counters.PiqueteC >= 30 Then
                                    .Counters.PiqueteC = 0
                                    .Counters.ContadorPiquete = 0
                                    Call Encarcelar(i, MinutosCarcelPiquete)
                                End If
                        Else
                            .Counters.PiqueteC = 0

                        End If

                    Else
                        .Counters.PiqueteC = 0

                    End If

                End If

            End If

        End With

    Next i

    Exit Sub

errHandler:
    Call LogError("Error en PasarSegundo. Err: " & Err.description & " - " & Err.Number & " - UserIndex: " & i)

    Resume Next

End Sub