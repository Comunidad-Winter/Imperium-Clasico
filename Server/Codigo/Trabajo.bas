Attribute VB_Name = "Trabajo"
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

Public Const EsfuerzoTalarGeneral As Byte = 4
Public Const EsfuerzoTalarLenador As Byte = 2

Public Const EsfuerzoPescarPescador As Byte = 1
Public Const EsfuerzoPescarGeneral As Byte = 3

Public Const EsfuerzoExcavarMinero As Byte = 2
Public Const EsfuerzoExcavarGeneral As Byte = 5

Public Const EsfuerzoRaicesBotanico As Byte = 2
Public Const EsfuerzoRaicesGeneral As Byte = 4

Private Const GASTO_ENERGIA As Byte = 6

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)

    '********************************************************
    'Autor: Nacho (Integer)
    'Last Modif: 11/19/2009
    'Chequea si ya debe mostrarse
    'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '13/01/2010: ZaMa - Arreglo condicional para que el bandido camine oculto.
    '********************************************************
    On Error GoTo errHandler

    With UserList(UserIndex)
    
        'Si tiene el anillo del ocultismo equipado, no contabiliza el tiempo
        If .Invent.AnilloEqpObjIndex > 0 Then
            If ObjData(.Invent.AnilloEqpObjIndex).Efectomagico = eEfectos.CaminaOculto Then Exit Sub
        End If

        .Counters.TiempoOculto = .Counters.TiempoOculto - 1

        If .Counters.TiempoOculto <= 0 Then
            If .clase = eClass.Hunter And .Stats.UserSkills(eSkill.Ocultarse) > 90 Then
                If .Invent.ArmourEqpObjIndex = 648 Or .Invent.ArmourEqpObjIndex = 360 Then
                    .Counters.TiempoOculto = IntervaloOculto
                    Exit Sub

                End If

            End If

            .Counters.TiempoOculto = 0
            .flags.Oculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Mercenario Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NingunAura, NingunAura)

                End If

            Else

                If .flags.invisible = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    Call SetInvisible(UserIndex, .Char.CharIndex, False)

                End If

            End If

        End If

    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en Sub DoPermanecerOculto")

End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 13/01/2010 (ZaMa)
    'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
    'Modifique la formula y ahora anda bien.
    '13/01/2010: ZaMa - El pirata se transforma en galeon fantasmal cuando se oculta en agua.
    '***************************************************

    On Error GoTo errHandler

    Dim Suerte As Double

    Dim res    As Integer

    Dim Skill  As Integer
    
    With UserList(UserIndex)
  
        If Not IntervaloPuedeOcultar(UserIndex) Then Exit Sub

        Skill = .Stats.UserSkills(eSkill.Ocultarse)
            
        Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
        
        If .clase = eClass.Thief Then Suerte = Suerte * 2
        
        res = RandomNumber(1, 100)
        
        If res <= Suerte Then
        
            .flags.Oculto = 1
            .Counters.TiempoOculto = IntervaloOculto
            
            ' No es pirata o es uno sin barca
            If .flags.Navegando = 0 Then
                Call SetInvisible(UserIndex, .Char.CharIndex, True)
        
                Call WriteConsoleMsg(UserIndex, "Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
                ' Es un pirata navegando
            Else
                ' Le cambiamos el body a galeon fantasmal
                .Char.body = iFragataFantasmal
                ' Actualizamos clientes
                Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NingunAura, NingunAura)

            End If
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse, True)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 4 Then
                Call WriteConsoleMsg(UserIndex, "No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 4

            End If

            '[/CDT]
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse, False)

        End If
        
        .Counters.Ocultando = .Counters.Ocultando + 1

    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, _
                    ByRef Barco As ObjData, _
                    ByVal Slot As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2020 (Recox)
'13/01/2010: ZaMa - El pirata pierde el ocultar si desequipa barca.
'16/09/2010: ZaMa - Ahora siempre se va el invi para los clientes al equipar la barca (Evita cortes de cabeza).
'10/12/2010: Pato - Limpio las variables del inventario que hacen referencia a la barca, sino el pirata que la ultima barca que equipo era el galeon no explotaba(Y capaz no la tenia equipada :P).
'12/01/2020: Recox - Se refactorizo un poco para reutilizar con monturas .
'***************************************************

    Dim ModNave As Single
    
    With UserList(UserIndex)
        If .flags.Equitando = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes navegar mientras estas en tu montura!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        
        ModNave = ModNavegacion(.clase, UserIndex)
            
        If .Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
    
        End If
        
        ' No estaba navegando
        If .flags.Navegando = 0 Then
            
            Call ComenzaraNavegar(UserIndex, Slot)
        
        ' Estaba navegando
        Else
        
            Call DejardeNavegar(UserIndex)

        End If

    End With
    
End Sub

Public Sub ComenzaraNavegar(ByVal UserIndex As Integer, ByVal Slot As Integer)

    With UserList(UserIndex)
    
        .Invent.BarcoObjIndex = .Invent.Object(Slot).ObjIndex
        .Invent.BarcoSlot = Slot
            
        .Char.Head = 0
            
        ' No esta muerto
        If .flags.Muerto = 0 Then
            Call ToggleBoatBody(UserIndex)
            Call SetVisibleStateForUserAfterNavigateOrEquitate(UserIndex)
                
        ' Esta muerto
        Else
            .Char.body = iFragataFantasmal
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            .Char.AuraAnim = NingunAura
            .Char.AuraColor = NingunAura
                
        End If
            
        ' Comienza a navegar
        .flags.Navegando = 1
        
        ' Actualizo clientes
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
    
        Call WriteNavigateToggle(UserIndex)
    
    End With

End Sub

Public Sub DejardeNavegar(ByVal UserIndex As Integer)

    With UserList(UserIndex)
    
        .Invent.BarcoObjIndex = 0
        .Invent.BarcoSlot = 0
    
        ' No esta muerto
        If .flags.Muerto = 0 Then
            .Char.Head = .OrigChar.Head
                
            Call SetEquipmentOnCharAfterNavigateOrEquitate(UserIndex)
                
            ' Al dejar de navegar, si estaba invisible actualizo los clientes
            If .flags.invisible = 1 Then
                Call SetInvisible(UserIndex, .Char.CharIndex, True)
            End If
                
        ' Esta muerto
        Else
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            .Char.AuraAnim = NingunAura
            .Char.AuraColor = NingunAura

        End If
            
        ' Termina de navegar
        .flags.Navegando = 0
        
        ' Actualizo clientes
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
    
        Call WriteNavigateToggle(UserIndex)
    
    End With
End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    With UserList(UserIndex)

        If .flags.TargetObjInvIndex > 0 Then
           
            If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill <= UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) / ModFundicion(UserList(UserIndex).clase) Then
                Call DoLingotes(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de mineria suficientes para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)

            End If
        
        End If

    End With

    Exit Sub

errHandler:
    Call LogError("Error en FundirMineral. Error " & Err.Number & " : " & Err.description)

End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, _
                      ByVal cant As Long, _
                      ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 10/07/2010
    '10/07/2010: ZaMa - Ahora cant es long para evitar un overflow.
    '***************************************************

    Dim i     As Integer

    Dim Total As Long

    For i = 1 To UserList(UserIndex).CurrentInventorySlots

        If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
            Total = Total + UserList(UserIndex).Invent.Object(i).Amount

        End If

    Next i
    
    If cant <= Total Then
        TieneObjetos = True
        Exit Function

    End If
        
End Function

Public Sub QuitarObjetos(ByVal ItemIndex As Integer, _
                         ByVal cant As Integer, _
                         ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 05/08/09
    '05/08/09: Pato - Cambie la funcion a procedimiento ya que se usa como procedimiento siempre, y fixie el bug 2788199
    '***************************************************

    Dim i As Integer

    For i = 1 To UserList(UserIndex).CurrentInventorySlots

        With UserList(UserIndex).Invent.Object(i)

            If .ObjIndex = ItemIndex Then
                If .Amount <= cant And .Equipped = 1 Then Call Desequipar(UserIndex, i)
                
                .Amount = .Amount - cant

                If .Amount <= 0 Then
                    cant = Abs(.Amount)
                    UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
                    .Amount = 0
                    .ObjIndex = 0
                Else
                    cant = 0

                End If
                
                Call UpdateUserInv(False, UserIndex, i)
                
                If cant = 0 Then Exit Sub

            End If

        End With

    Next i

End Sub

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex)
    If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex)
    If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex)
End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Lena, ObjData(ItemIndex).Madera, UserIndex)
End Sub

Sub SastreQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).PielLobo > 0 Then Call QuitarObjetos(PielLobo, ObjData(ItemIndex).PielLobo, UserIndex)
    If ObjData(ItemIndex).PielOsoPardo > 0 Then Call QuitarObjetos(PielOsoPardo, ObjData(ItemIndex).PielOsoPardo, UserIndex)
    If ObjData(ItemIndex).PielOsoPolar > 0 Then Call QuitarObjetos(PielOsoPolar, ObjData(ItemIndex).PielOsoPolar, UserIndex)
End Sub

Sub AlquimistaQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).Raices > 0 Then Call QuitarObjetos(Raices, ObjData(ItemIndex).Raices, UserIndex)
End Sub

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, _
                                   Optional ByVal ShowMsg As Boolean = False) As Boolean
    
    If ObjData(ItemIndex).Madera > 0 Then
            If Not TieneObjetos(Lena, ObjData(ItemIndex).Madera, UserIndex) Then
                    If ShowMsg Then Call WriteConsoleMsg(UserIndex, "No tenes suficientes madera.", FontTypeNames.FONTTYPE_INFO)
                    CarpinteroTieneMateriales = False
                    Exit Function
            End If
    End If
    
    CarpinteroTieneMateriales = True

End Function
 
Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, _
                                   Optional ByVal ShowMsg As Boolean = False) As Boolean
    If ObjData(ItemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex) Then
                    If ShowMsg Then Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex) Then
                    If ShowMsg Then Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex) Then
                    If ShowMsg Then Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    HerreroTieneMateriales = True
End Function

Function SastreTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, _
                                   Optional ByVal ShowMsg As Boolean = False) As Boolean
    If ObjData(ItemIndex).PielLobo > 0 Then
            If Not TieneObjetos(PielLobo, ObjData(ItemIndex).PielLobo, UserIndex) Then
                    If ShowMsg Then Call WriteConsoleMsg(UserIndex, "No tenes suficientes pieles de lobo.", FontTypeNames.FONTTYPE_INFO)
                    SastreTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).PielOsoPardo > 0 Then
            If Not TieneObjetos(PielOsoPardo, ObjData(ItemIndex).PielOsoPardo, UserIndex) Then
                    If ShowMsg Then Call WriteConsoleMsg(UserIndex, "No tenes suficientes pieles de oso pardo.", FontTypeNames.FONTTYPE_INFO)
                    SastreTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).PielOsoPolar > 0 Then
            If Not TieneObjetos(PielOsoPolar, ObjData(ItemIndex).PielOsoPolar, UserIndex) Then
                    If ShowMsg Then Call WriteConsoleMsg(UserIndex, "No tenes suficientes pieles de oso polar.", FontTypeNames.FONTTYPE_INFO)
                    SastreTieneMateriales = False
                    Exit Function
            End If
    End If
    SastreTieneMateriales = True
End Function

Function AlquimistaTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, _
                                   Optional ByVal ShowMsg As Boolean = False) As Boolean
    
    If ObjData(ItemIndex).Raices > 0 Then
            If Not TieneObjetos(Raices, ObjData(ItemIndex).Raices, UserIndex) Then
                    If ShowMsg Then Call WriteConsoleMsg(UserIndex, "No tenes suficientes raices.", FontTypeNames.FONTTYPE_INFO)
                    AlquimistaTieneMateriales = False
                    Exit Function
            End If
    End If
    
    AlquimistaTieneMateriales = True

End Function

Public Function PuedeConstruirItemHerrero(ByVal UserIndex As Integer, _
                               ByVal ItemIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 24/08/2009
    '24/08/2008: ZaMa - Validates if the player has the required skill
    '16/11/2009: ZaMa - Validates if the player has the required amount of materials, depending on the number of items to make
    '***************************************************
    PuedeConstruirItemHerrero = HerreroTieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) >= _
        ObjData(ItemIndex).SkHerreria

End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Dim i As Long

    For i = 1 To UBound(ArmasHerrero)

        If ArmasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i

    For i = 1 To UBound(ArmadurasHerrero)

        If ArmadurasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i

    PuedeConstruirHerreria = False

End Function

Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 30/05/2010
    '16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items.
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '30/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '***************************************************

    Dim TieneMateriales As Boolean

    Dim OtroUserIndex   As Integer

    With UserList(UserIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
            
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(UserIndex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
            
                Call LimpiarComercioSeguro(UserIndex)

            End If

        End If
        
        'Sacamos energia
        'Chequeamos que tenga los puntos antes de sacarselos
        If .Stats.MinSta >= GASTO_ENERGIA Then
            .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
            Call DejardeTrabajar(UserIndex) 'Paramos el macro
            Exit Sub

        End If

        Call HerreroQuitarMateriales(UserIndex, ItemIndex)
        ' AGREGAR FX
        
        'Mensajes de exito
        Select Case ObjData(ItemIndex).OBJType
            Case eOBJType.otWeapon
                Call WriteConsoleMsg(UserIndex, "Has construido el arma!.", FontTypeNames.FONTTYPE_INFO)
                    
            Case eOBJType.otEscudo
                Call WriteConsoleMsg(UserIndex, "Has construido el escudo!.", FontTypeNames.FONTTYPE_INFO)
                    
            Case eOBJType.otCasco
                Call WriteConsoleMsg(UserIndex, "Has construido el casco!.", FontTypeNames.FONTTYPE_INFO)
                    
            Case eOBJType.otArmadura
                Call WriteConsoleMsg(UserIndex, "Has construido la armadura!.", FontTypeNames.FONTTYPE_INFO)
        End Select
        
        Dim MiObj As obj
        
        MiObj.Amount = 1
        MiObj.ObjIndex = ItemIndex

        If Not MeterItemEnInventario(UserIndex, MiObj) Then _
            Call TirarItemAlPiso(.Pos, MiObj)
        
        'Log de construccion de Items. Pablo (ToxicWaste) 10/09/07
        If ObjData(MiObj.ObjIndex).Log = 1 Then _
            Call LogDesarrollo(.Name & " ha construido " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name)
        
        Call SubirSkill(UserIndex, eSkill.Herreria, True)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TRABAJO_HERRERO, .Pos.X, .Pos.Y))
        
        If Not criminal(UserIndex) Then
            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

            If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

        End If
        
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

End Sub

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 28/05/2010
    '24/08/2008: ZaMa - Validates if the player has the required skill
    '16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '***************************************************
    On Error GoTo errHandler

    Dim TieneMateriales As Boolean

    Dim WeaponIndex     As Integer

    Dim OtroUserIndex   As Integer
    
    With UserList(UserIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(UserIndex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)

            End If

        End If
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
    
        If WeaponIndex <> SERRUCHO_CARPINTERO Then
            Call WriteConsoleMsg(UserIndex, "Debes tener equipado el serrucho para trabajar.", FontTypeNames.FONTTYPE_INFO)
            Call DejardeTrabajar(UserIndex) 'Paramos el macro
            Exit Sub

        End If
    
        If .Stats.UserSkills(eSkill.Carpinteria) >= ObjData(ItemIndex).SkCarpinteria Then
           
            'Sacamos energia
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA
                Call WriteUpdateSta(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
                Call DejardeTrabajar(UserIndex) 'Paramos el macro
                Exit Sub

            End If
            
            Call CarpinteroQuitarMateriales(UserIndex, ItemIndex)
            Call WriteConsoleMsg(UserIndex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
            
            Dim MiObj As obj

            MiObj.Amount = 1
            MiObj.ObjIndex = ItemIndex

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If
            
            'Log de construccion de Items. Pablo (ToxicWaste) 10/09/07
            If ObjData(MiObj.ObjIndex).Log = 1 Then
                Call LogDesarrollo(.Name & " ha construido " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name)

            End If
            
            Call SubirSkill(UserIndex, eSkill.Carpinteria, True)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TRABAJO_CARPINTERO, .Pos.X, .Pos.Y))
            
            If Not criminal(UserIndex) Then
                .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

                If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

            End If
            
            .Counters.Trabajando = .Counters.Trabajando + 1

        Else
            Call WriteConsoleMsg(UserIndex, "Aun no posees la habilidad suficiente para construir ese objeto. Necesitas al menos " & ObjData(ItemIndex).SkCarpinteria & " Skills.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
    
    Exit Sub
errHandler:
    Call LogError("Error en CarpinteroConstruirItem. Error " & Err.Number & " : " & Err.description & ". UserIndex:" & UserIndex & ". ItemIndex:" & ItemIndex)

End Sub

Public Sub SastreConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 21/08/2020
    '***************************************************
    On Error GoTo errHandler

    Dim TieneMateriales As Boolean

    Dim WeaponIndex     As Integer

    Dim OtroUserIndex   As Integer
    
    With UserList(UserIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(UserIndex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)

            End If

        End If
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
    
        If WeaponIndex <> KIT_DE_COSTURA Then
            Call WriteConsoleMsg(UserIndex, "Debes tener equipado el kit de sastreria para trabajar.", FontTypeNames.FONTTYPE_INFO)
            Call DejardeTrabajar(UserIndex) 'Paramos el macro
            Exit Sub

        End If
    
        If .Stats.UserSkills(eSkill.Sastreria) >= ObjData(ItemIndex).SkSastreria Then
           
            'Sacamos energia
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA
                Call WriteUpdateSta(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
                Call DejardeTrabajar(UserIndex) 'Paramos el macro
                Exit Sub

            End If
            
            Call SastreQuitarMateriales(UserIndex, ItemIndex)
            Call WriteConsoleMsg(UserIndex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
            
            Dim MiObj As obj

            MiObj.Amount = 1
            MiObj.ObjIndex = ItemIndex

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If
            
            'Log de construccion de Items. Pablo (ToxicWaste) 10/09/07
            If ObjData(MiObj.ObjIndex).Log = 1 Then
                Call LogDesarrollo(.Name & " ha construido " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name)

            End If
            
            Call SubirSkill(UserIndex, eSkill.Sastreria, True)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TRABAJO_CARPINTERO, .Pos.X, .Pos.Y))
            
            If Not criminal(UserIndex) Then
                .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

                If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

            End If
            
            .Counters.Trabajando = .Counters.Trabajando + 1

        Else
            Call WriteConsoleMsg(UserIndex, "Aun no posees la habilidad suficiente para construir ese objeto. Necesitas al menos " & ObjData(ItemIndex).SkSastreria & " Skills.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
    
    Exit Sub
errHandler:
    Call LogError("Error en SastreConstruirItem. Error " & Err.Number & " : " & Err.description & ". UserIndex:" & UserIndex & ". ItemIndex:" & ItemIndex)

End Sub

Public Sub AlquimistaConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 21/08/2020
    '***************************************************
    On Error GoTo errHandler

    Dim TieneMateriales As Boolean

    Dim WeaponIndex     As Integer

    Dim OtroUserIndex   As Integer
    
    With UserList(UserIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(UserIndex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)

            End If

        End If
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
    
        If WeaponIndex <> OLLA_ALQUIMISTA Then
            Call WriteConsoleMsg(UserIndex, "Debes tener equipado la olla de alquimista para trabajar.", FontTypeNames.FONTTYPE_INFO)
            Call DejardeTrabajar(UserIndex) 'Paramos el macro
            Exit Sub

        End If
    
        If .Stats.UserSkills(eSkill.Alquimia) >= ObjData(ItemIndex).SkAlquimia Then
           
            'Sacamos energia
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA
                Call WriteUpdateSta(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
                Call DejardeTrabajar(UserIndex) 'Paramos el macro
                Exit Sub

            End If
            
            Call AlquimistaQuitarMateriales(UserIndex, ItemIndex)
            Call WriteConsoleMsg(UserIndex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
            
            Dim MiObj As obj

            MiObj.Amount = 1
            MiObj.ObjIndex = ItemIndex

            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If
            
            'Log de construccion de Items. Pablo (ToxicWaste) 10/09/07
            If ObjData(MiObj.ObjIndex).Log = 1 Then
                Call LogDesarrollo(.Name & " ha construido " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name)

            End If
            
            Call SubirSkill(UserIndex, eSkill.Alquimia, True)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TRABAJO_CARPINTERO, .Pos.X, .Pos.Y))
            
            If Not criminal(UserIndex) Then
                .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

                If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

            End If
            
            .Counters.Trabajando = .Counters.Trabajando + 1

        Else
            Call WriteConsoleMsg(UserIndex, "Aun no posees la habilidad suficiente para construir ese objeto. Necesitas al menos " & ObjData(ItemIndex).SkAlquimia & " Skills.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
    
    Exit Sub
errHandler:
    Call LogError("Error en AlquimistaConstruirItem. Error " & Err.Number & " : " & Err.description & ". UserIndex:" & UserIndex & ". ItemIndex:" & ItemIndex)

End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Select Case Lingote

        Case iMinerales.HierroCrudo
            MineralesParaLingote = 14

        Case iMinerales.PlataCruda
            MineralesParaLingote = 20

        Case iMinerales.OroCrudo
            MineralesParaLingote = 35

        Case Else
            MineralesParaLingote = 10000

    End Select

End Function

Public Sub DoLingotes(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
    '***************************************************
    '    Call LogTarea("Sub DoLingotes")
    Dim Slot           As Integer

    Dim obji           As Integer

    Dim CantidadItems  As Integer

    Dim TieneMinerales As Boolean

    Dim OtroUserIndex  As Integer
    
    With UserList(UserIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(UserIndex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)

            End If

        End If
        
        CantidadItems = MaximoInt(1, CInt((.Stats.ELV - 4) / 5))

        Slot = .flags.TargetObjInvSlot
        obji = .Invent.Object(Slot).ObjIndex
        
        While CantidadItems > 0 And Not TieneMinerales

            If .Invent.Object(Slot).Amount >= MineralesParaLingote(obji) * CantidadItems Then
                TieneMinerales = True
            Else
                CantidadItems = CantidadItems - 1

            End If

        Wend
        
        If Not TieneMinerales Or ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - MineralesParaLingote(obji) * CantidadItems

        If .Invent.Object(Slot).Amount < 1 Then
            .Invent.Object(Slot).Amount = 0
            .Invent.Object(Slot).ObjIndex = 0

        End If
        
        Dim MiObj As obj

        MiObj.Amount = CantidadItems
        MiObj.ObjIndex = ObjData(.flags.TargetObjInvIndex).LingoteIndex

        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)

        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
        Call WriteConsoleMsg(UserIndex, "Has obtenido " & CantidadItems & " lingote" & IIf(CantidadItems = 1, "", "s") & "!", FontTypeNames.FONTTYPE_INFO)
    
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

End Sub

Function ModNavegacion(ByVal clase As eClass, ByVal UserIndex As Integer) As Single

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 27/11/2009
    '12/04/2010: ZaMa - Arreglo modificador de pescador, para que navegue con 60 skills.
    '***************************************************
    Select Case clase

        Case eClass.Mercenario
            ModNavegacion = 1

        Case Else
            ModNavegacion = 2

    End Select

End Function

Function ModFundicion(ByVal clase As eClass) As Integer

    Select Case clase
        Case eClass.Minero
            ModFundicion = 1
        Case eClass.Herrero
            ModFundicion = 1.2
        Case Else
            ModFundicion = 3
    End Select

End Function

Function ModCarpinteria(ByVal clase As eClass) As Integer

    Select Case clase
        Case eClass.Carpintero
            ModCarpinteria = 1
        Case Else
            ModCarpinteria = 3
    End Select

End Function

Function ModHerreria(ByVal clase As eClass) As Integer

    Select Case clase
        Case eClass.Herrero
            ModHerreria = 1
        Case eClass.Minero
            ModHerreria = 1.2
        Case Else
            ModHerreria = 4
    End Select

End Function

Function ModSastreria(ByVal clase As eClass) As Integer

    Select Case clase
        Case eClass.Sastre
            ModSastreria = 1
        Case Else
            ModSastreria = 4
    End Select

End Function

Function ModAlquimia(ByVal clase As eClass) As Integer

    Select Case clase
        Case eClass.Druid
            ModAlquimia = 1
        Case Else
            ModAlquimia = 4
    End Select

End Function

Function ModDomar(ByVal clase As eClass) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Select Case clase

        Case eClass.Druid
            ModDomar = 6

        Case eClass.Hunter
            ModDomar = 6

        Case eClass.Cleric
            ModDomar = 7

        Case Else
            ModDomar = 10

    End Select

End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: 02/03/09
    '02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
    '***************************************************
    Dim j As Integer

    For j = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasType(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function

        End If

    Next j

End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Nacho (Integer)
    'Last Modification: 01/05/2010
    '12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
    '02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
    '01/05/2010: ZaMa - Agrego bonificacion 11% para domar con flauta magica.
    '***************************************************

    On Error GoTo errHandler

    Dim puntosDomar      As Integer

    Dim puntosRequeridos As Integer

    Dim CanStay          As Boolean

    Dim petType          As Integer

    Dim NroPets          As Integer
    
    If Npclist(NPCIndex).MaestroUser = UserIndex Then
        Call WriteConsoleMsg(UserIndex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    With UserList(UserIndex)

        If .NroMascotas < MAXMASCOTAS Then
            
            If Npclist(NPCIndex).MaestroNpc > 0 Or Npclist(NPCIndex).MaestroUser > 0 Then
                Call WriteConsoleMsg(UserIndex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            If Not PuedeDomarMascota(UserIndex, NPCIndex) Then
                Call WriteConsoleMsg(UserIndex, "No puedes domar mas de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            puntosDomar = CInt(.Stats.UserAtributos(eAtributos.Carisma)) * CInt(.Stats.UserSkills(eSkill.Domar))
            puntosRequeridos = Npclist(NPCIndex).flags.Domable
            
            If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then

                Dim index As Integer

                .NroMascotas = .NroMascotas + 1
                index = FreeMascotaIndex(UserIndex)
                .MascotasIndex(index) = NPCIndex
                .MascotasType(index) = Npclist(NPCIndex).Numero
                
                Npclist(NPCIndex).MaestroUser = UserIndex
                
                Call FollowAmo(NPCIndex)
                Call ReSpawnNpc(Npclist(NPCIndex))
                
                Call WriteConsoleMsg(UserIndex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
                
                ' Es zona segura?
                CanStay = (MapInfo(.Pos.Map).Pk = True)
                
                If Not CanStay Then
                    petType = Npclist(NPCIndex).Numero
                    NroPets = .NroMascotas
                    
                    Call QuitarNPC(NPCIndex)
                    
                    .MascotasType(index) = petType
                    .NroMascotas = NroPets
                    
                    Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. estas te esperaran afuera.", FontTypeNames.FONTTYPE_INFO)

                End If
                
                Call SubirSkill(UserIndex, eSkill.Domar, True)
        
            Else

                If Not .flags.UltimoMensaje = 5 Then
                    Call WriteConsoleMsg(UserIndex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 5

                End If
                
                Call SubirSkill(UserIndex, eSkill.Domar, False)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "No puedes controlar mas criaturas.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal UserIndex As Integer, _
                                   ByVal NPCIndex As Integer) As Boolean

    '***************************************************
    'Author: ZaMa
    'This function checks how many NPCs of the same type have
    'been tamed by the user.
    'Returns True if that amount is less than two.
    '***************************************************
    Dim i           As Long

    Dim numMascotas As Long
    
    For i = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasType(i) = Npclist(NPCIndex).Numero Then
            numMascotas = numMascotas + 1

        End If

    Next i
    
    If numMascotas <= 1 Then PuedeDomarMascota = True
    
End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010 (ZaMa)
    'Makes an admin invisible o visible.
    '13/07/2009: ZaMa - Now invisible admins' chars are erased from all clients, except from themselves.
    '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
    '***************************************************
    
    Dim tempData As String
    
    With UserList(UserIndex)

        If .flags.AdminInvisible = 0 Then

            ' Sacamos el mimetizmo
            If .flags.Mimetizado = 1 Then
                .Char.body = .CharMimetizado.body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                .Counters.Mimetismo = 0
                .flags.Mimetizado = 0
                ' Se fue el efecto del mimetismo, puede ser atacado por npcs
                .flags.Ignorado = False

            End If
            
            'Guardamos el antiguo body y head
            .flags.OldBody = .Char.body
            .flags.OldHead = .Char.Head
            
            .flags.AdminInvisible = 1
            .flags.invisible = 1
            .flags.Oculto = 1
            
            ' Solo el admin sabe que se hace invi
            tempData = PrepareMessageSetInvisible(.Char.CharIndex, True)
            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(tempData)
            
            'Le mandamos el mensaje para que borre el personaje a los clientes que esten cerca
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
            
        Else
            .flags.AdminInvisible = 0
            .flags.invisible = 0
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            ' Solo el admin sabe que se hace visible
            tempData = PrepareMessageCharacterChange(.Char.body, .Char.Head, .Char.Heading, .Char.CharIndex, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, .Char.loops, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(tempData)
            
            tempData = PrepareMessageSetInvisible(.Char.CharIndex, False)
            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(tempData)
             
            'Le mandamos el mensaje para crear el personaje a los clientes que esten cerca
            Call MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y, True)

        End If

    End With
    
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, _
                        ByVal X As Integer, _
                        ByVal Y As Integer, _
                        ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Suerte    As Byte

    Dim Exito     As Byte

    Dim obj       As obj

    Dim posMadera As WorldPos

    If Not LegalPos(Map, X, Y) Then Exit Sub

    With posMadera
        .Map = Map
        .X = X
        .Y = Y

    End With

    If MapData(Map, X, Y).ObjInfo.ObjIndex <> 58 Then
        Call WriteConsoleMsg(UserIndex, "Necesitas clickear sobre lena para hacer ramitas.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If Distancia(posMadera, UserList(UserIndex).Pos) > 2 Then
        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If MapData(Map, X, Y).ObjInfo.Amount < 3 Then
        Call WriteConsoleMsg(UserIndex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    Dim SupervivenciaSkill As Byte

    SupervivenciaSkill = UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia)

    If SupervivenciaSkill < 6 Then
        Suerte = 3
    ElseIf SupervivenciaSkill <= 34 Then
        Suerte = 2
    Else
        Suerte = 1

    End If

    Exito = RandomNumber(1, Suerte)

    If Exito = 1 Then
        obj.ObjIndex = FOGATA_APAG
        obj.Amount = MapData(Map, X, Y).ObjInfo.Amount \ 3
    
        Call WriteConsoleMsg(UserIndex, "Has hecho " & obj.Amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
    
        Call MakeObj(obj, Map, X, Y)
    
        'Seteamos la fogata como el nuevo TargetObj del user
        UserList(UserIndex).flags.TargetObj = FOGATA_APAG
    
        Call SubirSkill(UserIndex, eSkill.Supervivencia, True)
    Else

        '[CDT 17-02-2004]
        If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
            Call WriteConsoleMsg(UserIndex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 10

        End If

        '[/CDT]
    
        Call SubirSkill(UserIndex, eSkill.Supervivencia, False)

    End If

End Sub

Public Sub DoPescar(ByVal UserIndex As Integer, ByVal Red As Boolean)

    '***************************************************
    'Author: Unknown
    'Last Modification: 26/10/2018
    '26/10/2018: CHOTS - Multiplicador de oficios
    '***************************************************
    On Error GoTo errHandler

    Dim iSkill        As Integer

    Dim Suerte        As Integer

    Dim res           As Integer

    Dim MAXITEMS      As Integer

    Dim CantidadItems As Integer

    With UserList(UserIndex)
    
        Call QuitarStaExtraccion(UserIndex, eSkill.Pesca)

        iSkill = .Stats.UserSkills(eSkill.Pesca)
        
        ' m = (60-11)/(1-10)
        ' y = mx - m*10 + 11
        
        Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 49)

        If Suerte > 0 Then
            res = RandomNumber(1, Suerte)
            
            If res <= DificultadExtraer Then
            
                Dim MiObj As obj
                
                MAXITEMS = MaxItemsExtraibles(.Stats.ELV)
                   
                If UserList(UserIndex).clase = eClass.Pescador Then
                    CantidadItems = RandomNumber(1, MAXITEMS)
                Else
                    CantidadItems = 1
                End If
                
                CantidadItems = CantidadItems * OficioMultiplier
                MiObj.Amount = CantidadItems
                
                If Red Then
                    MiObj.ObjIndex = ListaPeces(RandomNumber(1, NUM_PECES))
                Else
                    MiObj.ObjIndex = Pescado
                End If
                
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)

                End If
                
                Call WriteConsoleMsg(UserIndex, "Has pescado algunos peces!", FontTypeNames.FONTTYPE_INFO)
                
                Call SubirSkill(UserIndex, eSkill.Pesca, True)
            Else

                If Not .flags.UltimoMensaje = 6 Then
                    Call WriteConsoleMsg(UserIndex, "No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 6

                End If
                
                Call SubirSkill(UserIndex, eSkill.Pesca, False)

            End If

        End If
        
        .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

        If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
        
        'Sonido
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
        
        .Counters.Trabajando = .Counters.Trabajando + 1
    
    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en DoPescar Red: " & Red)

End Sub

Public Sub DoInstrumentos(ByVal UserIndex As Integer, ByRef ObjIndex As Integer)

    Dim Suerte As Double

    Dim res    As Integer

    Dim Skill  As Integer
    
    With UserList(UserIndex)
  
        If Not IntervaloPuedeTocar(UserIndex) Then Exit Sub

        Skill = .Stats.UserSkills(eSkill.Musica)
            
        Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
        
        If .clase = eClass.Bard Then Suerte = Suerte * 2
        
        res = RandomNumber(1, 100)
        
        If res <= Suerte Then
        
            If ObjData(ObjIndex).Real Then 'Es el Cuerno Real?
                If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                    If MapInfo(.Pos.Map).Pk = False Then
                        Call WriteConsoleMsg(UserIndex, "No hay peligro aqui. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
                            
                    ' Los admin invisibles solo producen sonidos a si mismos
                    If .flags.AdminInvisible = 1 Then
                        Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(ObjData(ObjIndex).Snd1, .Pos.X, .Pos.Y))
                    Else
                        Call AlertarFaccionarios(UserIndex)
                        Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(ObjData(ObjIndex).Snd1, .Pos.X, .Pos.Y))
    
                    End If
                            
                    Exit Sub
                Else
                    Call WriteConsoleMsg(UserIndex, "Solo miembros del ejercito real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
    
                End If
                
            ElseIf ObjData(ObjIndex).Caos Then 'Es el Cuerno Legion?
    
                If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                    If MapInfo(.Pos.Map).Pk = False Then
                        Call WriteConsoleMsg(UserIndex, "No hay peligro aqui. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
                            
                    ' Los admin invisibles solo producen sonidos a si mismos
                    If .flags.AdminInvisible = 1 Then
                        Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(ObjData(ObjIndex).Snd1, .Pos.X, .Pos.Y))
                    Else
                        Call AlertarFaccionarios(UserIndex)
                        Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(ObjData(ObjIndex).Snd1, .Pos.X, .Pos.Y))
    
                    End If
                            
                    Exit Sub
                Else
                    Call WriteConsoleMsg(UserIndex, "Solo miembros de la legion oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
    
                End If
    
            End If
    
            'Si llega aca es porque es o Laud o Tambor o Flauta
            ' Los admin invisibles solo producen sonidos a si mismos
            If .flags.AdminInvisible = 1 Then
                Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(ObjData(ObjIndex).Snd1, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(ObjData(ObjIndex).Snd1, .Pos.X, .Pos.Y))
    
            End If
            
            Call SubirSkill(UserIndex, eSkill.Musica, True)
            
        Else
        
            If Not .flags.UltimoMensaje = 4 Then
                Call WriteConsoleMsg(UserIndex, "No has logrado tocar el instrumento!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 4

            End If
        
            Call SubirSkill(UserIndex, eSkill.Musica, False)
            
        End If
                
    End With

End Sub

''
' Try to steal an item / gold to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
    '*************************************************
    'Author: Unknown
    'Last modified: 05/04/2010
    'Last Modification By: ZaMa
    '24/07/08: Marco - Now it calls to WriteUpdateGold(VictimaIndex and LadrOnIndex) when the thief stoles gold. (MarKoxX)
    '27/11/2009: ZaMa - Optimizacion de codigo.
    '18/12/2009: ZaMa - Los ladrones ciudas pueden robar a pks.
    '01/04/2010: ZaMa - Los ladrones pasan a robar oro acorde a su nivel.
    '05/04/2010: ZaMa - Los armadas no pueden robarle a ciudadanos jamas.
    '23/04/2010: ZaMa - No se puede robar mas sin energia.
    '23/04/2010: ZaMa - El alcance de robo pasa a ser de 1 tile.
    '*************************************************

    On Error GoTo errHandler

    Dim OtroUserIndex As Integer

    If Not MapInfo(UserList(VictimaIndex).Pos.Map).Pk Then Exit Sub
    
    If UserList(VictimaIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a usuarios en consulta!!!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    With UserList(LadrOnIndex)
    
        If .flags.Seguro Then
            If Not criminal(VictimaIndex) Then
                Call WriteConsoleMsg(LadrOnIndex, "Debes quitarte el seguro para robarle a un ciudadano.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

        Else

            If .Faccion.ArmadaReal = 1 Then
                If Not criminal(VictimaIndex) Then
                    Call WriteConsoleMsg(LadrOnIndex, "Los miembros del ejercito real no tienen permitido robarle a ciudadanos.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub

                End If

            End If

        End If
        
        ' Caos robando a caos?
        If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And .Faccion.FuerzasCaos = 1 Then
            Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a otros miembros de la legion oscura.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
        
        If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
        
        ' Tiene energia?
        If .Stats.MinSta < 15 Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(LadrOnIndex, "Estas muy cansado para robar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "Estas muy cansada para robar.", FontTypeNames.FONTTYPE_INFO)

            End If
            
            Exit Sub

        End If
        
        ' Quito energia
        Call QuitarSta(LadrOnIndex, 15)
        
        If UserList(VictimaIndex).flags.Privilegios And PlayerType.User Then
            
            Dim Suerte     As Integer

            Dim res        As Integer

            Dim RobarSkill As Byte
            
            RobarSkill = .Stats.UserSkills(eSkill.Robar)
                
            If RobarSkill <= 10 Then
                Suerte = 35
            ElseIf RobarSkill <= 20 Then
                Suerte = 30
            ElseIf RobarSkill <= 30 Then
                Suerte = 28
            ElseIf RobarSkill <= 40 Then
                Suerte = 24
            ElseIf RobarSkill <= 50 Then
                Suerte = 22
            ElseIf RobarSkill <= 60 Then
                Suerte = 20
            ElseIf RobarSkill <= 70 Then
                Suerte = 18
            ElseIf RobarSkill <= 80 Then
                Suerte = 15
            ElseIf RobarSkill <= 90 Then
                Suerte = 10
            ElseIf RobarSkill < 100 Then
                Suerte = 7
            Else
                Suerte = 5

            End If
            
            res = RandomNumber(1, Suerte)
                
            If res < 3 Then 'Exito robo
                If UserList(VictimaIndex).flags.Comerciando Then
                    OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                        
                    If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                        Call WriteConsoleMsg(VictimaIndex, "Comercio cancelado, te estan robando!!", FontTypeNames.FONTTYPE_TALK)
                        Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                        
                        Call LimpiarComercioSeguro(VictimaIndex)

                    End If

                End If
               
                If (RandomNumber(1, 50) < 25) And (.clase = eClass.Thief) Then
                    If TieneObjetosRobables(VictimaIndex) Then
                        Call RobarObjeto(LadrOnIndex, VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else 'Roba oro

                    If UserList(VictimaIndex).Stats.Gld > 0 Then

                        Dim n As Long
                        
                        If .clase = eClass.Thief Then
                            n = RandomNumber(.Stats.ELV * 25, .Stats.ELV * 50)

                        Else
                            n = RandomNumber(1, 100)

                        End If

                        If n > UserList(VictimaIndex).Stats.Gld Then n = UserList(VictimaIndex).Stats.Gld
                        UserList(VictimaIndex).Stats.Gld = UserList(VictimaIndex).Stats.Gld - n
                        
                        .Stats.Gld = .Stats.Gld + n

                        If .Stats.Gld > MAXORO Then .Stats.Gld = MAXORO
                        
                        Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & n & " monedas de oro a " & UserList(VictimaIndex).Name, FontTypeNames.FONTTYPE_INFO)
                        Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                        
                        Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
                
                Call SubirSkill(LadrOnIndex, eSkill.Robar, True)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(VictimaIndex, "" & .Name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
                
                Call SubirSkill(LadrOnIndex, eSkill.Robar, False)

            End If
        
            If Not criminal(LadrOnIndex) Then
                If Not criminal(VictimaIndex) Then
                    Call VolverCriminal(LadrOnIndex)

                End If

            End If
            
            ' Se pudo haber convertido si robo a un ciuda
            If criminal(LadrOnIndex) Then
                .Reputacion.LadronesRep = .Reputacion.LadronesRep + vlLadron

                If .Reputacion.LadronesRep > MAXREP Then .Reputacion.LadronesRep = MAXREP

            End If

        End If

    End With

    Exit Sub

errHandler:
    Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Check if one item is stealable
'
' @param VictimaIndex Specifies reference to victim
' @param Slot Specifies reference to victim's inventory slot
' @return If the item is stealable
Public Function ObjEsRobable(ByVal VictimaIndex As Integer, _
                             ByVal Slot As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    ' Agregue los barcos
    ' Esta funcion determina que objetos son robables.
    ' 22/05/2010: Los items newbies ya no son robables.
    '***************************************************

    Dim OI As Integer

    OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

    ObjEsRobable = ObjData(OI).OBJType <> eOBJType.otLlaves And UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And ObjData(OI).Real = 0 And ObjData(OI).Caos = 0 And ObjData(OI).OBJType <> eOBJType.otBarcos And ObjData(OI).OBJType <> eOBJType.otMonturas And ObjData(OI).NoRobable = 1 And Not ItemNewbie(OI)

End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen
Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/04/2010
    '02/04/2010: ZaMa - Modifico la cantidad de items robables por el ladron.
    '***************************************************

    Dim flag As Boolean

    Dim i    As Integer

    flag = False

    With UserList(VictimaIndex)

        If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
            i = 1

            Do While Not flag And i <= .CurrentInventorySlots

                'Hay objeto en este slot?
                If .Invent.Object(i).ObjIndex > 0 Then
                    If ObjEsRobable(VictimaIndex, i) Then
                        If RandomNumber(1, 10) < 4 Then flag = True

                    End If

                End If

                If Not flag Then i = i + 1
            Loop
        Else
            i = .CurrentInventorySlots

            Do While Not flag And i > 0

                'Hay objeto en este slot?
                If .Invent.Object(i).ObjIndex > 0 Then
                    If ObjEsRobable(VictimaIndex, i) Then
                        If RandomNumber(1, 10) < 4 Then flag = True

                    End If

                End If

                If Not flag Then i = i - 1
            Loop

        End If
    
        If flag Then

            Dim MiObj     As obj

            Dim Num       As Integer

            Dim ObjAmount As Integer
        
            ObjAmount = .Invent.Object(i).Amount
        
            'Cantidad al azar entre el 5% y el 10% del total, con minimo 1.
            Num = MaximoInt(1, RandomNumber(ObjAmount * 0.05, ObjAmount * 0.1))
                                    
            MiObj.Amount = Num
            MiObj.ObjIndex = .Invent.Object(i).ObjIndex
        
            .Invent.Object(i).Amount = ObjAmount - Num
                    
            If .Invent.Object(i).Amount <= 0 Then
                Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)

            End If
                
            Call UpdateUserInv(False, VictimaIndex, CByte(i))
                    
            If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)

            End If
        
            If UserList(LadrOnIndex).clase = eClass.Thief Then
                Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)

            End If

        Else
            Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningun objeto.", FontTypeNames.FONTTYPE_INFO)

        End If

        'If exiting, cancel de quien es robado
        Call CancelExit(VictimaIndex)
        
        'Si esta casteando, lo cancelamos
        Call CancelCast(VictimaIndex)

    End With

End Sub

Public Sub DoApunalar(ByVal UserIndex As Integer, _
                      ByVal VictimNpcIndex As Integer, _
                      ByVal VictimUserIndex As Integer, _
                      ByVal dano As Long)

    '***************************************************
    'Autor: Nacho (Integer) & Unknown (orginal version)
    'Last Modification: 04/17/08 - (NicoNZ)
    'Simplifique la cuenta que hacia para sacar la suerte
    'y arregle la cuenta que hacia para sacar el dano
    '***************************************************
    Dim Suerte As Integer

    Dim Skill  As Integer

    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apunalar)

    Select Case UserList(UserIndex).clase

        Case eClass.Assasin
            Suerte = Int(((0.00004 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
    
        Case eClass.Cleric, eClass.Paladin, eClass.Mercenario
            Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
    
        Case eClass.Bard
            Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
    
        Case Else
            Suerte = Int(0.0361 * Skill + 4.39)

    End Select

    If RandomNumber(0, 100) < Suerte Then
        If VictimUserIndex <> 0 Then
            If UserList(UserIndex).clase = eClass.Assasin Then
                dano = Round(dano * 1.4, 0)
            Else
                dano = Round(dano * 1.5, 0)

            End If
        
            With UserList(VictimUserIndex)
                .Stats.MinHp = .Stats.MinHp - dano
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(dano, UserList(UserIndex).Char.CharIndex, vbBlue, True))
                
                Call WriteConsoleMsg(UserIndex, "Has apunalado a " & .Name & " por " & dano, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimUserIndex, "Te ha apunalado " & UserList(UserIndex).Name & " por " & dano, FontTypeNames.FONTTYPE_FIGHT)

            End With
        
        Else
            
            With Npclist(VictimNpcIndex)
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Int(dano * 2), .Char.CharIndex, vbBlue, True))
                
                Call WriteConsoleMsg(UserIndex, "Has apunalado la criatura por " & Int(dano * 2), FontTypeNames.FONTTYPE_FIGHT)
                Call CalcularDarExp(UserIndex, VictimNpcIndex, dano * 2)
            
            End With

        End If
    
        Call SubirSkill(UserIndex, eSkill.Apunalar, True)
    Else
        Call WriteConsoleMsg(UserIndex, "No has logrado apunalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
        Call SubirSkill(UserIndex, eSkill.Apunalar, False)

    End If

End Sub

Public Sub DoAcuchillar(ByVal UserIndex As Integer, _
                        ByVal VictimNpcIndex As Integer, _
                        ByVal VictimUserIndex As Integer, _
                        ByVal dano As Integer)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 12/01/2010
    '***************************************************

    If RandomNumber(1, 100) <= PROB_ACUCHILLAR Then
        dano = Int(dano * DANO_ACUCHILLAR)
        
        If VictimUserIndex <> 0 Then
        
            With UserList(VictimUserIndex)
                .Stats.MinHp = .Stats.MinHp - dano
                Call WriteConsoleMsg(UserIndex, "Has acuchillado a " & .Name & " por " & dano, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha acuchillado por " & dano, FontTypeNames.FONTTYPE_FIGHT)

            End With
            
        Else
            With Npclist(VictimNpcIndex)
                
                Call WriteConsoleMsg(UserIndex, "Has acuchillado a la criatura por " & dano, FontTypeNames.FONTTYPE_FIGHT)
                Call CalcularDarExp(UserIndex, VictimNpcIndex, dano)
            End With
        End If

    End If
    
End Sub

Public Sub DoGolpeCritico(ByVal UserIndex As Integer, _
                          ByVal VictimNpcIndex As Integer, _
                          ByVal VictimUserIndex As Integer, _
                          ByVal dano As Long)

    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 28/01/2007
    '01/06/2010: ZaMa - Valido si tiene arma equipada antes de preguntar si es vikinga.
    '***************************************************
    Dim Suerte      As Integer

    Dim Skill       As Integer

    Dim WeaponIndex As Integer
    
    With UserList(UserIndex)

        ' Es bandido?
        If .clase <> eClass.Bandit Then Exit Sub
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
        
        ' Es una espada vikinga?
        If WeaponIndex <> ESPADA_VIKINGA Then Exit Sub
    
        Skill = .Stats.UserSkills(eSkill.Marciales)

    End With
    
    Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0893) * 100)
    
    If RandomNumber(1, 100) <= Suerte Then
    
        dano = Int(dano * 0.75)
        
        If VictimUserIndex <> 0 Then
            
            With UserList(VictimUserIndex)
                .Stats.MinHp = .Stats.MinHp - dano
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Int(dano * 2), UserList(UserIndex).Char.CharIndex, vbBlue, True))
                
                Call WriteConsoleMsg(UserIndex, "Has golpeado criticamente a " & .Name & " por " & dano & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha golpeado criticamente por " & dano & ".", FontTypeNames.FONTTYPE_FIGHT)

            End With
            
        Else
            
            With Npclist(VictimNpcIndex)
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Int(dano * 2), UserList(UserIndex).Char.CharIndex, vbBlue, True))
                
                Call WriteConsoleMsg(UserIndex, "Has golpeado criticamente a la criatura por " & dano & ".", FontTypeNames.FONTTYPE_FIGHT)
                
                Call CalcularDarExp(UserIndex, VictimNpcIndex, dano)
            End With
            
           
            
        End If
        
    End If

End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler
    
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad

    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateSta(UserIndex)
    
    Exit Sub

errHandler:
    Call LogError("Error en QuitarSta. Error " & Err.Number & " : " & Err.description)
    
End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With UserList(UserIndex)
        .Counters.IdleCount = 0
        
        Dim Suerte       As Integer

        Dim res          As Integer

        Dim cant         As Integer

        Dim MeditarSkill As Byte
    
        'Barrin 3/10/03
        'Esperamos a que se termine de concentrar
        Dim TActual      As Long

        TActual = GetTickCount() And &H7FFFFFFF

        If TActual - .Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
            Exit Sub

        End If
        
        If .Counters.bPuedeMeditar = False Then
            .Counters.bPuedeMeditar = True

        End If
            
        If .Stats.MinMAN >= .Stats.MaxMAN Then
            Call WriteConsoleMsg(UserIndex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMeditateToggle(UserIndex)
            .flags.Meditando = False
            .Char.Particle = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticleChar(.Char.CharIndex, .Char.Particle, False, 0))
            Exit Sub

        End If
        
        MeditarSkill = .Stats.UserSkills(eSkill.Meditar)
        
        If MeditarSkill <= 10 Then
            Suerte = 35
        ElseIf MeditarSkill <= 20 Then
            Suerte = 30
        ElseIf MeditarSkill <= 30 Then
            Suerte = 28
        ElseIf MeditarSkill <= 40 Then
            Suerte = 24
        ElseIf MeditarSkill <= 50 Then
            Suerte = 22
        ElseIf MeditarSkill <= 60 Then
            Suerte = 20
        ElseIf MeditarSkill <= 70 Then
            Suerte = 18
        ElseIf MeditarSkill <= 80 Then
            Suerte = 15
        ElseIf MeditarSkill <= 90 Then
            Suerte = 10
        ElseIf MeditarSkill < 100 Then
            Suerte = 7
        Else
            Suerte = 5

        End If
        
        If .Invent.AnilloEqpObjIndex <> 0 Then
            If ObjData(.Invent.AnilloEqpObjIndex).Efectomagico = eEfectos.AceleraMana Then
                Suerte = Suerte - Porcentaje(Suerte, 30)
            End If
        End If

        res = RandomNumber(1, Suerte)
        
        If res = 1 Then
            
            cant = Porcentaje(.Stats.MaxMAN, PorcentajeRecuperoMana)

            If cant <= 0 Then cant = 1
            .Stats.MinMAN = .Stats.MinMAN + cant

            If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
            
            Call WriteUpdateMana(UserIndex)
            Call SubirSkill(UserIndex, eSkill.Meditar, True)
        Else
            Call SubirSkill(UserIndex, eSkill.Meditar, False)

        End If

    End With

End Sub

Public Sub DoDesequipar(ByVal UserIndex As Integer, ByVal victimIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modif: 15/04/2010
    'Unequips either shield, weapon or helmet from target user.
    '***************************************************

    Dim Probabilidad   As Integer

    Dim Resultado      As Integer

    Dim MarcialesSkill As Byte

    Dim AlgoEquipado   As Boolean
    
    With UserList(UserIndex)
        
        ' Si no esta solo con manos, no desequipa tampoco.
        If .Invent.WeaponEqpObjIndex > 0 Then Exit Sub
        
        MarcialesSkill = .Stats.UserSkills(eSkill.Marciales)
        
        Probabilidad = MarcialesSkill * 0.2 + .Stats.ELV * 0.66

    End With
   
    With UserList(victimIndex)

        ' Si tiene escudo, intenta desequiparlo
        If .Invent.EscudoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(victimIndex, .Invent.EscudoEqpSlot)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el escudo de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(victimIndex, "Tu oponente te ha desequipado el escudo!", FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
        
        ' No tiene escudo, o fallo desequiparlo, entonces trata de desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(victimIndex, .Invent.WeaponEqpSlot)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(victimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
        
        ' No tiene arma, o fallo desequiparla, entonces trata de desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(victimIndex, .Invent.CascoEqpSlot)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el casco de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(victimIndex, "Tu oponente te ha desequipado el casco!", FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
    
        If AlgoEquipado Then
            Call WriteConsoleMsg(UserIndex, "Tu oponente no tiene equipado items!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "No has logrado desequipar ningun item a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
    End With

End Sub

Public Sub DoHurtar(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modif: 03/03/2010
    'Implements the pick pocket skill of the Bandit :)
    '03/03/2010 - Pato: Solo se puede hurtar si no esta en trigger 6 :)
    '***************************************************
    Dim OtroUserIndex As Integer

    If TriggerZonaPelea(UserIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

    If UserList(UserIndex).clase <> eClass.Bandit Then Exit Sub

    Dim res As Integer

    res = RandomNumber(1, 100)

    If (res < 20) Then
        If TieneObjetosRobables(VictimaIndex) Then
    
            If UserList(VictimaIndex).flags.Comerciando Then
                OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                
                If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                    Call WriteConsoleMsg(VictimaIndex, "Comercio cancelado, te estan robando!!", FontTypeNames.FONTTYPE_WARNING)
                    Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_WARNING)
                
                    Call LimpiarComercioSeguro(VictimaIndex)

                End If

            End If
                
            Call RobarObjeto(UserIndex, VictimaIndex)
            Call WriteConsoleMsg(VictimaIndex, "" & UserList(UserIndex).Name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

        End If

    End If

End Sub

Public Sub Desarmar(ByVal UserIndex As Integer, ByVal victimIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/04/2010 (ZaMa)
    '02/04/2010: ZaMa - Nueva formula para desarmar.
    '***************************************************

    Dim Probabilidad   As Integer

    Dim Resultado      As Integer

    Dim MarcialesSkill As Byte
    
    With UserList(UserIndex)
        MarcialesSkill = .Stats.UserSkills(eSkill.Marciales)
        
        Probabilidad = MarcialesSkill * 0.2 + .Stats.ELV * 0.66
        
        Resultado = RandomNumber(1, 100)
        
        If Resultado <= Probabilidad Then
            Call Desequipar(victimIndex, UserList(victimIndex).Invent.WeaponEqpSlot)
            Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)

            If UserList(victimIndex).Stats.ELV < 20 Then
                Call WriteConsoleMsg(victimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

            End If

        End If

    End With
    
End Sub

Public Function MaxItemsConstruibles(ByVal UserIndex As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/01/2010
    '11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
    '05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
    '***************************************************
    
    With UserList(UserIndex)

    MaxItemsConstruibles = MaximoInt(1, CInt((.Stats.ELV - 2) * 0.2))

    End With

End Function

Public Function MaxItemsExtraibles(ByVal UserLevel As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/05/2010
    '***************************************************
    MaxItemsExtraibles = MaximoInt(1, CInt((UserLevel - 2) * 0.2)) + 1

End Function

Public Sub ImitateNpc(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    'Copies body, head and desc from previously clicked npc.
    '***************************************************
    
    With UserList(UserIndex)
        
        ' Copy desc
        .DescRM = Npclist(NPCIndex).Name
        
        ' Remove Anims (Npcs don't use equipment anims yet)
        .Char.CascoAnim = NingunCasco
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        
        ' If admin is invisible the store it in old char
        If .flags.AdminInvisible = 1 Or .flags.invisible = 1 Or .flags.Oculto = 1 Then
            
            .flags.OldBody = Npclist(NPCIndex).Char.body
            .flags.OldHead = Npclist(NPCIndex).Char.Head
        Else
            .Char.body = Npclist(NPCIndex).Char.body
            .Char.Head = Npclist(NPCIndex).Char.Head
            
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)

        End If
    
    End With
    
End Sub

Public Sub DoEquita(ByVal UserIndex As Integer, _
                    ByRef Montura As ObjData, _
                    ByVal Slot As Integer)
    '***************************************************
    'Author: Recox
    'Last Modification: 06/04/2020
    'Podemos usar monturas ahora
    '06/04/2020: FrankoH298 - Ahora hay un timer para poder montarte
    '***************************************************

    With UserList(UserIndex)
    
        If UserList(UserIndex).Stats.UserSkills(Equitacion) < Montura.MinSkill Then
            Call WriteConsoleMsg(UserIndex, "Para usar esta montura necesitas " & Montura.MinSkill & " puntos en equitaci�n.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes utilizar la montura mientras estas muerto !!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If .flags.Navegando = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes utilizar la montura mientras navegas !!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.BAJOTECHO Or MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.CASA Then
            'TODO: SACAR ESTA VALIDACION DE ACA, Y HACER UN legalpos HAY TECHO en el cliente
            If .flags.Equitando = 0 Then Exit Sub

            Call WriteConsoleMsg(UserIndex, "No puedes utilizar la montura bajo techo!", FontTypeNames.FONTTYPE_INFO)
        End If

        ' If .flags.Metamorfosis = 1 Then 'Metamorfosis
        '     Call WriteConsoleMsg(UserIndex, "No puedes montar mientras estas metamorfoseado.", FontTypeNames.FONTTYPE_INFO)
        '     Exit Sub
        ' End If

        ' No estaba equitando
        If .flags.Equitando = 0 Then

            .Invent.MonturaObjIndex = .Invent.Object(Slot).ObjIndex
            .Invent.MonturaEqpSlot = Slot
    
            Call ToggleMonturaBody(UserIndex)
            Call SetVisibleStateForUserAfterNavigateOrEquitate(UserIndex)
    
            '  Comienza a equitar
            .flags.Equitando = 1
            
            If ObjData(.Invent.MonturaObjIndex).Speed > 0 Then
                .flags.Velocidad = ObjData(.Invent.MonturaObjIndex).Speed
            Else
                .flags.Velocidad = 2.4
            End If
            
            Call WriteSetSpeed(UserIndex)
                
            Call WriteEquitandoToggle(UserIndex)

            'Mostramos solo el casco de los items equipados por que los demas items quedan mal en el render, solo es un tema visual (Recox)
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
            
        ' Estaba equitando
        Else
            Call UnmountMontura(UserIndex)
            Call WriteEquitandoToggle(UserIndex)

        End If


    End With

End Sub

Public Sub UnmountMontura(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        .Invent.MonturaObjIndex = 0
        .Invent.MonturaEqpSlot = 0

        .Char.Head = .OrigChar.Head

        ' Seteamos el equipo que tiene y lo mostramos en el render.
        Call SetEquipmentOnCharAfterNavigateOrEquitate(UserIndex)
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
  
        ' Termina de equitar
        .flags.Equitando = 0
        .flags.Velocidad = SPEED_NORMAL
        Call WriteSetSpeed(UserIndex)

    End With
End Sub

Private Sub SetVisibleStateForUserAfterNavigateOrEquitate(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        ' Pierde el ocultar
        If .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .Counters.Ocultando = 0
            Call SetInvisible(UserIndex, .Char.CharIndex, False)
            Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
        End If

        ' Siempre se ve la montura (Nunca esta invisible), pero solo para el cliente.
        If .flags.invisible = 1 Then
            Call SetInvisible(UserIndex, .Char.CharIndex, False)
        End If

    End With

End Sub

Private Sub SetEquipmentOnCharAfterNavigateOrEquitate(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        If .Invent.ArmourEqpObjIndex > 0 Then
            .Char.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(UserIndex)

        End If
        
        If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim

        If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Invent.WeaponEqpObjIndex)

        If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
        
    End With


End Sub

Public Sub DoExtraer(ByVal UserIndex As Integer, ByVal Profesion As Integer)

    '***************************************************
    'Autor: Lorwik
    'Fecha: 19/08/2020
    'Descripci�n: Extrae recursos de forma pasiva
    '***************************************************
    
    On Error GoTo errHandler

    Dim Suerte        As Integer
    Dim res           As Integer
    Dim MAXITEMS      As Integer
    Dim CantidadItems As Integer
    Dim MiObj As obj
    
    With UserList(UserIndex)

        If .flags.TargetObj = 0 Then Exit Sub

        Call QuitarStaExtraccion(UserIndex, Profesion)

        Dim Skill As Integer

        Skill = .Stats.UserSkills(Profesion)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
        res = RandomNumber(1, Suerte)

        If res <= DificultadExtraer Then
        
            MAXITEMS = MaxItemsExtraibles(.Stats.ELV)
            CantidadItems = RandomNumber(1, MAXITEMS)
            CantidadItems = CantidadItems * OficioMultiplier

            Select Case Profesion
                Case eSkill.talar
                    If .clase = eClass.Lenador Then
                        MiObj.Amount = 5
                    Else
                        MiObj.Amount = 1
                    End If
                    MiObj.ObjIndex = Lena
                    Call WriteConsoleMsg(UserIndex, "Has extraido algo de le�a!", FontTypeNames.FONTTYPE_INFO)
                    
                Case eSkill.Mineria
                    If .clase = eClass.Minero Then
                        MiObj.Amount = 5
                    Else
                        MiObj.Amount = 1
                    End If
                    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObj).MineralIndex 'Cambiar esto a un target permanente.
                    Call WriteConsoleMsg(UserIndex, "Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
                    
                Case eSkill.Botanica
                    If .clase = eClass.Druid Then
                        MiObj.Amount = 5
                    Else
                        MiObj.Amount = 1
                    End If
                    MiObj.ObjIndex = Raices
                    Call WriteConsoleMsg(UserIndex, "Has extraido algunas raices!", FontTypeNames.FONTTYPE_INFO)
                
            End Select
       
            If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)

            Call SubirSkill(UserIndex, Profesion, True)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 9 Then
                Call WriteConsoleMsg(UserIndex, "No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 9

            End If

            '[/CDT]
            Call SubirSkill(UserIndex, Profesion, False)

        End If
    
        If Not criminal(UserIndex) Then
            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

            If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

        End If
    
        .Counters.Trabajando = .Counters.Trabajando + 1
        
        'Play sound!
        If Profesion = eSkill.Mineria Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_MINERO, .Pos.X, .Pos.Y))
            
        ElseIf Profesion = eSkill.talar Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
            
        End If

    End With

    Exit Sub

errHandler:
    Call LogError("Error en Sub DoExtraer")

End Sub

Private Sub QuitarStaExtraccion(ByVal UserIndex As Integer, ByVal Profesion As eSkill)

    Dim Esfuerzo As Integer

    With UserList(UserIndex)
    
        Select Case Profesion
            Case eSkill.talar
                If .clase = eClass.Lenador Then
                    Esfuerzo = EsfuerzoTalarLenador
                Else
                    Esfuerzo = EsfuerzoTalarGeneral
                End If
                    
            Case eSkill.Mineria
                If .clase = eClass.Minero Then
                    Esfuerzo = EsfuerzoExcavarMinero
                Else
                    Esfuerzo = EsfuerzoExcavarGeneral
                End If
                    
            Case eSkill.Botanica
                If .clase = eClass.Druid Then
                    Esfuerzo = EsfuerzoRaicesBotanico
                Else
                    Esfuerzo = EsfuerzoRaicesGeneral
                End If
                
            Case eSkill.Pesca
                If .clase = eClass.Pescador Then
                    Esfuerzo = EsfuerzoPescarPescador
                Else
                    Esfuerzo = EsfuerzoPescarGeneral
                End If
                
            Case Else
                Esfuerzo = 5
            
        End Select
    
    End With
    
     Call QuitarSta(UserIndex, Esfuerzo)
End Sub

Sub EnviarArmasConstruibles(ByVal UserIndex As Integer)
    Call WriteBlacksmithWeapons(UserIndex)

End Sub
 
Sub EnviarObjConstruibles(ByVal UserIndex As Integer)
    Call WriteCarpenterObjects(UserIndex)

End Sub

Sub EnviarArmadurasConstruibles(ByVal UserIndex As Integer)
    Call WriteBlacksmithArmors(UserIndex)

End Sub

Sub EnviarRopasConstruibles(ByVal UserIndex As Integer)
    Call WriteSastreRopas(UserIndex)

End Sub

Sub EnviarPocionesConstruibles(ByVal UserIndex As Integer)
    Call WriteAlquimistaPociones(UserIndex)

End Sub

