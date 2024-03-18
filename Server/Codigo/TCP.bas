Attribute VB_Name = "TCP"
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

#If False Then

    Dim errHandler, Length, index As Variant

#End If

Option Explicit

#If False Then

    Dim X, Y, n, Mapa, Email, Length As Variant
    
#End If

Private MAX_OBJ_INICIAL As Byte
Private ItemsIniciales() As UserObj

Sub DarCuerpo(ByVal UserIndex As Integer)

    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 14/03/2007
    'Elije una cabeza para el usuario y le da un body
    '*************************************************
    Dim NewBody    As Integer

    Dim UserRaza   As Byte

    Dim UserGenero As Byte

    UserGenero = UserList(UserIndex).Genero
    UserRaza = UserList(UserIndex).Raza

    Select Case UserGenero

        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano
                    NewBody = 1

                Case eRaza.Elfo
                    NewBody = 2

                Case eRaza.Drow
                    NewBody = 3

                Case eRaza.Enano
                    NewBody = 95

                Case eRaza.Gnomo
                    NewBody = 52
                    
                Case eRaza.Orco
                    NewBody = 251

            End Select

        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano
                    NewBody = 351

                Case eRaza.Elfo
                    NewBody = 352

                Case eRaza.Drow
                    NewBody = 353

                Case eRaza.Gnomo
                    NewBody = 138

                Case eRaza.Enano
                    NewBody = 138
                    
                Case eRaza.Orco
                    NewBody = 252

            End Select

    End Select

    UserList(UserIndex).Char.body = NewBody

End Sub

Private Function ValidarCabeza(ByVal UserRaza As Byte, _
                               ByVal UserGenero As Byte, _
                               ByVal Head As Integer) As Boolean

    Select Case UserGenero

        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano
                    ValidarCabeza = (Head >= HUMANO_H_PRIMER_CABEZA And Head <= HUMANO_H_ULTIMA_CABEZA)

                Case eRaza.Elfo
                    ValidarCabeza = (Head >= ELFO_H_PRIMER_CABEZA And Head <= ELFO_H_ULTIMA_CABEZA)

                Case eRaza.Drow
                    ValidarCabeza = (Head >= DROW_H_PRIMER_CABEZA And Head <= DROW_H_ULTIMA_CABEZA)

                Case eRaza.Enano
                    ValidarCabeza = (Head >= ENANO_H_PRIMER_CABEZA And Head <= ENANO_H_ULTIMA_CABEZA)

                Case eRaza.Gnomo
                    ValidarCabeza = (Head >= GNOMO_H_PRIMER_CABEZA And Head <= GNOMO_H_ULTIMA_CABEZA)
                    
                Case eRaza.Orco
                    ValidarCabeza = (Head >= ORCO_H_PRIMER_CABEZA And Head <= ORCO_H_ULTIMA_CABEZA)

            End Select
    
        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano
                    ValidarCabeza = (Head >= HUMANO_M_PRIMER_CABEZA And Head <= HUMANO_M_ULTIMA_CABEZA)

                Case eRaza.Elfo
                    ValidarCabeza = (Head >= ELFO_M_PRIMER_CABEZA And Head <= ELFO_M_ULTIMA_CABEZA)

                Case eRaza.Drow
                    ValidarCabeza = (Head >= DROW_M_PRIMER_CABEZA And Head <= DROW_M_ULTIMA_CABEZA)

                Case eRaza.Enano
                    ValidarCabeza = (Head >= ENANO_M_PRIMER_CABEZA And Head <= ENANO_M_ULTIMA_CABEZA)

                Case eRaza.Gnomo
                    ValidarCabeza = (Head >= GNOMO_M_PRIMER_CABEZA And Head <= GNOMO_M_ULTIMA_CABEZA)
                    
                Case eRaza.Orco
                    ValidarCabeza = (Head >= ORCO_M_PRIMER_CABEZA And Head <= ORCO_M_ULTIMA_CABEZA)

            End Select

    End Select
        
End Function

Function AsciiValidos(ByVal cad As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim car As Byte

    Dim i   As Integer

    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
    
        If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
            AsciiValidos = False
            Exit Function

        End If
    
    Next i

    AsciiValidos = True

End Function

Public Function CheckMailString(ByVal sString As String) As Boolean

    On Error GoTo errHnd

    Dim lPos As Long

    Dim lX   As Long

    Dim iAsc As Integer

    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")

    If (lPos <> 0) Then

        '2do test: Busca un simbolo . despues de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then Exit Function

        '3er test: Recorre todos los caracteres y los valida
        For lX = 0 To Len(sString) - 1

            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))

                If Not CMSValidateChar_(iAsc) Then Exit Function

            End If

        Next lX

        'Finale
        CheckMailString = True

    End If

errHnd:

End Function

'  Corregida por Maraxus para que reconozca como validas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)

End Function

Function Numeric(ByVal cad As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim car As Byte

    Dim i   As Integer

    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
    
        If (car < 48 Or car > 57) Then
            Numeric = False
            Exit Function

        End If
    
    Next i

    Numeric = True

End Function

Function NombrePermitido(ByVal Nombre As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer

    For i = 1 To UBound(ForbidenNames)

        If InStr(Nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function

        End If

    Next i

    NombrePermitido = True

End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim LoopC As Integer

    For LoopC = 1 To NUMSKILLS

        If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then
            Exit Function

            If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100

        End If

    Next LoopC

    ValidateSkills = True
    
End Function

Sub ConnectNewUser(ByVal UserIndex As Integer, _
                   ByRef Name As String, _
                   ByVal UserRaza As eRaza, _
                   ByVal UserSexo As eGenero, _
                   ByVal UserClase As eClass, _
                   ByVal Hogar As eCiudad, _
                   ByVal Head As Integer, _
                   ByRef skills() As Byte, _
                   ByRef PetName As String, _
                   ByRef PetTipo As Byte)

    '*************************************************
    'Author: Unknown
    'Last modified: 3/12/2009
    'Conecta un nuevo Usuario
    '23/01/2007 Pablo (ToxicWaste) - Agregue ResetFaccion al crear usuario
    '24/01/2007 Pablo (ToxicWaste) - Agregue el nuevo mana inicial de los magos.
    '12/02/2007 Pablo (ToxicWaste) - Puse + 1 de const al Elfo normal.
    '20/04/2007 Pablo (ToxicWaste) - Puse -1 de fuerza al Elfo.
    '09/01/2008 Pablo (ToxicWaste) - Ahora los modificadores de Raza se controlan desde Balance.dat
    '11/19/2009: Pato - Modifico la mana inicial del bandido.
    '03/12/2009: Budi - Optimizacion del codigo.
    '12/10/2018: CHOTS - Sistema de cuentas
    '*************************************************
    With UserList(UserIndex)
        Dim i As Byte
        Dim Count As Integer
        Dim Suma As Long
        
        '�Intentaron hackear los atributos?
        For i = 1 To NUMATRIBUTOS
            If .Stats.UserAtributos(i) > 18 Then
                Call WriteErrorMsg(UserIndex, "Error en la asignacion de atributos, vuelva a asignarlos.")
                Exit Sub
            End If
            
            'Vamos sumando todos los atributos para luego comprobarlos
            Suma = Suma + .Stats.UserAtributos(i)
            Debug.Print Suma
        Next i
        
        If Suma <> 70 Then
            Call WriteErrorMsg(UserIndex, "Error en la asignacion de atributos, vuelva a asignarlos.")
            Exit Sub
        End If
        
        If Not AsciiValidos(Name) Or LenB(Name) = 0 Then
            Call WriteErrorMsg(UserIndex, "Nombre invalido.")
            Exit Sub

        End If
        
        'Solo permitimos 1 espacio en los nombres
        For i = 1 To Len(Name)
            If mid(Name, i, 1) = Chr(32) Then Count = Count + 1
            
        Next i
        
        If Count > 1 Then
            Call WriteErrorMsg(UserIndex, "Nombre invalido.")
            Exit Sub
        End If
        
        'Capitalizamos el nombre
        Name = StrConv(Name, vbProperCase)
        
        If UserList(UserIndex).flags.UserLogged Then
            Call LogCheating("El usuario " & UserList(UserIndex).Name & " ha intentado crear a " & Name & " desde la IP " & UserList(UserIndex).IP)
        
            'Kick player ( and leave character inside :D )!
            Call CloseSocketSL(UserIndex)
            Call Cerrar_Usuario(UserIndex)
        
            Exit Sub

        End If
    
        'Existe el personaje?
        If PersonajeExiste(Name) Then
            Call WriteErrorMsg(UserIndex, "Ya existe el personaje.")
            Exit Sub

        End If
        
        'Tiro los dados antes de llegar aca??
        If .Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
            Call WriteErrorMsg(UserIndex, "Debe tirar los dados antes de poder crear un personaje.")
            Exit Sub

        End If
    
        If Not ValidarCabeza(UserRaza, UserSexo, Head) Then
            Call LogCheating("El usuario " & Name & " ha seleccionado la cabeza " & Head & " desde la IP " & .IP)
        
            Call WriteErrorMsg(UserIndex, "Cabeza invalida, elija una cabeza seleccionable.")
            Exit Sub

        End If

        'Prevenir hackeo de skills
        Suma = 0
        For i = 1 To NUMSKILLS
            UserList(UserIndex).Stats.UserSkills(i) = skills(i - 1)
            Suma = Suma + Abs(UserList(UserIndex).Stats.UserSkills(i))
        Next i
        
        'Ya asigno los 10 primeros skills
        .Counters.AsignedSkills = 10
        
        If Suma > 10 Then
            Call LogHackAttemp(UserList(UserIndex).Name & " intento hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If

        'Mandamos a crear el familiar
        If Not CreateFamiliarNewUser(UserIndex, UserClase, PetName, PetTipo) Then
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    
        .flags.Muerto = 0
        .flags.Escondido = 0
    
        .Reputacion.AsesinoRep = 0
        .Reputacion.BandidoRep = 0
        .Reputacion.BurguesRep = 0
        .Reputacion.LadronesRep = 0
        .Reputacion.NobleRep = 1000
        .Reputacion.PlebeRep = 30
    
        .Reputacion.Promedio = 30 / 6
    
        .Name = Name
        .clase = UserClase
        .Raza = UserRaza
        .Genero = UserSexo
        .Hogar = Hogar
        
        For i = 0 To 1
            .Profesion(i).Profesion = 0
        Next i

        '???????????????? ATRIBUTOS
        Call SetAttributesToNewUser(UserIndex, UserClase, UserRaza)

        'Primero agregamos los items, ya que en caso de que el nivel
        'Inicial sea mayor al de un newbie, los items se borran automaticamente.
        '???????????????? INVENTARIO
        Call AddItemsToNewUser(UserIndex, UserClase, UserRaza)

        If EstadisticasInicialesUsarConfiguracionPersonalizada Then
            Call SetAttributesCustomToNewUser(UserIndex)
        End If

        Call DarCuerpo(UserIndex)
        .Char.Heading = eHeading.SOUTH
        .Char.Head = Head
    
        .OrigChar = .Char
        
        'Comenzara en la isla Newbie
        .Pos = Ciudades(Hogar)
            
        'De primeras podra hablar por global
        .flags.Global = 1
        .Counters.LastGlobalMsg = INTERVALO_GLOBAL

        #If ConUpTime Then
            .LogOnTime = Now
            .UpTime = 0
        #End If

    End With

    'Valores Default de facciones al Activar nuevo usuario
    Call ResetFacciones(UserIndex)
    
    Call SaveUser(UserIndex)
  
    'Open User
    Call ConnectUser(UserIndex, Name, True)

End Sub

Private Sub SetAttributesCustomToNewUser(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        .Stats.Gld = CLng(val(GetVar(ConfigPath & "Server.ini", "ESTADISTICASINICIALESPJ", "Oro")))
        .Stats.Banco = CLng(val(GetVar(ConfigPath & "Server.ini", "ESTADISTICASINICIALESPJ", "Banco")))

        Dim InitialLevel, Experiencia As Long
        InitialLevel = val(GetVar(ConfigPath & "Server.ini", "ESTADISTICASINICIALESPJ", "Nivel"))
        
        Dim i As Long
        For i = 1 To InitialLevel
            If i <> InitialLevel Then
                .Stats.Exp = .Stats.ELU
                
                'Se creo el parametro opcional en la funcion CheckUserLevel
                'Ya que al crear pjs con nivel mayor a 40 la cantidad de datos enviados hacia el
                'WriteConsole hacia que explote la aplicacion, con este parche se evita eso.
                Call CheckUserLevel(UserIndex, False)
            End If
        Next i
        
        .Stats.SkillPts = 0
    End With

End Sub

Private Sub SetAttributesToNewUser(ByVal UserIndex As Integer, ByVal UserClase As eClass, ByVal UserRaza As eRaza)

    Dim hSlot As Byte

    With UserList(UserIndex)
        '[Pablo (Toxic Waste) 9/01/08]
        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + ModRaza(UserRaza).Fuerza
        .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + ModRaza(UserRaza).Agilidad
        .Stats.UserAtributos(eAtributos.Inteligencia) = .Stats.UserAtributos(eAtributos.Inteligencia) + ModRaza(UserRaza).Inteligencia
        .Stats.UserAtributos(eAtributos.Carisma) = .Stats.UserAtributos(eAtributos.Carisma) + ModRaza(UserRaza).Carisma
        .Stats.UserAtributos(eAtributos.Constitucion) = .Stats.UserAtributos(eAtributos.Constitucion) + ModRaza(UserRaza).Constitucion
        '[/Pablo (Toxic Waste)]
    
        Dim i As Long
        For i = 1 To NUMSKILLS
            '.Stats.UserSkills(i) = 0
            Call CheckEluSkill(UserIndex, i, True)
        Next i
    
        Dim MiInt As Long

        MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Constitucion) \ 3)
    
        .Stats.MaxHp = 15 + MiInt
        .Stats.MinHp = 15 + MiInt
    
        MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) \ 6)

        If MiInt = 1 Then MiInt = 2
    
        .Stats.MaxSta = 20 * MiInt
        .Stats.MinSta = 20 * MiInt
    
        .Stats.MaxAGU = 100
        .Stats.MinAGU = 100
    
        .Stats.MaxHam = 100
        .Stats.MinHam = 100
    
        '<-----------------MANA----------------------->
        If UserClase = eClass.Mage Then 'Cambio en mana inicial (ToxicWaste)
            MiInt = .Stats.UserAtributos(eAtributos.Inteligencia) * 3
            .Stats.MaxMAN = MiInt
            .Stats.MinMAN = MiInt
        ElseIf UserClase = eClass.Cleric Or _
               UserClase = eClass.Druid Or _
               UserClase = eClass.Bard Or _
               UserClase = eClass.Assasin Or _
               UserClase = eClass.Bandit Or _
               UserClase = eClass.Paladin Then
            .Stats.MaxMAN = 50
            .Stats.MinMAN = 50
        Else
            .Stats.MaxMAN = 0
            .Stats.MinMAN = 0
        End If
    
        hSlot = 1
    
        If UserClase = eClass.Cleric Or _
           UserClase = eClass.Druid Or _
           UserClase = eClass.Bard Or _
           UserClase = eClass.Assasin Or _
           UserClase = eClass.Bandit Or _
           UserClase = eClass.Paladin Or _
           UserClase = eClass.Mage Then

            .Stats.UserHechizos(hSlot) = 2 'Proyectil m�gico
            
            hSlot = hSlot + 1
        
        End If
        
        'Familiar
        If UserClase = eClass.Druid Or UserClase = eClass.Mage Or UserClase = eClass.Hunter Then _
            .Stats.UserHechizos(hSlot) = 59 'Llamado al familiar
    
        .Stats.MaxHit = 2
        .Stats.MinHIT = 1
    
        .Stats.Gld = 0
    
        .Stats.Exp = 0
        .Stats.ELV = 1
        .Stats.ELU = 300
    End With

End Sub

Private Sub AddItemsToNewUser(ByVal UserIndex As Integer, ByVal UserClase As eClass, ByVal UserRaza As eRaza)
'*************************************************
'Author: Lucas Recoaro (Recox)
'Last modified: 19/03/2019
'Agrega items al usuario recien creado
'*************************************************
    Dim Slot As Byte
    Dim IsPaladin As Boolean

    IsPaladin = UserClase = eClass.Paladin
    
    With UserList(UserIndex)
        'Pociones Rojas (Newbie)
        Slot = 1
        .Invent.Object(Slot).ObjIndex = 461
        .Invent.Object(Slot).Amount = 100

        'Pociones azules (Newbie)
        If .Stats.MaxMAN > 0 Or IsPaladin Then
            Slot = Slot + 1
            .Invent.Object(Slot).ObjIndex = 1598
            .Invent.Object(Slot).Amount = 100

        End If

        ' Ropa (Newbie)
        Slot = Slot + 1
        Select Case UserRaza
            Case eRaza.Humano
                If .Genero = eGenero.Hombre Then
                    .Invent.Object(Slot).ObjIndex = 463
                Else
                    .Invent.Object(Slot).ObjIndex = 1283
                End If
                
            Case eRaza.Elfo
                If .Genero = eGenero.Hombre Then
                    .Invent.Object(Slot).ObjIndex = 464
                Else
                    .Invent.Object(Slot).ObjIndex = 1284
                End If
                
            Case eRaza.Drow
                If .Genero = eGenero.Hombre Then
                    .Invent.Object(Slot).ObjIndex = 465
                Else
                    .Invent.Object(Slot).ObjIndex = 1285
                End If
                
            Case eRaza.Enano
                If .Genero = eGenero.Hombre Then
                    .Invent.Object(Slot).ObjIndex = 562
                Else
                    .Invent.Object(Slot).ObjIndex = 563
                End If
                
            Case eRaza.Gnomo
                If .Genero = eGenero.Hombre Then
                    .Invent.Object(Slot).ObjIndex = 466
                Else
                    .Invent.Object(Slot).ObjIndex = 563
                End If
                
            Case eRaza.Orco
                If .Genero = eGenero.Hombre Then
                    .Invent.Object(Slot).ObjIndex = 988
                Else
                    .Invent.Object(Slot).ObjIndex = 1087
                End If
        End Select

        ' Equipo ropa
        .Invent.Object(Slot).Amount = 1
        .Invent.Object(Slot).Equipped = 1
'
        .Invent.ArmourEqpSlot = Slot
        .Invent.ArmourEqpObjIndex = .Invent.Object(Slot).ObjIndex

        'Arma (Newbie)
        Slot = Slot + 1
        Select Case UserClase
            Case eClass.Paladin, eClass.Warrior, _
                 eClass.Herrero, eClass.Mercenario, _
                 eClass.Minero, eClass.Cleric, _
                 eClass.Lenador
                 
                'Martillo de Guerra (Newbie)
                .Invent.Object(Slot).ObjIndex = 574
                
            Case eClass.Gladiador, eClass.Bard
                'Nudillos de Hierro (Newbie)
                .Invent.Object(Slot).ObjIndex = 1354
                
            Case eClass.Assasin, eClass.Thief, _
                 eClass.Sastre, eClass.Pescador
                 'Daga (Newbie)
                .Invent.Object(Slot).ObjIndex = 460
                
            Case eClass.Mage, eClass.Nigromante
                'Baculo (Newbie)
                .Invent.Object(Slot).ObjIndex = 1356
                .Invent.Object(Slot).Amount = 1
                
            Case eClass.Hunter, eClass.Druid
                'Arco (Newbie)
                .Invent.Object(Slot).ObjIndex = 1355
                .Invent.Object(Slot).Amount = 1
                Slot = Slot + 1
                'Daga (Newbie)
                .Invent.Object(Slot).ObjIndex = 1356
                
            Case Else
                'Daga (Newbie)
                .Invent.Object(Slot).ObjIndex = 460
        End Select

        ' Equipo arma
        .Invent.Object(Slot).Amount = 1
        .Invent.Object(Slot).Equipped = 1

        'Todos menos el bardo usan "Weapon", el bardo usa "nudillos"
        If UserClase <> eClass.Bard Then
            .Invent.WeaponEqpObjIndex = .Invent.Object(Slot).ObjIndex
            .Invent.WeaponEqpSlot = Slot
            .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Invent.WeaponEqpObjIndex)
            
        Else
            .Invent.NudiEqpIndex = .Invent.Object(Slot).ObjIndex
            .Invent.NudiEqpSlot = Slot
            
        End If

        ' Municiones (Newbie)
        If UserClase = eClass.Hunter Then
            Slot = Slot + 1
            'arco
            .Invent.Object(Slot).ObjIndex = 1355
            .Invent.Object(Slot).Amount = 500

            ' Equipo flechas
            .Invent.Object(Slot).Equipped = 1
            .Invent.MunicionEqpSlot = Slot
            .Invent.MunicionEqpObjIndex = 860
            
        ElseIf UserClase = eClass.Thief Then
            ' Daga Arrojadizas (Newbie)
            .Invent.Object(Slot).ObjIndex = 576
            .Invent.Object(Slot).Amount = 1
            
        End If

        ' Manzanas (Newbie)
        Slot = Slot + 1
        .Invent.Object(Slot).ObjIndex = 573
        .Invent.Object(Slot).Amount = 100

        ' Agua (Nwbie)
        Slot = Slot + 1
        .Invent.Object(Slot).ObjIndex = 572
        .Invent.Object(Slot).Amount = 100
        
        'Runa
        Slot = Slot + 1
        .Invent.Object(7).ObjIndex = 1599
        .Invent.Object(7).Amount = 1
        
        Dim tmpObj As obj
        tmpObj.ObjIndex = 879
        tmpObj.Amount = 1
        Call MeterItemEnInventario(UserIndex, tmpObj) 'Mapa

        ' Sin casco y escudo
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco

        ' Total Items
        .Invent.NroItems = Slot

     End With
End Sub

Private Sub CargarObjetosIniciales()

    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    Call Leer.Initialize(ConfigPath & "Server.ini")

    Dim Slot As Long, sTemp As String

    MAX_OBJ_INICIAL = val(Leer.GetValue("INVENTARIO", "CantidadItemsIniciales"))

    ReDim ItemsIniciales(1 To MAX_OBJ_INICIAL) As UserObj

    For Slot = 1 To MAX_OBJ_INICIAL

        sTemp = Leer.GetValue("INVENTARIO", "Item" & Slot)

        ItemsIniciales(Slot).ObjIndex = val(ReadField(1, sTemp, 45))
        ItemsIniciales(Slot).Amount = val(ReadField(2, sTemp, 45))
        ItemsIniciales(Slot).Equipped = val(ReadField(3, sTemp, 45))

    Next Slot

    Set Leer = Nothing

End Sub

Sub ConnectAccount(ByVal UserIndex As Integer, _
                   ByRef UserName As String, _
                   ByRef Password As String)

'*************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last modified: 12/10/2018
'Se conecta a una cuenta existente
'*************************************************
'SHA256
    Dim oSHA256 As CSHA256

    Dim Salt    As String

    Set oSHA256 = New CSHA256

    If LenB(UserName) > 24 Or LenB(UserName) = 0 Then
        Call WriteErrorMsg(UserIndex, "Nombre invalido.")
        Exit Sub

    End If

    'Controlamos no pasar el maximo de usuarios
    If NumUsers >= MaxUsers Then
        Call WriteErrorMsg(UserIndex, "El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
        Call CloseSocket(UserIndex)
        Exit Sub

    End If

    '�Existe la cuenta?
    If Not CuentaExiste(UserName) Then
        Call WriteErrorMsg(UserIndex, "La cuenta no existe.")
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    'Ya esta conectado el personaje?
    If CheckForSameNameAccount(UserName) Then
        Call WriteErrorMsg(UserIndex, "La cuenta ya esta conectada.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
        
    'Aca Guardamos y Hasheamos el password + Salt
    'Es el passwd valido?
    Salt = GetAccountSalt(UserName) ' Obtenemos la Salt

    If oSHA256.SHA256(Password & Salt) <> GetAccountPassword(UserName) Then
        Call WriteErrorMsg(UserIndex, "Password incorrecto.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If

    '�La cuenta esta verificada?
    If Not CuentaVerificada(UserName) Then
        Call WriteErrorMsg(UserIndex, "La cuenta aun no ha sido verificada, por favor revise su email.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If

    'Aqui solo vamos a hacer un request a los endpoints de la aplicacion en Node.js
    'el repositorio para hacer funcionar esto, es este: https://github.com/ao-libre/ao-api-server
    'Si no tienen interes en usarlo pueden desactivarlo en el Server.ini
    If ConexionAPI Then
        'Pasamos UserName tambien como email, ya que son lo mismo.... :(
        Call ApiEndpointSendLoginAccountEmail(UserName)
    End If

    Call SaveAccountLastLoginDatabase(UserName, UserList(UserIndex).IP)
    Call LoginAccountDatabase(UserIndex, UserName)

End Sub

Sub CloseSocket(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 4/4/2020
    '4/4/2020: FrankoH298 - Flusheamos el buffer antes de cerrar el socket.
    '
    '***************************************************
    On Error GoTo errHandler
    
    Call FlushBuffer(UserIndex)
    
    With UserList(UserIndex)

        Call SecurityIp.IpRestarConexion(GetLongIp(.IP))
        
        If .ConnID <> -1 Then
            Call CloseSocketSL(UserIndex)
        End If
            
        'Empty buffer for reuse
        Call .incomingData.ReadASCIIStringFixed(.incomingData.Length)

        'Si llegamos aqui, sacamos al usuario de la cuenta y reseteamos todo
        If .flags.AccountLogged Then
            Call CloseAccount(UserIndex)
            
        ElseIf .flags.UserLogged Then 'Llego aqui estando logeado en un PJ?
            Call CloseUser(UserIndex)
            Call CloseAccount(UserIndex)
            
        Else
            Call ResetUserSlot(UserIndex)
            If NumCuentas > 0 Then NumCuentas = NumCuentas - 1
        End If
        
        Call LiberarSlot(UserIndex)
            
    End With

    Exit Sub

errHandler:

    Call ResetUserSlot(UserIndex)
        
    Call LiberarSlot(UserIndex)
        
    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripcion = " & Err.description & " - UserIndex = " & UserIndex)

End Sub

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
        Call SecurityIp.IpRestarConexion(GetLongIp(UserList(UserIndex).IP))
        Call BorraSlotSock(UserList(UserIndex).ConnID)
        Call WSApiCloseSocket(UserList(UserIndex).ConnID)
        UserList(UserIndex).ConnIDValida = False

    End If

End Sub

Function EstaPCarea(index As Integer, Index2 As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim X As Integer, Y As Integer

    For Y = UserList(index).Pos.Y - MinYBorder + 1 To UserList(index).Pos.Y + MinYBorder - 1
        For X = UserList(index).Pos.X - MinXBorder + 1 To UserList(index).Pos.X + MinXBorder - 1

            If MapData(UserList(index).Pos.Map, X, Y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function

            End If
        
        Next X
    Next Y

    EstaPCarea = False

End Function

Function HayPCarea(Pos As WorldPos) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim X As Integer, Y As Integer

    For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1

            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.Map, X, Y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function

                End If

            End If

        Next X
    Next Y

    HayPCarea = False

End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim X As Integer, Y As Integer

    For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1

            If MapData(Pos.Map, X, Y).ObjInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function

            End If
        
        Next X
    Next Y

    HayOBJarea = False

End Function

Function ValidateChr(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    ValidateChr = UserList(UserIndex).Char.Head <> 0 And UserList(UserIndex).Char.body <> 0 And ValidateSkills(UserIndex)

End Function

Sub ConnectUser(ByVal UserIndex As Integer, _
                ByRef Name As String, Optional ByVal NewUser As Boolean = False)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 24/07/2010 (ZaMa)
    '26/03/2009: ZaMa - Agrego por default que el color de dialogo de los dioses, sea como el de su nick.
    '12/06/2009: ZaMa - Agrego chequeo de nivel al loguear
    '14/09/2009: ZaMa - Ahora el usuario esta protegido del ataque de npcs al loguear
    '11/27/2009: Budi - Se envian los InvStats del personaje y su Fuerza y Agilidad
    '03/12/2009: Budi - Optimizacion del codigo
    '24/07/2010: ZaMa - La posicion de comienzo es namehuak, como se habia definido inicialmente.
    '12/10/2019: CHOTS - Sistema de cuentas
    '***************************************************
    Dim n    As Integer

    Dim tStr As String

    With UserList(UserIndex)

        If .flags.UserLogged Then
            Call LogCheating("El usuario " & .Name & " ha intentado loguear a " & Name & " desde la IP " & .IP)
            'Kick player ( and leave character inside :D )!
            Call CloseSocketSL(UserIndex)
            Call Cerrar_Usuario(UserIndex)
            Exit Sub

        End If
    
        'Reseteamos los FLAGS
        .flags.Escondido = 0
        .flags.TargetNPC = 0
        .flags.TargetNpcTipo = eNPCType.Comun
        .flags.TargetObj = 0
        .flags.TargetUser = 0
        .Char.FX = 0
        .Char.Particle = 0
    
        'Controlamos no pasar el maximo de usuarios
        If NumUsers >= MaxUsers Then
            Call WriteErrorMsg(UserIndex, "El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
            Call CloseUser(UserIndex)
            Exit Sub

        End If
    
        'Este IP ya esta conectado?
        If AllowMultiLogins = False Then
            If CheckForSameIP(UserIndex, .IP) = True Then
                Call WriteErrorMsg(UserIndex, "No es posible usar mas de un personaje al mismo tiempo.")
                Call CloseUser(UserIndex)
                Exit Sub

            End If

        End If
    
        'Existe el personaje? (Se comprueba aqui de nuevo por el CrearPJ)
        If Not PersonajeExiste(Name) Then
            Call WriteErrorMsg(UserIndex, "El personaje no existe.")
            Call CloseUser(UserIndex)
            Exit Sub

        End If
    
        'El personaje pertenece a la cuenta
        If Not PersonajePerteneceCuenta(UserIndex, Name) Then
            Call WriteErrorMsg(UserIndex, "El personaje al que intentas acceder no pertenece a tu cuenta.")
            Call CloseUser(UserIndex)
            Exit Sub

        End If
    
        'Ya esta conectado el personaje?
        If CheckForSameName(Name) Then
            If UserList(NameIndex(Name)).Counters.Saliendo Then
                Call WriteErrorMsg(UserIndex, "El usuario esta saliendo.")
            Else
                Call WriteErrorMsg(UserIndex, "Un usuario con el mismo nombre esta conectado.")
                Call Cerrar_Usuario(NameIndex(Name))
            End If
            Exit Sub
        End If
    
        'Reseteamos los privilegios
        .flags.Privilegios = 0
    
        'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
        If EsAdmin(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
            Call LogGM(Name, "Se conecto con ip:" & .IP)
        ElseIf EsDios(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
            Call LogGM(Name, "Se conecto con ip:" & .IP)
        ElseIf EsSemiDios(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
        
            .flags.PrivEspecial = EsGmEspecial(Name)
        
            Call LogGM(Name, "Se conecto con ip:" & .IP)
        ElseIf EsConsejero(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Consejero
            Call LogGM(Name, "Se conecto con ip:" & .IP)
        Else
            .flags.Privilegios = .flags.Privilegios Or PlayerType.User
            .flags.AdminPerseguible = True

        End If
    
        'Add RM flag if needed
        If EsRolesMaster(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster

        End If
    
        'Nombre de sistema
        .Name = Name
        
        'Seteamos el nombre de su cuenta en su char
        .AccountName = .AccountInfo.UserName
    
        'Load the user here
        Call LoadUser(UserIndex)

        If Not ValidateChr(UserIndex) Then
            Call WriteErrorMsg(UserIndex, "Error en el personaje.")
            Call CloseUser(UserIndex)
            Exit Sub

        End If
    
        If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
        If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
        If .Invent.WeaponEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
    
        .CurrentInventorySlots = MAX_INVENTORY_SLOTS
    
        Call UpdateUserInv(True, UserIndex, 0)
        Call UpdateUserHechizos(True, UserIndex, 0)
                                                                                            
        If .flags.Paralizado Then
            Call WriteParalizeOK(UserIndex)

        End If
    
        Dim Mapa As Integer

        Mapa = .Pos.Map
    
        '�Mapa invalido? Lo llevamos a Ulla
        If Mapa = 0 Then

            'Dejo esto comentado aqui por si se quiere utilizar la ciudad elegida desde el menu
            'Crear personaje, ahora se utiliza solo Ramx ya que es la ciudad inicial de Winter
             .Pos = Ciudades(.Hogar)
             Mapa = Ciudades(.Hogar).Map
        Else
    
            If Not MapaValido(Mapa) Then
                Call WriteErrorMsg(UserIndex, "El PJ se encuenta en un mapa invalido.")
                Call CloseUser(UserIndex)
                Exit Sub

            End If
        
            ' If map has different initial coords, update it
            Dim StartMap As Integer

            StartMap = MapInfo(Mapa).StartPos.Map

            If StartMap <> 0 Then
                If MapaValido(StartMap) Then
                    .Pos = MapInfo(Mapa).StartPos
                    Mapa = StartMap

                End If

            End If
        
        End If
    
        'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
        'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martin Sotuyo Dodero (Maraxus)
        If MapData(Mapa, .Pos.X, .Pos.Y).UserIndex <> 0 Or MapData(Mapa, .Pos.X, .Pos.Y).NpcIndex <> 0 Then

            Dim FoundPlace As Boolean

            Dim esAgua     As Boolean

            Dim tX         As Long

            Dim tY         As Long
        
            FoundPlace = False
            esAgua = HayAgua(Mapa, .Pos.X, .Pos.Y)
        
            For tY = .Pos.Y - 1 To .Pos.Y + 1
                For tX = .Pos.X - 1 To .Pos.X + 1

                    If esAgua Then

                        'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
                        If LegalPos(Mapa, tX, tY, True, False) Then
                            FoundPlace = True
                            Exit For

                        End If

                    Else

                        'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
                        If LegalPos(Mapa, tX, tY, False, True) Then
                            FoundPlace = True
                            Exit For

                        End If

                    End If

                Next tX
            
                If FoundPlace Then Exit For
            Next tY
        
            If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
                .Pos.X = tX
                .Pos.Y = tY
            Else

                'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
                If MapData(Mapa, .Pos.X, .Pos.Y).UserIndex <> 0 Then

                    'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                    If UserList(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu > 0 Then

                        'Le avisamos al que estaba comerciando que se tuvo que ir.
                        If UserList(UserList(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                            Call FinComerciarUsu(UserList(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)
                            Call WriteConsoleMsg(UserList(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                        'Lo sacamos.
                        If UserList(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex).flags.UserLogged Then
                            Call FinComerciarUsu(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex)
                            Call WriteErrorMsg(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconectate...")
                        End If

                    End If
                
                    Call CloseUser(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex)

                End If

            End If

        End If
    
        .showName = True 'Por default los nombres son visibles
    
        'If in the water, and has a boat, equip it!
        If .Invent.BarcoObjIndex > 0 And (HayAgua(Mapa, .Pos.X, .Pos.Y) Or BodyIsBoat(.Char.body)) Then

            .Char.Head = 0

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
        
            .flags.Navegando = 1

        End If
        
        'Seteamos la velocidad
        If .flags.Muerto = 1 Then
            .flags.Velocidad = SPEED_MUERTO
        Else
            .flags.Velocidad = SPEED_NORMAL
        End If
        
        Call WriteSetSpeed(UserIndex)
        
        'Actualizamos los seguros
        If .flags.ModoCombate Then
            Call WriteMultiMessage(UserIndex, eMessages.CombatSafeOn)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.CombatSafeOff)

        End If

        If .flags.Seguro Then
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOff)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn)

        End If
    
        'Info
        Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
        Call WriteChangeMap(UserIndex, .Pos.Map, MapInfo(.Pos.Map).MapVersion) 'Carga el mapa
        
        If .flags.Privilegios = PlayerType.Dios Then
            .flags.ChatColor = RGB(250, 250, 150)
        ElseIf .flags.Privilegios <> PlayerType.User And .flags.Privilegios <> (PlayerType.User Or PlayerType.ChaosCouncil) And .flags.Privilegios <> (PlayerType.User Or PlayerType.RoyalCouncil) Then
            .flags.ChatColor = RGB(0, 255, 0)
        ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.RoyalCouncil) Then
            .flags.ChatColor = RGB(0, 255, 255)
        ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.ChaosCouncil) Then
            .flags.ChatColor = RGB(255, 128, 64)
        Else
            .flags.ChatColor = vbWhite

        End If
    
        ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
        #If ConUpTime Then
            .LogOnTime = Now
        #End If
    
        'Crea  el personaje del usuario
        If (.flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster)) = 0 Then
            Call DoAdminInvisible(UserIndex)
            .flags.SendDenounces = True
        End If
        
        Call MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
    
        Call WriteUserCharIndexInServer(UserIndex)
        ''[/el oso]
    
        Call DoTileEvents(UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
    
        Call CheckUserLevel(UserIndex)
        Call WriteUpdateUserStats(UserIndex)
        
        Call WriteEnviarMacros(UserIndex)
    
        Call WriteUpdateHungerAndThirst(UserIndex)
        Call WriteUpdateStrenghtAndDexterity(UserIndex)
        
        Call SendMOTD(UserIndex)
    
        If haciendoBK Then
            Call WritePauseToggle(UserIndex)
            Call WriteConsoleMsg(UserIndex, "Servidor> Por favor espera algunos segundos, el WorldSave esta ejecutandose.", FontTypeNames.FONTTYPE_SERVER)

        End If
    
        If EnPausa Then
            Call WritePauseToggle(UserIndex)
            Call WriteConsoleMsg(UserIndex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar mas tarde.", FontTypeNames.FONTTYPE_SERVER)

        End If
    
        If EnTesting And .Stats.ELV >= 18 Then
            Call WriteErrorMsg(UserIndex, "Servidor en Testing por unos minutos, conectese con PJs de nivel menor a 18. No se conecte con Pjs que puedan resultar importantes por ahora pues pueden arruinarse.")
            Call CloseUser(UserIndex)
            Exit Sub

        End If
    
        'Actualiza el Num de usuarios
        'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
        NumUsers = NumUsers + 1
        .flags.UserLogged = True
    
        'usado para borrar Pjs
        Call UpdateUserLogged(.Name, 1)
    
        If .Stats.SkillPts > 0 Then
            Call WriteSendSkills(UserIndex)
            Call WriteLevelUp(UserIndex, .Stats.SkillPts)

        End If
    
        MapInfo(.Pos.Map).NumUsers = MapInfo(.Pos.Map).NumUsers + 1
    
        If NumUsers > RecordUsuariosOnline Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultaneamente. Hay " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFOBOLD))
            RecordUsuariosOnline = NumUsers
            Call WriteVar(ConfigPath & "Server.ini", "INIT", "RECORD", Str(RecordUsuariosOnline))

            'Este ultimo es para saber siempre los records en el frmMain
            frmMain.txtRecordOnline.Text = RecordUsuariosOnline
        End If
    
        If .NroMascotas > 0 And MapInfo(.Pos.Map).Pk Then

            Dim i As Integer

            For i = 1 To MAXMASCOTAS

                If .MascotasType(i) > 0 Then
                    .MascotasIndex(i) = SpawnNpc(.MascotasType(i), .Pos, True, True)
                
                    If .MascotasIndex(i) > 0 Then
                        Npclist(.MascotasIndex(i)).MaestroUser = UserIndex
                        Call FollowAmo(.MascotasIndex(i))
                    Else
                        .MascotasIndex(i) = 0

                    End If

                End If

            Next i

        End If

        If .flags.Navegando = 1 Then
            Call WriteNavigateToggle(UserIndex)

        End If
    
        If .GuildIndex > 0 Then

            'welcome to the show baby...
            If Not modGuilds.m_ConectarMiembroAClan(UserIndex, .GuildIndex) Then
                Call WriteConsoleMsg(UserIndex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)

            End If

        End If
        
        '****A�adimos al user nuevo a la lista de PJ*****
        If NewUser Then _
            Call AddNewPJCuenta(UserIndex)

    
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
    
        Call WriteLoggedMessage(UserIndex)
    
        ' Esta protegido del ataque de npcs por 5 segundos, si no realiza ninguna accion
        Call IntervaloPermiteSerAtacado(UserIndex, True)
    
        'Estado del dia y del clima
        Call WriteActualizarClima(UserIndex)
    
        tStr = modGuilds.a_ObtenerRechazoDeChar(.Name)
    
        If LenB(tStr) <> 0 Then
            Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)

        End If
    
        'Load the user statistics
        Call Statistics.UserConnected(UserIndex)
    
        Call MostrarNumUsers
        
        Call modGuilds.SendGuildNews(UserIndex)

        n = FreeFile
        Open App.Path & "\logs\numusers.log" For Output As n
        Print #n, NumUsers
        Close #n
    
        n = FreeFile
        'Log
        Open App.Path & "\logs\Connect.log" For Append Shared As #n
        Print #n, .Name & " ha entrado al juego. UserIndex:" & UserIndex & " " & time & " " & Date
        Close #n

    End With

End Sub

Sub SendMOTD(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: Puse devuelta esto con WriteGuildChat (Recox)
    '
    '***************************************************

    Dim j As Long
    
    Call WriteGuildChat(UserIndex, "Mensajes de entrada:")

    For j = 1 To MaxLines
        Call WriteGuildChat(UserIndex, MOTD(j).texto)
    Next j

End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)

    '*************************************************
    'Author: Unknown
    'Last modified: 23/01/2007
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
    '*************************************************
    With UserList(UserIndex).Faccion
        .ArmadaReal = 0
        .CiudadanosMatados = 0
        .CriminalesMatados = 0
        .FuerzasCaos = 0
        .FechaIngreso = vbNullString
        .RecibioArmaduraCaos = 0
        .RecibioArmaduraReal = 0
        .RecibioExpInicialCaos = 0
        .RecibioExpInicialReal = 0
        .RecompensasCaos = 0
        .RecompensasReal = 0
        .Reenlistadas = 0
        .NivelIngreso = 0
        .MatadosIngreso = 0
        .NextRecompensa = 0

    End With

End Sub

Sub ResetContadores(ByVal UserIndex As Integer)

    '*************************************************
    'Author: Unknown
    'Last modified: 10/07/2010
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '05/20/2007 Integer - Agregue todas las variables que faltaban.
    '10/07/2010: ZaMa - Agrego los counters que faltaban.
    '*************************************************
    With UserList(UserIndex).Counters
        .TimeFight = 0
        .AGUACounter = 0
        .AsignedSkills = 0
        .AttackCounter = 0
        .bPuedeMeditar = True
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .failedUsageAttempts = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Lava = 0
        .Mimetismo = 0
        .Ocultando = 0
        .Paralisis = 0
        .Pena = 0
        .PiqueteC = 0
        .Saliendo = False
        .Salir = 0
        .STACounter = 0
        .TiempoOculto = 0
        .TimerEstadoAtacable = 0
        .TimerPuedeOcultar = 0
        .TimerPuedeTocar = 0
        .TimerGolpeMagia = 0
        .TimerGolpeUsar = 0
        .TimerLanzarSpell = 0
        .TimerMagiaGolpe = 0
        .TimerPerteneceNpc = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeSerAtacado = 0
        .TimerPuedeTrabajar = 0
        .TimerPuedeUsarArco = 0
        .TimerUsar = 0
        .MacroTrabajo = 0
        .Trabajando = 0
        .Veneno = 0

    End With
    
    Call modAntiCheat.ResetAllCount(UserIndex)

End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)

    '*************************************************
    'Author: Unknown
    'Last modified: 03/15/2006
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '*************************************************
    With UserList(UserIndex).Char
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Particle = 0
        .Head = 0
        .loops = 0
        .Heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
        .AuraAnim = 0
        .AuraColor = 0
    End With

End Sub

Sub ResetUserAccount(ByVal UserIndex As Integer)
'*****************************************
'Autor: lorwik
'Fecha: 13/09/2020
'Descripcion: Borramos todos los datos almacenados de una cuenta
'*****************************************
    Dim i As Byte

    With UserList(UserIndex)
    
        'Borro la informaci�n de la cuenta
        .AccountInfo.Id = 0
        .AccountInfo.UserName = vbNullString
        .AccountInfo.Password = vbNullString
        .AccountInfo.Salt = vbNullString
        .AccountInfo.creditos = 0
        .AccountInfo.status = False
        
        For i = 1 To .AccountInfo.NumPjs
            Call ResetPJAccountSlot(UserIndex, i)
        Next i
    
    End With
End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)

    '*************************************************
    'Author: Unknown
    'Last modified: 03/15/2006
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '*************************************************
    With UserList(UserIndex)
        .Name = vbNullString
        .Id = 0
        .AccountName = vbNullString
        .Desc = vbNullString
        .DescRM = vbNullString
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .clase = 0
        .Email = vbNullString
        .Genero = 0
        .Hogar = 0
        .Raza = 0
        
        .PartyIndex = 0
        .PartySolicitud = 0
        .FormandoGrupo = 0
        
        .CreoPortal = False
        .PortalPos.Map = 0
        .PortalPos.X = 0
        .PortalPos.Y = 0
        .PortalTiempo = 0
        
        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            .NPCsMuertos = 0
            .Muertes = 0
            .SkillPts = 0
            .Gld = 0
            .UserAtributos(1) = 0
            .UserAtributos(2) = 0
            .UserAtributos(3) = 0
            .UserAtributos(4) = 0
            .UserAtributos(5) = 0
            .UserAtributosBackUP(1) = 0
            .UserAtributosBackUP(2) = 0
            .UserAtributosBackUP(3) = 0
            .UserAtributosBackUP(4) = 0
            .UserAtributosBackUP(5) = 0

        End With
        
    End With

End Sub

Sub ResetReputacion(ByVal UserIndex As Integer)

    '*************************************************
    'Author: Unknown
    'Last modified: 03/15/2006
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '*************************************************
    With UserList(UserIndex).Reputacion
        .AsesinoRep = 0
        .BandidoRep = 0
        .BurguesRep = 0
        .LadronesRep = 0
        .NobleRep = 0
        .PlebeRep = 0
        .NobleRep = 0
        .Promedio = 0

    End With

End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If UserList(UserIndex).EscucheClan > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(UserIndex, UserList(UserIndex).EscucheClan)
        UserList(UserIndex).EscucheClan = 0

    End If

    If UserList(UserIndex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(UserIndex, UserList(UserIndex).GuildIndex)

    End If

    UserList(UserIndex).GuildIndex = 0

End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)

    '*************************************************
    'Author: Unknown
    'Last modified: 06/28/2008
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '06/28/2008 NicoNZ - Agrego el flag Inmovilizado
    '*************************************************
    With UserList(UserIndex).flags
        .SlotReto = 0
        .SlotRetoUser = 255
        .Comerciando = False
        .SlotCarcel = 0
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .StatsChanged = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Descuento = vbNullString
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .Vuela = 0
        .Navegando = 0
        .Equitando = 0
        .Oculto = 0
        .Envenenado = 0
        .Incinerado = 0
        .invisible = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .Privilegios = 0
        .PrivEspecial = False
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .Silenciado = 0
        .AdminPerseguible = False
        .MacroTrabajo = 0
        .lastMap = 0
        .AtacablePor = 0
        .AtacadoPorNpc = 0
        .AtacadoPorUser = 0
        .NoPuedeSerAtacado = False
        .ShareNpcWith = 0
        .EnConsulta = False
        .Ignorado = False
        .SendDenounces = False
        .ParalizedBy = vbNullString
        .ParalizedByIndex = 0
        .ParalizedByNpcIndex = 0
        .Global = 0
        .Subastando = False
        .ProfInstruyendo = 0
        .Instruyendo = 0
        .Trabajando = 0
        .Velocidad = 0

        Call ResetCasteo(UserIndex)
        
        If .OwnedNpc <> 0 Then
            Call PerdioNpc(UserIndex)

        End If
        
    End With
    
End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim LoopC As Long

    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
    Next LoopC

End Sub

Sub ResetUserProfesion(ByVal UserIndex As Integer)
    '***************************************************
    'Autor: Lorwik
    'Last Modification: 21/08/2020
    '
    '***************************************************
    
    Dim i As Byte
    Dim j As Integer
    
    For i = 0 To 1
    
        For j = 1 To MAXUSERRECETAS
            UserList(UserIndex).Profesion(i).Recetas(j) = 0
        Next j
        
    Next i

End Sub

Sub ResetUserPets(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim LoopC As Long
    
    UserList(UserIndex).NroMascotas = 0
        
    For LoopC = 1 To MAXMASCOTAS
        UserList(UserIndex).MascotasIndex(LoopC) = 0
        UserList(UserIndex).MascotasType(LoopC) = 0
    Next LoopC

End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim LoopC As Long
    
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
        UserList(UserIndex).BancoInvent.Object(LoopC).Amount = 0
        UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
        UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = 0
    Next LoopC
    
    UserList(UserIndex).BancoInvent.NroItems = 0

End Sub

Sub ResetMacros(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/03/2021
    '
    '***************************************************
    
    Dim i As Byte
    
    With UserList(UserIndex)
    
        For i = 1 To NUMMACROS
            .MacrosKey(i).TipoAccion = 0
            .MacrosKey(i).hList = 0
            .MacrosKey(i).InvObj = 0
            .MacrosKey(i).Comando = ""
        Next i
    
    End With
    
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With UserList(UserIndex).ComUsu

        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(UserIndex)

        End If

    End With

End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Long

    Call LimpiarComercioSeguro(UserIndex)
    Call ResetFacciones(UserIndex)
    Call ResetContadores(UserIndex)
    Call ResetGuildInfo(UserIndex)
    Call ResetCharInfo(UserIndex)
    Call ResetBasicUserInfo(UserIndex)
    Call ResetReputacion(UserIndex)
    Call ResetUserFlags(UserIndex)
    Call LimpiarInventario(UserIndex)
    Call ResetUserSpells(UserIndex)
    Call ResetUserProfesion(UserIndex)
    Call ResetUserPets(UserIndex)
    Call ResetUserBanco(UserIndex)
    Call ResetMacros(UserIndex)
    
    With UserList(UserIndex).ComUsu
        .Acepto = False
    
        For i = 1 To MAX_OFFER_SLOTS
            .cant(i) = 0
            .Objeto(i) = 0
        Next i
    
        .GoldAmount = 0
        .DestNick = vbNullString
        .DestUsu = 0

    End With
 
End Sub

Sub CloseUser(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Dim n    As Integer

    Dim Map  As Integer

    Dim Name As String

    Dim i    As Integer

    Dim aN   As Integer

    With UserList(UserIndex)
    
        'Subastas
        Call Revisar_Subasta(UserIndex)
    
        'Nuevo centinela - maTih.-
        If .CentinelaUsuario.centinelaIndex <> 0 Then
            Call modCentinela.UsuarioInActivo(UserIndex)
        End If
        
        'mato los comercios seguros
        If .ComUsu.DestUsu > 0 Then
            
            If UserList(.ComUsu.DestUsu).flags.UserLogged Then
                
                If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_WARNING)
                    Call FinComerciarUsu(.ComUsu.DestUsu)
                End If

            End If

        End If
            
        ' Retos nVSn. Usuario cierra conexion.
        If .flags.SlotReto > 0 Then
            Call Retos.UserDieFight(UserIndex, 0, True)
        End If

        ' Desequipamos la montura justo antes de cerrar el socket
        ' para prevenir que se la equipe durante el conteo de salida (WyroX)
        If .flags.Equitando = 1 Then
            Call UnmountMontura(UserIndex)
        End If
    
        aN = .flags.AtacadoPorNpc

        If aN > 0 Then
            Npclist(aN).Movement = Npclist(aN).flags.OldMovement
            Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
            Npclist(aN).flags.AttackedBy = vbNullString

        End If
    
        aN = .flags.NPCAtacado

        If aN > 0 Then
            If Npclist(aN).flags.AttackedFirstBy = .Name Then
                Npclist(aN).flags.AttackedFirstBy = vbNullString

            End If

        End If

        .flags.AtacadoPorNpc = 0
        .flags.NPCAtacado = 0
    
        Map = .Pos.Map
        Name = UCase$(.Name)
    
        .Char.FX = 0
        .Char.loops = 0
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
        
        .Char.Particle = 0
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateParticleChar(.Char.CharIndex, .Char.Particle, False, 0))
        
        If NumUsers > 0 Then NumUsers = NumUsers - 1
        .flags.UserLogged = False
        .Counters.Saliendo = False
    
        'Le devolvemos el body y head originales
        If .flags.AdminInvisible = 1 Then
            .Char.body = .flags.OldBody
            .Char.Head = .flags.OldHead
            .flags.AdminInvisible = 0

        End If

        'si esta en party le devolvemos la experiencia
        If .PartyIndex > 0 Then Call mdParty.SalirDeParty(UserIndex)
        
        'Si creo un portal, lo borramos
        If .CreoPortal = True Then Call Borrar_Portal_User(UserIndex)
        
        'Save statistics
        Call Statistics.UserDisconnected(UserIndex)
    
        ' Grabamos el personaje del usuario
        Call SaveUser(UserIndex)
    
        'usado para borrar Pjs
        Call UpdateUserLogged(.Name, 0)
        
        'Quitar el dialogo
        'If MapInfo(Map).NumUsers > 0 Then
        '    Call SendToUserArea(UserIndex, "QDL" & .Char.charindex)
        'End If
    
        If Map > 0 Then
            If MapInfo(Map).NumUsers > 0 Then _
                Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))

        End If
    
        'Borrar el personaje
        If .Char.CharIndex > 0 Then
            Call EraseUserChar(UserIndex, .flags.AdminInvisible = 1)

        End If
    
        'Borrar mascotas
        For i = 1 To MAXMASCOTAS

            If .MascotasIndex(i) > 0 Then
                If Npclist(.MascotasIndex(i)).flags.NPCActive Then Call QuitarNPC(.MascotasIndex(i))

            End If

        Next i
        
        'Borramos el familiar
        Call ResetFamiliar(UserIndex)
    
        'Update Map Users
        If Map > 0 Then
            MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1
    
            If MapInfo(Map).NumUsers < 0 Then
                MapInfo(Map).NumUsers = 0
            End If
            
        End If
    
        ' Si el usuario habia dejado un msg en la gm's queue lo borramos
        If Ayuda.Existe(.Name) Then Call Ayuda.Quitar(.Name)
    
        Call ActualizarPJCuentas(UserIndex)
    
        Call ResetUserSlot(UserIndex)
    
        Call MostrarNumUsers

        n = FreeFile(1)
        Open App.Path & "\logs\Connect.log" For Append Shared As #n
        Print #n, Name & " ha dejado el juego. " & "User Index:" & UserIndex & " " & time & " " & Date
        Close #n

    End With

    Exit Sub

errHandler:
    Call LogError("Error en CloseUser. Numero " & Err.Number & " Descripcion: " & Err.description)

End Sub

Sub ReloadSokcet()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    Else
        'Call apiclosesocket(SockListen)
        'SockListen = ListenForConnect(Puerto, hWndMsg, vbNullString)
    End If

    Exit Sub
    
errHandler:
    Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.description)

End Sub

Public Sub EcharPjsNoPrivilegiados()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim LoopC As Long
    
    For LoopC = 1 To LastUser

        If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
            If UserList(LoopC).flags.Privilegios And PlayerType.User Then
                Call CloseUser(LoopC)

            End If

        End If

    Next LoopC

End Sub

Function RandomString(cb As Integer) As String

    Randomize

    Dim rgch As String

    rgch = "abcdefghijklmnopqrstuvwxyz"
    rgch = rgch & UCase(rgch) & "0123456789" & "#@!~$()-_"

    Dim i As Long

    For i = 1 To cb
        RandomString = RandomString & mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next

End Function
