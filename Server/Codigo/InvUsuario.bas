Attribute VB_Name = "InvUsuario"
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

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    ' 22/05/2010: Los items newbies ya no son robables.
    '***************************************************

    '17/09/02
    'Agregue que la funcion se asegure que el objeto no es un barco

    On Error GoTo errHandler

    Dim i        As Integer

    Dim ObjIndex As Integer
    
    For i = 1 To UserList(UserIndex).CurrentInventorySlots
        ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex

        If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And ObjData(ObjIndex).OBJType <> eOBJType.otBarcos And Not ItemNewbie(ObjIndex)) Then
                TieneObjetosRobables = True
                Exit Function

            End If

        End If

    Next i
    
    Exit Function

errHandler:
    Call LogError("Error en TieneObjetosRobables. Error: " & Err.Number & " - " & Err.description)

End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, _
                            ByVal ObjIndex As Integer, _
                            Optional ByRef sMotivo As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 01/04/2019
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '08/08/2015: Shak - Hechizos por clase
    '01/04/2019: Recox - Se arreglo la prohibicion de hechizos por clase
    '***************************************************

    On Error GoTo manejador
  
    'Admins can use ANYTHING!
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then

            Dim i As Integer

            For i = 1 To NUMCLASES

                If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
              
                    '//Si es un hechizo
                    If ObjData(ObjIndex).OBJType = eOBJType.otPergaminos Then
                        sMotivo = "Tu clase no tiene la habilidad de aprender este hechizo."
                        ClasePuedeUsarItem = False
                        Exit Function
                    Else
                        sMotivo = "Tu clase no puede usar este objeto."
                        ClasePuedeUsarItem = False
                        Exit Function
                    End If

                End If
                
            Next i

        End If

    End If
  
    ClasePuedeUsarItem = True

    Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")

End Function

Public Function ItemNoUsaConUser(ByVal UserIndex As Integer, _
                            ByVal ObjIndex As Integer) As Boolean
    '***************************************************
    'Autor: Lorwik
    'Fecha: 14/07/2020
    'Descripcion Devuelve true si el usuario no puede usar el item debido a su raza, sexo o clase
    '***************************************************

    If ObjIndex = 0 Then
        ItemNoUsaConUser = False
        Exit Function
    End If

    Select Case ObjData(ObjIndex).OBJType

        Case eOBJType.otWeapon, eOBJType.otAnillo, eOBJType.otFlechas, eOBJType.otEscudo
            If ClasePuedeUsarItem(UserIndex, ObjIndex) And FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                ItemNoUsaConUser = False
            Else
                ItemNoUsaConUser = True
            End If

        Case eOBJType.otArmadura
            If ClasePuedeUsarItem(UserIndex, ObjIndex) And SexoPuedeUsarItem(UserIndex, ObjIndex) And CheckRazaUsaRopa(UserIndex, ObjIndex) And FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                ItemNoUsaConUser = False
            Else
                ItemNoUsaConUser = True
            End If

        Case eOBJType.otCasco, eOBJType.otPergaminos
            If ClasePuedeUsarItem(UserIndex, ObjIndex) Then
                ItemNoUsaConUser = False
            Else
                ItemNoUsaConUser = True
            End If

        Case Else
            ItemNoUsaConUser = False

    End Select

End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim j As Integer

    With UserList(UserIndex)

        For j = 1 To UserList(UserIndex).CurrentInventorySlots

            If .Invent.Object(j).ObjIndex > 0 Then
             
                If ObjData(.Invent.Object(j).ObjIndex).Newbie = 1 Then Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                Call UpdateUserInv(False, UserIndex, j)
        
            End If

        Next j
    
        '[Barrin 17-12-03] Si el usuario dejo de ser Newbie, y estaba en el Newbie Dungeon
        'es transportado a su hogar de origen ;)
        If MapInfo(.Pos.Map).Restringir = eRestrict.restrict_newbie Then
        
            Dim DeDonde As WorldPos
        
            Select Case .Hogar

                Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                    DeDonde = Lindos

                Case eCiudad.cUllathorpe
                    DeDonde = Ullathorpe
                    
                Case eCiudad.cNix
                    DeDonde = Nix

                Case eCiudad.cBanderbill
                    DeDonde = Banderbill
                    
                Case eCiudad.cArghal
                    DeDonde = Arghal
                    
                Case eCiudad.cRinkel
                    DeDonde = Rinkel

                Case Else
                    DeDonde = Nix

            End Select
        
            Call WarpUserChar(UserIndex, DeDonde.Map, DeDonde.X, DeDonde.Y, True)
    
        End If

        '[/Barrin]
    End With

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim j As Integer

    With UserList(UserIndex)

        For j = 1 To .CurrentInventorySlots
            .Invent.Object(j).ObjIndex = 0
            .Invent.Object(j).Amount = 0
            .Invent.Object(j).Equipped = 0
        Next j
    
        .Invent.NroItems = 0
    
        .Invent.ArmourEqpObjIndex = 0
        .Invent.ArmourEqpSlot = 0
    
        .Invent.WeaponEqpObjIndex = 0
        .Invent.WeaponEqpSlot = 0
        
        .Invent.NudiEqpIndex = 0
        .Invent.NudiEqpSlot = 0
    
        .Invent.CascoEqpObjIndex = 0
        .Invent.CascoEqpSlot = 0
    
        .Invent.EscudoEqpObjIndex = 0
        .Invent.EscudoEqpSlot = 0
    
        .Invent.AnilloEqpObjIndex = 0
        .Invent.AnilloEqpSlot = 0
    
        .Invent.MunicionEqpObjIndex = 0
        .Invent.MunicionEqpSlot = 0
    
        .Invent.BarcoObjIndex = 0
        .Invent.BarcoSlot = 0
        
        .Invent.MonturaObjIndex = 0
        .Invent.MonturaEqpSlot = 0
    End With

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 23/01/2007
    '23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
    '***************************************************
    On Error GoTo errHandler

    'If Cantidad > 100000 Then Exit Sub

    With UserList(UserIndex)

        'SI EL Pjta TIENE ORO LO TIRAMOS
        If (Cantidad > 0) And (Cantidad <= .Stats.Gld) Then

            Dim MiObj As obj

            'info debug
            Dim loops As Integer
            
            'Seguridad Alkon (guardo el oro tirado si supera los 50k)
            If Cantidad > 50000 Then

                Dim j        As Integer

                Dim K        As Integer

                Dim M        As Integer

                Dim Cercanos As String

                M = .Pos.Map

                For j = .Pos.X - 10 To .Pos.X + 10
                    For K = .Pos.Y - 10 To .Pos.Y + 10

                        If InMapBounds(M, j, K) Then
                            If MapData(M, j, K).UserIndex > 0 Then
                                Cercanos = Cercanos & UserList(MapData(M, j, K).UserIndex).Name & ","

                            End If

                        End If

                    Next K
                Next j

                Call LogDesarrollo(.Name & " tira oro. Cercanos: " & Cercanos)

            End If

            '/Seguridad
            Dim Extra    As Long

            Dim TeniaOro As Long

            TeniaOro = .Stats.Gld

            If Cantidad > 500000 Then 'Para evitar explotar demasiado
                Extra = Cantidad - 500000
                Cantidad = 500000

            End If
            
            Do While (Cantidad > 0)
                
                If Cantidad > MAX_INVENTORY_OBJS And .Stats.Gld > MAX_INVENTORY_OBJS Then
                    MiObj.Amount = MAX_INVENTORY_OBJS
                    Cantidad = Cantidad - MiObj.Amount
                Else
                    MiObj.Amount = Cantidad
                    Cantidad = Cantidad - MiObj.Amount

                End If
    
                MiObj.ObjIndex = iORO
                
                If EsGm(UserIndex) Then Call LogGM(.Name, "Tiro cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)

                Dim AuxPos As WorldPos
                
                If .clase = eClass.Mercenario And .Invent.BarcoObjIndex = 476 Then
                    AuxPos = TirarItemAlPiso(.Pos, MiObj, False)

                    If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                        .Stats.Gld = .Stats.Gld - MiObj.Amount

                    End If

                Else
                    AuxPos = TirarItemAlPiso(.Pos, MiObj, True)

                    If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                        .Stats.Gld = .Stats.Gld - MiObj.Amount

                    End If

                End If
                
                'info debug
                loops = loops + 1

                If loops > 100 Then
                    LogError ("Error en tiraroro")
                    Exit Sub

                End If
                
            Loop

            If TeniaOro = .Stats.Gld Then Extra = 0
            If Extra > 0 Then
                .Stats.Gld = .Stats.Gld - Extra

            End If
        
        End If

    End With

    Exit Sub

errHandler:
    Call LogError("Error en TirarOro. Error " & Err.Number & " : " & Err.description)

End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, _
                      ByVal Slot As Byte, _
                      ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
    
    With UserList(UserIndex).Invent.Object(Slot)

        If .Amount <= Cantidad And .Equipped = 1 Then
            Call Desequipar(UserIndex, Slot)

        End If
        
        'Quita un objeto
        .Amount = .Amount - Cantidad

        'Quedan mas?
        If .Amount <= 0 Then
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
            .ObjIndex = 0
            .Amount = 0

        End If

    End With

    Exit Sub

errHandler:
    Call LogError("Error en QuitarUserInvItem. Error " & Err.Number & " : " & Err.description)
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, _
                  ByVal UserIndex As Integer, _
                  ByVal Slot As Byte)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Dim NullObj As UserObj

    Dim LoopC   As Long

    With UserList(UserIndex)

        'Actualiza un solo slot
        If Not UpdateAll Then
    
            'Actualiza el inventario
            If .Invent.Object(Slot).ObjIndex > 0 Then
                Call ChangeUserInv(UserIndex, Slot, .Invent.Object(Slot))
            Else
                Call ChangeUserInv(UserIndex, Slot, NullObj)

            End If
    
        Else
    
            'Actualiza todos los slots
            For LoopC = 1 To .CurrentInventorySlots

                'Actualiza el inventario
                If .Invent.Object(LoopC).ObjIndex > 0 Then
                    Call ChangeUserInv(UserIndex, LoopC, .Invent.Object(LoopC))
                Else
                    Call ChangeUserInv(UserIndex, LoopC, NullObj)

                End If

            Next LoopC

        End If
    
        Exit Sub

    End With

errHandler:
    Call LogError("Error en UpdateUserInv. Error " & Err.Number & " : " & Err.description)

End Sub

Sub DropObj(ByVal UserIndex As Integer, _
            ByRef DropObj As obj, _
            ByVal Slot As Byte, _
            ByVal Num As Integer, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer, _
            Optional ByVal Muere As Boolean = False)
    '***************************************************
    'Author: Unknown
    'Last Modification: 11/5/2010
    '11/5/2010 - ZaMa: Arreglo bug que permitia apilar mas de 10k de items.
    '***************************************************

    Dim MapObj      As obj
    
    Dim TieneCarro  As Boolean
    
    With UserList(UserIndex)

        If Num > 0 Then
            
            'Validacion para que no podamos tirar nuestra monturas mientras la usamos.
            If .flags.Equitando = 1 And .Invent.MonturaObjIndex = DropObj.ObjIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes tirar tu montura mientras la estas usando.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            If (ItemNewbie(DropObj.ObjIndex) And (.flags.Privilegios And PlayerType.User)) Then
                Call WriteConsoleMsg(UserIndex, "No puedes tirar objetos newbie.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
            
            If ObjData(DropObj.ObjIndex).OBJType = otRuna Then
                Call WriteConsoleMsg(UserIndex, "No puedes tirar la Runa.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If

            'Check objeto en el suelo
            MapObj.ObjIndex = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
            MapObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount

            If MapObj.ObjIndex = 0 Or MapObj.ObjIndex = DropObj.ObjIndex Then
        
                If MapObj.Amount = MAX_INVENTORY_OBJS Then
                    Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub

                End If
            
                If DropObj.Amount + MapObj.Amount > MAX_INVENTORY_OBJS Then
                    DropObj.Amount = MAX_INVENTORY_OBJS - MapObj.Amount

                End If
            
                Call MakeObj(DropObj, Map, X, Y)
                Call QuitarUserInvItem(UserIndex, Slot, DropObj.Amount)
                Call UpdateUserInv(False, UserIndex, Slot)
            
                If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Tiro cantidad:" & Num & " Objeto:" & ObjData(DropObj.ObjIndex).Name)
            
                'Log de Objetos que se tiran al piso. Pablo (ToxicWaste) 07/09/07
                'Es un Objeto que tenemos que loguear?
                If ObjData(DropObj.ObjIndex).Log = 1 Then
                    Call LogDesarrollo(.Name & " tiro al piso " & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)
                ElseIf DropObj.Amount > 5000 Then 'Es mucha cantidad? > Subi a 5000 el minimo porque si no se llenaba el log de cosas al pedo. (NicoNZ)

                    'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(DropObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(.Name & " tiro al piso " & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)

                    End If

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

End Sub

Sub EraseObj(ByVal Num As Integer, _
             ByVal Map As Integer, _
             ByVal X As Integer, _
             ByVal Y As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim TieneLuz As Boolean

    With MapData(Map, X, Y)
    
        '�Es un objeto que crea luz en el mapa?
        If ObjData(.ObjInfo.ObjIndex).CreaLuz.Rango > 0 Then TieneLuz = True
    
        .ObjInfo.Amount = .ObjInfo.Amount - Num
    
        If .ObjInfo.Amount <= 0 Then
            .ObjInfo.ObjIndex = 0
            .ObjInfo.Amount = 0

            Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectDelete(X, Y, TieneLuz))

        End If

    End With

End Sub

Sub MakeObj(ByRef obj As obj, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    If obj.ObjIndex > 0 And obj.ObjIndex <= UBound(ObjData) Then
    
        With MapData(Map, X, Y)

            If .ObjInfo.ObjIndex = obj.ObjIndex Then
                .ObjInfo.Amount = .ObjInfo.Amount + obj.Amount
            Else
                .ObjInfo = obj
                
                Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(obj.ObjIndex).GrhIndex, ObjData(obj.ObjIndex).ParticulaIndex, ObjData(obj.ObjIndex).CreaLuz.Rango, ObjData(obj.ObjIndex).CreaLuz.Color, X, Y))

            End If
            
            '//Agregamos las pos de los objetos
            If ObjData(obj.ObjIndex).OBJType <> otFogata And ItemNoEsDeMapa(ObjData(obj.ObjIndex).OBJType) Then
                Dim xPos As WorldPos

                xPos.Map = Map
                xPos.X = X
                xPos.Y = Y
                If (MapData(xPos.Map, xPos.X, xPos.Y).Trigger <> eTrigger.CASA Or MapData(xPos.Map, xPos.X, xPos.Y).Trigger <> eTrigger.BAJOTECHO) And MapData(xPos.Map, xPos.X, xPos.Y).Blocked <> 1 Then AgregarObjetoLimpieza xPos

            End If

        End With

    End If

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As obj) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 17/12/2019
    '***************************************************

    On Error GoTo errHandler

    Dim Slot As Byte

    With UserList(UserIndex)

        .CurrentInventorySlots = MAX_INVENTORY_SLOTS
        
        'el user ya tiene un objeto del mismo tipo?
        Slot = 1
        
        Do Until .Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
            Slot = Slot + 1

            If Slot > .CurrentInventorySlots Then
                Exit Do

            End If

        Loop
            
        'Sino busca un slot vacio
        If Slot > .CurrentInventorySlots Then
            Slot = 1

            Do Until .Invent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1

                If Slot > .CurrentInventorySlots Then
                    Call WriteConsoleMsg(UserIndex, "No puedes cargar mas objetos.", FontTypeNames.FONTTYPE_WARNING)
                    MeterItemEnInventario = False
                    Exit Function

                End If

            Loop
            .Invent.NroItems = .Invent.NroItems + 1

        End If

        'Mete el objeto
        If .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            .Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + MiObj.Amount
        Else
            .Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS

        End If

    End With
    
    MeterItemEnInventario = True
           
    Call UpdateUserInv(False, UserIndex, Slot)
    
    Exit Function
errHandler:
    Call LogError("Error en MeterItemEnInventario. Error " & Err.Number & " : " & Err.description)

End Function

Sub GetObj(ByVal UserIndex As Integer)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 18/12/2009
    '18/12/2009: ZaMa - Oro directo a la billetera.
    '***************************************************

    Dim obj    As ObjData

    Dim MiObj  As obj

    Dim ObjPos As String
    
    With UserList(UserIndex)

        'Hay algun obj?
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex > 0 Then

            'Esta permitido agarrar este obj?
            If ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then

                Dim X As Integer

                Dim Y As Integer
                
                X = .Pos.X
                Y = .Pos.Y
                
                obj = ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex)
                MiObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
                MiObj.ObjIndex = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
                ' Oro directo a la billetera!
                If obj.OBJType = otOro Then

                    'Calculamos la diferencia con el maximo de oro permitido el cual es el valor de LONG
                    Dim RemainingAmountToMaximumGold As Long
                    RemainingAmountToMaximumGold = 2147483647 - .Stats.Gld

                    If Not .Stats.Gld > 2147483647 And RemainingAmountToMaximumGold >= MiObj.Amount Then
                        .Stats.Gld = .Stats.Gld + MiObj.Amount
                        'Quitamos el objeto
                        Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)
                            
                        Call WriteUpdateGold(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes juntar este oro por que tendrias mas del maximo disponible (2147483647)", FontTypeNames.FONTTYPE_INFO)
                    End If

                Else

                    If MeterItemEnInventario(UserIndex, MiObj) Then
                    
                        'Quitamos el objeto
                        Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)

                        If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
        
                        'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
                        'Es un Objeto que tenemos que loguear?
                        If ObjData(MiObj.ObjIndex).Log = 1 Then
                            ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
                            Call LogDesarrollo(.Name & " junto del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)
                        
                        ElseIf MiObj.Amount > 5000 Then 'Es mucha cantidad?

                            'Si no es de los prohibidos de loguear, lo logueamos.
                            If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
                                ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
                                Call LogDesarrollo(.Name & " junto del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)

                            End If

                        End If

                    End If

                End If

            End If

        Else
            If Not .flags.UltimoMensaje = 99 Then
                .flags.UltimoMensaje = 99
                
                Call WriteConsoleMsg(UserIndex, "No hay nada aqui.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

End Sub

Public Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    'Desequipa el item slot del inventario
    Dim obj As ObjData
    
    With UserList(UserIndex)
        With .Invent

            If (Slot < LBound(.Object)) Or (Slot > UBound(.Object)) Then
                Exit Sub
            ElseIf .Object(Slot).ObjIndex = 0 Then
                Exit Sub

            End If
            
            obj = ObjData(.Object(Slot).ObjIndex)

        End With
        
        Select Case obj.OBJType

            Case eOBJType.otWeapon, eOBJType.otHerramientas

                With .Invent
                    .Object(Slot).Equipped = 0
                    .WeaponEqpObjIndex = 0
                    .WeaponEqpSlot = 0

                End With
                
                If Not .flags.Mimetizado = 1 Then

                    With .Char
                        .WeaponAnim = NingunArma
                        .AuraAnim = NingunAura
                        .AuraColor = NingunAura
                        Call ChangeUserChar(UserIndex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraAnim, .AuraColor)

                    End With

                End If
                
                '�Esta trabajando?
                'Las herramientas estan dateadas como armas, asi que estoy va aqui.
                If .flags.MacroTrabajo <> 0 Then
                    Call DejardeTrabajar(UserIndex)
                End If
                
            Case eOBJType.otNudillos
                .Invent.NudiEqpIndex = 0
                .Invent.NudiEqpSlot = 0
                .Invent.Object(Slot).Equipped = 0
                
                With .Char
                    .WeaponAnim = NingunArma
                    Call ChangeUserChar(UserIndex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraAnim, .AuraColor)
                End With

            Case eOBJType.otFlechas

                With .Invent
                    .Object(Slot).Equipped = 0
                    .MunicionEqpObjIndex = 0
                    .MunicionEqpSlot = 0

                End With
            
            Case eOBJType.otAnillo

                With .Invent
                    .Object(Slot).Equipped = 0
                    .AnilloEqpObjIndex = 0
                    .AnilloEqpSlot = 0
                    
                End With
                
                If obj.Efectomagico = eEfectos.ModificaAtributo Then
                    If obj.QueAtributo <> 0 Then _
                        .Stats.UserAtributos(obj.QueAtributo) = .Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
                        
                ElseIf obj.Efectomagico = eEfectos.ModificaSkill Then
                    If obj.QueSkill <> 0 Then _
                        .Stats.UserSkills(obj.QueSkill) = .Stats.UserSkills(obj.QueSkill) - obj.CuantoAumento

                End If
            
            Case eOBJType.otArmadura

                With .Invent
                    .Object(Slot).Equipped = 0
                    .ArmourEqpObjIndex = 0
                    .ArmourEqpSlot = 0

                End With
                
                If Not .flags.Mimetizado = 1 And Not .flags.Navegando = 1 Then
                    Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado = 1)
    
                    With .Char
                        Call ChangeUserChar(UserIndex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraAnim, .AuraColor)
    
                    End With
                End If
                
            Case eOBJType.otCasco

                With .Invent
                    .Object(Slot).Equipped = 0
                    .CascoEqpObjIndex = 0
                    .CascoEqpSlot = 0

                End With
                
                If Not .flags.Mimetizado = 1 Or Not .flags.Navegando = 1 Then

                    With .Char
                        .CascoAnim = NingunCasco
                        Call ChangeUserChar(UserIndex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraAnim, .AuraColor)

                    End With

                End If
            
            Case eOBJType.otEscudo

                With .Invent
                    .Object(Slot).Equipped = 0
                    .EscudoEqpObjIndex = 0
                    .EscudoEqpSlot = 0

                End With
                
                If Not .flags.Mimetizado = 1 Then

                    With .Char
                        .ShieldAnim = NingunEscudo
                        Call ChangeUserChar(UserIndex, .body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .AuraAnim, .AuraColor)

                    End With

                End If

        End Select

    End With
    
    Call WriteUpdateUserStats(UserIndex)
    Call UpdateUserInv(False, UserIndex, Slot)
    
    Exit Sub

errHandler:
    Call LogError("Error en Desquipar. Error " & Err.Number & " : " & Err.description)

End Sub

Function EsUsable(ByVal ObjIndex As Integer)
'*************************************************
'Author: Jopi
'Revisa si el objeto puede ser equipado/usado.
'Todas las opciones no usables estan en un solo case en ves de muchos diferentes (Recox)
'*************************************************
    Dim obj As ObjData
        obj = ObjData(ObjIndex)
        
    Select Case obj.OBJType
    
         Case eOBJType.otArboles, _
              eOBJType.otCarteles, _
              eOBJType.otForos, _
              eOBJType.otFragua, _
              eOBJType.otMuebles, _
              eOBJType.otPuertas, _
              eOBJType.otTeleport, _
              eOBJType.otYacimiento, _
              eOBJType.otYunque
         
            EsUsable = False
            
            Exit Function
        
    End Select
    
    EsUsable = True

End Function

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, _
                           ByVal ObjIndex As Integer, _
                           Optional ByRef sMotivo As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '***************************************************

    On Error GoTo errHandler
    
    If ObjData(ObjIndex).Mujer = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Hombre
    ElseIf ObjData(ObjIndex).Hombre = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Mujer
    Else
        SexoPuedeUsarItem = True

    End If
    
    If Not SexoPuedeUsarItem Then sMotivo = "Tu genero no puede usar este objeto."
    
    Exit Function
errHandler:
    Call LogError("SexoPuedeUsarItem")

End Function

Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, _
                              ByVal ObjIndex As Integer, _
                              Optional ByRef sMotivo As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '***************************************************

    If ObjData(ObjIndex).Real = 1 Then
        If Not criminal(UserIndex) Then
            FaccionPuedeUsarItem = esArmada(UserIndex)
        Else
            FaccionPuedeUsarItem = False

        End If

    ElseIf ObjData(ObjIndex).Caos = 1 Then

        If criminal(UserIndex) Then
            FaccionPuedeUsarItem = esCaos(UserIndex)
        Else
            FaccionPuedeUsarItem = False

        End If

    Else
        FaccionPuedeUsarItem = True

    End If
    
    If Not FaccionPuedeUsarItem Then sMotivo = "Tu alineacion no puede usar este objeto."

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
    '*************************************************
    'Author: Unknown
    'Last modified: 03/02/2020
    '01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin
    '14/01/2010: ZaMa - Agrego el motivo especifico por el que no puede equipar/usar el item.
    '03/02/2020: WyroX - Nivel minimo y skill minimo para poder equipar
    '*************************************************

    On Error GoTo errHandler

    'Equipa un item del inventario
    Dim obj      As ObjData

    Dim ObjIndex As Integer

    Dim sMotivo  As String
    
    With UserList(UserIndex)
    
        'Prevenimos posibles errores de otros codigos
        If Slot < 1 Then Exit Sub
    
        ObjIndex = .Invent.Object(Slot).ObjIndex
        obj = ObjData(ObjIndex)
        
        ' No se pueden usar muebles.
        If Not EsUsable(ObjIndex) Then Exit Sub

        If .flags.Equitando = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes equiparte o desequiparte mientras estas en tu montura!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        ' Nivel minimo
        If .Stats.ELV < obj.MinLevel Then
            Call WriteConsoleMsg(UserIndex, "Necesitas ser nivel " & obj.MinLevel & " para poder equipar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Skills minimos
        If obj.SkillRequerido Then
            If .Stats.UserSkills(obj.SkillRequerido) < obj.SkillCantidad Then
                Call WriteConsoleMsg(UserIndex, "Necesitas " & obj.SkillCantidad & " puntos en " & SkillsNames(obj.SkillRequerido) & " para poder equipar este objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
                
        Select Case obj.OBJType

            Case eOBJType.otWeapon, eOBJType.otHerramientas

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then

                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)

                        'Animacion por defecto
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.WeaponAnim = NingunArma
                        Else
                            .Char.WeaponAnim = NingunArma
                            .Char.AuraAnim = NingunAura
                            .Char.AuraColor = NingunAura
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)

                        End If

                        Exit Sub

                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.WeaponEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)

                    End If
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.WeaponEqpObjIndex = ObjIndex
                    .Invent.WeaponEqpSlot = Slot
                    
                    'El sonido solo se envia si no lo produce un admin invisible
                    If Not (.flags.AdminInvisible = 1) Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                    
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.WeaponAnim = GetWeaponAnim(UserIndex, ObjIndex)
                    Else
                        .Char.WeaponAnim = GetWeaponAnim(UserIndex, ObjIndex)
                        .Char.AuraAnim = obj.GrhAura
                        .Char.AuraColor = obj.AuraColor
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eOBJType.otNudillos
            
                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)
                        'Animacion por defecto
                        .Char.WeaponAnim = NingunArma
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
                        Exit Sub
                    End If
                    
                    'Quitamos el arma si tiene alguna equipada
                    If .Invent.WeaponEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                    End If
                    
                    If .Invent.NudiEqpIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.NudiEqpSlot)
                    End If
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.NudiEqpIndex = .Invent.Object(Slot).ObjIndex
                    .Invent.NudiEqpSlot = Slot
        
                    .Char.WeaponAnim = obj.WeaponAnim
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
                    
               Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                    
               End If
            
            Case eOBJType.otAnillo

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then

                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)
                        Exit Sub

                    End If
                        
                    'Quitamos el elemento anterior
                    If .Invent.AnilloEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)

                    End If
                
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.AnilloEqpObjIndex = ObjIndex
                    .Invent.AnilloEqpSlot = Slot
                    
                    If obj.Efectomagico = eEfectos.ModificaAtributo Then
                        If obj.QueAtributo <> 0 Then
                            .Stats.UserAtributos(obj.QueAtributo) = .Stats.UserAtributos(obj.QueAtributo) + obj.CuantoAumento
                        End If
                    ElseIf obj.Efectomagico = eEfectos.ModificaSkill Then
                        If obj.QueSkill <> 0 Then
                            .Stats.UserSkills(obj.QueSkill) = .Stats.UserSkills(obj.QueSkill) + obj.CuantoAumento
                        End If
                    End If
                        
                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eOBJType.otFlechas

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                        
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)
                        Exit Sub

                    End If
                        
                    'Quitamos el elemento anterior
                    If .Invent.MunicionEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)

                    End If
                
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.MunicionEqpObjIndex = ObjIndex
                    .Invent.MunicionEqpSlot = Slot
                        
                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eOBJType.otArmadura

                'Parchesin para que no se saquen una armadura mientras estan en montura y dsp les queda el cuerpo de la armadura y velocidad de montura (Recox)
                If .flags.Equitando = 1 Then
                    Call WriteConsoleMsg(UserIndex, "No podes equiparte o desequiparte vestimentas o armaduras mientras estas en tu montura.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                'Nos aseguramos que puede usarla
                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And SexoPuedeUsarItem(UserIndex, ObjIndex, sMotivo) And CheckRazaUsaRopa(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then

                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)

                        If Not .flags.Mimetizado = 1 And .flags.Navegando = 0 Then
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)

                        End If

                        Exit Sub

                    End If
            
                    'Quita el anterior
                    If .Invent.ArmourEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)

                    End If
            
                    'Lo equipa
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.ArmourEqpObjIndex = ObjIndex
                    .Invent.ArmourEqpSlot = Slot
                        
                    If .flags.Mimetizado = 1 Or .flags.Navegando = 1 Then
                        .CharMimetizado.body = obj.Ropaje
                    Else
                        .Char.body = obj.Ropaje
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)

                    End If

                    .flags.Desnudo = 0
                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eOBJType.otCasco

                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then

                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)

                        If .flags.Mimetizado = 1 Or .flags.Navegando = 1 Then
                            .CharMimetizado.CascoAnim = NingunCasco
                        Else
                            .Char.CascoAnim = NingunCasco
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)

                        End If

                        Exit Sub

                    End If
            
                    'Quita el anterior
                    If .Invent.CascoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.CascoEqpSlot)

                    End If
            
                    'Lo equipa
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.CascoEqpObjIndex = ObjIndex
                    .Invent.CascoEqpSlot = Slot

                    If .flags.Mimetizado = 1 Or .flags.Navegando = 1 Then
                        .CharMimetizado.CascoAnim = obj.CascoAnim
                    Else
                        .Char.CascoAnim = obj.CascoAnim
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eOBJType.otEscudo


                If ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then

                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)

                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.ShieldAnim = NingunEscudo
                        Else
                            .Char.ShieldAnim = NingunEscudo
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)

                        End If

                        Exit Sub

                    End If
             
                    'Quita el anterior
                    If .Invent.EscudoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)

                    End If
             
                    'Lo equipa
                     
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.EscudoEqpObjIndex = ObjIndex
                    .Invent.EscudoEqpSlot = Slot
                     
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.ShieldAnim = obj.ShieldAnim
                    Else
                        .Char.ShieldAnim = obj.ShieldAnim
                         
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)

                End If

        End Select

    End With
    
    'Actualiza
    Call UpdateUserInv(False, UserIndex, Slot)
    
    Exit Sub
    
errHandler:
    Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.description)

End Sub

Private Function CheckRazaUsaRopa(ByVal UserIndex As Integer, _
                                  ItemIndex As Integer, _
                                  Optional ByRef sMotivo As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 14/01/2010 (ZaMa)
    '14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
    '***************************************************

    On Error GoTo errHandler

    With UserList(UserIndex)

        'Verifica si la raza puede usar la ropa
        If .Raza = eRaza.Humano Or .Raza = eRaza.Elfo Or .Raza = eRaza.Drow Or .Raza = eRaza.Orco Then
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
        Else
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)

        End If
        
        'Solo se habilita la ropa exclusiva para Drows por ahora. Pablo (ToxicWaste)
        If (.Raza <> eRaza.Drow) And ObjData(ItemIndex).RazaDrow Then
            CheckRazaUsaRopa = False

        End If

    End With
    
    If Not CheckRazaUsaRopa Then sMotivo = "Tu raza no puede usar este objeto."
    
    Exit Function
    
errHandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
    '*************************************************
    'Author: Unknown
    'Last modified: 03/02/2020
    'Handels the usage of items from inventory box.
    '24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legion.
    '24/01/2007 Pablo (ToxicWaste) - Utilizacion nueva de Barco en lvl 20 por clase Pirata y Pescador.
    '01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin, except to its own client
    '17/11/2009: ZaMa - Ahora se envia una orientacion de la posicion hacia donde esta el que uso el cuerno.
    '27/11/2009: Budi - Se envia indivualmente cuando se modifica a la Agilidad o la Fuerza del personaje.
    '08/12/2009: ZaMa - Agrego el uso de hacha de madera elfica.
    '10/12/2009: ZaMa - Arreglos y validaciones en todos las herramientas de trabajo.
    '03/02/2020: WyroX - Nivel minimo para poder usar y clase prohibida a pergaminos y barcas
    '*************************************************

    Dim obj      As ObjData

    Dim ObjIndex As Integer

    Dim TargObj  As ObjData

    Dim MiObj    As obj

    Dim sMotivo As String

    With UserList(UserIndex)
    
        'Prevenimos posibles errores de otros codigos
        If Slot < 1 Then Exit Sub
    
        If .Invent.Object(Slot).Amount = 0 Then Exit Sub
        
        obj = ObjData(.Invent.Object(Slot).ObjIndex)

        If obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        '�Est� trabajando?
        If .flags.MacroTrabajo <> 0 Then
            Call WriteConsoleMsg(UserIndex, "�No puedes usar objetos mientras estas trabajando!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If obj.OBJType = eOBJType.otWeapon Then
            If obj.proyectil = 1 Then
                
                'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
                If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
            Else

                'dagas
                If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub

            End If

        Else

            If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub

        End If
        
        ObjIndex = .Invent.Object(Slot).ObjIndex
        .flags.TargetObjInvIndex = ObjIndex
        .flags.TargetObjInvSlot = Slot
        
        ' No se pueden usar muebles.
        If Not EsUsable(ObjIndex) Then Exit Sub

        ' Nivel minimo
        If .Stats.ELV < obj.MinLevel Then
            Call WriteConsoleMsg(UserIndex, "Necesitas ser nivel " & obj.MinLevel & " para poder usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Select Case obj.OBJType

            Case eOBJType.otUseOnce

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                    Exit Sub

                End If
        
                'Usa el item
                .Stats.MinHam = .Stats.MinHam + obj.MinHam

                If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
                .flags.Hambre = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
                'Sonido
                
                If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
                    Call ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MORFAR_MANZANA)
                Else
                    Call ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_COMIDA)

                End If
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                Call UpdateUserInv(False, UserIndex, Slot)
        
            Case eOBJType.otOro

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                    Exit Sub

                End If
                
                .Stats.Gld = .Stats.Gld + .Invent.Object(Slot).Amount
                .Invent.Object(Slot).Amount = 0
                .Invent.Object(Slot).ObjIndex = 0
                .Invent.NroItems = .Invent.NroItems - 1
                
                Call UpdateUserInv(False, UserIndex, Slot)
                Call WriteUpdateGold(UserIndex)
                
            Case eOBJType.otWeapon

                If .flags.Equitando = 1 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes usar una herramienta mientras estas en tu montura!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                    Exit Sub

                End If
                
                If Not .Stats.MinSta > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Estas muy cansad" & IIf(.Genero = eGenero.Hombre, "o", "a") & ".", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
                If ObjData(ObjIndex).proyectil = 1 Then
                    If .Invent.Object(Slot).Equipped = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Proyectiles)  'Call WriteWorkRequestTarget(UserIndex, Proyectiles)
                ElseIf .flags.TargetObj = Lena Then

                    If .Invent.Object(Slot).ObjIndex = DAGA Then
                        If .Invent.Object(Slot).Equipped = 0 Then
                            Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If
                            
                        Call TratarDeHacerFogata(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY, UserIndex)

                    End If
                End If

                Case eOBJType.otHerramientas
                    
                    Select Case ObjIndex
                    
                        Case CANA_PESCA, RED_PESCA
                            
                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Pesca)  'Call WriteWorkRequestTarget(UserIndex, eSkill.Pesca)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                        Case HACHA_LENADOR
                            
                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.talar)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                        Case PIQUETE_MINERO
                        
                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Mineria)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                        Case TIJERAS_BOTANICA
                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Botanica)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                        Case MARTILLO_HERRERO

                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Herreria)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                        Case SERRUCHO_CARPINTERO
                            
                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call EnviarObjConstruibles(UserIndex)
                                Call WriteShowTrabajoForm(UserIndex, eSkill.Carpinteria)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                        Case KIT_DE_COSTURA

                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call EnviarRopasConstruibles(UserIndex)
                                Call WriteShowTrabajoForm(UserIndex, eSkill.Sastreria)
                                
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                        Case OLLA_ALQUIMISTA

                            ' Lo tiene equipado?
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call EnviarPocionesConstruibles(UserIndex)
                                Call WriteShowTrabajoForm(UserIndex, eSkill.Alquimia)
                                
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                        Case Else ' Las herramientas no se pueden fundir

                            If ObjData(ObjIndex).SkHerreria > 0 Then
                                ' Solo objetos que pueda hacer el herrero
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, FundirMetal) 'Call WriteWorkRequestTarget(UserIndex, FundirMetal)
                                Exit Sub
                            End If

                    End Select

            
            Case eOBJType.otPociones

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                    Exit Sub

                End If
                
                If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
                    Call WriteConsoleMsg(UserIndex, "Debes esperar unos momentos para tomar otra pocion!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
                .flags.TomoPocion = True
                .flags.TipoPocion = obj.TipoPocion
                        
                Select Case .flags.TipoPocion
                
                    Case 1 'Modif la agilidad
                        .flags.DuracionEfecto = obj.DuracionEfecto
                
                        'Usa el item
                        .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador)

                        If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                        
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))

                        End If

                        Call WriteUpdateDexterity(UserIndex)
                        
                    Case 2 'Modif la fuerza
                        .flags.DuracionEfecto = obj.DuracionEfecto
                
                        'Usa el item
                        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador)

                        If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                        
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))

                        End If

                        Call WriteUpdateStrenght(UserIndex)
                        
                    Case 3 'Pocion roja, restaura HP
                        'Usa el item
                        .Stats.MinHp = .Stats.MinHp + RandomNumber(obj.MinModificador, obj.MaxModificador)

                        If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
                        
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))

                        End If
                    
                    Case 4 'Pocion azul, restaura MANA
                        'Usa el item
                        'nuevo calculo para recargar mana
                        .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, 4) + .Stats.ELV \ 2 + 40 / .Stats.ELV

                        If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
                        
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))

                        End If
                        
                    Case 5 ' Pocion violeta

                        If .flags.Envenenado = 1 Then
                            .flags.Envenenado = 0
                            Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)

                        End If

                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))

                        End If
                        
                    Case 6  ' Pocion Negra
                        If .flags.SlotReto > 0 Then Exit Sub
                        
                        If .flags.Privilegios And PlayerType.User Then
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Call UserDie(UserIndex)
                            Call WriteConsoleMsg(UserIndex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)

                        End If
                        
                    Case 7 ' Pocion Apagar quemaduras

                        If .flags.Incinerado = 1 Then
                            .flags.Incinerado = 0
                            Call WriteConsoleMsg(UserIndex, "Tus llamas se han apagado.", FontTypeNames.FONTTYPE_INFO)

                        End If

                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))

                        End If

                End Select

                Call WriteUpdateUserStats(UserIndex)
                Call UpdateUserInv(False, UserIndex, Slot)
        
            Case eOBJType.otBebidas

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                    Exit Sub

                End If

                .Stats.MinAGU = .Stats.MinAGU + obj.MinSed

                If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
                .flags.Sed = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))

                End If
                
                Call UpdateUserInv(False, UserIndex, Slot)
            
            Case eOBJType.otLlaves

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                    Exit Sub

                End If
                
                If .flags.TargetObj = 0 Then Exit Sub
                TargObj = ObjData(.flags.TargetObj)

                'El objeto clickeado es una puerta?
                If TargObj.OBJType = eOBJType.otPuertas Then

                    'Esta cerrada?
                    If TargObj.Cerrada = 1 Then

                        'Cerrada con llave?
                        If TargObj.Llave > 0 Then
                            If TargObj.Clave = obj.Clave Then
                 
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                                Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            Else
                                Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If

                        Else

                            If TargObj.Clave = obj.Clave Then
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                                Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                                Exit Sub
                            Else
                                Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                End If
            
            Case eOBJType.otBotellaVacia

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                    Exit Sub

                End If

                If Not HayAgua(.Pos.Map, .flags.TargetX, .flags.TargetY) Then
                    Call WriteConsoleMsg(UserIndex, "No hay agua alli.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                MiObj.Amount = 1
                MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexAbierta
                Call QuitarUserInvItem(UserIndex, Slot, 1)

                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)

                End If
                
                Call UpdateUserInv(False, UserIndex, Slot)
            
            Case eOBJType.otBotellaLlena

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                    Exit Sub

                End If

                .Stats.MinAGU = .Stats.MinAGU + obj.MinSed

                If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
                .flags.Sed = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
                MiObj.Amount = 1
                MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexCerrada
                Call QuitarUserInvItem(UserIndex, Slot, 1)

                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)

                End If
                
                Call UpdateUserInv(False, UserIndex, Slot)
            
            Case eOBJType.otPergaminos

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                    Exit Sub

                End If
                
                If .Stats.MaxMAN > 0 Then
                    If .flags.Hambre = 0 And .flags.Sed = 0 Then

                        If Not ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                            Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If

                        Call AgregarHechizo(UserIndex, Slot)
                        Call UpdateUserInv(False, UserIndex, Slot)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)

                End If

            Case eOBJType.otMinerales

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                    Exit Sub

                End If

                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, FundirMetal) 'Call WriteWorkRequestTarget(UserIndex, FundirMetal)
               
            Case eOBJType.otInstrumentos

                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                    Exit Sub

                End If
                
                Call doInstrumentos(UserIndex, ObjIndex)
               
            Case eOBJType.otBarcos

                If Not ClasePuedeUsarItem(UserIndex, ObjIndex, sMotivo) Or Not FaccionPuedeUsarItem(UserIndex, ObjIndex, sMotivo) Then
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                'Verifica si esta aproximado al agua antes de permitirle navegar
                If .Stats.ELV < 25 Then

                    ' Solo pirata puede navegar antes
                    If .clase <> eClass.Mercenario Then
                        Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    Else

                        ' Pero a partir de 20
                        If .Stats.ELV < 20 Then
                            
                            If .Stats.UserSkills(eSkill.Pesca) <> 100 Then
                                Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 y ademas tu skill en pesca debe ser 100.", FontTypeNames.FONTTYPE_INFO)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)

                            End If
                            
                            Exit Sub

                        End If

                    End If

                End If
                
                If (LegalPos(.Pos.Map, .Pos.X - 1, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y - 1, True, False) Or LegalPos(.Pos.Map, .Pos.X + 1, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y + 1, True, False)) And .flags.Navegando = 0 Then
                    Call DoNavega(UserIndex, obj, Slot)
                
                ElseIf (LegalPos(.Pos.Map, .Pos.X - 1, .Pos.Y, False, True) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y - 1, False, True) Or LegalPos(.Pos.Map, .Pos.X + 1, .Pos.Y, False, True) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y + 1, False, True)) And .flags.Navegando = 1 Then
                    Call DoNavega(UserIndex, obj, Slot)
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "Debes aproximarte al agua para navegar y a la tierra para bajar!", FontTypeNames.FONTTYPE_INFO)

                End If

            '<-------------> MONTURAS <----------->
            Case eOBJType.otMonturas
                If ClasePuedeUsarItem(UserIndex, ObjIndex) Then

                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "Estas muerto, no puedes montarte ni desmontarte en este estado!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    If .flags.Navegando = 1 Then
                        Call WriteConsoleMsg(UserIndex, "Estas navegando, no puedes montarte ni desmontarte en este estado!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    Call DoEquita(UserIndex, obj, Slot)
                Else
                    Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                End If
                    
            Case eOBJType.otPasajes
            
                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                    Exit Sub

                End If

                If .flags.TargetNpcTipo <> Marinero Then
                    Call WriteConsoleMsg(UserIndex, "Primero debes hacer click sobre el marinero.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                    
                If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                    Call WriteConsoleMsg(UserIndex, "�Estas demasiado lejos!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                    
                If .Pos.Map <> obj.DesdeMap Then
                    Call WriteConsoleMsg(UserIndex, "El pasaje no lo compraste aqu�! Largate!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                    
                If Not MapaValido(obj.HastaMap) Then
                    Call WriteConsoleMsg(UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                    
                If obj.NecesitaNave = 1 And .Stats.UserSkills(eSkill.Navegacion) < 40 Then
                    Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje no puedo llevarte. Necesitas 40 skills para utilizar este pasaje. Consulta el manual del juego en http://winterao.com.ar/wiki/ para saber c�mo conseguirlos.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                    
                If .Stats.ELV < 10 Then
                    Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, necesitas ser nivel 10 como minimo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                    
                Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
                Call WriteConsoleMsg(UserIndex, "Has viajado por varios d�as, te sientes exhausto!", FontTypeNames.FONTTYPE_CENTINELA)
                
                'Penalizador:
                .Stats.MinAGU = 0
                .Stats.MinHam = 0
                .flags.Sed = 1
                .flags.Hambre = 1
                
                Call WriteUpdateHungerAndThirst(UserIndex)
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call UpdateUserInv(False, UserIndex, Slot)
                
            Case eOBJType.otMapas
                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                    Exit Sub

                End If
                
                Call WriteAbrirMapa(UserIndex)
                
            Case eOBJType.otBolsasOro
            
                If .flags.Muerto = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                    Exit Sub

                End If
                
                Call usarBolsadeOro(UserIndex, Slot)
                
            Case eOBJType.otEsposas
                    Call Encarcelar(UserIndex, 10)
                    Call LogGM(.Name, " fue encarcelado por las esposas.")
                
            Case eOBJType.otRuna
                If .flags.Muerto = 1 Then
    
                    'Si es un mapa comun y no esta en cana
                    If (MapInfo(.Pos.Map).Restringir = eRestrict.restrict_no) And (.Counters.Pena = 0) Then
                        If Ciudades(.Hogar).Map <> .Pos.Map Then
                            Call MandaraCasa(UserIndex)
                        Else
                            Call WriteConsoleMsg(UserIndex, "Ya te encuentras en tu hogar.", FontTypeNames.FONTTYPE_INFO)
    
                        End If
    
                    Else
                        Call WriteConsoleMsg(UserIndex, "Una fuerza misteriosa interfiere con la runa, no puedes utilizarla aqu�.", FontTypeNames.FONTTYPE_FIGHT)
    
                    End If
    
                Else
                    Call WriteConsoleMsg(UserIndex, "La runa no funciona si estas vivo.", FontTypeNames.FONTTYPE_INFO)
    
                End If
                  
            End Select
    
    End With

End Sub

Private Sub usarBolsadeOro(ByVal UserIndex As Integer, ByVal Slot As Byte)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 06/03/2021
    'Descripci�n: A�ade la cantidad de oro de la bolsa a la billetera
    '***************************************************
    
    Dim obj As ObjData
    
    With UserList(UserIndex)
    
        obj = ObjData(.Invent.Object(Slot).ObjIndex)

        'Actualizamos su billetera
        .Stats.Gld = .Stats.Gld + obj.CuantoAgrega
        Call WriteUpdateGold(UserIndex)

        'Eliminamos el objeto
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        Call UpdateUserInv(False, UserIndex, Slot)
        
        Call LogDesarrollo(.Name & " ha obtenido " & obj.CuantoAgrega & " monedas de oro de " & obj.Name & ". Tenia " & .Invent.Object(Slot).Amount + 1 & " bolsas.")
        
        Call WriteConsoleMsg(UserIndex, "�Has obtenido " & obj.CuantoAgrega & " monedas de oro de " & obj.Name, FontTypeNames.FONTTYPE_INFO)

    End With
    
End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    With UserList(UserIndex)

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.ZONAPELEA Then Exit Sub

        If TieneSacri(UserIndex) Then Exit Sub
        
        Call TirarTodosLosItems(UserIndex)
           
        ' Si estas en zona segura tampoco se tira el oro.
        If MapInfo(.Pos.Map).Pk Then
            
            'Si supera los 100k no se cae
            If .Stats.Gld < 100000 Then
                Call TirarOro(.Stats.Gld, UserIndex)
            End If
            
        End If
        
    End With

    Exit Sub

errHandler:
    Call LogError("Error en TirarTodo. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Function TieneSacri(ByVal UserIndex As Integer) As Boolean
'****************************************
'Autor: Lorwik
'Fecha: 09/03/2021
'Descripcion: Comprueba si tiene el pendiente del sacrificio
'****************************************

    Dim PendienteSacrificio As Boolean
    Dim Slot                As Byte
    Dim MiObj               As obj
    
    With UserList(UserIndex)
    
        '�Tiene pendiente del sacrificio?
        'NOTA: Quizas para este nuevo modelo de pendiente esto podria ser innecesario.
        If .Invent.AnilloEqpObjIndex > 0 Then
            If ObjData(.Invent.AnilloEqpObjIndex).Efectomagico = eEfectos.Sacrificio Then
                Slot = .Invent.AnilloEqpSlot
                PendienteSacrificio = True
            End If
        End If
    
        If PendienteSacrificio Then
            '�Es el pendiente intacto?
            If .Invent.AnilloEqpObjIndex = PendienteIntacto Then
                'Se lo quitamos
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call UpdateUserInv(False, UserIndex, Slot)
                                
                'Le quitamos un uso y tiramos el nuevo objeto
                MiObj.ObjIndex = PendienteMedio
                MiObj.Amount = 1
                Call TirarItemAlPiso(.Pos, MiObj)
                
                TieneSacri = True
                Exit Function
                                
            ElseIf .Invent.AnilloEqpObjIndex = PendienteMedio Then
                'Se lo quitamos
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call UpdateUserInv(False, UserIndex, Slot)
                                
                'Le quitamos un uso y tiramos el nuevo objeto
                MiObj.ObjIndex = PendienteRoto
                MiObj.Amount = 1
                Call TirarItemAlPiso(.Pos, MiObj)
        
                TieneSacri = True
                Exit Function
                            
            ElseIf .Invent.AnilloEqpObjIndex = PendienteRoto Then
                'Se lo quitamos y se destruye
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call UpdateUserInv(False, UserIndex, Slot)
                
                TieneSacri = True
                Exit Function
                
            Else '�Es un pendiente del sacrificio infinito?
                MiObj.ObjIndex = Slot
                MiObj.Amount = 1
                Call TirarItemAlPiso(.Pos, MiObj)
                
                TieneSacri = True
                Exit Function
                
            End If
        End If
    End With
    
    TieneSacri = False
End Function

Public Function ItemSeCae(ByVal index As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With ObjData(index)
        ItemSeCae = (.Real <> 1 Or .NoSeCae = 0) And (.Caos <> 1 Or .NoSeCae = 0) And .OBJType <> eOBJType.otLlaves And .OBJType <> eOBJType.otBarcos And .NoSeCae = 0

    End With

End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010 (ZaMa)
    '12/01/2010: ZaMa - Ahora los piratas no explotan items solo si estan entre 20 y 25
    '***************************************************
    On Error GoTo errHandler

    Dim i           As Byte

    Dim NuevaPos    As WorldPos

    Dim MiObj       As obj

    Dim ItemIndex   As Integer

    Dim DropAgua    As Boolean
    
    Dim TieneCarro  As Boolean
    
    With UserList(UserIndex)

        If .Invent.AnilloEqpObjIndex > 0 Then
            If ObjData(.Invent.AnilloEqpObjIndex).Efectomagico = eEfectos.CarroMinerales Then TieneCarro = True
        End If

        For i = 1 To .CurrentInventorySlots
            ItemIndex = .Invent.Object(i).ObjIndex

            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
        
                    'Creo el Obj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.ObjIndex = ItemIndex
                    
                    If TieneCarro Then
                        If ObjData(MiObj.ObjIndex).OBJType = eOBJType.otMinerales Then _
                            MiObj.Amount = Porcentaje(MiObj.Amount, 30) 'Salvamos el 70% de los minerales
                    End If
                    
                    DropAgua = True

                    ' Es pirata?
                    If .clase = eClass.Mercenario Then

                        ' Si tiene galeon equipado
                        If .Invent.BarcoObjIndex = 476 Then

                            ' Limitacion por nivel, despues dropea normalmente
                            If .Stats.ELV = 20 Then
                                ' No dropea en agua
                                DropAgua = False

                            End If

                        End If

                    End If
                    
                    Call Tilelibre(.Pos, NuevaPos, MiObj, DropAgua, True)
                    
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, MiObj, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)

                    End If

                End If

            End If

        Next i

    End With
    
    Exit Sub
    
errHandler:
    Call LogError("Error en TirarTodosLosItems. Error: " & Err.Number & " - " & Err.description)

End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
    ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 23/11/2009
    '07/11/09: Pato - Fix bug #2819911
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    Dim i           As Byte

    Dim NuevaPos    As WorldPos

    Dim MiObj       As obj

    Dim ItemIndex   As Integer
    
    Dim TieneCarro  As Boolean
    
    With UserList(UserIndex)
    
        If .Invent.AnilloEqpObjIndex > 0 Then
            If ObjData(.Invent.AnilloEqpObjIndex).Efectomagico = eEfectos.CarroMinerales Then TieneCarro = True
        End If
    
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.ZONAPELEA Then Exit Sub
        
        If TieneSacri(UserIndex) Then Exit Sub
        
        For i = 1 To UserList(UserIndex).CurrentInventorySlots
            ItemIndex = .Invent.Object(i).ObjIndex

            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo MiObj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.ObjIndex = ItemIndex
                    
                    If TieneCarro Then
                        If ObjData(MiObj.ObjIndex).OBJType = eOBJType.otMinerales Then _
                            MiObj.Amount = Porcentaje(MiObj.Amount, 30) 'Salvamos el 70% de los minerales
                    End If
                    
                    'Pablo (ToxicWaste) 24/01/2007
                    'Tira los Items no newbies en todos lados.
                    Tilelibre .Pos, NuevaPos, MiObj, True, True

                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, MiObj, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)

                    End If

                End If

            End If

        Next i

    End With

End Sub

Public Function getObjType(ByVal ObjIndex As Integer) As eOBJType
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If ObjIndex > 0 Then
        getObjType = ObjData(ObjIndex).OBJType

    End If
    
End Function

Public Sub moveItem(ByVal UserIndex As Integer, _
                    ByVal originalSlot As Integer, _
                    ByVal newSlot As Integer)

    Dim tmpObj      As UserObj

    Dim newObjIndex As Integer, originalObjIndex As Integer

    If (originalSlot <= 0) Or (newSlot <= 0) Then Exit Sub

    With UserList(UserIndex)

        If (originalSlot > .CurrentInventorySlots) Or (newSlot > .CurrentInventorySlots) Then Exit Sub
    
        tmpObj = .Invent.Object(originalSlot)
        .Invent.Object(originalSlot) = .Invent.Object(newSlot)
        .Invent.Object(newSlot) = tmpObj
    
        'Viva VB6 y sus putas deficiencias.
        If .Invent.AnilloEqpSlot = originalSlot Then
            .Invent.AnilloEqpSlot = newSlot
        ElseIf .Invent.AnilloEqpSlot = newSlot Then
            .Invent.AnilloEqpSlot = originalSlot

        End If
    
        If .Invent.ArmourEqpSlot = originalSlot Then
            .Invent.ArmourEqpSlot = newSlot
        ElseIf .Invent.ArmourEqpSlot = newSlot Then
            .Invent.ArmourEqpSlot = originalSlot

        End If
    
        If .Invent.BarcoSlot = originalSlot Then
            .Invent.BarcoSlot = newSlot
        ElseIf .Invent.BarcoSlot = newSlot Then
            .Invent.BarcoSlot = originalSlot

        End If
    
        If .Invent.CascoEqpSlot = originalSlot Then
            .Invent.CascoEqpSlot = newSlot
        ElseIf .Invent.CascoEqpSlot = newSlot Then
            .Invent.CascoEqpSlot = originalSlot

        End If
    
        If .Invent.EscudoEqpSlot = originalSlot Then
            .Invent.EscudoEqpSlot = newSlot
        ElseIf .Invent.EscudoEqpSlot = newSlot Then
            .Invent.EscudoEqpSlot = originalSlot

        End If
    
        If .Invent.MunicionEqpSlot = originalSlot Then
            .Invent.MunicionEqpSlot = newSlot
        ElseIf .Invent.MunicionEqpSlot = newSlot Then
            .Invent.MunicionEqpSlot = originalSlot

        End If
    
        If .Invent.WeaponEqpSlot = originalSlot Then
            .Invent.WeaponEqpSlot = newSlot
        ElseIf .Invent.WeaponEqpSlot = newSlot Then
            .Invent.WeaponEqpSlot = originalSlot

        End If

        Call UpdateUserInv(False, UserIndex, originalSlot)
        Call UpdateUserInv(False, UserIndex, newSlot)

    End With

End Sub

Public Function ObtenerSlotObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Integer
'****************************************
'Autor: Lorwik
'Fecha: 15/03/2021
'Descripci�n: Devuelve el slot en el inventario donde esta el objeto.
'****************************************

    Dim i As Integer
    
    With UserList(UserIndex)

        '�Tiene el objeto en el inventario? Si es asi obtenemos el slot en el que esta
        For i = 1 To .CurrentInventorySlots
            If .Invent.Object(i).ObjIndex = ObjIndex Then
                ObtenerSlotObj = i
                Exit Function
            End If
        Next i

        ObtenerSlotObj = 0
    
    End With
End Function
