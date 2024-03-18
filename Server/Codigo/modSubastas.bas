Attribute VB_Name = "modSubastas"
'***************************************************
'Author: Ezequiel Ju�rez (Standelf)
'Last Modification: 25/05/10
'Blisse-AO | Black And White AO | Sistema de Subastas 0.13.x
'***************************************************
 
Public Type c_subasta
    Actual As Boolean ' Sabemos si hay una subasta actualmente
    
    UserIndex As Integer ' UserIndex del Usuario Subastando
    OfertaIndex As Integer ' UserIndex del usuario con mayor oferta
    
    OfertaMayor As Long ' Oferta que vale la pena
    ValorBase As Long ' Valor base del item
    
    Objeto As obj ' Objeto
    
    Tiempo As Byte ' Tiempo de Subasta
End Type
    
Public Subasta As c_subasta
 
Public Sub Init_Subastas()
    With Subasta
        .Actual = False
        .UserIndex = 0
        .OfertaIndex = 0
        .ValorBase = 0
        .OfertaMayor = 0
        .Tiempo = 0
    End With
End Sub
 
Public Sub Consultar_Subasta(ByVal UserIndex As Integer)
    With Subasta
        If .Actual = True Then
                If .UserIndex <> -1 Then
                    Call WriteConsoleMsg(UserIndex, "[Subasta] " & UserList(.UserIndex).Name & " est� subastando " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).Name & ". La oferta actual es de " & .OfertaMayor & ". Esta subasta seguir� por " & .Tiempo & " minutos.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "[Subasta] Se est� subastando " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).Name & ". La oferta actual es de " & .OfertaMayor & ". Esta subasta seguir� por " & .Tiempo & " minutos.", FontTypeNames.FONTTYPE_INFO)
                End If
            Exit Sub
        Else
            Call WriteConsoleMsg(UserIndex, "Actualmente no hay ninguna subasta activa.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub
 
Public Sub Iniciar_Subasta(ByVal UserIndex As Integer, Slot As Integer, Amount As Integer, ValorBase As Long)

    With Subasta
        ' Si ya hay una subasta le informamos que debe esperar
        If .Actual = True Then
            Call WriteConsoleMsg(UserIndex, "Ya hay una subasta actualmente, deber�s esperar " & .Tiempo & " minutos para inciar una nueva subasta.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            
            'Hizo aparecer el form con hacks?
            If UserList(UserIndex).flags.Subastando = False Then Exit Sub
            
            'reseteamos el flag
            UserList(UserIndex).flags.Subastando = False
            
            'AntiHack (Ya se comprobo en el cliente, se vuelve a comprobar aqui)
            If Amount <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Debes subastar una cantidad mayor a 0 de �tems.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If ValorBase <= 100 Then
                Call WriteConsoleMsg(UserIndex, "Debes poner un valor mayor a 100 para poder subastar.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).Invent.Object(Slot).ObjIndex <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar un �tem para poder iniciar una subasta.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Comprobamos que el usuario tenga lo que intenta ofertar
            If UserList(UserIndex).Invent.Object(Slot).Amount < Amount Then
                Call WriteConsoleMsg(UserIndex, "No tienes la cantidad de �tems que deseas subastar.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Items de Newbie no se pueden vender
            If ItemNewbie(UserList(UserIndex).Invent.Object(Slot).ObjIndex) = True Then
                Call WriteConsoleMsg(UserIndex, "No puedes subastar �tems de Newbie.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Runa no se puede vender
            If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).OBJType = otRuna Then
                Call WriteConsoleMsg(UserIndex, "��Estas loco!? �Si vendes tu Runa no podras volver a casa!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Actualizamos los datos
            .Actual = True
            .UserIndex = UserIndex
            
            .OfertaIndex = 0 ' Mientras quede 0 es por que no hay ofertas ;)
            .ValorBase = val(ValorBase)
            .OfertaMayor = val(.ValorBase) ' La oferta mayor es igual al valor inicial ;)
            
            ' Creamos el Objeto
            .Objeto.Amount = Amount
            .Objeto.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            
            .Tiempo = 3
            
            ' Quitamos el Objeto del usuario
            Call QuitarObjetos(.Objeto.ObjIndex, .Objeto.Amount, .UserIndex)
            
            ' Ahora podemos informar:
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] " & UserList(.UserIndex).Name & " est� subastando " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).Name & " con un valor inicial de " & .ValorBase, FontTypeNames.FONTTYPE_INFO))
 
            Exit Sub
        End If
    End With
End Sub
 
Public Sub Revisar_Subasta(ByVal UserIndex As Integer)
    With Subasta
        If UserIndex = .OfertaIndex Then
            .OfertaIndex = -1
        End If
        
        If UserIndex = .UserIndex Then
            .UserIndex = -1
        End If
    End With
End Sub
 
Public Sub Ofertar_Subasta(ByVal UserIndex As Integer, Oferta As Long)
    With Subasta
        ' Nos fijamos si existe la subasta
        If .Actual = False Then
            Call WriteConsoleMsg(UserIndex, "No hay ninguna subasta actualmente.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            ' �Tiene la cantidad de oro?
            If UserList(UserIndex).Stats.Gld < Oferta Then
                Call WriteConsoleMsg(UserIndex, "No tienes la cantidad de oro que intentas ofrecer.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' �Ya hay una oferta mayor? �Existia alguna oferta?
            If Oferta <= .OfertaMayor And .OfertaIndex <> 0 Then
                Call WriteConsoleMsg(UserIndex, "Tu oferta es menor a la oferta de " & val(.OfertaMayor) & " de " & UserList(.OfertaIndex).Name, FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Si no hay oferta, revisamos la nueva oferta para que no sea menor al precio base
            If .OfertaIndex = 0 And Oferta <= .ValorBase Then
                Call WriteConsoleMsg(UserIndex, "Tu oferta es menor a la oferta del valor inicial de " & val(.ValorBase), FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Si no hay oferta previa esto no sirve ;)
            If .OfertaIndex <> 0 Then
                ' Antes de actualizar devolvemos las cosas al flaco anterior
                UserList(.OfertaIndex).Stats.Gld = UserList(.OfertaIndex).Stats.Gld + val(.OfertaMayor)
                Call WriteUpdateGold(.OfertaIndex)
            End If
            
            ' Ahora podemos actualizar tranquilos
            .OfertaIndex = UserIndex
            .OfertaMayor = Oferta
            
            ' Restamos el oro:
            UserList(.OfertaIndex).Stats.Gld = UserList(.OfertaIndex).Stats.Gld - val(.OfertaMayor)
            Call WriteUpdateGold(.OfertaIndex)
            
            ' Informamos a los usuarios sobre la nueva oferta;
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] El usuario " & UserList(.OfertaIndex).Name & " aument� la oferta a " & .OfertaMayor, FontTypeNames.FONTTYPE_INFO))
        End If
    End With
End Sub
 
Public Sub Actualizar_Subasta()
    With Subasta
        ' Revisamos
        If .Actual Then
            ' Restamos tiempo
            .Tiempo = .Tiempo - 1
            
            ' Terminamos si es necesario, sino solo recordamos
            If .Tiempo <= 0 Then
                Call Termina_Subasta
            Else
                If .UserIndex <> -1 Then
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] El usuario " & UserList(.UserIndex).Name & " est� subastando " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).Name & ". La oferta actual es de " & .OfertaMayor & ". Esta subasta seguir� por " & .Tiempo & " minutos.", FontTypeNames.FONTTYPE_INFO))
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] Se est� subastando " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).Name & ". La oferta actual es de " & .OfertaMayor & ". Esta subasta seguir� por " & .Tiempo & " minutos.", FontTypeNames.FONTTYPE_INFO))
                End If
            End If
        End If
    End With
End Sub
 
Public Sub Termina_Subasta()
    With Subasta
        If .OfertaIndex = 0 Then
            ' Informamos que la subasta termino, y que nadie ofert�
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] La subasta de " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).Name & " termin� sin ninguna oferta.", FontTypeNames.FONTTYPE_INFO))
       
            If .UserIndex <> -1 Then
                Call MeterItemEnInventario(.UserIndex, .Objeto)
            End If
            
            'Reseteamos los datos
            .Actual = False
            .UserIndex = 0
            .OfertaIndex = 0
            .ValorBase = 0
            .OfertaMayor = 0
            .Tiempo = 0
        Else
            ' Informamos que la subasta termino, y quien se lleva las cosas.
            
            ' Entregamos el Item, y el Oro
            If .OfertaIndex <> -1 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] El usuario " & UserList(.OfertaIndex).Name & " gan� la subasta de " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).Name & " por la cantidad de " & .OfertaMayor, FontTypeNames.FONTTYPE_INFO))
    
                If MeterItemEnInventario(.OfertaIndex, .Objeto) Then
                    Call WriteConsoleMsg(.OfertaIndex, "Felicitaciones, has ganado la subasta de " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).Name & " por la cantidad de " & .OfertaMayor, FontTypeNames.FONTTYPE_INFOBOLD)
                End If
            End If
            
            ' Enviamos el Oro
            If .UserIndex <> -1 Then
                UserList(.UserIndex).Stats.Gld = UserList(.UserIndex).Stats.Gld + val(.OfertaMayor)
                Call WriteUpdateGold(.UserIndex)
            End If
            
            'Reseteamos los datos
            .Actual = False
            .UserIndex = 0
            .OfertaIndex = 0
            .ValorBase = 0
            .OfertaMayor = 0
            .Tiempo = 0
        End If
    End With
End Sub
 
