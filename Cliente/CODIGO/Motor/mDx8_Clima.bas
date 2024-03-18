Attribute VB_Name = "mDx8_Clima"
Option Explicit

'***************************************************
'Autor: Lorwik
'Descripción: Este sistema es una adaptación del que hice en
'las versiones anteriores de Imperium que posteriormente mejore en
'AODrag. El sistema fue adaptado al que trae AOLibre que a su vez
'se basaba en el de Blisse.
'***************************************************

Public Enum e_estados
    Amanecer = 0
    MedioDia
    Tarde
    noche
    Lluvia
    Niebla
    FogLluvia 'Niebla mas lluvia
End Enum

Public Estados(0 To 8) As D3DCOLORVALUE
Public Estado_Actual As D3DCOLORVALUE
Public Estado_Actual_Date As Byte

'****************************
'Usado para las particulas
'****************************

Private RainParticle As Long
Private NieveParticle As Long

Public OnRampage As Long
Public OnRampageImg As Long
Public OnRampageImgGrh As Integer

Public Enum eWeather
    Rain
    Nieve
End Enum

Private m_Hora_Actual As Long
Private m_Last_Hora_Actual As Long

Public Sub Init_MeteoEngine()
'***************************************************
'Author: Standelf
'Last Modification: 15/05/10
'Initializate
'***************************************************
    With Estados(e_estados.Amanecer)
        .a = 255
        .b = 230
        .r = 200
        .g = 200
    End With
    
    With Estados(e_estados.MedioDia)
        .a = 255
        .r = 255
        .g = 255
        .b = 255
    End With
    
    With Estados(e_estados.Tarde)
        .a = 255
        .b = 200
        .r = 230
        .g = 200
    End With
  
    With Estados(e_estados.noche)
        .a = 255
        .b = 170
        .r = 170
        .g = 170
    End With
    
    With Estados(e_estados.Lluvia)
        .a = 255
        .r = 200
        .g = 200
        .b = 200
    End With
    
    Estado_Actual_Date = 1
    
End Sub

Public Sub Actualizar_Estado(ByVal Estado As Byte)
'***************************************************
'Author: Lorwik
'Last Modification: 09/08/2020
'Actualiza el estado del clima y del dia
'***************************************************
    Dim X As Byte, Y As Byte

    'Primero actualizamos la imagen del frmmain
    'Call ActualizarImgClima

    '¿El mapa tiene su propia luz?
    If mapInfo.LuzBase <> -1 Then
        
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                Call Engine_Long_To_RGB_List(MapData(X, Y).Engine_Light(), mapInfo.LuzBase)
            Next Y
        Next X
        
        Call LightRenderAll
        
        Exit Sub
    End If

    '¿Es un estado invalido?
     Estado = e_estados.noche
        
    Estado_Actual = Estados(Estado)
    Estado_Actual_Date = Estado
        
    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Actual)
        Next Y
    Next X
        
    Call LightRenderAll
    
    If Estado = (e_estados.Lluvia Or e_estados.FogLluvia) Then
        If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
    
        bTecho = (MapData(UserPos.X, UserPos.Y).Trigger = eTrigger.BAJOTECHO Or _
            MapData(UserPos.X, UserPos.Y).Trigger = eTrigger.CASA Or _
            MapData(UserPos.X, UserPos.Y).Trigger = eTrigger.ZONASEGURA)
        
    End If

End Sub

Public Sub Start_Rampage()
'***************************************************
'Author: Standelf
'Last Modification: 27/05/2010
'Init Rampage
'***************************************************
    Dim X As Byte, Y As Byte, TempColor As D3DCOLORVALUE
    TempColor.a = 255: TempColor.b = 255: TempColor.r = 255: TempColor.g = 255
    
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), TempColor)
            Next Y
        Next X
End Sub

Public Sub End_Rampage()

    '***************************************************
    'Author: Standelf
    'Last Modification: 27/05/2010
    'End Rampage
    '***************************************************
    
    OnRampageImgGrh = 0
    OnRampageImg = 0
    
    Dim X As Byte, Y As Byte

    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Actual)
        Next Y
    Next X

    Call LightRenderAll

End Sub

Public Function bRain() As Boolean
'*****************************************************************
'Author: Lorwik
'Fecha: 13/08/2020
'Devuelve un True o un False si hay lluvia
'*****************************************************************

    If Estado_Actual_Date = e_estados.FogLluvia Or Estado_Actual_Date = e_estados.Lluvia Then
        bRain = True
        Exit Function
    End If
    
    bRain = False
End Function

Public Sub Engine_Weather_Update()
'*****************************************************************
'Author: Lorwik
'Fecha: 13/08/2020
'Controla los climas, aqui se renderizan la lluvia, nieve, etc.
'*****************************************************************

    '¿Esta lloviendo y no esta en dungeon?
    If bRain And MeterologiaEnDungeon Then
            
        'Particula segun el terreno...
        Select Case mapInfo.Terreno
        
            Case "BOSQUE", "DESIERTO"
                If RainParticle <= 0 Then
                    'Creamos las particulas de lluvia
                    Call mDx8_Clima.LoadWeatherParticles(eWeather.Rain)
                ElseIf RainParticle > 0 Then
                    Call mDx8_Particulas.Particle_Group_Render(RainParticle, 250, -1)
                End If
                
                'EXTRA: Relampagos
                If RandomNumber(1, 200000) < 20 Then
                    Call Sound.Sound_Play(SND_RELAMPAGO)
                    Start_Rampage
                    OnRampage = GetTickCount
                    OnRampageImg = OnRampage
                    OnRampageImgGrh = 2837
            End If
            
                If OnRampageImg <> 0 Then
                    If GetTickCount - OnRampageImg > 36 Then
                    
                        OnRampageImgGrh = OnRampageImgGrh + 1
                        If OnRampageImgGrh = 2847 Then OnRampageImgGrh = 2837
            
                        OnRampageImg = GetTickCount
                    End If
                End If
                
                If OnRampage <> 0 Then 'Hay Uno en curso
                    If GetTickCount - OnRampage > 400 Then
                        End_Rampage
                        OnRampage = 0
                    End If
                End If
                
                If OnRampageImgGrh <> 0 Then
                    Call Draw_GrhIndex(OnRampageImgGrh, 0, 0, 0, Normal_RGBList(), , True)
                End If
            
            Case "NIEVE"
            
                If NieveParticle <= 0 Then
                    'Creamos las particulas de nieve
                    Call mDx8_Clima.LoadWeatherParticles(eWeather.Nieve)
                ElseIf NieveParticle > 0 Then
                    Call mDx8_Particulas.Particle_Group_Render(NieveParticle, 250, -1)
                End If
        
        End Select
    
    Else '¿No esta lloviendo o dejo de llover?
        
        Call RemoveWeatherParticlesAll
            
    End If

End Sub

Public Sub LoadWeatherParticles(ByVal Weather As Byte)
'*****************************************************************
'Author: Lucas Recoaro (Recox)
'Last Modify Date: 19/12/2019
'Crea las particulas de clima.
'*****************************************************************
    Select Case Weather

        Case eWeather.Rain
            RainParticle = mDx8_Particulas.General_Particle_Create(8, -1, -1)
            
        Case eWeather.Nieve
            NieveParticle = mDx8_Particulas.General_Particle_Create(56, -1, -1)

    End Select
End Sub

Public Sub RemoveWeatherParticlesAll()
'*****************************************************************
'Author: Lorwik
'Last Modify Date: 13/08/2020
'Comprobamos si hay alguna particula climatologica activa para eliminarla
'*****************************************************************

    'Si alguna de las siguientes particulas esta cargada, la eliminamos
    If RainParticle > 0 Then
        Call mDx8_Clima.RemoveWeatherParticles(eWeather.Rain)
            
    ElseIf NieveParticle > 0 Then
        Call mDx8_Clima.RemoveWeatherParticles(eWeather.Nieve)
    
    End If
End Sub

Public Sub RemoveWeatherParticles(ByVal Weather As Byte)
'*****************************************************************
'Author: Lorwik
'Fecha: 14/08/2020
'Elimina las particulas climatologicas segun la que reciba
'*****************************************************************
    Select Case Weather

        Case eWeather.Rain
            Particle_Group_Remove (RainParticle)
            RainParticle = 0
            
        Case eWeather.Nieve
            Particle_Group_Remove (NieveParticle)
            NieveParticle = 0

    End Select
End Sub

Public Function MeterologiaEnDungeon() As Boolean
'*********************************************
'Autor: Lorwik
'Fecha: 26/10/2020
'Descripcion: Comprueba si hay algun fenomeno meteorologico activo y si esta en dungeon
''*********************************************
    If bRain And mapInfo.Zona <> "DUNGEON" Then
        
        MeterologiaEnDungeon = True
        
    Else
    
        MeterologiaEnDungeon = False
    
    End If
            
End Function

Public Sub Time_Logic(ByVal hora_ac As Byte)
'*********************************************
'Autor: Lorwik
'Fecha: 08/03/2021
'Descripcion: Actualiza la imagen del clima del Main
''*********************************************
    m_Hora_Actual = hora_ac

    If m_Hora_Actual <> m_Last_Hora_Actual Then
        frmMain.imgHora.Picture = General_Load_Picture_From_Resource(Format(m_Hora_Actual, "0#") & ".bmp")
        m_Last_Hora_Actual = m_Hora_Actual
    End If

End Sub

Private Sub ActualizarImgClima(ByVal Estado As Byte)

    If bRain Then
        frmMain.imgHora.Picture = General_Load_Picture_From_Resource("lluvia.bmp")
    
    Else
        frmMain.imgHora.Picture = General_Load_Picture_From_Resource(Format(m_Hora_Actual, "0#") & ".bmp")
        
    End If

End Sub
