Attribute VB_Name = "mDx8_Ambiente"
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: ??/??/10
'Blisse-AO | Sistema de Ambientes
'***************************************************

Option Explicit

Type A_Light
    range As Byte
    r As Integer
    g As Integer
    b As Integer
End Type

Type MapAmbientBlock
    Light As A_Light
    Particle As Byte
End Type

Type MapAmbient
    MapBlocks() As MapAmbientBlock
    OwnAmbientLight As D3DCOLORVALUE
    Fog As Integer
    Snow As Boolean
    Rain As Boolean
End Type

Public CurMapAmbient As MapAmbient

Public Sub Init_Ambient(ByVal Map As Integer)
'***************************************************
'Author: Standelf
'Last Modification: 15/10/10
'***************************************************
    With CurMapAmbient
        .Fog = -1
        .OwnAmbientLight.a = 255
        .OwnAmbientLight.r = 0
        .OwnAmbientLight.g = 0
        .OwnAmbientLight.b = 0
        
        Call Actualizar_Estado(Estado_Actual_Date)
            
    End With
End Sub
