Attribute VB_Name = "mDx8_Dibujado"
Option Explicit

Private DrawBuffer As cDIBSection

Sub DrawGrhtoHdc(ByRef Pic As PictureBox, _
                 ByVal GrhIndex As Long, _
                 ByRef DestRect As RECT)

    '*****************************************************************
    'Draws a Grh's portion to the given area of any Device Context
    '*****************************************************************
         
    DoEvents
    
    Pic.AutoRedraw = False
        
    'Clear the inventory window
    Call Engine_BeginScene
        
    Call Draw_GrhIndex(GrhIndex, 0, 0, 0, Normal_RGBList())
        
    Call Engine_EndScene(DestRect, Pic.hWnd)
    
    Call DrawBuffer.LoadPictureBlt(Pic.hDC)

    Pic.AutoRedraw = True

    Call DrawBuffer.PaintPicture(Pic.hDC, 0, 0, Pic.Width, Pic.Height, 0, 0, vbSrcCopy)

    Pic.Picture = Pic.Image
        
End Sub

Public Sub PrepareDrawBuffer()
    Set DrawBuffer = New cDIBSection
    'El tamanio del buffer es arbitrario = 1024 x 1024
    Call DrawBuffer.Create(1024, 1024)
End Sub

Public Sub CleanDrawBuffer()
    Set DrawBuffer = Nothing
End Sub

Public Sub DrawPJ(ByVal Index As Byte)

    If LenB(cPJ(Index).Nombre) = 0 Then Exit Sub
    DoEvents
    
    Dim cColor       As Long
    Dim Head_OffSet  As Integer
    Dim PixelOffsetX As Integer
    Dim PixelOffsetY As Integer
    Dim RE           As RECT
    
    If cPJ(Index).GameMaster Then
        cColor = 2004510
    Else
        cColor = IIf(cPJ(Index).Criminal, 255, 16744448)
    End If
    
    With frmCharList.lblAccData(Index)
        .Caption = cPJ(Index).Nombre
        .ForeColor = cColor
    End With
    
    With frmCharList.picChar(Index - 1)
        RE.Left = 0
        RE.Top = 0
        RE.Bottom = .Height
        RE.Right = .Width
    End With

    PixelOffsetX = RE.Right \ 2 - 16
    PixelOffsetY = RE.Bottom \ 2
    
    Call Engine_BeginScene
    
    With cPJ(Index)
    
        If .Body <> 0 Then

            Call Draw_Grh(BodyData(.Body).Walk(3), PixelOffsetX, PixelOffsetY, 1, Normal_RGBList(), 0)

            If .Head <> 0 Then _
                Call Draw_Grh(HeadData(.Head).Head(3), PixelOffsetX + BodyData(.Body).HeadOffset.X + 1, PixelOffsetY + BodyData(.Body).HeadOffset.Y, 1, Normal_RGBList(), 0)

            If .helmet <> 0 Then _
                Call Draw_Grh(CascoAnimData(.helmet).Head(3), PixelOffsetX + BodyData(.Body).HeadOffset.X + 1, PixelOffsetY + BodyData(.Body).HeadOffset.Y + OFFSET_HEAD, 1, Normal_RGBList(), 0)

            If .weapon <> 0 Then _
                Call Draw_Grh(WeaponAnimData(.weapon).WeaponWalk(3), PixelOffsetX, PixelOffsetY, 1, Normal_RGBList(), 0)

            If .shield <> 0 Then _
                Call Draw_Grh(ShieldAnimData(.shield).ShieldWalk(3), PixelOffsetX, PixelOffsetY, 1, Normal_RGBList(), 0)
        
        End If
    
    End With

    Call Engine_EndScene(RE, frmCharList.picChar(Index - 1).hWnd)

    Call DrawBuffer.LoadPictureBlt(frmCharList.picChar(Index - 1).hDC)

    frmCharList.picChar(Index - 1).AutoRedraw = True

    Call DrawBuffer.PaintPicture(frmCharList.picChar(Index - 1).hDC, 0, 0, RE.Right, RE.Bottom, 0, 0, vbSrcCopy)

    frmCharList.picChar(Index - 1).Picture = frmCharList.picChar(Index - 1).Image
    
End Sub

Public Sub DibujarMenuMacros(Optional ActualizarCual As Byte = 0)
'************************************
'Autor: Lorwik
'Fecha: 07/03/2021
'Descripcion: Dibuja los macros del frmmain
'***********************************

    Dim i As Integer
    
    If ActualizarCual <= 0 Then
    
        For i = 1 To NUMMACROS
            Select Case MacrosKey(i).TipoAccion
                Case 1 'Envia comando
                    Call RenderItem(frmMain.picMacro(i - 1), 17506)
                    frmMain.picMacro(i - 1).ToolTipText = "Enviar comando: " & MacrosKey(i).Comando
                    
                Case 2 'Lanza hechizo
                    Call RenderItem(frmMain.picMacro(i - 1), 609)
                    frmMain.picMacro(i - 1).ToolTipText = "Lanzar hechizo: " & MacrosKey(i).SpellName
                    
                Case 3 'Equipa
                    If MacrosKey(i).InvGrh > 0 Then
                        Call RenderItem(frmMain.picMacro(i - 1), MacrosKey(i).InvGrh)
                        frmMain.picMacro(i - 1).ToolTipText = "Equipar objeto: " & MacrosKey(i).invName
                    End If
                    
                Case 4 'Usa
                    If MacrosKey(i).InvGrh > 0 Then
                        Call RenderItem(frmMain.picMacro(i - 1), MacrosKey(i).InvGrh)
                        frmMain.picMacro(i - 1).ToolTipText = "Usar objeto: " & MacrosKey(i).invName
                    End If
                    
                End Select
        Next i
    
    Else
    
        Select Case MacrosKey(ActualizarCual).TipoAccion
            Case 1 'Envia comando
                Call RenderItem(frmMain.picMacro(ActualizarCual - 1), 17506)
                frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Enviar comando: " & MacrosKey(ActualizarCual).Comando
                
            Case 2 'Lanza hechizo
                Call RenderItem(frmMain.picMacro(ActualizarCual - 1), 609)
                frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Lanzar hechizo: " & MacrosKey(ActualizarCual).SpellName
                
            Case 3 'Equipa
                If MacrosKey(ActualizarCual).InvGrh > 0 Then
                    Call RenderItem(frmMain.picMacro(ActualizarCual - 1), MacrosKey(ActualizarCual).InvGrh)
                    frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Equipar objeto: " & MacrosKey(ActualizarCual).invName
                End If
                
            Case 4 'Usa
                If MacrosKey(ActualizarCual).InvGrh > 0 Then
                    Call RenderItem(frmMain.picMacro(ActualizarCual - 1), MacrosKey(ActualizarCual).InvGrh)
                    frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Usar objeto: " & MacrosKey(ActualizarCual).invName
                End If
        End Select
    
        frmMain.picMacro(ActualizarCual - 1).Refresh
    
    End If

End Sub

Private Function EaseOutCubic(Time As Double)
    Time = Time - 1
    EaseOutCubic = Time * Time * Time + 1
End Function
