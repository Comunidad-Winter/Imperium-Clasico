Attribute VB_Name = "modClimas"
'********************************Modulo Climas*********************************
'Autor: Lorwik
'Last Modification: 09/082020
'Controla el clima y lo envia al cliente.
'Nota: Cuando reformemos el sistema de lluvias, todo va a ir aqui.
'******************************************************************************
Option Explicit

Enum eColorEstado
    Amanecer = 0
    MedioDia
    Tarde
    Noche
    Lluvia
End Enum

Public DayStatus As eColorEstado 'Establece el color actual del dia

Public Sub SortearHorario(Optional ByVal Clima As eColorEstado)
'***************************************************************************************
'Autor: Lorwik
'Ultima modificación: 23/12/2018
'Descripción: Sorteamos el clima, si hay tormenta y es de Mañana o de Dia
'ponemos el efecto de tarde, pero si es de Tarde o de Noche no ponemos nigun efecto.
'***************************************************************************************

    'Si esta lloviendo ignoramos el resto y solo mandamos el estado lluvia
    If Lloviendo Then
    
        'Solo para el Main del server:
        Select Case Clima
        
            Case eColorEstado.Lluvia
                frmMain.lblLloviendoInfo.Caption = "Hora: Lloviendo - [" & Hour(Now) & ":" & Minute(Now) & "]"
                
        End Select
        
        Call ColorClima(Clima)
        
    Else
        If (Hour(Now) >= 5 And Hour(Now) < 8) Then 'Amanecer
            Call ColorClima(eColorEstado.Amanecer)
            frmMain.lblLloviendoInfo.Caption = "Hora: Mañana - [" & Hour(Now) & ":" & Minute(Now) & "]"
            
        ElseIf (Hour(Now) >= 9 And Hour(Now) < 12) Then  'MedioDia
            Call ColorClima(eColorEstado.MedioDia)
            frmMain.lblLloviendoInfo.Caption = "Hora: MedioDia - [" & Hour(Now) & ":" & Minute(Now) & "]"
            
        ElseIf (Hour(Now) >= 13 And Hour(Now) < 18) Then 'Tarde
            Call ColorClima(eColorEstado.Tarde)
            frmMain.lblLloviendoInfo.Caption = "Hora: Tarde - [" & Hour(Now) & ":" & Minute(Now) & "]"
            
        ElseIf (Hour(Now) >= 19 And Hour(Now) < 4) Then 'Noche
            Call ColorClima(eColorEstado.Noche)
            frmMain.lblLloviendoInfo.Caption = "Hora: Noche - [" & Hour(Now) & ":" & Minute(Now) & "]"
        End If
        
    End If
    
End Sub

'Enviamos el Clima
Private Sub ColorClima(Clima As eColorEstado)
'****************************************
'Autor: Lorwik
'Ultima modificación: 09/08/2020
'Enviamos el clima
'****************************************

    Dim UserIndex As Integer
    Dim i As Long
    
    DayStatus = Clima

    Call SendData(SendTarget.ToAll, 0, PrepareMessageActualizarClima())
    
End Sub

Public Sub SortearClima(Optional ByVal Forzar As Byte = 0)
'**********************************************
'Autor: Lorwik
'Ultima modificación: 09/08/2020
'Descripción: En este sub vamos a sortear si va lloviar
'**********************************************

    Dim Clima As eColorEstado
    
    '¿Esta lloviendo?
    If Lloviendo Then

        If Forzar = 0 Then
            'Por el momento seteamos la lluvia, ya que no requiere probs
            Clima = eColorEstado.Lluvia
            
        Else '¿Queremos forzar la aparicion de algun fenomeno?
        
            If Forzar = 1 Then
                Clima = eColorEstado.Lluvia
            
            End If
        End If
        
    End If

    'Sea cual sea el resultado, lo mandamos
    Call SortearHorario(Clima)
    
End Sub
