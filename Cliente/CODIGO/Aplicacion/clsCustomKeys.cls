VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsCustomKeys.cls - Allows the User to Customize Keys.
'
' Creation Date: 08/03/07
' Created and Implemented by Rapsodius
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
'Customizes User Keys.
'
'@author Rapsodius
'@version 1.1.0
'@date 20080413

'03/08/2007 - Rapsodius
'   - First Release
'
'13/04/2008 - Juan Martin Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - Removed lots of unwanted buggy features keeping it clean. Revised documentation and coding standards.


Option Explicit

''
'Number of Customizable Keys
Private Const NUM_CUSTOM_KEYS As Byte = 34

''
'Error Raised in case the key file doesn't exist
Private Const ERROR_NO_KEY_FILE As Long = vbObjectError + 16

''
'The relative path from the application's location where the key config file is.
Private Const KEY_CONFIG_FILE As String = "BindKeys.bin"

''
'Array of mapped keys
Private mappedKeys(1 To NUM_CUSTOM_KEYS) As Integer


''
'Index of keys on array
'
'@param mKeyUp Key for moving character up
'@param mKeyDown Key for moving character down
'@param mKeyLeft Key for moving character left
'@param mKeyRight Key for moving character right
'@param mKeyToggleMusic Key for toggling background music
'@param mKeyToggleSound Key for toggling sound
'@param mKeyRequestRefresh Key for sending a request refreshing packet
'@param mKeyToggleNames Key for toggling characters' names
'@param mKeyGetObject Key for retrieving an object from the ground
'@param mKeyEquipObject Key for equiping an inventory object
'@param mKeyTamAnimal Key for taming animals
'@param mKeySteal Key for Stealing other users' gold or objects
'@param mKeyToggleSafeMode Key for toggling Safe Mode
'@param mKeyToggleCombatSafe Key for toggling Resuscitation Safe
'@param mKeyHide Key for hiding the character
'@param mKeyDropObject Key for dropping inventory objects
'@param mKeyUseObject Key for using inventory objects
'@param mKeyAttack Key for attacking other Users/NPCs
'@param mKeyTalk Key for talking (to all))
'@param mKeyTakeScreenShot Key for taking a snapshot of the screen
'@param mKeyShowOptions Key for showing the "Options" form
'@param mKeyExitGame Key for close the game
'@param mKeyChatNormal Acceso rapido de chat normal
'@param mKeyChatGritar Acceso rapido de chat gritar
'@param mKeyChatPrivado Acceso rapido de chat privado
'@param mKeyChatClan Acceso rapido de chat clan
'@param mKeyChatGrupo Acceso rapido de chat grupo
'@param mKeyChatGlobal Acceso rapido de chat global
'@param mKeyQuests Acceso rapido al panel de Quests

Public Enum eKeyType
    mKeyUp = 1
    mKeyDown
    mKeyLeft
    mKeyRight
    mKeyToggleMusic
    mKeyToggleSound
    mKeyRequestRefresh
    mKeyToggleNames
    mKeyGetObject
    mKeyEquipObject
    mKeyTamAnimal
    mKeySteal
    mKeyToggleSafeMode
    mKeyToggleCombatSafe
    mKeyHide
    mKeyDropObject
    mKeyUseObject
    mKeyAttack
    mKeyTalk
    mKeyTakeScreenShot
    mKeyShowOptions
    mKeyExitGame
    mKeyToggleFPS
    mKeyFuncionUno
    mKeyFuncionDos
    mKeyFuncionTres
    mKeyFuncionCuatro
    mKeyFuncionCinco
    mKeyFuncionSeis
    mKeyFuncionSiete
    mKeyFuncionOcho
    mKeyFuncionNueve
    mKeyFuncionDiez
    mKeyFuncionOnce
End Enum

''
'Shows a readable name for a specific Key code
'
'@param KeyCode Key code to process
'@return An User readable name for the key code or a null string if the key is not valid
'@remarks All keys not listed here are considered not valid

Public Function ReadableName(ByVal KeyCode As Integer) As String
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyA To KeyCodeConstants.vbKeyZ
            ReadableName = Chr$(KeyCode)
        
        Case KeyCodeConstants.vbKeyNumpad0 To KeyCodeConstants.vbKeyNumpad9
            ReadableName = KeyCode - KeyCodeConstants.vbKeyNumpad0 & " (teclado numerico)"
            
        Case KeyCodeConstants.vbKey0 To KeyCodeConstants.vbKey9
            ReadableName = Chr$(KeyCode)
        
        Case KeyCodeConstants.vbKeyF1 To KeyCodeConstants.vbKeyF16
            ReadableName = "F" & (KeyCode - KeyCodeConstants.vbKeyF1 + 1)
        
        Case KeyCodeConstants.vbKeyMultiply
            ReadableName = "* (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyAdd
            ReadableName = "+ (teclado numerico)"
        
        Case KeyCodeConstants.vbKeySeparator
            'Note: Separator appears in MSDN as the Enter key for
            'the numpad, but MS recognizes it as the normal
            'Enter key
            ReadableName = "Enter (teclado numerico)"
        
        Case KeyCodeConstants.vbKeySubtract
            ReadableName = "- (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyDecimal
            ReadableName = ". (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyDivide
            ReadableName = "/ (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyShift
            ReadableName = "Shift"
        
        Case KeyCodeConstants.vbKeyControl
            ReadableName = "Control"
        
        Case KeyCodeConstants.vbKeyMenu 'Alt
            ReadableName = "Alt"
        
        Case KeyCodeConstants.vbKeyPause
            ReadableName = "Pausa"
        
        Case KeyCodeConstants.vbKeyCapital
            ReadableName = "Blq Mayus"
        
        Case KeyCodeConstants.vbKeyEscape
            ReadableName = "Esc"
        
        Case KeyCodeConstants.vbKeyPageUp
            ReadableName = "Ret Pag"
        
        Case KeyCodeConstants.vbKeyPageDown
            ReadableName = "Av Pag"
        
        Case KeyCodeConstants.vbKeyEnd
            ReadableName = "Fin"
        
        Case KeyCodeConstants.vbKeyHome
            ReadableName = "Inicio"
        
        Case KeyCodeConstants.vbKeyLeft
            ReadableName = "Izquierda"
        
        Case KeyCodeConstants.vbKeyUp
            ReadableName = "Arriba"
        
        Case KeyCodeConstants.vbKeyRight
            ReadableName = "Derecha"
        
        Case KeyCodeConstants.vbKeyDown
            ReadableName = "Abajo"
        
        Case KeyCodeConstants.vbKeyInsert
            ReadableName = "Insertar"
        
        Case KeyCodeConstants.vbKeyNumlock
            ReadableName = "Blq Num"
        
        Case KeyCodeConstants.vbKeyScrollLock
            ReadableName = "Blq Despl"
        
        Case KeyCodeConstants.vbKeySpace
            ReadableName = "Barra Espaciadora"
        
        Case KeyCodeConstants.vbKeyBack
            ReadableName = "Borrar"
        
        Case KeyCodeConstants.vbKeyReturn
            ReadableName = "Enter"
        
        Case KeyCodeConstants.vbKeyDelete
            ReadableName = "Suprimir"
            
        Case KeyCodeConstants.vbKeySnapshot
            ReadableName = "Impr. Pant."
        
        Case Else
            'In all other cases, key is considered not valid.
            ReadableName = Constants.vbNullString
    End Select
    
End Function

''
'Reads custom keys from a file.
'
'@remarks   If the custom keys file is not found ERROR_NO_KEY_FILE is raised.

Public Sub LoadCustomKeys()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/21/07
'
'***************************************************
    Dim i As Long
    Dim FileNum As Integer
    FileNum = FreeFile()
    
    'Inits keyboard config with default configurations
    Call LoadDefaults(True)
        
    Open Carga.Path(Init) & KEY_CONFIG_FILE For Binary Access Read Lock Read Write As FileNum
    
    For i = 1 To NUM_CUSTOM_KEYS
        Get FileNum, , mappedKeys(i)
        If LenB(ReadableName(mappedKeys(i))) = 0 Then
            Close FileNum
            Call Err.Raise(ERROR_NO_KEY_FILE)
            Exit Sub
        End If
    Next i
    
    Close FileNum
End Sub

''
' Saves custom keys to a file

Public Sub SaveCustomKeys()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    Dim FileNum As Integer
    Dim i As Long
    FileNum = FreeFile()
    
    Open Carga.Path(Init) & KEY_CONFIG_FILE For Binary Access Write As FileNum
    
    For i = 1 To NUM_CUSTOM_KEYS
        Put FileNum, , mappedKeys(i)
    Next i
    
    Close #FileNum
End Sub

''
'Gets a key code
'
'@param index Index of the key code to retrieve
'@return The Key code that belongs to index

Public Property Get BindedKey(ByVal Index As eKeyType) As Integer
'***************************************************
'Author: Rapsodius
'Last Modification: 16/11/2010
'16/11/2010: Amraphen - Adapte la propiedad para que cargue el bind de la tecla en la configuracion activa.
'***************************************************
    If Index < 1 Or Index > NUM_CUSTOM_KEYS Then Exit Property
    
    BindedKey = mappedKeys(Index)
End Property

''
'Sets a key code
'
'@param index Index of the key to change
'@param NewVal New key code

Public Property Let BindedKey(ByVal Index As eKeyType, ByVal NewVal As Integer)
'***************************************************
'Author: Rapsodius
'Last Modification: 16/11/2010
'16/11/2010: Amraphen - Adapte la propiedad para que guarde el bind de la tecla en la configuracion activa.
'***************************************************
    If Index < 1 Or Index > NUM_CUSTOM_KEYS Then Exit Property
    
    If LenB(ReadableName(NewVal)) = 0 Then 'If key is not valid...
        Exit Property                      'Rejects the new assignment
    End If
    
    If KeyAssigned(NewVal) Then 'Also reject it in case key is already assigned
        Exit Property
    End If
    
    mappedKeys(Index) = NewVal
End Property

''
'Loads default keys
'
'@param KeyConfig Index of the key configuration to change

Public Sub LoadDefaults(ByVal Clasica As Boolean)
'***************************************************
'Author: Rapsodius
'Last Modification: 16/11/2010
'16/11/2010: Amraphen - Adapte el procedimiento para que cargue los bindings por default en una sola configuracion.
'27/04/2020: Lorwik - Añado 2 tipos de configuraciones por defecto
'***************************************************
    
    If Clasica Then 'Configuración clasica con las flechitas
        mappedKeys(eKeyType.mKeyUp) = KeyCodeConstants.vbKeyUp
        mappedKeys(eKeyType.mKeyDown) = KeyCodeConstants.vbKeyDown
        mappedKeys(eKeyType.mKeyLeft) = KeyCodeConstants.vbKeyLeft
        mappedKeys(eKeyType.mKeyRight) = KeyCodeConstants.vbKeyRight
        mappedKeys(eKeyType.mKeyGetObject) = KeyCodeConstants.vbKeyA
        mappedKeys(eKeyType.mKeyTamAnimal) = KeyCodeConstants.vbKeyD
        mappedKeys(eKeyType.mKeyDropObject) = KeyCodeConstants.vbKeyT
        mappedKeys(eKeyType.mKeyUseObject) = KeyCodeConstants.vbKeyU
        mappedKeys(eKeyType.mKeyAttack) = KeyCodeConstants.vbKeyControl
        mappedKeys(eKeyType.mKeyToggleCombatSafe) = KeyCodeConstants.vbKeyS
        
    Else 'Configuracion con AWSD
        mappedKeys(eKeyType.mKeyUp) = KeyCodeConstants.vbKeyW
        mappedKeys(eKeyType.mKeyDown) = KeyCodeConstants.vbKeyS
        mappedKeys(eKeyType.mKeyLeft) = KeyCodeConstants.vbKeyA
        mappedKeys(eKeyType.mKeyRight) = KeyCodeConstants.vbKeyD
        mappedKeys(eKeyType.mKeyGetObject) = KeyCodeConstants.vbKeyG
        mappedKeys(eKeyType.mKeyTamAnimal) = KeyCodeConstants.vbKeyP
        mappedKeys(eKeyType.mKeyUseObject) = KeyCodeConstants.vbKeySpace
        mappedKeys(eKeyType.mKeyAttack) = KeyCodeConstants.vbKeyF
        mappedKeys(eKeyType.mKeyToggleCombatSafe) = KeyCodeConstants.vbKeyV
        
    End If
    
    mappedKeys(eKeyType.mKeyToggleMusic) = KeyCodeConstants.vbKeyM
    mappedKeys(eKeyType.mKeyToggleSound) = KeyCodeConstants.vbKeyS
    mappedKeys(eKeyType.mKeyRequestRefresh) = KeyCodeConstants.vbKeyL
    mappedKeys(eKeyType.mKeyToggleNames) = KeyCodeConstants.vbKeyN
    mappedKeys(eKeyType.mKeyEquipObject) = KeyCodeConstants.vbKeyE
    mappedKeys(eKeyType.mKeySteal) = KeyCodeConstants.vbKeyR
    mappedKeys(eKeyType.mKeyToggleCombatSafe) = KeyCodeConstants.vbKeyC
    mappedKeys(eKeyType.mKeyHide) = KeyCodeConstants.vbKeyO
    mappedKeys(eKeyType.mKeyDropObject) = KeyCodeConstants.vbKeyT
    mappedKeys(eKeyType.mKeyTalk) = KeyCodeConstants.vbKeyReturn
    mappedKeys(eKeyType.mKeyTakeScreenShot) = KeyCodeConstants.vbKeySnapshot
    mappedKeys(eKeyType.mKeyShowOptions) = KeyCodeConstants.vbKeyI
    mappedKeys(eKeyType.mKeyExitGame) = KeyCodeConstants.vbKeyEscape
    mappedKeys(eKeyType.mKeyToggleFPS) = KeyCodeConstants.vbKeyMultiply
    
    mappedKeys(eKeyType.mKeyFuncionUno) = KeyCodeConstants.vbKeyF1
    mappedKeys(eKeyType.mKeyFuncionDos) = KeyCodeConstants.vbKeyF2
    mappedKeys(eKeyType.mKeyFuncionTres) = KeyCodeConstants.vbKeyF3
    mappedKeys(eKeyType.mKeyFuncionCuatro) = KeyCodeConstants.vbKeyF4
    mappedKeys(eKeyType.mKeyFuncionCinco) = KeyCodeConstants.vbKeyF5
    mappedKeys(eKeyType.mKeyFuncionSeis) = KeyCodeConstants.vbKeyF6
    mappedKeys(eKeyType.mKeyFuncionSiete) = KeyCodeConstants.vbKeyF7
    mappedKeys(eKeyType.mKeyFuncionOcho) = KeyCodeConstants.vbKeyF8
    mappedKeys(eKeyType.mKeyFuncionNueve) = KeyCodeConstants.vbKeyF9
    mappedKeys(eKeyType.mKeyFuncionDiez) = KeyCodeConstants.vbKeyF10
    mappedKeys(eKeyType.mKeyFuncionOnce) = KeyCodeConstants.vbKeyF11
    
End Sub

''
'Returns whether a key is binded to an action
'
'@param KeyCode Key code to check
'@return True if key is mapped, false otherwise

Public Function KeyAssigned(ByVal KeyCode As Integer) As Boolean
'***************************************************
'Author: Rapsodius
'Last Modification: 16/11/2010
'16/11/2010: Amraphen - Adapte la funcion para que trabaje sobre la configuracion activa.
'***************************************************
Dim Counter As Long
    
    KeyAssigned = False
    
    For Counter = 1 To NUM_CUSTOM_KEYS
        If mappedKeys(Counter) = KeyCode Then
            KeyAssigned = True
            Exit For
        End If
    Next Counter
End Function

''
'Retrieves the number of customizable keys
'
'@return The number of customizable keys

Public Property Get KeyCount() As Byte
'***************************************************
'Author: Rapsodius
'Last Modification: 16/11/2010
'16/11/2010: Amraphen - Cambie el nombre para evitar confusiones.
'***************************************************
    KeyCount = NUM_CUSTOM_KEYS
End Property

''
'Constructor. Loads Keys from file

Private Sub Class_Initialize()
'***************************************************
'Author: Rapsodius
'Last Modification: 27/04/2020
'27/04/2020: Lorwik - Ahora podemos manejar una configuracion default y una alternativa.
'***************************************************
    
On Error GoTo ErrHand:

    Call LoadCustomKeys
    Exit Sub
    
ErrHand:
    If Err.number = ERROR_NO_KEY_FILE Then
        'Por defecto usaremos la clasica
        Call LoadDefaults(True)
        Call SaveCustomKeys
        Resume Next
    End If
End Sub

''
'Destructor. Saves Keys to file

Private Sub Class_Terminate()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
On Error Resume Next

    Call SaveCustomKeys
    
End Sub

