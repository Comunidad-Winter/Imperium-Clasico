Attribute VB_Name = "modCompression"
Option Explicit

'Public Formato As String * 6
Public Const Formato As String * 6 = ".IAC"

Public PkContra() As Byte

'This structure will describe our binary file's
'size and number of contained files
Public Type FILEHEADER
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    lngNumFiles As Long                 'How many files are inside?
End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER
    
    lngFileStart As Long            'Where does the chunk start?
    lngFileSize As Long             'How big is this chunk of stored data?
    strFileName As String * 32      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed
End Type

Public Enum srcFileType
    Graphics
    Ambient
    Music
    Midi
    Wav
    Scripts
    Map
    Interface
    Fuentes
    Skin
    Patch
End Enum

Private Const SrcPath As String = "\Recursos\"
Public Windows_Temp_Dir As String

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, bytesTotal As Currency, FreeBytesTotal As Currency) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function UnCompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

'Loading pictures from byte arrays
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Public Sub GenerateContra()
'***************************************************
'Author: ^[GS]^
'Last Modification: 17/06/2012 - ^[GS]^
'
'***************************************************

'on error resume next
    Dim Contra As String
    Dim LoopC As Byte
    
    Contra = "$FlLrjB3JoliHdAPKA8&YaJR5"
    
    Erase PkContra
    
    If LenB(Contra) <> 0 Then
        ReDim PkContra(Len(Contra) - 1)
        For LoopC = 0 To UBound(PkContra)
            PkContra(LoopC) = Asc(mid(Contra, LoopC + 1, 1))
        Next LoopC
    End If
    
End Sub

Public Sub Compress_Data(ByRef Data() As Byte)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Compresses binary data avoiding data loses
'*****************************************************************

    Dim Dimensions As Long
    Dim DimBuffer As Long
    Dim BufTemp() As Byte
    Dim BufTemp2() As Byte
    Dim LoopC As Long
    
    'Get size of the uncompressed data
    Dimensions = UBound(Data)
    
    'Prepare a buffer 1.06 times as big as the original size
    DimBuffer = Dimensions * 1.06
    ReDim BufTemp(DimBuffer)
    
    'Compress data using zlib
    Compress BufTemp(0), DimBuffer, Data(0), Dimensions
    
    'Deallocate memory used by uncompressed data
    Erase Data
    
    'Get rid of unused bytes in the compressed data buffer
    ReDim Preserve BufTemp(DimBuffer - 1)
    
    'Copy the compressed data buffer to the original data array so it will return to caller
    Data = BufTemp
    
    'Deallocate memory used by the temp buffer
    Erase BufTemp
    
    'Encrypt the first byte of the compressed data for extra security
    If UBound(PkContra) <= UBound(Data) And UBound(PkContra) <> 0 Then
        For LoopC = 0 To UBound(PkContra)
            Data(LoopC) = Data(LoopC) Xor PkContra(LoopC)
        Next LoopC
    End If
    
End Sub

Public Sub Decompress_Data(ByRef Data() As Byte, ByVal OrigSize As Long)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Decompresses binary data
'*****************************************************************

    Dim BufTemp() As Byte
    Dim LoopC As Integer
    
    ReDim BufTemp(OrigSize - 1)
    
    'Des-encrypt the first byte of the compressed data
    If UBound(PkContra) <= UBound(Data) And UBound(PkContra) <> 0 Then
        For LoopC = 0 To UBound(PkContra)
            Data(LoopC) = Data(LoopC) Xor PkContra(LoopC)
        Next LoopC
    End If
    
    UnCompress BufTemp(0), OrigSize, Data(0), UBound(Data) + 1
    
    ReDim Data(OrigSize - 1)
    
    Data = BufTemp
    
    Erase BufTemp
    
End Sub

Private Sub encryptHeaderFile(ByRef FileHead As FILEHEADER)
    'Each different variable is encrypted with a different key for extra security
    With FileHead
        .lngNumFiles = .lngNumFiles Xor 37816
        .lngFileSize = .lngFileSize Xor 245378169
    End With
End Sub

Private Sub encryptHeaderInfo(ByRef InfoHead As INFOHEADER)
    Dim EncryptedFileName As String
    Dim LoopC As Long
    
    For LoopC = 1 To Len(InfoHead.strFileName)
        If LoopC Mod 2 = 0 Then
            EncryptedFileName = EncryptedFileName & Chr(Asc(mid(InfoHead.strFileName, LoopC, 1)) Xor 12)
        Else
            EncryptedFileName = EncryptedFileName & Chr(Asc(mid(InfoHead.strFileName, LoopC, 1)) Xor 23)
        End If
    Next LoopC
    
    'Each different variable is encrypted with a different key for extra security
    With InfoHead
        .lngFileSize = .lngFileSize Xor 221872469
        .lngFileSizeUncompressed = .lngFileSizeUncompressed Xor 447915732
        .lngFileStart = .lngFileStart Xor 172379447
        .strFileName = EncryptedFileName
    End With
End Sub

Public Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 6/07/2004
'
'**************************************************************

    Dim RetVal As Long
    Dim FB As Currency
    Dim BT As Currency
    Dim FBT As Currency
    
    RetVal = GetDiskFreeSpace(Left(DriveName, 2), FB, BT, FBT)
    
    General_Drive_Get_Free_Bytes = FB * 10000 'convert result to actual size in bytes
End Function

Public Function General_Get_Temp_Dir() As String
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'Gets windows temporary directory
'**************************************************************

 Const MAX_LENGTH = 512
   Dim s As String
   Dim c As Long
   s = Space$(MAX_LENGTH)
   c = GetTempPath(MAX_LENGTH, s)
   If c > 0 Then
       If c > Len(s) Then
           s = Space$(c + 1)
           c = GetTempPath(MAX_LENGTH, s)
       End If
   End If
   General_Get_Temp_Dir = IIf(c > 0, Left$(s, c), "")
End Function

Public Function extractMusic(ByVal file_name As String, Optional ByVal Midi As Boolean = False) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a resource file
'*****************************************************************

    Dim LoopC As Long
    
    Dim SourceFilePath As String
    Dim OutputFilePath As String
    
    Dim SourceData() As Byte
    Dim InfoHead As INFOHEADER
    Dim handle As Integer
    
'Set up the error handler
On Local Error GoTo errhandler
    
    '¿Queremos descomprimir en la carpeta temporal?
    OutputFilePath = Windows_Temp_Dir
    
    If Midi = False Then
        'Find the Info Head of the desired file
        SourceFilePath = App.Path & SrcPath & "Musica" & Formato
        InfoHead = File_Find(SourceFilePath, file_name & ".mp3")
    Else
        'Find the Info Head of the desired file
        SourceFilePath = App.Path & SrcPath & "Midi" & Formato
        InfoHead = File_Find(SourceFilePath, file_name & ".mid")
    End If
    
    If InfoHead.strFileName = "" Or InfoHead.lngFileSize = 0 Then
        extractMusic = False
        Exit Function
    End If

    'Open the binary file
    handle = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As handle
    
    'Check the file for validity
    'If LOF(handle) <> InfoHead.lngFileSize Then
    '    Close handle
    '    MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
    '    Exit Function
    'End If
    
    'Make sure there is enough space in the HD
    If InfoHead.lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
        Close handle
        MsgBox "There is not enough drive space to extract the compressed file.", , "Error"
        extractMusic = False
        Exit Function
    End If
    
    'Extract file from the binary file
    
    'Resize the byte data array
    ReDim SourceData(InfoHead.lngFileSize - 1)
    
    'Get the data
    Get handle, InfoHead.lngFileStart, SourceData
    
    'Decompress all data
    Decompress_Data SourceData, InfoHead.lngFileSizeUncompressed
    
    'Close the binary file
    Close handle
    
    'Get a free handler
    handle = FreeFile
    
    Open OutputFilePath & InfoHead.strFileName For Binary As handle
    
    Put handle, 1, SourceData
    
    Close handle
    
    Erase SourceData
        
    extractMusic = True
Exit Function

errhandler:
    Close handle
    Erase SourceData
    extractMusic = False
    'Display an error message if it didn't work
    'MsgBox "Unable to decode binary file. Reason: " & Err.number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Function Extract_File_Memory(ByVal File_Type As srcFileType, ByVal file_name As String, ByRef SourceData() As Byte) As Boolean
 '********************************************
'Author: ???
'Last Modify Date: ???
'Extra archivos en memoria
'*********************************************

    Dim LoopC As Long
    Dim SourceFilePath As String
    Dim InfoHead As INFOHEADER
    Dim handle As Integer
   
On Local Error GoTo errhandler
   
    Select Case File_Type
    
        Case Graphics
                SourceFilePath = App.Path & SrcPath & "Graficos" & Formato
            
        Case Music
                SourceFilePath = App.Path & SrcPath & "Musica" & Formato
                
        Case Midi
                SourceFilePath = App.Path & SrcPath & "Midi" & Formato
        
        Case Wav
                SourceFilePath = App.Path & SrcPath & "Sounds" & Formato

        Case Scripts
                SourceFilePath = App.Path & SrcPath & "Scripts" & Formato

        Case Interface
                SourceFilePath = App.Path & SrcPath & "Interface" & Formato

        Case Map
                SourceFilePath = App.Path & SrcPath & "Mapas" & Formato

        Case Ambient
                SourceFilePath = App.Path & SrcPath & "Ambient" & Formato
                
        Case Fuentes
                SourceFilePath = App.Path & SrcPath & "Fuentes" & Formato
                
        Case Skin
                SourceFilePath = App.Path & SrcPath & "Skins\" & ClientSetup.SkinSeleccionado & Formato
                
        Case Else
            Exit Function
    End Select
   
    InfoHead = File_Find(SourceFilePath, file_name)
   
    If InfoHead.strFileName = "" Or InfoHead.lngFileSize = 0 Then Exit Function
 
    handle = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As handle
   
    If InfoHead.lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
        Close handle
        MsgBox "There is not enough drive space to extract the compressed file.", , "Error"
        Exit Function
    End If
   
   
    ReDim SourceData(InfoHead.lngFileSize - 1)
   
    Get handle, InfoHead.lngFileStart, SourceData
        Decompress_Data SourceData, InfoHead.lngFileSizeUncompressed
    Close handle
       
    Extract_File_Memory = True
Exit Function
 
errhandler:
    Close handle
    Erase SourceData
End Function

Public Sub DeleteFile(ByVal file_path As String)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 3/03/2005
'Deletes a resource files
'*****************************************************************

    Dim handle As Integer
    Dim Data() As Byte
    
    On Error GoTo ERROR_HANDLER
    
    'We open the file to delete
    handle = FreeFile
    Open file_path For Binary Access Write Lock Read As handle
    
    'We replace all the bytes in it with 0s
    ReDim Data(LOF(handle) - 1)
    Put handle, 1, Data
    
    'We close the file
    Close handle
    
    'Now we delete it, knowing that if they retrieve it (some antivirus may create backup copies of deleted files), it will be useless
    Kill file_path
    
    Exit Sub
    
ERROR_HANDLER:
    Kill file_path
        
End Sub

Public Function General_File_Exists(ByVal file_path As String, ByVal File_Type As VbFileAttribute) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a file exists
'*****************************************************************

    If Dir(file_path, File_Type) = "" Then
        General_File_Exists = False
    Else
        General_File_Exists = True
    End If
End Function

Public Sub General_Quick_Sort(ByRef SortArray As Variant, ByVal First As Long, ByVal Last As Long)
 '********************************************
'Author: ???
'Last Modify Date: ???
'Extra archivos en memoria
'*********************************************

    Dim Low As Long, High As Long
    Dim temp As Variant
    Dim List_Separator As Variant
   
    Low = First
    High = Last
    List_Separator = SortArray((First + Last) / 2)
    Do While (Low <= High)
        Do While SortArray(Low) < List_Separator
            Low = Low + 1
        Loop
        Do While SortArray(High) > List_Separator
            High = High - 1
        Loop
        If Low <= High Then
            temp = SortArray(Low)
            SortArray(Low) = SortArray(High)
            SortArray(High) = temp
            Low = Low + 1
            High = High - 1
        End If
    Loop
    If First < High Then General_Quick_Sort SortArray, First, High
    If Low < Last Then General_Quick_Sort SortArray, Low, Last
End Sub

Public Function File_Find(ByVal resource_file_path As String, ByVal file_name As String) As INFOHEADER
 '********************************************
'Author: ???
'Last Modify Date: ???
'Extra archivos en memoria
'*********************************************
 
On Error GoTo errhandler
 
    Dim Max As Integer
    Dim Min As Integer
    Dim mid As Integer
    Dim file_handler As Integer
   
    Dim file_head As FILEHEADER
    Dim info_head As INFOHEADER
   
    If Len(file_name) < Len(info_head.strFileName) Then _
        file_name = file_name & Space$(Len(info_head.strFileName) - Len(file_name))
   
    file_handler = FreeFile
    Open resource_file_path For Binary Access Read Lock Write As file_handler
   
    Get file_handler, 1, file_head

    'Desencrypt File Header
    encryptHeaderFile file_head
   
    Min = 1
    Max = file_head.lngNumFiles
   
    Do While Min <= Max
        mid = (Min + Max) / 2
       
        Get file_handler, CLng(Len(file_head) + CLng(Len(info_head)) * CLng((mid - 1)) + 1), info_head
        
        'Once an InfoHead index is ready, we encrypt it
        encryptHeaderInfo info_head
               
        If file_name < info_head.strFileName Then
            If Max = mid Then
                Max = Max - 1
            Else
                Max = mid
            End If
        ElseIf file_name > info_head.strFileName Then
            If Min = mid Then
                Min = Min + 1
            Else
                Min = mid
            End If
        Else
            File_Find = info_head
           
            Close file_handler
            Exit Function
        End If
    Loop
   
errhandler:
    Close file_handler
    File_Find.strFileName = ""
    File_Find.lngFileSize = 0
End Function

Public Function General_Load_Picture_From_Resource(ByVal picture_file_name As String, Optional ByVal Skin As Boolean = False) As IPicture
'*****************************************
'Autor: Lorwik
'Fecha: 24/08/2020
'Descripción: Carga un archivo de interfaz de los recursos comprimidos
'*****************************************

On Error GoTo ErrorHandler

    Dim bytArr() As Byte
    
    If Skin Then
        If Extract_File_Memory(srcFileType.Skin, picture_file_name, bytArr()) Then
            Set General_Load_Picture_From_Resource = General_Load_Picture_From_BArray(bytArr())
        Else
            Set General_Load_Picture_From_Resource = Nothing
        End If
        
        Exit Function
        
    Else
        If Extract_File_Memory(srcFileType.Interface, picture_file_name, bytArr()) Then
            Set General_Load_Picture_From_Resource = General_Load_Picture_From_BArray(bytArr())
        Else
            Set General_Load_Picture_From_Resource = Nothing
        End If
        
        Exit Function
        
    End If

ErrorHandler:

End Function

Public Function General_Load_Picture_From_BArray(ByRef bytArr() As Byte) As IPicture
'*****************************************
'Autor: ???
'Fecha: ?????
'Descripción: Carga imagenes desde la memoria
'*****************************************
On Error GoTo ErrorHandler

    Dim LowerBound As Long
    Dim ByteCount As Long
    Dim hMem As Long
    Dim lpMem As Long
    Dim IID_IPicture(15) As Long
    Dim istm As stdole.IUnknown
        
    LowerBound = LBound(bytArr)
    ByteCount = (UBound(bytArr) - LowerBound) + 1
    hMem = GlobalAlloc(&H2, ByteCount)
    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)
        If lpMem <> 0 Then
            MoveMemory ByVal lpMem, bytArr(LowerBound), ByteCount
            Call GlobalUnlock(hMem)
            If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
                If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                    Call OleLoadPicture(ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture(0), General_Load_Picture_From_BArray)
                End If
            End If
        End If
    End If
    
    Exit Function

ErrorHandler:

End Function


