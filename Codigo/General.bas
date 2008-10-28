Attribute VB_Name = "General"
''
' Module description.
' Can be more than one line.
'
' @remarks
' @author Marco Vanotti (marco@vanotti.com.ar)
' @version 0.0.1
' @date 20081005


Option Explicit


Public Type tAoUpdateFile
    name As String
    Version As Integer
    MD5 As String * 32
    Path As String
    HasPatches As Boolean
    Comment As String
End Type

''
' Loads the AoUpdate Ini File to an struct array
'
' @param file Specifies reference to AoUpdateIniFile
' @return an array of tAoUpdate
Public Function ReadAoUFile(file As String) As tAoUpdateFile()
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 27/10/2008
'
'*************************************************

    Dim Leer As New clsIniReader
    Dim NumFiles As Integer
    Dim tmpAoUFile() As tAoUpdateFile
    Dim i As Integer
    
    On Error GoTo error
    
    Call Leer.Initialize(file)
    
    NumFiles = Leer.GetValue("INIT", "NumFiles")
    
    ReDim tmpAoUFile(1 To NumFiles) As tAoUpdateFile
    
    For i = 1 To NumFiles
        tmpAoUFile(i).name = Leer.GetValue("File" & i, "Name")
        tmpAoUFile(i).Version = Leer.GetValue("File" & i, "Version")
        tmpAoUFile(i).MD5 = Leer.GetValue("File" & i, "MD5")
        If Leer.KeyExists("Path") Then tmpAoUFile(i).Path = Leer.GetValue("File" & i, "Path")
        If Leer.KeyExists("HasPatches") Then tmpAoUFile(i).HasPatches = CBool(Leer.GetValue("File" & i, "HasPatches"))
        If Leer.KeyExists("Comment") Then tmpAoUFile(i).Comment = Leer.GetValue("File" & i, "Comment")
    Next
    
    ReadAoUFile = tmpAoUFile
    
    Set Leer = Nothing
    
Exit Function
error:
    MsgBox Err.Description, vbCritical, Err.Number
    Set Leer = Nothing
End Function

Public Function compareUpdateFiles(localUpdateFile() As tAoUpdateFile, remoteUpdateFile() As tAoUpdateFile) As Byte()
    Dim i As Integer
    Dim j As Integer
    Dim tmpArr() As Byte
    
    ReDim tmpArr(0)
    
    For i = 1 To UBound(remoteUpdateFile)
        If i > UBound(localUpdateFile) Then
            ReDim Preserve tmpArr(UBound(tmpArr) + UBound(remoteUpdateFile) - UBound(localUpdateFile))
            
            For j = i To UBound(remoteUpdateFile)
                tmpArr(j) = j
            Next
            compareUpdateFiles = tmpArr
            Exit Function
        End If
        
        If remoteUpdateFile(i).name <> localUpdateFile(i).name Then
            MsgBox "Erro critico en los archivos ini. Por favor descargue el AoUpdater nuevamente."
        End If
                
        If remoteUpdateFile(i).Version <> localUpdateFile(i).Version Then
            'Version Diffs, add to download queue.
            ReDim Preserve tmpArr(UBound(tmpArr) + 1)
            tmpArr(UBound(tmpArr)) = i
        End If
    Next
    
    compareUpdateFiles = tmpArr
End Function



