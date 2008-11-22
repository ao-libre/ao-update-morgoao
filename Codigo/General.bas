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

Public Const UPDATES_SITE As String = "http://www.argentuuum.com.ar/aoupdate/"
Public Const AOUPDATE_FILE As String = "AoUpdate.ini"
Public DownloadsPath As String

Public Type tAoUpdateFile
    name As String
    Version As Integer
    MD5 As String * 32
    Path As String
    HasPatches As Boolean
    Comment As String
    DownloadError As Boolean
End Type

Public AoUpdateFileDownloaded As Boolean

Public AoUpdateRemote() As tAoUpdateFile
Public AoUpdateLocal() As tAoUpdateFile
Public DownloadQueue() As Byte

'Cola de archivos a decargar
Public ColDownloadQueue As New Collection

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
    
'On Error GoTo error
    
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
    Next i
    
    ReadAoUFile = tmpAoUFile
    
    Set Leer = Nothing
Exit Function

error:
    MsgBox Err.Description, vbCritical, Err.Number
    Set Leer = Nothing
End Function

''
' Compares the local AoUpdate file with the one in the server
'
' @param localUpdateFile Specifies reference to Local Update File
' @param remoteUpdateFile Specifies reference to Remote Update File
' @return an array of bytes with the updates queue.
Public Function compareUpdateFiles(localUpdateFile() As tAoUpdateFile, remoteUpdateFile() As tAoUpdateFile) As Byte()
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 27/10/2008
'
'*************************************************
    Dim i As Long
    Dim j As Long
    Dim tmpArrIndex As Integer
    Dim tmpArr() As Byte
    
    ReDim tmpArr(0)
    
    For i = 1 To UBound(remoteUpdateFile)
        If i > UBound(localUpdateFile) Then
        
            tmpArrIndex = UBound(tmpArr)
            ReDim Preserve tmpArr(tmpArrIndex + UBound(remoteUpdateFile) - UBound(localUpdateFile))
            
            j = i
            While j <= UBound(remoteUpdateFile)
                tmpArrIndex = tmpArrIndex + 1
                
                tmpArr(tmpArrIndex) = j
                j = j + 1
            Wend

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
    Next i
    
    compareUpdateFiles = tmpArr
End Function

''
' Downloads the Updates from the UpdateQueue.
'
' @param DownloadQueue Specifies reference to UpdateQueue
' @param remoteUpdateFile Specifies reference to Remote Update File
Public Sub DownloadUpdates(DownloadQueue() As Byte)
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 27/10/2008
'
'*************************************************
    Dim i As Long
    
'On Error GoTo error

    For i = 1 To UBound(DownloadQueue)
        If AoUpdateRemote(DownloadQueue(i)).HasPatches Then
            
        Else
            frmDownload.Show
            
            ColDownloadQueue.Add DownloadQueue(i)
            Call frmDownload.DownloadFile(DownloadQueue(i))
        End If
    Next i
Exit Sub

error:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Sub checkAoUpdateIntegrity()
    Dim nF As Integer
    
    'Look if exists the TEMP folder, if not, create it.
    If Dir(DownloadsPath, vbDirectory) = vbNullString Then
        MkDir DownloadsPath
    End If
    
    'Do we have a local AoUpdateFile ? If not, create it.
    If Dir(App.Path & "\" & AOUPDATE_FILE) = vbNullString Then
        nF = FreeFile
        Open App.Path & "\" & AOUPDATE_FILE For Output As #nF
            Print #nF, "# Este archivo contiene las direcciones de los archivos del cliente, con sus respectivas versiones y sus respectivos md5"
            Print #nF, "[INIT]"
            Print #nF, "NumFiles=1"
            Print #nF, "[File1] 'Argentum Client"
            Print #nF, "Name=Argentum.exe"
            Print #nF, "Version=0"
            Print #nF, "MD5="
        Close #nF
    End If
End Sub

Public Sub Main()
    DownloadsPath = App.Path & "\TEMP\"
    frmDownload.filePath = DownloadsPath
    
    Dim i As Long
    
    
    checkAoUpdateIntegrity
    'Debug.Print "Checking AoUpdate integrity..."
    
    'Download the remote AoUpdate.ini to the TEMP folder
    ColDownloadQueue.Add 0
    Call frmDownload.DownloadFile(0, UPDATES_SITE & AOUPDATE_FILE)
    
    While frmDownload.Downloading = True
        DoEvents
    Wend
    
    AoUpdateFileDownloaded = True
    
    AoUpdateLocal = ReadAoUFile(App.Path & "\" & "AoUpdate.ini") 'Load the local file
    AoUpdateRemote = ReadAoUFile(DownloadsPath & "AoUpdate.ini") 'Load the Remote file

    
    DownloadQueue = compareUpdateFiles(AoUpdateLocal, AoUpdateRemote) 'Compare local vs remote.
    
    If UBound(DownloadQueue) > 1 Then
        Call DownloadUpdates(DownloadQueue)
        'Patch 'em!
        'Check MD5 integrity. If wrong, redo queue, only do this once. For everyFile that went right, remake LocalAoUpdateFile
    Else
    End If
End Sub
