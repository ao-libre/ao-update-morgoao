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

Public Type tAoUpdateFile
    name As String              'File name
    Version As Integer          'The version of the file
    MD5 As String * 32          'It's checksum
    Path As String              'Path in the client to the file from App.Path (the server path is the same, changing '\' with '/')
    HasPatches As Boolean       'Weather if patches are available for this file or not (if not the complete file has to be downloaded)
    Comment As String           'Any comments regarding this file.
End Type

Public DownloadsPath As String

Public AoUpdateRemote() As tAoUpdateFile
Public AoUpdateLocal() As tAoUpdateFile
Public DownloadQueue() As Long
Public DownloadQueueIndex As Long

''
' Loads the AoUpdate Ini File to an struct array
'
' @param file Specifies reference to AoUpdateIniFile
' @return an array of tAoUpdate

Public Function ReadAoUFile(ByVal file As String) As tAoUpdateFile()
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
    
    ReDim tmpAoUFile(NumFiles - 1) As tAoUpdateFile
    
    For i = 1 To NumFiles
        tmpAoUFile(i - 1).name = Leer.GetValue("File" & i, "Name")
        tmpAoUFile(i - 1).Version = CInt(Leer.GetValue("File" & i, "Version"))
        tmpAoUFile(i - 1).MD5 = Leer.GetValue("File" & i, "MD5")
        tmpAoUFile(i - 1).Path = Leer.GetValue("File" & i, "Path")
        
        If Leer.KeyExists("HasPatches") Then tmpAoUFile(i - 1).HasPatches = CBool(Leer.GetValue("File" & i, "HasPatches"))
        If Leer.KeyExists("Comment") Then tmpAoUFile(i - 1).Comment = Leer.GetValue("File" & i, "Comment")
    Next i
    
    ReadAoUFile = tmpAoUFile
    
    Set Leer = Nothing
Exit Function

error:
    Call MsgBox(Err.Description, vbCritical, Err.Number)
    Set Leer = Nothing
End Function

''
' Compares the local AoUpdate file with the one in the server
'
' @param localUpdateFile Specifies reference to Local Update File
' @param remoteUpdateFile Specifies reference to Remote Update File

Public Sub CompareUpdateFiles(ByRef localUpdateFile() As tAoUpdateFile, ByRef remoteUpdateFile() As tAoUpdateFile)
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 27/10/2008
'
'*************************************************
    Dim i As Long
    Dim j As Long
    Dim tmpArrIndex As Long
    
    'ReDim DownloadQueue(0) As Long
    tmpArrIndex = -1
    
    For i = 0 To UBound(remoteUpdateFile)
        If i > UBound(localUpdateFile) Then
            
            ReDim Preserve DownloadQueue(tmpArrIndex + UBound(remoteUpdateFile) - UBound(localUpdateFile)) As Long
            
            j = i
            While j <= UBound(remoteUpdateFile)
                tmpArrIndex = tmpArrIndex + 1
                
                DownloadQueue(tmpArrIndex) = j
                j = j + 1
            Wend
            
            Exit Sub
        End If
        
        If remoteUpdateFile(i).name <> localUpdateFile(i).name Then
            Call MsgBox("Error critico en los archivos ini. Por favor descargue el AoUpdater nuevamente.")
        End If
        
        If remoteUpdateFile(i).Version <> localUpdateFile(i).Version Then
            'Version Diffs, add to download queue.
            tmpArrIndex = tmpArrIndex + 1
            ReDim Preserve DownloadQueue(tmpArrIndex) As Long
            DownloadQueue(tmpArrIndex) = i
        ElseIf remoteUpdateFile(i).MD5 <> MD5File(App.Path & "\" & remoteUpdateFile(i).Path & remoteUpdateFile(i).name) Then
            'File checksum diffs (corrupted file?), add to download queue.
            tmpArrIndex = tmpArrIndex + 1
            ReDim Preserve DownloadQueue(tmpArrIndex) As Long
            DownloadQueue(tmpArrIndex) = i
        End If
    Next i
End Sub

''
' Downloads the Updates from the UpdateQueue.
'
' @param DownloadQueue Specifies reference to UpdateQueue
' @param remoteUpdateFile Specifies reference to Remote Update File

Public Sub NextDownload()
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 27/10/2008
'
'*************************************************
'On Error GoTo error
    
    If DownloadQueueIndex > UBound(DownloadQueue) Then
' TODO : TERMINAMOS!!
        ' Override local config file with remote one
        Call Kill(App.Path & "\" & AOUPDATE_FILE)
        Name DownloadsPath & "\" & AOUPDATE_FILE As App.Path & "\" & AOUPDATE_FILE
        
        'Overwrite / patch every file
        For DownloadQueueIndex = 0 To UBound(DownloadQueue)
            With AoUpdateRemote(DownloadQueue(DownloadQueueIndex))
                
                If .HasPatches Then
' TODO : Patch files!
                Else
                    If Dir$(App.Path & "\" & .name) <> vbNullString Then
                        Call Kill(App.Path & "\" & .name)
                    End If
                    
                    Name DownloadsPath & "\" & .name As App.Path & "\" & .Path & "\" & .name
                End If
            End With
        Next DownloadQueueIndex
        
        Call MsgBox("TERMINAMOS!")
        End
    Else
        If AoUpdateRemote(DownloadQueue(DownloadQueueIndex)).HasPatches Then
'TODO : Download and apply patches individually
        Else
            'Downlaod file. Map local paths to urls.
            Call frmDownload.DownloadFile(Replace("\", AoUpdateRemote(DownloadQueue(DownloadQueueIndex)).Path, "/") & AoUpdateRemote(DownloadQueue(DownloadQueueIndex)).name)
        End If
        
        'Move on to the next one
        DownloadQueueIndex = DownloadQueueIndex + 1
    End If
Exit Sub

error:
    Call MsgBox(Err.Description, vbCritical, Err.Number)
End Sub

Private Sub CheckAoUpdateIntegrity()
    Dim nF As Integer
    
    'Look if exists the TEMP folder, if not, create it.
    If Dir$(DownloadsPath, vbDirectory) = vbNullString Then
        Call MkDir(DownloadsPath)
    End If
    
    'Do we have a local AoUpdateFile ? If not, create it.
    If Dir$(App.Path & "\" & AOUPDATE_FILE) = vbNullString Then
        nF = FreeFile()
        
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

Public Sub ConfgFileDownloaded()
    AoUpdateLocal = ReadAoUFile(App.Path & "\" & AOUPDATE_FILE) 'Load the local file
    AoUpdateRemote = ReadAoUFile(DownloadsPath & AOUPDATE_FILE) 'Load the Remote file
    
    Call CompareUpdateFiles(AoUpdateLocal, AoUpdateRemote) 'Compare local vs remote.
    
    Call NextDownload
End Sub

Public Sub Main()
    Dim i As Long
    
    DownloadsPath = App.Path & "\TEMP\"
    frmDownload.filePath = DownloadsPath
    
    'Display form
    Call frmDownload.Show
    
    Call CheckAoUpdateIntegrity
    
    'Download the remote AoUpdate.ini to the TEMP folder and let the magic begin
    Call frmDownload.DownloadConfigFile
End Sub
