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
    version As Integer          'The version of the file
    MD5 As String * 32          'It's checksum
    Path As String              'Path in the client to the file from App.Path (the server path is the same, changing '\' with '/')
    HasPatches As Boolean       'Weather if patches are available for this file or not (if not the complete file has to be downloaded)
    Comment As String           'Any comments regarding this file.
End Type

Public Type tAoUpdatePatches
    name As String          'its location in the server
    MD5 As String * 32      'It's Checksum
End Type

Public DownloadsPath As String

Public AoUpdatePatches() As tAoUpdatePatches

Public AoUpdateRemote() As tAoUpdateFile
Public AoUpdateLocal() As tAoUpdateFile
Public DownloadQueue() As Long
Public DownloadQueueIndex As Long
Public PatchQueueIndex As Long

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
    
On Error GoTo error
    
    Call Leer.Initialize(file)
    
    NumFiles = Leer.GetValue("INIT", "NumFiles")
    
    ReDim tmpAoUFile(NumFiles - 1) As tAoUpdateFile
    
    For i = 1 To NumFiles
        tmpAoUFile(i - 1).name = Leer.GetValue("File" & i, "Name")
        tmpAoUFile(i - 1).version = CInt(Leer.GetValue("File" & i, "Version"))
        tmpAoUFile(i - 1).MD5 = Leer.GetValue("File" & i, "MD5")
        tmpAoUFile(i - 1).Path = Leer.GetValue("File" & i, "Path")
        tmpAoUFile(i - 1).HasPatches = CBool(Leer.GetValue("File" & i, "HasPatches"))
        tmpAoUFile(i - 1).Comment = Leer.GetValue("File" & i, "Comment")
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
    
' TODO : Check what happens if no files are to be downloaded....
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
        
        If remoteUpdateFile(i).version <> localUpdateFile(i).version Then
            'Version Diffs, add to download queue.
            tmpArrIndex = tmpArrIndex + 1
            ReDim Preserve DownloadQueue(tmpArrIndex) As Long
            DownloadQueue(tmpArrIndex) = i
        ElseIf remoteUpdateFile(i).MD5 <> MD5File(App.Path & "\" & remoteUpdateFile(i).Path & "\" & remoteUpdateFile(i).name) Then
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
On Error GoTo error
    
    If DownloadQueueIndex > UBound(DownloadQueue) Then
        
        ' Override local config file with remote one
        Call Kill(App.Path & "\" & AOUPDATE_FILE)
        Name DownloadsPath & "\" & AOUPDATE_FILE As App.Path & "\" & AOUPDATE_FILE
        
        'Overwrite every file not already patched
        For DownloadQueueIndex = 0 To UBound(DownloadQueue)
            With AoUpdateRemote(DownloadQueue(DownloadQueueIndex))
                
                If Not .HasPatches Then
                    If Dir$(App.Path & "\" & .Path & "\" & .name) <> vbNullString Then
                        Call Kill(App.Path & "\" & .Path & "\" & .name)
                    End If
                    
                    If Not FileExist(App.Path & "\" & .Path, vbDirectory) Then MkDir (App.Path & "\" & .Path)
                    
                    Name DownloadsPath & .name As App.Path & "\" & .Path & "\" & .name
                End If
            End With
        Next DownloadQueueIndex
        
        Call MsgBox("TERMINAMOS!")
        End
    Else
        With AoUpdateRemote(DownloadQueue(DownloadQueueIndex))
            If .HasPatches Then
                Dim localVersion As Long
                
                localVersion = -1
                
                If FileExist(App.Path & "\" & .Path & "\" & .name, vbArchive) Then 'Check if local version is too old to be patched.
                    localVersion = GetVersion(App.Path & "\" & .Path & "\" & .name)
                End If
                
                If ReadPatches(DownloadQueue(DownloadQueueIndex) + 1, localVersion, .version, App.Path & "\" & AOUPDATE_FILE) Then
                    'Reset index and download patches!
                    PatchQueueIndex = 0
                    Call frmDownload.DownloadPatch(AoUpdatePatches(PatchQueueIndex).name)
                Else
                    'Our version is too old to be patched (it doesn't exist in the server). Overwrite it!
                    .HasPatches = False
                    Call frmDownload.DownloadFile(Replace(.Path, "\", "/") & "/" & .name)
                End If
            Else
                'Downlaod file. Map local paths to urls.
                Call frmDownload.DownloadFile(Replace(.Path, "\", "/") & .name)
            End If
        End With
        
        'Move on to the next one
        DownloadQueueIndex = DownloadQueueIndex + 1
    End If
Exit Sub

error:
    Call MsgBox(Err.Description, vbCritical, Err.Number)
End Sub

Public Sub PatchDownloaded()
    Dim localVersion As Long
    
    localVersion = -1
    
    With AoUpdateRemote(DownloadQueue(DownloadQueueIndex - 1))
        'Apply downlaoded patch!
#If SeguridadAlkon Then
        Call Apply_Patch(App.Path & "\" & .Path & "\", DownloadsPath & "\", AoUpdatePatches(PatchQueueIndex).MD5, frmDownload.pbDownload)
#Else
        Call Apply_Patch(App.Path & "\" & .Path & "\", DownloadsPath & "\", frmDownload.pbDownload)
#End If
        
        localVersion = GetVersion(App.Path & "\" & .Path & "\" & .name)
        
        If .version = localVersion Then
            'We finished patching this file, continue!
            Call NextDownload
        Else
            PatchQueueIndex = PatchQueueIndex + 1
            Call frmDownload.DownloadPatch(AoUpdatePatches(PatchQueueIndex).name)
        End If
    End With
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
    
    'Start downloads!
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

Public Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

''
' Loads the patches and their md5. Check if a file isn't too old to be patched
'
' @param NumFile Specifies reference to File in AoUpdateFile file.
' @param begininVersion Specifies reference to LocalVersion
' @param endingVersion Specifies reference to last version of the file
' @param sFile specifies reference to ConfiFile to read data from.
'
' @returns True if the file can be patcheable or false if the file can't be patcheable

Private Function ReadPatches(ByVal numFile As Integer, ByVal beginingVersion As Long, ByVal endingVersion As Long, ByVal sFile As String) As Boolean
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 27/10/2008
'
'*************************************************
    Dim nF As Integer
    Dim i As Long
    Dim Leer As New clsIniReader
    
    nF = FreeFile
    
    Call Leer.Initialize(sFile)
    
    If Not Leer.KeyExists("PATCHES" & numFile & "-" & beginingVersion) Or beginingVersion = -1 Then Exit Function
    ReadPatches = True
    
    ReDim AoUpdatePatches(endingVersion - beginingVersion - 1) As tAoUpdatePatches
    
    For i = beginingVersion To endingVersion - 1
        AoUpdatePatches(i - beginingVersion).name = Leer.GetValue("PATCHES" & numFile & "-" & i, "name")
        AoUpdatePatches(i - beginingVersion).MD5 = Leer.GetValue("PATCHES" & numFile & "-" & i, "md5")
    Next i
End Function
