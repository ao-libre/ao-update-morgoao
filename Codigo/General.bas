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

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Const UPDATES_SITE As String = "http://www.argentuuum.com.ar/aoupdate/"
Public Const AOUPDATE_FILE As String = "AoUpdate.ini"
Public Const PARAM_UPDATED As String = "/uptodate"


Public Type tAoUpdateFile
    name As String              'File name
    version As Integer          'The version of the file
    md5 As String * 32          'It's checksum
    Path As String              'Path in the client to the file from App.Path (the server path is the same, changing '\' with '/')
    HasPatches As Boolean       'Weather if patches are available for this file or not (if not the complete file has to be downloaded)
    Comment As String           'Any comments regarding this file.
End Type

Public Type tAoUpdatePatches
    name As String          'its location in the server
    md5 As String * 32      'It's Checksum
End Type

Public DownloadsPath As String

Public AoUpdatePatches() As tAoUpdatePatches

Public AoUpdateRemote() As tAoUpdateFile
'Public AoUpdateLocal() As tAoUpdateFile
Public DownloadQueue() As Long
Public DownloadQueueIndex As Long
Public PatchQueueIndex As Long
Public ClientParams As String

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
        tmpAoUFile(i - 1).md5 = Leer.GetValue("File" & i, "MD5")
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

Public Sub CompareUpdateFiles(ByRef remoteUpdateFile() As tAoUpdateFile)
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
        'If i > UBound(localUpdateFile) Then
        '
        '    ReDim Preserve DownloadQueue(tmpArrIndex + UBound(remoteUpdateFile) - UBound(localUpdateFile)) As Long
        '
        '    j = i
        '    While j <= UBound(remoteUpdateFile)
        '        tmpArrIndex = tmpArrIndex + 1
        '
        '        DownloadQueue(tmpArrIndex) = j
        '        j = j + 1
        '    Wend
        '
        '    Exit Sub
        'End If
        
        'If remoteUpdateFile(i).name <> localUpdateFile(i).name Then
        '    Call MsgBox("Error critico en los archivos ini. Por favor descargue el AoUpdater nuevamente.")
        'End If
        
        'If remoteUpdateFile(i).version <> localUpdateFile(i).version Then
        '    'Version Diffs, add to download queue.
        '    tmpArrIndex = tmpArrIndex + 1
        '    ReDim Preserve DownloadQueue(tmpArrIndex) As Long
        '    DownloadQueue(tmpArrIndex) = i
        'ElseIf remoteUpdateFile(i).md5 <> MD5File(App.Path & "\" & remoteUpdateFile(i).Path & "\" & remoteUpdateFile(i).name) Then
        '    'File checksum diffs (corrupted file?), add to download queue.
        '    tmpArrIndex = tmpArrIndex + 1
        '    ReDim Preserve DownloadQueue(tmpArrIndex) As Long
        '    DownloadQueue(tmpArrIndex) = i
        'End If
        
        If Not FileExist(App.Path & remoteUpdateFile(i).Path & "\" & remoteUpdateFile(i).name, vbNormal) Then
            tmpArrIndex = tmpArrIndex + 1
            ReDim Preserve DownloadQueue(tmpArrIndex) As Long
            DownloadQueue(tmpArrIndex) = i
        ElseIf remoteUpdateFile(i).md5 <> MD5File(App.Path & remoteUpdateFile(i).Path & "\" & remoteUpdateFile(i).name) Then
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
On Error GoTo noqueue
    
    If DownloadQueueIndex > UBound(DownloadQueue) Then

On Error GoTo error
        ' Override local config file with remote one
        'Call Kill(App.Path & "\" & AOUPDATE_FILE)
        'Name DownloadsPath & "\" & AOUPDATE_FILE As App.Path & "\" & AOUPDATE_FILE
        
        'Overwrite every file not already patched
        For DownloadQueueIndex = 0 To UBound(DownloadQueue)
            With AoUpdateRemote(DownloadQueue(DownloadQueueIndex))
                If Not .HasPatches Then
                    If .name <> App.EXEName & ".exe" Then
                        If Dir$(App.Path & "\" & .Path & "\" & .name) <> vbNullString Then
                            Call Kill(App.Path & "\" & .Path & "\" & .name)
                        End If
                        
                        If Not FileExist(App.Path & "\" & .Path, vbDirectory) Then MkDir (App.Path & "\" & .Path)
                        
                        Name DownloadsPath & .name As App.Path & "\" & .Path & "\" & .name
                    Else
                        'We are trying to patch AoUpdate.exe, so we need to give an extra argument to client
                        ClientParams = "/patchao '" & App.EXEName & ".exe'"
                    End If
                End If
            End With
        Next DownloadQueueIndex
        
        'Call MsgBox("TERMINAMOS!")
        ClientParams = PARAM_UPDATED & " " & ClientParams
        Call AddtoRichTextBox(frmDownload.rtbDetalle, "Cliente de Argentum Online actualizado correctamente.", 255, 255, 255, True, False, False)
        frmDownload.cmdComenzar.Enabled = True
        
        If frmDownload.chkJugar.value = 1 Then
            Call ShellArgentum
        End If
        'End
    Else
        With AoUpdateRemote(DownloadQueue(DownloadQueueIndex))
            If .HasPatches Then
                Dim localVersion As Long
                
                localVersion = -1
                
                If FileExist(App.Path & "\" & .Path & "\" & .name, vbArchive) Then 'Check if local version is too old to be patched.
                    localVersion = GetVersion(App.Path & "\" & .Path & "\" & .name)
                End If
                
                If ReadPatches(DownloadQueue(DownloadQueueIndex) + 1, localVersion, .version, DownloadsPath & AOUPDATE_FILE) Then
                    'Reset index and download patches!
                    PatchQueueIndex = 0
                    Call frmDownload.DownloadPatch(AoUpdatePatches(PatchQueueIndex).name)
                Else
                    'Our version is too old to be patched (it doesn't exist in the server). Overwrite it!
                    .HasPatches = False
                    
                    Call AddtoRichTextBox(frmDownload.rtbDetalle, "Descargando " & .name & " - " & .Comment, 255, 255, 255, True, False, False)
                    
                    Call frmDownload.DownloadFile(Replace(.Path, "\", "/") & "/" & .name)
                End If
            Else
                'Downlaod file. Map local paths to urls.
                
                Call AddtoRichTextBox(frmDownload.rtbDetalle, "Descargando " & .name & " - " & .Comment, 255, 255, 255, True, False, False)
                
                Call frmDownload.DownloadFile(Replace(.Path, "\", "/") & .name)
            End If
        End With
        
        'Move on to the next one
        DownloadQueueIndex = DownloadQueueIndex + 1
    End If
Exit Sub

noqueue: 'If we get here, it means that there isn't any update.
    
    Call AddtoRichTextBox(frmDownload.rtbDetalle, "Descargas finalizadas", 255, 255, 255, True, False, False)
    frmDownload.cmdComenzar.Enabled = True
    
    ClientParams = PARAM_UPDATED & " " & ClientParams
    
    If frmDownload.chkJugar.value = 1 Then
        Call ShellArgentum
        End
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
        Call Apply_Patch(App.Path & "\" & .Path & "\", DownloadsPath & "\", AoUpdatePatches(PatchQueueIndex).md5, frmDownload.pbDownload)
#Else
        Call Apply_Patch(App.Path & "\" & .Path & "\", DownloadsPath & "\", frmDownload.pbDownload)
#End If
        'Delete patch after patching!
        Kill DownloadsPath & "\" & Right(AoUpdatePatches(PatchQueueIndex).name, Len(AoUpdatePatches(PatchQueueIndex).name) - InStrRev(AoUpdatePatches(PatchQueueIndex).name, "/"))
        
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
    'If Dir$(App.Path & "\" & AOUPDATE_FILE) = vbNullString Then
    '    nF = FreeFile()
    '
    '    Open App.Path & "\" & AOUPDATE_FILE For Output As #nF
    '        Print #nF, "# Este archivo contiene las direcciones de los archivos del cliente, con sus respectivas versiones y sus respectivos md5" & vbCrLf
    '        Print #nF, "[INIT]"
    '        Print #nF, "NumFiles=1" & vbCrLf & vbCrLf
    '        Print #nF, "[File1] 'Cliente"
    '        Print #nF, "Name=Argentum.exe"
    '        Print #nF, "Version=0"
    '        Print #nF, "MD5=4a52d8025392734793235bdb4f3a54fa"
    '        Print #nF, "Path=\"
    '        Print #nF, "HasPatches=0"
    '        Print #nF, "Comment=Cliente de Argentum Online, sin Alpha Blending"
    '    Close #nF
    'End If
End Sub

Public Sub ConfgFileDownloaded()
    'AoUpdateLocal = ReadAoUFile(App.Path & "\" & AOUPDATE_FILE) 'Load the local file
    AoUpdateRemote = ReadAoUFile(DownloadsPath & AOUPDATE_FILE) 'Load the Remote file
    
    Call CompareUpdateFiles(AoUpdateRemote)  'Compare local vs remote.
    
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

Public Sub ShellArgentum()
On Error GoTo error
    Call ShellExecute(0, "OPEN", App.Path & "\Argentum.exe", ClientParams, App.Path, 0)   'We open Argentum.exe updated
    End
    Exit Sub
error:
    MsgBox "Error al ejecutar el juego", vbCritical
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
        AoUpdatePatches(i - beginingVersion).md5 = Leer.GetValue("PATCHES" & numFile & "-" & i, "md5")
    Next i
End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'******************************************
    With RichTextBox
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
        RichTextBox.Refresh
    End With
End Sub
