VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDownload 
   BorderStyle     =   0  'None
   Caption         =   "AoUpdate Downloader"
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pbDownload 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   923
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin InetCtlsObjects.Inet iDownload 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblDownloadPath 
      Height          =   495
      Left            =   2539
      TabIndex        =   2
      Top             =   263
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Descargando Archivo: "
      Height          =   255
      Left            =   739
      TabIndex        =   1
      Top             =   263
      Width           =   1695
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CurrentDownload As Byte
Public Downloading As Boolean
Private fileName As String
Public filePath As String

Public Sub DownloadFile(ByVal Index As Long, Optional ByRef sIniUrl As String)
    Dim sURL As String
    
    'Si se pasó el parámetro optional, es la ruta del INI
    If LenB(sIniUrl) > 0 Then
        sURL = sIniUrl
    Else
        sURL = UPDATES_SITE & AoUpdateRemote(Index).name
    End If
    
    If Not Downloading Then
        Downloading = True
        With iDownload
            .AccessType = icUseDefault
            'Indicamos el url del archivo
            .URL = sURL
            
            'Indicamos que vamos a descargar o recuperar un archivo desde una url
            .Execute , "GET"
        End With
        fileName = ReturnFileOrFolder(sURL, True, True)
        
        lblDownloadPath.Caption = fileName
    End If
End Sub

Private Sub iDownload_StateChanged(ByVal State As Integer)
    Dim nF As Integer
    Dim tmpArr() As Byte
    Dim fileSize As Long
    Dim dDone As Boolean
    Dim dData
    
  '  On Error GoTo error
    
    Select Case State
        Case icResponseCompleted
            dDone = False
            Downloading = True
            fileSize = iDownload.GetHeader("Content-Length")
            
            
            pbDownload.max = fileSize
            pbDownload.value = 0
            
            'Create the file.
            nF = FreeFile
            
            Open filePath & fileName For Binary As #nF
                While Not dDone
                    dData = iDownload.GetChunk("1024", icByteArray)
                    
                    If Len(dData) = 0 Then dDone = True
                    
                    tmpArr = dData
                    
                    Put #nF, , tmpArr
                    
                    pbDownload.value = pbDownload.value + (Len(dData) * 2)
                    DoEvents
                Wend
            Close #nF
            
            pbDownload.value = 0
            Downloading = False
            
            'Check the MD5 file
            CheckMD5File
    End Select
    
    Exit Sub
error:
    MsgBox Err.Description, vbCritical, Err.Number
    On Error Resume Next
    iDownload.Cancel
    pbDownload.value = 0
End Sub

Private Sub CheckMD5File()
    If AoUpdateFileDownloaded = False Then
        ColDownloadQueue.Remove (1)
        Unload Me
        Exit Sub
    End If
    
    If AoUpdateRemote(ColDownloadQueue.Item(1)).MD5 <> MD5File(DownloadsPath & AoUpdateRemote(ColDownloadQueue.Item(1)).name) Then
        If AoUpdateRemote(ColDownloadQueue.Item(1)).DownloadError = True Then
            'DOS ERRORES, QUÉ HAGO?
            'Cola de archivos de errores?
            Debug.Print "Dos errores en " & AoUpdateRemote(ColDownloadQueue.Item(1)).name
            'Elimino el erroneo y sigo con el siguiente
            ColDownloadQueue.Remove (1)
            NextDownload
        Else
            'Lo pongo como erroneo y vuelvo a descargar
            AoUpdateRemote(ColDownloadQueue.Item(1)).DownloadError = True
            Call DownloadFile(ColDownloadQueue.Item(1))
        End If
    Else
        'Es correcto, lo remuevo y descargo el siguiente
        ColDownloadQueue.Remove (1)
        'Start next Download
        NextDownload
    End If
End Sub
    
Private Sub NextDownload()
    'Si hay cola sigo descargando sino finalizé correctamente el Update
    If ColDownloadQueue.Count > 0 Then
        Call DownloadFile(ColDownloadQueue.Item(1))
    Else
        MsgBox "Acabé!"
        Unload Me
    End If
End Sub

Public Function ReturnFileOrFolder(FullPath As String, _
                                   ReturnFile As Boolean, _
                                   Optional IsURL As Boolean = False) _
                                   As String
'*************************************************
'Author: Jeff Cockayne
'Last modified: ?/?/?
'*************************************************

' ReturnFileOrFolder:   Returns the filename or path of an
'                       MS-DOS file or URL.
'
' Author:   Jeff Cockayne 4.30.99
'
' Inputs:   FullPath:   String; the full path
'           ReturnFile: Boolean; return filename or path?
'                       (True=filename, False=path)
'           IsURL:      Boolean; Pass True if path is a URL.
'
' Returns:  String:     the filename or path
'

Dim intDelimiterIndex As Integer

intDelimiterIndex = InStrRev(FullPath, IIf(IsURL, "/", "\"))
ReturnFileOrFolder = IIf(ReturnFile, _
                         Right(FullPath, Len(FullPath) - intDelimiterIndex), _
                         Left(FullPath, intDelimiterIndex))

End Function

