VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
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
Public filePath As String

Private Downloading As Boolean
Private fileName As String

Private downloadingConfig As Boolean

Public Sub DownloadConfigFile()
    downloadingConfig = True
    
    Call DownloadFile(AOUPDATE_FILE)
End Sub

Public Sub DownloadFile(ByVal file As String)
    Dim sURL As String
    
    sURL = UPDATES_SITE & file
    
    If Not Downloading Then
        Downloading = True
        
        With iDownload
            .AccessType = icUseDefault
            
            'Indicamos que vamos a descargar o recuperar un archivo desde una url
            Call .Execute(sURL, "GET")
        End With
        
        fileName = ReturnFileOrFolder(sURL, True, True)
        
        lblDownloadPath.Caption = fileName
    End If
End Sub

Private Sub iDownload_StateChanged(ByVal State As Integer)
    Dim nF As Integer
    Dim tmpArr() As Byte
    Dim fileSize As Long
    Dim downloaded As Long
    
'On Error GoTo error
    nF = -1
    
    Select Case State
        Case icResponseCompleted
            fileSize = iDownload.GetHeader("Content-Length")
            downloaded = 0
            
            pbDownload.max = fileSize
            pbDownload.value = downloaded
            
            'Create the file.
            nF = FreeFile()
            
            Open filePath & fileName For Binary As nF
                While fileSize <> downloaded
                    tmpArr = iDownload.GetChunk(1024, icByteArray)
                    
                    Put nF, , tmpArr
                    
                    downloaded = downloaded + UBound(tmpArr) + 1
                    pbDownload.value = downloaded
                    
                    DoEvents
                Wend
            Close nF
            
            'Reset nF
            nF = -1
            
            Call DownloadComplete
    End Select
Exit Sub

error:
    Call MsgBox(Err.Description, vbCritical, Err.Number)
    
On Error Resume Next
    If nF <> -1 Then
        Close nF
    End If
    
    iDownload.Cancel
    pbDownload.value = 0
End Sub

Private Sub DownloadComplete()
    Downloading = False
    
    If downloadingConfig Then
        downloadingConfig = False
        
        Call ConfgFileDownloaded
    Else
        Call NextDownload
    End If
End Sub

Public Function ReturnFileOrFolder(ByVal FullPath As String, _
                                   ByVal ReturnFile As Boolean, _
                                   Optional ByVal IsURL As Boolean = False) _
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
                             Right$(FullPath, Len(FullPath) - intDelimiterIndex), _
                             Left$(FullPath, intDelimiterIndex))
End Function
