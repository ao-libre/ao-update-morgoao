VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDownload 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "AoUpdate Downloader"
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8955
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDownload.frx":22262
   ScaleHeight     =   5955
   ScaleWidth      =   8955
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerTimeOut 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   4800
      Top             =   3960
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox rtbDetalle 
      Height          =   2415
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4260
      _Version        =   393217
      BackColor       =   12632256
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDownload.frx":60868
   End
   Begin MSComctlLib.ProgressBar pbDownload 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   3360
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   0
   End
   Begin InetCtlsObjects.Inet iDownload 
      Left            =   120
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image imgCheck 
      Height          =   360
      Index           =   2
      Left            =   240
      Picture         =   "frmDownload.frx":608EB
      Top             =   5470
      Width           =   390
   End
   Begin VB.Image imgCheck 
      Height          =   360
      Index           =   1
      Left            =   960
      Picture         =   "frmDownload.frx":60C5D
      Top             =   0
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgCheck 
      Height          =   360
      Index           =   0
      Left            =   480
      Picture         =   "frmDownload.frx":61053
      Top             =   0
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image cmdComenzar 
      Height          =   645
      Left            =   1080
      Picture         =   "frmDownload.frx":613C5
      Top             =   4110
      Width           =   2700
   End
   Begin VB.Image imgExit 
      Height          =   645
      Left            =   1080
      Picture         =   "frmDownload.frx":66EBB
      Top             =   4750
      Width           =   2700
   End
   Begin VB.Label lblDownloadPath 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   3045
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descargando Archivo: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   3045
      Width           =   2055
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
Private FileName As String

Private downloadingConfig As Boolean
Private downloadingPatch As Boolean

Private WebTimeOut As Boolean

Public Sub DownloadConfigFile()
    downloadingConfig = True
    
    If Not WebTimeOut Then
        Call AddtoRichTextBox(frmDownload.rtbDetalle, "Descargando archivo de configuración.", 255, 255, 255, True, False, False)
        UPDATES_SITE = UPDATE_URL
    Else
        Call AddtoRichTextBox(frmDownload.rtbDetalle, "Descargando archivo de configuración desde página alternativa.", 255, 255, 255, True, False, False)
        UPDATES_SITE = UPDATE_URL_MIRROR
    End If
    
    Call DownloadFile(AOUPDATE_FILE)
End Sub

Public Sub DownloadPatch(ByVal file As String)
    downloadingPatch = True
    
    Call DownloadFile(file)
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
        
        FileName = ReturnFileOrFolder(sURL, True, True)
        
        lblDownloadPath.Caption = FileName
    End If
End Sub

'Private Sub chkJugar_Click()
'    NoExecute = chkJugar.value
'End Sub

Private Sub cmdComenzar_Click()
    Call ShellArgentum
    End
End Sub

Private Sub cmdExit_Click()
    If iDownload.StillExecuting Then Call iDownload.Cancel
    End
End Sub

Private Sub Form_Load()
    imgCheck(2).Picture = imgCheck(IIf(NoExecute, 0, 1)).Picture
    cmdComenzar.Enabled = False
End Sub

Private Sub iDownload_StateChanged(ByVal State As Integer)
    Dim nF As Integer
    Dim tmpArr() As Byte
    Dim fileSize As Long
    Dim downloaded As Long
    
On Error GoTo error
    nF = -1
    
    Select Case State
        Case icResponseCompleted
            fileSize = iDownload.GetHeader("Content-Length")
            downloaded = 0
            
            pbDownload.max = fileSize
            pbDownload.value = downloaded
            
            'Create the file.
            'Si existe el archivo que queremos bajar, lo borramos
            If FileExist(filePath & FileName, vbArchive) Then Kill filePath & FileName
            nF = FreeFile()
            
            Open filePath & FileName For Binary As nF
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
        Case icReceivingResponse
            If downloadingConfig Then TimerTimeOut.Enabled = True
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
    ElseIf downloadingPatch Then
        downloadingPatch = False
        
        Call PatchDownloaded
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

Private Sub imgCheck_Click(index As Integer)
    NoExecute = Not NoExecute
    imgCheck(index).Picture = imgCheck(IIf(NoExecute, 0, 1)).Picture
End Sub

Private Sub imgExit_Click()
    If iDownload.StillExecuting Then Call iDownload.Cancel
    End
End Sub

Private Sub TimerTimeOut_Timer()
If pbDownload.value = 0 And downloadingConfig = True Then
    If Not WebTimeOut Then
        If iDownload.StillExecuting Then Call iDownload.Cancel
        WebTimeOut = True
        Downloading = False
        
        Call DownloadConfigFile
    Else
        If MsgBox("No se ha podido acceder a la web y por lo tanto su cliente puede estar desactualizado" & vbCrLf & "¿Desea correr el cliente de todas formas?", vbYesNo) = vbYes Then
            Call ShellArgentum
        Else
            If iDownload.StillExecuting Then Call iDownload.Cancel
            End
        End If
    End If
End If

TimerTimeOut.Enabled = False
End Sub
