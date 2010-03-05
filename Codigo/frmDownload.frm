VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmDownload 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "AoUpdate Downloader"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDownload.frx":22262
   ScaleHeight     =   6360
   ScaleWidth      =   9360
   StartUpPosition =   1  'CenterOwner
   Begin MSWinsockLib.Winsock wskDownload 
      Left            =   240
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer TimerTimeOut 
      Interval        =   10000
      Left            =   240
      Top             =   3000
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox rtbDetalle 
      Height          =   2415
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4260
      _Version        =   393217
      BackColor       =   12632256
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDownload.frx":799EF
   End
   Begin MSComctlLib.ProgressBar pbDownload 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   3610
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label lblDescargado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2160
      TabIndex        =   7
      Top             =   4440
      Width           =   75
   End
   Begin VB.Image imgSalirClick 
      Height          =   465
      Left            =   3840
      Picture         =   "frmDownload.frx":79A72
      Top             =   0
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgJugarClick 
      Height          =   495
      Left            =   5040
      Picture         =   "frmDownload.frx":7D8A2
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgJugarRollover 
      Height          =   495
      Left            =   7440
      Picture         =   "frmDownload.frx":81B55
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgSalirRollover 
      Height          =   465
      Left            =   6360
      Picture         =   "frmDownload.frx":85E42
      Top             =   0
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3600
      TabIndex        =   6
      Top             =   4440
      Width           =   75
   End
   Begin VB.Label lblVelocidad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   4200
      Width           =   75
   End
   Begin VB.Image imgCheck 
      Height          =   360
      Left            =   420
      Top             =   5750
      Width           =   390
   End
   Begin VB.Image imgCheckBkp 
      Height          =   405
      Left            =   600
      Picture         =   "frmDownload.frx":89D13
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgExit 
      Height          =   405
      Left            =   3225
      Top             =   5310
      Width           =   1020
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
      Left            =   3000
      TabIndex        =   2
      Top             =   3230
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
      Left            =   960
      TabIndex        =   1
      Top             =   3230
      Width           =   2055
   End
   Begin VB.Image imgJugar 
      Height          =   405
      Left            =   3195
      Top             =   4830
      Width           =   1020
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long


Public WithEvents Download As CDownload
Attribute Download.VB_VarHelpID = -1

Public CurrentDownload As Byte
Public filePath As String

Private Downloading As Boolean
Private FileName As String

Private downloadingConfig As Boolean
Private downloadingPatch As Boolean

Private WebTimeOut As Boolean

Private Sub Download_Starting(ByVal FileSize As Long, ByVal Header As String)
If FileSize <> 0 Then
    pbDownload.max = FileSize
End If

pbDownload.value = 0
End Sub

Private Sub Download_DataArrival(ByVal bytesTotal As Long)
'TODO: Cambiar la interface y permitir lblBytes y lblRate para darle más información al usuario.
'lblBytes = Val(lblBytes) + bytesTotal
Static lastTime As Long

If Download.FileSize <> 0 Then
    pbDownload.value = pbDownload.value + bytesTotal
    
    If GetTickCount - lastTime > 500 Then
        lblDescargado.Caption = Round(Download.CurrentFileDownloadedBytes / 1048576, 2)
        lblTotal.Caption = Round(Download.FileSize / 1048576, 2)
        lblVelocidad.Caption = Round(Download.AverageDownloadSpeed / 1024, 2)
        lastTime = GetTickCount
    End If
End If
End Sub

Private Sub Download_Completed()
pbDownload.max = 100
pbDownload.value = 100

lblDescargado.Caption = Round(Download.CurrentFileDownloadedBytes / 1048576, 2)
lblTotal.Caption = Round(Download.CurrentFileDownloadedBytes / 1048576, 2)
Downloading = False

If downloadingConfig Then
    downloadingConfig = False
    Call ConfgFileDownloaded
    
ElseIf downloadingPatch Then
    downloadingPatch = False
    Call PatchDownloaded
Else
    With AoUpdateRemote(DownloadQueue(DownloadQueueIndex - 1))
        If Dir$(App.Path & "\" & .Path & "\" & .name) <> vbNullString Then
            Call Kill(App.Path & "\" & .Path & "\" & .name)
        End If
                
        If Not FileExist(App.Path & "\" & .Path, vbDirectory) Then MkDir (App.Path & "\" & .Path)
            
        Name DownloadsPath & .name As App.Path & "\" & .Path & "\" & .name
        
'        If .Critical Then
'            Call ShellExecute(0, "OPEN", App.Path & "\" & .Path & "\" & .name, Command, App.Path, SW_SHOWNORMAL)    'We open AoUpdate.exe updated
'            End
'        End If
    End With
    
    Call NextDownload
End If

End Sub

Private Sub Download_Error(ByVal Number As Integer, Description As String)
    'Manejar el error que hubo.
    'Si estabamos bajando el archivo de config y tiro error, tratamos de bajar del mirror
    'Connection is aborted due to timeout or other failure
    If Number = 10053 Then
        If downloadingConfig Then
            If Not WebTimeOut Then
                Download.Cancel
                WebTimeOut = True
                Downloading = False
                Call DownloadConfigFile
            Else
                If MsgBox("No se ha podido acceder a la web y por lo tanto su cliente puede estar desactualizado" & vbCrLf & "¿Desea correr el cliente de todas formas?", vbYesNo) = vbYes Then
                    Call ShellArgentum
                Else
                    Download.Cancel
                    End
                End If
            End If
        End If
    End If
End Sub


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
        
        FileName = ReturnFileOrFolder(sURL, True, True)
        If FileExist(filePath & FileName, vbArchive) Then Kill filePath & FileName
        
        If downloadingConfig Then
            Me.Download.Download sURL, filePath & FileName, True
        Else
            Me.Download.Download sURL, filePath & FileName, False
        End If
        
        lblDownloadPath.Caption = FileName
    End If
End Sub

Private Sub cmdComenzar_Click()

End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub Form_Load()
    Set Download = New CDownload
    Call Download.Init(Me.wskDownload)
    NoExecute = Not NoExecute
    Call imgCheck_Click
    ''''1imgJugar.Enabled = False
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgExit.Picture = Nothing
    imgJugar.Picture = Nothing
End Sub

Private Sub imgCheck_Click()
    NoExecute = Not NoExecute
    If NoExecute Then
        imgCheck.Picture = Nothing
    Else
        imgCheck.Picture = imgCheckBkp.Picture
    End If
End Sub

Private Sub imgExit_Click()
    Call Download.Cancel
    End
End Sub


Private Sub Label2_Click()

End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgExit.Picture = imgSalirClick.Picture
End Sub

Private Sub imgExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgExit.Picture = imgSalirRollover.Picture
    imgJugar.Picture = Nothing
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgExit.Picture = Nothing
End Sub

Private Sub imgJugar_Click()
    If StillDownloading Then
        Call AddtoRichTextBox(rtbDetalle, "¡No puedes ejecutar el juego mientras se está actualizando! Aguarda unos minutos por favor", , , , True)
        Exit Sub
    End If
    Call ShellArgentum
    End
End Sub

Private Sub imgJugar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgJugar.Picture = imgJugarClick.Picture
End Sub

Private Sub imgJugar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgJugar.Picture = imgJugarRollover.Picture
    imgExit.Picture = Nothing
End Sub

Private Sub imgJugar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgJugar.Picture = Nothing
End Sub

Private Sub TimerTimeOut_Timer()
If downloadingConfig = True Then
    If Not WebTimeOut Then
        Download.Cancel
        WebTimeOut = True
        Downloading = False
        
        Call DownloadConfigFile
    Else
        If MsgBox("No se ha podido acceder a la web y por lo tanto su cliente puede estar desactualizado" & vbCrLf & "¿Desea correr el cliente de todas formas?", vbYesNo) = vbYes Then
            Call ShellArgentum
        Else
            Download.Cancel
            End
        End If
    End If
End If

TimerTimeOut.Enabled = False
End Sub

