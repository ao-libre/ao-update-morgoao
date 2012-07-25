VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmDownload 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "AoUpdate Downloader"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   Icon            =   "frmDownload.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDownload.frx":22262
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   624
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
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmDownload.frx":72AD3
   End
   Begin MSComctlLib.ProgressBar pbDownload 
      Height          =   225
      Left            =   975
      TabIndex        =   0
      Top             =   3600
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblTotalArchivos 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2880
      TabIndex        =   8
      Top             =   4545
      Width           =   675
   End
   Begin VB.Label lblArchivo 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   7
      Top             =   4545
      Width           =   675
   End
   Begin VB.Label lblDescargado 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   4245
      Width           =   600
   End
   Begin VB.Image imgSalirClick 
      Height          =   465
      Left            =   3840
      Picture         =   "frmDownload.frx":72B55
      Top             =   0
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgJugarClick 
      Height          =   495
      Left            =   5040
      Picture         =   "frmDownload.frx":76985
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgJugarRollover 
      Height          =   495
      Left            =   7440
      Picture         =   "frmDownload.frx":7AC38
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgSalirRollover 
      Height          =   465
      Left            =   6360
      Picture         =   "frmDownload.frx":7EF25
      Top             =   0
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3495
      TabIndex        =   5
      Top             =   4245
      Width           =   555
   End
   Begin VB.Label lblVelocidad 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1965
      TabIndex        =   4
      Top             =   3975
      Width           =   795
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
      Picture         =   "frmDownload.frx":82DF6
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
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3000
      TabIndex        =   2
      Top             =   3195
      Width           =   75
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descargando Archivo: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1000
      TabIndex        =   1
      Top             =   3195
      Width           =   1935
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
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

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
        'Check if the MD5 matches.
        If UCase$(.MD5) <> UCase$(MD5.MD5File(DownloadsPath & .name)) Then
            Kill DownloadsPath & .name
            'If we are not downloading the file from the mirror, lets redownload it from there
            If (Not DownloadingFromMirror) And (General.UPDATES_SITE = General.UPDATE_URL) And (UPDATE_URL_MIRROR <> vbNullString) Then
                UPDATES_SITE = UPDATE_URL_MIRROR
                DownloadQueueIndex = DownloadQueueIndex - 1
                Call NextDownload
                Exit Sub
            Else
                Call AddtoRichTextBox(rtbDetalle, "No se pudo verificar la integridad del archivo " & .name & ", puede que el juego no funcione correctamente. Comuniquese con los administradores", 250, 20, 70, True, False, False)
                
                If Not DownloadingFromMirror Then UPDATES_SITE = UPDATE_URL
                
                Call NextDownload
                Exit Sub
            End If
        End If
        
        If FileExist(App.Path & "\" & .Path & "\" & .name, vbNormal) Then
            Call Kill(App.Path & "\" & .Path & "\" & .name)
        End If
                
        If Not FileExist(App.Path & "\" & .Path, vbDirectory) Then Call MkDir(App.Path & "\" & .Path)
            
        Name DownloadsPath & .name As App.Path & "\" & .Path & "\" & .name
        
        If .Critical Then
            Call ShellExecute(0, "OPEN", App.Path & "\" & .Path & "\" & .name, Command, App.Path, SW_SHOWNORMAL)    'We open AoUpdate.exe updated
            End
        End If
    End With
    
    If Not DownloadingFromMirror Then UPDATES_SITE = UPDATE_URL
    Call NextDownload
End If

End Sub

Private Sub Download_Error(ByVal Number As Integer, Description As String)
    'Manejar el error que hubo.
    'Si estabamos bajando el archivo de config y tiro error, tratamos de bajar del mirror
    '404 NOT FOUND
    'Connection is aborted due to timeout or other failure
    If Number = 10053 Or Number = 404 Then
        If downloadingConfig Then
            If Not WebTimeOut Then
                Download.Cancel
                WebTimeOut = True
                Downloading = False
                Call DownloadConfigFile
            Else
                If MsgBox("No se ha podido acceder a la web y por lo tanto su cliente puede estar desactualizado" & vbCrLf & "¿Desea correrlo igualmente?", vbYesNo) = vbYes Then
                    ClientParams = PARAM_UPDATED & " " & ClientParams
                    Call ShellArgentum
                Else
                    Download.Cancel
                    End
                End If
            End If
        Else
            If (Not DownloadingFromMirror) And (UPDATE_URL = UPDATES_SITE) And (UPDATE_URL_MIRROR <> vbNullString) Then
                UPDATES_SITE = UPDATE_URL_MIRROR
                If downloadingPatch Then
                    'Try to redownload it from the mirror
                    PatchQueueIndex = PatchQueueIndex - 1
                    downloadingPatch = False
                    Call PatchDownloaded
                Else
                    DownloadQueueIndex = DownloadQueueIndex - 1
                    Call NextDownload 'Try to redownload the file
                End If
            Else
                If Not DownloadingFromMirror Then UPDATES_SITE = UPDATE_URL
                
                If downloadingPatch Then
                    Call AddtoRichTextBox(rtbDetalle, "Fallo la descarga de Parches, bajando el archivo completo...", 250, 20, 70, True, False, False)
                    With AoUpdateRemote(DownloadQueue(DownloadQueueIndex))
                        .HasPatches = False
                        Call NextDownload
                    End With
                Else
                    With AoUpdateRemote(DownloadQueue(DownloadQueueIndex))
                        Call AddtoRichTextBox(rtbDetalle, "No se ha podido descargar correctamente el archivo " & .name & ", puede que el juego no funcione correctamente. Comuniquese con los administradores", 250, 20, 70, True, False, False)
                        Call NextDownload
                    End With
                    Exit Sub
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
        DownloadingFromMirror = True
    End If
    
    TimerTimeOut.Enabled = False
    TimerTimeOut.Enabled = True
    Call DownloadFile(AOUPDATE_FILE)
End Sub

Public Sub DownloadPatch(ByVal file As String)
    downloadingPatch = True
    
    Call DownloadFile(file)
End Sub

Public Sub DownloadFile(ByVal file As String)
    Dim sURL As String
    Dim antiProxy As String
    
    'Parche metido para evitar que se arme quilombo con el proxy fruta que puso speedy.
    antiProxy = "?speedy=" & (Int(Rnd * 30000)) & "&meteteel=proxy" & CLng(Timer) & "&en=el" & (CLng(Timer) Mod (Int(Rnd * 500) + 1)) & "&orto=" & (Int(Rnd * 35000) + 20)
    sURL = UPDATES_SITE & file

    If Not Downloading Then
        Downloading = True
        
        FileName = ReturnFileOrFolder(sURL, True, True)
        If FileExist(filePath & FileName, vbArchive) Then Kill filePath & FileName
        
        If downloadingConfig Then
            Call Me.Download.Download(sURL & antiProxy, filePath & FileName, True)
        Else
            Call Me.Download.Download(sURL & antiProxy, filePath & FileName, False)
            lblArchivo.Caption = DownloadQueueIndex + 1
        End If
        
        lblDownloadPath.Caption = FileName
        
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call Download.Cancel
        End
    End If
End Sub

Private Sub Form_Load()
    Set Download = New CDownload
    Call Download.Init(Me.wskDownload)
    NoExecute = Not NoExecute
    Call imgCheck_Click
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgExit.Picture = imgSalirClick.Picture
End Sub

Private Sub imgExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgExit.Picture = imgSalirRollover.Picture
    imgJugar.Picture = Nothing
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub imgJugar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgJugar.Picture = imgJugarClick.Picture
End Sub

Private Sub imgJugar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgJugar.Picture = imgJugarRollover.Picture
    imgExit.Picture = Nothing
End Sub

Private Sub imgJugar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgJugar.Picture = Nothing
End Sub

Private Sub TimerTimeOut_Timer()
If downloadingConfig Then
    If Not WebTimeOut Then
        Download.Cancel
        WebTimeOut = True
        Downloading = False
        
        Call DownloadConfigFile
    Else
        If MsgBox("No se ha podido acceder a la web y por lo tanto su cliente puede estar desactualizado" & vbCrLf & "¿Desea correrlo igualmente?", vbYesNo) = vbYes Then
            ClientParams = PARAM_UPDATED & " " & ClientParams
            Call ShellArgentum
        Else
            Download.Cancel
            End
        End If
    End If
End If
End Sub
