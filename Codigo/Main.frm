VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AoUpdate"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   2535
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MiMD5 As MD5

Private Sub Command1_Click()
    Dim AoUpdateLocal() As tAoUpdateFile
    Dim AoUpdateRemote() As tAoUpdateFile
    Dim holaprobando() As tAoUpdateFile
    Dim i As Integer
    Dim b() As Byte
    
    ReDim holaprobando(0)
    AoUpdateLocal = ReadAoUFile(App.Path & "\" & "AoUpdate.ini")
    AoUpdateRemote = ReadAoUFile(App.Path & "\" & "AAoUpdate.ini")

    b = compareUpdateFiles(AoUpdateLocal, AoUpdateRemote)
    
    For i = 1 To UBound(b)
        Print Int(b(i))
    Next
End Sub

Private Sub Form_Load()
    Set MiMD5 = New MD5
End Sub


