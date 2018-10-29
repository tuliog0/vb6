VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Music 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Toca CD"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl MMControl1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   1720
      _Version        =   393216
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      BackEnabled     =   -1  'True
      StepEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      EjectEnabled    =   -1  'True
   End
   Begin VB.Label lblNumero1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Número de músicas"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1395
   End
   Begin VB.Label lblNumero 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   150
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Número da música"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Music Center"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   2070
   End
End
Attribute VB_Name = "Music"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arquivo As String
Dim volta As Boolean

Private Sub Form_Load()
    volta = False
    arquivo = Dir("D:\*.*")
    MMControl1.FileName = "D:" & arquivo
    MMControl1.Command = "open"
    lblNumero1.Caption = MMControl1.Tracks
    MMControl1.Command = "play"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MMControl1.Command = "stop"
    MMControl1.Command = "close"
End Sub


Private Sub MMControl1_NextClick(Cancel As Integer)
    If lblNumero.Caption = lblNumero1.Caption Then
        MMControl1.Command = "stop"
        MMControl1.Command = "close"
        volta = False
        MMControl1.FileName = "D:" & arquivo
        MMControl1.Command = "open"
        MMControl1.Command = "play"
        lblNumero.Caption = 2
        Exit Sub
    End If
    lblNumero.Caption = Val(lblNumero.Caption) + 1
End Sub

Private Sub MMControl1_PrevClick(Cancel As Integer)
    If lblNumero.Caption = "1" Then
        Exit Sub
    End If
    If volta = True Then
        lblNumero.Caption = Val(lblNumero.Caption) - 1
        volta = False
    Else
        volta = True
    End If
End Sub
