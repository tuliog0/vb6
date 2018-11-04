VERSION 5.00
Begin VB.Form frmAbertura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jogo da Velha"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSair 
      Caption         =   "Sair"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton btnEntrar 
      Caption         =   "Entrar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtSegundo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1560
      TabIndex        =   4
      Top             =   1440
      Width           =   4125
   End
   Begin VB.TextBox txtPrimeiro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   4125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Jogador 2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Jogador 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Entre com os nomes dos jogadores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   4995
   End
End
Attribute VB_Name = "frmAbertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEntrar_Click()
    Primeiro = txtPrimeiro.Text
    Segundo = txtSegundo.Text
    Unload Me
    frmJogo.Show
End Sub

Private Sub btnSair_Click()
    End
End Sub

Private Sub txtPrimeiro_Change()
    If txtPrimeiro <> "" And txtSegundo <> "" Then
        btnEntrar.Enabled = True
    Else
        btnEntrar.enable = False
    End If
End Sub

Private Sub txtSegundo_Change()
    If txtSegundo.txt <> "" And txtPrimeiro <> "" Then
        btnEntrar.Enabled = True
    Else
        btnEntrar.Enabled = False
    End If
End Sub
