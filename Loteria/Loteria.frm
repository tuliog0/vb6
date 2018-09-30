VERSION 5.00
Begin VB.Form frmLoteria 
   Caption         =   "Loteria da sorte"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdVerificar 
      Caption         =   "Verificar"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdSortear 
      Caption         =   "Sortear"
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtApostado 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Numeros sorteados"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4455
      Begin VB.Label lblSorteado 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblSorteado 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblSorteado 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblSorteado 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblSorteado 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Numeros apostados"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      Begin VB.TextBox txtApostado 
         Height          =   285
         Index           =   4
         Left            =   3840
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtApostado 
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtApostado 
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtApostado 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label lblAcertos 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade de acertos:"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   1665
   End
End
Attribute VB_Name = "frmLoteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdLimpar_Click()
    
    Dim i As Byte
        For i = 0 To 4
            lblSorteado(i).Caption = ""
        Next i
        For i = 0 To 4
            txtApostado(i).Text = ""
        Next i
        lblAcertos.Caption = ""
        acertos = 0
        txtApostado(0).SetFocus

End Sub

Private Sub cmdSortear_Click()

    Dim i As Byte
    Randomize
    For i = 0 To 4
        lblSorteado(i).Caption = Int(Rnd * 10)
    Next i
    
End Sub

Private Sub cmdVerificar_Click()

    Dim acertos As Byte
    Dim i, j As Byte
    
    For i = 0 To 4
        For j = 0 To 4
            If txtApostado(i).Text = lblSorteado(j).Caption Then
                acertos = acertos + 1
            End If
        Next j
    Next i
    
    lblAcertos.Caption = acertos
    
    If acertos = 1 Then
        MsgBox "Parabéns, você fez um acerto", vbInformation, "Aviso"
    ElseIf acertos = 2 Then
        MsgBox "Parabéns, você fez dois acertos", vbInformation, "Aviso"
    ElseIf acertos = 3 Then
        MsgBox "Parabéns, você fez um TERNO!!!", vbInformation, "Aviso"
    ElseIf acertos = 4 Then
        MsgBox "Parabéns, você fez uma QUADRA!!!", vbInformation, "Aviso"
    ElseIf acertos = 5 Then
        MsgBox "Parabéns, você fez uma QUINA!!!", vbInformation, "Aviso"
    Else
        MsgBox "Que pena você não acertou nada.", vbInformation, "Aviso"
    End If
    
End Sub
