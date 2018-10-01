VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jogo da Memória"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   3360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   3360
   End
   Begin VB.Label lblTempo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Você tem 60s para vencer"
      Height          =   195
      Left            =   1800
      TabIndex        =   0
      Top             =   3000
      Width           =   1875
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   15
      Left            =   4440
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   14
      Left            =   3000
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   13
      Left            =   1560
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   12
      Left            =   120
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   11
      Left            =   4440
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   10
      Left            =   3000
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   9
      Left            =   1560
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   8
      Left            =   120
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   7
      Left            =   4440
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   6
      Left            =   3000
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   5
      Left            =   1560
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   4
      Left            =   120
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   3
      Left            =   4440
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   2
      Left            =   120
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   0
      Left            =   3000
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   1
      Left            =   1560
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Byte
Dim Primeiro As Boolean
Dim Clicou, Ganhou As Byte
Const NumPecas = 16

Private Sub Form_Load()

    iniciaJogo
    
End Sub

Private Sub imgMemo_Click(Index As Integer)

    If imagMemo(Index).Tag = "Word" Then
        imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\word.jpg")
    ElseIf imagMemo(Index).Tag = "Tabela" Then
        imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\table.jpg")
    ElseIf imagMemo(Index).Tag = "Cd" Then
        imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\cdrom01.jpg")
    ElseIf imagMemo(Index).Tag = "Copa" Then
        imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\copas.jpg")
    ElseIf imagMemo(Index).Tag = "Paus" Then
        imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\paus.jpg")
    ElseIf imagMemo(Index).Tag = "Ouro" Then
        imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\ouro.jpg")
    ElseIf imagMemo(Index).Tag = "Preto" Then
        imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\preto.jpg")
    ElseIf imagMemo(Index).Tag = "Abrir" Then
        imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\folder.jpg")
    End If
    
    imgMemo(Index).Enabled = False
    
    If Primeiro = True Then
        Clicou = imgMemo(Index).Index
    End If
    
    If imgMemo(Clicou).Tag <> imgMemo(Index).Tag And Primeiro = True Then
        imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\delphi.jpg")
        imgMemo(Clicou).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\delphi.jpg")
        imgMemo(Index).Enabled = True
        imgMemo(Clicou).Enabled = True
        Primeiro = True
    ElseIf imgMemo(Clicou).Tag = imgMemo(Index).Tag And Primeiro = False Then
        Ganhou = Ganhou + 1
        Primeiro = True
    Else
        Primeiro = False
    End If
    
    If Ganhou = 8 Then
        MsgBox "Você venceu"
        Timer1.enable = False
        iniciaJogo
    End If
    
End Sub
