VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jogo da Memória"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Acertos:"
      Height          =   195
      Left            =   5280
      TabIndex        =   4
      Top             =   3000
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Segundos:"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label lblAcertos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5520
      TabIndex        =   2
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblTempo 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   0
      TabIndex        =   1
      Top             =   3240
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Você tem 60s para vencer"
      Height          =   195
      Left            =   2040
      TabIndex        =   0
      Top             =   3360
      Width           =   1875
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   15
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   14
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   13
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   12
      Left            =   120
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   11
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   10
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   9
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   8
      Left            =   120
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   7
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   6
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   5
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   4
      Left            =   120
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   3
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "JogoDaMemoria.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   2
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgMemo 
      Height          =   495
      Index           =   1
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BorderColor     =   &H80000007&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   0
      Top             =   0
      Width           =   5895
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

    If imgMemo(Index).Tag = "Word" Then
        'imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\word.jpg")
        imgMemo(Index).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\word.jpg")
    ElseIf imgMemo(Index).Tag = "Tabela" Then
        'imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\table.jpg")
        imgMemo(Index).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\table.jpg")
    ElseIf imgMemo(Index).Tag = "Cd" Then
        'imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\cdrom01.jpg")
        imgMemo(Index).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\cdrom01.jpg")
    ElseIf imgMemo(Index).Tag = "Copa" Then
        'imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\copas.jpg")
        imgMemo(Index).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\copas.jpg")
    ElseIf imgMemo(Index).Tag = "Paus" Then
        'imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\paus.jpg")
        imgMemo(Index).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\paus.jpg")
    ElseIf imgMemo(Index).Tag = "Ouro" Then
        'imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\ouro.jpg")
        imgMemo(Index).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\ouro.jpg")
    ElseIf imgMemo(Index).Tag = "Preto" Then
        'imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\preto.jpg")
        imgMemo(Index).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\preto.jpg")
    ElseIf imgMemo(Index).Tag = "Abrir" Then
        'imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\folder.jpg")
        imgMemo(Index).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\folder.jpg")
    End If
    
    imgMemo(Index).Enabled = False
    
    If Primeiro = True Then
        Clicou = imgMemo(Index).Index
    End If
    
    If imgMemo(Clicou).Tag <> imgMemo(Index).Tag And Primeiro = False Then
        MsgBox " Tente novamente!"
        'imgMemo(Index).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\delphi.jpg")
        'imgMemo(Clicou).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\delphi.jpg")
        imgMemo(Index).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\delphi.jpg")
        imgMemo(Clicou).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\delphi.jpg")
        imgMemo(Index).Enabled = True
        imgMemo(Clicou).Enabled = True
        Primeiro = True
        
    ElseIf imgMemo(Clicou).Tag = imgMemo(Index).Tag And Primeiro = False Then
        Ganhou = Ganhou + 1
        lblAcertos.Caption = Ganhou
        Primeiro = True
    Else
        Primeiro = False
    End If
    
    If Ganhou = 8 Then
        MsgBox "Você venceu!"
        Timer1.Enabled = False
        iniciaJogo
    End If
    
End Sub

Private Sub iniciaJogo()

    Primeiro = True
    Ganhou = 0
    lblAcertos.Caption = "0"
    Randomize
    For I = 0 To 15
        imgMemo(I).Tag = "Vazio"
        imgMemo(I).Enabled = True
        imgMemo(I).Visible = False
        'imgMemo(I).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\delphi.jpg")
        imgMemo(I).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\delphi.jpg")
    Next
    
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If imgMemo(posicao).Tag = "Vazio" Then
            imgMemo(posicao).Tag = "Word"
            'imgMemo(posicao).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\word.jpg")
            imgMemo(posicao).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\word.jpg")
            I = I + 1
        End If
    Loop
    
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If imgMemo(posicao).Tag = "Vazio" Then
            imgMemo(posicao).Tag = "Tabela"
            'imgMemo(posicao).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\table.jpg")
            imgMemo(posicao).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\table.jpg")
            I = I + 1
        End If
    Loop
    
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If imgMemo(posicao).Tag = "Vazio" Then
            imgMemo(posicao).Tag = "Cd"
            'imgMemo(posicao).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\cdrom01.jpg")
            imgMemo(posicao).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\cdrom01.jpg")
            I = I + 1
        End If
    Loop
    
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If imgMemo(posicao).Tag = "Vazio" Then
            imgMemo(posicao).Tag = "Copa"
            'imgMemo(posicao).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\copas.jpg")
            imgMemo(posicao).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\copas.jpg")
            I = I + 1
        End If
    Loop
    
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If imgMemo(posicao).Tag = "Vazio" Then
            imgMemo(posicao).Tag = "Paus"
            'imgMemo(posicao).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\paus.jpg")
            imgMemo(posicao).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\paus.jpg")
            I = I + 1
        End If
    Loop
    
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If imgMemo(posicao).Tag = "Vazio" Then
            imgMemo(posicao).Tag = "Ouro"
            'imgMemo(posicao).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\ouro.jpg")
            imgMemo(posicao).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\ouro.jpg")
            I = I + 1
        End If
    Loop
    
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If imgMemo(posicao).Tag = "Vazio" Then
            imgMemo(posicao).Tag = "Preto"
            'imgMemo(posicao).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\preto.jpg")
            imgMemo(posicao).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\preto.jpg")
            I = I + 1
        End If
    Loop
    
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If imgMemo(posicao).Tag = "Vazio" Then
            imgMemo(posicao).Tag = "Abrir"
            'imgMemo(posicao).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\folder.jpg")
            imgMemo(posicao).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\folder.jpg")
            I = I + 1
        End If
    Loop
    
    lblTempo.Caption = "0"
    Timer2.Enabled = True
    
End Sub


Private Sub Timer1_Timer()
    lblTempo.Caption = Val(lblTempo.Caption) + 1
    If lblTempo.Caption = "60" Then
        MsgBox "Você perdeu!"
        iniciaJogo
    End If
End Sub

Private Sub Timer2_Timer()
    For I = O To 15
        'imgMemo(I).Picture = LoadPicture("C:\vb6-master\JogoDaMemoria\midia\img\delphi.jpg")
        imgMemo(I).Picture = LoadPicture("\\SUPORTE03-PC\vb6-master\JogoDaMemoria\midia\img\delphi.jpg")
        imgMemo(I).Visible = True
    Next I
    Timer1.Enabled = True
    Timer2.Enabled = False
End Sub
