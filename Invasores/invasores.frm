VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   """Invasores"" - F2 Inicia - F3 Pausa"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   150
      Left            =   4680
      Top             =   6480
   End
   Begin VB.Timer Timer2 
      Interval        =   30
      Left            =   2640
      Top             =   3600
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   2040
      Top             =   3600
   End
   Begin VB.Label lblDev 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "oncalves5_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   5
      Left            =   4800
      TabIndex        =   7
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   4
      Left            =   4920
      TabIndex        =   6
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   5
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   2
      Left            =   4920
      TabIndex        =   4
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   1
      Left            =   4920
      TabIndex        =   3
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   0
      Left            =   4920
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image alienDead 
      Height          =   600
      Index           =   5
      Left            =   5520
      Picture         =   "invasores.frx":0000
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   600
   End
   Begin VB.Image alienDead 
      Height          =   600
      Index           =   4
      Left            =   5520
      Picture         =   "invasores.frx":841E
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   600
   End
   Begin VB.Image alienDead 
      Height          =   600
      Index           =   3
      Left            =   5520
      Picture         =   "invasores.frx":1083C
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   600
   End
   Begin VB.Image alienDead 
      Height          =   600
      Index           =   2
      Left            =   5520
      Picture         =   "invasores.frx":18C5A
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image alienDead 
      Height          =   600
      Index           =   1
      Left            =   5520
      Picture         =   "invasores.frx":21078
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   600
   End
   Begin VB.Image alienDead 
      Height          =   600
      Index           =   0
      Left            =   5520
      Picture         =   "invasores.frx":29496
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   600
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblShowScore 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000004&
      X1              =   4680
      X2              =   4680
      Y1              =   0
      Y2              =   7320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   6240
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   6975
      Left            =   4680
      Top             =   0
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2280
      Picture         =   "invasores.frx":318B4
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   0
      Top             =   6960
      Width           =   6255
   End
   Begin VB.Image Alien 
      Height          =   525
      Index           =   5
      Left            =   3720
      Picture         =   "invasores.frx":3253F
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   405
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   3000
      X2              =   3000
      Y1              =   6120
      Y2              =   6300
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   3840
      Picture         =   "invasores.frx":3A95D
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   530
      Index           =   4
      Left            =   3000
      Picture         =   "invasores.frx":3C276
      Stretch         =   -1  'True
      Top             =   960
      Width           =   410
   End
   Begin VB.Image Alien 
      Height          =   530
      Index           =   3
      Left            =   2400
      Picture         =   "invasores.frx":44694
      Stretch         =   -1  'True
      Top             =   240
      Width           =   410
   End
   Begin VB.Image Alien 
      Height          =   530
      Index           =   2
      Left            =   1680
      Picture         =   "invasores.frx":4CAB2
      Stretch         =   -1  'True
      Top             =   240
      Width           =   410
   End
   Begin VB.Image Alien 
      Height          =   530
      Index           =   1
      Left            =   600
      Picture         =   "invasores.frx":54ED0
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   410
   End
   Begin VB.Image Alien 
      Height          =   530
      Index           =   0
      Left            =   120
      Picture         =   "invasores.frx":5D2EE
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   410
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   6975
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lugar As Boolean
Dim posicao As Integer
Dim matou As Byte
Dim xy As Long


Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal IpszSoundName As String, ByVal uFlags As Long) As Long


Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal IpstrCommand As String, ByVal IpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long


Private Sub form_keydown(KeyCode As Integer, Shift As Integer)
    
    Dim rc As Integer
    If KeyCode = vbKeyF2 Then
        IniciaJogo
        
        ElseIf KeyCode = vbKeyF3 Then ' F3 pausa o jogo
            Timer1.Enabled = False
            Timer1.Enabled = False
            MsgBox "Pausa", vbExclamation, "Jogo"
            Timer1.Enabled = True
            Timer1.Enabled = True
            
        ElseIf KeyCode = vbKeySpace And Line2.Visible = False Then ' Espaço atira
            rc = sndPlaySound("C:\vb6-master\Invasores\midia\Sounds\deagle-1.wav", SND_ASYNC) ' aciona som de tiro
            Line2.Visible = True
            Line2.X1 = Image1.Left + 230
            Line2.X2 = Image1.Left + 230
            Line2.Y1 = Image1.Top - 360
            Line2.Y2 = Image1.Top
            Timer2.Enabled = True
            
        ElseIf KeyCode = vbKeyRight Then ' Movimentação da arma para direita
        
            If Image1.Left < 4200 Then
                Image1.Left = Image1.Left + 200
            End If
            
        ElseIf KeyCode = vbKeyLeft Then ' Movimentação da arma para esquerda
        
            If Image1.Left > 0 Then
                Image1.Left = Image1.Left - 200
            End If
    End If

End Sub


Private Sub form_load()

    IniciaJogo ' Chama função Iniciar jogo
    
End Sub


Private Sub IniciaJogo()
        
    If MsgBox("Deseja um jogo mais difícil?" & vbCrLf & vbCrLf & "Sim - dificil" & "    /   " & "Não - fácil", vbYesNo, "Selecione o nível de jogo.") = vbNo Then
    ' Caso o jogador selecionar não na mensagem iniciara os "dab" com mais espaço.
        Line2.Visible = False
        lblScore.Caption = 0
        
        For i = 0 To 5 ' load 6 pictures dab
            Alien(i).Visible = True
            Alien(i).Picture = LoadPicture("C:\vb6-master\Invasores\midia\img\dab.jpg")
        Next i
        
        For i = 0 To 5 ' load 6 pictures dead
            Label1(i).Visible = False
            alienDead(i).Visible = False
            alienDead(i).Picture = LoadPicture("C:\vb6-master\Invasores\midia\img\emoji.jpg")
        Next i
        
        Image1.Picture = LoadPicture("C:\vb6-master\Invasores\midia\img\image2.jpg")
        Timer1.Enabled = True
        Timer2.Enabled = True
        Timer3.Enabled = True
        matou = 0
        Alien(0).Left = 4199
        Alien(0).Top = 0
        Alien(1).Left = 3675
        Alien(1).Top = -600
        Alien(2).Left = 3150
        Alien(2).Top = -1200
        Alien(3).Left = 2100
        Alien(3).Top = -2000
        Alien(4).Left = 1050
        Alien(4).Top = -2600
        Alien(5).Left = 525
        Alien(5).Top = -3400
        rc = sndPlaySound("C:\vb6-master\Invasores\midia\Sounds\hos2.wav", SND_ASYNC)
        
        Else ' Caso o jogador selecionar 'SIM' na mensagem ou qualquer outra coisa... o jogo iniciara com os "dab" mais proximos dificultando o jogo.
            Line2.Visible = False
            lblScore.Caption = 0
            
            For i = 0 To 5 ' load 6 pictures dab
                Label1(i).Visible = False
                Alien(i).Visible = True
                Alien(i).Picture = LoadPicture("C:\vb6-master\Invasores\midia\img\dab.jpg")
            Next i
            
            For i = 0 To 5 ' load 6 pictures dead
            alienDead(i).Visible = False ' Torna os deads invisíveis
            alienDead(i).Picture = LoadPicture("C:\vb6-master\Invasores\midia\img\emoji.jpg")
            Next i
            
            Image1.Picture = LoadPicture("C:\vb6-master\Invasores\midia\img\image2.jpg")
            Timer1.Enabled = True
            Timer2.Enabled = True
            Timer3.Enabled = True
            matou = 0
            Alien(0).Left = 4199
            Alien(0).Top = 0
            Alien(1).Left = 3675
            Alien(1).Top = -50
            Alien(2).Left = 3150
            Alien(2).Top = -100
            Alien(3).Left = 2100
            Alien(3).Top = -150
            Alien(4).Left = 1050
            Alien(4).Top = -200
            Alien(5).Left = 525
            Alien(5).Top = -250
            rc = sndPlaySound("C:\vb6-master\Invasores\midia\Sounds\hos3.wav", SND_ASYNC)
    End If
        
End Sub


Private Sub Timer1_Timer()

    For i = 0 To 5
    
        posicao = Int(Rnd * 100) ' Randomizando posição para o "dab" (dentro do campo)
        posicaoFimLeft = Int(Rnd * 4199) ' Randomizando posição para o "dab" quando acaba o campo

        If Alien(i).Top < 6360 Then 'limite de campo abaixo
            Alien(i).Top = Alien(i).Top + 100
            If lugar = False Then
                If Alien(i).Left < 4200 Then 'limite de campo a direita
                    Alien(i).Left = Alien(i).Left + posicao
                    lugar = True
                Else
                    Alien(i).Left = Alien(i).Left - posicaoFimLeft ' recolocando o "dab" quando acaba o campo
                    lugar = True
                End If
            Else
                If Alien(i).Left > 120 Then 'limite de campo a esquerda
                    Alien(i).Left = Alien(i).Left - posicao
                    lugar = False
                Else
                    Alien(i).Left = Alien(i).Left + posicaoFimLeft ' recolocando o "dab" quando acaba o campo
                    lugar = False
                End If
            End If
        ' Caso queira retornar o "dab" ao inicio depois de tocar a extremidade de baixo (descomentar abaixo)
        'Else
            'Alien(i).Top = -30
            
        End If
        
        If Alien(i).Visible = True Then
            ' Caso queira retornar o "dab" ao inicio depois de tocar a extremidade de baixo (descomentar abaixo)
            'If Alien(i).Left >= Image1.Left And Alien(i).Left <= Image1.Left + 480 Then
                'If Alien(i).Top + 480 >= 6000 And Alien(i).Top + 480 <= 6400 Then
            ' Caso queira retornar o "dab" ao inicio depois de tocar a extremidade de baixo (comentar somente o proximo if)
                    If Alien(i).Top >= 6400 Then
                        rc = sndPlaySound("C:\vb6-master\Invasores\midia\Sounds\gg_brass_bell.wav", SND_ASYNC)
                        Image1.Picture = Image2.Picture
                        Timer1.Enabled = False
                        Timer2.Enabled = False
                            If MsgBox("Você perdeu...", vbOK, "Que pena.") = vbOK Or vbCancel Then
                                IniciaJogo
                            Else
                                Exit Sub
                            End If
                    End If
                'End If ' Caso queira retornar o "dab" ao inicio depois de tocar a extremidade de baixo (descomentar esse ENDIF)
            'End If ' Caso queira retornar o "dab" ao inicio depois de tocar a extremidade de baixo (descomentar esse ENDIF)
        End If
    Next i

End Sub


Private Sub timer2_timer()

    Line2.Y1 = Line2.Y1 - 250
    Line2.Y2 = Line2.Y2 - 250
    conta = conta + 1
    
    If Line2.Y1 < 0 Then
        Line2.Visible = False
        Timer2.Enabled = False
    End If
    
    For x = 0 To 5 ' Verificação de acertos da bala vs "dab"
        pega = Alien(x).Left
            If Line2.X1 >= Alien(x).Left And Line2.X2 <= Alien(x).Left + 480 Then
                If Line2.Y1 >= Alien(x).Top And Line2.Y2 <= Alien(x).Top + 680 Then
                    If Alien(x).Visible = True And Line2.Visible = True Then
                        'Line2.Visible = False    Após matar o primeiro "dab" a bala não continua o curso
                        rc = sndPlaySound("C:\vb6-master\Invasores\midia\Sounds\bhit_flesh-1.wav", SND_ASYNC)
                        Alien(x).Visible = False ' Tira o "dab" que morreu
                        alienDead(x).Visible = True ' Mostra no controle o "dab" morto
                        matou = matou + 1 ' Contagem de acertos
                        Label1(x).Caption = x + 1 ' coloca no caption o numero do "dab" morto
                        Label1(x).Visible = True ' Mostra no controle o numero do "dab" morto
                        lblScore.Caption = matou
                    End If
                   ' voltar
                End If
            End If
    Next x
    
    If matou = 6 Then ' Após matar os seis "dab" o jogador ganha e fim de jogo
    rc = sndPlaySound("C:\vb6-master\Invasores\midia\Sounds\gg_brass_bell.wav", SND_ASYNC)
    
        If MsgBox("Você ganhou!!!", vbOK, "Parabéns!") = vbOK Or vbCancel Then
            Timer1.Enabled = False
            Timer2.Enabled = False
            IniciaJogo 'FIM DE JOGO (reinicia o jogo)
        Else
            Timer1.Enabled = False
            Timer2.Enabled = False
        End If
    End If

End Sub

Private Sub Timer3_Timer()
    Dim dev
    dev = Array("D", "e", "s", "e", "n", "v", "o", "l", "v", "i", "d", "o", "_", "p", "o", "r", "_", "t", "u", "l", "i", "o", "g", "o", "n", "c", "a", "l", "v", "e", "s", "5", "_")
    
    lblDev.Caption = lblDev.Caption + dev(xy)
    xy = xy + 1
    
    If xy > 32 Then '32 5  31 7
        lblDev.Caption = "oncalves5_"
        xy = 0
    End If
    
End Sub
