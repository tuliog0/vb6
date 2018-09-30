VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   """Invasores"" - F2 Inicia - F3 Pausa"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
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
   Begin VB.Line Line3 
      BorderColor     =   &H80000004&
      X1              =   4680
      X2              =   4680
      Y1              =   0
      Y2              =   6960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   4800
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   6975
      Left            =   4680
      Top             =   0
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2280
      Picture         =   "invasores.frx":0000
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
      Width           =   4815
   End
   Begin VB.Image Alien 
      Height          =   525
      Index           =   5
      Left            =   3720
      Picture         =   "invasores.frx":0C8B
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
      Picture         =   "invasores.frx":90A9
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   530
      Index           =   4
      Left            =   3000
      Picture         =   "invasores.frx":A9C2
      Stretch         =   -1  'True
      Top             =   960
      Width           =   410
   End
   Begin VB.Image Alien 
      Height          =   530
      Index           =   3
      Left            =   2400
      Picture         =   "invasores.frx":12DE0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   410
   End
   Begin VB.Image Alien 
      Height          =   530
      Index           =   2
      Left            =   1680
      Picture         =   "invasores.frx":1B1FE
      Stretch         =   -1  'True
      Top             =   240
      Width           =   410
   End
   Begin VB.Image Alien 
      Height          =   530
      Index           =   1
      Left            =   600
      Picture         =   "invasores.frx":2361C
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   410
   End
   Begin VB.Image Alien 
      Height          =   530
      Index           =   0
      Left            =   120
      Picture         =   "invasores.frx":2BA3A
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
        
    If MsgBox("Deseja um jogo mais Dificil?" & vbCrLf & "Yes - Dificil" & vbCrLf & "No - Fácil", vbYesNo, "Selecione o estilo.") = vbNo Then
    ' Caso o jogador selecionar não na mensagem iniciara os "dab" com mais espaço.
        Line2.Visible = False
        
        For i = 0 To 5 ' load 6 pictures
            Alien(i).Visible = True
            Alien(i).Picture = LoadPicture("C:\vb6-master\Invasores\midia\img\dab.jpg")
        Next i
        
        Image1.Picture = LoadPicture("C:\vb6-master\Invasores\midia\img\image2.jpg")
        Timer1.Enabled = True
        Timer2.Enabled = True
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
        
        Else ' Caso o jogador selecionar 'SIM' na mensagem ou qualquer outra coisa... o jogo iniciara com os "dab" mais proximos dificultando o jogo.
            Line2.Visible = False
            
            For i = 0 To 5 ' load 6 pictures
                Alien(i).Visible = True
                Alien(i).Picture = LoadPicture("C:\vb6-master\Invasores\midia\img\dab.jpg")
            Next i
            
            Image1.Picture = LoadPicture("C:\vb6-master\Invasores\midia\img\image2.jpg")
            Timer1.Enabled = True
            Timer2.Enabled = True
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
                            If MsgBox("Você perdeu...", vbOK, "Jogo") = vbOK Or vbCancel Then
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
                    If Alien(x).Visible = True Then
                        'Line2.Visible = False    Após matar o primeiro "dab" a bala continua o curso
                        Alien(x).Visible = False
                        matou = matou + 1 ' Contagem de acertos
                    End If
                   ' voltar
                End If
            End If
    Next x
    
    If matou = 6 Then ' Após matar os seis "dab" o jogador ganha e fim de jogo
        If MsgBox("Você ganhou!!!", vbOK, "Jogo") = vbOK Or vbCancel Then
            Timer1.Enabled = False
            Timer2.Enabled = False
            IniciaJogo 'FIM DE JOGO (reinicia o jogo)
        Else
            Timer1.Enabled = False
            Timer2.Enabled = False
        End If
    End If

End Sub
