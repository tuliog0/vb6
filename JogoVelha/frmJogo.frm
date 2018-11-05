VERSION 5.00
Begin VB.Form frmJogo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jogo da Velha"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Mesa8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   5160
      TabIndex        =   16
      Top             =   6000
      Width           =   1500
   End
   Begin VB.CommandButton Mesa7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3240
      TabIndex        =   15
      Top             =   6000
      Width           =   1500
   End
   Begin VB.CommandButton Mesa6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   1320
      TabIndex        =   14
      Top             =   6000
      Width           =   1500
   End
   Begin VB.CommandButton Mesa3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   1320
      TabIndex        =   13
      Top             =   4560
      Width           =   1500
   End
   Begin VB.CommandButton Mesa5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   5160
      TabIndex        =   12
      Top             =   4560
      Width           =   1500
   End
   Begin VB.CommandButton Mesa4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3240
      TabIndex        =   11
      Top             =   4560
      Width           =   1500
   End
   Begin VB.CommandButton Mesa2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   5160
      TabIndex        =   10
      Top             =   3120
      Width           =   1500
   End
   Begin VB.CommandButton Mesa1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3240
      TabIndex        =   9
      Top             =   3120
      Width           =   1500
   End
   Begin VB.CommandButton Mesa0 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   1320
      TabIndex        =   8
      Top             =   3120
      Width           =   1500
   End
   Begin VB.CommandButton btnIniciar 
      Caption         =   "S T A R T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   960
      TabIndex        =   7
      Top             =   2160
      Width           =   6000
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   4605
      Left            =   960
      Top             =   2760
      Width           =   6000
   End
   Begin VB.Label lblPlacarSegundo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   6360
      TabIndex        =   6
      Top             =   1440
      Width           =   1320
   End
   Begin VB.Label lblPlacarPrimeiro 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   6360
      TabIndex        =   5
      Top             =   960
      Width           =   1320
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   4920
      TabIndex        =   4
      Top             =   1440
      Width           =   1320
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   4920
      TabIndex        =   3
      Top             =   960
      Width           =   1320
   End
   Begin VB.Label lblSegundo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   4560
   End
   Begin VB.Label lblPrimeiro 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Quem tiver X em uma das caixas a frente do nome, começa o jogo."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   6885
   End
End
Attribute VB_Name = "frmJogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public verifica As Boolean
Public cont, X, y As Byte
Public ganhador, ganhador1 As String

Private Sub btnIniciar_Click()
    AtivaBotoes
    LimpaMesa
    cont = 0
End Sub

Private Sub Form_Load()
    lblPrimeiro.Caption = primeiro
    lblSegundo.Caption = segundo
    verifica = True
    cont = 0
    DesativaBotoes
End Sub

Private Sub Mesa0_Click()
    If verifica = True Then
        Mesa0.Caption = "X"
        verifica = False
        Mesa0.Enabled = False
        X = X + 1
    Else
        Mesa0.Caption = "O"
        verifica = True
        Mesa0.Enabled = False
        y = y + 1
    End If
    cont = cont + 1
    verificaGanhador
End Sub

Private Sub Mesa1_Click()
        If verifica = True Then
        Mesa1.Caption = "X"
        verifica = False
        Mesa1.Enabled = False
        X = X + 1
    Else
        Mesa1.Caption = "O"
        verifica = True
        Mesa1.Enabled = False
        y = y + 1
    End If
    cont = cont + 1
    verificaGanhador
End Sub

Private Sub Mesa2_Click()
    If verifica = True Then
        Mesa2.Caption = "X"
        verifica = False
        Mesa2.Enabled = False
        X = X + 1
    Else
        Mesa2.Caption = "O"
        verifica = True
        Mesa2.Enabled = False
        y = y + 1
    End If
    cont = cont + 1
    verificaGanhador
End Sub

Private Sub Mesa3_Click()
    If verifica = True Then
        Mesa3.Caption = "X"
        verifica = False
        Mesa3.Enabled = False
        X = X + 1
    Else
        Mesa3.Caption = "O"
        verifica = True
        Mesa3.Enabled = False
        y = y + 1
    End If
    cont = cont + 1
    verificaGanhador
End Sub

Private Sub Mesa4_Click()
    If verifica = True Then
        Mesa4.Caption = "X"
        verifica = False
        Mesa4.Enabled = False
        X = X + 1
    Else
        Mesa4.Caption = "O"
        verifica = True
        Mesa4.Enabled = False
        y = y + 1
    End If
    cont = cont + 1
    verificaGanhador
End Sub

Private Sub Mesa5_Click()
    If verifica = True Then
        Mesa5.Caption = "X"
        verifica = False
        Mesa5.Enabled = False
        X = X + 1
    Else
        Mesa5.Caption = "O"
        verifica = True
        Mesa5.Enabled = False
        y = y + 1
    End If
    cont = cont + 1
    verificaGanhador
End Sub

Private Sub Mesa6_Click()
    If verifica = True Then
        Mesa6.Caption = "X"
        verifica = False
        Mesa6.Enabled = False
        X = X + 1
    Else
        Mesa6.Caption = "O"
        verifica = True
        Mesa6.Enabled = False
        y = y + 1
    End If
    cont = cont + 1
    verificaGanhador
End Sub

Private Sub Mesa7_Click()
    If verifica = True Then
        Mesa7.Caption = "X"
        verifica = False
        Mesa7.Enabled = False
        X = X + 1
    Else
        Mesa7.Caption = "O"
        verifica = True
        Mesa7.Enabled = False
        y = y + 1
    End If
    cont = cont + 1
    verificaGanhador
End Sub

Private Sub Mesa8_Click()
    If verifica = True Then
        Mesa8.Caption = "X"
        verifica = False
        Mesa8.Enabled = False
        X = X + 1
    Else
        Mesa8.Caption = "O"
        verifica = True
        Mesa8.Enabled = False
        y = y + 1
    End If
    cont = cont + 1
    verificaGanhador
End Sub

Private Sub verificaGanhador()
    If (Mesa0.capition = "X" And Mesa1.capiton = "X" And Mesa2.Caption = "X") Or (Mesa0.capition = "O" And Mesa1.Caption = "O" And Mesa2.Caption = "O") Then
        QuemGanhou
        MsgBox "O ganhador foi: " & ganhador
        DesativaBotoes
        SomaPonto
        Exit Sub
    ElseIf (Mesa3.capition = "X" And Mesa4.capition = "X" And Mesa5.capition = "X") Or (Mesa3.capition = "O" And Mesa4.capition = "O" And Mesa5.capition = "O") Then
        QuemGanhou
        MsgBox "O ganhador foi: " & ganhador
        DesativaBotoes
        SomaPonto
        Exit Sub
    ElseIf (Mesa6.capition = "X" And Mesa7.capition = "X" And Mesa8.capition = "X") Or (Mesa6.capition = "O" And Mesa7.capition = "O" And Mesa8.capition = "O") Then
        QuemGanhou
        MsgBox "O ganhador foi: " & ganhador
        DesativaBotoes
        SomaPonto
        Exit Sub
    ElseIf (Mesa0.capition = "X" And Mesa3.capition = "X" And Mesa6.capition = "X") Or (Mesa0.capition = "O" And Mesa3.capition = "O" And Mesa6.capition = "O") Then
        QuemGanhou
        MsgBox "O ganhador foi: " & ganhador
        DesativaBotoes
        SomaPonto
        Exit Sub
    ElseIf (Mesa1.capition = "X" And Mesa4.capition = "X" And Mesa7.capition = "X") Or (Mesa1.capition = "O" And Mesa4.capition = "O" And Mesa7.capition = "O") Then
        QuemGanhou
        MsgBox "O ganhador foi: " & ganhador
        DesativaBotoes
        SomaPonto
        Exit Sub
    ElseIf (Mesa2.capition = "X" And Mesa5.capition = "X" And Mesa8.capition = "X") Or (Mesa2.capition = "O" And Mesa5.capition = "O" And Mesa8.capition = "O") Then
        QuemGanhou
        MsgBox "O ganhador foi: " & ganhador
        DesativaBotoes
        SomaPonto
        Exit Sub
    ElseIf (Mesa0.capition = "X" And Mesa4.capition = "X" And Mesa8.capition = "X") Or (Mesa0.capition = "O" And Mesa4.capition = "O" And Mesa8.capition = "O") Then
        QuemGanhou
        MsgBox "O ganhador foi: " & ganhador
        DesativaBotoes
        SomaPonto
        Exit Sub
    ElseIf (Mesa2.capition = "X" And Mesa4.capition = "X" And Mesa6.capition = "X") Or (Mesa2.capition = "O" And Mesa4.capition = "O" And Mesa6.capition = "O") Then
        QuemGanhou
        MsgBox "O ganhador foi: " & ganhador
        DesativaBotoes
        SomaPonto
        Exit Sub
    End If
    If cont = 9 Then
        MsgBox "Não houve ganhador", vbInformation, "Jogo da Velha"
        ganhador1 = "O"
        verifica = False
    End If
End Sub

Private Sub DesativaBotoes()
    Mesa0.Enabled = False
    Mesa1.Enabled = False
    Mesa2.Enabled = False
    Mesa3.Enabled = False
    Mesa4.Enabled = False
    Mesa5.Enabled = False
    Mesa6.Enabled = False
    Mesa7.Enabled = False
    Mesa8.Enabled = False
End Sub

Private Sub AtivaBotoes()
    Mesa0.Enabled = True
    Mesa1.Enabled = True
    Mesa2.Enabled = True
    Mesa3.Enabled = True
    Mesa4.Enabled = True
    Mesa5.Enabled = True
    Mesa6.Enabled = True
    Mesa7.Enabled = True
    Mesa8.Enabled = True
End Sub

Private Sub LimpaMesa()
    Mesa0.Caption = ""
    Mesa1.Caption = ""
    Mesa2.Caption = ""
    Mesa3.Caption = ""
    Mesa4.Caption = ""
    Mesa5.Caption = ""
    Mesa6.Caption = ""
    Mesa7.Caption = ""
    Mesa8.Caption = ""
End Sub

Private Sub QuemGanhou()
    If X > y Then
        ganhador = primeiro
        ganhador1 = segundo
    Else
        ganhador = segundo
        ganhador1 = primeiro
    End If
End Sub

Private Sub SomaPonto()
    If X > y Then
        lblPlacarPrimeiro.Caption = Val(lblPlacarPrimeiro.Caption) + 1
    Else
        lblPlacarSegundo.Caption = Val(lblPlacarSegundo.Caption) + 1
    End If
End Sub
