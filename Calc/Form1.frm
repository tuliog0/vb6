VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Calculadora Curso"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtResultado 
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdMultiplicar 
      Caption         =   "X"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdDividir 
      Caption         =   "/"
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSomar 
      Caption         =   "+"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSubtrair 
      Caption         =   "-"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtNum2 
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtNum1 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lbl2 
      Caption         =   "Segundo Número:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lbl1 
      Caption         =   "Primeiro Número:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim valor1 As Single
Dim valor2 As Single


Private Sub cmdDividir_Click()
    valor1 = txtNum1.Text
    valor2 = txtNum2.Text
    txtResultado.Text = valor1 / valor2
End Sub

Private Sub cmdLimpar_Click()
    txtResultado.Text = ""
    txtNum1.Text = ""
    txtNum2.Text = ""
    txtNum1.SetFocus
End Sub

Private Sub cmdMultiplicar_Click()
    valor1 = txtNum1.Text
    valor2 = txtNum2.Text
    txtResultado.Text = valor1 * valor2
End Sub

Private Sub cmdSomar_Click()
    valor1 = txtNum1.Text
    valor2 = txtNum2.Text
    txtResultado.Text = valor1 + valor2
End Sub

Private Sub cmdSubtrair_Click()
    valor1 = txtNum1.Text
    valor2 = txtNum2.Text
    txtResultado.Text = valor1 - valor2
End Sub
