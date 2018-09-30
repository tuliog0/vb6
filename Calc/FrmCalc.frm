VERSION 5.00
Begin VB.Form FrmCalc 
   Caption         =   "Calculadora"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtNumber2 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   500
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "use somente numeros"
      Top             =   720
      Width           =   1000
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Limpar"
      Height          =   1095
      Left            =   1560
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton btnMultiplicar 
      Caption         =   "*"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton btnSomar 
      Caption         =   "+"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton btnSubtrair 
      Caption         =   "-"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton btnDividir 
      Caption         =   "/"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtNumber1 
      Alignment       =   2  'Center
      Height          =   500
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1000
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1680
      TabIndex        =   7
      Top             =   360
      Width           =   120
   End
End
Attribute VB_Name = "FrmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim valor1 As Single
    Dim valor2 As Single
    
Private Sub btnClear_Click()
    txtNumber1.Text = ""
    txtNumber2.Text = ""
    lblResult.Caption = ""
    txtNumber1.SetFocus
End Sub

Private Sub btnMultiplicar_Click()

    valor1 = txtNumber1.Text
    valor2 = txtNumber2.Text
    
    lblResult.Caption = valor1 * valor2
    
End Sub

Private Sub Command1_Click()
    Form1.Show
End Sub




Private Sub Form_Load()
    txtNumber2.Enabled = False
    btnDividir.Enabled = False
    btnMultiplicar.Enabled = False
    btnSubtrair.Enabled = False
    btnSomar.Enabled = False
    btnClear.Enabled = False
End Sub

Private Sub txtNumber1_KeyPress(KeyAscii As Integer)

' Positive numbers only, with "." decimal separator
    Select Case KeyAscii
        Case Asc("0") To Asc("9")       ' if in range 0 to 9 then retain KeyAscii value - goto End Select
        Case Asc(",")
            If InStr(1, txtNumber1.Text, ".") > 0 Then  ' count number of periods
                KeyAscii = 0                            ' if 1 already
            End If                                      ' return Ascii 0 (Null) (for this KeyPress)
        Case Else
            KeyAscii = 0        ' any other character: return Ascii 0 (Null)
    End Select

End Sub



Private Sub txtNumber2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Asc(",")
            If InStr(1, txtNumber1.Text, ".") > 0 Then
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub
