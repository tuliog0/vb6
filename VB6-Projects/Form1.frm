VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Sistema de Vendas"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnExcluir 
      Caption         =   "Excluir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   485
      Left            =   5760
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton btnAlterar 
      Caption         =   "Alterar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   485
      Left            =   4560
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cadastro de Produtos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3240
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton btnPesquisar 
         Caption         =   "Pesquisar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   485
         Left            =   6840
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton btnInserir 
         Caption         =   "Inserir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   485
         Left            =   3240
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblPrecoProd 
         Caption         =   "Preço:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblDescProd 
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblCodProd 
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
