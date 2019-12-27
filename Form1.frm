VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnEncripta 
      Caption         =   "desencriptar"
      Height          =   435
      Left            =   6870
      TabIndex        =   4
      Top             =   1140
      Width           =   1320
   End
   Begin VB.TextBox txtencripta 
      Height          =   840
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   900
      Width           =   6720
   End
   Begin VB.TextBox txtResultado 
      Height          =   1380
      Left            =   75
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   8070
   End
   Begin VB.TextBox txtValor 
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Text            =   "Provider=SQLNCLI11;DataBase=SIGH;Uid=sa;Pwd=Password123.;Server=10.100.14.41"
      Top             =   360
      Width           =   6720
   End
   Begin VB.CommandButton cmdEncriptar 
      Caption         =   "encriptar"
      Height          =   435
      Left            =   6870
      TabIndex        =   0
      Top             =   315
      Width           =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEncriptar_Click()
  On Error Resume Next
  Dim oCrypKey As New CrypKey.Util
  CadenaConexion = oCrypKey.EncryptString(txtValor.Text)
  txtResultado.Text = CadenaConexion
End Sub

Private Sub btnEncripta_Click()
  On Error Resume Next
  Dim oCrypKey As New CrypKey.Util
  CadenaConexion = oCrypKey.DecryptString(txtencripta.Text)
  txtResultado.Text = CadenaConexion
End Sub
