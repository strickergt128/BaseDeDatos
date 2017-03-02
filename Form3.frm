VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8745
   LinkTopic       =   "Form3"
   ScaleHeight     =   3045
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Tipo de Pelicula"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pelicula"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cliente"
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cassette"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Alquiler"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actor"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Hide
Form1.Show
End Sub

Private Sub Command2_Click()
Form3.Hide
Form2.Show
End Sub

Private Sub Command3_Click()
Form3.Hide
Form4.Show
End Sub

Private Sub Command4_Click()
Form3.Hide
Form5.Show
End Sub

Private Sub Command5_Click()
Form3.Hide
Form6.Show
End Sub

Private Sub Command6_Click()
Form3.Hide
Form7.Show
End Sub

Private Sub Command8_Click()
End
End Sub
