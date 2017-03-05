VERSION 5.00
Begin VB.Form registro 
   BackColor       =   &H000080FF&
   Caption         =   "registro"
   ClientHeight    =   3675
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   16.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¡Bienvenido! Tienda de Discos DVD ""DELCANCHE"""
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton Anterior 
      Caption         =   "<---Regresar"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   4155
      Left            =   0
      Picture         =   "registro.frx":0000
      Top             =   0
      Width           =   6300
   End
   Begin VB.Menu tip 
      Caption         =   "TIPOS"
   End
   Begin VB.Menu peli 
      Caption         =   "PELICULAS"
   End
   Begin VB.Menu dissc 
      Caption         =   "DISCO"
   End
   Begin VB.Menu clien 
      Caption         =   "CLIENTE"
   End
   Begin VB.Menu aut 
      Caption         =   "AUTOR"
   End
   Begin VB.Menu alqu 
      Caption         =   "ALQUILER"
   End
End
Attribute VB_Name = "registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alqu_Click()
alquiler.Show
Me.Hide
End Sub

Private Sub aut_Click()
autor.Show
Me.Hide
End Sub

Private Sub clien_Click()
cliente.Show
Me.Hide
End Sub

Private Sub dissc_Click()
disco.Show
Me.Hide
End Sub

Private Sub Label1_Click()

End Sub

Private Sub peli_Click()
pelicula.Show
Me.Hide
End Sub

Private Sub Anterior_Click()
Login.Show
Me.Hide
If Click Then
User = ""
pass = ""
End If
End Sub



Private Sub tip_Click()
tipodepelicula.Show
Me.Hide

End Sub
