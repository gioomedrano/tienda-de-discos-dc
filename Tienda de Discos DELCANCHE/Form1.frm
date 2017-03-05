VERSION 5.00
Begin VB.Form Login 
   AutoRedraw      =   -1  'True
   Caption         =   "Login"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   6735
      Begin VB.CommandButton Salir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3600
         TabIndex        =   6
         Top             =   3360
         Width           =   2655
      End
      Begin VB.CommandButton Ingresar 
         Caption         =   "Ingresar"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox txtpass 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3240
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "password1234"
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtuser 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Text            =   "Pablo"
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Contraseña"
         BeginProperty Font 
            Name            =   "AR CARTER"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   3360
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "AR CARTER"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   1545
      End
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mostrar_Click()
con = con + 1
If (con / 2) = Int((con / 2)) Then
txtpass.PasswordChar = "#"
End If

End Sub

Private Sub Ingresar_Click()
If txtuser.Text = "Pablo" And txtpass.Text = "password1234" Then
registro.Show
Me.Hide
Else
MsgBox "ACCESO DENEGADO", , "Alerta"
End If
End Sub

Private Sub salir_Click()
End
End Sub

