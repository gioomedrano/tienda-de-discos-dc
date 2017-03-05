VERSION 5.00
Begin VB.Form autor 
   Caption         =   "autor"
   ClientHeight    =   5925
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame comandos 
      Caption         =   "COMANDOS"
      Height          =   1575
      Left            =   480
      TabIndex        =   7
      Top             =   4080
      Width           =   5295
      Begin VB.CommandButton eliminar 
         Caption         =   "ELIMINAR"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         TabIndex        =   9
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton new 
         Caption         =   "AGREGAR"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame principal 
      Caption         =   "PRINCIPAL"
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   -240
      Width           =   6015
      Begin VB.Data data1 
         Caption         =   "BASE DE DATOS"
         Connect         =   "Access"
         DatabaseName    =   "E:\reistro de una empresa de DISCOS\discos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "AUTOR"
         Top             =   3120
         Width           =   2775
      End
      Begin VB.TextBox fechnac 
         DataField       =   "Fecha de nacimiento"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   6
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox nombre 
         DataField       =   "Nombre"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   5
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox cod 
         DataField       =   "Codigo"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   4
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Nacimiento"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   720
         TabIndex        =   3
         Top             =   1800
         Width           =   2835
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   2
         Top             =   1080
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   600
         TabIndex        =   1
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.Menu menu 
      Caption         =   "MENU PRICIPAL"
   End
   Begin VB.Menu volver 
      Caption         =   "VOLVER"
   End
   Begin VB.Menu salir 
      Caption         =   "SALIR"
   End
End
Attribute VB_Name = "autor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub hola_Click()

End Sub

Private Sub eliminar_Click()
If data1.Recordset.EOF = False And data1.Recordset.BOF = False Then
    data1.Recordset.delete
    data1.Recordset.Requery
    End If
End Sub

Private Sub left_Click()
data1.Recordset.MovePrevious
If data1.Recordset.BOF = True Then
    data1.Recordset.MoveLast
 End If
End Sub

Private Sub menu_Click()
Login.Show
Me.Hide
End Sub

Private Sub new_Click()
If data1.Recordset.EOF = False And data1.Recordset.BOF = False Then
    data1.Recordset.AddNew
    data1.Recordset("Codigo") = cod.Text
    data1.Recordset("Nombre") = nombre.Text
    data1.Recordset("Fecha de nacimiento") = fechnac.Text
    data1.Recordset.Update
    End If
End Sub

Private Sub rigth_Click()
data1.Recordset.MoveNext
If data1.Recordset.EOF = True Then
data1.Recordset.MoveFirst
End If
End Sub

Private Sub salir_Click()
End
End Sub

Private Sub volver_Click()
registro.Show
Me.Hide
End Sub
