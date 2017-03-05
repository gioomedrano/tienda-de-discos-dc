VERSION 5.00
Begin VB.Form cliente 
   Caption         =   "Form1"
   ClientHeight    =   6915
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame comandos 
      Caption         =   "COMANDOS"
      Height          =   1575
      Left            =   600
      TabIndex        =   9
      Top             =   4800
      Width           =   6135
      Begin VB.CommandButton agregar 
         Caption         =   "Agregar"
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
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton eliminar 
         Caption         =   "Eliminar"
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
         Left            =   3120
         TabIndex        =   10
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame principal 
      Caption         =   "PRINCIPAL"
      Height          =   4455
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.TextBox tel 
         DataField       =   "Telefono"
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
         TabIndex        =   8
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox membre 
         DataField       =   "Num_membresia"
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
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox nomb 
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
         TabIndex        =   2
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox direc 
         DataField       =   "Direccion"
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
         TabIndex        =   1
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Data data1 
         Caption         =   "BASE DE DATOS"
         Connect         =   "Access"
         DatabaseName    =   "E:\reistro de una empresa de DISCOS\discos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CLIENTE"
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Telefono"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Num_Membresia"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Direccion"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   2295
      End
   End
   Begin VB.Menu mprincipal 
      Caption         =   "MENU PRINCIPAL"
   End
   Begin VB.Menu volver 
      Caption         =   "VOLVER"
   End
   Begin VB.Menu salir 
      Caption         =   "SALIR"
   End
End
Attribute VB_Name = "cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub agregar_Click()
If data1.Recordset.EOF = False And data1.Recordset.BOF = False Then
    data1.Recordset.AddNew
    data1.Recordset("Num_membresia") = membre.Text
    data1.Recordset("Nombre") = nomb.Text
    data1.Recordset("Direccion") = direc.Text
    data1.Recordset("Telefono") = tel.Text
    data1.Recordset.Update
    End If
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

Private Sub mprincipal_Click()
Login.Show
Me.Hide
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
