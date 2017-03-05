VERSION 5.00
Begin VB.Form disco 
   Caption         =   "disco"
   ClientHeight    =   6765
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame comandos 
      Caption         =   "COMANDOS"
      Height          =   1695
      Left            =   1080
      TabIndex        =   7
      Top             =   4920
      Width           =   5415
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
         Left            =   120
         TabIndex        =   9
         Top             =   600
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
         Left            =   2760
         TabIndex        =   8
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame principal 
      Caption         =   "PRINCIPAL"
      Height          =   4575
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      Begin VB.TextBox format 
         DataField       =   "Formato"
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
         Left            =   3360
         TabIndex        =   10
         Top             =   2880
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
         Left            =   3360
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox cop 
         DataField       =   "Num_copias"
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
         Left            =   3360
         TabIndex        =   2
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox peli 
         DataField       =   "Cod_pelicula"
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
         Left            =   3360
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
         Left            =   3120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "DISCO"
         Top             =   4080
         Width           =   2775
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Formato"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   21.75
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
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Num_Copias"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Cod_Pelicula"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   3135
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
Attribute VB_Name = "disco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub agregar_Click()
If data1.Recordset.EOF = False And data1.Recordset.BOF = False Then
    data1.Recordset.AddNew
    data1.Recordset("Codigo") = cod.Text
    data1.Recordset("Num_copias") = cop.Text
    data1.Recordset("Cod_pelicula") = peli.Text
    data1.Recordset("Formato") = format.Text
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
