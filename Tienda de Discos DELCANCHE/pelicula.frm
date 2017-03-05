VERSION 5.00
Begin VB.Form pelicula 
   Caption         =   "pelicula"
   ClientHeight    =   5985
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame comandos 
      Caption         =   "COMANDOS"
      Height          =   1935
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   6135
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
         Left            =   3000
         TabIndex        =   7
         Top             =   600
         Width           =   2415
      End
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
         TabIndex        =   6
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame principal 
      Caption         =   "PRINCIPAL"
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   6015
      Begin VB.Data data1 
         Caption         =   "BASE DE DATOS"
         Connect         =   "Access"
         DatabaseName    =   "E:\reistro de una empresa de DISCOS\discos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "PECICULA"
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox cactor 
         DataField       =   "Cod_actor"
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
         Left            =   3000
         TabIndex        =   2
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox ctipo 
         DataField       =   "Cod_tipo"
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
         Left            =   3000
         TabIndex        =   1
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Cod_Autor"
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
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Cod_Tipo"
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
         TabIndex        =   3
         Top             =   360
         Width           =   2655
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
Attribute VB_Name = "pelicula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub agregar_Click()
If data1.Recordset.EOF = False And data1.Recordset.BOF = False Then
    data1.Recordset.AddNew
    data1.Recordset("Cod_tipo") = ctipo.Text
    data1.Recordset("Cod_actor") = cactor.Text
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
