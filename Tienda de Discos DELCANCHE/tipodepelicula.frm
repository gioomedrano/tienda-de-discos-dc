VERSION 5.00
Begin VB.Form tipodepelicula 
   Caption         =   "tipo de pelicula"
   ClientHeight    =   5985
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8430
   Icon            =   "tipodepelicula.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame comandos 
      Caption         =   "COMANDOS"
      Height          =   1335
      Left            =   600
      TabIndex        =   5
      Top             =   3840
      Width           =   6495
      Begin VB.CommandButton new 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton delete 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3240
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame principal 
      Caption         =   "PRINCIPAL"
      Height          =   3495
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.TextBox tipo 
         DataField       =   "Tipo"
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
         Left            =   1440
         TabIndex        =   2
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox cate 
         DataField       =   "Categoria"
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
         Left            =   1440
         TabIndex        =   1
         Top             =   2160
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
         RecordSource    =   "TIPO DE PELICULA"
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   2280
         TabIndex        =   4
         Top             =   120
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Categoria"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   1800
         TabIndex        =   3
         Top             =   1440
         Width           =   1980
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
Attribute VB_Name = "tipodepelicula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub delete_Click()
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

Private Sub new_Click()
If data1.Recordset.EOF = False And data1.Recordset.BOF = False Then
    data1.Recordset.AddNew
    data1.Recordset("Tipo") = tipo.Text
    data1.Recordset("Categoria") = cate.Text
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
