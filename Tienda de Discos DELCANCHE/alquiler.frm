VERSION 5.00
Begin VB.Form alquiler 
   Caption         =   "alquiler"
   ClientHeight    =   6375
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame comandos 
      Caption         =   "COMANDOS"
      Height          =   1335
      Left            =   1560
      TabIndex        =   9
      Top             =   4800
      Width           =   7455
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
         Left            =   4080
         TabIndex        =   11
         Top             =   360
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
         Left            =   1200
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame principal 
      Caption         =   "PRINCIPAL"
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   11535
      Begin VB.TextBox cant 
         DataField       =   "Cantidad"
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
         Left            =   8400
         TabIndex        =   17
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox valalq 
         DataField       =   "valor alquiler"
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
         Left            =   8400
         TabIndex        =   15
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox fecdev 
         DataField       =   "Fecha de devolucion"
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
         Left            =   8400
         TabIndex        =   13
         Top             =   1080
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
         Left            =   4920
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ALQUILER"
         Top             =   3600
         Width           =   2775
      End
      Begin VB.TextBox ccliente 
         DataField       =   "Cod_cliente"
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
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox cdisco 
         DataField       =   "Cod_disco"
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
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox codigo 
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
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox fealqu 
         DataField       =   "Fecha de alquiler"
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
         Left            =   8520
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Roman"
            Size            =   24
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6120
         TabIndex        =   16
         Top             =   2640
         Width           =   1725
      End
      Begin VB.Label Label6 
         Caption         =   "VALOR ALQUILER"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   14
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "FECH_DEVOLUCION"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   12
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "COD_CLIENTE"
         BeginProperty Font 
            Name            =   "Roman"
            Size            =   15.75
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "COD_DISCO"
         BeginProperty Font 
            Name            =   "Roman"
            Size            =   26.25
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Roman"
            Size            =   26.25
            Charset         =   255
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
      Begin VB.Label Label4 
         Caption         =   "FECH_ALQUILER"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   5
         Top             =   480
         Width           =   2535
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
Attribute VB_Name = "alquiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub agregar_Click()
If data1.Recordset.EOF = False And data1.Recordset.BOF = False Then
    data1.Recordset.AddNew
    data1.Recordset("Codigo") = codigo.Text
    data1.Recordset("Cod_disco") = cdisco.Text
    data1.Recordset("Cod_cliente") = ccliente.Text
    data1.Recordset("Fecha de alquiler") = fealqu.Text
    data1.Recordset("Fecha de devolucion") = fecdev.Text
    data1.Recordset("valor alquiler") = valalq.Text
    data1.Recordset("Cantidad") = cant.Text
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
