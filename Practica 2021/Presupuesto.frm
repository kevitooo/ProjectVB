VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Presupuesto 
   BackColor       =   &H80000003&
   Caption         =   "Presupuesto"
   ClientHeight    =   10245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18405
   Icon            =   "Presupuesto.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10245
   ScaleWidth      =   18405
   Begin VB.ListBox ListProd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      ItemData        =   "Presupuesto.frx":1084A
      Left            =   12600
      List            =   "Presupuesto.frx":1084C
      TabIndex        =   49
      Top             =   720
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.TextBox txtDIVA 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   48
      Top             =   8280
      Width           =   975
   End
   Begin VB.TextBox txtDCant 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   5
      Top             =   8280
      Width           =   975
   End
   Begin VB.TextBox txtDIP 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15720
      TabIndex        =   47
      Top             =   8280
      Width           =   1815
   End
   Begin VB.TextBox txtDDesc 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14160
      TabIndex        =   6
      Top             =   8280
      Width           =   1575
   End
   Begin VB.TextBox txtDPrecio 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   46
      Top             =   8280
      Width           =   1695
   End
   Begin VB.TextBox txtDetalle 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   45
      Top             =   8280
      Width           =   8895
   End
   Begin VB.TextBox txtDCod 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   51
      Top             =   8280
      Width           =   1335
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16440
      TabIndex        =   44
      Top             =   9360
      Width           =   1575
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15360
      TabIndex        =   43
      Top             =   9360
      Width           =   1095
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14160
      TabIndex        =   42
      Top             =   9360
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12840
      TabIndex        =   41
      Top             =   9360
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   40
      Top             =   9360
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   39
      Top             =   9360
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   38
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   8
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton btnBorrar 
      Caption         =   "&Borrar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   9
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton btnSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   10
      Top             =   9120
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid grilla1 
      Height          =   2535
      Left            =   360
      TabIndex        =   23
      Top             =   5760
      Width           =   17535
      _ExtentX        =   30930
      _ExtentY        =   4471
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000003&
      Caption         =   "Datos del comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      TabIndex        =   20
      Top             =   3480
      Width           =   12015
      Begin MSMask.MaskEdBox maskVence 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox maskFecha 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbCV 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6360
         TabIndex        =   4
         Text            =   "Seleccionar venta"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox cmbTipo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5040
         TabIndex        =   2
         Text            =   "Seleccionar tipo"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtNum 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000003&
         Caption         =   "Número:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8280
         TabIndex        =   36
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000003&
         Caption         =   "Condición venta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   35
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000003&
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   34
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000003&
         Caption         =   "Vence:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   33
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000003&
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   32
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.ListBox ListPresu 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      ItemData        =   "Presupuesto.frx":1084E
      Left            =   12600
      List            =   "Presupuesto.frx":10850
      TabIndex        =   19
      Top             =   720
      Width           =   5295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "Datos del cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   360
      TabIndex        =   18
      Top             =   480
      Width           =   12015
      Begin MSMask.MaskEdBox maskCUIT 
         Height          =   375
         Left            =   2280
         TabIndex        =   52
         Top             =   1920
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbIVA 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8280
         TabIndex        =   13
         Text            =   "Seleccionar IVA"
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtNom 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   11
         Top             =   480
         Width           =   4695
      End
      Begin VB.TextBox txtProv 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   16
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox txtLoc 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   15
         Top             =   1440
         Width           =   4455
      End
      Begin VB.TextBox txtCP 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtDirec 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtNro 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   0
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000003&
         Caption         =   "IVA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   31
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000003&
         Caption         =   "Provincia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   30
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000003&
         Caption         =   "Localidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   29
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000003&
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   28
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000003&
         Caption         =   "C.U.I.T.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000003&
         Caption         =   "C. Postal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   26
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000003&
         Caption         =   "Dirección:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000003&
         Caption         =   "Número:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label lblPro 
      BackColor       =   &H80000003&
      Caption         =   "Listado de productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13800
      TabIndex        =   50
      Top             =   240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000003&
      Caption         =   "N. Grav. 21%          IVA 21%          N. Grav. 10%          IVA 10%          No Insc.          Exento          Total Compb."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   37
      Top             =   9000
      Width           =   9135
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000003&
      Caption         =   "Detalle de los items a facturar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   22
      Top             =   5400
      Width           =   3735
   End
   Begin VB.Label lblCli 
      BackColor       =   &H80000003&
      Caption         =   "Listado de clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13800
      TabIndex        =   21
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Presupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Desc As Double
Dim PrecioFinal As Double
Dim Producto As String

Private Sub Form_Load()
    Presupuesto.Height = 10830
    Presupuesto.Width = 18645
    Presupuesto.Top = ((Screen.Height - Me.Height) / 2) - 650
    Presupuesto.Left = ((Screen.Width - Me.Width) / 2)
    
    cmbTipo.AddItem "Presupuesto A"
    cmbTipo.AddItem "Presupuesto B"
    
    cmbCV.AddItem "Contado"
    cmbCV.AddItem "Tarjeta de crédito"
    cmbCV.AddItem "Tarjeta de débito"
    cmbCV.AddItem "Cheque"
    
    camino = App.Path
    Set db = OpenDatabase(camino & "\dbInfor.mdb")
    
    Set TBCLIENTES = db.OpenRecordset("TBCLIENTES")
    TBCLIENTES.Index = "NroI_cli"
    
    Set TBCODPOS = db.OpenRecordset("TBCODPOS")
    TBCODPOS.Index = "NroI_cpos"
    
    Set TBDETALLE = db.OpenRecordset("TBDETALLE")
    TBDETALLE.Index = "PrimaryKey"
    
    Set TBCONSULTA = db.OpenRecordset("TBCLIENTES")
    TBCONSULTA.Index = "AyNI_cli"
    
    Set TBPRODUCTOS = db.OpenRecordset("TBPRODUCTOS")
    TBPRODUCTOS.Index = "Nro_proI"
    
    Set TBCONSULTA2 = db.OpenRecordset("TBPRODUCTOS")
    TBCONSULTA2.Index = "Nro_proI"
    
    Set TBCONSULTA3 = db.OpenRecordset("TBPRODUCTOS")
    TBCONSULTA3.Index = "Nro_proI"
    
    Set TBENCABEZADO = db.OpenRecordset("TBENCABEZADO")
    TBENCABEZADO.Index = "PrimaryKey"
    
    CargaGrilla
    MostrarGrilla
    CargaList
    CargaListProd
    
    TBENCABEZADO.MoveLast
    txtNum.Text = TBENCABEZADO!Nro_Presupuesto + 1
    
    maskFecha.Mask = Format(Now, "dd/mm/yyyy")
    maskVence.Mask = Format(Now, "dd/mm/yyyy")
End Sub

Sub CargaListProd()
    TBCONSULTA2.MoveFirst
    Do Until TBCONSULTA2.EOF
        If i > 20 Then
            Exit Sub
        Else
            ListProd.AddItem Format(TBCONSULTA2!Nro_pro, "0000") & " - " & TBCONSULTA2!Nom_pro
        End If
        i = i + 1
        TBCONSULTA2.MoveNext
    Loop
End Sub

Sub LIMPIAR()
    txtNro.Text = ""
    txtNom.Text = ""
    txtDirec.Text = ""
    cmbIVA.Text = ""
    txtCP.Text = ""
    txtLoc.Text = ""
    maskCUIT.Mask = ""
    maskCUIT.Text = ""
    txtProv.Text = ""
    'maskFecha.Mask = Format(maskFecha.Mask, "##/##/####")
    maskFecha.Mask = Format(Now, "dd/mm/yyyy")
    cmbTipo.Text = ""
    txtNum.Text = ""
    'maskVence.Mask = Format(maskVence.Mask, "##/##/####")
    maskVence.Mask = Format(Now, "dd/mm/yyyy")
    cmbCV.Text = ""
    txtDCod.Text = ""
    txtDetalle.Text = ""
    txtDIVA.Text = ""
    txtDCant.Text = ""
    txtDPrecio.Text = ""
    txtDDesc.Text = ""
    txtDIP.Text = ""
    cmbTipo.Text = "Seleccionar tipo"
    cmbCV.Text = "Seleccionar venta"
    txtNro.SetFocus
    txtNom.Enabled = False
    txtDirec.Enabled = False
    cmbIVA.Enabled = False
    txtCP.Enabled = False
    txtLoc.Enabled = False
    maskCUIT.Enabled = False
    txtProv.Enabled = False
    maskFecha.Enabled = False
    cmbTipo.Enabled = False
    txtNum.Enabled = False
    maskVence.Enabled = False
    cmbCV.Enabled = False
    txtDCod.Enabled = False
    txtDetalle.Enabled = False
    txtDIVA.Enabled = False
    txtDCant.Enabled = False
    txtDPrecio.Enabled = False
    txtDDesc.Enabled = False
    txtDIP.Enabled = False
    btnAceptar.Enabled = False
    btnCancelar.Enabled = False
    btnBorrar.Enabled = False
    ListProd.Visible = False
    lblPro.Visible = False
    lblCli.Visible = True
    TBDETALLE.MoveLast
    txtNum.Text = TBDETALLE!NroPresu_Det + 1
End Sub

Private Sub grilla1_DblClick()
    btnBorrar.Enabled = True
    grilla1.Col = 0
    Z = grilla1.RowSel
    C = grilla1.Text
    TBDETALLE.Seek "=", Val(C)

With grilla1
    .Col = 0
    .RowSel = .Row
    .ColSel = .Cols - 1
End With
End Sub

Sub CargaList()
    TBCONSULTA.MoveFirst
    Do Until TBCONSULTA.EOF
        If i > 20 Then
            Exit Sub
        Else
            ListPresu.AddItem Format(TBCONSULTA!Nro_cli, "0000") & " - " & TBCONSULTA!AyN_cli
        End If
        i = i + 1
        TBCONSULTA.MoveNext
    Loop
End Sub

Sub CargaGrilla()
    grilla1.Text = ""
    grilla1.Cols = 7
    grilla1.Rows = 15
    grilla1.ColWidth(0) = 1250
    grilla1.ColWidth(1) = 8880
    grilla1.ColWidth(2) = 940
    grilla1.ColWidth(3) = 1070
    grilla1.ColWidth(4) = 1600
    grilla1.ColWidth(5) = 1600
    grilla1.ColWidth(6) = 1820
    grilla1.FixedAlignment(0) = 2
    grilla1.Col = 0
    grilla1.Row = 0
    grilla1.Text = "Cód.Producto"
    grilla1.Col = 1
    grilla1.FixedAlignment(1) = 2
    grilla1.Text = "Detalle del Bien"
    grilla1.Col = 2
    grilla1.FixedAlignment(2) = 2
    grilla1.Text = "%IVA"
    grilla1.Col = 3
    grilla1.FixedAlignment(3) = 2
    grilla1.Text = "Cantidad"
    grilla1.Col = 4
    grilla1.FixedAlignment(4) = 2
    grilla1.Text = "Prec.Unitario"
    grilla1.Col = 5
    grilla1.FixedAlignment(6) = 2
    grilla1.Text = "%Descuento "
    grilla1.Col = 6
    grilla1.Text = "Importe Parcial"
    grilla1.Row = 1
End Sub

Sub MostrarGrilla()
    grilla1.Col = 0
    grilla1.Row = 1

Set TBDETALLE = db.OpenRecordset("TBDETALLE")
TBDETALLE.Index = "PrimaryKey"

Dim i As Integer
i = 0

Do While Not TBDETALLE.EOF
    grilla1.Col = 0
    grilla1.Text = Format(TBDETALLE!Id, "0000000")
    i = TBDETALLE!CodProd_Det
    Set TBCONSULTA3 = db.OpenRecordset("SELECT TBPRODUCTOS.Nro_pro, TBPRODUCTOS.Nom_pro, TBDETALLE.CodProd_Det FROM TBPRODUCTOS, TBDETALLE WHERE TBPRODUCTOS.Nro_pro = TBDETALLE.CodProd_Det AND TBDETALLE.CodProd_Det = " & i & "")
    Producto = TBCONSULTA3!Nom_pro
    grilla1.Col = 1
    grilla1.Text = Producto
    grilla1.Col = 2
    grilla1.Text = Format("%" + TBDETALLE!IVA_Det)
    grilla1.Col = 3
    grilla1.Text = TBDETALLE!Cant_Det
    grilla1.Col = 4
    grilla1.Text = Format(TBDETALLE!PU_Det, "$######0.00")
    grilla1.Col = 5
    grilla1.Text = Format("%" + CStr(TBDETALLE!Desc_Det))
    grilla1.Col = 6
    grilla1.Text = Format(TBDETALLE!ImpParc_Det, "$######0.00")
    i = i + 1
    'If TBDETALLE!VENPRODUIVA = 0 Then
     '   EXENTO = EXENTO + TBDETALLE!VENPRODUPARC1
      '  Else
       ' CALCULOPREVIO = Val(TBDETALLE!VENPRODUPARC1) / ((100 + Val(RG227!VENPRODUIVA)) / 100)
       ' CALCULOPREVIO1 = Val(TBDETALLE!VENPRODUPARC1) - CALCULOPREVIO
       ' If TBDETALLE!VENPRODUIVA < 13 Then
        '    GRAVADO10 = GRAVADO10 + CALCULOPREVIO
         '   IVA10 = IVA10 + CALCULOPREVIO1
         ' Else
          '  GRAVADO21 = GRAVADO21 + CALCULOPREVIO
           ' IVA21 = IVA21 + CALCULOPREVIO1
       ' End If
       ' If Combo2.Text = "NO INSCRIPTO" Then
       '     NOINSC = ((IVA10 + IVA21) * 0.5)
        '    Else
         '   NOINSC = 0
        'End If
    'End If
    'TOTAL = EXENTO + GRAVADO21 + GRAVADO10 + IVA21 + IVA10 + NOINSC
    grilla1.Row = grilla1.Row + 1
    TBDETALLE.MoveNext
Loop
'Text12.Text = ""
'Text13.Text = ""
'Text14.Text = ""
'Text15.Text = ""
'Text16.Text = ""
'Text17.Text = ""
'Text18.Text = ""
'TOTALES
End Sub


Private Sub btnAceptar_Click()
    msg = MsgBox("¿Desea guadar presupuesto?", vbYesNo + vbQuestion, "Guardar presupuesto")
    If msg = 6 Then
        If E <> 1 Then
            TBENCABEZADO.AddNew
            TBDETALLE.AddNew
            
            TBENCABEZADO!Nro_Presupuesto = txtNum.Text
            TBENCABEZADO!Cli_Enca = txtNom.Text
            TBENCABEZADO!Fecha_Enca = maskFecha.Text
            TBENCABEZADO!TipoPresu_Enca = cmbTipo.Text
            TBENCABEZADO!FechaVto_Enca = maskVence.Text
            TBENCABEZADO!CondVenta_Enca = cmbCV.Text
            
            TBDETALLE!NroPresu_Det = txtNum.Text
            TBDETALLE!CodProd_Det = txtDCod.Text
            TBDETALLE!IVA_Det = txtDIVA.Text
            TBDETALLE!Cant_Det = txtDCant.Text
            TBDETALLE!PU_Det = txtDPrecio.Text
            TBDETALLE!Desc_Det = txtDDesc.Text
            TBDETALLE!ImpParc_Det = txtDIP.Text
        Else
            TBENCABEZADO!Nro_Presupuesto = txtNum.Text
            TBENCABEZADO!Cli_Enca = txtNom.Text
            TBENCABEZADO!Fecha_Enca = maskFecha.Text
            TBENCABEZADO!TipoPresu_Enca = cmbTipo.Text
            TBENCABEZADO!FechaVto_Enca = maskVence.Text
            TBENCABEZADO!CondVenta_Enca = cmbCV.Text
            
            TBDETALLE!NroPresu_Det = txtNum.Text
            TBDETALLE!CodProd_Det = txtDCod.Text
            TBDETALLE!IVA_Det = txtDIVA.Text
            TBDETALLE!Cant_Det = txtDCant.Text
            TBDETALLE!PU_Det = txtDPrecio.Text
            TBDETALLE!Desc_Det = txtDDesc.Text
            TBDETALLE!ImpParc_Det = txtDIP.Text
        End If
        TBENCABEZADO.Update
        TBDETALLE.Update
        btnCancelar.Enabled = False
        btnAceptar.Enabled = False
        LIMPIAR
    Else
        LIMPIAR
    End If
    TBDETALLE.MoveLast
    TBENCABEZADO.MoveLast
    txtNum.Text = TBENCABEZADO!Nro_Presupuesto + 1
    CargaGrilla
    MostrarGrilla
End Sub

Private Sub btnBorrar_Click()
    msg = MsgBox("¿Desea eliminar los datos?", vbYesNo + vbDefaultButton2 + vbQuestion, "Eliminar datos")
    If msg = 6 Then
        TBDETALLE.Delete
        TBENCABEZADO.Seek "=", TBDETALLE!NroPresu_Det
        TBENCABEZADO.Delete
        LIMPIAR
        CargaGrilla
        MostrarGrilla
        grilla1.RemoveItem (grilla1.Row)
    Else
        
    End If
    TBENCABEZADO.MoveLast
    txtNum.Text = TBENCABEZADO!Nro_Presupuesto + 1
    LIMPIAR
End Sub

Private Sub btnCancelar_Click()
    LIMPIAR
    TBENCABEZADO.MoveLast
    txtNum.Text = TBENCABEZADO!Nro_Presupuesto + 1
End Sub

Function ValidarFecha(qFecha) As Boolean
    If Not IsDate(qFecha) Then
        ValidarFecha = False
        Exit Function
    End If
    Dim mes As Integer
    mes = Mid(qFecha, 4, 2)
    If Val(mes) > 12 Then
        ValidarFecha = False
        Exit Function
    End If
    ValidarFecha = True
End Function

Private Sub ListProd_DblClick()
    Z = Mid(Val(ListProd.Text), 1, 4)
    C = Val(Z)
    TBPRODUCTOS.Seek "=", Val(C)
    If TBPRODUCTOS.NoMatch Then
        cmbCV.SetFocus
        msg = MsgBox("No se ha encontrado el producto", vbExclamation, "Error")
        btnCancelar.Enabled = True
        E = 0
    Else
        TBPRODUCTOS.Edit
        txtDCod.Text = TBPRODUCTOS!Nro_pro
        txtDetalle.Text = TBPRODUCTOS!Nom_pro
        txtDCant.Text = TBPRODUCTOS!Cantidad_pro
        txtDPrecio.Text = TBPRODUCTOS!PrecioVenta_pro
        txtDIVA.Text = TBPRODUCTOS!IVA_pro
        E = 1
        txtDCant.Enabled = True
        txtDCant.SetFocus
    End If
End Sub

Private Sub ListPresu_DblClick()
    Z = Mid(Val(ListPresu.Text), 1, 4)
    C = Val(Z)
    TBCLIENTES.Seek "=", Val(C)
    If TBCLIENTES.NoMatch Then
        txtNro.Enabled = True
        txtNro.SetFocus
        msg = MsgBox("No se ha encontrado el cliente", vbExclamation, "Error")
        btnCancelar.Enabled = True
        E = 0
    Else
        TBCLIENTES.Edit
        txtNro.Text = TBCLIENTES!Nro_cli
        txtNom.Text = TBCLIENTES!AyN_cli
        txtDirec.Text = TBCLIENTES!Dom_cli
        cmbIVA.Text = TBCLIENTES!Siva_cli
        Select Case cmbIVA.Text
            Case 1
                cmbIVA.Text = "1 - Resp. Inscripto"
            Case 2
                cmbIVA.Text = "2 - Monotributo"
            Case 3
                cmbIVA.Text = "3 - Exento"
            Case 4
                cmbIVA.Text = "4 - Consumidor final"
        End Select
        txtCP.Text = TBCLIENTES!Cpos_cli
        txtLoc.Text = TBCODPOS!Loc_cpos
        maskCUIT.Mask = TBCLIENTES!Cuit_cli
        txtProv.Text = TBCODPOS!Pcia_cpos
        Select Case txtProv.Text
            Case 1
                txtProv.Text = "01 - Buenos Aires"
            Case 2
                txtProv.Text = "02 - Capital Federal"
            Case 3
                txtProv.Text = "03 - Catamarca"
            Case 4
                txtProv.Text = "04 - Chaco"
            Case 5
                txtProv.Text = "05 - Chubut"
            Case 6
                txtProv.Text = "06 - Córdoba"
            Case 7
                txtProv.Text = "07 - Corrientes"
            Case 8
                txtProv.Text = "08 - Entre Ríos"
            Case 9
                txtProv.Text = "09 - Formosa"
            Case 10
                txtProv.Text = "10 - Jujuy"
            Case 11
                txtProv.Text = "11 - La Pampa"
            Case 12
                txtProv.Text = "12 - La Rioja"
            Case 13
                txtProv.Text = "13 - Mendoza"
            Case 14
                txtProv.Text = "14 - Misiones"
            Case 15
                txtProv.Text = "15 - Neuquén"
            Case 16
                txtProv.Text = "16 - Río Negro"
            Case 17
                txtProv.Text = "17 - Salta"
            Case 18
                txtProv.Text = "18 - San Juan"
            Case 19
                txtProv.Text = "19 - San Luis"
            Case 20
                txtProv.Text = "20 - Santa Cruz"
            Case 21
                txtProv.Text = "21 - Santa Fe"
            Case 22
                txtProv.Text = "22 - Santiago del Estero"
            Case 23
                txtProv.Text = "23 - Tierra del Fuego"
            Case 24
                txtProv.Text = "24 - Tucumán"
        End Select
        E = 1
        maskFecha.Enabled = True
        maskFecha.SetFocus
        btnCancelar.Enabled = True
        txtNom.Enabled = True
        txtDirec.Enabled = True
        cmbIVA.Enabled = True
        txtCP.Enabled = True
        txtLoc.Enabled = True
        maskCUIT.Enabled = True
        txtProv.Enabled = True
    End If
End Sub

Private Sub btnSalir_Click()
    Unload Presupuesto
End Sub

Private Sub txtDDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtDDesc.Text) < 0 Or txtDDesc.Text = "" Then
            txtDDesc.SetFocus
        Else
            Desc = PrecioFinal * Val(txtDDesc.Text) / 100
            txtDIP.Text = PrecioFinal - Desc
            txtDIP.Enabled = False
            btnAceptar.Enabled = True
            btnAceptar.SetFocus
        End If
    End If
End Sub

Private Sub txtDCant_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtDCant.Text) <= 0 Or txtDCant.Text = "" Then
            txtDCant.SetFocus
        Else
            PrecioFinal = Val(txtDCant.Text) * Val(txtDPrecio.Text)
            txtDDesc.Enabled = True
            txtDDesc.SetFocus
        End If
    End If
End Sub

Private Sub cmbCV_Click()
    If cmbCV.Text = "" Then
        cmbCV.SetFocus
    Else
        ListProd.Visible = True
        lblCli.Visible = False
        lblPro.Visible = True
        ListProd.SetFocus
    End If
End Sub

Private Sub maskVence_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not ValidarFecha(maskVence.Text) Then
            MsgBox "La fecha informada no es válida", vbExclamation, Me.Caption
            maskVence.SelStart = 0
            maskVence.SelLength = Len(maskFecha.Text)
            maskVence.SetFocus
            Exit Sub
        End If
        cmbCV.Enabled = True
        cmbCV.SetFocus
    End If
End Sub

Private Sub txtDCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtDCod.Text) <= 0 Or txtDCod.Text = "" Then
            txtDCod.SetFocus
        Else
            TBPRODUCTOS.Seek "=", Val(txtDCod.Text)
            If TBPRODUCTOS.NoMatch Then
                txtDCod.Enabled = True
                txtDCod.SetFocus
                btnCancelar.Enabled = True
                E = 0
            Else
                TBPRODUCTOS.Edit
                txtDCod.Text = TBPRODUCTOS!Nro_pro
                txtDetalle.Text = TBPRODUCTOS!Nom_pro
                txtDIVA.Text = TBPRODUCTOS!IVA_pro
                txtDCant.Text = TBPRODUCTOS!Cantidad_pro
                txtDPrecio.Text = TBPRODUCTOS!Precio_pro
                txtDIP.Text = TBPRODUCTOS!PrecioVenta_pro
                E = 1
                btnCancelar.Enabled = True
                btnAceptar.Enabled = True
                btnBorrar.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub cmbTipo_Click()
    If cmbTipo.Text = "" Then
        cmbTipo.SetFocus
    Else
        maskVence.Enabled = True
        maskVence.SetFocus
    End If
End Sub

Private Sub maskFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not ValidarFecha(maskFecha.Text) Then
            MsgBox "La fecha informada no es válida", vbExclamation, Me.Caption
            maskFecha.SelStart = 0
            maskFecha.SelLength = Len(maskFecha.Text)
            maskFecha.SetFocus
            Exit Sub
        End If
        cmbTipo.Enabled = True
        cmbTipo.SetFocus
    End If
End Sub

Private Sub txtNro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtNro.Text) <= 0 Or txtNro.Text = "" Then
            txtNro.SetFocus
        Else
            TBCLIENTES.Seek "=", Val(txtNro.Text)
            If TBCLIENTES.NoMatch Then
                txtNom.Enabled = True
                txtNom.SetFocus
                btnCancelar.Enabled = True
                E = 0
            Else
                TBCLIENTES.Edit
                txtNom.Text = TBCLIENTES!AyN_cli
                txtDirec.Text = TBCLIENTES!Dom_cli
                cmbIVA.Text = TBCLIENTES!Siva_cli
                Select Case cmbIVA.Text
                    Case 1
                        cmbIVA.Text = "1 - Resp. Inscripto"
                    Case 2
                        cmbIVA.Text = "2 - Monotributo"
                    Case 3
                        cmbIVA.Text = "3 - Exento"
                    Case 4
                        cmbIVA.Text = "4 - Consumidor final"
                End Select
                txtCP.Text = TBCLIENTES!Cpos_cli
                txtLoc.Text = TBCODPOS!Loc_cpos
                maskCUIT.Mask = TBCLIENTES!Cuit_cli
                txtProv.Text = TBCODPOS!Pcia_cpos
                Select Case txtProv.Text
                    Case 1
                        txtProv.Text = "01 - Buenos Aires"
                    Case 2
                        txtProv.Text = "02 - Capital Federal"
                    Case 3
                        txtProv.Text = "03 - Catamarca"
                    Case 4
                        txtProv.Text = "04 - Chaco"
                    Case 5
                        txtProv.Text = "05 - Chubut"
                    Case 6
                        txtProv.Text = "06 - Córdoba"
                    Case 7
                        txtProv.Text = "07 - Corrientes"
                    Case 8
                        txtProv.Text = "08 - Entre Ríos"
                    Case 9
                        txtProv.Text = "09 - Formosa"
                    Case 10
                        txtProv.Text = "10 - Jujuy"
                    Case 11
                        txtProv.Text = "11 - La Pampa"
                    Case 12
                        txtProv.Text = "12 - La Rioja"
                    Case 13
                        txtProv.Text = "13 - Mendoza"
                    Case 14
                        txtProv.Text = "14 - Misiones"
                    Case 15
                        txtProv.Text = "15 - Neuquén"
                    Case 16
                        txtProv.Text = "16 - Río Negro"
                    Case 17
                        txtProv.Text = "17 - Salta"
                    Case 18
                        txtProv.Text = "18 - San Juan"
                    Case 19
                        txtProv.Text = "19 - San Luis"
                    Case 20
                        txtProv.Text = "20 - Santa Cruz"
                    Case 21
                        txtProv.Text = "21 - Santa Fe"
                    Case 22
                        txtProv.Text = "22 - Santiago del Estero"
                    Case 23
                        txtProv.Text = "23 - Tierra del Fuego"
                    Case 24
                        txtProv.Text = "24 - Tucumán"
                End Select
                E = 1
                maskFecha.Enabled = True
                maskFecha.SetFocus
                btnCancelar.Enabled = True
                btnAceptar.Enabled = True
                btnBorrar.Enabled = True
                txtNom.Enabled = True
                txtDirec.Enabled = True
                cmbIVA.Enabled = True
                txtCP.Enabled = True
                txtLoc.Enabled = True
                maskCUIT.Enabled = True
                txtProv.Enabled = True
            End If
        End If
    End If
End Sub
