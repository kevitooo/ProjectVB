VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Clientes 
   BackColor       =   &H80000003&
   Caption         =   "Clientes"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15330
   Icon            =   "Clientes.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   15330
   Begin MSMask.MaskEdBox mbCUIT 
      Height          =   375
      Left            =   3240
      TabIndex        =   31
      Top             =   6720
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
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
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtBuscar 
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
      Left            =   9240
      TabIndex        =   30
      Top             =   720
      Width           =   4455
   End
   Begin VB.TextBox txtCT 
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
      Height          =   360
      Left            =   3240
      MaxLength       =   5
      TabIndex        =   26
      Top             =   4200
      Width           =   975
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
      Height          =   360
      Left            =   3240
      MaxLength       =   25
      TabIndex        =   25
      Top             =   2880
      Width           =   3135
   End
   Begin VB.ComboBox cmbProv 
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
      Left            =   3240
      TabIndex        =   24
      Text            =   "Seleccionar provincia"
      Top             =   3480
      Width           =   2655
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
      Left            =   9960
      TabIndex        =   12
      Top             =   9840
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
      Left            =   7800
      TabIndex        =   11
      Top             =   9840
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
      Left            =   5640
      TabIndex        =   10
      Top             =   9840
      Width           =   1815
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
      Left            =   3480
      TabIndex        =   9
      Top             =   9840
      Width           =   1815
   End
   Begin VB.ListBox ListCli 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6360
      ItemData        =   "Clientes.frx":1084A
      Left            =   8760
      List            =   "Clientes.frx":1084C
      TabIndex        =   13
      Top             =   1320
      Width           =   5055
   End
   Begin VB.TextBox txtObs 
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
      Height          =   1155
      Left            =   3240
      MaxLength       =   25
      TabIndex        =   8
      Top             =   8160
      Width           =   3135
   End
   Begin VB.ComboBox cmbCat 
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
      Left            =   3240
      TabIndex        =   7
      Text            =   "Seleccionar categoría"
      Top             =   7440
      Width           =   2775
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
      Left            =   3240
      TabIndex        =   6
      Text            =   "Seleccionar IVA"
      Top             =   6000
      Width           =   2175
   End
   Begin VB.TextBox txtMail 
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
      Height          =   360
      Left            =   3240
      MaxLength       =   25
      TabIndex        =   5
      Top             =   5400
      Width           =   3135
   End
   Begin VB.TextBox txtTel 
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
      Height          =   360
      Left            =   3240
      MaxLength       =   15
      TabIndex        =   4
      Top             =   4800
      Width           =   2175
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
      Height          =   360
      Left            =   3240
      MaxLength       =   4
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtDom 
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
      Height          =   360
      Left            =   3240
      MaxLength       =   25
      TabIndex        =   2
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox txtAyP 
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
      Height          =   360
      Left            =   3240
      MaxLength       =   25
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox txtNroC 
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
      Height          =   360
      Left            =   3240
      MaxLength       =   4
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   8880
      Picture         =   "Clientes.frx":1084E
      Stretch         =   -1  'True
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000003&
      Caption         =   "Característica telefónica:"
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
      Left            =   480
      TabIndex        =   29
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Label Label11 
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
      Left            =   2160
      TabIndex        =   28
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Localidad 
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
      Left            =   2040
      TabIndex        =   27
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000003&
      Caption         =   "Observaciones:"
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
      Left            =   1560
      TabIndex        =   23
      Top             =   8280
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000003&
      Caption         =   "Categoría:"
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
      Left            =   2040
      TabIndex        =   22
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000003&
      Caption         =   "CUIT:"
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
      Left            =   2520
      TabIndex        =   21
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000003&
      Caption         =   "Situación IVA:"
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
      Left            =   1680
      TabIndex        =   20
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000003&
      Caption         =   "Mail:"
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
      Left            =   2640
      TabIndex        =   19
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000003&
      Caption         =   "Teléfono:"
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
      Left            =   2160
      TabIndex        =   18
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000003&
      Caption         =   "Código postal:"
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
      TabIndex        =   17
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000003&
      Caption         =   "Domicilio:"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000003&
      Caption         =   "Apellido y nombre:"
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
      Left            =   1200
      TabIndex        =   15
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      Caption         =   "Número cliente:"
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
      Left            =   1440
      TabIndex        =   14
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim E As Integer
Dim X As Integer

Private Sub Form_Load()
    Clientes.Height = 11520
    Clientes.Width = 15570
    Clientes.Top = ((Screen.Height - Me.Height) / 2) - 650
    Clientes.Left = ((Screen.Width - Me.Width) / 2)
    
    cmbProv.AddItem "01 - Buenos Aires"
    cmbProv.AddItem "02 - Capital Federal"
    cmbProv.AddItem "03 - Catamarca"
    cmbProv.AddItem "04 - Chaco"
    cmbProv.AddItem "05 - Chubut"
    cmbProv.AddItem "06 - Córdoba"
    cmbProv.AddItem "07 - Corrientes"
    cmbProv.AddItem "08 - Entre Ríos"
    cmbProv.AddItem "09 - Formosa"
    cmbProv.AddItem "10 - Jujuy"
    cmbProv.AddItem "11 - La Pampa"
    cmbProv.AddItem "12 - La Rioja"
    cmbProv.AddItem "13 - Mendoza"
    cmbProv.AddItem "14 - Misiones"
    cmbProv.AddItem "15 - Neuquén"
    cmbProv.AddItem "16 - Río Negro"
    cmbProv.AddItem "17 - Salta"
    cmbProv.AddItem "18 - San Juan"
    cmbProv.AddItem "19 - San Luis"
    cmbProv.AddItem "20 - Santa Cruz"
    cmbProv.AddItem "21 - Santa Fe"
    cmbProv.AddItem "22 - Santiago del Estero"
    cmbProv.AddItem "23 - Tierra del Fuego"
    cmbProv.AddItem "24 - Tucumán"
    
    cmbIVA.AddItem "1 - Resp. Inscripto"
    cmbIVA.AddItem "2 - Monotributo"
    cmbIVA.AddItem "3 - Exento"
    cmbIVA.AddItem "4 - Consumidor final"
    
    cmbCat.AddItem "1 - Mayorista"
    cmbCat.AddItem "2 - Minorista"
    cmbCat.AddItem "3 - No categorizado"
    
    camino = App.Path
    Set db = OpenDatabase(camino & "\dbInfor.mdb")
    
    Set TBCLIENTES = db.OpenRecordset("TBCLIENTES")
    TBCLIENTES.Index = "NroI_cli"
    
    Set TBCODPOS = db.OpenRecordset("TBCODPOS")
    TBCODPOS.Index = "NroI_cpos"
    
    Set TBCONSULTA = db.OpenRecordset("TBCLIENTES")
    TBCONSULTA.Index = "AyNI_cli"
    
    CargaList
    
    TBCLIENTES.MoveLast
    txtNroC.Text = TBCLIENTES!Nro_cli + 1
End Sub

Sub CargaList()
    TBCONSULTA.MoveFirst
    Do Until TBCONSULTA.EOF
        If i > 20 Then
            Exit Sub
        Else
            ListCli.AddItem Format(TBCONSULTA!Nro_cli, "0000") & " - " & TBCONSULTA!AyN_cli
        End If
        i = i + 1
        TBCONSULTA.MoveNext
    Loop
End Sub

Sub BuscaCodigo()
    TBCODPOS.Seek "=", Val(txtCP.Text)
    If TBCODPOS.NoMatch Then
        txtLoc.Enabled = True
        txtLoc.SetFocus
        btnCancelar.Enabled = True
        X = 0
    Else
        TBCODPOS.Edit
        txtLoc.Text = TBCODPOS!Loc_cpos
        cmbProv.Text = TBCODPOS!Pcia_cpos
        Select Case cmbProv.Text
            Case 1
                cmbProv.Text = "01 - Buenos Aires"
            Case 2
                cmbProv.Text = "02 - Capital Federal"
            Case 3
                cmbProv.Text = "03 - Catamarca"
            Case 4
                cmbProv.Text = "04 - Chaco"
            Case 5
                cmbProv.Text = "05 - Chubut"
            Case 6
                cmbProv.Text = "06 - Córdoba"
            Case 7
                cmbProv.Text = "07 - Corrientes"
            Case 8
                cmbProv.Text = "08 - Entre Ríos"
            Case 9
                cmbProv.Text = "09 - Formosa"
            Case 10
                cmbProv.Text = "10 - Jujuy"
            Case 11
                cmbProv.Text = "11 - La Pampa"
            Case 12
                cmbProv.Text = "12 - La Rioja"
            Case 13
                cmbProv.Text = "13 - Mendoza"
            Case 14
                cmbProv.Text = "14 - Misiones"
            Case 15
                cmbProv.Text = "15 - Neuquén"
            Case 16
                cmbProv.Text = "16 - Río Negro"
            Case 17
                cmbProv.Text = "17 - Salta"
            Case 18
                cmbProv.Text = "18 - San Juan"
            Case 19
                cmbProv.Text = "19 - San Luis"
            Case 20
                cmbProv.Text = "20 - Santa Cruz"
            Case 21
                cmbProv.Text = "21 - Santa Fe"
            Case 22
                cmbProv.Text = "22 - Santiago del Estero"
            Case 23
                cmbProv.Text = "23 - Tierra del Fuego"
            Case 24
                cmbProv.Text = "24 - Tucumán"
        End Select
        txtCT.Text = TBCODPOS!Car_cpos
        X = 1
        txtLoc.Enabled = True
        cmbProv.Enabled = True
        txtCT.Enabled = True
        txtTel.Enabled = True
        txtTel.SetFocus
    End If
End Sub

Sub LIMPIAR()
    txtNroC.Text = ""
    txtAyP.Text = ""
    txtDom.Text = ""
    txtCP.Text = ""
    txtLoc.Text = ""
    cmbProv.Text = ""
    txtCT.Text = ""
    txtTel.Text = ""
    txtMail.Text = ""
    cmbIVA.Text = "Seleccionar IVA"
    cmbIVA.Enabled = False
    mbCUIT.Mask = ""
    mbCUIT.Text = ""
    mbCUIT.Enabled = False
    cmbCat.Text = "Seleccionar categoria"
    cmbCat.Enabled = False
    txtObs.Text = ""
    txtNroC.SetFocus
    txtAyP.Enabled = False
    txtDom.Enabled = False
    txtCP.Enabled = False
    txtLoc.Enabled = False
    cmbProv.Enabled = False
    txtCP.Enabled = False
    txtTel.Enabled = False
    txtMail.Enabled = False
    txtObs.Enabled = False
    btnAceptar.Enabled = False
    btnCancelar.Enabled = False
    btnBorrar.Enabled = False
End Sub

Private Sub btnSalir_Click()
    Unload Clientes
End Sub

Private Sub btnCancelar_Click()
    LIMPIAR
    TBCLIENTES.MoveLast
    txtNroC.Text = TBCLIENTES!Nro_cli + 1
End Sub

Private Sub btnBorrar_Click()
    msg = MsgBox("¿Desea eliminar los datos?", vbYesNo + vbDefaultButton2 + vbQuestion, "Eliminar archivo")
    If msg = 6 Then
        TBCLIENTES.Delete
        ListCli.RemoveItem (ListCli.ListIndex)
        LIMPIAR
    Else

    End If
    TBCLIENTES.MoveLast
    txtNroC.Text = TBCLIENTES!Nro_cli + 1
End Sub

Private Sub btnAceptar_Click()
    msg = MsgBox("¿Desea guadar los datos?", vbYesNo + vbQuestion, "Guardar archivo")
    If msg = 6 Then
        If E <> 1 Then
            TBCLIENTES.AddNew
            TBCLIENTES!Nro_cli = txtNroC.Text
            TBCLIENTES!AyN_cli = txtAyP.Text
            TBCLIENTES!Dom_cli = txtDom.Text
            TBCLIENTES!Cpos_cli = txtCP.Text
            TBCLIENTES!Tel_cli = txtTel.Text
            TBCLIENTES!Mail_cli = txtMail.Text
            TBCLIENTES!Siva_cli = Mid(Val(cmbIVA.Text), 1, 1)
            TBCLIENTES!Cuit_cli = Val(mbCUIT.Text)
            TBCLIENTES!Cat_cli = Mid(Val(cmbCat.Text), 1, 1)
            TBCLIENTES!Obs_cli = txtObs.Text
            ListCli.AddItem Format(TBCLIENTES!Nro_cli, "0000") & " - " & TBCLIENTES!AyN_cli
        Else
            TBCLIENTES!Nro_cli = txtNroC.Text
            TBCLIENTES!AyN_cli = txtAyP.Text
            TBCLIENTES!Dom_cli = txtDom.Text
            TBCLIENTES!Cpos_cli = txtCP.Text
            TBCLIENTES!Tel_cli = txtTel.Text
            TBCLIENTES!Mail_cli = txtMail.Text
            TBCLIENTES!Siva_cli = Mid(Val(cmbIVA.Text), 1, 1)
            TBCLIENTES!Cuit_cli = Val(mbCUIT.Text)
            TBCLIENTES!Cat_cli = Mid(Val(cmbCat.Text), 1, 1)
            TBCLIENTES!Obs_cli = txtObs.Text
        End If
        If X <> 1 Then
            TBCODPOS.AddNew
            TBCODPOS!Nro_cpos = txtCP.Text
            TBCODPOS!Loc_cpos = txtLoc.Text
            TBCODPOS!Pcia_cpos = Mid(Val(cmbProv.Text), 1, 2)
            TBCODPOS!Car_cpos = txtCT.Text
        Else
            TBCODPOS!Nro_cpos = txtCP.Text
            TBCODPOS!Loc_cpos = txtLoc.Text
            TBCODPOS!Pcia_cpos = Mid(Val(cmbProv.Text), 1, 2)
            TBCODPOS!Car_cpos = txtCT.Text
        End If
        TBCLIENTES.Update
        TBCODPOS.Update
        btnCancelar.Enabled = False
        btnAceptar.Enabled = False
        LIMPIAR
    Else
        LIMPIAR
    End If
    TBCLIENTES.MoveLast
    txtNroC.Text = TBCLIENTES!Nro_cli + 1
End Sub

Private Sub ListCli_DblClick()
    Z = Mid(Val(ListCli.Text), 1, 4)
    C = Val(Z)
    TBCLIENTES.Seek "=", Val(C)
    If TBCLIENTES.NoMatch Then
        txtAyP.Enabled = True
        txtAyP.SetFocus
        btnCancelar.Enabled = True
        E = 0
    Else
        TBCLIENTES.Edit
        txtNroC.Text = TBCLIENTES!Nro_cli
        txtAyP.Text = TBCLIENTES!AyN_cli
        txtDom.Text = TBCLIENTES!Dom_cli
        txtCP.Text = TBCLIENTES!Cpos_cli
        BuscaCodigo
        txtTel.Text = TBCLIENTES!Tel_cli
        txtMail.Text = TBCLIENTES!Mail_cli
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
        mbCUIT.Mask = TBCLIENTES!Cuit_cli
        cmbCat.Text = TBCLIENTES!Cat_cli
        Select Case cmbCat.Text
            Case 1
                cmbCat.Text = "1 - Mayorista"
            Case 2
                cmbCat.Text = "2 - Minorista"
            Case 3
                cmbCat.Text = "3 - No categorizado"
        End Select
        txtObs.Text = TBCLIENTES!Obs_cli
        E = 1
        btnCancelar.Enabled = True
        btnAceptar.Enabled = True
        btnBorrar.Enabled = True
        txtAyP.Enabled = True
        txtDom.Enabled = True
        txtCP.Enabled = True
        txtLoc.Enabled = True
        cmbProv.Enabled = True
        txtCT.Enabled = True
        txtTel.Enabled = True
        txtMail.Enabled = True
        cmbIVA.Enabled = True
        mbCUIT.Enabled = True
        cmbCat.Enabled = True
        txtObs.Enabled = True
        txtNroC.SetFocus
    End If
End Sub

Private Sub txtBuscar_Change()
    ListCli.Clear
    Set TBCONSULTA = db.OpenRecordset("select * from TBCLIENTES where AyN_cli >= '" & txtBuscar.Text & "' order by AyN_cli, Nro_cli")
        TBCONSULTA.MoveFirst
    If Not TBCONSULTA.EOF Then
        TBCONSULTA.MoveFirst
        Do While Not TBCONSULTA.EOF
            ListCli.AddItem Format(TBCONSULTA!Nro_cli, "0000") & " - " & TBCONSULTA!AyN_cli
            TBCONSULTA.MoveNext
        Loop
    End If
End Sub

Private Sub txtObs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtObs.Text = "" Then
            txtObs.SetFocus
        Else
            btnAceptar.Enabled = True
            btnAceptar.SetFocus
        End If
    End If
End Sub

Private Sub cmbCat_Click()
    If IsNumeric(cmbCat.Text) Or cmbCat.Text = "" Then
        cmbCat.SetFocus
    Else
        txtObs.Enabled = True
        txtObs.SetFocus
    End If
End Sub

Private Sub mbCUIT_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(mbCUIT.Text) <= 0 Or mbCUIT.Text = "" Then
            mbCUIT.SetFocus
        Else
            cmbCat.Enabled = True
            cmbCat.SetFocus
        End If
    End If
End Sub

Private Sub cmbIVA_Click()
    If Val(cmbIVA.Text) = 0 Then
        cmbIVA.SetFocus
    Else
        mbCUIT.Enabled = True
        mbCUIT.SetFocus
    End If
End Sub

Private Sub txtMail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtMail.Text = "" Then
            txtMail.SetFocus
        Else
            cmbIVA.Enabled = True
            cmbIVA.SetFocus
        End If
    End If
End Sub

Private Sub txtTel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtTel.Text) <= 0 Or txtTel.Text = "" Then
            txtTel.SetFocus
        Else
            txtMail.Enabled = True
            txtMail.SetFocus
        End If
    End If
End Sub

Private Sub txtCT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtCT.Text) <= 0 Or txtCT.Text = "" Then
            txtCT.SetFocus
        Else
            txtTel.Enabled = True
            txtTel.SetFocus
        End If
    End If
End Sub

Private Sub cmbProv_Click()
    If Val(cmbProv.Text) = 0 Then
        cmbProv.SetFocus
    Else
        txtCT.Enabled = True
        txtCT.SetFocus
    End If
End Sub

Private Sub txtLoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txtLoc.Text) Or txtLoc.Text = "" Then
            txtLoc.SetFocus
        Else
            txtLoc.Text = UCase(txtLoc.Text)
            cmbProv.Enabled = True
            cmbProv.SetFocus
        End If
    End If
End Sub

Private Sub txtCP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtCP.Text) < 1000 Then
            txtCP.SetFocus
        Else
            BuscaCodigo
            If txtLoc.Text = "" Then
                txtLoc.Enabled = True
                txtLoc.SetFocus
            Else
                
            End If
        End If
    End If
End Sub

Private Sub txtDom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txtDom.Text) Or txtDom.Text = "" Then
            txtDom.SetFocus
        Else
            txtDom.Text = UCase(txtDom.Text)
            txtCP.Enabled = True
            txtCP.SetFocus
        End If
    End If
End Sub

Private Sub txtAyP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txtAyP.Text) Or txtAyP.Text = "" Then
            txtAyP.SetFocus
        Else
            txtAyP.Text = UCase(txtAyP.Text)
            txtDom.Enabled = True
            txtDom.SetFocus
        End If
    End If
End Sub

Private Sub txtNroC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtNroC.Text) <= 0 Or txtNroC.Text = "" Then
            txtNroC.SetFocus
        Else
            TBCLIENTES.Seek "=", Val(txtNroC.Text)
            If TBCLIENTES.NoMatch Then
                txtAyP.Enabled = True
                txtAyP.SetFocus
                btnCancelar.Enabled = True
                E = 0
            Else
                TBCLIENTES.Edit
                txtAyP.Text = TBCLIENTES!AyN_cli
                txtDom.Text = TBCLIENTES!Dom_cli
                txtCP.Text = TBCLIENTES!Cpos_cli
                BuscaCodigo
                txtTel.Text = TBCLIENTES!Tel_cli
                txtMail.Text = TBCLIENTES!Mail_cli
                cmbIVA.Text = TBCLIENTES!Siva_cli
                Select Case cmbIVA.Text
                    Case 1
                        cmbIVA.Text = "Resp. Inscripto"
                    Case 2
                        cmbIVA.Text = "Monotributo"
                    Case 3
                        cmbIVA.Text = "Exento"
                    Case 4
                        cmbIVA.Text = "Consumidor final"
                End Select
                mbCUIT.Text = TBCLIENTES!Cuit_cli
                cmbCat.Text = TBCLIENTES!Cat_cli
                Select Case cmbCat.Text
                    Case 1
                        cmbCat.Text = "Mayorista"
                    Case 2
                        cmbCat.Text = "Minorista"
                    Case 3
                        cmbCat.Text = "No categorizado"
                End Select
                txtObs.Text = TBCLIENTES!Obs_cli
                E = 1
                btnCancelar.Enabled = True
                btnAceptar.Enabled = True
                btnBorrar.Enabled = True
                txtAyP.Enabled = True
                txtDom.Enabled = True
                txtCP.Enabled = True
                txtLoc.Enabled = True
                cmbProv.Enabled = True
                txtCT.Enabled = True
                txtTel.Enabled = True
                txtMail.Enabled = True
                cmbIVA.Enabled = True
                mbCUIT.Enabled = True
                cmbCat.Enabled = True
                txtObs.Enabled = True
            End If
        End If
    End If
End Sub
