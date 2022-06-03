VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Productos 
   BackColor       =   &H80000003&
   Caption         =   "Productos"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13200
   Icon            =   "Productos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   13200
   Begin VB.TextBox txtPVenta 
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
      Left            =   3360
      TabIndex        =   27
      Top             =   6480
      Width           =   1815
   End
   Begin MSMask.MaskEdBox maskFecha 
      Height          =   375
      Left            =   3360
      TabIndex        =   26
      Top             =   2880
      Width           =   1935
      _ExtentX        =   3413
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
      Left            =   3360
      TabIndex        =   25
      Text            =   "Seleccionar IVA"
      Top             =   5280
      Width           =   3135
   End
   Begin VB.ComboBox cmbUM 
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
      Left            =   3360
      TabIndex        =   24
      Text            =   "Seleccionar unidad"
      Top             =   3480
      Width           =   3135
   End
   Begin VB.ComboBox cmbSR 
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
      Left            =   3360
      TabIndex        =   23
      Text            =   "Seleccionar subrubro"
      Top             =   2280
      Width           =   3135
   End
   Begin VB.ComboBox cmbRub 
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
      Left            =   3360
      TabIndex        =   22
      Text            =   "Seleccionar rubro"
      Top             =   1680
      Width           =   3135
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
      Left            =   8880
      TabIndex        =   8
      Top             =   8040
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
      Left            =   6720
      TabIndex        =   7
      Top             =   8040
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
      Left            =   4560
      TabIndex        =   6
      Top             =   8040
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
      Left            =   2400
      TabIndex        =   5
      Top             =   8040
      Width           =   1815
   End
   Begin VB.ListBox ListPro 
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
      ItemData        =   "Productos.frx":1084A
      Left            =   7680
      List            =   "Productos.frx":1084C
      TabIndex        =   10
      Top             =   1080
      Width           =   5055
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
      Left            =   8160
      TabIndex        =   9
      Top             =   480
      Width           =   4455
   End
   Begin VB.TextBox txtPR 
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
      Left            =   3360
      MaxLength       =   25
      TabIndex        =   4
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtNroP 
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
      Left            =   3360
      MaxLength       =   4
      TabIndex        =   0
      Top             =   480
      Width           =   735
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
      Height          =   360
      Left            =   3360
      MaxLength       =   25
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox txtPrecio 
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
      Left            =   3360
      MaxLength       =   15
      TabIndex        =   3
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox txtCant 
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
      Left            =   3360
      MaxLength       =   5
      TabIndex        =   2
      Top             =   4080
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   7800
      Picture         =   "Productos.frx":1084E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      Caption         =   "Número producto:"
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
      Left            =   1320
      TabIndex        =   21
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000003&
      Caption         =   "Nombre de producto:"
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
      Left            =   960
      TabIndex        =   20
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000003&
      Caption         =   "Rubro de producto:"
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
      TabIndex        =   19
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000003&
      Caption         =   "Subrubro de producto:"
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
      Left            =   840
      TabIndex        =   18
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000003&
      Caption         =   "Precio:"
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
      TabIndex        =   17
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label6 
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
      Left            =   2760
      TabIndex        =   16
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000003&
      Caption         =   "Porcentaje de remarcación:"
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
      TabIndex        =   15
      Top             =   6000
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000003&
      Caption         =   "Precio de venta:"
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
      TabIndex        =   14
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Localidad 
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
      Left            =   2520
      TabIndex        =   13
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000003&
      Caption         =   "Unidad de medida:"
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
      Left            =   1320
      TabIndex        =   12
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000003&
      Caption         =   "Cantidad:"
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
      Left            =   2280
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
   End
End
Attribute VB_Name = "Productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim E As Integer
Dim PrecioIVA As Double
Dim PrecioVenta As Double
Dim Z As Double
Dim C As Double
Dim Y As Double
Dim V As Double
Dim vSR As Integer

Private Sub Form_Load()
    Productos.Height = 9735
    Productos.Width = 13440
    Productos.Top = ((Screen.Height - Me.Height) / 2) - 650
    Productos.Left = ((Screen.Width - Me.Width) / 2)
    
    cmbUM.AddItem "Unidades"
    cmbUM.AddItem "Kilos"
    
    cmbIVA.AddItem "27"
    cmbIVA.AddItem "21"
    cmbIVA.AddItem "10,5"
    
    camino = App.Path
    Set db = OpenDatabase(camino & "\dbInfor.mdb")
    
    Set TBPRODUCTOS = db.OpenRecordset("TBPRODUCTOS")
    TBPRODUCTOS.Index = "Nro_proI"
    
    Set TBCONSULTA = db.OpenRecordset("TBRUBROS")
    TBCONSULTA.Index = "Nro_RubI"
    
    Set TBCONSULTA2 = db.OpenRecordset("TBSUBRUBROS")
    TBCONSULTA2.Index = "PrimaryKey"
    
    Set TBCONSULTA3 = db.OpenRecordset("TBPRODUCTOS")
    TBCONSULTA3.Index = "Nom_proI"
    
    Set TBCONSULTA4 = db.OpenRecordset("TBSUBRUBROS")
    TBCONSULTA4.Index = "PrimaryKey"
    
    Set TBSUBRUBROS = db.OpenRecordset("TBSUBRUBROS")
    TBSUBRUBROS.Index = "PrimaryKey"
    
    CargaRubro
    CargaSubRubro
    CargaList
    
    TBPRODUCTOS.MoveLast
    txtNroP.Text = TBPRODUCTOS!Nro_pro + 1
End Sub

Sub CargaList()
    TBCONSULTA3.MoveFirst
    Do Until TBCONSULTA3.EOF
        If i > 20 Then
            Exit Sub
        Else
            ListPro.AddItem Format(TBCONSULTA3!Nro_pro, "0000") & " - " & TBCONSULTA3!Nom_pro
        End If
        i = i + 1
        TBCONSULTA3.MoveNext
    Loop
End Sub

Sub CargaSubRubro()
    A = Mid(cmbRub.Text, 1, 4)
    A = Val(A)
    Set TBCONSULTA2 = db.OpenRecordset("select * from TBSUBRUBROS where Rub_Subr = " & A & " order by Nro_Subr")
    If Not TBCONSULTA2.EOF Then
        TBCONSULTA2.MoveFirst
        Do While Not TBCONSULTA2.EOF
            cmbSR.AddItem Format(TBCONSULTA2!Nro_Subr, "0000") & " - " & TBCONSULTA2!Nom_Subr
            If vSR = TBCONSULTA2!Nro_Subr Then
                cmbSR.Text = TBCONSULTA2!Nro_Subr & " - " & TBCONSULTA2!Nom_Subr
            End If
            TBCONSULTA2.MoveNext
        Loop
    End If
End Sub

Sub CargaRubro()
    TBCONSULTA.MoveFirst
    Do Until TBCONSULTA.EOF
        If i > 20 Then
            Exit Sub
        Else
            cmbRub.AddItem Format(TBCONSULTA!Nro_Rub, "0000") & " - " & TBCONSULTA!Nom_Rub
        End If
        i = i + 1
        TBCONSULTA.MoveNext
    Loop
End Sub

Sub BuscaRubro()
    TBCONSULTA.Seek "=", Val(cmbRub.Text)
    If TBCONSULTA.NoMatch Then
        cmbRub.Enabled = True
        cmbRub.SetFocus
    Else
        cmbRub.Text = TBCONSULTA!Nro_Rub & " - " & TBCONSULTA!Nom_Rub
    End If
End Sub

Sub BuscaSubRubro()
    cmbSR.Clear
    TBCONSULTA4.Seek "=", Val(cmbRub.Text)
    If TBCONSULTA4.NoMatch Then
        cmbSR.Enabled = True
        cmbSR.SetFocus
    Else
        'cmbSR.Text = TBCONSULTA4!Nro_Subr & " - " & TBCONSULTA4!Nom_Subr
    End If
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

Sub LIMPIAR()
    txtNroP.Text = ""
    txtNom.Text = ""
    maskFecha.Mask = ""
    txtCant.Text = ""
    txtPrecio.Text = ""
    txtPR.Text = ""
    txtPVenta.Text = ""
    cmbRub.Text = "Seleccionar rubro"
    cmbRub.Enabled = False
    cmbSR.Text = "Seleccionar subrubro"
    cmbSR.Enabled = False
    cmbUM.Text = "Seleccionar unidad"
    cmbUM.Enabled = False
    cmbIVA.Text = "Seleccionar IVA"
    cmbIVA.Enabled = False
    txtNroP.SetFocus
    txtNom.Enabled = False
    maskFecha.Enabled = False
    txtCant.Enabled = False
    txtPrecio.Enabled = False
    txtPR.Enabled = False
    txtPVenta.Enabled = False
    btnAceptar.Enabled = False
    btnCancelar.Enabled = False
    btnBorrar.Enabled = False
    cmbSR.Clear
End Sub

Private Sub ListPro_DblClick()
    Z = Mid(Val(ListPro.Text), 1, 4)
    C = Val(Z)
    TBPRODUCTOS.Seek "=", Val(C)
    If TBPRODUCTOS.NoMatch Then
        txtNom.Enabled = True
        txtNom.SetFocus
        btnCancelar.Enabled = True
        E = 0
    Else
        TBPRODUCTOS.Edit
        txtNroP.Text = TBPRODUCTOS!Nro_pro
        txtNom.Text = TBPRODUCTOS!Nom_pro
        cmbRub.Text = TBPRODUCTOS!Rub_pro
        BuscaRubro
        vSR = Val(TBPRODUCTOS!SubRub_pro)
        cmbSR.Text = TBPRODUCTOS!SubRub_pro
        BuscaSubRubro
        CargaSubRubro
        maskFecha.Mask = TBPRODUCTOS!Fecha_pro
        cmbUM.Text = TBPRODUCTOS!UnidadMedida_pro
        txtCant.Text = TBPRODUCTOS!Cantidad_pro
        txtPrecio.Text = TBPRODUCTOS!Precio_pro
        cmbIVA.Text = TBPRODUCTOS!IVA_pro
        txtPR.Text = TBPRODUCTOS!PorcRemarcacion_pro
        txtPVenta.Text = TBPRODUCTOS!PrecioVenta_pro
        E = 1
        btnCancelar.Enabled = True
        btnAceptar.Enabled = True
        btnBorrar.Enabled = True
        txtNom.Enabled = True
        cmbRub.Enabled = True
        cmbSR.Enabled = True
        maskFecha.Enabled = True
        cmbUM.Enabled = True
        txtCant.Enabled = True
        txtPrecio.Enabled = True
        cmbIVA.Enabled = True
        txtPR.Enabled = True
        txtPVenta.Enabled = False
    End If
End Sub

Private Sub btnBorrar_Click()
    msg = MsgBox("¿Desea eliminar los datos?", vbYesNo + vbDefaultButton2 + vbQuestion, "Eliminar archivo")
    If msg = 6 Then
        TBPRODUCTOS.Delete
        ListPro.RemoveItem (ListPro.ListIndex)
        LIMPIAR
    Else

    End If
    TBPRODUCTOS.MoveLast
    txtNroP.Text = TBPRODUCTOS!Nro_pro + 1
End Sub

Private Sub btnAceptar_Click()
    msg = MsgBox("¿Desea guadar los datos?", vbYesNo + vbQuestion, "Guardar archivo")
    If msg = 6 Then
        If E <> 1 Then
            TBPRODUCTOS.AddNew
            TBPRODUCTOS!Nro_pro = txtNroP.Text
            TBPRODUCTOS!Nom_pro = txtNom.Text
            TBPRODUCTOS!Rub_pro = Mid(Val(cmbRub.Text), 1, 4)
            TBPRODUCTOS!SubRub_pro = Mid(Val(cmbSR.Text), 1, 4)
            TBPRODUCTOS!Fecha_pro = maskFecha.Text
            TBPRODUCTOS!UnidadMedida_pro = cmbUM.Text
            TBPRODUCTOS!Cantidad_pro = txtCant.Text
            TBPRODUCTOS!Precio_pro = txtPrecio.Text
            TBPRODUCTOS!IVA_pro = cmbIVA.Text
            TBPRODUCTOS!PorcRemarcacion_pro = txtPR.Text
            TBPRODUCTOS!PrecioVenta_pro = txtPVenta.Text
            ListPro.AddItem Format(TBPRODUCTOS!Nro_pro, "0000") & " - " & TBPRODUCTOS!Nom_pro
        Else
            TBPRODUCTOS!Nro_pro = txtNroP.Text
            TBPRODUCTOS!Nom_pro = txtNom.Text
            TBPRODUCTOS!Rub_pro = Mid(Val(cmbRub.Text), 1, 4)
            TBPRODUCTOS!SubRub_pro = Mid(Val(cmbSR.Text), 1, 4)
            TBPRODUCTOS!Fecha_pro = maskFecha.Text
            TBPRODUCTOS!UnidadMedida_pro = cmbUM.Text
            TBPRODUCTOS!Cantidad_pro = txtCant.Text
            TBPRODUCTOS!Precio_pro = txtPrecio.Text
            TBPRODUCTOS!IVA_pro = cmbIVA.Text
            TBPRODUCTOS!PorcRemarcacion_pro = txtPR.Text
            TBPRODUCTOS!PrecioVenta_pro = txtPVenta.Text
        End If
        TBPRODUCTOS.Update
        btnCancelar.Enabled = False
        btnAceptar.Enabled = False
        LIMPIAR
    Else
        LIMPIAR
    End If
    TBPRODUCTOS.MoveLast
    txtNroP.Text = TBPRODUCTOS!Nro_pro + 1
End Sub

Private Sub btnCancelar_Click()
    LIMPIAR
    TBPRODUCTOS.MoveLast
    txtNroP.Text = TBPRODUCTOS!Nro_pro + 1
End Sub

Private Sub btnSalir_Click()
    Unload Productos
End Sub

Private Sub txtBuscar_Change()
    ListPro.Clear
    Set TBCONSULTA = db.OpenRecordset("select * from TBPRODUCTOS where Nom_pro >= '" & txtBuscar.Text & "' order by Nom_pro, Nro_pro")
        TBCONSULTA.MoveFirst
    If Not TBCONSULTA.EOF Then
        TBCONSULTA.MoveFirst
        Do While Not TBCONSULTA.EOF
            ListPro.AddItem Format(TBCONSULTA!Nro_pro, "0000") & " - " & TBCONSULTA!Nom_pro
            TBCONSULTA.MoveNext
        Loop
    End If
End Sub

Private Sub txtPVenta_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtPVenta.Text) <= 0 Then
            txtPVenta.SetFocus
        Else
            btnAceptar.Enabled = True
            btnAceptar.SetFocus
        End If
    End If
End Sub

Private Sub txtPR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtPR.Text) <= 0 Then
            txtPR.SetFocus
        Else
            PrecioIVA = Val(txtPrecio.Text) * Val(cmbIVA.Text) / 100
            PrecioIVA = PrecioIVA + Val(txtPrecio.Text)
            PrecioVenta = PrecioIVA * Val(txtPR.Text) / 100
            txtPVenta.Text = PrecioIVA + PrecioVenta
            txtPVenta.Enabled = True
            txtPVenta.SetFocus
        End If
    End If
End Sub

Private Sub cmbIVA_Click()
    If cmbIVA.Text = "" Then
        cmbIVA.SetFocus
    Else
        txtPR.Enabled = True
        txtPR.SetFocus
    End If
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtPrecio.Text = "" Then
            txtPrecio.SetFocus
        Else
            cmbIVA.Enabled = True
            cmbIVA.SetFocus
        End If
    End If
End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtCant.Text = "" Then
            txtCant.SetFocus
        Else
            txtPrecio.Enabled = True
            txtPrecio.SetFocus
        End If
    End If
End Sub

Private Sub cmbUM_Click()
    If cmbUM.Text = "" Then
        cmbUM.SetFocus
    Else
        txtCant.Enabled = True
        txtCant.SetFocus
    End If
End Sub

Private Sub maskFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ValidarFecha(maskFecha.Text) Then
            MsgBox "La fecha informada no es válida", vbExclamation, Me.Caption
            maskFecha.SelStart = 0
            maskFecha.SelLength = Len(maskFecha.Text)
            maskFecha.SetFocus
            Exit Sub
        End If
        cmbUM.Enabled = True
        cmbUM.SetFocus
    End If
End Sub

Private Sub cmbSR_Click()
    If cmbSR.Text = "" Then
        cmbSR.SetFocus
    Else
        V = Mid(cmbSR.Text, 1, 4)
        V = Val(V)
        maskFecha.Enabled = True
        maskFecha.SetFocus
    End If
End Sub

Private Sub cmbRub_Click()
    If cmbRub.Text = "" Then
        cmbRub.SetFocus
    Else
        Y = Mid(cmbRub.Text, 1, 4)
        Y = Val(Y)
        CargaSubRubro
        cmbSR.Enabled = True
        cmbSR.SetFocus
    End If
End Sub

Private Sub txtNom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txtNom.Text) Or txtNom.Text = "" Then
            txtNom.SetFocus
        Else
            txtNom.Text = UCase(txtNom.Text)
            cmbRub.Enabled = True
            cmbRub.SetFocus
        End If
    End If
End Sub

Private Sub txtNroP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtNroP.Text) <= 0 Or txtNroP.Text = "" Then
            txtNroP.SetFocus
        Else
            TBPRODUCTOS.Seek "=", Val(txtNroP.Text)
            If TBPRODUCTOS.NoMatch Then
                txtNom.Enabled = True
                txtNom.SetFocus
                btnCancelar.Enabled = True
                E = 0
            Else
                TBPRODUCTOS.Edit
                txtNom.Text = TBPRODUCTOS!Nom_pro
                cmbRub.Text = TBPRODUCTOS!Rub_pro
                BuscaRubro
                cmbSR.Text = TBPRODUCTOS!SubRub_pro
                BuscaSubRubro
                maskFecha.Mask = TBPRODUCTOS!Fecha_pro
                cmbUM.Text = TBPRODUCTOS!UnidadMedida_pro
                txtCant.Text = TBPRODUCTOS!Cantidad_pro
                txtPrecio.Text = TBPRODUCTOS!Precio_pro
                cmbIVA.Text = TBPRODUCTOS!IVA_pro
                txtPR.Text = TBPRODUCTOS!PorcRemarcacion_pro
                txtPVenta.Text = TBPRODUCTOS!PrecioVenta_pro
                E = 1
                btnCancelar.Enabled = True
                btnAceptar.Enabled = True
                btnBorrar.Enabled = True
                txtNom.Enabled = True
                cmbRub.Enabled = True
                cmbSR.Enabled = True
                maskFecha.Enabled = True
                cmbUM.Enabled = True
                txtCant.Enabled = True
                txtPrecio.Enabled = True
                cmbIVA.Enabled = True
                txtPR.Enabled = True
                txtPVenta.Enabled = True
            End If
        End If
    End If
End Sub
