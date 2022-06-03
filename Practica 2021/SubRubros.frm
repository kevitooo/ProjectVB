VERSION 5.00
Begin VB.Form SubRubros 
   BackColor       =   &H80000003&
   Caption         =   "Subrubros"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12270
   Icon            =   "SubRubros.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   12270
   Begin VB.ComboBox cmbRub 
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
      Left            =   2760
      TabIndex        =   0
      Text            =   "Seleccionar rubro"
      Top             =   1560
      Width           =   3255
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
      Left            =   8400
      TabIndex        =   6
      Top             =   7200
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
      Left            =   6240
      TabIndex        =   5
      Top             =   7200
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
      Left            =   4080
      TabIndex        =   4
      Top             =   7200
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
      Left            =   1920
      TabIndex        =   3
      Top             =   7200
      Width           =   1815
   End
   Begin VB.TextBox txtNroSR 
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
      Left            =   3120
      MaxLength       =   4
      TabIndex        =   1
      Top             =   3840
      Width           =   855
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
      Left            =   3120
      MaxLength       =   25
      TabIndex        =   2
      Top             =   4440
      Width           =   2895
   End
   Begin VB.ListBox ListSR 
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
      Height          =   5160
      Left            =   6960
      TabIndex        =   8
      Top             =   1200
      Width           =   4935
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
      Left            =   7440
      TabIndex        =   7
      Top             =   600
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "Seleccionar rubro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   9
      Top             =   960
      Width           =   6135
      Begin VB.Label Label4 
         BackColor       =   &H80000003&
         Caption         =   "Nombre de rubro:"
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
         Left            =   480
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000003&
      Caption         =   "Agregar subrubro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   11
      Top             =   3240
      Width           =   6135
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Nombre de subrubro:"
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
         Left            =   480
         TabIndex        =   13
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000003&
         Caption         =   "Número de subrubro:"
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
         Left            =   480
         TabIndex        =   12
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   7080
      Picture         =   "SubRubros.frx":1084A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   255
   End
End
Attribute VB_Name = "SubRubros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim E As Integer
Dim X As Integer
Dim C As Integer
Dim Y As Integer

Private Sub Form_Load()
    Me.Height = 9045
    Me.Width = 12510
    Me.Top = ((Screen.Height - Me.Height) / 2) - 650
    Me.Left = ((Screen.Width - Me.Width) / 2)
    
    camino = App.Path
    Set db = OpenDatabase(camino & "\dbInfor.mdb")
    
    Set TBSUBRUBROS = db.OpenRecordset("TBSUBRUBROS")
    TBSUBRUBROS.Index = "NomI_Subr"
    
    Set TBCONSULTA = db.OpenRecordset("TBRUBROS")
    TBCONSULTA.Index = "Nro_RubI"
    
    Set TBSUBRUBROS = db.OpenRecordset("TBSUBRUBROS")
    TBSUBRUBROS.Index = "PrimaryKey"
    
    Set TBCONSULTA2 = db.OpenRecordset("TBSUBRUBROS")
    TBCONSULTA2.Index = "NomI_Subr"
    
    CargaRubro
    CargaList
End Sub

Sub ConsultaRubro()
    TBCONSULTA.Seek "=", Val(C)
    If TBCONSULTA.NoMatch Then
        MsgBox "No se ha encontrado el rubro", vbExclamation, "Error"
    Else
        cmbRub.Text = Format(TBCONSULTA!Nro_Rub, "0000") & " - " & TBCONSULTA!Nom_Rub
    End If
End Sub

Private Sub ListSR_DblClick()
    Z = Mid(ListSR.Text, 1, 4)
    C = Val(Z)
    V = Mid(ListSR.Text, 8, 4)
    F = Val(V)
    TBSUBRUBROS.Seek "=", Val(C), Val(F)
    If TBSUBRUBROS.NoMatch Then
        txtNom.Enabled = True
        txtNom.SetFocus
        btnCancelar.Enabled = True
        E = 0
    Else
        TBSUBRUBROS.Edit
        cmbRub.Text = TBSUBRUBROS!Rub_Subr
        ConsultaRubro
        txtNroSR.Text = TBSUBRUBROS!Nro_Subr
        txtNom.Text = TBSUBRUBROS!Nom_Subr
        E = 1
        btnCancelar.Enabled = True
        btnAceptar.Enabled = True
        btnBorrar.Enabled = True
        txtNom.Enabled = True
    End If
End Sub

Sub CargaList()
    If Not TBCONSULTA2.BOF Then
        TBCONSULTA2.MoveFirst
        Do Until TBCONSULTA2.EOF
            If i > 20 Then
                Exit Sub
            Else
                ListSR.AddItem Format(TBCONSULTA2!Rub_Subr, "0000") & " - " & Format(TBCONSULTA2!Nro_Subr, "0000") & " - " & TBCONSULTA2!Nom_Subr
            End If
            i = i + 1
            TBCONSULTA2.MoveNext
        Loop
    End If
End Sub

Sub CargaRubro()
    If Not TBCONSULTA.BOF Then
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
    End If
End Sub

Private Sub txtBuscar_Change()
    ListSR.Clear
    Set TBCONSULTA = db.OpenRecordset("select * from TBSUBRUBROS where Nom_Subr >= '" & txtBuscar.Text & "' order by Nom_Subr, Nro_Subr")
        TBCONSULTA.MoveFirst
    If Not TBCONSULTA.EOF Then
        TBCONSULTA.MoveFirst
        Do While Not TBCONSULTA.EOF
            ListSR.AddItem Format(TBCONSULTA!Nro_Subr, "0000") & " - " & TBCONSULTA!Nom_Subr
                TBCONSULTA.MoveNext
        Loop
    End If
End Sub

Sub LIMPIAR()
    cmbRub.SetFocus
    cmbRub.Text = ""
    cmbRub.Text = "Seleccionar rubro"
    txtNroSR.Text = ""
    txtNom.Text = ""
    txtNom.Enabled = False
    btnAceptar.Enabled = False
    btnCancelar.Enabled = False
    btnBorrar.Enabled = False
End Sub

Private Sub btnAceptar_Click()
    msg = MsgBox("¿Desea guadar los datos?", vbYesNo + vbQuestion, "Guardar archivo")
    If msg = 6 Then
        If E <> 1 Then
            TBSUBRUBROS.AddNew
            TBSUBRUBROS!Rub_Subr = Mid(Val(cmbRub.Text), 1, 1)
            TBSUBRUBROS!Nro_Subr = txtNroSR.Text
            TBSUBRUBROS!Nom_Subr = txtNom.Text
            ListSR.AddItem Format(TBSUBRUBROS!Rub_Subr, "0000") & " - " & Format(TBSUBRUBROS!Nro_Subr, "0000") & " - " & TBSUBRUBROS!Nom_Subr
        Else
            TBSUBRUBROS!Rub_Subr = Mid(Val(cmbRub.Text), 1, 1)
            TBSUBRUBROS!Nro_Subr = txtNroSR.Text
            TBSUBRUBROS!Nom_Subr = txtNom.Text
        End If
        TBSUBRUBROS.Update
        btnCancelar.Enabled = False
        btnAceptar.Enabled = False
        LIMPIAR
    Else
        LIMPIAR
    End If
End Sub

Private Sub btnCancelar_Click()
    LIMPIAR
End Sub

Private Sub btnBorrar_Click()
    msg = MsgBox("¿Desea eliminar los datos?", vbYesNo + vbDefaultButton2 + vbQuestion, "Eliminar archivo")
    If msg = 6 Then
        TBSUBRUBROS.Delete
        ListSR.RemoveItem (ListSR.ListIndex)
        LIMPIAR
    Else

    End If
End Sub

Private Sub btnSalir_Click()
    Unload SubRubros
End Sub

Private Sub txtNom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txtNom.Text) Or txtNom.Text = "" Then
            txtNom.SetFocus
        Else
            txtNom.Text = UCase(txtNom.Text)
            btnAceptar.Enabled = True
            btnAceptar.SetFocus
        End If
    End If
End Sub

Private Sub txtNroSR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtNroSR.Text) <= 0 Or txtNroSR.Text = "" Then
            txtNroSR.SetFocus
        Else
            TBSUBRUBROS.Seek "=", Y, Val(txtNroSR.Text)
            If TBSUBRUBROS.NoMatch Then
                txtNom.Enabled = True
                txtNom.SetFocus
                btnCancelar.Enabled = True
                E = 0
            Else
                TBSUBRUBROS.Edit
                txtNom.Text = TBSUBRUBROS!Nom_Subr
                E = 1
                btnCancelar.Enabled = True
                btnAceptar.Enabled = True
                btnBorrar.Enabled = True
                txtNom.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub cmbRub_Click()
    If cmbRub.Text = "" Then
        cmbRub.SetFocus
    Else
        Y = Mid(cmbRub.Text, 1, 4)
        Y = Val(Y)
        txtNroSR.Enabled = True
        txtNroSR.SetFocus
        btnCancelar.Enabled = True
    End If
End Sub
