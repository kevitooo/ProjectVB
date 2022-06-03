VERSION 5.00
Begin VB.Form Rubros 
   BackColor       =   &H80000003&
   Caption         =   "Rubros"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12270
   Icon            =   "Rubros.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   12270
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
      Left            =   8520
      TabIndex        =   9
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
      Left            =   6360
      TabIndex        =   4
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
      Left            =   4200
      TabIndex        =   3
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
      Left            =   2040
      TabIndex        =   2
      Top             =   7200
      Width           =   1815
   End
   Begin VB.TextBox txtNroR 
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
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   0
      Top             =   2400
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
      Left            =   2520
      MaxLength       =   25
      TabIndex        =   1
      Top             =   3000
      Width           =   3135
   End
   Begin VB.ListBox ListRub 
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
      Left            =   6600
      TabIndex        =   6
      Top             =   1200
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
      Left            =   7080
      TabIndex        =   5
      Top             =   600
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   6720
      Picture         =   "Rubros.frx":1084A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      Caption         =   "Número de rubro:"
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
      Left            =   600
      TabIndex        =   8
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label2 
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
      Left            =   600
      TabIndex        =   7
      Top             =   3120
      Width           =   1935
   End
End
Attribute VB_Name = "Rubros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim E As Integer
Dim X As Integer

Private Sub btnSalir_Click()
    Unload Rubros
End Sub

Private Sub Form_Load()
    Me.Height = 9045
    Me.Width = 12510
    Me.Top = ((Screen.Height - Me.Height) / 2) - 650
    Me.Left = ((Screen.Width - Me.Width) / 2)
    
    camino = App.Path
    Set db = OpenDatabase(camino & "\dbInfor.mdb")
    
    Set TBRUBROS = db.OpenRecordset("TBRUBROS")
    TBRUBROS.Index = "Nro_RubI"
    
    Set TBCONSULTA = db.OpenRecordset("TBRUBROS")
    TBCONSULTA.Index = "Nro_RubI"
    
    CargaList
    
    TBRUBROS.MoveLast
    txtNroR.Text = TBRUBROS!Nro_Rub + 1
End Sub

Private Sub ListRub_DblClick()
    Z = Mid(Val(ListRub.Text), 1, 4)
    C = Val(Z)
    TBRUBROS.Seek "=", Val(C)
    If TBRUBROS.NoMatch Then
        txtNom.Enabled = True
        txtNom.SetFocus
        btnCancelar.Enabled = True
        E = 0
    Else
        TBRUBROS.Edit
        txtNroR.Text = TBRUBROS!Nro_Rub
        txtNom.Text = TBRUBROS!Nom_Rub
        E = 1
        btnCancelar.Enabled = True
        btnAceptar.Enabled = True
        btnBorrar.Enabled = True
        txtNom.Enabled = True
    End If
End Sub

Sub CargaList()
    If Not TBCONSULTA.BOF Then
        TBCONSULTA.MoveFirst
        Do Until TBCONSULTA.EOF
            If i > 20 Then
                Exit Sub
            Else
                ListRub.AddItem Format(TBCONSULTA!Nro_Rub, "0000") & " - " & TBCONSULTA!Nom_Rub
            End If
            i = i + 1
            TBCONSULTA.MoveNext
        Loop
    End If
End Sub

Private Sub txtBuscar_Change()
    ListRub.Clear
    Set TBCONSULTA = db.OpenRecordset("select * from TBRUBROS where Nom_Rub >= '" & txtBuscar.Text & "' order by Nom_Rub, Nro_Rub")
        TBCONSULTA.MoveFirst
    If Not TBCONSULTA.EOF Then
        TBCONSULTA.MoveFirst
        Do While Not TBCONSULTA.EOF
            ListRub.AddItem Format(TBCONSULTA!Nro_Rub, "0000") & " - " & TBCONSULTA!Nom_Rub
                TBCONSULTA.MoveNext
        Loop
    End If
End Sub

Sub LIMPIAR()
    txtNroR.Text = ""
    txtNom.Text = ""
    txtNroR.SetFocus
    txtNom.Enabled = False
    btnAceptar.Enabled = False
    btnCancelar.Enabled = False
    btnBorrar.Enabled = False
End Sub

Private Sub btnAceptar_Click()
    msg = MsgBox("¿Desea guadar los datos?", vbYesNo + vbQuestion, "Guardar archivo")
    If msg = 6 Then
        If E <> 1 Then
            TBRUBROS.AddNew
            TBRUBROS!Nro_Rub = txtNroR.Text
            TBRUBROS!Nom_Rub = txtNom.Text
            ListRub.AddItem Format(TBRUBROS!Nro_Rub, "0000") & " - " & TBRUBROS!Nom_Rub
        Else
            TBRUBROS!Nro_Rub = txtNroSR.Text
            TBRUBROS!Nom_Rub = txtNom.Text
        End If
        TBRUBROS.Update
        btnCancelar.Enabled = False
        btnAceptar.Enabled = False
        LIMPIAR
    Else
        LIMPIAR
    End If
    TBRUBROS.MoveLast
    txtNroR.Text = TBRUBROS!Nro_Rub + 1
End Sub

Private Sub btnCancelar_Click()
    LIMPIAR
    TBRUBROS.MoveLast
    txtNroR.Text = TBRUBROS!Nro_Rub + 1
End Sub

Private Sub btnBorrar_Click()
    msg = MsgBox("¿Desea eliminar los datos?", vbYesNo + vbDefaultButton2 + vbQuestion, "Eliminar archivo")
    If msg = 6 Then
        TBRUBROS.Delete
        ListRub.RemoveItem (ListRub.ListIndex)
        LIMPIAR
    Else

    End If
    TBRUBROS.MoveLast
    txtNroR.Text = TBRUBROS!Nro_Rub + 1
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

Private Sub txtNroR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtNroR.Text) <= 0 Or txtNroR.Text = "" Then
            txtNroR.SetFocus
        Else
            TBRUBROS.Seek "=", Val(txtNroR.Text)
            If TBRUBROS.NoMatch Then
                txtNom.Enabled = True
                txtNom.SetFocus
                btnCancelar.Enabled = True
                E = 0
            Else
                TBRUBROS.Edit
                txtNom.Text = TBRUBROS!Nom_Rub
                E = 1
                btnCancelar.Enabled = True
                btnAceptar.Enabled = True
                btnBorrar.Enabled = True
                txtNom.Enabled = True
            End If
        End If
    End If
End Sub
