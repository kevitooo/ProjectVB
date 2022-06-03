VERSION 5.00
Begin VB.Form CodPostal 
   BackColor       =   &H80000003&
   Caption         =   "Código Postal"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12450
   Icon            =   "CodPostal.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   12450
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
      TabIndex        =   13
      Top             =   960
      Width           =   3495
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
      TabIndex        =   12
      Top             =   6000
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
      TabIndex        =   11
      Top             =   6000
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
      TabIndex        =   10
      Top             =   6000
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
      Left            =   8520
      TabIndex        =   9
      Top             =   6000
      Width           =   1815
   End
   Begin VB.ListBox ListCP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      ItemData        =   "CodPostal.frx":1084A
      Left            =   7680
      List            =   "CodPostal.frx":1084C
      TabIndex        =   4
      Top             =   1560
      Width           =   4095
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
      Left            =   3600
      TabIndex        =   2
      Text            =   "Seleccionar provincia"
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox txtCod 
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
      Left            =   3600
      MaxLength       =   4
      TabIndex        =   0
      Top             =   2040
      Width           =   855
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
      Left            =   3600
      MaxLength       =   25
      TabIndex        =   1
      Top             =   2760
      Width           =   2895
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
      Left            =   3600
      MaxLength       =   5
      TabIndex        =   3
      Top             =   4200
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   7800
      Picture         =   "CodPostal.frx":1084E
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      Caption         =   "Código Postal:"
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
      Left            =   1920
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
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
      Left            =   2400
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label3 
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
      Left            =   2520
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label4 
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
      Left            =   840
      TabIndex        =   5
      Top             =   4320
      Width           =   2655
   End
End
Attribute VB_Name = "CodPostal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Z As Integer
Dim C As Integer
Dim E As Integer
Dim VcmbProv As Integer

Private Sub Form_Load()
    Me.Height = 8205
    Me.Width = 12690
    Me.Top = ((Screen.Height - Me.Height) / 2) - 650
    Me.Left = ((Screen.Width - Me.Width) / 2)
    
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
    
    camino = App.Path
    Set db = OpenDatabase(camino & "\dbInfor.mdb")
    
    Set TBCODPOS = db.OpenRecordset("TBCODPOS")
    TBCODPOS.Index = "NroI_cpos"
    
    Set TBCONSULTA = db.OpenRecordset("TBCODPOS")
    TBCONSULTA.Index = "LocI_cpos"
    
    Set TBCONSULTA2 = db.OpenRecordset("TBCODPOS")
    TBCONSULTA2.Index = "LocI_cpos"
    
    CargaList
End Sub

Sub CargaList()
    TBCONSULTA.MoveFirst
    Do Until TBCONSULTA.EOF
        If i > 20 Then
            Exit Sub
        Else
            ListCP.AddItem Format(TBCONSULTA!Nro_cpos, "0000") & " - " & TBCONSULTA!Loc_cpos
        End If
        i = i + 1
        TBCONSULTA.MoveNext
    Loop
End Sub

Sub LIMPIAR()
    txtCod = ""
    txtLoc.Text = ""
    txtLoc.Enabled = False
    cmbProv.Text = "Seleccionar provincia"
    cmbProv.Enabled = False
    txtCT.Text = ""
    txtCT.Enabled = False
    txtCod.SetFocus
    btnAceptar.Enabled = False
    btnCancelar.Enabled = False
    btnBorrar.Enabled = False
End Sub

Private Sub btnBorrar_Click()
    msg = MsgBox("¿Desea eliminar los datos?", vbYesNo + vbQuestion, "Eliminar archivo")
    If msg = 6 Then
        TBCODPOS.Delete
        ListCP.RemoveItem (ListCP.ListIndex)
        LIMPIAR
    Else
        
    End If
End Sub

Private Sub btnAceptar_Click()
    msg = MsgBox("¿Desea guadar los datos?", vbYesNo + vbQuestion, "Guardar archivo")
    If msg = 6 Then
        If E = 0 Then
            TBCODPOS.AddNew
            TBCODPOS!Nro_cpos = txtCod.Text
            TBCODPOS!Loc_cpos = txtLoc.Text
            TBCODPOS!Pcia_cpos = Mid(Val(cmbProv.Text), 1, 2)
            TBCODPOS!Car_cpos = txtCT.Text
            ListCP.AddItem Format(TBCODPOS!Nro_cpos, "0000") & " - " & TBCODPOS!Loc_cpos
        Else
            TBCODPOS.Edit
            TBCODPOS!Nro_cpos = txtCod.Text
            TBCODPOS!Loc_cpos = txtLoc.Text
            TBCODPOS!Pcia_cpos = Mid(Val(cmbProv.Text), 1, 2)
            TBCODPOS!Car_cpos = txtCT.Text
        End If
        TBCODPOS.Update
        btnCancelar.Enabled = False
        btnAceptar.Enabled = False
        LIMPIAR
    Else
        LIMPIAR
    End If
End Sub

Private Sub btnSalir_Click()
    Unload CodPostal
End Sub

Private Sub btnCancelar_Click()
    LIMPIAR
End Sub

Private Sub ListCP_DblClick()
    Z = Mid(Val(ListCP.Text), 1, 4)
    C = Val(Z)
    TBCODPOS.Seek "=", Val(C)
    If TBCODPOS.NoMatch Then
        txtLoc.Enabled = True
        txtLoc.SetFocus
        btnCancelar.Enabled = True
        E = 0
    Else
        txtCod.Text = TBCODPOS!Nro_cpos
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
        txtLoc.Enabled = True
        cmbProv.Enabled = True
        txtCT.Enabled = True
        btnAceptar.Enabled = True
        btnCancelar.Enabled = True
        btnBorrar.Enabled = True
        E = 1
    End If
End Sub

Private Sub txtBuscar_Change()
    ListCP.Clear
    Set TBCONSULTA2 = db.OpenRecordset("select * from TBCODPOS where Loc_cpos >= '" & txtBuscar.Text & "' order by Loc_cpos, Nro_cpos")
        TBCONSULTA2.MoveFirst
    If Not TBCONSULTA2.EOF Then
        TBCONSULTA2.MoveFirst
        Do While Not TBCONSULTA2.EOF
            ListCP.AddItem Format(TBCONSULTA2!Nro_cpos, "0000") & " - " & TBCONSULTA2!Loc_cpos
            TBCONSULTA2.MoveNext
        Loop
    End If
End Sub

Private Sub txtCT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtCT.Text) <= 0 Or txtCT.Text = "" Then
            txtCT.SetFocus
        Else
            btnAceptar.Enabled = True
            btnAceptar.SetFocus
        End If
    End If
End Sub

Private Sub cmbProv_Click()
    If cmbProv.Text = "" Then
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

Private Sub txtCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtCod.Text) < 1000 Then
            txtCod.SetFocus
        Else
            TBCODPOS.Seek "=", Val(txtCod.Text)
            If TBCODPOS.NoMatch Then
                txtLoc.Enabled = True
                txtLoc.SetFocus
                btnCancelar.Enabled = True
                E = 0
            Else
                txtCod.Text = TBCODPOS!Nro_cpos
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
                txtLoc.Enabled = True
                cmbProv.Enabled = True
                txtCT.Enabled = True
                btnAceptar.Enabled = True
                btnCancelar.Enabled = True
                btnBorrar.Enabled = True
                E = 1
            End If
        End If
    End If
End Sub
