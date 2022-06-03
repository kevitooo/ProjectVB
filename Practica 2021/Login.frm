VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H80000003&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Iniciar Sesión"
   ClientHeight    =   2325
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1373.687
   ScaleMode       =   0  'User
   ScaleWidth      =   5267.486
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
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
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   1440
      Width           =   1380
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   1440
      Width           =   1380
   End
   Begin VB.TextBox txtPassword 
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
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000003&
      Caption         =   "Nombre de usuario:"
      Height          =   270
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   2160
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000003&
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   1320
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Public loginsucceeded As Boolean

Private Sub cmdCancel_Click()
    loginsucceeded = False
    End
End Sub

Private Sub cmdOK_Click()
    If txtPassword = "123" Then
        If txtUserName = "admin" Then
            loginsucceeded = True
            Me.Hide
            frmSplash.Show
            frmSplash.Timer1.Interval = 1
            frmSplash.Timer1.Enabled = True
        Else
            MsgBox "Los datos ingresados no son válidos. Vuelva a intentarlo.", vbExclamation, "Error al iniciar sesión"
            txtUserName.SetFocus
        End If
    Else
        MsgBox "Los datos ingresados no son válidos. Vuelva a intentarlo", vbExclamation, "Error al iniciar sesión"
        txtUserName.SetFocus
    End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtUserName.Text = "" Then
            txtUserName.SetFocus
        Else
            txtPassword.Enabled = True
            txtPassword.SetFocus
        End If
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtPassword.Text = "" Then
            txtPassword.SetFocus
        Else
            cmdOK.Enabled = True
            cmdOK.SetFocus
        End If
    End If
End Sub
