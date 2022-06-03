VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4335
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4050
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7080
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   3600
         Visible         =   0   'False
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Left            =   6480
         Top             =   3480
      End
      Begin VB.Image imgLogo 
         Height          =   1785
         Left            =   2640
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "¡Bienvenido!"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
    If ProgressBar1.Value = 100 Then
        Timer1.Interval = 0
        Me.Hide
        SisPres.Show
    Else
        ProgressBar1.Value = Val(ProgressBar1.Value) + Val(1)
    End If
End Sub
