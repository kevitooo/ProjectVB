VERSION 5.00
Begin VB.MDIForm SisPres 
   BackColor       =   &H80000002&
   Caption         =   "Sistema Presupuesto"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   Icon            =   "SisPres.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu A 
      Caption         =   "&Archivos"
      Begin VB.Menu C 
         Caption         =   "Clientes"
      End
      Begin VB.Menu Pr 
         Caption         =   "Presupuestos"
      End
      Begin VB.Menu CP 
         Caption         =   "Códigos Postales"
      End
      Begin VB.Menu P 
         Caption         =   "Productos"
      End
      Begin VB.Menu R 
         Caption         =   "Rubros"
      End
      Begin VB.Menu SR 
         Caption         =   "Subrubros"
      End
   End
   Begin VB.Menu Salir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "SisPres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub C_Click()
    Clientes.Show
End Sub

Private Sub CP_Click()
    CodPostal.Show
End Sub

Private Sub pr_click()
    Presupuesto.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub P_Click()
    Productos.Show
End Sub

Private Sub R_Click()
    Rubros.Show
End Sub

Private Sub SR_Click()
    SubRubros.Show
End Sub

Private Sub Salir_Click()
    msg = MsgBox("¿Estás seguro de salir de la aplicación?", vbYesNo + vbDefaultButton2 + vbQuestion, "Salir del sistema")
    If msg = 6 Then
        End
    Else
        
    End If
End Sub
