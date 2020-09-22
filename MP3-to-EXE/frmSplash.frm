VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5340
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8475
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":08CA
   ScaleHeight     =   5340
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   0
      Top             =   3480
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdModify_Click()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Timer1_Timer()
Unload Me
frmWriter.Show vbModal
End Sub
