VERSION 5.00
Begin VB.Form frmCatcher 
   BackColor       =   &H000000FF&
   Caption         =   "Joe Mauer"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    frmCatcher.Hide 'hides joe mauer
    frmStarting.Show 'back to the starting line-up
End Sub
