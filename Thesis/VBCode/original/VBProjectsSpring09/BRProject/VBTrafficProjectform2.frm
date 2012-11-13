VERSION 5.00
Begin VB.Form frmSignInName 
   BackColor       =   &H80000003&
   Caption         =   "Traffic Project(2)"
   ClientHeight    =   7095
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   ScaleHeight     =   7095
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   1935
      Left            =   3120
      Picture         =   "VBTrafficProjectform2.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   4080
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   6480
      Picture         =   "VBTrafficProjectform2.frx":911E
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   4
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdFindSignInName 
      Caption         =   "Find Sign In Name"
      Height          =   975
      Left            =   4680
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   1935
      Left            =   6480
      Picture         =   "VBTrafficProjectform2.frx":13724
      ScaleHeight     =   1875
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   3360
      Picture         =   "VBTrafficProjectform2.frx":1ECCE
      ScaleHeight     =   1875
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdShowFormOne 
      Caption         =   "Back to Form One"
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   3360
      Width           =   1815
   End
End
Attribute VB_Name = "frmSignInName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Traffic Project Sign-In Name Finder 'Traffic Program: This program was written between the dates of March 16th, 2009 and March 24th, 2009 by Bill Roiger.

Private Sub cmdFindSignInName_Click() 'Enables user to find out what a given student's school sign-in name will be.
    Dim Title As String, signinname As String, N As Integer, txtName As String
    Dim middle As String, last As String, first As String
        frmTrafficProject.picResults.Cls
            Title = InputBox("Enter your name below", "Name")
                N = InStr(Title, " ")
                first = Left(Title, N - 1)
                last = Right(Title, Len(Title) - (N + 2))
                middle = Mid(Title, N + 1, 1)
                signinname = Left(first, 1) & " " & middle & Left(last, 6)
                frmSignInName.Hide
                frmTrafficProject.Show
                frmTrafficProject.picResults.Print "Your sign-in name is "; signinname
End Sub

Private Sub cmdShowFormOne_Click() 'Hides the sign-in name page and shows the main traffic project page.
frmSignInName.Hide
frmTrafficProject.Show
End Sub
