VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FF8080&
   Caption         =   "Form Main"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBooks 
      BackColor       =   &H00C000C0&
      Caption         =   "List of Books"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   3855
   End
   Begin VB.ListBox lstEmployee 
      Height          =   1815
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000C0&
      Caption         =   "Exit"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H000080FF&
      Caption         =   "Edit Employee"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H000000C0&
      Caption         =   "Browse"
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H000080FF&
      Caption         =   "Remove Employee"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H000080FF&
      Caption         =   "Display Employee"
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H000080FF&
      Caption         =   "Add Employee"
      Default         =   -1  'True
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdAdd_Click()
 ' Display the Employee file and Add the item to the list
 With frmEmployee
  .txtLastName = " "
  .txtFirstName = " "
  .txtStreet = " "
  .txtCity = " "
  .txtState = " "
  .txtZip = " "
  .txtPhone = " "
  .txtEmail = " "
 End With
 
  With frmEmployee
  .Action = "A"
  .show vbModal
 End With
End Sub


Private Sub cmdBooks_Click()
Load frmBookData
frmBookData.show
End Sub

Private Sub cmdBrowse_Click()
 'Display the first Employee in list
 With frmEmployee
  .txtLastName = " "
  .txtFirstName = " "
  .txtStreet = " "
  .txtCity = " "
  .txtState = " "
  .txtZip = " "
  .txtPhone = " "
  .txtEmail = " "
 End With
 With lstEmployee
  .ListIndex = 0
  frmEmployee.Key = Trim(Str(.ItemData(0)))
 End With
 With frmEmployee
  .Action = "B"
  .show vbModal
 End With
End Sub

Private Sub cmdDisplay_Click()
 'Display the Employee member
 With frmEmployee
  .txtLastName = " "
  .txtFirstName = " "
  .txtStreet = " "
  .txtCity = " "
  .txtState = " "
  .txtZip = " "
  .txtPhone = " "
  .txtEmail = " "
 End With
  frmEmployee.Action = "D"
  DisplayForm
 

End Sub

Private Sub cmdEdit_Click()
 'Edit the Employee member
 With frmEmployee
  .txtLastName = " "
  .txtFirstName = " "
  .txtStreet = " "
  .txtCity = " "
  .txtState = " "
  .txtZip = " "
  .txtPhone = " "
  .txtEmail = " "
 End With

 frmEmployee.Action = "E"
  DisplayForm
End Sub

Private Sub cmdExit_Click()
 'Terminate the project
 
 Unload frmEmployee
 Unload Me
 End
End Sub

Private Sub cmdRemove_Click()
 'Remove the Employee member
 With frmEmployee
  .txtLastName = " "
  .txtFirstName = " "
  .txtStreet = " "
  .txtCity = " "
  .txtState = " "
  .txtZip = " "
  .txtPhone = " "
  .txtEmail = " "
 End With
 frmEmployee.Action = "R"
 DisplayForm
End Sub

Private Sub Form_Load()
 'Force file creation from employee form
 'frmNewMain.show
 frmMain.Hide
 Splash.show
 Load frmEmployee
End Sub

Private Sub DisplayForm()
 'Set key and show frmEmployee
 Dim strMsg  As String
 
 If lstEmployee.ListIndex <> -1 Then
   With lstEmployee
     frmEmployee.Key = Trim(Str(.ItemData(.ListIndex)))
   End With
  frmEmployee.show vbModal
 Else
   strMsg = "Please select an employee from the list."
   MsgBox strMsg, vbInformation, "Employee File"
 End If
End Sub


