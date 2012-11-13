VERSION 5.00
Begin VB.Form frmInventory 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00404040&
      Caption         =   "Load Inventory"
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdInventory 
      BackColor       =   &H00404040&
      Caption         =   "Calculate Inventory"
      Height          =   1095
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox txtSource4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtFresnel 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtPars 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtS4Par 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtCyc 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   3375
      Left            =   360
      ScaleHeight     =   3315
      ScaleWidth      =   8835
      TabIndex        =   1
      Top             =   3120
      Width           =   8895
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main Menu"
      Height          =   975
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6720
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Enter the number of Source 4 lights used:"
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
      TabIndex        =   12
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Enter the number of Fresnel lights used:"
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
      TabIndex        =   11
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "Enter the number of Par Can lights used:"
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
      TabIndex        =   10
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "Enter the number of Source 4 Par lights used:"
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
      TabIndex        =   9
      Top             =   2040
      Width           =   5055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      Caption         =   "Enter the number of Cyc lights used:"
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
      TabIndex        =   8
      Top             =   2520
      Width           =   5055
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808080&
      Caption         =   "Fill each box, add a zero if no lights are being added"
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Theater Lighting
'Form Name: frmInventory
'Author: Kurt Oostra
'Date Written:3/11/08
'Objective: Inventory the lights you have on hand
Option Explicit
Dim Fixtures(1 To 6) As String, Units(1 To 6) As Single
Private Sub cmdReturn_Click()
'Returns to main menu
frmMainMenu.Show
frmInventory.Hide
End Sub

Private Sub cmdInventory_Click()
Dim S4 As Single, S4Par As Single, Cyc As Single, Fres As Single, Par As Single
S4Par = txtS4Par
Fres = txtFresnel
S4 = txtSource4
Par = txtPars
Cyc = txtCyc
'prints the new inventory numbers
picResults.Cls
picResults.Print "Fixture"; Tab(30); "Units Available"
picResults.Print "*******************************************************"
picResults.Print Fixtures(1); Tab(30); Units(1) - S4
picResults.Print Fixtures(2); Tab(30); Units(2) - Fres
picResults.Print Fixtures(3); Tab(30); Units(3) - Par
picResults.Print Fixtures(4); Tab(30); Units(4) - S4Par
picResults.Print Fixtures(5); Tab(30); Units(5) - Cyc
End Sub

Private Sub cmdLoad_Click()
Dim ctr As Integer, j As Integer
picResults.Cls
'loads the inventory array
Open App.Path & "\inventory.txt" For Input As #1
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Fixtures(ctr), Units(ctr)
Loop
'prints a heading
picResults.Print "Fixture"; Tab(30); "Units Available"
picResults.Print "*******************************************************"
'prints the array
For j = 1 To ctr
    picResults.Print Fixtures(j); Tab(30); Units(j)
Next
'closes the file path so that it can be reloaded after changing the array
Close #1
'opens the txt boxes so you can type in the number of lights you use
txtSource4.Enabled = True
txtFresnel.Enabled = True
txtPars.Enabled = True
txtCyc.Enabled = True
txtS4Par.Enabled = True
End Sub

Private Sub txtCyc_Change()
'the next 5 subs make it so you can't type in a number greater than the number of lights you have on hand
'a message box is printed if the number is too high
If txtCyc > Units(5) Then MsgBox ("Their are not that many units available.")
End Sub

Private Sub txtFresnel_Change()
If txtFresnel > Units(2) Then MsgBox ("Their are not that many units available.")
End Sub

Private Sub txtPars_Change()
If txtPars > Units(3) Then MsgBox ("Their are not that many units available.")
End Sub

Private Sub txtS4Par_Change()
If txtS4Par > Units(4) Then MsgBox ("Their are not that many units available.")
End Sub

Private Sub txtSource4_Change()
If txtSource4 > Units(1) Then MsgBox ("Their are not that many units available.")
End Sub
