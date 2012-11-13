VERSION 5.00
Begin VB.Form frmCircle 
   BackColor       =   &H0000FF00&
   Caption         =   "Return to Main Menu"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkFlood 
      Caption         =   "Check2"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkSpot 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txtThrow 
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CheckBox chk36 
      Caption         =   "Check3"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox chk26 
      Caption         =   "Check2"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox chk19 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton cmdFres 
      BackColor       =   &H0000FFFF&
      Caption         =   "Calculate the diameter of the cirle made by a Fresnel"
      Height          =   1095
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdS4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Calculate the diameter of the circle made by a Source Four"
      Height          =   1215
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main Menu"
      Height          =   855
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   6495
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FF00&
      Caption         =   "Fresnel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   16
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Caption         =   "Source Four"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   15
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      Caption         =   "Flood"
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FF00&
      Caption         =   "Spot"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FF00&
      Caption         =   "Enter the distance between the fixture and the focus point"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "36 degree fixture"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "26 degree fixture"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "19 degree fixture"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Theater Lighting
'Form Name: frmCircle
'Author: Kurt Oostra
'Date Written:3/26/08
'Objective: Determine the diameter of a circle of 'x' amount of feet away of different light fixtures.
Option Explicit
Dim throw As Single, diameter As Single
Private Sub chk19_Click()
'the next 5 subs make it so that only one check box can be filled at a time
If chk19.Value = 1 Then
    chk26.Value = 0
    chk36.Value = 0
End If
End Sub

Private Sub chk26_Click()
If chk26.Value = 1 Then
    chk19.Value = 0
    chk36.Value = 0
End If
End Sub

Private Sub chk36_Click()
If chk36.Value = 1 Then
    chk26.Value = 0
    chk19.Value = 0
End If
End Sub

Private Sub chkFlood_Click()
If chkFlood.Value = 1 Then chkSpot.Value = 0
End Sub

Private Sub chkSpot_Click()
If chkSpot.Value = 1 Then chkFlood.Value = 0
End Sub

Private Sub cmdFres_Click()
Dim size As Single, diameter As Single
'input box asking the size of the Fresnel
throw = txtThrow
size = InputBox("Enter '6' for a 6 inch Fresnel, or '8' for an 8 inch Fresnel.")
'if statements determining whether the fresnel is spotted or flooded and then calculating the diamter of the circle.
If size = 6 Then
    If chkSpot = 1 Then
        diameter = throw * 0.1 + 0.1
        MsgBox ("The diameter of a 6 inch Fresnel on spot at " & throw & " is " & diameter & " feet.")
    End If
    If chkFlood = 1 Then
        diameter = throw * 1.13
        MsgBox ("The diameter of a 6 inch Fresnel on flood at " & throw & " is " & diameter & " feet.")
    End If
End If
If size = 8 Then
    If chkSpot = 1 Then
        diameter = throw * 0.14
        MsgBox ("The diameter of a 8 inch Fresnel on spot at " & throw & " is " & FormatNumber(diameter, 1) & " feet.")
    End If
    If chkFlood = 1 Then
        diameter = throw * 1.35
        MsgBox ("The diameter of a 8 inch Fresnel on flood at " & throw & " is " & FormatNumber(diameter, 1) & " feet.")
    End If
End If
End Sub

Private Sub cmdReturn_Click()
'Returns to main menu
frmMainMenu.Show
frmCircle.Hide
End Sub

Private Sub cmdS4_Click()
diameter = 0
throw = txtThrow
'check boxes determine what degree of source four it is and then determine the size of the circle.
If chk19.Value = 1 Then
    diameter = throw * 0.32
    MsgBox ("The diameter of the circle " & throw & " feet away from a 19 degree Source Four is " & FormatNumber(diameter, 1) & " feet.")
End If
If chk26.Value = 1 Then
    diameter = throw * 0.45
    MsgBox ("The diameter of the circle " & throw & " feet away from a 26 degree Source Four is " & FormatNumber(diameter, 1) & " feet.")
End If
If chk36.Value = 1 Then
    diameter = throw * 0.61
    MsgBox ("The diameter of the circle " & throw & " feet away from a 36 degree Source Four is " & FormatNumber(diameter, 1) & " feet.")
End If
End Sub
