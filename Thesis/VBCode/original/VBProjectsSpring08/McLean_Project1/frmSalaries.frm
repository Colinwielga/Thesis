VERSION 5.00
Begin VB.Form frmSalaries 
   Caption         =   "Salaries"
   ClientHeight    =   11772
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   15192
   LinkTopic       =   "Form1"
   ScaleHeight     =   11772
   ScaleWidth      =   15192
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9840
      Width           =   3012
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6960
      Width           =   3012
   End
   Begin VB.CommandButton cmdContents 
      Caption         =   "Back to the Table of Contents"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   12000
      TabIndex        =   25
      Top             =   9960
      Width           =   2892
   End
   Begin VB.CommandButton cmdMedium 
      Caption         =   "Medium Firm/Company"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5640
      Width           =   1812
   End
   Begin VB.CommandButton cmdSmall 
      Caption         =   "Small Firm/Company"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5640
      Width           =   1812
   End
   Begin VB.CommandButton cmdLarge 
      Caption         =   "Large Firm/Company"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5640
      Width           =   1812
   End
   Begin VB.CommandButton cmdPartner 
      Caption         =   "Partner"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3960
      Width           =   1932
   End
   Begin VB.CommandButton cmdDirector 
      Caption         =   "Director"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3960
      Width           =   1812
   End
   Begin VB.CommandButton cmdManager 
      Caption         =   "Manager"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3960
      Width           =   1812
   End
   Begin VB.CommandButton cmdSenior 
      Caption         =   "Senior"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3960
      Width           =   1812
   End
   Begin VB.CommandButton cmdOnetoThree 
      Caption         =   "1 to 3 years"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3960
      Width           =   1812
   End
   Begin VB.CommandButton cmdOne 
      Caption         =   "Up to 1 year"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3960
      Width           =   1812
   End
   Begin VB.CommandButton cmdCFO 
      Caption         =   "Chief Financial Officer"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
      Width           =   1932
   End
   Begin VB.CommandButton cmdController 
      Caption         =   "Controller"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   1812
   End
   Begin VB.CommandButton cmdTreasurer 
      Caption         =   "Treasurer"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1812
   End
   Begin VB.CommandButton cmdAnalyst 
      Caption         =   "Financial Analyst"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1812
   End
   Begin VB.CommandButton cmdTax 
      Caption         =   "Tax Accountant"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   1812
   End
   Begin VB.CommandButton cmdAudit 
      Caption         =   "Audit Accountant"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   1812
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1932
      Left            =   4080
      ScaleHeight     =   1884
      ScaleWidth      =   6564
      TabIndex        =   1
      Top             =   7800
      Width           =   6612
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Size of the Company"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   5040
      TabIndex        =   24
      Top             =   4800
      Width           =   4812
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Experience"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   5880
      TabIndex        =   23
      Top             =   3240
      Width           =   3132
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   6360
      TabIndex        =   22
      Top             =   1560
      Width           =   2052
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Salary Calculator"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1212
      Left            =   2880
      TabIndex        =   21
      Top             =   120
      Width           =   9012
   End
   Begin VB.Image Image2 
      Height          =   12216
      Left            =   0
      Picture         =   "frmSalaries.frx":0000
      Top             =   0
      Width           =   16488
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Size of the Company"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1212
      Left            =   0
      TabIndex        =   5
      Top             =   4800
      Width           =   6612
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Experience"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1212
      Left            =   0
      TabIndex        =   4
      Top             =   3240
      Width           =   6612
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1212
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   6612
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Salary Calculator"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1212
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   6612
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Salary Calculator"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5412
   End
   Begin VB.Image Image1 
      Height          =   14400
      Left            =   0
      Picture         =   "frmSalaries.frx":2C2E8
      Top             =   0
      Width           =   15408
   End
End
Attribute VB_Name = "frmSalaries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Accounting Project
'Introduction Form
'Tony McLean
'3.31.2008
'The purpose of this form is to provide the user with
'a highly interactive interface that will allow for the user
'to choose a profession that he or she finds interesting and
'calculate a corresponding salary for that position, experience, etc.
Option Explicit
'This subroutine disables certain command buttons when this buttion is chosen
Private Sub cmdAnalyst_Click()
    cmdTax.Enabled = False
    cmdAudit.Enabled = False
    cmdTreasurer.Enabled = False
    cmdController.Enabled = False
    cmdCFO.Enabled = False
    cmdDirector.Enabled = False
    cmdPartner.Enabled = False
    picResults.Print "Financial Analyst"
End Sub
'This subroutine disables certain command buttons when this buttion is chosen
Private Sub cmdAudit_Click()
    cmdTax.Enabled = False
    cmdAnalyst.Enabled = False
    cmdTreasurer.Enabled = False
    cmdController.Enabled = False
    cmdCFO.Enabled = False
    picResults.Print "Audit Accountant"
End Sub
'The following is a lengthy "if, then" statement that calculates and displays a salary
'figure from the selections made by the user
Private Sub cmdCalculate_Click()
    picResults.Print "*******************************************************************"
    If cmdController.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $64,250 to 83,750"
    End If
    If cmdController.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $88,500 to 117,000"
    End If
    If cmdController.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $126,750 to 161,500"
    End If
    If cmdTreasurer.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $91,000 to 122,250"
    End If
    If cmdTreasurer.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $121,250 to 170,250"
    End If
    If cmdTreasurer.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $257,500 to 370,500"
    End If
    If cmdCFO.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $91,000 to 122,250"
    End If
    If cmdCFO.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $121,250 to 170,250"
    End If
    If cmdCFO.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $257,500 to 370,500"
    End If
    If cmdAudit.Enabled = True And cmdOne.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $47,500 to 57,500"
    End If
    If cmdAudit.Enabled = True And cmdOnetoThree.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $54,500 to 66,250"
    End If
    If cmdAudit.Enabled = True And cmdSenior.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $64,000 to 83,000"
    End If
    If cmdAudit.Enabled = True And cmdManager.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $80,000 to 106,250"
    End If
    If cmdAudit.Enabled = True And cmdDirector.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $98,750 to 151,500"
    End If
    If cmdAudit.Enabled = True And cmdPartner.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $151,500+"
    End If
    If cmdAudit.Enabled = True And cmdOne.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $41,500 to 51,000"
    End If
    If cmdAudit.Enabled = True And cmdOnetoThree.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $48,750 to 59,500"
    End If
    If cmdAudit.Enabled = True And cmdSenior.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $57,500 to 76,250"
    End If
    If cmdAudit.Enabled = True And cmdManager.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $74,250 to 93,500"
    End If
    If cmdAudit.Enabled = True And cmdDirector.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $88,250 to 129,250"
    End If
    If cmdAudit.Enabled = True And cmdPartner.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: 129,250+"
    End If
    If cmdAudit.Enabled = True And cmdOne.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $40,000 to 47,250"
    End If
    If cmdAudit.Enabled = True And cmdOnetoThree.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $44,750 to 53,250"
    End If
    If cmdAudit.Enabled = True And cmdSenior.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $52,000 to 66,500"
    End If
    If cmdAudit.Enabled = True And cmdManager.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $66,500 to 82,000"
    End If
    If cmdAudit.Enabled = True And cmdDirector.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $80,750 to 105,500"
    End If
    If cmdAudit.Enabled = True And cmdPartner.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: 105,500+"
    End If
    If cmdTax.Enabled = True And cmdOne.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $47,500 to 57,500"
    End If
    If cmdTax.Enabled = True And cmdOnetoThree.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $54,500 to 66,250"
    End If
    If cmdTax.Enabled = True And cmdSenior.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $64,000 to 83,000"
    End If
    If cmdTax.Enabled = True And cmdManager.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $80,000 to 106,250"
    End If
    If cmdTax.Enabled = True And cmdDirector.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $98,750 to 151,500"
    End If
    If cmdTax.Enabled = True And cmdPartner.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $151,500+"
    End If
    If cmdTax.Enabled = True And cmdOne.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $41,500 to 51,000"
    End If
    If cmdTax.Enabled = True And cmdOnetoThree.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $48,750 to 59,500"
    End If
    If cmdTax.Enabled = True And cmdSenior.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $57,500 to 76,250"
    End If
    If cmdTax.Enabled = True And cmdManager.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $74,250 to 93,500"
    End If
    If cmdTax.Enabled = True And cmdDirector.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $88,250 to 129,250"
    End If
    If cmdTax.Enabled = True And cmdPartner.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: 129,250+"
    End If
    If cmdTax.Enabled = True And cmdOne.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008:$40,000 to 47,250"
    End If
    If cmdTax.Enabled = True And cmdOnetoThree.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $44,750 to 53,250"
    End If
    If cmdTax.Enabled = True And cmdSenior.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $52,000 to 66,500"
    End If
    If cmdTax.Enabled = True And cmdManager.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $66,500 to 82,000"
    End If
    If cmdTax.Enabled = True And cmdDirector.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $80,750 to 105,500"
    End If
    If cmdTax.Enabled = True And cmdPartner.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: 105,500+"
    End If
    If cmdAnalyst.Enabled = True And cmdOne.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $38,250 to 47,500"
    End If
    If cmdAnalyst.Enabled = True And cmdOnetoThree.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $46,500 to 61,250"
    End If
    If cmdAnalyst.Enabled = True And cmdSenior.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $61,000 to 77,250"
    End If
    If cmdAnalyst.Enabled = True And cmdManager.Enabled = True And cmdLarge.Enabled = True Then
                picResults.Print "Salary for 2008: $74,750 to 99,000"
    End If
    If cmdAnalyst.Enabled = True And cmdOne.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $36,500 to 44,000"
    End If
    If cmdAnalyst.Enabled = True And cmdOnetoThree.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $43,500 to 55,750"
    End If
    If cmdAnalyst.Enabled = True And cmdSenior.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $55,750 to 70,000"
    End If
    If cmdAnalyst.Enabled = True And cmdManager.Enabled = True And cmdMedium.Enabled = True Then
                picResults.Print "Salary for 2008: $67,500 to 85,000"
    End If
    If cmdAnalyst.Enabled = True And cmdOne.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $34,400 to 40,500"
    End If
    If cmdAnalyst.Enabled = True And cmdOnetoThree.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $39,500 ot 50,750"
    End If
    If cmdAnalyst.Enabled = True And cmdSenior.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $48,500 to 61,250"
    End If
    If cmdAnalyst.Enabled = True And cmdManager.Enabled = True And cmdSmall.Enabled = True Then
                picResults.Print "Salary for 2008: $80,000 to 106,250"
    End If
End Sub
'This subroutine disables certain command buttons when this buttion is chosen
Private Sub cmdCFO_Click()
    cmdTax.Enabled = False
    cmdAudit.Enabled = False
    cmdTreasurer.Enabled = False
    cmdController.Enabled = False
    cmdAnalyst.Enabled = False
    cmdOne.Enabled = False
    cmdOnetoThree.Enabled = False
    cmdSenior.Enabled = False
    cmdManager.Enabled = False
    cmdDirector.Enabled = False
    cmdPartner.Enabled = False
    picResults.Print "Chief Financial Officer"
End Sub
'This subroutine allows the user to return to the contents page of the program
Private Sub cmdContents_Click()
    frmProfessions.Hide
    frmFirms.Hide
    frmSalaries.Hide
    frmDidKnow.Hide
    frmContents.Show
    frmIntroduction.Hide
End Sub
'This subroutine disables certain command buttons when this buttion is chosen
Private Sub cmdController_Click()
    cmdTax.Enabled = False
    cmdAudit.Enabled = False
    cmdTreasurer.Enabled = False
    cmdAnalyst.Enabled = False
    cmdCFO.Enabled = False
    cmdOne.Enabled = False
    cmdOnetoThree.Enabled = False
    cmdSenior.Enabled = False
    cmdManager.Enabled = False
    cmdDirector.Enabled = False
    cmdPartner.Enabled = False
    picResults.Print "Corporate Controller"
End Sub
'This subroutine disables certain command buttons when this buttion is chosen
Private Sub cmdDirector_Click()
    cmdOnetoThree.Enabled = False
    cmdSenior.Enabled = False
    cmdManager.Enabled = False
    cmdOne.Enabled = False
    cmdPartner.Enabled = False
    picResults.Print "Director"
End Sub
'This subroutine disables certain command buttons when this buttion is chosen
Private Sub cmdLarge_Click()
    cmdMedium.Enabled = False
    cmdSmall.Enabled = False
    picResults.Print "Large Company"
End Sub
'This subroutine disables certain command buttons when this buttion is chosen
Private Sub cmdManager_Click()
    cmdOnetoThree.Enabled = False
    cmdSenior.Enabled = False
    cmdOne.Enabled = False
    cmdDirector.Enabled = False
    cmdPartner.Enabled = False
    picResults.Print "Manager"
End Sub
'This subroutine disables certain command buttons when this buttion is chosen
Private Sub cmdMedium_Click()
    cmdLarge.Enabled = False
    cmdSmall.Enabled = False
    picResults.Print "Medium Company"
End Sub
'This subroutine disables certain command buttons when this buttion is chosen
Private Sub cmdOne_Click()
    cmdOnetoThree.Enabled = False
    cmdSenior.Enabled = False
    cmdManager.Enabled = False
    cmdDirector.Enabled = False
    cmdPartner.Enabled = False
    picResults.Print "Up to 1 year of experience"
End Sub
'This subroutine disables certain command buttons when this buttion is chosen
Private Sub cmdOnetoThree_Click()
    cmdOne.Enabled = False
    cmdSenior.Enabled = False
    cmdManager.Enabled = False
    cmdDirector.Enabled = False
    cmdPartner.Enabled = False
    picResults.Print "1 to 3 years of experience"
End Sub
'This subroutine disables certain command buttons when this buttion is chosen
Private Sub cmdPartner_Click()
    cmdOnetoThree.Enabled = False
    cmdSenior.Enabled = False
    cmdManager.Enabled = False
    cmdDirector.Enabled = False
    cmdOne.Enabled = False
    picResults.Print "Partner"
End Sub
'This subroutine resets all of the command buttons back to an enabled state for the user
'to use again.
Private Sub cmdReset_Click()
    cmdTax.Enabled = True
    cmdAudit.Enabled = True
    cmdAnalyst.Enabled = True
    cmdController.Enabled = True
    cmdTreasurer.Enabled = True
    cmdCFO.Enabled = True
    cmdOne.Enabled = True
    cmdOnetoThree.Enabled = True
    cmdSenior.Enabled = True
    cmdManager.Enabled = True
    cmdDirector.Enabled = True
    cmdPartner.Enabled = True
    cmdLarge.Enabled = True
    cmdMedium.Enabled = True
    cmdSmall.Enabled = True
    picResults.Cls
End Sub
'This subroutine disables certain command buttons when this buttion is chosen
Private Sub cmdSenior_Click()
    cmdOnetoThree.Enabled = False
    cmdOne.Enabled = False
    cmdManager.Enabled = False
    cmdDirector.Enabled = False
    cmdPartner.Enabled = False
    picResults.Print "Senior"
End Sub
'This subroutine disables certain command buttons when this buttion is chosen
Private Sub cmdSmall_Click()
    cmdMedium.Enabled = False
    cmdLarge.Enabled = False
    picResults.Print "Small Company"
End Sub
'This subroutine disables certain command buttons when this buttion is chosen
Private Sub cmdTax_Click()
    cmdAudit.Enabled = False
    cmdAnalyst.Enabled = False
    cmdTreasurer.Enabled = False
    cmdController.Enabled = False
    cmdCFO.Enabled = False
    picResults.Print "Tax Accountant"
End Sub
'This subroutine disables certain command buttons when this buttion is chosen
Private Sub cmdTreasurer_Click()
    cmdTax.Enabled = False
    cmdAudit.Enabled = False
    cmdAnalyst.Enabled = False
    cmdController.Enabled = False
    cmdCFO.Enabled = False
    cmdOne.Enabled = False
    cmdOnetoThree.Enabled = False
    cmdSenior.Enabled = False
    cmdManager.Enabled = False
    cmdDirector.Enabled = False
    cmdPartner.Enabled = False
    picResults.Print "Treasurer"
End Sub
