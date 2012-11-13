VERSION 5.00
Begin VB.Form frmpolicebac 
   BackColor       =   &H00B75C3E&
   Caption         =   "BAC Calculator"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H0003CCE9&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1800
      Width           =   2535
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2235
      ScaleWidth      =   6435
      TabIndex        =   11
      Top             =   4800
      Width           =   6495
   End
   Begin VB.CommandButton cmdjoe 
      BackColor       =   &H0003CCE9&
      Caption         =   "Continue on your tour de st. joe"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CommandButton cmdreturn 
      BackColor       =   &H0003CCE9&
      Caption         =   "Return to Police "
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdcalculate 
      BackColor       =   &H0003CCE9&
      Caption         =   "Calculate your Blood Alcohol Content (BAC)"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox txttime 
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox txtweight 
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txtpercent 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtoz 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lbltime 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter the time in hours spent drinking"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label lblweight 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter your body weight in pounds"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label lblpercent 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter the alcohol percent"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label lblounces 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter the number of ounces consumed"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmpolicebac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
    'Project name:  Tour De St. Joe
    'Form:  frmpolicebac, "BAC"
    'Author:  Brooke and Josh
    'Date:  3/26/08
    'Objective: To calculate blood alcohol levels by asking a user for key information and then putting that into a formula.

Private Sub cmdcalculate_Click()

    
    Dim ounces As Single
    Dim percent As Single
    Dim weight As Single
    Dim time As Single
    Dim bac As Single
    
    ounces = txtoz.Text
    percent = txtpercent.Text
    weight = txtweight.Text
    time = txttime.Text
    bac = picresults.Picture
    
    'calculation
    bac = (((ounces * percent * 7.5) / (1000 * weight)) - (0.0015 * time)) * 10
    
    Select Case bac
        Case Is <= 0
            picresults.Print "Your B.A.C. is " & FormatNumber(bac, 2) & " You are sober!"
        Case 0 To 0.04
            picresults.Print "Your B.A.C. is " & FormatNumber(bac, 2) & " You are leagally buzzed."
        Case 0.04 To 0.08
            picresults.Print "Your B.A.C. is " & FormatNumber(bac, 2) & " You have impared judgement."
        Case 0.08 To 0.13
            picresults.Print "Your B.A.C. is " & FormatNumber(bac, 2) & " You have poor coordination."
        Case 0.13 To 0.19
            picresults.Print "Your B.A.C. is " & FormatNumber(bac, 2) & " You have no coordination."
        Case 0.19 To 0.29
            picresults.Print "Your B.A.C. is " & FormatNumber(bac, 2) & " You need medical assistance."
        Case Is >= 0.3
            picresults.Print "Your B.A.C. is " & FormatNumber(bac, 2) & " which is so high that you will probably die."
    End Select
    

End Sub

Private Sub cmdclear_Click()

    picresults.Cls
 
    txtoz.Text = ""
    txtpercent.Text = ""
    txtweight.Text = ""
    txttime.Text = ""
    
End Sub

Private Sub cmdjoe_Click()

    frmjoetown.Show
    frmpolicebac.Hide

End Sub

Private Sub cmdreturn_Click()

    frmpolice.Show
    frmpolicebac.Hide

End Sub
