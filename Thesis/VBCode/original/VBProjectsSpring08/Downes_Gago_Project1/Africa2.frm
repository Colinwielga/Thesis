VERSION 5.00
Begin VB.Form Africa2 
   BackColor       =   &H00008080&
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H00008080&
      Height          =   3615
      Left            =   3720
      ScaleHeight     =   3555
      ScaleWidth      =   4755
      TabIndex        =   5
      Top             =   1680
      Width           =   4815
   End
   Begin VB.TextBox txtYears 
      BackColor       =   &H0080FF80&
      Height          =   975
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   4815
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H00008080&
      Caption         =   "Calculate Population in X Number of Years"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   3375
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Calculations for Top 10 Populated Countries in Africa"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label LblYears 
      BackColor       =   &H00008080&
      Caption         =   "Enter the Number of Years for Estimated Population of Top 10 Countries =================>"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "Africa2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: The Globe Trotter Experience
'Form name: Africa2.frm
'Author: Marta Gago & Brian Downes
'Date Written: Thursday March 27th, 2008
'Objective of form:  This form gives the user the ability to calculate population growth

Option Explicit
'Hides the Africa2 Form and then shows the Africa form
Private Sub cmdBack_Click()
Africa2.Hide
Africa.Show
End Sub
'Calculation of Population growth of the top 10 countries with highest population
Private Sub cmdCalculate_Click()
'dim the variables for calculation
Dim Years As Single, country(1 To 10) As String, population(1 To 90000000) As Single
Dim GrowthRate(1 To 10) As Single, k As Single
'Counter set at zero
ctr = 0
'Years is the amount of years the user enters
Years = txtYears.Text
'Open the data file with the population statistics
    Open App.Path & "\Population.txt" For Input As #1
    'using the Do While loop, the data is put into an array
        Do While Not EOF(1)
            ctr = ctr + 1
                Input #1, country(ctr), population(ctr), GrowthRate(ctr)
        Loop
'printing of the labels for the data
picResults.Print "Country", "Population in "; Years; " Years:"
picResults.Print "*********************************************************"
'using a For/Next loop, the population after the amount of years entered is calculated (i)
For j = 1 To ctr
            k = (population(j) * GrowthRate(j) ^ Years) + population(j)
                    'printing of the countries and their growth after amount of years with a For/Next loop
                picResults.Print country(j), Tab(20), Int(Round(k)) 'Int takes off the decimals from the population and Round
                                                                    'rounds the number to its nearest whole number
Next j

Close   'Close data array
End Sub
