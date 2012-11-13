VERSION 5.00
Begin VB.Form frmCalculator 
   BackColor       =   &H00404040&
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   18
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   16
      Top             =   4560
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   15
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   14
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Calculate Total Costs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   13
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox txtTransportation 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtHousing 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
   Begin VB.PictureBox picAverage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtYears 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtTuition 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox cbo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Text            =   "Pick University"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdAverage 
      Caption         =   "Calculate Average"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00404040&
      Caption         =   "Total:"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label lblTransportation 
      BackColor       =   &H00404040&
      Caption         =   "Transportation Costs per year:"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label lblHousing 
      BackColor       =   &H00404040&
      Caption         =   "Housing Costs per year:"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label lblYears 
      BackColor       =   &H00404040&
      Caption         =   "Number of Years in College:"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00404040&
      Caption         =   "Average Tuition Costs:"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project Name: College Bound
' Form Name: Calculator
' Authors: Magdalena Adamczyk & Leszek Nowacki
' 9-25 March 2009
' This form is designed to let the user see calculate the total costs of attending collage
' as well as calculate the average costs of tuition
Option Explicit
Dim ctr As Single, I As Single, total As Single, sum As Single, average As Single, A As Single, B As Single, C As Single, D As Single


Private Sub cmdClear_Click() ' this button clears the form in order for the user to be able to start a new calculation
txtTuition.Text = " "
txtYears.Text = " "
txtHousing.Text = " "
txtTransportation.Text = " "
picResults.Cls

End Sub

Private Sub cmdLoad_Click() ' this button loads the file in to arrays and loads the content of the scroll down list(combobox)

Open App.Path & "\list.txt" For Input As #1 ' when the button is clicked the file with the list of the universities is loaded into parallel arrays
Do Until EOF(1)
    ctr = ctr + 1
    Input #1, university(ctr), tuition(ctr), sort(ctr), major1(ctr), major2(ctr), major3(ctr)
cbo1.AddItem university(ctr) ' this comment loads the name of each university to the combobox
Loop
Close #1 ' closing the file
cmdAverage.Enabled = True ' enabling the button for calculating the average tuition
cmdOK.Enabled = True ' enabling the button for choising the university
End Sub
Private Sub cmdAverage_Click() ' this button allows the user to calculate the average tuition costs from all of the universities in the database

For I = 1 To ctr 'at this moment the programs goes through the array that holds tuition and sums up the tuitions
    sum = sum + tuition(I) ' sum of the tuitions is equal to the previous sum plus current tuition
Next I
average = sum / ctr 'the average tuition is egual to the total sum devided by number of universities
picAverage.Print FormatCurrency(average) ' displayig the average
End Sub

Private Sub cmdOK_Click() 'this button allows the user to confirm the name of the university that he or she has choisen in the combobox list
For I = 1 To ctr
    If cbo1.Text = university(I) Then ' for the appriopiate name of the university the tuition amount athat is assigned to it will be displayed in a text box
        txtTuition.Text = FormatCurrency(tuition(I))
    End If
Next I
End Sub

Private Sub cmdTotal_Click() ' this button lets the user calculate all of the costs of going to a univerity
picResults.Cls

A = txtTuition.Text
B = txtYears.Text
C = txtHousing.Text
D = txtTransportation.Text


total = A * B + C * B + D * B ' this formula takes the amount of tuition and costs of housing and transportation, multiplys it by the number of years
picResults.Print FormatCurrency(total) 'then the total costs are displayed
End Sub


Private Sub cmdBack_Click() ' this button alows the user to go back to the main form
frmCalculator.Hide
frmUniversitySearch.Show
End Sub

Private Sub cmdQuit_Click() ' this button ends the program
End
End Sub

