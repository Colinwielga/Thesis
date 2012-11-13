VERSION 5.00
Begin VB.Form frmMonthAverage 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16395
   BeginProperty Font 
      Name            =   "MV Boli"
      Size            =   8.25
      Charset         =   1
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   16395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBackToStates 
      BackColor       =   &H00FF0000&
      Caption         =   "Return to State Average Page"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7680
      Width           =   4215
   End
   Begin VB.CommandButton cmdBackBeg 
      BackColor       =   &H00FF0000&
      Caption         =   "Return to Home Page"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Width           =   4215
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   26.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8400
      Width           =   3015
   End
   Begin VB.PictureBox picResults3 
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8160
      ScaleHeight     =   1875
      ScaleWidth      =   7035
      TabIndex        =   13
      Top             =   1560
      Width           =   7095
   End
   Begin VB.CommandButton cmdDec 
      BackColor       =   &H000000FF&
      Caption         =   "December"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdNov 
      BackColor       =   &H000000FF&
      Caption         =   "November"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdOct 
      BackColor       =   &H000000FF&
      Caption         =   "October"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdSept 
      BackColor       =   &H000000FF&
      Caption         =   "September"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdAug 
      BackColor       =   &H000000FF&
      Caption         =   "August"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdJune 
      BackColor       =   &H000000FF&
      Caption         =   "June"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdMay 
      BackColor       =   &H000000FF&
      Caption         =   "May"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdJuly 
      BackColor       =   &H000000FF&
      Caption         =   "July"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdApril 
      BackColor       =   &H000000FF&
      Caption         =   "April"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdMarch 
      BackColor       =   &H000000FF&
      Caption         =   "March"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdFeb 
      BackColor       =   &H000000FF&
      Caption         =   "February"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdJan 
      BackColor       =   &H000000FF&
      Caption         =   "January"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblAverage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select the Month You Want to See the Average         Unemployment Rate in the United States"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   12
      Top             =   360
      Width           =   6975
   End
End
Attribute VB_Name = "frmMonthAverage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Unemployment in the 2008 Recession
'Form Name: Month Average
'Author: Josh Overman
'March 22th 2008
'Objective: To anaylize the unemployment on a monthly basis

'Declare all these varaibles as global for this form
Dim Sum As Double
Dim I As Integer
Dim Average As Double
'This buttom will take the average of the unemployment rates 51 (D.C.) total and display it.
'each button in this form will do the same thing,
'the only difference will be the number in the table(I, #)
'which will correspond to the month we are finding

Private Sub cmdApril_Click()
'Clear the picture box
picResults3.Cls
'Initialize the variables
Sum = 0
'Run through the table and add the unemployment rates for the month up
'The 4 means the 4th month or April
For I = 1 To CTR1
    Sum = Sum + Table(I, 4)
Next I
'Take the average of the sum of the unemployment rates
'Divide by 51 for the 50 states and the District of Columbia
Average = Sum / 51
'Print the average unemplyment rate for the month
picResults3.Print "The Average Unemployment Rate in the U.S. for the "
picResults3.Print "Month of April was "; FormatNumber(Average, 2); "%"
End Sub

Private Sub cmdAug_Click()
'For the month of August
picResults3.Cls
Sum = 0
For I = 1 To CTR1
    Sum = Sum + Table(I, 8)
Next I
Average = Sum / 51
picResults3.Print "The Average Unemployment Rate in the U.S. for the "
picResults3.Print "Month of August was "; FormatNumber(Average, 2); "%"
End Sub

Private Sub cmdBackBeg_Click()
'Naviagte ffrom the average form to the Start up menu
frmStartUp.Show
frmMonthAverage.Hide
End Sub

Private Sub cmdBackToStates_Click()
'Nagivate from the average form to the individual state form
frmStates.Show
frmMonthAverage.Hide
End Sub

Private Sub cmdDec_Click()
'For the month of December
picResults3.Cls
Sum = 0
For I = 1 To CTR1
    Sum = Sum + Table(I, 12)
Next I
Average = Sum / 51
picResults3.Print "The Average Unemployment Rate in the U.S. for the "
picResults3.Print "Month of December was "; FormatNumber(Average, 2); "%"
End Sub

Private Sub cmdFeb_Click()
'For the month of February
picResults3.Cls
Sum = 0
For I = 1 To CTR1
    Sum = Sum + Table(I, 2)
Next I
Average = Sum / 51
picResults3.Print "The Average Unemployment Rate in the U.S. for the "
picResults3.Print "Month of February was "; FormatNumber(Average, 2); "%"
End Sub

Private Sub cmdJan_Click()
'For the month of January
picResults3.Cls
Sum = 0
For I = 1 To CTR1
    Sum = Sum + Table(I, 1)
Next I
Average = Sum / 51
picResults3.Print "The Average Unemployment Rate in the U.S. for the "
picResults3.Print "Month of January was "; FormatNumber(Average, 2); "%"

End Sub

Private Sub cmdJuly_Click()
'For the month of July
picResults3.Cls
Sum = 0
For I = 1 To CTR1
    Sum = Sum + Table(I, 7)
Next I
Average = Sum / 51
picResults3.Print "The Average Unemployment Rate in the U.S. for the "
picResults3.Print "Month of July was "; FormatNumber(Average, 2); "%"
End Sub

Private Sub cmdJune_Click()
'For the month of June
picResults3.Cls
Sum = 0
For I = 1 To CTR1
    Sum = Sum + Table(I, 6)
Next I
Average = Sum / 51
picResults3.Print "The Average Unemployment Rate in the U.S. for the "
picResults3.Print "Month of June was "; FormatNumber(Average, 2); "%"
End Sub

Private Sub cmdMarch_Click()
'For the month of March
picResults3.Cls
Sum = 0
For I = 1 To CTR1
    Sum = Sum + Table(I, 3)
Next I
Average = Sum / 51
picResults3.Print "The Average Unemployment Rate in the U.S. for the "
picResults3.Print "Month of March was "; FormatNumber(Average, 2); "%"
End Sub

Private Sub cmdMay_Click()
'For the month of May
picResults3.Cls
Sum = 0
For I = 1 To CTR1
    Sum = Sum + Table(I, 5)
Next I
Average = Sum / 51
picResults3.Print "The Average Unemployment Rate in the U.S. for the "
picResults3.Print "Month of May was "; FormatNumber(Average, 2); "%"
End Sub

Private Sub cmdNov_Click()
'For the month of November
picResults3.Cls
Sum = 0
For I = 1 To CTR1
    Sum = Sum + Table(I, 11)
Next I
Average = Sum / 51
picResults3.Print "The Average Unemployment Rate in the U.S. for the "
picResults3.Print "Month of November was "; FormatNumber(Average, 2); "%"
End Sub

Private Sub cmdOct_Click()
'For the month of October
picResults3.Cls
Sum = 0
For I = 1 To CTR1
    Sum = Sum + Table(I, 10)
Next I
Average = Sum / 51
picResults3.Print "The Average Unemployment Rate in the U.S. for the "
picResults3.Print "Month of October was "; FormatNumber(Average, 2); "%"
End Sub

'ends the program
Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSept_Click()
'for the month of september
picResults3.Cls
Sum = 0
For I = 1 To CTR1
    Sum = Sum + Table(I, 9)
Next I
Average = Sum / 51
picResults3.Print "The Average Unemployment Rate in the U.S. for the "
picResults3.Print "Month of September was "; FormatNumber(Average, 2); "%"
End Sub

