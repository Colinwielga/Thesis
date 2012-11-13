VERSION 5.00
Begin VB.Form YearsDancedForm 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   5280
      ScaleHeight     =   4515
      ScaleWidth      =   4635
      TabIndex        =   5
      Top             =   720
      Width           =   4695
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Print Number of Years Danced By each Dancer"
      Height          =   975
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Put in Order from Most Years Danced to Least Years Danced"
      Enabled         =   0   'False
      Height          =   975
      Left            =   1200
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate the Average Number of Years Danced"
      Enabled         =   0   'False
      Height          =   975
      Left            =   1230
      TabIndex        =   2
      Top             =   3750
      Width           =   1935
   End
   Begin VB.CommandButton cmdMainForm 
      BackColor       =   &H000000FF&
      Caption         =   "Main Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6240
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lauren Doyscher"
      Height          =   255
      Left            =   9360
      TabIndex        =   6
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "YearsDancedForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: SophmoreDancers (VBProject.vbp)
'Form Name: YearsDancedForm (YearsDanced.frm)
'Author: Lauren Doyscher
'Date Written: 10/27/03
'This form fills two arrays from a file.  It is used to show how
'many years each dancer has danced and then calculate the average
'amount of years the dancers have danced.
'Dimensions These variables for the whole form
Option Explicit
Dim A As Integer
Dim YearsDanced(1 To 12) As Integer
Dim Dancer(1 To 12) As String
'This Button fills two arrays: Years in Dance and Dancer.
'It prints the number of years danced and the dancer's name next to it.
'Enables the Order and Calculation buttons
Private Sub cmdArray_Click()
picResults.Cls
Open PATH & "YearsDanced.txt" For Input As #1
picResults.Print "Years In Dance"; Tab(25); "Dancer"
picResults.Print "__________________________________________"
    For A = 1 To 12
        Input #1, YearsDanced(A), Dancer(A)
        picResults.Print YearsDanced(A); Tab(25); Dancer(A)
    Next A
cmdOrder.Enabled = True
cmdCalc.Enabled = True
Close #1
End Sub

Private Sub cmdCalc_Click()
Dim Sum As Integer
Dim Average As Single
    For A = 1 To 12
        Sum = Sum + YearsDanced(A)
    Next A
Average = Sum / 12
picResults.Print "  "
picResults.Print "The average amount of years danced ="
picResults.Print Average; " Years"
End Sub

Private Sub cmdMainForm_Click()
'Brings you back to Main Page
YearsDancedForm.Hide
MainForm.Show
End Sub

'This button will put the dancers in order from the
'girl that has danced the most amount of years to the
'girl that has danced the least amount of years.
'If the dancers have danced the same amount of years, the program
'will list the dancers in alphabetical order.
Private Sub cmdOrder_Click()
Dim Pass As Integer
Dim Comp As Integer
Dim Temp1 As Integer
Dim Temp2 As String
picResults.Cls
picResults.Print "Years In Dance"; Tab(25); "Dancer"
picResults.Print "__________________________________________"
    For Pass = 1 To 11
        For Comp = 1 To 12 - Pass
            If YearsDanced(Comp) < YearsDanced(Comp + 1) Then
                Temp1 = YearsDanced(Comp)
                YearsDanced(Comp) = YearsDanced(Comp + 1)
                YearsDanced(Comp + 1) = Temp1
                Temp2 = Dancer(Comp)
                Dancer(Comp) = Dancer(Comp + 1)
                Dancer(Comp + 1) = Temp2
            End If
            If YearsDanced(Comp) = YearsDanced(Comp + 1) Then
                If Dancer(Comp) > Dancer(Comp + 1) Then
                    Temp2 = Dancer(Comp)
                    Dancer(Comp) = Dancer(Comp + 1)
                    Dancer(Comp + 1) = Temp2
                End If
            End If
        Next Comp
    Next Pass
    For A = 1 To 12
        picResults.Print YearsDanced(A); Tab(25); Dancer(A)
    Next A
End Sub
Private Sub cmdQuit_Click()
End
End Sub
