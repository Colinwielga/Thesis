VERSION 5.00
Begin VB.Form frmFormMales 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   12255
   ClientLeft      =   4035
   ClientTop       =   2595
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   12255
   ScaleWidth      =   8040
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Get Data from Query"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   9720
      Width           =   5415
   End
   Begin VB.CommandButton cmdLowtoHighWeight 
      Caption         =   "Low to High Weight"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   6
      Top             =   10680
      Width           =   1695
   End
   Begin VB.CommandButton cmdLowtoHighBMI 
      Caption         =   "Low to High BMI"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      TabIndex        =   5
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton cmdTalltoShort 
      Caption         =   "Shortest to tallest"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      TabIndex        =   4
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Survey"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   3
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      TabIndex        =   2
      Top             =   10680
      Width           =   1695
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      TabIndex        =   1
      Top             =   10680
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   6375
      Left            =   1080
      ScaleHeight     =   6315
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   1560
      Width           =   5415
   End
   Begin VB.Label lblIntro 
      BackColor       =   &H0080FFFF&
      Caption         =   "NOW LET'S FIND THE PERFECT            MALE FOR YOU!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      TabIndex        =   7
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmFormMales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDisplay_Click()
picResults.Cls

Open App.Path & "\Males.txt" For Input As #1

Dim Name(1 To 100) As String
Dim Height(1 To 100) As Integer
Dim Weight(1 To 100) As Integer
Dim BMI(1 To 100) As Single
Dim Pos As Integer

Do While Not EOF(1)
    Pos = Pos + 1
    Input #1, Name(Pos), Height(Pos), Weight(Pos), BMI(Pos)
    picResults.Print Name(Pos), Height(Pos); " inches", Weight(Pos); " lbs", BMI(Pos)
Loop
Close #1
End Sub

Private Sub cmdEnter_Click()
Dim DB As Database
Dim QD As QueryDef

Dim RS As Recordset2

Set DB = OpenDatabase(App.Path & "\Weight_Range.accdb")
Set QD = DB.QueryDefs("SpecificMales")
QD.Parameters(0) = InputBox("Please enter preferred minimum height in inches.")
QD.Parameters(1) = InputBox("Please enter preferred maximum height in inches.")
QD.Parameters(2) = InputBox("Please enter preferred minimum weight.")
QD.Parameters(3) = InputBox("Please enter preferred maximum weight.")

Set RS = QD.OpenRecordset()
picResults.Cls
picResults.Print "These are the men that met your criteria."
picResults.Print "***********************************************************"
Do Until RS.EOF
    picResults.Print RS![FirstName], RS![Height]; " inches", RS![Weight]; " lbs", RS![Body Index]
    RS.MoveNext
Loop
RS.Close
DB.Close
End Sub

Private Sub cmdGoBack_Click()
frmFormMales.Hide
frmForm3.Show
End Sub

Private Sub cmdLowtoHighBMI_Click()
Dim tempHeight As Integer
Dim tempWeight As Integer
Dim tempName As String
Dim tempBMI As Single
Dim NameA(1 To 100) As String
Dim Height(1 To 100) As Integer
Dim Weight(1 To 100) As Integer
Dim BMI(1 To 100) As Single
Dim CTR As Integer
Dim Pos As Integer
Dim Pass As Integer
picResults.Cls

Open App.Path & "\Males.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, NameA(CTR), Height(CTR), Weight(CTR), BMI(CTR)
Loop
    
For Pass = 1 To CTR - 1
    For Pos = 1 To (CTR - Pass)
        If BMI(Pos) > BMI(Pos + 1) Then
            
            tempHeight = Height(Pos)
            Height(Pos) = Height(Pos + 1)
            Height(Pos + 1) = tempHeight
            
            tempWeight = Weight(Pos)
            Weight(Pos) = Weight(Pos + 1)
            Weight(Pos + 1) = tempWeight
            
            tempBMI = BMI(Pos)
            BMI(Pos) = BMI(Pos + 1)
            BMI(Pos + 1) = tempBMI
            
            tempName = NameA(Pos)
            NameA(Pos) = NameA(Pos + 1)
            NameA(Pos + 1) = tempName
           
        End If
    Next Pos
Next Pass
For Pos = 1 To CTR
    picResults.Print NameA(Pos), Height(Pos); " inches", Weight(Pos); " lbs", BMI(Pos)
Next Pos
Close #1
End Sub

Private Sub cmdLowtoHighWeight_Click()
Dim tempHeight As Integer
Dim tempWeight As Integer
Dim tempName As String
Dim tempBMI As Single
Dim NameA(1 To 100) As String
Dim Height(1 To 100) As Integer
Dim Weight(1 To 100) As Integer
Dim BMI(1 To 100) As Single
Dim CTR As Integer
Dim Pos As Integer
Dim Pass As Integer
picResults.Cls

Open App.Path & "\Males.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, NameA(CTR), Height(CTR), Weight(CTR), BMI(CTR)
Loop
    
For Pass = 1 To CTR - 1
    For Pos = 1 To (CTR - Pass)
        If Weight(Pos) > Weight(Pos + 1) Then
            
            tempHeight = Height(Pos)
            Height(Pos) = Height(Pos + 1)
            Height(Pos + 1) = tempHeight
            
            tempWeight = Weight(Pos)
            Weight(Pos) = Weight(Pos + 1)
            Weight(Pos + 1) = tempWeight
            
            tempBMI = BMI(Pos)
            BMI(Pos) = BMI(Pos + 1)
            BMI(Pos + 1) = tempBMI
            
            tempName = NameA(Pos)
            NameA(Pos) = NameA(Pos + 1)
            NameA(Pos + 1) = tempName
           
        End If
    Next Pos
Next Pass
For Pos = 1 To CTR
    picResults.Print NameA(Pos), Height(Pos); " inches", Weight(Pos); " lbs", BMI(Pos)
Next Pos
Close #1
End Sub
Private Sub cmdNext_Click()
frmFormMales.Hide
frmForm4.Show
End Sub

Private Sub cmdTalltoShort_Click()
Dim tempHeight As Integer
Dim tempWeight As Integer
Dim tempName As String
Dim tempBMI As Single
Dim NameA(1 To 100) As String
Dim Height(1 To 100) As Integer
Dim Weight(1 To 100) As Integer
Dim BMI(1 To 100) As Single
Dim CTR As Integer
Dim Pos As Integer
Dim Pass As Integer
picResults.Cls

Open App.Path & "\Males.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, NameA(CTR), Height(CTR), Weight(CTR), BMI(CTR)
Loop
    
For Pass = 1 To CTR - 1
    For Pos = 1 To (CTR - Pass)
        If Height(Pos) > Height(Pos + 1) Then
            
            tempHeight = Height(Pos)
            Height(Pos) = Height(Pos + 1)
            Height(Pos + 1) = tempHeight
            
            tempWeight = Weight(Pos)
            Weight(Pos) = Weight(Pos + 1)
            Weight(Pos + 1) = tempWeight
            
            tempBMI = BMI(Pos)
            BMI(Pos) = BMI(Pos + 1)
            BMI(Pos + 1) = tempBMI
            
            tempName = NameA(Pos)
            NameA(Pos) = NameA(Pos + 1)
            NameA(Pos + 1) = tempName
           
        End If
    Next Pos
Next Pass
For Pos = 1 To CTR
    picResults.Print NameA(Pos), Height(Pos); " inches", Weight(Pos); " lbs", BMI(Pos)
Next Pos
Close #1
End Sub

