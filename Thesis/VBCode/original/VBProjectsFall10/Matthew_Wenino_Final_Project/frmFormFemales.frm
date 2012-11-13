VERSION 5.00
Begin VB.Form frmFormFemales 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   11055
   ClientLeft      =   3615
   ClientTop       =   2370
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   11055
   ScaleWidth      =   7290
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
      Height          =   615
      Left            =   840
      TabIndex        =   8
      Top             =   8880
      Width           =   5415
   End
   Begin VB.CommandButton cmdLowtoHighWeight 
      Caption         =   "Low  to High Weight"
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
      Left            =   840
      TabIndex        =   7
      Top             =   9720
      Width           =   1815
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
      Left            =   2880
      TabIndex        =   6
      Top             =   7680
      Width           =   1575
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
      Left            =   4680
      TabIndex        =   5
      Top             =   7680
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
      Left            =   840
      TabIndex        =   3
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   9720
      Width           =   1575
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
      Left            =   4680
      TabIndex        =   1
      Top             =   9720
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   6015
      Left            =   840
      ScaleHeight     =   5955
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label lblIntro 
      BackColor       =   &H0080FFFF&
      Caption         =   "NOW LET'S FIND THE PERFECT FEMALE                     FOR YOU!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   4
      Top             =   480
      Width           =   6135
   End
End
Attribute VB_Name = "frmFormFemales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDisplay_Click()

Dim NameA(1 To 100) As String
Dim Height(1 To 100) As Integer
Dim Weight(1 To 100) As Integer
Dim BMI(1 To 100) As Single
Dim CTR As Integer
Dim Pos As Integer

picResults.Cls

Open App.Path & "\Females.txt" For Input As #1
Do While Not EOF(1)
    Pos = Pos + 1
    CTR = CTR + 1
    Input #1, NameA(Pos), Height(Pos), Weight(Pos), BMI(Pos)
    picResults.Print NameA(Pos), Height(Pos); " inches", Weight(Pos); " lbs", BMI(Pos)
Loop
Close #1
End Sub

Private Sub cmdEnter_Click()
Dim DB As Database
Dim QD As QueryDef

Dim RS As Recordset2

Set DB = OpenDatabase(App.Path & "\Weight_Range.accdb")
Set QD = DB.QueryDefs("SpecificFemale")
QD.Parameters(0) = InputBox("Please enter preferred minimum height in inches.")
QD.Parameters(1) = InputBox("Please enter preferred maximum height in inches.")
QD.Parameters(2) = InputBox("Please enter preferred minimum weight.")
QD.Parameters(3) = InputBox("Please enter preferred maximum weight.")

Set RS = QD.OpenRecordset()
picResults.Cls
picResults.Print "These are the women that met your criteria."
picResults.Print "***********************************************************"
Do Until RS.EOF
    picResults.Print RS![FirstName], RS![Height]; " inches", RS![Weight]; " lbs", RS![Body Index]
    RS.MoveNext
Loop
RS.Close
DB.Close
End Sub

Private Sub cmdGoBack_Click()
frmFormFemales.Hide
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

Open App.Path & "\Females.txt" For Input As #1
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

Open App.Path & "\Females.txt" For Input As #1
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
frmFormFemales.Hide
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

Open App.Path & "\Females.txt" For Input As #1
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
