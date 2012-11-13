VERSION 5.00
Begin VB.Form frmCost 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Cost of Living"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   Begin VB.CommandButton cmdFutureCost 
      BackColor       =   &H00C000C0&
      Caption         =   "Find Future Cost of Living"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1460
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   11280
      Width           =   2550
   End
   Begin VB.CommandButton cmdAverageCost 
      BackColor       =   &H00C000C0&
      Caption         =   "Find Average Cost of Living"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1460
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   10560
      Width           =   2550
   End
   Begin VB.CommandButton cmdReturntoProf 
      BackColor       =   &H00C000C0&
      Caption         =   "Return to Professions Selection"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1460
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   11280
      Width           =   2550
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C000C0&
      Caption         =   "End Simulation"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1460
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   11280
      Width           =   2550
   End
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00C000C0&
      Caption         =   "Select a State"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1460
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8880
      Width           =   2550
   End
   Begin VB.CommandButton cmdleast 
      BackColor       =   &H00C000C0&
      Caption         =   "Cheapest State to Live in"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1460
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7200
      Width           =   2550
   End
   Begin VB.CommandButton CmdMost 
      BackColor       =   &H00C000C0&
      Caption         =   "Most Expensive State to live in"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1460
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5520
      Width           =   2550
   End
   Begin VB.CommandButton cmdAscending 
      BackColor       =   &H00C000C0&
      Caption         =   "List States by Cost of Living (Least Expensive to Most Expensive)"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1460
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   2550
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   3360
      ScaleHeight     =   8955
      ScaleWidth      =   11235
      TabIndex        =   1
      Top             =   1920
      Width           =   11295
      Begin VB.VScrollBar MyScroll 
         Height          =   8055
         Left            =   10560
         TabIndex        =   16
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdDescending 
      BackColor       =   &H00C000C0&
      Caption         =   "List States by Cost of Living (Most Expensive to Least Expensive)"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1460
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   2550
   End
   Begin VB.Label lblFS5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "5. Two Adults and Two Children"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label lblFS4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "4. Two Adults and One Child"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label lblFS3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "3. Two Adults"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   10200
      TabIndex        =   6
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblFS2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "2. One Adult and One Child"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label lblFS1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "1. One Adult"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblFamilySize 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Cost of living data has been categorized by the following: "
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   10215
   End
End
Attribute VB_Name = "frmCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is solely associated to the data regarding cost of living. This form allow the user to search, sort, and manipulate
' the cost of livig data regarding all the States of the United states and one Federal District. The data is categorized by 5 family sizes and the
' task of this form is to execute the manipulation of the data across all those categories.

'Ensures that all variables are declared and serves as a spell checker for variables
Option Explicit

'Declare Form Level variable
Dim FamilySize As Integer
    
'Sort by cost of living in the U.S. in ascending order but within the categories of the 5 family sizes
Private Sub cmdAscending_Click()
    Dim Pos As Integer
    Dim Pass As Integer
    Dim tempUsState As String
    Dim tempOneAdult As Single
    Dim tempOneAdultOneChild As Single
    Dim tempTwoAdults As Single
    Dim tempTwoAdultsOneChild As Single
    Dim tempTwoAdultsTwoChildren As Single
    Dim FamilySize As Integer
    Dim ctr2 As Integer
    
    picResults.Cls

    FamilySize = InputBox("Please specify family size", "Family Size Required")
    
    If FamilySize = 1 Then
        For Pass = 1 To ctrRead - 1
            For Pos = 1 To ctrRead - Pass
            
            If OneAdult(Pos) > OneAdult(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of living for One Adult (Least expensive to most expensive)"
    picResults.Print "**********************************************************************************************************************"
    For Pos = 1 To ctrRead
        picResults.Print UsState(Pos); Tab(1); , , FormatCurrency(OneAdult(Pos))
    Next Pos
    ElseIf FamilySize = 2 Then
    For Pass = 1 To ctrRead - 1
            For Pos = 1 To ctrRead - Pass
            
            If OneAdultOneChild(Pos) > OneAdultOneChild(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for One Adult & One Child (Least expensive to most expensive)"
    picResults.Print "**********************************************************************************************************************"
    For Pos = 1 To ctrRead
        picResults.Print UsState(Pos); Tab(1); , , FormatCurrency(OneAdultOneChild(Pos))
    Next Pos
    ElseIf FamilySize = 3 Then
        For Pass = 1 To ctrRead - 1
            For Pos = 1 To ctrRead - Pass
            
            If TwoAdults(Pos) > TwoAdults(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for Two Adults (Least expensive to most expensive)"
    picResults.Print "**********************************************************************************************************************"
    For Pos = 1 To ctrRead
        picResults.Print UsState(Pos); Tab(1); , , FormatCurrency(TwoAdults(Pos))
    Next Pos
    
    ElseIf FamilySize = 4 Then
        For Pass = 1 To ctrRead - 1
            For Pos = 1 To ctrRead - Pass
                
            If TwoAdultsOneChild(Pos) > TwoAdultsOneChild(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for Two Adults & One Child (Least expensive to most expensive)"
    picResults.Print "**********************************************************************************************************************"
    For Pos = 1 To ctrRead
        picResults.Print UsState(Pos); Tab(1); , , FormatCurrency(TwoAdultsOneChild(Pos))
    Next Pos
    
    ElseIf FamilySize = 5 Then
    For Pass = 1 To ctrRead - 1
        For Pos = 1 To ctrRead - Pass
            
            If TwoAdultsTwoChildren(Pos) > TwoAdultsTwoChildren(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for Two Adults & Two Children (Least expensive to most expensive)"
    picResults.Print "**********************************************************************************************************************"
    For Pos = 1 To ctrRead
        picResults.Print UsState(Pos); Tab(1); , , FormatCurrency(TwoAdultsTwoChildren(Pos))
    Next Pos
    
    Else
    
        MsgBox "You must select a valid Family size!", , "Invalid Entry"
       
End If
End Sub

'Calculate average cost of living across all categories of family sizes
Private Sub cmdAverageCost_Click()

    FamilySize = InputBox("Please specify family size", "Family Size Required")
    
    If FamilySize = 1 Then
    
        Dim sum As Single
        Dim avg As Single
        Dim Pos As Integer
        For Pos = 1 To ctrRead
            sum = sum + OneAdult(Pos)
        Next Pos
        avg = sum / ctrRead
        MsgBox "The average cost of living in the United States for families consisting of only one adult is " & FormatCurrency(avg) & ""
    
    ElseIf FamilySize = 2 Then
    
        For Pos = 1 To ctrRead
            sum = sum + OneAdultOneChild(Pos)
        Next Pos
        avg = sum / ctrRead
        MsgBox "The average cost of living in the United States for families consisting of one adult and one child is " & FormatCurrency(avg) & ""
    
    ElseIf FamilySize = 3 Then
    
        For Pos = 1 To ctrRead
            sum = sum + TwoAdults(Pos)
        Next Pos
        avg = sum / ctrRead
        MsgBox "The average cost of living in the United States for families consisting of two adults is " & FormatCurrency(avg) & ""
    
    ElseIf FamilySize = 4 Then
    
        
        For Pos = 1 To ctrRead
            sum = sum + TwoAdultsOneChild(Pos)
        Next Pos
        avg = sum / ctrRead
        MsgBox "The average cost of living in the United States for families consisting of two adults and one child " & FormatCurrency(avg) & ""
    
    ElseIf FamilySize = 5 Then
    
        
        For Pos = 1 To ctrRead
            sum = sum + TwoAdultsTwoChildren(Pos)
        Next Pos
        avg = sum / ctrRead
        MsgBox "The average cost of living in the United States for families consisting of two adults and two children " & FormatCurrency(avg) & ""
        
    Else
        MsgBox "You must select a valid Family size!", , "Invalid Entry"
    
    End If
    
End Sub


'Sort by cost of living in the U.S. in descending order but within the categories of the 5 family sizes
Private Sub cmdDescending_Click()
    Dim Pos As Integer
    Dim Pass As Integer
    Dim tempUsState As String
    Dim tempOneAdult As Single
    Dim tempOneAdultOneChild As Single
    Dim tempTwoAdults As Single
    Dim tempTwoAdultsOneChild As Single
    Dim tempTwoAdultsTwoChildren As Single
    Dim FamilySize As Integer
    
    picResults.Cls

    FamilySize = InputBox("Please specify family size", "Family Size Required")
    
    If FamilySize = 1 Then
        For Pass = 1 To ctrRead - 1
            For Pos = 1 To ctrRead - Pass
            
            If OneAdult(Pos) < OneAdult(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of living for One Adult (Most expensive to least expensive)"
    picResults.Print "**********************************************************************************************************************"
    For Pos = 1 To ctrRead
        picResults.Print UsState(Pos); Tab(1); , , FormatCurrency(OneAdult(Pos))
    Next Pos
    ElseIf FamilySize = 2 Then
        For Pass = 1 To ctrRead - 1
            For Pos = 1 To ctrRead - Pass
            
            If OneAdultOneChild(Pos) < OneAdultOneChild(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for One Adult & One Child (Most expensive to least expensive)"
    picResults.Print "**********************************************************************************************************************"
    For Pos = 1 To ctrRead
        picResults.Print UsState(Pos); Tab(1); , , FormatCurrency(OneAdultOneChild(Pos))
    Next Pos
    ElseIf FamilySize = 3 Then
        For Pass = 1 To ctrRead - 1
            For Pos = 1 To ctrRead - Pass
            
            If TwoAdults(Pos) < TwoAdults(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for Two Adults (Most expensive to least expensive)"
    picResults.Print "**********************************************************************************************************************"
    For Pos = 1 To ctrRead
        picResults.Print UsState(Pos); Tab(1); , , FormatCurrency(TwoAdults(Pos))
    Next Pos
    
    ElseIf FamilySize = 4 Then
        For Pass = 1 To ctrRead - 1
            For Pos = 1 To ctrRead - Pass
            
            If TwoAdultsOneChild(Pos) < TwoAdultsOneChild(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for Two Adults & One Child (Most expensive to least expensive)"
    picResults.Print "**********************************************************************************************************************"
    For Pos = 1 To ctrRead
        picResults.Print UsState(Pos); Tab(1); , , FormatCurrency(TwoAdultsOneChild(Pos))
    Next Pos
    
    ElseIf FamilySize = 5 Then
    For Pass = 1 To ctrRead - 1
        For Pos = 1 To ctrRead - Pass
            
            If TwoAdultsTwoChildren(Pos) < TwoAdultsTwoChildren(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for Two Adults & Two Children (Most expensive to least expensive)"
    picResults.Print "**********************************************************************************************************************"
    For Pos = 1 To ctrRead
        picResults.Print UsState(Pos); Tab(1); , , FormatCurrency(TwoAdultsTwoChildren(Pos))
    Next Pos
    
    Else
    
        MsgBox "You must select a valid Family size!", , "Invalid Entry"
       
End If
End Sub

'Calculate future cost of living in United states based on the average across all 5 family sizes
Private Sub cmdFutureCost_Click()

    Dim FutureCost As Single
    Dim FutureYears As Integer
    Dim FutureValue As Single

    FamilySize = InputBox("Please specify family size", "Family Size Required")
    FutureYears = InputBox("Please enter number of years", "Number of years into the future")
        
    
    If FamilySize = 1 Then
    
        Dim sum As Single
        Dim avg As Single
        
        Dim Pos As Integer
        For Pos = 1 To ctrRead
            sum = sum + OneAdult(Pos)
        Next Pos
        avg = sum / ctrRead
        
        FutureValue = (1 + 0.018) ^ FutureYears * avg
        
        If FutureYears <= 0 Then
        MsgBox "Number of years must be greater than zero!", , "Invalid Entry"
        
        Else
        
        MsgBox "The cost of living in the United States for families consisting of only one adult " & FutureYears & " years from now is  " & FormatCurrency(FutureValue) & ""
        End If
        
    ElseIf FamilySize = 2 Then
    
        For Pos = 1 To ctrRead
            sum = sum + OneAdultOneChild(Pos)
        Next Pos
        avg = sum / ctrRead
        
        FutureValue = (1 + 0.018) ^ FutureYears * avg
        
        If FutureYears <= 0 Then
        MsgBox "Number of years must be greater than zero!", , "Invalid Entry"
        
        Else
        
        MsgBox "The cost of living in the United States for families consisting of one adult and one child " & FutureYears & " years from now is  " & FormatCurrency(FutureValue) & ""
        End If
    
    ElseIf FamilySize = 3 Then
    
        For Pos = 1 To ctrRead
            sum = sum + TwoAdults(Pos)
        Next Pos
        avg = sum / ctrRead
        
        FutureValue = (1 + 0.018) ^ FutureYears * avg
        
        If FutureYears <= 0 Then
        MsgBox "Number of years must be greater than zero!", , "Invalid Entry"
        
        Else
        
        MsgBox "The cost of living in the United States for families consisting of two adults " & FutureYears & " years from now is  " & FormatCurrency(FutureValue) & ""
        End If
    
    ElseIf FamilySize = 4 Then
    
        
        For Pos = 1 To ctrRead
            sum = sum + TwoAdultsOneChild(Pos)
        Next Pos
        avg = sum / ctrRead
        
        FutureValue = (1 + 0.018) ^ FutureYears * avg
        
        If FutureYears <= 0 Then
        MsgBox "Number of years must be greater than zero!", , "Invalid Entry"
        
        Else
        
        MsgBox "The cost of living in the United States for families consisting of two adults and one child " & FutureYears & " years from now is  " & FormatCurrency(FutureValue) & ""
        End If
    ElseIf FamilySize = 5 Then
    
        
        For Pos = 1 To ctrRead
            sum = sum + TwoAdultsTwoChildren(Pos)
        Next Pos
        avg = sum / ctrRead
        
        FutureValue = (1 + 0.018) ^ FutureYears * avg
        
        If FutureYears <= 0 Then
        MsgBox "Number of years must be greater than zero!", , "Invalid Entry"
        
        Else
        
        MsgBox "The cost of living in the United States for families consisting of two adults and two children " & FutureYears & " years from now is  " & FormatCurrency(FutureValue) & ""
        End If
    Else
        MsgBox "You must select a valid Family size!", , "Invalid Entry"
    
    End If
        
End Sub

'Find the least expensive state across all 5 family size categories
Private Sub cmdleast_Click()
Dim Pos As Integer
    Dim Pass As Integer
    Dim tempUsState As String
    Dim tempOneAdult As Single
    Dim tempOneAdultOneChild As Single
    Dim tempTwoAdults As Single
    Dim tempTwoAdultsOneChild As Single
    Dim tempTwoAdultsTwoChildren As Single
    Dim FamilySize As Integer
    
    picResults.Cls

    FamilySize = InputBox("Please specify family size", "Family Size Required")
    
If FamilySize = 1 Then
    For Pass = 1 To ctrRead - 1
        For Pos = 1 To ctrRead - Pass
            
            If OneAdult(Pos) > OneAdult(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of living for One Adult"
    picResults.Print "**********************************************************************************************************************"
    picResults.Print UsState(1); Tab(1); , , FormatCurrency(OneAdult(1))
    
ElseIf FamilySize = 2 Then
For Pass = 1 To ctrRead - 1
        For Pos = 1 To ctrRead - Pass
            
            If OneAdultOneChild(Pos) > OneAdultOneChild(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for One Adult & One Child"
    picResults.Print "**********************************************************************************************************************"
    picResults.Print UsState(1); Tab(1); , , FormatCurrency(OneAdultOneChild(1))
    
ElseIf FamilySize = 3 Then
For Pass = 1 To ctrRead - 1
        For Pos = 1 To ctrRead - Pass
            
            If TwoAdults(Pos) > TwoAdults(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for Two Adults"
    picResults.Print "**********************************************************************************************************************"
    picResults.Print UsState(1); Tab(1); , , FormatCurrency(TwoAdults(1))
    
    
ElseIf FamilySize = 4 Then
    For Pass = 1 To ctrRead - 1
        For Pos = 1 To ctrRead - Pass
            
            If TwoAdultsOneChild(Pos) > TwoAdultsOneChild(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for Two Adults & One Child"
    picResults.Print "**********************************************************************************************************************"
    picResults.Print UsState(1); Tab(1); , , FormatCurrency(TwoAdultsOneChild(1))
    
    
    ElseIf FamilySize = 5 Then
    For Pass = 1 To ctrRead - 1
        For Pos = 1 To ctrRead - Pass
            
            If TwoAdultsTwoChildren(Pos) > TwoAdultsTwoChildren(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for Two Adults & Two Children"
    picResults.Print "**********************************************************************************************************************"
    picResults.Print UsState(1); Tab(1); , , FormatCurrency(TwoAdultsTwoChildren(1))
    
    
    Else
    
        MsgBox "You must select a valid Family size!", , "Invalid Entry"
       
End If
End Sub

'Find the mosst expensive state across all 5 family size categories
Private Sub CmdMost_Click()
Dim Pos As Integer
    Dim Pass As Integer
    Dim tempUsState As String
    Dim tempOneAdult As Single
    Dim tempOneAdultOneChild As Single
    Dim tempTwoAdults As Single
    Dim tempTwoAdultsOneChild As Single
    Dim tempTwoAdultsTwoChildren As Single
    Dim FamilySize As Integer
    
    picResults.Cls

    FamilySize = InputBox("Please specify family size", "Family Size Required")
    
If FamilySize = 1 Then
    For Pass = 1 To ctrRead - 1
        For Pos = 1 To ctrRead - Pass
            
            If OneAdult(Pos) < OneAdult(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of living for One Adult"
    picResults.Print "**********************************************************************************************************************"
    picResults.Print UsState(1); Tab(1); , , FormatCurrency(OneAdult(1))
    
ElseIf FamilySize = 2 Then
For Pass = 1 To ctrRead - 1
        For Pos = 1 To ctrRead - Pass
            
            If OneAdultOneChild(Pos) < OneAdultOneChild(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for One Adult & One Child"
    picResults.Print "**********************************************************************************************************************"
    picResults.Print UsState(1); Tab(1); , , FormatCurrency(OneAdultOneChild(1))
    
ElseIf FamilySize = 3 Then
For Pass = 1 To ctrRead - 1
        For Pos = 1 To ctrRead - Pass
            
            If TwoAdults(Pos) < TwoAdults(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for Two Adults"
    picResults.Print "**********************************************************************************************************************"
    picResults.Print UsState(1); Tab(1); , , FormatCurrency(TwoAdults(1))
    
    
ElseIf FamilySize = 4 Then
    For Pass = 1 To ctrRead - 1
        For Pos = 1 To ctrRead - Pass
            
            If TwoAdultsOneChild(Pos) < TwoAdultsOneChild(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for Two Adults & One Child"
    picResults.Print "**********************************************************************************************************************"
    picResults.Print UsState(1); Tab(1); , , FormatCurrency(TwoAdultsOneChild(1))
    
    
    ElseIf FamilySize = 5 Then
    For Pass = 1 To ctrRead - 1
        For Pos = 1 To ctrRead - Pass
            
            If TwoAdultsTwoChildren(Pos) < TwoAdultsTwoChildren(Pos + 1) Then
                tempUsState = UsState(Pos)
                UsState(Pos) = UsState(Pos + 1)
                UsState(Pos + 1) = tempUsState
                
                tempOneAdult = OneAdult(Pos)
                OneAdult(Pos) = OneAdult(Pos + 1)
                OneAdult(Pos + 1) = tempOneAdult
                
                tempOneAdultOneChild = OneAdultOneChild(Pos)
                OneAdultOneChild(Pos) = OneAdultOneChild(Pos + 1)
                OneAdultOneChild(Pos + 1) = tempOneAdultOneChild
                
                tempTwoAdults = TwoAdults(Pos)
                TwoAdults(Pos) = TwoAdults(Pos + 1)
                TwoAdults(Pos + 1) = tempTwoAdults
                
                tempTwoAdultsOneChild = TwoAdultsOneChild(Pos)
                TwoAdultsOneChild(Pos) = TwoAdultsOneChild(Pos + 1)
                TwoAdultsOneChild(Pos + 1) = tempTwoAdultsOneChild
                
                tempTwoAdultsTwoChildren = TwoAdultsTwoChildren(Pos)
                TwoAdultsTwoChildren(Pos) = TwoAdultsTwoChildren(Pos + 1)
                TwoAdultsTwoChildren(Pos + 1) = tempTwoAdultsTwoChildren
                
            End If
        Next Pos
    Next Pass


    
    picResults.Cls
    picResults.Print "Cost of Living for Two Adults & Two Children"
    picResults.Print "**********************************************************************************************************************"
    picResults.Print UsState(1); Tab(1); , , FormatCurrency(TwoAdultsTwoChildren(1))
    
    
    Else
    
        MsgBox "You must select a valid Family size!", , "Invalid Entry"
       
End If
End Sub
'End program
Private Sub cmdQuit_Click()
    End
End Sub

'Return to profession selection screen
Private Sub cmdReturntoProf_Click()
    frmStart.Visible = True
    frmCost.Visible = False
End Sub

'Allow the user to search for a particular United state and explore cost across all five categories
Private Sub cmdSelect_Click()

    Dim Pos As Integer
    Dim Found As Boolean
    Dim StateName As String
    Dim FamilySize As Integer
    
    picResults.Cls
    
    Found = False
    StateName = InputBox("Please enter The U.S. State of Choice")
    FamilySize = InputBox("Please specify family size", "Family Size Required")
    
    If FamilySize = 1 Then
        Do Until Found = True Or Pos >= ctrRead
            Pos = Pos + 1
            If LCase(UsState(Pos)) = LCase(StateName) Then
                Found = True
            End If
        Loop
        If Found = True Then
            picResults.Print "The cost of living for one adult in " & LCase(StateName) & " is " & FormatCurrency(OneAdult(Pos)) & ""
        
        Else
            MsgBox "Not Found In Database"
        End If
        
    ElseIf FamilySize = 2 Then
        Do Until Found = True Or Pos >= ctrRead
            Pos = Pos + 1
            If LCase(UsState(Pos)) = LCase(StateName) Then
                Found = True
            End If
        Loop
        If Found = True Then
            picResults.Print "The cost of living for one adult and one child in " & LCase(StateName) & " is " & FormatCurrency(OneAdultOneChild(Pos)) & ""
        
        Else
            MsgBox "Not Found In Database"
        End If
        
    ElseIf FamilySize = 3 Then
        Do Until Found = True Or Pos >= ctrRead
            Pos = Pos + 1
            If LCase(UsState(Pos)) = LCase(StateName) Then
                Found = True
            End If
        Loop
        If Found = True Then
            picResults.Print "The cost of living for two adults in " & LCase(StateName) & " is " & FormatCurrency(TwoAdults(Pos)) & ""
        
        Else
            MsgBox "Not Found In Database"
        End If
        
    ElseIf FamilySize = 4 Then
        Do Until Found = True Or Pos >= ctrRead
            Pos = Pos + 1
            If LCase(UsState(Pos)) = LCase(StateName) Then
                Found = True
            End If
        Loop
        If Found = True Then
            picResults.Print "The cost of living for two adults and one child in " & LCase(StateName) & " is " & FormatCurrency(TwoAdultsOneChild(Pos)) & ""
        
        Else
            MsgBox "Not Found In Database"
        End If
        
    ElseIf FamilySize = 5 Then
        Do Until Found = True Or Pos >= ctrRead
            Pos = Pos + 1
            If LCase(UsState(Pos)) = LCase(StateName) Then
                Found = True
            End If
        Loop
        If Found = True Then
            picResults.Print "The cost of living for two adults and two children in " & LCase(StateName) & " is " & FormatCurrency(TwoAdultsTwoChildren(Pos)) & ""
        
        Else
            MsgBox "Not Found In Database"
        End If
    
    Else
            MsgBox "You must select a valid Family size!", , "Invalid Entry"
    
    End If
End Sub


