VERSION 5.00
Begin VB.Form frmSortingPage 
   BackColor       =   &H00000000&
   Caption         =   "Sorting Page"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10740
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdexplanation 
      BackColor       =   &H000000FF&
      Caption         =   "Go to Explanation Page"
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6720
      Width           =   3615
   End
   Begin VB.CommandButton cmdyearsorting 
      BackColor       =   &H000000FF&
      Caption         =   "Sort By Ascending Year"
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   3615
   End
   Begin VB.CommandButton cmdweightsorting 
      BackColor       =   &H000000FF&
      Caption         =   "Sort By Descending Weight Class"
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   3615
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H000000FF&
      Caption         =   "Go to Searching Page"
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   3615
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "End"
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7320
      Width           =   3615
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back to Home page"
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   3615
   End
   Begin VB.CommandButton cmdSort4 
      BackColor       =   &H000000FF&
      Caption         =   "Sort By Descending Year"
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   3615
   End
   Begin VB.CommandButton cmdSort3 
      BackColor       =   &H000000FF&
      Caption         =   "Sort By Ascending Weight Class"
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   3615
   End
   Begin VB.CommandButton cmdSort2 
      BackColor       =   &H000000FF&
      Caption         =   "Sort Reverse Alphabetically "
      Height          =   495
      Left            =   7440
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   3615
   End
   Begin VB.PictureBox Picture1 
      Height          =   9255
      Left            =   480
      Negotiate       =   -1  'True
      ScaleHeight     =   9195
      ScaleWidth      =   6795
      TabIndex        =   1
      Top             =   480
      Width           =   6855
   End
   Begin VB.CommandButton cmdSort1 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "Sort Alphabetically "
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   3615
   End
End
Attribute VB_Name = "frmSortingPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    frmWrestlingSorter.Visible = True
    frmSortingPage.Visible = False 'this button allows the user go from the sorting page to the home page.
End Sub

Private Sub cmdExit_Click()
    End 'this button ends the progam
End Sub

Private Sub cmdexplanation_Click()
    frmSortingPage.Visible = False
    frmexplanation.Visible = True 'this button allows the use to go from the sorting page to the explanation page
End Sub

Private Sub cmdSearch_Click()
    frmSortingPage.Visible = False 'this button allows the user to go from the sorting page to the searching page
    FrmNameSearch.Visible = True
End Sub

Private Sub cmdSort1_Click()
    Dim Pos As Integer, Pass As Integer
    Dim Temp1 As String ' this button allows the user to sort the roster alphabetically
    Dim Temp2 As Integer
    Dim Temp3 As String
    Dim temp4 As String
    Dim results As Integer
    
        For Pass = 1 To ctr - 1
            For Pos = 1 To ctr - Pass
                If LastName(Pos) > LastName(Pos + 1) Then
                    Temp1 = LastName(Pos)
                    LastName(Pos) = LastName(Pos + 1)
                    LastName(Pos + 1) = Temp1
                    
                    temp4 = FirstName(Pos)
                    FirstName(Pos) = FirstName(Pos + 1)
                    FirstName(Pos + 1) = temp4
                    
                    Temp2 = Weights(Pos)
                    Weights(Pos) = Weights(Pos + 1)
                    Weights(Pos + 1) = Temp2
                    
                    Temp3 = Year(Pos)
                    Year(Pos) = Year(Pos + 1)
                    Year(Pos + 1) = Temp3
                End If
             Next Pos
        Next Pass
    Picture1.Cls
    
    For results = 1 To ctr
         Picture1.Print LastName(results); Tab(30); FirstName(results); Tab(50); Weights(results); Tab(60); Year(results)
    Next results
End Sub

Private Sub cmdSort2_Click()
    Dim Pos As Integer, Pass As Integer
    Dim Temp1 As String
    Dim Temp2 As Integer 'this button allows the user to sort the roster reverse alphabetically
    Dim Temp3 As String
    Dim temp4 As String
    Dim results As Integer
    
        For Pass = 1 To ctr - 1
            For Pos = 1 To ctr - Pass
                If LastName(Pos) < LastName(Pos + 1) Then
                    Temp1 = LastName(Pos)
                    LastName(Pos) = LastName(Pos + 1)
                    LastName(Pos + 1) = Temp1
                    
                    temp4 = FirstName(Pos)
                    FirstName(Pos) = FirstName(Pos + 1)
                    FirstName(Pos + 1) = temp4
                    
                    
                    
                    Temp2 = Weights(Pos)
                    Weights(Pos) = Weights(Pos + 1)
                    Weights(Pos + 1) = Temp2
                    
                    Temp3 = Year(Pos)
                    Year(Pos) = Year(Pos + 1)
                    Year(Pos + 1) = Temp3
                End If
             Next Pos
        Next Pass
    
    Picture1.Cls
    
    For results = 1 To ctr
         Picture1.Print LastName(results); Tab(30); FirstName(results); Tab(50); Weights(results); Tab(60); Year(results)
    Next results
End Sub

Private Sub cmdSort3_Click()
    Dim Pos As Integer, Pass As Integer
    Dim Temp1 As Integer
    Dim Temp2 As String
    Dim Temp3 As String 'this button allows the user to sort the roster by accending weight class
    Dim temp4 As String
    Dim results As Integer
    
        For Pass = 1 To ctr - 1
            For Pos = 1 To ctr - Pass
                If Weights(Pos) > Weights(Pos + 1) Then
                    Temp1 = Weights(Pos)
                    Weights(Pos) = Weights(Pos + 1)
                    Weights(Pos + 1) = Temp1
                    
                    Temp2 = LastName(Pos)
                    LastName(Pos) = LastName(Pos + 1)
                    LastName(Pos + 1) = Temp2
                    
                    temp4 = FirstName(Pos)
                    FirstName(Pos) = FirstName(Pos + 1)
                    FirstName(Pos + 1) = temp4
                    
                    
                    Temp3 = Year(Pos)
                    Year(Pos) = Year(Pos + 1)
                    Year(Pos + 1) = Temp3
                End If
             Next Pos
        Next Pass
    
    Picture1.Cls
    
    For results = 1 To ctr
        Picture1.Print LastName(results); Tab(30); FirstName(results); Tab(50); Weights(results); Tab(60); Year(results)
    Next results
End Sub

Private Sub cmdSort4_Click()
  Dim Pos As Integer, Pass As Integer
    Dim Temp1 As Integer
    Dim Temp2 As String
    Dim Temp3 As String
    Dim temp4 As String
    Dim results As Integer 'this button allows the user to sort the roster but decending year
    
        For Pass = 1 To ctr - 1
            For Pos = 1 To ctr - Pass
                If Year(Pos) < Year(Pos + 1) Then
                    Temp3 = Year(Pos)
                    Year(Pos) = Year(Pos + 1)
                    Year(Pos + 1) = Temp3
                    
                    Temp1 = Weights(Pos)
                    Weights(Pos) = Weights(Pos + 1)
                    Weights(Pos + 1) = Temp1
                    
                    Temp2 = LastName(Pos)
                    LastName(Pos) = LastName(Pos + 1)
                    LastName(Pos + 1) = Temp2
                    
                    temp4 = FirstName(Pos)
                    FirstName(Pos) = FirstName(Pos + 1)
                    FirstName(Pos + 1) = temp4
                    
                End If
             Next Pos
        Next Pass
    
    Picture1.Cls
    
    For results = 1 To ctr
        Picture1.Print LastName(results); Tab(30); FirstName(results); Tab(50); Weights(results); Tab(60); Year(results)
    Next results
End Sub

Private Sub cmdweightsorting_Click()
  Dim Pos As Integer, Pass As Integer
    Dim Temp1 As String
    Dim Temp2 As String 'this button allows the user to sort the roster by decending weight class
    Dim Temp3 As Integer
    Dim temp4 As String
    Dim results As Integer
    
        For Pass = 1 To ctr - 1
            For Pos = 1 To ctr - Pass
                If Weights(Pos) < Weights(Pos + 1) Then
                    Temp3 = Weights(Pos)
                    Weights(Pos) = Weights(Pos + 1)
                    Weights(Pos + 1) = Temp3
                    
                    Temp1 = Year(Pos)
                    Year(Pos) = Year(Pos + 1)
                    Year(Pos + 1) = Temp1
                    
                    Temp2 = LastName(Pos)
                    LastName(Pos) = LastName(Pos + 1)
                    LastName(Pos + 1) = Temp2
                    
                    temp4 = FirstName(Pos)
                    FirstName(Pos) = FirstName(Pos + 1)
                    FirstName(Pos + 1) = temp4
                    
                End If
             Next Pos
        Next Pass
    
    Picture1.Cls
    
    For results = 1 To ctr
        Picture1.Print LastName(results); Tab(30); FirstName(results); Tab(50); Weights(results); Tab(60); Year(results)
    Next results
End Sub

Private Sub cmdyearsorting_Click()
  Dim Pos As Integer, Pass As Integer
    Dim Temp1 As String
    Dim Temp2 As String
    Dim Temp3 As Integer 'this button allows the user to sort the roster by Accending year
    Dim temp4 As String
    Dim results As Integer
    
        For Pass = 1 To ctr - 1
            For Pos = 1 To ctr - Pass
                If Year(Pos) > Year(Pos + 1) Then
                    Temp1 = Year(Pos)
                    Year(Pos) = Year(Pos + 1)
                    Year(Pos + 1) = Temp1
                    
                    Temp2 = LastName(Pos)
                    LastName(Pos) = LastName(Pos + 1)
                    LastName(Pos + 1) = Temp2
                    
                    temp4 = FirstName(Pos)
                    FirstName(Pos) = FirstName(Pos + 1)
                    FirstName(Pos + 1) = temp4
                    
                    Temp3 = Weights(Pos)
                    Weights(Pos) = Weights(Pos + 1)
                    Weights(Pos + 1) = Temp3
                End If
             Next Pos
        Next Pass
    
    Picture1.Cls
    
    For results = 1 To ctr
        Picture1.Print LastName(results); Tab(30); FirstName(results); Tab(50); Weights(results); Tab(60); Year(results)
    Next results
End Sub

Private Sub cmdyearsort2_Click()

End Sub
