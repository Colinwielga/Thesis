VERSION 5.00
Begin VB.Form FrmNameSearch 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexplanation 
      BackColor       =   &H000000FF&
      Caption         =   "Go to Explanation page"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H000000FF&
      Caption         =   "Search!"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdend 
      BackColor       =   &H000000FF&
      Caption         =   "End Program"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdsearchyear 
      BackColor       =   &H000000FF&
      Caption         =   "Search By Year"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortingPage 
      BackColor       =   &H000000FF&
      Caption         =   "Go To Sorting Page"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdback2 
      BackColor       =   &H000000FF&
      Caption         =   "Back to Home Page"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   4215
      Left            =   2640
      ScaleHeight     =   4155
      ScaleWidth      =   4755
      TabIndex        =   2
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox txtNsearch 
      BackColor       =   &H000000FF&
      Height          =   735
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Lblsearch1 
      BackColor       =   &H000000FF&
      Caption         =   "Enter the name of a wrestler you are looking for and then click Search"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "FrmNameSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdback2_Click()
    FrmNameSearch.Visible = False
    frmWrestlingSorter.Visible = True ' this button allows the user to go from the searching page to the homepage
End Sub

Private Sub cmdend_Click()
    End 'this button ends the program
End Sub

Private Sub cmdexplanation_Click()
    frmexplanation.Visible = True
    FrmNameSearch.Visible = False 'this button allows the user to go to explanation page to the searching page
End Sub

Private Sub cmdSearch_Click()
    Dim Pos As Integer '   this button allows the user to search names through out the file, if there are duplicates it will print all of them.
    Dim Sname As String
    Dim Found As Boolean
    Found = False

        Sname = txtNsearch.Text
        Picture2.Cls
        Found = False
        For Pos = 1 To ctr
            If LCase(Sname) = LCase(LastName(Pos)) Then
                Found = True
                Picture2.Print FirstName(Pos); " "; LastName(Pos), Tab(30); Weights(Pos), Tab(50); Year(Pos)
            End If
            If LCase(Sname) = LCase(FirstName(Pos)) Then
                Found = True
                Picture2.Print FirstName(Pos); " "; LastName(Pos), Tab(30); Weights(Pos), Tab(50); Year(Pos)
            End If
    
    Next Pos
    
        If Not Found Then
             MsgBox "That name does not appear on the on the roster", , "Error"
        End If
    
   
    
    
    
    
 
End Sub

Private Sub cmdsearchyear_Click() 'this allows the user to search for all matches of a year within the array.
    Dim Syear As String
    Dim Found As Boolean
    Dim Pos As Integer

    Syear = InputBox("Enter a Year you would like to find  Fr, So, Jr, Sr")
    Picture2.Cls
    Found = False
        For Pos = 1 To ctr
            If Syear = Year(Pos) Then
                Found = True
                Picture2.Print FirstName(Pos); " "; LastName(Pos), Tab(30); Weights(Pos), Tab(50); Year(Pos)
            End If
      
        Next Pos
End Sub

Private Sub cmdSortingPage_Click()
    FrmNameSearch.Visible = False
    frmSortingPage.Visible = True
'this allows the user to go to the sorting page from the searching page
End Sub

