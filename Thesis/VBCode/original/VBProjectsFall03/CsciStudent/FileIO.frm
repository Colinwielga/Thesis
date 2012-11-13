VERSION 5.00
Begin VB.Form FileIO 
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   3615
   ClientTop       =   4500
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   4680
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdBinary 
      Caption         =   " Binary Search"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton CmdSort 
      Caption         =   "Sort"
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdSequential 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sequential SEARCH "
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0080FF80&
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   1635
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdCreateArray 
      Caption         =   "Put data from file in array"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdCreatefile 
      BackColor       =   &H00C0C0C0&
      Caption         =   "create data file"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   2175
   End
End
Attribute VB_Name = "FileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CTR, number As Integer
Dim AGES(1 To 500) As Integer

Private Sub Cmdcreatefile_Click()
'Create an output file
'Open "M:\vbsamp\vbdata.txt" For Output As #1
Open PATH & "vbdata.txt" For Output As #1
'generate sample data and write to file
For j = 1 To 50
    Write #1, (j * 17 Mod 501)
Next j
Close
cmdCreateArray.Enabled = True

End Sub

Private Sub CmdCreateArray_Click()
Picture1.Cls
'Open "M:\vbsamp\vbDATA.TXT" For Input As #1
Open PATH & "vbDATA.TXt" For Input As #1
CTR = 0
'Put data from file into an array
Do While (CTR < 50)
    CTR = CTR + 1
    Input #1, AGES(CTR)
    Picture1.Print AGES(CTR);
    If CTR Mod 10 = 0 Then Picture1.Print
Loop
Close
End Sub

Private Sub cmdSequential_Click()
'sequential search
CTR = 0
Found = False
number = InputBox("Enter your #", "Searching")

'Search until the number is found
'or until you have looked through the entire list.
Do While ((Not Found) And (CTR < 50))
    CTR = CTR + 1
    If AGES(CTR) = number Then
        Found = True
    End If
Loop

'The CTR tells you the position in the list
Picture2.Print "FOUND IS "; Found
Picture2.Print Tab(10); "Position "; CTR


End Sub

Private Sub CmdSort_Click()
Dim temp As Integer
Dim comp, pass As Integer

'Bubble sort the ages
For pass = 1 To 49
    For comp = 1 To 50 - pass
        If AGES(comp) > AGES(comp + 1) Then
            temp = AGES(comp)
            AGES(comp) = AGES(comp + 1)
            AGES(comp + 1) = temp
        End If
    Next comp
Next pass
Picture1.Cls

'print sorted list
For j = 1 To 50
    Picture1.Print AGES(j);
    If j Mod 10 = 0 Then Picture1.Print
Next j
cmdBinary.Enabled = True

End Sub

Private Sub CmdBinary_Click()
'Binary Search used to look for a # in an array.
Dim first, last, middle, looked As Integer
Picture2.Cls
looked = 0
first = 1
last = 50
middle = Int((first + last) / 2)
Found = False
number = InputBox("Enter your #", "Searching")
'The list is partitioned into smaller
'search spaces each time the looking fails
'until the list is exhausted or the #
'is found.  (If the number you are
'searching for is not in the first half
'of the list, then the search continues
'using only the second half of the list.)

Do While (first <= last) And (Not Found)
    looked = looked + 1
    If number = AGES(middle) Then
        Found = True
    Else
        If number > AGES(middle) Then
            first = middle + 1
        Else
            last = middle - 1
        End If
    End If
    middle = Int((first + last) / 2)
Loop

Picture2.Print "Found is "; Found
Picture2.Print Tab(10); looked; " tries"
        
End Sub

Private Sub CmdExit_Click()
End

End Sub

Private Sub Form_Load()

    Dim PATH As String
    PATH = "N:\CS130\handin\CsciStudent\"
    FileIO.Show
    


End Sub
