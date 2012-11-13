VERSION 5.00
Begin VB.Form frmDetails 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13530
   LinkTopic       =   "Form2"
   Picture         =   "View Details.frx":0000
   ScaleHeight     =   10230
   ScaleWidth      =   13530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMaga 
      Caption         =   "Move to magazines page"
      Height          =   735
      Left            =   9720
      TabIndex        =   15
      Top             =   9360
      Width           =   3015
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to files"
      Height          =   735
      Left            =   10080
      TabIndex        =   14
      Top             =   7200
      Width           =   2655
   End
   Begin VB.CommandButton cmdAdvancedSearch 
      Caption         =   "Advanced Search"
      Height          =   735
      Left            =   9960
      TabIndex        =   13
      Top             =   6000
      Width           =   2775
   End
   Begin VB.TextBox txtTo 
      Height          =   495
      Left            =   11280
      TabIndex        =   12
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox txtFrom 
      Height          =   615
      Left            =   11280
      TabIndex        =   11
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtTitle 
      Height          =   495
      Left            =   11160
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Go back to the main Menu"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   9360
      Width           =   2775
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "List Books"
      Height          =   495
      Left            =   10080
      TabIndex        =   5
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdPrice 
      Caption         =   "Search by the price"
      Height          =   495
      Left            =   10080
      TabIndex        =   4
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton cmdAuthor 
      Caption         =   "Search by catergory"
      Height          =   615
      Left            =   10080
      TabIndex        =   3
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton cmdTitle 
      Caption         =   "Search by the Title"
      Height          =   615
      Left            =   10080
      TabIndex        =   2
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Search by ISBN"
      Height          =   495
      Left            =   10200
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   360
      ScaleHeight     =   8835
      ScaleWidth      =   9315
      TabIndex        =   0
      Top             =   240
      Width           =   9375
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   495
      Left            =   10080
      TabIndex        =   10
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   375
      Left            =   10080
      TabIndex        =   9
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   255
      Left            =   10080
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   10200
      Left            =   0
      Picture         =   "View Details.frx":BC0A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DB As Database
Dim RS As Recordset2
Dim InputISBN As String
Dim Found As Boolean

Private Sub cmdSearch_Click()

End Sub

Private Sub cmdAdvancedSearch_Click()
Dim InputPrice1 As Single
Dim InputPrice2 As Single
Dim InputName As String
Dim Temp As Single
'Open Access Database
Set DB = OpenDatabase(App.Path & "\Books.accdb")
Set RS = DB.OpenRecordset("Books")
RS.Index = "PrimaryKey"
'Input the textbooks information and its prices to search by advanced search
InputName = txtTitle.Text
InputPrice1 = txtFrom.Text
InputPrice2 = txtTo.Text
' Compare the lower bound and upper bound prices and then swap them if needed
    If InputPrice1 > InputPrice2 Then
        Temp = InputPrice1
        InputPrice1 = InputPrice2
        InputPrice2 = Temp
    End If
        
Found = False
picResults2.Cls
picResults2.Print " ISBN "; Tab(25); " Book's Name "; Tab(65); " Author "; Tab(90); " Price "
picResults2.Print "*************************************************************************************"
CurrentExp = 0
    'Exhaustive search
    Do Until RS.EOF
        If InStr(LCase(RS![Book Name]), LCase(InputName)) <> 0 Then
            
            If RS![Price] >= InputPrice1 And RS![Price] <= InputPrice2 Then
                Found = True
                CurrentExp = CurrentExp + 1
                ExportItem(CurrentExp) = RS![Book Name] & ", " & RS![Price]
                picResults2.Print RS![ISBN]; Tab(25); RS![Book Name]; Tab(65); RS![Author]; Tab(90); FormatCurrency(RS![Price])
           
            End If
            
        End If
        RS.MoveNext
    Loop
    If Found = False Then
        MsgBox " No match Found "
        
    End If

End Sub

Private Sub cmdAuthor_Click()
Dim InputType As String
InputType = InputBox(" Search for keywords", " Catergory ")
Set DB = OpenDatabase(App.Path & "\Books.accdb")
Set RS = DB.OpenRecordset("Books")
RS.Index = "PrimaryKey"
Found = False
picResults2.Cls
picResults2.Print " ISBN "; Tab(25); " Book's Name "; Tab(60); " Author "; Tab(90); " Price "
picResults2.Print "*************************************************************************************"
CurrentExp = 0
Do Until RS.EOF
        If InStr(LCase(RS![Type]), LCase(InputType)) <> 0 Then
            Found = True
            CurrentExp = CurrentExp + 1
            ExportItem(CurrentExp) = RS![Book Name] & ", " & RS![Price]
            picResults2.Print RS![ISBN]; Tab(25); RS![Book Name]; Tab(65); RS![Author]; Tab(90); FormatCurrency(RS![Price])
        
        End If
        RS.MoveNext
Loop
If Found = False Then
        MsgBox " No match Found "
        
End If
End Sub

Private Sub CmdBack_Click()
frmChoosing.Show
frmDetails.Hide
End Sub

Private Sub cmdExport_Click()
'Either create a exported file or continue to append to the current files
Open App.Path & "\ExportItems.txt" For Append As #1
    
    Dim I As Integer
    For I = 1 To CurrentExp
        Print #1, ExportItem(I)
    Next I
    CurrentExp = 0
    MsgBox "Successfully Export"
    Close #1
End Sub

Private Sub cmdList_Click()
Set DB = OpenDatabase(App.Path & "\Books.accdb")
Set RS = DB.OpenRecordset("Books")
RS.Index = "PrimaryKey"
picResults2.Cls
picResults2.Print " ISBN "; Tab(25); " Book's Name "; Tab(65); " Author "
picResults2.Print "*************************************************************************************"
CurrentExp = 0
    Do Until RS.EOF
        'add book to the export list
        CurrentExp = CurrentExp + 1
        ExportItem(CurrentExp) = RS![Book Name] & ", " & RS![Price]
        picResults2.Print RS![ISBN]; Tab(25); RS![Book Name]; Tab(65); RS![Author]
        RS.MoveNext
    Loop
    
End Sub


Private Sub cmdMaga_Click()
frmDetails.Hide
frmMagazine.Show
End Sub

Private Sub cmdPrice_Click()
'Search by price from the lower bound and upper bound prices
Dim InputPrice1 As Single
Dim InputPrice2 As Single
Found = False
InputPrice1 = InputBox(" Enter your lower bound price", " Price ")
InputPrice2 = InputBox(" Enter your upper bound price ", " Price ")

Set DB = OpenDatabase(App.Path & "\Books.accdb")
Set RS = DB.OpenRecordset("Books")
RS.Index = "PrimaryKey"
picResults2.Cls
picResults2.Print " ISBN "; Tab(25); " Book's Name "; Tab(60); " Author "; Tab(90); " Price "
picResults2.Print "*************************************************************************************"
CurrentExp = 0
Do Until RS.EOF
        If RS![Price] >= InputPrice1 And RS![Price] <= InputPrice2 Then
            Found = True
            CurrentExp = CurrentExp + 1
            ExportItem(CurrentExp) = RS![Book Name] & ", " & RS![Price]
            picResults2.Print RS![ISBN]; Tab(25); RS![Book Name]; Tab(60); RS![Author]; Tab(90); FormatCurrency(RS![Price])
        
        End If
        RS.MoveNext
Loop
If Found = False Then
        MsgBox " No match Found "
        
End If
End Sub


Private Sub cmdTitle_Click()
'Search by Title by exhaustive search
Dim InputName As String
Found = False

InputName = InputBox(" Search for keywords", "Title ")
Set DB = OpenDatabase(App.Path & "\Books.accdb")
Set RS = DB.OpenRecordset("Books")
RS.Index = "PrimaryKey"
picResults2.Cls
picResults2.Print " ISBN "; Tab(25); " Book's Name "; Tab(60); " Author "; Tab(90); " Price "
picResults2.Print "*************************************************************************************"
CurrentExp = 0
Do Until RS.EOF
        'Find keywords using InStr
        If InStr(LCase(RS![Book Name]), LCase(InputName)) <> 0 Then
            Found = True
            CurrentExp = CurrentExp + 1
            ExportItem(CurrentExp) = RS![Book Name] & ", " & RS![Price]
            
            picResults2.Print RS![ISBN]; Tab(25); RS![Book Name]; Tab(60); RS![Author]; Tab(90); FormatCurrency(RS![Price])
        End If
        RS.MoveNext
Loop
If Found = False Then
        MsgBox " No match Found "
        
End If
End Sub

Private Sub cmdView_Click()
'Match and stop search to search for ISBN of the book
InputISBN = InputBox(" Enter the ISBN ")
Set DB = OpenDatabase(App.Path & "\Books.accdb")
Set RS = DB.OpenRecordset("Books")
RS.Index = "PrimaryKey"
CurrentExp = 0
    
    Do Until RS.EOF
        If RS![ISBN] = InputISBN Then
            CurrentExp = CurrentExp + 1
            ExportItem(CurrentExp) = RS![Book Name] & ", " & RS![Price]
            Found = True
            picResults2.Cls
            picResults2.Print " ISBN "; Tab(25); " Book's Name "; Tab(60); " Author "; Tab(90); " Price "
            picResults2.Print "*************************************************************************************"
            picResults2.Print RS![ISBN]; Tab(25); RS![Book Name]; Tab(60); RS![Author]; Tab(90); FormatCurrency(RS![Price])
        
        End If
        RS.MoveNext
    Loop
    If Found = False Then
        MsgBox " No match Found "
        
    End If
End Sub

