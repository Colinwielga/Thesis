VERSION 5.00
Begin VB.Form frmLoaded 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Works Cited"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdView 
      Caption         =   "Click here to view your formated Works Cited Page"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   8175
   End
   Begin VB.PictureBox picWorksCited 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   120
      ScaleHeight     =   7455
      ScaleWidth      =   8175
      TabIndex        =   4
      Top             =   480
      Width           =   8175
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print your Works Cited Page"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   8040
      Width           =   2415
   End
   Begin VB.CommandButton cmdAddMore 
      Caption         =   "Not done yet? Click here to add another source"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   8040
      Width           =   2055
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save your work"
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   8040
      Width           =   1455
   End
End
Attribute VB_Name = "frmLoaded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pos As Integer
'this form displays the bibliography that the user loaded

Private Sub cmdAddMore_Click()
    'sends the user back to the previous forms to add more sources to their bibliographies
    frmLoaded.Hide
    
    'if the type is MLA, then it sends it to the mla adding form
    If MLA = True Then
        frmMLABook.Show
    'otherwise, it goes to the APA form
    Else
        frmAPABook.Show
    End If
End Sub

Private Sub cmdPrint_Click()
    'this command allows the user to print their works cited page on a printer. I learned how to do this with help from: http://www.devarticles.com/c/a/Visual-Basic/Printing-With-Visual-Basic/1/
    Printer.Print ""
    Printer.Print UsersName
    Printer.Print Tab(75); "Works Cited"
    
    'determines the type of format needed
    If MLA = True Then
        'runs through the loop, printing each citation
        For pos = 1 To ctr
            Printer.Print ""
            Printer.Print AuthorsLastName(pos) & ", " & AuthorsFirstName(pos) & ". " & Title(pos) & ". " & CityPublished(pos) & ": " & Publisher(pos) & ", " & Year(pos)
           
        Next pos
    Else
        'runs through the loop, printing each citation
        For pos = 1 To ctr
            Printer.Print ""
            Printer.Print AuthorsLastName(pos) & ", " & AuthorsFirstName(pos) & ". " & AuthorsMiddleName(pos) & ". (" & Year(pos) & "). " & Title(pos) & ". " & CityPublished(pos) & ": " & Publisher(pos) & "."
        Next pos
    End If
    'tells the printer that this is the end of the information to be printed
    Printer.EndDoc
End Sub

Private Sub CmdQuit_Click()
    'ends the program
    End
End Sub

Private Sub cmdSave_Click()
    'allows the user to save their work. I learned to do this with help from http://www.developerfusion.co.uk/show/37/4/
        Dim pos As Integer
        Dim fileName As String
        pos = 0
        
        'asks the user to input a filename to save their work as
        fileName = InputBox("Please type the file name that you wish to save your works cited as:")
        
        'opens the file. The previous contents will be overwritten. Or if the file doesn't already exist, a new file will be created.
        Open App.Path & "\saved\" & fileName & ".txt" For Output As #1
        
        'runs through the loop, saving each citation in the file
        Do Until pos = ctr
            pos = pos + 1
            Write #1, AuthorsLastName(pos), AuthorsFirstName(pos), AuthorsMiddleName(pos), Title(pos), CityPublished(pos), Publisher(pos), Year(pos)
        Loop
        
        'closes the file
        Close #1
        
        'tells the user what their work has been saved as
        MsgBox "Your file has been saved in the \saved folder as " & fileName & ".txt"
End Sub

    Private Sub cmdView_Click()
    'Sorts all of the arrays alphabetically and then outputs them to be viewed by the user
    
    Dim pass As Integer, comp As Integer
    Dim tempAuthorsLastName As String
    Dim tempAuthorsFirstName As String
    Dim tempAuthorsMiddleName As String
    Dim tempTitle As String
    Dim tempCityPublished As String
    Dim tempPublisher As String
    Dim tempYear As String
    Dim tempMiddleName
    
    'sorts bibliography alphabetically
    For pass = 1 To (ctr - 1)
        For comp = 1 To (ctr - pass)
            If LCase(AuthorsLastName(comp)) > LCase(AuthorsLastName(comp + 1)) Then
                'sorts author's last name
                tempAuthorsLastName = AuthorsLastName(comp)
                AuthorsLastName(comp) = AuthorsLastName(comp + 1)
                AuthorsLastName(comp + 1) = tempAuthorsLastName
                
                'sorts authors first name
                tempAuthorsFirstName = AuthorsFirstName(comp)
                AuthorsFirstName(comp) = AuthorsFirstName(comp + 1)
                AuthorsFirstName(comp + 1) = tempAuthorsFirstName
                
                'sorts authors middle initial
                   'sorts authors first name
                tempAuthorsMiddleName = AuthorsMiddleName(comp)
                AuthorsMiddleName(comp) = AuthorsMiddleName(comp + 1)
               AuthorsMiddleName(comp + 1) = tempAuthorsMiddleName
                
                'sorts Title
                tempTitle = Title(comp)
                Title(comp) = Title(comp + 1)
                Title(comp + 1) = tempTitle
                
                'sorts City published
                tempCityPublished = CityPublished(comp)
                CityPublished(comp) = CityPublished(comp + 1)
                CityPublished(comp + 1) = tempCityPublished
                
                'sorts publisher
                tempPublisher = Publisher(comp)
                Publisher(comp) = Publisher(comp + 1)
                Publisher(comp + 1) = tempPublisher
                
                'sorts year
                tempYear = Year(comp)
                Year(comp) = Year(comp + 1)
                Year(comp + 1) = tempYear
            End If
        Next comp
    Next pass
    
    'UsersName = InputBox("Please enter your name", "Enter your name")
    
    'enables the other buttons on the form
    cmdAddMore.Enabled = True
    cmdPrint.Enabled = True
    pos = 0
    
    'displays the users bibiliography
    picWorksCited.Cls
    picWorksCited.Print UsersName
    picWorksCited.Print Tab(50); "Works Cited"
    'MLA format
    If MLA = True Then
        For pos = 1 To ctr
            picWorksCited.Print ""
            picWorksCited.Print AuthorsLastName(pos) & ", " & AuthorsFirstName(pos) & ". " & Title(pos) & ". " & CityPublished(pos) & ": " & Publisher(pos) & ", " & Year(pos)
           
        Next pos
    'APA format
    Else
        For pos = 1 To ctr
            picWorksCited.Print ""
            picWorksCited.Print AuthorsLastName(pos) & ", " & AuthorsFirstName(pos) & ". " & AuthorsMiddleName(pos) & ". (" & Year(pos) & "). " & Title(pos) & ". " & CityPublished(pos) & ": " & Publisher(pos) & "."
        Next pos
    End If
End Sub


Private Sub Form_Load()
    'tells the user what file they have opened
    MsgBox "Your works cited, " & fileName & ".txt has been loaded"
End Sub
