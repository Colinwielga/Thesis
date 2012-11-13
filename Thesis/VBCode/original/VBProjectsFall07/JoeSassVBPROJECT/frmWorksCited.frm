VERSION 5.00
Begin VB.Form frmWorksCited 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Here is your finished bibliography"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save your work"
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Click here to view your formated Works Cited Page"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   8175
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   8040
      Width           =   975
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print your Works Cited Page"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   8040
      Width           =   2415
   End
   Begin VB.PictureBox picWorksCited 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   120
      ScaleHeight     =   7455
      ScaleWidth      =   8175
      TabIndex        =   0
      Top             =   480
      Width           =   8175
   End
End
Attribute VB_Name = "frmWorksCited"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pos As Integer
Dim saved As Boolean


Private Sub cmdAddMore_Click()
    'brings the user back to the appropriate format form to add more sources
    frmWorksCited.Hide
    If MLA = True Then
        frmMLABook.Show
    Else
        frmAPABook.Show
    End If
End Sub

Private Sub cmdPrint_Click()
    'Prints the user's works cited page on a printer. I learned how to do this with help from http://www.devarticles.com/c/a/Visual-Basic/Printing-With-Visual-Basic/1/
    Printer.Print ""
    Printer.Print UsersName
    Printer.Print Tab(75); "Works Cited"
    
    'checks which format it is, and then runs through the loop printing each citation in the right format
    If MLA = True Then 'mla format
        For pos = 1 To ctr
           Printer.Print ""
           Printer.Print AuthorsLastName(pos) & ", " & AuthorsFirstName(pos) & ". " & Title(pos) & ". " & CityPublished(pos) & ": " & Publisher(pos) & ", " & Year(pos) & "."
           
        Next pos
    Else
        For pos = 1 To ctr 'apa format
            Printer.Print ""
            Printer.Print AuthorsLastName(pos) & ", " & AuthorsFirstName(pos) & ". " & AuthorsMiddleName(pos) & ". (" & Year(pos) & "). " & Title(pos) & ". " & CityPublished(pos) & ": " & Publisher(pos) & "."
        Next pos
    End If
    'tells the printer that there is no more data to print
    Printer.EndDoc
End Sub

Private Sub CmdQuit_Click()
    'checks to see if the user has already saved their work
    If saved = False Then 'if they haven't saved
        frmQuit.Show
    Else 'quits the program if they have
        End
    End If
End Sub

Private Sub cmdSave_Click()
    'allows the user to save their bibliography. I learned to do this with help from http://www.developerfusion.co.uk/show/37/4/
    Dim pos As Integer
    Dim fileName As String
    pos = 0
    
    'asks the user for a filename to save their work as
    fileName = InputBox("Please type the file name that you wish to save your works cited as:")
    
    'opens the file or creates a new one if it does not already exist
    Open App.Path & "\saved\" & fileName & ".txt" For Output As #1
    
    'loops through writing the arrays to the file
    Do Until pos = ctr
        pos = pos + 1
        Write #1, AuthorsLastName(pos), AuthorsFirstName(pos), AuthorsMiddleName(pos), Title(pos), CityPublished(pos), Publisher(pos), Year(pos)
    Loop
    
    Close #1
    
    'displays for the user the filename they have saved their work as
    MsgBox "Your file has been saved in the \saved folder as " & fileName & ".txt"
    
    'allows the computer to know that the user has saved
    saved = True
End Sub

Private Sub cmdView_Click()
    'sorts and displays the bibliography
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
    
    'enables the other buttons on the form now that the data has been loaded
    cmdAddMore.Enabled = True
    cmdPrint.Enabled = True
    
        
    pos = 0
    
    'displays the works cited page in the picture box
    picWorksCited.Cls
    picWorksCited.Print UsersName
    picWorksCited.Print Tab(50); "Works Cited"
    
    'checks to see which format the user used and then displays the bibliography in the appropriate format
    If MLA = True Then 'MLA format
        For pos = 1 To ctr
            picWorksCited.Print ""
            picWorksCited.Print AuthorsLastName(pos) & ", " & AuthorsFirstName(pos) & ". " & Title(pos) & ". " & CityPublished(pos) & ": " & Publisher(pos) & ", " & Year(pos)
        Next pos
    Else
        For pos = 1 To ctr 'APA format
            picWorksCited.Print ""
            picWorksCited.Print AuthorsLastName(pos) & ", " & AuthorsFirstName(pos) & ". " & AuthorsMiddleName(pos) & ". (" & Year(pos) & "). " & Title(pos) & ". " & CityPublished(pos) & ": " & Publisher(pos) & "."
        Next pos
    End If
End Sub

Private Sub Form_Load()
    'tells the computer the user has not yet saved their work
    saved = False
End Sub


