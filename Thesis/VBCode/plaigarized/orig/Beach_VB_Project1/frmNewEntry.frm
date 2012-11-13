VERSION 5.00
Begin VB.Form frmNewEntry 
   Caption         =   "Form2"
   ClientHeight    =   7905
   ClientLeft      =   3375
   ClientTop       =   2280
   ClientWidth     =   12615
   LinkTopic       =   "Form2"
   ScaleHeight     =   7905
   ScaleWidth      =   12615
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   1095
      Left            =   960
      TabIndex        =   0
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   1095
      Left            =   960
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtContent 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2400
      Width           =   5535
   End
   Begin VB.TextBox txtSubject 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9000
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtSource 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   7575
      Left            =   480
      Picture         =   "frmNewEntry.frx":0000
      ScaleHeight     =   7515
      ScaleWidth      =   11475
      TabIndex        =   5
      Top             =   120
      Width           =   11535
      Begin VB.Label lblContent 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Content Here starting with the Page Number:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   8
         Top             =   1920
         Width           =   4935
      End
      Begin VB.Label lblSubject 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblSourceName 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Source Name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmNewEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NoteCard VB CS130 Project
'FrmNewEntry
'LauraBeach
'Written between February 1-24, 2010
'This form is perhaps the most difficult of the three forms...I use functions that were not discussed in class.
'The point of this form is to allow the user to input data into a file via the append function. The data will be saved and allow the user to search it under the frmSearch form
'There are three input boxes, all of which must be full in order for the entry to be saved via the append function. I used message boxes to inform the User if a text box
'is empty AND which box that is...another message box tells the User that the data has been saved...If the data is saved (Because all of the text boxes contain data) the
'message box appears after the save button is clicked and the text boxes are cleared. I had to search via the internet to find the code to clear the text boxes as I was unable
'to find the code in the VB examples on the N:\ drive. I got the code from a website -- freevbcode.com -- I also cite this website below where I use that code.
'The back button reads the code and takes the User back to the Main page (frmMain) so that the user may go to the search page (frmSearch) or Quit...cmdQuit on the main page
'After cmdSave is clicked...the User may input another 'notecard' (more data) and it will save as many times as the User wishes...


Option Explicit
Private Sub Command1_Click()

End Sub

Private Sub cmdBack_Click()

    'Return to Main Form
    frmNewEntry.Hide
    frmMain.Show
    
    'Read the File so that when the search is performed the File will not have to be read each time you search, the file will
    'Only have to be opened once
    
    'Open the File
    Open App.Path & "\notecarddata.txt" For Input As #1
    
    'Start the counter so that the data will be read into a specific slot in the array and we know what slot that is
    
    CTR = 0
    
    'Start the loop so that the file will be read until the end
    
    Do While Not EOF(1)
        CTR = CTR + 1 'This starts the actual counter
        Input #1, SourceName(CTR), Subject(CTR), Content(CTR) 'Label the data to put in the proper arrays.
    Loop 'Go back and read the next line
    
    Close #1
    'Now the data has been read and is available for the next form
    
End Sub

Private Sub cmdSave_Click()

    'The Variables need to be dimmed here...They are not in arrays yet because they haven't been read into them...
    'This is just putting data into the file, not taking data from a file and putting it into an array
    
    Dim SourceNameEntry As String, SubjectEntry As String, ContentEntry As String, txt As Control
    
    
    'Assign the variables with data
    SourceNameEntry = txtSource.Text
    SubjectEntry = txtSubject.Text
    ContentEntry = txtContent.Text
    
    
    Open App.Path & "\notecarddata.txt" For Append As #1      'note the word "Append" in this line of code
    'Provide for possible errors if nothing is typed into the open textboxes
        If SourceNameEntry = "" Then
            MsgBox "Error, must insert a Source", , Error
             If SubjectEntry = "" Then
                MsgBox "Error, must insert a Subject Title", , Error
                    If ContentEntry = "" Then
                        MsgBox "Error, must insert Content", , Error
                    End If
             End If
        Else
            Write #1, SourceNameEntry, SubjectEntry, ContentEntry
            'The data written will be appended to the end of the file.
            
     'Let the user know that the data has indeed been written to the file...
     
        MsgBox "Thank You, Your Data Has Been Saved."
            
        End If
    Close
    
        
    'Clear the text boxes so that a new notecard/more data can be entered...I got this code from online -- freevbcode.com as I could not find it in the book...
    
        For Each txt In frmNewEntry
        
              If TypeOf txt Is TextBox Then txt.Text = ""

        Next

        
    

End Sub

Private Sub lblSource_Click()

End Sub

