VERSION 5.00
Begin VB.Form frmBreeds 
   BackColor       =   &H00FF00FF&
   Caption         =   "Info on Breeds"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picpug 
      Height          =   2175
      Left            =   7920
      Picture         =   "frmBreeds.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   2475
      TabIndex        =   10
      Top             =   4680
      Width           =   2535
   End
   Begin VB.PictureBox piclab 
      Height          =   2655
      Left            =   5640
      Picture         =   "frmBreeds.frx":1AF4
      ScaleHeight     =   2595
      ScaleWidth      =   2115
      TabIndex        =   9
      Top             =   3600
      Width           =   2175
   End
   Begin VB.PictureBox picbull 
      Height          =   2535
      Left            =   3120
      Picture         =   "frmBreeds.frx":59D0
      ScaleHeight     =   2475
      ScaleWidth      =   2235
      TabIndex        =   8
      Top             =   4560
      Width           =   2295
   End
   Begin VB.PictureBox picgol 
      Height          =   3615
      Left            =   120
      Picture         =   "frmBreeds.frx":8745
      ScaleHeight     =   3555
      ScaleWidth      =   2715
      TabIndex        =   7
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton cmdPug 
      Caption         =   "Pug"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      TabIndex        =   6
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmdBull 
      Caption         =   "Bulldog"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3120
      TabIndex        =   5
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton cmdLab 
      Caption         =   "Labrador Retriever"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      TabIndex        =   4
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton cmdGolden 
      Caption         =   "Golden Retriever"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   3
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Back to the main screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "More Breeds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5760
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblInfobr 
      BackColor       =   &H00FF00FF&
      Caption         =   $"frmBreeds.frx":CC4B
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   9015
   End
End
Attribute VB_Name = "frmBreeds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Dogs(VB-project.vbp)
'Form Name: frmBreeds (frmBreeds.frm)
'Author: Libby Owen
'Date: Wednesday October 19
'Purpose of the project: to have the user interact with the program to decide
                        ' what kind of dog would be best for them.  The program
                        'will educate the user on some common types of dogs and
                        ' what to look for when picking one to take home
'Purpose of the form: this is where the user will go if they already know the breed
                        'they are looking to get.  They can get more info on the specific
                        'breed by clicking on the correct button.  If the bred they are
                        'looking for is not there then the name of a website where
                        ' they can go find the info on their own will pop up.


Private Sub cmdBull_Click()
'displays a message box that show the user info about the dog that they selected by pushing this button
MsgBox "You have picked a bulldog.  Bulldogs have usually have a kind temperament and tend to be very social with people.  The average weight of a girl is 40 lbs and the average weight of a boy is about 50 lbs.  For more information visit the American Kennel Club's website.", , "Bulldog"
End Sub

Private Sub cmdGolden_Click()
'displays a message box that show the user info about the dog that they selected by pushing this button
MsgBox "You have picked a Golden Retriever.  Golden Retrievers are in the sport group and can be used for hunting purposes.  They are larger dogs and the average weights for boys is 65-75 pounds, average weight for girls is 55-65 pounds. Golden Retrievers are Friendly, reliable, and trustworthy, so they intereact well with children. For more information visit the American Kennel Club's website", , "Golden Retriever"



End Sub

Private Sub cmdLab_Click()
'displays a message box that show the user info about the dog that they selected by pushing this button
    MsgBox "You have picked a Labrador Retriever.  Labs are in the sport group and are great for hunting.  They have a very friendly demeamor and make great family companions.  For more information visit the American Kennel Club's website.", , "Lab"
End Sub

Private Sub cmdMain_Click()
    frmFirstscreen.Show  'goes to main screen
    frmBreeds.Hide
End Sub

Private Sub cmdMore_Click()
'displays a message box that gives the user a website to find more info
MsgBox "Go to this website for more info; http://www.akc.org/breeds/index.cfm", , "Website"
End Sub

Private Sub cmdPug_Click()
'displays a message box that show the user info about the dog that they selected by pushing this button
MsgBox "You have picked a pug.  Pugs are in the toy group, they are very playful and make wonderful family pets.  For more information visit the American Kennel Club's website. ", , "Pug"

End Sub
