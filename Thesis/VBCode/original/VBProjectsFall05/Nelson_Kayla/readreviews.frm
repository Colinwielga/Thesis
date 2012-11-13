VERSION 5.00
Begin VB.Form frmreadreviews 
   BackColor       =   &H00FF8080&
   Caption         =   "Read Reviews; Project by Kayla Nelson"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   FillColor       =   &H00FF8080&
   ForeColor       =   &H00800080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picoutput 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   11.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   8475
      TabIndex        =   6
      Top             =   4440
      Width           =   8535
   End
   Begin VB.PictureBox piccity 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2565
      Left            =   7680
      Picture         =   "readreviews.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   1500
      TabIndex        =   4
      Top             =   1200
      Width           =   1530
   End
   Begin VB.PictureBox pickite 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2490
      Left            =   5520
      Picture         =   "readreviews.frx":1CA1
      ScaleHeight     =   2460
      ScaleWidth      =   1575
      TabIndex        =   3
      Top             =   1320
      Width           =   1605
   End
   Begin VB.PictureBox picfirst 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   3120
      Picture         =   "readreviews.frx":A51A
      ScaleHeight     =   2700
      ScaleWidth      =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   1830
   End
   Begin VB.PictureBox picmillion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   600
      Picture         =   "readreviews.frx":C188
      ScaleHeight     =   2970
      ScaleWidth      =   1950
      TabIndex        =   1
      Top             =   1080
      Width           =   1980
   End
   Begin VB.CommandButton cmdreturnmm 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8280
      Width           =   2775
   End
   Begin VB.Label lblbookreviews 
      BackColor       =   &H00FF8080&
      Caption         =   "Click on picture of book to see book reviews."
      BeginProperty Font 
         Name            =   "@MingLiU"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "frmreadreviews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Kayla's Book Club (MainMenu.vbp)
'Form Name: Read Reviews (frmreadreviews.frm)
'Author: Kayla Nelson
'Date: 10-27-05
'purpose of the form: This form allows the user to click on the picture of the book they would like to read reviews on.  3 reviews pop up into the picturebox.

Private Sub cmdreturnmm_Click() 'This closes the Read reviews form and brings the user back to the Main Menu.
    frmreadreviews.Hide
    frmmainmenu.Show
End Sub

Private Sub piccity_Click() 'When the user clicks on the book, it clears the picture box and then prints 3 reviews.
    picoutput.Cls
    picoutput.Print "Once again, Mr. Berendt makes erudite, inquisitive, nicely skeptical company as he"
    picoutput.Print "leads the reader through the shadows of what was heretofore better known as a"
    picoutput.Print "tourist attraction. [A]n urbane, beautifully fashioned book with much exotic charm."
    picoutput.Print "—Janet Maslin, The New York Times"
    picoutput.Print "   "
    picoutput.Print "An intriguing tour of mysterious Venice and its most fascinating residents."
    picoutput.Print "—Kirkus Reviews"
    picoutput.Print "   "
    picoutput.Print "This is journalism at its most accomplished; it is creative nonfiction as"
    picoutput.Print "enveloping and heart embracing as good fiction. —Booklist"
End Sub

Private Sub picfirst_Click() 'When the user clicks on the picture, it first clears the picture box and then prints 3 reviews.
    picoutput.Cls
    picoutput.Print "This volume truly brings a satisfying end to Jeremy and Lexie's story. You won't want"
    picoutput.Print "to miss it if you like reading Mr. Sparks."
    picoutput.Print
    picoutput.Print "This book was excellent. Nicholas Sparks really knows how to twist a person's heart!"
    picoutput.Print "This book was wonderful but left me a little depressed at the end. However, Sparks has"
    picoutput.Print "a balance of sadness and sweetness at the end. I found myself crying and smiling as I"
    picoutput.Print "read the last few pages. Excellent book!"
    picoutput.Print
    picoutput.Print "Nicholas Sparks is known for his ability to touch the heart of the reader through the"
    picoutput.Print "relationships of his characters. As with True Believer, At First Sight is a solid book,"
    picoutput.Print "but not one of Sparks' best. Here we are brought up to speed with characters we first"
    picoutput.Print "met in True Believer, but here the realistic character portrayals that Sparks usually"
    picoutput.Print "captures are somewhat lacking."
End Sub

Private Sub pickite_Click() 'When the user clicks on the picture it first clears the picture box then prints 3 reviews.
    picoutput.Cls
    picoutput.Print "A wonderful work… This is one of those unforgettable stories that stay with you for"
    picoutput.Print "years. All the great themes of literature and of life are the fabric of this"
    picoutput.Print "extraordinary novel: love, honor, guilt, fear redemption…It is so powerful that for"
    picoutput.Print "a long time everything I read after seemed bland."
    picoutput.Print "   "
    picoutput.Print "Truly a thought-provoking novel one will not put down until its last page."
    picoutput.Print "   "
    picoutput.Print "Brilliant. . . both as a political chronicle and a deeply personal tale about how"
    picoutput.Print "childhood choices affect our adult lives. Publishers Weekly, starred review"
End Sub
Private Sub picmillion_Click() ' When the user first clicks on the picture it clears the message box and then prints 3 reviews of that book.
    picoutput.Cls
    picoutput.Print "Thoroughly engrossing... Hard-bitten existentialism bristles on every page... Frey's"
    picoutput.Print "prose is muscular and tough, ideal for conveying extreme physical anguish and steely"
    picoutput.Print "determination. Entertainment Weekly"
    picoutput.Print "   "
    picoutput.Print "Insistent as it is demanding.... A story that cuts to the nerve of addiction by"
    picoutput.Print "clank-clank-clanking through the skull of the addicted... A critical milestone in"
    picoutput.Print "modern literature. Orlando Weekly"
    picoutput.Print "   "
    picoutput.Print "One of the most compelling books of the year... Incredibly bold... Somehow"
    picoutput.Print "accomplishes what three decades' worth of cheesy public service announcements and"
    picoutput.Print "after-school specials have failed to do: depict hard-core drug addiction as the"
    picoutput.Print "self-inflicted apocalypse that it is. The New York Post"
End Sub

