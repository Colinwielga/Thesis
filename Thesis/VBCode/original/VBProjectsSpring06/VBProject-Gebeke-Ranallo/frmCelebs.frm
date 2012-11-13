VERSION 5.00
Begin VB.Form frmCelebs 
   Caption         =   "More about Celeb's Style"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   Picture         =   "frmCelebs.frx":0000
   ScaleHeight     =   8325
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1815
   End
   Begin VB.PictureBox picJLo 
      BackColor       =   &H00000000&
      Height          =   2775
      Left            =   6480
      Picture         =   "frmCelebs.frx":938F
      ScaleHeight     =   2715
      ScaleWidth      =   2595
      TabIndex        =   4
      Top             =   5160
      Width           =   2655
   End
   Begin VB.PictureBox picJennifer 
      BackColor       =   &H00000000&
      Height          =   3615
      Left            =   6000
      Picture         =   "frmCelebs.frx":B23C
      ScaleHeight     =   3555
      ScaleWidth      =   3555
      TabIndex        =   3
      Top             =   1200
      Width           =   3615
   End
   Begin VB.PictureBox picParis 
      BackColor       =   &H00000000&
      Height          =   3975
      Left            =   600
      Picture         =   "frmCelebs.frx":1094E
      ScaleHeight     =   3915
      ScaleWidth      =   3915
      TabIndex        =   2
      Top             =   4080
      Width           =   3975
   End
   Begin VB.PictureBox picReese 
      BackColor       =   &H00000000&
      Height          =   2775
      Left            =   360
      Picture         =   "frmCelebs.frx":17007
      ScaleHeight     =   2715
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label lblNames 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenna Gebeke ~ Katie Ranallo"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   8040
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click on your favorite celeb to read more..."
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "frmCelebs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Form Name: Celebs
'Form Objective: To allow the user to view and read more about her favorite celebs and to read more about which one matches her style.

Private Sub cmdExit_Click()
'This command button allows the user to return to the Startup page.
    frmCelebs.Hide
    frmStart.Show
End Sub

Private Sub cmdCelebBios_Click()
End Sub

Private Sub Form_Load()
    frmCelebs.Caption = "Welcome " & userName & "  - More about Celeb's Style"
End Sub

'By clicking on the picture of Reese Witherspoon, the user is able to read about Reese's style.
Private Sub picReese_Click()
    MsgBox "Reese Witherspoon is whom PEOPLE style director Susan Kaufman likens to a modern-day Grace Kelly, is known for her classic, yet elegant look.  On the red carpet, she's been wearing very feminine styles, classic yet elegant displays of bows and tulle. Her everyday look is a more familiar one for Reese. While on the go, running errands, and playing with her kids, she still manages to be a classic trend setter displaying a relaxed image. Reese is very family oriented and portrays this through her clothing. She is most comfortable in jeans, a shirt, and minimal make-up on a day to day basis and still manages to embody the elegance she displays on the red carpet.", , "Reese's Style!"
End Sub
'By clicking on the picture of Paris Hilton, the user is able to read more about Paris's style.
Private Sub picParis_Click()
    MsgBox "Paris Hilton, the heiress known for controversial press coverage, and skimpy, trendy clothing. This trend setting behavior has even graced the English dictionary, 'That's hot!', the phrase librally used by Hilton has successfully been coined and is, like the creater of it, a pop culture icon. Hilton is using her pop culture icon status to enter the fashion world as well. She has always been known for her trendy, controvercial and fashion forward style and intends to establish this unique trend in brand by marketing clothing with the phrase 'That's Hot!' prominently emblazoned on the garments.", , "Paris's Style!"
End Sub
'By clicking on the picture of Jennifer Garner, the user is able to read more about Jennifer's style.
Private Sub picJennifer_Click()
    MsgBox "Due to Jennifer Garner's ballet background and her intense workout schedule to stay in shape for her 'Alias' sitcom she maintains a fit physique which matches her naturally athletic mannerism and style. Due to the roles she has played in many films she and her style have been type casted. She displays a flirty and sweet style in the movie '13 Going On 30', and a sporty athletic style from her role in 'Alias'. Both styles combine in her everyday life and embody the actress for who she really is.", , "Jennifer's Style!"
End Sub
'By clicking on the picture of Jennifer Lopez, the user is able to read more about JLo's style.
Private Sub picJLo_Click()
    MsgBox "Jennifer Lopez has come a long way from her trendy hard edge style in the Bronx, so far that she now owns her own sexy, trendy, yet chic clothing line. JLO by Jennifer Lopez is a line that hits every edge of the fashion arena. She is one of the pioneers in the curvy clothing industry as she is well known for her curvy dancers physique. While she has made waves in the fashion industry she has also made a mark on the music and film industries as well. Her presence in all of these industries has given her an Icon status that carries her name and image across cultures and genres.", , "JLO's Style!"
End Sub



