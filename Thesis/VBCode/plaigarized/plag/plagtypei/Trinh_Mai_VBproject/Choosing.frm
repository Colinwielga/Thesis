VERSION 5.00
Begin VB.Form frmChoosing
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   11430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   ScaleHeight     =   11430
   ScaleWidth      =   14985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit
      Caption         =   "Exit"
      BeginProperty Font
         Name            =   "Elephant"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   4
      Top             =   8640
      Width           =   3375
   End
   Begin VB.PictureBox picMaga
      AutoSize        =   -1  'True
      Height          =   4560
      Left            =   8040
      Picture         =   "Choosing.frx":0000
      ScaleHeight     =   4500
      ScaleWidth      =   4500
      TabIndex        =   3
      Top             =   3480
      Width           =   4565
   End
   Begin VB.PictureBox picResults
      AutoSize        =   -1  'True
      Height          =   4560
      Left            =   2520
      Picture         =   "Choosing.frx":7FB7
      ScaleHeight     =   4500
      ScaleWidth      =   4500
      TabIndex        =   2
      Top             =   3480
      Width           =   4560
   End
   Begin VB.CommandButton cmdListMagazines
      BackColor       =   &H0080FFFF&
      Caption         =   "Magazines"
      BeginProperty Font
         Name            =   "Engravers MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   1
      Top             =   1800
      Width           =   3375
   End
   Begin VB.CommandButton cmdListBooks
      BackColor       =   &H00FFFFFF&
      Caption         =   "Books"
      BeginProperty Font
         Name            =   "Engravers MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      MaskColor       =   &H00808080&
      TabIndex        =   0
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label1
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Welcome to Trinh's bookstore"
      BeginProperty Font
         Name            =   "Ravie"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1275
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   10275
   End
   Begin VB.Image Image1
      Height          =   11415
      Left            =   0
      Picture         =   "Choosing.frx":E5DA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14895
   End
End
Attribute VB_Name = "frmChoosing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '  hey  words  comment  spaces  stuff  things


  Dim MyLogo As String  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
  Dim ctrBook As Integer  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
  Dim ctrMaga As Integer  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things



  Private Sub Form_Load()  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
      'check if the file exist  '  hey  words  comment  spaces  stuff  things
        '  hey  words  comment  spaces  stuff  things
        '  hey  words  comment  spaces  stuff  things


        '  hey  words  comment  spaces  stuff  things
      If Dir("ExportItems.txt") <> "" Then  '  hey  words  comment  spaces  stuff  things
        '  hey  words  comment  spaces  stuff  things
          Kill "ExportItems.txt"  '  hey  words  comment  spaces  stuff  things


            '  hey  words  comment  spaces  stuff  things
      End If  '  hey  words  comment  spaces  stuff  things



        '  hey  words  comment  spaces  stuff  things
  End Sub  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
  Private Sub cmdExit_Click()  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
  End  '  hey  words  comment  spaces  stuff  things


    '  hey  words  comment  spaces  stuff  things
  End Sub  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things


    '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
  Private Sub cmdListBooks_Click()  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
  frmChoosing.Hide  '  hey  words  comment  spaces  stuff  things


    '  hey  words  comment  spaces  stuff  things
  frmDetails.Show  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
  End Sub  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things


  Private Sub cmdListMagazines_Click()  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
  frmChoosing.Hide  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things


  frmMagazine.Show  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
  End Sub  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
  Private Sub picMaga_Click()  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things


      'Dynamic picture loading  '  hey  words  comment  spaces  stuff  things
        '  hey  words  comment  spaces  stuff  things
      ctrMaga = (ctrMaga + 1) Mod 3  '  hey  words  comment  spaces  stuff  things
        '  hey  words  comment  spaces  stuff  things



        '  hey  words  comment  spaces  stuff  things
        '  hey  words  comment  spaces  stuff  things
      MyLogo = "Maga" & ctrMaga  '  hey  words  comment  spaces  stuff  things


        '  hey  words  comment  spaces  stuff  things
      picMaga.Picture = LoadPicture(App.Path & "\magazines\" & MyLogo & ".jpg")  '  hey  words  comment  spaces  stuff  things
        '  hey  words  comment  spaces  stuff  things
  End Sub  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things


    '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
  Private Sub picResults_Click()  '  hey  words  comment  spaces  stuff  things



    '  hey  words  comment  spaces  stuff  things
      ctrBook = (ctrBook + 1) Mod 3  '  hey  words  comment  spaces  stuff  things
        '  hey  words  comment  spaces  stuff  things


        '  hey  words  comment  spaces  stuff  things
        '  hey  words  comment  spaces  stuff  things
      MyLogo = "book" & ctrBook  '  hey  words  comment  spaces  stuff  things

        '  hey  words  comment  spaces  stuff  things
      picResults.Picture = LoadPicture(App.Path & "\pic\" & MyLogo & ".jpg")  '  hey  words  comment  spaces  stuff  things
        '  hey  words  comment  spaces  stuff  things


    '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
  End Sub  '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things


    '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
    '  hey  words  comment  spaces  stuff  things
