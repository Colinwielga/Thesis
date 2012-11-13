VERSION 5.00
Begin VB.Form frmMuseums 
   BackColor       =   &H00000040&
   Caption         =   "Museum Attractions"
   ClientHeight    =   11025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11025
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTubeResults 
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   5775
      TabIndex        =   9
      Top             =   9840
      Width           =   5775
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9480
      Width           =   2415
   End
   Begin VB.CommandButton cmdGoToPopularAttractions 
      BackColor       =   &H00808080&
      Caption         =   "Return to Popular Attractions Page"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8640
      Width           =   2415
   End
   Begin VB.CommandButton cmdGoToHome 
      BackColor       =   &H00808080&
      Caption         =   "Return to Home Page"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7800
      Width           =   2415
   End
   Begin VB.TextBox txtResults 
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   7320
      Width           =   8175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   3480
      ScaleHeight     =   6615
      ScaleWidth      =   8175
      TabIndex        =   4
      Top             =   480
      Width           =   8175
   End
   Begin VB.CommandButton cmdTateModern 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tate Modern"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdBritish 
      BackColor       =   &H00C0C0C0&
      Caption         =   "British Museum"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdNaturalHistory 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Natural History Museum"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton cmdNationalGallery 
      BackColor       =   &H00C0C0C0&
      Caption         =   "National Gallery"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frmMuseums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: London Attractions
'Form Name: Museums
'Author: Heather Arnhalt
'Date Written: October 18, 2009
'Objective: The user can click on the attraction they would like to learn more about and a picture of that attraction is displayed
'along with information about that attraction.

Private Sub cmdBritish_Click()
    'clear both picture boxes
    picResults.Cls
    picTubeResults.Cls
    
    'Display a picture of the attraction
    picResults.Picture = LoadPicture(App.Path & "\London_Pictures\BritishMuseum.jpg")
    
    'Print information about the attraction in the text results box
    txtResults.Text = "If you wanted to thoroughly explore the British Museum, it would take months, if not years. Over seven millions objects from all over the world are housed in this impressive museum of human history and culture (many of the artifacts are stored underneath the museum due to lack of space). Founded in 1753, displays ranging from prehistoric to modern times were primarily based on the collections of physician and scientist, Sir Hans Sloane. Notable objects include the Parthenon Sculptures (also sometimes called the Elgin Marbles), the Rosetta Stone, the Sutton Hoo and Mildenhall treasures, and the Portland Vase. The hieroglyphics and classical sculptures are instantly recognisable and world famous. The museum's collection of ancient Egyptian mummies is world famous as well."

    'Print the tube station the attraction is nearest to in the PicTubeResults box
    picTubeResults.Print "Tube Station: Tottenham Court Road"
End Sub

Private Sub cmdNationalGallery_Click()
    'clear both picture boxes
    picResults.Cls
    picTubeResults.Cls
    
    'Display a picture of the attraction
    picResults.Picture = LoadPicture(App.Path & "\London_Pictures\NationalGallery.jpg")
    
    'Print information about the attraction in the text results box
    txtResults.Text = "The National Gallery dominates London's Trafalgar Square with its neo-classical columns and portico designed by William Wilkins adjoining the square on its east side where it has been pedestrianised. Some of the finest examples of European art, ranging from 1260 to 1900, are included among the 2,300 paintings filling its halls and rooms. Holbein's 'The Ambassadors', 'The Hay Wain' by Constable, and Jan Van Eyck's 'Arnolfini Marriage' are just some of the major attractions. Works on display also include those of Botticelli, Monet, Constable, Van Gogh and Rembrandt. This really is the place to come for top quality artwork spanning a wide spectrum of styles and periods."

    'Print the tube station the attraction is nearest to in the PicTubeResults box
    picTubeResults.Print "Tube Station: Charing Cross"
End Sub

Private Sub cmdNaturalHistory_Click()
    'clear both picture boxes
    picResults.Cls
    picTubeResults.Cls
    
    'Display a picture of the attraction
    picResults.Picture = LoadPicture(App.Path & "\London_Pictures\NaturalHistory.jpg")
    
    'Print information about the attraction in the text results box
    txtResults.Text = "The Natural History Museum is sure to impress even the most jaded of children. This ornate museum is home to more than 70 million specimens from across the natural world, including insects, fossils and rocks. The Dinosaur gallery is one of the most popular exhibits in the museum, with a giant T.rex, the horned Triceratops and the fossilised skin of an Edmontosaurus."

    'Print the tube station the attraction is nearest to in the PicTubeResults box
    picTubeResults.Print "Tube Station: South Kensington"
End Sub

Private Sub cmdTateModern_Click()
    'clear both picture boxes
    picResults.Cls
    picTubeResults.Cls
    
    'Display a picture of the attraction
    picResults.Picture = LoadPicture(App.Path & "\London_Pictures\TateModern.jpg")
    
    'Print information about the attraction in the text results box
    txtResults.Text = "Located along the banks of the River Thames in the former Bankside Power Station, originally designed by Sir Giles Gilbert Scott in 1947, the architect of Battersea Power Station, Tate Modern opened to great acclaim in 2000. Since then it has welcomed millions of visitors through its imposing doors. The gallery pays homage to art from 1900 to the present day while the awesome Turbine Hall creates a stunning entrance and a vast space, used to display temporary installations on a grand scale. There are three levels of galleries enclosed by a spectacular two-storey glass roof that provides fantastic views of London and a great caf. It offers a full set of iconic twentieth century artists, from Matisse to Moore, Dali to Picasso. Justifiably the most popular art gallery in Europe."

    'Print the tube station the attraction is nearest to in the PicTubeResults box
    picTubeResults.Print "Tube Station: Southwark"
End Sub

Private Sub cmdGoToHome_Click()
    'hides the Museum form and shows the Home Page form
    frmMuseums.Hide
    frmHomePage.Show
End Sub

Private Sub cmdGoToPopularAttractions_Click()
    'Hides the Museum form and shows the Popular Attractions form
    frmMuseums.Hide
    frmPopularAttractions.Show
End Sub

Private Sub cmdQuit_Click()
    'ends the program
    End
End Sub
