VERSION 5.00
Begin VB.Form frmOtherAttractions 
   BackColor       =   &H00400040&
   Caption         =   "Other Attractions"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   12255
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   7800
      Width           =   2415
   End
   Begin VB.PictureBox picTubeResults 
      BackColor       =   &H00400040&
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
      Left            =   3360
      ScaleHeight     =   495
      ScaleWidth      =   4935
      TabIndex        =   6
      Top             =   9840
      Width           =   4935
   End
   Begin VB.TextBox txtResults 
      BackColor       =   &H00400040&
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
      Height          =   2895
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   7320
      Width           =   8175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   3360
      ScaleHeight     =   6615
      ScaleWidth      =   8175
      TabIndex        =   4
      Top             =   480
      Width           =   8175
   End
   Begin VB.CommandButton cmdHarrods 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Harrods Department Store"
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
   Begin VB.CommandButton cmdCathedral 
      BackColor       =   &H00FFC0C0&
      Caption         =   "St. Paul's Cathedral"
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
   Begin VB.CommandButton cmdLondonEye 
      BackColor       =   &H00FFC0C0&
      Caption         =   "London Eye"
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
   Begin VB.CommandButton cmdHydePark 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Hyde Park"
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
Attribute VB_Name = "frmOtherAttractions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: London Attractions
'Form Name: Other Attractions
'Author: Heather Arnhalt
'Date Written: October 18, 2009
'Objective: The user can click on the attraction they would like to learn more about and a picture of that attraction is displayed
'along with information about that attraction.

Private Sub cmdCathedral_Click()
    'clear both picture boxes
    picResults.Cls
    picTubeResults.Cls
    
    'Display a picture of the attraction
    picResults.Picture = LoadPicture(App.Path & "\London_Pictures\StPauls.jpg")
    
    'Print information about the attraction in the text results box
    txtResults.Text = "Being the cathedral of the capital city, St Paul's is officially the spiritual home of Great Britain. The funerals of Lord Nelson, the Duke of Wellington and Sir Winston Churchill were conducted inside the church's fortress-like walls, as was the elaborate fairy-tale wedding of Prince Charles and Lady Diana Spencer. Built by court architect, Sir Christopher Wren, after the Great Fire of London in 1666, the cathedral miraculously survived the Blitz in World War Two when most of the surrounding area was flattened by German bombing raids. It consequently served to act as an inspirational symbol of British strength in the nation's darkest hour. Three curving galleries lead up to the dome - one of the largest in the world and one of the best viewing points in the City."

    'Print the tube station the attraction is nearest to in the PicTubeResults box
    picTubeResults.Print "Tube Station: St.Paul's"
End Sub

Private Sub cmdHarrods_Click()
    'clear both picture boxes
    picResults.Cls
    picTubeResults.Cls
    
    'Display a picture of the attraction
    picResults.Picture = LoadPicture(App.Path & "\London_Pictures\Harrods.jpg")
    
    'Print information about the attraction in the text results box
    txtResults.Text = "Britain's most famous store and possibly the most famous store in the world, Harrods features on many tourist 'must-see' lists - and with good reason. Its humble beginnings date back to 1849, when Henry Charles Harrod opened a small grocery shop that emphasised impeccable service over value. Today, it occupies a vast site in London's fashionable Knightsbridge and boasts a phenomenal range of products from pianos and cooking pans to pets and perfumery with a large Hair and Beauty department its crowning glory on the top floor. The Food Hall is ostentatious to the core and mouth-wateringly exotic, and the store as a whole is well served with restaurants. At Christmas time, Harrods boasts an enchanting Santa's Grotto for the kids and an extensive range of festive decorations."

    'Print the tube station the attraction is nearest to in the PicTubeResults box
    picTubeResults.Print "Tube Station: Knightsbridge"
End Sub

Private Sub cmdHydePark_Click()
    'clear both picture boxes
    picResults.Cls
    picTubeResults.Cls
    
    'Display a picture of the attraction
    picResults.Picture = LoadPicture(App.Path & "\London_Pictures\HydePark.bmp")
    
    'Print information about the attraction in the text results box
    txtResults.Text = "Technically two different parks, Hyde Park and Kensington Gardens are in practical terms one huge, merging expanse. The 'split' dates back to 1728 when Queen Caroline, wife of George II, took almost 300 acres from Hyde Park to form Kensington Gardens. The 350 acres that remained has become one of London's best-loved parks. Almost every kind of outdoor pursuit takes place within its lush green landscape. Horse riding, rollerblading, bowls, putting and tennis are all catered for while informal games of cricket, rounders and frizbee spring up on the area to the south of the park known as The Sports Field. A number of famous London attractions are also housed within this central space. Hyde Park boasts Speakers' Corner, the Princess of Wales Memorial Fountain and the Serpentine Boating Lake - another feature owed to Queen Caroline who started a new landscaping trend by designing this natural-looking lake."

    'Print the tube station the attraction is nearest to in the PicTubeResults box
    picTubeResults.Print "Tube Station: Knightsbridge"
End Sub

Private Sub cmdLondonEye_Click()
    'clear both picture boxes
    picResults.Cls
    picTubeResults.Cls
    
    'Display a picture of the attraction
    picResults.Picture = LoadPicture(App.Path & "\London_Pictures\LondonEye.jpg")
    
    'Print information about the attraction in the text results box
    txtResults.Text = "The London Eye (or Millennium Wheel) elbowed its way onto the capital's tourist scene as one of the statement pieces to mark the turn of the century (see also The Millennium Bridge) and quickly became a definitive part of the London experience. This spectacularly streamlined riverside wheel stands an impressive 135 meters tall and allows inhabitants of its sleek, modern, totally see-through glass-pods an unrivalled 360 degree view of London and beyond. The Houses of Parliament, Canary Wharf, Big Ben, the glorious old winding Father Thames and Windsor Castle are just a few of the 55 attractions that can be admired from the top of the arc. On a clear day the view extends to 25 miles in each direction. It moves at a slow but steady pace, taking 30 minutes to complete its flight."

    'Print the tube station the attraction is nearest to in the PicTubeResults box
    picTubeResults.Print "Tube Station: Waterloo"
End Sub

Private Sub cmdGoToHome_Click()
    'Show the Home Page form and hide the Other Attractions form
    frmHomePage.Show
    frmOtherAttractions.Hide
End Sub

Private Sub cmdGoToPopularAttractions_Click()
    'Hide the Other Attractions form and show the Popular Attractions form
    frmOtherAttractions.Hide
    frmPopularAttractions.Show
End Sub

Private Sub cmdQuit_Click()
    'end the program
    End
End Sub

