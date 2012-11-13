VERSION 5.00
Begin VB.Form FrmDest 
   BackColor       =   &H8000000C&
   Caption         =   "SURF Destinations"
   ClientHeight    =   5085
   ClientLeft      =   3705
   ClientTop       =   3585
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   8295
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4875
      Left            =   0
      Picture         =   "FrmDest.frx":0000
      ScaleHeight     =   4815
      ScaleWidth      =   8250
      TabIndex        =   0
      Top             =   240
      Width           =   8310
      Begin VB.CommandButton Cmdclose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4440
         Width           =   975
      End
      Begin VB.Image Imgphil 
         Height          =   4815
         Left            =   5160
         ToolTipText     =   "Aisa & South Pacific"
         Top             =   0
         Width           =   3135
      End
      Begin VB.Image ImgEur 
         Height          =   4815
         Left            =   3240
         ToolTipText     =   "Europe & Africa"
         Top             =   0
         Width           =   1935
      End
      Begin VB.Image ImgNorth 
         Height          =   2295
         Left            =   0
         ToolTipText     =   "North America"
         Top             =   0
         Width           =   3255
      End
      Begin VB.Image Imgsouth 
         Height          =   2535
         Left            =   0
         ToolTipText     =   "Central & South America"
         Top             =   2280
         Width           =   3255
      End
   End
   Begin VB.Label lblchoose 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Please click an area on the map you would like to know more about."
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8295
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmDest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: SurfProject (SurfingProject.vbp)
'Form Name: frmDest (frmDest.frm)
'Author: Benjamin Luther
'Purpose of Form: this form displays a map of the world
                'and enables the user to click on an area of the map.
                'when the area is click it informs the user of the surf
                'spots and importance of the area.

Private Sub cmdclose_Click()
    FrmDest.Hide 'when button is clicked the form will disapear
End Sub

Private Sub ImgEur_Click() 'when this section of the picture is clicked on a message box appears displaying facts about the area
    MsgBox "Europe offers a wide variety of cultures, languages and currencies but also a large variety of surf spots. From the world-class left hander of Mundaca and the great beach breaks of Hossegor to cold big wave spots of Ireland. If you are going on an extended surf trip to some countries, you will need lots of different wetsuit and boards. Crowds can cause some problems in the busy summer months along with long flat spells. In Winter you will need a thick wetsuit on most countries except for the Canary island where board shorts are sufficient year-round. Price have risen drastically with the new Euro Currency in most countries. Africa is probably the least explored continent as far as surfing goes. There are lots of countries around the equator on both sides of Africa, so you will not need to travel with thick wetsuits, what you do need however is lots of patience since traveling around takes time there.", , "Europe & Africa"
End Sub

Private Sub ImgNorth_Click() 'when this section of the picture is clicked on a message box appears displaying facts about the area
    MsgBox "The surf in Northern America is quite abundant. The coasts are filled with anxious surfers up and down both the Pacific and Atlantic coasts. Lots of surf spots can be found here in all different kind of varieties: point breaks, beach breaks, big wave spots etc. It's hard to find secret spots in the US most of the spots have been well documented by travelers over the last 30 years. Canada also has some good surf spots but it's not your typical tropical surf experience here, expect to wear lots of neoprene year-round. Some of the most popular surfing destinations include California, and Hawaii. In California beaches there are many world famous beaches such as Maverics, Huntington Beach, and Santa Monica. Hawaii is also a world renound destination, said to be the birth place of surfing. There are many big name beaches in Hawaii including the infamous Pipeline, Jaws, and Waikiki. If you are in Hawaii you are in the surf capital of the world.", , "Northern America- U.S. & Canada"
End Sub

Private Sub Imgphil_Click() 'when this section of the picture is clicked on a message box appears displaying facts about the area
    MsgBox "Asia has only been a major surf destination since 25 years or so. Indonesia still being the most famous place in the area with world class left waves like Padang Padang, G-land and Uluwatu. But the Philippines and Sri Lanka have also been growing in popularity in the last 10 years. The weather is comfortable in most countries year round, the swell consistency can vary with the seasons. The Pacific area, a large area blessed with countless goof surf spots, from Australia, to the exclusive surf resorts on Fiji, to one of the longest left waves in the world of Raglan. Surf can be found year round here since lost of place pick up swell from many different directions.", , "Aisa, Australia & Philipines"
End Sub

Private Sub Imgsouth_Click() 'when this section of the picture is clicked on a message box appears displaying facts about the area
    MsgBox "South American surfing has adventure written all over it. If you can find a road that's not overgrown by dense jungle or cracked on a high plain desert chances are you may stumble across some very good, very uncrowded waves. South America is wide open. Juicy southern hemi storms send consistent surf to almost all coastlines on the Pacific side during the winter months (Northern Hemisphere's summer). The Atlantic coast is rich with wave potential too, although it may not be as consistent or big as the West Coast. Central America isa great area for surf trips: quality surf, secret spots and adverture is what you can expect here. Costa Rica being the most visited by US citizens because of the large variety of quality breaks, like Pavones and Ollie's point. Other great places to surf are El Salvador, Panama and of course the Caribbean version of Hawaii: Puerto Rico. Also plenty of surfing on some of the Caribbean islands like Barabados and Jamaica.", , "Central America & South America"
End Sub
