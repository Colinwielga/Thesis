VERSION 5.00
Begin VB.Form frmTedBio 
   Caption         =   "Ted's Biography"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form3"
   Picture         =   "frmTedBio.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquickfacts 
      BackColor       =   &H000000FF&
      Caption         =   "Click for Quick Facts about Ted"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   2535
   End
   Begin VB.PictureBox picresults 
      AutoSize        =   -1  'True
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   9570
      Left            =   7680
      ScaleHeight     =   9510
      ScaleWidth      =   5715
      TabIndex        =   2
      Top             =   480
      Width           =   5775
   End
   Begin VB.CommandButton cmdClickBio 
      BackColor       =   &H000000FF&
      Caption         =   "Click for Ted's Biography"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.CommandButton cmdTedBioBack 
      BackColor       =   &H000000FF&
      Caption         =   "Click to go back to Contents"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8880
      Width           =   4815
   End
End
Attribute VB_Name = "frmTedBio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClickBio_Click()
    picresults.Cls
    picresults.Print ""
    picresults.Print ""
    picresults.Print "  TED WILLIAMS"
    picresults.Print "  ******************************************************************************************"
    picresults.Print ""
    picresults.Print "  Williams remains the last man to hit over .400"
    picresults.Print "  for a complete major league baseball season."
    picresults.Print "  He finished the 1941 season with a .406 batting average"
    picresults.Print "  ,going 6-for-8 during a season-ending double-header "
    picresults.Print "  to push himself over the .400 mark."
    picresults.Print ""
    picresults.Print " Williams played for the Boston Red Sox from 1939 to 1942"
    picresults.Print " and from 1946 to 1960."
    picresults.Print ""
    picresults.Print "  During World War II he served in the U.S. Navy."
    picresults.Print "  His 1952 and 1953 seasons were interrupted "
    picresults.Print "  by his service as a Marine Corps pilot; he flew 39 combat missions in Korea. "
    picresults.Print ""
    picresults.Print "  Williams hit a home run in his last at-bat at Boston's Fenway Park,"
    picresults.Print "  and finished his career with 521 homers and 2654 hits. "
    picresults.Print "  He was inducted into baseball's Hall of Fame in 1966. "
    picresults.Print ""
    picresults.Print "  His 1969 autobiography was titled - My Turn At Bat -, "
    picresults.Print "  and his 1971 book -The Science of Hitting - remains a popular baseball manual."
    picresults.Print ""
    picresults.Print "  Late in life he suffered from congestive heart trouble and a series of strokes."
    picresults.Print "  He died at age 83 in July of 2002."
    
    
End Sub

Private Sub cmdquickfacts_Click()
    picresults.Cls
    picresults.Print ""
    picresults.Print ""
    picresults.Print "  QUICK FACTS"
    picresults.Print "  *******************************************************************************************"
    picresults.Print "  Williams wore uniform #9"
    picresults.Print ""
    picresults.Print "  His nicknames include The Kid, Teddy Ballgame, and The Splendid Splinter."
    picresults.Print ""
    picresults.Print "  Williams hit .400 in 6 games in 1952, and .407 in 37 games in 1953,"
    picresults.Print "  but both seasons featured too few at-bats to be considered official"
    picresults.Print ""
    picresults.Print "  Williams managed the Washington Senators (later the Texas Rangers)"
    picresults.Print "  from 1969-72..."
    picresults.Print ""
    picresults.Print "  In Korea Williams flew as wing man for future astronaut John Glenn."
End Sub

Private Sub cmdTedBioBack_Click()
    frmTedmenu.Show
    frmTedBio.Hide
End Sub

