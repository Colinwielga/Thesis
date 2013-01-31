VERSION 5.00
Begin VB.Form frmGame
   AutoRedraw      =   -1  'True
   Caption         =   "Vegtable Maze"
   ClientHeight    =   8160
   ClientLeft      =   2955
   ClientTop       =   2610
   ClientWidth     =   8670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "game.frx":0000
   ScaleHeight     =   565.427
   ScaleMode       =   0  'User
   ScaleWidth      =   554.504
   Begin VB.PictureBox picSprite
      AutoRedraw      =   -1  'True
      Height          =   975
      Left            =   8280
      Picture         =   "game.frx":17BB42
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   773
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   11655
   End
   Begin VB.PictureBox picMask
      AutoRedraw      =   -1  'True
      Height          =   975
      Left            =   8280
      Picture         =   "game.frx":28BD44
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   773
      TabIndex        =   0
      Top             =   7200
      Visible         =   0   'False
      Width           =   11655
   End
   Begin VB.Image imgeggplant
      Height          =   570
      Index           =   21
      Left            =   5400
      Picture         =   "game.frx":39BF46
      Top             =   4440
      Width           =   555
   End
   Begin VB.Image imgeggplant
      Height          =   570
      Index           =   20
      Left            =   0
      Picture         =   "game.frx":39D028
      Top             =   -120
      Width           =   555
   End
   Begin VB.Image imgeggplant
      Height          =   570
      Index           =   19
      Left            =   7080
      Picture         =   "game.frx":39E10A
      Top             =   1440
      Width           =   555
   End
   Begin VB.Image imgeggplant
      Height          =   570
      Index           =   18
      Left            =   4920
      Picture         =   "game.frx":39F1EC
      Top             =   960
      Width           =   555
   End
   Begin VB.Image imgeggplant
      Height          =   570
      Index           =   17
      Left            =   3480
      Picture         =   "game.frx":3A02CE
      Top             =   4560
      Width           =   555
   End
   Begin VB.Image imgeggplant
      Height          =   570
      Index           =   16
      Left            =   5040
      Picture         =   "game.frx":3A13B0
      Top             =   2280
      Width           =   555
   End
   Begin VB.Image imgeggplant
      Height          =   570
      Index           =   15
      Left            =   2760
      Picture         =   "game.frx":3A2492
      Top             =   960
      Width           =   555
   End
   Begin VB.Image imgbroccoli
      Height          =   585
      Index           =   21
      Left            =   7080
      Picture         =   "game.frx":3A3574
      Top             =   5040
      Width           =   540
   End
   Begin VB.Image imgbroccoli
      Height          =   585
      Index           =   20
      Left            =   5400
      Picture         =   "game.frx":3A462A
      Top             =   5040
      Width           =   540
   End
   Begin VB.Image imgbroccoli
      Height          =   585
      Index           =   19
      Left            =   6600
      Picture         =   "game.frx":3A56E0
      Top             =   3600
      Width           =   540
   End
   Begin VB.Image imgbroccoli
      Height          =   585
      Index           =   18
      Left            =   7080
      Picture         =   "game.frx":3A6796
      Top             =   2040
      Width           =   540
   End
   Begin VB.Image imgbroccoli
      Height          =   585
      Index           =   17
      Left            =   3360
      Picture         =   "game.frx":3A784C
      Top             =   1200
      Width           =   540
   End
   Begin VB.Image imgbroccoli
      Height          =   585
      Index           =   16
      Left            =   1560
      Picture         =   "game.frx":3A8902
      Top             =   840
      Width           =   540
   End
   Begin VB.Image imgbroccoli
      Height          =   585
      Index           =   15
      Left            =   2040
      Picture         =   "game.frx":3A99B8
      Top             =   1920
      Width           =   540
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   22
      Left            =   7560
      Picture         =   "game.frx":3AAA6E
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   540
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   21
      Left            =   5400
      Picture         =   "game.frx":3ABA4C
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   540
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   20
      Left            =   6000
      Picture         =   "game.frx":3ACA2A
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   540
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   19
      Left            =   7080
      Picture         =   "game.frx":3ADA08
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   540
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   18
      Left            =   6600
      Picture         =   "game.frx":3AE9E6
      Stretch         =   -1  'True
      Top             =   840
      Width           =   540
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   17
      Left            =   3840
      Picture         =   "game.frx":3AF9C4
      Stretch         =   -1  'True
      Top             =   960
      Width           =   540
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   16
      Left            =   4920
      Picture         =   "game.frx":3B09A2
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   540
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   15
      Left            =   3600
      Picture         =   "game.frx":3B1980
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   540
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   14
      Left            =   1080
      Picture         =   "game.frx":3B295E
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   540
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   13
      Left            =   2040
      Picture         =   "game.frx":3B393C
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   540
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   22
      Left            =   5400
      Picture         =   "game.frx":3B491A
      Top             =   6240
      Width           =   480
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   21
      Left            =   5520
      Picture         =   "game.frx":3B573C
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   20
      Left            =   7080
      Picture         =   "game.frx":3B655E
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   19
      Left            =   7080
      Picture         =   "game.frx":3B7380
      Top             =   840
      Width           =   480
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   18
      Left            =   4440
      Picture         =   "game.frx":3B81A2
      Top             =   960
      Width           =   480
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   17
      Left            =   3600
      Picture         =   "game.frx":3B8FC4
      Top             =   3480
      Width           =   480
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   16
      Left            =   2640
      Picture         =   "game.frx":3B9DE6
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   12
      Left            =   1080
      Picture         =   "game.frx":3BAC08
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   11
      Left            =   5640
      Picture         =   "game.frx":3BBA2A
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   10
      Left            =   8160
      Picture         =   "game.frx":3BC84C
      Top             =   5040
      Width           =   480
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   9
      Left            =   120
      Picture         =   "game.frx":3BD66E
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   12
      Left            =   2040
      Picture         =   "game.frx":3BE490
      Stretch         =   -1  'True
      Top             =   960
      Width           =   540
   End
   Begin VB.Image imgbroccoli
      Height          =   585
      Index           =   9
      Left            =   600
      Picture         =   "game.frx":3BF46E
      Top             =   2160
      Width           =   540
   End
   Begin VB.Image imgcmdQuit
      Height          =   585
      Left            =   5520
      Picture         =   "game.frx":3C0524
      Top             =   7560
      Width           =   1260
   End
   Begin VB.Image imgcmdMain
      Height          =   555
      Left            =   1200
      Picture         =   "game.frx":3C2BCA
      Top             =   7560
      Width           =   2730
   End
   Begin VB.Image imgeggplant
      Height          =   570
      Index           =   5
      Left            =   6840
      Picture         =   "game.frx":3C7B40
      Top             =   7080
      Width           =   555
   End
   Begin VB.Image imgeggplant
      Height          =   570
      Index           =   4
      Left            =   0
      Picture         =   "game.frx":3C8C22
      Top             =   7080
      Width           =   555
   End
   Begin VB.Image imgeggplant
      Height          =   570
      Index           =   3
      Left            =   960
      Picture         =   "game.frx":3C9D04
      Top             =   4200
      Width           =   555
   End
   Begin VB.Image imgeggplant
      Height          =   570
      Index           =   2
      Left            =   1800
      Picture         =   "game.frx":3CADE6
      Top             =   5280
      Width           =   555
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   5
      Left            =   3480
      Picture         =   "game.frx":3CBEC8
      Top             =   5160
      Width           =   480
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   4
      Left            =   1080
      Picture         =   "game.frx":3CCCEA
      Top             =   5280
      Width           =   480
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   3
      Left            =   7440
      Picture         =   "game.frx":3CDB0C
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   2
      Left            =   5160
      Picture         =   "game.frx":3CE92E
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   5
      Left            =   1800
      Picture         =   "game.frx":3CF750
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   540
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   4
      Left            =   3360
      Picture         =   "game.frx":3D072E
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   540
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   3
      Left            =   3120
      Picture         =   "game.frx":3D170C
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   540
   End
   Begin VB.Image imgbroccoli
      Height          =   585
      Index           =   5
      Left            =   1560
      Picture         =   "game.frx":3D26EA
      Top             =   4200
      Width           =   540
   End
   Begin VB.Image imgbroccoli
      Height          =   585
      Index           =   4
      Left            =   5640
      Picture         =   "game.frx":3D37A0
      Top             =   7080
      Width           =   540
   End
   Begin VB.Image imgbroccoli
      Height          =   585
      Index           =   3
      Left            =   3360
      Picture         =   "game.frx":3D4856
      Top             =   6240
      Width           =   540
   End
   Begin VB.Image imgbroccoli
      Height          =   585
      Index           =   2
      Left            =   8040
      Picture         =   "game.frx":3D590C
      Top             =   6960
      Width           =   540
   End
   Begin VB.Image imgeggplant
      Height          =   570
      Index           =   1
      Left            =   2280
      Picture         =   "game.frx":3D69C2
      Top             =   7080
      Width           =   555
   End
   Begin VB.Image imgbroccoli
      Height          =   585
      Index           =   1
      Left            =   1080
      Picture         =   "game.frx":3D7AA4
      Top             =   7080
      Width           =   540
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   1
      Left            =   2880
      Picture         =   "game.frx":3D8B5A
      Top             =   6960
      Width           =   480
   End
   Begin VB.Image imgwatermelon
      Height          =   615
      Left            =   4320
      Picture         =   "game.frx":3D997C
      Top             =   6120
      Width           =   600
   End
   Begin VB.Image imgpepper
      Height          =   555
      Index           =   0
      Left            =   600
      Picture         =   "game.frx":3DACF6
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image imgfruitbasket
      Height          =   450
      Left            =   120
      Picture         =   "game.frx":3DBB18
      Top             =   1440
      Width           =   540
   End
   Begin VB.Image Image5
      Height          =   450
      Left            =   1200
      Picture         =   "game.frx":3DC802
      Top             =   4800
      Width           =   525
   End
   Begin VB.Image imgeggplant
      Height          =   570
      Index           =   0
      Left            =   4560
      Picture         =   "game.frx":3DD4EC
      Top             =   6960
      Width           =   555
   End
   Begin VB.Image imgbroccoli
      Height          =   585
      Index           =   0
      Left            =   3360
      Picture         =   "game.frx":3DE5CE
      Top             =   7080
      Width           =   540
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   2
      Left            =   3960
      Picture         =   "game.frx":3DF684
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   540
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   1
      Left            =   6240
      Picture         =   "game.frx":3E0662
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   540
   End
   Begin VB.Image Image2
      Height          =   615
      Left            =   720
      Top             =   3840
      Width           =   735
   End
   Begin VB.Image imgbanana
      Height          =   570
      Left            =   8040
      Picture         =   "game.frx":3E1640
      Top             =   5880
      Width           =   525
   End
   Begin VB.Image Imgtomatoes
      Height          =   555
      Index           =   0
      Left            =   1680
      Picture         =   "game.frx":3E268A
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   540
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Code is based on a tutorial from: https://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=6077&lngWId=1
'Sprite images from www.spriters-resource.com/
Option Explicit
Dim Alt As Integer          'Allows two alturnating images to be displayed making the dinosaur look like it's flying.
Dim XPos As Integer, YPos As Integer    'these are used as the location of the dinosaur
Dim LeftX As Integer, TopY As Integer   'this shows where the frame of the image is.  these are the variables I adjust to display each of the dinosaurs images from the .bmp
Dim tomatoup As Integer, tomatodown As Integer, tomatoleft As Integer, tomatoright As Integer ' tomato boundries
Dim pepperup As Integer, pepperdown As Integer, pepperleft As Integer, pepperright As Integer 'pepper boundries
Const Speed = 20   'this is how far the dinosaur moves each time you push a direction button.
Dim spriteup As Long, spritedown As Long, spriteleft As Long, spriteright As Long 'declares the boundries of the sprite
Dim eggplantup As Long, eggplantdown As Long, eggplantleft As Long, eggplantright As Long ' declare eggplant boundries
Dim broccoliup As Integer, brocollidown As Integer, brocollileft As Integer, brocolliright As Integer 'broccoli boundries
Dim watermelonup As Integer, watermelondown As Integer, watermelonleft As Integer, watermelonright As Integer ' watermelon boundries
Dim fruitbasketup As Integer, fruitbasketdown As Integer, fruitbasketleft As Integer, fruitbasketright As Integer 'fruitbasket boundries
Dim bananaup As Integer, bananadown As Integer, bananaleft As Integer, bananaright As Integer 'banana boundries
Dim sprite As Object
Dim upcross As Boolean, downcross As Boolean, leftcross As Boolean, rightcross As Boolean 'these are used to determine if one of the boundries has been crossed by the sprite.





Private Sub cmdQuit_Click()
    End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)  'allows for keyboard input in the game
If KeyCode = vbKeyLeft Then    'this if statement moves the sprite left while alternating between pictures to make the dinosaur appear to be flapping it's wings.
    If Alt = 0 Then       'this is for the first time the left button is pressed.
        XPos = XPos - Speed 'moves the dinosaur by the constant speed
        Alt = 1           'this alturnates between the two if statements to display two different images each time you push the left button.
        LeftX = 52          'this designates the area of the picture box to be used as each image
        TopY = 0            'so does this
        MoveIt              'tells the dinosaur to move by the indicated speed.
    ElseIf Alt = 1 Then     'this If is for the other image
        Alt = 0
        TopY = 0
        LeftX = 110
        XPos = XPos - Speed
        MoveIt
    End If
ElseIf KeyCode = vbKeyUp Then    'moves the dinosaur up one space.
        If Alt = 1 Then
        TopY = 0
        Alt = 0
        LeftX = 360
        YPos = YPos - Speed
        MoveIt
    ElseIf Alt = 0 Then
        LeftX = 420
        Alt = 1
        TopY = 0
        YPos = YPos - Speed
        MoveIt
    End If
ElseIf KeyCode = vbKeyRight Then    'this is the same as the left button except that it moves the dinosaur to the right.
    If Alt = 1 Then
        Alt = 0
        TopY = 0
        LeftX = 164
        XPos = XPos + Speed
        MoveIt
    ElseIf Alt = 0 Then
        LeftX = 254
        Alt = 1
        TopY = 0
        XPos = XPos + Speed
        MoveIt
    End If
ElseIf KeyCode = vbKeyDown Then    'down one space
    If Alt = 0 Then
        Alt = 1
        LeftX = 478
        TopY = 0
        YPos = YPos + Speed
        MoveIt
    ElseIf Alt = 1 Then
        LeftX = 540
        TopY = 0
        Alt = 0
        YPos = YPos + Speed
        MoveIt
    End If
End If

End Sub


'this refreshes the images so that the dinosaur doesn't appear as a long blur as it moves across the field
Sub MoveIt()
'frmGame.cls erases the older image so that it isn't just a streak of color.
frmGame.Cls
'Call BitBlt gets the sprite and mask pictures and sets their size at 64,64.  srcand and srcinvert remove the color from the background of this image so it is in the shape of a dinosaur and not a square.
Call BitBlt(frmGame.hDC, XPos, YPos, 64, 64, picMask.hDC, LeftX, TopY, SRCAND)
Call BitBlt(frmGame.hDC, XPos, YPos, 64, 64, picSprite.hDC, LeftX, TopY, SRCINVERT)

'this refreshes the images so that you can see them all of the time.
frmGame.Refresh
End Sub

Private Sub Form_Load()
    XPos = 0    'starts the dinosaur at the left side of the form
    YPos = 300  'starts the dinosaur at the bottem of the form
    LeftX = 52  'This describes the position of the first image the form uses when starting the game
    TopY = 0    'This does too but for the Y axis
    MoveIt      'This displays the first image
End Sub

'goes to main menu
Private Sub imgcmdMain_Click()
    frmMain.Show
End Sub

'quits program
Private Sub imgcmdQuit_Click()
    End
End Sub
