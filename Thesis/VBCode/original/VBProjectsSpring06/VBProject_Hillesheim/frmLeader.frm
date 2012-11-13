VERSION 5.00
Begin VB.Form frmLeader 
   BackColor       =   &H80000001&
   Caption         =   "Famous Commanders"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Return to Main Page"
      Height          =   855
      Left            =   6720
      TabIndex        =   2
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   8520
      TabIndex        =   0
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Image imgSpruance 
      Height          =   3135
      Left            =   5400
      Picture         =   "frmLeader.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2415
   End
   Begin VB.Image imgTanaka 
      Height          =   3255
      Left            =   8040
      Picture         =   "frmLeader.frx":74CD
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Image imgOzawa 
      Height          =   3255
      Left            =   5400
      Picture         =   "frmLeader.frx":137B3
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Image imgNagumo 
      Height          =   3255
      Left            =   2760
      Picture         =   "frmLeader.frx":18606
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Image imgYamamoto 
      Height          =   3255
      Left            =   240
      Picture         =   "frmLeader.frx":1DEC5
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Image imgTurner 
      Height          =   3135
      Left            =   8040
      Picture         =   "frmLeader.frx":1F40B
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2415
   End
   Begin VB.Image imgHalsey 
      Height          =   3105
      Left            =   2760
      Picture         =   "frmLeader.frx":2139A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2460
   End
   Begin VB.Image imgNimitz 
      Height          =   3135
      Left            =   240
      Picture         =   "frmLeader.frx":22A2A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblAdm 
      BackColor       =   &H80000001&
      Caption         =   "DOUBLE CLICK ON DESIRED LEADER FOR A BRIEF BIOGRAPHY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000001&
      Caption         =   "By Jacob Hillesheim"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   8160
      Width           =   2535
   End
End
Attribute VB_Name = "frmLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Naval History (Naval.vpb)
'Leaders (frmLeader.frm)
'Jacob Hillesheim
'March 20,2006
'The purpose of this form is to teach the users about famous leaders of the War in the Pacific

Private Sub cmdBack_Click()
    'Returns user to main page
    frmMain.Show
    frmLeader.Hide
    
    'Clears welcome box and greets user
    frmMain.picWelcome.Cls
    frmMain.picWelcome.Print "Welcome, Admiral "; Left(x, 1); ". " & Left(y, 1); ". " & z
End Sub
Private Sub cmdQuit_Click()
    'Ends program
    End
End Sub
Private Sub ImgHalsey_DblClick()
    'displays Halsey's bio
    MsgBox "William Halsey was known as an extremely aggressive admiral who would take the fight to the Japanese. This is precisely why he was placed in charge of South Pacific operations. His arrival instantly revived the flagging morale of US soldiers and sailors in the region. Halsey's aggression got him into trouble, however. He charged after a decoy carrier force in the Battle if Leyte Gulf, leaving the transports open to attack. He was finally quietly relieved of active command after he took heavy casualties as he sailed his fleet into a typhoon for the second time.", , "Admiral William Halsey"
End Sub
Private Sub imgNagumo_DblClick()
    'Displays Nagumo's bio
    MsgBox "Chuichi Nagumo was Yamamoto's immediate subordinate and led the Japanese Carrier Force in the attack on Pearl Harbor and the Indian Ocean Raid. His fame took a turn for a worse and his generally remembered as an indecisive commander following defeats at Midway, Eastern Solomons, and the Santa Cruz Islands.", , "Admiral Chuichi Nagumo"
End Sub
Private Sub imgNimitz_DblClick()
    'Displays Nimitz's bio
    MsgBox "After the Japanese attack on Pearl Harbor, Chester Nimitz was promoted to Commander in Chief of the US Pacific Fleet. He was a brilliant strategist and an excellent judge of character. He knew the right man for every job. Nimitz was the architect of the US drive through the Central Pacific to Iwo Jima, Okinawa, and V-J Day.", , "Admiral Chester Nimitz"
End Sub
Private Sub imgOzawa_DblClick()
    'Displays Ozawa's bio
    MsgBox "Admiral Jisaburo Ozawa was an excellent Carrier commander who arrived on the scene too late. By the time he ascended to Commander in Chief of the Japanese Fleet, there was not much of an effective fleet left to take into battle. His inexperienced pilots were massacred in the Battle of the Philippine Sea, known to Americans as the Marianas Turkey Shoot. He successfully decoyed Admiral Halsey into chasing him in the Battle of Leyte Gulf, which would have been remembered as a great Japanese victory if it were not for extraordinary courage of American sailors and pilots in Battle of the San Bernardino Strait and the incompetence of his subordinates.", , "Admiral Jisaburo Ozawa"
End Sub
Private Sub imgSpruance_DblClick()
    'Displays Spruance's bio
    MsgBox "As the US Navy prepared for the Japanese attack on Midway, the expected commander of US forces, William Halsey, became ill. Nimitz assigned an obscure cruiser division commander named Raymond Spruance to command of the US task force. Spruance's confidence and intellect won the Battle of Midway, seen as the turning point of the War in the Pacific. He went on to smash the Japanese Fleet in the Battle of the Philippine Sea and command the US Fleet that captured the Gilbert, Marshall, and Mariana Islands, as well as Iwo Jima and Okinawa.", , "Admiral Raymond Spruance"
End Sub
Private Sub imgTanaka_DblClick()
    'Displays Tanaka's bio
    MsgBox "Raizo Tanaka was perhaps the most brilliant destroyer division commander of the war. He set up what became known as the Tokyo Express, midnight runs through Ironbottom Sound to resupply Japanese troops on Guadalcanal. He won a decisive victory at Tassafaronga, in which he sunk one and damaged three US cruisers without losing any of his own ships. Samuel Eliot Morrison dubbed him Tenacious Tanaka because he would press on and accomplish his mission no matter how hopeless it seemed.", , "Admiral Raizo Tanaka"
End Sub
Private Sub ImgTurner_DblClick()
    'Displays Turner's bio
    MsgBox "Richmond Kelly Turner became the world's foremost expert on amphibious assaults. He commanded amphibious landings in the South, Central, and West Pacific, including Guadalcanal, Makin, Guam, Saipan, Iwo Jima, and Okinawa. He was planning the assault on the Japanese Home Islands when the Japanese surrendered.", , "Admiral Richmond Kelly Turner"
End Sub
Private Sub imgYamamoto_DblClick()
    'Displays Yamamoto's bio
    MsgBox "Isoroku Yamamoto was the Commander in Chief of the Japanese Navy. Although he was opposed to war with the United States, he accepted responsibility for planning the attack on Pearl Harbor, which was successfully carried out. He planned Japanese attacks on Allied fleets and bases as the Japanese Fleet ran roughshod over the Allies in the Pacific for six months after Pearl Harbor. He was finally defeated in the decisive Battle of Midway in 1942. In 1943, Yamamoto proposed to visit the front lines in order to boost morale. US forces decrypted this message and assassinated Yamamoto by sending American fighters to intercept his transport plane.", , "Admiral Isoroku Yamamoto"
End Sub
