VERSION 5.00
Begin VB.Form frmfriend 
   Caption         =   "Kitty's friends"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   LinkTopic       =   "Form2"
   Picture         =   "frmfriend.frx":0000
   ScaleHeight     =   7935
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPochacco 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pochacco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10320
      Picture         =   "frmfriend.frx":4B80B
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdPurin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9000
      Picture         =   "frmfriend.frx":4BD80
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdMonkichi 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Monkichi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7560
      Picture         =   "frmfriend.frx":4C2D6
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdLittleTwinStars 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Little Twin Stars"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      Picture         =   "frmfriend.frx":4C94F
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdchococat 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chococat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      Picture         =   "frmfriend.frx":4D099
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdPandapple 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pandapple"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10320
      Picture         =   "frmfriend.frx":4D413
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdCKitty 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Charmmy Kitty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9000
      Picture         =   "frmfriend.frx":4DA1B
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdMinnaNoTabo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Minna No Tabo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6120
      Picture         =   "frmfriend.frx":4E0CF
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdPekkle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pekkle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7560
      Picture         =   "frmfriend.frx":4E662
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdChiChaiMonchan 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chi Chai Monchan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      Picture         =   "frmfriend.frx":4EC2A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdKeroppi 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Keroppi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3240
      Picture         =   "frmfriend.frx":4F1C9
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdChibimaru 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chibimaru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Picture         =   "frmfriend.frx":4F7C5
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmddeerylou 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Deerylou"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6120
      Picture         =   "frmfriend.frx":4FE14
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdbadtz 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Badtz-Maru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Picture         =   "frmfriend.frx":5032D
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdkitty 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hello Kitty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10320
      Picture         =   "frmfriend.frx":50676
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdMelody 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Melody"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1800
      Picture         =   "frmfriend.frx":50BB1
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.PictureBox picpicture 
      BackColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   480
      Picture         =   "frmfriend.frx":510A6
      ScaleHeight     =   3195
      ScaleWidth      =   3435
      TabIndex        =   3
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton cmdcinnamoroll 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cinnamoroll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3240
      Picture         =   "frmfriend.frx":524CC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      Picture         =   "frmfriend.frx":528D1
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to the main page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lblintro 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   4800
      TabIndex        =   18
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmfriend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Picname As String


Private Sub cmdback_Click()
frmfriend.Visible = False
frmmain.Visible = True
End Sub

Private Sub cmdcinnamaru_Click()
frmdisplay.Visible = True
frmfriend.Visible = False

End Sub

Private Sub cmdbadtz_Click()
Picname = App.Path & "\friendpic\" & "batz1" & ".JPG"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "Badtz-Maru, literally 'XO', is one of the many fictional characters produced by Japanese corporation Sanrio. Unlike the more popular Hello Kitty, Badtz-Maru has an attitude and is one of the few characters that is marketed to both males and females."
End Sub

Private Sub cmdChibimaru_Click()
Picname = App.Path & "\friendpic\" & "Chibimaru1" & ".GIF"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "Chibimaru is an active dog to say the least! He lives in a house with a red roof with his favorite toys¡ªhis stuffed animals. His favorite treat is a milk-flavored cookie shaped like a bone. Mmm! "
End Sub

Private Sub cmdChiChaiMonchan_Click()
Picname = App.Path & "\friendpic\" & "ChiChaiMonchan1" & ".GIF"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "A fun-loving little boy monkey with a whirly-curly tail, Chi Chai Monchan lives on a small tropical island in the South Seas where he spends his days climbing trees and eating bananas."
End Sub

Private Sub cmdchococat_Click()
Picname = App.Path & "\friendpic\" & "chococat1" & ".JPG"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "Chococat is one of the many fictional characters produced by Japanese corporation Sanrio. He is drawn as a black cat with huge black eyes, four whiskers, and like counterpart Hello Kitty, no mouth. His name comes from his chocolate-coloured nose. Like other Sanrio characters, he appears on a variety of merchandise, including stationery, coffee mugs, plush toys, bath towels, etc."
End Sub


Private Sub cmdcinnamoroll_Click()
Picname = App.Path & "\friendpic\" & "cinnamoroll1" & ".JPG"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "One day, while the owner of CafeCinnamon was admiring the sky, a tiny white rabbit came floating by, looking just like a small, fluffy cloud. She thought, Maybe he caught a whiff of the cinnamon rolls and came to check them out. The curious rabbit took a shine to the cafeowner and her delicious cinnamon rolls, so he decided to stay. Since his tail was plump and curled up like a cinnamon roll, she decided to call him Cinnamoroll. Sweet, little Cinnamoroll was instantly popular with customers and soon became Café Cinnamon's official mascot. Now, when he is not napping on the café terrace, you may find Cinnamoroll flying around the town looking for fun and new adventures with his friends Chiffon, Mocha, Espresso, Cappuccino, and Milk. His birthday is March 6th."

End Sub

Private Sub cmdCKitty_Click()
Picname = App.Path & "\friendpic\" & "CharmmyKitty1" & ".JPG"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "Charmmy Kitty is a white persian cat that Papa gave to Hello Kitty as a gift. She is well-mannered, quiet, and listens to whatever Hello Kitty says. She loves objects that are bright and sparkly! Charmmy Kitty wears a lace-lined ribbon on her left ear, and a necklace which holds the key to Hello Kitty¡¯s jewelry box. "
End Sub

Private Sub cmddeerylou_Click()
Picname = App.Path & "\friendpic\" & "deerylou1" & ".JPG"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "Deerylou the cheerful Fawm lives happily with his friends in the rainbow forest. Deery-lou spends his time playing in the sun, chasing butterfly and making friends. When he's really happy he always swings his tail! His birhtday is Jan 8th."
End Sub

Private Sub cmdKeroppi_Click()
Picname = App.Path & "\friendpic\" & "Keroppi1" & ".GIF"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "Keroppi lives with his brother, sister and parents in a big house on the edge of Donut Pond, the largest and bluest pond around. Keroppi and his friends share his love for baseball and boomerangs. Most often he is seen with his little snail friend Den Den, always tagging along a little behind. "
End Sub

Private Sub cmdkitty_Click()
Picname = App.Path & "\friendpic\" & "kitty1" & ".JPG"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "Hello Kitty was born on November 1st, and she lives in London England with her parent and her twin sister, Mimmy. They have a lot of friends at school with whom they share many adventures. Her hobby includes travelling, reading, music, eating yummy cookies her sister Mimmy backes, and best of all making new friends. As Hello Kitty says, you can never have too many friends."
End Sub

Private Sub cmdLittleTwinStars_Click()
Picname = App.Path & "\friendpic\" & "LittleTwinStars1" & ".GIF"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "Kiki and Lala, the Little Twin Stars, were born on the Star of Compassion. With permission from Mother-Star and Father-Star they set out for a visit to Earth. Lala¡¯s star wand led them on their journey. Ever since they arrived, the Little Twin Stars have been spreading happiness to everyone they meet. "
End Sub

Private Sub cmdMelody_Click()
Picname = App.Path & "\friendpic\" & "melody1" & ".JPG"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "My Melody is one of the many fictional characters produced by the Japanese company Sanrio. The character looks like a rabbit and always wears a red or pink hood. It is especially popular in Asia and can be found on children's toys and merchandise."

End Sub

Private Sub cmdMinnaNoTabo_Click()
Picname = App.Path & "\friendpic\" & "MinnaNoTabo1" & ".GIF"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "Always bright and cheerful, Minna No Tabo doesn¡¯t have a dishonest bone in his body. His straightforward and can-do approach to life wins him many friends, although sometimes he can get into a bit of a panic."
End Sub

Private Sub cmdMonkichi_Click()
Picname = App.Path & "\friendpic\" & "Monkichi1" & ".GIF"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "Most monkeys like bananas, but only Monkichi can eat ten bananas in one minute! He lives high in the mountains with all his friends. Monkichi loves keeping everyone entertained with stories, jokes and poems. His dream is to one day be a poet or maybe a professional comedian."
End Sub

Private Sub cmdPandapple_Click()
Picname = App.Path & "\friendpic\" & "Pandapple1" & ".GIF"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "Pandapple is a boy panda who absolutely LOVES apples! He feels the most relaxed when he is seated in his apple chair in his apple-scented house. Pandapple is always playing with his pet caterpillar Imomushi. He is happy and cheerful, but when it comes to apples, he can be a little picky. "
End Sub

Private Sub cmdPekkle_Click()
Picname = App.Path & "\friendpic\" & "Pekkle1" & ".GIF"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "found singing or dancing, two things he excels at. He is currently enrolled in a tap dance class. "
End Sub

Private Sub cmdPochacco_Click()
Picname = App.Path & "\friendpic\" & "Pochacco1" & ".GIF"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "Pochacco is the most popular purebred in the neighborhood. This sports-minded pup is the best three-on-three basketball player on the playground and a not-so-shabby soccer goalie, too. He¡¯s a real original¡ªhow many vegetarian canines do you know? Pochacco loves carrots but banana ice cream is his all-time favorite! "
End Sub

Private Sub cmdPurin_Click()
Picname = App.Path & "\friendpic\" & "Purin1" & ".GIF"
picpicture.Picture = LoadPicture(Picname)
lblintro.Caption = "Purin is a good-natured Golden Retriever who dreams of growing up and becoming a big dog, just like his mother and father! He spends most of his time napping or going on walks with his best friend Muffin. He loves to drink milk and eat foods that are nice and soft, like pudding! "
End Sub

Private Sub cmdquit_Click()
End
End Sub
