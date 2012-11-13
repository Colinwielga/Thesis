VERSION 5.00
Begin VB.Form Russia 
   Appearance      =   0  'Flat
   BackColor       =   &H80000006&
   Caption         =   "Russia"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13725
   LinkTopic       =   "Form1"
   Picture         =   "Russia.frx":0000
   ScaleHeight     =   9405
   ScaleWidth      =   13725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00E0E0E0&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton cmdFamousRussian 
      BackColor       =   &H00E0E0E0&
      Caption         =   "FAMOUS RUSSIANS"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdGeography 
      BackColor       =   &H00E0E0E0&
      Caption         =   "GEOGRAPHY"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton cmdCulture 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CULTURE"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Image picImage 
      Height          =   5295
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "Russia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: The Globe Trotter Experience
'Form name: Russia.frm
'Author: Marta Gago & Brian Downes
'Date Written: Thursday March 27th, 2008
'Objective of form:  is to give the user some information on Russia with pictures of famous russians,
'pictures and information on geography, and Culture

'The Russia Form Disappears whereas the Asia Form Appears
Private Sub cmdBack_Click()
Russia.Hide
Asia.Show
picImage = Nothing  'The picImage is cleared of any picture
End Sub
'Gives pictures and messages of what pictures are shown, and other cultural items and landmarks
Private Sub cmdCulture_Click()

picImage = Nothing  'the picImage box is cleared

Russia.Picture = LoadPicture(App.Path & "\moscow3.jpg") 'The background becomes the picture loaded from this file
MsgBox ("The Cathedral of Intercession of the Virgin on the Moat or simply Pokrovskiy Cathedral, better known as the Cathedral of Saint Basil the Blessed , Saint Basil's Cathedral , or The Cathedral of the Protection of the Mother of God - is a multi-tented church on the Red Square in Moscow that also features distinctive onion domes. The cathedral is traditionally perceived as symbolic of the unique position of Russia between Europe and Asia.")
    'Message box is accompanied with the picture to explain what the picture is
Russia.Picture = LoadPicture(App.Path & "\Matroshka.jpg")
MsgBox ("A matryoshka doll or a Russian nested doll /also called a stacking doll or Babooshka doll/ is a set of dolls of decreasing sizes placed one inside another. Matryoshka is a derivative of the Russian female first name Matryona, which is traditionally associated with a fat, robust, rustic Russian woman.")

Russia.Picture = LoadPicture(App.Path & "\rf.jpg")
MsgBox ("Russia has a rich culinary history and offers a wide variety of soups, dishes made from fish, cereal based products and drinks. In addition to meat culinary, vegetables, fruit, mushrooms, berries and herbs also play a major part in the Russian diet. Primordial Russian products such as caviar, smetana (sour cream), buckwheat, rye flour, etc. have had a great influence on world-wide cuisine. Also known for their making of the beer.")


End Sub
'This sub command shows pictures of famous russians and is accompanied by their stories
Private Sub cmdFamousRussian_Click()

Russia.Picture = Nothing    'Clears the background of the form


picImage.Picture = LoadPicture(App.Path & "\brez.jpg")  'Loads the picture into the picImage box
MsgBox ("Brezhnev Leonid (1906-1982) Soviet politician. Started the politician career in the Ukraine, then Moldavia until he toppled Khruschev in Moscow in 1964. The grim conservatism of his rule was best exemplified by crushing of the “Prague Spring” in 1968. He proclaimed the right of Soviet intervention in any client state where Communism was threatened. Brezhnev was for conservative tendencies, no positive reforms during his 18-year ruling period.")
MsgBox ("He has also been a bearer of thousands of jokes. Here is one of them: At Lenin it was like in a tunnel: it is dark around and the light is ahead. At Stalin it was like in a bus: one is driving, half is seating, half is trembling with fear. At Khruschev it was like in a circus – one is speaking, everybody else is laughing. At Brezhnev – like in a movie – everyone is waiting till the end.")
        'Two message boxes display information on the picture
picImage.Picture = LoadPicture(App.Path & "\gagarin_yuri.jpg")
MsgBox ("Gagarin Yuri (1934-1968) Soviet Cosmonaut. The first man in space, Gagarin was rocketed into orbit on April 12, 1961, aboard the Vostok I spacecraft. His famous phrase at the very start “Poehali” (Let’s go) will always be a motto for world pioneers. Unable to steer the spacecraft, he orbited the earth once and after 108 minutes his craft parachuted safely down.")

picImage.Picture = LoadPicture(App.Path & "\gorbachev_mikhail.jpg")
MsgBox ("Gorbachev Mikhail (1931-)Soviet and Russian statesman. Gorbachev was the youngest leader since Stalin succeeded Lenin. He was a remarkable and forceful leader who changed the course of Russian history. General Secretary of the Communist Party from 1985, he embarked on a radical program of reform based on two premises: perestroika (restructuring) and glasnost (openness). The Soviet people achieved greater freedom of expression than they had enjoyed for over 50 years, but perestroika introduced dramatic socio-economic changes which only gradually revealed their benefits. Gorbachev was a good diplomat and a bad economist – the economy of the Soviet Union nearly collapsed by 1990 and facing growing criticism in public, he was overthrown in 1991 in a hardline coup. We would guess that the reason is in his remaining Soviet mentality. Gorbachev resigned all his offices in December 1991, announcing the official dissolution of the Soviet Union into independent states.")

picImage.Picture = LoadPicture(App.Path & "\lenin.jpg")
MsgBox ("Lenin Vladimir (1870-1924)Russian statesman. The revolutionary and founder of the Bolsheviks party, Lenin was upheld for over 65 years as the founder of the Soviet Union. Having studied Marxism at the University of St.-Petersburg, his involvement with revolutionary politics earned him three years exile in Siberia from 1897. He moved to Switzerland in 1900, becoming leader of the Bolsheviks during the abortive revolution of 1905. After the deposition of the Tsar, Lenin returned to Russia with German connivance in March 1917 in a “sealed train” and won power in the October Revolution that year. He had a murderous assault in 1918, which led to a long-time recovery. Unsatisfied with the idea of “Military communism” he instituted the New Economic Policy (NEP) in 1921 which permitted limited free enterprise. After his death Stalin abolished NEP and redirected Lenin’s ideas into the more severe and cruel interpretation.")

picImage.Picture = LoadPicture(App.Path & "\nicholas.jpg")
MsgBox ("Nicholas II (1868-1918) Tsar. Nicholas succeeded to the Russian imperial throne in 1894 on the death of his father, Alexander III. He inherited a vast, unruly empire, riven by political and social discontents, which required a ruthless autocratic ruler to control it. Resentment boiled over into revolution of 1905, during the disastrous Russian-Japanese war, and although Nicholas was quite prepared to order his troops to suppress the uprising, he also accepted the creation of an elected Duma /parliament/. Unfortunately he refused to allow it any power to introduce reforms, further alienating his people. In March, faced with implacable and almost universal opposition, Nicholas abdicated; in July 1918 he and his entire family were executed by the Bolsheviks, at Ekaterinburg.")

picImage.Picture = LoadPicture(App.Path & "\stalin.jpg")
MsgBox ("Stalin Josef (1879-1953)Statesman. The failed priest and bank robber who made the Soviet Union a superpower, from 1903, Stalin was successful as a propagandist for Bolshevism in his native Caucasus, and in raising funds at gunpoint. Lenin dubbed him “the wonderful Georgian”, and coopted him on to the party’s Central Committee. As a political Commissar he helped his future chief of the armed forces, Voroshilov, defend Tsaritsyn (later Stalingrad, now Volgograd) against the Whites. In 1922 Lenin appointed him Secretary of the Central Committee of the Party, the key post he held for the next 30 years. Lenin soon regretted the promotion and in his pre-deathbed “Testament” specifically warned the other old Bolsheviks against him. The Yalta peace conference (April 1945) confirmed his conquests, which were held by extending his secret police terror and slave-labor system. Paranoia affected his judgement; he miscalculated over the Korean war and the Berlin Airlift and died in 1953.")

End Sub
'gives the user a picture of the geography of Russia and a message box with information on Russian Geography
Private Sub cmdGeography_Click()

picImage = Nothing      'clears the picImage box

Russia.Picture = LoadPicture(App.Path & "\russiam.jpg") 'The background of the form displays the picture
MsgBox ("Russia is a country located in Europe and in North Asia. The European part of the country includes the territories to the west of the Ural Mountains. Russia is the largest country in the world in terms of area, but is unfavorably located in relation to major sea lanes of the world. Despite its size, much of the country lacks proper soils and climates (either too cold or too dry) for agriculture. Russia's topology includes Europe's highest mountain, its longest river, and the world's deepest lake. The topography and climate, however, resemble those of the northernmost portion of the North American continent. The northern forests and the plains bordering them to the south find their closest counterparts in the Yukon Territory and in the wide swath of land extending across most of Canada. The terrain, climate, and settlement patterns of Siberia are similar to those of Alaska and Canada.")
'Message box displays the information on the Russian Geography
End Sub


