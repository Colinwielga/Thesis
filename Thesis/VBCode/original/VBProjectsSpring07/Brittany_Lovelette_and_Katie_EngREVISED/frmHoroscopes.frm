VERSION 5.00
Begin VB.Form frmHoroscopes 
   BackColor       =   &H00C00000&
   Caption         =   "Horoscopes"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form2"
   Picture         =   "frmHoroscopes.frx":0000
   ScaleHeight     =   6945
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080C0FF&
      Caption         =   "Go to Main Menu"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   6015
      Left            =   4680
      ScaleHeight     =   5955
      ScaleWidth      =   4635
      TabIndex        =   2
      Top             =   240
      Width           =   4695
   End
   Begin VB.CommandButton cmdLuckydays 
      BackColor       =   &H0080C0FF&
      Caption         =   "What are your lucky days?"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdHoroscope 
      BackColor       =   &H0080C0FF&
      Caption         =   "What is your horoscope?"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   1815
   End
End
Attribute VB_Name = "frmHoroscopes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'spell check, error check, Cited: Lecture 11

'This button takes the user back to the form frmHome
'Cited: Lecture 18
Private Sub cmdBack_Click()
    frmHoroscopes.Hide
    'makes the form frmHoroscopes invisible to the user
    frmHome.Show
    'makes the form frmHome visible to the user
End Sub

'Cited (for this button and also for the Lucky Days button): http://www.helium.com/tm/101081/astrology-predicting-lucky, http://images.google.com/imgres?imgurl=, http://www.tattoosymbol.com/zodiac/capricorn.jpg&imgrefurl=http://www.tattoosymbol.com/zodiac/capricorn.html&h=148&w=144&sz=4&hl=en&start=19&um=1&tbnid=0T5NNHuBtQJRiM:&tbnh=95&tbnw=92&prev=/images%3Fq%3DCapricorn%26svnum%3D10%26um%3D1%26hl%3Den
'This button tells the user their horoscope and displays its sign based on the birthdate they enter into an input box
Private Sub cmdHoroscope_Click()
    Dim number As String, month As String, day As String, year As String
    Dim birthdate As Integer
    Dim Space As String, Pos As Integer
    
    
    'Declared variables
    'Cited: Lecture 11
    Space = " "
    Pos = 0
    number = InputBox("Enter Date of birth.  e.g. '01/15/1986'", "Enter DOB")
    month = Left(number, 2)
    day = Mid(number, 4, 2)
    year = Right(number, 4)
    birthdate = month + day
    'initialized variables
    'Cited: TA Chris Kerber
    Dim myString As String, stringLength As Integer, tempString As String, lineLength As Integer, i As Integer
    'Declared string variables
    'Cited: TA Chris Kerber
    lineLength = 50
    'initialized variables
    'Cited: TA Chris Kerber
    picResults.Cls
    'clears picture box
    'Cited: Lecture 11 and TA Chris Kerber
    
    'This entire case is Cited from TA Chris Kerber, he and Katie Eng worked together to learn how to format the picture box
    Select Case birthdate
        'The documentation in this first section applies to the entire case
        Case 101 To 119
            picResults.Picture = LoadPicture(App.Path & "\Newcapricorn.jpg")
            picResults.Print "Capricorn"
            'displays picture of sign and horoscope
            myString = "A tall tale you tell at the beginning of the month will land you into an unexpected pot of hot water when you accidentally embellish the details of a story. Although you were only trying to make it sound a little more interesting, you will inadvertently ruffle a few feathers! The New Moon will pass through your house of communications on the 19th, making you want to express your inner most thoughts through an artistic medium. Whether it's poetry, writing or even painting, putting pen to paper will be a healthy outlet and have a beneficial effect on your emotions. Around the 31st, a social gathering could leave you with more than you bargained for when your magnetic charm will find you more attention that you know what to do with. While good news for the single Capricorn, the difficult part will be short listing all the potential suitors!"
            stringLength = Len(myString)
            i = 1
            'this allows a lengthy paragraph to be displayed in the picture box
    'This While Wend statement formats the paragraph to fit in the picture box
    While i + lineLength < stringLength
            tempString = Mid(myString, i, lineLength)
            Pos = InStrRev(tempString, Space)
            tempString = Mid(myString, i, Pos)
            
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 120 To 131
            picResults.Picture = LoadPicture(App.Path & "\Newaquarius.jpg")
            picResults.Print "Aquarius"
            'displays picture of sign and horoscope
            myString = "When the Full Moon passes through your house of joint finances on the 3rd of the month you will want to take stock of your long term financial obligations. Reading the fine print on existing arrangements such as mortgage, insurance or loans could potentially save you money in the long run. Mid month, you will have a flare for expressing yourself as a smooth connection between Mercury, planet of communications and charming Venus will give you a refined way of getting your point across. The influence of this transit is especially favourable for social events or job interviews as you are easily able to analyse the human dynamic, telling people what they want to hear. Your charisma will be compelling around the 25th when forceful Mars teams up with ethereal Neptune to make you a powerful influence. Those around you will use your ambitious approach in getting what you want out of life as an example."
    stringLength = Len(myString)
    i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
    picResults.Print Right(myString, stringLength - i + 1)
        
        Case 201 To 218
            picResults.Picture = LoadPicture(App.Path & "\Newaquarius.jpg")
            picResults.Print "Aquarius"
            'displays picture of sign and horoscope
            myString = "When the Full Moon passes through your house of joint finances on the 3rd of the month you will want to take stock of your long term financial obligations. Reading the fine print on existing arrangements such as mortgage, insurance or loans could potentially save you money in the long run. Mid month, you will have a flare for expressing yourself as a smooth connection between Mercury, planet of communications and charming Venus will give you a refined way of getting your point across. The influence of this transit is especially favourable for social events or job interviews as you are easily able to analyse the human dynamic, telling people what they want to hear. Your charisma will be compelling around the 25th when forceful Mars teams up with ethereal Neptune to make you a powerful influence. Those around you will use your ambitious approach in getting what you want out of life as an example."
    stringLength = Len(myString)
    i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
    picResults.Print Right(myString, stringLength - i + 1)
            
        Case 219 To 228
            picResults.Picture = LoadPicture(App.Path & "\Newpisces.jpg")
            picResults.Print "Pisces"
            'displays picture of sign and horoscope
            myString = "Freedom of expression will be of importance to you at the beginning of the month, as you want to be seen for your individual style. Those who know you well may notice a radical change as you will want to spice things up and not be seen as the usual predictable you. When the New Moon comes to pass in your house of self-development on the 19th you will want to get in touch with yourself and nurture your own emotions. This is a good time for reflection or meditation as you will be able deal with any unresolved issues from the past. Towards the end of the month when a competitive career venture presents itself, you will whole-heartedly jump in with both feet. Although the workload may be more than you were bargaining for, you know in the long run this is the opportunity you've been waiting for."
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
    picResults.Print Right(myString, stringLength - i + 1)
        Case 301 To 320
            picResults.Picture = LoadPicture(App.Path & "\Newpisces.jpg")
            picResults.Print "Pisces"
            'displays picture of sign and horoscope
            myString = "Freedom of expression will be of importance to you at the beginning of the month, as you want to be seen for your individual style. Those who know you well may notice a radical change as you will want to spice things up and not be seen as the usual predictable you. When the New Moon comes to pass in your house of self-development on the 19th you will want to get in touch with yourself and nurture your own emotions. This is a good time for reflection or meditation as you will be able deal with any unresolved issues from the past. Towards the end of the month when a competitive career venture presents itself, you will whole-heartedly jump in with both feet. Although the workload may be more than you were bargaining for, you know in the long run this is the opportunity you've been waiting for."
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
    picResults.Print Right(myString, stringLength - i + 1)
        Case 321 To 331
            picResults.Picture = LoadPicture(App.Path & "\Newaries.jpg")
            picResults.Print "Aries"
            'displays picture of sign and horoscope
            myString = "When the Full Moon occurs on the 3rd of the month, your day-to-day life will take a turn to the practical side as you realise you could do with an overhaul. Whether it's setting up a system to keep the chequebook balanced or starting a new exercise programme, this fine-tuning will make a difference. A delightful passing of charming Venus mid month will put you in the right place at the right time when your persuasive energy convinces others that your high standards are needed. It may be tricky business, but your constructive expression will be the catalyst for positive change. When the Sun moves into your sign on the 21st you will get a real energy boost, as you will want to burst into action. This period of time will give you the vitality needed to kick start a personal project or enable you the me time you could really do with."
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 401 To 420
            picResults.Picture = LoadPicture(App.Path & "\Newaries.jpg")
            picResults.Print "Aries"
            'displays picture of sign and horoscope
            myString = "When the Full Moon occurs on the 3rd of the month, your day-to-day life will take a turn to the practical side as you realise you could do with an overhaul. Whether it's setting up a system to keep the chequebook balanced or starting a new exercise programme, this fine-tuning will make a difference. A delightful passing of charming Venus mid month will put you in the right place at the right time when your persuasive energy convinces others that your high standards are needed. It may be tricky business, but your constructive expression will be the catalyst for positive change. When the Sun moves into your sign on the 21st you will get a real energy boost, as you will want to burst into action. This period of time will give you the vitality needed to kick start a personal project or enable you the 'Me' time you could really do with."
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 421 To 430
            picResults.Picture = LoadPicture(App.Path & "\Newtaurus.jpg")
            picResults.Print "Taurus"
            'displays picture of sign and horoscope
            myString = "As Mercury, planet of communications goes direct on the 9th you will finally be able to say exactly what's on your mind at work. For the last three weeks you may have been tongue-tied or promising more than you can deliver but as things begin to run smoothly, you can carry on in your usual methodical way. Mid month your steady determination will allow you to take things one step at a time as you set out to reach your goal. Your practicality and optimism will help you focus on your target keeping you on the right track - but watch out for wanting too much too soon. Towards the 25th, many people will be drawn to your compelling presence as a potent connection between assertive Mars and Neptune the illusionist will give you sway at the office. People may seem to follow your lead, but be sure to keep your intentions on the level."
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 501 To 520
            picResults.Picture = LoadPicture(App.Path & "\Newtaurus.jpg")
            picResults.Print "Taurus"
            'displays picture of sign and horoscope
            myString = "As Mercury, planet of communications goes direct on the 9th you will finally be able to say exactly what's on your mind at work. For the last three weeks you may have been tongue-tied or promising more than you can deliver but as things begin to run smoothly, you can carry on in your usual methodical way. Mid month your steady determination will allow you to take things one step at a time as you set out to reach your goal. Your practicality and optimism will help you focus on your target keeping you on the right track - but watch out for wanting too much too soon. Towards the 25th, many people will be drawn to your compelling presence as a potent connection between assertive Mars and Neptune the illusionist will give you sway at the office. People may seem to follow your lead, but be sure to keep your intentions on the level."
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 521 To 531
            picResults.Picture = LoadPicture(App.Path & "\Newgemini.jpg")
            picResults.Print "Gemini"
            'displays picture of sign and horoscope
            myString = "A change in your domestic setting may be called for on the 3rd when the Full Moon encourages you to spruce things up. Getting rid of clutter and straightening out the cupboards will make your house feel like more of a home and a place where you can begin to relax! Mid month you will get down to the root of the matter within an intimate relationship, as new information will put a different perspective on the situation. While this may not excuse their behaviour, it will shed light into why things have been out of sorts lately. Around the 31st you will be on fine form at work as you will not only come up with innovating solutions but people will enjoy your motivating spirit. During this time you will be more productive than usual as you encourage the team to work together for a common purpose."
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 601 To 621
            picResults.Picture = LoadPicture(App.Path & "\Newgemini.jpg")
            picResults.Print "Gemini"
            'displays picture of sign and horoscope
            myString = "A change in your domestic setting may be called for on the 3rd when the Full Moon encourages you to spruce things up. Getting rid of clutter and straightening out the cupboards will make your house feel like more of a home and a place where you can begin to relax! Mid month you will get down to the root of the matter within an intimate relationship, as new information will put a different perspective on the situation. While this may not excuse their behaviour, it will shed light into why things have been out of sorts lately. Around the 31st you will be on fine form at work as you will not only come up with innovating solutions but people will enjoy your motivating spirit. During this time you will be more productive than usual as you encourage the team to work together for a common purpose."
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 622 To 630
            picResults.Picture = LoadPicture(App.Path & "\Newcancer.jpg")
            picResults.Print "Cancer"
            'displays picture of sign and horoscope
            myString = "With Mercury, planet of communications going direct on the 9th you will finally be able to open up and really get something off your chest. For the last three weeks you may have had to be tight-lipped and secretive on certain matters but the time has come when you can talk about it. Mid month you may have to do a bit of convincing, but you will eventually be able to influence those you're close to to lend a helping hand towards your cause. Once the ball gets rolling you will be quite pleased to see the constructive changes can all make together. Towards the 22nd, a tense connection between two planets will make you feel torn in two different directions. While speedy Mars wants instant gratification wise old Saturn is digging his heels in making you feel like you're running in circles. Being patient during this time will be easier said than done!"
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 701 To 722
            picResults.Picture = LoadPicture(App.Path & "\Newcancer.jpg")
            picResults.Print "Cancer"
            'displays picture of sign and horoscope
            myString = "With Mercury, planet of communications going direct on the 9th you will finally be able to open up and really get something off your chest. For the last three weeks you may have had to be tight-lipped and secretive on certain matters but the time has come when you can talk about it. Mid month you may have to do a bit of convincing, but you will eventually be able to influence those you're close to to lend a helping hand towards your cause. Once the ball gets rolling you will be quite pleased to see the constructive changes can all make together. Towards the 22nd, a tense connection between two planets will make you feel torn in two different directions. While speedy Mars wants instant gratification wise old Saturn is digging his heels in making you feel like you're running in circles. Being patient during this time will be easier said than done!"
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 723 To 731
            picResults.Picture = LoadPicture(App.Path & "\Newleo.jpg")
            picResults.Print "Leo"
            'displays picture of sign and horoscope
            myString = "When the Full Moon passes through your house of finances at the beginning of the month, it will prompt you to get your head around your spending habits. By doing a basic budget you will be able to cut out unnecessary spending and start saving up for your future. When rational Mercury teams up with powerful Pluto around the 16th you will come up with a creative solution to get to the bottom of a sticky business situation. Although you could choose to turn a blind eye to what's going on, putting a clever spin on the situation could work to your favour. Around the end of the month, you will jump at the chance when an exciting yet time consuming new opportunity comes your way. While you will have to work under immense pressure, you will rise to the challenge and produce some of your most impressive work during this influence!"
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 801 To 822
            picResults.Picture = LoadPicture(App.Path & "\Newleo.jpg")
            picResults.Print "Leo"
            'displays picture of sign and horoscope
            myString = "When the Full Moon passes through your house of finances at the beginning of the month, it will prompt you to get your head around your spending habits. By doing a basic budget you will be able to cut out unnecessary spending and start saving up for your future. When rational Mercury teams up with powerful Pluto around the 16th you will come up with a creative solution to get to the bottom of a sticky business situation. Although you could choose to turn a blind eye to what's going on, putting a clever spin on the situation could work to your favour. Around the end of the month, you will jump at the chance when an exciting yet time consuming new opportunity comes your way. While you will have to work under immense pressure, you will rise to the challenge and produce some of your most impressive work during this influence!"
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 823 To 831
            picResults.Picture = LoadPicture(App.Path & "\Newvirgo.jpg")
            picResults.Print "Virgo"
            'displays picture of sign and horoscope
            myString = "You may be inclined to take on the role of counsellor at the beginning of the month as the burden of old family issues may begin to take their toll. Although the issues at hand may have come from past conflicts, be sure to keep this separate from your own personal relationships. When the New Moon occurs on the 19th you will be interested in brushing up on your inter-personal skills, as you will be overcome with feelings of understanding and empathy. Now would be a good time to lend a shoulder to cry on for a friend in need. Towards the end of the month your charismatic mood will persuade your co-workers to follow your lead in order to make positive changes to the work environment. While your may have ulterior motives for self advancement, everyone will enjoy the joint venture and team spirits will be lifted."
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 901 To 922
            picResults.Picture = LoadPicture(App.Path & "\Newvirgo.jpg")
            picResults.Print "Virgo"
            'displays picture of sign and horoscope
            myString = "You may be inclined to take on the role of counsellor at the beginning of the month as the burden of old family issues may begin to take their toll. Although the issues at hand may have come from past conflicts, be sure to keep this separate from your own personal relationships. When the New Moon occurs on the 19th you will be interested in brushing up on your inter-personal skills, as you will be overcome with feelings of understanding and empathy. Now would be a good time to lend a shoulder to cry on for a friend in need. Towards the end of the month your charismatic mood will persuade your co-workers to follow your lead in order to make positive changes to the work environment. While your may have ulterior motives for self advancement, everyone will enjoy the joint venture and team spirits will be lifted."
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 923 To 930
            picResults.Picture = LoadPicture(App.Path & "\Newlibra.jpg")
            picResults.Print "Libra"
            'displays picture of sign and horoscope
            myString = "When Mercury, planet of mental agility goes direct on the 9th, you will regain confidence in your own creativity. Although you may have had a frustrating couple of weeks of writer's block, the ideas will begin to flow again letting you pick up where you left off. Around the 17th, you won't be shy in letting your affections be known as an amorous connection between romantic Venus and intense Pluto will see you spelling your intentions out - loud and clear. Whether attached or single, the object of your desire doesn't stand a chance against your charms! The end of the month will feel like a battle of wills as a tense combination of planetary energy could leave you feeling despondent. Even though you have enough energy and drive to compete with the best, you can help but feel as though your resources are limited. Use this time to reassess how achievable your goals are."
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 1001 To 1023
            picResults.Picture = LoadPicture(App.Path & "\Newlibra.jpg")
            picResults.Print "Libra"
            'displays picture of sign and horoscope
            myString = "When Mercury, planet of mental agility goes direct on the 9th, you will regain confidence in your own creativity. Although you may have had a frustrating couple of weeks of writer's block, the ideas will begin to flow again letting you pick up where you left off. Around the 17th, you won't be shy in letting your affections be known as an amorous connection between romantic Venus and intense Pluto will see you spelling your intentions out - loud and clear. Whether attached or single, the object of your desire doesn't stand a chance against your charms! The end of the month will feel like a battle of wills as a tense combination of planetary energy could leave you feeling despondent. Even though you have enough energy and drive to compete with the best, you can help but feel as though your resources are limited. Use this time to reassess how achievable your goals are."
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 1024 To 1031
            picResults.Picture = LoadPicture(App.Path & "\Newscorpio.jpg")
            picResults.Print "Scorpio"
            'displays picture of sign and horoscope
            myString = "At the beginning of the month, you'll step out into the world with rose coloured glasses as your outlook towards life becomes glowingly optimistic. The normally cautious Scorpion will have an outgoing and friendly disposition, as you will feel there is no problem big enough that can't be solved. Around the 16th you will be able to cleverly manipulate a domestic situation to your benefit as Mercury the trickster teams up with invincible Pluto. You will be able to put your foot down on unnecessary household spending while striking up a compromising deal with the rest of those you live with. Those who know you well will think you're acting cheeky around the 31st, as your devil-may-care attitude is sure to stir up a bit of controversy within your close relationships. At the moment, you will feel the need to have your own personal space but may be going about it in an unusual way."
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 1101 To 1121
            picResults.Picture = LoadPicture(App.Path & "\Newscorpio.jpg")
            picResults.Print "Scorpio"
            'displays picture of sign and horoscope
            myString = "At the beginning of the month, you'll step out into the world with rose coloured glasses as your outlook towards life becomes glowingly optimistic. The normally cautious Scorpion will have an outgoing and friendly disposition, as you will feel there is no problem big enough that can't be solved. Around the 16th you will be able to cleverly manipulate a domestic situation to your benefit as Mercury the trickster teams up with invincible Pluto. You will be able to put your foot down on unnecessary household spending while striking up a compromising deal with the rest of those you live with. Those who know you well will think you're acting cheeky around the 31st, as your devil-may-care attitude is sure to stir up a bit of controversy within your close relationships. At the moment, you will feel the need to have your own personal space but may be going about it in an unusual way."
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 1122 To 1130
            picResults.Picture = LoadPicture(App.Path & "\Newsagittarius.jpg")
            picResults.Print "Sagittarius"
            'displays picture of sign and horoscope
            myString = "At the beginning of the month, your family or people you live with have a difficult time agreeing with recent changes you've made to your living arrangements. While you maintain your right to live your life on your own terms, other traditionally minded people may not see eye to eye with this. When the New Moon passes through your house of career on the 19th your understanding and flexible attitude will make you stand out in a competitive work environment. Being a good listener goes a long way and soon you will find people opening up to you more than usual. Around the 22nd when Mars, powerhouse of energy meets it's match in slow, restrictive Saturn you feel like throwing in the towel. At the moment it may seem like all of your energies are being wasted on all the endless obligations being thrown your way - but no one ever said perseverance would be easy!"
        stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 1201 To 1221
            picResults.Picture = LoadPicture(App.Path & "\Newsagittarius.jpg")
            picResults.Print "Sagittarius"
            'displays picture of sign and horoscope
            myString = "At the beginning of the month, your family or people you live with have a difficult time agreeing with recent changes you've made to your living arrangements. While you maintain your right to live your life on your own terms, other traditionally minded people may not see eye to eye with this. When the New Moon passes through your house of career on the 19th your understanding and flexible attitude will make you stand out in a competitive work environment. Being a good listener goes a long way and soon you will find people opening up to you more than usual. Around the 22nd when Mars, powerhouse of energy meets it's match in slow, restrictive Saturn you feel like throwing in the towel. At the moment it may seem like all of your energies are being wasted on all the endless obligations being thrown your way - but no one ever said perseverance would be easy!"
            stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
        Case 1222 To 1231
            picResults.Picture = LoadPicture(App.Path & "\Newcapricorn.jpg")
            picResults.Print "Capricorn"
            'displays picture of sign and horoscope
            myString = "A tall tale you tell at the beginning of the month will land you into an unexpected pot of hot water when you accidentally embellish the details of a story. Although you were only trying to make it sound a little more interesting, you will inadvertently ruffle a few feathers!The New Moon will pass through your house of communications on the 19th, making you want to express your inner most thoughts through an artistic medium. Whether it's poetry, writing or even painting, putting pen to paper will be a healthy outlet and have a beneficial effect on your emotions. Around the 31st, a social gathering could leave you with more than you bargained for when your magnetic charm will find you more attention that you know what to do with. While good news for the single Capricorn, the difficult part will be short listing all the potential suitors!"
    stringLength = Len(myString)
        i = 1
    While i + lineLength < stringLength
       tempString = Mid(myString, i, lineLength)
       Pos = InStrRev(tempString, Space)
       tempString = Mid(myString, i, Pos)
       
       picResults.Print tempString
       i = i + Pos
    Wend
    End Select
End Sub

Private Sub cmdLuckydays_Click()
    Dim AstrologicalSign As String
    'Declares variable
    'Cited: Lecture 11
    picResults.Cls
    'Clears picture box
    'Cited: Lecture 11
    AstrologicalSign = InputBox("Enter your Astrological Sign", "What is your sign?")
    'initializes variable
    'Cited: Lecture 11
    Select Case AstrologicalSign
        Case "Capricorn"
            picResults.Print "You are in-line with the planet Saturn"
            picResults.Print " Your lucky days are the 4th and the 8th of the month"
        Case "Aquarius"
            picResults.Print " You are in-line with the planets Saturn and Uranius"
            picResults.Print "Your lucky days are the 4th and the 11th of the month"
        Case "Pisces"
            picResults.Print "You are in-line with the planets Jupiter and Neptune"
            picResults.Print "You lucky days are the 3rd and the 7th of the month"
        Case "Aries"
            picResults.Print "You are in-line with the planet Mars"
            picResults.Print "Your lucky days are the 8th and the 1st of the month"
        Case "Taurus"
            picResults.Print "You are in-line with the planet Venus"
            picResults.Print "Your lucky day is the 6th of the month"
        Case "Gemini"
            picResults.Print "You are in-line with the planet Mercury"
            picResults.Print " Your lucky days are the 3rd and 5th of the month"
        Case "Cancer"
            picResults.Print "You are in-line with the moon"
            picResults.Print "Your luck days are the 2nd and 7th of the month"
        Case "Leo"
            picResults.Print "You are in-line with the sun"
            picResults.Print "Your luck days are the 1st and the 4th of the month"
        Case "Virgo"
            picResults.Print "You are in-line with the planet Mercury"
            picResults.Print "Your lucky days are the 3rd and 5th of the month"
        Case "Libra"
            picResults.Print "You are in-line with the planet Venus"
            picResults.Print "Your lucky day is the 6th of the month"
        Case "Scorpio"
            picResults.Print "You are in-line with the planets Mars and Pluto"
            picResults.Print "Your lucky days are the 8th and 9th of the month"
        Case "Sagittarius"
        picResults.Print "You are in-line with the planet Jupiter"
        picResults.Print "Your lucky days are the 7th and 3rd of the month"
        Case Else
        picResults.Print "You have entered an invalid sign."
        picResults.Print "Please check your spelling and capitalize the first letter only."
    End Select
    'searches for condition that applies to the data input by the user.
    'Cited entire case: Lab 8, Problem 2
    
End Sub
