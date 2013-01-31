VERSION 5.00
Begin VB.Form frmMainScreen
   BackColor       =   &H0000FFFF&
   Caption         =   "LEMONADE!"
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13890
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   9840
   ScaleWidth      =   13890
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMake
      Height          =   4455
      Left            =   9600
      ScaleHeight     =   4395
      ScaleWidth      =   4035
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.PictureBox picShow
      Height          =   4455
      Left            =   120
      ScaleHeight     =   4395
      ScaleWidth      =   4395
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton cmdMake
      Caption         =   "Make some lemonade!"
      Height          =   975
      Left            =   10560
      TabIndex        =   8
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdStats
      Caption         =   "Show my statistics (Cash, Stuff, Advertisements)"
      Height          =   735
      Left            =   4560
      TabIndex        =   7
      Top             =   120
      Width           =   4815
   End
   Begin VB.PictureBox picStats
      Height          =   3015
      Left            =   3960
      ScaleHeight     =   2955
      ScaleWidth      =   5955
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.CommandButton cmdQuit
      Caption         =   "Quit"
      Height          =   735
      Left            =   12600
      TabIndex        =   5
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton cmdAds
      Caption         =   "Advertise"
      Height          =   975
      Left            =   10080
      TabIndex        =   4
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdRecipe
      Caption         =   "Change Recipe or the Price per Cup"
      Height          =   975
      Left            =   2040
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdBuy
      Caption         =   "Buy Stuff"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.PictureBox picScene
      Height          =   4455
      Left            =   4680
      ScaleHeight     =   4395
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.CommandButton cmdBeginDay
      Caption         =   "Commence with the commerce!"
      Height          =   735
      Left            =   4440
      MaskColor       =   &H00404040&
      TabIndex        =   0
      Top             =   9000
      Width           =   5295
   End
End
Attribute VB_Name = "frmMainScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAds_Click()
    picScene.Visible = False
    frmMainScreen.Hide
    frmAdScreen.Show
End Sub

Private Sub cmdBeginDay_Click()
    If fullpitcher = 0 Then
        MsgBox EnterName & ", You need to make some lemonade, duh. Go make up a pitcher and try again."
    Else
    Dim Tinput As String, CTR As Integer, Tactual(1 To 100) As String, Ttemp(1 To 100) As Integer, Pos1 As Integer, foundit As Boolean, temp As Integer
    Tinput = InputBox(EnterName & ", As an 8 year old, your weather prediction strategy is mildly haphazard. Type a letter (lowercase or capital) to predict.")
    Day = Day + 1
    foundit = False

    Open App.Path & "\Temps.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Tactual(CTR), Ttemp(CTR)
    Loop
    Close #1

    Do Until foundit = True Or Pos1 > CTR
        Pos1 = Pos1 + 1
        If Tinput = Tactual(Pos1) Then
            foundit = True
        End If
    Loop

    Dim RandTemp As Integer
    Randomize
    RandTemp = Int((2 * Rnd) + 1)

    If foundit = True Then
        temp = Ttemp(Pos1 - RandTemp)
    End If

    Dim RandWeather As Integer
    Randomize
    RandWeather = Int((100 * Rnd) + 1)

    Dim WType(1 To 100) As String, WBonus(1 To 100) As Single, CTR2 As Integer, WTypeActual As String, WBonusActual As Single

    Open App.Path & "\Weather.txt" For Input As #2
    Do Until EOF(2)
        CTR2 = CTR2 + 1
        Input #2, WType(CTR2), WBonus(CTR2)
    Loop
    Close #2

    If RandWeather <= 5 Then
        WTypeActual = WType(1)
        WBonusActual = WBonus(1)
        ElseIf RandWeather <= 18 Then
            WTypeActual = WType(2)
            WBonusActual = WBonus(2)
        ElseIf RandWeather <= 30 Then
            WTypeActual = WType(3)
            WBonusActual = WBonus(3)
        ElseIf RandWeather <= 50 Then
            WTypeActual = WType(4)
            WBonusActual = WBonus(4)
        ElseIf RandWeather <= 76 Then
            WTypeActual = WType(5)
            WBonusActual = WBonus(5)
        Else
            WTypeActual = WType(6)
            WBonusActual = WBonus(6)
    End If

    Dim totalppl As Long, Recipe1 As Single, Recipe2 As Single, willingppl As Long
    totalppl = (2 * temp * Fame * WBonusActual) - 100

    If totalppl < 0 Then
        totalppl = 0
    End If

    If RecipeS / RecipeL <= 0.1 Then
        Recipe1 = 0
        ElseIf RecipeS / RecipeL <= 0.2 Then
            Recipe1 = 0.05
        ElseIf RecipeS / RecipeL <= 0.25 Then
            Recipe1 = 0.1
        ElseIf RecipeS / RecipeL <= 0.34 Then
            Recipe1 = 0.2
        ElseIf RecipeS / RecipeL <= 0.4 Then
            Recipe1 = 0.25
        ElseIf RecipeS / RecipeL <= 0.5 Then
            Recipe1 = 0.35
        ElseIf RecipeS / RecipeL <= 0.6 Then
            Recipe1 = 0.4
        ElseIf RecipeS / RecipeL <= 0.7 Then
            Recipe1 = 0.3
        ElseIf RecipeS / RecipeL <= 0.75 Then
            Recipe1 = 0.175
            Dim aaaa as String
        ElseIf RecipeS / RecipeL <= 0.8 Then
            Recipe1 = 0.1
        ElseIf RecipeS / RecipeL <= 1 Then
            Recipe1 = 0.075
        ElseIf RecipeS / RecipeL <= 1.5 Then
            Recipe1 = 0.05
        Else
            Recipe1 = 0
    End If

    If RecipeI / RecipeS < 0.5 Then
        Recipe2 = 1
        ElseIf RecipeI / RecipeS <= 0.75 Then
            Recipe2 = 1.5
        ElseIf RecipeI / RecipeS <= 0.99 Then
            Recipe2 = 1.75
        ElseIf RecipeI / RecipeS <= 1.25 Then
            Recipe2 = 2
        ElseIf RecipeI / RecipeS = 1.5 Then
            Recipe2 = 2.5
        ElseIf RecipeI / RecipeS <= 2 Then
            Recipe2 = 1.75
        ElseIf RecipeI / RecipeS <= 4 Then
            Recipe2 = 1.25
        Else
            Recipe2 = 1
    End If

    Dim possPrice As Single, Recipecalc As Single, PWilling As Single, chargematch As Boolean
    chargematch = False
    Recipecalc = Recipe1 * Recipe2
    willingppl = totalppl * Recipecalc

    If Recipecalc <= 0.1 Then
        possPrice = 0.1
        ElseIf Recipecalc <= 0.2 Then
            possPrice = 0.25
        ElseIf Recipecalc <= 0.3 Then
            possPrice = 0.5
        ElseIf Recipecalc <= 0.4 Then
            possPrice = 0.75
        ElseIf Recipecalc <= 0.5 Then
            possPrice = 1
        ElseIf Recipecalc <= 0.6 Then
            possPrice = 3
        ElseIf Recipecalc <= 0.7 Then
            possPrice = 5
        ElseIf Recipecalc <= 0.8 Then
            possPrice = 10
        ElseIf Recipecalc <= 0.9 Then
            possPrice = 25
        ElseIf Recipecalc < 1 Then
            possPrice = 50
        Else
            possPrice = 1000
    End If

    Dim CTR3 As Integer, possPricemultiplier(1 To 100) As Single, percEarned(1 To 100) As Single
    Open App.Path & "\MoneyEarnedPercent.txt" For Input As #3
    Do Until EOF(3)
        CTR3 = CTR3 + 1
        Input #3, possPricemultiplier(CTR3), percEarned(CTR3)
    Loop
    Close #3

    Dim count As Integer, finder As Boolean, Subtotal As Single
    finder = False

    Do Until finder = True Or count > CTR3
    count = count + 1
        If charged <= possPrice * possPricemultiplier(count) Then
            finder = True
        End If
    Loop

    Dim totalpurchases As Single
    If finder = True Then
        totalpurchases = willingppl * percEarned(count)
    End If

    Dim Cupcounter As Integer
    Cupcounter = fullpitcher * 25

    Do Until fullpitcher = 0
        fullpitcher = fullpitcher - 1
        Pitchers = Pitchers + 1
    Loop
    Dim finalcount As Integer
        If Cupcounter > totalpurchases Then
            Cups = Cups - Int(totalpurchases)
            Subtotal = Int(totalpurchases) * charged
            finalcount = Int(totalpurchases)
            Dim bbbb as Integer
        ElseIf Cupcounter <= totalpurchases Then
            Cups = Cups - Cupcounter
            Subtotal = Cupcounter * charged
            finalcount = Cupcounter
        End If

    Cash = Cash + Subtotal
    MsgBox "The Temperature of Day " & Day & " was " & temp & " degrees. Today's weather was " & WTypeActual & ". " & totalppl & " people walked past your stand today, but only " & willingppl & " people were interested in your lemonade. " & Int(finalcount) & " people actually bought your lemonade. You earned " & FormatCurrency(Subtotal) & " today."
    picScene.Visible = True
    picScene.Picture = LoadPicture(App.Path & "\lemonade.bmp")
    End If

    If Day = 30 Then
        frmMainScreen.Hide
        frmFinal.Show
    End If

End Sub

Private Sub cmdBuy_Click()
    picScene.Visible = False
    frmMainScreen.Hide
            Dim cccc as Long
    frmBuy.Show
End Sub

Private Sub cmdMake_Click()
    picScene.Visible = False

    If Lemons >= RecipeL And Ice >= RecipeI And Sugar >= RecipeS And Pitchers - 1 >= 0 Then
        Lemons = Lemons - RecipeL
        Ice = Ice - RecipeI
        Sugar = Sugar - RecipeS
        Pitchers = Pitchers - 1
        fullpitcher = fullpitcher + 1
    Else
        MsgBox "Buy more stuff or adjust your recipe -- you don't have enough for a pitcher!"
    End If

    picMake.Visible = True
    Dim Randpic As Integer
    Randomize
    Randpic = Int((3 * Rnd) + 1)
    If Randpic = 1 Then
        picMake.Picture = LoadPicture(App.Path & "\lemonpitcher.jpg")
    ElseIf Randpic = 2 Then
        picMake.Picture = LoadPicture(App.Path & "\lemons.jpg")
    Else
        picMake.Picture = LoadPicture(App.Path & "\lemonade.bmp")
    End If

        picStats.Visible = True
    picStats.Cls
    picStats.Print "Cash = " & FormatCurrency(Cash)
    picStats.Print "Price per cup = " & FormatCurrency(charged)
    picStats.Print "***************************************************************************************************"
    picStats.Print
            Dim dddd as Single
    picStats.Print "Lemons", "Sugar", "Ice", "Cups", "Pitchers"
    picStats.Print Lemons, Sugar, Ice, Cups, Pitchers
    picStats.Print
    picStats.Print "***************************************************************************************************"
    picStats.Print "Bonus from Ads = "; Fame - 1
    picStats.Print
    picStats.Print "***************************************************************************************************"
    picStats.Print "Number of full pitchers: " & fullpitcher
    picStats.Print
    picStats.Print "***************************************************************************************************"
    picStats.Print "Today is the morning of Day: " & Day + 1

End Sub

Private Sub cmdQuit_Click()
    frmMainScreen.Hide
    frmFinal.Show
End Sub

Private Sub cmdRecipe_Click()
    picScene.Visible = False
    frmMainScreen.Hide
    frmRecipe.Show
End Sub

Private Sub cmdStats_Click()
    picShow.Visible = True
    picStats.Visible = True
    picStats.Cls
    picStats.Print "Cash = " & FormatCurrency(Cash)
    picStats.Print "Price per cup = " & FormatCurrency(charged)
    picStats.Print "***************************************************************************************************"
    picStats.Print
    picStats.Print "Lemons", "Sugar", "Ice", "Cups", "Pitchers"
    picStats.Print Lemons, Sugar, Ice, Cups, Pitchers
    picStats.Print
    picStats.Print "***************************************************************************************************"
    picStats.Print "Bonus from Ads = "; Fame - 1
    picStats.Print
    picStats.Print "***************************************************************************************************"
    picStats.Print "Number of full pitchers: " & fullpitcher
    picStats.Print
    picStats.Print "***************************************************************************************************"
    picStats.Print "Today is the morning of Day: " & Day + 1

    Dim Randpic2 As Integer
    Randomize
    Randpic2 = Int((3 * Rnd) + 1)
    If Randpic2 = 1 Then
        picShow.Picture = LoadPicture(App.Path & "\lemonwedge.jpg")
    ElseIf Randpic2 = 2 Then
        picShow.Picture = LoadPicture(App.Path & "\lemonpeel.jpg")
    Else
        picShow.Picture = LoadPicture(App.Path & "\lemonade.bmp")
    End If
End Sub

