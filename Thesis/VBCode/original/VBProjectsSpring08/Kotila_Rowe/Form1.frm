VERSION 5.00
Begin VB.Form frm1 
   BackColor       =   &H00000000&
   Caption         =   "Map"
   ClientHeight    =   12540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16455
   LinkTopic       =   "Form1"
   ScaleHeight     =   12540
   ScaleWidth      =   16455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTarget 
      BackColor       =   &H000000FF&
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdPWP 
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdWestRiverPark 
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdGoldMedalPark 
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdCurriePark 
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdElliotPark 
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdTIR 
      BackColor       =   &H00FFFF00&
      Height          =   375
      Left            =   15000
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdJeuneLune 
      BackColor       =   &H00FFFF00&
      Height          =   375
      Left            =   7320
      MaskColor       =   &H00FFFF00&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Return To Main Page"
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10680
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear Screen"
      Height          =   1095
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10680
      Width           =   1335
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Show All"
      Height          =   1095
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10680
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   10095
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   10035
      ScaleWidth      =   15675
      TabIndex        =   6
      Top             =   0
      Width           =   15735
      Begin VB.CommandButton cmdSpooner 
         BackColor       =   &H00FF00FF&
         Height          =   375
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2640
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdMB 
         BackColor       =   &H00FFFF00&
         Height          =   375
         Left            =   14040
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   6360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdOrpheum 
         BackColor       =   &H00FFFF00&
         Height          =   375
         Left            =   4440
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdPantages 
         BackColor       =   &H00FFFF00&
         Height          =   375
         Left            =   5280
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdPlazaParking 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   6480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmd7thStParking 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmd4thStParking 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdMillParking 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdMacysParking 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdOldSpaghetti 
         BackColor       =   &H00FF00FF&
         Height          =   375
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdCafeHavana 
         BackColor       =   &H00FF00FF&
         Height          =   375
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdZelo 
         BackColor       =   &H00FF00FF&
         Height          =   375
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdIchiban 
         BackColor       =   &H00FF00FF&
         Height          =   375
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   6840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdPalomino 
         BackColor       =   &H00FF00FF&
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdNicollet 
         BackColor       =   &H0000FFFF&
         Height          =   495
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdHHH 
         BackColor       =   &H000000FF&
         Height          =   495
         Left            =   10920
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdLoringPark 
         BackColor       =   &H0000FF00&
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdGuthrie 
         BackColor       =   &H00FFFF00&
         Height          =   375
         Left            =   11760
         MaskColor       =   &H00FFFF00&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2880
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdRestaurant 
      BackColor       =   &H00FF00FF&
      Caption         =   "Show Restaurants"
      Height          =   1095
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10680
      Width           =   1455
   End
   Begin VB.CommandButton cmdShop 
      BackColor       =   &H0000FFFF&
      Caption         =   "Show Shopping"
      Height          =   1095
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10680
      Width           =   1455
   End
   Begin VB.CommandButton cmdSports 
      BackColor       =   &H000000FF&
      Caption         =   "Show Sports"
      Height          =   1095
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   10680
      Width           =   1455
   End
   Begin VB.CommandButton cmdpark 
      BackColor       =   &H0000FF00&
      Caption         =   "Show Parks"
      Height          =   1095
      Left            =   9360
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   10680
      Width           =   1455
   End
   Begin VB.CommandButton cmdParkinglot 
      BackColor       =   &H000080FF&
      Caption         =   "Show Parking"
      Height          =   1095
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10680
      Width           =   1455
   End
   Begin VB.CommandButton cmdThea 
      BackColor       =   &H00FFFF00&
      Caption         =   "Show Theaters"
      Height          =   1095
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10680
      Width           =   1455
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Minneapolis Travel Guide
'frmMap
'Kayla Kotila and Chris Rowe
'March 30
'Objective of this form is to display locations and give information about the venues

Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmd4thStParking_Click()
'info about parking
MsgBox "Park here for $7 a day.", , "Parking"
End Sub

Private Sub cmd7thStParking_Click()
'info about parking
MsgBox "Park here for $9 a day.", , "Parking"
End Sub

Private Sub cmdAll_Click()
'shows all places on map
    cmdSpooner.Visible = True
    cmdPWP.Visible = True
    cmdLoringPark.Visible = True
    cmdCurriePark.Visible = True
    cmdElliotPark.Visible = True
    cmdGoldMedalPark.Visible = True
    cmdWestRiverPark.Visible = True
    cmdMillParking.Visible = True
    cmdMacysParking.Visible = True
    cmd4thStParking.Visible = True
    cmd7thStParking.Visible = True
    cmdPlazaParking.Visible = True
    cmdIchiban.Visible = True
    cmdPalomino.Visible = True
    cmdZelo.Visible = True
    cmdCafeHavana.Visible = True
    cmdOldSpaghetti.Visible = True
    cmdNicollet.Visible = True
    cmdHHH.Visible = True
    cmdTarget.Visible = True
    cmdTIR.Visible = True
    cmdGuthrie.Visible = True
    cmdJeuneLune.Visible = True
    cmdMB.Visible = True
    cmdPantages.Visible = True
    cmdOrpheum.Visible = True
    cmdThea.Caption = "Hide Theaters"
    cmdpark.Caption = "Hide Parks"
    cmdParkinglot.Caption = "Hide Parking"
    cmdRestaurant.Caption = "Hide Restaurants"
    cmdShop.Caption = "Hide Shopping"
    cmdSports.Caption = "Hide Sports"
    
End Sub

Private Sub cmdCafeHavana_Click()
'info about cafe havana
MsgBox "Cafe Havana serves delicious Cuban food in an elegant setting. We reccomend trying the tuna. Free Valet.", , "Cafe Havana"
End Sub

Private Sub cmdClear_Click()
'Clears all locations from map
    cmdSpooner.Visible = False
    cmdMB.Visible = False
    cmdPantages.Visible = False
    cmdOrpheum.Visible = False
    cmdPWP.Visible = False
    cmdLoringPark.Visible = False
    cmdCurriePark.Visible = False
    cmdElliotPark.Visible = False
    cmdGoldMedalPark.Visible = False
    cmdWestRiverPark.Visible = False
    cmdMillParking.Visible = False
    cmdMacysParking.Visible = False
    cmd4thStParking.Visible = False
    cmd7thStParking.Visible = False
    cmdPlazaParking.Visible = False
    cmdIchiban.Visible = False
    cmdPalomino.Visible = False
    cmdZelo.Visible = False
    cmdCafeHavana.Visible = False
    cmdOldSpaghetti.Visible = False
    cmdNicollet.Visible = False
    cmdHHH.Visible = False
    cmdTarget.Visible = False
    cmdTIR.Visible = False
    cmdGuthrie.Visible = False
    cmdJeuneLune.Visible = False
    cmdThea.Caption = "Show Theaters"
    cmdpark.Caption = "Show Parks"
    cmdParkinglot.Caption = "Show Parking"
    cmdRestaurant.Caption = "Show Restaurants"
    cmdShop.Caption = "Show Shopping"
    cmdSports.Caption = "Show Sports"
End Sub



Private Sub cmdCurriePark_Click()
'info about park
ShellExecute Me.hWnd, "open", "http://www.minneapolisparks.org/default.asp?pageid=4&parkid=423", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdElliotPark_Click()
'info about park
ShellExecute Me.hWnd, "open", "http://www.minneapolisparks.org/default.asp?PageID=4&parkid=278", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdGoldMedalPark_Click()
'info about park

ShellExecute Me.hWnd, "open", "http://www.oaala.com/projects/gold_medal_park/gold_medal.htm", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdGuthrie_Click()
'info about theater

ShellExecute Me.hWnd, "open", "http://www.guthrietheater.org", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdHHH_Click()
'info about metrodome
    ShellExecute Me.hWnd, "open", "http://www.msfc.com", "", "", SW_SHOW Or SW_NORMAL
End Sub


Private Sub cmdIchiban_Click()
'info about ichiban
ShellExecute Me.hWnd, "open", "http://www.ichiban.ca", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdJeuneLune_Click()
'info about theater
ShellExecute Me.hWnd, "open", "http://www.jeunelune.com", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdLoringPark_Click()
'info about park
ShellExecute Me.hWnd, "open", "http://www.minneapolisparks.org/default.asp?PageID=4&parkid=199", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdMacysParking_Click()
'info about parking
MsgBox "Park here for $10 a day.", , "Parking"
End Sub

Private Sub cmdMain_Click()
    frm1.Hide
    frm2.Show
End Sub



Private Sub cmdMB_Click()
'info about place
ShellExecute Me.hWnd, "open", "http://www.mixedblood.org", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdMillParking_Click()
'info about place
MsgBox "Park here for only $8 a day.", , "Parking"
End Sub

Private Sub cmdNicollet_Click()
'info about place

ShellExecute Me.hWnd, "open", "http://en.wikipedia.org/wiki/Nicollet_Mall", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdOldSpaghetti_Click()
'info about place
ShellExecute Me.hWnd, "open", "http://www.osf.com", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdOrpheum_Click()
'info about place
ShellExecute Me.hWnd, "open", "http://www.hennepintheaterdistrict.org", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdPalomino_Click()
'info about place
MsgBox "Palamino serves Mediterranean food at a moderate price. Great Service"
End Sub

Private Sub cmdPantages_Click()
'info about place
ShellExecute Me.hWnd, "open", "http://www.hennepintheaterdistrict.org", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdpark_Click()
'shows/hides parks on map
If cmdPWP.Visible = True Then
    cmdPWP.Visible = False
    cmdLoringPark.Visible = False
    cmdCurriePark.Visible = False
    cmdElliotPark.Visible = False
    cmdGoldMedalPark.Visible = False
    cmdWestRiverPark.Visible = False
    cmdpark.Caption = "Show Parks"
Else
    cmdPWP.Visible = True
    cmdLoringPark.Visible = True
    cmdCurriePark.Visible = True
    cmdElliotPark.Visible = True
    cmdGoldMedalPark.Visible = True
    cmdWestRiverPark.Visible = True
    cmdpark.Caption = "Hide Parks"
End If
    
End Sub


    


Private Sub cmdParkinglot_Click()
'show/hide parking on map
If cmdMillParking.Visible = True Then
    cmdMillParking.Visible = False
    cmdMacysParking.Visible = False
    cmd4thStParking.Visible = False
    cmd7thStParking.Visible = False
    cmdPlazaParking.Visible = False
    cmdParkinglot.Caption = "Show Parking"
Else
    cmdMillParking.Visible = True
    cmdMacysParking.Visible = True
    cmd4thStParking.Visible = True
    cmd7thStParking.Visible = True
    cmdPlazaParking.Visible = True
    cmdParkinglot.Caption = "Hide Parking"
End If
    
End Sub

Private Sub cmdPlazaParking_Click()
'info about place
MsgBox "Park here for $10 a day.", , "Parking"
End Sub

Private Sub cmdPWP_Click()
'info about place
ShellExecute Me.hWnd, "open", "http://www.citiesarchitecture.com/Building/1983/Pillsbury_Park.php", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdRestaurant_Click()
'show/hide resturants on map
If cmdIchiban.Visible = True Then
    cmdIchiban.Visible = False
    cmdPalomino.Visible = False
    cmdZelo.Visible = False
    cmdCafeHavana.Visible = False
    cmdOldSpaghetti.Visible = False
    cmdSpooner.Visible = False
    cmdRestaurant.Caption = "Show Restaurants"
Else
     cmdIchiban.Visible = True
    cmdPalomino.Visible = True
    cmdZelo.Visible = True
    cmdCafeHavana.Visible = True
    cmdOldSpaghetti.Visible = True
    cmdSpooner.Visible = True
    cmdRestaurant.Caption = "Hide Restaurants"
End If
    
End Sub

Private Sub cmdShop_Click()
'show/hide shopping places on map
If cmdNicollet.Visible = True Then
    cmdNicollet.Visible = False
    cmdShop.Caption = "Show Shopping"
Else
    cmdShop.Caption = "Hide Shopping"
    cmdNicollet.Visible = True
End If

End Sub

Private Sub cmdSpooner_Click()
'info about place
ShellExecute Me.hWnd, "open", "http://www.spoonriverrestaurant.com", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdSports_Click()
'show/hide sporting places on map
If cmdHHH.Visible = True Then
    cmdHHH.Visible = False
    cmdTarget.Visible = False
    cmdSports.Caption = "Show Sports"
Else
    cmdHHH.Visible = True
    cmdTarget.Visible = True
    cmdSports.Caption = "Hide Sports"
End If
End Sub

Private Sub cmdTarget_Click()
'info about place
   
    ShellExecute Me.hWnd, "open", "http://www.targetcenter.com", "", "", SW_SHOW Or SW_NORMAL
    
    
End Sub

Private Sub cmdThea_Click()
'show/hide theaters on map
If cmdTIR.Visible = True Then
    cmdTIR.Visible = False
    cmdGuthrie.Visible = False
    cmdJeuneLune.Visible = False
    cmdMB.Visible = False
    cmdPantages.Visible = False
    cmdOrpheum.Visible = False
    cmdThea.Caption = "Show Theaters"
Else
    cmdTIR.Visible = True
    cmdGuthrie.Visible = True
    cmdJeuneLune.Visible = True
    cmdMB.Visible = True
    cmdPantages.Visible = True
    cmdOrpheum.Visible = True
    cmdThea.Caption = "Hide Theaters"
End If
End Sub

Private Sub cmdTIR_Click()
'info about place
ShellExecute Me.hWnd, "open", "http://www.theatreintheround.org", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub cmdWestRiverPark_Click()
'info about place
MsgBox "West River Park is located near the historic Mill City Museum. We highly reccomend taking a bike ride through this park.", , "West River Park"
End Sub

Private Sub cmdZelo_Click()
'info about place
ShellExecute Me.hWnd, "open", "http://www.zelomn.com", "", "", SW_SHOW Or SW_NORMAL
End Sub

