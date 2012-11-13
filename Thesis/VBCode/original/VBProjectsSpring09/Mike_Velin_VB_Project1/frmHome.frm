VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H000000C0&
   Caption         =   "frmHome"
   ClientHeight    =   14415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   Picture         =   "frmHome.frx":0000
   ScaleHeight     =   14415
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEntertainmentCenters 
      Caption         =   "Entertainment Center"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   79
      Top             =   13440
      Width           =   1215
   End
   Begin VB.CommandButton cmdUmbrellas 
      Caption         =   "Umbrella"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   78
      Top             =   13440
      Width           =   1215
   End
   Begin VB.CommandButton cmdGazebos 
      Caption         =   "Gazebo"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   77
      Top             =   13440
      Width           =   1215
   End
   Begin VB.CommandButton cmdPatioDinningSets 
      Caption         =   "Patio Dinning Set"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   76
      Top             =   13440
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrills 
      Caption         =   "Grill"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   75
      Top             =   13440
      Width           =   1215
   End
   Begin VB.CommandButton cmdDinningTables 
      Caption         =   "Dinner Table"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   74
      Top             =   12600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSofas 
      Caption         =   "Sofa"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   73
      Top             =   12600
      Width           =   1215
   End
   Begin VB.CommandButton cmdChairs 
      Caption         =   "Chair"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   72
      Top             =   12600
      Width           =   1215
   End
   Begin VB.CommandButton cmdBookcases 
      Caption         =   "Bookcase"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   71
      Top             =   12600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOfficeChairs 
      Caption         =   "Office Chair"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   70
      Top             =   12600
      Width           =   1215
   End
   Begin VB.CommandButton cmdDesks 
      Caption         =   "Desk"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   69
      Top             =   11760
      Width           =   1215
   End
   Begin VB.CommandButton cmdDressers 
      Caption         =   "Dresser"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   68
      Top             =   11760
      Width           =   1215
   End
   Begin VB.CommandButton cmdNightstands 
      Caption         =   "Nightstand"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   67
      Top             =   11760
      Width           =   1215
   End
   Begin VB.CommandButton cmdTowels 
      Caption         =   "Towel"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   66
      Top             =   11760
      Width           =   1215
   End
   Begin VB.CommandButton cmdMattresses 
      Caption         =   "Mattress"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   65
      Top             =   11760
      Width           =   1215
   End
   Begin VB.CommandButton cmdBlankets 
      Caption         =   "Blanket"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   64
      Top             =   10920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSheets 
      Caption         =   "Sheets"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   63
      Top             =   10920
      Width           =   1215
   End
   Begin VB.CommandButton cmdDuvets 
      Caption         =   "Duvet"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   62
      Top             =   10920
      Width           =   1215
   End
   Begin VB.CommandButton cmdComforters 
      Caption         =   "Comforter"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   61
      Top             =   10920
      Width           =   1215
   End
   Begin VB.CommandButton cmdDrinkware 
      Caption         =   "Drinkware"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   60
      Top             =   10920
      Width           =   1215
   End
   Begin VB.TextBox txtPriceCheck 
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      TabIndex        =   23
      Top             =   7200
      Width           =   3375
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Price Check"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   22
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdStorageBins 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Storage Bin"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   21
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton cmdLandryBins 
      Caption         =   "Laundry Bin"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   20
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton cmdShelving 
      Caption         =   "Shelving"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   19
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton cmdRugs 
      Caption         =   "Rugs"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   18
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCurtains 
      Caption         =   "Curtain"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   17
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton cmdTableLamps 
      Caption         =   "Table Lamp"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   16
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton cmdFloorLamp 
      Caption         =   "Floor Lamp"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   15
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton cmdPillows 
      Caption         =   "Pillow"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   14
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton cmdMirrors 
      Caption         =   "Mirror"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   13
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton cmdClocks 
      Caption         =   "Clock"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   12
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton cmdMicrowaves 
      Caption         =   "Microwave"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton cmdBlenders 
      Caption         =   "Blender"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   10
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefridgerator 
      Caption         =   "Refridgerator"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   9
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton cmdToasters 
      Caption         =   "Toaster"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   8
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton cmdDinnerware 
      Caption         =   "Dinnerware"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   7
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton cmdElectronics 
      Caption         =   "Electronics Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdApparel 
      Caption         =   "Apparel Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdShoes 
      Caption         =   "Shoe Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6960
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdHomePage 
      Caption         =   "Home Page"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdToys 
      Caption         =   "Toy Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9240
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdShoppingCart 
      Caption         =   "Shopping Cart"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblEntertainmentCenters 
      BackColor       =   &H000000C0&
      Caption         =   "35.  Entertainment         Centers"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   59
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label lblUmbrellas 
      BackColor       =   &H000000C0&
      Caption         =   "34.  Umbrellas"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   58
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label lblGazebos 
      BackColor       =   &H000000C0&
      Caption         =   "33.  Gazebos"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   57
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label lblPatioDinningSets 
      BackColor       =   &H000000C0&
      Caption         =   "32.  Patio Dinning Sets"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   56
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label lblGrills 
      BackColor       =   &H000000C0&
      Caption         =   "31.  Grills"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   55
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label lblDinningTables 
      BackColor       =   &H000000C0&
      Caption         =   "30.  Dinning Tables"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   54
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label lblSofas 
      BackColor       =   &H000000C0&
      Caption         =   "29.  Sofas"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   53
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblChairs 
      BackColor       =   &H000000C0&
      Caption         =   "28.  Chairs"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   52
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblBookcases 
      BackColor       =   &H000000C0&
      Caption         =   "27.  Bookcases"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   51
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblOfficeChairs 
      BackColor       =   &H000000C0&
      Caption         =   "26.  Office Chairs"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   50
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblDesks 
      BackColor       =   &H000000C0&
      Caption         =   "25.  Desks"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   49
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label lblDressers 
      BackColor       =   &H000000C0&
      Caption         =   "24.  Dressers"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   48
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label lblNightstands 
      BackColor       =   &H000000C0&
      Caption         =   "23.  Nightstands"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   47
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label lblTowels 
      BackColor       =   &H000000C0&
      Caption         =   "22.  Towels"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   46
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label lblMattresses 
      BackColor       =   &H000000C0&
      Caption         =   "21.  Mattresses"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   45
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label lblBlankets 
      BackColor       =   &H000000C0&
      Caption         =   "20.  Blankets"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   44
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblSheets 
      BackColor       =   &H000000C0&
      Caption         =   "19.  Sheets"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   43
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblDuvets 
      BackColor       =   &H000000C0&
      Caption         =   "18.  Duvets"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   42
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblComforters 
      BackColor       =   &H000000C0&
      Caption         =   "17.  Comforters"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   41
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblDrinkware 
      BackColor       =   &H000000C0&
      Caption         =   "16.  Drinkware"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblStorageBins 
      BackColor       =   &H000000C0&
      Caption         =   "1.  Storage Bins"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblLaundryBins 
      BackColor       =   &H000000C0&
      Caption         =   "2.  Laundry Bins"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   38
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblShelving 
      BackColor       =   &H000000C0&
      Caption         =   "3.  Shelving"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   37
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblRugs 
      BackColor       =   &H000000C0&
      Caption         =   "4.  Rugs"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   36
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblCurtains 
      BackColor       =   &H000000C0&
      Caption         =   "5.  Curtains"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   35
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblTableLamps 
      BackColor       =   &H000000C0&
      Caption         =   "6.  Table Lamps"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblFloorLamps 
      BackColor       =   &H000000C0&
      Caption         =   "7.  Floor Lamps"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   33
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblPillows 
      BackColor       =   &H000000C0&
      Caption         =   "8.  Pillows"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   32
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label lblMirrors 
      BackColor       =   &H000000C0&
      Caption         =   "9.  Mirrors"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   31
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblClocks 
      BackColor       =   &H000000C0&
      Caption         =   "10.  Clocks"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   30
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblMicrowaves 
      BackColor       =   &H000000C0&
      Caption         =   "11.  Microwaves"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   29
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lblBlenders 
      BackColor       =   &H000000C0&
      Caption         =   "12.  Blenders"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   28
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lblRefridgerators 
      BackColor       =   &H000000C0&
      Caption         =   "13.  Refridgerators"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   27
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label lblToasters 
      BackColor       =   &H000000C0&
      Caption         =   "14.  Toasters"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   26
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblDinnerware 
      BackColor       =   &H000000C0&
      Caption         =   "15.  Dinnerware"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   25
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lblEnter 
      BackColor       =   &H000000C0&
      Caption         =   "Enter a Number to Price Check the Item   -------->"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   24
      Top             =   7200
      Width           =   4095
   End
   Begin VB.Label lblHome 
      BackColor       =   &H000000C0&
      Caption         =   "Home Department"
      BeginProperty Font 
         Name            =   "GentiumAlt"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApparel_Click()
        frmHome.Hide
        frmApparel.Show
End Sub

Private Sub cmdBlankets_Click()
        Blankets = InputBox("Enter how many Blankets you would like to buy", "Blankets")
End Sub

Private Sub cmdBlenders_Click()
        Blenders = InputBox("Enter how many Blenders you would like to buy", "Blenders")
End Sub

Private Sub cmdBookcases_Click()
        Bookcases = InputBox("Enter how many Bookcases you would like to buy", "Bookcases")
End Sub

Private Sub cmdChairs_Click()
        Chairs = InputBox("Enter how many Chairs you would like to buy", "Chairs")
End Sub

Private Sub cmdCheck_Click()
        Dim CTR As Integer, Home(1 To 50) As String, HomeNumber(1 To 50) As Integer, HomePrice(1 To 50) As Single, Found As Boolean, Pos As Integer, Number As String
        Open App.Path & "\Home.txt" For Input As #1
        CTR = 0
        Do While Not EOF(1)
            CTR = CTR + 1
            Input #1, HomeNumber(CTR), Home(CTR), HomePrice(CTR)
        Loop
        Close #1
        Found = False
        Pos = 1
        Number = txtPriceCheck.Text
        Do While Not Found And Pos <= CTR
            If Number = HomeNumber(Pos) Then
                MsgBox Home(Pos) & " cost " & HomePrice(Pos) & " dollars", , "Price Check"
                Found = True
            End If
            Pos = Pos + 1
        Loop
        If Found = False Then
            MsgBox "Try Again!", , "Error"
        End If
End Sub

Private Sub cmdClocks_Click()
        Clocks = InputBox("Enter how many Clocks you would like to buy", "Clocks")
End Sub

Private Sub cmdComforters_Click()
        Comforters = InputBox("Enter how many Comforters you would like to buy", "Comforters")
End Sub

Private Sub cmdCurtains_Click()
        Curtains = InputBox("Enter how many Curtains you would like to buy", "Curtains")
End Sub

Private Sub cmdDesks_Click()
        Desks = InputBox("Enter how many Desks you would like to buy", "Desks")
End Sub

Private Sub cmdDinnerware_Click()
        Dinnerware = InputBox("Enter how many Dinnerware you would like to buy", "Dinnerware")
End Sub

Private Sub cmdDinningTables_Click()
        DinningTables = InputBox("Enter how many Dinning Tables you would like to buy", "Dinning Tables")
End Sub

Private Sub cmdDressers_Click()
        Dressers = InputBox("Enter how many Dressers you would like to buy", "Dressers")
End Sub

Private Sub cmdDrinkware_Click()
        Drinkware = InputBox("Enter how many Drinkware you would like to buy", "Drinkware")
End Sub

Private Sub cmdDuvets_Click()
        Duvets = InputBox("Enter how many Duvets you would like to buy", "Duvets")
End Sub

Private Sub cmdElectronics_Click()
        frmHome.Hide
        frmElectronics.Show
End Sub

Private Sub cmdEntertainmentCenters_Click()
        EntertainmentCenters = InputBox("Enter how many Entertainment Centers you would like to buy", "Entertainment Centers")
End Sub

Private Sub cmdFloorLamp_Click()
        FloorLamps = InputBox("Enter how many Floor Lamps you would like to buy", "Floor Lamps")
End Sub

Private Sub cmdGazebos_Click()
        Gazebos = InputBox("Enter how many Gazebo's you would like to buy", "Gazebo's")
End Sub

Private Sub cmdGrills_Click()
        Grills = InputBox("Enter how many Grills you would like to buy", "Grills")
End Sub

Private Sub cmdHomePage_Click()
        frmHome.Hide
        frmTarget.Show
End Sub

Private Sub cmdLandryBins_Click()
        LaundryBins = InputBox("Enter how many Laundry Bins you would like to buy", "Laundry Bins")
End Sub

Private Sub cmdMattresses_Click()
        Mattresses = InputBox("Enter how many Mattresses you would like to buy", "Mattresses")
End Sub

Private Sub cmdMicrowaves_Click()
        Microwaves = InputBox("Enter how many Microwaves you would like to buy", "Microwaves")
End Sub

Private Sub cmdMirrors_Click()
        Mirrors = InputBox("Enter how many Mirrors you would like to buy", "Mirrors")
End Sub

Private Sub cmdNightstands_Click()
        Nightstands = InputBox("Enter how many Nightstands you would like to buy", "Nightstands")
End Sub

Private Sub cmdOfficeChairs_Click()
        OfficeChairs = InputBox("Enter how many Office Chairs you would like to buy", "Office Chairs")
End Sub

Private Sub cmdPatioDinningSets_Click()
        PatioDinningSets = InputBox("Enter how many Patio Dinning Set's you would like to buy", "Patio TDinning Set's")
End Sub

Private Sub cmdPillows_Click()
        Pillows = InputBox("Enter how many Pillows you would like to buy", "Pillows")
End Sub

Private Sub cmdRefridgerator_Click()
        Refridgerators = InputBox("Enter how many Refridgerators you would like to buy", "Refridgerators")
End Sub

Private Sub cmdRugs_Click()
        Rugs = InputBox("Enter how many Rugs you would like to buy", "Rugs")
End Sub

Private Sub cmdSheets_Click()
        Sheets = InputBox("Enter how many Sheets you would like to buy", "Sheets")
End Sub

Private Sub cmdShelving_Click()
        Shelving = InputBox("Enter how much Shelving you would like to buy", "Shelving")
End Sub

Private Sub cmdShoes_Click()
        frmHome.Hide
        frmShoes.Show
End Sub

Private Sub cmdShoppingCart_Click()
        frmHome.Hide
        frmShoppingCart.Show
End Sub

Private Sub cmdSofas_Click()
        Sofas = InputBox("Enter how many Sofas you would like to buy", "Sofas")
End Sub

Private Sub cmdStorageBins_Click()
        StorageBins = InputBox("Enter how many Storage Bins you would like to buy", "Storage Bins")
End Sub

Private Sub cmdTableLamps_Click()
        TableLamps = InputBox("Enter how many Table Lamps you would like to buy", "Table Lamps")
End Sub

Private Sub cmdToasters_Click()
        Toasters = InputBox("Enter how many Toasters you would like to buy", "Toasters")
End Sub

Private Sub cmdTowels_Click()
        Towels = InputBox("Enter how many Towels you would like to buy", "Towels")
End Sub

Private Sub cmdToys_Click()
        frmHome.Hide
        frmToys.Show
End Sub

Private Sub cmdUmbrellas_Click()
        Umbrellas = InputBox("Enter how many Umbrellas you would like to buy", "Umbrellas")
End Sub

Private Sub Form_Load()
        'Target, Corp.
        'frmHome.frm
        'Mike Velin
        'March 23rd, 2009
        'To give the user options on home furnishings
End Sub
