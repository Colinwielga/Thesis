VERSION 5.00
Begin VB.Form frmStartUp 
   BackColor       =   &H00400000&
   Caption         =   "The Beer Experience"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBar 
      BackColor       =   &H00FFFF80&
      Caption         =   "Go to the Mini Bar"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton cmdTop50 
      BackColor       =   &H00FFFF80&
      Caption         =   "View the Top 50 Beer Brewers In the Country"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton cmdMiniStore 
      BackColor       =   &H00FFFF80&
      Caption         =   "Go to the Mini Store"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdAds 
      BackColor       =   &H00FFFF80&
      Caption         =   "View Advertisements from Leading Beer Companies"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmdBAC 
      BackColor       =   &H00FFFF80&
      Caption         =   "Calculate Your Blood Alcohol Content"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdCompanies 
      BackColor       =   &H00FFFF80&
      Caption         =   "Learn More About Leading Beer Breweries"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lblNames 
      BackColor       =   &H00400000&
      Caption         =   "By: Lauren Gooley and Tim Janssen"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   735
      Left            =   2040
      TabIndex        =   7
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label lblBeer 
      BackColor       =   &H00400000&
      Caption         =   "The Beer Experience"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   1095
      Left            =   600
      TabIndex        =   6
      Top             =   480
      Width           =   7455
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the start up form, from this form the user can navigate throughout the program.
Option Explicit

Private Sub cmdAds_Click()
frmStartUp.Hide
frmAds.Show
End Sub

Private Sub cmdBAC_Click()
frmStartUp.Hide
frmBAC.Show
End Sub

Private Sub cmdBar_Click()
frmStartUp.Hide
frmBar.Show
End Sub

Private Sub cmdCompanies_Click()
frmStartUp.Hide
Companies.Show
End Sub

Private Sub cmdMiniStore_Click()
frmStartUp.Hide
frmMiniStore.Show
End Sub

Private Sub cmdTop50_Click()
frmStartUp.Hide
frmTopBeers.Show
End Sub
