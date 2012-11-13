VERSION 5.00
Begin VB.Form frmSectionView 
   BackColor       =   &H00000080&
   Caption         =   "Sections"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   735
      Left            =   4080
      TabIndex        =   3
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View Sections"
      Height          =   735
      Left            =   4080
      TabIndex        =   2
      Top             =   5760
      Width           =   1695
   End
   Begin VB.PictureBox picChart 
      Height          =   4335
      Left            =   3120
      Picture         =   "frmSectionView.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label lblView 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "View your section"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   3750
   End
End
Attribute VB_Name = "frmSectionView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmSectionView
'Cole and John
'10/30/06
'Objective: The objective of this form is to allow the user to input a section
'they wish to view.  The user can input the desired section by clicking on a
'command button that brings up the input box.

Option Explicit

Private Sub cmdBack_Click()
    frmTicketPurchase.Visible = True
    frmSectionView.Visible = False
End Sub

Private Sub cmdView_Click()
Dim Section As Integer

    Section = InputBox("Enter a Gold Section", "Section") 'whatever the user inputs now becomes the Section variable
                                                            
    Select Case Section     'compares Section variable with single values
    
    Case Is = 1
        frmSection1.Show    'if user inputs Section 1, then the view from that section will appear
        frmSectionView.Hide
    
    Case Is = 3
        frmSection3.Show
        frmSectionView.Hide
    
    Case Is = 7
        frmSection7.Show
        frmSectionView.Hide
    
    Case Is = 11
        frmSection11.Show
        frmSectionView.Hide
    
    Case Is = 13
        frmSection13.Show
        frmSectionView.Hide
    
    Case Is = 15
        frmSection15.Show
        frmSectionView.Hide
    
    Case Is = 19
        frmSection19.Show
        frmSectionView.Hide
    
    Case Is = 23
        frmSection23.Show
        frmSectionView.Hide
        
    Case Else                   'if the user inputs an invalid Section
        frmSectionNone.Show
        frmSectionView.Hide
        
    End Select

End Sub
