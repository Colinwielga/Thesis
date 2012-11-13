VERSION 5.00
Begin VB.Form frmOrder1 
   BackColor       =   &H00FF80FF&
   Caption         =   "Lets Order"
   ClientHeight    =   8595
   ClientLeft      =   4515
   ClientTop       =   4350
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10290
   Begin VB.CommandButton cmdfinaltotal 
      BackColor       =   &H00FF8080&
      Caption         =   "Lets Get the Final Total"
      Height          =   1215
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdordercolors 
      BackColor       =   &H008080FF&
      Caption         =   "Order the Colors"
      Height          =   1335
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton cmdorderskirts 
      BackColor       =   &H00FFFF80&
      Caption         =   "Order the Skirts"
      Height          =   1215
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdorderlettering 
      BackColor       =   &H0080FFFF&
      Caption         =   "Order the Lettering"
      Height          =   1215
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdordershells 
      BackColor       =   &H0080FF80&
      Caption         =   "Order the Shells"
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Hold my poms while I stunt with your Boyfriend!"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   11.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   5880
      Width           =   5415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "All women are created equal, then a few become cheerleaders!"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   11.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   7215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "If cheerleading is so easy...why can't you pick up any girls? "
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   11.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   3120
      Width           =   6975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Any man can hold a girl's hand, but only the elite can hold her feet!"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   11.25
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4920
      Width           =   7815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "LETS ORDER A CHEERLEADING UNIFORM!!!!"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "frmOrder1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Cheerleading (Cheerleading.vbp)
'Form Name : Lets Order (Order1.frm)
'Author: Heather Johnson
'Date Written: October 28, 2003
'Purpose of Form:'this form will bring you from form to form to
                 'order yourcheerleading uniform

Private Sub cmdfinaltotal_Click()
frmShells1.Hide 'when you pick the Final Total button you can't see the shells form
frmLettering1.Hide 'when you pick the Final Total button you can't see the lettering form
frmSkirts1.Hide 'when you pick the Final Total button you can't see the skirts form
frmColors1.Hide 'when you pick the Final Total button you can't see the colors form
frmTotal1.Show 'when you pick the Final Total button you go the final total form
End Sub

Private Sub cmdordercolors_Click()
frmShells1.Hide 'when you pick the Order Colors button you can't see the shells form
frmLettering1.Hide 'when you pick the Order Colors button you can't see the lettering form
frmSkirts1.Hide 'when you pick the Order Colors button you can't see the skirts form
frmColors1.Show 'when you pick the Order Colors button you will go to the colors form
frmTotal1.Hide 'when you pick the Final Total button you can't see the final total form
End Sub

Private Sub cmdorderlettering_Click()
frmShells1.Hide 'when you pick the Order lettering button you can't see the shells form
frmLettering1.Show 'when you pick the Order lettering button you go to the lettering form
frmSkirts1.Hide 'when you pick the Order lettering button you can't see the skirts form
frmColors1.Hide 'when you pick the Order lettering button you can't see the colors form
frmTotal1.Hide 'when you pick the Final Total button you can't see the final total form
End Sub

Private Sub cmdordershells_Click()
frmShells1.Show 'when you pick the Order Shells button you go to the shells form
frmLettering1.Hide 'when you pick the Order Shells button you can't see the lettering form
frmSkirts1.Hide 'when you pick the Order Shells button you can't see the skirts form
frmColors1.Hide 'when you pick the Order Shells button you can't see the colors form
frmTotal1.Hide 'when you pick the Final Total button you can't see the final total form
End Sub

Private Sub cmdorderskirts_Click()
frmShells1.Hide 'when you pick the Order Skirts button you can't see the shells form
frmLettering1.Hide 'when you pick the Order Skirts button you can't see the lettering form
frmSkirts1.Show 'when you pick the Order Skirts button you go to the skirts form
frmColors1.Hide 'when you pick the Order Skirts button you can't see the colors form
frmTotal1.Hide 'when you pick the Final Total button you can't see the final total form
End Sub

