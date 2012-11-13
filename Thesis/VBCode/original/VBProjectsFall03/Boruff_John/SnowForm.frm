VERSION 5.00
Begin VB.Form SnowForm 
   BackColor       =   &H00C0C000&
   Caption         =   "Wax Project by John Boruff"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdOld 
      BackColor       =   &H00C0C000&
      Caption         =   "Old, Wet Corn, and Frozen Corn Snow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CommandButton cmdNewfine 
      BackColor       =   &H00C0C000&
      Caption         =   "New and Fine Snow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CommandButton cmdReturn2 
      BackColor       =   &H00C0C000&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C000&
      Caption         =   $"SnowForm.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   9
      Top             =   2760
      Width           =   10215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C000&
      Caption         =   "Ice / Frozen Corn Snow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9000
      TabIndex        =   8
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Image imgIce 
      Height          =   1950
      Left            =   8880
      Picture         =   "SnowForm.frx":008E
      Top             =   3720
      Width           =   1950
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      Caption         =   "Slush / Wet Corn Snow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      TabIndex        =   7
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Image imgSlush 
      Height          =   1950
      Left            =   6720
      Picture         =   "SnowForm.frx":2E7B
      Top             =   3720
      Width           =   1950
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "Crust / Old Snow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   6
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Image imgCrust 
      Height          =   1950
      Left            =   4560
      Picture         =   "SnowForm.frx":5A5D
      Top             =   3720
      Width           =   1950
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      Caption         =   "Crud / Fine Snow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   5
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Image imgCrud 
      Height          =   1950
      Left            =   2400
      Picture         =   "SnowForm.frx":88B1
      Top             =   3720
      Width           =   1950
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Powder / New Snow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Image imgPowder 
      Height          =   1950
      Left            =   240
      Picture         =   "SnowForm.frx":B858
      Top             =   3720
      Width           =   1950
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Please select the snow condition that you will be on"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   9135
   End
End
Attribute VB_Name = "SnowForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : WaxProject (John Boruff's VB-project.vbp)
'Form Name : SnowForm (SnowForm.frm)
'Author: John Boruff
'purpose of the form:  This form will help the user decide what
                    'type of snow he/she will be encountering
                    'it explains 5 different snow types and
                    ' groups them into two different wax conditions'
                    'the user can choose between

Private Sub cmdNewfine_Click()
    SnowForm.Hide 'brings user to the NewSnowForm
    NewSnowForm.Show
End Sub

Private Sub cmdOld_Click()
    SnowForm.Hide 'brings user to the OldSnowForm
    OldSnowForm.Show
End Sub

Private Sub cmdReturn2_Click()
    MainForm1.Show 'returns user to the MainForm1
    SnowForm.Hide
End Sub

Private Sub imgCrud_Click()  'Informs user about Crud snow type
    MsgBox "Crud or Fine snow could be considered the next step from powder. As more and more people ride throught the powder the snow gets piled at certain places and packed together at other places. This results in an unever surface with lumps of soft powder-like snow and slippery patches. Riding crud is more challenging than powder but it does not have to be less fun.", , "Crud / Fine Snow"
End Sub

Private Sub imgCrust_Click()  'informs user about Crust snow type
    MsgBox "As the names suggest crust described snow that has a harder crust on top of softer powder snow. Crust is formed as sunrays and wind melt the top layer and cold makes it freeze solid again. If the crust is hard then you will remain riding on top of the harder, icy surface. If the crust is soft you will punch through it breaking the crust with your ankles as you ride through it. Something that is less fun is an intermediate crust where you are riding on top of the crust, punch through it and then bump against a harder part again", , "Old / Crust Snow"
End Sub

Private Sub imgIce_Click()  'informs user about Ice snow type
    MsgBox "Ice is as hated as powder is loved. It is the exact opposite of powder, it is hard, it is slippery, it is hell. You will actually never find real ice on the slopes, you will only find snow that has been melted and frozen again for a number of times. This forms a solid surface of icy compact snow that is often called ice.", , "Ice / Wet Corn Snow"
End Sub

Private Sub imgPowder_Click() 'informs user about Powder snow type
    MsgBox " Powder is freshly fallen, untouched, soft snow. As the name explains, it is literally powder, tiny flakes and crystals that form a smooth and soft surface. Most snowboarders and skieers find powder the ultimate surface. It forms a soft smooth surface that will give you the feeling that you are floating in a weightless environment. Powder is often packed in thick layers that form a natural pillow for any crashes. Thick powder is the ultimate surface for trying new tricks and increasing your speed record.", , "Powder or New Snow"
End Sub

Private Sub imgSlush_Click() 'informs user about  Slush snow type
    MsgBox "Slush is snow that is starting to melt and thus becomes more wet. The snow crystals become larger and what was once soft powder turns into ice grains. People who have had a slush puppies (an icy snack) will know what the word slush means", , "Slush / Wet Corn Snow"
End Sub
