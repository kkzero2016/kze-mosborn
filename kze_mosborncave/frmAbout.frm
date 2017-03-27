VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H0080FF80&
   Caption         =   "About Koonhai Zero Expeditions: Mosborn Cave"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15600
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   15600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWhat 
      BackColor       =   &H0000FF00&
      Caption         =   "But just WHAT is this whole ""Koonhai Zero"" thing!? "
      Height          =   1575
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdBackfromabout 
      BackColor       =   &H000000FF&
      Caption         =   "Back"
      Height          =   855
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Image imgOriginal 
      Height          =   2775
      Left            =   6000
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label lblAudiovisual 
      BackColor       =   &H0080FF80&
      Caption         =   $"frmAbout.frx":0D25
      Height          =   735
      Left            =   5160
      TabIndex        =   4
      Top             =   5280
      Width           =   5175
   End
   Begin VB.Label lblCopyright 
      BackColor       =   &H0080FF80&
      Caption         =   "(c) 2016-2017 kkzero Interactive"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label lblCodelicense 
      BackColor       =   &H0080FF80&
      Caption         =   "The program in source and binary form is distributed under absolutely no license. The program itself is Public Domain."
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   4560
      Width           =   5055
   End
   Begin VB.Label lblTitleabout 
      BackColor       =   &H0080FF80&
      Caption         =   "KOONHAI ZERO EXPEDITIONS: MOSBORN CAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Oakley S. Puc
'Period 5
'Koonhai Zero Expeditions: Mosborn Cave

Private Sub cmdBackfromabout_Click()
frmAbout.Hide
frmStart.Show
End Sub

Private Sub cmdWhat_Click()
frmAbout.Hide
frmAboutKZ.Show
End Sub
