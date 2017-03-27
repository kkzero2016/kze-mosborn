VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H0080FF80&
   Caption         =   "Koonhai Zero Expeditions: Mosborn Cave"
   ClientHeight    =   9540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17160
   LinkTopic       =   "Form1"
   ScaleHeight     =   9540
   ScaleWidth      =   17160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000040C0&
      Caption         =   "Exit"
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9000
      Width           =   6015
   End
   Begin VB.CommandButton cmdAbout 
      BackColor       =   &H0000FF00&
      Caption         =   "About"
      Height          =   1335
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   2535
   End
   Begin VB.CommandButton cmdInstructions 
      BackColor       =   &H0000FFFF&
      Caption         =   "How to Play"
      Height          =   1335
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H000000FF&
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Image imgTitle 
      Height          =   4455
      Left            =   3480
      Picture         =   "frmStart.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   9975
   End
   Begin VB.Label lblCreatedby 
      BackColor       =   &H0080FF80&
      Caption         =   "Created by Oakley S. Puc of kkzero Interactive"
      Height          =   255
      Left            =   13560
      TabIndex        =   4
      Top             =   9240
      Width           =   3615
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Oakley S. Puc
'Period 5
'Koonhai Zero Expeditions: Mosborn Cave

Private Sub cmdAbout_Click()
frmStart.Hide
frmAbout.Show
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdInstructions_Click()
frmStart.Hide
frmHowtoplay.Show

End Sub

Private Sub cmdStart_Click()
frmStart.Hide
frmGameplay.Show

End Sub

