VERSION 5.00
Begin VB.Form frmHowtoplay 
   BackColor       =   &H0080FF80&
   Caption         =   "How To Play"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back to Menu"
      Height          =   1095
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H0000FF00&
      Caption         =   $"frmHowtoplay.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   9015
   End
   Begin VB.Label lblHow 
      BackColor       =   &H0080FF80&
      Caption         =   "HOW TO PLAY THIS GAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "frmHowtoplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Oakley S. Puc
'Period 5
'Koonhai Zero Expeditions: Mosborn Cave

Private Sub cmdBack_Click()
frmHowtoplay.Hide
frmStart.Show

End Sub

