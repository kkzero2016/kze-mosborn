VERSION 5.00
Begin VB.Form frmAboutKZ 
   Caption         =   "What even IS Koonhai Zero!?"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15570
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   15570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNowstart 
      BackColor       =   &H000000FF&
      Caption         =   "Now start the game."
      Height          =   735
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   4095
   End
   Begin VB.Label lblAboutKZ 
      Caption         =   $"frmAboutKZ.frx":0000
      Height          =   4455
      Left            =   3960
      TabIndex        =   1
      Top             =   960
      Width           =   6975
   End
End
Attribute VB_Name = "frmAboutKZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Oakley S. Puc
'Period 5
'Koonhai Zero Expeditions: Mosborn Cave

Private Sub cmdNowstart_Click()
frmAboutKZ.Hide
frmGameplay.Show
End Sub
