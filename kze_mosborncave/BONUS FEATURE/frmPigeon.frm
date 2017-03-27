VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19110
   LinkTopic       =   "Form1"
   ScaleHeight     =   11145
   ScaleWidth      =   19110
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgBrickwall 
      Height          =   9975
      Left            =   13680
      Picture         =   "frmPigeon.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label lblTitle 
      Caption         =   "PIGEON"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7920
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Image imgPigeon 
      Height          =   3855
      Left            =   840
      Picture         =   "frmPigeon.frx":1A7CA
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const SND_APPLICATION = &H80         '  look for application specific association
Private Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_FILENAME = &H20000     '  name is a file name
Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Private Const SND_PURGE = &H40               '  purge non-static events for task
Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Private Const SND_SYNC = &H0         '  play synchronously (default)
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Sub Form_Load()
    PlaySound ".\pigeon_theme_song.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
    If imgPigeon.Left > 11999 Then
        MsgBox "Crashed into a brick wall", vbOKOnly
        End
    End If
End Sub

Private Sub Form_KeyDown(Keycode As Integer, Shift As Integer)
If Keycode = vbKeyUp Then
    imgPigeon.Top = imgPigeon.Top - 300
End If
If Keycode = vbKeyDown Then
    imgPigeon.Top = imgPigeon.Top + 300
End If
If Keycode = vbKeyRight Then
    imgPigeon.Left = imgPigeon.Left + 300
End If
If Keycode = vbKeyLeft Then
    imgPigeon.Left = imgPigeon.Left - 300
End If
If imgPigeon.Left > 11999 Then
    MsgBox "Crashed into a brick wall", vbOKOnly
    End
End If

End Sub

