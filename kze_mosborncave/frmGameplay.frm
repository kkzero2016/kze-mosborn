VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmGameplay 
   Caption         =   "Koonhai Zero Expeditions: Mosborn Cave"
   ClientHeight    =   10260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   ScaleHeight     =   10260
   ScaleWidth      =   15225
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAnim2 
      Interval        =   100
      Left            =   6960
      Top             =   840
   End
   Begin VB.Timer tmrAnim 
      Interval        =   100
      Left            =   4920
      Top             =   360
   End
   Begin VB.Timer tmrBattle 
      Left            =   8760
      Top             =   480
   End
   Begin VB.Frame fraBattleoptions 
      BackColor       =   &H00004080&
      Caption         =   "Battle Options"
      Height          =   4935
      Left            =   10560
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton cmdEngulf 
         BackColor       =   &H000000FF&
         Caption         =   "Fiery Engulf (5 PP)"
         Height          =   1695
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdSlash 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Energy Slash (4 PP)"
         Height          =   1695
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdBash 
         BackColor       =   &H0000FF00&
         Caption         =   "Sparky Bash (3 PP)"
         Height          =   1695
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdRegular 
         Caption         =   "Regular Attack (No PP Cost)"
         Height          =   1095
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.CommandButton cmdRevivewater 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Revival Water"
         Height          =   855
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdHealwater 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Healing Water"
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdBatoptback 
         BackColor       =   &H000000FF&
         Caption         =   "BACK"
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4440
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdFlee 
         BackColor       =   &H00FF0000&
         Caption         =   "FLEE"
         BeginProperty Font 
            Name            =   "Prestige Elite Std"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3360
         Width           =   3735
      End
      Begin VB.CommandButton cmdItems 
         BackColor       =   &H0000FFFF&
         Caption         =   "ITEMS"
         BeginProperty Font 
            Name            =   "Prestige Elite Std"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1920
         Width           =   3735
      End
      Begin VB.CommandButton cmdAttack 
         BackColor       =   &H000000C0&
         Caption         =   "ATTACK"
         BeginProperty Font 
            Name            =   "Prestige Elite Std"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.Frame fraFlamestats 
      BackColor       =   &H000000FF&
      Caption         =   "Flame"
      Height          =   975
      Left            =   12360
      TabIndex        =   5
      Top             =   9120
      Width           =   1935
      Begin VB.Label lblFlamepp 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblFlamehp 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblFlamepplabel 
         BackColor       =   &H000000FF&
         Caption         =   "PP"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblFlamehplabel 
         BackColor       =   &H000000FF&
         Caption         =   "HP"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame fraSetstats 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Set"
      Height          =   975
      Left            =   10320
      TabIndex        =   4
      Top             =   9120
      Width           =   1935
      Begin VB.Label lblSetpp 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblSethp 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblSetpplabel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "PP"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblSethplabel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "HP"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame fraJohnstats 
      BackColor       =   &H0000FF00&
      Caption         =   "John"
      Height          =   975
      Left            =   8280
      TabIndex        =   3
      Top             =   9120
      Width           =   1935
      Begin VB.Label lblJohnpp 
         BackColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblJohnhp 
         BackColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblJohnpplabel 
         BackColor       =   &H0000FF00&
         Caption         =   "PP"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblJohnhplabel 
         BackColor       =   &H0000FF00&
         Caption         =   "HP"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame fraEnemy 
      BackColor       =   &H00000040&
      Caption         =   "Enemy Name"
      Height          =   6135
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Frame fraEnemystats 
         BackColor       =   &H000080FF&
         Caption         =   "Enemy Stats"
         Height          =   975
         Left            =   4800
         TabIndex        =   6
         Top             =   5160
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Label lblEnemypp 
            BackColor       =   &H000080FF&
            Height          =   255
            Left            =   480
            TabIndex        =   30
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblEnemyhp 
            BackColor       =   &H000080FF&
            Height          =   255
            Left            =   480
            TabIndex        =   29
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblEnemyhplbl 
            BackColor       =   &H000080FF&
            Caption         =   "HP"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Image imgBeater 
         Height          =   5055
         Index           =   2
         Left            =   240
         Picture         =   "frmGameplay.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Image imgBeater 
         Height          =   5055
         Index           =   1
         Left            =   360
         Picture         =   "frmGameplay.frx":C0042
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Image imgBeater 
         Height          =   5055
         Index           =   0
         Left            =   240
         Picture         =   "frmGameplay.frx":180084
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Image imgSwordsman 
         Height          =   4800
         Index           =   2
         Left            =   360
         Picture         =   "frmGameplay.frx":2400C6
         Stretch         =   -1  'True
         Top             =   480
         Visible         =   0   'False
         Width           =   4920
      End
      Begin VB.Image imgSwordsman 
         Height          =   4800
         Index           =   1
         Left            =   360
         Picture         =   "frmGameplay.frx":300108
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   4920
      End
      Begin VB.Image imgSwordsman 
         Height          =   4800
         Index           =   0
         Left            =   360
         Picture         =   "frmGameplay.frx":3C014A
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   4920
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpWDH 
      Height          =   735
      Left            =   360
      TabIndex        =   34
      Top             =   9240
      Visible         =   0   'False
      Width           =   1215
      URL             =   "H:\VB6 Stuff\finalproject\music\wdh.mp3"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2143
      _cy             =   1296
   End
   Begin VB.Label lblEnemyrand 
      Caption         =   "Label1"
      Height          =   375
      Left            =   11280
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgPlayerRight 
      Height          =   975
      Left            =   7320
      Picture         =   "frmGameplay.frx":48018C
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   975
   End
   Begin VB.Image imgPlayerLeft 
      Height          =   975
      Left            =   7320
      Picture         =   "frmGameplay.frx":48148C
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   975
   End
   Begin VB.Image imgPlayerFront 
      Height          =   975
      Left            =   7320
      Picture         =   "frmGameplay.frx":4825FF
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   975
   End
   Begin VB.Image imgPlayerBack 
      Height          =   975
      Left            =   7320
      Picture         =   "frmGameplay.frx":483C92
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblLeft 
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblTop 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgPlayer 
      Height          =   975
      Left            =   7320
      Picture         =   "frmGameplay.frx":4856B5
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   975
   End
   Begin VB.Image imgBGSTART 
      Height          =   11520
      Left            =   -120
      Picture         =   "frmGameplay.frx":485B3B
      Top             =   -240
      Width           =   16110
   End
End
Attribute VB_Name = "frmGameplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Oakley S. Puc
'Period 5
'Koonhai Zero Expeditions: Mosborn Cave

Option Explicit
'There are plenty of variables following.

Dim canmove As Integer 'Determines if the player can move on the overworld or not. For example, if someone is talking in the game, this will set to zero and the player canot move.
Dim battle As Integer 'Determines if a battle is underway or not. 0 if not and 1 if so.

'The overworld coordinates for the player
Dim playerx As Integer
Dim playery As Integer

'John's battle stats
Dim johnhp As Integer
Dim johnpp As Integer
Dim johnatk As Integer
Dim johndef As Integer

'Set's battle stats
Dim sethp As Integer
Dim setpp As Integer
Dim setatk As Integer
Dim setdef As Integer

'Flame's battle stats
Dim flamehp As Integer
Dim flamepp As Integer
Dim flameatk As Integer
Dim flamedef As Integer

'In-battle enemy stats
Dim enemyhp As Integer
Dim enemyatk As Integer
Dim enemydef As Integer
Dim enemy As Integer

Dim hpcost As Integer 'To determine how much damage to be inflicted on the target

'Items
'These variables determine how many of each item the player has
Dim healwater As Integer
Dim revivewater As Integer

Dim i As Integer 'For timers

Dim randomizer As Integer 'This randomizes several things, like if a wild battle will occur after every step, which enemy will appear, if the player can flee when they clcik flee, etc.
Dim enemyrand As Integer 'To determine enemy

Private Sub cmdAttack_Click()
'Hides every option not used on the attack menu
cmdAttack.Visible = False
cmdItems.Visible = False
cmdFlee.Visible = False
cmdBatoptback.Visible = True
cmdRegular.Visible = True
cmdBash.Visible = True
cmdSlash.Visible = True
cmdEngulf.Visible = True
End Sub

Private Sub cmdBash_Click()
'When the attack starts, it checks if John has enough powerpoints to use the attack.
If johnpp > 2 And johnhp > 0 Then
    cmdBatoptback.Visible = False
    cmdRegular.Visible = False
    cmdBash.Visible = False
    cmdSlash.Visible = False
    cmdEngulf.Visible = False
    MsgBox ("John starts leaking sparks. He rams into the foe!"), vbOKOnly
    hpcost = 10 - enemydef 'Calculates how much HP will be taken from the target. It is stored into a variable in order to be processed in the next operation
    If hpcost < 1 Then
        hpcost = 1 'If the HP cost is less than 1, it will be adjusted to 1. Imagine a violent assault that healed someone. Makes no sense, right?
    End If
    enemyhp = enemyhp - hpcost 'The enemy HP is now lowered
    lblEnemyhp.Caption = enemyhp
    'The below checks if the enemy's HP is lower than 1. If so, the enemy has lost the battle, and the player is directed back to the overworld.
    If enemyhp < 1 Then
        battle = 0
        canmove = 1
        fraBattleoptions.Visible = False
        fraEnemy.Visible = False
    ElseIf enemyhp > 0 Then 'But if theyu're still fighting fit, they will attack.
        MsgBox ("The enemy attacks!")
        'Whoever attacked the last turn will have their HP lowered. In this case, John takes the hit.
        johnhp = johnhp - enemyatk + johndef
        lblJohnhp.Caption = johnhp
    End If
    
    'All options and labels on the base menu will regenerate after the enemy's attack
    fraBattleoptions.Visible = True
    cmdAttack.Visible = True
    cmdItems.Visible = True
    cmdFlee.Visible = True
    johnpp = johnpp - 3
    lblJohnpp.Caption = johnpp
End If
'If every teammate is less than 1 HP, the game is over.
If johnhp < 1 And sethp < 1 And flamehp < 1 Then
    frmGameplay.Hide
    frmGameover.Show
End If


End Sub

Private Sub cmdBatoptback_Click()
'Clicking back on any menu that has it will direct you back to the base battle menu
cmdAttack.Visible = True
cmdItems.Visible = True
cmdFlee.Visible = True
cmdBatoptback.Visible = False
cmdRevivewater.Visible = False
cmdHealwater.Visible = False
cmdRegular.Visible = False
cmdBash.Visible = False
cmdSlash.Visible = False
cmdEngulf.Visible = False
End Sub

Private Sub cmdEngulf_Click()
'Same as the first attack, except with Falme.
If flamepp > 2 And flamehp > 0 Then
    cmdBatoptback.Visible = False
    cmdRegular.Visible = False
    cmdBash.Visible = False
    cmdSlash.Visible = False
    cmdEngulf.Visible = False
    MsgBox ("Flame soars up into a great big flame and hurts the enemy!"), vbOKOnly
    hpcost = 15 - enemydef
    If hpcost < 1 Then
        hpcost = 1
    End If
    enemyhp = enemyhp - hpcost
    lblEnemyhp.Caption = enemyhp
    
    If enemyhp < 1 Then
        battle = 0
        canmove = 1
        fraBattleoptions.Visible = False
        fraEnemy.Visible = False
    ElseIf enemyhp > 0 Then
        MsgBox ("The enemy attacks!")
        flamehp = flamehp - enemyatk + flamedef
        lblFlamehp.Caption = flamehp
        End If
    End If
    

    fraBattleoptions.Visible = True
    cmdAttack.Visible = True
    cmdItems.Visible = True
    cmdFlee.Visible = True
    flamepp = flamepp - 3
    lblFlamepp.Caption = flamepp

If johnhp < 1 And sethp < 1 And flamehp < 1 Then
    frmGameplay.Hide
    frmGameover.Show
End If
End Sub

Private Sub cmdFlee_Click()
'You have a chance of fleeing with this option.
randomizer = Int((Rnd * 99) + 1)
If randomizer < 25 Then 'The escape will fail if the randomizer hits below 25
    MsgBox ("Escape failed!"), vbOKOnly
    
    fraBattleoptions.Visible = False
    cmdAttack.Visible = False
    cmdItems.Visible = False
    cmdFlee.Visible = False
    
    
    
    'The enemy attacks, just like after you attacked them
    MsgBox ("The enemy attacks!")
    randomizer = Int((Rnd * 2) + 1)
    johnhp = johnhp - enemyatk + johndef
    sethp = sethp - enemyatk + setdef
    flamehp = flamehp - enemyatk + flamedef
    lblJohnhp.Caption = johnhp
    lblSethp.Caption = sethp
    lblFlamehp.Caption = flamehp
    'Once again, the game is over once the whole team is down
    If johnhp < 1 And sethp < 1 And flamehp < 1 Then
        frmGameplay.Hide
        frmGameover.Show
    End If

fraBattleoptions.Visible = True
cmdAttack.Visible = True
cmdItems.Visible = True
cmdFlee.Visible = True
Else 'If the randomizer is above 24, the battle will be escaped, and you will be directed back to the overworld.
    MsgBox ("The battle has been terminated."), vbOKOnly
    fraBattleoptions.Visible = False
    cmdAttack.Visible = False
    cmdItems.Visible = False
    cmdFlee.Visible = False
    fraEnemy.Visible = False
    fraEnemystats.Visible = False
    canmove = 1
    battle = 0
End If


    
End Sub

Private Sub cmdHealwater_Click()
'This item replenishes some health of the teammates.
If healwater > 0 Then 'Detects if you still have any haling waters
    MsgBox ("The group intakes the healing water. They all heal up a bit."), vbOKOnly
    If johnhp > 0 Then 'Keep in mind you can't heal more if the teammate is down.
        johnhp = johnhp + 7
    End If
    If sethp > 0 Then
        sethp = sethp + 7
    End If
    If flamehp > 0 Then
        flamehp = flamehp + 7
    End If
    healwater = healwater - 1
End If
lblJohnhp.Caption = johnhp
lblSethp.Caption = sethp
lblFlamehp.Caption = flamehp

cmdHealwater.Visible = False
cmdRevivewater.Visible = False
cmdBatoptback.Visible = False

MsgBox ("The enemy attacks!")
johnhp = johnhp - enemyatk + johndef
sethp = sethp - enemyatk + setdef
flamehp = flamehp - enemyatk + flamedef
lblJohnhp.Caption = johnhp
lblSethp.Caption = sethp
lblFlamehp.Caption = flamehp

fraBattleoptions.Visible = True
cmdAttack.Visible = True
cmdItems.Visible = True
cmdFlee.Visible = True

If johnhp < 1 And sethp < 1 And flamehp < 1 Then
    frmGameplay.Hide
    frmGameover.Show
End If

End Sub

Private Sub cmdItems_Click()
'Gets everything for the item options showing
cmdAttack.Visible = False
cmdItems.Visible = False
cmdFlee.Visible = False
cmdBatoptback.Visible = True
cmdHealwater.Visible = True
cmdRevivewater.Visible = True
End Sub

Private Sub cmdRegular_Click()
'A regular attack every team mate uses to simply bash the foe. Everyone gets damaged after this.
cmdBatoptback.Visible = False
cmdRegular.Visible = False
cmdBash.Visible = False
cmdSlash.Visible = False
cmdEngulf.Visible = False
MsgBox ("The group scrambles toward the foe, inflicting damage on them."), vbOKOnly
hpcost = 6 - enemydef
If hpcost < 1 Then
    hpcost = 1
End If
enemyhp = enemyhp - hpcost
lblEnemyhp.Caption = enemyhp


If enemyhp < 1 Then
        battle = 0
        canmove = 1
        fraBattleoptions.Visible = False
        fraEnemy.Visible = False
ElseIf enemyhp > 0 Then
        MsgBox ("The enemy attacks!")
        johnhp = johnhp - enemyatk + johndef
        sethp = sethp - enemyatk + setdef
        flamehp = flamehp - enemyatk + flamedef
        lblJohnhp.Caption = johnhp
        lblSethp.Caption = sethp
        lblFlamehp.Caption = flamehp
End If

fraBattleoptions.Visible = True
cmdAttack.Visible = True
cmdItems.Visible = True
cmdFlee.Visible = True

If johnhp < 1 And sethp < 1 And flamehp < 1 Then
    frmGameplay.Hide
    frmGameover.Show
End If


End Sub

Private Sub cmdRevivewater_Click()
'This revives every team member if they're down
If revivewater > 0 Then
    MsgBox ("The group intakes the revivial water. They all get back up and running."), vbOKOnly
    johnhp = 20
    sethp = 15
    flamehp = 17
    revivewater = revivewater - 1
End If



cmdHealwater.Visible = False
cmdRevivewater.Visible = False
cmdBatoptback.Visible = False

MsgBox ("The enemy attacks!")

johnhp = johnhp - enemyatk + johndef
sethp = sethp - enemyatk + setdef
flamehp = flamehp - enemyatk + flamedef
lblJohnhp.Caption = johnhp
lblSethp.Caption = sethp
lblFlamehp.Caption = flamehp
fraBattleoptions.Visible = True
cmdAttack.Visible = True
cmdItems.Visible = True
cmdFlee.Visible = True

If johnhp < 1 And sethp < 1 And flamehp < 1 Then
    frmGameplay.Hide
    frmGameover.Show
End If
End Sub

Private Sub cmdSlash_Click()
'Set's cool knife attack
If setpp > 3 And sethp > 0 Then
    cmdBatoptback.Visible = False
    cmdRegular.Visible = False
    cmdBash.Visible = False
    cmdSlash.Visible = False
    cmdEngulf.Visible = False
    MsgBox ("Set charges his knife and strikes the foe!"), vbOKOnly
    hpcost = 10 - enemydef
    If hpcost < 1 Then
        hpcost = 1
    End If
    enemyhp = enemyhp - hpcost
    lblEnemyhp.Caption = enemyhp
End If
    
    
    
    If enemyhp < 1 Then
        battle = 0
        canmove = 1
        fraBattleoptions.Visible = False
        fraEnemy.Visible = False
    ElseIf enemyhp > 0 Then
        MsgBox ("The enemy attacks!")
        randomizer = Int((Rnd * 2) + 1)
        sethp = sethp - enemyatk + setdef
        lblSethp.Caption = sethp
        
    End If
    

    fraBattleoptions.Visible = True
    cmdAttack.Visible = True
    cmdItems.Visible = True
    cmdFlee.Visible = True
    setpp = setpp - 3
    lblSetpp.Caption = setpp
    
    If johnhp < 1 And sethp < 1 And flamehp < 1 Then
        frmGameplay.Hide
        frmGameover.Show
    
End If
End Sub

Private Sub Form_KeyDown(Keycode As Integer, Shift As Integer)
'Here's the movement of the player
If canmove = 1 Then
    If Keycode = vbKeyUp Then
        'the player moves
        playery = playery - 300
        imgPlayerFront.Visible = False
        imgPlayerLeft.Visible = False
        imgPlayerRight.Visible = False
        imgPlayerBack.Visible = True
        imgPlayerBack.Top = playery
        imgPlayerBack.Left = playerx
        'Collision detection
        If Top < 0 Then
            playery = playery + 300
            imgPlayerBack.Top = playery
            imgPlayerBack.Left = playerx
        End If
        lblTop.Caption = playery
        'Checks if a wild enemy appears and attacks
        randomizer = Int((Rnd * 99) + 1)
        If randomizer < 3 Then 'If the randomizer goes below 3, a battle will commence. I changed this from 10 on the write-up to 3.
            canmove = 0
            battle = 1
            
            fraEnemy.Visible = True
            fraEnemystats.Visible = True
            fraBattleoptions.Visible = True
            cmdAttack.Visible = True
            cmdItems.Visible = True
            cmdFlee.Visible = True
            'All visual elements are ready for the battle
            enemyrand = Int((Rnd * 99) + 1)
            Randomize
            lblEnemyrand.Caption = randomizer
            enemyrand = Int((Rnd * 5) + 1)
            lblEnemyrand.Caption = randomizer
            Randomize
            'Determines which enemy will appear based on the randomizer
            If enemyrand > 2 Then
                enemy = 1
                enemyhp = 10
                enemyatk = 7
                enemydef = 1
                fraEnemy.Caption = "Swordsman"
                
            Else
                enemy = 2
                enemyhp = 14
                enemyatk = 5
                enemydef = 4
                fraEnemy.Caption = "Beater"
            End If
            lblEnemyhp.Caption = enemyhp
        End If
    End If
    
    If Keycode = vbKeyDown Then
        'the player moves
        playery = playery + 300
        imgPlayerFront.Visible = True
        imgPlayerLeft.Visible = False
        imgPlayerRight.Visible = False
        imgPlayerBack.Visible = False
        imgPlayerFront.Top = playery
        imgPlayerFront.Left = playerx
        'Collision detection
        If Top > 9240 Then
            playery = playery - 300
            imgPlayerFront.Top = playery
            imgPlayerFront.Left = playerx
        End If
         lblTop.Caption = playery
        'Checks if a wild enemy appears and attacks
        randomizer = Int((Rnd * 99) + 1)
        If randomizer < 3 Then
            canmove = 0
            battle = 1
            
            fraEnemy.Visible = True
            fraEnemystats.Visible = True
            fraBattleoptions.Visible = True
            cmdAttack.Visible = True
            cmdItems.Visible = True
            cmdFlee.Visible = True
            enemyrand = Int((Rnd * 5) + 1)
            lblEnemyrand.Caption = randomizer
            Randomize
            If enemyrand > 2 Then
                enemy = 1
                enemyhp = 10
                enemyatk = 7
                enemydef = 1
                fraEnemy.Caption = "Swordsman"
            Else
                enemy = 2
                enemyhp = 14
                enemyatk = 5
                enemydef = 4
                fraEnemy.Caption = "Beater"
            End If
            lblEnemyhp.Caption = enemyhp
        End If
    End If
    
    If Keycode = vbKeyRight Then
        'the player moves
        playerx = playerx + 300
        imgPlayerRight.Left = playerx
        imgPlayerRight.Top = playery
        imgPlayerFront.Visible = False
        imgPlayerLeft.Visible = False
        imgPlayerRight.Visible = True
        imgPlayerBack.Visible = False
        'Collision detection
        If Left < 0 Then
            playerx = playerx - 300
            imgPlayerRight.Left = playerx
            imgPlayerRight.Top = playery
        End If
         lblLeft.Caption = playerx
        'Checks if a wild enemy appears and attacks
        randomizer = Int((Rnd * 99) + 1)
        If randomizer < 3 Then
            canmove = 0
            battle = 1
            
            fraEnemy.Visible = True
            fraEnemystats.Visible = True
            fraBattleoptions.Visible = True
            cmdAttack.Visible = True
            cmdItems.Visible = True
            cmdFlee.Visible = True
            enemyrand = Int((Rnd * 5) + 1)
            lblEnemyrand.Caption = randomizer
            Randomize
            If enemyrand > 2 Then
                enemy = 1
                enemyhp = 10
                enemyatk = 7
                enemydef = 1
                fraEnemy.Caption = "Swordsman"
            Else
                enemy = 2
                enemyhp = 14
                enemyatk = 5
                enemydef = 4
                fraEnemy.Caption = "Beater"
            End If
            lblEnemyhp.Caption = enemyhp
        End If
    End If
    
    
    If Keycode = vbKeyLeft Then
        'the player moves
        playerx = playerx - 300
        imgPlayerFront.Visible = False
        imgPlayerLeft.Visible = True
        imgPlayerRight.Visible = False
        imgPlayerBack.Visible = False
        imgPlayerLeft.Left = playerx
        imgPlayerLeft.Top = playery
        'Collision detection
         If Left > 14280 Then
            playerx = playerx + 300
            imgPlayerLeft.Left = playerx
            imgPlayerLeft.Top = playery
        End If
        lblLeft.Caption = playerx
        'Checks if a wild enemy appears and attacks
        randomizer = Int((Rnd * 99) + 1)
        If randomizer < 3 Then
            canmove = 0
            battle = 1
            
            fraEnemy.Visible = True
            fraEnemystats.Visible = True
            fraBattleoptions.Visible = True
            cmdAttack.Visible = True
            cmdItems.Visible = True
            cmdFlee.Visible = True
            enemyrand = Int((Rnd * 5) + 1)
            lblEnemyrand.Caption = randomizer
            Randomize
            If enemyrand > 2 Then
                enemy = 1
                enemyhp = 10
                enemyatk = 7
                enemydef = 1
                fraEnemy.Caption = "Swordsman"
            Else
                enemy = 2
                enemyhp = 14
                enemyatk = 5
                enemydef = 4
                fraEnemy.Caption = "Beater"
            End If
            lblEnemyhp.Caption = enemyhp
        End If
    End If
    
End If
End Sub

Private Sub Form_Load()
MsgBox ("Get out there and stay strong!"), vbOKOnly 'Starts with this




i = 0 'Gets timers' stuff ready

'You can move at the start
canmove = 1
battle = 0

'Items are stocked as follows
healwater = 3
revivewater = 1

'Stats are set

'John's
johnhp = 20
johnpp = 15

'Set's
sethp = 15
setpp = 20

'Flame's
flamehp = 17
flamepp = 18

'Labels have stats captioned to them
lblJohnhp.Caption = johnhp
lblJohnpp.Caption = johnpp
lblSetpp.Caption = setpp
lblSethp.Caption = sethp
lblFlamehp.Caption = flamehp
lblFlamepp.Caption = flamepp

'Player coords are set
playerx = 7200
playery = 5280
imgPlayer.Visible = False
imgPlayerFront.Visible = False
imgPlayerLeft.Visible = False
imgPlayerRight.Visible = False
imgPlayerBack.Visible = True
imgPlayerBack.Top = playery
imgPlayerBack.Left = playerx



End Sub

Private Sub tmrAnim_Timer()
Reset

If battle = 1 And enemy = 1 Then
    imgBeater(0).Visible = False
    imgBeater(1).Visible = False
    imgBeater(2).Visible = False
    imgSwordsman(i).Visible = False
        If i = 0 Then 'If the animation is at its end, i will shift back to zero.
            i = 2
        Else
            i = i - 1
        End If
    imgSwordsman(i).Visible = True
End If

If battle = 0 Then
    i = 0
End If

End Sub

Private Sub tmrAnim2_Timer()
If battle = 1 And enemy = 2 Then
        imgSwordsman(0).Visible = False
        imgSwordsman(1).Visible = False
        imgSwordsman(2).Visible = False
        If i = 0 Then 'If the animation is at its end, i will shift back to zero.
            i = 2
        Else
            i = i - 1
        End If
    imgBeater(i).Visible = True
End If

If battle = 0 Then
    i = 0
    fraBattleoptions.Visible = False
    canmove = 1
End If

End Sub

Private Sub tmrBattle_Timer()
lblEnemyrand.Caption = randomizer
If battle = 1 Then
    'HP and PP Display
    lblEnemyhplbl.Visible = True
    
    lblEnemyhp.Caption = enemyhp
    
    lblJohnhp.Caption = johnhp
    lblJohnpp.Caption = johnpp
    lblSethp.Caption = sethp
    lblSetpp.Caption = setpp
    lblFlamehp.Caption = flamehp
    lblFlamepp.Caption = flamepp
End If
    
'Enemy specific stats
If enemy = "Swordsman" Then
    enemyatk = 7
    enemydef = 2
    lblEnemyhp.Caption = enemyhp
End If
End Sub

