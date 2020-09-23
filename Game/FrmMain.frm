VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "R Quest 2 Engine"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5295
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrTXT 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   4680
      Top             =   4080
   End
   Begin VB.PictureBox picNpcMSK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   9660
      Left            =   5400
      Picture         =   "FrmMain.frx":0000
      ScaleHeight     =   640
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox picNPC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   9660
      Left            =   5400
      Picture         =   "FrmMain.frx":3C042
      ScaleHeight     =   640
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5400
      Picture         =   "FrmMain.frx":78084
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   135
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox PicTiles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9150
      Left            =   8160
      Picture         =   "FrmMain.frx":7B0C6
      ScaleHeight     =   608
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   288
      TabIndex        =   1
      Top             =   120
      Width           =   4350
   End
   Begin VB.PictureBox PicMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   0
      ScaleHeight     =   351
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   351
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.Timer tmrFollow 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   4680
         Top             =   4680
      End
      Begin VB.Label lblMSG 
         BackStyle       =   0  'Transparent
         Caption         =   "Messages popup in this label."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   4080
         Visible         =   0   'False
         Width           =   5055
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
OPENMAP 1
LoadNPCs
End Sub

Private Sub PicMap_KeyDown(KeyCode As Integer, Shift As Integer)
'If you not allowed to move dont take any input
If HALTmov = True Then Exit Sub

'Turn On Move
If KeyCode = vbKeyT Then
If tmrFollow.Enabled = False Then
tmrFollow.Enabled = True
Else
tmrFollow.Enabled = False
LoadNPCs
End If


End If

'Chat to npc's
If KeyCode = vbKeySpace Then
For a = 0 To 100
If Npc(a).X = PlayerX And Npc(a).Y = PlayerY - 1 Then
lblMSG.Caption = Npc(a).MSG
lblMSG.Visible = True
tmrTXT.Enabled = True
HALTmov = True
End If
Next a
End If


'These are pretty self explanitory if u press up Move up lol
If KeyCode = vbKeyUp Then
'Check Up block
If MAP(PlayerX, PlayerY - 1).Blocked = 1 Then Exit Sub
'Move
OffsetY = OffsetY - 1
PlayerDir = 0
End If

If KeyCode = vbKeyDown Then
'Check Down Block
If MAP(PlayerX, PlayerY + 1).Blocked = 1 Then Exit Sub
'Move
OffsetY = OffsetY + 1
PlayerDir = 1
End If

If KeyCode = vbKeyLeft Then
'Check Left Block
If MAP(PlayerX - 1, PlayerY).Blocked = 1 Then Exit Sub
'Move
OffsetX = OffsetX - 1
PlayerDir = 2
End If

If KeyCode = vbKeyRight Then
'Check Right Block
If MAP(PlayerX + 1, PlayerY).Blocked = 1 Then Exit Sub
'Move
OffsetX = OffsetX + 1
PlayerDir = 3
End If

'because -5 willset player at center (since he is 5 tiles in the middle)
If OffsetX < -5 Then OffsetX = -5
If OffsetY < -5 Then OffsetY = -5
'Same as above but this is just for the map boundry on the other side
If OffsetX > 195 Then OffsetX = 195
If OffsetY > 195 Then OffsetY = 195

'Redraw the screen
Refresh_Screen
Label1.Caption = PlayerX
Label2.Caption = PlayerY
End Sub


Private Sub tmrFollow_Timer()
Call procMoveNPC(1)
End Sub

Private Sub tmrTXT_Timer()
lblMSG.Visible = False
HALTmov = False
tmrTXT.Enabled = False
End Sub
