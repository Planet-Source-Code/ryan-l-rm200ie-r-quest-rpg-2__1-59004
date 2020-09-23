VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "R Quest 2 Map Editor"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Fill Map With Sel"
      Height          =   255
      Left            =   5760
      TabIndex        =   19
      Top             =   5400
      Width           =   1935
   End
   Begin VB.PictureBox PicColl 
      AutoRedraw      =   -1  'True
      Height          =   615
      Left            =   5280
      Picture         =   "FrmMain.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   18
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mini Map"
      Height          =   1455
      Left            =   7800
      TabIndex        =   16
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Events"
      Height          =   1095
      Left            =   7800
      TabIndex        =   15
      Top             =   3000
      Width           =   2655
      Begin VB.CheckBox ChkBlock 
         Caption         =   "Blocked"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load"
      Height          =   375
      Left            =   7800
      TabIndex        =   14
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   7800
      TabIndex        =   13
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Timer tmrFPS 
      Interval        =   1000
      Left            =   5880
      Top             =   3960
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   190
      TabIndex        =   7
      Top             =   5400
      Width           =   5295
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   5295
      Left            =   5400
      Max             =   190
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox PicSel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6960
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1935
      Left            =   10200
      Max             =   100
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   5760
      ScaleHeight     =   1875
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.PictureBox PicTiles 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   9150
         Left            =   0
         Picture         =   "FrmMain.frx":0C42
         ScaleHeight     =   608
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   288
         TabIndex        =   2
         Top             =   0
         Width           =   4350
      End
   End
   Begin VB.PictureBox PicMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   120
      ScaleHeight     =   351
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   351
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.Shape shpMouse 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.Label lblFPS 
      Caption         =   "0"
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "FPS Count:"
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label oY 
      Caption         =   "0"
      Height          =   255
      Left            =   6480
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label oX 
      Caption         =   "0"
      Height          =   255
      Left            =   6480
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Offsets:"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Selected Tile :"
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'These are temp holders for the selcted tile
Dim SelX, SelY As Integer

Private Sub Command1_Click()
num = InputBox("Map Index for save")
Open App.Path & "/Maps/" & num & ".map" For Binary As #1
Put #1, , Map
Close #1
End Sub

Private Sub Command2_Click()
num = InputBox("Map Index for save")
Open App.Path & "/Maps/" & num & ".map" For Binary As #1
Get #1, , Map
Close #1
Refresh_Screen
End Sub

Private Sub Command3_Click()
For a = 0 To 200
For b = 0 To 200
Map(a, b).X = SelX
Map(a, b).Y = SelY
Next b
Next a
End Sub

Private Sub HScroll1_Change()
OffsetX = HScroll1.Value
Refresh_Screen
End Sub

Private Sub PicMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'if an error occurs it will refresh the screen
'an example of this is the users mouse goes off the screen the engine will crash
'because X = like -1000 and the variable cant go below 0
On Error GoTo Err:
Dim X2, Y2 As Integer

X2 = Int(X / 32) ' This just makes each 32 pixels 1 number so 32,32 = 1,1 64,64 = 2,2
Y2 = Int(Y / 32)

X2 = X2 + OffsetX
Y2 = Y2 + OffsetY

If X2 < 0 Then X2 = 0
If Y2 < 0 Then Y2 = 0
If X2 > 100 Then X2 = 200
If Y2 > 100 Then Y2 = 200


' if the button = 1 then
' Set the map array
If Button = 1 Then
Map(X2, Y2).X = SelX 'Put the sel tile into the map location
Map(X2, Y2).Y = SelY
Map(X2, Y2).Blocked = 0
    'if the blocked checkbox is checked then set the block
    If ChkBlock.Value = 1 Then
    Map(X2, Y2).Blocked = 1
    End If
End If

Refresh_Screen 'Redraw the game
Err:
Refresh_Screen
End Sub

Private Sub PicMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim X2, Y2 As Integer

X2 = Int(X / 32) ' This just makes each 32 pixels 1 number so 32,32 = 1,1 64,64 = 2,2
Y2 = Int(Y / 32)

shpMouse.Top = Y2 * 32
shpMouse.Left = X2 * 32

'if the left mouse button is pressed
'CLICK at that spot
If Button = 1 Then
Call PicMap_MouseDown(Button, Shift, X, Y)
End If
End Sub

Private Sub PicTiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim X2, Y2 As Integer

X2 = Int(X / 32) ' This just makes each 32 pixels 1 number so 32,32 = 1,1 64,64 = 2,2
Y2 = Int(Y / 32)

SelX = X2 ' Put where you clicked into sel X and Y
SelY = Y2
'Draw the selected tile to the PicSel
BitBlt PicSel.hDC, 0, 0, 32, 32, PicTiles.hDC, X2 * 32, Y2 * 32, vbSrcCopy
'Refresh the Selected Picture Box
PicSel.Refresh
End Sub

Private Sub tmrFPS_Timer()
lblFPS.Caption = FPS
FPS = 0
End Sub



Private Sub VScroll1_Change()
'This just moves the tileset up and down
PicTiles.Top = -VScroll1.Value * 490
End Sub

Private Sub VScroll2_Change()
OffsetY = VScroll2.Value
Refresh_Screen
End Sub
