Attribute VB_Name = "ModScreen"
Public Sub Refresh_Screen()
'Define to X and Y vars X1,2 and Y1,2
'X,y(1) are used for looping through the map array and getting the tile X
'X,y(2) is used to show where to draw the screen
Dim X(1 To 2) As Integer
Dim Y(1 To 2) As Integer

'Loop from Your Possition (top Left) to top left + 10 more tiles making the visable map
'10 x 10 tiles
'320 x 320 pixels
For X(1) = OffsetX To OffsetX + 10
For Y(1) = OffsetY To OffsetY + 10

'This here blits the tiles from Tileset to MapScreen
BitBlt FrmMain.PicMap.hDC, X(2) * 32, Y(2) * 32, 32, 32, FrmMain.PicTiles.hDC, MAP(X(1), Y(1)).X * 32, MAP(X(1), Y(1)).Y * 32, vbSrcCopy

'BEGIN NPC
Call BlitNPCS(X(1), X(2), Y(1), Y(2))
'END NPC

'These are just for scrolling its like a Loop without the For
'This line Scrolls to the bottem of the screen
Y(2) = Y(2) + 1
Next Y(1)
'Then it gets set to 0 (back to the top)
Y(2) = 0
'Then It moves over 1 tile and starts again
X(2) = X(2) + 1
Next X(1)

'Blits the char in the center of the screen
BitBlt FrmMain.PicMap.hDC, 5 * 32, 5 * 32, 32, 32, FrmMain.picChar.hDC, PlayerDir * 32, 0, vbSrcAnd

'Adds to the fps
FPS = FPS + 1

'Just says where the player is he is at offsetx + 5 and same for y
PlayerX = OffsetX + 5
PlayerY = OffsetY + 5

'Refresh the main screen without this you wont see anything on the display
FrmMain.PicMap.Refresh
End Sub



Sub BlitNPCS(x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer)
Dim X(1 To 2) As Integer
Dim Y(1 To 2) As Integer
X(1) = x1
X(2) = x2
Y(1) = y1
Y(2) = y2

For S = 0 To 100
If Npc(S).X > PlayerX - 5 And Npc(S).X < PlayerX + 5 And Npc(S).Y > PlayerY - 5 And Npc(S).Y < PlayerY + 5 Then
If Npc(S).X = X(1) And Npc(S).Y = Y(1) Then
If Npc(S).Visible = 1 Then
Call BitBlt(FrmMain.PicMap.hDC, X(2) * 32, Y(2) * 32, 32, 32, FrmMain.picNpcMSK.hDC, Npc(S).Direction * 32, 0, vbSrcPaint)
Call BitBlt(FrmMain.PicMap.hDC, X(2) * 32, Y(2) * 32, 32, 32, FrmMain.picNPC.hDC, Npc(S).Direction * 32, Npc(S).Sprite * 32, vbSrcAnd)
End If
End If
End If
Next S
End Sub
