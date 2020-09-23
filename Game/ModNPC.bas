Attribute VB_Name = "ModNPC"
Public Sub procMoveNPC(Monno As Integer)
'Get distance to hero
d2h = Abs(Npc(Monno).X - PlayerX)

If d2h > 3 Then
DX = 1 - Int(Rnd * 3) ' random movement
DY = 1 - Int(Rnd * 3)
End If


'if your inrange then move the npc
If d2h <= 3 Then
DX = 0
If Npc(Monno).X > PlayerX Then DX = -1
If Npc(Monno).X < PlayerX Then DX = 1
DY = 0
If Npc(Monno).Y > PlayerY Then DY = -1
If Npc(Monno).Y < PlayerY Then DY = 1
End If

'Update the values
Npc(Monno).X = Npc(Monno).X + DX
Npc(Monno).Y = Npc(Monno).Y + DY

Refresh_Screen
End Sub
