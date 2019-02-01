'!/bin/vbs
' Open and Close CD-DVD Deader Randomly
' Version v1
' Coded by: Othmane Moutaouakkil [ WHOAMI2507 ]  (You don't become a coder by just changing the credits)
' Github: https://github.com/whoami2507



Set oWMP = CreateObject("WMPlayer.OCX.7")
Set colCDROMs = oWMP.cdromCollection
do
if colCDROMs.Count >= 1 then
For i = 0 to colCDROMs.Count - 1
colCDROMs.Item(i).Eject
Next
For i = 0 to colCDROMs.Count - 1
colCDROMs.Item(i).Eject
Next
End If
wscript.sleep 5000
loop
