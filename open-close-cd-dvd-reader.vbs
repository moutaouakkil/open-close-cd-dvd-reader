'!/bin/vbs
' open-close-cd-dvd-reader
' Version v1
' Coded by: Othmane Moutaouakkil [ moutaouakkil]
' Github: https://github.com/moutaouakkil



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
