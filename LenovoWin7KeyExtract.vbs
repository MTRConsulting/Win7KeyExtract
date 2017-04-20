Set WshShell = CreateObject("WScript.Shell")
'MsgBox ConvertToKey(WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId"))

Dim KeyArray(165)
KeyArray(0) = &Ha4
KeyArray(1) = &H00
KeyArray(2) = &H00 
KeyArray(3) = &H00 
KeyArray(4) = &H03 
KeyArray(5) = &H00 
KeyArray(6) = &H00 
KeyArray(7) = &H00 
KeyArray(8) = &H30 
KeyArray(9) = &H30 
KeyArray(10) = &H33 
KeyArray(11) = &H37 
KeyArray(12) = &H31 
KeyArray(13) = &H2d 
KeyArray(14) = &H4f 
KeyArray(15) = &H45 
KeyArray(16) = &H4d 
KeyArray(17) = &H2d 
KeyArray(18) = &H38 
KeyArray(19) = &H39 
KeyArray(20) = &H39 
KeyArray(21) = &H32 
KeyArray(22) = &H36 
KeyArray(23) = &H37 
KeyArray(24) = &H31 
KeyArray(25) = &H2d 
KeyArray(26) = &H30 
KeyArray(27) = &H30 
KeyArray(28) = &H34 
KeyArray(29) = &H33 
KeyArray(30) = &H37 
KeyArray(31) = &H00 
KeyArray(32) = &Hb2 
KeyArray(33) = &H00
KeyArray(34) = &H00 
KeyArray(35) = &H00 
KeyArray(36) = &H58 
KeyArray(37) = &H31 
KeyArray(38) = &H35 
KeyArray(39) = &H2d 
KeyArray(40) = &H33 
KeyArray(41) = &H37 
KeyArray(42) = &H33 
KeyArray(43) = &H37
KeyArray(44) = &H37 
KeyArray(45) = &H00 
KeyArray(46) = &H00 
KeyArray(47) = &H00 
KeyArray(48) = &H00 
KeyArray(49) = &H00 
KeyArray(50) = &H00 
KeyArray(51) = &H00 
KeyArray(52) = &H75 
KeyArray(53) = &H3c 
KeyArray(54) = &H3e 
KeyArray(55) = &H34 
KeyArray(56) = &H65 
KeyArray(57) = &H00 
KeyArray(58) = &H48 
KeyArray(59) = &H82 
KeyArray(60) = &Hfe 
KeyArray(61) = &Hd6 
KeyArray(62) = &H8e 
KeyArray(63) = &H82 
KeyArray(64) = &Had 
KeyArray(65) = &H91 
KeyArray(66) = &H04 
KeyArray(67) = &H00 
KeyArray(68) = &H00 
KeyArray(69) = &H00 
KeyArray(70) = &H00 
KeyArray(71) = &H00 
KeyArray(72) = &H89 
KeyArray(73) = &Ha5 
KeyArray(74) = &Hfa 
KeyArray(75) = &H55 
KeyArray(76) = &H1f 
KeyArray(77) = &H3c 
KeyArray(78) = &Hc4 
KeyArray(79) = &H55 
KeyArray(80) = &H02 
KeyArray(81) = &H00 
KeyArray(82) = &H00 
KeyArray(83) = &H00 
KeyArray(84) = &H00 
KeyArray(85) = &H00 
KeyArray(86) = &H00 
KeyArray(87) = &H00 
KeyArray(88) = &H00 
KeyArray(89) = &H00 
KeyArray(90) = &H00 
KeyArray(91) = &H00 
KeyArray(92) = &H00 
KeyArray(93) = &H00 
KeyArray(94) = &H00 
KeyArray(95) = &H00 
KeyArray(96) = &H00 
KeyArray(97) = &H00 
KeyArray(98) = &H00 
KeyArray(99) = &H00 
KeyArray(100) = &H00 
KeyArray(101) = &H00 
KeyArray(102) = &H00 
KeyArray(103) = &H00 
KeyArray(104) = &H00 
KeyArray(105) = &H00 
KeyArray(106) = &H00 
KeyArray(107) = &H00 
KeyArray(108) = &H00 
KeyArray(109) = &H00 
KeyArray(110) = &H00 
KeyArray(111) = &H00 
KeyArray(112) = &H00 
KeyArray(113) = &H00 
KeyArray(114) = &H00 
KeyArray(115) = &H00 
KeyArray(116) = &H00 
KeyArray(117) = &H00 
KeyArray(118) = &H00 
KeyArray(119) = &H00 
KeyArray(120) = &H00 
KeyArray(121) = &H00 
KeyArray(122) = &H00 
KeyArray(123) = &H00 
KeyArray(124) = &H00 
KeyArray(125) = &H00 
KeyArray(126) = &H00 
KeyArray(127) = &H00 
KeyArray(128) = &H00 
KeyArray(129) = &H00 
KeyArray(130) = &H00 
KeyArray(131) = &H00 
KeyArray(132) = &H00 
KeyArray(133) = &H00 
KeyArray(134) = &H00 
KeyArray(135) = &H00 
KeyArray(136) = &H00 
KeyArray(137) = &H00 
KeyArray(138) = &H00 
KeyArray(139) = &H00 
KeyArray(140) = &H00 
KeyArray(141) = &H00 
KeyArray(142) = &H00 
KeyArray(143) = &H00 
KeyArray(144) = &H00 
KeyArray(145) = &H00 
KeyArray(146) = &H00 
KeyArray(147) = &H00 
KeyArray(148) = &H00 
KeyArray(149) = &H00 
KeyArray(150) = &H00 
KeyArray(151) = &H00 
KeyArray(152) = &H00 
KeyArray(153) = &H00 
KeyArray(154) = &H00 
KeyArray(155) = &H00 
KeyArray(156) = &H00 
KeyArray(157) = &H00 
KeyArray(158) = &H00 
KeyArray(159) = &H00 
KeyArray(160) = &H91 
KeyArray(161) = &H8c 
KeyArray(162) = &H20 
KeyArray(163) = &H48
KeyArray(164) - &H99


'{&Ha4,&H00,&H00,&H00,&H03,&H00,&H00,&H00,&H30,&H30,&H33,&H37,&H31,&H2d,&H4f,&H45,&H4d,&H2d,&H\
'  38,&H39,&H39,&H32,&H36,&H37,&H31,&H2d,&H30,&H30,&H34,&H33,&H37,&H00,&Hb2,&H00,&H00,&H00,&H58,&H31,&H35,&H2d,&H33,&H37,&H33,&H\
'  37,&H37,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H75,&H3c,&H3e,&H34,&H65,&H00,&H48,&H82,&Hfe,&Hd6,&H8e,&H82,&Had,&H91,&H04,&H00,&H\
'  00,&H00,&H00,&H00,&H89,&Ha5,&Hfa,&H55,&H1f,&H3c,&Hc4,&H55,&H02,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H\
'  00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H\
'  00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H\
'  00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H91,&H8c,&H20,&H48,&H99}

MsgBox ConvertToKey(KeyArray)

'"DigitalProductId"=hex:a4,00,00,00,03,00,00,00,30,30,33,37,31,2d,4f,45,4d,2d,\
'  38,39,39,32,36,37,31,2d,30,30,34,33,37,00,b2,00,00,00,58,31,35,2d,33,37,33,\
'  37,37,00,00,00,00,00,00,00,75,3c,3e,34,65,00,48,82,fe,d6,8e,82,ad,91,04,00,\
'  00,00,00,00,89,a5,fa,55,1f,3c,c4,55,02,00,00,00,00,00,00,00,00,00,00,00,00,\
'  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
'  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
'  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,91,8c,20,48,99

Function ConvertToKey(Key)
Const KeyOffset = 52
i = 28
Chars = "BCDFGHJKMPQRTVWXY2346789"
Do
Cur = 0
x = 14
Do
Cur = Cur * 256
Cur = Key(x + KeyOffset) + Cur
Key(x + KeyOffset) = (Cur \ 24) And 255
Cur = Cur Mod 24
x = x -1
Loop While x >= 0
i = i -1
KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
If (((29 - i) Mod 6) = 0) And (i <> -1) Then
i = i -1
KeyOutput = "-" & KeyOutput
End If
Loop While i >= 0
ConvertToKey = KeyOutput
End Function


