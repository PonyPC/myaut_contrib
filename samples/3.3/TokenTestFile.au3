#NoTrayIcon
#region
	#AutoIt3Wrapper_UseUpx=n
#endregion
FileInstall(">>>AUTOIT SCRIPT<<<", @ScriptDir & "\TokenTestFile_Extracted.au3")
Exit
$DUMMY1 = (1 + 2 - 3 / 4) ^ 2
$DUMMY2 += 1 > 2 < 3 <> 4 >= 5 <= 6 = 7 & 8
$DUMMY3 -= True == True = -True
$DUMMY4 /= 1
$DUMMY5 /= (-1)
$DUMMY6 *= 1.123
$DUMMY7 &= -1.53
$DUMMY8 = 1234567887654321
$DUMMY9 = -1234567887654321
MYFUNC($DUMMY1, $DUMMY2)
$OSHELL = ObjCreate("shell.application")
$OSHELLWINDOWS = $OSHELL.windows
Func MYFUNC($VALUE, $VALUE2)
	Dim $ARRAY1[4]
	$ARRAY1[2] = 1
	$DUMMY1 = ' " '
	$DUMMY2 = ' " '
	$DUMMY3 = " ' ' """"  "
EndFunc
Exit
; DeTokenise by myAut2Exe >The Open Source AutoIT/AutoHotKey script decompiler< 2.12 build(196)
