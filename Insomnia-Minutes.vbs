'Insomnia Minutes,created by Gulam Mustafa, me_salman148@yahoo.com

Dim VarMns 

VarMns = InputBox("Enter the approximate Execution time in Minutes","Insomnia Minutes") 

If TypeName(VarMns) = "Empty"  Then 

MsgBox("You pressed cancel")

Else If Len(Trim(VarMns)) = 0 Then 

MsgBox("You entered nothing")

Else

exptime = DateAdd("n",VarMns,Now) 

et=FormatDateTime(exptime,4)

tv=TimeValue(et)

MsgBox("Execution time is  " & VarMns & " minutes " & "i.e until " & tv)

exptime = DateAdd("n",VarMns,Now) 

Set Wshell = WScript.CreateObject("WScript.Shell") 

While Now < exptime 

Wscript.Sleep 120000 

Wshell.SendKeys "{NUMLOCK}" 

Wshell.SendKeys "{NUMLOCK}" 

Wend 

MsgBox("Expired at " & Time)

End If

End If
