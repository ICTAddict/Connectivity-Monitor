/*
Tests internet connectivity; plays an audio file whenever the internet gets disconnected or reconnected
Finally got it to work, plays a notification only once, every connect/disconnect until state change.
Still needs fine tuning, because 128 seconds is a bit long to know that it's down.
Oh, and I still need to add something to reset connection after a certain number of 'Failed Pings'. Done!
Latest change (Circa Oct, 2019) reset the connection via telnet when disconnected, but only do it once until the next time it disconnects.
Script path E:\Africano\Programs\System Tweak\AHK Scripts\Ping\ninix ping.ahk

To Do:
1-Gui (showing Current IP, last disconnect, last reconnect, Connection uptime)
2-Calculate how long it was disconnected (working on it using .ini file). Finished, Now waiting for testing
3-Maybe log that information to the Excel file, but that's too difficult
4-

*/

#NoEnv
#WinActivateForce
#SingleInstance force
SendMode, Input
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2
DetectHiddenWindows, On
DetectHiddenText, On
SetKeyDelay, 0
SetControlDelay -1
count:=0  ; Number of successful pings
Fcount:=0 ; Number of failed pings
sl:=1000  ; Waiting between...
;N:=1 ;can't find that variable anywhere in the script.
If(Ping("213.158.163.119")) ;To avoid startup connection report. 
	count:=3
IP:= WhatIsMyIP()
IniRead, CurrentIP, settings.ini, history, StoredIP
if (currentIP=IP)
	MsgBox, 0, Ninix Ping, No change in IP`nCurrent IP: %IP%, 5
else
{
	CurrentIP=%IP%
	IniWrite, %IP%, settings.ini, history, StoredIP
	FormatTime, Time,, hh:mm:ss tt dd/MM/yyyy
	iniWrite, %Time%, settings.ini, history, ConnectedSince
}
;Run indefinitely
loop, 
{
	If(Ping("102.41.128.1")) ;current is TE-Data's ;213.158.163.74 ;213.158.163.119 ;173.194.164.103
	{
		;Need a message when it's up after a disconnection.
		if (count=2) ; I think it should be 'count=1' instead of 'count>1' otherwise, it'll keep repeating both messages
		{
			if (Fcount>=4) ; That means the 'down' message has been played. This will be reset after announcing that it's back up. Must be >=n from the down message. line 63
			{
				SoundPlay, e:\Africano\Numbers\Internet is back up.wav
				tempIP := WhatIsMyIP()
				if (tempIP != currentIP)
				{
					IniWrite, %tempIP%, settings.ini, history, StoredIP
					FormatTime, Time,, hh:mm:ss tt dd/MM/yyyy
					iniWrite, %Time%, settings.ini, history, ConnectedSince
					
				}
				SplashTextOn, 300, 20, Connection Status, It's Up. Hurray!!
				FormatTime, Time,, hh:mm:ss tt dd/MM/yyyy
				Sleep, 1500
				SplashTextOff
				IP := tempIP
				;WhatIsMyIP() ;It was here, but I moved it outside the if statement
				;FileAppend, IP Address is %IP%`n, E:\Africano\Programs\Lan programs\Logs\Internet connection log.txt
				FileAppend,
(
%time%		`t Connection restored
Current IP Address: %IP%

), E:\Africano\Programs\Lan programs\Logs\Internet connection log.txt ; No size limit
				if (CurrentIP != IP)
				{
					;calculating how long the internet has been disconnected
					rcDuration:=A_now 
					EnvSub, rcDuration, %dcTime%, Seconds
					;rcDuration:=FormatSeconds(rctime) - FormatSeconds(dcTime)
					rcDuration := FormatSeconds(rcDuration)
					FileAppend,  Duration:`t%rcDuration%`n,  E:\Africano\Programs\Lan programs\Logs\Internet connection log.txt
					CurrentIP := tempIP
				}
				;Replacing nslookup with 'MyExternalIP' API 25092020
		
				;run,%comspec% /c nslookup myip.opendns.com 208.67.222.222 >> "...\Internet connection log.txt" ;run, nslookup ... >> c:\test.txt ;experimental 10/9/2020. Not working without CMD, works with %comspec%
		
				Fcount:=0 ;Reset failed ping count.
			}
		}
	;Probably commented out to reset the connection quickly during noise spikes.
		Sleep, %sl%
/* 	If (count<=1)
 * 	{
 * 		sl+=sl ;increases sleep duration all the way to (32 second)
 * 	}
 */
		
		count++
	}
	Else												; Ping Failed
	{
		Fcount++
		if (Fcount=4) ; 'Fcount>1' was causing it to execute every single time. orig: Fcount=2. Tied to line 48 in if (Ping()), be mindful of that.
		{
			SoundPlay, e:\Africano\Numbers\Internet is down.wav
			SplashTextOn, 300, 20, Connection Status, It's Down, Down, Downhill :\
			FormatTime, Time,, hh:mm:ss tt dd/MM/yyyy
			Sleep, 1500
			; Disconnect timestamp
			dcTime:=A_Now
			SplashTextOff
			FileAppend,
(

%time%		`t Connection lost`n
), E:\Africano\Programs\Lan programs\Logs\Internet connection log.txt
		}
		/*
		if (Fcount=30) ; change back to 10 after going to the hg520b router. 
		{
			soundplay, E:\Africano\numbers\rebooting router.wav
			Sleep, 1000
			MsgBox, 257, Reboot Router?, Actually`, just resetting connection., 2 ;OK,Cancel - default 2nd.
			ifmsgbox, cancel
			{}
			ifmsgbox, Timeout
/*			{
				run, chrome.exe
				;winwaitactive, New Tab - Google Chrome ;Something has changed. I don't know what was opening a new tab before, but now it just opens the recent tabs and keeps waiting for a new tab that won't open, making the script stuck. I'm gonna comment it.
				WinWaitActive, New Tab - Google Chrome, , 5 ; This should timeout after 5 seconds
				Sleep, 500
				;send, admin:a6o1i1i1@192.168.1.1/DiagADSL.html{Enter} ;I don't think this would work, maybe it was just filling in the credentials from the browser itself.
			}
              ;
			{
				;reset connection via telnet 22/10/19
				IfWinNotExist, Telnet 192.168.1.1
				{
					run telnet 192.168.1.1 23{Enter}
					Sleep, 500
					controlsend,, a6o1i1i1{Enter}, Telnet 192.168.1.1
				}
				WinActivate, Telnet 192.168.1.1
				Sleep, 50
				Send, {Enter} ;if an earlier connection has timed out but the window is still there, this could cause problems
				Sleep, 50
				
				IfWinNotExist, Telnet 192.168.1.1 ; to avoid a lingering window of a timed-out connection
				{
					run telnet 192.168.1.1 23{Enter}
					Sleep, 500
					controlsend,, a6o1i1i1{Enter}, Telnet 192.168.1.1
				}
				
				Sleep, 50
				controlsend,, set wan adsl reset{Enter}, Telnet 192.168.1.1 ;Send, set wan adsl reset{Enter}
			}

			 ;~ Not working, the protocol isn't identified as html because it doesn't start with http
			;~ run, "admin:a6o1i1i1@http://192.168.1.1/DiagADSL.html" ; Yipee!!! Not working in AHK
			;~ Sleep, 1000
			;~ WinWaitActive, 192.168.1.1/DiagADSL.html - Google Chrome
			;~ controlsend, , {Tab 3}, 192.168.1.1/DiagADSL.html - Google Chrome ;resetting connection
			
			
		}
		*/
		;Need a message only the first time it goes down, until it goes up again.
		count:=0 ; This caused %sl% sleep duration to continue incrementing, so I have to reset it whenever I reset 'count'. 'it' means %sl%. Note to self: write clearer comments
		sl:=1000
	}
}
;*******************************************[ Ping Function ]***************************
Ping(IP, Timeout = 3500, Counts = 2)
	{
	 global
	 
	 ;If your windows is not in English, please modify this!!!
	 CMD_RGX       = Reply from %IP%: bytes=32 time
	 CMD_END       = #########
	 ConWinWidth   = 45  ;Maximum line length (CMD line length)
	 ConWinHeight  = 15  ;How many lines to read from CMD output
	 Buff_Ext_Text = 0

	 Run, %COMSPEC% /C MODE CON: cols=75 lines=25& PING -n %Counts% -w %Timeout% %IP% &ECHO %CMD_END%&PAUSE >NUL ,,Hide, PID_CMD
	 WinWait, ahk_pid %PID_CMD%,, 15
	 If ErrorLevel
		{
		 Return -1
		}

	 If(!AttachConsole(PID_CMD))
		{
		 Process, Close, %PID_CMD%
		 Return -1
		}

	 Loop, 40 ;10 seconds+Execution time...
		{
		 Buff_Ext_Text := GetConsoleText(ConWinWidth, ConWinHeight)
		 If(Buff_Ext_Text = 0)
			{
			 FreeConsole(PID_CMD)
			 Return -1
			}
		 If(InStr(Buff_Ext_Text, CMD_END))
			{
			 If(RegExMatch(Buff_Ext_Text, CMD_RGX))
				{
				 FreeConsole(PID_CMD)
				 Return 1 ;Ping ok...
				}
			 Else
				{
				 FreeConsole(PID_CMD)
				 Return 0 ;Ping fail...
				}
			}
		 Else
			{
			 Sleep, 250
			}
		}
	}
;*******************************************[ Hook to CMD Process ]***************************
AttachConsole(PID)
	{
	 global hConOut

	 If (!DllCall("AttachConsole", "uint", PID))
		{
		 MsgBox,,, AttachConsole failed - error %A_LastError%
		 Return 0
		}

	 hConOut:=DllCall("CreateFile","str","CONOUT$","uint",0xC0000000,"uint",7,"uint",0,"uint",3,"uint",0,"uint",0)
	 If (hConOut = -1)
		{
		 MsgBox,,, CreateFile failed - error %A_LastError%
		 Return 0
		}

	 Return 1
	}
;*******************************************[ Free console and kill CMD Process ]***********************
FreeConsole(PID)
	{
	 global hConOut

	 DllCall("FreeConsole")
	 Process, Close, %PID%

	 hConOut :=
	}
;*******************************************[ Extract CMD Window Output ]********************************
GetConsoleText(ConWinWidth, ConWinHeight)
	{
	 global hConOut

	 VarSetCapacity(info, 24, 0)
	 If (!DllCall("GetConsoleScreenBufferInfo","uint",hConOut,"uint",&info))
		{
		 MsgBox,,, GetConsoleScreenBufferInfo failed - error %A_LastError%
		 Return 0
		}

	 VarSetCapacity(buf, ConWinWidth*ConWinHeight*4, 0)
	 If (!DllCall("ReadConsoleOutput","uint",hConOut,"uint",&buf,"uint",ConWinWidth|ConWinHeight<<16,"uint",0,"uint",&info+10))
		{
		 MsgBox,,, ReadConsoleOutput failed - error %A_LastError%
		 Return 0
		}

	 VarSetCapacity(Temp_Buffer, ConWinWidth*ConWinHeight)
	 Loop % ConWinWidth*ConWinHeight ;%
		{
		 Temp_Buffer .= Chr(NumGet(buf, 4*(A_Index-1), "Char"))
		}

	 Return Temp_Buffer
	}

FormatSeconds(NumberOfSeconds)  ; Convert the specified number of seconds to hh:mm:ss format.
{
    time = 19990101  ; *Midnight* of an arbitrary date.
    time += %NumberOfSeconds%, seconds
    FormatTime, mmss, %time%, mm:ss
    return NumberOfSeconds//3600 ":" mmss
}


^!F10::
if  (Fcount > 0)
{
	;IniRead, 
	MsgBox, Internet is down since %dctime% `nLast IP:%CurrentIP%
}
else
{
	IniRead, ConnectedSince, settings.ini, history, ConnectedSince
	MsgBox, Current IP is: %CurrentIP%`nConnected since: %ConnectedSince% ;I'm going to store last rctime in .ini and get it from there
}


WhatIsMyIP()
{
	request := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	timeoutVal := 59000
	request.SetTimeouts(timeoutVal, timeoutVal, timeoutVal, timeoutVal)   
	request.Open("GET", "http://myexternalip.com/raw")
	request.Send()
	return request.ResponseText
}