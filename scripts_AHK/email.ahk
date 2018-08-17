#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#include Gdip.ahk
#include Gdip_ImageSearch.ahk

EnvAdd, Yesterday, -1, days
FormatTime, Yesterday, %Yesterday%,MMddyyyy
	
	Loop
	{
		CoordMode, pixel, screen
		coordmode, mouse, screen
		ImageSearch, FoundX, FoundY, 0,0, A_ScreenWidth, A_ScreenHeight, *50 %A_ScriptDir%\img\open_outlook.png
		if (ErrorLevel = 0)
		{
			Click, %FoundX%, %FoundY%, 2
			sleep, 9000
			break
		}
	}
	
	
	Loop
	{
		CoordMode, pixel, screen
		coordmode, mouse, screen
		ImageSearch, FoundX, FoundY, 0,0, A_ScreenWidth, A_ScreenHeight, *50 %A_ScriptDir%\img\open_new_email.png
		if (ErrorLevel = 0)
		{
			Click, %FoundX%, %FoundY%, 1
			sleep, 5000
			break
		}
	}
	
	
	Loop
	{
		CoordMode, pixel, screen
		coordmode, mouse, screen
		ImageSearch, FoundX, FoundY, 0,0, A_ScreenWidth, A_ScreenHeight, *50 %A_ScriptDir%\img\attach_file.png
		if (ErrorLevel = 0)
		{
			mouseclick, left, 139, 160
			sleep, 3000
			Send, al
			sleep, 3000
			mouseclick, left, 149, 182
			sleep, 4000
			break
		}
	}
	
	Loop
	{
		CoordMode, pixel, screen
		coordmode, mouse, screen
		ImageSearch, FoundX, FoundY, 0,0, A_ScreenWidth, A_ScreenHeight, *50 %A_ScriptDir%\img\attach_file.png
		if (ErrorLevel = 0)
		{
			mouseclick, left, 242, 160
			sleep, 3000
			Send, pa
			sleep, 3000
			mouseclick, left, 256, 182
			sleep, 4000
			break
		}
	}
	
	
	Loop
	{
		CoordMode, pixel, screen
		coordmode, mouse, screen
		ImageSearch, FoundX, FoundY, 0,0, A_ScreenWidth, A_ScreenHeight, *50 %A_ScriptDir%\img\attach_file.png
		if (ErrorLevel = 0)
		{
			mouseclick, left, 145, 210
			sleep, 3000
			Send, ^v
			sleep, 3000
			break
		}
	}
	
	Loop
	{
		CoordMode, pixel, screen
		coordmode, mouse, screen
		ImageSearch, FoundX, FoundY, 0,0, A_ScreenWidth, A_ScreenHeight, *50 %A_ScriptDir%\img\attach_file.png
		if (ErrorLevel = 0)
		{
			Click, %FoundX%, %FoundY%, 1
			sleep, 5000
			break
		}
	}
	
	
	Loop
	{
		CoordMode, pixel, screen
		coordmode, mouse, screen
		ImageSearch, FoundX, FoundY, 0,0, A_ScreenWidth, A_ScreenHeight, *50 %A_ScriptDir%\img\select_file.png
		if (ErrorLevel = 0)
		{
			Click, %FoundX%, %FoundY%, 2
			sleep, 5000
			break
		}
	}
	
	Loop
	{
		CoordMode, pixel, screen
		coordmode, mouse, screen
		ImageSearch, FoundX, FoundY, 0,0, A_ScreenWidth, A_ScreenHeight, *50 %A_ScriptDir%\img\file_attached.png
		if (ErrorLevel = 0)
		{
			mouseclick, left, 101, 317
			sleep, 3000
			Send, ^v
			sleep, 3000
			break
		}
	}
	
	Loop
	{
		CoordMode, pixel, screen
		coordmode, mouse, screen
		ImageSearch, FoundX, FoundY, 0,0, A_ScreenWidth, A_ScreenHeight, *50 %A_ScriptDir%\img\send_email.png
		if (ErrorLevel = 0)
		{
			Click, %FoundX%, %FoundY%, 1
			sleep, 5000
			ExitApp
		}
		
		return
	}
	





F5::ExitApp




Gdip_ImageSearchWithHwnd(Hwnd,Byref X,Byref Y,Image,Variation=0,Trans="",sX = 0,sY = 0,eX = 0,eY = 0)
{
	SysGet, wFrame, 7
	SysGet, wCaption, 4
	gdipToken := Gdip_Startup()
	bmpHaystack := Gdip_BitmapFromHwnd(Hwnd)
	bmpNeedle := Gdip_CreateBitmapFromFile(Image)
	if( sX!= 0 || sY!= 0 || eX!= 0 || eY != 0)
	{
	sX := sX + wFrame
	sY := sY + wCaption + wFrame
	eX := eX + wFrame
	eY := eY + wCaption + wFrame
	}
	RET := Gdip_ImageSearch(bmpHaystack,bmpNeedle,LIST,sX,sY,eX,eY,Variation,Trans,1,1)
	Gdip_DisposeImage(bmpHaystack)
	Gdip_DisposeImage(bmpNeedle)
	Gdip_Shutdown(gdipToken)
	StringSplit, LISTArray, LIST, `,
	X := LISTArray1 - wFrame
	Y := LISTArray2 - wCaption - wFrame

	if(RET = 1)
	{
	return true
	}
	else
	{
	return false
	}
}