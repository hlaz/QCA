#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


#include Gdip.ahk
#include Gdip_ImageSearch.ahk

	Loop
	{
		CoordMode, pixel, screen
		coordmode, mouse, screen
		ImageSearch, FoundX, FoundY, 0,0, A_ScreenWidth, A_ScreenHeight, *50 %A_ScriptDir%\img\select_file.png
		if (ErrorLevel = 0)
		{
			Click, right, %FoundX%, %FoundY%
			sleep, 3000
			Send, m
			sleep, 3000
			FormatTime, CurrentDateTime,, MMddyyyy
			Send, Repair %CurrentDateTime%
			sleep, 3000
			Send, {SHIFT}+{HOME}
			sleep, 1000
			Send, ^c
			sleep, 1000
			Send, {Enter}
			ExitApp
		}
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