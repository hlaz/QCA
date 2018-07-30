/*
scripts to automatically closes orders from HP Global Newton


for now, it only closes orders that ends with 01. I need to make conditions to check either it ends with 01 or 02. 
What I'm thinking about right now is take a screenshot of the last 2 numbers on order, 01 and 02, then use image detection as a condition to type either 01 or 02.

at this point, what I've done is, I finished orders that end with 02 manually, then ran this script to do other things.


pages we use
start - Order Search - Manual Logout - Google Chrome
shipping - shipping - Google Chrome
finished - Shipping Submit - Google Chrome
excel - Microsoft Excel - Shipping-7-20-2018
*/

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.



Gui, Add, Text, x60 y25 w50 h20 vA, Ready
Gui, Add, Text, x60 y45 w50 h20 vB, Closed 0 orders
Gui, Add, Button, x20 y85 w110 h20, Start
Gui, Add, Button, x20 y110 w110 h20, Stop
Gui, Show

return


ButtonStop:
{
	Start_macro := false
	ExitApp
}
return




F2::
WinActivate,Microsoft Excel - Shipping-7-26-2018
Send, ^c
Sleep, 300
WinActivate,Order Search - Manual Logout - Google Chrome
mouseclick, left, 306, 344
Send, ^v
Sleep, 300
mouseclick, left, 410, 344
Send, 01 
mouseclick, left, 505, 340
Sleep, 3000
WinActivate,shipping - Google Chrome
Send, {Down}
Send, {Down}
Send, {Down}
Send, {Down}
Send, {Down}
Sleep, 600
WinActivate,Microsoft Excel - Shipping-7-26-2018
Send {Right}
Send, ^c
Sleep, 600
WinActivate,shipping - Google Chrome
Sleep, 600
mouseclick, left, 968, 886
Send, ^v
Sleep, 3000
WinActivate,Microsoft Excel - Shipping-7-26-2018
Send {Left}
Send {Down}
Sleep, 600
WinActivate,shipping - Google Chrome
mouseclick, left, 1052, 960
Sleep, 2000
mouseclick, left, 144, 92 ;favorite
return

F8::
ExitApp
return

F9::
mousegetpos, xx, yy
msgbox X%xx% Y%yy%
return
