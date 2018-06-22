#NoEnv
#SingleInstance Force
SetBatchLines -1
#Include ExcelToArray.ahk

Gui, Add, ListView, xm w700 r10 Grid, A|B|C|D|E|F|G|H|I
Loop, 9
	LV_ModifyCol(A_Index, 75)
Gui, Show,, Get Listview data from excel

SplashTextOn, , , Loading...
Gosub, ImportData
SplashTextOff
Return

ImportData:
	arr := ExcelToArray("test.xlsx")

	for i, dat in arr
		LV_Add("", dat*)
Return

GuiClose:
ExitApp