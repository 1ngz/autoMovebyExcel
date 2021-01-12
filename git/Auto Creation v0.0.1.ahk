#Singleinstance Force

;GUI 파트
Gui, Add, Edit, x20 y80 w200 h50 vEdit +ReadOnly -VScroll
Gui, Add, Text, x20 y5 w230 h20, [Auto Creation]
Gui, Add, Text, x20 y25 w230 h20, F1키를 통해 매크로를 준비하세요


Gui, Add, Text, x20 y65 w230 h15, 상태창

Gui, Add, Button, x20 y235 w95 h25, 시작(F3)
Gui, Add, Button, x120 y235 w95 h25, 종료

Gui, Add, Text, x20 y140 w230 h20, Ctrl+Q를 통해 좌표를 탐색하세요
Gui, Add, Text, x25 y173 w20 h20, X:
Gui, Add, Text, x25 y203 w20 h20, Y:


Gui, Add, Button, x130 y170 w55 h55, Excel

Gui, Add, Edit, x45 y170 w60 h20 vX -VScroll +Number
Gui, Add, Edit, x45 y200 w60 h20 vY -VScroll +Number

Gui, Show, w240 h270, Auto Creation
return


editPrint(str){
	Guicontrol,,Edit,%str%
}

copy(){
SendInput, ^{c}
return
}

paste(){
SendInput, ^{v}
return
}

select(){
	SendInput, ^{a}
}

iEnter(){
	SendInput, {Enter}
}
;좌표측정 : ctrl+q
^q::
	CoordMode, Mouse, Screen
	MouseGetPos,vEX,vEY
	GuiControl,,X,%vEX%
	GuiControl,,Y,%vEY%
return

;파일로드 : F1
F1::	
	path = %A_ScriptDir%
	FileSelectFile, path
	xl := ComObjCreate("Excel.Application")
	try
	{
	xl.Workbooks.Open(path)
	}
	catch e
	{
    editPrint("File Error code:001") 
 	Exit
	}

	xl.Visible:=TRUE

	xl.Range("A:A").NumberFormat := "@"
	;엑셀 객체의 A열 전체를 TEXT로 설정

	editPrint("준비 중 . . . (1/2)")
	
	cA1:=object()
	cX:=Object()
	cY:=Object() ;엑셀에서 가져온 명령과 좌표 데이터들
	
	row := 2 ;2열부터 탐색
		
	while(xl.range("A"row).value) ;A열 셀 값이 존재할 동안 반복
	{
		cA1.Push(xl.range("A"row).value)
		cX.Push(xl.range("B"row).value)
		cY.Push(xl.range("C"row).value)
		row := row + 1
	}

		fileLoadCheck:=TRUE ;파일 로드 체크하는 인덱스

		editPrint("준비 중 . . . (2/2)")
		editPrint("준비 완료!")

return

	

F3:: ;F3:시작
{
CoordMode, Mouse, Screen
	if !fileLoadCheck
	{
		editPrint("No files")
	}
	Else
	{
		For index, value in cA1
		{
			if value=Left
			{
				MouseClick,Left,cX[index],cY[index]
				sleep,500
			}
			else if value = DLeft
			{
				MouseClick,Left,cX[index],cY[index]
				sleep,10
				MouseClick,Left,cX[index],cY[index]
				sleep,500
			}
			else if value = Right
			{
				MouseClick,Right,cX[index],cY[index]
				sleep,500
			}
			else if value = Key
			{
				keyval := cX[index]
				if keyval=copy
					copy()
				else if keyval=paste
					paste()	
				else if keyval=select
					select()
				else if keyval=enter
					iEnter()
				Else
					SendInput %keyval%
				
				sleep,500
			}
			Else ;오더 이름 틀렸을 시 에러
			{
				editPrint("Error")
			}
		}
			editPrint("작업 완료!")
	}
}
return



F4::ExitApp ; 강제종료 F3

;버튼 작동

ButtonExcel:
{
	xlIndex:=2
	while(xl.range("A"xlIndex).value) ;A열 셀 값이 존재할 동안 반복
	{
		xlIndex := xlIndex + 1
	}

	xl.range("A"xlIndex).value =
	
}


Button시작(F3):
{
	CoordMode, Mouse, Screen
	if !fileLoadCheck
	{
		editPrint("No files")
	}
	Else
	{
		For index, value in cA1
		{
			if value=Left
			{
				MouseClick,Left,cX[index],cY[index]
				sleep,500
			}
			else if value = DLeft
			{
				MouseClick,Left,cX[index],cY[index]
				sleep,10
				MouseClick,Left,cX[index],cY[index]
				sleep,500
			}
			else if value = Right
			{
				MouseClick,Right,cX[index],cY[index]
				sleep,500
			}
			else if value = Key
			{
				keyval := cX[index]
				if keyval=copy
					copy()
				else if keyval=paste
					paste()	
				else if keyval=select
					select()
				else if keyval=enter
					iEnter()
				Else
					SendInput %keyval%
				
				sleep,500
			}
			Else ;오더 이름 틀렸을 시 에러
			{
				editPrint("Error")
			}
		}
			editPrint("작업 완료!")
	}
}
return

Button종료:
{
Guiclose:
ExitApp
}