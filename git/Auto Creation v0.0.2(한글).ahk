#Singleinstance Force

;GUI ��Ʈ
Gui, Add, Edit, x20 y80 w200 h50 vEdit +ReadOnly -VScroll
Gui, Add, Text, x20 y5 w230 h20, [Auto Creation]
Gui, Add, Text, x20 y25 w230 h20, F1,F2Ű�� ���� ��ũ�θ� �غ��ϼ���


Gui, Add, Text, x20 y65 w230 h15, ����â

Gui, Add, Button, x20 y235 w95 h25, ����
Gui, Add, Button, x120 y235 w95 h25, ����

Gui, Add, Text, x20 y140 w230 h20, Ctrl+Q�� ���� ��ǥ�� Ž���ϼ���
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
;��ǥ���� : ctrl+q
^q::
	CoordMode, Mouse, Screen
	MouseGetPos,vEX,vEY
	GuiControl,,X,%vEX%
	GuiControl,,Y,%vEY%
return

;���Ϸε� : F1
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
	;���� ��ü�� A�� ��ü�� TEXT�� ����

	editPrint("�غ� �� . . . (1/2)")
	
return

;��ɾ� �ε� : F2
F2::
	cA1:=object()
	cX:=Object()
	cY:=Object() ;�������� ������ ��ɰ� ��ǥ �����͵�
	
	row := 2 ;2������ Ž��
		
	while(xl.range("A"row).value) ;A�� �� ���� ������ ���� �ݺ�
	{
		cA1.Push(xl.range("A"row).value)
		cX.Push(xl.range("B"row).value)
		cY.Push(xl.range("C"row).value)
		row := row + 1
	}

		fileLoadCheck:=TRUE ;���� �ε� üũ�ϴ� �ε���

		editPrint("�غ� �� . . . (2/2)")
		editPrint("�غ� �Ϸ�!")
return

F3::ExitApp ; �������� F3

;��ư �۵�

ButtonExcel:
{
	xlIndex:=2
	while(xl.range("A"xlIndex).value) ;A�� �� ���� ������ ���� �ݺ�
	{
		xlIndex := xlIndex + 1
	}

	xl.range("A"xlIndex).value =
	
	
	
}


Button����:
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
			Else ;���� �̸� Ʋ���� �� ����
			{
				editPrint("Error")
			}
		}
			editPrint("�۾� �Ϸ�!")
	}
}
return

Button����:
{
Guiclose:
ExitApp
}