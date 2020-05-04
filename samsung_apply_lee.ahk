FileSelectFile, FilePath
Xl := ComObjGet(FilePath)
MySlider := 0
Grade = A+|A|A-|B+|B|B-|C+|C|C-|D+|D|D-|F|PASS|FAIL
Gui, Add, Text, x15 y15, 속도
Gui, Add, Slider, x60 y10 h30 w150 vMySlider gSlide ToolTipBottom Range-30-30, 0
Gui, Add, Edit, x20 y43 h20 w30 vEdit, 2
Gui, Add, Text, x58 y46, 부터
Gui, Add, Button, x83 y41 h20 w40 gApyBtn, 적용
Gui, Add, Text, x20 y65 w190 vApyTxt, 적용 버튼을 눌러주세요.
Gui, Add, Button, x20 y80 h30 w70 gStBtn, 시작
Gui, Add, Button, x100 y80 h30 w70 gConBtn, 계속
Gui, Add, Button, x180 y80 h30 w70 gReBtn, 재시작
GuiControl, Disabled, 계속
Gui, Add, Text, x45 y120, 과정, 전공 미리 선택해주세요!
Gui, Add, Text, x45 y140, 일시정지 : F2
Gui, Add, Text, x80 y160, MadeBy.RandomLynn

Gui, Show,, SS_Auto
Return 

GuiClose:
ExitApp

ApyBtn:
Gui, Submit, nohide
num := Edit
SubName := Xl.Sheets(1).Range("D" . num).Value
GuiControl,, ApyTxt, %SubName%부터 시작
return

ConBtn:
Gui, Hide
Pause, Off
return

ReBtn:
Reload
return

Slide:
GuiControlGet, MySlider
Return

StBtn:
{	
	Gui, Submit, nohide
	GuiControl, Disable, 시작
	Gui, Hide
	num := Edit
	short_delay := 300 - MySlider * 8
	long_delay := 700 - MySlider * 10
	sleep, 2000
	loop, 
	{
		;A : Major:1 Or liberal arts Course:2
		Str := Xl.Sheets(1).Range("A" . num).Value
		;If Done -> break
		if(Str == "")
		{
			MsgBox, 입력 완료 . 꼭 저장하세요!
			break
		}
		ImageSearch, VX, VY, 1, 1, A_ScreenWidth, A_ScreenHeight, 01course.PNG
		if(errorlevel==0)
		{
			MouseClick, left, VX, VY, 1
			Send, {Enter}
			Sleep, %long_delay%	
			loop % Str-1
			{							
				Send, {Down}	 
			}
			Send, {Enter}
		}
		else
		{
			Send, {WheelUp}
			Send, {WheelUp}
			Send, {WheelUp}
			Sleep, %short_delay%				
			ImageSearch, VX, VY, 1, 1, A_ScreenWidth, A_ScreenHeight, 01course.PNG
			if(errorlevel==0)
			{
				MouseClick, left, VX, VY, 1
				Send, {Enter}
				Sleep, %long_delay%	
				loop % Str-1
				{							
					Send, {Down}	 
				}
				Send, {Enter}
			}
			else
			{
				MsgBox, 교과목 입력 화면을 띄워주세요!
				Reload
				break
			}
		}
		;B : Year 2016~
		Str := Xl.Sheets(1).Range("B" . num).Value		
		{			
			Sleep, %Short_delay%
			Send, {Tab}
			Sleep, %Short_delay%
			Send, {Tab}
			Sleep, %Short_delay%
			Send, {Enter}
			Sleep, %Short_delay%
			loop % 2016-Str
			{							
				Send, {Down}	
			}
			Send, {Enter}
		}
		;C : Semester 1/2/3/Sum:4/Win:5
		Str := Xl.Sheets(1).Range("C" . num).Value
		{			
			Sleep, %Short_delay%
			Send, {Tab}
			Sleep, %Short_delay%
			Send, {Tab}
			Sleep, %Short_delay%
			Send, {Enter}
			Sleep, %Short_delay%
			loop % Str-1
			{
				Send, {Down}	
			}
			Send, {Enter}
		}		
		;D : Subject Name
		Str := Xl.Sheets(1).Range("D" . num).Value
		{
			Sleep, %Short_delay%
			Send, {Tab}
			Sleep, %Short_delay%
			Send, {Tab}
			Sleep, %Short_delay%
			Clipboard := Str
			Send, {Ctrl Down}
			Sleep, %Short_delay%
			Send, {V}
			Sleep, %Short_delay%
			Send, {Ctrl Up}
			Sleep, %Short_delay%
		}		
		;E : Retake Y:1/N:2
		Str := Xl.Sheets(1).Range("E" . num).Value
		{
			Send, {Tab}
			Sleep, %Short_delay%
			Send, {Enter}
			Sleep, %Short_delay%
			loop % Str-1
			{							
				Send, {Down}	
			}
			Send, {Enter}
		}				
		;F : Credit 1~10
		Str := Xl.Sheets(1).Range("F" . num).Value
		{
			Sleep, %Short_delay%
			Send, {Tab}
			Sleep, %Short_delay%
			Send, {Tab}
			Sleep, %Short_delay%
			Send, {Enter}
			Sleep, %Short_delay%
			loop % Str-1
			{							
				Send, {Down}	
			}
			Send, {Enter}
		}		
		
		;G : Grade A+/A0/A-/F/PASS/FAIL
		Str := Xl.Sheets(1).Range("G" . num).Value
		For i in Array := StrSplit(Grade, "|")
		{
			if(Str == Array[i])
			{
				Str := i
			}
		}
		{
			Sleep, %Short_delay%
			Send, {Tab}
			Sleep, %Short_delay%
			Send, {Tab}
			Sleep, %Short_delay%
			Send, {Enter}
			Sleep, %Short_delay%
			loop % Str-1
			{							
				Send, {Down}	
			}
			Send, {Enter}		
			Sleep, %Short_delay%
			Send, {Tab}
			Sleep, %Short_delay%
			Send, {Tab}
			Sleep, %Short_delay%
			Send, {Enter}	
			Sleep, %long_delay%	
			Sleep, 1000
		}	
		num := num + 1
		ImageSearch, VX, VY, 1, 1, A_ScreenWidth, A_ScreenHeight, 04OK.PNG
		if(errorlevel==0)
		{
			MouseClick, left, VX+24, VY+9
			Sleep, 1000			
		}
	}
}
return

F2::
GuiControl, Enabled, 계속
Gui, Show
Pause, On
return
