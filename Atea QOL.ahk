#Persistent
SetMouseDelay,-1
global fileLoc = "C:\Users\"+%A_UserName%+"\Desktop\values.ini"

;Vars
global clipVar = 0 ;Stores the clipboard
global readVar = 0 
global placeHolderVar = 0 
global csnOrder = 0	;HP Channel Services Network number
global pobOrder = 0	;POBnummer
global partNo = 0   	;partnumber
global hpSN = 0		;Serial number
global hpPN = 0		;Product number
global noDays = 7 ;Antal dagar att suspenda


;########################################
;Keybinds 
;########################################

;POB + olika hp nummer
<!1::SendInput %pobOrder% ;Skriver ut pobNR
<!2::SendInput %csnOrder% ;Skriver ut XC/TP nummer
<!3::SendInput %partNo%	  ;Skriver ut HP part nummer
<!z::SendInput %hpSN%	  ;Skriver ut HP Serienummer
<!x::SendInput %hpPN%	  ;Skriver ut HP Produktnummer
<!4::Clean()			  ;Tömmer ovan sparad data

;Diverse pobgrejer
^<:: SendInput %A_UserName%
^q:: SendInput %A_YYYY%-%A_MM%-%A_DD% ;Skriver ut dagens datum
^1:: Tomorrow()						  ;Skriver ut morgondagens datum
^2:: DayAfterTomorrow()				  ;Skriver ut dagen efter imorgons datum
^3:: ThreeDaysFromNow()				  ;Skriver ut om 3 dagars datum
^l:: SuspOneWeek("leverans")		  ;Suspendar ärende som leverans från 3:e part
^i:: SuspOneWeek("komplettering")	  ;Suspendar ärende som information från 3:e part
^k:: SuspOneWeek("kund")			  ;Suspendar ärende som information från kund
^.:: SuspOneWeek("levKund")			  ;Suspendar ärende som leverans från kund
^SC01A:: SuspOneWeek("återkoppling")  ;Suspendar ärende som återkoppling/uppföljning (SC01A = Å)
^SPACE:: SendInput {enter}			  ;Trycker på enter

;Claim 
^Numpad1::SendInput 1153.75
^Numpad2::SendInput 1108.75
^Numpad3::SendInput 428.75


OnClipboardChange("Main")


SaveData(val, key){
	IniWrite, %val% , %fileLoc%, section1, %key%
}

TrimSpace(){
	if(StrLen(clipboard) > 35){
		return
	}
	else{
		clipboard := Trim(clipboard, OmitChars := " 	")
		clipVar := clipboard
		ClassifyInput(clipVar)
	}
}

ClassifyInput(clipIn){
;CSN
	if(RegExMatch(clipIn, "((?!^[LPHUVGSD])[A-Z]{2}?[0-9]{8})") != 0){
		RegExMatch(clipIn, "((?!^[LPHUVGSD])[A-Z]{2}?[0-9]{8})", csnOrder)
		SaveData(csnOrder, "CSN")	
				
	}
	
	else{
		IniRead, csnOrder, %fileLoc%, section1, CSN
	}
;POB
	if(RegExMatch(clipIn, "(\b10)([0-9]{7})") != 0 ){
		RegExMatch(clipIn, "(\b10)([0-9]{7})", pobOrder)
		SaveData(pobOrder, "POB")	
	}
	
	else{
		IniRead, pobOrder, %fileLoc%, section1, POB
	}
;HP - part
	if(RegExMatch(clipIn, "([0-9,A-Z]{6}(\b-)([0-9,A-Z]{3})(?!.*A))") != 0 ){
		RegExMatch(clipIn, "([0-9,A-Z]{6}(\b-)([0-9,A-Z]{3})(?!.*A))", partNo)
		SaveData(partNo, "HPPART")	
	}
	
	else{
		IniRead, partNo, %fileLoc%, section1, HPPART
	}
;HP S/N
	if(RegExMatch(clipIn, "((?!^[TP,10,X])[A-Z, 0-9]{10})") != 0 ){
		RegExMatch(clipIn, "((?!^[TP,10,X])[A-Z, 0-9]{10})", hpSN)
		SaveData(hpSN, "HPSN")	
	}
	
	else{
		IniRead, hpSN, %fileLoc%, section1, HPSN
	}
	
	;MsgBox CSN = %csnOrder%`nPOB= %pobOrder%`npartNo = %partNo%


;HP Prod/N
	if(RegExMatch(clipIn, "([A-Z 0-9]{3}([0-9]{2})([EAV]{2}))") !=0 ){
		RegExMatch(clipIn, "([A-Z 0-9]{3}([0-9]{2})([EAV]{2}))", hpPN)
		SaveData(hpPN, "HPPN")
	
	}
	
	else{
		IniRead, hpPN, %fileLoc%, section1, HPPN
	}
}
Main(){
	IniRead, partNo, %fileLoc%, section1, HPPART
	IniRead, pobOrder, %fileLoc%, section1, POB
	IniRead, csnOrder, %fileLoc%, section1, CSN
	TrimSpace()
	ClaimCopy()
}


Clean(){
	SaveData("", "HPPART")
	SaveData("", "POB")
	SaveData("", "CSN")
}

Tomorrow(){
	tomorrow = %a_now%
	tomorrow += 01, days
	FormatTime, tomorrow, %tomorrow%, yyyy-MM-dd
	SendInput %tomorrow%
	LiftAllKeys()
	return
}

DayAfterTomorrow(){
	tomorrow = %a_now%
	tomorrow += 02, days
	FormatTime, tomorrow, %tomorrow%, yyyy-MM-dd
	SendInput %tomorrow%
	LiftAllKeys()
	return
}

ThreeDaysFromNow(){
	tomorrow = %a_now%
	tomorrow += 03, days
	FormatTime, tomorrow, %tomorrow%, yyyy-MM-dd
	SendInput %tomorrow%
	LiftAllKeys()
	return
}

SuspOneWeek(typ){
	BlockInput, MouseMove
	foundIcon := ClickSuspend()
	if(foundIcon == true){

	if(typ == "leverans"){
		week = %a_now%
		week += %noDays%, days
		FormatTime, week, %week%, yyyy-MM-dd
		SendInput l{tab} %week% {tab} 0900 {tab}{tab} 1234 
	}
	if(typ == "komplettering"){
		week = %a_now%
		week += %noDays%, days
		FormatTime, week, %week%, yyyy-MM-dd
		SendInput i{tab} %week% {tab} 0900 {tab}{tab} 1234 
	}

	
	if(typ == "kund"){
		week = %a_now%
		week += %noDays%, days
		FormatTime, week, %week%, yyyy-MM-dd
		SendInput ii{tab} %week% {tab} 0900 {tab}{tab} 1234 
	}

	if(typ == "återkoppling"){
		week = %a_now%
		week += %noDays%, days
		FormatTime, week, %week%, yyyy-MM-dd
		SendInput å{tab} %week% {tab} 0900 {tab}{tab} 1234 
	}
	if(typ == "levKund"){
		week = %a_now%
		week += %noDays%, days
		FormatTime, week, %week%, yyyy-MM-dd
		SendInput ll{tab} %week% {tab} 0900 {tab}{tab} 1234 
	}
	ClickOK()
	}
	LiftAllKeys()
	BlockInput, MouseMoveOff
}
	
LiftAllKeys(){

	SendInput {ctrl}
	Sleep 80
	SendInput {alt}
	Sleep 80
	SendInput {shift}
	BlockInput On
	Sleep, 10
	BlockInput Off
}

ReqETA(){
	KeyWait Alt
	BlockInput On
	SendInput ANBRO02 %pobOrder% {tab}Hello,{enter}I would like an ETA in event %csnOrder% please.{enter}Thank you.{enter}{enter}Best regards,{enter}Anton{tab}h
	BlockInput Off	
	LiftAllKeys()
}







ClaimCopy(){
	
	SetTitleMatchMode, 1
		if WinActive("Print Details - Internet Explorer"){
										;102432766-VGR-GOTKON5 																XC33494606   102432766-VGR-GOTKON5				
			if((RegExMatch(clipboard, "(\b10)([0-9]{7})(\b-)([A-Z]{3,4}(\b-)([A-Z]{6})([0-9]{1,2})?)"))){
				if(StrLen(clipboard)<40){			
					newClip = %csnOrder%   %clipVar%
					clipboard := newClip
				}
			}
			
		}
}


;Hotstring
::etaa::SendInput, Hello,{enter}I would like an ETA in event %csnOrder% please.{enter}Thank you.{enter}{enter}Best regards,{enter}Anton{ctrl up}{alt up}{shift up}

;################
;FindSuspend

ClickSuspend(){	
	
	WinActivate, POB 21
	sleep 15
	WinGetPos, X, Y, winWidth, winHeight, POB 21
	CoordMode, Pixel, Window
	;ScreenHeight= 1440
	;ScreenWidth= 5120
	ImageSearch, FoundX, FoundY, 0, 0, winWidth, winHeight, *80 C:\Users\%A_UserName%\Desktop\suspendIcons\clicked.png 
	if (ErrorLevel = 2){
    		MsgBox Could not conduct the search.
		return false
	}
	else if (ErrorLevel = 1){
		ImageSearch, FoundX, FoundY, 0, 0, winWidth, winHeight, *80 C:\Users\%A_UserName%\Desktop\suspendIcons\new.png
		if (ErrorLevel = 2){
			MsgBox Could not conduct the search.
			return false
		}
		else if (ErrorLevel = 1){
			MsgBox, 0,,Could not find the icon.,1
			return false
		}
		else{
			CoordMode, Mouse, Window
			Click %FoundX%,%FoundY%
			return true
			}
		}
	else{
		CoordMode, Mouse, Window
		Click %FoundX%,%FoundY%
		return true
	}

	sleep 25
		
}

ClickOK(){
	WinActivate, POB 21
	sleep 500
	WinGetPos, X, Y, winWidth, winHeight, POB 21
	CoordMode, Pixel, Window
	ImageSearch, FoundX, FoundY, 0, 0, winWidth, winHeight, *80 C:\Users\%A_UserName%\Desktop\suspendIcons\OK2.png 
	if (ErrorLevel = 2){
		MsgBox Could not conduct the search.
	}
	else if (ErrorLevel = 1){
		MsgBox Could not find the icon.
	}
	else{
		Click %FoundX%,%FoundY%
	}

}
