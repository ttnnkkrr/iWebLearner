;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;	Written by Tank
;;	Based on Seans most excelent work COM.ahk
;;	http://www.autohotkey.com/forum/viewtopic.php?t=22923
;;	some credit due to Lexikos for ideas arrived at from ScrollMomentum
;;	http://www.autohotkey.com/forum/viewtopic.php?t=24264
;;	1-17-2009
;;	Please use and distribute freely
;;	Please do not claim it as your own
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; <COMPILER: v1.0.48.3>

	;	###### CONSTANTS/PARAMETERS ######

	GuiWinTitle = iWebBrowser2 Learner Build ID: 2.6 ; added by jethrow
	link=linked,traversed
	iwf=iWeb Examples (press Ctrl+e when mouse is over desired element)
	WM_LBUTTONDOWN = 0x201
	mintab=2,4,6
	ClickToClip=1
	Copyable=Title,URLL,MouseX,MouseY,EleIndex,EleIDs,EleName
	,theFrame,html_value,html_text,iWeb,CodeWindow,Source,Form,SampleScript
	framecols=Type/Level|Index|Name|ID
	tabs=Viewer|Script Writer|Source|Forms|Templates|About
	funcs=
	(LTrim Join
	iWeb_Init()|iWeb_Term()|iWeb_Release()|iWeb_getWin()
	|iWeb_DomWin()|iWeb_Nav()|iWeb_Complete()|iWeb_getDomObj()
	|iWeb_setDomObj()|iWeb_TableParse()|iWeb_clickDomObj()
	|iWeb_clickText()|iWeb_clickHref()|iWeb_clickValue()
	|iWeb_execScript()|iWeb_SelectOption()
	)

	;	###### DIRECTIVES ######

	DetectHiddenWindows,on
	SetTitleMatchMode,slow	
	SetTitleMatchMode,2
	SetControlDelay, -1
	ListLines, On


	;	###### AUTO-EXECUTE ######

	COM_CoInitialize()
	COM_Error(0)
	fu=bar

	;	###### MENU BAR LAYOUT ######

	Menu, FileMenu, Add, &Restart	F10, MenuRestart
	Menu, FileMenu, Add, &Close	Alt+F4, MenuClose
	Menu, FileMenu, Add, Next Tab	Ctrl+Right Arrow,TabRight
	Menu, FileMenu, Add, Previous Tab	Ctrl+Left Arrow,TabLeft
	Menu, FunctionMenu, Add, Show HTML Map, htmlmap
	Menu, FunctionMenu, Add, Show Form Map, formmap
	Menu, FunctionMenu, Add, Show Table Map, tablemap
	Menu, FunctionMenu, Add, Show Link Map, linkmap
	Menu, FunctionMenu, Add, Show ImageSource Map, imagesourcemap
	
	Menu, OptionsMenu, Add, Always On Top, ontop
	Menu, OptionsMenu, Add, Element Outline, ToggleOutline ; added by jethrow
	Menu, OptionsMenu, Check, Element Outline ; added by jethrow
	Menu, OptionsMenu, Add, Left Click/Copy, MenuLeftClick
	Menu, OptionsMenu, Check, Left Click/Copy ; added by jethrow
	Menu, HelpMenu, Add, About iWebBrowser2 Learner..., MenuAbout
	Menu,MenuGroup,Add,&File, :FileMenu
	Menu,MenuGroup,Add,F&unctions, :FunctionMenu
	Menu,MenuGroup,Add,&Options, :OptionsMenu
	Menu,MenuGroup,Add,&Help, :HelpMenu

	;	###### GUI LAYOUT ######

	Gui, Menu, MenuGroup
	Gui, Add, Tab2, x0 y0 w390 h572 0x40 gIsTab vCurrentTab AltSubmit, %tabs%
	Gui, Tab, Viewer

	;	###### VIEWER TAB ######

	;		~~~ Main GroupBox layout~~~

	Gui, Add, GroupBox, x10 y30 w365 h110 , Browser	
	Gui, Add, GroupBox, xp yp+115 wp hp-20 , Element Info	;	x15 y200 w365 h40
	Gui, Add, GroupBox, xp yp+95 wp hp-20 , OuterHTML		;	x25 y295 w385 h65
	Gui, Add, GroupBox, xp yp+75 wp hp+15 , Frames	;	x25 y365 w385 h125	
	Gui, Add, GroupBox, xp yp+90 wp hp-15 , %iwf%		;	x15 y510 w405 h70

	;		~~~ Browser sub-Groupbox layout ~~~

	Gui, Add, GroupBox, x22 y45 w345 h40, Page Title	;	x25 y45 w385 h40
	Gui, Add, GroupBox, xp yp+45 wp hp, URL/Address		;	x25 y90 w385 h40

	;		~~~Title/URL layout~~~

	Gui, Add, Edit, x30 y60 w330 h20 vTitle ReadOnly,	;	x33 y60 w330 h20
	Gui, Add, Edit, xp yp+45 wp hp vURLL ReadOnly,		;	x33 y105 w330 h20

	;		~~~Element Info layout~~~

	Gui, Add, Text, x25 y167 w65 h17 , Index		;	x25 y218 w65 h17
	Gui, Add, Text, xp+100 yp wp hp , Name			;	x125 y218 w65 h17
	Gui, Add, Text, xp+135 yp wp hp , ID			;	x260 y218 w65 h17
	Gui, Add, Edit, x60 yp-2 wp-10 hp vEleIndex ReadOnly,	;	x60 y216 w55 h17
	Gui, Add, Edit, xp+105 yp wp+25 hp vEleName ReadOnly,	;	x175 y216 w80 h17
	Gui, Add, Edit, xp+115 yp wp hp vEleIDs ReadOnly,	;	x310 y216 w80 h17
	Gui, Add, GroupBox, x20 yp+20 w345 h40, Value/InnerText
	Gui, Add, Edit, xp+8 yp+15 wp-16 hp-20 vhtml_value ReadOnly,	

	;		~~~ OuterHTML Edit box ~~~

	Gui, Add, Edit, x18 yp+57 wp+20 hp+23 Border vhtml_text ReadOnly,	;	x30 y390 w325 h40

	;		~~~ Frames ListView ~~~

	Gui, Add, ListView, xp yp+75 wp h60 Background0xECE9D8 Border -LV0x10 -multi NoSortHdr AltSubmit vVarListView gSubListView, %framecols%
	 LV_ModifyCol(1, 75), LV_ModifyCol(2,75), LV_ModifyCol(3,97), LV_ModifyCol(4,97)	;	, LV_ModifyCol(2,"Center")

	;		~~~ iWeb Functions ListBox/Clipboard Checkbox ~~~

	Gui, Add, ListBox, xp yp+90 wp r3 gCopyToWriter viWeb
	Gui, Add, CheckBox, xp+280 yp+63 viClip, To Clipboard

	;	###### SCRIPT WRITER TAB ######

	Gui, Tab, Script Writer
	Gui, Add, GroupBox, x10 y30 w370 h205 , Code Testing Sandbox
	Gui, Add, DropDownList, xp+10 yp+15 w200 r8 vFunc, List of functions||%funcs%
	Gui, Add, Button, xp+205 yp-1 gAddToScript, Add
	Gui, Add, Edit, xp-205 yp+30 w348 h130 vCodeWindow,
	Gui, Add, Button, x264 y207 w50 h20 gRunScript, Test
	Gui, Add, Button, xp+55 yp wp hp gSaveScript, Save
	Gui, Add, CheckBox, x20 y211 vsClip, To Clipboard	;	gIsBottom 
;	Gui, Add, CheckBox, xp+80 yp gIsClip vBottom, To End of Script

	;	###### SOURCE TAB ######

	Gui, Tab, Source
	Gui, Add, Edit, x10 y35 w370 h450 vSource ReadOnly,

	;	###### FORMS TAB ######

	Gui, Tab, Forms
	Gui, Add, Button, x10 y35 h22 gGetData, Get Data
	Gui, Add, DropDownList, xp+65 yp+1 w200 gLoadForm vSelForm AltSubmit,
	Gui, Add, Edit, xp-65 yp+30 w370 h150 Border vForm,
;	Gui, Add, Text, xp yp+220 w150 h40 , Sample Script
;	Gui, Add, Edit, xp yp+25 w370 h160 Border vSampleScript,
;	Gui, Add, Button, xp+275 yp+185 hh22 gFormToClip, Copy to Clipboard

	;	###### ABOUT TAB ######

	Gui, Tab, About
	Gui, Add, GroupBox, x10 y30 w370 h70 , Main contributors to this project
	Gui, Add, Text, xp+10 yp+20 w50 h40 , Tank`nJethrow`nSinkfaze
;	Gui, Add, GroupBox, xp-10 yp+55 w370 h70 , Special thanks
;	Gui, Add, Text, xp+10 yp+20 w300, % "Chris Mallett - Creator of AutoHotkey`n"
;		. "Sean - Creator of COM Standard Library for AutoHotkey`n"
;		. "Lexikos - Creator of AutoHotkey_L (and much more)"
	Gui, Add, Picture, x80 y108 , ahklogo.png
	Gui, +Delimiter`n
	Gui, Show, Center w390 h510, %GuiWinTitle% ; modified by jethrow
	OnMessage(WM_LBUTTONDOWN, "WM_LBUTTONDOWN")
	Gosub,ontop
	
	ToggleOutline:=CurrentTab:=True ; added by jethrow
	Loop, 4 { ; added by jethrow
		Gui, % A_Index+1 ": -Caption +ToolWindow"
		Gui, % A_Index+1 ": Color" , Red
		Gui, % A_Index+1 ": Show", NA h0 w0, outline%A_Index%
		outline%A_Index% := WinExist("outline" A_Index " ahk_class AutoHotkeyGUI")
	}
	
	;##########~~~BEGIN GETWIN/OUTLINE SUBS~~~##########
	
	GetWin:

	SetBatchLines, % !GetKeyState("CTRL","P") ? "10ms" : -1

	If ToggleOutline { ; skip it Menu Item "Element Ouline" isn't selected
		While, GetKeyState("LButton","P") {
			WinGetPos, Wx, Wy, , , ahk_id %WinHWND%
			If(Wx <> Stored_Wx || Wy <> Stored_Wy)
				Outline("Hide")
			Sleep, 10
		}
		WinGetTitle, WinTitle, ahk_id %WinHWND%
		WinGetPos, Wx, Wy, Ww, Wh, ahk_id %WinHWND%
		If(Ww <> Stored_Ww || Wh <> Stored_Wh || WinTitle <> Stored_WinTitle) && (WinHWND = Stored_WinHWND)
			Outline("Hide"), Resized:=True ; Hide outline if the window either changes size or WinTitle (set variable "Resized" as true)
		Else If(Wx <> Stored_Wx || Wy <> Stored_Wy) && !Resized && (WinHWND = Stored_WinHWND) { ; move outline if Window moves (not if window has been resized)
			Xmove:=Wx-Stored_Wx, Ymove:=Wy-Stored_Wy
			Loop, 2
				Fx%A_Index%+0 ? Fx%A_Index%+=Xmove:"", Fy%A_Index%+0 ? Fy%A_Index%+=Ymove:""
			Outline( x1+=Xmove, y1+=Ymove, x2+=Xmove, y2+=Ymove ), Stored_xorg+=Xmove, Stored_yorg+=Ymove
		}
		Stored_Wx:=Wx, Stored_Wy:=Wy, Stored_Ww:=Ww, Stored_Wh:=Wh, Stored_WinTitle:=WinTitle, Stored_WinHWND:=WinHWND
		GoSub, SetOutlineLevel
	}
	
			COM_Error(0)
	% !GetKeyState("CTRL","P") ? "" : IE_HtmlElement() ;GetKeyState("LButton","P") ? "" : IE_HtmlElement()
	Goto,GetWin

	ontop:

	ontop:=!ontop
	Menu,OptionsMenu,ToggleCheck,Always On Top ; moved from above - jethrow
	Gui, % !ontop ? "-AlwaysOnTop" : "+AlwaysOnTop"
	Return

	ToggleOutline:	; added by jethrow

	Menu,OptionsMenu,ToggleCheck,Element Outline
	ToggleOutline := !ToggleOutline
	Outline("Hide")
	Return

	;##########~~~END GETWIN/OUTLINE SUBS~~~##########

	;##########~~~BEGIN HOTKEYS~~~##########

	#s::
	psv:=COM_CreateObject("SAPI.SpVoice")
	COM_Invoke(psv, "Speak", textOfObj)
	COM_Release(psv)
	Return

	^e::
	MouseGetPos, , , wID
	Gui, Submit, NoHide
	if (WinGetClass(wID)<>"IEFrame")
		return
	if theFrame {
		Pos=1
		While Pos:=RegExMatch(theFrame,"is)sourceIndex]=(.*?) \*\*\[name]= (.*?) \*\*\[id]= (\V*)",f,Pos+StrLen(f))
			fpath.=((f1) ? (f1) : ((f3) ? (f3) : (f2))) . (A_Index=1 ? "" : ",")
	}
	oFrm:=fpath ? ",""" fpath """" : ""
	dObj:=((EleName) ? (EleName) : ((EleIDs) ? (EleIDs) : (EleIndex)))
	pacc := iWebacc_AccessibleObjectFromPoint()
	oRole:=((paccChild:=iWebacc_Child(pacc, _idChild_)) ? iWebacc_Role(paccChild) : iWebacc_Role(pacc,_idChild_))
	oValue:=((paccChild:=iWebacc_Child(pacc, _idChild_)) ? iWebacc_Value(paccChild) : iWebacc_Value(pacc,_idChild_))
	if InStr(HTMLTag,"INPUT") {
		if RegExMatch(oRole,"(?:push|radio) button")
			res:="iWeb_clickDomObj(pwb,""" dObj """" oFrm ")"
			. "`niWeb_clickValue(pwb,""" html_value """" oFrm ")"
		else if (oRole="check box")
			res:="iWeb_clickDomObj(pwb,""" dObj """" oFrm ")"
			. "`niWeb_Checked(pwb,""" dObj """" oFrm ")"
		else
			res:="iWeb_setDomObj(pwb,""" dObj 
			. ((html_value) ? """,""" html_value """" oFrm ")" : """,""<< enter value >>""" oFrm ")")
	}
	else if (oRole="combo box") {
		Pos=1
		While Pos:=RegExMatch(html_text,"is).*?\</OPTION\>",m,pos+strlen(m))
		{
			if InStr(m,html_value) {
			    sIndex:=A_Index-1
			    break
			}
		}
		res:="iWeb_setDomObj(pwb,""" dObj
		 . ((html_value) ? """,""" html_value """" oFrm ")" : """,""<< enter value >>""" oFrm ")")
		 . "`niWeb_selectOption(pwb,""" dObj """,""" sIndex """" oFrm ")"
	}
	else if InStr(link,RegExReplace(oRole,"\W"))
		res:="iWeb_clickDomObj(pwb,""" dObj """" oFrm ")"
		 . ((html_value) ? "`niWeb_clickText(pwb,""" html_value """" oFrm ")" : "")
		. ((RegExMatch(oValue,"^javascript")) ? "`niWeb_execScript(pwb,""" oValue """" oFrm ")" : "")
		. ((oValue) ? "`niWeb_clickHref(pwb,""" oValue """" oFrm ")" : "")
	else
		res:="iWeb_getDomObj(pwb,""" dObj """" oFrm ")"
	GuiControl, , iWeb, `n%res%
	TabActivate(0)
	WinGet, st, MinMax, %GuiWinTitle%
	if st=-1
		WinRestore, %GuiWinTitle%
	WinGetPos, x, y, , , %GuiWinTitle%
	WinMove, %GuiWinTitle%, , % x > A_ScreenWidth - 396 ? A_ScreenWidth - 396 : 
	 , % y > A_ScreenHeight - 562 ? 0 : , , % CurrentTab=1 ? 562 :  ; 278
	VarSetCapacity(res,0), VarSetCapacity(fpath,0)
	return

	^t::Gosub, getTables
	
	#IfWinActive iWebBrowser2 Learner
	^RIGHT::
	TabRight:
	Send ^{PGDN}
	return

	^LEFT::
	TabLeft:
	Send ^{PGUP}
	return

	F1::
	Gui, Submit, NoHide
	WinMove, %GuiWinTitle%, , , , , % CurrentTab=1 ? (GuiHeight()=562 ? 293 : 562) :  ; 582, 278
	return

	F10::
	MenuRestart:
	Reload

	!F4::
	GuiClose:
	MenuClose:
	COM_CoUninitialize()
	ExitApp

	;##########~~~END HOTKEYS~~~##########

IE_HtmlElement()
{
			COM_Error(0)
	CoordMode, Mouse
	MouseGetPos, xpos, ypos,, hCtl, 3
	WinGetClass, sClass, ahk_id %hCtl%
	If Not   sClass == "Internet Explorer_Server"
		|| Not   pdoc := IE_GetDocument(hCtl)
			Return


	GuiControl,Text,MouseX,%	xpos
	GuiControl,Text,MouseY,%	ypos
	pwin :=   COM_QueryService(pdoc ,"{332C4427-26CB-11D0-B483-00C04FD90119}")
	IID_IWebBrowserApp := "{0002DF05-0000-0000-C000-000000000046}"
	iWebBrowser2 := COM_QueryService(pwin,IID_IWebBrowserApp,IID_IWebBrowserApp)
	; GuiControl,Text,WindowTitle,%	COM_Invoke(iWebBrowser2,"LocationName")
	GuiControl,Text,Title,%	COM_Invoke(iWebBrowser2,"LocationName")
	GuiControl,Text,URLL,% COM_Invoke(iWebBrowser2,"LocationURL")
	; GuiControl,Text,browserHeight,%	 COM_Invoke(iWebBrowser2,"height")
	; GuiControl,Text,browserWidth,%	 COM_Invoke(iWebBrowser2,"width")
	
	If   pelt := COM_Invoke(pwin , "document.elementFromPoint", xpos-xorg:=COM_Invoke(pwin ,"screenLeft"), ypos-yorg:=COM_Invoke(pwin ,"screenTop"))
	{
		framepath:=
		COM_Release(pwin)
		While   (type:=COM_Invoke(pelt,"tagName"))="IFRAME" || type="FRAME"
		{
			selt .=   "[" type "]." A_Index " **[sourceIndex]=" COM_Invoke(pelt,"sourceindex") " **[name]= " COM_Invoke(pelt,"name") " **[id]= " COM_Invoke(pelt,"id") "`n"
;vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
			LV_R%A_Index%_C1 := type "." A_Index
			LV_R%A_Index%_C2 := COM_Invoke(pelt,"sourceindex")
			LV_R%A_Index%_C3 := COM_Invoke(pelt,"name")
			LV_R%A_Index%_C4 := COM_Invoke(pelt,"id")
;^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^	
			framepath.=(COM_Invoke(pelt,"id") ? COM_Invoke(pelt,"id") :  COM_Invoke(pelt,"sourceindex")) ","
			pwin :=   COM_QueryService(pbrt:=COM_Invoke(pelt,"contentWindow"), "{332C4427-26CB-11D0-B483-00C04FD90119}"), COM_Release(pbrt), COM_Release(pdoc)
			pdoc :=   COM_Invoke(pwin, "document"), COM_Release(pwin)
			pbrt :=   COM_Invoke(pdoc, "elementFromPoint", xpos-xorg+=COM_Invoke(pelt,"getBoundingClientRect.left"), ypos-yorg+=COM_Invoke(pelt,"getBoundingClientRect.top")), COM_Release(pelt), pelt:=pbrt
		}

		pbrt :=   COM_Invoke(pelt, "getBoundingClientRect")
		l  :=   COM_Invoke(pbrt, "left")
		t  :=   COM_Invoke(pbrt, "top")
		r  :=   COM_Invoke(pbrt, "right")
		b  :=   COM_Invoke(pbrt, "bottom")
;vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
		global ToggleOutline
	If ToggleOutline { ; skip if Menu Item "Element Ouline" isn't selected
		global WinHWND,x1,y1,x2,y2,Fx1,Fy1,Fx2,Fy2,Resized,Stored_xorg,Stored_yorg
		If Resized
			Stored_xorg:=Stored_yorg:=""
		If(x1 <> l+xorg || y1 <> t+yorg || x2 <> r+xorg || y2 <> b+yorg) { ; if it's a different element
			If selt { ; if the element is in a frame, get frame dimensions
				Loop, Parse, framepath, `, ; loop framepath from above & insert code
					If A_LoopField ; prevent error if extra comma at the end
						frame_path .= "document.all[" A_LoopField "].contentWindow."
				StringTrimRight, frame_path, frame_path, 14
				FRect := COM_Invoke(iWebBrowser2, frame_path "getBoundingClientRect") ; get the Frame Rectangle
				Fx1:=xorg, Fy1:=yorg, Fx2:=COM_Invoke(FRect,"right")+xorg, Fy2:=COM_Invoke(FRect,"bottom")+yorg
				If(Fx2="" || Fy2="") { ; **if frame doesn't have "getBoundingClientRect" dimensions, use previous-level Frame
; **BUG** - "Stored_xorg" & "Stored_yorg" will empty unless the user has already hovered over a Frame that DID have "getBoundingClientRect" dimensions (& the window hasn't been resized)
; **Possible BUG** - if this situation occurs in the top-level Frame, or if the previous-level Frame has the same situation
					COM_Release(FRect)
					frame_path := RegExReplace(frame_path,"contentWindow\.document\.all\[.*?\]\.$") ; access previous level frame
					FRect := COM_Invoke(iWebBrowser2, frame_path "getBoundingClientRect") ; get the Frame Rectangle
					Fx1:=Stored_xorg, Fy1:=Stored_yorg, Fx2:=COM_Invoke(FRect,"right")+Stored_xorg, Fy2:=COM_Invoke(FRect,"bottom")+Stored_yorg
				} Else Stored_xorg:=xorg, Stored_yorg:=yorg
				COM_Release(FRect)
				If(Fx2="" || Fy2="") ; **if previous-level frame doesn't have "getBoundingClientRect" dimensions, set Frame dimensions as "NA"
					Fx1:=Fy1:=Fx2:=Fy2:="NA"
			} Else Fx1:=Fy1:=Fx2:=Fy2:="NA" ; if there isn't any frames, assign frame coords "NA"
			Outline( x1:=l+xorg, y1:=t+yorg, x2:=r+xorg, y2:=b+yorg )
		}
		WinHWND := COM_Invoke(iWebBrowser2, "HWND"), COM_Release(iWebBrowser2)
	}
/*
**Example situation - navigate to following link, hover over "Share This" (index 17), and hover over the pop-up Frame

http://www.google.com/imgres?imgurl=http://www.crystalinks.com/dragon.gif&imgrefurl=http://www.crystalinks.com/dragons.html&h=318&w=356&sz=23&tbnid=NNY7T-tm8bUUSM:&tbnh=108&tbnw=121&prev=/images%3Fq%3Ddragon&hl=en&usg=__wNcn-FaHPVFZJZM9HZvn8E4XtNQ=&ei=rjUTS7P5FcagnQeDyr3RAw&sa=X&oi=image_result&resnum=1&ct=image&ved=0CBAQ9QEwAA
*/
		static Stored_selt, Stored_textOfObj, Stored_outerHTML
;^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^	
		StringTrimRight,framepath,framepath,1
		If(Stored_selt <> selt) { ; added by jethrow
			LV_Delete()
			While( LV_R%A_Index%_C1 )
				LV_Add( "",LV_R%A_Index%_C1,LV_R%A_Index%_C2,LV_R%A_Index%_C3,LV_R%A_Index%_C4 )
			Stored_selt := selt ; added by jethrow
		}
		GuiControl,Text,theFrame,%	selt
		GuiControl,Text,EleIndex,%	sI:=COM_Invoke(pelt,"sourceindex")
		GuiControl,Text,EleName,%	sName:=COM_Invoke(pelt,"name")
		GuiControl,Text,EleIDs,%	sID:=COM_Invoke(pelt,"id")
		GuiControl,Text,Source,%	COM_Invoke(pdoc,"documentelement.outerhtml")
		global textOfObj
		textOfObj:=inpt(pelt)
		If(Stored_textOfObj <> textOfObj) {	; added by jethrow
			GuiControl, Text, html_value, % RegExReplace(textOfObj,"\[.*\]=")
			Stored_textOfObj := textOfObj	; added by jethrow
		}
		outerHTML := RegExReplace(COM_Invoke(pelt, "outerhtml"),"\[.*\]=") ; modified by jethrow
		If(Stored_outerHTML <> outerHTML) {	; added by jethrow
			GuiControl, Text, html_text, %outerHTML% ; modified by jethrow
			Stored_outerHTML := outerHTML	; added by jethrow
		}
		global GuiWinTitle, HTMLTag
		; GuiControl,Text,HTMLTag,%	COM_Invoke(pelt,"tagName")
		HTMLTag:=COM_Invoke(pelt,"tagName")	; added by jethrow
		WinSetTitle, % GuiWinTitle "ahk_class AutoHotkeyGUI",, % GuiWinTitle (HTMLTag ? " - [" HTMLTag "]":"")  ; added by jethrow
		innert:=COM_Invoke(pelt, "innerHTML")
		StringReplace, textOfObj, textOfObj,]=,?
		StringSplit,textOfObjs,textOfObj,?

		StringReplace,textOfObj,textOfObjs2,`,,&#44;,all
		optFrames:=framepath ? ", """ framepath """" : ""
		global element,optFrames
		element:=sID ? sID : sNames ? sName : sI
		optFrames:=framepath ? ", """ framepath """" : ""




		COM_Release(pbrt)
		COM_Release(pelt)

	}
	COM_Release(pdoc)
	Return
}

inpt(i)
{

	typ		:=	COM_Invoke(i,	"tagName")
	inpt	:=	"BUTTON,INPUT,OPTION,SELECT,TEXTAREA"
	Loop,Parse,inpt,`,
		if (typ	=	A_LoopField	?	1	:	"")
			Return "[value]=" COM_Invoke(i,	"value")
	Return "[innertext]=" COM_Invoke(i,	"innertext")
}

IE_GetDocument(hWnd)
{
   Static
   If Not   pfn
      pfn := DllCall("GetProcAddress", "Uint", DllCall("LoadLibrary", "str", "oleacc.dll"), "str", "ObjectFromLresult")
   ,   msg := DllCall("RegisterWindowMessage", "str", "WM_HTML_GETOBJECT")
   ,   COM_GUID4String(iid, "{00020400-0000-0000-C000-000000000046}")
   If   DllCall("SendMessageTimeout", "Uint", hWnd, "Uint", msg, "Uint", 0, "Uint", 0, "Uint", 2, "Uint", 1000, "UintP", lr:=0) && DllCall(pfn, "Uint", lr, "Uint", &iid, "Uint", 0, "UintP", pdoc:=0)=0
   Return   pdoc
}

;vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
Outline(x1,y1="",x2="",y2="") {
	GoSub, SetOutlineTransparent
	Loop, 4
		Gui, % A_Index+1 ": Hide"
	If x1 = Hide
		Return
		
	global WinHWND,Resized,Fx1,Fy1,Fx2,Fy2
	WinGetPos, Wx, Wy, , , ahk_id %WinHWND%
	ControlGetPos, Cx1, Cy1, Cw, Ch, Internet Explorer_Server1, ahk_id %WinHWND%
	Cx1+=Wx, Cy1+=Wy, Cx2:=Cx1+Cw, Cy2:=Cy1+Ch, Resized := False ; set "Internet Explorer_Server1" dimensions (set variable "Resized" as true)
	Final_x1:=Val(x1,Cx1,Fx1,">"), Final_y1:=Val(y1,Cy1,Fy1,">"), Final_x2:=Val(x2,Cx2,Fx2,"<"), Final_y2:=Val(y2,Cy2,Fy2,"<") ; set outline dimensions
	If ElemCoord(y1,Cy1,Fy1,">") ; TOP side of GUI outline
		Gui, 2:Show, % "NA X" Final_x1-2 " Y" Final_y1-2 " W" Final_x2-Final_x1+4 " H" 2,outline1
	If ElemCoord(x2,Cx2,Fx2,"<") ; RIGHT side of GUI outline
		Gui, 3:Show, % "NA X" Final_x2 " Y" Final_y1 " W" 2 " H" Final_y2-Final_y1,outline2
	If ElemCoord(y2,Cy2,Fy2,"<") ; BOTTOM side of GUI outline
		Gui, 4:Show, % "NA X" Final_x1-2 " Y" Final_y2 " W" Final_x2-Final_x1+4 " H" 2,outline3
	If ElemCoord(x1,Cx1,Fx1,">") ; LEFT side of GUI outline
		Gui, 5:Show, % "NA X" Final_x1-2 " Y" Final_y1 " W" 2 " H" Final_y2-Final_y1,outline4
	GoSub, SetOutlineLevel
	Return
}
SetOutlineTransparent:
	Loop, 4
		WinSet, Transparent, 0, % "ahk_id" outline%A_Index%
Return
SetOutlineLevel:
; http://www.autohotkey.com/forum/topic5672.html&highlight=getnextwindow
	hwnd_above := DllCall("GetWindow", "uint", WinHWND, "uint", 0x3) ; get window directly above "WinHWND"
	While(hwnd_above=outline1 || hwnd_above=outline2 || hwnd_above=outline3 || hwnd_above=outline4) ; don't use 4 AHK GUIs
		hwnd_above := DllCall("GetWindow", "uint", hwnd_above, "uint", 0x3)
; http://www.autohotkey.com/forum/topic22763.html&highlight=setwindowpos
	Loop, 4 { ; set 4 "outline" GUI's directly below "hwnd_above"
		DllCall("SetWindowPos", "uint", outline%A_Index%, "uint", hwnd_above
			, "int", 0, "int", 0, "int", 0, "int", 0
			, "uint", 0x13) ; NOSIZE | NOMOVE | NOACTIVATE ( 0x1 | 0x2 | 0x10 )
		WinSet, Transparent, 255, % "ahk_id" outline%A_Index% ; set outline GUIs visible
	}
Return
Val(E,C,F,option=">") {
	If F is digit
		Return, option=">" ? (E>=C ? (E>=F ? E:F) : (C>=F ? C:F)) : (E<=C ? (E<=F ? E:F) : (C<=F ? C:F))
	Else Return, option=">" ? (E>=C ? E:C) : (E<=C ? E:C)
}
ElemCoord(E,C,F,option=">") {
	If F is digit
		Return, option=">" ? (E>=C && E>=F ? 1:0):(E<=C && E<=F ? 1:0)
	Else Return, option=">" ? (E>=C ? 1:0):(E<=C ? 1:0)
}
;^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^	



NewWindow:
{
	Gui,Submit,NoHide

	initString=
(
iWeb_Init()
pwb:=iWeb_newIe()
iWeb_nav("%URLL%")
)
	writeScript(initString)
	Return
}
InitWindow:
{
	Gui,Submit,NoHide
	initString=
(
iWeb_Init()
pwb:=iWeb_getwin("%Title%")
)
	writeScript(initString)


	Return
}

writeScript(code)
{
	global
	Gui,Submit,NoHide
	GuiControl,text,CodeWindow,% CodeWindow "`n" code
	Return



}

getURL(t)
{
	If	psh	:=	COM_CreateObject("Shell.Application") {
		If	psw	:=	COM_Invoke(psh,	"Windows") {
			Loop, %	COM_Invoke(psw,	"Count")
				If	url	:=	(InStr(COM_Invoke(psw,"Item[" A_Index-1 "].LocationName"),t) && InStr(COM_Invoke(psw,"Item[" A_Index-1 "].FullName"), "iexplore.exe")) ? COM_Invoke(psw,"Item[" A_Index-1 "].LocationURL") :
					Break
			COM_Release(psw)
		}
		COM_Release(psh)
	}
	Return	url
}



AddToScript:
{

	Gui, Submit, NoHide
	if InStr(Func,"List of functions")
		return
	StringReplace, Func, Func, `(`)
	 , % (RegExMatch(Func,"(?:Init|Term|newIe)") ? "()"
	 : (RegExMatch(Func,"(?:Release|Complete)") ? "(pwb)"
	 : (InStr(Func,"DomWin") ? "(pwb,"""")"
	 : (InStr(Func,"getWin") ? "(""" Title """)"
	 : (InStr(Func,"Nav") ? "(""" URLL """)"
	 : (InStr(Func,"setDomObj") ? "(pwb,"""","""")"
	 : "(pwb,"""")"))))))
	if sClip
		Clipboard:=InStr(Func,"getWin") ? "pwb:=" Func : Func
	else
		GuiControl, Text, CodeWindow, % (!CodeWindow || RegExMatch(CodeWindow,"\v+$"))
		 ? CodeWindow (InStr(Func,"getWin") ? "pwb:=" Func : Func)
		 : CodeWindow "`n" (InStr(Func,"getWin") ? "pwb:=" Func : Func)
	return
}

AddJscript:
{

	Gui,Submit,NoHide
	js=
	(
js=`n(`n%Javascript%`n)
iWeb_execScript(pwb,js %optFrames%)
	)
	writeScript(js)
	GuiControl,text,Javascript,
	Return
}

TestJscript:
{

	Gui,Submit,NoHide
	pwb:=iWeb_getwin(Title)
	iWeb_execScript(pwb,Javascript %optFrames%)
	COM_Release(pwb)
	Return
}

RunScript:

Gui,Submit,NoHide
if !ErrCheckPwb() || !ErrCheckInitTerm() || !ErrCheckBounds()
	return
pipe_name := "iWebBrowser2 Script Writer"
pipe_ga := CreateNamedPipe(pipe_name, 2)
pipe    := CreateNamedPipe(pipe_name, 2)
if (pipe=-1 or pipe_ga=-1) {
    MsgBox CreateNamedPipe failed.
    return
}
Run, %A_AhkPath% "\\.\pipe\%pipe_name%"
DllCall("ConnectNamedPipe","uint",pipe_ga,"uint",0)
DllCall("CloseHandle","uint",pipe_ga)
DllCall("ConnectNamedPipe","uint",pipe,"uint",0)
Script := chr(239) chr(187) chr(191) CodeWindow
if !DllCall("WriteFile","uint",pipe,"str",Script,"uint",StrLen(Script)+1,"uint*",0,"uint",0)
    MsgBox WriteFile failed: %ErrorLevel%/%A_LastError%
DllCall("CloseHandle","uint",pipe)
Return


SaveScript:
{


	Gui,Submit,NoHide
	FileSelectFile,script
	FileDelete,%script%
	FileAppend,%CodeWindow%,%script%
	Return
}

CopyToWriter:
Gui,default	
Gui, Submit, NoHide
if (A_GuiEvent<>"DoubleClick")
	return
if iClip
	Clipboard:=iWeb
else
	GuiControl, Text, CodeWindow
	 , % (!CodeWindow ? "iWeb_Init()`npwb:=iWeb_getWin(""" Title """)`n" iWeb
	 : (RegExMatch(CodeWindow,"\v$") ? CodeWindow iWeb
	 : CodeWindow "`n" iWeb))
WinActivate, %GuiWinTitle%
TabActivate(1)
return

IsTab:
Gui, Submit, NoHide
if (currentTab<>1) {
	if (currentTab=2) {
		if ClickToClip {
			Menu, OptionsMenu, ToggleCheck, Left Click/Copy
			ClickToClip:=!ClickToClip
		}
		if !ToggleOutline {
			Menu,OptionsMenu,ToggleCheck,Element Outline
			ToggleOutline:=!ToggleOutline
		}
	}
	else {
		if !ClickToClip {	
			Menu, OptionsMenu, ToggleCheck, Left Click/Copy
			ClickToClip:=!ClickToClip
		}			
		if ToggleOutline {
			Menu,OptionsMenu,ToggleCheck,Element Outline
			ToggleOutline:=!ToggleOutline
			Outline("Hide")
		}
	}
}
else if (prevTog<>ToggleOutline) {
	Menu,OptionsMenu,ToggleCheck,Element Outline
	ToggleOutline:=!ToggleOutline
	Outline("Hide")
}
else if !ClickToClip {
	Menu, OptionsMenu, ToggleCheck, Left Click/Copy
	ClickToClip:=!ClickToClip
}
WinMove, %GuiWinTitle%, , , , , % InStr(mintab,CurrentTab) ? 303 : 562
return

GetData:
Pos:=1
Gui, Submit, NoHide
While Pos:=RegExMatch(Source,"is)(?P<t><FORM.*?>).*?</FORM>",o,Pos+StrLen(o))	;	"is)<FORM.*? (?:name|id)=\s?(\w+).*?>.*?</FORM>"
{
	RegExMatch(ot," (?:name|id)=\s?(?P<ni>\w+)",m)
	c:=A_Index+1
	flist.=(A_Index=1 ? mni "`n" : "`n" mni), f%c%:=o	;	flist.=(A_Index=1 ? m1 "`n" : "`n" m1), f%A_Index%:=m
}
GuiControl, , SelForm, % (flist ? "`nAvailable forms`n" flist : "`n")
GuiControl, , Form, % (flist ? f1 : "")
VarSetCapacity(flist,0)
return

LoadForm:
Gui, Submit, NoHide
GuiControl, , Form, % f%SelForm%
return

Get_forms(fSelForm) {

	global formname,optFrames,SelForm,Title
	pdoc:=iWeb_Txt2Doc(fSelForm)
	formname:=COM_Invoke(pdoc,"forms[0].elements.name") ? COM_Invoke(pdoc,"forms[0].elements.name") : SelForm-1
	iwebstring=
	(
	/*
****Disclaimer:
This code is experimental and should not be relied upon without thorough testing


*/
iWeb_Init()
pwb:=iWeb_getwin("%Title%")
	)
	
	Loop % COM_Invoke(pdoc,"forms[0].elements.length")
	{
		ordinal:=A_Index-1
		tagname:=COM_Invoke(pdoc,"forms[0].elements[" ordinal "].tagname")
		selectedindex:=
		Checked	:=
		If	tagname=select
			selectedindex:=COM_Invoke(pdoc,"forms[0].elements[" ordinal "].selectedindex")
		Else	type:=COM_Invoke(pdoc,"forms[0].elements[" ordinal "].type")
		If	(type="radio" || type="checkbox")
			Checked	:=COM_Invoke(pdoc,"forms[0].elements[" ordinal "].checked") ? 1 : 0
		ID:=COM_Invoke(pdoc,"forms[0].elements[" ordinal "].id")
		name:=COM_Invoke(pdoc,"forms[0].elements[" ordinal "].name")
		value:=COM_Invoke(pdoc,"forms[0].elements[" ordinal "].value")
		elementRef:=ID ? ID : name ? name : ordinal
		element.=elementRef "=" value "`n"  ;; might one day use this to create a simple postdata string
		StringReplace,value,value,`,,&#44;,all	;	escape all commas in text extracted always
		StringReplace,elementRef,elementRef,`,,&#44;,all	;	escape all commas in text extracted always
		If	(StrLen(selectedindex) || StrLen(Checked))
			iwebstring.= StrLen(selectedindex) ? "iWeb_SelectOption(pwb,""" elementRef """," selectedindex optFrames ")" : StrLen(Checked) ? "iWeb_Checked(pwb,""" elementRef """," Checked optFrames ")" : ""
		Else	If	(type <> "button" &&  type <> "submit" && type <> "reset")
		{
			values.=value ","
			elementRefs.=elementRef ","
		}
		iwebstring.= "`n"
	}
	If	elementRefs
		iwebstring.="iWeb_setDomObj(pwb,""" elementRefs """,""" values """" optFrames ")`n"
	footer=
	(
COM_Invoke(pWin:=iWeb_DomWin(pwb%optFrames%),"document.forms[%formname%].submit")
iWeb_Release(pWin)
iWeb_Release(pwb)
iWeb_Term()	
	)
	iwebstring.=footer
	Loop,Parse,iwebstring,`n
		If	A_LoopField
			iwebstrings.=A_LoopField "`n"
;~ 	MsgBox	% iwebstrings
	COM_Release(pdoc)
	Return	iwebstrings
}

FormToClip:
Gui, Submit, NoHide
Clipboard:=Form
return

MenuLeftClick:
Menu, OptionsMenu, ToggleCheck, Left Click/Copy
ClickToClip:=!ClickToClip
return

MenuDebug:
Debug:=!Debug
Menu, HelpMenu, ToggleCheck, &Debug with ListLines
ListLines, % Debug ? "On" : "Off"
return

MenuAbout:
TabActivate(5)
return

iWebacc_Query(pacc, bunk = "")
{
	If	DllCall(NumGet(NumGet(1*pacc)+0), "Uint", pacc, "Uint", COM_GUID4String(IID_IAccessible,bunk ? "{00020404-0000-0000-C000-000000000046}" : "{618736E0-3C3D-11CF-810C-00AA00389B71}"), "UintP", pobj)=0
		DllCall(NumGet(NumGet(1*pacc)+8), "Uint", pacc), pacc:=pobj
	Return	pacc
}

iWebacc_AccessibleObjectFromPoint(x = "", y = "", ByRef _idChild_ = "")
{
	VarSetCapacity(varChild,16,0)
	x<>""&&y<>"" ? pt:=x&0xFFFFFFFF|y<<32 : DllCall("GetCursorPos", "int64P", pt)
	DllCall("oleacc\AccessibleObjectFromPoint", "int64", pt, "UintP", pacc, "Uint", &varChild)
	_idChild_ := NumGet(varChild,8)
	Return	pacc
}

iWebacc_Child(pacc, idChild)
{
	If	DllCall(NumGet(NumGet(1*pacc)+36), "Uint", pacc, "int64", 3, "int64", idChild, "UintP", paccChild)=0 && paccChild
	Return	iWebacc_Query(paccChild)
}

iWebacc_Name(pacc, idChild = 0)
{
	If	DllCall(NumGet(NumGet(1*pacc)+40), "Uint", pacc, "int64", 3, "int64", idChild, "UintP", pName)=0 && pName
	Return	COM_Ansi4Unicode(pName) . SubStr(COM_SysFreeString(pName),1,0)
}

iWebacc_Value(pacc, idChild = 0)
{
	If	DllCall(NumGet(NumGet(1*pacc)+44), "Uint", pacc, "int64", 3, "int64", idChild, "UintP", pValue)=0 && pValue
	Return	COM_Ansi4Unicode(pValue) . SubStr(COM_SysFreeString(pValue),1,0)
}

iWebacc_Role(pacc, idChild = 0)
{
	VarSetCapacity(var,16,0)
	If	DllCall(NumGet(NumGet(1*pacc)+52), "Uint", pacc, "int64", 3, "int64", idChild, "Uint", &var)=0
	Return	iWebacc_GetRoleText(NumGet(var,8))
}

iWebacc_State(pacc, idChild = 0)
{
	VarSetCapacity(var,16,0)
	If	DllCall(NumGet(NumGet(1*pacc)+56), "Uint", pacc, "int64", 3, "int64", idChild, "Uint", &var)=0
	Return	iWebacc_GetStateText(nState:=NumGet(var,8)) . "`t(" . iWebacc_Hex(nState) . ")"
}

iWebacc_GetRoleText(nRole)
{
	nSize := DllCall("oleacc\GetRoleTextA", "Uint", nRole, "Uint", 0, "Uint", 0)
	VarSetCapacity(sRole, nSize)
	DllCall("oleacc\GetRoleTextA", "Uint", nRole, "str", sRole, "Uint", nSize+1)
	Return	sRole
}

iWebacc_GetStateText(nState)
{
	nSize := DllCall("oleacc\GetStateTextA", "Uint", nState, "Uint", 0, "Uint", 0)
	VarSetCapacity(sState, nSize)
	DllCall("oleacc\GetStateTextA", "Uint", nState, "str", sState, "Uint", nSize+1)
	Return	sState
}

iWebacc_Hex(num)
{
	old := A_FormatInteger
	SetFormat, Integer, H
	num += 0
	SetFormat, Integer, %old%
	Return	num
}

CreateNamedPipe(Name, OpenMode=3, PipeMode=0, MaxInstances=255) {
    return DllCall("CreateNamedPipe","str","\\.\pipe\" Name,"uint",OpenMode
        ,"uint",PipeMode,"uint",MaxInstances,"uint",0,"uint",0,"uint",0,"uint",0)
}

GuiHeight() {
	global GuiWinTitle
	WinGetPos, , , , h, %GuiWinTitle%
	return h
}

TabActivate(no) {
	global GuiWinTitle
	SendMessage, 0x1330, %no%,, SysTabControl321, %GuiWinTitle%
	Sleep 50
	SendMessage, 0x130C, %no%,, SysTabControl321, %GuiWinTitle%
	return
}

WinGetClass(ID) {
	WinGetClass, res, ahk_id %ID%
	return res
}

ErrCheckInitTerm() {
	global CodeWindow
	i:=j:=errlvl:=0
	Loop, Parse, CodeWindow, `n
	{
		if InStr(A_LoopField,"iWeb_Init()") {
			++i
			init%i%:=A_Index
		}
		if InStr(A_LoopField,"iWeb_Term()") {	
			++j
			term%j%:=A_Index
		}
	}
	if (i <> j) {
		GuiControl,1: Text, CodeWindow
		 , % CodeWindow "`niWeb_Term()" 
		Gui,Submit,NoHide
;~ 		MsgBox, 48, iWebBrowser2 Script Writer Warning, % "Script Writer has detected that you do not have an equal number of:`n`n"
;~ 			. "`tiWeb_Init() and iWeb_Term() statements`n`n"
;~ 			. "Please double check your code and try again."
		return 1
	}
	Loop % i {
		if (init%A_Index% > term%A_Index%) {
			errlvl=1
			Break
		}
	}
	if errlvl {
		MsgBox, 48, iWebBrowser2 Script Writer Warning, % "Script Writer has detected that the following sequences are not in order:`n`n"
			. "`tiWeb_Init() and iWeb_Term() statements`n`n"
			. "Please double check your code and try again."
		return 0
	}
	return 1
}

ErrCheckPwb() {
	global CodeWindow
	i:=j:=errlvl:=0
	Loop, Parse, CodeWindow, `n
	{
		if RegExMatch(A_LoopField,"pwb:=iWeb_getWin\(.*?\)") {
			++i
			gwin%i%:=A_Index
		}
		if InStr(A_LoopField,"iWeb_Release(pwb)") {
			++j
			crls%j%:=A_Index
		}
	}
	if (i <> j) {
		GuiControl,1: Text, CodeWindow
		 , % CodeWindow "`niWeb_Release(pwb)" 
		Gui,Submit,NoHide
;~ 		MsgBox, 48, iWebBrowser2 Script Writer Warning, % "Script Writer has detected that you do not have an equal number of:`n`n"
;~ 			. "`tiWeb_getWin() and iWeb_Release() statements`n`n"
;~ 			. "Please double check your code and try again."
		return 1
	}
	Loop % i {
		if (gwin%A_Index% > crls%A_Index%) {
			errlvl=1
			Break
		}
	}
	if errlvl {
		MsgBox, 48, iWebBrowser2 Script Writer Warning, % "Script Writer has detected that the following sequences are not in order:`n`n"
			. "`tiWeb_getWin() and iWeb_Release() statements`n`n"
			. "Please double check your code and try again."
		return 0
	}
	return 1
}

ErrCheckBounds() {
	global CodeWindow
	ipos:=InStr(CodeWindow,"iWeb_Init"), tpos:=InStr(CodeWindow,"iWeb_Term")
	Pos:=1,errlvl=0
	While Pos:=RegExMatch(CodeWindow,"(?:iWeb_getWin\(.*?\)|iWeb_Release\(pwb\))",p,Pos+StrLen(p))
	{
		chk:=Pos+StrLen(p)
		if chk not between %ipos% and %tpos%
		{
				errlvl=1
				break
		}
	}
	if errlvl {
		MsgBox, 48, iWebBrowser2 Script Writer Warning, % "Script Writer has detected that your iWeb_getWin() and iWeb_Release() statements`n"
		. "         are not properly bound between iWeb_Init() and iWeb_Term() statements.`n`n"
		. "Please double check your code and try again."
		return 0
	}
	return 1
}

;	credit below to toralf
;	http://www.autohotkey.com/forum/topic8976.html
WM_LBUTTONDOWN(wParam, lParam, msg, hwnd){       ;Copy-On-Click for controls
    global 

	Gui, Submit, NoHide
	If !ClickToClip
		Return
    If A_GuiControl is space                     ;Control is not known
        Return
	If InStr(Copyable,A_GuiControl) {
		Clipboard:=%A_GuiControl%
		if !Clipboard
			return
        ToolTip("Contents copied to Clipboard.`n" (StrLen(Clipboard) > 25 ? SubStr(Clipboard,1,25) "..." : Clipboard))
		return
	}
	return
}

SubListView:
If(A_GuiEvent = "Normal") && ClickToClip { ; if Right Clicked
	LV_GetText(LVselection, A_EventInfo, column_num)
	If LVselection ; if listview item contains data
		ToolTip("Contents copied to Clipboard.`n" (StrLen(Clipboard:=LVSelection) > 25 ? SubStr(Clipboard,1,25) "..." : Clipboard))
}
Return

ToolTip(Text, TimeOut=2000){
    ToolTip, %Text%
    SetTimer, RemoveToolTip, %TimeOut%
    Return
}
RemoveToolTip:
ToolTip
Return
getTables:

if OnTop {
	OnTop:=!OnTop, prevOnTop=1
	Menu, OptionsMenu, ToggleCheck, Always on Top
	Gui, % !OnTop ? "-AlwaysOnTop" : "+AlwaysOnTop"
}
vgui=
;~ (LTrim
;~ 	`#NoEnv
;~ 	DetectHiddenWindows,on
;~ 	SetTitleMatchMode,2
;~ 	SetControlDelay, -1
	Gui,Submit,NoHide
	Gui,99: Add, ListView, r15 w500 ggetTableData, Table|Row|Cell|Data
;~ 	ControlGetText, Title, Edit1, iWebBrowser2 Learner
;~ 	iWeb_Init()
;~ MsgBox	% Title
	Gui 99:Default 
	pwbtable:=iWeb_getWin(Title)
	table=0
	col:=COM_Invoke(pwbtable,"document.all")
	iWeb_Release(pwbtable)
	Loop % l:=COM_Invoke(col,"tags[table].length")
	{
	  row=0
	  Loop % l:=COM_Invoke(col,"tags[table].item[" table "].rows.length")
	  {
	    cell=0
	    Loop % l:=COM_Invoke(col,"tags[table].item[" table "].rows[" row "].cells.length")
	    {
	      LV_Add("",table,row,cell
	       ,COM_Invoke(col,"tags[table].item[" table "].rows[" row "].cells[" cell "].innerText"))
	      cell++
	    }
	    row++
	  }
	  table++
	}
	iWeb_Release(col)
	Gui,99: Show, Center, iWebBrowser2 Table Data View
	return

	getTableData:

	if A_GuiEvent=DoubleClick
	{
	  ControlGet, LV_Data, List, Focused, SysListView321, iWebBrowser2 Table Data View
	  StringSplit, i, LV_Data, %A_Tab%
	  call:="iWeb_tableParse(pwb," i1 "," i2 "," i3 ")"
	  newcode:="tabledata%" i1 i2 i3 "%:=iWeb_tableParse(pwb," i1 "," i2 "," i3 ")"
	  ToolTip % iweb
;~ 	  clipboard	:=iweb
		Gui,1:default	
		Gui,1:Submit,NoHide
	  GuiControl,1: Text, CodeWindow
		 , % (!CodeWindow ? "iWeb_Init()`npwb:=iWeb_getWin(""" Title """)`n" newcode
		 : (RegExMatch(CodeWindow,"\v$") ? CodeWindow newcode
		 : CodeWindow "`n" newcode))
;~ 	  GuiControl, Text, CodeWindow,% CodeWindow "`n" iWeb
		WinActivate, %GuiWinTitle%
		TabActivate(1)
	  
		Gui,99:default	
	  ToolTip
	}
	return
	
	99GuiCancel:
	99GuiClose:
	Gui,99: Destroy
	Gui,Default 
;~ 	ExitApp
		
;~ )
;~ pipe_name := "iWebBrowser2 Script Writer"
;~ pipe_ga := CreateNamedPipe(pipe_name, 2)
;~ pipe    := CreateNamedPipe(pipe_name, 2)
;~ if (pipe=-1 or pipe_ga=-1) {
;~     MsgBox CreateNamedPipe failed.
;~     return
;~ }
;~ Run, %A_AhkPath% "\\.\pipe\%pipe_name%"
;~ DllCall("ConnectNamedPipe","uint",pipe_ga,"uint",0)
;~ DllCall("CloseHandle","uint",pipe_ga)
;~ DllCall("ConnectNamedPipe","uint",pipe,"uint",0)
;~ Script := chr(239) chr(187) chr(191) vgui
;~ if !DllCall("WriteFile","uint",pipe,"str",Script,"uint",StrLen(Script)+1,"uint*",0,"uint",0)
;~     MsgBox WriteFile failed: %ErrorLevel%/%A_LastError%
;~ DllCall("CloseHandle","uint",pipe)
;~ WinWait, iWebBrowser2 Table Data View
;~ Loop {
;~   if !WinExist("iWebBrowser2 Table Data View") {
;~     if prevOnTop {
;~       OnTop:=!OnTop, prevOnTop=0
;~       Menu, OptionsMenu, ToggleCheck, Always on Top
;~       Gui, % !OnTop ? "-AlwaysOnTop" : "+AlwaysOnTop"
;~     }
;~     break
;~   }
;~   Sleep, 1000
;~ }		
Return


imagesourcemap:
{
	theGui=98
	Gui,Submit,NoHide
	Gui,%theGui%: Add, ListView, r15 w500 ggetimagesourcemap, Image|Ordinal|ID|Name|Source URL
	Gui %theGui%:Default 
	pwbtable:=iWeb_getWin(Title)
	img=0
	col:=COM_Invoke(pwbtable,"document.images")
	iWeb_Release(pwbtable)
	Loop % l:=COM_Invoke(col,"length")
	{
	  LV_Add("",img,COM_Invoke(col,"item[" img "].sourceIndex"),COM_Invoke(col,"item[" img "].id"),COM_Invoke(col,"item[" img "].name"),COM_Invoke(col,"item[" img "].src"))
	  img++
	}
	iWeb_Release(col)
	Gui,%theGui%: Show, Center, iWebBrowser2 Table Data View
	return

	getimagesourcemap:

	if A_GuiEvent=DoubleClick
	{
	  ControlGet, LV_Data, List, Focused, SysListView321, iWebBrowser2 Table Data View
	  StringSplit, i, LV_Data, %A_Tab%
	  newcode:="imgSource%" i1 "%:=iWeb_tableParse(pwb," i1 "," i2 "," i3 ")"
		Gui,1:default	
		Gui,1:Submit,NoHide
		GuiControl,1: Text, CodeWindow
		 , % (!CodeWindow ? "iWeb_Init()`npwb:=iWeb_getWin(""" Title """)`n" newcode
		 : (RegExMatch(CodeWindow,"\v$") ? CodeWindow newcode
		 : CodeWindow "`n" newcode))
		WinActivate, %GuiWinTitle%
		TabActivate(1)
	  
		Gui,%theGui%:default	
	}
	return
	
	%theGui%GuiCancel:
	%theGui%GuiClose:
	Gui,%theGui%: Destroy
	Gui,1:Default 
	Return
}

linkmap:
{
	
	Return
}
formmap:
{
	
	Return
}
htmlmap: 
{
	
	Return
}
tablemap:
{
	Gosub,getTables
	Return
}


