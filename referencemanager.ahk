


;==========================================================================
;==========================================================================
; Reference Manager (comandiers the key shared with the tilde)
;Adds citation to reference history and makes parsed citation available for pasting
;==========================================================================
;==========================================================================
	
	`:: 
		ifwinexist, Microsoft Excel - referencehistory.csv
			{
				msgbox, Close referencehistory.csv and try again
				winactivate, Microsoft Excel - referencehistory.csv
				return
			}
		referenceCB := getText() 
		if StrLen(referenceCB) > 0
			{
				AppendToReferences(referenceCB)
			}
	return

; just copies to my reference clipboard thing

	#`::
		referenceCB := getText()
		referenceparser(referenceCB)	
		tooltip, saved to clipboard "%referenceCB%..."
		SetTimer, tooltiptoggle, 2500
	return


;------------------------------------------------------------------------------
; Paste Author(s), Year, Title, Journal, (i.e., full reference)

	^`:: 
		keywait, ^
		sendinput %raw%

	return

;------------------------------------------------------------------------------
; Author(s) & Year only ; parenthetic citation. EX: (Ruchty et al., 2009)

	^#`:: 
		keywait, # 
		keywait, ^
		if strlen(authorname) > 0 & Strlen(DateYear) > 0
		{
			IfInString, raw, &
				newcite = (%AuthorName% et al., %DateYear%)
			else
				newcite = (%AuthorName%, %DateYear%)
			sendinput %newcite% 
		}
		sendinput, {LWIN up}{RWIN up}{CTRL up}
	return

;------------------------------------------------------------------------------
; Author(s) & Year only; in-text citation, EX:  Ruchty et al. (2009)

	!#`:: 

		ifwinactive, , Reference History Formatted.xlsm
		winclose
		else
			{
				ifwinnotexist, , Reference History Formatted.xlsm
				run, C:\Users\Josh\Dropbox\Research\Reference History Formatted.xlsm
				else
				winactivate
			}
	return
		

	/* - this used to run inserting an in-text citation

		keywait, !
		keywait, #
		if strlen(authorname) > 0 & Strlen(DateYear) > 0
			{
				IfInString, raw, &
					newcite = %AuthorName% et al. (%DateYear%)
				else
					newcite = %AuthorName% (%DateYear%)
				sendinput %newcite%
			}
	return 
	*/ 

;------------------------------------------------------------------------------
; Author(s), Year, and Title only (replaces system characters, e.g., COLONS)

	^+`:: 
		keywait, ^
		keywait, +
		;combine the data, evaluate multiple authors (with "&"), send to keyboard
			if raw > 0
				{
					IfInString, raw, &
						;newcite = %AuthorName% et al. (%DateYear%) - %journalname% - %fileTitle%
						newcite = %AuthorName% et al. (%DateYear%) - %fileTitle%
					else
						;newcite = %AuthorName% (%DateYear%) - %journalname% - %fileTitle%
						newcite = %AuthorName% (%DateYear%) - %fileTitle%
					sendinput %newcite%
				}
	return


	;------------------------------------------------------------------------------
	; open reference history

	!`:: 
		
		keywait, ``
		keywait, !
		
		ifwinexist, Microsoft Excel - referencehistory.csv
			{
				ifwinnotactive, Microsoft Excel - referencehistory.csv
					{
						winactivate, Microsoft Excel - referencehistory.csv
						return
					}
				ifwinactive, Microsoft Excel - referencehistory.csv
					{
						sendinput, ^s
						winwait, ahk_class #32770^
						sendinput, {ENTER}
						winwait, ahk_class XLMAIN
						sendinput, ^w
						winwait, ahk_class NUIDialog
						sendinput, n
						sleep, 100
						; CLOSES EXCEL IF NO OTHER WORKBOOK OPEN
						exceltitle := ""
						ifwinexist, Microsoft Excel
						WinGetTitle, exceltitle
						if exceltitle = Microsoft Excel
						winClose
						return
						
					}
			}

		ifwinnotexist, Microsoft Excel - referencehistory.csv
			{
				run, referencehistory.csv ; 
				winwait, Microsoft Excel - referencehistory.csv
    			sendinput, ^{HOME}^{HOME}!wfr
    			sleep, 100
    			sendinput, ^gA1:A10000{ENTER}
    			sendinput, ^g!s!k{ENTER}{Appskey}dr{enter}
				sleep, 50
				sendinput, ^ga10000{enter}
				sleep, 50
				sendinput, {end}{up}
				Loop 5
	    			Click Wheelup
    			sleep, 50

				return
			}
	return


;------------------------------------------------------------------------------
;Name:       masterSave         



	;Save highlighted text to MasterSave CSV document
	^#c:: MasterStore()

	; Intelligently runs, activates, or closes MasterSave
	#!c::

		keywait, #
		keywait, !
		keywait, c
		
		ifwinnotexist, Microsoft Excel - masterSave.CSV
			{
				run, masterSave.csv ; (e.g., the contents of master save)
				winwait, Microsoft Excel - masterSave.CSV
				sleep, 100
				sendinput, ^g
				sendinput, a10000{enter}
				sleep, 50
				sendinput, {end}{up}
				Loop 5
	    			Click Wheelup
				return
			}

		ifwinexist, Microsoft Excel - masterSave.CSV
			{
				ifwinactive, Microsoft Excel - masterSave.CSV
					{
						sendinput, ^s
						winwait, ahk_class #32770
						sendinput, {ENTER}n
						winwait, ahk_class XLMAIN
						sendinput, ^w
						winwait, ahk_class NUIDialog
						sendinput, n
						sleep, 100
						; CLOSES EXCEL IF NO OTHER WORKBOOK OPEN
						exceltitle := ""
						ifwinexist, Microsoft Excel
						WinGetTitle, exceltitle
						if exceltitle = Microsoft Excel
						winClose
						return
					}

				ifwinnotactive, Microsoft Excel - masterSave.CSV
					{
						winactivate, Microsoft Excel - masterSave.CSV
						return
					}
				
			}

	return


;▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒
;   FUNCTIONS stored for retrieval
;▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒


;------------------------------------------------------------------------------
; get Text. To use, call the function's contents to a variable. for example, mysavedtext := getText()

getText() ; 
	{
		global
		tempCB := 0
		tempCB := clipboard ; 					begins loading clipboard contents into tempCB for storage
		storageloopcount := 0
		while, StrLen(tempCB) = 0 ; 		keep checking the stringlength of the tempCB to see when the clipboard has been loaded. exit when done.
			{
				storageloopcount++
				sleep, 10
				if storageloopcount > 300
					{
						break
					}
			}

		clipboard := 
		sendinput, ^c
		loadtextloopcount := 0
		while, StrLen(clipboard) = 0 ; 			this loop checks the stringlength of clipboard to wait for it to load after pressing ^c
			{
				loadtextloopcount++
				sleep, 10
				if loadtextloopcount > 200
					{
						;msgbox, timed out loading highlighted text to clipboard (ctrl-c) (> 2000 ms). script aborted
						return
					}
			}

		
		taskvariable := 
		taskvariable := clipboard ; 				begins tranferring the catured text to taskvariable for manipulation
		taskVloopcount := 0
		while, StrLen(taskvariable) = 0
			{
				taskVloopcount++
				sleep, 10
				if taskVloopcount > 200
					{
						msgbox, timed out loading highlighted text on clipboard to task variable (> 2000 ms). script aborted
						return
					}
			}
		clipboard :=
		clipboard := tempCB ; there could be a lag here, troubleshoot
		if tempCB > 0
		clipwait, 5
		if ErrorLevel = 1
			{
				msgbox, failed
				return
			}
		tempCB :=
		return taskvariable
	} 




;------------------------------------------------------------------------------
;Saving references to reference history using FileAppend. The text in each column must be surrounded by parentheses (can still resolve variables). The literal comma (`,) separates columns but commas in the normal text do not. use normal comma prior to filename
;The first 10 columns are reserved for the reference information
;The next 20 columns (11-30) are reserved for topics
;The next 20 columns (31-50) are reserved for relevant projects
;The next 20 columns (51-70) are reserved for domain

	AppendToReferences(what)
		{
			global
			referenceparser(what)
			GuiControl, , abstract ; these lines remove the text previously entered into the abstract and comments text edit boxes
			;GuiControl, , CiteComments
			fullreference := getText()
			if SavedGUI <> 1 ; clear button = savedgui = 0, "close" or "ok" = saved gui = 1
				{
					Gui, Destroy ; deletes GUI controls to reload them fresh. I believe leaves former variables intact

					;left column topics (csv columns 11-30)
					Gui, Add, GroupBox,  x20 y20 w250 h200, Topics
					Gui, Add, Checkbox, vc11 x40 y40 section, &Empathy ; remember, the v designates it as the variable being created (e.g., variable name is "c11")
					Gui, Add, Checkbox, vc12, &Persuasion ; and the "&" in the text facilitates the alt-letter jump 
					Gui, Add, Checkbox, vc13, Emo&tion
					Gui, Add, Checkbox, vc14, &Ingroup/outgroup
					Gui, Add, Checkbox, vc15, &Morality
					Gui, Add, Checkbox, vc16, &Status
					Gui, Add, Checkbox, vc17, Sel&fPerception
					Gui, Add, Checkbox, vc18, &Neuro
					Gui, Add, Checkbox, vc19, &Attitudes
					Gui, Add, Checkbox, vc20 ys, Arousal/threat 
					Gui, Add, Checkbox, vc21, Social
					;etc. up to and including c30

					;Projects  (csv columns 31-50)
					Gui, Add, GroupBox,  x290 y20 w140 h200, Projects
					Gui, Add, Checkbox, vc31 x300 y40 section, &2nd year paper
					Gui, Add, Checkbox, vc32, E&xclusion
					Gui, Add, Checkbox, vc33, Li&z
					Gui, Add, Checkbox, vc34, Emotion && Fee&dback
					;Gui, Add, Checkbox, vc35, Display Text
					;Gui, Add, Checkbox, vc36, Display Text
					;etc. up to 50


					;Domain (csv columns 51-70)
					Gui, Add, GroupBox,  x450 y20 w150 h200, Domain					
					Gui, Add, Checkbox, vc51 x460 y40, &Stats
					Gui, Add, Checkbox, vc52, &Methods
					Gui, Add, Checkbox, vc53, &Writing && publishing
					Gui, Add, Checkbox, vc54, &Career
					Gui, Add, Checkbox, vc55, Measu&re
					;etc. up to 50

					;abstract
					Gui, Add, text, x20 y230, Abstract:
					Gui, Add, Edit, x20 w580 r10 vabstract

					;Comments
					Gui, Add, text, x20, &Comments:
					Gui, Add, Edit, x20 w580 r3 vCiteComments

					

					;Buttons
					Gui, Add, Button, w50 x200 y620, Cancel
					Gui, Add, Button, w50 x270 y620, Clear
					Gui, Add, Button, w50 x340 y620, OK
				}
			
			;Display
					;Full reference
					Gui, Add, text, x30 y500 w400, Reference:
					Gui, Add, text, x45 y520 w400, %referencecb%

				Gui, -maximizebox -minimizebox
				Gui, Show, h650, Reference Manager ("The RM")
				return
			
			;Gui behavior

				GuiClose:
				GuiEscape:
				ButtonCancel: 
				SavedGUI = 1
				Gui, Cancel
				return

				ButtonClear:
				SavedGUI = 0
				Gui, cancel
				return

				ButtonOK:
				Gui, Submit
				SavedGUI = 1
				Gui, Cancel
				Gosub, Appendit
				return

			Appendit:
			FileAppend, `n"%raw%"`,"%accessDate%"`,"%WindowName%"`,"%dateyear%"`,"%journalName%"`,"%AuthorName%"`,"%titleText%"`,"%_%"`,"%abstract%"`,"%CiteComments%"`,"%c11%"`,"%c12%"`,"%c13%"`,"%c14%"`,"%c15%"`,"%c16%"`,"%c17%"`,"%c18%"`,"%c19%"`,"%c20%"`,"%c21%"`,"%c22%"`,"%c23%"`,"%c24%"`,"%c25%"`,"%c26%"`,"%c27%"`,"%c28%"`,"%c29%"`,"%c30%"`,"%c31%"`,"%c32%"`,"%c33%"`,"%c34%"`,"%c35%"`,"%c36%"`,"%c37%"`,"%c38%"`,"%c39%"`,"%c40%"`,"%c41%"`,"%c42%"`,"%c43%"`,"%c44%"`,"%c45%"`,"%c46%"`,"%c47%"`,"%c48%"`,"%c49%"`,"%c50%"`,"%c51%"`,"%c52%"`,"%c53%"`,"%c54%"`,"%c55%"`,"%c56%"`,"%c57%"`,"%c58%"`,"%c59%"`,"%c60%"`,"%c61%"`,"%c62%"`,"%c63%"`,"%c64%"`,"%c65%"`,"%c66%"`,"%c67%"`,"%c68%"`,"%c69%"`,"%c70%", referencehistory.csv ; columns separated by literal commas

			;empties variables ; leaves checked gui intact
				c11 :=
				c12 :=
				c13 :=
				c14 :=
				c15 :=
				c16 :=
				c17 :=
				c18 :=
				c19 :=
				c20 :=
				c21 :=
				c22 :=
				c23 :=
				c24 :=
				c25 :=
				c26 :=
				c27 :=
				c28 :=
				c29 :=
				c30 :=
				c31 :=
				c32 :=
				c33 :=
				c34 :=
				c35 :=
				c36 :=
				c37 :=
				c38 :=
				c39 :=
				c40 :=
				c41 :=
				c42 :=
				c43 :=
				c44 :=
				c45 :=
				c46 :=
				c47 :=
				c48 :=
				c49 :=
				c50 :=
				c51 :=
				c52 :=
				c53 :=
				c54 :=
				c55 :=
				c56 :=
				c57 :=
				c58 :=
				c59 :=
				c60 :=
				c61 :=
				c62 :=
				c63 :=
				c64 :=
				c65 :=
				c66 :=
				c67 :=
				c68 :=
				c69 :=
				c70 :=
				CiteComments :=

			;preview
				stringleft, preview, ParsedText, 20
				tooltip, Saved "%preview%..."
				SetTimer, tooltiptoggle, 2500
			return
		}

;------------------------------------------------------------------------------
; reference parser function

referenceparser(usedvar)
	{
		global
		ParsedText := usedvar
		StringReplace, ParsedText, ParsedText, `r`n, %A_Space%%A_Space%, All; replace line breaks and carriage returns with a double space
		StringReplace, ParsedText, ParsedText, ", "", All; replace double quotes with single quotes; necessary to keep commas from delimiting
		raw := ParsedText
		StringReplace, ParsedText, ParsedText, ?, ., All; replace question mark with period; helps separate better.
		WinGetActiveTitle, WindowName
		accessDate = %A_MM%-%A_DD%-%A_YYYY%
		; Last Name of first Author "AuthorName"
			StringGetPos, AuthorLength, ParsedText, `,
			StringLeft, AuthorName, ParsedText, %AuthorLength%

		; Year published ---- "dateYear"
			StringGetPos, DateOpen, ParsedText, (
			StringGetPos, DateClose, ParsedText, )
			DateLength := DateClose - DateOpen + 1 
			StringMid, dateYear, ParsedText, DateOpen + 2, DateLength - 2 ; stores year of publication in variable DateYear - excludes parenthesis when assigning year to variable dateYear

		; Title of the article ---- "titleText"
			StringGetPos, titleStartPos, ParsedText, ). , L
			titleStartPos += 3
			StringGetPos, titleEndPos, ParsedText, ., L, titleStartPos
			titleLength := titleEndPos - titleStartPos
			StringMid, titleText, ParsedText, DateClose + 4, titleLength
			stringreplace, fileTitle, titleText, :, -, A ; creates a variable "fileTitle" and replaces colon with a dash
			
		; Journal ----"journalName"
			stringgetpos, startTitle, ParsedText, ). , L
			stringgetpos, endTitle, ParsedText, ., L, startTitle + 3
			stringgetpos, endjournal, ParsedText, `,, L, endTitle + 1
			stringmid, journalName, ParsedText, endTitle + 3, endJournal - endTitle - 2
	}


------------------------------------------------------------------------------
;masterstore function

	MasterStore()
		{
			global
			MasterStore := getText()
			WinGetActiveTitle, WindowName
			StringReplace, MasterStore, MasterStore, `r`n, %A_Space%%A_Space%, All; replace line breaks and enter with a double space
			StringReplace, MasterStore, MasterStore, ", "", All; replace double quotes with single quotes; necessary to keep commas from delimiting
			stringMS = "%MasterStore%"
			stringWindowName = "%WindowName%"
			stringDate = "%A_MM%-%A_DD%-%A_YYYY%"
			FileAppend, %stringDate%`,%stringWindowName%`,%stringMS%`n, masterSave.CSV
			stringleft, preview, MasterStore, 20
			tooltip, Saved "%preview%..."
			SetTimer, tooltiptoggle, 2500
			return
		}