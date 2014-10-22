;**************************************************************************************************
; GL Monthly File AutoHotkey Script
; Mark Richardson
; Nov 22, 2013
;**************************************************************************************************


;**************************************************************************************************
; Start
; Windows Key + M Shortcut Key
;**************************************************************************************************

#MaxThreadsPerHotkey 3
#m::
#MaxThreadsPerHotkey 1
if KeepWinZRunning  ; This means an underlying thread is already running the loop below.
{
    KeepWinZRunning := false  ; Signal that thread's loop to stop.
    return  ; End this thread so that the one underneath will resume and see the change made by the line above.
}


;Variables for Batch Selection String
Iterator = 0
;StartDate = 01-Nov-2013
;EndDate = 02-Nov-2013

InputBox, StartDate, , Start Date:,,,,,,,,01-%A_MMM%-%A_YYYY%
InputBox, EndDate  , , End Date:  ,,,,,,,,02-%A_MMM%-%A_YYYY%

;/* ; <--- comment Out ----------------\\\\\\\\\\\\\\\\\\//////////////////////---------------------------

;**************************************************************************************************
; Load the Modify Job Window
;**************************************************************************************************

IfWinExist System Mgmt: DB OpsView Scheduler - \\Remote
{
	WinActivate
}
else
{
	MsgBox 1, , OpsViewScheduler Not found
        ifMsgBox, Ok
		Exit ;Exit the Script
}

;Ensure that the Procure - GL Outbound Interface is the top item in the Schedule List
Click 800,140 ;Click on the Procure - Gl Outbound Interface
Click right   ;Right Click on the Procure - Gl Outbound Interface
Click 810,143 ;Move and Click on the "Modify" option
Sleep, 1000   ; 1 second

;**************************************************************************************************
; Setup and modify the 7 instances that generate output files
;**************************************************************************************************

;Move the Scheduled Task Window to the 0,0 cordinate of the first monitor
IfWinExist Scheduled Task - \\Remote
{
	MouseMove, 486, 40
	Click			;move to "Steps" tab and click on it
	Loop 8
	{
		MouseMove, 500, 182
		Click			;move/click to "Batch Selection" Text Area
		Send {Shift Down}{Home}
		Send {Shift Up}	
		Send CMPY=%Iterator%; SDATE=%StartDate%; EDATE=%EndDate%;CMPYFLG=1;
		Iterator := Iterator + 1
		MouseMove, 570, 353
		Sleep, 500   ; 1 second
		Click			;move/click down arrow
	}
	MouseMove, 380,405
	Click			;Move/Click on "Ok"
}
else
{
	MsgBox 1, , Scheduled Task Window Not found
        ifMsgBox, Ok
		Exit ;Exit the Script
}


;**************************************************************************************************
; Load the test Job Window
;**************************************************************************************************

IfWinExist System Mgmt: DB OpsView Scheduler - \\Remote
{
	WinActivate
}
else
{
	MsgBox 1, , OpsViewScheduler Not found
        ifMsgBox, Ok
		Exit ;Exit the Script
}
;Ensure that the Procure - GL Outbound Interface is the top item in the Schedule List
Click 800,140 		;Click on the Procure - Gl Outbound Interface
Click right   		;Right Click on the Procure - Gl Outbound Interface
Click 818, 229 		;Move and Click on the "Modify" option
Sleep, 1000   		;1 second


;**************************************************************************************************
; Run the Ops Job
;**************************************************************************************************

IfWinExist Test Job - \\Remote
{
	WinActivate
}
else
{
	MsgBox 1, , Test Job - \\Remote Window Not found
        ifMsgBox, Ok
		Exit ;Exit the Script
}

MouseMove 350,389		
Click 			;Move/Click on the "Execute" button


;**************************************************************************************************
; Loop and Respond to the Ops Job Status windows
;**************************************************************************************************
KeepWinZRunning := true
counter = 0
MessagesAcknowledged = 0
Loop
{
    	; The next four lines are the action you want to repeat (update them to suit your preferences):
	Sleep 2000
	counter := counter + 1
	ToolTip, Ops Job Responder Waiting for Job to finish. Count: %counter%, 0, 0
	
	IfWinExist  Request Execution Error - \\Remote
	{
		WinActivate
		MouseMove 194,120
		Click 			;Move/Click on the "Confirm" button
		MessagesAcknowledged := MessagesAcknowledged + 1
		if MessagesAcknowledged = 8
		{
			break 	; Break out of this loop.
		} 
	}

    	if not KeepWinZRunning  ; The user signaled the loop to stop by pressing Win-Z again.
        	break  		; Break out of this loop.
}
KeepWinZRunning := false  	; Reset in preparation for the next press of this hotkey.
ToolTip 			; Remove the tooltip

;*/  ;<--- Comment Out ----------------\\\\\\\\\\\\\\\\\\//////////////////////---------------------------

return
