#SingleInstance, Force
#MaxHotkeysPerInterval 2000
#HotkeyInterval 1000
#Include VA.ahk
#Include ProcessInfo.ahk

DetectHiddenWindows, on

AdjustAmount := 2
VolChannel := 2

Volume = null
HeldDown := false

InDownMethod := false
UpPending := false

SetMaster := false


soundParentExeExclusionsPath := A_ScriptDir . "\soundParentExeExclusions.txt"

if (!FileExist(soundParentExeExclusionsPath))
{
	FileAppend, explorer.exe, %soundParentExeExclusionsPath%
}

FileRead, soundParentExeExclusionsRaw, %A_ScriptDir%\soundParentExeExclusions.txt

soundParentExeExclusionsRaw := StrReplace(soundParentExeExclusionsRaw, "`r", "`n")
While InStr(soundParentExeExclusionsRaw, "`n`n")
{
	soundParentExeExclusionsRaw := StrReplace(soundParentExeExclusionsRaw, "`n`n", "`n")
}
soundParentExeExclusionsRaw := trim(soundParentExeExclusionsRaw, " `n`r`t")

parentExclusions := StrSplit(soundParentExeExclusionsRaw, ["`n","`r"], " `n`r`t")

parentExclusionsMatchList := Join(",", parentExclusions)

return


Volume_Up::

	InDownMethod := true
	if (!HeldDown && !SetMaster)
	{
		pid := getActiveProcessID()

		if !(Volume := GetVolumeObject(pid))
		{
			SetMaster := true
		}
		else
		{
			SetMaster := false
		}
		HeldDown := true
	}

	if (SetMaster)
	{
		vol := VA_GetMasterVolume()

		vol := vol + AdjustAmount

		if (vol <= 100)
		{
			VA_SetMasterVolume(vol)
		}
	}
	else
	{
		VA_ISimpleAudioVolume_GetMasterVolume(Volume, vol)
		
		vol := vol + (AdjustAmount/100)

		if (vol <= 1)
		{
			VA_ISimpleAudioVolume_SetMasterVolume(Volume, vol)
		}
		else
		{
			percAdjust := ((vol - 1) / (AdjustAmount/100))
			
			VA_ISimpleAudioVolume_SetMasterVolume(Volume, 1)

			vol := VA_GetMasterVolume()

			vol := vol + (AdjustAmount * percAdjust)
			
			if (vol <= 100)
			{
				VA_SetMasterVolume(vol)
			}
		}
	}
	if (UpPending)
	{
		UpPending := false
		if (!SetMaster)
		{
			ObjRelease(Volume)
		}
		SetMaster := false
		HeldDown := false
	}
	InDownMethod := false
return
 
 
Volume_Up Up::
	if (InDownMethod)
	{
		UpPending := true
	}
	else
	{
		UpPending := false
		if (!SetMaster)
		{
			ObjRelease(Volume)
		}
		SetMaster := false
		HeldDown := false
	}
return


 
Volume_Down::

	InDownMethod := true
	if (!HeldDown && !SetMaster)
	{
		pid := getActiveProcessID()

		if !(Volume := GetVolumeObject(pid))
		{
			; Display("Couldn't GetVolumeObject for the Active Window: " . test)
			SetMaster := true
		}
		else
		{
			SetMaster := false
		}
		HeldDown := true
	}
	
	if (SetMaster)
	{
		vol := VA_GetMasterVolume()
		
		vol := vol - AdjustAmount
		
		if (vol >= 0)
		{
			VA_SetMasterVolume(vol)
		}
	}
	else
	{
		VA_ISimpleAudioVolume_GetMasterVolume(Volume, vol)

		if (vol > 0)
		{
			vol := vol - (AdjustAmount/100)

			VA_ISimpleAudioVolume_SetMasterVolume(Volume, vol)
		}
	}
	if (UpPending)
	{
		UpPending := false
		if (!SetMaster)
		{
			ObjRelease(Volume)
		}
		SetMaster := false
		HeldDown := false
	}
	InDownMethod := false
return

Volume_Down Up::
	if (InDownMethod)
	{
		UpPending := true
	}
	else
	{
		UpPending := false
		if (!SetMaster)
		{
			ObjRelease(Volume)
		}
		SetMaster := false
		HeldDown := false
	}
return 


GetVolumeObject(Param)
{
	global parentExclusionsMatchList
	
    static IID_IASM2 := "{77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}"
    , IID_IASC2 := "{bfb7ff88-7239-4fc9-8fa2-07c950be9c6d}"
    , IID_ISAV := "{87CE5498-68D6-44E5-9215-6DA47EF883D8}"

	debugStr := ""

    ; Turn empty into integer
    if !Param
        Param := 0
    
    ; Get PID from process name
    if Param is not Integer
    {
        Process, Exist, %Param%
        Param := ErrorLevel
    }

	ParamProcessPath := GetModuleFileNameEx(Param)
	parentParamProcessID := GetParentProcessID(Param)
	parentParamProcessPath := GetModuleFileNameEx(parentParamProcessID)
	
    ; GetDefaultAudioEndpoint
    DAE := VA_GetDevice()
    
    ; activate the session manager
    VA_IMMDevice_Activate(DAE, IID_IASM2, 0, 0, IASM2)
    
    ; enumerate sessions for on this device
    VA_IAudioSessionManager2_GetSessionEnumerator(IASM2, IASE)
    VA_IAudioSessionEnumerator_GetCount(IASE, Count)

	; Display("VA_IAudioSessionEnumerator_GetCount: " . Count)

	SplitPath, ParamProcessPath, ParamProcessName
	SplitPath, parentParamProcessPath, parentParamProcessName

	if ParamProcessName in %parentExclusionsMatchList%
	{
		return 0
	}
	if parentParamProcessName in %parentExclusionsMatchList%
	{
		parentParamProcessID := 0
		parentParamProcessName := ""
		parentParamProcessPath := ""
	}
	
	debugStr := debugStr . Param . ": " . ParamProcessName . " - " . ParamProcessPath . ", Parent (" . parentParamProcessID . "): " . parentParamProcessPath . "`n"
	
	debugStr := debugStr . "GetCount: " . Count . "`n"
	
	bestIASAV := 0
	bestIASAVByParent := 0
	bestIASAVBySibling := 0
	ISAV := 0
	
    ; search for an audio session with the required name
    Loop, % Count
    {
        ; Get the IAudioSessionControl object
        VA_IAudioSessionEnumerator_GetSession(IASE, A_Index-1, IASC)
        
        ; Query the IAudioSessionControl for an IAudioSessionControl2 object
        IASC2 := ComObjQuery(IASC, IID_IASC2)
        
		if (VA_IAudioSessionControl2_IsSystemSoundsSession(IASC2))
		{
			; Get the sessions process ID
			VA_IAudioSessionControl2_GetProcessID(IASC2, SPID)
		
			;VA_IAudioSessionControl_GetDisplayName(IASC, lDispalyName)
			VA_IAudioSessionControl2_GetSessionIdentifier(IASC2, lSessionIdentifier)
			
			
			processPath := GetModuleFileNameEx(SPID)
			parentSPID := GetParentProcessID(SPID)
			parentProcessPath := GetModuleFileNameEx(parentSPID)
			
			SplitPath, processPath, processName
			SplitPath, parentProcessPath, parentProcessName
			
			if processName in %parentExclusionsMatchList%
			{
				ObjRelease(IASC)
				ObjRelease(IASC2)
				continue
			}
			
			if parentProcessName in %parentExclusionsMatchList%
			{
				parentSPID := 0
				parentProcessName := ""
				parentProcessPath := ""
			}
			
			if (Param == parentSPID && parentProcessName != "")
			{
				bestIASAVByParent := ComObjQuery(IASC2, IID_ISAV)
			}
			if (parentParamProcessID == SPID && parentParamProcessName != "")
			{
				bestIASAVByParent := ComObjQuery(IASC2, IID_ISAV)
			}
			
			if (parentParamProcessID == parentSPID && parentParamProcessID > 0)
			{
				bestIASAVBySibling := ComObjQuery(IASC2, IID_ISAV)
			}
			
			if (ParamProcessName == processName)
			{
				bestIASAV := ComObjQuery(IASC2, IID_ISAV)
			}
					
			debugStr := debugStr . SPID . ": " . processName . " - " . processPath . " (" . parentSPID . ", " . parentProcessPath . ")`n"
			; If the process name is the one we are looking for
			if (SPID == Param)
			{
				; Query for the ISimpleAudioVolume
				ISAV := ComObjQuery(IASC2, IID_ISAV)
				
				ObjRelease(IASC)
				ObjRelease(IASC2)
				break
			}
		}
        
		ObjRelease(IASC)
        ObjRelease(IASC2)
    }

	if (ISAV == 0)
	{
		if (bestIASAV)
		{
			debugStr := debugStr . "bestIASAV: " . bestIASAV . "`n"
			ISAV := bestIASAV
		}
		else if (bestIASAVByParent)
		{
			debugStr := debugStr . "bestIASAVByParent: " . bestIASAVByParent . "`n"
			ISAV := bestIASAVByParent
		}
		else
		{
			debugStr := debugStr . "bestIASAVBySibling: " . bestIASAVBySibling . "`n"
			ISAV := bestIASAVBySibling
		}
	}
	
	; Display(debugStr)
	
    ObjRelease(IASE)
    ObjRelease(IASM2)
    ObjRelease(DAE)
    return ISAV
}



 Join(s,p*)
 {
	static _:="".base.Join:=Func("Join")
	for k,v in p
	{
		if isobject(v)
		{
			for k2, v2 in v
			{
				o.=s v2
			}
		}
		else
		{
			o.=s v
		}
	}
	return SubStr(o,StrLen(s)+1)
}


;
; ISimpleAudioVolume : {87CE5498-68D6-44E5-9215-6DA47EF883D8}
;
VA_ISimpleAudioVolume_SetMasterVolume(this, ByRef fLevel, GuidEventContext="") {
    return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "float", fLevel, "ptr", VA_GUID(GuidEventContext))
}
VA_ISimpleAudioVolume_GetMasterVolume(this, ByRef fLevel) {
    return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "float*", fLevel)
}
VA_ISimpleAudioVolume_SetMute(this, ByRef Muted, GuidEventContext="") {
    return DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "int", Muted, "ptr", VA_GUID(GuidEventContext))
}
VA_ISimpleAudioVolume_GetMute(this, ByRef Muted) {
    return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "int*", Muted)
}



RemoveToolTip:
SetTimer, RemoveToolTip, Off
ToolTip
return

Display(displayStr)
{
	ToolTip, %A_ScriptName%`n`n%displayStr%
	SetTimer, RemoveToolTip, 5000
}
