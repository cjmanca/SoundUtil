; ProcessInfo.ahk - Function library to retrieve various application process informations:
; - Script's own process identifier
; - Parent process ID of a process (the caller application)
; - Process name by process ID (filename without path)
; - Thread count by process ID (number of threads created by process)
; - Full filename by process ID (GetModuleFileNameEx() function)
;
; Tested with AutoHotkey 1.0.46.10
;
; Created by HuBa
; Contact: http://www.autohotkey.com/forum/profile.php?mode=viewprofile&u=4693
;
; Portions of the script are based upon the GetProcessList() function by wOxxOm
; (http://www.autohotkey.com/forum/viewtopic.php?p=65983#65983)

GetCurrentProcessID()
{
  Return DllCall("GetCurrentProcessId")  ; http://msdn2.microsoft.com/ms683180.aspx
}

GetCurrentParentProcessID()
{
  Return GetParentProcessID(GetCurrentProcessID())
}

GetProcessName(ProcessID)
{
  Return GetProcessInformation(ProcessID, "Str", 260 * (A_IsUnicode ? 2 : 1), 32 + A_PtrSize)  ; TCHAR szExeFile[MAX_PATH]
}

GetParentProcessID(ProcessID)
{
  Return GetProcessInformation(ProcessID, "UInt *", 8, 20 + A_PtrSize)  ; DWORD th32ParentProcessID
}

GetProcessThreadCount(ProcessID)
{
  Return GetProcessInformation(ProcessID, "UInt *", 8, 16 + A_PtrSize)  ; DWORD cntThreads
}

;{
; The function retrieves a value of the field from the
; PROCESSENTRY32 structure of the specified process.
;
; Parameters:
; - ProcessID - the PID of the process for which to retrieve the PROCESSENTRY32
; information
; - CallVariableType - type of value to get (~type of DllCall parameter)
; - VariableCapacity - size of the buffer [in bytes] to which to retrieve the value
; - DataOffset - how far from beginning of PROCESSENTRY32 structure to search
; for data
;
; Returns:
; - th32DataEntry - a value read from PROCESSENTRY32 structure
;
; Remarks:
; - values that are possible to read:
; http://msdn.microsoft.com/en-us/library/windows/desktop/ms684839%28v=vs.85%29.aspx
;
    ;~ typedef struct tagPROCESSENTRY32 {
      ;~ DWORD     dwSize;
      ;~ DWORD     cntUsage;
      ;~ DWORD     th32ProcessID;
      ;~ ULONG_PTR th32DefaultHeapID;
      ;~ DWORD     th32ModuleID;
      ;~ DWORD     cntThreads;
      ;~ DWORD     th32ParentProcessID;
      ;~ LONG      pcPriClassBase;
      ;~ DWORD     dwFlags;
      ;~ TCHAR     szExeFile[MAX_PATH];
    ;~ } PROCESSENTRY32, *PPROCESSENTRY32;
;
;}
GetProcessInformation(ProcessID, CallVariableType, VariableCapacity, DataOffset)
{
    static PE32_size := 8 * 4 + A_PtrSize + 260 * (A_IsUnicode ? 2 : 1)
	
	
  hSnapshot := DLLCall("CreateToolhelp32Snapshot", "UInt", 2, "UInt", 0)  ; TH32CS_SNAPPROCESS = 2
  if (hSnapshot >= 0)
  {
  
	VarSetCapacity(PE32, PE32_size, 0)  ; PROCESSENTRY32 structure -> http://msdn2.microsoft.com/ms684839.aspx
	DllCall("ntdll.dll\RtlFillMemoryUlong", "Ptr", &PE32, "UInt", 4, "UInt", PE32_size)  ; Set dwSize
	VarSetCapacity(th32ProcessID, 4, 0)
	DllCall("Process32First" . (A_IsUnicode ? "W" : ""), "Ptr", hSnapshot, "Ptr", &PE32)
	
    if (DllCall("Kernel32.dll\Process32First" . (A_IsUnicode ? "W" : ""), "Ptr", hSnapshot, "Ptr", &PE32))  ; http://msdn2.microsoft.com/ms684834.aspx
	{
      Loop
      {
        DllCall("RtlMoveMemory", "Ptr*", th32ProcessID, "Ptr", &PE32 + 8, "UInt", 4)  ; http://msdn2.microsoft.com/ms803004.aspx
        if (ProcessID = th32ProcessID)
        {
          VarSetCapacity(th32DataEntry, VariableCapacity, 0)
          
          DllCall("RtlMoveMemory", CallVariableType, th32DataEntry, "Ptr", &PE32 + DataOffset, "UInt", VariableCapacity)
          DllCall("CloseHandle", "Ptr", hSnapshot)  ; http://msdn2.microsoft.com/ms724211.aspx
          Return th32DataEntry  ; Process data found
        }
        if not DllCall("Process32Next" (A_IsUnicode ? "W" : ""), "Ptr", hSnapshot, "Ptr", &PE32)  ; http://msdn2.microsoft.com/ms684836.aspx
          Break
      }
	}
	error := DllCall("GetLastError")
	;MsgBox, error Found: %error%
    DllCall("CloseHandle", "Ptr", hSnapshot)
  }
  Return  ; Cannot find process
}



GetModuleFileNameEx(ProcessID)  ; modified version of shimanov's function
{
  if A_OSVersion in WIN_95, WIN_98, WIN_ME
    Return GetProcessName(ProcessID)
 
  ; #define PROCESS_VM_READ           (0x0010)
  ; #define PROCESS_QUERY_INFORMATION (0x0400)
  hProcess := DllCall( "OpenProcess", "UInt", 0x10|0x400, "Int", False, "UInt", ProcessID)
  if (ErrorLevel or hProcess = 0)
    Return
  FileNameSize := 260 * (A_IsUnicode ? 2 : 1)
  VarSetCapacity(ModuleFileName, FileNameSize, 0)
  CallResult := DllCall("Psapi.dll\GetModuleFileNameEx", "Ptr", hProcess, "Ptr", 0, "Str", ModuleFileName, "UInt", FileNameSize)
  DllCall("CloseHandle", "Ptr", hProcess)
  Return ModuleFileName
}

getActiveProcessID() {
	handle := DllCall("GetForegroundWindow", "Ptr")
	DllCall("GetWindowThreadProcessId", "Int", handle, "int*", pid)
	global true_pid := pid
	return pid
}

getActiveProcessName() {
	handle := DllCall("GetForegroundWindow", "Ptr")
	DllCall("GetWindowThreadProcessId", "Int", handle, "int*", pid)
	global true_pid := pid
	callback := RegisterCallback("enumChildCallback", "Fast")
	DllCall("EnumChildWindows", "Int", handle, "ptr", callback, "int", pid)
	handle := DllCall("OpenProcess", "Int", 0x0400, "Int", 0, "Int", true_pid)
	length := 259 ;max path length in windows
	VarSetCapacity(name, length)
	DllCall("QueryFullProcessImageName", "Int", handle, "Int", 0, "Ptr", &name, "int*", length)
	SplitPath, name, pname
	return pname
}
 
enumChildCallback(hwnd, pid) {
	DllCall("GetWindowThreadProcessId", "Int", hwnd, "int*", child_pid)
	if (child_pid != pid)
		global true_pid := child_pid
	return 1
}

EnumProcesses() {
	IfEqual, A_OSType, WIN32_WINDOWS, Return 0
	List_Sz := VarSetCapacity(Pid_List, 4000)  
	Res := DllCall("psapi.dll\EnumProcesses", "UInt",&Pid_List, "Int",List_Sz, "UInt *",PID_List_Actual) 
	IfLessOrEqual,Res,0, Return, Res
	_a := &PID_List 
	arr := []
	Loop, % (PID_List_Actual/4) 
	{ 
		arr.Push( (*(_a)+(*(_a+1)<<8)+(*(_a+2)<<16)+(*(_a+3)<<24)) )
		_a += 4 
	} 
	StringTrimLeft, arr, arr, 1 
	return arr
}

/*
GetModuleFileNameEx( p_pid )
{
	if A_OSVersion in WIN_95,WIN_98,WIN_ME 
	{
		MsgBox, This Windows version (%A_OSVersion%) is not supported.
		return 
	}
	
	if (p_pid == 0)
	{
		return "zero"
	}

    ; h 			:= DllCall("OpenProcess", "UInt", 0x10 | 0x400, "Int", false, "UInt", id, "Ptr")
	h_process 	:= DllCall("OpenProcess", "uint", 0x10 | 0x400, "int", false, "uint", p_pid, "Ptr")
	if ( ErrorLevel or h_process = 0 )
	{
		MsgBox, [OpenProcess] failed for: %p_pid%
		return
	}
	
	hModule := DllCall("LoadLibrary", "Str", "Psapi.dll")  ; Increase performance by preloading the library.
	s := 4096  ; size of buffers and arrays (4 KB)
    VarSetCapacity(name, s, 0)  ; a buffer that receives the base name of the module:
    e := DllCall("Psapi.dll\GetModuleBaseName", "Ptr", h_process, "Ptr", 0, "Str", name, "UInt", A_IsUnicode ? s/2 : s)
    if !e    ; fall-back method for 64-bit processes when in 32-bit mode:
        if e := DllCall("Psapi.dll\GetProcessImageFileName", "Ptr", h_process, "Str", name, "UInt", A_IsUnicode ? s/2 : s)
            SplitPath name, name
			
			
	DllCall( "CloseHandle", "uint", h_process )
	
	DllCall("FreeLibrary", "Ptr", hModule)  ; Unload the library to free memory.
	
	return, name
}

*/