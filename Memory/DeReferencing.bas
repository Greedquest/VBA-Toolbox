Attribute VB_Name = "DeReferencing"
'@Folder("VBAProject")
Option Explicit

Private Enum BOOL
    API_FALSE = 0
    'Use NOT (result = API_FALSE) for API_TRUE, as TRUE is just non-zero
End Enum

Private Enum VirtualProtectFlags 'See Memory Protection constants: https://docs.microsoft.com/en-gb/windows/win32/memory/memory-protection-constants
    PAGE_EXECUTE_READWRITE = &H40
    PAGE_READONLY = &H2
    RESET_TO_PREVIOUS = -1
End Enum

#If Win64 Then 'To decide whether to use 8 or 4 bytes per chunk of memory
    Private Declare Function GetMem Lib "msvbvm60" Alias "GetMem8" (ByRef src As Any, ByRef dest As Any) As Long
#Else
    Private Declare Function GetMem Lib "msvbvm60" Alias "GetMem4" (ByRef src As Any, ByRef dest As Any) As Long
#End If

Private Declare Function VirtualProtect Lib "kernel32" (ByRef location As Any, ByVal numberOfBytes As Long, ByVal newProtectionFlags As VirtualProtectFlags, ByVal lpOldProtectionFlags As LongPtr) As BOOL

'@Description("Pointer dereferencing; reads/ writes a single 4 byte (32-bit) or 8 byte (64-bit) block of memory at the address specified. Performs any necessary unprotecting")
Public Property Let DeReference(ByVal address As LongPtr, ByVal Value As LongPtr)
Attribute DeReference.VB_Description = "Pointer dereferencing; reads/ writes a single 4 byte (32-bit) or 8 byte (64-bit) block of memory at the address specified. Performs any necessary unprotecting"
    If ToggleMemoryProtection(address, LenB(Value), PAGE_EXECUTE_READWRITE) Then
        GetMem Value, ByVal address
        ToggleMemoryProtection address, LenB(Value)
    Else
        Err.Raise 5, Description:="That address is protected memory which cannot be accessed"
    End If
End Property

Public Property Get DeReference(ByVal address As LongPtr) As LongPtr
    If ToggleMemoryProtection(address, LenB(DeReference), PAGE_EXECUTE_READWRITE) Then
        GetMem ByVal address, DeReference
        ToggleMemoryProtection address, LenB(DeReference)
    Else
        Err.Raise 5, Description:="That address is protected memory which cannot be accessed"
    End If
End Property

Private Static Function ToggleMemoryProtection(ByVal address As LongPtr, ByVal numberOfBytes As Long, Optional ByVal newMemoryFlag As VirtualProtectFlags = RESET_TO_PREVIOUS) As Boolean
    Dim previousMemoryState As VirtualProtectFlags
    Dim unprotectWasNOOP As Boolean
    If newMemoryFlag = RESET_TO_PREVIOUS Then
        If unprotectWasNOOP Then
            ToggleMemoryProtection = True
        Else
            ToggleMemoryProtection = VirtualProtect(ByVal address, numberOfBytes, previousMemoryState, VarPtr(newMemoryFlag)) <> API_FALSE
        End If
    Else
        ToggleMemoryProtection = VirtualProtect(ByVal address, numberOfBytes, newMemoryFlag, VarPtr(previousMemoryState)) <> API_FALSE
        'check whether unprotecting even had an effect - if not then no need to toggle back
        unprotectWasNOOP = newMemoryFlag = previousMemoryState
    End If
End Function

