Attribute VB_Name = "MFncPtr"
Option Explicit

#If VBA7 = 0 Then
    Public Enum LongPtr: [_]: End Enum
#End If

#If VBA7 Then
    Private Declare PtrSafe Sub RtlMoveMemory Lib "Kernel32" (ByVal dst As LongPtr, ByVal src As LongPtr, ByVal sz As Long)
    Private Declare PtrSafe Function CreateThread Lib "Kernel32" (ByVal ThreadAttributes As LongPtr, ByVal StackSize As LongPtr, ByVal StartAddress As LongPtr, ByVal Parameter As LongPtr, ByVal CreationFlags As Long, Optional ByVal ThreadId As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetExitCodeThread Lib "Kernel32" (ByVal ThreadHandle As LongPtr, Optional ByVal ExitCode As LongPtr) As Long
    Private Declare PtrSafe Function WaitForSingleObject Lib "Kernel32" (ByVal Handle As LongPtr, ByVal Timeout As Long) As Long
    Private Declare PtrSafe Function CloseHandle Lib "Kernel32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function VirtualAlloc Lib "Kernel32" (ByVal BaseAddr As LongPtr, ByVal sz As LongPtr, ByVal AllocType As Long, ByVal Protect As Long) As LongPtr
    Private Declare PtrSafe Function VirtualFree Lib "Kernel32" (ByVal BaseAddr As LongPtr, Optional ByVal sz As LongPtr = 0, Optional ByVal FreeType As Long = &H8000) As Long
    Private Declare PtrSafe Function VirtualProtect Lib "Kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long ' Bool
#Else
    Private Declare Sub RtlMoveMemory Lib "Kernel32" (ByVal dst As LongPtr, ByVal src As LongPtr, ByVal sz As Long)
    Private Declare Function CreateThread Lib "Kernel32" (ByVal ThreadAttributes As LongPtr, ByVal StackSize As LongPtr, ByVal StartAddress As LongPtr, ByVal Parameter As LongPtr, ByVal CreationFlags As Long, Optional ByVal ThreadId As LongPtr) As LongPtr
    Private Declare Function GetExitCodeThread Lib "Kernel32" (ByVal ThreadHandle As LongPtr, Optional ByVal ExitCode As LongPtr) As Long
    Private Declare Function WaitForSingleObject Lib "Kernel32" (ByVal Handle As LongPtr, ByVal Timeout As Long) As Long
    Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As LongPtr) As Long
    Private Declare Function VirtualAlloc Lib "Kernel32" (ByVal BaseAddr As LongPtr, ByVal sz As LongPtr, ByVal AllocType As Long, ByVal Protect As Long) As LongPtr
    Private Declare Function VirtualFree Lib "Kernel32" (ByVal BaseAddr As LongPtr, Optional ByVal sz As LongPtr = 0, Optional ByVal FreeType As Long = &H8000) As Long
    Private Declare Function VirtualProtect Lib "Kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long ' Bool
#End If
' for Timeout
Private Const INFINITE As Long = -1

' ThreadProc spec:
' Function ThreadProc(ByVal Parameter As LongPtr) As Long
'   or (VB)
' Function ThreadProc(ByRef Parameter As T) As Long

' for AllocType
Private Const MEM_COMMIT             As Long = &H1000
Private Const MEM_RESERVE            As Long = &H2000

' for Protect
Private Const PAGE_NOACCESS          As Long = &H1
Private Const PAGE_READONLY          As Long = &H2
Private Const PAGE_READWRITE         As Long = &H4
Private Const PAGE_EXECUTE           As Long = &H10
Private Const PAGE_EXECUTE_READ      As Long = &H20
Private Const PAGE_EXECUTE_READWRITE As Long = &H40

Public Function Thread_CallFunctionPtr(ByVal pFunction As LongPtr, ByVal pArguments As LongPtr, Optional ByRef Result_out As Long) As Boolean
    Dim hThread As LongPtr: hThread = CreateThread(0, 0, pFunction, pArguments, 0)
    If hThread = 0 Then Exit Function
    If WaitForSingleObject(hThread, INFINITE) <> 0 Then Exit Function
    If GetExitCodeThread(hThread, VarPtr(Result_out)) = 0 Then Exit Function
    Thread_CallFunctionPtr = True
    CloseHandle hThread
End Function

Public Function FncPtr(ByVal FnAddress As LongPtr) As LongPtr
    FncPtr = FnAddress
End Function

Public Function MakeFunction(ByRef FnCode() As Byte) As LongPtr
    Dim FuncSize As Long: FuncSize = UBound(FnCode) - LBound(FnCode) + 1
    MakeFunction = VirtualAlloc(0, FuncSize, MEM_COMMIT Or MEM_RESERVE, PAGE_READWRITE)   'PAGE_EXECUTE_READWRITE)   ' see below
    If MakeFunction = 0 Then Exit Function
    RtlMoveMemory MakeFunction, VarPtr(FnCode(0)), FuncSize
    ' VirtualProtect always fails; not sure why
    '   for now we can just map the page initially as PAGE_EXECUTE_READWRITE;
    '   it'd be better to allocate it as PAGE_READWRITE and then change it after
    '   the code is written
    'If False Then
    Dim oldProt As Long
    If VirtualProtect(MakeFunction, FuncSize, PAGE_EXECUTE_READWRITE, oldProt) = 0 Then
        VirtualFree MakeFunction
    Else
        Debug.Print "VirtualProtect OK"
    End If
    'End If
End Function

Public Sub ReleaseFn(ByVal fn As LongPtr)
    Call VirtualFree(fn)
End Sub
