Option Explicit

Private Declare Function CreateThread Lib "kernel32" ( _
  ByVal ThreadAttributes As LongPtr, ByVal StackSize As LongPtr, _
  ByVal StartAddress As LongPtr, ByVal Parameter As LongPtr, _
  ByVal CreationFlags As Long, Optional ByVal ThreadId As LongPtr) As LongPtr

Private Declare Function GetExitCodeThread Lib "kernel32" ( _
  ByVal ThreadHandle As LongPtr, _
  Optional ByVal ExitCode As LongPtr) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" ( _
  ByVal Handle As LongPtr, ByVal Timeout As Long) As Long
' for Timeout
Private Const INFINITE As Long = -1

' ThreadProc spec:
' Function ThreadProc(ByVal Parameter As LongPtr) As Long
'   or (VB)
' Function ThreadProc(ByRef Parameter As T) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal dst As LongPtr, _
  ByVal src As LongPtr, ByVal sz As LongPtr)

Private Declare Function VirtualAlloc Lib "kernel32" ( _ 
  ByVal BaseAddr As LongPtr, ByVal sz As LongPtr, ByVal AllocType As Long, _
  ByVal Protect As Long) As LongPtr
' for AllocType
Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
' for Protect
Private Const PAGE_EXECUTE As Long = &H10
Private Const PAGE_EXECUTE_READ As Long = &H20
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const PAGE_NOACCESS As Long = &H01
Private Const PAGE_READONLY As Long = &H02
Private Const PAGE_READWRITE As Long = &H04

Private Declare Function VirtualFree Lib "kernel32" ( _
  ByVal BaseAddr As LongPtr, Optional ByVal sz As LongPtr = 0, _
  Optional ByVal FreeType As Long = &H8000) As Long

Private Declare Function VirtualProtect Lib "kernel32" ( _
  ByVal BaseAddr As LongPtr, ByVal sz As LongPtr, ByVal Protect As Long, _
  ByVal OldProtect As LongPtr) As Long

Public Function CallFn(ByVal FnPtr As LongPtr, ByVal ArgPtr As LongPtr, _
  Optional ByRef Result As Long) As Boolean
    CallFn = False

    Dim hthread As LongPtr
    hthread = CreateThread(0, 0, FnPtr, ArgPtr, 0)
    If hthread = 0 Then
        Exit Function
    End If

    If WaitForSingleObject(hthread, INFINITE) <> 0 Then
        Exit Function
    End If

    If GetExitCodeThread(hthread, VarPtr(Result)) = 0 Then
        Exit Function
    End If
    
    CallFn = True
End Function

Public Function FnPtr(ByVal FnAddress As LongPtr) As LongPtr
    FnPtr = FnAddress
End Function

Private Function LaunchMe(ByRef x As Long) As Long
    LaunchMe = x + 5
End Function

Private Function OrLaunchMe(ByRef x As Long) As Long
    OrLaunchMe = x * 7
End Function

Private Function MakeFn(ByRef FnCode() As Byte) As LongPtr
    MakeFn = 0

    Dim code_len As LongPtr
    code_len = UBound(FnCode) - LBound(FnCode) + 1

    Dim page As LongPtr
    page = VirtualAlloc(0, code_len, MEM_COMMIT Or MEM_RESERVE, _
      PAGE_EXECUTE_READWRITE) ' see below
    If page = 0 Then
        Exit Function
    End If

    Call RtlMoveMemory(page, VarPtr(FnCode(LBound(FnCode))), code_len)

    ' VirtualProtect always fails; not sure why
    '   for now we can just map the page initially as PAGE_EXECUTE_READWRITE;
    '   it'd be better to allocate it as PAGE_READWRITE and then change it after
    '   the code is written
    If False Then
        If VirtualProtect(page, code_len, PAGE_EXECUTE_READ, 0) = 0 Then
            Call VirtualFree(page)
            Exit Function
        End If
    End If

    MakeFn = page
End Function

Private Sub ReleaseFn(ByVal fn As LongPtr)
    Call VirtualFree(fn)
End Sub

Private Function TimesFive() As Byte()
    Dim bytes (0 To 17) As Byte
    ' push %ebp
    bytes(0) = &H55
    ' movl 8(%esp),%eax
    bytes(1) = &H8b
    bytes(2) = &H44
    bytes(3) = &H24
    bytes(4) = &H08
    ' movl (%eax),%eax
    bytes(5) = &H8b
    bytes(6) = &H00
    ' movl $5,%ecx
    bytes(7) = &Hb9
    bytes(8) = &H05
    bytes(9) = &H00
    bytes(10) = &H00
    bytes(11) = &H00
    ' mull %ecx
    bytes(12) = &Hf7
    bytes(13) = &He1
    ' pop %ebp
    bytes(14) = &H5d
    ' ret $4
    bytes(15) = &Hc2
    bytes(16) = &H04
    bytes(17) = &H00
    TimesFive = bytes
End Function

Public Sub Main
    Dim x As Long, result As Long
    x = 7
    Dim whichfn As LongPtr
    If Rnd() > 0.5 Then
        whichfn = FnPtr(AddressOf LaunchMe)
    Else
        whichfn = FnPtr(AddressOf OrLaunchMe)
    End If
    If Not CallFn(whichfn, VarPtr(x), result) Then
        MsgBox "Error calling function."
        Exit Sub
    End If
    Debug.Print "whichfn(" & x & ") = " & result

    Dim asmfn As LongPtr
    asmfn = MakeFn(TimesFive())
    If asmfn = 0 Then
        MsgBox "Error generating function."
        Exit Sub
    End If
    If Not CallFn(asmfn, VarPtr(x), result) Then
        Call ReleaseFn(asmfn)
        MsgBox "Error calling function."
        Exit Sub
    End If
    Call ReleaseFn(asmfn)
    Debug.Print "asmfn(" & x & ") = " & result
End Sub