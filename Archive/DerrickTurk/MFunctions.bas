Attribute VB_Name = "MFunctions"
Option Explicit

Public Function AddFive(ByRef x As Long) As Long
    AddFive = x + 5
End Function

Public Function TimesSeven(ByRef x As Long) As Long
    TimesSeven = x * 7
End Function

Public Function FuncTimesFive_GetAsm() As Byte()
    Dim ba(0 To 17) As Byte
    
    ' push %ebp:
    ba(0) = &H55
    
    ' movl 8(%esp),%eax:
    ba(1) = &H8B:  ba(2) = &H44: ba(3) = &H24: ba(4) = &H8
    
    ' movl (%eax),%eax
    ba(5) = &H8B:  ba(6) = &H0
        
    ' movl $5,%ecx:
    ba(7) = &HB9:  ba(8) = &H5:  ba(9) = &H0: ba(10) = &H0: ba(11) = &H0
    
    ' mull %ecx:
    ba(12) = &HF7: ba(13) = &HE1
    
    ' pop %ebp:
    ba(14) = &H5D
    
    ' ret $4:
    ba(15) = &HC2: ba(16) = &H4: ba(17) = &H0
    FuncTimesFive_GetAsm = ba
End Function

