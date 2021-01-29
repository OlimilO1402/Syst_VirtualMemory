Attribute VB_Name = "MVirtualMem"
'Option Explicit
'
'Private Const MEM_COMMIT         As Long = &H1000
'Private Const MEM_RESERVE        As Long = &H2000
'Private Const MEM_DECOMMIT       As Long = &H4000
'Private Const MEM_RELEASE        As Long = &H8000
'
'Private Const MEM_FREE           As Long = &H10000
'Private Const MEM_PRIVATE        As Long = &H20000
'Private Const MEM_MAPPED         As Long = &H40000
'Private Const MEM_RESET          As Long = &H80000
'
'Private Const MEM_TOP_DOWN       As Long = &H100000
'Private Const MEM_WRITE_WATCH    As Long = &H200000
'Private Const MEM_PHYSICAL       As Long = &H400000
'
'Private Const MEM_IMAGE          As Long = &H1000000
'
'Private Const MEM_4MB_PAGES      As Long = &H80000000
'
'Private Const MEM_E_INVALID_LINK As Long = &H80080010
'Private Const MEM_E_INVALID_ROOT As Long = &H80080009
'Private Const MEM_E_INVALID_SIZE As Long = &H80080011
'
'Private Const PAGE_NOACCESS          As Long = &H1
'Private Const PAGE_READONLY          As Long = &H2
'Private Const PAGE_READWRITE         As Long = &H4
'Private Const PAGE_WRITECOPY         As Long = &H8
'Private Const PAGE_EXECUTE           As Long = &H10
'Private Const PAGE_EXECUTE_READ      As Long = &H20
'Private Const PAGE_EXECUTE_READWRITE As Long = &H40
'Private Const PAGE_EXECUTE_WRITECOPY As Long = &H80
'Private Const PAGE_GUARD             As Long = &H100
'Private Const PAGE_NOCACHE           As Long = &H200
'Private Const PAGE_WRITECOMBINE      As Long = &H400
'
'Public Type SYSTEM_INFO
'    wProcessorArchitecture  As Integer
'    wReserved               As Integer
'    dwPageSize              As Long
'    lpMinimumApplicationAddress As Long
'    lpMaximumApplicationAddress As Long
'    dwActiveProcessorMask   As Long
'    dwNumberOfProcessors    As Long
'    dwProcessorType         As Long
'    dwAllocationGranularity As Long
'    wProcessorLevel         As Integer
'    wProcessorRevision      As Integer
'End Type
'
'Public Type MEMORY_BASIC_INFORMATION
'    BaseAddress As Long
'    AllocationBase As Long
'    AllocationProtect As Long
'    RegionSize As Long
'    State As Long
'    Protect As Long
'    lType As Long
'End Type
''
'Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLength As Long)
'
'Public Declare Sub GetSystemInfo Lib "kernel32" ( _
'    ByRef lpSystemInfo As SYSTEM_INFO _
')
'Public Declare Sub GetNativeSystemInfo Lib "kernel32" ( _
'    ByRef lpSystemInfo As SYSTEM_INFO _
')
''lpAddress As Any ist eigentlich Quatsch
''es wird dort nichts zurückgegeben
''Public Declare Function VirtualAlloc Lib "kernel32.dll" ( _
''    ByRef lpAddress As Any, _
''    ByVal dwSize As Long, _
''    ByVal flAllocationType As Long, _
''    ByVal flProtect As Long _
'') As Long
'Public Declare Function VirtualAlloc Lib "kernel32" ( _
'    ByVal lpAddress As Long, _
'    ByRef dwSize As Long, _
'    ByVal flAllocationType As Long, _
'    ByVal flProtect As Long _
') As Long
'Public Declare Function VirtualFree Lib "kernel32" ( _
'    ByRef lpAddress As Any, _
'    ByVal dwSize As Long, _
'    ByVal dwFreeType As Long _
') As Long
'Public Declare Function VirtualLock Lib "kernel32" ( _
'    ByRef lpAddress As Any, _
'    ByVal dwSize As Long _
') As Long
'Public Declare Function VirtualProtect Lib "kernel32" ( _
'    ByRef lpAddress As Any, _
'    ByVal dwSize As Long, _
'    ByVal flNewProtect As Long, _
'    ByRef lpflOldProtect As Long _
') As Long
'Public Declare Function VirtualQuery Lib "kernel32" ( _
'    ByRef lpAddress As Any, _
'    ByRef lpBuffer As MEMORY_BASIC_INFORMATION, _
'    ByVal dwLength As Long _
') As Long
'Public Declare Function VirtualUnlock Lib "kernel32" ( _
'    ByRef lpAddress As Any, _
'    ByVal dwSize As Long _
') As Long
'
'Public Declare Function VirtualAllocEx Lib "kernel32" ( _
'    ByVal hProcess As Long, _
'    ByRef lpAddress As Any, _
'    ByRef dwSize As Long, _
'    ByVal flAllocationType As Long, _
'    ByVal flProtect As Long _
') As Long
'Public Declare Function VirtualQueryEx Lib "kernel32" ( _
'    ByVal hProcess As Long, _
'    ByRef lpAddress As Any, _
'    ByRef lpBuffer As MEMORY_BASIC_INFORMATION, _
'    ByVal dwLength As Long _
') As Long
