VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'Public Enum AllocationFlags
'    MemCommit = &H1000&             'MEM_COMMIT
'    MemReserve = &H2000&            'MEM_RESERVE
'    MemCommitReserve = &H3000&
'    MemDecommit = &H4000&           'MEM_DECOMMIT
'    MemRelease = &H8000&            'MEM_RELEASE
'    MemFree = &H10000               'MEM_FREE
'    MemPrivate = &H20000            'MEM_PRIVATE
'    MemMapped = &H40000             'MEM_MAPPED
'    MemReset = &H80000              'MEM_RESET
'    MemTopDown = &H100000           'MEM_TOP_DOWN
'    MemWriteWatch = &H200000        'MEM_WRITE_WATCH
'    MemPhysical = &H400000          'MEM_PHYSICAL
'    MemImage = &H1000000            'MEM_IMAGE
'    Mem4MbPages = &H80000000        'MEM_4MB_PAGES
'
''Private Const MEM_E_INVALID_LINK 'As Long = &H80080010
''Private Const MEM_E_INVALID_ROOT 'As Long = &H80080009
''Private Const MEM_E_INVALID_SIZE 'As Long = &H80080011
'
'End Enum
'Public Enum ProtectionFlags
'    PageNoAccess = &H1&             'PAGE_NOACCESS          As Long = &H1
'    PageReadOnly = &H2&             'PAGE_READONLY          As Long = &H2
'    PageReadWrite = &H4&            'PAGE_READWRITE         As Long = &H4
'    PageWriteCopy = &H8&            'PAGE_WRITECOPY         As Long = &H8
'    PageExecute = &H10&             'PAGE_EXECUTE           As Long = &H10
'    PageExecuteRead = &H20&         'PAGE_EXECUTE_READ      As Long = &H20
'    PageExecuteReadWrite = &H40&    'PAGE_EXECUTE_READWRITE As Long = &H40
'    PageExecuteWriteCopy = &H80&    'PAGE_EXECUTE_WRITECOPY As Long = &H80
'    PageGuard = &H100&              'PAGE_GUARD             As Long = &H100
'    PageNoCache = &H200&            'PAGE_NOCACHE           As Long = &H200
'    PageWiteCombine = &H400&        'PAGE_WRITECOMBINE      As Long = &H400
'End Enum
'Private Type TMemHandle
'    EAllocation As AllocationFlags
'    EProtection As ProtectionFlags
'    Size        As Long
'    Handle      As Long
'End Type
'Private m_VMemHandles() As TMemHandle
'Private m_Size  As Long
'Private m_Count As Long
''Ja so ist das Quatsch, die Klasse muss
''1. den VirtualMEmory verwalten
''2. alle Handles verwalten die den Speicher brauchen
'
'
'Private Sub Class_Initialize()
'    Dim si As SYSTEM_INFO: Call GetNativeSystemInfo(si)
'    m_Size = si.dwAllocationGranularity
'    ReDim m_VMemHandles(0 To 3)
'End Sub
'Private Sub Class_Terminate()
'    Dim i As Long
'    For i = 0 To m_Count - 1
'        With m_VMemHandles(i)
'            Call VirtualFree(.Handle, .Size, AllocationFlags.MemFree)
'        End With
'    Next
'End Sub
'Private Function NewAlloc(ByVal flAllocation As AllocationFlags, _
'                          ByVal flProtection As ProtectionFlags) As TMemHandle
'    With NewAlloc
'        .EAllocation = flAllocation
'        .EProtection = flProtection
'        .Size = m_Size
'        .Handle = VirtualAlloc(ByVal 0&, .Size, .EAllocation, .EProtection)
'        Debug.Print s
'    End With
'End Function
'
'Public Sub Add(ByVal Size As Long, Optional ByVal flAllocation As AllocationFlags = MemCommitReserve, _
'               Optional ByVal flProtection As ProtectionFlags = PageExecuteReadWrite)
'    Dim u As Long: u = UBound(m_VMemHandles)
'    If u < m_Count Then
'        ReDim Preserve m_VMemHandles(0 To 2 * u + 1)
'    End If
'    m_VMemHandles(m_Count) = NewAlloc(flAllocation, flProtection)
'    m_Count = m_Count + 1
'End Sub
'
'Public Property Get Item(Index As Long) As Long
'    Item = m_VMemHandles(Index).Handle
'End Property
'Public Property Get Size(Index As Long) As Long
'    Size = m_VMemHandles(Index).Size
'End Property
''erst dann ein neues erstellen, wenn das alte voll aufgebraucht ist
'
'
