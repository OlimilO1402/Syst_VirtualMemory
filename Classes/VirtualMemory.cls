VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VirtualMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub GetNativeSystemInfo Lib "kernel32" ( _
    ByRef lpSystemInfo As SYSTEM_INFO)
Private Declare Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As Long, ByRef dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" ( _
    ByRef lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function VirtualLock Lib "kernel32" ( _
    ByRef lpAddress As Any, ByVal dwSize As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" ( _
    ByRef lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
    
Private Type SYSTEM_INFO
    wProcessorArchitecture  As Integer
    wReserved               As Integer
    dwPageSize              As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask   As Long
    dwNumberOfProcessors    As Long
    dwProcessorType         As Long
    dwAllocationGranularity As Long
    wProcessorLevel         As Integer
    wProcessorRevision      As Integer
End Type

Public Enum AllocationFlags
    MemCommit = &H1000&             'MEM_COMMIT
    MemReserve = &H2000&            'MEM_RESERVE
    MemCommitReserve = &H3000&
    MemDecommit = &H4000&           'MEM_DECOMMIT
    MemRelease = &H8000&            'MEM_RELEASE
    MemFree = &H10000               'MEM_FREE
    MemPrivate = &H20000            'MEM_PRIVATE
    MemMapped = &H40000             'MEM_MAPPED
    MemReset = &H80000              'MEM_RESET
    MemTopDown = &H100000           'MEM_TOP_DOWN
    MemWriteWatch = &H200000        'MEM_WRITE_WATCH
    MemPhysical = &H400000          'MEM_PHYSICAL
    MemImage = &H1000000            'MEM_IMAGE
    Mem4MbPages = &H80000000        'MEM_4MB_PAGES

'Private Const MEM_E_INVALID_LINK 'As Long = &H80080010
'Private Const MEM_E_INVALID_ROOT 'As Long = &H80080009
'Private Const MEM_E_INVALID_SIZE 'As Long = &H80080011
    
End Enum
Public Enum ProtectionFlags
    PageNoAccess = &H1&             'PAGE_NOACCESS          As Long = &H1
    PageReadOnly = &H2&             'PAGE_READONLY          As Long = &H2
    PageReadWrite = &H4&            'PAGE_READWRITE         As Long = &H4
    PageWriteCopy = &H8&            'PAGE_WRITECOPY         As Long = &H8
    PageExecute = &H10&             'PAGE_EXECUTE           As Long = &H10
    PageExecuteRead = &H20&         'PAGE_EXECUTE_READ      As Long = &H20
    PageExecuteReadWrite = &H40&    'PAGE_EXECUTE_READWRITE As Long = &H40
    PageExecuteWriteCopy = &H80&    'PAGE_EXECUTE_WRITECOPY As Long = &H80
    PageGuard = &H100&              'PAGE_GUARD             As Long = &H100
    PageNoCache = &H200&            'PAGE_NOCACHE           As Long = &H200
    PageWiteCombine = &H400&        'PAGE_WRITECOMBINE      As Long = &H400
End Enum
Private Type TVMemory
    Pointer     As Long ' der Zeiger zum Speicher
    Size        As Long ' die Größe des Speicherbereichs
End Type
Private Type TVMemPage
    EAllocation As AllocationFlags
    EProtection As ProtectionFlags
    Arr()  As TVMemory
    Count  As Long ' die Anzahl der Elemente im Array
    Size   As Long ' die Summer aller Sizes in dem Momentanen Speicherbereich
End Type
Private m_VMemPages()  As TVMemPage
Private m_Capacity     As Long       ' die Gesamt-Speicherkapazität (= m_HandlesCount * m_AGranularity)
Private m_Size         As Long       ' die Summe aller hinzugefügten Alloc Size
Private m_PagesCount   As Long       ' die Anzahl der für den User zur Verfügung gestellten Speicherbereich-Handles
Private m_PageSize     As Long
Private m_AGranularity As Long       ' AllocationGranularity, die Größe eines einzelnen virtuellen Speicherberichs
'Private m_VMemorySize  As Long
'Ja so ist das Quatsch, die Klasse muss
'1. den VirtualMemory verwalten
'2. alle Handles verwalten die den Speicher brauchen


Private Sub Class_Initialize()
    Dim si As SYSTEM_INFO: Call GetNativeSystemInfo(si)
    m_AGranularity = si.dwAllocationGranularity
    m_PageSize = si.dwPageSize
End Sub
Private Sub Class_Terminate()
    '
    ' Speicher freigegeben wird nur ganz am Ende
    ' vorher kann man nichts löschen, da VirtualMemory normalerweise für ausführbaren Code ist,
    ' und der wird normalerweise ständig wiederverwendet
    '
    Dim hr As Long
    Dim i As Long
    For i = m_PagesCount - 1 To 0 Step -1
        With m_VMemPages(i)
            If .Count > 0 Then
                hr = VirtualFree(.Arr(0).Pointer, m_PageSize, AllocationFlags.MemFree)
                'If hr <> 0 Then MsgBox hr
                Debug.Print hr
            End If
        End With
    Next
End Sub
Public Property Get PageSize() As Long
    PageSize = m_PageSize
End Property
Public Property Get AllocationGranularity() As Long
    AllocationGranularity = m_AGranularity
End Property
Public Property Get Capacity() As Long
    Capacity = m_Capacity
End Property
Public Property Get Size() As Long
    Size = m_Size
End Property
Private Function AddPage(aPage As TVMemPage) As Long
    If m_Capacity = 0 Then
        ReDim m_VMemPages(0 To 3)
    Else
        Dim u As Long: u = UBound(m_VMemPages)
        If m_PagesCount > u Then
            ReDim Preserve m_VMemPages(0 To (u + 1) * 2 - 1)
        End If
    End If
    m_VMemPages(m_PagesCount) = aPage
    m_PagesCount = m_PagesCount + 1
    m_Capacity = m_PagesCount * m_PageSize 'AGranularity
    AddPage = aPage.Arr(0).Pointer
End Function
Private Function New_Page(ByVal Size As Long, _
                          ByVal flAl As AllocationFlags, _
                          ByVal flPr As ProtectionFlags) As TVMemPage
    With New_Page
        .Size = Size
        .EAllocation = flAl
        .EProtection = flPr
        ReDim .Arr(0 To 3)
        With .Arr(0)
            .Size = Size
            .Pointer = VirtualAlloc(ByVal 0&, m_PageSize, flAl, flPr)
        End With
        .Count = 1
    End With
End Function
Public Function Alloc(ByVal Size As Long, _
                      Optional ByVal flAllocat As AllocationFlags = AllocationFlags.MemCommitReserve, _
                      Optional ByVal flProtect As ProtectionFlags = ProtectionFlags.PageReadWrite) As Long
    Dim hr As Long
    If Size > m_PageSize Then
        'Definition: es soll nur für ausführbaren Speicher verwendet werden
        'deshalb hier eine Beschränkung
        'vielleicht machen wir das später wieder frei
        MsgBox "Can not allocate more than: " & CStr(m_PageSize) & " bytes at once"
        Exit Function
    End If
    If m_Capacity < m_Size + Size Then
        Alloc = AddPage(New_Page(Size, flAllocat, flProtect))
    Else
        With m_VMemPages(m_PagesCount - 1)
            'count ist hier nie null, weil schon in AddPage
            Dim u As Long: u = UBound(.Arr)
            If u < .Count Then
                ReDim Preserve .Arr(0 To (u + 1) * 2 - 1)
            End If
            Alloc = .Arr(0).Pointer + .Size
            With .Arr(.Count)
                .Pointer = Alloc
                .Size = Size
            End With
            .Size = .Size + Size
            .Count = .Count + 1
        End With
    End If
    m_Size = m_Size + Size
End Function
Public Property Get Pointer(ByVal Index As Long) As Long
    Dim c As Long
    Dim i As Long
    
    For i = 0 To m_PagesCount - 1
        c = c + m_VMemPages(i).Count
        If Index <= c Then Exit For
    Next
    With m_VMemPages(i) 'i - 1)
        c = c - Index
        Pointer = .Arr(c).Pointer
    End With
End Property
'Private Sub Add(ByVal Size As Long, _
'                Optional ByVal flAllocation As AllocationFlags = MemCommitReserve, _
'                Optional ByVal flProtection As ProtectionFlags = PageExecuteReadWrite)
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
'erst dann ein neues erstellen, wenn das alte voll aufgebraucht ist
