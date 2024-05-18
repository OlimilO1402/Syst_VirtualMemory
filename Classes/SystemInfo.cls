VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub GetSystemInfo Lib "kernel32" (ByRef lpSystemInfo As SYSTEM_INFO)

Private Declare Sub GetNativeSystemInfo Lib "kernel32" (ByRef lpSystemInfo As SYSTEM_INFO)

Private Declare Function RegOpenKeyExA Lib "advapi32" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegOpenKeyA Lib "advapi32" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegQueryValueExA Lib "advapi32" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Const sCPURegKey = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const REG_SZ     As Long = 1
Private Const REG_BINARY As Long = 3
Private Const REG_DWORD  As Long = 4
Private Const ERROR_SUCCESS          As Long = 0&
'Private Const STANDARD_RIGHTS_ALL    As Long = &H1F0000
'Private Const KEY_QUERY_VALUE        As Long = &H1&
'Private Const KEY_SET_VALUE          As Long = &H2&
'Private Const KEY_CREATE_SUB_KEY     As Long = &H4&
'Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8&
'Private Const KEY_NOTIFY             As Long = &H10&
'Private Const KEY_CREATE_LINK        As Long = &H20&
'Private Const SYNCHRONIZE            As Long = &H100000
'Private Const KEY_ALL_ACCESS         As Long = (( _
'    STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or _
'    KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
    
Private Const KEY_QUERY_VALUE As Long = &H1&
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8&
Private Const KEY_NOTIFY As Long = &H10&

Private Const KEY_READ = KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
    
    
Private Type SYSTEM_INFO
    wProcessorArchitecture  As Integer
    wReserved               As Integer
    dwPageSize              As Long
    lpMinimumApplicationAddress As Long 'unsigned
    lpMaximumApplicationAddress As Long 'unsigned
    dwActiveProcessorMask   As Long
    dwNumberOfProcessors    As Long
    dwProcessorType         As Long
    dwAllocationGranularity As Long
    wProcessorLevel         As Integer
    wProcessorRevision      As Integer
End Type
Private Type ULong
    Value As Long
End Type
Private Type LongLong
    Value As Currency
End Type
Private m_System_Info As SYSTEM_INFO
Private m_CPUSpeed    As Long
Private m_CPUName     As String
'Private Const PROCESSOR_INTEL_386     As Long = 386
'Private Const PROCESSOR_INTEL_486     As Long = 486
'Private Const PROCESSOR_INTEL_PENTIUM As Long = 586
'Private Const PROCESSOR_INTEL_IA64    As Long = 2200
'Private Const PROCESSOR_AMD_X8664     As Long = 8664
'Private Const PROCESSOR_ARM (Reserved)

Private Sub Class_Initialize()
    'GetSystemInfo m_System_Info
    'If iswow64 Then
    GetNativeSystemInfo m_System_Info
    m_CPUSpeed = GetCPUSpeed
    m_CPUName = GetCPUName
End Sub

'Public Property Get OemID() As Long
'    OemID = m_System_Info.dwOemID
'    'entweder OemID oder
'    'wProcessorArchitecture
'    'wReserved
'End Property

Public Sub CallSysInfoExe()
    Shell "msinfo32.exe"
End Sub

Public Property Get ProcessorArchitecture() As String
    Dim s As String
    Select Case m_System_Info.wProcessorArchitecture
    Case 0: s = "Intel x86"
    Case 6: s = "Intel Itanium-based"
    Case 9: s = "x64 (AMD or Intel)"
    Case &HFFFF: s = "Unknown architecture"
    End Select
    ProcessorArchitecture = s
End Property

Private Function GetCPUSpeed() As Long
    Dim hKey As Long
    Dim hr As Long
    hr = RegOpenKeyA(HKEY_LOCAL_MACHINE, sCPURegKey, hKey)
    Dim CPUSpeed As Long
    hr = RegQueryValueExA(hKey, "~MHz", 0, 0, CPUSpeed, 4)
    hr = RegCloseKey(hKey)
    GetCPUSpeed = CPUSpeed
End Function

Public Property Get CPUSpeed() As Long
    CPUSpeed = m_CPUSpeed
End Property

Private Function GetCPUName() As String
    
    Dim hr As Long, hKey As Long
    hr = RegOpenKeyExA(HKEY_LOCAL_MACHINE, sCPURegKey, 0&, KEY_READ, hKey)
    If hr <> ERROR_SUCCESS Then Exit Function
    
    Dim RegType As Long, sLen As Long, sField As String: sField = "ProcessorNameString"
    hr = RegQueryValueExA(hKey, sField, 0&, RegType, ByVal 0&, sLen)
    If hr <> ERROR_SUCCESS Then
        hr = RegCloseKey(hKey)
        Exit Function
    End If
    
    If RegType = REG_SZ Then
        GetCPUName = Space$(sLen) & vbNullChar
        hr = RegQueryValueExA(hKey, sField, 0&, RegType, ByVal GetCPUName, sLen)
        If hr <> ERROR_SUCCESS Then
            hr = RegCloseKey(hKey)
            Exit Function
        End If
        GetCPUName = Left$(GetCPUName, sLen)
    End If
    hr = RegCloseKey(hKey)
        
End Function
Public Property Get CPUName() As String
    CPUName = m_CPUName
End Property

Public Property Get PageSize() As Long
    PageSize = m_System_Info.dwPageSize
End Property

Public Property Get MinimumApplicationAddress() As Long
    MinimumApplicationAddress = m_System_Info.lpMinimumApplicationAddress
End Property
Private Function Long_ToULong(ByVal l As Long) As Variant 'Decimal!
    Dim ul As ULong:   ul.Value = l
    Dim ll As LongLong: LSet ll = ul
    Long_ToULong = ll.Value * 10000
End Function
Public Property Get MaximumApplicationAddress()
    'must be ULong
    MaximumApplicationAddress = Long_ToULong(m_System_Info.lpMaximumApplicationAddress)
End Property
Public Property Get ActiveProcessorMask() As Long
    ActiveProcessorMask = m_System_Info.dwActiveProcessorMask
End Property
Public Property Get NumberOfProcessors() As Long
    NumberOfProcessors = m_System_Info.dwNumberOfProcessors
End Property
Public Property Get AllocationGranularity() As Long
    AllocationGranularity = m_System_Info.dwAllocationGranularity
End Property
Public Property Get ProcessorLevel() As Long
    ProcessorLevel = m_System_Info.wProcessorLevel
End Property
Public Property Get ProcessorRevision() As Integer
    ProcessorRevision = m_System_Info.wProcessorRevision
End Property

Private Function HiByte(ByVal i As Integer) As Byte
    HiByte = (i And &HFF00&) \ (&H100)
End Function
Private Function LoByte(ByVal i As Integer) As Byte
    LoByte = i And &HFF&
End Function

Public Function ToStr() As String
    Dim s As String
    'With m_System_Info
    With Me
        s = s & "ProcessorArchitecture: " & .ProcessorArchitecture & vbCrLf
        s = s & "NumberOfProcessors   : " & CStr(.NumberOfProcessors) & vbCrLf
        s = s & "ActiveProcessorMask  : " & CStr(.ActiveProcessorMask) & vbCrLf
        s = s & "ProcessorLevel       : " & CStr(.ProcessorLevel) & vbCrLf
        s = s & "ProcessorRevision    : " & CStr(.ProcessorRevision) & " &&H" & Hex(.ProcessorRevision) & vbCrLf
        s = s & "PageSize             : " & CStr(.PageSize) & vbCrLf
        s = s & "Min App Address      : " & CStr(.MinimumApplicationAddress) & vbCrLf
        s = s & "Max App Address      : " & CStr(.MaximumApplicationAddress) & vbCrLf
        s = s & "AllocationGranularity: " & CStr(.AllocationGranularity) & vbCrLf
        s = s & "CPUSpeed             : " & CStr(.CPUSpeed) & "MHz" & vbCrLf
        s = s & "Family : " & CStr(.ProcessorLevel) & ", Model: " & HiByte(.ProcessorRevision) & ", Stepping: " & LoByte(.ProcessorRevision) & vbCrLf
        s = s & "CPU : " & .CPUName & vbCrLf
    End With
    ToStr = s
End Function

'SYSTEM_INFO structure
'Contains information about the current computer system. This includes the architecture and type of the processor, the number of processors in the system, the page size, and other such information.
'
'Syntax
'
'typedef struct _SYSTEM_INFO {
'  union {
'    DWORD  dwOemId;
'    struct {
'      WORD wProcessorArchitecture;
'      WORD wReserved;
'    };
'  };
'  DWORD     dwPageSize;
'  LPVOID    lpMinimumApplicationAddress;
'  LPVOID    lpMaximumApplicationAddress;
'  DWORD_PTR dwActiveProcessorMask;
'  DWORD     dwNumberOfProcessors;
'  DWORD     dwProcessorType;
'  DWORD     dwAllocationGranularity;
'  WORD      wProcessorLevel;
'  WORD      wProcessorRevision;
'} SYSTEM_INFO;
'
'Members
' dwOemID
'   An obsolete member that is retained for compatibility. Applications should use the wProcessorArchitecture branch of the union.
' wProcessorArchitecture
'   The processor architecture of the installed operating system. This member can be one of the following values.
'
'Value                                      Meaning
'PROCESSOR_ARCHITECTURE_AMD64 9             x64 (AMD Or Intel)
'PROCESSOR_ARCHITECTURE_IA64 6              Intel Itanium - based
'PROCESSOR_ARCHITECTURE_INTEL 0             x86
'PROCESSOR_ARCHITECTURE_UNKNOWN 0xffff      Unknown architecture.
'
' wReserved
'   This member is reserved for future use.
' dwPageSize
'   The page size and the granularity of page protection and commitment. This is the page size used by the VirtualAlloc function.
' lpMinimumApplicationAddress
'   A pointer to the lowest memory address accessible to applications and dynamic-link libraries (DLLs).
' lpMaximumApplicationAddress
'   A pointer to the highest memory address accessible to applications and DLLs.
' dwActiveProcessorMask
'   A mask representing the set of processors configured into the system. Bit 0 is processor 0; bit 31 is processor 31.
' dwNumberOfProcessors
'   The number of logical processors in the current group. To retrieve this value, use the GetLogicalProcessorInformation function.
'   Note
'   For information about the physical processors shared by logical processors, call GetLogicalProcessorInformationEx with the RelationshipType parameter set to RelationProcessorPackage (3).
'
' dwProcessorType
'   An obsolete member that is retained for compatibility.
'   Use the wProcessorArchitecture, wProcessorLevel, and wProcessorRevision members to determine the type of processor.
'   PROCESSOR_INTEL_386 (386)
'   PROCESSOR_INTEL_486 (486)
'   PROCESSOR_INTEL_PENTIUM (586)
'   PROCESSOR_INTEL_IA64 (2200)
'   PROCESSOR_AMD_X8664 (8664)
' dwAllocationGranularity
'   The granularity for the starting address at which virtual memory can be allocated. For more information, see VirtualAlloc.
' wProcessorLevel
'   The architecture-dependent processor level. It should be used only for display purposes. To determine the feature set of a processor, use the
'   IsProcessorFeaturePresent function.
'   If wProcessorArchitecture is PROCESSOR_ARCHITECTURE_INTEL, wProcessorLevel is defined by the CPU vendor.
'   If wProcessorArchitecture is PROCESSOR_ARCHITECTURE_IA64, wProcessorLevel is set to 1.
' wProcessorRevision
'   The architecture-dependent processor revision. The following table shows how the revision value is assembled for each type of processor
'   architecture.

' Processor      Value
'   Intel Pentium,
'   Cyrix, or
'   NextGen 586
'
'       The high byte is the model and the low byte is the stepping.
'       For example, if the value is xxyy, the model number and stepping can be displayed as follows:
'           Model xx, Stepping yy
'   Intel 80386 Or 80486
'       A value of the form xxyz.
'       If xx is equal to 0xFF, y - 0xA is the model number, and z is the stepping identifier.
'       If xx is not equal to 0xFF, xx + 'A' is the stepping letter and yz is the minor stepping.
'Examples
'
'For an example, see Getting Hardware Information.
'
'Requirements
'Minimum supported client: Windows 2000 Professional
'Minimum supported server: Windows 2000 Server
'
'Header
'Winbase.h (include Windows.h)
'
'See also
' GetNativeSystemInfo GetSystemInfo MapViewOfFile MapViewOfFileEx
'
'Getting Hardware Information
'
'The following example uses the GetSystemInfo function to obtain hardware information such as the OEM identifier, processor type, page size, and so on. The example displays the information in the console.
'
'#include <windows.h>
'#include <stdio.h>
'#pragma comment(lib, "user32.lib")
'
'void main()
'{
'   SYSTEM_INFO siSysInfo;
'
'   // Copy the hardware information to the SYSTEM_INFO structure.
'
'   GetSystemInfo(&siSysInfo);
'
'   // Display the contents of the SYSTEM_INFO structure.
'
'   printf("Hardware information: \n");
'   printf("  OEM ID: %u\n", siSysInfo.dwOemId);
'   printf("  Number of processors: %u\n",
'      siSysInfo.dwNumberOfProcessors);
'   printf("  Page size: %u\n", siSysInfo.dwPageSize);
'   printf("  Processor type: %u\n", siSysInfo.dwProcessorType);
'   printf("  Minimum application address: %lx\n",
'      siSysInfo.lpMinimumApplicationAddress);
'   printf("  Maximum application address: %lx\n",
'      siSysInfo.lpMaximumApplicationAddress);
'   printf("  Active processor mask: %u\n",
'      siSysInfo.dwActiveProcessorMask);
'}
'
'
