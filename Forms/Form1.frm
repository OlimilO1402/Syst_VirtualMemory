VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   6135
   End
   Begin VB.CommandButton BtnCopyArrayFromVMem 
      Caption         =   "Copy Array From VMem"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton BtnCopyArrayToVMem 
      Caption         =   "Copy Array To V-Mem"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton BtnCallMsInfo32 
      Caption         =   "msinfo32.exe"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLength As Long)

Dim m_SysInfo As SystemInfo
Dim m_VMem    As VirtualMemory
Dim m_p0      As Long

Private Sub Form_Load()
    Set m_SysInfo = New SystemInfo
    Label1.Caption = m_SysInfo.ToStr
    Set m_VMem = New VirtualMemory
End Sub

Private Sub BtnCallMsInfo32_Click()
    m_SysInfo.CallSysInfoExe
End Sub

Private Sub BtnCopyArrayToVMem_Click()
    
    'create array fill with numbers
    Randomize
    Dim aSize As Long: aSize = Rnd * 4096 'in Bytes
    
    Dim u As Long: u = aSize \ 4 - 1
    ReDim Arr(0 To u) As Long
    Dim i As Long
    For i = 0 To u
        Arr(i) = i
    Next

    'allocate virtual memory the size of the array in bytes
    Dim p As Long: p = m_VMem.Alloc(aSize)
    List1.AddItem "Allocated: " & aSize & " Bytes"
    'copy the array to virtual memory
    RtlMoveMemory ByVal p, Arr(0), aSize
    
    If (m_VMem.PageSize Mod aSize) = 0 Then
        m_p0 = p
        List1.AddItem "p0: " & m_p0
    End If
    List1.AddItem "VMemAlloc: " & p & "   " & p - m_p0
    'm_ix = m_ix + 1
End Sub

Private Sub BtnCopyArrayFromVMem_Click()
    Randomize
    Dim ix As Long: ix = Rnd * m_VMem.PagesCount
    ix = IIf(ix = 0, 1, ix)
    Dim p As Long: p = m_VMem.Pointer(ix)
    Dim sz As Long: sz = m_VMem.SizeForIndex(ix)
    If p = 0 Then
        MsgBox "Not enough virtual memory for index: " & ix
        Exit Sub
    End If
    Dim u As Long: u = sz \ 4 - 1
    ReDim Arr(0 To u) As Long
    RtlMoveMemory Arr(0), ByVal p, sz
    'test if the values are all there
    Dim i As Long: i = Rnd * u
    Dim v As Long: v = Arr(i)
    MsgBox "On page " & ix & " on index " & i & " the value is " & Arr(i) & " this is " & IIf(i = v, "", "in") & "correct!"
    'on position i the value is i, if not there must be something wrong
    
End Sub
