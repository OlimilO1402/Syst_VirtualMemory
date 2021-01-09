VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14520
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   14520
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
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
      Top             =   5280
      Width           =   6135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   4800
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
      Height          =   4560
      Left            =   120
      TabIndex        =   4
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

Dim mSysInfo As SystemInfo
Dim mVMem As VirtualMemory
Dim p0 As Long
Dim s As Long
Dim Arr() As Long
Dim ix As Long

Private Sub Form_Load()
    Set mSysInfo = New SystemInfo
    Label1.Caption = mSysInfo.ToStr
    Set mVMem = New VirtualMemory
    s = 512 '2048
    ReDim Arr(0 To s \ 4 - 1)
    Dim i As Long
    For i = 0 To s \ 4 - 1
        Arr(i) = i
    Next
End Sub

Private Sub Command3_Click()
    Dim p As Long: p = mVMem.Alloc(s)
    RtlMoveMemory ByVal p, Arr(0), s
    If (mVMem.PageSize Mod s) = 0 Then
        p0 = p
        List1.AddItem "p0: " & p0
    End If
    List1.AddItem "VMemAlloc: " & p & "   " & p - p0
    ix = ix + 1
End Sub

Private Sub Command4_Click()
    Dim i As Long: i = 3
    Dim p As Long: p = mVMem.Pointer(i)
    If p = 0 Then
        MsgBox "Not enough virtual memory for index: " & i
        Exit Sub
    End If
    ReDim marr(0 To s \ 4 - 1) As Long
    RtlMoveMemory marr(0), ByVal p, s
    MsgBox marr(67)
    
End Sub

Sub TextAdd(s As String)
    'text1.Text = text1.Text & s & vbCrLf
    List1.AddItem s
End Sub
