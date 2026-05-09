VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    Dim pFunc As LongPtr, result As Long, x As Long: x = Rnd * 10 + 1
    Dim bRnd As Boolean: bRnd = Rnd() > 0.5
    
    pFunc = IIf(bRnd, FncPtr(AddressOf MFunctions.AddFive), FncPtr(AddressOf MFunctions.TimesSeven))
    Dim FuncName As String: FuncName = IIf(bRnd, "AddFive", "TimesSeven")
    
    If Thread_CallFunctionPtr(pFunc, VarPtr(x), result) Then
        Debug_Print FuncName & "(" & x & ") = " & result
    Else
        MsgBox "Error calling function."
        Exit Sub
    End If
    
    pFunc = MFncPtr.MakeFunction(MFunctions.FuncTimesFive_GetAsm())
    If pFunc = 0 Then
        MsgBox "Error generating function."
        Exit Sub
    End If
    If Thread_CallFunctionPtr(pFunc, VarPtr(x), result) Then
        Debug_Print "TimesFive(" & x & ") = " & result
    Else
        MsgBox "Error calling function."
        Exit Sub
    End If
    
    ReleaseFn pFunc
    
End Sub

Sub Debug_Print(ByVal s As String)
    Text1.Text = Text1.Text & s & vbCrLf
End Sub

