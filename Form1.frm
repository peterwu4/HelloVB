VERSION 5.00
Begin VB.Form Ex7_Sub_Function 
   Caption         =   "Ex7_Sub_Function"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   5400
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command4 
      Caption         =   "Function"
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Variable life cycle"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   3960
      TabIndex        =   3
      Text            =   "10"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Subprogram"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Subroutine"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Ex1"
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Ex7_Sub_Function"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'Option Compare Binary

Private Sub Label1_Click()
    Dim testArr(4) As Integer
    Dim i As Integer
    For i = 0 To 3
        testArr(i) = i
    Next i
    
    If 1 + 1 = 2 Xor 1 + 1 = 11 Then
        Label1.Caption = "VB�d��"
        Ex7_Sub_Function.Caption = "Ex7_Sub_Function <== Title"
        Print UBound(testArr) '����
        MsgBox "���ߡI"
    Else
        MsgBox "�����ߡI"
    End If
        
End Sub

' �ϥΪ̦ۭq���    def fn... end def
' ���`��    gosub... return
' �Ƶ{��    sub... end sub
'                function... end function


' Subroutine
Private Sub Command1_Click()
    Dim k As Integer
    k = 1
    Cls
    Print "A Subroutine Test"
    Print "This is the starter k="; k
    GoSub addk
    Print "First time return from subroutine k="; k
    GoSub addk
    Print "Second time return from subroutine k="; k
    'End
    GoTo subend
    
addk:
    k = k + 1
    Return

subend:
End Sub

' Subprogram
Private Sub Command2_Click()
    'declare sub prodnum (x!)
    Dim x As Integer
    x = Val(Text1.Text)
    Call prodnum(x)
End Sub

Sub prodnum(x)
    Dim sum As Double
    Dim i As Integer
    sum = 1
    For i = 1 To x
        sum = sum * i
    Next i
    Print sum
End Sub

' �ܼƥͩR�P��
' �ܼƪ��ǻ�
' �ǭȩI�scall by value�G��޼ƻP��޼Ʀ��Τ��P���O����Ŷ���}�A
' ��޼Ƹ���ܰʡA��޼Ƥ��|�ܰʡC�i�H�O�d��޼Ƥ��e�C
' �޼Ƭ��`�ơB�B�⦡�ݩ�ǭȩI�s�A�Y�ܼƥ~�[()��ܥH�ǭȩI�s�ǻ��޼Ƹ�ơC
'   call ADD(3)
'   call ADD(x+3)
'   call ADD((x), (y))
' �ǧ}�I�scall by address�G��޼ƻP��޼Ʀ��άۦP���O����Ŷ��A
' ���޼Ƥ��e���ܡA��޼Ƥ]�H�ۧ��ܡC��޼ƪ��ȵL�k�O�s�C
' �޼Ƭ��ܼơB�}�C�ݩ�ǧ}�I�s
'   call ADD(x, y)
Private Sub Command3_Click()
    'declare sub test()
    Dim x As Integer
    Dim y As Integer
    Dim i As Integer
    Cls
    x = x + 1: y = y + 1
    For i = 1 To 5
        Call test(x, y)
    Next i
    For i = 1 To 5
        Call test2((x), (y))
    Next i
End Sub

Static Sub test(x, y)
    x = x + 1
    y = y + 1
    Print "x="; x, "y="; y
End Sub

Sub test2(x, y)
    x = x + 1
    y = y + 1
    Print "x="; x, "y="; y
End Sub

' ��Ʀ�Function
Private Sub Command4_Click()
    Dim x As Integer
    x = Val(Text1.Text)
    Print factorial&(x%)
End Sub

' factorial ����
Function factorial&(n%)
    If n% > 0 Then
        factorial& = n% * factorial&(n% - 1)
    Else
        factorial& = 1
    End If
End Function
