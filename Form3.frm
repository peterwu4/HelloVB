VERSION 5.00
Begin VB.Form Ex4_Flow_Control 
   Caption         =   "Form3"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8745
   LinkTopic       =   "Form3"
   ScaleHeight     =   4410
   ScaleWidth      =   8745
   StartUpPosition =   3  '╰参箇砞
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Text            =   "12"
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Ex4_Flow_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim a As Integer
    Dim Name As String
    Dim salary As Double
    a% = 10
    salary = 54321.123
    Name = "peter wu"
    
    ' if else end if
    Cls
    a% = Text1.Text
    If a Mod 2 Then
        Print "块计"
    Else
        Print "块案计"
    End If
    
    ' select case
    Select Case a%
    Case 100
        Print "Best"
    Case Is > 90
        Print "Good"
    Case 60 To 89
        Print "so so"
    Case Is < 60
        Print "bad"
    End Select
    
    Print a; Spc(5); Name; Tab(1); salary
   
   ' for next
    Print " "
    Print " "
    For x = 1 To 9
        For y = 1 To 9
            Print x; "*"; y; "="; x * y; " ";
        Next y
        Print " "
    Next x
    
    ' do while loop
    Print
    s = 0: x = 1
    Do While x <= a%
        s = s + x
        x = x + 1
    Loop
        
    Print "1+ ... +"; a%; "="; s
    
    ' do loop while
    s2 = 0
    x = a%
    Do
        s2 = s2 + x
        x = x - 1
    Loop While x >= 0
    
    Print a%; "+...+1="; s2
    
    ' while wend
    Print
    bincode$ = ""
    Source% = a
    While Source <> 0
        remain = Source Mod 2
        Source = Source \ 2
        bincode$ = Str$(remain) + bincode$
    Wend
    Print bincode$
End Sub
