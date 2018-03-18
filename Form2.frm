VERSION 5.00
Begin VB.Form Ex6_Function 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command4 
      Caption         =   "Randomize function"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "built-in string function"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "built-in value function"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "自定函數"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "Ex6_Function"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 所謂函數，為程式語言為設計時，便已預先設計好的運算程序。
' 使用者只需套用該程序，便可以直接得到結果。
' 函數一般可以分為：
'   內儲函數 Built-in Function
'   自訂函數 User defined Function

Private Sub Command2_Click()
    ' built-in function
    ' 數值函數：傳回數值資料
    Cls
    Print Abs(6); " "; Abs(-3)  ' 傳回絕對值
    Print Sgn(-50); " "; Sgn(0); " "; Sgn(34)   ' 傳回符號值
    Print Fix(34.253); " "; Fix(-25.13) ' 傳回整數值
    Print Int(24.32); " "; Int(-26.78)  ' 傳回小於或等於的最大值
    Print CInt(1.49); " "; CInt(1.51); " "; CInt(-1.46); " "; CInt(-1.56) ' 傳回四捨五入的整數值
    Print Sqr(4)    ' 傳回平方根
    Print Sin(1)
    Print Cos(1)
    Print Tan(1)
    Print Exp(2)    ' 傳回e^x值
    Print Log(2)    ' 傳回自然對數值LN(x)
    Print Asc("ABC")    ' 傳回x$的第一個字元的ASCII碼
    Print Val("1234")   ' 將x$轉為數值
    Dim ab As Double
    Print Len("ABC我"); " "; Len(ab) ' 傳回x$或是x變數的位元數
    'Print clen("ABC我") ' 傳回x$的字元數
End Sub

Private Sub Command3_Click()
    ' built-in function
    ' 字串函數：傳回字串資料
    Cls
    Print Chr$(65)  ' 傳回字元碼所代表的字元
    Print Str$(123)   ' 將數值型態資料改為字元型態
    Print LCase$("ABCdef")    ' 將字元改為小寫
    Print UCase$("ABCdef")    ' 將字元改為大寫
    Print "-"; Space$(5); "+"   ' 傳回n個空白字元
    Print String$(5, "ABC")   ' 重複x$內的第一個字元n次
    Print Left$("ABCDEF", 3)    ' 由字串x$左邊開始取n byte
    Print Right$("ABCDEF", 4)    ' 由字串x$右邊n byte位置開始取剩餘byte
    Print Mid$("ABCDEF", 3, 2)  ' 由字軍左邊n byte位置開始取m個byte
    'Print cmid$("ABC我DEF", 3, 2)  ' 由字軍左邊n byte位置開始取m個字元
    Print "-"; LTrim$("    abcd   123       "); "+" ' 將x$內的前置空白字元刪除
    Print "-"; RTrim$("    abcd   123       "); "+" ' 將x$內尾隨的空白字元刪除
    Print Date$ ' 取得系統日期 顯示格式：mm-dd-yyyy
    Print Time$ ' 取得系統時間 顯示格式：hh:mm:ss
    Print Hex$(16) ' 傳回代表x(10進位制)的16進位制字串
    Print Oct$(16) ' 傳回代表x(10進位制)的8進位制字串
End Sub

Private Sub Command1_Click()
    
    a = 10: b = 20
    c = addab(a, b)
    Cls
    Print c

End Sub

' 亂數函數
Private Sub Command4_Click()
    Dim i As Integer
    Dim n As Double
    Randomize (Timer)
    Cls
    For i = 1 To 10
        n = Rnd
        Print CInt(n * 100); " "; n
    Next i
End Sub

Private Sub Form_Load()
'MsgBox c
End Sub

Function addab(a, b)
    addab = a + b
End Function
