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
   StartUpPosition =   3  '�t�ιw�]��
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
      Caption         =   "�۩w���"
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
' �ҿר�ơA���{���y�����]�p�ɡA�K�w�w���]�p�n���B��{�ǡC
' �ϥΪ̥u�ݮM�θӵ{�ǡA�K�i�H�����o�쵲�G�C
' ��Ƥ@��i�H�����G
'   ���x��� Built-in Function
'   �ۭq��� User defined Function

Private Sub Command2_Click()
    ' built-in function
    ' �ƭȨ�ơG�Ǧ^�ƭȸ��
    Cls
    Print Abs(6); " "; Abs(-3)  ' �Ǧ^�����
    Print Sgn(-50); " "; Sgn(0); " "; Sgn(34)   ' �Ǧ^�Ÿ���
    Print Fix(34.253); " "; Fix(-25.13) ' �Ǧ^��ƭ�
    Print Int(24.32); " "; Int(-26.78)  ' �Ǧ^�p��ε��󪺳̤j��
    Print CInt(1.49); " "; CInt(1.51); " "; CInt(-1.46); " "; CInt(-1.56) ' �Ǧ^�|�ˤ��J����ƭ�
    Print Sqr(4)    ' �Ǧ^�����
    Print Sin(1)
    Print Cos(1)
    Print Tan(1)
    Print Exp(2)    ' �Ǧ^e^x��
    Print Log(2)    ' �Ǧ^�۵M��ƭ�LN(x)
    Print Asc("ABC")    ' �Ǧ^x$���Ĥ@�Ӧr����ASCII�X
    Print Val("1234")   ' �Nx$�ର�ƭ�
    Dim ab As Double
    Print Len("ABC��"); " "; Len(ab) ' �Ǧ^x$�άOx�ܼƪ��줸��
    'Print clen("ABC��") ' �Ǧ^x$���r����
End Sub

Private Sub Command3_Click()
    ' built-in function
    ' �r���ơG�Ǧ^�r����
    Cls
    Print Chr$(65)  ' �Ǧ^�r���X�ҥN���r��
    Print Str$(123)   ' �N�ƭȫ��A��Ƨאּ�r�����A
    Print LCase$("ABCdef")    ' �N�r���אּ�p�g
    Print UCase$("ABCdef")    ' �N�r���אּ�j�g
    Print "-"; Space$(5); "+"   ' �Ǧ^n�Ӫťզr��
    Print String$(5, "ABC")   ' ����x$�����Ĥ@�Ӧr��n��
    Print Left$("ABCDEF", 3)    ' �Ѧr��x$����}�l��n byte
    Print Right$("ABCDEF", 4)    ' �Ѧr��x$�k��n byte��m�}�l���Ѿlbyte
    Print Mid$("ABCDEF", 3, 2)  ' �Ѧr�x����n byte��m�}�l��m��byte
    'Print cmid$("ABC��DEF", 3, 2)  ' �Ѧr�x����n byte��m�}�l��m�Ӧr��
    Print "-"; LTrim$("    abcd   123       "); "+" ' �Nx$�����e�m�ťզr���R��
    Print "-"; RTrim$("    abcd   123       "); "+" ' �Nx$�����H���ťզr���R��
    Print Date$ ' ���o�t�Τ�� ��ܮ榡�Gmm-dd-yyyy
    Print Time$ ' ���o�t�ήɶ� ��ܮ榡�Ghh:mm:ss
    Print Hex$(16) ' �Ǧ^�N��x(10�i���)��16�i���r��
    Print Oct$(16) ' �Ǧ^�N��x(10�i���)��8�i���r��
End Sub

Private Sub Command1_Click()
    
    a = 10: b = 20
    c = addab(a, b)
    Cls
    Print c

End Sub

' �üƨ��
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
