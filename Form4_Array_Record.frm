VERSION 5.00
Begin VB.Form Ex5_Array_Record 
   Caption         =   "Form4"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command2 
      Caption         =   "Type demo"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Array"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "Ex5_Array_Record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' dim �}�C�W��(�U���� to �W����, �U���� to �W����) [as type]
' option base n (n=0,1) ���ܤ��w�}�C���ФU���Ȭ�0��1
' �R�A�}�C�G�b�{���sĶ�ɫK�t�m�O����Ŷ�; �H�`�ƫŧi�}�C����
' �ʺA�}�C�G�b�{������ɤ~�t�m�O����Ŷ�; �H�ܼƫŧi�}�C����
' �w�]��ȡG�r�ꬰ�Ŧr��A�ƭȬ�0
' erase�G�R�A�}�C�gerase�ᬰ�M���}�C���e�F�ʺA�}�C������O����Ŷ�
Option Base 1

Private Sub Command1_Click()
    Cls
    
    
    'Dim x As Integer: x = 3
    'Dim y As Integer: y = 3
    'Dim a(x, y)
    Dim a(3, 3) As Integer
    For i = 1 To 3
        For j = 1 To 3
            a(i, j) = 4 * i + j
        Next j, i
            
    Call output(a, 3, 3)
    
    Erase a
    
    Call output(a, 3, 3)
    
End Sub

Private Function output(a, x, y)
    For i = 1 To x
        For j = 1 To y
            Print a(i, j); " ";
        Next j
        Print
    Next i
End Function

Private Sub Command2_Click()

    Dim students(2) As ScoreRec
    
    students(1).rName = "Peter"
    students(1).rChi = 80
    students(1).rMath = 70
    students(2).rName = "Mary"
    students(2).rChi = 66
    students(2).rMath = 77
    For i = 1 To 2
        students(i).rAve = (students(i).rChi + students(i).rMath) / 2
    Next i
    
    Cls
    For i = 1 To 2
        Print students(i).rName; " "; students(i).rChi; " "; students(i).rMath; " "; students(i).rAve
    Next i
End Sub
