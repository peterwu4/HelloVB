VERSION 5.00
Begin VB.Form Ex8_Filesystem 
   Caption         =   "Ex8_FileSystem"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   7005
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   5280
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Ex8_Filesystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' �ɮת��}��
' �y�k1�G
'   open �ɮצW�� [for �Ҧ�1] [access �s��] [��w]
'               as [#]�ɮץN�X [len=�O������]
'       �Ҧ�1 ���w�s���Ҧ�
'               OUTPUT   �`�ǿ�X
'               INPUT       �`�ǿ�J
'               APPEND   �`�ǼW�K
'               RANDOM   �H���s��
'               BINARY   �G�i��s��
'       �s�� �]�w�ɮת��ާ@�覡
'               READ    ��Ū
'               WRITE   ��X
'               READWRITE   Ū�g
'       ��w �bmultiprocessing���ҤU�A���w��Lprocessing�s���Ҷ}�Ҫ��ɮ�
'               SHARED  �@��
'               LOCK READ   ��Ū��w
'               LOCK WRITE  �g�J��w
'               LOCK READWRITE  Ū�g��w
' Open "test.dat" For Output As #1
' Open "test.dat" For Random As #1 Len = 25
'
' �y�k2�G
'   open �Ҧ�2, [#] �ɮץN�X, �ɮצW�� [, �O������]
'       �Ҧ�2 ���w�s���Ҧ�
'               O   �`�ǿ�X
'               I       �`�ǿ�J
'               A   �`�ǼW�K
'               R   �H���s��
'               B   �G�i��s��
' open "o", #1, "test.dat"

' �ɮת�����
' �y�k�G
'   close [[#]�ɮץN�X] [,[#]�ɮץN�X...]

' �P�ɮצ��������
'   EOF (1) end of file���
'   Loc (1) �W�@���s����ƪ��ɮ׫��Ц�}�C
'   LOF (1) length of file���
'   seek(1)

Private Sub Command1_Click()
End Sub
