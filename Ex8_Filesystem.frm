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
   StartUpPosition =   3  '系統預設值
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
' 檔案的開啟
' 語法1：
'   open 檔案名稱 [for 模式1] [access 存取] [鎖定]
'               as [#]檔案代碼 [len=記錄長度]
'       模式1 指定存取模式
'               OUTPUT   循序輸出
'               INPUT       循序輸入
'               APPEND   循序增添
'               RANDOM   隨機存取
'               BINARY   二進位存取
'       存取 設定檔案的操作方式
'               READ    唯讀
'               WRITE   輸出
'               READWRITE   讀寫
'       鎖定 在multiprocessing環境下，限定其他processing存取所開啟的檔案
'               SHARED  共享
'               LOCK READ   唯讀鎖定
'               LOCK WRITE  寫入鎖定
'               LOCK READWRITE  讀寫鎖定
' Open "test.dat" For Output As #1
' Open "test.dat" For Random As #1 Len = 25
'
' 語法2：
'   open 模式2, [#] 檔案代碼, 檔案名稱 [, 記錄長度]
'       模式2 指定存取模式
'               O   循序輸出
'               I       循序輸入
'               A   循序增添
'               R   隨機存取
'               B   二進位存取
' open "o", #1, "test.dat"

' 檔案的關閉
' 語法：
'   close [[#]檔案代碼] [,[#]檔案代碼...]

' 與檔案有關的函數
'   EOF (1) end of file函數
'   Loc (1) 上一次存取資料的檔案指標位址。
'   LOF (1) length of file函數
'   seek(1)

Private Sub Command1_Click()
End Sub
