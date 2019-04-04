VERSION 5.00
Begin VB.Form Admin2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "数据库管理"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4530
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton OpenImage 
      Caption         =   "打开图片文件夹"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   18
      Top             =   5400
      Width           =   1935
   End
   Begin VB.TextBox Number 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Names 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox FirstMoneys 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox MinMoneys 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Passages 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3360
      Width           =   4215
   End
   Begin VB.CommandButton Open 
      Caption         =   "打开"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Save 
      Caption         =   "保存"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Up 
      Caption         =   "上一个"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton down 
      Caption         =   "下一个"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox PeopleNames 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "图片请放置在软件根目录下的""Image""文件夹下，并命名为相对应的序号（纯数字），目前只支持""*.jpg""""*.jpeg""格式的图片。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   4680
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "拍卖软件 数据库管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "序号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "最低增幅"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "起拍价"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "介绍（200字以内）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "拍卖者"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
End
Attribute VB_Name = "Admin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FormAdmin2 As Boolean
Dim datNo As Integer
Dim datPeopleName As String
Dim datName As String
Dim datPassage As String
Dim datFirstMoney As Single
Dim datMinMoney As Single
'定义Excel模块变量
Dim objExcelFile As Excel.Application
Dim objWorkBook As Excel.Workbook
Dim objImportSheet As Excel.Worksheet

Private Sub LoadDat()
intCountI = Number.Text + 1
'Check if Empty Data Row
blnNullRow = True
'如果第1到第10个单元格的值均为空或空格，则视为空行
For intI = 1 To 6
    If Trim$(objImportSheet.Cells(intCountI, intI).Value) <> "" Then
        blnNullRow = False
    Else
        datName = ""
        datPeopleName = ""
        datFirstMoney = 0
        datMinMoney = 0
        datPassage = ""
    End If
Next intI
'若不是空行，则进行读取动作，否则继续向后遍历Excel中的行
If blnNullRow = False Then
    '获取单元格中的数据，做有效性Check，并将合法数据创建为实体存入对象数组中
    datName = objImportSheet.Cells(intCountI, 2)
    datPeopleName = objImportSheet.Cells(intCountI, 3)
    datFirstMoney = objImportSheet.Cells(intCountI, 4)
    datMinMoney = objImportSheet.Cells(intCountI, 5)
    datPassage = objImportSheet.Cells(intCountI, 6)
End If
'读取数据
Names.Text = datName
PeopleNames.Text = datPeopleName
Passages.Text = datPassage
FirstMoneys.Text = datFirstMoney
MinMoneys.Text = datMinMoney
Number.SetFocus
End Sub

Private Sub OpenImage_Click()
If Dir(App.Path + "\Images", vbDirectory) = "" Then
    MkDir (App.Path + "\Images")
End If
Shell "explorer " + App.Path + "\Images", 1
End Sub

Private Sub Down_Click()
Number.Text = Number.Text + 1
Call LoadDat
End Sub

Private Sub FirstMoneys_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
MinMoneys.SetFocus
End If
End Sub

Private Sub Form_Load()
FormAdmin2 = True
'加载Excel模块
Set objExcelFile = New Excel.Application
objExcelFile.DisplayAlerts = False
Set objWorkBook = objExcelFile.Workbooks.Open(App.Path + "\data.xlsx")
Set objImportSheet = objWorkBook.Sheets(1)
'获取行数
Number.Text = 1
Call LoadDat
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormAdmin2 = False
objWorkBook.SaveAs App.Path + "\data.xlsx"
'结束Excel模块
objExcelFile.Quit
Set objWorkBook = Nothing
Set objImportSheet = Nothing
Set objExcelFile = Nothing
End Sub

Private Sub MinMoneys_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Passages.SetFocus
End If
End Sub

Private Sub Names_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
PeopleNames.SetFocus
End If
End Sub

Private Sub Number_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Open_Click
Names.SetFocus
End If
End Sub

Private Sub Open_Click()
'读取数据
If Number.Text < 1 Then
    Number.Text = 1
End If
Call LoadDat
End Sub

Private Sub PeopleNames_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
FirstMoneys.SetFocus
End If
End Sub

Private Sub Save_Click()
On Error GoTo error
intCountI = Number.Text + 1
'写入数据
datName = Names.Text
datPeopleName = PeopleNames.Text
datPassage = Passages.Text
datFirstMoney = FirstMoneys.Text
datMinMoney = MinMoneys.Text
objImportSheet.Cells(intCountI, 1) = Number.Text
objImportSheet.Cells(intCountI, 2) = datName
objImportSheet.Cells(intCountI, 3) = datPeopleName
objImportSheet.Cells(intCountI, 4) = datFirstMoney
objImportSheet.Cells(intCountI, 5) = datMinMoney
objImportSheet.Cells(intCountI, 6) = datPassage
Exit Sub
error:
MsgBox "数据类型错误"
End Sub

Private Sub Up_Click()
Number.Text = Number.Text - 1
If Number.Text < 1 Then
    Number.Text = 1
End If
Call LoadDat
End Sub
