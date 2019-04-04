VERSION 5.00
Begin VB.Form AdminForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "管理界面"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4590
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton buy 
      Caption         =   "成交"
      Height          =   495
      Left            =   2880
      TabIndex        =   20
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Turn 
      Caption         =   "更新"
      Height          =   495
      Left            =   2880
      TabIndex        =   19
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox NewMoney 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   18
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox buyName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Down 
      Caption         =   "下一个"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Up 
      Caption         =   "上一个"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   480
      Width           =   1095
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
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label PeopleNames 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   22
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label8 
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
      Left            =   2400
      TabIndex        =   21
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "拍卖价"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   4200
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      X1              =   4320
      X2              =   240
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label9 
      Caption         =   "竞买者"
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
      Left            =   240
      TabIndex        =   15
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Moneys 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "当前拍卖价"
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
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label buyNames 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "当前竞买者"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label MinMoneys 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
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
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label FirstMoneys 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
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
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Names 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   3615
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
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   615
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
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "拍卖软件 管理界面"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "AdminForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FormAdmin As Boolean
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
'定义实时变量
Dim nowMoney As Single

Private Sub LoadDat() '载入数据
'读取数据
If Number.Text < 1 Then
    Number.Text = 1
End If
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
Names.Caption = datName
PeopleNames.Caption = datPeopleName
FirstMoneys.Caption = datFirstMoney
MinMoneys.Caption = datMinMoney
'同步至大屏幕
Form1.Number.Caption = Number.Text
Form1.PeopleName.Caption = datPeopleName
Form1.Names.Caption = datName
Form1.Passage.Caption = datPassage
Form1.FirstMoney = datFirstMoney
Form1.MinMoney = datMinMoney
Form1.Money.Caption = datFirstMoney
Form1.buyName.Caption = ""
'载入图片
If Dir(App.Path + "\Images\" & Number.Text & ".jpg") <> "" Then
    Form1.Image1.Picture = LoadPicture(App.Path + "\Images\" & Number.Text & ".jpg")
ElseIf Dir(App.Path + "\Images\" & Number.Text & ".jpeg") <> "" Then
    Form1.Image1.Picture = LoadPicture(App.Path + "\Images\" & Number.Text & ".jpeg")
Else
    Form1.Image1.Picture = LoadPicture
End If
'实时变量赋值
nowMoney = datFirstMoney
'判断是否成交
For intI = 8 To 9
    If Trim$(objImportSheet.Cells(intCountI, intI)) <> "" Then
        Form1.State.Caption = "已成交"
        Form1.State.BackColor = &H8080FF
        Form1.Money.Caption = objImportSheet.Cells(intCountI, 9)
        Moneys.Caption = objImportSheet.Cells(intCountI, 9)
        Form1.buyName.Caption = objImportSheet.Cells(intCountI, 8)
        buyNames.Caption = objImportSheet.Cells(intCountI, 8)
        nowMoney = objImportSheet.Cells(intCountI, 9)
    Else
        Form1.State.Caption = "拍卖中"
        Form1.State.BackColor = &H80FF80
        Form1.Money.Caption = datMinMoney
        Form1.buyName.Caption = ""
        buyNames.Caption = ""
        Moneys.Caption = datMinMoney
    End If
Next intI
End Sub

Private Sub buy_Click()
On Error GoTo error
'成交
objImportSheet.Cells(Number.Text + 1, 8) = buyNames.Caption
objImportSheet.Cells(Number.Text + 1, 9) = nowMoney
Form1.State.Caption = "已成交"
Form1.State.BackColor = &H8080FF
buyName.Text = ""
NewMoney.Text = ""
buyName.SetFocus
Exit Sub
error:
MsgBox "数据类型错误"
End Sub

Private Sub buyName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    NewMoney.SetFocus
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormAdmin = False
objWorkBook.SaveAs App.Path + "\data.xlsx"
'结束Excel模块
objExcelFile.Quit
Set objWorkBook = Nothing
Set objImportSheet = Nothing
Set objExcelFile = Nothing
End Sub

Private Sub NewMoney_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Turn_Click
End If
End Sub

Private Sub Number_KeyPress(KeyAscii As Integer)
On Error GoTo error
If KeyAscii = 13 Then
    buyName.SetFocus
    Call LoadDat
End If
error:
MsgBox "数据类型错误"
End Sub

Private Sub Turn_Click()
On Error GoTo error
'更新数据
If NewMoney.Text = "" Or buyName.Text = "" Then
    MsgBox "数值为空！"
Else
    If NewMoney.Text - nowMoney < datMinMoney Then
        MsgBox "小于最小增幅！"
    Else
        nowMoney = NewMoney.Text
        buyNames.Caption = buyName.Text
        Moneys.Caption = NewMoney.Text
        Form1.Money.Caption = NewMoney.Text
        Form1.buyName.Caption = buyName.Text
    End If
End If
buyName.Text = ""
NewMoney.Text = ""
buyName.SetFocus
Exit Sub
error:
MsgBox "数据类型错误"
End Sub

Private Sub Down_Click()
On Error GoTo error
Number.Text = Number.Text + 1
Call LoadDat
Exit Sub
error:
MsgBox "数据类型错误"
End Sub

Private Sub Form_Load()
On Error GoTo error
FormAdmin = True
'加载Excel模块
Set objExcelFile = New Excel.Application
objExcelFile.DisplayAlerts = False
Set objWorkBook = objExcelFile.Workbooks.Open(App.Path + "\data.xlsx")
Set objImportSheet = objWorkBook.Sheets(1)
'获取行数
Number.Text = 1
Call LoadDat
Exit Sub
error:
MsgBox "数据类型错误"
End Sub

Private Sub Up_Click()
On Error GoTo error
If Number.Text > 1 Then
    Number.Text = Number.Text - 1
End If
Call LoadDat
Exit Sub
error:
MsgBox "数据类型错误"
End Sub
