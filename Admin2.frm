VERSION 5.00
Begin VB.Form Admin2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ݿ����"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton OpenImage 
      Caption         =   "��ͼƬ�ļ���"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "��"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Save 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��һ��"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton down 
      Caption         =   "��һ��"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox PeopleNames 
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ͼƬ������������Ŀ¼�µ�""Image""�ļ����£�������Ϊ���Ӧ����ţ������֣���Ŀǰֻ֧��""*.jpg""""*.jpeg""��ʽ��ͼƬ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "������� ���ݿ����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���ļ�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���ܣ�200�����ڣ�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
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
'����Excelģ�����
Dim objExcelFile As Excel.Application
Dim objWorkBook As Excel.Workbook
Dim objImportSheet As Excel.Worksheet

Private Sub LoadDat()
intCountI = Number.Text + 1
'Check if Empty Data Row
blnNullRow = True
'�����1����10����Ԫ���ֵ��Ϊ�ջ�ո�����Ϊ����
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
'�����ǿ��У�����ж�ȡ�������������������Excel�е���
If blnNullRow = False Then
    '��ȡ��Ԫ���е����ݣ�����Ч��Check�������Ϸ����ݴ���Ϊʵ��������������
    datName = objImportSheet.Cells(intCountI, 2)
    datPeopleName = objImportSheet.Cells(intCountI, 3)
    datFirstMoney = objImportSheet.Cells(intCountI, 4)
    datMinMoney = objImportSheet.Cells(intCountI, 5)
    datPassage = objImportSheet.Cells(intCountI, 6)
End If
'��ȡ����
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
'����Excelģ��
Set objExcelFile = New Excel.Application
objExcelFile.DisplayAlerts = False
Set objWorkBook = objExcelFile.Workbooks.Open(App.Path + "\data.xlsx")
Set objImportSheet = objWorkBook.Sheets(1)
'��ȡ����
Number.Text = 1
Call LoadDat
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormAdmin2 = False
objWorkBook.SaveAs App.Path + "\data.xlsx"
'����Excelģ��
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
'��ȡ����
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
'д������
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
MsgBox "�������ʹ���"
End Sub

Private Sub Up_Click()
Number.Text = Number.Text - 1
If Number.Text < 1 Then
    Number.Text = 1
End If
Call LoadDat
End Sub
