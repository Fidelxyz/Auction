VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�������"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14115
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   14115
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton OpenAdmin2 
      Caption         =   "���ݿ����"
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
      Left            =   10440
      TabIndex        =   25
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Exit 
      Caption         =   "EXIT"
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
      Left            =   10440
      MaskColor       =   &H8000000C&
      TabIndex        =   21
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton OpenAdmin 
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8520
      MaskColor       =   &H8000000C&
      TabIndex        =   20
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label State 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   26
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label16 
      Caption         =   "www.fidelxzyz.icoc.bz"
      Height          =   255
      Left            =   12000
      TabIndex        =   24
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Label Label15 
      Caption         =   "��Ӣר�����ư� - 2017��4�����Ʒ� ʹ��Visual Basic��д"
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
      Left            =   2640
      TabIndex        =   23
      Top             =   240
      Width           =   8895
   End
   Begin VB.Label Copyright 
      Caption         =   "Copyright @2017 �������          Powered By Fidel Version 1.0.0 "
      Height          =   735
      Left            =   12000
      TabIndex        =   22
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label BuyName 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   19
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   18
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Ԫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label MinMoney 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   16
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Money 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   13
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "ʵʱ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   5400
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   600
      Width           =   7200
   End
   Begin VB.Label Label12 
      Caption         =   "Ԫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label FirstMoney 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "���ļ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Passage 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   960
      TabIndex        =   8
      Top             =   2040
      Width           =   5655
   End
   Begin VB.Label Label8 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label PeopleName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Names 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label Label4 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Number 
      BackColor       =   &H00C0C0C0&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Title 
      Caption         =   "������� v1.0.0"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Ԫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   7200
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      X1              =   360
      X2              =   14040
      Y1              =   6240
      Y2              =   6240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private FormOldWidth As Long
'���洰���ԭʼ���
Private FormOldHeight As Long
'���洰���ԭʼ�߶�

Public Sub ResizeInit(FormName As Form)
    Dim Obj As Control
    FormOldWidth = Form1.ScaleWidth
    FormOldHeight = Form1.ScaleHeight
    On Error Resume Next
    For Each Obj In FormName
        Obj.Tag = Obj.Left & " " & Obj.Top & " " & Obj.Width & " " & Obj.Height & " "
    Next Obj
    On Error GoTo 0
End Sub

Public Sub ResizeForm(FormName As Form)
    Dim Pos(4) As Double
    Dim i      As Long, TempPos As Long, StartPos As Long
    Dim Obj    As Control
    Dim ScaleX As Double, ScaleY As Double

    ScaleX = FormName.ScaleWidth / FormOldWidth
    ScaleY = FormName.ScaleHeight / FormOldHeight
    On Error Resume Next
    For Each Obj In FormName
        StartPos = 1
        For i = 0 To 4
            TempPos = InStr(StartPos, Obj.Tag, " ", vbTextCompare)
            If TempPos > 0 Then
                Pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
                StartPos = TempPos + 1
            Else
                Pos(i) = 0
            End If

            Obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
        Next i
    Next Obj
    On Error GoTo 0
End Sub

Private Sub Form_Resize()
    Call ResizeForm(Me) 'ȷ������ı�ʱ�ؼ���֮�ı�
End Sub

Private Sub Form_Load()
    Call ResizeInit(Me) '�ڳ���װ��ʱ����
    '-------------------------------------
    Copyright.Caption = "Copyright @2017" & vbLf & "�������" & vbLf & "Powered By Fidel" & vbLf & "Version " & App.Major & "." & App.Minor & "." & App.Revision
Title.Caption = "������� v" & App.Major & "." & App.Minor & "." & App.Revision
'�ж�Image�ļ����Ƿ����
If Dir(App.Path + "\Images", vbDirectory) = "" Then
    MkDir (App.Path + "\Images")
End If
'�ж�data.xls�Ƿ����
If Dir(App.Path + "\data.xlsx") = "" Then
    '����data.xls
    Dim VBExcel As Excel.Application
    Set VBExcel = CreateObject("Excel.Application")
    With VBExcel
        .Workbooks.Add
            With ActiveWorkbook
                .SaveAs App.Path + "\data.xlsx"
                .Close
            End With
        .Quit
    End With
    '����Excelģ��
    Set objExcelFile = New Excel.Application
    objExcelFile.DisplayAlerts = False
    Set objWorkBook = objExcelFile.Workbooks.Open(App.Path + "\data.xlsx")
    Set objImportSheet = objWorkBook.Sheets(1)
    'data.xlsx��ʼ��
    objImportSheet.Cells(1, 1) = "No."
    objImportSheet.Cells(1, 2) = "����"
    objImportSheet.Cells(1, 3) = "������"
    objImportSheet.Cells(1, 4) = "���ļ�"
    objImportSheet.Cells(1, 5) = "�������"
    objImportSheet.Cells(1, 6) = "���ܣ�200�����ڣ�"
    objImportSheet.Cells(1, 8) = "������"
    objImportSheet.Cells(1, 9) = "�ɽ���"
    objWorkBook.SaveAs App.Path + "\data.xlsx"
    '����Excelģ��
    objExcelFile.Quit
    Set objWorkBook = Nothing
    Set objImportSheet = Nothing
    Set objExcelFile = Nothing
End If
End Sub

Private Function CheckExeIsRun(exeName As String) As Boolean
    On Error GoTo Err
    Dim WMI
    Dim Obj
    Dim Objs
    CheckExeIsRun = False
    Set WMI = GetObject("WinMgmts:")
    Set Objs = WMI.InstancesOf("Win32_Process")
    For Each Obj In Objs
      If (InStr(UCase(exeName), UCase(Obj.Description)) <> 0) Then
            CheckExeIsRun = True
            If Not Objs Is Nothing Then Set Objs = Nothing
            If Not WMI Is Nothing Then Set WMI = Nothing
            Exit Function
      End If
    Next
    If Not Objs Is Nothing Then Set Objs = Nothing
    If Not WMI Is Nothing Then Set WMI = Nothing
    Exit Function
Err:
    If Not Objs Is Nothing Then Set Objs = Nothing
    If Not WMI Is Nothing Then Set WMI = Nothing
End Function

Private Sub Copyright_Click()
frmSplash.Show
End Sub

Private Sub Exit_Click()
If AdminForm.FormAdmin = True Or Admin2.FormAdmin2 = True Then
    MsgBox "���ݿ�δ�رգ�"
Else
    End
End If
End Sub

Private Sub Label14_Click()
frmSplash.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
If AdminForm.FormAdmin = True Or Admin2.FormAdmin2 = True Then
    MsgBox "���ݿ�δ�رգ�"
    Cancel = True
End If
End Sub

Private Sub Label15_Click()
frmSplash.Show
End Sub

Private Sub Label16_Click()
frmSplash.Show
End Sub

Private Sub OpenAdmin_Click()
If CheckExeIsRun("EXCEL.EXE") Then
  MsgBox "����ͬʱ������������棡" & Chr(10) & "����һ�������������������"
Else
  AdminForm.Show
End If
End Sub

Private Sub OpenAdmin2_Click()
If CheckExeIsRun("EXCEL.EXE") Then
  MsgBox "����ͬʱ������������棡" & Chr(10) & "����һ�������������������"
Else
  Admin2.Show
End If
End Sub

Private Sub Title_Click()
frmSplash.Show
End Sub
