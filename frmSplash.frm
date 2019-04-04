VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Windows 平台"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5160
         TabIndex        =   6
         Top             =   2040
         Width           =   1590
      End
      Begin VB.Label Label2 
         Caption         =   "http://www.fidelxyz.icoc.bz"
         Height          =   255
         Left            =   4440
         TabIndex        =   9
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "华英专属定制版"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   8
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   2415
      End
      Begin VB.Label lblCopyright 
         Caption         =   "版权所有 @2017"
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "Powered by Fidel"
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Caption         =   "警告 版权所有，盗版必究。"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Width           =   6765
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "版本"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6240
         TabIndex        =   5
         Top             =   2400
         Width           =   510
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "拍卖软件"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   32.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2640
         TabIndex        =   7
         Top             =   960
         Width           =   2640
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "Fidel 授权"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub imgLogo_Click()
    Unload Me
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub lblCompany_Click()
    Unload Me
End Sub

Private Sub lblCopyright_Click()
    Unload Me
End Sub

Private Sub lblLicenseTo_Click()
    Unload Me
End Sub

Private Sub lblPlatform_Click()
    Unload Me
End Sub

Private Sub lblProductName_Click()
    Unload Me
End Sub

Private Sub lblVersion_Click()
    Unload Me
End Sub

Private Sub lblWarning_Click()
    Unload Me
End Sub
