VERSION 5.00
Begin VB.Form WinActFormMain 
   Caption         =   "Windows Activer"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   5340
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox WinActServerList 
      Height          =   300
      ItemData        =   "WinActFormMain.frx":0000
      Left            =   360
      List            =   "WinActFormMain.frx":0007
      TabIndex        =   6
      Text            =   "（请选择预设KMS服务器或输入现有服务器地址）"
      Top             =   3480
      Width           =   4575
   End
   Begin VB.ComboBox WinActKeyList 
      Height          =   300
      ItemData        =   "WinActFormMain.frx":001A
      Left            =   360
      List            =   "WinActFormMain.frx":0021
      TabIndex        =   5
      Text            =   "（请选择预设产品密钥或输入现有产品密钥）"
      Top             =   2280
      Width           =   4575
   End
   Begin VB.CommandButton WinActUtils 
      Caption         =   "其他功能"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton WinActExec 
      Caption         =   "点击激活"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label WinActServerTips 
      Caption         =   "请在此处选择KMS服务器："
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   3120
      Width           =   4575
   End
   Begin VB.Label WinActKeyTips 
      Caption         =   "请在此处选择产品密钥："
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Label WinActLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Windows Activer"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "WinActFormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Activator()
Set objShell = CreateObject("WScript.Shell")
objShell.Run "slmgr.vbs /upk", vbWaitOnReturn
objShell.Run ("slmgr.vbs /ipk " + WinActKeyList.Text), vbWaitOnReturn
objShell.Run ("slmgr.vbs /skms " + WinActServerList.Text), vbWaitOnReturn
objShell.Run "slmgr.vbs /ato", vbWaitOnReturn
'Retry Test
FailMsg = MsgBox("该Windows副本激活失败，原因未知。", vbCritical + vbRetryCancel, "激活失败")
If FailMsg = vbRetry Then
Activator
Else
WinActExec.Enabled = True
WinActExec.Caption = "点击激活"
End If
End Sub

Sub RunAsAdmin()
    ' 定义要以管理员身份运行的exe文件的完整路径
    Dim filePath As String
    
    ' 使用Shell函数以管理员身份运行程序，并提示用户确认权限
    Shell "runas /user:Administrator """"" & "%1" & """", vbNormalFocus
End Sub
Private Sub Form_Load()
'RunAsAdmin
End Sub

Private Sub WinActExec_Click()
If WinActKeyList.Text = "（请选择预设产品密钥或输入现有产品密钥）" Then
MsgBox ("你尚未输入Windows产品密钥，无法激活。"), 48, ("激活失败")
ElseIf WinActServerList.Text = "（请选择预设KMS服务器或输入现有服务器地址）" Then
MsgBox ("你尚未设置KMS服务器，无法激活。"), 48, ("激活失败")
Else
ExecMsg = MsgBox("确认要激活Windows副本吗？" & vbCrLf & "你选择的产品密钥是：" + WinActKeyList.Text & vbCrLf & "你选择的KMS服务器是：" + WinActServerList.Text, vbQuestion + vbYesNo, "二次确认")
End If
If ExecMsg = vbYes Then
WinActExec.Enabled = False
WinActExec.Caption = "正在激活"
Activator
End If
WinActExec.Enabled = True
WinActExec.Caption = "点击激活"
End Sub

'For debugging only describe it to disable
Private Sub WinActUtils_Click()
WinActExec.Enabled = True
WinActExec.Caption = "点击激活"
MsgBox ("恭喜，已成功激活该Windows副本。"), 64, ("激活完毕")
End Sub

