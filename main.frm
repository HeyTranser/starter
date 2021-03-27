VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form main 
   Caption         =   "主页"
   ClientHeight    =   3690
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   5025
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   5025
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   3480
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start!"
      Height          =   615
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog Cg1 
      Left            =   4440
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   8438015
      DefaultExt      =   "*.htp"
      Filter          =   "HT Project Files(*.htp)|*.htp"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "让我们从头开始吧"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   3480
   End
   Begin VB.Menu about 
      Caption         =   "关于"
      Begin VB.Menu a1 
         Caption         =   "关于项目..."
      End
      Begin VB.Menu a2 
         Caption         =   "项目网站"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pd As String



Dim pf As String




Private Sub a1_Click()
'MsgBox ("About Us" + vbCrLf + "HeyTranser Project" + vbCrLf + "maicarons@outlook.com")
frmAbout.Show
End Sub

Private Sub a2_Click()
Shell ("explorer https://github.com/HeyTranser/starter")
End Sub

Private Sub Command1_Click()
Me.Enabled = False

On Error GoTo connerror
  
Cg1.Action = 2
'action控件一共有6个值：1为打开文件showopen，2为保存文件showsave,3为选择颜色showcolor,4为选择字体showfont，5为打印showprint,6不常用
  'CommonDialog1.ShowSave
  'Open Cg1.FileName For Output As #1
  'Print #1,
  'Close 1
pf = Cg1.FileName
pd =

connerror:

End Sub
