VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfig 
   Caption         =   "配置 - VB6 To C#"
   ClientHeight    =   3264
   ClientLeft      =   60
   ClientTop       =   408
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   3264
   ScaleWidth      =   6900
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdSave 
      Caption         =   "..."
      Height          =   288
      Left            =   6120
      TabIndex        =   10
      Top             =   1080
      Width           =   372
   End
   Begin VB.CommandButton cmdOpenproject 
      Caption         =   "..."
      Height          =   288
      Left            =   6120
      TabIndex        =   9
      Top             =   600
      Width           =   372
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Configuration:"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin MSComDlg.CommonDialog dia2 
         Left            =   960
         Top             =   2280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "VB工程|*.vbp"
      End
      Begin VB.TextBox txtAssemblyName 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   1440
         Width           =   4215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消"
         Height          =   495
         Left            =   3720
         TabIndex        =   7
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定"
         Default         =   -1  'True
         Height          =   495
         Left            =   5160
         TabIndex        =   8
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtOutput 
         Height          =   372
         Left            =   1680
         TabIndex        =   4
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtVBPFile 
         Height          =   492
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label lblAssemblyName 
         Alignment       =   1  'Right Justify
         Caption         =   "Assembly Name:"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1452
      End
      Begin VB.Label lblOutput 
         Alignment       =   1  'Right Justify
         Caption         =   "输出文件夹："
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1452
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSrc 
         Alignment       =   2  'Center
         Caption         =   "工程文件："
         Height          =   252
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1332
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Config form
Dim sh As New Shell

Private Sub cmdOpenproject_Click()
dia2.ShowOpen
txtVBPFile.Text = dia2.FileName
End Sub

Private Sub cmdSave_Click()
    Dim str
    str = GetFolder(Me.hWnd, "浏览文件夹")
    If str <> "" Then
        txtOutput.Text = str
    End If
End Sub

Private Sub Form_Load()
  modConfig.Hush = True
  With Me
    .txtVBPFile = modConfig.vbpFile
    .txtOutput = modConfig.OutputFolder
    .txtAssemblyName = modConfig.AssemblyName
  End With
  modConfig.Hush = False
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  modINI.INIWrite INISection_Settings, INIKey_VBPFile, txtVBPFile, INIFile
  modINI.INIWrite INISection_Settings, INIKey_OutputFolder, txtOutput, INIFile
  modINI.INIWrite INISection_Settings, INIKey_AssemblyName, txtAssemblyName, INIFile
  modConfig.LoadSettings True
  Unload Me
End Sub

Private Sub fraConfig_DblClick()
  If MsgBox("Reset to default?", vbOKCancel, "Config Reset") = vbCancel Then Exit Sub
  txtVBPFile = App.Path & "\prj.vbp"
  txtOutput = App.Path & "\quick"
  txtAssemblyName = "VB2CS"
End Sub

Private Sub txtOutput_Validate(ByRef Cancel As Boolean)
  If Dir(txtOutput, vbDirectory) = "" Then
    MsgBox "Output folder does not exist.  Please create to prevent errors."
  End If
End Sub

Private Sub txtVBPFile_Validate(ByRef Cancel As Boolean)
  If Dir(txtVBPFile) = "" Then
    MsgBox "Project file does not exist.  Please give a valid project to prevent errors."
  End If
End Sub

Private Sub txtAssemblyName_Validate(ByRef Cancel As Boolean)
  If txtAssemblyName = "" Then
    MsgBox "Please enter something for an assembly name."
  End If
End Sub

Private Sub txtVBPFile_GotFocus(): txtVBPFile.SelStart = 0: txtVBPFile.SelLength = Len(txtVBPFile): End Sub
Private Sub txtOutput_GotFocus(): txtOutput.SelStart = 0: txtOutput.SelLength = Len(txtOutput): End Sub
Private Sub txtAssemblyName_GotFocus(): txtAssemblyName.SelStart = 0: txtAssemblyName.SelLength = Len(txtAssemblyName): End Sub


