VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "长度单位转换器"
   ClientHeight    =   1545
   ClientLeft      =   3390
   ClientTop       =   2070
   ClientWidth     =   3030
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3030
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox Combo2 
      Height          =   375
      ItemData        =   "Form1.frx":048A
      Left            =   1680
      List            =   "Form1.frx":04A3
      Sorted          =   -1  'True
      TabIndex        =   6
      Text            =   "厘米(㎝)"
      ToolTipText     =   "结果单位"
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   375
      ItemData        =   "Form1.frx":04EB
      Left            =   1680
      List            =   "Form1.frx":0504
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "米(m)"
      ToolTipText     =   "原单位"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空(&C)"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "清空数字"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出(&E)"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "退出程序"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "结果"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "原数据"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "by Session"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
    '以下为公制单位
        If Combo1.Text = "厘米(㎝)" Then
            Let m = Val(Text1.Text) * 100
        End If
        If Combo1.Text = "米(m)" Then
            Let m = Val(Text1.Text)
        End If
        If Combo1.Text = "毫米(㎜)" Then
            Let m = Val(Text1.Text) * 1000
        End If
        If Combo1.Text = "分米(dm)" Then
            Let m = Val(Text1.Text) * 10
        End If
        If Combo1.Text = "纳米(nm)" Then
            Let m = Val(Text1.Text) / (10 ^ -9)
        End If
        If Combo1.Text = "微米(μm)" Then
            Let m = Val(Text1.Text) / (10 ^ -6)
        End If
        If Combo1.Text = "千米(㎞)" Then
            Let m = Val(Text1.Text) * 1000
        End If
    '公制单位完成
    '以下为市制单位（仅中国）
        'If Combo1.Text = "
  
  
  
    MsgBox m, , "m="                   'debug
End Sub
Private Sub Command1_Click()
    End                                                                                     '结束进程
End Sub

Private Sub Command2_Click()
    Let Text1.Text = ""                                                                     '清除输入文字
    Let Text2.Text = ""                                                                     '清除输出文字
End Sub

Private Sub Label1_DblClick()
    MsgBox "版权所有 (C) 2022 XhuOffice  保留所有权利", vbInformation, "XhuOffice"          '版权信息
End Sub
