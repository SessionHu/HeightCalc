VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "长度单位转换器"
   ClientHeight    =   1560
   ClientLeft      =   3750
   ClientTop       =   2070
   ClientWidth     =   3615
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
   ScaleHeight     =   1560
   ScaleWidth      =   3615
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox Combo2 
      Height          =   375
      ItemData        =   "Form1.frx":048A
      Left            =   1980
      List            =   "Form1.frx":04D9
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "结果单位"
      Top             =   600
      Width           =   1515
   End
   Begin VB.ComboBox Combo1 
      Height          =   375
      ItemData        =   "Form1.frx":05E7
      Left            =   1980
      List            =   "Form1.frx":0636
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "原单位"
      Top             =   120
      Width           =   1515
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空(&C)"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "清空数字"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出(&E)"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "退出程序"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "结果"
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "原数据"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "by Session"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1150
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
    '以下为市制单位
        If Combo1.Text = "（市）里" Then
            Let m = Val(Text1.Text) * 500
        End If
        If Combo1.Text = "（市）引" Then
            Let m = Val(Text1.Text) * (100 / 3)
        End If
        If Combo1.Text = "（市）丈" Then
            Let m = Val(Text1.Text) / 0.3
        End If
        If Combo1.Text = "（市）尺" Then
            Let m = Val(Text1.Text) / 3
        End If
        If Combo1.Text = "（市）寸" Then
            Let m = Val(Text1.Text) / 30
        End If
        If Combo1.Text = "（市）分" Then
            Let m = Val(Text1.Text) / 300
        End If
        If Combo1.Text = "（市）厘" Then
            Let m = Val(Text1.Text) / 3000
        End If
    '以上为市制单位
    '以下为英制单位
        If Combo1.Text = "英里(mi)" Then
            Let m = Val(Text1.Text) * 1690.9344
        End If
        If Combo1.Text = "码(yd)" Then
            Let m = Val(Text1.Text) * 0.9144
        End If
        If Combo1.Text = "英尺(ft)" Then
            Let m = Val(Text1.Text) * 0.3048
        End If
        If Combo1.Text = "英寸(in)" Then
            Let m = Val(Text1.Text) * 0.0254
        End If
    '以上为英制单位
    '以下为不常用单位
        If Combo1.Text = "海里(n mile)" Then
            Let m = Val(Text1.Text) * 1852
        End If
        If Combo1.Text = "光年(ly)" Then
            Let m = Val(Text1.Text) * 9.4607304725808E+15
        End If
        If Combo1.Text = "天文单位(A.U.)" Then
            Let m = Val(Text1.Text) * 149597870700#
        End If
    Form1.Cls                   'debug
    Print ""                    'debug
    Print ""                    'debug
    Print ""                    'debug
    Print ""                    'debug
    Print ""                    'debug
    Print "DEBUG: m:"; m        'debug
End Sub

Private Sub Text1_Change()
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
    '以下为市制单位
        If Combo1.Text = "（市）里" Then
            Let m = Val(Text1.Text) * 500
        End If
        If Combo1.Text = "（市）引" Then
            Let m = Val(Text1.Text) * (100 / 3)
        End If
        If Combo1.Text = "（市）丈" Then
            Let m = Val(Text1.Text) / 0.3
        End If
        If Combo1.Text = "（市）尺" Then
            Let m = Val(Text1.Text) / 3
        End If
        If Combo1.Text = "（市）寸" Then
            Let m = Val(Text1.Text) / 30
        End If
        If Combo1.Text = "（市）分" Then
            Let m = Val(Text1.Text) / 300
        End If
        If Combo1.Text = "（市）厘" Then
            Let m = Val(Text1.Text) / 3000
        End If
    '以上为市制单位
    '以下为英制单位
        If Combo1.Text = "英里(mi)" Then
            Let m = Val(Text1.Text) * 1690.9344
        End If
        If Combo1.Text = "码(yd)" Then
            Let m = Val(Text1.Text) * 0.9144
        End If
        If Combo1.Text = "英尺(ft)" Then
            Let m = Val(Text1.Text) * 0.3048
        End If
        If Combo1.Text = "英寸(in)" Then
            Let m = Val(Text1.Text) * 0.0254
        End If
    '以上为英制单位
    '以下为不常用单位
        If Combo1.Text = "海里(n mile)" Then
            Let m = Val(Text1.Text) * 1852
        End If
        If Combo1.Text = "光年(ly)" Then
            Let m = Val(Text1.Text) * 9.4607304725808E+15
        End If
        If Combo1.Text = "天文单位(A.U.)" Then
            Let m = Val(Text1.Text) * 149597870700#
        End If
    Form1.Cls                   'debug
    Print ""                    'debug
    Print ""                    'debug
    Print ""                    'debug
    Print ""                    'debug
    Print ""                    'debug
    Print "DEBUG: m:"; m        'debug
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
