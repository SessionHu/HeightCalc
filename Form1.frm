VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ȵ�λת����"
   ClientHeight    =   1560
   ClientLeft      =   3750
   ClientTop       =   2070
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "΢���ź�"
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
   StartUpPosition =   1  '����������
   Begin VB.ComboBox Combo2 
      Height          =   375
      ItemData        =   "Form1.frx":048A
      Left            =   1980
      List            =   "Form1.frx":04D9
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "�����λ"
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
      ToolTipText     =   "ԭ��λ"
      Top             =   120
      Width           =   1515
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���(&C)"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "�������"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�˳�(&E)"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "�˳�����"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "���"
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "ԭ����"
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
    '����Ϊ���Ƶ�λ
        If Combo1.Text = "����(�M)" Then
            Let m = Val(Text1.Text) * 100
        End If
        If Combo1.Text = "��(m)" Then
            Let m = Val(Text1.Text)
        End If
        If Combo1.Text = "����(�L)" Then
            Let m = Val(Text1.Text) * 1000
        End If
        If Combo1.Text = "����(dm)" Then
            Let m = Val(Text1.Text) * 10
        End If
        If Combo1.Text = "����(nm)" Then
            Let m = Val(Text1.Text) / (10 ^ -9)
        End If
        If Combo1.Text = "΢��(��m)" Then
            Let m = Val(Text1.Text) / (10 ^ -6)
        End If
        If Combo1.Text = "ǧ��(�N)" Then
            Let m = Val(Text1.Text) * 1000
        End If
    '���Ƶ�λ���
    '����Ϊ���Ƶ�λ
        If Combo1.Text = "���У���" Then
            Let m = Val(Text1.Text) * 500
        End If
        If Combo1.Text = "���У���" Then
            Let m = Val(Text1.Text) * (100 / 3)
        End If
        If Combo1.Text = "���У���" Then
            Let m = Val(Text1.Text) / 0.3
        End If
        If Combo1.Text = "���У���" Then
            Let m = Val(Text1.Text) / 3
        End If
        If Combo1.Text = "���У���" Then
            Let m = Val(Text1.Text) / 30
        End If
        If Combo1.Text = "���У���" Then
            Let m = Val(Text1.Text) / 300
        End If
        If Combo1.Text = "���У���" Then
            Let m = Val(Text1.Text) / 3000
        End If
    '����Ϊ���Ƶ�λ
    '����ΪӢ�Ƶ�λ
        If Combo1.Text = "Ӣ��(mi)" Then
            Let m = Val(Text1.Text) * 1690.9344
        End If
        If Combo1.Text = "��(yd)" Then
            Let m = Val(Text1.Text) * 0.9144
        End If
        If Combo1.Text = "Ӣ��(ft)" Then
            Let m = Val(Text1.Text) * 0.3048
        End If
        If Combo1.Text = "Ӣ��(in)" Then
            Let m = Val(Text1.Text) * 0.0254
        End If
    '����ΪӢ�Ƶ�λ
    '����Ϊ�����õ�λ
        If Combo1.Text = "����(n mile)" Then
            Let m = Val(Text1.Text) * 1852
        End If
        If Combo1.Text = "����(ly)" Then
            Let m = Val(Text1.Text) * 9.4607304725808E+15
        End If
        If Combo1.Text = "���ĵ�λ(A.U.)" Then
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
    '����Ϊ���Ƶ�λ
        If Combo1.Text = "����(�M)" Then
            Let m = Val(Text1.Text) * 100
        End If
        If Combo1.Text = "��(m)" Then
            Let m = Val(Text1.Text)
        End If
        If Combo1.Text = "����(�L)" Then
            Let m = Val(Text1.Text) * 1000
        End If
        If Combo1.Text = "����(dm)" Then
            Let m = Val(Text1.Text) * 10
        End If
        If Combo1.Text = "����(nm)" Then
            Let m = Val(Text1.Text) / (10 ^ -9)
        End If
        If Combo1.Text = "΢��(��m)" Then
            Let m = Val(Text1.Text) / (10 ^ -6)
        End If
        If Combo1.Text = "ǧ��(�N)" Then
            Let m = Val(Text1.Text) * 1000
        End If
    '���Ƶ�λ���
    '����Ϊ���Ƶ�λ
        If Combo1.Text = "���У���" Then
            Let m = Val(Text1.Text) * 500
        End If
        If Combo1.Text = "���У���" Then
            Let m = Val(Text1.Text) * (100 / 3)
        End If
        If Combo1.Text = "���У���" Then
            Let m = Val(Text1.Text) / 0.3
        End If
        If Combo1.Text = "���У���" Then
            Let m = Val(Text1.Text) / 3
        End If
        If Combo1.Text = "���У���" Then
            Let m = Val(Text1.Text) / 30
        End If
        If Combo1.Text = "���У���" Then
            Let m = Val(Text1.Text) / 300
        End If
        If Combo1.Text = "���У���" Then
            Let m = Val(Text1.Text) / 3000
        End If
    '����Ϊ���Ƶ�λ
    '����ΪӢ�Ƶ�λ
        If Combo1.Text = "Ӣ��(mi)" Then
            Let m = Val(Text1.Text) * 1690.9344
        End If
        If Combo1.Text = "��(yd)" Then
            Let m = Val(Text1.Text) * 0.9144
        End If
        If Combo1.Text = "Ӣ��(ft)" Then
            Let m = Val(Text1.Text) * 0.3048
        End If
        If Combo1.Text = "Ӣ��(in)" Then
            Let m = Val(Text1.Text) * 0.0254
        End If
    '����ΪӢ�Ƶ�λ
    '����Ϊ�����õ�λ
        If Combo1.Text = "����(n mile)" Then
            Let m = Val(Text1.Text) * 1852
        End If
        If Combo1.Text = "����(ly)" Then
            Let m = Val(Text1.Text) * 9.4607304725808E+15
        End If
        If Combo1.Text = "���ĵ�λ(A.U.)" Then
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
    End                                                                                     '��������
End Sub

Private Sub Command2_Click()
    Let Text1.Text = ""                                                                     '�����������
    Let Text2.Text = ""                                                                     '����������
End Sub

Private Sub Label1_DblClick()
    MsgBox "��Ȩ���� (C) 2022 XhuOffice  ��������Ȩ��", vbInformation, "XhuOffice"          '��Ȩ��Ϣ
End Sub
