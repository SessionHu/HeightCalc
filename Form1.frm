VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ȵ�λת����"
   ClientHeight    =   1545
   ClientLeft      =   3390
   ClientTop       =   2070
   ClientWidth     =   3030
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
   ScaleHeight     =   1545
   ScaleWidth      =   3030
   StartUpPosition =   1  '����������
   Begin VB.ComboBox Combo2 
      Height          =   375
      ItemData        =   "Form1.frx":048A
      Left            =   1680
      List            =   "Form1.frx":04A3
      Sorted          =   -1  'True
      TabIndex        =   6
      Text            =   "����(�M)"
      ToolTipText     =   "�����λ"
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
      Text            =   "��(m)"
      ToolTipText     =   "ԭ��λ"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���(&C)"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "�������"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�˳�(&E)"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "�˳�����"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "���"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "ԭ����"
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
    '����Ϊ���Ƶ�λ�����й���
        'If Combo1.Text = "
  
  
  
    MsgBox m, , "m="                   'debug
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
