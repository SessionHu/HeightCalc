VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��߼�����"
   ClientHeight    =   1545
   ClientLeft      =   3390
   ClientTop       =   2040
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command4 
      Caption         =   "���(&C)"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�˳�(&E)"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "������"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���ס�"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "by Session"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "���ף�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "���ף�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Let Text2.Text = Val(Text1.Text) * 100                                                 '��������Ϊ�׵�100��
End Sub
Private Sub Command2_Click()
 Let Text1.Text = Val(Text2.Text) * 0.01                                                '������Ϊ���׵�0.01��
End Sub
Private Sub Command3_Click()
 End                                                                                    '��������
End Sub
Private Sub Command4_Click()
 Let Text1.Text = ""                                                                    '�����
 Let Text2.Text = ""                                                                    '�������
End Sub
Private Sub Label3_DblClick()
 MsgBox "��Ȩ���� (C) 2022 XhuOffice  ��������Ȩ��", vbInformation, "XhuOffice"         '��Ȩ��Ϣ
End Sub
