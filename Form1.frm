VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��߼�����"
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
      Left            =   1680
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���(&C)"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�˳�(&E)"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
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
Private Sub Command1_Click()
 End                                                                                    '��������
End Sub
Private Sub Command2_Click()
 Let Text1.Text = ""                                                                    '�����
 Let Text2.Text = ""                                                                    '�������
End Sub
Private Sub Label1_DblClick()
 MsgBox "��Ȩ���� (C) 2022 XhuOffice  ��������Ȩ��", vbInformation, "XhuOffice"         '��Ȩ��Ϣ
End Sub
