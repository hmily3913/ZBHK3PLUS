VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SelectFileForm 
   Caption         =   "EXCEL����"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   5805
   StartUpPosition =   3  '����ȱʡ
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "����ѡ���ļ�"
      Top             =   480
      Width           =   4815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "excel�ļ�|*.xls"
   End
End
Attribute VB_Name = "SelectFileForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

        On Error GoTo ms:
        With Me.CommonDialog1
            .ShowOpen
        End With
        
        Dim strFileName  As String
        Text1 = Me.CommonDialog1.FileName
        Exit Sub
ms:
      MsgBox Err.Description, vbCritical, "�����ʾ"
End Sub


Private Sub Text1_Click()
    Me.CommonDialog1.ShowOpen
End Sub
