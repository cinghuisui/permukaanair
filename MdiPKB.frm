VERSION 5.00
Begin VB.MDIForm MdiPKB 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuLevelAir 
      Caption         =   "&Menu"
      Index           =   3000
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
      Index           =   2000
   End
End
Attribute VB_Name = "MdiPKB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
    Public cn As ADODB.Connection
    Public UserID, NamaUser, Password, Jabatan, Wilayah, GroupUser As String
    Public Periode As Date
    Public no As Integer
    Dim mDeactivated As Boolean
Private Sub MDIForm_Activate()
    If mDeactivated = False Then
        Me.MousePointer = vbHourglass
        Form1.Show vbModal
        Me.MousePointer = vbNormal
    End If
End Sub
Private Sub MDIForm_Deactivate()
mDeactivated = True
End Sub
Private Sub MDIForm_Initialize()
mDeactivated = False
End Sub

Private Sub mnuExit_Click(Index As Integer)
End
End Sub

Private Sub mnuLevelAir_Click(Index As Integer)
    FrmLevelAir.Show
End Sub
