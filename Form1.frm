VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Pantau Permukaan Air Kanal"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5415
      Begin VB.Label lblKOPERASIKARYAWAN 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PANTAU PERMUKAAN AIR KANAL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   315
         TabIndex        =   7
         Top             =   240
         Width           =   4785
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   3000
      Width           =   5415
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H008080FF&
         Caption         =   "&Keluar"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H0000FF00&
      Caption         =   "&Masuk"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   3180
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   720
      Width           =   5400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MenuAkses As String
Public LoginSucceeded As Boolean

Private Sub CmdCancel_Click()
End
End Sub

Private Sub Command1_Click()
    'On Error GoTo ErrorHandler
Dim rsPassWord As ADODB.Recordset
Dim sql As String
Me.MousePointer = vbHourglass
MdiPKB.MousePointer = vbHourglass

If Trim(txtUserName.Text) = "" Or Trim(txtPassword.Text) = "" Then
   MsgBox "UserName atau Password harus diisi"
   txtUserName.SetFocus
   SendKeys "{Home}+{End}"
   Me.MousePointer = vbNormal
   MdiPKB.MousePointer = vbDefault
   Exit Sub
End If

Set rsPassWord = New ADODB.Recordset
sql = "select * from tblUtlUser where UserID='" & Trim(txtUserName.Text) & "'"
rsPassWord.Open sql, ActiveCn, adOpenKeyset, adLockOptimistic

If rsPassWord.RecordCount = 0 Then
   MsgBox "User Anda dengan User Name " & Trim(txtUserName.Text) & " belum terdaftar, silahkan cek ", vbInformation, At
   txtUserName.SetFocus
   SendKeys "{Home}+{end}"
   Me.MousePointer = vbNormal
   MdiPKB.MousePointer = vbDefault
   Exit Sub
Else
   
   'check for correct password
   'EncryptText(PW, Trim(Text1.text))
   'If UCase(Trim(txtPassword.Text)) = DecryptText(rsPassWord!Password, Trim(txtUserName.Text)) Then 'rsPassWord!Password Then
   If Trim(txtPassword.Text) = DecryptText(rsPassWord!Password, Trim(txtUserName.Text)) Then 'rsPassWord!Password Then
      MdiPKB.UserID = rsPassWord!UserID
      MdiPKB.NamaUser = rsPassWord!Nama
      MdiPKB.Password = DecryptText(rsPassWord!Password, Trim(txtUserName.Text)) 'rsPassWord!Password
      MdiPKB.Jabatan = rsPassWord!Jabatan
      MdiPKB.Wilayah = rsPassWord!KodeWilayah
      MdiPKB.GroupUser = rsPassWord!NamaGroup
      MdiPKB.Periode = Date
      'mdipkb.StatusBar1.Panels("Periode").text = "Periode : " & Format(Date, "MMMM yyyy")
      'mdipkb.StatusBar1.Panels(1).text = rsPassWord!Nama
      'mdipkb.StatusBar1.Panels(2).text = rsPassWord!Bagian
      'mdipkb.StatusBar1.Panels(3).text = "Periode : " & Format(Date, "MMMM yyyy")
      MenuAkses = rsPassWord!NamaGroup
'      EnabledMenu
      'Masuk ke tabel LogOut
      'mainform.cnPKB.Execute "insert into tlogout (Nik,Login) values ('" & rsPassWord!NIK & "',  '" & CDate(Now) & "'  ) "
      'place code to here to pass the
      'success to the calling sub
      'setting a global var is the easiest
      LoginSucceeded = True
      
      If rsPassWord!NeedUpdatedPwd = True Then
            MsgBox "PERHATIAN  !!!!" & vbCrLf & _
                "Password Login Lemah, Silahkan untuk mengganti password anda setelah Login. Terima Kasih", vbInformation, At
      End If
      Me.Hide
    Else
        MsgBox "Password anda salah, coba lagi!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End If

Me.MousePointer = vbDefault
MdiPKB.MousePointer = vbDefault

'ErrorHandler:
'If Err.Number <> 0 Then
'    MsgBox Err.Number & " : " & Err.Description
'    Me.MousePointer = vbDefault
'End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
'Me.Left = (MdiPKB.Width - Me.Width) / 2
'Me.Top = (MdiPKB.Height - Me.Height) / 2
End Sub
