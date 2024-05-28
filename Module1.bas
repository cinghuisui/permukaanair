Attribute VB_Name = "Module1"

Option Explicit
Public rs As ADODB.Recordset
Public cn As ADODB.Connection
Public Const At As String = "Aplikasi Pantau Permukaan Air Kanal"

Public Const AppServer As String = "192.168.9.9" '"192.168.13.67"
Public Const AppDB     As String = "MyKTTATest"
Public Const AppUser   As String = "pkbKoneksi"
Public Const AppPass   As String = "s@mbu56#2024" '"12345"
Function ActiveCn() As String
    ActiveCn = "Provider=SQLOLEDB.1; " & _
               "Persist Security Info=False; " & _
               "User ID=" & AppUser & "; " & _
               "Password=" & AppPass & "; " & _
               "Initial Catalog=" & AppDB & "; " & _
               "Data Source=" & AppServer
End Function
Public Sub HapusGrid(ocx As VSFlexGrid, baris As Integer)
    Dim i As Long
    ocx.Rows = baris + 1
    
    For i = 1 To ocx.Cols - 1
        ocx.TextMatrix(baris, i) = ""
    Next i
End Sub
Function GetDate() As String
    Dim rs As New ADODB.Recordset
    On Error GoTo EH
    Call rs.Open("SELECT CONVERT(varchar, GETDATE(), 103) + ' ' + CONVERT(varchar, GETDATE(), 108) AS Now", ActiveCn)
    rs.MoveFirst
    GetDate = rs!Now
    Set rs = Nothing
    Exit Function
EH:
    Set rs = Nothing
    Call ErrMSg(Err)
End Function
Function sqllogin() As String
    Dim rs As New ADODB.Recordset
    On Error GoTo EH
    Call rs.Open("Select LEFT(SUSER_NAME(),25)As login", ActiveCn)
    rs.MoveFirst
    sqllogin = rs!login
    Set rs = Nothing
    Exit Function
EH:
    Set rs = Nothing
    Call ErrMSg(Err)
End Function
Function namacom() As String
    Dim rs As New ADODB.Recordset
    On Error GoTo EH
    Call rs.Open("Select LEFT(HOST_NAME(),15) as namacom", ActiveCn)
    rs.MoveFirst
    namacom = rs!namacom
    Set rs = Nothing
    Exit Function
EH:
    Set rs = Nothing
    Call ErrMSg(Err)
End Function
Sub ErrMSg(Error As Object)
    Dim Msg As String
    Select Case Error.Number
        Case 0
        Case -2147467259: Msg = "Connection to database server is broken !"
        Case 13: Msg = "Numeric value is not valid !"
        Case 3021: Msg = "No data !"
        Case Else: Msg = "Error : " & Error.Number & " : " & Error.Description
    End Select
    If Error.Number <> 0 Then Call MsgBox(Msg, vbExclamation, At)
End Sub
Public Sub ControlCentreForm(frm As Form, ctrl As Control)
    ctrl.Left = ((frm.Width - ctrl.Width) / 2) - 100
    ctrl.Top = ((frm.Height - ctrl.Height) / 2) - 100
End Sub
Public Sub LoadCentreForm(frm As Form)
    frm.Left = ((MdiPKB.Width - frm.Width) / 2)
    frm.Top = (MdiPKB.Height - frm.Height) / 8
End Sub
Public Sub FormSize(Tinggi As Double, Lebar As Double, frm As Form)
    frm.Height = Tinggi
    frm.Width = Lebar
End Sub
Public Function DecryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer, strBuff As String, d As Integer
    #If Not CASE_SENSITIVE_PASSWORD Then
        strPwd = UCase$(strPwd)
    #End If
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    DecryptText = strBuff
End Function
Public Function EncryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String
    
    #If Not CASE_SENSITIVE_PASSWORD Then
        strPwd = UCase$(strPwd)
    #End If
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    EncryptText = strText
End Function
'Function LoadPeminjaman(cbo As ComboBoxLB)
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Dim sql As String
'    Dim i As Long
'
'    sql = "Select DetailID HeaderID "
'    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly
'
'    If rs.RecordCount > 0 Then
'        cbo.ColumnCount = 2
'        cbo.ColumnWidths = "2500;0"
'        rs.MoveFirst
'
'        For i = 0 To rs.RecordCount - 1
'            cbo.AddItem rs!Peminjam & "'" & rs!DetailID
'            rs.MoveNext
'        Next i
'End Function
Public Sub HLText(ByRef srcText As TextBox)
    srcText.BackColor = &HC0FFFF
End Sub
'Function LoadMess(cbo As ComboBox)
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Dim sql As String
'    Dim i As Long
'
'    sql = "Select id, NamaMess From tblmstMess iif(MdiMess.GroupUser = "", MDIMess.id, "") & " ' Order by NamaMess Asc"
'    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly
'
'    If rs.RecordCount > 0 Then
'        cbo.ColumnCount = 2
'        cbo.ColumnWidths = "2500;0"
'        rs.MoveFirst
'
'        For i = 0 To rs.RecordCount - 1
'            cbo.AddItem rs!ID & ";" & rs!NamaMess
'            rs.MoveNext
'        Next i
'    End If
'End Function


