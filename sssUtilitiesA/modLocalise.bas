Attribute VB_Name = "modLocalise"
Option Explicit
Option Base 1

Global g_cn As ADODB.Connection
Global g_blnLogMode As Boolean
Global g_rs As ADODB.Recordset
Global g_cmd As ADODB.Command

'Global g_dicStrings As Dictionary

Public Function Translate(ByRef strString As String) As String
    GoTo EXIT_PROC
    If strString = "" Then GoTo EXIT_PROC
    
    If g_cn Is Nothing Then
        OpenDatabase
    End If
    
'        If g_rs Is Nothing Then Set rs = New ADODB.Recordset
        Set g_cmd = New ADODB.Command
        Dim strSQL As String
        Dim param As ADODB.Parameter
        strSQL = "select value from strings where key = ? "
        g_cmd.ActiveConnection = g_cn
        Set param = g_cmd.CreateParameter("key", adVarChar, adParamInput, 1000, strString)
        g_cmd.Parameters.Append param
        g_cmd.CommandText = strSQL
        Set g_rs = Nothing
        Set g_rs = g_cmd.Execute
        If g_rs.EOF Then
            If g_blnLogMode Then
                Dim param2 As ADODB.Parameter
                strSQL = "insert into strings values(?,?)"
                g_cmd.CommandText = strSQL
                Set param2 = g_cmd.CreateParameter("value", adVarChar, adParamInput, 1000, strString)
                g_cmd.Parameters.Append param2
                g_cmd.Execute
            End If
            GoTo EXIT_PROC
        Else
            Translate = g_rs(0).Value
        End If
        Set g_cmd = Nothing
    
    Exit Function
EXIT_PROC:
    
   Translate = strString
End Function
Public Function FormatCurr(ByRef curCurrency As Currency) As Currency
'    If Not g_dicStrings Is Nothing Then
'    End If
    FormatCurr = Format(curCurrency, "##########.$$")
End Function
Public Function FormatDate(ByRef dt As Date) As Date
'    If Not g_dicStrings Is Nothing Then
'    End If
    FormatDate = dt
End Function
Public Sub TranslateForm(frm As Form)
    Dim idx As Integer
    Dim kdx As Integer
    Dim ctrl As Control
    Dim strTmp As String
    
    frm.Caption = Translate(frm.Caption)
    
    For idx = 0 To frm.Controls.Count - 1
        Set ctrl = frm.Controls.Item(idx)
        strTmp = TypeName(ctrl)
        
        
        
        Select Case UCase(TypeName(ctrl))
        Case "COMMANDBUTTON"
            ctrl.Caption = Translate(ctrl.Caption)
        Case "FRAME", "LABEL", "CHECKBOX", "MENU"
            ctrl.Caption = Translate(ctrl.Caption)
        Case "TEXTBOX"
            ctrl.Text = Translate(ctrl.Text)
        Case "STATUSBAR"
            For kdx = 1 To ctrl.Panels.Count
                ctrl.Panels(kdx).Text = Translate(ctrl.Panels(kdx).Text)
            Next
        Case "LISTVIEW"
            For kdx = 1 To ctrl.ColumnHeaders.Count
                ctrl.ColumnHeaders(kdx).Text = Translate(ctrl.ColumnHeaders(kdx).Text)
            Next
        Case "SSTAB"
            For kdx = 0 To ctrl.Tabs - 1
                'ctrl.Tab = kdx
                ctrl.TabCaption(kdx) = Translate(ctrl.TabCaption(kdx))
            Next
            ctrl.Tab = 0
        Case "DTPICKER", "LISTBOX", "IMAGELIST"
        Case Else
'            MsgBox strTmp
        End Select
    
    Next
End Sub
Private Sub OpenDatabase()
    Dim strDatabase As String
    
    Set g_cn = New ADODB.Connection
    g_cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    strDatabase = App.Path & "\sssLocalise.mdb"
    g_cn.Open (strDatabase)
    g_blnLogMode = False
End Sub
