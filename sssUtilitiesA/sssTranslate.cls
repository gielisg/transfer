VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sssTranslate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private g_cn As ADODB.Connection
Private g_blnLogMode As Boolean
Private g_rs As ADODB.Recordset
Private g_cmd As ADODB.Command

Private m_strPath As String

Public Property Let Path(strPath As String)
    m_strPath = strPath
End Property

Public Function Translate(ByRef strString As String) As String
    If strString = "" Then GoTo EXIT_PROC
    Translate = strString
    GoTo EXIT_PROC
    
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
Private Sub OpenDatabase()
    Dim strDatabase As String
    
    Set g_cn = New ADODB.Connection
    g_cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    strDatabase = m_strPath & "\sssLocalise.mdb"
    g_cn.Open (strDatabase)
    g_blnLogMode = False
End Sub

