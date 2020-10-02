VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sssMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

'***    Message variables
Private m_sMessageId As String
Private m_sCategory As String
Private m_dMessageDate As Date
Private m_iPriority As Integer
Private m_lSessionId As Long
Private m_sMessage As String
Private m_bMessage As Boolean
Private m_bValidMessage As Boolean

'***    user variables
Private m_strLogonCode As String
Private m_strFirstName As String
Private m_strName As String
Private m_strContactCode As String
Private m_strEmail As String
Private m_strPassword As String
Private m_strDSN As String

'***    these may not be needed rely on DSN
'Private m_strDatabaseName As String
'Private m_strDBInstance As String
'Private m_strDBHost As String
'Private m_strConnectMode As String

Private m_strCurrentBusinessUnitCode As String
Private m_strCurrentBusinessUnit As String
Public Property Get ValidMessage() As Boolean
    ValidMessage = m_bValidMessage
End Property
Public Property Get AnSssMessage() As Boolean
    AnSssMessage = m_bMessage
End Property
Public Property Let MessageId(vData As String)
    m_sMessageId = vData
End Property
Public Property Get MessageId() As String
    MessageId = m_sMessageId
End Property

Public Property Let Category(vData As String)
    m_sCategory = vData
End Property
Public Property Get Category() As String
    Category = m_sCategory
End Property

Public Property Let MessageDate(vData As Date)
    m_dMessageDate = vData
End Property
Public Property Get MessageDate() As Date
    MessageDate = m_dMessageDate
End Property
Public Property Let Priority(vData As Integer)
    m_iPriority = vData
End Property
Public Property Get Priority() As Integer
    Priority = m_iPriority
End Property
Public Property Let SessionId(vData As Long)
    m_lSessionId = vData
End Property
Public Property Get SessionId() As Long
    SessionId = m_lSessionId
End Property
Public Property Let Message(vData As String)
    m_sMessage = vData
End Property
Public Property Get Message() As String
    Message = m_sMessage
End Property
Public Property Let LogonCode(vData As String)
    m_strLogonCode = vData
End Property
Public Property Get LogonCode() As String
    LogonCode = m_strLogonCode
End Property
Public Property Let FirstName(vData As String)
    m_strFirstName = vData
End Property
Public Property Get FirstName() As String
    FirstName = m_strFirstName
End Property

Public Property Let Name(vData As String)
    m_strName = vData
End Property
Public Property Get Name() As String
    Name = m_strName
End Property
Public Property Let ContactCode(vData As String)
    m_strContactCode = vData
End Property
Public Property Get ContactCode() As String
    ContactCode = m_strContactCode
End Property
Public Property Let Email(vData As String)
    m_strEmail = vData
End Property
Public Property Get Email() As String
    Email = m_strEmail
End Property
Public Property Let Password(vData As String)
    m_strPassword = vData
End Property
Public Property Get Password() As String
    Password = m_strPassword
End Property
Public Property Let DSN(vData As String)
    m_strDSN = vData
End Property
Public Property Get DSN() As String
    DSN = m_strDSN
End Property

Public Property Let CurrentBusinessUnitCode(vData As String)
    m_strCurrentBusinessUnitCode = vData
End Property
Public Property Get CurrentBusinessUnitCode() As String
    CurrentBusinessUnitCode = m_strCurrentBusinessUnitCode
End Property
Public Property Let CurrentBusinessUnit(vData As String)
    m_strCurrentBusinessUnit = vData
End Property
Public Property Get CurrentBusinessUnit() As String
    CurrentBusinessUnit = m_strCurrentBusinessUnit
End Property
Public Sub DeSerialize(sData As String)
    
    Dim objPB As XMLPropertyBag
    Set objPB = New XMLPropertyBag
    
    '***    is the data stream a message from another sssMessage
    If objPB.ReadProperty("sssMessage") <> "TRUE" Then
        m_bMessage = True
    Else
        Exit Sub
    End If
    
    m_sMessageId = objPB.ReadProperty("MessageId")
    m_sCategory = objPB.ReadProperty("Category")
    m_dMessageDate = objPB.ReadProperty("MessageDate")
    m_iPriority = objPB.ReadProperty("Priority")
    m_lSessionId = objPB.ReadProperty("SessionId")
    m_sMessage = objPB.ReadProperty("Message")
    m_strLogonCode = objPB.ReadProperty("LogonCode")
    m_strFirstName = objPB.ReadProperty("FirstName")
    m_strName = objPB.ReadProperty("Name")
    m_strContactCode = objPB.ReadProperty("ContactCode")
    m_strEmail = objPB.ReadProperty("Email")
    m_strPassword = objPB.ReadProperty("Password")
    m_strDSN = objPB.ReadProperty("DSN")
    m_strCurrentBusinessUnitCode = objPB.ReadProperty("CurrentBusinessUnitCode")
    m_strCurrentBusinessUnit = objPB.ReadProperty("CurrentBusinessUnit")
    
    '***    presumably the message is valid
    m_bValidMessage = True
    
End Sub
Public Function Serialize() As String
    Dim objPB As XMLPropertyBag
    
    Set objPB = New XMLPropertyBag
    
    ' Write the root node to the XML document
    objPB.Initialize "sssMessageRoot"
    objPB.WriteProperty "sssMessage", "TRUE"
    objPB.WriteProperty "MessageId", m_sMessageId
    objPB.WriteProperty "Category", m_sCategory
    objPB.WriteProperty "MessageDate", m_dMessageDate
    objPB.WriteProperty "Priority", m_iPriority
    objPB.WriteProperty "SessionId", m_lSessionId
    objPB.WriteProperty "Message", m_sMessage
    objPB.WriteProperty "LogonCode", m_strLogonCode
    objPB.WriteProperty "FirstName", m_strFirstName
    objPB.WriteProperty "Name", m_strName
    objPB.WriteProperty "ContactCode", m_strContactCode
    objPB.WriteProperty "Email", m_strEmail
    objPB.WriteProperty "Password", m_strPassword
    objPB.WriteProperty "DSN", m_strDSN
    objPB.WriteProperty "CurrentBusinessUnitCode", m_strCurrentBusinessUnitCode
    objPB.WriteProperty "CurrentBusinessUnit", m_strCurrentBusinessUnit

    Serialize = objPB.Contents
    
    Set objPB = Nothing

End Function
    
Private Sub Class_Initialize()
    m_dMessageDate = Now
End Sub
