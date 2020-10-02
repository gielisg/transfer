VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLPropertyBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' --------------------------------------
' Name:     XMLPropertyBag
' Author:   Darren Weir
'           Select Software Solutions
' Date:     3 Oct 2000
' Description:
'          This class is an XML version of the
'          VB6 PropertyBag. It allows for the
'          Serializing and Restoring of property name/value pairs.
'          It is particulalry useful for saving and restoring the
'          state of business objects as/using XML strings.
'
' NOTE:
'   In order to use this XMLPropertyBag class, you must have Microsoft XML, version 2.0 (or higher) installed
'
' Revisions:
'
' --------------------------------------

' The underlying storage mechanism for the XML PropertyBag
Private mdocXML As MSXML2.DOMDocument

' Flag to indicate if the PropertyBag has been initialized
Private mbIsInitialized As Boolean

Public Enum sssValueType
  icValue = 1
  icChildObject = 2
End Enum

' ----------------------------------
' Start of VB-accelerator Code Block
' ----------------------------------
Private Enum epbtTypeConstants
   epbtImplementsPropBag
   epbtEmpty
   epbtNull
   epbtNothing
   epbtByteArray
   epbtDate
   epbtNumber
   epbtString
   epbtStringCData
   epbtError
End Enum

Public Enum epbtErrorConstants
   eeXMLPropertyBagErrorBase = 29450
   eeInvalidVariantType
   eeInvalidXMLFile
   eeParseError ' Added by Darren
   eeNotInitialized ' Added by Darren
   eePropertyNotFound ' Added by Darren
   eeDuplicateEntry ' Added by Darren
End Enum
' ----------------------------------
' End of VB-accelerator Code Block
' ----------------------------------

' ------------------------------------------------------
' This method is part of the PUBLIC interface.
'
' This method writes a value into the XML PropertyBag
' Credit goes to VB-accelerator for some of the datatype
' handling code.
' ------------------------------------------------------
Public Sub WriteProperty(ByVal strName As String, ByVal vValue As Variant, Optional ByVal vDefault As Variant)
  Dim objNode As MSXML2.IXMLDOMElement
  Dim eType As epbtTypeConstants
  
  ' TODO: When does the default value get used???
  ' The default value is included for compatibility with the
  ' VB6 Property Bag

  ' Check to see if the object has been initialized
  If mbIsInitialized = False Then
    Err.Raise vbObjectError + 1048 + eeNotInitialized, "WriteProperty", "Property Bag has not been initialized"
  End If
    
  ' Evaluate the type of variant we are adding to the XML:
  eType = ValidateValueType(vValue)
   
  ' Write out accordingly:
  If eType = epbtError Then
    ' Can't write!
    Err.Raise (vbObjectError + 1048 + eeInvalidVariantType), App.EXEName & ".XMLPropertyBag", "Invalid PropertyValue type."
  ElseIf eType = epbtStringCData Then
    ' The Value is an XML string
    WriteChildProperty strName, vValue
  Else
  
    ' Check for duplicate values
    Set objNode = mdocXML.documentElement.selectSingleNode(strName)
    If Not objNode Is Nothing Then
      ' Duplicate nodes are not allowed, so raise an error!
      Err.Raise (vbObjectError + 1048 + eeDuplicateEntry), App.EXEName & ".XMLPropertyBag", "A " & strName & " property already exists in the XMLPropertyBag. Duplicate Names are not allowed."
    End If
  
    ' Create the node in the XML tree for this property
    Set objNode = mdocXML.createElement(strName)
  
    ' Write the appropriate value to the xml node according to the datatype
    Select Case eType
    
      Case epbtDate
        objNode.Text = vValue
        
      Case epbtNumber
        objNode.Text = vValue
        
      Case epbtString
        objNode.Text = vValue
                
      Case epbtByteArray
        With objNode
          .dataType = "bin.base64"
          .nodeTypedValue = vValue
        End With
        
      Case epbtEmpty
        ' Special values:
        objNode.setAttribute "vartype", "empty"
        
      Case epbtNothing
        ' Special values:
        objNode.setAttribute "vartype", "nothing"
        
      Case epbtNull
        ' Special values:
        objNode.setAttribute "vartype", "null"
        
    End Select
    
    ' Append the node to the XML tree
    mdocXML.documentElement.appendChild objNode
  
  End If
  
End Sub

' -------------------------------------------------------
' This method accepts the name and value of the property
' where the value is an xml subtree
' -------------------------------------------------------
Private Sub WriteChildProperty(ByVal strName As String, ByVal strXML As String, Optional ByVal docNode As MSXML2.IXMLDOMNode)

  Dim docXML As DOMDocument
  'Dim docNode As MSXML.IXMLDOMNode
  Dim objPropertyNode As MSXML2.IXMLDOMNode
  
  ' Check to see if the XML node for the child object
  ' has been passed in by another internal method
  ' If it hasn't, then create one
  If docNode Is Nothing Then
    ' Create the xml subtree of the Child object
    ' Example of the strXML parameter:
    ' <PhoneNumber>
    '   <AreaCode>
    '   </AreaCode>
    ' </PhoneNumber>
    Set docXML = New DOMDocument
    docXML.loadXML strXML
    
    ' Get the root Node of the child objects XML
    Set docNode = docXML.documentElement
    Set docXML = Nothing
  End If
  
  ' If a name is specified for this property, then create an element
  ' with the name passed in as a parameter. Otherwise, the XML value
  ' is just appended to the PropertyBag
  If Len(strName) <> 0 Then
  
    ' Create the Node for this property
    ' Example: strName = "HomePhoneNumber"
    ' <HomePhoneNumber></HomePhoneNumber>
    Set objPropertyNode = mdocXML.createElement(strName)
      
    ' Add the XML subtree of the child object to Node for this Property
    objPropertyNode.appendChild docNode
    ' Now the XML for this Child Property looks like:
    ' <HomePhoneNumber>
    '   <PhoneNumber>
    '     <AreaCode>
    '     </AreaCode>
    '   </PhoneNumber>
    ' </HomePhoneNumber>
  Else
    ' In this scenario, the <HomePhoneNumber></HomePhoneNumber> tags are
    ' not included as part of the subtree, only the child XML subtree
    Set objPropertyNode = docNode
  End If
  
  ' Now Append the whole lot to the XMLPropertyBag
  mdocXML.documentElement.appendChild objPropertyNode

End Sub

' -------------------------------------------------------------------------
' This method is part of the PUBLIC interface.
'
' This method attempts to find and return the value of the named Property
' If the property is not found, the default is returned if present, otherwise
' an error is generated
' -------------------------------------------------------------------------
Public Function ReadProperty(ByVal strName As String, Optional ByVal vDefault As Variant, Optional ByVal lngIndex As Long = -1) As Variant
  Dim objNode As MSXML.IXMLDOMElement
  
  ' Check to see if the object has been initialized
  If mbIsInitialized = False Then
    Err.Raise vbObjectError + 1048 + eeNotInitialized, "ReadProperty", "Property Bag has not been initialized"
  End If
  
  ' If an Index has been passed as a parameter, then look for
  ' one property out of several properties with the same name.
  If lngIndex <> -1 Then
    ReadProperty = ReadChildProperty(strName, lngIndex, vDefault)
    Exit Function
  End If
  
  On Error Resume Next
  ' Try to retrieve the node for the Property
  Set objNode = mdocXML.documentElement.selectSingleNode(strName)
  
  ' Check for any errors
  If Err Then
    On Error GoTo 0
    ' Raise an error
    Err.Raise vbObjectError + 1048 + eeParseError, "ReadProperty", Err.Description
  ElseIf objNode Is Nothing Then
    If IsMissing(vDefault) Then
      ' Raise an error
      Err.Raise vbObjectError + 1048 + eePropertyNotFound, "ReadProperty", "The " & strName & " property could not be found in the XMLPropertyBag."
    Else
      ' Property not found, so return the default value
      ReadProperty = vDefault
    End If
  Else
    ' Return the value
    If objNode.childNodes.length > 0 Then
      ReadProperty = objNode.firstChild.xml
    Else
      If objNode.dataType = "bin.base64" Then
        ReadProperty = objNode.nodeTypedValue
      Else
        ReadProperty = objNode.Text
      End If
    End If
    
  End If

End Function

' -------------------------------------------------------------------------------
' This method attempts to find and return the XML subtree of the child property.
' If the property is not found, the default is returned if present, otherwise
' an error is generated
' -------------------------------------------------------------------------------
Private Function ReadChildProperty(ByVal strName As String, ByVal lngIndex As Long, Optional ByVal vDefault As Variant) As Variant
  Dim objNode As MSXML.IXMLDOMElement
  Dim objTmpNode As MSXML.IXMLDOMElement
  Dim lngCount As Long
  Dim bFound As Boolean

  bFound = False

  On Error GoTo errHandler

  ' Get a child object in a collection of child elements
  lngCount = 1
  For Each objTmpNode In mdocXML.documentElement.selectNodes(strName)
    If lngCount = lngIndex Then
      Set objNode = objTmpNode
      bFound = True
      Exit For
    End If
    lngCount = lngCount + 1
  Next

errHandler:

  ' Check to see if the value was found
  If Err Then
    ' Raise an error
    Err.Raise vbObjectError + 1048 + eeParseError, "ReadProperty", Err.Description
  ElseIf bFound = False Or objNode Is Nothing Then
    ' The Property was not found
    If IsMissing(vDefault) Then
      ' Raise an error
      Err.Raise vbObjectError + 1048 + eePropertyNotFound, "ReadProperty", "The " & strName & " property could not be found in the XMLPropertyBag."
    Else
      ' Return the default value
      ReadChildProperty = vDefault
    End If
  Else
    ' Return the XML subtree of the child property object
    ReadChildProperty = objNode.xml
  End If

End Function

Private Sub Class_Initialize()
  mbIsInitialized = False
End Sub

Private Sub Class_Terminate()
  Set mdocXML = Nothing
End Sub

' -----------------------------------------------------
' This method is part of the PUBLIC interface.
'
' This method MUST be called before the XMLPropertyBag
' can be used. It inserts the root node into the
' underlying XML DOMDocument.
'
' Example:
' <Person></Person>
' where Person is the name of the class using the
' XMLPropertyBag
' -----------------------------------------------------
Public Sub Initialize(ByVal strObjectName As String)
  Dim objNode As MSXML2.IXMLDOMElement
  
  ' Create a new xml document
  Set mdocXML = New MSXML2.DOMDocument
  
  ' Create the Root node
  Set objNode = mdocXML.createElement(strObjectName)
  
  ' Add the root element to the xml tree
  mdocXML.appendChild objNode
  
  ' set the flag to true
  mbIsInitialized = True
    
End Sub

' -----------------------------------------------------
' This property is part of the PUBLIC interface.
'
' This property returns the XML representation of the
' contents of the XMLPropertyBag
' -----------------------------------------------------
Public Property Get Contents() As String
  Contents = mdocXML.xml
End Property

' -----------------------------------------------------
' This property is part of the PUBLIC interface.
'
' Causes the XMLProperyBag to load itself from the XML
' string passed in as a parameter.
' -----------------------------------------------------
Public Property Let Contents(ByVal strXML As String)

  On Error Resume Next

  ' Create a new XML document, overwriting any previous values
  Set mdocXML = New MSXML.DOMDocument
    
  ' Attempt to Load the XML into the XML DOMDocument
  mdocXML.loadXML strXML
  
  ' Check for any errors
  If Err Or mdocXML.parsed = False Then
    ' Raise an error
    Err.Raise (vbObjectError + 1048 + eeInvalidXMLFile), App.EXEName & ".XMLPropertyBag", "Invalid XML file."
  ElseIf mdocXML.parsed Then
    mbIsInitialized = True
  Else
    ' Raise an error
    Err.Raise (vbObjectError + 1048 + eeParseError), App.EXEName & ".XMLPropertyBag", "Unable to parse the XML file."
  End If
  
End Property

' ----------------------------------
' Start of VB-accelerator Code Block
' ----------------------------------
Private Function ValidateValueType(Value As Variant) As epbtTypeConstants
   
   ' Here we evaluate what type of variant we're
   ' trying to write out to the XML:
   ValidateValueType = epbtError
   If IsEmpty(Value) Then
      ValidateValueType = epbtEmpty
   ElseIf IsNull(Value) Then
      ValidateValueType = epbtNull
   ElseIf IsObject(Value) Then
      If Value Is Nothing Then
         ValidateValueType = epbtNothing
'      ElseIf pbImplementsXMLPropBag(Value) Then ' Commented out by Darren
'         ValidateValueType = epbtImplementsPropBag ' Commented out by Darren
      Else
         ' Evalute for default value.
         ValidateValueType = GetOtherType(Value)
      End If
   ElseIf IsError(Value) Then
      ValidateValueType = epbtError
   ElseIf IsNumeric(Value) Then
      ValidateValueType = epbtNumber
   ElseIf IsDate(Value) Then
      ValidateValueType = epbtDate
   ElseIf IsArray(Value) Then
      ' only byte arrays are supported:
      If (VarType(Value) And vbByte) = vbByte Then
         ValidateValueType = epbtByteArray
      End If
   Else
      ValidateValueType = GetOtherType(Value)
   End If
   
End Function

Private Function GetOtherType(Value As Variant) As epbtTypeConstants
Dim sThis As String
   
   On Error Resume Next
   sThis = Value
   If Err.Number = 0 Then
      GetOtherType = epbtString
      
      ' Check to see if the Value is an XML string
      Dim docXML As MSXML2.DOMDocument
      Dim bLoaded As Boolean
      
      Set docXML = New MSXML2.DOMDocument
      docXML.async = False
      
      bLoaded = docXML.loadXML(Value)
      If bLoaded And docXML.parsed = True Then
      
        GetOtherType = epbtStringCData
      End If
      
      ' Does sThis contain XML delimiters?
      ' NB: rough check.  If you're worried, pass the
      ' information into WriteProperties as a byte
      ' array & the bin.base64 will do its work, then
      ' you'll never have difficulties.
'      If Not (InStr(sThis, "<") = 0) Then
'         GetOtherType = epbtStringCData
'      ElseIf Not (InStr(sThis, ">") = 0) Then
'         GetOtherType = epbtStringCData
'      End If
      
   Else
      GetOtherType = epbtError
   End If
   Err.Clear
   
End Function

' ----------------------------------
' End of VB-accelerator Code Block
' ----------------------------------
