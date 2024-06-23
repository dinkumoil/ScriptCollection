'///////////////////////////////////////////////////////////////////////////////
'
' Header file for handling *.xml files (create, read, change, update, transform)
'
' Author: Andreas Heim
' Date:   30.09.2022
'
' Required header files (to be included before):
'   - ADO.vbs
'   - Utils.vbs
'
'///////////////////////////////////////////////////////////////////////////////


'Pretty-print methods
Const PrettyPrintType_TextNodes = 0  'Indentation is configurable, made of space characters, consumes more resources
Const PrettyPrintType_SAX       = 1  'Indentation is always one tab character per indent level, consumes less resources,
                                     'always inserts "standalone" attribute into XML header

'Error codes, see property LastError
Const ERR_NO_ERROR              =  0
Const ERR_FILE_NOT_FOUND        =  1
Const ERR_FILE_ALREADY_EXISTS   =  2
Const ERR_LOADING_FILE_FAILED   =  3
Const ERR_LOADING_STRING_FAILED =  4
Const ERR_SAVING_FILE_FAILED    =  5
Const ERR_NODE_PATH_NOT_EXISTS  =  6
Const ERR_NODE_NOT_EXISTS       =  7
Const ERR_NODE_ALREADY_EXISTS   =  8
Const ERR_ATTR_NOT_EXISTS       =  9
Const ERR_ATTR_ALREADY_EXISTS   = 10




Class clsXmlFile
  Private NODE_ELEMENT
  Private NODE_ATTRIBUTE
  Private NODE_TEXT
  Private NODE_CDATA_SECTION
  Private NODE_ENTITY_REFERENCE
  Private NODE_ENTITY
  Private NODE_PROCESSING_INSTRUCTION
  Private NODE_COMMENT
  Private NODE_DOCUMENT
  Private NODE_DOCUMENT_TYPE
  Private NODE_DOCUMENT_FRAGMENT
  Private NODE_NOTATION

  Private objFSO
  Private objXmlDoc
  Private objXslDoc
  Private objOutStream

  Private strXmlFile
  Private strXslFile
  Private strOutFile

  Private strEncoding
  Private strStandalone
  Private strNamespaces

  Private intIndentSize
  Private intPrettyPrintType
  Private intLastError


  '-----------------------------------------------------------------------------
  'Constructor
  '-----------------------------------------------------------------------------
  Private Sub Class_Initialize()
    'Init pseudo private constants
    NODE_ELEMENT                =  1
    NODE_ATTRIBUTE              =  2
    NODE_TEXT                   =  3
    NODE_CDATA_SECTION          =  4
    NODE_ENTITY_REFERENCE       =  5
    NODE_ENTITY                 =  6
    NODE_PROCESSING_INSTRUCTION =  7
    NODE_COMMENT                =  8
    NODE_DOCUMENT               =  9
    NODE_DOCUMENT_TYPE          = 10
    NODE_DOCUMENT_FRAGMENT      = 11
    NODE_NOTATION               = 12

    Call Clear()
    Call Init()

    intPrettyPrintType = PrettyPrintType_TextNodes
    intIndentSize      = 4
  End Sub


  '-----------------------------------------------------------------------------
  'Destructor
  '-----------------------------------------------------------------------------
  Private Sub Class_Terminate()
    Call Clear()
  End Sub


  '-----------------------------------------------------------------------------
  'Discard all internal objects
  '-----------------------------------------------------------------------------
  Private Sub Clear
    Set objFSO       = Nothing
    Set objXmlDoc    = Nothing
    Set objXslDoc    = Nothing
    Set objOutStream = Nothing
    intLastError     = ERR_NO_ERROR
  End Sub


  '-----------------------------------------------------------------------------
  'Init internal objects
  '-----------------------------------------------------------------------------
  Private Sub Init
    strEncoding   = "utf-8"
    strStandalone = "yes"
    strNamespaces = ""

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    Set objXmlDoc = CreateObject("MSXml2.DOMDocument.6.0")
    objXmlDoc.async              = False
    objXmlDoc.preserveWhiteSpace = False
    objXmlDoc.validateOnParse    = False

    Set objXslDoc = CreateObject("MSXml2.FreeThreadedDOMDocument.6.0")
    objXslDoc.async              = False
    objXslDoc.preserveWhiteSpace = False
    objXslDoc.validateOnParse    = False

    Set objOutStream = CreateObject("ADODB.Stream")
    objOutStream.Mode = adModeReadWrite Or adModeShareDenyWrite
    objOutStream.Type = adTypeBinary
  End Sub


  '-----------------------------------------------------------------------------
  'Pretty-Print XML data into output stream
  '-----------------------------------------------------------------------------
  Private Function PrettyPrint(ByRef strOutEncoding)
    Select Case intPrettyPrintType
      Case PrettyPrintType_TextNodes
        PrettyPrint = PrettyPrintTextNodes(strOutEncoding)

      Case PrettyPrintType_SAX
        PrettyPrint = PrettyPrintSAX(strOutEncoding)
    End Select
  End Function


  '-----------------------------------------------------------------------------
  'Pretty-Print XML data by a recursive algorithm using self-inserted
  'XML text nodes
  '-----------------------------------------------------------------------------
  Private Function PrettyPrintTextNodes(ByRef strOutEncoding)
    Dim objTmpXmlDoc, objNode

    PrettyPrintTextNodes = True

    'Use a temporary XML DOM object for pretty-printing ...
    Set objTmpXmlDoc = CreateObject("MSXml2.DOMDocument.6.0")
    objTmpXmlDoc.async              = False
    objTmpXmlDoc.preserveWhiteSpace = False
    objTmpXmlDoc.validateOnParse    = False

    '... in order to avoid changing the original XML DOM object
    Call objXmlDoc.save(objTmpXmlDoc)
    Call FormatXmlNode(objTmpXmlDoc.documentElement, 0, intIndentSize)

    'Set character encoding of temporary XML DOM object ...
    Set objNode = objTmpXmlDoc.firstChild.attributes.getNamedItem("encoding")

    If Not objNode Is Nothing Then
      objNode.nodeValue = strOutEncoding
    End If

    '... and write it to the output stream
    objOutStream.Open
    Call objTmpXmlDoc.save(objOutStream)
  End Function


  '-----------------------------------------------------------------------------
  'Recursive function to pretty-print XML data using XML text nodes
  '-----------------------------------------------------------------------------
  Private Sub FormatXmlNode(ByRef objNode, intIndent, intIndentSize)
    Dim objChild, bolTextOnly

    'Do nothing if this is a text node
    If objNode.nodeType = NODE_TEXT Then Exit Sub

    'Check if this node contains only text
    bolTextOnly = True

    If objNode.hasChildNodes Then
      For Each objChild In objNode.childNodes
        If objChild.nodeType <> NODE_TEXT Then
          bolTextOnly = False
          Exit For
        End If
      Next
    End If

    'Process child nodes
    If objNode.hasChildNodes Then
      'Add a carriage return before the children
      If Not bolTextOnly Then
        Call objNode.insertBefore(objNode.ownerDocument.createTextNode(vbCrLf), objNode.firstChild)
      End If

      'Format the children
      For Each objChild In objNode.childNodes
        Call FormatXmlNode(objChild, intIndent + intIndentSize, intIndentSize)
      Next
    End If

    'Format this element
    If intIndent > 0 Then
      'Indent before this element
      Call objNode.parentNode.insertBefore(objNode.ownerDocument.createTextNode(String(intIndent, " ")), objNode)

      'Indent after the last child node
      If Not bolTextOnly Then
        Call objNode.appendChild(objNode.ownerDocument.createTextNode(String(intIndent, " ")))
      End If

      'Add a carriage return after this node
      If objNode.nextSibling Is Nothing Then
        Call objNode.parentNode.appendChild(objNode.ownerDocument.createTextNode(vbCrLf))
      Else
        Call objNode.parentNode.insertBefore(objNode.ownerDocument.createTextNode(vbCrLf), objNode.nextSibling)
      End If
    End If
  End Sub


  '-----------------------------------------------------------------------------
  'Pretty-Print XML data using a SAXXMLReader and a MXXMLWriter ActiveX object
  '-----------------------------------------------------------------------------
  Private Function PrettyPrintSAX(ByRef strOutEncoding)
    Dim objXmlReader, objXmlWriter

    PrettyPrintSAX = True

    With objOutStream
      .Open

      'Create XML writer that writes to the output stream
      Set objXmlWriter = CreateObject("MSXML2.MXXMLWriter")

      With objXmlWriter
        .omitXMLDeclaration = False
        .standalone         = (LCase(strStandalone) = "yes")
        .byteOrderMark      = False  'If not set (even to False) then '.encoding'
        .indent             = True   'is ignored
        .encoding           = strOutEncoding
        .output             = objOutStream

        'Create XML reader to parse input XML data and
        'hand over parsing result to XML writer
        Set objXmlReader = CreateObject("MSXML2.SAXXMLReader")

        With objXmlReader
          Set .contentHandler = objXmlWriter
          Set .dtdHandler     = objXmlWriter
          Set .errorHandler   = objXmlWriter

          Call .putProperty("http://xml.org/sax/properties/lexical-handler", objXmlWriter)
          Call .putProperty("http://xml.org/sax/properties/declaration-handler", objXmlWriter)

          'Pretty-print XML data and write result to output stream
          Call .parse(objXmlDoc)
        End With
      End With
    End With
  End Function


  '-----------------------------------------------------------------------------
  'Determine the size (in bytes) of a character encoding's BOM (Byte Order Mark)
  '-----------------------------------------------------------------------------
  Private Function EncodingBOMSize(ByRef strEncoding)
    Dim arrEncodings2ByteBom, arrEncodings3ByteBom

    arrEncodings2ByteBom = Array("utf-16", "utf-16LE", "unicode", "utf-16BE", "unicodeFFFE")
    arrEncodings3ByteBom = Array("utf-7", "utf-8")

    If UBound(Filter(arrEncodings3ByteBom, strEncoding, True, vbTextCompare)) >= 0 Then
      EncodingBOMSize = 3

    ElseIf UBound(Filter(arrEncodings2ByteBom, strEncoding, True, vbTextCompare)) >= 0 Then
      EncodingBOMSize = 2

    Else
      EncodingBOMSize = 0
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Save output stream to file without writing a BOM (Byte Order Mark)
  '-----------------------------------------------------------------------------
  Private Sub SaveOutStreamWithoutBOM(ByRef strFilePath, ByRef strEncoding)
    Dim objTmpStream

    'Create temporary binary stream
    Set objTmpStream = CreateObject("ADODB.Stream")
    objTmpStream.Mode = adModeReadWrite Or adModeShareDenyWrite
    objTmpStream.Type = adTypeBinary

    'Copy output stream to temporary binary stream
    objTmpStream.Open
    Call objOutStream.CopyTo(objTmpStream)

    'Skip BOM and save temporary binary stream to file
    objTmpStream.Position = EncodingBOMSize(strEncoding)
    Call objTmpStream.SaveToFile(strFilePath, adSaveCreateOverWrite)

    'Close temporary binary stream
    objTmpStream.Close
  End Sub


  '-----------------------------------------------------------------------------
  'Get/Set character encoding of currently loaded XML data
  '-----------------------------------------------------------------------------
  Public Property Get Encoding
    Encoding = strEncoding
  End Property


  Public Property Let Encoding(strValue)
    strEncoding = strValue
  End Property


  '-----------------------------------------------------------------------------
  'Get/Set standalone attribute of currently loaded XML data
  '-----------------------------------------------------------------------------
  Public Property Get Standalone
    Standalone = strStandalone
  End Property


  Public Property Let Standalone(strValue)
    strStandalone = strValue
  End Property


  '-----------------------------------------------------------------------------
  'Get/Set namespaces of currently loaded XML data. Namespaces must be given as
  'space-separated list in the form of "xmlns:<namespace alias>='<namespace id'"
  '-----------------------------------------------------------------------------
  Public Property Get Namespaces
    Namespaces = strNamespaces
  End Property


  Public Property Let Namespaces(strValue)
    strNamespaces = strValue

    If Not objXmlDoc Is Nothing Then
      Call objXmlDoc.setProperty("SelectionNamespaces", strNameSpaces)
    End If
  End Property


  '-----------------------------------------------------------------------------
  'Get/Set pretty-printing method
  '-----------------------------------------------------------------------------
  Public Property Get PrettyPrintType
    PrettyPrintType = intPrettyPrintType
  End Property


  Public Property Let PrettyPrintType(intValue)
    intPrettyPrintType = intValue
  End Property


  '-----------------------------------------------------------------------------
  'Get/Set indentation size when pretty-printing currently loaded XML data
  '-----------------------------------------------------------------------------
  Public Property Get IndentSize
    IndentSize = intIndentSize
  End Property


  Public Property Let IndentSize(intValue)
    intIndentSize = intValue
  End Property


  '-----------------------------------------------------------------------------
  'Get result of last call
  '-----------------------------------------------------------------------------
  Public Property Get LastError
    LastError = intLastError
  End Property


  '-----------------------------------------------------------------------------
  'Load XML file
  '-----------------------------------------------------------------------------
  Public Function LoadFromFile(ByRef strFilePath)
    Dim objAttr

    'Error-exit if input file doesn't exist
    If Not objFSO.FileExists(strFilePath) Then
      intLastError = ERR_FILE_NOT_FOUND
      LoadFromFile = False
      Exit Function
    End If

    'Discard existing internal objects and reinit instance
    Call Clear()
    Call Init()

    'Error-exit if loading input file failed
    If Not objXmlDoc.load(strFilePath) Then
      intLastError = ERR_LOADING_FILE_FAILED
      LoadFromFile = False
      Exit Function
    End If

    'Retrieve character encoding from input XML
    Set objAttr = objXmlDoc.firstChild.attributes.getNamedItem("encoding")

    If Not objAttr Is Nothing Then
      strEncoding = objAttr.nodeValue
    End If

    'Retrieve standalone attribute from input XML
    Set objAttr = objXmlDoc.firstChild.attributes.getNamedItem("standalone")

    If Not objAttr Is Nothing Then
      strStandalone = objAttr.nodeValue
    End If

    'Set namespaces to be able to query XML nodes that are defined in namespaces
    If strNamespaces <> "" Then
      Call objXmlDoc.setProperty("SelectionNamespaces", strNameSpaces)
    End If

    'Return success
    intLastError = ERR_NO_ERROR
    LoadFromFile = True
  End Function


  '-----------------------------------------------------------------------------
  'Load XML from string
  '-----------------------------------------------------------------------------
  Public Function LoadFromString(ByRef strXmlString)
    Dim objAttr

    'Discard existing internal objects and reinit instance
    Call Clear()
    Call Init()

    'Error-exit if loading input string failed
    If Not objXmlDoc.loadXml(strXmlString) Then
      intLastError   = ERR_LOADING_STRING_FAILED
      LoadFromString = False
      Exit Function
    End If

    'Retrieve character encoding from input XML
    Set objAttr = objXmlDoc.firstChild.attributes.getNamedItem("encoding")

    If Not objAttr Is Nothing Then
      strEncoding = objAttr.nodeValue
    End If

    'Retrieve standalone attribute from input XML
    Set objAttr = objXmlDoc.firstChild.attributes.getNamedItem("standalone")

    If Not objAttr Is Nothing Then
      strStandalone = objAttr.nodeValue
    End If

    'Return success
    intLastError   = ERR_NO_ERROR
    LoadFromString = True
  End Function


  '-----------------------------------------------------------------------------
  'Save XML data to file with original encoding, no pretty-printing
  '-----------------------------------------------------------------------------
  Public Function SaveToFile(ByRef strFilePath)
    SaveToFile = SaveToFileAs(strFilePath, strEncoding)
  End Function


  '-----------------------------------------------------------------------------
  'Save XML data to file with original encoding, with pretty-printing
  '-----------------------------------------------------------------------------
  Public Function SaveToFilePretty(ByRef strFilePath)
    SaveToFilePretty = SaveToFilePrettyAs(strFilePath, strEncoding)
  End Function


  '-----------------------------------------------------------------------------
  'Save XML data to file with new encoding, no pretty-printing
  '-----------------------------------------------------------------------------
  Public Function SaveToFileAs(ByRef strFilePath, ByRef strOutEncoding)
    Dim objNode

    'Set character encoding for output
    Set objNode = objXmlDoc.firstChild.attributes.getNamedItem("encoding")

    If Not objNode Is Nothing Then
      objNode.nodeValue = strOutEncoding
    End If

    'Write XML DOM to output stream
    objOutStream.Open
    Call objXmlDoc.save(objOutStream)

    'Save output stream to file
    objOutStream.Position = 0
    objOutStream.Type     = adTypeText
    objOutStream.CharSet  = strOutEncoding

    'If output encoding is not UTF-X ...
    If EncodingBOMSize(strOutEncoding) = 0 Then
      '... save output stream directly to output file
      Call objOutStream.SaveToFile(strFilePath, adSaveCreateOverWrite)
    Else
      '... otherwise save output stream without writing a BOM (Byte Order Mark)
      Call SaveOutStreamWithoutBOM(strFilePath, strOutEncoding)
    End If

    'Reset output stream
    objOutStream.Close
    objOutStream.Type = adTypeBinary

    'Return success
    intLastError = ERR_NO_ERROR
    SaveToFileAs = True
  End Function


  '-----------------------------------------------------------------------------
  'Save XML data to file with new encoding, with pretty-printing
  '-----------------------------------------------------------------------------
  Public Function SaveToFilePrettyAs(ByRef strFilePath, ByRef strOutEncoding)
    'Pretty-print XML DOM to output stream using provided character encoding
    'Error-exit if pretty-printing failed
    If Not PrettyPrint(strOutEncoding) Then
      intLastError       = ERR_SAVING_FILE_FAILED
      SaveToFilePrettyAs = False
      Exit Function
    End If

    'Save output stream to file
    objOutStream.Position = 0
    objOutStream.Type     = adTypeText
    objOutStream.CharSet  = strOutEncoding

    'If output encoding is not UTF-X ...
    If EncodingBOMSize(strOutEncoding) = 0 Then
      '... save output stream directly to output file
      Call objOutStream.SaveToFile(strFilePath, adSaveCreateOverWrite)
    Else
      '... otherwise save output stream without writing a BOM (Byte Order Mark)
      Call SaveOutStreamWithoutBOM(strFilePath, strOutEncoding)
    End If

    'Reset output stream
    objOutStream.Close
    objOutStream.Type = adTypeBinary

    'Return success
    intLastError       = ERR_NO_ERROR
    SaveToFilePrettyAs = True
  End Function


  '-----------------------------------------------------------------------------
  'Return XML data as string, no pretty-printing
  '-----------------------------------------------------------------------------
  Public Function ToString
    'Write XML DOM to output stream
    objOutStream.Open
    Call objXmlDoc.save(objOutStream)

    'Write content of output stream to result string
    objOutStream.Position = 0
    objOutStream.Type     = adTypeText
    objOutStream.CharSet  = strEncoding

    ToString = objOutStream.ReadText(adReadAll)

    'Reset output stream
    objOutStream.Close
    objOutStream.Type = adTypeBinary

    'Return success
    intLastError = ERR_NO_ERROR
  End Function


  '-----------------------------------------------------------------------------
  'Return XML data as string, with pretty-printing
  '-----------------------------------------------------------------------------
  Public Function ToStringPretty
    'Pretty-print XML DOM to output stream
    'Error-exit if pretty-printing failed
    If Not PrettyPrint(strEncoding) Then
      intLastError   = ERR_SAVING_FILE_FAILED
      ToStringPretty = ""
      Exit Function
    End If

    'Write content of output stream to result string
    objOutStream.Position = 0
    objOutStream.Type     = adTypeText
    objOutStream.CharSet  = strEncoding

    ToStringPretty = objOutStream.ReadText(adReadAll)

    'Reset output stream
    objOutStream.Close
    objOutStream.Type = adTypeBinary

    'Return success
    intLastError = ERR_NO_ERROR
  End Function


  '-----------------------------------------------------------------------------
  'Perform XSL transformation on currently loaded XML data
  'using provided XSL file
  '-----------------------------------------------------------------------------
  Public Function XslTransform(ByRef strXslFilePath)
    XslTransform = XslTransformWithParams(strXslFilePath, Array())
  End Function


  '-----------------------------------------------------------------------------
  'Perform XSL transformation on currently loaded XML data using provided XSL
  'file and array of parameters. Parameters must be given as strings containing
  'key-value pairs in the form of "key='value'" or "key=value".
  '-----------------------------------------------------------------------------
  Public Function XslTransformWithParams(ByRef strXslFilePath, ByRef arrParams)
    Dim objRegEx, colMatches, objMatch, intIdx, dicParams
    Dim objXslt, objXslProc
    Dim objNode, objAttr
    Dim strParam

    'Error-exit if XSLT file does not exist
    If Not objFSO.FileExists(strXslFilePath) Then
      intLastError           = ERR_FILE_NOT_FOUND
      XslTransformWithParams = False
      Exit Function
    End If

    'Read XSLT file, error-exit when failed
    If Not objXslDoc.load(strXslFilePath) Then
      intLastError           = ERR_LOADING_FILE_FAILED
      XslTransformWithParams = False
      Exit Function
    End If

    'Decode parameters array
    Set dicParams = CreateObject("Scripting.Dictionary")

    If IsArray(arrParams) Then
      Set objRegEx        = New RegExp
      objRegEx.Global     = True
      objRegEx.IgnoreCase = True
      objRegEx.Pattern    = "\s*([^\s=]+)\s*=\s*'?([^\s']+)'?\s*"

      For intIdx = 0 To UBound(arrParams)
        Set colMatches = objRegEx.Execute(arrParams(intIdx))

        For Each objMatch In colMatches
          If objMatch.SubMatches.Count = 2 Then
            dicParams.Item(objMatch.SubMatches(0)) = objMatch.SubMatches(1)
          End If
        Next
      Next
    End If

    'Create XSL template object and open output stream
    Set objXslt = CreateObject("Msxml2.XSLTemplate.6.0")
    objOutStream.Open

    'Try to retrieve 'output' node from XSLT file
    Set objNode = objXslDoc.documentElement.selectSingleNode("./*[local-name()='output']")

    'When failed, create that node
    If objNode Is Nothing Then
      Set objNode = objXslDoc.CreateNode(NODE_ELEMENT, "output", "http://www.w3.org/1999/XSL/Transform")
      Call objXslDoc.documentElement.insertBefore(objNode, objXslDoc.documentElement.firstChild)
    End If

    'Try to retrieve 'encoding' attribute from 'output' node
    Set objAttr = objNode.attributes.getNamedItem("encoding")

    'When failed, create that attribute
    If objAttr Is Nothing Then
      Set objAttr = objXslDoc.CreateNode(NODE_ATTRIBUTE, "encoding", "")
      objNode.attributes.setNamedItem(objAttr)
    End If

    'Set XSLT file's 'output[@encoding]' attribute
    'to encoding of XML file to transform
    objAttr.nodeValue = strEncoding

    'Set XSL template's XSL file and create XSL processor
    Set objXslt.stylesheet = objXslDoc
    Set objXslProc         = objXslt.createProcessor()

    'Set XSL processor's input and output object
    objXslProc.input       = objXmlDoc
    objXslProc.output      = objOutStream

    'Perform XSL transformation without parameters
    If dicParams.Count = 0 Then
      Call objXslProc.transform()

    'Perform XSL transformation with parameters provided on command line
    Else
      For Each strParam In dicParams.Keys
        'Set output stream's file pointer back to beginning
        objOutStream.Position = 0

        'Add XSL parameter to XSL processor
        Call objXslProc.addParameter(strParam, dicParams.Item(strParam))

        'Perform XSL transformation with last added parameter
        Call objXslProc.transform()
      Next
    End If

    'Replace XML DOM with transformed version
    objOutStream.Position = 0
    Call objXmlDoc.load(objOutStream)

    'Reset output stream
    objOutStream.Close
    objOutStream.Type = adTypeBinary

    'Return success
    intLastError           = ERR_NO_ERROR
    XslTransformWithParams = True
  End Function


  '-----------------------------------------------------------------------------
  'Create XML document node, the document's root node with an optional attribute
  '(optionally set to a certain value) and an optional node text
  '-----------------------------------------------------------------------------
  Public Function CreateDocument(ByRef strRootNodeName, ByRef strAttrName, ByRef strAttrValue, ByRef strNodeValue)
    Dim objPINode, objRootNode, objNewAttr

    'Discard existing internal objects and reinit instance
    Clear()
    Init()

    'Create XML header node and insert it into XML DOM
    Set objPINode = objXmlDoc.createProcessingInstruction("xml", FormatString("version=""1.0"" encoding=""%1"" standalone=""%2""", Array(strEncoding, strStandalone)))
    Call objXmlDoc.insertBefore(objPINode, objXmlDoc.childNodes.item(0))

    'Error-exit if XML header node already exists
    If Not objXmlDoc.documentElement Is Nothing Then
      intLastError   = ERR_NODE_ALREADY_EXISTS
      CreateDocument = False
    Else
      'Create XML root node ...
      Set objRootNode = objXmlDoc.createElement(strRootNodeName)

      '... and set its provided value
      If strNodeValue <> "" Then
        objRootNode.text = strNodeValue
      End If

      'Create an attribute and set its value if provided
      If strAttrName <> "" Then
        Set objNewAttr   = objXmlDoc.createAttribute(strAttrName)
        objNewAttr.Value = strAttrValue

        Call objRootNode.setAttributeNode(objNewAttr)
      End If

      'Append root node to XML document
      Call objXmlDoc.appendChild(objRootNode)

      'Return success
      intLastError   = ERR_NO_ERROR
      CreateDocument = True
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Create an XML node with an optional attribute (optionally set to a certain
  'value) and an optional node text and add it as a child node to a provided
  'XML node
  '-----------------------------------------------------------------------------
  Public Function AddNode(ByRef strParentNodePath, ByRef strNodeName, ByRef strAdrAttrName, ByRef strAdrAttrValue, ByRef strNodeValue)
    Dim colNodes, colTmpNodes, objParentNode, objNewNode, objNewAttr, bolExists

    Set colNodes = objXmlDoc.selectNodes(strParentNodePath)

    If colNodes.length = 0 Then
      intLastError = ERR_NODE_PATH_NOT_EXISTS
      AddNode      = False
    Else
      bolExists = False

      For Each objParentNode In colNodes
        If strAdrAttrName <> "" Then
          Set colTmpNodes = objParentNode.selectNodes(FormatString("./%1[@%2='%3']", Array(strNodeName, strAdrAttrName, strAdrAttrValue)))

          If colTmpNodes.length > 0 Then
            bolExists = True
          Else
            Set objNewNode = objXmlDoc.createElement(strNodeName)

            If strNodeValue <> "" Then
              objNewNode.text = strNodeValue
            End If

            Set objNewAttr   = objXmlDoc.createAttribute(strAdrAttrName)
            objNewAttr.Value = strAdrAttrValue

            Call objNewNode.setAttributeNode(objNewAttr)
            Call objParentNode.appendChild(objNewNode)
          End If
        Else
          Set objNewNode = objXmlDoc.createElement(strNodeName)

          If strNodeValue <> "" Then
            objNewNode.text = strNodeValue
          End If

          Call objParentNode.appendChild(objNewNode)
        End If
      Next

      If bolExists Then
        intLastError = ERR_NODE_ALREADY_EXISTS
        AddNode      = False
      Else
        intLastError = ERR_NO_ERROR
        AddNode      = True
      End IF
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Create an XML node with an optional attribute (optionally set to a certain
  'value) and an optional node text and insert it beside a provided sibling node
  '-----------------------------------------------------------------------------
  Public Function InsertNode(ByRef strSiblingNodePath, ByRef strNodeName, ByRef strAdrAttrName, ByRef strAdrAttrValue, ByRef strNodeValue)
    Dim colNodes, colTmpNodes, objRefNode, objParentNode, objNewNode, objNewAttr, bolExists

    Set colNodes = objXmlDoc.selectNodes(strSiblingNodePath)

    If colNodes.length = 0 Then
      intLastError = ERR_NODE_PATH_NOT_EXISTS
      InsertNode   = False
    Else
      bolExists = False

      For Each objRefNode In colNodes
        If strAdrAttrName <> "" Then
          Set objParentNode = objRefNode.parentNode

          Set colTmpNodes = objParentNode.selectNodes(FormatString("./%1[@%2='%3']", Array(strNodeName, strAdrAttrName, strAdrAttrValue)))

          If colTmpNodes.length > 0 Then
            bolExists = True
          Else
            Set objNewNode = objXmlDoc.createElement(strNodeName)

            If strNodeValue <> "" Then
              objNewNode.text = strNodeValue
            End If

            Set objNewAttr   = objXmlDoc.createAttribute(strAdrAttrName)
            objNewAttr.Value = strAdrAttrValue

            Call objNewNode.setAttributeNode(objNewAttr)

            If Not IsNull(objRefNode.nextSibling) Then
              Call objParentNode.insertBefore(objNewNode, objRefNode.nextSibling)
            Else
              Call objParentNode.appendChild(objNewNode)
            End If
          End If
        Else
          Set objParentNode = objRefNode.parentNode

          Set objNewNode = objXmlDoc.createElement(strNodeName)

          If strNodeValue <> "" Then
            objNewNode.text = strNodeValue
          End If

          If Not IsNull(objRefNode.nextSibling) Then
            Call objParentNode.insertBefore(objNewNode, objRefNode.nextSibling)
          Else
            Call objParentNode.appendChild(objNewNode)
          End If
        End If
      Next

      If bolExists Then
        intLastError = ERR_NODE_ALREADY_EXISTS
        InsertNode   = False
      Else
        intLastError = ERR_NO_ERROR
        InsertNode   = True
      End IF
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Update an XML node's node text
  '-----------------------------------------------------------------------------
  Public Function UpdateNodeValue(ByRef strNodePath, ByRef strNodeValue)
    Dim colNodes, objNode

    Set colNodes = objXmlDoc.selectNodes(strNodePath)

    If colNodes.length = 0 Then
      intLastError    = ERR_NODE_NOT_EXISTS
      UpdateNodeValue = False
    Else
      For Each objNode In colNodes
        objNode.text = strNodeValue
      Next

      intLastError    = ERR_NO_ERROR
      UpdateNodeValue = True
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Read an XML node's node text
  '-----------------------------------------------------------------------------
  Public Function ReadNodeValue(ByRef strNodePath)
    Dim colNodes, objNode, intIdx, arrOutput()

    Set colNodes = objXmlDoc.selectNodes(strNodePath)

    If colNodes.length = 0 Then
      intLastError  = ERR_NODE_NOT_EXISTS
      ReadNodeValue = Array()
    Else
      ReDim arrOutput(colNodes.length - 1)
      intIdx = 0

      For Each objNode In colNodes
        arrOutput(intIdx) = objNode.text
        intIdx = intIdx + 1
      Next

      intLastError  = ERR_NO_ERROR
      ReadNodeValue = arrOutput
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Delete an XML node's node text
  '-----------------------------------------------------------------------------
  Public Function DeleteNodeValue(ByRef strNodePath)
    Dim colNodes, objNode, objChildNode

    Set colNodes = objXmlDoc.selectNodes(strNodePath)

    If colNodes.length = 0 Then
      intLastError    = ERR_NODE_NOT_EXISTS
      DeleteNodeValue = False
    Else
      For Each objNode In colNodes
        For Each objChildNode In objNode.childNodes
          If objChildNode.nodeType = NODE_TEXT Then
            Call objNode.removeChild(objChildNode)
            Exit For
          End If
        Next
      Next

      intLastError    = ERR_NO_ERROR
      DeleteNodeValue = True
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Delete a whole XML node
  '-----------------------------------------------------------------------------
  Public Function DeleteNode(ByRef strNodePath)
    Dim colNodes, objNode

    Set colNodes = objXmlDoc.selectNodes(strNodePath)

    If colNodes.length = 0 Then
      intLastError = ERR_NODE_NOT_EXISTS
      DeleteNode   = False
    Else
      For Each objNode In colNodes
        Call objNode.parentNode.removeChild(objNode)
      Next

      intLastError = ERR_NO_ERROR
      DeleteNode   = True
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Create an attribute for the provided XML node and optionally set its value
  '-----------------------------------------------------------------------------
  Public Function AddAttribute(ByRef strNodePath, ByRef strAttrName, ByRef strAttrValue)
    Dim colNodes, objNode, objNewAttr, bolExists

    Set colNodes = objXmlDoc.selectNodes(strNodePath)

    If colNodes.length = 0 Then
      intLastError = ERR_NODE_NOT_EXISTS
      AddAttribute = False
    Else
      bolExists = False

      For Each objNode In colNodes
        Set objNewAttr = objNode.getAttributeNode(strAttrName)

        If Not objNewAttr Is Nothing Then
          bolExists = True
        Else
          Set objNewAttr   = objXmlDoc.createAttribute(strAttrName)
          objNewAttr.Value = strAttrValue

          Call objNode.setAttributeNode(objNewAttr)
        End If
      Next

      If bolExists Then
        intLastError = ERR_ATTR_ALREADY_EXISTS
        AddAttribute = False
      Else
        intLastError = ERR_NO_ERROR
        AddAttribute = True
      End If
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Update an attribute's value of the provided XML node
  '-----------------------------------------------------------------------------
  Public Function UpdateAttributeValue(ByRef strNodePath, ByRef strAttrName, ByRef strAttrValue)
    Dim colNodes, objNode, bolFound, objAttr

    Set colNodes = objXmlDoc.selectNodes(strNodePath)

    If colNodes.length = 0 Then
      intLastError         = ERR_NODE_PATH_NOT_EXISTS
      UpdateAttributeValue = False
    Else
      bolFound = False

      For Each objNode In colNodes
        Set objAttr = objNode.getAttributeNode(strAttrName)

        If Not objAttr Is Nothing Then
          objAttr.nodeValue = strAttrValue
          bolFound          = True
        End If
      Next

      If Not bolFound Then
        intLastError         = ERR_ATTR_NOT_EXISTS
        UpdateAttributeValue = False
      Else
        intLastError         = ERR_NO_ERROR
        UpdateAttributeValue = True
      End If
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Read an attribute's value of the provided XML node
  '-----------------------------------------------------------------------------
  Public Function ReadAttributeValue(ByRef strNodePath, ByRef strAttrName)
    Dim colNodes, objNode, objAttr, intIdx, bolFound, arrOutput()

    Set colNodes = objXmlDoc.selectNodes(strNodePath)

    If colNodes.length = 0 Then
      intLastError       = ERR_NODE_PATH_NOT_EXISTS
      ReadAttributeValue = Array()
    Else
      ReDim arrOutput(colNodes.length - 1)
      intIdx   = 0
      bolFound = False

      For Each objNode In colNodes
        Set objAttr = objNode.getAttributeNode(strAttrName)

        If Not objAttr Is Nothing Then
          arrOutput(intIdx) = objAttr.nodeValue
          intIdx            = intIdx + 1
          bolFound          = True
        End If
      Next

      If Not bolFound Then
        intLastError       = ERR_ATTR_NOT_EXISTS
        ReadAttributeValue = Array()
      Else
        ReDim Preserve arrOutput(intIdx - 1)
        intLastError       = ERR_NO_ERROR
        ReadAttributeValue = arrOutput
      End If
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Delete an attribute's value of the provided XML node
  '-----------------------------------------------------------------------------
  Public Function DeleteAttributeValue(ByRef strNodePath, ByRef strAttrName)
    Dim colNodes, objNode, objAttr, bolFound

    Set colNodes = objXmlDoc.selectNodes(strNodePath)

    If colNodes.length = 0 Then
      intLastError         = ERR_NODE_PATH_NOT_EXISTS
      DeleteAttributeValue = False
    Else
      bolFound = False

      For Each objNode In colNodes
        Set objAttr = objNode.getAttributeNode(strAttrName)

        If Not objAttr Is Nothing Then
          objAttr.nodeValue = ""
          bolFound          = True
        End If
      Next

      If Not bolFound Then
        intLastError         = ERR_ATTR_NOT_EXISTS
        DeleteAttributeValue = False
      Else
        intLastError         = ERR_NO_ERROR
        DeleteAttributeValue = True
      End If
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Delete a whole attribute of the provided XML node
  '-----------------------------------------------------------------------------
  Public Function DeleteAttribute(ByRef strNodePath, ByRef strAttrName)
    Dim colNodes, objNode

    Set colNodes = objXmlDoc.selectNodes(strNodePath)

    If colNodes.length = 0 Then
      intLastError    = ERR_NODE_PATH_NOT_EXISTS
      DeleteAttribute = False
    Else
      For Each objNode In colNodes
        Call objNode.removeAttribute(strAttrName)
      Next

      intLastError    = ERR_NO_ERROR
      DeleteAttribute = True
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Create an XML comment with an optional node text and add it as a child
  'node to a provided XML node
  '-----------------------------------------------------------------------------
  Public Function AddComment(ByRef strParentNodePath, ByRef strNodeValue)
    Dim colNodes, objParentNode, objNewNode

    Set colNodes = objXmlDoc.selectNodes(strParentNodePath)

    If colNodes.length = 0 Then
      intLastError = ERR_NODE_PATH_NOT_EXISTS
      AddComment   = False
    Else
      For Each objParentNode In colNodes
        Set objNewNode = objXmlDoc.createComment(strNodeValue)
        Call objParentNode.appendChild(objNewNode)
      Next

      intLastError = ERR_NO_ERROR
      AddComment   = True
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Create an XML comment with an optional node text and insert it beside
  'a provided sibling node
  '-----------------------------------------------------------------------------
  Public Function InsertComment(ByRef strSiblingNodePath, ByRef strNodeValue)
    Dim colNodes, objRefNode, objParentNode, objNewNode

    Set colNodes = objXmlDoc.selectNodes(strSiblingNodePath)

    If colNodes.length = 0 Then
      intLastError  = ERR_NODE_PATH_NOT_EXISTS
      InsertComment = False
    Else
      For Each objRefNode In colNodes
        Set objParentNode = objRefNode.parentNode
        Set objNewNode    = objXmlDoc.createComment(strNodeValue)

        If Not IsNull(objRefNode.nextSibling) Then
          Call objParentNode.insertBefore(objNewNode, objRefNode.nextSibling)
        Else
          Call objParentNode.appendChild(objNewNode)
        End If
      Next

      intLastError  = ERR_NO_ERROR
      InsertComment = True
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Create an XML CData section with an optional node text and add it as a child
  'node to a provided XML node
  '-----------------------------------------------------------------------------
  Public Function AddCDataSection(ByRef strParentNodePath, ByRef strNodeValue)
    Dim colNodes, objParentNode, objNewNode

    Set colNodes = objXmlDoc.selectNodes(strParentNodePath)

    If colNodes.length = 0 Then
      intLastError    = ERR_NODE_PATH_NOT_EXISTS
      AddCDataSection = False
    Else
      For Each objParentNode In colNodes
        Set objNewNode = objXmlDoc.createCDATASection(strNodeValue)
        Call objParentNode.appendChild(objNewNode)
      Next

      intLastError    = ERR_NO_ERROR
      AddCDataSection = True
    End If
  End Function


  '-----------------------------------------------------------------------------
  'Create an XML CData section with an optional node text and insert it beside
  'a provided sibling node
  '-----------------------------------------------------------------------------
  Public Function InsertCDataSection(ByRef strSiblingNodePath, ByRef strNodeValue)
    Dim colNodes, objRefNode, objParentNode, objNewNode

    Set colNodes = objXmlDoc.selectNodes(strSiblingNodePath)

    If colNodes.length = 0 Then
      intLastError       = ERR_NODE_PATH_NOT_EXISTS
      InsertCDataSection = False
    Else
      For Each objRefNode In colNodes
        Set objParentNode = objRefNode.parentNode
        Set objNewNode    = objXmlDoc.createCDATASection(strNodeValue)

        If Not IsNull(objRefNode.nextSibling) Then
          Call objParentNode.insertBefore(objNewNode, objRefNode.nextSibling)
        Else
          Call objParentNode.appendChild(objNewNode)
        End If
      Next

      intLastError       = ERR_NO_ERROR
      InsertCDataSection = True
    End If
  End Function
End Class
