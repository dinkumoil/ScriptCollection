'*******************************************************************************
'
' Data container to hold MS SQL Server related data.
'
' Author: Andreas Heim
' Date  : June 2017
'
'
' Disclaimer
' ==========
'
' This code is released "as is" without any guarantee to work proper for a
' certain subject. The author is not responsible for any damage caused by
' running this code.
'
'
' License
' =======
'
' This code can be changed and used for free by everyone. It is not allowed to
' claim any rights on this code or to charge any fees for it. If you change this
' file or deploy it to other persons the disclaimer and license text has to be
' kept included as well as a reference to the original author.
'
'*******************************************************************************

Class clsSQLServer
  Private strInstanceName
  Private strInstanceExePath
  Private strVersion
  Private strVersionName
  Private strEdition
  Private intSPLevel
  Private bolIsDefaultInstance
  Private bolIsDomainMember


  Private Sub Class_Initialize()
    strInstanceName      = ""
    strInstanceExePath   = ""
    strVersion           = ""
    strVersionName       = ""
    strEdition           = ""
    intSPLevel           = 0
    bolIsDefaultInstance = True
    bolIsDomainMember    = GetIsDomainMember()
  End Sub


  Private Function GetIsDomainMember
    Dim objWMIService, colItems, objItem

    GetIsDomainMember = False

    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems      = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem", , 48)

    For Each objItem In colItems
      If objItem.PartOfDomain Then
        GetIsDomainMember = True
        Exit For
      End If
    Next
  End Function


  Public Property Let InstanceName(strValue)
    Dim objWshNetwork

    Set objWshNetwork = CreateObject("WScript.Network")

    If StrComp(strValue, "MSSQLSERVER", vbTextCompare) = 0 Then
      strInstanceName      = objWshNetwork.ComputerName
      bolIsDefaultInstance = True
    Else
      strInstanceName      = objWshNetwork.ComputerName & "\" & strValue
      bolIsDefaultInstance = False
    End If
  End Property

  Public Property Get InstanceName
    InstanceName = strInstanceName
  End Property


  Public Property Let InstanceExePath(strValue)
    strInstanceExePath = strValue
  End Property

  Public Property Get InstanceExePath
    InstanceExePath = strInstanceExePath
  End Property


  Public Property Let Edition(strValue)
    strEdition = strValue
  End Property

  Public Property Get Edition
    Edition = strEdition
  End Property


  Public Property Let SPLevel(intValue)
    intSPLevel = intValue
  End Property

  Public Property Get SPLevel
    SPLevel = intSPLevel
  End Property


  Public Property Let Version(strValue)
    strVersion = strValue

    Select Case Split(strVersion, ".")(0)
      Case "15" strVersionName = "SQL Server 2019"
      Case "14" strVersionName = "SQL Server 2017"
      Case "13" strVersionName = "SQL Server 2016"
      Case "12" strVersionName = "SQL Server 2014"
      Case "11" strVersionName = "SQL Server 2012"
      Case "10" Select Case Left(Split(strVersion, ".")(1), 1)
                  Case  "0" strVersionName = "SQL Server 2008"
                  Case  "5" strVersionName = "SQL Server 2008 R2"
                  Case Else strVersionName = "SQL Server " & strVersion
                End Select
      Case  "9" strVersionName = "SQL Server 2005"
      Case  "8" strVersionName = "SQL Server 2000"
      Case  "7" strVersionName = "SQL Server 7.0"
      Case Else strVersionName = "SQL Server " & strVersion
    End Select
  End Property

  Public Property Get Version
    Version = strVersion
  End Property


  Public Property Get VersionName
    VersionName = strVersionName
  End Property


  Public Property Get FullName
    FullName = "MS " & strVersionName

    If SPLevel > 0 Then
      FullName = FullName & " (SP" & SPLevel & ")"
    End If

    FullName = FullName & " " & Edition
  End Property


  Public Property Get IsDefaultInstance
    IsDefaultInstance = bolIsDefaultInstance
  End Property


  Public Property Get IsDomainMember
    IsDomainMember = bolIsDomainMember
  End Property


  Public Property Get IsValid
    Dim objFSO

    IsValid = False

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If InstanceName    = ""                   Then Exit Property
    If InstanceExePath = ""                   Then Exit Property
    If Version         = ""                   Then Exit Property
    If Edition         = ""                   Then Exit Property
    If Not objFSO.FileExists(InstanceExePath) Then Exit Property

    IsValid = True
  End Property
End Class
