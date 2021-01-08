Function CreateWMINamespaceObject(ByRef objWMINamespaceConnection, ByVal strTargetComputerName, ByVal strTargetNamespace)
    'region FunctionMetadata ####################################################
    ' Safely creates a WMI object with a connection to the specified namespace.
    '
    ' Function takes three positional arguments:
    '   The first argument (objWMINamespaceConnection) will be populated with the WMI
    '       connection object upon successful connection.
    '   The second argument (strTargetComputerName) can be the name of a target computer (for
    '       remote WMI connections). For connections to the local computer, it should be "." or
    '       vbNullString (""), or Null.
    '   The third argument (strTargetNamespace) specifies the namespace target to which this
    '       function will connect. If vbNullString ("") or Null is passed, the function
    '       defaults to "root\cimv2", which is the most commonly-used WMI namespace.
    '
    ' The function returns 0 if the WMI object was created successfully; -1 otherwise.
    '
    ' Example:
    '   intReturnCode = CreateWMINamespaceObject(objWMI, Null, Null)
    '   If intReturnCode = 0 Then
    '       'objWMI is initialized and connected to the root\CIMv2 namespace
    '   End If
    '
    ' Version: 1.0.20210107.0
    'endregion FunctionMetadata ####################################################

    'region License ####################################################
    ' Copyright 2021 Frank Lesniak
    '
    ' Permission is hereby granted, free of charge, to any person obtaining a copy of this
    ' software and associated documentation files (the "Software"), to deal in the Software
    ' without restriction, including without limitation the rights to use, copy, modify, merge,
    ' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
    ' persons to whom the Software is furnished to do so, subject to the following conditions:
    '
    ' The above copyright notice and this permission notice shall be included in all copies or
    ' substantial portions of the Software.
    '
    ' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
    ' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
    ' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
    ' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
    ' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
    ' DEALINGS IN THE SOFTWARE.
    'endregion License ####################################################

    'region DownloadLocationNotice ####################################################
    ' The most up-to-date version of this script can be found on the author's GitHub repository
    ' at https://github.com/franklesniak/VBScript_Resources
    'endregion DownloadLocationNotice ####################################################

    Dim strEffectiveComputerName
    Dim strEffectiveNamespace
    Dim intReturnCode
    Dim objWMI
    
    If TestObjectForData(strTargetComputerName) = False Then
        strEffectiveComputerName = "."
    Else
        strEffectiveComputerName = strTargetComputerName
    End If

    If TestObjectForData(strTargetNamespace) = False Then
        strEffectiveNamespace = "root\cimv2"
    Else
        strEffectiveNamespace = strTargetNamespace
    End If

    intReturnCode = 0

    On Error Resume Next
    Set objWMI = GetObject("winmgmts:\\" + strEffectiveComputerName + "\" + strEffectiveNamespace)
    If Err Then
        On Error Goto 0
        Err.Clear
        intReturnCode = -1
    Else
        On Error Goto 0
        Set objWMI = Nothing
        Set objWMINamespaceConnection = GetObject("winmgmts:\\" + strEffectiveComputerName + "\" + strEffectiveNamespace)
    End If

    CreateWMINamespaceObject = intReturnCode
End Function
