Function TestObjectForData(ByVal objToCheck)
    'region FunctionMetadata ####################################################
    ' Checks an object or variable to see if it "has data".
    ' If any of the following are true, then objToCheck is regarded as NOT having data:
    '   VarType(objToCheck) = 0
    '   VarType(objToCheck) = 1
    '   objToCheck Is Nothing
    '   IsEmpty(objToCheck)
    '   IsNull(objToCheck)
    '   objToCheck = vbNullString (or "")
    '   IsArray(objToCheck) = True And UBound(objToCheck) throws an error
    '   IsArray(objToCheck) = True And UBound(objToCheck) < 0
    ' In any of these cases, the function returns False. Otherwise, it returns True.
    '
    ' Version: 1.1.20210115.0
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
    ' at https://github.com/franklesniak/Test_Object_For_Data
    'endregion DownloadLocationNotice ####################################################

    'region Acknowledgements ####################################################
    ' Thanks to Scott Dexter for writing the article "Empty Nothing And Null How Do You Feel
    ' Today", which inspired me to create this function. https://evolt.org/node/346
    '
    ' Thanks also to "RhinoScript" for the article "Testing for Empty Arrays" for providing
    ' guidance for how to test for the empty array condition in VBScript.
    ' https://wiki.mcneel.com/developer/scriptsamples/emptyarray
    '
    ' Thanks also "iamresearcher" who posted this and inspired the test case for vbNullString:
    ' https://www.vbforums.com/showthread.php?684799-The-Differences-among-Empty-Nothing-vbNull-vbNullChar-vbNullString-and-the-Zero-L
    'endregion Acknowledgements ####################################################

    Dim boolTestResult
    Dim boolFunctionReturn
    Dim intArrayUBound

    Err.Clear

    boolFunctionReturn = True

    'Check VarType(objToCheck) = 0
    On Error Resume Next
    boolTestResult = (VarType(objToCheck) = 0)
    If Err Then
        'Error occurred
        Err.Clear
        On Error Goto 0
    Else
        'No Error
        On Error Goto 0
        If boolTestResult = True Then
            'vbEmpty
            boolFunctionReturn = False
        End If
    End If

    'Check VarType(objToCheck) = 1
    On Error Resume Next
    boolTestResult = (VarType(objToCheck) = 1)
    If Err Then
        'Error occurred
        Err.Clear
        On Error Goto 0
    Else
        'No Error
        On Error Goto 0
        If boolTestResult = True Then
            'vbNull
            boolFunctionReturn = False
        End If
    End If

    'Check to see if objToCheck Is Nothing
    If boolFunctionReturn = True Then
        On Error Resume Next
        boolTestResult = (objToCheck Is Nothing)
        If Err Then
            'Error occurred
            Err.Clear
            On Error Goto 0
        Else
            'No Error
            On Error Goto 0
            If boolTestResult = True Then
                'No data
                boolFunctionReturn = False
            End If
        End If
    End If

    'Check IsEmpty(objToCheck)
    If boolFunctionReturn = True Then
        On Error Resume Next
        boolTestResult = IsEmpty(objToCheck)
        If Err Then
            'Error occurred
            Err.Clear
            On Error Goto 0
        Else
            'No Error
            On Error Goto 0
            If boolTestResult = True Then
                'No data
                boolFunctionReturn = False
            End If
        End If
    End If

    'Check IsNull(objToCheck)
    If boolFunctionReturn = True Then
        On Error Resume Next
        boolTestResult = IsNull(objToCheck)
        If Err Then
            'Error occurred
            Err.Clear
            On Error Goto 0
        Else
            'No Error
            On Error Goto 0
            If boolTestResult = True Then
                'No data
                boolFunctionReturn = False
            End If
        End If
    End If
    
    'Check objToCheck = vbNullString
    If boolFunctionReturn = True Then
        On Error Resume Next
        boolTestResult = (objToCheck = vbNullString)
        If Err Then
            'Error occurred
            Err.Clear
            On Error Goto 0
        Else
            'No Error
            On Error Goto 0
            If boolTestResult = True Then
                'No data
                boolFunctionReturn = False
            End If
        End If
    End If

    If boolFunctionReturn = True Then
        On Error Resume Next
        boolTestResult = IsArray(objToCheck)
        If Err Then
            'Error occurred
            Err.Clear
            On Error Goto 0
            boolTestResult = False
        Else
            'No Error
            On Error Goto 0
        End If
        If boolTestResult = True Then
            ' objToCheck is an array
            On Error Resume Next
            intArrayUBound = UBound(objToCheck)
            If Err Then
                'Undimensioned array
                Err.Clear
                On Error Goto 0
                intArrayUBound = -1
            Else
                On Error Goto 0
            End If
            If intArrayUBound < 0 Then
                boolFunctionReturn = False
            End If
        End If
    End If

    TestObjectForData = boolFunctionReturn
End Function

Function TestObjectIsAnyTypeOfInteger(ByRef objToTest)
    'region FunctionMetadata ####################################################
    ' Safely determines if the specified object is an integer (of any kind)
    '
    ' Function takes one positional argument (objToTest), which is the object to be tested to
    '   determine if it is an integer number.
    '
    ' The function returns boolean True if the specified object is an integer number, boolean
    ' False otherwise
    '
    ' Example 1:
    '   objToTest = "12345"
    '   boolResult = TestObjectIsAnyTypeOfInteger(objToTest)
    '   ' boolResult is equal to False
    '
    ' Example 2:
    '   objToTest = 0
    '   boolResult = TestObjectIsAnyTypeOfInteger(objToTest)
    '   ' boolResult is equal to True
    '
    ' Example 3:
    '   objToTest = 12345
    '   boolResult = TestObjectIsAnyTypeOfInteger(objToTest)
    '   ' boolResult is equal to True
    '
    ' Example 4:
    '   objToTest = 12345.678
    '   boolResult = TestObjectIsAnyTypeOfInteger(objToTest)
    '   ' boolResult is equal to False
    '
    ' Example 5:
    '   objToTest = True
    '   boolResult = TestObjectIsAnyTypeOfInteger(objToTest)
    '   ' boolResult is equal to False
    '
    ' Version: 1.0.20210220.0
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

    Dim boolFunctionReturn
    Dim boolTest
    Dim intVarType

    If TestObjectForData(objToTest) = False Then
        boolFunctionReturn = False
    Else
        ' objToTest has data
        On Error Resume Next
        intVarType = VarType(objToTest)
        If Err Then
            On Error Goto 0
            Err.Clear
            boolFunctionReturn = False
        Else
            boolTest = (intVarType <> 2 And intVarType <> 3)
            If Err Then
                On Error Goto 0
                Err.Clear
                boolFunctionReturn = False
            Else
                On Error Goto 0
                If boolTest = True Then
                    ' VarType(objToTest) <> 2 And VarType(objToTest) <> 3
                    boolFunctionReturn = False
                Else
                    ' VarType(objToTest) = 2 Or VarType(objToTest) = 3
                    boolFunctionReturn = True
                End If
            End If
        End If
    End If

    TestObjectIsStringContainingData = boolFunctionReturn
End Function

Function NewWMIBitWidthContext(ByRef objSWbemNamedValueSetContext, ByVal intTargetWMIProviderArchitectureBitWidth)
    'region FunctionMetadata ####################################################
    ' Safely creates a SWbemNamedValueSet object for use with setting the bit-width "context"
    ' when connecting to or working with WMI.
    '
    ' Function takes three positional arguments:
    '   The first argument (objSWbemNamedValueSetContext) will be populated with the
    '       SWbemNamedValueSet (WMI context) object upon successful creation and configuration.
    '   The second argument (intTargetWMIProviderArchitectureBitWidth) specifies a target bit
    '       width "context" to use when opening WMI. For example, supplying 32 or 64 will force
    '       a respective 32- or 64-bit context when opening the WMI connection. This feature is
    '       commonly used when connecting to the "root\default" WMI namespace and then using
    '       the StdRegProv class to connect to the Windows registry.
    '
    ' The function returns 0 if the SWbemNamedValueSet (WMI context) object
    '       objSWbemNamedValueSetContext was created successfully; a negative number otherwise.
    '
    ' Example:
    '   intReturnCode = NewWMIBitWidthContext(objWMIContext, 32)
    '   If intReturnCode = 0 Then
    '       ' objWMIContext is initialized and configured to instruct WMI to use a 32-bit
    '       ' context.
    '   End If
    '
    ' Version: 1.1.20210220.0
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

    Dim intReturnCode
    Dim objSWbemNamedValueSetTemp

    Err.Clear

    intReturnCode = 0

    If TestObjectIsAnyTypeOfInteger(intTargetWMIProviderArchitectureBitWidth) = False Then
        intReturnCode = -1
    Else
        On Error Resume Next
        Set objSWbemNamedValueSetTemp = CreateObject("WbemScripting.SWbemNamedValueSet")
        If Err Then
            On Error Goto 0
            Err.Clear
            intReturnCode = -2
        Else
            objSWbemNamedValueSetTemp.Add "__ProviderArchitecture", intTargetWMIProviderArchitectureBitWidth
            If Err Then
                On Error Goto 0
                Err.Clear
                intReturnCode = -3
            Else
                objSWbemNamedValueSetTemp.Add "__RequiredArchitecture", True
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intReturnCode = -4
                Else
                    On Error Goto 0
                End If
            End If
        End If
    End If

    If intReturnCode = 0 Then
        ' No error occurred
        ' At this point, we've only configured a temporary variable; we still need to configure
        ' objSWbemNamedValueSetContext:
        Set objSWbemNamedValueSetTemp = Nothing
        On Error Resume Next
        Set objSWbemNamedValueSetContext = CreateObject("WbemScripting.SWbemNamedValueSet")
        If Err Then
            On Error Goto 0
            Err.Clear
            intReturnCode = -5
        Else
            objSWbemNamedValueSetContext.Add "__ProviderArchitecture", intTargetWMIProviderArchitectureBitWidth
            If Err Then
                On Error Goto 0
                Err.Clear
                intReturnCode = -6
            Else
                objSWbemNamedValueSetContext.Add "__RequiredArchitecture", True
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intReturnCode = -7
                Else
                    On Error Goto 0
                End If
            End If
        End If
    End If

    NewWMIBitWidthContext = intReturnCode
End Function

Function ConnectLocalWMINamespace(ByRef objSWbemServicesWMINamespace, ByVal strTargetWMINamespace, ByVal objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth)
    'region FunctionMetadata ####################################################
    ' Safely creates a SWbemServices object with a connection to the specified WMI namespace on
    ' the local computer.
    '
    ' Function takes three positional arguments:
    '   The first argument (objSWbemServicesWMINamespace) will be populated with the
    '       SWbemServices (WMI connection) object upon successful connection.
    '   The second argument (strTargetWMINamespace) specifies the namespace target to which
    '       this function will connect. If vbNullString ("") or Null is passed, the function
    '       defaults to "root\cimv2", which is the most commonly-used WMI namespace.
    '   The third argument
    '       (objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth) specifies
    '       either a SWbemNamedValueSet that sets the required bit-width to use when opening
    '       the WMI connection, **or** it specifies an integer target bit width "context" to
    '       use when opening WMI. For example, supplying 32 or 64 will force a respective 32-
    '       or 64-bit context when opening the WMI connection. Generally, when using this
    '       function, it is recommended to use SWbemNamedValueSet instead of an integer. This
    '       feature is commonly used when connecting to the "root\default" WMI namespace and
    '       then using the StdRegProv class to connect to the Windows registry. If Null is
    '       passed, the function defaults to the context supplied by the VBScript process that
    '       is running this script.
    '
    ' The function returns 0 if the SWbemServices (WMI connection) object
    '       objSWbemServicesWMINamespace was created successfully; a negative number otherwise.
    '
    ' Example 1:
    '   intReturnCode = ConnectLocalWMINamespace(objWMI, Null, Null)
    '   If intReturnCode = 0 Then
    '       ' objWMI is initialized and connected to the root\CIMv2 namespace
    '       Set colOS = objWMI.InstancesOf("Win32_OperatingSystem")
    '       For Each objOS in colOS
    '           WScript.Echo(objOS.Caption)
    '       Next
    '   End If
    '
    ' Example 2:
    '   Const HKEY_CLASSES_ROOT     = &H80000000
    '   Const HKEY_CURRENT_USER     = &H80000001
    '   Const HKEY_LOCAL_MACHINE    = &H80000002
    '   Const HKEY_USERS            = &H80000003
    '   intReturnCode = NewWMIBitWidthContext(objWMIContext, 32)
    '   If intReturnCode = 0 Then
    '       intReturnCode = ConnectLocalWMINamespace(objWMI, "root\default", objWMIContext)
    '       If intReturnCode = 0 Then
    '           ' objWMI is initialized and connected to the root\default namespace
    '           ' Create the StdRegProv:
    '           Set objStdRegProv = objWMI.Get("StdRegProv")
    '           ' Create a registry key in the 32-bit process context:
    '           Set objInParams = objStdRegProv.Methods_("CreateKey").Inparameters
    '           objInParams.hDefKey = HKEY_CURRENT_USER
    '           objInParams.sSubKeyName = "SOFTWARE\West Monroe Partners\Temp"
    '           Set objOutParams = objStdRegProv.ExecMethod_("CreateKey",objInParams,,objWMIContext)
    '           intReturnCode = objOutParams.ReturnValue
    '       End If
    '   End If
    '
    ' Example 3:
    '   intReturnCode = ConnectLocalWMINamespace(objWMI, Null, 64)
    '   If intReturnCode = 0 Then
    '       ' objWMI is initialized and connected to the root\cimv2 namespace
    '       Set colWinSATs = objWMI.ExecQuery("Select * From Win32_WinSAT")
    '       For Each objWinSAT in colWinSATs
    '           WScript.Echo(objWinSAT.WinSATAssessmentState)
    '       Next
    '   End If
    '
    ' Version: 2.2.20210220.0
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
    Dim intReturnCode
    Dim strEffectiveNamespace
    Dim objSWbemLocator
    Dim objSWbemNamedValueSetContext
    Dim objSWbemServicesTemp

    Const wbemImpersonationLevelImpersonate = 3
    strEffectiveComputerName = "."

    Err.Clear

    intReturnCode = 0
    
    If TestObjectForData(strTargetWMINamespace) = False Then
        strEffectiveNamespace = "root\cimv2"
    Else
        strEffectiveNamespace = strTargetWMINamespace
    End If

    On Error Resume Next
    Set objSWbemLocator = CreateObject("Wbemscripting.SWbemLocator")
    If Err Then
        On Error Goto 0
        Err.Clear
        intReturnCode = -1
    Else
        On Error Goto 0
    End If

    If intReturnCode = 0 Then
        ' No error occurred
        If TestObjectForData(objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth) = True Then
            ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth parameter
            ' was supplied
            If TestObjectIsAnyTypeOfInteger(objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth) = True Then
                ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth is an
                ' integer
                On Error Resume Next
                Set objSWbemNamedValueSetContext = CreateObject("WbemScripting.SWbemNamedValueSet")
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intReturnCode = -2
                Else
                    objSWbemNamedValueSetContext.Add "__ProviderArchitecture", objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        intReturnCode = -3
                    Else
                        objSWbemNamedValueSetContext.Add "__RequiredArchitecture", True
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intReturnCode = -4
                        Else
                            Set objSWbemServicesTemp = objSWbemLocator.ConnectServer(strEffectiveComputerName, strEffectiveNamespace,,,,,,objSWbemNamedValueSetContext)
                            If Err Then
                                On Error Goto 0
                                Err.Clear
                                intReturnCode = -5
                            Else
                                On Error Goto 0
                            End If
                        End If
                    End If
                End If
            Else
                ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth is not
                ' an integer; it is probably a SWbemNamedValueSet
                On Error Resume Next
                Set objSWbemServicesTemp = objSWbemLocator.ConnectServer(strEffectiveComputerName, strEffectiveNamespace,,,,,,objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intReturnCode = -6
                Else
                    On Error Goto 0
                End If
            End If
        Else
            ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth parameter
            ' was not supplied
            On Error Resume Next
            Set objSWbemServicesTemp = objSWbemLocator.ConnectServer(strEffectiveComputerName, strEffectiveNamespace)
            If Err Then
                On Error Goto 0
                Err.Clear
                intReturnCode = -7
            Else
                On Error Goto 0
            End If
        End If

        If intReturnCode = 0 Then
            ' No error occurred
            On Error Resume Next
            objSWbemServicesTemp.Security_.ImpersonationLevel = wbemImpersonationLevelImpersonate
            If Err Then
                On Error Goto 0
                Err.Clear
                intReturnCode = -8
            Else
                On Error Goto 0
            End If
        End If
    End If

    If intReturnCode = 0 Then
        ' No error occurred
        ' We fully connected to WMI, but did so with a "dummy" object...
        ' ... so, let's connect using the real object
        Set objSWbemServicesTemp = Nothing
        If TestObjectForData(objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth) = True Then
            ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth parameter
            ' was supplied
            If TestObjectIsAnyTypeOfInteger(objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth) = True Then
                ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth is an
                ' integer
                ' objSWbemNamedValueSetContext already constructed
                On Error Resume Next
                Set objSWbemServicesWMINamespace = objSWbemLocator.ConnectServer(strEffectiveComputerName, strEffectiveNamespace,,,,,,objSWbemNamedValueSetContext)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intReturnCode = -9
                Else
                    On Error Goto 0
                End If
            Else
                ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth is not
                ' an integer; it is probably a SWbemNamedValueSet
                On Error Resume Next
                Set objSWbemServicesWMINamespace = objSWbemLocator.ConnectServer(strEffectiveComputerName, strEffectiveNamespace,,,,,,objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intReturnCode = -10
                Else
                    On Error Goto 0
                End If
            End If
        Else
            ' objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth parameter
            ' was not supplied
            On Error Resume Next
            Set objSWbemServicesWMINamespace = objSWbemLocator.ConnectServer(strEffectiveComputerName, strEffectiveNamespace)
            If Err Then
                On Error Goto 0
                Err.Clear
                intReturnCode = -11
            Else
                On Error Goto 0
            End If
        End If
        If intReturnCode = 0 Then
            ' No error occurred
            On Error Resume Next
            objSWbemServicesWMINamespace.Security_.ImpersonationLevel = wbemImpersonationLevelImpersonate
            If Err Then
                On Error Goto 0
                Err.Clear
                intReturnCode = -12
            Else
                On Error Goto 0
            End If
        End If
    End If

    ConnectLocalWMINamespace = intReturnCode
End Function
