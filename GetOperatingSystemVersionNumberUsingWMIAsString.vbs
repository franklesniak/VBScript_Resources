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
    ' Version: 1.0.20210115.0
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

    If TestObjectForData(intTargetWMIProviderArchitectureBitWidth) = False Then
        intReturnCode = -1
    Else
        If VarType(intTargetWMIProviderArchitectureBitWidth) <> 2 Then
            intReturnCode = -2
        End If
    End If

    If intReturnCode = 0 Then
        ' No error occurred
        On Error Resume Next
        Set objSWbemNamedValueSetTemp = CreateObject("WbemScripting.SWbemNamedValueSet")
        If Err Then
            On Error Goto 0
            Err.Clear
            intReturnCode = -3
        Else
            objSWbemNamedValueSetTemp.Add "__ProviderArchitecture", intTargetWMIProviderArchitectureBitWidth
            If Err Then
                On Error Goto 0
                Err.Clear
                intReturnCode = -4
            Else
                objSWbemNamedValueSetTemp.Add "__RequiredArchitecture", True
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
            intReturnCode = -6
        Else
            objSWbemNamedValueSetContext.Add "__ProviderArchitecture", intTargetWMIProviderArchitectureBitWidth
            If Err Then
                On Error Goto 0
                Err.Clear
                intReturnCode = -7
            Else
                objSWbemNamedValueSetContext.Add "__RequiredArchitecture", True
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intReturnCode = -8
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
    ' Version: 2.0.20210115.0
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
            If VarType(objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth) = 2 Then
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
            If VarType(objSWbemNamedValueSetContextOrIntTargetWMIProviderArchitectureBitWidth) = 2 Then
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

Function GetOperatingSystemVersionNumberUsingWMIAsString(ByRef strOperatingSystemVersion, ByVal boolExcludeRevisionNumber)
    'region FunctionMetadata ####################################################
    ' Safely obtains the operating system version number from Win32_OperatingSystem using WMI
    '
    ' Function takes two positional arguments:
    '   The first argument (strOperatingSystemVersion) will be populated with the operating
    '       system version in string format upon success
    '   The second argument (boolExcludeRevisionNumber) indicates whether the function should
    '       remove the "revision" portion of the operating system version (i.e.,
    '       major.minor.build.revision) from the operating system version string that the
    '       function returns via the first argument. Doing so can be useful because, at the
    '       time of writing, WMI does not accurately retrieve the revision number, which can be
    '       misleading or cause confusion. The valid values for boolExcludeRevisionNumber are:
    '           True = return just the major.minor.build portions of the operating system
    '               version; if WMI provides a revision number, it is removed.
    '           False = return the version number exactly as it was provided by WMI
    '           Null = same as False
    '
    ' The function returns 0 or a positive number if the operating system version number was
    ' retrieved successfully. A negative number is returned if the operating system version
    ' number was not retrieved successfully.
    '
    ' Example:
    '   intReturnCode = GetOperatingSystemVersionNumberUsingWMIAsString(strOperatingSystemVersion, True)
    '   If intReturnCode = 0 Then
    '       ' strOperatingSystemVersion is populated with the operating system version number
    '       ' in string format. If applicable, the revision portion of the version number was
    '       ' trimmed.
    '   Else
    '       ' The operating system version number could not be retrieved via WMI. Usually this
    '       ' occurs because something in the operating system blocked the creation of the WMI
    '       ' object, something is wrong with the WMI database, or in the case of Windows 95,
    '       ' Windows 98, or Windows NT 4.0, it is likely that WMI is not installed.
    '   End If
    '
    ' Version: 1.0.20210120.0
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

    'region Acknowledgements ####################################################
    ' User "Shem Sargent" on Super User, who provided sample code for augmenting WMI-based
    ' version numbers with their revision number:
    ' https://superuser.com/a/1160428/334370
    'endregion Acknowledgements ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim objWMI
    Dim intReturnCode
    Dim colOperatingSystems
    Dim objOperatingSystem
    Dim strWorkingOSVersion
    Dim arrVersionNumber

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 1

    intReturnCode = ConnectLocalWMINamespace(objWMI, Null, Null)
    If intReturnCode <> 0 Then
        intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
    Else
        intReturnMultiplier = intReturnMultiplier * 16
        On Error Resume Next
        Set colOperatingSystems = objWMI.ExecQuery("Select Version From Win32_OperatingSystem")
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
        Else
            For Each objOperatingSystem in colOperatingSystems
                strWorkingOSVersion = objOperatingSystem.Version
            Next
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
            Else
                On Error Goto 0
                If TestObjectForData(strWorkingOSVersion) = False Then
                    intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        If boolExcludeRevisionNumber = True Then
            arrVersionNumber = Split(strWorkingOSVersion, ".")
            If UBound(arrVersionNumber) >= 3 Then
                ' Revision portion of version number is present
                strWorkingOSVersion = arrVersionNumber(0) & "." & arrVersionNumber(1) & "." & arrVersionNumber(2)
            End If
        End If
        strOperatingSystemVersion = strWorkingOSVersion
    End If
    
    GetOperatingSystemVersionNumberUsingWMIAsString = intFunctionReturn
End Function
