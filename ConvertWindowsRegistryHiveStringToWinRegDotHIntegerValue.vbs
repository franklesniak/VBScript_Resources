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

Function ConvertWindowsRegistryHiveStringToWinRegDotHIntegerValue(ByRef lngRegistryHive, ByVal strRegistryHiveName)
    'region FunctionMetadata ####################################################
    ' Safely takes a string that contains a registry hive and converts it to a numerical value
    ' compatible with Windows subsystems that require the integer values specified in WinReg.h
    '
    ' Function takes two positional arguments:
    '   The first argument (lngRegistryHive) is set to a 32-bit integer upon success:
    '           "HKCU" Or "HKEY_CURRENT_USER" specified in the second argument
    '               (strRegistryHiveName): this argument (lngRegistryHive) will be set to
    '               &H80000001 (hex) = 2147483649
    '           "HKLM" Or "HKEY_LOCAL_MACHINE" specified in the second argument
    '               (strRegistryHiveName): this argument (lngRegistryHive) will be set to
    '               &H80000002 (hex) = 2147483650
    '           "HKDU" Or "HKEY_DEFAULT_USER" specified in the second argument
    '               (strRegistryHiveName): this argument (lngRegistryHive) will be set to &H4D2
    '               (hex) = 1234
    '               NOTE: This is a fake registry hive designation created by the function
    '               author to handle automatic mounting and unmounting of the default user
    '               profile's HKCU registry hive. This value should not be passed to Windows
    '               system calls that use WinReg.h values as it will result in an error.
    '               NOTE 2: If "HKDU" Or "HKEY_DEFAULT_USER" was specified in the second
    '               argument (strRegistryHiveName), the function will return 1
    '           "HKCR" Or "HKEY_CLASSES_ROOT" specified in the second argument
    '               (strRegistryHiveName): this argument (lngRegistryHive) will be set to
    '               &H80000000 (hex) = 2147483648
    '           "HKU" Or "HKEY_USERS" specified in the second argument (strRegistryHiveName):
    '               this argument (lngRegistryHive) will be set to &H80000003 (hex) =
    '               2147483651
    '           "HKCC" Or "HKEY_CURRENT_CONFIG" specified in the second argument
    '               (strRegistryHiveName): this argument (lngRegistryHive) will be set to
    '               &H80000005 (hex) = 2147483653
    '           "HKDD" Or "HKEY_DYN_DATA" specified in the second argument
    '               (strRegistryHiveName): this argument (lngRegistryHive) will be set to
    '               &H80000006 (hex) = 2147483654
    '           "HKPD" Or "HKEY_PERFORMANCE_DATA" specified in the second argument
    '               (strRegistryHiveName): this argument (lngRegistryHive) will be set to
    '               &H80000004 (hex) = 2147483652
    '   The second argument (strRegistryHiveName) is a string containing one of the following
    '       values:
    '           "HKCU" or "HKEY_CURRENT_USER"
    '           "HKLM" or "HKEY_LOCAL_MACHINE"
    '           "HKDU" or "HKEY_DEFAULT_USER" - a "fake" registry hive that references the
    '               default user profile's HKCU registry hive. This designation was created by
    '               the function author to facilitate downstream processing, i.e., automatic
    '               mounting and dismounting of the default user profile's HKCU registry hive.
    '           "HKCR" or "HKEY_CLASSES_ROOT" - a "fake" registry hive that represents a
    '               joining of HKCU\Software\Classes and HKLM\Software\Classes. Per Wikipedia,
    '               if a given value exists in both HKCU\Software\Classes and
    '               HKLM\Software\Classes, the one in HKCU\Software\Classes takes precedence.
    '           "HKU" or "HKEY_USERS"
    '           "HKCC" or "HKEY_CURRENT_CONFIG" - a "fake" registry hive that serves as an
    '               alias for "HKLM\SYSTEM\CurrentControlSet\Hardware Profiles\Current".
    '           "HKDD" or "HKEY_DYN_DATA" - only present in Windows 95, 98, and ME.
    '           "HKPD" or "HKEY_PERFORMANCE_DATA" - a "fake" registry hive that exposes
    '               performance information; not persistent/not stored on disk.
    '
    ' The function returns 0 if the registry hive was successfully converted to the equivalent
    ' integer value specified in WinReg.h. The function returns 1 if the registry hive
    ' specified was the fake "HKDU" / "HKEY_DEFAULT_USER" hive created by the function author
    ' to facilitate downstream processing and automatic mounting/unmounting of the default user
    ' profile's HKCU registry hive. A negative number is returned if the registry hive name was
    ' invalid and could not be converted.
    '
    ' Example:
    '   intReturnCode = ConvertWindowsRegistryHiveStringToWinRegDotHIntegerValue(lngRegistryHive, "HKEY_LOCAL_MACHINE")
    '   If intReturnCode >= 0 Then
    '       ' Conversion completed successfully
    '       ' lngRegistryHive equals 2147483650
    '   End If
    '
    ' Version: 1.0.20210123.0
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
    ' Microsoft, who published the list of Windows Registry hives present in WinReg.h on the
    ' following page:
    ' https://docs.microsoft.com/en-us/previous-versions/windows/desktop/regprov/enumkey-method-in-class-stdregprov
    '
    ' Stack Overflow user "TheMadTechnician", who listed additional values from WinReg.h:
    ' https://stackoverflow.com/a/24892338/2134110
    'endregion Acknowledgements ####################################################

    Dim intFunctionReturn
    Dim intVariableType
    Dim lngRegistryHiveStaging

    Err.Clear

    intFunctionReturn = 0

    If TestObjectForData(strRegistryHiveName) = False Then
        intFunctionReturn = -1
    Else
        On Error Resume Next
        intVariableType = VarType(strRegistryHiveName)
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = -2
        Else
            On Error Goto 0
            If intVariableType <> 8 Then
                'Was not a string
                intFunctionReturn = -3
            Else
                On Error Resume Next
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        ' strRegistryHiveName is a string and contains data
        Select Case UCase(strRegistryHiveName)
            Case "HKCU"
                lngRegistryHiveStaging = &H80000001
            Case "HKEY_CURRENT_USER"
                lngRegistryHiveStaging = &H80000001
            Case "HKLM"
                lngRegistryHiveStaging = &H80000002
            Case "HKEY_LOCAL_MACHINE"
                lngRegistryHiveStaging = &H80000002
            Case "HKDU"
                lngRegistryHiveStaging = 1234
                intFunctionReturn = 1
            Case "HKEY_DEFAULT_USER"
                lngRegistryHiveStaging = 1234
                intFunctionReturn = 1
            Case "HKCR"
                lngRegistryHiveStaging = &H80000000
            Case "HKEY_CLASSES_ROOT"
                lngRegistryHiveStaging = &H80000000
            Case "HKU"
                lngRegistryHiveStaging = &H80000003
            Case "HKEY_USERS"
                lngRegistryHiveStaging = &H80000003
            Case "HKCC"
                lngRegistryHiveStaging = &H80000005
            Case "HKEY_CURRENT_CONFIG"
                lngRegistryHiveStaging = &H80000005
            Case "HKDD"
                lngRegistryHiveStaging = &H80000006
            Case "HKEY_DYN_DATA"
                lngRegistryHiveStaging = &H80000006
            Case "HKPD"
                lngRegistryHiveStaging = &H80000004
            Case "HKEY_PERFORMANCE_DATA"
                lngRegistryHiveStaging = &H80000004
            Case Else
                intFunctionReturn = -4
        End Select
    End If

    If intFunctionReturn >= 0 Then
        lngRegistryHive = lngRegistryHiveStaging
    End If

    ConvertWindowsRegistryHiveStringToWinRegDotHIntegerValue = intFunctionReturn
End Function
