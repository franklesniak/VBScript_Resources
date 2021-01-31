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

Function ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(ByRef lngMajor, ByRef lngMinor, ByRef lngBuild, ByRef lngRevision, ByVal strVersionNumber)
    'region FunctionMetadata ####################################################
    ' Safely takes a string that contains a version number and converts it to a series of four
    ' integers representing the major, minor, build, and revision portions of the version
    ' string.
    '
    ' Function takes five positional arguments:
    '   The first argument (lngMajor) is set to a 32-bit integer upon success with the major
    '       portion of the version number specified in the fifth argument.
    '   The second argument (lngMinor) is set to a 32-bit integer upon success with the minor
    '       portion of the version number specified in the fifth argument.
    '   The third argument (lngBuild) is set to a 32-bit integer upon success with the build
    '       portion of the version number specified in the fifth argument.
    '   The fourth argument (lngRevision) is set to a 32-bit integer upon success with the
    '       revision portion of the version number specified in the fifth argument.
    '   The fifth argument (strVersionNumber) contains the version number in string format to
    '       be converted. The version number should be in "major.minor.build.revision",
    '       "major.minor.build", or "major.minor" format.
    '
    ' The function returns 0 if the version number was successfully converted to its integer
    ' components. A negative number is returned if the version number could not be converted.
    '
    ' Example:
    '   intReturnCode = ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngMajor, lngMinor, lngBuild, lngRevision, "6.1.7601")
    '   If intReturnCode = 0 Then
    '       ' Conversion completed successfully
    '       ' lngMajor equals 6
    '       ' lngMinor equals 1
    '       ' lngBuild equals 7601
    '       ' lngRevision equals -1 (was not specified)
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

    Dim intFunctionReturn
    Dim arrVersion
    Dim intCountOfVersionSections
    Dim boolVersionSectionCountTest
    Dim lngTempMajor
    Dim lngTempMinor
    Dim lngTempBuild
    Dim lngTempRevision

    Err.Clear

    intFunctionReturn = 0

    If TestObjectForData(strVersionNumber) = False Then
        ' No data was passed to function
        intFunctionReturn = -1
    Else
        On Error Resume Next
        arrVersion = Split(strVersionNumber, ".")
        If Err Then
            Err.Clear
            On Error Goto 0
            ' Object passed to function was not a string, or an error occurred splitting
            ' the string
            intFunctionReturn = -2
        Else
            intCountOfVersionSections = UBound(arrVersion)
            If Err Then
                Err.Clear
                On Error Goto 0
                ' Something went wrong reading the upper boundary of the array resulting
                ' from the Split() function
                intFunctionReturn = -3
            Else
                boolVersionSectionCountTest = (intCountOfVersionSections > 3) Or (intCountOfVersionSections < 1)
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    ' Something went wrong comparing the upper boundary to an integer
                    intFunctionReturn = -4
                Else
                    On Error Goto 0
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        If boolVersionSectionCountTest = True Then
            ' Less than two parts of the version string were passed (e.g., "1")
            ' or
            ' More than four parts of the version string were passed (e.g., "1.2.3.4.5")
            ' Neither is allowed here, nor the System.Version .NET analog
            intFunctionReturn = -5
        Else
            ' String appears valid so far and has 2-4 parts, e.g.:
            ' 1.2
            ' 1.2.3
            ' 1.2.3.4
            If TestObjectForData(arrVersion(0)) = False Then
                ' Blank sections of the version number are not allowed during conversion
                ' from string
                intFunctionReturn = -6
            Else
                On Error Resume Next
                lngTempMajor = CLng(arrVersion(0))
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    ' The "major" portion of the version number was not a valid long
                    ' integer
                    intFunctionReturn = -7
                Else
                    On Error Goto 0
                    If lngTempMajor < CLng(0) Then
                        ' Cannot have negative version numbers
                        intFunctionReturn = -8
                    Else
                        lngTempMinor = CLng(0)
                        lngTempBuild = CLng(0)
                        lngTempRevision = CLng(0)
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        If TestObjectForData(arrVersion(1)) = False Then
            ' Blank sections of the version number are not allowed during conversion
            ' from string
            intFunctionReturn = -9
        Else
            On Error Resume Next
            lngTempMinor = CLng(arrVersion(1))
            If Err Then
                Err.Clear
                On Error Goto 0
                ' The "minor" portion of the version number was not a valid long integer
                intFunctionReturn = -10
            Else
                On Error Goto 0
                If lngTempMinor < CLng(0) Then
                    ' Cannot have negative version numbers
                    intFunctionReturn = -11
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        If intCountOfVersionSections >= 2 Then
            ' Build portion of version should be present
            If TestObjectForData(arrVersion(2)) = False Then
                ' Blank sections of the version number are not allowed during conversion
                ' from string
                intFunctionReturn = -12
            Else
                On Error Resume Next
                lngTempBuild = CLng(arrVersion(2))
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    ' The "build" portion of the version number was not a valid long integer
                    intFunctionReturn = -13
                Else
                    On Error Goto 0
                    If lngTempBuild < CLng(0) Then
                        ' Cannot have negative version numbers
                        intFunctionReturn = -14
                    End If
                End If
            End If
        Else
            lngTempBuild = CLng(-1)
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        If intCountOfVersionSections = 3 Then
            ' Revision portion of version should be present
            If TestObjectForData(arrVersion(3)) = False Then
                ' Blank sections of the version number are not allowed during conversion
                ' from string
                intFunctionReturn = -15
            Else
                On Error Resume Next
                lngTempRevision = CLng(arrVersion(3))
                If Err Then
                    Err.Clear
                    On Error Goto 0
                    ' The "revision" portion of the version number was not a valid long integer
                    intFunctionReturn = -16
                Else
                    On Error Goto 0
                    If lngTempRevision < CLng(0) Then
                        ' Cannot have negative version numbers
                        intFunctionReturn = -17
                    End If
                End If
            End If
        Else
            lngTempRevision = CLng(-1)
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        lngMajor = lngTempMajor
        lngMinor = lngTempMinor
        lngBuild = lngTempBuild
        lngRevision = lngTempRevision
    End If

    ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers = intFunctionReturn
End Function

Function TestWindowsRegistryKeyShouldRedirectWithWinRegDotHIntegerHiveAndOSVersionNumber(ByRef boolRedirect, ByVal lngRegistryHive, ByVal strRegistryPathWithoutHive, ByVal strOSVersionNumber)
    'region FunctionMetadata ####################################################
    ' Safely determines if a registry key would redirect on 64-bit Windows installations
    ' beginning with Windows XP and Windows Server 2003, and in 32-bit ARM processor
    ' architecture Windows installations beginning with Windows RT (Windows 8). Registry
    ' redirection occurs on these operating systems when the running process's architecture
    ' does not match the processor architecture of the operating system.
    '
    ' This version of the function requires the registry hive specified using WinReg.h integer
    ' format, and the OS version supplied in string format, with at least major.minor portions
    ' of the version number.
    '
    ' Function takes four positional arguments:
    '   The first argument (boolRedirect) is populated with True or False upon successful
    '       evaluation of whether the registry key should redirect given the information
    '       supplied in the other function arguments.
    '   The second argument (lngRegistryHive) is a 32-bit integer aligned with the definitions
    '       in WinReg.h:
    '           &H4D2 (hex) = 1234 means the default user profile's HKCU registry hive.
    '               NOTE: This is a fake registry hive designation created by the function
    '               author to handle automatic mounting and unmounting of the default user
    '               profile's HKCU registry hive. This value should not be passed to Windows
    '               system calls that use WinReg.h values as it will result in an error.
    '               NOTE 2: This registry hive is processed identially to HKCU, since it would
    '               have the same behavior
    '           &H80000000 (hex) = 2147483648 means HKCR / HKEY_CLASSES_ROOT - a "fake"
    '               registry hive that represents a joining of HKCU\Software\Classes and
    '               HKLM\Software\Classes. Per Wikipedia, if a given value exists in both
    '               HKCU\Software\Classes and HKLM\Software\Classes, the one in
    '               HKCU\Software\Classes takes precedence.
    '           &H80000001 (hex) = 2147483649 = HKCU / HKEY_CURRENT_USER
    '           &H80000002 (hex) = 2147483650 means HKLM / HKEY_LOCAL_MACHINE
    '           &H80000003 (hex) = 2147483651 means HKU / HKEY_USERS
    '           &H80000004 (hex) = 2147483652 means HKPD / HKEY_PERFORMANCE_DATA - a "fake"
    '               registry hive that exposes performance information; not persistent/not
    '               stored on disk.
    '           &H80000005 (hex) = 2147483653 means HKCC / HKEY_CURRENT_CONFIG - a "fake"
    '               registry hive that serves as an alias for
    '               "HKLM\SYSTEM\CurrentControlSet\Hardware Profiles\Current".
    '           &H80000006 (hex) = 2147483654 means HKDD / HKEY_DYN_DATA - only present in
    '               Windows 95, 98, and ME.
    '   The third argument (strRegistryHiveName) is a string containing the path to the
    '       registry value to be tested, minus the registry hive. For example, if the full path
    '       to the registry key to be tested is "HKLM\SOFTWARE\Microsoft", then this argument
    '       should be "SOFTWARE\Microsoft".
    '   The fourth argument (strOSVersionNumber) is a string containing the operating system's
    '       version number, at least the major and minor portion, in the format
    '       "major.minor.build.revision", "major.minor.build", or "major.minor".
    '
    ' The function returns 0 if the registry key was evaluated successfully for redirection. A
    ' negative number is returned if the registry key could not be evaluated.
    '
    ' Example:
    '   Const HKEY_LOCAL_MACHINE = &H80000002
    '   intReturnCode = TestWindowsRegistryKeyShouldRedirectWithWinRegDotHIntegerHiveAndOSVersionNumber(boolRedirect, HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft", "6.1.7601")
    '   If intReturnCode = 0 Then
    '       ' Registry key was tested successfully
    '       ' boolRedirect is set to True because HKLM\SOFTWARE\Microsoft should redirect
    '   End If
    '
    ' Version: 1.0.20210131.0
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
    '
    ' Microsoft, who published the list of redirected registry paths on the following page:
    ' https://docs.microsoft.com/en-us/windows/win32/winprog64/shared-registry-keys
    'endregion Acknowledgements ####################################################

    Dim intFunctionReturn
    Dim intReturnCode
    Dim lngOSMajor
    Dim lngOSMinor
    Dim lngOSBuild
    Dim lngOSRevision
    Dim strWorkingRegPath
    Dim intVariableType
    Dim boolRedirectStaging
    Dim boolXPWS2003VistaOrWS2008
    Dim lngWorkingRegistryHive
    Dim strTestPath

    Const REG_PATH_SEPARATOR = "\"
    Const HKEY_DEFAULT_USER = &H4D2
    Const HKEY_CLASSES_ROOT = &H80000000
    Const HKEY_CURRENT_USER = &H80000001
    Const HKEY_LOCAL_MACHINE = &H80000002
    Const HKEY_USERS = &H80000003
    Const HKEY_PERFORMANCE_DATA = &H80000004
    Const HKEY_CURRENT_CONFIG = &H80000005
    Const HKEY_DYN_DATA = &H80000006

    Err.Clear

    intFunctionReturn = 0

    intReturnCode = ConvertStringVersionNumberToMajorMinorBuildRevisionIntegers(lngOSMajor, lngOSMinor, lngOSBuild, lngOSRevision, strOSVersionNumber)
    If intReturnCode <> 0 Then
        ' Invalid OS version number supplied
        intFunctionReturn = -1
    Else
        ' Minimally, lngOSMajor and lngOSMinor are populated
        ' lngOSBuild is either -1, or it contains the OS build number
        ' lngOSRevision is either -1, or it contains the OS revision number
        If TestObjectForData(strRegistryPathWithoutHive) = False Then
            strWorkingRegPath = ""
        Else
            On Error Resume Next
            intVariableType = VarType(strRegistryPathWithoutHive)
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = -2
            Else
                On Error Goto 0
                If intVariableType <> 8 Then
                    ' strRegistryPathWithoutHive was not a string
                    intFunctionReturn = -3
                Else
                    ' strRegistryPathWithoutHive is a string and not a blank string
                    strWorkingRegPath = strRegistryPathWithoutHive
                    If Right(strWorkingRegPath, 1) <> REG_PATH_SEPARATOR Then
                        strWorkingRegPath = strWorkingRegPath & REG_PATH_SEPARATOR
                    End If
                    If lngRegistryHive <> HKEY_DEFAULT_USER And lngRegistryHive <> HKEY_CLASSES_ROOT And lngRegistryHive <> HKEY_CURRENT_USER And lngRegistryHive <> HKEY_LOCAL_MACHINE And lngRegistryHive <> HKEY_USERS And lngRegistryHive <> HKEY_PERFORMANCE_DATA And lngRegistryHive <> HKEY_CURRENT_CONFIG And lngRegistryHive <> HKEY_DYN_DATA Then
                        ' Invalid registry hive specified
                        intFunctionReturn = -4
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred
        ' lngOSMajor and lngOSMinor are populated
        ' lngOSBuild is either -1, or it contains the OS build number
        ' lngOSRevision is either -1, or it contains the OS revision number
        ' strWorkingRegPath is either an empty string, or it is a registry path and ends in REG_PATH_SEPARATOR
        ' lngRegistryHive is specified and valid
        
        If lngOSMajor < 5 Or (lngOSMajor = 5 And lngOSMinor = 0) Then
            ' OS is Windows 2000 or older; registry redirection does not exist
            boolRedirectStaging = False
        Else
            ' OS is at least Windows XP

            ' Determine if OS is Windows XP, Windows Server 2003, Windows Vista, or Windows
            ' Server 2008; if it is, then an older set of redirection behavior applies in some
            ' cases, depending on the registry key.
            If lngOSMajor = 5 Or (lngOSMajor = 6 and lngOSMinor = 0) Then
                boolXPWS2003VistaOrWS2008 = True
            Else
                boolXPWS2003VistaOrWS2008 = False
            End If

            ' Convert registry aliases
            If lngRegistryHive = HKEY_DEFAULT_USER Then
                lngWorkingRegistryHive = HKEY_CURRENT_USER
            ElseIf lngRegistryHive = HKEY_CLASSES_ROOT Then
                lngWorkingRegistryHive = HKEY_CURRENT_USER
                strWorkingRegPath = "SOFTWARE\CLASSES\" & strWorkingRegPath
            ElseIf lngRegistryHive = HKEY_CURRENT_CONFIG Then
                lngWorkingRegistryHive = HKEY_LOCAL_MACHINE
                strWorkingRegPath = "SYSTEM\CURRENTCONTROLSET\HARDWARE PROFILES\CURRENT\" & strWorkingRegPath
            Else
                lngWorkingRegistryHive = lngRegistryHive
            End If

            strWorkingRegPath = UCase(strWorkingRegPath)

            boolRedirectStaging = False
            If lngWorkingRegistryHive = HKEY_LOCAL_MACHINE Then
                strTestPath = "SOFTWARE\"
                If Len(strWorkingRegPath) >= Len(strTestPath) Then
                    If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                        boolRedirectStaging = True
                        strTestPath = "SOFTWARE\CLASSES\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                                strTestPath = "SOFTWARE\CLASSES\CLSID\"
                                If Len(strWorkingRegPath) >= Len(strTestPath) Then
                                    If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                        'TO-DO: Check the list of redirected registry keys
                                        ' (https://docs.microsoft.com/en-us/windows/win32/winprog64/shared-registry-keys)
                                        ' for updated text from Microsoft on the
                                        ' HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID registry
                                        ' key case for Windows Server 2008, Windows Vista,
                                        ' Windows Server 2003, and Windows XP. Its text is
                                        ' currently a bit ambiguous and may mean that this key
                                        ' does not redirect under some circumstances.
                                        boolRedirectStaging = True
                                    End If
                                End If
                                strTestPath = "SOFTWARE\CLASSES\DIRECTSHOW\"
                                If Len(strWorkingRegPath) >= Len(strTestPath) Then
                                    If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                        boolRedirectStaging = True
                                    End If
                                End If
                                strTestPath = "SOFTWARE\CLASSES\HCP\"
                                If Len(strWorkingRegPath) >= Len(strTestPath) Then
                                    If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                        boolRedirectStaging = False
                                    End If
                                End If
                                strTestPath = "SOFTWARE\CLASSES\INTERFACE\"
                                If Len(strWorkingRegPath) >= Len(strTestPath) Then
                                    If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                        boolRedirectStaging = True
                                    End If
                                End If
                                strTestPath = "SOFTWARE\CLASSES\MEDIA TYPE\"
                                If Len(strWorkingRegPath) >= Len(strTestPath) Then
                                    If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                        boolRedirectStaging = True
                                    End If
                                End If
                                strTestPath = "SOFTWARE\CLASSES\MEDIAFOUNDATION\"
                                If Len(strWorkingRegPath) >= Len(strTestPath) Then
                                    If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                        boolRedirectStaging = True
                                    End If
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\CLIENTS\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\COM3\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\CRYPTOGRAPHY\CALAIS\CURRENT\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\CRYPTOGRAPHY\CALAIS\READERS\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\CRYPTOGRAPHY\SERVICES\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\CTF\SYSTEMSHARED\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\CTF\TIP\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\DFS\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\DRIVER SIGNING\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\ENTERPRISECERTIFICATES\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\EVENTSYSTEM\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\MSMQ\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\NON-DRIVER SIGNING\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\NOTEPAD\DEFAULTFONTS\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\OLE\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\RAS\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\RPC\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\SOFTWARE\MICROSOFT\SHARED TOOLS\MSINFO" ' Not a typo
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\SYSTEMCERTIFICATES\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\TERMSERVLICENSING\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\TRANSACTIONSERVER\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\APP PATHS\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\CONTROL PANEL\CURSORS\SCHEMES\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\EXPLORER\AUTOPLAYHANDLERS\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\EXPLORER\DRIVEICONS\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\EXPLORER\KINDMAP\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\GROUP POLICY\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\POLICIES\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\PREVIEWHANDLERS\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\SETUP\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\TELEPHONY\LOCATIONS\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\CONSOLE\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\FONTDPI\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\FONTLINK\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\FONTMAPPER\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\FONTS\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\FONTSUBSTITUTES\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\GRE_INITIALIZE\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\IMAGE FILE EXECUTION OPTIONS\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\LANGUAGE PACK\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If boolXPWS2003VistaOrWS2008 = False Then
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\NETWORKCARDS\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\PERFLIB\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\PORTS\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\PRINT\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\PROFILELIST\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\TIME ZONES\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\POLICIES\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = False
                            End If
                        End If
                        strTestPath = "SOFTWARE\REGISTEREDAPPLICATIONS\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                If lngOSMajor > 5 Then
                                    ' Key present starting with Windows Vista
                                    boolRedirectStaging = False
                                End If
                            End If
                        End If
                    End If
                End If
            ElseIf lngWorkingRegistryHive = HKEY_CURRENT_USER Then
                strTestPath = "SOFTWARE\CLASSES\"
                If Len(strWorkingRegPath) >= Len(strTestPath) Then
                    If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                        If boolXPWS2003VistaOrWS2008 = True Then
                            boolRedirectStaging = True
                        End If
                        strTestPath = "SOFTWARE\CLASSES\CLSID\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = True
                            End If
                        End If
                        strTestPath = "SOFTWARE\CLASSES\DIRECTSHOW\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = True
                            End If
                        End If
                        strTestPath = "SOFTWARE\CLASSES\INTERFACE\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = True
                            End If
                        End If
                        strTestPath = "SOFTWARE\CLASSES\MEDIA TYPE\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = True
                            End If
                        End If
                        strTestPath = "SOFTWARE\CLASSES\MEDIAFOUNDATION\"
                        If Len(strWorkingRegPath) >= Len(strTestPath) Then
                            If Left(strWorkingRegPath, Len(strTestPath)) = strTestPath Then
                                boolRedirectStaging = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        boolRedirect = boolRedirectStaging
    End If

    TestWindowsRegistryKeyShouldRedirectWithWinRegDotHIntegerHiveAndOSVersionNumber = intFunctionReturn
End Function
