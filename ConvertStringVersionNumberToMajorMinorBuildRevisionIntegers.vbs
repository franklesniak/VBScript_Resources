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
