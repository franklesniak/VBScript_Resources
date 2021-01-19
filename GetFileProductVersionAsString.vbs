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

Function GetFileProductVersionAsString(ByRef strFileProductVersion, ByVal strFilePath)
    'region FunctionMetadata ####################################################
    ' Safely obtains the "product version" of a binary file. This is the "product version"
    ' displayed in the properties of the file, details tab, when viewed from Windows Explorer.
    '
    ' Function takes three positional arguments:
    '   The first argument (strFileProductVersion) will be the string representation of the
    '       file's product version (whose path is strFilePath).
    '   The second argument (strFilePath) is the path to the file for which we want to know the
    '       product version.
    '
    ' The function returns 0 if the file's product version was retrieved successfully. A
    ' negative number is returned if the file's product version was not retrieved successfully.
    '
    ' Example:
    ' strFilePath = "C:\Windows\System32\hal.dll"
    ' intReturnCode = GetFileProductVersionAsString(strFileProductVersion, strFilePath)
    ' If intReturnCode = 0 Then
    '   ' The product version of hal.dll was retrieved successfully and is stored in
    '   ' strFileProductVersion in string format.
    ' End If
    '
    ' Note: this function requires Windows 95 with at least Internet Explorer 4.0 installed and
    ' Windows Scripting Host 2.0 or newer installed, Windows 98 with Windows Scripting Host 2.0
    ' or newer installed, Windows ME, Windows NT 4.0 with Internet Explorer 4.0 installed and
    ' Windows Scripting Host 2.0 or newer installed, Windows 2000, or newer.
    '
    ' Version: 1.0.20210119.0
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
    ' Stack Overflow user "Maputi", who provided a sample function for reading a file's
    ' Product Version and pointed me in the right direction.
    ' https://stackoverflow.com/a/2990698/2134110
    '
    ' Microsoft, for providing documentation on the Shell Object in MSDN (Jan 2003), which
    ' clarified the requirements for using the Shell Object.
    '
    ' Andrew Clinick, for his article "If It Moves, Script It" (available in the MSDN library
    ' published 2003 Jan), which tipped me off that FileSystemObject is available starting in
    ' Windows Scripting Host 2.0.
    '
    ' Jerry Lee Ford, Jr., for providing a history of VBScript and Windows Scripting Host in
    ' his book, "Microsoft WSH and VBScript Programming for the Absolute Beginner".
    '
    ' Gunter Born, for providing a history of Windows Scripting Host in his book "Microsoft
    ' Windows Script Host 2.0 Developer's Guide" that corrected some points.
    'endregion Acknowledgements ####################################################

    Dim MAXINT
    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim objFileSystemObject
    Dim strWorkingFileProductVersion
    Dim boolResult
    Dim objFSOFile
    Dim objFSOParentFolder
    Dim strParentFolderPath
    Dim arrFilePath
    Dim strFileName
    Dim objShell
    Dim objShell32GlobalFolder
    Dim objShell32GlobalFolderItem
    Dim intCounter
    Dim intHeaderNumberForProductVersion
    Dim strHeaderName

    MAXINT = (2 ^ 15) - 1

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 1

    On Error Resume Next
    Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
    If Err Then
        On Error Goto 0
        Err.Clear
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        boolResult = objFileSystemObject.FileExists(strFilePath)
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            On Error Goto 0
            If boolResult = False Then
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                On Error Resume Next
                Set objFSOFile = objFileSystemObject.GetFile(strFilePath)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                Else
                    Set objFSOParentFolder = objFSOFile.ParentFolder
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        intFunctionReturn = intFunctionReturn + (-5 * intReturnMultiplier)
                    Else
                        strParentFolderPath = objFSOParentFolder.Path
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intFunctionReturn = intFunctionReturn + (-6 * intReturnMultiplier)
                        Else
                            On Error Goto 0
                            If TestObjectForData(strParentFolderPath) = False Then
                                intFunctionReturn = intFunctionReturn + (-7 * intReturnMultiplier)
                            Else
                                If Right(strParentFolderPath, 1) <> "\" Then
                                    strParentFolderPath = strParentFolderPath & "\"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' File specified by strFilePath exists
        ' strParentFolderPath contains file's parent folder (appended with trailing backslash)
        intReturnMultiplier = intReturnMultiplier * 8
        arrFilePath = Split(LCase(strFilePath), LCase(strParentFolderPath))
        If UBound(arrFilePath) <> 1 Then
            intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
        Else
            strFileName = arrFilePath(1)
        End If
        On Error Resume Next
        Set objShell = CreateObject("Shell.Application")
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            Set objShell32GlobalFolder = objShell.Namespace(strParentFolderPath)
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                Set objShell32GlobalFolderItem = objShell32GlobalFolder.ParseName(strFileName)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                Else
                    On Error Goto 0
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' File specified by strFilePath exists
        ' strParentFolderPath contains file's parent folder (appended with trailing backslash)
        ' strFileName initialized with file name in lowercase
        ' objShell32GlobalFolder initalized to parent folder
        ' objShell32GlobalFolderItem to file
        intReturnMultiplier = intReturnMultiplier * 8
        intHeaderNumberForProductVersion = -1
        For intCounter = 0 To MAXINT
            On Error Resume Next
            strHeaderName = objShell32GlobalFolder.GetDetailsOf(objShell32GlobalFolder.Items, intCounter)
            If Err Then
                Err.Clear
            End If
            On Error Goto 0
            If TestObjectForData(strHeaderName) = True Then
                If LCase(strHeaderName) = "product version" Then
                    ' This is the correct header index
                    intHeaderNumberForProductVersion = intCounter
                    Exit For
                End If
            End If
        Next
        If intHeaderNumberForProductVersion = -1 Then
            intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
        Else
            On Error Resume Next
            strWorkingFileProductVersion = objShell32GlobalFolder.GetDetailsOf(objShell32GlobalFolderItem, intHeaderNumberForProductVersion)
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
            Else
                On Error Goto 0
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        strFileProductVersion = strWorkingFileProductVersion
    End If
    
    GetFileProductVersionAsString = intFunctionReturn
End Function
