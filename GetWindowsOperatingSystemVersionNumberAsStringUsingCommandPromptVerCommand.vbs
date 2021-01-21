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

Function GetWindowsPath(ByRef strWindowsPath)
    'region FunctionMetadata ####################################################
    ' Safely obtains the path to the Windows folder
    '
    ' Function takes one positional argument (strWindowsPath) that is populated upon success
    ' with the path to the Windows folder. The path is appended with a trailing backslash.
    '
    ' The function returns 0 if the Windows path was retrieved successfully. A negative number
    ' is returned if the Windows path was not retrieved successfully.
    '
    ' Example:
    ' intReturnCode = GetWindowsPath(strWindowsPath)
    ' If intReturnCode = 0 Then
    '   ' Windows path was retrieved successfully and stored in strWindowsPath.
    ' End If
    '
    ' Note: the technique used in this function requires Windows Scripting Host 2.0 or newer,
    ' which was included in Windows releases beginning with Windows 2000 and Windows ME. It was
    ' available as a separate download for Windows 95, 98, and NT 4.0.
    '
    ' Version: 1.0.20210118.1
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

    Dim intFunctionReturn
    Dim objFileSystemObject
    Dim objFolder
    Dim strTempFolderPath

    Err.Clear

    intFunctionReturn = 0

    On Error Resume Next
    Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
    If Err Then
        On Error Goto 0
        Err.Clear
        intFunctionReturn = -1
    Else
        Set objFolder = objFileSystemObject.GetSpecialFolder(0)
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = -2
        Else
            strTempFolderPath = objFolder.Path
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = -3
            Else
                On Error Goto 0
                If TestObjectForData(strTempFolderPath) = False Then
                    intFunctionReturn = -4
                Else
                    If Right(strTempFolderPath, 1) <> "\" Then
                        strTempFolderPath = strTempFolderPath & "\"
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        strWindowsPath = strTempFolderPath
    End If
    
    GetWindowsPath = intFunctionReturn
End Function

Function GetWindowsSystemPath(ByRef strWindowsSystemPath)
    'region FunctionMetadata ####################################################
    ' Safely obtains the path to the Windows system folder (i.e., on a VBScript process whose
    ' processor architecture matches the operating system process architecture, the Windows
    ' system folder is usually C:\Windows\System32)
    '
    ' Function takes one positional argument (strWindowsSystemPath) that is populated upon
    ' success with the path to the Windows system folder. The path is appended with a trailing
    ' backslash.
    '
    ' The function returns 0 if the Windows system path was retrieved successfully. A negative
    ' number is returned if the Windows system path was not retrieved successfully.
    '
    ' Example:
    ' intReturnCode = GetWindowsSystemPath(strWindowsSystemPath)
    ' If intReturnCode = 0 Then
    '   ' Windows system path was retrieved successfully and stored in strWindowsSystemPath.
    ' End If
    '
    ' Note: the technique used in this function requires Windows Scripting Host 2.0 or newer,
    ' which was included in Windows releases beginning with Windows 2000 and Windows ME. It was
    ' available as a separate download for Windows 95, 98, and NT 4.0.
    '
    ' Note: if the processor architecture of the VBScript process does not match the operating
    ' system's processor architecture (e.g., 32-bit Intel IA32/x86 VBScript process running on
    ' 64-bit AMD64/Intel x86-64 Windows), then the path to the Windows System folder may be
    ' automatically substituted for the Windows-on-Windows (WOW) equivalent path.
    '
    ' Version: 1.0.20210118.1
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

    Dim intFunctionReturn
    Dim objFileSystemObject
    Dim objFolder
    Dim strTempFolderPath

    Err.Clear

    intFunctionReturn = 0

    On Error Resume Next
    Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
    If Err Then
        On Error Goto 0
        Err.Clear
        intFunctionReturn = -1
    Else
        Set objFolder = objFileSystemObject.GetSpecialFolder(1)
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = -2
        Else
            strTempFolderPath = objFolder.Path
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = -3
            Else
                On Error Goto 0
                If TestObjectForData(strTempFolderPath) = False Then
                    intFunctionReturn = -4
                Else
                    If Right(strTempFolderPath, 1) <> "\" Then
                        strTempFolderPath = strTempFolderPath & "\"
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        strWindowsSystemPath = strTempFolderPath
    End If
    
    GetWindowsSystemPath = intFunctionReturn
End Function

Function GetCommandPromptPath(ByRef strCommandPromptPath)
    'region FunctionMetadata ####################################################
    ' Safely determines the path to the Windows Command Prompt (command interpreter)
    '
    ' This function takes one argument (strCommandPromptPath) that is populated upon success
    ' with the path to the Command Prompt executable.
    '
    ' The function returns 0 or a positive number if the path to the Command Prompt was
    ' retrieved successfully; it returns a negative number if the path to the Command Prompt
    ' was not retrived successfully.
    '
    ' Example:
    ' intReturnCode = GetCommandPromptPath(strCommandPromptPath)
    ' If intReturnCode = 0 Then
    '   ' Path to command prompt executable was retrieved successfully and stored in
    '   ' strCommandPromptPath.
    ' End If
    '
    ' Version: 1.0.20210118.0
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
    ' Microsoft, for including in the MSDN Library Jan 2003 information on the nuiances in
    ' accessing environment variables on pre-Windows 2000 and Windows ME-and-prior operating
    ' systems (namely that VBScript in Windows 9x can only access per-process environment
    ' variables)
    ' (link unavailable, check Internet Archive for source)
    'endregion Acknowledgements ####################################################

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim objWSHShell
    Dim objEnvironment
    Dim strWorkingCommandPromptPath
    Dim intReturnCode
    Dim strWindowsSystemPath
    Dim objFileSystemObject
    Dim boolResult
    Dim strWindowsPath

    Err.Clear
    
    intFunctionReturn = 0
    intReturnMultiplier = 1

    ' Try shell environment variable approach
    On Error Resume Next
    Set objWSHShell = WScript.CreateObject("WScript.Shell")
    If Err Then
        Err.Clear
        On Error Goto 0
        intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
    Else
        Set objEnvironment = objWSHShell.Environment("Process")
        If Err Then
            Err.Clear
            On Error Goto 0
            intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
        Else
            strWorkingCommandPromptPath = objEnvironment("COMSPEC")
            If Err Then
                Err.Clear
                On Error Goto 0
                intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
            Else
                On Error Goto 0
                If TestObjectForData(strWorkingCommandPromptPath) = False Then
                    intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                End If
            End If
        End If
    End If

    If intFunctionReturn < 0 Then
        intReturnMultiplier = intReturnMultiplier * 8
        intReturnCode = GetWindowsSystemPath(strWindowsSystemPath)
        If intReturnCode <> 0 Then
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        Else
            intReturnMultiplier = intReturnMultiplier * 8
            On Error Resume Next
            Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
            Else
                ' Try cmd.exe
                boolResult = objFileSystemObject.FileExists(strWindowsSystemPath & "cmd.exe")
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
                Else
                    On Error Goto 0
                    If boolResult = True Then
                        strWorkingCommandPromptPath = strWindowsSystemPath & "cmd.exe"
                        intFunctionReturn = 1
                    Else
                        ' Try command.com
                        On Error Resume Next
                        boolResult = objFileSystemObject.FileExists(strWindowsSystemPath & "command.com")
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
                        Else
                            On Error Goto 0
                            If boolResult = True Then
                                strWorkingCommandPromptPath = strWindowsSystemPath & "command.com"
                                intFunctionReturn = 1
                            Else
                                intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn < 0 Then
        intReturnMultiplier = intReturnMultiplier * 8
        intReturnCode = GetWindowsPath(strWindowsPath)
        If intReturnCode <> 0 Then
            intFunctionReturn = intFunctionReturn + (intReturnCode * intReturnMultiplier)
        Else
            intReturnMultiplier = intReturnMultiplier * 8
            On Error Resume Next
            Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (-1 * intReturnMultiplier)
            Else
                ' Try cmd.exe
                boolResult = objFileSystemObject.FileExists(strWindowsPath & "cmd.exe")
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intFunctionReturn = intFunctionReturn + (-2 * intReturnMultiplier)
                Else
                    On Error Goto 0
                    If boolResult = True Then
                        strWorkingCommandPromptPath = strWindowsPath & "cmd.exe"
                        intFunctionReturn = 2
                    Else
                        ' Try command.com
                        On Error Resume Next
                        boolResult = objFileSystemObject.FileExists(strWindowsPath & "command.com")
                        If Err Then
                            On Error Goto 0
                            Err.Clear
                            intFunctionReturn = intFunctionReturn + (-3 * intReturnMultiplier)
                        Else
                            On Error Goto 0
                            If boolResult = True Then
                                strWorkingCommandPromptPath = strWindowsPath & "command.com"
                                intFunctionReturn = 2
                            Else
                                intFunctionReturn = intFunctionReturn + (-4 * intReturnMultiplier)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If intFunctionReturn >= 0 Then
        strCommandPromptPath = strWorkingCommandPromptPath
    End If
    
    GetCommandPromptPath = intFunctionReturn
End Function

Function GetTempFolderPath(ByRef strTempFolderPath)
    'region FunctionMetadata ####################################################
    ' Safely obtains the path to the temporary files folder
    '
    ' Function takes one positional argument (strTempFolderPath) that is populated upon
    ' success with the path to the temprary files folder. The path is appended with a trailing
    ' backslash.
    '
    ' The function returns 0 if the temporary folder path was retrieved successfully. A
    ' negative number is returned if the temporary folder path was not retrieved successfully.
    '
    ' Example:
    ' intReturnCode = GetTempFolderPath(strTempFolderPath)
    ' If intReturnCode = 0 Then
    '   ' Temporary folder path was retrieved successfully and stored in strTempFolderPath.
    ' End If
    '
    ' Note: the technique used in this function requires Windows Scripting Host 2.0 or newer,
    ' which was included in Windows releases beginning with Windows 2000 and Windows ME. It was
    ' available as a separate download for Windows 95, 98, and NT 4.0.
    '
    ' Note: if the processor architecture of the VBScript process does not match the operating
    ' system's processor architecture (e.g., 32-bit Intel IA32/x86 VBScript process running on
    ' 64-bit AMD64/Intel x86-64 Windows), then the path to the Windows System folder may be
    ' automatically substituted for the Windows-on-Windows (WOW) equivalent path.
    '
    ' Version: 1.1.20210118.0
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

    Dim intFunctionReturn
    Dim objFileSystemObject
    Dim objFolder
    Dim strWorkingTempFolderPath

    Err.Clear

    intFunctionReturn = 0

    On Error Resume Next
    Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
    If Err Then
        On Error Goto 0
        Err.Clear
        intFunctionReturn = -1
    Else
        Set objFolder = objFileSystemObject.GetSpecialFolder(2)
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = -2
        Else
            strWorkingTempFolderPath = objFolder.Path
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = -3
            Else
                On Error Goto 0
                If TestObjectForData(strWorkingTempFolderPath) = False Then
                    intFunctionReturn = -4
                Else
                    If Right(strWorkingTempFolderPath, 1) <> "\" Then
                        strWorkingTempFolderPath = strWorkingTempFolderPath & "\"
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        strTempFolderPath = strWorkingTempFolderPath
    End If
    
    GetTempFolderPath = intFunctionReturn
End Function

Function GetTempFilePath(ByRef strTempFilePath)
    'region FunctionMetadata ####################################################
    ' Safely obtains the path to the temporary files folder
    '
    ' Function takes one positional argument (strTempFilePath) that is populated upon
    ' success with the path to the temprary files folder. The path is appended with a trailing
    ' backslash.
    '
    ' The function returns 0 if the temporary folder path was retrieved successfully. A
    ' negative number is returned if the temporary folder path was not retrieved successfully.
    '
    ' Example:
    ' intReturnCode = GetTempFilePath(strTempFilePath)
    ' If intReturnCode = 0 Then
    '   ' Temporary folder path was retrieved successfully and stored in strTempFilePath.
    ' End If
    '
    ' Note: the technique used in this function requires Windows Scripting Host 2.0 or newer,
    ' which was included in Windows releases beginning with Windows 2000 and Windows ME. It was
    ' available as a separate download for Windows 95, 98, and NT 4.0.
    '
    ' Note: if the processor architecture of the VBScript process does not match the operating
    ' system's processor architecture (e.g., 32-bit Intel IA32/x86 VBScript process running on
    ' 64-bit AMD64/Intel x86-64 Windows), then the path to the Windows System folder may be
    ' automatically substituted for the Windows-on-Windows (WOW) equivalent path.
    '
    ' Version: 1.0.20210118.0
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

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim strTempFolderPath
    Dim objFileSystemObject
    Dim strTempFile
    Dim strWorkingTempFilePath

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 1

    intReturnCode = GetTempFolderPath(strTempFolderPath)
    If intReturnCode <> 0 Then
        intFunctionReturn = intReturnCode
    Else
        intReturnMultiplier = intReturnMultiplier * 8
        On Error Resume Next
        Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = -1 * intReturnMultiplier
        Else
            strTempFile = objFileSystemObject.GetTempName
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = -2 * intReturnMultiplier
            Else
                On Error Goto 0
                If TestObjectForData(strTempFile) = False Then
                    intFunctionReturn = -3 * intReturnMultiplier
                Else
                    strWorkingTempFilePath = strTempFolderPath & strTempFile
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        strTempFilePath = strWorkingTempFilePath
    End If
    
    GetTempFilePath = intFunctionReturn
End Function

Function GetWindowsOperatingSystemVersionNumberAsStringUsingCommandPromptVerCommand(ByRef strOperatingSystemVersion)
    'region FunctionMetadata ####################################################
    ' Safely obtains the operating system version number from the Command Prompt using the
    ' "ver" command
    '
    ' Function takes one positional arguments (strOperatingSystemVersion), which will be
    '       populated with the operating system version in string format upon success
    '
    ' The function returns 0 or a positive number if the operating system version number was
    ' retrieved successfully. A negative number is returned if the operating system version
    ' number was not retrieved successfully.
    '
    ' Example:
    '   intReturnCode = GetWindowsOperatingSystemVersionNumberAsStringUsingCommandPromptVerCommand(strOperatingSystemVersion)
    '   If intReturnCode = 0 Then
    '       ' strOperatingSystemVersion is populated with the operating system version number
    '       ' in string format.
    '   Else
    '       ' The operating system version number could not be retrieved via the Command
    '       ' Prompt's "ver" command.
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

    Dim intFunctionReturn
    Dim intReturnMultiplier
    Dim intReturnCode
    Dim strCommandPromptPath
    Dim strTempFilePath
    Dim objWSHShell
    Dim objFileSystemObject
    Dim strWorkingOperatingSystemVersion
    Dim objTextStreamTempFile
    Dim boolFoundLine
    Dim strLine
    Dim arrLine
    Dim intCounter
    Dim arrLinePortion
    Dim arrLinePortion2
    Dim boolFoundVersionNumber

    Const forReading = 1

    Err.Clear

    intFunctionReturn = 0
    intReturnMultiplier = 1

    intReturnCode = GetCommandPromptPath(strCommandPromptPath)
    If intReturnCode < 0 Then
        intFunctionReturn = intFunctionReturn + (intReturnMultiplier * intReturnCode)
    Else
        intReturnMultiplier = intReturnMultiplier * 8
        intReturnMultiplier = intReturnMultiplier * 8
        intReturnMultiplier = intReturnMultiplier * 8
        intReturnCode = GetTempFilePath(strTempFilePath)
        If intReturnCode <> 0 Then
            intFunctionReturn = intFunctionReturn + (intReturnMultiplier * intReturnCode)
        Else
            intReturnMultiplier = intReturnMultiplier * 8
            intReturnMultiplier = intReturnMultiplier * 8
            On Error Resume Next
            Set objWSHShell = CreateObject("WScript.Shell")
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -1)
            Else
                Set objFileSystemObject = CreateObject("Scripting.FileSystemObject")
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -2)
                Else
                    intReturnCode = objWSHShell.Run("""" & strCommandPromptPath & """ /c ""ver > """ & strTempFilePath & """""", 0, True)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -3)
                    Else
                        On Error Goto 0
                    End If
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No errors have occurred
        ' The file at strTempFilePath is populated with the output of the "ver" command
        ' objFileSystemObject was created successfully
        On Error Resume Next
        Set objTextStreamTempFile = objFileSystemObject.OpenTextFile(strTempFilePath, forReading, False)
        If Err Then
            Err.Clear
            WScript.Sleep(100*Rnd())
            Set objTextStreamTempFile = objFileSystemObject.OpenTextFile(strTempFilePath, forReading, False)
            If Err Then
                Err.Clear
                WScript.Sleep(200*Rnd())
                Set objTextStreamTempFile = objFileSystemObject.OpenTextFile(strTempFilePath, forReading, False)
                If Err Then
                    Err.Clear
                    WScript.Sleep(800*Rnd())
                    Set objTextStreamTempFile = objFileSystemObject.OpenTextFile(strTempFilePath, forReading, False)
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -4)
                    Else
                        On Error Goto 0
                    End If
                Else
                    On Error Goto 0
                End If
            Else
                On Error Goto 0
            End If
        Else
            On Error Goto 0
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No errors have occurred
        ' The file at strTempFilePath is populated with the output of the "ver" command
        ' objFileSystemObject was created successfully
        ' The temp file is open for reading using objTextStreamTempFile
        boolFoundLine = False
        On Error Resume Next
        Do Until ((objTextStreamTempFile.AtEndOfStream) Or (boolFoundLine = True))
            strLine = objTextStreamTempFile.ReadLine
            If TestObjectForData(strLine) = True Then
                arrLine = Split(strLine, " ")
                If UBound(arrLine) > 0 Then
                    boolFoundLine = True
                End If
            End If
        Loop
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -5)
        Else
            On Error Goto 0
            If boolFoundLine = False Then
                intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -6)
            Else
                'arrLine is already a split of strLine by space
                boolFoundVersionNumber = False
                For intCounter = 0 To UBound(arrLine)
                    arrLinePortion = Split(arrLine(intCounter), ".")
                    If UBound(arrLinePortion) > 0 Then
                        'arrLine(intCounter) contains what appears to be a version number
                        boolFoundVersionNumber = True
                        arrLinePortion2 = Split(arrLine(intCounter), "]")
                        strWorkingOperatingSystemVersion = arrLinePortion2(0)
                        Exit For
                    End If
                Next
                If boolFoundVersionNumber = False Then
                    intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -7)
                Else
                    If TestObjectForData(strWorkingOperatingSystemVersion) = False Then
                        intFunctionReturn = intFunctionReturn + (intReturnMultiplier * -8)
                    End If
                End If
            End If
        End If
        On Error Resume Next
        objTextStreamTempFile.Close
        If Err Then
            On Error Goto 0
            Err.Clear
            ' do not return error
        Else
            Set objTextStreamTempFile = Nothing
            If Err Then
                On Error Goto 0
                Err.Clear
                ' do not return error
            Else
                objFileSystemObject.DeleteFile strTempFilePath, True
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    ' do not return error
                Else
                    On Error Goto 0
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        strOperatingSystemVersion = strWorkingOperatingSystemVersion
    End If
    
    GetWindowsOperatingSystemVersionNumberAsStringUsingCommandPromptVerCommand = intFunctionReturn
End Function
