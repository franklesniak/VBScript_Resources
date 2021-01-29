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

Function ReadCharactersFromFile(ByRef strData, ByVal strPathToFile, ByVal lngMaxNumberOfCharactersToRead, ByVal boolContinueOnError)
    'region FunctionMetadata ####################################################
    ' Safely reads-in characters from a file at path strPathToFile and stores them in a string
    ' (strData)
    '
    ' This function takes four arguments:
    '   - The first argument strData) is populated upon success with a string containing all
    '       of the characters that were read-in from the file.
    '   - The second argument (strPathToFile) is a string containing the path to the file to be
    '       read-in by this function.
    '   - The third argument (lngMaxNumberOfCharactersToRead) allows the caller to set an upper
    '       boundary on the number of characters read-in from the file. It can be set to an
    '       integrer, or set to Null if there is no limit.
    '   - The fourth argument (boolContinueOnError) allows the caller to specify whether the
    '       function should continue reading-in characters if the operation to read one
    '       character fails. If set to True, the process continues and drops the character that
    '       resulted in a read error. If set to False or Null, the process would stop on error
    '       and return a numerical code indicating failure (see below).
    '
    ' The function returns 0 if the characters were read-in from the specified file
    ' successfully; it returns a negative number if the characters were not able to be read
    '
    ' Example:
    '   intReturnCode = ReadCharactersFromFile(strData, "C:\Users\flesniak\Desktop\TestFile.txt", 60000, Null)
    '   If intReturnCode = 0 Then
    '       ' The file was read successfully and was capped at a maximum of 60,000 characters
    '       ' strData contains the characters read from the file
    '   End If
    '
    ' Version: 1.1.20210129.0
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
    ' StackExchange user "jumpjack", for the post that inspired the creation of this function:
    ' https://superuser.com/a/1027161/334370
    'endregion Acknowledgements ####################################################

    Dim intFunctionReturn
    Dim boolWorkingContinueOnError
    Dim fileSystemObject
    Dim boolTest
    Dim fileObjectSource
    Dim textStreamObjectSource
    Dim strWorkingOutput
    Dim lngCounter
    Dim boolBreakOut

    Err.Clear
    
    intFunctionReturn = 0

    If TestObjectForData(boolContinueOnError) = False Then
        boolWorkingContinueOnError = False
    Else
        On Error Resume Next
        boolTest = (boolContinueOnError = True)
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = -1
        Else
            On Error Goto 0
            If boolTest Then
                boolWorkingContinueOnError = True
            Else
                boolWorkingContinueOnError = False
            End If
        End If
    End If

    On Error Resume Next
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    If Err Then
        On Error Goto 0
        Err.Clear
        intFunctionReturn = -2
    Else
        boolTest = fileSystemObject.FileExists(strPathToFile)
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = -3
        Else
            On Error Goto 0
            If boolTest = False Then
                ' File specified by strPathToFile did not exist
                intFunctionReturn = -4
            Else
                On Error Resume Next
                Set fileObjectSource = fileSystemObject.GetFile(strPathToFile)
                If Err Then
                    On Error Goto 0
                    Err.Clear
                    intFunctionReturn = -5
                Else
                    On Error Goto 0
                End If
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        ' No error occurred yet
        ' fileObjectSource is a FileObject consisting of the source file
        On Error Resume Next
        Set textStreamObjectSource = fileObjectSource.OpenAsTextStream()
        If Err Then
            On Error Goto 0
            Err.Clear
            intFunctionReturn = -6
        Else
            On Error Goto 0
            strWorkingOutput = ""
            lngCounter = CLng(0)
            boolBreakOut = False
            On Error Resume Next
            boolTest = ((textStreamObjectSource.AtEndOfStream = False) And (lngCounter < lngMaxNumberOfCharactersToRead) And (boolBreakOut = False))
            If Err Then
                On Error Goto 0
                Err.Clear
                intFunctionReturn = -7
            Else
                While boolTest = True
                    strWorkingOutput = strWorkingOutput + textStreamObjectSource.Read(1)
                    If Err Then
                        Err.Clear
                        If boolWorkingContinueOnError = False Then
                            On Error Goto 0
                            intFunctionReturn = -8
                            boolBreakOut = True
                        End If
                    End If
                    lngCounter = lngCounter + 1
                    boolTest = ((textStreamObjectSource.AtEndOfStream = False) And (lngCounter < lngMaxNumberOfCharactersToRead) And (boolBreakOut = False))
                    If Err Then
                        On Error Goto 0
                        Err.Clear
                        intFunctionReturn = -9
                        boolBreakOut = True
                        boolTest = False
                    End If
                Wend
                On Error Goto 0
            End If
        End If
    End If

    If intFunctionReturn = 0 Then
        strData = strWorkingOutput
    End If

    ReadCharactersFromFile = intFunctionReturn
End Function
    