VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Const CSV_QUOTATION As String = """"
Private Const CSV_DELIMITER As String = ","
Private InitialFileName As String
Private InitialFolderName As String
Private Const BOM_BIG_ENDIAN As Long = 65279        ' -> &HFEFF
Private Const BOM_LITTLE_ENDIAN As Long = 65534     ' -> &HFFFE

Public Enum Endianess
    BigEndian
    LittleEndian
End Enum

Public Function CSVFileToArray(FileName As String, Optional TrimText As Boolean = True) As Variant
    If Split(FileName, ".")(1) Like "csv" Then
        CSVFileToArray = TextFileToArray(FileName, CSV_DELIMITER, CSV_QUOTATION, TrimText)
    End If
End Function

Public Function ArrayToTextFile(FileName As String, ByRef SaveArray As Variant, Delimiter As String, Quotation As String, Optional TrimText As Boolean = True, Optional AlwaysAddQuotation As Boolean = False, Optional Unicode As Boolean, Optional Endian As Endianess = BigEndian) As Boolean
    If IsEmpty(SaveArray) Then Exit Function
'------------------------------------------------------------------------------------
    Dim FileNumber As Integer: FileNumber = FreeFile
    Dim TextLine As String
    Dim TextStream() As Byte
    Dim DelimiterFound As Boolean
    Dim TextValue As String
    Dim ReplacedText As String
    Dim TextPosition As Long
    Dim Row As Long
    Dim Column As Long
    
'------------------------------------------------------------------------------------
'----------------------------START OF WRITING TEXT FILE------------------------------
'------------------------------------------------------------------------------------
    '** Cleans the file if it exists
    Open FileName For Output As #FileNumber: Close #FileNumber
    
    Open FileName For Binary Access Write As #FileNumber
    On Error GoTo CloseFunction
    
    For Row = LBound(SaveArray, 1) To UBound(SaveArray, 1)
        TextLine = ""
        For Column = LBound(SaveArray, 2) To UBound(SaveArray, 2)
            TextValue = CStr(SaveArray(Row, Column))
            DelimiterFound = InStr(1, TextValue, Delimiter)
            ReplacedText = Replace(TextValue, Quotation, Quotation & Quotation)
            If ReplacedText <> TextValue Or DelimiterFound Or AlwaysAddQuotation Then
            '** If the Text in TextValue contatins the quotation mark then
            '   change them for double quotation, and closed between
            '   quotation marks
                TextValue = Quotation & ReplacedText & Quotation
            End If
            If TrimText Then Trim (TextValue)
            If Column = LBound(SaveArray, 2) Then
                TextLine = TextValue
            Else
                TextLine = TextLine & Delimiter & TextValue
            End If
        Next Column
        
        If Unicode Then
            If Row = 1 Then
                Dim LeftValue As Long
                Dim RightValue As Long
                Select Case Endian
                    Case Endianess.LittleEndian
                        LeftValue = (BOM_LITTLE_ENDIAN And &HFF00) / CLng(&H100)
                        RightValue = (BOM_LITTLE_ENDIAN And &HFF)
                    Case Endianess.BigEndian
                        LeftValue = (BOM_BIG_ENDIAN And &HFF00) / CLng(&H100)
                        RightValue = (BOM_BIG_ENDIAN And &HFF)
                    Case Else
                        Err.Raise 1001, , "Invalid Endianess"
                        Exit Function
                End Select
                TextLine = ChrW(BOM_BIG_ENDIAN) & TextLine
            End If
            TextLine = StrConv(TextLine & vbCrLf, vbUnicode)
            TextStream = StrConv(TextLine, vbFromUnicode)
        Else
            TextStream = TextLine & vbCrLf
        End If
        
        Put #FileNumber, , TextStream
        
    Next Row
    
    '** If everything succeed the function returns 'True'
    ArrayToTextFile = True

'------------------------------------------------------------------------------------
'---------------------------FINISH OF WRITING TEXT FILE------------------------------
'------------------------------------------------------------------------------------
CloseFunction:
    Close #FileNumber
    If Not IsEmpty(TextStream) Then Erase TextStream
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Public Function TextFileToArray(FileName As String, Delimiter As String, Quotation As String, Optional TrimText As Boolean = True) As Variant
    Dim objCollection As Collection: Set objCollection = New Collection
    Dim FileNumber As Integer
    Dim TextLine As String
    Dim BOM As Long
    Dim MainString As String
    Dim SubString As String
    Dim ValueString As String
    Dim CurrentChar As String
    Dim PreviousSubChar As String
    Dim CurrentSubChar As String
    Dim NextSubChar As String
    Dim IsQuote As Boolean
    Dim IsFinish As Boolean
    Dim CurrentDelimiterPosition As Long
    Dim PreviousDelimiterPosition As Long: PreviousDelimiterPosition = 0
    Dim XY As String
    Dim Columns As Long
    Dim Column As Long: Column = 0
    Dim Rows As Long
    Dim Row As Long: Row = 0
    Dim i, j, x, y As Long
'------------------------------------------------------------------------------------
'----------------------------START OF READING TEXT FILE------------------------------
'------------------------------------------------------------------------------------

If Dir(FileName) = "" Or IsFileOpen(FileName) Then Exit Function

FileNumber = FreeFile()
Open FileName For Input As #FileNumber

On Error GoTo ForceClose
Do
    '** Get the text of the next line in the Opened text file
    Line Input #FileNumber, TextLine

'------------------------------------------------------------------------------------
'-----------------------START OF DECOUPLING TEXT LINE--------------------------------
'------------------------------------------------------------------------------------
    PreviousDelimiterPosition = 0
    SubString = TextLine
    Row = Row + 1
    Column = 0

    '** Removes BOM if found
    If Row = 1 Then
        BOM = CLng(AscW(Mid(TextLine, 1, 1))) * CLng(&H100) + CLng(AscW(Mid(TextLine, 2, 2)))
        If BOM = BOM_LITTLE_ENDIAN Or BOM = BOM_BIG_ENDIAN Then
            TextLine = Mid(TextLine, 3, Len(TextLine))
        End If
    End If

    If Row > Rows Then Rows = Row
    For i = 1 To Len(TextLine)
    If i > Len(TextLine) Then Exit For
        SubString = ""
        PreviousSubChar = ""
        CurrentSubChar = ""
        NextSubChar = ""
        
        '** Stores the next found delimiter position
        CurrentDelimiterPosition = InStr(PreviousDelimiterPosition + 1, TextLine, Delimiter)
        
        If CurrentDelimiterPosition > 0 Then
        '** Checks the value between CurrentDelimiterPostion and PreviousDelimiterPosistion
            SubString = Mid(TextLine, PreviousDelimiterPosition + 1, CurrentDelimiterPosition - PreviousDelimiterPosition - 1)
        Else
        '** If nothing is found means the SubString is between the the PreviousDelimiterPosition
        '    and End position, and send virtual delimiter position to the end of the line
            CurrentDelimiterPosition = Len(TextLine)
            SubString = Mid(TextLine, PreviousDelimiterPosition + 1, Len(TextLine))
        End If
        '** --------------------------  ------------------------------------------------------------
        
        '** Now check if the current SubString is Quoted there are only 2 cases:
        '   1.- If the sum of the totals current Quotes marks is an "Odd" number
        '       means is OpenQuoted.
        '   2.- If the sum of the totals current Quotes marks is a "Pair" number
        '       means is an Ordinary String, hence:
        '       a) May be inbetween Quote marks, this way must be removed those marks
        '       b) May is an Ordinary string
        x = 0
        For y = 1 To Len(SubString)
            If Mid(SubString, y, 1) Like Quotation Then
                x = x + 1
            End If
        Next y
        IsQuote = x Mod 2 <> 0
        '** --------------------------------------------------------------------------------------
        
        '** If is OpenQuote then loop until find a Finish Quote -> Quotation & Delimiter
        If IsQuote Then
        '** Takes the remaining String from previous delimiter found to the end of the line
            MainString = ""
            IsFinish = False
            Do
            SubString = Mid(TextLine, PreviousDelimiterPosition + 1, Len(TextLine) - PreviousDelimiterPosition)
            For j = 1 To Len(SubString)
                CurrentSubChar = Mid(SubString, j, 1)
                NextSubChar = Mid(SubString, j + 1, 1)
                MainString = MainString & CurrentSubChar
                If ((CurrentSubChar = Quotation) And (NextSubChar = Delimiter)) Or (CurrentSubChar = Quotation And j = Len(SubString)) Or (j = Len(SubString) And EOF(FileNumber)) Then
        '** If Current Substring loop finds the finish combination then now it will check
        '   backwards if sum of "Continuos" quotations marks found:
        '   1.- If is "Odd" number means the current substring is a real finish mark
        '   2.- If is "Pair" number means is a real finsh quote mark, so from
        '       this point all MainString stored is the ValueString
                    x = 0
                    For y = j To 1 Step -1
                        If Mid(SubString, y, 1) <> Quotation Then
                            Exit For ' <-- Exits "y" for loop
                        Else
                            x = x + 1
                        End If
                    Next y
                    IsFinish = x Mod 2 <> 0
        '** Special case where MainString = Quotation && NextSubChar = ", will be treat
        '   MainString as a FinishQuote itself, hence it must be skipped until a
        '   real FinishQuote is found
                    IsFinish = IsFinish And Not (MainString = Quotation And NextSubChar = Delimiter)
        
        '** If IsFinish then stop searching FinishQuote
                    If IsFinish Then
                        IsQuote = False
                        Exit For ' <-- Exits "j" for loop
                    End If
                End If
            Next j
        '** If is Finished send all stored values in MainString to ValueString
            If IsFinish Then
        '** If MainString = Quotation and Delimiter = "" means only a quotation mark was found, therefore
        '   should be threaten as an empty value
                If MainString = Quotation And NextSubChar = "" Then MainString = ""
                CurrentDelimiterPosition = PreviousDelimiterPosition + j + 1
                ValueString = MainString
            Else
        '** Else End of the line reached and not finish found
                If Not (EOF(FileNumber)) Then
        '   1.- If is not the end of the line, read next text line and loop again
                    MainString = MainString & vbCrLf
                    PreviousDelimiterPosition = 0
                    Line Input #FileNumber, TextLine
                Else
        '   2.- Else send all stored values in MainString to ValueString
                    If IsQuote And PreviousDelimiterPosition + j > Len(TextLine) Then
        '** If there exists special case where is the end of the line and the end of the file,
        '   and the open quoted value was never closed with a quotation mark, this close it.
                        If CurrentSubChar <> Quotation Then MainString = MainString & Quotation
                    End If
                    ValueString = MainString
                    IsFinish = True
                End If
            End If
            Loop Until IsFinish
        Else
            ValueString = SubString
        End If
        
        
        If Not ValueString Like "" Then
            PreviousSubChar = Mid(ValueString, 1, 1)
            NextSubChar = Mid(ValueString, Len(ValueString), 1)
            If PreviousSubChar Like Quotation And NextSubChar Like Quotation Then
            'Check if is inbetween the quotation marks
                ValueString = Left(Right(ValueString, Len(ValueString) - 1), Len(ValueString) - 2)
            End If
        End If
        
        
        i = CurrentDelimiterPosition
        PreviousDelimiterPosition = CurrentDelimiterPosition
        Column = Column + 1
        If Column > Columns Then Columns = Column
        ValueString = Replace(ValueString, Quotation & Quotation, Quotation)
        If TrimText Then ValueString = Trim(ValueString)
        XY = CStr(Row) & ":" & CStr(Column)
        objCollection.Add ValueString, XY
        Debug.Print "--> ValueString is " & ValueString
    Next i
    
'------------------------------------------------------------------------------------
'----------------------FINISH OF DECOUPLING TEXT LINE--------------------------------
'------------------------------------------------------------------------------------

Loop Until EOF(FileNumber)
ForceClose:
Close #FileNumber
If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Error " & Err.Number, Err.HelpFile, Err.HelpContext
Dim FinalArray() As Variant
ReDim FinalArray(1 To Rows, 1 To Columns)
On Error GoTo 0

On Error Resume Next
Dim Item As String
For Row = LBound(FinalArray, 1) To UBound(FinalArray, 1)
    For Column = LBound(FinalArray, 2) To UBound(FinalArray, 2)
            Item = ""
            Item = objCollection.Item(CStr(Row) & ":" & CStr(Column))
            FinalArray(Row, Column) = Item
    Next Column
Next Row
On Error GoTo 0

'------------------------------------------------------------------------------------
'---------------------------FINISH OF READING TEXT FILE------------------------------
'------------------------------------------------------------------------------------
TextFileToArray = FinalArray
Erase FinalArray
Set objCollection = Nothing
End Function

Public Function GetFiles(ByRef SelectionCanceled As Boolean) As String()
 
 If InitialFileName = "" Then InitialFileName = ThisWorkbook.Path
 SelectionCanceled = True
 
 '** Declare a variable as a FileDialog object.
    Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim FileNames() As String
 
    With fd
    
    '** Set filters values for search text-a-like files
    With .Filters
        .Clear
        .Add "Text Files", "*.txt"
        .Add "Comma Separated Values (CSV)", "*.csv"
        .Add "Other", "*.*"
    End With
    
    '** Set multiples selections available
    .AllowMultiSelect = True
    
    'Preset Route this
    .FilterIndex = 1
    
    'Preset Route this
    .InitialFileName = InitialFileName
    
    'Get Path
    If .Show = -1 Then
        SelectionCanceled = False
        ReDim FileNames(.SelectedItems.Count - 1)
        For i = 1 To .SelectedItems.Count
            FileNames(i - 1) = .SelectedItems(i)
        Next i
        GetFiles = FileNames
        Dim ss() As String: ss = Split(FileNames(0), Application.PathSeparator)
        InitialFileName = Split(FileNames(0), Application.PathSeparator & ss(UBound(ss)))(0)
        Erase FileNames
        Erase ss
    End If
    End With
  
 'Set the object variable to Nothing.
    Set fd = Nothing
 
End Function

Public Function IsFileOpen(FileName As String) As Boolean
    Dim FileNumber As Integer, ErrNum As Integer

    On Error Resume Next   ' Turn error checking off.
    FileNumber = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open FileName For Input Lock Read As #FileNumber
    Close FileNumber          ' Close the file.
    ErrNum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case ErrNum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileOpen = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        
        Case 70
            Debug.Print "--> File already open: " & vbCrLf & FileName
            IsFileOpen = True
            
        ' Another error occurred.
        Case Else
            Error ErrNum
    End Select

End Function

'****************************
' Get Destination Folder
'****************************

Friend Function GetFolders(ByRef Canceled As Boolean, Optional MultiSelect As Boolean = True) As String()
    Dim FolderPaths() As String
    Canceled = True
    If InitialFolderName = "" Then
        If InitialFileName = "" Then
            InitialFolderName = InitialFileName
        Else
            InitialFolderName = "C:\Users\" & Environ("UserName") & "\Documents"
        End If
    End If
    Dim FolderPath As String:  FolderPath = ""
    Dim fl As FileDialog: Set fl = Application.FileDialog(msoFileDialogFolderPicker)
    With fl
        .Title = "Destination Folder"
        If MultiSelect Then
            .ButtonName = "Select"
        Else
            .ButtonName = "Here"
        End If
        .InitialFileName = InitialFolderName
        .AllowMultiSelect = MultiSelect
        If .Show = -1 Then
            Canceled = False
            ReDim FolderPaths(.SelectedItems.Count - 1)
            For i = 1 To .SelectedItems.Count
                FolderPaths(i - 1) = .SelectedItems(i)
            Next i
            InitialFolderName = FolderPaths(0)
            Debug.Print "--> Save last folder path: " & vbCrLf & InitialFolderName
        End If
    End With
    GetFolders = FolderPaths
    Set fl = Nothing
End Function
