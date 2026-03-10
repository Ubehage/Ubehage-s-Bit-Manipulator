Attribute VB_Name = "modBit"
Option Explicit

Private Const MAX_DATASIZE = SIZE_MEGA * 10

Public Enum Bit_Manipulation_Method
  bmFlip = &H1
  bmRemove = &H2
End Enum

Public Function ManipulateBitInFile(FileName As String, BytePosition As Double, BitIndex As Integer, BitOperation As Bit_Manipulation_Method) As Boolean
  If (BitIndex < 1 Or BitIndex > 8) Then Exit Function
  Dim bFile As BinaryFileReader, bData() As Byte
  Set bFile = New BinaryFileReader
  ReDim bData(1 To 1) As Byte
  If OpenFileInFileReader(bFile, FileName, False) = False Then GoTo ExitManipulate
  bFile.SeekAbsolute (BytePosition - 1)
  If ReadFromFileReader(bFile, bData()) = False Then GoTo ExitManipulate
  ManipulateSingleBitArray bData(), BitIndex, BitOperation
  bFile.SeekAbsolute (BytePosition - 1)
  If WriteToFileReader(bFile, bData()) = False Then GoTo ExitManipulate
  ManipulateBitInFile = True
ExitManipulate:
  CloseFileReader bFile
End Function

Public Function ManipulateBitToNewFile(SourceFileName As String, TargetFileName As String, BytePosition As Double, BitIndex As Integer, BitOperation As Bit_Manipulation_Method) As Boolean
  If (BitIndex < 1 Or BitIndex > 8) Then Exit Function
  Dim bFileRead As BinaryFileReader, bFileWrite As BinaryFileReader, bData() As Byte, TotalSize As Double
  TotalSize = GetFileSizeA(SourceFileName)
  Set bFileRead = New BinaryFileReader
  Set bFileWrite = New BinaryFileReader
  If OpenFileInFileReader(bFileRead, SourceFileName, True) = False Then GoTo ExitManipulate
  If OpenFileInFileReader(bFileWrite, TargetFileName, False, True) = False Then GoTo ExitManipulate
  If CopyBytesFromFileToFile(bFileRead, bFileWrite, (BytePosition - 1)) = False Then GoTo ExitManipulate
  ReDim bData(1 To 1) As Byte
  If ReadFromFileReader(bFileRead, bData()) = False Then GoTo ExitManipulate
  ManipulateSingleBitArray bData(), BitIndex, BitOperation
  If WriteToFileReader(bFileWrite, bData()) = False Then GoTo ExitManipulate
  If CopyBytesFromFileToFile(bFileRead, bFileWrite, (TotalSize - BytePosition)) = False Then GoTo ExitManipulate
  ManipulateBitToNewFile = True
ExitManipulate:
  CloseFileReader bFileRead
  CloseFileReader bFileWrite
End Function

Private Sub ManipulateSingleBitArray(BitArray() As Byte, BitIndex As Integer, BitOperation As Bit_Manipulation_Method)
  If BitOperation = bmFlip Then
    BitArray(1) = FlipBit(BitArray(1), BitIndex)
  ElseIf BitOperation = bmRemove Then
    BitArray(1) = RemoveBit(BitArray(1), BitIndex)
  End If
End Sub

Private Function FlipBit(InByte As Byte, BitIndex As Integer) As Byte
  If (BitIndex < 1 Or BitIndex > 8) Then Exit Function
  Dim bitmask As Byte
  bitmask = 2 ^ (BitIndex - 1)
  FlipBit = InByte Xor bitmask
End Function

Private Function RemoveBit(InByte As Byte, BitIndex As Integer) As Byte
  If (BitIndex < 1 Or BitIndex > 8) Then Exit Function
  Dim highByte As Byte, lowByte As Byte, bitmask As Byte, r As Byte
  bitmask = 2 ^ (BitIndex - 1)
  lowByte = InByte And (bitmask - 1)
  highByte = InByte \ (bitmask * 2)
  r = (highByte * bitmask) Or lowByte
  RemoveBit = (r * 2) And &HFF
End Function

Private Function OpenFileInFileReader(FileReader As BinaryFileReader, FilePath As String, ReadOnly As Boolean, Optional ForceDeleteExistingFile As Boolean = False) As Boolean
  On Error GoTo OpenError
  If ReadOnly = True Then
    FileReader.OpenForRead FilePath
  Else
    FileReader.OpenFile FilePath, ForceDeleteExistingFile
  End If
  OpenFileInFileReader = True
ExitOpen:
  On Error GoTo 0
  Exit Function
OpenError:
  Select Case ShowErrorMessage("Opening file for " & IIf(ReadOnly, "read", "write"), Err.Description, Err.Number, FilePath)
    Case MessageBoxButtons.mbbRetry
      Resume
    Case Else
      Resume ExitOpen
  End Select
End Function

Private Sub CloseFileReader(FileReader As BinaryFileReader)
  If Not FileReader Is Nothing Then If FileReader.IsOpen() = True Then FileReader.CloseFile
End Sub

Private Function CopyBytesFromFileToFile(SourceFileReader As BinaryFileReader, TargetFileReader As BinaryFileReader, BytesToCopy As Double) As Boolean
  Dim fData() As Byte, bSize As Double, dSize As Double
  bSize = BytesToCopy
  dSize = IIf(bSize >= MAX_DATASIZE, MAX_DATASIZE, bSize)
  If dSize <= 0 Then GoTo ExitCopySuccess
  ReDim fData(1 To dSize) As Byte
  Do While bSize > 0
    If bSize < dSize Then
      dSize = bSize
      ReDim fData(1 To dSize) As Byte
    End If
    If ReadFromFileReader(SourceFileReader, fData()) = False Then GoTo ExitCopy
    If WriteToFileReader(TargetFileReader, fData()) = False Then GoTo ExitCopy
    bSize = (bSize - dSize)
  Loop
ExitCopySuccess:
  CopyBytesFromFileToFile = True
ExitCopy:
End Function

Private Function ReadFromFileReader(SourceFileReader As BinaryFileReader, DataArray() As Byte) As Boolean
  On Error GoTo ReadError
  Call SourceFileReader.ReadBytes(DataArray())
  ReadFromFileReader = True
ExitRead:
  On Error GoTo 0
  Exit Function
ReadError:
  Select Case ShowErrorMessage("Reading from file", Err.Description, Err.Number, SourceFileReader.FileName)
    Case MessageBoxButtons.mbbRetry
      Resume
    Case Else
      Resume ExitRead
  End Select
End Function

Private Function WriteToFileReader(TargetFileReader As BinaryFileReader, DataArray() As Byte) As Boolean
  On Error GoTo WriteError
  Call TargetFileReader.WriteBytes(DataArray())
  WriteToFileReader = True
ExitWrite:
  On Error GoTo 0
  Exit Function
WriteError:
  Select Case ShowErrorMessage("Writing to file", Err.Description, Err.Number, TargetFileReader.FileName)
    Case MessageBoxButtons.mbbRetry
      Resume
    Case Else
      Resume ExitWrite
  End Select
End Function
