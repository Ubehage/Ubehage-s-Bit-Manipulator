Attribute VB_Name = "modBit"
Option Explicit

Public Enum Bit_Manipulation_Method
  bmFlip = &H1
  bmRemove = &H2
End Enum

Public Function ManipulateBitInFile(FileName As String, BytePosition As Double, BitIndex As Integer, BitOperation As Bit_Manipulation_Method) As Boolean
  If (BitIndex < 1 Or BitIndex > 8) Then Exit Function
  Dim bFile As BinaryFileReader, bData() As Byte
  Set bFile = New BinaryFileReader
  ReDim bData(1 To 1) As Byte
  On Error GoTo ManipulateError
  bFile.OpenForRead FileName
  bFile.SeekAbsolute (BytePosition - 1)
  Call bFile.ReadBytes(bData())
  bFile.CloseFile
  ManipulateSingleBitArray bData(), BitIndex, BitOperation
  bFile.OpenForWrite FileName
  bFile.SeekAbsolute (BytePosition - 1)
  Call bFile.WriteBytes(bData)
  bFile.CloseFile
  ManipulateBitInFile = True
DoneManipulating:
  On Error GoTo 0
  Exit Function
ManipulateError:
  'do something...
  Resume
End Function

Public Function ManipulateBitToNewFile(SourceFileName As String, TargetFileName As String, BytePosition As Double, BitIndex As Integer, BitOperation As Bit_Manipulation_Method) As Boolean
  If (BitIndex < 1 Or BitIndex > 8) Then Exit Function
  Dim bFileRead As BinaryFileReader, bFileWrite As BinaryFileReader, bData() As Byte
  Set bFileRead = New BinaryFileReader
  Set bFileWrite = New BinaryFileReader

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
