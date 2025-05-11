Attribute VB_Name = "Module1"
'****************************************************************
'Read 24-bit BitMap file and fill in cells of Excel
'
'
'08/12/2022
'Tom Yang
'*****************************************************************
Option Explicit

Const BMP_START = 54 'Raw Bitmap Data ADDRESS

Const ADD_biWidth = &H12    '18
Const ADD_biHeight = &H16   '22
Const ADD_biBitCount = &H1C  '28

Dim BMP_arr() As Byte
Dim width As String
Dim height As String
Dim row_size As Integer
Dim size_file As Long
Dim BitCount As Integer

Function read_BMP(file As String)

Dim arrByte() As Byte, fNo#

fNo = FreeFile
    Open file For Binary Access Read As fNo
    size_file = LOF(fNo)
    ReDim arrByte(size_file - 1)
    Get fNo, , arrByte
Close fNo

read_BMP = arrByte
width = WorksheetFunction.Dec2Hex(arrByte(ADD_biWidth + 3), 2) & WorksheetFunction.Dec2Hex(arrByte(ADD_biWidth + 2), 2) & WorksheetFunction.Dec2Hex(arrByte(ADD_biWidth + 1), 2) & WorksheetFunction.Dec2Hex(arrByte(ADD_biWidth), 2)
height = WorksheetFunction.Hex2Dec(WorksheetFunction.Dec2Hex(arrByte(ADD_biHeight + 3), 2) & WorksheetFunction.Dec2Hex(arrByte(ADD_biHeight + 2), 2) & WorksheetFunction.Dec2Hex(arrByte(ADD_biHeight + 1), 2) & WorksheetFunction.Dec2Hex(arrByte(ADD_biHeight), 2))
row_size = WorksheetFunction.RoundDown((WorksheetFunction.Hex2Dec(width) * 24 + 31) / 32, 0) * 4
BitCount = WorksheetFunction.Hex2Dec(WorksheetFunction.Dec2Hex(arrByte(ADD_biBitCount + 1), 2) & WorksheetFunction.Dec2Hex(arrByte(ADD_biBitCount), 2))


End Function

Sub PUTOUT_BMP()

Dim filename As String
filename = Application.GetOpenFilename("24-bit BitMap(*.bmp;*.dib),*.bmp")

BMP_arr() = read_BMP(filename)
If size_file > 900000 Then
    MsgBox "The file is too big! the file must smaller then 900 KB!", vbCritical
    Exit Sub
ElseIf Not BitCount = 24 Then
    MsgBox "The file is not 24-bit BitMap, please double check!", vbCritical
    Exit Sub

End If



Application.ScreenUpdating = False
Dim i As Long
Dim j As Long
Dim k As Long
Dim c As Long

k = row_size - WorksheetFunction.Hex2Dec(width) * 3

'For i = 0 To size_file - 1
'    Cells(i + 1, 3) = BMP_arr(i)
'Next


c = BMP_START
For j = height To 1 Step -1
    For i = 1 To WorksheetFunction.Hex2Dec(width)
    
            Cells(0 + j, 3 + i).Interior.Color = RGB(CInt(BMP_arr(c + 2)), CInt(BMP_arr(c + 1)), CInt(BMP_arr(c))) 'BMP color is BGR
            c = c + 3
    Next
    c = c + k
Next

'For j = 1 To height
'    For i = 1 To WorksheetFunction.Hex2Dec(width)
    
'            Cells(0 + j, 3 + i).Interior.Color = RGB(CInt(BMP_arr(c+2)), CInt(BMP_arr(c + 1)), CInt(BMP_arr(c))) 'BMP color is BGR
'            c = c + 3
'    Next
'    c = c + k
'Next
Application.ScreenUpdating = True

MsgBox "Done!"

End Sub

Sub clearAll()

    Worksheets("sheet1").UsedRange.ClearFormats
    Worksheets("sheet1").UsedRange.ClearContents

End Sub
