VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHexClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'I rushed th Hex Editor a bit
'Ended up wrapping into this class
'seems to work OK
Dim mHex() As String
Dim mBin() As Byte
Dim mAsc() As String
Dim HexString As String
Dim AscString As String
Public Property Get HexData() As Variant
    HexData = mHex
End Property
Public Property Let HexData(ByVal vNewValue As Variant)
    Dim Z As Long
    Erase mHex
    Erase mAsc
    mHex = vNewValue
    ReDim mAsc(0 To UBound(mHex))
    For Z = 0 To UBound(mHex)
        mAsc(Z) = HexToAsc(mHex(Z))
    Next
    HexString = Join(mHex, Chr(32))
    AscString = Join(mAsc, Chr(32))
End Property
Public Property Get BinData() As Variant
    BinData = mBin
End Property
Public Property Let BinData(ByVal vNewValue As Variant)
    mBin = vNewValue
End Property
Public Property Get AscData() As Variant
    AscData = mAsc
End Property
Public Property Let AscData(ByVal vNewValue As Variant)
    mAsc = vNewValue
End Property
Public Property Get HexStr() As String
    HexStr = HexString
End Property
Public Property Let HexStr(ByVal vNewValue As String)
    HexString = vNewValue
End Property
Public Property Get AscStr() As String
    AscStr = AscString
End Property
Public Property Let AscStr(ByVal vNewValue As String)
    AscString = vNewValue
End Property
Public Sub LoadRawHex(mSrc As String)
    Dim mVar As Variant, Z As Long, z1 As Long
    HexString = Replace(mSrc, vbLf, Chr(32))
    HexString = Replace(HexString, vbCrLf, Chr(32))
    HexString = Replace(HexString, vbCr, Chr(32))
    mVar = Split(Trim(HexString), Chr(32))
    For Z = 0 To UBound(mVar)
        If Len(mVar(Z)) <> 0 Then z1 = z1 + 1
    Next
    ReDim mHex(0 To z1)
    ReDim mAsc(0 To z1)
    z1 = 0
    For Z = 0 To UBound(mVar)
        If Len(mVar(Z)) <> 0 Then
            mHex(z1) = mVar(Z)
            mAsc(z1) = HexToAsc(mVar(Z))
            z1 = z1 + 1
        End If
    Next
End Sub
Public Sub LoadRawAsc(mSrc As String)
    Dim Z As Long, mVar As Variant
    AscString = Replace(mSrc, vbLf, Chr(32))
    AscString = Replace(AscString, vbCrLf, Chr(32))
    AscString = Replace(AscString, vbCr, Chr(32))
    mVar = Split(AscString, Chr(32))
    ReDim mHex(0 To UBound(mVar))
    ReDim mAsc(0 To UBound(mVar))
    For Z = 0 To UBound(mVar)
        mHex(Z) = Hex$(Asc((mVar(Z))))
        If Len(mHex(Z)) = 1 Then mHex(Z) = "0" + mHex(Z)
        If Len(mHex(Z)) = 0 Then mHex(Z) = "00"
        mAsc(Z) = mVar(Z)
    Next
    HexString = Join(mHex, Chr(32))
    AscString = Join(mAsc, Chr(32))
End Sub
Public Sub LoadRawBin(arrByte As Variant)
    Dim Z As Long, q As Long
    ReDim mHex(0 To UBound(arrByte))
    ReDim mAsc(0 To UBound(arrByte))
    For Z = 0 To UBound(arrByte)
        mHex(Z) = Format(Hex$(arrByte(Z)), "00")
        If Len(mHex(Z)) = 1 Then mHex(Z) = "0" + mHex(Z)
        If Len(mHex(Z)) = 0 Then mHex(Z) = "00"
        mAsc(Z) = Str(arrByte(Z))
        q = Val(mAsc(Z))
        If q < 33 Or (q > 126 And q < 144) Or (q > 147 And q < 161) Then
            mAsc(Z) = Chr(46)
        Else
            mAsc(Z) = Chr(q)
        End If
    Next
    HexString = Join(mHex, Chr(32))
    AscString = Join(mAsc, Chr(32))
End Sub
Public Function DataSplit(mWidth As Long, mSrc As Variant, Optional mSeparator As String = "") As String
    Dim Z As Long, Temp As String
    For Z = 0 To UBound(mSrc)
        If ((Z + 1) Mod mWidth) = 0 Then
            Temp = Temp + mSrc(Z) + vbLf
        Else
            Temp = Temp + mSrc(Z) + mSeparator
        End If
        DoEvents
    Next
    DataSplit = Temp
End Function
Private Function HexToAsc(ByVal HexStr As String) As String
    Dim mult As Double
    Dim DecNum As Long
    Dim ch As String
    Dim tempASC As String
    mult = 1
    DecNum = 0
    Dim i As Integer
    For i = Len(HexStr) To 1 Step -1
        ch = Mid(HexStr, i, 1)
        If (ch >= "0") And (ch <= "9") Then
            DecNum = DecNum + (Val(ch) * mult)
        Else
            If (ch >= "A") And (ch <= "F") Then
                DecNum = DecNum + ((Asc(ch) - Asc("A") + 10) * mult)
            Else
                If (ch >= "a") And (ch <= "f") Then
                    DecNum = DecNum + ((Asc(ch) - Asc("a") + 10) * mult)
                Else
                    HexToAsc = ""
                    Exit Function
                End If
            End If
        End If
        mult = mult * 16
    Next i
    tempASC = Chr(DecNum)
    If DecNum < 33 Or (DecNum > 126 And DecNum < 144) Or (DecNum > 147 And DecNum < 161) Then
        tempASC = Chr(46)
    Else
        tempASC = Chr(DecNum)
    End If
    HexToAsc = tempASC
End Function
Public Sub EditByteByAsc(mStartIndex As Long, mSrc As Variant)
    Dim Z As Long
    If mStartIndex > UBound(mAsc) Then
        Z = mStartIndex
        If (mStartIndex - 1 Mod 2) <> 0 Then Z = Z + 1
        ReDim Preserve mAsc(0 To Z)
        ReDim Preserve mHex(0 To Z)
    End If
    For Z = mStartIndex To mStartIndex + UBound(mSrc)
        mAsc(Z) = mSrc(Z - mStartIndex)
        mHex(Z) = Hex$(Asc((mSrc(Z - mStartIndex))))
        If Len(mHex(Z)) = 1 Then mHex(Z) = "0" + mHex(Z)
        If Len(mHex(Z)) = 0 Then mHex(Z) = "00"
    Next
End Sub
Public Sub EditByteByHex(mStartIndex As Long, mSrc As Variant)
    Dim Z As Long
    If mStartIndex + UBound(mSrc) > UBound(mAsc) Then
        Z = mStartIndex + UBound(mSrc)
        ReDim Preserve mAsc(0 To Z)
        ReDim Preserve mHex(0 To Z)
    End If
    For Z = mStartIndex To mStartIndex + UBound(mSrc)
        mAsc(Z) = HexToAsc(mSrc(Z - mStartIndex))
        mHex(Z) = mSrc(Z - mStartIndex)
        If Len(mHex(Z)) = 1 Then mHex(Z) = "0" + mHex(Z)
        If Len(mHex(Z)) = 0 Then mHex(Z) = "00"
    Next
    HexData = mHex
End Sub
Public Sub RemoveBytes(mStart As Long, mFin As Long)
    Dim NewCnt As Long, tmpBytes() As String, Z As Long, nz As Long
    NewCnt = UBound(mHex) - (mFin - mStart)
    ReDim tmpBytes(0 To NewCnt)
    For Z = 0 To mStart - 1
        tmpBytes(Z) = mHex(Z)
    Next
    For nz = mFin + 1 To UBound(mHex)
        tmpBytes(Z) = mHex(nz)
        Z = Z + 1
    Next
    HexData = tmpBytes
End Sub
Public Sub AddBytesAsBytes(mStart As Long, mInsert As Variant)
    Dim NewCnt As Long, tmpBytes() As String, Z As Long, nz As Long, nzz As Long
    Dim Cnt As Long
    NewCnt = UBound(mHex) + UBound(mInsert) + 1
    ReDim tmpBytes(0 To NewCnt)
    For Z = 0 To mStart - 1
        tmpBytes(Z) = mHex(Z)
    Next
    For nz = mStart To mStart + UBound(mInsert)
        tmpBytes(nz) = mInsert(Cnt)
        Cnt = Cnt + 1
    Next
    For nzz = mStart To UBound(mHex)
        tmpBytes(nz) = mHex(nzz)
        nz = nz + 1
    Next
    HexData = tmpBytes
End Sub
Public Sub AddBytes(mStart As Long, mInsert As String)
    Dim NewCnt As Long, tmpBytes() As String, Z As Long, nz As Long, nzz As Long
    Dim tmpStr As String, Temp As String
    tmpStr = Replace(mInsert, vbLf, "")
    tmpStr = Replace(tmpStr, vbCrLf, "")
    tmpStr = Replace(tmpStr, vbCr, "")
    NewCnt = UBound(mHex) + Len(tmpStr)
    ReDim tmpBytes(0 To NewCnt)
    For Z = 0 To mStart
        tmpBytes(Z) = mHex(Z)
    Next
    For nz = 1 To Len(tmpStr)
        Temp = Mid(tmpStr, nz, 1)
        If Trim(Temp) <> "" Then mHex(Z + nz - 1) = Hex$(Asc(Temp))
        If Len(mHex(Z + nz - 1)) = 1 Then mHex(Z + nz - 1) = "0" + mHex(Z + nz - 1)
        If Len(mHex(Z + nz - 1)) = 0 Then mHex(Z + nz - 1) = "00"
    Next
    For nzz = mStart + 1 To UBound(mHex)
        tmpBytes(Z + nz - 1) = mHex(nzz)
        Z = Z + 1
    Next
    HexData = tmpBytes
End Sub
Public Sub AddEmptyLine()
    Dim tmpBytes() As String, Z As Long
    Dim tmpStr As String, Temp As String
    ReDim tmpBytes(0 To UBound(mHex) + 8)
    For Z = 0 To UBound(mHex)
        tmpBytes(Z) = mHex(Z)
    Next
    For Z = UBound(mHex) + 1 To UBound(mHex) + 8
        tmpBytes(Z) = "00"
    Next
    HexData = tmpBytes
End Sub
Public Function GetBytes(mStart As Long, mEnd As Long) As Variant
    Dim Z As Long, tmpStr() As String, Cnt As Long
    ReDim tmpStr(0 To mEnd - mStart)
    For Z = mStart To mEnd
        If Z > UBound(mHex) Then Exit For
        tmpStr(Cnt) = mHex(Z)
        Cnt = Cnt + 1
    Next
    GetBytes = tmpStr
End Function
Public Function CreateByteList() As String
    Dim bList() As String, Cnt As Long
    ReDim bList(0 To Int(UBound(mHex) / 8) + 1)
    For Z = 0 To UBound(mHex) Step 8
        bList(Cnt) = Hex$(Z)
        If Len(bList(Cnt)) = 1 Then bList(Cnt) = "000" + bList(Cnt)
        If Len(bList(Cnt)) = 2 Then bList(Cnt) = "00" + bList(Cnt)
        If Len(bList(Cnt)) = 3 Then bList(Cnt) = "0" + bList(Cnt)
        If Len(bList(Cnt)) > 4 Then bList(Cnt) = Right(bList(Cnt), 4)
        Cnt = Cnt + 1
    Next Z
    CreateByteList = Join(bList, vbLf)
End Function
