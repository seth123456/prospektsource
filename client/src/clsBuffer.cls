VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private buffer() As Byte
Private BufferSize As Long
Private WriteHead As Long
Private ReadHead As Long

Private Sub Class_Initialize()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Flush
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Class_Initialize", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub PreAllocate(ByVal nLength As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    WriteHead = 0
    ReadHead = 0
    BufferSize = nLength - 1
    ReDim buffer(0 To BufferSize)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "PreAllocate", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub Allocate(ByVal nLength As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If BufferSize = 0 And nLength > 1 Then nLength = nLength - 1
    BufferSize = BufferSize + nLength
    ReDim Preserve buffer(0 To BufferSize)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Allocate", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub Flush()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    WriteHead = 0
    ReadHead = 0
    BufferSize = 0
    ReDim buffer(0)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Flush", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub Trim()
Dim c As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' If the readhead is past the buffersize, this means everything has been read in the array, flush it
    If ReadHead >= Count Then
        Flush
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Trim", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub WriteByte(ByVal nByte As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If WriteHead > BufferSize Then Allocate 1
 
    buffer(WriteHead) = nByte
    WriteHead = WriteHead + 1
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "WriteByte", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub WriteBytes(ByRef nByte() As Byte)
Dim nLength As Long
        
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    nLength = (UBound(nByte) - LBound(nByte)) + 1
    
    If WriteHead + nLength - 1 > BufferSize Then Allocate nLength
    
    CopyMemory buffer(WriteHead), nByte(0), nLength
    WriteHead = WriteHead + nLength
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "WriteBytes", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub WriteInteger(ByVal nInteger As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If WriteHead + 1 > BufferSize Then Allocate 2
        
    CopyMemory buffer(WriteHead), nInteger, 2
    WriteHead = WriteHead + 2
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "WriteInteger", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub WriteLong(ByVal nLong As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If WriteHead + 3 > BufferSize Then Allocate 4
    
    CopyMemory buffer(WriteHead), nLong, 4
    WriteHead = WriteHead + 4
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "WriteLong", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub WriteString(ByRef nString As String)
Dim sBytes() As Byte
Dim sLength As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    sLength = Len(nString)
    sBytes = StrConv(nString, vbFromUnicode)
    
    WriteLong sLength
    
    If sLength <= 0 Then Exit Sub
    
    If WriteHead + sLength - 1 > BufferSize Then Allocate sLength
      
    CopyMemory buffer(WriteHead), sBytes(0), sLength
    WriteHead = WriteHead + sLength
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "WriteString", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Function ReadByte(Optional MoveReadHead As Boolean = True) As Byte
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ReadHead > BufferSize Then Exit Function
    
    ReadByte = buffer(ReadHead)
    If MoveReadHead Then ReadHead = ReadHead + 1
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "ReadByte", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function ReadBytes(ByVal nLength As Long, Optional MoveReadHead As Boolean = True) As Byte()
Dim Data() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If nLength = 0 Then Exit Function
    If ReadHead + nLength - 1 > BufferSize Then Exit Function
        
    ReDim Data(nLength - 1)
        
    CopyMemory Data(0), buffer(ReadHead), nLength
    If MoveReadHead Then ReadHead = ReadHead + nLength
    
    ReadBytes = Data
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "ReadBytes", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function ReadInteger(Optional MoveReadHead As Boolean = True) As Integer
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ReadHead + 1 > BufferSize Then Exit Function
    
    CopyMemory ReadInteger, buffer(ReadHead), 2
    If MoveReadHead Then ReadHead = ReadHead + 2
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "ReadInteger", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function ReadLong(Optional MoveReadHead As Boolean = True) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ReadHead + 3 > BufferSize Then Exit Function
    
    CopyMemory ReadLong, buffer(ReadHead), 4
    If MoveReadHead Then ReadHead = ReadHead + 4
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "ReadLong", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function ReadString(Optional MoveReadHead As Boolean = True) As String
Dim sLength As Long
Dim sBytes() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    sLength = ReadLong(False)
    If sLength <= 0 Then
        If MoveReadHead Then ReadHead = ReadHead + 4
        Exit Function
    End If
    
    ReDim sBytes(sLength - 1)
    
    CopyMemory sBytes(0), buffer(ReadHead + 4), sLength
    
    ReadString = StrConv(sBytes, vbUnicode)
    If MoveReadHead Then ReadHead = ReadHead + sLength + 4
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "ReadString", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function Count() As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Count = (UBound(buffer) - LBound(buffer)) + 1
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "Count", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function Length() As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Length = Count - ReadHead
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "Length", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function ToArray() As Byte()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ToArray = buffer()
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "ToArray", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function ToString() As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ToString = StrConv(buffer, vbUnicode)
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "ToString", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function
