Attribute VB_Name = "modText"
Option Explicit

' Stuffs
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

Private Type VFH
    BitmapWidth As Long
    BitmapHeight As Long
    CellWidth As Long
    CellHeight As Long
    BaseCharOffset As Byte
    CharWidth(0 To 255) As Byte
    CharVA(0 To 255) As CharVA
End Type

Public Type CustomFont
    HeaderInfo As VFH
    Texture As Direct3DTexture8
    RowPitch As Integer
    RowFactor As Single
    ColFactor As Single
    CharHeight As Byte
End Type

' Chat Buffer
Public ChatVA() As TLVERTEX
Public ChatVAS() As TLVERTEX

Public Const ChatTextBufferSize As Integer = 200
Public ChatBufferChunk As Single

'Text buffer
Public Type ChatTextBuffer
    Text As String
    color As Long
End Type

'Chat vertex buffer information
Public ChatArrayUbound As Long
Public ChatVB As Direct3DVertexBuffer8
Public ChatVBS As Direct3DVertexBuffer8
Public ChatTextBuffer(1 To ChatTextBufferSize) As ChatTextBuffer

Public Font_Default As CustomFont
Public Font_Georgia As CustomFont

Public Sub DrawPlayerName(ByVal Index As Long)
Dim textX As Long, textY As Long, Text As String, textSize As Long, Colour As Long
    
    Text = Trim$(GetPlayerName(Index))
    textSize = EngineGetTextWidth(Font_Default, Text)
    
    ' get the colour
    If GetPlayerAccess(Index) > 0 Then
        Colour = Yellow
    Else
        Colour = White
    End If
    
    textX = Player(Index).X * PIC_X + Player(Index).xOffset + (PIC_X \ 2) - (textSize \ 2)
    textY = Player(Index).Y * PIC_Y + Player(Index).yOffset - 32
    
    If GetPlayerSprite(Index) >= 1 And GetPlayerSprite(Index) <= Count_Char Then
        textY = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - (gTexture(Tex_Char(GetPlayerSprite(Index))).RHeight / 4) + 12
    End If
    
    Call RenderText(Font_Default, Text, ConvertMapX(textX), ConvertMapY(textY), Colour)
End Sub

Public Sub DrawNpcName(ByVal Index As Long)
Dim textX As Long, textY As Long, Text As String, textSize As Long, npcNum As Long, Colour As Long
    
    npcNum = MapNpc(Index).Num
    Text = Trim$(Npc(npcNum).name)
    textSize = EngineGetTextWidth(Font_Default, Text)
    
    If Npc(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKWHENATTACKED Then
        ' get the colour
        If Npc(npcNum).Level <= GetPlayerLevel(MyIndex) - 3 Then
            Colour = Grey
        ElseIf Npc(npcNum).Level <= GetPlayerLevel(MyIndex) - 2 Then
            Colour = Green
        ElseIf Npc(npcNum).Level > GetPlayerLevel(MyIndex) Then
            Colour = Red
        Else
            Colour = White
        End If
    Else
        Colour = White
    End If
    
    textX = MapNpc(Index).X * PIC_X + MapNpc(Index).xOffset + (PIC_X \ 2) - (textSize \ 2)
    textY = MapNpc(Index).Y * PIC_Y + MapNpc(Index).yOffset - 32
    
    If Npc(npcNum).Sprite >= 1 And Npc(npcNum).Sprite <= Count_Char Then
        textY = MapNpc(Index).Y * PIC_Y + MapNpc(Index).yOffset - (gTexture(Tex_Char(Npc(npcNum).Sprite)).RHeight / 4) + 12
    End If
    
    Call RenderText(Font_Default, Text, ConvertMapX(textX), ConvertMapY(textY), Colour)
End Sub

Public Sub RenderText(ByRef UseFont As CustomFont, ByVal Text As String, ByVal X As Long, ByVal Y As Long, ByVal color As Long, Optional ByVal Alpha As Long = 255, Optional Shadow As Boolean = True)
Dim TempVA(0 To 3)  As TLVERTEX
Dim TempVAS(0 To 3) As TLVERTEX
Dim TempStr() As String
Dim Count As Integer
Dim Ascii() As Byte
Dim Row As Integer
Dim u As Single
Dim v As Single
Dim I As Long
Dim j As Long
Dim KeyPhrase As Byte
Dim TempColor As Long
Dim ResetColor As Byte
Dim srcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim yOffset As Single

    ' set the color
    color = dx8Colour(color, Alpha)

    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
    
    'Get the text into arrays (split by vbCrLf)
    TempStr = Split(Text, vbCrLf)
    
    'Set the temp color (or else the first character has no color)
    TempColor = color
    
    'Set the texture
    D3DDevice8.SetTexture 0, UseFont.Texture
    CurrentTexture = -1
    
    'Loop through each line if there are line breaks (vbCrLf)
    For I = 0 To UBound(TempStr)
        If Len(TempStr(I)) > 0 Then
            yOffset = I * UseFont.CharHeight
            Count = 0
            'Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(I), vbFromUnicode)
            
            'Loop through the characters
            For j = 1 To Len(TempStr(I))
                'Copy from the cached vertex array to the temp vertex array
                Call CopyMemory(TempVA(0), UseFont.HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), FVF_Size * 4)
                
                'Set up the verticies
                TempVA(0).X = X + Count
                TempVA(0).Y = Y + yOffset
                TempVA(1).X = TempVA(1).X + X + Count
                TempVA(1).Y = TempVA(0).Y
                TempVA(2).X = TempVA(0).X
                TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
                TempVA(3).X = TempVA(1).X
                TempVA(3).Y = TempVA(2).Y
                
                'Set the colors
                TempVA(0).color = TempColor
                TempVA(1).color = TempColor
                TempVA(2).color = TempColor
                TempVA(3).color = TempColor
                
                'Draw the verticies
                Call D3DDevice8.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, TempVA(0), FVF_Size)
                
                'Shift over the the position to render the next character
                Count = Count + UseFont.HeaderInfo.CharWidth(Ascii(j - 1))
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = color
                End If
            Next j
        End If
    Next I
End Sub

Public Function dx8Colour(ByVal colourNum As Long, ByVal Alpha As Long) As Long
    Select Case colourNum
        Case 0 ' Black
            dx8Colour = D3DColorARGB(Alpha, 0, 0, 0)
        Case 1 ' Blue
            dx8Colour = D3DColorARGB(Alpha, 16, 104, 237)
        Case 2 ' Green
            dx8Colour = D3DColorARGB(Alpha, 119, 188, 84)
        Case 3 ' Cyan
            dx8Colour = D3DColorARGB(Alpha, 16, 224, 237)
        Case 4 ' Red
            dx8Colour = D3DColorARGB(Alpha, 201, 0, 0)
        Case 5 ' Magenta
            dx8Colour = D3DColorARGB(Alpha, 255, 0, 255)
        Case 6 ' Brown
            dx8Colour = D3DColorARGB(Alpha, 175, 149, 92)
        Case 7 ' Grey
            dx8Colour = D3DColorARGB(Alpha, 192, 192, 192)
        Case 8 ' DarkGrey
            dx8Colour = D3DColorARGB(Alpha, 128, 128, 128)
        Case 9 ' BrightBlue
            dx8Colour = D3DColorARGB(Alpha, 126, 182, 240)
        Case 10 ' BrightGreen
            dx8Colour = D3DColorARGB(Alpha, 126, 240, 137)
        Case 11 ' BrightCyan
            dx8Colour = D3DColorARGB(Alpha, 157, 242, 242)
        Case 12 ' BrightRed
            dx8Colour = D3DColorARGB(Alpha, 255, 0, 0)
        Case 13 ' Pink
            dx8Colour = D3DColorARGB(Alpha, 255, 118, 221)
        Case 14 ' Yellow
            dx8Colour = D3DColorARGB(Alpha, 255, 255, 0)
        Case 15 ' White
            dx8Colour = D3DColorARGB(Alpha, 255, 255, 255)
        Case 16 ' dark brown
            dx8Colour = D3DColorARGB(Alpha, 98, 84, 52)
    End Select
End Function

Public Function EngineGetTextWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer
Dim LoopI As Integer

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
    
    'Loop through the text
    For LoopI = 1 To Len(Text)
        EngineGetTextWidth = EngineGetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(Mid$(Text, LoopI, 1)))
    Next LoopI

End Function

Sub DrawActionMsg(ByVal Index As Integer)
Dim X As Long, Y As Long, I As Long, Time As Long
Dim LenMsg As Long, Alpha As Long

    If ActionMsg(Index).Message = vbNullString Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(Index).Type
        Case ACTIONMSG_STATIC
            Time = 1500
            
            LenMsg = EngineGetTextWidth(Font_Default, Trim$(ActionMsg(Index).Message))

            If ActionMsg(Index).Y > 0 Then
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - (LenMsg / 2)
                Y = ActionMsg(Index).Y + PIC_Y
            Else
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - (LenMsg / 2)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) + 18
            End If

        Case ACTIONMSG_SCROLL
            Time = 1500
        
            If ActionMsg(Index).Y > 0 Then
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) - 2 - (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            Else
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) + 18 + (ActionMsg(Index).Scroll * 0.001)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            End If
            
            ActionMsg(Index).Alpha = ActionMsg(Index).Alpha - 5
            If ActionMsg(Index).Alpha <= 0 Then ClearActionMsg Index: Exit Sub

        Case ACTIONMSG_SCREEN
            Time = 3000

            ' This will kill any action screen messages that there in the system
            For I = MAX_BYTE To 1 Step -1
                If ActionMsg(I).Type = ACTIONMSG_SCREEN Then
                    If I <> Index Then
                        ClearActionMsg Index
                        Index = I
                    End If
                End If
            Next
    
            X = (400) - ((EngineGetTextWidth(Font_Default, Trim$(ActionMsg(Index).Message)) \ 2))
            Y = 24

    End Select
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    If ActionMsg(Index).created > 0 Then
        RenderText Font_Numbers, ActionMsg(Index).Message, X, Y, ActionMsg(Index).color, ActionMsg(Index).Alpha
    End If

End Sub

Public Sub AddText(ByVal Text As String, ByVal tColor As Long, Optional ByVal Alpha As Long = 255)
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim Size As Long
Dim I As Long
Dim B As Long
Dim color As Long

    color = dx8Colour(tColor, Alpha)
    Text = SwearFilter_Replace(Text)

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(Text, vbCrLf)
    
    For TSLoop = 0 To UBound(TempSplit)
    
        'Clear the values for the new line
        Size = 0
        B = 1
        lastSpace = 1
        
        'Loop through all the characters
        For I = 1 To Len(TempSplit(TSLoop))
        
            'If it is a space, store it so we can easily break at it
            Select Case Mid$(TempSplit(TSLoop), I, 1)
                Case " ": lastSpace = I
                Case "_": lastSpace = I
                Case "-": lastSpace = I
            End Select
            
            'Add up the size
            Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), I, 1)))
            
            'Check for too large of a size
            If Size > ChatWidth Then
                
                'Check if the last space was too far back
                If I - lastSpace > 10 Then
                
                    'Too far away to the last space, so break at the last character
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), B, (I - 1) - B)), color
                    B = I - 1
                    Size = 0
                Else
                    'Break at the last space to preserve the word
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), B, lastSpace - B)), color
                    B = lastSpace + 1
                    'Count all the words we ignored (the ones that weren't printed, but are before "i")
                    Size = EngineGetTextWidth(Font_Default, Mid$(TempSplit(TSLoop), lastSpace, I - lastSpace))
                End If
            End If
            
            'This handles the remainder
            If I = Len(TempSplit(TSLoop)) Then
                If B <> I Then AddToChatTextBuffer_Overflow Mid$(TempSplit(TSLoop), B, I), color
            End If
        Next I
    Next TSLoop
    
    'Only update if we have set up the text (that way we can add to the buffer before it is even made)
    If Font_Default.RowPitch = 0 Then Exit Sub
    
    If ChatScroll > 8 Then ChatScroll = ChatScroll + 1

    'Update the array
    UpdateChatArray
End Sub

Private Sub AddToChatTextBuffer_Overflow(ByVal Text As String, ByVal color As Long)
Dim LoopC As Long

    'Move all other text up
    For LoopC = (ChatTextBufferSize - 1) To 1 Step -1
        ChatTextBuffer(LoopC + 1) = ChatTextBuffer(LoopC)
    Next LoopC
    
    'Set the values
    ChatTextBuffer(1).Text = Text
    ChatTextBuffer(1).color = color
    
    ' set the total chat lines
    totalChatLines = totalChatLines + 1
    If totalChatLines > ChatTextBufferSize - 1 Then totalChatLines = ChatTextBufferSize - 1
End Sub

Public Sub WordWrap_Array(ByVal Text As String, ByVal MaxLineLen As Long, ByRef theArray() As String)
Dim lineCount As Long, I As Long, Size As Long, lastSpace As Long, B As Long
    
    'Too small of text
    If Len(Text) < 2 Then
        ReDim theArray(1 To 1) As String
        theArray(1) = Text
        Exit Sub
    End If
    
    ' default values
    B = 1
    lastSpace = 1
    Size = 0
    
    For I = 1 To Len(Text)
        ' if it's a space, store it
        Select Case Mid$(Text, I, 1)
            Case " ": lastSpace = I
            Case "_": lastSpace = I
            Case "-": lastSpace = I
        End Select
        
        'Add up the size
        Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(Text, I, 1)))
        
        'Check for too large of a size
        If Size > MaxLineLen Then
            'Check if the last space was too far back
            If I - lastSpace > 12 Then
                'Too far away to the last space, so break at the last character
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(Text, B, (I - 1) - B))
                B = I - 1
                Size = 0
            Else
                'Break at the last space to preserve the word
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(Text, B, lastSpace - B))
                B = lastSpace + 1
                
                'Count all the words we ignored (the ones that weren't printed, but are before "i")
                Size = EngineGetTextWidth(Font_Default, Mid$(Text, lastSpace, I - lastSpace))
            End If
        End If
        
        ' Remainder
        If I = Len(Text) Then
            If B <> I Then
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = theArray(lineCount) & Mid$(Text, B, I)
            End If
        End If
    Next
End Sub

Public Function WordWrap(ByVal Text As String, ByVal MaxLineLen As Integer) As String
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim Size As Long
Dim I As Long
Dim B As Long

    'Too small of text
    If Len(Text) < 2 Then
        WordWrap = Text
        Exit Function
    End If

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(Text, vbNewLine)
    
    For TSLoop = 0 To UBound(TempSplit)
    
        'Clear the values for the new line
        Size = 0
        B = 1
        lastSpace = 1
        
        'Add back in the vbNewLines
        If TSLoop < UBound(TempSplit()) Then TempSplit(TSLoop) = TempSplit(TSLoop) & vbNewLine
        
        'Only check lines with a space
        If InStr(1, TempSplit(TSLoop), " ") Or InStr(1, TempSplit(TSLoop), "-") Or InStr(1, TempSplit(TSLoop), "_") Then
            
            'Loop through all the characters
            For I = 1 To Len(TempSplit(TSLoop))
            
                'If it is a space, store it so we can easily break at it
                Select Case Mid$(TempSplit(TSLoop), I, 1)
                    Case " ": lastSpace = I
                    Case "_": lastSpace = I
                    Case "-": lastSpace = I
                End Select
    
                'Add up the size
                Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), I, 1)))
 
                'Check for too large of a size
                If Size > MaxLineLen Then
                    'Check if the last space was too far back
                    If I - lastSpace > 12 Then
                        'Too far away to the last space, so break at the last character
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, (I - 1) - B)) & vbNewLine
                        B = I - 1
                        Size = 0
                    Else
                        'Break at the last space to preserve the word
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, lastSpace - B)) & vbNewLine
                        B = lastSpace + 1
                        
                        'Count all the words we ignored (the ones that weren't printed, but are before "i")
                        Size = EngineGetTextWidth(Font_Default, Mid$(TempSplit(TSLoop), lastSpace, I - lastSpace))
                    End If
                End If
                
                'This handles the remainder
                If I = Len(TempSplit(TSLoop)) Then
                    If B <> I Then
                        WordWrap = WordWrap & Mid$(TempSplit(TSLoop), B, I)
                    End If
                End If
            Next I
        Else
            WordWrap = WordWrap & TempSplit(TSLoop)
        End If
    Next TSLoop
End Function

Public Sub UpdateShowChatText()
Dim CHATOFFSET As Long, I As Long, X As Long

    CHATOFFSET = 52
    
    If EngineGetTextWidth(Font_Default, MyText) > GUIWindow(GUI_CHAT).Width - CHATOFFSET Then
        For I = Len(MyText) To 1 Step -1
            X = X + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(MyText, I, 1)))
            If X > GUIWindow(GUI_CHAT).Width - CHATOFFSET Then
                RenderChatText = Right$(MyText, Len(MyText) - I + 1)
                Exit For
            End If
        Next
    Else
        RenderChatText = MyText
    End If
End Sub

Public Sub LoadFontHeader(ByRef theFont As CustomFont, ByVal FileName As String)
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single

    'Load the header information
    FileNum = FreeFile
    Open App.path & Path_Font & FileName For Binary As #FileNum
        Get #FileNum, , theFont.HeaderInfo
    Close #FileNum
    
    'Calculate some common values
    theFont.CharHeight = theFont.HeaderInfo.CellHeight - 4
    theFont.RowPitch = theFont.HeaderInfo.BitmapWidth \ theFont.HeaderInfo.CellWidth
    theFont.ColFactor = theFont.HeaderInfo.CellWidth / theFont.HeaderInfo.BitmapWidth
    theFont.RowFactor = theFont.HeaderInfo.CellHeight / theFont.HeaderInfo.BitmapHeight
    
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - theFont.HeaderInfo.BaseCharOffset) \ theFont.RowPitch
        u = ((LoopChar - theFont.HeaderInfo.BaseCharOffset) - (Row * theFont.RowPitch)) * theFont.ColFactor
        v = Row * theFont.RowFactor
        
        'Set the verticies
        With theFont.HeaderInfo.CharVA(LoopChar)
            .Vertex(0).color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).RHW = 1
            .Vertex(0).tu = u
            .Vertex(0).tv = v
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).z = 0
            .Vertex(1).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).tu = u + theFont.ColFactor
            .Vertex(1).tv = v
            .Vertex(1).X = theFont.HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).z = 0
            .Vertex(2).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = v + theFont.RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = theFont.HeaderInfo.CellHeight
            .Vertex(2).z = 0
            .Vertex(3).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).tu = u + theFont.ColFactor
            .Vertex(3).tv = v + theFont.RowFactor
            .Vertex(3).X = theFont.HeaderInfo.CellWidth
            .Vertex(3).Y = theFont.HeaderInfo.CellHeight
            .Vertex(3).z = 0
        End With
    Next LoopChar
End Sub
' Chat Box

Public Sub UpdateChatArray()
Dim Chunk As Integer
Dim Count As Integer
Dim LoopC As Byte
Dim Ascii As Byte
Dim Row As Long
Dim Pos As Long
Dim u As Single
Dim v As Single
Dim X As Single
Dim Y As Single
Dim Y2 As Single
Dim I As Long
Dim j As Long
Dim Size As Integer
Dim KeyPhrase As Byte
Dim ResetColor As Byte
Dim TempColor As Long
Dim yOffset As Long

    ' set the offset of each line
    yOffset = 14

    'Set the position
    If ChatBufferChunk <= 1 Then ChatBufferChunk = 1
    
    Chunk = ChatScroll
    
    'Get the number of characters in all the visible buffer
    Size = 0
    
    For LoopC = (Chunk * ChatBufferChunk) - (8 - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        Size = Size + Len(ChatTextBuffer(LoopC).Text)
    Next
    
    Size = Size - j
    ChatArrayUbound = Size * 6 - 1
    If ChatArrayUbound < 0 Then Exit Sub
    ReDim ChatVA(0 To ChatArrayUbound) 'Size our array to fix the 6 verticies of each character
    ReDim ChatVAS(0 To ChatArrayUbound)
    
    'Set the base position
    X = GUIWindow(GUI_CHAT).X + ChatOffsetX
    Y = GUIWindow(GUI_CHAT).Y + ChatOffsetY

    'Loop through each buffer string
    For LoopC = (Chunk * ChatBufferChunk) - (8 - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        If ChatBufferChunk * Chunk > ChatTextBufferSize Then ChatBufferChunk = ChatBufferChunk - 1
        
        'Set the temp color
        TempColor = ChatTextBuffer(LoopC).color
        
        'Set the Y position to be used
        Y2 = Y - (LoopC * yOffset) + (Chunk * ChatBufferChunk * yOffset) - 32
        
        'Loop through each line if there are line breaks (vbCrLf)
        Count = 0   'Counts the offset value we are on
        If LenB(ChatTextBuffer(LoopC).Text) <> 0 Then  'Dont bother with empty strings
            
            'Loop through the characters
            For j = 1 To Len(ChatTextBuffer(LoopC).Text)
            
                'Convert the character to the ascii value
                Ascii = Asc(Mid$(ChatTextBuffer(LoopC).Text, j, 1))
                
                'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
                Row = (Ascii - Font_Default.HeaderInfo.BaseCharOffset) \ Font_Default.RowPitch
                u = ((Ascii - Font_Default.HeaderInfo.BaseCharOffset) - (Row * Font_Default.RowPitch)) * Font_Default.ColFactor
                v = Row * Font_Default.RowFactor

                ' ****** Rectangle | Top Left ******
                With ChatVA(0 + (6 * Pos))
                    .color = TempColor
                    .X = (X) + Count
                    .Y = (Y2)
                    .tu = u
                    .tv = v
                    .RHW = 1
                End With
                
                ' ****** Rectangle | Bottom Left ******
                With ChatVA(1 + (6 * Pos))
                    .color = TempColor
                    .X = (X) + Count
                    .Y = (Y2) + Font_Default.HeaderInfo.CellHeight
                    .tu = u
                    .tv = v + Font_Default.RowFactor
                    .RHW = 1
                End With
                
                ' ****** Rectangle | Bottom Right ******
                With ChatVA(2 + (6 * Pos))
                    .color = TempColor
                    .X = (X) + Count + Font_Default.HeaderInfo.CellWidth
                    .Y = (Y2) + Font_Default.HeaderInfo.CellHeight
                    .tu = u + Font_Default.ColFactor
                    .tv = v + Font_Default.RowFactor
                    .RHW = 1
                End With
                
                
                'Triangle 2 (only one new vertice is needed)
                ChatVA(3 + (6 * Pos)) = ChatVA(0 + (6 * Pos)) 'Top-left corner
                
                ' ****** Rectangle | Top Right ******
                With ChatVA(4 + (6 * Pos))
                    .color = TempColor
                    .X = (X) + Count + Font_Default.HeaderInfo.CellWidth
                    .Y = (Y2)
                    .tu = u + Font_Default.ColFactor
                    .tv = v
                    .RHW = 1
                End With

                ChatVA(5 + (6 * Pos)) = ChatVA(2 + (6 * Pos))

                'Update the character we are on
                Pos = Pos + 1

                'Shift over the the position to render the next character
                Count = Count + Font_Default.HeaderInfo.CharWidth(Ascii)
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = ChatTextBuffer(LoopC).color
                End If
            Next
        End If
    Next LoopC
        
    If Not D3DDevice8 Is Nothing Then   'Make sure the D3DDevice exists - this will only return false if we received messages before it had time to load
        Set ChatVBS = D3DDevice8.CreateVertexBuffer(FVF_Size * Pos * 6, 0, FVF, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVBS, 0, FVF_Size * Pos * 6, 0, ChatVAS(0)
        Set ChatVB = D3DDevice8.CreateVertexBuffer(FVF_Size * Pos * 6, 0, FVF, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVB, 0, FVF_Size * Pos * 6, 0, ChatVA(0)
    End If
    Erase ChatVAS()
    Erase ChatVA()
    
End Sub
Public Sub RenderChatTextBuffer()
Dim srcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim I As Long

    'Clear the LastTexture, letting the rest of the engine know that the texture needs to be changed for next rect render
    D3DDevice8.SetTexture 0, Font_Default.Texture
    CurrentTexture = -1

    If ChatArrayUbound > 0 Then
        D3DDevice8.SetStreamSource 0, ChatVBS, FVF_Size
        D3DDevice8.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
        D3DDevice8.SetStreamSource 0, ChatVB, FVF_Size
        D3DDevice8.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
    End If
    
End Sub

Public Function SwearFilter_Replace(ByVal Message As String) As String
    Dim I As Long

    ' Check to see if there are any swear words in memory.
    If MaxSwearWords = 0 Then
        SwearFilter_Replace = Message
        Exit Function
    End If

    ' Loop through all of the words.
    For I = 1 To MaxSwearWords
        ' Check if the word exists in the sentence.
        If InStr(LCase(Message), LCase(SwearFilter(I).BadWord)) Then
            ' Replace the bad words with the replacement words.
            Message = Replace$(LCase(Message), LCase(SwearFilter(I).BadWord), SwearFilter(I).NewWord, 1, -1, vbTextCompare)
        End If
    Next I

    ' Return the filtered word message.
    SwearFilter_Replace = Message
End Function
