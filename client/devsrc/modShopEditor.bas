Attribute VB_Name = "modShopEditor"
Option Explicit

Public Shop(1 To MAX_SHOPS) As ShopRec
Public TempShop As ShopRec

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem As Long
    CostValue As Long
End Type

Public Type ShopRec
    name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

' /////////////////
' // Shop Editor //
' /////////////////
Public Sub ShopEditorInit()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If frmEditor_Shop.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Shop.lstIndex.ListIndex + 1

    With frmEditor_Shop
        .txtName.Text = Trim$(Shop(EditorIndex).name)
        If Shop(EditorIndex).BuyRate > 0 Then
            .scrlBuy.Value = Shop(EditorIndex).BuyRate
        Else
            .scrlBuy.Value = 100
        End If
        
        .cmbItem.Clear
        .cmbItem.AddItem "None"
        .cmbCostItem.Clear
        .cmbCostItem.AddItem "None"
        
        For I = 1 To MAX_ITEMS
        .cmbItem.AddItem I & ": " & Trim$(Item(I).name)
        .cmbCostItem.AddItem I & ": " & Trim$(Item(I).name)
        Next
        
        .cmbItem.ListIndex = 0
        .cmbCostItem.ListIndex = 0
        
        .cmbCostItem.ListIndex = Shop(EditorIndex).TradeItem(EditorIndex).CostItem
        .txtCostValue.Text = Shop(EditorIndex).TradeItem(EditorIndex).CostValue
        .cmbItem.ListIndex = Shop(EditorIndex).TradeItem(EditorIndex).Item
        .txtItemValue.Text = Shop(EditorIndex).TradeItem(EditorIndex).ItemValue
    End With
    
    UpdateShopTrade
    
    Shop_Changed(EditorIndex) = True

' Error handler
Exit Sub
errorhandler:
HandleError "ShopEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
Err.Clear
Exit Sub
End Sub

Public Sub UpdateShopTrade(Optional ByVal tmpPos As Long = 0)
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    frmEditor_Shop.lstTradeItem.Clear

    For I = 1 To MAX_TRADES
        With Shop(EditorIndex).TradeItem(I)
            ' if none, show as none
            If .Item = 0 And .CostItem = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem "Empty Trade Slot"
            Else
                frmEditor_Shop.lstTradeItem.AddItem I & ": " & .ItemValue & "x " & Trim$(Item(.Item).name) & " for " & .CostValue & "x " & Trim$(Item(.CostItem).name)
            End If
        End With
    Next

    frmEditor_Shop.lstTradeItem.ListIndex = tmpPos
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateShopTrade", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorOk()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    For I = 1 To MAX_SHOPS
        If Shop_Changed(I) Then
            Call SendSaveShop(I)
        End If
    Next
    
    Unload frmEditor_Shop
    Editor = 0
    ClearChanged_Shop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorCancel()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Shop
    ClearChanged_Shop
    ClearShops
    SendRequestShops
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Shop()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    ZeroMemory Shop_Changed(1), MAX_SHOPS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Shop", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearShop(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearShops()
Dim I As Long
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    For I = 1 To MAX_SHOPS
        Call ClearShop(I)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
