Attribute VB_Name = "modEnumerations"
Option Explicit

' The order of the packets must match with the client's packet enumeration

' Packets sent by server to client
Public Enum ServerPackets
    SAlertMsg = 1
    SLoginOk
    SDevLoginOk
    SNewCharClasses
    SClassesData
    SInGame
    SPlayerInv
    SPlayerInvUpdate
    SPlayerWornEq
    SPlayerHp
    SPlayerMp
    SPlayerStats
    SPlayerData
    SPlayerMove
    SNpcMove
    SPlayerDir
    SNpcDir
    SPlayerXY
    SPlayerXYMap
    SAttack
    SNpcAttack
    SCheckForMap
    SMapData
    SMapItemData
    SMapNpcData
    SMapDone
    SGlobalMsg
    SAdminMsg
    SPlayerMsg
    SMapMsg
    SSpawnItem
    SItemEditor
    SUpdateItem
    SREditor
    SSpawnNpc
    SNpcDead
    SNpcEditor
    SUpdateNpc
    SEditMap
    SShopEditor
    SUpdateShop
    SSpellEditor
    SUpdateSpell
    SSpells
    SLeft
    SResourceCache
    SResourceEditor
    SUpdateResource
    SSendPing
    SActionMsg
    SPlayerEXP
    SBlood
    SAnimationEditor
    SUpdateAnimation
    SAnimation
    SMapNpcVitals
    SCooldown
    SClearSpellBuffer
    SSayMsg
    SOpenShop
    SStunned
    SMapWornEq
    SBank
    STrade
    SCloseTrade
    STradeUpdate
    STradeStatus
    STarget
    SHotbar
    SHighIndex
    SSound
    STradeRequest
    SPartyInvite
    SPartyUpdate
    SPartyVitals
    SStartTutorial
    SChatBubble
    SMapReport
    SEventData
    SEventEditor
    SEventUpdate
    SSwitchesAndVariables
    SEventOpen
    SCreateProjectile
    SEventGraphic
    SClientTime
    SPlaySound
    SPlayBGM
    SFadeoutBGM
    SEffectEditor
    SUpdateEffect
    SEffect
    SSpecialEffect
    SThreshold
    SSwearFilter
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CNewAccount = 1
    CLogin
    CDevLogin
    CAddChar
    CUseChar
    CSayMsg
    CEmoteMsg
    CBroadcastMsg
    CPlayerMsg
    CPlayerMove
    CPlayerDir
    CUseItem
    CAttack
    CUseStatPoint
    CWarpMeTo
    CWarpToMe
    CWarpTo
    CSetSprite
    CRequestNewMap
    CMapData
    CNeedMap
    CMapGetItem
    CMapDropItem
    CMapRespawn
    CMapReport
    CKickPlayer
    CBanList
    CBanDestroy
    CBanPlayer
    CRequestEditMap
    CRequestEditItem
    CSaveItem
    CRequestEditNpc
    CSaveNpc
    CRequestEditShop
    CSaveShop
    CRequestEditSpell
    CSaveSpell
    CSetAccess
    CWhosOnline
    CSetMotd
    CTarget
    CSpells
    CCast
    CQuit
    CSwapInvSlots
    CRequestEditResource
    CSaveResource
    CCheckPing
    CUnequip
    CRequestPlayerData
    CRequestItems
    CRequestNPCS
    CRequestResources
    CSpawnItem
    CRequestEditAnimation
    CSaveAnimation
    CRequestAnimations
    CRequestSpells
    CRequestShops
    CRequestLevelUp
    CForgetSpell
    CCloseShop
    CBuyItem
    CSellItem
    CChangeBankSlots
    CDepositItem
    CWithdrawItem
    CCloseBank
    CAdminWarp
    CTradeRequest
    CAcceptTrade
    CDeclineTrade
    CTradeItem
    CUntradeItem
    CHotbarChange
    CHotbarUse
    CSwapSpellSlots
    CAcceptTradeRequest
    CDeclineTradeRequest
    CPartyRequest
    CAcceptParty
    CDeclineParty
    CPartyLeave
    CFinishTutorial
    CSwitchesAndVariables
    CRequestSwitchesAndVariables
    CSaveEventData
    CRequestEditEvents
    CRequestEventData
    CRequestEventsData
    CChooseEventOption
    CRequestEditEffect
    CSaveEffect
    CRequestEffects
    CDevMap
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

Public HandleDataSub(CMSG_COUNT) As Long

' Stats used by Players, Npcs and Classes
Public Enum Stats
    Strength = 1
    Endurance
    Intelligence
    Agility
    Willpower
    ' Make sure Stat_Count is below everything else
    Stat_Count
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    HP = 1
    MP
    ' Make sure Vital_Count is below everything else
    Vital_Count
End Enum

' Equipment used by Players
Public Enum Equipment
    Weapon = 1
    Armor
    Helmet
    shield
    ' Make sure Equipment_Count is below everything else
    Equipment_Count
End Enum

' Layers in a map
Public Enum MapLayer
    Ground = 1
    Mask
    Mask2
    Fringe
    Fringe2
    Roof
    ' Make sure Layer_Count is below everything else
    Layer_Count
End Enum

' Sound entities
Public Enum SoundEntity
    seAnimation = 1
    seItem
    seNpc
    seResource
    seSpell
    seEffect
    ' Make sure SoundEntity_Count is below everything else
    SoundEntity_Count
End Enum

Public Enum EventType
    Evt_Message = 0
    Evt_Menu
    Evt_Quit
    Evt_OpenShop
    Evt_OpenBank
    Evt_GiveItem
    Evt_ChangeLevel
    Evt_PlayAnimation
    Evt_Warp
    Evt_GOTO
    Evt_Switch
    Evt_Variable
    Evt_AddText
    Evt_Chatbubble
    Evt_Branch
    Evt_ChangeSkill
    Evt_ChangeSprite
    Evt_ChangePK
    Evt_ChangeClass
    Evt_ChangeSex
    Evt_ChangeExp
    Evt_SetAccess
    Evt_CustomScript
    Evt_OpenEvent
    Evt_Changegraphic
    Evt_ChangeVitals
    Evt_PlaySound
    Evt_PlayBGM
    Evt_FadeoutBGM
    Evt_SpecialEffect
    'EventType_Count should be below everything else
    EventType_Count
End Enum

Public Enum ComparisonOperator
    GEQUAL = 0
    LEQUAL
    GREATER
    LESS
    EQUAL
    NOTEQUAL
End Enum

Public Enum NPCMoralType
    NPC_MORAL_NONE = 0
    NPC_MORAL_GOOD
    NPC_MORAL_BAD
End Enum

Public Enum SEffectType
    SEFFECT_TYPE_FADEIN = 1
    SEFFECT_TYPE_FADEOUT
    SEFFECT_TYPE_FLASH
    SEFFECT_TYPE_FOG
    SEFFECT_TYPE_WEATHER
    SEFFECT_TYPE_TINT
End Enum
