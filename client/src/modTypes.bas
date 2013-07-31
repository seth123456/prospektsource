Attribute VB_Name = "modTypes"
Option Explicit

    

' Public data structures
Public Map As MapRec
Public Bank As BankRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Effect(1 To MAX_EFFECTS) As EffectRec
Public Events(1 To MAX_EVENTS) As EventWrapperRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String

Public Class() As ClassRec
Public GameTime As TimeRec
Public SwearFilter() As SwearFilterRec


' client-side stuff
Public MapResource() As MapResourceRec
Public Hotbar(1 To MAX_HOTBAR) As HotbarRec
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public Blood(1 To MAX_BYTE) As BloodRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec
Public Party As PartyRec
Public PlayerInv(1 To MAX_INV) As PlayerInvRec   ' Inventory
Public GUIWindow() As GUIWindowRec
Public Buttons(1 To Button_Count) As ButtonRec
Public Autotile() As AutotileRec
Public CurrentEvent As SubEventRec
Public ProjectileList() As ProjectileRec
Public EffectData(1 To MAX_BYTE) As Effect   'List of all the active effects
Public WeatherParticle(1 To MAX_WEATHER_PARTICLES) As WeatherParticleRec
Public EditorEffectData As Effect
Public MapSounds() As MapSoundRec
Public chatBubble(1 To MAX_BYTE) As ChatBubbleRec
Public TradeYourOffer(1 To MAX_INV) As PlayerInvRec
Public TradeTheirOffer(1 To MAX_INV) As PlayerInvRec
Public autoInner(1 To 4) As PointRec
Public autoNW(1 To 4) As PointRec
Public autoNE(1 To 4) As PointRec
Public autoSW(1 To 4) As PointRec
Public autoSE(1 To 4) As PointRec

Public Options As OptionsRec

' Type recs
Private Type OptionsRec
    Game_Name As String
    savePass As Byte
    Password As String * NAME_LENGTH
    Username As String * ACCOUNT_LENGTH
    IP As String
    Port As Long
    MenuMusic As String
    Music As Byte
    sound As Byte
    Debug As Byte
    noAuto As Byte
    render As Byte
    Fullscreen As Byte
End Type

Private Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Private Type PlayerInvRec
    Num As Long
    Value As Long
    bound As Byte
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Private Type SpellAnim
    spellnum As Long
    timer As Long
    FramePointer As Long
End Type

Private Type PlayerRec
    ' General
    name As String
    Class As Long
    Sprite As Long
    Level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    MaxVital(1 To Vitals.Vital_Count - 1) As Long
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    ' Position
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    EventOpen(1 To MAX_EVENTS) As Byte
    EventGraphic(1 To MAX_EVENTS) As Byte
    Threshold As Byte
    ' Client use only
    xOffset As Integer
    yOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    Step As Byte
    Anim As Long
    AnimTimer As Long
End Type

Private Type TileDataRec
    X As Long
    Y As Long
    Tileset As Long
End Type

Private Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    DirBlock As Byte
End Type

Private Type MapRec
    name As String * NAME_LENGTH
    Music As String * MUSIC_LENGTH
    BGS As String * MUSIC_LENGTH
    
    Revision As Long
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
    BossNpc As Long
    
    Fog As Byte
    FogSpeed As Byte
    FogOpacity As Byte
    
    Red As Byte
    Green As Byte
    Blue As Byte
    Alpha As Byte
    
    Panorama As Byte
    
    Weather As Long
    WeatherIntensity As Long
End Type

Private Type ClassRec
    name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
End Type

Private Type ItemRec
    name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH
    
    Pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Rarity As Byte
    Speed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Byte
    Animation As Long
    Paperdoll As Long
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    Projectile As Long
    Range As Byte
    Rotation As Integer
    Ammo As Long
    isTwoHanded As Byte
    Stackable As Byte
    Effect As Long
End Type

Private Type MapItemRec
    playerName As String
    Num As Long
    Value As Long
    Frame As Byte
    X As Byte
    Y As Byte
    bound As Boolean
End Type

Private Type NpcRec
    name As String * NAME_LENGTH
    AttackSay As String * 100
    sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    Stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    EXP As Long
    Animation As Long
    Damage As Long
    Level As Long
    
    ' Npc drops
    DropChance(1 To MAX_NPC_DROPS) As Double
    DropItem(1 To MAX_NPC_DROPS) As Byte
    DropItemValue(1 To MAX_NPC_DROPS) As Integer
    
    ' Casting
    Spell(1 To MAX_NPC_SPELLS) As Long
    Event As Long
    Projectile As Long
    ProjectileRange As Byte
    Rotation As Integer
    Moral As Byte
    Effect As Long
End Type

Private Type MapNpcRec
    Num As Long
    target As Long
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    ' Client use only
    xOffset As Long
    yOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
    Anim As Long
    AnimTimer As Long
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem As Long
    CostValue As Long
End Type

Private Type ShopRec
    name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH
    
    Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    X As Long
    Y As Long
    Dir As Byte
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    Vital(1 To Vitals.Vital_Count - 1) As Long
    VitalType(1 To Vitals.Vital_Count - 1) As Byte
    Effect As Long
End Type

Private Type MapResourceRec
    X As Long
    Y As Long
    ResourceState As Byte
End Type

Private Type ResourceRec
    name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    
    ResourceType As Byte
    ResourceImage As Long
    ExhaustedImage As Long
    ItemReward As Long
    ToolRequired As Long
    health As Long
    RespawnTime As Long
    Animation As Long
    Effect As Long
End Type

Private Type ActionMsgRec
    Message As String
    created As Long
    Type As Long
    color As Long
    Scroll As Long
    X As Long
    Y As Long
    timer As Long
    Alpha As Long
End Type

Private Type BloodRec
    Sprite As Long
    timer As Long
    X As Long
    Y As Long
    Alpha As Byte
End Type

Private Type AnimationRec
    name As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
End Type

Private Type AnimInstanceRec
    Animation As Long
    X As Long
    Y As Long
    ' used for locking to players/npcs
    LockIndex As Long
    LockType As Byte
    ' timing
    timer(0 To 1) As Long
    ' rendering check
    Used(0 To 1) As Boolean
    ' counting the loop
    LoopIndex(0 To 1) As Long
    frameIndex(0 To 1) As Long
End Type

Private Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Private Type ButtonRec
    state As Byte
    X As Long
    Y As Long
    Width As Long
    Height As Long
    visible As Boolean
    PicNum As Long
End Type

Private Type GUIWindowRec
    X As Long
    Y As Long
    Width As Long
    Height As Long
    visible As Boolean
End Type

Private Type PointRec
    X As Long
    Y As Long
End Type

Private Type QuarterTileRec
    QuarterTile(1 To 4) As PointRec
    renderState As Byte
    srcX(1 To 4) As Long
    srcY(1 To 4) As Long
End Type

Private Type AutotileRec
    Layer(1 To MapLayer.Layer_Count - 1) As QuarterTileRec
End Type

Private Type ChatBubbleRec
    Msg As String
    Colour As Long
    target As Long
    TargetType As Byte
    timer As Long
    active As Boolean
End Type

Private Type SubEventRec
    Type As EventType
    HasText As Boolean
    Text() As String
    HasData As Boolean
    Data() As Long
End Type

Private Type EventWrapperRec
    name As String
    chkSwitch As Byte
    chkVariable As Byte
    chkHasItem As Byte
    
    SwitchIndex As Long
    SwitchCompare As Byte
    VariableIndex As Long
    VariableCompare As Byte
    VariableCondition As Long
    HasItemIndex As Long
    
    HasSubEvents As Boolean
    SubEvents() As SubEventRec
    
    Trigger As Byte
    WalkThrought As Byte
    Animated As Byte
    Graphic(0 To 2) As Long
    Layer As Byte
End Type

Private Type ProjectileRec
    X As Long
    Y As Long
    tx As Long
    ty As Long
    RotateSpeed As Byte
    Rotate As Single
    Graphic As Long
End Type

Private Type TimeRec
     Minute As Byte
     Hour As Byte
     Day As Byte
     Month As Byte
     Year As Long
End Type

Private Type EffectRec
    name As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    isMulti As Byte
    MultiParticle(1 To MAX_MULTIPARTICLE) As Long
    Type As Long
    Sprite As Long
    Particles As Long
    Size As Single
    Alpha As Single
    Decay As Single
    Red As Single
    Green As Single
    Blue As Single
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Duration As Single
    XSpeed As Single
    YSpeed As Single
    XAcc As Single
    YAcc As Single
    Modifier As Byte
End Type

Private Type Effect
    X As Single                 'Location of effect
    Y As Single
    GoToX As Single             'Location to move to
    GoToY As Single
    KillWhenAtTarget As Boolean     'If the effect is at its target (GoToX/Y), then Progression is set to 0
    KillWhenTargetLost As Boolean   'Kill the effect if the target is lost (sets progression = 0)
    Gfx As Byte                 'Particle texture used
    Used As Boolean 'If the effect is in use
    Alpha As Single
    Decay As Single
    Red As Single
    Green As Single
    Blue As Single
    XSpeed As Single
    YSpeed As Single
    XAcc As Single
    YAcc As Single
    EffectNum As Byte           'What number of effect that is used
    Modifier As Integer         'Misc variable (depends on the effect)
    FloatSize As Long           'The size of the particles
    Particles() As clsParticle  'Information on each particle
    Progression As Single       'Progression state, best to design where 0 = effect ends
    PartVertex() As TLVERTEX    'Used to point render particles ' Cant use in .NET maybe change
    PreviousFrame As Long       'Tick time of the last frame
    ParticleCount As Integer    'Number of particles total
    ParticlesLeft As Integer    'Number of particles left - only for non-repetitive effects
    BindType As Byte
    BindIndex As Long       'Setting this value will bind the effect to move towards the character
    BindSpeed As Single         'How fast the effect moves towards the character
End Type

Private Type WeatherParticleRec
    Type As Long
    X As Long
    Y As Long
    Velocity As Long
    InUse As Long
End Type

Private Type MapSoundRec
    X As Long
    Y As Long
    SoundHandle As Long
    InUse As Boolean
    Channel As Long
End Type

Private Type SwearFilterRec
    BadWord As String
    NewWord As String
End Type
