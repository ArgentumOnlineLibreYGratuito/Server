Attribute VB_Name = "modEngine_Protocol"
'**************************************************************************
' This program is free software; you can redistribute it and/or modify
' it under the terms of the Affero General Public License;
' either version 1 of the License, or any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' Affero General Public License for more details.
'
' You should have received a copy of the Affero General Public License
' along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

Option Explicit


Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
End Enum

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
End Enum


Private Enum ServerPacketID
    logged                                       ' LOGGED
    RemoveDialogs                                ' QTDL
    RemoveCharDialog                             ' QDL
    NavigateToggle                               ' NAVEG
    Disconnect                                   ' FINOK
    CommerceEnd                                  ' FINCOMOK
    BankEnd                                      ' FINBANOK
    CommerceInit                                 ' INITCOM
    BankInit                                     ' INITBANCO
    UserCommerceInit                             ' INITCOMUSU
    UserCommerceEnd                              ' FINCOMUSUOK
    ShowBlacksmithForm                           ' SFH
    ShowCarpenterForm                            ' SFC
    NPCSwing                                     ' N1
    NPCKillUser                                  ' 6
    BlockedWithShieldUser                        ' 7
    BlockedWithShieldOther                       ' 8
    UserSwing                                    ' U1
    SafeModeOn                                   ' SEGON
    SafeModeOff                                  ' SEGOFF
    ResuscitationSafeOn
    ResuscitationSafeOff
    NobilityLost                                 ' PN
    CantUseWhileMeditating                       ' M!
    UpdateSta                                    ' ASS
    UpdateMana                                   ' ASM
    UpdateHP                                     ' ASH
    UpdateGold                                   ' ASG
    UpdateExp                                    ' ASE
    ChangeMap                                    ' CM
    PosUpdate                                    ' PU
    NPCHitUser                                   ' N2
    UserHitNPC                                   ' U2
    UserAttackedSwing                            ' U3
    UserHittedByUser                             ' N4
    UserHittedUser                               ' N5
    ChatOverHead                                 ' ||
    ConsoleMsg                                   ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat                                    ' |+
    ShowMessageBox                               ' !!
    UserIndexInServer                            ' IU
    UserCharIndexInServer                        ' IP
    CharacterCreate                              ' CC
    CharacterRemove                              ' BP
    CharacterMove                                ' MP, +, * and _ '
    CharacterChange                              ' CP
    ObjectCreate                                 ' HO
    ObjectDelete                                 ' BO
    BlockPosition                                ' BQ
    PlayMIDI                                     ' TM
    PlayWave                                     ' TW
    guildList                                    ' GL
    AreaChanged                                  ' CA
    PauseToggle                                  ' BKW
    RainToggle                                   ' LLU
    CreateFX                                     ' CFX
    UpdateUserStats                              ' EST
    WorkRequestTarget                            ' T01
    ChangeInventorySlot                          ' CSI
    ChangeBankSlot                               ' SBO
    ChangeSpellSlot                              ' SHS
    Atributes                                    ' ATR
    BlacksmithWeapons                            ' LAH
    BlacksmithArmors                             ' LAR
    CarpenterObjects                             ' OBR
    RestOK                                       ' DOK
    ErrorMsg                                     ' ERR
    ChangeNPCInventorySlot                       ' NPCI
    UpdateHungerAndThirst                        ' EHYS
    Fame                                         ' FAMA
    MiniStats                                    ' MEST
    LevelUp                                      ' SUNI
    SetInvisible                                 ' NOVER
    DiceRoll                                     ' DADOS
    MeditateToggle                               ' MEDOK
    SendSkills                                   ' SKILLS
    TrainerCreatureList                          ' LSTCRI
    guildNews                                    ' GUILDNE
    OfferDetails                                 ' PEACEDE & ALLIEDE
    AlianceProposalsList                         ' ALLIEPR
    PeaceProposalsList                           ' PEACEPR
    CharacterInfo                                ' CHRINFO
    GuildLeaderInfo                              ' LEADERI
    GuildDetails                                 ' CLANDET
    ShowGuildFundationForm                       ' SHOWFUN
    ParalizeOK                                   ' PARADOK
    ShowUserRequest                              ' PETICIO
    TradeOK                                      ' TRANSOK
    BankOK                                       ' BANCOOK
    ChangeUserTradeSlot                          ' COMUSUINV
    Pong
    UpdateTagAndStatus
    
    'GM messages
    SpawnList                                    ' SPL
    ShowSOSForm                                  ' MSOS
    ShowMOTDEditionForm                          ' ZMOTD
    ShowGMPanelForm                              ' ABPANEL
    UserNameList                                 ' LISTUSU
End Enum

Private Enum ClientPacketID
    LoginExistingChar                            'OLOGIN
    ThrowDices                                   'TIRDAD
    LoginNewChar                                 'NLOGIN
    Talk                                         ';
    Yell                                         '-
    Whisper                                      '\
    Walk                                         'M
    RequestPositionUpdate                        'RPU
    Attack                                       'AT
    PickUp                                       'AG
    CombatModeToggle                             'TAB        - SHOULD BE HANLDED JUST BY THE CLIENT!!
    SafeToggle                                   '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle
    RequestGuildLeaderInfo                       'GLINFO
    RequestAtributes                             'ATR
    RequestFame                                  'FAMA
    RequestSkills                                'ESKI
    RequestMiniStats                             'FEST
    CommerceEnd                                  'FINCOM
    UserCommerceEnd                              'FINCOMUSU
    BankEnd                                      'FINBAN
    UserCommerceOk                               'COMUSUOK
    UserCommerceReject                           'COMUSUNO
    Drop                                         'TI
    CastSpell                                    'LH
    LeftClick                                    'LC
    DoubleClick                                  'RC
    Work                                         'UK
    UseSpellMacro                                'UMH
    UseItem                                      'USA
    CraftBlacksmith                              'CNS
    CraftCarpenter                               'CNC
    WorkLeftClick                                'WLC
    CreateNewGuild                               'CIG
    SpellInfo                                    'INFS
    EquipItem                                    'EQUI
    ChangeHeading                                'CHEA
    ModifySkills                                 'SKSE
    Train                                        'ENTR
    CommerceBuy                                  'COMP
    BankExtractItem                              'RETI
    CommerceSell                                 'VEND
    BankDeposit                                  'DEPO
    MoveSpell                                    'DESPHE
    ClanCodexUpdate                              'DESCOD
    UserCommerceOffer                            'OFRECER
    GuildAcceptPeace                             'ACEPPEAT
    GuildRejectAlliance                          'RECPALIA
    GuildRejectPeace                             'RECPPEAT
    GuildAcceptAlliance                          'ACEPALIA
    GuildOfferPeace                              'PEACEOFF
    GuildOfferAlliance                           'ALLIEOFF
    GuildAllianceDetails                         'ALLIEDET
    GuildPeaceDetails                            'PEACEDET
    GuildRequestJoinerInfo                       'ENVCOMEN
    GuildAlliancePropList                        'ENVALPRO
    GuildPeacePropList                           'ENVPROPP
    GuildDeclareWar                              'DECGUERR
    GuildNewWebsite                              'NEWWEBSI
    GuildAcceptNewMember                         'ACEPTARI
    GuildRejectNewMember                         'RECHAZAR
    GuildKickMember                              'ECHARCLA
    GuildUpdateNews                              'ACTGNEWS
    GuildMemberInfo                              '1HRINFO<
    GuildOpenElections                           'ABREELEC
    GuildRequestMembership                       'SOLICITUD
    GuildRequestDetails                          'CLANDETAILS
    Online                                       '/ONLINE
    Quit                                         '/SALIR
    GuildLeave                                   '/SALIRCLAN
    RequestAccountState                          '/BALANCE
    PetStand                                     '/QUIETO
    PetFollow                                    '/ACOMPAÑAR
    TrainList                                    '/ENTRENAR
    Rest                                         '/DESCANSAR
    Meditate                                     '/MEDITAR
    Resucitate                                   '/RESUCITAR
    Heal                                         '/CURAR
    Help                                         '/AYUDA
    RequestStats                                 '/EST
    CommerceStart                                '/COMERCIAR
    BankStart                                    '/BOVEDA
    Enlist                                       '/ENLISTAR
    Information                                  '/INFORMACION
    Reward                                       '/RECOMPENSA
    RequestMOTD                                  '/MOTD
    UpTime                                       '/UPTIME
    PartyLeave                                   '/SALIRPARTY
    PartyCreate                                  '/CREARPARTY
    PartyJoin                                    '/PARTY
    GuildMessage                                 '/CMSG
    PartyMessage                                 '/PMSG
    GuildOnline                                  '/ONLINECLAN
    PartyOnline                                  '/ONLINEPARTY
    CouncilMessage                               '/BMSG
    RoleMasterRequest                            '/ROL
    GMRequest                                    '/GM
    ChangeDescription                            '/DESC
    GuildVote                                    '/VOTO
    Punishments                                  '/PENAS
    ChangePassword                               '/CONTRASEÑA
    Gamble                                       '/APOSTAR
    LeaveFaction                                 '/RETIRAR ( with no arguments )
    BankExtractGold                              '/RETIRAR ( with arguments )
    BankDepositGold                              '/DEPOSITAR
    Denounce                                     '/DENUNCIAR
    GuildFundate                                 '/FUNDARCLAN
    PartyKick                                    '/ECHARPARTY
    PartySetLeader                               '/PARTYLIDER
    PartyAcceptMember                            '/ACCEPTPARTY
    Ping                                         '/PING
    
    'GM messages
    GMMessage                                    '/GMSG
    showName                                     '/SHOWNAME
    OnlineRoyalArmy                              '/ONLINEREAL
    OnlineChaosLegion                            '/ONLINECAOS
    GoNearby                                     '/IRCERCA
    comment                                      '/REM
    serverTime                                   '/HORA
    Where                                        '/DONDE
    CreaturesInMap                               '/NENE
    WarpMeToTarget                               '/TELEPLOC
    WarpChar                                     '/TELEP
    Silence                                      '/SILENCIAR
    SOSShowList                                  '/SHOW SOS
    SOSRemove                                    'SOSDONE
    GoToChar                                     '/IRA
    invisible                                    '/INVISIBLE
    GMPanel                                      '/PANELGM
    RequestUserList                              'LISTUSU
    Working                                      '/TRABAJANDO
    Hiding                                       '/OCULTANDO
    Jail                                         '/CARCEL
    KillNPC                                      '/RMATA
    WarnUser                                     '/ADVERTENCIA
    EditChar                                     '/MOD
    RequestCharInfo                              '/INFO
    RequestCharStats                             '/STAT
    RequestCharGold                              '/BAL
    RequestCharInventory                         '/INV
    RequestCharBank                              '/BOV
    RequestCharSkills                            '/SKILLS
    ReviveChar                                   '/REVIVIR
    OnlineGM                                     '/ONLINEGM
    OnlineMap                                    '/ONLINEMAP
    Forgive                                      '/PERDON
    Kick                                         '/ECHAR
    Execute                                      '/EJECUTAR
    BanChar                                      '/BAN
    UnbanChar                                    '/UNBAN
    NPCFollow                                    '/SEGUIR
    SummonChar                                   '/SUM
    SpawnListRequest                             '/CC
    SpawnCreature                                'SPA
    ResetNPCInventory                            '/RESETINV
    CleanWorld                                   '/LIMPIAR
    ServerMessage                                '/RMSG
    NickToIP                                     '/NICK2IP
    IPToNick                                     '/IP2NICK
    GuildOnlineMembers                           '/ONCLAN
    TeleportCreate                               '/CT
    TeleportDestroy                              '/DT
    RainToggle                                   '/LLUVIA
    SetCharDescription                           '/SETDESC
    ForceMIDIToMap                               '/FORCEMIDIMAP
    ForceWAVEToMap                               '/FORCEWAVMAP
    RoyalArmyMessage                             '/REALMSG
    ChaosLegionMessage                           '/CAOSMSG
    CitizenMessage                               '/CIUMSG
    CriminalMessage                              '/CRIMSG
    TalkAsNPC                                    '/TALKAS
    DestroyAllItemsInArea                        '/MASSDEST
    AcceptRoyalCouncilMember                     '/ACEPTCONSE
    AcceptChaosCouncilMember                     '/ACEPTCONSECAOS
    ItemsInTheFloor                              '/PISO
    CouncilKick                                  '/KICKCONSE
    SetTrigger                                   '/TRIGGER
    AskTrigger                                   '/TRIGGER with no arguments
    BannedIPList                                 '/BANIPLIST
    BannedIPReload                               '/BANIPRELOAD
    GuildMemberList                              '/MIEMBROSCLAN
    GuildBan                                     '/BANCLAN
    BanIP                                        '/BANIP
    UnbanIP                                      '/UNBANIP
    CreateItem                                   '/CI
    DestroyItems                                 '/DEST
    ChaosLegionKick                              '/NOCAOS
    RoyalArmyKick                                '/NOREAL
    ForceMIDIAll                                 '/FORCEMIDI
    ForceWAVEAll                                 '/FORCEWAV
    RemovePunishment                             '/BORRARPENA
    TileBlockedToggle                            '/BLOQ
    KillNPCNoRespawn                             '/MATA
    KillAllNearbyNPCs                            '/MASSKILL
    LastIP                                       '/LASTIP
    ChangeMOTD                                   '/MOTDCAMBIA
    SetMOTD                                      'ZMOTD
    SystemMessage                                '/SMSG
    CreateNPC                                    '/ACC
    CreateNPCWithRespawn                         '/RACC
    NavigateToggle                               '/NAVE
    ServerOpenToUsersToggle                      '/HABILITAR
    TurnOffServer                                '/APAGAR
    TurnCriminal                                 '/CONDEN
    ResetFactions                                '/RAJAR
    RemoveCharFromGuild                          '/RAJARCLAN
    RequestCharMail                              '/LASTEMAIL
    AlterPassword                                '/APASS
    AlterMail                                    '/AEMAIL
    AlterName                                    '/ANAME
    DoBackUp                                     '/DOBACKUP
    ShowGuildMessages                            '/SHOWCMSG
    SaveMap                                      '/GUARDAMAPA
    ChangeMapInfoPK                              '/MODMAPINFO PK
    ChangeMapInfoBackup                          '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted                      '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic                         '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi                          '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu                          '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand                            '/MODMAPINFO TERRENO
    ChangeMapInfoZone                            '/MODMAPINFO ZONA
    SaveChars                                    '/GRABAR
    CleanSOS                                     '/BORRAR SOS
    ShowServerForm                               '/SHOW INT
    KickAllChars                                 '/ECHARTODOSPJS
    ChatColor                                    '/CHATCOLOR
    Ignored                                      '/IGNORADO
    CheckSlot                                    '/SLOT
End Enum

Private Writer_ As BinaryWriter

Public Sub Initialize()

    Set Writer_ = New BinaryWriter
    
End Sub


Public Sub OnConnect(ByVal Connection As Network_Client)


'==========================================================
'USO DE LA API DE WINSOCK
'========================
    
    Dim NewIndex As Integer
    Dim i As Long
  
    Dim Address As String
    Address = Connection.GetStatistics().Address
    
    If Not SecurityIp.IpSecurityAceptarNuevaConexion(GetLongIp(Address)) Then
        Call Connection.Close(True)
        Exit Sub
    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   BIENVENIDO AL SERVIDOR!!!!!!!!
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Mariano: Baje la busqueda de slot abajo de CondicionSocket y limite x ip
    NewIndex = NextOpenUser ' Nuevo indice
    
    If NewIndex <= MaxUsers Then
        Call Connection.SetAttachment(NewIndex)
        
        UserList(NewIndex).ip = Address
        'Busca si esta banneada la ip
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = UserList(NewIndex).ip Then
                'Call apiclosesocket(NuevoSock)
                Call WriteErrorMsg(NewIndex, "Su IP se encuentra bloqueada en este servidor.")
                Call Connection.Close(False)
                Exit Sub
            End If
        Next i
        
        If NewIndex > LastUser Then LastUser = NewIndex
        
        Set UserList(NewIndex).Connection = Connection
        UserList(NewIndex).ConnIDValida = True
        UserList(NewIndex).ConnID = NewIndex
    Else
    
        Call Connection.Write(PrepareMessageErrorMsg("El server se encuentra lleno en este momento. Disculpe las molestias ocasionadas."))
        Call Writer_.Clear
        Call Connection.Close(False)
        
    End If
    
End Sub

Public Sub OnClose(ByVal Connection As Network_Client)
    Dim UserIndex As Long
    UserIndex = Connection.GetAttachment()
    
    If (UserIndex > 0) Then

        If UserList(UserIndex).flags.UserLogged Then
            Call Cerrar_Usuario(UserIndex)
        Else
            Call CloseSocket(UserIndex)
        End If

        Set UserList(UserIndex).Connection = Nothing
        
    End If
    
End Sub


Public Sub Encode(ByVal Connection As Network_Client, ByVal Message As BinaryReader)

    ' Here goes encode function
    
End Sub

Public Sub Decode(ByVal Connection As Network_Client, ByVal Message As BinaryReader)

    ' Here goes decode function
    
End Sub


''
' Handles incoming data.
'
' @param    userIndex The index of the user sending the message.

Public Sub Handle(ByVal Connection As Network_Client, ByVal Message As BinaryReader)
    Dim UserIndex As Long
    UserIndex = Connection.GetAttachment()
    
    If (UserIndex <= 0) Then
        Exit Sub
    End If
    
    
    Dim PacketID As Long
    PacketID = Message.ReadInt
    
    'Does the packet requires a logged user??
    If Not (PacketID = ClientPacketID.ThrowDices _
      Or PacketID = ClientPacketID.LoginExistingChar _
      Or PacketID = ClientPacketID.LoginNewChar) Then
        
        'Is the user actually logged?
        If Not UserList(UserIndex).flags.UserLogged Then
            Call CloseSocket(UserIndex)
            Exit Sub
        
        'He is logged. Reset idle counter if id is valid.
        Else
            UserList(UserIndex).Counters.IdleCount = 0
        End If
    Else
        UserList(UserIndex).Counters.IdleCount = 0
    End If
    
    Select Case PacketID
        Case ClientPacketID.LoginExistingChar       'OLOGIN
            Call HandleLoginExistingChar(Message, UserIndex)
        
        Case ClientPacketID.ThrowDices              'TIRDAD
            Call HandleThrowDices(Message, UserIndex)
        
        Case ClientPacketID.LoginNewChar            'NLOGIN
            Call HandleLoginNewChar(Message, UserIndex)
        
        Case ClientPacketID.Talk                    ';
            Call HandleTalk(Message, UserIndex)
        
        Case ClientPacketID.Yell                    '-
            Call HandleYell(Message, UserIndex)
        
        Case ClientPacketID.Whisper                 '\
            Call HandleWhisper(Message, UserIndex)
        
        Case ClientPacketID.Walk                    'M
            Call HandleWalk(Message, UserIndex)
        
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(Message, UserIndex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(Message, UserIndex)
        
        Case ClientPacketID.PickUp                  'AG
            Call HandlePickUp(Message, UserIndex)
        
        Case ClientPacketID.CombatModeToggle        'TAB        - SHOULD BE HANLDED JUST BY THE CLIENT!!
            Call HanldeCombatModeToggle(Message, UserIndex)
        
        Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
            Call HandleSafeToggle(Message, UserIndex)
        
        Case ClientPacketID.ResuscitationSafeToggle
            Call HandleResuscitationToggle(Message, UserIndex)
        
        Case ClientPacketID.RequestGuildLeaderInfo  'GLINFO
            Call HandleRequestGuildLeaderInfo(Message, UserIndex)
        
        Case ClientPacketID.RequestAtributes        'ATR
            Call HandleRequestAtributes(Message, UserIndex)
        
        Case ClientPacketID.RequestFame             'FAMA
            Call HandleRequestFame(Message, UserIndex)
        
        Case ClientPacketID.RequestSkills           'ESKI
            Call HandleRequestSkills(Message, UserIndex)
        
        Case ClientPacketID.RequestMiniStats        'FEST
            Call HandleRequestMiniStats(Message, UserIndex)
        
        Case ClientPacketID.CommerceEnd             'FINCOM
            Call HandleCommerceEnd(Message, UserIndex)
        
        Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
            Call HandleUserCommerceEnd(Message, UserIndex)
        
        Case ClientPacketID.BankEnd                 'FINBAN
            Call HandleBankEnd(Message, UserIndex)
        
        Case ClientPacketID.UserCommerceOk          'COMUSUOK
            Call HandleUserCommerceOk(Message, UserIndex)
        
        Case ClientPacketID.UserCommerceReject      'COMUSUNO
            Call HandleUserCommerceReject(Message, UserIndex)
        
        Case ClientPacketID.Drop                    'TI
            Call HandleDrop(Message, UserIndex)
        
        Case ClientPacketID.CastSpell               'LH
            Call HandleCastSpell(Message, UserIndex)
        
        Case ClientPacketID.LeftClick               'LC
            Call HandleLeftClick(Message, UserIndex)
        
        Case ClientPacketID.DoubleClick             'RC
            Call HandleDoubleClick(Message, UserIndex)
        
        Case ClientPacketID.Work                    'UK
            Call HandleWork(Message, UserIndex)
        
        Case ClientPacketID.UseSpellMacro           'UMH
            Call HandleUseSpellMacro(Message, UserIndex)
        
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(Message, UserIndex)
        
        Case ClientPacketID.CraftBlacksmith         'CNS
            Call HandleCraftBlacksmith(Message, UserIndex)
        
        Case ClientPacketID.CraftCarpenter          'CNC
            Call HandleCraftCarpenter(Message, UserIndex)
        
        Case ClientPacketID.WorkLeftClick           'WLC
            Call HandleWorkLeftClick(Message, UserIndex)
        
        Case ClientPacketID.CreateNewGuild          'CIG
            Call HandleCreateNewGuild(Message, UserIndex)
        
        Case ClientPacketID.SpellInfo               'INFS
            Call HandleSpellInfo(Message, UserIndex)
        
        Case ClientPacketID.EquipItem               'EQUI
            Call HandleEquipItem(Message, UserIndex)
        
        Case ClientPacketID.ChangeHeading           'CHEA
            Call HandleChangeHeading(Message, UserIndex)
        
        Case ClientPacketID.ModifySkills            'SKSE
            Call HandleModifySkills(Message, UserIndex)
        
        Case ClientPacketID.Train                   'ENTR
            Call HandleTrain(Message, UserIndex)
        
        Case ClientPacketID.CommerceBuy             'COMP
            Call HandleCommerceBuy(Message, UserIndex)
        
        Case ClientPacketID.BankExtractItem         'RETI
            Call HandleBankExtractItem(Message, UserIndex)
        
        Case ClientPacketID.CommerceSell            'VEND
            Call HandleCommerceSell(Message, UserIndex)
        
        Case ClientPacketID.BankDeposit             'DEPO
            Call HandleBankDeposit(Message, UserIndex)

        Case ClientPacketID.MoveSpell               'DESPHE
            Call HandleMoveSpell(Message, UserIndex)
        
        Case ClientPacketID.ClanCodexUpdate         'DESCOD
            Call HandleClanCodexUpdate(Message, UserIndex)
        
        Case ClientPacketID.UserCommerceOffer       'OFRECER
            Call HandleUserCommerceOffer(Message, UserIndex)
        
        Case ClientPacketID.GuildAcceptPeace        'ACEPPEAT
            Call HandleGuildAcceptPeace(Message, UserIndex)
        
        Case ClientPacketID.GuildRejectAlliance     'RECPALIA
            Call HandleGuildRejectAlliance(Message, UserIndex)
        
        Case ClientPacketID.GuildRejectPeace        'RECPPEAT
            Call HandleGuildRejectPeace(Message, UserIndex)
        
        Case ClientPacketID.GuildAcceptAlliance     'ACEPALIA
            Call HandleGuildAcceptAlliance(Message, UserIndex)
        
        Case ClientPacketID.GuildOfferPeace         'PEACEOFF
            Call HandleGuildOfferPeace(Message, UserIndex)
        
        Case ClientPacketID.GuildOfferAlliance      'ALLIEOFF
            Call HandleGuildOfferAlliance(Message, UserIndex)
        
        Case ClientPacketID.GuildAllianceDetails    'ALLIEDET
            Call HandleGuildAllianceDetails(Message, UserIndex)
        
        Case ClientPacketID.GuildPeaceDetails       'PEACEDET
            Call HandleGuildPeaceDetails(Message, UserIndex)
        
        Case ClientPacketID.GuildRequestJoinerInfo  'ENVCOMEN
            Call HandleGuildRequestJoinerInfo(Message, UserIndex)
        
        Case ClientPacketID.GuildAlliancePropList   'ENVALPRO
            Call HandleGuildAlliancePropList(Message, UserIndex)
        
        Case ClientPacketID.GuildPeacePropList      'ENVPROPP
            Call HandleGuildPeacePropList(Message, UserIndex)
        
        Case ClientPacketID.GuildDeclareWar         'DECGUERR
            Call HandleGuildDeclareWar(Message, UserIndex)
        
        Case ClientPacketID.GuildNewWebsite         'NEWWEBSI
            Call HandleGuildNewWebsite(Message, UserIndex)
        
        Case ClientPacketID.GuildAcceptNewMember    'ACEPTARI
            Call HandleGuildAcceptNewMember(Message, UserIndex)
        
        Case ClientPacketID.GuildRejectNewMember    'RECHAZAR
            Call HandleGuildRejectNewMember(Message, UserIndex)
        
        Case ClientPacketID.GuildKickMember         'ECHARCLA
            Call HandleGuildKickMember(Message, UserIndex)
        
        Case ClientPacketID.GuildUpdateNews         'ACTGNEWS
            Call HandleGuildUpdateNews(Message, UserIndex)
        
        Case ClientPacketID.GuildMemberInfo         '1HRINFO<
            Call HandleGuildMemberInfo(Message, UserIndex)
        
        Case ClientPacketID.GuildOpenElections      'ABREELEC
            Call HandleGuildOpenElections(Message, UserIndex)
        
        Case ClientPacketID.GuildRequestMembership  'SOLICITUD
            Call HandleGuildRequestMembership(Message, UserIndex)
        
        Case ClientPacketID.GuildRequestDetails     'CLANDETAILS
            Call HandleGuildRequestDetails(Message, UserIndex)
                  
        Case ClientPacketID.Online                  '/ONLINE
            Call HandleOnline(Message, UserIndex)
        
        Case ClientPacketID.Quit                    '/SALIR
            Call HandleQuit(Message, UserIndex)
        
        Case ClientPacketID.GuildLeave              '/SALIRCLAN
            Call HandleGuildLeave(Message, UserIndex)
        
        Case ClientPacketID.RequestAccountState     '/BALANCE
            Call HandleRequestAccountState(Message, UserIndex)
        
        Case ClientPacketID.PetStand                '/QUIETO
            Call HandlePetStand(Message, UserIndex)
        
        Case ClientPacketID.PetFollow               '/ACOMPAÑAR
            Call HandlePetFollow(Message, UserIndex)
        
        Case ClientPacketID.TrainList               '/ENTRENAR
            Call HandleTrainList(Message, UserIndex)
        
        Case ClientPacketID.Rest                    '/DESCANSAR
            Call HandleRest(Message, UserIndex)
        
        Case ClientPacketID.Meditate                '/MEDITAR
            Call HandleMeditate(Message, UserIndex)
        
        Case ClientPacketID.Resucitate              '/RESUCITAR
            Call HandleResucitate(Message, UserIndex)
        
        Case ClientPacketID.Heal                    '/CURAR
            Call HandleHeal(Message, UserIndex)
        
        Case ClientPacketID.RequestStats            '/EST
            Call HandleRequestStats(Message, UserIndex)
        
        Case ClientPacketID.CommerceStart           '/COMERCIAR
            Call HandleCommerceStart(Message, UserIndex)
        
        Case ClientPacketID.BankStart               '/BOVEDA
            Call HandleBankStart(Message, UserIndex)
        
        Case ClientPacketID.Enlist                  '/ENLISTAR
            Call HandleEnlist(Message, UserIndex)
        
        Case ClientPacketID.Information             '/INFORMACION
            Call HandleInformation(Message, UserIndex)
        
        Case ClientPacketID.Reward                  '/RECOMPENSA
            Call HandleReward(Message, UserIndex)
        
        Case ClientPacketID.RequestMOTD             '/MOTD
            Call HandleRequestMOTD(Message, UserIndex)
        
        Case ClientPacketID.UpTime                  '/UPTIME
            Call HandleUpTime(Message, UserIndex)
        
        Case ClientPacketID.PartyLeave              '/SALIRPARTY
            Call HandlePartyLeave(Message, UserIndex)
        
        Case ClientPacketID.PartyCreate             '/CREARPARTY
            Call HandlePartyCreate(Message, UserIndex)
        
        Case ClientPacketID.PartyJoin               '/PARTY
            Call HandlePartyJoin(Message, UserIndex)

        Case ClientPacketID.GuildMessage            '/CMSG
            Call HandleGuildMessage(Message, UserIndex)
        
        Case ClientPacketID.PartyMessage            '/PMSG
            Call HandlePartyMessage(Message, UserIndex)

        Case ClientPacketID.GuildOnline             '/ONLINECLAN
            Call HandleGuildOnline(Message, UserIndex)
        
        Case ClientPacketID.PartyOnline             '/ONLINEPARTY
            Call HandlePartyOnline(Message, UserIndex)
        
        Case ClientPacketID.CouncilMessage          '/BMSG
            Call HandleCouncilMessage(Message, UserIndex)
        
        Case ClientPacketID.RoleMasterRequest       '/ROL
            Call HandleRoleMasterRequest(Message, UserIndex)
        
        Case ClientPacketID.GMRequest               '/GM
            Call HandleGMRequest(Message, UserIndex)

        Case ClientPacketID.ChangeDescription       '/DESC
            Call HandleChangeDescription(Message, UserIndex)
        
        Case ClientPacketID.GuildVote               '/VOTO
            Call HandleGuildVote(Message, UserIndex)
        
        Case ClientPacketID.Punishments             '/PENAS
            Call HandlePunishments(Message, UserIndex)
        
        Case ClientPacketID.ChangePassword          '/CONTRASEÑA
            Call HandleChangePassword(Message, UserIndex)
        
        Case ClientPacketID.Gamble                  '/APOSTAR
            Call HandleGamble(Message, UserIndex)
  
        Case ClientPacketID.LeaveFaction            '/RETIRAR ( with no arguments )
            Call HandleLeaveFaction(Message, UserIndex)
        
        Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
            Call HandleBankExtractGold(Message, UserIndex)
        
        Case ClientPacketID.BankDepositGold         '/DEPOSITAR
            Call HandleBankDepositGold(Message, UserIndex)
        
        Case ClientPacketID.Denounce                '/DENUNCIAR
            Call HandleDenounce(Message, UserIndex)
        
        Case ClientPacketID.GuildFundate            '/FUNDARCLAN
            Call HandleGuildFundate(Message, UserIndex)
        
        Case ClientPacketID.PartyKick               '/ECHARPARTY
            Call HandlePartyKick(Message, UserIndex)
        
        Case ClientPacketID.PartySetLeader          '/PARTYLIDER
            Call HandlePartySetLeader(Message, UserIndex)
        
        Case ClientPacketID.PartyAcceptMember       '/ACCEPTPARTY
            Call HandlePartyAcceptMember(Message, UserIndex)
        
        Case ClientPacketID.GuildMemberList         '/MIEMBROSCLAN
            Call HandleGuildMemberList(Message, UserIndex)
        
        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(Message, UserIndex)
        
        
        'GM messages
        Case ClientPacketID.GMMessage               '/GMSG
            Call HandleGMMessage(Message, UserIndex)
        
        Case ClientPacketID.showName                '/SHOWNAME
            Call HandleShowName(Message, UserIndex)
        
        Case ClientPacketID.OnlineRoyalArmy         '/ONLINEREAL
            Call HandleOnlineRoyalArmy(Message, UserIndex)
        
        Case ClientPacketID.OnlineChaosLegion       '/ONLINECAOS
            Call HandleOnlineChaosLegion(Message, UserIndex)
        
        Case ClientPacketID.GoNearby                '/IRCERCA
            Call HandleGoNearby(Message, UserIndex)
        
        Case ClientPacketID.comment                 '/REM
            Call HandleComment(Message, UserIndex)
        
        Case ClientPacketID.serverTime              '/HORA
            Call HandleServerTime(Message, UserIndex)
        
        Case ClientPacketID.Where                   '/DONDE
            Call HandleWhere(Message, UserIndex)
        
        Case ClientPacketID.CreaturesInMap          '/NENE
            Call HandleCreaturesInMap(Message, UserIndex)
        
        Case ClientPacketID.WarpMeToTarget          '/TELEPLOC
            Call HandleWarpMeToTarget(Message, UserIndex)
        
        Case ClientPacketID.WarpChar                '/TELEP
            Call HandleWarpChar(Message, UserIndex)
        
        Case ClientPacketID.Silence                 '/SILENCIAR
            Call HandleSilence(Message, UserIndex)
        
        Case ClientPacketID.SOSShowList             '/SHOW SOS
            Call HandleSOSShowList(Message, UserIndex)
        
        Case ClientPacketID.SOSRemove               'SOSDONE
            Call HandleSOSRemove(Message, UserIndex)
        
        Case ClientPacketID.GoToChar                '/IRA
            Call HandleGoToChar(Message, UserIndex)
        
        Case ClientPacketID.invisible               '/INVISIBLE
            Call HandleInvisible(Message, UserIndex)
        
        Case ClientPacketID.GMPanel                 '/PANELGM
            Call HandleGMPanel(Message, UserIndex)
        
        Case ClientPacketID.RequestUserList         'LISTUSU
            Call HandleRequestUserList(Message, UserIndex)
        
        Case ClientPacketID.Working                 '/TRABAJANDO
            Call HandleWorking(Message, UserIndex)
        
        Case ClientPacketID.Hiding                  '/OCULTANDO
            Call HandleHiding(Message, UserIndex)
        
        Case ClientPacketID.Jail                    '/CARCEL
            Call HandleJail(Message, UserIndex)
        
        Case ClientPacketID.KillNPC                 '/RMATA
            Call HandleKillNPC(Message, UserIndex)
        
        Case ClientPacketID.WarnUser                '/ADVERTENCIA
            Call HandleWarnUser(Message, UserIndex)
        
        Case ClientPacketID.EditChar                '/MOD
            Call HandleEditChar(Message, UserIndex)
            
        Case ClientPacketID.RequestCharInfo         '/INFO
            Call HandleRequestCharInfo(Message, UserIndex)
        
        Case ClientPacketID.RequestCharStats        '/STAT
            Call HandleRequestCharStats(Message, UserIndex)
            
        Case ClientPacketID.RequestCharGold         '/BAL
            Call HandleRequestCharGold(Message, UserIndex)
            
        Case ClientPacketID.RequestCharInventory    '/INV
            Call HandleRequestCharInventory(Message, UserIndex)
            
        Case ClientPacketID.RequestCharBank         '/BOV
            Call HandleRequestCharBank(Message, UserIndex)
        
        Case ClientPacketID.RequestCharSkills       '/SKILLS
            Call HandleRequestCharSkills(Message, UserIndex)
        
        Case ClientPacketID.ReviveChar              '/REVIVIR
            Call HandleReviveChar(Message, UserIndex)
        
        Case ClientPacketID.OnlineGM                '/ONLINEGM
            Call HandleOnlineGM(Message, UserIndex)
        
        Case ClientPacketID.OnlineMap               '/ONLINEMAP
            Call HandleOnlineMap(Message, UserIndex)
        
        Case ClientPacketID.Forgive                 '/PERDON
            Call HandleForgive(Message, UserIndex)
            
        Case ClientPacketID.Kick                    '/ECHAR
            Call HandleKick(Message, UserIndex)
            
        Case ClientPacketID.Execute                 '/EJECUTAR
            Call HandleExecute(Message, UserIndex)
            
        Case ClientPacketID.BanChar                 '/BAN
            Call HandleBanChar(Message, UserIndex)
            
        Case ClientPacketID.UnbanChar               '/UNBAN
            Call HandleUnbanChar(Message, UserIndex)
            
        Case ClientPacketID.NPCFollow               '/SEGUIR
            Call HandleNPCFollow(Message, UserIndex)
            
        Case ClientPacketID.SummonChar              '/SUM
            Call HandleSummonChar(Message, UserIndex)
            
        Case ClientPacketID.SpawnListRequest        '/CC
            Call HandleSpawnListRequest(Message, UserIndex)
            
        Case ClientPacketID.SpawnCreature           'SPA
            Call HandleSpawnCreature(Message, UserIndex)
            
        Case ClientPacketID.ResetNPCInventory       '/RESETINV
            Call HandleResetNPCInventory(Message, UserIndex)
            
        Case ClientPacketID.CleanWorld              '/LIMPIAR
            Call HandleCleanWorld(Message, UserIndex)
            
        Case ClientPacketID.ServerMessage           '/RMSG
            Call HandleServerMessage(Message, UserIndex)
            
        Case ClientPacketID.NickToIP                '/NICK2IP
            Call HandleNickToIP(Message, UserIndex)
        
        Case ClientPacketID.IPToNick                '/IP2NICK
            Call HandleIPToNick(Message, UserIndex)
            
        Case ClientPacketID.GuildOnlineMembers      '/ONCLAN
            Call HandleGuildOnlineMembers(Message, UserIndex)
        
        Case ClientPacketID.TeleportCreate          '/CT
            Call HandleTeleportCreate(Message, UserIndex)
            
        Case ClientPacketID.TeleportDestroy         '/DT
            Call HandleTeleportDestroy(Message, UserIndex)
            
        Case ClientPacketID.RainToggle              '/LLUVIA
            Call HandleRainToggle(Message, UserIndex)
        
        Case ClientPacketID.SetCharDescription      '/SETDESC
            Call HandleSetCharDescription(Message, UserIndex)
        
        Case ClientPacketID.ForceMIDIToMap          '/FORCEMIDIMAP
            Call HanldeForceMIDIToMap(Message, UserIndex)
            
        Case ClientPacketID.ForceWAVEToMap          '/FORCEWAVMAP
            Call HandleForceWAVEToMap(Message, UserIndex)
            
        Case ClientPacketID.RoyalArmyMessage        '/REALMSG
            Call HandleRoyalArmyMessage(Message, UserIndex)
                        
        Case ClientPacketID.ChaosLegionMessage      '/CAOSMSG
            Call HandleChaosLegionMessage(Message, UserIndex)
            
        Case ClientPacketID.CitizenMessage          '/CIUMSG
            Call HandleCitizenMessage(Message, UserIndex)
            
        Case ClientPacketID.CriminalMessage         '/CRIMSG
            Call HandleCriminalMessage(Message, UserIndex)
            
        Case ClientPacketID.TalkAsNPC               '/TALKAS
            Call HandleTalkAsNPC(Message, UserIndex)
        
        Case ClientPacketID.DestroyAllItemsInArea   '/MASSDEST
            Call HandleDestroyAllItemsInArea(Message, UserIndex)
            
        Case ClientPacketID.AcceptRoyalCouncilMember '/ACEPTCONSE
            Call HandleAcceptRoyalCouncilMember(Message, UserIndex)
            
        Case ClientPacketID.AcceptChaosCouncilMember '/ACEPTCONSECAOS
            Call HandleAcceptChaosCouncilMember(Message, UserIndex)
            
        Case ClientPacketID.ItemsInTheFloor         '/PISO
            Call HandleItemsInTheFloor(Message, UserIndex)
            
        Case ClientPacketID.CouncilKick             '/KICKCONSE
            Call HandleCouncilKick(Message, UserIndex)
        
        Case ClientPacketID.SetTrigger              '/TRIGGER
            Call HandleSetTrigger(Message, UserIndex)
        
        Case ClientPacketID.AskTrigger               '/TRIGGER
            Call HandleAskTrigger(Message, UserIndex)
            
        Case ClientPacketID.BannedIPList            '/BANIPLIST
            Call HandleBannedIPList(Message, UserIndex)
        
        Case ClientPacketID.BannedIPReload          '/BANIPRELOAD
            Call HandleBannedIPReload(Message, UserIndex)
        
        Case ClientPacketID.GuildBan                '/BANCLAN
            Call HandleGuildBan(Message, UserIndex)
        
        Case ClientPacketID.BanIP                   '/BANIP
            Call HandleBanIP(Message, UserIndex)
        
        Case ClientPacketID.UnbanIP                 '/UNBANIP
            Call HandleUnbanIP(Message, UserIndex)
        
        Case ClientPacketID.CreateItem              '/CI
            Call HandleCreateItem(Message, UserIndex)
        
        Case ClientPacketID.DestroyItems            '/DEST
            Call HandleDestroyItems(Message, UserIndex)
        
        Case ClientPacketID.ChaosLegionKick         '/NOCAOS
            Call HandleChaosLegionKick(Message, UserIndex)
        
        Case ClientPacketID.RoyalArmyKick           '/NOREAL
            Call HandleRoyalArmyKick(Message, UserIndex)
        
        Case ClientPacketID.ForceMIDIAll            '/FORCEMIDI
            Call HandleForceMIDIAll(Message, UserIndex)
        
        Case ClientPacketID.ForceWAVEAll            '/FORCEWAV
            Call HandleForceWAVEAll(Message, UserIndex)
        
        Case ClientPacketID.RemovePunishment        '/BORRARPENA
            Call HandleRemovePunishment(Message, UserIndex)
        
        Case ClientPacketID.TileBlockedToggle       '/BLOQ
            Call HandleTileBlockedToggle(Message, UserIndex)
        
        Case ClientPacketID.KillNPCNoRespawn        '/MATA
            Call HandleKillNPCNoRespawn(Message, UserIndex)
        
        Case ClientPacketID.KillAllNearbyNPCs       '/MASSKILL
            Call HandleKillAllNearbyNPCs(Message, UserIndex)
        
        Case ClientPacketID.LastIP                  '/LASTIP
            Call HandleLastIP(Message, UserIndex)
        
        Case ClientPacketID.ChangeMOTD              '/MOTDCAMBIA
            Call HandleChangeMOTD(Message, UserIndex)
        
        Case ClientPacketID.SetMOTD                 'ZMOTD
            Call HandleSetMOTD(Message, UserIndex)
        
        Case ClientPacketID.SystemMessage           '/SMSG
            Call HandleSystemMessage(Message, UserIndex)
        
        Case ClientPacketID.CreateNPC               '/ACC
            Call HandleCreateNPC(Message, UserIndex)
        
        Case ClientPacketID.CreateNPCWithRespawn    '/RACC
            Call HandleCreateNPCWithRespawn(Message, UserIndex)

        Case ClientPacketID.NavigateToggle          '/NAVE
            Call HandleNavigateToggle(Message, UserIndex)
        
        Case ClientPacketID.ServerOpenToUsersToggle '/HABILITAR
            Call HandleServerOpenToUsersToggle(Message, UserIndex)
        
        Case ClientPacketID.TurnOffServer           '/APAGAR
            Call HandleTurnOffServer(Message, UserIndex)
        
        Case ClientPacketID.TurnCriminal            '/CONDEN
            Call HandleTurnCriminal(Message, UserIndex)
        
        Case ClientPacketID.ResetFactions           '/RAJAR
            Call HandleResetFactions(Message, UserIndex)
        
        Case ClientPacketID.RemoveCharFromGuild     '/RAJARCLAN
            Call HandleRemoveCharFromGuild(Message, UserIndex)
        
        Case ClientPacketID.RequestCharMail         '/LASTEMAIL
            Call HandleRequestCharMail(Message, UserIndex)
        
        Case ClientPacketID.AlterPassword           '/APASS
            Call HandleAlterPassword(Message, UserIndex)
        
        Case ClientPacketID.AlterMail               '/AEMAIL
            Call HandleAlterMail(Message, UserIndex)
        
        Case ClientPacketID.AlterName               '/ANAME
            Call HandleAlterName(Message, UserIndex)

        Case ClientPacketID.DoBackUp                '/DOBACKUP
            Call HandleDoBackUp(Message, UserIndex)
        
        Case ClientPacketID.ShowGuildMessages       '/SHOWCMSG
            Call HandleShowGuildMessages(Message, UserIndex)
        
        Case ClientPacketID.SaveMap                 '/GUARDAMAPA
            Call HandleSaveMap(Message, UserIndex)
        
        Case ClientPacketID.ChangeMapInfoPK         '/MODMAPINFO PK
            Call HandleChangeMapInfoPK(Message, UserIndex)
        
        Case ClientPacketID.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
            Call HandleChangeMapInfoBackup(Message, UserIndex)
    
        Case ClientPacketID.ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
            Call HandleChangeMapInfoRestricted(Message, UserIndex)
            
        Case ClientPacketID.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
            Call HandleChangeMapInfoNoMagic(Message, UserIndex)
            
        Case ClientPacketID.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
            Call HandleChangeMapInfoNoInvi(Message, UserIndex)
            
        Case ClientPacketID.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
            Call HandleChangeMapInfoNoResu(Message, UserIndex)
            
        Case ClientPacketID.ChangeMapInfoLand       '/MODMAPINFO TERRENO
            Call HandleChangeMapInfoLand(Message, UserIndex)
            
        Case ClientPacketID.ChangeMapInfoZone       '/MODMAPINFO ZONA
            Call HandleChangeMapInfoZone(Message, UserIndex)
        
        Case ClientPacketID.SaveChars               '/GRABAR
            Call HandleSaveChars(Message, UserIndex)
        
        Case ClientPacketID.CleanSOS                '/BORRAR SOS
            Call HandleCleanSOS(Message, UserIndex)
        
        Case ClientPacketID.ShowServerForm          '/SHOW INT
            Call HandleShowServerForm(Message, UserIndex)
    
        Case ClientPacketID.KickAllChars            '/ECHARTODOSPJS
            Call HandleKickAllChars(Message, UserIndex)

        Case ClientPacketID.ChatColor               '/CHATCOLOR
            Call HandleChatColor(Message, UserIndex)
        
        Case ClientPacketID.Ignored                 '/IGNORADO
            Call HandleIgnored(Message, UserIndex)
        
        Case ClientPacketID.CheckSlot               '/SLOT
            Call HandleCheckSlot(Message, UserIndex)
            
        Case Else
            Call CloseSocket(UserIndex)

    End Select

End Sub



''
' Handles the "LoginExistingChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    Dim UserName As String
    Dim Password As String
    Dim version As String
    
    UserName = Message.ReadString16()
    Password = Message.ReadString16()
    
    'Convert version number to string
    version = CStr(Message.ReadInt()) & "." & CStr(Message.ReadInt()) & "." & CStr(Message.ReadInt())
    
    If Not AsciiValidos(UserName) Then
        Call WriteErrorMsg(UserIndex, "Nombre invalido.")
        'TODO Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    If Not PersonajeExiste(UserName) Then
        Call WriteErrorMsg(UserIndex, "El personaje no existe.")
        'TODO Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If

        If BANCheck(UserName) Then
            Call WriteErrorMsg(UserIndex, "Se te ha prohibido la entrada a Argentum debido a tu mal comportamiento. Puedes consultar el reglamento y el sistema de soporte desde www.argentumonline.com.ar")
        ElseIf Not VersionOK(version) Then
            Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en www.argentumonline.com.ar")
        Else
            Call ConnectUser(UserIndex, UserName, Password)
        End If

End Sub

''
' Handles the "ThrowDices" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleThrowDices(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex).Stats
        .UserAtributos(eAtributos.Fuerza) = 9 + RandomNumber(0, 4) + RandomNumber(0, 5)
        .UserAtributos(eAtributos.Agilidad) = 9 + RandomNumber(0, 4) + RandomNumber(0, 5)
        .UserAtributos(eAtributos.Inteligencia) = 12 + RandomNumber(0, 3) + RandomNumber(0, 3)
        .UserAtributos(eAtributos.Carisma) = 12 + RandomNumber(0, 3) + RandomNumber(0, 3)
        .UserAtributos(eAtributos.Constitucion) = 12 + RandomNumber(0, 3) + RandomNumber(0, 3)
    End With
    
    Call WriteDiceRoll(UserIndex)
End Sub

''
' Handles the "LoginNewChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    Dim UserName As String
    Dim Password As String
    Dim version As String
    Dim skills(NUMSKILLS - 1) As Byte
    Dim race As eRaza
    Dim gender As eGenero
    Dim homeland As eCiudad
    Dim Class As eClass
    Dim mail As String

    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(UserIndex, "La creacion de personajes en este servidor se ha deshabilitado.")
        'TODO Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    If ServerSoloGMs <> 0 Then
        Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Consulte la página oficial o el foro oficial para mas información.")
        'TODO Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    If aClon.MaxPersonajes(UserList(UserIndex).ip) Then
        Call WriteErrorMsg(UserIndex, "Has creado demasiados personajes.")
        'TODO Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    UserName = Message.ReadString16()
    Password = Message.ReadString16()

    'Convert version number to string
    version = CStr(Message.ReadInt()) & "." & CStr(Message.ReadInt()) & "." & CStr(Message.ReadInt())

    race = Message.ReadInt()
    gender = Message.ReadInt()
    Class = Message.ReadInt()
    Call Message.ReadSafeArrayInt8(skills)
    
    mail = Message.ReadString16()
    homeland = Message.ReadInt()
    

        If Not VersionOK(version) Then
            Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en www.argentumonline.com.ar")
        Else
            Call ConnectNewUser(UserIndex, UserName, Password, race, gender, Class, skills, mail, homeland)
        End If

End Sub

''
' Handles the "Talk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)
    

        
        Dim chat As String
        
        chat = Message.ReadString16()
        
        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.name, "Dijo: " & chat)
        End If
        
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            If .flags.invisible = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        If LenB(chat) <> 0 Then

            If .flags.Muerto = 1 Then
                Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, .flags.ChatColor))
            End If
        End If

    End With

End Sub

''
' Handles the "Yell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleYell(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)
    

        
        Dim chat As String
        
        chat = Message.ReadString16()
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos.", FontTypeNames.FONTTYPE_INFO)
        Else
            '[Consejeros & GMs]
            If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                Call LogGM(.name, "Grito: " & chat)
            End If
            
            'I see you....
            If .flags.Oculto > 0 Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                If .flags.invisible = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
            
            If LenB(chat) <> 0 Then

                If .flags.Privilegios And PlayerType.User Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, vbRed))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_GM_YELL))
                End If
            End If
        End If

    End With

End Sub

''
' Handles the "Whisper" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim chat As String
        Dim targetCharIndex As Integer
        Dim targetUserIndex As Integer
        Dim targetPriv As PlayerType
        
        targetCharIndex = Message.ReadInt()
        chat = Message.ReadString16()
        
        targetUserIndex = CharIndexToUserIndex(targetCharIndex)
        
        targetPriv = UserList(targetUserIndex).flags.Privilegios
        
        If .flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. ", FontTypeNames.FONTTYPE_INFO)
        Else
            If targetUserIndex = INVALID_INDEX Then
                Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (targetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Then
                    'A los dioses y admins no vale susurrarles si no sos uno vos mismo (así no pueden ver si están conectados o no)
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Dioses y Admins.", FontTypeNames.FONTTYPE_INFO)
                
                ElseIf (.flags.Privilegios And PlayerType.User) <> 0 And (Not targetPriv And PlayerType.User) <> 0 Then
                    'A los Consejeros y SemiDioses no vale susurrarles si sos un PJ común.
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los GMs.", FontTypeNames.FONTTYPE_INFO)
                
                ElseIf Not EstaPCarea(UserIndex, targetUserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Estas muy lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                
                Else
                    '[Consejeros & GMs]
                    If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                        Call LogGM(.name, "Le dijo a '" & UserList(targetUserIndex).name & "' " & chat)
                    End If
                    
                    If LenB(chat) <> 0 Then

                        Call WriteChatOverHead(UserIndex, chat, .Char.CharIndex, vbBlue)
                        Call WriteChatOverHead(targetUserIndex, chat, .Char.CharIndex, vbBlue)

                        '[CDT 17-02-2004]
                        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                            Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageChatOverHead("a " & UserList(targetUserIndex).name & "> " & chat, .Char.CharIndex, vbYellow))
                        End If
                    End If
                End If
            End If
        End If

    End With

End Sub

''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Dim dummy As Long
    Dim TempTick As Long
    Dim heading As eHeading
    
    With UserList(UserIndex)

        
        heading = Message.ReadInt()
        
        'Prevent SpeedHack
        If .flags.TimesWalk >= 30 Then
            TempTick = GetTickCount And &H7FFFFFFF
            dummy = (TempTick - .flags.StartWalk)
            
            ' 5800 is actually less than what would be needed in perfect conditions to take 30 steps
            '(it's about 193 ms per step against the over 200 needed in perfect conditions)
            If dummy < 5800 Then
                If TempTick - .flags.CountSH > 30000 Then
                    .flags.CountSH = 0
                End If
                
                If Not .flags.CountSH = 0 Then
                    If dummy <> 0 Then _
                        dummy = 126000 \ dummy
                    
                    Call LogHackAttemp("Tramposo SH: " & .name & " , " & dummy)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha sido echado por el servidor por posible uso de SH.", FontTypeNames.FONTTYPE_SERVER))
                    Call CloseSocket(UserIndex)
                    
                    Exit Sub
                Else
                    .flags.CountSH = TempTick
                End If
            End If
            .flags.StartWalk = TempTick
            .flags.TimesWalk = 0
        End If
        
        .flags.TimesWalk = .flags.TimesWalk + 1
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        If .flags.Paralizado = 0 Then
            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                .Char.FX = 0
                .Char.loops = 0
                
                Call WriteMeditateToggle(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            Else
                'Move user
                Call MoveUserChar(UserIndex, heading)
                
                'Stop resting if needed
                If .flags.Descansar Then
                    .flags.Descansar = False
                    
                    Call WriteRestOK(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        Else    'paralized
            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                
                Call WriteConsoleMsg(UserIndex, "No podes moverte porque estas paralizado.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.CountSH = 0
        End If
        
        'Can't move while hidden except he is a thief
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
            If .clase <> eClass.Thief Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                
                'If not under a spell effect, show char
                If .flags.invisible = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                End If
            End If
        End If
        
        If .flags.Muerto = 1 Then
            Call Empollando(UserIndex)
        Else
            .flags.EstaEmpo = 0
            .EmpoCont = 0
        End If
    End With
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Call WritePosUpdate(UserIndex)
End Sub

''
' Handles the "Attack" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'If dead, can't attack
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡No podes atacar a nadie porque estas muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'If not in combat mode, can't attack
        If Not .flags.ModoCombate Then
            Call WriteConsoleMsg(UserIndex, "No estás en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'If user meditates, can't attack
        If .flags.Meditando Then
            Exit Sub
        End If
        
        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(UserIndex, "No podés usar así esta arma.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Attack!
        Call UsuarioAtaca(UserIndex)
        
        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            If .flags.invisible = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
End Sub

''
' Handles the "PickUp" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'If dead, it can't pick up objects
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Los muertos no pueden tomar objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Lower rank administrators can't pick up items
        If .flags.Privilegios And PlayerType.Consejero Then
            If Not .flags.Privilegios And PlayerType.RoleMaster Then
                Call WriteConsoleMsg(UserIndex, "No puedes tomar ningun objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        Call GetObj(UserIndex)
    End With
End Sub

''
' Handles the "CombatModeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeCombatModeToggle(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.ModoCombate Then
            Call WriteConsoleMsg(UserIndex, "Has salido del modo de combate.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "Has pasado al modo de combate.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        .flags.ModoCombate = Not .flags.ModoCombate
    End With
End Sub

''
' Handles the "SafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Seguro Then
            Call WriteSafeModeOff(UserIndex)
        Else
            Call WriteSafeModeOn(UserIndex)
        End If
        
        .flags.Seguro = Not .flags.Seguro
    End With
End Sub

''
' Handles the "ResuscitationSafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResuscitationToggle(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)
        Call Message.ReadInt
        
        .flags.SeguroResu = Not .flags.SeguroResu
        
        If .flags.SeguroResu Then
            Call WriteResuscitationSafeOn(UserIndex)
        Else
            Call WriteResuscitationSafeOff(UserIndex)
        End If
    End With
End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestGuildLeaderInfo(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Call modGuilds.SendGuildLeaderInfo(UserIndex)
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Call WriteAttributes(UserIndex)
End Sub

''
' Handles the "RequestFame" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestFame(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Call EnviarFama(UserIndex)
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Call WriteSendSkills(UserIndex)
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Call WriteMiniStats(UserIndex)
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    'User quits commerce mode
    UserList(UserIndex).flags.Comerciando = False
    Call WriteCommerceEnd(UserIndex)
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'Quits commerce mode with user
        If .ComUsu.DestUsu > 0 And UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
            Call WriteConsoleMsg(.ComUsu.DestUsu, .name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(.ComUsu.DestUsu)
        End If
        
        Call FinComerciarUsu(UserIndex)
    End With
End Sub

''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'User exits banking mode
        .flags.Comerciando = False
        Call WriteBankEnd(UserIndex)
    End With
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    'Trade accepted
    Call AceptarComercioUsu(UserIndex)
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    Dim otherUser As Integer
    
    With UserList(UserIndex)

        
        otherUser = .ComUsu.DestUsu
        
        'Offer rejected
        If otherUser > 0 Then
            If UserList(otherUser).flags.UserLogged Then
                Call WriteConsoleMsg(otherUser, .name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(otherUser)
            End If
        End If
        
        Call WriteConsoleMsg(UserIndex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)
    End With
End Sub

''
' Handles the "Drop" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Dim Slot As Byte
    Dim amount As Integer
    
    With UserList(UserIndex)


        Slot = Message.ReadInt()
        amount = Message.ReadInt()
        
        'low rank admins can't drop item. Neither can the dead nor those sailing.
        If .flags.Navegando = 1 Or _
           .flags.Muerto = 1 Or _
           ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub

        'Are we dropping gold or other items??
        If Slot = FLAGORO Then
            If amount > 10000 Then Exit Sub 'Don't drop too much gold

            Call TirarOro(amount, UserIndex)
            
            Call WriteUpdateGold(UserIndex)
        Else
            'Only drop valid slots
            If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).ObjIndex = 0 Then
                    Exit Sub
                End If
                
                Call DropObj(UserIndex, Slot, amount, .Pos.map, .Pos.X, .Pos.Y)
            End If
        End If
    End With
End Sub

''
' Handles the "CastSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim Spell As Byte
        
        Spell = Message.ReadInt()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .flags.Hechizo = Spell
        
        If .flags.Hechizo < 1 Then
            .flags.Hechizo = 0
        ElseIf .flags.Hechizo > MAXUSERHECHIZOS Then
            .flags.Hechizo = 0
        End If
    End With
End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

        Dim X As Byte
        Dim Y As Byte
        
        X = Message.ReadInt()
        Y = Message.ReadInt()
        
        Call LookatTile(UserIndex, UserList(UserIndex).Pos.map, X, Y)

End Sub

''
' Handles the "DoubleClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

        
        Dim X As Byte
        Dim Y As Byte
        
        X = Message.ReadInt()
        Y = Message.ReadInt()
        
        Call Accion(UserIndex, UserList(UserIndex).Pos.map, X, Y)

End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim Skill As eSkill
        
        Skill = Message.ReadInt()
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
            Case Robar, Magia, Domar
                Call WriteWorkRequestTarget(UserIndex, Skill)
            Case Ocultarse
                If .flags.Navegando = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 3 Then
                        Call WriteConsoleMsg(UserIndex, "No podés ocultarte si estás navegando.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 3
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                
                If .flags.Oculto = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 2 Then
                        Call WriteConsoleMsg(UserIndex, "Ya estás oculto.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 2
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                
                Call DoOcultarse(UserIndex)
        End Select
    End With
End Sub

''
' Handles the "UseSpellMacro" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseSpellMacro(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(.name & " fue expulsado por Anti-macro de hechizos", FontTypeNames.FONTTYPE_VENENO))
        Call WriteErrorMsg(UserIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros")
        'TODO Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
    End With
End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim Slot As Byte
        
        Slot = Message.ReadInt()
        
        If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
            If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
        End If
        
        If .flags.Meditando Then
            Exit Sub    'The error message should have been provided by the client.
        End If
        
        Call UseInvItem(UserIndex, Slot)
    End With
End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


        
        Dim Item As Integer
        
        Item = Message.ReadInt()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkHerreria = 0 Then Exit Sub
        
        Call HerreroConstruirItem(UserIndex, Item)

End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

        Dim Item As Integer
        
        Item = Message.ReadInt()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
        
        Call CarpinteroConstruirItem(UserIndex, Item)

End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim X As Byte
        Dim Y As Byte
        Dim Skill As eSkill
        Dim DummyInt As Integer
        Dim tU As Integer   'Target user
        Dim tN As Integer   'Target NPC
        
        X = Message.ReadInt()
        Y = Message.ReadInt()
        
        Skill = Message.ReadInt()
        
        
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando _
                        Or Not InMapBounds(.Pos.map, X, Y) Then
            Exit Sub
        End If
        
        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
            Case eSkill.Proyectiles
            
                'Check attack interval
                If Not IntervaloPermiteAtacar(UserIndex, False) Then Exit Sub
                'Check Magic interval
                If Not IntervaloPermiteLanzarSpell(UserIndex, False) Then Exit Sub
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                'Make sure the item is valid and there is ammo equipped.
                With .Invent
                    If .WeaponEqpObjIndex = 0 Then
                        DummyInt = 1
                    ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
                        DummyInt = 1
                    ElseIf .MunicionEqpSlot < 1 Or .MunicionEqpSlot > MAX_INVENTORY_SLOTS Then
                        DummyInt = 1
                    ElseIf .MunicionEqpObjIndex = 0 Then
                        DummyInt = 1
                    ElseIf ObjData(.WeaponEqpObjIndex).proyectil <> 1 Then
                        DummyInt = 2
                    ElseIf ObjData(.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
                        DummyInt = 1
                    ElseIf .Object(.MunicionEqpSlot).amount < 1 Then
                        DummyInt = 1
                    End If
                    
                    If DummyInt <> 0 Then
                        If DummyInt = 1 Then
                            Call WriteConsoleMsg(UserIndex, "No tenés municiones.", FontTypeNames.FONTTYPE_INFO)
                            
                            Call Desequipar(UserIndex, .WeaponEqpSlot)
                        End If
                        
                        Call Desequipar(UserIndex, .MunicionEqpSlot)
                        Exit Sub
                    End If
                End With
                
                'Quitamos stamina
                If .Stats.MinSta >= 10 Then
                    Call QuitarSta(UserIndex, RandomNumber(1, 10))
                Else
                    Call WriteConsoleMsg(UserIndex, "Estás muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                Call LookatTile(UserIndex, .Pos.map, X, Y)
                
                tU = .flags.TargetUser
                tN = .flags.TargetNPC
                
                'Validate target
                If tU > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                        Exit Sub
                    End If
                    
                    'Prevent from hitting self
                    If tU = UserIndex Then
                        Call WriteConsoleMsg(UserIndex, "¡No puedes atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Attack!
                    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub 'TODO: Por ahora pongo esto para solucionar lo anterior.
                    Call UsuarioAtacaUsuario(UserIndex, tU)
                ElseIf tN > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(Npclist(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(Npclist(tN).Pos.X - .Pos.X) > RANGO_VISION_X Then
                        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                        Exit Sub
                    End If
                    
                    'Is it attackable???
                    If Npclist(tN).Attackable <> 0 Then
                        
                        'Attack!
                        Call UsuarioAtacaNpc(UserIndex, tN)
                    End If
                End If
                
                With .Invent
                    DummyInt = .MunicionEqpSlot
                    
                    'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
                    Call QuitarUserInvItem(UserIndex, DummyInt, 1)
                    
                    If .Object(DummyInt).amount > 0 Then
                        'QuitarUserInvItem unequipps the ammo, so we equip it again
                        .MunicionEqpSlot = DummyInt
                        .MunicionEqpObjIndex = .Object(DummyInt).ObjIndex
                        .Object(DummyInt).Equipped = 1
                    Else
                        .MunicionEqpSlot = 0
                        .MunicionEqpObjIndex = 0
                    End If
                    Call UpdateUserInv(False, UserIndex, DummyInt)
                End With
                '-----------------------------------
            
            Case eSkill.Magia
                'Check the map allows spells to be casted.
                If MapInfo(.Pos.map).MagiaSinEfecto > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energía", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                'Target whatever is in that tile
                Call LookatTile(UserIndex, .Pos.map, X, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Call LogCheating("Ataque fuera de rango de " & .name & "(" & .Pos.map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .ip & " a la posicion (" & .Pos.map & "/" & X & "/" & Y & ")")
                    Exit Sub
                End If
                
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
                
                'Check Spell-Hit interval
                If Not IntervaloPermiteGolpeMagia(UserIndex) Then
                    'Check Magic interval
                    If Not IntervaloPermiteLanzarSpell(UserIndex) Then
                        Exit Sub
                    End If
                End If
                
                
                'Check intervals and cast
                If .flags.Hechizo > 0 Then
                    Call LanzarHechizo(.flags.Hechizo, UserIndex)
                    .flags.Hechizo = 0
                Else
                    Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Pesca
                DummyInt = .Invent.WeaponEqpObjIndex
                If DummyInt = 0 Then Exit Sub
                
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                'Basado en la idea de Barrin
                'Comentario por Barrin: jah, "basado", caradura ! ^^
                If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 1 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If HayAgua(.Pos.map, X, Y) Then
                    Select Case DummyInt
                        Case CAÑA_PESCA
                            Call DoPescar(UserIndex)
                        
                        Case RED_PESCA
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            Call DoPescarRed(UserIndex)
                        
                        Case Else
                            Exit Sub    'Invalid item!
                    End Select
                    
                    'Play sound!
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Robar
                'Does the map allow us to steal here?
                If MapInfo(.Pos.map).Pk Then
                    
                    'Check interval
                    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                    'Target whatever is in that tile
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.map, X, Y)
                    
                    tU = .flags.TargetUser
                    
                    If tU > 0 And tU <> UserIndex Then
                        'Can't steal administrative players
                        If UserList(tU).flags.Privilegios And PlayerType.User Then
                            If UserList(tU).flags.Muerto = 0 Then
                                 If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                     Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                     Exit Sub
                                 End If
                                 
                                 '17/09/02
                                 'Check the trigger
                                 If MapData(UserList(tU).Pos.map, X, Y).trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(UserIndex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(UserIndex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 Call DoRobar(UserIndex, tU)
                            End If
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "No a quien robarle!.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "¡No podés robar en zonas seguras!.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Talar
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                If .Invent.WeaponEqpObjIndex = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Deberías equiparte el hacha.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .Invent.WeaponEqpObjIndex <> HACHA_LEÑADOR Then
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                DummyInt = MapData(.Pos.map, X, Y).ObjInfo.ObjIndex
                
                If DummyInt > 0 Then
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Barrin 29/9/03
                    If .Pos.X = X And .Pos.Y = Y Then
                        Call WriteConsoleMsg(UserIndex, "No podés talar desde allí.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    '¿Hay un arbol donde clickeo?
                    If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                        Call DoTalar(UserIndex)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Mineria
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                                
                If .Invent.WeaponEqpObjIndex = 0 Then Exit Sub
                
                If .Invent.WeaponEqpObjIndex <> PIQUETE_MINERO Then
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                'Target whatever is in the tile
                Call LookatTile(UserIndex, .Pos.map, X, Y)
                
                DummyInt = MapData(.Pos.map, X, Y).ObjInfo.ObjIndex
                
                If DummyInt > 0 Then
                    'Check distance
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    DummyInt = MapData(.Pos.map, X, Y).ObjInfo.ObjIndex 'CHECK
                    '¿Hay un yacimiento donde clickeo?
                    If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then
                        Call DoMineria(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Ahí no hay ningun yacimiento.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Domar
                'Modificado 25/11/02
                'Optimizado y solucionado el bug de la doma de
                'criaturas hostiles.
                
                'Target whatever is that tile
                Call LookatTile(UserIndex, .Pos.map, X, Y)
                tN = .flags.TargetNPC
                
                If tN > 0 Then
                    If Npclist(tN).flags.Domable > 0 Then
                        If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
                            Call WriteConsoleMsg(UserIndex, "No podés domar una criatura que está luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        Call DoDomar(UserIndex, tN)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No podés domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay ninguna criatura alli!.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                'Check there is a proper item there
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then
                        'Validate other items
                        If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > MAX_INVENTORY_SLOTS Then
                            Exit Sub
                        End If
                        
                        ''chequeamos que no se zarpe duplicando oro
                        If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                            If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).amount = 0 Then
                                Call WriteConsoleMsg(UserIndex, "No tienes más minerales", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            ''FUISTE
                            Call WriteErrorMsg(UserIndex, "Has sido expulsado por el sistema anti cheats.")
                            'TODO Call FlushBuffer(UserIndex)
                            Call CloseSocket(UserIndex)
                            Exit Sub
                        End If
                        
                        Call FundirMineral(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Herreria
                'Target wehatever is in that tile
                Call LookatTile(UserIndex, .Pos.map, X, Y)
                
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otYunque Then
                        Call EnivarArmasConstruibles(UserIndex)
                        Call EnivarArmadurasConstruibles(UserIndex)
                        Call WriteShowBlacksmithForm(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateNewGuild(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim desc As String
        Dim GuildName As String
        Dim site As String
        Dim codex() As String
        Dim errorStr As String
        
        desc = Message.ReadString16()
        GuildName = Message.ReadString16()
        site = Message.ReadString16()
        codex = Split(Message.ReadString16(), SEPARATOR)
        
        If modGuilds.CrearNuevoClan(UserIndex, desc, GuildName, site, codex, .FundandoGuildAlineacion, errorStr) Then
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg(.name & " fundó el clan " & GuildName & " de alineación " & modGuilds.GuildAlignment(.guildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))

            
            'Update tag
             Call RefreshCharStatus(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If

    End With

End Sub

''
' Handles the "SpellInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim spellSlot As Byte
        Dim Spell As Integer
        
        spellSlot = Message.ReadInt()
        
        'Validate slot
        If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo.!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate spell in the slot
        Spell = .Stats.UserHechizos(spellSlot)
        If Spell > 0 And Spell < NumeroHechizos + 1 Then
            With Hechizos(Spell)
                'Send information
                Call WriteConsoleMsg(UserIndex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf _
                                               & "Nombre:" & .Nombre & vbCrLf _
                                               & "Descripción:" & .desc & vbCrLf _
                                               & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf _
                                               & "Mana necesario: " & .ManaRequerido & vbCrLf _
                                               & "Stamina necesaria: " & .StaRequerido & vbCrLf _
                                               & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%", FontTypeNames.FONTTYPE_INFO)
            End With
        End If
    End With
End Sub

''
' Handles the "EquipItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim itemSlot As Byte
        
        itemSlot = Message.ReadInt()
        
        'Dead users can't equip items
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo podés usar items cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate item slot
        If itemSlot > MAX_INVENTORY_SLOTS Or itemSlot < 1 Then Exit Sub
        
        If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
        Call EquiparInvItem(UserIndex, itemSlot)
    End With
End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim heading As eHeading
        Dim posX As Integer
        Dim posY As Integer
                
        heading = Message.ReadInt()
        
        If .flags.Paralizado = 1 And .flags.Inmovilizado = 0 Then
            Select Case heading
                Case eHeading.NORTH
                    posY = -1
                Case eHeading.EAST
                    posX = 1
                Case eHeading.SOUTH
                    posY = 1
                Case eHeading.WEST
                    posX = -1
            End Select
            
                If LegalPos(.Pos.map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then
                    Exit Sub
                End If
        End If
        
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If heading > 0 And heading < 5 Then
            .Char.heading = heading
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    End With
End Sub

''
' Handles the "ModifySkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        Dim i As Long
        Dim Count As Integer
        Dim points(1 To NUMSKILLS) As Byte
        
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            points(i) = Message.ReadInt()
            
            If points(i) < 0 Then
                Call LogHackAttemp(.name & " IP:" & .ip & " trató de hackear los skills.")
                .Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            Count = Count + points(i)
        Next i
        
        If Count > .Stats.SkillPts Then
            Call LogHackAttemp(.name & " IP:" & .ip & " trató de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        With .Stats
            For i = 1 To NUMSKILLS
                .SkillPts = .SkillPts - points(i)
                .UserSkills(i) = .UserSkills(i) + points(i)
                
                'Client should prevent this, but just in case...
                If .UserSkills(i) > 100 Then
                    .SkillPts = .SkillPts + .UserSkills(i) - 100
                    .UserSkills(i) = 100
                End If
            Next i
        End With
    End With
End Sub

''
' Handles the "Train" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim SpawnedNpc As Integer
        Dim petIndex As Byte
        
        petIndex = Message.ReadInt()
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
            If petIndex > 0 And petIndex < Npclist(.flags.TargetNPC).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(petIndex).NpcIndex, Npclist(.flags.TargetNPC).Pos, True, False)
                
                If SpawnedNpc > 0 Then
                    Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNPC
                    Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1
                End If
            End If
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
        End If
    End With
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim Slot As Byte
        Dim amount As Integer
        
        Slot = Message.ReadInt()
        amount = Message.ReadInt()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
            
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "No estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'User compra el item
        Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC, Slot, amount)
    End With
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim Slot As Byte
        Dim amount As Integer
        
        Slot = Message.ReadInt()
        amount = Message.ReadInt()
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿Es el banquero?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User retira el item del slot
        Call UserRetiraItem(UserIndex, Slot, amount)
    End With
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim Slot As Byte
        Dim amount As Integer
        
        Slot = Message.ReadInt()
        amount = Message.ReadInt()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'User compra el item del slot
        Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC, Slot, amount)
    End With
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim Slot As Byte
        Dim amount As Integer
        
        Slot = Message.ReadInt()
        amount = Message.ReadInt()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User deposita el item del slot rdata
        Call UserDepositaItem(UserIndex, Slot, amount)
    End With
End Sub

''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


        Dim dir As Integer
        
        If Message.ReadBool() Then
            dir = 1
        Else
            dir = -1
        End If
        
        Call DesplazarHechizo(UserIndex, dir, Message.ReadInt())

End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleClanCodexUpdate(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim desc As String
        Dim codex() As String
        
        desc = Message.ReadString16()
        codex = Split(Message.ReadString16(), SEPARATOR)
        
        Call modGuilds.ChangeCodexAndDesc(desc, codex, .guildIndex)

    End With

End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim amount As Long
        Dim Slot As Byte
        Dim tUser As Integer
        
        Slot = Message.ReadInt()
        amount = Message.ReadInt()
        
        'Get the other player
        tUser = .ComUsu.DestUsu
        
        'If amount is invalid, or slot is invalid and it's not gold, then ignore it.
        If ((Slot < 1 Or Slot > MAX_INVENTORY_SLOTS) And Slot <> FLAGORO) _
                        Or amount <= 0 Then Exit Sub
        
        'Is the other player valid??
        If tUser < 1 Or tUser > MaxUsers Then Exit Sub
        
        'Is the commerce attempt valid??
        If UserList(tUser).ComUsu.DestUsu <> UserIndex Then
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        End If
        
        'Is he still logged??
        If Not UserList(tUser).flags.UserLogged Then
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        Else
            'Is he alive??
            If UserList(tUser).flags.Muerto = 1 Then
                Call FinComerciarUsu(UserIndex)
                Exit Sub
            End If
            
            'Has he got enough??
            If Slot = FLAGORO Then
                'gold
                If amount > .Stats.GLD Then
                    Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            Else
                'inventory
                If amount > .Invent.Object(Slot).amount Then
                    Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            
            'Prevent offer changes (otherwise people would ripp off other players)
            If .ComUsu.Objeto > 0 Then
                Call WriteConsoleMsg(UserIndex, "No puedes cambiar tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
            
            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 Then
                If .Invent.BarcoSlot = Slot Then
                    Call WriteConsoleMsg(UserIndex, "No podés vender tu barco mientras lo estés usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            
            .ComUsu.Objeto = Slot
            .ComUsu.cant = amount
            
            'If the other one had accepted, we turn that back and inform of the new offer (just to be cautious).
            If UserList(tUser).ComUsu.Acepto = True Then
                UserList(tUser).ComUsu.Acepto = False
                Call WriteConsoleMsg(tUser, .name & " ha cambiado su oferta.", FontTypeNames.FONTTYPE_TALK)
            End If
            
            Call EnviarObjetoTransaccion(tUser)
        End If
    End With
End Sub

''
' Handles the "GuildAcceptPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptPeace(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = Message.ReadString16()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .guildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.guildIndex), FontTypeNames.FONTTYPE_GUILD))
        End If

    End With

End Sub

''
' Handles the "GuildRejectAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectAlliance(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = Message.ReadString16()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .guildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.guildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", FontTypeNames.FONTTYPE_GUILD))
        End If

    End With

End Sub

''
' Handles the "GuildRejectPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectPeace(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = Message.ReadString16()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .guildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.guildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", FontTypeNames.FONTTYPE_GUILD))
        End If

    End With

End Sub

''
' Handles the "GuildAcceptAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptAlliance(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = Message.ReadString16()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .guildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.guildIndex), FontTypeNames.FONTTYPE_GUILD))
        End If

    End With

End Sub

''
' Handles the "GuildOfferPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferPeace(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim guild As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = Message.ReadString16()
        proposal = Message.ReadString16()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de paz enviada", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If

    End With

End Sub

''
' Handles the "GuildOfferAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferAlliance(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim guild As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = Message.ReadString16()
        proposal = Message.ReadString16()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de alianza enviada", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If

    End With

End Sub

''
' Handles the "GuildAllianceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAllianceDetails(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim guild As String
        Dim errorStr As String
        Dim details As String
        
        guild = Message.ReadString16()
        
        details = modGuilds.r_VerPropuesta(UserIndex, guild, RELACIONES_GUILD.ALIADOS, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)
        End If

    End With

End Sub

''
' Handles the "GuildPeaceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeaceDetails(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim guild As String
        Dim errorStr As String
        Dim details As String
        
        guild = Message.ReadString16()
        
        details = modGuilds.r_VerPropuesta(UserIndex, guild, RELACIONES_GUILD.PAZ, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)
        End If

    End With

End Sub

''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestJoinerInfo(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim User As String
        Dim details As String
        
        User = Message.ReadString16()
        
        details = modGuilds.a_DetallesAspirante(UserIndex, User)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, "El personaje no ha mandado solicitud, o no estás habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(UserIndex, details)
        End If

    End With

End Sub

''
' Handles the "GuildAlliancePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAlliancePropList(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Call WriteAlianceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.ALIADOS))
End Sub

''
' Handles the "GuildPeacePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeacePropList(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Call WritePeaceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.PAZ))
End Sub

''
' Handles the "GuildDeclareWar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildDeclareWar(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim guild As String
        Dim errorStr As String
        Dim otherGuildIndex As Integer
        
        guild = Message.ReadString16()
        
        otherGuildIndex = modGuilds.r_DeclararGuerra(UserIndex, guild, errorStr)
        
        If otherGuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            'WAR shall be!
            Call SendData(SendTarget.ToGuildMembers, .guildIndex, PrepareMessageConsoleMsg("TU CLAN HA ENTRADO EN GUERRA CON " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.guildIndex) & " LE DECLARA LA GUERRA A TU CLAN", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .guildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        End If

    End With

End Sub

''
' Handles the "GuildNewWebsite" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildNewWebsite(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Call modGuilds.ActualizarWebSite(UserIndex, Message.ReadString16())

    End With

End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptNewMember(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim errorStr As String
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        If Not modGuilds.a_AceptarAspirante(UserIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                Call modGuilds.m_ConectarMiembroAClan(tUser, .guildIndex)
                Call RefreshCharStatus(tUser)
            End If
            
            Call SendData(SendTarget.ToGuildMembers, .guildIndex, PrepareMessageConsoleMsg(UserName & " ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .guildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))
        End If

    End With

End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectNewMember(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim errorStr As String
        Dim UserName As String
        Dim reason As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        reason = Message.ReadString16()
        
        If Not modGuilds.a_RechazarAspirante(UserIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call WriteConsoleMsg(tUser, errorStr & " : " & reason, FontTypeNames.FONTTYPE_GUILD)
            Else
                'hay que grabar en el char su rechazo
                Call modGuilds.a_RechazarAspiranteChar(UserName, .guildIndex, reason)
            End If
        End If

    End With

End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildKickMember(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim guildIndex As Integer
        
        UserName = Message.ReadString16()
        
        guildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
        
        If guildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, guildIndex, PrepareMessageConsoleMsg(UserName & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, guildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        Else
            Call WriteConsoleMsg(UserIndex, "No puedes expulsar ese personaje del clan.", FontTypeNames.FONTTYPE_GUILD)
        End If

    End With

End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildUpdateNews(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Call modGuilds.ActualizarNoticias(UserIndex, Message.ReadString16())

    End With

End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberInfo(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Call modGuilds.SendDetallesPersonaje(UserIndex, Message.ReadString16())

    End With

End Sub

''
' Handles the "GuildOpenElections" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOpenElections(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        Dim Error As String
        
        If Not modGuilds.v_AbrirElecciones(UserIndex, Error) Then
            Call WriteConsoleMsg(UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .guildIndex, PrepareMessageConsoleMsg("¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & .name, FontTypeNames.FONTTYPE_GUILD))
        End If
    End With
End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestMembership(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim guild As String
        Dim application As String
        Dim errorStr As String
        
        guild = Message.ReadString16()
        application = Message.ReadString16()
        
        If Not modGuilds.a_NuevoAspirante(UserIndex, guild, application, errorStr) Then
           Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
           Call WriteConsoleMsg(UserIndex, "Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & guild & ".", FontTypeNames.FONTTYPE_GUILD)
        End If

    End With

End Sub

''
' Handles the "GuildRequestDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestDetails(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Call modGuilds.SendGuildDetails(UserIndex, Message.ReadString16())

    End With

End Sub

''
' Handles the "Online" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnline(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    Dim i As Long
    Dim Count As Long
    
    With UserList(UserIndex)

        
        For i = 1 To LastUser
            If LenB(UserList(i).name) <> 0 Then
                If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then _
                    Count = Count + 1
            End If
        Next i
        
        Call WriteConsoleMsg(UserIndex, "Número de usuarios: " & CStr(Count), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    Dim tUser As Integer
    
    With UserList(UserIndex)

        
        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu
            
            If UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(tUser)
                End If
            End If
            
            Call WriteConsoleMsg(UserIndex, "Comercio cancelado. ", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(UserIndex)
        End If
        
        Call Cerrar_Usuario(UserIndex)
    End With
End Sub

''
' Handles the "GuildLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildLeave(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    Dim guildIndex As Integer
    
    With UserList(UserIndex)

        
        'obtengo el guildindex
        guildIndex = m_EcharMiembroDeClan(UserIndex, .name)
        
        If guildIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, guildIndex, PrepareMessageConsoleMsg(.name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
        Else
            Call WriteConsoleMsg(UserIndex, "Tu no puedes salir de ningún clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    Dim earnings As Integer
    Dim percentage As Integer
    
    With UserList(UserIndex)

        
        'Dead people can't check their accounts
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenes que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case Npclist(.flags.TargetNPC).NPCtype
            Case eNPCType.Banquero
                Call WriteChatOverHead(UserIndex, "Tenes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Case eNPCType.Timbero
                If Not .flags.Privilegios And PlayerType.User Then
                    earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    
                    If earnings >= 0 And Apuestas.Ganancias <> 0 Then
                        percentage = Int(earnings * 100 / Apuestas.Ganancias)
                    End If
                    
                    If earnings < 0 And Apuestas.Perdidas <> 0 Then
                        percentage = Int(earnings * 100 / Apuestas.Perdidas)
                    End If
                    
                    Call WriteConsoleMsg(UserIndex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub

''
' Handles the "PetStand" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenás que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's his pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it!
        Npclist(.flags.TargetNPC).Movement = TipoAI.ESTATICO
        
        Call Expresar(.flags.TargetNPC, UserIndex)
    End With
End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenás que seleccionar un personaje, hace click izquierdo sobre ál.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it
        Call FollowAmo(.flags.TargetNPC)
        
        Call Expresar(.flags.TargetNPC, UserIndex)
    End With
End Sub

''
' Handles the "TrainList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrainList(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's the trainer
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC)
    End With
End Sub

''
' Handles the "Rest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRest(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Solo podés usar items cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(UserIndex)
            
            If Not .flags.Descansar Then
                Call WriteConsoleMsg(UserIndex, "Te acomodás junto a la fogata y comenzás a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.Descansar = Not .flags.Descansar
        Else
            If .flags.Descansar Then
                Call WriteRestOK(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                
                .flags.Descansar = False
                Exit Sub
            End If
            
            Call WriteConsoleMsg(UserIndex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Meditate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Solo podés usar meditar cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Can he meditate?
        If .Stats.MaxMAN = 0 Then
             Call WriteConsoleMsg(UserIndex, "Sólo las clases mágicas conocen el arte de la meditación", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        'Admins don't have to wait :D
        If Not .flags.Privilegios And PlayerType.User Then
            .Stats.MinMAN = .Stats.MaxMAN
            Call WriteConsoleMsg(UserIndex, "Mana restaurado", FontTypeNames.FONTTYPE_VENENO)
            Call WriteUpdateMana(UserIndex)
            Exit Sub
        End If
        
        Call WriteMeditateToggle(UserIndex)
        
        If .flags.Meditando Then _
           Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        
        .flags.Meditando = Not .flags.Meditando
        
        'Barrin 3/10/03 Tiempo de inicio al meditar
        If .flags.Meditando Then
            .Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
            
            Call WriteConsoleMsg(UserIndex, "Te estás concentrando. En " & Fix(TIEMPO_INICIOMEDITAR / 1000) & " segundos comenzarás a meditar.", FontTypeNames.FONTTYPE_INFO)
            
            .Char.loops = INFINITE_LOOPS
            
            'Show proper FX according to level
            If .Stats.ELV < 13 Then
                .Char.FX = FXIDs.FXMEDITARCHICO
            
            ElseIf .Stats.ELV < 25 Then
                .Char.FX = FXIDs.FXMEDITARMEDIANO
            
            ElseIf .Stats.ELV < 35 Then
                .Char.FX = FXIDs.FXMEDITARGRANDE
            
            ElseIf .Stats.ELV < 42 Then
                .Char.FX = FXIDs.FXMEDITARXGRANDE
            
            Else
                .Char.FX = FXIDs.FXMEDITARXXGRANDE
            End If
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
        Else
            .Counters.bPuedeMeditar = False
            
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
        End If
    End With
End Sub

''
' Handles the "Resucitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate NPC and make sure player is dead
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor _
            And (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) _
            Or .flags.Muerto = 0 Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede resucitarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call RevivirUsuario(UserIndex)
        Call WriteConsoleMsg(UserIndex, "¡¡Hás sido resucitado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Heal" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor _
            And Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie) _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Stats.MinHP = .Stats.MaxHP
        
        Call WriteUpdateHP(UserIndex)
        
        Call WriteConsoleMsg(UserIndex, "¡¡Hás sido curado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "RequestStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Call SendUserStatsTxt(UserIndex, UserIndex)
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Is it already in commerce mode??
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            'Does the NPC want to trade??
            If Npclist(.flags.TargetNPC).Comercia = 0 Then
                If LenB(Npclist(.flags.TargetNPC).desc) <> 0 Then
                    Call WriteChatOverHead(UserIndex, "No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                End If
                
                Exit Sub
            End If
            
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Start commerce....
            Call IniciarComercioNPC(UserIndex)
        '[Alejo]
        ElseIf .flags.TargetUser > 0 Then
            'User commerce...
            'Can he commerce??
            If .flags.Privilegios And PlayerType.Consejero Then
                Call WriteConsoleMsg(UserIndex, "No puedes vender items.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            
            'Is the other one dead??
            If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡¡No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is it me??
            If .flags.TargetUser = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con vos mismo...", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Check distance
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is he already trading?? is it with me or someone else??
            If UserList(.flags.TargetUser).flags.Comerciando = True And _
                UserList(.flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Initialize some variables...
            .ComUsu.DestUsu = .flags.TargetUser
            .ComUsu.DestNick = UserList(.flags.TargetUser).name
            .ComUsu.cant = 0
            .ComUsu.Objeto = 0
            .ComUsu.Acepto = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, .flags.TargetUser)
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "BankStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'If it's the banker....
            If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                Call IniciarDeposito(UserIndex)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Enlist" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnlist(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            Call EnlistarArmadaReal(UserIndex)
        Else
            Call EnlistarCaos(UserIndex)
        End If
    End With
End Sub

''
' Handles the "Information" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInformation(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
                Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
             If .Faccion.ArmadaReal = 0 Then
                 Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, cada 100 criminales que derrotes te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
             If .Faccion.FuerzasCaos = 0 Then
                 Call WriteChatOverHead(UserIndex, "No perteneces a la legión oscura!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
    End With
End Sub

''
' Handles the "Reward" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReward(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
             If .Faccion.ArmadaReal = 0 Then
                 Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call RecompensaArmadaReal(UserIndex)
        Else
             If .Faccion.FuerzasCaos = 0 Then
                 Call WriteChatOverHead(UserIndex, "No perteneces a la legión oscura!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call RecompensaCaos(UserIndex)
        End If
    End With
End Sub

''
' Handles the "RequestMOTD" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMOTD(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Call SendMOTD(UserIndex)
End Sub

''
' Handles the "UpTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Dim time As Long
    Dim UpTimeStr As String
    
    'Get total time in seconds
    time = ((GetTickCount() And &H7FFFFFFF) - tInicioServer) \ 1000
    
    'Get times in dd:hh:mm:ss format
    UpTimeStr = (time Mod 60) & " segundos."
    time = time \ 60
    
    UpTimeStr = (time Mod 60) & " minutos, " & UpTimeStr
    time = time \ 60
    
    UpTimeStr = (time Mod 24) & " horas, " & UpTimeStr
    time = time \ 24
    
    If time = 1 Then
        UpTimeStr = Time & " día, " & UpTimeStr
    Else
        UpTimeStr = Time & " días, " & UpTimeStr
    End If
    
    Call WriteConsoleMsg(UserIndex, "Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
End Sub

''
' Handles the "PartyLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyLeave(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Call mdParty.SalirDeParty(UserIndex)
End Sub

''
' Handles the "PartyCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyCreate(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub
    
    Call mdParty.CrearParty(UserIndex)
End Sub

''
' Handles the "PartyJoin" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyJoin(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Call mdParty.SolicitarIngresoAParty(UserIndex)
End Sub

''
' Handles the "GuildMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim chat As String
        
        chat = Message.ReadString16()
        
        If LenB(chat) <> 0 Then

            If .guildIndex > 0 Then
                Call SendData(SendTarget.ToDiosesYclan, .guildIndex, PrepareMessageGuildChat(.name & "> " & chat))
'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
                'Call SendData(SendTarget.ToClanArea, userindex, UserList(userindex).Pos.Map, "||" & vbYellow & "°< " & rData & " >°" & CStr(UserList(userindex).Char.CharIndex))
            End If
        End If

    End With

End Sub

''
' Handles the "PartyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyMessage(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim chat As String
        
        chat = Message.ReadString16()
        
        If LenB(chat) <> 0 Then
   
            Call mdParty.BroadCastParty(UserIndex, chat)
'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
            'Call SendData(SendTarget.ToPartyArea, userindex, UserList(userindex).Pos.Map, "||" & vbYellow & "°< " & mid$(rData, 7) & " >°" & CStr(UserList(userindex).Char.CharIndex))
        End If

    End With

End Sub

''
' Handles the "GuildOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnline(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        Dim onlineList As String
        
        onlineList = modGuilds.m_ListaDeMiembrosOnline(UserIndex, .guildIndex)
        
        If .guildIndex <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Compañeros de tu clan conectados: " & onlineList, FontTypeNames.FONTTYPE_GUILDMSG)
        Else
            Call WriteConsoleMsg(UserIndex, "No pertences a ningún clan.", FontTypeNames.FONTTYPE_GUILDMSG)
        End If
    End With
End Sub

''
' Handles the "PartyOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyOnline(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Call mdParty.OnlineParty(UserIndex)
End Sub

''
' Handles the "CouncilMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim chat As String
        
        chat = Message.ReadString16()
        
        If LenB(chat) <> 0 Then

            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejo, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))
            End If
        End If

    End With

End Sub

''
' Handles the "RoleMasterRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoleMasterRequest(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim request As String
        
        request = Message.ReadString16()
        
        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Su solicitud ha sido enviada", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))
        End If

    End With

End Sub

''
' Handles the "GMRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMRequest(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If Not Ayuda.Existe(.name) Then
            Call WriteConsoleMsg(UserIndex, "El mensaje ha sido entregado, ahora sólo debes esperar que se desocupe algún GM.", FontTypeNames.FONTTYPE_INFO)
            Call Ayuda.Push(.name)
        Else
            Call Ayuda.Quitar(.name)
            Call Ayuda.Push(.name)
            Call WriteConsoleMsg(UserIndex, "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim Description As String
        
        Description = Message.ReadString16()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedés cambiar la descripción estando muerto.", FontTypeNames.FONTTYPE_INFO)
        Else
            If Not AsciiValidos(Description) Then
                Call WriteConsoleMsg(UserIndex, "La descripción tiene caractéres inválidos.", FontTypeNames.FONTTYPE_INFO)
            Else
                .desc = Trim$(Description)
                Call WriteConsoleMsg(UserIndex, "La descripción a cambiado.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

    End With

End Sub

''
' Handles the "GuildVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildVote(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim vote As String
        Dim errorStr As String
        
        vote = Message.ReadString16()
        
        If Not modGuilds.v_UsuarioVota(UserIndex, vote, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Voto NO contabilizado: " & errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)
        End If

    End With

End Sub

''
' Handles the "Punishments" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePunishments(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim name As String
        Dim Count As Integer
        
        name = Message.ReadString16()
        
        If LenB(name) <> 0 Then
            If (InStrB(name, "\") <> 0) Then
                name = Replace(name, "\", "")
            End If
            If (InStrB(name, "/") <> 0) Then
                name = Replace(name, "/", "")
            End If
            If (InStrB(name, ":") <> 0) Then
                name = Replace(name, ":", "")
            End If
            If (InStrB(name, "|") <> 0) Then
                name = Replace(name, "|", "")
            End If
            
            If FileExist(CharPath & name & ".chr", vbNormal) Then
                Count = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
                If Count = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                Else
                    While Count > 0
                        Call WriteConsoleMsg(UserIndex, Count & " - " & GetVar(CharPath & name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
                        Count = Count - 1
                    Wend
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Personaje """ & name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

    End With

End Sub

''
' Handles the "ChangePassword" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangePassword(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        Dim oldPass As String
        Dim newPass As String
        Dim oldPass2 As String
        

        oldPass = Message.ReadString16()
        newPass = Message.ReadString16()

        
        If LenB(newPass) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Debe especificar una contraseña nueva, inténtelo de nuevo", FontTypeNames.FONTTYPE_INFO)
        Else
            oldPass2 = GetVar(CharPath & UserList(UserIndex).name & ".chr", "INIT", "Password")
            
            If oldPass2 <> oldPass Then
                Call WriteConsoleMsg(UserIndex, "La contraseña actual proporcionada no es correcta. La contraseña no ha sido cambiada, inténtelo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteVar(CharPath & UserList(UserIndex).name & ".chr", "INIT", "Password", newPass)
                Call WriteConsoleMsg(UserIndex, "La contraseña fue cambiada con éxito", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

    End With

End Sub


''
' Handles the "Gamble" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim amount As Integer
        
        amount = Message.ReadInt()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
        ElseIf .flags.TargetNPC = 0 Then
            'Validate target NPC
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
            Call WriteChatOverHead(UserIndex, "No tengo ningún interés en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf amount < 1 Then
            Call WriteChatOverHead(UserIndex, "El mínimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf amount > 5000 Then
            Call WriteChatOverHead(UserIndex, "El máximo de apuesta es 5000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf .Stats.GLD < amount Then
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            If RandomNumber(1, 100) <= 47 Then
                .Stats.GLD = .Stats.GLD + amount
                Call WriteChatOverHead(UserIndex, "Felicidades! Has ganado " & CStr(amount) & " monedas de oro!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Perdidas = Apuestas.Perdidas + amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.GLD = .Stats.GLD - amount
                Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & CStr(amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Ganancias = Apuestas.Ganancias + amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
            End If
            
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
            Call WriteUpdateGold(UserIndex)
        End If
    End With
End Sub

''
' Handles the "BankExtractGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractGold(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim amount As Long
        
        amount = Message.ReadInt()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If amount > 0 And amount <= .Stats.Banco Then
             .Stats.Banco = .Stats.Banco - amount
             .Stats.GLD = .Stats.GLD + amount
             Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
             Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
        
        Call WriteUpdateGold(UserIndex)
    End With
End Sub

''
' Handles the "LeaveFaction" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeaveFaction(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Noble Then
           'Quit the Royal Army?
           If .Faccion.ArmadaReal = 1 Then
               If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
                   Call ExpulsarFaccionReal(UserIndex)
                   Call WriteChatOverHead(UserIndex, "Serás bienvenido a las fuerzas imperiales si deseas regresar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
               Else
                   Call WriteChatOverHead(UserIndex, "¡¡¡Sal de aquí bufón!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
               End If
            'Quit the Chaos Legion??
           ElseIf .Faccion.FuerzasCaos = 1 Then
               If Npclist(.flags.TargetNPC).flags.Faccion = 1 Then
                   Call ExpulsarFaccionCaos(UserIndex)
                   Call WriteChatOverHead(UserIndex, "Ya volverás arrastrandote.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
               Else
                   Call WriteChatOverHead(UserIndex, "Sal de aquí maldito criminal", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
               End If
           Else
               Call WriteChatOverHead(UserIndex, "¡No perteneces a ninguna facción!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
           End If
        End If
    End With
End Sub

''
' Handles the "BankDepositGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDepositGold(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim amount As Long
        
        amount = Message.ReadInt()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If amount > 0 And amount <= .Stats.GLD Then
            .Stats.Banco = .Stats.Banco + amount
            .Stats.GLD = .Stats.GLD - amount
            Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Call WriteUpdateGold(UserIndex)
        Else
            Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
    End With
End Sub

''
' Handles the "Denounce" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim Text As String
        
        Text = Message.ReadString16()
        
        If .flags.Silenciado = 0 Then

            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(LCase$(.name) & " DENUNCIA: " & Text, FontTypeNames.FONTTYPE_GUILDMSG))
            Call WriteConsoleMsg(UserIndex, "Denuncia enviada, espere..", FontTypeNames.FONTTYPE_INFO)
        End If

    End With

End Sub

''
' Handles the "GuildFundate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundate(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim clanType As eClanType
        Dim Error As String
        
        clanType = Message.ReadInt()
        
        Select Case UCase$(Trim(clanType))
            Case eClanType.ct_RoyalArmy
                .FundandoGuildAlineacion = ALINEACION_ARMADA
            Case eClanType.ct_Evil
                .FundandoGuildAlineacion = ALINEACION_LEGION
            Case eClanType.ct_Neutral
                .FundandoGuildAlineacion = ALINEACION_NEUTRO
            Case eClanType.ct_GM
                .FundandoGuildAlineacion = ALINEACION_MASTER
            Case eClanType.ct_Legal
                .FundandoGuildAlineacion = ALINEACION_CIUDA
            Case eClanType.ct_Criminal
                .FundandoGuildAlineacion = ALINEACION_CRIMINAL
            Case Else
                Call WriteConsoleMsg(UserIndex, "Alineación inválida.", FontTypeNames.FONTTYPE_GUILD)
                Exit Sub
        End Select
        
        If modGuilds.PuedeFundarUnClan(UserIndex, .FundandoGuildAlineacion, Error) Then
            Call WriteShowGuildFundationForm(UserIndex)
        Else
            .FundandoGuildAlineacion = 0
            Call WriteConsoleMsg(UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub

''
' Handles the "PartyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyKick(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        tUser = NameIndex(UserName)
        If tUser > 0 Then
            Call mdParty.ExpulsarDeParty(UserIndex, tUser)
        Else
            If InStr(UserName, "+") Then
                UserName = Replace(UserName, "+", " ")
            End If
            
            Call WriteConsoleMsg(UserIndex, UserName & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
        End If

    End With

End Sub

''
' Handles the "PartySetLeader" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartySetLeader(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        tUser = NameIndex(UserName)
        If tUser > 0 Then
            Call mdParty.TransformarEnLider(UserIndex, tUser)
        Else
            Call WriteConsoleMsg(UserIndex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
        End If

    End With

End Sub

''
' Handles the "PartyAcceptMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyAcceptMember(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        Dim rank As Integer
        
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        UserName = Message.ReadString16()
        
        tUser = NameIndex(UserName)
        If tUser > 0 Then
            'Validate administrative ranks - don't allow users to spoof online GMs
            If (UserList(tUser).flags.Privilegios And rank) <= (.flags.Privilegios And rank) Then
                Call mdParty.AprobarIngresoAParty(UserIndex, tUser)
            Else
                Call WriteConsoleMsg(UserIndex, "No puedes incorporar a tu party a personajes de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            If InStr(UserName, "+") Then
                UserName = Replace(UserName, "+", " ")
            End If
            
            'Don't allow users to spoof online GMs
            If (UserDarPrivilegioLevel(UserName) And rank) <= (.flags.Privilegios And rank) Then
                Call WriteConsoleMsg(UserIndex, UserName & " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
            Else
                Call WriteConsoleMsg(UserIndex, "No puedes incorporar a tu party a personajes de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

    End With

End Sub

''
' Handles the "GuildMemberList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberList(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim guild As String
        Dim memberCount As Integer
        Dim i As Long
        Dim UserName As String
        
        guild = Message.ReadString16()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If (InStrB(guild, "\") <> 0) Then
                guild = Replace(guild, "\", "")
            End If
            If (InStrB(guild, "/") <> 0) Then
                guild = Replace(guild, "/", "")
            End If
            
            If Not FileExist(App.Path & "\guilds\" & guild & "-members.mem") Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & guild, FontTypeNames.FONTTYPE_INFO)
            Else
                memberCount = val(GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "INIT", "NroMembers"))
                
                For i = 1 To memberCount
                    UserName = GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "Members", "Member" & i)
                    
                    Call WriteConsoleMsg(UserIndex, UserName & "<" & guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i
            End If
        End If

    End With

End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim Text As String
        
        Text = Message.ReadString16()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.name, "Mensaje a Gms:" & Text)
        
            If LenB(Text) <> 0 Then
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & "> " & Text, FontTypeNames.FONTTYPE_GMMSG))
            End If
        End If

    End With

End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            .showName = Not .showName 'Show / Hide the name
            
            Call RefreshCharStatus(UserIndex)
        End If
    End With
End Sub

''
' Handles the "OnlineRoyalArmy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineRoyalArmy(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)
        'Remove packet ID
        Message.ReadInt
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i As Long
        Dim list As String

        For i = 1 To LastUser
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.ArmadaReal = 1 Then
                    If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Or _
                      .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        list = list & UserList(i).name & ", "
                    End If
                End If
            End If
        Next i
    End With
    
    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Armadas conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay Armadas conectados", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)
        'Remove packet ID
        Message.ReadInt
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i As Long
        Dim list As String

        For i = 1 To LastUser
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.FuerzasCaos = 1 Then
                    If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Or _
                      .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        list = list & UserList(i).name & ", "
                    End If
                End If
            End If
        Next i
    End With

    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay Caos conectados", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

''
' Handles the "GoNearby" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        
        UserName = Message.ReadString16()
        
        Dim tIndex As Integer
        Dim X As Long
        Dim Y As Long
        Dim i As Long
        Dim found As Boolean
        
        tIndex = NameIndex(UserName)
        
        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                If tIndex <= 0 Then 'existe el usuario destino?
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
                        For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                            For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i
                                If MapData(UserList(tIndex).Pos.map, X, Y).UserIndex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.map, X, Y, True, True) Then
                                        Call WarpUserChar(UserIndex, UserList(tIndex).Pos.map, X, Y, True)
                                        found = True
                                        Exit For
                                    End If
                                End If
                            Next Y
                            
                            If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                        Next X
                        
                        If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next i
                    
                    'No space found??
                    If Not found Then
                        Call WriteConsoleMsg(UserIndex, "Todos los lugares están ocupados.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        End If

    End With

End Sub

''
' Handles the "Comment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleComment(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim comment As String
        comment = Message.ReadString16()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.name, "Comentario: " & comment)
            Call WriteConsoleMsg(UserIndex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)
        End If

    End With

End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

    
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Call LogGM(.name, "Hora.")
    End With
    
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & time & " " & Date, FontTypeNames.FONTTYPE_INFO))
End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call WriteConsoleMsg(UserIndex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, "/Donde " & UserName)
                End If
            End If
        End If

    End With

End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim map As Integer
        Dim i, j As Long
        Dim NPCcount1, NPCcount2 As Integer
        Dim NPCcant1() As Integer
        Dim NPCcant2() As Integer
        Dim List1() As String
        Dim List2() As String
        
        map = Message.ReadInt()
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        If MapaValido(map) Then
            For i = 1 To LastNPC
                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If Npclist(i).Pos.map = map Then
                    '¿esta vivo?
                    If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 And Npclist(i).Stats.Alineacion = 2 Then
                        If NPCcount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim NPCcant1(0) As Integer
                            NPCcount1 = 1
                            List1(0) = Npclist(i).name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant1(0) = 1
                        Else
                            For j = 0 To NPCcount1 - 1
                                If Left$(List1(j), Len(Npclist(i).name)) = Npclist(i).name Then
                                    List1(j) = List1(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant1(j) = NPCcant1(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = NPCcount1 Then
                                ReDim Preserve List1(0 To NPCcount1) As String
                                ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
                                NPCcount1 = NPCcount1 + 1
                                List1(j) = Npclist(i).name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant1(j) = 1
                            End If
                        End If
                    Else
                        If NPCcount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim NPCcant2(0) As Integer
                            NPCcount2 = 1
                            List2(0) = Npclist(i).name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant2(0) = 1
                        Else
                            For j = 0 To NPCcount2 - 1
                                If Left$(List2(j), Len(Npclist(i).name)) = Npclist(i).name Then
                                    List2(j) = List2(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant2(j) = NPCcant2(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = NPCcount2 Then
                                ReDim Preserve List2(0 To NPCcount2) As String
                                ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
                                NPCcount2 = NPCcount2 + 1
                                List2(j) = Npclist(i).name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant2(j) = 1
                            End If
                        End If
                    End If
                End If
            Next i
            
            Call WriteConsoleMsg(UserIndex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount1 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay NPCS Hostiles", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount1 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call WriteConsoleMsg(UserIndex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount2 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay más NPCS", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call LogGM(.name, "Numero enemigos en mapa " & map)
        End If
    End With
End Sub

''
' Handles the "WarpMeToTarget" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpMeToTarget(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call WarpUserChar(UserIndex, .flags.TargetMap, .flags.TargetX, .flags.TargetY, True)
        Call LogGM(.name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.map)
    End With
End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim map As Integer
        Dim X As Byte
        Dim Y As Byte
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        map = Message.ReadInt()
        X = Message.ReadInt()
        Y = Message.ReadInt()
        
        If Not .flags.Privilegios And PlayerType.User Then
            If MapaValido(map) And LenB(UserName) <> 0 Then
                If UCase$(UserName) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(UserName)
                    End If
                Else
                    tUser = UserIndex
                End If
            
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                ElseIf InMapBounds(map, X, Y) Then
                    Call WarpUserChar(tUser, map, X, Y, True)
                    Call WriteConsoleMsg(UserIndex, UserList(tUser).name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, "Transportó a " & UserList(tUser).name & " hacia " & "Mapa" & map & " X:" & X & " Y:" & Y)
                End If
            End If
        End If

    End With

End Sub

''
' Handles the "Silence" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
        
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    Call WriteConsoleMsg(UserIndex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "ESTIMADO USUARIO, ud ha sido silenciado por los administradores. Sus denuncias serán ignoradas por el servidor de aquí en más. Utilice /GM para contactar un administrador.")
                    Call LogGM(.name, "/silenciar " & UserList(tUser).name)
                
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(UserIndex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, "/DESsilenciar " & UserList(tUser).name)
                End If
            End If
        End If

    End With

End Sub

''
' Handles the "SOSShowList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowSOSForm(UserIndex)
    End With
End Sub

''
' Handles the "SOSRemove" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        UserName = Message.ReadString16()
        
        If Not .flags.Privilegios And PlayerType.User Then _
            Call Ayuda.Quitar(UserName)

    End With

End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        tUser = NameIndex(UserName)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WarpUserChar(UserIndex, UserList(tUser).Pos.map, UserList(tUser).Pos.X, UserList(tUser).Pos.Y + 1, True)
                    
                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
         
                    End If
                    
                    Call LogGM(.name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)
                End If
            End If
        End If

    End With

End Sub

''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call DoAdminInvisible(UserIndex)
        Call LogGM(.name, "/INVISIBLE")
    End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call WriteShowGMPanelForm(UserIndex)
    End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    Dim i As Long
    Dim names() As String
    Dim Count As Long
    
    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        ReDim names(1 To LastUser) As String
        Count = 1
        
        For i = 1 To LastUser
            If (LenB(UserList(i).name) <> 0) Then
                If UserList(i).flags.Privilegios And PlayerType.User Then
                    names(Count) = UserList(i).name
                    Count = Count + 1
                End If
            End If
        Next i
        
        If Count > 1 Then Call WriteUserNameList(UserIndex, names(), Count - 1)
    End With
End Sub

''
' Handles the "Working" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorking(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    Dim i As Long
    Dim users As String
    
    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                users = users & ", " & UserList(i).name
            End If
        Next i
        
        If LenB(users) <> 0 Then
            users = Right$(users, Len(users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios trabajando: " & users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios trabajando", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Hiding" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHiding(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    Dim i As Long
    Dim users As String
    
    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser
            If (LenB(UserList(i).name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                users = users & UserList(i).name & ", "
            End If
        Next i
        
        If LenB(users) <> 0 Then
            users = Left$(users, Len(users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios ocultandose: " & users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios ocultandose", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Jail" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleJail(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim reason As String
        Dim jailTime As Byte
        Dim Count As Byte
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        reason = Message.ReadString16()
        jailTime = Message.ReadInt()
        
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")
        End If
        
        '/carcel nick@motivo@<tiempo>
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(UserName)
                
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                        Call WriteConsoleMsg(UserIndex, "No podés encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > 60 Then
                        Call WriteConsoleMsg(UserIndex, "No podés encarcelar por más de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")
                        End If
                        If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")
                        End If
                        
                        If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                            Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(reason) & " " & Date & " " & time)
                        End If
                        
                        Call Encarcelar(tUser, jailTime, .name)
                        Call LogGM(.name, " encarcelo a " & UserName)
                    End If
                End If
            End If
        End If

    End With

End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Dim tNPC As Integer
        Dim auxNPC As npc
        
        'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
        If .flags.Privilegios And PlayerType.Consejero Then
            If .Pos.map = MAPA_PRETORIANO Then
                Call WriteConsoleMsg(UserIndex, "Los consejeros no pueden usar este comando en el mapa pretoriano.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        tNPC = .flags.TargetNPC
        
        If tNPC > 0 Then
            Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & Npclist(tNPC).name, FontTypeNames.FONTTYPE_INFO)
            
            auxNPC = Npclist(tNPC)
            Call QuitarNPC(tNPC)
            Call ReSpawnNpc(auxNPC)
            
            .flags.TargetNPC = 0
        Else
            Call WriteConsoleMsg(UserIndex, "Debes hacer click sobre el NPC antes", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "WarnUser" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim reason As String
        Dim privs As PlayerType
        Dim Count As Byte
        
        UserName = Message.ReadString16()
        reason = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                privs = UserDarPrivilegioLevel(UserName)
                
                If Not privs And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "No podés advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")
                    End If
                    If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")
                    End If
                    
                    If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                        Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.name) & ": ADVERTENCIA por: " & LCase$(reason) & " " & Date & " " & time)
                        
                        Call WriteConsoleMsg(UserIndex, "Has advertido a " & UCase$(UserName), FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.name, " advirtio a " & UserName)
                    End If
                End If
            End If
        End If

    End With

End Sub

''
' Handles the "EditChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEditChar(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        Dim opcion As Byte
        Dim Arg1 As String
        Dim Arg2 As String
        Dim valido As Boolean
        Dim LoopC As Byte
        Dim commandString As String
        
        UserName = Replace(Message.ReadString16(), "+", " ")
        
        If UCase$(UserName) = "YO" Then
            tUser = UserIndex
        Else
            tUser = NameIndex(UserName)
        End If
        
        opcion = Message.ReadInt()
        Arg1 = Message.ReadString16()
        Arg2 = Message.ReadString16()
        
        If .flags.Privilegios And PlayerType.RoleMaster Then
            Select Case .flags.Privilegios And (PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)
                Case PlayerType.Consejero
                    ' Los RMs consejeros sólo se pueden editar su head, body y level
                    valido = tUser = UserIndex And _
                            (opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head Or opcion = eEditOptions.eo_Level)
                
                Case PlayerType.SemiDios
                    ' Los RMs sólo se pueden editar su level y el head y body de cualquiera
                    valido = (opcion = eEditOptions.eo_Level And tUser = UserIndex) _
                            Or opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head
                
                Case PlayerType.Dios
                    ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    ' pero si quiere modificar el level sólo lo puede hacer sobre sí mismo
                    valido = (opcion = eEditOptions.eo_Level And tUser = UserIndex) Or _
                            opcion = eEditOptions.eo_Body Or _
                            opcion = eEditOptions.eo_Head Or _
                            opcion = eEditOptions.eo_CiticensKilled Or _
                            opcion = eEditOptions.eo_CriminalsKilled Or _
                            opcion = eEditOptions.eo_Class Or _
                            opcion = eEditOptions.eo_Skills
            End Select
            
        ElseIf .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then   'Si no es RM debe ser dios para poder usar este comando
            valido = True
        End If
        
        If valido Then
            Select Case opcion
                Case eEditOptions.eo_Gold
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(Arg1) < 5000000 Then
                            UserList(tUser).Stats.GLD = val(Arg1)
                            Call WriteUpdateGold(tUser)
                        Else
                            Call WriteConsoleMsg(UserIndex, "No esta permitido utilizar valores mayores. Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                
                Case eEditOptions.eo_Experience
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(Arg1) < 15995001 Then
                            UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(Arg1)
                            Call CheckUserLevel(tUser)
                            Call WriteUpdateExp(tUser)
                        Else
                            Call WriteConsoleMsg(UserIndex, "No esta permitido utilizar valores mayores a mucho. Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                
                Case eEditOptions.eo_Body
                    If tUser <= 0 Then
                        Call WriteVar(CharPath & UserName & ".chr", "INIT", "Body", Arg1)
                        Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call ChangeUserChar(tUser, val(Arg1), UserList(tUser).Char.Head, UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                    End If
                
                Case eEditOptions.eo_Head
                    If tUser <= 0 Then
                        Call WriteVar(CharPath & UserName & ".chr", "INIT", "Head", Arg1)
                        Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call ChangeUserChar(tUser, UserList(tUser).Char.body, val(Arg1), UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                    End If
                
                Case eEditOptions.eo_CriminalsKilled
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(Arg1) > MAXUSERMATADOS Then
                            UserList(tUser).Faccion.CriminalesMatados = MAXUSERMATADOS
                        Else
                            UserList(tUser).Faccion.CriminalesMatados = val(Arg1)
                        End If
                    End If
                
                Case eEditOptions.eo_CiticensKilled
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(Arg1) > MAXUSERMATADOS Then
                            UserList(tUser).Faccion.CiudadanosMatados = MAXUSERMATADOS
                        Else
                            UserList(tUser).Faccion.CiudadanosMatados = val(Arg1)
                        End If
                    End If
                
                Case eEditOptions.eo_Level
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(Arg1) > STAT_MAXELV Then
                            Arg1 = CStr(STAT_MAXELV)
                            Call WriteConsoleMsg(UserIndex, "No puedes tener un nivel superior a " & STAT_MAXELV & ".", FONTTYPE_INFO)
                        End If
                        
                        UserList(tUser).Stats.ELV = val(Arg1)
                    End If
                    
                    Call WriteUpdateUserStats(UserIndex)
                
                Case eEditOptions.eo_Class
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        For LoopC = 1 To NUMCLASES
                            If UCase$(ListaClases(LoopC)) = UCase$(Arg1) Then Exit For
                        Next LoopC
                        
                        If LoopC > NUMCLASES Then
                            Call WriteConsoleMsg(UserIndex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            UserList(tUser).clase = LoopC
                        End If
                    End If
                
                Case eEditOptions.eo_Skills
                    For LoopC = 1 To NUMSKILLS
                        If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg1) Then Exit For
                    Next LoopC
                    
                    If LoopC > NUMSKILLS Then
                        Call WriteConsoleMsg(UserIndex, "Skill Inexistente!", FontTypeNames.FONTTYPE_INFO)
                    Else
                        If tUser <= 0 Then
                            Call WriteVar(CharPath & UserName & ".chr", "Skills", "SK" & LoopC, Arg2)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            UserList(tUser).Stats.UserSkills(LoopC) = val(Arg2)
                        End If
                    End If
                
                Case eEditOptions.eo_SkillPointsLeft
                    If tUser <= 0 Then
                        Call WriteVar(CharPath & UserName & ".chr", "STATS", "SkillPtsLibres", Arg1)
                        Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        UserList(tUser).Stats.SkillPts = val(Arg1)
                    End If
                
                Case eEditOptions.eo_Nobleza
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(Arg1) > MAXREP Then
                            UserList(tUser).Reputacion.NobleRep = MAXREP
                        Else
                            UserList(tUser).Reputacion.NobleRep = val(Arg1)
                        End If
                    End If
                
                Case eEditOptions.eo_Asesino
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        If val(Arg1) > MAXREP Then
                            UserList(tUser).Reputacion.AsesinoRep = MAXREP
                        Else
                            UserList(tUser).Reputacion.AsesinoRep = val(Arg1)
                        End If
                    End If
                
                Case eEditOptions.eo_Sex
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Arg1 = UCase$(Arg1)
                        If (Arg1 = "MUJER") Then
                            UserList(tUser).genero = eGenero.Mujer
                        ElseIf (Arg1 = "HOMBRE") Then
                            UserList(tUser).genero = eGenero.Hombre
                        End If
                    End If
                
                Case eEditOptions.eo_Raza
                    If tUser <= 0 Then
                        Call WriteConsoleMsg(UserIndex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Arg1 = UCase$(Arg1)
                        If (Arg1 = "HUMANO") Then
                            UserList(tUser).raza = eRaza.Humano
                        ElseIf (Arg1 = "ELFO") Then
                            UserList(tUser).raza = eRaza.Elfo
                        ElseIf (Arg1 = "DROW") Then
                            UserList(tUser).raza = eRaza.Drow
                        ElseIf (Arg1 = "ENANO") Then
                            UserList(tUser).raza = eRaza.Enano
                        ElseIf (Arg1 = "GNOMO") Then
                            UserList(tUser).raza = eRaza.Gnomo
                        End If
                    End If
                
                Case Else
                    Call WriteConsoleMsg(UserIndex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)
            End Select
        End If
        
        'Log it!
        commandString = "/MOD "
        
        Select Case opcion
            Case eEditOptions.eo_Gold
                commandString = commandString & "ORO "
            
            Case eEditOptions.eo_Experience
                commandString = commandString & "EXP "
            
            Case eEditOptions.eo_Body
                commandString = commandString & "BODY "
            
            Case eEditOptions.eo_Head
                commandString = commandString & "HEAD "
            
            Case eEditOptions.eo_CriminalsKilled
                commandString = commandString & "CRI "
            
            Case eEditOptions.eo_CiticensKilled
                commandString = commandString & "CIU "
            
            Case eEditOptions.eo_Level
                commandString = commandString & "LEVEL "
            
            Case eEditOptions.eo_Class
                commandString = commandString & "CLASE "
            
            Case eEditOptions.eo_Skills
                commandString = commandString & "SKILLS "
            
            Case eEditOptions.eo_SkillPointsLeft
                commandString = commandString & "SKILLSLIBRES "
                
            Case eEditOptions.eo_Nobleza
                commandString = commandString & "NOB "
                
            Case eEditOptions.eo_Asesino
                commandString = commandString & "ASE "
                
            Case eEditOptions.eo_Sex
                commandString = commandString & "SEX "
                
            Case eEditOptions.eo_Raza
                commandString = commandString & "RAZA "
                
            Case Else
                commandString = commandString & "UNKOWN "
        End Select
        
        commandString = commandString & Arg1 & " " & Arg2
        
        If valido Then _
            Call LogGM(.name, commandString & " " & UserName)

    End With


End Sub

''
' Handles the "RequestCharInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

                
        Dim targetName As String
        Dim targetIndex As Integer
        
        targetName = Replace$(Message.ReadString16(), "+", " ")
        targetIndex = NameIndex(targetName)
        
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            'is the player offline?
            If targetIndex <= 0 Then
                'don't allow to retrieve administrator's info
                If Not (EsDios(targetName) Or EsAdmin(targetName)) Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, Buscando en Charfile.", FontTypeNames.FONTTYPE_INFO)
                    Call SendUserStatsTxtOFF(UserIndex, targetName)
                End If
            Else
                'don't allow to retrieve administrator's info
                If UserList(targetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(UserIndex, targetIndex)
                End If
            End If
        End If

    End With

End Sub

''
' Handles the "RequestCharStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharStats(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        UserName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call LogGM(.name, "/STAT " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo Charfile... ", FontTypeNames.FONTTYPE_INFO)
                
                Call SendUserMiniStatsTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserMiniStatsTxt(UserIndex, tUser)
            End If
        End If

    End With


End Sub

''
' Handles the "RequestCharGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharGold(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/BAL " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserOROTxtFromChar(UserIndex, UserName)
            Else
                Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.Banco & " en el banco", FontTypeNames.FONTTYPE_TALK)
            End If
        End If

    End With


End Sub

''
' Handles the "RequestCharInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInventory(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/INV " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo del charfile...", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserInvTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserInvTxt(UserIndex, tUser)
            End If
        End If

    End With


End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/BOV " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserBovedaTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserBovedaTxt(UserIndex, tUser)
            End If
        End If

    End With


End Sub

''
' Handles the "RequestCharSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharSkills(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Long
        Dim Text As String
        
        UserName = Message.ReadString16()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.name, "/STATS " & UserName)
            
            If tUser <= 0 Then
                If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                End If
                
                For LoopC = 1 To NUMSKILLS
                    Text = Text & "CHAR>" & SkillsNames(LoopC) & " = " & GetVar(CharPath & UserName & ".chr", "SKILLS", "SK" & LoopC) & vbCrLf
                Next LoopC
                
                Call WriteConsoleMsg(UserIndex, Text & "CHAR> Libres:" & GetVar(CharPath & UserName & ".chr", "STATS", "SKILLPTSLIBRES"), FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendUserSkillsTxt(UserIndex, tUser)
            End If
        End If

    End With


End Sub

''
' Handles the "ReviveChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = UserIndex
            End If
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                With UserList(tUser)
                    'If dead, show him alive (naked).
                    If .flags.Muerto = 1 Then
                        .flags.Muerto = 0
                        
                        Call DarCuerpoDesnudo(tUser)
                        
                        Call ChangeUserChar(tUser, .Char.body, .OrigChar.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    
                    .Stats.MinHP = .Stats.MaxHP
                End With
                
                Call WriteUpdateHP(tUser)
              
                Call LogGM(.name, "Resucito a " & UserName)
            End If
        End If

    End With


End Sub

''
' Handles the "OnlineGM" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    Dim i As Long
    Dim list As String
    Dim priv As PlayerType
    
    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        priv = PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or PlayerType.Admin
        
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged Then
                If UserList(i).flags.Privilegios And priv Then _
                    list = list & UserList(i).name & ", "
            End If
        Next i
        
        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(UserIndex, list & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "OnlineMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Dim LoopC As Long
        Dim list As String
        Dim priv As PlayerType
        
        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or PlayerType.Admin)
        
        For LoopC = 1 To LastUser
            If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).Pos.map = .Pos.map Then
                If UserList(LoopC).flags.Privilegios And priv Then _
                    list = list & UserList(LoopC).name & ", "
            End If
        Next LoopC
        
        If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
        
        Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Forgive" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForgive(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If EsNewbie(tUser) Then
                    Call VolverCiudadano(tUser)
                Else
                    Call LogGM(.name, "Intento perdonar un personaje de nivel avanzado.")
                    Call WriteConsoleMsg(UserIndex, "Solo se permite perdonar newbies.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

    End With


End Sub

''
' Handles the "Kick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKick(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        Dim rank As Integer
        
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        UserName = Message.ReadString16()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(UserIndex, "No podes echar a alguien con jerarquia mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " echo a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                    Call CloseSocket(tUser)
                    Call LogGM(.name, "Echo a " & UserName)
                End If
            End If
        End If

    End With


End Sub

''
' Handles the "Execute" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "Estás loco?? como vas a piñatear un gm!!!! :@", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call UserDie(tUser)
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " ha ejecutado a " & UserName, FontTypeNames.FONTTYPE_EJECUCION))
                    Call LogGM(.name, " ejecuto a " & UserName)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "No está online", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

    End With


End Sub

''
' Handles the "BanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim reason As String
        
        UserName = Message.ReadString16()
        reason = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanCharacter(UserIndex, UserName, reason)
        End If

    End With


End Sub

''
' Handles the "UnbanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim cantPenas As Byte
        
        UserName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            
            If Not FileExist(CharPath & UserName & ".chr", vbNormal) Then
                Call WriteConsoleMsg(UserIndex, "Charfile inexistente (no use +)", FontTypeNames.FONTTYPE_INFO)
            Else
                If (val(GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban")) = 1) Then
                    Call UnBan(UserName)
                
                    'penas
                    cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.name) & ": UNBAN. " & Date & " " & time)
                
                    Call LogGM(.name, "/UNBAN a " & UserName)
                    Call WriteConsoleMsg(UserIndex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & " no esta baneado. Imposible unbanear", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

    End With


End Sub

''
' Handles the "NPCFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.TargetNPC > 0 Then
            Call DoFollow(.flags.TargetNPC, .name)
            Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
            Npclist(.flags.TargetNPC).flags.Paralizado = 0
            Npclist(.flags.TargetNPC).Contadores.Paralisis = 0
        End If
    End With
End Sub

''
' Handles the "SummonChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El jugador no esta online.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or _
                  (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.User)) <> 0 Then
                    Call WriteConsoleMsg(tUser, .name & " te há trasportado.", FontTypeNames.FONTTYPE_INFO)
                    Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y + 1, True)
                    Call LogGM(.name, "/SUM " & UserName & " Map:" & .Pos.map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

    End With


End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call EnviarSpawnList(UserIndex)
    End With
End Sub

''
' Handles the "SpawnCreature" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim npc As Integer
        npc = Message.ReadInt()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If npc > 0 And npc <= UBound(Declaraciones.SpawnList()) Then _
              Call SpawnNpc(Declaraciones.SpawnList(npc).NpcIndex, .Pos, True, False)
            
            Call LogGM(.name, "Sumoneo " & Declaraciones.SpawnList(npc).NpcName)
        End If
    End With
End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call ResetNpcInv(.flags.TargetNPC)
        Call LogGM(.name, "/RESETINV " & Npclist(.flags.TargetNPC).name)
    End With
End Sub

''
' Handles the "CleanWorld" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCleanWorld(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)


        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LimpiarMundo
    End With
End Sub

''
' Handles the "ServerMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim Text As String
        Text = Message.ReadString16()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(Text) <> 0 Then
                Call LogGM(.name, "Mensaje Broadcast:" & Text)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Text, FontTypeNames.FONTTYPE_TALK))
            End If
        End If

    End With


End Sub

''
' Handles the "NickToIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        Dim priv As PlayerType
        
        UserName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            Call LogGM(.name, "NICK2IP Solicito la IP de " & UserName)

            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
                priv = PlayerType.User
            End If
            
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And priv Then
                    Call WriteConsoleMsg(UserIndex, "El ip de " & UserName & " es " & UserList(tUser).ip, FontTypeNames.FONTTYPE_INFO)
                    Dim ip As String
                    Dim lista As String
                    Dim LoopC As Long
                    ip = UserList(tUser).ip
                    For LoopC = 1 To LastUser
                        If UserList(LoopC).ip = ip Then
                            If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                If UserList(LoopC).flags.Privilegios And priv Then
                                    lista = lista & UserList(LoopC).name & ", "
                                End If
                            End If
                        End If
                    Next LoopC
                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "No hay ningun personaje con ese nick", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

    End With


End Sub

''
' Handles the "IPToNick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim ip As String
        Dim LoopC As Long
        Dim lista As String
        Dim priv As PlayerType
        
        ip = Message.ReadInt() & "."
        ip = ip & Message.ReadInt() & "."
        ip = ip & Message.ReadInt() & "."
        ip = ip & Message.ReadInt()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, "IP2NICK Solicito los Nicks de IP " & ip)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            priv = PlayerType.User
        End If

        For LoopC = 1 To LastUser
            If UserList(LoopC).ip = ip Then
                If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).flags.Privilegios And priv Then
                        lista = lista & UserList(LoopC).name & ", "
                    End If
                End If
            End If
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "GuildOnlineMembers" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnlineMembers(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim GuildName As String
        Dim tGuild As Integer
        
        GuildName = Message.ReadString16()
        
        If (InStrB(GuildName, "+") <> 0) Then
            GuildName = Replace(GuildName, "+", " ")
        End If
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tGuild = guildIndex(GuildName)
            
            If tGuild > 0 Then
                Call WriteConsoleMsg(UserIndex, "Clan " & UCase(GuildName) & ": " & _
                  modGuilds.m_ListaDeMiembrosOnline(UserIndex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)
            End If
        End If

    End With


End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        mapa = Message.ReadInt()
        X = Message.ReadInt()
        Y = Message.ReadInt()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call LogGM(.name, "/CT " & mapa & "," & X & "," & Y)
        
        If Not MapaValido(mapa) Or Not InMapBounds(mapa, X, Y) Then _
            Exit Sub
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then _
            Exit Sub
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).TileExit.map > 0 Then _
            Exit Sub
        
        If MapData(mapa, X, Y).ObjInfo.ObjIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapData(mapa, X, Y).TileExit.map > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim ET As Obj
        ET.amount = 1
        ET.ObjIndex = 378
        
        Call MakeObj(ET, .Pos.map, .Pos.X, .Pos.Y - 1)
        
        With MapData(.Pos.map, .Pos.X, .Pos.Y - 1)
            .TileExit.map = mapa
            .TileExit.X = X
            .TileExit.Y = Y
        End With
    End With
End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        

        
        '/dt
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        
        If Not InMapBounds(mapa, X, Y) Then Exit Sub
        
        With MapData(mapa, X, Y)
            If .ObjInfo.ObjIndex = 0 Then Exit Sub
            
            If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.map > 0 Then
                Call LogGM(UserList(UserIndex).name, "/DT: " & mapa & "," & X & "," & Y)
                
                Call EraseObj(.ObjInfo.amount, mapa, X, Y)
                
                If MapData(.TileExit.map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
                    Call EraseObj(1, .TileExit.map, .TileExit.X, .TileExit.Y)
                End If
                
                .TileExit.map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
            End If
        End With
    End With
End Sub

''
' Handles the "RainToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRainToggle(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call LogGM(.name, "/LLUVIA")
        Lloviendo = Not Lloviendo
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
    End With
End Sub

''
' Handles the "SetCharDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetCharDescription(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim tUser As Integer
        Dim desc As String
        
        desc = Message.ReadString16()
        
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
            tUser = .flags.TargetUser
            If tUser > 0 Then
                UserList(tUser).DescRM = desc
            Else
                Call WriteConsoleMsg(UserIndex, "Haz click sobre un personaje antes!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

    End With


End Sub

''
' Handles the "ForceMIDIToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim midiID As Byte
        Dim mapa As Integer
        
        midiID = Message.ReadInt
        mapa = Message.ReadInt
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, 50, 50) Then
                mapa = .Pos.map
            End If
        
            If midiID = 0 Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMidi(MapInfo(.Pos.map).Music))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMidi(midiID))
            End If
        End If
    End With
End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim waveID As Byte
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        waveID = Message.ReadInt()
        mapa = Message.ReadInt()
        X = Message.ReadInt()
        Y = Message.ReadInt()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
        'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, X, Y) Then
                mapa = .Pos.map
                X = .Pos.X
                Y = .Pos.Y
            End If
            
            'Ponemos el pedido por el GM
            Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayWave(waveID, X, Y))
        End If
    End With
End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim Text As String
        Text = Message.ReadString16()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("ARMADA REAL> " & Text, FontTypeNames.FONTTYPE_TALK))
        End If

    End With


End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim Text As String
        Text = Message.ReadString16()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & Text, FontTypeNames.FONTTYPE_TALK))
        End If

    End With


End Sub

''
' Handles the "CitizenMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCitizenMessage(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim Text As String
        Text = Message.ReadString16()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & Text, FontTypeNames.FONTTYPE_TALK))
        End If

    End With


End Sub

''
' Handles the "CriminalMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCriminalMessage(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim Text As String
        Text = Message.ReadString16()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & Text, FontTypeNames.FONTTYPE_TALK))
        End If

    End With


End Sub

''
' Handles the "TalkAsNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim Text As String
        Text = Message.ReadString16()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(Text, Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

    End With


End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.map, X, Y).ObjInfo.ObjIndex > 0 Then
                        If ItemNoEsDeMapa(MapData(.Pos.map, X, Y).ObjInfo.ObjIndex) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.map, X, Y)
                        End If
                    End If
                End If
            Next X
        Next Y
        
        Call LogGM(UserList(UserIndex).name, "/MASSDEST")
    End With
End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    
                    Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If

    End With


End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Consejo de la Legión Oscura.", FontTypeNames.FONTTYPE_CONSEJO))
                
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

                    Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If

    End With


End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim tObj As Integer
        Dim X As Long
        Dim Y As Long
        
        For X = 5 To 95
            For Y = 5 To 95
                tObj = MapData(.Pos.map, X, Y).ObjInfo.ObjIndex
                If tObj > 0 Then
                    If ObjData(tObj).OBJType <> eOBJType.otArboles Then
                        Call WriteConsoleMsg(UserIndex, "(" & X & "," & Y & ") " & ObjData(tObj).name, FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Next Y
        Next X
    End With
End Sub

''
' Handles the "CouncilKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, Echando de los consejos", FontTypeNames.FONTTYPE_INFO)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECE", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECECAOS", 0)
                Else
                    Call WriteConsoleMsg(UserIndex, "No se encuentra el charfile " & CharPath & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                        
                        Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                    
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de la Legión Oscura", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                        
                        Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de la Legión Oscura", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                End With
            End If
        End If

    End With


End Sub

''
' Handles the "SetTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim tTrigger As Byte
        Dim tLog As String
        
        tTrigger = Message.ReadInt()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If tTrigger >= 0 Then
            MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.X & "," & .Pos.Y
            
            Call LogGM(.name, tLog)
            Call WriteConsoleMsg(UserIndex, tLog, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "AskTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    Dim tTrigger As Byte
    
    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        tTrigger = MapData(.Pos.map, .Pos.X, .Pos.Y).trigger
        
        Call LogGM(.name, "Miro el trigger en " & .Pos.map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        
        Call WriteConsoleMsg(UserIndex, _
            "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.X & ", " & .Pos.Y _
            , FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "BannedIPList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim lista As String
        Dim LoopC As Long
        
        Call LogGM(.name, "/BANIPLIST")
        
        For LoopC = 1 To BanIps.Count
            lista = lista & BanIps.Item(LoopC) & ", "
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
        Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call BanIpGuardar
        Call BanIpCargar
    End With
End Sub

''
' Handles the "GuildBan" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildBan(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim GuildName As String
        Dim cantMembers As Integer
        Dim LoopC As Long
        Dim member As String
        Dim Count As Byte
        Dim tIndex As Integer
        Dim tFile As String
        
        GuildName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tFile = App.Path & "\guilds\" & GuildName & "-members.mem"
            
            If Not FileExist(tFile) Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " banned al clan " & UCase$(GuildName), FontTypeNames.FONTTYPE_FIGHT))
                
                'baneamos a los miembros
                Call LogGM(.name, "BANCLAN a " & UCase$(GuildName))
                
                cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))
                
                For LoopC = 1 To cantMembers
                    member = GetVar(tFile, "Members", "Member" & LoopC)
                    'member es la victima
                    Call Ban(member, "Administracion del servidor", "Clan Banned")
                    
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & "> ha sido expulsado del servidor.", FontTypeNames.FONTTYPE_FIGHT))
                    
                    tIndex = NameIndex(member)
                    If tIndex > 0 Then
                        'esta online
                        UserList(tIndex).flags.Ban = 1
                        Call CloseSocket(tIndex)
                    End If
                    
                    'ponemos el flag de ban a 1
                    Call WriteVar(CharPath & member & ".chr", "FLAGS", "Ban", "1")
                    'ponemos la pena
                    Count = val(GetVar(CharPath & member & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & member & ".chr", "PENAS", "Cant", Count + 1)
                    Call WriteVar(CharPath & member & ".chr", "PENAS", "P" & Count + 1, LCase$(.name) & ": BAN AL CLAN: " & GuildName & " " & Date & " " & time)
                Next LoopC
            End If
        End If

    End With


End Sub

''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim bannedIP As String
        Dim tUser As Integer
        Dim reason As String
        Dim i As Long
        
        ' Is it by ip??
        If Message.ReadBool() Then
            bannedIP = Message.ReadInt() & "."
            bannedIP = bannedIP & Message.ReadInt() & "."
            bannedIP = bannedIP & Message.ReadInt() & "."
            bannedIP = bannedIP & Message.ReadInt()
        Else
            tUser = NameIndex(Message.ReadString16())
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
            Else
                bannedIP = UserList(tUser).ip
            End If
        End If
        
        reason = Message.ReadString16()
        
        If LenB(bannedIP) > 0 Then
            If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                Call LogGM(.name, "/BanIP " & bannedIP & " por " & reason)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                Call BanIpAgrega(bannedIP)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " baneó la IP " & bannedIP & " por " & reason, FontTypeNames.FONTTYPE_FIGHT))
                
                'Find every player with that ip and ban him!
                For i = 1 To LastUser
                    If UserList(i).ConnIDValida Then
                        If UserList(i).ip = bannedIP Then
                            Call BanCharacter(UserIndex, UserList(i).name, "IP POR " & reason)
                        End If
                    End If
                Next i
            End If
        End If

    End With


End Sub

''
' Handles the "UnbanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim bannedIP As String
        
        bannedIP = Message.ReadInt() & "."
        bannedIP = bannedIP & Message.ReadInt() & "."
        bannedIP = bannedIP & Message.ReadInt() & "."
        bannedIP = bannedIP & Message.ReadInt()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If BanIpQuita(bannedIP) Then
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "CreateItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)


        Dim tObj As Integer
        tObj = Message.ReadInt()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call LogGM(.name, "/CI: " & tObj)
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then _
            Exit Sub
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y - 1).TileExit.map > 0 Then _
            Exit Sub
        
        If tObj < 1 Or tObj > NumObjDatas Then _
            Exit Sub
        
        'Is the object not null?
        If LenB(ObjData(tObj).name) = 0 Then Exit Sub
        
        Dim Objeto As Obj
        Call WriteConsoleMsg(UserIndex, "ATENCION: FUERON CREADOS ***100*** ITEMS!, TIRE Y /DEST LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
        
        Objeto.amount = 100
        Objeto.ObjIndex = tObj
        Call MakeObj(Objeto, .Pos.map, .Pos.X, .Pos.Y - 1)
    End With
End Sub

''
' Handles the "DestroyItems" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex = 0 Then Exit Sub
        
        Call LogGM(.name, "/DEST")
        
        If ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport Then
            Call WriteConsoleMsg(UserIndex, "No puede destruir teleports así. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call EraseObj(10000, .Pos.map, .Pos.X, .Pos.Y)
    End With
End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionKick(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)
            
            Call LogGM(.name, "ECHO DEL CAOS A: " & UserName)
    
            If tUser > 0 Then
                UserList(tUser).Faccion.FuerzasCaos = 0
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoCaos", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .name)
                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

    End With


End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyKick(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)
            
            Call LogGM(.name, "ECHO DE LA REAL A: " & UserName)
            
            If tUser > 0 Then
                UserList(tUser).Faccion.ArmadaReal = 0
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "EjercitoReal", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Reenlistadas", 200)
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .name)
                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

    End With


End Sub

''
' Handles the "ForceMIDIAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceMIDIAll(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    
    With UserList(UserIndex)


        Dim midiID As Byte
        midiID = Message.ReadInt()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " broadcast musica: " & midiID, FontTypeNames.FONTTYPE_SERVER))
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))
    End With
End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)


        Dim waveID As Byte
        waveID = Message.ReadInt()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))
    End With
End Sub

''
' Handles the "RemovePunishment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRemovePunishment(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim punishment As Byte
        Dim NewText As String
        
        UserName = Message.ReadString16()
        punishment = Message.ReadInt
        NewText = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
            Else
                If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                End If
                
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    Call LogGM(.name, " borro la pena: " & punishment & "-" & _
                      GetVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment) _
                      & " de " & UserName & " y la cambió por: " & NewText)
                    
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment, LCase$(.name) & ": <" & NewText & "> " & Date & " " & time)
                    
                    Call WriteConsoleMsg(UserIndex, "Pena Modificada.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

    End With


End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call LogGM(.name, "/BLOQ")
        
        If MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked = 1
        Else
            MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked = 0
        End If
        
        Call Bloquear(True, .Pos.map, .Pos.X, .Pos.Y, MapData(.Pos.map, .Pos.X, .Pos.Y).Blocked)
    End With
End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call QuitarNPC(.flags.TargetNPC)
        Call LogGM(.name, "/MATA " & Npclist(.flags.TargetNPC).name)
    End With
End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.map, X, Y).NpcIndex)
                End If
            Next X
        Next Y
        Call LogGM(.name, "/MASSKILL")
    End With
End Sub

''
' Handles the "LastIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim lista As String
        Dim LoopC As Byte
        Dim priv As Integer
        Dim validCheck As Boolean
        
        priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            'Handle special chars
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            If (InStrB(UserName, "+") <> 0) Then
                UserName = Replace(UserName, "+", " ")
            End If
            
            'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
            If NameIndex(UserName) > 0 Then
                validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            Else
                validCheck = (UserDarPrivilegioLevel(UserName) And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            End If
            
            If validCheck Then
                Call LogGM(.name, "/LASTIP " & UserName)
                
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    lista = "Las ultimas IPs con las que " & UserName & " se conectó son:"
                    For LoopC = 1 To 5
                        lista = lista & vbCrLf & LoopC & " - " & GetVar(CharPath & UserName & ".chr", "INIT", "LastIP" & LoopC)
                    Next LoopC
                    Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, UserName & " es de mayor jerarquía que vos.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

    End With


End Sub

''
' Handles the "ChatColor" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim color As Long
        
        color = RGB(Message.ReadInt(), Message.ReadInt(), Message.ReadInt())
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            .flags.ChatColor = color
        End If
    End With
End Sub

''
' Handles the "Ignored" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible
        End If
    End With
End Sub

''
' Handles the "CheckSlot" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleCheckSlot(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim Slot As Byte
        Dim tIndex As Integer
        
        UserName = Message.ReadString16() 'Que UserName?
        Slot = Message.ReadInt() 'Que Slot?
        tIndex = NameIndex(UserName)  'Que user index?
        
        Call LogGM(.name, .name & " Checkeo el slot " & Slot & " de " & UserName)
           
        If tIndex > 0 Then
            If Slot > 0 And Slot <= MAX_INVENTORY_SLOTS Then
                If UserList(tIndex).Invent.Object(Slot).ObjIndex > 0 Then
                    Call WriteConsoleMsg(UserIndex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).ObjIndex).name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).amount, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay Objeto en slot seleccionado", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Slot Inválido.", FontTypeNames.FONTTYPE_TALK)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
        End If

    End With

End Sub

''
' Handle the "KickAllChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha echado a todos los personajes.")
        
        Call EcharPjsNoPrivilegiados
    End With
End Sub

''
' Handle the "ShowServerForm" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowServerForm(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha solicitado mostrar el formulario del servidor.")
        Call frmMain.mnuMostrar_Click
    End With
End Sub

''
' Handle the "CleanSOS" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha borrado los SOS")
        
        Call Ayuda.Reset
    End With
End Sub

''
' Handle the "SaveChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha guardado todos los chars")
        
        Call mdParty.ActualizaExperiencias
        Call GuardarUsuarios
    End With
End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim doTheBackUp As Boolean
        
        doTheBackUp = Message.ReadBool()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha cambiado la información sobre el BackUp")
        
        'Change the boolean to byte in a fast way
        If doTheBackUp Then
            MapInfo(.Pos.map).BackUp = 1
        Else
            MapInfo(.Pos.map).BackUp = 0
        End If
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.map & ".dat", "Mapa" & .Pos.map, "backup", MapInfo(.Pos.map).BackUp)
        
        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Backup: " & MapInfo(.Pos.map).BackUp, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim isMapPk As Boolean
        
        isMapPk = Message.ReadBool()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha cambiado la informacion sobre si es PK el mapa.")
        
        MapInfo(.Pos.map).Pk = isMapPk
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.map & ".dat", "Mapa" & .Pos.map, "Pk", IIf(isMapPk, "1", "0"))

        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " PK: " & MapInfo(.Pos.map).Pk, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    Dim tStr As String
    
    With UserList(UserIndex)

        
        tStr = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
                Call LogGM(.name, .name & " ha cambiado la informacion sobre si es Restringido el mapa.")
                MapInfo(UserList(UserIndex).Pos.map).Restringir = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "Restringir", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Restringido: " & MapInfo(.Pos.map).Restringir, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

    End With


End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Dim nomagic As Boolean
    
    With UserList(UserIndex)

        
        nomagic = Message.ReadBool
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar la Magia el mapa.")
            MapInfo(UserList(UserIndex).Pos.map).MagiaSinEfecto = nomagic
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "MagiaSinEfecto", nomagic)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " MagiaSinEfecto: " & MapInfo(.Pos.map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoNoInvi" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Dim noinvi As Boolean
    
    With UserList(UserIndex)

        
        noinvi = Message.ReadBool()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar Invisibilidad el mapa.")
            MapInfo(UserList(UserIndex).Pos.map).InviSinEfecto = noinvi
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "InviSinEfecto", noinvi)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " InviSinEfecto: " & MapInfo(.Pos.map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub
            
''
' Handle the "ChangeMapInfoNoResu" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    Dim noresu As Boolean
    
    With UserList(UserIndex)

        
        noresu = Message.ReadBool()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar Resucitar el mapa.")
            MapInfo(UserList(UserIndex).Pos.map).ResuSinEfecto = noresu
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "ResuSinEfecto", noresu)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " ResuSinEfecto: " & MapInfo(.Pos.map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    Dim tStr As String
    
    With UserList(UserIndex)

        
        tStr = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.name, .name & " ha cambiado la informacion del Terreno del mapa.")
                MapInfo(UserList(UserIndex).Pos.map).Terreno = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "Terreno", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Terreno: " & MapInfo(.Pos.map).Terreno, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el único útil es 'NIEVE' ya que al ingresarlo, la gente muere de frio en el Mapa", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

    End With


End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    Dim tStr As String
    
    With UserList(UserIndex)

        
        tStr = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.name, .name & " ha cambiado la informacion de la Zona del mapa.")
                MapInfo(UserList(UserIndex).Pos.map).Zona = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "Zona", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Zona: " & MapInfo(.Pos.map).Zona, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el único útil es 'DUNGEON' ya que al ingresarlo, NO se sentirá el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

    End With


End Sub

''
' Handle the "SaveMap" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha guardado el mapa " & CStr(.Pos.map))
        
        Call GrabarMapa(.Pos.map, App.Path & "\WorldBackUp\Mapa" & .Pos.map)
        
        Call WriteConsoleMsg(UserIndex, "Mapa Guardado", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ShowGuildMessages" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowGuildMessages(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim guild As String
        
        guild = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaClan(UserIndex, guild)
        End If

    End With


End Sub

''
' Handle the "DoBackUp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, .name & " ha hecho un backup")
        
        Call ES.DoBackUp 'Sino lo confunde con la id del paquete
    End With
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        'Reads the userName and newUser Packets
        Dim UserName As String
        Dim newName As String
        Dim changeNameUI As Integer
        Dim guildIndex As Integer
        
        UserName = Message.ReadString16()
        newName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
            Else
                changeNameUI = NameIndex(UserName)
                
                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El Pj esta online, debe salir para el cambio", FontTypeNames.FONTTYPE_WARNING)
                Else
                    If Not FileExist(CharPath & UserName & ".chr") Then
                        Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " es inexistente ", FontTypeNames.FONTTYPE_INFO)
                    Else
                        guildIndex = val(GetVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX"))
                        
                        If guildIndex > 0 Then
                            Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If Not FileExist(CharPath & newName & ".chr") Then
                                Call FileCopy(CharPath & UserName & ".chr", CharPath & UCase$(newName) & ".chr")
                                
                                Call WriteConsoleMsg(UserIndex, "Transferencia exitosa", FontTypeNames.FONTTYPE_INFO)
                                
                                Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                                
                                Dim cantPenas As Byte
                                
                                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", CStr(cantPenas + 1))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & CStr(cantPenas + 1), LCase$(.name) & ": BAN POR Cambio de nick a " & UCase$(newName) & " " & Date & " " & time)
                                
                                Call LogGM(.name, "Ha cambiado de nombre al usuario " & UserName & ". Ahora se llama " & newName)
                            Else
                                Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End With


End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterMail(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim newMail As String
        
        UserName = Message.ReadString16()
        newMail = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newMail) = 0 Then
                Call WriteConsoleMsg(UserIndex, "usar /AEMAIL <pj>-<nuevomail>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "No existe el charfile " & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteVar(CharPath & UserName & ".chr", "CONTACTO", "Email", newMail)
                    Call WriteConsoleMsg(UserIndex, "Email de " & UserName & " cambiado a: " & newMail, FontTypeNames.FONTTYPE_INFO)
                End If
                
                Call LogGM(.name, "Le ha cambiado el mail a " & UserName)
            End If
        End If

    End With


End Sub

''
' Handle the "AlterPassword" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterPassword(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim copyFrom As String
        Dim Password As String
        
        UserName = Replace(Message.ReadString16(), "+", " ")
        copyFrom = Replace(Message.ReadString16(), "+", " ")
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "Ha alterado la contraseña de " & UserName)
            
            If LenB(UserName) = 0 Or LenB(copyFrom) = 0 Then
                Call WriteConsoleMsg(UserIndex, "usar /APASS <pjsinpass>@<pjconpass>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not FileExist(CharPath & UserName & ".chr") Or Not FileExist(CharPath & copyFrom & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "Alguno de los PJs no existe " & UserName & "@" & copyFrom, FontTypeNames.FONTTYPE_INFO)
                Else
                    Password = GetVar(CharPath & copyFrom & ".chr", "INIT", "Password")
                    Call WriteVar(CharPath & UserName & ".chr", "INIT", "Password", Password)
                    
                    Call WriteConsoleMsg(UserIndex, "Password de " & UserName & " ha cambiado por la de " & copyFrom, FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

    End With


End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim NpcIndex As Integer
        
        NpcIndex = Message.ReadInt()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)
        
        If NpcIndex <> 0 Then
            Call LogGM(.name, "Sumoneo a " & Npclist(NpcIndex).name & " en mapa " & .Pos.map)
        End If
    End With
End Sub


''
' Handle the "CreateNPCWithRespawn" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal Message As BinaryReader, ByVal UserIndex As Integer)


    
    With UserList(UserIndex)

        
        Dim NpcIndex As Integer
        
        NpcIndex = Message.ReadInt()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)
        
        If NpcIndex <> 0 Then
            Call LogGM(.name, "Sumoneo con respawn " & Npclist(NpcIndex).name & " en mapa " & .Pos.map)
        End If
    End With
End Sub

''
' Handle the "NavigateToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
        Else
            .flags.Navegando = 1
        End If
        
        'Tell the client that we are navigating.
        Call WriteNavigateToggle(UserIndex)
    End With
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If ServerSoloGMs > 0 Then
            Call WriteConsoleMsg(UserIndex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
        Else
            Call WriteConsoleMsg(UserIndex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1
        End If
    End With
End Sub

''
' Handle the "TurnOffServer" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    Dim Handle As Integer
    
    With UserList(UserIndex)

        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.name, "/APAGAR")
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))
        
        'Log
        Handle = FreeFile
        Open App.Path & "\logs\Main.log" For Append Shared As #Handle
        
        Print #Handle, Date & " " & time & " server apagado por " & .name & ". "
        
        Close #Handle
        
        Unload frmMain
    End With
End Sub

''
' Handle the "TurnCriminal" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnCriminal(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "/CONDEN " & UserName)
            
            tUser = NameIndex(UserName)
            If tUser > 0 Then _
                Call VolverCriminal(tUser)
        End If
        
    End With


End Sub

''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "/RAJAR " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then _
                Call ResetFacciones(tUser)
        End If

    End With


End Sub

''
' Handle the "RemoveCharFromGuild" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim guildIndex As Integer
        
        UserName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "/RAJARCLAN " & UserName)
            
            guildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
            
            If guildIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, "No pertenece a ningún clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, guildIndex, PrepareMessageConsoleMsg(UserName & " ha sido expulsado del clan por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))
            End If
        End If

    End With


End Sub

''
' Handle the "RequestCharMail" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRequestCharMail(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim UserName As String
        Dim mail As String
        
        UserName = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If FileExist(CharPath & UserName & ".chr") Then
                mail = GetVar(CharPath & UserName & ".chr", "CONTACTO", "email")
                
                Call WriteConsoleMsg(UserIndex, "Last email de " & UserName & ":" & mail, FontTypeNames.FONTTYPE_INFO)
            End If
        End If

    End With


End Sub

''
' Handle the "SystemMessage" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim Text As String
        Text = Message.ReadString16()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "Mensaje de sistema:" & Text)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(Text))
        End If

    End With


End Sub

''
' Handle the "SetMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetMOTD(ByVal Message As BinaryReader, ByVal UserIndex As Integer)



    With UserList(UserIndex)

        
        Dim newMOTD As String
        Dim auxiliaryString() As String
        Dim LoopC As Long
        
        newMOTD = Message.ReadString16()
        auxiliaryString = Split(newMOTD, vbCrLf)
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "Ha fijado un nuevo MOTD")
            
            MaxLines = UBound(auxiliaryString()) + 1
            
            ReDim MOTD(1 To MaxLines)
            
            Call WriteVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines", CStr(MaxLines))
            
            For LoopC = 1 To MaxLines
                Call WriteVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))
                
                MOTD(LoopC).texto = auxiliaryString(LoopC - 1)
            Next LoopC
            
            Call WriteConsoleMsg(UserIndex, "Se ha cambiado el MOTD con exito", FontTypeNames.FONTTYPE_INFO)
        End If

    End With


End Sub

''
' Handle the "ChangeMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMOTD(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    With UserList(UserIndex)

        
        If (.flags.Privilegios And (PlayerType.RoleMaster Or PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) Then
            Exit Sub
        End If
        
        Dim auxiliaryString As String
        Dim LoopC As Long
        
        For LoopC = LBound(MOTD()) To UBound(MOTD())
            auxiliaryString = auxiliaryString & MOTD(LoopC).texto & vbCrLf
        Next LoopC
        
        If Len(auxiliaryString) >= 2 Then
            If Right$(auxiliaryString, 2) = vbCrLf Then
                auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)
            End If
        End If
        
        Call WriteShowMOTDEditionForm(UserIndex, auxiliaryString)
    End With
End Sub

''
' Handle the "Ping" message
'
' @param userIndex The index of the user sending the message

Public Sub HandlePing(ByVal Message As BinaryReader, ByVal UserIndex As Integer)

    Dim Time As Long
    Time = Message.ReadInt
    
    Call WritePong(UserIndex, Time)
    
End Sub


''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)


    Call Writer_.WriteInt(ServerPacketID.Logged)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.RemoveDialogs)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.NavigateToggle)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.Disconnect)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.CommerceEnd)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.BankEnd)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.CommerceInit)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.BankInit)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.UserCommerceInit)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.UserCommerceEnd)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowBlacksmithForm(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.ShowBlacksmithForm)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCarpenterForm(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.ShowCarpenterForm)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "NPCSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCSwing(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.NPCSwing)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "NPCKillUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCKillUser(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.NPCKillUser)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldUser(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.BlockedWithShieldUser)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldOther(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.BlockedWithShieldOther)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UserSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserSwing(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.UserSwing)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "SafeModeOn" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOn(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.SafeModeOn)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "SafeModeOff" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOff(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.SafeModeOff)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ResuscitationSafeOn" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationSafeOn(ByVal UserIndex As Integer)

    Call Writer_.WriteInt(ServerPacketID.ResuscitationSafeOn)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ResuscitationSafeOff" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationSafeOff(ByVal UserIndex As Integer)

'Author: Rapsodius
'Last Modification: 10/10/07



    Call Writer_.WriteInt(ServerPacketID.ResuscitationSafeOff)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "NobilityLost" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNobilityLost(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.NobilityLost)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UpdateSta)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MinSta)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UpdateMana)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MinMAN)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UpdateHP)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MinHP)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UpdateGold)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.GLD)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UpdateExp)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.Exp)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal map As Integer, ByVal version As Integer)
        Call Writer_.WriteInt(ServerPacketID.ChangeMap)
        Call Writer_.WriteInt(map)
        Call Writer_.WriteInt(version)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.PosUpdate)
        Call Writer_.WriteInt(UserList(UserIndex).Pos.X)
        Call Writer_.WriteInt(UserList(UserIndex).Pos.Y)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "NPCHitUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCHitUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal damage As Integer)
        Call Writer_.WriteInt(ServerPacketID.NPCHitUser)
        Call Writer_.WriteInt(Target)
        Call Writer_.WriteInt(damage)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UserHitNPC" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    damage The number of HP lost by the target creature.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHitNPC(ByVal UserIndex As Integer, ByVal damage As Long)
        Call Writer_.WriteInt(ServerPacketID.UserHitNPC)
        
        'It is a long to allow the "drake slayer" (matadracos) to kill the great red dragon of one blow.
        Call Writer_.WriteInt(damage)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UserAttackedSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex       User to which the message is intended.
' @param    attackerIndex   The user index of the user that attacked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserAttackedSwing(ByVal UserIndex As Integer, ByVal attackerIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UserAttackedSwing)
        Call Writer_.WriteInt(UserList(attackerIndex).Char.CharIndex)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UserHittedByUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    attackerChar Char index of the user hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHittedByUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal attackerChar As Integer, ByVal damage As Integer)
        Call Writer_.WriteInt(ServerPacketID.UserHittedByUser)
        Call Writer_.WriteInt(attackerChar)
        Call Writer_.WriteInt(Target)
        Call Writer_.WriteInt(damage)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UserHittedUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    attackedChar Char index of the user hitted.
' @param    damage The number of HP lost by the oponent hitted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHittedUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal attackedChar As Integer, ByVal damage As Integer)
        Call Writer_.WriteInt(ServerPacketID.UserHittedUser)
        Call Writer_.WriteInt(attackedChar)
        Call Writer_.WriteInt(Target)
        Call Writer_.WriteInt(damage)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal UserIndex As Integer, ByVal chat As String, ByVal CharIndex As Integer, ByVal color As Long)

    Call modSendData.SendData(ToUser, UserIndex, PrepareMessageChatOverHead(chat, CharIndex, color))
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, ByVal chat As String, ByVal FontIndex As FontTypeNames)

    Call modSendData.SendData(ToUser, UserIndex, PrepareMessageConsoleMsg(chat, FontIndex))
End Sub

''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal UserIndex As Integer, ByVal chat As String)


    Call modSendData.SendData(ToUser, UserIndex, PrepareMessageGuildChat(chat))
End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal Message As String)
        Call Writer_.WriteInt(ServerPacketID.ShowMessageBox)
        Call Writer_.WriteString16(Message)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UserIndexInServer)
        Call Writer_.WriteInt(UserIndex)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UserCharIndexInServer)
        Call Writer_.WriteInt(UserList(UserIndex).Char.CharIndex)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal name As String, ByVal criminal As Byte, _
                                ByVal privileges As Byte)



    Call modSendData.SendData(ToUser, UserIndex, PrepareMessageCharacterCreate(body, Head, heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, helmet, name, criminal, privileges))

End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal UserIndex As Integer, ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

    Call modSendData.SendData(ToUser, UserIndex, PrepareMessageObjectCreate(GrhIndex, X, Y))
End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean)
        Call Writer_.WriteInt(ServerPacketID.BlockPosition)
        Call Writer_.WriteInt(X)
        Call Writer_.WriteInt(Y)
        Call Writer_.WriteBool(Blocked)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMidi(ByVal UserIndex As Integer, ByVal midi As Byte)

    Call modSendData.SendData(ToUser, UserIndex, PrepareMessagePlayMidi(midi))
End Sub

''
' Writes the "GuildList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef guildList() As String)




    Dim Tmp As String
    Dim i As Long
        Call Writer_.WriteInt(ServerPacketID.guildList)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call Writer_.WriteString16(Tmp)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.AreaChanged)
        Call Writer_.WriteInt(UserList(UserIndex).Pos.X)
        Call Writer_.WriteInt(UserList(UserIndex).Pos.Y)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal UserIndex As Integer)

    Call modSendData.SendData(ToUser, UserIndex, PrepareMessagePauseToggle())
End Sub

''
' Writes the "RainToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle(ByVal UserIndex As Integer)

    Call modSendData.SendData(ToUser, UserIndex, PrepareMessageRainToggle())
End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UpdateUserStats)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MaxHP)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MinHP)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MaxMAN)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MinMAN)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MaxSta)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MinSta)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.GLD)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.ELV)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.ELU)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.Exp)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkRequestTarget(ByVal UserIndex As Integer, ByVal Skill As eSkill)
        Call Writer_.WriteInt(ServerPacketID.WorkRequestTarget)
        Call Writer_.WriteInt(Skill)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
        Call Writer_.WriteInt(ServerPacketID.ChangeInventorySlot)
        Call Writer_.WriteInt(Slot)
        
        Dim ObjIndex As Integer
        Dim obData As ObjData
        
        ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        
        Call Writer_.WriteInt(ObjIndex)
        Call Writer_.WriteString16(obData.name)
        Call Writer_.WriteInt(UserList(UserIndex).Invent.Object(Slot).amount)
        Call Writer_.WriteBool(UserList(UserIndex).Invent.Object(Slot).Equipped)
        Call Writer_.WriteInt(obData.GrhIndex)
        Call Writer_.WriteInt(obData.OBJType)
        Call Writer_.WriteInt(obData.MaxHIT)
        Call Writer_.WriteInt(obData.MinHIT)
        Call Writer_.WriteInt(obData.def)
        Call Writer_.WriteReal32(SalePrice(obData.Valor))


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
        Call Writer_.WriteInt(ServerPacketID.ChangeBankSlot)
        Call Writer_.WriteInt(Slot)
        
        Dim ObjIndex As Integer
        Dim obData As ObjData
        
        ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex
        
        Call Writer_.WriteInt(ObjIndex)
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        
        Call Writer_.WriteString16(obData.name)
        Call Writer_.WriteInt(UserList(UserIndex).BancoInvent.Object(Slot).amount)
        Call Writer_.WriteInt(obData.GrhIndex)
        Call Writer_.WriteInt(obData.OBJType)
        Call Writer_.WriteInt(obData.MaxHIT)
        Call Writer_.WriteInt(obData.MinHIT)
        Call Writer_.WriteInt(obData.def)
        Call Writer_.WriteInt(obData.Valor)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)
        Call Writer_.WriteInt(ServerPacketID.ChangeSpellSlot)
        Call Writer_.WriteInt(Slot)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.UserHechizos(Slot))
        
        If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
            Call Writer_.WriteString16(Hechizos(UserList(UserIndex).Stats.UserHechizos(Slot)).Nombre)
        Else
            Call Writer_.WriteString16("(None)")
        End If


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.atributes)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call Writer_.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call Writer_.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call Writer_.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
        Call Writer_.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)

    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
        Call Writer_.WriteInt(ServerPacketID.BlacksmithWeapons)
        
        For i = 1 To UBound(ArmasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call Writer_.WriteInt(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ArmasHerrero(validIndexes(i)))
            Call Writer_.WriteString16(Obj.name)
            Call Writer_.WriteInt(Obj.LingH)
            Call Writer_.WriteInt(Obj.LingP)
            Call Writer_.WriteInt(Obj.LingO)
            Call Writer_.WriteInt(ArmasHerrero(validIndexes(i)))
        Next i


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)

    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
        Call Writer_.WriteInt(ServerPacketID.BlacksmithArmors)
        
        For i = 1 To UBound(ArmadurasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call Writer_.WriteInt(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ArmadurasHerrero(validIndexes(i)))
            Call Writer_.WriteString16(Obj.name)
            Call Writer_.WriteInt(Obj.LingH)
            Call Writer_.WriteInt(Obj.LingP)
            Call Writer_.WriteInt(Obj.LingO)
            Call Writer_.WriteInt(ArmadurasHerrero(validIndexes(i)))
        Next i


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCarpenterObjects(ByVal UserIndex As Integer)




    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ObjCarpintero()))
        Call Writer_.WriteInt(ServerPacketID.CarpenterObjects)
        
        For i = 1 To UBound(ObjCarpintero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(UserIndex).clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call Writer_.WriteInt(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ObjCarpintero(validIndexes(i)))
            Call Writer_.WriteString16(Obj.name)
            Call Writer_.WriteInt(Obj.Madera)
            Call Writer_.WriteInt(ObjCarpintero(validIndexes(i)))
        Next i


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.RestOK)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal Message As String)

    Call modSendData.SendData(ToUser, UserIndex, PrepareMessageErrorMsg(Message))
End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Obj As Obj, ByVal price As Single)

'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 06/13/08
'Last Modified by: Nicolas Ezequiel Bouhid (NicoNZ)



    Dim ObjInfo As ObjData
    
    If Obj.ObjIndex >= LBound(ObjData()) And Obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(Obj.ObjIndex)
    End If
        Call Writer_.WriteInt(ServerPacketID.ChangeNPCInventorySlot)
        Call Writer_.WriteInt(Slot)
        Call Writer_.WriteString16(ObjInfo.name)
        Call Writer_.WriteInt(Obj.amount)
        Call Writer_.WriteReal32(price)
        Call Writer_.WriteInt(ObjInfo.GrhIndex)
        Call Writer_.WriteInt(Obj.ObjIndex)
        Call Writer_.WriteInt(ObjInfo.OBJType)
        Call Writer_.WriteInt(ObjInfo.MaxHIT)
        Call Writer_.WriteInt(ObjInfo.MinHIT)
        Call Writer_.WriteInt(ObjInfo.def)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UpdateHungerAndThirst)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MaxAGU)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MinAGU)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MaxHam)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MinHam)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "Fame" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteFame(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.Fame)
        
        Call Writer_.WriteInt(UserList(UserIndex).Reputacion.AsesinoRep)
        Call Writer_.WriteInt(UserList(UserIndex).Reputacion.BandidoRep)
        Call Writer_.WriteInt(UserList(UserIndex).Reputacion.BurguesRep)
        Call Writer_.WriteInt(UserList(UserIndex).Reputacion.LadronesRep)
        Call Writer_.WriteInt(UserList(UserIndex).Reputacion.NobleRep)
        Call Writer_.WriteInt(UserList(UserIndex).Reputacion.PlebeRep)
        Call Writer_.WriteInt(UserList(UserIndex).Reputacion.Promedio)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.MiniStats)
        
        Call Writer_.WriteInt(UserList(UserIndex).Faccion.CiudadanosMatados)
        Call Writer_.WriteInt(UserList(UserIndex).Faccion.CriminalesMatados)
        
'TODO : Este valor es calculable, no debería NI EXISTIR, ya sea en el servidor ni en el cliente!!!
        Call Writer_.WriteInt(UserList(UserIndex).Stats.UsuariosMatados)
        
        Call Writer_.WriteInt(UserList(UserIndex).Stats.NPCsMuertos)
        
        Call Writer_.WriteInt(UserList(UserIndex).clase)
        Call Writer_.WriteInt(UserList(UserIndex).Counters.Pena)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data buffer.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)
        Call Writer_.WriteInt(ServerPacketID.LevelUp)
        Call Writer_.WriteInt(skillPoints)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean)

    Call modSendData.SendData(ToUser, UserIndex, PrepareMessageSetInvisible(CharIndex, invisible))
End Sub

''
' Writes the "DiceRoll" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDiceRoll(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.DiceRoll)
        
        Call Writer_.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call Writer_.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call Writer_.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call Writer_.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
        Call Writer_.WriteInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.MeditateToggle)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal UserIndex As Integer)




    Dim i As Long
        Call Writer_.WriteInt(ServerPacketID.SendSkills)
        
        For i = 1 To NUMSKILLS
            Call Writer_.WriteInt(UserList(UserIndex).Stats.UserSkills(i))
        Next i


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)




    Dim i As Long
    Dim str As String
        Call Writer_.WriteInt(ServerPacketID.TrainerCreatureList)
        
        For i = 1 To Npclist(NpcIndex).NroCriaturas
            str = str & Npclist(NpcIndex).Criaturas(i).NpcName & SEPARATOR
        Next i
        
        If LenB(str) > 0 Then _
            str = Left$(str, Len(str) - 1)
        
        Call Writer_.WriteString16(str)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "GuildNews" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildNews The guild's news.
' @param    enemies The list of the guild's enemies.
' @param    allies The list of the guild's allies.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNews(ByVal UserIndex As Integer, ByVal guildNews As String, ByRef enemies() As String, ByRef allies() As String)




    Dim i As Long
    Dim Tmp As String
        Call Writer_.WriteInt(ServerPacketID.guildNews)
        
        Call Writer_.WriteString16(guildNews)
        
        'Prepare enemies' list
        For i = LBound(enemies()) To UBound(enemies())
            Tmp = Tmp & enemies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call Writer_.WriteString16(Tmp)
        
        Tmp = vbNullString
        'Prepare allies' list
        For i = LBound(allies()) To UBound(allies())
            Tmp = Tmp & allies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call Writer_.WriteString16(Tmp)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)




        Call Writer_.WriteInt(ServerPacketID.OfferDetails)
        
        Call Writer_.WriteString16(details)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlianceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)




    Dim i As Long
    Dim Tmp As String
        Call Writer_.WriteInt(ServerPacketID.AlianceProposalsList)
        
        ' Prepare guild's list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call Writer_.WriteString16(Tmp)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePeaceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)




    Dim i As Long
    Dim Tmp As String
        Call Writer_.WriteInt(ServerPacketID.PeaceProposalsList)
                
        ' Prepare guilds' list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call Writer_.WriteString16(Tmp)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "CharacterInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    charName The requested char's name.
' @param    race The requested char's race.
' @param    class The requested char's class.
' @param    gender The requested char's gender.
' @param    level The requested char's level.
' @param    gold The requested char's gold.
' @param    reputation The requested char's reputation.
' @param    previousPetitions The requested char's previous petitions to enter guilds.
' @param    currentGuild The requested char's current guild.
' @param    previousGuilds The requested char's previous guilds.
' @param    RoyalArmy True if tha char belongs to the Royal Army.
' @param    CaosLegion True if tha char belongs to the Caos Legion.
' @param    citicensKilled The number of citicens killed by the requested char.
' @param    criminalsKilled The number of criminals killed by the requested char.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterInfo(ByVal UserIndex As Integer, ByVal charName As String, ByVal race As eRaza, ByVal Class As eClass, _
                            ByVal gender As eGenero, ByVal level As Byte, ByVal gold As Long, ByVal bank As Long, ByVal reputation As Long, _
                            ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, _
                            ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)
        Call Writer_.WriteInt(ServerPacketID.CharacterInfo)
        
        Call Writer_.WriteString16(charName)
        Call Writer_.WriteInt(race)
        Call Writer_.WriteInt(Class)
        Call Writer_.WriteInt(gender)
        
        Call Writer_.WriteInt(level)
        Call Writer_.WriteInt(gold)
        Call Writer_.WriteInt(bank)
        Call Writer_.WriteInt(reputation)
        
        Call Writer_.WriteString16(previousPetitions)
        Call Writer_.WriteString16(currentGuild)
        Call Writer_.WriteString16(previousGuilds)
        
        Call Writer_.WriteBool(RoyalArmy)
        Call Writer_.WriteBool(CaosLegion)
        
        Call Writer_.WriteInt(citicensKilled)
        Call Writer_.WriteInt(criminalsKilled)

End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @param    guildNews The guild's news.
' @param    joinRequests The list of chars which requested to join the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeaderInfo(ByVal UserIndex As Integer, ByRef guildList() As String, ByRef MemberList() As String, _
                            ByVal guildNews As String, ByRef joinRequests() As String)




    Dim i As Long
    Dim Tmp As String
        Call Writer_.WriteInt(ServerPacketID.GuildLeaderInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call Writer_.WriteString16(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString
        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call Writer_.WriteString16(Tmp)
        
        ' Store guild news
        Call Writer_.WriteString16(guildNews)
        
        ' Prepare the join request's list
        Tmp = vbNullString
        For i = LBound(joinRequests()) To UBound(joinRequests())
            Tmp = Tmp & joinRequests(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call Writer_.WriteString16(Tmp)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "GuildDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildName The requested guild's name.
' @param    founder The requested guild's founder.
' @param    foundationDate The requested guild's foundation date.
' @param    leader The requested guild's current leader.
' @param    URL The requested guild's website.
' @param    memberCount The requested guild's member count.
' @param    electionsOpen True if the clan is electing it's new leader.
' @param    alignment The requested guild's alignment.
' @param    enemiesCount The requested guild's enemy count.
' @param    alliesCount The requested guild's ally count.
' @param    antifactionPoints The requested guild's number of antifaction acts commited.
' @param    codex The requested guild's codex.
' @param    guildDesc The requested guild's description.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDetails(ByVal UserIndex As Integer, ByVal GuildName As String, ByVal founder As String, ByVal foundationDate As String, _
                            ByVal leader As String, ByVal URL As String, ByVal memberCount As Integer, ByVal electionsOpen As Boolean, _
                            ByVal alignment As String, ByVal enemiesCount As Integer, ByVal AlliesCount As Integer, _
                            ByVal antifactionPoints As String, ByRef codex() As String, ByVal guildDesc As String)




    Dim i As Long
    Dim temp As String
        Call Writer_.WriteInt(ServerPacketID.GuildDetails)
        
        Call Writer_.WriteString16(GuildName)
        Call Writer_.WriteString16(founder)
        Call Writer_.WriteString16(foundationDate)
        Call Writer_.WriteString16(leader)
        Call Writer_.WriteString16(URL)
        
        Call Writer_.WriteInt(memberCount)
        Call Writer_.WriteBool(electionsOpen)
        
        Call Writer_.WriteString16(alignment)
        
        Call Writer_.WriteInt(enemiesCount)
        Call Writer_.WriteInt(AlliesCount)
        
        Call Writer_.WriteString16(antifactionPoints)
        
        For i = LBound(codex()) To UBound(codex())
            temp = temp & codex(i) & SEPARATOR
        Next i
        
        If Len(temp) > 1 Then _
            temp = Left$(temp, Len(temp) - 1)
        
        Call Writer_.WriteString16(temp)
        
        Call Writer_.WriteString16(guildDesc)

End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildFundationForm(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.ShowGuildFundationForm)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal UserIndex As Integer)

'And updates user position


    Call Writer_.WriteInt(ServerPacketID.ParalizeOK)
    Call WritePosUpdate(UserIndex)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)
        Call Writer_.WriteInt(ServerPacketID.ShowUserRequest)
        
        Call Writer_.WriteString16(details)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "TradeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTradeOK(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.TradeOK)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "BankOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankOK(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.BankOK)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal amount As Long)
        Call Writer_.WriteInt(ServerPacketID.ChangeUserTradeSlot)
        
        Call Writer_.WriteInt(ObjIndex)
        Call Writer_.WriteString16(ObjData(ObjIndex).name)
        Call Writer_.WriteInt(amount)
        Call Writer_.WriteInt(ObjData(ObjIndex).GrhIndex)
        Call Writer_.WriteInt(ObjData(ObjIndex).OBJType)
        Call Writer_.WriteInt(ObjData(ObjIndex).MaxHIT)
        Call Writer_.WriteInt(ObjData(ObjIndex).MinHIT)
        Call Writer_.WriteInt(ObjData(ObjIndex).def)
        Call Writer_.WriteInt(SalePrice(ObjData(ObjIndex).Valor))


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByRef npcNames() As String)




    Dim i As Long
    Dim Tmp As String
        Call Writer_.WriteInt(ServerPacketID.SpawnList)
        
        For i = LBound(npcNames()) To UBound(npcNames())
            Tmp = Tmp & npcNames(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call Writer_.WriteString16(Tmp)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)




    Dim i As Long
    Dim Tmp As String
        Call Writer_.WriteInt(ServerPacketID.ShowSOSForm)
        
        For i = 1 To Ayuda.Longitud
            Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i
        
        If LenB(Tmp) <> 0 Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call Writer_.WriteString16(Tmp)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, ByVal currentMOTD As String)
        Call Writer_.WriteInt(ServerPacketID.ShowMOTDEditionForm)
        
        Call Writer_.WriteString16(currentMOTD)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.ShowGMPanelForm)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal UserIndex As Integer, ByRef userNamesList() As String, ByVal cant As Integer)

    Dim i As Long
    Dim Tmp As String
        Call Writer_.WriteInt(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
        For i = 1 To cant
            Tmp = Tmp & userNamesList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call Writer_.WriteString16(Tmp)


    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal UserIndex As Integer, ByVal Time As Long)




    Call Writer_.WriteInt(ServerPacketID.Pong)
    Call Writer_.WriteInt(Time)



    Call modSendData.SendData(ToUser, UserIndex, Writer_)
End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal invisible As Boolean) As BinaryWriter

'Prepares the "SetInvisible" message and returns it.
        Call Writer_.WriteInt(ServerPacketID.SetInvisible)
        
        Call Writer_.WriteInt(CharIndex)
        Call Writer_.WriteBool(invisible)
        
        


    Set PrepareMessageSetInvisible = Writer_
End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal chat As String, ByVal CharIndex As Integer, ByVal color As Long) As BinaryWriter

'Prepares the "ChatOverHead" message and returns it.
        Call Writer_.WriteInt(ServerPacketID.ChatOverHead)
        Call Writer_.WriteString16(chat)
        Call Writer_.WriteInt(CharIndex)
        
        ' Write rgb channels and save one byte from long :D
        Call Writer_.WriteInt(color And &HFF)
        Call Writer_.WriteInt((color And &HFF00&) \ &H100&)
        Call Writer_.WriteInt((color And &HFF0000) \ &H10000)
        
        


    Set PrepareMessageChatOverHead = Writer_
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal chat As String, ByVal FontIndex As FontTypeNames) As BinaryWriter

'Prepares the "ConsoleMsg" message and returns it.
        Call Writer_.WriteInt(ServerPacketID.ConsoleMsg)
        Call Writer_.WriteString16(chat)
        Call Writer_.WriteInt(FontIndex)
        
        


    Set PrepareMessageConsoleMsg = Writer_
End Function

''
' Prepares the "CreateFX" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer) As BinaryWriter

'Prepares the "CreateFX" message and returns it
        Call Writer_.WriteInt(ServerPacketID.CreateFX)
        Call Writer_.WriteInt(CharIndex)
        Call Writer_.WriteInt(FX)
        Call Writer_.WriteInt(FXLoops)
        
        


    Set PrepareMessageCreateFX = Writer_
End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte) As BinaryWriter

'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
        Call Writer_.WriteInt(ServerPacketID.PlayWave)
        Call Writer_.WriteInt(wave)
        Call Writer_.WriteInt(X)
        Call Writer_.WriteInt(Y)
        
        


    Set PrepareMessagePlayWave = Writer_
End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal chat As String) As BinaryWriter

'Prepares the "GuildChat" message and returns it
        Call Writer_.WriteInt(ServerPacketID.GuildChat)
        Call Writer_.WriteString16(chat)
        
        


    Set PrepareMessageGuildChat = Writer_
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal chat As String) As BinaryWriter

'Prepares the "ShowMessageBox" message and returns it
        Call Writer_.WriteInt(ServerPacketID.ShowMessageBox)
        Call Writer_.WriteString16(chat)
        
        


    Set PrepareMessageShowMessageBox = Writer_
End Function


''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMidi(ByVal midi As Byte) As BinaryWriter


        Call Writer_.WriteInt(ServerPacketID.PlayMidi)
        Call Writer_.WriteInt(midi)

        


    Set PrepareMessagePlayMidi = Writer_
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As BinaryWriter

'Prepares the "PauseToggle" message and returns it
        Call Writer_.WriteInt(ServerPacketID.PauseToggle)
        


    Set PrepareMessagePauseToggle = Writer_
End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRainToggle() As BinaryWriter

'Prepares the "RainToggle" message and returns it
        Call Writer_.WriteInt(ServerPacketID.RainToggle)
        
        


    Set PrepareMessageRainToggle = Writer_
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As BinaryWriter

'Prepares the "ObjectDelete" message and returns it
        Call Writer_.WriteInt(ServerPacketID.ObjectDelete)
        Call Writer_.WriteInt(X)
        Call Writer_.WriteInt(Y)
        
        


    Set PrepareMessageObjectDelete = Writer_
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean) As BinaryWriter

'Prepares the "BlockPosition" message and returns it
        Call Writer_.WriteInt(ServerPacketID.BlockPosition)
        Call Writer_.WriteInt(X)
        Call Writer_.WriteInt(Y)
        Call Writer_.WriteBool(Blocked)
        
        


    Set PrepareMessageBlockPosition = Writer_
End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As BinaryWriter

'prepares the "ObjectCreate" message and returns it
        Call Writer_.WriteInt(ServerPacketID.ObjectCreate)
        Call Writer_.WriteInt(X)
        Call Writer_.WriteInt(Y)
        Call Writer_.WriteInt(GrhIndex)
        
        


    Set PrepareMessageObjectCreate = Writer_
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As BinaryWriter

'Prepares the "CharacterRemove" message and returns it
        Call Writer_.WriteInt(ServerPacketID.CharacterRemove)
        Call Writer_.WriteInt(CharIndex)
        
        


    Set PrepareMessageCharacterRemove = Writer_
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As BinaryWriter
        Call Writer_.WriteInt(ServerPacketID.RemoveCharDialog)
        Call Writer_.WriteInt(CharIndex)
        
        


    Set PrepareMessageRemoveCharDialog = Writer_
End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterCreate(ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal name As String, ByVal criminal As Byte, _
                                ByVal privileges As Byte) As BinaryWriter

'Prepares the "CharacterCreate" message and returns it
        Call Writer_.WriteInt(ServerPacketID.CharacterCreate)
        
        Call Writer_.WriteInt(CharIndex)
        Call Writer_.WriteInt(body)
        Call Writer_.WriteInt(Head)
        Call Writer_.WriteInt(heading)
        Call Writer_.WriteInt(X)
        Call Writer_.WriteInt(Y)
        Call Writer_.WriteInt(weapon)
        Call Writer_.WriteInt(shield)
        Call Writer_.WriteInt(helmet)
        Call Writer_.WriteInt(FX)
        Call Writer_.WriteInt(FXLoops)
        Call Writer_.WriteString16(name)
        Call Writer_.WriteInt(criminal)
        Call Writer_.WriteInt(privileges)
        
        


    Set PrepareMessageCharacterCreate = Writer_
End Function

''
' Prepares the "CharacterChange" message and returns it.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterChange(ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer) As BinaryWriter

'Prepares the "CharacterChange" message and returns it
        Call Writer_.WriteInt(ServerPacketID.CharacterChange)
        
        Call Writer_.WriteInt(CharIndex)
        Call Writer_.WriteInt(body)
        Call Writer_.WriteInt(Head)
        Call Writer_.WriteInt(heading)
        Call Writer_.WriteInt(weapon)
        Call Writer_.WriteInt(shield)
        Call Writer_.WriteInt(helmet)
        Call Writer_.WriteInt(FX)
        Call Writer_.WriteInt(FXLoops)
        
        


    Set PrepareMessageCharacterChange = Writer_
End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As BinaryWriter

'Prepares the "CharacterMove" message and returns it
        Call Writer_.WriteInt(ServerPacketID.CharacterMove)
        Call Writer_.WriteInt(CharIndex)
        Call Writer_.WriteInt(X)
        Call Writer_.WriteInt(Y)
        
        


    Set PrepareMessageCharacterMove = Writer_
End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, isCriminal As Boolean, Tag As String) As BinaryWriter

'Prepares the "UpdateTagAndStatus" message and returns it
        Call Writer_.WriteInt(ServerPacketID.UpdateTagAndStatus)
        
        Call Writer_.WriteInt(UserList(UserIndex).Char.CharIndex)
        Call Writer_.WriteBool(isCriminal)
        Call Writer_.WriteString16(Tag)
        
        


    Set PrepareMessageUpdateTagAndStatus = Writer_
End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal Message As String) As BinaryWriter

'Prepares the "ErrorMsg" message and returns it
        Call Writer_.WriteInt(ServerPacketID.ErrorMsg)
        Call Writer_.WriteString16(Message)
        
        


    Set PrepareMessageErrorMsg = Writer_
End Function


