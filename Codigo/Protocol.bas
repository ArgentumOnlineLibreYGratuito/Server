Attribute VB_Name = "Protocol"


Option Explicit

Private Writer_ As BinaryWriter

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.Logged)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.RemoveDialogs)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)




    Call Writer_.WriteString16(PrepareMessageRemoveCharDialog(CharIndex))


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.NavigateToggle)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.Disconnect)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.CommerceEnd)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.BankEnd)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.CommerceInit)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.BankInit)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.UserCommerceInit)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.UserCommerceEnd)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowBlacksmithForm(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.ShowBlacksmithForm)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCarpenterForm(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.ShowCarpenterForm)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "NPCSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCSwing(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.NPCSwing)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "NPCKillUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCKillUser(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.NPCKillUser)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldUser(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.BlockedWithShieldUser)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldOther(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.BlockedWithShieldOther)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "UserSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserSwing(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.UserSwing)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "UpdateNeeded" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateNeeded(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.UpdateNeeded)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "SafeModeOn" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOn(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.SafeModeOn)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "SafeModeOff" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOff(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.SafeModeOff)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "ResuscitationSafeOn" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationSafeOn(ByVal UserIndex As Integer)

    Call Writer_.WriteInt(ServerPacketID.ResuscitationSafeOn)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "NobilityLost" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNobilityLost(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.NobilityLost)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCantUseWhileMeditating(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.CantUseWhileMeditating)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UpdateSta)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MinSta)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UpdateMana)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MinMAN)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UpdateHP)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.MinHP)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UpdateGold)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.GLD)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UpdateExp)
        Call Writer_.WriteInt(UserList(UserIndex).Stats.Exp)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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




    Call Writer_.WriteString16(PrepareMessageChatOverHead(chat, CharIndex, color))


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, ByVal chat As String, ByVal FontIndex As FontTypeNames)




    Call Writer_.WriteString16(PrepareMessageConsoleMsg(chat, FontIndex))


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal UserIndex As Integer, ByVal chat As String)




    Call Writer_.WriteString16(PrepareMessageGuildChat(chat))


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UserIndexInServer)
        Call Writer_.WriteInt(UserIndex)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.UserCharIndexInServer)
        Call Writer_.WriteInt(UserList(UserIndex).Char.CharIndex)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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




    Call Writer_.WriteString16(PrepareMessageCharacterCreate(body, Head, heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, _
                                                            helmet, name, criminal, privileges))

End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal CharIndex As Integer)




    Call Writer_.WriteString16(PrepareMessageCharacterRemove(CharIndex))


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte)




    Call Writer_.WriteString16(PrepareMessageCharacterMove(CharIndex, X, Y))


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer)




    Call Writer_.WriteString16(PrepareMessageCharacterChange(body, Head, heading, CharIndex, weapon, shield, FX, FXLoops, helmet))

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




    Call Writer_.WriteString16(PrepareMessageObjectCreate(GrhIndex, X, Y))


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)




    Call Writer_.WriteString16(PrepareMessageObjectDelete(X, Y))


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMidi(ByVal UserIndex As Integer, ByVal midi As Byte, Optional ByVal loops As Integer = -1)




    Call Writer_.WriteString16(PrepareMessagePlayMidi(midi, loops))


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal UserIndex As Integer, ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte)

    Call Writer_.WriteString16(PrepareMessagePlayWave(wave, X, Y))


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal UserIndex As Integer)




    Call Writer_.WriteString16(PrepareMessagePauseToggle())


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "RainToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle(ByVal UserIndex As Integer)




    Call Writer_.WriteString16(PrepareMessageRainToggle())


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)




    Call Writer_.WriteString16(PrepareMessageCreateFX(CharIndex, FX, FXLoops))


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.RestOK)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal Message As String)




    Call Writer_.WriteString16(PrepareMessageErrorMsg(Message))


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.Blind)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.Dumb)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
        Call Writer_.WriteInt(ServerPacketID.ShowSignal)
        Call Writer_.WriteString16(ObjData(ObjIndex).texto)
        Call Writer_.WriteInt(ObjData(ObjIndex).GrhSecundario)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data buffer.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)
        Call Writer_.WriteInt(ServerPacketID.LevelUp)
        Call Writer_.WriteInt(skillPoints)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data buffer.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAddForumMsg(ByVal UserIndex As Integer, ByVal title As String, ByVal Message As String)
        Call Writer_.WriteInt(ServerPacketID.AddForumMsg)
        Call Writer_.WriteString16(title)
        Call Writer_.WriteString16(Message)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowForumForm(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.ShowForumForm)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean)




    Call Writer_.WriteString16(PrepareMessageSetInvisible(CharIndex, invisible))


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.MeditateToggle)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.BlindNoMore)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.DumbNoMore)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)




    Dim i As Long
        Call Writer_.WriteInt(ServerPacketID.OfferDetails)
        
        Call Writer_.WriteString16(details)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "TradeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTradeOK(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.TradeOK)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "BankOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankOK(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.BankOK)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "SendNight" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendNight(ByVal UserIndex As Integer, ByVal night As Boolean)

'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
        Call Writer_.WriteInt(ServerPacketID.SendNight)
        Call Writer_.WriteBool(night)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.ShowGMPanelForm)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal UserIndex As Integer, ByRef userNamesList() As String, ByVal cant As Integer)
 NIGO:



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


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal UserIndex As Integer)




    Call Writer_.WriteInt(ServerPacketID.Pong)


    Call modSendData.SendData(ToUser, UserIndex, "Empty")
End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal invisible As Boolean) As String

'Prepares the "SetInvisible" message and returns it.
        Call Writer_.WriteInt(ServerPacketID.SetInvisible)
        
        Call Writer_.WriteInt(CharIndex)
        Call Writer_.WriteBool(invisible)
        
        

End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal chat As String, ByVal CharIndex As Integer, ByVal color As Long) As String

'Prepares the "ChatOverHead" message and returns it.
        Call Writer_.WriteInt(ServerPacketID.ChatOverHead)
        Call Writer_.WriteString16(chat)
        Call Writer_.WriteInt(CharIndex)
        
        ' Write rgb channels and save one byte from long :D
        Call Writer_.WriteInt(color And &HFF)
        Call Writer_.WriteInt((color And &HFF00&) \ &H100&)
        Call Writer_.WriteInt((color And &HFF0000) \ &H10000)
        
        

End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal chat As String, ByVal FontIndex As FontTypeNames) As String

'Prepares the "ConsoleMsg" message and returns it.
        Call Writer_.WriteInt(ServerPacketID.ConsoleMsg)
        Call Writer_.WriteString16(chat)
        Call Writer_.WriteInt(FontIndex)
        
        

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

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer) As String

'Prepares the "CreateFX" message and returns it
        Call Writer_.WriteInt(ServerPacketID.CreateFX)
        Call Writer_.WriteInt(CharIndex)
        Call Writer_.WriteInt(FX)
        Call Writer_.WriteInt(FXLoops)
        
        

End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte) As String

'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
        Call Writer_.WriteInt(ServerPacketID.PlayWave)
        Call Writer_.WriteInt(wave)
        Call Writer_.WriteInt(X)
        Call Writer_.WriteInt(Y)
        
        

End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal chat As String) As String

'Prepares the "GuildChat" message and returns it
        Call Writer_.WriteInt(ServerPacketID.GuildChat)
        Call Writer_.WriteString16(chat)
        
        

End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal chat As String) As String

'Prepares the "ShowMessageBox" message and returns it
        Call Writer_.WriteInt(ServerPacketID.ShowMessageBox)
        Call Writer_.WriteString16(chat)
        
        

End Function


''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMidi(ByVal midi As Byte, Optional ByVal loops As Integer = -1) As String

'Prepares the "GuildChat" message and returns it
        Call Writer_.WriteInt(ServerPacketID.PlayMidi)
        Call Writer_.WriteInt(midi)
        Call Writer_.WriteInt(loops)
        
        

End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String

'Prepares the "PauseToggle" message and returns it
        Call Writer_.WriteInt(ServerPacketID.PauseToggle)
        

End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRainToggle() As String

'Prepares the "RainToggle" message and returns it
        Call Writer_.WriteInt(ServerPacketID.RainToggle)
        
        

End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As String

'Prepares the "ObjectDelete" message and returns it
        Call Writer_.WriteInt(ServerPacketID.ObjectDelete)
        Call Writer_.WriteInt(X)
        Call Writer_.WriteInt(Y)
        
        

End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean) As String

'Prepares the "BlockPosition" message and returns it
        Call Writer_.WriteInt(ServerPacketID.BlockPosition)
        Call Writer_.WriteInt(X)
        Call Writer_.WriteInt(Y)
        Call Writer_.WriteBool(Blocked)
        
        

End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String

'prepares the "ObjectCreate" message and returns it
        Call Writer_.WriteInt(ServerPacketID.ObjectCreate)
        Call Writer_.WriteInt(X)
        Call Writer_.WriteInt(Y)
        Call Writer_.WriteInt(GrhIndex)
        
        

End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As String

'Prepares the "CharacterRemove" message and returns it
        Call Writer_.WriteInt(ServerPacketID.CharacterRemove)
        Call Writer_.WriteInt(CharIndex)
        
        

End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String
        Call Writer_.WriteInt(ServerPacketID.RemoveCharDialog)
        Call Writer_.WriteInt(CharIndex)
        
        

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
                                ByVal privileges As Byte) As String

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
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer) As String

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
        
        

End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String

'Prepares the "CharacterMove" message and returns it
        Call Writer_.WriteInt(ServerPacketID.CharacterMove)
        Call Writer_.WriteInt(CharIndex)
        Call Writer_.WriteInt(X)
        Call Writer_.WriteInt(Y)
        
        

End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, isCriminal As Boolean, Tag As String) As String

'Prepares the "UpdateTagAndStatus" message and returns it
        Call Writer_.WriteInt(ServerPacketID.UpdateTagAndStatus)
        
        Call Writer_.WriteInt(UserList(UserIndex).Char.CharIndex)
        Call Writer_.WriteBool(isCriminal)
        Call Writer_.WriteString16(Tag)
        
        

End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal Message As String) As String

'Prepares the "ErrorMsg" message and returns it
        Call Writer_.WriteInt(ServerPacketID.ErrorMsg)
        Call Writer_.WriteString16(Message)
        
        

End Function
