Attribute VB_Name = "modSendData"
'**************************************************************
' SendData.bas - Has all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' Implemented by Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

''
' Contains all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20070107

Option Explicit

Public Enum SendTarget
    ToUser = 1
    ToAll
    toMap
    ToPCArea
    ToAllButIndex
    ToMapButIndex
    ToGM
    ToNPCArea
    ToGuildMembers
    ToAdmins
    ToPCAreaButIndex
    ToAdminsAreaButConsejeros
    ToDiosesYclan
    ToConsejo
    ToClanArea
    ToConsejoCaos
    ToRolesMasters
    ToDeadArea
    ToCiudadanos
    ToCriminales
    ToPartyArea
    ToReal
    ToCaos
    ToCiudadanosYRMs
    ToCriminalesYRMs
    ToRealYRMs
    ToCaosYRMs
End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndData As BinaryWriter)
On Error GoTo OnException

    Dim LoopC As Long
    
    Select Case sndRoute
        Case SendTarget.ToUser
            If UserList(sndIndex).ConnID <> -1 Then
                Call UserList(sndIndex).Connection.Write(sndData)
            End If

        Case SendTarget.ToPCArea
            Call SendToUserArea(sndIndex, sndData)

        Case SendTarget.ToAdmins
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
                        Call UserList(LoopC).Connection.Write(sndData)
                   End If
                End If
            Next LoopC

        Case SendTarget.ToAll
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call UserList(LoopC).Connection.Write(sndData)
                    End If
                End If
            Next LoopC

        Case SendTarget.ToAllButIndex
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call UserList(LoopC).Connection.Write(sndData)
                    End If
                End If
            Next LoopC

        Case SendTarget.toMap
            Call SendToMap(sndIndex, sndData)

        Case SendTarget.ToMapButIndex
            Call SendToMapButIndex(sndIndex, sndData)

        Case SendTarget.ToGuildMembers
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            While LoopC > 0
                If (UserList(LoopC).ConnID <> -1) Then
                    Call UserList(LoopC).Connection.Write(sndData)
                End If
                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend

        Case SendTarget.ToDeadArea
            Call SendToDeadUserArea(sndIndex, sndData)

        Case SendTarget.ToPCAreaButIndex
            Call SendToUserAreaButindex(sndIndex, sndData)

        Case SendTarget.ToClanArea
            Call SendToUserGuildArea(sndIndex, sndData)

        Case SendTarget.ToPartyArea
            Call SendToUserPartyArea(sndIndex, sndData)

        Case SendTarget.ToAdminsAreaButConsejeros
            Call SendToAdminsButConsejerosArea(sndIndex, sndData)

        Case SendTarget.ToNPCArea
            Call SendToNpcArea(sndIndex, sndData)

        Case SendTarget.ToDiosesYclan
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            While LoopC > 0
                If (UserList(LoopC).ConnID <> -1) Then
                    Call UserList(LoopC).Connection.Write(sndData)
                End If
                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            
            LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
            While LoopC > 0
                If (UserList(LoopC).ConnID <> -1) Then
                    Call UserList(LoopC).Connection.Write(sndData)
                End If
                LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
            Wend

        Case SendTarget.ToConsejo
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call UserList(LoopC).Connection.Write(sndData)
                    End If
                End If
            Next LoopC
 
        Case SendTarget.ToConsejoCaos
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call UserList(LoopC).Connection.Write(sndData)
                    End If
                End If
            Next LoopC

        Case SendTarget.ToRolesMasters
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster Then
                        Call UserList(LoopC).Connection.Write(sndData)
                    End If
                End If
            Next LoopC

        Case SendTarget.ToCiudadanos
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If Not criminal(LoopC) Then
                        Call UserList(LoopC).Connection.Write(sndData)
                    End If
                End If
            Next LoopC

        Case SendTarget.ToCriminales
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If criminal(LoopC) Then
                        Call UserList(LoopC).Connection.Write(sndData)
                    End If
                End If
            Next LoopC

        Case SendTarget.ToReal
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Then
                        Call UserList(LoopC).Connection.Write(sndData)
                    End If
                End If
            Next LoopC

        Case SendTarget.ToCaos
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
                        Call UserList(LoopC).Connection.Write(sndData)
                    End If
                End If
            Next LoopC

        Case SendTarget.ToCiudadanosYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If Not criminal(LoopC) Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call UserList(LoopC).Connection.Write(sndData)
                    End If
                End If
            Next LoopC

        Case SendTarget.ToCriminalesYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If criminal(LoopC) Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call UserList(LoopC).Connection.Write(sndData)
                    End If
                End If
            Next LoopC

        Case SendTarget.ToRealYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call UserList(LoopC).Connection.Write(sndData)
                    End If
                End If
            Next LoopC

        Case SendTarget.ToCaosYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call UserList(LoopC).Connection.Write(sndData)
                    End If
                End If
            Next LoopC
    End Select

OnException:
    
    Call sndData.Clear
End Sub

Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sndData As BinaryWriter)
On Error GoTo OnException

    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida Then
                    Call UserList(tempIndex).Connection.Write(sndData)
                End If
            End If
        End If
    Next LoopC

OnException:
    
    Call sndData.Clear
    
End Sub

Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal sndData As BinaryWriter)
On Error GoTo OnException

    Dim LoopC As Long
    Dim TempInt As Integer
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

    If Not MapaValido(map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
            
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
            If TempInt Then
                If tempIndex <> UserIndex Then
                    If UserList(tempIndex).ConnIDValida Then
                        Call UserList(tempIndex).Connection.Write(sndData)
                    End If
                End If
            End If
        End If
    Next LoopC

OnException:
    
    Call sndData.Clear
    
End Sub

Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, ByVal sndData As BinaryWriter)
On Error GoTo OnException

    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                'Dead and admins read
                If UserList(tempIndex).ConnIDValida = True And (UserList(tempIndex).flags.Muerto = 1 Or (UserList(tempIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) <> 0) Then
                    Call UserList(tempIndex).Connection.Write(sndData)
                End If
            End If
        End If
    Next LoopC

OnException:
    
    Call sndData.Clear
    
End Sub

Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal sndData As BinaryWriter)
On Error GoTo OnException

    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    If UserList(UserIndex).guildIndex = 0 Then Exit Sub
    
    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida And UserList(tempIndex).guildIndex = UserList(UserIndex).guildIndex Then
                    Call UserList(tempIndex).Connection.Write(sndData)
                End If
            End If
        End If
    Next LoopC

OnException:
    
    Call sndData.Clear
    
End Sub

Private Sub SendToUserPartyArea(ByVal UserIndex As Integer, ByVal sndData As BinaryWriter)
On Error GoTo OnException

    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    If UserList(UserIndex).PartyIndex = 0 Then Exit Sub
    
    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida And UserList(tempIndex).PartyIndex = UserList(UserIndex).PartyIndex Then
                    Call UserList(tempIndex).Connection.Write(sndData)
                End If
            End If
        End If
    Next LoopC

OnException:
    
    Call sndData.Clear
    
End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal UserIndex As Integer, ByVal sndData As BinaryWriter)
On Error GoTo OnException

    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida Then
                    If UserList(tempIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then _
                        Call UserList(tempIndex).Connection.Write(sndData)
                End If
            End If
        End If
    Next LoopC

OnException:
    
    Call sndData.Clear
    
End Sub

Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sndData As BinaryWriter)
On Error GoTo OnException

    Dim LoopC As Long
    Dim TempInt As Integer
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = Npclist(NpcIndex).Pos.map
    AreaX = Npclist(NpcIndex).AreasInfo.AreaPerteneceX
    AreaY = Npclist(NpcIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
            If TempInt Then
                If UserList(tempIndex).ConnIDValida Then
                    Call UserList(tempIndex).Connection.Write(sndData)
                End If
            End If
        End If
    Next LoopC

OnException:
    
    Call sndData.Clear
    
End Sub

Public Sub SendToAreaByPos(ByVal map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ByVal sndData As BinaryWriter)
On Error GoTo OnException

    Dim LoopC As Long
    Dim TempInt As Integer
    Dim tempIndex As Integer
    
    AreaX = 2 ^ (AreaX \ 9)
    AreaY = 2 ^ (AreaY \ 9)
    
    If Not MapaValido(map) Then Exit Sub

    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
            
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
            If TempInt Then
                If UserList(tempIndex).ConnIDValida Then
                    Call UserList(tempIndex).Connection.Write(sndData)
                End If
            End If
        End If
    Next LoopC

OnException:
    
    Call sndData.Clear
    
End Sub

Public Sub SendToMap(ByVal map As Integer, ByVal sndData As BinaryWriter)
On Error GoTo OnException

    Dim LoopC As Long
    Dim tempIndex As Integer
    
    If Not MapaValido(map) Then Exit Sub

    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
        If UserList(tempIndex).ConnIDValida Then
            Call UserList(tempIndex).Connection.Write(sndData)
        End If
    Next LoopC

OnException:
    
    Call sndData.Clear
    
End Sub

Public Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal sndData As BinaryWriter)
On Error GoTo OnException

    Dim LoopC As Long
    Dim map As Integer
    Dim tempIndex As Integer
    
    map = UserList(UserIndex).Pos.map
    
    If Not MapaValido(map) Then Exit Sub

    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
        If tempIndex <> UserIndex And UserList(tempIndex).ConnIDValida Then
            Call UserList(tempIndex).Connection.Write(sndData)
        End If
    Next LoopC

OnException:
    
    Call sndData.Clear
    
End Sub

