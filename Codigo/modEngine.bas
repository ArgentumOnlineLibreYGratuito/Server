Attribute VB_Name = "modEngine"
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

' -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
' [Engine::Network]
' -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

Private NetListener_ As Network_Server
Private NetProtocol_ As Network_Protocol

' -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
' [Engine::Main]
' -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

Public Sub Initialize()

    Dim Configuration As Kernel_Properties
    
    Call Kernel.Initialize(eKernelModeServer, Configuration)

End Sub

Public Sub Tick()

    Call Kernel.Tick
    
End Sub

' -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
' [Engine::Network]
' -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

Public Sub NetListen(ByVal Address As String, ByVal Port As Long)
    
    Set NetProtocol_ = New Network_Protocol
    Call NetProtocol_.Attach(AddressOf Network_OnAttach, AddressOf Network_OnDetach, AddressOf Network_OnRecv, AddressOf Network_OnSend, AddressOf Network_OnError)
    
    Set NetListener_ = Kernel.Network.Listen(Address, Port)
    Call NetListener_.SetProtocol(NetProtocol_)
    
    Call modEngine_Protocol.Initialize
    
End Sub

Public Sub NetClose()
    
    Set NetListener_ = Nothing
    
End Sub

Public Sub NetFlush()
    
    If (Not NetListener_ Is Nothing) Then
    
        Call NetListener_.Flush
        
    End If

End Sub

Private Sub Network_OnAttach(ByVal Connection As Network_Client)
    
    Call modEngine_Protocol.OnConnect(Connection)
    
End Sub

Private Sub Network_OnDetach(ByVal Connection As Network_Client)
    
    Call modEngine_Protocol.OnClose(Connection)
    
End Sub

Private Sub Network_OnRecv(ByVal Connection As Network_Client, ByVal Message As BinaryReader)
    
    Call modEngine_Protocol.Decode(Connection, Message)
    Call modEngine_Protocol.Handle(Connection, Message)
    
End Sub

Private Sub Network_OnSend(ByVal Connection As Network_Client, ByVal Message As BinaryReader)
    
    Call modEngine_Protocol.Encode(Connection, Message)
    
End Sub

Private Sub Network_OnError(ByVal Connection As Network_Client, ByVal Error As Long, ByVal Description As String)

    ' TODO: Log.Error(...)
    
End Sub



