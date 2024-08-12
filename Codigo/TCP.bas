Attribute VB_Name = "TCP"
'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Sub DarCuerpoYCabeza(ByVal UserIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 14/03/2007
'Elije una cabeza para el usuario y le da un body
'*************************************************
Dim NewBody As Integer
Dim NewHead As Integer
Dim UserRaza As Byte
Dim UserGenero As Byte
UserGenero = UserList(UserIndex).genero
UserRaza = UserList(UserIndex).raza
Select Case UserGenero
   Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Humano
                NewHead = RandomNumber(1, 40)
                NewBody = 1
            Case eRaza.Elfo
                NewHead = RandomNumber(101, 112)
                NewBody = 2
            Case eRaza.Drow
                NewHead = RandomNumber(200, 210)
                NewBody = 3
            Case eRaza.Enano
                NewHead = RandomNumber(300, 306)
                NewBody = 300
            Case eRaza.Gnomo
                NewHead = RandomNumber(401, 406)
                NewBody = 300
        End Select
   Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                NewHead = RandomNumber(70, 79)
                NewBody = 1
            Case eRaza.Elfo
                NewHead = RandomNumber(170, 178)
                NewBody = 2
            Case eRaza.Drow
                NewHead = RandomNumber(270, 278)
                NewBody = 3
            Case eRaza.Gnomo
                NewHead = RandomNumber(370, 372)
                NewBody = 300
            Case eRaza.Enano
                NewHead = RandomNumber(470, 476)
                NewBody = 300
        End Select
End Select
UserList(UserIndex).Char.Head = NewHead
UserList(UserIndex).Char.body = NewBody
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean

Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then
        Exit Function
        If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100
    End If
Next LoopC

ValidateSkills = True
    
End Function

Sub ConnectNewUser(ByVal UserIndex As Integer, ByRef name As String, ByRef Password As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass, _
                    ByRef skills() As Byte, ByRef UserEmail As String, ByVal Hogar As eCiudad)
'*************************************************
'Author: Unknown
'Last modified: 20/4/2007
'Conecta un nuevo Usuario
'23/01/2007 Pablo (ToxicWaste) - Agregué ResetFaccion al crear usuario
'24/01/2007 Pablo (ToxicWaste) - Agregué el nuevo mana inicial de los magos.
'12/02/2007 Pablo (ToxicWaste) - Puse + 1 de const al Elfo normal.
'20/04/2007 Pablo (ToxicWaste) - Puse -1 de fuerza al Elfo.
'09/01/2008 Pablo (ToxicWaste) - Ahora los modificadores de Raza se controlan desde Balance.dat
'*************************************************

If Not AsciiValidos(name) Or LenB(name) = 0 Then
    Call WriteErrorMsg(UserIndex, "Nombre invalido.")
    Exit Sub
End If

If UserList(UserIndex).flags.UserLogged Then
    Call LogCheating("El usuario " & UserList(UserIndex).name & " ha intentado crear a " & name & " desde la IP " & UserList(UserIndex).ip)
    
    'Kick player ( and leave character inside :D )!
    Call CloseSocketSL(UserIndex)
    Call Cerrar_Usuario(UserIndex)
    
    Exit Sub
End If

Dim LoopC As Long
Dim totalskpts As Long

'¿Existe el personaje?
If FileExist(CharPath & UCase$(name) & ".chr", vbNormal) = True Then
    Call WriteErrorMsg(UserIndex, "Ya existe el personaje.")
    Exit Sub
End If

'Tiró los dados antes de llegar acá??
If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
    Call WriteErrorMsg(UserIndex, "Debe tirar los dados antes de poder crear un personaje.")
    Exit Sub
End If

UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).flags.Escondido = 0



UserList(UserIndex).Reputacion.AsesinoRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.BurguesRep = 0
UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.NobleRep = 1000
UserList(UserIndex).Reputacion.PlebeRep = 30

UserList(UserIndex).Reputacion.Promedio = 30 / 6


UserList(UserIndex).name = name
UserList(UserIndex).clase = UserClase
UserList(UserIndex).raza = UserRaza
UserList(UserIndex).genero = UserSexo
UserList(UserIndex).email = UserEmail
UserList(UserIndex).Hogar = Hogar

'[Pablo (Toxic Waste) 9/01/08]
UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + ModRaza(UserRaza).Fuerza
UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + ModRaza(UserRaza).Agilidad
UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) + ModRaza(UserRaza).Inteligencia
UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) + ModRaza(UserRaza).Carisma
UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) + ModRaza(UserRaza).Constitucion
'[/Pablo (Toxic Waste)]

For LoopC = 1 To NUMSKILLS
    UserList(UserIndex).Stats.UserSkills(LoopC) = skills(LoopC - 1)
    totalskpts = totalskpts + Abs(UserList(UserIndex).Stats.UserSkills(LoopC))
Next LoopC


If totalskpts > 10 Then
    Call LogHackAttemp(UserList(UserIndex).name & " intento hackear los skills.")
    Call BorrarUsuario(UserList(UserIndex).name)
    Call CloseSocket(UserIndex)
    Exit Sub
End If
'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%

UserList(UserIndex).Char.heading = eHeading.SOUTH

Call DarCuerpoYCabeza(UserIndex)
UserList(UserIndex).OrigChar = UserList(UserIndex).Char
   
 
UserList(UserIndex).Char.WeaponAnim = NingunArma
UserList(UserIndex).Char.ShieldAnim = NingunEscudo
UserList(UserIndex).Char.CascoAnim = NingunCasco

Dim MiInt As Long
MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 3)

UserList(UserIndex).Stats.MaxHP = 15 + MiInt
UserList(UserIndex).Stats.MinHP = 15 + MiInt

MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) \ 6)
If MiInt = 1 Then MiInt = 2

UserList(UserIndex).Stats.MaxSta = 20 * MiInt
UserList(UserIndex).Stats.MinSta = 20 * MiInt


UserList(UserIndex).Stats.MaxAGU = 100
UserList(UserIndex).Stats.MinAGU = 100

UserList(UserIndex).Stats.MaxHam = 100
UserList(UserIndex).Stats.MinHam = 100


'<-----------------MANA----------------------->
If UserClase = eClass.Mage Then 'Cambio en mana inicial (ToxicWaste)
    MiInt = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) * 3
    UserList(UserIndex).Stats.MaxMAN = MiInt
    UserList(UserIndex).Stats.MinMAN = MiInt
ElseIf UserClase = eClass.Cleric Or UserClase = eClass.Druid _
    Or UserClase = eClass.Bard Or UserClase = eClass.Assasin Then
        UserList(UserIndex).Stats.MaxMAN = 50
        UserList(UserIndex).Stats.MinMAN = 50
ElseIf UserClase = eClass.Bandit Then 'Mana Inicial del Bandido (ToxicWaste)
        UserList(UserIndex).Stats.MaxMAN = 150
        UserList(UserIndex).Stats.MinMAN = 150
Else
    UserList(UserIndex).Stats.MaxMAN = 0
    UserList(UserIndex).Stats.MinMAN = 0
End If

If UserClase = eClass.Mage Or UserClase = eClass.Cleric Or _
   UserClase = eClass.Druid Or UserClase = eClass.Bard Or _
   UserClase = eClass.Assasin Then
        UserList(UserIndex).Stats.UserHechizos(1) = 2
End If

UserList(UserIndex).Stats.MaxHIT = 2
UserList(UserIndex).Stats.MinHIT = 1

UserList(UserIndex).Stats.GLD = 0

UserList(UserIndex).Stats.Exp = 0
UserList(UserIndex).Stats.ELU = 300
UserList(UserIndex).Stats.ELV = 1

'???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
UserList(UserIndex).Invent.NroItems = 4

UserList(UserIndex).Invent.Object(1).ObjIndex = 467
UserList(UserIndex).Invent.Object(1).amount = 100

UserList(UserIndex).Invent.Object(2).ObjIndex = 468
UserList(UserIndex).Invent.Object(2).amount = 100

UserList(UserIndex).Invent.Object(3).ObjIndex = 460
UserList(UserIndex).Invent.Object(3).amount = 1
UserList(UserIndex).Invent.Object(3).Equipped = 1

Select Case UserRaza
    Case eRaza.Humano
        UserList(UserIndex).Invent.Object(4).ObjIndex = 463
    Case eRaza.Elfo
        UserList(UserIndex).Invent.Object(4).ObjIndex = 464
    Case eRaza.Drow
        UserList(UserIndex).Invent.Object(4).ObjIndex = 465
    Case eRaza.Enano
        UserList(UserIndex).Invent.Object(4).ObjIndex = 466
    Case eRaza.Gnomo
        UserList(UserIndex).Invent.Object(4).ObjIndex = 466
End Select

UserList(UserIndex).Invent.Object(4).amount = 1
UserList(UserIndex).Invent.Object(4).Equipped = 1

UserList(UserIndex).Invent.ArmourEqpSlot = 4
UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(4).ObjIndex

UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(3).ObjIndex
UserList(UserIndex).Invent.WeaponEqpSlot = 3
 
#If ConUpTime Then
    UserList(UserIndex).LogOnTime = Now
    UserList(UserIndex).UpTime = 0
#End If

'Valores Default de facciones al Activar nuevo usuario
Call ResetFacciones(UserIndex)

Call WriteVar(CharPath & UCase$(name) & ".chr", "INIT", "Password", Password) 'grabamos el password aqui afuera, para no mantenerlo cargado en memoria

Call SaveUser(UserIndex, CharPath & UCase$(name) & ".chr")
  
'Open User
Call ConnectUser(UserIndex, name, Password)
  
End Sub

Sub CloseSocket(ByVal UserIndex As Integer)

On Error GoTo Errhandler
    
    If UserIndex = LastUser Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If
    
    'Call SecurityIp.IpRestarConexion(GetLongIp(UserList(UserIndex).ip))
    
    If UserList(UserIndex).ConnID <> -1 Then
        Call CloseSocketSL(UserIndex)
    End If

    'mato los comercios seguros
    If UserList(UserIndex).ComUsu.DestUsu > 0 Then
        If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
            If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call WriteConsoleMsg(UserList(UserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)

            End If
        End If
    End If

    If UserList(UserIndex).flags.UserLogged Then
        If NumUsers > 0 Then NumUsers = NumUsers - 1
        Call CloseUser(UserIndex)
    Else
        Call ResetUserSlot(UserIndex)
    End If
    
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False

Exit Sub

Errhandler:
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False

    Call ResetUserSlot(UserIndex)

    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.Description & " - UserIndex = " & UserIndex)
End Sub

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal UserIndex As Integer)

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    Call UserList(UserIndex).Connection.Close(False)
    
    UserList(UserIndex).ConnIDValida = False
End If

End Sub


Function EstaPCarea(index As Integer, Index2 As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = UserList(index).Pos.Y - MinYBorder + 1 To UserList(index).Pos.Y + MinYBorder - 1
        For X = UserList(index).Pos.X - MinXBorder + 1 To UserList(index).Pos.X + MinXBorder - 1

            If MapData(UserList(index).Pos.map, X, Y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next X
Next Y
EstaPCarea = False
End Function

Function HayPCarea(Pos As WorldPos) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.map, X, Y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next X
Next Y
HayPCarea = False
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If MapData(Pos.map, X, Y).ObjInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next X
Next Y
HayOBJarea = False
End Function
Function ValidateChr(ByVal UserIndex As Integer) As Boolean

ValidateChr = UserList(UserIndex).Char.Head <> 0 _
                And UserList(UserIndex).Char.body <> 0 _
                And ValidateSkills(UserIndex)

End Function

Sub ConnectUser(ByVal UserIndex As Integer, ByRef name As String, ByRef Password As String)
Dim N As Integer
Dim tStr As String

If UserList(UserIndex).flags.UserLogged Then
    Call LogCheating("El usuario " & UserList(UserIndex).name & " ha intentado loguear a " & name & " desde la IP " & UserList(UserIndex).ip)
    
    'Kick player ( and leave character inside :D )!
    Call CloseSocketSL(UserIndex)
    Call Cerrar_Usuario(UserIndex)
    
    Exit Sub
End If

'Reseteamos los FLAGS
UserList(UserIndex).flags.Escondido = 0
UserList(UserIndex).flags.TargetNPC = 0
UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
UserList(UserIndex).flags.TargetObj = 0
UserList(UserIndex).flags.TargetUser = 0
UserList(UserIndex).Char.FX = 0

'Controlamos no pasar el maximo de usuarios
If NumUsers >= MaxUsers Then
    Call WriteErrorMsg(UserIndex, "El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
    'TODO Call FlushBuffer(UserIndex)
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'¿Este IP ya esta conectado?
If AllowMultiLogins = 0 Then
    If CheckForSameIP(UserIndex, UserList(UserIndex).ip) = True Then
        Call WriteErrorMsg(UserIndex, "No es posible usar mas de un personaje al mismo tiempo.")
        'TODO Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
End If

'¿Existe el personaje?
If Not FileExist(CharPath & UCase$(name) & ".chr", vbNormal) Then
    Call WriteErrorMsg(UserIndex, "El personaje no existe.")
    'TODO Call FlushBuffer(UserIndex)
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'¿Es el passwd valido?
If UCase$(Password) <> UCase$(GetVar(CharPath & UCase$(name) & ".chr", "INIT", "Password")) Then
    Call WriteErrorMsg(UserIndex, "Password incorrecto.")
    'TODO Call FlushBuffer(UserIndex)
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'¿Ya esta conectado el personaje?
If CheckForSameName(name) Then
    If UserList(NameIndex(name)).Counters.Saliendo Then
        Call WriteErrorMsg(UserIndex, "El usuario está saliendo.")
    Else
        Call WriteErrorMsg(UserIndex, "Perdon, un usuario con el mismo nombre se há logoeado.")
    End If
    'TODO Call FlushBuffer(UserIndex)
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'Reseteamos los privilegios
UserList(UserIndex).flags.Privilegios = 0

'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
If EsAdmin(name) Then
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.Admin
    Call LogGM(UserList(UserIndex).name, "Se conecto con ip:" & UserList(UserIndex).ip)
ElseIf EsDios(name) Then
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.Dios
    Call LogGM(UserList(UserIndex).name, "Se conecto con ip:" & UserList(UserIndex).ip)
ElseIf EsSemiDios(name) Then
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.SemiDios
    Call LogGM(UserList(UserIndex).name, "Se conecto con ip:" & UserList(UserIndex).ip)
ElseIf EsConsejero(name) Then
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.Consejero
    Call LogGM(UserList(UserIndex).name, "Se conecto con ip:" & UserList(UserIndex).ip)
Else
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.User
    UserList(UserIndex).flags.AdminPerseguible = True
End If

'Add RM flag if needed
If EsRolesMaster(name) Then
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.RoleMaster
End If

If ServerSoloGMs > 0 Then
    If (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
        Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
        'TODO Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
End If

'Cargamos el personaje
Dim Leer As New clsIniReader

Call Leer.Initialize(CharPath & UCase$(name) & ".chr")

'Cargamos los datos del personaje
Call LoadUserInit(UserIndex, Leer)

Call LoadUserStats(UserIndex, Leer)

If Not ValidateChr(UserIndex) Then
    Call WriteErrorMsg(UserIndex, "Error en el personaje.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

Call LoadUserReputacion(UserIndex, Leer)

Set Leer = Nothing

If UserList(UserIndex).Invent.EscudoEqpSlot = 0 Then UserList(UserIndex).Char.ShieldAnim = NingunEscudo
If UserList(UserIndex).Invent.CascoEqpSlot = 0 Then UserList(UserIndex).Char.CascoAnim = NingunCasco
If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then UserList(UserIndex).Char.WeaponAnim = NingunArma

If (UserList(UserIndex).flags.Muerto = 0) Then
    UserList(UserIndex).flags.SeguroResu = False
    Call WriteResuscitationSafeOff(UserIndex)
Else
    UserList(UserIndex).flags.SeguroResu = True
    Call WriteResuscitationSafeOn(UserIndex)
End If

Call UpdateUserInv(True, UserIndex, 0)
Call UpdateUserHechizos(True, UserIndex, 0)

If UserList(UserIndex).flags.Paralizado Then
    Call WriteParalizeOK(UserIndex)
End If

'Posicion de comienzo
If UserList(UserIndex).Pos.map = 0 Then
    Select Case UserList(UserIndex).Hogar
        Case eCiudad.cNix
            UserList(UserIndex).Pos = Nix
        Case eCiudad.cUllathorpe
            UserList(UserIndex).Pos = Ullathorpe
        Case eCiudad.cBanderbill
            UserList(UserIndex).Pos = Banderbill
        Case eCiudad.cLindos
            UserList(UserIndex).Pos = Lindos
        Case eCiudad.cArghal
            UserList(UserIndex).Pos = Arghal
        Case Else
            UserList(UserIndex).Hogar = eCiudad.cUllathorpe
            UserList(UserIndex).Pos = Ullathorpe
    End Select
Else
    If Not MapaValido(UserList(UserIndex).Pos.map) Then
        Call WriteErrorMsg(UserIndex, "EL PJ se encuenta en un mapa invalido.")
        'TODO Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
End If

'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)
If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex <> 0 Or MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).NpcIndex <> 0 Then
    Dim FoundPlace As Boolean
    Dim tX As Long
    Dim tY As Long
    
    FoundPlace = False
    
    For tY = UserList(UserIndex).Pos.Y - 1 To UserList(UserIndex).Pos.Y + 1
        For tX = UserList(UserIndex).Pos.X - 1 To UserList(UserIndex).Pos.X + 1
            'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
            If LegalPos(UserList(UserIndex).Pos.map, tX, tY, False, True) Then
                FoundPlace = True
                Exit For
            End If
        Next tX
        
        If FoundPlace Then _
            Exit For
    Next tY
    
    If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
        UserList(UserIndex).Pos.X = tX
        UserList(UserIndex).Pos.Y = tY
    Else
        'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
        If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex <> 0 Then
            'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
            If UserList(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex).ComUsu.DestUsu > 0 Then
                'Le avisamos al que estaba comerciando que se tuvo que ir.
                If UserList(UserList(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    Call FinComerciarUsu(UserList(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex).ComUsu.DestUsu)
                    Call WriteConsoleMsg(UserList(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                End If
                'Lo sacamos.
                If UserList(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex).flags.UserLogged Then
                    Call FinComerciarUsu(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex)
                    Call WriteErrorMsg(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                End If
            End If
            
            Call CloseSocket(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex)
        End If
    End If
End If

If UserList(UserIndex).flags.Muerto = 1 Then
    Call Empollando(UserIndex)
End If

'Nombre de sistema
UserList(UserIndex).name = name

UserList(UserIndex).showName = True 'Por default los nombres son visibles

'If in the water, and has a boat, equip it!
If UserList(UserIndex).Invent.BarcoObjIndex > 0 And _
        (HayAgua(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Or BodyIsBoat(UserList(UserIndex).Char.body)) Then
    Dim Barco As ObjData
    Barco = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
    UserList(UserIndex).Char.Head = 0
    If UserList(UserIndex).flags.Muerto = 0 Then

        If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
            UserList(UserIndex).Char.body = iFragataReal
        ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
            UserList(UserIndex).Char.body = iFragataCaos
        Else
            If criminal(UserIndex) Then
                If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.body = iBarcaPk
                If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.body = iGaleraPk
                If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.body = iGaleonPk
            Else
                If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.body = iBarcaCiuda
                If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.body = iGaleraCiuda
                If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.body = iGaleonCiuda
            End If
        End If
    Else
        UserList(UserIndex).Char.body = iFragataFantasmal
    End If
    
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
    UserList(UserIndex).flags.Navegando = 1
End If


'Info
Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
Call WriteChangeMap(UserIndex, UserList(UserIndex).Pos.map, MapInfo(UserList(UserIndex).Pos.map).MapVersion) 'Carga el mapa
Call WritePlayMidi(UserIndex, val(ReadField(1, MapInfo(UserList(UserIndex).Pos.map).Music, 45)))

If UserList(UserIndex).flags.Privilegios <> PlayerType.User And UserList(UserIndex).flags.Privilegios <> (PlayerType.User Or PlayerType.ChaosCouncil) And UserList(UserIndex).flags.Privilegios <> (PlayerType.User Or PlayerType.RoyalCouncil) Then
    UserList(UserIndex).flags.ChatColor = RGB(0, 255, 0)
ElseIf UserList(UserIndex).flags.Privilegios = (PlayerType.User Or PlayerType.RoyalCouncil) Then
    UserList(UserIndex).flags.ChatColor = RGB(0, 255, 255)
ElseIf UserList(UserIndex).flags.Privilegios = (PlayerType.User Or PlayerType.ChaosCouncil) Then
    UserList(UserIndex).flags.ChatColor = RGB(255, 128, 64)
Else
    UserList(UserIndex).flags.ChatColor = vbWhite
End If


''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
#If ConUpTime Then
    UserList(UserIndex).LogOnTime = Now
#End If

'Crea  el personaje del usuario
Call MakeUserChar(True, UserList(UserIndex).Pos.map, UserIndex, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

Call WriteUserCharIndexInServer(UserIndex)
''[/el oso]

Call WriteUpdateUserStats(UserIndex)

Call WriteUpdateHungerAndThirst(UserIndex)

Call SendMOTD(UserIndex)

If haciendoBK Then
    Call WritePauseToggle(UserIndex)
    Call WriteConsoleMsg(UserIndex, "Servidor> Por favor espera algunos segundos, WorldSave esta ejecutandose.", FontTypeNames.FONTTYPE_SERVER)
End If

If EnPausa Then
    Call WritePauseToggle(UserIndex)
    Call WriteConsoleMsg(UserIndex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
End If

'Actualiza el Num de usuarios
'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
NumUsers = NumUsers + 1
UserList(UserIndex).flags.UserLogged = True

'usado para borrar Pjs
Call WriteVar(CharPath & UserList(UserIndex).name & ".chr", "INIT", "Logged", "1")


MapInfo(UserList(UserIndex).Pos.map).NumUsers = MapInfo(UserList(UserIndex).Pos.map).NumUsers + 1

If UserList(UserIndex).Stats.SkillPts > 0 Then
    Call WriteSendSkills(UserIndex)
    Call WriteLevelUp(UserIndex, UserList(UserIndex).Stats.SkillPts)
End If

If NumUsers > recordusuarios Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultaniamente." & "Hay " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFO))
    recordusuarios = NumUsers
    Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
    
End If

If UserList(UserIndex).NroMascotas > 0 And MapInfo(UserList(UserIndex).Pos.map).Pk Then
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(i) > 0 Then
            UserList(UserIndex).MascotasIndex(i) = SpawnNpc(UserList(UserIndex).MascotasType(i), UserList(UserIndex).Pos, True, True)
            
            If UserList(UserIndex).MascotasIndex(i) > 0 Then
                Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
                Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
            Else
                UserList(UserIndex).MascotasIndex(i) = 0
            End If
        End If
    Next i
End If

If UserList(UserIndex).flags.Navegando = 1 Then
    Call WriteNavigateToggle(UserIndex)
End If

If criminal(UserIndex) Then
    Call WriteSafeModeOff(UserIndex)
    UserList(UserIndex).flags.Seguro = False
Else
    UserList(UserIndex).flags.Seguro = True
    Call WriteSafeModeOn(UserIndex)
End If

If UserList(UserIndex).guildIndex > 0 Then
    'welcome to the show baby...
    If Not modGuilds.m_ConectarMiembroAClan(UserIndex, UserList(UserIndex).guildIndex) Then
        Call WriteConsoleMsg(UserIndex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)
    End If
End If

Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXIDs.FXWARP, 0))

Call WriteLoggedMessage(UserIndex)

Call modGuilds.SendGuildNews(UserIndex)

If Lloviendo Then
    Call WriteRainToggle(UserIndex)
End If

tStr = modGuilds.a_ObtenerRechazoDeChar(UserList(UserIndex).name)

If LenB(tStr) <> 0 Then
    Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
End If

Call MostrarNumUsers

N = FreeFile
Open App.Path & "\logs\numusers.log" For Output As N
Print #N, NumUsers
Close #N

N = FreeFile
'Log
Open App.Path & "\logs\Connect.log" For Append Shared As #N
Print #N, UserList(UserIndex).name & " ha entrado al juego. UserIndex:" & UserIndex & " " & time & " " & Date
Close #N

End Sub

Sub SendMOTD(ByVal UserIndex As Integer)
    Dim j As Long
    
    Call WriteGuildChat(UserIndex, "Mensajes de entrada:")
    For j = 1 To MaxLines
        Call WriteGuildChat(UserIndex, MOTD(j).texto)
    Next j
End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'*************************************************
    With UserList(UserIndex).Faccion
        .ArmadaReal = 0
        .CiudadanosMatados = 0
        .CriminalesMatados = 0
        .FuerzasCaos = 0
        .FechaIngreso = "No ingresó a ninguna Facción"
        .RecibioArmaduraCaos = 0
        .RecibioArmaduraReal = 0
        .RecibioExpInicialCaos = 0
        .RecibioExpInicialReal = 0
        .RecompensasCaos = 0
        .RecompensasReal = 0
        .Reenlistadas = 0
        .NivelIngreso = 0
        .MatadosIngreso = 0
        .NextRecompensa = 0
    End With
End Sub

Sub ResetContadores(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'05/20/2007 Integer - Agregue todas las variables que faltaban.
'*************************************************
    With UserList(UserIndex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .COMCounter = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Pasos = 0
        .Pena = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Trabajando = 0
        .Ocultando = 0
        .bPuedeMeditar = False
        .Lava = 0
        .Mimetismo = 0
        .Saliendo = False
        .Salir = 0
        .TiempoOculto = 0
        .TimerMagiaGolpe = 0
        .TimerGolpeMagia = 0
        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeUsarArco = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0
    End With
End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).Char
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex)
        .name = vbNullString
        .modName = vbNullString
        .desc = vbNullString
        .DescRM = vbNullString
        .Pos.map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .ip = vbNullString
        .clase = 0
        .email = vbNullString
        .genero = 0
        .Hogar = 0
        .raza = 0
        
        .EmpoCont = 0
        .PartyIndex = 0
        .PartySolicitud = 0
        
        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            '.CriminalesMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .SkillPts = 0
            .GLD = 0
            .UserAtributos(1) = 0
            .UserAtributos(2) = 0
            .UserAtributos(3) = 0
            .UserAtributos(4) = 0
            .UserAtributos(5) = 0
            .UserAtributosBackUP(1) = 0
            .UserAtributosBackUP(2) = 0
            .UserAtributosBackUP(3) = 0
            .UserAtributosBackUP(4) = 0
            .UserAtributosBackUP(5) = 0
        End With
        
    End With
End Sub

Sub ResetReputacion(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).Reputacion
        .AsesinoRep = 0
        .BandidoRep = 0
        .BurguesRep = 0
        .LadronesRep = 0
        .NobleRep = 0
        .PlebeRep = 0
        .NobleRep = 0
        .Promedio = 0
    End With
End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)
    If UserList(UserIndex).EscucheClan > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(UserIndex, UserList(UserIndex).EscucheClan)
        UserList(UserIndex).EscucheClan = 0
    End If
    If UserList(UserIndex).guildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(UserIndex, UserList(UserIndex).guildIndex)
    End If
    UserList(UserIndex).guildIndex = 0
End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 06/28/2008
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'06/28/2008 NicoNZ - Agrego el flag Inmovilizado
'*************************************************
    With UserList(UserIndex).flags
        .Comerciando = False
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .StatsChanged = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Descuento = vbNullString
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .ModoCombate = False
        .Vuela = 0
        .Navegando = 0
        .Oculto = 0
        .Envenenado = 0
        .invisible = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .Privilegios = 0
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .EstaEmpo = 0
        .Silenciado = 0
        .AdminPerseguible = False
    End With
End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
    Dim LoopC As Long
    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserPets(ByVal UserIndex As Integer)
    Dim LoopC As Long
    
    UserList(UserIndex).NroMascotas = 0
        
    For LoopC = 1 To MAXMASCOTAS
        UserList(UserIndex).MascotasIndex(LoopC) = 0
        UserList(UserIndex).MascotasType(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
    Dim LoopC As Long
    
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
          UserList(UserIndex).BancoInvent.Object(LoopC).amount = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = 0
    Next LoopC
    
    UserList(UserIndex).BancoInvent.NroItems = 0
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
    With UserList(UserIndex).ComUsu
        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(UserIndex)
        End If
    End With
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)

UserList(UserIndex).ConnIDValida = False
UserList(UserIndex).ConnID = -1

Call LimpiarComercioSeguro(UserIndex)
Call ResetFacciones(UserIndex)
Call ResetContadores(UserIndex)
Call ResetCharInfo(UserIndex)
Call ResetBasicUserInfo(UserIndex)
Call ResetReputacion(UserIndex)
Call ResetGuildInfo(UserIndex)
Call ResetUserFlags(UserIndex)
Call LimpiarInventario(UserIndex)
Call ResetUserSpells(UserIndex)
Call ResetUserPets(UserIndex)
Call ResetUserBanco(UserIndex)
With UserList(UserIndex).ComUsu
    .Acepto = False
    .cant = 0
    .DestNick = vbNullString
    .DestUsu = 0
    .Objeto = 0
End With

End Sub

Sub CloseUser(ByVal UserIndex As Integer)
'Call LogTarea("CloseUser " & UserIndex)
On Error GoTo Errhandler

Dim N As Integer
Dim map As Integer
Dim name As String
Dim i As Integer

Dim aN As Integer

aN = UserList(UserIndex).flags.AtacadoPorNpc
If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = vbNullString
End If
aN = UserList(UserIndex).flags.NPCAtacado
If aN > 0 Then
    If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).name Then
        Npclist(aN).flags.AttackedFirstBy = vbNullString
    End If
End If
UserList(UserIndex).flags.AtacadoPorNpc = 0
UserList(UserIndex).flags.NPCAtacado = 0

map = UserList(UserIndex).Pos.map
name = UCase$(UserList(UserIndex).name)

UserList(UserIndex).Char.FX = 0
UserList(UserIndex).Char.loops = 0
Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, 0))


UserList(UserIndex).flags.UserLogged = False
UserList(UserIndex).Counters.Saliendo = False

'Le devolvemos el body y head originales
If UserList(UserIndex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)

'si esta en party le devolvemos la experiencia
If UserList(UserIndex).PartyIndex > 0 Then Call mdParty.SalirDeParty(UserIndex)

' Grabamos el personaje del usuario
Call SaveUser(UserIndex, CharPath & name & ".chr")

'usado para borrar Pjs
Call WriteVar(CharPath & UserList(UserIndex).name & ".chr", "INIT", "Logged", "0")


'Quitar el dialogo
'If MapInfo(Map).NumUsers > 0 Then
'    Call SendToUserArea(UserIndex, "QDL" & UserList(UserIndex).Char.charindex)
'End If

If MapInfo(map).NumUsers > 0 Then
    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(UserList(UserIndex).Char.CharIndex))
End If



'Borrar el personaje
If UserList(UserIndex).Char.CharIndex > 0 Then
    Call EraseUserChar(UserIndex)
End If

'Borrar mascotas
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).flags.NPCActive Then _
            Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
    End If
Next i

'Update Map Users
MapInfo(map).NumUsers = MapInfo(map).NumUsers - 1

If MapInfo(map).NumUsers < 0 Then
    MapInfo(map).NumUsers = 0
End If

' Si el usuario habia dejado un msg en la gm's queue lo borramos
If Ayuda.Existe(UserList(UserIndex).name) Then Call Ayuda.Quitar(UserList(UserIndex).name)

Call ResetUserSlot(UserIndex)

Call MostrarNumUsers

N = FreeFile(1)
Open App.Path & "\logs\Connect.log" For Append Shared As #N
Print #N, name & " há dejado el juego. " & "User Index:" & UserIndex & " " & Time & " " & Date
Close #N

Exit Sub

Errhandler:
Call LogError("Error en CloseUser. Número " & Err.Number & " Descripción: " & Err.Description)

End Sub

Public Sub EcharPjsNoPrivilegiados()
Dim LoopC As Long

For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
        If UserList(LoopC).flags.Privilegios And PlayerType.User Then
            Call CloseSocket(LoopC)
        End If
    End If
Next LoopC

End Sub

