Attribute VB_Name = "InvNpc"
'Argentum Online 0.12.2
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
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Inv & Obj
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Modulo para controlar los objetos y los inventarios.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Public Function TirarItemAlPiso(Pos As WorldPos, Obj As Obj, Optional NotPirata As Boolean = True) As WorldPos
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim NuevaPos As WorldPos
    NuevaPos.X = 0
    NuevaPos.Y = 0
    
    Tilelibre Pos, NuevaPos, Obj, NotPirata, True
    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
        Call MakeObj(Obj, Pos.Map, NuevaPos.X, NuevaPos.Y)
    End If
    TirarItemAlPiso = NuevaPos

    Exit Function
Errhandler:

End Function

'AyudandOh
Public Sub NPC_TIRAR_ITEMS(ByVal UserIndex As Integer, ByRef npc As npc, ByVal IsPretoriano As Boolean)

On Error Resume Next

    With npc
        
        Dim i As Byte
        Dim MiObj As Obj
        Dim NroDrop As Integer
        Dim Random As Integer
        Dim ObjIndex As Integer
        
        If IsPretoriano Then
            For i = 1 To MAX_INVENTORY_SLOTS
                If .Invent.Object(i).ObjIndex > 0 Then
                      MiObj.Amount = .Invent.Object(i).Amount
                      MiObj.ObjIndex = .Invent.Object(i).ObjIndex
                      Call TirarItemAlPiso(.Pos, MiObj)
                End If
            Next i

            If .GiveGLD > 0 Then _
                Call TirarOroNpc(UserIndex, .GiveGLD, .Pos)
                
            Exit Sub
        End If
        
        Random = RandomNumber(1, 100)
        
        If Random <= 90 Then
            NroDrop = 1
            
            If Random <= 10 Then
                NroDrop = NroDrop + 1
                 
                For i = 1 To 3

                    If RandomNumber(1, 100) <= 10 Then
                        NroDrop = NroDrop + 1
                    Else
                        Exit For
                    End If
                Next i
                
            End If
            

            ObjIndex = .Drop(NroDrop).ObjIndex
            If ObjIndex > 0 Then
            
                If ObjIndex = iORO Then
                    Call TirarOroNpc(UserIndex, .Drop(NroDrop).Amount, npc.Pos)
                Else
                    MiObj.Amount = .Drop(NroDrop).Amount
                    MiObj.ObjIndex = .Drop(NroDrop).ObjIndex
                    
                    Call TirarItemAlPiso(.Pos, MiObj)
                End If
            End If

        End If

    End With

End Sub

''
' Gets the amount of a certain item that an npc has.
'
' @param npcIndex Specifies reference to npcmerchant
' @param ObjIndex Specifies reference to object
' @return   The amount of the item that the npc has
' @remarks This function reads the Npc.dat file
Function EncontrarCant(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: 03/09/08
'Last Modification By: Marco Vanotti (Marco)
' - 03/09/08 EncontrarCant now returns 0 if the npc doesn't have it (Marco)
'***************************************************
On Error Resume Next
'Devuelve la cantidad original del obj de un npc

    Dim ln As String, npcfile As String
    Dim i As Integer
    
    npcfile = DatPath & "NPCs.dat"
     
    For i = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & i)
        If ObjIndex = val(ReadField(1, ln, 45)) Then
            EncontrarCant = val(ReadField(2, ln, 45))
            Exit Function
        End If
    Next
                       
    EncontrarCant = 0

End Function

Sub ResetNpcInv(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim i As Integer
    
    With Npclist(NpcIndex)
        .Invent.NroItems = 0
        
        For i = 1 To MAX_INVENTORY_SLOTS
           .Invent.Object(i).ObjIndex = 0
           .Invent.Object(i).Amount = 0
        Next i
        
        .InvReSpawn = 0
    End With

End Sub

''
' Removes a certain amount of items from a slot of an npc's inventory
'
' @param npcIndex Specifies reference to npcmerchant
' @param Slot Specifies reference to npc's inventory's slot
' @param antidad Specifies amount of items that will be removed
Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 23/11/2009
'Last Modification By: Marco Vanotti (Marco)
' - 03/09/08 Now this sub checks that te npc has an item before respawning it (Marco)
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************
    Dim ObjIndex As Integer
    Dim iCant As Integer
    
    With Npclist(NpcIndex)
        ObjIndex = .Invent.Object(Slot).ObjIndex
    
        'Quita un Obj
        If ObjData(.Invent.Object(Slot).ObjIndex).Crucial = 0 Then
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - Cantidad
            
            If .Invent.Object(Slot).Amount <= 0 Then
                .Invent.NroItems = .Invent.NroItems - 1
                .Invent.Object(Slot).ObjIndex = 0
                .Invent.Object(Slot).Amount = 0
                If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
                   Call CargarInvent(NpcIndex) 'Reponemos el inventario
                End If
            End If
        Else
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - Cantidad
            
            If .Invent.Object(Slot).Amount <= 0 Then
                .Invent.NroItems = .Invent.NroItems - 1
                .Invent.Object(Slot).ObjIndex = 0
                .Invent.Object(Slot).Amount = 0
                
                    'Check if the item is in the npc's dat.
                    iCant = EncontrarCant(NpcIndex, ObjIndex)
                    If iCant Then
                        .Invent.Object(Slot).ObjIndex = ObjIndex
                        .Invent.Object(Slot).Amount = iCant
                        .Invent.NroItems = .Invent.NroItems + 1
                    End If
                End If
                
                If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
                   Call CargarInvent(NpcIndex) 'Reponemos el inventario
                End If
            End If
    End With
End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    'Vuelve a cargar el inventario del npc NpcIndex
    Dim LoopC As Integer
    Dim ln As String
    Dim npcfile As String
    
    npcfile = DatPath & "NPCs.dat"
    
    With Npclist(NpcIndex)
        .Invent.NroItems = val(GetVar(npcfile, "NPC" & .Numero, "NROITEMS"))
        
        For LoopC = 1 To .Invent.NroItems
            ln = GetVar(npcfile, "NPC" & .Numero, "Obj" & LoopC)
            .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
            .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
            
        Next LoopC
    End With

End Sub


'AyudandOh
Public Sub TirarOroNpc(ByVal UserIndex As Integer, ByVal Cantidad As Long, ByRef Pos As WorldPos)

On Error GoTo Errhandler

    If Cantidad > 0 Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Cantidad
        Call WriteUpdateGold(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has obtenido " & Cantidad & " monedas de oro.", FontTypeNames.FONTTYPE_INFO) 'Ponele una comilla adelante de esta linea si no queres q salga por consola
    End If

    Exit Sub

Errhandler:
    Call LogError("Error en TirarOro. Error " & Err.Number & " : " & Err.description)
End Sub
