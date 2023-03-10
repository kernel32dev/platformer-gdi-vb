VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EntityBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim e() As EntityCls
Dim Max As Long
Dim X As Long

Public Function IsEntityAlive(index As Long) As Boolean
On Error Resume Next
IsEntityAlive = Not e(index) Is Nothing
End Function

Function TickContact(PosX As Single, PosY As Single, Width As Single, Height As Single, Optional IsPlayer As Boolean = True, Optional Exclude As Long = -1) As Boolean
For X = 0 To Max
    If Not (e(X) Is Nothing Or Exclude = X) Then
        With e(X)
            If IsPlayer Then
            .WasInContactWithPlayer = .IsInContactWithPlayer
            .IsInContactWithPlayer = (PosX < .PosX + .Width And PosY < .PosY + .Height And .PosX < PosX + Width And .PosY < PosY + Height)
            TickContact = TickContact Or .IsInContactWithPlayer
            Else
            .IsInContactWithObj = (PosX < .PosX + .Width And PosY < .PosY + .Height And .PosX < PosX + Width And .PosY < PosY + Height)
            TickContact = TickContact Or .IsInContactWithObj
            End If
        End With
    End If
Next
End Function

Sub CreateFromMarkerBag(Markers As MarkerBag, Optional Grid As GridCls)
Max = -1
Erase e
For X = 0 To Markers.MarkerCount
    With Markers.GetMarker(X)
        If .T > 0 Then
            Max = Max + 1
            ReDim Preserve e(Max)
            Set e(Max) = New EntityCls
            e(Max).EntityType = .T
            'If .T = Gallomba Then
            '.T = .T
            'End If
            e(Max).Width = EPs(.T).EntityW
            e(Max).Height = EPs(.T).EntityH
            e(Max).PosX = .X
            e(Max).PosY = .Y + e(Max).Height - Int(e(Max).Height)
            Set e(Max).Grid = Grid
        Else
        .T = .T
        End If
    End With
Next
End Sub

Property Get EntityCount() As Long
EntityCount = Max
End Property

Function AddEntity(Optional Grid As GridCls) As EntityCls
For X = 0 To Max
    If e(X) Is Nothing Then
    Set e(X) = New EntityCls
    Set e(X).Grid = Grid
    Set AddEntity = e(X)
    End If
Next
Max = Max + 1
ReDim Preserve e(Max)
Set e(Max) = New EntityCls
Set e(Max).Grid = Grid
Set AddEntity = e(Max)
End Function

Sub TickEntities()
For X = 0 To Max
    If Not e(X) Is Nothing Then
        If e(X).DeadFlag Then
        Set e(X) = Nothing
        Else
        e(X).Tick
        End If
    End If
Next
End Sub

Function GetEntity(index As Long) As EntityCls
On Error Resume Next
Set GetEntity = e(index)
End Function

Sub KillEntity(index As Long)
On Error GoTo Err
Set e(index) = Nothing
Exit Sub
If Max = -1 Then Exit Sub
If e(Max) Is Nothing Then ' TrimArray
    For X = Max To 0 Step -1
        If Not e(X) Is Nothing Then
            If X <> Max Then
            Max = X
            ReDim Preserve e(Max)
            End If
        Exit Sub
        End If
    Next
    Max = -1
    Erase e
End If
Err:
End Sub

Function FindEntityIn(PosX As Single, PosY As Single, Width As Single, Height As Single) As Long
For X = 0 To Max
    If Not e(X) Is Nothing Then
        With e(X)
            If PosX < .PosX + .Width And PosY < .PosY + .Height And .PosX < PosX + Width And .PosY < PosY + Height Then
            FindEntityIn = X
            Exit Function
            End If
        End With
    End If
Next
FindEntityIn = -1
End Function

Sub RenderEntities(Buff As GDIBuffer, OffX As Long, OffY As Long, GameTick As Long)
For X = 0 To Max
If Not e(X) Is Nothing Then e(X).RenderEntity Buff, OffX, OffY, GameTick
Next
End Sub

Private Sub Class_Initialize()
Max = -1
End Sub
