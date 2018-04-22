Attribute VB_Name = "mMap"
Option Explicit

Public MapUp As Boolean
Public HighlightedSystem As Integer
Public ScaleFactor As Integer
Public Const MaxScaleFactor As Integer = 12
Public Const MinScaleFactor As Integer = 4

Private hyperoffsetX As Integer
Private hyperoffsetY As Integer

Private SysMapRadius As Single
Private YouMapRadius As Single

Sub DoDraw()
 
   If ScaleFactor < MinScaleFactor Then ScaleFactor = MinScaleFactor
   If ScaleFactor > MaxScaleFactor Then ScaleFactor = MaxScaleFactor
   
   SysMapRadius = (8 * MinScaleFactor) / ScaleFactor
   YouMapRadius = (4 * MinScaleFactor) / ScaleFactor
   
   HighlightedSystem = -1
   
   If MapUp Then Draw

End Sub

Private Sub Draw()
Dim RefSys As Integer
   
   EnableBlendColour
   
   DrawTexture ViewImages(0), srcRECTNorm, NewfRECT(MapDims.y, MapDims.y + MapDims.Height, MapDims.x, MapDims.x + MapDims.Width), 0, , Blue - &H80000000
   DrawXYWH ViewImages(0), MapDims, Green - &H80000000
   
   For RefSys = 0 To UBound(Systems)
      DrawSysToMap RefSys, SysMapRadius
   Next RefSys
   
   DrawText "Zoom: " & ScaleFactor, MapDims.x + 20, MapDims.y + 62, 14, 14, &HFFB0B0FF
   
   DrawYou

End Sub

Sub DrawSysToMap(ByVal pRefSys As Integer, ByVal pSize As Integer)
Dim MapLoc As Point

   If Ships(You.Ship).InHyperspace Then
      hyperoffsetX = PolToX(Ships(You.Ship).InitialHyperspaceCruiseDistanceLeft - Ships(You.Ship).HyperspaceCruiseDistanceLeft, Ships(You.Ship).maBearing)
      hyperoffsetY = PolToY(Ships(You.Ship).InitialHyperspaceCruiseDistanceLeft - Ships(You.Ship).HyperspaceCruiseDistanceLeft, Ships(You.Ship).maBearing)
   Else
      hyperoffsetX = 0
      hyperoffsetY = 0
   End If
   
   With Systems(pRefSys)
   
      MapLoc.x = (.x - Systems(Ships(You.Ship).System).x - hyperoffsetX) / ScaleFactor + MapDims.x + MapDims.Width / 2
      MapLoc.y = (.y - Systems(Ships(You.Ship).System).y + hyperoffsetY) / ScaleFactor + MapDims.y + MapDims.Height / 2
      
      EnableBlendOne ' blends the circle
   
      If MapLoc.x > MapDims.x And MapLoc.x < MapDims.x + MapDims.Width And _
         MapLoc.x > MapDims.y And MapLoc.y < MapDims.y + MapDims.Height Then
         '---------------------
         If Sqr((mCursor.x - MapLoc.x) ^ 2 + (mCursor.y - MapLoc.y) ^ 2) <= pSize And pRefSys <> Ships(You.Ship).System Then
            HighlightedSystem = pRefSys
            If .Government = -1 Then
               DrawCircle MapLoc.x, MapLoc.y, pSize, White, 32
               DrawCircle MapLoc.x, MapLoc.y, pSize - 1, White, 32
            Else
               DrawCircle MapLoc.x, MapLoc.y, pSize, RelationColour(GovRelations(Ships(You.Ship).Government, .Government)), 32
               DrawCircle MapLoc.x, MapLoc.y, pSize - 1, RelationColour(GovRelations(Ships(You.Ship).Government, .Government)), 32
            End If
            DrawCircle MapLoc.x, MapLoc.y, pSize / 2, &HFFFFFFFF, 18
            DrawText .Name, mWinDims.MapDims.x + 20, mWinDims.MapDims.y + 30, 14, 14, &HFFB0B0FF
            DrawText Int(Sqr((.x - Systems(Ships(You.Ship).System).x) ^ 2 + (MapLoc.y - Systems(Ships(You.Ship).System).y) ^ 2)) & " ly", mWinDims.MapDims.x + 20, mWinDims.MapDims.y + 46, 14, 14, &HFFB0B0FF
         Else
            If .Government = -1 Then
               DrawCircle MapLoc.x, MapLoc.y, pSize, White, 32
               DrawCircle MapLoc.x, MapLoc.y, pSize - 1, White, 32
            Else
               DrawCircle MapLoc.x, MapLoc.y, pSize, RelationColour(GovRelations(Ships(You.Ship).Government, .Government)), 32
               DrawCircle MapLoc.x, MapLoc.y, pSize - 1, RelationColour(GovRelations(Ships(You.Ship).Government, .Government)), 32
            End If
         End If
      End If
   
   End With

End Sub

Sub DrawYou()

   DrawCircle mWinDims.MapDims.x + mWinDims.MapDims.Width / 2, mWinDims.MapDims.y + mWinDims.MapDims.Height / 2, YouMapRadius, &HFF60FF60, 18

End Sub
, True
   DrawXYWH ViewImages(0), MapRightBarDims, Black, DarkGrey, CT_Rect_LeftRight, False, True

End Sub

Sub DrawYou()

   DrawBasicCircle mWinDims.MapDims.X + mWinDims.MapDims.Width / 2, mWinDims.MapDims.Y + mWinDims.MapDims.Height / 2, YouMapRadius, &HFF60FF60, 18

End Sub
