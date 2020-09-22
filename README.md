<div align="center">

## Map\_Line\_Intersect


</div>

### Description

Determines the intersection of two line segments
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bob RIchards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bob-richards.md)
**Level**          |Beginner
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bob-richards-map-line-intersect__1-5365/archive/master.zip)





### Source Code

```
Function Map_Line_Intersect(x1 As Long, y1 As Long, x2 As Long, y2 As Long, _
      x3 As Long, y3 As Long, x4 As Long, y4 As Long, _
      ByRef intersect As Boolean, ByRef x As Long, ByRef y As Long) As Boolean
'Call with x1,y1,x2,y2,x3,y3,x4,y4 and returns intersect,x,y
'
'Where:
' x1,y1,x2,y2,x3,y3,x4,y4 are the end points of two line segments
'Returns:
' intersect is true/false, and x,y is the interecting point if intersect is true
'
'Description:
'
'Line intersection test, requires a form with object Picture1
'
'Equations for the lines are:
' Pa = P1 + Ua(P2 - P1)
' Pb = P3 + Ub(P4 - P3)
'
'Solving for the point where Pa = Pb gives the following equations for ua and ub
' Ua = ((x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)) / ((y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1))
' Ub = ((x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)) / ((y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1))
'
'Substituting either of these into the corresponding equation for the line gives the intersection point.
'For example the intersection point (x,y) is
' x = x1 + Ua(x2 - x1)
' y = y1 + Ua(y2 - y1)
'
'Notes:
' - The denominators are the same.
'
' - If the denominator above is 0 then the two lines are parallel.
'
' - If the denominator and numerator are 0 then the two lines are coincident.
'
' - The equations above apply to lines, if the intersection of line segments is required then it is only
'  necessary to test if ua and ub lie between 0 and 1. Whichever one lies within that range then the
'  corresponding line segment contains the intersection point. If both lie within the range of 0 to 1 then
'  the intersection point is within both line segments.
Dim d As Double
Dim Ua As Double
Dim Ub As Double
'Pre calc the denominator, if zero then both lines are parallel and there is no intersection
d = ((y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1))
If d <> 0 Then
  'Solve for the simultaneous equations
  Ua = ((x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)) / d
  Ub = ((x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)) / d
End If
'Could the lines intersect?
If Ua > 0 And Ua < 1 And Ub > 0 And Ub < 1 Then
  'Calculate the intersection point
  x = x1 + Ua * (x2 - x1)
  y = y1 + Ua * (y2 - y1)
  'Yes, they do
  Map_Line_Intersect = True
  intersect = True
Else
  'No, they do not
  Map_Line_Intersect = False
  intersect = False
End If
End Function
```

