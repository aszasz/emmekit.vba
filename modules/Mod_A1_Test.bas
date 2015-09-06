Attribute VB_Name = "Mod_A1_Test"
Sub UnitTestAngle()
Load_Network_Parameters
ResetNetWork
Dim P1 As Long, P2 As Long
P1 = Add_Point(10, 10)
P2 = Add_Point(5, 10)
p3 = Add_Point(5, 5)
Debug.Assert Angle(point(P1), point(P2)) = 180
Debug.Assert Angle(point(p3), point(P1)) = 45
Debug.Assert Angle(point(P1), point(p3)) = 225
End Sub

Sub UnitTestGetCentralAngle()
Load_Network_Parameters
ResetNetWork
Dim P1 As Long, P2 As Long, p3 As Long, p4 As Long
P1 = Add_Point(10, 10)
P2 = Add_Point(5, 10)
p3 = Add_Point(10, 5)
p4 = Add_Point(0, 0)
Dim Dummy As Link_type
Dummy.op = P1: Dummy.dp = P2
L1 = Add_Link(Dummy)
Dummy.op = P1: Dummy.dp = p3
L2 = Add_Link(Dummy)
Dummy.op = p4: Dummy.dp = P1
L3 = Add_Link(Dummy)

Debug.Assert GetAngletoSplitPoint(P1) = 45

p4 = Add_Point(0, 0)
Dummy.op = p4: Dummy.dp = P2
L3 = Add_Link(Dummy)
Dummy.op = p4: Dummy.dp = p3
L4 = Add_Link(Dummy)
Debug.Assert GetAngletoSplitPoint(p4) = 225



End Sub
Sub UnitTestInsertPosition()
Load_Network_Parameters
ResetNetWork

Dim P1 As Long, P2 As Point_type
P1 = Add_Point(100, 1000)
P2 = PointInsertPosition(point(P1), 60, 1, 1, 25 * approx_one_meter_in_network)
Debug.Assert P2.x = 112.5
Debug.Print P2.x
End Sub



