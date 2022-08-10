Attribute VB_Name = "Test11"
Dim swApp As Object
Dim swModeldoc2 As SldWorks.ModelDoc2
Dim swSelectionManager As SldWorks.SelectionMgr
Dim swPart1 As SldWorks.PartDoc
Dim vswBody1 As Variant

Dim kaka As Integer

Sub main()
Set swApp = Application.SldWorks
Set swModeldoc2 = swApp.ActiveDoc
Set swPart1 = swModeldoc2
vswBody1 = swPart1.GetBodies2(swSolidBody, True)
Debug.Print (vswBody1(0).Name)

Dim outx1 As Double
Dim outy1 As Double
Dim outZ1 As Double

Dim outx2 As Double
Dim outy2 As Double
Dim outZ2 As Double

Dim arrDim(2) As Double
Dim bool As Boolean
bool = vswBody1(0).GetExtremePoint(1, 0, 0, outx1, outy1, outZ1)
bool = vswBody1(0).GetExtremePoint(-1, 0, 0, outx2, outy2, outZ2)
arrDim(0) = outx1 - outx2
bool = vswBody1(0).GetExtremePoint(0, 1, 0, outx1, outy1, outZ1)
bool = vswBody1(0).GetExtremePoint(0, -1, 0, outx2, outy2, outZ2)
arrDim(1) = outy1 - outy2
bool = vswBody1(0).GetExtremePoint(0, 0, 1, outx1, outy1, outZ1)
bool = vswBody1(0).GetExtremePoint(0, 0, -1, outx2, outy2, outZ2)
arrDim(2) = outZ1 - outZ2

'Sort-----------------
Dim i As Integer
Dim temp As Double
bool = True
While (bool)
i = 0
bool = False
    While (i < 1)
    If (arrDim(i) > arrDim(i + 1)) Then
    temp = arrDim(i)
    arrDim(i) = arrDim(i + 1)
    arrDim(1 + i) = temp
    bool = True
    End If
    i = i + 1
    Wend
Wend
'--------------------

Debug.Print (arrDim(0))
Debug.Print (arrDim(1))
Debug.Print (arrDim(2))

Dim docUserUnit As SldWorks.UserUnit
Set docUserUnit = swModeldoc2.GetUserUnit(swLengthUnit)

Dim thick As String
thick = docUserUnit.ConvertToUserUnit(arrDim(0), False, False)
Debug.Print (thick)

Dim width As String
width = docUserUnit.ConvertToUserUnit(arrDim(1), False, False)
Debug.Print (width)

Dim length As String
length = docUserUnit.ConvertToUserUnit(arrDim(2), False, False)
Debug.Print (length)




End Sub







