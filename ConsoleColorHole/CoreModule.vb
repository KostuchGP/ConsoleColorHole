Imports INFITF
Imports PARTITF 'Include Hole
Imports MECMOD 'Include Bodies, Body, PartDocument, Part, Shapes
Imports ProductStructureTypeLib 'Include Product

Module CoreModule
    Public CATIA As Object
    Public mainDoc As INFITF.Document
    Private czyShim As Boolean = False

    'Zestawienie kołków:
    Public arrayOfKolki As Double() = {6, 8, 10}
    Public arrayOfSruby As String() = {"M4", "M5", "M6", "M8", "M10", "M12"}
    Public resztaHole As New List(Of String)

    'Startowy Subroutine
    Public Sub Main()
        Dim iErr As Integer

        On Error Resume Next
        CATIA = GetObject(, "CATIA.Application")
        iErr = Err.Number
        If (iErr <> 0) Then
            MsgBox("There is no open CATIA Application")
            Exit Sub
        End If

        mainDoc = CATIA.ActiveDocument

        If Err.Number <> 0 Then
            MsgBox("There is no open any component in CATIA")
            Exit Sub
        End If
        On Error GoTo 0
        If TypeName(mainDoc) <> "ProductDocument" Then
            MsgBox("In CATIA Active window must be the Assembly (.CATProduct)")
            Exit Sub
        Else ' Jeżeli wszystko działa to wykonuje się to co poniżej
            Select Case MsgBox("Do you want start to color holes?", MsgBoxStyle.YesNo, "Tool to color Holes")
                Case MsgBoxResult.Yes
                    Exit Select
                Case MsgBoxResult.No
                    Exit Sub
            End Select
            arrayWithFeatures()
            'DetectHoles()
        End If
    End Sub
    'Loading All 

    Sub DetectHoles()
        Dim iMatrix(11)
        iMatrix(0) = 1.0
        iMatrix(1) = 0.0
        iMatrix(2) = 0.0
        iMatrix(3) = 0.0
        iMatrix(4) = 1.0
        iMatrix(5) = 0.0
        iMatrix(6) = 0.0
        iMatrix(7) = 0.0
        iMatrix(8) = 1.0
        iMatrix(9) = 100.0
        iMatrix(10) = 100.0
        iMatrix(11) = 100.0

        Dim arrayHoles(1, 3) As String
        Dim oSelection2
        Dim dwaElementy As Integer
        Dim czyGwintowanyOtwor As CatHoleThreadingMode
        Dim dlugoscHole1 As Double
        Dim dlugoscHole2 As Double
        Dim dlugoscSruby As Double
        Dim arrayOfVariantOfBSTR1(0)

        Dim newMatrix(11)

        Dim product1 As Product
        product1 = mainDoc.Product

        Dim products1 As Products
        products1 = product1.Products

        Dim product2 As Product
        Dim products2 As Products

        oSelection2 = mainDoc.Selection
        dwaElementy = oSelection2.Count

        'Ładowanie arrayHoles
        For i = 0 To oSelection2.Count - 1
            czyGwintowanyOtwor = oSelection2.ITEM(i + 1).Value.ThreadingMode
            If czyGwintowanyOtwor = CatHoleThreadingMode.catThreadedHoleThreading Then
                'jeżeli hole jest z gwintem to:
                arrayHoles(i, 0) = oSelection2.Item(i + 1).Value.HoleThreadDescription.Value ' tylko dla thread np M10
            Else
                arrayHoles(i, 0) = oSelection2.Item(i + 1).Value.Diameter.Value 'np 10 
            End If

            'arrayHoles(i, 0) = oSelection2.Item(i + 1).Value.Name 'np Hole.1
            'arrayHoles(i, 1) = oSelection2.Item(i + 1).LeafProduct.PartNumber 'np Konsola
            arrayHoles(i, 1) = oSelection2.Item(i + 1).LeafProduct.Parent.Parent.Name ' np BG_04.1
            arrayHoles(i, 2) = oSelection2.Item(i + 1).Value.BottomLimit.Dimension.Value 'np Hole.1
        Next

        'przypisujemy wartosc
        product2 = products1.Item(arrayHoles(0, 1))
        products2 = product2.Products

        'Liczymy

        Double.TryParse(arrayHoles(0, 2), dlugoscHole1)
        Double.TryParse(arrayHoles(1, 2), dlugoscHole2)

        If czyShim = True Then
            dlugoscSruby = dlugoscHole1 + dlugoscHole2 + 10
            'MsgBox(dlugoscSruby)
        Else
            dlugoscSruby = dlugoscHole1 + dlugoscHole2
            'MsgBox(dlugoscSruby)
        End If

        Dim proba As Product
        proba = products2.Item(3)

        'Dim pozycjaCompasu As DNBASY.AsyMotionTarget

        'GetCompassPosition(iMatrix, DNBASY.AsyMotionTargetDataFormat.AsyMotionTarget3x4Matrix)

        proba.Move.Apply(iMatrix)

        'Wstawiamy elementy
        If dlugoscSruby > 55 Then
            arrayOfVariantOfBSTR1(0) = "E:\Pliki 3D\SrubaM10x55.CATPart"
        ElseIf dlugoscSruby > 50 Then
            arrayOfVariantOfBSTR1(0) = "E:\Pliki 3D\SrubaM10x50.CATPart"
        End If

        products2.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

        Dim sruba As Product
        sruba = products2.Item(products2.Count)

        sruba.Move.Apply(iMatrix)

    End Sub

    Sub arrayWithFeatures()
        Dim arrayPomocne(0, 2) As String
        Dim oSelection
        Dim ileHole As Integer
        Dim visPropertySet As VisPropertySet
        Dim czyGwintowanyOtwor As CatHoleThreadingMode
        Dim pomalowanychOtworow As Integer = 0

        oSelection = mainDoc.Selection
        oSelection.Clear()

        visPropertySet = oSelection.VisProperties

        oSelection.Search("n:*Hole.*,all")

        ileHole = oSelection.Count
        ReDim arrayPomocne(ileHole - 1, 3)

        'Ładowanie arrayPomocne
        For i = 0 To oSelection.Count - 1
            czyGwintowanyOtwor = oSelection.Item(i + 1).Value.ThreadingMode
            arrayPomocne(i, 0) = oSelection.Item(i + 1).Value.Name
            arrayPomocne(i, 1) = oSelection.Item(i + 1).LeafProduct.PartNumber
            If czyGwintowanyOtwor = CatHoleThreadingMode.catThreadedHoleThreading Then
                'jeżeli hole jest z gwintem to:
                arrayPomocne(i, 2) = oSelection.Item(i + 1).Value.HoleThreadDescription.Value ' tylko dla thread np M10
                If arrayOfSruby.Contains(arrayPomocne(i, 2)) Then
                    arrayPomocne(i, 3) = "Gwint"
                    pomalowanychOtworow = pomalowanychOtworow + 1
                Else
                    arrayPomocne(i, 3) = "Z poza zakresu"
                End If
            Else
                arrayPomocne(i, 2) = oSelection.Item(i + 1).Value.Diameter.Value 'np 10 
                If arrayOfKolki.Contains(arrayPomocne(i, 2)) Then
                    arrayPomocne(i, 3) = "Kolek"
                    pomalowanychOtworow = pomalowanychOtworow + 1
                Else
                    arrayPomocne(i, 3) = "Z poza zakresu"
                End If
            End If
        Next

        'Druga Petla
        For InxSel = 0 To ileHole - 1
            oSelection.Clear()

            Dim documents1 As Documents
            documents1 = CATIA.Documents

            Dim partDocument1 As MECMOD.PartDocument
            partDocument1 = documents1.Item(arrayPomocne(InxSel, 1) & ".CATPart")

            Dim part1 As MECMOD.Part
            part1 = partDocument1.Part

            Dim bodies1 As MECMOD.Bodies
            bodies1 = part1.Bodies

            Dim body1 As MECMOD.Body
            body1 = bodies1.Item("PartBody")

            Dim shapes1 As MECMOD.Shapes
            shapes1 = body1.Shapes

            Dim hole1 As Hole
            hole1 = shapes1.Item(arrayPomocne(InxSel, 0))

            oSelection.Add(hole1)
            If arrayPomocne(InxSel, 3) = "Gwint" Then

                oSelection.VisProperties.SetRealColor(0, 0, 0, 0)
            ElseIf arrayPomocne(InxSel, 3) = "Kolek" Then

                oSelection.VisProperties.SetRealColor(0, 255, 0, 0)
            ElseIf arrayPomocne(InxSel, 3) = "Z poza zakresu" Then

                resztaHole.Add(arrayPomocne(InxSel, 2))
            End If
        Next
        oSelection.Clear()

        Dim str As String
        Dim sResult As String = ""
        For Each str In resztaHole
            sResult &= str & Environment.NewLine
        Next

        MsgBox("No colored holes: " & Environment.NewLine & sResult & Environment.NewLine & Environment.NewLine & "Colored: " & pomalowanychOtworow & " elements.")

    End Sub

End Module