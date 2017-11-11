Imports INFITF
Imports PARTITF 'Include Hole
Imports MECMOD 'Include Bodies, Body, PartDocument, Part, Shapes
Imports ProductStructureTypeLib 'Include Product

Module CoreModule
    Public CATIA As Object
    Public mainDoc As INFITF.Document
    Private pomalowanychOtworow As Integer = 0
    Private ileHole As Integer = 0

    'Zestawienie kołków:
    Public arrayOfKolki As Double() = {6, 8, 10}
    Public arrayOfSruby As String() = {"M4", "M5", "M6", "M8", "M10", "M12", "M16", "M20"}
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
            colorWithHole()
            If ileHole = 0 Then
                MsgBox("Any holes detected")
                Exit Sub
            End If
            colorWithUserPattern()
            colorWithRectPattern()
            Result()
        End If
    End Sub
    'Coloring all holes
    Sub colorWithHole()
        Dim arrayPomocneHole(0, 2) As String
        Dim oSelection
        Dim visPropertySet As VisPropertySet
        Dim czyGwintowanyOtwor As CatHoleThreadingMode
        Dim partDocumentToCheck As MECMOD.PartDocument

        oSelection = mainDoc.Selection
        oSelection.Clear()

        visPropertySet = oSelection.VisProperties

        oSelection.Search("n:*Hole.*,all")

        ileHole = oSelection.Count

        If ileHole = 0 Then Exit Sub

        ReDim arrayPomocneHole(ileHole - 1, 4)

        'Ładowanie arrayPomocne
        For i = 0 To oSelection.Count - 1
            czyGwintowanyOtwor = oSelection.Item(i + 1).Value.ThreadingMode
            arrayPomocneHole(i, 0) = oSelection.Item(i + 1).Value.Name 'Nazwa np Hole.1
            arrayPomocneHole(i, 1) = oSelection.Item(i + 1).Value.Parent.Parent.Parent.Parent.Parent.Name 'Zmieniamy to na bezpośrednią nazwę pliku ]:->
            arrayPomocneHole(i, 4) = oSelection.Item(i + 1).Value.Parent.Parent.Name 'Np PartBody / PartBody.1
            partDocumentToCheck = CATIA.Documents.Item(arrayPomocneHole(i, 1))
            If InStr(1, partDocumentToCheck.Path, "pp") = 1 Then
                arrayPomocneHole(i, 3) = "Z poza zakresu"
                GoTo OverHoleCheckingThreadingMode
            End If

            If czyGwintowanyOtwor = CatHoleThreadingMode.catThreadedHoleThreading Then
                'jeżeli hole jest z gwintem to:
                arrayPomocneHole(i, 2) = oSelection.Item(i + 1).Value.HoleThreadDescription.Value ' tylko dla thread np M10
                If arrayOfSruby.Contains(arrayPomocneHole(i, 2)) Then
                    arrayPomocneHole(i, 3) = "Gwint"
                    pomalowanychOtworow = pomalowanychOtworow + 1
                Else
                    arrayPomocneHole(i, 3) = "Z poza zakresu"
                End If
            Else
                arrayPomocneHole(i, 2) = oSelection.Item(i + 1).Value.Diameter.Value 'np 10 
                If arrayOfKolki.Contains(arrayPomocneHole(i, 2)) Then
                    arrayPomocneHole(i, 3) = "Kolek"
                    pomalowanychOtworow = pomalowanychOtworow + 1
                Else
                    arrayPomocneHole(i, 3) = "Z poza zakresu"
                End If
            End If

OverHoleCheckingThreadingMode:

        Next

        'Druga Petla
        For InxSel = 0 To ileHole - 1
            oSelection.Clear()

            Dim documents1 As Documents
            documents1 = CATIA.Documents

            Dim partDocument1 As MECMOD.PartDocument
            partDocument1 = documents1.Item(arrayPomocneHole(InxSel, 1))

            Dim part1 As MECMOD.Part
            part1 = partDocument1.Part

            Dim bodies1 As MECMOD.Bodies
            bodies1 = part1.Bodies

            Dim body1 As MECMOD.Body
            body1 = bodies1.Item(arrayPomocneHole(InxSel, 4))

            Dim shapes1 As MECMOD.Shapes
            shapes1 = body1.Shapes

            Dim hole1 As Hole
            hole1 = shapes1.Item(arrayPomocneHole(InxSel, 0))

            oSelection.Add(hole1)
            If arrayPomocneHole(InxSel, 3) = "Gwint" Then

                oSelection.VisProperties.SetRealColor(0, 0, 0, 0)
            ElseIf arrayPomocneHole(InxSel, 3) = "Kolek" Then

                oSelection.VisProperties.SetRealColor(0, 255, 0, 0)
            ElseIf arrayPomocneHole(InxSel, 3) = "Z poza zakresu" Then
                'USTAWIAM szary kolor dla innych otworów
                oSelection.VisProperties.SetRealColor(210, 210, 255, 0)
                resztaHole.Add(arrayPomocneHole(InxSel, 2))
            End If
        Next
        oSelection.Clear()
    End Sub
    'Coloring all UserPatterns
    Sub colorWithUserPattern()
        Dim arrayPomocneUserPattern(0, 2) As String
        Dim oSelection
        Dim ileUserPattern As Integer
        Dim visPropertySet As VisPropertySet
        Dim czyGwintowanyOtwor As CatHoleThreadingMode
        Dim partDocumentToCheck As MECMOD.PartDocument

        oSelection = mainDoc.Selection
        oSelection.Clear()

        visPropertySet = oSelection.VisProperties

        oSelection.Search("n:*UserPattern.*,all")

        ileUserPattern = oSelection.Count

        If ileUserPattern = 0 Then Exit Sub

        ReDim arrayPomocneUserPattern(ileUserPattern - 1, 4)

        'Ładowanie arrayPomocne
        For i = 0 To oSelection.Count - 1
            czyGwintowanyOtwor = oSelection.Item(i + 1).Value.ItemToCopy.ThreadingMode
            arrayPomocneUserPattern(i, 0) = oSelection.Item(i + 1).Value.Name
            arrayPomocneUserPattern(i, 1) = oSelection.Item(i + 1).Value.Parent.Parent.Parent.Parent.Parent.Name
            arrayPomocneUserPattern(i, 4) = oSelection.Item(i + 1).Value.ItemToCopy.Parent.Parent.Name 'Nazwa np PartBody
            partDocumentToCheck = CATIA.Documents.Item(arrayPomocneUserPattern(i, 1))
            If InStr(1, partDocumentToCheck.Path, "pp") = 1 Then
                arrayPomocneUserPattern(i, 3) = "Z poza zakresu"
                GoTo OverUserPatternCheckingThreadingMode
            End If

            If czyGwintowanyOtwor = CatHoleThreadingMode.catThreadedHoleThreading Then
                'jeżeli hole jest z gwintem to:
                arrayPomocneUserPattern(i, 2) = oSelection.Item(i + 1).Value.ItemToCopy.HoleThreadDescription.Value ' tylko dla thread np M10
                If arrayOfSruby.Contains(arrayPomocneUserPattern(i, 2)) Then
                    arrayPomocneUserPattern(i, 3) = "Gwint"
                    pomalowanychOtworow = pomalowanychOtworow + 1
                Else
                    arrayPomocneUserPattern(i, 3) = "Z poza zakresu"
                End If
            Else
                arrayPomocneUserPattern(i, 2) = oSelection.Item(i + 1).Value.ItemToCopy.Diameter.Value 'np 10 
                If arrayOfKolki.Contains(arrayPomocneUserPattern(i, 2)) Then
                    arrayPomocneUserPattern(i, 3) = "Kolek"
                    pomalowanychOtworow = pomalowanychOtworow + 1
                Else
                    arrayPomocneUserPattern(i, 3) = "Z poza zakresu"
                End If
            End If

OverUserPatternCheckingThreadingMode:

        Next

        'Druga Petla
        For InxSel = 0 To ileUserPattern - 1
            oSelection.Clear()

            Dim documents1 As Documents
            documents1 = CATIA.Documents

            Dim partDocument1 As MECMOD.PartDocument
            partDocument1 = documents1.Item(arrayPomocneUserPattern(InxSel, 1))

            Dim part1 As MECMOD.Part
            part1 = partDocument1.Part

            Dim bodies1 As MECMOD.Bodies
            bodies1 = part1.Bodies

            Dim body1 As MECMOD.Body
            body1 = bodies1.Item(arrayPomocneUserPattern(InxSel, 4))

            Dim shapes1 As MECMOD.Shapes
            shapes1 = body1.Shapes

            Dim userPattern1 As UserPattern
            userPattern1 = shapes1.Item(arrayPomocneUserPattern(InxSel, 0))

            oSelection.Add(userPattern1)
            If arrayPomocneUserPattern(InxSel, 3) = "Gwint" Then
                oSelection.VisProperties.SetRealColor(0, 0, 0, 0)
            ElseIf arrayPomocneUserPattern(InxSel, 3) = "Kolek" Then
                oSelection.VisProperties.SetRealColor(0, 255, 0, 0)
            ElseIf arrayPomocneUserPattern(InxSel, 3) = "Z poza zakresu" Then
                'USTAWIAM szary kolor dla innych otworów
                oSelection.VisProperties.SetRealColor(210, 210, 255, 0)
            End If
        Next
        oSelection.Clear()
    End Sub
    'Coloring all RectPatterns
    Sub colorWithRectPattern()
        Dim arrayPomocneRectPattern(0, 2) As String
        Dim oSelection
        Dim ileRectPattern As Integer
        Dim visPropertySet As VisPropertySet
        Dim czyGwintowanyOtwor As CatHoleThreadingMode
        Dim partDocumentToCheck As MECMOD.PartDocument

        oSelection = mainDoc.Selection
        oSelection.Clear()

        visPropertySet = oSelection.VisProperties

        oSelection.Search("n:*RectPattern.*,all")

        ileRectPattern = oSelection.Count

        If ileRectPattern = 0 Then Exit Sub

        ReDim arrayPomocneRectPattern(ileRectPattern - 1, 4)

        'Ładowanie arrayPomocne
        For i = 0 To oSelection.Count - 1
            czyGwintowanyOtwor = oSelection.Item(i + 1).Value.ItemToCopy.ThreadingMode
            arrayPomocneRectPattern(i, 0) = oSelection.Item(i + 1).Value.Name
            arrayPomocneRectPattern(i, 1) = oSelection.Item(i + 1).Value.Parent.Parent.Parent.Parent.Parent.Name
            arrayPomocneRectPattern(i, 4) = oSelection.Item(i + 1).Value.ItemToCopy.Parent.Parent.Name 'Nazwa np PartBody
            partDocumentToCheck = CATIA.Documents.Item(arrayPomocneRectPattern(i, 1))
            If InStr(1, partDocumentToCheck.Path, "pp") = 1 Then
                arrayPomocneRectPattern(i, 3) = "Z poza zakresu"
                GoTo OverRectPatternCheckingThreadingMode
            End If
            If czyGwintowanyOtwor = CatHoleThreadingMode.catThreadedHoleThreading Then
                'jeżeli hole jest z gwintem to:
                arrayPomocneRectPattern(i, 2) = oSelection.Item(i + 1).Value.ItemToCopy.HoleThreadDescription.Value ' tylko dla thread np M10
                If arrayOfSruby.Contains(arrayPomocneRectPattern(i, 2)) Then
                    arrayPomocneRectPattern(i, 3) = "Gwint"
                    pomalowanychOtworow = pomalowanychOtworow + 1
                Else
                    arrayPomocneRectPattern(i, 3) = "Z poza zakresu"
                End If
            Else
                arrayPomocneRectPattern(i, 2) = oSelection.Item(i + 1).Value.ItemToCopy.Diameter.Value 'np 10 
                If arrayOfKolki.Contains(arrayPomocneRectPattern(i, 2)) Then
                    arrayPomocneRectPattern(i, 3) = "Kolek"
                    pomalowanychOtworow = pomalowanychOtworow + 1
                Else
                    arrayPomocneRectPattern(i, 3) = "Z poza zakresu"
                End If
            End If

OverRectPatternCheckingThreadingMode:

        Next

        'Druga Petla
        For InxSel = 0 To ileRectPattern - 1
            oSelection.Clear()

            Dim documents1 As Documents
            documents1 = CATIA.Documents

            Dim partDocument1 As MECMOD.PartDocument
            partDocument1 = documents1.Item(arrayPomocneRectPattern(InxSel, 1))

            Dim part1 As MECMOD.Part
            part1 = partDocument1.Part

            Dim bodies1 As MECMOD.Bodies
            bodies1 = part1.Bodies

            Dim body1 As MECMOD.Body
            body1 = bodies1.Item(arrayPomocneRectPattern(InxSel, 4))

            Dim shapes1 As MECMOD.Shapes
            shapes1 = body1.Shapes

            Dim rectPattern1 As RectPattern
            rectPattern1 = shapes1.Item(arrayPomocneRectPattern(InxSel, 0))

            oSelection.Add(rectPattern1)
            If arrayPomocneRectPattern(InxSel, 3) = "Gwint" Then
                oSelection.VisProperties.SetRealColor(0, 0, 0, 0)
            ElseIf arrayPomocneRectPattern(InxSel, 3) = "Kolek" Then
                oSelection.VisProperties.SetRealColor(0, 255, 0, 0)
            ElseIf arrayPomocneRectPattern(InxSel, 3) = "Z poza zakresu" Then
                'USTAWIAM szary kolor dla innych otworów
                oSelection.VisProperties.SetRealColor(210, 210, 255, 0)
            End If
        Next
        oSelection.Clear()
    End Sub
    'Display result of coloring
    Sub Result()
        Dim str As String
        Dim sResult As String = ""

        For Each str In resztaHole
            sResult &= str & Environment.NewLine
        Next
        MsgBox("No colored holes: " & Environment.NewLine & sResult & Environment.NewLine & Environment.NewLine & "Colored: " & pomalowanychOtworow & " elements.")

    End Sub

End Module