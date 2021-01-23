Imports MECMOD

Module Module1

    Sub Main()

        Dim CATIA As INFITF.Application = GetObject("Catia.Application")
        Dim partDocument1 As PartDocument = CATIA.ActiveDocument
        Dim part1 As Part = partDocument1.Part
        part1.Name = "Hello"

    End Sub

End Module
