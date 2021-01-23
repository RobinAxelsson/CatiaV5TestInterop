Imports System

Module Program
    Public Shared CATIA As Application = Interaction.GetObject(, "CATIA.APPLICATION")
    Sub Main(args As String())
        Console.WriteLine("Hello World!")
    End Sub

End Module
