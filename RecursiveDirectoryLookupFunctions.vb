Imports System.IO

Module RecursiveDirectoryLookupFunctions


    Private Function recursive_directory_lookup(ByRef topDirectory As String, ByVal targetDirectory As String) As String
        'returns the full path to the directory in question, reading up the filesystem tree until it finds (or fails to find)
        'the directory in question

        'If you are looking at C:/ then either you are finished and should return an error (directory not found) 
        'or return C:/targetDirectory with just 1 slash. 
        If topDirectory = Directory.GetDirectoryRoot(topDirectory).ToString Then

            If Directory.Exists(topDirectory & targetDirectory) Then
                Return (topDirectory & targetDirectory)
            Else
                Throw New DirectoryNotFoundException("This directory does not exist anywhere on this branch of the directory tree")
            End If

            'Otherwise you are not on the root, so try to find the directory you are looking for
        Else
            'if it exists, return it (with the slash in the right place)
            If Directory.Exists(topDirectory & "\" & targetDirectory) Then
                Return topDirectory & "\" & targetDirectory
            Else
                'otherwise move up one directory and recursively try again
                topDirectory = Directory.GetParent(topDirectory).ToString()
                Return recursive_directory_lookup(topDirectory, targetDirectory)
            End If

        End If
    End Function

    Public Function find_my_directory(ByVal targetDirectory As String) As String
        'Set the top directory to be the current path of the executable. Then search for the directory you want. 
        Dim globalTopDirectory As String = My.Application.Info.DirectoryPath
        Return recursive_directory_lookup(globalTopDirectory, targetDirectory)
    End Function


End Module
