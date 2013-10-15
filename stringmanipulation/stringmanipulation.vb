Public Class stringmanipulation

    '=====================================================================
    'This function removes all special charaters shown in the array then it returns the string with the chars removed
    'Preconditions:
    '   a string is required to search and remove any of the special characters

    Function strip_special_chars(ByVal strip_string As String)
        Dim str_org_filename As String
        Dim int_counter As Integer
        Dim arr_special_chars() As String

        arr_special_chars = {",", ".", "<", ">", "/", "=", """", "@", "!", "$", "%", "^", "&", "*", "(", ")", "{", "}", "[", "]", "?", "\", "|"}

        str_org_filename = strip_string

        int_counter = 0

        Dim i As Integer

        For i = 0 To arr_special_chars.Length - 1
            Do Until int_counter = 23
                strip_string = Replace(str_org_filename, arr_special_chars(i), "")
                int_counter = int_counter + 1
                str_org_filename = strip_string
            Loop
            int_counter = 0
        Next

        Return str_org_filename

    End Function
End Class