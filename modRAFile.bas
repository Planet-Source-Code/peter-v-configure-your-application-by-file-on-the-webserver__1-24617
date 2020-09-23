Attribute VB_Name = "modRAFile"
'modRAFile  => Remote Access File
'Module that search for the Subject & item in the subject file
'that was downloaded thru the internet - webpage..

Function GetValue(Str As String, Subject As String, Item As String) As String
Dim A, B, C, D
Dim beginItem, endItem
Dim Result
Dim NextSpace

'Str is the File downloaded on the net..
'if Subject = CHECKBOX  =>   Subject = [CHECKBOX]
Subject = "[" & Subject & "]"
A = InStr(1, Str, Subject, vbTextCompare)
If A <> 0 Then    '0 = Not FOUND !!
    'Now We have the Subject..=> searh for next subject..
    D = InStr(A + Len(Subject) + 1, Str, "[", vbTextCompare)
        If D = 0 Then
            'Next Subject is not found..Use MID function To End String
            B = A + Len(Subject) + 1
            'ITem Search
            C = InStr(B, Str, Item, vbTextCompare)   'OR Substring > Item...
                    If C <> 0 Then
                    NextSpace = InStr(C, Str, " ", vbTextCompare)
                    End If
                If (NextSpace - C) > Len(Item) Or C = 0 Then
                    MsgBox "Item : " & Item & " not  found"
                Exit Function
                Else
                'Found./..
                 beginItem = InStr(C, Str, "=", vbTextCompare)
                 endItem = InStr(C + 1, Str, ";", vbTextCompare)
                    'if VbCrLf not is found => Last line of file...
                    If endItem = 0 Then
                        endItem = Len(Str)
                    End If
                ' Clear Space's  in the string..
                 Result = Trim(Mid(Str, beginItem + 1, (endItem - beginItem - 1)))
                GetValue = Result
                End If
            
        Else
         'A next Subject is found.. Use MID to Next Subject..
         B = A + Len(Subject) + 1
        'ITem Search
        'Debug.Print Mid(Str, B, D - B)
        C = InStr(1, Mid(Str, B, D - B), Item)
        If C <> 0 Then
            NextSpace = InStr(C, Mid(Str, B, D - B), " ")
        End If
                If (NextSpace - C) > Len(Item) Or C = 0 Then
                MsgBox "Item : " & Item & " not  found"
                Exit Function
            Else
                'Found./..
                 beginItem = InStr(B + C, Str, "=", vbTextCompare)
                 endItem = InStr(B + C, Str, ";", vbTextCompare)
                ' Clear Space's  in the string..
                Result = Trim(Mid(Str, beginItem + 1, (endItem - beginItem - 1)))
                GetValue = Result
            End If
      
        End If
    
Else
MsgBox "Subject " & Subject & " NOT FOUND !"
Exit Function
End If

End Function
