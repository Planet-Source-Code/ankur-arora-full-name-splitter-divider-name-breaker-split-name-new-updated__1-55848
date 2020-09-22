Attribute VB_Name = "Module1"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''     '''''''''''          ''''''''  ''''''''         ''''''''  ''''''''                ''''''''  '''''''''''''''''''''''''   '''''''''''''''''''''
''''''''   '''''''       '''''''    ''''''''''''         ''''''''  ''''''''        ''''''''   ''''''''                ''''''''  ''''''''''''''''''''''''''  '''''''''''''''''''''
''''''''  '''''''         '''''''   '''''''''''''        ''''''''  ''''''''       ''''''''    ''''''''                ''''''''  '''''''''''''''''''''''''''  ''''''''''''''''''''
''''''''  ''''''           '''''''  ''''''''''''''       ''''''''  ''''''''      ''''''''     ''''''''                ''''''''  ''''''''' ''''''''''''''''''  '''''''''''''''''''
''''''''  ''''''            ''''''  '''''''''''''''      ''''''''  ''''''''     ''''''''      ''''''''                ''''''''  ''''''''' '''''  ''''''''''''  ''''''''''''''''''
''''''''  ''''''            ''''''  ''''''''''''''''     ''''''''  ''''''''    ''''''''       ''''''''                ''''''''  ''''''''' ''''''  ''''''''''''  '''''''''''''''''
''''''''  ''''''            ''''''  '''''''''''''''''    ''''''''  ''''''''   ''''''''        ''''''''                ''''''''  ''''''''' '''''''  '''''''''''  '''''''''''''''''
''''''''  ''''''            ''''''  ''''''''''''''''''   ''''''''  ''''''''  ''''''''         ''''''''                ''''''''  ''''''''' '''''''   ''''''''''  '''''''''''''''''
''''''''  ''''''''''''''''''''''''  '''''''''''''''''''  ''''''''  '''''''' ''''''''          ''''''''                ''''''''  ''''''''' '''''''   '''''''''  ''''''''''''''''''
''''''''  ''''''''''''''''''''''''  '''''''''''''''''''' ''''''''  ''''''''''''''''           ''''''''                ''''''''  ''''''''' ''''''    ''''''''  '''''''''''''''''''
''''''''  ''''''''''''''''''''''''  '''''''''''''''''''''''''''''  '''''''''''''''            ''''''''                ''''''''  '''''''''           '''''''  ''''''''''''''''''''
''''''''  ''''''''''''''''''''''''  ''''''''    '''''''''''''''''  '''''''''''''''            ''''''''                ''''''''  ''''''''''''''''''''''''''  '''''''''''''''''''''
''''''''  ''''''          ''''''''  ''''''''     ''''''''''''''''  '''''''''''''''            ''''''''                ''''''''  '''''''''''''''''''''''''  ''''''''''''''''''''''
''''''''  ''''''          ''''''''  ''''''''      '''''''''''''''  ''''''''''''''''           ''''''''                ''''''''  ''''''''''''''''''  '''''''''''''''''''''''''''''
''''''''  ''''''''''''''''''''''''  ''''''''       ''''''''''''''  '''''''' ''''''''          ''''''''                ''''''''  '''''''' ''''''''''  ''''''''''''''''''''''''''''
''''''''  ''''''''''''''''''''''''  ''''''''        '''''''''''''  ''''''''  ''''''''         ''''''''                ''''''''  ''''''''  ''''''''''  '''''''''''''''''''''''''''
''''''''  ''''''''''''''''''''''''  ''''''''         ''''''''''''  ''''''''   ''''''''        ''''''''                ''''''''  ''''''''   ''''''''''  ''''''''''''''''''''''''''
''''''''  ''''''''''''''''''''''''  ''''''''          '''''''''''  ''''''''    ''''''''       '''''''''              '''''''''  ''''''''    ''''''''''  '''''''''''''''''''''''''
''''''''  ''''''          ''''''''  ''''''''           ''''''''''  ''''''''     ''''''''      ''''''''''            ''''''''''  ''''''''     ''''''''''  ''''''''''''''''''''''''
''''''''  ''''''          ''''''''  ''''''''            '''''''''  ''''''''      ''''''''     '''''''''''          '''''''''''  ''''''''      ''''''''''  '''''''''''''''''''''''
''''''''  ''''''          ''''''''  ''''''''             ''''''''  ''''''''       ''''''''    ''''''''''''        ''''''''''''  ''''''''       ''''''''''  ''''''''''''''''''''''
''''''''  ''''''          ''''''''  ''''''''              '''''''  ''''''''        ''''''''   ''''''''''''''''''''''''''''''''  ''''''''        ''''''''''  '''''''''''''''''''''
''''''''  ''''''          ''''''''  ''''''''               ''''''  ''''''''         ''''''''  ''''''''''''''''''''''''''''''''  ''''''''         ''''''''''  ''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ------------------------------------------------------------------- '
' THIS APPLICATION IS WRITTEN BY
' ANKUR ARORA FOR SPLITTING A COMPLETE NAME INTO FIVE PARTS
' I.E. INTIALS/FIRSTNAME/MIDDLENAME/LASTNAME/SUFFIX
' ------------------------------------------------------------------- '
Option Explicit
Dim setIndex As Integer
Public server_name, Database, Userid, Password As String
Public nameTitle, nameSuffix, firstName, middleName, lastName As String
' ------------------------------------------------------------------- '


Public Function SplitName(ByVal strFullname As String) As String
    Dim i As Integer
    Dim rsTitles(), rsSuffix() As String
    ReDim rsTitles(100)
    ReDim rsSuffix(100)
    
    ' ------------------------------------------------------------------- '
    ' FOR STORING INTIALS OF NAMES
    ' IF YOU WANT MORE JUST ADD MORE VALUES BY INCREASING ARRAY'S INDEX
    ' ------------------------------------------------------------------- '
    rsTitles(1) = "Mr."
    rsTitles(2) = "Mrs."
    rsTitles(3) = "Dr."
    rsTitles(4) = "Gen."
    rsTitles(5) = "Hon."
    rsTitles(6) = "Ms."
    rsTitles(7) = "Msgr."
    rsTitles(8) = "Prof."
    rsTitles(9) = "Rep."
    rsTitles(10) = "Rev."
    rsTitles(11) = "Mr"
    rsTitles(12) = "Mrs"
    rsTitles(13) = "Dr"
    rsTitles(14) = "Gen"
    rsTitles(15) = "Hon"
    rsTitles(16) = "Ms"
    rsTitles(17) = "Msgr"
    rsTitles(18) = "Prof"
    rsTitles(19) = "Rep"
    rsTitles(20) = "Rev"
    ' ------------------------------------------------------------------- '
    

    
    ' ------------------------------------------------------------------- '
    ' FOR STORING SUFFIX OF NAMES
    ' IF YOU WANT MORE JUST ADD MORE VALUES BY INCREASING ARRAY'S INDEX
    ' ------------------------------------------------------------------- '
    rsSuffix(1) = "/sp"
    rsSuffix(2) = "C.P.A."
    rsSuffix(3) = "Ed.D."
    rsSuffix(4) = "Esq."
    rsSuffix(5) = "I"
    rsSuffix(6) = "II"
    rsSuffix(7) = "III"
    rsSuffix(8) = "IV"
    rsSuffix(9) = "V"
    rsSuffix(10) = "Jr."
    rsSuffix(11) = "M.D."
    rsSuffix(12) = "Ph.D."
    rsSuffix(13) = "Sr."
    rsSuffix(14) = "DDS"
    rsSuffix(15) = "CPA"
    rsSuffix(16) = "EdD"
    rsSuffix(17) = "Esq"
    rsSuffix(18) = "Jr"
    rsSuffix(19) = "MD"
    rsSuffix(20) = "PhD"
    rsSuffix(21) = "Sr"
    rsSuffix(22) = "DDS"
    ' ------------------------------------------------------------------- '
    



    ' ------------------------------------------------------------------- '
    ' FOR FINDING THE INTIALS FROM THE NAME
    ' ------------------------------------------------------------------- '
    Dim instr1, instr2
    For i = 1 To UBound(rsTitles)
        If rsTitles(i) <> "" Or rsTitles(i) <> Empty Then
            instr1 = InStr(1, LCase(strFullname), LCase(rsTitles(i)))
            If instr1 <> 0 Then
                Dim nextChar, prevChar As String
                nextChar = Mid(strFullname, instr1 + Len(rsTitles(i)), 1)
                prevChar = Mid(strFullname, instr1 + Len(rsTitles(i)) - 1, 1)
                If (LCase(nextChar) < "a" And LCase(nextChar) > "z") Or nextChar = "," Or nextChar = " " Or prevChar = "." Then _
                    nameTitle = Mid(strFullname, instr1, Len(rsTitles(i)))
                Exit For
            End If
        End If
    Next
    ' ------------------------------------------------------------------- '
    
    
    
    
    ' ------------------------------------------------------------------- '
    ' FOR FINDING THE SUFFIX FROM THE NAME
    ' ------------------------------------------------------------------- '
    For i = 1 To UBound(rsSuffix)
        If rsSuffix(i) <> "" Or rsSuffix(i) <> Empty Then
            instr1 = InStr(1, LCase(strFullname), LCase(rsSuffix(i)))
            If instr1 <> 0 Then
                prevChar = Mid(strFullname, instr1 - 1, 1)
                If (instr1 + Len(rsSuffix(i)) - 1) = Len(strFullname) Then
                    'If (LCase(prevChar) < "a" And LCase(prevChar) > "z")Then
                    If prevChar = " " Or prevChar = "," Then _
                        nameSuffix = Mid(strFullname, instr1, Len(rsSuffix(i)))
                    Exit For
                End If
            End If
        End If
    Next
    ' ------------------------------------------------------------------- '
    
    
    
    
    ' ------------------------------------------------------------------- '
    ' BREAKING THE FULLNAME & FETCHING THE TITLE OUT OF IT
    ' ------------------------------------------------------------------- '
    If nameTitle <> "" Then
        instr1 = InStr(1, strFullname, nameTitle)
        Dim startChar As String
        Dim cnt As Integer
        While Not (LCase(startChar) >= "a" And LCase(startChar) <= "z")
            cnt = cnt + 1
            startChar = Mid(strFullname, instr1 + Len(nameTitle) + cnt, 1)
        Wend
        If cnt > 1 Then
            strFullname = Mid(strFullname, instr1 + Len(nameTitle) + cnt)
        Else
            strFullname = Mid(strFullname, instr1 + Len(nameTitle))
        End If
    End If
    ' ------------------------------------------------------------------- '
    
    
    
    
    ' ------------------------------------------------------------------- '
    ' BREAKING THE FULLNAME & FETCHING THE SUFFIX OUT OF IT
    ' ------------------------------------------------------------------- '
    If nameSuffix <> "" Then
        instr1 = InStr(1, strFullname, nameSuffix)
        strFullname = Mid(strFullname, 1, instr1 - 1)
    End If
    ' ------------------------------------------------------------------- '
    
    
    
    ' ------------------------------------------------------------------- '
    ' RECREATING THE FULLNAME WITH COMMA(,) AS SEPRATOR BETWEEN
    ' FIRSTNAME MIDDLENAME & LASTNAME AND REMOVING SPACES FROM THE FULLNAME
    ' ------------------------------------------------------------------- '
    strFullname = Trim(strFullname)
    Dim instr3, c, c1, spaceCount
    For i = 1 To Len(strFullname)
        c = Mid(strFullname, i, 1)
        If c = " " Then
            c1 = Mid(strFullname, i - 1, 1)
            If c1 <> "," Then
                strFullname = Replace(strFullname, " ", ",")
            Else
                strFullname = Replace(strFullname, " ", "")
            End If
        End If
    Next
    c = Mid(strFullname, 1, 1)
    If c = "," Then _
        strFullname = Mid(strFullname, 2)
    
    c = Mid(strFullname, Len(strFullname), 1)
    If c = "," Then _
        strFullname = Mid(strFullname, 1, Len(strFullname) - 1)
    ' ------------------------------------------------------------------- '
    
    
    
    
    ' ------------------------------------------------------------------- '
    ' STORING COUNT OF COMMA(S)(,) IN FULLNAME
    ' ------------------------------------------------------------------- '
    For i = 1 To Len(strFullname)
        c = Mid(strFullname, i, 1)
        If c = "," Then
            spaceCount = spaceCount + 1
        End If
    Next
    ' ------------------------------------------------------------------- '
    
    
    
    ' ------------------------------------------------------------------- '
    ' BREAKING NAME INTO FIRST MIDDLE & LAST NAME AS PER COMMA(S) REPETION
    ' CASE1: WIL GIVE
    '   FIRSTNAME & LASTNAME
    ' CASE2: WIL GIVE
    '   FIRSTNAME & MIDDLENAME & LASTNAME
    ' ------------------------------------------------------------------- '
    instr1 = InStr(1, strFullname, ",")
    On Error GoTo ExtraChars
    Select Case spaceCount
        Case Empty:
            firstName = strFullname
        Case 1:
            firstName = Mid(strFullname, 1, instr1 - 1)
            lastName = Mid(strFullname, instr1 + 1)
        Case 2:
            firstName = Mid(strFullname, 1, instr1 - 1)
            instr2 = InStr(instr1 + 1, strFullname, ",")
            middleName = Mid(strFullname, instr1 + 1, instr2 - instr1 - 1)
            lastName = Mid(strFullname, instr2 + 1)
            
    End Select
    ' ------------------------------------------------------------------- '
    
    
    
    
    MsgBox "Title is: " & nameTitle, , "Split Name"
    MsgBox "First Name is: " & firstName, , "Split Name"
    MsgBox "Middle Name is: " & middleName, , "Split Name"
    MsgBox "Last Name is: " & lastName, , "Split Name"
    MsgBox "Suffix is: " & nameSuffix, , "Split Name"
    
    nameTitle = ""
    firstName = ""
    middleName = ""
    lastName = ""
    nameSuffix = ""
    
    Exit Function

ExtraChars:
            MsgBox "This Name Contains Invalid Comma(s) or Space(s)" & vbCrLf & "Unable to break it...", vbOKOnly + vbExclamation
            Exit Function
End Function
