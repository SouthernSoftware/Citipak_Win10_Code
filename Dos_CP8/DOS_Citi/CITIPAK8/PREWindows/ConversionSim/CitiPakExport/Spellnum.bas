Attribute VB_Name = "Module1"

'----- Look up tables for number words
Const NumTbl$ = "123456789"
Const NumNames$ = "One  Two  ThreeFour Five Six  SevenEightNine Ten"
Const Teens$ = "Eleven    Twelve    Thirteen  Fourteen  Fifteen   Sixteen   Seventeen Eighteen  Nineteen"
Const Tens$ = "Ten     Twenty  Thirty  Forty   Fifty   Sixty   Seventy Eighty  Ninety"
Const Powers3$ = "Thousand Million  Billion  Trillion"

'DECLARE FUNCTION SpellNumber$ (Number$)

'******* Returns a spelled out version of a number
Function SpellNumber$(StrNum$)    ' STATIC

  Dim Num$
  Dim x$
  Dim n As Integer
  Dim Temp As Integer
  Dim Word$
  Dim Sentence$
  
    SpellNumber$ = ""                           'Clear the function
    Num$ = LTrim$(RTrim$(StrNum$))              'Trim off any spaces

    x = InStr(Num$, ".")                        'Trim off any decimal places
    If x Then Num$ = Left$(Num$, x - 1)
    
    Length = Len(Num$)                          'Get the length
    If Length > 15 Then Exit Function           'Exit if bigger than trillions

    For n = Length To 1 Step -1                 'Step backwards through number

        x = InStr(NumTbl$, Mid$(Num$, n, 1)) - 1 'Look up the digit in table

        Select Case (Length - n) Mod 3          'Branch according to digit
                                                '  position
           '----- Ones digit
           Case 0
              If n < Length Then                'If not on last digit, look
                 For Temp = n To n - 2 Step -1  '  for non 0 digit
                    If Temp > 0 Then            'If not past end of number
                       Word$ = Mid$(Num$, Temp, 1)
                                                'If this is a non 0 digit,
                                                '  put power word in sentence
                       If Word$ <> "0" And Word$ <> "-" Then
                          Temp = ((Length - n) \ 3 - 1) * 9 + 1
                          Word$ = RTrim$(Mid$(Powers3$, Temp, 9))
                          Sentence$ = Word$ + " " + Sentence$
                          Exit For              'Bail out of search loop
                       End If
                    End If
                 Next
              End If

              If x > -1 Then                    'If digit found, get the word
                 Word$ = Mid$(NumNames$, x * 5 + 1, 5)
                 
                 If n > 1 Then                  'If left digit is one, use
                                                '  "Teen" table
                    If Mid$(Num$, n - 1, 1) = "1" Then
                       Word$ = Mid$(Teens$, x * 10 + 1, 10)
                       n = n - 1                'Skip the Tens digit
                    End If
                 End If
              End If

           '----- Tens digit
           Case 1
              If x > -1 Then                    'Find word in "Tens" table
                 Word$ = Mid$(Tens$, x * 8 + 1, 8)
              End If

           '----- Hundreds digit
           Case 2
              If x > -1 Then                    'Find word in number table
                 Word$ = Mid$(NumNames$, x * 5 + 1, 5)
                                                'Add the word "Hundred"
                 Word$ = RTrim$(Word$) + " Hundred"
              End If

        End Select

        If n = 1 And x = -1 Then                'Look for a minus sign at
           If Mid$(Num$, n, 1) = "-" Then       '  digit one
              Word$ = "Negative"                'Add it to sentence
              x = 0
           End If
        End If
                                                'If digit is non zero, add
                                                '  the word to the sentence
        If x > -1 Then Sentence$ = RTrim$(Word$) + " " + Sentence$
    Next
                                       
'****** Added "dollars and cents" directly to spellnum
'       02/22/94

    Sentence$ = RTrim$(Sentence$)
    Sentence$ = Sentence$ + " Dollars and "

    'Sentence$ = Sentence$ + " Dollar"
    'IF INT(VAL(Num$)) <> 1 THEN
    '  Sentence$ = Sentence$ + "s and "    'Anything but "One" is plural
    'ELSE
    '  Sentence$ = Sentence$ + " and "
    'END IF

    Sentence$ = Sentence$ + Mid$(StrNum$, InStr(StrNum$, ".") + 1) + " Cents"
    'Do cents part

    SpellNumber$ = RTrim$(Sentence$)            'Assign the function

    Num$ = ""                                   'Clean up work strings
    Word$ = ""
    Sentence$ = ""

End Function

