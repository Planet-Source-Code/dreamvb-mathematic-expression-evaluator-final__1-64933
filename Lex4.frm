VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recursive Decent Parsing v2"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Test 3 Functions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5220
      TabIndex        =   3
      Top             =   1650
      Width           =   1860
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5220
      TabIndex        =   2
      Top             =   900
      Width           =   1860
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5220
      TabIndex        =   1
      Top             =   2370
      Width           =   1860
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5220
      TabIndex        =   0
      Top             =   120
      Width           =   1860
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By DreamVB
'This is my Recursive Decent Parsing project. I made this project to maybe help others
'That may need to use something like this in there projects.
'I made this after reading though one of my C++ books and decided to convert my C++ project over to VB


'Token Types
Enum token_type
    LERROR = -1
    NONE = 0
    DELIMITER = 1
    DIGIT = 2
    LSTRING = 3
    VARIABLE = 4
    IDENTIFIER = 6
    HEXDIGIT = 5
    FINISHED = 7
End Enum

'Relational
Const GE = 1 ' Greator than or equal to
Const NE = 2 ' not equal to
Const LE = 3 ' Less than or equal to

'Bitwise
Const cAND = 4
Const cOR = 5
'Bitshift

Const shr = 6
Const shl = 7
Const cXor = 8
Const cIMP = 9
Const cEqv = 11
Const cINC = 12

Const Str_Ops = "AND,OR,XOR,MOD,DIV,EXP,SHL,SHR,IMP,EQV,NOT"
Const Str_Funcs = "ABS,ATN,COS,EXP,LOG,RND,ROUND,SGN,SIN,SQR,TAN,SUM"

'We use this to store variables
Private Type vars
    vName As String
    vValue As Variant
End Type

Dim Token As String         'Current processing token
Dim tok_type As token_type  'Used to idenfiy the tokens
Dim Look_Pos As Long    'Current processing char pointer
Dim ExprLine As String  'The Expression line to scan

Dim lVars() As vars   '26 variables
Dim lVarCount As Integer

Function isIdent(sIdentName As String) As Boolean
Dim x As Integer, Idents As Variant
    Idents = Split(Str_Ops, ",")
    
    For x = 0 To UBound(Idents)
        If LCase(Idents(x)) = LCase(sIdentName) Then
            isIdent = True
            Exit For
        End If
    Next x
    
    x = 0
    Erase Idents
End Function

Function IsIdentFunc(sIdentName As String) As Boolean
Dim x As Integer, Idents As Variant
    Idents = Split(Str_Funcs, ",")
    
    For x = 0 To UBound(Idents)
        If LCase(Idents(x)) = LCase(sIdentName) Then
            IsIdentFunc = True
            Exit For
        End If
    Next x
    
    x = 0
    Erase Idents
End Function

Sub PushBack()
Dim tok_len As Integer
    tok_len = Len(Token)
    Look_Pos = Look_Pos - tok_len
End Sub

Sub init()
    lVarCount = 0
    Erase lVars
End Sub

Sub abort(code As Integer, Optional aStr As String = "")
Dim lMsg As String
    Select Case code
        Case 0: lMsg = "Undeclared variable '" & aStr & "'"
        Case 1: lMsg = "Division by zero"
        Case 2: lMsg = "Expected parenthesized closing bracket ')'"
        Case 3: lMsg = "Invalid Digit found '" & aStr & "'"
        Case 4: lMsg = "Unknown character found '" & aStr & "'"
        Case 5: lMsg = "The Variable '" & aStr & "' is an identifier and can't be used."
        Case 6: lMsg = "Expected expression."
        Case 7: lMsg = "Invalid Hexadecimal value found '0x" & UCase(aStr) & "'"
    End Select
    
    MsgBox lMsg, vbInformation, "Error"
    Look_Pos = Len(ExprLine) + 1
    End
End Sub

Sub AddVar(name As String, Optional lValue As Variant = 0)
    'Add a new variable along with the variables value.
    
    ReDim Preserve lVars(lVarCount)     'Resize variable stack
    lVars(lVarCount).vName = name       'Add variable name
    lVars(lVarCount).vValue = lValue    'Add varaible data
    lVarCount = lVarCount + 1           'INC variable Counter
End Sub

Function FindVarIdx(name As String) As Integer
Dim x As Integer, idx As Integer
    'Locate a variables position in the variables array
    idx = -1 'Bad position
    For x = 0 To UBound(lVars)
        If LCase(name) = LCase(lVars(x).vName) Then
            idx = x
            Exit For
        End If
    Next x
    FindVarIdx = idx
End Function

Sub SetVar(vIdx, Optional lData As Variant = 0)
    'Set a variables value, by using the variables index vIdx
    lVars(vIdx).vValue = lData
End Sub

Function GetVarData(name As String) As Variant
    'Return data from a variable stored in the variable stack
    GetVarData = lVars(FindVarIdx(name)).vValue
End Function

Function IsDelim(c As String) As Boolean
    'Return true if we have a Delimiter
    If InStr(" ;,+-<>^=(*)/\%&|!", c) Then IsDelim = True
End Function

Function isAlphaNum(c As String) As Boolean
    isAlphaNum = (isDigit(c) Or isAlpha(c))
End Function

Function isAlpha(c As String) As Boolean
    'Return true if we only have letters a-z  A-Z
    isAlpha = UCase(c) >= "A" And UCase(c) <= "Z"
End Function

Function isWhite(c As String) As Boolean
    'Return true if we find a white space
    isWhite = (c = " ") Or (c = vbTab)
End Function

Function isDigit(c As String) As Boolean
    'Return true when we only have a digit
    isDigit = (c >= "0") And (c <= "9")
End Function

Function isHex(HexVal As String) As Boolean
Dim x As Integer, c As String
    For x = 1 To Len(HexVal)
        c = Mid(HexVal, x, 1)

        Select Case UCase(c)
            Case 0 To 9: isHex = True
            Case "A", "B", "C", "D", "E", "F": isHex = True
            Case Else
                isHex = False
                Exit For
        End Select
    Next x
End Function

Function Eval(expression As String)
    ExprLine = expression 'Store the expression to scan
    Look_Pos = 1    'Default state of char pos
    GetToken        'Kick start and Get the first token.
    
    If (tok_type = FINISHED) Or Len(Trim(expression)) = 0 Then
        MsgBox "Error No Expression Found", vbExclamation, "Nothing to phase"
        Exit Function
    Else
        Eval = Exp0()
    End If
    
End Function

Sub GetToken()
Dim Temp As String
Dim idx As Integer
Dim dTmp


    Temp = ""
    'This is the main part of the pharser and is used to.
    'Identfiy all the tokens been scanned and return th correct token type
    
    'Clear current token info
    Token = ""
    tok_type = NONE
    
    If Look_Pos > Len(ExprLine) Then tok_type = FINISHED: Exit Sub
    'Above exsits the sub if we are passed expr len
    
    Do While (Look_Pos <= Len(ExprLine) And isWhite(Mid(ExprLine, Look_Pos, 1)))
        'Skip over white spaces. and stay within the expr len
        Look_Pos = Look_Pos + 1 'INC
        If Look_Pos > Len(ExprLine) Then Exit Sub
    Loop
    
    'Some little test I was doing to do Increment/Decrement operators -- ++
    If ((Mid(ExprLine, Look_Pos, 1) = "+") Or Mid(ExprLine, Look_Pos, 1) = "-") Then
        If ((Mid(ExprLine, Look_Pos + 1, 1) = "+") Or Mid(ExprLine, Look_Pos + 1, 1) = "-") Then
            Temp = Mid(ExprLine, 1, Look_Pos - 1)
            If Mid(ExprLine, Look_Pos + 1, 1) = "+" Then
                dTmp = GetVarData(Temp) + 1
            ElseIf Mid(ExprLine, Look_Pos + 1, 1) = "-" Then
                 dTmp = GetVarData(Temp) - 1
            End If
            
            SetVar FindVarIdx(Temp), dTmp
            Token = Temp
            Exit Sub
        End If
    End If
    ''
    If (Mid(ExprLine, Look_Pos, 1) = "&") Or Mid(ExprLine, Look_Pos, 1) = "|" Then
        'Bitwise code, I still got some work to do on this yet but it does the ones
        ' that are listed below fine
        Select Case Mid(ExprLine, Look_Pos, 1)
        Case "&"
            If Mid(ExprLine, Look_Pos + 1, 1) = "&" Then
                Look_Pos = Look_Pos + 2
                Token = Chr(cAND)
                Exit Sub
            Else
                Look_Pos = Look_Pos + 1
                Token = "&"
                Exit Sub
            End If
        Case "|"
            If Mid(ExprLine, Look_Pos + 1, 1) = "|" Then
                Look_Pos = Look_Pos + 2
                Token = Chr(cOR)
                Exit Sub
            Else
                Look_Pos = Look_Pos + 1
                Token = "|"
                Exit Sub
            End If
        tok_type = DELIMITER
        End Select
    End If

    If (Mid(ExprLine, Look_Pos, 1) = "<") Or (Mid(ExprLine, Look_Pos, 1) = ">") Then
        'Check for Relational operators < > <= >= <>
        'check for not equal to get first op <
        Select Case Mid(ExprLine, Look_Pos, 1)
            Case "<"
                If Mid(ExprLine, Look_Pos + 1, 1) = ">" Then
                    'Not Equal to
                    Look_Pos = Look_Pos + 2
                    Token = Chr(NE)
                    Exit Sub
                ElseIf Mid(ExprLine, Look_Pos + 1, 1) = "=" Then
                    'Less then of equal to
                    Look_Pos = Look_Pos + 2
                    Token = Chr(LE)
                    Exit Sub
                ElseIf Mid(ExprLine, Look_Pos + 1, 1) = "<" Then
                    'Bitshift left
                    Look_Pos = Look_Pos + 2
                    Token = Chr(shl)
                    Exit Sub
                Else
                    'Less then
                    Look_Pos = Look_Pos + 1
                    Token = "<"
                    Exit Sub
                End If
            Case ">"
                If Mid(ExprLine, Look_Pos + 1, 1) = "=" Then
                    'Greator than or equal to
                    Look_Pos = Look_Pos + 2
                    Token = Chr(GE)
                    Exit Sub
                ElseIf Mid(ExprLine, Look_Pos + 1, 1) = ">" Then
                    Look_Pos = Look_Pos + 2
                    Token = Chr(shr)
                    Exit Sub
                Else
                    'Greator than
                    Look_Pos = Look_Pos + 1
                    Token = ">"
                    Exit Sub
                End If
                tok_type = DELIMITER
        End Select
    End If
    
    If IsDelim(Mid(ExprLine, Look_Pos, 1)) Then
        'Check if we have a Delimiter ;,+-<>^=(*)/\%
        Token = Token + Mid(ExprLine, Look_Pos, 1) 'Get next char
        Look_Pos = Look_Pos + 1 'INC
        tok_type = DELIMITER 'Delimiter Token type
    ElseIf isDigit(Mid(ExprLine, Look_Pos, 1)) Then
        'See if we are dealing with a Hexadecimal Value
        If Mid(ExprLine, Look_Pos + 1, 1) = "x" Then
            Do While (isAlphaNum(Mid(ExprLine, Look_Pos, 1)))
                Token = Token + Mid(ExprLine, Look_Pos, 1)
                Look_Pos = Look_Pos + 1
                tok_type = HEXDIGIT
            Loop
            Exit Sub
        End If
        'Check if we are dealing with only digits 0 .. 9
        Do While (IsDelim(Mid(ExprLine, Look_Pos, 1))) = 0
            Token = Token + Mid(ExprLine, Look_Pos, 1) 'Get next char
            Look_Pos = Look_Pos + 1 'INC
        Loop
        tok_type = DIGIT 'Digit token type
        
    ElseIf isAlpha(Mid(ExprLine, Look_Pos, 1)) Then
        'Check if we have strings Note no string support in this version
        ' this is only used for variables.
        Do While Not IsDelim(Mid(ExprLine, Look_Pos, 1))
            Token = Token + Mid(ExprLine, Look_Pos, 1)
            Look_Pos = Look_Pos + 1 'INC
            'tok_type = VARIABLE
            tok_type = LSTRING 'String token type
        Loop
    Else
        abort 4, Mid(ExprLine, Look_Pos, 1)
        tok_type = FINISHED
    End If
    
    If tok_type = LSTRING Then
        'check for identifiers
        If isIdent(Token) Then
            Select Case UCase(Token)
                Case "AND"
                    Token = Chr(cAND)
                    Exit Sub
                Case "OR"
                    Token = Chr(cOR)
                    Exit Sub
                Case "NOT"
                    Token = "!"
                    Exit Sub
                Case "IMP"
                    Token = Chr(cIMP)
                    Exit Sub
                Case "EQV"
                    Token = Chr(cEqv)
                    Exit Sub
                Case "DIV"
                    Token = "\"
                    Exit Sub
                Case "MOD"
                    Token = "%"
                    Exit Sub
                Case "XOR"
                    Token = Chr(cXor)
                    Exit Sub
                Case "SHL"
                    Token = Chr(shl)
                    Exit Sub
                Case "SHR"
                    Token = Chr(shr)
                    Exit Sub
            End Select
                tok_type = DELIMITER
                Exit Sub
            
        ElseIf IsIdentFunc(Token) Then
            tok_type = IDENTIFIER
           ' GetToken
            Exit Sub
        Else
            tok_type = VARIABLE
            Exit Sub
        End If
    End If
    
    
    
End Sub

Function Exp0()
Dim Tmp_tokType As token_type
Dim Tmp_Token As String
Dim Var_Idx As Integer
Dim Temp
    'Assignments
    If (tok_type = VARIABLE) Then
        'Store temp type and token
        'we first need to check if the variable name is not an identifier
        If isIdent(Token) Then abort 5, Token
        Tmp_tokType = tok_type
        Tmp_Token = Token
        'Locate the variables index
        Var_Idx = FindVarIdx(Token)
        'If we have an invaild var index -1 we Must add a new variable
        If (Var_Idx = -1) Then
            'Add the new variable
            AddVar Token
            'Now get the variable index agian
            Var_Idx = FindVarIdx(Token)
        Else
            Exp0 = Exp1
            Exit Function
        End If
        'Get the next token
        Call GetToken
        If (Token <> "=") Then
            PushBack 'Move expr pointer back
            Token = Tmp_Token       'Restore temp token
            tok_type = Tmp_tokType  'Restore temp token type
        Else
            'Carry on processing the expression
            Call GetToken
            'Set the variables value
            Temp = Exp1()
            SetVar Var_Idx, Temp
            Exp0 = Temp
        End If
    End If
    Exp0 = Exp1
    
End Function

Function Exp1()
Dim op As String, Relops As String
Dim Temp, rPos As Integer

    'Relational operators
    Relops = Chr(GE) + Chr(NE) + Chr(LE) + "<" + ">" + "=" + "!" + Chr(0)
    Exp1 = Exp2()
    
    op = Token 'Get operator
    rPos = InStr(1, Relops, op) 'Check for other ops in token <> =
    If rPos > 0 Then
        GetToken 'Get next token
        Temp = Exp2 'Store temp val
        Select Case op
            Case "<" 'less then
                Exp1 = CDbl(Exp1) < CDbl(Temp)
            Case ">" 'greator than
                Exp1 = CDbl(Exp1) > CDbl(Temp)
            Case Chr(NE)
                Exp1 = CDbl(Exp1) <> CDbl(Temp)
            Case Chr(LE)
                Exp1 = CDbl(Exp1) <= CDbl(Temp)
            Case Chr(GE)
                Exp1 = CDbl(Exp1) >= CDbl(Temp)
            Case "=" 'equal to
                Exp1 = CDbl(Exp1) = CDbl(Temp)
            Case "!"
                Exp1 = Not CDbl(Temp)
        End Select
        'op = Token
    End If
    
End Function

Function Exp2()
Dim op As String
Dim Temp
    'Add or Subtact two terms
    Exp2 = Exp3()
    op = Token 'Get operator
    
    Do While (op = "+" Or op = "-")
        GetToken 'Get next token
        Temp = Exp3() 'Temp value
        'Peform the expresion for the operator
        Select Case op
            Case "-"
                Exp2 = CDbl(Exp2) - CDbl(Temp)
            Case "+"
                Exp2 = CDbl(Exp2) + CDbl(Temp)
        End Select
        op = Token
    Loop
    
End Function

Function Exp3()
Dim op As String
Dim Temp
    'Multiply or Divide two factors
    Exp3 = Exp4()
    op = Token 'Get operator
    Do While (op = "*" Or op = "/" Or op = "\" Or op = "%")
        GetToken 'Get next token
        Temp = Exp4() 'Temp value
        'Peform the expresion for the operator
        Select Case op
            Case "*"
                Exp3 = CDbl(Exp3) * CDbl(Temp)
            Case "/"
                If Temp = 0 Then abort 1
                Exp3 = CDbl(Exp3) / CDbl(Temp)
            Case "\"
                If Temp = 0 Then abort 1
                Exp3 = CDbl(Exp3) \ CDbl(Temp)
            Case "%"
                If Temp = 0 Then abort 1
                Exp3 = CDbl(Exp3) Mod CDbl(Temp)
        End Select
        op = Token
    Loop

End Function

Function Exp4()
Dim op As String, BitWOps As String
Dim Temp, rPos As Integer

    'Bitwise operators ^ | & || &&
    BitWOps = Chr(cAND) + Chr(cOR) + Chr(shl) + Chr(shr) + _
    Chr(cXor) + Chr(cIMP) + Chr(cEqv) + "^" + "|" + "&" + Chr(0)
    Exp4 = Exp5()
    
    op = Token 'Get operator
    rPos = InStr(1, BitWOps, op) 'Check for other ops in token <> =
    If rPos > 0 Then
        GetToken 'Get next token
        Temp = Exp5 'Store temp val
        Select Case op
            Case "^" 'Excompnent
                Exp4 = CDbl(Exp4) ^ CDbl(Temp)
            Case "&"
                Exp4 = CDbl(Exp4) & CDbl(Temp)
            Case Chr(cAND)
                Exp4 = CDbl(Exp4) And CDbl(Temp)
            Case Chr(cOR)
                Exp4 = CDbl(Exp4) Or CDbl(Temp)
            Case Chr(shl)
                'Bitshift Shift left
                Exp4 = CDbl(Exp4) * (2 ^ CDbl(Temp))
            Case Chr(shr)
                'bitshift right
                Exp4 = CDbl(Exp4) \ (2 ^ CDbl(Temp))
            Case Chr(cXor)
                'Xor
                Exp4 = CDbl(Exp4) Xor CDbl(Temp)
            Case Chr(cIMP)
                'IMP
                Exp4 = CDbl(Exp4) Imp CDbl(Temp)
            Case Chr(cEqv)
                Exp4 = CDbl(Exp4) Eqv CDbl(Temp)
        End Select
        'op = Token
    End If
    
End Function

Function Exp5()
Dim op As String
Dim Temp As Variant
    op = ""
    'Unary +,-
    If ((tok_type = DELIMITER) And (Token = "+" Or Token = "-")) Then
        op = Token
        GetToken
    End If
    
    Exp5 = Exp6()
    If (op = "-") Then Exp5 = -CDbl(Exp5)
End Function

Function Exp6()
    'Check for Parenthesized expression
    If Token = "(" Then
        GetToken 'Get next token
        Exp6 = Exp1()
        'Check that we have a closeing bracket
        If (Token <> ")") Then abort 2
        GetToken 'Get next token
    Else
        Exp6 = atom()
    End If
    
End Function

Function atom()
Dim Temp As String

    'Check for Digits ,Hexadecimal,Functions, Variables
    Select Case tok_type
        Case HEXDIGIT 'Hexadecimal
            Temp = Trim(Right(Token, Len(Token) - 2))
            If Len(Temp) = 0 Then
                abort 6
            ElseIf Not isHex(Temp) Then
                abort 7, Temp
            Else
                atom = CDec("&H" & Temp)
                GetToken
            End If
            
        Case IDENTIFIER 'Inbuilt Functions
           atom = CallIntFunc(Token)
            GetToken
        Case DIGIT 'Digit const found
            If Not IsNumeric(Token) Then abort 3, Token 'Check we have a real digit
            atom = Token 'Return the value
            GetToken 'Get next token
        Case LERROR 'Expression phase error
            abort 0, Token 'Show error message
        Case VARIABLE 'Variable found
            If FindVarIdx(Token) = -1 Then abort 0, Token
            atom = GetVarData(Token) 'Return variable value
            GetToken 'Get next token
    End Select
    
End Function

Function CallIntFunc(sFunction As String) As Double
Dim Temp
Dim UserFuncID As Integer, x As Integer, sFunction_Str As String
Dim ArgList

'ABS,ATN,COS,EXP,LOG,RND,ROUND,SGN,SIN,SQR,TAN

    On Error Resume Next
    Select Case UCase(sFunction)
        Case "ABS"
            GetToken
            Temp = Exp6
            CallIntFunc = Abs(Temp)
            PushBack
        Case "ATN"
            GetToken
            Temp = Exp6
            CallIntFunc = Atn(Temp)
            PushBack
        Case "COS"
            GetToken
            Temp = Exp6
            CallIntFunc = Cos(Temp)
            PushBack
        Case "EXP"
            GetToken
            Temp = Exp6
            CallIntFunc = Exp(Temp)
            PushBack
        Case "LOG"
            GetToken
            Temp = Exp6
            CallIntFunc = Log(Temp)
            PushBack
        Case "RND"
            GetToken
            Temp = Exp6
            CallIntFunc = Rnd(Temp)
            PushBack
        Case "ROUND"
            GetToken
            Temp = Exp6
            CallIntFunc = Round(Temp)
            PushBack
        Case "SGN"
            GetToken
            Temp = Exp6
            CallIntFunc = Sgn(Temp)
            PushBack
        Case "SIN"
            GetToken
            Temp = Exp6
            CallIntFunc = Sin(Temp)
            PushBack
        Case "SQR"
            GetToken
            Temp = Exp6
            CallIntFunc = Sqr(Temp)
            PushBack
        Case "TAN"
            GetToken
            Temp = Exp6
            CallIntFunc = Tan(Temp)
            PushBack
        Case "SUM"
            ArgList = GetArgs
            Temp = 0
            For x = 0 To UBound(ArgList)
                Temp = CDbl(Temp) + CDbl(ArgList(x))
            Next x
            
            GetToken
            CallIntFunc = Temp
            PushBack
    End Select
    
End Function

Function GetArgs()
Dim Count As Integer
Dim Value
Dim Temp() As Variant

    GetToken
    If Token <> "(" Then Exit Function
    
    Do
        
        GetToken
        Value = Exp1
        ReDim Preserve Temp(0 To Count)
        Temp(Count) = Value
        Count = Count + 1

    Loop Until (Token = ")")
    
    GetArgs = Temp
    Erase Temp
    Count = 0
    Value = 0
    
    
End Function

Sub PutHead(head As String)
    Me.FontBold = True
    Print head
    Me.FontBold = False
End Sub

Private Sub cmdexit_Click()
    End
End Sub

Private Sub cmdTest_Click()
    'Just some tests
    Cls
    'Me.Caption = Eval("sum(log(2),5) * sum(10,2) \ sum(10,1)") Now it's working
    PutHead "Normal VB expresion tests"
    Print ""
    Print "2 + 2 * (5 + 5) = " & 2 + 2 * (5 + 5)
    Print "area + area = " & 180 + 180
    Print "8+1/7*4+(9*4+1*(2+8))*6 = " & 8 + 1 / 7 * 4 + (9 * 4 + 1 * (2 + 8)) * 6
    Print "pi =" & 22 / 7
    Print "2 ^ 2 = " & 2 ^ 2
    Print "15 MOD 2 = " & 15 Mod 2
    Print String(160, "-")
    PutHead "Normal VB Relational tests"
    Print "5 > 5 " & (5 > 5)
    Print "4 > 1 " & (4 > 2)
    Print "8 = 5 " & (8 = 5)
    Print "5 = 5 " & (5 = 5)
    Print "16 >=16 " & (16 >= 16)
    Print "6 <= 0 " & (6 <= 0)
    Print "5 <> 6 " & (5 <> 6)
    Print "5 <> 5 " & (5 <> 5)
    Print String(160, "-")
    PutHead "Eval tests expected as tests above"
    'Add the area variable
    Call Eval("area = 180")
    Print "2 + 2 * (5 + 5) = " & Eval("2 + 2 * (5 + 5)")
    Print "area + area = " & Eval("area + area")
    Print "8+1/7*4+(9*4+1*(2+8))*6 = " & Eval("8+1/7*4+(9*4+1*(2+8))*6")
    Print "pi =" & Eval("22 / 7")
    Print "2 ^ 2 = " & Eval("2 ^ 2")
    Print "MOD Test %"
    Print "15 % 2 = " & Eval("15 % 2")
    Print String(160, "-")
    
    PutHead "Eval Relational tests"
    Print "5 > 5 " & CBool(Eval("5 > 5"))
    Print "4 > 1 " & CBool(Eval("4 > 2"))
    Print "8 = 5 " & CBool(Eval("8 = 5"))
    Print "5 = 5 " & CBool(Eval("5 = 5"))
    Print "16 >=16 " & CBool(Eval("16 >= 16"))
    Print "6 <= 0 " & CBool(Eval("6 <= 0"))
    Print "5 <> 6 " & CBool(Eval("5 <> 6"))
    Print "5 <> 5 " & CBool(Eval("5 <> 5"))
    Print "(10 = 10) " & CBool(Eval("(10 = 10)"))
    Print String(160, "-")
    'Now better variable support
    PutHead "Now with better variable support"
    Print "Call Eval(a = 5)"
    Call Eval("a=5")
    Print "And here is the result of the above = " & Eval("a")
    Print String(160, "-")
    Print "Random test (-5 + 3) = " & Eval("(-5 + 3)")
    Print "This should be 5 (-5 + 3) / -5 * (-2.5) + 6  and was it ?  " & Eval("(-5 + 3) / -5 * (-2.5) + 6")
    Print "So were the result right? " & " Let me know"

End Sub

Private Sub Command1_Click()
    Cls
    Print String(160, "-")
    PutHead "VB Function Tests"
    Print String(160, "-")
    'ABS,ATN,COS,EXP,LOG,RND,ROUND,SGN,SIN,SQR,TAN,SUM"
    Print "Abs(-6) = " & Abs(-6)
    Print "ATN(5) = " & Atn(5)
    Print "COS(180) = " & Cos(180)
    Print "EXP(6) = " & Exp(8)
    Print "LOG(1000)\LOG(10) = " & Log(1000) \ Log(10)
    Print "Rnd(8) = " & Rnd(8)
    Print "SQR(9) + SQR(9) = " & Sqr(9) + Sqr(9)
    Print String(160, "-")
    PutHead "Our Eval Engine Tests"
    Print String(160, "-")
    
    Print "Eval Abs(-6) = " & Eval("Abs(-6)")
    Print "Eval ATN(5) = " & Eval("Atn(5)")
    Print "Eval COS(180) = " & Eval("Cos(180)")
    Print "Eval EXP(6) = " & Eval("Exp(8)")
    Print "Eval LOG(1000)\LOG(10) = " & Eval("Log(1000) \ Log(10)")
    Print "Eval Rnd(8) = " & Eval("Rnd(8)")
    Print "Eval SQR(9) + SQR(9) = " & Eval("Sqr(9) + Sqr(9)")
    Print "Eval Function with Arglist"
    
    Print "VB test of above function first"
    Print "(5 + 5) * (10 + 2) \ (10 + 2) + 9 = " & (5 + 5) * (10 + 2) \ (10 + 2) + 9
    Print ""
    Print "Eval Test"
    Print "Eval(sum(5,5) * sum(10,2) \ sum(10 + 2) + 9) = " & Eval("sum(5,5) * sum(10,2) \ sum(10 + 2) + 9")
    
    Print "Hexadecimal Tests"
    
    Print "VB Tests"
    Print "&H256 = " & &H256
    Print "Eval(0x256) = " & Eval("0x256")
    Print "Eval(0x256 + 0xffffff) = " & Eval("0x256 + 0xffffff")
    Print String(160, "-")
    
End Sub

Private Sub Command2_Click()
    'Bitwise tests
    Cls
    Print String(160, "-")
    PutHead "VB bitwise Tests and other tests"
    Print String(160, "-")
    Print "5 And 6 = " & (5 And 6)
    Print "1 Or 8 = " & (1 Or 8)
    Print "16 Xor 8 = " & (16 Xor 8)
    Print "15 Mod 4 = " & (15 Mod 4)
    Print "8 Imp 4 = " & (8 Imp 4)
    Print "3 Eqv 3 = " & (3 Eqv 3)
    Print "22 \ 7 = " & (22 \ 7)
    Print String(160, "-")
    PutHead "Now Tests above using our Eval"
    Print String(160, "-")
    Print "5 And 6 = " & Eval("5 And 6") & vbTab & "5 && 6 = " & Eval("5 && 6")
    Print "1 Or 8 = " & Eval("1 or 8") & vbTab & "1 || 8 = " & Eval("1 || 8")
    Print "16 Xor 8 = " & Eval("16 Xor 8")
    Print "15 Mod 4 = " & Eval("15 Mod 4")
    Print "8 Imp 4 = " & Eval("8 Imp 4")
    Print "3 Eqv 3 = " & Eval("3 eqv 3")
    Print "27 \ 7 = " & Eval("27 \ 7") & vbTab & "27 Div 7 = " & Eval("27 Div 7")
    Print String(160, "-")
    PutHead "Our Bit Shift Tests"
    Print String(160, "-")
    Print "Shift Left (2 << 4)" & " = " & Eval("(2 << 4)")
    Print "Shift Right (32 >> 4)" & " = " & Eval("(32 >> 4)")
    Print ""
    Print "Shift Left (2 shl 4)" & " = " & Eval("(2 shl 4)")
    Print "Shift right (32 shr 4)" & " = " & Eval("(32 shr 4)")
    Print String(160, "-")
    PutHead "Increment/Decrement operators -- ++"
    Print String(160, "-")
    Print "x = 1"
    Call Eval("x = 1")
    Print "x++"
    Call Eval("x++")
    Print "Value of a is now " & Eval("x")
    Print "x--"
    Call Eval("x--")
    Print "Value of a is now " & Eval("x")
    Print String(160, "-")
    PutHead "O almost forgot the NOT operator"
    Call Eval("test = 5")
    Print "test = 5"
    Print "!5 " & Eval("!5")
    Call Eval("b = Not 5")
    Print "Eval(""b = Not 5"")"
    Print "b = " & Eval("b")
    Print "Well that is a fair amount of operators ant it :)"
    
End Sub

Private Sub Form_Load()
    'Add some variables
    Call init
    AddVar "pi", 3.14159265358979
    AddVar "e", 2.71828182845905
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
End Sub

