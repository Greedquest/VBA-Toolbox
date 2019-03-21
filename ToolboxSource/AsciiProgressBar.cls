VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AsciiProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@PredeclaredId
Option Explicit
'@Folder(AsciiProgressBar)

Private Type tProgressBar
    percentComplete As Double
    size As Long
    base As String
    bar As String
    character As String
    whitespace As String
    mask As String
End Type

Private Enum progressError
    percentOutOfBoundsError = vbObjectError + 513 'to get into custom error raising territory
    barSizeOutOfRangeError
    singleCharacterRequiredError
    baseIsNotAStringError
    maskMissingPositionalArgumentError
End Enum

Private Const DEFAULT_CHAR As String = "|"
Private Const DEFAULT_SIZE As Long = 10
Private Const DEFAULT_BASE As String = vbNullString
Private Const DEFAULT_WHITESPACE As String = " "
Private Const DEFAULT_MASK As String = "{0}{1}{2}%"

Private this As tProgressBar

Public Function Update(ByVal fractionComplete As Double) As String
    
    'check if valid input (0-100%)
    If fractionComplete < 0# Or fractionComplete > 1# Then raiseError percentOutOfBoundsError
    
    'set number of charachters in progress bar
    this.percentComplete = fractionComplete
    Dim numberOfChars As Long
    numberOfChars = Round(this.size * this.percentComplete, 0)
    this.bar = String(numberOfChars, this.character) & String(this.size - numberOfChars, this.whitespace)
    
    Update = repr
End Function

Public Property Get repr() As String
    repr = Toolbox.Strings.Format(this.mask, this.base, this.bar, Round(this.percentComplete * 100, 0))
End Property

Private Sub raiseError(ByVal errNum As progressError, ParamArray args() As Variant)
    Select Case errNum
    
    Case percentOutOfBoundsError
        Err.Description = "Percent must lie between 0.0 and 1.0"
        
    Case barSizeOutOfRangeError
        Err.Description = Toolbox.Strings.printf("Bar size must be at least {0} characters", args(0))
        
    Case singleCharacterRequiredError
        Err.Description = Toolbox.Strings.printf("Only a single character should be used as {0}, not '{1}'", args(0), args(1))
        
    Case baseIsNotAStringError
        Err.Description = Toolbox.Strings.printf("Base must be of type string or left blank, not '{0}'", TypeName(args(0)))
        
    Case maskMissingPositionalArgumentError
        Err.Description = Toolbox.Strings.printf("formatMask must contain all three positional tokens ({0,1,2}){0}'{1}' does not", _
                                                 vbCrLf, args(0))
                                                 
    Case Else                                    'some errNum we don't know what to do with
        On Error Resume Next                     'fake raise to grab description text
        Err.Raise errNum
        Dim errDescription As String
        errDescription = Err.Description
        On Error GoTo 0
Debug.Print Toolbox.Strings.printf("Warning: Unexpected error '{0}' with description '{1}'", errNum, errDescription)
    End Select
    Err.Raise errNum
End Sub

Public Sub Init(Optional ByVal size As Long = 0, Optional ByVal base As Variant, _
                Optional ByVal character As String = vbNullString, Optional ByVal whitespace As String = vbNullString, _
                Optional ByVal formatMask As String = vbNullString)
    'Method to set appearence and other properties of the progress bar

    'check is inputs were missing - if so leave as they were
    'Base can be any string so can't be checked in this way, needs special handling
    size = IIf(size = 0, this.size, size)
    character = IIf(character = vbNullString, this.character, character)
    whitespace = IIf(whitespace = vbNullString, this.whitespace, whitespace)
    formatMask = IIf(formatMask = vbNullString, this.mask, formatMask)
    
    'check for valid inputs
    Const minBarSize As Long = 2
    If size < minBarSize Then
        raiseError barSizeOutOfRangeError, minBarSize
    ElseIf Len(character) <> 1 Then
        raiseError singleCharacterRequiredError, "'character'", character
    ElseIf Len(whitespace) <> 1 Then
        raiseError singleCharacterRequiredError, "'whitespace'", whitespace
    ElseIf MaskIsInvalid(formatMask) Then
        raiseError maskMissingPositionalArgumentError, formatMask
    ElseIf Not IsMissing(base) Then
        'base is variant so requires type checking
        On Error Resume Next
        this.base = base                         'may be type error if base can't be converted; e.g an object was passed
        Dim errNum As Long
        errNum = Err.Number
        On Error GoTo 0
        If errNum <> 0 Then
            raiseError baseIsNotAStringError, base
        End If
    End If

    'If we've got here then inputs are valid, so we can commit them
    this.size = size
    this.whitespace = whitespace
    this.character = character
    this.mask = formatMask
    
End Sub

Private Function MaskIsInvalid(ByVal mask As String) As Boolean
    'check whether any of the positional tokens don't appear in the mask
    Const matchPattern As String = "{0} {1} {2}"
    Dim tokens() As String
    tokens = Split(matchPattern)
    MaskIsInvalid = False
    Dim token As Variant
    For Each token In tokens
        MaskIsInvalid = Not CBool(InStr(mask, token))
        If MaskIsInvalid Then Exit Function
    Next
End Function

Private Sub Class_Initialize()
    ResetDefaults
    Update this.percentComplete
End Sub

Public Sub ResetDefaults()
    this.character = DEFAULT_CHAR
    this.base = DEFAULT_BASE
    this.whitespace = DEFAULT_WHITESPACE
    this.size = DEFAULT_SIZE
    this.mask = DEFAULT_MASK
End Sub

Public Function Create(Optional ByVal size As Long = 0, Optional ByVal base As Variant, _
                       Optional ByVal character As String = vbNullString, Optional ByVal whitespace As String = vbNullString, _
                       Optional ByVal formatMask As String = vbNullString) As AsciiProgressBar
                
    Dim result As New AsciiProgressBar
    result.Init size, base, character, whitespace, formatMask
    Set Create = result

End Function


