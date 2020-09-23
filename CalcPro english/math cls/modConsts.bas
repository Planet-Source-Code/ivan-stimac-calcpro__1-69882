Attribute VB_Name = "modConsts"
'
'consts
'ascii
Public Const ASC_DOLLAR  As Byte = 36    '$
Public Const ASC_DOWNLN As Byte = 95    '_
Public Const ASC_MINUS As Byte = 45    '-
Public Const ASC_PLUS As Byte = 43    '+
'
'provjere dali treba nešto raditi (da li ima operatora)
Public Const CHK_CALLFUNC   As String = "+-*/%^<=>=!=&&||"  'zvati ce se Calculate ako u provjeravanom stringu ima nešto od ovoga
'
Public Const LNG_DOLLAR     As String = "$"
Public Const LNG_OPENST     As String = "("
Public Const LNG_CLOSEST    As String = ")"
Public Const LNG_REPLACEB   As String = "\"
Public Const LNG_DOWNLN     As String = "_"
'
Public Const LNG_ARGDELIMITER As String = ";"
Public Const LNG_SYSVAR     As String = "_sysVar_"
'Public Const LNG_REPL       As String = "$pre\_sysVar_$num_\$aft"
'
'lib code
Public Const LIB_COMMENT    As String = "//"
'operators
Public Const OP_PLUS        As String = "+"
Public Const OP_MINUS       As String = "-"
Public Const OP_DIV         As String = "/"
Public Const OP_MUL         As String = "*"
Public Const OP_POT         As String = "^"
'LOGIC
Public Const OP_AND         As String = "&&"
Public Const OP_OR          As String = "||"
Public Const OP_IS          As String = "=="
Public Const OP_ISNOT       As String = "!="
Public Const OP_LESSTHAN    As String = "<"
Public Const OP_GREATERTHAN As String = ">"
Public Const OP_LESSORIS    As String = "<="
Public Const OP_GREATERORIS As String = ">="
'

'
'Complex
Public Const CPL_IMAG       As String = "i"

'math consts
Public Const PI             As Double = 3.14159265358979

'lenghts
Public Const LEN_FUNCTION   As Integer = 8   'LEN_FUNCTION
Public Const LEN_DEFINE     As Integer = 6   'LEN_DEFINE
Public Const LEN_ENDFUNCT   As Integer = 12

'errors
'
Public Const ERR_FunctionExist      As String = "Invalid function '$fname', function already exist!"
Public Const ERR_FunctionExistN     As Integer = -11
'
Public Const ERR_InvalidVarName     As String = "Invalid variable name"
Public Const ERR_InvalidVarNameN    As Integer = -11
'
'
Public Const ERR_InvalidLogBase     As String = "Invalid logN base"
Public Const ERR_InvalidLogBaseN    As Integer = -11
'
Public Const ERR_InvalidChar        As String = "Statment can't contain any of the folowing characters: \ : _"
Public Const ERR_InvalidCharN       As Integer = -11
'
Public Const ERR_InvalidFunction    As String = "Invalid function or missing operator"
Public Const ERR_InvalidFunctionN   As Integer = -11
'
Public Const ERR_FunctionErr        As String = "Function error (check domain): "
Public Const ERR_FunctionErrN       As Integer = -11
'
Public Const ERR_BesselErr          As String = "Bessel function error: "
Public Const ERR_BesselErrN         As Integer = -11
'
'
Public Const ERR_BetaIncErr         As String = "BetaInc function error:"
Public Const ERR_BetaIncErrN        As Integer = -11
'
Public Const ERR_SphericalErr       As String = "Spherical function error: "
Public Const ERR_SphericalErrN      As Integer = -11
'
Public Const ERR_BetaErr            As String = "Beta function error: "
Public Const ERR_BetaErrN           As Integer = -11
'
Public Const ERR_AiryErr            As String = "Airy function error: "
Public Const ERR_AiryErrN           As Integer = -11
'
Public Const ERR_ExpEndOfLn         As String = "Expected ';'"
Public Const ERR_ExpEndOfLnN        As Integer = -11
'
Public Const ERR_ExpEndOfSt         As String = "Expected end of statment"
Public Const ERR_ExpEndOfStN        As Integer = -11
'
Public Const ERR_ElseIfBefIf        As String = "Elseif before block if"
Public Const ERR_ElseIfBefIfN       As Integer = -11
'
Public Const ERR_ElseBefIf          As String = "Else before block if"
Public Const ERR_ElseBefIfN         As Integer = -11
'
Public Const ERR_EndIfBefIf         As String = "End if before block if"
Public Const ERR_EndIfBefIfN        As Integer = -11
'
Public Const ERR_ExpectedEndIf      As String = "Block if without end if"
Public Const ERR_ExpectedEndIfN     As Integer = -11
'
Public Const ERR_ExpectedLoop       As String = "While without loop"
Public Const ERR_ExpectedLoopN      As Integer = -11
'
Public Const ERR_LoopBefWhile       As String = "Loop before block while"
Public Const ERR_LoopBefWhileN      As Integer = -11
'
Public Const ERR_BreakWthWhile      As String = "Break without block while"
Public Const ERR_BreakWthWhileN     As Integer = -11
'
Public Const ERR_InvalidChars       As String = "There is some invalid chars or expressions. You should not use: \, :, ? "
Public Const ERR_InvalidCharsN      As Integer = -11
'
Public Const ERR_ExpectedVarNm      As String = "Expected variable name"
Public Const ERR_ExpectedVarNmN     As Integer = -11
'
Public Const ERR_InvalidVarNm       As String = "Invalid variable name"
Public Const ERR_InvalidVarNmN      As Integer = -11
'
Public Const ERR_ExpectedExpres     As String = "Expected expression"
Public Const ERR_ExpectedExpresN    As Integer = -11
'
Public Const ERR_CodeOutside        As String = "Only comments may appear outside of function"
Public Const ERR_CodeOutsideN       As Integer = -11
'
Public Const ERR_UndefVar           As String = "Undefined variable "
Public Const ERR_UndefVarN          As Integer = -11
'
Public Const ERR_DivByZer           As String = "Division by zero "
Public Const ERR_DivByZerN          As Integer = -11
'
Public Const ERR_InfRes             As String = "Infinite"
Public Const ERR_InfResN            As Integer = -11
'
Public Const ERR_NegArg             As String = "Argument must be nonnegative"
Public Const ERR_NegArgN            As Integer = -11
'
Public Const ERR_NonIntArg          As String = "Argument must be integer"
Public Const ERR_NonIntArgN         As Integer = -11
'
Public Const ERR_ExpectArgs         As String = "Expected arguments "
Public Const ERR_ExpectArgsN        As Integer = -11
'
Public Const ERR_InvalidArg         As String = "Invalid argument"
Public Const ERR_InvalidArgN        As Integer = -11
'
Public Const ERR_UnknownCmd         As String = "Unknown command"
Public Const ERR_UnknownCmdN        As Integer = -11
'
Public Const ERR_ExpectedStatm      As String = "Expected statment"
Public Const ERR_ExpectedStatmN     As Integer = -11
'
Public Const ERR_ExpectedFunc       As String = "Expected function"
Public Const ERR_ExpectedFuncN      As Integer = -11
'
Public Const ERR_DeadCode           As String = "Dead code detected"
Public Const ERR_DeadCodeN          As Integer = -11
'
Public Const ERR_UndefFunc          As String = "Undefined function"
Public Const ERR_UndefFuncN         As Integer = -11
'
Public Const ERR_FileError          As String = "Unable to load file"
Public Const ERR_FileErrorN         As Integer = -11
'
'binomCDF
Public Const ERR_BinomNegPar        As String = "Binom error: first and second argument must be positive"
Public Const ERR_BinomNegParN       As Integer = -11
'
Public Const ERR_BinomArgErr        As String = "Binom error: first argument must be <= second argument"
Public Const ERR_BinomArgErrN       As Integer = -11
'
Public Const ERR_BinomLastErr       As String = "Binom error: last argument must be from interval [0, 1]"
Public Const ERR_BinomLastErrN      As Integer = -11
'
'HypergeometricPMF
Public Const ERR_hiperGeomPMF       As String = "HypergeometricPMF error: ivalid domain"
Public Const ERR_hiperGeomPMFN      As Integer = -11
'
'GeometricPMF
Public Const ERR_geometricFirst     As String = "Geometric error: first argumet must be from interval [0, 1]"
Public Const ERR_geometricFirstN    As Integer = -11
'
Public Const ERR_geometricLast      As String = "Geometric error: second argument must >= 1"
Public Const ERR_geometricLastN     As Integer = -11
'
'ChiSquarePDF
Public Const ERR_ChiSquareNeg       As String = "ChiSquare error: arguments must be positive"
Public Const ERR_ChiSquareNegN      As Integer = -11
'
Public Const ERR_ChiSquareSecArg    As String = "ChiSquare error: second argument must be > 0"
Public Const ERR_ChiSquareSecArgN   As Integer = -11
'
'ExponentialPDF
Public Const ERR_ExpNeg             As String = "Exponential error: arguments must be positive"
Public Const ERR_ExpNegN            As Integer = -11
'
'Fisher
Public Const ERR_FisherSecLastArg   As String = "Fisher error: arguments must be > 0"
Public Const ERR_FisherSecLastArgN  As Integer = -11
'normal
Public Const ERR_NormlaLastArg      As String = "Normal distribution error: last argument must be > 0"
Public Const ERR_NormlaLastArgN     As Integer = -11
''
Public Const ERR_PoissonNeg         As String = "Poisson error: arguments must be positive"
Public Const ERR_PoissonNegN        As Integer = -11
'
''
Public Const ERR_SumArg             As String = "Sum error: upper limit must be >= lower limit"
Public Const ERR_SumArgN            As Integer = -11

'
'pretvorbe
'hex
Public Const ERR_InvalidHexVal      As String = "Hex value found with characters other than 0-9, a-f, or A-F!"
Public Const ERR_InvalidHexValN     As Integer = -11
'
Public Const ERR_InvalidHexLen      As String = "Hex value length must be 1-8 characters!"
Public Const ERR_InvalidHexLenN     As Integer = -11
'
'oct
Public Const ERR_InvalidOctVal      As String = "Oct value found with characters other than 0-7!"
Public Const ERR_InvalidOctValN     As Integer = -11
'
Public Const ERR_InvalidOctLen      As String = "Oct value length must be 1-10 characters!"
Public Const ERR_InvalidOctLenN     As Integer = -11
'
'bin
Public Const ERR_InvalidBinVal      As String = "Bin value found with characters other than 0, 1!"
Public Const ERR_InvalidBinValN     As Integer = -11
'
Public Const ERR_InvalidBinLen      As String = "Bin value length must be 1-10 characters!"
Public Const ERR_InvalidBinLenN     As Integer = -11
'
'
Public Const ERR_NonIntVal          As String = ": argument must be integer!"
Public Const ERR_NonIntValN         As Integer = -11
'
Public Const ERR_OutsideInterval    As String = ": argument must be from interval [i]!"
Public Const ERR_OutsideIntervalN   As Integer = -11
