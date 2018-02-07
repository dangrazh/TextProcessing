VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPorterStemEn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Option Base 1
'Option Compare Text

'# <author> Daniel Grass based on the work of Navonil Mustafee
'# <mail> dani.grass@bluewin.ch

'#Region
'#======================================================================================================================
'# About this class
'#======================================================================================================================
'Porter Stemmer in VISUAL BASIC 6. It follow the algorithm definition
'presented in :
'   Porter, 1980, An algorithm for suffix stripping, Program, Vol. 14,
'   no. 3, pp 130-137,
'   (http://www.tartarus.org/~martin/PorterStemmer)

'Author : Navonil Mustafee
'Brunel University - student
'Algorithm Implemented as part for assignment on document visualization

'TO USE THE PROGRAM CALL THE FUNCTION PORTERALGORITHM. THE WORD
'TO BE STEMMED SHOULD BE PASSED AS THE ARGUEMENT ARGUEMENT. THE STRING
'RETURNED BY THE FUNCTION IS THE STEMMED WORD


'#======================================================================================================================
'# Accessible in this class
'#======================================================================================================================
'     |> Get | --- About                :: Returns description of the class.
'     |> Get | --- Name                 :: Returns name of the class.
'     |> Get | --- Version              :: Returns version string for the class [e.g. #.# (year)].


'#======================================================================================================================
'# References
'#======================================================================================================================
'Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal bytes As Long)

'#======================================================================================================================
'# Dependencies to other classes
'#======================================================================================================================
'none

'#======================================================================================================================
'# Application Constants, Enumerations & Types
'#======================================================================================================================
Private Const C_Name As String = "cPorterStemEn.cls"


'#======================================================================================================================
'# Private Variables
'#======================================================================================================================
Private step1a(3, 1) 	As String		'mappings used in step 1a
Private step2(20, 1) 	As String		'mappings used in step 2
Private step3(6, 1) 	As String		'mappings used in step 3
Private step4(18) 		As String		'mappings used in step 4




'#Region
'#======================================================================================================================
'# Class Initialization, Termination & Properties
'#======================================================================================================================


Private Sub Class_Initialize()
' ************************************************
' Class constructor.
' ************************************************

    'Debug.Print "|> Initializing:= " & Me.Name
    
	'initializing contents of 2D array - step 1
	step1a(0, 0) = "sses"
	step1a(0, 1) = "ss"
	step1a(1, 0) = "ies"
	step1a(1, 1) = "i"
	step1a(2, 0) = "ss"
	step1a(2, 1) = "ss"
	step1a(3, 0) = "s"
	step1a(3, 1) = ""
	
	
	'initializing contents of 2D array - step 2
	step2(0, 0) = "ational"
	step2(0, 1) = "ate"
	step2(1, 0) = "tional"
	step2(1, 1) = "tion"
	step2(2, 0) = "enci"
	step2(2, 1) = "ence"
	step2(3, 0) = "anci"
	step2(3, 1) = "ance"
	step2(4, 0) = "izer"
	step2(4, 1) = "ize"
	step2(5, 0) = "bli"
	step2(5, 1) = "ble"
	step2(6, 0) = "alli"
	step2(6, 1) = "al"
	step2(7, 0) = "entli"
	step2(7, 1) = "ent"
	step2(8, 0) = "eli"
	step2(8, 1) = "e"
	step2(9, 0) = "ousli"
	step2(9, 1) = "ous"
	step2(10, 0) = "ization"
	step2(10, 1) = "ize"
	step2(11, 0) = "ation"
	step2(11, 1) = "ate"
	step2(12, 0) = "ator"
	step2(12, 1) = "ate"
	step2(13, 0) = "alism"
	step2(13, 1) = "al"
	step2(14, 0) = "iveness"
	step2(14, 1) = "ive"
	step2(15, 0) = "fulness"
	step2(15, 1) = "ful"
	step2(16, 0) = "ousness"
	step2(16, 1) = "ous"
	step2(17, 0) = "aliti"
	step2(17, 1) = "al"
	step2(18, 0) = "iviti"
	step2(18, 1) = "ive"
	step2(19, 0) = "biliti"
	step2(19, 1) = "ble"
	step2(20, 0) = "logi"
	step2(20, 1) = "log"
	
	'initializing contents of 2D array - step 3
	step3(0, 0) = "icate"
	step3(0, 1) = "ic"
	step3(1, 0) = "ative"
	step3(1, 1) = ""
	step3(2, 0) = "alize"
	step3(2, 1) = "al"
	step3(3, 0) = "iciti"
	step3(3, 1) = "ic"
	step3(4, 0) = "ical"
	step3(4, 1) = "ic"
	step3(5, 0) = "ful"
	step3(5, 1) = ""
	step3(6, 0) = "ness"
	step3(6, 1) = ""
	
	'initializing contents of 2D array - step 4
	step4(0) = "al"
	step4(1) = "ance"
	step4(2) = "ence"
	step4(3) = "er"
	step4(4) = "ic"
	step4(5) = "able"
	step4(6) = "ible"
	step4(7) = "ant"
	step4(8) = "ement"
	step4(9) = "ment"
	step4(10) = "ent"
	step4(11) = "ion"
	step4(12) = "ou"
	step4(13) = "ism"
	step4(14) = "ate"
	step4(15) = "iti"
	step4(16) = "ous"
	step4(17) = "ive"
	step4(18) = "ize"	

End Sub


Private Sub Class_Terminate()
' ************************************************
' Class destructor.
' ************************************************
    
    'Debug.Print "|> Terminating:= " & Me.Name

End Sub

Public Property Get Version() As String
' ************************************************
' Version string of the current class.
' Contains a list of (historical) changes to the class within the comments of the procedure.
' ************************************************
   Version = "Version 1.0 (2017)"  'Initial release

End Property

Public Property Get About() As String
' ***********************************************
' String that describes the current class.
' ***********************************************
    About = "Implements the classic Porter Stemming algorithm for english text. Version: " & Me.Version & "." & VBA.vbCrLf & VBA.vbCrLf
    About = About & "For additional details please contact the author."
End Property

Public Property Get Name() As String
' ***********************************************
' Returns the name of the class.
' ***********************************************
    Name = C_Name
    
End Property


Public Function StemWord(str As String) As String
On Error Resume Next

	'only strings greater than 2 are stemmed
	If Len(Trim(str)) > 2 Then
			str = porterAlgorithmStep1(str)
			str = porterAlgorithmStep2(str)
			str = porterAlgorithmStep3(str)
			str = porterAlgorithmStep4(str)
			str = porterAlgorithmStep5(str)
	End If

	'End of Porter's algorithm.........returning the word
	StemWord = str

End Function

Private Function porterAlgorithmStep1(str As String) As String
On Error Resume Next

	'STEP 1A
	'
	'    SSES -> SS                         caresses  ->  caress
	'    IES  -> I                          ponies    ->  poni
	'                                       ties      ->  ti
	'    SS   -> SS                         caress    ->  caress
	'    S    ->                            cats      ->  cat


	'declaring local variables
	Dim i As Byte
	Dim j As Byte

	'checking word
	For i = 0 To 3 Step 1
			If porterEndsWith(str, step1a(i, 0)) Then
							str = porterTrimEnd(str, Len(step1a(i, 0)))
							str = porterAppendEnd(str, step1a(i, 1))
							Exit For
			End If
	Next i


	'--------------------------------------------------------------------------------------------------------

	'STEP 1B
	'
	'   If
	'       (m>0) EED -> EE                     feed      ->  feed
	'                                           agreed    ->  agree
	'   Else
	'       (*v*) ED  ->                        plastered ->  plaster
	'                                           bled      ->  bled
	'       (*v*) ING ->                        motoring  ->  motor
	'                                           sing      ->  sing
	'
	'If the second or third of the rules in Step 1b is successful, the following
	'is done:
	'
	'    AT -> ATE                       conflat(ed)  ->  conflate
	'    BL -> BLE                       troubl(ed)   ->  trouble
	'    IZ -> IZE                       siz(ed)      ->  size
	'    (*d and not (*L or *S or *Z))
	'       -> single letter
	'                                    hopp(ing)    ->  hop
	'                                    tann(ed)     ->  tan
	'                                    fall(ing)    ->  fall
	'                                    hiss(ing)    ->  hiss
	'                                    fizz(ed)     ->  fizz
	'    (m=1 and *o) -> E               fail(ing)    ->  fail
	'                                    fil(ing)     ->  file
	'
	'The rule to map to a single letter causes the removal of one of the double
	'letter pair. The -E is put back on -AT, -BL and -IZ, so that the suffixes
	'-ATE, -BLE and -IZE can be recognised later. This E may be removed in step
	'4.

	'declaring local variables
	Dim m As Byte
	Dim temp As String
	Dim second_third_success As Boolean

	'initializing contents of 2D array
	second_third_success = False

	'(m>0) EED -> EE..else..(*v*) ED  ->(*v*) ING  ->
	If porterEndsWith(str, "eed") Then

			'counting the number of m's
			temp = porterTrimEnd(str, Len("eed"))
			m = porterCountm(temp)

			If m > 0 Then
							str = porterTrimEnd(str, Len("eed"))
							str = porterAppendEnd(str, "ee")
			End If

	ElseIf porterEndsWith(str, "ed") Then
	  
					'trim and check for vowel
					temp = porterTrimEnd(str, Len("ed"))

					If porterContainsVowel(temp) Then
							str = porterTrimEnd(str, Len("ed"))
							second_third_success = True
					End If
					
	ElseIf porterEndsWith(str, "ing") Then

					'trim and check for vowel
					temp = porterTrimEnd(str, Len("ing"))
					
					If porterContainsVowel(temp) Then
							str = porterTrimEnd(str, Len("ing"))
							second_third_success = True
					End If

	End If

	'If the second or third of the rules in Step 1b is SUCCESSFUL, the following
	'is done:
	'
	'    AT -> ATE                       conflat(ed)  ->  conflate
	'    BL -> BLE                       troubl(ed)   ->  trouble
	'    IZ -> IZE                       siz(ed)      ->  size
	'    (*d and not (*L or *S or *Z))
	'       -> single letter
	'                                    hopp(ing)    ->  hop
	'                                    tann(ed)     ->  tan
	'                                    fall(ing)    ->  fall
	'                                    hiss(ing)    ->  hiss
	'                                    fizz(ed)     ->  fizz
	'    (m=1 and *o) -> E               fail(ing)    ->  fail
	'                                    fil(ing)     ->  file


	If second_third_success = True Then             'If the second or third of the rules in Step 1b is SUCCESSFUL
					
			If porterEndsWith(str, "at") Then           'AT -> ATE
							str = porterTrimEnd(str, Len("at"))
							str = porterAppendEnd(str, "ate")
			ElseIf porterEndsWith(str, "bl") Then       'BL -> BLE
							str = porterTrimEnd(str, Len("bl"))
							str = porterAppendEnd(str, "ble")
			ElseIf porterEndsWith(str, "iz") Then       'IZ -> IZE
							str = porterTrimEnd(str, Len("iz"))
							str = porterAppendEnd(str, "ize")
			ElseIf porterEndsDoubleConsonent(str) Then  '(*d and not (*L or *S or *Z))-> single letter
							If Not (porterEndsWith(str, "l") Or porterEndsWith(str, "s") Or porterEndsWith(str, "z")) Then
									str = porterTrimEnd(str, 1)
							End If
			ElseIf porterCountm(str) = 1 Then                           '(m=1 and *o) -> E
							If porterEndsCVC(str) Then
									   str = porterAppendEnd(str, "e")
							End If
			End If
			
	End If


	'--------------------------------------------------------------------------------------------------------
	'
	'STEP 1C
	'
	'    (*v*) Y -> I                    happy        ->  happi
	'                                    sky          ->  sky

	If porterEndsWith(str, "y") Then
					
					'trim and check for vowel
					temp = porterTrimEnd(str, 1)

					If porterContainsVowel(temp) Then
							str = porterTrimEnd(str, Len("y"))
							str = porterAppendEnd(str, "i")
					End If
			   
	End If

	'retuning the word
	porterAlgorithmStep1 = str

End Function


Private Function porterAlgorithmStep2(str As String) As String
On Error Resume Next

	'STEP 2
	'
	'    (m>0) ATIONAL ->  ATE           relational     ->  relate
	'    (m>0) TIONAL  ->  TION          conditional    ->  condition
	'                                    rational       ->  rational
	'    (m>0) ENCI    ->  ENCE          valenci        ->  valence
	'    (m>0) ANCI    ->  ANCE          hesitanci      ->  hesitance
	'    (m>0) IZER    ->  IZE           digitizer      ->  digitize
	'Also,
	'    (m>0) BLI    ->   BLE           conformabli    ->  conformable
	'
	'    (m>0) ALLI    ->  AL            radicalli      ->  radical
	'    (m>0) ENTLI   ->  ENT           differentli    ->  different
	'    (m>0) ELI     ->  E             vileli        - >  vile
	'    (m>0) OUSLI   ->  OUS           analogousli    ->  analogous
	'    (m>0) IZATION ->  IZE           vietnamization ->  vietnamize
	'    (m>0) ATION   ->  ATE           predication    ->  predicate
	'    (m>0) ATOR    ->  ATE           operator       ->  operate
	'    (m>0) ALISM   ->  AL            feudalism      ->  feudal
	'    (m>0) IVENESS ->  IVE           decisiveness   ->  decisive
	'    (m>0) FULNESS ->  FUL           hopefulness    ->  hopeful
	'    (m>0) OUSNESS ->  OUS           callousness    ->  callous
	'    (m>0) ALITI   ->  AL            formaliti      ->  formal
	'    (m>0) IVITI   ->  IVE           sensitiviti    ->  sensitive
	'    (m>0) BILITI  ->  BLE           sensibiliti    ->  sensible
	'Also,
	'    (m>0) LOGI    ->  LOG           apologi        -> apolog
	'
	'The test for the string S1 can be made fast by doing a program switch on
	'the penultimate letter of the word being tested. This gives a fairly even
	'breakdown of the possible values of the string S1. It will be seen in fact
	'that the S1-strings in step 2 are presented here in the alphabetical order
	'of their penultimate letter. Similar techniques may be applied in the other
	'steps.

	'declaring local variables
	
	Dim i As Byte
	Dim temp As String


	'checking word
	For i = 0 To 20 Step 1
			If porterEndsWith(str, step2(i, 0)) Then
							temp = porterTrimEnd(str, Len(step2(i, 0)))
							If porterCountm(temp) > 0 Then
									str = porterTrimEnd(str, Len(step2(i, 0)))
									str = porterAppendEnd(str, step2(i, 1))
							End If
							Exit For
			End If
	Next i

	'retuning the word
	porterAlgorithmStep2 = str

End Function



Private Function porterAlgorithmStep3(str As String) As String
On Error Resume Next

	'STEP 3
	'
	'    (m>0) ICATE ->  IC              triplicate     ->  triplic
	'    (m>0) ATIVE ->                  formative      ->  form
	'    (m>0) ALIZE ->  AL              formalize      ->  formal
	'    (m>0) ICITI ->  IC              electriciti    ->  electric
	'    (m>0) ICAL  ->  IC              electrical     ->  electric
	'    (m>0) FUL   ->                  hopeful        ->  hope
	'    (m>0) NESS  ->                  goodness       ->  good


	'declaring local variables
	Dim i As Byte
	Dim temp As String

	'checking word
	For i = 0 To 6 Step 1
			If porterEndsWith(str, step3(i, 0)) Then
							temp = porterTrimEnd(str, Len(step3(i, 0)))
							If porterCountm(temp) > 0 Then
									str = porterTrimEnd(str, Len(step3(i, 0)))
									str = porterAppendEnd(str, step3(i, 1))
							End If
							Exit For
			End If
	Next i

	'retuning the word
	porterAlgorithmStep3 = str

End Function



Private Function porterAlgorithmStep4(str As String) As String
On Error Resume Next
	'STEP 4
	'
	'    (m>1) AL    ->                  revival        ->  reviv
	'    (m>1) ANCE  ->                  allowance      ->  allow
	'    (m>1) ENCE  ->                  inference      ->  infer
	'    (m>1) ER    ->                  airliner       ->  airlin
	'    (m>1) IC    ->                  gyroscopic     ->  gyroscop
	'    (m>1) ABLE  ->                  adjustable     ->  adjust
	'    (m>1) IBLE  ->                  defensible     ->  defens
	'    (m>1) ANT   ->                  irritant       ->  irrit
	'    (m>1) EMENT ->                  replacement    ->  replac
	'    (m>1) MENT  ->                  adjustment     ->  adjust
	'    (m>1) ENT   ->                  dependent      ->  depend
	'    (m>1 and (*S or *T)) ION ->     adoption       ->  adopt
	'    (m>1) OU    ->                  homologou      ->  homolog
	'    (m>1) ISM   ->                  communism      ->  commun
	'    (m>1) ATE   ->                  activate       ->  activ
	'    (m>1) ITI   ->                  angulariti     ->  angular
	'    (m>1) OUS   ->                  homologous     ->  homolog
	'    (m>1) IVE   ->                  effective      ->  effect
	'    (m>1) IZE   ->                  bowdlerize     ->  bowdler
	'
	'The suffixes are now removed. All that remains is a little tidying up.

	'declaring local variables
	Dim i As Byte
	Dim temp As String

	'checking word
	For i = 0 To 18 Step 1

			If porterEndsWith(str, step4(i)) Then
			
							temp = porterTrimEnd(str, Len(step4(i)))
							
							If porterCountm(temp) > 1 Then
							
									If porterEndsWith(str, "ion") Then
											If porterEndsWith(temp, "s") Or porterEndsWith(temp, "t") Then
													str = porterTrimEnd(str, Len(step4(i)))
													str = porterAppendEnd(str, "")
											End If
									Else
											str = porterTrimEnd(str, Len(step4(i)))
											str = porterAppendEnd(str, "")
									End If

							End If
							
							Exit For
							
			End If
			
	Next i

	'retuning the word
	porterAlgorithmStep4 = str

End Function



Private Function porterAlgorithmStep5(str As String) As String
On Error Resume Next

	'STEP 5a
	'
	'    (m>1) E     ->                  probate        ->  probat
	'                                    rate           ->  rate
	'    (m=1 and not *o) E ->           cease          ->  ceas
	'
	'STEP 5b
	'
	'    (m>1 and *d and *L) -> single letter
	'                                    controll       ->  control
	'                                    roll           ->  roll

	'declaring local variables
	Dim i As Byte
	Dim temp As String


	'Step5a
	If porterEndsWith(str, "e") Then            'word ends with e
			temp = porterTrimEnd(str, 1)
			If porterCountm(temp) > 1 Then          'm>1
					str = porterTrimEnd(str, 1)
			ElseIf porterCountm(temp) = 1 Then      'm=1
					If Not porterEndsCVC(temp) Then     'not *o
							str = porterTrimEnd(str, 1)
					End If
			End If
	End If


	'--------------------------------------------------------------------------------------------------------
	'
	'Step5b
	If porterCountm(str) > 1 Then
			If porterEndsDoubleConsonent(str) And porterEndsWith(str, "l") Then
					str = porterTrimEnd(str, 1)
			End If
	End If

	'retuning the word
	porterAlgorithmStep5 = str

End Function


Private Function porterEndsWith(str As String, ends As String) As Boolean
On Error Resume Next

	'declaring local variables
	Dim length_str As Byte
	Dim length_ends As Byte
	Dim hold_ends As String

	'finding the length of the string
	length_str = Len(str)
	length_ends = Len(ends)

	'if length of str is greater than the length of length_ends, only then proceed..else return false
	If length_ends >= length_str Then

			porterEndsWith = False
			
	Else

			'extract characters from right of str
			hold_ends = Right$(str, length_ends)
			
			'comparing to see whether hold_ends=ends
			If StrComp(hold_ends, ends) = 0 Then
					porterEndsWith = True
			Else
					porterEndsWith = False
			End If
			
	End If

End Function

Private Function porterContains(str As String, present As String) As Boolean
On Error Resume Next

	'checking whether strr contains present
	If InStr(str, present) = 0 Then
			porterContains = False
	Else
			porterContains = True
	End If

End Function


Private Function porterContainsVowel(str As String) As Boolean
'checking word to see if vowels are present

	Dim chars() As Byte
	Dim i As Byte
	Dim pattern As String

	If Len(str) >= 0 Then

			'find out the CVC pattern
			pattern = returnCVCpattern(str)
			
			'check to see if the return pattern contains a vowel
			If InStr(pattern, "v") = 0 Then
					porterContainsVowel = False
			Else
					porterContainsVowel = True
			End If

	Else
			porterContainsVowel = False
	End If

End Function

Private Function porterEndsDoubleConsonent(str As String) As Boolean
On Error Resume Next
'checking whether word ends with a double consonant (e.g. -TT, -SS).

	'declaring local variables
	Dim holds_ends As String
	Dim hold_third_last As String
	Dim chars() As Byte

	'first check whether the size of the word is >= 2
	If Len(str) >= 2 Then

			'extract 2 characters from right of str
			holds_ends = Right$(str, 2)
			
			'converting string to byte array
			chars = StrConv(holds_ends, vbFromUnicode)
			
			'checking if both the characters are same
			If chars(0) = chars(1) Then
			
					'check for double consonent
					If holds_ends = "aa" Or holds_ends = "ee" Or holds_ends = "ii" Or holds_ends = "oo" Or holds_ends = "uu" Then
							
							porterEndsDoubleConsonent = False
							
					Else
					
							'if the second last character is y, and there are atleast three letters in str
							If holds_ends = "yy" And Len(str) > 2 Then
							
									'extracting the third last character
									hold_third_last = Right$(str, 3)
									hold_third_last = Left$(str, 1)
									
									If Not (hold_third_last = "a" Or hold_third_last = "e" Or hold_third_last = "i" Or hold_third_last = "o" Or hold_third_last = "u") Then
											
											porterEndsDoubleConsonent = False
											
									Else
									
											porterEndsDoubleConsonent = True
											
									End If
							
							Else
							
									porterEndsDoubleConsonent = True
									
							End If
							
					End If
			
			Else
			
					porterEndsDoubleConsonent = False
					
			End If
			
	Else

			porterEndsDoubleConsonent = False
			
	End If

End Function


Private Function porterEndsCVC(str As String) As Boolean
On Error Resume Next
	
	'*o  - the stem ends cvc, where the second c is not W, X or Y (e.g. -WIL, -HOP).

	'declaring local variables
	Dim chars() As Byte
	Dim const_vowel As String
	Dim i As Byte
	Dim pattern As String

	'check to see if atleast 3 characters are present
	If Len(str) >= 3 Then
			
			'converting string to byte array
			chars = StrConv(str, vbFromUnicode)
			
			'find out the CVC pattern
			pattern = returnCVCpattern(str)
			
			'we need to check only the last three characters
			pattern = Right$(pattern, 3)
			  
			'check to see if the letters in str match the sequence cvc
			If pattern = "cvc" Then
					If Not (Chr(chars(UBound(chars))) = "w" Or Chr(chars(UBound(chars))) = "x" Or Chr(chars(UBound(chars))) = "y") Then
							porterEndsCVC = True
					Else
							porterEndsCVC = False
					End If
			Else
					porterEndsCVC = False
			End If
			
	Else

			porterEndsCVC = False

	End If

End Function


Private Function porterTrimEnd(str As String, length As Byte) As String
On Error Resume Next
	'returning the trimmed string
	porterTrimEnd = Left$(str, Len(str) - length)

	End Function

	Private Function porterAppendEnd(str As String, ends As String) As String

	On Error Resume Next

	'returning the appended string
	porterAppendEnd = str + ends

End Function


Private Function porterCountm(str As String) As Byte
On Error Resume Next

	'A \consonant\ in a word is a letter other than A, E, I, O or U, and other
	'than Y preceded by a consonant. (The fact that the term `consonant' is
	'defined to some extent in terms of itself does not make it ambiguous.) So in
	'TOY the consonants are T and Y, and in SYZYGY they are S, Z and G. If a
	'letter is not a consonant it is a \vowel\.

	'declaring local variables
	Dim chars() As Byte
	Dim const_vowel As String
	Dim i As Byte
	Dim m As Byte
	Dim flag As Boolean
	Dim pattern As String

	'initializing
	const_vowel = ""
	m = 0
	flag = False


	If Not Len(str) = 0 Then

			'find out the CVC pattern
			pattern = returnCVCpattern(str)
			
			'converting const_vowel to byte array
			chars = StrConv(pattern, vbFromUnicode)
			
			'counting the number of m's...
			For i = 0 To UBound(chars) Step 1
					If Chr(chars(i)) = "v" Or flag = True Then
							flag = True
							If Chr(chars(i)) = "c" Then
									m = m + 1
									flag = False
							End If
					End If
			Next i
			
	End If

	porterCountm = m

End Function


Private Function returnCVCpattern(str As String) As String

	'local variables
	Dim chars() As Byte
	Dim const_vowel As String
	Dim i As Byte

	'converting string to byte array
	chars = StrConv(str, vbFromUnicode)
			
	'checking each character to see if it is a consonent or a vowel. also inputs the information in const_vowel
	For i = 0 To UBound(chars) Step 1
			
			If Chr(chars(i)) = "a" Or Chr(chars(i)) = "e" Or Chr(chars(i)) = "i" Or Chr(chars(i)) = "o" Or Chr(chars(i)) = "u" Then
					const_vowel = const_vowel + "v"
			ElseIf Chr(chars(i)) = "y" Then
					'if y is not the first character, only then check the previous character
					If i > 0 Then
							'check to see if previous character is a consonent
							If Not (Chr(chars(i - 1)) = "a" Or Chr(chars(i - 1)) = "e" Or Chr(chars(i - 1)) = "i" Or Chr(chars(i - 1)) = "o" Or Chr(chars(i - 1)) = "u") Then
									const_vowel = const_vowel + "v"
							Else
									const_vowel = const_vowel + "c"
							End If
					Else
							const_vowel = const_vowel + "c"
					End If
			Else
					const_vowel = const_vowel + "c"
			End If
					
	Next i
			
	returnCVCpattern = const_vowel

End Function


