Option Explicit
Option Base 1
Option Compare Text

'# <author> Daniel Grass
'# <mail> dani.grass@bluewin.ch

'# Public Subs, Functions and Properties
'#======================================================================================================================
'# Accessible in this class
'#======================================================================================================================
'     |> Get | --- About                :: Returns description of the class.
'     |> Get | --- Name                 :: Returns name of the class.
'     |> Get | --- Version              :: Returns version string for the class [e.g. #.# (year)].


'#======================================================================================================================
'# References to external API's
'#======================================================================================================================
'none

'#======================================================================================================================
'# Application Constants & Enumerations
'#======================================================================================================================
Private Const C_Name             As String = "cLanguageDetector.cls"
Private Const C_NoOfLanguages    As Integer = 4


Public Enum cdLanguage
	lEnglish = 1
	lGerman = 2
	lFrench = 3
	lItalian = 4
End Enum

'#======================================================================================================================
'# Private Variables
'#======================================================================================================================
Private m_English() As Variant                    ' The N-Gram of English 
Private m_German() as Variant                     ' The N-Gram of German 
Private m_French() as Variant                     ' The N-Gram of French
Private m_Italian() as Variant                    ' The N-Gram of Italian
Private m_NGramInput() as Variant                 ' The N-Gram of the input text
Private m_CompResult(1 to C_NoOfLanguages, 1 to 2) as Variant 'Array holding the results of the N-Gram comparison

'#======================================================================================================================
'# Class Initialization, Termination & Properties
'#======================================================================================================================

Private Sub Class_Initialize()
' ************************************************
' Class constructor.
' ************************************************

    Debug.Print "|> Initializing:= " & Me.Name
	
        
    'initialze the language N-Grams
    

End Sub


Private Sub Class_Terminate()
' ************************************************
' Class destructor.
' ************************************************

    Debug.Print "|> Terminating:= " & Me.Name
    
    'empty the lists
        ReDim m_English(1, 1)
        redim m_German(1, 1)
		redim m_French(1, 1)
		redim m_Italian(1, 1)

End Sub

Public Property Get Version() As String
' ************************************************
' Version string of the current class.
' Contains a list of (historical) changes to the class within the comments of the procedure.
' ************************************************
   Version = "Version 1.0 (2016)"  'Initial (official) release.

End Property

Public Property Get About() As String
' ***********************************************
' String that describes the current class.
' ***********************************************
    About = "Language detection class supporting N-Gram based language detection " & vba.vbCrLf & "of English, German, French and Italian. Version: " & Me.Version & "." & VBA.vbCrLf & VBA.vbCrLf
    About = About & "For additional details please contact the author."
End Property

Public Property Get Name() As String
' ***********************************************
' Returns the name of the class.
' ***********************************************
    Name = C_Name
End Property


'#======================================================================================================================
'# N-Gram handling 
'#======================================================================================================================

Public Sub LoadText(Text as string)

	m_NGramInput = CreateNGram(Text, 2)

End Sub

Public Sub DetermineLanguage




End Sub

'#======================================================================================================================
'# N-Gram Base Functions 
'#======================================================================================================================

Private Function CreateNGram(strInput as string, intN as long) as variant
' ***********************************************
' Caltulate the N-Gram
' ***********************************************

 Dim arrNGram, intBound, i, j, strGram, didInc, arrTemp

 If Len(strInput) = 0 Then Exit Function

 ReDim arrNGram(Len(strInput) + 1, 1)
 strInput = Chr(0) & LCase(Trim(strInput)) & Chr(0)
 intBound = -1

 For i = 1 To Len(strInput)-intN+1
	 strGram = Mid(strInput, i, intN)
	 
	 didInc = False

	 For j = 0 To intBound
		 If strGram = arrNGram(j, 0) Then
			 arrNGram(j, 1) = arrNGram(j, 1) + 1
			 didInc = True
			 Exit For
		 End If
	 Next

	 If Not didInc Then
		 intBound = intBound + 1
		 arrNGram(intBound, 0) = strGram
		 arrNGram(intBound, 1) = 1
	 End If
 Next

 ReDim arrTemp(intBound, 1)
 For i = 0 To intBound
	 arrTemp(i, 0) = arrNGram(i, 0)
	 arrTemp(i, 1) = arrNGram(i, 1)
 Next

 CreateNGram = arrTemp
End Function
  
Private Function CompareNGram(arr1() as variant, arr2() as variant) as double
' ***********************************************
' Caltulate Ratio of similarity of the N-Gram
' ***********************************************

 Dim i, j, intMatches, intCount1, intCount2

 intMatches = 0
 intCount1 = 0

 For i = 0 To UBound(arr1)
	 intCount1 = intCount1 + arr1(i, 1)
	 intCount2 = 0

	 For j = 0 To UBound(arr2)
		 intCount2 = intCount2 + arr2(j, 1)

		 If arr1(i, 0) = arr2(j, 0) Then
			 If arr1(i, 1) >= arr2(j, 1) Then
				 intMatches = intMatches + arr2(j, 1)
			 Else
				 intMatches = intMatches + arr1(i, 1)
			 End If
		 End If
	 Next
 Next

 CompareNGram = 2 * intMatches / (intCount1 + intCount2)
End Function


'#======================================================================================================================
'# Language N-Grams  
'#======================================================================================================================

Private Sub LoadEnglish()


End Sub

Private Sub LoadGerman()


End Sub

Private Sub LoadFrench()


End Sub

Private Sub LoadItalian()


End Sub