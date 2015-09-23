'**************************************************************************************************
'*
'*   NAME
'*      TagDefinition.vb
'*
'*   DESCRIPTION
'*      Implements an outline tag definition.
'*
'*   AUTHOR
'*      Christian Vigh, 11/2011.
'*
'*   HISTORY
'*   [Version : 1.0]    [Date : 2011/11/19]     [Author : CV]
'*      Initial version.
'*
'**************************************************************************************************
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions


#Region "TagDefinitionOptions enumeration"
'==========================================================================================
'=
'=  TagDefinitionOptions class -
'=      Options used for tag definition creation.
'=
'==========================================================================================
Public Enum  TagDefinitionOptions
	ParametersAllowed	=  &H0001	' Parameters are allowed after the tag name
	RemoveQuotes		=  &H0002	' If a string parameter is specified, then enclosing double quotes are removed

	All			=  &HFFFF	' All of the above
End Enum
#End Region


#Region  "TagDefinition class"
'==========================================================================================
'=
'=  TagDefinition class -
'=      Implements a tag definition object.
'=
'==========================================================================================
Friend Class  TagDefinition
	#Region "Data members" 
	' Default Regex options
	Private Const	RegexOptions  As  RegexOptions	=  RegexOptions.Compiled +
							   RegexOptions.IgnoreCase +
							   RegexOptions.IgnorePatternWhitespace +
							   RegexOptions.ExplicitCapture
							
		
	' Public members
	Public		StartTag		As  String			' Start tag string
	Public		EndTag			As  String			' End tag string
	Public		Prefix			As  String			' Prefix string
	Public		Options			As  TagDefinitionOptions	' Tag definition options

	' Private members
	Friend		RegexString		As  String			' Built regex string to match either the opening or closing tag
	Friend		RegexObject		As  Regex			' Compiled regex
	#End Region


	#Region "Constructor"
        '==========================================================================================
        '=
        '=   NAME
        '=      Constructor
        '=
        '=   DESCRIPTION
	'=	Builds a tag definition object/
        '=
	'=   PARAMETERS
	'=	ByVal StartTag As String -
	'=		Starting tag.
	'=
	'=	ByVal EndTag As String -
	'=		Ending tag.
	'=
	'=	Optional ByVal Options As TagDefinitionOptions -
	'=		Specifies the tag definition options. Can be any combination of :
	'=
	'=		- ParametersAllowed :
	'=			Indicates if parameters are allowed after the opening tag.
	'=		- RemoveQuotes :
	'=			Double quotes enclosing tag parameters will be removed.
	'=
	'=		The specifal value 'All' implies all options.
	'=		The default value is 'All'.
	'=
	'=	Optional ByVal Prefix As String -
	'=		Prefix string to be placed before the tag. The default value is "#".
	'=
        '==========================================================================================
	Public Sub  New  ( ByVal		StartTag	As  String,
			   ByVal		EndTag		As  String,
			   Optional ByVal	Options		As  TagDefinitionOptions	=  TagDefinitionOptions.All,
			   Optional ByVal	Prefix		As  String			=  "#" )
		Me.StartTag		=  StartTag
		Me.EndTag		=  EndTag
		Me.Options		=  Options
		Me.Prefix		=  Prefix
		
		Me.RegexString		=  "(?isx-mn:)^" & BuildRegex ( ) & "$"
		Me.RegexObject		=  New Regex ( Me.RegexString, RegexOptions ) 
	End Sub
	#End Region

	#Region "Public functions"

	#Region "Matches function"
        '==========================================================================================
        '=
        '=   NAME
        '=      Matches
        '=
        '=   DESCRIPTION
	'=	Checks if the specified input text matches this tag definition.
	'=
	'=   PARAMETERS
	'=	ByVal Text As String -
	'=		String to find a match for.
	'=
	'=   RETURN VALUE
	'=	Returns a TagMatch object.
        '=
        '==========================================================================================
	Public Function  Matches  ( ByVal  Text  As  String )  As  TagMatch 
		Return New TagMatch ( Me, Text, Me.RegexObject )
	End Function
	#End Region

	#End Region

	#Region "Private functions" 

	#Region "BuildRegex function"
        '==========================================================================================
        '=
        '=   NAME
        '=      BuildRegex
        '=
        '=   DESCRIPTION
	'=	Builds the regex string to match the specified opening and closing tags.
	'=
	'=   RETURN VALUE
	'=	Returns the regex string.
        '=
        '==========================================================================================
	Private Function  BuildRegex (  )  As  String
		Dim	Regex	As  String

		Regex	=  "(?<spaces> \s*) " & "(?<prefix> " & Regexify ( Me.Prefix ) & ") " & "\s* " & 
			   "( ( (?<starttag> " & Regexify ( Me.StartTag ) & ") "

		If  ( Me.Options  &  TagDefinitionOptions.ParametersAllowed )  Then
			Regex = Regex & "\s+ (?<params> .+?) \s* ) "
		Else
			Regex = Regex & "\s* ) "
		End If

		Regex   =  Regex  &  " | ( (?<endtag> " & Regexify ( Me.EndTag ) & ") \s* ) )"

		Return ( Regex )
	End Function
	#End Region

	#Region "Regexify function"
        '==========================================================================================
        '=
        '=   NAME
        '=      Regexify
        '=
        '=   DESCRIPTION
	'=	Escapes all the characters in the specified string so that it can be put into a
	'=	regular expression.
	'=
	'=   PARAMETERS
	'=	ByVal Text As String -
	'=		String to find a match for.
	'=
	'=   RETURN VALUE
	'=	Returns the escaped string.
        '=
        '==========================================================================================
	Private Function  Regexify ( ByVal  Value  As  String )  As  String
		Dim	Result		As  String	=  ""

		For  I As Integer = 0 To Value.Length - 1
			Dim	Ch  As  Char	=  Value (I)

			Select Case  Ch
				Case  "[", "]", "\", "-" :
					Result  =  Result & "[\" & Ch & "]"

				Case Else :
					Result  =  Result & "[" & Ch & "]"
			End Select
		Next  I

		Return ( Result )
	End Function
	#End Region

	#End Region
End Class
#End Region