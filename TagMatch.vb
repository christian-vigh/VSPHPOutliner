'**************************************************************************************************
'*
'*   NAME
'*      OutlineTagDefinition.vb
'*
'*   DESCRIPTION
'*      Implements an outline tag definition.
'*
'*   AUTHOR
'*      Christian Vigh, 12/2011.
'*
'*   HISTORY
'*   [Version : 1.0]    [Date : 2011/12/07]     [Author : CV]
'*      Initial version.
'*
'**************************************************************************************************
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions


#Region  "TagMatch class"
'==========================================================================================
'=
'=  TagMatch class -
'=      Implements a tag definition match.
'=
'==========================================================================================
Friend Class TagMatch
	#Region "Data members"
	Public		Success		As  Boolean		=  False	' True when the match is successful
	Public		LeadingSpaces	As  String		=  ""		' Leading spaces string
	Public		Tag		As  String		=  ""		' Found tag
	Public		Parameters	As  String		=  ""		' Tag parameters
	Public		Opening		As  Boolean		=  False	' When true, the tag is an opening tag
	Public		Definition	As  TagDefinition			' Tag definition
	#End Region


	#Region "Constructor"
        '==========================================================================================
        '=
        '=   NAME
        '=      FindMatch
        '=
        '=   DESCRIPTION
	'=	Builds a TagMatch object.
	'=
	'=   PARAMETERS
	'=	ByVal TagDefinition As TagDefinition -
	'=		Tag definition object.
	'=
	'=	ByVal Text As String -
	'=		String to find a match for.
	'=
	'=	ByVal Regex As String -
	'=		Regex object to be used for matching.
	'=
        '==========================================================================================
	Public Sub New ( ByVal  TagDefinition  As  TagDefinition, ByVal  Text  As  String, ByVal Regex  As  Regex )
		Dim     Match		As  Match	=  Regex.Match ( Text )
		
		If  ( Match.Success )  Then
			Me.Success		=  True
			Me.LeadingSpaces	=  Match.Groups ( "spaces" ).Value
			Me.Definition		=  TagDefinition

			If  ( Match.Groups ( "starttag" ).Value  =  "" )  Then
				Me.Tag		=  Match.Groups ( "endtag" ).Value
				Me.Opening	=  False
			Else
				Me.Tag		=  Match.Groups ( "starttag" ).Value
				Me.Opening	=  True
				Me.Parameters	=  Utilities.UnquoteString ( Match.Groups ( "params" ).Value )
			End If
		End If
	End Sub
	#End Region
End Class
#End Region