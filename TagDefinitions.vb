'**************************************************************************************************
'*
'*   NAME
'*      TagDefinitions.vb
'*
'*   DESCRIPTION
'*      Implements a collection of tag definitions.
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


#Region  "TagDefinitions class"
'==========================================================================================
'=
'=  TagDefinitions class -
'=      Implements a collection of tag definition objects.
'=
'==========================================================================================
Friend Class  TagDefinitions
	Inherits  List ( Of TagDefinition )

	#Region "Add method"
        '==========================================================================================
        '=
        '=   NAME
        '=      Add
        '=
        '=   DESCRIPTION
	'=	A shortcut to add a new TagDefinition object to this list. The Add() method
	'=	directly creates the object based on the supplied parameters.
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
	Public Overloads Sub  Add  ( ByVal		StartTag	As  String,
				     ByVal		EndTag		As  String,
				     Optional ByVal	Options		As  TagDefinitionOptions	=  TagDefinitionOptions.All,
				     Optional ByVal	Prefix		As  String			=  "#" )
		MyBase.Add ( New TagDefinition ( StartTag, EndTag, Options, Prefix ) )
	End Sub
	#End Region

	#Region "FindMatch function"
        '==========================================================================================
        '=
        '=   NAME
        '=      FindMatch
        '=
        '=   DESCRIPTION
	'=	Finds a matching tag definition for the supplied input text.
	'=
	'=   PARAMETERS
	'=	ByVal Text As String -
	'=		String to find a match for.
	'=
	'=   RETURN VALUE
	'=	Returns a TagMatch object, or Nothing if no match has been found.
        '=
        '==========================================================================================
	Public Function  FindMatch ( ByVal  Text  As  String )  As  TagMatch
		For Each  Def  As  TagDefinition  In  Me
			If  ( Text.IndexOf ( Def.StartTag )  =  -1  And  Text.IndexOf ( Def.EndTag )  =  -1 )  Then
				Return ( Nothing )
			End If

			Dim	Result	As  TagMatch	=  Def.Matches ( Text )

			If  ( Result.Success )  Then
				Return ( Result ) 
			End If
		Next

		Return ( Nothing )
	End Function
	#End Region
End Class

#End Region