'**************************************************************************************************
'*
'*   NAME
'*      Utilities.vb
'*
'*   DESCRIPTION
'*	Utility functions.
'*
'*   AUTHOR
'*      Christian Vigh, 11/2011.
'*
'*   HISTORY
'*   [Version : 1.0]    [Date : 2011/11/25]     [Author : CV]
'*      Initial version.
'*
'**************************************************************************************************
Imports System
Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Security.Cryptography
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE90
Imports EnvDTE100
Imports VsLangProj
Imports Microsoft.VisualStudio.OLE.Interop
Imports Microsoft.VisualStudio.Shell.Interop
Imports Microsoft.VisualStudio.Shell
Imports Microsoft.VisualStudio.Text.Classification
Imports Microsoft.VisualStudio.Text.Editor
Imports Microsoft.VisualStudio.Text.Operations
Imports Microsoft.VisualStudio.Text.Outlining
Imports Microsoft.VisualStudio.Text.Tagging
Imports Microsoft.VisualStudio.Utilities
Imports Microsoft.VisualStudio.Text
Imports PHPOutliner



Module Utilities
	#Region "Data members"
	' Flag to prevent overlapping exception windows
	Private		ShowExceptionRunning		=  False 
	#End Region

	#Region "External functions"
	'
	' External functions
	'
	Public Declare Function  GetTickCount		Lib "Kernel32"	( )					As Integer

	<DllImport ( "Vroom.dll", CallingConvention := CallingConvention.Cdecl )>
	Public Function vroom_strchrln ( ByVal  Str  As  String, ByVal  Length  As  Integer, ByVal  Ch  As  Char ) As IntPtr
	End Function

	<DllImport ( "Vroom.dll", CallingConvention := CallingConvention.Cdecl )>
	Public Function vroom_strchrn ( ByVal  Str  As  String, ByVal  Length  As  Integer, ByVal  Ch  As  Char ) As Integer
	End Function

	<DllImport ( "Vroom.dll", CallingConvention := CallingConvention.Cdecl )>
	Public Function vroom_strxln ( ByVal  Str		As  String, 
				       ByVal  StartPosition	As  Integer, 
				       ByVal  EndPosition	As  Integer, 
				       ByVal  Count		As  Integer,
				       ByVal  TabWidth		As  Integer,
			               ByVal  OverflowText	As  String ) As IntPtr
	End Function
	#End Region

	#Region  "GetCurrentDTEInstance"
        '==========================================================================================
        '=
        '=   NAME
        '=      GetCurrentDTEInstance
        '=
        '=   DESCRIPTION
        '=      Returns the currently running DTE instance.
        '=
        '==========================================================================================
	Public Function  GetCurrentDTEInstance ( )  As  EnvDTE80.DTE2
		Dim	DTE  As  EnvDTE80.DTE2

		DTE =  System.Runtime.InteropServices.Marshal.GetActiveObject ("VisualStudio.DTE")
		
		Return ( DTE )
	End Function
	#End Region

	#Region  "GetDocumentFromFilename"
        '==========================================================================================
        '=
        '=   NAME
        '=      GetDocumentFromFilename 
        '=
        '=   DESCRIPTION
        '=      Returns the Document item corresponding to the specified file name, or Nothing if
	'=	the document does not exist.
	'=
	'=   PARAMETERS
	'=	ByVal Filename As String -
	'=		Filename whose Document object is to be retrieved.
        '=
	'=   NOTES
	'=	So far, I did not found the way to directly access DTE members, so I used the
	'=	Microsoft.VisualBasic.CallByName method.
	'=
        '==========================================================================================
	Public Function  GetDocumentFromFilename ( ByVal  Filename  As  String )  As  EnvDTE.Document
		Dim	DTE		As  EnvDTE80.DTE2		=  GetCurrentDTEInstance ( )

		If  ( DTE.Documents  Is  Nothing )  Then
			Return ( Nothing )
		End If

		Dim	Item	As	EnvDTE.ProjectItem		=  DTE.Solution.FindProjectItem (FileName)

		If  ( Item  IsNot  Nothing )  Then
			Return ( Item.Document )
		Else
			Return ( Nothing )
		End If
	End Function
	#End Region

	#Region  "HasExtension"
        '==========================================================================================
        '=
        '=   NAME
        '=      HasExtension
        '=
        '=   DESCRIPTION
        '=      Checks if the supplied path has one of the specified extensions.
	'=
	'=   PARAMETERS
	'=	ByVal Path As String -
	'=		Path whose extension is to be checked.
	'=
	'=	ByVal Extension As Object -
	'=		Specifies the extension against which the path is to be checked.
	'=		This parameter can either be :
	'=		- A string
	'=		- An array of strings
	'=		- A list(of string)
	'=
	'=   RETURNS
	'=	True if the supplied path has one of the specified extensions, false otherwise.
	'=
        '==========================================================================================
	Public Function  HasExtension  ( ByVal  Path  As  String, ByVal  Extension  As  Object )  As  Boolean
		Dim	PathExtension	As  String	=  System.IO.Path.GetExtension ( Path ).ToLower ( )	
		Dim     Extensions	As  String()	=  Nothing

		' Check argument type
		If  ( TypeOf Extension Is String )  Then
			Extensions	=  { Extension }
		Else If  ( TypeOf  Extension  Is  List(Of String) )  Then
			Extensions	=  Extension.ToArray ( )
		Else If  ( TypeOf  Extension  Is  Array )  Then
			Extensions	=  Extension
		' In case of improper type, return failure
		Else
			Return ( False )
		End If

		' Normalize each extension :
		' - Convert to lower case
		' - Add a leading dot if needed
		For  I As Integer  =  LBound ( Extensions )  To  UBound ( Extensions )
			Dim  E		As  String	=  Trim ( Extensions(I).ToLower ( ) )

			If  ( E.Length = 0 )  Then 
				Continue For
			End If

			If  ( E(0)  <>  "." )  Then
				E = "." & E
			End If

			If  ( PathExtension  =  E )  Then
				Return ( True )
			End If
		Next

		' Not found
		Return ( False )
	End Function
	#End Region

	#Region  "MD5"
        '==========================================================================================
        '=
        '=   NAME
        '=      MD5 - Computes an MD5 sum.
        '=
        '=   DESCRIPTION
        '=      Computes The MD5 sum of the specified object.
	'=
	'=   PARAMETERS
	'=	ByVal Filename As String -
	'=		Filename whose Document object is to be retrieved.
        '=
        '==========================================================================================
	Public Function  MD5  ( ByVal  Value  As  String )
		Dim	MD5Service	As  New MD5CryptoServiceProvider 
		Dim	Hash()		As  Byte	=  Encoding.ASCII.GetBytes ( Value )
		Dim	Result		As  String	=  ""

		Hash = MD5Service.ComputeHash ( Hash )

		For Each  HashByte  As  Byte  In  Hash 
			Result = Result & HashByte.ToString( "x2" )
		Next

		Return ( Result ) 
	End Function
	#End Region

	#Region "MkDir"
        '==========================================================================================
        '=
        '=   NAME
        '=      MkDir - Recursive MkDir
        '=
        '=   DESCRIPTION
        '=      Recursively creates a directory and its sub-directories.
	'=
	'=   PARAMETERS
	'=	ByVal Path As String -
	'=		Directory to be created.
	'=
	'=	Optional ByVal Recursive As Boolean -
	'=		When true (the default), directories are recursively created.
	'=
	'=   RETURN VALUE
	'=	The function returns true if the directory tree has been successfully create or
	'=	was already existing, or false if an error occured.
        '=
        '==========================================================================================
	Public Function  MkDir ( ByVal Path As String, Optional ByVal Recursive As  Boolean = True )
		' Paranoia
		If  ( Path  =  "" )  Then
			Return ( False )
		End If

		' When recursive, check the existence of each path element
		If  ( Recursive )  Then
			Dim	Elements	As  String()	=  Path.Split ( "\" )		' Path elements
			Dim	StartIndex	As  Integer	=  0				' Start index will be 1 if path begins with a drive letter
			Dim	FirstElement	As  String	=  Elements ( 0 )		' Get first path element
			Dim	Prepend		As  String	=  ""				' Prepend will contain the drive letter if specified

			' If a drive letter has been specified, then keep it apart to build successive paths
			If ( FirstElement.Length  >  1  And  FirstElement(1) = ":" )  Then
				StartIndex	=  1 
				Prepend		=  FirstElement & "\"
			End If

			' Current directory path
			Dim		CurrentPath	As  String	=  Prepend

			' Loop through path elements
			For  I As Integer = StartIndex  To  UBound(Elements)
				CurrentPath = CurrentPath & Elements(I) & "\"		' Catenate current path element to current path

				' Then try to create this directory if it does not exist
				If  ( Not  System.IO.Directory.Exists ( CurrentPath ) )  Then
					Try
						System.IO.Directory.CreateDirectory ( CurrentPath )
					Catch ex As Exception
						Return ( False )
					End Try
				End If
			Next

			' Directory tree has been successfully created or was already existing : tell that everything is ok
			Return ( True )
		' Otherwise, simply try to create the supplied directory
		Else
			Try
				System.IO.Directory.CreateDirectory ( Path )
				Return ( True )
			Catch ex As Exception
				Return ( False )
			End Try
		End If


	End Function
	#End Region

	#Region "ShowException method"
       '==========================================================================================
        '=
        '=   NAME
        '=      ShowException
        '=
        '=   DESCRIPTION
        '=      Shows the specified exception in a modal form.
	'=
	'=   PARAMETERS
	'=	ByVal E As Exception -
	'=		Exception object.
	'=
	'=	ByVal Title As  String -
	'=		Exception window title.
        '=
        '==========================================================================================
	Public Sub  ShowException ( ByVal  E  As  Exception, ByVal  Title  As  String )
		If  ( Not  ShowExceptionRunning )  Then
			ShowExceptionRunning		=  True 

			Dim	Form		As  ShowExceptionForm			=  New ShowExceptionForm ( )
			Dim	Trace		As  System.Diagnostics.StackTrace	=  New System.Diagnostics.StackTrace ( E, True )
			Dim	Frames		As  System.Diagnostics.StackFrame()	=  Trace.GetFrames()
			Dim	Message		As  String				=  E.Message
			Dim	InnerException	As  Exception				=  E.InnerException
			Dim	Separator	As  String				=  "----------------------------------------------------------------------"

			' Build a complete message by tracing inner exceptions if any
			While  ( InnerException  IsNot  Nothing )
				Message		=  Message & vbCrLf & Separator & vbCrLf
				Message		=  Message & InnerException.ToString ( )
				InnerException	=  InnerException.InnerException
			End While

			' Add a stacktrace
			If  ( Frames.Count  >  0 )  Then
				Message		=  Message & vbCrLf & Separator & vbCrLf & _
						   "STACK TRACE :" & vbCrLf

				For  Each  Frame  As  System.Diagnostics.StackFrame  In  Frames 
					Dim	Line		As  Integer	=  Frame.GetFileLineNumber ( )
					Dim	Column		As  Integer	=  Frame.GetFileColumnNumber ( )
					Dim	Filename	As  String	=  Frame.GetFileName ( )
					Dim	Method		As  String	=  Frame.GetMethod ( ).ToString ( ) 

					Method		=  Method.Replace ( "(", " ( " )
					Method		=  Method.Replace ( ")", " ) " ) 
					Method		=  Method.Trim ( ) 

					If  ( Filename  <>  "" )  Then
						Message		=  Message & "    - Method in file """ & Filename & """, line #" & Line.ToString ( ) & _
								   ", column #" & Column.ToString ( ) & vbCrLf
					Else
						Message		=  Message & "    - Method : " & vbCrLf
					End If

					Message = Message & "        " & Method & vbCrLf & vbCrLf
				Next
			End If

			Message = RTrim ( Message )

			' Show the form
			Form.Text		=  Title
			Form.Message.Text	=  Message 

			Form.ShowDialog ( )

			ShowExceptionRunning		=  False
		End If
	End Sub
	#End Region

	#Region "UnquoteString method"
       '==========================================================================================
        '=
        '=   NAME
        '=      UnquoteString
        '=
        '=   DESCRIPTION
        '=      Unquotes a quoted string.
	'=
	'=   PARAMETERS
	'=	ByVal Value As String -
	'=		String to be unquoted.
	'=
	'=   RETURN VALUE
	'=	The input string, without its surrounding quotes. Two consecutive double quotes
	'=	inside the string will be replace by a single one.
	'=	If the input string is less than two characters long, or is not enclosed withing
	'=	double quotes, it will be returned as is.
        '=
        '==========================================================================================
	Public Function  UnquoteString ( ByVal  Value  As  String )
		If  ( Value.Length  >  1 )  Then
			If  ( Value(0)  =  """"  And  Value( Value.Length - 1 )  =  """" )  Then
				Value	=  Value.Substring ( 1, Value.Length - 2 ).Replace ( """""", """" )
			End If
		End If

		Return ( Value )
	End Function
	#End Region
End Module
