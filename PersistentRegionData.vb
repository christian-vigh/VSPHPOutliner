'**************************************************************************************************
'*
'*   NAME
'*      PersistentRegionData.vb
'*
'*   DESCRIPTION
'*	Handles persistent region data information (collapsed/expanded state).
'*	The data stored and retrieved is a text file in which each line has two fields :
'*	- The region collapsed state (0 or 1)
'*	- The region unique ID
'*
'*   AUTHOR
'*      Christian Vigh, 11/2011.
'*
'*   HISTORY
'*   [Version : 1.0]    [Date : 2011/11/27]     [Author : CV]
'*      Initial version.
'*
'**************************************************************************************************
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Timers
Imports System.ComponentModel.Composition
Imports Microsoft.VisualStudio.Text.Classification
Imports Microsoft.VisualStudio.Text.Editor
Imports Microsoft.VisualStudio.Text.Operations
Imports Microsoft.VisualStudio.Text.Outlining
Imports Microsoft.VisualStudio.Text.Tagging
Imports Microsoft.VisualStudio.Utilities
Imports Microsoft.VisualStudio.Text
Imports System.Windows.Controls
Imports EnvDTE
Imports PHPOutliner


'==========================================================================================
'=
'=  PersistentRegionData class -
'=      Holds information about the regions collapsing state to be loaded or stored.
'=
'==========================================================================================
Public Class PersistentRegionData
	#Region "Data members"
	Public Shared		PersistentDataDirectory		As  String		=  ""		' Directory for #region collapsing state files
	Private Shared		PersistentDataDirectoryCreated	As  Boolean		=  False	' True when the directory has been created

	Public			RegionDataFilename		As  String				' Region data filename corresponding to this document
	Public			RegionStates			As  String		=  ""		' Region states (0 or 1) in sequential order
	#End Region

	#Region "Constructor"
        '==========================================================================================
        '=
        '=   NAME
        '=      Constructor
        '=
        '=   DESCRIPTION
        '=      Builds the PersistentRegionData object and loads region data, if any.
	'=	Region data is associated with a text document located within a project located
	'=	within a project etc...
        '=
        '==========================================================================================
	Public Sub  New  ( ByVal  FilePath  As  String, Optional ByVal  Autoload  As  Boolean  =  True )
		' Create our own user-dependent directory if needed
		If  ( Not  PersistentDataDirectoryCreated )  Then
			PersistentDataDirectory	=  Environment.GetEnvironmentVariable ( "USERPROFILE" ) & "\" & "Thrak\Visual Studio\PHP\RegionData"

			Utilities.MkDir ( PersistentDataDirectory ) 
			PersistentDataDirectoryCreated = True
		End If

		' A region file of name Filename.{MD5 of full file path} will be created
		Dim	Filename	As  String	=  System.IO.Path.GetFileName ( FilePath )
		Dim	Hash		As  String	=  Utilities.MD5 ( FilePath )

		RegionDataFilename	=  PersistentDataDirectory & "\" & Filename & ".{" & Hash & "}"

		' Load file contents if needed
		If  ( Autoload )  Then
			Load ( )
		End If
	End Sub
	#End Region

	#Region "Properties"

	#Region "IsCollapsed property"
	'==============================================================================================================
	'=
	'=   NAME
	'=      IsCollapsed property.
	'=
	'=   PROTOTYPE
	'=      IsCollapsed(Index) = boolean
	'=	boolean = IsCollapsed(Index)
	'=
	'=   DESCRIPTION
	'=      Sets the collapsed state for the corresponding region index.
	'=
	'=   PARAMETERS
	'=      ByVal Index As Integer -
	'=              Index of the collapsing state to be set or retrieved.
	'=
	'=   RETURN VALUE
	'=      Region collapsing state.
	'=
	'==============================================================================================================
	Public Property  IsCollapsed ( ByVal  Index  As  Integer )  As  Boolean
 		Get
			If  ( Index  >=  0  And  Index  <  Me.RegionStates.Length )  Then
				Dim	Ch	As  Char	=  Me.RegionStates(Index)

				If  ( Ch  =  "1" )  Then
					Return ( True )
				End If
			End If

			Return ( False )
 		End Get

 		Set ( value As Boolean )
			Dim	Ch	As  Char

			If  ( value )
				Ch	=  "1" 
			Else
				Ch	=  "0"
			End If

			Mid ( Me.RegionStates, Index + 1, 1 ) =  Ch
 		End Set
	End Property
	#End Region

	#End Region

	#Region "Public methods"

	#Region  "Load method"
        '==========================================================================================
        '=
        '=   NAME
        '=      Load
        '=
        '=   DESCRIPTION
	'=	Loads region collapsing state data.
	'=	This data is used once, when loading a text document, then discarded.
        '=
        '==========================================================================================
	Public Sub  Load ( )
		' If the file exists, load its contents
		If  ( System.IO.File.Exists ( RegionDataFilename ) )  Then
			Dim	Reader		As  System.IO.TextReader	=  System.IO.File.OpenText ( RegionDataFilename )
			Dim	Contents	As  String			=  Reader.ReadToEnd ( ).Trim ( )

			' Close the reader object
			Reader.Close ( )
			Reader = Nothing 

			Me.RegionStates		=  Contents
		End If
	End Sub
	#End Region

	#Region "Save method"
        '==========================================================================================
        '=
        '=   NAME
        '=      Save
        '=
        '=   DESCRIPTION
	'=	Saves region collapsing state data.
        '=
        '==========================================================================================
	Friend Sub Save ( )
		' Save file contents
		Try
			Dim	Writer		As  System.IO.StreamWriter	=  New System.IO.StreamWriter ( RegionDataFilename )

			Writer.Write ( Me.RegionStates )
			Writer.Close ( )
		Catch  E  As  Exception
			Utilities.ShowException ( E, "Exception encountered while saving persistent data" )
		End Try
	End Sub
	#End Region

	#End Region

End Class
