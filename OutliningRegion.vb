'**************************************************************************************************
'*
'*   NAME
'*      OutliningRegion.vb
'*
'*   DESCRIPTION
'*      Implements an outline region.
'*
'*   AUTHOR
'*      Christian Vigh, 11/2011.
'*
'*   HISTORY
'*   [Version : 1.0]    [Date : 2011/11/17]     [Author : CV]
'*      Initial version.
'*
'**************************************************************************************************
Imports System
Imports System.Runtime.InteropServices
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
'=  OutliningRegionPosition class -
'=      Holds information about an outlining region start or end position.
'=
'==========================================================================================
Public Class  OutliningRegionPosition
	Public  LineNumber		As  Integer		' Line number
	Public  LineStartPosition	As  Integer		' Offset of start of line in file
	Public  LineEndPosition		As  Integer		' Offset of end of line in file
	Public  TagOffset		As  Integer		' Tag offset in line

	'
	' Returns the TagPosition (ie, LineStartPosition + TagOffset)
	'
	Public ReadOnly	Property  TagPosition()  As  Integer
		Get
			Return ( Me.LineStartPosition + Me.TagOffset )
		End Get
	End Property

	'
	'  Shifts the point
	'
	Public Sub  ShiftBy ( ByVal  LineCount  As  Integer, ByVal  CharacterCount  As  Integer )
		Me.LineNumber		+=  LineCount
		Me.LineStartPosition	+=  CharacterCount
		Me.LineEndPosition	+=  CharacterCount
	End Sub
End Class


'==========================================================================================
'=
'=  OutliningRegion class -
'=      Holds information about an outlining region
'=
'==========================================================================================
Friend Class OutliningRegion

        #Region "Members"
	' Default hover text
	public Const	DefaultHoverText	As  String		=  "..."

	' Max hover text lines
	Private Const   MaxHoverTextLines	As  Integer		=  20	

        '
        ' Members
        '
        Private		_IsComplete		As  Boolean			=  False				' True if the region definition is complete
	Private		_Title			As  String			=  ""					' Region title (after the '#tag' construct)
	Private		_HoverText		As  String			=  ""					' Text to be displayed when the cursor is over the region
	Private		_IsCollapsed		As  Boolean								' Collapsing state
	Private		_StartPoint		As  OutliningRegionPosition	=  New OutliningRegionPosition ( )	' Start and end positions of the region
	Private		_EndPoint		As  OutliningRegionPosition	=  New OutliningRegionPosition ( )
	Private		_StartTag		As  String								' Starting and ending tags
	Private		_EndTag			As  String		
	Private		_UniqueID		As  String			=  ""					' Unique region ID
	Private		_Parent			As  OutliningRegion		=  Nothing
	Private		_NestingLevel		As  Integer			=  0					' Region nesting level
	Private		_Index			As  Integer								' Region sequential index
	Private		_MatchedText		As  String			=  ""					' The text that matched the regular expression
	Private		_Match			As  TagMatch								' The match result
        #End Region

        #Region "Properties"

	#Region "EndTag property"
	'
	' Gets/sets the ending tag name.
	'
	Public Property EndTag()  As  String
		Get
			Return ( Me._EndTag )
		End Get
		Set ( value As String )
			Me._EndTag = value
		End Set
	End Property
	#End Region

	#Region "EndPoint property"
	'
	' Gets/Sets the startpoint property
	'
	Public Property EndPoint()  As  OutliningRegionPosition 
		Get
			Return ( Me._EndPoint )
		End Get
		Set ( value As OutliningRegionPosition )
			Me._EndPoint = value
		End Set
	End Property
	#End Region

        #Region "Index property"
        '
        ' Index property -
        '       Gets/Sets the region index.
        '
        Public Property  Index ( )  As  Integer
                Get
                        Return ( Me._Index )
                End Get

                Set ( value As Integer )
                        Me._Index = value 
                End Set
        End Property
        #End Region

        #Region "Match property"
        '
        ' Match property -
        '       Gets/Sets the text that matched the regex.
        '
        Public Property  Match ( )  As  TagMatch
                Get
                        Return ( Me._Match )
                End Get

                Set ( value As TagMatch )
                        Me._Match = value 
                End Set
        End Property
        #End Region

        #Region "MatchedText property"
        '
        ' MatchedText property -
        '       Gets/Sets the text that matched the regex.
        '
        Public Property  MatchedText ( )  As  String
                Get
                        Return ( Me._MatchedText )
                End Get

                Set ( value As String )
                        Me._MatchedText = value 
                End Set
        End Property
        #End Region

        #Region "IsCollapsed property"
        '
        ' IsCollapsed property -
        '       Gets/Sets the region collapsing state.
        '
        Public Property  IsCollapsed ( )  As  Boolean
                Get
                        Return ( Me._IsCollapsed )
                End Get

                Set ( value As Boolean )
                        Me._IsCollapsed = value 
                End Set
        End Property
        #End Region

        #Region "IsComplete property"
        '
        ' IsComplete property -
        '       Gets/Sets the region 'complete' flag (when the ending region tag has been encountered).
        '
        Public Property  IsComplete ( )  As  Boolean
                Get
                        Return ( Me._IsComplete )
                End Get

                Set ( value As Boolean )
                        Me._IsComplete = value 
                End Set
        End Property
        #End Region

        #Region "HoverText property"
        '
        ' HoverText property -
        '       Gets/Sets the region hover text.
        '
        Public Property  HoverText ( )  As  Object
                Get
			Return ( _HoverText )
                End Get
                Set ( value As Object )
                        Me._HoverText = value
                End Set
        End Property
        #End Region

	#Region "NestingLevel property"
	'
	' Gets/sets the region nesting level
	'
	Public Property NestingLevel()  As  Integer
		Get
			Return ( Me._NestingLevel )
		End Get
		Set ( value As Integer )
			Me._NestingLevel = value 
		End Set
	End Property
	#End Region

	#Region "Parent property"
	'
	' Gets/Sets the Parent property
	'
	Public Property Parent()  As  OutliningRegion
		Get
			Return ( Me._Parent )
		End Get
		Set ( value As OutliningRegion )
			Me._Parent = value
		End Set
	End Property
	#End Region

	#Region "StartPoint property"
	'
	' Gets/Sets the startpoint property
	'
	Public Property StartPoint()  As  OutliningRegionPosition 
		Get
			Return ( Me._StartPoint )
		End Get
		Set ( value As OutliningRegionPosition )
			Me._StartPoint = value
		End Set
	End Property
	#End Region

	#Region "StartTag property"
	'
	' Gets/sets the starting tag name.
	'
	Public Property StartTag()  As  String
		Get
			Return ( Me._StartTag )
		End Get
		Set ( value As String )
			Me._StartTag = value
		End Set
	End Property
	#End Region

        #Region "Title property"
        '
        ' Title property -
        '       Gets/Sets the region title.
        '
        Public Property  Title ( )  As  Object 
                Get
			Return ( Me._Title )
                End Get

                Set ( value As Object )
			Me._Title	=  value 
                End Set
        End Property
        #End Region

	#Region "UniqueID"
	'
	'  Gets/sets the region unique ID, which is used to track editor changes
	'  The region's unique ID is generated upon the first UniqueID value
	'  retrieval ; this means that all the region fields must have been set
	'  before.
	'
	Public Readonly Property  UniqueID ( )  As  String
		Get
			If  ( Me._UniqueID  =  "" )  Then
				Me._UniqueID = BuildUniqueID ( )
			End If

			Return ( Me._UniqueID )
		End Get
	End Property
	#End Region

	#End Region

	#Region "Private functions"

	#Region "BuildUniqueID function"
        '==========================================================================================
        '=
        '=   NAME
        '=      BuildUniqueID
        '=
        '=   DESCRIPTION
	'=	Builds the unique ID of this region.
	'=
	'=   RETURNS
	'=	An MD5 string built from various member values.
	'=
        '==========================================================================================
	Private Function  BuildUniqueID ( )  As  String
		Dim	Value	As  String	=  Utilities.GetTickCount ( ).ToString ( )

		Return ( Utilities.MD5 ( Value ) )
	End Function
	#End Region

	#End Region

	#Region  "Public functions"

	#Region "ShiftRegion functions"
        '==========================================================================================
        '=
        '=   NAME
        '=      ShiftBy
        '=
        '=   DESCRIPTION
	'=	Adds the specified line/count offset to the region.
	'=
        '==========================================================================================
	Public Sub  ShiftBy ( ByVal  LineCount  As  Integer, ByVal  CharacterCount  As  Integer )
		Me.StartPoint.ShiftBy ( LineCount, CharacterCount )
		Me.EndPoint.ShiftBy   ( LineCount, CharacterCount ) 
	End Sub

	#End Region

	#Region	"SetHoverText method"
        '==========================================================================================
        '=
        '=   NAME
        '=      SetHoverText
        '=
        '=   DESCRIPTION
	'=	Defines the hover text from the specified range in a snapshot.
	'=
	'=   PARAMETERS
	'=	ByVal Snapshot As ITextSnapshot -
	'=		Snapshot where the region is located.
	'=
	'=	ByVal StartPosition As Integer -
	'=		Character start position.
	'=
	'=	ByVal EndPosition As Integer -
	'=		Character end position.
	'=		
	'=
        '==========================================================================================
	Friend Sub  SetHoverText ( ByVal  Snapshot		As  ITextSnapshot, 
				   ByVal  StartPosition		As  Integer, 
				   ByVal  EndPosition		As  Integer )

		If  ( StartPosition  <  EndPosition )  Then
			Dim  Text		As  String	=  Snapshot.GetText ( )
			Dim  HoverPtr		As  IntPtr	=  vroom_strxln ( Text, StartPosition, EndPosition, MaxHoverTextLines, 8, DefaultHoverText )
		
			Me._HoverText	=  Marshal.PtrToStringAnsi ( HoverPtr )
		Else
			Me._HoverText	=  ""
		End If
	End Sub
	#End Region

	#End Region
End Class

