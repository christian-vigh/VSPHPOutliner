'**************************************************************************************************
'*
'*   NAME
'*      OutliningTagger.vb
'*
'*   DESCRIPTION
'*      Implements an outline tagger handler.
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


#Region "ExtendedOutliningRegionTag class"
'==========================================================================================
'=
'=  ExtendedOutliningRegionTag class -
'=      Implements an outlining region tag that holds extra information.
'=
'==========================================================================================
Friend Class  ExtendedOutliningRegionTag
	Inherits  OutliningRegionTag

	Friend OutliningRegion		As  OutliningRegion	


	Public Sub  New  ( ByRef  Region  As  OutliningRegion, ByVal  Title  As  Object, ByVal  Hint  As  Object )
		MyBase.New ( Region.IsCollapsed, False, Title, Hint )
		Me.OutliningRegion	=  Region
	End Sub
End Class
#End Region


#Region "OutliningTagger class"
'==========================================================================================
'=
'=  
'=  OutliningTagger class -
'=      Implements a tag handling object.
'=
'==========================================================================================
Friend NotInheritable Class OutliningTagger
        Implements	ITagger(Of IOutliningRegionTag)

        #Region "Data members"
        '==========================================================================================
        '=
        '= Data members
        '=
        '==========================================================================================
 
        ' Internal data
        Private			TextBuffer			As  ITextBuffer								' Text buffer
	Private			Snapshot			As  ITextSnapshot							' Text snapshot

	' Timer data
	Private	Shared		BufferChanged			As  Boolean				=  False			' True when regions have changed
	Private	Shared		LastBufferChangedTickCount	As  Integer				=  0				' Last tick count on buffer changed event
	Private	Shared		LastGetTagsTickCount		As  Integer				=  0				' Last tick count of GetTags() call
	Private	Shared		BufferChangedTimer		As  Timer								' Buffer update timer
	Private	Shared		InBufferChangedTimer		As  Boolean				=  False
	Private	Shared		InGetTags			As  Boolean				=  False			' True when the GetTags function is running
	Private	Shared		InParseDocument			As  Boolean				=  False			' True when the ParseDocument function is running
	Private	Shared		BufferTimerFrenquencies(,)	As  Integer				=
	   {
	   	{ 1024		,   100 },
		{ 2048		,   150 },
		{ 4096		,   200 },
		{ 8192		,   400 },
		{ 16384		,   600 },
		{ 32768		,   900 },
		{ 65536		,  1000 },
		{ 131072	,  1500 },
		{ 524288	,  2000 },
		{ 1048576	,  4000 },
		{ Int32.MaxValue,  6000 }
	    }	   

	' Properties set by the CreateIViewTagger() method, and other general data
	Friend	WithEvents	Document			As  ITextDocument							' Text document associated with this OUtliningTagger
	Friend	WithEvents	OutliningManager		As  IOutliningManager			=  Nothing			' Outlining manager
	Friend			OutliningRegions		As  OutliningRegions			=  New OutliningRegions ( )	' Outlining regions
	Friend			PersistentRegionData		As  PersistentRegionData						' Persistent region data
	Friend			TagDefinitions			As  TagDefinitions							' Tag definitions
        #End Region

        #Region "Constructor"
        '==========================================================================================
        '=
        '=   NAME
        '=      Constructor
        '=
        '=   DESCRIPTION
        '=      Creates an OutliningTagger object.
        '=
        '=   PARAMETERS
        '=      ByVal Buffer As ITextBuffer -
        '=              Document text buffer.
        '=
        '==========================================================================================
        Public Sub New ( ByVal buffer  As  ITextBuffer )
        End Sub


	Public Sub  FinalizeObject ( ByVal  buffer  As  ITextBuffer )
		Me.TextBuffer	=  buffer
		Me.Snapshot	=  buffer.CurrentSnapshot

		' Add handler for various events
                AddHandler Me.TextBuffer.Changed	, AddressOf OnBufferChanged		' Buffer has changed event

		BufferChangedTimer		=  New Timer ( 100 )
		AddHandler BufferChangedTimer.Elapsed, AddressOf OnBufferChangedTimerElapsed	' Timer taking care of buffer changed state
		BufferChangedTimer.Enabled	=  True
	End Sub
        #End Region

	#Region "ITagger interface implementation"
        #Region "ITagger::GetTags"
        '==========================================================================================
        '=
        '=   NAME
        '=      GetTags - Gets the available tags
        '=
        '=   DESCRIPTION
        '=      Creates an OutliningTagger object.
        '=
        '=   PARAMETERS
        '=      ByVal Buffer As ITextBuffer -
        '=              Document text buffer.
        '=
        '==========================================================================================
        Public Function GetTags ( ByVal spans As NormalizedSnapshotSpanCollection )  As  IEnumerable ( Of ITagSpan ( Of IOutliningRegionTag ) ) _
                                        Implements ITagger ( Of IOutliningRegionTag ).GetTags
		Try
			InGetTags	=  True

			'
			' No span, no cry...
			'
			If ( spans.Count = 0  Or  OutliningRegions.Count = 0 )  Then
				Return ( Nothing )
			End If

			' Return value
			Dim List                        As List ( Of ITagSpan ( Of IOutliningRegionTag) ) _
											=  New List ( Of ITagSpan ( Of IOutliningRegionTag ) )
			'
			' Good example from Microsoft, but no useful comments on that...
			'
			Dim CurrentRegions              As OutliningRegions		=  OutliningRegions
			Dim CurrentSnapshot             As ITextSnapshot                =  Me.Snapshot
			Dim AllSpans			As SnapshotSpan			=  New SnapshotSpan ( spans(0).Start, spans (spans.Count - 1).[End]).TranslateTo ( CurrentSnapshot, SpanTrackingMode.EdgeExclusive )
			Dim StartLineNumber             As Integer                      =  AllSpans.Start.GetContainingLine ( ).LineNumber
			Dim EndLineNumber               As Integer                      =  AllSpans.[End].GetContainingLine ( ).LineNumber
			Dim CurrentTickCount		As Integer			=  Utilities.GetTickCount ( )

			'
			' Loop through regions
			'
			For I  As  Integer  =  CurrentRegions.Count - 1  To  0  Step  -1 
				Dim   Region	As  OutliningRegion	=  CurrentRegions(I)

				If  ( Region.StartPoint.LineNumber  <=  EndLineNumber  AndAlso  Region.EndPoint.LineNumber >=  StartLineNumber  )  Then
					Dim StartLine		As  ITextSnapshotLine		=  CurrentSnapshot.GetLineFromLineNumber ( Region.StartPoint.LineNumber )
					Dim EndLine		As  ITextSnapshotLine		=  CurrentSnapshot.GetLineFromLineNumber ( Region.EndPoint.LineNumber )

					' Until I find how to handle this issue, not extra decoration will take place
					' Issue is the following : For each TextBlock supplied to Visual Studio 
					' - An Initialized() event is sent 
					' - Then a Loaded() event
					' Unfortunately, the Initialized() event is sent to more controls that the currently visible collapsed regions.
					' Upon first keyboard events (just after the document has been loaded), if a scroll down occurs and brings a 
					' collapsed region into view (a collapsed region that received the Initialized event but not the Loaded event)
					' then an exception occurs, telling to disconnect the logical child from its logical parent.
					'Dim  TitleDecoration	As  CollapsedRegionDecoration	=  New CollapsedRegionTitleDecoration ( Region.Title )
					Dim  TitleDecoration	As  String			=  Region.Title
					Dim  HintDecoration	As  CollapsedRegionDecoration	=  New CollapsedRegionHintDecoration ( Region.HoverText )

					' Create the region tag
					Dim RegionTag		As  ExtendedOutliningRegionTag	=  New ExtendedOutliningRegionTag ( Region, TitleDecoration, HintDecoration )

					' The region starts at the beginning of the opening tag, and goes until the end of the line that contains the closing tag.
					List.Add ( New TagSpan ( Of IOutliningRegionTag ) ( New SnapshotSpan ( StartLine.Start + Region.StartPoint.TagOffset, EndLine.End ), RegionTag ) )
				End If
			Next

			' Update tick count
			LastGetTagsTickCount		=  CurrentTickCount
			InGetTags			=  False

			' All done, return
			If  ( List.Count  =  0 )  Then
				Return ( Nothing )
			Else
				Return ( List )
			End If
		Catch  E  As  Exception
			Utilities.ShowException ( E, OutliningTaggerProvider.ExceptionTitle )
			InGetTags	=  False
			Return ( Nothing ) 
		End Try
        End Function
        #End Region

        #Region "ITagger::TagsChanged event"
        '==========================================================================================
        '=
        '=   NAME
        '=      TagsChanged event.
        '=
        '==========================================================================================
        Public Event TagsChanged As EventHandler ( Of SnapshotSpanEventArgs ) Implements ITagger ( Of IOutliningRegionTag ).TagsChanged
        #End Region
	#End Region

	#Region "General Event Handlers"

        #Region "OnBufferChanged event"
        '==========================================================================================
        '=
        '=   NAME
        '=      OnBufferChanged event
        '=
        '=   DESCRIPTION
        '=      Reponds to a buffer changed event.
        '=
        '==========================================================================================
        Private Sub OnBufferChanged ( ByVal Sender As Object, ByVal e As TextContentChangedEventArgs )
		Me.Snapshot			=  e.After.TextBuffer.CurrentSnapshot

		Try
			'
			' If this isn't the most up-to-date version of the buffer, then ignore it for now (we'll eventually get another change event).
			'
			If ( e. After.Version.VersionNumber >=  Me.Snapshot.Version.VersionNumber ) Then
				Me. Snapshot	=  e.After

				' Decide of which action to perform, depending on the modification type
				Dim	Impact		As  ImpactType	=  OutliningRegions.Impacts ( e. Changes )

				Select Case  Impact
					' Modification changed line count : we need to reparse the document
					Case	ImpactType.ImpactsRegionLayout :
						Me.OptimizedReparsing ( e )

					' Modification changed only a region tag contents : wait for the timer to do the reparse
					Case	ImpactType.ImpactsRegionTag :
						BufferChanged		=  True
				End Select
			End If
		Catch  Ex  As  Exception
			Utilities.ShowException ( Ex, OutliningTaggerProvider.ExceptionTitle )
		End Try
        End Sub
        #End Region

        #Region "OnBufferChangedTimerElapsed event"
        '==========================================================================================
        '=
        '=   NAME
        '=      OnBufferChangedTimerElapsed event
        '=
        '=   DESCRIPTION
        '=      Reponds to a buffer changed event.
        '=
        '==========================================================================================
        Private Sub OnBufferChangedTimerElapsed ( ByVal Sender As Object, ByVal EventArgs  As  ElapsedEventArgs )
		If  ( InGetTags )  Then
			Return
		End If

		Try
			' Check if not already running this timer
			If  ( InBufferChangedTimer )  Then
				Return
			End If

			' Don't do anything if buffer did not change
			If  ( Not BufferChanged )  Then
				Return
			End If

			' Set the flag that tells that a timer is already running
			InBufferChangedTimer		=  True

			' Check if we need to update something
			Dim	CurrentTickCount	As  Integer	=  Utilities.GetTickCount ( )
			Dim	CurrentFileSize		As  Integer	=  Me.Document.TextBuffer.CurrentSnapshot.Length

			For I  As  Integer = LBound ( BufferTimerFrenquencies )  To  UBound ( BufferTimerFrenquencies )
				Dim	FileSize	As  Integer	=  BufferTimerFrenquencies ( I, 0 )
				Dim	Interval	As  Integer	=  BufferTimerFrenquencies ( I, 1 )

				' Select the tick count appropriate to the current file size
				If  ( CurrentFileSize  <  FileSize )  Then
					If  ( CurrentTickCount  - LastBufferChangedTickCount  >=  Interval )  Then
						ParseDocument ( )

						' Save region collapsing state
						Me.PersistentRegionData.Save ( )

						LastBufferChangedTickCount	=  CurrentTickCount
					End If

					Exit For
				End If
			Next

			' Reset fields for next event
			InBufferChangedTimer		=  False
		Catch E  As  Exception
			If  ( Me.Document  IsNot  Nothing  And  Me.Document.TextBuffer  IsNot  Nothing )  Then
				Utilities.ShowException ( E, "OnBufferChanged() event" )
			End If
		End Try
	End Sub
	#End Region

	#End Region

        #Region "ParseDocument method"
        '==========================================================================================
        '=
        '=   NAME
        '=      ParseDocument - Parses the region delimiters.
        '=
        '=   DESCRIPTION
        '=      Parses the region delimiters from the supplied input buffer.
	'=
	'=   PARAMETERS
	'=	Optional Byval UsePersistentData As Boolean -
	'=		Tells if persistent data region states should be used instead of current
	'=		region states.
        '=
        '==========================================================================================
	Public Sub  ParseDocument ( Optional Byval UsePersistentData As Boolean = False )
		' Ignore empty documents
		If  ( Me.Snapshot.Length  =  0 )  Then
			Return
		End If

		InParseDocument	=  True

		' Current snapshot
		Dim  SnapShot		As  ITextSnapshot			=  Me.Snapshot
		' Parsed regions
		Dim  ParsedRegions	As  OutliningRegions			=  New OutliningRegions ( )
		' Current region
		Dim  CurrentRegion	As  OutliningRegion			=  Nothing
		Dim  CurrentRegionIndex	As  Integer				=  0
		' Current nesting level
		Dim  NestingLevel	As  Integer				=  0

		' Region positions
		Dim	SnapShotText			As  String		=  SnapShot.GetText ( ) 
		Dim	RegionList			As  IntPtr		=  vroom_strchrln ( SnapShotText, SnapShotText.Length, "#" )
		Dim	RegionCount			As  Integer		=  Marshal.ReadInt32 ( RegionList, 0 )
		Dim	RegionOffsets(RegionCount)	As  Integer
		
		' For the INTEL version of strchrln(), there is some mess when trying to retrieve an offset corresponding
		' to the searched character. For safety reasons, we copy all DWORDS to the RegionOffsets array
		For I As Integer = 0 to RegionCount - 1
			RegionOffsets (I) =  Marshal.ReadInt32 ( RegionList, ( I + 1 ) * 4 )
		Next

		' Stack of regions being processed
		Dim	RegionStack	As  Stack (Of OutliningRegion)	=  New Stack (Of OutliningRegion)

		' Loop through snapshot lines
		For I = 0 To RegionCount - 1
			Dim  Position	As  Integer		=  RegionOffsets (I) 'Here is the mess : Marshal.ReadInt32 ( RegionList, ( I + 1 ) * 4 ) 
			Dim  Line	As  ITextSnapshotLine

			Try
				Line	=  SnapShot.GetLineFromPosition ( Position )
			Catch E As Exception
				MsgBox ( "EXCEPTION ON POSITION " & Position & ", INDEX = " & I & ", COUNT = " & RegionCount)
				Exit For
			End Try 

			Dim  Text	As  String		=  Line.GetText ( )

			' Simple optimization
			If  ( Text.Length  =  0  OrElse  Text.Trim ( ).Length  =  0  OrElse  Text.IndexOf ( "#" )  =  -1 )  Then
				Continue For
			End If

			' Try to find a match
			Dim  Match	As  TagMatch	=  TagDefinitions.FindMatch ( Text )

			' If not tag match found, skip line
			If  ( Match Is Nothing )  Then
				Continue For
			End If

			' If a match has been found for an opening tag then process it
			If  ( Match.Opening )  Then
				Dim	Parent		As  OutliningRegion	=  Nothing

				' Add one more nesting level
				NestingLevel	=  NestingLevel + 1 

				' Start region creation
				CurrentRegion   =  New OutliningRegion ( )
				
				CurrentRegion.NestingLevel	=  NestingLevel
				CurrentRegion.StartTag		=  Match.Tag	
				CurrentRegion.Title		=  Match.Parameters
				CurrentRegion.MatchedText	=  Text
				CurrentRegion.Match		=  Match 
				CurrentRegion.Parent		=  Parent

				' Get region coordinate info
				With  CurrentRegion.StartPoint
					.LineNumber		=  Line.LineNumber
					.LineStartPosition	=  Line.Extent.Start.Position
					.TagOffset		=  Match.LeadingSpaces.Length
					.LineEndPosition	=  Line.Extent.End.Position
				End With

				RegionStack.Push ( CurrentRegion )
			' Otherwise, close the current region and add it to the list
			Else
				Dim	ClosingRegion		As  OutliningRegion 

				' Get last pushed region
				If  ( RegionStack.Count )  Then
					ClosingRegion	=  RegionStack.Pop ( )
				Else
					Continue For
				End If


				' Get region coordinate info
				With  ClosingRegion.EndPoint
					.LineNumber		=  Line.LineNumber
					.LineStartPosition	=  Line.Extent.Start.Position
					.TagOffset		=  Match.LeadingSpaces.Length
					.LineEndPosition	=  Line.Extent.End.Position
				End With

				' Say the region is complete
				ClosingRegion.IsComplete		=  True
				ClosingRegion.SetHoverText ( SnapShot, ClosingRegion.StartPoint.LineEndPosition + 2, ClosingRegion.EndPoint.LineStartPosition )

				' Set collaped state
				ClosingRegion.IsCollapsed		=  Me.PersistentRegionData.IsCollapsed ( CurrentRegionIndex )

				' If current region is closed then add it and decrease nesting level
				ClosingRegion.Index	 =  CurrentRegionIndex
				ParsedRegions.Add ( ClosingRegion )
				CurrentRegionIndex	+=  1
				NestingLevel		-=  1
			End If
		Next

		'
		' Perform the region update
		'
		Dim	RangeLow		As  Integer
		Dim	RangeHigh		As  Integer

		' If not initial load, then try to optimize the update range
		If  ( ParsedRegions.Count  =  0 )  then
			RangeLow	=  -1 
		Else
			RangeLow	=  ParsedRegions(0).StartPoint.LineStartPosition
			RangeHigh	=  ParsedRegions ( ParsedRegions.Count - 1 ).EndPoint.LineEndPosition
		End If

		' Save the parsed regions 
		OutliningRegions	=  ParsedRegions 

		' Set buffer state to clean
		BufferChanged		=  False
		LastBufferChangedTickCount	=  GetTickCount ( )
		InParseDocument		=  False

		' Then fire the TagsChanged event
		If  ( RangeLow  <>  -1 )  Then
			RaiseEvent TagsChanged ( Me, New SnapshotSpanEventArgs ( New SnapshotSpan ( Snapshot, Span.FromBounds ( RangeLow, RangeHigh ) ) ) )
		End If
	End Sub
	#End Region

	#Region "OutliningManager Events"

	#Region "RegionsCollapsed Event"
        '==========================================================================================
        '=
        '=   NAME
	'=	OnRegionsCollapsed event
        '=
        '=   DESCRIPTION
	'=	Updates corresponding internal collapsed state when regions have been collapsed.
	'=	Sets the Regions Modified Flag to true.
	'=
	'=   PARAMETERS
	'=	ByVal Sender As Object -
	'=		Sender object.
	'=
	'=	ByVal EventArgs As RegionCollapsedEventArgs -
	'=		List of regions that have been collapsed.
        '=
        '==========================================================================================
	Public Sub  OnRegionsCollapsed ( ByVal  Sender  As  Object, ByVal  EventArgs  As  RegionsCollapsedEventArgs )  Handles OutliningManager.RegionsCollapsed
		UpdateRegionsState ( EventArgs.CollapsedRegions )
	End Sub
	#End Region

	#Region "RegionsExpanded Event"
        '==========================================================================================
        '=
        '=   NAME
	'=	OnRegionsExpanded event
        '=
        '=   DESCRIPTION
	'=	Updates corresponding internal collapsed state when regions have been expandeded.
	'=	Sets the Regions Modified Flag to true.
	'=
	'=   PARAMETERS
	'=	ByVal Sender As Object -
	'=		Sender object.
	'=
	'=	ByVal EventArgs As RegionExpandedEventArgs -
	'=		List of regions that have been expanded.
        '=
        '==========================================================================================
	Public Sub  OnRegionsExpanded ( ByVal  Sender  As  Object, ByVal  EventArgs  As  RegionsExpandedEventArgs )  Handles OutliningManager.RegionsExpanded
		UpdateRegionsState ( EventArgs.ExpandedRegions )
	End Sub
	#End Region

	#End Region

	#Region "Private functions"

	#Region "LengthWithoutCrLf function"
        '==========================================================================================
        '=
        '=   NAME
	'=	LengthWithoutCrLf
        '=
        '=   DESCRIPTION
	'=	Returns the length of a string without counting carriage returns/line feeds.
	'=
	'=   PARAMETERS
	'=	ByVal Value As String
	'=		String whose length is to be computed.
        '=
        '==========================================================================================
	Public Function  LengthWithoutCrLf ( ByVal  Value  As  String )
		Dim	Result		As  Integer	=  0 
		Dim	Ch		As  Char 

		For  I  As  Integer	=  0  To  Value.Length - 1
			Ch	=  Value (I)

			If  ( Ch  <>  vbCr  And  Ch  <>  vbLf )  Then
				Result +=  1
			End If
		Next

		Return ( Result )
	End Function
	#End Region

	#Region "OptimizeReparsing method method"
        '==========================================================================================
        '=
        '=   NAME
	'=	OptimizedReparsing
        '=
        '=   DESCRIPTION
	'=	Reparses the document only if needed.
	'=
	'=   PARAMETERS
	'=	ByVal EventArg As TextContentChangedEventArgs
	'=		Change list.
        '=
        '==========================================================================================
	Private Sub  OptimizedReparsing ( ByVal  EventArgs  As  TextContentChangedEventArgs )
		' If no change occurred, then no reparsing is necessary
		If  ( EventArgs.Changes  Is Nothing  OrElse  EventArgs.Changes.Count  =  0 )  Then
			Return
		End If

		' We don't handle multiple changes and apparently VS2010 neither...
		If  ( EventArgs.Changes.Count  >  1 )  Then
			ParseDocument ( ) 
			Return 
		End If

		' Check that the change impacted the line count
		Dim	Change			As  ITextChange	=  EventArgs.Changes(0)
		Dim	ChangedText		As  String
		Dim	ChangedTextPosition	As  Integer

		If  ( Change.LineCountDelta  =  0 )  Then
			Return
		End If

		' Changed text data
		If  ( Change.LineCountDelta  >  0 )  Then
			ChangedText		=  Change.NewText
			ChangedTextPosition	=  Change.NewPosition
		Else
			ChangedText		=  Change.OldText
			ChangedTextPosition	=  Change.OldPosition
		End If

		' If the changed text contains a sharp, then reparse document
		If  ( vroom_strchrn ( ChangedText, ChangedText.Length, "#" )  <>  -1 )  Then
			ParseDocument ( ) 
			Return
		End If

		' If no region defined, return
		If (  OutliningRegions.Count  =  0 )  Then
			Return
		End If

		Dim  ActualLength	As  Integer		=  LengthWithoutCrLf ( Change.NewText )
		Dim  Position		As  Integer		=  Change.NewPosition
		Dim  StartPosition	As  Integer		=  Me.OutliningRegions.ShiftBy ( Position, Change.LineCountDelta, ActualLength )
		Dim  EndPosition	As  Integer	

		If  ( StartPosition  <>  -1 )  Then
			' Compute ending position
			EndPosition	=  Math.Min ( OutliningRegions ( OutliningRegions.Count - 1 ).EndPoint.LineEndPosition, 
								Me.Snapshot.Length - 1 )

			' Then raise the TagsChanged event, so that VS will call our GetTags() method to update regions
			RaiseEvent TagsChanged ( Me, New SnapshotSpanEventArgs ( 
							New SnapshotSpan ( Snapshot, Span.FromBounds ( 0, EndPosition ) ) ) )
		End If
	End Sub
	#End Region

	#Region "UpdateRegionsState method"
        '==========================================================================================
        '=
        '=   NAME
	'=	UpdateRegionsState
        '=
        '=   DESCRIPTION
	'=	Updates the regions collapsing state.
	'=
	'=   PARAMETERS
	'=	ByVal Regions As IEnumerable ( Of ICollapsible ) -
	'=		List of regions whose collapsed state is to be updated.
        '=
        '==========================================================================================
	Private Sub  UpdateRegionsState ( ByVal  Regions  As  IEnumerable ( Of ICollapsible ) )
		Dim  Index		As  Integer

		For Each  Region  As  ICollapsible  In  Regions 
			Dim  ExtendedRegion	As  ExtendedOutliningRegionTag	=  DirectCast ( Region.Tag, ExtendedOutliningRegionTag )
			Dim  RegionToUpdate	As  OutliningRegion		=  ExtendedRegion.OutliningRegion

			' Set region collapsing state
			If  ( RegionToUpdate  IsNot  Nothing )  Then
				RegionToUpdate.IsCollapsed			=  Region.IsCollapsed
				Index						=  RegionToUpdate.Index
				Me.PersistentRegionData.IsCollapsed ( Index )	=  Region.IsCollapsed
			End If
		Next

		BufferChanged	=  True
	End Sub
	#End Region

	#End Region
End Class
#End Region
