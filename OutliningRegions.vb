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


#Region "ImpactType type"
'==========================================================================================
'=
'=  ImpactType enumeration -
'=      Defines the impact of a buffer modification.
'=
'==========================================================================================
Public Enum  ImpactType
	ImpactsRegionContents		=  0	' Impacts only region inner contents, so do nothing
	ImpactsRegionLayout		=  1	' Impacts region layout, document needs to be reparsed
	ImpactsRegionTag		=  2	' Impacts a region opening or ending tag, so we can wait for the update timer to proceed
End Enum
#End Region

#Region "RegionRange type"
'==========================================================================================
'=
'=  RegionRange type -
'=      Holds information about a region range.
'=
'==========================================================================================
Friend Class  RegionRange 
	Public	X	As  Integer		' Region range, between end of start tag and beginning of end tag
	Public  Y	As  Integer
	Public  RX	As  Integer		' Region range, including the surrounding tags
	Public  RY	As  Integer


	Public Sub New ( ByVal  X  As  Integer, ByVal  Y  As  Integer, ByVal  RX  As  Integer, ByVal  RY  As  Integer )
		Me.X	=  X
		Me.Y	=  Y
		Me.RX	=  RX
		Me.RY	=  RY
	End Sub

	'
	' After, Before -
	'	Checks if the specified point is after/before this region.
	'
	Public Function  After  ( ByVal  Point  As  Integer )  As  Boolean
		If  ( Point  >  Me.RX )  Then
			Return ( True )
		Else
			Return ( False )
		End If
	End Function

	Public Function  Before  ( ByVal  Point  As  Integer )  As  Boolean
		If  ( Point  <  Me.RX )  Then
			Return ( True )
		Else
			Return ( False )
		End If
	End Function

	'
	' Within -
	'	Checks if the specified point is within this region.
	'
	Public Function  Within ( ByVal  Point  As  Integer )
		If  ( Point  >=  Me.RX  And  Point  <=  Me.RY )  Then
			Return ( True )
		Else
			Return ( False )
		End If
	End Function

	'
	' Just for facilitating the debugging...
	'
	Public Overrides Function  ToString ( )  As  String
		Return ( Me.X.ToString ( ) &  ".."  &  Me.Y.ToString ( ) & "   [" & Me.RX.ToString ( ) & ".." & Me.RY.ToString ( ) & "]" )
	End Function
End Class
#End Region

#Region "OutliningRegion class"
'==========================================================================================
'=
'=  OutliningRegions class -
'=      Manages a list of OutliningRegion objects.
'=
'==========================================================================================
Friend Class OutliningRegions
	Inherits  List ( Of  OutliningRegion )

	#Region "Overloaded methods"
        '==========================================================================================
        '=
        '=   NAME
        '=      Add
        '=
        '=   DESCRIPTION
	'=	Adds the specified region and adds an entry in the RegionStates dictionary to
	'=	store the initial region collapsing state.
        '=
	'=   PARAMETERS
	'=	ByVal Region  As  OutliningRegion -
	'=		Region to be added.
	'=
        '==========================================================================================
	Public Overloads Sub  Add ( ByVal  Region  As  OutliningRegion ) 
		Dim	Ch		As  Char


		If  ( Region.IsCollapsed )  Then
			Ch = "1"
		Else
			Ch = "0"
		End If

		MyBase.Add ( Region )
	End Sub
	#End Region

	#Region "Public methods"
	
	#Region "Impacts function"
        '==========================================================================================
        '=
        '=   NAME
        '=      Impacts
        '=
        '=   DESCRIPTION
	'=	Checks if the supplied text buffer changes impacts on the regions layout.
        '=
	'=   PARAMETERS
	'=	ByVal Changes As INormalizedTextChangeCollection -
	'=		List of changes that have occurred in the text buffer.
	'=
	'=   RETURNS   
	'=	True if region layout is impacted, false otherwise.
	'=
        '==========================================================================================
	Public Function  Impacts ( ByVal  Changes  As  INormalizedTextChangeCollection )  As  ImpactType
		' First, try to see if changes had an impact on line count, which would impact regions layout
		For Each  Change  In  Changes 
			If  ( Change.LineCountDelta )  Then
				Return ( ImpactType.ImpactsRegionLayout )
			End If
		Next

		' Then try to check if the changes impacted something in the starting and ending tag lines
		For Each  Change  In  Changes 
			Dim	StartPosition	As  Integer	=  Change.NewPosition
			Dim	EndPosition	As  Integer	=  Change.NewEnd

			For Each  Region  As  OutliningRegion  In  Me
				If  ( StartPosition  >=  Region.StartPoint.LineStartPosition  And  StartPosition  <=  Region.StartPoint.LineEndPosition )  Then
					Return ( ImpactType.ImpactsRegionTag )
				End If

				If  ( StartPosition  >=  Region.EndPoint.LineStartPosition  And  StartPosition  <=  Region.EndPoint.LineEndPosition )  Then
					Return ( ImpactType.ImpactsRegionTag )
				End If

				If  ( EndPosition  >=  Region.StartPoint.LineStartPosition  And  EndPosition  <=  Region.StartPoint.LineEndPosition )  Then
					Return ( ImpactType.ImpactsRegionTag )
				End If

				If  ( EndPosition  >=  Region.EndPoint.LineStartPosition  And  EndPosition  <=  Region.EndPoint.LineEndPosition )  Then
					Return ( ImpactType.ImpactsRegionTag )
				End If
			Next
		Next

		' No impact found on region layout
		Return ( ImpactType.ImpactsRegionContents )
	End Function
	#End Region

	#Region "ShiftBy function"
	'==============================================================================================================
	'=
	'=   NAME
	'=      ShiftBy
	'=
	'=   PROTOTYPE
	'=	Function ShiftBy   ( ByVal  Position		As  Integer, 
	'=			     ByVal  LineDelta		As  Integer, 
	'=			     ByVal  ColumnDelta		As  Integer ) As Integer
	'=
	'=   DESCRIPTION
	'=      Shifts any region impacted by the specified position.
	'=
	'=   PARAMETERS
	'=      ByVal Position As Integer -
	'=              Position to be searched.
	'=
	'=	ByVal LineDelta As Integer -
	'=		Number of shifted lines.
	'=
	'=	ByVal ColumnDelta As Integer -
	'=		Number of shifted column.
	'=
	'=   RETURN VALUE
	'=      The starting position of the first impacted region, or zero if no region impacted.
	'=
	'==============================================================================================================
	Public Function  ShiftBy ( ByVal  Position	As  Integer, 
				   ByVal  LineDelta	As  Integer,
				   ByVal  ColumnDelta	As  Integer )  As  Integer
		Dim  Result	As  Integer	=  -1

		For  I  As  Integer  =  0  To  Me.Count - 1
			Dim  Region	As  OutliningRegion	=  Me(I)

			If  ( Position  <  Region.StartPoint.LineStartPosition )  Then
				If  ( Result  =  -1 )  Then
					Result	=  Region.StartPoint.LineStartPosition
				End If

				Region.ShiftBy ( LineDelta, ColumnDelta )
			Else If  ( Position  >  Region.StartPoint.LineEndPosition  And  Position  <=  Region.EndPoint.LineStartPosition )  Then
				If  ( Result  =  -1 )  Then
					Result	=  Region.StartPoint.LineStartPosition
				End If

				Region.EndPoint.ShiftBy ( LineDelta, ColumnDelta ) 
			End If
		Next

		If  ( Result = -1 )  Then
			Result = 0 
		End If

		Return ( Result )
	End Function
	#End Region

	#End Region

End Class
#End Region
