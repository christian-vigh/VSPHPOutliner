'**************************************************************************************************
'*
'*   NAME
'*      CollapsedRegionTitle.vb
'*
'*   DESCRIPTION
'*      Implements a collapsed region title.
'*
'*   AUTHOR
'*      Christian Vigh, 12/2011.
'*
'*   HISTORY
'*   [Version : 1.0]    [Date : 2011/12/05]     [Author : CV]
'*      Initial version.
'*
'**************************************************************************************************
Imports System
Imports System.Collections.Generic
Imports System.Windows.Media
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Timers
Imports System.Windows
Imports System.Windows.Markup
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


#Region  "CollapsedRegionDecoration class"
'==========================================================================================
'=
'=  CollapsedRegionDecoration class -
'=      Base class for decoration objects that are used for collapsed region text and hint
'=	text.
'=
'==========================================================================================
Friend MustInherit Class CollapsedRegionDecoration
	Inherits	TextBlock
	Implements	IContentHost, IAddChild, IServiceProvider


	Public		InitializedState	As  Boolean		=  False
	Public		LoadedState		As  Boolean		=  False


	#Region "Constructor"
        '==========================================================================================
        '=
        '=   NAME
        '=      Constructor
        '=
        '=   DESCRIPTION
	'=	Builds a decoration object (a TextBlock).
        '=
        '==========================================================================================
	Friend Sub  New  ( ByVal  Title  As  String )
		MyBase.New ( )
		Me.Text = Title		' Save text
		SetAppearance ( )	' Set appearance settings
	End Sub
	#End Region

	#Region  "GetObject method"
        '==========================================================================================
        '=
        '=   NAME
        '=      GetObject
        '=
        '=   DESCRIPTION
	'=	Returns the underlying object.
	'=	The outlining tagger classes use this method to retrieve the current object but
	'=	should testing needs require it, it could be overridden to return a single string,
	'=	for example.
        '=
        '==========================================================================================
	Public Function  GetObject ( )  As  Object
		Return ( Me )
	End Function
	#End Region

	#Region  "ToString method"
        '==========================================================================================
        '=
        '=   NAME
        '=      ToString
        '=
        '=   DESCRIPTION
	'=	Overrides the standard ToString() method to return the text put into the Texblock
	'=	control instead of returning the object name.
        '=
        '==========================================================================================
	Public Overrides Function ToString ( )  As  String
		Return ( Me.Text )
	End Function
	#End Region

	#Region  "SetAppearance method"
        '==========================================================================================
        '=
        '=   NAME
        '=      SetAppearance
        '=
        '=   DESCRIPTION
	'=	Sets the default TextBlock appearance settings.
        '=
        '==========================================================================================
	Protected Overridable Sub  SetAppearance ( )
		Me.FontFamily		=  New FontFamily ( "Consolas" )
	End Sub
	#End Region

	#Region "RemoveQuotes function"
	'
	'  Removes the quotes around a title text
	'
	Protected Shared Function  RemoveQuotes ( ByVal  Value  As  String )
		If  ( Value  Is  Nothing )
			Return ( "" )
		End If

		Dim	Length		As  Integer	=  Value.Length 

		If  ( Length  <  2 )  Then
			Return ( Value )
		End If

		If  ( Value.Chars(0)  =  """"  And  Value.Chars ( Length - 1 )  =  """" )  Then
			Value	=  Value.Substring ( 1, Length - 2 ).Replace ( """""", """" )
		End If

		Return ( Value )
	End Function
	#End Region

'	Public Sub OnInitializedEvent ( ByVal  Sender  As  Object, ByVal  EventArgs  As EventArgs ) Handles  Me.Initialized
'		Me.InitializedState = True
'	End Sub

'	Public Sub OnLoadedEvent ( ByVal  Sender  As  Object, ByVal  EventArgs  As RoutedEventArgs ) Handles  Me.Loaded
'		Me.LoadedState		=  True
'	End Sub

'	Public Sub OnVisibleChangedEvent ( ByVal  Sender  As  Object, ByVal  EventArgs  As DependencyPropertyChangedEventArgs ) Handles  Me.IsVisibleChanged
'	End Sub

'	Public Sub OnRequestBringIntoView ( ByVal  Sender  As  Object, ByVal  EventArgs  As RequestBringIntoViewEventArgs ) Handles  Me.RequestBringIntoView
'	End Sub
End Class
#End Region


#Region  "CollapsedRegionDecoration class"
'==========================================================================================
'=
'=  CollapsedRegionTitleDecoration class -
'=      Implements the control used for displaying a collapsed region title.
'=
'==========================================================================================
Friend Class  CollapsedRegionTitleDecoration 
	Inherits	CollapsedRegionDecoration 

	#Region "Constructor"
	'
	'  Class constructor
	'
	Friend Sub New ( ByVal  Title  As  String )
		MyBase.New ( RemoveQuotes ( Title ) )
	End Sub
	#End Region

	#Region "SetAppearance method"
	'
	' Set the object's appearance settings
	'
	Protected Overrides Sub SetAppearance ( ) 
		MyBase.SetAppearance ( )
		Me.Background		= Brushes.FloralWhite
		Me.FontSize		= 12
		Me.Padding		= New Thickness(5, 2, 5, 2)
	End Sub
	#End Region
End Class
#End Region


#Region  "CollapsedRegionHintDecoration class"
'==========================================================================================
'=
'=  CollapsedRegionHintDecoration class -
'=      Implements the control used for displaying a collapsed region hover text.
'=
'==========================================================================================
Friend Class  CollapsedRegionHintDecoration 
	Inherits	CollapsedRegionDecoration 

	#Region "Constructor"
	'
	'  Class constructor
	'
	Friend Sub New ( ByVal  Title  As  String )
		MyBase.New ( Title )
	End Sub
	#End Region

	#Region "SetAppearance method"
	'
	' Set the object's appearance settings
	'
	Protected Overrides Sub SetAppearance ( ) 
		MyBase.SetAppearance ( )
		Me.FontSize		= 12
		Me.Padding		= New Thickness(6, 3, 6, 3)
	End Sub
	#End Region
End Class
#End Region