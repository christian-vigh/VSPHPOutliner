'**************************************************************************************************
'*
'*   NAME
'*      OutliningTaggerProvider.vb
'*
'*   DESCRIPTION
'*      Implements an outline tagger provider.
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
Imports System.ComponentModel.Composition
Imports Microsoft.VisualStudio.Text
Imports Microsoft.VisualStudio.Text.Classification
Imports Microsoft.VisualStudio.Text.Editor
Imports Microsoft.VisualStudio.Text.Outlining
Imports Microsoft.VisualStudio.Text.Tagging
Imports Microsoft.VisualStudio.Utilities
Imports PHPOutliner


#Region "OutliningTaggerProviderInfo class"
'==========================================================================================
'=
'=  OutliningTaggerProviderInfo class -
'=      Our OutliningTaggerProvider class implements two interfaces :
'=	- ITaggerProvider, for region tagging
'=	- IViewTaggerProvider. The sole reason of its existence is to retrieve a TextView
'=	  object that will later be used by the OutliningTagger class (this is the only
'=	  way I found to retrieve a TextView from a TextBuffer).
'=
'=	The CreateITagger() method is called twice, then the CreateIViewTagger() one.
'=	We thus need to remember a certain quantity of data between two calls, which is
'=	the goal of this class.
'=
'==========================================================================================
Friend Class  OutliningTaggerProviderInfo
	#Region "Data members"
	'
	' Data stored between two successive calls
	'
	Public	TextDocument		As  ITextDocument		' Text document
	Public  OutliningTagger		As  OutliningTagger		' Our outlining tagger processor object
	Public  TaggerObject		As  Object			' The return value of the CreatexxxTagger() method
	Public  TextView		As  ITextView			' The TextView we need to provide visual feedback
	#End Region

	#Region "Constructor"
	'
	' Constructor -
	'	Builds the current instance of this class. Any unspecified argument can be directly assigned later
	'
	Public Sub  New  ( Optional ByVal  TextDocument		As  ITextDocument	=  Nothing, _
			   Optional ByVal  OutliningTagger	As  OutliningTagger	=  Nothing, _
			   Optional ByVal  TaggerObject		As  Object		=  Nothing, _
			   Optional ByVal  TextView		As  ITextView		=  Nothing )
		Me.TextDocument		=  TextDocument
		Me.OutliningTagger	=  OutliningTagger
		Me.TaggerObject		=  TaggerObject
		Me.TextView		=  TextView
	End Sub
	#End Region
End Class
#End Region


#Region "OutliningTaggerProvider class"
'==========================================================================================
'=
'=  OutliningTaggerProvider class -
'=      Exports the ITaggerProvider interface.
'=
'==========================================================================================
'< ContentType ( "text" ) >									_
'< ContentType ( PHPOutliner.FileAndContentTypeDefinitions.PHPTypeName ) >			_
< Export ( GetType ( IViewTaggerProvider ) ) >							_
< TagType ( GetType ( TextMarkerTag ) ) >							_
< Export ( GetType ( ITaggerProvider ) ) >							_
< TagType ( GetType ( IOutliningRegionTag ) ) >							_
< ContentType ( "text" ) >									_
Public NotInheritable Class OutliningTaggerProvider
	Implements ITaggerProvider
	Implements IViewTaggerProvider

	#Region "Data members"
	' Title for exceptions
	Public Const	ExceptionTitle			=  "Exception encountered in the PHPOutliner Visual Studio Extension"

	' Allowed PHP extensions
	Private Shared	PHPExtensions			=  { ".php", ".phs", ".phpclass", ".phpinclude", ".phpinc", ".phpscript" }

	' OutliningManagerService
	<Import()>
	Dim		OutliningManagerService		As  IOutliningManagerService


	' List of taggers, one per opened document
	Dim		OpenedDocuments			As Dictionary ( Of String, OutliningTaggerProviderInfo )  =  _
								New Dictionary ( Of String, OutliningTaggerProviderInfo )
	#End Region


	#Region "CreateITagger"
        '==========================================================================================
        '=
        '=   NAME
        '=      CreateITagger
        '=
        '=   DESCRIPTION
	'=	Called by Visual Studio to create the ITagger object that will be used to outline
	'=	code regions (ie, the OutliningTagger class).
	'=	For some unknown reason, this method is called twice.
        '=
        '==========================================================================================
	Public Function CreateITagger ( Of T As ITag ) ( ByVal Buffer As ITextBuffer ) As ITagger ( Of T )  Implements ITaggerProvider.CreateTagger
		Try
			' Get the TextDocument object from the supplied buffer
			Dim  TextDocument		As  ITextDocument			=  Nothing

			Buffer.Properties.TryGetProperty ( GetType ( ITextDocument ), TextDocument ) 

			' Check if this document is not already opened
			Dim  ProviderInfo		As  OutliningTaggerProviderInfo	=  Nothing

			If  ( OpenedDocuments. TryGetValue ( TextDocument.FilePath, ProviderInfo ) )  Then
				ProviderInfo.TextDocument	=  TextDocument		' Remember the document object if not already done

				Return ( TryCast ( ProviderInfo.TaggerObject, ITagger ( Of T ) ) )
			End If

			' Create the outlining tagger object
			Dim	OutliningTagger		As  OutliningTagger			=  New OutliningTagger ( buffer ) 

			'Create a single ITagger property for this ItextBuffer object.
			Dim	SC			As  Func ( Of ITagger ( Of T ) )	=  Function ( )  TryCast ( OutliningTagger, ITagger ( Of T ) )
			Dim	Tagger			As  ITagger ( Of T )			=  buffer.Properties.GetOrCreateSingletonProperty ( Of ITagger ( Of T ) ) ( SC )

			' Add the current document to our list of opened documents
			Dim	Info	As  OutliningTaggerProviderInfo				=  New  OutliningTaggerProviderInfo ( TextDocument, OutliningTagger, Tagger )

			OpenedDocuments.Add ( TextDocument.FilePath, Info )

			' All done, return
			Return ( Tagger )
		Catch  E  As  Exception
			Utilities.ShowException ( E, ExceptionTitle )
			Return ( Nothing ) 
		End Try
	End Function
	#End Region


	#Region "CreateIViewTagger"
        '==========================================================================================
        '=
        '=   NAME
        '=      CreateIViewTagger
        '=
        '=   DESCRIPTION
	'=	The main purpose of this method is to retrieve a TextView object, since I did not
	'=	find any way to do that starting from a TextBuffer.
	'=	The additional purpose of this method is to finalize the initialization of our
	'=	OutliningTagger object.
        '=
        '==========================================================================================
	Public Function CreateIViewTagger ( Of T As ITag ) ( ByVal TextView As ITextView, ByVal Buffer As ITextBuffer ) As ITagger ( Of T ) Implements IViewTaggerProvider.CreateTagger
		Try
			' That was in the Microsoft walkthrough sample...
			If ( TextView Is Nothing ) Then
				Return Nothing
			End If

			' Get the TextDocument corresponding to this view
			Dim     TextDocument		As  ITextDocument		=  Nothing

			TextView.TextDataModel.DocumentBuffer.Properties.TryGetProperty ( GetType ( Microsoft.VisualStudio.Text.ITextDocument ), TextDocument )

			' Check if this document is not already opened
			Dim  ProviderInfo		As  OutliningTaggerProviderInfo	=  OpenedDocuments ( TextDocument.FilePath )

			ProviderInfo.TextDocument	=  TextDocument

			' Check that file extension is correct
			If  ( Not Utilities.HasExtension ( TextDocument.FilePath, PHPExtensions ) ) Then
				Return ( Nothing )
			End If

			'  Create the outlining manager here, since I have found no way to retrieve the ITextView from an ITextBuffer object
			Dim     OutliningManager	As  IOutliningManager		=  OutliningManagerService.GetOutliningManager ( textview )

			' Create persistent region data
			Dim	PersistentRegionData	As  PersistentRegionData	=  New  PersistentRegionData ( TextDocument.FilePath ) 

			' Create tag definitions
			Dim	TagDefinitions		As  TagDefinitions		=  New  TagDefinitions ( )

			TagDefinitions.Add ( "section"		, "endsection"		)
			TagDefinitions.Add ( "phpsection"	, "endphpsection"	)
			TagDefinitions.Add ( "phpregion"	, "endphpregion"	)
			TagDefinitions.Add ( "phpcode"		, "endphpcode"		) 

			' Save the listener and text document objects
			With ProviderInfo.OutliningTagger
				.Document		=  TextDocument
				.OutliningManager	=  OutliningManager
				.PersistentRegionData	=  PersistentRegionData
				.TagDefinitions		=  TagDefinitions

				' Finalize object creation and parse document
				.FinalizeObject ( Buffer )
				.ParseDocument ( True ) 
			End With

			' Return nothing for now, since we don't have to process any view tag (but maybe in the future)
			Return ( Nothing )
		Catch  E  As  Exception
			Utilities.ShowException ( E, ExceptionTitle )
			Return ( Nothing ) 
		End Try
	End Function
	#End Region
End Class
#End Region
