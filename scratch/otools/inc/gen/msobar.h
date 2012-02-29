///////////////////////////////////////////////////////////////////////////////
//
// msobar.h
// Generated file, do not tamper with the order of the items below!
//
///////////////////////////////////////////////////////////////////////////////

#ifndef __MSOBAR_H__
#define __MSOBAR_H__

#define msotbidUser                                           32768   // app-custom toolbars should use tbid >= msotbidUser
#define msotbidMin                                            0
#define msotbidNil                                            0   // Label: 'Nil'	 NOTE: Note: \"Menu\"=dropdowns, \"Popup\"=shortcut menus, \"Well\"=customize well
#define msotbidCustom                                         1   // Label: 'Custom'	 NOTE: used only for user-custom bars
#define msotbidGenericContainer                               2   // Label: 'MSO Generic Control Container'	 NOTE: used for GCC's
#define msotbidToolbarPopup                                   3   // Label: 'Toolbar List'	 NOTE: Warning: for toolbar internal use only
#define msotbidToolbarCustomize                               4   // Label: 'Customize'	 NOTE: Warning: for toolbar internal use only
#define msotbidAddCommand                                     5   // Label: 'Add Command'	 NOTE: Warning: for toolbar internal use only
#define msotbidModifySelection                                6   // Label: 'Modify Selection'	 NOTE: Warning: for toolbar internal use only
#define msotbidResetHideEtc                                   7   // Label: 'Reset or Hide'	 NOTE: Warning: for toolbar internal use only
#define msotbidCommandBars                                    8   // Label: 'Toolbars'	 NOTE: Warning: for toolbar internal use only
#define msotbidStandard                                       9   // Label: 'Standard'
#define msotbidFormatting                                     10   // Label: 'Formatting'
#define msotbidFormatBorder                                   11   // Label: 'Borders'
#define msotbidListMgmt                                       12   // Label: 'Database'
#define msotbidDrawing                                        13   // Label: 'Drawing'	 NOTE: If you really want it to say \"Drawing\", use msotbidDrawing2
#define msotbidFullScreen                                     14   // Label: 'Full Screen'
#define msotbidFormDesign                                     15   // Label: 'Forms'
#define msotbidEditPicture                                    16   // Label: 'Edit Picture'
#define msotbidMacro                                          17   // Label: 'Visual Basic'	 NOTE: PGM: OBSOLETE? (cf. msotbidVisualBasic)
#define msotbidMacroRecord                                    18   // Label: 'Macro Record'
#define msotbidMailMergeMain                                  19   // Label: 'Mail Merge'
#define msotbidLotusNotesFlow                                 20   // Label: 'NotesFlow'	 NOTE: NotesFlow feature
#define msotbidMicrosoft                                      21   // Label: 'Microsoft'
#define msotbidMasterPageview                                 22   // Label: 'Master Document'
#define msotbidOutline                                        23   // Label: 'Outlining'
#define msotbidPrintPreview                                   24   // Label: 'Print Preview'
#define msotbidWinword50                                      25   // Label: 'Word for Macintosh 5.0'
#define msotbidWinword20                                      26   // Label: 'Word for Windows 2.0'
#define msotbidMacWord51                                      27   // Label: 'Word for Macintosh 5.1'
#define msotbidMail                                           28   // Label: 'Mail'
#define msotbidPowerBook                                      29   // Label: 'PowerBook'
#define msotbidTipWizard                                      30   // Label: 'TipWizard'
#define msotbidHeaderFooter                                   31   // Label: 'Header and Footer'
#define msotbidLanguageMenu                                   32   // Label: 'Language'	 NOTE: Word
#define msotbidAllCommandsWell                                33   // Label: 'All Commands Well'	 NOTE: Word
#define msotbidPowerTalkMenu                                  34   // Label: 'PowerTalk'	 NOTE: Word (Mac specific)
#define msotbidMailEdRead                                     35   // Label: 'Read Mail'	 NOTE: Capone-Word mail
#define msotbidMailEdSend                                     36   // Label: 'Send Mail'	 NOTE: Capone-Word mail
#define msotbidMenuBarFull                                    37   // Label: 'Menu Bar'	 NOTE: the main menu bar
#define msotbidMenuBarSmall                                   38   // Label: 'Menu Bar (no document)'	 NOTE: no-doc menu bar; OBSOLETE
#define msotbidMenuBarEmerg                                   39   // Label: 'Menu Bar (low memory)'	 NOTE: emergency ?
#define msotbidFileMenu                                       40   // Label: 'File'
#define msotbidEditMenu                                       41   // Label: 'Edit'
#define msotbidViewMenu                                       42   // Label: 'View'
#define msotbidInsertMenu                                     43   // Label: 'Insert'
#define msotbidFormatMenu                                     44   // Label: 'Format'
#define msotbidToolsMenu                                      45   // Label: 'Tools'
#define msotbidTableMenu                                      46   // Label: 'Table'
#define msotbidWindowMenu                                     47   // Label: 'Window'
#define msotbidHelpMenu                                       48   // Label: 'Help'
#define msotbidDebugMenu                                      49   // Label: 'Debug'	 NOTE: debug version only -- for VB use msotcidVisualBasicDebugMenu
#define msotbidFileMenuSmall                                  50   // Label: 'File (no document)'	 NOTE: no-doc File menu; OBSOLETE
#define msotbidHelpMenuSmall                                  51   // Label: 'Help (no document)'	 NOTE: no-doc Help menu; OBSOLETE
#define msotbidDebugMenuSmall                                 52   // Label: 'Debug (no document)'	 NOTE: no-doc Debug menu; OBSOLETE
#define msotbidFileMenuEmerg                                  53   // Label: 'File (low memory)'	 NOTE: emergency ?
#define msotbidDrawnObjectPopup                               54   // Label: 'Shapes'	 NOTE: Escher - single or multiple shape selection
#define msotbidDropCapPopup                                   55   // Label: 'Drop Caps'
#define msotbidEdnPopup                                       56   // Label: 'Endnotes'
#define msotbidFieldPopup                                     57   // Label: 'Fields'
#define msotbidFieldDisplayPopup                              58   // Label: 'Display Fields'
#define msotbidFieldFormPopup                                 59   // Label: 'Form Fields'
#define msotbidFtnPopup                                       60   // Label: 'Footnotes'
#define msotbidFramePopup                                     61   // Label: 'Frames'
#define msotbidHeadingPopup                                   62   // Label: 'Headings'
#define msotbidHeadingLinkedPopup                             63   // Label: 'Linked Headings'
#define msotbidListPopup                                      64   // Label: 'Lists'
#define msotbidPictPopup                                      65   // Label: 'Pictures Context Menu'
#define msotbidTablePopup                                     66   // Label: 'Tables'
#define msotbidTableCellPopup                                 67   // Label: 'Table Cells'
#define msotbidHeadingTablePopup                              68   // Label: 'Table Headings'
#define msotbidListTablePopup                                 69   // Label: 'Table Lists'
#define msotbidPictTablePopup                                 70   // Label: 'Table Pictures'
#define msotbidTextTablePopup                                 71   // Label: 'Table Text'
#define msotbidTableWholePopup                                72   // Label: 'Whole Table'
#define msotbidTableWholeLinkedPopup                          73   // Label: 'Linked Table'
#define msotbidTextPopup                                      74   // Label: 'Text'
#define msotbidTextLinkedPopup                                75   // Label: 'Linked Text'
#define msotbidToolbarCustPopup                               76   // Label: 'Customize Shortcut'
#define msotbidRTEFontPopup                                   77   // Label: 'Font Popup'
#define msotbidRTEFontParaPopup                               78   // Label: 'Font Paragraph'
#define msotbidSpellPopup                                     79   // Label: 'Spelling'
#define msotbidFileWell                                       80   // Label: 'File Well'	 NOTE: Wells used in the Customize, Add Command dropdown
#define msotbidEditWell                                       81   // Label: 'Edit Well'
#define msotbidViewWell                                       82   // Label: 'View Well'
#define msotbidInsertWell                                     83   // Label: 'Insert Well'
#define msotbidFormattingWell                                 84   // Label: 'Format Well'
#define msotbidToolsWell                                      85   // Label: 'Tools Well'
#define msotbidTableWell                                      86   // Label: 'Table Well'
#define msotbidWindowWell                                     87   // Label: 'Window Well'
#define msotbidDrawingWell                                    88   // Label: 'Drawing Well'
#define msotbidBordersWell                                    89   // Label: 'Borders Well'
#define msotbidMailMergeWell                                  90   // Label: 'Mail Merge Well'
#define msotbidChart                                          91   // Label: 'Chart'
#define msotbidData                                           92   // Label: 'Query and Pivot'
#define msotbidAuditing                                       93   // Label: 'Formula Auditing'
#define msotbidWorkGroup                                      94   // Label: 'WorkGroup'
#define msotbidCustomBtn                                      95   // Label: 'Custom Button'	 NOTE: ?? is this being used?
#define msotbidFlowChartPalette                               96   // Label: 'Flow Chart'	 NOTE: Escher
#define msotbidOrgChartPalette                                97   // Label: 'Org Chart'	 NOTE: Escher
#define msotbidBasicPalette                                   98   // Label: 'Shapes'	 NOTE: Escher
#define msotbidMapPalette                                     99   // Label: 'Map'	 NOTE: Escher?
#define msotbidVisualBasic                                    100   // Label: 'Visual Basic'
#define msotbidStopRecording                                  101   // Label: 'Stop Recording'
#define msotbidGrammarPopup                                   102   // Label: 'Grammar'
#define msotbidMiniStatLine                                   103   // Label: 'MiniStatLine'	 NOTE: View buttons in lower left corner of Word
#define msotbidInsertMergeField                               104   // Label: 'Insert Merge Field'	 NOTE: Word - Menu for inserting a mail merge field
#define msotbidInsertMergeKeyword                             105   // Label: 'Insert Word Field'	 NOTE: Word - Menu for inserting a Word field code
#define msotbidMacroWell                                      106   // Label: 'Macro Well'	 NOTE: more well ID's
#define msotbidFontWell                                       107   // Label: 'Font Well'
#define msotbidAutoTextWell                                   108   // Label: 'AutoText Well'
#define msotbidStyleWell                                      109   // Label: 'Style Well'
#define msotbidChangeIconBar                                  110   // Label: 'Change Icon'	 NOTE: Warning: for toolbar internal use only
#define msotbidAddCommandOther                                111   // Label: 'Other'	 NOTE: Warning: for toolbar internal use only
#define msotbidCustomizeContextMenu                           112   // Label: 'Shortcut Menus'	 NOTE: Warning: for toolbar internal use only
#define msotbidDatabaseContainer                              113   // Label: 'Database'	 NOTE: 113-126 are for Access
#define msotbidRelationships                                  114   // Label: 'Relationships'
#define msotbidTableDesign                                    115   // Label: 'Table Design'
#define msotbidTableDatasheet                                 116   // Label: 'Table Datasheet'
#define msotbidQueryDesign                                    117   // Label: 'Query Design'
#define msotbidQueryDatasheet                                 118   // Label: 'Query Datasheet'
#define msotbidFormDesign2                                    119   // Label: 'Form Design'
#define msotbidFormView                                       120   // Label: 'Form View'
#define msotbidFilterSort                                     121   // Label: 'Filter/Sort'
#define msotbidReportDesign                                   122   // Label: 'Report Design'
#define msotbidFormattingForm                                 123   // Label: 'Formatting (Form/Report)'
#define msotbidFormattingDatasheet                            124   // Label: 'Formatting (Datasheet)'
#define msotbidUtility1                                       125   // Label: 'Utility 1'	 NOTE: an empty custom bar
#define msotbidUtility2                                       126   // Label: 'Utility 2'	 NOTE: another empty custom bar
#define msotbidToolbox                                        127   // Label: 'Toolbox'	 NOTE: use this one for OLE Controls
#define msotbidSectionMenu                                    128   // Label: 'Section'	 NOTE: for Binder
#define msotbidBinderHelpMenu                                 129   // Label: 'Binder Help'	 NOTE: when Binder help menu is on a cascade
#define msotbidSectionHelpMenu                                130   // Label: 'Help'	 NOTE: Binder puts app's help on cascade (override label)
#define msotbidLeftPanePopup                                  131   // Label: 'Left Pane'	 NOTE: Rt click menu for Binder's left pane
#define msotbidAutoText                                       132   // Label: 'AutoText'	 NOTE: Word
#define msotbidLetter                                         133   // Label: 'Letter'	 NOTE: Word
#define msotbidDrawMenu                                       134   // Label: 'Draw'	 NOTE: Escher - dropdown menu on Drawing bar
#define msotbidOrderMenu                                      135   // Label: 'Order'	 NOTE: Escher - cascade off of Draw menu
#define msotbidAlignMenu                                      136   // Label: 'Align'	 NOTE: Escher, Forms
#define msotbidDistributeMenu                                 137   // Label: 'Distribute'	 NOTE: Escher - cascade off of Draw menu
#define msotbidRotateFlipMenu                                 138   // Label: 'Rotate or Flip'	 NOTE: Escher - cascade off of Draw menu
#define msotbidTextEffects                                    139   // Label: 'WordArt'	 NOTE: Escher -
#define msotbidPictureToolbar                                 140   // Label: 'Picture'	 NOTE: Escher -
#define msotbidChangeShapeMenu                                141   // Label: 'Change Shape'	 NOTE: Escher - cascade to change the shape of a drawing
#define msotbidReviewComments                                 142   // Label: 'Reviewing'
#define msotbidBuiltInMenus                                   143   // Label: 'Built-in Menus'	 NOTE: used by Office customize
#define msotbidFieldAutoTextListPopup                         144   // Label: 'Field AutoText'	 NOTE: Word
#define msotbidReviewingMenu                                  145   // Label: 'Reviewing'	 NOTE: revision review (?)
#define msotbidFillColorMenu                                  146   // Label: 'Fill Color'	 NOTE: Escher -
#define msotbidLineColorMenu                                  147   // Label: 'Line Color'	 NOTE: Escher -
#define msotbidShadowMenu                                     148   // Label: 'Shadow'	 NOTE: Escher - shadow settings
#define msotbidLineWidthMenu                                  149   // Label: 'Line Width'	 NOTE: Escher -
#define msotbidLineDashMenu                                   150   // Label: 'Line Dash'	 NOTE: Escher -
#define msotbidArrowMenu                                      151   // Label: 'Arrow'	 NOTE: Escher -
#define msotbid3DEffectMenu                                   152   // Label: '3-D Effect'	 NOTE: Escher -
#define msotbidInsertGraphicMenu                              153   // Label: 'Insert Graphic'	 NOTE: Escher - main entry point on Std toolbar
#define msotbidLineToolMenu                                   154   // Label: 'Lines'	 NOTE: Escher - draws lines
#define msotbidQueryLayoutPopup                               155   // Label: 'Query Layout'	 NOTE: Excel - query in page break view
#define msotbidInlinePicturePopup                             156   // Label: 'Inline Picture'	 NOTE: Word - used for inline pictures
#define msotbidFloatingPicturePopup                           157   // Label: 'Floating Picture'	 NOTE: Word - used for floating pictures
#define msotbidBroadcastMenu                                  158   // Label: 'Broadcast'	 NOTE: PP broadcast
#define msotbidSynchronizeHTML                                159   // Label: 'Refresh'	 NOTE: Programmability - View/Source
#define msotbidPreviewMenu                                    160   // Label: 'Print/Web Preview'	 NOTE: Print Preview split dropdown
#define msotbidFramesetBorderPopup                            161   // Label: 'Frame Properties'	 NOTE: Word (see bug 29942)
#define msotbiddummy162                                       162   // Label: ' '
#define msotbidPhoneticEditPopup                              163   // Label: 'Phonetic Information'	 NOTE: Excel FE
#define msotbidTextShapeMenu                                  164   // Label: 'WordArt Shape'	 NOTE: Escher - these are dropdowns off TextEffects bar
#define msotbidTextEffectAlignmentMenu                        165   // Label: 'WordArt Alignment'	 NOTE: Escher - ditto
#define msotbidTextEffectTrackingMenu                         166   // Label: 'WordArt Tracking'	 NOTE: Escher - ditto
#define msotbidShadowSettings                                 167   // Label: 'Shadow Settings'	 NOTE: Escher - a \"More...\" item in Shadow
#define msotbidShadowColor                                    168   // Label: 'Shadow Color'	 NOTE: Escher - a \"More...\" item in Shadow
#define msotbid3DSettings                                     169   // Label: '3-D Settings'	 NOTE: Escher - a \"More...\" item in 3DEffect
#define msotbid3DExtrusionDepthMenu                           170   // Label: 'Extrusion Depth'	 NOTE: Escher - menu on 3DSettings bar
#define msotbid3DExtrusionDirectionMenu                       171   // Label: 'Extrusion Direction'	 NOTE: Escher -
#define msotbid3DExtrusionColor                               172   // Label: 'Extrusion Color'	 NOTE: Escher -
#define msotbid3DLightingMenu                                 173   // Label: '3-D Lighting'	 NOTE: Escher -
#define msotbid3DSurfaceMaterialMenu                          174   // Label: '3-D Surface Material'	 NOTE: Escher -
#define msotbidNudgeMenu                                      175   // Label: 'Nudge'	 NOTE: Escher -
#define msotbiddummy176                                       176   // Label: ' '
#define msotbidStandardShapesMenu                             177   // Label: 'AutoShapes'	 NOTE: Escher - menu on draw toolbar, contains common shapes
#define msotbidMoreShapesMenu                                 178   // Label: 'Basic Shapes'	 NOTE: Escher - menu with cascades to more shapes
#define msotbidMoreShapes1Menu                                179   // Label: 'Callouts'	 NOTE: Escher - groups of shapes; will get new label later
#define msotbidMoreShapes2Menu                                180   // Label: 'Flowchart'	 NOTE: Escher - ditto
#define msotbidMoreShapes3Menu                                181   // Label: 'Block Arrows'	 NOTE: Escher - ditto
#define msotbidMoreShapes4Menu                                182   // Label: 'Stars & Banners'	 NOTE: Escher - ditto
#define msotbiddummy183                                       183   // Label: ' '
#define msotbiddummy184                                       184   // Label: ' '
#define msotbiddummy185                                       185   // Label: ' '
#define msotbiddummy186                                       186   // Label: ' '
#define msotbiddummy187                                       187   // Label: ' '
#define msotbiddummy188                                       188   // Label: ' '
#define msotbidWeb                                            189   // Label: 'Web'	 NOTE: OfficeWeb toolbar
#define msotbidHistoryMenu                                    190   // Label: 'History'	 NOTE: OfficeWeb
#define msotbidFavoritesMenu                                  191   // Label: 'Favorites'	 NOTE: OfficeWeb
#define msotbidJournalMenu                                    192   // Label: 'Actions'	 NOTE: Ren
#define msotbidMDISystem                                      193   // Label: 'System'	 NOTE: Office - used for MDI child system menu when maximized
#define msotbidSnapToMenu                                     194   // Label: 'Snap To'	 NOTE: Escher
#define msotbidChangeShape0Menu                               195   // Label: 'Change AutoShape - Basic Shapes'	 NOTE: Escher
#define msotbidChangeShape1Menu                               196   // Label: 'Change AutoShape - Callouts'	 NOTE: Escher
#define msotbidChangeShape2Menu                               197   // Label: 'Change AutoShape - Flowchart'	 NOTE: Escher
#define msotbidChangeShape3Menu                               198   // Label: 'Change AutoShape - Block Arrows'	 NOTE: Escher
#define msotbidChangeShape4Menu                               199   // Label: 'Change AutoShape - Stars & Banners'	 NOTE: Escher
#define msotbidImageControlMenu                               200   // Label: 'Image Control'	 NOTE: Escher
#define msotbidTextWrapping                                   201   // Label: 'Text Wrapping'	 NOTE: Escher
#define msotbidSingleColorGradient                            202   // Label: 'Single Color Gradient Fill'	 NOTE: Escher
#define msotbidTwoColorGradient                               203   // Label: 'Two Color Gradient Fill'	 NOTE: Escher
#define msotbidTwoColorGradient2                              204   // Label: 'Two Color Gradient Fill'	 NOTE: Escher - supposedly both TwoColorGradient bars are needed
#define msotbidPatternForegroundColor                         205   // Label: 'Pattern Foreground Color'	 NOTE: Escher
#define msotbidPatternBackgroundColor                         206   // Label: 'Pattern Background Color'	 NOTE: Escher
#define msotbidFormatDODlgLineColor                           207   // Label: 'Line Color'	 NOTE: Escher - the DODlg bars are menus inside the format obj dialog
#define msotbidFormatDODlgLineStyle                           208   // Label: 'Line Style'	 NOTE: Escher
#define msotbidFormatDODlgLineWidth                           209   // Label: 'Line Width'	 NOTE: Escher
#define msotbidFormatDODlgDashStyle                           210   // Label: 'Dash Style'	 NOTE: Escher
#define msotbidFormatDODlgConnectorStyle                      211   // Label: 'Connector Style'	 NOTE: Escher
#define msotbidFormatDODlgArrow1Style                         212   // Label: 'Arrow 1 Style'	 NOTE: Escher
#define msotbidFormatDODlgArrow1Size                          213   // Label: 'Arrow 1 Size'	 NOTE: Escher
#define msotbidFormatDODlgArrow2Style                         214   // Label: 'Arrow 2 Style'	 NOTE: Escher
#define msotbidFormatDODlgArrow2Size                          215   // Label: 'Arrow 2 Size'	 NOTE: Escher
#define msotbidTFCPopup                                       216   // Label: 'Assistant'	 NOTE: the shortcut menu of the assistant
#define msotbidExternalData                                   217   // Label: 'External Data'	 NOTE: XL
#define msotbidAsyncRefresh                                   218   // Label: 'Asynchronous Refresh'	 NOTE: XL
#define msotbidWholeDocumentPopup                             219   // Label: 'Document'	 NOTE: all apps - right-click the title bar icon for an open doc.
#define msotbidPivotTable                                     220   // Label: 'PivotTable'	 NOTE: XL
#define msotbidPivotTableSelectMenu                           221   // Label: 'Select'	 NOTE: XL
#define msotbidPivotTableGroup                                222   // Label: 'Group and Outline'	 NOTE: XL
#define msotbidPivotTableInsert                               223   // Label: 'Insert'	 NOTE: XL
#define msotbidNewCustomMenu                                  224   // Label: 'new menu'	 NOTE: Office internal use only
#define msotbidChartTypeMenu                                  225   // Label: 'Chart Type'	 NOTE: XL - the list of chart types
#define msotbidCellPatternMenu                                226   // Label: 'Cell Pattern'	 NOTE: XL OBSOLETE
#define msotbidDlgColor                                       227   // Label: 'Color'	 NOTE: the color dropdowns inside the formatting dialogs
#define msotbidDlgPattern                                     228   // Label: 'Pattern'	 NOTE: ditto
#define msotbiddummy229                                       229   // Label: ' '
#define msotbiddummy230                                       230   // Label: ' '
#define msotbiddummy231                                       231   // Label: ' '
#define msotbidCurvePopup                                     232   // Label: 'Curve'	 NOTE: Escher
#define msotbidCurveNodePopup                                 233   // Label: 'Curve Node'	 NOTE: Escher
#define msotbidCurveSegmentPopup                              234   // Label: 'Curve Segment'	 NOTE: Escher
#define msotbidOLEObjectPopup                                 235   // Label: 'OLE Object'	 NOTE: Escher (?)
#define msotbidConnectorPopup                                 236   // Label: 'Connector'	 NOTE: Escher
#define msotbidTextEffectPopup                                237   // Label: 'WordArt Context Menu'	 NOTE: Escher
#define msotbidRotateModePopup                                238   // Label: 'Rotate Mode'	 NOTE: Escher - used when app is in rotate mode
#define msotbidRightDragDropPopup                             239   // Label: 'Nondefault Drag and Drop'	 NOTE: use this popup on a non-default drag and drop operation
#define msotbidHorizontalSpacingMenu                          240   // Label: 'Horizontal Spacing'	 NOTE: Forms
#define msotbidVerticalSpacingMenu                            241   // Label: 'Vertical Spacing'	 NOTE: Forms
#define msotbidCenterInFormMenu                               242   // Label: 'Center In Form'	 NOTE: Forms
#define msotbidArrangeButtonsMenu                             243   // Label: 'Arrange Buttons'	 NOTE: Forms
#define msotbidControlToolbox                                 244   // Label: 'Control Toolbox'	 NOTE: Forms - toolbar with OCX controls
#define msotbidHyperlinkMenu                                  245   // Label: 'Hyperlink Menu'	 NOTE: Web
#define msotbidSendToMenu                                     246   // Label: 'Send To'	 NOTE: Send doc somewhere
#define msotbidFixSpellingAutoCorrectMenu                     247   // Label: 'AutoCorrect'	 NOTE: Word
#define msotbidAutoManagerMenu                                248   // Label: 'AutoManager'	 NOTE: Word
#define msotbidSynonymsMenu                                   249   // Label: 'Synonyms'	 NOTE: Word
#define msotbidFontMenu                                       250   // Label: 'Font'	 NOTE: Word - for the Mac
#define msotbidWorkMenu                                       251   // Label: 'Work'	 NOTE: Word - for Mac, like in Word 5.1
#define msotbidGetExternalDataMenu                            252   // Label: 'Get External Data'	 NOTE: Access
#define msotbidChangeControlToMenu                            253   // Label: 'Change To'	 NOTE: Access
#define msotbidSQLSpecificMenu                                254   // Label: 'SQL Specific'	 NOTE: Access
#define msotbidFilterMenu                                     255   // Label: 'Filter Menu'	 NOTE: Access
#define msotbidSortMenu                                       256   // Label: 'Sort'	 NOTE: Access
#define msotbidAnalysisMenu                                   257   // Label: 'Analyze'	 NOTE: Access
#define msotbidSecurityMenu                                   258   // Label: 'Security'	 NOTE: Access
#define msotbidReplicationMenu                                259   // Label: 'Replication'	 NOTE: Access
#define msotbidOfficeLinksMenu                                260   // Label: 'OfficeLinks Menu'	 NOTE: Access
#define msotbidDatabaseObjectsMenu                            261   // Label: 'Database Objects'	 NOTE: Access
#define msotbidZoomMenu                                       262   // Label: 'Zoom'	 NOTE: Access
#define msotbidPagesMenu                                      263   // Label: 'Pages'	 NOTE: Access
#define msotbidArrangeIconsMenu                               264   // Label: 'Arrange Icons'	 NOTE: Access
#define msotbidWorksheetMenuBar                               265   // Label: 'Worksheet Menu Bar'	 NOTE: Excel - one of it's menu bars
#define msotbidChartMenuBar                                   266   // Label: 'Chart Menu Bar'	 NOTE: Excel
#define msotbidTextDirectionMenu                              267   // Label: 'Paragraph Direction'
#define msotbidEscherRightDragDropPopup                       268   // Label: 'Non-default drag-and-drop'	 NOTE: Escher only; use msotbidRightDragDropPopup in apps
#define msotbidBiDiFormatting                                 269   // Label: 'BiDi Formatting'	 NOTE: extra formatting toolbar used by Middle East apps
#define msotbidNewDatabaseObjectMenu                          270   // Label: 'New Object'	 NOTE: Access - creates a new db object
#define msotbidDocumentViewsMenu                              271   // Label: 'Views'	 NOTE: Access - switches between run, design, and other views
#define msotbidQueryTypeMenu                                  272   // Label: 'Query Type'	 NOTE: Access
#define msotbidDatabaseAnalysisMenu                           273   // Label: 'Analysis'	 NOTE: Access - database analysis options
#define msotbidOfficeLinksSplitMenu                           274   // Label: 'OfficeLinks Button Menu'	 NOTE: Access
#define msotbidSpecialEffectMenu                              275   // Label: 'Special Effect'	 NOTE: Access - chiseled, embossed, etc.
#define msotbidDatasheetBackColorMenu                         276   // Label: 'Fill/Back Color'	 NOTE: Access - the \"datasheet\" ID's exist to distinguish from the normal formatting ID's for controls
#define msotbidDatasheetTextColorMenu                         277   // Label: 'Font/Fore Color'	 NOTE: Access
#define msotbidDatasheetGridColorMenu                         278   // Label: 'Line/Border Color'	 NOTE: Access
#define msotbidDatasheetSpecialEffectMenu                     279   // Label: 'Datasheet Special Effect'	 NOTE: Access
#define msotbidDatasheetGridlinesMenu                         280   // Label: 'Gridlines'	 NOTE: Access
#define msotbidAllTablesWell                                  281   // Label: 'All Tables Well'	 NOTE: Access - one of the customize wells
#define msotbidAllQueriesWell                                 282   // Label: 'All Queries Well'	 NOTE: Access
#define msotbidAllFormsWell                                   283   // Label: 'All Forms Well'	 NOTE: Access
#define msotbidAllReportsWell                                 284   // Label: 'All Reports Well'	 NOTE: Access
#define msotbidAllMacrosWell                                  285   // Label: 'All Macros Well'	 NOTE: Access
#define msotbidAllModulesWell                                 286   // Label: 'All Modules Well'	 NOTE: Access
#define msotbidDatabaseContainerPopup                         287   // Label: 'Database TitleBar'	 NOTE: Access - used when right-click outside the object list
#define msotbidTableDesignGeneralPopup                        288   // Label: 'Table DesignTitleBar'
#define msotbidTableDesignUpperPanePopup                      289   // Label: 'Table Design Upper Pane'
#define msotbidTableDesignLowerPanePopup                      290   // Label: 'Table Design Lower Pane'
#define msotbidIndexGeneralPopup                              291   // Label: 'Index TitleBar'
#define msotbidIndexUpperPanePopup                            292   // Label: 'Index Upper Pane'
#define msotbidIndexLowerPanePopup                            293   // Label: 'Index Lower Pane'
#define msotbidRelationshipGeneralPopup                       294   // Label: 'Relationship'
#define msotbidRelationshipColumnListPopup                    295   // Label: 'Relationship TableFieldList'
#define msotbidRelationshipColumnListQueryPopup               296   // Label: 'Relationship QueryFieldList'
#define msotbidRelationshipJoinPopup                          297   // Label: 'Relationship Join'
#define msotbidTableDesignDatasheetPopup                      298   // Label: 'Table Design Datasheet'
#define msotbidTableDesignDatasheetColumnPopup                299   // Label: 'Table Design Datasheet Column'
#define msotbidTableDesignDatasheetRowPopup                   300   // Label: 'Table Design Datasheet Row'
#define msotbidTableDesignDatasheetCellPopup                  301   // Label: 'Table Design Datasheet Cell'
#define msotbidQueryDesignGeneralPopup                        302   // Label: 'Query DesignGeneral'
#define msotbidQueryDesignColumnListPopup                     303   // Label: 'Query DesignFieldList'
#define msotbidQueryDesignJoinPopup                           304   // Label: 'Query Design Join'
#define msotbidQueryDesignQBEPopup                            305   // Label: 'Query DesignGrid'
#define msotbidQueryDesignSQLPopup                            306   // Label: 'Query SQLTitleBar'
#define msotbidQueryDesignSQLTextPopup                        307   // Label: 'Query SQLText'
#define msotbidQueryDesignPropertiesPopup                     308   // Label: 'Query Design Properties'
#define msotbidQueryDesignDatasheetPopup                      309   // Label: 'Query Design Datasheet'
#define msotbidQueryDesignDatasheetColumnPopup                310   // Label: 'Query Design Datasheet Column'
#define msotbidQueryDesignDatasheetRowPopup                   311   // Label: 'Query Design Datasheet Row'
#define msotbidQueryDesignDatasheetCellPopup                  312   // Label: 'Query Design Datasheet Cell'
#define msotbidFilterGeneralPopup                             313   // Label: 'Filter General Context Menu'
#define msotbidFilterQBEPopup                                 314   // Label: 'Filter Field'
#define msotbidQuickFilterGeneralPopup                        315   // Label: 'Filter FilterByForm'
#define msotbidFormDesignGeneralPopup                         316   // Label: 'Form DesignTitleBar'
#define msotbidFormDesignFormPopup                            317   // Label: 'Form Design Form'
#define msotbidFormDesignSectionPopup                         318   // Label: 'Form Design Section'
#define msotbidFormDesignControlPopup                         319   // Label: 'Form Design Control'
#define msotbidFormDesignControlLabelPopup                    320   // Label: 'Form Design Control Label'
#define msotbidFormDesignControlLabelXPopup                   321   // Label: 'Form Design Control Label'
#define msotbidFormDesignControlOLEPopup                      322   // Label: 'Form Design Control OLE'
#define msotbidReportDesignGeneralPopup                       323   // Label: 'Report DesignTitleBar'
#define msotbidReportDesignFormPopup                          324   // Label: 'Report DesignReport'
#define msotbidReportDesignSectionPopup                       325   // Label: 'Report Design Section'
#define msotbidReportDesignControlPopup                       326   // Label: 'Report Design Control'
#define msotbidFormReportDesignPropertyPopup                  327   // Label: 'Form/Report Properties'
#define msotbidFormViewGeneralPopup                           328   // Label: 'Form View Popup'
#define msotbidFormViewSubformGeneralPopup                    329   // Label: 'Form View Subform'
#define msotbidFormViewControlPopup                           330   // Label: 'Form View Control'
#define msotbidFormViewSubformControlPopup                    331   // Label: 'Form View Subform Control'
#define msotbidFormViewRecordPopup                            332   // Label: 'Form View Record'
#define msotbidFormDatasheetGeneralPopup                      333   // Label: 'Form Datasheet'
#define msotbidFormDatasheetColumnPopup                       334   // Label: 'Form Datasheet Column'
#define msotbidFormDatasheetSubColumnPopup                    335   // Label: 'Form Datasheet Subcolumn'
#define msotbidFormDatasheetRowPopup                          336   // Label: 'Form Datasheet Row'
#define msotbidFormDatasheetCellPopup                         337   // Label: 'Form Datasheet Cell'
#define msotbidPrintPreviewPopup                              338   // Label: 'Print Preview Popup'	 NOTE: Access - may be suitable for other apps, if needed
#define msotbidMacroDesignGeneralPopup                        339   // Label: 'Macro TitleBar'
#define msotbidMacroDesignUpperPanePopup                      340   // Label: 'Macro UpperPane'
#define msotbidMacroDesignConditionPopup                      341   // Label: 'Macro Condition'
#define msotbidMacroDesignArgumentPopup                       342   // Label: 'Macro Argument'
#define msotbidModuleWaitModePopup                            343   // Label: 'ModuleUncompiled'
#define msotbidModuleBreakModePopup                           344   // Label: 'ModuleCompiled'
#define msotbidModuleImmediatePanePopup                       345   // Label: 'Module Immediate'
#define msotbidModuleImmediatePaneWatchPopup                  346   // Label: 'Module Watch'
#define msotbidDatabaseContainerTableQueryPopup               347   // Label: 'Database Table/Query'
#define msotbidDatabaseContainerFormPopup                     348   // Label: 'Database Form'
#define msotbidDatabaseContainerReportPopup                   349   // Label: 'Database Report'
#define msotbidDatabaseContainerMacroPopup                    350   // Label: 'Database Macro'
#define msotbidDatabaseContainerModulePopup                   351   // Label: 'Database Module'
#define msotbidDatabaseContainerBackgroundPopup               352   // Label: 'Database Background'
#define msotbidAccessOLESharedPopup                           353   // Label: 'OLE Shared'
#define msotbidAccessGlobalPopup                              354   // Label: 'Global'
#define msotbidUtilitiesMenu                                  355   // Label: 'Utilities'	 NOTE: Access
#define msotbidViewDatabaseObjectsMenu                        356   // Label: 'Database Objects'	 NOTE: Access - switches tabs in the DBC
#define msotbidGoToMenu                                       357   // Label: 'Go To'	 NOTE: Access
#define msotbidEditObjectMenu                                 358   // Label: 'Edit Object'	 NOTE: Edit object cascade menu
#define msotbidArrangeIconsPopup                              359   // Label: 'Arrange Icons'	 NOTE: Access
#define msotbidNumberOfPagesMenu                              360   // Label: 'Pages'	 NOTE: Access - how many pages to view
#define msotbidAccessFilterPopup                              361   // Label: 'Filter Context Menu'	 NOTE: Access
#define msotbidDatasheetWell                                  362   // Label: 'Datasheet Well'	 NOTE: Access
#define msotbidRecordsWell                                    363   // Label: 'Records Well'	 NOTE: Access
#define msotbidQueryDesignWell                                364   // Label: 'Query Design Well'	 NOTE: Access
#define msotbidFormReportDesignWell                           365   // Label: 'Form/Report Design Well'	 NOTE: Access
#define msotbidModuleDesignWell                               366   // Label: 'Module Design Well'	 NOTE: Access
#define msotbidTableDesignPropertiesPopup                     367   // Label: 'Table Design Properties'	 NOTE: Access
#define msotbidIndexPropertiesPopup                           368   // Label: 'Index Properties'	 NOTE: Access
#define msotbidGetExternalDataPopup                           369   // Label: 'Get External Data'	 NOTE: Access
#define msotbidChangeControlTypeMenu                          370   // Label: 'Change To'	 NOTE: Access - morphs controls into other types
#define msotbidSizeMenu                                       371   // Label: 'Size'	 NOTE: Access
#define msotbidAddinsMenu                                     372   // Label: 'Add-ins'	 NOTE: Access
#define msotbidRemote                                         373   // Label: 'Remote'	 NOTE: Ren - remote mode
#define msotbidComposeMenu                                    374   // Label: 'Actions'	 NOTE: Ren
#define msotbiddummy375                                       375   // Label: ' '
#define msotbidContactMenu                                    376   // Label: 'Actions'	 NOTE: Ren
#define msotbidNoteMenu                                       377   // Label: 'Actions'	 NOTE: Ren
#define msotbidTaskMenu                                       378   // Label: 'Actions'	 NOTE: Ren
#define msotbidSupernoteMenu                                  379   // Label: 'Form Design'	 NOTE: Ren
#define msotbiddummy391                                       380   // Label: ' '
#define msotbiddummy381                                       381   // Label: ' '
#define msotbiddummy382                                       382   // Label: ' '
#define msotbiddummy383                                       383   // Label: ' '
#define msotbidSheetFileMenu                                  384   // Label: 'File'	 NOTE: Excel
#define msotbidFontColor                                      385   // Label: 'Font Color'	 NOTE: ?
#define msotbidPrintSetupMenu                                 386   // Label: 'Print Setup'	 NOTE: File submenu, contains printing commands
#define msotbiddummy387                                       387   // Label: ' '
#define msotbidSendMailMenuMac                                388   // Label: 'Mail'	 NOTE: Mac version of mail submenu
#define msotbidFillMenu                                       389   // Label: 'Fill'	 NOTE: Excel - Edit submenu
#define msotbidClearMenu                                      390   // Label: 'Clear'	 NOTE: Excel - Edit submenu
#define msotbidPublishSubscribeMenu                           391   // Label: 'Publishing'	 NOTE: Excel - Edit submenu (Mac only)
#define msotbidObjectMenu                                     392   // Label: 'Object'	 NOTE: Edit submenu, contains list of OLE object verbs
#define msotbidCommentsMenu                                   393   // Label: 'Comments'	 NOTE: View submenu, contains comment/annotation commands
#define msotbidInsertMacroMenu                                394   // Label: 'Macro'	 NOTE: Excel - Insert submenu with macro commands
#define msotbidNameMenu                                       395   // Label: 'Names'	 NOTE: Excel - Insert submenu with name/bookmark commands
#define msotbidFormatRowMenu                                  396   // Label: 'Row'	 NOTE: Excel - format submenu
#define msotbidFormatColumnMenu                               397   // Label: 'Column'	 NOTE: Excel - format submenu
#define msotbidPageFormatMenu                                 398   // Label: 'Page Format'	 NOTE: Excel - format submenu
#define msotbidPhoneticMenu                                   399   // Label: 'Phonetic Guide'	 NOTE: Excel - used in FE for phonetic properties
#define msotbidConditionsMenu                                 400   // Label: 'Conditions'	 NOTE: Excel
#define msotbidObjectPlacementMenu                            401   // Label: 'Object Placement Menu'	 NOTE: submenu with bring to front, send to back, etc.
#define msotbidRevisionsMenu                                  402   // Label: 'Track Changes Menu'
#define msotbidAuditingMenu                                   403   // Label: 'Formula Auditing'	 NOTE: Excel
#define msotbidProtectionMenu                                 404   // Label: 'Protection'	 NOTE: Excel
#define msotbidDataValidationMenu                             405   // Label: 'Validation'	 NOTE: Excel
#define msotbidOutlineMenu                                    406   // Label: 'Outline'	 NOTE: Excel
#define msotbidChartFileMenu                                  407   // Label: 'File'	 NOTE: Excel - these are the menus on the Chart menu bar
#define msotbidChartEditMenu                                  408   // Label: 'Edit'	 NOTE: Excel
#define msotbidChartViewMenu                                  409   // Label: 'View'
#define msotbidChartInsertMenu                                410   // Label: 'Insert'
#define msotbidChartFormatMenu                                411   // Label: 'Format'
#define msotbidChartToolsMenu                                 412   // Label: 'Tools'
#define msotbidChartWindowMenu                                413   // Label: 'Window'
#define msotbidChartHelpMenu                                  414   // Label: 'Help'
#define msotbidChartSendMailMenu                              415   // Label: 'Mail'
#define msotbidChartClearMenu                                 416   // Label: 'Clear'
#define msotbidChartPublishSubscribeMenu                      417   // Label: 'Publsh/Subscribe'
#define msotbidChartInsertMacroMenu                           418   // Label: 'Macro'
#define msotbidChartPageFormatMenu                            419   // Label: 'Page Format'
#define msotbidChartObjectPlacementMenu                       420   // Label: 'Object Placement Menu (Chart)'
#define msotbidChartRevisionsMenu                             421   // Label: 'Track Changes Menu (chart)'
#define msotbidChartProtectionMenu                            422   // Label: 'Protection'
#define msotbidAllPliesPopup                                  423   // Label: 'Workbook tabs'	 NOTE: Excel - pops up list of all plies in current workbook
#define msotbidGridCellPopup                                  424   // Label: 'Cell'	 NOTE: use this on cells in grids
#define msotbidGridRowPopup                                   425   // Label: 'Row'	 NOTE: use this on whole rows in grids
#define msotbidGridColumnPopup                                426   // Label: 'Column'	 NOTE: use this on whole columns in grids
#define msotbidGridCellLayoutPopup                            427   // Label: 'Cell'	 NOTE: Excel - used when in Layout View
#define msotbidGridRowLayoutPopup                             428   // Label: 'Row'	 NOTE: Excel
#define msotbidGridColumnLayoutPopup                          429   // Label: 'Column'	 NOTE: Excel
#define msotbidWorkbookPlyPopup                               430   // Label: 'Ply'	 NOTE: Excel - popup menu for a workbook ply
#define msotbidXLMMacroCellPopup                              431   // Label: 'XLM Cell'	 NOTE: Excel - cells in XLM plies
#define msotbidMDIDesktopPopup                                432   // Label: 'Desktop'	 NOTE: Excel - popup menu for empty desktop area
#define msotbidAutoFillPopup                                  433   // Label: 'AutoFill'	 NOTE: Excel - used for non-default AutoFill
#define msotbidFormsButtonPopup                               434   // Label: 'Button'	 NOTE: Excel - used for native button controls
#define msotbidFormsTextboxPopup                              435   // Label: 'Textbox'	 NOTE: Excel
#define msotbidDialogPlyPopup                                 436   // Label: 'Dialog'	 NOTE: Excel - generic dialog design commands
#define msotbidSeriesPopup                                    437   // Label: 'Series'	 NOTE: Chart
#define msotbidPlotAreaPopup                                  438   // Label: 'Plot Area'	 NOTE: Chart
#define msotbidFloorAndWallsPopup                             439   // Label: 'Floor and Walls'	 NOTE: Chart
#define msotbidTrendlinePopup                                 440   // Label: 'Trendline'	 NOTE: Chart
#define msotbidChartGeneralPopup                              441   // Label: 'Chart'	 NOTE: Chart - generic popup when no selection
#define msotbidFormulaBarPopup                                442   // Label: 'Formula Bar'	 NOTE: Excel - used to edit contents of formula bar
#define msotbidPivotTablePopup                                443   // Label: 'PivotTable Context Menu'	 NOTE: Excel
#define msotbidQueryPopup                                     444   // Label: 'Query'	 NOTE: Excel
#define msotbidAutoCalculatePopup                             445   // Label: 'AutoCalculate'	 NOTE: Excel - list of functions for status bar AlwaysCalc feature
#define msotbidChartObjectPlotPopup                           446   // Label: 'Object/Plot'	 NOTE: Chart - \"merged chart object and chart area context menu\"
#define msotbidChartTitleBarPopup                             447   // Label: 'Title Bar (Charting)'	 NOTE: Chart
#define msotbidDataMenu                                       448   // Label: 'Data'	 NOTE: Excel
#define msotbidRecordMacroMenu                                449   // Label: 'Record Macro'	 NOTE: Excel (any other apps?)
#define msotbidEditOLEObjectMenu                              450   // Label: 'Object'	 NOTE: the submenu off Edit with the object's verbs, and Convert...
#define msotbidAttachmentsShortcutsMenu                       451   // Label: 'View Attachments'	 NOTE: Ren
#define msotbidFolderMenu                                     452   // Label: 'Folder'	 NOTE: Ren
#define msotbidMessageMenu                                    453   // Label: 'Message'	 NOTE: Ren
#define msotbidNewItemMenu                                    454   // Label: 'New Item'	 NOTE: Ren and Word New split dropdown
#define msotbidViewNextMenu                                   455   // Label: 'Next'	 NOTE: Ren
#define msotbidViewPreviousMenu                               456   // Label: 'Previous'	 NOTE: Ren
#define msotbidViewNumberOfDaysMenu                           457   // Label: 'Number of Days'	 NOTE: Ren
#define msotbidAppointmentMenu                                458   // Label: 'Actions'	 NOTE: Ren
#define msotbidItemColorMenu                                  459   // Label: 'Color'	 NOTE: Ren
#define msotbidCurrentViewMenu                                460   // Label: 'Current View'	 NOTE: Ren
#define msotbidOpenSharedFolderMenu                           461   // Label: 'Open Shared Folder'	 NOTE: Ren
#define msotbidSaveAttachmentsMenu                            462   // Label: 'Save Attachments'	 NOTE: Ren
#define msotbidTaskPadSettingsMenu                            463   // Label: 'TaskPad Settings'	 NOTE: Ren
#define msotbidTaskPadFilterMenu                              464   // Label: 'TaskPad Filter'	 NOTE: Ren
#define msotbidPatternMenu                                    465   // Label: 'Pattern'	 NOTE: XL, Escher
#define msotbidWebOptionsMenu                                 466   // Label: 'Options'	 NOTE: Web - contains set start page & set search page
#define msotbidLineStyle                                      467   // Label: 'Line/Border Style'	 NOTE: Access
#define msotbidBorderColor                                    468   // Label: 'Line/Border Color'	 NOTE: Access
#define msotbidRun                                            469   // Label: 'Run'	 NOTE: Access
#define msotbidOLEControlsWell                                470   // Label: 'ActiveX Controls Well'	 NOTE: Access
#define msotbidShortcutMenuBar                                471   // Label: 'Shortcut Menus'	 NOTE: apps put their shortcut menus on this bar for customization
#define msotbidRecords                                        472   // Label: 'Records'	 NOTE: Access
#define msotbidFormatSheetMenu                                473   // Label: 'Sheet'	 NOTE: Excel
#define msotbidGroupAndOutlineMenu                            474   // Label: 'Group and Outline'	 NOTE: Excel
#define msotbidDlgColorNone                                   475   // Label: 'Color'	 NOTE: color dropdown with \"None\" instead of \"Automatic\"
#define msotbidAccess2Macro                                   476   // Label: 'Macro Design'	 NOTE: Access (cf. msotbidMacro, which says \"Visual Basic\")
#define msotbidDrawingObjectDefaultsMenu                      477   // Label: 'dummy'	 NOTE: UNUSED
#define msotbidContactsMenu                                   478   // Label: 'Contacts'	 NOTE: Ren
#define msotbidCalendarMenu                                   479   // Label: 'Calendar'	 NOTE: Ren
#define msotbidArrangeMenu                                    480   // Label: 'Arrange'	 NOTE: Ren
#define msotbidSpeedDialMenu                                  481   // Label: 'Speed Dial'	 NOTE: Ren
#define msotbidFormMenu                                       482   // Label: 'Form'	 NOTE: Ren
#define msotbidLayoutMenu                                     483   // Label: 'Layout'	 NOTE: Ren
#define msotbidRedialMenu                                     484   // Label: 'Redial'	 NOTE: Ren
#define msotbidSizeToMenu                                     485   // Label: 'Size To'	 NOTE: Ren
#define msotbidTasksMenu                                      486   // Label: 'Tasks'	 NOTE: Ren
#define msotbidRemoteMenu                                     487   // Label: 'Remote'	 NOTE: Ren
#define msotbidDialPhoneMenu                                  488   // Label: 'Dial Phone'	 NOTE: Ren
#define msotbidPageSetupMenu                                  489   // Label: 'Page Setup'	 NOTE: Ren
#define msotbidVoting                                         490   // Label: 'Response'	 NOTE: Ren
#define msotbidRenFieldMenu                                   491   // Label: 'Field'	 NOTE: Ren
#define msotbidMoveMRUMenu                                    492   // Label: 'Move'	 NOTE: Ren
#define msotbidCharacterScalingMenu                           493   // Label: 'Character Scaling'	 NOTE: Word
#define msotbidTablesAndBorders                               494   // Label: 'Tables and Borders'	 NOTE: Word
#define msotbidInsertTextboxMenu                              495   // Label: 'Textbox'	 NOTE: Word - submenu off insert
#define msotbidInsertFrameMenu                                496   // Label: 'Frame'	 NOTE: Word - submenu off insert
#define msotbidVisualBasicDebugMenu                           497   // Label: 'Debug'	 NOTE: VBE
#define msotbidGenericContextMenu                             498   // Label: 'Context Menu'	 NOTE: Ren - in 97 they use one ID for all their context menus
#define msotbidNoteColorMenu                                  499   // Label: 'Note Color'	 NOTE: Ren
#define msotbidChartMenu                                      500   // Label: 'Chart'	 NOTE: Excel
#define msotbidFormatDODlgFillColor                           501   // Label: 'Fill Color'	 NOTE: Escher - used in a dialog
#define msotbidExitDesignMode                                 502   // Label: 'Exit Design Mode'	 NOTE: VBE - gets you out of form design mode
#define msotbidRenInsertFieldMenu                             503   // Label: 'Insert Field'	 NOTE: Ren
#define msotbidEmptyListMenu                                  504   // Label: 'Empty List'	 NOTE: Ren
#define msotbidNoteItemMenu                                   505   // Label: 'Note Item'	 NOTE: Ren
#define msotbidRecurrenceMenu                                 506   // Label: 'Recurrence'	 NOTE: Ren
#define msotbidFindAllMenu                                    507   // Label: 'Find All'	 NOTE: Ren
#define msotbidFontSizeMenu                                   508   // Label: 'Font Size'	 NOTE: Ren
#define msotbidManageRemoteMailMenu                           509   // Label: 'Manage Remote Mail'	 NOTE: Ren
#define msotbidDialMenu                                       510   // Label: 'Dial'	 NOTE: Ren
#define msotbidAlignDistributeMenu                            511   // Label: 'Align or Distribute'	 NOTE: Escher
#define msotbidTitleBarPopup                                  512   // Label: 'Title Bar'	 NOTE: when you right click the title bar
#define msotbidPivotTableSelectPopup                          513   // Label: 'PivotTable Select'	 NOTE: Excel
#define msotbidPivotTableGroupingPopup                        514   // Label: 'PivotTable Grouping'	 NOTE: Excel
#define msotbidPivotTableFormulasPopup                        515   // Label: 'PivotTable Formulas'	 NOTE: Excel
#define msotbidFormDesignBackColorMenu                        516   // Label: 'Fill/Back Color'	 NOTE: Access
#define msotbidFormDesignTextColorMenu                        517   // Label: 'Font/Fore Color'	 NOTE: Access
#define msotbidFormDesignGridColorMenu                        518   // Label: 'Grid Color'	 NOTE: Access
#define msotbidCommandBarPropertySheet                        519   // Label: 'Command Bar Property Sheet'	 NOTE: Access - the property sheet for setting advanced properties
#define msotbidGroupingMenu                                   520   // Label: 'Grouping'	 NOTE: Escher - cascade off context menu
#define msotbidTableDesignWell                                521   // Label: 'Table Design Well'	 NOTE: Access
#define msotbidToolboxWell                                    522   // Label: 'Toolbox Well'	 NOTE: Access
#define msotbidMacroDesignWell                                523   // Label: 'Macro Design Well'	 NOTE: Access
#define msotbidLocalsPopup                                    524   // Label: 'Module LocalsPane'	 NOTE: Access (VBE?)
#define msotbidObjectBrowserPopup                             525   // Label: 'Object Browser'	 NOTE: Access (VBE?)
#define msotbidRightClickNYI                                  526   // Label: 'NYI'	 NOTE: Access - not yet implemented placeholder
#define msotbidInternetMenu                                   527   // Label: 'Internet'	 NOTE: Access - used to publish to Web
#define msotbidSourceCodeControlMenu                          528   // Label: 'Source Code Control Menu'	 NOTE: Access - SourceSafe integration
#define msotbidInsertVBEObjectSplitMenu                       529   // Label: 'Insert VBE Object'	 NOTE: VBE - allows insertion of dialogs, modules, etc.
#define msotbidTextBox                                        530   // Label: 'Text Box'	 NOTE: Word - used for text box linking
#define msotbidSourceCodeControl                              531   // Label: 'Source Code Control'	 NOTE: Access - SourceSafe toolbar
#define msotbidExcelControlPopup                              532   // Label: 'Excel Control'	 NOTE: Excel
#define msotbidConnectorToolMenu                              533   // Label: 'Connectors'	 NOTE: Escher - appears near LineToolsMenu and StandardShapesMenu
#define msotbidSymbolPanel1                                   534   // Label: 'Symbol Panel 1'	 NOTE: Word internal use
#define msotbidSymbolPanel2                                   535   // Label: 'Symbol Panel 2'	 NOTE: Word internal use
#define msotbidExtendedFormatting                             536   // Label: 'Extended Formatting'	 NOTE: Word
#define msotbidAutoSummary                                    537   // Label: 'AutoSummarize'	 NOTE: Word
#define msotbidPictureMenu                                    538   // Label: 'Picture Menu'	 NOTE: Word, others?
#define msotbidInsertAutoTextMenu                             539   // Label: 'AutoText'	 NOTE: Word
#define msotbidAutoTextListMenu                               540   // Label: 'AutoText List'	 NOTE: Word
#define msotbidBorderColorMenu                                541   // Label: 'Border Color'	 NOTE: Word, others?
#define msotbidRevisionPopup                                  542   // Label: 'Track Changes'	 NOTE: Word, others?
#define msotbidCommentPopup                                   543   // Label: 'Comment'	 NOTE: Word, others?
#define msotbidRevisionMarkingStatusBarPopup                  544   // Label: 'Track Changes Indicator'	 NOTE: Word
#define msotbidHyperlinkPopup                                 545   // Label: 'Hyperlink Context Menu'	 NOTE: Web
#define msotbidDataWell                                       546   // Label: 'Data Well'	 NOTE: Excel - command well
#define msotbidWindowAndHelpWell                              547   // Label: 'Window and Help Well'	 NOTE: command well
#define msotbidChartingWell                                   548   // Label: 'Charting Well'	 NOTE: command well
#define msotbidFormsWell                                      549   // Label: 'Forms Well'	 NOTE: command well
#define msotbidMicrosoftOnTheWebMenu                          550   // Label: 'Microsoft on the Web'	 NOTE: Help menu cascade
#define msotbidCustomWell                                     551   // Label: 'Custom Well'	 NOTE: command well
#define msotbidAutoShapesWell                                 552   // Label: 'AutoShapes Well'	 NOTE: command well
#define msotbidControlToolboxWell                             553   // Label: 'Control Toolbox Well'	 NOTE: command well
#define msotbidPivotTableWell                                 554   // Label: 'PivotTable Well'	 NOTE: command well
#define msotbidWebWell                                        555   // Label: 'Web Well'	 NOTE: command well
#define msotbidCustomMainMenu                                 556   // Label: 'Menu Bar (custom)'	 NOTE: Access
#define msotbidCustomDropDownMenu                             557   // Label: 'Dropdown Menu (custom)'	 NOTE: Access
#define msotbidCustomPopup                                    558   // Label: 'Custom'	 NOTE: Access
#define msotbidTabControlPopup                                559   // Label: 'Tab Control'	 NOTE: Access
#define msotbidSourceCodeControlWell                          560   // Label: 'Source Code Control Well'	 NOTE: Access - SourceSafe integration
#define msotbidChartWholeDocumentPopup                        561   // Label: 'Whole Document (Chart)'	 NOTE: Excel
#define msotbidInsertMoviesAndSoundsMenu                      562   // Label: 'Movies and Sounds'	 NOTE: PPT
#define msotbidMasterMenu                                     563   // Label: 'Master Menu'	 NOTE: PPT
#define msotbidSlideShowScreenMenu                            564   // Label: 'Screen'	 NOTE: PPT - Slide show popup cascade
#define msotbidSlideShowMenu                                  565   // Label: 'Slide Show'	 NOTE: PPT
#define msotbidSlideShowGoToMenu                              566   // Label: 'Slide Show Go To'	 NOTE: PPT - Slide show popup cascade
#define msotbidSlideShowPointerOptionsMenu                    567   // Label: 'Slide Show Pointer Options'	 NOTE: PPT - slide show popup cascade
#define msotbidSlideShowPenColorMenu                          568   // Label: 'Slide Show Pen Color'	 NOTE: PPT - slide show popup cascade
#define msotbidCommonTasks                                    569   // Label: 'Common Tasks'	 NOTE: PPT
#define msotbidSlideShowMenuBar                               570   // Label: 'Menu Bar (Slide Show)'	 NOTE: PPT
#define msotbidSlideSorter                                    571   // Label: 'Slide Sorter'	 NOTE: PPT
#define msotbidAnimationEffects                               572   // Label: 'Animation Effects'	 NOTE: PPT
#define msotbidSlideShowPopup                                 573   // Label: 'Slide Show'	 NOTE: PPT
#define msotbidSlideShowFileMenu                              574   // Label: 'File'	 NOTE: PPT
#define msotbidSlideShowEditMenu                              575   // Label: 'Edit'	 NOTE: PPT
#define msotbidSlideShowGoMenu                                576   // Label: 'Browse'	 NOTE: PPT
#define msotbidSlideShowWindowMenu                            577   // Label: 'Window'	 NOTE: PPT
#define msotbidSlideMaster                                    578   // Label: 'Slide Master View'	 NOTE: PPT
#define msotbidHandoutMaster                                  579   // Label: 'Handout Master View'	 NOTE: PPT
#define msotbidSlideMiniaturePopup                            580   // Label: 'Slide Miniature'	 NOTE: PPT
#define msotbidSlideSorterPopup                               581   // Label: 'Slider Sorter'	 NOTE: PPT
#define msotbidSlideOutlinerPopup                             582   // Label: 'Outliner'	 NOTE: PPT
#define msotbidTracking                                       583   // Label: 'Tracking'	 NOTE: Project
#define msotbidResourceManagement                             584   // Label: 'Resource Management'	 NOTE: Project
#define msotbidUtility                                        585   // Label: 'Utility'	 NOTE: Project
#define msotbidCustomForms                                    586   // Label: 'Custom Forms'	 NOTE: Project
#define msotbidProject3                                       587   // Label: 'Microsoft Project 3.0'	 NOTE: Project
#define msotbidOutMailEdSend                                  588   // Label: 'Outlook Send Mail'	 NOTE: Outlook-Word mail
#define msotbidOutMailEdRead                                  589   // Label: 'Outlook Read Mail'	 NOTE: Outlook-Word mail
#define msotbidChartPictureMenu                               590   // Label: 'Picture Menu (Chart)'	 NOTE: Excel
#define msotbidPivotTableLayout                               591   // Label: 'Layout'	 NOTE: Excel
#define msotbidBookmarkMenu                                   592   // Label: 'Bookmark'	 NOTE: Access
#define msotbidDrawing2                                       593   // Label: 'Drawing'	 NOTE: Project, other non-Escher apps
#define msotbidPivotTableGroupMenu                            594   // Label: 'Group and Outline'	 NOTE: Excel
#define msotbidPivotTableMenu                                 595   // Label: 'PivotTable Menu'	 NOTE: Excel
#define msotbidPivotTableFormulaMenu                          596   // Label: 'Formula'	 NOTE: Excel
#define msotbidPrintAreaMenu                                  597   // Label: 'Print Area'	 NOTE: Excel
#define msotbidOCXPopup                                       598   // Label: 'ActiveX Control'	 NOTE: Escher
#define msotbidChangeShape5Menu                               599   // Label: 'Change AutoShape - Action Buttons'	 NOTE: Escher
#define msotbidWebGoMenu                                      600   // Label: 'Go'	 NOTE: Web - top-level menu on menu bar for web navigation
#define msotbidMoreShapes5Menu                                601   // Label: 'Action Buttons'	 NOTE: Escher
#define msotbidAlignmentMenu                                  602   // Label: 'Alignment'	 NOTE: use this for text alignment
#define msotbidBlackWhitePopup                                603   // Label: 'Black and White'	 NOTE: PPT
#define msotbidFavoritesSubmenu0000                           604   // Label: 'Favorites Submenu0000'	 NOTE: Web - there are potentially a lot of subdirs to Favorites, and each needs its own ID for merging
#define msotbidFavoritesSubmenu0001                           605   // Label: 'Favorites Submenu0001'
#define msotbidFavoritesSubmenu0002                           606   // Label: 'Favorites Submenu0002'
#define msotbidFavoritesSubmenu0003                           607   // Label: 'Favorites Submenu0003'
#define msotbidFavoritesSubmenu0004                           608   // Label: 'Favorites Submenu0004'
#define msotbidFavoritesSubmenu0005                           609   // Label: 'Favorites Submenu0005'
#define msotbidFavoritesSubmenu0006                           610   // Label: 'Favorites Submenu0006'
#define msotbidFavoritesSubmenu0007                           611   // Label: 'Favorites Submenu0007'
#define msotbidFavoritesSubmenu0008                           612   // Label: 'Favorites Submenu0008'
#define msotbidFavoritesSubmenu0009                           613   // Label: 'Favorites Submenu0009'
#define msotbidFavoritesSubmenu0010                           614   // Label: 'Favorites Submenu0010'
#define msotbidFavoritesSubmenu0011                           615   // Label: 'Favorites Submenu0011'
#define msotbidFavoritesSubmenu0012                           616   // Label: 'Favorites Submenu0012'
#define msotbidFavoritesSubmenu0013                           617   // Label: 'Favorites Submenu0013'
#define msotbidFavoritesSubmenu0014                           618   // Label: 'Favorites Submenu0014'
#define msotbidFavoritesSubmenu0015                           619   // Label: 'Favorites Submenu0015'
#define msotbidFavoritesSubmenu0016                           620   // Label: 'Favorites Submenu0016'
#define msotbidFavoritesSubmenu0017                           621   // Label: 'Favorites Submenu0017'
#define msotbidFavoritesSubmenu0018                           622   // Label: 'Favorites Submenu0018'
#define msotbidFavoritesSubmenu0019                           623   // Label: 'Favorites Submenu0019'
#define msotbidFavoritesSubmenu0020                           624   // Label: 'Favorites Submenu0020'
#define msotbidFavoritesSubmenu0021                           625   // Label: 'Favorites Submenu0021'
#define msotbidFavoritesSubmenu0022                           626   // Label: 'Favorites Submenu0022'
#define msotbidFavoritesSubmenu0023                           627   // Label: 'Favorites Submenu0023'
#define msotbidFavoritesSubmenu0024                           628   // Label: 'Favorites Submenu0024'
#define msotbidFavoritesSubmenu0025                           629   // Label: 'Favorites Submenu0025'
#define msotbidFavoritesSubmenu0026                           630   // Label: 'Favorites Submenu0026'
#define msotbidFavoritesSubmenu0027                           631   // Label: 'Favorites Submenu0027'
#define msotbidFavoritesSubmenu0028                           632   // Label: 'Favorites Submenu0028'
#define msotbidFavoritesSubmenu0029                           633   // Label: 'Favorites Submenu0029'
#define msotbidFavoritesSubmenu0030                           634   // Label: 'Favorites Submenu0030'
#define msotbidFavoritesSubmenu0031                           635   // Label: 'Favorites Submenu0031'
#define msotbidFavoritesSubmenu0032                           636   // Label: 'Favorites Submenu0032'
#define msotbidFavoritesSubmenu0033                           637   // Label: 'Favorites Submenu0033'
#define msotbidFavoritesSubmenu0034                           638   // Label: 'Favorites Submenu0034'
#define msotbidFavoritesSubmenu0035                           639   // Label: 'Favorites Submenu0035'
#define msotbidFavoritesSubmenu0036                           640   // Label: 'Favorites Submenu0036'
#define msotbidFavoritesSubmenu0037                           641   // Label: 'Favorites Submenu0037'
#define msotbidFavoritesSubmenu0038                           642   // Label: 'Favorites Submenu0038'
#define msotbidFavoritesSubmenu0039                           643   // Label: 'Favorites Submenu0039'
#define msotbidFavoritesSubmenu0040                           644   // Label: 'Favorites Submenu0040'
#define msotbidFavoritesSubmenu0041                           645   // Label: 'Favorites Submenu0041'
#define msotbidFavoritesSubmenu0042                           646   // Label: 'Favorites Submenu0042'
#define msotbidFavoritesSubmenu0043                           647   // Label: 'Favorites Submenu0043'
#define msotbidFavoritesSubmenu0044                           648   // Label: 'Favorites Submenu0044'
#define msotbidFavoritesSubmenu0045                           649   // Label: 'Favorites Submenu0045'
#define msotbidFavoritesSubmenu0046                           650   // Label: 'Favorites Submenu0046'
#define msotbidFavoritesSubmenu0047                           651   // Label: 'Favorites Submenu0047'
#define msotbidFavoritesSubmenu0048                           652   // Label: 'Favorites Submenu0048'
#define msotbidFavoritesSubmenu0049                           653   // Label: 'Favorites Submenu0049'
#define msotbidFavoritesSubmenu0050                           654   // Label: 'Favorites Submenu0050'
#define msotbidFavoritesSubmenu0051                           655   // Label: 'Favorites Submenu0051'
#define msotbidFavoritesSubmenu0052                           656   // Label: 'Favorites Submenu0052'
#define msotbidFavoritesSubmenu0053                           657   // Label: 'Favorites Submenu0053'
#define msotbidFavoritesSubmenu0054                           658   // Label: 'Favorites Submenu0054'
#define msotbidFavoritesSubmenu0055                           659   // Label: 'Favorites Submenu0055'
#define msotbidFavoritesSubmenu0056                           660   // Label: 'Favorites Submenu0056'
#define msotbidFavoritesSubmenu0057                           661   // Label: 'Favorites Submenu0057'
#define msotbidFavoritesSubmenu0058                           662   // Label: 'Favorites Submenu0058'
#define msotbidFavoritesSubmenu0059                           663   // Label: 'Favorites Submenu0059'
#define msotbidFavoritesSubmenu0060                           664   // Label: 'Favorites Submenu0060'
#define msotbidFavoritesSubmenu0061                           665   // Label: 'Favorites Submenu0061'
#define msotbidFavoritesSubmenu0062                           666   // Label: 'Favorites Submenu0062'
#define msotbidFavoritesSubmenu0063                           667   // Label: 'Favorites Submenu0063'
#define msotbidFavoritesSubmenu0064                           668   // Label: 'Favorites Submenu0064'
#define msotbidFavoritesSubmenu0065                           669   // Label: 'Favorites Submenu0065'
#define msotbidFavoritesSubmenu0066                           670   // Label: 'Favorites Submenu0066'
#define msotbidFavoritesSubmenu0067                           671   // Label: 'Favorites Submenu0067'
#define msotbidFavoritesSubmenu0068                           672   // Label: 'Favorites Submenu0068'
#define msotbidFavoritesSubmenu0069                           673   // Label: 'Favorites Submenu0069'
#define msotbidFavoritesSubmenu0070                           674   // Label: 'Favorites Submenu0070'
#define msotbidFavoritesSubmenu0071                           675   // Label: 'Favorites Submenu0071'
#define msotbidFavoritesSubmenu0072                           676   // Label: 'Favorites Submenu0072'
#define msotbidFavoritesSubmenu0073                           677   // Label: 'Favorites Submenu0073'
#define msotbidFavoritesSubmenu0074                           678   // Label: 'Favorites Submenu0074'
#define msotbidFavoritesSubmenu0075                           679   // Label: 'Favorites Submenu0075'
#define msotbidFavoritesSubmenu0076                           680   // Label: 'Favorites Submenu0076'
#define msotbidFavoritesSubmenu0077                           681   // Label: 'Favorites Submenu0077'
#define msotbidFavoritesSubmenu0078                           682   // Label: 'Favorites Submenu0078'
#define msotbidFavoritesSubmenu0079                           683   // Label: 'Favorites Submenu0079'
#define msotbidFavoritesSubmenu0080                           684   // Label: 'Favorites Submenu0080'
#define msotbidFavoritesSubmenu0081                           685   // Label: 'Favorites Submenu0081'
#define msotbidFavoritesSubmenu0082                           686   // Label: 'Favorites Submenu0082'
#define msotbidFavoritesSubmenu0083                           687   // Label: 'Favorites Submenu0083'
#define msotbidFavoritesSubmenu0084                           688   // Label: 'Favorites Submenu0084'
#define msotbidFavoritesSubmenu0085                           689   // Label: 'Favorites Submenu0085'
#define msotbidFavoritesSubmenu0086                           690   // Label: 'Favorites Submenu0086'
#define msotbidFavoritesSubmenu0087                           691   // Label: 'Favorites Submenu0087'
#define msotbidFavoritesSubmenu0088                           692   // Label: 'Favorites Submenu0088'
#define msotbidFavoritesSubmenu0089                           693   // Label: 'Favorites Submenu0089'
#define msotbidFavoritesSubmenu0090                           694   // Label: 'Favorites Submenu0090'
#define msotbidFavoritesSubmenu0091                           695   // Label: 'Favorites Submenu0091'
#define msotbidFavoritesSubmenu0092                           696   // Label: 'Favorites Submenu0092'
#define msotbidFavoritesSubmenu0093                           697   // Label: 'Favorites Submenu0093'
#define msotbidFavoritesSubmenu0094                           698   // Label: 'Favorites Submenu0094'
#define msotbidFavoritesSubmenu0095                           699   // Label: 'Favorites Submenu0095'
#define msotbidFavoritesSubmenu0096                           700   // Label: 'Favorites Submenu0096'
#define msotbidFavoritesSubmenu0097                           701   // Label: 'Favorites Submenu0097'
#define msotbidFavoritesSubmenu0098                           702   // Label: 'Favorites Submenu0098'
#define msotbidFavoritesSubmenu0099                           703   // Label: 'Favorites Submenu0099'
#define msotbidFavoritesSubmenu0100                           704   // Label: 'Favorites Submenu0100'
#define msotbidFavoritesSubmenu0101                           705   // Label: 'Favorites Submenu0101'
#define msotbidFavoritesSubmenu0102                           706   // Label: 'Favorites Submenu0102'
#define msotbidFavoritesSubmenu0103                           707   // Label: 'Favorites Submenu0103'
#define msotbidFavoritesSubmenu0104                           708   // Label: 'Favorites Submenu0104'
#define msotbidFavoritesSubmenu0105                           709   // Label: 'Favorites Submenu0105'
#define msotbidFavoritesSubmenu0106                           710   // Label: 'Favorites Submenu0106'
#define msotbidFavoritesSubmenu0107                           711   // Label: 'Favorites Submenu0107'
#define msotbidFavoritesSubmenu0108                           712   // Label: 'Favorites Submenu0108'
#define msotbidFavoritesSubmenu0109                           713   // Label: 'Favorites Submenu0109'
#define msotbidFavoritesSubmenu0110                           714   // Label: 'Favorites Submenu0110'
#define msotbidFavoritesSubmenu0111                           715   // Label: 'Favorites Submenu0111'
#define msotbidFavoritesSubmenu0112                           716   // Label: 'Favorites Submenu0112'
#define msotbidFavoritesSubmenu0113                           717   // Label: 'Favorites Submenu0113'
#define msotbidFavoritesSubmenu0114                           718   // Label: 'Favorites Submenu0114'
#define msotbidFavoritesSubmenu0115                           719   // Label: 'Favorites Submenu0115'
#define msotbidFavoritesSubmenu0116                           720   // Label: 'Favorites Submenu0116'
#define msotbidFavoritesSubmenu0117                           721   // Label: 'Favorites Submenu0117'
#define msotbidFavoritesSubmenu0118                           722   // Label: 'Favorites Submenu0118'
#define msotbidFavoritesSubmenu0119                           723   // Label: 'Favorites Submenu0119'
#define msotbidFavoritesSubmenu0120                           724   // Label: 'Favorites Submenu0120'
#define msotbidFavoritesSubmenu0121                           725   // Label: 'Favorites Submenu0121'
#define msotbidFavoritesSubmenu0122                           726   // Label: 'Favorites Submenu0122'
#define msotbidFavoritesSubmenu0123                           727   // Label: 'Favorites Submenu0123'
#define msotbidFavoritesSubmenu0124                           728   // Label: 'Favorites Submenu0124'
#define msotbidFavoritesSubmenu0125                           729   // Label: 'Favorites Submenu0125'
#define msotbidFavoritesSubmenu0126                           730   // Label: 'Favorites Submenu0126'
#define msotbidFavoritesSubmenu0127                           731   // Label: 'Favorites Submenu0127'
#define msotbidFavoritesSubmenu0128                           732   // Label: 'Favorites Submenu0128'
#define msotbidFavoritesSubmenu0129                           733   // Label: 'Favorites Submenu0129'
#define msotbidFavoritesSubmenu0130                           734   // Label: 'Favorites Submenu0130'
#define msotbidFavoritesSubmenu0131                           735   // Label: 'Favorites Submenu0131'
#define msotbidFavoritesSubmenu0132                           736   // Label: 'Favorites Submenu0132'
#define msotbidFavoritesSubmenu0133                           737   // Label: 'Favorites Submenu0133'
#define msotbidFavoritesSubmenu0134                           738   // Label: 'Favorites Submenu0134'
#define msotbidFavoritesSubmenu0135                           739   // Label: 'Favorites Submenu0135'
#define msotbidFavoritesSubmenu0136                           740   // Label: 'Favorites Submenu0136'
#define msotbidFavoritesSubmenu0137                           741   // Label: 'Favorites Submenu0137'
#define msotbidFavoritesSubmenu0138                           742   // Label: 'Favorites Submenu0138'
#define msotbidFavoritesSubmenu0139                           743   // Label: 'Favorites Submenu0139'
#define msotbidFavoritesSubmenu0140                           744   // Label: 'Favorites Submenu0140'
#define msotbidFavoritesSubmenu0141                           745   // Label: 'Favorites Submenu0141'
#define msotbidFavoritesSubmenu0142                           746   // Label: 'Favorites Submenu0142'
#define msotbidFavoritesSubmenu0143                           747   // Label: 'Favorites Submenu0143'
#define msotbidFavoritesSubmenu0144                           748   // Label: 'Favorites Submenu0144'
#define msotbidFavoritesSubmenu0145                           749   // Label: 'Favorites Submenu0145'
#define msotbidFavoritesSubmenu0146                           750   // Label: 'Favorites Submenu0146'
#define msotbidFavoritesSubmenu0147                           751   // Label: 'Favorites Submenu0147'
#define msotbidFavoritesSubmenu0148                           752   // Label: 'Favorites Submenu0148'
#define msotbidFavoritesSubmenu0149                           753   // Label: 'Favorites Submenu0149'
#define msotbidFavoritesSubmenu0150                           754   // Label: 'Favorites Submenu0150'
#define msotbidFavoritesSubmenu0151                           755   // Label: 'Favorites Submenu0151'
#define msotbidFavoritesSubmenu0152                           756   // Label: 'Favorites Submenu0152'
#define msotbidFavoritesSubmenu0153                           757   // Label: 'Favorites Submenu0153'
#define msotbidFavoritesSubmenu0154                           758   // Label: 'Favorites Submenu0154'
#define msotbidFavoritesSubmenu0155                           759   // Label: 'Favorites Submenu0155'
#define msotbidFavoritesSubmenu0156                           760   // Label: 'Favorites Submenu0156'
#define msotbidFavoritesSubmenu0157                           761   // Label: 'Favorites Submenu0157'
#define msotbidFavoritesSubmenu0158                           762   // Label: 'Favorites Submenu0158'
#define msotbidFavoritesSubmenu0159                           763   // Label: 'Favorites Submenu0159'
#define msotbidFavoritesSubmenu0160                           764   // Label: 'Favorites Submenu0160'
#define msotbidFavoritesSubmenu0161                           765   // Label: 'Favorites Submenu0161'
#define msotbidFavoritesSubmenu0162                           766   // Label: 'Favorites Submenu0162'
#define msotbidFavoritesSubmenu0163                           767   // Label: 'Favorites Submenu0163'
#define msotbidFavoritesSubmenu0164                           768   // Label: 'Favorites Submenu0164'
#define msotbidFavoritesSubmenu0165                           769   // Label: 'Favorites Submenu0165'
#define msotbidFavoritesSubmenu0166                           770   // Label: 'Favorites Submenu0166'
#define msotbidFavoritesSubmenu0167                           771   // Label: 'Favorites Submenu0167'
#define msotbidFavoritesSubmenu0168                           772   // Label: 'Favorites Submenu0168'
#define msotbidFavoritesSubmenu0169                           773   // Label: 'Favorites Submenu0169'
#define msotbidFavoritesSubmenu0170                           774   // Label: 'Favorites Submenu0170'
#define msotbidFavoritesSubmenu0171                           775   // Label: 'Favorites Submenu0171'
#define msotbidFavoritesSubmenu0172                           776   // Label: 'Favorites Submenu0172'
#define msotbidFavoritesSubmenu0173                           777   // Label: 'Favorites Submenu0173'
#define msotbidFavoritesSubmenu0174                           778   // Label: 'Favorites Submenu0174'
#define msotbidFavoritesSubmenu0175                           779   // Label: 'Favorites Submenu0175'
#define msotbidFavoritesSubmenu0176                           780   // Label: 'Favorites Submenu0176'
#define msotbidFavoritesSubmenu0177                           781   // Label: 'Favorites Submenu0177'
#define msotbidFavoritesSubmenu0178                           782   // Label: 'Favorites Submenu0178'
#define msotbidFavoritesSubmenu0179                           783   // Label: 'Favorites Submenu0179'
#define msotbidFavoritesSubmenu0180                           784   // Label: 'Favorites Submenu0180'
#define msotbidFavoritesSubmenu0181                           785   // Label: 'Favorites Submenu0181'
#define msotbidFavoritesSubmenu0182                           786   // Label: 'Favorites Submenu0182'
#define msotbidFavoritesSubmenu0183                           787   // Label: 'Favorites Submenu0183'
#define msotbidFavoritesSubmenu0184                           788   // Label: 'Favorites Submenu0184'
#define msotbidFavoritesSubmenu0185                           789   // Label: 'Favorites Submenu0185'
#define msotbidFavoritesSubmenu0186                           790   // Label: 'Favorites Submenu0186'
#define msotbidFavoritesSubmenu0187                           791   // Label: 'Favorites Submenu0187'
#define msotbidFavoritesSubmenu0188                           792   // Label: 'Favorites Submenu0188'
#define msotbidFavoritesSubmenu0189                           793   // Label: 'Favorites Submenu0189'
#define msotbidFavoritesSubmenu0190                           794   // Label: 'Favorites Submenu0190'
#define msotbidFavoritesSubmenu0191                           795   // Label: 'Favorites Submenu0191'
#define msotbidFavoritesSubmenu0192                           796   // Label: 'Favorites Submenu0192'
#define msotbidFavoritesSubmenu0193                           797   // Label: 'Favorites Submenu0193'
#define msotbidFavoritesSubmenu0194                           798   // Label: 'Favorites Submenu0194'
#define msotbidFavoritesSubmenu0195                           799   // Label: 'Favorites Submenu0195'
#define msotbidFavoritesSubmenu0196                           800   // Label: 'Favorites Submenu0196'
#define msotbidFavoritesSubmenu0197                           801   // Label: 'Favorites Submenu0197'
#define msotbidFavoritesSubmenu0198                           802   // Label: 'Favorites Submenu0198'
#define msotbidFavoritesSubmenu0199                           803   // Label: 'Favorites Submenu0199'
#define msotbidFavoritesSubmenu0200                           804   // Label: 'Favorites Submenu0200'
#define msotbidFavoritesSubmenu0201                           805   // Label: 'Favorites Submenu0201'
#define msotbidFavoritesSubmenu0202                           806   // Label: 'Favorites Submenu0202'
#define msotbidFavoritesSubmenu0203                           807   // Label: 'Favorites Submenu0203'
#define msotbidFavoritesSubmenu0204                           808   // Label: 'Favorites Submenu0204'
#define msotbidFavoritesSubmenu0205                           809   // Label: 'Favorites Submenu0205'
#define msotbidFavoritesSubmenu0206                           810   // Label: 'Favorites Submenu0206'
#define msotbidFavoritesSubmenu0207                           811   // Label: 'Favorites Submenu0207'
#define msotbidFavoritesSubmenu0208                           812   // Label: 'Favorites Submenu0208'
#define msotbidFavoritesSubmenu0209                           813   // Label: 'Favorites Submenu0209'
#define msotbidFavoritesSubmenu0210                           814   // Label: 'Favorites Submenu0210'
#define msotbidFavoritesSubmenu0211                           815   // Label: 'Favorites Submenu0211'
#define msotbidFavoritesSubmenu0212                           816   // Label: 'Favorites Submenu0212'
#define msotbidFavoritesSubmenu0213                           817   // Label: 'Favorites Submenu0213'
#define msotbidFavoritesSubmenu0214                           818   // Label: 'Favorites Submenu0214'
#define msotbidFavoritesSubmenu0215                           819   // Label: 'Favorites Submenu0215'
#define msotbidFavoritesSubmenu0216                           820   // Label: 'Favorites Submenu0216'
#define msotbidFavoritesSubmenu0217                           821   // Label: 'Favorites Submenu0217'
#define msotbidFavoritesSubmenu0218                           822   // Label: 'Favorites Submenu0218'
#define msotbidFavoritesSubmenu0219                           823   // Label: 'Favorites Submenu0219'
#define msotbidFavoritesSubmenu0220                           824   // Label: 'Favorites Submenu0220'
#define msotbidFavoritesSubmenu0221                           825   // Label: 'Favorites Submenu0221'
#define msotbidFavoritesSubmenu0222                           826   // Label: 'Favorites Submenu0222'
#define msotbidFavoritesSubmenu0223                           827   // Label: 'Favorites Submenu0223'
#define msotbidFavoritesSubmenu0224                           828   // Label: 'Favorites Submenu0224'
#define msotbidFavoritesSubmenu0225                           829   // Label: 'Favorites Submenu0225'
#define msotbidFavoritesSubmenu0226                           830   // Label: 'Favorites Submenu0226'
#define msotbidFavoritesSubmenu0227                           831   // Label: 'Favorites Submenu0227'
#define msotbidFavoritesSubmenu0228                           832   // Label: 'Favorites Submenu0228'
#define msotbidFavoritesSubmenu0229                           833   // Label: 'Favorites Submenu0229'
#define msotbidFavoritesSubmenu0230                           834   // Label: 'Favorites Submenu0230'
#define msotbidFavoritesSubmenu0231                           835   // Label: 'Favorites Submenu0231'
#define msotbidFavoritesSubmenu0232                           836   // Label: 'Favorites Submenu0232'
#define msotbidFavoritesSubmenu0233                           837   // Label: 'Favorites Submenu0233'
#define msotbidFavoritesSubmenu0234                           838   // Label: 'Favorites Submenu0234'
#define msotbidFavoritesSubmenu0235                           839   // Label: 'Favorites Submenu0235'
#define msotbidFavoritesSubmenu0236                           840   // Label: 'Favorites Submenu0236'
#define msotbidFavoritesSubmenu0237                           841   // Label: 'Favorites Submenu0237'
#define msotbidFavoritesSubmenu0238                           842   // Label: 'Favorites Submenu0238'
#define msotbidFavoritesSubmenu0239                           843   // Label: 'Favorites Submenu0239'
#define msotbidFavoritesSubmenu0240                           844   // Label: 'Favorites Submenu0240'
#define msotbidFavoritesSubmenu0241                           845   // Label: 'Favorites Submenu0241'
#define msotbidFavoritesSubmenu0242                           846   // Label: 'Favorites Submenu0242'
#define msotbidFavoritesSubmenu0243                           847   // Label: 'Favorites Submenu0243'
#define msotbidFavoritesSubmenu0244                           848   // Label: 'Favorites Submenu0244'
#define msotbidFavoritesSubmenu0245                           849   // Label: 'Favorites Submenu0245'
#define msotbidFavoritesSubmenu0246                           850   // Label: 'Favorites Submenu0246'
#define msotbidFavoritesSubmenu0247                           851   // Label: 'Favorites Submenu0247'
#define msotbidFavoritesSubmenu0248                           852   // Label: 'Favorites Submenu0248'
#define msotbidFavoritesSubmenu0249                           853   // Label: 'Favorites Submenu0249'
#define msotbidFavoritesSubmenu0250                           854   // Label: 'Favorites Submenu0250'
#define msotbidFavoritesSubmenu0251                           855   // Label: 'Favorites Submenu0251'
#define msotbidFavoritesSubmenu0252                           856   // Label: 'Favorites Submenu0252'
#define msotbidFavoritesSubmenu0253                           857   // Label: 'Favorites Submenu0253'
#define msotbidFavoritesSubmenu0254                           858   // Label: 'Favorites Submenu0254'
#define msotbidFavoritesSubmenu0255                           859   // Label: 'Favorites Submenu0255'
#define msotbidFavoritesSubmenu0256                           860   // Label: 'Favorites Submenu0256'
#define msotbidFavoritesSubmenu0257                           861   // Label: 'Favorites Submenu0257'
#define msotbidFavoritesSubmenu0258                           862   // Label: 'Favorites Submenu0258'
#define msotbidFavoritesSubmenu0259                           863   // Label: 'Favorites Submenu0259'
#define msotbidFavoritesSubmenu0260                           864   // Label: 'Favorites Submenu0260'
#define msotbidFavoritesSubmenu0261                           865   // Label: 'Favorites Submenu0261'
#define msotbidFavoritesSubmenu0262                           866   // Label: 'Favorites Submenu0262'
#define msotbidFavoritesSubmenu0263                           867   // Label: 'Favorites Submenu0263'
#define msotbidFavoritesSubmenu0264                           868   // Label: 'Favorites Submenu0264'
#define msotbidFavoritesSubmenu0265                           869   // Label: 'Favorites Submenu0265'
#define msotbidFavoritesSubmenu0266                           870   // Label: 'Favorites Submenu0266'
#define msotbidFavoritesSubmenu0267                           871   // Label: 'Favorites Submenu0267'
#define msotbidFavoritesSubmenu0268                           872   // Label: 'Favorites Submenu0268'
#define msotbidFavoritesSubmenu0269                           873   // Label: 'Favorites Submenu0269'
#define msotbidFavoritesSubmenu0270                           874   // Label: 'Favorites Submenu0270'
#define msotbidFavoritesSubmenu0271                           875   // Label: 'Favorites Submenu0271'
#define msotbidFavoritesSubmenu0272                           876   // Label: 'Favorites Submenu0272'
#define msotbidFavoritesSubmenu0273                           877   // Label: 'Favorites Submenu0273'
#define msotbidFavoritesSubmenu0274                           878   // Label: 'Favorites Submenu0274'
#define msotbidFavoritesSubmenu0275                           879   // Label: 'Favorites Submenu0275'
#define msotbidFavoritesSubmenu0276                           880   // Label: 'Favorites Submenu0276'
#define msotbidFavoritesSubmenu0277                           881   // Label: 'Favorites Submenu0277'
#define msotbidFavoritesSubmenu0278                           882   // Label: 'Favorites Submenu0278'
#define msotbidFavoritesSubmenu0279                           883   // Label: 'Favorites Submenu0279'
#define msotbidFavoritesSubmenu0280                           884   // Label: 'Favorites Submenu0280'
#define msotbidFavoritesSubmenu0281                           885   // Label: 'Favorites Submenu0281'
#define msotbidFavoritesSubmenu0282                           886   // Label: 'Favorites Submenu0282'
#define msotbidFavoritesSubmenu0283                           887   // Label: 'Favorites Submenu0283'
#define msotbidFavoritesSubmenu0284                           888   // Label: 'Favorites Submenu0284'
#define msotbidFavoritesSubmenu0285                           889   // Label: 'Favorites Submenu0285'
#define msotbidFavoritesSubmenu0286                           890   // Label: 'Favorites Submenu0286'
#define msotbidFavoritesSubmenu0287                           891   // Label: 'Favorites Submenu0287'
#define msotbidFavoritesSubmenu0288                           892   // Label: 'Favorites Submenu0288'
#define msotbidFavoritesSubmenu0289                           893   // Label: 'Favorites Submenu0289'
#define msotbidFavoritesSubmenu0290                           894   // Label: 'Favorites Submenu0290'
#define msotbidFavoritesSubmenu0291                           895   // Label: 'Favorites Submenu0291'
#define msotbidFavoritesSubmenu0292                           896   // Label: 'Favorites Submenu0292'
#define msotbidFavoritesSubmenu0293                           897   // Label: 'Favorites Submenu0293'
#define msotbidFavoritesSubmenu0294                           898   // Label: 'Favorites Submenu0294'
#define msotbidFavoritesSubmenu0295                           899   // Label: 'Favorites Submenu0295'
#define msotbidFavoritesSubmenu0296                           900   // Label: 'Favorites Submenu0296'
#define msotbidFavoritesSubmenu0297                           901   // Label: 'Favorites Submenu0297'
#define msotbidFavoritesSubmenu0298                           902   // Label: 'Favorites Submenu0298'
#define msotbidFavoritesSubmenu0299                           903   // Label: 'Favorites Submenu0299'
#define msotbidFavoritesSubmenu0300                           904   // Label: 'Favorites Submenu0300'
#define msotbidMacroMenu                                      905   // Label: 'Macro Menu'	 NOTE: Access
#define msotbidBorderWidthMenu                                906   // Label: 'Line/Border Width'	 NOTE: Access
#define msotbidFormDesignBorderColorMenu                      907   // Label: 'Line/Border Color'	 NOTE: Access
#define msotbidSaveFormToMenu                                 908   // Label: 'Save Form To'	 NOTE: Ren
#define msotbidViewDBCIconsMenu                               909   // Label: 'View Icons'	 NOTE: Access - large icons, small icons, list, etc.
#define msotbidFontColorMenu                                  910   // Label: 'Font Color'	 NOTE: Escher/PPT
#define msotbidSlideShowCustomShowMenu                        911   // Label: 'Custom Show'	 NOTE: PPT
#define msotbidOutlineWell                                    912   // Label: 'Outline Well'	 NOTE: PPT
#define msotbidDummy913                                       913   // Label: ' '	 NOTE: Formerly AnimationWell
#define msotbidSlideShowWell                                  914   // Label: 'Slide Show Well'	 NOTE: PPT
#define msotbidPresetAnimationMenu                            915   // Label: 'Preset Animation'	 NOTE: PPT
#define msotbidSlideShowGoByTitleMenu                         916   // Label: 'By Title'	 NOTE: PPT
#define msotbidDlgFontColorMenu                               917   // Label: 'Dlg Font Color'	 NOTE: PPT
#define msotbidDlgFontColorNoAutomaticMenu                    918   // Label: 'Dlg Font Color No Automatic'	 NOTE: PPT
#define msotbidDlgFillEffectsColorMenu                        919   // Label: 'Dlg Fill Effects Color'	 NOTE: PPT
#define msotbidDlgFillColorMenu                               920   // Label: 'Dlg Fill Color'	 NOTE: PPT
#define msotbidDlgBuildSlideAfterMenu                         921   // Label: 'Dlg Build Slide After'	 NOTE: PPT
#define msotbidEmptyPopup                                     922   // Label: 'Empty'	 NOTE: PPT
#define msotbidNotesViewSlidePopup                            923   // Label: 'Notes View Slide'	 NOTE: PPT
#define msotbidAlignmentFEMenu                                924   // Label: 'Alignment (Far East)'	 NOTE: PPT - FE version of text alignment menu
#define msotbidAnalysisSplitMenu                              925   // Label: 'Analysis Menu 2'	 NOTE: Access
#define msotbidShortcutMenuCategory_Database                  926   // Label: 'Database'	 NOTE: Access - categories in the shortcut menu bar
#define msotbidShortcutMenuCategory_Filter                    927   // Label: 'Filter'	 NOTE: Access
#define msotbidShortcutMenuCategory_Form                      928   // Label: 'Form'	 NOTE: Access
#define msotbidShortcutMenuCategory_Index                     929   // Label: 'Index'	 NOTE: Access
#define msotbidShortcutMenuCategory_Macro                     930   // Label: 'Macro'	 NOTE: Access
#define msotbidShortcutMenuCategory_Module                    931   // Label: 'Module'	 NOTE: Access
#define msotbidShortcutMenuCategory_Query                     932   // Label: 'Query'	 NOTE: Access
#define msotbidShortcutMenuCategory_Relationship              933   // Label: 'Relationship'	 NOTE: Access
#define msotbidShortcutMenuCategory_Report                    934   // Label: 'Report'	 NOTE: Access
#define msotbidShortcutMenuCategory_Table                     935   // Label: 'Table'	 NOTE: Access
#define msotbidShortcutMenuCategory_Custom                    936   // Label: 'Custom'	 NOTE: Access
#define msotbidCircularReference                              937   // Label: 'Circular Reference'	 NOTE: Excel
#define msotbidGoToSlideByTitleMenu                           938   // Label: 'Go To Slide By Title'	 NOTE: PPT
#define msotbidGoToSlideByCustomShowMenu                      939   // Label: 'Go To Slide By Custom Show'	 NOTE: PPT
#define msotbidSlideShowHyperlinkPopup                        940   // Label: 'Hyperlinked Object'	 NOTE: PPT - ctxt menu for object with hyperlink
#define msotbidSlideShowBrowsePopup                           941   // Label: 'Slide Show Browse'	 NOTE: PPT - slide show in browse mode
#define msotbidSaveToHTMLMenu                                 942   // Label: 'Save to HTML'	 NOTE: File submenu
#define msotbidInactiveChartPopup                             943   // Label: 'Inactive Chart'	 NOTE: Excel
#define msotbidOutlookHelpMenu                                944   // Label: 'Outlook Help'	 NOTE: Outlook - cf. BinderHelpMenu
#define msotbidVBEToggle                                      945   // Label: 'Toggle'	 NOTE: VBE, Access
#define msotbidShortcutMenuCategory_Draw                      946   // Label: 'Draw Shortcut Menu Category'	 NOTE: PPT
#define msotbidShortcutMenuCategory_Text                      947   // Label: 'Text Shortcut Menu Category'	 NOTE: PPT
#define msotbidShortcutMenuCategory_SlideShow                 948   // Label: 'Slide Show Shortcut Menu Category'	 NOTE: PPT
#define msotbidMicrosoftOnTheWeb_ContainerMenu                949   // Label: 'Microsoft on the Web'	 NOTE: used by Docobj containers like Binder
#define msotbidAlignment                                      950   // Label: 'Alignment'	 NOTE: Word
#define msotbidHighlightMenu                                  951   // Label: 'Highlight'	 NOTE: Word
#define msotbidShadingColor                                   952   // Label: 'Shading Color'	 NOTE: Word
#define msotbidNavigationPanePopup                            953   // Label: 'Document Map'	 NOTE: Word
#define msotbidFieldDisplayListNumPopup                       954   // Label: 'Field Display List Numbers'	 NOTE: Word
#define msotbidFormatBackgroundMenu                           955   // Label: 'Format Background'	 NOTE: Word
#define msotbidChartWell                                      956   // Label: 'Chart Well'	 NOTE: Graph - cf. msotbidChartingWell
#define msotbidBackgroundProofingStatusBarPopup               957   // Label: 'Background Proofing Status Bar'	 NOTE: Word
#define msotbidAutoTextCategory                               958   // Label: 'AutoText Category'	 NOTE: Word
#define msotbidHelpWell                                       959   // Label: 'Help Well'	 NOTE: Graph - cf. msotbidWindowAndHelpWell
#define msotbidTableCellVerticalAlignmentMenu                 960   // Label: 'Alignment'	 NOTE: Word
#define msotbidOCXVerbMenu                                    961   // Label: 'OCX Verbs'	 NOTE: PPT - different from EditOLEObjectMenu
#define msotbidMovieToolbar                                   962   // Label: 'Movie'	 NOTE: Escher - Movie Toolbar
#define msotbidWizardMenu                                     963   // Label: 'Wizard'	 NOTE: Tools Wizard menu
#define msotbidRibbon                                         964   // Label: 'Ribbon'
#define msotbidRuler                                          965   // Label: 'Ruler'
#define msotbidFormatBackground                               966   // Label: 'Background'	 NOTE: Word
#define msotbidThesTablePopup                                 967   // Label: 'Table w/ Thesaurus'	 NOTE: Word
#define msotbidThesPopup                                      968   // Label: 'Text w/ Thesaurus'	 NOTE: Word
#define msotbidSynonymsPopup                                  969   // Label: 'Synonyms'	 NOTE: Word
#define msotbidWorkWell                                       970   // Label: 'Work'	 NOTE: Word
#define msotbidGlossaryMenu                                   971   // Label: 'Glossary'
#define msotbidFile51Menu                                     972   // Label: 'File51'	 NOTE: Mac TBID
#define msotbidEdit51Menu                                     973   // Label: 'Edit51'	 NOTE: Mac TBID
#define msotbidView51Menu                                     974   // Label: 'View51'	 NOTE: Mac TBID
#define msotbidInsert51Menu                                   975   // Label: 'Insert51'	 NOTE: Mac TBID
#define msotbidFormat51Menu                                   976   // Label: 'Format51'	 NOTE: Mac TBID
#define msotbidTools51Menu                                    977   // Label: 'Tools51'	 NOTE: Mac TBID
#define msotbidDummy978                                       978   // Label: ' '
#define msotbidToolboxJavaApplets                             979   // Label: 'Java Applets'	 NOTE: Access control toolbox
#define msotbidToolboxHTMLControls                            980   // Label: 'HTML Controls'	 NOTE: Access control toolbox
#define msotbidToolboxDatapageMain                            981   // Label: 'Web Tools'	 NOTE: Access control toolbox
#define msotbidOfficeAssistantMenu                            982   // Label: 'Office Assistant'
#define msotbidCollectCopy                                    983   // Label: 'Clipboard'
#define msotbidScriptAnchorPopup                              984   // Label: 'Script Anchor Popup'
#define msotbidHorizRulePopup                                 985   // Label: 'Horizontal Line Popup'
#define msotbidConference                                     986   // Label: 'Online Meeting'	 NOTE: Net Meeting toolbar by PPT in all apps
#define msotbidAutoSigList                                    987   // Label: '&Open Signature'	 NOTE: Word AutoSignature menu
#define msotbidFieldAutoSigListPopup                          988   // Label: 'AutoSignature Popup'	 NOTE: Word AutoSignature popup
#define msotbidAutoSigBold                                    989   // Label: 'Bold'	 NOTE: Word AutoSig dialog bold button toolbar uses one tb per button
#define msotbidAutoSigItalic                                  990   // Label: 'Italic'
#define msotbidAutoSigUnderline                               991   // Label: 'Underline'
#define msotbidAutoSigLeftPara                                992   // Label: 'Left Justify Paragraph'
#define msotbidAutoSigCenterPara                              993   // Label: 'Center Justify Paragraph'
#define msotbidAutoSigRightPara                               994   // Label: 'Right Justify Paragraph'
#define msotbidAutoSigCharColor                               995   // Label: 'Character Color'
#define msotbidAutoSigFont                                    996   // Label: 'Font'
#define msotbidAutoSigFontSize                                997   // Label: 'Font Size'
#define msotbidAutoSigFontColor                               998   // Label: 'Font Color'
#define msotbidAutoSigPicture                                 999   // Label: 'Insert Picture'
#define msotbidAutoSigHyperlink                               1000   // Label: 'Insert Hyperlink'
#define msotbidFixSpellingLangMenu                            1001   // Label: 'Language'	 NOTE: from Word97 FE
#define msotbidBiLanguageMenu                                 1002   // Label: 'BiDi Language Menu'	 NOTE: Word97 BiDi
#define msotbidDummy1003                                      1003   // Label: ' '
#define msotbidFontColorNoTearoff                             1004   // Label: 'Font Color Dropdown'	 NOTE: Word AutoSig dialog
#define msotbidAssignHyperlinkBar                             1005   // Label: 'Assign Hyperlink Menu'
#define msotbidFramesetMenu                                   1006   // Label: 'Frames Menu'
#define msotbidStyleCheckerPopup                              1007   // Label: 'Style Checker Popup'
#define msotbidCondFormatDlg                                  1008   // Label: 'Conditional Formatting Dialog'
#define msotbidCondFormatDlgFC                                1009   // Label: 'Conditional Formatting Dialog ForeColor'
#define msotbidCondFormatDlgBC                                1010   // Label: 'Conditional Formatting Dialog BackColor'
#define msotbidDaVinciTableDesignPopup                        1011   // Label: 'View Table Design Popup'
#define msotbidDaVinciQueryDesignBackgroundPopup              1012   // Label: 'View Design Background Popup'
#define msotbidDaVinciQueryDesignDiagramPanePopup             1013   // Label: 'View Design Diagram Pane Popup'
#define msotbidDaVinciQueryDesignFieldPopup                   1014   // Label: 'View Design Field Popup'
#define msotbidDaVinciQueryDesignGridPanePopup                1015   // Label: 'View Design Grid Pane Popup'
#define msotbidDaVinciQueryDesignSQLPanePopup                 1016   // Label: 'View Design SQL Pane Popup'
#define msotbidDaVinciSchemaDesignDesignPopup                 1017   // Label: 'View Schema Design Design Popup'
#define msotbidDaVinciSchemaDesignDiagramPopup                1018   // Label: 'View Schema Design Diagram Popup'
#define msotbidDaVinciQueryDesignMultiSelectPopup             1019   // Label: 'View Design Multiple Select Popup'
#define msotbidDaVinciQueryDesignJoinLinePopup                1020   // Label: 'View Design Join Line Popup'
#define msotbidDaVinciTableViewModeMenu                       1021   // Label: 'View Table View Mode Submenu'
#define msotbidDaVinciShowPanesMenu                           1022   // Label: 'View Show Panes Submenu'
#define msotbidDatapagePalette                                1023   // Label: 'Formatting (Page)'
#define msotbidDatapagePopup                                  1024   // Label: 'Page Popup'
#define msotbidDatapage                                       1025   // Label: 'Page Design'
#define msotbidDBCToolbar                                     1026   // Label: 'DBC Toolbar'
#define msotbidDBCDatapagePopup                               1027   // Label: 'Database Page Popup'
#define msotbidFunctionKeys                                   1028   // Label: 'Function Key Display'
#define msotbidCommonTasksMenu                                1029   // Label: 'Common Tasks'
#define msotbidMailHeader                                     1030   // Label: 'Envelope'	 NOTE: toolbar of the envelope
#define msotbidCommandBarOptionsMenu                          1031   // Label: 'Toolbar Options Menu'
#define msotbidQuickCustomizeMenu                             1032   // Label: 'Quick Customize'
#define msotbidCollectCopyMenu                                1033   // Label: 'Clipboard Paste Menu'	 NOTE: dropdown on docked active clipboard
#define msotbidDummy1034                                      1034   // Label: ' '
#define msotbidRehearsal                                      1035   // Label: 'Rehearsal'	 NOTE: PP rehearsal toolbar
#define msotbidSubdatasheetMenu                               1036   // Label: 'Subdatasheet'
#define msotbidViewDirMenu                                    1037   // Label: 'View Direction'	 NOTE: BiDi
#define msotbidDiagramMenu                                    1038   // Label: 'Diagram'
#define msotbidDiagramPopup                                   1039   // Label: 'Diagram'
#define msotbidReportDesignTabControlPopup                    1040   // Label: 'Tab Control on Report Design'
#define msotbidGrammarVersion2Popup                           1041   // Label: 'Grammar (2)'	 NOTE: like GrammarPopup, but for v2 grammar checkers
#define msotbidTableAutoFitShortMenu                          1042   // Label: 'AutoFit'
#define msotbidTableCellAlignment                             1043   // Label: 'Cell Alignment'
#define msotbidTableInsertLongMenu                            1044   // Label: 'Insert Table'
#define msotbidAccessConvertMenu                              1045   // Label: 'Convert Database'
#define msotbidPivotChartPopup                                1046   // Label: 'Pivot Chart Popup'	 NOTE: shown when the user right clicks on a pivot table field button on a pivot chart
#define msotbidAccessDatapageLayoutPopup                      1047   // Label: 'Web Page Layout'
#define msotbidUnderlineColor                                 1048   // Label: 'Underline Color'
#define msotbidUnderlineStyle                                 1049   // Label: 'Underline Style'
#define msotbidUnderlineColorMenu                             1050   // Label: 'Underline Color'
#define msotbidFormatAsianLayoutMenu                          1051   // Label: 'Asian Layout'	 NOTE: FE
#define msotbidBackgroundMenu                                 1052   // Label: 'Background'	 NOTE: cascade under Format in Web Page design
#define msotbidPPTMultipleMonitorsSlideShow                   1053   // Label: 'Slide Show'	 NOTE: PPT multiple monitor support during slide show
#define msotbidTableAutoFitLongMenu                           1054   // Label: 'AutoFit'
#define msotbidDaVinciDiagramDesignDesignPopup                1055   // Label: 'Diagram Design Popup'
#define msotbidDaVinciDiagramDesignDiagramPopup               1056   // Label: 'Diagram Popup'
#define msotbidDaVinciDiagramDesignJoinLinePopup              1057   // Label: 'Join Line Popup'
#define msotbidAllDiagramsWell                                1058   // Label: 'All Database Diagrams Well'
#define msotbidDatapageBrowse                                 1059   // Label: 'Page View'
#define msotbidPivotChartMenu                                 1060   // Label: 'PivotChart Menu'	 NOTE: the menu on the left of the pivot table toolbar when a chart is active
#define msotbidCollaborationMenu                              1061   // Label: 'Online Collaboration'	 NOTE: a cascade under Tools menu
#define msotbidViewDesign                                     1062   // Label: 'View Design'	 NOTE: New toolbar for designing DaVinci views.
#define msotbidDiagramDesign                                  1063   // Label: 'Diagram Design'	 NOTE: New toolbar for designing DaVinci diagrams.
#define msotbidTableModesMenu                                 1064   // Label: 'Table'	 NOTE: New dropdown submenu for table modes in Davinci view and schema design
#define msotbidZoomModesMenu                                  1065   // Label: 'Zoom'	 NOTE: New dropdown submenu for zoom modes in Davinci schema design
#define msotbidSPDesign                                       1066   // Label: 'Stored Procedure Design'	 NOTE: New toolbar for designing stored procedures.
#define msotbidAllViewsWell                                   1067   // Label: 'All Views Well'	 NOTE: Access: Customize dialog well
#define msotbidAllStoredProceduresWell                        1068   // Label: 'All Stored Procedures Well'	 NOTE: Access: Customize dialog well
#define msotbidAllDataPagesWell                               1069   // Label: 'All Pages'	 NOTE: Access: Customize dialog well
#define msotbidNotesPanePopup                                 1070   // Label: 'Notes Pane'	 NOTE: PowerPoint: Popup for notes pane of Normal View
#define msotbidDPSpecialEffectMenu                            1071   // Label: 'Appearance'	 NOTE: Border styles for datapages
#define msotbidPivotChartSeriesPopup                          1072   // Label: 'Format Data Series'	 NOTE: Excel: pivot chart drilling on data series
#define msotbidPivotChartCatAxisPopup                         1073   // Label: 'Format Axis'	 NOTE: Excel: pivot chart drilling on category axis labels
#define msotbidPivotChartLdEntryPopup                         1074   // Label: 'Format Legend Entry'	 NOTE: Excel: pivot chart drilling on legend entry
#define msotbidStoredProcedureDesignDatasheetPopup            1075   // Label: 'Stored Procedure Design Datasheet'	 NOTE: Access: stored procedure design datasheet popup
#define msotbidViewDesignDatasheetPopup                       1076   // Label: 'View Design Datasheet'	 NOTE: Access: view design datasheet popup
#define msotbidDiagramDesignLabelPopup                        1077   // Label: 'Diagram Design Label Popup'	 NOTE: Access: diagram design right click menu popup for labels
#define msotbidStoredProcedureDesignPopup                     1078   // Label: 'Stored Procedure Design Popup'	 NOTE: Access: stored procedure design popup
#define msotbidTriggerDesignPopup                             1079   // Label: 'Trigger Design Popup'	 NOTE: Access: trigger design popup
#define msotbidTriggerDesign                                  1080   // Label: 'Trigger Design'	 NOTE: Access: trigger design toolbar
#define msotbidReadingOrderMenu                               1081   // Label: 'Reading Order'	 NOTE: XL, BIDI Left-to-Rignt, Right-to-Left or context reading order for the active cell or a selection
#define msotbidAddToDBCGroupMenu                              1082   // Label: 'Add to Group'	 NOTE: Access: Submenu off of right-click menu in DBC
#define msotbidDBCShortcutPopup                               1083   // Label: 'Database Shortcut Popup'	 NOTE: Access: popup in DBC
#define msotbidDBCCustomGroupPopup                            1084   // Label: 'Database Custom Group Popup'	 NOTE: Access: popup in DBC
#define msotbidFramesetToolbar                                1085   // Label: 'Frames'	 NOTE: add and remove frameset frames
#define msotbidDMToolsMenu                                    1086   // Label: 'Tools'	 NOTE: Tools menu in the FileOpen dialog
#define msotbidDMViewMenu                                     1087   // Label: 'View'	 NOTE: View menu in the FileOpen dialog
#define msotbidDMArrangeMenu                                  1088   // Label: 'Arrange'	 NOTE: ArrangeBy menu in the FileOpen dialog
#define msotbidWebToolbox                                     1089   // Label: 'Web Tools'	 NOTE: compare with ToolboxDatapageMain
#define msotbidTCSCTranslation                                1090   // Label: 'Chinese Translation'
#define msotbidSymbolBar                                      1091   // Label: 'Symbols'	 NOTE: Word: Symbol Toolbar (Chinese feature)
#define msotbidFixSynonymChange                               1092   // Label: 'Synonym'	 NOTE: Word: Popup menu synonym
#define msotbidDiacriticColorMenu                             1093   // Label: 'Diacritic Color'	 NOTE: Word: This is for the Tools/Options/RTL Dialog only.
#define msotbidSxReordering                                   1094   // Label: 'Reordering'	 NOTE: Excel: Popup Data Menu for Reordering Multi DataItems
#define msotbidDPAlignSize                                    1095   // Label: 'Alignment and Sizing'	 NOTE: Datapage alignment & sizing toolbar
#define msotbidInlineOCXPopup                                 1096   // Label: 'Inline ActiveX Control'	 NOTE: Word: Popup menu inline ocx
#define msotbidDPBackColorMenu                                1097   // Label: 'Color'	 NOTE: Color palette from Datapage format background menu (Bug 53792)
#define msotbidTableDirMenu                                   1098   // Label: 'Table Direction'	 NOTE: BiDi PP (tomian)
#define msotbidPivotChartGroupMenu                            1099   // Label: 'Group and Outline'	 NOTE: XL: pivot chart popup menu - WentaoC - O9 bug 82158
#define msotbidDummy1100DontReuse                             1100   // Label: ' '	 NOTE: I'm confused
#define msotbidToolboxDatapageFragments                       1101   // Label: 'HTML Fragments'	 NOTE: toolbox tearoff - daveand
#define msotbidViewDesignWell                                 1102   // Label: 'View Design'	 NOTE: Access View Design Command Well - cbryant
#define msotbidDiagramWell                                    1103   // Label: 'Diagram'	 NOTE: Diagram command well - cbryant
#define msotbidSxReorderingMenu                               1104   // Label: 'Order'	 NOTE: Excel popup menu - reordering in data menu - elenacam
#define msotbidFutureUse1105                                  1105   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1106                                  1106   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1107                                  1107   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1108                                  1108   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1109                                  1109   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1110                                  1110   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1111                                  1111   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1112                                  1112   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1113                                  1113   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1114                                  1114   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1115                                  1115   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1116                                  1116   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1117                                  1117   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1118                                  1118   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1119                                  1119   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1120                                  1120   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1121                                  1121   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1122                                  1122   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1123                                  1123   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1124                                  1124   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1125                                  1125   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1126                                  1126   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1127                                  1127   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1128                                  1128   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1129                                  1129   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1130                                  1130   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1131                                  1131   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1132                                  1132   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1133                                  1133   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1134                                  1134   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1135                                  1135   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1136                                  1136   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1137                                  1137   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1138                                  1138   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1139                                  1139   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1140                                  1140   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1141                                  1141   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1142                                  1142   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1143                                  1143   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1144                                  1144   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1145                                  1145   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1146                                  1146   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1147                                  1147   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1148                                  1148   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1149                                  1149   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1150                                  1150   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1151                                  1151   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1152                                  1152   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1153                                  1153   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1154                                  1154   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1155                                  1155   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1156                                  1156   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1157                                  1157   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1158                                  1158   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1159                                  1159   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1160                                  1160   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1161                                  1161   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1162                                  1162   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1163                                  1163   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1164                                  1164   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1165                                  1165   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1166                                  1166   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1167                                  1167   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1168                                  1168   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1169                                  1169   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1170                                  1170   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1171                                  1171   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1172                                  1172   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1173                                  1173   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1174                                  1174   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1175                                  1175   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1176                                  1176   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1177                                  1177   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1178                                  1178   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1179                                  1179   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1180                                  1180   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1181                                  1181   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1182                                  1182   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1183                                  1183   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1184                                  1184   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1185                                  1185   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1186                                  1186   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1187                                  1187   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1188                                  1188   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1189                                  1189   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1190                                  1190   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1191                                  1191   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1192                                  1192   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1193                                  1193   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1194                                  1194   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1195                                  1195   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1196                                  1196   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1197                                  1197   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1198                                  1198   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1199                                  1199   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1200                                  1200   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1201                                  1201   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1202                                  1202   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1203                                  1203   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1204                                  1204   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1205                                  1205   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1206                                  1206   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1207                                  1207   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1208                                  1208   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1209                                  1209   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1210                                  1210   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1211                                  1211   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1212                                  1212   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1213                                  1213   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1214                                  1214   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1215                                  1215   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1216                                  1216   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1217                                  1217   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1218                                  1218   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1219                                  1219   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1220                                  1220   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1221                                  1221   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1222                                  1222   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1223                                  1223   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1224                                  1224   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1225                                  1225   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1226                                  1226   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1227                                  1227   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1228                                  1228   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1229                                  1229   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1230                                  1230   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1231                                  1231   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1232                                  1232   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1233                                  1233   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1234                                  1234   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1235                                  1235   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1236                                  1236   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1237                                  1237   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1238                                  1238   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1239                                  1239   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1240                                  1240   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1241                                  1241   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1242                                  1242   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1243                                  1243   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1244                                  1244   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1245                                  1245   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1246                                  1246   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1247                                  1247   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1248                                  1248   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidFutureUse1249                                  1249   // Label: ' '	 NOTE: Reserved for non-core apps
#define msotbidOutlook1250                                    1250   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1251                                    1251   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1252                                    1252   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1253                                    1253   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1254                                    1254   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1255                                    1255   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1256                                    1256   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1257                                    1257   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1258                                    1258   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1259                                    1259   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1260                                    1260   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1261                                    1261   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1262                                    1262   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1263                                    1263   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1264                                    1264   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1265                                    1265   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1266                                    1266   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1267                                    1267   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1268                                    1268   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1269                                    1269   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1270                                    1270   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1271                                    1271   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1272                                    1272   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1273                                    1273   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1274                                    1274   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1275                                    1275   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1276                                    1276   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1277                                    1277   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1278                                    1278   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1279                                    1279   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1280                                    1280   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1281                                    1281   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1282                                    1282   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1283                                    1283   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1284                                    1284   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1285                                    1285   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1286                                    1286   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1287                                    1287   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1288                                    1288   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1289                                    1289   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1290                                    1290   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1291                                    1291   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1292                                    1292   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1293                                    1293   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1294                                    1294   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1295                                    1295   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1296                                    1296   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1297                                    1297   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1298                                    1298   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1299                                    1299   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidOutlook1300                                    1300   // Label: ' '	 NOTE: Reserved for Outlook
#define msotbidTable                                          1301   // Label: 'Tables'	 NOTE: NOTE: FP
#define msotbidImage                                          1302   // Label: 'Pictures'	 NOTE: NOTE: FP
#define msotbidReport                                         1303   // Label: 'Reporting'	 NOTE: NOTE: FP
#define msotbidPositioning                                    1304   // Label: 'Positioning'	 NOTE: NOTE: FP
#define msotbidDHTMLEffects                                   1305   // Label: 'DHTML Effects'	 NOTE: NOTE: FP
#define msotbidNewMenu                                        1306   // Label: 'New Menu'	 NOTE: NOTE: FP  File Menu - New Sub Menu
#define msotbidReportsViewMenu                                1307   // Label: 'Reports View Menu'	 NOTE: NOTE: FP  View Menu - Reports Sub Menu
#define msotbidInsertComponentMenu                            1308   // Label: 'Component'	 NOTE: NOTE: FP  Insert Menu - Component Sub Menu
#define msotbidInsertDatabaseMenu                             1309   // Label: 'Database'	 NOTE: NOTE: FP  Insert Menu - Database Sub Menu
#define msotbidInsertFormMenu                                 1310   // Label: 'Form'	 NOTE: NOTE: FP  Insert Menu - Form Field Sub Menu
#define msotbidInsertAdvancedMenu                             1311   // Label: 'Advanced'	 NOTE: NOTE: FP  Insert Menu - Advanced Sub Menu
#define msotbidGeneralTableInsertMenu                         1312   // Label: 'Table Insert Menu'	 NOTE: FP  Table changed AHW - Insert Sub Menu
#define msotbidGeneralTableSelectMenu                         1313   // Label: 'Table Select Menu'	 NOTE: FP  Table changed AHW - Select Sub Menu
#define msotbidGeneralTableConvertMenu                        1314   // Label: 'Table Convert Menu'	 NOTE: FP  Table changed AHW - Convert Sub Menu
#define msotbidGeneralTablePropertiesMenu                     1315   // Label: 'Table Properties Menu'	 NOTE: FP  Table changed AHW - Properties Sub Menu
#define msotbidOpenMenu                                       1316   // Label: 'Open Menu'	 NOTE: NOTE: FP  Standard Toolbar - Open Sub Menu
#define msotbidViewBarPopup                                   1317   // Label: 'View Bar Context Menu'	 NOTE: fp bulk import
#define msotbidAllFilesReportPopup                            1318   // Label: 'All Files Report Context Menu'	 NOTE: NOTE: FP  Reports View
#define msotbidAllFilesFilePopup                              1319   // Label: 'All Files File Context Menu'	 NOTE: NOTE: FP  All Files Report (Reports View)
#define msotbidTasksViewPopup                                 1320   // Label: 'Tasks View Context Menu'	 NOTE: fp bulk import
#define msotbidTaskPopup                                      1321   // Label: 'Task Context Menu'	 NOTE: NOTE: FP  Tasks View
#define msotbidFoldersViewPopup                               1322   // Label: 'Folders View Context Menu'	 NOTE: fp bulk import
#define msotbidFolderTreeViewPopup                            1323   // Label: 'Folder Tree View Context Menu'	 NOTE: NOTE: FP  Page & Folders Views
#define msotbidFolderTreePopup                                1324   // Label: 'Folder Tree Context Menu'	 NOTE: NOTE: FP  Folder Tree View
#define msotbidHyperlinksReportPopup                          1325   // Label: 'Hyperlinks Report Context Menu'	 NOTE: NOTE: FP  Reports View
#define msotbidHyperlinksHyperlinkPopup                       1326   // Label: 'Hyperlinks Hyperlink Context Menu'	 NOTE: NOTE: FP  Hyperlinks Report (Reports View)
#define msotbidHyperlinksViewPopup                            1327   // Label: 'Hyperlinks View Context Menu'	 NOTE: fp bulk import
#define msotbidHyperlinksPopup                                1328   // Label: 'Hyperlinks Context Menu'	 NOTE: NOTE: FP  Hyperlinks View
#define msotbidNavigationViewPopup                            1329   // Label: 'Navigation View Context Menu'	 NOTE: fp bulk import
#define msotbidNavigationPopup                                1330   // Label: 'Navigation Context Menu'	 NOTE: NOTE: FP  Navigation View
#define msotbidConnectionSpeedPopup                           1331   // Label: 'Connection Speed Context Menu'	 NOTE: fp bulk import
#define msotbidNormalPageViewPopup                            1332   // Label: 'Normal Page View Context Menu'	 NOTE: NOTE: FP  Page View
#define msotbidHtmlPageViewPopup                              1333   // Label: 'Html Page View Context Menu'	 NOTE: NOTE: FP  Page View
#define msotbidFileMRUMenu                                    1334   // Label: 'File MRU Menu'	 NOTE: NOTE: FP
#define msotbidWebMRUMenu                                     1335   // Label: 'Web MRU Menu'	 NOTE: NOTE: FP
#define msotbidImageWell                                      1336   // Label: 'Pictures Well'	 NOTE: FP AHW man add
#define msotbidFramesWell                                     1337   // Label: 'Frames Well'	 NOTE: FP AHW man add
#define msotbidRightDragDropExplorerToEditorPopup             1338   // Label: 'Explorer to Editor Drag and Drop'	 NOTE: FP AHW man add
#define msotbidRightDragDropEditorToEditorPopup               1339   // Label: 'Editor to Editor Drag and Drop'	 NOTE: FP AHW man add
#define msotbidRightDragDropExternalToExplorerPopup           1340   // Label: 'External to Explorer Drag and Drop'	 NOTE: FP AHW man add
#define msotbidDeco1341                                       1341   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidDeco1342                                       1342   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidDeco1343                                       1343   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidDeco1344                                       1344   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidDeco1345                                       1345   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidDeco1346                                       1346   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidDeco1347                                       1347   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidDeco1348                                       1348   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidDeco1349                                       1349   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidDeco1350                                       1350   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidDeco1351                                       1351   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidDeco1352                                       1352   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidDeco1353                                       1353   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidDeco1354                                       1354   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidDeco1355                                       1355   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidDeco1356                                       1356   // Label: ' '	 NOTE: Reserved for Deco
#define msotbidPublisher1357                                  1357   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1358                                  1358   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1359                                  1359   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1360                                  1360   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1361                                  1361   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1362                                  1362   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1363                                  1363   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1364                                  1364   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1365                                  1365   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1366                                  1366   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1367                                  1367   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1368                                  1368   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1369                                  1369   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1370                                  1370   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1371                                  1371   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1372                                  1372   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1373                                  1373   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1374                                  1374   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1375                                  1375   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidPublisher1376                                  1376   // Label: ' '	 NOTE: Reserved for Publisher
#define msotbidWebServerDiscussionMenu                        1377   // Label: 'Web Discussion Menu'	 NOTE: ravindra
#define msotbidWebServerCollabMain                            1378   // Label: 'Web Discussions'	 NOTE: ravindra
#define msotbidWebServerCollabComment                         1379   // Label: 'Web Discussions Pane'	 NOTE: ravindra
#define msotbidDBCViewPopup                                   1380   // Label: 'Database View'	 NOTE: View rt-click menu in DBC
#define msotbidDBCStoredProcedurePopup                        1381   // Label: 'Database Stored Procedure'	 NOTE: Stored Procedure rt-click in DBC
#define msotbidDBCDiagramPopup                                1382   // Label: 'Database Diagram'	 NOTE: Diagram rt-click menu in DBC
#define msotbidShortcutMenuCategory_ViewDesign                1383   // Label: 'View Design'	 NOTE: Access customize mode
#define msotbidShortcutMenuCategory_Other                     1384   // Label: 'Other'	 NOTE: Access customize mode
#define msotbidDistributionListActionMenu                     1385   // Label: 'Actions'	 NOTE: Actions drop down menu for new Distribution List Item
#define msotbidNewEnvelopeDocumentMenu                        1386   // Label: 'Microsoft Office'	 NOTE: Cascade for new mail message using Microsoft Office app
#define msotbidDialupMenu                                     1387   // Label: 'Dialup'	 NOTE: New dialup cascade menu
#define msotbidFormatEncodingMRU                              1388   // Label: 'Format'
#define msotbidViewEncodingMRU                                1389   // Label: 'View Language'
#define msotbidContactQuickFindList                           1390   // Label: 'Find a Contact'	 NOTE: New quickfind combo dropdown
#define msotbidContactLink                                    1391   // Label: 'Link'	 NOTE: Contact link cascade to items and files
#define msotbidCollectCopyItemsMenu                           1392   // Label: 'Clipboard Items Menu'	 NOTE: in docked state use will be able to get to individual items in the clipboard
#define msotbidFormFieldMenu                                  1393   // Label: 'HTML Form Field'	 NOTE: MAC
#define msotbidNewOutlookBar                                  1394   // Label: 'New Navigation Pane'	 NOTE: Outlook add
#define msotbidSynchronize                                    1395   // Label: 'Synchronize'	 NOTE: Outlook add
#define msotbidShare                                          1396   // Label: 'Share'	 NOTE: Outlook add
#define msotbidMSMailTools                                    1397   // Label: 'MSMailTools'	 NOTE: Outlook add
#define msotbidFaxTools                                       1398   // Label: 'FaxTools'	 NOTE: Outlook add
#define msotbidStyle                                          1399   // Label: 'Style'	 NOTE: Outlook add
#define msotbidAddressMenu                                    1400   // Label: 'Addr Menu'	 NOTE: Outlook add
#define msotbidSpecial                                        1401   // Label: 'Special'	 NOTE: Outlook add
#define msotbidNavigation                                     1402   // Label: 'Navigation'	 NOTE: Outlook add
#define msotbidFormatLanguage                                 1403   // Label: 'Format Language'	 NOTE: Outlook add
#define msotbidViewLanguage                                   1404   // Label: 'View Language'	 NOTE: Outlook add
#define msotbidBackColor                                      1405   // Label: 'Back Color'	 NOTE: Outlook add
#define msotbidFontZoom                                       1406   // Label: 'Font Zoom'	 NOTE: Outlook add
#define msotbidRetrieveMailFrom                               1407   // Label: 'Retrieve Mail From'	 NOTE: Outlook add
#define msotbidOMISendUsing                                   1408   // Label: 'OMI Send Using'	 NOTE: Outlook add
#define msotbidSignatureMRU                                   1409   // Label: 'Signature MRU'	 NOTE: Outlook add
#define msotbidStationeryMRU                                  1410   // Label: 'Stationery MRU'	 NOTE: Outlook add
#define msotbidToolsForm                                      1411   // Label: 'Tools Form'	 NOTE: Outlook add
#define msotbidImapDownload                                   1412   // Label: 'Imap Download'	 NOTE: Outlook add
#define msotbidJunkEmail                                      1413   // Label: 'Junk Email'	 NOTE: Outlook add
#define msotbidInternetCall                                   1414   // Label: 'Internet Call'	 NOTE: Outlook add
#define msotbidToolsMacro                                     1415   // Label: 'Tools Macro'	 NOTE: Outlook add
#define msotbidFormulaWatch                                   1416   // Label: 'Formula Watch'	 NOTE: Excel Formula Watch toolbar
#define msotbidCellWatch                                      1417   // Label: 'Watch Window'	 NOTE: Excel Cell Watch toolbar
#define msotbidFileNew                                        1418   // Label: 'FileNew'	 NOTE: Word FileNew drop down
#define msotbidDMBackMenu                                     1419   // Label: 'Back Menu'	 NOTE: File Open Back button drop down
#define msotbidLineSpacing                                    1420   // Label: 'Line spacing'	 NOTE: Line spacing drop down in Word
#define msotbidHHCPopup                                       1421   // Label: 'HHC Popup'	 NOTE: HHC Popup menu in Word
#define msotbidOrgChartToolbar                                1422   // Label: 'Organization Chart'	 NOTE: For prototyping OrgChart in PPT
#define msotbidWorkPane                                       1423   // Label: 'Task Pane'	 NOTE: The new workpane toolbar
#define msotbidWordCount                                      1424   // Label: 'Word Count'	 NOTE: Word's Word count toolbar
#define msotbidFieldList                                      1425   // Label: 'PivotTable Field List'	 NOTE: Excel Field List toolbar
#define msotbidWebPane                                        1426   // Label: 'Web Pane'	 NOTE: The web pane is a webbrowser hosted by a toolbar
#define msotbidDummy1427                                      1427   // Label: ' '	 NOTE: Formerly Speech
#define msotbidSpeechMenu                                     1428   // Label: 'Speech Menu'	 NOTE: Speech Menu
#define msotbidAutoSumMenu                                    1429   // Label: 'Auto Sum'	 NOTE: Excel AutoSum drop down
#define msotbidRecUIToolbar                                   1430   // Label: 'Reconciliation Toolbar'	 NOTE: Reconciliation UI toolbar in PPT
#define msotbidSlideFormat                                    1431   // Label: 'Format Slides'	 NOTE: Slide formatting work pane in PPT
#define msotbidDlgSlideFormatSchemeMenu                       1432   // Label: 'Color Scheme'	 NOTE: Dropdown menu for scheme gallery on slide formatting work pane in PPT
#define msotbidPasteSpecialDDMenu                             1433   // Label: 'Paste Special Dropdown'	 NOTE: Excel paste special dropdown
#define msotbidOrgChartPopup                                  1434   // Label: 'Organization Chart Popup'	 NOTE: OrgChart context menu;
#define msotbidDiagramToolbar                                 1435   // Label: 'Diagram'	 NOTE: Diagrams
#define msotbidOrgChartAddMenu                                1436   // Label: 'Add'	 NOTE: OrgChart
#define msotbidOrgChartLayoutMenu                             1437   // Label: 'Layout'	 NOTE: OrgChart Layout menu
#define msotbidDlgSlideApplyTemplateMenu                      1438   // Label: 'Apply Template'	 NOTE: Dropdown menu for templates gallery on slide formatting work pane in PPT
#define msotbidCanvasToolbar                                  1439   // Label: 'Drawing Canvas'	 NOTE: Draw canvas
#define msotbidCanvasPopup                                    1440   // Label: 'Canvas Popup'	 NOTE: Draw popup menu for canvas shape
#define msotbidColorModeMenu                                  1441   // Label: 'Color/&Grayscale'	 NOTE: PPT color mode button menu
#define msotbidPrintOptionsMenu                               1442   // Label: '&Options'	 NOTE: PPT printing options dropdown menu
#define msotbidPrintHandoutsOrderMenu                         1443   // Label: 'Printing &Order'	 NOTE: Print Preview printing order sub menu
#define msotbidPivotAutoCalcMenu                              1444   // Label: 'PivotTable AutoCalc Menu'	 NOTE: Access Pivot\AutoCalc sub menu
#define msotbidPivotChart                                     1445   // Label: 'PivotChart'	 NOTE: Access toolbar for Pivot Chart
#define msotbidPivotFormatting                                1446   // Label: 'Formatting (PivotTable/PivotChart)'	 NOTE: Access toolbar for Pivot formatting
#define msotbidClipboardContext                               1447   // Label: 'Clipboard Context'	 NOTE: Right click on the clipboard
#define msotbidFavoritesContext                               1448   // Label: 'Favorites Context'	 NOTE: Right click on the Favorites
#define msotbidEditContext                                    1449   // Label: 'Edit Context'	 NOTE: Right click on an edit control
#define msotbidWPFormatting                                   1450   // Label: 'Styles and Formatting'	 NOTE: Word: Formatting Work Pane
#define msotbidDlgSlideApplyBackgroundMenu                    1451   // Label: 'Apply Background'	 NOTE: Dropdown menu for backgrounds gallery on slide formatting work pane in PPT
#define msotbidCommentMarkerPopup                             1452   // Label: 'Comment Popup'	 NOTE: PPT popup menu for in slide comment markers
#define msotbidMailMergeMenu                                  1453   // Label: 'Mail Merge Menu'	 NOTE: Word Tools|MailMerge drop down
#define msotbidWPMMWizardStep1                                1454   // Label: 'Mail Merge'	 NOTE: Mail Merge Wizard Step #1
#define msotbidWPMMWizardStep2                                1455   // Label: 'Mail Merge'	 NOTE: Mail Merge Wizard Step #2
#define msotbidWPMMWizardStep3                                1456   // Label: 'Mail Merge'	 NOTE: Mail Merge Wizard Step #3
#define msotbidWPMMWizardStep4                                1457   // Label: 'Mail Merge'	 NOTE: Mail Merge Wizard Step #4
#define msotbidWPMMWizardStep5                                1458   // Label: 'Mail Merge'	 NOTE: Mail Merge Wizard Step #5
#define msotbidWPMMWizardStep6                                1459   // Label: 'Mail Merge'	 NOTE: Mail Merge Wizard Step #6
#define msotbidWPSlideFormatTemplate                          1460   // Label: 'Slide Design'	 NOTE: PPT: Slide formatting work pane (template tab)
#define msotbidWPSlideFormatColorScheme                       1461   // Label: 'Slide Design'	 NOTE: PPT: Slide formatting work pane (color scheme tab)
#define msotbidWPSlideFormatBackground                        1462   // Label: 'Slide Design'	 NOTE: PPT: Slide formatting work pane (background tab)
#define msotbidWPSlideFormatAnimation                         1463   // Label: 'Slide Design'	 NOTE: PPT: Slide formatting work pane (animation tab)
#define msotbidOrgChartStyleMenu                              1464   // Label: 'OrgChart Styles'	 NOTE: PPT: Styles Menu for orgcharts
#define msotbidPublishPopup                                   1465   // Label: 'Publish Context Menu'	 NOTE: FrontPage context menu for 'Publish' dialogue
#define msotbidPreviewPageViewPopup                           1466   // Label: 'Preview Page View Context Menu'	 NOTE: FrontPage context menu for 'Preview' Page View
#define msotbidFindFormatMenu                                 1467   // Label: 'Find Format'	 NOTE: XL's find format split button dropdown menu
#define msotbidReplFormatMenu                                 1468   // Label: 'Replace Format'	 NOTE: XL's replace format split button dropdown menu
#define msotbidWPStartWorking                                 1469   // Label: 'Start Working'	 NOTE: File New workpane at app boot
#define msotbidWPFileNew                                      1470   // Label: 'New'	 NOTE: File New workpane from File|New menu command
#define msotbidImportOwsComments                              1471   // Label: '&Import Web Discusisons'
#define msotbidFlushOwsComments                               1472   // Label: '&Unimport Web Discussions'
#define msotbidReportsViewFilesMenu                           1473   // Label: 'Files'	 NOTE: brings up menu of 'files' oriented reports
#define msotbidReportsViewProblemsMenu                        1474   // Label: 'Problems'	 NOTE: brings up menu of 'Problems' oriented reports
#define msotbidReportsViewWorkFlowMenu                        1475   // Label: 'WorkFlow'	 NOTE: brings up menu of 'WorkFlow' oriented reports
#define msotbidReportsViewUsageMenu                           1476   // Label: 'Usage'	 NOTE: brings up menu of 'Usage' oriented reports
#define msotbidReportsViewCustom                              1477   // Label: 'Custom...'	 NOTE: brings up menu of 'Custom' reports
#define msotbidRecUIGalleryToolbar                            1478   // Label: 'Revisions'	 NOTE: Revisions Pane in PPT
#define msotbidNavTreePagePopup                               1479   // Label: 'Navigation Tree Page Context Menu'	 NOTE: right clicking on a node in the nav tree in FP
#define msotbidNavTreeNoPagePopup                             1480   // Label: 'Navigation Tree No Page Context Menu'	 NOTE: right clicking on anything but a node in the nav tree in FP
#define msotbidTintBaseColor                                  1481   // Label: 'Tint Base Color'	 NOTE: for Escher Fill Effects dialog Tint tab color dropdown
#define msotbidInlineCanvasPopup                              1482   // Label: 'Inline Canvas'	 NOTE: for inline Canvas context menu in Word
#define msotbidOrgChartBackgroundPopup                        1483   // Label: 'OrgChart Background Context Menu'	 NOTE: PPT: Orgchart Background shape context menu
#define msotbidFolderTreeOpenMenu                             1484   // Label: 'Open'	 NOTE: FP: submenu of folder tree context menu
#define msotbidFolderTreeNewMenu                              1485   // Label: 'New'	 NOTE: FP: submenu of folder tree context menu
#define msotbidThumbnailPopup                                 1486   // Label: 'Thumbnails'	 NOTE: PPT: popup menu for thumbnail view
#define msotbidSlideGapPopup                                  1487   // Label: 'Slide Gap'	 NOTE: PPT: popup menu for thumbnail\sorter
#define msotbidReferenceMenu                                  1488   // Label: 'Reference Menu'	 NOTE: Word Insert submenu
#define msotbidLettersMailMenu                                1489   // Label: 'Letters and Mailings Menu'	 NOTE: Word Tools submenu
#define msotbidCustomGroupsMenu                               1490   // Label: 'Custom Groups'	 NOTE: Access: submenu of Edit menu
#define msotbidBorders                                        1491   // Label: 'Borders'	 NOTE: Excel: Borders toolbar
#define msotbidFormatBorder2                                  1492   // Label: 'Borders'	 NOTE: Excel: Borders swatch toolbar
#define msotbidFormatConsistency                              1493   // Label: 'Format consistency'	 NOTE: Word: FCC
#define msotbidWPFontScheme                                   1494   // Label: 'Font Schemes'	 NOTE: PUB:
#define msotbidJapaneseGreeting                               1495   // Label: 'Japanese Greetings'	 NOTE: Word: Japanese Greeting Wizard
#define msotbidWPCustomAnimation                              1496   // Label: 'Custom Animation'	 NOTE: PPT: Custom Animation Workpane
#define msotbidFormatInspectorPopup                           1497   // Label: 'Formatting Properties Popup'	 NOTE: popup menu inside Word's Format Inspector
#define msotbidWPFormatInspector                              1498   // Label: 'Reveal Formatting'	 NOTE: Word-
#define msotbidDiagramChangeShapeMenu                         1499   // Label: 'Change Shape'	 NOTE: Item on diagram toolbar
#define msotbidDiagramLayoutMenu                              1500   // Label: 'Layout'	 NOTE: Item on diagram toolbar
#define msotbidDiagramConvertToMenu                           1501   // Label: 'Convert Diagram'	 NOTE: Item on diagram toolbar
#define msotbidWPTranslate                                    1502   // Label: 'Translate'	 NOTE: WORD: translation work pane
#define msotbidWPDocRecovery                                  1503   // Label: 'Document Recovery'	 NOTE: word/excel boot after crash with diry documents
#define msotbidDpDataOutline                                  1504   // Label: 'Data Outline'	 NOTE: Access: DataAccessPages Data Outline tool
#define msotbidDpFieldList                                    1505   // Label: 'Field List'	 NOTE: Access: DataAccessPages Field List tool
#define msotbidOrgChartSelectMenu                             1506   // Label: 'Select'	 NOTE: DRAW: Select menu for orgchartt
#define msotbidWPWizardPublicationOptions                     1507   // Label: 'Publication Options'	 NOTE: Pub: wizard publication options workpane
#define msotbidWPWizardPublicationDesigns                     1508   // Label: 'Publication Designs'	 NOTE: Pub: wizard publication designs workpane
#define msotbidWPWizardColorSchemes                           1509   // Label: 'Color Schemes'	 NOTE: Pub: wizard color schemes workpane
#define msotbidWPWizardPageContent                            1510   // Label: 'Page Content'	 NOTE: Pub: wizard page content workpane
#define msotbidSampleDBsMenu                                  1511   // Label: 'Sample Databases Menu'	 NOTE: Access: Sample databases menu
#define msotbidDummyDlgSlideApplyAnimationMenu                1512   // Label: 'Animation'	 NOTE: PPT: Easy animation workpane gallery menu
#define msotbidSearchUI                                       1513   // Label: 'File Search'	 NOTE: Office: Search Workpane
#define msotbidBorderDrawTool                                 1514   // Label: 'Draw Border'	 NOTE: Excel: Borders toolbar Draw Border dropdown toolbar
#define msotbidQuickCustomizeMainMenu                         1515   // Label: 'Quick Customize Menu'	 NOTE: Mso: Add/Remove Button menu
#define msotbidDPGroupLevelPopup                              1516   // Label: 'Group Level Properties'	 NOTE: Access: Dropdown arrow on datapage section banners
#define msotbidStyleModifyTable                               1517   // Label: 'Modify Table Formatting'	 NOTE: WORD: Formatting button in ModifyStyle dialog
#define msotbidApplyRevisionMenu                              1518   // Label: 'Apply Revision Menu'	 NOTE: PPT: dropdown arrow on apply revision on reviewing toolbar
#define msotbidUnapplyRevisionMenu                            1519   // Label: 'Unapply Revision Menu'	 NOTE: PPT: dropdown arrow on unapply revision on reviewing toolbar
#define msotbidRevisionMarkerPopup                            1520   // Label: 'Revision Marker Popup'	 NOTE: PPT: popup from revision marker
#define msotbidCagSearchUI                                    1521   // Label: 'Clip Art'	 NOTE: Office:
#define msotbidGrayscale                                      1522   // Label: 'Grayscale View'	 NOTE: PPT: floating toolbar in Grayscale view
#define msotbidWPSlideLayout                                  1523   // Label: 'Slide Layout'	 NOTE: PPT: slide layout workpane
#define msotbidDlgSlideLayoutMenu                             1524   // Label: 'Slide Layout'	 NOTE: PPT: Slide layout workpane gallery menu
#define msotbidDlgCustomAnimationEffectContextMenu            1525   // Label: 'Context Menu'	 NOTE: PPT: Custom Animation effect context menu
#define msotbidDlgCustomAnimationTimelineZoomMenu             1526   // Label: 'Zoom Menu'	 NOTE: PPT: Custom Animation timeline zoom in/out menu
#define msotbidDlgCustomAnimationEffectTypeMenu               1527   // Label: 'Effect Type Menu'	 NOTE: PPT: Custom Animation add/change effect
#define msotbidDlgCustomAnimationEffectTypeEnterMRUMenu       1528   // Label: 'Effect Type - Enter MRU Menu'	 NOTE: PPT: Custom Animation add/change effect - entrances submenu
#define msotbidDlgCustomAnimationEffectTypeEmphasisMRUMenu    1529   // Label: 'Effect Type - Emphasis MRU Menu'	 NOTE: PPT: Custom Animation add/change effect - emphasis submenu
#define msotbidDlgCustomAnimationEffectTypeExitMRUMenu        1530   // Label: 'Effect Type - Exit MRU Menu'	 NOTE: PPT: Custom Animation add/change effect - exit submenu
#define msotbidDlgCustomAnimationEffectTypePathMRUMenu        1531   // Label: 'Effect Type - Path MRU Menu'	 NOTE: PPT: Custom Animation add/change effect - path submenu
#define msotbidDlgCustomAnimationEffectTypeVerbMRUMenu        1532   // Label: 'Effect Type - Verb MRU Menu'	 NOTE: PPT: Custom Animation add/change effect - verb submenu
#define msotbidAnimEffectPropsInOut                           1533   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsFromDirs8                       1534   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsToDirs8                         1535   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsFromDirs4                       1536   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsToDirs4                         1537   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsHorizVert                       1538   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsUpDownLeftRight                 1539   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsToLeftRight                     1540   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsFromLeftRight                   1541   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsAcrossDown                      1542   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsSplit                           1543   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsStretch                         1544   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsCollapse                        1545   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsStrips                          1546   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsWheel                           1547   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsZoomIn                          1548   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsZoomOut                         1549   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsFontStyle                       1550   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsGrowShrink                      1551   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsSpin                            1552   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsTransparency                    1553   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectPropsMotionPath                      1554   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectColorSwatch                          1555   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidShowAcetateMarkup                              1556   // Label: 'Show Markup'	 NOTE: WORD: Show Markup split dropdown
#define msotbidPreviewRejectAll                               1557   // Label: 'Close Reject All Preview'	 NOTE: WORD: Toolbar to remind user they want to get out of Preview Reject All mode
#define msotbidShowAcetateAuthor                              1558   // Label: 'Reviewers'	 NOTE: WORD: Revision Author Filter Menu
#define msotbidAcceptChanges                                  1559   // Label: 'Accept'	 NOTE: WORD: Split off accept button
#define msotbidRejectChanges                                  1560   // Label: 'Reject'	 NOTE: WORD: Split off reject/delete button
#define msotbidAnnotations                                    1561   // Label: 'Comments'	 NOTE: WORD: Split off reject/delete button
#define msotbidPropertySheet                                  1562   // Label: 'Property Sheet'	 NOTE: Access: Property sheet tool
#define msotbidWorkpaneOptionsMenu                            1563   // Label: 'Workpane Options Menu'	 NOTE: Mso: Options menu for the workpane.
#define msotbidDlgRecUIGalleryMenu                            1564   // Label: 'Revision Gallery Menu'	 NOTE: PPT: Reconciliation gallery menu
#define msotbidDPPagePopup                                    1565   // Label: 'Page Selection Popup'	 NOTE: Access: Datapage: Right-click menu when clicking on the caption (page is selected)
#define msotbidDPDocumentPopup                                1566   // Label: 'Document Selection Popup'	 NOTE: Access: Datapage: Right-click menu when clicking on the page background
#define msotbidDPControlPopup                                 1567   // Label: 'Control Selection Popup'	 NOTE: Access: Datapage: Right-click menu when clicking on a control
#define msotbidDPExpandoPopup                                 1568   // Label: 'Expando Selection Popup'	 NOTE: Access: Datapage: Right-click menu when clicking on an expando control
#define msotbidDPSectionPopup                                 1569   // Label: 'Section Selection Popup'	 NOTE: Access: Datapage: Right-click menu when clicking on an section
#define msotbidDPNavigationPopup                              1570   // Label: 'Navigation Control Popup'	 NOTE: Access: Datapage: Right-click menu when clicking on an navigation control
#define msotbidDPPivotPopup                                   1571   // Label: 'PivotList Popup'	 NOTE: Access: Datapage: Right-click menu when clicking on a pivot list
#define msotbidDPNavigationButtons                            1572   // Label: 'Navigation &Buttons'	 NOTE: Access:
#define msotbidPivotWell                                      1573   // Label: 'Pivot Well'	 NOTE: ACCESS: Customize Toolbar well for Pivot items
#define msotbidTablePivotTablePopup                           1574   // Label: 'Table PivotTable Popup'	 NOTE: ACCESS: Context menu for pivot views
#define msotbidTablePivotChartPopup                           1575   // Label: 'Table PivotChart Popup'	 NOTE: ACCESS: Context menu for pivot views
#define msotbidQueryPivotTablePopup                           1576   // Label: 'Query PivotTable Popup'	 NOTE: ACCESS: Context menu for pivot views
#define msotbidQueryPivotChartPopup                           1577   // Label: 'Query PivotChart Popup'	 NOTE: ACCESS: Context menu for pivot views
#define msotbidFormPivotTablePopup                            1578   // Label: 'Form PivotTable Popup'	 NOTE: ACCESS: Context menu for pivot views
#define msotbidFormPivotChartPopup                            1579   // Label: 'Form PivotChart Popup'	 NOTE: ACCESS: Context menu for pivot views
#define msotbidStoredProcedurePivotTablePopup                 1580   // Label: 'Stored Procedure PivotTable Popup'	 NOTE: ACCESS: Context menu for pivot views
#define msotbidStoredProcedurePivotChartPopup                 1581   // Label: 'Stored Procedure PivotChart Popup'	 NOTE: ACCESS: Context menu for pivot views
#define msotbidViewPivotTablePopup                            1582   // Label: 'View PivotTable Popup'	 NOTE: ACCESS: Context menu for pivot views
#define msotbidViewPivotChartPopup                            1583   // Label: 'View PivotChart Popup'	 NOTE: ACCESS: Context menu for pivot views
#define msotbidWPSlideTransition                              1584   // Label: 'Slide Transition'	 NOTE: PPT: Slide Transition workpane (invoked from Slide Show menu)
#define msotbidDummyDlgSlideApplyTransitionMenu               1585   // Label: 'Apply Slide Transition'	 NOTE: PPT: Slide Transition workpane gallery menu
#define msotbidPivotConditionalFilterMenu                     1586   // Label: 'Conditional Filter'	 NOTE: Access submenu tcid for PivotTable menu
#define msotbidPivotShowOnlyTheTopMenu                        1587   // Label: 'Show Only the Top'	 NOTE: Access submenu tcid for Conditional Filter menu
#define msotbidPivotShowOnlyTheBottomMenu                     1588   // Label: 'Show Only the Bottom'	 NOTE: Access submenu tcid for Conditional Filter menu
#define msotbidPivotCalcTotalItemFieldMenu                    1589   // Label: 'Calculated Totals and Fields'	 NOTE: Access
#define msotbidDpFieldListTop                                 1590   // Label: 'FieldList Top'	 NOTE: ACCESS: for GCC at top of DAP fieldlist
#define msotbidDpFieldListBottom                              1591   // Label: 'FieldList Bottom'	 NOTE: ACCESS: for GCC at bottom of DAP fieldlist
#define msotbidWPMMManager                                    1592   // Label: 'MultiMonitor Manager'	 NOTE: PPT: Workpane ID for MultiMonitor Manager
#define msotbidFPServerMenu                                   1593   // Label: 'Ser&ver'	 NOTE: FP: submenu off FP Tools Menu for accessing server properties
#define msotbidAddressBlockPopup                              1594   // Label: 'Address Block Popup'	 NOTE: WORD: Context menu for Address Block fields
#define msotbidGreetingLinePopup                              1595   // Label: 'Greeting Line Popup'	 NOTE: WORD: Context menu for Greeting Line fields
#define msotbidNewMenu2                                       1596   // Label: 'New Menu'	 NOTE: FP: Separate from Split-down menu for File Menu - New Sub Menu
#define msotbidFmtInspNormalPopup                             1597   // Label: 'Format Inspector Popup in Normal Mode'	 NOTE: popup menu inside Word's Format Inspector in normal mode
#define msotbidFmtInspComparePopup                            1598   // Label: 'Format Inspector Popup in Compare Mode'	 NOTE: popup menu inside Word's Format Inspector in compare mode
#define msotbidWPFormatPageBackground                         1599   // Label: 'Background'	 NOTE: Publisher: Format Background
#define msotbidObjects                                        1600   // Label: 'Objects'	 NOTE: PUB: object insertion toolbar (initially in left-dock)
#define msotbidTextFrameLinking                               1601   // Label: 'Connect Text Boxes'	 NOTE: PUB: text frame linking toolbar (initially top-docked)
#define msotbidMeasurement                                    1602   // Label: 'Measurement'	 NOTE: PUB: measurement toolbar (direct object formatting)
#define msotbidHTMLControlsMenu                               1603   // Label: 'Form Control Tool'	 NOTE: PUB: web form controls flyout toolbar (off Object toolbar and Insert menu)
#define msotbidRulerGuidesMenu                                1604   // Label: 'Ruler Guides'	 NOTE: PUB: ruler guides menu (off Arrange menu)
#define msotbidIdentityMenu                                   1605   // Label: 'Personal Information'	 NOTE: PUB: identity menu (on Insert menu)
#define msotbidCopyFitMenu                                    1606   // Label: 'AutoFit Text'	 NOTE: PUB: copyfitting menu (on Format menu)
#define msotbidSpellingMenu                                   1607   // Label: 'Spelling'	 NOTE: PUB: spelling menu (on Tools menu)
#define msotbidPrepressMenu                                   1608   // Label: 'Prepress'	 NOTE: PUB: prepress menu (on Tools menu)
#define msotbidTrappingMenu                                   1609   // Label: 'Trapping'	 NOTE: PUB: trapping menu (on Prepress menu)
#define msotbidPackAndGoMenu                                  1610   // Label: 'Pack and Go'	 NOTE: PUB: pack and go menu (on File menu)
#define msotbidChineseTranslatorMenu                          1611   // Label: 'Chinese Translator'	 NOTE: PUB: TC/SC conversion (on Language menu on Tools menu)
#define msotbidCustomShapesMenu                               1612   // Label: 'Custom Shapes'	 NOTE: PUB: custom shapes menu (on Object toolbar)
#define msotbidGeneralTableDeleteMenu                         1613   // Label: 'Delete'	 NOTE: PUB: table deletion menu (on Table menu)
#define msotbidScannerMenu                                    1614   // Label: 'Scanner'	 NOTE: PUB: Scanning images menu (on Insert menu)
#define msotbidArrangeWell                                    1615   // Label: 'Arrange'	 NOTE: PUB: customize well for Arrange menu items
#define msotbidTextFrameLinkingWell                           1616   // Label: 'Connect Text Boxes'	 NOTE: PUB: customize well for Text Frame Linking toolbar
#define msotbidPictureWell                                    1617   // Label: 'Picture'	 NOTE: PUB: customize well for Picture toolbar items
#define msotbidHTMLControlPopup                               1618   // Label: 'HTML Control Popup'	 NOTE: PUB: popup over form controls
#define msotbidHTMLCodeFragmentPopup                          1619   // Label: 'HTML Code Fragment Popup'	 NOTE: PUB: popup over HTML code fragments
#define msotbidGroupPopup                                     1620   // Label: 'Group Popup'	 NOTE: PUB: popup over groups
#define msotbidMultiplePopup                                  1621   // Label: 'Multiple Popup'	 NOTE: PUB: popup over multiple-selection
#define msotbidDesignGalleryObjectPopup                       1622   // Label: 'Design Gallery Object Popup'	 NOTE: PUB: popup over object in Design Gallery
#define msotbidDesignGalleryCategoryPopup                     1623   // Label: 'Design Gallery Category Popup'	 NOTE: PUB: popup over category in Design Gallery
#define msotbidDesignGalleryOptionsPopup                      1624   // Label: 'Design Gallery Options Popup'	 NOTE: PUB: popup over options in Design Gallery
#define msotbidWorkspacePopup                                 1625   // Label: 'Workspace Popup'	 NOTE: PUB: popup over workspace
#define msotbidTextPopupTextMenu                              1626   // Label: 'Text Popup Text Menu'	 NOTE: PUB: submenu on tbidTextPopup
#define msotbidTextPopupToolsMenu                             1627   // Label: 'Text Popup Tools Menu'	 NOTE: PUB: submenu on tbidTextPopup and TablePopup
#define msotbidTablePopupTextMenu                             1628   // Label: 'Table Popup Text Menu'	 NOTE: PUB: submenu on tbidTablePopup
#define msotbidTablePopupFrameMenu                            1629   // Label: 'Table Popup Frame Menu'	 NOTE: PUB: submenu on tbidTablePopup
#define msotbidPictPopupPictureMenu                           1630   // Label: 'Pict Popup Picture Menu'	 NOTE: PUB: submenu on tbidGraphicPopup
#define msotbidWPWizardSmartObjectDesign                      1631   // Label: 'Smart Object Design'	 NOTE: PUB: new workpane
#define msotbidWPWizardSmartObjectOptions                     1632   // Label: 'Smart Object Options'	 NOTE: PUB: new workpane
#define msotbidDRPWorkpaneLite                                1633   // Label: 'Document Recovery'	 NOTE: Mso: Document recovery
#define msotbidWebComponentMenu                               1634   // Label: 'Web Components Menu'	 NOTE: PUB: web component menu
#define msotbidWebServiceMenu                                 1635   // Label: 'Web Services Menu'	 NOTE: PUB: web services menu
#define msotbidCompareDocumentMenu                            1636   // Label: 'Co&mpare Documents'	 NOTE: WORD: File | Compare Documents menu
#define msotbidFNDesign                                       1637   // Label: 'Function Design'	 NOTE: Access: shows when editing a SQL function
#define msotbidFunctionDesignPopup                            1638   // Label: 'Function Design'	 NOTE: Access: popup when editing a function
#define msotbidAllFunctionsWell                               1639   // Label: 'All Functions Well'	 NOTE: Access: customize dialog well
#define msotbidAnimEffectFillColorSwatch                      1640   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectLineColorSwatch                      1641   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidAnimEffectTextColorSwatch                      1642   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu
#define msotbidSubformMenu                                    1643   // Label: 'Subform'	 NOTE: Access: menu tbid for View\Subform menu
#define msotbidPivotShowAsMenu                                1644   // Label: 'Show Data As'	 NOTE: Access: menutbid for PivotTable\Show Data As menu
#define msotbidOrgChartTextEditPopup                          1645   // Label: 'OrgChart Text Edit PopUp'	 NOTE: PPT word excel: OrgChart Text Edit PopUp
#define msotbidOrgChartInsertDropdown                         1646   // Label: 'Insert Shape'	 NOTE: icon: PPT word excel: OrgChart Insert Shape
#define msotbidNotesMaster                                    1647   // Label: 'Notes Master View'	 NOTE: PPT: Notes Master View floating toolbar
#define msotbidDeleteMarkerMenu                               1648   // Label: 'Delete Marker Menu'	 NOTE: PPT: Dropdown menu on Reviewing Toolbar
#define msotbidFormatBorderNoTearoff                          1649   // Label: 'Borders'	 NOTE: Word: msotbidFormatBorder except doesn't tear off
#define msotbidShadingColorNoTearoff                          1650   // Label: 'Shading'	 NOTE: Word: msotbidShadingColor except doesn't tear off
#define msotbidBorderColorNoTearoff                           1651   // Label: 'Border'	 NOTE: Word: msotbidMenuBorderColor except doesn't tear off
#define msotbidNavViewFolderTreePopup                         1652   // Label: 'Folder Tree Context Menu for Nav View'	 NOTE: FP: context menu for nav view
#define msotbidFolderPanePopup                                1653   // Label: 'Context Menu for Right Folder Pane'	 NOTE: FP: context menu for right folder pane
#define msotbidNavOnPageNewMenu                               1654   // Label: 'Submenu New for Nav tree and Nav View On Page'	 NOTE: FP: submenu when selection On Page
#define msotbidNavOffPageNewMenu                              1655   // Label: 'Submenu New for Nav tree and Nav View Off Page'	 NOTE: FP: submenu when selection Off Page
#define msotbidPaneViewsMenu                                  1656   // Label: 'Pane Views'	 NOTE: FP: Menu to change the view within the pane
#define msotbidDlgCustomAnimationEffectTypeMediaMRUMenu       1657   // Label: 'Submenu for Media Animation Effects'	 NOTE: PPT: submenu when media is selected
#define msotbidFPTableMenu                                    1658   // Label: 'Table'	 NOTE: FrontPage: Differs from msotbidTableMenu by translation (see bug #139367)
#define msotbidMotionPathContextMenu                          1659   // Label: 'Motion Path Context Menu'	 NOTE: PPT: Context menu for custom animation motion paths
#define msotbidMotionPathPointContextMenu                     1660   // Label: 'Motion Path Point Context Menu'	 NOTE: PPT: Context menu for custom animation motion paths
#define msotbidMotionPathSegmentContextMenu                   1661   // Label: 'Motion Path Segment Context Menu'	 NOTE: PPT: Context menu for custom animation motion paths
#define msotbidAcetateResolve                                 1662   // Label: 'Resolve'	 NOTE: Word: Resolution options for the Reviewing toolbar
#define msotbidProtectionOptions                              1663   // Label: 'Protection'	 NOTE: Excel: toolbar for workbook- sheet- and cell-protection options (O10 #124216)
#define msotbidAnimEffectPropsAcrossUp                        1664   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu (O10 #174133)
#define msotbidTextToSpeech                                   1665   // Label: 'Text To Speech'	 NOTE: XL: text-to-speech toolbar PGM TGetsch
#define msotbidPivotTableMainPopup                            1666   // Label: 'PivotTable Popup'	 NOTE: Access: right-click menu for PivotTable view
#define msotbidPivotChartMainPopup                            1667   // Label: 'PivotChart Popup'	 NOTE: Access: right-click menu for PivotChart view
#define msotbidFunctionDesignDatasheetPopup                   1668   // Label: 'Function Datasheet Popup'	 NOTE: ACCESS: context menu in function datasheet
#define msotbidFunctionPivotTablePopup                        1669   // Label: 'Function Pivot Table Popup'	 NOTE: ACCESS: context menu in function pivot table view
#define msotbidFunctionPivotChartPopup                        1670   // Label: 'Function Pivot Chart Popup'	 NOTE: ACCESS: context menu in function pivot chart view
#define msotbidMotionPathCustomTools                          1671   // Label: 'Custom Motion Path'	 NOTE: PPT: Custom Animation custom path submenu (O10#207146)
#define msotbidFPTasksMenu                                    1672   // Label: 'Tasks'	 NOTE: FP: Same as msotbidTasksMenu diffferent localization RAID #205337
#define msotbidFPInsertFormMenu                               1673   // Label: 'Form'	 NOTE: FP: Differs from msotbidInsertFormMenu in localization.  Office 10 #216532
#define msotbidAnimEffectPropsFontSize                        1674   // Label: 'Properties Menu'	 NOTE: PPT: Custom Animation effect properties menu office 10 #262357
#define msotbidDataSourcePane                                 1675   // Label: 'Data Sources'
#define msotbidDataSourceSearchPane                           1676   // Label: 'Find a Data Source'
#define msotbidDSCMainPopup                                   1677   // Label: 'Data Source Context Menu'	 NOTE: context menu for the data source catelog item
#define msotbidDSCSubPopup                                    1678   // Label: 'Send Data Source To'	 NOTE: Just add one in case I will use it.
#define msotbidOISEditBrightContrast                          1679   // Label: 'Brightness and Contrast'	 NOTE: OIS Brightness and Contrast work pane
#define msotbidWPMasterPages                                  1680   // Label: 'Master Pages'	 NOTE: Master pages workpane
#define msotbidXmlDocMap                                      1681   // Label: 'XML Structure'	 NOTE: Xml Document work pane
#define msotbidWordMailStandard                               1682   // Label: 'E-mail'	 NOTE: Default WordMail formatting toolbar
#define msotbidOISEditColorBal                                1683   // Label: ' '	 NOTE: unused
#define msotbidOISEditHueSat                                  1684   // Label: 'Color'	 NOTE: OIS Hue and Saturation work pane
#define msotbidOISEditRename                                  1685   // Label: 'Rename'	 NOTE: OIS Rename work pane
#define msotbidOISEditRotate                                  1686   // Label: 'Rotate and Flip'	 NOTE: OIS Rotate work pane
#define msotbidOISEditRedEye                                  1687   // Label: 'Red Eye Removal'	 NOTE: OIS Red Eye work pane
#define msotbidOISEditCrop                                    1688   // Label: 'Crop'	 NOTE: OIS Crop work pane
#define msotbidFpRulerMenu                                    1689   // Label: 'Ruler'	 NOTE: Ruler menu in FrontPage
#define msotbidWPDisplayMap                                   1690   // Label: 'XML Source'	 NOTE: From Excel work pane.
#define msotbidFPPackageMenu                                  1691   // Label: 'Packages'
#define msotbidFPMasterTemplates                              1692   // Label: 'Dynamic Web Template'	 NOTE: FP Master Templates toolbar
#define msotbidFPMasterTemplatesMenu                          1693   // Label: 'Dynamic Web Template Popup'	 NOTE: FP master page menu
#define msotbidWPSelectTheme                                  1694   // Label: 'Theme'	 NOTE: Format->Theme workpane in FrontPage
#define msotbidFPDataMenu                                     1695   // Label: 'Data'
#define msotbidDataSourceDetailsPane                          1696   // Label: 'Data Source Details'	 NOTE: FP DSD workpane
#define msotbidFPDataTreePopup                                1697   // Label: 'Data Tree Context Menu'	 NOTE: context menu for the data tree
#define msotbidFPDataViewPopup                                1698   // Label: 'Data View Context Menu'	 NOTE: context menu for the data tree
#define msotbidLayersPane                                     1699   // Label: 'Layers'	 NOTE: FP Layers pane
#define msotbidLayersPopUp                                    1700   // Label: 'Layer's context menu'	 NOTE: FP layers context menu
#define msotbidPanelsPane                                     1701   // Label: 'Panels'	 NOTE: FP Panels pane
#define msotbidTTP                                            1702   // Label: 'Template Help'	 NOTE: shows automatically when opening a file with a valid TemplateID custom doc property
#define msotbidEditorSizesPopup                               1703   // Label: 'Editor Sizes'	 NOTE: FP; activated when user clicks browser size control on status bar
#define msotbidPageNavPopup                                   1704   // Label: 'Page Menu'	 NOTE: Pub: Context Menu on Page Navigator icons
#define msotbidPageNavMasterPopup                             1705   // Label: 'Master Page Menu'	 NOTE: Pub: Context Menu on Page Navigate master-page icons
#define msotbidWPGettingStarted                               1706   // Label: 'Getting Started'	 NOTE: Getting Started workpane
#define msotbidOISNavigationPane                              1707   // Label: 'Picture Shortcuts'
#define msotbidSRP                                            1708   // Label: 'Search Results'	 NOTE: shows search results after issuing a help query
#define msotbidFPPackageImportPopup                           1709   // Label: 'Import'	 NOTE: Package Import dlg tree control context menu
#define msotbidReadingMode                                    1710   // Label: 'Reading Layout'	 NOTE: Reading Mode
#define msotbidSaveAsMenu                                     1711   // Label: 'Save As'	 NOTE: from XDocs File menu
#define msotbidXDSolutionViews                                1712   // Label: 'Solution Views'	 NOTE: from XDocs View menu
#define msotbidScriptMenu                                     1713   // Label: 'Script'	 NOTE: from XDocs Tools menu
#define msotbidXDDataEvents                                   1714   // Label: 'Data Events'	 NOTE: from XDocs Script menu
#define msotbidXDDataSourcePanePopup                          1715   // Label: 'Data Source Pane Menu'	 NOTE: right-click on XDocs Data Source Pane
#define msotbidConvertMenu                                    1716   // Label: 'Convert'	 NOTE: from XDocs Component right-click menu
#define msotbidOISShortcutPopup                               1717   // Label: ' '	 NOTE: OIS Shortcut section context menu
#define msotbidOISUnsavedImagesPopup                          1718   // Label: ' '	 NOTE: OIS Unsaved images section context menu
#define msotbidOISImageLibraryPopup                           1719   // Label: ' '	 NOTE: unused
#define msotbidOISAddToShortcutPopup                          1720   // Label: ' '	 NOTE: unused
#define msotbidOISThumbnailViewPopup                          1721   // Label: ' '	 NOTE: OIS Thumbnail view context menu
#define msotbidOISSingleImageViewPopup                        1722   // Label: ' '	 NOTE: OIS Single image view context menu
#define msotbidOISZoomSliderToolbar                           1723   // Label: ' '	 NOTE: unused
#define msotbidOISZoomMini                                    1724   // Label: ' '	 NOTE: unused
#define msotbidWPRefPane                                      1725   // Label: 'Research'	 NOTE: Tools:Research
#define msotbidWPDocWorkspace                                 1726   // Label: 'Document Workspace'	 NOTE: DWS netpane
#define msotbidReadingModeMarkup                              1727   // Label: ' '	 NOTE: Reading Mode
#define msotbidOISEditHome                                    1728   // Label: 'Edit Pictures'	 NOTE: OIS Edit Home pane
#define msotbidJotBulletListContextMenu                       1729   // Label: 'Bullet List Context Menu'	 NOTE: Jot's Bullet List Item Context Menu
#define msotbidJotNumberListContextMenu                       1730   // Label: 'Number List Context Menu'	 NOTE: Jot's Number List Item Context Menu
#define msotbidJotOutlineElementHandleContextMenu             1731   // Label: 'Outline Element Handle Context Menu'	 NOTE: Jot's Outline Element Handle Context Menu
#define msotbidJotBodyTextHandleContextMenu                   1732   // Label: 'Body Text Handle Context Menu'	 NOTE: Jot's Body Text Handle Context Menu
#define msotbidJotAudioRecordingToolbar                       1733   // Label: 'Audio Recording'	 NOTE: One of Jot's Toolbars
#define msotbidJotAudioPlaybackToolbar                        1734   // Label: 'Audio Recording'	 NOTE: One of Jot's Toolbars
#define msotbidJotPenToolbar                                  1735   // Label: 'Drawing and Writing Tools'	 NOTE: One of Jot's Toolbars
#define msotbidJotCommandWellFile                             1736   // Label: 'File Command Well'	 NOTE: toolbar name of the category of the file category of controls in the toolbar well for jot
#define msotbidJotCommandWellView                             1737   // Label: 'View Command Well'	 NOTE: toolbar name of the category of the view category of controls in the toolbar well for jot
#define msotbidJotAudioFlyoutMenu                             1738   // Label: 'Audio Recording'	 NOTE: Flyout Recording Menu under tools
#define msotbidJotExpandOutlineElementToLevel                 1739   // Label: 'Expand to &Level'	 NOTE: Found on Outline and Outline Element Context Menus
#define msotbidJotCollapseOutlineElementToLevel               1740   // Label: 'Collaps&e to Level'	 NOTE: Found on Outline and Outline Element Context Menus
#define msotbidJotANPageTopRecordingMenu                      1741   // Label: 'Recording'	 NOTE: Menu that drops when the user clicks on the AN page top recording icon
#define msotbidFPNuiSelectTheme                               1742   // Label: 'Theme'	 NOTE: NetUI Format->Theme workpane in frontpage
#define msotbidApplyStructureMenu                             1743   // Label: 'Apply XML Elements'	 NOTE: Context menu for applying arbitrary XML in Word
#define msotbidListRange                                      1744   // Label: 'List'	 NOTE: Excel List Range toolbar
#define msotbidListTBMenu                                     1745   // Label: 'List Range'	 NOTE: List menu on Excel List Range toolbar
#define msotbidListRangePopup                                 1746   // Label: 'List Range Popup'	 NOTE: Excel List range context menu
#define msotbidDocStructNodePopup                             1747   // Label: 'XML Structure Node Popup'	 NOTE: Word XML structure right click popup
#define msotbidWPHelpSearch                                   1748   // Label: 'Help'	 NOTE: shows the new Search net pane when F1 or help is invoked
#define msotbidWPAdhocSharing                                 1749   // Label: 'Attachment Options'	 NOTE: Shows when adding an attachment to an Outlook mail item
#define msotbidXDPreviewMenu                                  1750   // Label: 'Preview'	 NOTE: XDocs File Preview menu
#define msotbidDesignMenu                                     1751   // Label: 'Design'	 NOTE: XD's Design menu
#define msotbidWPEditorFileNew                                1752   // Label: 'New Document'	 NOTE: XD's Editor File New Pane
#define msotbidWPDesignerFileNew                              1753   // Label: 'Design a Form'	 NOTE: XD's Designer File New Pane
#define msotbidWPSpelling                                     1754   // Label: 'Spelling'	 NOTE: XD's Spell Checking Pane
#define msotbidWPFindAndReplace                               1755   // Label: 'Find'	 NOTE: XD's Find and Replace Pane
#define msotbidWPViews                                        1756   // Label: 'Views'	 NOTE: XD's View Management Pane
#define msotbidWPControlToolbox                               1757   // Label: 'Controls'	 NOTE: XD's Control Toolbox Pane
#define msotbidWPThemes                                       1758   // Label: 'Themes'	 NOTE: XD's Themes Pane
#define msotbidWPSolutionControls                             1759   // Label: 'Section Controls'	 NOTE: XD's Solution Controls Pane
#define msotbidWPSolutionDefined                              1760   // Label: 'Solution'	 NOTE: XD's Solution Defined Pane
#define msotbidWPBordersAndShading                            1761   // Label: 'Borders and Shading'	 NOTE: XD's Borders and Shading Pane
#define msotbidWPAsianLayout                                  1762   // Label: 'Asian Layout'	 NOTE: XD's Asian Layout Pane
#define msotbidWPTableGrid                                    1763   // Label: 'Table Grid'	 NOTE: XD's Table Grid Pane
#define msotbidFPNuiDSSPane                                   1764   // Label: 'Find a Data Source'	 NOTE: NetUI data->insert data workpane in FrontPage.
#define msotbidFPPreviewInBrowserMenu                         1765   // Label: 'Preview in Browser'	 NOTE: FP File->Preview in Browser
#define msotbidFPTagToolbarPopup                              1766   // Label: 'Tag Selector Popup'	 NOTE: FP Tag Selector popup menu
#define msotbidFPNuiPageLayout                                1767   // Label: 'Layout Tables and Cells'	 NOTE: NetUI PageLayout workpane in FrontPage
#define msotbidNetZoneMenu                                    1768   // Label: ' '	 NOTE: Menu that pops up when Netzone menu button is pressed
#define msotbidFPNuiBehaviors                                 1769   // Label: 'Behaviors'	 NOTE: FrontPage DHTML scripting task pane
#define msotbidOLMOfficeNETDetailsPane                        1770   // Label: 'Meeting Services'	 NOTE: Outlook Online Meetings Office.NET Details
#define msotbidOLMOfficeNETSettingsPane                       1771   // Label: 'Meeting Services'	 NOTE: Outlook Online Meetings Office.NET Settings
#define msotbidOLMSelectProviderPane                          1772   // Label: 'Meeting Services'	 NOTE: Outlook Online Meetings Select Provider
#define msotbidOLMNetMeetingPane                              1773   // Label: 'Meeting Services'	 NOTE: Outlook Online Meetings NetMeeting Provider
#define msotbidOLMWindowsMediaPane                            1774   // Label: 'Meeting Services'	 NOTE: Outlook Online Meetings Windows Media Services Provider
#define msotbidOLMExchangeConferencingPane                    1775   // Label: 'Meeting Services'	 NOTE: Outlook Online Meetings Exchange Conferencing Services Provider
#define msotbidWPSmPane                                       1776   // Label: 'Document Updates'	 NOTE: Document Services (SyncMan) Office.NET Synchronization pane
#define msotbidWPMeetWorkspace                                1777   // Label: 'Meeting Workspace'	 NOTE: Outlook Meetings workspace picker workpane
#define msotbidStructuredDocumentMenu                         1778   // Label: '&XML'	 NOTE: Structured aka XML mapped document
#define msotbidEditMasterPageToolbar                          1779   // Label: 'Edit Master Pages'	 NOTE: Publisher edit master page floating toolbar
#define msotbidWPProtection                                   1780   // Label: 'Protect Document'	 NOTE: Word's document protection work pane
#define msotbidOISSortByMenu                                  1781   // Label: 'Sort By'	 NOTE: OIS Sort By cascade menu
#define msotbidOISMyImageShortcutsPopup                       1782   // Label: ' '	 NOTE: OIS My Image Shortcuts context menu
#define msotbidOISImageShortcutSubFolderPopup                 1783   // Label: ' '	 NOTE: OIS sub-folder of shortcut context menu
#define msotbidOISMyImageLibrariesPopup                       1784   // Label: ' '	 NOTE: unused
#define msotbidOISImageLibrarySubFolderPopup                  1785   // Label: ' '	 NOTE: unused
#define msotbidOISDetailsViewPopup                            1786   // Label: ' '	 NOTE: unused
#define msotbidOISBrowseWebPopup                              1787   // Label: ' '	 NOTE: unused
#define msotbidWPSmartDoc                                     1788   // Label: 'Document Actions'	 NOTE: Smart Document Pane in apps
#define msotbidFPNuiPanelProperties                           1789   // Label: 'Cell Formatting'	 NOTE: Panel Property dialog
#define msotbidSlideShowShortPopup                            1790   // Label: 'Slide Show Short Popup'	 NOTE: PPT shorter ver of slideshow context menu
#define msotbidWPDependantObjects                             1791   // Label: 'Object Dependencies'	 NOTE: NetPane for Dependent objects in Access
#define msotbidSensitivityCB                                  1792   // Label: 'Sensitivity'	 NOTE: Combobox in the item toolbar for the sensitivity and drm settings
#define msotbidJotOutlineHandleContextMenu                    1793   // Label: 'Outline Handle Context Menu'	 NOTE: Jot's Outline Handle Context Menu
#define msotbidJotApplyBulletsMenu                            1794   // Label: 'Apply Bullet Button Menu'	 NOTE: Menu off of Jot's Apply Bullet Button
#define msotbidJotApplyNumberingMenu                          1795   // Label: 'Apply Number Button Menu'	 NOTE: Menu off of Jot's Apply Number Button
#define msotbidJotPictureContextMenu                          1796   // Label: 'Picture Context Menu'	 NOTE: Jot's Picture's Context Menu
#define msotbidJotFindButtonMenu                              1797   // Label: 'Find Button Menu'	 NOTE: Menu off of Jot's Find Toolbar Button
#define msotbidJotActionItemButtonMenu                        1798   // Label: 'Note Flag Button Menu'	 NOTE: Menu off of Jot's Action Item Toolbar Button
#define msotbidJotActionItemHandleMenu                        1799   // Label: 'Note Flag Context Menu'	 NOTE: Jot's Action item context menu
#define msotbidJotNotebookManagerTaskpane                     1800   // Label: 'Section Manager'	 NOTE: Jot's Notebook manager taskpane
#define msotbidJotSelectMenu                                  1801   // Label: 'Select Menu'	 NOTE: Jot's selection menu
#define msotbidFPNuiDataSourceDetails                         1802   // Label: 'Data View Details'	 NOTE: NetUI Data Source Details pane in FrontPage
#define msotbidInsertInListSubMenu                            1803   // Label: '&Insert'	 NOTE: submenu on Excel List toolbar menu
#define msotbidDeleteInListSubMenu                            1804   // Label: '&Delete'	 NOTE: submenu on Excel List toolbar menu
#define msotbidListSubMenu                                    1805   // Label: 'L&ist'	 NOTE: under Excel Data menu
#define msotbidXDExportToExcelMenu                            1806   // Label: 'Export'	 NOTE: Xdocs file->Export to Excel flyout
#define msotbidXDInsertSectionMenu                            1807   // Label: 'Section'	 NOTE: Xdocs Insert->Section Flyout
#define msotbidNumberingMenu                                  1808   // Label: 'Numbering'	 NOTE: Xdocs Formatting toolbar Numbering flyout
#define msotbidBulletsMenu                                    1809   // Label: 'Bullets'	 NOTE: Xdocs Formatting toolbar Bullets flyout
#define msotbidComponents                                     1810   // Label: 'Components'	 NOTE: Xdocs Components Toolbar
#define msotbidXDComponentPopup                               1811   // Label: 'Component'	 NOTE: Xdocs Component Context Menu
#define msotbidXDDSPInsertPopup                               1812   // Label: 'Insert'	 NOTE: Xdocs Data Source Pane Insert Context Menu
#define msotbidWPReplace                                      1813   // Label: 'Replace'	 NOTE: Xdocs Replace work pane, seperate from WPFindAndReplace
#define msotbidWPAsianTypo                                    1814   // Label: 'Asian Typography'	 NOTE: Xdocs Asian Typography Work Pane
#define msotbidWPGrid                                         1815   // Label: 'Grid'	 NOTE: Xdocs Grid Work Pane
#define msotbidWPPhoneticGuide                                1816   // Label: 'Phonetic Guide'	 NOTE: Xdocs Phonetic Guide Work Pane
#define msotbidWPResizeObject                                 1817   // Label: 'Resize Object'	 NOTE: Xdocs Resize Object Work Pane
#define msotbidFPSaveWebPartMenu                              1818   // Label: 'Save Web Part'	 NOTE: on fp editor contextmenu for webpart
#define msotbidFPWebPartGallery                               1819   // Label: 'Web Parts'	 NOTE: NetUI WebPart Gallery in FrontPage
#define msotbidWPFaxPanel                                     1820   // Label: 'Fax Service'	 NOTE: Outlook/WordMail Fax pane
#define msotbidMute                                           1821   // Label: 'Mute'	 NOTE: Mute menu in the infoman dialog in Outlook
#define msotbidInfoManagerActions                             1822   // Label: 'Actions'	 NOTE: Actions menu in the infoman dialog in Outlook
#define msotbidOLArrangeByMenu                                1823   // Label: 'Arrange By Menu'	 NOTE: submenu off of Outlook's View, Arrange By
#define msotbidInkMarkerPopup                                 1824   // Label: 'Slide View Ink Annotation Popup'	 NOTE: PPT slide view ink annotation context menu
#define msotbidReportsViewSharedContentMenu                   1825   // Label: 'Shared Content'	 NOTE: FP Reports view flyout menu
#define msotbidWPAlternateCoordinates                         1826   // Label: ' '	 NOTE: Used to store the coordinates for the toolbar in landscape mode
#define msotbidXDComponentConvertMenu                         1827   // Label: 'C&hange To'	 NOTE: Xdocs Convert flyout from component popup menu
#define msotbidWPDataSource                                   1828   // Label: 'Data Source'	 NOTE: Xdocs XML Data Source Pane
#define msotbidWPBulletsAndNumbering                          1829   // Label: 'Bullets and Numbering'	 NOTE: Xdocs Bullets and Numbering work pane
#define msotbidWPBIDataSource                                 1830   // Label: 'Data Sources'	 NOTE: BI Data Source pane
#define msotbidFPCodeViewMenu                                 1831   // Label: 'Code View Menu'	 NOTE: FP Edit - Code View submenu
#define msotbidFPArrangeSubMenu                               1832   // Label: 'Arrange'	 NOTE: FP any list context menu - Arrange submenu
#define msotbidWPHTMLTask                                     1833   // Label: 'Form Tasks'	 NOTE: Xdocs HTML Task pane
#define msotbidFPNuiDSCPane                                   1834   // Label: 'Data Source Catalog'	 NOTE: NetUI data->insert data workpane in FrontPage.
#define msotbidFPDataWell                                     1835   // Label: 'Data Menu Well'	 NOTE: Command Well for data menu.
#define msotbidInfoManagerRules                               1836   // Label: 'Rules'	 NOTE: TB in the rules tab of the infoman dlg
#define msotbidInfoManagerSubscriptions                       1837   // Label: 'Subscriptions'	 NOTE: TB in the subscriptions tab of the infoman dlg
#define msotbidListTBSubMenu                                  1838   // Label: '&List'	 NOTE: List menu on Excel List Range toolbar
#define msotbidWPStartHelp                                    1839   // Label: 'Help'	 NOTE: Start Help pane
#define msotbidEditHandwrittenComment                         1840   // Label: 'Handwritten Comment'	 NOTE: Excel: OBSOLETE
#define msotbidJotConvertMenu                                 1841   // Label: 'Convert Selection to Text'	 NOTE: Jot Tool-Convert Selection to Text submenu
#define msotbidJotTreatAsMenu                                 1842   // Label: 'Treat As'	 NOTE: Jot Tool-Treat As submenu
#define msotbidPubBrowserNavBar                               1843   // Label: 'Navigation'	 NOTE: TB in the new modal Trident window for Pub Services
#define msotbidOISEditResize                                  1844   // Label: 'Resize'	 NOTE: OIS Resize workpane
#define msotbidOISFileFormatOpt                               1845   // Label: 'File Format Option'	 NOTE: OIS File Format Option
#define msotbidOISEditDpiSetting                              1846   // Label: ' '	 NOTE: unused
#define msotbidRejectComments                                 1847   // Label: 'Reject Comments'	 NOTE: Word's Reading Mode
#define msotbidWPRmsStart                                     1848   // Label: 'Start'	 NOTE: Rms Start Task pane
#define msotbidWPRmsAddressEdit                               1849   // Label: 'Address Edit'	 NOTE: Rms Edit pane
#define msotbidWPRmsDesignerFormatting                        1850   // Label: 'Format'	 NOTE: Rms Designer Format Task pane
#define msotbidWPRmsMyProfile                                 1851   // Label: 'My Profile'	 NOTE: Rms My Profile Task pane
#define msotbidWPRmsExternalService                           1852   // Label: 'External Service'	 NOTE: Rms External Service
#define msotbidWPRmsExternalService2                          1853   // Label: 'External Service 2'	 NOTE: Rms External Service 2
#define msotbidWPRmsExternalService3                          1854   // Label: 'External Service 3'	 NOTE: Rms External Service 3
#define msotbidOISPublish                                     1855   // Label: 'Export'	 NOTE: OIS Publish Taskpane title
#define msotbidOISCompressPictures                            1856   // Label: 'Compress Pictures'	 NOTE: OIS Optimize (compress picture) taskpane
#define msotbidWPPubFindAndReplace                            1857   // Label: 'Find and Replace'	 NOTE: Pub find/replace workpane
#define msotbidXmlErrorPopup                                  1858   // Label: 'XML Error Options'	 NOTE: Word: a menu of options for invalid XML node
#define msotbidFpRulerAndGridMenu                             1859   // Label: 'Ruler and Grid'	 NOTE: Bring the rule and grid menu
#define msotbidFpTracingImageMenu                             1860   // Label: 'Tracing Image'	 NOTE: Image Tracing menu
#define msotbidFpRulerOriginMenu                              1861   // Label: 'Ruler's Origin'	 NOTE: Bring the ruler's origin pop up menu
#define msotbidDocStatPane                                    1862   // Label: 'Shared Workspace'	 NOTE: noicon; on RegisteredPane's dropdown
#define msotbidFollowUpMenuBar                                1863   // Label: 'Follow &Up'	 NOTE: noicon; Follow Up Menu Bar
#define msotbidFollowUpDefaultMenuBar                         1864   // Label: 'Set &Default Flag'	 NOTE: noicon; Follow Up Default Menu Bar
#define msotbidFPCondFormatWP                                 1865   // Label: 'Conditional Formatting'	 NOTE: data view task pane
#define msotbidXDDSPCMEPopup                                  1866   // Label: 'CME'	 NOTE: Data source pane, CME context menu
#define msotbidXDSectionControlsPopup                         1867   // Label: 'Section Controls'	 NOTE: Section controls context menu
#define msotbidWPBIFieldList                                  1868   // Label: 'Field List'	 NOTE: BI app Field List task pane
#define msotbidWebViewToolbar                                 1869   // Label: 'Web Tools'	 NOTE: toolbar for Publisher web view
#define msotbidPublisherServicesMenu                          1870   // Label: 'Services'	 NOTE: menu containing available Publisher services
#define msotbidWunderBarButtonConfigPopup                     1871   // Label: ' '	 NOTE: popup menu for configuration of the big wunder bar buttons
#define msotbidWunderBarAddRemoveButtonsMenu                  1872   // Label: '&Add or Remove Buttons'	 NOTE: cascading menu under ButtonConfigPopup to check or uncheck buttons
#define msotbidBaselineGuidesMenu                             1873   // Label: 'Baseline Guides'	 NOTE: menu containing the horizontal andvertical baseline guides items in FE install
#define msotbidVersionMenu                                    1874   // Label: 'Ve&rsion History'	 NOTE: Word menu to select local/STS versions
#define msotbidWPActionsPane                                  1875   // Label: 'Actions'	 NOTE: Access: NetUI actions pane
#define msotbidFpLayoutTableRowMenu                           1876   // Label: 'Layout Table Row Menu'	 NOTE: Layout table row context menu
#define msotbidFpLayoutTableColumnMenu                        1877   // Label: 'Layout Table Column Menu'	 NOTE: Layout column context menu
#define msotbidFpLayoutTablePanelMenu                         1878   // Label: 'Layout Table Panel Menu'	 NOTE: Layout table panel context menu
#define msotbidFpIntelliSenseMenu                             1879   // Label: 'IntelliSense'	 NOTE: IntelliSense Menu for code view
#define msotbidFPRegularExpressionFind                        1880   // Label: 'Find Regular Expressions'	 NOTE: Context menu in FrontPage's Edit/Find dialog
#define msotbidFPRegularExpressionFind2                       1881   // Label: 'Regular Expressions'	 NOTE: Sub-context menu in FrontPage's Edit/Find dialog
#define msotbidFPRegularExpressionReplace                     1882   // Label: 'Replace Regular Expressions'	 NOTE: Context menu in FrontPage's Edit/Replace dialog
#define msotbidFPHTMLRulesContext                             1883   // Label: 'HTML Rule'	 NOTE: Context menu in FrontPage's HTML rules dialog
#define msotbidOISUploadBarPane                               1884   // Label: 'Upload'	 NOTE: OIS upload mode button bar
#define msotbidOISEditBarPane                                 1885   // Label: 'Edit'	 NOTE: OIS edit mode button bar
#define msotbidOISUploadPane                                  1886   // Label: 'Upload'	 NOTE: OIS Upload workpane
#define msotbidFPCodeViewToolbar                              1887   // Label: 'Code View'	 NOTE: FP Code View toolbar
#define msotbidArrangeSideBySide                              1888   // Label: 'Compare Side by Side'	 NOTE: Arrange Side By Side
#define msotbidInkColorMenu                                   1889   // Label: 'Drawing and Writing Pens'	 NOTE: Escher ink color
#define msotbidMeetingRequestWith                             1890   // Label: 'New Meeting Request Wit&h'	 NOTE: Outlook menu containing list of attendees
#define msotbidSlideShowArrowOptionsMenu                      1891   // Label: 'Arrow &Options'	 NOTE: what used to be in PPT's pointer options menu; now a submenu of that menu
#define msotbidSlideShowPenWidthMenu                          1892   // Label: 'Pen &Width'	 NOTE: Ink width during slide show
#define msotbidPaneTransform                                  1893   // Label: 'XML Document'	 NOTE: workpane to apply transforms to an XML file loaded in Word
#define msotbidXML                                            1894   // Label: 'Nil'	 NOTE: not used
#define msotbidXMLSubMenu                                     1895   // Label: 'Nil'	 NOTE: not used
#define msotbidWPXDFormLayout                                 1896   // Label: 'Layout'	 NOTE: Xdocs: Form Layout task pane
#define msotbidOISEditRenameSub                               1897   // Label: 'Rename'	 NOTE: Rename subpane of Publish task pane
#define msotbidOISEditResizeSub                               1898   // Label: 'Resize'	 NOTE: Resize subpane of Publish task pane
#define msotbidXDViewManagementPopup                          1899   // Label: 'View Management'	 NOTE: Xdocs View management Context Menu
#define msotbidWPRmsContactSearch                             1900   // Label: 'Contact Search'	 NOTE: Contact Search task pane
#define msotbidWPRmsActivitySearch                            1901   // Label: 'Activity Search'	 NOTE: Activity Search task pane
#define msotbidJotNewPageMenu                                 1902   // Label: 'New Page Menu'	 NOTE: jeffreyk
#define msotbidJotPublish                                     1903   // Label: 'Publish Menu'	 NOTE: simonc
#define msotbidJotSetLanguageTaskpane                         1904   // Label: 'Set Language'	 NOTE: levl
#define msotbidJotNoteTagTaskpane                             1905   // Label: 'Note Flags Summary'	 NOTE: mskim
#define msotbidJotCustomNoteTagTaskpane                       1906   // Label: 'Customize My Note Flags'	 NOTE: mskim
#define msotbidJotPageSetupTaskpane                           1907   // Label: 'Page Setup'	 NOTE: zhenyas
#define msotbidJotNumberListTaskpane                          1908   // Label: 'Numbering'	 NOTE: fviton
#define msotbidJotNumberListCreationTaskpane                  1909   // Label: 'Customize Numbering'	 NOTE: jeffreyk
#define msotbidJotBulletListTaskpane                          1910   // Label: 'Bullets'	 NOTE: fviton
#define msotbidJotOutlineFormatTaskpane                       1911   // Label: 'Outline AutoFormat'	 NOTE: knyul
#define msotbidJotStationeryTaskpane                          1912   // Label: 'Stationery'	 NOTE: knyul
#define msotbidJotInkStyleMenu                                1913   // Label: 'Ink Style Menu'	 NOTE: jgriffin
#define msotbidJotTextContextMenu                             1914   // Label: 'Text Context Menu'	 NOTE: kentu
#define msotbidJotSeriesInfoContextMenu                       1915   // Label: 'Series Info context menu'	 NOTE: zhenyas
#define msotbidJotMovePagesMenu                               1916   // Label: 'Move Pages'	 NOTE: jeffreyk
#define msotbidJotPageContextMenu                             1917   // Label: 'Page Context Menu'	 NOTE: jeffreyk
#define msotbidJotNotesMenu                                   1918   // Label: 'Notes'	 NOTE: mskim
#define msotbidDrmTemplateMenu                                1919   // Label: 'Unprotected'	 NOTE: Drm Template Menu
#define msotbidFPOpenWithMenu                                 1920   // Label: 'Open With'	 NOTE: Menu for opening files with different editors in FrontPage
#define msotbidDeprecatedWell                                 1921   // Label: 'Deprecated Commands Well'	 NOTE: Command Well used to hold deprecated toolbar controls in FrontPage.  The string is never displayed.
#define msotbidFp2DLayoutPanels                               1922   // Label: 'Layout Tables and Cells Formatting'	 NOTE: 2D Layout toolbar
#define msotbidWPRmsDesignTemplate                            1923   // Label: 'Design Template'	 NOTE: Rms Design Template Task pane
#define msotbidOISFileProperties                              1924   // Label: 'Properties'	 NOTE: OIS Property work pane
#define msotbidFPCodeViewOptionsMenu                          1925   // Label: 'Code View Options'	 NOTE: FP Edit - Code View options sub-menu
#define msotbidWPDesignChecker                                1926   // Label: 'Design Checker'	 NOTE: Design Checker workpane in Pub
#define msotbidToolPictureMenu                                1927   // Label: 'Picture Frame'	 NOTE:  // Pub: picture flyout menu on drawing toolbar
#define msotbidJotSearchTaskpane                              1928   // Label: 'Page List'	 NOTE:  // Search taskpane
#define msotbidOISEditEmail                                   1929   // Label: 'E-mail'	 NOTE:  // OIS Email workpane
#define msotbidWPGraphicsManager                              1930   // Label: 'Graphics Manager'	 NOTE:  // Publisher Graphics Manager workpane
#define msotbidInkToolbar                                     1931   // Label: 'Ink Drawing and Writing'	 NOTE:  // Shared Ink toolbar
#define msotbidInkAnnotationToolbar                           1932   // Label: 'Ink Annotations'	 NOTE:  // Shared Ink Annotation toolbar
#define msotbidInkAnnotationColorMenu                         1933   // Label: 'Annotation Pens'	 NOTE:  //Ink Annotation color menu
#define msotbidXDErrorMenu                                    1934   // Label: 'Errors'	 NOTE: Xdocs tools->errors flyout
#define msotbidImageInURLPopup                                1935   // Label: 'Image in URL'	 NOTE: Image in URL Cntxt menu
#define msotbidWPFillOutAForm                                 1936   // Label: 'Fill Out a Form'	 NOTE: Fill out a form WP (Xdocs)
#define msotbidImportMenu                                     1937   // Label: 'Import'	 NOTE: Xdocs designer file->import flyout
#define msotbidSnorePopup                                     1938   // Label: 'XML Range Popup'	 NOTE: XML range context menu
#define msotbidXdtRangePopup                                  1939   // Label: 'XML List Popup'	 NOTE: XML List range context menu
#define msotbidRmsMyRelationshipToolbar                       1940   // Label: 'Standard'	 NOTE:  Standard Toolbar
#define msotbidRmsAddressBookToolbar                          1941   // Label: 'Address Book'	 NOTE:  Rms Address Book Toolbar
#define msotbidRmsFillColorFloat                              1942   // Label: 'Fill Color'	 NOTE:  Rms Fill Color Float
#define msotbidRmsLineColorFloat                              1943   // Label: 'Line Color'	 NOTE:  Rms Line Color Float
#define msotbidRmsFontColorFloat                              1944   // Label: 'Font Color'	 NOTE:  Rms Font Color Float
#define msotbidRmsRotateFlipFloat                             1945   // Label: 'Rotate Flip'	 NOTE:  Rms Rotate Flip Float
#define msotbidRmsOrderFloat                                  1946   // Label: 'Order'	 NOTE:  Rms Order Float
#define msotbidRmsCardDesignToolbar                           1947   // Label: 'Standard'	 NOTE:  Rms Card Design Toolbar
#define msotbidRmsFormDefaultToolbar                          1948   // Label: 'Default'	 NOTE:  Rms Form Default Toolbar
#define msotbidRmsJumpMenu                                    1949   // Label: 'Jump'	 NOTE:  Rms Jump Menu
#define msotbidRmsToolbarsMenu                                1950   // Label: 'Toolbars'	 NOTE:  Rms Toolbars Menu
#define msotbidRmsPageMenu                                    1951   // Label: 'Page'	 NOTE:  Rms Page Menu
#define msotbidRmsSearchMenu                                  1952   // Label: 'Search'	 NOTE:  Rms SearchMenu
#define msotbidRmsFormFileMenu                                1953   // Label: 'File'	 NOTE:  Rms Form File Menu
#define msotbidRmsFormViewMenu                                1954   // Label: 'View'	 NOTE:  Rms Form View Menu
#define msotbidRmsFormToolsMenu                               1955   // Label: 'Tools'	 NOTE:  Rms Form Tools Menu
#define msotbidRmsFormHelpMenu                                1956   // Label: 'Help'	 NOTE:  Rms Form Help Menu
#define msotbidRmsExternalServicesMenu                        1957   // Label: 'Help'	 NOTE:  Rms External Services Menu
#define msotbidFPXmlView                                      1958   // Label: 'XML View'	 NOTE: the toolbar that appears when you open an xml file in FP (willbr)
#define msotbidDPFLConnectionNodePopup                        1959   // Label: 'Connection'	 NOTE: Access popup menu of datapage fieldlist
#define msotbidDPFLAddToPagePopup                             1960   // Label: 'Add to Page'	 NOTE: Access popup menu of datapage fieldlist
#define msotbidDPDORelNodePopup                               1961   // Label: 'Relationship'	 NOTE: Access popup menu of datapage dataoutline
#define msotbidDPDOFldNodePopup                               1962   // Label: 'Field Node'	 NOTE: Access popup menu of datapage dataoutline
#define msotbidJotDeletedPagesContextMenu                     1963   // Label: 'Deleted Pages Menu'
#define msotbidJotActionItemTearOff                           1964   // Label: 'Note Flags'
#define msotbidJotInkStyleTearOff                             1965   // Label: 'Pens'
#define msotbidJotFileNewTaskpane                             1966   // Label: 'New'
#define msotbidJotAlternateControls                           1967   // Label: 'Alternate Controls'
#define msotbidJotSpellingTaskPane                            1968   // Label: 'Spelling'
#define msotbidJotListTaskpane                                1969   // Label: 'List'
#define msotbidJotScribblet                                   1970   // Label: 'Side Note'
#define msotbidJotFormatFontTaskPane                          1971   // Label: 'Font'
#define msotbidJotAudioNoteContextMenu                        1972   // Label: 'Audio Notes'
#define msotbidReadingModeInk                                 1973   // Label: 'Document Layout'	 NOTE: joey
#define msotbidJotInkDrawingContextMenu                       1974   // Label: 'Ink Drawing'
#define msotbidDesignAFormMenu                                1975   // Label: 'Design a Form'
#define msotbidWPFont                                         1976   // Label: 'Font'
#define msotbidApplyFontToMenu                                1977   // Label: 'Apply Font To'
#define msotbidJotStationeryPropertiesTaskpane                1978   // Label: 'Stationery Properties'
#define msotbidJotFindByDateMenu                              1979   // Label: 'Find'
#define msotbidJotAudioNoteSubMenu                            1980   // Label: 'Audio Notes'
#define msotbidWPDesignTasks                                  1981   // Label: 'Design Tasks'	 NOTE: Xdocs Design Tasks pane
#define msotbidAcetateMixModeMenu                             1982   // Label: 'Balloons'	 NOTE: WORD: Show bubbles menu
#define msotbidInsertNavBarMenu                               1983   // Label: '&Navigation Bar'	 NOTE: PUB: Insert Navbar menu
#define msotbidJotNotebookTabContextMenu                      1984   // Label: 'Section Tab Context Menu'
#define msotbidJotFolderTabContextMenu                        1985   // Label: 'Folder Tab Context Menu'
#define msotbidJotNotebookColorMenu                           1986   // Label: 'Section Color Menu'
#define msotbidFPDataFormatItemMenu                           1987   // Label: 'Format Item'	 NOTE: on fp editor contextmenu for format data view item
#define msotbidFPDataFormatHyperlinkMenu                      1988   // Label: 'Format Hyperlink'	 NOTE: on fp FormatItemAs submenu
#define msotbidFPInsertasHyperlinktoMenu                      1989   // Label: 'Insert as H&yperlink to'	 NOTE: on fp datatree contextmenu for data insertion
#define msotbidInkSplitMenu                                   1990   // Label: 'Insert Ink'	 NOTE: OfficeInkSplitMenu
#define msotbidInsertInkMenu                                  1991   // Label: 'Insert Ink'	 NOTE: OfficeInsertInkMenu
#define msotbidWordAnnotInk                                   1992   // Label: 'Ink Comment'	 NOTE: Default Ink Comment formatting toolbar
#define msotbidRulerGuidePopup                                1993   // Label: 'Ruler Guide'	 NOTE:  PUB:Conext menu for Ruler guide
#define msotbidOISHomePane                                    1994   // Label: 'Getting Started'	 NOTE: OIS Home taskpane
#define msotbidOISLocatePicturesPane                          1995   // Label: 'Locate Pictures'	 NOTE: OIS Locate Pictures taskpane
#define msotbidExchangeConnectionBar                          1996   // Label: 'Exc&hange Connection'	 NOTE: Ren menu to allow setting of fast/slow exchange connection
#define msotbidJotRecognition                                 1997   // Label: 'Recognition'
#define msotbidJotInkDrawingMenu                              1998   // Label: 'Ink Drawing'
#define msotbidJotActionItemMenu                              1999   // Label: 'Note Flag Menu'
#define msotbidJotCommandWellNotes                            2000   // Label: 'Notes'	 NOTE:  toolbar name of the category of the notes category of controls in the toolbar well for jot
#define msotbidListRangeLayoutPopup                           2001   // Label: 'List Range Layout Popup'	 NOTE: Excel List range Layout context menu
#define msotbidSnoreLayoutPopup                               2002   // Label: 'List Range Layout Popup'	 NOTE: XML range Layout context menu
#define msotbidXdtRangeLayoutPopup                            2003   // Label: 'List Range Layout Popup'	 NOTE: XML List range Layout context menu
#define msotbidInkAtn                                         2004   // Label: 'Ink Comment'	 NOTE: Word: Ink Comment
#define msotbidJotNotebookMenu                                2005   // Label: 'Section Menu'
#define msotbidJotFolderMenu                                  2006   // Label: 'Folder Menu'
#define msotbidJotRuleLinesMenu                               2007   // Label: 'Rule Lines Menu'
#define msotbidJotOEShowThroughMenu                           2008   // Label: 'Hide Levels Below Menu'
#define msotbidJotAudioNotebookButtonMenu                     2009   // Label: 'AudioNotebook Button Menu'
#define msotbidJotOutlineHandleSelectMenu                     2010   // Label: 'Outline Handle Select Menu'
#define msotbidWPRmsDesignObject                              2011   // Label: 'Insert Object'	 NOTE: Rms Insert Object
#define msotbidRmsFormEditMenu                                2012   // Label: 'Edit'	 NOTE:  Rms Form Edit Menu
#define msotbidRmsNewMenu                                     2013   // Label: 'New'	 NOTE:  Rms New Menu
#define msotbidRmsImportMenu                                  2014   // Label: 'Import'	 NOTE:  Rms Import Menu
#define msotbidRmsExportMenu                                  2015   // Label: 'Export'	 NOTE:  Rms Export Menu
#define msotbidRmsFormMenuBar                                 2016   // Label: 'Menu Bar'	 NOTE:  Rms Form Menu
#define msotbidJotInkingBiasMenu                              2017   // Label: 'Pen Mode'
#define msotbidJotOutlineFormatMenu                           2018   // Label: 'Outline Format Menu'
#define msotbidSLP                                            2019   // Label: 'Microsoft Office Online'	 NOTE: Spotlight Pane which comes up by clicking the More... link in the Getting Started Pane
#define msotbidJotPenActsAsMenu                               2020   // Label: 'Drawing and Writing Tools Menu'
#define msotbidJunkEmailMenu                                  2021   // Label: 'Junk E-mail Menu'	 NOTE:  Junk E-mail Menu
#define msotbidRMS00                                          2022   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS01                                          2023   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS02                                          2024   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS03                                          2025   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS04                                          2026   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS05                                          2027   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS06                                          2028   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS07                                          2029   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS08                                          2030   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS09                                          2031   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS10                                          2032   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS11                                          2033   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS12                                          2034   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS13                                          2035   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS14                                          2036   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS15                                          2037   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS16                                          2038   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS17                                          2039   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS18                                          2040   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS19                                          2041   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS20                                          2042   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS21                                          2043   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS22                                          2044   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS23                                          2045   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS24                                          2046   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS25                                          2047   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS26                                          2048   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS27                                          2049   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS28                                          2050   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidRMS29                                          2051   // Label: ' '	 NOTE: Reserved for RMS
#define msotbidJotNoteFlagsSubMenu                            2052   // Label: 'Note Flags Menu'
#define msotbidJotOutliningToolsWell                          2053   // Label: 'Outlining Tools'
#define msotbidJotWritingDrawingToolsWell                     2054   // Label: 'Writing and Drawing Tools'
#define msotbidFPPackageMenu2                                 2055   // Label: 'Packages'
#define msotbidXDInkToolbar                                   2056   // Label: 'Ink'
#define msotbidInkAnnotationColorFloatingMenu                 2057   // Label: 'Annotation Pens'	 NOTE:  //Ink Annotation color floating menu
#define msotbidInkColorFloatingMenu                           2058   // Label: 'Drawing and Writing Pens'	 NOTE:  //Ink Drawing and Writing color floating menu
#define msotbidMax                                            2059

#endif //__MSOMENU_H__
