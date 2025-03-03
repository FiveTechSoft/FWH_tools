#include "FiveWin.ch"
#include "dbinfo.ch"

REQUEST HB_LANG_EN
REQUEST HB_LANG_ES
REQUEST HB_LANG_IT
REQUEST HB_LANG_DE

#define adUseClient 3

#ifdef __XHARBOUR__
   #define hb_CurDrive() CurDrive()
#endif

#define STRIM( cStr, nChr ) Left( cStr, Len( cStr ) - nChr )
#define NTRIM( nNumber ) LTrim( Str( nNumber ) )

#define DT_RIGHT      2

#define STYLEBAR      2010

REQUEST DBFCDX, DESCEND

static oWndMain
static oMruDBFs
static oMruODBCConnections, oMruADOConnections
static aSearches       := {}
static cRDD
static cDefRdd         := "DBFCDX"
static lShared         := .T.
static cConnection
static oCon
static nLanguage

// Variables configuration
static oFontX
static nClrTxtBrw      := CLR_BLACK
static nClrBackBrw     := CLR_WHITE
static lPijama         := .T.
//----------------------------------------------------------------------------//
static aMariaCn        := {}
static hAdo

//----------------------------------------------------------------------------//

function Main( cDbfName )

   local oBmpTiled

   FWNumFormat( "A", .t. )

   SetAutoHelp( .T. )

   hADO := { "DBASE" => {}, "FOXPRO" => {}, "EXCEL"  => {}, "MSACCESS" => {}, ;
             "MSSQL" => {}, "MYSQL"  => {}, "SQLITE" => {}, "ORACLE"   => {}, ;
             "POSTGRE" =>  {} }

   TPreview():bSaveAsPDF := { |o,c,l| FWSavePreviewToPDF(o,c,l) }

   SetResDebug( .T. )

   cConnection = Space( 100 )

   SET DELETED OFF

   SET DATE FORMAT TO "DD/MM/YYYY"
   LoadPreferences()

   if lPijama
      SetDlgGradient( { { 1, RGB( 199, 216, 237 ), RGB( 237, 242, 248 ) } } )
   else
      SetDlgGradient( { { 1, RGB( 199, 216, 237 ), CLR_WHITE } } )
   endif

   DEFINE FONT oFontX     NAME "Calibri" SIZE 0, -14
   DEFINE BITMAP oBmpTiled RESOURCE "background"

   DEFINE WINDOW oWndMain TITLE "FiveDBU" MDI MENU BuildMenu() VSCROLL HSCROLL
   oWndMain:SetFont( oFontX )

   BuildMainBar()

   DEFINE MSGBAR PROMPT "FiveDBU 32/64 bits, (c) FiveTech Software 2014-2020" ;
      OF oWndMain 2010 KEYBOARD DATE

   if ! Empty( cDbfName )
      Open( cDbfName )
   endif

   ACTIVATE WINDOW oWndMain MAXIMIZED ;
      VALID ( FWMissingStrings(), MsgYesNo( FWString( "Want to end ?" ),;
                                            FWString( "Select an option" ) ) ) ;
      ON PAINT DrawTiled( hDC, oWndMain, oBmpTiled ) ;
      ON RIGHT CLICK BuildPopupMenu( nRow, nCol, oWndMain )

   oBmpTiled:End()

   RELEASE FONT oFontX

   AEval( aMariaCn, { |a| a[ 3 ]:Close() } )
   HEval( hAdo, { |k,v| AEval( v, { |a| a[ 3 ]:Close() } ) } )

   FErase( "checkres.txt" )
   CheckRes()
   // WinExec( "notepad checkres.txt" )

return nil

//----------------------------------------------------------------------------//

function BuildPopupMenu( nRow, nCol, oWnd )

   local oPopup, oItem

   static lServer := .F., hServer

   MENU oPopup POPUP
      if lServer
         #ifndef __XHARBOUR__
            MENUITEM "WebApp" CHECKED ACTION ( lServer := .F., hb_threadQuitRequest( hServer ) ) 
         #else   
            MENUITEM "WebApp" ACTION MsgAlert( "not available for xHarbour" )
         #endif    
      else   
         #ifndef __XHARBOUR__
            MENUITEM "WebApp" ACTION If( ! lServer, ( lServer := .T., hServer := hb_threadStart( @WebServer() ) ),)
         #else
            MENUITEM "WebApp" ACTION MsgAlert( "not available for xHarbour" )
         #endif   
      endif   
      SEPARATOR
      MENUITEM "Controller" ACTION WA_Controller()
      SEPARATOR
      MENUITEM "Dialogs" ACTION WA_ShowCodeDialogs()
   ENDMENU

   ACTIVATE POPUP oPopup OF oWnd AT nRow, nCol

return nil

//----------------------------------------------------------------------------//

#ifndef __XHARBOUR__
static function WebServer()

   local oServer := HbWebServer()

   oServer:bOnGet = { | cF, cR | WndMain():Html( cF, cR ) }
   oServer:Run()

return nil   
#endif

//----------------------------------------------------------------------------//

function BuildMainBar()

   local oBar

   if oWndMain:oBar != nil
      oWndMain:oBar:End()
   endif

   DEFINE BUTTONBAR oBar OF oWndMain STYLEBAR SIZE 70, 70

   DEFINE BUTTON OF oBar PROMPT FWString( "New" ) RESOURCE "new" ACTION New()

   DEFINE BUTTON OF oBar PROMPT FWString( "DBF" ) RESOURCE "open" ACTION Open()

   DEFINE BUTTON OF oBar PROMPT FWString( "ODBC" ) RESOURCE "open" ACTION OdbcConnect()

   DEFINE BUTTON OF oBar PROMPT FWString( "ADO" ) RESOURCE "open" ;
      MENU { || AdoTypePopup() } ACTION This:ShowPopup()

   DEFINE BUTTON OF oBar PROMPT FWString( "MySQL" ) RESOURCE "open" ;
      MENU { || MariaPopup() } ACTION If( Empty( aMariaCn ), MariaConnect(), This:ShowPopup() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Copy" ) RESOURCE "copy" ACTION Copy() GROUP

   DEFINE BUTTON OF oBar PROMPT FWString( "Paste" ) RESOURCE "paste" ACTION Paste()

   DEFINE BUTTON OF oBar PROMPT FWString( "Previous" ) RESOURCE "prev" ;
      ACTION oWndMain:PrevWindow() GROUP WHEN Len( oWndMain:oWndClient:aWnd ) > 1

   DEFINE BUTTON OF oBar PROMPT FWString( "Next" ) RESOURCE "next" ;
      ACTION oWndMain:NextWindow() WHEN Len( oWndMain:oWndClient:aWnd ) > 1

   DEFINE BUTTON OF oBar PROMPT FWString( "Exit" ) RESOURCE "exit" ;
      ACTION oWndMain:End() GROUP

return nil

//----------------------------------------------------------------------------//

function BuildMenu()

   local oMenu

   //MENU oMenu
   MENU oMenu FONT oFontX NOBORDER ;
      COLORPNEL CLR_BLACK, CLR_WHITE

      MENUITEM FWString( "Databases" )
      MENU
         MENUITEM FWString( "New" ) + "..." RESOURCE "new(16x16)" ACTION New()
         MENUITEM FWString( "Open" ) + "..." ACTION Open()
         MENUITEM FWString( "Recent files" )
         MENU
            MRU oMruDBFs ;
               FILENAME GetEnv( "APPDATA" ) + "\FiveDBU.ini" ;    // .INI to manipulate
               SECTION  "Recent DBF files" ; // The name of the INI section
               ACTION   Open( cMruItem ) ;   // cMruItem is automatically provided
               MESSAGE  FWString( "Open this file" ) ;   // The message for all of them
               SIZE     10
         ENDMENU

         SEPARATOR

         MENUITEM FWString( "ODBC open" ) + "..." ACTION OdbcConnect()

         MENUITEM FWString( "Recent ODBC connections" )
         MENU
            MRU oMruODBCConnections ;
               FILENAME GetEnv( "APPDATA" ) + "\FiveDBU.ini" ;    // .INI to manipulate
               SECTION  FWString( "Recent ODBC connections strings" ) ; // The name of the INI section
               ACTION   ( cConnection := cMruItem, OdbcOpen( "." ) ) ;   // cMruItem is automatically provided
               MESSAGE  FWString( "Connect to this ODBC database" ) ;   // The message for all of them
               SIZE     10
         ENDMENU

         MENUITEM FWString( "ADO open" ) + "..." ACTION AdoConnect()

         MENUITEM FWString( "Recent ADO connections" )
         MENU
            MRU oMruADOConnections ;
               FILENAME GetEnv( "APPDATA" ) + "\FiveDBU.ini" ;    // .INI to manipulate
               SECTION  FWString( "Recent ADO connections strings" ) ; // The name of the INI section
               ACTION   ( cConnection := cMruItem, AdoConnect( cConnection ) ) ;   // cMruItem is automatically provided
               MESSAGE  FWString( "Connect to this ADO database" ) ;   // The message for all of them
               SIZE     10
         ENDMENU

         SEPARATOR
         MENUITEM FWString( "Exit" ) ACTION oWndMain:End()
      ENDMENU

      MENUITEM "Setup"
      MENU
         MENUITEM FWString( "Preferences" ) + "..." ACTION Preferences()
         SEPARATOR
         MENUITEM "Select Printer" ACTION PrinterSetup()
         MENUITEM "&Map Network Drive" ACTION WNetConnect()
         MENUITEM "&Disconnect Network Drive" ACTION WNetDisconnectDialog()
      ENDMENU

      MENUITEM "View"
      MENU
         MENUITEM FWString( "Workareas" ) + "..." ACTION Workareas( .t. )
         MENUITEM FWString( "Sets" ) + "..."      ACTION Sets()
         MENUITEM FWString( "Console" ) + "..."   ACTION Console()
      ENDMENU

      // oMenu:AddEdit()
      oMenu:AddMdi()
      oMenu:AddHelp( "FiveDBU", "(c) FiveTech Software 2020" )
   ENDMENU

return oMenu

//----------------------------------------------------------------------------//

static function AdoTypePopup()

   local oPop

   MENU oPop POPUP STYLEBAR FONT oFontX
      HEval( hADO, { |k,v,i| AdoMenuItem( k, v, i ) } )
   ENDMENU

return oPop

//----------------------------------------------------------------------------//

static function AdoMenuItem( cRdbms, aConnected )

   local oItem, aCn, n

   if Empty( aConnected )
      MENUITEM cRdbms ACTION AdoConnect( oMenuItem:cPrompt )
   else
      MENUITEM cRdbms
      MENU
         MENUITEM "New Connection" BOLD ACTION AdoConnect( cRdbms )
         for n := 1 to Len( aConnected )
            aCn   := aConnected[ n ]
            MENUITEM oItem PROMPT aCn[ 1 ] + If( Empty( aCn[ 2 ] ), "", ";" + aCn[ 2 ] )
            oItem:Cargo := { aCn[ 3 ], n }
            oItem:bAction := { |o| AdoTableSelect( o:Cargo[ 1 ], cRdbms, o:Cargo[ 2 ] ) }
         next
      ENDMENU
   endif

return nil

//----------------------------------------------------------------------------//

static function MariaPopup()

   local oPop, oItem, aCn, oCn, aDbs, aTables, cDb, cTable, nFor

   MENU oPop POPUP STYLEBAR FONT oFontX COLORS
      MENUITEM "New Connection" BOLD FILE "..\bitmaps\b20_open.bmp" ACTION MariaConnect()
      if !Empty( aMariaCn )
         SEPARATOR
         for nFor := 1 to Len( aMariaCn )
            aCn   := aMariaCn[ nFor ]
            oCn   := aCn[ 3 ]
            MENUITEM aCn[ 1 ] + ";" + aCn[ 2 ]
            MENU
               aDbs  := oCn:ListDbs()
               MENUITEM "DATABASES" SEPARATOR BOLD
               for each cDb in aDbs
                  MENUITEM cDb
                  if aCn[ 2 ] == "root" .or. !( cDb == "mysql" .or. cDb == "sys" .or. "_schema" $ cDb )
                     aTables := oCn:ListBaseTables( nil, cDb )
                     if !Empty( aTables )
                        MENU
                        MENUITEM "TABLES" SEPARATOR BOLD
                        for each cTable in aTables
                           MENUITEM oItem PROMPT cTable
                           oItem:Cargo := { oCn, cDb }
                           oItem:bAction := <|o|
                              local oRs
                              o:Cargo[ 1 ]:SelectDB( o:Cargo[ 2 ] )
                              oRs   := o:Cargo[ 1 ]:RowSet( "select * from `" + o:cPrompt + "`" )
                              if oRs == nil
                                 MsgInfo( "can not open " + o:cPrompt )
                              else
                                 AdoOpen( o:cPrompt, oRs )
                              endif
                              return nil
                              >
                        next
                        ENDMENU
                     endif
                  endif
               next
               SEPARATOR
               MENUITEM oItem PROMPT "Close Connection" BOLD FILE "..\bitmaps\close.bmp" ;
                  ACTION CloseMariaCn( oMenuItem:Cargo[ 1 ], "MYSQL", oMenuItem:Cargo[ 2 ] )

//                     oMenuItem:Cargo[ 1 ]:Close(), HB_ADel( aMariaCn, oMenuItem:Cargo[ 2 ], .t. ) )
               oItem:Cargo := { oCn, nFor }
            ENDMENU
         next
      endif
   ENDMENU

return oPop

//----------------------------------------------------------------------------//

static function BuildIndexesDropMenu( cAlias, oBrw )

   local oPopup, oItem, n, nTags

   MENU oPopup POPUP FONT oFontX
      MENUITEM FWString( "Natural order" ) ;
         ACTION ( ( cAlias )->( DbSetOrder( 0 ) ), oBrw:Refresh(), oBrw:SetFocus(),;
                  ( cAlias )->( Eval( oBrw:bChange ) ) ) ;
         WHEN ( oMenuItem:SetCheck( IndexOrd() == 0 ), .T. )

      if ( nTags := ( cAlias )->( OrdCount() ) ) > 0
         SEPARATOR
      endif

      for n = 1 to nTags
         if ! Empty( OrdName( n ) )
            // if ! Empty( OrdName( 1 ) )
            //    DbSetOrder( OrdName( 1 ) )
            //    DbGoTop()
            // endif
            MENUITEM oItem PROMPT OrdName( n ) ;
               ACTION ( ( cAlias )->( DbSetOrder( oMenuItem:cPrompt ) ),;
                     oBrw:Refresh(), ( cAlias )->( Eval( oBrw:bChange ) ), oBrw:SetFocus() ) ;
               WHEN ( oMenuItem:SetCheck( IndexOrd() == oMenuItem:nPos ), .T. )
            oItem:nPos = n
         endif
      next
   ENDMENU

return oPopup

//----------------------------------------------------------------------------//

function Open( cFileName )

   local oWnd, oBar, oBrw
   local oMsgBar, oMsgRecNo, oMsgDeleted, oMsgTagName, oMsgRdd
   local oPopup, oPopupZap, cAlias, cClrBack, n
   local oBtnIndexes, cFileState

   DEFAULT cFileName := cGetFile( FWString( "DBF file| *.dbf|" + ;
                                            "Access file| *.mdb|" + ;
                                            "Access 2010 file | *.accdb|" + ;
                                            "SQLite file| *.db|" ),;
                                  FWString( "Please select a database" ) )

   if Empty( cFileName )
      return nil
   endif

   if ! "." $ cFileName
      cFileName += ".dbf"
   endif

   if ! File( cFileName )
      MsgStop( FWString( "File not found" ) + ": " + cFileName )
      return nil
   endif

   if Upper( cFileExt( cFileName( cFileName ) ) ) == "MDB" .or. ; // Access old
      Upper( cFileExt( cFileName( cFileName ) ) ) == "DB" .or. ;  // SQLite
      Upper( cFileExt( cFileName( cFileName ) ) ) == "ACCDB"      // Access 2010
      return OdbcOpen( cFileName )
   endif

   if Upper( cFileExt( cFileName( cFileName ) ) ) == "SCX"
      cRDD = "DBFCDX"
      RddSetDefault( "DBFCDX" )
      RddInfo( RDDI_MEMOEXT, ".sct" )

   elseif File( cFileSetExt( cFileName, "CDX" ) ) .or. File( cFileSetExt( cFileName, "FPT" ) )
      cRDD = "DBFCDX"

   elseif File( cFileSetExt( cFileName, "DBT" ) ) .or. File( cFileSetExt( cFileName, "NTX" ) )
      cRDD = "DBFNTX"

   else
      cRDD = cDefRDD
   endif

   cAlias  = cGetNewAlias( cFileNoExt( cFileName ) )
   if lShared
      USE ( cFileName ) VIA cRDD NEW SHARED ALIAS ( cAlias )
   else
      USE ( cFileName ) VIA cRDD NEW EXCLUSIVE ALIAS ( cAlias )
   endif

   if RddInfo( RDDI_MEMOEXT ) == ".sct"
      RddInfo( RDDI_MEMOEXT, ".fpt" )
   endif

   oMruDBFs:Save( cFileName )

   MENU oPopupZap POPUP FONT oFontX
      MENUITEM FWSTring( "Delete for" ) + "..." ;
         ACTION ( cAlias )->( DeleteFor( cAlias, cFileName ) ),;
                ( cAlias )->( Eval( oBrw:bChange ) )
      MENUITEM FWSTring( "Recall" ) ;
         ACTION ( cAlias )->( Recall( oBrw, cFileName ) ),;
                ( cAlias )->( Eval( oBrw:bChange ) )
      MENUITEM "Pack" ACTION ExclusiveExec( oBrw, "PACK" )
      //Pack( oBrw, cFileName )
      MENUITEM "Zap" ACTION ExclusiveExec( oBrw, "ZAP" )
      //Zap( oBrw, cFileName )
   ENDMENU

   cFileState  := cFileSetExt( TrueName( cFileName ), "brw" )

   DEFINE WINDOW oWnd TITLE "View: " + Upper( cAlias ) MDICHILD OF oWndMain

   oWnd:SetSize( 1150, 565 )

   oWnd:bCopy = { || MsgInfo( "copy" ) }

   DEFINE BUTTONBAR oBar OF oWnd STYLEBAR SIZE 70, 70

   DEFINE BUTTON OF oBar PROMPT FWString( "Add" ) RESOURCE "add" ;
      ACTION ( ( oBrw:cAlias )->( DbAppend() ), oBrw:Refresh(), oBrw:SetFocus() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Edit" ) RESOURCE "edit" ;
      ACTION oBrw:EditSource() // ( oBrw:cAlias )->( Edit() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Del" ) RESOURCE "del" ;
      MENU oPopupZap ACTION ( oBrw:cAlias )->( DelRecord( oBrw, oMsgDeleted ) )

   DEFINE BUTTON OF oBar PROMPT FWString( "Top" ) RESOURCE "top" ;
      ACTION ( oBrw:GoTop(), oBrw:SetFocus() ) GROUP

   DEFINE BUTTON OF oBar PROMPT FWString( "Bottom" ) RESOURCE "bottom" ;
      ACTION ( oBrw:GoBottom(), oBrw:SetFocus() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Search" ) RESOURCE "search" ;
      GROUP ACTION ( cAlias )->( Search( oBrw ) )

   DEFINE BUTTON oBtnIndexes OF oBar PROMPT FWString( "Index" ) RESOURCE "index" ;
      ACTION ( cAlias )->( Indexes( oBrw, oBtnIndexes, cAlias ) )

   DEFINE BUTTON OF oBar PROMPT FWString( "Filter" ) RESOURCE "filter" ;
      ACTION ( cAlias )->( Filter( oBrw ) )

   DEFINE BUTTON OF oBar PROMPT FWString( "Relations" ) RESOURCE "relation" ;
      ACTION ( cAlias )->( Relations() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Process" ) RESOURCE "process" ;
      ACTION ( cAlias )->( Processes() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Replacement" ) RESOURCE "impexp" ;
      ACTION ( cAlias )->( GlobalReplacement( oBrw ) )

   DEFINE BUTTON OF oBar PROMPT FWString( "Struct" ) RESOURCE "struct" ;
      ACTION ( oBrw:cAlias )->( Struct( cFileName, oBrw ) ) GROUP

   DEFINE BUTTON OF oBar PROMPT FWString( "Imp/Exp" ) RESOURCE "impexp" ;
      ACTION ( oBrw:cAlias )->( ImportExport( cFileName, oBrw ) )

   DEFINE BUTTON OF oBar PROMPT FWString( "Report" ) RESOURCE "report" ;
      ACTION oBrw:Report()

   DEFINE BUTTON OF oBar PROMPT FWString( "Exit" ) RESOURCE "exit" ACTION oWnd:End() GROUP

   @ 0, 0 XBROWSE oBrw OF oWnd LINES ;
      AUTOCOLS ALIAS Alias() AUTOSORT FOOTERS NOBORDER STYLE FLAT ;
      ON CHANGE ( oMsgBar:cMsgDef := FWString( "FileName" ) + ": " + cFileName,;
                                     oMsgRecNo:SetText( "RecNo: " + ;
                                         AllTrim( Str( ( cAlias )->( RecNo() ) ) ) + " / " + ;
                                         AllTrim( Str( ( cAlias )->( OrdKeyCount() ) ) ) ),;
                                     oMsgTagName:SetText( FWString( "Ordered by" ) + ": " + ;
                                        If( ! Empty( ( cAlias )->( OrdName() ) ), ( cAlias )->( OrdName() ), "Natural order" ) ),;
                                     oMsgDeleted:SetText( If( ( oBrw:cAlias )->( Deleted() ),;
                                     FWString( "DELETED" ), FWString( "NON DELETED" ) ) ),;
                                     oMsgDeleted:SetBitmap( If( ( oBrw:cAlias )->( Deleted() ),;
                                     "deleted", "nondeleted" ) ) )


   oBrw:bOnSort         := { || oBrw:GoTop(), oBrw:Refresh() }
   obRW:bHRClickMenus   := { |o| ( oBrw:cAlias )->( BrwColRClickHeader( o ) ) }

   StyleBrowse( oBrw )

   if lPijama
      oBrw:bClrStd = { || If( oBrw:KeyNo() % 2 == 0, ;
                            { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrTxtBrw ),;
                              RGB( 198, 255, 198 ) }, ;
                            { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrTxtBrw ),;
                              RGB( 232, 255, 232 ) } ) }
      oBrw:bClrSel = { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrBackBrw ),;
                              RGB( 0x33, 0x66, 0xCC ) } }
   else
      oBrw:bClrStd := { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrTxtBrw ),;
                          nClrBackBrw } }
      oBrw:bClrSel := { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrBackBrw ),;
                           RGB( 0x33, 0x66, 0xCC ) } }
   endif
   cClrBack     := Eval( oBrw:bClrSelFocus )[ 2 ]
   oBrw:bClrSelFocus = { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrBackBrw ),;
                              cClrBack } }

   WITH OBJECT oBrw
//      :RecSelShowRecNo()
      :bRecSelClick  := { |o| o:SetOrderNatural(), o:Refresh() }
   END

   oBrw:CreateFromCode()

   if File( cFileState )
//      oBrw:RestoreState( MemoRead( cFileSetExt( cFileName, "brw" ) ) )
      BrwRestState( oBrw, cFileState )
   endif

   oBrw:SetFocus()
   oBrw:bLDblClick := { || oBrw:EditSource() } // ( oBrw:cAlias )->( Edit() ) }
   // oBrw:lMultiSelect = .T.
   oBrw:bLClickHeaders = { || Eval( oBrw:bChange ) }
   oBrw:SetChecks()
   oBrw:bKeyDown   := { | nKey, nFlags | If( nKey == VK_DELETE,;
                     ( oBrw:cAlias )->( DelRecord( oBrw, oMsgDeleted ) ),) }

   oWnd:oClient    := oBrw
   oWnd:oControl   := oBrw


   DEFINE MSGBAR oMsgBar PROMPT FWString( "FileName" ) + ": " + cFileName OF oWnd 2010

   DEFINE MSGITEM oMsgRecNo OF oMsgBar ;
          PROMPT "RecNo: " + ;
          AllTrim( Str( RecNo() ) ) + " / " + ;
          AllTrim( Str( OrdKeyCount() ) ) ;
          SIZE 150 ;
          ACTION ( cAlias )->( DlgGoTo( oBrw ) )

   DEFINE MSGITEM oMsgDeleted OF oMsgBar ;
          PROMPT If( ( oBrw:cAlias )->( Deleted() ), FWString( "DELETED" ), FWString( "NON DELETED" ) ) ;
          SIZE 120 ;
          BITMAPS If( ( oBrw:cAlias )->( Deleted() ), "deleted", "nondeleted" ) ;
          ACTION ( oBrw:cAlias )->( DelRecord( oBrw, oMsgDeleted ) )

   DEFINE MSGITEM oMsgTagName OF oMsgBar ;
          PROMPT FWString( "Ordered by" ) + ": " + ;
          If( ! Empty( ( oBrw:cAlias )->( OrdName() ) ),;
              ( oBrw:cAlias )->( OrdName() ), FWString( "Natural order" ) ) ;
          SIZE 150

   DEFINE MSGITEM oMsgRdd OF oMsgBar ;
          PROMPT "Rdd: " + ( cAlias )->( RddName() ) ;
          SIZE 120

   oBtnIndexes:oPopup := BuildIndexesDropMenu( cAlias, oBrw )

   oWnd:bPostEnd  := { || ( cAlias )->( DbCloseArea() ), WorkAreas() }

   ACTIVATE WINDOW oWnd ;
      VALID ( BrwSaveState( oBrw, cFileState ), .T. )

   WorkAreas()   


return oBrw

//----------------------------------------------------------------------------//

static function BrwColRClickHeader( oCol )

   local oBrw  := oCol:oBrw
   local oItem, cText

   MENUITEM "COLUMNS" BOLD BREAK SEPARATOR COLORPNEL CLR_BLACK, oBrw:nRecSelColor
   SEPARATOR
   MENUITEM oItem PROMPT "Add Column" ACTION ( oBrw:cAlias )->( BrwDbfAddCol( oBrw ) )
   MENUITEM oItem PROMPT "Delete Column" WHEN !Empty( oCol:Cargo ) ;
      ACTION If( MsgNoYes( "Delete column : " + oCol:cHeader ), ;
               ( oBrw:DelCol( oCol:nPos ), oBrw:Refresh() ), nil )
   MENUITEM oItem PROMPT "Change Header" WHEN !Empty( oCol:Cargo ) ;
      ACTION ( cText := PadR( oCol:cHeader, 30 ), ;
               If( MsgGet( "Enter New name", "HEADER", @cText ) .and. !Empty( cText ), ;
               ( oCol:cHeader := AllTrim( cText ), oBrw:Refresh() ), nil ) )

return nil

//----------------------------------------------------------------------------//

static function BrwDbfAddCol( oBrw )

   local cHeader  := Space( 15 )
   local cExprn   := Space( 300 )
   local uVal     := ""
   local oDlg, bExpr, lOk := .f.

   DEFINE DIALOG oDlg SIZE 600,250 PIXEL TRUEPIXEL TITLE "Add Column" ;
      FONT WndMain():oFont OF oBrw:oWnd

   @  20, 20 SAY "Header     : " GET cHeader SIZE 200,26 PIXEL OF oDlg
   @  50, 20 SAY "Expression : " GET cExprn SIZE 560,26 PIXEL OF oDlg ;
      ACTION FWExpBuilder( @cExprn ) ;
      VALID If( ! Empty( cExprn ) .and. FWCheckExpression( cExprn,,@uVal ), ;
         ( oDlg:Update(), .t. ), ( oDlg:Update(), .f. ) )

   @ 100, 20 SAY { || "Value : " + cValToChar( uVal ) } SIZE 580,26 PIXEL OF oDlg UPDATE CENTER

   @ 150, 20 BUTTON "Cancel" SIZE 100,40 PIXEL OF oDlg ACTION oDlg:End()
   @ 150,480 BUTTON "Create" SIZE 100,40 PIXEL OF oDlg ACTION ( lOk := .t., oDlg:End() ) ;
      WHEN !Empty( cHeader ) .and. !Empty( cExprn )

   ACTIVATE DIALOG oDlg CENTERED IN PARENT

   if lOK
      TRY
         cExprn   := AllTrim( cExprn )
         bExpr := &( "{|x| " + cExprn + " }" )
         WITH OBJECT oBrw:AddCol()
            :cHeader    := AllTrim( cHeader )
            :cExpr      := cExprn
            :bEditValue := bExpr
            :bHRClickMenu := oBrw:aCols[ 1 ]:bHRClickMenu
            :Cargo      := .t.
            :Adjust()
         END
         oBrw:Refresh()
      CATCH
      END
   endif

return nil

//----------------------------------------------------------------------------//

static function BrwSaveState( oBrw, cFile )

   local cText    := oBrw:SaveState( { "cExprs" } )
   local aNew     := {}
   local cNew     := ""
   local n, oCol

   for n := 1 to Len( oBrw:aCols )
      oCol  := oBrw:oCol( n )
      if oCol:Cargo == .t.
         AAdd( aNew, { oCol:cHeader, oCol:cExpr } )
      endif
   next

   if !Empty( aNew )
      cText    += ( "|" + FW_ValToExp( aNew ) )
   endif

   if cFile == nil
      cFile    := cFileSetExt( TrueName( ( oBrw:cAlias )->( DBINFO( DBI_FULLPATH ) ) ), "brw" )
   endif

   HB_MEMOWRIT( cFile, cText )

return nil

//----------------------------------------------------------------------------//

static function BrwRestState( oBrw, cFile )

   local cState, aNew, aCol, bExpr

   if cFile == nil
      cFile    := cFileSetExt( TrueName( ( oBrw:cAlias )->( DBINFO( DBI_FULLPATH ) ) ), "brw" )
   endif

   cState   := MEMOREAD( cFile )
   if "|" $ cState
      aNew     := &( TOKEN( cState, "|" ) )
      cState   := TOKEN( cState, "|", 1 )
   endif

   if !Empty( aNew )
      for each aCol in aNew
         TRY
            bExpr := &( "{|x|" + aCol[ 2 ] + "}" )
         CATCH
            bExpr := nil
         END
         if bExpr != nil
            WITH OBJECT oBrw:AddCol
               :cHeader       := aCol[ 1 ]
               :bEditValue    := bExpr
               :cExpr         := aCol[ 2 ]
               :bHRClickMenu := oBrw:aCols[ 1 ]:bHRClickMenu
               :Cargo      := .t.
               :Adjust()
            END
         endif
      next
   endif

   oBrw:RestoreState( cState )

return nil

//----------------------------------------------------------------------------//

static function ExclusiveExec( oBrw, cPackZap )

   local cDbf, cAlias1, cAlias2, cRDD
   local lShared

   if Empty( oBrw:cAlias ); return nil; endif
   if cPackZap == "PACK"
      if !MsgNoYes( "Do you want to completely remove the deleted records ?" )
         return nil
      endif
   elseif cPackZap == "ZAP"
      if !MsgNoYes( FWString( "WARNING: This will remove all records from the DBF. Are you sure ?" ) )
         return nil
      endif
   else
      return nil
   endif

   cAlias1  := oBrw:cAlias
   lShared  := ( oBrw:cAlias )->( DBINFO( DBI_SHARED ) )
   cDbf     := TrueName( ( oBrw:cAlias )->( DBINFO( DBI_FULLPATH ) ) )
   cRDD     := ( oBrw:cAlias )->( RDDNAME() )
   if !( cRDD == "DBFCDX" )
      ? "Now ready for DBFCDX only"
      return nil
   endif
   oBrw:lScreenUpdating := .f.
   oBrw:Refresh()

   if lShared
      ( cAlias1 )->( DBCLOSEAREA() )
      USE ( cDbf ) NEW EXCLUSIVE VIA ( cRDD ) ALIAS "PCKZAP"
      cAlias2  := "PCKZAP"
      if !USED()
         USE ( cDbf ) NEW SHARED VIA ( cRDD ) ALIAS ( cAlias1 )
         oBrw:GoTop()
         oBrw:lScreenUpdating := .t.
         oBrw:Refresh()
         ? "Can not use the table exclusively"
         return nil
      endif
   else
      cAlias2  := cAlias1
   endif

   if cPackZap == "ZAP"
      ( cAlias2 )->( __dbZap() )
   else
      MsgRun( "Packing DBF", cDbf, { || ( cAlias2 )->( __dbPack() ) } )
   endif

   if lShared
      ( cAlias2 )->( dbCloseArea() )
      USE ( cDbf ) NEW SHARED VIA ( cRDD ) ALIAS ( cAlias1 )
   endif

   oBrw:GoTop()
   oBrw:lScreenUpdating := .t.
   oBrw:Refresh()
   Eval( oBrw:bChange )


return nil

//----------------------------------------------------------------------------//
//----------------------------------------------------------------------------//

function OdbcConnect()

   local oDL

   cConnection = ""

   oDL = CreateObject( "Datalinks" ):PromptNew()

   if ! Empty( oDL )
      cConnection = oDL:ConnectionString
   endif

   if ! Empty( cConnection )
      OdbcOpen( "." )
   endif

return nil

//----------------------------------------------------------------------------//

function MariaConnect( cConnection )

   local cRdbms   := "MYSQL"
   local cServer, cDatabase, cUserName, cPassword, cTable
   local nAt, oRs

   if ! Empty( cConnection )
      cServer   = StrToken( cConnection, 1, ";" )
      cDatabase = StrToken( cConnection, 2, ";" )
      cUserName = StrToken( cConnection, 3, ";" )
      cPassword = Decrypt( StrToken( cConnection, 4, ";" ) )
   endif

   EDITVARS cServer, cDatabase, cUserName, cPassword TITLE "Server login"

   if ! Empty( cServer )

      if ( nAt := AScan( aMariaCn, { |a| a[ 1 ] == cServer .and. a[ 2 ] == cUserName } ) ) > 0
         oCon  := aMariaCn[ nAt, 3 ]
      else
         oCon = maria_Connect( cServer, cDatabase, cUserName, cPassword )
         if oCon == nil
            ? "Can not connect"
            return nil
         endif
         AAdd( aMariaCn, { cServer, cUserName, oCon } )
      endif
/*
      do while !Empty( cTable := MariaSelectTable( oCon ) )
         ? cTable
      enddo
*/
      XBROWSER oCon:ListTables() TITLE FWString( "Select a table" ) ;
         SELECT cTable := oBrw:aCols[ 1 ]:Value

      if ! Empty( cTable )
//         oMruADOConnections:Save( cServer + ";" + cDatabase + ";" + cUserName + ";" + Encrypt( cPassword ) )

         oRs   := oCon:RowSet( "select * from `" + cTable + "`" )
         if oRs == nil
            ? "Failed to open " + cTable
         else
            AdoOpen( cTable, oRs )
         endif
         //XBROWSER oCon:RowSet( cTable )
      endif
   endif


return nil

//----------------------------------------------------------------------------//

function MariaSelectTable( oCon )

   local oTree, oDlg, oBrw
   local cDb, cTable

   oTree := MariaTableTree( oCon )

   DEFINE DIALOG oDlg SIZE 300,600 PIXEL TRUEPIXEL

   @ 20,20 XBROWSE oBrw SIZE -20,-70 PIXEL OF oDlg ;
      DATASOURCE oTree AUTOCOLS CELL LINES NOBORDER

   WITH OBJECT oBrw
      :nStretchCol   := 1
      WITH OBJECT :aCols[ 1 ]
         :cHeader    := "DB>TABLE"
         :AddBitmap( { FWDArrow(), FWRArrow(), "..\bitmaps\tree2.bmp" } )
      END
      :CreateFromCode()
   END

   ACTIVATE DIALOG oDlg CENTERED


return cTable

//----------------------------------------------------------------------------//

function MariaTableTree( oCon )

   local oTree, oItem, aDb, aTables
   local cDb, cTable
   local cCurDb   := oCon:CurrentDB()

   aDb      := oCon:ListDbs()
   TREE oTree
      for each cDb in aDb
         TREEITEM oItem PROMPT cDb
         if !( cDb == "sys" .or. cDb == "mysql" .or. "_schema" $ cDb )
            aTables  := oCon:ListBaseTables( nil, cDb )
            if !Empty( aTables )
               TREE
                  for each cTable in aTables
                     TREEITEM cTable
                  next
               ENDTREE
            endif
         endif
         if cDb == cCurDb
            oItem:Open()
         endif
      next
   ENDTREE

return oTree

//----------------------------------------------------------------------------//

function AdoConnect( cConnection )

   static cInitDir

   local cRdbms   := "MYSQL"
   local cServer, cDatabase, cUserName, cPassword, cTable
   local oCn, oErr, oRs, nAt
   local cPw   := Space( 50 )
   local cMsg  := "Enter password"

   DEFAULT cInitDir := TrueName( ".\" )

   if ! Empty( cConnection )
      cRdbms    = StrToken( cConnection, 1, ";" )
      cServer   = StrToken( cConnection, 2, ";" )
      cDatabase = StrToken( cConnection, 3, ";" )
      cUserName = StrToken( cConnection, 4, ";" )
      cPassword = Decrypt( StrToken( cConnection, 5, ";" ) )
   endif

   if ( nAt := AScan( hADO[ cRdbms ], ;
      { |a| Lower( a[ 1 ] ) == Lower( cServer ) .and. ;
            Lower( a[ 2 ] ) == Lower( IfNil( cUserName, "" ) ) } ) ) > 0
      oCn   := hADO[ cRdbms ][ nAt, 3 ]
   else

      do case
      case cRdbms $ "DBASE,FOXPRO"
         if !Empty( cServer := cGetDir( "SELECT FOLDER", cInitDir ) )
            cInitDir := cServer + "\"
            if cRdbms == "DBASE"
               oCn   := FW_OpenAdoConnection( cServer, .t. )
            else
               oCn   := FW_OpenAdoConnection( "FOXPRO," + cServer, .t. )
            endif
         endif
      case cRdbms == "EXCEL"
         if !Empty( cServer := cGetFile( "Excel File (*.xls;*.xlsx)|*.xls;*.xlsx|", ;
            "Select Excel File" ) )

            oCn      := FW_OpenADOExcelBook( cServer, .t. )
         endif
      case cRdbms == "MSACCESS"
         if !Empty( cServer := cGetFile( "Access File (*.mdb;*.accdb)|*.mdb;*.accdb|", ;
            "Select Access Database" ) )
            oCn      := FW_OpenAdoConnection( cServer, .f., @oErr )
            if oCn == nil
               do while oErr != nil .and. "password" $ Lower( oErr:Description ) .and. ;
                  MsgGet( "MSACCESS : " + cServer, cMsg, @cPw, nil, nil, .t. )
                  //
                  oCn   := FW_OpenAdoConnection( { cServer, AllTrim( cPw ) }, .f., @oErr )
                  if oCn != nil
                     EXIT
                  endif
                  cMsg  := "Enter correct password"
               enddo
               if oCn == nil .and. oErr != nil
                  ? oErr:Description
               endif
            endif
         endif
      case cRdbms == "SQLITE"
         if !Empty( cServer := cGetFile( "(*.*)|*.*|", ;
            "Select SqLite Database" ) )

            oCn      := FW_OpenAdoConnection( { "SQLITE", nil, cServer }, .t. )
         endif
      otherwise
         EDITVARS cServer, cDatabase, cUserName, cPassword TITLE cRdbms + " : Server login"

         if ! Empty( cServer )
            oCn = FW_OpenAdoConnection( { cRdbms, cServer, cDatabase, cUserName, cPassword }, .t. )
         endif
      endcase

      if oCn != nil
         AAdd( hAdo[ cRdbms ], { cServer, cUserName, oCn } )
         nAt   := Len( hADO[ cRdbms ] )
      endif

   endif

   if oCn != nil
      AdoTableSelect( oCn, cRdbms, nAt )
   endif

return oCn

//----------------------------------------------------------------------------//

static function AdoTableSelect( oCn, cRdbms, nPos )

   local cTable, oRs
   local oDlg, oFont, oBold, oBrwDb, oBrw, oBtn
   local aDb, aTable, cDb
   local nDlgW, nDlgH, nBrwCol

   aDb      := FW_AdoCatalogs( oCn )
   if !Empty( aDb )
      cDb   := FW_AdoCurrentDB( oCn )
   endif

   aTable   := FW_AdoTables( oCn )

   nDlgW    := If( Empty( aDb ), 400, 600 )
   nDlgH    := 550
   nBrwCol  := If( Empty( aDb ),  20, 220 )

   oBold    := oFontX:Bold()

   DEFINE DIALOG oDlg SIZE nDlgW,nDlgH PIXEL TRUEPIXEL FONT oFontX ;
      TITLE "SELECT TABLE TO VIEW"

   if !Empty( aDb )
      @ 20, 20 XBROWSE oBrwDb SIZE 200,-120 PIXEL OF oDlg DATASOURCE aDb ;
         AUTOCOLS HEADERS "DATABASE" ;
         CELL LINES NOBORDER

      WITH OBJECT oBrwDB
         :nStretchCol   := 1
         :lHScroll      := .f.
         if !Empty( aDb )
            :nArrayAt := :nRowSel := AScan( aDb, { |c| c == cDb } )
         endif
         :bChange := { |o| cDb := oBrwDB:aCols[ 1 ]:Value, ;
                           oBrw:aArrayData := FW_AdoTables( oCn, cDb ), ;
                           oBrw:GoTop(), oDlg:Update() }
         :CreateFromCode()
      END
   endif

   @ 20,nBrwCol XBROWSE oBrw SIZE -20,-120 PIXEL OF oDlg DATASOURCE aTable ;
      AUTOCOLS HEADERS "TABLE" ;
      CELL LINES NOBORDER FOOTERS UPDATE

   WITH OBJECT oBrw
      :nStretchCol   := 1
      :lHScroll      := .f.
//      :RecSelShowRecNo()
      :aCols[ 1 ]:bLDClickData := { || cTable := oBrw:aCols[ 1 ]:Value, oDlg:End() }
      :bChange       := { || oDlg:Update() }
      :CreateFromCode()
   END

   @ nDlgH-90, 20 BTNBMP oBtn PROMPT { || "Open " + ;
      If( Empty( cDb ), "", cDb + "." ) + ;
      If( Empty( oBrw:aCols[ 1 ]:Value ), "none", oBrw:aCols[ 1 ]:Value ) } ;
      WHEN oBrw:nLen > 0 ;
      SIZE 170, 30 PIXEL OF oDlg FLAT UPDATE COLOR CLR_WHITE,CLR_GREEN ;
      ACTION ( cTable := oBrw:aCols[ 1 ]:Value, oDlg:End() )
   @ nDlgH - 90, nDlgW - 100 BTNBMP PROMPT "Cancel" SIZE 80,30 PIXEL OF oDlg FLAT ;
      COLOR CLR_WHITE,CLR_HRED ACTION oDlg:End()

   @ nDlgH- 50,20 BTNBMP PROMPT "CLOSE CONNECTION" SIZE nDlgW-40,30 PIXEL ;
      OF oDlg FLAT COLOR CLR_WHITE,METRO_ORANGE ;
      ACTION ( CloseAdoCn( oCn, cRdbms, nPos ), oDlg:End() )

   ACTIVATE DIALOG oDlg CENTERED
   RELEASE FONT oBold

   if ! Empty( cTable )
      if !Empty( cDb )
         FW_AdoSelectDB( oCn, cDb )
      endif
      if cRdbms $ "DBASE,FOXPRO,EXCEL,MSACCESS"
         cTable   := "[" + cTable + "]"
      elseif cRdbms == "MSSQL"
         if "." $ cTable
            cTable   := StrTran( cTable, ".", ".[" ) + "]"
         else
            cTable   := "[" + cTable + "]"
         endif
      elseif cRdbms == "MYSQL"
         cTable   := "`" + cTable + "`"
      endif
      oRs   := FW_OpenRecordSet( oCn, "select * from " + cTable )
      if oRs != nil
         AdoOpen( cTable, oRs )
      endif
   endif

return nil

//----------------------------------------------------------------------------//

static function CloseAdoCn( oCn, cRdbms, nPos )

   local nAt

   do while ( nAt := AScan( WndMain():oWndClient:aWnd, { |oWnd| IsWndAdoCn( oWnd, oCn ) } ) ) > 0
      WndMain():oWndClient:aWnd[ nAt ]:End()
   enddo

   oCn:Close()
   HB_ADel( hADO[ cRdbms ], nPos, .t. )

return nil

//----------------------------------------------------------------------------//

static function IsWndAdoCn( oWnd, oCn )

   local lRet  := .f.
   local oBrw, oRs

   if HB_ISOBJECT( oWnd:oClient ) .and. ( oBrw := oWnd:oClient ):IsKindOf( "TXBROWSE" )
      if ( oRs := oBrw:oRs ) != nil
         lRet  := ( oCn:ConnectionString == oRs:ActiveConnection:ConnectionString )
      endif
   endif

return lRet

//----------------------------------------------------------------------------//

static function CloseMariaCn( oCn, cRdbms, nPos )

   local nAt

   do while ( nAt := AScan( WndMain():oWndClient:aWnd, { |oWnd| IsWndMariaCn( oWnd, oCn ) } ) ) > 0
      WndMain():oWndClient:aWnd[ nAt ]:End()
   enddo

   oCn:Close()
   HB_ADel( aMariaCn, nPos, .t. )

return nil

//----------------------------------------------------------------------------//

static function IsWndMariaCn( oWnd, oCn )

   local lRet  := .f.
   local oBrw, oRs

   if HB_ISOBJECT( oWnd:oClient ) .and. ( oBrw := oWnd:oClient ):IsKindOf( "TXBROWSE" )
      if ( oRs := oBrw:oDbf ) != nil
         lRet  := ( oCn == oRs:oCn )
      endif
   endif

return lRet

//----------------------------------------------------------------------------//

function AdoOpen( cTable, oRs )

   local oWnd, oBar, oBrw
   local oMsgBar, oMsgRecNo, oMsgDeleted, oMsgTagName, oMsgRdd
   local oPopup, oPopupZap, cAlias, cClrBack, n
   local oBtnIndexes, oBtnQuery
   local cRdd     := If( oRs:ClassName == "TOLEAUTO", "ADO", "MYSQL" )
   local cRdbms
   local cTitle   := "Browse " + cTable

   if cRdd == "MYSQL"
      cTitle   := "FWH:" + If( oRs:oCn:IsMariaDB(), "MARIA", "MYSQL" ) + ":" + ;
                  oRs:oCn:CurrentDB() + "." + cTable

   else
      cRdd     := "ADO:" + FW_RdbmsName( oRs:ActiveConnection, .f. )
      cTitle   := cRdd + ":" + oRs:ActiveConnection:DefaultDatabase + "." + cTable

   endif

   DEFINE WINDOW oWnd TITLE cTitle MDICHILD

   oWnd:SetSize( 1150, 565 )
   oWnd:bCopy = { || MsgInfo( "copy" ) }

   DEFINE BUTTONBAR oBar OF oWnd STYLEBAR SIZE 70, 70

   DEFINE BUTTON OF oBar PROMPT FWString( "Add" ) RESOURCE "add" ACTION ( oBrw:EditSource( .t. ) )
   DEFINE BUTTON OF oBar PROMPT FWString( "Edit" ) RESOURCE "edit" ;
      ACTION oBrw:EditSource()  //ADOEdit( oBrw:oRS, cTable )

   DEFINE BUTTON OF oBar PROMPT FWString( "Del" ) RESOURCE "del" ;
      /* MENU oPopupZap */ ACTION oBrw:Delete( .t. )


   DEFINE BUTTON OF oBar PROMPT FWString( "Top" ) RESOURCE "top" ;
      ACTION ( oBrw:GoTop(), oBrw:SetFocus() ) GROUP

   DEFINE BUTTON OF oBar PROMPT FWString( "Bottom" ) RESOURCE "bottom" ;
      ACTION ( oBrw:GoBottom(), oBrw:SetFocus() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Search" ) RESOURCE "search" ;
      GROUP WHEN .f. ACTION Search( oBrw )

   DEFINE BUTTON oBtnQuery OF oBar PROMPT FWString( "Query" ) RESOURCE "help" ;
      WHEN .f. ACTION ( RSQuery( oBrw ),;
               oBtnQuery:LoadBitmaps( If( Empty( oBrw:oRS:Filter ),;
               "help", "help3" ) ), oBtnQuery:Refresh() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Index" ) RESOURCE "index" ;
      WHEN .f. MENU oPopup ACTION AdoIndexes( oBrw, oRS, cTable )

   DEFINE BUTTON OF oBar PROMPT FWString( "Filter" ) RESOURCE "filter" ;
      ACTION Filter( oBrw )

   DEFINE BUTTON OF oBar PROMPT FWString( "Relations" ) RESOURCE "relation" ;
      when .f. ACTION Relations()

   DEFINE BUTTON OF oBar PROMPT FWString( "Process" ) RESOURCE "process" ;
      ACTION Processes()

   DEFINE BUTTON OF oBar PROMPT FWString( "Struct" ) RESOURCE "struct" GROUP ;
      ACTION Struct( cTable, oBrw )

   DEFINE BUTTON OF oBar PROMPT FWString( "Imp/Exp" ) RESOURCE "impexp" ;
      WHEN .f. ACTION ImportExport()

   DEFINE BUTTON OF oBar PROMPT FWString( "Report" ) RESOURCE "report" ;
      ACTION oBrw:Report()

   DEFINE BUTTON OF oBar PROMPT FWString( "Exit" ) RESOURCE "exit" ;
      ACTION oWnd:End() GROUP

   @ 0, 0 XBROWSE oBrw OF oWnd LINES ;
      AUTOCOLS DATASOURCE oRs AUTOSORT FOOTERS ;
      ON CHANGE oMsgBar:Refresh()


   StyleBrowse( oBrw )
   if lPijama
      oBrw:bClrStd := { || If( oBrw:KeyNo() % 2 == 0, ;
                        { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrTxtBrw ),;
                              RGB( 198, 255, 198 ) }, ;
                        { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrTxtBrw ),;
                              RGB( 232, 255, 232 ) } ) }
      oBrw:bClrSel := { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrBackBrw ),;
                              RGB( 0x33, 0x66, 0xCC ) } }
   else
      oBrw:bClrStd := { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED,  ),;
                          nClrBackBrw } }

      oBrw:bClrSel := { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrBackBrw ),;
                           RGB( 0x33, 0x66, 0xCC ) } }
   endif
   cClrBack := Eval( oBrw:bClrSelFocus )[ 2 ]

   WITH OBJECT oBrw
      :bClrSelFocus := { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrBackBrw ),;
                                 cClrBack } }
      :bLDblClick := { || oBrw:EditSource() }
      :bLClickHeaders = { || Eval( oBrw:bChange ) }
      :bKeyDown := { | nKey, nFlags | If( nKey == VK_DELETE, ( oBrw:Delete( .t. ), 0 ), nil ) }

//      :RecSelShowRecNo()
      :bRecSelClick  := { |o| o:SetOrderNatural(), o:Refresh() }

      oBrw:CreateFromCode()

   END

   oWnd:oClient  := oBrw
   oWnd:oControl := oBrw

   if File( cTable + ".brw" )
      oBrw:RestoreState( MemoRead( cTable + ".brw" ) )
   endif

   DEFINE MSGBAR oMsgBar PROMPT FWString( "Table" ) + ": " + cTable OF oWnd 2010

   DEFINE MSGITEM oMsgRecNo OF oMsgBar ;
          PROMPT { || "RecNo: " + ;
                  AllTrim( Str( Int( oBrw:BookMark ) ) ) + " / " + ;
                  AllTrim( Str( oBrw:nLen ) ) } ;
          SIZE 150 ;
          ACTION ( cAlias )->( DlgGoTo( oBrw ) )

   DEFINE MSGITEM oMsgTagName OF oMsgBar ;
          PROMPT { || If( Empty( oRs:Sort ), FWString( "natural order" ), oRs:Sort ) } ;
          BITMAP "..\bitmaps\sort16.bmp" ;
          SIZE 110

   DEFINE MSGITEM oMsgRdd OF oMsgBar ;
          PROMPT cRdd ;
          SIZE 120

   oWnd:bPostEnd  := { || oRs:Close() }

   ACTIVATE WINDOW oWnd ON INIT oBrw:SetFocus() ;
      VALID ( MemoWrit( cTable + ".brw", oBrw:SaveState() ), .T. )

return oBrw

//----------------------------------------------------------------------------//

function AdoCreateTable( cMot, cTab, aFld )

    local cQuery := "CREATE TABLE " + cTab + " ( "
    local cType, i

    if cMot == "JET"
       cQuery += "Id COUNTER PRIMARY KEY, "
    elseif cMot == "MSSQL"
       cQuery += "Id INT IDENTITY PRIMARY KEY, "
    elseif cMot == "MYSQL"
       cQuery += "Id SERIAL, "
    endif

    for i = 1 to LEN( aFld )
        cType = aFld[ i, 2 ]

        do case
            case cType = "C"
                cQuery += aFld[ i, 1 ] + " VARCHAR ( " + NTRIM( aFld[ i, 3 ] ) + " ), "

            case cType = "N"
                cQuery += aFld[ i, 1 ] + " NUMERIC ( " + NTRIM( aFld[ i, 3 ] ) + ", " + NTRIM( aFld[ i, 4 ] ) + " ), "

            case cType = "D"
                cQuery += aFld[ i, 1 ] + " DATETIME, "

            case cType = "L"
                cQuery += aFld[ i, 1 ] + " INT, "
            case cType = "M"
                IF cMot == "JET"
                    cQuery += "[" + aFld[ i, 1 ] + "]" + " MEMO, "
                ELSEIF cMot == "MSSQL"
                    cQuery += "[" + aFld[ i, 1 ] + "]" + " TEXT, "
                ELSEIF cMot == "MYSQL"
                    cQuery += aFld[ i, 1 ] + " TEXT, "
                ENDIF
        endcase
    next

    cQuery = STRIM( cQuery, 2 ) + " )"

    SQLEXEC( cQuery )

return nil

//----------------------------------------------------------------------------//

function OdbcOpen( cFileName )

   local oRs, aTables := {}, cTable
   local oWnd, oBar, oBrw, oError
   local oMsgBar, oMsgRecNo, oMsgTagName, oMsgRdd
   local oPopup, oPopupZap, cAlias, n, cClrBack, nTags
   local aTokens, nAt, oBtnQuery

   if Empty( cFileName )
      return nil
   endif

   if Upper( cFileExt( cFileName( cFileName ) ) ) == "MDB"
      cConnection = "Provider='Microsoft.Jet.OLEDB.4.0'; Data Source='" + cFileName + "';"
   endif

   if Upper( cFileExt( cFileName( cFileName ) ) ) == "DB"
      cConnection = "Driver={SQLite3 ODBC Driver};Database=" + cFileName
   endif

   if Upper( cFileExt( cFileName( cFileName ) ) ) == "ACCDB"
      cConnection = "Provider='Microsoft.ACE.OLEDB.12.0'; Data Source=" + cFileName
   endif

   TRY
      if ! Empty( cConnection )
         oCon = FW_OpenAdoConnection( cConnection )
         if oCon == nil
            MsgAlert( FWString( "can't connect to" ) + " " + cConnection )
            break
         endif
         aTokens = hb_ATokens( cConnection, ";" )
         if Empty( cFileName ) .or. cFileName == "."
            if ( nAt := AScan( aTokens, { | c | "data source" $ Lower( c ) .or. ;
                 "data source" $ Lower( c ) .or. "database" $ Lower( c ) } ) ) != 0
               cFileName = hb_TokenGet( aTokens[ nAt ], 2, "=" )
            endif
         endif
      endif
      oCon:CursorLocation = 3 // adUseClient
   CATCH oError
      MsgInfo( oError:Description )
   END

   oMruODBCConnections:Save( cConnection )

   TRY
      oRs = oCon:OpenSchema( 20 ) // adSchemaTables
      oRs:Filter = "TABLE_TYPE='TABLE'"
      XBROWSER oRS TITLE FWString( "Select a table" ) ;
         COLUMNS { "TABLE_NAME" } ;
         SELECT cTable := oBrw:aCols[ 1 ]:Value
      oRs:Close()
   CATCH oError
   END

   oRs = TOleAuto():New( "ADODB.Recordset" )
   oRs:CursorType     = 1        // opendkeyset
   oRs:CursorLocation = 3        // local cache
   oRs:LockType       = 3        // lockoptimistic

   if Empty( cTable )
      return nil
   endif

   cTable = FW_QuotedColSQL( cTable )

   TRY
      oRs:Open( "SELECT * FROM " + cTable, oCon ) // Password="abc" )
   CATCH oError
      MsgInfo( oError:Description, "Aqui" )
   END

   DEFINE WINDOW oWnd TITLE "Browse " + cTable MDICHILD

   oWnd:SetSize( 1150, 565 )
   oWnd:bCopy = { || MsgInfo( "copy" ) }

   DEFINE BUTTONBAR oBar OF oWnd STYLEBAR SIZE 70, 70

   DEFINE BUTTON OF oBar PROMPT FWString( "Add" ) RESOURCE "add" ;
      ACTION ( RSAppendBlank( oBrw:oRS ),;
               Eval( oBrw:bChange ), oBrw:Refresh(), oBrw:SetFocus() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Edit" ) RESOURCE "edit" ;
      ACTION oBrw:EditSource()  //ADOEdit( oBrw:oRS, cTable )

   DEFINE BUTTON OF oBar PROMPT FWString( "Del" ) RESOURCE "del" ;
      MENU oPopupZap ACTION RSDelRecord( oBrw )

   DEFINE BUTTON OF oBar PROMPT FWString( "Top" ) RESOURCE "top" ;
      ACTION ( oBrw:GoTop(), oBrw:SetFocus() ) GROUP

   DEFINE BUTTON OF oBar PROMPT FWString( "Bottom" ) RESOURCE "bottom" ;
      ACTION ( oBrw:GoBottom(), oBrw:SetFocus() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Search" ) RESOURCE "search" ;
      GROUP ACTION Search( oBrw )

   DEFINE BUTTON oBtnQuery OF oBar PROMPT FWString( "Query" ) RESOURCE "help" ;
      ACTION ( RSQuery( oBrw ),;
               oBtnQuery:LoadBitmaps( If( Empty( oBrw:oRS:Filter ),;
               "help", "help3" ) ), oBtnQuery:Refresh() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Index" ) RESOURCE "index" ;
      MENU oPopup ACTION AdoIndexes( oBrw, oRS, cTable )

   DEFINE BUTTON OF oBar PROMPT FWString( "Filter" ) RESOURCE "filter" ;
      ACTION Filter( oBrw )

   DEFINE BUTTON OF oBar PROMPT FWString( "Relations" ) RESOURCE "relation" ;
      ACTION Relations()

   DEFINE BUTTON OF oBar PROMPT FWString( "Process" ) RESOURCE "process" ;
      ACTION Processes()

   DEFINE BUTTON OF oBar PROMPT FWString( "Struct" ) RESOURCE "struct" ;
      ACTION Struct( cFileName, oBrw ) GROUP

   DEFINE BUTTON OF oBar PROMPT FWString( "Imp/Exp" ) RESOURCE "impexp" ;
      ACTION ImportExport()

   DEFINE BUTTON OF oBar PROMPT FWString( "Report" ) RESOURCE "report" ;
      ACTION oBrw:Report()

   DEFINE BUTTON OF oBar PROMPT FWString( "Exit" ) RESOURCE "exit" ;
      ACTION oWnd:End() GROUP

   @ 0, 0 XBROWSE oBrw OF oWnd LINES ;
      AUTOCOLS RECORDSET oRs AUTOSORT ;
      ON CHANGE ( oMsgBar:cMsgDef := FWString( "FileName" ) + ": " + cFileName,;
                                     oMsgBar:Refresh(),;
                                     oMsgRecNo:SetText( "RecNo: " + ;
                                                        AllTrim( Str( If( oRS:AbsolutePosition == -3,;
                                                        oRS:RecordCount() + 1,;
                                                        oRS:AbsolutePosition ) ) ) + " / " + ;
                                                        AllTrim( Str( oRS:RecordCount() ) ) ),;
                                     oMsgTagName:SetText( FWString( "Ordered by" ) + ": " + ;
                                     If( Empty( oRs:Sort ), FWString( "natural order" ), oRs:Sort ) ) )

   for n := 1 to Len( oBrw:aCols )
      if oBrw:aCols[ n ]:cDataType == 'M'
         oBrw:aCols[ n ]:bStrData = GenLocalBlock( oBrw:aCols, n )
      endif
   next

   StyleBrowse( oBrw )
   if lPijama
      oBrw:bClrStd := { || If( oBrw:KeyNo() % 2 == 0, ;
                        { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrTxtBrw ),;
                              RGB( 198, 255, 198 ) }, ;
                        { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrTxtBrw ),;
                              RGB( 232, 255, 232 ) } ) }
      oBrw:bClrSel := { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrBackBrw ),;
                              RGB( 0x33, 0x66, 0xCC ) } }
   else
      oBrw:bClrStd := { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED,  ),;
                          nClrBackBrw } }

      oBrw:bClrSel := { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrBackBrw ),;
                           RGB( 0x33, 0x66, 0xCC ) } }
   endif
   cClrBack := Eval( oBrw:bClrSelFocus )[ 2 ]
   oBrw:bClrSelFocus := { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrBackBrw ),;
                              cClrBack } }
   oBrw:CreateFromCode()
   oBrw:SetFocus()
   oBrw:bLDblClick := { || oBrw:EditSource() } // ADOEdit( oBrw:oRS, cTable ) }
   oBrw:bLClickHeaders = { || Eval( oBrw:bChange ) }
   oBrw:bKeyDown := { | nKey, nFlags | If( nKey == VK_DELETE,;
                     RSDelRecord( oBrw ),) }

   oWnd:oClient  := oBrw
   oWnd:oControl := oBrw

   if File( Alias() + ".brw" )
      oBrw:RestoreState( MemoRead( Alias() + ".brw" ) )
   endif

   DEFINE MSGBAR oMsgBar PROMPT FWString( "FileName" ) + ": " + If( ValType( cFileName ) == "C", cFileName, cTable ) OF oWnd 2010

   DEFINE MSGITEM oMsgRecNo OF oMsgBar ;
          PROMPT "RecNo: " + ;
          AllTrim( Str( If( oRS:AbsolutePosition == -3, oRS:RecordCount() + 1,;
                   oRS:AbsolutePosition ) ) ) + " / " + ;
          AllTrim( Str( oRS:RecordCount() ) ) ;
          SIZE 150 ;
          ACTION ( cAlias )->( DlgGoTo( oBrw ) )

   DEFINE MSGITEM oMsgTagName OF oMsgBar ;
          PROMPT FWString( "Ordered by" ) + ": " + If( Empty( oRs:Sort ),;
                 FWString( "natural order" ), oRs:Sort ) ;
          SIZE 150

   DEFINE MSGITEM oMsgRdd OF oMsgBar ;
          PROMPT "Rdd: " + "ADO" ;
          SIZE 120

   ACTIVATE WINDOW oWnd ;
      VALID (  MemoWrit( Alias() + ".brw", oBrw:SaveState() ),;
               ( cAlias )->( DbCloseArea() ), oBrw:cAlias := "", .T. )

return oBrw

//----------------------------------------------------------------------------//

static function RSAppendBlank( oRS )

   local n, aValues := {}

   if oRs:RecordCount() > 0
      oRS:MoveLast()

      for n = 1 to oRS:Fields:Count
         AAdd( aValues, oRS:Fields[ n - 1 ]:Value )
         if n == 1 .and. ValType( aValues[ 1 ] ) == "N"
            aValues[ 1 ]++
         else
            aValues[ n ] = uValBlank( aValues[ n ] )
            if ValType( aValues[ n ] ) == "D" .and. Empty( aValues[ n ] )
               aValues[ n ] = Date()
            endif
         endif
      next
   else
      aValues = Array( oRS:Fields:Count )
      for n = 1 to oRS:Fields:Count
         do case
            case oRS:Fields[ n - 1 ]:Type == 3
                 aValues[ n ] = 1

            case oRS:Fields[ n - 1 ]:Type == 202 .or. ;
                 oRS:Fields[ n - 1 ]:Type == 203
                 aValues[ n ] = Space( oRS:Fields[ n - 1 ]:DefinedSize )

            case oRS:Fields[ n - 1 ]:Type == 131 .or. oRS:Fields[ n - 1 ]:Type == 16 .or. ;
                 oRS:Fields[ n - 1 ]:Type == 2 .or. oRS:Fields[ n - 1 ]:Type == 11
                 aValues[ n ] = 0

            case oRS:Fields[ n - 1 ]:Type == 135 .or. oRS:Fields[ n - 1 ]:Type == 7
                 aValues[ n ] = Date()

            otherwise
                 MsgInfo( oRS:Fields[ n - 1 ]:Type )
         endcase
      next
   endif

   oRS:AddNew()

   if ! Empty( aValues )
      for n = 1 to oRS:Fields:Count
         TRY
            oRS:Fields[ n - 1 ]:Value = aValues[ n ]
         CATCH
         END
      next
   endif

   oRS:Update()

return nil

//----------------------------------------------------------------------------//

function Pack( oBrw, cFileName )

   if MsgYesNo( "Do you want to completely remove the deleted records ?" )
      TRY
         ( oBrw:cAlias )->( __dbPack() )
         oBrw:Refresh()
      CATCH
         MsgStop( FWString( "Please change shared mode in preferences and " + ;
                  "open the file again in exclusive mode" ) )
      END
   endif

return nil

//----------------------------------------------------------------------------//

function Zap( oBrw, cFileName )

   if MsgYesNo( FWString( "WARNING: This will remove all records from the DBF. Are you sure ?" ) )
      TRY
         ( oBrw:cAlias )->( __dbZap() )
         oBrw:Refresh()
      CATCH
         MsgStop( FWString( "Please change shared mode in preferences and " + ;
                  "open the file again in exclusive mode" ) )
      END
   endif

return nil

//----------------------------------------------------------------------------//

function GenLocalBlock( aCols, n )

   // non fixed length fields are translated into Memo so we can not see them
   // return { || If( Empty( Eval( aCols[ n ]:bEditValue ) ), "<memo>", "<Memo>" ) }

return { || Eval( aCols[ n ]:bEditValue ) }

//----------------------------------------------------------------------------//

function Console()

   local oWndCmds, oBar, oBrw, oMsgBar
   local aCmds := { "oFiveDBU:New()" }, aResults := { oWndMain }

   DEFINE WINDOW oWndCmds TITLE "Commands" MDICHILD OF oWndMain

   DEFINE BUTTONBAR oBar OF oWndCmds STYLEBAR SIZE 70, 70

   DEFINE BUTTON OF oBar PROMPT FWString( "Add" ) RESOURCE "add" ;
      ACTION ( AAdd( aCmds, Space( 200 ) ), AAdd( aResults, "" ), oBrw:Refresh(), oBrw:SetFocus() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Exit" ) RESOURCE "exit" ACTION oWndCmds:End() GROUP

   @ 0, 0 XBROWSE oBrw OF oWndCmds LINES ARRAY aCmds

   ADD TO oBrw HEADER 'Command' DATA aCmds[ oBrw:nArrayAt ] SIZE 150
   ADD TO oBrw HEADER 'Result'  DATA cValToChar( aResults[ oBrw:nArrayAt ] ) SIZE 150
   ADD TO oBrw HEADER 'Type'    DATA ValType( aResults[ oBrw:nArrayAt ] ) SIZE 150

   oBrw:CreateFromCode()
   oWndCmds:oClient  = oBrw
   oWndCmds:oControl = oBrw
   StyleBrowse( oBrw )

   DEFINE MSGBAR oMsgBar OF oWndCmds 2010

   ACTIVATE WINDOW oWndCmds

return nil

//----------------------------------------------------------------------------//

function Copy()

   if oWndMain:oWndActive != nil
      oWndMain:oWndActive:Copy()
   endif

return nil

//----------------------------------------------------------------------------//

function Recall( oBrw, cFileName )

   local nRecNo := RecNo()

   if ! FLock()
      MsgStop( cFileName + " " + FWString( "can not be locked" ),;
               FWString( "Recall records of" ) + " " + Alias() )
      return nil
   endif

   RECALL ALL
   DbUnLock()

   DbGoTo( nRecNo )
   oBrw:Refresh()
   oBrw:SetFocus()

return nil

//----------------------------------------------------------------------------//

function DeleteFor( cAlias, cFileName )

   local oDlg, cFor := Space( 80 ), cWhile := Space( 80 ),;
         cScope := FWString( "All" ), nRecord := 0, nDeleted := 0, oPgr

   DEFINE DIALOG oDlg TITLE FWString( "Delete records of" ) + " " + cAlias ;
      SIZE 600, 300

   @ 0.5,  2 SAY FWString( "For condition" ) OF oDlg SIZE 50, 8

   @ 1.5, 1 GET cFor OF oDlg SIZE 285, 11 ACTION FWExpBuilder( @cFor ) ;
      VALID If( ! Empty( cFor ), FWCheckExpression( cFor ), .T. )

   @ 2.5,  2 SAY FWString( "While condition" ) OF oDlg SIZE 50, 8

   @ 3.8, 1 GET cWhile OF oDlg SIZE 285, 11 ACTION FWExpBuilder( @cWhile ) ;
      VALID If( ! Empty( cWhile ), FWCheckExpression( cWhile ), .T. )

   @ 4.5,  2 SAY FWString( "Scope" ) OF oDlg SIZE 50, 8

   @ 5.5, 1 COMBOBOX cScope ;
      ITEMS { FWString( "All" ), FWString( "Next" ), FWString( "Record" ),;
              FWString( "Rest" ) } ;
      OF oDlg SIZE 70, 11

   @ 4.5,  16 SAY FWString( "Number of records" ) OF oDlg SIZE 60, 8

   @ 5.94, 11 GET nRecord OF oDlg SIZE 70, 11

   @ 4.5,  29 SAY FWString( "Total deleted" ) OF oDlg SIZE 60, 8

   @ 5.94, 21 GET nDeleted OF oDlg SIZE 70, 11 UPDATE

   @ 6.8, 1.4 PROGRESS oPgr OF oDlg SIZE 285, 11

   @ 7, 15 BUTTON FWString( "&Ok" ) OF oDlg SIZE 45, 13 ;
      ACTION DeleteGroup( oDlg, oPgr, cFor, cWhile, cScope, nRecord, RecNo(),;
                          @nDeleted, cFileName )

   @ 7, 26 BUTTON FWString( "&Cancel" ) OF oDlg SIZE 45, 13 ACTION oDlg:End()

   ACTIVATE DIALOG oDlg CENTERED

return nil

//----------------------------------------------------------------------------//

function DeleteGroup( oDlg, oPgr, cFor, cWhile, cScope, nReg, nActual,;
                      nDeleted, cFileName )

   local nRecNo := RecNo()

   if ! FLock()
      MsgStop( cFileName + " " + FWString( "can not be locked" ),;
               FWString( "Delete records of" ) + " " + Alias() )
      return nil
   endif

   nDeleted = 0

   ASend( oDlg:aControls, "disable" )

   oDlg:bValid := { || .F. }

   DbEval( { || nActual++,;
                If( ! Deleted(), nDeleted++,),;
               dbDelete() },;
               { || oPgr:SetPos( nActual ), Eval( FWGENBLOCK( cFor ) ) },;
               If( ! Empty( cWhile ), FWGENBLOCK( cWhile ), nil ),;
               If( cScope == "Next", nReg, nil ),;
               If( cScope == "Record", nReg, nil ),;
               ( cScope == FWString( "Rest" ) ) )

   DbUnlock()

   Asend( oDlg:aControls, "enable")

   oDlg:bValid = nil

   oPgr:SetPos( oPgr:nMax )
   oDlg:Update()

   MessageBeep()
   SysWait( 1 )
   nActual = 0
   DbGoTo( nRecNo )

return nil

//----------------------------------------------------------------------------//

function DelField( aFields, cFieldName, oGet, oBrw )

   if Len( aFields ) >= 1
      ADel( aFields, oBrw:nArrayAt, .T. )
      // ASize( aFields, Len( aFields ) - 1 )  bug fixed 08-Sept-2024
      oBrw:SetArray( aFields )
   endif

   oGet:VarPut( cFieldName := Space( 10 ) )
   oGet:SetPos( 0 )
   oGet:SetFocus()
   oBrw:GoBottom()

return nil

//----------------------------------------------------------------------------//

function DelRecord( oBrw, oMsgDeleted )

   if ! Deleted()
      if ! MsgYesNo( FWString( "Want to delete this record ?" ) )
         return nil
      endif
      DbRLock()
      DbDelete()
      DbUnlock()
      oMsgDeleted:SetText( FWString( "DELETED" ) )
      oMsgDeleted:SetBitmap( "deleted" )
   else
      DbRLock()
      DbRecall()
      DbUnlock()
      oMsgDeleted:SetText( FWString( "NON DELETED" ) )
      oMsgDeleted:SetBitmap( "nondeleted" )
      MsgInfo( "UnDeleted record" )
   endif

   if oBrw != nil
      if ! Empty( oBrw:bChange )
         Eval( oBrw:bChange )
         oBrw:Refresh()
      endif
   endif

return nil

//----------------------------------------------------------------------------//

static function RSDelRecord( oBrw )

   local n := oBrw:oRs:AbsolutePosition

   if ! MsgYesNo( FWString( "Want to delete this record ?" ) )
      return nil
   endif

   oBrw:oRs:Delete()
   oBrw:oRs:Update()
   if n > oBrw:oRs:RecordCount()  // Happens only when last record is deleted
      n--
   endif
   oBrw:oRs:AbsolutePosition = n  // in most cases this n is not changed. but this assignment is necessary.

   Eval( oBrw:bChange )
   oBrw:Refresh()

return nil

//----------------------------------------------------------------------------//

function ImportExport( cFileName, oBrwParent )

   local oDlg
   local oBrw
   local aFields := DbStruct()
   local n
   local oCol

   if Empty( aFields ) .and. ! Empty( oBrwParent:oRS )
      for n = 1 to oBrwParent:oRS:Fields:Count
         AAdd( aFields, { oBrwParent:oRS:Fields[ n - 1 ]:Name,;
                          oBrwParent:oRS:Fields[ n - 1 ]:Type,;
                          FWAdoFieldSize( oBrwParent:oRS, n ),;
                          FWAdoFieldDec( oBrwParent:oRS, n ), .T. } )
      next
   else
      for n = 1 to Len( aFields )
         AAdd( aFields[ n ], .T. )
      next
   endif

   DEFINE DIALOG oDlg TITLE FWString( "Import/Export for " ) + Alias() SIZE 400, 600

   @ 3.0, 1 XBROWSE oBrw ARRAY aFields AUTOCOLS NOBORDER STYLE FLAT ;
      HEADERS "", FWString( "Name" ), FWString( "Type" ), FWString( "Len" ),;
              FWString( "Dec" ) ;
      COLUMNS 5, 1, 2, 3, 4 ;
      COLSIZES 40, 120, 55, 40, 40 ;
      SIZE 183, 185 OF oDlg

   StyleBrowse( oBrw )
   WITH OBJECT oBrw:oCol( 1 )
      :nWidth      := 40
      :nEditType   := EDIT_GET
      :SetCheck()
      :nHeadBmpNo  := 1
   END
   oCol := oBrw:oCol( 1 )
   oCol:bLClickHeader := <||
         oCol:nHeadBmpNo   := If( oCol:nHeadBmpNo == 1, 2, 1 )
         AEval( oBrw:aArrayData, { |a| a[ 5 ] := ( oCol:nHeadBmpNo == 1 ) } )
         return nil
         >

   oBrw:lHScroll   := .F.
   if lPijama
      oBrw:bClrStd := { || If( oBrw:KeyNo() % 2 == 0, ;
                            { CLR_BLACK, RGB( 198, 255, 198 ) }, ;
                            { CLR_BLACK, RGB( 232, 255, 232 ) } ) }
      oBrw:bClrSel := { || { CLR_WHITE, RGB( 0x33, 0x66, 0xCC ) } }
   else
      oBrw:bClrStd := { || { nClrTxtBrw, nClrBackBrw } }
      oBrw:bClrSel := { || { nClrBackBrw, RGB( 0x33, 0x66, 0xCC ) } }
   endif
   oBrw:bLClicked  := { | r,c,f,o | if( o:MouseColPos( c ) == 1, ;
                            ( aFields[ oBrw:nArrayAt, 5 ] := !aFields[ oBrw:nArrayAt, 5 ], ;
                              o:RefreshCurrent() ), nil ) }

   BrwRecSel( oBrw, "KEYNO" )
   oBrw:CreateFromCode()

   ACTIVATE DIALOG oDlg CENTERED ;
      ON INIT ( BuildImpExpBar( oDlg, oBrw, cFileName, oBrwParent ), oDlg:Resize(), oBrw:SetFocus() )


return nil

//----------------------------------------------------------------------------//

static function BrwRecSel( oBrwI, cHead )

   local nSel   := 0
   WITH OBJECT oBrwI
      if "REC" $ Upper( cHead )
         :bRecSelHeader    := { || "RecNo" }
         :bRecSelData      := { |o| o:BookMark }
      else
         :bRecSelHeader    := { || "SlNo" }
         :bRecSelData      := { |o| o:KeyNo }
      endif
      :nRecSelWidth     := Replicate( '9', Len( cValToChar( Eval( oBrwI:bKeyCount, oBrwI ) ) ) + 1 )
   END

return nil

//----------------------------------------------------------------------------//

function BuildImpExpBar( oDlg, oBrw, cFileName, oBrwParent )

   local oBar

   DEFINE BUTTONBAR oBar OF oDlg STYLEBAR SIZE 70, 70

   DEFINE BUTTON OF oBar PROMPT "Import" RESOURCE "code" //;
   //   ACTION ( TxtStruct( oBrwParent ), oBrw:SetFocus() )

   DEFINE BUTTON OF oBar PROMPT "Export" RESOURCE "edit" //;
   //   ACTION ( New( Alias(), cFileName ) ) //, oDlg:End() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Exit" ) RESOURCE "exit" ;
      ACTION oDlg:End() GROUP BTNRIGHT

return nil

//----------------------------------------------------------------------------//

function DlgGoTo( oBrw )

   local oDlg, nRecNo := Int( oBrw:BookMark ) //RecNo()

   DEFINE DIALOG oDlg TITLE Alias() + FWString( " go to record" ) SIZE 400, 200
/*
   @ 1.2, 1.5 SAY FWString( "Current record" ) + ": " + AllTrim( Str( RecNo() ) ) + ;
      " / " + AllTrim( Str( OrdKeyCount() ) ) OF oDlg
*/
   @ 1.2, 1.5 SAY FWString( "Current record" ) + ": " + AllTrim( Str( nRecNo ) ) + ;
      " / " + AllTrim( Str( oBrw:nLen ) ) OF oDlg


   @ 2.4, 1.2 GET nRecNo OF oDlg SIZE 180, 11

   @ 4, 7 BUTTON FWString( "&Ok" ) OF oDlg SIZE 45, 13 ;
      ACTION ( oBrw:BookMark := nRecNo * 1.0, oBrw:Refresh(), oDlg:End() )

   @ 4, 18 BUTTON FWString( "&Cancel" ) OF oDlg SIZE 45, 13 ACTION oDlg:End()

   ACTIVATE DIALOG oDlg CENTERED

return nil

//----------------------------------------------------------------------------//
/*
function ADOEdit( oRS, cTableName )

   local oWnd, aRecord, oBar, oBrw, oMsgBar
   local oBtnSave, nRecNo := oRS:BookMark
   local oMsgDeleted

   aRecord = RSLoadRecord( oRS )

   DEFINE WINDOW oWnd TITLE FWString( "Edit" ) + " " + cTableName MDICHILD
   //oWnd:SetSize( 1150, 565 )
   oWndMain:oBar:AEvalWhen()

   DEFINE BUTTONBAR oBar OF oWnd STYLEBAR SIZE 70, 70

   DEFINE BUTTON oBtnSave OF oBar PROMPT FWString( "Save" ) RESOURCE "save" ;
      ACTION ( FWAdoSaveRecord( oRS, aRecord, nRecNo ), oBtnSave:Disable(), oBrw:SetFocus() )

   oBtnSave:Disable()

   DEFINE BUTTON OF oBar PROMPT FWString( "Prev" ) RESOURCE "prev" ;
      ACTION ( If( oRS:AbsolutePosition > 1,;
               ( oRS:MovePrevious(),;
                 nRecNo := oRs:BookMark,;
                 oBrw:SetArray( RSLoadRecord( oRS ) ),;
                 oBrw:Refresh(), Eval( oBrw:bChange ) ),), oBrw:SetFocus() ) GROUP

   DEFINE BUTTON OF oBar PROMPT FWString( "Next" ) RESOURCE "next" ;
      ACTION ( If( oRS:AbsolutePosition < oRS:RecordCount,;
               ( oRS:MoveNext(),;
                 nRecNo := oRs:BookMark,;
                 oBrw:SetArray( RSLoadRecord( oRS ) ),;
                 oBrw:Refresh(), Eval( oBrw:bChange ) ),), oBrw:SetFocus() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Report" ) RESOURCE "report" ;
      ACTION oBrw:Report() GROUP

   DEFINE BUTTON OF oBar PROMPT FWString( "View" ) RESOURCE "view" ;
      ACTION View( oBrw:aRow[ 2 ], oWnd )

   DEFINE BUTTON OF oBar PROMPT FWString( "Exit" ) RESOURCE "exit" ;
      ACTION oWnd:End() GROUP

   @ 0, 0 XBROWSE oBrw OF oWnd ARRAY aRecord AUTOCOLS LINES ;
      HEADERS FWString( "FieldName" ), FWString( "Value" ) COLSIZES 150, 400 FASTEDIT ;
      ON CHANGE ( SetEditType( oRs, oBrw, oBtnSave ), oBrw:DrawLine( .T. ),;
                  oMsgBar:cMsgDef := " RecNo: " + AllTrim( Str( oRS:AbsolutePosition ) ) + ;
                                 "/" + AllTrim( Str( oRS:RecordCount ) ),;
                  oMsgBar:Refresh() )

   oBrw:nEditTypes = EDIT_GET
   oBrw:aCols[ 1 ]:nEditType = 0 // Don't allow to edit first column
   oBrw:aCols[ 2 ]:bOnChange = { || oBtnSave:Enable() }
   oBrw:aCols[ 2 ]:lWillShowABtn = .T.
   StyleBrowse( oBrw )
   if lPijama
      oBrw:bClrStd := { || If( oBrw:KeyNo() % 2 == 0, ;
                            { CLR_BLACK, RGB( 198, 255, 198 ) }, ;
                            { CLR_BLACK, RGB( 232, 255, 232 ) } ) }
      oBrw:bClrSel := { || { CLR_WHITE, RGB( 0x33, 0x66, 0xCC ) } }
   else
      oBrw:bClrStd := { || { nClrTxtBrw, nClrBackBrw } }
      oBrw:bClrSel := { || { nClrBackBrw, RGB( 0x33, 0x66, 0xCC ) } }
   endif
   oBrw:CreateFromCode()
   oBrw:SetFocus()

   oWnd:oClient = oBrw

   DEFINE MSGBAR oMsgBar ;
      PROMPT " RecNo: " + AllTrim( Str( oRS:AbsolutePosition ) ) + "/" + ;
      AllTrim( Str( oRS:RecordCount ) ) OF oWnd 2010

   ACTIVATE WINDOW oWnd ;
      VALID If( oBtnSave:IsEnabled(),;
         MsgYesNo( "Do you want to exit without saving your changes ? " ), .T. )

return nil
*/
//----------------------------------------------------------------------------//

#ifndef __XHARBOUR__
static function RSLoadRecord( oRS )

   local aRecord := {}, n

   for n = 1 to oRS:Fields:Count
      AAdd( aRecord, { oRS:Fields[ n - 1 ]:Name, oRS:Fields[ n - 1 ]:Value } )
      If ValType( ATail( aRecord )[ 2 ] ) == "C"
         ATail( aRecord )[ 2 ] = PadR( ATail( aRecord )[ 2 ], Min( oRS:Fields[ n - 1 ]:DefinedSize, 255 ) )
       endif
   next

return aRecord
#endif

//----------------------------------------------------------------------------//

function AdoIndexes( oBrwPrev, oRS, cTableName )

   local aIndexes := FW_AdoIndexes( oCon,;
                          SubStr( cTablename, 2, Len( cTableName ) - 2 ) )
   local oWnd, oBar, oBrw, oMsgBar

   DEFINE WINDOW oWnd TITLE FWString( "Indexes of " ) + ctableName MDICHILD
   //oWnd:SetSize( 1150, 565 )
   oWndMain:oBar:AEvalWhen()

   DEFINE BUTTONBAR oBar OF oWnd STYLEBAR SIZE 70, 70

   DEFINE BUTTON OF oBar PROMPT FWString( "Add" ) RESOURCE "add" ;
      ACTION ( FW_AdoCreateIndex( oCon, cTableName, "IndexName",;
               { oRS:Fields[ 0 ]:Name }, .F. ),;
               oBrw:SetArray( FW_AdoIndexes( oCon,;
               SubStr( cTablename, 2, Len( cTableName ) - 2 ) ) ),;
               oBrw:Refresh(), oBrw:SetFocus() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Edit" ) RESOURCE "edit" ;
      ACTION ( oBrw:Edit() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Del" ) RESOURCE "del" ;
      ACTION If( Len( aIndexes ) > 0 .and. ;
                 MsgYesNo( FWString( "Want to delete this index key ?" ),;
                           FWString( "Please select" ) ),;
                 ( FW_AdoDropIndex( oCon, cTableName, "[" + aIndexes[ 1 ][ 1 ] + "]" ),;
                   oBrw:SetArray( FW_AdoIndexes( oCon,;
                   SubStr( cTablename, 2, Len( cTableName ) - 2 ) ) ),;
                   oBrw:Refresh(), oBrw:SetFocus() ),)

   DEFINE BUTTON OF oBar PROMPT FWString( "Report" ) RESOURCE "report" ;
      ACTION oBrw:Report() GROUP

   DEFINE BUTTON OF oBar PROMPT FWString( "Exit" ) RESOURCE "exit" ;
      ACTION oWnd:End() GROUP

   @ 0, 0 XBROWSE oBrw OF oWnd ARRAY aIndexes LINES ;
      HEADERS "Order", "IndexName", "Column Name", "IsUnique" ;
      COLUMNS { || oBrw:nArrayAt },;
              { || aIndexes[ oBrw:nArrayAt ][ 1 ] },;
              { || aIndexes[ oBrw:nArrayAt ][ 2 ] },;
              { || aIndexes[ oBrw:nArrayAt ][ 3 ] } ;
      COLSIZES 55, 317, 317, 317

   StyleBrowse( oBrw )
   if lPijama
      oBrw:bClrStd := { || If( oBrw:KeyNo() % 2 == 0, ;
                            { CLR_BLACK, RGB( 198, 255, 198 ) }, ;
                            { CLR_BLACK, RGB( 232, 255, 232 ) } ) }
      oBrw:bClrSel := { || { CLR_WHITE, RGB( 0x33, 0x66, 0xCC ) } }
   else
      oBrw:bClrStd := { || { nClrTxtBrw, nClrBackBrw } }
      oBrw:bClrSel := { || { CLR_WHITE, RGB( 0x33, 0x66, 0xCC ) } }
   endif
   oBrw:CreateFromCode()

   oBrw:SetFocus()

   oWnd:oClient = oBrw

   DEFINE MSGBAR oMsgBar 2010

   ACTIVATE WINDOW oWnd

return nil

//----------------------------------------------------------------------------//
/*
function Edit()

   local oWnd, aRecord := ( Alias() )->( LoadRecord() ), oBar, oBrw, oMsgBar
   local cAlias := Alias(), oBtnSave, nRecNo := ( Alias() )->( RecNo() )
   local oMsgDeleted, n

   DEFINE WINDOW oWnd TITLE FWString( "Edit" ) + " " + Alias() MDICHILD
   oWnd:SetSize( 620, 544 )
   oWndMain:oBar:AEvalWhen()

   DEFINE BUTTONBAR oBar OF oWnd STYLEBAR SIZE 70, 70

   DEFINE BUTTON oBtnSave OF oBar PROMPT FWString( "Save" ) RESOURCE "save" ;
      ACTION ( ( cAlias )->( SaveRecord( aRecord, nRecNo ) ), oBtnSave:Disable() )

   oBtnSave:Disable()

   DEFINE BUTTON OF oBar PROMPT FWString( "Prev" ) RESOURCE "prev" ;
      ACTION ( cAlias )->( GoPrevRecord( oBrw, aRecord, oMsgDeleted ) ) GROUP

   DEFINE BUTTON OF oBar PROMPT FWString( "Next" ) RESOURCE "next" ;
      ACTION ( cAlias )->( GoNextRecord( oBrw, aRecord, oMsgDeleted ) )

   DEFINE BUTTON OF oBar PROMPT FWString( "Report" ) RESOURCE "report" ;
      ACTION oBrw:Report() GROUP

   DEFINE BUTTON OF oBar PROMPT FWString( "View" ) RESOURCE "view" ;
      ACTION View( oBrw:aRow[ 2 ], oWnd )

   DEFINE BUTTON OF oBar PROMPT FWString( "Exit" ) RESOURCE "exit" ;
      ACTION oWnd:End() GROUP

   @ 0, 0 XBROWSE oBrw OF oWnd ARRAY aRecord AUTOCOLS LINES ;
      HEADERS "FieldName", FWString( "Value" ) COLSIZES 150, 400 FASTEDIT ;
      ON CHANGE ( ( cAlias )->( SetEditType( nil, oBrw, oBtnSave ) ), oBrw:DrawLine( .T. ),;
                  oMsgBar:cMsgDef := " RecNo: " + AllTrim( Str( ( cAlias )->( RecNo() ) ) ) + ;
                                 "/" + AllTrim( Str( ( cAlias )->( OrdKeyCount() ) ) ),;
                  oMsgBar:Refresh() )

   oBrw:nEditTypes = EDIT_GET
   oBrw:aCols[ 1 ]:nEditType = 0 // Don't allow to edit first column
   oBrw:aCols[ 2 ]:bOnChange = { || If( ( cAlias )->( FieldType( oBrw:nArrayAt ) != "M" ), oBtnSave:Enable(),) }
   oBrw:aCols[ 2 ]:lWillShowABtn = .T.
   //oBrw:nMarqueeStyle = MARQSTYLE_HIGHLROW
   StyleBrowse( oBrw )
   if lPijama
      oBrw:bClrStd := { || If( oBrw:KeyNo() % 2 == 0, ;
                            { CLR_BLACK, RGB( 198, 255, 198 ) }, ;
                            { CLR_BLACK, RGB( 232, 255, 232 ) } ) }
      oBrw:bClrSel := { || { CLR_WHITE, RGB( 0x33, 0x66, 0xCC ) } }
   else
      oBrw:bClrStd := { || { nClrTxtBrw, nClrBackBrw } }
      oBrw:bClrSel := { || { nClrBackBrw, RGB( 0x33, 0x66, 0xCC ) } }
   endif
   oBrw:CreateFromCode()

   oBrw:aCols[ 2 ]:bPaintText = { | oCol, hDC, cText, aCoors, aColors, lHighlight | ;
        DrawText( hDC, ( cAlias )->( If( FieldType( oBrw:nArrayAt ) == "M",;
        If( ! Empty( FieldGet( oBrw:nArrayAt ) ), "<Memo>", "<memo>" ), cText ) ), aCoors, DT_RIGHT ) }

   oBrw:bKeyDown = { | nKey, nFlags | If( nKey == VK_DELETE,;
                     ( cAlias )->( DelRecord( oBrw, oMsgDeleted ) ),) }

   oBrw:SetFocus()
   oWnd:oClient = oBrw

   DEFINE MSGBAR oMsgBar ;
      PROMPT " RecNo: " + AllTrim( Str( ( cAlias )->( RecNo() ) ) ) + "/" + ;
      AllTrim( Str( ( cAlias )->( OrdKeyCount() ) ) ) OF oWnd 2010

   DEFINE MSGITEM oMsgDeleted OF oMsgBar ;
          PROMPT If( ( cAlias )->( Deleted() ), FWString( "DELETED" ), FWString( "NON DELETED" ) ) ;
          SIZE 130 ;
          BITMAPS If( ( cAlias )->( Deleted() ), "deleted", "nondeleted" ) ;
          ACTION ( cAlias )->( DelRecord( oBrw, oMsgDeleted ) )

   ACTIVATE WINDOW oWnd ;
      VALID If( oBtnSave:IsEnabled(),;
         MsgYesNo( "Do you want to exit without saving your changes ? " ), .T. )

return nil
*/
//----------------------------------------------------------------------------//

function Filter( oBrw )

   local oDlg, cFilter, oGet, oBtn
   local lDbf  := !Empty( oBrw:cAlias )
   local oRs   := If( oBrw:oRs == nil, oBrw:oDbf, oBrw:oRs )

   cFilter  := PadR( If( lDBF, ( oBrw:cAlias )->( DbFilter() ), oRs:Filter ), 200 )

   DEFINE DIALOG oDlg TITLE FWString( "Filter" ) + " " + Alias() SIZE 600, 200

   @ 1, 1.5 SAY FWString( "Expression" ) + ":" OF oDlg SIZE 80, 10

   @ 2, 1 GET oGet VAR cFilter OF oDlg SIZE 285, 11 ;
      ACTION FWExpBuilder( @cFilter, oRs, ;
                           if( lPijama, nil, nClrTxtBrw ), ;
                           if( lPijama, nil, nClrBackBrw ) ) ;
      VALID If( ! Empty( cFilter ) .and. lDbf, ( oBrw:cAlias )->( FWCheckExpression( cFilter ) ), .T. )

   @ 4, 15 BUTTON oBtn PROMPT FWString( "&Ok" ) OF oDlg SIZE 45, 13
   oBtn:bAction   := <||
      cFilter     := AllTrim( cFilter )
      if Empty( cFilter )
         if lDbf
            ( oBrw:cAlias )->( DbClearFilter() )
         else
            oRs:Filter  := ""
         endif
      else
         if lDbf
            ( oBrw:cAlias )->( DbSetFilter( FWGENBLOCK( cFilter ), cFilter ) )
         else
            TRY
               oRs:Filter := cFilter
            CATCH
               ? "Invalid filter"
            END
         endif
      endif
      oBrw:GoTop()
      oBrw:Refresh()
      oDlg:End()
      return nil
      >

   @ 4, 26 BUTTON FWString( "&Cancel" ) OF oDlg SIZE 45, 13 ACTION oDlg:End() CANCEL

   ACTIVATE DIALOG oDlg CENTERED

return nil

//----------------------------------------------------------------------------//

function GetFields()

   local aStruct := DbStruct(), n, aFields := {}

   for n = 1 to Len( aStruct )
      AAdd( aFields, aStruct[ n ][ 1 ] )
   next

return aFields

//----------------------------------------------------------------------------//

function GlobalReplacement( oBrw )

   local oDlg, oCbxFieldName
   local cFieldName, cExpression := Space( 200 )
   local aFieldNames := {}
   local oRadScope, nScope := 1
   local cStep := "  1", lStep := .F.
   local cForCondition := Space( 200 )
   local cWhileCondition := Space( 200 )
   local oProgress, oGetTotal, cTotal := "0", nRecNo, n := 1
   local oBtnOk

   AEval( ( Alias() )->( DbStruct() ),;
      { | aField | AAdd( aFieldNames, aField[ 1 ] ) } )

   cFieldName = aFieldNames[ 1 ]

   DEFINE DIALOG oDlg TITLE FWString( "Global replacement in" ) + " " + Alias() ;
      SIZE 350, 400

   @ 0.6, 1.5 SAY FWString( "Field" ) + ":" OF oDlg SIZE 80, 10

   @ 1.6, 1 COMBOBOX oCbxFieldName VAR cFieldName ITEMS aFieldNames ;
      OF oDlg SIZE 40, 11

   @ 0.6, 11 SAY FWString( "Expression" ) + ":" OF oDlg SIZE 80, 10

   @ 1.7, 8 GET cExpression OF oDlg SIZE 104, 11 ;
      ACTION FWExpBuilder( @cExpression, oBrw:oRs ) ;
      VALID If( ! Empty( cExpression ),;
            FWCheckExpression( cExpression, FieldGet( AScan( aFieldNames, cFieldName ) ) ), .F. )

   @ 2.6, 1.5 SAY FWString( "Scope" ) + ":" OF oDlg SIZE 80, 10

   @ 3.35, 0.8 RADIO oRadScope VAR nScope ITEMS "All", "Rest" OF oDlg SIZE 18, 15

   @ 2.6, 11 SAY FWString( "Options" ) + ":" OF oDlg SIZE 80, 10

   @ 4.1, 9.2 CHECKBOX lStep PROMPT "Incremental" OF oDlg SIZE 60, 10

   @ 3.55, 17.6 SAY FWString( "Step" ) + ":" OF oDlg SIZE 60, 10

   @ 4.1, 17 GET cStep OF oDlg SIZE 32, 11 WHEN lStep RIGHT

   @ 4.6, 1.5 SAY FWString( "FOR condition" ) + ":" OF oDlg SIZE 80, 10

   @ 6.4, 1 GET cForCondition OF oDlg SIZE 159, 11 ;
      ACTION FWExpBuilder( @cForCondition, oBrw:oRs ) ;
      VALID If( ! Empty( cForCondition ), FWCheckExpression( cForCondition ), .T. )

   @ 6.6, 1.5 SAY FWString( "WHILE condition" ) + ":" OF oDlg SIZE 80, 10

   @ 8.6, 1 GET cWhileCondition OF oDlg SIZE 159, 11 ;
      ACTION FWExpBuilder( @cWhileCondition, oBrw:oRs ) ;
      VALID If( ! Empty( cWhileCondition ), FWCheckExpression( cWhileCondition ), .T. )

   @ 8.6, 1.5 SAY FWString( "Progress" ) + ":" OF oDlg SIZE 80, 10

   @ 9.5, 1.4 PROGRESS oProgress SIZE 105, 12

   @ 9.5, 20 SAY FWString( "Total" ) + ":" OF oDlg SIZE 80, 10

   @ 11, 17 GET oGetTotal VAR cTotal OF oDlg SIZE 32, 11 WHEN .F. RIGHT

   @ 9.5, 5.9 BUTTON oBtnOk PROMPT FWString( "&Ok" ) OF oDlg SIZE 45, 13 ;
      ACTION ( oBtnOk:Disable(), nRecNo := RecNo(),;
               DbEval( { || oGetTotal:SetText( AllTrim( Str( n ) ) ),;
                  oProgress:SetPos( n++ / 5 ), If( DbRLock(),;
                  ( FieldPut( AScan( aFieldNames, cFieldName ),;
                  &( AllTrim( cExpression ) ) ), DbUnLock() ),) },;
                  If( ! Empty( cForCondition ), &( "{ ||" + cForCondition + "}" ),),;
                  If( ! Empty( cWhileCondition ), &( "{ ||" + cWhileCondition + "}" ),),;
                  If( lStep, Val( cStep ),) ),;
                  DbGoTo( nRecNo ), oBrw:Refresh(), MsgInfo( FWString( "done" ) ), oDlg:End() )

   /*
   #command REPLACE [<f1> WITH <x1> [, <fN> WITH <xN>]] ;
                 [FOR <for>] [WHILE <while>] [NEXT <next>] ;
                 [RECORD <rec>] [<rest:REST>] [ALL] => ;
         dbEval( {|| _FIELD-><f1> := <x1>[, _FIELD-><fN> := <xN>] }, ;
                 <{for}>, <{while}>, <next>, <rec>, <.rest.> )
   */

   @ 9.5, 15.9 BUTTON FWString( "&Cancel" ) OF oDlg SIZE 45, 13 ;
      ACTION oDlg:End()

   ACTIVATE DIALOG oDlg CENTERED ;
      ON INIT ( oRadScope:aItems[ 2 ]:Move( oRadScope:aItems[ 1 ]:nTop,;
                oRadScope:aItems[ 1 ]:nRight + 40 ),;
                oProgress:SetRange( 0, 100 ), oProgress:SetStep( 1 ), .T. )

return nil

//----------------------------------------------------------------------------//

function GoPrevRecord( oBrw, aRecord, oMsgDeleted )

   DbSkip( -1 )

   oBrw:SetArray( aRecord := LoadRecord() )
   oBrw:SetFocus()
   Eval( oBrw:bChange )

   oMsgDeleted:SetText( FWString( If( Deleted(), "DELETED", "NON DELETED" ) ) )
   oMsgDeleted:SetBitmap( If( Deleted(), "deleted", "nondeleted" ) )

return nil

//----------------------------------------------------------------------------//

function GoNextRecord( oBrw, aRecord, oMsgDeleted )

   DbSkip( 1 )

   If Eof()
      DbSkip( -1 )
   else
      oBrw:SetArray( aRecord := LoadRecord() )
      oBrw:SetFocus()
      Eval( oBrw:bChange )
   endif

   oMsgDeleted:SetText( FWString( If( Deleted(), "DELETED", "NON DELETED" ) ) )
   oMsgDeleted:SetBitmap( If( Deleted(), "deleted", "nondeleted" ) )

return nil

//----------------------------------------------------------------------------//

function SetEditType( oRs, oBrw, oBtnSave )

   local cType, cAlias

   if Empty( oRs )
      cType  = FieldType( oBrw:nArrayAt )
      cAlias = Alias()
   else
      cType = FWAdoFieldType( oRs, oBrw:nArrayAt )
   endif

   do case
      case cType == "M"
           oBrw:aCols[ 2 ]:nEditType = EDIT_BUTTON
           if Empty( oRs )
              oBrw:aCols[ 2 ]:bEditBlock = ;
                 { || If( ( cAlias )->( EditMemo( oBrw ) ), oBtnSave:Enable(),) }
           else
              oBrw:aCols[ 2 ]:bEditBlock = ;
                 { || If( EditMemo( oBrw ), oBtnSave:Enable(),) }
           endif

      case cType == "D"
           oBrw:aCols[ 2 ]:nEditType = EDIT_BUTTON
           oBrw:aCols[ 2 ]:bEditBlock = { || If( ! Empty( oBrw:aRow[ 2 ] ) .and. ;
                                                 ! AllTrim( DtoC( oBrw:aRow[ 2 ] ) ) == "/  /",;
                                                 MsgDate( oBrw:aRow[ 2 ] ),;
                                                 MsgDate( Date() ) ) }

      case cType == "L"
           oBrw:aCols[ 2 ]:nEditType = EDIT_LISTBOX
           oBrw:aCols[ 2 ]:aEditListTxt   = { ".T.", ".F." }
           oBrw:aCols[ 2 ]:aEditListBound = { .T., .F. }

      otherwise
           oBrw:aCols[ 2 ]:nEditType = EDIT_GET
   endcase

return nil

//----------------------------------------------------------------------------//

function EditMemo( oBrw )

   local cTemp  := oBrw:aRow[ 2 ]
   local lResult := .F.

   if lResult := MemoEdit( @cTemp, oBrw:aRow[ 1 ] )
      oBrw:aRow[ 2 ] = cTemp
      oBrw:DrawLine()
   endif

return lResult

//----------------------------------------------------------------------------//

function IndexBuilder()

   local oDlg, cExp := Space( 80 ), cTo := Space( 80 ), cTag := Space( 20 )
   local cFor := Space( 80 ), cWhile := Space( 80 ), lUnique := .F.
   local lDescend := .F., lMemory := .F., cScope := FWString( "All" ), nRecNo
   local oPgr, nMeter := 0, nStep := 10, lTag := .T.
   local oBtnCreate, oBtnCancel

   DEFINE DIALOG oDlg TITLE FWString( "Index builder" ) SIZE 540, 380

   @ 0.4, 3 SAY FWString( "Expression" ) + ":" OF oDlg SIZE 40, 8

   @ 0.5, 6 GET cExp OF oDlg SIZE 200, 11 ACTION FWExpBuilder( @cExp ) ;
      VALID If( ! Empty( cExp ), FWCheckExpression( cExp ), ! Empty( cExp ) )

   @ 1.5, 5 SAY FWString( "To" ) + ":" OF oDlg SIZE 40, 8

   @ 1.7, 6 GET cTo OF oDlg SIZE 200, 11

   @ 2.6, 5 SAY FWString( "Tag" ) + ":" OF oDlg SIZE 40, 8

   @   3, 6 GET cTag OF oDlg SIZE 100, 11

   @ 3.7, 5 SAY FWString( "For" ) + ":" OF oDlg SIZE 40, 8

   @ 4.3, 6 GET cFor OF oDlg SIZE 200, 11 ACTION FWExpBuilder() ;
      VALID If( ! Empty( cFor ), FWCheckExpression( cFor ), .T. )

   @ 4.8, 4.5 SAY FWString( "While" ) + ":" OF oDlg SIZE 40, 8

   @ 5.6, 6 GET cWhile OF oDlg SIZE 200, 11 ACTION FWExpBuilder() ;
      VALID If( ! Empty( cWhile ), FWCheckExpression( cWhile ), .T. )

   @ 6.8, 6.9 CHECKBOX lUnique PROMPT FWString( "&Unique" ) OF oDlg SIZE 50, 8

   @ 6.8, 13.9 CHECKBOX lDescend PROMPT FWString( "&Descending" ) OF oDlg SIZE 80, 8

   @ 6.8, 23 CHECKBOX lMemory PROMPT FWString( "&Memory" ) OF oDlg SIZE 80, 8

   @ 6.8, 8 SAY FWString( "Scope" ) + ":" OF oDlg SIZE 40, 8

   @ 8.2, 6 COMBOBOX cScope ;
      ITEMS { FWString( "All" ), FWString( "Next" ), FWString( "Record" ), FWString( "Rest" ) } OF oDlg

   @ 6.8, 16.1 SAY FWString( "Record" ) + ":" OF oDlg SIZE 40, 8

   @ 8.8, 12.1 GET nRecNo OF oDlg SIZE 30, 11

   @ 8.5, 8 SAY FWString( "Progress" ) + ":" OF oDlg SIZE 40, 8

   @ 9.2, 8 PROGRESS oPgr OF oDlg SIZE 200, 10

   @ 9.2, 13 BUTTON oBtnCreate PROMPT FWString( "C&reate" ) OF oDlg SIZE 45, 13 ;
      ACTION ( oDlg:bValid := { || .F. },;
               oBtnCreate:Disable(),;
               oBtnCancel:Hide(),;
               oBtnCreate:SetText( "Please wait..." ),;
               oBtnCreate:nLeft += 70,;
               OrdCondSet( If( ! Empty( cFor ), cFor,),;
                           If( ! Empty( cFor ), FWGENBLOCK( cFor ),),;
                           If( cScope == "All", .T.,),;
                           If( ! Empty( cWhile ), FWGENBLOCK( cWhile ),),;
                           { || nMeter += nStep, oPgr:SetPos( nMeter ), SysRefresh() },;
                           nStep, Recno(),;
                           If( cScope == "Next", nRecNo,),;
                           If( cScope == "Record", nRecNo,),;
                           If( cScope == "Rest", .T.,),;
                           If( lDescend, .T.,),;
                           lTag,, .F., .F., .F., .T., .F., .F. ),;
               OrdCreate( OrdBagName(), If( lTag, cTag,), cExp, FWGENBLOCK( cExp ),;
                          If( lUnique, .T.,) ), oDlg:bValid := { || .T. }, oDlg:End() )

   @ 9.2, 24 BUTTON oBtnCancel PROMPT FWString( "&Cancel" ) OF oDlg SIZE 45, 13 ;
      ACTION oDlg:End() CANCEL

   ACTIVATE DIALOG oDlg CENTERED

return nil

//----------------------------------------------------------------------------//

function Indexes( oBrwParent, oBtnIndexes )

   local oWnd, oBar, oBrw, oMsgBar
   local cAlias := Alias(), aIndexes := Array( OrdCount() )
   local oBttEdit

   DEFINE WINDOW oWnd TITLE FWString( "Indexes of " ) + Alias() MDICHILD
   oWnd:SetSize( 1320, 570 )
   oWndMain:oBar:AEvalWhen()

   DEFINE BUTTONBAR oBar OF oWnd STYLEBAR SIZE 70, 70

   if RddName() == "DBFNTX"
      DEFINE BUTTON OF oBar PROMPT FWString( "Open" ) RESOURCE "open" ;
         ACTION ( DbSetIndex( cGetFile( FWString( "NTX file" ) + "| *.ntx|",;
                              FWString( "Please select a NTX file" ) ) ),;
                  oBrw:SetArray( Array( OrdCount() ) ), oBrw:Refresh(),;
                  Eval( oBrwParent:bChange ), oBtnIndexes:oPopup:End(),;
                  oBtnIndexes:oPopup := BuildIndexesDropMenu( cAlias, oBrwParent ), oBrw:SetFocus() )
   endif

   DEFINE BUTTON OF oBar PROMPT FWString( "Add" ) RESOURCE "add" ;
      ACTION ( IndexBuilder(), oBrw:SetArray( Array( OrdCount() ) ),;
               oBrw:Refresh(), oBrw:SetFocus() )

   DEFINE BUTTON oBttEdit OF oBar PROMPT FWString( "Edit" ) RESOURCE "edit" ;
      ACTION ( oBrw:Edit() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Del" ) RESOURCE "del" ;
      ACTION If( MsgYesNo( FWString( "Want to delete this tag ?" ),;
                           FWString( "Please select" ) ),;
                ( ( cAlias )->( OrdDestroy( oBrw:nArrayAt ) ),;
                oBrw:SetArray( Array( OrdCount() ) ), oBrw:Refresh() ),)

   DEFINE BUTTON OF oBar PROMPT "Rebuild" RESOURCE "index" GROUP //;
      // ACTION ( cAlias )->( Indexes( oBrw, oBtnIndexes, cAlias ) )

   DEFINE BUTTON OF oBar PROMPT FWString( "Report" ) RESOURCE "report" ;
      ACTION oBrw:Report() GROUP

   DEFINE BUTTON OF oBar PROMPT FWString( "Exit" ) RESOURCE "exit" ;
      ACTION oWnd:End() GROUP

   @ 0, 0 XBROWSE oBrw OF oWnd ARRAY aIndexes AUTOCOLS LINES ;
      HEADERS FWString( "Order" ), "Active", "Custom", FWString( "TagName" ), FWString( "Expression" ),;
              FWString( "For" ), FWString( "BagName" ), ; //FWString( "BagExt" ), ;
              "ScopeTop", "ScopeBottom" ;
      COLUMNS { || oBrw:nArrayAt },;
              { |x| ( cAlias )->( If( x == .T., OrdSetFocus( oBrw:nArrayAt ), IndexOrd() == oBrw:nArrayAt ) ) },;
              { || ( cAlias )->( dbOrderInfo( DBOI_CUSTOM, OrdName( oBrw:nArrayAt ), oBrw:nArrayAt ) ) },;
              { || ( cAlias )->( OrdName( oBrw:nArrayAt ) ) },;
              { || ( cAlias )->( OrdKey ( oBrw:nArrayAt ) ) },;
              { || If( ! Empty( ( cAlias )->( OrdFor ( oBrw:nArrayAt ) ) ), ;
                     ( cAlias )->( OrdFor( oBrw:nArrayAt ) ), Space( 60 ) ) },;
              { || ( cAlias )->( OrdBagName( oBrw:nArrayAt ) ) }, ;
              { |x| ( cAlias )->( SetGet_Scope( oBrw, 0, x ) ) }, ;
              { |x| ( cAlias )->( SetGet_Scope( oBrw, 1, x ) ) } ;
      COLSIZES 50, 50, 100, 400, 240, 100, 100, 100, 100

   WITH OBJECT oBrw:aCols[ 2 ]
      :SetCheck( nil, .t. )
      :bEditWhen  := { |o| o:Value == .f. }
      :bOnChange  := { || oBrwParent:GoTop(), oBrwParent:Refresh(), oBrw:Refresh() }
   END
   WITH OBJECT oBrw:aCols[ 3 ]
      :SetCheck( nil, .t. )
      :bEditWhen  := { |o| o:Value == .f. }
      :bOnChange  := { || oBrwParent:GoTop(), oBrwParent:Refresh(), oBrw:Refresh() }
   END
   oBrw:aCols[ 7 ]:nEditType := ;
   oBrw:aCols[ 8 ]:nEditType := EDIT_GET
   oBrw:aCols[ 7 ]:bOnChange := <|oCol|
      if ( cAlias )->( IndexOrd() ) == oBrw:nArrayAt
         oBrwParent:GoTop()
         oBrwParent:Refresh()
      endif
      return nil
      >
   oBrw:aCols[ 8 ]:bOnChange := <|oCol|
      if ( cAlias )->( IndexOrd() ) == oBrw:nArrayAt
         oBrwParent:GoBottom()
         oBrwParent:Refresh()
      endif
      return nil
      >


   oBrw:bLDblClick = { || oBrw:Edit() }

   StyleBrowse( oBrw )
   if lPijama
      oBrw:bClrStd := { || If( oBrw:KeyNo() % 2 == 0, ;
                            { CLR_BLACK, RGB( 198, 255, 198 ) }, ;
                            { CLR_BLACK, RGB( 232, 255, 232 ) } ) }
      oBrw:bClrSel := { || { CLR_WHITE, RGB( 0x33, 0x66, 0xCC ) } }
   else
      oBrw:bClrStd = { || { nClrTxtBrw, nClrBackBrw } }
      oBrw:bClrSel = { || { nClrBackBrw, RGB( 0x33, 0x66, 0xCC ) } }
   endif
   oBrw:CreateFromCode()

   oBrw:SetFocus()

   oWnd:oClient := oBrw

   DEFINE MSGBAR oMsgBar 2010

   ACTIVATE WINDOW oWnd

return nil

//----------------------------------------------------------------------------//

static function SetGet_Scope( oBrw, nTopBot, uNewVal )

   local uVal
   local nOrder   := oBrw:nArrayAt
   local cBagName := OrdBagName( nOrder )
   local cType    := DbOrderInfo( DBOI_KEYTYPE, cBagName, nOrder )

   if uNewVal  != nil
      if Empty( uNewVal )
         DbOrderInfo( If( nTopBot == 0, DBOI_SCOPETOPCLEAR, DBOI_SCOPEBOTTOMCLEAR ), cBagName, nOrder )
      else
         uNewVal  := uCharToVal( AllTrim( uNewVal ), cType )
         DbOrderInfo( If( nTopBot == 0, DBOI_SCOPETOP, DBOI_SCOPEBOTTOM ), cBagName, nOrder, uNewVal )
      endif
   endif

   uVal  := cValToChar( DbOrderInfo( If( nTopBot == 0, DBOI_SCOPETOP, DBOI_SCOPEBOTTOM ), cBagName, nOrder ) )
   uVal  := PadR( uVal, 50 )

return uVal

//----------------------------------------------------------------------------//

function LoadRecord()

   local aRecord := {}, n

   for n = 1 to FCount()
      AAdd( aRecord, { FieldName( n ), FieldGet( n ) } )
   next

return aRecord

//----------------------------------------------------------------------------//

function Search( oBrw )

   local oDlg, oCbx, cSearch := Space( 50 )
   local nRecNo := RecNo(), lInc := .T.

   DEFINE DIALOG oDlg TITLE FWString( "Search" ) + ": " + Alias() SIZE 400, 200

   @ 0.5, 1.5 SAY FWString( "Ordered by" ) + ": " + ;
      If( ! Empty( oBrw:oRS ), oBrw:oRS:Sort, OrdName() ) OF oDlg

   @ 1.2, 1.5 SAY FWString( "Key" ) + ": " + ;
                  If( ! Empty( oBrw:oRS ), oBrw:oRS:Sort, OrdKey() ) OF oDlg

   @ 2.4, 1.2 COMBOBOX oCbx VAR cSearch ITEMS aSearches OF oDlg SIZE 180, 150 ;
     STYLE CBS_DROPDOWN

   oCbx:oGet:bChange = { |n,f,oGet| If( lInc, oBrw:Seek( Trim( oGet:oGet:Buffer ) ), .t. ) }


   @ 3.7, 1.5 CHECKBOX lInc PROMPT FWString( "&Incremental" ) OF oDlg SIZE 80, 10

   @ 4, 7 BUTTON FWString( "&Ok" ) OF oDlg SIZE 45, 13 ;
      ACTION ( oBrw:Seek( Trim( cSearch ) ), oDlg:End() )

   @ 4, 18 BUTTON FWString( "&Cancel" ) OF oDlg SIZE 45, 13 ACTION oDlg:End()

   ACTIVATE DIALOG oDlg CENTERED

return nil

//----------------------------------------------------------------------------//

function Preferences()

   local oDlg

   DEFINE DIALOG oDlg TITLE FWString( "Preferences" ) SIZE 400, 300

   @ 0.6, 1.5 CHECKBOX lPijama PROMPT "Efect Pijama" OF oDlg SIZE 110, 11

   @ 1.4, 1.5 SAY FWString( "Default RDD" ) + ": " + cDefRdd OF oDlg

   @ 2.8, 1.2 COMBOBOX cDefRdd ITEMS { "DBFNTX", "DBFCDX", "RDDADS" } ;
      OF oDlg SIZE 180, 150

   @ 4.4, 1.5 CHECKBOX lShared PROMPT FWString( "Open in shared mode" ) OF oDlg ;
      SIZE 110, 11

   @ 5.0, 1.5 SAY FWString( "Language" ) OF oDlg

   DEFAULT nLanguage := FWSetLanguage()

   @ 6.4, 1.2 COMBOBOX nLanguage ;
      ITEMS { FWString( "English" ), FWString( "Spanish" ), FWString( "French" ),;
              FWString( "Portuguese" ), FWString( "German" ) , FWString( "Italian" ) } ;
      OF oDlg SIZE 180, 150 ;
      ON CHANGE ( FWSetLanguage( nLanguage ), oWndMain:SetMenu( BuildMenu() ),;
                  BuildMainBar() )

   @ 7, 7 BUTTON FWString( "&Ok" ) OF oDlg SIZE 45, 13 ;
      ACTION ( SavePreferences(), oDlg:End() )

   @ 7, 18 BUTTON FWString( "&Cancel" ) OF oDlg SIZE 45, 13 ACTION oDlg:End()

   ACTIVATE DIALOG oDlg CENTERED

return nil

//----------------------------------------------------------------------------//

function Processes()

   local oWnd, oBar, oBrw, cClrBack, oMsgBar, cAlias, oMsgDeleted

   if ! File( "process.dbf" )
      DbCreate( "process.dbf", { { "NAME", "C", 50, 0 },;
                                 { "CODE", "M", 0, 0 } }, "DBFCDX" )
   endif

   cAlias = cGetNewAlias( "processes" )

   USE process VIA "DBFCDX" NEW SHARED ALIAS ( cAlias )

   if OrdKeyCount() == 0
      APPEND BLANK
      if RLock()
         Field->Name := "Test"
         Field->Code := "function Test()" + CRLF + CRLF + ;
                        "   MsgInfo( 'Hello world!' )" + ;
                        CRLF + CRLF + "return nil"
         DbUnlock()
      endif
   endif

   DEFINE WINDOW oWnd TITLE FWString( "Processes" ) MDICHILD
   //oWnd:SetSize( 1150, 565 )
   DEFINE BUTTONBAR oBar OF oWnd STYLEBAR SIZE 70, 70

   DEFINE BUTTON OF oBar PROMPT FWString( "Add" ) RESOURCE "add" ;
      ACTION ( ( oBrw:cAlias )->( DbAppend() ), oBrw:Refresh(), oBrw:SetFocus() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Edit" ) RESOURCE "edit" ;
      ACTION oBrw:EditSource()  //( oBrw:cAlias )->( Edit() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Del" ) RESOURCE "del" ;
      ACTION ( obrw:cAlias )->( DelRecord( oBrw, oMsgDeleted ) )

   DEFINE BUTTON OF oBar PROMPT FWString( "Run" ) RESOURCE "run" ;
      ACTION Execute( ( cAlias )->Code ) GROUP

   DEFINE BUTTON OF oBar PROMPT FWString( "Exit" ) RESOURCE "exit" ;
      ACTION oWnd:End() GROUP

   @ 0, 0 XBROWSE oBrw OF oWnd LINES AUTOSORT ;
      AUTOCOLS ALIAS Alias()

   StyleBrowse( oBrw )
   if lPijama
      oBrw:bClrStd := { || If( oBrw:KeyNo() % 2 == 0, ;
                        { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrTxtBrw ),;
                              RGB( 198, 255, 198 ) }, ;
                        { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrTxtBrw ),;
                              RGB( 232, 255, 232 ) } ) }
      oBrw:bClrSel := { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrBackBrw ),;
                              RGB( 0x33, 0x66, 0xCC ) } }
   else
      oBrw:bClrStd := { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrTxtBrw ),;
                          nClrBackBrw } }
      oBrw:bClrSel := { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrBackBrw ),;
                           RGB( 0x33, 0x66, 0xCC ) } }
   endif
   cClrBack = Eval( oBrw:bClrSelFocus )[ 2 ]
   oBrw:bClrSelFocus = { || { If( ( oBrw:cAlias )->( Deleted() ), CLR_HRED, nClrBackBrw ),;
                              cClrBack } }
   oBrw:CreateFromCode()
   oBrw:SetFocus()
   oBrw:bLDblClick = { || oBrw:EditSource() } //( oBrw:cAlias )->( Edit() ) }
   oBrw:bLClickHeaders = { || Eval( oBrw:bChange ) }
   oBrw:bKeyDown = { | nKey, nFlags | If( nKey == VK_DELETE,;
                     ( cAlias )->( DelRecord( oBrw, oMsgDeleted ) ),) }

   oWnd:oClient = oBrw
   oWnd:oControl = oBrw

   if File( Alias() + ".brw" )
      oBrw:RestoreState( MemoRead( Alias() + ".brw" ) )
   endif

   DEFINE MSGBAR oMsgBar OF oWnd 2010

   DEFINE MSGITEM oMsgDeleted OF oMsgBar ;
          PROMPT FWString( If( ( cAlias )->( Deleted() ), "DELETED", "NON DELETED" ) ) ;
          SIZE 130 ;
          BITMAPS If( ( cAlias )->( Deleted() ), "deleted", "nondeleted" ) ;
          ACTION ( cAlias )->( DelRecord( oBrw, oMsgDeleted ) )

   ACTIVATE WINDOW oWnd ;
      VALID ( ( cAlias )->( DbCloseArea() ),;
              MemoWrit( Alias() + ".brw", oBrw:SaveState() ),;
              oBrw:cAlias := "", .T. )

return nil

//----------------------------------------------------------------------------//

static function RSQuery( oBrw )

   local oDlg
   local oCbx1, cFieldName1, oCbx2, cOperation1, cValue1, oCbx3, cCon1
   local oCbx4, cFieldName2, oCbx5, cOperation2, cValue2, oCbx6, cCon2
   local oCbx7, cFieldName3, oCbx8, cOperation3, cValue3, oCbx9, cCon3
   local oCbx10, cFieldName4, oCbx11, cOperation4, cValue4, oCbx12, cCon4
   local oCbx13, cFieldName5, oCbx14, cOperation5, cValue5, oCbx15, cCon5
   local oCbx16, cFieldName6, oCbx17, cOperation6, cValue6, oCbx18, cCon6
   local aFields := {}, n
   local nRow := 1.4

   if ! Empty( oBrw:oRS:Filter )
      oBrw:oRS:Filter = ""
      oBrw:Refresh()
      MsgInfo( FWString( "current query removed" ), FWString( "Information" ) )
      return nil
   endif

   cValue1 := cValue2 := cValue3 := cValue4 := cValue5 := cValue6 := Space( 50 )

   for n = 1 to oBrw:oRS:Fields:Count
      AAdd( aFields, oBrw:oRS:Fields[ n - 1 ]:Name )
   next

   DEFINE DIALOG oDlg TITLE FWString( "Query builder" ) SIZE 458, 400

   @ 0.5, 1.6 SAY FWString( "Field Name" ) OF oDlg SIZE 40, 8

   @ 0.5, 10 SAY FWString( "Operation" ) OF oDlg SIZE 40, 8

   @ 0.5, 18 SAY FWString( "Value" ) OF oDlg SIZE 40, 8

   @ nRow, 1 COMBOBOX oCbx1 VAR cFieldName1 ;
      ITEMS aFields OF oDlg SIZE 41, 11

   @ nRow,  7 COMBOBOX oCbx2 VAR cOperation1 ;
      ITEMS { FWString( "Equal" ), FWString( "Different" ), FWString( "Like" ) } ;
      OF oDlg SIZE 41, 11

   @ nRow + 0.1, 13 GET cValue1 OF oDlg SIZE 80, 11

   @ nRow, 23.8 COMBOBOX oCbx3 VAR cCon1 ;
      ITEMS { "", "AND", "OR" } OF oDlg SIZE 31, 11

   @ nRow * 2, 1 COMBOBOX oCbx4 VAR cFieldName2 ;
      ITEMS aFields OF oDlg SIZE 41, 11

   @ nRow * 2,  7 COMBOBOX oCbx5 VAR cOperation2 ;
      ITEMS { FWString( "Equal" ), FWString( "Different" ), FWString( "Like" ) } ;
      OF oDlg SIZE 41, 11

   @ ( nRow + 0.1 ) * 2, 13 GET cValue2 OF oDlg SIZE 80, 11

   @ nRow * 2, 23.8 COMBOBOX oCbx6 VAR cCon2 ;
      ITEMS { "", "AND", "OR" } OF oDlg SIZE 31, 11

   @ nRow * 3, 1 COMBOBOX oCbx7 VAR cFieldName3 ;
      ITEMS aFields OF oDlg SIZE 41, 11

   @ nRow * 3,  7 COMBOBOX oCbx8 VAR cOperation3 ;
      ITEMS { FWString( "Equal" ), FWString( "Different" ), FWString( "Like" ) } ;
      OF oDlg SIZE 41, 11

   @ ( nRow + 0.1 ) * 3, 13 GET cValue3 OF oDlg SIZE 80, 11

   @ nRow * 3, 23.8 COMBOBOX oCbx9 VAR cCon3 ;
      ITEMS { "", "AND", "OR" } OF oDlg SIZE 31, 11

   @ nRow * 4, 1 COMBOBOX oCbx10 VAR cFieldName4 ;
      ITEMS aFields OF oDlg SIZE 41, 11

   @ nRow * 4,  7 COMBOBOX oCbx11 VAR cOperation4 ;
      ITEMS { FWString( "Equal" ), FWString( "Different" ), FWString( "Like" ) } ;
      OF oDlg SIZE 41, 11

   @ ( nRow + 0.1 ) * 4, 13 GET cValue4 OF oDlg SIZE 80, 11

   @ nRow * 4, 23.8 COMBOBOX oCbx12 VAR cCon4 ;
      ITEMS { "", "AND", "OR" } OF oDlg SIZE 31, 11

   @ nRow * 5, 1 COMBOBOX oCbx13 VAR cFieldName5 ;
      ITEMS aFields OF oDlg SIZE 41, 11

   @ nRow * 5,  7 COMBOBOX oCbx14 VAR cOperation5 ;
      ITEMS { FWString( "Equal" ), FWString( "Different" ), FWString( "Like" ) } ;
      OF oDlg SIZE 41, 11

   @ ( nRow + 0.1 ) * 5, 13 GET cValue5 OF oDlg SIZE 80, 11

   @ nRow * 5, 23.8 COMBOBOX oCbx15 VAR cCon5 ;
      ITEMS { "", "AND", "OR" } OF oDlg SIZE 31, 11

   @ nRow * 6, 1 COMBOBOX oCbx16 VAR cFieldName6 ;
      ITEMS aFields OF oDlg SIZE 41, 11

   @ nRow * 6,  7 COMBOBOX oCbx17 VAR cOperation6 ;
      ITEMS { FWString( "Equal" ), FWString( "Different" ), FWString( "Like" ) } ;
      OF oDlg SIZE 41, 11

   @ ( nRow + 0.1 ) * 6, 13 GET cValue6 OF oDlg SIZE 80, 11

   @ nRow * 6, 23.8 COMBOBOX oCbx18 VAR cCon6 ;
      ITEMS { "", "AND", "OR" } OF oDlg SIZE 31, 11

   @ 10, 10 BUTTON FWString( "&Ok" ) OF oDlg SIZE 45, 13 ;
      ACTION ( ProcessQuery( oBrw, { cFieldName1, cOperation1, cValue1, cCon1,;
                                     cFieldName2, cOperation2, cValue2, cCon2,;
                                     cFieldName3, cOperation3, cValue3, cCon3,;
                                     cFieldName4, cOperation4, cValue4, cCon4,;
                                     cFieldName5, cOperation5, cValue5, cCon5,;
                                     cFieldName6, cOperation6, cValue6, cCon6 } ),;
               oDlg:End() )

   @ 10, 21 BUTTON FWString( "&Cancel" ) OF oDlg SIZE 45, 13 ACTION oDlg:End()

   ACTIVATE DIALOG oDlg CENTERED

return nil

//----------------------------------------------------------------------------//

function ProcessQuery( oBrw, aValues )

   local cQuery, n

   // ? oBrw:oRS:Fields[ 0 ]:Type

   aValues[ 2 ] = AScan( { FWString( "Equal" ), FWString( "Different" ),;
                           FWString( "Like" ) }, aValues[ 2 ] )
   if aValues[ 2 ] == 0
      MsgInfo( FWString( "empty query!" ), FWString( "Information" ) )
      return nil
   endif
   aValues[ 2 ] = { "=", "<>", "LIKE" }[ aValues[ 2 ] ]
   cQuery = aValues[ 1 ] + " " + aValues[ 2 ] + " " + ;
            If( oBrw:oRS:Fields[ 0 ]:Type == 3, "'", "" ) + ;
            AllTrim( aValues[ 3 ] ) + ;
            If( oBrw:oRS:Fields[ 0 ]:Type == 3, "'", "" )

   for n = 4 to Len( aValues ) step 4
      if ! Empty( aValues[ n ] )
         aValues[ n + 2 ] = AScan( { FWString( "Equal" ), FWString( "Different" ),;
                                     FWString( "Like" ) }, aValues[ n + 2 ] )
         aValues[ n + 2 ] = { "=", "<>", "LIKE" }[ aValues[ n + 2 ] ]
         cQuery += " " + aValues[ n ] + " " + aValues[ n + 1 ] + " " + ;
                   aValues[ n + 2 ] + " " + If( oBrw:oRS:Fields[ ( n / 4 ) - 1 ]:Type == 3, "'", "" ) + ;
                   AllTrim( aValues[ n + 3 ] ) + ;
                   If( oBrw:oRS:Fields[ ( n / 4 ) - 1 ]:Type == 3, "'", "" )
                   // ? oBrw:oRS:Fields[ ( n / 4 ) - 1 ]:Type
      endif
   next

   if ! Empty( cQuery )
      oBrw:oRS:Filter = cQuery
      oBrw:Refresh()
   else
      MsgInfo( FWString( "Empty query!" ), FWString( "Information" ) )
   endif

return nil

//----------------------------------------------------------------------------//

function Relations()

   local oWnd, oBar, oBrw, aRelations := Array( 8 ), oMsgBar
   local nTarget, cAlias := Alias()

   DEFINE WINDOW oWnd TITLE FWString( "Relations of" ) + " " + Alias() MDICHILD
   //oWnd:SetSize( 1150, 565 )
   oWndMain:oBar:AEvalWhen()

   DEFINE BUTTONBAR oBar OF oWnd STYLEBAR SIZE 70, 70

   DEFINE BUTTON OF oBar PROMPT FWString( "Add" ) RESOURCE "add" ;
      ACTION MsgInfo( "Add" )

   DEFINE BUTTON OF oBar PROMPT FWString( "Edit" ) RESOURCE "edit" ;
      ACTION ( MsgInfo( "Edit" ) )

   DEFINE BUTTON OF oBar PROMPT FWString( "Del" ) RESOURCE "del" ;
      ACTION MsgInfo( "del" )

   DEFINE BUTTON OF oBar PROMPT FWString( "Report" ) RESOURCE "report" ;
      ACTION oBrw:Report() GROUP

   DEFINE BUTTON OF oBar PROMPT FWString( "Exit" ) RESOURCE "exit" ;
      ACTION oWnd:End() GROUP

   @ 0, 0 XBROWSE oBrw OF oWnd ARRAY aRelations AUTOCOLS LINES ;
      HEADERS FWString( "Rel." ), FWString( "Expression" ),;
              FWString( "Child Alias" ), FWString( "Additive" ),;
              FWString( "Scoped" ) ;
      COLUMNS { || oBrw:nArrayAt },;
              { || ( cAlias )->( DbRelation( oBrw:nArrayAt ) ) },;
              { || If( ( cAlias )->( DbRSelect( oBrw:nArrayAt ) ) != 0,;
                   Alias( ( cAlias )->( DbRSelect( oBrw:nArrayAt ) ) ), "" ) },;
              { || ".F." }, { || ".F." } ;
      COLSIZES 40, 150, 400, 100, 100

   StyleBrowse( oBrw )
   if lPijama
      oBrw:bClrStd := { || If( oBrw:KeyNo() % 2 == 0, ;
                            { CLR_BLACK, RGB( 198, 255, 198 ) }, ;
                            { CLR_BLACK, RGB( 232, 255, 232 ) } ) }
      oBrw:bClrSel := { || { CLR_WHITE, RGB( 0x33, 0x66, 0xCC ) } }
   else
      oBrw:bClrStd = { || { nClrTxtBrw, nClrBackBrw } }
      oBrw:bClrSel = { || { nClrBackBrw, RGB( 0x33, 0x66, 0xCC ) } }
   endif
   oBrw:CreateFromCode()
   oBrw:SetFocus()

   oWnd:oClient = oBrw

   DEFINE MSGBAR oMsgBar 2010

   ACTIVATE WINDOW oWnd

return nil

//----------------------------------------------------------------------------//

function LoadPreferences()

   local oIni

   if File( GetEnv( "APPDATA" ) + "\fivedbu.ini" )
      INI oIni FILE GetEnv( "APPDATA" ) + "\fivedbu.ini"
         GET cDefRdd SECTION "Default" ENTRY "RDD"    DEFAULT "DBFCDX" OF oIni
         GET lShared SECTION "Default" ENTRY "Shared" DEFAULT .T. OF oIni
         GET nLanguage SECTION "Default" ENTRY "Language" ;
            DEFAULT FWSetLanguage() OF oIni
         GET lPijama SECTION "Default" ENTRY "Pijama" DEFAULT .T. OF oIni
         FWSetLanguage( nLanguage )
      ENDINI
   endif

return nil

//----------------------------------------------------------------------------//

function SavePreferences()

   local oIni

   DEFAULT nLanguage := FWSetLanguage()

   INI oIni FILE GetEnv( "APPDATA" ) + "\fivedbu.ini"
      SET SECTION "Default" ENTRY "RDD" OF oIni TO cDefRdd
      SET SECTION "Default" ENTRY "Shared" OF oIni TO lShared
      SET SECTION "Default" ENTRY "Language" OF oIni TO nLanguage
      SET SECTION "Default" ENTRY "Pijama" OF oIni TO lPijama
   ENDINI

return nil

//----------------------------------------------------------------------------//

function SaveRecord( aRecord, nRecNo )

   local n

   ( Alias() )->( DbGoTo( nRecNo ) )

   if ( Alias() )->( DbRLock( nRecNo ) )
      for n = 1 to Len( aRecord )
         if ( Alias() )->( FieldType( n ) ) != "+"
            ( Alias() )->( FieldPut( n, aRecord[ n ][ 2 ] ) )
         endif
      next
      ( Alias() )->( DbUnLock() )
      MsgInfo( FWString( "Record updated" ), FWString( "Information" ) )
   else
      MsgAlert( FWString( "Record in use, please try it again" ),;
                FWString( "Alert" ) )
   endif

return nil

//----------------------------------------------------------------------------//

function Sets()

   local oWndSets, oBar, oBrw, oMsgBar, a
   local aSets

   aSets := ;
   { { _SET_EXACT, "EXACT" }, { _SET_FIXED, "FIXED" }, { _SET_DECIMALS, "DECIMALS" }, { _SET_DATEFORMAT, "DATEFORMAT" } ;
   , { _SET_EPOCH, "EPOCH" }, { _SET_PATH,  "PATH", { || cGetDir( "Select Path", CURDIR() ) }  } ;
   , { _SET_DEFAULT,  "DEFAULT", { || cGetDir( "Default Directory" ) }  }, { _SET_EXCLUSIVE,  "EXCLUSIVE"  } ;
   , { _SET_SOFTSEEK, "SOFTSEEK" }, { _SET_UNIQUE, "UNIQUE" }, { _SET_DELETED, "DELETED" }, { _SET_STRICTREAD, "STRICTREAD" } ;
   , { _SET_OPTIMIZE, "OPTIMZE" }, { _SET_AUTOPEN, "AUTOOPEN" }, { _SET_AUTOSHARE, "AUTOSHARE" } ;
   , { _SET_LANGUAGE, "LANGUAGE", { "en", "es", "de", "it" } }, { _SET_HARDCOMMIT, "HARDCOMMIT" }, { _SET_CODEPAGE, "CODEPAGE" } ;
   , { _SET_OSCODEPAGE, "OSCODEPAGE" }, { _SET_TIMEFORMAT, "TIMEFORMAT" } }

   #ifndef __XHARBOUR__
      AAdd( aSets, { _SET_DBCODEPAGE, "DBCODEPAGE" } )
   #endif   

   AEval( aSets, { |a| ASize( a, 3 ) } )

   DEFINE WINDOW oWndSets TITLE "Sets" MDICHILD OF oWndMain

   DEFINE BUTTONBAR oBar OF oWndSets STYLEBAR SIZE 70, 70

   DEFINE BUTTON OF oBar PROMPT FWString( "Exit" ) RESOURCE "exit" ACTION oWndSets:End() GROUP

   @ 0, 0 XBROWSE oBrw OF oWndSets ARRAY aSets ;
      COLUMNS 2, { |x| If( x == nil, SET( oBrw:aRow[ 1 ] ), SET( oBrw:aRow[ 1 ], x ) ) } ;
      HEADERS "Set", "Value" FASTEDIT CELL LINES ;
      COLSIZES 100, 200 JUSTIFY .T.

   StyleBrowse( oBrw )
   if lPijama
      oBrw:bClrStd := { || If( oBrw:KeyNo() % 2 == 0, ;
                            { CLR_BLACK, RGB( 198, 255, 198 ) }, ;
                            { CLR_BLACK, RGB( 232, 255, 232 ) } ) }
      oBrw:bClrSel := { || { CLR_WHITE, RGB( 0x33, 0x66, 0xCC ) } }
   else
      oBrw:bClrStd := { || { nClrTxtBrw, nClrBackBrw } }
      oBrw:bClrSel := { || { nClrBackBrw, RGB( 0x33, 0x66, 0xCC ) } }
   endif
   WITH OBJECT oBrw
      :nMarqueeStyle := 3
      :nFreeze       := 1
      :lLockFreeze   := .t.
      :lHScroll      := .f.
      :bChange       := <||
         if HB_ISBLOCK( oBrw:aRow[ 3 ] )
            oBrw:aCols[ 2 ]:nEditType := EDIT_BUTTON
            oBrw:aCols[ 2 ]:bEditBlock := oBrw:aRow[ 3 ]
            oBrw:RefreshCurrent()
         elseif HB_ISARRAY( oBrw:aRow[ 3 ] )
            oBrw:aCols[ 2 ]:nEditType := EDIT_LISTBOX
            oBrw:aCols[ 2 ]:aEditListTxt := ;
            oBrw:aCols[ 2 ]:aEditListBound := oBrw:aRow[ 3 ]
            oBrw:RefreshCurrent()
         else
            oBrw:aCols[ 2 ]:nEditType := EDIT_GET
            oBrw:aCols[ 2 ]:bEditBlock := nil
            oBrw:aCols[ 2 ]:aEditListTxt := ;
            oBrw:aCols[ 2 ]:aEditListBound := nil
            oBrw:RefreshCurrent()
         endif
         return nil
         >
   END
   WITH OBJECT oBrw:aCols[ 2 ]
      :lWillShowABtn := .T.
      :nEditType     := EDIT_GET
      :bStrData      := { |x,o| If( HB_ISLOGICAL( o:Value ), If( o:Value, "TRUE", "FALSE" ), cValToChar( o:Value ) ) }
      :bKeyChar      := <|nKey,nFlags,brw,oCol|
         local cKey
         if HB_ISLOGICAL( oCol:Value ) .and. ;
            ( cKey := Chr( nKey ) ) $ "TtFf"
            if cKey $ "Tt"
               oCol:VarPut( .t. )
            else
               oCol:VarPut( .f. )
            endif
            brw:GoDown()
            return 0
         endif
         return nil
         >
   END
   oBrw:CreateFromCode()
   oWndSets:oClient  = oBrw
   oWndSets:oControl = oBrw

   oWndSets:nWidth   := 400

   DEFINE MSGBAR oMsgBar OF oWndSets 2010

   ACTIVATE WINDOW oWndSets

return nil

//----------------------------------------------------------------------------//

function Struct( cFileName, oBrwParent )

   local oDlg, oBrw, aFields := DbStruct(), n

/*
   if Empty( aFields )
      for n = 1 to oBrwParent:oRS:Fields:Count
         AAdd( aFields, { oBrwParent:oRS:Fields[ n - 1 ]:Name,;
                          oBrwParent:oRS:Fields[ n - 1 ]:Type,;
                          FWAdoFieldSize( oBrwParent:oRS, n ),;
                          FWAdoFieldDec( oBrwParent:oRS, n ) } )

      next
   endif
*/
   if oBrwParent:oRs != nil
      aFields  := FWAdoStruct( oBrwParent:oRs )
   elseif oBrwParent:oDbf != nil
      aFields  := oBrwParent:oDbf:aStructure
   endif

   DEFINE DIALOG oDlg TITLE Alias() + " " + FWString( "fields" ) SIZE 415, 600

   @ 0, 0 XBROWSE oBrw ARRAY aFields AUTOCOLS LINES NOBORDER STYLE FLAT ;
      HEADERS FWString( "Name" ), FWString( "Type" ), FWString( "Len" ),;
              FWString( "Dec" ) ;
      COLSIZES 150, 50, 80, 80

   StyleBrowse( oBrw )
   if lPijama
      oBrw:bClrStd := { || If( oBrw:KeyNo() % 2 == 0, ;
                            { CLR_BLACK, RGB( 198, 255, 198 ) }, ;
                            { CLR_BLACK, RGB( 232, 255, 232 ) } ) }
      oBrw:bClrSel := { || { CLR_WHITE, RGB( 0x33, 0x66, 0xCC ) } }
   else
      oBrw:bClrStd = { || { nClrTxtBrw, nClrBackBrw } }
      oBrw:bClrSel = { || { nClrBackBrw, RGB( 0x33, 0x66, 0xCC ) } }
   endif
   oBrw:CreateFromCode()

   oDlg:oClient = oBrw

   ACTIVATE DIALOG oDlg CENTERED ;
      ON INIT ( BuildStructBar( oDlg, oBrw, cFileName, oBrwParent ), oDlg:Resize(), oBrw:SetFocus() )

return nil

//----------------------------------------------------------------------------//

function BuildStructBar( oDlg, oBrw, cFileName, oBrwParent )

   local oBar

   DEFINE BUTTONBAR oBar OF oDlg STYLEBAR SIZE 70, 70

   DEFINE BUTTON OF oBar PROMPT FWString( "Code" ) RESOURCE "code" ;
      ACTION ( TxtStruct( oBrwParent ), oBrw:SetFocus() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Edit" ) RESOURCE "edit" ;
      ACTION ( New( Alias(), cFileName ) ) //, oDlg:End() )

   DEFINE BUTTON OF oBar PROMPT FWString( "Exit" ) RESOURCE "exit" ;
      ACTION oDlg:End() GROUP

return nil

//----------------------------------------------------------------------------//

function TxtStruct( oBrw )

   local cCode := "local aFields := { ", n, bClipboard

   if Empty( oBrw:oRs )
      for n = 1 to FCount()
         if n > 1
            cCode += Space( 19 )
         endif
         cCode += '{ "' + FieldName( n ) + '", "' + ;
                  FieldType( n ) + '", ' + ;
                  AllTrim( Str( FieldLen( n ) ) ) + ", " + ;
                  AllTrim( Str( FieldDec( n ) ) ) + " },;" + CRLF
      next
   else
      for n = 1 to oBrw:oRS:Fields:Count
         if n > 1
            cCode += Space( 19 )
         endif
         cCode += '{ "' + oBrw:oRS:Fields[ n - 1 ]:Name + '", "' + ;
                  FWAdoFieldType( oBrw:oRs, n ) + '", ' + ;
                  AllTrim( Str( FWAdoFieldSize( oBrw:oRs, n ) ) ) + ", " + ;
                  AllTrim( Str( FWAdoFieldDec( oBrw:oRs, n ) ) ) + ;
                  " },;" + CRLF
      next
   endif

   cCode = SubStr( cCode, 1, Len( cCode ) - 4 ) + " }" + CRLF + CRLF

   if Empty( oBrw:oRs )
      cCode += 'DbCreate( "myfile.dbf", aFields, "' + RddName() + '" )' + CRLF + CRLF
      cCode += 'USE myfile.dbf NEW VIA "' + RddName() + '"'
   endif

   n := 1
   while ! Empty( IndexKey( n ) )
      cCode += CRLF + CRLF + "INDEX ON " + IndexKey( n ) + ;
               If( RddName() == "DBFNTX", " TO ", " TAG " ) + ;
               OrdName( n )
      if ! Empty( OrdFor( n ) )
         cCode += ' FOR ' + OrdFor( n )
      endif
      n++
   end

   bClipboard = { | oDlg | AddClipboardButton( oDlg ), cCode }

   MemoEdit( bClipboard, FWString( "Code" ) )

return nil

//----------------------------------------------------------------------------//

function CopyToClipboard( cText )

   local oClip := TClipBoard():New()

   if oClip:Open()
      oClip:SetText( cText )
      oClip:Close()
   endif

   oClip:End()

return nil

//----------------------------------------------------------------------------//

function AddClipboardButton( oDlg, cText )

   local oBtn

   @ 242, 6 BUTTON oBtn PROMPT FWString( "Copy to clipboard" ) ;
      OF oDlg SIZE 180, 28 PIXEL ;
      ACTION ( CopyToClipboard( oDlg:aControls[ 1 ]:GetText() ),;
               MsgInfo( FWString( "Code copied to the clipboard" ) ) )

   @ 242, 191 BTNBMP BITMAP FWBitmap( "printer" ) OF oDlg SIZE 28, 28 ;
      ACTION oDlg:aControls[ 1 ]:Print() TOOLTIP FWString( "Print" )

return nil

//----------------------------------------------------------------------------//

function View( cFileName, oWnd )

   local cExt

   if ! File( cFileName )
      return nil
   endif

   cExt = Lower( cFileExt( cFileName ) )

   do case
      case cExt == "bmp"
           WinExec( "mspaint" + " " + cFileName )

      case cExt == "txt"
           WinExec( "notepad" + " " + cFileName )

      otherwise
           ShellExecute( oWnd:hWnd, "Open", cFileName )
   endcase

return nil

//----------------------------------------------------------------------------//

function Workareas( lShow )

   local oWndAreas, oBar, oBrw, oMsgBar, nAt
   local aAreas

   if ( nAt := AScan( WndMain():oWndClient:aWnd, { |o| o:cCaption == "Workareas" } ) ) > 0
      oWndAreas   := WndMain():oWndClient:aWnd[ nAt ]
      WITH OBJECT oWndAreas:oClient
         :aArrayData := ReadWorkAreas()
         :Refresh()
      END
      if lShow == .t.
         WITH OBJECT WndMain():oWndClient
            :ChildRestore( oWndAreas )
            :ChildActivate( oWndAreas )
         END
      endif
      return nil
   endif

   if !( lShow == .t. )
      return nil
   endif

   aAreas := ReadWorkAreas()

   DEFINE WINDOW oWndAreas TITLE "Workareas" MDICHILD OF oWndMain

   oWndAreas:SetSize( 1150, 565 )

   DEFINE BUTTONBAR oBar OF oWndAreas STYLEBAR SIZE 70, 70

   DEFINE BUTTON OF oBar PROMPT FWString( "Exit" ) RESOURCE "exit" ACTION oWndAreas:End() GROUP

   @ 0, 0 XBROWSE oBrw OF oWndAreas LINES ARRAY aAreas ;
      COLUMNS 1,2,3,4,5,6,7 ;
      HEADERS "Area", "Alias", "RddName", "RecCount", "RecNo", "Shared", "Path"

   oBrw:CreateFromCode()
   oBrw:bGotFocus := { || oBrw:aArrayData := ReadWorkAreas(), oBrw:Refresh() }

   oWndAreas:oClient  = oBrw
   oWndAreas:oControl = oBrw
   StyleBrowse( oBrw )

   DEFINE MSGBAR oMsgBar OF oWndAreas 2010

   ACTIVATE WINDOW oWndAreas

return nil

//----------------------------------------------------------------------------//

static function ReadWorkAreas()

   local aAreas   := {}
   local n

   for n := 1 to 255
      if !Empty( Alias( n ) )
         AAdd( aAreas, { n, Alias( n ), ( n )->( RDDNAME() ), ( n )->( LASTREC() ), ( n )->( RECNO() ), ;
                         ( n )->( DBINFO( DBI_SHARED ) ), TrueName( ( n )->( DBINFO( DBI_FULLPATH ) ) ) } )
      endif
   next

return aAreas

//----------------------------------------------------------------------------//

function New( cAlias, cFileName )

   local oDlg, oGet, oBrw, oBtn, cTitle, cNewAlias, oBrwNew, lCopy := .F.
   local cFieldName := Space( 10 ), cType := "Character", nLen := 10, nDec := 0
   local aFields := { Array( 4 ) }, cDbfName := Space( 8 ), aTemp
   local oLen, oDec, aType := { "AutoIncr", "Character", "Number", "Date", "Logical", "Memo" }
   local bChange := {|| If( cType == "AutoIncr",  ( nLen := 4,  nDec := 0, oDec:Disable()  ),),;
                        If( cType == "Character", ( nLen := 10, nDec := 0, oDec:Disable() ),),;
                        If( cType == "Number",    ( nLen := 10, nDec := 0, oDec:Enable()  ),),;
                        If( cType == "Date",      ( nLen := 8,  nDec := 0, oDec:Disable() ),),;
                        If( cType == "Logical",   ( nLen := 1,  nDec := 0, oDec:Disable() ),),;
                        If( cType == "Memo",      ( nLen := 10, nDec := 0, oDec:Disable() ),),;
                        oDlg:Update() }
   local bEdit := {|| IF ( !Empty (aFields[1,1]) ,;
                      (oBtn:Enable (),;
                       cFieldName := aFields[oBrw:nArrayAt,1] ,;
                       cType := aFields[oBrw:nArrayAt,2] ,;
                       cType := aType[ aScan(aType, {|x| Left(x,1) = cType} )],;
                       Eval (bChange) ,;
                       nLen := aFields[oBrw:nArrayAt,3] ,;
                       nDec := aFields[oBrw:nArrayAt,4] ,;
                       oGet:SetPos( 0 ),;
                       oGet:SetFocus(),;
                       oDlg:Update() ) ,) ;
                     }
   local bSave := { || oBtn:Disable (),;
                      aFields[ oBrw:nArrayAt, 1 ] := cFieldname,;
                      aFields[ oBrw:nArrayAt, 2 ] := if( Left( cType, 1 ) = "A", "+", Left( cType, 1 ) ), ;
                      aFields[ oBrw:nArrayAt, 3 ] := nLen,;
                      aFields[ oBrw:nArrayAt, 4 ] := nDec,;
                      oBrw:SetArray( aFields ),;
                      cFieldName := Space( 10 ),;
                      Eval( bChange ) ,;
                      oDlg:Update() ,;
                      oGet:SetPos( 0 ),;
                      oGet:SetFocus(),;
                      oBrw:GoBottom();
                      }

   if ! Empty( cAlias )
      aFields = ( cAlias )->( DbStruct() )
      cTitle = FWString( "Modify DBF structure" )
   else
      cTitle = FWString( "DBF builder" )
   endif

   DEFINE DIALOG oDlg TITLE cTitle SIZE 415, 500

   @ 0.5,  2 SAY FWString( "Field Name" ) OF oDlg SIZE 40, 8
   @ 0.5, 10 SAY FWString( "Type" ) OF oDlg SIZE 40, 8
   @ 0.5, 17 SAY FWString( "Len" )  OF oDlg SIZE 40, 8
   @ 0.5, 22 SAY FWString( "Dec" )  OF oDlg SIZE 20, 8

   @ 1.4, 1 GET oGet VAR cFieldName PICTURE "!!!!!!!!!!" OF oDlg SIZE 41, 11 UPDATE

   @ 1.3, 6.5 COMBOBOX cType ITEMS aType ;
      OF oDlg ON CHANGE Eval (bChange) UPDATE

   @ 1.4, 11.9 GET oLen VAR nLen PICTURE "999" OF oDlg SIZE 25, 11 UPDATE

   @ 1.4, 15.4 GET oDec VAR nDec PICTURE "999" OF oDlg SIZE 25, 11 UPDATE

   @ 0.9, 26 BUTTON FWString( "&Add" ) OF oDlg SIZE 45, 13 ;
      ACTION AddField( @aFields, @cFieldName, @cType, @nLen, @nDec, oGet, oBrw )

   @ 2.4, 26 BUTTON oBtn PROMPT FWString( "&Edit" ) OF oDlg SIZE 45, 13 ;
     ACTION Eval (bSave)

   @ 3.4, 26 BUTTON FWString( "&Delete" ) OF oDlg SIZE 45, 13 ;
     ACTION DelField( @aFields, @cFieldName, oGet, oBrw )

   @ 4.4, 26 BUTTON FWString( "Move &Up" ) OF oDlg SIZE 45, 13 ;
      ACTION If( oBrw:nArrayAt > 1,;
                 ( aTemp := aFields[ oBrw:nArrayAt ],;
                   aFields[ oBrw:nArrayAt ] := aFields[ oBrw:nArrayAt - 1 ],;
                   aFields[ oBrw:nArrayAt - 1 ] := aTemp,;
                   oBrw:GoUp() ),)

   @ 5.4, 26 BUTTON FWString( "Move D&own" ) OF oDlg SIZE 45, 13 ;
      ACTION If( oBrw:nArrayAt < Len( aFields ),;
                 ( aTemp := aFields[ oBrw:nArrayAt ],;
                   aFields[ oBrw:nArrayAt ] := aFields[ oBrw:nArrayAt + 1 ],;
                   aFields[ oBrw:nArrayAt + 1 ] := aTemp,;
                   oBrw:GoDown() ),)

   @ 11.8, 26 BUTTON FWString( "&Cancel" ) OF oDlg SIZE 45, 13 ;
      ACTION oDlg:End()

   @ 2.0, 1.8 SAY "Struct" OF oDlg SIZE 40, 8

   @ 3.0, 1 XBROWSE oBrw ARRAY aFields AUTOCOLS NOBORDER STYLE FLAT ;
      HEADERS FWString( "Name" ), FWString( "Type" ), FWString( "Len" ),;
              FWString( "Dec" ) ;
      COLSIZES 90, 55, 40, 40 ;
      SIZE 138, 183 OF oDlg ;
      ON DBLCLICK Eval (bEdit)

   StyleBrowse( oBrw )
   if lPijama
      oBrw:bClrStd := { || If( oBrw:KeyNo() % 2 == 0, ;
                            { CLR_BLACK, RGB( 198, 255, 198 ) }, ;
                            { CLR_BLACK, RGB( 232, 255, 232 ) } ) }
      oBrw:bClrSel := { || { CLR_WHITE, RGB( 0x33, 0x66, 0xCC ) } }
   else
      oBrw:bClrStd := { || { nClrTxtBrw, nClrBackBrw }  }
      oBrw:bClrSel := { || { nClrBackBrw, RGB( 0x33, 0x66, 0xCC ) } }
   endif
   oBrw:CreateFromCode()

   @ 15.3, 1.4 SAY FWString( "DBF Name:" ) OF oDlg SIZE 40, 8

   if ! Empty( cAlias )
      cDbfName = cGetNewAlias( cAlias )
   endif

   @ 17.7, 6 GET cDbfName PICTURE "!!!!!!!!!!!!" OF oDlg SIZE 100, 11

   @ 12.8, 26 BUTTON If( Empty( cAlias ), FWString( "&Create" ), FWString( "&Save" ) ) ;
      OF oDlg SIZE 45, 13 ;
      ACTION ( If( ! Empty( cDbfName ) .and. Len( aFields ) > 0,;
          DbCreate( AllTrim( cDbfName ), aFields, If( Empty( cAlias ),,(cAlias)->(RDDNAME()) ) ),), oDlg:End(),;
          lCopy := .T.,;
          oBrwNew := Open( hb_CurDrive() + ":\" + CurDir() + "\" + AllTrim( cDbfName ) ) )

   ACTIVATE DIALOG oDlg CENTERED ;
      ON INIT ( Eval ( bChange ), oBtn:Disable() ) ;
      VALID ! GETKEYSTATE( VK_ESCAPE )

   if ! Empty( cAlias ) .and. lCopy
      APPEND FROM ( cFileName )
      oBrwNew:Refresh()
   endif

return nil

//----------------------------------------------------------------------------//

function AddField( aFields, cFieldName, cType, nLen, nDec, oGet, oBrw )

   local cSymbol     := ""
   if Empty( cFieldName )
      oGet:SetPos( 0 )
      return nil
   endif

   if Len( aFields ) == 1 .and. Empty( aFields[ 1 ][ 1 ] )
      aFields = { { cFieldName, ;
                    if( Upper( Left( cType, 1 ) ) = "A", "+", Upper( Left( cType, 1 ) ) ), nLen, nDec } }
   else
      AAdd( aFields, { cFieldName, ;
                       if( Upper( Left( cType, 1 ) ) = "A", "+", Upper( Left( cType, 1 ) ) ), nLen, nDec } )
   endif

   oBrw:SetArray( aFields )
   oGet:VarPut( cFieldName := Space( 10 ) )
   oGet:SetPos( 0 )
   oGet:SetFocus()
   oBrw:GoBottom()

return nil

//----------------------------------------------------------------------------//

function Paste()

   if oWndMain:oWndActive != nil
      oWndMain:oWndActive:Paste()
   endif

return nil

//----------------------------------------------------------------------------//

function ChrKeep()

return ""

function Soundex()

return ""

FUNCTION SQLEXEC( cQuery )

    LOCAL cCns := "Your connectionstring here"

    LOCAL oCn := CREATEOBJECT( "ADODB.Connection" )

    oCn:CursorLocation = adUseClient

    oCn:Open( cCns )

    oCn:Execute( cQuery )

    oCn:Close()

RETURN NIL

//----------------------------------------------------------------------------//

Function FI_EdtHex()

   local cFile
   local cData

   cFile := cGetFile32( "Executable file (*.exe) |*.exe|" + ;
                        "Lib file (*.lib) |*.lib|" + ;
                        "Dll file (*.dll) |*.dll|" + ;
                        "Obj file (*.obj) |*.obj|" + ;
                        "Obj file (*.obj) |*.dbf|" + ;
                        "Obj file (*.obj) |*.ntx|" + ;
                        "Obj file (*.obj) |*.cdx|" + ;
                        "All Files (*.*) |*.*|", ;
                        "Select a file to open", 0, ;
                        hb_CurDrive() + ":\" + CurDir(), , )
   cData := MEMOREAD( cFile )

   XbrHexEdit( @cData, cFile, .t., .t. )

Return nil

//----------------------------------------------------------------------------//

Static Function StyleBrowse( oBrw )

   WITH OBJECT oBrw
      :l2007               := .F.
      :nMarqueeStyle       := MARQSTYLE_HIGHLROW
      :lFullGrid           := .T.
      :lRowDividerComplete := .T.
      :lColDividerComplete := .T.
      //:nRowDividerStyle    := LINESTYLE_NOLINES
      :nColDividerStyle    := LINESTYLE_LIGHTGRAY
      :nRowDividerStyle    := LINESTYLE_LIGHTGRAY
      :nHeaderHeight       := 23
      //:nFooterHeight     := oBrw:nHeaderHeight
      :nRowHeight          := oBrw:nHeaderHeight
      //:nStretchCol         := 1
      :nFreeze             := 1
      if lPijama
         :SetColor( CLR_BLACK, RGB( 232, 255, 232 ) )
      else
         :SetColor( nClrTxtBrw, nClrBackBrw )
      endif
   END

Return nil

//----------------------------------------------------------------------------//
