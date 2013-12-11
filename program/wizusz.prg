#INCLUDE Include\ALLDEFS.H
#DEFINE MAX_PARAMS 254
#DEFINE MAX_ROWS 180
#DEFINE WIZARD_CLOSING	.T.
#DEFINE BOOLEAN_TRUE ".T."
#DEFINE BOOLEAN_FALSE ".F."
#DEFINE BOOLEAN_SQL_TRUE	"1"
#DEFINE	BOOLEAN_SQL_FALSE	"0"
#DEFINE VFP_BETWEEN	"BETWEEN"
#DEFINE VFP_COMMA	", "

* jvf 9/1/99
#DEFINE VIEW_NAME_EXTENSION_LOC	"v_"
* New local table name for wiz report if was dropped
#DEFINE DROPPED_TABLE_STATUS_LOC "Table Dropped"
#DEFINE BULK_INSERT_FILENAME "BulkIns.out"

* Use tilde rather than comma in case commas in char fields
#DEFINE BULK_INSERT_FIELD_DELIMITER	~

*{ Add JEI -RKR - 2005.03.25
#DEFINE gcQT  CHR( 34 )
#DEFINE gc2QT  ['']
*} Add JEI -RKR - 2005.03.25

*******************************************
DEFINE CLASS UpsizeEngine AS WizEngineAll of WZEngine.prg
*******************************************

*ODBC-related properties
MasterConnHand    = 0
CONNECTSTRING     = ""
DataSourceName    = ""
ServerType        = "SQL Server"
CurrentServerType = ""
ServerVer         = 0
UserConnection    = ""
ViewConnection    = ""
UserName          = ""
USERID            = 0
MyError           = 0
FETCHMEMO         = .F.


*Navigation variables
DeviceRecalc         = .T.
AnalyzeTablesRecalc  = .T.
AnalyzeFieldsRecalc  = .T.
AnalyzeIndexesRecalc = .T.
ChooseTargetDBRecalc = .T.
TableCboRecalc       = .T.
GetRiInfoRecalc      = .T.
EligibleRelsRecalc   = .T.
DataSourceChosen     = .F.
DSNChange            = .F.
NoDataSourceRightNow = .F.
GetConnDefsRecalc    = .T.
DeviceLogChosen      = .F.
DeviceDBChosen       = .F.
SourceDBChosen       = .F.
GridFilled           = .F.


*Device properties
DeviceNumbersFree = 0
DeviceDBName      = ""
DeviceDBPName     = ""
DeviceDBSize      = 0
DeviceDBNumber    = 0
DeviceDBNew       = .F.
DeviceLogName     = ""
DeviceLogPName    = ""
DeviceLogSize     = 0
DeviceLogNumber   = 0
DeviceLogNew      = .F.
MasterPath        = ""
NewDeviceCount    = 0
DefaultFreeSpace  = 0
DeviceDBInDefa    = .F.		&&used to in Device class, DeviceThingOK method
DeviceLogInDefa   = .F.
DBonDefault       = .F.


*Server Database properties
CreateNewDB     = .F.
ServerFreeSpace = -5
ServerDBName    = ""
ServerDBSize    = 0
ServerLogSize   = 0


*Oracle Wizard properties

*Tables Selection Step properties
TBFoxTableSize = 0
TBFoxIndexSize = 0

*Tablespaces Step properties
TSNewTableTS      = .F.
TSNewIndexTS      = .F.
TSTableTSName     = ""
TSIndexTSName     = ""
TSDefaultTSName   = ""
TSPermTablespaces = .T.
TSNew             = .F.
TSDone            = .F.


*Tablespace File Step properties
TSFTableFileName = ""
TSFIndexFileName = ""
TSFTableFileSize = 0
TSFIndexFileSize = 0
TSFDone          = .F.

*Cluster Table Step properties
CLClustersDone = .F.
CLTablesDone   = .F.
CLKeysDone     = .F.

*Export properties
SourceDB            = ""
ExportIndexes       = .T.
ExportValidation    = .T.
ExportRelations     = .T.
ExportDRI           = .F.
ExportStructureOnly = .F.
ExportDefaults      = .T.
ExportTimeStamp     = .T.
ExportTableToView   = .F.
ExportViewToRmt     = .T.
ExportSavePwd       = .F.
OVERWRITE           = .F.					&&If .T., existing tables are overwritten
NullOverride        = 1
ExportClustered     = .F.				&& Default to Primary Keys not being Clustered
ViewNameExtension   = VIEW_NAME_EXTENSION_LOC
ViewPrefixOrSuffix  = 1
DropLocalTables     = .F.


*Names of tables created and aliases
EnumFieldsTbl    = ""
MappingTable     = ""					&& Same as EnumfieldsTbl, used again
EnumTablesTbl    = ""
EnumClustersTbl  = ""				&& Cluster table name
EnumIndexesTbl   = ""
DeviceTable      = ""					&& Same table, used twice
DeviceTableAlias = ""				&& Same as DeviceTable, used again by the log device screen
ViewsTbl         = ""
EnumRelsTbl      = ""
ScriptTbl        = ""					&& Basically a big memo field for holding generated sql
ErrTbl           = ""
OraNames         = ""						&& Just a cursor of Oracle index names

*Action properties
DoUpsize=.T.
DoScripts=.F.
DoReport=.T.
* jvf 08/13/99 Now can choose for all tables
TimeStampAll=0					&& add timestamp column to all tables
IdentityAll=0					&& add identity column to all tables

*Permissions properties
Perm_Device		=	.T.
Perm_Table		=	.T.
Perm_Database	=	.T.
Perm_Default	=	.T.
Perm_Sproc		=	.T.
Perm_Index		=	.T.
Perm_Trigger	=	.T.
Perm_AltTS		=	.T.
Perm_CreaTS		=	.T.
Perm_Cluster	=	.T.
Perm_UnlimTS	=	.T.

UserUpsizeMethod = 0	&& upsize method chosen by the user.

*Arrays of server datatypes, one for each FoxPro type
DIMENSION C[1]
DIMENSION N[1]
DIMENSION B[1]
DIMENSION L[1]
DIMENSION M[1]
DIMENSION Y[1]
DIMENSION D[1]
DIMENSION T[1]
DIMENSION P[1]
DIMENSION G[1]
DIMENSION F[1]
DIMENSION I[1]

* rmk - 01/06/2004 - support for new types in VFP 9
DIMENSION V[1]
DIMENSION Q[1]  && varbinary
DIMENSION W[1]	&& blob

*Other
UserInput          = ""			&& Inputbox always puts user input here
ZeroDefaultCreated = .F.	&& Flag set true after zero default has been created
ScriptTblCreated   = .F.
ProcessingOutput   = .F.	&& Set to .T. after user clicks Finish button
OldRow             = 1				&& Used by type mapping grid to see if row changed
OldType            = ""				&& Used by type mapping grid in case user wants to undo change
NormalShutdown     = .F.		&& Flag used to prevent analysis files from getting nuked
DataErrors         = .F.
NewProjName        = ""
PwdInDef           = .F.			&& See comment in page9 activate method
RealClick          = .T.			&& Flag for page9
SaveErrors         = .T.			&& Save error tables, set in BuildReport method
ZDUsed             = .F.				&& So that sql for zero default included ( if appropriate ) in sql script
FiltCond           = ""				&& Used on type mapping page
NewDir             = ""				&& Directory created to store tables etc. the upsizing wizard creates
CreatedNewDir      = .F.
KeepNewDir         = .F.
SQLServer          = .T.			&& True if we're connected to SQL Server
TimeStampName      = ""		&& Name used for all timestamp fields ( if any ) that are added
IdentityName       = ""		&& Name used for all identity fields ( if any ) that are added
TruncLog           = -1			&& Status of Trunc.log on chkpt. option of database
cFinishMsg         = ""			&& message to display at end

* Properties added from JEI
DIMENSION aChooseViews[1]
DIMENSION aDataErrTbls[1]
ServerISLocal    = .F.		&& .T. - SQL Server is on current machine, .F. - SQL Server is on remote machine
NotUseBulkInsert = .f.	&& .T. - Always not use Bulk Insert

* Extension object.
oExtension = .NULL.

* Value to use for blank Date and DateTime values.
BlankDateValue = .NULL.

* Default folder to use for report information.
ReportDir = ''

FUNCTION Init( tcReturnToProc, tcProcedure )
dodefault( tcReturnToProc, tcProcedure )
this.ReportDir = fullpath( NEW_DIRNAME_LOC )


* When the connection handle is set, get the connection string, set the
* connection properties, and get the server version.

FUNCTION MasterConnHand_Assign( tnHandle )
with This
	.MasterConnHand = tnHandle
*** DH 12/06/2012: added IF as suggested by Matt Slay to prevent problem in
*** SetConnProps when tnHandle is 0.
	IF tnHandle > 0
		.ServerVer      = .GetServerVersion()
		.ConnectString  = sqlgetprop( tnHandle, 'ConnectString' )
		.SetConnProps()
*** DH 12/06/2012: added ENDIF to match IF
	ENDIF
endwith


* Determine the server version.

FUNCTION GetServerVersion
local lnhdbc, ;
	lcName, ;
	lcbName, ;
	lnReturn
declare short SQLGetInfo in odbc32 ;
	integer hdbc, integer fInfoType, string @cName, ;
	integer cNameMax, integer @cbName
lnhdbc   = sqlgetprop( this.MasterConnHand, 'ODBChdbc' )
lcName   = space( 128 )
lcbName  = 0
lnReturn = sqlgetinfo( lnhdbc, SQL_DBMS_VER, @lcName, 128, @lcbName )
IF lnReturn = SQL_SUCCESS
	lnReturn = val( left( lcName, at( chr( 0 ), lcName ) - 1 ))
ELSE
	lnReturn = -1
ENDIF
return lnReturn


FUNCTION ProcessOutput
LOCAL lcSQL, lcMsg
public oEngine

this.ProcessingOutput=.T.

*Let user bail if they want
oEngine = This
ON ESCAPE OEngine.Esc_proc

this.InitTherm( 'Preparing server for upsizing', 100 )
this.GetChooseView()

IF this.SQLServer
	* SQL Server: create devices
	this.UpdateTherm( 5, 'Setting up devices' )
	this.DealWithDevices()

	* SQL Server: create database
	this.UpdateTherm( 10, 'Creating target database' )
	this.CreateTargetDB()

	this.UpdateTherm( 15, 'Checking for local server' )
	this.CheckForLocalServer()

	* Connect to target database
	this.UpdateTherm( 30, 'Connecting to database' )
	lcSQL="use " + ALLTRIM( this.ServerDBName )
	this.ExecuteTempSPT( lcSQL )

	* set option Trunc.Log on chkpt. for database on server
	this.UpdateTherm( 35, 'Setting database options' )
	this.TruncLogOn()
ENDIF

*Make sure everything's been analyzed that's going to be upsized
this.UpdateTherm( 90, 'Setting database options' )
this.AnalyzeFields()

*create tablespaces, clusters and cluster indexes
IF this.ServerType="Oracle"
	this.CreateTablespaces()
	this.CreateClusters()
ENDIF
raiseevent( This, 'CompleteProcess' )

*create tables
this.CreateTables()

*send data
this.SendData()

*create indexes
this.AnalyzeIndexes()

*build RI code
IF this.ExportRelations THEN
	this.BuildRiCode()
ENDIF

IF this.ExportIndexes THEN
	this.CreateIndexes()
ENDIF

*deal with defaults and validation rules
this.DefaultsAndRules()

*create put rules and RI code into triggers
this.CreateTriggers()

*redirect app
this.RedirectApp()

*do report stuff
this.CreateScript

this.BuildReport()

*done
*reset option trunc. log on chkpt. to initial value
IF this.SQLServer
    this.TruncLogOff()
ENDIF

*test if all the tables were upsized. If not display warning message.
this.UpsizeComplete()
release oEngine


* Ensures the specified table is selected for export.

FUNCTION SelectTable( tcTable )
local llReturn, ;
	lnSelect
llReturn = used( this.EnumTablesTbl )
IF llReturn
	lnSelect = select()
	select ( this.EnumTablesTbl )
	locate for lower( trim( TBLNAME )) == lower( alltrim( tcTable ))
	llReturn = found()
	IF llReturn
		replace EXPORT with .T.
	ENDIF
	select ( lnSelect )
ENDIF
return llReturn


* Select all tables for export.

FUNCTION SelectAllTables( tcTable )
local llReturn
llReturn = used( this.EnumTablesTbl )
IF llReturn
	replace all EXPORT with .T. in ( this.EnumTablesTbl )
ENDIF
return llReturn


FUNCTION ERROR
PARAMETERS nError, cMethod, nLine, oObject
LOCAL lcErrMsg, lnServerError

=AERROR( aErrArray )
nError=aErrArray[1]
this.MyError=nError

DO CASE
CASE nError=1523
    *User hit cancel button in ODBC dialog
    this.HadError=.T.
    RETURN .F.

CASE nError=15
    *Not a table ( probably means the table is corrupt )
    this.HadError=.T.
    RETURN

CASE nError=108
    *File opened exclusive by someone ELSE
    this.HadError=.T.
    RETURN

CASE nError=1705 OR nError=3
    *File access denied
    *Should be caused only when the wizard tries to open
    *all tables in database exclusively
    this.HadError=.T.
    RETURN

CASE nError=1976
    *Table is not marked for current database
    *Like 1705 above, should only happen in the this.Upsizable function
    this.HadError=.T.
    RETURN

CASE nError=1984
    *Table definition and DBC are out of sync
    *( Another error that could result from this.Upsizable )
    this.HadError=.T.
    RETURN

CASE nError=1498
    *Attempt to create remote view failed because of bogus SQL
    this.HadError=.T.
    RETURN

CASE nError=1160 and not this.lQuiet
    *Out of disk space
    lnUserChoice=MESSAGEBOX( NO_DISK_SPACE_LOC, RETRY_CANCEL+ICON_EXCLAMATION, TITLE_TEXT_LOC )
    IF lnUserChoice=RETRY_CHOICE THEN
        RETRY
    ENDIF
CASE nError=1577
    *Table "name" is referenced in a relation
    *Occurs when drop table that's in a relation
    this.HadError=.T.
    RETURN

ENDCASE

*Unhandled errors->We're dead
WizEngineAll::ERROR( nError, cMethod, nLine, oObject )



    FUNCTION Esc_proc

        *If the user hits escape when the wizard is processing output
        CLEAR TYPEAHEAD
        IF this.lQuiet or ;
        	MESSAGEBOX( ESCAPE_CONT_LOC, ICON_EXCLAMATION+YES_NO_BUTTONS, TITLE_TEXT_LOC )=USER_YES
            SET ESCAPE OFF
            this.Die()
            RETURN TO MASTER
        ELSE
            RETURN
        ENDIF



    FUNCTION DESTROY
        LOCAL lcAction, lcDelDir, lcSQL, llRetVal

        *don't nuke the analysis tables if the user wants a report
        lcAction=IIF( this.DoReport AND this.ProcessingOutput, "Close", "Delete" )

        *Deal with tables that are part of the upsizing wizard project ( that don't need to be deleted )
        this.DisposeTable( "Keywords", "Close" )
        this.DisposeTable( "ExprMap", "Close" )
        this.DisposeTable( "TypeMap", "Close" )

        *Close this cursor ( created in SQL 95 case )
        this.DisposeTable( this.OraNames, "Close" )

        *Clean up device stuff if it exists
        this.DeviceCleanUp( WIZARD_CLOSING )

        * jvf 8/17/99 Don't close them if we've already dropped 'em.
        IF !this.DropLocalTables
            *Close user tables
            this.CloseUserTables
        ENDIF

        *Close/delete analysis tables
        this.AnalCleanUp( lcAction, WIZARD_CLOSING )

        * Restore SQL Server 7.0 compatibility levels
*** DH 07/24/2013: compatibility level isn't used anymore.
*        IF ATC( "SQL Server", this.ServerType )#0 AND this.nSQL7CompLevel>=70
*            lcSQL=[sp_dbcmptlevel ]+ALLTRIM( this.ServerDBName )+[, ]+TRANS( this.nSQL7CompLevel )
*            llRetVal=this.ExecuteTempSPT( lcSQL )
*        ENDIF

        *Connection cleanup
        IF this.MasterConnHand>0 THEN
            SQLDISCONN( this.MasterConnHand )
        ENDIF

        *Clean up error tables
        IF !this.SaveErrors THEN
            this.DisposeTable( this.ErrTbl, "Delete" )
            IF !EMPTY( this.aDataErrTbls ) THEN
                FOR I=1 TO ALEN( this.aDataErrTbls, 1 )
                    this.DisposeTable( this.aDataErrTbls[i, 2], "Delete" )
                NEXT
            ENDIF
        ENDIF

        *Nix newly created directory if appropriate
        IF this.CreatedNewDir AND !this.KeepNewDir THEN
            lcDelDir=FULLPATH( SYS( 5 ))
*** DH 07/02/2013: CD to original folder
*            CD ..
			cd ( this.aEnvironment[29, 1] )
*** DH 07/02/2013: added TRY around RD command just in case
			try
	            RMDIR ( lcDelDir )
	        catch
	        endtry
        ENDIF

        *Release memory variables
        RELEASE aOpenDatabases, aDataSources, aServerDatabases, ;
            aConnDefs, aExport, aTablesToExport, ;
            aDataErrTbls, aDeviceNumbers, ;
            aClusters, aValidTables, aClusterTables, aServerTablespaces, ;
            aDataFiles, aSelectedTablespaces, aSelectList, aFiles

        * restore fetch memo option
        =CURSORSETPROP( 'FetchMemo', this.FETCHMEMO, 0 )

        *- save messagebox until we've have finished clearnup
        IF !EMPTY( this.cFinishMsg ) and not this.lQuiet
            =MESSAGEBOX( this.cFinishMsg, ICON_EXCLAMATION, TITLE_TEXT_LOC )
        ENDIF

        WizEngineAll::DESTROY



    FUNCTION AnalCleanUp
        *Called by main Cleanup proc and if the user changes the source database
        PARAMETERS lcAction, llWizardClosing

        IF !llWizardClosing

            *Reset flags
            this.AnalyzeTablesRecalc=.T.
            this.AnalyzeIndexesRecalc=.T.
            this.AnalyzeFieldsRecalc=.T.
            this.GetConnDefsRecalc=.T.
            this.GridFilled=.F.
            this.EligibleRelsRecalc=.T.
            this.GetRiInfoRecalc=.T.

        ENDIF

        IF this.NormalShutdown THEN
            IF this.DoReport AND this.ProcessingOutput THEN
                lcAction="Close"
            ELSE
                lcAction="Delete"
            ENDIF
        ELSE
            lcAction="Delete"
        ENDIF

        this.DisposeTable( this.MappingTable, "Close" )
        this.DisposeTable( this.EnumFieldsTbl, lcAction )
        this.DisposeTable( this.EnumTablesTbl, lcAction )
        this.DisposeTable( this.EnumIndexesTbl, lcAction )
        this.DisposeTable( this.ViewsTbl, lcAction )
        this.DisposeTable( this.EnumRelsTbl, lcAction )

        *Don't want to delete error or script table unless user hit cancel or error
        IF this.NormalShutdown THEN
            lcAction="Close"
        ENDIF
        this.DisposeTable( this.ErrTbl, lcAction )
        this.DisposeTable( this.ScriptTbl, lcAction )

        IF !llWizardClosing
            *Reset analysis table-related variables to original default values
            this.EnumRelsTbl=""
            this.MappingTable=""
            this.EnumFieldsTbl=""
            this.EnumTablesTbl=""
            this.EnumIndexesTbl=""
            this.ViewsTbl=""
            this.UserConnection=""
        ENDIF



    FUNCTION DeviceCleanUp
        PARAMETERS WizardClosing
        *Called by main Cleanup proc and if the user changes the data source


        *Clean up the device tables
        this.DisposeTable( this.DeviceTableAlias, "Close" )
        this.DisposeTable( this.DeviceTable, "Delete" )

        IF !WizardClosing THEN
            *Reset all Device-related properties to their original default values
            this.UserInput=""
            this.DeviceTable=""
            this.DeviceTableAlias=""
            this.DeviceNumbersFree=0
            this.DeviceDBName=""
            this.DeviceDBPName=""
            this.DeviceDBSize=0
            this.DeviceDBNumber=0
            this.DeviceDBNew=.F.
            this.DeviceDBChosen=.F.
            this.DeviceLogName=""
            this.DeviceLogPName=""
            this.DeviceLogSize=0
            this.DeviceLogNumber=0
            this.DeviceLogNew=.F.
            this.DeviceLogChosen=.F.
            this.MasterPath=""
            this.NewDeviceCount=0
            this.ChooseTargetDBRecalc=.T.
            this.DataSourceChosen=.F.
            DeviceDBInDefa=.F.
            DeviceLogInDefa=.F.
            DBonDefault=.F.

            *Set recalc flag on
            this.DeviceRecalc=.T.

        ENDIF



    FUNCTION CloseUserTables
        LOCAL lcEnumTablesTbl

        *Closes user tables that weren't open before running the upsizing wizard

        lcEnumTablesTbl=this.EnumTablesTbl
        IF lcEnumTablesTbl=="" THEN
            RETURN
        ENDIF

        DIMENSION aTableArray[1]
        SELECT CursName FROM ( lcEnumTablesTbl ) ;
            WHERE Upsizable=.T. AND PreOpened=.F. ;
            INTO ARRAY aTableArray
        IF !EMPTY( aTableArray ) THEN
            FOR I=1 TO ALEN( aTableArray )
                *close the table
                SELECT ( RTRIM( aTableArray[i] ))
                USE
            NEXT
        ENDIF



    FUNCTION DisposeTable
        PARAMETERS lcTableName, lcAction
        *Close the table if it's open
        IF USED( lcTableName )
            SELECT ( lcTableName )
            USE
        ENDIF

        *Clean up any backup files incidentally created
        IF FILE( lcTableName+".bak" ) THEN
            DELETE FILE lcTableName+".bak"
        ENDIF
        IF FILE( lcTableName+".tbk" ) THEN
            DELETE FILE lcTableName+".tbk"
        ENDIF

        *Delete file if appropriate
        IF lcAction="Delete" THEN
            IF FILE( lcTableName+".dbf" ) THEN
                DELETE FILE lcTableName+".dbf"
            ENDIF
            IF FILE( lcTableName+".cdx" ) THEN
                DELETE FILE lcTableName+".cdx"
            ENDIF
            IF FILE( lcTableName+".fpt" ) THEN
                DELETE FILE lcTableName+".fpt"
            ENDIF
        ENDIF



    FUNCTION InitTherm
        PARAMETERS lcTitle, lnBasis, lnInterval

        *This routine creates the thermometer or initializes it if it already exists

        IF !this.ProcessingOutput THEN
            RETURN
        ENDIF
		raiseevent( This, 'InitProcess', lcTitle, lnBasis )


	FUNCTION InitProcess( tcTitle, tnBasis )

	FUNCTION UpdateProcess( tnProgress, tcTask )

	FUNCTION CompleteProcess


    FUNCTION UpDateTherm
        PARAMETERS lnProgress, lcTask
		local lcTaskDescrip

        *This routine updates the thermometer if processing output

        IF !this.ProcessingOutput THEN
            RETURN
        ELSE
			lcTaskDescrip = iif( pcount() = 2, lcTask, '' )
			raiseevent( This, 'UpdateProcess', lnProgress, lcTaskDescrip )
        ENDIF



    FUNCTION DealWithDevices
        *Creates devices if appropriate
        LOCAL lnRetVal, lcDevicePName, lnDeviceNumber

        IF this.ServerVer>=7
            this.DeviceDBNew = .F.
            this.DeviceLogNew = .F.
        ENDIF

        lcDevicePName=""
        lnDeviceNumber=0

        IF this.DeviceDBNew = .T. THEN
            lnRetVal=this.CreateDevice( "DB", @lcDevicePName, @lnDeviceNumber )
            this.DeviceDBPName=lcDevicePName
            this.DeviceDBNumber=lnDeviceNumber
        ENDIF
        IF this.DeviceLogNew = .T. THEN
            lnRetVal=this.CreateDevice( "Log", @lcDevicePName, @lnDeviceNumber )
            this.DeviceLogPName=lcDevicePName
            this.DeviceLogNumber=lnDeviceNumber
        ENDIF



    FUNCTION CreateTargetDB
        LOCAL lcSQL, lcMsg, lnErr, lcErrMsg, lnCompLevel

        *Build the SQL statement
        IF this.CreateNewDB THEN
            IF this.ServerVer>=7
                lcSQL= "create database " + RTRIM( this.ServerDBName )
            ELSE
                lcSQL= "create database " + RTRIM( this.ServerDBName ) + " on "
                lcSQL= lcSQL + RTRIM( this.DeviceDBName ) + "=" + ALLTRIM( STR( this.ServerDBSize ))
                IF this.ServerLogSize<> 0 THEN
                    lcSQL= lcSQL+ " log on " + RTRIM( this.DeviceLogName )  + "=" + ALLTRIM( STR( this.ServerLogSize ))
                ENDIF
            ENDIF
        ELSE
            RETURN
        ENDIF

* If we have an extension object and it has a CreateTargetDB method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'CreateTargetDB', 5 ) and ;
			not this.oExtension.CreateTargetDB( This )
			return
		ENDIF

        *Execute if appropriate
        IF this.DoUpsize THEN
            lcMsg=STRTRAN( CREATING_DATABASE_LOC, '|1', RTRIM( this.ServerDBName ))
            this.InitTherm( lcMsg, 0, 0 )
            this.UpDateTherm( 0, TAKES_AWHILE_LOC )
            SQLSETPROP( this.MasterConnHand, "QueryTimeOut", 600 )
            this.MyError=0
            IF !this.ExecuteTempSPT( lcSQL, @lnErr, @lcErrMsg ) THEN
                IF lnErr=262 THEN
                    *User doesn't have CREATE DATABASE permissions
                    lcMsg=STRTRAN( NO_CREATEDB_PERM_LOC, '|1', RTRIM( this.DataSourceName ))
                ELSE
                    *Something else went wrong
                    lcMsg=STRTRAN( CREATE_DB_FAILED_LOC, '|1', RTRIM( this.ServerDBName ))
                ENDIF
				IF this.lQuiet
					.HadError     = .T.
					.ErrorMessage = lcMsg
				ELSE
	                MESSAGEBOX( lcMsg, ICON_EXCLAMATION, TITLE_TEXT_LOC )
				ENDIF
                this.Die
            ENDIF
            SQLSETPROP( this.MasterConnHand, "QueryTimeOut", 30 )
			raiseevent( This, 'CompleteProcess' )
        ENDIF

        *Stash sql for script
        this.StoreSQL( lcSQL, CREATE_DBSQL_LOC )

*** DH 07/24/2013: compatibility level isn't used anymore.
*        IF this.ServerVer>=7
*            lnCompLevel=0
*            lcSQL = "select cmptlevel from master.dbo.sysdatabases where name='"+this.ServerDBName+"'"
*            IF this.SingleValueSPT( lcSQL, @lnCompLevel, "cmptlevel" )
*                this.nSQL7CompLevel = lnCompLevel
*                lcSQL=[sp_dbcmptlevel ]+this.ServerDBName+[, 65]
*                this.ExecuteTempSPT( lcSQL )
*            ENDIF
*        ENDIF



    FUNCTION CreateTablespaces

* If we have an extension object and it has a CreateTableSpaces method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'CreateTableSpaces', 5 ) and ;
			not this.oExtension.CreateTableSpaces( This )
			return
		ENDIF

        * create new Table tablespace or add file to an existing one
        IF this.TSNewTableTS
            this.CreateTS( this.TSTableTSName, this.TSFTableFileName, this.TSFTableFileSize )
        ELSE
            IF !EMPTY( this.TSFTableFileName )
                this.CreateDataFile( this.TSTableTSName, this.TSFTableFileName, this.TSFTableFileSize )
            ENDIF
        ENDIF

        *If the tablespace for tables and indexes are the same, bail
        IF this.TSTableTSName=this.TSIndexTSName THEN
            RETURN
        ENDIF

        * create new Index tablespace or add file to an existing one
        IF this.TSNewIndexTS
            this.CreateTS( this.TSIndexTSName, this.TSFIndexFileName, this.TSFIndexFileSize )
        ELSE
            IF !EMPTY( this.TSFIndexFileName )
                this.CreateDataFile( this.TSIndexTSName, this.TSFIndexFileName, this.TSFIndexFileSize )
            ENDIF
        ENDIF



    FUNCTION CreateTables
        LOCAL lcEnumTablesTbl, llRetVal, dummy, llTableExists, MyMessageBox, ;
            lnTableCount, lcMsg, lcSQL, lnError, lcErrMsg, lcCRLF, lcRmtTable

* If we have an extension object and it has a CreateTables method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'CreateTables', 5 ) and ;
			not this.oExtension.CreateTables( This )
			return
		ENDIF

        dummy = "x"
        lcCRLF = CHR( 10 ) + CHR( 13 )

        * first go create a SQL statement for all tables the user chose to export
        this.CreateTableSQL

        * then if the user wants to upsize and not just create scripts, export the tables
        IF this.DoUpsize THEN
            lcEnumTablesTbl = this.EnumTablesTbl
            SELECT ( lcEnumTablesTbl )

            * For thermometer
            SELECT COUNT( * ) FROM ( lcEnumTablesTbl ) WHERE EXPORT = .T. INTO ARRAY aTableCount
            lnTableCount = 0
            this.InitTherm( CREATING_TABLES_LOC, aTableCount, 0 )

            SCAN FOR EXPORT = .T.
                * For thermometer
                lcMessage = STRTRAN( THIS_TABLE_LOC, "|1", RTRIM( &lcEnumTablesTbl..RmtTblName ))
                this.UpDateTherm( lnTableCount, lcMessage )
                lnTableCount = lnTableCount + 1

                * If we're Oracle and table is part of cluster, make sure cluster was created
                IF this.ServerType = "Oracle" AND !EMPTY( ClustName )
                    IF !this.ClusterExported( RTRIM( ClustName ))
                        * If cluster wasn't created, skip the table, log error message
                        lcMsg = STRTRAN( CLUST_NOT_CREATED_LOC, '|1', RTRIM( ClustName ))
                        REPLACE Exported WITH .F., TblError WITH lcMsg ADDITIVE
                        lnTableCount = lnTableCount + 1
                        LOOP
                    ENDIF
                ENDIF

                *Check to see if table already exists
                lcRmtTable = RTRIM( &lcEnumTablesTbl..RmtTblName )
                IF this.TableExists( lcRmtTable ) THEN
                    IF !this.OVERWRITE THEN
                        this.OverWriteOK( RTRIM( &lcEnumTablesTbl..RmtTblName ), "Table" )
                        DO CASE
                            CASE this.UserInput = '3'
                                * user says leave it alone ( 'NO' ) THEN
                                lcMsg = TABLE_NOT_CREATED_LOC
                                REPLACE Exported WITH .F., TblError WITH lcMsg ADDITIVE
                                lnTableCount = lnTableCount + 1
                                LOOP
                            CASE this.UserInput = '2'
                                * user says overwrite all
                                this.OVERWRITE = .T.
                                *CASE ELSE
                                * just keep going
                        ENDCASE
                        this.UserInput = ""
                    ENDIF
                    * Drop the table; skip the table if it can't be dropped for some reason
                    llRetVal = this.DropTable( &lcEnumTablesTbl..RmtTblName )
                    IF !llRetVal THEN
                        REPLACE Exported WITH .F., TblError WITH CANT_DROP2_LOC ADDITIVE
                        lnTableCount = lnTableCount+1
                        LOOP
                    ENDIF
                ENDIF

                lcSQL = TableSQL
                llRetVal = this.ExecuteTempSPT( lcSQL, @lnError, @lcErrMsg )
                IF llRetVal THEN
                    REPLACE Exported WITH .T.
                ELSE
                    this.StoreError( lnError, lcErrMsg, lcSQL, CANT_CREATE_LOC, lcRmtTable, TABLE_LOC )
                    REPLACE Exported WITH .F., TblError WITH CANT_CREATE_LOC ADDITIVE, ;
                        TblErrNo WITH lnError
                ENDIF
            ENDSCAN
			raiseevent( This, 'CompleteProcess' )
        ENDIF



    FUNCTION CreateTableSQL
        LOCAL lcEnumTablesTbl, lcEnumFieldsTbl, lnOldArea, lcTableName, llTStamp, ;
            lnTableCount, lcCRLF, lcTimeStamp, I, lcExact
        LOCAL llAddedIdentity

        lcCR=CHR( 10 )
        lcEnumTablesTbl = this.EnumTablesTbl
        lcEnumFieldsTbl = this.EnumFieldsTbl
        lnOldArea = SELECT()
        SELECT ( lcEnumTablesTbl )

        *Update thermometer
        SELECT COUNT( * ) FROM ( lcEnumTablesTbl ) WHERE EXPORT=.T. INTO ARRAY aTableCount
        this.InitTherm( TABLE_SQL_LOC, aTableCount, 0 )
        lnTableCount=0

        *SQL Server bit fields don't support NULLs
        SELECT ( lcEnumFieldsTbl )
        REPLACE RmtNull WITH .F. FOR RmtType="bit"

* If we're supposed to replace blank date values with nulls, ensure D and T
* fields support nulls.

		IF isnull( this.BlankDateValue )
			replace RMTNULL with .T. for inlist( DATATYPE, 'D', 'T' )
		ENDIF

        * Handle Null overrides
        DO CASE
            CASE this.NullOverride=1	&&general only
                REPLACE RmtNull WITH .T. FOR RmtType#"bit" AND INLIST( DATATYPE, "G" )
            CASE this.NullOverride=2	&&general and memo only
                REPLACE RmtNull WITH .T. FOR RmtType#"bit" AND INLIST( DATATYPE, "G", "M" )
            CASE this.NullOverride=3	&&all fields
                REPLACE RmtNull WITH .T. FOR RmtType#"bit"
        ENDCASE

        *Make sure if we add timestamp that we don't get duplicate field names
        lcTimeStamp=LOWER( TIMESTAMP_LOC )
        lcExact=SET( "EXACT" )
        SET EXACT ON
        LOCATE FOR FldName=( lcTimeStamp )
        I=1
        DO WHILE LEN( lcTimeStamp )<MAX_NAME_LENGTH
            IF EOF()
                EXIT
            ENDIF
            lcTimeStamp=TIMESTAMP_LOC+LTRIM( STR( I ))
            LOCATE FOR FldName=( lcTimeStamp )
            I=I+1
        ENDDO
        SET EXACT &lcExact
        this.TimeStampName=lcTimeStamp

        *Make sure if we add identity column that we don't get duplicate field names
        lcIdentityCol = LOWER( IDENTCOL_LOC )
        lcExact=SET( "EXACT" )
        SET EXACT ON
        LOCATE FOR FldName=( lcIdentityCol )
        I=1
        DO WHILE LEN( lcIdentityCol )<MAX_NAME_LENGTH
            IF EOF()
                EXIT
            ENDIF
            lcIdentityCol=IDENTCOL_LOC+LTRIM( STR( I ))
            LOCATE FOR FldName=( lcIdentityCol )
            I=I+1
        ENDDO
        SET EXACT &lcExact
        this.IdentityName=lcIdentityCol


        SELECT ( lcEnumTablesTbl )
        SCAN FOR EXPORT=.T.
			llAddedIdentity = .F. && set to true if we add an identity column

            lcTableName=RTRIM( &lcEnumTablesTbl..TblName )
            lcMsg=STRTRAN( THIS_TABLE_LOC, '|1', lcTableName )
            this.UpDateTherm( lnTableCount, lcMsg )

			*{Change JEI - RKR - 28.03.2005
            *lcCreateString="CREATE TABLE " + RTRIM( &lcEnumTablesTbl..RmtTblName ) +" ( "
			lcCreateString="CREATE TABLE [" + RTRIM( &lcEnumTablesTbl..RmtTblName ) +"] ( "
			*{Change JEI - RKR - 28.03.2005

            SELECT ( lcEnumFieldsTbl )
            SCAN FOR RTRIM( &lcEnumFieldsTbl..TblName )==lcTableName

            	&&{ Change JEI - RKR 28.03.2005 : upsize field names which are reserved words in SQL
				*lcCreateString = lcCreateString + RTRIM( &lcEnumFieldsTbl..RmtFldname )
                lcCreateString = lcCreateString + "[" + RTRIM( &lcEnumFieldsTbl..RmtFldname ) + "]"
            	&&} Change JEI - RKR 28.03.2005

                lcCreateString = lcCreateString + " " + RTRIM( &lcEnumFieldsTbl..RmtType )
                IF &lcEnumFieldsTbl..RmtLength<>0 THEN
                    lcCreateString = lcCreateString + "( " + ALLTRIM( STR( &lcEnumFieldsTbl..RmtLength ))
                    IF this.ServerType="Oracle" OR this.ServerType=="SQL Server95" THEN
                        IF &lcEnumFieldsTbl..RmtPrec<>0 THEN
                            lcCreateString = lcCreateString + ", " + ALLTRIM( STR( &lcEnumFieldsTbl..RmtPrec ))
                        ENDIF
                    ENDIF
                    lcCreateString = lcCreateString + " )"
                ENDIF
                lcCreateString = lcCreateString + " " + IIF( &lcEnumFieldsTbl..RmtNull=.T., "NULL ", "NOT NULL " )
                * JVF 11/02/02 Since VFP 8 has autoinc, need to account for ID column at field level.
                * Add SQL Server Seed and Step VFP Increment and Step provided.
                IF this.SQLServer AND RmtType= "int ( Ident )" AND !llAddedIdentity
                    lcCreateString = STRTRAN( lcCreateString, "( Ident )" ) + ;
                        " IDENTITY( " + ALLTRIM( STR( AutoInNext-1 )) + ", " + ;
                        ALLTRIM( STR( AutoInStep )) + " )"

					llAddedIdentity = .T.
                ENDIF
                lcCreateString = lcCreateString + ", "
            ENDSCAN
            SELECT ( lcEnumTablesTbl )

            * Add timestamp and identity if appropriate
            * Account for choosing All
            * jvf: 08/13/99
            llTStamp = ( this.SQLServer AND ( TStampAdd OR this.TimeStampAll=1 ))

            * rmk - 01/27/2004 - don't add identify column if upsizing AutoInc column
            llIdentity = ( this.SQLServer AND ( IdentAdd OR this.IdentityAll=1 ) AND !llAddedIdentity )

            IF llTStamp
                lcCreateString = lcCreateString + lcTimeStamp + " timestamp, "
            ENDIF

            IF llIdentity
                lcCreateString = lcCreateString + lcIdentityCol + " int IDENTITY( 1, 1 ), "
            ENDIF

            * peel off the extra comma at the end and add a closing parenthesis
            lcCreateString = LEFT( lcCreateString, LEN( RTRIM( lcCreateString )) - 1 ) + " )"

            * Oracle only: add tablespace or cluster name if appropriate
            IF this.ServerType = ORACLE_SERVER
                * clustered tables are created in the tablespace of the cluster index
                IF !EMPTY( &lcEnumTablesTbl..ClustName )
                    lcClustName = TRIM( &lcEnumTablesTbl..ClustName )
                    lcClusterKey = this.CreateClusterKey( lcTableName, .F. )
                    lcCreateString = lcCreateString + " CLUSTER " + lcClustName + " ( " + lcClusterKey + " )"
                ELSE
                    IF !EMPTY( this.TSTableTSName )
                        lcCreateString = lcCreateString + " TABLESPACE " + this.TSTableTSName
                    ENDIF
                ENDIF
            ENDIF

            REPLACE TableSQL WITH lcCreateString
            lnTableCount = lnTableCount + 1

        ENDSCAN

		raiseevent( This, 'CompleteProcess' )
        SELECT ( lnOldArea )




    FUNCTION TableExists
        PARAMETERS lcTableName
        LOCAL dummy, lcSQuote

        *Checks to see if a table of the same name already exists on the server

        dummy='x'
        lcSQuote=CHR( 39 )
        IF this.ServerType="Oracle" THEN
            lcSQL="SELECT TABLE_NAME FROM USER_TABLES WHERE TABLE_NAME=" + lcSQuote+ UPPER( lcTableName ) + lcSQuote
            lcField="TABLE_NAME"
        ELSE
            lcSQL="select uid from sysobjects where uid = user_id() and name =" + lcSQuote + lcTableName + lcSQuote
            lcField="uid"
        ENDIF

        RETURN this.SingleValueSPT( lcSQL, dummy, lcField )




    FUNCTION AddTableToCluster
        PARAMETERS lcClustName, aKeyFields
        LOCAL lnOldArea, lcEnumFieldsTbl, lcTableName, I

        * Adds the cluster name to the table record and the ClustOrder
        * to the fields participating in the key expression
        * The <table> parameter is given by the current record in 'Tables'

        DIMENSION aKeyFields[ALEN( aKeyFields )]
        lnOldArea = SELECT()
        lcEnumFieldsTbl = this.EnumFieldsTbl
        lcTableName = LOWER( TRIM( TblName ))
        REPLACE ClustName WITH lcClustName

        * set the ClusterOrder field for each Key column
        SELECT ( lcEnumFieldsTbl )
        REPLACE ClustOrder WITH 0 FOR TblName == lcTableName
        FOR m.I = 1 TO ALEN( aKeyFields, 1 )
            LOCATE FOR TblName = lcTableName AND FldName = LOWER( aKeyFields[m.i] )
            IF EOF()
                * we should never get here
                LOOP
            ENDIF
            REPLACE ClustOrder WITH m.I
        ENDFOR

        SELECT ( lnOldArea )



    FUNCTION DeleteClusterInfo( lcClusterName )
        LOCAL lnOldArea, lcEnumTablesTbl, lcEnumFieldsTbl, llAll

        * removes <lcClusterName> from its tables in Tables and
        * removes ClustOrder from all corresponding field records in Fields
        * if called with no parameter, clears all cluster info from Tables and Fields

        llAll = ( PARAMETERS() = 0 )
        lnOldArea = SELECT()
        lcEnumTablesTbl = this.EnumTablesTbl
        lcEnumFieldsTbl = this.EnumFieldsTbl

        IF !llAll
            SELECT ( lcEnumTablesTbl )
            SCAN FOR ClustName = lcClusterName
                lcTableName = LOWER( TRIM( TblName ))
                SELECT ( lcEnumFieldsTbl )
                REPLACE ClustOrder WITH 0 FOR TblName == lcTableName
                SELECT ( lcEnumTablesTbl )
                REPLACE ClustName WITH ""
            ENDSCAN
        ELSE
            SELECT ( lcEnumFieldsTbl )
            REPLACE ClustOrder WITH 0 ALL
            SELECT ( lcEnumTablesTbl )
            REPLACE ClustName WITH "" ALL
        ENDIF

        SELECT ( lnOldArea )




    FUNCTION GetDefaultClusters
        LOCAL lcEnumRelsTbl, lnOldArea, lcEnumTablesTbl, lcClustName, ;
            aParentKey, aChildKey, lcChildExpr, myarray, lcMsg
        DIMENSION aParentKey[1], aChildKey[1], myarray[1]

        this.GetRiInfo
        SELECT COUNT( * ) FROM ( this.EnumRelsTbl ) INTO ARRAY myarray
        IF EMPTY( myarray )
        	IF this.lQuiet
        		.HadError     = .T.
        		.ErrorMessage = CANTDEFCLUSTERS_LOC
        	ELSE
	            =MESSAGEBOX( CANTDEFCLUSTERS_LOC, ICON_EXCLAMATION, TITLE_TEXT_LOC )
        	ENDIF
            RETURN .F.
        ELSE
            this.DeleteClusterInfo()

            lnOldArea = SELECT()
            lcEnumRelsTbl = this.EnumRelsTbl
            lcEnumTablesTbl = this.EnumTablesTbl

            * select all parent tables in aParents
            SELECT DISTINCT  DD_PARENT, DD_PAREXPR  FROM ( lcEnumRelsTbl ) INTO ARRAY aParents
            FOR m.I = 1 TO ALEN( aParents, 1 )
                * find next available parent table
                SELECT ( lcEnumTablesTbl )
                LOCATE FOR TblName = LOWER( aParents( m.I, 1 ))
                IF EOF() OR !EMPTY( ClustName )
                    LOOP
                ENDIF

                * get cluster name and parent key fields
                lcClustName = CL_LOC + TRIM( aParents[m.i, 1] )
                aParentKey[1] = ""
                this.KeyArray( aParents[m.i, 2], @aParentKey )

                * add child table to cluster
                this.AddTableToCluster( lcClustName, @aParentKey )

                * find available child tables and add them to the cluster
                SELECT ( lcEnumRelsTbl )
                SCAN FOR DD_PARENT = aParents[m.i, 1] AND DD_PAREXPR = aParents[m.i, 2]
                    lcChild = TRIM( DD_CHILD )
                    lcChildExpr = DD_CHIEXPR
                    SELECT( lcEnumTablesTbl )
                    LOCATE FOR TblName = LOWER( lcChild )
                    IF EOF() OR !EMPTY( ClustName )
                        LOOP
                    ENDIF

                    * get child key fields, compare with parent key fields
                    aChildKey[1] = ""
                    this.KeyArray( lcChildExpr, @aChildKey )
                    IF ALEN( aParentKey ) != ALEN( aChildKey )
                        LOOP
                    ENDIF

                    * add child table to cluster
                    this.AddTableToCluster( lcClustName, @aChildKey )

                    SELECT ( lcEnumRelsTbl )
                ENDSCAN
            ENDFOR

            * initialise the aClusters array
            DIMENSION aClusters[1, 5]
            aClusters[1, 1] = ""
            SELECT DISTINCT ClustName FROM ( lcEnumTablesTbl ) INTO ARRAY aParents
            FOR m.I = 1 TO ALEN( aParents, 1 )
                IF ( m.I = 1 )
                    aClusters[1, 1] = aParents[1]
                    aClusters[1, 2] = "INDEX"
                    aClusters[1, 3] = 0
                    aClusters[1, 4] = 2
                    aClusters[1, 5] = .F.
                ELSE
					this.InsArrayRow( @aClusters, aParents[m.i], "INDEX", 0, 2, .F. )
                ENDIF
            ENDFOR
        ENDIF

        SELECT ( lnOldArea )



    FUNCTION GetClusterKey
        PARAMETERS lcTableName, lcClustName
        LOCAL lcEnumRelsTbl, lnOldArea

        * Returns the key of a table that's in a cluster

        lnOldArea = SELECT()
        lcEnumRelsTbl = this.EnumRelsTbl
        SELECT ( lcEnumRelsTbl )
        LOCATE FOR RTRIM( ClustName ) == lcClustName
        IF RTRIM( DD_PARENT )==lcTableName THEN
            lcClusterKey=DD_PAREXPR
        ELSE
            lcClusterKey=DD_CHIEXPR
        ENDIF

        SELECT ( lnOldArea )
        RETURN ALLTRIM( lcClusterKey )



    FUNCTION AddTimeStamp
        PARAMETERS lcTableName
        LOCAL lcEnumFieldsTbl, aCount[1]

        *This routine returns True if a table has float, real, binary,
        *varbinary, image, or text data types in them

        lcEnumFieldsTbl = this.EnumFieldsTbl

        SELECT COUNT( * ) FROM ( lcEnumFieldsTbl ) ;
            WHERE TblName = lcTableName ;
            AND ( DATATYPE = "M" OR ;
            DATATYPE = "G" OR ;
            DATATYPE = "P" ) ;
            INTO ARRAY aCount

        RETURN aCount[1] > 0



    FUNCTION DropTable
        PARAMETERS lcTable
        LOCAL lcSQL

        IF this.ServerType="Oracle" THEN
            lcSQL="drop table " + RTRIM( lcTable ) + " CASCADE CONSTRAINTS"
        ELSE
            lcSQL="drop table " + RTRIM( this.UserName ) + "." + RTRIM( lcTable )
        ENDIF
        lnRetVal=this.ExecuteTempSPT( lcSQL )
        RETURN lnRetVal



    FUNCTION CreateClusters
        LOCAL lcSQL, lcEnumClustersTbl, llClusterCreated, dummy, lcClusterName, lcThermMsg, ;
            lnClustCount, lnError, lcErrMsg, aClustCount

* If we have an extension object and it has a CreateClusters method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'CreateClusters', 5 ) and ;
			not this.oExtension.CreateClusters( This )
			return
		ENDIF

        *In this routine, the SQL for clusters is created and ( possibly ) executed
        *and cluster indexes are created as appropriate

        *Bail if there aren't any clusters to create
        IF !this.CreateClusterSQL() THEN
            RETURN
        ENDIF

        dummy = "x"
        IF this.DoUpsize THEN
            lcEnumClustersTbl = this.EnumClustersTbl
            SELECT ( lcEnumClustersTbl )
            aClustCount = RECCOUNT()
            this.InitTherm( CREATING_CLUSTERS_LOC, aClustCount, 0 )
            lnClustCount = 0

            SCAN
                *check for existing cluster
                lcClusterName = RTRIM( ClustName )
                lcThermMsg = STRTRAN( THIS_CLUST_LOC, "|1", lcClusterName )
                this.UpDateTherm( lnClustCount, lcThermMsg )
                lnClustCount = lnClustCount + 1
                lcSQL = "SELECT CLUSTER_NAME FROM USER_CLUSTERS WHERE CLUSTER_NAME=" + "'" + UPPER( lcClusterName ) + "'"
                IF this.SingleValueSPT( lcSQL, dummy, "CLUSTER_NAME" ) THEN
                    IF !this.OVERWRITE THEN
                        *Pop up dialog and ask user what they want to do
                        this.OverWriteOK( RTRIM( ClustName ), "Cluster" )
                        DO CASE
                            CASE this.UserInput = '3'
                                *user says leave it alone ( 'NO' ) THEN
                                REPLACE Exported WITH .F., ClustErr WITH CLUST_EXISTS_LOC
                                LOOP
                            CASE this.UserInput = '2'
                                *user says overwrite all
                                this.OVERWRITE = .T.
                                *CASE ELSE
                                *just keep going
                        ENDCASE
                        this.UserInput = ""
                    ENDIF
                    * Drop the cluster and all tables; otherwise you can't know
                    * if the cluster key is going to work for the tables
                    * the wizard will try to add
                    *
                    *If dropping cluster and tables doesn't work, bag the cluster
                    *and the tables

                    IF !this.DropCluster( RTRIM( ClustName ), @lnError, @lcErrMsg )
                        REPLACE Exported WITH .F., ClustErr WITH CANT_DROP_LOC ADDITIVE
                        this.StoreError( lnError, lcErrMsg, lcSQL, CANT_DROP_LOC, lcClusterName, CLUSTER_LOC )
                        LOOP
                    ENDIF
                ENDIF

                lcSQL = &lcEnumClustersTbl..ClusterSQL
                IF !this.ExecuteTempSPT( lcSQL, @lnError, @lcErrMsg ) THEN
                    this.StoreError( lnError, lcErrMsg, lcSQL, CANT_CREATE_CLUST_LOC, lcClusterName, CLUSTER_LOC )
                    REPLACE Exported WITH .F., ClustErr WITH CANT_CREATE_CLUST_LOC ADDITIVE, ;
                        ClustErrNo WITH lnError
                ELSE
                    REPLACE Exported WITH .T.
                ENDIF
            ENDSCAN
			raiseevent( This, 'CompleteProcess' )
        ENDIF

        * Now go create cluster indexes, if any
        this.CreateClusterIndexes



    FUNCTION CreateClusterIndexes
        LOCAL lcEnumRelsTbl, lcEnumIndexesTbl, llRetVal, lnError, lcMessage, lnOldArea

        * If there are clusters, create the index sql now
        *
        * DRI on clusters requires that indexes be created before the DRI is executed
        * On the other hand, indexes can conflict with DRI.  This way the wizard
        * creates cluster indexes, does the DRI ( which resolves potential index-DRI conflicts
        * and then creates non-cluster indexes
        *

        *Create table to hold index names and expressions
        IF RTRIM( this.EnumIndexesTbl ) == ""
            this.EnumIndexesTbl = this.CreateWzTable( "Indexes" )
        ENDIF

        lcEnumClustersTbl = this.EnumClustersTbl
        lcEnumIndexesTbl = this.EnumIndexesTbl
        lnOldArea = SELECT()
        SELECT ( lcEnumClustersTbl )

        SCAN FOR ClustType = "INDEX"
            lcClusterName = RTRIM( ClustName )
            lcTagName = CLUST_INDEX_PREFIX + LEFT( lcClusterName, MAX_NAME_LENGTH-LEN( CLUST_INDEX_PREFIX ))
            lcTagName = this.UniqueOraName( lcTagName )
            lcSQL = "CREATE INDEX [" + lcTagName + "] ON CLUSTER " + lcClusterName

            IF !EMPTY( this.TSIndexTSName )
                lcSQL = lcSQL + " TABLESPACE " + LTRIM( this.TSIndexTSName )
            ENDIF

            IF this.DoUpsize THEN
                llRetVal = this.ExecuteTempSPT( lcSQL, @lnError, @lcMessage )
                IF !llRetVal
                    this.StoreError( lnError, lcMessage, lcSQL, CLUST_IDX_FAILED_LOC, lcTagName, INDEX_LOC )
                ENDIF
            ENDIF

            SELECT ( lcEnumIndexesTbl )
            APPEND BLANK
            REPLACE RmtName WITH lcTagName, ;
                RmtTable WITH lcClusterName, ;
                IndexSQL WITH lcSQL, ;
                Exported WITH llRetVal, ;
                DontCreate WITH .T.

            IF !EMPTY( lcMessage ) THEN
                REPLACE IdxError WITH CLUST_IDX_FAILED_LOC, ;
                    IdxErrNo WITH lnError
            ENDIF

            SELECT ( lcEnumClustersTbl )
        ENDSCAN

        SELECT( lnOldArea )



    FUNCTION UniqueOraName
        PARAMETERS lcObjName, llAddToTable
        LOCAL lcSQL, lnOldArea, lcOraNames, I, lcNewName, lnN, lcM, lcFieldName

        *Returns a name that is not in use by an index, constraint, or trigger
        *Only an issue for Oracle and SQL '95
        *
        *SQL Server 4.x and '95 let you have the same index name as long as they're
        *on different tables; however must have unique name for constraints

        lnOldArea=SELECT()
        IF PARAMETERS()=1
            llAddToTable=.T.
        ENDIF

        * Get the names of all the user's indexes and constraints
        * this cursor is built once, then saved in <this.OraNames>
        IF this.OraNames == "" THEN
            SELECT 0
            this.OraNames = this.UniqueCursorName( "OraIdx" )

            IF this.ServerType="Oracle" THEN
                lcSQL=        "SELECT INDEX_NAME FROM USER_INDEXES "
                lcSQL=lcSQL + "UNION SELECT CONSTRAINT_NAME FROM USER_CONSTRAINTS "
                lcSQL=lcSQL + "UNION SELECT TRIGGER_NAME FROM USER_TRIGGERS"
            ELSE
                lcSQL="SELECT name FROM sysobjects"
            ENDIF

            *If it doesn't work, just send back original name; fatal
            *error is probably imminent anyway
            IF !this.ExecuteTempSPT( lcSQL, lnN, lcM, this.OraNames ) THEN
                this.OraNames=""
                RETURN lcObjName
            ENDIF
        ENDIF

        IF this.ServerType = ORACLE_SERVER THEN
            lcFieldName = "INDEX_NAME"
        ELSE
            lcFieldName = "name"
        ENDIF

        *See if there's a conflict
        lcOraNames = this.OraNames
        SELECT ( lcOraNames )
        lcNewName = lcObjName
        LOCATE FOR LOWER( RTRIM( &lcFieldName )) == LOWER( lcNewName )

        I=1
        DO WHILE FOUND()
            *If there's a duplicate name, fiddle with the new object name til it's unique
            lcNewName=LEFT( lcObjName, MAX_NAME_LENGTH-( LEN( LTRIM( STR( I )) )) ) + LTRIM( STR( I ))
            I=I+1
            LOCATE FOR LOWER( RTRIM( &lcFieldName ))==LOWER( lcNewName )
        ENDDO

        *Add the new name to the list of those in use
        IF llAddToTable THEN
            APPEND BLANK
            REPLACE &lcFieldName WITH lcNewName
        ENDIF

        SELECT ( lnOldArea )
        RETURN lcNewName



    FUNCTION DropCluster
        PARAMETERS lcClusterName, lnErrNo, lcErrMsg

        *Drop cluster, its tables, and any RI on the tables
        lcSQL="DROP CLUSTER " + RTRIM( lcClusterName )+ " INCLUDING TABLES CASCADE CONSTRAINTS"
        RETURN this.ExecuteTempSPT( lcSQL, @lnErrNo, @lcErrMsg )



    FUNCTION ClusterExported
        PARAMETERS lcClustName
        LOCAL lcEnumClustersTbl, lnOldArea, llExported

        * Checks to see if the cluster was actually created

        lcEnumClustersTbl = this.EnumClustersTbl
        lnOldArea = SELECT()
        SELECT ( lcEnumClustersTbl )
        LOCATE FOR RTRIM( ClustName ) == lcClustName
        llExported = Exported
        SELECT( lnOldArea )
        RETURN llExported



    FUNCTION HashDefault
        PARAMETERS lcRel
        LOCAL lcParent, lcChild, lnDupID, lnHashKeys, lnOldArea

        *Don't change this ratio w/o letting UE know
        #DEFINE HASH_RATIO 			1.1

        #DEFINE HASHKEYS_FLOOR		500
        #DEFINE HASHKEYS_CEILING	2147483647

        *
        *Derives a number for default hashkey value
        *

        *Figures out how many records are in the parent table, multiplies it by some number

        this.ParseRel ( lcRel, @lcParent, @lcChild, @lnDupID )

        lnOldArea=SELECT()
        SELECT ( lcParent )

        lnHashKeys=RECCOUNT()*HASH_RATIO
        lnHashKeys=IIF( lnHashKeys>HASHKEYS_FLOOR, lnHashKeys, HASHKEYS_FLOOR )
        lnHashKeys=IIF( lnHashKeys>HASHKEYS_CEILING, HASHKEYS_CEILING, lnHashKeys )
        SELECT ( lnOldArea )

        RETURN INT( lnHashKeys )



    FUNCTION OverWriteOK
        PARAMETERS lcObjectName, lcObjectType

        LOCAL aButtonNames, lcMessageText, MyMessageBox

        *display message box displaying Yes, Yes to all, No buttons
        IF lcObjectType="Table" THEN
            lcMessageText=STRTRAN( TABLE_EXISTS_LOC, "|1", lcObjectName )
        ELSE
            lcMessageText=STRTRAN( CLUSTER_EXISTS_LOC, "|1", lcObjectName )
        ENDIF

        DIMENSION aButtonNames[3]
        aButtonNames[1]=YES_LOC
        aButtonNames[2]=YESALL_LOC
        aButtonNames[3]=NO_LOC
        MyMessageBox = newobject( 'MessageBox2', 'VFPCtrls.vcx', '', ;
        	lcMessageText, @aButtonNames )
        MyMessageBox.Show()
        IF vartype( MyMessageBox ) = 'O'
        	this.UserInput = MyMessageBox.cChoice
        	MyMessageBox.Release()
        ENDIF



    FUNCTION CreateClusterSQL
        LOCAL lcClustName, lcPkey, aClusterTables, lcEnumClustersTbl, lnOldArea
        DIMENSION aClusterTables[1]

        *If the clusters table doesn't exist yet, there won't be any clusters
        IF EMPTY( this.EnumClustersTbl )
            RETURN .F.
        ENDIF

        * needs to be of the form:
        * CREATE CLUSTER <clustername> ( <key_name> <datatype>, etc. )

        lnOldArea = SELECT()
        lcEnumClustersTbl = this.EnumClustersTbl
        SELECT ( lcEnumClustersTbl )
        SCAN
            lcClustName = RTRIM( ClustName )
            this.GetEligibleClusterTables( @aClusterTables, lcClustName )
            lcPkey = this.CreateClusterKey( aClusterTables[1], .T. )

            lcSQL = "CREATE CLUSTER " + lcClustName + " ( " + lcPkey + " )"

            IF !EMPTY( this.TSTableTSName )
                lcSQL = lcSQL + " TABLESPACE " + LTRIM( this.TSTableTSName )
            ENDIF

            IF ClustType = RTRIM( "HASH" )
                lcSQL = lcSQL + " HASHKEYS " + LTRIM( STR( HashKeys ))
            ENDIF

            IF !EMPTY( ClustSize )
                lcSQL = lcSQL + " SIZE " + LTRIM( STR( ClustSize )) + " K"
            ENDIF

            REPLACE ClusterSQL WITH lcSQL
        ENDSCAN
        SELECT( lnOldArea )

        RETURN .T.



    FUNCTION CreateClusterKey
        PARAMETERS lcTableName, llDataType

        * Creates the Cluster Key with data types for CreateClustersSQL
        * Creates the Cluster Key list for CreateTablesSQL

        LOCAL aKeyArray, lcClustKey, I
        DIMENSION aKeyArray[1, 4]
        this.GetInfoTableFields( @aKeyArray, lcTableName, .T. )
        lcClustKey = ""

        FOR m.I = 1 TO ALEN( aKeyArray, 1 )
            lcClustKey = lcClustKey + RTRIM( aKeyArray[m.i, 1] )

            IF llDataType
                lcClustKey = lcClustKey + " " + RTRIM( aKeyArray[m.i, 2] )
                IF !EMPTY( aKeyArray[m.i, 3] )
                    lcClustKey = lcClustKey + "( " + ALLTRIM( STR( aKeyArray[m.i, 3] ))
                    IF !EMPTY( aKeyArray[m.i, 4] )
                        lcClustKey = lcClustKey + ", " + ALLTRIM( STR( aKeyArray[m.i, 4] ))
                    ENDIF
                    lcClustKey = lcClustKey + " )"
                ENDIF
            ENDIF

            lcClustKey = lcClustKey + ", "
        ENDFOR

        *peel off the extra comma at the end
        lcClustKey = LEFT( lcClustKey, LEN( lcClustKey )-2 )
        RETURN lcClustKey



    FUNCTION SendData
        LOCAL lcOldArea, lcEnumTables, lcSprocSQL, lnBigRows, ;
            lnSmallRows,  lnBigBlocks, lnSmallBlocks, lnExportErrors, ;
            ll255, llMaxErrExceeded, lcErrMsg, lcCursorName, lcErrTblName, lcExportType, lReturnVal
		local llDone

        IF !this.DoUpsize OR this.ExportStructureOnly THEN
            RETURN
        ENDIF

* If we have an extension object and it has a SendData method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'SendData', 5 ) and ;
			not this.oExtension.SendData( This )
			return
		ENDIF

        lcOldArea=SELECT()
        lcErrMsg=""
        lcEnumTables=this.EnumTablesTbl
        SELECT ( lcEnumTables )

        *Don't send deleted records
        lcDelStatus=SET( "DELETED" )
        SET DELETED ON

        SCAN FOR &lcEnumTables..EXPORT=.T. AND Exported=.T.

            lcExportType = ""  && can resolve to BULKINSERT, FASTEXPORT, or JIMEXPORT
            lcTablePath=&lcEnumTables..TblPath
            lcRmtTableName=RTRIM( &lcEnumTables..RmtTblName )
            lcTableName=RTRIM( &lcEnumTables..TblName )
            lcCursorName=RTRIM( &lcEnumTables..CursName )

            lcEnumFieldsTbl=RTRIM( this.EnumFieldsTbl )
            IF this.ServerType="Oracle" THEN
                *this query eliminates long, raw, and long raw data types
                SELECT COUNT( FldName ) FROM  &lcEnumFieldsTbl ;
                    WHERE ( RmtType="raw" OR RmtType="long" ) ;
                    AND RTRIM( TblName )==lcTableName INTO ARRAY aExportType
                IF !EMPTY( aExportType )
                    lcExportType = "JIMEXPORT"
                ENDIF
            ELSE
                lReturnVal = SQLEXEC( this.MasterConnHand, "SET IDENTITY_INSERT " + lcRmtTableName + " ON" )
            ENDIF

            *	Text and image datatypes won't work with BULKINSERT OR FASTEXPORT
            *	( can't pass these as parameters to stored procedures, nor can they
            *	reside in text files to be used as source of SQL>=7 bulk insert.
            IF EMPTY( lcExportType )
                SELECT COUNT( FldName ) FROM &lcEnumFieldsTbl ;
                    WHERE ( RmtType="text" OR RmtType="image" ) ;
                    AND RTRIM( TblName )==lcTableName INTO ARRAY aExportType
                IF !EMPTY( aExportType ) OR !this.ServerISLocal
                    lcExportType = "JIMEXPORT"
                ENDIF
            ENDIF

            IF EMPTY( lcExportType ) AND this.ServerType#"Oracle" AND this.ServerVer>=7
                * jvf Determine if we can use SQL >=7's BULK INSERT function:
                * 	- Can't bulk insert tables with NULL columns of non-character
                * 		datatype b/c COPY TO loses the null value by creating
                * 		empty values in the text file ( eg, /  /, F ).

                *{ JEI RKR 23.11.2005 Change
*!*	                SELECT COUNT( FldName ) FROM  &lcEnumFieldsTbl ;
*!*	                    WHERE ( RmtType <> "char" AND lclnull ) ;
*!*	                    AND RTRIM( TblName )==lcTableName INTO ARRAY aExportType
                SELECT COUNT( FldName ) FROM  &lcEnumFieldsTbl ;
                    WHERE lclnull ;
                    AND RTRIM( TblName )==lcTableName INTO ARRAY aExportType
                *} JEI RKR 23.11.2005
					do case
						case empty( aExportType )
						case not this.NotUseBulkInsert and ;
							not this.CheckForNullValuesInTable( lcTableName )
							lcExportType = "BULKINSERT"
						case this.NotUseBulkInsert
							lcExportType = "JIMEXPORT"
						otherwise
							lcExportType = "FASTEXPORT"
					endcase
            ENDIF

            * If survived all conditions above, we can go fast
            IF EMPTY( lcExportType )
                IF this.ServerType#"Oracle" AND this.ServerVer>=7
                	*( JEI RKR 23.11.2005 Change
                	IF this.NotUseBulkInsert = .t.
                		lcExportType = "FASTEXPORT"
                	ELSE
						IF this.NotUseBulkInsert or ;
							this.CheckForNullValuesInTable( lcTableName )
                			lcExportType = "FASTEXPORT"
                		ELSE
		                    lcExportType = "BULKINSERT"  && specific to SQL vers >= 7
		                ENDIF
					ENDIF
				    *lcExportType = "BULKINSERT"  && specific to SQL vers >= 7
				    *} JEI RKR 23.11.2005 Change
                ELSE
                    lcExportType = "FASTEXPORT"
                ENDIF
            ENDIF

            * If the table has 255 fields, we have to use JimExport
            * Stored procs can only pass 255 parameters and one of them
            * is used for something other than data leaving 254.
            * jvf 1/9/01 SQL >= 7 can handle 1024 SP Parameters. Handled in Case below.
            ll255=( MAX_PARAMS=this.CountFields( lcCursorName ))

* If we're upsizing to SQL Server 2005 or later, try to do a bulk XML import
* since it's way faster.

			IF this.SQLServer and this.ServerVer >= 9
				llDone = this.BulkXMLLoad( lcTableName, lcCursorName, ;
					lcRmtTableName )
				lnExportErrors = iif( llDone, 0, 1 )
			ENDIF

* If we didn't use bulk XML load, or if it failed, use one of the other
* mechanisms.

            DO CASE
				case llDone
                CASE lcExportType = "BULKINSERT" AND this.Perm_Database
                    * jvf Use SQL 7's BULK INSERT technique
                    lnExportErrors=this.BulkInsert( lcTableName, lcCursorName, ;
                        lcRmtTableName, @llMaxErrExceeded )
                    IF ( lnExportErrors == -1 ) THEN
                    	RETURN
                    ENDIF
                    * jvf SQL >=7 can handle 1024 SP Parameters.
                CASE lcExportType = "FASTEXPORT" AND this.Perm_Sproc AND ;
                        ( !ll255 OR ( this.ServerType="Oracle" OR this.ServerVer<7 )) AND !TStampAdd
                    *go fast if possible and user can create sprocs
                    lnExportErrors=this.FastExport( lcTableName, lcCursorName, lcRmtTableName, @llMaxErrExceeded )
                OTHERWISE
*** 11/20/2012: pass .T. to JimExport so it updates the thermometer
***                    lnExportErrors=this.JimExport( lcTableName, lcCursorName, lcRmtTableName, @llMaxErrExceeded )
                    lnExportErrors=this.JimExport( lcTableName, lcCursorName, lcRmtTableName, @llMaxErrExceeded, '', .T. )
            ENDCASE

            IF lnExportErrors<>0 THEN
                lcErrorTable=ERR_TBL_PREFIX_LOC + ;
                    LEFT( lcTableName, MAX_NAME_LENGTH-LEN( ERR_TBL_PREFIX_LOC ))
            ELSE
                lcErrorTable=""
            ENDIF

            IF llMaxErrExceeded THEN
                lcErrMsg=DEXPORT_FAILED_LOC
                this.StoreError( .NULL., "", "", lcErrMsg, lcTableName, DATA_EXPORT_LOC )
            ELSE
                IF lnExportErrors<>0 THEN
                    lcErrMsg=STRTRAN( SOME_ERRS_LOC, "|1", LTRIM( STR( lnExportErrors )) )
                    this.StoreError( .NULL., "", "", lcErrMsg, lcTableName, DATA_EXPORT_LOC )
                ENDIF
            ENDIF

            IF this.NormalShutdown=.F.
                EXIT
            ENDIF

            *For report
            REPLACE &lcEnumTables..DataErrs WITH lnExportErrors, ;
                &lcEnumTables..DataErrMsg WITH lcErrMsg, ;
                &lcEnumTables..ErrTable WITH lcErrorTable

            lnExportErrors=0
            llMaxErrExceeded=.F.
            lcErrMsg=""

            SQLEXEC( this.MasterConnHand, "SET IDENTITY_INSERT " + lcRmtTableName + " OFF" )
        ENDSCAN

        SET DELETED &lcDelStatus
        SELECT ( lcOldArea )



    FUNCTION CountFields
        PARAMETERS lcCursorName
        LOCAL lnFldCount, lnOldArea

        lnOldArea=SELECT()
        SELECT ( lcCursorName )
        lnFldCount=AFIELDS( zoo )
        SELECT ( lnOldArea )

        RETURN lnFldCount


* SQL Server 2005 ( or later ) bulk XML import. We need to get the server, so
* either parse it out of the connection string or get it from the DSNs entry
* in the Registry. Parse the user name and password from the connection
* string, then call BulkXMLLoad.PRG.

	FUNCTION BulkXMLLoad( tcTableName, tcCursorName, tcRmtTableName )
		local lcServer, ;
			loRegistry, ;
			laValues[1], ;
			lnPos, ;
			lcUserName, ;
			lcPassword, ;
			lcMsg, ;
			llReturn
		IF empty( this.DataSourceName )
			lcServer = strextract( this.ConnectString, 'server=', ';', 1, 3 )
		ELSE
			loRegistry = newobject( 'ODBCReg', home() + 'ffc\registry.vcx' )
			loRegistry.EnumODBCData( @laValues, this.DataSourceName )
			lnPos = ascan( laValues, 'Server', -1, -1, 1, 15 )
			IF lnPos > 0
				lcServer = laValues[lnPos, 2]
			ENDIF
		ENDIF
		IF not 'TRUSTED_CONNECTION=YES' $ upper( this.ConnectString ) and ;
			not 'INTEGRATED SECURITY=SSPI' $ upper( this.ConnectString )
			lcUserName = strextract( this.ConnectString, 'uid=', ';', 1, 3 )
			lcPassword = strextract( this.ConnectString, 'pwd=', ';', 1, 3 )
		ENDIF
		this.InitTherm( SEND_DATA_LOC, 100 )
		lcMsg = strtran( THIS_TABLE_LOC, '|1', tcTableName )
		this.UpDateTherm( 0, lcMsg )
		llReturn = empty( BulkXMLLoad( tcCursorName, tcRmtTableName, ;
			this.BlankDateValue, this.ServerDBName, lcServer, lcUserName, ;
			lcPassword ))
		IF llReturn
			raiseevent( This, 'CompleteProcess' )
		ENDIF
		return llReturn


    FUNCTION JimExport
        *Thanks Jim Lewallen for this code ( or the important and bug-free parts of it anyway )

*** 11/20/2012: added tlUpdateProgress parameter
***        PARAMETERS lcTableName, lcCursorName, lcRmtTableName, llMaxErrExceeded, tcDataErrTable &&  JEI RKR 24.11.2005 Add tcDataErrTable
        PARAMETERS lcTableName, lcCursorName, lcRmtTableName, llMaxErrExceeded, tcDataErrTable, tlUpdateProgress

        LOCAL lnOldArea, lnFieldCount, lcInsertString, lcInsertFinal, llRetVal, lnRecs, ;
            lnRecordsCopied, lcMsg, I, lnExportErrors, lcDataErrTable, lnMaxErrors, ;
            aTblFields, lcSQLErrMsg, lnSQLErrno

		*{ JEI RKR 24.11.2005 Add
		LOCAL lcEmptyDateValue, lcEmptyDateValueEnd as String
		*} JEI RKR 24.11.2005 Add
		*{ JEI RKR 05.01.2006 Add
		LOCAL lnNumberOfPar as Integer
		lnNumberOfPar = PARAMETERS()
		IF lnNumberOfPar > 4
			lcDataErrTable = tcDataErrTable
		ENDIF
		*} JEI RKR 05.01.2006 Add

        #DEFINE STAGGER_COUNT		5

        *Create array of local field names and their remote equivalents
        lcEnumFields=this.EnumFieldsTbl
        DIMENSION aTblFields[1]
        *{  JEI RKR Add field DataType, RmtNull
		SELECT FldName, RmtFldname, DataType, RmtNull, RmtType FROM ( lcEnumFieldsTbl ) WHERE RTRIM( TblName )==lcTableName ;
		    INTO ARRAY aTblFields

        lnOldArea=SELECT()
        SELECT ( lcCursorName )
        lcRemoteName=this.RemotizeName( lcTableName )

        *Thermometer stuff
        lnRecs=RECCOUNT()
        lnRecordsCopied=0
*** 11/20/2012: wrap this code in IF so we only update the progress if we're supposed to
        IF tlUpdateProgress
	        this.InitTherm( SEND_DATA_LOC, lnRecs, 0 )
	        lcMsg=STRTRAN( THIS_TABLE_LOC, '|1', lcTableName )
	        this.UpDateTherm( lnRecordsCopied, lcMsg )
        ENDIF

        *Use the remote field name here
        lcInsertString = 'INSERT INTO ['+ALLTRIM( lcRmtTableName ) + '] ( '
        lnFieldCount = ALEN( aTblFields, 1 )
        FOR ii = 1 TO lnFieldCount
            lcInsertString = lcInsertString + "[" + ALLTRIM( aTblFields[ii, 2] ) + "]" ;
                + IIF( ii < lnFieldCount, ', ', ' )' )
        ENDFOR

        lcInsertString = lcInsertString + ' VALUES ( '
        *{ JEI RKR 24.11.2005 Add

        lcEmptyDateValue = ""
        lcEmptyDateValueEnd = ""
        *} JEI RKR 24.11.2005 Add
        *Use the local field name here
        lcInsertFinal = lcInsertString
        FOR ii = 1 TO lnFieldCount
            SELECT ( lcCursorName )
            *{ JEI RKR 24.11.2005 Add
			do case
				case upper( alltrim( aTblFields[ii, 3] )) $ 'DT' and aTblFields[ii, 4]
					lcEmptyDateValue = "IIF( EMPTY( " + ALLT( lcCursorName )+'.'+ALLT( aTblFields[ii, 1] )+ " ), " + ;
						transform( this.BlankDateValue ) + ", "
        			lcEmptyDateValueEnd = " )"
				case upper( alltrim( aTblFields[ii, 3] )) $ 'YBFIN'
        			IF aTblFields[ii, 4] = .t. && Allow Null
        				lcEmptyDateValue = "IIF( [*] $  TRANSFORM( " + ALLT( lcCursorName )+'.'+ALLT( aTblFields[ii, 1] ) + " ), .Null., "
        			ELSE
        				lcEmptyDateValue = "IIF( [*] $  TRANSFORM( " + ALLT( lcCursorName )+'.'+ALLT( aTblFields[ii, 1] ) + " ), 0, "
        			ENDIF
        			lcEmptyDateValueEnd = " )"
				case lower( aTblFields[ii, 5] ) = 'varchar'
			    	lcEmptyDateValue = 'trim( '
					lcEmptyDateValueEnd = ' )'
			endcase
            *{ JEI RKR 24.11.2005 Add
            lcInsertFinal = lcInsertFinal + RMT_OPERATOR + lcEmptyDateValue + ALLT( lcCursorName ) ;
                + '.'+ALLT( aTblFields[ii, 1] )  + lcEmptyDateValueEnd ;
                + IIF( ii < lnFieldCount, ' , ', ' )' )

            *{ JEI RKR 24.11.2005 Add
          	lcEmptyDateValue = ""
	        lcEmptyDateValueEnd = ""
	        *} JEI RKR 24.11.2005 Add
        ENDFOR

        *Set maximum number of errors allowed so user's disk doesn't fill up if
        *something goes wrong over and over
        lnMaxErrors=lnRecs*DATA_ERR_FRACTION
        IF lnMaxErrors<DATA_ERR_MIN THEN
            lnMaxErrors=DATA_ERR_MIN
        ENDIF

        I=0
        lnExportErrors=0

        *undone: need to set number of inserts per batch

        SCAN
            IF !this.ExecuteTempSPT( lcInsertFinal, @lnSQLErrno, @lcSQLErrMsg ) THEN
                COPY TO ARRAY aErrData NEXT 1
                this.DataExportErr( @aErrData, lcTableName, @lcDataErrTable, lnSQLErrno, lcSQLErrMsg )
                lnExportErrors=lnExportErrors+1
            ENDIF

            IF lnExportErrors>lnMaxErrors THEN
*** 11/20/2012: wrap this code in IF so we only update the progress if we're supposed to
		        IF tlUpdateProgress
        	        this.UpDateTherm( lnRecordsCopied, CANCELED_LOC )
		        ENDIF
                llMaxErrExceeded=.T.
                EXIT
            ENDIF

            IF I=STAGGER_COUNT THEN
                lnRecordsCopied=lnRecordsCopied+STAGGER_COUNT
*** 11/20/2012: wrap this code in IF so we only update the progress if we're supposed to
		        IF tlUpdateProgress
	                IF lnExportErrors<>0 THEN
	                    this.UpDateTherm( lnRecordsCopied, lcMsg + ", " + LTRIM( STR( lnExportErrors ))+ " " + ERROR_COUNT_LOC )
	                ELSE
	                    this.UpDateTherm( lnRecordsCopied, lcMsg )
	                ENDIF
		        ENDIF
                I=1
            ELSE
                I=I+1
            ENDIF
        ENDSCAN

*** 11/20/2012: wrap this code in IF so we only update the progress if we're supposed to
***        IF !llMaxErrExceeded THEN
        IF !llMaxErrExceeded and tlUpdateProgress
			raiseevent( This, 'CompleteProcess' )
        ENDIF

        *Close the errors table if it exists
        IF lnExportErrors<>0  AND lnNumberOfPar < 5 THEN && JEI RKR 06.01.2006 Add lnNumberOfPar < 5
            SELECT ( lcDataErrTable )
            USE
        ENDIF
        SELECT ( lnOldArea )

		tcDataErrTable = lcDataErrTable

        RETURN lnExportErrors



    FUNCTION DataExportErr
        PARAMETERS aErrData, lcTableName, lcDataErrTable, lnSQLErrno, lcSQLErrMsg
        LOCAL lnAlen, lnArrayPos, lnOldArea, lcExact

        *If record( s ) is not exported successfully, it's placed in an error table

        lnOldArea=SELECT()

        IF EMPTY( lcDataErrTable ) THEN
            IF !this.DataErrors THEN
                DIMENSION this.aDataErrTbls [1, 3]
                aDatErrTbls=.F.
                this.DataErrors=.T.
            ELSE
                DIMENSION this.aDataErrTbls [ALEN( this.aDataErrTbls, 1 )+1, 3]
            ENDIF
            lnAlen=ALEN( this.aDataErrTbls, 1 )
            this.aDataErrTbls[lnAlen, 1]=lcTableName
            lcDataErrTable=this.UniqueTableName( ERRT_LOC + LTRIM( STR( lnAlen )) )
            this.aDataErrTbls[lnAlen, 2]=lcDataErrTable

            *Create table with same structure
           	*{ JEI RKR 24.11.2005 Change
           	LOCAL lcTempTableName as String
           	lcTempTableName = this.UniqueTableName( ERRT_LOC + "TableStructure" )
           	COPY STRUCTURE EXTENDED TO &lcTempTableName
            SELECT 0
            USE ( lcTempTableName )
            Replace FIELD_NEXT WITH 0, ;
					FIELD_STEP WITH 0 ALL IN ( ALIAS() )

			USE
			CREATE &lcDataErrTable. FROM &lcTempTableName.
*** DH 07/02/2013: ensure DBF extension is included or the file isn't erased.
*			ERASE ( lcTempTableName )
			ERASE ( forceext( lcTempTableName, 'dbf' ))
			ERASE ( FORCEEXT( lcTempTableName, "fpt" ))
*            COPY STRUCTURE TO &lcDataErrTable
            *{ JEI RKR 24.11.2005 Change
            USE ( lcDataErrTable ) EXCLUSIVE

            *May not be able to store the error if the table has 255 fields already
            IF this.CountFields( lcDataErrTable ) < MAX_FIELDS THEN
                this.aDataErrTbls[lnAlen, 3]=.T.		&&this column says whether it's <255 or not
                ALTER TABLE ( lcDataErrTable ) ADD COLUMN SQLErrNo N( 10, 0 )
                ALTER TABLE ( lcDataErrTable ) ADD COLUMN SQLErrMsg M
            ENDIF
        ELSE
            SELECT ( lcDataErrTable )
        ENDIF

        *If a field has numeric overflow condition, can get an error here; ignore it
        this.SetErrorOff=.T.
        APPEND FROM ARRAY aErrData
        this.SetErrorOff=.F.

        lcExact=SET( 'EXACT' )
        SET EXACT ON
        IF this.aDataErrTbls[ASCAN( this.aDataErrTbls, lcDataErrTable )+1] THEN
            REPLACE SQLErrNo WITH lnSQLErrno, SQLErrMsg WITH lcSQLErrMsg
        ENDIF
        SET EXACT &lcExact

        SELECT ( lnOldArea )

        *The data export routine will close the error table, so it's not closed here




    FUNCTION FastExport
        PARAMETERS lcTableName, lcCursorName, lcRmtTableName, llMaxErrExceeded
        LOCAL lnBigRows, lnSmallRows, lnBigBlocks, lnSmallBlocks, llTStamp, lnExportErrors

        *
        *This thing creates stored procedures that slam data in batches
        *( i.e. the sproc has multiple insert statements
        *

        *lnBigRows is the number of insert statements to put into the sproc; this
        *depends on how many fields are in the table
        *
        *lnBigBlocks is the number of times to execute the sproc
        *
        *lnSmallRows is basically the number of rows left after the sproc has
        *been executed lnBigBlocks number of times
        *
        *lnSmallBlocks is either 0 ( if lnBigRows divides evenly into the number of
        *records in a table ) or 1
        *
        lnSmallRows=0
        lnBigBlocks=0
        lnBigRows=0
        lnSmallBlocks=0

        *See if the table has a timestamp field
        *( Note: this proc is getting called from within a SCAN/ENDSCAN so we
        *know we're on the right record already
        llTStamp=TStampAdd

        *Figure out how many rows to send at a time
        this.RowHeuristics( lcCursorName, @lnBigRows, ;
            @lnSmallRows, @lnBigBlocks, @lnSmallBlocks )

        *Create stored procedures that do the inserts
        this.CreateSQLServerSproc( lcTableName, lcRmtTableName, lnBigRows, lnSmallRows, llTStamp )

        *Execute the stored procedures
        lnExportErrors=this.ExecuteSproc( lcTableName, lcCursorName, lcRmtTableName, lnBigBlocks, ;
            lnSmallBlocks, lnBigRows, lnSmallRows, llTStamp, @llMaxErrExceeded )

        RETURN lnExportErrors



    FUNCTION BulkInsert
        * Assuption: We won't get here if this table has image, text, or non-char nulls
        * columns. Also assume we are upsizing to SQL 7+.
        LPARAMETERS tcTableName, tcCursorName, tcRmtTableName, tlMaxErrExceeded
*** DH 11/20/2013: added llRetVal
*        LOCAL lnServerError, lcErrMsg, lcBulkInsertString, lcBulkFileNameWithPath, lnUserDecision
        LOCAL lnServerError, lcErrMsg, lcBulkInsertString, lcBulkFileNameWithPath, lnUserDecision, llRetVal

        * See if the table has a timestamp and identity field
        * ( Note: this proc is getting called from within a SCAN/ENDSCAN so we
        * know we're on the right record already.

        lnServerError = 0
        lcErrMsg = ""
*** DH 11/20/2013: removed duplicate assignment
*        lnServerError = 0

        * 06/23/01 jvf Account for bulk inserting into a remote machine. Use UNC to specify source path.
        * lcBulkFileNameWithPath = ;
        *	"\\"+ ADDBS( ALLTRIM( LEFT( sys( 0 ), AT( "#", sys( 0 ))-1 )) )+StrTran( SYS( 5 ), ":", "$" ) + ;
        *	CURDIR() + BULK_INSERT_FILENAME

        * JVF 10/30/02 Issue 16653: Invalid path of filename on COPY TO ( tcTargetTextFile ).
        * On 6/23/01 when I tried to fix the bug that occurs on COPY TO when your default directory is a network drive
        * I made the assumption that the network drive will be a share. The following much simpler code seems to work in
        * all situations, namingly: default dir is local, default dir is a mapped network drive, default dir is an unmapped
        * network drive ( UNC ). ( 6/23/01 code commented out. )

        lcBulkFileNameWithPath = ADDBS( FULLPATH( CURDIR() )) + BULK_INSERT_FILENAME

        * Clean up possible previous bulkins.out file.
        IF FILE( lcBulkFileNameWithPath )
            DELETE FILE ( lcBulkFileNameWithPath )
        ENDIF

        IF this.GenBulkInsertTextFile( tcTableName, tcCursorName, lcBulkFileNameWithPath )
            lcBulkInsertString = this.GetBulkInsertString( tcRmtTableName, lcBulkFileNameWithPath )
            * Execute the BULK INSERT into SQL 7 database
            IF this.ExecuteTempSPT( lcBulkInsertString, @lnServerError, @lcErrMsg )
                llRetVal = .T.
*** DH 12/06/2012: handle GenBulkInsertTextFile failing like ExecuteTempSPT failing: ask the
*** user for another mechanism, so comment out ELSE and add ENDIFs to close the above two IFs.
*            ELSE
			endif
		endif

*** DH 07/02/2013: erase the bulk insert file after we're done with it.

        IF FILE( lcBulkFileNameWithPath )
            DELETE FILE ( lcBulkFileNameWithPath )
        ENDIF

*** DH 07/02/2013: added check for llRetVal or tells users bulk insert failed even if worked
		IF not llRetVal
            	IF ( this.UserUpsizeMethod == 0 ) THEN
					lnUserDecision = messagebox( 'The Upsizing Wizard was ' + ;
						'unable to upsize the data using the bulk insert ' + ;
						'technique. However, the data still may be ' + ;
						'upsized using the fast export method.' + chr( 13 ) + ;
						 chr( 13 ) + 'Yes : Use fast export for all tables' + ;
						 chr( 13 ) + 'No : Skip data export for current table' + ;
						 chr( 13 ) + 'Cancel : Skip exporting data', 35, 'Upsize Wizard' )
            	ELSE
            		lnUserDecision = this.UserUpsizeMethod
            	ENDIF

            	IF ( lnUserDecision == 7 ) THEN
            		RETURN lnServerError
            	ENDIF

            	IF ( lnUserDecision == 2 ) THEN
            		RETURN -1
            	ENDIF

            	IF ( lnUserDecision == 6 ) THEN
            		this.UserUpsizeMethod = 6
            	ENDIF

            	* JEI RKR 18.11.2005 Change: Skip all Bulk insert error
*            	IF ( lnServerError <> 4861 ) THEN
*** DH 12/06/2012: handle lnServerError = 0
*				IF !BETWEEN( lnServerError , 4860, 4882 )
				IF lnServerError <> 0 and !BETWEEN( lnServerError , 4860, 4882 )
            		RETURN lnServerError
            	ENDIF
            	&& chandsr added code here.
            	&& If bulkinsert failed due to the fact that the SQL server is on a remote box and the bulk insert file is on the local box
            	&& we retry the operation with fastexport. This would ensure that we export data albeit slower than bulkinsert.
            	lnServerError = this.FastExport ( tcTableName, tcCursorName, tcRmtTableName, @tlMaxErrExceeded )
            	&& chandsr added code here
*** DH 12/06/2012: comment out the ENDIFs because ENDIFs were added above
*            ENDIF
*** DH 07/02/2013: uncommented one of the ENDIFs due to added IF above.
        ENDIF
        RETURN lnServerError


    FUNCTION GenBulkInsertTextFile
        * This function generate a text file that will be the source of the BULK INSERT statement
        * The fastest technique is to use COPY TO, and then clean it up to support BULK INSERT.
        * So far, that means removing double quotes and filling in empty dates.
        LPARAMETERS tcTableName, tcCursorName, tcTargetTextFile
        LOCAL lcFieldList, llRetVal, lnOldSele, lcMsg, lnRecordsCopied, lnRecs, lnxx, lnBytes, ;
            lcFileStr, lcEnumTablesTbl, lcCursor, llHasTimeStamp, llHasIdentity, lnCodePage
		local laFieldNames[1], ;
			lnI, ;
			lcField, ;
			lnField
        lnOldSele = SELECT()

        SELECT ( tcTableName )

        *Thermometer stuff
        lnRecs=RECCOUNT()
        lnRecordsCopied=0
        this.InitTherm( SEND_DATA_LOC, lnRecs, 0 )
        lcMsg=STRTRAN( THIS_TABLE_LOC, '|1', tcTableName )
        this.UpDateTherm( lnRecordsCopied, lcMsg )

        * If adding timestamp or identity, add additional placeholders ( more delimiters )
        * To accomplish this we add additional fields to the source cursor so that COPY TO will
        * automatically add additional placeholders and BULK INSERT will succeed.

        lcEnumTablesTbl = RTRIM( this.EnumTablesTbl )

* Create an array of fields being upsized to Varchar.

		select FLDNAME from ( this.EnumFieldsTbl ) ;
			into array laFieldNames ;
			where trim( TblName ) == lcTableName and RmtType = 'varchar'

* Create a list of fields to output, trimming any Varchar values.

		lcFieldList = ''
		for lnI = 1 to afields( laFields )
			lcField = laFields[lnI, 1]
			lnField = ascan( laFieldNames, lcField, -1, -1, 1, 15 )
			IF lnField > 0
				lcField = 'cast( trim( ' + lcField + ' ) as V( ' + ;
					transform( laFields[lnI, 3] ) + ' )) as ' + lcField
			ENDIF
			lcFieldList = lcFieldList + iif( empty( lcFieldList ), '', ', ' ) + ;
				lcField
		next lnI
        IF &lcEnumTablesTbl..TStampAdd OR ( this.TimeStampAll=1 )
            lcFieldList = lcFieldList + ", '' as tstampcol "
            llHasTimeStamp = .T.
        ENDIF
        IF &lcEnumTablesTbl..IdentAdd OR ( this.IdentityAll=1 )
            lcFieldList = lcFieldList + ", '' as identcol "
            llHasIdentity = .T.
        ENDIF

        lcCursor = this.UniqueCursorName( tcCursorName )
        SELECT &lcFieldList FROM ( tcTableName ) INTO CURSOR ( lcCursor )

		*{ JEI RKR - 2005.04.02 Add Code page
		* Get VFP code page
		lnCodePage = CPCURRENT( 1 )
        * Create the text file
*** MJP 07/15/2013: don't put delimiters around the character fields.
*        COPY TO ( tcTargetTextFile ) DELIMITED WITH CHARACTER BULK_INSERT_FIELD_DELIMITER AS lnCodePage
        COPY TO ( tcTargetTextFile ) DELIMITED WITH "" WITH CHARACTER BULK_INSERT_FIELD_DELIMITER AS lnCodePage
   		*{ JEI RKR - 2005.04.02
        USE IN ( lcCursor )

*** DH 12/06/2012: ensure the file isn't too large
		adir( laFiles, tcTargetTextFile )
		IF laFiles[1, 2] > 30000000
			llRetVal = .F.
			erase ( tcTargetTextFile )
		ELSE
*** DH 12/06/2012: end of new code

	        * Clean up bulk insert file:
	        * - Remove double quotes - BULK INSERT treats them literally
*** MJP 07/15/2013: if we don't add delimiters in the first place, there's no need to
***	remove them.  But we still need to read the file into a variable.
*	        lcFileStr = STRTRAN( FILETOSTR( tcTargetTextFile ), ["] )
            lcFileStr = FILETOSTR( tcTargetTextFile )

	        * - COPY TO leaves blank dates as such: /  /, so
	        *   make empty dates 01/01/1900 - the SQL 7 default
	        lcFileStr = STRTRAN( lcFileStr, "/  /", SQL_SERVER_EMPTY_DATE_CHAR )
	        * 1/6/01 jvf Convert VFP logicals to 0 and 1 for SQL Server.
	        * - COPY TO makes VFP logicals as such: F and T, so
	        *   make them 0 and 1, resp. for SQL Server
	        lcFileStr = this.ConvertVFPLogicalToSQLServerBit( tcTableName, lcFileStr, llHasTimeStamp, llHasIdentity )

	        lnBytes = STRTOFILE( lcFileStr, tcTargetTextFile )
	        llRetVal = ( lnBytes > 0 )
*** DH 12/06/2012: ENDIF to match IF added above
		ENDIF

        SELECT ( lnOldSele )
        RETURN llRetVal


    FUNCTION ConvertVFPLogicalToSQLServerBit
        LPARAMETERS tcTableName, lcFileStr, llHasTimeStamp, llHasIdentity
*** DH 07/11/2013: rewrote this to be more efficient and handle files > 16 MB in size
*!*	        LOCAL ii, jj, lnLines, lnPos, lcValue, llPosAdj
*!*	        lnLines = ALINES( laLines, lcFileStr )
*!*	        lnSele = SELECT()
*!*	        SELECT ( tcTableName )
*!*	        llPosAdj = 0
*!*	        llPosAdj = llPosAdj + IIF( llHasTimeStamp, 1, 0 )
*!*	        llPosAdj = llPosAdj + IIF( llHasIdentity, 1, 0 )
*!*	        FOR ii = 1 TO FCOUNT()
*!*	            IF ( TYPE( FIELD( ii )) = "L" )
*!*	                FOR jj = 1 TO lnLines
*!*	                    IF jj = 1
*!*	                        lnPos = AT( "~", lcFileStr, ( ii-1 )*jj )+1
*!*	                    ELSE
*!*	                        llNewPosAdj = ( llPosAdj + FCOUNT()-1 ) * ( jj-1 )
*!*	                        lnPos = AT( "~", lcFileStr, ( ii-1+llNewPosAdj ))+1
*!*	                    ENDIF
*!*	                    lcValue = SUBSTR( lcFileStr, lnPos, 1 )
*!*	                    lcFileStr = STUFF( lcFileStr, lnPos, 1, IIF( lcValue = "T", "1", "0" ))
*!*	                ENDFOR
*!*	            ENDIF
*!*	        ENDFOR
*!*	        SELECT ( lnSele )
		local laLines[1], lnLines, laFields[1], lnFields, ii, jj, lcLine, lnPos, lcValue
		lnLines  = alines( laLines, lcFileStr )
		lnFields = afields( laFields, tcTableName )
		for ii = 1 to lnFields
			IF laFields[ii, 2] = 'L'
				for jj = 1 to lnLines
					lcLine      = laLines[jj] + '~'
					lnPos       = at( '~', lcLine, ii )
					lcValue     = substr( lcLine, lnPos - 1, 1 )
					laLines[jj] = left( stuff( lcLine, lnPos - 1, 1, ;
						iif( lcValue = 'T', '1', '0' )), len( lcLine ) - 1 )
				next jj
			ENDIF
        next ii
		lcFileStr = ''
		for jj = 1 to lnLines
			lcFileStr = lcFileStr + laLines[jj] + chr( 13 ) + chr( 10 )
		next jj
*** DH 07/11/2013: end of new code
        RETURN lcFileStr


    FUNCTION GetBulkInsertString
        * jvf 9/1/99 Build bulk insert string per export table
        * Used COPY TO ( textfile ) DELI to create source file
        * ( while extracting binary and memo fields ).
        LPARAMETERS tcRmtTableName, tcSourceTextFile
        lcBulkInsStr = "BULK INSERT " + tcRmtTableName + ;
            " FROM '" + tcSourceTextFile + "' WITH " + ;
            " ( CODEPAGE = 'ACP', " + ;
            " FIELDTERMINATOR = '" + [BULK_INSERT_FIELD_DELIMITER] + "', " + ;
            " TABLOCK, KEEPIDENTITY, KEEPNULLS )"

        RETURN lcBulkInsStr


    FUNCTION CreateSQLServerSproc
        PARAMETERS lcTableName, lcRmtTableName, lnBigRows, lnSmallRows, llTStamp

        LOCAL lcEnumFields, lcSQL, lcParamString, lcInsertString, ;
            lcEnumTables, lcCR, lcLF, lcParamName, lcFieldName, llSmallCondition, ;
            lcType, lcLength, I, ii, j, lcSmallParamString, lcSmallSQL, llRetVal, ;
            lnOldArea, LCAS, lcParamChar, lcDecimals, lcColumnsList

        #DEFINE SQL_PARAM_CHAR			"@a"
        #DEFINE ORACLE_PARAM_CHAR		"x"
        #DEFINE BIG_PARAM "Big"

        *If the table has no data, don't create the sproc
        IF lnBigRows=0 	THEN
            RETURN
        ENDIF

        lnOldArea=SELECT()
        lcEnumTables=this.EnumTablesTbl

        lcInsertString=""
        lcParamString=""
        lcCR=CHR( 10 )
        lcLF=CHR( 13 )
        lnOldArea=SELECT()

        lcEnumFields=this.EnumFieldsTbl
        SELECT ( lcEnumFields )
        COPY TO ARRAY aFieldNames FIELDS RmtType, RmtLength, RmtPrec, fldname FOR RTRIM( TblName )==lcTableName

        lcSQL="CREATE PROCEDURE " + DATA_PROC_NAME

        *Oracle () around the whole parameter string

        IF this.ServerType="Oracle" THEN
            *Oracle doesn't allow params to begin with "@"
            lcParamString=" ( " + ORACLE_PARAM_CHAR + BIG_PARAM + " CHAR, "
            lcParamChar=ORACLE_PARAM_CHAR
        ELSE
            *SQL Server requires that params begin with "@"
            lcParamString=" " + SQL_PARAM_CHAR + BIG_PARAM + " char( 4 ), "
            lcParamChar=SQL_PARAM_CHAR
        ENDIF

        *@Big parameter is used to execute all ( for "big" inserts ) or part ( for "small" )
        *inserts ) of the insert statements in the sprocs that get created.  This way
        *we don't need two sprocs; all other parameters are for data.  This means
        *we can only send 254 data parameters at a time so this routine only works
        *for tables with 254 or less fields

        lcColumnNames = " ( "

        FOR iRow = 1 TO ALEN( aFieldNames ) / 4 STEP 1				&& compute the number of rows by dividing the length by the number of columns in the array
        	* JEI RKR 18.11.2005 Add "[" and "]" for field with name like User
        	lcColumnNames = lcColumnNames + "[" + ALLTRIM ( aFieldNames( iRow, 4 )) + "]"

        	IF ( iRow < ALEN ( aFieldNames ) / 4 ) THEN
	        	lcColumnNames = lcColumnNames + ", "
	        ELSE
	        	lcColumnNames = lcColumnNames + " ) "
	        ENDIF
        ENDFOR

        j=0

        FOR I=1 TO lnBigRows
			*{ JEI RKR 18.11.2005 Add "[" and "]"
			lcRmtTableName = ALLTRIM( lcRmtTableName )
			IF LEFT( lcRmtTableName, 1 ) <> "["
				lcRmtTableName = "[" + lcRmtTableName + "]"
			ENDIF
			*} JEI RKR 18.11.2005 Add "[" and "]"
            lcInsertString=lcInsertString + "INSERT INTO " + RTRIM( lcRmtTableName ) + lcColumnNames + ;
                lcCR + "VALUES " + "( "

            FOR ii=1 TO ALEN( aFieldNames, 1 )
                j=j+1
                lcParamName=lcParamChar + LTRIM( STR( j ))
                lcType=RTRIM( aFieldNames[ii, 1] )
                lcLength=LTRIM( STR( aFieldNames[ii, 2] ))
                lcDecimals = LTRIM( STR( aFieldNames[ii, 3] ))
                IF lcLength="0" OR this.ServerType="Oracle" THEN
                    *Note: Oracle parameters take a datatype but no length or precision
                    lcLength=""
                ELSE
                    lcLength= "( " + lcLength + IIF( lcDecimals # '0', ", " + lcDecimals, "" ) + " )"
                ENDIF

                *builds sproc param string like "@a1 char( 4 ), @a2 text, " etc. or
                *"x1 varchar2, x2 number, " etc. for Oracle

                && chandsr added code.
                lctype = SUBSTR( lctype, 1, IIF( AT( '( ', lcType ) = 0, LEN( lcType ), AT ( '( ', lcType ) - 1 ) )
                && Complicated expression just for fun ( just kidding ).
                && if expression  is of the form int ( ident ) the result of executing this expression would be int.
                && this has been done to ensure that the datatype definition such as the one mentioned above do not cause errors on SQL server.
                && chandsr added code.

                lcParamString=lcParamString + lcParamName + " " + lcType + lcLength + ", "

                *builds list of values for INSERT string like "@1, @2, " etc.
                lcInsertString=lcInsertString  + lcParamName + ", "

            NEXT ii

            *If column has a timstamp column, add parameter for that
            *IF llTStamp THEN
            *	lcInsertString=lcInsertString + "NULL )" + lcCR +lcLF
            *ELSE
            *otherwise, peel off last comma, add closing paren
            lcInsertString=LEFT( lcInsertString, LEN( lcInsertString )-2 )
            *Oracle likes a semicolon after each SQL command
            IF this.ServerType="Oracle" THEN
                lcInsertString=lcInsertString  + " );" + lcCR +lcLF
            ELSE
                lcInsertString=lcInsertString  + " )" + lcCR +lcLF
            ENDIF
            *ENDIF

            *if the sproc code includes enough row inserts for the "small"
            *insert sproc, tack on code that will skip over the remaining inserts
            IF I=lnSmallRows THEN
                lcParamName=lcParamChar + BIG_PARAM
                lcInsertString=lcInsertString  + "IF " + lcParamName + " = " + ;
                    "'" + "TRUE" + "'" + lcCR + lcLF
                IF this.ServerType="Oracle" THEN
                    lcInsertString=lcInsertString  + "THEN" + lcCR
                ELSE
                    lcInsertString=lcInsertString  + "BEGIN" + lcCR
                ENDIF
                llSmallCondition=.T.
            ENDIF

        NEXT I

        *peel off last comma, put whole string together
        lcParamString=LEFT( lcParamString, LEN( lcParamString )-2 )

        IF this.ServerType="Oracle" THEN
            *Add closing paren, BEGIN statement
            lcParamString=lcParamString + " )"
            LCAS= " AS BEGIN"
        ELSE
            LCAS= " AS "
        ENDIF

        lcSQL=lcSQL + lcParamString + LCAS + lcCR + lcLF + lcInsertString
        IF llSmallCondition OR this.ServerType="Oracle" THEN
            IF this.ServerType="Oracle" THEN
                *The Oracle IF..THEN construct requires an "END IF"
                IF llSmallCondition THEN
                    lcSQL=lcSQL + "END IF;" + lcCR + lcLF
                ENDIF
                *Oracle sprocs have to be ended explicitly
                lcSQL=lcSQL + "END " + DATA_PROC_NAME + ";"
            ELSE
                lcSQL=lcSQL + "END"
            ENDIF
        ENDIF

        *drop any existing sproc
        =this.ExecuteTempSPT( "drop procedure " + DATA_PROC_NAME )

        *Create the stored procedure
        llRetVal=this.ExecuteTempSPT( lcSQL )

        SELECT ( lnOldArea )



    FUNCTION RowHeuristics
        PARAMETERS lcCursorName, lnBigRows, lnSmallRows,  lnBigBlocks, lnSmallBlocks
        LOCAL lnReccount, lnFieldCount, lnTotalParams, lnOldArea

        *Figures out how many rows stored procedures should insert at a time

        lnOldArea=SELECT()
        SELECT ( lcCursorName )

        *used to be lnRecCount=reccount() but that counts deleted records too
        SELECT COUNT( * ) FROM ( lcTableName ) WHERE !DELETED() INTO ARRAY aReccount
        lnReccount=aReccount
        lnFieldCount=AFIELDS( aTemp )
        lnTotalParams=aReccount*lnFieldCount
        SELECT ( lnOldArea )

        IF lnTotalParams<=MAX_PARAMS THEN
            IF lnReccount<=MAX_ROWS
                lnBigRows=lnReccount
                lnBigBlocks=1
                lnSmallRows=0
                lnSmallBlocks=0
            ELSE
                lnBigRows=MAX_ROWS
                lnBigBlocks=1
                lnSmallRows=lnReccount-MAX_ROWS
                lnSmallBlocks=1
            ENDIF
        ELSE
            IF INT( MAX_PARAMS/lnFieldCount )> MAX_ROWS
                lnBigRows=MAX_ROWS
                lnBigBlocks=INT( lnReccount/lnBigRows )
                lnSmallRows=lnReccount-( lnBigBlocks*lnBigRows )
                lnSmallBlocks=1
            ELSE
                lnBigRows=INT( MAX_PARAMS/lnFieldCount )
                lnBigBlocks=INT( lnReccount/lnBigRows )
                lnSmallRows=lnReccount-( lnBigBlocks*lnBigRows )
                lnSmallBlocks=1
            ENDIF
        ENDIF



    FUNCTION ExecuteSproc
        PARAMETERS lcTableName, lcCursorName, lcRmtTableName, lnBigBlocks, lnSmallBlocks, ;
            lnBigRows, lnSmallRows, llTStamp, llMaxErrExceeded

        LOCAL lnOldArea, lnNumberOfFields, lnLoopMax, lnRecs, lcMsg, ;
            lcNumberofRows, lcArrayType, llRetVal, lcLoopLimiter, lnRowsToCopy, ;
            lnRecordsCopied, lcSQL, lnExportErrors, lcDataErrTable, lnMaxErrors, ;
            llBail, lcData, lnSQLErrno, lcSQLErrMsg, I, ii
		local laFieldNames[1], ;
			lcField, ;
			lnField, ;
			lcValue
*** DH 07/02/2013: added more variables to LOCAL
		local laFields[1], lnFields

		*{ JEI RKR  23.11.2005 ADD
		LOCAL lcDateFieldList, lcNumericFieldsListAllowNull, lcNumericFieldsListNotAllowNull, lcCurentField, lcTempCursorName as String
		LOCAL lnCurrentField, lnCurrentRow, lnCurrentRowInArray, lnRes, lnOldWorkArea as Integer
		*} JEI RKR  23.11.2005 ADD

        PRIVATE aBigArray, aSmallArray

        *If the table has no data, bail
        IF lnBigRows=0 THEN
            RETURN 0
        ENDIF

        lnOldArea=SELECT()
        SELECT ( lcCursorName )
        GO TOP

        lnNumberOfFields=AFIELDS( aTemp )

        *Thermometer stuff
        lnRecs=RECCOUNT()
        lnRecordsCopied=0
        this.InitTherm( SEND_DATA_LOC, lnRecs, 0 )
        lcMsg=STRTRAN( THIS_TABLE_LOC, '|1', lcTableName )
        this.UpDateTherm( lnRecordsCopied, lcMsg )

        *Set maximum number of errors allowed so user's disk doesn't fill up if
        *something goes wrong over and over
        lnMaxErrors=lnRecs*DATA_ERR_FRACTION
        IF lnMaxErrors<DATA_ERR_MIN THEN
            lnMaxErrors=DATA_ERR_MIN
        ENDIF
        lnExportErrors=0

        DIMENSION aBigArray[lnBigRows, lnNumberOfFields]

        lcArrayType="aBigArray"
        lcNumberofRows="lnBigRows"

        *Create big strings that look like this:
        *:aBigArray[1, 1], :aBigArray[1, 2], :aBigArray[1, 3], :aBigArray[1, 4]

*** DH 11/20/2012: moved this block of code inside main loop because lcNumberOfRows can change
*        lcData=""
*        FOR I=1 TO &lcNumberofRows
*            FOR ii=1 TO lnNumberOfFields
*                IF !lcData=="" THEN
*                    lcData=lcData + " , "
*                ENDIF
*                lcData=lcData+ RMT_OPERATOR + "&lcArrayType[" + LTRIM( STR( I )) + ", " + ;
*                    LTRIM( STR( ii )) + "]"
*            NEXT ii
*        NEXT I

        *Set variables for first pass through loop
        lcArrayType="aBigArray"
        lcLoopLimiter="lnBigBlocks"
        lnRowsToCopy=lnBigRows
        lnRecordsCopied=1
        lcBigInsert="TRUE"		&&determines if all inserts in sproc are executed

		*{ JEI RKR  23.11.2005 ADD
		lcDateFieldList = this.GetColumnNum( lcTableName, "DT", .T. )
		lcNumericFieldsListAllowNull = this.GetColumnNum( lcTableName, "YBFIN", .T. )
		lcNumericFieldsListNotAllowNull = this.GetColumnNum( lcTableName, "YBFIN", .F. )
		*} JEI RKR  23.11.2005 ADD

* Create an array of fields being upsized to Varchar.

		select FLDNAME from ( this.EnumFieldsTbl ) ;
			into array laFieldNames ;
			where trim( TblName ) == lcTableName and RmtType = 'varchar'

*** DH 07/02/2013: Used CurrBlock instead of I because I is used for an inner loop as well
*        FOR I=1 TO IIF( lnSmallRows=0, 1, 2 )
        FOR CurrBlock=1 TO IIF( lnSmallRows=0, 1, 2 )
            *In case massive data export errors occurred
            IF lnExportErrors>lnMaxErrors THEN
                this.UpDateTherm( lnRecordsCopied, CANCELED_LOC )
                llMaxErrExceeded=.T.
                EXIT FOR
            ENDIF

*** DH 11/20/2012: moved this block of code inside main loop because lcNumberOfRows can change
            *Build the array string
*            IF this.ServerType="Oracle" THEN
*                lcSQL="BEGIN " + DATA_PROC_NAME + " ( " + "'" + lcBigInsert + "'" + ;
*                    ", " + lcData+ 	" ); END;"
*            ELSE
*                lcSQL="EXECUTE " + DATA_PROC_NAME + "'" + lcBigInsert + "' " + ;
*                    ", " + lcData
*            ENDIF

            *Fill it and send it
            FOR ii= 1 TO &lcLoopLimiter
*** DH 11/20/2012: release the array because at the end of the table, we may
*** have fewer records left than existing rows in the array, which means there
*** are superfluous rows in the array.
				release &lcArrayType
                *get a load of data
                COPY TO ARRAY &lcArrayType NEXT lnRowsToCopy
                skip
*** DH 11/20/2012: recalculate lcNumberOfRows so we only work with the number
*** of rows actually in the array. Since lcNumberOfRows may change, we'll also
*** determine lcSQL here rather than earlier.
				lcNumberOfRows = transform( alen( &lcArrayType., 1 ))
				lcData=""
				FOR I=1 TO &lcNumberofRows
				    FOR ij=1 TO lnNumberOfFields
				        IF !lcData=="" THEN
				            lcData=lcData + " , "
				        ENDIF
				        lcData=lcData+ RMT_OPERATOR + "&lcArrayType[" + LTRIM( STR( I )) + ", " + ;
				            LTRIM( STR( ij )) + "]"
				    NEXT ij
				NEXT I
*Build the array string
				IF this.ServerType="Oracle" THEN
				    lcSQL="BEGIN " + DATA_PROC_NAME + " ( " + "'" + lcBigInsert + "'" + ;
				        ", " + lcData+ 	" ); END;"
				ELSE
				    lcSQL="EXECUTE " + DATA_PROC_NAME + "'" + lcBigInsert + "' " + ;
				        ", " + lcData
				ENDIF
*** DH 11/20/2012: end of moved code

				*{ JEI RKR 23.11.2005 Add
				IF !EMPTY( STRTRAN( lcDateFieldList + ;
					lcNumericFieldsListAllowNull + ;
					lcNumericFieldsListNotAllowNull , ", ", "" )) or ;
					not empty( laFieldNames[1] )
					FOR lnCurrentField = 1 TO lnNumberOfFields
						lcCurentField = ", " + TRANSFORM( lnCurrentField ) + ", "
						IF lcCurentField $ lcDateFieldList
							FOR lnCurrentRow = 1 TO &lcNumberofRows.
								IF EMPTY( &lcArrayType.[lnCurrentRow, lnCurrentField] )
*** DH 07/02/2013: use the chosen blank date value rather than assuming null
*									&lcArrayType.[lnCurrentRow, lnCurrentField] = .Null.
									&lcArrayType.[lnCurrentRow, lnCurrentField] = this.BlankDateValue
								ENDIF
							ENDFOR
						ELSE
							IF lcCurentField $ ", " + lcNumericFieldsListAllowNull + ", " + lcNumericFieldsListNotAllowNull + ", "
								FOR lnCurrentRow = 1 TO &lcNumberofRows.
									IF "*" $ TRANSFORM( &lcArrayType.[lnCurrentRow, lnCurrentField] )
										IF lcCurentField $ ", " + lcNumericFieldsListAllowNull
											&lcArrayType.[lnCurrentRow, lnCurrentField] = .NULL.
										ELSE
											&lcArrayType.[lnCurrentRow, lnCurrentField] = 0
										ENDIF
									ENDIF
								ENDFOR
							ENDIF
						ENDIF

* If this field is being upsized to Varchar, trim its values.

						lcField = field( lnCurrentField )
						lnField = ascan( laFieldNames, lcField, -1, -1, 1, 15 )
						IF lnField > 0
							for lnCurrentRow = 1 to &lcNumberofRows
								lcValue = &lcArrayType.[lnCurrentRow, lnCurrentField]
								&lcArrayType.[lnCurrentRow, lnCurrentField] = trim( lcValue )
							next lnCurrentRow
						ENDIF
					ENDFOR
				ENDIF
				*} JEI RKR 23.11.2005 Add

                *send it to the server
                IF !this.ExecuteTempSPT( lcSQL, @lnSQLErrno, @lcSQLErrMsg ) THEN
                		*{ JEI RKR 05.01.2006 Add and coment
                		lnOldWorkArea = SELECT()
               			lcTempCursorName = SYS( 2015 )
						lcTempCursorName = this.UniqueTableName( lcTempCursorName )
*** DH 07/02/2013: create a cursor that allows nulls in all fields to avoid an
*** error if the target allows nulls but the source doesn't.
*						SELECT * FROM ( lcCursorName ) WHERE .F. INTO CURSOR ( lcTempCursorName ) READWRITE
						lnFields = afields( laFields, lcCursorName )
						for lnField = 1 to lnFields
							laFields[lnField,  5] = .T.
*** DH 07/10/2013: blank out the other values so they don't cause problems ( e.g.
*** field rules, triggers, etc. )
							laFields[lnField,  7] = ''
							laFields[lnField,  8] = ''
							laFields[lnField,  9] = ''
							laFields[lnField, 10] = ''
							laFields[lnField, 11] = ''
							laFields[lnField, 12] = ''
							laFields[lnField, 13] = ''
							laFields[lnField, 14] = ''
							laFields[lnField, 15] = ''
							laFields[lnField, 16] = ''
							laFields[lnField, 17] = 0
							laFields[lnField, 18] = 0
						next lnField
						create cursor ( lcTempCursorName ) from array laFields
*** DH 07/02/2013: end of new code
						SELECT ( lcTempCursorName )
						APPEND FROM ARRAY &lcArrayType.

						lnRes = this.JimExport( lcTableName, lcTempCursorName, lcRmtTableName, @llMaxErrExceeded, @lcDataErrTable )
						IF lnRes <> 0
							lnExportErrors = lnExportErrors + lnRes
						ENDIF
						IF USED( lcTempCursorName )
							USE IN ( lcTempCursorName )
						ENDIF
						SELECT( lnOldWorkArea )
*!*	                    SKIP -1*lnRowsToCopy
*!*	                    COPY TO ARRAY aDatErr NEXT lnRowsToCopy
*!*	                    SKIP
*!*	                    this.DataExportErr( @aDatErr, lcTableName, @lcDataErrTable, lnSQLErrno, lcSQLErrMsg )
*!*	                    lnExportErrors=lnExportErrors+lnRowsToCopy
						*} JEI RKR 05.01.2006
                ENDIF

                *Thermometer
                lnRecordsCopied=lnRecordsCopied+lnRowsToCopy

                *If massive export errors, bail
                IF lnExportErrors>lnMaxErrors THEN
                    this.UpDateTherm( lnRecordsCopied, CANCELED_LOC )
                    llMaxErrExceeded=.T.
                    EXIT FOR
                ENDIF

                IF lnExportErrors<>0 THEN
                    this.UpDateTherm( lnRecordsCopied, lcMsg + ", " + LTRIM( STR( lnExportErrors ))+ " " + ERROR_COUNT_LOC )
                ELSE
                    this.UpDateTherm( lnRecordsCopied, lcMsg )
                ENDIF

            NEXT ii

            *if there are leftover rows after the "big" copies, do the inner loop again with
            *the smaller array
            *lcArrayType="aSmallArray"
            *The extra parameters not used in the small insert are set to null here

            aBigArray=.NULL.
            lcLoopLimiter="lnSmallBlocks"
            lnRowsToCopy=lnSmallRows
            this.BitifyArray( lcTableName, lnSmallRows )
            lcBigInsert="FALS"		&&only takes four characters

        NEXT CurrBlock

        *If export errors occurred, close the error table
        IF lnExportErrors<>0 THEN
			IF USED( lcDataErrTable ) && JEI RKR 05.01.2006 ADD
				SELECT ( lcDataErrTable )
				USE
			ENDIF
        ENDIF

        IF !llMaxErrExceeded THEN
			raiseevent( This, 'CompleteProcess' )
        ENDIF

        *drop the stored procedure
        lcSQL="DROP PROCEDURE rwf_insert_"
        llRetVal=this.ExecuteTempSPT( lcSQL )

        SELECT ( lnOldArea )

        RETURN lnExportErrors


    FUNCTION BitifyArray
        PARAMETERS lcTableName, lnRowsToCopy
        LOCAL lnOldArea, lcEnumTablesTbl, lnOldRecno, I, j

        *When ExecuteSproc does the "small" insert, the extra ( empty ) rows in
        *the array sent to the server have to be acceptable as parameters.
        *Null is fine for everything but bit fields
        lnOldArea=SELECT()

        *Grab the sql ( table definition ) of the table
        lcEnumTablesTbl=this.EnumTablesTbl
        SELECT ( lcEnumTablesTbl )
        lnOldRecno=RECNO()
        LOCATE FOR LOWER( RTRIM( TblName ))==LOWER( RTRIM( lcTableName ))
        lcSQL=TableSQL

        *Parse it looking for bit fields
        lcSQL=SUBSTR( lcSQL, AT( "( ", lcSQL ))

        * jvf 2/4/01 Bug ID 192139 - subscipt outside defined range on: aBigArray[j, i]=.F.
        * Add space( 1 ) after comma as search string to differentiate
        * field separator vs. numeric column struc with precision ( eg, numeric( 8, 2 )).
        lcSQL=STUFF( lcSQL, RAT( " )", lcSQL ), 1, ", " )
        I=1
        DO WHILE ", " $ lcSQL
            lcSubStr=LEFT( lcSQL, AT( ", ", lcSQL ))
            IF " bit " $ lcSubStr
                FOR j=lnRowsToCopy +1 TO ALEN( aBigArray, 1 )
                    aBigArray[j, i]=.F.
                NEXT j
            ENDIF

            * jvf 2/04/01
            * Money wouldn't upsize in FastExport for same reason bit don't.
            * Added this extra loop to prefill array with 0.00 instead of null
            * for money columns.
            IF " money " $ lcSubStr
                FOR j=lnRowsToCopy +1 TO ALEN( aBigArray, 1 )
                    aBigArray[j, i]=0.00
                NEXT j
            ENDIF

            lcSQL=SUBSTR( lcSQL, AT( ", ", lcSQL )+1 )
            I=I+1
        ENDDO

        GO lnOldRecno
        SELECT ( lnOldArea )



    FUNCTION CreateIndexes
        LOCAL lnOldArea, lcEnum_Indexes, lcSQL, lcScanCondition, I, lnError, ;
            llRetVal, lcClusterName, lnLoopLimiter, lnIndexCount, lcDel, lcErrMsg, ;
            lcTagName, llTableUpsized, lnOldTO, lcMsg

* If we have an extension object and it has a CreateIndexes method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'CreateIndexes', 5 ) and ;
			not this.oExtension.CreateIndexes( This )
			return
		ENDIF

        lnOldArea=SELECT()
        this.InitTherm( STARTING_COMMENT_LOC, 0, 0 )

        *Generate all the sql for the indexes
        this.BuildIndexSQL

        *Create the indexes
        IF this.DoUpsize AND this.Perm_Index  THEN
            *Make sure we don't create indexes with logical/bit fields
            IF this.ServerType<>"Oracle" THEN
                this.MarkBitIndexes
            ENDIF

            lcEnum_Indexes=this.EnumIndexesTbl
            SELECT ( lcEnum_Indexes )

            *Create indexes on tables ( clustered indexes first, otherwise existing
            *indexes automatically are regenerated )

            IF this.ServerType="Oracle" THEN
                lcScanCondition="DontCreate=.F. AND Exported=.F."
                lnLoopLimiter=1
            ELSE
                lcScanCondition="Clustered=.T. AND DontCreate=.F. AND Exported=.F."
                lnLoopLimiter=2
            ENDIF

            *note: if the user selected a table for export, moved ahead a few
            *pages causing the indexes to be analyzed, then deselected that table,
            *then the thermometer reading won't be right--just means that it will
            *jump up fast when some indexes are skipped
            SELECT COUNT( * ) FROM ( lcEnum_Indexes ) WHERE DontCreate=.F. ;
                AND Exported=.F. INTO ARRAY aIndexCount
            lnIndexCount=0
            this.InitTherm( STARTING_COMMENT_LOC, aIndexCount, 0 )

            *Never timeout
            lnOldTO=SQLGETPROP( this.MasterConnHand, "QueryTimeOut" )
            =SQLSETPROP( this.MasterConnHand, "QueryTimeOut", 0 )

            FOR I=1 TO lnLoopLimiter
                SCAN FOR &lcScanCondition
                    *Make sure table was upsized
                    lcTableName=RTRIM( &lcEnum_Indexes..IndexName )
                    lcErrMsg=""
                    lcMsg=STRTRAN( THIS_TABLE_LOC, "|1", lcTableName )
                    IF this.TableUpsized( lcTableName, @llTableUpsized ) THEN
                        *for thermometer
                        this.UpDateTherm( lnIndexCount, lcMsg )
                        lcSQL=&lcEnum_Indexes..IndexSQL

						IF !( "DELETED()" $ UPPER( ALLTRIM( &lcEnum_Indexes..Lclexpr )) ) && JEI RKR 21.11.2005 Add
	                        llRetVal=this.ExecuteTempSPT( lcSQL, @lnError, @lcErrMsg )
	                    ELSE
	                    	&&{ JEI RKR 21.11.2005 Add
	                    	llRetVal = .t.
	                    	lnError = 0
	                    	&&} JEI RKR 21.11.2005 Add
	                    ENDIF

                        IF !llRetVal THEN
                            REPLACE &lcEnum_Indexes..IdxErrNo WITH lnError
                            lcTagName=&lcEnum_Indexes..TagName
                            this.StoreError( lnError, lcErrMsg, lcSQL, INDEX_FAILED_LOC, lcTagName, INDEX_LOC )
                        ENDIF
                    ELSE
                        llRetVal=.F.
                        lcErrMsg=TABLE_NOT_EXPORTED_LOC
                    ENDIF
                    REPLACE &lcEnum_Indexes..Exported WITH llRetVal, ;
                        &lcEnum_Indexes..IdxError WITH lcErrMsg, ;
                        &lcEnum_Indexes..TblUpszd WITH llTableUpsized

                    lnIndexCount=lnIndexCount+1
                    this.UpDateTherm( lnIndexCount, lcMsg )
                ENDSCAN
                lcScanCondition="Clustered=.F. AND DontCreate=.F. AND Exported=.F."
            NEXT

            *Put this back
            IF TYPE( 'lnOldTO' ) = 'N' AND lnOldTO >= 0
                =SQLSETPROP( this.MasterConnHand, "QueryTimeOut", lnOldTO )
            ENDIF

			raiseevent( This, 'CompleteProcess' )
        ENDIF

        SELECT ( lnOldArea )



    FUNCTION MarkBitIndexes
        LOCAL lcEnumIndexesTbl, lcEnumFieldsTbl, aLogicals, lnOldArea, llDontCreate, ;
            lcTableName

        *
        *Marks indexes that have logical fields in them as un-createable
        *

        lcEnumFieldsTbl=this.EnumFieldsTbl
        lcEnumIndexesTbl=this.EnumIndexesTbl

        lnOldArea=SELECT()
        SELECT ( lcEnumIndexesTbl )
        SCAN
            lcTableName=RTRIM( IndexName )
            *Find out which ( if any ) fields in the table are logical fields
            DIMENSION aLogicals[1]
            aLogicals=.F.
            SELECT RmtFldname FROM ( lcEnumFieldsTbl ) WHERE RTRIM( TblName )==lcTableName AND ;
                DATATYPE="L" INTO ARRAY aLogicals
            llDontCreate=.F.

            *See if any of the logical fields are part of the index expression
            IF !EMPTY( aLogicals ) THEN
                FOR jj=1 TO ALEN( aLogicals, 1 )
                    IF RTRIM( aLogicals[jj] ) $ &lcEnumIndexesTbl..RmtExpr THEN
                        llDontCreate=.T.
                        EXIT
                    ENDIF
                NEXT jj
            ENDIF

            IF llDontCreate THEN
                lcMsg=CANT_CREATE_INDEX_LOC
            ELSE
                lcMsg=""
            ENDIF

            REPLACE DontCreate WITH llDontCreate, IdxError WITH lcMsg

        ENDSCAN

        SELECT ( lnOldArea )



    FUNCTION TableUpsized
        PARAMETERS lcTableName, llChosenForExport
        LOCAL lcEnumTablesTbl, lnOldArea, llExport

        *Returns whether a table was actually created on the server successfully or not
        *If the user is generating a script but not upsizing, then the function
        *returns whether the table was selected for upsizing

        lnOldArea=SELECT()
        lcEnumTablesTbl=RTRIM( this.EnumTablesTbl )
        SELECT ( lcEnumTablesTbl )
        LOCATE FOR TblName=lcTableName
        IF this.DoUpsize THEN
            llExport=&lcEnumTablesTbl..Exported
        ELSE
            llExport=&lcEnumTablesTbl..EXPORT
        ENDIF
        llChosenForExport=&lcEnumTablesTbl..EXPORT
        SELECT ( lnOldArea )
        RETURN llExport



    FUNCTION AnalyzeIndexes
        LOCAL lnOldArea, lcEnum_Tables, lcEnum_Indexes, I, ii, lcExprLeftover, ;
            lcTablePath, lcEnum_Fields, lcExpression, lcTagName, lcRemoteExpression, lcRemoteTagName, ;
            lcSQL, lcRemoteTable, lcClustered, llUserTableOpened, lcTableName, lcMsg, llCreateIndexes, ;
            aKeyFields, llDontCreate, lcErrMsg, lcLclIdxType

        *Don't do this routine if not necessary
        IF this.ProcessingOutput THEN
            IF  !this.ExportIndexes AND ;
                    !this.ExportRelations AND ;
                    !this.ExportTableToView AND ;
                    !this.ExportViewToRmt THEN
                RETURN
            ENDIF
        ENDIF

* If we have an extension object and it has a AnalyzeIndexes method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'AnalyzeIndexes', 5 ) and ;
			not this.oExtension.AnalyzeIndexes( This )
			return
		ENDIF

        lnOldArea=SELECT()
        lcEnum_Fields=this.EnumFieldsTbl
        lcEnum_Tables=this.EnumTablesTbl
        SELECT COUNT( * ) FROM ( lcEnum_Tables ) WHERE EXPORT=.T. AND EMPTY( ClustName )=.T. ;
            AND CDXAnald=.F. INTO ARRAY aTableCount

        *Thermometer stuff
        IF aTableCount=0 THEN
            RETURN
        ENDIF
        lnTableCount=0
        IF this.ProcessingOutput or this.lQuiet
            lcMsg=STRTRAN( ANALYZING_INDEXES_LOC, "..." )
            this.InitTherm( lcMsg, aTableCount, 0 )
        ELSE
            WAIT ANALYZING_INDEXES_LOC WINDOW NOWAIT
        ENDIF

        *Create table to hold index names and expressions
        IF RTRIM( this.EnumIndexesTbl )==""
            lcEnum_Indexes=this.CreateWzTable( "Indexes" )
            this.EnumIndexesTbl=lcEnum_Indexes
            llCreateIndexes=.T.
        ELSE
            lcEnum_Indexes=this.EnumIndexesTbl
            llCreateIndexes=.F.
        ENDIF

        *read tables-to-be-exported one at a time and see if they have .CDXs
        SELECT ( lcEnum_Tables )

        SCAN FOR EXPORT=.T. AND CDXAnald=.F. AND EMPTY( ClustName )=.T.
            lcCursorName=RTRIM( &lcEnum_Tables..CursName )
            lcTableName=RTRIM( &lcEnum_Tables..TblName )
            lcRemoteTable=RTRIM( &lcEnum_Tables..RmtTblName )
            lcTablePath=RTRIM( &lcEnum_Tables..TblPath )

            *Therm stuff
            lcMsg=STRTRAN( THIS_TABLE_LOC, '|1', lcTableName )
            this.UpDateTherm( lnTableCount, lcMsg )
            lnTableCount=lnTableCount+1

            *Read information for each tag

            SELECT ( lcCursorName )
            FOR I=1 TO TAGCOUNT( STRTRAN( lcTablePath, ".DBF", ".CDX" ))
                lcRemoteExpression=""
                lcClustered=.F.
                lcExpression=LOWER( SYS( 14, I ))		&&tag expression
                lcTagName=TAG( I, lcCursorName )		&&tag name

                *Figure out index type
                DO CASE
                    CASE PRIMARY( I )
                        lcLclIdxType="Primary key"
                        IF this.ServerType="Oracle" OR this.ServerType=="SQL Server95" THEN
                            lcTagType="PRIMARY KEY"
                            IF this.ServerType="Oracle" THEN
                                lcClustered=.F.
                            ELSE
                                lcClustered=.T.
                            ENDIF
                        ELSE
                            lcTagType="UNIQUE CLUSTERED"
                            lcClustered=.T.
                        ENDIF
                    CASE CANDIDATE( I )
                        lcLclIdxType="Candidate"
                        *same for both Oracle and SQL Server
                        lcTagType="UNIQUE"
                    OTHERWISE
                        IF UNIQUE() THEN
                            lcLclIdxType="Unique"
                        ELSE
                            lcLclIdxType="Regular"
                        ENDIF
                        *if UNIQUE() or just a regular index
                        lcTagType=""
                ENDCASE

                *pull the field names out of each expression into comma separated list
                lcRemoteExpression=this.ExtractFieldNames( lcExpression, lcTableName )
                DIMENSION aKeyFields[1]
                aKeyFields=.F.
                this.KeyArray( lcRemoteExpression, @aKeyFields )
                IF ALEN( aKeyFields, 1 )>MAX_INDEX_FIELDS THEN
                    llDontCreate=.T.
                    lcErrMsg=TOO_MANY_FIELDS_LOC
                ELSE
                    llDontCreate=.F.
                    lcErrMsg=""
                ENDIF

                *Write info into index analysis table
                SELECT ( lcEnum_Indexes )
                APPEND BLANK
                REPLACE IndexName WITH lcTableName, ;
                    TagName WITH LOWER( lcTagName ), ;
                    LclExpr WITH lcExpression, 	;
                    LclIdxType WITH lcLclIdxType, ;
                    RmtExpr WITH lcRemoteExpression, ;
                    RmtName WITH LOWER( this.RemotizeName( lcTagName )), ;
                    RmtType WITH lcTagType, ;
                    Clustered WITH lcClustered, ;
                    RmtTable WITH lcRemoteTable, ;
                    Exported WITH .F., ;
                    DontCreate WITH llDontCreate, ;
                    IdxError WITH lcErrMsg

                *Fox will let you have multiple "primary keys" on a table.  Oracle and SQL 95
                *declarative RI won't.  Consequently, the wizard needs to be
                *sure that the primary key on a table is the one used
                *for RI; the other "primary keys" can just be regular indexes.

                *Here the pkey expression gets stored for comparison later
                IF lcTagType="PRIMARY KEY" THEN
                    SELECT ( lcEnum_Tables )
                    REPLACE PkeyExpr WITH lcRemoteExpression, ;
                        PKTagName WITH lcTagName
                ENDIF

                SELECT ( lcCursorName )

            NEXT I

            SELECT ( lcEnum_Tables )
            REPLACE CDXAnald WITH .T.

        ENDSCAN

        SELECT ( lcEnum_Indexes )
        SELECT ( lnOldArea )
        this.AnalyzeIndexesRecalc=.F.

        IF this.ProcessingOutput or this.lQuiet
			raiseevent( This, 'CompleteProcess' )
        ELSE
            WAIT CLEAR
        ENDIF



    FUNCTION BuildIndexSQL
        LOCAL lcEnum_Indexes, lcSQL, lcRmtTable, lcRmtIdxName, lcRmtType, lcConstraint, lcTSClause, lcClustered

        *Build the CREATE INDEX sql string
        lcEnum_Indexes = this.EnumIndexesTbl
        SELECT ( lcEnum_Indexes )

        SCAN FOR Exported=.F.
            *Get remote table name
            lcRmtTable = LEFT( this.RemotizeName( RTRIM( &lcEnum_Indexes..IndexName )), 30 )
            lcRmtType = TRIM( &lcEnum_Indexes..RmtType )
            lcRmtExpr = TRIM( &lcEnum_Indexes..RmtExpr )
            lcRmtName = TRIM( &lcEnum_Indexes..RmtName )
            IF this.ServerType = ORACLE_SERVER
                lcRmtName = this.UniqueOraName( lcRmtName )
            ENDIF

            *check index type; deal with Oracle and Primary differently
            IF lcRmtType = "PRIMARY KEY" OR ;
                    ( lcRmtType = "UNIQUE" AND this.ServerType = "Oracle" ) ;
                    OR ( lcRmtType = "UNIQUE" AND this.ServerType == "SQL Server95" ) THEN

                * Oracle and SQL95 Unique and Primary Key indexes implemented via ALTER TABLE
                IF this.ServerType = ORACLE_SERVER
                    lcTSClause = IIF( !EMPTY( this.TSIndexTSName ), " USING INDEX TABLESPACE " + this.TSIndexTSName, "" )
                    lcSQL = "ALTER TABLE " + lcRmtTable + " ADD ( CONSTRAINT " + lcRmtName + ;
                        " " + lcRmtType + " ( " + lcRmtExpr + " )" + lcTSClause + " )"
                ELSE
                    * jvf 08/13/99
                    * Use this convention for unique db objects...
                    * PRIMARY KEY: "PK_" + lcRmtName, 3
                    * CANDIDATE KEYS: "UQ_" + lcRmtName

                    * ## Add NONCLUSTERED clause when creating a PRIMARY KEY b/c SS
                    * defaults to CLUSTERED INDEX.

                    *{ JEI RKR 14.07.2005 Change: Add check for CLUSTERED clause in lcRmtType
                    lcClustered = ""
                    IF !( " CLUSTERED" $ lcRmtType )
	                    lcClustered = " NONCLUSTERED "
                    ENDIF
                    *} JEI RKR 14.07.2005
                    IF lcRmtType = "UNIQUE"
                        * Can have multiple, so get unique name
                        lcRmtName = "UQ_" + this.UniqueTableName( lcRmtTable )
                    ELSE &&Primary Key
                        lcRmtName = "PK_" + lcRmtTable
                        IF !EMPTY( this.ExportClustered )
                            lcClustered = " CLUSTERED "
                        ENDIF
                    ENDIF
                    lcSQL = "ALTER TABLE [" + lcRmtTable + "] ADD CONSTRAINT " + lcRmtName + ;
                        " " + lcRmtType + lcClustered + "( " + lcRmtExpr + " )"
                    *!*End Mark
                ENDIF
            ELSE
                * All other index types ( including regular Oracle indexes ) are handled here
                lcSQL = "CREATE " + lcRmtType + " INDEX [" + lcRmtName + "] ON [" + lcRmtTable + "] ( " + lcRmtExpr + " )"
                IF this.ServerType = ORACLE_SERVER
                    lcTSClause = IIF( !EMPTY( this.TSIndexTSName ), " TABLESPACE " + this.TSIndexTSName, "" )
                    lcSQL = lcSQL + lcTSClause
                ENDIF
            ENDIF

            REPLACE &lcEnum_Indexes..IndexSQL WITH lcSQL, &lcEnum_Indexes..RmtName WITH lcRmtName
        ENDSCAN




    FUNCTION CreateDevice
        PARAMETERS lcDeviceType, lcDevicePhysName, lnDeviceNumber

        LOCAL lcDeviceLogicalName, lcSQL1, lcSQL2, lcSQL3, I, lnRetVal, ;
            lcRoot, lcMasterPath, lnServerErr, lnUserChoice, lcErrMsg, ;
            lcDevicePhysPath, lcSQT
*** This used to be set in GetDeviceNumbers
        local aDeviceNumbers[1]

* If we have an extension object and it has a CreateDevice method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'CreateDevice', 5 ) and ;
			not this.oExtension.CreateDevice( lcDeviceType, lcDevicePhysName, ;
			lnDeviceNumber, This )
			return
		ENDIF

        lcSQT=CHR( 39 )
        *Fill local variables from global
        IF lcDeviceType="Log" THEN
            *If the log and database names are the same, bail
            IF this.DeviceLogName=this.DeviceDBName THEN
                RETURN
            ENDIF
            *Put up thermometer
            this.InitTherm( CREATING_LOGDEVICE_LOC, 0, 0 )
            lcDeviceLogicalName=RTRIM( this.DeviceLogName )
            lnDeviceSize=this.DeviceLogSize
        ELSE
            *Put up thermometer
            this.InitTherm( CREATING_DBDEVICE_LOC, 0, 0 )
            lcDeviceLogicalName=RTRIM( this.DeviceDBName )
            lnDeviceSize=this.DeviceDBSize
        ENDIF
        this.UpDateTherm( 0, TAKES_AWHILE_LOC )

        *convert from megabytes to 2k pages
        lnDeviceSize=lnDeviceSize*512

        * Build path for physical device based on location of Master database
        IF this.MasterPath=="" THEN
            lcMasterPath=""
            lcSQL = "select phyname from sysdevices where name = " + lcSQT + "master" + lcSQT
            lnRetVal = this.SingleValueSPT( lcSQL, @lcMasterPath, "phyname" )
            this.MasterPath=RTRIM( lcMasterPath )
        ENDIF
        lcDevicePhysPath = this.JUSTPATH( this.MasterPath )+ "\"

        *Build device physical name
        lcRoot = LEFT( lcDeviceLogicalName, 6 )
        lcDevicePhysName = lcDevicePhysPath + lcRoot + ".DAT"

        *Get a device number and mark it as taken
        lnDeviceNumber=ASCAN( aDeviceNumbers, .F. )
        aDeviceNumbers[lnDeviceNumber]=.T.

        *Build sql string
        lcSQL1 = "disk init name=" + lcSQT+ lcDeviceLogicalName + lcSQT+ ", "
        lcSQL2 = "physname=" + lcSQT+ lcDevicePhysName + lcSQT+ ", "
        lcSQL3 = "vdevno=" + ALLTRIM( STR( lnDeviceNumber )) + ", "
        lcSQL3 = lcSQL3 + "size=" + ALLTRIM( STR( lnDeviceSize ))
        lcSQL = lcSQL1 + lcSQL2 + lcSQL3

        *
        * Create device, incrementing physical name if already taken
        * Since error SQS_ERR_DISK_INIT_FAIL can result from other problems, only try 20 times
        *

        IF this.DoUpsize THEN
            *Set query time out to huge value while this is running
            =SQLSETPROP( this.MasterConnHand, "QueryTimeOut", 600 )

            lnServerErr = 0
            lcErrMsg = ""

            *- if problems creating devices ( "DISK command not allowed within multi-statement transactions" ),
            *- uncomment this line
            *- this.ExecuteTempSPT( "COMMIT TRANSACTION", @lnServerErr, @lcErrMsg )

            *Try 20 times because physical name may be taken
            FOR I=1 TO 20
                lnServerErr=0
                IF !this.ExecuteTempSPT( lcSQL, @lnServerErr, @lcErrMsg ) THEN
                    lcDevicePhysName = lcDevicePhysPath + lcRoot + ALLTRIM( STR( I )) + ".DAT"
                    lcSQL2 = "physname=" + lcSQT + lcDevicePhysName + lcSQT+ ", "
                    lcSQL = lcSQL1 + lcSQL2 + lcSQL3
                ELSE
                    EXIT
                ENDIF
            NEXT
            *If the device creation fails, quit
            IF lnServerErr<>0 THEN
	        	IF this.lQuiet
	        		.HadError     = .T.
	        		.ErrorMessage = CANT_CREATE_DEVICE_LOC
				ELSE
	                MESSAGEBOX( CANT_CREATE_DEVICE_LOC, ICON_EXCLAMATION, TITLE_TEXT_LOC )
	        	ENDIF
                this.Die
            ELSE
                *Set the query timeout back to a more reasonable figure
                =SQLSETPROP( this.MasterConnHand, "QueryTimeOut", 30 )
            ENDIF

        ENDIF

        *Store sql for script
        this.StoreSQL( lcSQL, DEVICE_SCRIPT_COMMENT_LOC )
		raiseevent( This, 'CompleteProcess' )



    FUNCTION ExecuteTempSPT
        PARAMETERS lcSQL, lnServerError, lcErrMsg, lcCursor
        LOCAL nRetVal, lnButtons, lcMsg, lcNewSQL, lcEscape

        lcEscape=IIF( this.ProcessingOutput, "ON", SET ( "ESCAPE" ))
        SET ESCAPE OFF

        IF PARAMETERS()=4 THEN
            nRetVal=SQLEXEC( this.MasterConnHand, lcSQL, lcCursor )
        ELSE
            nRetVal=SQLEXEC( this.MasterConnHand, lcSQL )
        ENDIF

        SET ESCAPE &lcEscape

        DO CASE
                *It worked
            CASE nRetVal=1
                lnServerError=0
                lcErrMsg=""
                RETURN .T.

                *Server error occurred
            CASE nRetVal=-1
                =AERROR( aErrArray )
                lnServerError=aErrArray[1]
                lcErrMsg=aErrArray[2]

                IF lnServerError=1526 AND !ISNULL( aErrArray[5] )THEN
                    lnServerError=aErrArray[5]
                ENDIF

                DO CASE
                    CASE lnServerError=1105
                        *If Log is full, try to dump it ( but only if upsizing to existing db )
			        	IF this.lQuiet
			        		.HadError     = .T.
			        		.ErrorMessage = LOG_FULL_LOC
						ELSE
                        	MESSAGEBOX( LOG_FULL_LOC, ICON_EXCLAMATION, TITLE_TEXT_LOC )
			        	ENDIF
                    CASE  lnServerError=1101 OR lnServerError=1510
                        *Device full
			        	IF this.lQuiet
			        		.HadError     = .T.
			        		.ErrorMessage = DEVICE_FULL_LOC
						ELSE
                        	MESSAGEBOX( DEVICE_FULL_LOC, ICON_EXCLAMATION, TITLE_TEXT_LOC )
			        	ENDIF

                    CASE lnServerError=1804
                        *SQL Server bug having to do with dropped device
			        	IF this.lQuiet
			        		.HadError     = .T.
			        		.ErrorMessage = SQL_BUG_LOC
						ELSE
                        	MESSAGEBOX( SQL_BUG_LOC, ICON_EXCLAMATION, TITLE_TEXT_LOC )
			        	ENDIF

                        *this error happens when dropping sproc that doesn't exist
                    CASE lnServerError=3701
                        RETURN .F.

                    CASE lnServerError=102
                        *syntax error in sql
                        RETURN .F.

                    CASE lnServerError=2615
                        *duplicate record entered
                        *should be caused only by running sp_foreignkey on a relation
                        *with the same foreign key twice
                        RETURN .F.

                    OTHERWISE
                        *unknown error
                        RETURN .F.
                ENDCASE

                *Connection level error occurred
            CASE nRetVal=-2
                *This is trouble; continue to generate script if user wants; otherwise bail
                lcMsg=STRTRAN( CONNECT_FAILURE_LOC, "|1", LTRIM( STR( lnServerErr )) )
	        	IF this.lQuiet
	        		.HadError     = .T.
	        		.ErrorMessage = lcMsg
				ELSE
	                MESSAGEBOX( lcMsg, ICON_EXCLAMATION, TITLE_TEXT_LOC )
	        	ENDIF

        ENDCASE

        this.Die


    FUNCTION Die
        IF this.SQLServer
            this.TruncLogOff
        ENDIF
        this.NormalShutdown=.F.
        release oEngine
		cancel



    FUNCTION SingleValueSPT
        PARAMETERS lcSQL, lcReturnValue, lcFieldName, llReturnedOneValue
        LOCAL lcMsg, lcErrMsg, llRetVal, lcCursor, lnOldArea, lnServerError

        *
        *Executes a server query and sees if it return one value or not
        *If it returns one value, that value gets placed in a variable passed by reference
        *

        lnOldArea=SELECT()
        lcCursor=this.UniqueCursorName( "_spt" )
        SELECT 0
        IF this.ExecuteTempSPT( lcSQL, @lnServerError, @lcErrMsg, lcCursor ) THEN
            IF RECCOUNT( lcCursor )=0 THEN
                llReturnedOneValue= .F.
            ELSE
                lcReturnValue=&lcCursor..&lcFieldName
                llReturnedOneValue=.T.
            ENDIF
            USE
        ELSE
            lcMsg=STRTRAN( QUERY_FAILURE_LOC, "|1", LTRIM( STR( lnServerError )) )
        	IF this.lQuiet
        		.HadError     = .T.
        		.ErrorMessage = lcMsg
			ELSE
            	MESSAGEBOX( lcMsg, ICON_EXCLAMATION, TITLE_TEXT_LOC )
        	ENDIF
            this.Die
            RETURN
        ENDIF

        SELECT ( lnOldArea )
        RETURN llReturnedOneValue



    FUNCTION AnalyzeTables
        LOCAL lcEnum_Tables, lcTableName, llUTableOpen, lcSourceDB, lnOldArea, ;
            llUpsizable, lcTablePath, llAlreadyOpened, llWarnUser, lcCursorName, ;
            aOpenTables, I, ctmpTblName

        lcSourceDB=this.SourceDB
        lnOldArea=SELECT()
        SET DATABASE TO ( lcSourceDB )
        lcEnum_Tables=this.CreateWzTable( "Tables" )
        this.EnumTablesTbl=lcEnum_Tables

        IF ADBOBJECTS( aTblArray, "table" )=0 THEN
            RETURN
        ENDIF

        DIMENSION aOpenTables[1, 2]
        =AUSED( aOpenTables )
        FOR I=1 TO ALEN( aTblArray, 1 )
            APPEND BLANK
            lcTablePath= FULL( DBGETPROP( aTblArray( I ), "TABLE", "PATH" ), DBC() )
            llUpsizable=this.Upsizable( aTblArray( I ), lcTablePath, @llAlreadyOpened, @lcCursorName, @aOpenTables )

            REPLACE TblName WITH LOWER( aTblArray( I )), ;
                CursName WITH lcCursorName,  ;
                TblPath WITH lcTablePath, ;
                Upsizable WITH llUpsizable, ;
                PreOpened WITH llAlreadyOpened, ;
                Type	WITH "T" && Add JEI RKR 2005.05.09

            * Check for DBCS
            ctmpTblName = ALLTRIM( TblName )
            IF LEN( m.ctmpTblName )#LENC( m.ctmpTblName )
                ctmpTblName=STRTRAN( ctmpTblName, CHR( 32 ), "_" )
                IF LEN( m.ctmpTblName )>29
                    IF ISLEADBYTE( SUBSTR( m.ctmpTblName, 30, 1 ))
                        ctmpTblName = LEFT( m.ctmpTblName, 29 )
                    ENDIF
                ENDIF
            ENDIF
            REPLACE RmtTblName WITH this.RemotizeName( m.ctmpTblName )

            IF !llUpsizable THEN
                llWarnUser=.T.
            ENDIF
            lcCursorName=""
        NEXT

        IF llWarnUser and not this.lQuiet
            =MESSAGEBOX( NO_OPEN_EXCLU_LOC, ICON_EXCLAMATION, TITLE_TEXT_LOC )
        ENDIF

        this.AnalyzeTablesRecalc=.F.
        SELECT ( lnOldArea )



    FUNCTION AnalyzeClusters
        PARAMETERS aClusters
        LOCAL lcEnumClusters, lnOldArea, lcTableName, I, lnTableCount, aClusterCount

        * this creates the cluster table

        lnOldArea = SELECT()
        aClusterCount = ALEN( aClusters, 1 )
        DIMENSION aClusters( ALEN( aClusters, 1 ), 5 )	&& won't compile without dimension
        IF EMPTY( aClusters[1, 1] ) OR aClusterCount = 0
            RETURN
        ENDIF

        *Create table for clusters if it doesn't exist yet
        IF RTRIM( this.EnumClustersTbl ) == ""
            lcEnumClusters = this.CreateWzTable( "Clusters" )
            this.EnumClustersTbl = lcEnumClusters
        ELSE
            lcEnumClusters = RTRIM( this.EnumClustersTbl )
            SELECT &lcEnumClusters
            ZAP
        ENDIF

        * copy data from aClusters to Clusters table
        lnTableCount = 0
        SELECT ( lcEnumClusters )
        FOR m.I = 1 TO aClusterCount
            lcClusterName = LOWER( aClusters[m.i, 1] )
            lcMsg = STRTRAN( THIS_TABLE_LOC, "|1", lcClusterName )
            this.UpDateTherm( lnTableCount, lcMsg )

            APPEND BLANK
            REPLACE ClustName WITH  lcClusterName, ;
                ClustType WITH aClusters[m.i, 2], ;
                HashKeys WITH aClusters[m.i, 3], ;
                ClustSize WITH aClusters[m.i, 4], ;
                EXPORT WITH .T., ;
                Exported WITH .F.

        ENDFOR

        SELECT ( lnOldArea )



    FUNCTION AnalyzeFields
    	LPARAMETERS llAnalizeAllTables as Logical && JEI RKR 06.12.2005 Add for analize fields

        LOCAL lcEnum_Fields, lnOldArea, lcEnum_Tables, lcTableName, lcSourceDB, ;
            I, lnTableCount, lnListIndex, lnListIndexTemp, lcCursorName, lcErrMsg

* If we have an extension object and it has a AnalyzeFields method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'AnalyzeFields', 5 ) and ;
			not this.oExtension.AnalyzeFields( llAnalizeAllTables, This )
			return
		ENDIF

        *
        * this creates the enumerates all the fields, their types etc.
        * in the tables selected to be upsized
        *
        * gets called when the tables selected for upsizing have changed

        lcEnum_Tables=this.EnumTablesTbl
        lnOldArea=SELECT()

        SELECT COUNT( * ) FROM ( lcEnum_Tables ) WHERE FldsAnald=.F. AND EXPORT=.T. ;
            INTO ARRAY aTableCount
        IF aTableCount=0 THEN
            RETURN
        ENDIF

        *Tell the user what's going on
        IF this.ProcessingOutput or this.lQuiet
            lcMsg=STRTRAN( ANALYZING_FIELDS_LOC, "..." )
            this.InitTherm( lcMsg, aTableCount, 0 )
        ELSE
            WAIT ANALYZING_FIELDS_LOC WINDOW NOWAIT
        ENDIF
        lnTableCount=0

        *Create table for fields if it doesn't exist yet
        IF RTRIM( this.EnumFieldsTbl )=="" THEN
            lcEnum_Fields=this.CreateWzTable( "Fields" )
            this.EnumFieldsTbl=lcEnum_Fields
            llCreateIndexes=.T.
        ELSE
            lcEnum_Fields=RTRIM( this.EnumFieldsTbl )
            llCreateIndexes=.F.
        ENDIF

        *only look at tables that haven't been crunched through this procedure before
        SELECT ( lcEnum_Tables )
        SCAN FOR FldsAnald=.F. AND ( EXPORT=.T. OR llAnalizeAllTables )
            lcCursorName=RTRIM( &lcEnum_Tables..CursName )
            lcTableName=RTRIM( &lcEnum_Tables..TblName )
            lcMsg=STRTRAN( THIS_TABLE_LOC, "|1", lcTableName )
            this.UpDateTherm( lnTableCount, lcMsg )
            SELECT ( lcCursorName )
            =AFIELDS( aFldArray )

            * test for number of fields and record size( for SQL Server )
            IF this.SQLServer
                lcErrMsg = ""

                * rmk - 01/06/2004 - adjust limits for SQL Server 7.0 and above
		        IF this.ServerVer>=7
	                IF ALEN( aFldArray, 1 )> 1024
	                    lcErrMsg = STRTRAN( CANT_UPSIZE_LOC , "|1", lcTableName )+ EXCEED_FIELDS_LOC
	                ENDIF
	                IF RECSIZE() > 8060
	                    IF !EMPTY( lcErrMsg )
	                        lcErrMsg = lcErrMsg + " and " + EXCEED_RECSIZE_LOC
	                    ELSE
	                        lcErrMsg = STRTRAN( CANT_UPSIZE_LOC , "|1", lcTableName )+ EXCEED_RECSIZE_LOC
	                    ENDIF
	                ENDIF
		        ELSE
	                IF ALEN( aFldArray, 1 )> 250
	                    lcErrMsg = STRTRAN( CANT_UPSIZE_LOC , "|1", lcTableName )+ EXCEED_FIELDS_LOC
	                ENDIF
	                IF RECSIZE() > 1962
	                    IF !EMPTY( lcErrMsg )
	                        lcErrMsg = lcErrMsg + " and " + EXCEED_RECSIZE_LOC
	                    ELSE
	                        lcErrMsg = STRTRAN( CANT_UPSIZE_LOC , "|1", lcTableName )+ EXCEED_RECSIZE_LOC
	                    ENDIF
	                ENDIF
	            ENDIF
                IF !EMPTY( lcErrMsg ) and not this.lQuiet
                    =MESSAGEBOX( lcErrMsg, ICON_EXCLAMATION, TITLE_TEXT_LOC )
                ENDIF
            ENDIF

            * 11/02/02 JVF Added: AutoInNext and AutoInStep o account
            * for VFP 8.0 autoinc attrib. Using Afields( i, 18 ) Step Value > 0 to determine
            * if auto inc column.
            SELECT ( lcEnum_Fields )
            FOR I=1 TO ALEN( aFldArray, 1 )
                APPEND BLANK

                lcFldName=LOWER( aFldArray( I, 1 ))
                REPLACE TblName WITH lcTableName, ;
                    FldName WITH lcFldName, ;
                    DATATYPE WITH aFldArray( I, 2 ), ;
                    LENGTH WITH aFldArray( I, 3 ) ;
                    PRECISION WITH aFldArray( I, 4 ), ;
                    RmtFldname WITH this.RemotizeName( lcFldName ), ;
                    RmtLength WITH aFldArray( I, 3 ) ;
                    RmtPrec WITH aFldArray( I, 4 ), ;
                    lclnull WITH aFldArray( I, 5 ), ;
                    RmtNull WITH aFldArray( I, 5 ) AND !( (aFldArray( I, 17 ) <> 0 ) OR ( aFldArray( I, 18 ) <> 0 )), ; && JEI RKR 2005.03.28 From aFldArray( I, 5 )
                    NOCPTRANS WITH aFldArray( I, 6 ), ;
                    AutoInNext WITH aFldArray( I, 17 ), ;
                    AutoInStep WITH aFldArray( I, 18 )
            NEXT I

            SELECT ( lcEnum_Tables )
            * set default mapping
            this.DefaultMapping( lcTableName )
            * set TimeStamp default ( M, G, P )
            REPLACE TStampAdd WITH ( this.SQLServer AND this.AddTimeStamp( TblName ))
            REPLACE FldsAnald WITH .T.

            IF this.ProcessingOutput THEN
                lnTableCount=lnTableCount+1
                this.UpDateTherm( lnTableCount )
            ENDIF
        ENDSCAN

        SELECT ( lcEnum_Fields )

        *Only do this the first time through
        *IF llCreateIndexes THEN
        *	INDEX ON TblName TAG TblName
        *	INDEX ON FldName TAG FldName
        *	SET ORDER TO
        *ENDIF

        *Deal with thermometer or wait window
        IF this.ProcessingOutput or this.lQuiet
			raiseevent( This, 'CompleteProcess' )
        ELSE
            WAIT CLEAR
        ENDIF

        this.AnalyzeFieldsRecalc=.F.
        SELECT ( lnOldArea )



    FUNCTION DefaultMapping
        PARAMETERS lcTableName
        LOCAL lnOldArea, aDefaultMapping, lcEnum_Fields, lnLength, lnPrecision, ;
            lcTypeString, llSkippedFirst, lcNocpType, lcRemoteType

        lnOldArea=SELECT()
        lcEnum_Fields=RTRIM( this.EnumFieldsTbl )
        *grab array of data types for setting default datatypes
        DIMENSION aDefaultMapping( 11, 4 )
        this.GetDefaultMapping( @aDefaultMapping )
        *Columns in aDefaultMapping
        *Column 1: LocalType ( local FoxPro data type )
        *Column 2: RemoteType ( default server data type )
        *Column 3: VarLength ( whether the data type is variable length or not )
        *Column 4: FullLocal	( full name of FP data type, e.g. "C"->"character" )

        *Go through the fields and stick in default data type

        SELECT ( lcEnum_Fields )
        FOR I=1 TO ALEN( aDefaultMapping, 1 )

            *If the remote datatype doesn't take a length argument, put 0 in there.
            *Otherwise, put a 1 in and then transfer the length and precision values of the local type.
            *The 0 will cause the field to be ignored by the CreateTableSQL routine

            IF aDefaultMapping( I, 3 )=.T. THEN
                lnLength=1
                lnPrecision=1
            ELSE
                lnLength=0
                lnPrecision=0
            ENDIF

            REPLACE RmtType WITH aDefaultMapping( I, 2 ), ;
                FullType WITH aDefaultMapping( I, 4 ) ;
                RmtLength WITH lnLength, RmtPrec WITH lnPrecision ;
                FOR DATATYPE = aDefaultMapping( I, 1 ) AND RTRIM( TblName ) == lcTableName
        NEXT

        * Set up ComboType field of the type mapping grid
        * create a string like "character ( 14 )" or "numeric ( 3, 2 )" or "memo" or "memo-nocp"
        SCAN FOR RTRIM( TblName ) == lcTableName
            * change default FullType for NoCptrans

            * rmk - 01/06/2004
            * lcTypeNocp = IIF( DATATYPE = 'C', 'char_nocp', 'memo_nocp' )
            DO CASE
            CASE DATATYPE = 'C'
	            lcTypeNocp = 'char_nocp'
            CASE DATATYPE = 'V'
	            lcTypeNocp = 'varchar_nocp'
            OTHERWISE
            	lcTypeNocp = 'memo_nocp'
			ENDCASE

            lcTypeString = IIF( NOCPTRANS, lcTypeNocp, RTRIM( FullType ))

            * add field length and decimals for Fox variable length types
            IF INLIST( DATATYPE, 'C', 'V', 'Q', 'N', 'F' )
                lcTypeString = lcTypeString + " ( " + LTRIM( STR( LENGTH ))
                IF PRECISION <> 0 THEN
                    lcTypeString = lcTypeString + ", " + LTRIM( STR( PRECISION ))
                ENDIF
                lcTypeString = lcTypeString+" )"
            ELSE
                * JVF 11/02/02 Add AutoInc string to Local type
                IF DATATYPE = "I" AND this.SQLServer
                    IF AutoInStep > 0
                        lcTypeString = lcTypeString + " ( AutoInc )"

                        * JVF 11/02/02 # 1478 Add "( Identity )" string to RemoteType if AutoInc,
                        REPLACE RmtType WITH "int ( Ident )"
                    ELSE
                        * Since we added record it to typemap, strip it off if necc.
                        REPLACE RmtType WITH "int"
                    ENDIF
                ENDIF
            ENDIF
            REPLACE ComboType WITH lcTypeString
        ENDSCAN

        * Set up the RemLength, RemPrec fields of the type mapping grid
        * Replace all the 1s in the Rmtlength field with local length and precision values
        * jvf 08/16/99
        SCAN FOR RmtLength <> 0 AND RTRIM( TblName ) == lcTableName
            IF PRECISION <> 0 AND DATATYPE = 'N'
                REPLACE RmtLength WITH LENGTH-1, RmtPrec WITH PRECISION
            ELSE
                REPLACE RmtLength WITH LENGTH, RmtPrec WITH PRECISION
            ENDIF
        ENDSCAN

        * implement the server-specific cases for RmtType, Rmtlength, RmtPrec
        IF this.SQLServer THEN
            REPLACE ALL RmtType WITH "char" ;
                FOR DATATYPE = "C" AND NOCPTRANS AND RTRIM( TblName )==lcTableName
            * jvf 1/8/01 RmtType s/b Text, not Image, for Binary Memos - Bug ID 45752
            REPLACE ALL RmtType WITH "text" ;
                FOR DATATYPE = "M" AND NOCPTRANS AND RTRIM( TblName )==lcTableName
        ENDIF

        * If we're upsizing to Oracle, need to put default values in for money and logical types
        * Same for int when converted to numeric
        IF this.ServerType == ORACLE_SERVER THEN

            REPLACE ALL RmtType WITH "raw" ;
                FOR DATATYPE = "C" AND NOCPTRANS AND RTRIM( TblName )==lcTableName

            REPLACE ALL RmtType WITH "long raw" ;
                FOR DATATYPE = "M" AND NOCPTRANS AND RTRIM( TblName )==lcTableName

            REPLACE ALL RmtLength WITH 19, RmtPrec WITH 4 ;
                FOR DATATYPE="Y" AND RTRIM( TblName )==lcTableName

            REPLACE ALL RmtLength WITH 1, RmtPrec WITH 0 ;
                FOR DATATYPE="L" AND RTRIM( TblName )==lcTableName

            REPLACE ALL RmtLength WITH 11, RmtPrec WITH 0 ;
                FOR DATATYPE="I" AND RTRIM( TblName )==lcTableName

            *Oracle only allows one LONG or LONG RAW column per table; change all
            *but the first LONG RAW, ie General, to RAW( 255 ); change extra LONG fields
            *to VARCHAR2( 2000 )

            *This handles general fields
            LOCATE FOR RTRIM( RmtType )=="long raw" AND RTRIM( TblName )==lcTableName
            DO WHILE FOUND()
                IF llSkippedFirst THEN
                    REPLACE RmtType WITH "raw", RmtLength WITH 255
                ELSE
                    llSkippedFirst=.T.
                ENDIF
                CONTINUE
            ENDDO

            *This handles memo fields
            LOCATE FOR RTRIM( RmtType )=="long" AND RTRIM( TblName )==lcTableName
            DO WHILE FOUND()
                IF llSkippedFirst THEN
                    REPLACE RmtType WITH "varchar2", RmtLength WITH 2000
                ELSE
                    llSkippedFirst=.T.
                ENDIF
                CONTINUE
            ENDDO
        ENDIF

        SELECT ( lnOldArea )



    FUNCTION GetEligibleTables
        PARAMETERS aTableArray
        LOCAL lcEnumTables, lnOldArea, I

        lnOldArea=SELECT()
        lcEnumTables=this.EnumTablesTbl
        SELECT ( lcEnumTables )
        SET FILTER TO Upsizable=.T.
        GO TOP
        I=1
        IF EOF()
            RETURN 0
        ELSE
            SCAN
                IF !EMPTY( aTableArray[1] ) THEN
                    DIMENSION aTableArray[ALEN( aTableArray, 1 )+1]
                ENDIF
                aTableArray[i]=TblName
                I=I+1
            ENDSCAN
        ENDIF
        SELECT ( lnOldArea )
        RETURN



    FUNCTION GetEligibleClusterTables
        PARAMETERS aTableArray, lcSelectClustName
        LOCAL lcEnumTables, lnOldArea, I, lcFilter

        * if EOF() the array is unchanged
        DIMENSION aTableArray[1]
        aTableArray[1] = ""
        lnOldArea=SELECT()
        lcEnumTables=this.EnumTablesTbl
        SELECT ( lcEnumTables )

        lcSelectClustName = TRIM( lcSelectClustName )
        IF lcSelectClustName = ""
            lcFilter = "Export = .T. AND EMPTY( ClustName )"
        ELSE
            lcFilter = "Export = .T. AND TRIM( ClustName ) = lcSelectClustName"
        ENDIF

        GO TOP
        I=1
        IF EOF()
            RETURN 0
        ELSE
            SCAN FOR &lcFilter
                IF !EMPTY( aTableArray[1] ) THEN
                    DIMENSION aTableArray[ALEN( aTableArray, 1 )+1]
                ENDIF
                aTableArray[i] = TRIM( TblName )
                I=I+1
            ENDSCAN
        ENDIF
        SELECT ( lnOldArea )
        RETURN



    FUNCTION GetEligibleTableFields
        PARAMETERS aFieldArray, lcTableName, llKeys
        LOCAL lcEnumFields, lnOldArea, I, lcFilter

        DIMENSION aFieldArray[1]
        aFieldArray[1] = ""
        lnOldArea = SELECT()
        lcEnumFields = this.EnumFieldsTbl
        SELECT ( lcEnumFields )

        lcTableName = TRIM( lcTableName )
        IF llKeys
            lcFilter = "TblName = lcTableName  AND !EMPTY( ClustOrder )"
        ELSE
            lcFilter = "TblName = lcTableName  AND EMPTY( ClustOrder )"
        ENDIF

        GO TOP
        I = 1
        IF EOF()
            RETURN 0
        ELSE
            SCAN FOR &lcFilter
                IF !EMPTY( aFieldArray[1] ) THEN
                    DIMENSION aFieldArray[ALEN( aFieldArray, 1 )+1]
                ENDIF
                aFieldArray[i] = TRIM( FldName )
                I = I + 1
            ENDSCAN
        ENDIF
        SELECT ( lnOldArea )
        RETURN


    FUNCTION GetInfoTableFields
        PARAMETERS aInfoFieldArray, lcTableName, llKeys
        LOCAL lcEnumFields, lnOldArea, I, lcFilter

        DIMENSION aInfoFieldArray[1, 5]
        aInfoFieldArray[1, 1] = ""
        lnOldArea = SELECT()
        lcEnumFields = this.EnumFieldsTbl
        SELECT ( lcEnumFields )
        lcTableName = TRIM( lcTableName )
        IF llKeys
            lcFilter = "TblName = lcTableName  AND !EMPTY( ClustOrder )"
        ELSE
            lcFilter = "TblName = lcTableName  AND EMPTY( ClustOrder )"
        ENDIF

        GO TOP
        I = 1
        IF EOF()
            RETURN 0
        ELSE
            SCAN FOR &lcFilter
                IF !EMPTY( aInfoFieldArray[1, 1] )
                    DIMENSION aInfoFieldArray[ALEN( aInfoFieldArray, 1 )+1, 5]
                ENDIF
                aInfoFieldArray[i, 1] = TRIM( FldName )
                aInfoFieldArray[i, 2] = TRIM( RmtType )
                aInfoFieldArray[i, 3] = RmtLength
                aInfoFieldArray[i, 4] = RmtPrec
                IF llKeys
                    aInfoFieldArray[i, 5] = ClustOrder
                ELSE
                    aInfoFieldArray[i, 5] = .T.
                ENDIF
                I = I + 1
            ENDSCAN
        ENDIF
        =ASORT( aInfoFieldArray, 5 )
        SELECT ( lnOldArea )
        RETURN



    * verify if the tables selected in cluster have common fields

    FUNCTION VerifyClusterTablesFields
        PARAMETERS aClusterTables
        LOCAL llValid, aInfoTable1Fields[1, 5], aInfoTableFields[1, 5]

        * False if no table selected
        IF EMPTY( aClusterTables[1, 1] )
            RETURN .F.
        ENDIF

        this.GetInfoTableFields( @aInfoTable1Fields, aClusterTables[1], .F. )

        * True if we have a single cluster with at least one key
        IF ALEN( aClusterTables, 1 ) = 1
            RETURN IIF( EMPTY( aClusterTables[1, 1] ), .F., .T. )
        ENDIF

        * we have at least two tables here and first table has at least a field selected
        FOR m.I = 2 TO ALEN( aClusterTables, 1 )
            this.GetInfoTableFields( @aInfoTableFields, aClusterTables[m.i], .F. )
            FOR m.j = 1 TO ALEN( aInfoTable1Fields, 1 )
                IF aInfoTable1Fields[m.j, 5]
                    FOR m.k = 1 TO ALEN( aInfoTableFields, 1 )
                        llValid = .F.
                        IF ( aInfoTable1Fields[m.j, 2] = aInfoTableFields[m.k, 2] AND ;
                                aInfoTable1Fields[m.j, 3] = aInfoTableFields[m.k, 3] AND ;
                                aInfoTable1Fields[m.j, 4] = aInfoTableFields[m.k, 4] )
                            llValid = .T.
                            EXIT
                        ENDIF
                    ENDFOR
                    aInfoTable1Fields[m.j, 5] = llValid
                ENDIF
            ENDFOR
        ENDFOR

        RETURN ASCAN( aInfoTable1Fields, .T. ) > 0




    FUNCTION VerifyClusterKeyFields
        PARAMETERS aClusterTables
        LOCAL llValid, aInfoTable1Fields[1, 5], aInfoTableFields[1, 5]

        * verify if the fields selected in each table( part of the cluster ) are the same

        * False if no clusters
        IF EMPTY( aClusterTables[1, 1] )
            RETURN .F.
        ENDIF

        this.GetInfoTableFields( @aInfoTable1Fields, aClusterTables[1, 1], .T. )

        * True if we have a single table with a valid key
        IF ALEN( aClusterTables, 1 ) = 1
            RETURN IIF( EMPTY( aInfoTable1Fields[1, 1] ), .F., .T. )
        ENDIF

        * we have at least two tables here and first table has a valid key
        FOR m.I = 2 TO ALEN( aClusterTables, 1 )
            this.GetInfoTableFields( @aInfoTableFields, aClusterTables[m.i], .T. )

            * keys for both clusters should match in number of fields, type and size
            IF EMPTY( aInfoTableFields[1, 1] ) OR ;
                    ALEN( aInfoTableFields, 1 ) != ALEN( aInfoTable1Fields, 1 )
                RETURN .F.
            ENDIF

            FOR m.j = 1 TO ALEN( aInfoTable1Fields, 1 )
                IF ( aInfoTable1Fields[m.j, 2] = aInfoTableFields[m.j, 2] AND ;
                        aInfoTable1Fields[m.j, 3] = aInfoTableFields[m.j, 3] AND ;
                        aInfoTable1Fields[m.j, 4] = aInfoTableFields[m.j, 4] )
                    LOOP
                ELSE
                    RETURN .F.
                ENDIF
            ENDFOR
        ENDFOR

        RETURN .T.



    FUNCTION GetDefaultMapping
        PARAMETERS aPassedArray
        LOCAL lnOldArea, lcServerConstraint

        lnOldArea=SELECT()
        IF NOT USED( "TypeMap" )
            SELECT 0
*** DH 09/05/2013: don't open exclusively
***            USE TypeMap EXCLUSIVE
            USE TypeMap
        ELSE
            SELECT TypeMap
        ENDIF

        *Didn't foresee a problem, thus this cheezy snippet
*** DH 09/05/2013: use only SQL Server rather than variants of it
***        IF this.ServerType=="SQL Server95" THEN
		IF left( this.ServerType, 10 ) = 'SQL Server'
            lcServerConstraint="SQL Server"
        ELSE
***            IF this.ServerType=="SQL Server" THEN
***                lcServerConstraint="SQL Server4x"
***            ELSE
                lcServerConstraint=RTRIM( this.ServerType )
***            ENDIF
        ENDIF

        SELECT LocalType, RemoteType, VarLength, FullLocal FROM TypeMap ;
            WHERE  TypeMap.DEFAULT=.T. AND TypeMap.SERVER=lcServerConstraint ;
            INTO ARRAY aPassedArray

        SELECT( lnOldArea )



    #IF SUPPORT_ORACLE
    FUNCTION DealWithTypeLong
        LOCAL lcEnumFieldsTbl, lnOldArea

        *
        *Oracle tables only allow one field to be Long or LongRaw; this warns
        *the user about the problem.  The DefaultMapping routine deals with it
        *

        this.AnalyzeFields

        lcEnumFieldsTbl=this.EnumFieldsTbl
        lnOldArea=SELECT()
        SELECT 0
        lcCursor=this.UniqueCursorName( "_foo" )
        SELECT COUNT( * ) FROM ( lcEnumFieldsTbl ) ;
            WHERE RTRIM( DATATYPE )=="M" OR ;
            RTRIM( DATATYPE )=="G" OR ;
            RTRIM( DATATYPE )=="P" ;
            GROUP BY TblName ;
            INTO CURSOR &lcCursor
        SELECT COUNT( * ) FROM ( lcCursor ) WHERE CNT>1 INTO ARRAY aMemoCount
        USE

        IF aMemoCount>0 THEN
            lcMsg=STRTRAN( LONG_TYPE_LOC, '|1', LTRIM( STR( aMemoCount )) )
            lcMsg=STRTRAN( lcMsg, '|2', IIF( aMemoCount>1, TABLES_HAVE_LOC, TABLE_HAS_LOC ))
        	IF not this.lQuiet
            	MESSAGEBOX( lcMsg, ICON_EXCLAMATION, TITLE_TEXT_LOC )
        	ENDIF
        ENDIF

        SELECT ( lnOldArea )

#ENDIF


#IF SUPPORT_ORACLE
    FUNCTION TwoLongs
        PARAMETERS lcTableName, lcFirstLong, lcOtherLong
        LOCAL lcEnumFieldsTbl, lnOldArea

        *Checks to see if a field in a table already has a type of Long or Long Raw
        *Returns the name of the field if there is one

        lcEnumFieldsTbl=this.EnumFieldsTbl
        lnOldArea=SELECT()
        SELECT ( lcEnumFieldsTbl )
        LOCATE FOR RTRIM( TblName )==lcTableName AND RmtType="long"
        IF RTRIM( FldName )==lcFirstLong THEN
            CONTINUE
            IF FOUND() THEN
                lcOtherLong=FldName
            ENDIF
        ELSE
            lcOtherLong=FldName
        ENDIF

        SELECT ( lnOldArea )
        RETURN !EMPTY( lcOtherLong )

#ENDIF


    FUNCTION Upsizable
        PARAMETERS lcTableName, lcTablePath, llAlreadyOpen, lcCursorName, aOpenTables
        LOCAL lnOldArea, I, lcNewCursName

        *
        *This function checks to see that a table is actually marked as part of the
        *selected database
        *
        *It also opens all tables exclusively if they aren't already
        *

        *Substitute underscores for any spaces ( as FoxPro does )

        *See if the table is already open, possibly with an alias different from the table name
        IF !EMPTY( aOpenTables ) THEN
            FOR I=1 TO ALEN( aOpenTables, 1 )
                IF DBF( aOpenTables[i, 2] )==RTRIM( UPPER( lcTablePath )) THEN
                    lcCursorName=aOpenTables[i, 1]
                    EXIT
                ENDIF
            NEXT
        ENDIF

        *If it's not open already, handle table names with spaces
        IF EMPTY( lcCursorName ) THEN
            lcCursorName=RTRIM( STRTRAN( lcTableName, CHR( 32 ), "_" ))
            *Handle the case of table name being an important Fox keyword
            *Note the base wizard class ensures that no tables are already open
            *with these keywords, so we only worry about opening them here
            IF INLIST( UPPER( lcCursorName ), "THIS", "THISFORMSET" )
                lcCursorName=LEFT( lcCursorName, MAX_FIELDNAME_LEN-1 )+"_"
            ENDIF
        ENDIF

        lnOldArea=SELECT()
        IF !FILE( lcTablePath ) THEN
            SELECT ( lnOldArea )
            RETURN .F.
        ENDIF
        IF !USED( lcCursorName ) THEN
            this.SetErrorOff=.T.
            this.HadError=.F.
            llAlreadyOpened=.F.
            SELECT 0
            USE ( lcTableName ) ALIAS ( lcCursorName ) EXCLUSIVE
            this.SetErrorOff=.F.
            IF this.HadError
                SELECT ( lnOldArea )
                RETURN .F.
            ENDIF
        ELSE
            *Make sure that if a table's open, it belongs to the database
            *to be upsized
            SELECT ( lcCursorName )
            IF !LOWER( CURSORGETPROP( 'database' ))==LOWER( ALLTRIM( this.SourceDB )) THEN
                lcCursorName=this.UniqueTorVName( "Namewithmanycharacters" )
                this.SetErrorOff=.T.
                this.HadError=.F.
                SELECT 0
                USE ( lcTableName ) EXCLUSIVE ALIAS ( lcCursorName )
                this.SetErrorOff=.F.
                IF this.HadError
                    USE
                    SELECT ( lnOldArea )
                    RETURN .F.
                ENDIF
            ELSE
                llAlreadyOpened=.T.
            ENDIF

            *If it's open, make sure it's open exclusive
            SELECT ( lcCursorName )
            IF !ISFLOCKED()
                USE
                this.SetErrorOff=.T.
                this.HadError=.F.
                USE ( lcTableName ) EXCLUSIVE
                this.SetErrorOff=.F.
                IF this.HadError
                    USE ( lcTableName ) SHARED
                    SELECT ( lnOldArea )
                    RETURN .F.
                ENDIF
            ENDIF

        ENDIF

        SELECT ( lnOldArea )



    #IF SUPPORT_ORACLE
    FUNCTION RemoveCluster
        PARAMETERS lcClusterName
        LOCAL lcClusterNamesTbl, lcClusterKeysTbl, lcEnumTablesTbl, lnOldArea

        *Called by "Remove" button on Create Cluster screen and by Table Selection screen

        lcClusterNamesTbl=this.ClusterNamesTbl
        lcClusterKeysTbl=this.ClusterKeysTbl
        lcEnumTablesTbl=this.EnumTablesTbl
        lnOldArea=SELECT()

        *delete the record for the cluster from the cluster names table
        SELECT ( lcClusterNamesTbl )
        DELETE ALL FOR &lcClusterNamesTbl..ClustName=lcClusterName

        *delete any related records in the cluster keys table
        SELECT ( lcClusterKeysTbl )
        DELETE ALL FOR &lcClusterKeysTbl..ClustName=lcClusterName

        *Mark any tables that were in the cluster as available
        SELECT ( lcEnumTablesTbl )
        REPLACE ALL &lcEnumTablesTbl..lcClusterName WITH ""

        SELECT ( lnOldArea )

#ENDIF


#IF SUPPORT_ORACLE
    FUNCTION ChangeClusterStatus
        PARAMETERS lcRel, lcClustName, lcClustType, lnHashKeys
        LOCAL lcEnumTables, lnOldArea, lcEnumIndexesTbl, aClustTables, lcParent, ;
            lcChild, lnDupeID

        *Called from cluster creation page when user adds or removes a cluster

        lnOldArea=SELECT()
        lcEnumTables=this.EnumTablesTbl
        lcEnumIndexesTbl=this.EnumIndexesTbl
        lcEnumRelsTbl=this.EnumRelsTbl
        IF EMPTY( lnHashKeys ) THEN
            lnHashKeys=0
        ENDIF

        *Parse the relation
        lcParent=""
        lcChild=""
        lnDupeID=0
        this.ParseRel( lcRel, @lcParent, @lcChild, @lnDupeID )
        SELECT ( lcEnumRelsTbl )

        *If the clustertype wasn't passed, the cluster is being added or deleted
        *so see if the name is already in use
        IF TYPE( "lcClustType" )="L" THEN
            IF !lcClustName=="" THEN
                SET ORDER TO ClustName
                SEEK RTRIM( lcClustName )
                IF FOUND() THEN
                    *Give error message if cluster name already exists
                    lcMessage=STRTRAN( DUP_CLUSTNAME_LOC, "|1", RTRIM( this.UserInput ))
		        	IF this.lQuiet
		        		.HadError     = .T.
		        		.ErrorMessage = lcMessage
					ELSE
	                    MESSAGEBOX( lcMessage, ICON_EXCLAMATION, TITLE_TEXT_LOC )
		        	ENDIF
                    SET ORDER TO
                    RETURN .F.
                ENDIF
                SET ORDER TO
            ENDIF
        ENDIF

        *If it's not in use or we're just changing the cluster type, find the right record
        *and toss the values in

        LOCATE FOR DD_CHILD=lcChild AND DD_PARENT=lcParent AND Duplicates=lnDupeID
        IF TYPE( "lcClustType" )="L" THEN
            lcClustType=IIF( lcClustName=="", "", "INDEX" )
            REPLACE ClustName WITH lcClustName, ClustType WITH lcClustType, ;
                HashKeys WITH lnHashKeys

            *Now associate the cluster name ( "" if table is being removed )
            *with the tables in the cluster and store default cluster type of " INDEX"
            SELECT ( lcEnumTables )
            REPLACE &lcEnumTables..ClustName WITH lcClustName FOR TblName=lcParent OR TblName=lcChild
        ELSE
            REPLACE ClustType WITH lcClustType, HashKeys WITH lnHashKeys
        ENDIF

        SELECT ( lnOldArea )

#ENDIF


#IF SUPPORT_ORACLE
    FUNCTION RemoveTableFromClust
        PARAMETERS lcClustName

        *Called by page 6 when a user deselects a table that was going to be exported

        LOCAL lcEnumRelsTbl, lcEnumTablesTbl, lnOldArea, lnRecNo
        lcEnumTablesTbl=this.EnumTablesTbl
        lcEnumRelsTbl=this.EnumRelsTbl
        lnOldArea=SELECT()

        SELECT ( lcEnumTablesTbl )
        lnRecNo=RECNO()
        REPLACE ClustName WITH "" FOR RTRIM( ClustName )==RTRIM( lcClustName )
        GO lnRecNo

        SELECT ( lcEnumRelsTbl )
        REPLACE ClustName WITH "", ClustType WITH "" FOR RTRIM( ClustName )==RTRIM( lcClustName )

        SELECT ( lnOldArea )

#ENDIF


#IF SUPPORT_ORACLE
    FUNCTION InOneCluster
        PARAMETERS lcRel
        LOCAL lcParent, lcChild, lnDupeID, lcEnumTablesTbl, lcMsg

        *
        *Checked when the user creates a cluster; ensures that a given table is only
        *in one relation
        *

        *Parse the relation
        lcParent=""
        lcChild=""
        lnDupeID=0
        this.ParseRel( lcRel, @lcParent, @lcChild, @lnDupeID )

        *See if the tables are already in a cluster
        lcEnumTablesTbl=this.EnumTablesTbl
        SELECT ( lcEnumTablesTbl )

        FOR I=1 TO 2
            LOCATE FOR RTRIM( TblName )==lcParent
            IF !EMPTY( ClustName ) THEN
                lcMsg=STRTRAN( ONE_CLUSTER_LOC, "|1", lcParent )
                lcMsg=STRTRAN( lcMsg, "|2", RTRIM( ClustName ))
	        	IF this.lQuiet
	        		.HadError     = .T.
	        		.ErrorMessage = lcMsg
				ELSE
	                MESSAGEBOX( lcMsg, 48, TITLE_TEXT_LOC )
	        	ENDIF
                RETURN .F.
            ENDIF
            lcParent=lcChild
        NEXT

#ENDIF


    FUNCTION InIndex
        PARAMETERS aIndexes, lcFldName, lcTableName
        LOCAL lcEnumIndexesTbl, lnOldArea, lcTagName, llInIndex

        *
        *Returns array of tag names where a given field is part of the tag expression
        *

        lcEnumIndexesTbl=this.EnumIndexesTbl
        lnOldArea=SELECT()
        SELECT ( lcEnumIndexesTbl )
        LOCATE FOR RTRIM( IndexName )=lcTableName
        llInIndex=.F.
        DO WHILE FOUND()
            IF lcFldName $ LclExpr THEN
                lcTagName=RTRIM( TagName )
                this.InsaItem( @aIndexes, lcTagName )
                llInIndex=.T.
            ENDIF
            CONTINUE
        ENDDO

        SELECT ( lnOldArea )
        RETURN llInIndex



    FUNCTION INKEY
        PARAMETERS lcRmtFieldName, lcTableName
        LOCAL lcEnumRelsTbl, lnOldArea, aRels

        *
        *Checks to see if a given field is in a key ( primary or foreign )
        *

        lcTableName=LOWER( lcTableName )
        lcEnumRelsTbl=this.EnumRelsTbl
        lnOldArea=SELECT()
        SELECT ( lcEnumRelsTbl )
        LOCATE FOR RTRIM( DD_PARENT )==lcTableName OR RTRIM( DD_CHILD )==lcTableName
        llInKey=.F.
        DO WHILE FOUND()
            IF lcRmtFieldName $ DD_CHIEXPR
                lcRelatedTable=IIF( RTRIM( DD_PARENT )=lcTableName, RTRIM( DD_CHILD ), RTRIM( DD_PARENT ))
                llInKey=.T.
                EXIT
            ENDIF
            CONTINUE
        ENDDO

        SELECT ( lnOldArea )
        RETURN llInKey



    FUNCTION DontIndex
        PARAMETERS lcFieldName, lcTableName
        LOCAL lnOldArea, lcEnumIndexsTbl

        *
        *If user changed data type to something unindexable, don't create the indexes
        *that include the unindexable field
        *

        lnOldArea=SELECT()
        lcEnumIndexsTbl=this.EnumIndexesTbl
        SELECT ( lcEnumIndexsTbl )
        LOCATE FOR RTRIM( IndexName )==RTRIM( lcTableName ) AND lcFieldName $ LclExpr
        DO WHILE FOUND()
            REPLACE IdxError WITH STRTRAN( IDX_NOT_CREATED_LOC, "|1", lcFieldName ), ;
                Exported WITH .F., ;
                DontCreate WITH .T.
            CONTINUE
        ENDDO

        SELECT ( lnOldArea )



    FUNCTION DefaultsAndRules

        *
        *This proc converts FoxPro defaults and rules to server equivalents
        *
        *In the case of SQL Server, defaults are converted to defaults and rules to stored
        *procedures which are then called from insert and update triggers
        *
        *In the case of Oracle, defaults become ALTER TABLE statements and rules are
        *converted to SQL statements that will wind up in one trigger that executes on
        *the update or insert event
        *
        * If llOraFieldRules, create Oracle field rules to become part of CREATE TABLE

        LOCAL lcEnumTables, lcTableName, lcTableRule, lcRmtTableName, ;
            lcRuleExpression, lcRemoteRule, lcRuleText, lcFldName, ;
            lcDefaultExpression, lcRemoteDefault, lcEnumFields, llRuleSprocCreated, ;
            llDefaultCreated, llDefaultBound, llTableExported, lcRemoteRuleName, ;
            lcRemoteDefaultName, lcFldType, lcRuleError, lcDefaError, lcRmtFldName, ;
            lcConstName, lnTableCount, lcThermMsg, lnError, llShowTherm, llRuleCreated

* If we have an extension object and it has a DefaultsAndRules method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'DefaultsAndRules', 5 ) and ;
			not this.oExtension.DefaultsAndRules( This )
			return
		ENDIF

        lcEnumTables = this.EnumTablesTbl
        SELECT ( lcEnumTables )

        * thermometer
        SELECT COUNT( * ) FROM ( lcEnumTables ) WHERE EXPORT = .T. INTO ARRAY aTableCount
        IF this.ExportValidation THEN
            lcThermMsg =  CONVERT_RULE_LOC
        ENDIF
        IF this.ExportDefaults THEN
            IF EMPTY( lcThermMsg ) THEN
                lcThermMsg = CONVERT_DEFAS_LOC
            ELSE
                lcThermMsg = lcThermMsg + AND_LOC + CONVERT_DEFAS_LOC
            ENDIF
        ENDIF
        IF !this.ExportDefaults AND !this.ExportValidation THEN
            IF this.ServerType = "Oracle" THEN
                RETURN
            ELSE
                lcThermMsg = BIND_DEFAS_LOC
            ENDIF
        ELSE
            lcThermMsg = CONVERT_STEM_LOC + lcThermMsg
        ENDIF
        this.InitTherm( lcThermMsg, aTableCount, 0 )
        lnTableCount = 0

        SCAN FOR &lcEnumTables..EXPORT = .T.
            llTableExported = &lcEnumTables..Exported
            lcTableName = RTRIM( &lcEnumTables..TblName )
            lcRmtTableName = RTRIM( &lcEnumTables..RmtTblName )

            lcThermMsg = STRTRAN( THIS_TABLE_LOC, '|1', lcTableName )
            this.UpDateTherm( lnTableCount, lcThermMsg )
            lnTableCount = lnTableCount + 1

            lcRuleError=""
            lcDefaultError=""
            lcRemoteRuleName=""
            lcRemoteRule=""
            lcRuleText=""
            lcRemoteDefault=""
            llRuleSprocCreated=.F.
            llDefaultCreated=.F.
            llDefaultBound=.F.
            lcRemoteDefaultName=""
            lcDefaError=""
            lnError=.F.

            *Turn table rules into stored procedures ( SQL Server ) or constraints ( Oracle )
            *Grab the rule

            IF this.ExportValidation
                lcRuleExpression = DBGETPROP( lcTableName, "Table", "RuleExpression" )
                lcRuleText = DBGETPROP( lcTableName, "Table", "RuleText" )

                *If there is a rule, convert it
                IF !EMPTY( lcRuleExpression ) THEN

                    *For Oracle, convert rules to trigger code
                    IF this.ServerType = ORACLE_SERVER THEN
                        lcConstName = ORA_CONST_TAB_PREFIX + LEFT( lcRmtTableName, MAX_NAME_LENGTH - LEN( ORA_CONST_TAB_PREFIX ))
                        lcConstName = this.UniqueOraName( lcConstName, .T. )
                        lcRemoteRule = this.ConvertToConstraint( lcRuleExpression, lcTableName, lcRmtTableName, lcConstName )

                        * Create table constraint if user is upsizing and has create constraint permissions
                        IF this.DoUpsize AND llTableExported && AND this.Perm_default
                            IF this.ExecuteTempSPT( lcRemoteRule, @lnError, @lcDefaError )
                                lcRemoteRuleName = lcConstName
                            ELSE
                                REPLACE RuleErrNo WITH lnError
                                this.StoreError( lnError, lcRuleError, lcRemoteRule, CONVTO_TRIG_ERR_LOC, lcTableName, TABLE_RULE_LOC )
                            ENDIF
                        ENDIF

                        REPLACE &lcEnumTables..LocalRule WITH lcRuleExpression, ;
                            &lcEnumTables..RRuleName WITH lcRemoteRuleName, ;
                            &lcEnumTables..RmtRule WITH lcRemoteRule, ;
                            &lcEnumTables..RuleError WITH lcRuleError
                    ELSE

                        *For SQL Server, create a sproc
                        lcRemoteRule=this.ConvertToSproc( lcRuleExpression, lcRuleText, lcTableName, ;
                            lcRmtTableName, "Table", @lcRemoteRuleName )

                        *Create the rule if user is upsizing and has sproc permission
                        IF this.DoUpsize AND llTableExported AND this.Perm_Sproc THEN
                            *Might have to drop existing sproc
                            IF this.MaybeDrop( lcRemoteRuleName, "procedure" ) THEN
                                llRuleSprocCreated=this.ExecuteTempSPT( lcRemoteRule, @lnError, @lcRuleError )
                                IF !llRuleSprocCreated THEN
                                    REPLACE &lcEnumTables..RuleErrNo WITH lnError
                                    this.StoreError( lnError, lcRuleError, lcRemoteRule, SPROC_ERR_LOC, lcTableName, TABLE_RULE_LOC )
                                ENDIF
                            ELSE
                                lcRuleError=CANT_DROP_SPROC_LOC
                            ENDIF
                        ENDIF

                        *Store the information for the upsizing report
                        REPLACE &lcEnumTables..LocalRule WITH lcRuleExpression, ;
                            &lcEnumTables..RRuleName WITH lcRemoteRuleName, ;
                            &lcEnumTables..RmtRule WITH lcRemoteRule, ;
                            &lcEnumTables..RuleExport WITH llRuleSprocCreated, ;
                            &lcEnumTables..RuleError WITH lcRuleError

                        lcRuleError=""
                        lnError=.F.

                    ENDIF

                ENDIF

            ENDIF

            *
            *Deal with field rules: change to sprocs ( SQL Server ) or constraints ( Oracle )
            lcEnumFields=this.EnumFieldsTbl

            SELECT ( lcEnumFields )

            IF this.ExportValidation THEN

                SCAN FOR RTRIM( TblName )==lcTableName
                    lcFldName = RTRIM( &lcEnumFields..FldName )
                    lcRmtFldName = RTRIM( &lcEnumFields..RmtFldname )
                    lcRuleExpression = DBGETPROP( lcTableName+"."+lcFldName, "Field", "RuleExpression" )
                    lcRuleText = DBGETPROP( lcTableName+"."+lcFldName, "Field", "RuleText" )

                    *do nothing if there's no local rule
                    IF !EMPTY( lcRuleExpression ) THEN

                        * For Oracle, convert rules to table constraints
                        IF this.ServerType = ORACLE_SERVER
                            lcConstName = ORA_CONST_COL_PREFIX + LEFT( lcRmtFldName, MAX_NAME_LENGTH - LEN( ORA_CONST_COL_PREFIX ))
                            lcConstName = this.UniqueOraName( lcConstName, .T. )
                            lcRemoteRule = this.ConvertToConstraint( lcRuleExpression, lcTableName, lcRmtTableName, lcConstName )

                            * Create table constraint if user is upsizing and has create constraint permissions
                            IF this.DoUpsize AND llTableExported && AND this.Perm_default
                                IF this.ExecuteTempSPT( lcRemoteRule, @lnError, @lcDefaError )
                                    lcRemoteRuleName = lcConstName
                                ELSE
                                    REPLACE &lcEnumFields..RuleErrNo WITH lnError
                                    this.StoreError( lnError, lcRuleError, lcRemoteRule, CONVTO_TRIG_ERR_LOC, lcTableName+"."+lcFldName, FIELD_RULE_LOC )
                                ENDIF
                            ENDIF

                            REPLACE &lcEnumFields..LocalRule WITH lcRuleExpression, ;
                                &lcEnumFields..RRuleName WITH lcRemoteRuleName, ;
                                &lcEnumFields..RmtRule WITH lcRemoteRule, ;
                                &lcEnumFields..RuleError WITH lcRuleError
                        ELSE

                            lcRemoteRule=this.ConvertToSproc( lcRuleExpression, lcRuleText, lcTableName, ;
                                lcRmtTableName, "Field", @lcRemoteRuleName, lcRmtFldName )

                            *Create the sprocs if the user is actually upsizing and has permissions
                            IF this.DoUpsize AND llTableExported AND this.Perm_Sproc THEN

                                *Create the sproc if there's a rule
                                IF !EMPTY( lcRuleExpression ) THEN
                                    *Might have to drop existing sproc
                                    IF this.MaybeDrop( lcRemoteRuleName, "procedure" ) THEN
                                        llRuleSprocCreated=this.ExecuteTempSPT( lcRemoteRule, @lnError, @lcRuleError )
                                        IF !llRuleSprocCreated THEN
                                            REPLACE &lcEnumFields..RuleErrNo WITH lnError
                                            this.StoreError( lnError, lcRuleError, lcRemoteRule, SPROC_ERR_LOC, lcTableName+"."+lcFldName, FIELD_RULE_LOC )
                                        ENDIF
                                    ELSE
                                        lcRuleError=CANT_DROP_SPROC_LOC
                                    ENDIF
                                ENDIF

                            ENDIF

                            *Store all this stuff
                            REPLACE &lcEnumFields..LocalRule WITH lcRuleExpression, ;
                                &lcEnumFields..RmtRule WITH lcRemoteRule, ;
                                &lcEnumFields..RRuleName WITH lcRemoteRuleName, ;
                                &lcEnumFields..RuleExport WITH llRuleSprocCreated, ;
                                &lcEnumFields..RuleError WITH lcRuleError

                            lcRuleError=""
                            lnError=.F.

                        ENDIF

                    ENDIF

                ENDSCAN

            ENDIF

            * Deal with defaults ( depends on server type )
            *
            * Unlike the above code, the difference between Oracle and SQL Server is handled
            * in the procedure ConvertToDefault rather than here

            SCAN FOR RTRIM( TblName ) == lcTableName
                lcFldName = RTRIM( &lcEnumFields..FldName )
                llBitType = IIF( RTRIM( &lcEnumFields..RmtType ) = "bit", .T., .F. )
                lcDefaultExpression = DBGETPROP( lcTableName + "." + lcFldName, "Field", "DefaultValue" )

                IF ( this.ExportDefaults AND !EMPTY( lcDefaultExpression )) OR llBitType THEN

                    *Convert Fox default to server default
                    do case
						case this.ExportDefaults and ;
                    		not empty( lcDefaultExpression )
							lcRemoteDefault = this.ConvertToDefault( lcDefaultExpression, ;
								lcFldName, lcTableName, lcRmtTableName, ;
								@lcRemoteDefaultName )
						case this.SQLServer and this.ServerVer >= 9
						otherwise
							lcRemoteDefault = "0"
					endcase

                    *If the default expression is 0 or.F., bind the zero default to field
					IF this.SQLServer and this.ServerVer < 9 and ;
						( alltrim( lcRemoteDefault ) == '0' or ;
						lcDefaultExpression = '0' or ;
						lcDefaultExpression = '.F.' )
                        llZD_field=.T.
                        lcRemoteDefault="0"
                        lcRemoteDefaultName=ZERO_DEFAULT_NAME
                        this.ZDUsed=.T.
                    ELSE
                        llZD_field=.F.
                    ENDIF

                    IF !this.ExportDefaults AND !llZD_field AND EMPTY( lcDefaultExpression ) THEN
                        LOOP
                    ENDIF

                    * Create the default if user is upsizing and has create default permissions
                    IF this.DoUpsize AND llTableExported AND this.Perm_Default THEN
                        IF llZD_field THEN
                            llDefaultCreated = this.ZeroDefault()
                        ELSE
                            * If we're Oracle or a non-zero default, just create the default
                            * ( for SQL Server ) or alter the table ( Oracle )
                            * Might have to drop existing default
                            IF this.ServerType = "Oracle" OR this.MaybeDrop( lcRemoteDefaultName, "default" ) THEN
                                llDefaultCreated = this.ExecuteTempSPT( lcRemoteDefault, @lnError, @lcDefaError )
                                IF !llDefaultCreated THEN
                                    REPLACE &lcEnumFields..DefaErrNo WITH lnError
                                    this.StoreError( lnError, lcDefaError, lcRemoteDefault, DEFA_ERR_LOC, lcTableName+"."+lcFldName, DEFAULT_LOC )
                                ENDIF
                            ELSE
                                lcDefaError = CANT_DROP_DEFA_LOC
                            ENDIF
                        ENDIF

                        *If we're upsizing to SQL Server, need to bind default if successfully created
						IF llDefaultCreated and this.SQLServer and ;
							this.ServerVer < 9
                            llDefaultBound = this.BindDefault( lcRemoteDefaultName, lcRmtTableName, lcFldName )
                        ELSE
                            llDefaultBound=.F.
                        ENDIF
                    ENDIF

                    REPLACE &lcEnumFields..DEFAULT WITH lcDefaultExpression, ;
                        &lcEnumFields..RmtDefault WITH lcRemoteDefault, ;
                        &lcEnumFields..RDName WITH lcRemoteDefaultName, ;
                        &lcEnumFields..DefaExport WITH llDefaultCreated, ;
                        &lcEnumFields..DefaBound WITH llDefaultBound, ;
                        &lcEnumFields..DefaError WITH lcDefaError

                    lcDefaultError=""
                    lnError=.F.
                ENDIF
            ENDSCAN
            SELECT ( lcEnumTables )
        ENDSCAN
		raiseevent( This, 'CompleteProcess' )



    FUNCTION MungeXbase
        PARAMETERS lcLocalExpression, lcObjectType, lcLocalTableName, lcRemoteTableName

        *Takes an Xbase expression and replaces as many mappable keywords as possible
        *This leaves tons of potential keywords that will not work on SQL Server or Oracle

        LOCAL lcServerSQL, lcRemoteExpression, lnOldArea, lcSetTalk, lcDelimiter, ;
            lnPos, lnPos1, lnPos2, lnPosMax

        *select expression mapping table
        lnOldArea = SELECT()
        lcSetTalk = SET( 'TALK' )
        SET TALK OFF

        IF !USED( "ExprMap" ) THEN
            SELECT 0
            USE ExprMap EXCLUSIVE
            IF this.ServerType = "Oracle" THEN
                SET FILTER TO !EMPTY( ExprMap.ORACLE )
            ELSE
                SET FILTER TO !EMPTY( ExprMap.SQLServer )
            ENDIF
        ELSE
            SELECT ExprMap
        ENDIF

        lcRemoteExpression = ''
        DO WHILE !EMPTY( lcLocalExpression )

            * find next language string ( i.e. smallest positive lnPos or 0 )
            lnPos  = AT( "'", lcLocalExpression )
            lnPos1 = AT( '"', lcLocalExpression )
            lnPos2 = AT( '[', lcLocalExpression )
            lnMax  = LEN( lcLocalExpression ) + 1
            lnPos = MIN( IIF( lnPos > 0, lnPos, lnMax ), IIF( lnPos1 > 0, lnPos1, lnMax ), IIF( lnPos2 > 0, lnPos2, lnMax ))
            lnPos = IIF( lnPos = lnMax, 0, lnPos )

            IF lnPos = 0
                lcLanguageString = lcLocalExpression
                lcLocalExpression = ''
            ELSE
                lcLanguageString = LEFT( lcLocalExpression, lnPos - 1 )
                lcDelimiter = SUBSTR( lcLocalExpression, lnPos, 1 )
                lcLocalExpression = SUBSTR( lcLocalExpression, lnPos + 1 )
            ENDIF

            * convert language string to server native syntax
            IF ( !EMPTY( lcLanguageString ))
                lcRemoteExpression = lcRemoteExpression + this.ConvertLanguageString( lcLanguageString, lcObjectType, ;
                    lcLocalTableName, lcRemoteTableName )
            ENDIF

            * find next constant string
            IF !EMPTY( lcLocalExpression )
                lcDelimiter = IIF( lcDelimiter == '[', ']', lcDelimiter )
                lnPos = AT( lcDelimiter, lcLocalExpression )
                IF lnPos = 0
                    lcConstantString = lcLocalExpression
                    lcLocalExpression = ''
                ELSE
                    lcConstantString = LEFT( lcLocalExpression, lnPos - 1 )
                    lcLocalExpression = SUBSTR( lcLocalExpression, lnPos + 1 )
                ENDIF

                * convert language string to server native syntax
                IF ( !EMPTY( lcConstantString ))
                    lcRemoteExpression = lcRemoteExpression + "'" + this.ConvertConstantString( lcConstantString ) + "'"
                ENDIF
            ENDIF

        ENDDO

        SELECT ( lnOldArea )
        SET TALK &lcSetTalk
        RETURN ALLTRIM( lcRemoteExpression )


    FUNCTION ConvertLanguageString
        PARAMETERS lcLocalExpression, lcObjectType, lcLocalTableName, lcRemoteTableName

        * Takes an Xbase expression and replaces mappable keywords, functions, names and date constants.
        * Fox function replacement is case-insensitive; Fox name replacement is case-sensitive.
        * This leaves some potential keywords that will not work on SQL Server or Oracle.

        LOCAL lcServerSQL, lnOldArea, lcServerType, lcXbase, lcEnumFields, ;
            aFieldNames, lcExpression, lnOccurence, lnPos, I

        * go through all the keywords for the appropriate server type
        lcLocalExpression = LOWER( lcLocalExpression )
        lcLocalTableName = LOWER( RTRIM( lcLocalTableName ))
        lcRemoteTableName = LOWER( RTRIM( lcRemoteTableName ))
        lcMapField = IIF( this.ServerType = ORACLE_SERVER, "Oracle", "SqlServer" )

        lcExpression = CHR( 1 )+lcLocalExpression
        SCAN
            IF !ExprMap.PAD
                lcServerSQL = RTRIM( ExprMap.&lcMapField. )
            ELSE
                lcServerSQL = " "+ RTRIM( ExprMap.&lcMapField. ) + " "
            ENDIF
            lcXbase = LOWER( RTRIM( ExprMap.FoxExpr ))
            lcExpression = STRTRAN( lcExpression, lcXbase, LOWER( lcServerSQL ))
        ENDSCAN
        lcExpression = SUBSTR( lcExpression, 2 )

        *create an array of local and remote field names ( where the two are different )
        lcEnumFields = RTRIM( this.EnumFieldsTbl )
        DIMENSION aFieldNames[1, 2]
        SELECT LOWER( FldName ), LOWER( RmtFldname ) FROM ( lcEnumFields ) WHERE ;
            RTRIM( &lcEnumFields..TblName )=lcLocalTableName AND ;
            &lcEnumFields..FldName<>&lcEnumFields..RmtFldname ;
            INTO ARRAY aFieldNames

        *replace field names with remotized names in table validation rules
        IF !EMPTY( aFieldNames ) THEN
            FOR I=1 TO ALEN( aFieldNames, 1 )
                lcExpression = STRTRAN( lcExpression, RTRIM( aFieldNames[i, 1] ), RTRIM( aFieldNames[i, 2] ))
            NEXT
        ENDIF

        * replace table name with remotized table name
        lcExpression = STRTRAN( lcExpression, lcLocalTableName, lcRemoteTableName )

        * Convert and format date constants
        lcExpression = this.ConvertDates( lcExpression )

        RETURN ALLTRIM( lcExpression )


    FUNCTION ConvertConstantString
        PARAMETERS lcString
        LOCAL lnPos, lnOccurrence

        * Takes a constant string expression that should be delimited by single quotes
        * in the remote expression. Replaces all internal single quotes ( ' ) with two single quotes ( '' )

        lnOccurence=1
        DO WHILE AT( "'", lcString, lnOccurence ) <> 0
            *Just add another chr( 39 ) in front of each one we find
            lnPos = AT( "'", lcString, lnOccurence )
            lcString = STUFF( lcString, lnPos, 1, CHR( 39 )+CHR( 39 ))
            lnOccurence = lnOccurence+2
        ENDDO

        RETURN lcString


    FUNCTION HandleQuotes
        PARAMETERS lcExpr, llNoContractions
        LOCAL lnPos, lnOccurrence

        *Chr( 39 ) is always a contraction in validation rule expressions, default
        *expressions, and validation messages.
        *Start with : "don't" ==> "'don''t'"
        *"create default foo_defa2 as 'don''t'"

        *If I know the string has no contractions, just replace doubles with singles

        IF PARAMETERS()=1
            lnOccurence=1
            DO WHILE ATCC( "'", lcExpr, lnOccurence )<>0
                *Just add another chr( 39 ) in front of each one we find
                lnPos=ATCC( "'", lcExpr, lnOccurence )
                lcExpr=STUFF( lcExpr, lnPos, 1, CHR( 39 )+CHR( 39 ))
                lnOccurence=lnOccurence+2
            ENDDO
        ENDIF

        lcExpr=STRTRAN( lcExpr, CHR( 34 ), CHR( 39 ))
        RETURN lcExpr


    FUNCTION ConvertDates
        PARAMETERS lcExpression
        LOCAL lcOccurence, lnPos1, lnPos2, ltDateTime, lcLocalDTString, lcRemoteDTString, ;
            lcCentury, lcDate, lnHour, lcSeconds, lcMark

        lnOccurence = 1
        DO WHILE .T.

            * find next date string
            lnPos1 = AT( "{", lcExpression, lnOccurence )
            lnPos2 = AT( "}", lcExpression, lnOccurence )
            IF ( lnPos1 = 0 OR lnPos2 = 0 OR lnPos1 > lnPos2 )
                EXIT
            ENDIF

            lcLocalDTString = SUBSTR( lcExpression, lnPos1, lnPos2 - lnPos1 + 1 )
            ltDateTime = CTOT( SUBSTR( lcExpression, lnPos1 + 1, lnPos2 - lnPos1 -1 ))

            lcCentury = SET( 'CENTURY' )
            lcDate = SET( 'DATE' )
            lnHour = SET( 'HOUR' )
            lcSeconds = SET( 'SECONDS' )
            lcMark = SET( 'MARK' )

            SET CENTURY ON
            SET DATE TO AMERICAN
            SET HOURS TO 12
            SET SECONDS ON
            SET MARK TO '/'

            lcRemoteDTString = DTOC( ltDateTime )

            SET CENTURY &lcCentury
            SET DATE TO &lcDate
            SET HOURS TO lnHour
            SET SECONDS &lcSeconds
            SET MARK TO lcMark

            * need exact format for Oracle
            IF ( this.ServerType == ORACLE_SERVER )
                lcRemoteDTString = "TO_DATE( '" + lcRemoteDTString + "', 'MM/DD/YYYY HH:MI:SS AM' )"
            ENDIF

            * need quotes for SqlServer
            IF ( this.SQLServer )
                lcRemoteDTString = "'" + lcRemoteDTString + "'"
            ENDIF

            lcExpression = STRTRAN( lcExpression, lcLocalDTString, lcRemoteDTString )
            lnOccurence	= lnOccurence + 1
        ENDDO

        RETURN lcExpression


    #IF SUPPORT_ORACLE
    FUNCTION ConvertToTrigger
        PARAMETERS lcRemoteExpression, lcRuleText, lcTableName, lcRmtTableName, lcRmtFldName
        LOCAL lcCRLF, lcEnumFieldsTbl

        lcCRLF=CHR( 10 )+CHR( 13 )

        *Try to make expression Oracle-ish
        lcRemoteExpression=this.MungeXbase( lcRemoteExpression, "foo", lcTableName, lcRmtTableName )

        *make sure the user doesn't have stuff of the form <tblname>.<fldname> in there
        lcRemoteExpression=STRTRAN( lcRemoteExpression, lcRmtTableName+"." )

        IF EMPTY( lcRuleText ) THEN
            IF !EMPTY( lcRmtFldName ) THEN
                *Oracle wants error messages to look like this:
                *'Value entered violates rule for field 'cust' in table 'customer'."
                *Note the use of single and double quotes is the exact opposite of SQL Server
                lcRuleText=STRTRAN( DEFLT_FLDMSG_LOC, "'|1'", gc2QT + lcRmtFldName + gc2QT )
                lcRuleText=STRTRAN( lcRuleText, "'|2'", gc2QT + lcRmtTableName + gc2QT )
            ELSE
                lcRuleText=STRTRAN( DEFLT_TBLMSG_LOC, "'|1'", gc2QT + lcRmtTableName + gc2QT )
            ENDIF
            lcRuleText="'"+lcRuleText+"'"
        ELSE

            *Replace existing single quotes with two single quotes ( they should then appear as one single quote mark )
            lcSingle=CHR( 39 )
            lcRuleText=STRTRAN( lcRuleText, "'", lcSingle+lcSingle )
            *Replace all double quote marks with single quote marks
            lcRuleText=STRTRAN( lcRuleText, gcQT, "'" )
        ENDIF

        *Build comment
        IF !EMPTY( lcRmtFldName ) THEN
            *We're dealing with a field validation rule
            *Need to replace <table>.<fldname> name with ":new.<fldname>"
            lcRemoteExpression=STRTRAN( lcRemoteExpression, lcRmtFldName, ":new." + lcRmtFldName )
            lcSQL=this.BuildComment( FTRIG_COMMENT_LOC, lcRmtFldName )
        ELSE
            *Table validation rule
            lcSQL=this.BuildComment( TTRIG_COMMENT_LOC, "blah blah blah" )

            *Get array of field names
            lcEnumFieldsTbl=this.EnumFieldsTbl
            SELECT RmtFldname FROM ( lcEnumFieldsTbl ) WHERE RTRIM( TblName )==lcTableName ;
                INTO ARRAY aFieldNames

            *Need to replace <fldname> name with ":new.<fldname>"
            FOR I=1 TO ALEN( aFieldNames, 1 )
                lcRemoteExpression=STRTRAN( lcRemoteExpression, RTRIM( aFieldNames[i] ), ":new." + RTRIM( aFieldNames[i] ))
            NEXT

        ENDIF

        *Complete SQL string
        lcSQL=lcSQL +	"IF NOT ( " + lcRemoteExpression + " )" + lcCRLF
        lcSQL=lcSQL + 	"     THEN raise_application_error( " + ERR_SVR_RULEVIO_ORA + ", " + lcRuleText + " );" + lcCRLF
        lcSQL=lcSQL + 	"END IF ; " + lcCRLF

        RETURN lcSQL

#ENDIF


#IF SUPPORT_ORACLE
    FUNCTION TestTrigger
        PARAMETERS lcTrigger, lcTable, lcFieldName, lnError, lcErrMsg

        *
        *Checks to see if a rule can be successfully converted to an Oracle trigger
        *

        *If user is just generating a script, just return .T.
        IF !this.DoUpsize THEN
            RETURN .T.
        ENDIF

        *Put the trigger together
        lcSQL="CREATE TRIGGER " + lcTable + TRIG_NAME
        lcSQL=lcSQL+" BEFORE INSERT OR UPDATE "
        IF !EMPTY( lcFieldName ) THEN
            lcSQL=lcSQL+" OF " + lcFieldName
        ENDIF
        lcSQL=lcSQL + " ON " + lcTable + " FOR EACH ROW BEGIN "
        lcSQL=lcSQL + lcTrigger + "END;"

        *Run it up the flag pole...

        IF !this.ExecuteTempSPT( lcSQL, @lnError, @lcErrMsg ) THEN
            RETURN .F.
        ELSE
            *Drop the trigger
            lcSQL="DROP TRIGGER " + lcTable + TRIG_NAME
            =this.ExecuteTempSPT( lcSQL )
            RETURN .T.
        ENDIF

#ENDIF


    FUNCTION ConvertToSproc
        PARAMETERS lcExpression, lcMessage, lcTableName, lcRmtTableName,  lcObjectType, lcSprocName, lcFldName

        *
        *Takes an Xbase rule and turns it into a stored procedure
        *

        *Do what you can to make the Xbase expression work on the server
        lcExpression=this.MungeXbase( lcExpression, lcObjectType, lcTableName, lcRmtTableName )

        LOCAL lcSQL, lcCRLF
        lcCR=CHR( 10 )
        lcCRLF=CHR( 10 )+CHR( 13 )

        *if there's no error message, build a default one
        IF EMPTY( lcMessage ) THEN
            IF lcObjectType="Field" THEN
                lcMessage=STRTRAN( DEFLT_FLDMSG_LOC, '|1', lcFldName )
                lcMessage=STRTRAN( lcMessage, '|2', lcTableName )
            ELSE
                lcMessage=STRTRAN( DEFLT_TBLMSG_LOC, '|1', lcTableName )
            ENDIF
        ENDIF

        *need to operate on error message to make sure quotes are right
        lcMessage=this.HandleQuotes( lcMessage )
        IF LEFT( lcMessage, 1 )<>CHR( 39 )
            lcMessage=CHR( 39 )+ lcMessage + CHR( 39 )
        ENDIF

        *Operate on the sproc name
        IF lcObjectType="Field" THEN
            *For field validation sprocs, the aim is to get something
            *like "vrf_customer_lname"
            lcSprocName=this.NameObject( lcRmtTableName, lcFldName, FLD_SPROC_PREFIX, MAX_NAME_LENGTH )
        ELSE
            lcSprocName=TBL_SPROC_PREFIX + LEFT( lcRmtTableName, MAX_NAME_LENGTH-LEN( FLD_SPROC_PREFIX ))
        ENDIF

        lcSQL=        "CREATE PROCEDURE " + lcSprocName + " @status char( 10 ) output AS" + lcCRLF
        lcSQL=lcSQL + "/*" + lcCR
        lcSQL=lcSQL + " * TABLE VALIDATION RULE FOR " + gcQT + lcRmtTableName + gcQT + lcCR
        lcSQL=lcSQL + " */"  + lcCRLF + lcCRLF
        lcSQL=lcSQL + "IF @status='Failed'"+lcCR
        lcSQL=lcSQL + "      RETURN" +lcCRLF
        lcSQL=lcSQL + "IF ( SELECT Count( * ) FROM " + lcRmtTableName + " WHERE NOT ( " + lcExpression + " )) > 0" + lcCR
        lcSQL=lcSQL + "      BEGIN" + lcCR
        lcSQL=lcSQL + "           RAISERROR " + ERR_SVR_RULEVIO_SQL +  " " + lcMessage +  lcCR
        lcSQL=lcSQL + "           SELECT @status='Failed'" + lcCR
        lcSQL=lcSQL + "      END" + lcCR
        lcSQL=lcSQL + "ELSE" + lcCR
        lcSQL=lcSQL + "      BEGIN" + lcCR
        lcSQL=lcSQL + "          SELECT @status='Succeeded'" + lcCR
        lcSQL=lcSQL + "      END" + lcCRLF

        RETURN lcSQL



    FUNCTION ConvertToDefault
        PARAMETERS lcDefaultExpression, lcFieldName, lcTableName, lcRemTableName, lcRemoteDefaultName
        LOCAL lcSQL

        * Defaults become ALTER TABLE statements for Oracle and Defaults on SQL Server
        * Try to make Xbase expression more server like

        lcDefaultExpression = this.MungeXbase( lcDefaultExpression, "Field", lcTableName, "" )

        * 1/8/01 jvf, Bug ID 155006, VFP7:Upsizing Wizard Does not Create Default Value
        * If the defaul tvalue = [""] ( without brackets ) then an error occurs
        * because 'lcDefaultExpression' will be empty...
        * ( eg, "CREATE DEFAULT <lcRemoteDefaultName> AS <lcDefaultExpression>" )
        * so checking for this and setting 'lcDefaultExpression' to "''" will solve the problem.
        IF EMPTY( lcDefaultExpression ) THEN
            lcDefaultExpression = "''"
        ENDIF

        IF this.ServerType = ORACLE_SERVER THEN
            lcSQL = "ALTER TABLE " + lcRemTableName + " MODIFY ( " + lcFieldName + " DEFAULT " + lcDefaultExpression + " )"
        ENDIF

        IF this.SQLServer THEN
            lcRemoteDefaultName = this.NameObject( lcTableName, lcFieldName, DEFAULT_PREFIX, MAX_NAME_LENGTH )
			IF this.ServerVer >= 9
				lcSQL = 'ALTER TABLE ' + lcRemTableName + ;
					' ADD CONSTRAINT [' + lcRemoteDefaultName + ;
					'] DEFAULT ' + lcDefaultExpression + ;
					' FOR ' + lcFieldName
			ELSE
            	lcSQL = "CREATE DEFAULT [" + lcRemoteDefaultName + ;
            		"] AS " + lcDefaultExpression
			ENDIF
        ENDIF

        RETURN lcSQL



    FUNCTION ConvertToConstraint
        PARAMETERS lcRuleExpression, lcTableName, lcRemTableName, lcConstName
        LOCAL lcSQL

        * Convert rules into table and field constraints for Oracle and SQL Server
        lcRuleExpression = this.MungeXbase( lcRuleExpression, "Field", lcTableName, "" )

        IF this.ServerType = ORACLE_SERVER
            lcSQL = "ALTER TABLE " + lcRemTableName + ;
                " ADD ( CONSTRAINT " + lcConstName + " CHECK( " + lcRuleExpression + " ))"
        ENDIF

        * For future use
        IF this.SQLServer THEN
        ENDIF

        RETURN lcSQL



    FUNCTION BindDefault
        PARAMETERS lcRemoteDefaultName, lcRmtTableName, lcFldName
        LOCAL lcSQL, llRetVal

        *bind a default to a field

        lcSQL="sp_bindefault " + lcRemoteDefaultName + ", " ;
            + "'" + lcRmtTableName + "." + lcFldName + "'"

        llRetVal=this.ExecuteTempSPT( lcSQL )

        RETURN llRetVal



    FUNCTION ZeroDefault
        LOCAL lcSQL, llRetVal, llDefaultExists, dummy, lcSQT

        IF this.ZeroDefaultCreated THEN
            RETURN .T.
        ELSE
            lcSQT=CHR( 39 )
            lcSQL="select uid from sysobjects where name =" + lcSQT + ZERO_DEFAULT_NAME + lcSQT
            llDefaultExists=this.SingleValueSPT( lcSQL, dummy, "uid" )
            *Only create it if it doesn't exist already
            IF !llDefaultExists THEN
                lcSQL= "CREATE DEFAULT " + ZERO_DEFAULT_NAME + " AS 0"
                llRetVal=this.ExecuteTempSPT( lcSQL )
                this.ZeroDefaultCreated=llRetVal
                RETURN llRetVal
            ELSE
                RETURN .T.
            ENDIF
        ENDIF



    FUNCTION BuildRiCode

        *
        * Generates an ALTER TABLE statement for Oracle, code for trigger for SQL Server
        * for all relations between tables that are being upsized
        *

        LOCAL lcEnumTables, lnOldArea, lcTableName, aFldNames, lcCRLF, ;
            aNewForeign, aNewPrimary, llRetVal, lcErr, llSkipChildTbl, ;
            lcEnumRelsTbl, lcEnum_Indexes, lcUpdateType, lcDeleteType, ;
            lcInsertType, lnTableCount, lcThermMsg, lcXPkey, ;
            lcParentLoc, lcChildLoc, lnError, lcErrMsg, lcOParen, lcCParen

* If we have an extension object and it has a BuildRICode method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'BuildRICode', 5 ) and ;
			not this.oExtension.BuildRICode( This )
			return
		ENDIF

        lcCRLF = CHR( 13 )
        lnOldArea = SELECT()
        lcEnumTables = this.EnumTablesTbl

        * Go grab all the relation information for tables in the source database
        this.GetRiInfo
        lcEnumRelsTbl = this.EnumRelsTbl
        lcEnum_Indexes = this.EnumIndexesTbl

        * We only want to deal with relations where both tables were successfully upsized or,
        * if we're generating a script, relations where both tables were selected to upsize
        SELECT ( lcEnumRelsTbl )
        SCAN
            IF this.TableUpsized( RTRIM( &lcEnumRelsTbl..DD_PARENT )) ;
                    AND this.TableUpsized( RTRIM( &lcEnumRelsTbl..DD_CHILD )) THEN
                REPLACE EXPORT WITH .T.
            ELSE
                REPLACE EXPORT WITH .F.
            ENDIF
        ENDSCAN

        SELECT COUNT( * ) FROM ( lcEnumRelsTbl ) WHERE EXPORT=.T. INTO ARRAY aTableCount
        lnTableCount=0
        this.InitTherm( BUILDING_RI_LOC, aTableCount, 0 )

        SCAN FOR EXPORT = .T.
            lcParent=RTRIM( &lcEnumRelsTbl..Dd_RmtPar )
            lcChild=RTRIM( &lcEnumRelsTbl..Dd_RmtChi )
            lcParentLoc=RTRIM( &lcEnumRelsTbl..DD_PARENT )
            lcChildLoc=RTRIM( &lcEnumRelsTbl..DD_CHILD )
            lcNewPrimary=RTRIM( &lcEnumRelsTbl..DD_PAREXPR )
            lcNewForeign=RTRIM( &lcEnumRelsTbl..DD_CHIEXPR )
            llClustIdxOK=.T.

            * Therm stuff
            lcThermMsg = STRTRAN( RI_THIS_LOC, '|1', lcParentLoc )
            lcThermMsg = STRTRAN( lcThermMsg, '|2', lcChildLoc )
            this.UpDateTherm( lnTableCount, lcThermMsg )
            lnTableCount = lnTableCount+1

            * Pick up what kind of relation type this is
            lcUpdateType = &lcEnumRelsTbl..dd_update
            lcDeleteType = &lcEnumRelsTbl..dd_delete
            lcInsertType = &lcEnumRelsTbl..dd_insert

            * For report
            IF lcUpdateType = "I" AND lcDeleteType = "I" AND lcInsertType = "I" ;
                    AND EMPTY( &lcEnumRelsTbl..ClustName ) THEN
                REPLACE &lcEnumRelsTbl..Exported WITH .F.
            ELSE
                REPLACE &lcEnumRelsTbl..Exported WITH .T.
            ENDIF

            * Turn fields in keys into an array ( used later all over the place )
            DIMENSION aNewForeign[1], aNewPrimary[1]
            aNewForeign[1]=""
            aNewPrimary[1]=""
            this.KeyArray( lcNewForeign, @aNewForeign )
            this.KeyArray( lcNewPrimary, @aNewPrimary )

            * do simple comparison of keys in case they won't match up
            IF ALEN( aNewForeign, 1 )<>ALEN( aNewPrimary, 1 ) THEN
                *mark the relation as unupsizable and move to the next relation
                this.StoreError( .NULL., "", "", KEYS_MISMATCH_LOC, lcParent+":"+lcChild, RI_LOC )
                REPLACE RIError WITH KEYS_MISMATCH_LOC, Exported WITH .F.
                LOOP
            ENDIF

            * Make sure keyfields are less than 17
            IF ALEN( aNewForeign, 1 )>MAX_INDEX_FIELDS THEN
                *mark the relation as unupsizable and move to the next relation
                REPLACE RIError WITH TOO_MANY_FIELDS_LOC, Exported WITH .F.
                this.StoreError( .NULL., "", "", TOO_MANY_FIELDS_LOC, lcParent+":"+lcChild, RI_LOC )
                LOOP
            ENDIF

            * Here's the RI plan:
            *SQL Server: use triggers for everything
            *Oracle: create two triggers that enforces RI via SQL *OR*
            *use DRI ( which won't cascade updates )
            *SQL '95: create triggers for everything *OR*
            *use DRI ( which won't cascade updates or deletes )
            *note: this code currently is not aware of SQL '95

            *
            * This block of code handles DRI for SQL '95 and Oracle
            * For SQL '95 it implements Update-restrict and Delete-restrict
            * For Oracle, it also implements Delete-cascades
            *

            IF this.ExportDRI AND ;
                    ( this.ServerType = "Oracle" OR this.ServerType = "SQL Server95" )
                *Implement RI constraints at table level since foreign key may be compound

                *Deal with parent table first ( or child constraints will fail )

                *See if the table already has a primary key that's correct for RI purposes
                *( i.e. the same as the one we'd create anyway )

                SELECT ( lcEnumTables )
                LOCATE FOR RTRIM( TblName )==lcParent

                *Here are the cases handled below:
                *has no primary key: add pkey, convert old non-pkey index to type pkey,
                *mark as already created so it doesn't get recreated
                *has primary key, it's right: create the pkey now and mark the index
                *as created
                *has primary key, it's wrong: log error

                lcXPkey=RTRIM( &lcEnumTables..PkeyExpr )
                lcTagName=RTRIM( &lcEnumTables..PKTagName )
                IF lcXPkey==lcNewPrimary OR EMPTY( lcXPkey )THEN

                    IF lcTagName=="" THEN
                        lcConstraintName=CHRTRAN( lcNewPrimary, ", []", "" )
                        lcConstraintName=PRIMARY_KEY_PREFIX + LEFT( lcConstraintName, MAX_NAME_LENGTH-LEN( PRIMARY_KEY_PREFIX ))
                    ELSE
                        lcConstraintName=lcTagName
                    ENDIF
                    lcConstraintName=this.UniqueOraName( lcConstraintName )

                    IF this.ServerType="Oracle" THEN
                        lcOParen="( "
                        lcCParen=" )"
                    ELSE
                        lcOParen=""
                        lcCParen=""
                    ENDIF

                    *Add primary key constraint
                    lcSQL="ALTER TABLE [" + lcParent + "]"
                    lcSQL=lcSQL + " ADD " + lcOParen + "CONSTRAINT "
                    lcSQL=lcSQL + "[" + lcConstraintName + "] PRIMARY KEY"
                    lcSQL=lcSQL + " ( " + lcNewPrimary + " )" + lcCParen

                    *Execute the statement if appropriate

                    IF this.DoUpsize THEN
                        llRetVal=this.ExecuteTempSPT( lcSQL, @lnError, @lcErrMsg )
                        IF !llRetVal THEN
                            this.StoreError( lnError, lcErrMsg, lcSQL, DRI_ERR_LOC, lcParent, RI_LOC )
                        ENDIF
                    ENDIF

                    SELECT ( lcEnum_Indexes )
                    LOCATE FOR UPPER( RTRIM( IndexName ))==UPPER( lcParentLoc ) ;
                        AND UPPER( RTRIM( RmtExpr ))==UPPER( lcNewPrimary )
                    IF EMPTY( lcXPkey ) THEN
                        *If the table had no true primary key, change the index acting
                        *as primary key to a real primary key ( even though we just created it )
                        *This is done just for the purposes of the report and the SQL script
                        REPLACE &lcEnum_Indexes..RmtType WITH "PRIMARY KEY", ;
                            &lcEnum_Indexes..Exported WITH llRetVal, ;
                            &lcEnum_Indexes..IndexSQL WITH lcSQL, ;
                            &lcEnum_Indexes..RmtName WITH lcConstraintName
                    ELSE
                        *Mark the index as already created
                        REPLACE &lcEnum_Indexes..Exported WITH llRetVal, ;
                            &lcEnum_Indexes..IndexSQL WITH lcSQL
                    ENDIF
                    SELECT ( lcEnumTables )
                    llSkipChildTbl=.F.

                ELSE

                    *log error that this rel couldn't be created because table has
                    *more than one primary key
                    lcErr=STRTRAN( MULTIPLE_PKEYS_LOC, "|1", lcParent )
                    SELECT ( lcEnumRelsTbl )
                    REPLACE RIError WITH lcErr ADDITIVE
                    llSkipChildTbl=.T.

                ENDIF

                *Deal with child table
                IF !llSkipChildTbl THEN

                    lcConstraintName=CHRTRAN( lcNewForeign, ", []", "" )
                    lcConstraintName=FOREIGN_KEY_PREFIX + LEFT( lcConstraintName, MAX_NAME_LENGTH-LEN( FOREIGN_KEY_PREFIX ))
                    lcConstraintName=this.UniqueOraName( lcConstraintName )

                    lcSQL="ALTER TABLE [" + lcChild + "] WITH NOCHECK" &&  NOCHECK Add JEI RKR 2005.03.29
                    lcSQL=lcSQL + " ADD " +lcOParen + "CONSTRAINT "
                    lcSQL=lcSQL + "[" + lcConstraintName + "]" +  " FOREIGN KEY "
                    lcSQL=lcSQL + " ( " + lcNewForeign + " )"
                    lcSQL=lcSQL + " REFERENCES [" + lcParent + "] ( " + lcNewPrimary +" )"

                    *SQL95 does not support cascading deletes
                    *Neither Oracle nor SQL95 support cascading updates via DRI
                    *{Change JEI - RKR 2005.03.28
*                    IF lcDeleteType=CASCADE_CHAR_LOC AND this.ServerType="Oracle" THEN
                    IF ( lcDeleteType=CASCADE_CHAR_LOC AND this.ServerType="Oracle" ) OR ;
                    	( (this.ServerType = "SQL Server95" AND this.ServerVer >= 8 ) AND ;
                    	( lcUpdateType = CASCADE_CHAR_LOC OR  lcDeleteType = CASCADE_CHAR_LOC ))THEN

                    	IF this.ServerType="Oracle"
                    		lcSQL=lcSQL + " ON DELETE CASCADE"
                    	ELSE
                    		* SQL Server
                    		IF lcDeleteType = CASCADE_CHAR_LOC
                    			lcSQL=lcSQL + " ON DELETE CASCADE"
                    		ENDIF
                    		IF lcUpdateType = CASCADE_CHAR_LOC
                    			lcSQL=lcSQL + " ON UPDATE CASCADE"
                    		ENDIF
                    	ENDIF
*                        lcSQL=lcSQL + " ON DELETE CASCADE )"
                    *}Change JEI - RKR 2005.03.28
                    ELSE
                        lcSQL=lcSQL + lcCParen
                    ENDIF

                    *Execute the statement if appropriate
                    IF this.DoUpsize AND llRetVal THEN
                        llRetVal=this.ExecuteTempSPT( lcSQL, @lnError, @lcErrMsg )
                        IF !llRetVal THEN
                            this.StoreError( lnError, lcErrMsg, lcSQL, DRI_ERR_LOC, lcChild, RI_LOC )
                        ENDIF
                    ENDIF

                    *Add comment and tack on to existing table definition sql
                    lcSQL = lcCRLF + lcCRLF + ORA_FKEY_COMMENT_LOC + lcCRLF + lcSQL
                    this.StoreRiCode( lcChild, "TableSQL", lcSQL, "FKeyCrea", llRetVal )

                    SELECT ( lcEnumRelsTbl )

                ENDIF

                LOOP	&& continue after DRI code

            ENDIF
            *End of DRI code

            * PARENT DELETE RI
            * Prevents deleting a PARENT record for which CHILD records exist,
            * or deletes dependent CHILD records ( cascading ).

            * PARENT DELETE for SQL 4.x or SQL '95 and cascade delete
            IF ( this.SQLServer AND lcDeleteType <> IGNORE_CHAR_LOC ) THEN

                lcRestr = this.BuildRestr( @aNewPrimary, "deleted", @aNewForeign, lcChild, "AND" )
                IF lcDeleteType = CASCADE_CHAR_LOC THEN
                    lcSQL = this.BuildComment( CASCADE_DELETES_LOC, lcChild )
                    lcSQL = lcSQL + "DELETE " + lcChild + " FROM deleted, " + lcChild +  " WHERE " + lcRestr + lcCRLF
                ELSE
                    lcErrMsg = this.HandleQuotes( gcQT + STRTRAN( DEPENDENT_ROWS_LOC, "|1", lcChild ) + gcQT )
                    lcSQL= this.BuildComment( PREVENT_DELETES_LOC, lcChild )
                    lcSQL= lcSQL + "IF ( SELECT COUNT( * ) FROM deleted, " + lcChild + " WHERE ( " + lcRestr + " )) > 0" + lcCRLF
                    lcSQL= lcSQL + "    BEGIN" + lcCRLF
                    lcSQL= lcSQL + "    RAISERROR " + ERR_SVR_DELREFVIO + " " + lcErrMsg + lcCRLF
                    lcSQL= lcSQL + "    SELECT @status='Failed'" + lcCRLF
                    lcSQL= lcSQL + "    END" + lcCRLF
                ENDIF

                *save this code
                this.StoreRiCode( lcParent, "DeleteRI", lcSQL )
            ENDIF

            * PARENT DELETE for Oracle
            IF this.ServerType = "Oracle" AND lcDeleteType <> IGNORE_CHAR_LOC

                lcRestr = this.BuildRestr( @aNewForeign, lcChild, @aNewPrimary, ":old", "AND" )
                IF lcDeleteType = CASCADE_CHAR_LOC THEN
                    lcSQL = this.BuildComment( CASCADE_DELETES_LOC, lcChild )
                    lcSQL = lcSQL + "IF DELETING THEN " + lcCRLF
                    lcSQL = lcSQL + "    DELETE FROM " + lcChild +  " WHERE " + lcRestr + ";" + lcCRLF
                    lcSQL = lcSQL + "END IF;" + lcCRLF
                ELSE
                    lcErrMsg = "'" + STRTRAN( DEPENDENT_ROWS_LOC, "'|1'", gc2QT + lcChild + gc2QT ) + "'"
                    lcSQL = this.BuildComment( PREVENT_DELETES_LOC, lcChild )
                    lcSQL = lcSQL + "IF DELETING THEN " + lcCRLF
                    lcSQL = lcSQL + "    SELECT COUNT( * ) INTO " + REC_COUNT_VAR + " FROM " + lcChild + " WHERE ( " + lcRestr + " );" + lcCRLF
                    lcSQL = lcSQL + "    IF " + REC_COUNT_VAR + " > 0 THEN " + lcCRLF
                    lcSQL = lcSQL + "        raise_application_error( " + ERR_SVR_DELREFVIO_ORA + ", " + lcErrMsg + " );" + lcCRLF
                    lcSQL = lcSQL + "    END IF;" + lcCRLF
                    lcSQL = lcSQL + "END IF;" + lcCRLF
                ENDIF

                * save this code
                this.StoreRiCode( lcParent, "DeleteRI", lcSQL )

            ENDIF

            * PARENT UPDATE trigger
            * Prevents changing a PARENT key for which CHILD records exist,
            * or keeps CHILD keys in sync with PARENT keys ( cascading ).
            * Executed only if SQL Server or if Oracle or SQL '95 when updates are cascaded
            * Handle SQL Server ( 4.x or '95 ) case here

            * PARENT UPDATE for Sql Server
            IF this.SQLServer AND lcUpdateType <> IGNORE_CHAR_LOC

                lcRestr = this.BuildRestr( @aNewPrimary, "deleted", @aNewForeign, lcChild, "AND" )
                IF lcUpdateType = CASCADE_CHAR_LOC THEN
                    lcSQL = this.BuildComment( CASCADE_UPDATES_LOC, lcChild )
                ELSE
                    lcSQL = this.BuildComment( PREVENT_M_UPDATES_LOC, lcChild )
                ENDIF
                lcSQL = lcSQL + "IF " + this.BuildUpdateTest( @aNewPrimary )
                lcSQL = lcSQL + " AND @status<>'Failed'" + lcCRLF
                lcSQL = lcSQL + "    BEGIN" + lcCRLF
                IF lcUpdateType=CASCADE_CHAR_LOC THEN
                    lcSetKeys = this.BuildRestr( @aNewForeign, lcChild, @aNewPrimary, "inserted", ", " )
                    lcSQL = lcSQL + "         UPDATE " + lcChild + lcCRLF
                    lcSQL = lcSQL + "         SET " + lcSetKeys + lcCRLF
                    lcSQL = lcSQL + "         FROM " + lcChild + ", deleted, inserted" + lcCRLF
                    lcSQL = lcSQL + "         WHERE " + lcRestr + lcCRLF
                ELSE
                    lcErrMsg=this.HandleQuotes( gcQT+STRTRAN( DEPENDENT_ROWS_LOC, "|1", lcChild )+ gcQT )
                    lcSQL = lcSQL + "    IF ( SELECT COUNT( * ) FROM deleted, " + lcChild + " WHERE ( " + lcRestr + " )) > 0" + lcCRLF
                    lcSQL = lcSQL + "        BEGIN" + lcCRLF
                    lcSQL = lcSQL + "            RAISERROR " + ERR_SVR_DELREFVIO + " " + lcErrMsg + lcCRLF
                    lcSQL = lcSQL + "            SELECT @status='Failed'" + lcCRLF
                    lcSQL = lcSQL + "            END" + lcCRLF
                ENDIF
                lcSQL = lcSQL + "    END" + lcCRLF

                *save this code
                this.StoreRiCode( lcParent, "UpdateRI", lcSQL )
            ENDIF

            * PARENT UPDATE for Oracle
            IF this.ServerType = "Oracle" AND lcUpdateType <> IGNORE_CHAR_LOC

                lcRestr = this.BuildRestr( @aNewForeign, lcChild, @aNewPrimary, ":old", "AND" )
                lcUpdTest = this.BuildRestr( @aNewForeign, ":old", @aNewPrimary, ":new", "OR" )
                lcUpdTest = STRTRAN( lcUpdTest, "=", "!=" )
                IF lcUpdateType = CASCADE_CHAR_LOC THEN
                    lcSQL = this.BuildComment( CASCADE_UPDATES_LOC, lcChild )
                ELSE
                    lcSQL = this.BuildComment( PREVENT_M_UPDATES_LOC, lcChild )
                ENDIF
                IF lcUpdateType = CASCADE_CHAR_LOC THEN
                    lcSetKeys = this.BuildRestr( @aNewForeign, lcChild, @aNewPrimary, ":new", ", " )
                    lcSQL = lcSQL + "IF UPDATING AND " + lcUpdTest + " THEN" + lcCRLF
                    lcSQL = lcSQL + "    UPDATE " + lcChild + lcCRLF
                    lcSQL = lcSQL + "    SET " + lcSetKeys + lcCRLF
                    lcSQL = lcSQL + "    WHERE " + lcRestr + ";" + lcCRLF
                    lcSQL = lcSQL + "END IF;" + lcCRLF
                ELSE
                    lcErrMsg = "'" + STRTRAN( DEPENDENT_ROWS_LOC, "'|1'", gc2QT + lcChild + gc2QT ) +  "'"
                    lcSQL = lcSQL + "IF UPDATING AND " + lcUpdTest + " THEN" + lcCRLF
                    lcSQL = lcSQL + "    SELECT COUNT( * ) INTO " + REC_COUNT_VAR + " FROM " + lcChild + " WHERE ( " + lcRestr + " );" + lcCRLF
                    lcSQL = lcSQL + "    IF " + REC_COUNT_VAR + " > 0 THEN" + lcCRLF
                    lcSQL = lcSQL + "        raise_application_error( " + ERR_SVR_UPDREFVIO_ORA + ", " + lcErrMsg + " );" + lcCRLF
                    lcSQL = lcSQL + "    END IF;" + lcCRLF
                    lcSQL = lcSQL + "END IF;" + lcCRLF
                ENDIF

                this.StoreRiCode( lcParent, "UpdateRI", lcSQL )

            ENDIF

            * CHILD UPDATE trigger
            * Prevents changing or adding a CHILD record to a key not in the PARENT table
            * CHILD INSERT trigger
            * Prevents adding a CHILD record for which no PARENT record exists

            * CHILD UPDATE AND INSERT for SQL Server 4.x and 95
            IF this.SQLServer THEN

                * CHILD UPDATE trigger
                IF lcUpdateType = RESTRICT_CHAR_LOC THEN
                    lcRestr = this.BuildRestr( @aNewPrimary, lcParent, @aNewForeign, "inserted", "AND" )
                    lcErrMsg = this.HandleQuotes( gcQT + STRTRAN( CANT_ORPHAN_LOC, "|1", lcParent ) + gcQT )

                    lcSQL = this.BuildComment( PREVENT_C_UPDATES_LOC, lcParent )
                    lcSQL = lcSQL + "IF " + this.BuildUpdateTest( @aNewForeign )
                    lcSQL =lcSQL + " AND @status<>'Failed'" + lcCRLF
                    lcSQL = lcSQL + "    BEGIN" + lcCRLF
                    lcSQL = lcSQL + "IF ( SELECT COUNT( * ) FROM inserted ) !=" + lcCRLF
                    lcSQL = lcSQL + "           ( SELECT COUNT( * ) FROM " + lcParent + ", inserted WHERE ( " + lcRestr + " ))" + lcCRLF
                    lcSQL = lcSQL + "            BEGIN" + lcCRLF
                    lcSQL = lcSQL + "                RAISERROR " + ERR_SVR_UPDREFVIO + " " + lcErrMsg + lcCRLF
                    lcSQL = lcSQL + "                SELECT @status = 'Failed'" + lcCRLF
                    lcSQL = lcSQL + "            END" + lcCRLF
                    lcSQL = lcSQL + "    END" + lcCRLF

                    *save this code
                    this.StoreRiCode( lcChild, "UpdateRI", lcSQL )

                ENDIF

                * CHILD INSERT trigger
                IF lcInsertType = RESTRICT_CHAR_LOC
                    lcErrMsg = this.HandleQuotes( gcQT+ STRTRAN( CANT_ORPHAN_LOC, "|1", lcParent ) + gcQT )
                    lcRestr = this.BuildRestr( @aNewPrimary, lcParent, @aNewForeign, "inserted", "AND" )
                    lcSQL = this.BuildComment( PREVENT_INSERTS_LOC, lcParent )
                    lcSQL = lcSQL + "IF @status<>'Failed'" + lcCRLF
                    lcSQL = lcSQL + "    BEGIN" + lcCRLF
                    lcSQL = lcSQL + "    IF( SELECT COUNT( * ) FROM inserted ) !=" + lcCRLF
                    lcSQL = lcSQL + "   ( SELECT COUNT( * ) FROM " + lcParent + ", inserted WHERE ( " + lcRestr + " ))" + lcCRLF
                    lcSQL = lcSQL + "        BEGIN" + lcCRLF
                    lcSQL = lcSQL + "            RAISERROR " + ERR_SVR_UPDREFVIO + " " + lcErrMsg + lcCRLF
                    lcSQL = lcSQL + "            SELECT @status='Failed'" + lcCRLF
                    lcSQL = lcSQL + "        END" + lcCRLF
                    lcSQL = lcSQL + "    END" + lcCRLF

                    *save this code
                    this.StoreRiCode( lcChild, "InsertRI", lcSQL )

                ENDIF
            ENDIF

            * CHILD UPDATE AND INSERT for Oracle
            IF this.ServerType = "Oracle" AND ;
                    ( lcUpdateType = RESTRICT_CHAR_LOC OR lcInsertType = RESTRICT_CHAR_LOC )

                lcUpdTest = this.BuildRestr( @aNewForeign, ":old", @aNewPrimary, ":new", "OR" )
                lcUpdTest = STRTRAN( lcUpdTest, "=", "!=" )

                lcRestr = this.BuildRestr( @aNewPrimary, lcParent, @aNewForeign, ":new", "AND" )
                lcErrMsg = "'" + STRTRAN( CANT_ORPHAN_LOC, "'|1'", gc2QT + lcParent + gc2QT ) + "'"
                lcSQL = this.BuildComment( PREVENT_SELF_O_LOC, lcParent )
                lcSQL = lcSQL + "IF ( UPDATING AND " + lcUpdTest + " ) OR INSERTING THEN" + lcCRLF
                lcSQL = lcSQL + "    SELECT COUNT( * ) INTO " + REC_COUNT_VAR + " FROM " + lcParent + " WHERE ( " + lcRestr + " );" + lcCRLF
                lcSQL = lcSQL + "    IF " + REC_COUNT_VAR + " = 0 THEN" + lcCRLF
                lcSQL = lcSQL + "        raise_application_error( " + ERR_SVR_UPDREFVIO_ORA + ", " + lcErrMsg + " );" + lcCRLF
                lcSQL = lcSQL + "    END IF;" + lcCRLF
                lcSQL = lcSQL + "END IF;" + lcCRLF

                *save this code
                this.StoreRiCode( lcChild, "UpdateRI", lcSQL )
            ENDIF

            *If we're dealing with SQL Server, run sp_primarykey, sp_foreignkey
            IF this.ServerType <> "Oracle" AND ALEN( aNewPrimary, 1 ) <= 8 AND !this.ExportDRI THEN
                *Check if the table is in multiple rels
                SELECT COUNT( * ) FROM ( lcEnumRelsTbl ) WHERE RTRIM( DD_PARENT )==lcParent ;
                    AND !DD_CHIEXPR=="" AND !DD_PAREXPR=="" INTO ARRAY aDupeCount
                IF aDupeCount>1 THEN
                    *check to see if the local table has a primary key index
                    SELECT RmtExpr FROM ( lcEnum_Indexes ) WHERE RTRIM( IndexName )==lcParent ;
                        AND LclIdxType="Primary key" INTO ARRAY aIndexExpr
                    *If primary key index expression is same as RI primary key, run sp_primary key
                    IF RTRIM( aIndexExpr )==lcNewPrimary THEN
                        this.SetPKey( lcParent, lcNewPrimary )
                        this.SetFKey( lcChild, lcNewForeign, lcParent )
                    ELSE
                        this.SetCommonKey( lcParent, @aNewPrimary, lcChild, @aNewForeign )
                    ENDIF
                ELSE
                    this.SetPKey( lcParent, lcNewPrimary )
                    this.SetFKey( lcChild, lcNewForeign, lcParent )
                ENDIF
            ENDIF

        ENDSCAN
		raiseevent( This, 'CompleteProcess' )
        SELECT ( lnOldArea )



    FUNCTION SetCommonKey
        PARAMETERS lcTable1, aNewPrimary, lcTable2, aNewForeign
        LOCAL lcSQL
        lcSQL="sp_commonkey " + lcTable1 + ", " + lcTable2
        FOR I=1 TO ALEN( aNewPrimary, 1 )
            lcSQL=lcSQL + ", " + aNewPrimary[i] + ", " + aNewForeign[i]
        NEXT
        =this.ExecuteTempSPT( lcSQL )



    FUNCTION SetPKey
        PARAMETERS lcTable, lcKey
        LOCAL lcSQL
        lcSQL="sp_primarykey " + lcTable + ", " + lcKey
        =this.ExecuteTempSPT( lcSQL )



    FUNCTION SetFKey
        PARAMETERS lcChild, lcChildKey, lcParent
        LOCAL lcSQL
        lcSQL="sp_foreignkey " + lcChild +", " + lcParent + ", " + lcChildKey
        =this.ExecuteTempSPT( lcSQL )



    FUNCTION BuildComment
        PARAMETERS lcMainComment, lcToInsert
        LOCAL lcCRLF
        #DEFINE SEARCH_TOKEN "|1"

        lcCRLF = CHR( 13 )
        lcMainComment = STRTRAN( lcMainComment, SEARCH_TOKEN, lcToInsert )
        lcComment= lcCRLF + "/* " + lcMainComment + " */" + lcCRLF
        RETURN lcComment



    FUNCTION KeyArray
        PARAMETERS lcKeyString, aKeyArray

        *Takes comma separated list of fields in a key and converts it to an array

        lcKeyString=lcKeyString+ ", "
        DO WHILE !lcKeyString==""
            this.InsaItem( @aKeyArray, ALLTRIM( LEFT( lcKeyString, AT( ", ", lcKeyString )-1 )) )
            lcKeyString=SUBSTR( lcKeyString, AT( ", ", lcKeyString )+1 )
        ENDDO



    FUNCTION BuildRestr
        PARAMETERS aNewPrimary, lcParent, aNewForeign, lcChild, lcConjunction
        LOCAL lcBuildRestr

        *Creates a restriction string, e.g. "customer.cust.id=order.cust_id"

        lcBuildRestr=""
        FOR I=1 TO ALEN( aNewPrimary, 1 )
            IF I>1 THEN
                lcBuildRestr=lcBuildRestr + " " + lcConjunction + " "
            ENDIF
            lcBuildRestr=lcBuildRestr + lcParent + "." + aNewPrimary[i] + " = " + 	;
                lcChild + "." + aNewForeign[i]
        NEXT
        RETURN lcBuildRestr



    FUNCTION BuildUpdateTest
        PARAMETERS aKeyFields
        LOCAL lcSQL, I
        lcSQL=""

        FOR I = 1 TO ALEN( aKeyFields, 1 )
            IF I > 1 THEN
                lcSQL= lcSQL + " OR "
            ENDIF
            lcSQL= lcSQL + "UPDATE( " + aKeyFields[i] + " )"
        NEXT
        RETURN lcSQL



    FUNCTION StoreRiCode
        PARAMETERS lcRmtTblName, lcFldName1, lcSQL, lcFldName2, llRetVal
        LOCAL lnOldArea

        lnOldArea=SELECT()
        lcEnumTablesTbl=this.EnumTablesTbl
        SELECT ( lcEnumTablesTbl )
        LOCATE FOR RTRIM( RmtTblName )==RTRIM( lcRmtTblName )

        REPLACE &lcEnumTablesTbl..&lcFldName1 WITH lcSQL ADDITIVE

        IF PARAM()=4 THEN
            REPLACE &lcEnumTablesTbl..&lcFldName2 WITH llRetVal
        ENDIF

        SELECT ( lnOldArea )



    FUNCTION CreateTriggers
        LOCAL lcTableName, lcCRLF, lcDeleteRI, lcInsertRI, lcUpdateRI, lcSproc, ;
            lcEnumFields, lnOldArea, lnTableCount, lcTrigName, lnError, lcErrMsg, ;
            lcUpdateType, lcDeleteType, lcInsertType
        lnOldArea = SELECT()
        llRetVal = .F.
        lcCRLF = CHR( 13 )

        lcEnumTables = RTRIM( this.EnumTablesTbl )
        lcEnumRelsTbl = RTRIM( this.EnumRelsTbl )

        SELECT ( lcEnumTables )
        lnTableCount = 0

        IF !this.ExportRelations AND !this.ExportDefaults AND !this.ExportValidation THEN
            RETURN
        ENDIF

* If we have an extension object and it has a CreateTriggers method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'CreateTriggers', 5 ) and ;
			not this.oExtension.CreateTriggers( This )
			return
		ENDIF

        *Thermometer stuff
        SELECT COUNT( * ) FROM ( lcEnumTables ) WHERE &lcEnumTables..EXPORT=.T. ;
            INTO ARRAY aTableCount
        this.InitTherm( CREA_TRIGGERS_LOC, aTableCount, 0 )
        lnTableCount=0

        IF this.ServerType = "Oracle"
            #IF SUPPORT_ORACLE
                * Oracle: Create up to two row triggers
                * before update and insert: insert ri, restrict update ri, restrict delete ri ( AT )
                * after update or delete: cascade update ri, cascade delete ri ( AT )

                SCAN FOR EXPORT = .T. AND ( !EMPTY( InsertRI ) OR !EMPTY( UpdateRI ) OR !EMPTY( DeleteRI ))
                    lcTableName = RTRIM( &lcEnumTables..RmtTblName )

                    * Thermometer stuff
                    lcThermMsg = STRTRAN( THIS_TABLE_LOC, '|1', lcTableName )
                    this.UpDateTherm( lnTableCount, lcThermMsg )
                    lnTableCount = lnTableCount+1

                    * Grab RI and validation rule code ( Note that all the validation rule
                    * SQL has already been placed in the InsertRI field )
                    lcInsertRI = &lcEnumTables..InsertRI
                    lcUpdateRI = &lcEnumTables..UpdateRI
                    lcDeleteRI = &lcEnumTables..DeleteRI

                    *if update code is restrict, need to
                    *toss in a variable declaration at the top of the trigger
                    *( it can't come inside of the BEGIN...END commands )
                    IF AT( REC_COUNT_VAR, lcDeleteRI ) <> 0 OR AT( REC_COUNT_VAR, lcUpdateRI ) <> 0 THEN
                        lcDecl= "DECLARE " + REC_COUNT_VAR + " NUMBER;" + lcCRLF
                    ELSE
                        lcDecl=""
                    ENDIF

                    *Assemble before trigger
                    IF !EMPTY( lcInsertRI ) OR !EMPTY( lcUpdateRI ) OR !EMPTY( lcDeleteRI )
                        lcTrigName = ORA_BIUD_TRIG_PREFIX + LEFT( lcTableName, MAX_NAME_LENGTH-LEN( ORA_BIUD_TRIG_PREFIX ))
                        lcTrigName = this.UniqueOraName( lcTrigName )

                        lcSQL = "CREATE TRIGGER " + lcTrigName + lcCRLF
                        lcSQL = lcSQL + "BEFORE INSERT OR UPDATE OR DELETE"
                        lcSQL = lcSQL + " ON " + lcTableName + " FOR EACH ROW " + lcCRLF
                        lcSQL = lcSQL + lcDecl + "BEGIN " + lcCRLF
                        lcSQL = lcSQL + lcInsertRI + lcUpdateRI + lcDeleteRI + lcCRLF + "END;"

                        IF this.DoUpsize AND this.Perm_Trigger THEN
                            llRetVal = this.ExecuteTempSPT( lcSQL, @lnError, @lcErrMsg )
                            IF !llRetVal THEN
                                this.StoreError( lnError, lcErrMsg, lcSQL, TRIG_ERR_LOC, lcTableName, TRIGGER_LOC )
                                REPLACE &lcEnumTables..RIError WITH lcErrMsg ADDITIVE, ;
                                    RIErrNo WITH lnError
                            ENDIF

                        ENDIF
                        REPLACE &lcEnumTables..InsertRI WITH lcSQL, ;
                            &lcEnumTables..ItrigName WITH lcTrigName, ;
                            &lcEnumTables..InsertX WITH llRetVal

                        * Trigger sql is appended in the script, so clean it up
                        REPLACE &lcEnumTables..UpdateRI WITH "", ;
                            &lcEnumTables..DeleteRI WITH ""
                    ENDIF

                    *Assemble after trigger
                    IF 	.F.
                        lcTrigName = ORA_AUD_TRIG_PREFIX + LEFT( lcTableName, MAX_NAME_LENGTH-LEN( ORA_AUD_TRIG_PREFIX ))
                        lcTrigName = this.UniqueOraName( lcTrigName )
                        lcSQL =         "CREATE TRIGGER " + lcTrigName + lcCRLF
                        lcSQL = lcSQL + "AFTER UPDATE OR DELETE"
                        lcSQL = lcSQL + " ON " + lcTableName + " FOR EACH ROW " + lcCRLF
                        lcSQL  =lcSQL + lcDecl + "BEGIN " + lcCRLF
                        lcSQL = lcSQL + lcUpdateRI + lcDeleteRI + lcCRLF + "END;"

                        IF this.DoUpsize AND this.Perm_Trigger THEN
                            llRetVal = this.ExecuteTempSPT( lcSQL, @lnError, @lcErrMsg )
                            IF !llRetVal THEN
                                this.StoreError( lnError, lcErrMsg, lcSQL, TRIG_ERR_LOC, lcTableName, TRIGGER_LOC )
                                REPLACE &lcEnumTables..RIError WITH lcErrMsg ADDITIVE, ;
                                    RIErrNo WITH lnError
                            ENDIF
                        ENDIF
                        REPLACE &lcEnumTables..DeleteRI WITH lcSQL, ;
                            &lcEnumTables..DtrigName WITH lcTrigName, ;
                            &lcEnumTables..DeleteX WITH llRetVal, ;
                            &lcEnumTables..UpdateRI WITH ""
                    ENDIF

                ENDSCAN

            #ENDIF

        ELSE

            *SQL Server: Create up to three triggers
            *update trigger: updateRI, rules/sprocs
            *insert trigger: insertRI, rules/sprocs
            *delete trigger: deleteRI

            SCAN FOR &lcEnumTables..EXPORT=.T.
                lcSproc=""
                lcSQL=""
                lcTableName=RTRIM( &lcEnumTables..RmtTblName )

                lcThermMsg=STRTRAN( THIS_TABLE_LOC, '|1', lcTableName )
                this.UpDateTherm( lnTableCount, lcThermMsg )
                lnTableCount=lnTableCount+1

                *Build sproc string which will be used in Insert and Update triggers

                *Grab table validation rules ( i.e. sprocs ) that were successfully created
                *Grab them regardless if the user is just generating a script
                IF &lcEnumTables..RuleExport =.T. ;
                        OR ( !this.DoUpsize AND !EMPTY( &lcEnumTables..RmtRule )) THEN
                    lcSproc=TBLRULE_COMMENT_LOC
                    lcSproc=lcSproc+ "execute "+RTRIM( &lcEnumTables..RRuleName )  ;
                        + " @status output" + lcCRLF
                ENDIF

                *Grab RI code
                lcInsertRI=&lcEnumTables..InsertRI
                lcUpdateRI=&lcEnumTables..UpdateRI
                lcDeleteRI=&lcEnumTables..DeleteRI

                *Grab field validation sprocs
                lcEnumFields=RTRIM( this.EnumFieldsTbl )
                SELECT ( lcEnumFields )
                SCAN FOR RTRIM( &lcEnumFields..TblName )==lcTableName
                    lcFieldName=RTRIM( &lcEnumFields..RmtFldname )

                    *Only add it to the string if the sproc was successfully created

                    IF &lcEnumFields..RuleExport =.T. ;
                            OR ( !this.DoUpsize AND !EMPTY( &lcEnumFields..RmtRule )) THEN
                        lcSproc=lcSproc + STRTRAN( FLDRULE_COMMENT_LOC, "|1", lcFieldName )
                        lcSproc=lcSproc + "execute " + RTRIM( &lcEnumFields..RRuleName ) + " @status output" + lcCRLF
                    ENDIF

                ENDSCAN

                SELECT ( lcEnumTables )

                *Strings used in all the triggers:
                lcStatus="DECLARE @status char( 10 )  " + STATUS_COMMENT_LOC + lcCRLF
                lcStatus=lcStatus + "SELECT @status='Succeeded'" + lcCRLF
                lcRollBack=ROLLBACK_LOC + lcCRLF + "IF @status='Failed'" + lcCRLF +;
                    "ROLLBACK TRANSACTION" + lcCRLF

                IF !EMPTY( lcInsertRI ) OR !EMPTY( lcSproc ) THEN

                    *Create insert trigger
                    lcTrigName=ITRIG_PREFIX + LEFT( lcTableName, MAX_NAME_LENGTH-LEN( ITRIG_PREFIX ))
                    lcSQL="CREATE TRIGGER " + lcTrigName
                    lcSQL=lcSQL + " ON " + lcTableName + " FOR INSERT AS " + lcCRLF
                    lcSQL=lcSQL + lcStatus + lcSproc + lcInsertRI
                    lcSQL=lcSQL + lcRollBack
                    IF this.DoUpsize THEN
                        llRetVal=this.ExecuteTempSPT( lcSQL, @lnError, @lcErrMsg )
                        IF !llRetVal THEN
                            this.StoreError( lnError, lcErrMsg, lcSQL, TRIG_ERR_LOC, lcTableName, TRIGGER_LOC )
                            REPLACE &lcEnumTables..RIError WITH lcErrMsg ADDITIVE, ;
                                RIErrNo WITH lnError
                        ENDIF
                    ENDIF
                    REPLACE &lcEnumTables..InsertRI WITH lcSQL, ;
                        &lcEnumTables..ItrigName WITH lcTrigName, ;
                        &lcEnumTables..InsertX WITH llRetVal
                ENDIF

                IF !EMPTY( lcUpdateRI ) OR !EMPTY( lcSproc ) THEN

                    *Create update trigger
                    lcTrigName=UTRIG_PREFIX + LEFT( lcTableName, MAX_NAME_LENGTH-LEN( UTRIG_PREFIX ))
                    lcSQL="CREATE TRIGGER " + lcTrigName
                    lcSQL=lcSQL + " ON " + lcTableName + " FOR UPDATE AS " + lcCRLF
                    lcSQL=lcSQL + lcStatus + lcSproc + lcUpdateRI
                    lcSQL=lcSQL + lcRollBack
                    IF this.DoUpsize THEN
                        llRetVal=this.ExecuteTempSPT( lcSQL, @lnError, @lcErrMsg )
                        IF !llRetVal THEN
                            this.StoreError( lnError, lcErrMsg, lcSQL, TRIG_ERR_LOC, lcTableName, TRIGGER_LOC )
                            REPLACE &lcEnumTables..RIError WITH lcErrMsg ADDITIVE, ;
                                RIErrNo WITH lnError
                        ENDIF
                    ENDIF
                    REPLACE &lcEnumTables..UpdateRI WITH lcSQL, ;
                        &lcEnumTables..UtrigName WITH lcTrigName, ;
                        &lcEnumTables..UpdateX WITH llRetVal
                ENDIF

                IF !EMPTY( lcDeleteRI ) THEN

                    *Create delete trigger
                    lcTrigName=DTRIG_PREFIX + LEFT( lcTableName, MAX_NAME_LENGTH-LEN( DTRIG_PREFIX ))
                    lcSQL="CREATE TRIGGER " + lcTrigName
                    lcSQL=lcSQL + " ON " + lcTableName + " FOR DELETE AS " + lcCRLF
                    lcSQL=lcSQL + lcStatus + lcDeleteRI
                    lcSQL=lcSQL + lcRollBack
                    IF this.DoUpsize THEN
                        llRetVal=this.ExecuteTempSPT( lcSQL, @lnError, @lcErrMsg )
                        IF !llRetVal THEN
                            this.StoreError( lnError, lcErrMsg, lcSQL, TRIG_ERR_LOC, lcTableName, TRIGGER_LOC )
                            REPLACE &lcEnumTables..RIError WITH lcErrMsg ADDITIVE, ;
                                RIErrNo WITH lnError
                        ENDIF
                    ENDIF
                    REPLACE &lcEnumTables..DeleteRI WITH lcSQL, ;
                        &lcEnumTables..DtrigName WITH lcTrigName, ;
                        &lcEnumTables..DeleteX WITH llRetVal
                ENDIF

            ENDSCAN

        ENDIF
		raiseevent( This, 'CompleteProcess' )
        SELECT ( lnOldArea )



    FUNCTION GetRiInfo
        LOCAL lcEnumRelsTbl, lnOldArea, p_dbc, l_rmtchitable, l_rmtpartable, lcTableName
        PRIVATE aDupeCount

        *Thanks, George Goley, for the heart of this code
        IF !this.GetRiInfoRecalc THEN
            RETURN
        ENDIF

        #DEFINE KEY_CHILDTAG		13	&& For RELATION objects: name of child ( from ) index tag
        #DEFINE KEY_RELTABLE		18	&& For RELATION objects: name of related table
        #DEFINE KEY_RELTAG			19	&& For RELATION objects: name of related index tag
        #DEFINE d_updatespot		1
        #DEFINE d_deletespot		2
        #DEFINE d_insertspot		3

        *Make sure that all tables selected for upsizing have had their
        *indexes analyzed

        lnOldArea=SELECT()
        *Make sure index info is up-to-date
        IF this.AnalyzeIndexesRecalc THEN
            this.AnalyzeIndexes
        ENDIF

        p_dbc=RTRIM( this.SourceDB )

        lcEnumIndexesTbl=this.EnumIndexesTbl
        SELECT ( lcEnumIndexesTbl )
        lcEnumRelsTbl=this.CreateWzTable( "Relation" )
        this.EnumRelsTbl=lcEnumRelsTbl

        USE ( p_dbc ) IN 0 AGAIN ALIAS mydbc
        USE ( p_dbc ) IN 0 AGAIN ALIAS mydbcpar
        SELECT mydbc
        lcExact=SET( "EXACT" )
        SET EXACT ON
        LOCATE FOR LOWER( objecttype )="table" AND TRIM( LOWER( objectname ))=="ridd"

        SCAN FOR UPPER( objecttype )="RELATION"

            GOTO ( mydbc.parentid ) IN mydbcpar
            l_chitable=LOWER( mydbcpar.objectname )
            l_start=1
            DO WHILE l_start<=LEN( property )
                l_size=ASC( SUBSTR( property, l_start, 1 ))+;
                    ( ASC( SUBSTR( property, l_start+1, 1 ))*256 )+;
                    ( ASC( SUBSTR( property, l_start+2, 1 ))*256^2 )+;
                    ( ASC( SUBSTR( property, l_start+3, 1 ))*256^3 )
                l_key=SUBSTR( property, l_start+6, 1 )
                l_value=SUBSTR( property, l_start+7, l_size-8 )
                DO CASE
                    CASE l_key==CHR( KEY_CHILDTAG )
                        l_chitag=l_value
                    CASE l_key==CHR( KEY_RELTABLE )
                        l_partable=LOWER( l_value )
                    CASE l_key==CHR( KEY_RELTAG )
                        l_partag=l_value
                ENDCASE
                l_start=l_start+l_size
            ENDDO

            l_area=SELECT( 1 )

            *Grab tag expression
            SELECT ( lcEnumIndexesTbl )

            *Parent table
            LOCATE FOR IndexName=RTRIM( l_partable ) AND TagName=RTRIM( l_partag )
            l_parexpr=&lcEnumIndexesTbl..RmtExpr

            *Child table
            LOCATE FOR IndexName=RTRIM( l_chitable ) AND TagName=RTRIM( l_chitag )
            l_chiexpr=&lcEnumIndexesTbl..RmtExpr

            *Translate ri characters ( RCI ) into words ( Restrict, Cascade, Ignore )
            l_delete=UPPER( SUBSTR( mydbc.riinfo, d_deletespot, 1 ))
            l_delete=IIF( l_delete<>CASCADE_CHAR_LOC AND l_delete<>RESTRICT_CHAR_LOC, ;
                IGNORE_CHAR_LOC, l_delete )
            l_update=UPPER( SUBSTR( mydbc.riinfo, d_updatespot, 1 ))
            l_update=IIF( l_update<>CASCADE_CHAR_LOC AND l_update<>RESTRICT_CHAR_LOC, ;
                IGNORE_CHAR_LOC, l_update )
            l_insert=UPPER( SUBSTR( mydbc.riinfo, d_insertspot, 1 ))
            l_insert=IIF( l_insert<>CASCADE_CHAR_LOC AND l_insert<>RESTRICT_CHAR_LOC, ;
                IGNORE_CHAR_LOC, l_insert )

            l_rmtpartable=this.RemotizeName( l_partable )
            l_rmtchitable=this.RemotizeName( l_chitable )

            *See if there are multiple relations between the same two tables
            SELECT COUNT( * ) FROM ( lcEnumRelsTbl ) ;
                WHERE RTRIM( Dd_RmtPar )==RTRIM( l_rmtpartable ) ;
                AND RTRIM( Dd_RmtChi )==RTRIM( l_rmtchitable ) ;
                INTO ARRAY aDupeCount

            INSERT INTO &lcEnumRelsTbl ( DD_CHILD, Dd_RmtChi, DD_PARENT, Dd_RmtPar, ;
                DD_CHIEXPR, DD_PAREXPR, Duplicates, dd_update, dd_insert, dd_delete ) ;
                VALUES ( l_chitable, l_rmtchitable, l_partable, l_rmtpartable, ;
                l_chiexpr, l_parexpr, aDupeCount, l_update, l_insert, l_delete )

            SELECT mydbc

        ENDSCAN

        *clean up
        USE
        SELECT mydbcpar
        USE
        SET EXACT &lcExact
        this.GetRiInfoRecalc=.F.

        SELECT( lnOldArea )



    #IF SUPPORT_ORACLE
    FUNCTION GetEligibleRels
        PARAMETERS aEligibleRels, llChoices
        LOCAL lcEnumRelsTbl, lcEnumTablesTbl, lnOldArea, lcDupeString, lcConstraint, ;
            aExportTables, lcExact

        lnOldArea=SELECT()
        lcEnumRelsTbl=RTRIM( this.EnumRelsTbl )
        lcEnumTablesTbl=RTRIM( this.EnumTablesTbl )
        *This determines if we return relations that have cluster names or that don't
        *If we want the choices of rels for clusters, we want the rels w/o cluster names
        IF llChoices THEN
            lcConstraint="Export =.T. AND EMPTY( ClustName )"
        ELSE
            lcConstraint="Export =.T. AND !EMPTY( ClustName )"
        ENDIF

        DIMENSION aExportTables[1]
        SELECT LOWER( TblName ) FROM ( lcEnumTablesTbl ) WHERE &lcConstraint INTO ARRAY aExportTables

        IF !EMPTY( aExportTables ) THEN
            SELECT ( lcEnumRelsTbl )
            I=1
            lcExact=SET( 'EXACT' )
            SET EXACT ON
            SCAN
                *check to see if both parent and child table are being exported
                *also make sure it's not a self-join
                IF ASCAN( aExportTables, RTRIM( DD_CHILD ))<>0 ;
                        AND ASCAN( aExportTables, RTRIM( DD_PARENT ))<>0 ;
                        AND !DD_PARENT==DD_CHILD THEN
                    *Make sure number of fields in primary and foreign keys is the same
                    DIMENSION aParentKeys[1], aChildKeys[1]
                    this.KeyArray( DD_PAREXPR, @aParentKeys )
                    this.KeyArray( DD_CHIEXPR, @aChildKeys )

                    IF ALEN( aParentKeys, 1 )=ALEN( aChildKeys, 1 ) THEN
                        DIMENSION aEligibleRels [i, 2]
                        IF &lcEnumRelsTbl..Duplicates <> 0 THEN
                            lcDupeString = "( " + LTRIM( STR( Duplicates+1 )) + " )"
                        ELSE
                            lcDupeString=""
                        ENDIF
                        aEligibleRels[i, 1]=RTRIM( DD_PARENT ) + ":" + RTRIM( DD_CHILD ) + lcDupeString
                        aEligibleRels[i, 2]=aEligibleRels[i, 1]
                        I=I+1
                    ENDIF
                ENDIF
            ENDSCAN
            SET EXACT &lcExact
        ENDIF
        SELECT ( lnOldArea )

#ENDIF



#IF SUPPORT_ORACLE
    FUNCTION ParseRel
        PARAMETERS lcRelName, lcParent, lcChild, lnDupID
        LOCAL lnOPos, lnCPos, lnColPos

        *
        *Takes a string of the form "customer:orders" and parses it into table names
        *Handles case where two tables may have multiple relations between each other
        *

        *Check if the rel is one of several with duplicate parent and child tables
        lnOPos=AT( "( ", lcRelName )
        lnCPos=AT( " )", lcRelName )
        IF lnOPos<>0 THEN
            lnDupID=VAL( SUBSTR( lcRelName, lnOPos+1, lnCPos-lnOPos-1 ))
            lcRelName=LEFT( lcRelName, lnOPos-1 )
        ELSE
            lnDupID=0
        ENDIF

        *Separate parent and child tablenames out of relation name
        lnColPos=AT( ":", lcRelName )
        lcParent=LEFT( lcRelName, lnColPos-1 )
        lcChild=SUBSTR( lcRelName, lnColPos+1 )

#ENDIF


    FUNCTION RedirectApp
        LOCAL lnDispLogin, lcOldConnString, lcRenamedUserConn, llRenamed

* If we have an extension object and it has a RedirectApp method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'RedirectApp', 5 ) and ;
			not this.oExtension.RedirectApp( This )
			return
		ENDIF

        *All the renaming gyrations are to avoid the problem of the user
        *getting a login dialog when remote views are created

		IF !EMPTY( this.UserConnection ) AND !this.PwdInDef
            *create new conndef; name is placed into this.ViewConnection
            this.CreateConnDef()
            DBSETPROP( this.ViewConnection, "connection", "ConnectString", this.CONNECTSTRING )
            lcRenamedUserConn=this.UniqueConnDefName()
            *rename old connection
            RENAME CONNECTION ( this.UserConnection ) TO ( lcRenamedUserConn )
            DBSETPROP( lcRenamedUserConn, "connection", "ConnectString", this.CONNECTSTRING )
            *give new connection the original's name
            RENAME CONNECTION ( this.ViewConnection ) TO ( this.UserConnection )
            this.ViewConnection=this.UserConnection
            llRenamed=.T.
        ELSE
            llRenamed=.F.
        ENDIF

        IF this.ExportViewToRmt AND this.DoUpsize THEN
            this.RemotizeViews()
        ENDIF

        IF this.ExportTableToView AND this.DoUpsize THEN
            this.CreateRmtViews()
        ENDIF

        *need to take out password if conndef created and user doesn't want pwd saved
        IF !this.ViewConnection=="" AND !this.ExportSavePwd AND !llRenamed THEN
            lcRConnString=DBGETPROP( this.ViewConnection, "connection", "connectstring" )
            lcRPwd=this.ParseConnectString( lcRConnString, "pwd=" )
            lcRConnString=STRTRAN( lcRConnString, "PWD="+lcRPwd+";" )
            lcRConnString=STRTRAN( lcRConnString, "PWD="+lcRPwd )
            lcRConnString=STRTRAN( lcRConnString, "pwd="+lcRPwd+";" )
            lcRConnString=STRTRAN( lcRConnString, "pwd="+lcRPwd )
            =DBSETPROP( this.ViewConnection, "connection", "connectstring", lcRConnString )
        ENDIF

        IF llRenamed
            *Delete temporary connection definition
            DELETE CONNECTION ( this.ViewConnection )
            RENAME CONNECTION ( lcRenamedUserConn ) TO ( this.UserConnection )
        ENDIF


    FUNCTION TrimDBCNameFromTables( lcTables )
        * trim off "tastrade!" from "tastrade!customer, tastrade!orders"
        LOCAL P, q, m.RetVal
        m.RetVal = ""
        DO WHILE "!" $ m.lcTables
            m.P = AT( "!", m.lcTables )
            m.q = AT( ", ", m.lcTables )
            IF m.q > 0
                m.RetVal = m.RetVal + SUBSTRC( m.lcTables, m.P+1, m.q - m.P )
                m.lcTables = SUBSTRC( m.lcTables, m.q+1 )
            ELSE
                m.RetVal = m.RetVal + SUBSTRC( m.lcTables, m.P+1 )
                m.lcTables = ""
            ENDIF
        ENDDO
        RETURN m.RetVal


    FUNCTION ProcessFromClause
        * parses the FROM clause, adds the odbc oj escape to the Sql statement and
        * returns the list of tables participating in this view
        PARAMETERS cStr, cTables
        LOCAL lcTalk, lcPath, llResult, aTables[1], lcLen

        * disable errors
        this.SetErrorOff = .T.
        this.HadError = .F.

        * needed by substr() to return empty strings at eos
        lcTalk = SET( 'talk' )
        SET TALK OFF
        lcProc = SET( 'proc' )
        SET PROCEDURE TO "wizjoin.prg"
        llResult = ParseFromClause( @cStr, @aTables )
        SET PROCEDURE TO &lcProc
        SET TALK &lcTalk

        * bail on error
        this.SetErrorOff = .F.
		IF this.HadError
            RETURN .F.
        ENDIF

        * build table list
        IF llResult
            cTables = ""
            lcLen = ALEN( aTables, 1 )
            FOR m.I = 1 TO lcLen
                cTables = cTables + aTables[m.i] + IIF( m.I < lcLen, ", ", "" )
            ENDFOR
        ENDIF

        RETURN llResult


    FUNCTION RemotizeViews
        LOCAL aTableNames, lcEnumTables, lcTablesInView, lcTablesNotUpsized, ;
            lcNewTableString, lcViewSQL, lcRmtViewSQL, lnOldArea, lcViewsTbl, ;
            llTableUpsized, aWhip, lcEnumFieldsTbl, aFldArray, lcNewViewName, lcDBCAlias, ;
            aFieldsInView, mm, ii, jj, I, llSendUpdates, lcDelStatus, lcViewErr, lcCRLF, ;
            lcErrString, lnParentRec, lnViewCount, lnServerError, lcErrMsg, lcRCursor, ;
            lcRCursor, lcViewParms, llShareConnection, lcDatatype, lcRmtViewSQL2

        PRIVATE aViews
        *
        *Views that consist of tables which were upsized are modified so they point to
        *remote data and are executed on the back end
        *

        *If the user isn't upsizing, bail
        IF !this.DoUpsize THEN
            RETURN
        ENDIF

        lnOldArea=SELECT()
        lcCRLF=CHR( 10 )+CHR( 13 )
        *Check each view; see if all its tables were upsized
        *Get array of views; if no views, bail
        lnViewCount=ADBOBJECTS( aViews, "View" )
        IF lnViewCount=0 THEN
            RETURN
        ENDIF
        FOR I =1 TO ALEN( aViews, 1 )
            aViews[i] = LOWER( aViews[i] )
        ENDFOR
        this.InitTherm( RMTZING_VIEW_LOC, lnViewCount, 0 )
        lnViewCount=0

        *Create array of tables that were successfully exported, local names and remote names )
        *Sort them so the longest fields are first, otherwise string replacements could
        *get messed up
        lcEnumTables=this.EnumTablesTbl
        DIMENSION aTableNames[1]
        SELECT LOWER( TblName ), ;
            LOWER( RmtTblName ), ;
            ISNULL( TblName ), ;
            1/LEN( ALLTRIM( TblName )) AS foo ;
            FROM &lcEnumTables WHERE Exported=.T. ;
            ORDER BY foo ;
            INTO ARRAY aTableNames
        IF EMPTY( aTableNames )
			raiseevent( This, 'CompleteProcess' )
            RETURN
        ENDIF

        lcViewsTbl=this.CreateWzTable( "Views" )
        this.ViewsTbl=lcViewsTbl
        llViewConnCreated=.F.

        llShareConnection = .F.
        oReg = NEWOBJECT( "FoxReg", home() + "ffc\registry.vcx" )
        cOptionValue = ""
        cOptionName = "CrsShareConnection"
        m.nErrNum = oReg.GetFoxOption( m.cOptionName, @cOptionValue )
        IF nErrNum=0 AND cOptionValue="1"
            llShareConnection = .T.
        ENDIF

        FOR I=1 TO ALEN( aViews, 1 )

            *Thermometer stuff
            lcMessage=STRTRAN( THIS_VIEW_LOC, "|1", LOWER( RTRIM( aViews[i] )) )
            this.UpDateTherm( lnViewCount, lcMessage )
            lnViewCount=lnViewCount+1

            *If this is a remote view, skip it
            IF DBGETPROP( aViews[i], "View", "SourceType" )=2 THEN
                LOOP
            ENDIF

			*{Add JEI RKR 2005.03.25
			* Check for current View is Choosen for Upsizing
			IF !this.CheckUpsizeView( aViews[i] )
				LOOP
			ENDIF
			*}Add JEI RKR 2005.03.25

            *flag set true if at least one table in a view was upsized;
            *we can ignore the view if this stays .F.
            llTableUpsized=.F.

            *Grab the view's SQL string ( may need to change table names to remote versions )
            lcViewSQL=LOWER( DBGETPROP( aViews[i], "View", "SQL" ))

			&&chandsr added comments
			&&SANITIZE THE SQL STATEMENT BY REPLACING .F. and .T. WITH 0 and 1 RESPECTIVELY.
			&& Go through the FOX statements and make the statement ANSI92 compatible. This means replacing the occurrences of boolean expression with
			&& ANSI92 compatible boolean expressions.
            lcViewSQL = this.MakeStringANSI92Compatible ( lcViewSQL )
			&&chandsr added comments end.

            * Get parameter list
            lcViewParms=LOWER( DBGETPROP( aViews[i], "View", "Parameterlist" ))

            * try parsing the from clause to get the tables and remotize oj
            * on error get table list from Tables property
            IF !this.ProcessFromClause( @lcViewSQL, @lcTablesInView )
                lcTablesInView = LOWER( DBGETPROP( aViews[i], "View", "Tables" ))
            ENDIF

            lcTablesInView = lcTablesInView + ", "
            * add comma so there's no confusing strings like "dupe" and "dupe1"

            *will not remove free tables from the view list
            lcDBName=LOWER( this.JustStem2( this.SourceDB ))
            lcTablesNotUpsized = STRTRAN( m.lcTablesInView, lcDBName+"!" )

            *Remove local database name from sql string
            *If it has a space in it, it will be enclosed in quotes
            IF AT( lcDBName, " " )<>0 THEN
                lcRmtViewSQL=STRTRAN( lcViewSQL, gcQT+lcDBName+"!"+gcQT )
            ELSE
                lcRmtViewSQL=STRTRAN( lcViewSQL, lcDBName+"!" )
            ENDIF

            *Check which ( if any ) tables were upsized
            lcNewTableString=""

            DIMENSION aFieldsInView[1, 2]
            aFieldsInView=.F.

            FOR ii=1 TO ALEN( aTableNames, 1 )
                IF LOWER( RTRIM( aTableNames[ii, 1] )) + ", " $ lcTablesInView THEN
                    *set flag
                    llTableUpsized=.T.

                    *replace all field names changed by remotizing names
                    lcEnumFieldsTbl=this.EnumFieldsTbl
                    DIMENSION aFldArray[1, 2]
                    aFldArray=.F.
                    SELECT FldName, RmtFldname FROM ( lcEnumFieldsTbl ) ;
                        WHERE TblName=RTRIM( aTableNames[ii, 1] ) ;
                        AND FldName<>RmtFldname INTO ARRAY aFldArray

                    IF !EMPTY( aFldArray ) THEN
                        FOR jj=1 TO ALEN( aFldArray, 1 )
                            IF RTRIM( aFldArray[jj, 1] ) $ lcRmtViewSQL THEN
                                lcRmtViewSQL=STRTRAN( lcRmtViewSQL, RTRIM( aFldArray[jj, 1] ), RTRIM( aFldArray[jj, 2] ))
                            ENDIF
                        NEXT jj
                    ENDIF

                    *Replace table names with remotized table names
                    *First look for stuff of the form "from me" ( change to "from_me" )
                    *then look for "from" ( change to "from_" ).  This prevent "from me"
                    *becoming "from__me".

                    lcTablesNotUpsized=STRTRAN( lcTablesNotUpsized, gcQT+RTRIM( aTableNames[ii, 1] )+", "+gcQT )
                    lcTablesNotUpsized=STRTRAN( lcTablesNotUpsized, RTRIM( aTableNames[ii, 1] )+", " )
                    lcRmtViewSQL=STRTRAN( lcRmtViewSQL, gcQT+RTRIM( aTableNames[ii, 1] )+gcQT, ;
                        RTRIM( aTableNames[ii, 2] ))
                    lcRmtViewSQL=STRTRAN( lcRmtViewSQL, " " + RTRIM( aTableNames[ii, 1] ), ;
                        " " + RTRIM( aTableNames[ii, 2] ))

                    IF !lcNewTableString=="" THEN
                        lcNewTableString=lcNewTableString+", "
                    ENDIF
                    lcNewTableString=lcNewTableString + RTRIM( aTableNames[ii, 2] )

                    aTableNames[ii, 3]=.T.	&&Table is part of view

                ENDIF
            NEXT ii

            *Go on to the next view if this one has no tables upsized in it
            IF !llTableUpsized THEN
                LOOP
            ENDIF

            *If all the tables were upsized, then remotize the view
            IF ALLTRIM( STRTRAN( lcTablesNotUpsized, ", " ))=="" THEN

                *Need a connection to associate with remote view
                *create one if user didn't connect with one
                *or if they created a new database
                IF this.ViewConnection=="" THEN
                    IF this.UserConnection=="" OR this.CreateNewDB THEN
                        this.CreateConnDef()
                    ELSE
                        this.ViewConnection=this.UserConnection
                    ENDIF
                ENDIF

                *Rename local view
                lcNewViewName=this.UniqueTorVName( aViews[i] )
                RENAME VIEW ( RTRIM( aViews[i] )) TO ( lcNewViewName )

                * Check for "==" used in SQL Server which is not allowed and replace with "="
                lcRmtViewSQL2 = lcRmtViewSQL
                IF this.SQLServer AND ATCC( "==", lcRmtViewSQL )#0
                    lcRmtViewSQL2 = STRTRAN( lcRmtViewSQL, "==", "=" )
                ENDIF

                *Create remote view with same name
                this.SetErrorOff=.T.
                this.HadError=.F.
                CREATE SQL VIEW RTRIM( aViews[i] ) REMOTE CONNECT RTRIM( this.ViewConnection ) ;
                    AS &lcRmtViewSQL2
                this.SetErrorOff=.F.
                IF this.HadError THEN
                    *If the view creation fails, store the error and loop
                    =AERROR( aErrArray )
                    lnServerError=aErrArray[1]
                    lcErrMsg=aErrArray[2]
                    IF lnServerError=1526 THEN
                        lnServerError=aErrArray[5]
                    ENDIF

                    *Put things back and store error info
                    RENAME VIEW ( lcNewViewName ) TO ( RTRIM( aViews[i] ))
                    lcViewErr=aErrArray[2]
                    SELECT ( lcViewsTbl )
                    APPEND BLANK
                    REPLACE ViewName WITH aViews[i], NewName WITH aViews[i], ;
                        ViewSQL WITH lcViewSQL, RmtViewSQL WITH lcRmtViewSQL, ;
                        Remotized WITH .F., ViewErr WITH lcErrMsg, ;
                        ViewErrNo WITH lnServerError
                    this.StoreError( lnServerError, lcErrMsg, lcRmtViewSQL, RMTIZE_VIEW_FAILED_LOC, aViews[i], VIEW_LOC )
                    LOOP I
                ENDIF

                *Set various and sundry view properties
                llSendUpdates=DBGETPROP( lcNewViewName, "view", "SendUpdates" )
                lcViewErr=""
                IF !DBSETPROP( RTRIM( aViews[i] ), "view", "SendUpdates", llSendUpdates )
                    this.AddToError( @lcViewErr, UPDATE_PROP_FAILED_LOC )
                ENDIF
                IF !DBSETPROP( RTRIM( aViews[i] ), "view", "Tables", lcNewTableString )
                    this.AddToError( @lcViewErr, TABLES_PROP_FAILED_LOC )
                ENDIF
                IF !EMPTY( lcViewParms ) AND !DBSETPROP( RTRIM( aViews[i] ), "view", "Parameterlist", lcViewParms )
                    this.AddToError( @lcViewErr, TABLES_PROP_FAILED_LOC )
                ENDIF

                IF llShareConnection AND !DBSETPROP( RTRIM( aViews[i] ), "view", "ShareConnection", llShareConnection )
                    this.AddToError( @lcViewErr, TABLES_PROP_FAILED_LOC )
                ENDIF

                *Get properties of fields in local view; set remote field properties to same

                *Get cursors of fields in local view and for remote view
                lcDBCAlias=this.UniqueCursorName( "dbcalias" )
                lcDelStatus=SET( "deleted" )
                SET DELETED ON
                USE RTRIM( this.SourceDB ) IN 0 AGAIN ALIAS &lcDBCAlias
                SELECT ( lcDBCAlias )

                LOCATE FOR RTRIM( LOWER( objectname ))==RTRIM( LOWER( lcNewViewName ))
                lcParent=&lcDBCAlias..ObjectID
                lcLCursor=this.UniqueCursorName( "localfields" )
                SELECT 0
                SELECT objectname, RECNO() FROM ( lcDBCAlias ) WHERE parentid=lcParent AND objecttype="Field" ;
                    INTO CURSOR &lcLCursor
                SELECT( lcLCursor )
                lnLFields=RECCOUNT()

                *Find the record for the remote view
                SELECT ( lcDBCAlias )
                LOCATE FOR LOWER( RTRIM( objectname ))==LOWER( RTRIM( aViews[i] ))
                lcParent=&lcDBCAlias..ObjectID
                lcRCursor=this.UniqueCursorName( "remotefields" )
                SELECT 0
                *Only get fields which aren't timestamps added by the wizard
                SELECT objectname, RECNO() FROM ( lcDBCAlias ) ;
                    WHERE parentid=lcParent ;
                    AND objecttype="Field" ;
                    AND ATC( IDENTCOL_LOC, RTRIM( objectname )) == 0 ;
                    AND ATC( TIMESTAMP_LOC, RTRIM( objectname )) == 0 ;
                    INTO CURSOR &lcRCursor
                SELECT( lcRCursor )
                lnRFields=RECCOUNT()

                IF lnRFields<>lnLFields
                    this.AddToErr( @lcViewErr, FIELDS_UNEQUAL_LOC )
                ELSE
                    SELECT ( lcLCursor )
                    SCAN
                        llKeyField=DBGETPROP( lcNewViewName+"."+RTRIM( objectname ), "field", "keyfield" )
                        llUpdatable=DBGETPROP( lcNewViewName+"."+RTRIM( objectname ), "field", "updatable" )
                        lcUpdateName=DBGETPROP( lcNewViewName+"."+RTRIM( objectname ), "field", "updatename" )
                        lcDatatype=DBGETPROP( lcNewViewName+"."+RTRIM( objectname ), "field", "datatype" )

                        SELECT ( lcRCursor )
                        IF !DBSETPROP( RTRIM( aViews[i] )+"."+ RTRIM( objectname ), "field", "keyfield", llKeyField )
                            lcRmtViewField=RTRIM( objectname )
                            lcErrString=STRTRAN( KEYFIELD_PROP_FAILED_LOC, '|1', lcRmtViewField )
                            this.AddToErr( @lcViewErr, lcErrString )
                        ENDIF
                        IF !DBSETPROP( RTRIM( aViews[i] )+"."+ RTRIM( objectname ), "field", "updatable", llUpdatable )
                            lcRmtViewField=RTRIM( objectname )
                            lcErrString=STRTRAN( UPDATABLE_PROP_FAILED_LOC, '|1', lcRmtViewField )
                            this.AddToErr( @lcViewErr, lcErrString )
                        ENDIF

                        *Since we map Date->DateTime field in SQL Server, we need to make sure the new
                        *remote view resets DateTime->Date.
                        IF lcDatatype="D" AND !DBSETPROP( RTRIM( aViews[i] )+"."+ RTRIM( objectname ), "field", "datatype", lcDatatype )
                            this.AddToErr( @lcViewErr, lcErrString )
                        ENDIF
                        SKIP
                        SELECT ( lcLCursor )
                    ENDSCAN
                ENDIF
                USE
                SET DELETED &lcDelStatus
                USE IN ( lcDBCAlias )
                IF USED( m.lcRCursor )
                    USE IN ( m.lcRCursor )
                ENDIF

                *store all this stuff
                SELECT ( lcViewsTbl )
                APPEND BLANK
                REPLACE ViewName WITH aViews[i], NewName WITH lcNewViewName, ;
                    ViewSQL WITH lcViewSQL, RmtViewSQL WITH lcRmtViewSQL, ;
                    CONNECTION WITH this.ViewConnection, Remotized WITH .T., ;
                    ViewErr WITH lcViewErr, TblsUpszd WITH lcNewTableString
                IF !lcViewErr=="" THEN
                    this.StoreError( .NULL., "", "", VIEW_PROPS_FAILED_LOC, aViews[i], VIEW_LOC )
                ENDIF
            ELSE
                *create remote views of tables that were upsized and leave the current view alone
                *Don't create remote view here if the user wants all upsized tables
                *to have remote views created for them
                IF !this.ExportTableToView THEN
                    FOR zz=1 TO ALEN( aTableNames, 1 )
                        IF aTableNames[zz, 3]=.T. THEN
                            *Function expects array even though we'll always pass just one element from here
                            DIMENSION aWhip[1, 2]
                            aWhip[1, 1]=aTableNames[zz, 1]
                            aWhip[1, 2]=aTableNames[zz, 2]
                            this.CreateRmtViews( @aWhip )
                        ENDIF
                    NEXT zz
                ENDIF

                *Store stuff for report
                SELECT ( lcViewsTbl )
                APPEND BLANK
                REPLACE ViewName WITH aViews[i], NewName WITH "", ;
                    Remotized WITH .F., TblsUpszd WITH lcNewTableString, ;
                    NotUpszd WITH LTRIM( STRTRAN( lcTablesNotUpsized, ", " ))

            ENDIF

        NEXT I
		raiseevent( This, 'CompleteProcess' )

        *Get rid of table if no records were added; report will not be added either
        IF RECCOUNT()=0 THEN
            this.DisposeTable( lcViewsTbl, "Delete" )
            this.ViewsTbl=""
        ENDIF

        SELECT ( lnOldArea )



    FUNCTION AddToErr
        PARAMETERS lcErrString, lcAddToString

        IF lcErrString==""
            lcErrString=lcAddToString
        ELSE
            lcErrString=CHR( 10 )+CHR( 13 )+lcAddToString
        ENDIF



    FUNCTION UniqueConnDefName
        LOCAL lcConnName, I, lcExact

        *Make sure name is unique
        lcConnName=CONN_NAME_LOC

        *{Add JEI RKR 2005.03.28
        * IF in Database no existing connection
        LOCAL aConnDefs[1, 1]
        aConnDefs[1, 1] = SYS( 2015 )
        *}Add JEI RKR 2005.03.28

        =ADBOBJECTS( aConnDefs, "connection" )
        I=1
        lcExact=SET( 'EXACT' )
        SET EXACT ON
        DO WHILE ASCAN( aConnDefs, UPPER( lcConnName ))<>0
            lcConnName=CONN_NAME_LOC+LTRIM( STR( I ))
            I=I+1
        ENDDO
        SET EXACT &lcExact
        RETURN lcConnName



    FUNCTION CreateConnDef
        LOCAL I, llFound, lcConnString, lcConnName, lcNewConnString

        lcConnName=this.UniqueConnDefName()
        this.ViewConnection=lcConnName
        IF this.UserConnection=="" THEN

            *If the user didn't connect with a connection definition, we
            *need to create one from scratch based on the connection string
            *created when they logged into ODBC

            lcConnString=RTRIM( this.CONNECTSTRING )

            *parse connection string to get user id, password, and database
            IF !EMPTY( this.DataSourceName ) && Add JEI RKR 2005.03.28
	            lcNewConnString="dsn="+RTRIM( this.DataSourceName )+";"

	            llFound=.F.
	            lcUserID=this.ParseConnectString( this.CONNECTSTRING, "uid=", @llFound )

	            IF llFound THEN
	                lcNewConnString=lcNewConnString+"uid="+RTRIM( lcUserID )+";"
	            ENDIF

	            lcPassword=this.ParseConnectString( this.CONNECTSTRING, "pwd=", llFound )
	            IF llFound THEN
	                lcNewConnString=lcNewConnString+"pwd="+RTRIM( lcPassword )+";"
	            ENDIF

	            *add database name only for SQL Server
	            IF this.ServerType <> ORACLE_SERVER
	                lcNewConnString=lcNewConnString+"database="+RTRIM( this.ServerDBName )+";"
	            ENDIF
	            *Create new connection and set its connection string property
	            CREATE CONNECTION &lcConnName DATASOURCE RTRIM( this.DataSourceName )
	            =DBSETPROP( lcConnName, "connection", "ConnectString", lcNewConnString )
	        ELSE
	        	*{ Add RKR 2005.03.28
	        	IF EMPTY( this.ParseConnectString( lcConnString, "database=", llFound )) And !EMPTY( this.ServerDBName )
	        		 lcConnString = lcConnString + IIF( RIGHT( lcConnString, 1 )<> ";", ";", "" ) + "database="+RTRIM( this.ServerDBName )+";"
	        	ENDIF
				CREATE CONNECTION &lcConnName CONNSTRING ( lcConnString )
	        	*} Add RKR 2005.03.28
			ENDIF
        ELSE

            *user specified a connection def so new conndef should be just like it
            *except for the database specified

            *Get all properties ( except datasource which we already know )
            llAsynchronous=		DBGETPROP( this.UserConnection, "connection", "Asynchronous" )
            llBatchMode=		DBGETPROP( this.UserConnection, "connection", "BatchMode" )
            lcConnectString=	DBGETPROP( this.UserConnection, "connection", "ConnectString" )
            lcConnectTimeOut=	DBGETPROP( this.UserConnection, "connection", "ConnectTimeOut" )
            lnDispLogin=		DBGETPROP( this.UserConnection, "connection", "DispLogin" )
            llDispWarnings=		DBGETPROP( this.UserConnection, "connection", "DispWarnings" )
            lnIdleTimeOut=		DBGETPROP( this.UserConnection, "connection", "IdleTimeOut" )
            lcPassword=			DBGETPROP( this.UserConnection, "connection", "PassWord" )
            lnQueryTimeout=		DBGETPROP( this.UserConnection, "connection", "QueryTimeout" )
            lnTransactions=		DBGETPROP( this.UserConnection, "connection", "Transactions" )
            lcUserID=			DBGETPROP( this.UserConnection, "connection", "UserID" )
            lnWaitTime=			DBGETPROP( this.UserConnection, "connection", "WaitTime" )

            *Create bare-bones connection
            CREATE CONNECTION &lcConnName DATASOURCE RTRIM( this.DataSourceName )

            *Hack connection string so it points at new database
            lcOldDatabase=this.ParseConnectString( lcConnectString, "database=" )
            IF !lcOldDatabase=="" THEN
                lcConnectString=STRTRAN( lcConnectString, lcOldDatabase, RTRIM( this.ServerDBName ))
            ENDIF

            *Make new connection's properties just like this.UserConnection's
            =DBSETPROP( this.ViewConnection, "connection", "Asynchronous", llAsynchronous )
            =DBSETPROP( this.ViewConnection, "connection", "BatchMode", llBatchMode )
            =DBSETPROP( this.ViewConnection, "connection", "ConnectString", lcConnectString )
            =DBSETPROP( this.ViewConnection, "connection", "ConnectTimeOut", lcConnectTimeOut )
            =DBSETPROP( this.ViewConnection, "connection", "DispLogin", lnDispLogin )
            =DBSETPROP( this.ViewConnection, "connection", "DispWarnings", llDispWarnings )
            =DBSETPROP( this.ViewConnection, "connection", "IdleTimeOut", lnIdleTimeOut )
            =DBSETPROP( this.ViewConnection, "connection", "PassWord", lcPassword )
            =DBSETPROP( this.ViewConnection, "connection", "QueryTimeout", lnQueryTimeout )
            =DBSETPROP( this.ViewConnection, "connection", "Transactions", lnTransactions )
            =DBSETPROP( this.ViewConnection, "connection", "UserID", lcUserID )
            =DBSETPROP( this.ViewConnection, "connection", "WaitTime", lnWaitTime )

        ENDIF

        this.ViewConnection=lcConnName



    FUNCTION ParseConnectString
        PARAMETERS lcConnString, lcStringPart, llFound
        LOCAL lcLowConnString, lnStartPos

        *Takes a connection string; returns the part of string asked for

        lcStringPart=LOWER( lcStringPart )
        lcLowConnString=LOWER( lcConnString )
        lnStartPos=AT( lcStringPart, lcLowConnString )
        IF lnStartPos<>0 THEN
            lcFoundString=SUBSTR( lcConnString, lnStartPos+LEN( lcStringPart ))
            IF AT( ";", lcFoundString )<>0 THEN
                lcFoundString=LEFT( lcFoundString, AT( ";", lcFoundString )-1 )
            ENDIF
            llFound=.T.
        ELSE
            lcFoundString=""
            llFound=.F.
        ENDIF

        RETURN lcFoundString



    FUNCTION CreateRmtViews
        PARAMETERS aTableNames

        *
        *Creates remote views for tables passed in or all tables that were upsized
        *

        LOCAL lcEnumTablesTbl, lcNewTblName, lnOldArea, lcSQL, lcEnumIndexesTbl, ;
            aPkey, lcEnumFieldsTbl, lcViewErr, lcTblError, lnErrNo, llShowTherm, lnTableCount, I, ;
            lcCRLF, lcViewExtension, lcViewName
        lnOldArea=SELECT()
        lcEnumTablesTbl=this.EnumTablesTbl
        lcEnumIndexesTbl=this.EnumIndexesTbl
        lcEnumFieldsTbl=this.EnumFieldsTbl
        lcCRLF=CHR( 10 )+CHR( 13 )

        IF EMPTY( aTableNames )
            DIMENSION aTableNames[1]
            SELECT TblName, RmtTblName FROM ( lcEnumTablesTbl ) WHERE Exported=.T. ;
                INTO ARRAY aTableNames

            IF EMPTY( aTableNames )
                RETURN	&& no tables were actually upsized
            ENDIF

            lnTableCount=ALEN( aTableNames, 1 )
            llShowTherm=.T.
        ENDIF

        IF llShowTherm THEN
            *Only display thermometer if this method is called by ProcessOutput;
            *otherwise the thermometer will be showing the progress
            *for the RemotizeView method already
            this.InitTherm( RMTZING_TABLE_LOC, lnTableCount, 0 )
            lnTableCount=0
        ENDIF

        IF this.ViewConnection=="" THEN
            IF this.UserConnection=="" OR this.CreateNewDB THEN
                this.CreateConnDef
            ELSE
                this.ViewConnection=this.UserConnection
            ENDIF
        ENDIF

        SELECT ( lcEnumTablesTbl )
        FOR I=1 TO ALEN( aTableNames, 1 )
            IF llShowTherm THEN
                lcMessage=STRTRAN( THIS_TABLE_LOC, "|1", RTRIM( aTableNames[i, 1] ))
                this.UpDateTherm( lnTableCount, lcMessage )
                lnTableCount=lnTableCount+1
            ENDIF

            this.HadError=.F.
            this.MyError=0
            lcViewErr=""
            lcTblError="" && jvf 9/25/99 for table drop
            lnTblErrNo=0 && jvf 9/25/99 for drop table
            lcViewExtension=IIF( this.ViewPrefixOrSuffix=3, "", RTRIM( this.ViewNameExtension ))

            * rename or drop the original table
            * jvf: 08/16/99 Added drop functionality
            lcNewTblName=RTRIM( aTableNames[i, 1] )
            IF this.DropLocalTables
                * jvf 9/15/99
                * Can't drop table unless we drop its relation first.
                * But we can't easily tell if it's in a relation. If the table to drop is on the
                * child side, easy to see that in the dbc and drop if needed. But if the table's the
                * parent side of the relation, like customer is to orders, it's hard to know what relation
                * to delete. You actually have to delete the relation of the order on cust_id. So, we'll just
                * trap for error on drop, log it, and move on.
                DROP TABLE ( RTRIM( aTableNames[i, 1] ))
                IF this.HadError && probably error 1577: table is in a relation
                    lcTblError=CANT_DROP2_LOC+" "+MESSAGE()
                    lnTblErrNo=this.MyError
                ELSE
                    lcNewTblName = DROPPED_TABLE_STATUS_LOC  && Table Dropped
                ENDIF
            ENDIF
            * If couldn't drop table, rename it if they didn't choose a prefix/suffix for
            * the new remote view name.
            * If the user chose not to drop tables, but did not to give the view a suffix/prefix,
            * need to add _local to the table name to prevent duplicate object names in dbc.
            IF ( NOT this.DropLocalTables OR lnTblErrNo>0 ) AND EMPTY( lcViewExtension )
                lcNewTblName=this.UniqueTorVName( RTRIM( aTableNames[i, 1] ))
                RENAME TABLE ( RTRIM( aTableNames[i, 1] )) TO ( lcNewTblName )
            ENDIF

            lcSQL="SELECT * FROM " + aTableNames[i, 2]
            *create the view
            * jvf 08/16/99 Add prefix/suffix to new view name

            lcViewName = IIF( this.ViewPrefixOrSuffix=1, lcViewExtension+RTRIM( aTableNames[i, 1] ), ;
                RTRIM( aTableNames[i, 1] )+lcViewExtension )

            CREATE SQL VIEW &lcViewName REMOTE CONNECT RTRIM( this.ViewConnection ) ;
                AS &lcSQL

            *See if table has primary key or candidate
            DIMENSION aPkey[1]
            aPkey=.F.
            SELECT RmtExpr FROM ( lcEnumIndexesTbl ) WHERE RTRIM( IndexName )==RTRIM( aTableNames[i, 1] ) ;
                AND LclIdxType="Primary key" INTO ARRAY aPkey

            *If not but unique index is available, use that
            IF EMPTY( aPkey ) THEN
                SELECT RmtExpr FROM ( lcEnumIndexesTbl ) WHERE RTRIM( IndexName )==RTRIM( aTableNames[i, 1] ) ;
                    AND RmtType="UNIQUE" ;
                    INTO ARRAY aPkey
            ENDIF

            IF !EMPTY( aPkey ) THEN
                *Make the whole thing updatable
                * jvf 08/16/99 just using var here now instead of array element
                IF !DBSETPROP( lcViewName, "view", "SendUpdates", .T. ) THEN
                    lcViewErr=UPDATE_PROP_FAILED_LOC
                ENDIF

                *Set the keyfields
                DIMENSION aKeyFields[1]
                aKeyFields=.F.
                this.KeyArray( aPkey, @aKeyFields )
                FOR ii=1 TO ALEN( aKeyFields, 1 )
                    * jvf 08/16/99 just using lcViewName var here now instead of array element
                    * JEI RKR 2005.07.14 add ChrTran
                    IF !DBSETPROP( lcViewName+"."+CHRTRAN( RTRIM( aKeyFields[ii] ), "[]", "" ), "field", "keyfield", .T. ) THEN
                        lcErrString=STRTRAN( KEYFIELD_PROP_FAILED_LOC, '|1', RTRIM( aKeyFields[ii] ))
                        lcViewErr=lcViewErr+lcCRLF+lcErrString
                    ENDIF
                NEXT ii

                *Make all the fields updatable
                DIMENSION aFldNames[1]
                aFldNames=.F.
                SELECT RmtFldname FROM ( lcEnumFieldsTbl ) WHERE RTRIM( TblName )==RTRIM( aTableNames[i, 1] ) ;
                    INTO ARRAY aFldNames
                IF !EMPTY( aFldNames ) THEN
                    FOR ii=1 TO ALEN( aFldNames, 1 )
                        * jvf 08/16/99 just using lcViewName var here now instead of array element
                        IF !DBSETPROP( lcViewName+"."+RTRIM( aFldNames[ii] ), "field", "updatable", .T. ) THEN
                            lcErrString=STRTRAN( UPDATABLE_PROP_FAILED_LOC, '|1', RTRIM( aKeyFields[ii] ))
                            lcViewErr=lcViewErr+lcCRLF+lcErrString
                        ENDIF
                    NEXT ii
                ENDIF

            ELSE

                lcViewErr=NO_UNIQUEKEY_LOC

            ENDIF

            *store these table and view names somewhere
            LOCATE FOR LOWER( RTRIM( &lcEnumTablesTbl..TblName ))==LOWER( RTRIM( aTableNames[i, 1] ))
            * jvf 08/16/99 just using lcNewTblName var here now instead of array element
            REPLACE NewTblName WITH lcNewTblName, RmtView WITH lcViewName, ;
                ViewErr WITH lcViewErr, TblError WITH lcTblError, TblErrNo WITH lnTblErrNo

            lcViewErr=""
            lcTblError=""
            lnTblErrNo=0

        NEXT I

        SELECT ( lnOldArea )

		raiseevent( This, 'CompleteProcess' )



    FUNCTION UniqueTorVName
        PARAMETERS lcTableName
        LOCAL lcNewTableName, lcOldTableName, I, lcExact

        *Need to make sure that when renaming tables and views that we don't overwrite
        *existing ones; this function returns a unique name

        DIMENSION aViewsArr[1]
        =ADBOBJECTS( aViewsArr, 'view' )
        =ADBOBJECTS( aTablesArr, 'table' )
        lcOldTableName=lcTableName
        lcNewTblName=LEFT( RTRIM( lcTableName ), MAX_FIELDNAME_LEN-LEN( LOCAL_SUFFIX_LOC ))+LOCAL_SUFFIX_LOC
        I=1
        lcExact = SET( 'EXACT' )
        SET EXACT ON
        DO WHILE ASCAN( aViewsArr, UPPER( lcNewTblName ))<>0 ;
                OR ASCAN( aTablesArr, UPPER( lcNewTblName ))<>0
            IF LEN( lcNewTblName )+ LEN( LTRIM( STR( I )) )>=MAX_FIELDNAME_LEN THEN
                lcOldTableName=LEFT( (lcTableName ), LEN( LTRIM( STR( I )) ))
                lcNewTblName=RTRIM( lcOldTableName )+LTRIM( STR( I ))+ LOCAL_SUFFIX_LOC
            ELSE
                *Just stick a number on the end
                lcNewTblName=RTRIM( lcOldTableName )+LOCAL_SUFFIX_LOC+LTRIM( STR( I ))
            ENDIF
            I=I+1
        ENDDO
        SET EXACT &lcExact

        RETURN lcNewTblName



    FUNCTION BuildReport
        LOCAL I, lcNewDB, lcNewProj, lcMisc, lcErrTblName, lcPath, lcPathAndFile, ;
            aRepArray
        *
        *Creates a project with report or just export error tables if the user
        *didn't ask for a report
        *
        IF ( this.DataErrors OR !this.ErrTbl=="" ) AND !this.DoReport THEN
            IF this.lQuiet or !MESSAGEBOX( DATA_ERRORS_LOC, ICON_EXCLAMATION+YES_NO_BUTTONS, TITLE_TEXT_LOC )=USER_YES
                this.SaveErrors=.F.
            ELSE
                this.SaveErrors=.T.
            ENDIF
        ELSE
            IF this.DoReport THEN
                this.SaveErrors=.T.
            ELSE
                this.SaveErrors=.F.
            ENDIF
        ENDIF

        *Bail if nothing to do
        IF !this.DoReport AND !this.SaveErrors AND !this.DoScripts THEN
            RETURN
        ENDIF

* If we have an extension object and it has a BuildReport method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'BuildReport', 5 ) and ;
			not this.oExtension.BuildReport( This )
			return
		ENDIF

        IF this.DoReport THEN
            this.InitTherm( BUILDING_REPORT_LOC, 0, 0 )

            *Can't use table name as primary and foreign keys because it's too long
            *Stuff integers into fields
            this.PurgeTable( this.EnumTablesTbl, "Export=.F. OR Upsizable = .F." )
            this.ReorderTable()
            this.Integerize

            *Close analysis tables
            this.DisposeTable( this.MappingTable, "close" )
            this.DisposeTable( this.ViewsTbl, "close" )

            *Get rid of records where stuff wasn't chosen for upsizing by user
            this.PurgeTable2	&&Deals with this.EnumFieldsTbl
            USE IN ( this.EnumTablesTbl )

            IF !this.EnumRelsTbl=="" THEN
                this.PurgeTable( this.EnumRelsTbl, "Exported=.F." )
            ENDIF
            IF !this.EnumIndexesTbl=="" THEN
                this.PurgeTable( this.EnumIndexesTbl, "TblUpszd=.F." )
            ENDIF

        ELSE
            IF this.DoScripts THEN
                this.InitTherm( SCRIPT_INDB_LOC, 0, 0 )
            ELSE
                this.InitTherm( PREP_ERR_LOC, 0, 0 )
            ENDIF
        ENDIF

* Create the reports folder if it doesn't exist, copy all of our tables to it,
* then change to that folder so all files are created there.

		IF not sys( 5 ) + upper( curdir() ) == upper( addbs( this.ReportDir ))
			IF not directory( this.ReportDir )
				md ( this.ReportDir )
			ENDIF
			copy file ( this.EnumFieldsTbl + '.*' ) to ;
				forcepath( this.EnumFieldsTbl + '.*', this.ReportDir )
			erase ( this.EnumFieldsTbl + '.*' )
			copy file ( this.EnumTablesTbl + '.*' ) to ;
				forcepath( this.EnumTablesTbl + '.*', this.ReportDir )
			erase ( this.EnumTablesTbl + '.*' )
			IF this.ExportIndexes
				copy file ( this.EnumIndexesTbl + '.*' ) to ;
					forcepath( this.EnumIndexesTbl + '.*', this.ReportDir )
				erase ( this.EnumIndexesTbl + '.*' )
			ENDIF
			IF this.ExportRelations
				copy file ( this.EnumRelsTbl + '.*' ) to ;
					forcepath( this.EnumRelsTbl + '.*', this.ReportDir )
				erase ( this.EnumRelsTbl + '.*' )
			ENDIF
			IF not empty( this.ViewsTbl )
				copy file ( this.ViewsTbl + '.*' ) to ;
					forcepath( this.ViewsTbl + '.*', this.ReportDir )
				erase ( this.ViewsTbl + '.*' )
			ENDIF
			IF this.DoScripts
				this.DisposeTable( this.ScriptTbl, 'close' )
				copy file ( this.ScriptTbl + '.*' ) to ;
					forcepath( this.ScriptTbl + '.*', this.ReportDir )
				erase ( this.ScriptTbl + '.*' )
			ENDIF
			IF this.SaveErrors and not empty( this.ErrTbl )
				use in select( this.ErrTbl )
				copy file ( this.ErrTbl + '.*' ) to ;
					forcepath( this.ErrTbl + '.*', this.ReportDir )
				erase ( this.ErrTbl + '.*' )
			ENDIF
			IF this.SaveErrors and not empty( this.aDataErrTbls )
				for I = 1 to alen( this.aDataErrTbls, 1 )
					lcFile = this.aDataErrTbls[I, 2] + '.*'
					copy file ( lcFile ) to forcepath( lcFile, this.ReportDir )
					erase ( lcFile )
			    next I
			ENDIF
			cd ( this.ReportDir )
		ENDIF

        *Create new database
        lcNewDB = this.UniqueFileName( NEWDB_NAME_LOC, 'dbc' )
        CREATE DATABASE &lcNewDB
        SET DATABASE TO &lcNewDB

        IF this.DoReport THEN
            *Create table that contains 'one time only' data and put the data in it
            lcMisc=this.CreateWzTable( MISC_NAME_LOC )
            this.PutDataInMisc( lcMisc )

            *Add analysis tables to database
            ADD TABLE ( RTRIM( this.EnumFieldsTbl )) NAME FIELD_NAME_LOC
            ADD TABLE ( RTRIM( this.EnumTablesTbl )) NAME TABLE_NAME_LOC
            IF this.ExportIndexes THEN
                ADD TABLE ( RTRIM( this.EnumIndexesTbl )) NAME INDEX_NAME_LOC
            ENDIF

            IF this.ExportRelations THEN
                ADD TABLE ( RTRIM( this.EnumRelsTbl )) NAME REL_NAME_LOC
            ENDIF
            IF !this.ViewsTbl=="" THEN
                ADD TABLE ( RTRIM( this.ViewsTbl )) NAME VIEW_NAME_LOC
            ENDIF

            *Set relations between them

* For some reason, the ALTER TABLE commands sometimes fail the first time,
* so try a second time if so.
			try
	            ALTER TABLE TABLE_NAME_LOC ALTER COLUMN TblID I PRIMARY KEY
			catch
            	ALTER TABLE TABLE_NAME_LOC ALTER COLUMN TblID I PRIMARY KEY
			endtry
            *This "tables" table needs to be open for cleanup later on
            USE
            USE TABLE_NAME_LOC ALIAS ( this.EnumTablesTbl )
            SELECT 0
			try
            	ALTER TABLE FIELD_NAME_LOC ALTER COLUMN TblID I REFERENCES TABLE_NAME_LOC TAG TblID
			catch
	            ALTER TABLE FIELD_NAME_LOC ALTER COLUMN TblID I REFERENCES TABLE_NAME_LOC TAG TblID
			endtry
            USE
            IF this.ExportIndexes
				try
                	ALTER TABLE INDEX_NAME_LOC ALTER COLUMN TblID I REFERENCES TABLE_NAME_LOC TAG TblID
				catch
                	ALTER TABLE INDEX_NAME_LOC ALTER COLUMN TblID I REFERENCES TABLE_NAME_LOC TAG TblID
				endtry
                USE
            ENDIF

        ENDIF

        IF this.DoScripts THEN
            this.DisposeTable( this.ScriptTbl, "close" )
            ADD TABLE ( RTRIM( this.ScriptTbl )) NAME SCRIPT_NAME_LOC
        ENDIF

        *Toss in error table
        IF this.SaveErrors AND !this.ErrTbl=="" THEN
            this.DisposeTable( this.ErrTbl, "close" )
            ADD TABLE ( RTRIM( this.ErrTbl )) NAME ERROR_NAME_LOC
        ENDIF

        *Add tables that contain failed data exports
        IF this.SaveErrors AND !EMPTY( this.aDataErrTbls ) THEN
            FOR I=1 TO ALEN( this.aDataErrTbls, 1 )
                lcErrTblName=ERR_TBL_PREFIX_LOC + ;
                    LEFT( this.aDataErrTbls[i, 1], MAX_NAME_LENGTH-LEN( ERR_TBL_PREFIX_LOC ))
                ADD TABLE ( RTRIM( this.aDataErrTbls[i, 2] )) NAME ( lcErrTblName )
            NEXT
        ENDIF

        *
        *Create new project
        *

        lcNewProj = this.UniqueFileName( NEWPROJ_NAME_LOC, 'pjx' )
        this.NewProjName=lcNewProj
        USE Project1.PJX
        COPY TO lcNewProj+".pjx"
        USE lcNewProj+".pjx"

        *change project path to its new directory
        lcPath=DBC()
        lcPath=STRTRAN( lcPath, SET( 'DATA' )+".DBC" )
        lcPathAndFile=lcPath+lcNewProj+".PJX" + CHR( 0 )
        lcPath=lcPath+CHR( 0 )

        REPLACE NAME WITH lcPathAndFile, ;
            HOMEDIR WITH lcPath, ;
            OBJECT WITH lcPath, ;
            Reserved1 WITH lcPathAndFile

        *Add database to project
        LOCATE FOR LOWER( TYPE )="d"
        REPLACE NAME WITH LOWER( lcNewDB )+".dbc", KEY WITH UPPER( lcNewDB )

        *Create reports and add report records to project
        IF this.DoReport
            DIMENSION aRepArray[1, 2]
            this.AddToRepArray( FIELDS_REPORT_LOC, @aRepArray )
            this.AddToRepArray( TABLES_REPORT_LOC, @aRepArray )
            IF !this.ErrTbl=="" AND this.SaveErrors THEN
                this.AddToRepArray( ERR_REPORT_LOC, @aRepArray )
            ENDIF
            IF this.ExportIndexes THEN
                this.AddToRepArray( INDEX_REPORT_LOC, @aRepArray )
            ENDIF
            IF this.ExportRelations THEN
                this.AddToRepArray( RELS_REPORT_LOC, @aRepArray )
            ENDIF
            IF this.ExportViewToRmt AND !this.ViewsTbl=="" THEN
                this.AddToRepArray( VIEWS_REPORT_LOC, @aRepArray )
            ENDIF

            FOR I=1 TO ALEN( aRepArray, 1 )

                *Create copy of each upsizing report
                aRepArray[i, 2]=this.UniqueFileName( aRepArray[i, 1], "frx" )
                SELECT 0
                USE ( aRepArray[i, 1] ) + ".frx"
                COPY TO aRepArray[i, 2] + ".frx"
                USE ( aRepArray[i, 2] ) + ".frx"
                LOCATE FOR NAME="cursor"
                *Need to fiddle with the DE
                DO WHILE FOUND()
                    REPLACE EXPR WITH STRTRAN( EXPR, "upsize1", DBC() )
                    CONTINUE
                ENDDO
                USE

                *Add to project table
                SELECT ( lcNewProj )
                APPEND BLANK
                REPLACE NAME WITH LOWER( aRepArray[i, 2] )+".frx", ;
                    KEY WITH UPPER( aRepArray[i, 2] ), ;
                    EXCLUDE WITH .F., ;
                    TYPE WITH "R"

            ENDFOR

        ENDIF
        USE

		raiseevent( This, 'CompleteProcess' )

        lcNewProj = ADDBS( FULLPATH( SYS( 5 )) ) + lcNewProj
        IF not this.lQuiet
	        _SHELL=[MODIFY PROJECT "&lcNewProj" NOWAIT]
        ENDIF
        this.KeepNewDir=.T.



    FUNCTION Integerize

        LOCAL lcEnumTablesTbl, lcEnumFieldsTbl, lcEnumIndexesTbl, lcTemp, lcTemp1

        *Add unique IDs to table names and propagate to child tables
        lcEnumTablesTbl=this.EnumTablesTbl
        lcEnumFieldsTbl=this.EnumFieldsTbl
        lcEnumIndexesTbl=this.EnumIndexesTbl

        SELECT ( lcEnumTablesTbl )
        REPLACE ALL TblID WITH RECNO()
        SCAN
            lcTemp =&lcEnumTablesTbl..TblID
            lcTemp1 = &lcEnumTablesTbl..TblName
            UPDATE ( lcEnumFieldsTbl );
                SET &lcEnumFieldsTbl..TblID=lcTemp ;
                WHERE &lcEnumFieldsTbl..TblName==lcTemp1

            IF this.ExportIndexes
                lcTemp =&lcEnumTablesTbl..TblID
                lcTemp1 =&lcEnumTablesTbl..TblName
                UPDATE ( lcEnumIndexesTbl );
                    SET &lcEnumIndexesTbl..TblID=lcTemp ;
                    WHERE &lcEnumIndexesTbl..IndexName==lcTemp1
            ENDIF

        ENDSCAN


    FUNCTION AddToRepArray
        PARAMETERS lcRepName, aRepArray
        LOCAL lnArrayLen
        *This method assumes that the passed array is 2D; the string passed is placed
        *in the first column

        IF EMPTY( aRepArray ) THEN
            aRepArray[1, 1]=lcRepName
        ELSE
            lnArrayLen=ALEN( aRepArray, 1 )
            DIMENSION aRepArray[lnArrayLen+1, 2]
            aRepArray[lnArrayLen+1, 1]=lcRepName
        ENDIF



    FUNCTION PutDataInMisc
        PARAMETERS lcMisc

        INSERT INTO &lcMisc ;
            ( SvType, ;
            DataSourceName, ;
            UserConnection, ;
            ViewConnection, ;
            DeviceDBName, ;
            DeviceDBPName, ;
            DeviceDBSize, ;
            DeviceDBNumber, ;
            DeviceLogName, ;
            DeviceLogPName, ;
            DeviceLogSize, ;
            DeviceLogNumber, ;
            ServerDBName, ;
            ServerDBSize, ;
            ServerLogSize, ;
            SourceDB, ;
            ExportIndexes, ;
            ExportValidation, ;
            ExportRelations, ;
            ExportStructureOnly, ;
            ExportDefaults, ;
            ExportTimeStamp, ;
            ExportTableToView, ;
            ExportViewToRmt, ;
            ExportDRI, ;
            ExportSavePwd, ;
            DoUpsize, ;
            DoScripts, ;
            DoReport ) ;
            VALUES ;
            ( this.ServerType, ;
            this.DataSourceName, ;
            this.UserConnection, ;
            this.ViewConnection, ;
            this.DeviceDBName, ;
            this.DeviceDBPName, ;
            this.DeviceDBSize, ;
            this.DeviceDBNumber, ;
            this.DeviceLogName, ;
            this.DeviceLogPName, ;
            this.DeviceLogSize, ;
            this.DeviceLogNumber, ;
            this.ServerDBName, ;
            this.ServerDBSize, ;
            this.ServerLogSize, ;
            this.SourceDB, ;
            this.ExportIndexes, ;
            this.ExportValidation, ;
            this.ExportRelations, ;
            this.ExportStructureOnly, ;
            this.ExportDefaults, ;
            this.ExportTimeStamp, ;
            this.ExportTableToView, ;
            this.ExportViewToRmt, ;
            this.ExportDRI, ;
            this.ExportSavePwd, ;
            this.DoUpsize, ;
            this.DoScripts, ;
            this.DoReport )

        USE



    FUNCTION ReorderTable
        LOCAL lcNewName, lcAlias

        *Copy to new table sorted by table name ( can't index on 128 character field
        *using general code page )

        *		lcNewName=SUBSTR( SYS( 2015 ), 3, 10 )
        lcNewName="_"+SUBSTR( SYS( 2015 ), 4, 10 )
        RENAME ( this.EnumTablesTbl )+".dbf" TO ( lcNewName )+".dbf"
        RENAME ( this.EnumTablesTbl )+".fpt" TO ( lcNewName )+".fpt"
        *{ JEI RKR 18.11.2005 Add
        IF FILE( this.EnumTablesTbl+".cdx" )
        	RENAME ( this.EnumTablesTbl )+".cdx" TO ( lcNewName )+".cdx"
        ENDIF
        *{ JEI RKR 18.11.2005 Add
        SELECT 0
        USE ( lcNewName )
        lcAlias=ALIAS()

        SELECT * FROM ( lcNewName );
            ORDER BY TblName;
            INTO TABLE ( this.EnumTablesTbl )

        USE IN ( lcAlias )

        DELETE FILE ( lcNewName )+".dbf"
        DELETE FILE ( lcNewName )+".fpt"
        DELETE FILE ( lcNewName )+".cdx"



    FUNCTION PurgeTable
        PARAMETERS lcTableName, lcCondition
        *This removes objects which were analyzed but not selected for upsizing

        SELECT ( lcTableName )
        SET FILTER TO
        DELETE FOR &lcCondition
        PACK
        USE



    FUNCTION PurgeTable2
        LOCAL lcTableName

        *Gets rid of field info related to tables not selected for upsizing
        *( This info gets added if the user selects a table, changes pages, and
        *then deselects the table )

        SELECT ( this.EnumTablesTbl )
        SCAN FOR EXPORT=.F.
            lcTableName=RTRIM( TblName )
            SELECT ( this.EnumFieldsTbl )
            DELETE FOR RTRIM( TblName )==lcTableName
            SELECT ( this.EnumTablesTbl )
        ENDSCAN

        SELECT ( this.EnumFieldsTbl )
        PACK
        USE



    FUNCTION CreateScript
        LOCAL lcEnumRelsTbl, lcEnumTablesTbl, lcCRLF, lcComment, lcEnumIndexesTbl, ;
            lcEnumFieldsTbl, lnTableCount, lcEnumClustersTbl

        *Device and database sql ( if any ) are already in the script memo field by now
        IF !this.DoScripts THEN
            RETURN
        ENDIF

* If we have an extension object and it has a CreateScript method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'CreateScript', 5 ) and ;
			not this.oExtension.CreateScript( This )
			return
		ENDIF

        lcCRLF = CHR( 13 )
        lcEnumClustersTbl = this.EnumClustersTbl
        lcEnumTablesTbl = this.EnumTablesTbl
        lcEnumIndexesTbl = this.EnumIndexesTbl
        lcEnumRelsTbl = this.EnumRelsTbl
        lcEnumFieldsTbl = this.EnumFieldsTbl

        SELECT COUNT( * ) FROM ( lcEnumTablesTbl ) WHERE EXPORT=.T. INTO ARRAY aTableCount
        this.InitTherm( BUILDING_SCRIPT_LOC, aTableCount, 0 )
        lnTableCount = 0
        this.UpDateTherm( lnTableCount, "" )

        IF this.ServerType = "Oracle"
            * Grab cluster SQL ( if any )
            IF !EMPTY( lcEnumClustersTbl )
                SELECT ( lcEnumClustersTbl )
                SCAN FOR !EMPTY( ClustName )

                    * Get cluster SQL
                    lcClustName = RTRIM( ClustName )
                    lcSQL = this.BuildComment( CLUST_COMMENT_LOC, lcClustName )
                    lcSQL = lcSQL + ClusterSQL
                    this.StoreSQL( lcSQL, "" )
                    lcSQL = ""

                    *Grab cluster index SQL
                    SELECT ( lcEnumIndexesTbl )
                    LOCATE FOR RTRIM( RmtTable ) == lcClustName
                    IF FOUND()
                        lcSQL = this.BuildComment( CLUST_INDEX_LOC, RmtName )
                        lcSQL = lcSQL + IndexSQL
                        this.StoreSQL( lcSQL, "" )
                    ENDIF
                    lcSQL = ""

                    * Grab SQL of tables ( and their triggers ) in cluster
                    SELECT ( lcEnumTablesTbl )
                    SCAN FOR RTRIM( ClustName ) == lcClustName
                        lcTableName = RTRIM( TblName )
                        lcRmtTblName = RTRIM( RmtTblName )
                        this.OracleScript( lcTableName, lcRmtTblName )
                        lnTableCount = lnTableCount + 1
                        this.UpDateTherm( lnTableCount )
                    ENDSCAN
                    SELECT ( lcEnumClustersTbl )
                ENDSCAN
            ENDIF
            * Deal with tables not in clusters
            * Grab SQL of tables ( and their triggers ) in cluster
            SELECT ( lcEnumTablesTbl )
            SCAN FOR EMPTY( ClustName ) AND EXPORT=.T.
                lcTableName = RTRIM( TblName )
                lcRmtTblName = RTRIM( RmtTblName )
                this.OracleScript( lcTableName, lcRmtTblName )
                lnTableCount = lnTableCount + 1
                this.UpDateTherm( lnTableCount )
            ENDSCAN
        ELSE
            * Deal with SQL Server
            SELECT ( lcEnumTablesTbl )
            IF this.ZDUsed THEN
                lcSQL = ZD_DESC_LOC + lcCRLF
                lcSQL=lcSQL + "CREATE DEFAULT " + ZERO_DEFAULT_NAME + " AS 0" + lcCRLF
                this.StoreSQL( lcSQL, "" )
            ENDIF

            SCAN FOR EXPORT=.T.

                *Grab table SQL
                lcTableName=RTRIM( TblName )
                lcRmtTblName=RTRIM( RmtTblName )
                lcSQL = this.BuildComment( TABLE_COMMENT_LOC, lcRmtTblName ) + TableSQL
                this.StoreSQL( lcSQL, "" )

                *Triggers
                lcSQL=""
                IF !EMPTY( InsertRI ) THEN
                    lcSQL = InsertRI
                ENDIF
                IF !EMPTY( UpdateRI ) THEN
                    lcSQL = IIF( EMPTY( lcSQL ), UpdateRI, lcSQL + UpdateRI + lcCRLF )
                ENDIF
                IF !EMPTY( DeleteRI ) THEN
                    lcSQL = IIF( EMPTY( lcSQL ), DeleteRI, lcSQL + DeleteRI + lcCRLF )
                ENDIF
                IF !lcSQL=="" THEN
                    this.StoreSQL( lcSQL, TRIGGER_COMMENT_LOC )
                ENDIF

                *Grab index SQL
                IF !EMPTY( lcEnumIndexesTbl )
                    SELECT ( lcEnumIndexesTbl )
                    lcSQL = lcCRLF + INDEX_COMMENT_LOC + lcCRLF
                    SCAN FOR RTRIM( IndexName )==lcTableName AND DontCreate=.F.
                        lcSQL=lcSQL+ IndexSQL
                        this.StoreSQL( lcSQL, "" )
                        lcSQL=""
                    ENDSCAN
                ENDIF

                *Grab default SQL
                SELECT ( lcEnumFieldsTbl )
                lcSQL = lcCRLF + DEFAULT_COMMENT_LOC + lcCRLF
                SCAN FOR RTRIM( TblName )==lcTableName AND !EMPTY( RmtDefault )
                    IF RmtDefault<>"0" THEN
                        lcSQL = lcSQL + RmtDefault + lcCRLF
                    ENDIF
                    lcSQL = lcSQL + "sp_bindefault " + RTRIM( RDName ) + ", '" + RTRIM( TblName ) + "." + RTRIM( FldName ) +"'" + lcCRLF
                    this.StoreSQL( lcSQL, "" )
                    lcSQL=""

                ENDSCAN

                *Stored procedures
                lcSQL=SPROC_COMMENT_LOC
                SCAN FOR RTRIM( TblName )==lcTableName AND !EMPTY( RmtRule )
                    lcSQL=lcSQL + RmtDefault
                    this.StoreSQL( lcSQL, "" )
                    lcSQL=""
                ENDSCAN

                SELECT ( lcEnumTablesTbl )
                lnTableCount=lnTableCount+1
                this.UpDateTherm( lnTableCount )

            ENDSCAN

        ENDIF

		raiseevent( This, 'CompleteProcess' )



    FUNCTION OracleScript
        PARAMETERS lcTableName, lcRmtTblName
        LOCAL lcEnumTablesTbl, lcEnumIndexsTbl, lcEnumFieldsTbl, lcCRLF

        *
        *Called by CreateScript; puts together all the sql for a table
        *including the table, triggers, indexes and defaults
        *

        lcEnumTablesTbl = this.EnumTablesTbl
        lcEnumIndexesTbl = this.EnumIndexesTbl
        lcEnumFieldsTbl = this.EnumFieldsTbl
        lcCRLF = CHR( 13 )

        * Grab SQL of tables ( and their triggers )
        lcSQL = this.BuildComment( TABLE_COMMENT_LOC, lcRmtTblName ) + lcCRLF + TableSQL
        this.StoreSQL( lcSQL, "" )

        * Triggers
        lcSQL = ""
        IF !EMPTY( InsertRI ) THEN
            lcSQL = InsertRI + lcCRLF
        ENDIF
        IF !EMPTY( UpdateRI ) THEN
            lcSQL = lcSQL + UpdateRI + lcCRLF
        ENDIF
        IF !EMPTY( DeleteRI ) THEN
            lcSQL = lcSQL + DeleteRI + lcCRLF
        ENDIF
        IF !EMPTY( lcSQL )
            this.StoreSQL( lcSQL, TRIGGER_COMMENT_LOC )
        ENDIF

        * Grab index sql
        SELECT ( lcEnumIndexesTbl )
        lcSQL = lcCRLF + INDEX_COMMENT_LOC + lcCRLF + lcCRLF
        SCAN FOR RTRIM( IndexName ) == lcTableName AND !EMPTY( IndexSQL )
            lcSQL = lcSQL + IndexSQL
            this.StoreSQL( lcSQL, "" )
            lcSQL = ""
        ENDSCAN

        *Grab default sql
        SELECT ( lcEnumFieldsTbl )
        lcSQL = lcCRLF + DEFAULT_COMMENT_LOC + lcCRLF + lcCRLF
        SCAN FOR RTRIM( TblName ) == lcTableName AND !EMPTY( RmtDefault )
            lcSQL = lcSQL + RmtDefault
            this.StoreSQL( lcSQL, "" )
            lcSQL = ""
        ENDSCAN

        SELECT ( lcEnumTablesTbl )



    FUNCTION UniqueFileName
        PARAMETERS lcFileName, lcExtension
        LOCAL I, lcNewName

        lcNewName=lcFileName
        I=1
        DO WHILE FILE( lcNewName + "." + lcExtension )
            lcNewName=LEFT( lcFileName, MAX_DOSNAME_LEN-LEN( LTRIM( STR( I )) )) + LTRIM( STR( I ))
            I=I+1
        ENDDO
        RETURN lcNewName



    FUNCTION JustStem2

        * Return just the stem name from "filname"
        * Unlike JustStem, this returns file name in same case it came in as

        LPARAMETERS m.filname
        IF RAT( '\', m.filname ) > 0
            m.filname = SUBSTR( m.filname, RAT( '\', m.filname )+1, 255 )
        ENDIF
        IF RAT( ':', m.filname ) > 0
            m.filname = SUBSTR( m.filname, RAT( ':', m.filname )+1, 255 )
        ENDIF
        IF AT( '.', m.filname ) > 0
            m.filname = SUBSTR( m.filname, 1, AT( '.', m.filname )-1 )
        ENDIF
        RETURN ALLTRIM( m.filname )



    FUNCTION RemotizeName
        PARAMETERS lcLocalName
        LOCAL lcResult, lnLength, I, lcChar, lnOldArea, lnAsc, lcExact, lcServerConstraint

        *all expressions and objects everywhere in the Upsizing Wizard are
        *lower cased, otherwise STRTRAN transformations won't work reliably
        lcExact=SET( 'EXACT' )
        SET EXACT ON
        lcResult = LOWER( ALLTRIM( lcLocalName ))
        lnOldArea=SELECT()

        * Check keyword table
        IF !USED( "Keywords" )
            SELECT 0
            USE Keywords
            SET ORDER TO Keyword
        ELSE
            SELECT Keywords
        ENDIF
        IF RTRIM( this.ServerType )=="SQL Server95" THEN
            lcServerConstraint="SQL Server"
        ELSE
            lcServerConstraint=RTRIM( this.ServerType )
        ENDIF
        SET FILTER TO ServerType=lcServerConstraint

        SEEK lcResult

        IF FOUND() THEN
            lcResult = lcResult + "_"

        ELSE

            *if it starts with a number, stick a "_" in front of it
            IF LEFT( lcLocalName, 1 ) >= "0" AND LEFT( lcLocalName, 1 ) <= "9" THEN
                lcResult= "_" + lcResult
            ENDIF

            lnLength = LEN( lcResult )

            *ISALPHA() will return true but SQL Server will reject when...
            *Codepage 1252 ( US ): 156, 207
            *Codepage 1250 ( E.Eur. ): 156, 190, 207
            *Codepage ???? ( Russia ): 220
            *So these characters are always ( on all code pages ) turned to underscores
            *Skip for DBCS
            IF lnLength = LENC( lcResult )
                FOR I = 1 TO lnLength
                    lcChar = SUBSTR( lcResult, I, 1 )
					lnAsc = ASC( lcChar )
						IF ( not isalpha( lcChar ) and ;
							not between( lcChar, '0', '9' ) and ;
							lcChar <> ' ' ) or ;
                            inlist( lnAsc, 156, 190, 207, 220 )
                        lcResult=STUFF( lcResult, I, 1, "_" )
                    ENDIF
                NEXT I
            ENDIF
        ENDIF

        SET EXACT &lcExact
        SELECT ( lnOldArea )
        RETURN lcResult



    FUNCTION UniqueTableName
        PARAMETERS lcStem
        LOCAL lcTest, lcResult, I, lnLength

        lcStem=ALLTRIM( lcStem )
        lcResult=lcStem
        lcTest=lcResult + ".dbf"
        lnLength=LEN( lcStem )
        FOR I=1 TO 10^lnLength-1
            IF FILE( lcTest ) THEN
                lcResult=LEFT( lcStem, lnLength-LEN( ALLTRIM( STR( I )) )) + ALLTRIM( STR( I ))
                lcTest=lcResult + ".dbf"
            ELSE
                EXIT
            ENDIF
        NEXT
        RETURN lcResult



    * 10/30/02 JVF Replaced this function with the new one below. Left intact, but retired, b/c of
    * its legacy nature.
    FUNCTION UniqueCursorName_Retired
        PARAMETERS lcStem
        LOCAL lcResult, I, lnLength

        lcStem=ALLTRIM( lcStem )
        lcResult=lcStem
        lnLength=LEN( lcStem )
        FOR I=1 TO lnLength-1
            IF USED( lcResult ) THEN
                lcResult=LEFT( lcStem, lnLength-LEN( ALLTRIM( STR( I )) )) + ALLTRIM( STR( I ))
            ELSE
                EXIT
            ENDIF
        NEXT
        RETURN lcResult

        * 10/30/02 JVF Issue 16455 Alias in use error with similar tables ending in numeric chars.
        * I see the fatal flaw of original code:
        * For I=1 TO lnLength-1
        * It only loops 6 times b/c the stem length was 7. I guess the USW never ran into situation
        * where their were more cursors open than length of the stem -1.

        * I agree we can loop 999 times and don't have to worry about names longer than 8, so we do not
        * have to keep shortnening the stem. But the following approach is a little better than the
        * recommended b/c we maintain the stem instead of just appending to it.

    FUNCTION UniqueCursorName
        PARAMETERS lcStem

        LOCAL lcResult, I, lnLength

        lcStem=ALLTRIM( lcStem )
        lcResult=lcStem
        lnLength=LEN( lcStem )
        FOR I=1 TO 999
            IF USED( lcResult ) THEN
                lcResult=lcStem + PADL( I, 3, "0" )
            ELSE
                EXIT
            ENDIF
        NEXT
        RETURN lcResult



    FUNCTION CreateTypeArrays
        *Creates an array for each FoxPro datatype; each array of possible remote datatypes
        *has the same name as the local FoxPro datatype
        PRIVATE aArrays
        LOCAL lcServerConstraint, I

*** DH 09/05/2012: use only SQL Server rather than variants of it
***        IF RTRIM( this.ServerType )=="SQL Server95" THEN
		IF left( this.ServerType, 10 ) = 'SQL Server'
            lcServerConstraint="SQL Server"
        ELSE
***            IF this.ServerType="SQL Server" THEN
***                lcServerConstraint="SQL Server4x"
***            ELSE
                lcServerConstraint=RTRIM( this.ServerType )
***            ENDIF
        ENDIF

        *Find all the local types
        SELECT LocalType, "this." + LocalType FROM TypeMap ;
            WHERE DEFAULT=.T. AND SERVER=lcServerConstraint ;
            INTO ARRAY aArrays

        FOR I=1 TO ALEN( aArrays, 1 )
            SELECT RemoteType FROM TypeMap WHERE LocalType=aArrays[i, 1] AND SERVER=lcServerConstraint ;
                INTO ARRAY &aArrays[i, 2]
        NEXT



    FUNCTION ValidName
        PARAMETERS lcName
        LOCAL lcNewName

        lcNewName=this.RemotizeName( lcName )
        IF LOWER( lcNewName )<>LOWER( lcName ) THEN
            *display error message
			IF not this.lQuiet
	            MESSAGEBOX( INVALID_NAME_LOC, ICON_EXCLAMATION, TITLE_TEXT_LOC )
			ENDIF
            lcName=lcNewName
            RETURN .F.
        ENDIF
        RETURN .T.



    FUNCTION NameObject
        PARAMETERS lcRmtTableName, lcFldName, lcPrefix, lnMaxLength
        LOCAL lnTblNameLength, lnFldNameLength, lnCharsLeft

        lnTblNameLength=LEN( lcRmtTableName )
        lnFldNameLength=LEN( lcFldName )
        lnCharsLeft=( lnMaxLength )-LEN( lcPrefix )-LEN( SEP_CHARACTER )

        *If all the components of the string are too big, clip
        *the table and/or field names

        IF lnCharsLeft<lnTblNameLength+lnFldNameLength THEN
            DO CASE
                    *If each name is bigger than half of what's left, clip them both
                CASE lnTblNameLength>( lnCharsLeft/2 ) AND lnFldNameLength>( lnCharsLeft/2 )
                    lcRmtTableName=LEFT( lcRmtTableName, ( lnCharsLeft/2 ))
                    lcFldName=LEFT( lcFldName, ( lnCharsLeft/2 ))

                    *If the field name is super long, clip it
                CASE lnTblNameLength<=( lnCharsLeft/2 ) AND lnFldNameLength>( lnCharsLeft/2 )
                    lcFldName=LEFT( lcFldName, ( lnCharsLeft-lnTblNameLength ))

                    *If the table name is super long, clip it
                CASE lnTblNameLength>( lnCharsLeft/2 ) AND lnFldNameLength<=( lnCharsLeft/2 )
                    lcTmpTblName=LEFT( lcRmtTableName, ( lnCharsLeft-lnFldNameLength ))

            ENDCASE
        ENDIF

        lcSprocName=lcPrefix+lcRmtTableName + SEP_CHARACTER + lcFldName

        RETURN lcSprocName



    FUNCTION MaybeDrop
        PARAMETERS lcObjectName, lcObjectType
        LOCAL llObjectExists, lcSQL, dummy, lcSQT

        *This is called in several places by this.DefaultsAndRules
        *It will drop a sproc or default if it already exists

        *Check to see if the object already exists
        lcSQT=CHR( 39 )
        lcSQL="select uid from sysobjects where name =" + lcSQT + lcObjectName + lcSQT
        dummy="x"
        llObjectExists=this.SingleValueSPT( lcSQL, dummy, "uid" )

        IF llObjectExists THEN
            lcSQL="drop " + lcObjectType + " " + lcObjectName
            lnRetVal=this.ExecuteTempSPT( lcSQL )
            RETURN lnRetVal
        ELSE
            RETURN .T.
        ENDIF



    FUNCTION ExtractFieldNames
        PARAMETERS lcExpression, lcTableName, lnKeyCount, aFieldNames
        LOCAL ii, lcReturnExpression, lcEnum_Fields, lcFieldName, lnRow

        *
        *Takes a FoxPro expression and returns comma separated list of
        *remotized version of all the field names that were in the expression
        *
        *Called by AnalyzeIndexes and BuildRICode
        *

        *Build the array of field names if it wasn't passed
        IF EMPTY( aFieldNames ) THEN
            DIMENSION aFieldNames[1]
            lcEnum_Fields=RTRIM( this.EnumFieldsTbl )
            *Be sure they come back in order of the longest field names first
            *or the shorter ones which are substrings of longer ones ( if any )
            *will mess things up
            SELECT FldName, RmtFldname, 1/LEN( RTRIM( RmtFldname )) AS foo ;
                FROM &lcEnum_Fields ;
                WHERE &lcEnum_Fields..TblName=lcTableName ;
                ORDER BY foo ;
                INTO ARRAY aFieldNames
        ENDIF

        lnKeyCount=0
        lcReturnExpression=""

        *!* jvf: 08/16/99
        * Replaced code that caused indexes to be created out of the intended field sequence.
        * The subsequent code corrects this issue.
        DIMENSION laExp[1]
        this.StringToArray( lcExpression, @laExp, "+" )
        FOR lnRow = 1 TO ALEN( laExp, 1 )
            lcFieldName=this.StripFunction( laExp[lnRow] )
            IF lcReturnExpression=="" THEN
                lcReturnExpression="[" + lcFieldName + "]"
            ELSE
                lcReturnExpression=lcReturnExpression+", ["+lcFieldName+"]"
            ENDIF

            *Keep track of how many fields are in the index expression
            lnKeyCount=lnKeyCount+1
        ENDFOR
        RETURN lcReturnExpression



    FUNCTION StoreError
        PARAMETERS lnError, lcErrMsg, lcSQL, lcWizErrMsg, lcObjName, lcObjType
        LOCAL lcErrTbl, lnOldArea

        *Stores errors for report

        lnOldArea=SELECT()

        IF this.ErrTbl=="" THEN
            this.ErrTbl=this.CreateWzTable( "Errors" )
        ENDIF
        lcErrTbl=this.ErrTbl
        IF EMPTY( lcWizErrMsg ) THEN
            lcWizErrMsg=""
        ENDIF
        INSERT INTO &lcErrTbl ( ErrNumber, ErrMsg, WizErr, FailedSQL, ObjName, ObjType ) ;
            VALUES ( lnError, lcErrMsg, lcWizErrMsg, lcSQL, lcObjName, lcObjType )
        SELECT ( lnOldArea )



    FUNCTION StoreSQL
        PARAMETERS lcSQL, lcComment
        LOCAL lcCRLF, lnOldArea, lcScriptTbl

        *Get out of here if user doesn't want a script
        IF !this.DoScripts THEN
            RETURN
        ENDIF

        lcCRLF = CHR( 13 )
        lnOldArea=SELECT()

        IF RTRIM( this.ScriptTbl ) == "" THEN
            lcScriptTbl = this.CreateWzTable( "Script" )
            this.ScriptTbl = lcScriptTbl
        ELSE
            lcScriptTbl = RTRIM( this.ScriptTbl )
        ENDIF

        SELECT ( lcScriptTbl )

        * There should only be one record in this table
        IF RECCOUNT()=0 THEN
            APPEND BLANK
        ENDIF

        *Add some carriage returns/linefeeds and then stick everything together
        IF !EMPTY( lcComment ) THEN
            lcSQL = lcCRLF + lcComment + lcCRLF + lcSQL
        ENDIF
        REPLACE &lcScriptTbl..ScriptSQL WITH lcSQL + lcCRLF ADDITIVE

        * Prevent bulk build-up of memo
        PACK MEMO

        SELECT ( lnOldArea )



    FUNCTION SetConnProps
        *Sets connection properties to where the upsizing wizard wants them
        =SQLSETPROP( this.MasterConnHand, "Asynchronous", .F. )
        =SQLSETPROP( this.MasterConnHand, "Batchmode", .T. )
        =SQLSETPROP( this.MasterConnHand, "ConnectTimeOut", 45 )
        =SQLSETPROP( this.MasterConnHand, "DispWarnings", .F. )
        *QueryTimeOut is set to 600 when creating devices and databases
        =SQLSETPROP( this.MasterConnHand, "QueryTimeOut", 45 )
        =SQLSETPROP( this.MasterConnHand, "Transactions", 1 )
        =SQLSETPROP( this.MasterConnHand, "WaitTime", 100 )
        *Never timeout if idle
        =SQLSETPROP( this.MasterConnHand, "IdleTimeOut", 0 )
        *Default wait time of 100 milliseconds


    FUNCTION CreateTS
        PARAMETERS lcTSName, lcTSFName, lcTSFSize
        LOCAL lcSQL, lcSQL1, lcMsg, lnErr, lcErrMsg

        * create new tablespace on Oracle server and an associated data file
        * allocates unlimited space quota for the user on the new tablespace
        lcSQL = "CREATE TABLESPACE " + lcTSName + " DATAFILE '" + lcTSFName + "' SIZE " + ALLTRIM( STR( lcTSFSize )) + " K"
        lcSQL1 = "ALTER USER " + this.UserName + " QUOTA UNLIMITED ON " + lcTSName

        *Execute if appropriate
        IF this.DoUpsize THEN
            lcMsg=STRTRAN( CREATING_TABLESPACE_LOC, '|1', RTRIM( lcTSName ))
            this.InitTherm( lcMsg, 0, 0 )
            this.UpDateTherm( 0, TAKES_AWHILE_LOC )
            =SQLSETPROP( this.MasterConnHand, "QueryTimeOut", 600 )
            this.MyError=0

            * create tablespace
            IF !this.ExecuteTempSPT( lcSQL, @lnErr, @lcErrMsg ) THEN
                IF lnErr = 01543 THEN
                    *User doesn't have CREATE TABLESPACE permissions
                    lcMsg=STRTRAN( NO_CREATETS_PERM_LOC, '|1', RTRIM( this.DataSourceName ))
                ELSE
                    *Something else went wrong
                    lcMsg=STRTRAN( CREATE_TS_FAILED_LOC, '|1', RTRIM( lcTSName ))
                ENDIF
				IF this.lQuiet
					.HadError     = .T.
					.ErrorMessage = lcMsg
				ELSE
                	MESSAGEBOX( lcMsg, ICON_EXCLAMATION, TITLE_TEXT_LOC )
				ENDIF
                this.Die
            ENDIF

            * allocate unlimited quota on tablespace
            IF !this.ExecuteTempSPT( lcSQL1, @lnErr, @lcErrMsg ) THEN
                IF lnErr = 01543 THEN
                    *User doesn't have CREATE TABLESPACE permissions
                    lcMsg=STRTRAN( NO_CREATETS_PERM_LOC, '|1', RTRIM( this.DataSourceName ))
                ELSE
                    *Something else went wrong
                    lcMsg=STRTRAN( CREATE_TS_FAILED_LOC, '|1', RTRIM( lcTSName ))
                ENDIF
				IF this.lQuiet
					.HadError     = .T.
					.ErrorMessage = lcMsg
				ELSE
                	MESSAGEBOX( lcMsg, ICON_EXCLAMATION, TITLE_TEXT_LOC )
				ENDIF
                this.Die
            ENDIF

            =SQLSETPROP( this.MasterConnHand, "QueryTimeOut", 30 )
        ENDIF

        *Stash sql for script
        this.StoreSQL( lcSQL, CREATE_DBSQL_LOC )
        this.StoreSQL( lcSQL1, CREATE_DBSQL_LOC )
		raiseevent( This, 'CompleteProcess' )


    * create new datafile in existing tablespace
    FUNCTION CreateDataFile
        PARAMETERS lcTSName, lcTSFName, lcTSFSize
        lcSQL = "ALTER TABLESPACE " + lcTSName + " ADD DATAFILE '" + lcTSFName + "' SIZE " + ALLTRIM( STR( lcTSFSize )) + " K"
		llRetVal = this.ExecuteTempSPT( lcSQL )


    FUNCTION CreateWzTable
        PARAMETERS lcPassed
        LOCAL lcTableName, lcOldDir

        *
        *All tables used internally ( except device table ) get created here, indexes where possible
        *

        SELECT 0
        IF this.NewDir == "" THEN
            this.CreateNewDir
        ENDIF
        lcTableName = this.UniqueTableName( lcPassed )

        DO CASE
            CASE lcPassed="Tables"
                CREATE TABLE &lcTableName FREE;
                    ( TblName C ( 128 ) NOT NULL, ;
                    TblID I, ;
                    CursName C ( 128 ) NOT NULL, ;
                    TblPath M, ;
                    RmtTblName C ( 30 ) NOT NULL, ;
                    NewTblName C ( 128 ), ;
                    Upsizable L, ;
                    PreOpened L, ;
                    EXPORT L, ;
                    Exported L, ;
                    DataErrs N( 9 ), ;
                    DataErrMsg M, ;
                    ErrTable C ( 128 ), ;
                    FldsAnald L, ;
                    CDXAnald L, ;
                    ClustName C( 30 ), ;
                    TableSQL M, ;
                    TStampAdd L, ;
                    IdentAdd L, ;
                    LocalRule M , ;
                    RmtRule M, ;
                    RRuleName C ( 30 ), ;
                    RuleExport L, ;
                    RuleError M, ;
                    RuleErrNo N ( 5 ), ;
                    ItrigName C ( 30 ), ;
                    InsertRI M, ;
                    InsertX L, ;
                    DtrigName C ( 30 ), ;
                    DeleteRI M, ;
                    DeleteX L, ;
                    UtrigName C ( 30 ), ;
                    UpdateRI M, ;
                    UpdateX L, ;
                    RIError M, ;
                    RIErrNo N ( 5 ), ;
                    FKeyCrea L, ;
                    PKeyCrea L, ;
                    PkeyExpr M, ;
                    PKTagName C( 10 ), ;
                    TblError M, ;
                    TblErrNo N ( 5 ), ;
                    RmtView C( 128 ), ;
                    ViewErr M, ;
                    Type C( 1 ), ;&& Add JEI - RKR - 24.03.2005
                    SQLStr M ) && Add JEI - RKR - 24.03.2005

                *DataSent indicates if data was successfully appended to the new table
                *Exported indicates if the table was successfully created

                *"RmtView" contains the name of the "SELECT *" view created if the table was part of a view
                *only some of whose tables were upsized; this field maybe ""

                *RRuleName is the name of remote rule

                *"NewTblName" is name of table after renaming if a remote view was created based on the table
                *Index actually created elsewhere for speed reasons
                *INDEX ON TblName TAG TblName

                *Insert, Delete, and Update contain trigger code of the same name
                *InsertX, DeleteX, and UpdateX show whether the triggers were created successfully

                *FKeyCrea and PKeyCrea are used only in the Oracle or SQL 95 case to indicate
                *whether ALTER TABLE statements to add RI succeeded or not

                * JVF 11/02/02 Added column AutoInNext I, AutoInStep I, to account for VFP 8.0 autoinc attrib.
*** DH 12/06/2012: changed size of RmtFldName from 30 to 128
*** DH 09/09/2013: changed size of RmtType from 13 to 20
            CASE lcPassed="Fields"
                CREATE TABLE &lcTableName FREE;
                    ( TblName C ( 128 ) NOT NULL, ;
                    FldName C ( 128 ) NOT NULL, ;
                    DATATYPE C ( 1 ) NOT NULL, ;
                    ComboType C ( 20 ) NOT NULL, ;
                    FullType C ( 10 ) NOT NULL, ;
                    LENGTH N ( 3 ) NOT NULL, ;
                    PRECISION N ( 3 ) NOT NULL, ;
                    NOCPTRANS L, ;
                    lclnull L, ;
                    RmtFldname C ( 128 ), ;
                    RmtType C ( 20 ), ;
                    RmtLength N ( 4 ), ;
                    RmtPrec N ( 3 ), ;
                    RmtNull L, ;
                    LocalRule M, ;
                    RmtRule M, ;
                    RRuleName C ( 30 ), ;
                    RuleExport L, ;
                    RuleError M, ;
                    RuleErrNo N ( 5 ), ;
                    DEFAULT M, ;
                    RmtDefault M, ;
                    RDName C ( 30 ), ;
                    DefaExport L, ;
                    DefaBound L, ;
                    DefaError M, ;
                    DefaErrNo N ( 5 ), ;
                    InCluster L, ;
                    ClustOrder I, ;
                    TblID I, ;
                    AutoInNext I, ;
                    AutoInStep I )



                *Index actually created elsewhere for speed reasons
                *INDEX ON TblName TAG TblName
                *INDEX ON FldName TAG FldName

            CASE lcPassed="Indexes"
                CREATE TABLE &lcTableName FREE;
                    ( TblID I, ;
                    IndexName C ( 128 ) NOT NULL, ;
                    TagName C ( 10 ), ;
                    LclExpr M, ;
                    LclIdxType C ( 12 ), ;
                    RmtTable C ( 30 ), ;
                    RmtName C ( 10 ), ;
                    RmtExpr M, ;
                    RmtType C ( 20 ), ;
                    Clustered L, ;
                    Exported L, ;
                    TblUpszd L, ;
                    DontCreate L, ;
                    IdxError M, ;
                    IdxErrNo N ( 5 ), ;
                    IndexSQL M )

                *Value in IndexName is exactly the same as local tablename
                *INDEX ON RTRIM( IndexName )+TagName TAG TblAndTag ( performed elsewhere )

            CASE lcPassed="Views"
                CREATE TABLE &lcTableName FREE;
                    ( ViewName C ( 128 ) NOT NULL, ;
                    NewName C ( 128 ), ;
                    ViewSQL M, ;
                    RmtViewSQL M, ;
                    TblsUpszd M, ;
                    NotUpszd  M, ;
                    CONNECTION C ( 128 ), ;
                    Remotized L, ;
                    ViewErr M, ;
                    ViewErrNo N( 5 ) NULL )

            CASE lcPassed="Script"
                CREATE TABLE &lcTableName FREE;
                    ( ScriptSQL M )

            CASE lcPassed="Errors"
                CREATE TABLE &lcTableName FREE;
                    ( ObjType C( 30 ), ;
                    ObjName M, ;
                    ErrNumber N( 5 ) NULL, ;
                    ErrMsg M, ;
                    WizErr M, ;
                    FailedSQL M )

            CASE lcPassed=MISC_NAME_LOC
                *This table gets created in the analysis database
                CREATE TABLE &lcTableName NAME MISC_NAME_LOC;
                    ( SvType C ( 20 ), ;
                    DataSourceName M, ;
                    UserConnection C ( 128 ), ;
                    ViewConnection C ( 128 ), ;
                    DeviceDBName C ( 30 ), ;
                    DeviceDBPName C ( 12 ), ;
                    DeviceDBSize N ( 6 ), ;
                    DeviceDBNumber N ( 3 ), ;
                    DeviceLogName C ( 30 ), ;
                    DeviceLogPName C ( 12 ), ;
                    DeviceLogSize N ( 6 ), ;
                    DeviceLogNumber N ( 3 ), ;
                    ServerDBName C ( 30 ), ;
                    ServerDBSize N ( 6 ), ;
                    ServerLogSize N ( 6 ), ;
                    SourceDB M , ;
                    ExportIndexes L, ;
                    ExportValidation L, ;
                    ExportRelations L, ;
                    ExportStructureOnly L, ;
                    ExportDefaults L, ;
                    ExportTimeStamp L, ;
                    ExportTableToView L, ;
                    ExportViewToRmt L, ;
                    ExportDRI L, ;
                    ExportSavePwd L, ;
                    DoUpsize L, ;
                    DoScripts L, ;
                    DoReport L )

            CASE lcPassed="Relation"
                CREATE TABLE &lcTableName FREE;
                    ( DD_CHIEXPR M, ;
                    DD_CHILD C( 128 ), ;
                    Dd_RmtChi C( 128 ), ;
                    dd_delete C( 1 ), ;
                    dd_insert C( 1 ), ;
                    DD_PARENT C( 128 ), ;
                    Dd_RmtPar C( 128 ), ;
                    DD_PAREXPR M, ;
                    dd_update C( 1 ), ;
                    ClustName C( 30 ), ;
                    ClustType C( 8 ), ;
                    ClusterSQL M, ;
                    HashKeys N( 12 ), ;
                    RIError M, ;
                    RIErrNo N ( 5 ), ;
                    ClustErr M, ;
                    ClustErrNo N ( 5 ), ;
                    EXPORT L, ;
                    Exported L, ;
                    Duplicates N( 3 ))

            CASE lcPassed = "Clusters"
                CREATE TABLE &lcTableName FREE;
                    ( ClustName C ( 30 ), ;
                    ClustType C ( 5 ), ;
                    HashKeys N ( 6 ), ;
                    ClustSize N ( 6 ), ;
                    ClusterSQL M, ;
                    ClustErr M, ;
                    ClustErrNo N ( 5 ), ;
                    EXPORT L, ;
                    Exported L, ;
                    Duplicates N( 3 ))

        ENDCASE

        RETURN lcTableName



    FUNCTION CreateNewDir
        LOCAL aDirArray, lcDirName
        *Create directory for upsizing files if it doesn't exist already
        IF this.NewDir=="" THEN
            this.NewDir = NEW_DIRNAME_LOC
            DIMENSION aDirArray[1]
            IF ADIR( aDirArray, this.NewDir, 'D' )=0 THEN
                MD ( this.NewDir )
                this.CreatedNewDir=.T.
            ENDIF
            CD ( this.NewDir )
            SET DEFAULT TO CURDIR()
            this.NewDir=CURDIR()
        ENDIF



    FUNCTION GetFoxDataSize
        LOCAL lcEnumTablesTbl, lcFileName, lcTableStem

        lnTableSize = 0
        lnIndexSize = 0
        this.TBFoxTableSize = 0
        this.TBFoxIndexSize = 0

        lcEnumTablesTbl = this.EnumTablesTbl
        IF EMPTY( lcEnumTablesTbl )
            RETURN
        ENDIF

        SELECT ( lcEnumTablesTbl )

        SCAN FOR EXPORT=.T.
            * get table + memo size
            lcFileName = &lcEnumTablesTbl..TblPath
            IF ADIR( laDir, lcFileName ) <> 0
                lnTableSize = lnTableSize + laDir[1, 2] / 1024.

                IF RAT( '.', lcFileName ) > 0
                    lcTableStem = LEFT( lcFileName, RAT( '.', lcFileName )-1 )
                ENDIF

                * lcTableStem = this.JustStem2( &lcEnumTablesTbl..TblPath )
                lcFileName = lcTableStem + ".fpt"
                IF ADIR( laDir, lcFileName ) <> 0
                    lnTableSize = lnTableSize + laDir[1, 2] / 1024.
                ENDIF

                * get index size
                lcFileName = lcTableStem + ".cdx"
                IF ADIR( laDir, lcFileName ) <> 0
                    lnIndexSize = lnIndexSize + laDir[1, 2] / 1024.
                ENDIF
            ELSE
                * die
            ENDIF
        ENDSCAN

        this.TBFoxTableSize = lnTableSize
        this.TBFoxIndexSize = lnIndexSize


    FUNCTION InsArrayRow
        LPARAMETERS aArray, lcElement1, lcElement2, lcElement3, lcElement4, lcElement5, lcElement6F
        LOCAL lnParams, lnRow

        lnParams = PARAMETERS() - 1

        IF ALEN ( aArray, 1 ) = 1 AND EMPTY( aArray[1, 1] )
            lnRow = 1
        ELSE
            DIMENSION aArray[ALEN( aArray, 1 )+1, lnParams]
            lnRow = ALEN( aArray, 1 )
        ENDIF

        IF lnParams >= 2
            aArray[lnRow, 1] = lcElement1
        ENDIF
        IF lnParams >= 2
            aArray[lnRow, 2] = lcElement2
        ENDIF
        IF lnParams >= 3
            aArray[lnRow, 3] = lcElement3
        ENDIF
        IF lnParams >= 4
            aArray[lnRow, 4] = lcElement4
        ENDIF
        IF lnParams >= 5
            aArray[lnRow, 5] = lcElement5
        ENDIF
        IF lnParams >= 6
            aArray[lnRow, 6] = lcElement6
        ENDIF


    FUNCTION DEBUG
        ACTIVATE WINDOW trace
        ACTIVATE WINDOW DEBUG
        SUSPEND


    *set trunc.log on chkpt. option for database on server
    FUNCTION TruncLogOn
        LOCAL lnOldArea, lcDBName, lnRes

* If we have an extension object and it has a TruncLogOn method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'TruncLogOn', 5 ) and ;
			not this.oExtension.TruncLogOn( This )
			return
		ENDIF

        lcDBName = ALLTRIM( this.ServerDBName )
        lnOldArea = SELECT()
        IF ( SQLEXEC( this.MasterConnHand, "sp_helpdb " ) = 1 ) AND ;
        	!EMPTY( ALIAS() ) and type( 'STATUS' ) <> 'U'
            LOCATE FOR NAME = lcDBName
            IF !EOF()
                this.TruncLog = IIF( ATC( "trunc. log", STATUS ) > 0 or ;
                	ATC( "trunc. log", strconv( STATUS, 6 )) > 0 or ;
                	ATC( "Recovery=SIMPLE", strconv( STATUS, 6 )) > 0, 1, 0 )
            ENDIF
            USE
        ENDIF
        lnRes = SQLEXEC( this.MasterConnHand, "sp_dboption " + lcDBName + ", 'trunc. log on chkpt.', true" ) = 1
        SELECT ( lnOldArea )


    FUNCTION TruncLogOff
        LOCAL lnRes

* If we have an extension object and it has a TruncLogOff method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'TruncLogOff', 5 ) and ;
			not this.oExtension.TruncLogOff( This )
			return
		ENDIF

        * if false or unknown, set to false
        IF this.TruncLog <> 1
            lnRes = SQLEXEC( this.MasterConnHand, "sp_dboption " + ALLTRIM( this.ServerDBName ) + ", 'trunc. log on chkpt.', false" )
        ENDIF
        this.TruncLog = -1


    FUNCTION UpsizeComplete
        LOCAL lcEnumTablesTbl, lcCRLF, myarray[1]

* If we have an extension object and it has a UpsizeComplete method, call it.
* If it returns .F., that means don't continue with the usual processing.

		IF vartype( this.oExtension ) = 'O' and ;
			pemstatus( this.oExtension, 'UpsizeComplete', 5 ) and ;
			not this.oExtension.UpsizeComplete( This )
			return
		ENDIF

        this.cFinishMsg = ALL_DONE_LOC

        lcCRLF = CHR( 10 ) + CHR( 13 )
        lcEnumTablesTbl	= this.EnumTablesTbl
        SELECT( lcEnumTablesTbl )
		SELECT COUNT( * ) FROM ( lcEnumTablesTbl ) WHERE EXPORT and not Exported INTO ARRAY myarray
		IF ( this.DoScripts AND this.DoUpsize ) OR this.DoUpsize
            IF !EMPTY( myarray )
                this.cFinishMsg = ALL_DONE_LOC + lcCRLF + CANTUPSIZE_TABLE_LOC
                SCAN FOR Exported = .F.
                    this.cFinishMsg = this.cFinishMsg + RTRIM( LOWER( &lcEnumTablesTbl..TblName )) + ", "
                ENDSCAN
                this.cFinishMsg = LEFT( this.cFinishMsg, LEN( RTRIM( this.cFinishMsg )) - 1 )
            ENDIF
        ENDIF


    FUNCTION StripFunction
        LPARAMETER tcString

        LOCAL lnPos, lcString, lcDelim, lnOpenDelimOccurs, lnEndPos

        * Chandsr added code for 40543
        && stripping the fox function names from the expression before sending them over to SQL
        lnOpenDelimOccurs = OCCURS( "( ", tcString )

        lnPos = AT( "( ", tcString, IIF ( lnOpenDelimOccurs = 0, 1, lnOpenDelimOccurs ) )
        * End chandsr added code for 40543

        lcString = tcString

        lcDelim = IIF( AT( ", ", tcString ) > 0, ", ", " )" )
        IF lnPos > 0
        	*{ Add and Change JEI RKR 2005.03.30
			lnEndPos = AT( lcDelim, tcString ) - 1 - lnPos
			IF lnEndPos > 0
				lcString = SUBSTR( tcString, lnPos + 1, lnEndPos )
			ELSE
				lcString = SUBSTR( tcString, lnPos + 1 )
			ENDIF
*			lcString = SUBSTR( tcString, lnPos + 1, AT( lcDelim, tcString ) - 1 - lnPos )
        	*} Add JEI RKR 2005.03.30
		ELSE
        	*{ Add and Change JEI RKR 2005.03.30
        	lnEndPos = AT( lcDelim, tcString ) - 1
        	IF lnEndPos > 0
        		lcString = ALLTRIM( LEFT( tcString, lnEndPos ))
        	ENDIF
        	*} Add JEI RKR 2005.03.30
        ENDIF

        RETURN lcString


	&& chandsr added function to make the string ANSI 92 compatible
    FUNCTION MakeStringANSI92Compatible
        LPARAMETER	cSQLString
        LOCAL cRet, lnBetweenLocation, lnComma, lcBetweenExpression, lcBetweenStart, lcBetweenEnd, lnBetweenBegin, lnBetweenEnd
        LOCAL lcBetweenColumnName

        cRet = STRTRAN ( cSQLString, BOOLEAN_FALSE, BOOLEAN_SQL_FALSE, 1, 1, 1 )
        cRet = STRTRAN ( cRet, BOOLEAN_TRUE, BOOLEAN_SQL_TRUE, 1, 1, 1 )

        lnBetweenLocation = ATC ( VFP_BETWEEN, cRet )

        IF ( lnBetweenLocation <> 0 ) THEN
            lcBetweenExpression = SUBSTR ( cRet, lnBetweenLocation )
            lnBetweenStart = ATC ( "( ", lcBetweenExpression )
            lnBetweenEnd = ATC ( " )", lcBetweenExpression )
            lcBetweenExpression = SUBSTR ( cRet, lnBetweenLocation, lnBetweenEnd )

            lnComma = ATC( VFP_COMMA, lcBetweenExpression )
            lcBetweenColumnName = SUBSTR ( lcBetweenExpression, lnBetweenStart + 1, lnComma - lnBetweenStart - 1 )

            lnBetweenStart = lnComma + 1
            lnComma = ATC( VFP_COMMA, lcBetweenExpression, 2 )
            lcBetweenStart = SUBSTR ( lcBetweenExpression, lnBetweenStart, lnComma - lnBetweenStart )
            lcBetweenEnd = SUBSTR ( lcBetweenExpression, lnComma + 1 )
            cRet = STRTRAN ( cRet, lcBetweenExpression, lcBetweenColumnName + " BETWEEN " + this.ConvertToSQLType ( lcBetweenStart ) + " AND " + this.ConvertToSQLType ( lcBetweenEnd ), 1, 1, 1 )
        ENDIF
        RETURN cRet


    FUNCTION ConvertToSQLType
        LPARAMETERS cData
        LOCAL lcLeadingChar, lcData

        lcData = ALLTRIM ( cData )
        lcLeadingChar = SUBSTR( lcData, 1, 1 )

        IF ( lcLeadingChar = "{" ) THEN
            lcData = STRTRAN ( lcData, "{^", "'" )
            lcData = STRTRAN ( lcData, "} )", "'" )
            lcData = STRTRAN ( lcData, "}", "'" )
        ENDIF

        RETURN lcData

	&& chandsr added function to make the string ANSI 92 compatible


	*ADD JEI - RKR - 2005.03.24
	FUNCTION ReadViews
		LOCAL lcSource, laViews[1, 1], lcSQLStr As String
		LOCAL lnOldArea, lnViewCount, lnCurrentView as Integer


        lcSourceDB=this.SourceDB
        lnOldArea=SELECT()
        SET DATABASE TO ( lcSourceDB )
		lnViewCount = ADBOBJECTS( laViews, "VIEW" )
		FOR lnCurrentView = 1 TO lnViewCount
			*If this is a remote view, skip it
			IF DBGETPROP( laViews[lnCurrentView], "View", "SourceType" )=2 THEN
				LOOP
			ENDIF
            lcSQLStr = DBGETPROP( laViews[lnCurrentView], "View", "SQL" )
            INSERT INTO ( this.EnumTablesTbl ) ( TblName, Type, SQlStr,  EXPORT ) ;
            			VALUES ( LOWER( laViews[lnCurrentView] ), "V", lcSQLStr, .t. )

		ENDFOR
		SELECT( this.EnumTablesTbl )
		LOCATE
		SELECT( lnOldArea )


	FUNCTION GetChooseView

		LOCAL lnOldWorkArea as Integer

		DIMENSION this.aChooseViews[1, 1]

		this.aChooseViews[1, 1] = ""
		lnOldWorkArea = SELECT()

		SELECT TblName FROM ( this.EnumTablesTbl ) WHERE Type = "V" AND EXPORT = .T. INTO ARRAY this.aChooseViews

		SELECT( this.EnumTablesTbl )
		Replace EXPORT WITH .F. ALL FOR Type == "V"
		DELETE ALL FOR Type == "V"

		LOCATE
		SELECT( lnOldWorkArea )


	FUNCTION CheckUpsizeView
		LPARAMETERS tcViewName

		RETURN ( ASCAN( this.aChooseViews, tcViewName, 1, ALEN( this.aChooseViews ), 1, 15 ) > 0 )


	&& JEI RKR Add Procedure
	FUNCTION CheckForLocalServer
		LOCAL lcTemDBName, lcTempDBFileName, lcCurrentPath, lcSQLStr, laTemp[1, 1] as String
		LOCAL lnResult as Integer
		LOCAL llServerIsLocal as Logical

		llServerIsLocal = .f.

		lcCurrentPath = FULLPATH( "." )
		lcTemDBName = SYS( 2015 )
		lcTempDBFileName = ADDBS( lcCurrentPath ) + FORCEEXT( lcTemDBName, "mdf" )
		lcSQLStr = "CREATE DATABASE " + lcTemDBName + " ON ( NAME = " + lcTemDBName + "_DAT , " + ;
					"FILENAME = '" +lcTempDBFileName + "' )"

		lnResult = SQLEXEC( this.MasterConnHand, lcSQLStr )
		IF lnResult = 1
			llServerIsLocal = ( ADIR( laTemp, lcTempDBFileName ) = 1 )
			=SQLEXEC( this.MasterConnHand, "DROP DATABASE " + lcTemDBName )
		ENDIF

		this.ServerISLocal = llServerIsLocal

		RETURN llServerIsLocal


	FUNCTION CheckForNullValuesInTable
		LPARAMETERS tcTablaName as String

		LOCAL lcEnumFieldsTbl, lcTempCursorName, lcNullExistStr, lcEmptyDateExist, lcNullExistOR, lcEmptyDateOR as String
		LOCAL lnOldWorkArea as Integer
		LOCAL llRes as Logical
		LOCAL loException as Exception

		lnOldWorkArea = SELECT()
		tcTablaName = UPPER( ALLTRIM( tcTablaName ))
		lcEnumFieldsTbl = RTRIM( this.EnumFieldsTbl )
		lcTempCursorName = this.UniqueCursorName( "CheckForNull" )
		lcNullExistStr = "LOCATE FOR "
		lcEmptyDateExist = "LOCATE FOR "
		lcNullExistOR = ""
		lcEmptyDateOR = ""
		llRes = .T.

		SELECT * FROM &lcEnumFieldsTbl. WHERE upper( ALLTRIM( TblName )) == tcTablaName INTO CURSOR &lcTempCursorName.
		SELECT( lcTempCursorName )
		SCAN
			lcNullExistStr = lcNullExistStr + lcNullExistOR + " ISNULL( " + ALLTRIM( FldName ) + " )"
			lcNullExistOR = " OR "
			IF ALLTRIM( UPPER( DATATYPE )) $ "DT"
				lcEmptyDateExist = lcEmptyDateExist + lcEmptyDateOR + " EMPTY( " + ALLTRIM( FldName ) + " )"
				lcEmptyDateOR = " OR "
			ENDIF
		ENDSCAN
		IF USED( tcTablaName )
			SELECT( tcTablaName )
			TRY
				&lcNullExistStr.
				IF !FOUND()
					IF !EMPTY( lcEmptyDateOR )
						&lcEmptyDateExist.
						IF !FOUND()
							llRes = .F.
						ENDIF
					ELSE
						llRes = .F.
					ENDIF
				ENDIF
			CATCH TO loException
				* Do Nothing
			ENDTRY
		ENDIF
		USE IN ( lcTempCursorName )
		SELECT( lnOldWorkArea )
		RETURN llRes


	FUNCTION GetColumnNum
		LPARAMETERS tcTableName as String, tcDataTypes as String, tlAllowNull as Logical

		LOCAL lnOldWorkArea, lnFieldCount, lnCurrentField as Integer
		LOCAL laDummy[1], lcRes as String

		IF VARTYPE( tcDataTypes ) <> "C" OR EMPTY( tcDataTypes )
			RETURN ""
		ENDIF

		lcRes = ", "
		IF USED( tcTableName )
			lnOldWorkArea = SELECT()
			SELECT( tcTableName )
			lnFieldCount = AFIELDS( laDummy )
			FOR lnCurrentField = 1 TO lnFieldCount
*** DH 07/02/2013: use all DT fields whether they're nullable or not since they may contain
***					blank values
*				IF laDummy[lnCurrentField, 2] $ tcDataTypes AND laDummy[lnCurrentField, 5] = tlAllowNull
				IF laDummy[lnCurrentField, 2] $ tcDataTypes AND ;
					( laDummy[lnCurrentField, 5] = tlAllowNull or laDummy[lnCurrentField, 2] $ 'DT' )
					lcRes = lcRes + TRANSFORM( lnCurrentField ) + ", "
				ENDIF
			ENDFOR
			SELECT( lnOldWorkArea )
		ENDIF

		RETURN lcRes

ENDDEFINE