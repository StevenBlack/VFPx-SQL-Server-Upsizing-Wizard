VFPx-SQL-Server-Upsizing-Wizard
===============================

A fork of the VFP Sedna SQL Server Upsizing Wizard

This is an update to the Visual FoxPro 9.0 SP2 Upsizing Wizard. The Sedna Upsizing Wizard will install to the `Microsoft Visual FoxPro 9\Sedna\UpsizingWizard` folder under `Program Files`. To launch the new wizard run the `UpsizingWizard.app` from this location.

This Sedna update release includes:

* Updated, cleaner look and feel
* Streamlined, simpler steps
* Support for bulk insert to improve performance.
* Allows you to specify the connection as a DBC, a DSN, one of the existing connections or a new connection string.
* Fields using names with reserved SQL keywords are now delimited.
* If lQuiet is set to true when calling the wizard, no UI is displayed. It uses RAISEEVENT() during the progress of the upsizing so the caller can show progress.
* Performance improvement when upsizing to Microsoft SQL Server 2005.
* Trims all Character fields being upsized to Varchar.
* BlankDateValue property available. It specifies that blank dates should be upsized as nulls. (Old behavior was to set them to 01/01/1900).
* Support for an extension object. This allows developers to hook into every step of the upsizing process and change the behavior. Another way is to subclass the engine.
* Support for table names with spaces.
* UpsizingWizard.APP can be started with default settings (via params) for source name and path, target db, and a Boolean indicating if the target database is to be created.


###2013.07.24 Release
Th
is update has the following bug fixes (thanks to Mike Potjer for finding and even fixing some of these issues):

* Upsizing a logical field to a bit field no longer causes an error when the table has a lot of records.
* You no longer get an error upsizing tables that have field rules, table rules, or triggers.
* Quotes in the content of fields are preserved.
* You no longer get a warning that 6.5 compatibility cannot be used.


###2013.07.02 Release

This update has the following bug fixes (thanks to Mike Potjer for helping with these issues):

* Under some conditions, an error occurred while trying to clean up (specifically removing a working directory the wizard created, which may not be empty) after the wizard is done. This code is now wrapped in a TRY to avoid the error.
* A temporary DBF file wasn't deleted but its FPT was, causing an error if you tried to open it. The file is now deleted.
* A temporary file used for bulk upload wasn't erased when the wizard is done with it. It is now.
* One of the export mechanisms ("FastExport") used null dates rather than the defined date for blank VFP date fields. This was fixed.
* You are no longer told that the bulk insert mechanism failed when it in fact succeeds. This also resolves a problem with duplicate records being created.
* Under some conditions, the last few records in a VFP table weren't imported into the SQL Server table. This was fixed.


###2012.12.06 Release
This update has the following bug fixes:

* Handles converting Memo to Varchar(Max)
* You can now change the date used in SQL Server for empty VFP dates in one place: SQLSERVEREMPTYDATEY2K in AllDefs.H
* One of the export mechanisms ("JimExport") used null dates rather than the defined date for blank VFP date fields. This was fixed.
* The progress meter now correctly shows the progress of the sending data for each table.
* A bug that sometimes caused an error sending the last bit of data for a table to SQL Server was fixed.
* You no longer get a "string too long" error when using bulk insert with a record with more than 30 MB of data.
* You can now use field names up to 128 characters; the former limit was 30.


Review [article on CODE Magazine](http://www.code-magazine.com/Article.aspx?quickid=0703052)
