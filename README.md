# Report-Automator
Used to automate the process getting the SCCM reports ready to be sent to principals.

The program takes an unedited Hardware 09A report that has been exported from Configuration Manager Console. It can handles some small edits to the original .xlsx file, but it works best if it's just straight from MCM.

CURRENTLY WORKING ON:
- Taking out the automatic report downloading functionality for now - it is broken due to changing values for the "Collection" dropdown menu at the report server. As of now, I've got no way of getting it right.

Common Errors:
"There was a problem removing one or more unused columns"
- This means that either the columns or column headers in the report file have been edited. You can just manually delete the unused column after the script has done its thing; it should still run fine.
