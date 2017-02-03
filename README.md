# MitelSMDR
Import a Mitel SMDR file and load into SQL Server

This console application runs on the server, and reads the MITEL.LOG produced from the Mitel Phone system and transfers it to a SQL Server table. The log file is in Station Messaging Detail Record (SMDR) format, a universal format for recording telecommunications system activity.

Version    Date        Comment

1.0.0      16-Dec-2011 Migration from VB.NET to C#
1.1.0      19-Jul-2013 Added SQL Parameters, and email warning for suspicious activity
1.2.0      13-Aug-2013 Suspicious Activity logs to MITEL_Error.log
1.3.0      05-Dec-2013 Added SpeedCall field
1.4.0      12-May-2014 Outputs first 300 incidents to the email body
1.4.1      16-May-2014 Changed suspicious activity from 23.00-06:00, to 23.00-07:00
	           						... changed wording in email
						          	... scheduled task now runs at 07:00 hours
1.4.2      19-May-2014 Only email suspicious activity if greater than 3 occurrences
1.4.3      18-Mar-2015 Added feature to ignore local calls and mobiles from suspicious activity                           
1.4.4      01-Aug-2015 Changed SQL2K8 to SQLdb
1.4.5      09-Jan-2017 If last day of year, change date to previous year

TO DO :                
