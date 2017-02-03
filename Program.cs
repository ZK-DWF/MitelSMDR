using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.Net.Mail;

class MitelSMDR
{
    /* This program reads the Mitel SMDR log file and transfers the fields into SQL Server

       See MitelSMDR Development History.TXT for version history
    */

    static string _connectionString = "Data Source=SQLdb;Initial Catalog=HBRCdb;Integrated Security=True;MultipleActiveResultSets=True";
    public static SqlConnection SQLConnection1 = new SqlConnection(_connectionString);
    public static SqlCommand SQLCommand1 = new SqlCommand();

    // used to store constant SQL string
    static string SQLInsertString = "INSERT INTO Phone_Records (DateTime,DateM,TimeM,Duration,CallingParty,CalledParty,TimeToAnswer,Extension,DigitsDialled,CallCompletionStatus,TransferConferenceCall,ThirdParty,SpeedCall) VALUES (@DateTime,@DateM,@TimeM,@Duration,@CallingParty,@CalledParty,@TimeToAnswer,@Extension,@DigitsDialled,@CallCompletionStatus,@TransferConferenceCall,@ThirdParty,@SpeedCall)";

    // used to store variable SQL string
    //static string sql2String = "";

    // command line filename
    static string CommandLineFilename = "";

    static Int32 countInteger;
    static Int32 countSuspiciousActivity;
    static string Incidentlog = "";

    // -------------------------------------------------------------------------------------------
    static void Main(string[] args)
    {
        // Main engine

        TimeStamp("*************** Transfer started");

        //StartDate = Convert.ToDateTime("06-Oct-2011");
        //EndDate = Convert.ToDateTime("20-Oct-2011");

        // get command line filename
        CommandLineFilename = args[0];

        countInteger = 0;
        countSuspiciousActivity = 0;

        // Keep reading files until we reach the EndDate

        ReadLogFile();  // (DateString);

        TimeStamp(@"Records imported :" + countInteger.ToString());
        TimeStamp("*************** Transfer ended");

        System.Environment.Exit(0);
    }

    // -------------------------------------------------------------------------------------------
    private static void ReadLogFile()
    {
        // const string FileLocation = "//ISA/C$/Program Files/Microsoft ISA Server/ISALogs/ISALOG_";
        string StartupPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);

        string textLine = null;
        //string[] testArray = null;

        string DateTimex;
        string DateM = null;
        string TimeM = null;
        Int32 HourM = 0;

        Int32 Duration;
        Int32 Widthx;
        string CallingParty = null;
        string CalledParty = null;
        Int32 TimeToAnswer = 0;
        string Extension = null;
        string DigitsDialled = null;
        string CallCompletionStatus = null;
        string TransferConferenceCall = null;
        string ThirdParty = null;
        string fileNameString = null;
        string SpeedCall = null;

        SqlTransaction myTransaction = null;

        // Open up input log file

        fileNameString = @StartupPath + @"\" + CommandLineFilename; // @"\Mitel_14-02@0115.LOG";
        fileNameString = fileNameString.Substring(6, fileNameString.Length - 6);

        if (!File.Exists(fileNameString))
        {
            Console.WriteLine(@"{0} does not exist.", fileNameString);
            TimeStamp(fileNameString + @" does not exist");
            return;
        }

        // Open up the SQL Server connection
        SQLConnection1.Open();

        // Start the SQL transaction
        myTransaction = SQLConnection1.BeginTransaction();

        SqlCommand Insert_command = new SqlCommand(SQLInsertString, SQLConnection1, myTransaction);

        // Build SQL Parameters
        Insert_command.Parameters.Add("@DateTime", SqlDbType.DateTime);
        Insert_command.Parameters.Add("@DateM", SqlDbType.VarChar, 11);
        Insert_command.Parameters.Add("@TimeM", SqlDbType.VarChar, 5);
        Insert_command.Parameters.Add("@Duration", SqlDbType.Int);
        Insert_command.Parameters.Add("@CallingParty", SqlDbType.VarChar, 25);
        Insert_command.Parameters.Add("@CalledParty", SqlDbType.VarChar, 25);
        Insert_command.Parameters.Add("@TimeToAnswer", SqlDbType.Int);
        Insert_command.Parameters.Add("@Extension", SqlDbType.VarChar, 4);
        Insert_command.Parameters.Add("@DigitsDialled", SqlDbType.VarChar, 40);
        Insert_command.Parameters.Add("@CallCompletionStatus", SqlDbType.Char, 1);
        Insert_command.Parameters.Add("@TransferConferenceCall", SqlDbType.Char, 1);
        Insert_command.Parameters.Add("@ThirdParty", SqlDbType.VarChar, 4);
        Insert_command.Parameters.Add("@SpeedCall", SqlDbType.Char, 1);

        // Read in values from file, and transfer to database

        //try
        //{
        StreamReader MitelSMDRFileReader = new StreamReader(fileNameString);

        // read until End of File is hit
        while (MitelSMDRFileReader.Peek() != -1)
        {
            textLine = MitelSMDRFileReader.ReadLine();

            // ignore any lines that are blank
            if (textLine.Trim().Length > 5)
            {
                DateTimex = ConvertDate1(textLine.Substring(1, 12));
                DateM = ConvertDate2(textLine.Substring(1, 12)).Substring(0, 8);
                TimeM = textLine.Substring(7, 5);
                HourM = Convert.ToInt32(textLine.Substring(7, 2));

                if (textLine.Substring(15, 8) == "        ")
                {
                    Duration = 0;
                }
                else
                {
                    Duration = Convert.ToInt32(textLine.Substring(14, 2)) * 60 * 60 + (Convert.ToInt32(textLine.Substring(17, 2)) * 60) + Convert.ToInt32(textLine.Substring(20, 2));
                }

                // Check Calling Party and Called Party
                // These columns 34-59 can have 1, 2, 3 or 4 sub fields ... just read the first two

                CallingParty = textLine.Substring(23, 4);
                CallingParty = CallingParty.Trim();

                Widthx = 25;

                if (textLine.Length < 59) Widthx = textLine.Length - 33;
                CalledParty = textLine.Substring(33, Widthx).Trim();

                if (textLine.Substring(28, 4) == "    ")
                {
                    TimeToAnswer = 0;
                }
                else if (textLine.Substring(28, 4) == "****")
                {
                    TimeToAnswer = -1;
                }
                else
                {
                    TimeToAnswer = Convert.ToInt32(textLine.Substring(28, 4));
                }

                Extension = "";  // textLine.Substring(33, 4);
                DigitsDialled = ""; //  textLine.Substring(38, 20).TrimEnd();

                if (textLine.Length >= 60)
                {
                    CallCompletionStatus = textLine.Substring(59, 1);
                }
                else
                {
                    CallCompletionStatus = " ";
                }

                if (textLine.Length > 61)
                {
                    SpeedCall = textLine.Substring(60, 1);
                }
                else
                {
                    SpeedCall = " ";
                }

                if (textLine.Length >= 66)
                {
                    TransferConferenceCall = textLine.Substring(65, 1);
                }
                else
                {
                    TransferConferenceCall = " ";
                }

                if (textLine.Length >= 71)
                {
                    ThirdParty = textLine.Substring(67, 4);
                }
                else
                {
                    ThirdParty = " ";
                }

                // Insert line into SQL
                Insert_command.Parameters["@DateTime"].Value = DateTimex;
                Insert_command.Parameters["@DateM"].Value = DateM;
                Insert_command.Parameters["@TimeM"].Value = TimeM;
                Insert_command.Parameters["@Duration"].Value = Duration;
                Insert_command.Parameters["@CallingParty"].Value = CallingParty;
                Insert_command.Parameters["@CalledParty"].Value = CalledParty;
                Insert_command.Parameters["@TimeToAnswer"].Value = TimeToAnswer;
                Insert_command.Parameters["@Extension"].Value = Extension;
                Insert_command.Parameters["@DigitsDialled"].Value = DigitsDialled;
                Insert_command.Parameters["@CallCompletionStatus"].Value = CallCompletionStatus;
                Insert_command.Parameters["@TransferConferenceCall"].Value = TransferConferenceCall;
                Insert_command.Parameters["@ThirdParty"].Value = ThirdParty;
                Insert_command.Parameters["@SpeedCall"].Value = SpeedCall;
                Insert_command.ExecuteNonQuery();

                countInteger += 1;

                // Check for Suspicious Activity
                // e.g. accessing voice mail from outside between line 2300 to 0700 hours

                if (HourM == 23 || HourM <= 7)
                {
                    if (CallingParty.StartsWith("T") && CalledParty.Contains("9200 7777"))
                    {
                        if (CallingParty.StartsWith("T68") || CallingParty.StartsWith("T63") || CallingParty.StartsWith("T27"))
                        {
                            // ignore any calls from local numbers or mobiles
                        }
                        else
                        {
                            countSuspiciousActivity += 1;
                            OutputError(textLine);
                            if (countSuspiciousActivity <= 200)
                            {
                                Incidentlog += textLine + "\n";
                            }
                        }
                    }
                }
            }
        }
        MitelSMDRFileReader.Close();
        //}
        //catch (Exception e)
        //{
        //   Console.WriteLine("{0} Exception caught.", e);
        //   TimeStamp(e.Message);
        //}

        // if Suspicious Activity found, email out a warning
        if (countSuspiciousActivity > 3)
        {
            EmailWarning();
        }
        myTransaction.Commit();
        SQLConnection1.Close();
    }

    // -------------------------------------------------------------------------------------------
    public static void TimeStamp(string status)
    {
        const String timestampFilename = @"\\FileServ\Common\Timesheets\MITEL_Phone.log";

        // Writes current time and comment to the Audit Log

        if (File.Exists(timestampFilename))
        {
            StreamWriter timestampSW;
            timestampSW = File.AppendText(timestampFilename);
            timestampSW.WriteLine('"' + DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss") + '"' + ',' + '"' + status + '"');
            timestampSW.Close();
        }
    }

    // -------------------------------------------------------------------------------------------
    public static void OutputError(string status)
    {
        string StartupPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
        string errorFilename = @StartupPath + @"\" + @"MITEL_Error.log";
        string localpath = new Uri(errorFilename).LocalPath;

        // Writes current time and comment to the Audit Log

        if (File.Exists(localpath))
        {
            StreamWriter errorSW;
            errorSW = File.AppendText(localpath);
            errorSW.WriteLine(status);
            errorSW.Close();
        }
    }

    // -------------------------------------------------------------------------------------------
    public static string QuoteMarks(string inString)
    {
        // Firstly, replace single quotes with backward quote mark
        //string outString;

        //outString = inString.Replace(@"'", @"''");

        return "\'" + inString + "\'";
    }

    // -------------------------------------------------------------------------------------------
    //public static bool IsNumeric(string strToCheck)
    //{
    //   return Regex.IsMatch(strToCheck, "^\\d+(\\.\\d+)?$");
    //}

    // -------------------------------------------------------------------------------------------
    public static String ConvertDate1(string inDate)
    {
        string Year = null;
        string Mnth = null;

        Year = DateTime.Now.ToString("yyyy");
        // if last day of year, change to previous year
        if (inDate.Substring(0, 5) == "12/31") Year = Convert.ToString(Convert.ToInt32(Year) - 1);

        switch (inDate.Substring(0, 2))
        {
            case "01":
                Mnth = "Jan";
                break;
            case "02":
                Mnth = "Feb";
                break;
            case "03":
                Mnth = "Mar";
                break;
            case "04":
                Mnth = "Apr";
                break;
            case "05":
                Mnth = "May";
                break;
            case "06":
                Mnth = "Jun";
                break;
            case "07":
                Mnth = "Jul";
                break;
            case "08":
                Mnth = "Aug";
                break;
            case "09":
                Mnth = "Sep";
                break;
            case "10":
                Mnth = "Oct";
                break;
            case "11":
                Mnth = "Nov";
                break;
            case "12":
                Mnth = "Dec";
                break;
        }

        return inDate.Substring(3, 2) + "-" + Mnth + "-" + Year + " " + inDate.Substring(6, 5);
    }

    // -------------------------------------------------------------------------------------------
    public static String ConvertDate2(string inDate)
    {
        string Year = null;

        Year = DateTime.Now.ToString("yyyy");
        // if last day of year, change to previous year
        if (inDate.Substring(0, 5) == "12/31") Year = Convert.ToString(Convert.ToInt32(Year) - 1);

        return Year + inDate.Substring(0, 2) + inDate.Substring(3, 2) + inDate.Substring(6, 5);
    }

    // -------------------------------------------------------------------------------------------
    public static void EmailWarning()
    {
        // Create and configure the SMTPclient that will send the email.
        // Specify the host name of the SMTP server and the port used to send email.
        SmtpClient client = new SmtpClient("mail64.hbrc.govt.nz", 25);

        // Create the MailMessage to represent the e-mail being sent.
        using (MailMessage msg = new MailMessage())
        {
            // Configure the e-mail sender and subject.
            msg.From = new MailAddress("Administrator@hbrc.govt.nz");
            msg.To.Add("dave@hbrc.govt.nz");
            msg.To.Add("Helpdesk@hbrc.govt.nz");
            msg.Subject = "Possible Suspicious Activity found in Mitel System";

            // Configure the e-mail body.
            string MessageM = "The overnight process transferring the Mitel Log data to SQL Server has detected occurrences of possible {0} suspicious logons attempts overnight. Please check the error log files  \\\\Techserv\\C$\\Program Files\\Utilities\\Mitel\\Mitel_Error.LOG as soon as possible.\n\n" + Incidentlog;

            msg.Body = string.Format(MessageM, countSuspiciousActivity);

            // Send the message
            client.Send(msg);
        }
    }
}