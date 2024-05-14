using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Data;

namespace SPCAccessModule
{
    public class SPCAccessDB
    {
        bool flag = true;
        public OleDbConnection connection;
        OleDbDataReader reader;
        string str;
        string conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
        static readonly object locker = new object(); 

        public SPCAccessDB(string filePath)
        {
            //lock (locker)
            //{
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
           //}
        }

        #region Process Table
        public void insertIntoProcessTable(string filePath, string processName, string partName, string partNo)
        {
            try
            {
                connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
                if (connection.State == System.Data.ConnectionState.Closed)
                    connection.Open();

                OleDbTransaction transaction;
                //start transaction
                transaction = connection.BeginTransaction();
                try
                {
                    OleDbCommand cmdInsertIntoProcess = new OleDbCommand("Insert into Process(ProcName,PartName,PartNo) values(@ProcName,@PartName,@PartNo)", connection);
                    cmdInsertIntoProcess.Transaction = transaction;
                    cmdInsertIntoProcess.Parameters.Add("@ProcName", OleDbType.VarChar).Value = processName;
                    if (partName.Equals(string.Empty) == false)
                        cmdInsertIntoProcess.Parameters.Add("@PartName", OleDbType.VarChar).Value = partName;
                    else
                        cmdInsertIntoProcess.Parameters.Add("@PartName", OleDbType.VarChar).Value = String.Empty;

                    if (partNo.Equals(string.Empty) == false)
                        cmdInsertIntoProcess.Parameters.Add("@PartNo", OleDbType.VarChar).Value = partNo;
                    else
                        cmdInsertIntoProcess.Parameters.Add("@PartNo", OleDbType.VarChar).Value = String.Empty;

                    //if (procDoc.Equals(string.Empty) == false)
                    //    cmdInsertIntoProcess.Parameters.Add("@ProcDoc", OleDbType.VarChar).Value = procDoc;
                    //else
                    //    cmdInsertIntoProcess.Parameters.Add("@ProcDoc", OleDbType.VarChar).Value = String.Empty;

                    cmdInsertIntoProcess.ExecuteNonQuery();
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception : " + ex.Message);
                    // Attempt to roll back the transaction.
                    try
                    {
                        transaction.Rollback();
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred
                        // on the server that would cause the rollback to fail, such as
                        // a closed connection.
                        MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                        MessageBox.Show("Message:" + ex2.Message);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception : "+e.ToString());
            }
        }

        public void deleteFromProcessTable(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteAllCharacteristics = new OleDbCommand("Delete from Process", connection);
                cmdDeleteAllCharacteristics.Transaction = transaction;
                cmdDeleteAllCharacteristics.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Characterstic Table
        public void insert_Characteristics(string filePath, string charName, string charType, int chartType, double SGSize, double? target, double? USL, double? LSL, int DataEntry)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd_insertCharacteristics = new OleDbCommand("Insert into Characterstic(CharName,CharType,ChartType,SGSize,Target,USL,LSL,DataEntry) values(@CharName,@CharType,@ChartType,@SGSize,@Target,@USL,@LSL,@DataEntry)", connection);
                cmd_insertCharacteristics.Transaction = transaction;

                cmd_insertCharacteristics.Parameters.Add("@CharName", OleDbType.VarChar).Value = charName;
                cmd_insertCharacteristics.Parameters.Add("@CharType", OleDbType.VarChar).Value = charType;
                cmd_insertCharacteristics.Parameters.Add("@ChartType", OleDbType.Integer).Value = chartType;
                cmd_insertCharacteristics.Parameters.Add("@SGSize", OleDbType.Double).Value = SGSize;

                if (target != null)
                    cmd_insertCharacteristics.Parameters.Add("@Target", OleDbType.Double).Value = target;
                else
                    cmd_insertCharacteristics.Parameters.Add("@Target", OleDbType.Double).Value = DBNull.Value;

                if (USL != null)
                    cmd_insertCharacteristics.Parameters.Add("@USL", OleDbType.Double).Value = USL;
                else
                    cmd_insertCharacteristics.Parameters.Add("@USL", OleDbType.Double).Value = DBNull.Value;

                if (LSL != null)
                    cmd_insertCharacteristics.Parameters.Add("@LSL", OleDbType.Double).Value = LSL;
                else
                    cmd_insertCharacteristics.Parameters.Add("@LSL", OleDbType.Double).Value = DBNull.Value;
                cmd_insertCharacteristics.Parameters.Add("@DataEntry", OleDbType.Integer).Value = DataEntry;

                cmd_insertCharacteristics.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void delete_All_Characteristice(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);//+ ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteAllCharacteristics = new OleDbCommand("Delete from Characterstic", connection);
                cmdDeleteAllCharacteristics.Transaction = transaction;
                cmdDeleteAllCharacteristics.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        
        public void UpdateCharacteristics(string filePath,string oldcharName,string newcharName, double? target, double? USL, double? LSL,int DataEntry)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateChar = new OleDbCommand("Update Characterstic set CharName=@newcharName,Target=@target,USL=@USL,LSL=@LSL,DataEntry=@DataEntry where CharName='" + oldcharName + "'", connection);
                cmdUpdateChar.Transaction = transaction;
                cmdUpdateChar.Parameters.Add("@newcharName", OleDbType.VarChar).Value = newcharName;
                if (target != null)
                    cmdUpdateChar.Parameters.Add("@target", OleDbType.Double).Value = target;
                else
                    cmdUpdateChar.Parameters.Add("@target", OleDbType.Double).Value = System.DBNull.Value;
                if (USL != null)
                    cmdUpdateChar.Parameters.Add("@USL", OleDbType.Double).Value = USL;
                else
                    cmdUpdateChar.Parameters.Add("@USL", OleDbType.Double).Value = System.DBNull.Value;
                if (LSL != null)
                    cmdUpdateChar.Parameters.Add("@LSL", OleDbType.Double).Value = LSL;
                else
                    cmdUpdateChar.Parameters.Add("@LSL", OleDbType.Double).Value = System.DBNull.Value;
                cmdUpdateChar.Parameters.Add("@DataEntry", OleDbType.Integer).Value = DataEntry;
                cmdUpdateChar.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        #endregion

        #region trace.dat

        #region Trace data
        public OleDbDataReader read_TraceHeaders(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadTraceHeaders = new OleDbCommand("Select * from TraceHeader", connection);
                cmdReadTraceHeaders.Transaction = transaction;
                reader = cmdReadTraceHeaders.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public void insert_TraceHeader(string filePath, string traceDesc)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdInsertTraceHeader = new OleDbCommand("Insert into TraceHeader(TraceDes) values(@TraceDes)", connection);
                cmdInsertTraceHeader.Transaction = transaction;
                cmdInsertTraceHeader.Parameters.Add("@TraceDes", OleDbType.VarChar).Value = traceDesc;
                cmdInsertTraceHeader.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void rename_TraceHeader(string filePath, string oldtraceDesc, string newtraceDesc)
        {
            connection = new OleDbConnection(conString + filePath );//+ ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateTraceHeader = new OleDbCommand("Update TraceHeader set TraceDes=@newTraceDes where TraceDes=@oldTraceDes", connection);
                cmdUpdateTraceHeader.Transaction = transaction;
                cmdUpdateTraceHeader.Parameters.Add("@newTraceDes", OleDbType.VarChar).Value = newtraceDesc;
                cmdUpdateTraceHeader.Parameters.Add("@oldTraceDes", OleDbType.VarChar).Value = oldtraceDesc;
                cmdUpdateTraceHeader.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void delete_traceHeader(string filePath, string traceDesc)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteTrace = new OleDbCommand("Delete from TraceHeader where TraceDes=@TraceDes", connection);
                cmdDeleteTrace.Transaction = transaction;
                cmdDeleteTrace.Parameters.Add("@TraceDes", OleDbType.VarChar).Value = traceDesc;
                cmdDeleteTrace.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public bool check_traceHeader(string filePath, string traceDesc)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdCheckTraceHeader = new OleDbCommand("Select * from TraceHeader where TraceDes like @TraceDes", connection);
                cmdCheckTraceHeader.Transaction = transaction;
                cmdCheckTraceHeader.Parameters.Add("@TraceDes", OleDbType.VarChar).Value = traceDesc;
                reader = cmdCheckTraceHeader.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            if (reader.Read())
                return true;
            else
                return false;
        }

        public int read_TGID(string filePath, string traceDesc)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadTGID = new OleDbCommand("Select TGID from TraceHeader where TraceDes=@TraceDes", connection);
                cmdReadTGID.Transaction = transaction;
                cmdReadTGID.Parameters.Add("@TraceDes", OleDbType.VarChar).Value = traceDesc;
                str = cmdReadTGID.ExecuteScalar().ToString();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return Convert.ToInt32(str);
        }

        public void delete_TGID(string filePath, int TGID)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteTGID = new OleDbCommand("Delete * from TraceDetail where TGID=@TGID", connection);
                cmdDeleteTGID.Transaction = transaction;
                cmdDeleteTGID.Parameters.Add("@TGID", OleDbType.Integer).Value = TGID;

                cmdDeleteTGID.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public OleDbDataReader read_Tracecat(string filePath, int TGID)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd_Tracecat = new OleDbCommand("Select distinct(Tracecat) from TraceDetail where TGID=@TGID", connection);
                cmd_Tracecat.Transaction = transaction;
                cmd_Tracecat.Parameters.Add("@TGID", OleDbType.Integer).Value = TGID;
                reader = cmd_Tracecat.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader read_TraceDetail(string filePath, int TGID)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd_Tracecat = new OleDbCommand("Select * from TraceDetail where TGID=@TGID", connection);
                cmd_Tracecat.Transaction = transaction;
                cmd_Tracecat.Parameters.Add("@TGID", OleDbType.Integer).Value = TGID;
                reader = cmd_Tracecat.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public bool check_Tracecat(string filePath, string tracecat, int TGID)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdCheckTraceCat = new OleDbCommand("Select * from TraceDetail where Tracecat=@Tracecat and TGID=@TGID", connection);
                cmdCheckTraceCat.Transaction = transaction;
                cmdCheckTraceCat.Parameters.Add("@Tracecat", OleDbType.VarChar).Value = tracecat;
                cmdCheckTraceCat.Parameters.Add("@TGID", OleDbType.Integer).Value = TGID;
                reader = cmdCheckTraceCat.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            if (reader.Read())
                return true;
            else
                return false;
        }

        public void update_Tracecat(string filePath, int TGID, string oldTracecat, string newTracecat)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateTracecat = new OleDbCommand("Update TraceDetail set Tracecat=@newTracecat where Tracecat=@oldTracecat and TGID=@TGID", connection);
                cmdUpdateTracecat.Transaction = transaction;
                cmdUpdateTracecat.Parameters.Add("@newTracecat", OleDbType.VarChar).Value = newTracecat;
                cmdUpdateTracecat.Parameters.Add("@oldTracecat", OleDbType.VarChar).Value = oldTracecat;
                cmdUpdateTracecat.Parameters.Add("@TGID", OleDbType.Integer).Value = TGID;

                cmdUpdateTracecat.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void delete_Tracecat(string filePath, int TGID, string Tracecat)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteTracecat = new OleDbCommand("Delete from TraceDetail where Tracecat=@Tracecat and TGID=@TGID", connection);
                cmdDeleteTracecat.Transaction = transaction;
                cmdDeleteTracecat.Parameters.Add("@Tracecat", OleDbType.VarChar).Value = Tracecat;
                cmdDeleteTracecat.Parameters.Add("@TGID", OleDbType.Integer).Value = TGID;

                cmdDeleteTracecat.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public OleDbDataReader read_Tracevalue(string filePath, int TGID, string Tracecat)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdreadTracevalue = new OleDbCommand("Select Tracevalue from TraceDetail where Tracecat=@Tracecat and TGID=@TGID", connection);
                cmdreadTracevalue.Transaction = transaction;
                cmdreadTracevalue.Parameters.Add("@Tracecat", OleDbType.VarChar).Value = Tracecat;
                cmdreadTracevalue.Parameters.Add("@TGID", OleDbType.Integer).Value = TGID;
                reader = cmdreadTracevalue.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public void insert_TraceValue(string filePath, int TGID, string Tracecat, string traceValue)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
              OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdinsertTraceValue = new OleDbCommand("Insert into TraceDetail(TGID,Tracecat,Tracevalue) values(@TGID,@Tracecat,@Tracevalue)", connection);
                cmdinsertTraceValue.Transaction = transaction;
                cmdinsertTraceValue.Parameters.Add("@TGID", OleDbType.Integer).Value = TGID;
                cmdinsertTraceValue.Parameters.Add("@Tracecat", OleDbType.VarChar).Value = Tracecat;
                cmdinsertTraceValue.Parameters.Add("@Tracevalue", OleDbType.VarChar).Value = traceValue;
              

                cmdinsertTraceValue.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void update_TraceValue(string filePath, int TGID, string Tracecat, string oldTraceValue, string newTraceValue)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdupdateTraceValue = new OleDbCommand("Update TraceDetail set TraceValue=@newTraceValue where TGID=@TGID and Tracecat=@Tracecat and TraceValue=@oldTraceValue", connection);
                cmdupdateTraceValue.Transaction = transaction;
                cmdupdateTraceValue.Parameters.Add("@newTraceValue", OleDbType.VarChar).Value = newTraceValue;
        
                cmdupdateTraceValue.Parameters.Add("@TGID", OleDbType.Integer).Value = TGID;
                cmdupdateTraceValue.Parameters.Add("@Tracecat", OleDbType.VarChar).Value = Tracecat;
                cmdupdateTraceValue.Parameters.Add("@oldTraceValue", OleDbType.VarChar).Value = oldTraceValue;
                cmdupdateTraceValue.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void delete_TraceValue(string filePath, int TGID, string Tracecat, string TraceValue)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmddeleteTraceValue = new OleDbCommand("Delete from TraceDetail where TGID=@TGID and Tracecat=@Tracecat and TraceValue=@oldTraceValue", connection);
                cmddeleteTraceValue.Transaction = transaction;
                cmddeleteTraceValue.Parameters.Add("@TGID", OleDbType.Integer).Value = TGID;
                cmddeleteTraceValue.Parameters.Add("@Tracecat", OleDbType.VarChar).Value = Tracecat;
                cmddeleteTraceValue.Parameters.Add("@oldTraceValue", OleDbType.VarChar).Value = TraceValue;
                cmddeleteTraceValue.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        #endregion

        #region Event data

        public OleDbDataReader read_EventHeader(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadEventHeader = new OleDbCommand("Select * from EventHeader", connection);
                cmdReadEventHeader.Transaction = transaction;
                reader = cmdReadEventHeader.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public void insert_EventHeader(string filePath, string EventDes)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdInsertEventHeader = new OleDbCommand("Insert into EventHeader(EventDes) values(@EventDes)", connection);
                cmdInsertEventHeader.Transaction = transaction;
                cmdInsertEventHeader.Parameters.Add("@EventDes", OleDbType.VarChar).Value = EventDes;
                cmdInsertEventHeader.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void rename_EventHeader(string filePath, string oldEventDes, string newEventDes)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateEventHeader = new OleDbCommand("Update EventHeader set EventDes=@newEventDes where EventDes=@oldEventDes", connection);
                cmdUpdateEventHeader.Transaction = transaction;
                cmdUpdateEventHeader.Parameters.Add("@newEventDes", OleDbType.VarChar).Value = newEventDes;
                cmdUpdateEventHeader.Parameters.Add("@oldEventDes", OleDbType.VarChar).Value = oldEventDes;
                cmdUpdateEventHeader.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void delete_EventHeader(string filePath, string EventDes)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteEventHeader = new OleDbCommand("Delete from EventHeader where EventDes=@EventDes", connection);
                cmdDeleteEventHeader.Transaction = transaction;
                cmdDeleteEventHeader.Parameters.Add("@EventDes", OleDbType.VarChar).Value = EventDes;
                cmdDeleteEventHeader.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public bool check_EventHeader(string filePath, string EventDes)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
              OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdCheckEventHeader = new OleDbCommand("Select * from EventHeader where EventDes like @EventDes", connection);
                cmdCheckEventHeader.Transaction = transaction;
                cmdCheckEventHeader.Parameters.Add("@EventDes", OleDbType.VarChar).Value = EventDes;
                reader = cmdCheckEventHeader.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            if (reader.Read())
                return true;
            else
                return false;
        }

        public int read_EventID(string filePath, string EventDes)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
              OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadEventID = new OleDbCommand("Select EventID from EventHeader where EventDes=@EventDes", connection);
                cmdReadEventID.Transaction = transaction;
                cmdReadEventID.Parameters.Add("@EventDes", OleDbType.VarChar).Value = EventDes;
                str = cmdReadEventID.ExecuteScalar().ToString();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return Convert.ToInt32(str);
        }

        public string Dummy(string filePath, string EventDes)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadEventID = new OleDbCommand("Select EventID from EventHeader where EventDes=@EventDes", connection);
                cmdReadEventID.Transaction = transaction;
                cmdReadEventID.Parameters.Add("@EventDes", OleDbType.VarChar).Value = EventDes;
                str = cmdReadEventID.ExecuteScalar().ToString();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return str;
        }

        public void delete_EventID(string filePath, int EventID)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteEventID = new OleDbCommand("Delete * from EventDetail where EventID=@EventID", connection);
                cmdDeleteEventID.Transaction = transaction;
                cmdDeleteEventID.Parameters.Add("@EventID", OleDbType.Integer).Value = EventID;

                cmdDeleteEventID.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public OleDbDataReader read_Eventcat(string filePath, int EventID)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd_Eventcat = new OleDbCommand("Select distinct(Eventcat) from EventDetail where EventID=@EventID", connection);
                cmd_Eventcat.Transaction = transaction;
                cmd_Eventcat.Parameters.Add("@EventID", OleDbType.Integer).Value = EventID;
                reader = cmd_Eventcat.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader read_EventDetail(string filePath, int EventID)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd_EventDetail = new OleDbCommand("Select * from EventDetail where EventID=@EventID", connection);
                cmd_EventDetail.Transaction = transaction;
                cmd_EventDetail.Parameters.Add("@EventID", OleDbType.Integer).Value = EventID;
                reader = cmd_EventDetail.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public bool check_Eventcat(string filePath, string Eventcat, int EventID)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdCheckEventcat = new OleDbCommand("Select * from EventDetail where Eventcat=@Eventcat and EventID=@EventID", connection);
                cmdCheckEventcat.Transaction = transaction;
                cmdCheckEventcat.Parameters.Add("@Eventcat", OleDbType.VarChar).Value = Eventcat;
                cmdCheckEventcat.Parameters.Add("@EventID", OleDbType.Integer).Value = EventID;
                reader = cmdCheckEventcat.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            if (reader.Read())
                return true;
            else
                return false;
        }

        public void update_Eventcat(string filePath, int EventID, string oldEventcat, string newEventcat)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateEventcat = new OleDbCommand("Update EventDetail set Eventcat=@newEventcat where Eventcat=@oldEventcat and EventID=@EventID", connection);
                cmdUpdateEventcat.Transaction = transaction;
                cmdUpdateEventcat.Parameters.Add("@newEventcat", OleDbType.VarChar).Value = newEventcat;
                cmdUpdateEventcat.Parameters.Add("@oldEventcat", OleDbType.VarChar).Value = oldEventcat;
                cmdUpdateEventcat.Parameters.Add("@EventID", OleDbType.Integer).Value = EventID;

                cmdUpdateEventcat.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void delete_Eventcat(string filePath, int EventID, string Eventcat)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteEventcat = new OleDbCommand("Delete from EventDetail where Eventcat=@Eventcat and EventID=@EventID", connection);
                cmdDeleteEventcat.Transaction = transaction;
                cmdDeleteEventcat.Parameters.Add("@Eventcat", OleDbType.VarChar).Value = Eventcat;
                cmdDeleteEventcat.Parameters.Add("@EventID", OleDbType.Integer).Value = EventID;

                cmdDeleteEventcat.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public OleDbDataReader read_Eventvalue(string filePath, int EventID, string Eventcat)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdreadEventvalue = new OleDbCommand("Select Eventvalue from EventDetail where Eventcat=@Eventcat and EventID=@EventID", connection);
                cmdreadEventvalue.Transaction = transaction;
                cmdreadEventvalue.Parameters.Add("@Eventcat", OleDbType.VarChar).Value = Eventcat;
                cmdreadEventvalue.Parameters.Add("@EventID", OleDbType.Integer).Value = EventID;
                reader = cmdreadEventvalue.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public void insert_Eventvalue(string filePath, int EventID, string Eventcat, string Eventvalue)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdinsertEventvalue = new OleDbCommand("Insert into EventDetail(EventID,Eventcat,Eventvalue) values(@EventID,@Eventcat,@Eventvalue)", connection);
                cmdinsertEventvalue.Transaction = transaction;
                cmdinsertEventvalue.Parameters.Add("@EventID", OleDbType.Integer).Value = EventID;
                cmdinsertEventvalue.Parameters.Add("@Eventcat", OleDbType.VarChar).Value = Eventcat;
                cmdinsertEventvalue.Parameters.Add("@Eventvalue", OleDbType.VarChar).Value = Eventvalue;
              
                
                cmdinsertEventvalue.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void update_Eventvalue(string filePath, int EventID, string Eventcat, string oldEventvalue, string newEventvalue)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdupdateEventvalue = new OleDbCommand("Update EventDetail set Eventvalue=@newEventvalue where EventID=@EventID and Eventcat=@Eventcat and Eventvalue=@oldEventvalue", connection);
                cmdupdateEventvalue.Transaction = transaction;
                cmdupdateEventvalue.Parameters.Add("@newEventvalue", OleDbType.VarChar).Value = newEventvalue;
                cmdupdateEventvalue.Parameters.Add("@EventID", OleDbType.Integer).Value = EventID;
                cmdupdateEventvalue.Parameters.Add("@Eventcat", OleDbType.VarChar).Value = Eventcat;
                cmdupdateEventvalue.Parameters.Add("@oldEventvalue", OleDbType.VarChar).Value = oldEventvalue;
                cmdupdateEventvalue.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void delete_Eventvalue(string filePath, int EventID, string Eventcat, string Eventvalue)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmddeleteEventvalue = new OleDbCommand("Delete from EventDetail where EventID=@EventID and Eventcat=@Eventcat and Eventvalue=@Eventvalue", connection);
                cmddeleteEventvalue.Transaction = transaction;
                cmddeleteEventvalue.Parameters.Add("@EventID", OleDbType.Integer).Value = EventID;
                cmddeleteEventvalue.Parameters.Add("@Eventcat", OleDbType.VarChar).Value = Eventcat;
                cmddeleteEventvalue.Parameters.Add("@Eventvalue", OleDbType.VarChar).Value = Eventvalue;
                cmddeleteEventvalue.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        #endregion

        #endregion

        # region Close Connection
        public OleDbConnection CloseConnection()
        {
            lock (locker)
            {
                if (connection.State == System.Data.ConnectionState.Open)
                    connection.Close();
            }
            return connection;
        }
        # endregion

        #region Table : Characterstic

        #region Main screen :Read CharId , CharName,USL,LSL
        /// <summary>
        /// Read CharId,CharName,USL,LSL,SGSize from Characterstic table
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <returns>All read data</returns>
        public OleDbDataReader readUniqueCharId(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUniqueCharId = new OleDbCommand("Select CharID,CharName,USL,LSL,SGSize,CharType from Characterstic", connection);
                cmdUniqueCharId.Transaction = transaction;
                reader = cmdUniqueCharId.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        public OleDbDataReader ReadCharacterstic(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdChar = new OleDbCommand("Select CharName from Characterstic", connection);
                cmdChar.Transaction = transaction;
                reader = cmdChar.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader ReadCharInfo(string filePath, string Characterstic)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdChar = new OleDbCommand("Select CharName,CharType,ChartType,SGSize,Target,USL,LSL,DataEntry from Characterstic where CharName = '" + Characterstic + "'", connection);
                cmdChar.Transaction = transaction;
                reader = cmdChar.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader ReadCharId_USL_LSL(string filePath, string Characterstic)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdChar = new OleDbCommand("Select Charid,USL,LSL,ChartType,SGSize,Target from Characterstic where CharName = '" + Characterstic + "'", connection);
                cmdChar.Transaction = transaction;
                reader = cmdChar.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        #endregion

        #region Table : SGDATA

        #region Main screen:Check if spc file new or old
        /// <summary>
        /// To check if spc is old or new by identifiying if SGNO column is present in SGDATA table
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <returns>true if column present(new file) else false</returns>
        public bool checkSPCFile(string filePath)
        {
            string colName = "";
            int counter = 0;
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadSGDATA = new OleDbCommand("Select * from SGDATA", connection);
                cmdReadSGDATA.Transaction = transaction;
                reader = cmdReadSGDATA.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            while (counter <= reader.FieldCount - 1)
            {
                colName = reader.GetName(counter);
                if (colName.Equals("SGNO"))
                {
                    CloseConnection();
                    return true;
                }
                else
                    counter++;
            }
            CloseConnection();
            return false;
        }
        #endregion

        #region Main screen:Delete column SGCharID from table SGDATA
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        public void deleteColumnSGCharID(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteColumn = new OleDbCommand("Alter table SGDATA drop column SGCharID", connection);
                cmdDeleteColumn.Transaction = transaction;
                cmdDeleteColumn.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Main screen: Add Column SGNO in table SGDATA
        /// <summary>
        /// Adding column SGNo in SGDATA table 
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        public void addColumnSGNO(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdAddColumn = new OleDbCommand("Alter table SGDATA add SGNO integer", connection);
                cmdAddColumn.Transaction = transaction;
                cmdAddColumn.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Main screen : read unique SGCharId
        /// <summary>
        /// Reading unique SGCharID according to charId from SGDATA table
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="charId">CharId whose respecive unique SGCharId is to be read</param>
        /// <returns>All SGCharId to respective charId</returns>
        public OleDbDataReader readUniqueSGCharId(string filePath, int charId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUniqueSGCharId = new OleDbCommand("Select distinct(SGCharID) from SGDATA where CharID=" + charId, connection);
                cmdUniqueSGCharId.Transaction = transaction;
                reader = cmdUniqueSGCharId.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Main Screen : Update SGDATA (i.e insert SGNO enteries)
        /// <summary>
        /// Updating SGDATA by setting SGNo to respective SGCharId and charId in SGDATA table
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="charId">CharId</param>
        /// <param name="SGChatId">SGCharId</param>
        /// <param name="SGNo">SGNo</param>
        public void updateSGDATA(string filePath, int charId, int SGChatId, int SGNo)
        {
            try
            {
                connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
                if (connection.State == System.Data.ConnectionState.Closed)
                {
                    connection.Open();
                }
                OleDbTransaction transaction;
                //start transaction
                transaction = connection.BeginTransaction();
                try
                {
                    OleDbCommand cmdUpdateSGData = new OleDbCommand("Update SGDATA set SGNO=" + SGNo + " where CharID=" + charId + "and SGCharID=" + SGChatId, connection);
                    cmdUpdateSGData.Transaction = transaction;
                    cmdUpdateSGData.ExecuteNonQuery();
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception : " + ex.Message);
                    // Attempt to roll back the transaction.
                    try
                    {
                        transaction.Rollback();
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred
                        // on the server that would cause the rollback to fail, such as
                        // a closed connection.
                        MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                        MessageBox.Show("Message:" + ex2.Message);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception : "+e.ToString());
            }
        }
        #endregion

        #region Main screen:Read values from SGDATA
        /// <summary>
        /// Read all values from SGDATA for respective charId,SGNo and rdgNo
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="charId">CharId</param>
        /// <param name="SGNo">SGNo</param>
        /// <param name="rdgNo">RdgNo</param>
        /// <returns></returns>
        public OleDbDataReader readValues(string filePath, int charId, int SGNo, int rdgNo)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();

            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadValues = new OleDbCommand("Select * from SGDATA where CharID=" + charId + " and SGNO=" + SGNo + " and RdgNo=" + rdgNo, connection);
                cmdReadValues.Transaction = transaction;
                reader = cmdReadValues.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readValSgnoWiseFromSgdata(string filePath, int charId, int SGNo)
        {
            //connection = new OleDbConnection(conString + filePath);
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadValues = new OleDbCommand("Select value from SGDATA where CharID=" + charId + " and SGNO=" + SGNo, connection);
                cmdReadValues.Transaction = transaction;
                reader = cmdReadValues.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        #endregion

        #region Main Screen : Read unique SGNo from SGDATA
        /// <summary>
        /// Read distinct SGNo to corresponding charId from SGDATA table
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="charId">charId</param>
        /// <returns>All SGNo corresponding to specified CharId</returns>
        public OleDbDataReader readSGNo(string filePath, int charId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadSGNo = new OleDbCommand("Select distinct(SGNO) from SGDATA where CharID=" + charId, connection);
                cmdReadSGNo.Transaction = transaction;
                reader = cmdReadSGNo.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Main Screen : Update Value in SGData
        /// <summary>
        /// Update Value in SGDATA table for given charId,SGNo,RdgNo
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="CharId">CharId whosw value is to be updated</param>
        /// <param name="SGNo">SGNo whosw value is to be updated</param>
        /// <param name="rdgNo">RdgNo whosw value is to be updated</param>
        /// <param name="value">Update Value</param>
        public void UpdateValue(string filePath, int CharId, int SGNo, int rdgNo, double? value)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateValue = new OleDbCommand("Update SGDATA set [Value]= @Value where CharID=@CharId  and RdgNo=@rdgNo and SGNO=@SGNo", connection);
                cmdUpdateValue.Transaction = transaction;
                if (value != null)
                    cmdUpdateValue.Parameters.Add("@Value", OleDbType.VarChar).Value = value;
                else
                    cmdUpdateValue.Parameters.Add("@Value", OleDbType.VarChar).Value = Double.NaN;
                cmdUpdateValue.Parameters.Add("@CharId", OleDbType.Integer).Value = CharId;
                cmdUpdateValue.Parameters.Add("@rdgNo", OleDbType.Integer).Value = rdgNo;
                cmdUpdateValue.Parameters.Add("@SGNo", OleDbType.Integer).Value = SGNo;

                cmdUpdateValue.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Main Screen : Delete Value in SGData
        /// <summary>
        /// Delete entry from SGDATA table for given combination of charId,SGNo,RdgNo
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="CharId">CharId</param>
        /// <param name="SGNo">SGNo</param>
        /// <param name="rdgNo">RDGNo</param>
        public void DeleteValue(string filePath, int CharId, int SGNo, int rdgNo)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteValue = new OleDbCommand("Delete from SGDATA where CharID=" + CharId + " and RdgNo=" + rdgNo + " and SGNO=" + SGNo, connection);
                cmdDeleteValue.Transaction = transaction;
                cmdDeleteValue.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Main Screen : Insert  Value in SGData
        /// <summary>
        /// Insert into SGDATA table
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="CharId">CharId</param>
        /// <param name="SGNo">SGNo</param>
        /// <param name="rdgNo">RdgNo</param>
        /// <param name="value">Value(may be null)</param>
        public void InsertValueIntoSGData(string filePath, int CharId,int SGCharID ,int SGNo, int rdgNo, double? value)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            //if (value == null)
            //{
            //    OleDbCommand cmdInsertNullValue = new OleDbCommand("Insert into SGDATA(CharID,RdgNo,SGNO) values(@CharId,@rdgNo,@SGNo)", connection);
            //    cmdInsertNullValue.Parameters.Add("@CharId", OleDbType.Integer).Value = CharId;
            //    cmdInsertNullValue.Parameters.Add("@rdgNo", OleDbType.Integer).Value = rdgNo;
            //    cmdInsertNullValue.Parameters.Add("@SGNo", OleDbType.Integer).Value = SGNo;
            //    cmdInsertNullValue.ExecuteNonQuery();
            //}
            //else

            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdInsertNullValue = new OleDbCommand("Insert into SGDATA(CharID,SGCharID,RdgNo,SGNO,[Value]) values(@CharId,@SGCharID,@rdgNo,@SGNo,@value)", connection);
                cmdInsertNullValue.Transaction = transaction;
                cmdInsertNullValue.Parameters.Add("@CharId", OleDbType.Integer).Value = CharId;
                cmdInsertNullValue.Parameters.Add("@SGCharID", OleDbType.Integer).Value = SGCharID;
                cmdInsertNullValue.Parameters.Add("@rdgNo", OleDbType.Integer).Value = rdgNo;
                cmdInsertNullValue.Parameters.Add("@SGNo", OleDbType.Integer).Value = SGNo;
                if (value != null)
                    cmdInsertNullValue.Parameters.Add("@value", OleDbType.Double).Value = value;
                else
                    cmdInsertNullValue.Parameters.Add("@value", OleDbType.Double).Value = Double.NaN;
                cmdInsertNullValue.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Main Screen : Read all rdgNo
        /// <summary>
        /// Select distinct RdgNo from SGDATA table
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <returns>distinct RdgNo</returns>
        public OleDbDataReader readRdgNo(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
            {
                connection.Open();
            }
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadRdgNo = new OleDbCommand("Select distinct(RdgNo) from SGDATA", connection);
                cmdReadRdgNo.Transaction = transaction;
                reader = cmdReadRdgNo.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Main Screen : Drop SGNo feild from SGDATA
        /// <summary>
        /// Drop column SGNo from SGDATA
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        public void dropColumnSGNo(string filePath)
        {
            try
            {
                connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
                if (connection.State == System.Data.ConnectionState.Closed)
                    connection.Open();
                OleDbTransaction transaction;
                //start transaction
                transaction = connection.BeginTransaction();
                try
                {
                    OleDbCommand cmdDropColumnSGNo = new OleDbCommand("ALTER TABLE SGDATA DROP COLUMN SGNO", connection);
                    cmdDropColumnSGNo.Transaction = transaction;
                    cmdDropColumnSGNo.ExecuteNonQuery();
                    transaction.Commit();
                    //connection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception : " + ex.Message);
                    // Attempt to roll back the transaction.
                    try
                    {
                        transaction.Rollback();
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred
                        // on the server that would cause the rollback to fail, such as
                        // a closed connection.
                        MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                        MessageBox.Show("Message:" + ex2.Message);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception : "+e.ToString());
            }
        }
        #endregion

        #endregion

        #region Table :SGStat

        #region Main screen: read SGNO
        /// <summary>
        /// Select SGNo for given CharId and SGCharId from SGStat table
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="charId">CharId</param>
        /// <param name="SGChatId">SGCharId</param>
        /// <returns>Reader containing SGNo</returns>
        public OleDbDataReader readSGNO(string filePath, int charId, int SGChatId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadSGNO = new OleDbCommand("Select SGNO from SGStat where CharID=" + charId + " and SGCharID=" + SGChatId, connection);
                cmdReadSGNO.Transaction = transaction;
                reader = cmdReadSGNO.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Main Screen : Read Max SGNo
        /// <summary>
        /// Select max SGNo from SGStat table 
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <returns>max SGNo</returns>
        public int ReadMaxSGNo(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();

            OleDbCommand cmdReadMaxSGNo = new OleDbCommand("Select max(SGNO) from SGStat", connection);
            if (cmdReadMaxSGNo.ExecuteScalar() != DBNull.Value)
                return Convert.ToInt32(cmdReadMaxSGNo.ExecuteScalar());
            else
                return 0;
        }

        public int ReadMinSGNo(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();

            OleDbCommand cmdReadMinSGNo = new OleDbCommand("Select min(SGNO) from SGStat", connection);
            if (cmdReadMinSGNo.ExecuteScalar() != DBNull.Value)
                return Convert.ToInt32(cmdReadMinSGNo.ExecuteScalar());
            else
                return 0;
        }

        public int ReadMaxSGNo(string filePath, int charId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();

            OleDbCommand cmdReadMaxSGNo = new OleDbCommand("Select max(SGNO) from SGData where CharId = " + charId, connection);
            if (cmdReadMaxSGNo.ExecuteScalar() != DBNull.Value)
                return Convert.ToInt32(cmdReadMaxSGNo.ExecuteScalar());
            else
                return 0;
        }

        public int ReadMinSGNo(string filePath, int charId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();

            OleDbCommand cmdReadMinSGNo = new OleDbCommand("Select min(SGNO) from SGData where CharId = " + charId, connection);
            if (cmdReadMinSGNo.ExecuteScalar() != DBNull.Value)
                return Convert.ToInt32(cmdReadMinSGNo.ExecuteScalar());
            else
                return 0;
        }

        public OleDbDataReader ReadValueFromSGData(string filePath, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select value,SGNO from sgdata where Charid = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Main Screen : Insert into SGStat
        /// <summary>
        /// Insert into SGStat table 
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="CharId">CharId</param>
        /// <param name="SGNo">SGNo</param>
        public void InsertIntoSGStat(string filePath, int CharId, int SGNo)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdInsertIntoSGStat = new OleDbCommand("Insert into SGStat(CharID,SGNO) values(@CharId,@SGNo)", connection);
                cmdInsertIntoSGStat.Transaction = transaction;
                cmdInsertIntoSGStat.Parameters.Add("@CharId", OleDbType.Integer).Value = CharId;
                cmdInsertIntoSGStat.Parameters.Add("@SGNo", OleDbType.Integer).Value = SGNo;
                cmdInsertIntoSGStat.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #endregion

        #region Table : Trace

        #region Main screen :Read trace details
        /// <summary>
        /// Read trace category and value from Trace Table for given traceId
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="traceTagId">traceId</param>
        /// <returns>OleDbDataReader</returns>
        public OleDbDataReader readTraceDetls(string filePath, int traceTagId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadTraceDetls = new OleDbCommand("Select TraceCat,Tracevalue from Trace where TraceID=" + traceTagId, connection);
                cmdReadTraceDetls.Transaction = transaction;
                reader = cmdReadTraceDetls.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Main screen :Read All trace details
        /// <summary>
        /// Read trace category and value from Trace Table
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <returns>OleDbDataReader</returns>
        public OleDbDataReader readTraceDetls(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadTraceDetls = new OleDbCommand("Select TraceCat,TraceType,Tracevalue from Trace", connection);
                cmdReadTraceDetls.Transaction = transaction;
                reader = cmdReadTraceDetls.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Event_Trace Screen : Read distinct trace category
        /// <summary>
        /// Read distinct trace category from Trace table
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <returns>OleDbDataReader</returns>
        public OleDbDataReader readDistinctTraceCategory(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdreadDistinctTraceCategory = new OleDbCommand("Select TraceCat from Trace group by TraceCat order by min(TraceID)", connection);
                cmdreadDistinctTraceCategory.Transaction = transaction;
                reader = cmdreadDistinctTraceCategory.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Event_Trace Screen : Read trace value for given trace category
        /// <summary>
        /// Read traceId and value for specified trace category
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="category">category</param>
        /// <returns>OleDbDataReader</returns>
        public OleDbDataReader readTraceValue(string filePath, string category)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdreadCatValue = new OleDbCommand("Select Tracevalue,TraceID from Trace where TraceCat=@category ", connection);
                cmdreadCatValue.Transaction = transaction;
                cmdreadCatValue.Parameters.Add("@category", OleDbType.VarChar).Value = category;
                reader = cmdreadCatValue.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Event_Trace Screen : Update Trace Value
        /// <summary>
        /// Update trace value in Trace table for specified traceId
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="traceValue">trace value</param>
        /// <param name="traceId">trace Id</param>
        public void updateTraceValue(string filePath, string traceValue, int traceId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateTraceCategoryValue = new OleDbCommand("Update Trace set Tracevalue=@traceValue where TraceID=@traceId", connection);
                cmdUpdateTraceCategoryValue.Transaction = transaction;
                cmdUpdateTraceCategoryValue.Parameters.Add("@traceValue", OleDbType.VarChar).Value = traceValue;
                cmdUpdateTraceCategoryValue.Parameters.Add("@traceId", OleDbType.Integer).Value = traceId;
                cmdUpdateTraceCategoryValue.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Event_Trace Screen : Insert trace category and value
        /// <summary>
        /// Insert trace category and value in trace table
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="traceValue">trace value</param>
        /// <param name="traceCategory">trace category</param>
        public void insertTraceCategoryValue(string filePath, string traceValue, string traceCategory,string traceType)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdInsertTraceCategoryValue = new OleDbCommand("Insert into Trace(TraceCat,TraceType,Tracevalue) values(@traceCategory,@traceType,@traceValue)", connection);
                cmdInsertTraceCategoryValue.Transaction = transaction;
                cmdInsertTraceCategoryValue.Parameters.Add("@traceCategory", OleDbType.VarChar).Value = traceCategory;
                cmdInsertTraceCategoryValue.Parameters.Add("@traceType", OleDbType.VarChar).Value = traceType;
                cmdInsertTraceCategoryValue.Parameters.Add("@traceValue", OleDbType.VarChar).Value = traceValue;
                cmdInsertTraceCategoryValue.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void deleteFromTraceTable(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteAll = new OleDbCommand("Delete from Trace", connection);
                cmdDeleteAll.Transaction = transaction;
                cmdDeleteAll.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Event_Trace Screen : Delete trace value as per traceId
        /// <summary>
        /// Delete entry from Trace table specified trace Id
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="traceId">Trace Id</param>
        public void deleteTraceCategoryValue(string filePath, int traceId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteTraceCategoryValue = new OleDbCommand("Delete from Trace where TraceID=" + traceId, connection);
                cmdDeleteTraceCategoryValue.Transaction = transaction;
                cmdDeleteTraceCategoryValue.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region addNewEvent_TraceCategory : Update Category for trace
        /// <summary>
        /// Update Trace table with new trace category for specified old trace category 
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="oldCategory">trace category to be updated</param>
        /// <param name="newCategory">trace category to be updated with</param>
        public void updateTraceCategory(string filePath, string oldCategory, string newCategory)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateTraceCategory = new OleDbCommand("Update Trace set TraceCat=@newCategory where TraceCat=@oldCategory", connection);
                cmdUpdateTraceCategory.Transaction = transaction;
                cmdUpdateTraceCategory.Parameters.Add("@newCategory", OleDbType.VarChar).Value = newCategory;
                cmdUpdateTraceCategory.Parameters.Add("@oldCategory", OleDbType.VarChar).Value = oldCategory;
                cmdUpdateTraceCategory.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Main Screen : Get TraceId using TraceCategory and value
        /// <summary>
        /// Get traceId for specified trace category and value
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="traceCategory">Trace category</param>
        /// <param name="traceValue">Trace Value</param>
        /// <returns>TraceId</returns>
        public int getTraceId(string filePath, string traceCategory, string traceValue)
        {
            int id = 0;

            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                //Prepared stmt here
                OleDbCommand cmdGetTraceId = new OleDbCommand("Select TraceID from Trace where TraceCat=@traceCategory and Tracevalue=@traceValue", connection);
                cmdGetTraceId.Transaction = transaction;
                cmdGetTraceId.Parameters.Add("@traceCategory", OleDbType.VarChar).Value = traceCategory;
                cmdGetTraceId.Parameters.Add("@traceValue", OleDbType.VarChar).Value = traceValue;
                reader = cmdGetTraceId.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            if (reader.Read())
            {
                id = Convert.ToInt32(reader["TraceID"]);
            }
            return id;
        }
        #endregion

        #endregion

        #region Table : SGTrace

        #region Main Screen: Get TraceID from SgTrace
        /// <summary>
        /// Get Trac Id from SGTrace table for specified SGNo
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="SGNo">SGNo</param>
        /// <returns></returns>
        public OleDbDataReader readTraceId(string filePath, int SGNo)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadTraceId = new OleDbCommand("Select TraceID from SgTrace where SgNo=" + SGNo, connection);
                cmdReadTraceId.Transaction = transaction;
                reader = cmdReadTraceId.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Event_Trace Screen : Check if given trace id is assigned to any SGNo
        /// <summary>
        /// Identifies if specified TraceId is assigned to any SGNo
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="traceId">Trace Id</param>
        /// <returns>bool</returns>
        public bool checkIfTraceIdUsed(string filePath, int traceId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdCheckIfTraceIdUsed = new OleDbCommand("Select SgNo from SgTrace where TraceID=" + traceId, connection);
                cmdCheckIfTraceIdUsed.Transaction = transaction;
                reader = cmdCheckIfTraceIdUsed.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            if (reader.Read())
            {
                CloseConnection();
                return true;
            }
            else
            {
                CloseConnection();
                return false;
            }
        }
        #endregion

        #region Event_Trace Screen : Delete from SGTrace as per TraceId
        /// <summary>
        /// Delete entry from SGTrace for specified TraceId
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="traceId">TraceId</param>
        public void deleteFromSGTrace(string filePath, int traceId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteFromSGTrace = new OleDbCommand("Delete from SgTrace where TraceID=" + traceId, connection);
                cmdDeleteFromSGTrace.Transaction = transaction;
                cmdDeleteFromSGTrace.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Main Screen :Check if Trace Id and SGNo combination exists
        /// <summary>
        /// Check if Trace Id and SGNo combination exists
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="traceId">Trace Id</param>
        /// <param name="SGNo">SGNo</param>
        /// <returns>bool</returns>
        public bool checkForTraceId_SGNoCombination(string filePath, int traceId, int SGNo)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
            {
                connection.Open();
            }
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdCheckForTraceId_SGNoCombination = new OleDbCommand("Select * from SgTrace where SgNo=" + SGNo + " and TraceID=" + traceId, connection);
                cmdCheckForTraceId_SGNoCombination.Transaction = transaction;
                reader = cmdCheckForTraceId_SGNoCombination.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            if (reader.Read())
            {
                //CloseConnection();
                return true;
            }
            else
            {
                //connection.Close();
                //connection.Dispose();
                return false;
            }
        }
        #endregion

        #region MainScreen : Insert into SGTrace
        /// <summary>
        /// Insert into SGTrace
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="SGno">SGNo</param>
        /// <param name="TraceId">Trace Id</param>
        public void insertSGno_TraceId(string filePath, int SGno, int TraceId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdInsertSGno_TraceId = new OleDbCommand("Insert into SgTrace values(@SGno,@traceId)", connection);
                cmdInsertSGno_TraceId.Transaction = transaction;
                cmdInsertSGno_TraceId.Parameters.Add("@SGno", OleDbType.Integer).Value = SGno;
                cmdInsertSGno_TraceId.Parameters.Add("@traceId", OleDbType.Integer).Value = TraceId;
                cmdInsertSGno_TraceId.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Main Screen : Delete from SGTrace with given SGNo and TraceId
        /// <summary>
        /// Delete enrty from SGTrace for specified SGNo and TraceId
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="SGNo">SGNo</param>
        /// <param name="TraceId">Trace Id</param>
        public void deleteSGNo_TraceIdFromSGTrace(string filePath, int SGNo, int TraceId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteSGNo_TraceId = new OleDbCommand("Delete from SgTrace where SgNo=" + SGNo + " and TraceID=" + TraceId, connection);
                cmdDeleteSGNo_TraceId.Transaction = transaction;
                cmdDeleteSGNo_TraceId.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Main Screen : Update TraceId in SGTrace for given SGNo and earlier TraceId
        /// <summary>
        /// Update TraceId feild of SGTrace for specified SGNo and earlier TraceId
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="SGNo">SGNo</param>
        /// <param name="newtraceId">TraceId to be updated with</param>
        /// <param name="oldTraceId">Trace Id to be updated</param>
        public void updateTraceId_SGTrace(string filePath, int SGNo, int newtraceId, int oldTraceId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateTraceId_SGTrace = new OleDbCommand("Update SgTrace set TraceID=@newtraceId where SgNo=@SGNo and TraceID=@oldTraceId", connection);
                cmdUpdateTraceId_SGTrace.Transaction = transaction;
                cmdUpdateTraceId_SGTrace.Parameters.Add("@newtraceId", OleDbType.Integer).Value = newtraceId;
                cmdUpdateTraceId_SGTrace.Parameters.Add("@SGNo", OleDbType.Integer).Value = SGNo;
                cmdUpdateTraceId_SGTrace.Parameters.Add("@oldTraceId", OleDbType.Integer).Value = oldTraceId;
                cmdUpdateTraceId_SGTrace.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion


        #endregion

        #region Table : Event

        #region Main screen :Read all event details
        /// <summary>
        /// Read Event category,value from Event table
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <returns>OleDbDataReader</returns>
        public OleDbDataReader readEventDetls(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadEventDetls = new OleDbCommand("Select EventCat,EventValue,Eventtype from Event", connection);
                cmdReadEventDetls.Transaction = transaction;
                reader = cmdReadEventDetls.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Main screen :Read event details
        /// <summary>
        /// Read Event category,value from Event table for specified event Id
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="eventId">Event Id</param>
        /// <returns>OleDbDataReader</returns>
        public OleDbDataReader readEventDetls(string filePath, int eventId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadEventDetls = new OleDbCommand("Select EventCat,EventValue from Event where EventId=" + eventId, connection);
                cmdReadEventDetls.Transaction = transaction;
                reader = cmdReadEventDetls.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Event_Trace Screen : Read distinct Event Category
        /// <summary>
        /// read distinct event category
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <returns>OleDbDataReader</returns>
        public OleDbDataReader readDistinctEventCategory(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();

            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdreadDistinctEventCategory = new OleDbCommand("Select EventCat from Event group by EventCat order by min(EventID)", connection);
                cmdreadDistinctEventCategory.Transaction = transaction;
                reader = cmdreadDistinctEventCategory.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Event_Trace Screen read event value for given event category
        /// <summary>
        /// Read event ID,value from Event table for specified cartegory
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="category">category</param>
        /// <returns>OledbDataReader</returns>
        public OleDbDataReader readEventValue(string filePath, string category)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadEventValue = new OleDbCommand("Select EventId,EventValue from Event where EventCat=@category", connection);
                cmdReadEventValue.Transaction = transaction;
                cmdReadEventValue.Parameters.Add("@category", OleDbType.VarChar).Value = category;
                reader = cmdReadEventValue.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Event_Trace Screen : Update Event Value
        /// <summary>
        /// Update Event Value in Event table for specified Event id 
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="eventValue">Event Value</param>
        /// <param name="eventId">Event Id</param>
        public void updateEventValue(string filePath, string eventValue, int eventId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateEventValue = new OleDbCommand("Update Event set EventValue=@eventValue where EventId=@eventId", connection);
                cmdUpdateEventValue.Transaction = transaction;
                cmdUpdateEventValue.Parameters.Add("@eventValue", OleDbType.VarChar).Value = eventValue;
                cmdUpdateEventValue.Parameters.Add("@eventId", OleDbType.Integer).Value = eventId;
                cmdUpdateEventValue.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Event_Trace Screen : Insert Event category and Value
        /// <summary>
        /// Insert into Event table 
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="eventValue">Event Value</param>
        /// <param name="eventCategory">Event Category</param>
        public void insertEventCategoryValue(string filePath, string eventValue, string eventCategory,string eventType)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdInsertEventCategoryValue = new OleDbCommand("Insert into Event(EventCat,EventValue,Eventtype) values(@eventCategory,@eventValue,@eventType)", connection);
                cmdInsertEventCategoryValue.Transaction = transaction;
                cmdInsertEventCategoryValue.Parameters.Add("@eventCategory", OleDbType.VarChar).Value = eventCategory;
                cmdInsertEventCategoryValue.Parameters.Add("@eventValue", OleDbType.VarChar).Value = eventValue;
                cmdInsertEventCategoryValue.Parameters.Add("@eventType", OleDbType.VarChar).Value = eventType;
                cmdInsertEventCategoryValue.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void deleteFromEvent(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteAll = new OleDbCommand("Delete from Event", connection);
                cmdDeleteAll.Transaction = transaction;
                cmdDeleteAll.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Event_Trace screen : Delete event value as per eventId
        /// <summary>
        /// Delete from Event table for specified event Id
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="eventId">Event Id</param>
        public void deleteEventCategoryValue(string filePath, int eventId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteEventCategoryValue = new OleDbCommand("Delete from Event where EventID=" + eventId, connection);
                cmdDeleteEventCategoryValue.Transaction = transaction;
                cmdDeleteEventCategoryValue.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region addNewEvent_TraceCategory : Update Category for event
        /// <summary>
        /// Update Event Category in Event Table for specified Event Id and earlier Event Category
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="oldCategory">Event category to be updated</param>
        /// <param name="newCategory">Event Category to update with</param>
        public void updateEventCategory(string filePath, string oldCategory, string newCategory)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdupdateEventCategory = new OleDbCommand("Update Event set EventCat=@newCategory where EventCat=@oldCategory", connection);
                cmdupdateEventCategory.Transaction = transaction;
                cmdupdateEventCategory.Parameters.Add("@newCategory", OleDbType.VarChar).Value = newCategory;
                cmdupdateEventCategory.Parameters.Add("@oldCategory", OleDbType.VarChar).Value = oldCategory;
                cmdupdateEventCategory.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region MainScreen : Get EventId using Event Category and value
        /// <summary>
        /// Get EventId for specified Event category and value
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="traceCategory">Event category</param>
        /// <param name="traceValue">Event Value</param>
        /// <returns>Event Id</returns>
        public int getEventId(string filePath, string eventCategory, string eventValue)
        {
            int id = 0;

            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdGetEventId = new OleDbCommand("Select EventId from Event where EventCat=@eventCategory and EventValue=@eventValue", connection);
                cmdGetEventId.Transaction = transaction;
                cmdGetEventId.Parameters.Add("@eventCategory", OleDbType.VarChar).Value = eventCategory;
                cmdGetEventId.Parameters.Add("@eventValue", OleDbType.VarChar).Value = eventValue;

                reader = cmdGetEventId.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            if (reader.Read())
            {
                id = Convert.ToInt32(reader["EventId"]);
            }
            return id;
        }
        #endregion

        #endregion

        #region Table : SGEvent

        #region Event_Trace Screen : Check if given event id is assigned to any SGNo
        /// <summary>
        /// Check if event id is assigned for specified SGNo
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="eventId">Event Id</param>
        /// <returns>bool</returns>
        public bool checkIfEventIdUsed(string filePath, int eventId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdCheckIfEventIdUsed = new OleDbCommand("Select SgNo from SgEvent where EventID=" + eventId, connection);
                cmdCheckIfEventIdUsed.Transaction = transaction;
                reader = cmdCheckIfEventIdUsed.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            if (reader.Read())
            {
                //connection.Close();
                return true;
            }
            else
            {
                //connection.Close();
                return false;
            }
        }
        #endregion

        #region Main Screen: Get EventID from SgEvent
        /// <summary>
        /// Read event id for specified SGNo
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="SGNo">SGNo</param>
        /// <returns>OleDbDataReader</returns>
        public OleDbDataReader readEventId(string filePath, int SGNo)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadEventId = new OleDbCommand("Select EventID from SgEvent where SgNo=" + SGNo, connection);
                cmdReadEventId.Transaction = transaction;
                reader = cmdReadEventId.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Event_Trace Screen : Delete from SGEvent as per EventId
        /// <summary>
        /// Delete from SgEvent for specified event Id 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="eventId"></param>
        public void deleteFromSGEvent(string filePath, int eventId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteFromSGEvent = new OleDbCommand("Delete from SgEvent where EventID=" + eventId, connection);
                cmdDeleteFromSGEvent.Transaction = transaction;
                cmdDeleteFromSGEvent.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Main Screen :Check if Event Id and SGNo combination exists
        /// <summary>
        /// Check if Event Id and SGNo combination exists
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="eventId">Event Id</param>
        /// <param name="SGNo">SGno</param>
        /// <returns>bool</returns>
        public bool checkForEventId_SGNoCombination(string filePath, int eventId, int SGNo)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdCheckForEventId_SGNoCombination = new OleDbCommand("Select * from SgEvent where SgNo=" + SGNo + " and EventID=" + eventId, connection);
                cmdCheckForEventId_SGNoCombination.Transaction = transaction;
                reader = cmdCheckForEventId_SGNoCombination.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            if (reader.Read())
            {
                //CloseConnection();
                return true;
            }
            else
            {
                //connection.Close();
                return false;
            }
        }
        #endregion

        #region  Main screen : Delete from SGEvent with given SGNo and EventId
        /// <summary>
        /// Delete enrty from SGEvent for specified SGNo and EventId
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="SGNo">SGNo</param>
        /// <param name="EventId">EventId</param>

        public void deleteSGNo_EventId_FromSGEvent(string filePath, int SGNo, int eventId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteSGNo_EventId = new OleDbCommand("Delete from SgEvent where SgNo=" + SGNo + " and EventID=" + eventId, connection);
                cmdDeleteSGNo_EventId.Transaction = transaction;
                cmdDeleteSGNo_EventId.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Main Screen : Insert into SGEvent
        public void InsertSGNo_EventId(string filePath, int SGNo, int EventId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdInsertSGNo_EventId = new OleDbCommand("Insert into SgEvent values(@SGNo,@EventId)", connection);
                cmdInsertSGNo_EventId.Transaction = transaction;
                cmdInsertSGNo_EventId.Parameters.Add("@SGNo", OleDbType.Integer).Value = SGNo;
                cmdInsertSGNo_EventId.Parameters.Add("@EventId", OleDbType.Integer).Value = EventId;
                cmdInsertSGNo_EventId.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Main Screen : Update EventId for given SGNo and and earlier EventId
        /// <summary>
        /// Update EventId feild of SGEvent for specified SGNo and earlier EventId
        /// </summary>
        /// <param name="filePath">file name which is to be read</param>
        /// <param name="SGNo">SGNo</param>
        /// <param name="newEventId">Event to be updated with</param>
        /// <param name="oldEventId">Event Id to be updated</param>
        public void updateEventId_InSGEvent(string filePath, int SGNo, int newEventId, int oldEventId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateEventId_InSGEvent = new OleDbCommand("Update SgEvent set EventID=" + newEventId + " where SgNo=" + SGNo + " and EventID=" + oldEventId, connection);
                cmdUpdateEventId_InSGEvent.Transaction = transaction;
                cmdUpdateEventId_InSGEvent.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        #region Get sgno from sgevent table
        public OleDbDataReader GetSgnoFromSgEvent(string filePath,int fromsgno,int tosgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdSgno = new OleDbCommand("Select SgNo from SgEvent where SgNo between " + fromsgno + " and " + tosgno, connection);
                cmdSgno.Transaction = transaction;
                reader = cmdSgno.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader GetSgnoFromSgEvent(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdSgno = new OleDbCommand("Select SgNo from SgEvent", connection);
                cmdSgno.Transaction = transaction;
                reader = cmdSgno.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #endregion

        /*For Attribute data*/

        public OleDbDataReader readNCId_Values(string filePath, int CharId, int SGNo)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdreadNCId_Values = new OleDbCommand("Select NCId,[value] from NcParato where CharID=@CharId and Sgno=@SGNo", connection);
                cmdreadNCId_Values.Transaction = transaction;
                cmdreadNCId_Values.Parameters.Add("@CharId", OleDbType.Integer).Value = CharId;
                cmdreadNCId_Values.Parameters.Add("@SGNo", OleDbType.Integer).Value = SGNo;
                transaction.Commit();
                reader = cmdreadNCId_Values.ExecuteReader();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readNCId_Values(string filePath, int CharId, int SGNo, int NCId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdreadNCId_Values = new OleDbCommand("Select NCId,[value] from NcParato where CharID=@CharId and Sgno=@SGNo and NCId=@NCId", connection);
                cmdreadNCId_Values.Transaction = transaction;
                cmdreadNCId_Values.Parameters.Add("@CharId", OleDbType.Integer).Value = CharId;
                cmdreadNCId_Values.Parameters.Add("@SGNo", OleDbType.Integer).Value = SGNo;
                cmdreadNCId_Values.Parameters.Add("@NCId", OleDbType.Integer).Value = NCId;
                transaction.Commit();
                reader = cmdreadNCId_Values.ExecuteReader();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readNCCat(string filePath, int NCId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdreadNCCat = new OleDbCommand("Select NCCat from NC where NCID=@NCId", connection);
                cmdreadNCCat.Transaction = transaction;
                cmdreadNCCat.Parameters.Add("@NCId", OleDbType.Integer).Value = NCId;
                transaction.Commit();
                reader = cmdreadNCCat.ExecuteReader();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readNCCat_CharId(string filePath, int CharId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdreadNCCat = new OleDbCommand("Select NCCat,NCID from NC where CharID=@CharID", connection);
                cmdreadNCCat.Transaction = transaction;
                cmdreadNCCat.Parameters.Add("@CharID", OleDbType.Integer).Value = CharId;
                transaction.Commit();
                reader = cmdreadNCCat.ExecuteReader();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public void update_Attribute_Value(string filePath, int CharId, int NCId, int SGNo, int Value)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdate_Attribute_Value = new OleDbCommand("Update NcParato set [value]=@Value where CharID=@CharId and Sgno=@SGNo and NCId=@NCId", connection);
                cmdUpdate_Attribute_Value.Transaction = transaction;
                cmdUpdate_Attribute_Value.Parameters.Add("@Value", OleDbType.Integer).Value = Value;
                cmdUpdate_Attribute_Value.Parameters.Add("@CharId", OleDbType.Integer).Value = CharId;
                cmdUpdate_Attribute_Value.Parameters.Add("@SGNo", OleDbType.Integer).Value = SGNo;
                cmdUpdate_Attribute_Value.Parameters.Add("@NCId", OleDbType.Integer).Value = NCId;
                cmdUpdate_Attribute_Value.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void update_Attribute_Total(string filePath, int CharId, int SGNo, int rdgNo, int value)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdate_Attribute_Total = new OleDbCommand("Update SGDATA set [Value]=@Value where CharID=@CharId and RdgNo=@rdgNo and SGNO=@SGNo", connection);
                cmdUpdate_Attribute_Total.Transaction = transaction;
                cmdUpdate_Attribute_Total.Parameters.Add("@Value", OleDbType.Integer).Value = value;
                cmdUpdate_Attribute_Total.Parameters.Add("@CharId", OleDbType.Integer).Value = CharId;
                cmdUpdate_Attribute_Total.Parameters.Add("@rdgNo", OleDbType.Integer).Value = rdgNo;
                cmdUpdate_Attribute_Total.Parameters.Add("@SGNo", OleDbType.Integer).Value = SGNo;
                cmdUpdate_Attribute_Total.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public OleDbDataReader readSPCFileType(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadSPCFileType = new OleDbCommand("Select CharType,ChartType from Characterstic", connection);
                cmdReadSPCFileType.Transaction = transaction;
                reader = cmdReadSPCFileType.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                throw ex;
                //MessageBox.Show("Exception : " + ex.Message,"SPC WorkBench");
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public void deleteFromSGData(string filePath, int SGNO)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteFromSGData = new OleDbCommand("Delete from SGDATA where SGNO=" + SGNO, connection);
                cmdDeleteFromSGData.Transaction = transaction;
                cmdDeleteFromSGData.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void deleteFromSGHeader(string filePath, int SGNO)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteFromSGHeader = new OleDbCommand("Delete from SGHeader where SGNO=" + SGNO, connection);
                cmdDeleteFromSGHeader.Transaction = transaction;
                cmdDeleteFromSGHeader.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void deleteFromSGStat(string filePath, int SGNO)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteFromSGStat = new OleDbCommand("Delete from SGStat where SGNO=" + SGNO, connection);
                cmdDeleteFromSGStat.Transaction = transaction;
                cmdDeleteFromSGStat.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void deleteFromSGEventSGNO(string filePath, int SGNO)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteFromSGEvent = new OleDbCommand("Delete from SgEvent where SgNo=" + SGNO, connection);
                cmdDeleteFromSGEvent.Transaction = transaction;
                cmdDeleteFromSGEvent.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void deleteFromSGTraceSGNO(string filePath, int SGNO)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteFromSGTrace = new OleDbCommand("Delete from SgTrace where SgNo=" + SGNO, connection);
                cmdDeleteFromSGTrace.Transaction = transaction;
                cmdDeleteFromSGTrace.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public int MaxSgno(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbCommand cmdReadMaxSGNo = new OleDbCommand("Select max(SGNO) from SGStat", connection);
            if (cmdReadMaxSGNo.ExecuteScalar() != DBNull.Value)
                return Convert.ToInt32(cmdReadMaxSGNo.ExecuteScalar());
            else
                return 0;
        }

        //Select Charid and Sgno from sgstat of excluded subgroup
        public OleDbDataReader getCharIdNSgnoFromSgstat(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadCharIdNSgno = new OleDbCommand("Select SGNO,CharID from SGStat where Exclude = 'y'", connection);
                cmdReadCharIdNSgno.Transaction = transaction;
                reader = cmdReadCharIdNSgno.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader getCharIdNSgnoFromSgstat(string filePath ,int fromsgno,int tosgno)
        {
            connection = new OleDbConnection(conString + filePath);//+ ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadCharIdNSgno = new OleDbCommand("Select SGNO,CharID from SGStat where Exclude = 'y' and SGNO between " + fromsgno + " and " + tosgno, connection);
                cmdReadCharIdNSgno.Transaction = transaction;
                reader = cmdReadCharIdNSgno.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public void updateExcludeFromSgstat(string filePath, int sgno, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateExclude = new OleDbCommand("Update SGStat set Exclude = 'Y' where SGNO = " + sgno + " and CharID = " + charid, connection);
                cmdUpdateExclude.Transaction = transaction;
                cmdUpdateExclude.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void updateIncludeIntoSgstat(string filePath, int sgno, int charid)
        {
            //connection = new OleDbConnection(conString + filePath);
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateInclude = new OleDbCommand("Update SGStat set Exclude = 'N' where SGNO = " + sgno + " and CharID = " + charid, connection);
                cmdUpdateInclude.Transaction = transaction;
                cmdUpdateInclude.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public string readExcludeFromSGStat(string filePath, int sgno, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbCommand cmdreadExcludeFromSGStat = new OleDbCommand("Select Exclude from SGStat where SGNO = " + sgno + " and CharID = " + charid, connection);
            if (cmdreadExcludeFromSGStat.ExecuteScalar() != DBNull.Value)
                return Convert.ToString(cmdreadExcludeFromSGStat.ExecuteScalar());
            else
                return String.Empty;
        }

        public OleDbDataReader readTraceIdForSgno(string filePath,int sgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdreadTraceIdForSgno = new OleDbCommand("Select TraceID from SgTrace where SgNo =" + sgno, connection);
                cmdreadTraceIdForSgno.Transaction = transaction;
                reader = cmdreadTraceIdForSgno.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader ReadSGDateTime(string filePath, int sgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadSGDateTime = new OleDbCommand("Select SGDate,SgTime from SGHeader where SGNO = " + sgno, connection);
                cmdReadSGDateTime.Transaction = transaction;
                reader = cmdReadSGDateTime.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader ReadTraceValForTraceId(string filePath, int traceid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadTraceValForTraceId = new OleDbCommand("Select TraceCat,Tracevalue from Trace where TraceID = " + traceid, connection);
                cmdReadTraceValForTraceId.Transaction = transaction;
                reader = cmdReadTraceValForTraceId.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader ReadTraceValForTraceCatnId(string filePath, int StartTraceid, int EndTraceid,string Tracecategory)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadTraceValForTraceId = new OleDbCommand("Select Tracevalue from Trace where TraceID between " + StartTraceid + " and " + EndTraceid + " and TraceCat='" + Tracecategory + "'", connection);
                cmdReadTraceValForTraceId.Transaction = transaction;
                reader = cmdReadTraceValForTraceId.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader ReadEventValForEventId(string filePath, int eventid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadEventValForEventId = new OleDbCommand("Select EventCat,EventValue from Event where EventId = " + eventid, connection);
                cmdReadEventValForEventId.Transaction = transaction;
                reader = cmdReadEventValForEventId.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readEventIdForSgno(string filePath, int sgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdreadEventIdForSgno = new OleDbCommand("Select EventID from SgEvent where SgNo =" + sgno, connection);
                cmdreadEventIdForSgno.Transaction = transaction;
                reader = cmdreadEventIdForSgno.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readUCLnLCLFromSGStat(string filePath, int charid, int sgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdreadUCLnLCLFromSGStat = new OleDbCommand("Select UCLAvg,LCLAvg,MeanAVG,UCLDis,LCLDis,MeanDispersion from SGStat where CharID = " + charid + "and SGNO = " + sgno, connection);
                cmdreadUCLnLCLFromSGStat.Transaction = transaction;
                reader = cmdreadUCLnLCLFromSGStat.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        #region spctemplate.dat
        public OleDbDataReader read_DocDetails(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadDocDetails = new OleDbCommand("Select * from Documentation", connection);
                cmdReadDocDetails.Transaction = transaction;
                reader = cmdReadDocDetails.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public void update_DocDetails(string filePath,int FormatNo,string DocNO,string RevNo,string DocDate)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateDocDetails = new OleDbCommand("Update Documentation set DocNo = @DocNO,RevNo = @RevNo,DocDate = @DocDate where FormatNo = " + FormatNo + "", connection);
                cmdUpdateDocDetails.Transaction = transaction;
                if (DocNO.Equals(string.Empty) == false)
                    cmdUpdateDocDetails.Parameters.Add("@DocNO", OleDbType.VarChar).Value = DocNO;
                else
                    cmdUpdateDocDetails.Parameters.Add("@DocNO", OleDbType.VarChar).Value = String.Empty;
                if (RevNo.Equals(string.Empty) == false)
                    cmdUpdateDocDetails.Parameters.Add("@RevNo", OleDbType.VarChar).Value = RevNo;
                else
                    cmdUpdateDocDetails.Parameters.Add("@RevNo", OleDbType.VarChar).Value = String.Empty;
                if (DocDate.Equals(string.Empty) == false)
                    cmdUpdateDocDetails.Parameters.Add("@DocDate", OleDbType.VarChar).Value = DocDate;
                else
                    cmdUpdateDocDetails.Parameters.Add("@DocDate", OleDbType.VarChar).Value = String.Empty;
                cmdUpdateDocDetails.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void insert_into_NC(string filePath, string NCCat, int NCWfg,int CharID)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdinsert_into_NC = new OleDbCommand("Insert into NC(NCCat,NCWfg,CharID) values(@NCCat,@NCWfg,@CharID)", connection);
                cmdinsert_into_NC.Transaction = transaction;
                cmdinsert_into_NC.Parameters.Add("@NCCat", OleDbType.VarChar).Value = NCCat;
                cmdinsert_into_NC.Parameters.Add("@NCWfg", OleDbType.Integer).Value = NCWfg;
                cmdinsert_into_NC.Parameters.Add("@CharID", OleDbType.Integer).Value = CharID;
                cmdinsert_into_NC.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void insertintoAttributeSetting(string filePath, int CharID)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdinsert_into_NC = new OleDbCommand("Insert into AttributeSettting(,NCWfg,CharID) values(@NCCat,@NCWfg,@CharID)", connection);
                cmdinsert_into_NC.Transaction = transaction;
                cmdinsert_into_NC.Parameters.Add("@CharID", OleDbType.Integer).Value = CharID;
                cmdinsert_into_NC.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void deleteFromNC(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteAll = new OleDbCommand("Delete from NC", connection);
                cmdDeleteAll.Transaction = transaction;
                cmdDeleteAll.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public int ReadCharIdFromCharacterstic(string filePath, string CharName)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();

            OleDbCommand cmdReadCharIdFromCharacterstic = new OleDbCommand("Select CharID from Characterstic where CharName = '" +CharName+ "' ", connection);
            if (cmdReadCharIdFromCharacterstic.ExecuteScalar() != DBNull.Value)
                return Convert.ToInt32(cmdReadCharIdFromCharacterstic.ExecuteScalar());
            else
                return 0;
        }

        #endregion

        #region AttributeSetting Table

        public void UpdateProvisionalCL(string filePath, string AttVal)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateProvisionalCL = new OleDbCommand("Update AttributeSetting set AttributeValue = '" + AttVal + "' where AttributeName = 'ProvisionalCL'", connection);
                cmdUpdateProvisionalCL.Transaction = transaction;
                cmdUpdateProvisionalCL.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void UpdateScaleAll(string filePath, string AttVal)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateScaleAll = new OleDbCommand("Update AttributeSetting set AttributeValue = '" + AttVal + "' where AttributeName = 'ScaleAll'", connection);
                cmdUpdateScaleAll.Transaction = transaction;
                cmdUpdateScaleAll.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void UpdateTrensMsg(string filePath, string AttVal)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateTrensMsg = new OleDbCommand("Update AttributeSetting set AttributeValue = '" + AttVal + "' where AttributeName = 'TrensMsg'", connection);
                cmdUpdateTrensMsg.Transaction = transaction;
                cmdUpdateTrensMsg.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void UpdateIncludeData(string filePath, string AttVal)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateTrensMsg = new OleDbCommand("Update AttributeSetting set AttributeValue = '" + AttVal + "' where AttributeName = 'IncludeData'", connection);
                cmdUpdateTrensMsg.Transaction = transaction;
                cmdUpdateTrensMsg.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void UpdateIncludeCharts(string filePath, string AttVal)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateTrensMsg = new OleDbCommand("Update AttributeSetting set AttributeValue = '" + AttVal + "' where AttributeName = 'IncludeCharts'", connection);
                cmdUpdateTrensMsg.Transaction = transaction;
                cmdUpdateTrensMsg.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void UpdateIncludeResult(string filePath, string AttVal)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateTrensMsg = new OleDbCommand("Update AttributeSetting set AttributeValue = '" + AttVal + "' where AttributeName = 'IncludeResult'", connection);
                cmdUpdateTrensMsg.Transaction = transaction;
                cmdUpdateTrensMsg.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void UpdateAutoDateTime(string filePath, string AttVal)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if(connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateAutoDateTime = new OleDbCommand("Update AttributeSetting set AttributeValue = '" + AttVal + "' where AttributeName = 'AutoDateTime'", connection);
                cmdUpdateAutoDateTime.Transaction = transaction;
                cmdUpdateAutoDateTime.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void UpdateDataWindow(string filePath, string DataWindow)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateDataWindow = new OleDbCommand("Update AttributeSetting set AttributeValue = '" + DataWindow + "' where AttributeName = 'DataWindow'", connection);
                cmdUpdateDataWindow.Transaction = transaction;
                cmdUpdateDataWindow.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void UpdateDataEntryHV(string filePath, string DataEntryHV)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDataEntryHV = new OleDbCommand("Update AttributeSetting set AttributeValue = '" + DataEntryHV + "' where AttributeName = 'DataEntryHV'", connection);
                cmdDataEntryHV.Transaction = transaction;
                cmdDataEntryHV.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void UpdatePrefill(string filePath, string CharName,string prefill)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdPrefill = new OleDbCommand("Update AttributeSetting set AttributeValue = '" + prefill + "' where AttributeName = '"+CharName+"'", connection);
                cmdPrefill.Transaction = transaction;
                cmdPrefill.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public OleDbDataReader ReadFromAttributeSetting(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadFromAttributeSetting = new OleDbCommand("select * from AttributeSetting", connection);
                cmdReadFromAttributeSetting.Transaction = transaction;
                reader = cmdReadFromAttributeSetting.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }
        #endregion

        #region Delete Trace Category from trace table
        public void DeleteTraceCatFromTrace(string filePath, string traceCat)
        {
            if (connection.State == ConnectionState.Closed)
                connection.Open();
             OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteTraceVal = new OleDbCommand("Delete from Trace where TraceCat = '"+traceCat+"'", connection);
                cmdDeleteTraceVal.Transaction = transaction;
                cmdDeleteTraceVal.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void deleteFromSgTraceTraceId(string filePath, int traceId)
        {
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteTraceId = new OleDbCommand("Delete from SgTrace where TraceID = " + traceId, connection);
                cmdDeleteTraceId.Transaction = transaction;
                cmdDeleteTraceId.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        #endregion

        #region Delete Event Category from event table
        public void DeleteEventCatFromEvent(string filePath, string eventCat)
        {
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteEventVal = new OleDbCommand("Delete from Event where EventCat = '" + eventCat + "'", connection);
                cmdDeleteEventVal.Transaction = transaction;
                cmdDeleteEventVal.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void deleteFromSgEventEventId(string filePath, int eventId)
        {
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteEventId = new OleDbCommand("Delete from SgEvent where EventID = " + eventId, connection);
                cmdDeleteEventId.Transaction = transaction;
                cmdDeleteEventId.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        #endregion

        public void updateTraceTable(string filePath, string category,string tracevalue,int traceid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateTraceCategoryValue = new OleDbCommand("Update Trace set TraceCat=@category,tracevalue=@tracevalue where TraceID=@traceid", connection);
                cmdUpdateTraceCategoryValue.Transaction = transaction;
                cmdUpdateTraceCategoryValue.Parameters.Add("@category", OleDbType.VarChar).Value = category;
                cmdUpdateTraceCategoryValue.Parameters.Add("@tracevalue", OleDbType.VarChar).Value = tracevalue;
                cmdUpdateTraceCategoryValue.Parameters.Add("@traceid", OleDbType.Integer).Value = traceid;
                cmdUpdateTraceCategoryValue.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void updateEventTable(string filePath, string category, string eventvalue, int eventid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateTraceCategoryValue = new OleDbCommand("Update Event set EventCat=@category,eventvalue=@eventvalue where EventID=@eventid", connection);
                cmdUpdateTraceCategoryValue.Transaction = transaction;
                cmdUpdateTraceCategoryValue.Parameters.Add("@category", OleDbType.VarChar).Value = category;
                cmdUpdateTraceCategoryValue.Parameters.Add("@eventvalue", OleDbType.VarChar).Value = eventvalue;
                cmdUpdateTraceCategoryValue.Parameters.Add("@eventid", OleDbType.Integer).Value = eventid;
                cmdUpdateTraceCategoryValue.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        #region insert into SGHeader
        public void insertIntoSGHeaderTable(string filePath, int SGNo, DateTime? Date, DateTime? Time)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();

            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdInsertIntosgheader;
                if (Date.HasValue && Time.HasValue)
                {
                    cmdInsertIntosgheader = new OleDbCommand("Insert into SgHeader(SGNO,SGDate,SgTime) values(@SGNo,@Date,@Time)", connection);
                    cmdInsertIntosgheader.Transaction = transaction;
                    cmdInsertIntosgheader.Parameters.Add("@SGNo", OleDbType.Integer).Value = SGNo;
                    cmdInsertIntosgheader.Parameters.Add("@Date", OleDbType.DBDate).Value = Date.Value;
                    cmdInsertIntosgheader.Parameters.Add("@Time", OleDbType.DBTime).Value = Time.Value.TimeOfDay;
                }
                else
                {
                    cmdInsertIntosgheader = new OleDbCommand("Insert into SgHeader(SGNO,SGDate,SgTime) values(@SGNo,@Date,@Time)", connection);
                    cmdInsertIntosgheader.Transaction = transaction;
                    cmdInsertIntosgheader.Parameters.Add("@SGNo", OleDbType.Integer).Value = SGNo;
                    cmdInsertIntosgheader.Parameters.Add("@Date", OleDbType.DBDate).Value = DBNull.Value;
                    cmdInsertIntosgheader.Parameters.Add("@Time", OleDbType.DBTime).Value = DBNull.Value;
                }
                
                cmdInsertIntosgheader.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }
        #endregion

        public int getSGSize(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();

            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SGSize from Characterstic", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            reader.Read();
            return Convert.ToInt32(reader[0]);
        }

        public void insertIntoSGStatTable(string filePath, int SGNo, int Charid, double SGAvg, double SGDisp, double UCLAvg, double LCLAvg, double MeanAvg, double MeanDisp, string Exclude, double UCLDisp, double LCLDisp, string LimitsApply)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Insert into SGStat(SGNO,CharID,SGAVG,SGDISPERSION,UCLAvg,LCLAvg,MeanAVG,MeanDispersion,Exclude,UCLDis,LCLDis,LimitsApply) values(@SGNo,@Charid,@SGAvg,@SGDisp,@UCLAvg,@LCLAvg,@MeanAvg,@MeanDisp,@Exclude,@UCLDisp,@LCLDisp,@LimitsApply)", connection);
                cmd.Transaction = transaction;
                cmd.Parameters.Add("@SGNo", OleDbType.Integer).Value = SGNo;
                cmd.Parameters.Add("@Charid", OleDbType.Integer).Value = Charid;
                cmd.Parameters.Add("@SGAvg", OleDbType.Double).Value = SGAvg;
                cmd.Parameters.Add("@SGDisp", OleDbType.Double).Value = SGDisp;
                cmd.Parameters.Add("@UCLAvg", OleDbType.Double).Value = UCLAvg;
                cmd.Parameters.Add("@LCLAvg", OleDbType.Double).Value = LCLAvg;
                cmd.Parameters.Add("@MeanAvg", OleDbType.Double).Value = MeanAvg;
                cmd.Parameters.Add("@MeanDisp", OleDbType.Double).Value = MeanDisp;
                cmd.Parameters.Add("@Exclude", OleDbType.VarChar).Value = Exclude;
                cmd.Parameters.Add("@UCLDisp", OleDbType.Double).Value = UCLDisp;
                cmd.Parameters.Add("@LCLDisp", OleDbType.Double).Value = LCLDisp;
                cmd.Parameters.Add("@LimitsApply", OleDbType.VarChar).Value = LimitsApply;
                cmd.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void insertIntoSGStatSomeVal(string filePath, int SGNo, int Charid, double SGAvg, double SGDisp, string Exclude,string LimitsApply)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Insert into SGStat(SGNO,CharID,SGAVG,SGDISPERSION,Exclude,LimitsApply) values(@SGNo,@Charid,@SGAvg,@SGDisp,@Exclude,@LimitsApply)", connection);
                cmd.Transaction = transaction;
                cmd.Parameters.Add("@SGNo", OleDbType.Integer).Value = SGNo;
                cmd.Parameters.Add("@Charid", OleDbType.Integer).Value = Charid;
                cmd.Parameters.Add("@SGAvg", OleDbType.Double).Value = SGAvg;
                cmd.Parameters.Add("@SGDisp", OleDbType.Double).Value = SGDisp;
                cmd.Parameters.Add("@Exclude", OleDbType.VarChar).Value = Exclude;
                cmd.Parameters.Add("@LimitsApply", OleDbType.VarChar).Value = LimitsApply;
                cmd.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public int SelectSGCharIdFromSGStat(string filePath, int SGNo,int charid)
        {
            int SGCharId = 0;
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SGCharID from SGStat where SGNO = " + SGNo + "and CharID = "+charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            while (reader.Read())
            {
                SGCharId = Convert.ToInt32(reader[0]);
            }
            reader.Close();
            return SGCharId;
        }

        public void UpdateSGStat(string filePath, double UCLAvg, double LCLAvg, double MeanAVG, double MeanDispersion, double UCLDis, double LCLDis,int charId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Update SGStat set UCLAvg = @UCLAvg,LCLAvg = @LCLAvg,MeanAVG = @MeanAVG ,MeanDispersion = @MeanDispersion,UCLDis = @UCLDis ,LCLDis= @LCLDis  Where CharID = " + charId, connection);
                cmd.Parameters.Add("@UCLAvg", OleDbType.Double).Value = UCLAvg;
                cmd.Parameters.Add("@LCLAvg", OleDbType.Double).Value = LCLAvg;
                cmd.Parameters.Add("@MeanAVG", OleDbType.Double).Value = MeanAVG;
                cmd.Parameters.Add("@MeanDispersion", OleDbType.Double).Value = MeanDispersion;
                cmd.Parameters.Add("@UCLDis", OleDbType.Double).Value = UCLDis;
                cmd.Parameters.Add("@LCLDis", OleDbType.Double).Value = LCLDis;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public OleDbDataReader getSGAVG_SGDispFromSGStat(string filePath, int Charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SGAVG,SGDISPERSION,Exclude from SGStat where CharID = " + Charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public double ReadSGAVG_ForSGCharId(string filePath, int SGNo, int Charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SGAVG from SGStat where SGNO = " + SGNo + "and CharID = " + Charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            reader.Read();
            return Convert.ToDouble(reader[0]);
        }

        public double ReadSGDisp_ForSGCharId(string filePath, int SGNo, int Charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SGDISPERSION from SGStat where SGNO = " + SGNo + "and CharID = " + Charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            reader.Read();
            return Convert.ToDouble(reader[0]);
        }

        public OleDbDataReader readSgNoCharidSGCharidfromSGStat(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SGNO,CharID,SGCharID from SGStat", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public void UpdateUSLLSL(string filePath,string USL,string LSL,string charName)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Update Characterstic set USL = @USL,LSL = @LSL Where CharName = '" + charName+"'", connection);
                cmd.Parameters.Add("@USL", OleDbType.VarChar).Value = USL;
                cmd.Parameters.Add("@LSL", OleDbType.VarChar).Value = LSL;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public OleDbDataReader getCharName_USL_LSL_Target_FromCharacterstic(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select CharName,Target,USL,LSL from Characterstic", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readSG_Avg_Disp_forSgno_charid(string filePath, int Sgno, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SGAVG,SGDISPERSION,Exclude,UCLAvg,LCLAvg from SGStat where SGNO = " + Sgno + "and CharID = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //Attempt to roll back the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public void updateSGStat_forSgnoCharid(string filePath, int sgno, int charid, double UCLAvg, double LCLAvg, double MeanAVG, double MeanDispersion, double UCLDis, double LCLDis, string LimitsApply)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Update SGStat set UCLAvg = @UCLAvg ,LCLAvg=@LCLAvg,MeanAVG=@MeanAVG,MeanDispersion=@MeanDispersion,UCLDis=@UCLDis,LCLDis=@LCLDis,LimitsApply ='" + LimitsApply + "' where SGNO = " + sgno + "and CharID =" + charid, connection);
                cmd.Parameters.Add("@UCLAvg", OleDbType.Double).Value = UCLAvg;
                cmd.Parameters.Add("@LCLAvg", OleDbType.Double).Value = LCLAvg;
                cmd.Parameters.Add("@MeanAVG", OleDbType.Double).Value = MeanAVG;
                cmd.Parameters.Add("@MeanDispersion", OleDbType.Double).Value = MeanDispersion;
                cmd.Parameters.Add("@UCLDis", OleDbType.Double).Value = UCLDis;
                cmd.Parameters.Add("@LCLDis", OleDbType.Double).Value = LCLDis;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //Attempt to roll back the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public OleDbDataReader readSGStat(string filePath,int sgno,int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SGAVG,SGDISPERSION,UCLAvg,LCLAvg,MeanAVG,MeanDispersion,UCLDis,LCLDis from SGStat where SGNO = "+sgno+"and CharID = "+charid,connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            } 
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //Attempt to roll back the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : "+ex2.Message);
                }
            }
            return reader;
        }

        public void insertIntoCharthistory(string filePath, int CharId, int BasisStart, int BasisEnd, int AppStart, int AppEnd)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Insert into Charthistory(CharId,BasisStart,BasisEnd,AppStart,AppEnd) values(@CharId,@BasisStart,@BasisEnd,@AppStart,@AppEnd)", connection);
                cmd.Transaction = transaction;
                cmd.Parameters.Add("@CharId", OleDbType.Integer).Value = CharId;
                cmd.Parameters.Add("@BasisStart", OleDbType.Integer).Value = BasisStart;
                cmd.Parameters.Add("@BasisEnd", OleDbType.Integer).Value = BasisEnd;
                cmd.Parameters.Add("@AppStart", OleDbType.Integer).Value = AppStart;
                cmd.Parameters.Add("@AppEnd", OleDbType.Integer).Value = AppEnd;
                cmd.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : "+ex.ToString());
                //Attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type : "+ex2.GetType());
                    MessageBox.Show("Message : "+ex2.Message);
                }
            }
        }

        public OleDbDataReader readFromCharthistory(string filePath, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select BasisStart,BasisEnd,AppStart,AppEnd from Charthistory where CharId = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //Attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message: " + ex2.Message);
                }
            }
            return reader;
        }

        public void UpdateSGAVG_SGDISPERSION_InSGStat(string filePath, int Sgno, int charid, double SGAvg, double SGDisp)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Update SGStat set SGAVG =@SGAvg,SGDISPERSION = @SGDisp where SGNO = " + Sgno + "and CharID = " + charid, connection);
                cmd.Parameters.Add("@SGAvg", OleDbType.Double).Value = SGAvg;
                cmd.Parameters.Add("@SGDisp", OleDbType.Double).Value = SGDisp;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //Attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : "+ex2.GetType());
                    MessageBox.Show("Message : "+ex2.Message);
                }
            }
        }

        public void UpdateSGStatForApplyControlLimits(string filePath, int sgno, int charid, double UCLAvg, double LCLAvg, double MeanAVG, double MeanDispersion, double UCLDis, double LCLDis)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Update SGStat set UCLAvg = @UCLAvg,LCLAvg = @LCLAvg,MeanAVG = @MeanAVG,MeanDispersion = @MeanDispersion ,UCLDis = @UCLDis,LCLDis=@LCLDis where SGNO = " + sgno + "and CharID = " + charid, connection);
                cmd.Parameters.Add("@UCLAvg", OleDbType.Double).Value = UCLAvg;
                cmd.Parameters.Add("@LCLAvg", OleDbType.Double).Value = LCLAvg;
                cmd.Parameters.Add("@MeanAVG", OleDbType.Double).Value = MeanAVG;
                cmd.Parameters.Add("@MeanDispersion", OleDbType.Double).Value = MeanDispersion;
                cmd.Parameters.Add("@UCLDis", OleDbType.Double).Value = UCLDis;
                cmd.Parameters.Add("@LCLDis", OleDbType.Double).Value = LCLDis;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
        }

        public OleDbDataReader readAppControlLimits_SGStat(string filePath, int charid, int sgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select LimitsApply from SGStat where CharID = " + charid + "and SGNO = " + sgno, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public int readChartType(string filePath, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select ChartType from Characterstic where CharID = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToInt32(reader[0]);
            else
                return 0;
        }

        public OleDbDataReader readSgnoFromSgstat(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select distinct SGNO from SGStat", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readSgnoFromSgstat(string filePath,int sgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select distinct top 20 SGNO from SGStat where SGNO <= " + sgno + " order by SGNO desc", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readCharidFromCharthistory(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select distinct CharId from Charthistory", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader ReadValFromSGdataForCharidNRdgNo(string filePath, int charId, int rdgNo)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SGDATA.Value,SGStat.Exclude from SGDATA,SGStat where SGDATA.CharID = " + charId + "and SGDATA.RdgNo = " + rdgNo + " and SGDATA.SGNO = SGStat.SGNO", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public double readSgSizeForCharid(string filePath, int charId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SGSize from Characterstic where CharID = " + charId, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    //MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            reader.Read();
            double sgsize = Convert.ToDouble(reader[0]);
            return sgsize;
        }

        public double readSgSizeForCharname(string filePath, string  charname)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SGSize from Characterstic where CharName = '" + charname + "'", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    //MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            reader.Read();
            return Convert.ToDouble(reader[0]);
        }


        public OleDbDataReader readNCCatFromNC(string filePath, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select NCCat from NC where CharID = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readValueFromNcParato(string filePath, int charId, int NCId, int Sgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select value from NcParato where CharID = " + charId + "and NCId = " + NCId + "and Sgno = " + Sgno, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader Read_NCId_NCCat_From_NC(string filePath, int CharId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select NCID,NCCat from NC where CharID = " + CharId, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollbak the transaction
                try
                {
                    transaction.Commit();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Transaction Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readNCId_Charid_FromNC(string filePath,string NCCat)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select NCID,CharID from NC where NCCat = '"+NCCat+"'", connection);
                cmd.Transaction = transaction; 
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollbak the transaction
                try
                {
                    transaction.Commit();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Transaction Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public int readNCIdFromNC(string filePath, string NCCat,int CharId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select NCID from NC where NCCat = '" + NCCat + "' and CharID = " + CharId, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollbak the transaction
                try
                {
                    transaction.Commit();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Transaction Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToInt32(reader[0]);
            else
                return 0;
        }

        public void insertIntoNcParato(string filePath, int CharID, int NCId, int value, int Sgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Insert into NcParato values('"+CharID+"','"+NCId+"','"+value+"','"+Sgno+"')", connection);
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //Attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
        }

        public void UpdateValueOfNcParato(string filePath, int CharID, int NCId, int Sgno, int value)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Update NcParato set [value] = "+value+" where CharID = " + CharID + " and NCId = " + NCId + " and Sgno = " + Sgno, connection);
                cmd.Transaction = transaction;
                //cmd.Parameters.Add("@value", OleDbType.Integer).Value = value;
                cmd.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
        }

        public int readDataEntryMode(string filePath, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select DataEntry from Characterstic where CharID = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            reader.Read();
            return Convert.ToInt32(reader[0]);
        }

        public int readDataEntryModeForCharName(string filePath, string charName)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select DataEntry from Characterstic where CharName = '" + charName+ "'", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            reader.Read();
            return Convert.ToInt32(reader[0]);
        }

        public OleDbDataReader readCharName_Target_USL_LSL(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select CharName,Target,USL,LSL from Characterstic", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public int getSgSize(string filePath, string charName)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SGSize from Characterstic where CharName = '" + charName + "'", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            reader.Read();
            return Convert.ToInt32(reader[0]);
        }

        public OleDbDataReader readTraceValForTraceCat(string filePath,string traceCat)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Tracevalue from Trace where TraceCat = '"+traceCat+"'", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readTracIdForTraceCatVal(string filePath, string traceCat,string traceVal)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select TraceId from Trace where TraceCat = '" + traceCat + "' and Tracevalue = '"+traceVal+"'", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readSgnoForTracId(string filePath, int traceId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                //Commented and Added by Dhanashri S on 7 May 2018
                //OleDbCommand cmd = new OleDbCommand("Select SgNo from SgTrace where TraceID = "+traceId, connection);
                OleDbCommand cmd = new OleDbCommand("Select DISTINCT SgNo from SgTrace where TraceID = " + traceId, connection);
                //End of Comment and Addition by Dhanashri S
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public int readChartTypeForCharName(string filePath, string charname)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select ChartType from Characterstic where CharName = '"+ charname+"'", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToInt32(reader[0]);
            else
                return 0;
        }

        public OleDbDataReader readNCParato(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select CharID,NCId,value,Sgno from NcParato", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public string readCharNameForCharid(string filePath, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select CharName from Characterstic where CharID = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToString(reader[0]);
            else
                return "";
        }

        public string readNCCatForNCid(string filePath, int NCid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select NCCat from NC where NCID = " + NCid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToString(reader[0]);
            else
                return "";
        }

        public OleDbDataReader getProcessNamePartId_NameFromProcess(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select ProcName,PartName,PartNo from Process", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public double readTargetForCharName(string filePath, string charname)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Target from Characterstic where CharName = '" + charname + "'", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public OleDbDataReader readCharName_SGSize(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select CharName,SGSize from Characterstic", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readChartType_Charname(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select ChartType,CharName from Characterstic", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public int readNcIdForNcCat_Charid(string filePath, int charid ,string NcCat)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select NCID from NC where CharID = " + charid + " and NCCat = '" + NcCat + "'" , connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToInt32(reader[0]);
            else
                return 0;
        }

        public OleDbDataReader readNCCat(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdreadNCCat = new OleDbCommand("Select NCCat from NC", connection);
                cmdreadNCCat.Transaction = transaction;
                transaction.Commit();
                reader = cmdreadNCCat.ExecuteReader();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public void UpdateNccatFromNCForNCID(string filePath, int NCID, string NCCat)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdreadNCCat = new OleDbCommand("Update NC set NCCat = @NCCat where NCID = @NCID", connection);
                cmdreadNCCat.Parameters.Add("@NCCat", OleDbType.VarChar).Value = NCCat;
                cmdreadNCCat.Parameters.Add("@NCID", OleDbType.Integer).Value = NCID;
                cmdreadNCCat.Transaction = transaction;
                transaction.Commit();
                reader = cmdreadNCCat.ExecuteReader();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public double getMaxUCL(string filePath, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Max(UCLAvg) from SGStat where CharID = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public double getMinLCL(string filePath, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Min(LCLAvg) from SGStat where CharID = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public double getMaxUCLDis(string filePath, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Max(UCLDis) from SGStat where CharID = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public double getMinLCLDis(string filePath, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Min(LCLDis) from SGStat where CharID = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public double getMinSGAVG(string filePath, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Min(SGAVG) from SGStat where CharID = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public double getMaxSGAVG(string filePath, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Max(SGAVG) from SGStat where CharID = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public double getMinSGDISP(string filePath, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Min(SGDISPERSION) from SGStat where CharID = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public double getMaxSGDISP(string filePath, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Max(SGDISPERSION) from SGStat where CharID = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public void UpdateSGStat(string filePath, double UCLAvg, double LCLAvg, double MeanAVG, double MeanDispersion, double UCLDis, double LCLDis, int charId,int SGNo)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Update SGStat set UCLAvg = @UCLAvg,LCLAvg = @LCLAvg ,MeanAVG = @MeanAVG ,MeanDispersion = @MeanDispersion ,UCLDis=@UCLDis ,LCLDis=@LCLDis Where CharID = " + charId + "and SGNO = " + SGNo, connection);
                cmd.Parameters.Add("@UCLAvg", OleDbType.Double).Value = UCLAvg;
                cmd.Parameters.Add("@LCLAvg", OleDbType.Double).Value = LCLAvg;
                cmd.Parameters.Add("@MeanAVG", OleDbType.Double).Value = MeanAVG;
                cmd.Parameters.Add("@MeanDispersion", OleDbType.Double).Value = MeanDispersion;
                cmd.Parameters.Add("@UCLDis", OleDbType.Double).Value = UCLDis;
                cmd.Parameters.Add("@LCLDis", OleDbType.Double).Value = LCLDis;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void UpdateSGStat(string filePath, double UCLAvg, double LCLAvg, int charId, int SGNo)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Update SGStat set UCLAvg = @UCLAvg ,LCLAvg =@LCLAvg Where CharID = " + charId + "and SGNO = " + SGNo, connection);
                cmd.Parameters.Add("@UCLAvg", OleDbType.Double).Value = UCLAvg;
                cmd.Parameters.Add("@LCLAvg", OleDbType.Double).Value = LCLAvg;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void UpdateSGStat(string filePath, double MeanAVG, double MeanDispersion, double UCLDis, double LCLDis, int charId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Update SGStat set MeanAVG = @MeanAVG,MeanDispersion =@MeanDispersion ,UCLDis=@UCLDis,LCLDis=@LCLDis Where CharID = " + charId, connection);
                cmd.Parameters.Add("@MeanAVG", OleDbType.Double).Value = MeanAVG;
                cmd.Parameters.Add("@MeanDispersion", OleDbType.Double).Value = MeanDispersion;
                cmd.Parameters.Add("@UCLDis", OleDbType.Double).Value = UCLDis;
                cmd.Parameters.Add("@LCLDis", OleDbType.Double).Value = LCLDis;

                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public OleDbDataReader ReadValFromSGdataForCharidNRdgNo(string filePath, int charId, int rdgNo,int sgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SGDATA.Value,SGStat.Exclude from SGDATA,SGStat where SGDATA.CharID = " + charId + "and SGDATA.RdgNo = " + rdgNo + "and SGDATA.SGNO = " + sgno + " and SGDATA.SGNO = SGStat.SGNO", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public void updateSGStat_forSgnoCharid(string filePath, int sgno, int charid, double MeanAVG, string LimitsApply)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Update SGStat set MeanAVG = @MeanAVG ,LimitsApply ='" + LimitsApply + "' where SGNO = " + sgno + "and CharID =" + charid, connection);
                cmd.Parameters.Add("@MeanAVG", OleDbType.Double).Value = MeanAVG;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //Attempt to roll back the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void renameTraceCatInTrace(string filePath, string oldtraceCat, string newtraceCat)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Update Trace set TraceCat = '" + newtraceCat + "' where TraceCat = '" + oldtraceCat + "'", connection);
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
        }

        public void renameEventCatInEvent(string filePath, string oldeventCat, string neweventCat)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Update Event set EventeCat = '" + neweventCat + "' where EventeCat = '" + oldeventCat + "'", connection);
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
        }

        public bool checksgnointosgtrace(string filePath, int Sgno)
        {

            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            OleDbDataReader rd = null;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SgNo from SgTrace where SgNo = "+Sgno, connection);
                cmd.Transaction = transaction;
                rd = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (rd.Read())
                return true;
            else
                return false;
        }

        public OleDbDataReader getSMTPDetails(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select * from SMTP", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader getUserDetails(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select * from UserList", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader getEmailEvents(string filePath, string charName)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                //OleDbCommand cmd = new OleDbCommand("Select EventName,EventValue from EmailEvents where EventApplied = " + true, connection);
                OleDbCommand cmd = new OleDbCommand("Select EventName,EventValue,Operator from EmailEvents where CharName = '" + charName + "'", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader getEmailEvents(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select EventName,EventValue from EmailEvents", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public void UpdateSGDISPERSION(string filePath, double SGDISPERSION, int charId, int SGNo)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Update SGStat set SGDISPERSION = @SGDISPERSION Where CharID = " + charId + "and SGNO = " + SGNo, connection);
                cmd.Parameters.Add("@SGDISPERSION", OleDbType.Double).Value = SGDISPERSION;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public OleDbDataReader readdatacharwise(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("select sgstat.SGNO,SGAVG,SGDISPERSION,characterstic.charname,sgheader.sgdate,sgheader.sgtime,sgstat.UCLAvg,sgstat.LCLAvg from sgstat,characterstic,sgheader where Sgstat.charID=characterstic.charID and Sgstat.sgno=sgheader.sgno Order by Sgstat.charID,sgstat.sgno asc", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readdatavaluecharwise(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("select sgno,rdgno,value,charid from sgdata order by charid,sgno,rdgno asc", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readdatasgnowise(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("select sgstat.SGNO,SGAVG,SGDISPERSION,characterstic.charname,sgheader.sgdate,sgheader.sgtime,sgstat.UCLAvg,sgstat.LCLAvg from sgstat,characterstic,sgheader where Sgstat.charID=characterstic.charID and Sgstat.sgno=sgheader.sgno Order by sgstat.sgno asc,characterstic.charname;", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readdatavaluesgnowise(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("select sgno,rdgno,value,charid from sgdata order by sgno,charid,rdgno asc;", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("RollBack Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public void updateProcessTable(string filePath, string ProcessName,string PartName,string PartNo)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Update Process set ProcName = '" + ProcessName + "',PartName ='" + PartName + "',PartNo='" + PartNo + "'", connection);
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //Attempt to roll back the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public double getMaxUCL(string filePath, int charid, int fromSgno,int toSgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Max(UCLAvg) from SGStat where CharID = " + charid + " and SGNO between " + fromSgno + " and " + toSgno, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public double getMinLCL(string filePath, int charid, int fromSgno, int toSgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Min(LCLAvg) from SGStat where CharID = " + charid + " and SGNO between " + fromSgno + " and " + toSgno, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public double getMaxUCLDis(string filePath, int charid, int fromSgno, int toSgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Max(UCLDis) from SGStat where CharID = " + charid + " and SGNO between " + fromSgno + " and " + toSgno, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public double getMinLCLDis(string filePath, int charid, int fromSgno, int toSgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Min(LCLDis) from SGStat where CharID = " + charid + " and SGNO between " + fromSgno + " and " + toSgno, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public double getMinSGAVG(string filePath, int charid, int fromSgno, int toSgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Min(SGAVG) from SGStat where CharID = " + charid + " and SGNO between " + fromSgno + " and " + toSgno, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public double getMaxSGAVG(string filePath, int charid, int fromSgno, int toSgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Max(SGAVG) from SGStat where CharID = " + charid + " and SGNO between " + fromSgno + " and " + toSgno, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public double getMinSGDISP(string filePath, int charid, int fromSgno, int toSgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Min(SGDISPERSION) from SGStat where CharID = " + charid + " and SGNO between " + fromSgno + " and " + toSgno, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public double getMaxSGDISP(string filePath, int charid, int fromSgno, int toSgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Max(SGDISPERSION) from SGStat where CharID = " + charid + " and SGNO between " + fromSgno + " and " + toSgno, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToDouble(reader[0]);
            else
                return 0;
        }

        public void deleteFromSGData(string filePath, int fromSgno, int toSgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteFromSGData = new OleDbCommand("Delete from SGDATA where SGNO between " + fromSgno + " and " + toSgno, connection);
                cmdDeleteFromSGData.Transaction = transaction;
                cmdDeleteFromSGData.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void deleteFromSGHeader(string filePath, int fromSgno, int toSgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteFromSGHeader = new OleDbCommand("Delete from SGHeader where SGNO between " + fromSgno + " and " + toSgno, connection);
                cmdDeleteFromSGHeader.Transaction = transaction;
                cmdDeleteFromSGHeader.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void deleteFromSGStat(string filePath, int fromSgno, int toSgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteFromSGStat = new OleDbCommand("Delete from SGStat where SGNO between " + fromSgno + " and " + toSgno, connection);
                cmdDeleteFromSGStat.Transaction = transaction;
                cmdDeleteFromSGStat.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void deleteFromSGEventSGNO(string filePath, int fromSgno, int toSgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteFromSGEvent = new OleDbCommand("Delete from SgEvent where SgNo between " + fromSgno + " and " + toSgno, connection);
                cmdDeleteFromSGEvent.Transaction = transaction;
                cmdDeleteFromSGEvent.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void deleteFromSGTraceSGNO(string filePath, int fromSgno, int toSgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteFromSGTrace = new OleDbCommand("Delete from SgTrace where SgNo between " + fromSgno + " and " + toSgno, connection);
                cmdDeleteFromSGTrace.Transaction = transaction;
                cmdDeleteFromSGTrace.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void deleteFromNcParato(string filePath, int fromSgno, int toSgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdDeleteFromSGTrace = new OleDbCommand("Delete from NcParato where Sgno between " + fromSgno + " and " + toSgno, connection);
                cmdDeleteFromSGTrace.Transaction = transaction;
                cmdDeleteFromSGTrace.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public OleDbDataReader ReadSgnoFromSGStat(string filePath, int fromSgno, int toSgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdSelectFromSGStat = new OleDbCommand("Select distinct SGNO from SGStat where SGNO between " + fromSgno + " and " + toSgno, connection);
                cmdSelectFromSGStat.Transaction = transaction;
                reader =  cmdSelectFromSGStat.ExecuteReader();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader ReadEventCatValForSgno(string filePath, int sgno)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdSelectFromSGStat = new OleDbCommand("Select EventCat,EventValue from Event,SgEvent where Event.EventID = SgEvent.EventID and SgEvent.SgNo = " + sgno, connection);
                cmdSelectFromSGStat.Transaction = transaction;
                reader = cmdSelectFromSGStat.ExecuteReader();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readValforNonExcludesgnosFromSgdata(string filePath, int charId, int frmoSGNo, int toSGNo)
        {
            //connection = new OleDbConnection(conString + filePath);
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadValues = new OleDbCommand("Select value,Exclude from SGDATA inner join SGStat on SGDATA.SGCharID=SGStat.SGCharID where SGDATA.CharID=" + charId + " and SGStat.SGNO between " + frmoSGNo+" and "+toSGNo+" and Exclude = 'N'", connection);
                cmdReadValues.Transaction = transaction;
                reader = cmdReadValues.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readSG_Avg_Disp_forNonExcludeSgno_charid(string filePath, int frmoSGNo, int toSGNo, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SGAVG,SGDISPERSION from SGStat where SGNO between " + frmoSGNo + " and " + toSGNo + " and Exclude = 'N' and CharID = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //Attempt to roll back the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public OleDbDataReader readSGStatFromTosgnos(string filePath, int fromsgno,int tosgno, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select SGAVG,SGDISPERSION,UCLAvg,LCLAvg,MeanAVG,MeanDispersion,UCLDis,LCLDis,SGNO from SGStat where SGNO between " + fromsgno + " and " + tosgno + " and CharID = " + charid, connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //Attempt to roll back the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            return reader;
        }

        public void Save_ComLogo(string Db_path, Byte[] ImgLogoByte, string logo_name)
        {
            #region Delete Image & Replace it with new Image

            try
            {
                connection = new OleDbConnection(conString + Db_path);
                if (connection.State == System.Data.ConnectionState.Closed)
                    connection.Open();

                OleDbTransaction transaction;
                //start transaction
                transaction = connection.BeginTransaction();
                try
                {
                    OleDbCommand DeleteCmd = new OleDbCommand("Delete from AttributeSettingForCompanyLogo where AttributeName = 'CompanyLogo'", connection);
                    DeleteCmd.Transaction = transaction;
                    DeleteCmd.ExecuteNonQuery();
                    DeleteCmd = new OleDbCommand("Delete from AttributeSetting where AttributeName = 'CompanyLogoName'", connection);
                    DeleteCmd.Transaction = transaction;
                    DeleteCmd.ExecuteNonQuery();
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception : " + ex.Message);
                    // Attempt to roll back the transaction.
                    try
                    {
                        transaction.Rollback();
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred
                        // on the server that would cause the rollback to fail, such as
                        // a closed connection.
                        MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                        MessageBox.Show("Message:" + ex2.Message);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception : " + e.ToString());
            }
            #endregion

            #region Save Image
            try
            {
                connection = new OleDbConnection(conString + Db_path);
                if (connection.State == System.Data.ConnectionState.Closed)
                    connection.Open();

                OleDbTransaction transaction;
                //start transaction
                transaction = connection.BeginTransaction();
                try
                {
                    OleDbCommand command;
                    if (ImgLogoByte.Length == 0)
                    {
                        command = new OleDbCommand("Insert into AttributeSettingForCompanyLogo(AttributeName,AttributeValue) values(@AttributeName,@AttributeValue)", connection);
                        command.Transaction = transaction;
                        command.Parameters.Add("@AttributeName", OleDbType.VarChar).Value = "CompanyLogo";
                        command.Parameters.Add("@AttributeValue", OleDbType.LongVarBinary, ImgLogoByte.Length).Value = DBNull.Value;
                        command.ExecuteNonQuery();

                        command = new OleDbCommand("Insert into AttributeSetting(AttributeName,AttributeValue) values(@AttributeName,@AttributeValue)", connection);
                        command.Transaction = transaction;
                        command.Parameters.Add("@AttributeName", OleDbType.VarChar).Value = "CompanyLogoName";
                        command.Parameters.Add("@AttributeValue", OleDbType.VarChar).Value = DBNull.Value;
                        command.ExecuteNonQuery();
                    }
                    else
                    {
                        command = new OleDbCommand("Insert into AttributeSettingForCompanyLogo(AttributeName,AttributeValue) values(@AttributeName,@AttributeValue)", connection);
                        command.Transaction = transaction;
                        command.Parameters.Add("@AttributeName", OleDbType.VarChar).Value = "CompanyLogo";
                        command.Parameters.Add("@AttributeValue", OleDbType.LongVarBinary, ImgLogoByte.Length).Value = ImgLogoByte;
                        command.ExecuteNonQuery();

                        command = new OleDbCommand("Insert into AttributeSetting(AttributeName,AttributeValue) values(@AttributeName,@AttributeValue)", connection);
                        command.Transaction = transaction;
                        command.Parameters.Add("@AttributeName", OleDbType.VarChar).Value = "CompanyLogoName";
                        command.Parameters.Add("@AttributeValue", OleDbType.VarChar).Value = logo_name;
                        command.ExecuteNonQuery();
                    }
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception : " + ex.Message);
                    // Attempt to roll back the transaction.
                    try
                    {
                        transaction.Rollback();
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred
                        // on the server that would cause the rollback to fail, such as
                        // a closed connection.
                        MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                        MessageBox.Show("Message:" + ex2.Message);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception : " + e.ToString());
            }
            #endregion
        }

        public Byte[] get_logoForReport(string DB_path)
        {
            // Get All Format Details For  Selected study
            Byte[] returnimage = null;
            try
            {
                connection = new OleDbConnection(conString + DB_path);
                if (connection.State == System.Data.ConnectionState.Closed)
                    connection.Open();
                OleDbTransaction transaction;
                //start transaction
                transaction = connection.BeginTransaction();
                try
                {
                    OleDbCommand read_image = new OleDbCommand("Select AttributeValue from AttributeSettingForCompanyLogo where AttributeName = 'CompanyLogo'", connection);
                    read_image.Transaction = transaction;
                    object obj = read_image.ExecuteScalar();
                    returnimage = (Byte[])read_image.ExecuteScalar();
                    transaction.Commit();
                    CloseConnection();
                    return returnimage;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception : " + ex.Message);
                    // Attempt to roll back the transaction.
                    try
                    {
                        transaction.Rollback();
                        return returnimage;
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred
                        // on the server that would cause the rollback to fail, such as
                        // a closed connection.
                        MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                        MessageBox.Show("Message:" + ex2.Message);
                        return returnimage;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception : " + e.ToString());
                return returnimage;
            }
        }

        public string get_CompanyLogoName(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdSelectFromSGStat = new OleDbCommand("Select AttributeValue from AttributeSetting where AttributeName = 'CompanyLogoName'", connection);
                cmdSelectFromSGStat.Transaction = transaction;
                reader = cmdSelectFromSGStat.ExecuteReader();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            string str = "";
            if (reader.Read())
            {
                str = Convert.ToString(reader[0]);
            }
            return str;
        }

        public void Delete_ComLogo(string Db_path)
        {
            #region Delete Image & Replace it with new Image

            try
            {
                connection = new OleDbConnection(conString + Db_path);
                if (connection.State == System.Data.ConnectionState.Closed)
                    connection.Open();

                OleDbTransaction transaction;
                //start transaction
                transaction = connection.BeginTransaction();
                try
                {
                    OleDbCommand DeleteCmd = new OleDbCommand("Delete from AttributeSettingForCompanyLogo where AttributeName = 'CompanyLogo'", connection);
                    DeleteCmd.Transaction = transaction;
                    DeleteCmd.ExecuteNonQuery();
                    DeleteCmd = new OleDbCommand("Delete from AttributeSetting where AttributeName = 'CompanyLogoName'", connection);
                    DeleteCmd.Transaction = transaction;
                    DeleteCmd.ExecuteNonQuery();
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception : " + ex.Message);
                    // Attempt to roll back the transaction.
                    try
                    {
                        transaction.Rollback();
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred
                        // on the server that would cause the rollback to fail, such as
                        // a closed connection.
                        MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                        MessageBox.Show("Message:" + ex2.Message);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception : " + e.ToString());
            }
            #endregion
        }

        public string get_ReportInMultipleTabValue(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdSelectFromSGStat = new OleDbCommand("Select AttributeValue from AttributeSetting where AttributeName = 'ReportInMultipleTab'", connection);
                cmdSelectFromSGStat.Transaction = transaction;
                reader = cmdSelectFromSGStat.ExecuteReader();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            string str = "";
            if (reader.Read())
            {
                str = Convert.ToString(reader[0]);
            }
            return str;
        }

        public OleDbDataReader Read_Sgno_From_SGHeader(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);//+ ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadCharIdNSgno = new OleDbCommand("Select SGNo from SgHeader", connection);
                cmdReadCharIdNSgno.Transaction = transaction;
                reader = cmdReadCharIdNSgno.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public void UpdateReportInMultipleTab(string filePath, string AttVal)
        {
            try
            {
                connection = new OleDbConnection(conString + filePath);
                if (connection.State == System.Data.ConnectionState.Closed)
                    connection.Open();

                OleDbTransaction transaction;
                //start transaction
                transaction = connection.BeginTransaction();
                try
                {
                    OleDbCommand DeleteCmd = new OleDbCommand("Delete from AttributeSetting where AttributeName = 'ReportInMultipleTab'", connection);
                    DeleteCmd.Transaction = transaction;
                    DeleteCmd.ExecuteNonQuery();
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception : " + ex.Message);
                    // Attempt to roll back the transaction.
                    try
                    {
                        transaction.Rollback();
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred
                        // on the server that would cause the rollback to fail, such as
                        // a closed connection.
                        MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                        MessageBox.Show("Message:" + ex2.Message);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception : " + e.ToString());
            }

            try
            {
                connection = new OleDbConnection(conString + filePath);
                if (connection.State == System.Data.ConnectionState.Closed)
                    connection.Open();

                OleDbTransaction transaction;
                //start transaction
                transaction = connection.BeginTransaction();
                try
                {
                    OleDbCommand command;

                    command = new OleDbCommand("Insert into AttributeSetting(AttributeName,AttributeValue) values(@AttributeName,@AttributeValue)", connection);
                    command.Transaction = transaction;
                    command.Parameters.Add("@AttributeName", OleDbType.VarChar).Value = "ReportInMultipleTab";
                    command.Parameters.Add("@AttributeValue", OleDbType.VarChar).Value = AttVal;
                    command.ExecuteNonQuery();

                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception : " + ex.Message);
                    // Attempt to roll back the transaction.
                    try
                    {
                        transaction.Rollback();
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred
                        // on the server that would cause the rollback to fail, such as
                        // a closed connection.
                        MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                        MessageBox.Show("Message:" + ex2.Message);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception : " + e.ToString());
            }
        }

        public OleDbDataReader ReadDateTime(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdReadSGDateTime = new OleDbCommand("Select SGDate,SgTime from SGHeader", connection);
                cmdReadSGDateTime.Transaction = transaction;
                reader = cmdReadSGDateTime.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public void InsertTrendMessages(string filePath, string TrendMsg, int SGNO,int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateTraceCategoryValue = new OleDbCommand("Insert into TrendMessages(SGNO,TrendMsg,CharId) values(@SGNO,@TrendMsg,@CharId)", connection);
                cmdUpdateTraceCategoryValue.Transaction = transaction;
                cmdUpdateTraceCategoryValue.Parameters.Add("@SGNO", OleDbType.Integer).Value = SGNO;
                cmdUpdateTraceCategoryValue.Parameters.Add("@TrendMsg", OleDbType.VarChar).Value = TrendMsg;
                cmdUpdateTraceCategoryValue.Parameters.Add("@CharId", OleDbType.Integer).Value = charid;
                cmdUpdateTraceCategoryValue.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void DeleteTrendMessages(string filePath, int SGNO, int CharId)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateTraceCategoryValue = new OleDbCommand("Delete from TrendMessages where CharId = " + CharId + " and SGNO = " + SGNO, connection);
                cmdUpdateTraceCategoryValue.Transaction = transaction;
                cmdUpdateTraceCategoryValue.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void DeleteSGNOTrendMessages(string filePath, int SGNO)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateTraceCategoryValue = new OleDbCommand("Delete from TrendMessages where SGNO = " + SGNO, connection);
                cmdUpdateTraceCategoryValue.Transaction = transaction;
                cmdUpdateTraceCategoryValue.ExecuteNonQuery();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public OleDbDataReader ReadSGNOTrendMessages(string filePath, int SGNO,int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdUpdateTraceCategoryValue = new OleDbCommand("Select TrendMsg from TrendMessages where SGNO = " + SGNO + " and CharId = " + charid, connection);
                cmdUpdateTraceCategoryValue.Transaction = transaction;
                reader = cmdUpdateTraceCategoryValue.ExecuteReader();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

        public void InsertDataEntryBy(string filePath, string DataEntryBy_Val)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand command = new OleDbCommand("Insert into AttributeSetting(AttributeName,AttributeValue) values(@AttributeName,@AttributeValue)", connection);
                command.Transaction = transaction;
                command.Parameters.Add("@AttributeName", OleDbType.VarChar).Value = "DataEntryBy";
                command.Parameters.Add("@AttributeValue", OleDbType.VarChar).Value = DataEntryBy_Val;
                command.ExecuteNonQuery();

                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void InsertDataEntryChar(string filePath, string DataEntryChar_Val)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand command = new OleDbCommand("Insert into AttributeSetting(AttributeName,AttributeValue) values(@AttributeName,@AttributeValue)", connection);
                command.Transaction = transaction;
                command.Parameters.Add("@AttributeName", OleDbType.VarChar).Value = "DataEntryChar";
                command.Parameters.Add("@AttributeValue", OleDbType.VarChar).Value = DataEntryChar_Val;
                command.ExecuteNonQuery();

                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void DeleteDataEntryBy(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand DeleteCmd = new OleDbCommand("Delete from AttributeSetting where AttributeName = 'DataEntryBy'", connection);
                DeleteCmd.Transaction = transaction;
                DeleteCmd.ExecuteNonQuery();
                DeleteCmd.ExecuteNonQuery();

                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public void DeleteDataEntryChar(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand DeleteCmd = new OleDbCommand("Delete from AttributeSetting where AttributeName = 'DataEntryChar'", connection);
                DeleteCmd.Transaction = transaction;
                DeleteCmd.ExecuteNonQuery();
                DeleteCmd.ExecuteNonQuery();

                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
        }

        public int readMaxSGSize(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select max(SGSize) from Characterstic", connection);
                cmd.Transaction = transaction;
                reader = cmd.ExecuteReader();
                transaction.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.ToString());
                //attempt to rollback the transaction
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection
                    MessageBox.Show("Rollback Exception Type : " + ex2.GetType());
                    MessageBox.Show("Message : " + ex2.Message);
                }
            }
            if (reader.Read())
                return Convert.ToInt32(reader[0]);
            else
                return 0;
        }

        public string get_DataEntryChar(string filePath)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                OleDbCommand cmdSelectFromSGStat = new OleDbCommand("Select AttributeValue from AttributeSetting where AttributeName = 'DataEntryChar'", connection);
                cmdSelectFromSGStat.Transaction = transaction;
                reader = cmdSelectFromSGStat.ExecuteReader();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            string str = "";
            if (reader.Read())
            {
                str = Convert.ToString(reader[0]);
            }
            return str;
        }

        public OleDbDataReader GetProcessInformation(string filepath)
        {

            try
            {
                string conString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=";
                connection = new OleDbConnection(conString + filepath);
                if (connection.State == System.Data.ConnectionState.Closed)
                    connection.Open();

                OleDbTransaction transaction;
                //start transaction
                transaction = connection.BeginTransaction();
                try
                {
                    OleDbCommand cmd = new OleDbCommand("Select ProcName,PartName,PartNo from Process ", connection);
                    cmd.Transaction = transaction;
                    reader = cmd.ExecuteReader();
                    transaction.Commit();
                   // CloseConnection();
                    return reader;

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception : " + ex.Message);
                    // Attempt to roll back the transaction.
                    try
                    {
                        return reader;
                        transaction.Rollback();
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred
                        // on the server that would cause the rollback to fail, such as
                        // a closed connection.
                        MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                        MessageBox.Show("Message:" + ex2.Message);
                        return reader;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception : " + e.ToString());
                return reader;
            }

        }

        public OleDbDataReader ReadDataForConsolidtedReport(string filePath,string SGNO, string TraceValue,int TraceId, int charid)
        {
            connection = new OleDbConnection(conString + filePath);// + ";Persist Security Info=False;Mode= Share Deny Write;Jet OLEDB:Database Locking Mode=1");
            if (connection.State == System.Data.ConnectionState.Closed)
                connection.Open();
            OleDbTransaction transaction;
            //start transaction
            transaction = connection.BeginTransaction();
            try
            {
                // OleDbCommand cmdUpdateTraceCategoryValue = new OleDbCommand("select sgdata.sgno,Trace.TraceValue,SGHeader.SGDate,sgdata.value,SGDATA.charid,Characterstic.USL,Characterstic.Target,Characterstic.LSL from(((SGDATA left join sgtrace on sgtrace.[SGNO]= SGDATA.SGNO)left join Trace on Trace.TraceID = sgTrace.TraceID) left join Characterstic on Characterstic.CharID = SGDATA.CharID)left join SGHeader on SGHeader.SGNO = SGDATA.SGNO  Where SGDATA.CharID = " + charid + " and  Trace.TraceValue ='" + TraceValue + "' and SGDATA.SGNO=" + SGNO + " order by SGDATA.CharId,SGDATA.RdgNo asc,SGDATA.sgno asc", connection);
               //  OleDbCommand cmdUpdateTraceCategoryValue = new OleDbCommand("select distinct sgdata.sgno, SGDATA.RdgNo,Trace.TraceCat,Trace.TraceValue,SGHeader.SGDate,sgdata.value,SGDATA.charid,Characterstic.USL,Characterstic.Target,Characterstic.LSL from(((SGDATA left join sgtrace on sgtrace.[SGNO]= SGDATA.SGNO)left join Trace on Trace.TraceID = sgTrace.TraceID) left join Characterstic on Characterstic.CharID = SGDATA.CharID)left join SGHeader on SGHeader.SGNO = SGDATA.SGNO  Where SGDATA.CharID = " + charid + " and Trace.TraceValue ='" + TraceValue + "' and SGDATA.SGNO=" + SGNO +" order by SGDATA.CharId,SGDATA.sgno asc,SGDATA.RdgNo asc", connection);
                OleDbCommand cmdUpdateTraceCategoryValue = new OleDbCommand("SELECT DISTINCT sgdata.sgno, SGDATA.RdgNo, Trace.TraceCat, Trace.TraceValue, SGHeader.SGDate, sgdata.value, SGDATA.charid, Characterstic.USL, Characterstic.Target, Characterstic.LSL, Trace.Traceid FROM(((SGDATA LEFT JOIN sgtrace ON sgtrace.[SGNO] = SGDATA.SGNO) LEFT JOIN Trace ON Trace.TraceID = sgTrace.TraceID) LEFT JOIN Characterstic ON Characterstic.CharID = SGDATA.CharID) LEFT JOIN SGHeader ON SGHeader.SGNO = SGDATA.SGNO Where SGDATA.CharID = " + charid + " and Trace.TraceValue ='" + TraceValue + "' and Trace.TraceId="+TraceId+" and SGDATA.SGNO=" + SGNO + " GROUP BY sgdata.sgno, SGDATA.RdgNo, Trace.TraceCat, Trace.TraceValue, SGHeader.SGDate, sgdata.value, SGDATA.charid, Characterstic.USL, Characterstic.Target, Characterstic.LSL, Trace.TraceValue, Trace.Traceid  ORDER BY Trace.Traceid,sgdata.sgno", connection);
                // OleDbCommand cmdUpdateTraceCategoryValue = new OleDbCommand("select sgdata.sgno,Trace.TraceValue,SGHeader.SGDate,sgdata.value,SGDATA.charid,Characterstic.USL,Characterstic.Target,Characterstic.LSL from(((SGDATA left join sgtrace on sgtrace.[SGNO]= SGDATA.SGNO)left join Trace on Trace.TraceID = sgTrace.TraceID) left join Characterstic on Characterstic.CharID = SGDATA.CharID)left join SGHeader on SGHeader.SGNO = SGDATA.SGNO  Where Trace.TraceValue ='" + TraceValue + "' and SGDATA.CharID = " + charid + " order by SGDATA.charid,SGDATA.sgno asc", connection);
                // OleDbCommand cmdUpdateTraceCategoryValue = new OleDbCommand("select distinct sgdata.sgno, SGDATA.RdgNo,Trace.TraceCat,Trace.TraceValue,SGHeader.SGDate,sgdata.value,SGDATA.charid,Characterstic.USL,Characterstic.Target,Characterstic.LSL from(((SGDATA left join sgtrace on sgtrace.[SGNO]= SGDATA.SGNO)left join Trace on Trace.TraceID = sgTrace.TraceID) left join Characterstic on Characterstic.CharID = SGDATA.CharID)left join SGHeader on SGHeader.SGNO = SGDATA.SGNo order by SGDATA.CharId,SGDATA.sgno asc,SGDATA.RdgNo asc", connection);
                cmdUpdateTraceCategoryValue.Transaction = transaction;
                reader = cmdUpdateTraceCategoryValue.ExecuteReader();
                transaction.Commit();
                //connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception : " + ex.Message);
                // Attempt to roll back the transaction.
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex2)
                {
                    // This catch block will handle any errors that may have occurred
                    // on the server that would cause the rollback to fail, such as
                    // a closed connection.
                    MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                    MessageBox.Show("Message:" + ex2.Message);
                }
            }
            return reader;
        }

    }
}