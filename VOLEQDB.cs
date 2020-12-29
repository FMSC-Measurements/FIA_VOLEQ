using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data.SQLite;
using System.Windows.Forms;
using System.Data;
using System.Reflection;

namespace FIA_VOLEQ
{
    class VOLEQDB
    {
        string dllPath = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        string dbName = "\\FIAVOLEQCONFIG.s3db";

        public SQLiteConnection OpenDB()
        {
            //MessageBox.Show("DB Path and name: " + dllPath+dbName);
            try
            {
                var conn = new SQLiteConnection("Data Source=" + dllPath + dbName + "; FailIfMissing=True");
                return conn;
            }
            catch (SQLiteException e)
            {
                throw new Exception(e.Message);
            }
        }

        //Get list of states for the combobox
        public DataTable getStateList()
        {
            DataTable dTable = new DataTable();
            try
            {
                using (SQLiteConnection cnn = this.OpenDB())
                using (SQLiteCommand cmd = cnn.CreateCommand())
                {
                    cnn.Open();
                    string cmdText = String.Format("SELECT Name FROM States ORDER BY Name");
                    using (SQLiteDataAdapter dAdapter = new SQLiteDataAdapter(cmdText, cnn))
                    {
                        dAdapter.Fill(dTable);
                    }
                    cnn.Close();
                }
            }
            catch (Exception excp)
            {
                throw new Exception("getStateList -- error accessing database." + excp.Message);
            }
            return dTable;
        }
        //Get list of SPN for the combobox
        public DataTable getSpList(string configID)
        {
            DataTable dTable = new DataTable();
            try
            {
                using (SQLiteConnection cnn = this.OpenDB())
                using (SQLiteCommand cmd = cnn.CreateCommand())
                {
                    cnn.Open();
                    string cmdText = String.Format("SELECT SPN FROM MCF_Config WHERE Config_id = '{0}' ORDER BY SPN",configID);
                    using (SQLiteDataAdapter dAdapter = new SQLiteDataAdapter(cmdText, cnn))
                    {
                        dAdapter.Fill(dTable);
                    }
                    cnn.Close();
                }
            }
            catch (Exception excp)
            {
                throw new Exception("getSpList -- error accessing database." + excp.Message);
            }
            return dTable;
        }
        //Get SPN and name list for the combobox
        public DataTable getSpList2(string configID)
        {
            DataTable dTable = new DataTable();
            try
            {
                using (SQLiteConnection cnn = this.OpenDB())
                using (SQLiteCommand cmd = cnn.CreateCommand())
                {
                    cnn.Open();
                    string cmdText = String.Format("SELECT MCF_Config.SPN, TreeSpecies.SPN||' - '||TreeSpecies.Name spn_name FROM MCF_Config, TreeSpecies WHERE MCF_Config.SPN = TreeSpecies.SPN AND MCF_config.Config_id = '{0}' ORDER BY MCF_Config.SPN", configID);
                    using (SQLiteDataAdapter dAdapter = new SQLiteDataAdapter(cmdText, cnn))
                    {
                        dAdapter.Fill(dTable);
                    }
                    cnn.Close();
                }
            }
            catch (Exception excp)
            {
                throw new Exception("getSpList2 -- error accessing database." + excp.Message);
            }
            return dTable;
        }
        //Get selected state abbrevation
        public string getStateAb(string stateName)
        {
            string queryResult = "error";
            try
            {
                using (SQLiteConnection cnn = this.OpenDB())
                using (SQLiteCommand cmd = cnn.CreateCommand())
                {
                    cnn.Open();
                    string cmdText = String.Format("SELECT Abbr FROM States WHERE Name = '{0}'", stateName);
                    cmd.CommandText = cmdText;
                    object retVal = cmd.ExecuteScalar();
                    queryResult = retVal.ToString();
                    cnn.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("error accessing VolEq database.");
                throw new Exception(e.Message);
            }
            return queryResult;
        }
        //Get voleq for the selected species
        public string getVoleq(string tableName, string configID, string sp)
        {
            string queryResult = "";
            try
            {
                using (SQLiteConnection cnn = this.OpenDB())
                using (SQLiteCommand cmd = cnn.CreateCommand())
                {
                    cnn.Open();
                    string cmdText = String.Format("SELECT EQ FROM {0} WHERE Config_ID = '{1}' AND SPN = {2}", tableName, configID, sp);
                    cmd.CommandText = cmdText;
                    object retVal = cmd.ExecuteScalar();
                    if (retVal != null) queryResult = retVal.ToString();
                    cnn.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("error accessing VolEq database.");
                throw new Exception(e.Message);
            }
            return queryResult;
        }
        //Get voleq species for the selected species
        public string getVolSp(string tableName, string configID, string sp)
        {
            string queryResult = "";
            try
            {
                using (SQLiteConnection cnn = this.OpenDB())
                using (SQLiteCommand cmd = cnn.CreateCommand())
                {
                    cnn.Open();
                    string cmdText = String.Format("SELECT VOLSPN FROM {0} WHERE Config_ID = '{1}' AND SPN = {2}", tableName, configID, sp);
                    cmd.CommandText = cmdText;
                    object retVal = cmd.ExecuteScalar();
                    if (retVal != null) queryResult = retVal.ToString();
                    cnn.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("error accessing VolEq database.");
                throw new Exception(e.Message);
            }
            return queryResult;
        }
        //Get sawlog top DIA
        public string getSawTop(string voleq)
        {
            string queryResult = "";
            try
            {
                using (SQLiteConnection cnn = this.OpenDB())
                using (SQLiteCommand cmd = cnn.CreateCommand())
                {
                    cnn.Open();
                    string cmdText = String.Format("SELECT substr(VolType,3,3) FROM Voleq_ref WHERE EQ = '{0}' ", voleq);
                    cmd.CommandText = cmdText;
                    object retVal = cmd.ExecuteScalar();
                    if (retVal != null) queryResult = retVal.ToString();
                    cnn.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("error accessing VolEq database.");
                throw new Exception(e.Message);
            }
            return queryResult;
        }
        //Get voleq ref nnum
        public string getRefNo(string voleq)
        {
            string queryResult;
            try
            {
                using (SQLiteConnection cnn = this.OpenDB())
                using (SQLiteCommand cmd = cnn.CreateCommand())
                {
                    cnn.Open();
                    string cmdText = String.Format("SELECT RefNo FROM VOLEQ_Ref WHERE EQ = '{0}' ", voleq);
                    cmd.CommandText = cmdText;
                    object retVal = cmd.ExecuteScalar();
                    queryResult = retVal.ToString();
                    cnn.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("error accessing VolEq database.");
                throw new Exception(e.Message);
            }
            return queryResult;
        }
        //Get NVEL Eq nnum
        public string getNVEL(string voleq)
        {
            string queryResult;
            try
            {
                using (SQLiteConnection cnn = this.OpenDB())
                using (SQLiteCommand cmd = cnn.CreateCommand())
                {
                    cnn.Open();
                    string cmdText = String.Format("SELECT NVEL_EQ FROM VOLEQ_Ref WHERE EQ = '{0}' ", voleq);
                    cmd.CommandText = cmdText;
                    object retVal = cmd.ExecuteScalar();
                    queryResult = retVal.ToString();
                    cnn.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("error accessing VolEq database.");
                throw new Exception(e.Message);
            }
            return queryResult;
        }
        //get the list of reference
        public DataTable getReference(string sRefNoList)
        {
            DataTable dTable = new DataTable();
            string cmdText;
            try
            {
                using (SQLiteConnection cnn = this.OpenDB())
                using (SQLiteCommand cmd = cnn.CreateCommand())
                {
                    cnn.Open();
                    cmdText = String.Format("SELECT RefNo,  Reference FROM Ref_Citation WHERE RefNo IN {0}", sRefNoList);
                    using (SQLiteDataAdapter dAdapter = new SQLiteDataAdapter(cmdText, cnn))
                    {
                        dAdapter.Fill(dTable);
                    }
                    cnn.Close();

                }
            }
            catch (Exception excp)
            {
                MessageBox.Show("getReference -- error accessing biomass database." + excp.Message);
                throw new Exception(excp.Message);
            }
            return dTable;
        }
    }
}
