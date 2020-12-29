using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections;
using System.Reflection;

namespace FIA_VOLEQ
{
    public partial class Form1 : Form
    {
        [DllImport("vollib.dll")]//, CallingConvention = CallingConvention.StdCall)]
        static extern void VOLLIBCS(ref int regn, StringBuilder forst, StringBuilder voleq, ref float mtopp, ref float mtops,
         ref float stump, ref float dbhob, ref float drcob, StringBuilder httype, ref float httot, ref int htlog, ref float ht1prd, ref float ht2prd,
         ref float upsht1, ref float upsht2, ref float upsd1, ref float upsd2, ref int htref, ref float avgz1, ref float avgz2, ref int fclass, ref float dbtbh,
         ref float btr, ref int i3, ref int i7, ref int i15, ref int i20, ref int i21, float[] vol, float[,] logvol,
         float[,] logdia, float[] loglen, float[] bohlt, ref int tlogs, ref float nologp, ref float nologs, ref int cutflg, ref int bfpflg,
         ref int cupflg, ref int cdpflg, ref int spflg, StringBuilder conspec, StringBuilder prod, ref int httfll, StringBuilder live,
         ref int ba, ref int si, StringBuilder ctype, ref int errflg, ref int indeb, ref int pmtflg, ref MRules mRules, ref int dist, int ll1, int ll2, int ll3, int ll4, int ll5, int ll6, int ll7, int charLen);
        [DllImport("vollib.dll", CallingConvention = CallingConvention.Cdecl)]//EntryPoint = "VERNUM2",
        static extern void VERNUM2(ref int a);

        // standard variables
        int REGN, HTLOG, HTREF, FCLASS, HTTFLL, ERRFLAG, TLOGS, BA, SI, SPCODE, INDEB, PMTFLG, IDIST;
        int CUTFLG, BFPFLG, CUPFLG, CDPFLG, CUSFLG, CDSFLG, SPFLG, VERSION;
        float DBHOB, DRCOB, HTTOT, HT1PRD, HT2PRD, STUMP;
        float UPSHT1, UPSHT2, UPSD1, UPSD2, AVGZ1, AVGZ2;
        float DBTBH, BTR, MTOPP, MTOPS, NOLOGP, NOLOGS, DIB, DIBHT, DOB;

        // test MRULESCS variable
        int OPT;
        float TRIM, MINLEN, MAXLEN, MERCHL, MINLENT;
        // array defintions
        int I3 = 3;
        int I7 = 7;
        int I15 = 15;
        int I20 = 20;
        int I21 = 21;


        float[] VOL;
        float[] LOGLEN;
        float[] BOLHT;

        // 2-dimentional array definitions
        float[,] LOGVOL;
        float[,] LOGDIA;

        // character length parameters
        const int len2 = 2;
        const int len3 = 3;
        const int len5 = 5;
        const int len11 = 11;
        const int strlen = 256;
        const int charLen = 1;

        string m_sDist;
        string sCtype;
        string m_sLive;
        string m_sConSpec;
        string MDL;

        //strings for passing to/from fortran.  Pass fixed length strings
        StringBuilder FORST = new StringBuilder(256);
        StringBuilder DIST = new StringBuilder(256);
        StringBuilder PROD = new StringBuilder(256);
        StringBuilder HTTYPE = new StringBuilder(256);
        StringBuilder VOLEQ = new StringBuilder(256);
        StringBuilder LIVE = new StringBuilder(256);
        StringBuilder CONSPEC = new StringBuilder(256);
        StringBuilder CTYPE = new StringBuilder(256);
        StringBuilder VAR = new StringBuilder(256);

        MRules mRules;

        //variables for this form
        VOLEQDB vdb = new VOLEQDB();
        List<string> StateAbList = new List<string>() { "AZ", "CO", "MT", "NM", "UT", "WY", "CA", "OR", "WA", "AK" };
        List<string> CSList = new List<string>() { "IL", "IN", "IA", "MO" };
        List<string> LSList = new List<string>() { "MI", "MN", "WI"};
        List<string> PSList = new List<string>() { "KS", "NE", "ND", "SD" };
        List<string> S24List = new List<string>() { "CT", "DE", "ME", "MD", "MA", "NH", "NJ", "NY", "OH", "PA", "RI", "VT", "WV" };
        List<string> S33List = new List<string>() { "AL", "AR", "FL", "GA", "LA", "KY", "MS", "OK", "NC", "SC", "TN", "TX", "VA" };
        List<string> EWList = new List<string>() { "CO", "OR", "WA", "MT", "WY" };
        List<string> R2List = new List<string>() { "CO", "SD", "NE", "KS", "WY" };
        string ConfigID;
        string sSawTop;
        string[] errmsg = new string[14]{"No volume equation match",
                                            "No form class",
                                            "DBH less than one",
                                            "Tree height less than 4.5",
                                            "D2H is out of bounds",
                                            "No species match",
                                            "Illegal primary product log height (Ht1prd)",
                                            "Illegal secondary product log height (Ht2prd)",
                                            "Upper stem measurements required",
                                            "Illegal upper stem height (UPSHT1)",
                                            "Unable to fit profile given dbh, merch ht and top dia",
                                            "Tree has more than 20 logs",
                                            "Illegal top diameter",
                                            "The VOLEQ does not have a bark thickness equation"};

        public Form1()
        {
            InitializeComponent();
            loadStateList();
        }

        //load items to State list
        private void loadStateList()
        {
            StateName.Items.Clear();
            DataTable dt = vdb.getStateList();
            //load state Name to combobox
            foreach (DataRow dr in dt.Rows)
            {
                StateName.Items.Add(dr["Name"].ToString());
            }
        }

        //load items to Species list
        private void loadSpList(string configID)
        {
            SPN.Items.Clear();
            DataTable dt = vdb.getSpList(configID);
            //load species number to combobox
            foreach (DataRow dr in dt.Rows)
            {
                SPN.Items.Add(dr["SPN"].ToString());
            }
        }
        //load items to species list with valuemember and displayname
        private void loadSpList2(string configID)
        {
            cbSPN.ValueMember = "SPN";
            cbSPN.DisplayMember = "spn_name";
            cbSPN.DataSource = vdb.getSpList2(configID);
            cbSPN.SelectedIndex = -1;
        }
        //State sunregion selected from list
        private void StateSub_SelectedIndexChanged(object sender, EventArgs e)
        {
            //SPN.Items.Clear();
            cbSPN.DataSource = null;
            clearFields();
            string geoSub = "";
            if (StateSub.Text == "Mixed conifer forest") geoSub = "MIX";
            else if (StateSub.Text == "Northern and Eastern") geoSub = "NE";
            else if (StateSub.Text == "Southern and Western") geoSub = "SW";
            else if (StateSub.Text == "Northern") geoSub = "N";
            else if (StateSub.Text == "Southern") geoSub = "S";
            else if (StateSub.Text == "Eastern") geoSub = "E";
            else if (StateSub.Text == "Western") geoSub = "W";
            else if (StateSub.Text == "Jackson and Josephine Counties") geoSub = "JJ";
            else if (StateSub.Text == "Silver Fir Zone") geoSub = "CF";
            else if (StateSub.Text == "Coastal Alaska Southeast and Central") geoSub = "1AB";
            else if (StateSub.Text == "Coastal Alaska Kodiak and Afognak Islands") geoSub = "1C";

            if (StateAbbr.Text == "CA")
            {
                ConfigID = "S26LCA";
                if (geoSub == "MIX") ConfigID = "S26LCAMIX";
            }
            else if (StateAbbr.Text == "AK")
            {
                ConfigID = "S27LAK";
                if (geoSub == "1AB") ConfigID = "S27LAK1AB";
                else if (geoSub == "1C") ConfigID = "S27LAK1C";
            }
            else if (StateAbbr.Text == "OR" || StateAbbr.Text == "WA")
            {
                if (geoSub == "E" || geoSub == "W") ConfigID = "S26L" + geoSub + StateAbbr.Text;
                else if (geoSub == "JJ") ConfigID = "S26LORJJ";
                else if (geoSub == "CF") ConfigID = "S26LWACF";
            }
            else
            {
                ConfigID = "S22L" + StateAbbr.Text + geoSub;
            }
            loadSpList(ConfigID);
            loadSpList2(ConfigID);
        }
        //State name selected from the list
        private void StateName_SelectedIndexChanged(object sender, EventArgs e)
        {
            StateSub.Items.Clear();
            //SPN.Items.Clear();
            cbSPN.DataSource = null;
            clearFields();
            StateAbbr.Text = vdb.getStateAb(StateName.Text);
            string stateAB = StateAbbr.Text;
            //set default REGN number
            if (stateAB == "AK") REGN = 10;
            else if (stateAB == "MT" || stateAB == "ND") REGN = 1;
            else if (R2List.Contains(stateAB)) REGN = 2;
            else if (stateAB == "AZ" || stateAB == "NM") REGN = 3;
            else if (stateAB == "NV" || stateAB == "UT") REGN = 4;
            else if (stateAB == "CA" || stateAB == "HI") REGN = 5;
            else if (stateAB == "OR" || stateAB == "WA") REGN = 6;
            else if (S33List.Contains(stateAB)) REGN = 8;
            else REGN = 9;

            if (StateAbList.Contains(StateAbbr.Text) == true)
            {
                //the state has subregion
                GeoSub.Visible = true;
                StateSub.Visible = true;
                if (StateAbbr.Text=="CA")
                {
                    StateSub.Items.Add("Mixed conifer forest");
                    StateSub.Items.Add("Other than mixed conifer forest");
                }
                else if (StateAbbr.Text == "AK")
                {
                    //StateSub.Items.Add("Coastal Alaska Southeast");
                    StateSub.Items.Add("Coastal Alaska Southeast and Central");
                    StateSub.Items.Add("Coastal Alaska Kodiak and Afognak Islands");
                }
                else if (StateAbbr.Text == "UT")
                {
                    StateSub.Items.Add("Northern and Eastern");
                    StateSub.Items.Add("Southern and Western");
                }
                else if (StateAbbr.Text == "AZ" || StateAbbr.Text == "NM")
                {
                    StateSub.Items.Add("Northern");
                    StateSub.Items.Add("Southern");
                }
                else if (EWList.Contains(StateAbbr.Text))
                {
                    StateSub.Items.Add("Eastern");
                    StateSub.Items.Add("Western");
                }

                if (StateAbbr.Text == "OR") StateSub.Items.Add("Jackson and Josephine Counties");
                if (StateAbbr.Text == "WA") StateSub.Items.Add("Silver Fir Zone");
            }
            else
            {
                GeoSub.Visible = false;
                StateSub.Visible = false;
                if (StateAbbr.Text == "HI") ConfigID = "S26LPI";
                else if (CSList.Contains(StateAbbr.Text)) ConfigID = "S23LCS";
                else if (LSList.Contains(StateAbbr.Text)) ConfigID = "S23LLS";
                else if (PSList.Contains(StateAbbr.Text)) ConfigID = "S23LPS";
                else if (S24List.Contains(StateAbbr.Text)) ConfigID = "S24";
                else if (S33List.Contains(StateAbbr.Text)) ConfigID = "S33";
                else if (StateAbbr.Text == "ID") ConfigID = "S22LID";
                else if (StateAbbr.Text == "NV") ConfigID = "S22LNV";

                loadSpList(ConfigID);
                loadSpList2(ConfigID);
            }
            //if (ConfigID == "S24")
            if (S24List.Contains(StateAbbr.Text))
            {
                lbBoleHt.Visible = true;
                lbSawHt.Visible = true;
                txBolHt.Visible = true;
                txSawHt.Visible = true;
            }
            else
            {
                lbBoleHt.Visible = false;
                lbSawHt.Visible = false;
                txBolHt.Visible = false;
                txSawHt.Visible = false;
            }
        }

        //Species selected from the list
        private void Species_SelectedIndexChanged(object sender, EventArgs e)
        {
            //string sRefNo;
            //string sNVEL;
            //string sRefNoList = "(";
            //clearFields();
            //mcfEQ.Text = vdb.getVoleq("MCF_Config", ConfigID, SPN.Text);
            //mcfSP.Text = vdb.getVolSp("MCF_Config", ConfigID, SPN.Text);
            //if (mcfEQ.Text.Length > 0)
            //{
            //    sRefNo = vdb.getRefNo(mcfEQ.Text);
            //    mcfRefNo.Text = formatRefNo(sRefNo);
            //    sNVEL = vdb.getNVEL(mcfEQ.Text);
            //    mcfNVEL.Text = formatNVEL(sNVEL, mcfSP.Text, SPN.Text);
            //    sRefNoList = sRefNoList + mcfRefNo.Text + ",";
            //}
            //tcfEQ.Text = vdb.getVoleq("TCF_Config", ConfigID, SPN.Text);
            //tcfSP.Text = vdb.getVolSp("TCF_Config", ConfigID, SPN.Text);
            //if (tcfEQ.Text.Length > 0)
            //{
            //    sRefNo = vdb.getRefNo(tcfEQ.Text);
            //    tcfRefNo.Text = formatRefNo(sRefNo);
            //    sNVEL = vdb.getNVEL(tcfEQ.Text);
            //    tcfNVEL.Text = formatNVEL(sNVEL, tcfSP.Text, SPN.Text);
            //    sRefNoList = sRefNoList + tcfRefNo.Text + ",";
            //}
            //scfEQ.Text = vdb.getVoleq("SCF_Config", ConfigID, SPN.Text);
            //scfSP.Text = vdb.getVolSp("SCF_Config", ConfigID, SPN.Text);
            //if (scfEQ.Text.Length > 0)
            //{
            //    sRefNo = vdb.getRefNo(scfEQ.Text);
            //    scfRefNo.Text = formatRefNo(sRefNo);
            //    sNVEL = vdb.getNVEL(scfEQ.Text);
            //    scfNVEL.Text = formatNVEL(sNVEL, scfSP.Text, SPN.Text);
            //    sRefNoList = sRefNoList + scfRefNo.Text + ",";
            //    sSawTop = vdb.getSawTop(scfEQ.Text);
            //    if (sSawTop == "*") sSawTop = "6";
            //    else if (sSawTop.Length == 3)
            //    {
            //        if (sSawTop.Substring(1, 2) == "16" || sSawTop.Substring(1, 2) == "32") sSawTop = sSawTop.Substring(0, 1);
            //        else if (sSawTop == "7/9" || sSawTop == "6/8")
            //        {
            //            if (int.Parse(SPN.Text) < 300) sSawTop = sSawTop.Substring(0, 1);
            //            else sSawTop = sSawTop.Substring(2, 1);
            //        }
            //    }
                
            //}
            //scrbfEQ.Text = vdb.getVoleq("ScrBF_Config", ConfigID, SPN.Text);
            //scrbfSP.Text = vdb.getVolSp("ScrBF_Config", ConfigID, SPN.Text);
            //if (scrbfEQ.Text.Length > 0)
            //{
            //    sRefNo = vdb.getRefNo(scrbfEQ.Text);
            //    scrbfRefNo.Text = formatRefNo(sRefNo);
            //    sNVEL = vdb.getNVEL(scrbfEQ.Text);
            //    scrbfNVEL.Text = formatNVEL(sNVEL, scrbfSP.Text, SPN.Text);
            //    sRefNoList = sRefNoList + scrbfRefNo.Text + ",";
            //}
            //intbfEQ.Text = vdb.getVoleq("IntBF_Config", ConfigID, SPN.Text);
            //intbfSP.Text = vdb.getVolSp("IntBF_Config", ConfigID, SPN.Text);
            //if (intbfEQ.Text.Length > 0)
            //{
            //    sRefNo = vdb.getRefNo(intbfEQ.Text);
            //    intbfRefNo.Text = formatRefNo(sRefNo);
            //    sNVEL = vdb.getNVEL(intbfEQ.Text);
            //    intbfNVEL.Text = formatNVEL(sNVEL, intbfSP.Text, SPN.Text);
            //    sRefNoList = sRefNoList + intbfRefNo.Text + ",";
            //}
            //sRefNoList = sRefNoList + "0)";
            ////get list of reference to display
            //DataView dv = new DataView(vdb.getReference(sRefNoList));
            //dataGridView1.DataSource = dv;
        }
        //combobox spn and name selection changed
        private void cbSpecies_SelectedIndexChanged2(object sender, EventArgs e)
        {
            //MessageBox.Show(cbSPN.SelectedValue.ToString());
            string sRefNo;
            string sNVEL;
            string sRefNoList = "(";
            clearFields();
            mcfEQ.Text = vdb.getVoleq("MCF_Config", ConfigID, cbSPN.SelectedValue.ToString());
            mcfSP.Text = vdb.getVolSp("MCF_Config", ConfigID, cbSPN.SelectedValue.ToString());
            if (mcfEQ.Text.Length > 0)
            {
                sRefNo = vdb.getRefNo(mcfEQ.Text);
                mcfRefNo.Text = formatRefNo(sRefNo);
                sNVEL = vdb.getNVEL(mcfEQ.Text);
                mcfNVEL.Text = formatNVEL(sNVEL, mcfSP.Text, cbSPN.SelectedValue.ToString());
                sRefNoList = sRefNoList + mcfRefNo.Text + ",";
            }
            tcfEQ.Text = vdb.getVoleq("TCF_Config", ConfigID, cbSPN.SelectedValue.ToString());
            tcfSP.Text = vdb.getVolSp("TCF_Config", ConfigID, cbSPN.SelectedValue.ToString());
            if (tcfEQ.Text.Length > 0)
            {
                sRefNo = vdb.getRefNo(tcfEQ.Text);
                tcfRefNo.Text = formatRefNo(sRefNo);
                sNVEL = vdb.getNVEL(tcfEQ.Text);
                tcfNVEL.Text = formatNVEL(sNVEL, tcfSP.Text, cbSPN.SelectedValue.ToString());
                sRefNoList = sRefNoList + tcfRefNo.Text + ",";
            }
            scfEQ.Text = vdb.getVoleq("SCF_Config", ConfigID, cbSPN.SelectedValue.ToString());
            scfSP.Text = vdb.getVolSp("SCF_Config", ConfigID, cbSPN.SelectedValue.ToString());
            if (scfEQ.Text.Length > 0)
            {
                sRefNo = vdb.getRefNo(scfEQ.Text);
                scfRefNo.Text = formatRefNo(sRefNo);
                sNVEL = vdb.getNVEL(scfEQ.Text);
                scfNVEL.Text = formatNVEL(sNVEL, scfSP.Text, cbSPN.SelectedValue.ToString());
                sRefNoList = sRefNoList + scfRefNo.Text + ",";
                sSawTop = vdb.getSawTop(scfEQ.Text);
                if (sSawTop == "*") sSawTop = "6";
                else if (sSawTop.Length == 3)
                {
                    if (sSawTop.Substring(1, 2) == "16" || sSawTop.Substring(1, 2) == "32") sSawTop = sSawTop.Substring(0, 1);
                    else if (sSawTop == "7/9" || sSawTop == "6/8")
                    {
                        if (int.Parse(cbSPN.SelectedValue.ToString()) < 300) sSawTop = sSawTop.Substring(0, 1);
                        else sSawTop = sSawTop.Substring(2, 1);
                    }
                }

            }
            scrbfEQ.Text = vdb.getVoleq("ScrBF_Config", ConfigID, cbSPN.SelectedValue.ToString());
            scrbfSP.Text = vdb.getVolSp("ScrBF_Config", ConfigID, cbSPN.SelectedValue.ToString());
            if (scrbfEQ.Text.Length > 0)
            {
                sRefNo = vdb.getRefNo(scrbfEQ.Text);
                scrbfRefNo.Text = formatRefNo(sRefNo);
                sNVEL = vdb.getNVEL(scrbfEQ.Text);
                scrbfNVEL.Text = formatNVEL(sNVEL, scrbfSP.Text, cbSPN.SelectedValue.ToString());
                sRefNoList = sRefNoList + scrbfRefNo.Text + ",";
            }
            intbfEQ.Text = vdb.getVoleq("IntBF_Config", ConfigID, cbSPN.SelectedValue.ToString());
            intbfSP.Text = vdb.getVolSp("IntBF_Config", ConfigID, cbSPN.SelectedValue.ToString());
            if (intbfEQ.Text.Length > 0)
            {
                sRefNo = vdb.getRefNo(intbfEQ.Text);
                intbfRefNo.Text = formatRefNo(sRefNo);
                sNVEL = vdb.getNVEL(intbfEQ.Text);
                intbfNVEL.Text = formatNVEL(sNVEL, intbfSP.Text, cbSPN.SelectedValue.ToString());
                sRefNoList = sRefNoList + intbfRefNo.Text + ",";
            }
            sRefNoList = sRefNoList + "0)";
            //get list of reference to display
            DataView dv = new DataView(vdb.getReference(sRefNoList));
            dataGridView1.DataSource = dv;
        }
        //format reference number to display
        private string formatRefNo(string sRefNo)
        {
            int iRefNo;
            if (sRefNo.Length > 0)
            {
                iRefNo = int.Parse(sRefNo);
                if (iRefNo > 100) sRefNo = (iRefNo - Math.Floor(iRefNo / 100d) * 100).ToString() + ", " + (Math.Floor(iRefNo / 100d) * 100).ToString();
            }
            return sRefNo;
        }
        //format NVEL equation number to replace *** with species number
        private string formatNVEL(string sNVEL, string volsp, string spec)
        {
            char pad = '0';
            if (sNVEL.Length > 0)
            {
                if (sNVEL.Substring(6, 4) == "****")
                {
                    if (volsp.Length > 0) sNVEL = sNVEL.Replace("****", volsp.PadLeft(4, pad));
                    else sNVEL = sNVEL.Replace("****", spec.PadLeft(4, pad));
                }
                else if (sNVEL.Substring(7, 3) == "***")
                {
                    if (volsp.Length > 0) sNVEL = sNVEL.Replace("***", volsp.PadLeft(3, pad));
                    else sNVEL = sNVEL.Replace("***", spec.PadLeft(3, pad));
                }
            }
            return sNVEL;
        }
        //clear voleq fields
        private void clearFields()
        {
            tcfEQ.Clear();
            tcfSP.Clear();
            tcfRefNo.Clear();
            tcfNVEL.Clear();

            mcfEQ.Clear();
            mcfSP.Clear();
            mcfRefNo.Clear();
            mcfNVEL.Clear();

            scfEQ.Clear();
            scfSP.Clear();
            scfRefNo.Clear();
            scfNVEL.Clear();

            scrbfEQ.Clear();
            scrbfSP.Clear();
            scrbfRefNo.Clear();
            scrbfNVEL.Clear();

            intbfEQ.Clear();
            intbfSP.Clear();
            intbfRefNo.Clear();
            intbfNVEL.Clear();

            if (this.dataGridView1.DataSource != null)
            {
                this.dataGridView1.DataSource = null;
            }
            else
            {
                this.dataGridView1.Rows.Clear();
            }
            txCV4.Clear();
            txCVsaw.Clear();
            txCVTS.Clear();
            txScrBF.Clear();
            txIntBF.Clear();
            lbSawTop.Text = "*Sawlog top DIA";
            lbErr.Visible = false;
        }
        //Button click event
        private void calcButton_Click(object sender, EventArgs e)
        {
            string sNVEL = "";
            //string spec = SPN.Text;
            string spec = cbSPN.SelectedValue.ToString();
            int ver = 0;
            VERNUM2(ref ver);
            NVELver.Text = "NVEL Version: " + ver;
            if (txDBH.Text.Length > 0 && txTHT.Text.Length > 0)
            {
                DBHOB = float.Parse(txDBH.Text);
                HTTOT = float.Parse(txTHT.Text);
            }
            else
            {
                MessageBox.Show("Please enter DBH and Total Ht");
                return;
            }
            //Northeast states need input for Bole height and Saw height
            if (ConfigID == "S24")
            {
                if (txBolHt.Text.Length > 0 && txSawHt.Text.Length > 0)
                {
                    HT1PRD = float.Parse(txSawHt.Text);
                    HT2PRD = float.Parse(txBolHt.Text);
                }
                else
                {
                    MessageBox.Show("Please enter BoleHt and SawHt");
                    return;
                }
            }
            CONSPEC.Length = 0;
            VOLEQ.Length = 0;
            if (mcfNVEL.Text.Length > 0)
            {
                sNVEL = mcfNVEL.Text;
                if (mcfSP.Text.Length > 0) spec = mcfSP.Text;
                CONSPEC.Append(spec);
                VOLEQ.Append(sNVEL);
                callNVELdll();
                if (ERRFLAG == 0)
                {
                    float cv4 = VOL[3] + VOL[6];
                    txCV4.Text = Math.Round(cv4,1).ToString();
                }
            }
            if (tcfNVEL.Text.Length > 0)
            {
                float cvts;
                if (tcfNVEL.Text == sNVEL)
                {
                    cvts = VOL[0] + VOL[13];
                    txCVTS.Text = Math.Round(cvts, 1).ToString();
                }
                else
                {
                    CONSPEC.Length = 0;
                    VOLEQ.Length = 0;
                    sNVEL = tcfNVEL.Text;
                    if (tcfSP.Text.Length > 0) spec = tcfSP.Text;
                    CONSPEC.Append(spec);
                    VOLEQ.Append(sNVEL);
                    callNVELdll();
                    cvts = VOL[0] + VOL[13];
                    if (ERRFLAG == 0) txCVTS.Text = Math.Round(cvts, 1).ToString();
                }
            }
            if (scfNVEL.Text.Length > 0)
            {
                if (sSawTop.Length > 0) lbSawTop.Text = "*Sawlog top DIA " + sSawTop + "\"";
                if (scfNVEL.Text == sNVEL) txCVsaw.Text = Math.Round(VOL[3],1).ToString();
                else
                {
                    CONSPEC.Length = 0;
                    VOLEQ.Length = 0;
                    sNVEL = scfNVEL.Text;
                    if (scfSP.Text.Length > 0) spec = scfSP.Text;
                    CONSPEC.Append(spec);
                    VOLEQ.Append(sNVEL);
                    callNVELdll();
                    if (ERRFLAG == 0) txCVsaw.Text = Math.Round(VOL[3],1).ToString();
                    
                }
            }
            if (scrbfNVEL.Text.Length > 0)
            {
                if (scrbfNVEL.Text == sNVEL) txScrBF.Text = Math.Round(VOL[1]).ToString();
                else
                {
                    CONSPEC.Length = 0;
                    VOLEQ.Length = 0;
                    sNVEL = scrbfNVEL.Text;
                    if (scrbfSP.Text.Length > 0) spec = scrbfSP.Text;
                    CONSPEC.Append(spec);
                    VOLEQ.Append(sNVEL);
                    callNVELdll();
                    if (ERRFLAG == 0) txScrBF.Text = Math.Round(VOL[1]).ToString();
                }
            }
            if (intbfNVEL.Text.Length > 0)
            {
                if (intbfNVEL.Text == sNVEL) txIntBF.Text = Math.Round(VOL[9]).ToString();
                else
                {
                    CONSPEC.Length = 0;
                    VOLEQ.Length = 0;
                    sNVEL = intbfNVEL.Text;
                    if (intbfSP.Text.Length > 0) spec = intbfSP.Text;
                    CONSPEC.Append(spec);
                    VOLEQ.Append(sNVEL);
                    callNVELdll();
                    if (ERRFLAG == 0) txIntBF.Text = Math.Round(VOL[9]).ToString();
                }
            }
        }
        //Call NVEL to calculate volumes
        private void callNVELdll()
        {
            clearStrings();
            FORST.Append("01");
            DIST.Append("01");
            IDIST = 1;
            PROD.Append("01");
            HTTYPE.Append("F");
            //VOLEQ.Append(volEqTB.Text);
            //CONSPEC.Append(m_sConSpec);
            LIVE.Append("L");
            CTYPE.Append("F");

            VOL = new float[I15];
            LOGLEN = new float[I20];
            BOLHT = new float[I21];
            // 2-dimentional array definitions
            LOGVOL = new float[I20, I7];
            LOGDIA = new float[I3, I21];

            //REGN = int.Parse(regionTB.Text);
            MTOPP = 6.0F;
            if (sSawTop.Length > 0 && sSawTop != "*")
            {
                if (sSawTop == "7/9" || sSawTop == "6/8")
                {
                    if (int.Parse(cbSPN.SelectedValue.ToString()) < 300) MTOPP = float.Parse(sSawTop.Substring(0, 1));
                    else MTOPP = float.Parse(sSawTop.Substring(2, 1));
                }
                else if (float.TryParse(sSawTop, out MTOPP) == true)
                {
                    MTOPP = float.Parse(sSawTop);
                }
            }
            MTOPS = 4.0F;
            STUMP = 1.0F;
            DRCOB = 0.0F;
            //Northeastern states need input for bole height and sawlog height
            if (ConfigID != "S24")
            {
                HT1PRD = 0.0F;
                HT2PRD = 0.0F;
            }
            HTLOG = 0;
            UPSHT1 = 0.0F;
            UPSHT2 = 0;
            HTTFLL = 0;
            DRCOB = 0;
            UPSD1 = 0;//upper stem diameter 1
            UPSD2 = 0;//upper stem diameter 2
            HTREF = 0;//reference height
            AVGZ1 = 0;
            AVGZ2 = 0;
            FCLASS = 80;
            DBTBH = 0.0F;
            BTR = 0.0F;//bark thickness ratio
            CUTFLG = 1;//cut flag
            BFPFLG = 1;
            CUPFLG = 1;//cubic foot flag;
            CDPFLG = 1;//cord volume flag;
            SPFLG = 1;//;
            PMTFLG = 0;//flag for which profile model to run
            TLOGS = 0;//total # logs;
            NOLOGP = 0.0F;//# logs primary product
            NOLOGS = 0.0F;//# logs secondary product
            BA = 0; //basal area
            SI = 0; //site index
            INDEB = 0;// debugCB.Checked ? 1 : 0;
            ERRFLAG = 0;//error flag
            //call vollib.dll to calculate tree volume
            VOLLIBCS(ref REGN, FORST, VOLEQ, ref MTOPP, ref MTOPS, ref STUMP, ref DBHOB, ref DRCOB,
              HTTYPE, ref HTTOT, ref HTLOG, ref HT1PRD, ref HT2PRD, ref UPSHT1, ref UPSHT2, ref UPSD1, ref UPSD2,
              ref HTREF, ref AVGZ1, ref AVGZ2, ref FCLASS, ref DBTBH, ref BTR, ref I3, ref I7, ref I15, ref I20, ref I21, VOL, LOGVOL,
              LOGDIA, LOGLEN, BOLHT, ref TLOGS, ref NOLOGP, ref NOLOGS, ref CUTFLG, ref BFPFLG, ref CUPFLG,
              ref CDPFLG, ref SPFLG, CONSPEC, PROD, ref HTTFLL, LIVE, ref BA, ref SI, CTYPE, ref ERRFLAG, ref INDEB, ref PMTFLG,
              ref mRules, ref IDIST, strlen, strlen, strlen, strlen, strlen, strlen, strlen, charLen);
            if (ERRFLAG > 0)
            {
                lbErr.Visible = true;
                lbErr.ForeColor = Color.Red;
                lbErr.Text = "Error: " + errmsg[ERRFLAG - 1];
            }
            else
            {
                lbErr.Visible = false;
                lbErr.Text = "";
            }
        }
        private void clearStrings()
        {
            FORST.Length = 0;
            DIST.Length = 0;
            PROD.Length = 0;
            HTTYPE.Length = 0;
            LIVE.Length = 0;
            CTYPE.Length = 0;
        }
    }
}
