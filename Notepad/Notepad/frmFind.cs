using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;



namespace RichTextEditor
{
    public partial class frmFind : Form
    {
        // local member variable to hold main form
        frmMain mMain;


        // default constructor
        public frmFind()
        {
            InitializeComponent();
        }


        // overloaded constructor - permits passing in main form
        public frmFind(frmMain f)
        {
            InitializeComponent();
            mMain = f;
        }




        private void btnFind_Click(object sender, System.EventArgs e)
        {
            try
            {
                int StartPosition;
                StringComparison SearchType;

                if (chkMatchCase.Checked == true)
                {
                    SearchType = StringComparison.Ordinal;
                }
                else
                {
                    SearchType = StringComparison.OrdinalIgnoreCase;
                }

                StartPosition = mMain.rtbDoc.Text.IndexOf(txtSearchTerm.Text, SearchType);

                if (StartPosition == 0)
                {
                    MessageBox.Show("String: " + txtSearchTerm.Text.ToString() + " not found", "No Matches", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    return;
                }

                mMain.rtbDoc.Select(StartPosition, txtSearchTerm.Text.Length);
                mMain.rtbDoc.ScrollToCaret();
                mMain.Focus();
                btnFindNext.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }




        private void btnFindNext_Click(object sender, System.EventArgs e)
        {
            try
            {
                int StartPosition = mMain.rtbDoc.SelectionStart + 2;

                StringComparison SearchType;

                if (chkMatchCase.Checked == true)
                {
                    SearchType = StringComparison.Ordinal;
                }
                else
                {
                    SearchType = StringComparison.OrdinalIgnoreCase;
                }

                //StartPosition = Microsoft.VisualBasic.Strings.InStr(StartPosition, mMain.rtbDoc.Text, txtSearchTerm.Text, SearchType);
                StartPosition = mMain.rtbDoc.Text.IndexOf(txtSearchTerm.Text, StartPosition, SearchType);

                if (StartPosition == 0 || StartPosition < 0)
                {
                    MessageBox.Show("String: " + txtSearchTerm.Text.ToString() + " not found", "No Matches", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    return;
                }

                mMain.rtbDoc.Select(StartPosition, txtSearchTerm.Text.Length);
                mMain.rtbDoc.ScrollToCaret();
                mMain.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }


        private void txtSearchTerm_TextChanged(object sender, EventArgs e)
        {
            btnFindNext.Enabled = false;
        }


    }
}