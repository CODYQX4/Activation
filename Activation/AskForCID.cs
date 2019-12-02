using System;
using System.Drawing;
using System.Windows.Forms;

namespace Activation
{
    /// <summary>
    /// Form for obtaining a Phone Activation Code (Confirmation ID) for an OfflineInstallationID
    /// </summary>
    public partial class AskForCID : Form
    {
        /// <summary>
        /// Publicly Accessible Input Confirmation ID
        /// </summary>
        public static string ConfirmationID { get; private set; }

        /// <summary>
        /// Constructor for AskForCID Form
        /// </summary>
        /// <param name="oid">Offline Installation ID to Obtain Confirmation ID for</param>
        public AskForCID(string oid)
        {
            InitializeComponent();

            // How long is a group of digits?
            int sectionLength = 0;
            if (oid.Length % 6 == 0)
            {
                sectionLength = 6;
            }
            else if (oid.Length % 7 == 0)
            {
                sectionLength = 7;
            }

            // Add Dashes
            for (int index = 0; index < oid.Length; index++ )
            {
                if (index != 0 && index != oid.Length - 1 && index % sectionLength == 0)
                {
                    labelOID.Text += "-";
                }
                labelOID.Text += oid[index];
            }
            // Reset ConfirmationID on Form Creation
            ConfirmationID = string.Empty;
        }

        private void BtnEnterCIDClick(object sender, EventArgs e)
        {
            // Get Input of All Textboxes and add Dashes
            string input = txtEnterCID1.Text.ToUpper() + "-" + txtEnterCID2.Text.ToUpper() + "-" + txtEnterCID3.Text.ToUpper() + "-" + txtEnterCID4.Text.ToUpper() + "-" + txtEnterCID5.Text.ToUpper() + "-" + txtEnterCID6.Text.ToUpper() + "-" + txtEnterCID7.Text.ToUpper() + "-" + txtEnterCID8.Text.ToUpper();

            // Check if CID is Valid
            if (Phone.IsValidCID(input) == false)
            {
                // Set error colors.
                txtEnterCID1.BackColor = Color.Red;
                txtEnterCID1.ForeColor = Color.White;
                txtEnterCID2.BackColor = Color.Red;
                txtEnterCID2.ForeColor = Color.White;
                txtEnterCID3.BackColor = Color.Red;
                txtEnterCID3.ForeColor = Color.White;
                txtEnterCID4.BackColor = Color.Red;
                txtEnterCID4.ForeColor = Color.White;
                txtEnterCID5.BackColor = Color.Red;
                txtEnterCID5.ForeColor = Color.White;
                txtEnterCID6.BackColor = Color.Red;
                txtEnterCID6.ForeColor = Color.White;
                txtEnterCID7.BackColor = Color.Red;
                txtEnterCID7.ForeColor = Color.White;
                txtEnterCID8.BackColor = Color.Red;
                txtEnterCID8.ForeColor = Color.White;

                MessageBox.Show("The CID has an invalid format" + Environment.NewLine + "The correct format is: XXXXXX-XXXXXX-XXXXXX-XXXXXX-XXXXXX-XXXXXX-XXXXXX-XXXXXX");
                return;
            }
            // Remove Dashes and Assign
            ConfirmationID = input.Replace("-", string.Empty);

            // Close Form
            Close();
        }

        private void TxtEnterCID1TextChanged(object sender, EventArgs e)
        {
            if (txtEnterCID1.BackColor == Color.Red)
            {
                txtEnterCID1.BackColor = Color.White;
                txtEnterCID1.ForeColor = Color.Black;
            }
            if (txtEnterCID1.TextLength == 6)
            {
                txtEnterCID2.Focus();
            }
        }

        private void TxtEnterCID2TextChanged(object sender, EventArgs e)
        {
            if (txtEnterCID2.BackColor == Color.Red)
            {
                txtEnterCID2.BackColor = Color.White;
                txtEnterCID2.ForeColor = Color.Black;
            }
            if (txtEnterCID2.TextLength == 6)
            {
                txtEnterCID3.Focus();
            }
        }

        private void TxtEnterCID3TextChanged(object sender, EventArgs e)
        {
            if (txtEnterCID3.BackColor == Color.Red)
            {
                txtEnterCID3.BackColor = Color.White;
                txtEnterCID3.ForeColor = Color.Black;
            }
            if (txtEnterCID3.TextLength == 6)
            {
                txtEnterCID4.Focus();
            }
        }

        private void TxtEnterCID4TextChanged(object sender, EventArgs e)
        {
            if (txtEnterCID4.BackColor == Color.Red)
            {
                txtEnterCID4.BackColor = Color.White;
                txtEnterCID4.ForeColor = Color.Black;
            }
            if (txtEnterCID4.TextLength == 6)
            {
                txtEnterCID5.Focus();
            }
        }

        private void TxtEnterCID5TextChanged(object sender, EventArgs e)
        {
            if (txtEnterCID5.BackColor == Color.Red)
            {
                txtEnterCID5.BackColor = Color.White;
                txtEnterCID5.ForeColor = Color.Black;
            }
            if (txtEnterCID5.TextLength == 6)
            {
                txtEnterCID6.Focus();
            }
        }

        private void TxtEnterCID6TextChanged(object sender, EventArgs e)
        {
            if (txtEnterCID6.BackColor == Color.Red)
            {
                txtEnterCID6.BackColor = Color.White;
                txtEnterCID6.ForeColor = Color.Black;
            }
            if (txtEnterCID6.TextLength == 6)
            {
                txtEnterCID7.Focus();
            }
        }

        private void TxtEnterCID7TextChanged(object sender, EventArgs e)
        {
            if (txtEnterCID7.BackColor == Color.Red)
            {
                txtEnterCID7.BackColor = Color.White;
                txtEnterCID7.ForeColor = Color.Black;
            }
            if (txtEnterCID7.TextLength == 6)
            {
                txtEnterCID8.Focus();
            }
        }

        private void TxtEnterCID8TextChanged(object sender, EventArgs e)
        {
            if (txtEnterCID8.BackColor != Color.Red)
            {
                return;
            }
            txtEnterCID8.BackColor = Color.White;
            txtEnterCID8.ForeColor = Color.Black;
        }
    }
}