using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SPCWBConsolidatedReport
{
    public partial class TraceabilitySelection : Form
    {
        // RadioButton radioButton = new RadioButton();
        FlowLayoutPanel pnl = new FlowLayoutPanel();
        public static string TraceCat;
        private string _value { get; set; }

        public TraceabilitySelection()
        {
            InitializeComponent();
        }
        public TraceabilitySelection(DataTable dttracecat)
        {
           // RadioButton radioButton = new RadioButton();
            InitializeComponent();
           
            pnl.Dock = DockStyle.Fill;
            foreach (System.Data.DataRow dtTracerow in dttracecat.Rows)
             {
                 TraceCat = (dtTracerow["TraceCat"]).ToString();
                pnl.Controls.Add(new RadioButton() { Text =TraceCat });
                
             }

            this.Controls.Add(pnl);

        }
        bool ContainsDynamicRadioButtons(Panel panel)
        {
         
              foreach (Control control in panel.Controls)
                {
                    if (control is RadioButton && control.Name.StartsWith(TraceCat))
                    {
                        // Found a dynamic radio button
                        return true;
                    }
                }
            

            // No dynamic radio buttons found
            return false;
        }
        private void radioButton_CheckedChanged(object sender,EventArgs e)
        {
            // RadioButton radiobutton = new RadioButton();
            //RadioButton selectedRadioButton = (RadioButton)sender;
            //if (selectedRadioButton.Checked)
            //    this._value = radioButton.Text;
        }

        //private void radioButton_CheckedChanged(object sender, EventArgs e)
        //{
        //  //  RadioButton radioButton = (RadioButton)sender;
           
        //    // Assign the radio button text as value Ex: AAA
        //}

        public string GetValue()
        {
            RadioButton rbSelected = pnl.Controls
                       .OfType<RadioButton>()
                       .FirstOrDefault(r => r.Checked);
            if (rbSelected!=null)
            {
                this._value = rbSelected.Text;
            }
           return this._value;
         }

        private void TraceabilitySelection_Load(object sender, EventArgs e)
        {
           bool checkradiobuttonpresent= ContainsDynamicRadioButtons(pnl);
            if(checkradiobuttonpresent==false)
            {
                //DialogResult dialogResult = MessageBox.Show("There are no traceabilities defined in this file. This file and its data will not be exported.Do you want to continue the export?", "Traceability Selection", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                //if (dialogResult == DialogResult.Yes)
                //{
                //    this.Close();
                //}
                //else if (dialogResult == DialogResult.No)
                //{
                //    return;
                //}
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            return;
            //this.Close();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            RadioButton rbSelected = pnl.Controls
                      .OfType<RadioButton>()
                      .FirstOrDefault(r => r.Checked);
            //if (rbSelected == null)
            //{
            //    DialogResult dialogResult= MessageBox.Show("There are no traceabilities defined in this file. This file and its data will not be exported.Do you want to continue the export?", "Traceability Selection", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            //    if(dialogResult==DialogResult.Yes)
            //    {

            //    }
            //    else if(dialogResult == DialogResult.No)
            //    {

            //    }

            //}

            //radiobutton.CheckedChanged += RadioButton_CheckedChanged;
            this.Close();
        }
    }
}
