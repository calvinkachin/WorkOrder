using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using System.Net.Mail;
using System.Net.Mail;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Timers;
using System.Drawing.Printing;
using System.Drawing.Text;
using Novacode;
using System.Web;



namespace WorkOrder
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void cmbWorkType_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (cmbWorkType.SelectedIndex == 4)
            {
                if (lblObservation.Visible == false)
                {
                    Util.Animate(lblObservation, Util.Effect.Slide, 100, 90);
                    Util.Animate(cmbObservation, Util.Effect.Slide, 100, 90);
                    lblObservation.Visible = true;
                    cmbObservation.Visible = true;
                }
            }
            else
            {
                if (lblObservation.Visible == true)
                {
                    Util.Animate(lblObservation, Util.Effect.Slide, 100, 90);
                    Util.Animate(cmbObservation, Util.Effect.Slide, 100, 90);
                    lblObservation.Visible = false;
                    cmbObservation.Visible = false;
                }
            }

            if (cmbWorkType.SelectedIndex == cmbWorkType.Items.Count - 1)
            {
                if (lblOtherWork.Visible == false)
                {
                    Util.Animate(lblOtherWork, Util.Effect.Slide, 100,0);
                    Util.Animate(txtOtherWork, Util.Effect.Slide, 100,0);
                    lblOtherWork.Visible = true;
                    txtOtherWork.Visible = true;
                }
            }
            else
            {
                if (lblOtherWork.Visible == true)
                {
                    Util.Animate(lblOtherWork, Util.Effect.Slide, 100, 0);
                    Util.Animate(txtOtherWork, Util.Effect.Slide, 100, 0);
                    lblOtherWork.Visible = false;
                    txtOtherWork.Visible = false;
                }
            }
        }


        private void ReportAdd()
        {
            if (txtSerial.Text == "") { txtSerial.Text = "N/A"; }

            if (txtComplaint.Text == "") { txtComplaint.Text = "N/A"; }
            if (txtFieldReport.Text == "") { txtFieldReport.Text = "N/A"; }

            ListViewItem lvi = new ListViewItem(txtSerial.Text);
            lvi.SubItems.Add(txtComplaint.Text);
            lvi.SubItems.Add(txtFieldReport.Text);
            lvi.SubItems.Add(cmbRFU.Text);
            lvi.SubItems.Add(cmbObservation.Text);
            lvi.SubItems.Add("N/A");

            listViewReport.Items.Add(lvi);

            txtSerial.Clear();
            txtComplaint.Clear();
            txtFieldReport.Clear();
            cmbRFU.SelectedIndex = -1;
        }

        private void DefaultAnswers()
        {
            txtAge.Text = "UNK";
            cmbGender.SelectedIndex = 2;
            txtTreatment.Text = "";
            datePicker.Text = DateTime.Today.ToString();
            cmbActionTaken.SelectedIndex = 2;
            
            txtSettings.Text = "";
        }

        private bool CheckPass(string complaint,string report)
        {
            if (complaint.ToLower().Contains("ecg"))
            {
                if (report.ToLower().Contains("lead") || report.ToLower().Contains("pads"))
                {
                    return true;
                }
                else
                {
                    MessageBox.Show("Complaint mentions \"ECG\", but Field report does not mention \"Lead\" or \"Pads\"");
                    return false;
                }
            }
            return true;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            

            if (chkMandatory.Checked == true)
            {
                if (txtSerial.Text == "")
                {
                    MessageBox.Show("No Serial number entered");
                    return;
                }

                if (cmbRFU.SelectedIndex < 0)
                {
                    MessageBox.Show("Please select a RFU status for this unit");
                    return;
                }

                if (CheckPass(txtComplaint.Text, txtFieldReport.Text) == false)
                {
                    return;
                }

                if (cmbObservation.Visible == true && cmbObservation.SelectedIndex < 0)
                {
                    MessageBox.Show("Please select your observation during your Defib Evaluation");
                    return;
                }
            }
                if (cmbObservation.SelectedIndex == 2)
                {
                    tabControl1.SelectedIndex = 1;
                    lblUnitSerial.Text = "Serial: " + txtSerial.Text;
                    DefaultAnswers();
                    grpPatient.Visible = true;

                    return;
                }
            
            ReportAdd();
            txtSerial.Focus();
        }

        private void listViewReport_DoubleClick(object sender, EventArgs e)
        {
            if (listViewReport.SelectedItems.Count == 0) return;
            listViewReport.SelectedItems[0].Remove();
        }

        private void chkSign_CheckedChanged(object sender, EventArgs e)
        {
            System.Drawing.Image Simage;

            if(chkSign.Checked== true)
            {
                Process.Start("Signature.exe");
                MessageBox.Show("Click OK once signature is submitted.");

                Simage = System.Drawing.Image.FromFile("Signature.png");


                pictureBox1.Image = Simage;

                //pictureBox1.Height = image.Height;
                //pictureBox1.Width = image.Width;

                pictureBox1.Show();
            }
            else
            {
                pictureBox1.Image.Dispose();
                //Simage.Dispose();
            }
        }

        private void txtFieldReport_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (chkMandatory.Checked == true)
                {
                    if (txtSerial.Text == "")
                    {
                        MessageBox.Show("No Serial number entered");
                        return;
                    }

                    if (cmbRFU.SelectedIndex < 0)
                    {
                        MessageBox.Show("Please select a RFU status for this unit");
                        return;
                    }

                    if (CheckPass(txtComplaint.Text, txtFieldReport.Text) == false)
                    {
                        return;
                    }

                    if (cmbObservation.Visible == true && cmbObservation.SelectedIndex < 0)
                    {
                        MessageBox.Show("Please select your observation during your Defib Evaluation");
                        return;
                    }
                }
                if (cmbObservation.SelectedIndex == 2)
                {
                    tabControl1.SelectedIndex = 1;
                    lblUnitSerial.Text = "Serial: " + txtSerial.Text;
                    DefaultAnswers();
                    grpPatient.Visible = true;

                    return;
                }

                ReportAdd();
                txtSerial.Focus();
            }
        }


        private void LoadNewWorkOrder()
        {
            StreamReader wo_input = new StreamReader("wo.txt");
            string wonumber = wo_input.ReadLine();
            wo_input.Close();



            txtWorkOrderNum.Text = wonumber;
            txtTimeIn.Text = DateTime.Now.ToString("HH:mm");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadNewWorkOrder();
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            if (listViewReport.Items.Count == 0)
            {
                MessageBox.Show("There are no items in your report! Remember to hit 'Add to Report'!");
                return;
            }

            if (txtCustID.Text == "" || txtCustAddress.Text == "" || txtContact.Text == "")
            {
                MessageBox.Show("Please fill in all mandatory Customer Information fields.");
                return;
            }

            if (txtTimeIn.Text==""||txtTimeOut.Text == "" || cmbWorkType.SelectedIndex < 0)
            {
                MessageBox.Show("Please fill in all mandatory Work Information fields.");
                return;
            }

            if (cmbWorkType.Text == "PM")
            {
                for (int i = 0; i < listViewReport.Items.Count; i++)
                {
                    DocX maint = DocX.Load("eseries_main.docx");
                    string outputmaint = Environment.CurrentDirectory + "\\Maintenance\\WO "+txtWorkOrderNum.Text+" - " + listViewReport.Items[i].SubItems[0].Text+" - " +" Maintenance.docx";
                    maint.ReplaceText("#serial#", listViewReport.Items[i].SubItems[0].Text);
                    maint.ReplaceText("#date#", DateTime.Today.ToString("MMMM dd yyyy"));
                    maint.ReplaceText("#custid#", txtCustID.Text);
                    maint.SaveAs(outputmaint);
                    maint.Dispose();
                }
            }


            DocX letter = DocX.Load("worktemplate.docx");
            DocX officeletter = DocX.Load("officetemplate.docx");
            //if(txtSR.Text.Contains("/")||txtSR.Text.Contains("\"||))
            string outputFileName =Environment.CurrentDirectory+"\\Work Orders\\WO " + txtWorkOrderNum.Text + " - "+txtCustID.Text+" - "+ DateTime.Today.ToString("MMMM dd yyyy")+".docx";
            string officeoutputFileName = Environment.CurrentDirectory + "\\Work Orders\\WO " + txtWorkOrderNum.Text + " - " + txtCustID.Text + " - " + DateTime.Today.ToString("MMMM dd yyyy") + " OFFICE USE.docx";

            // Perform the replace:
            letter.ReplaceText("#wo#", txtWorkOrderNum.Text);
            letter.ReplaceText("#date#", DateTime.Today.ToString("MMMM dd, yyyy"));
            letter.ReplaceText("#custid#", txtCustID.Text);
            letter.ReplaceText("#timein#", txtTimeIn.Text);
            letter.ReplaceText("#timeout#", txtTimeOut.Text);
            letter.ReplaceText("#po#", txtPO.Text);
            letter.ReplaceText("#address#", txtCustAddress.Text);

            officeletter.ReplaceText("#wo#", txtWorkOrderNum.Text);
            officeletter.ReplaceText("#date#", DateTime.Today.ToString("MMMM dd, yyyy"));
            officeletter.ReplaceText("#custid#", txtCustID.Text);
            officeletter.ReplaceText("#timein#", txtTimeIn.Text);
            officeletter.ReplaceText("#timeout#", txtTimeOut.Text);
            officeletter.ReplaceText("#po#", txtPO.Text);
            officeletter.ReplaceText("#address#", txtCustAddress.Text);

            if (txtOtherWork.Text == "")
            {
                letter.ReplaceText("#worktype#", cmbWorkType.Text);
                officeletter.ReplaceText("#worktype#", cmbWorkType.Text);
            }
            else
            {
                letter.ReplaceText("#worktype#", txtOtherWork.Text);
                officeletter.ReplaceText("#worktype#", txtOtherWork.Text);
            }

            letter.ReplaceText("#contact#", txtContact.Text);
            officeletter.ReplaceText("#contact#", txtContact.Text);

            for (int i = 0; i < listViewReport.Items.Count; i++)
            {
                letter.ReplaceText("#serial" + i.ToString() + "#", listViewReport.Items[i].SubItems[0].Text);
                letter.ReplaceText("#comp" + i.ToString() + "#", listViewReport.Items[i].SubItems[1].Text);
                letter.ReplaceText("#report"+i.ToString()+"#", listViewReport.Items[i].SubItems[2].Text);

                officeletter.ReplaceText("#serial" + i.ToString() + "#", listViewReport.Items[i].SubItems[0].Text);
                officeletter.ReplaceText("#comp" + i.ToString() + "#", listViewReport.Items[i].SubItems[1].Text);
                officeletter.ReplaceText("#report" + i.ToString() + "#", listViewReport.Items[i].SubItems[2].Text);
                officeletter.ReplaceText("#rfu" + i.ToString() + "#", listViewReport.Items[i].SubItems[3].Text);
                officeletter.ReplaceText("#use" + i.ToString() + "#", listViewReport.Items[i].SubItems[4].Text);
                officeletter.ReplaceText("#add" + i.ToString() + "#", listViewReport.Items[i].SubItems[5].Text);
            }

            for (int i = 0; i < 15; i++)
            {
                letter.ReplaceText("#serial" + i.ToString() + "#", "");
                letter.ReplaceText("#comp" + i.ToString() + "#", "");
                letter.ReplaceText("#report" + i.ToString() + "#", "");

                officeletter.ReplaceText("#serial" + i.ToString() + "#", "");
                officeletter.ReplaceText("#comp" + i.ToString() + "#", "");
                officeletter.ReplaceText("#report" + i.ToString() + "#", "");
                officeletter.ReplaceText("#rfu" + i.ToString() + "#", "");
                officeletter.ReplaceText("#use" + i.ToString() + "#", "");
                officeletter.ReplaceText("#add" + i.ToString() + "#", "");
            }

            if (chkSign.Checked == true)
            {
                var logo = letter.AddImage("Signature.png");
                var logo2 = officeletter.AddImage("Signature.png");
                Picture Image = logo.CreatePicture(100, 250);
                Picture Image2 = logo2.CreatePicture(100, 250);
                Paragraph p = letter.InsertParagraph("");
                Paragraph p2 = officeletter.InsertParagraph("");
                p.AppendPicture(Image);
                p2.AppendPicture(Image2);
            }
            letter.SaveAs(outputFileName);
            officeletter.SaveAs(officeoutputFileName);

            letter.Dispose();
            officeletter.Dispose();

            // Open in word:
            Process.Start("WINWORD.EXE", "\"" + outputFileName + "\"");
            Process.Start("WINWORD.EXE", "\"" + officeoutputFileName + "\"");
            
            if(chkEmail.Checked== true)
            {
                SendEmail(outputFileName);
            }

            try
            {
                StreamWriter wo_output = new StreamWriter("wo.txt");
                wo_output.WriteLine((Int32.Parse(txtWorkOrderNum.Text) + 1).ToString());
                wo_output.Close();
            }
            catch
            {
                MessageBox.Show("wo.txt is inaccessible. Work Order Number will not be updated.");
            }
            ClearForm();
            LoadNewWorkOrder();
        }

        private void cmbObservation_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbObservation.SelectedIndex == 2)
            {
                grpPatient.Visible = true;
            }
            else
            {
                grpPatient.Visible = false;
            }
        }

        private void txtSerial_TextChanged(object sender, EventArgs e)
        {
            if (txtSerial.Text.Length > 2)
            {
                if (txtSerial.Text[0].ToString() + txtSerial.Text[1].ToString() == "AB"||txtSerial.Text[0].ToString()=="T")
                {
                    cmbRFU.SelectedIndex = 2;
                }
            }
        }

        private void cmbObservation_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (cmbObservation.SelectedIndex == 2)
            {
                btnAdd.Text = "Add to Report\n(Requires additional info)";
            }
            else
            {
                btnAdd.Text = "Add to Report";
            }
        }

        private void cmbError_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbError.SelectedIndex == 0)
            {
                if (lblError.Visible == false)
                {
                    Util.Animate(lblError, Util.Effect.Slide, 100, 0);
                    Util.Animate(txtErrorMessages, Util.Effect.Slide, 100, 0);
                    lblError.Visible = true;
                    txtErrorMessages.Visible = true;
                }
            }
            else
            {
                if (lblError.Visible == true)
                {
                    Util.Animate(lblError, Util.Effect.Slide, 100, 0);
                    Util.Animate(txtErrorMessages, Util.Effect.Slide, 100, 0);
                    lblError.Visible = false;
                    txtErrorMessages.Visible = false;
                }
            }
        }

        private void ClearAdditionalInfo()
        {
            txtAge.Text = "";
            cmbGender.SelectedIndex = -1;
            txtTreatment.Text = "";
            datePicker.Value = DateTime.Today;
            txtSettings.Text = "";
            cmbActionTaken.SelectedIndex = -1;
            cmbAdverse.SelectedIndex = -1;
            cmbDuring.SelectedIndex = -1;

            cmbMalDupe.SelectedIndex = -1;
            cmbMalStatus.SelectedIndex = -1;
            cmbError.SelectedIndex = -1;
            cmbReportAvail.SelectedIndex = -1;
            cmbStripsAvail.SelectedIndex = -1;
            cmbDataAvail.SelectedIndex = -1;

            lblUnitSerial.Text = "";
        }

        private void btnAddAdditional_Click(object sender, EventArgs e)
        {
            if (chkMandatory.Checked == true)
            {
                if (txtSerial.Text == "") return;
            }

            if (txtComplaint.Text == "") { txtComplaint.Text = "N/A"; }
            if (txtFieldReport.Text == "") { txtFieldReport.Text = "N/A"; }

            string additional = "Patient Age: " + txtAge.Text + "\nGender: " + cmbGender.Text + "\nTreatment: " + txtTreatment.Text + "\nEvent Date: " + datePicker.Value.ToString("mm - dd - yyyy") + "\nAction taken: " + cmbActionTaken.Text + "\nAdverse Effects: " + cmbAdverse.Text + "\nDuring: " + cmbDuring.Text + "\nSettings used: " + txtSettings.Text + "\n\nMalfunction Status: " + cmbMalStatus.Text + "\nMalfunction Duplicated?: " + cmbMalDupe.Text + "\nError Messages?: " + cmbError.Text + "\nError Messages Found: " + txtErrorMessages.Text + "\nReport Available?: " + cmbReportAvail.Text + "\nPrint-Outs Available?: " + cmbStripsAvail.Text + "\nData File Available?: " + cmbDataAvail.Text + "\nAdditional Info: " + txtAdditional.Text; ;


            ListViewItem lvi = new ListViewItem(txtSerial.Text);
            lvi.SubItems.Add(txtComplaint.Text);
            lvi.SubItems.Add(txtFieldReport.Text);
            lvi.SubItems.Add(cmbRFU.Text);
            lvi.SubItems.Add(cmbObservation.Text);
            lvi.SubItems.Add(additional);

            listViewReport.Items.Add(lvi);

            txtSerial.Clear();
            txtComplaint.Clear();
            txtFieldReport.Clear();
            cmbRFU.SelectedIndex = -1;
            grpPatient.Visible = false;
            tabControl1.SelectedIndex = 0;
            ClearAdditionalInfo();
        }

        private void grpPatient_VisibleChanged(object sender, EventArgs e)
        {
            if(grpPatient.Visible== true)
            {
                lblQA.Visible = false;
                lblDisable.Visible = true;
                grp1.Visible = false;
                grp2.Visible = false;
                grp3.Visible = false;
            }
            else
            {
                lblQA.Visible = true;
                lblDisable.Visible = false;
                grp1.Visible = true;
                grp2.Visible = true;
                grp3.Visible = true;
            }
        }

        private void SendEmail(string attach)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.MailItem oMsg = (Microsoft.Office.Interop.Outlook.MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(txtEmail.Text);
                oRecip.Resolve();

                oMsg.Subject = "Work Order #" + txtWorkOrderNum.Text + " - Customer Copy";
                oMsg.Body = "Hi " + txtContact.Text.Split(' ')[0] + ",\n\nHere is a copy of Work Order #"+txtWorkOrderNum.Text+" to keep for your records.\n\nThanks,\n\n";
                oMsg.Attachments.Add(attach, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                oMsg.Display(true);
            }
            catch
            {
                MessageBox.Show("Email function failed. Make sure Outlook is open");
            }
        }

        private void chkEmail_CheckedChanged(object sender, EventArgs e)
        {
            if (chkEmail.Checked == true)
            {
                if (lblEmail.Visible == false)
                {
                    Util.Animate(lblEmail, Util.Effect.Slide, 90, 180);
                    Util.Animate(txtEmail, Util.Effect.Slide, 90, 180);
                    lblEmail.Visible = true;
                    txtEmail.Visible = true;
                }
                
            }
            else
            {
                if (lblEmail.Visible == true)
                {
                    Util.Animate(lblEmail, Util.Effect.Slide, 90, 180);
                    Util.Animate(txtEmail, Util.Effect.Slide, 90, 180);
                    lblEmail.Visible = false;
                    txtEmail.Visible = false;
                    txtEmail.Text = "";
                }
            }
        }

        private void ClearForm()
        {
            txtTimeIn.Text = DateTime.Now.ToString("HH:mm");
            txtTimeOut.Text = "";
            cmbWorkType.SelectedIndex = -1;
            txtOtherWork.Text = "";
            txtSerial.Text = "";
            txtComplaint.Text = "";
            cmbRFU.SelectedIndex = -1;
            cmbObservation.SelectedIndex = -1;
            txtFieldReport.Text = "";

            listViewReport.Items.Clear();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            ClearForm();
        }

        private void datePicker_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
