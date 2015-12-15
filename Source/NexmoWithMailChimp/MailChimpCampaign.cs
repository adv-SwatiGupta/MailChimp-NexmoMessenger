using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MailChimp;
using MailChimp.Lists;
using MailChimp.Campaigns;
using MailChimp.Templates;
using MailChimp.Helper;
using CRMNexmo.Plugins;
using System.Xml;
using System.IO;
using MailChimp.Reports;
using System.Text.RegularExpressions;
using System.Net.Mail;
using System.Threading;

namespace NexmoWithMailChimp
{
    public partial class MailChimpCampaign : Form
    {
        MailChimpManager mc = null;
        SmsSender SmsSender = new SmsSender();
        string NexmoAPI = string.Empty, NexmoSecretKey = string.Empty, NexmoFromNumber = string.Empty;
        string MailchimpAPI = string.Empty;


        public MailChimpCampaign()
        {
            InitializeComponent();
            try
            {

                ReadSettings();
                if (mc != null)
                {
                    LoadCampaign();

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        public void ReadSettings()
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load("settings.xml");

                XmlNodeList nodeList = xmlDoc.DocumentElement.SelectNodes("/settings/nexmo");
                foreach (XmlNode node in nodeList)
                {
                    NexmoAPI = node.SelectSingleNode("api") != null ? node.SelectSingleNode("api").InnerText : "";
                    NexmoSecretKey = node.SelectSingleNode("secret-key") != null ? node.SelectSingleNode("secret-key").InnerText : "";
                    NexmoFromNumber = node.SelectSingleNode("from-number") != null ? node.SelectSingleNode("from-number").InnerText : "";
                }

                XmlNodeList mailchimpList = xmlDoc.DocumentElement.SelectNodes("/settings/mailchimp");
                foreach (XmlNode node in mailchimpList)
                {
                    MailchimpAPI = node.SelectSingleNode("api") != null ? node.SelectSingleNode("api").InnerText : "";
                }

                if (string.IsNullOrEmpty(NexmoAPI)
                      || string.IsNullOrEmpty(NexmoSecretKey)
                      || string.IsNullOrEmpty(NexmoFromNumber)
                      || string.IsNullOrEmpty(MailchimpAPI))
                {
                    OpenSettingsFrom();
                }
                else
                {
                    mc = new MailChimpManager(MailchimpAPI);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                OpenSettingsFrom();
            }

        }
        private void btnNextCampaign_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                if (ValidateFields())
                {
                    string campaignId = cmbCampaign.SelectedValue.ToString();

                    if (chkEnableSMS.Checked)
                    {
                        BindFields();
                        pnlNexmoMessage.Visible = true;
                        pnlCampaign.Visible = false;
                        //MailChimpCampaign campaign = new MailChimpCampaign();
                        //campaign.Height = 400;
                    }
                    else
                    {
                        SendCampaign();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor.Current = Cursors.Default;

            }
        }

        public bool ValidateFields()
        {
            if (cmbCampaign.SelectedIndex <= 0 || cmbCampaign.SelectedItem == null)
            {
                MessageBox.Show("Please select Campaign.", "Campaign", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cmbCampaign.Focus();
                return false;
            }

            if (string.IsNullOrEmpty(txtCampaignTitle.Text.Trim()))
            {
                MessageBox.Show("Please enter Campaign Title.", "Campaign Title", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtCampaignTitle.Focus();
                return false;
            }
            if (string.IsNullOrEmpty(txtCampaignSubject.Text.Trim()))
            {
                MessageBox.Show("Please enter Campaign Subject.", "Campaign Subject", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtCampaignSubject.Focus();
                return false;
            }
            if (string.IsNullOrEmpty(txtCampaignFromName.Text.Trim()))
            {
                MessageBox.Show("Please enter Campaign From Name.", "Campaign From Name", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtCampaignFromName.Focus();
                return false;
            }
            if (string.IsNullOrEmpty(txtCampaignFromEmail.Text.Trim()))
            {
                MessageBox.Show("Please enter Campaign From Email.", "Campaign From Email", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtCampaignFromEmail.Focus();
                return false;
            }
            else
            {
                if (txtCampaignFromEmail.Text.Trim().Length > 0)
                {
                    if (!Regex.IsMatch(txtCampaignFromEmail.Text.Trim(), @"\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)\Z", RegexOptions.IgnoreCase))
                    {
                        MessageBox.Show("Please provide valid value in 'Campaign From Email' field.", "Campaign From Email", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        txtCampaignFromEmail.SelectAll();
                        txtCampaignFromEmail.Focus();
                        return false;
                    }
                }
            }


            if (chkEnableSMS.Checked && pnlNexmoMessage.Visible)
            {
                if (cmbFieldPhone.SelectedIndex <= 0 || cmbFieldPhone.SelectedItem == null)
                {
                    MessageBox.Show("Please select recipient field.", "Recipient field", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    cmbFieldPhone.Focus();
                    return false;
                }
                if (string.IsNullOrEmpty(txtMessage.Text.Trim()))
                {
                    MessageBox.Show("Please enter Message.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtMessage.Focus();
                    return false;
                }
            }
            else
            {
                if (cmbCampaign.SelectedIndex > 0)
                {
                    bool isFoundPhoneField = false;
                    string listId = string.Empty;
                    string campaignId = cmbCampaign.SelectedValue.ToString();
                    CampaignFilter options = new CampaignFilter();
                    options.CampaignId = campaignId;
                    CampaignListResult campaign = mc.GetCampaigns(options);
                    foreach (var data in campaign.Data)
                    {
                        listId = data.ListId;
                        break;
                    }
                    IEnumerable<string> listEnum = new string[] { listId.ToString() };
                    MergeVarResult results = mc.GetMergeVars(listEnum);
                    foreach (var list in results.Data)
                    {
                        foreach (var mergeVars in list.MergeVars)
                        {
                            if (mergeVars.FieldType == "phone")
                            {
                                isFoundPhoneField = true;
                            }
                        }
                        if (!isFoundPhoneField)
                        {
                            MessageBox.Show("No Phone type field available in " + list.Name + "list.", "Phone field", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return false;
                        }
                    }
                }
            }
            return true;
        }
        public void LoadCampaign()
        {

            try
            {
                List<MailChimpList> mailChimpList = new List<MailChimpList>();
                mailChimpList.Add(new MailChimpList("0", "Select Campaign"));
                bool flag = true;
                int count = 0;

                while (flag)
                {

                    CampaignFilter filter = new CampaignFilter();
                    filter.Status = "save";
                    CampaignListResult lists = mc.GetCampaigns(filter, count, 100);
                    if (lists.Data.Count == 0)
                    {
                        flag = false;
                        break;
                    }
                    foreach (var list in lists.Data)
                    {
                        mailChimpList.Add(new MailChimpList(list.Id.ToString(), list.Title));
                    }
                    count++;
                }
                if (mailChimpList != null && mailChimpList.Count > 0)
                {
                    cmbCampaign.DisplayMember = "Name";
                    cmbCampaign.ValueMember = "Id";
                    cmbCampaign.DataSource = mailChimpList;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        public void OpenSettingsFrom()
        {
            this.Hide();
            Settings settings = new Settings();
            settings.ShowDialog();
            this.Close();

        }

        private void cmbCampaign_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbCampaign.SelectedIndex > 0 && cmbCampaign.SelectedItem != null)
            {
                string campaignId = cmbCampaign.SelectedValue.ToString();
                CampaignFilter options = new CampaignFilter();
                options.CampaignId = campaignId;
                CampaignListResult lists = mc.GetCampaigns(options);
                foreach (var data in lists.Data)
                {
                    txtCampaignTitle.Text = data.Title != null ? data.Title.Trim() : "";
                    txtCampaignSubject.Text = data.Subject != null ? data.Subject.Trim() : "";
                    txtCampaignFromName.Text = data.FromName != null ? data.FromName.Trim() : "";
                    txtCampaignFromEmail.Text = data.FromEmail != null ? data.FromEmail.Trim() : "";
                }
            }
            if (cmbCampaign.SelectedIndex == 0)
            {
                txtCampaignTitle.Text = "";
                txtCampaignSubject.Text = "";
                txtCampaignFromName.Text = "";
                txtCampaignFromEmail.Text = "";
            }
        }
        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Settings settings = new Settings();
            settings.ShowDialog();
            this.Close();
        }
        private void btnSendCampain_Click(object sender, EventArgs e)
        {
            SendCampaign();
        }
        public void SendCampaign()
        {
            if (ValidateFields())
            {
                var confirmResult = MessageBox.Show("Do you want to send Campaign ?",
                                                "Send Campaign",
                                                MessageBoxButtons.YesNo);
                if (confirmResult == DialogResult.Yes)
                {

                    string listId = string.Empty;
                    string campaignId = cmbCampaign.SelectedValue.ToString();
                    CampaignFilter options = new CampaignFilter();
                    options.CampaignId = campaignId;

                    //var updateResult = mc.UpdateCampaign(campaignId, "options", new
                    //{
                    //    title = txtCampaignTitle.Text,
                    //    subject = txtCampaignSubject.Text,
                    //    from_email = txtCampaignFromEmail.Text,
                    //    from_name = txtCampaignFromName.Text,
                    //});

                    Cursor.Current = Cursors.WaitCursor;
                    try
                    {
                        CampaignActionResult result = mc.SendCampaign(campaignId);
                        if (result.Complete)
                        {
                            if (chkEnableSMS.Checked)
                            {
                                bool flag = true;
                                int count = 0;
                                while (flag)
                                {
                                    List<EmailParameter> emails = new List<EmailParameter>();
                                    SentToLimits opt = new SentToLimits();
                                    opt.Status = "sent";
                                    opt.Start = count;
                                    opt.Limit = 100;

                                    bool isCampaignSend = false;
                                    while (!isCampaignSend)
                                    {
                                        CampaignFilter filter = new CampaignFilter();
                                        filter.CampaignId = campaignId;
                                        CampaignListResult lists = mc.GetCampaigns(filter);
                                        foreach (var l in lists.Data)
                                        {
                                            if (l.Status.Trim() == "sent")
                                            {
                                                isCampaignSend = true;
                                                SentToMembers results = mc.GetReportSentTo(campaignId, opt);
                                                if (results.Data.Count == 0)
                                                {
                                                    flag = false;
                                                    break;
                                                }
                                                foreach (var i in results.Data)
                                                {
                                                    string message = txtMessage.Text.Trim();
                                                    string phone = string.Empty;
                                                    foreach (var j in i.Member.MemberMergeInfo)
                                                    {
                                                        if (message.IndexOf(j.Key.Trim()) != -1)
                                                        {

                                                            message = message.Replace("*|" + j.Key.Trim() + "|*", j.Value != null ? j.Value.ToString() : "");
                                                        }
                                                        if (cmbFieldPhone.SelectedIndex > 0 && cmbFieldPhone.SelectedItem != null)
                                                        {
                                                            if (cmbFieldPhone.Text.Trim() == j.Key.ToString())
                                                            {
                                                                if (j.Value != null && !string.IsNullOrEmpty(j.Value.ToString()))
                                                                {
                                                                    phone = j.Value.ToString();
                                                                }
                                                            }
                                                        }
                                                    }
                                                    if (message.IndexOf("EMAIL") != -1)
                                                    {
                                                        message = message.Replace("*|EMAIL|*", i.Member.Email.ToString());
                                                    }
                                                    if (!string.IsNullOrEmpty(phone))
                                                    {
                                                        string smsResult = SmsSender.SendSMS(phone.Trim(), NexmoFromNumber, NexmoAPI, NexmoSecretKey, Uri.EscapeUriString(message));
                                                    }
                                                }
                                                break;
                                            }
                                        }
                                    }

                                    count++;
                                }
                            }
                        }
                        MessageBox.Show("Campaign sent successfully");
                        this.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        Cursor.Current = Cursors.Default;
                    }
                }
            }
        }

        public void BindFields()
        {

            List<MailChimpList> mailChimpField = new List<MailChimpList>();
            List<MailChimpList> mailChimpPhoneField = new List<MailChimpList>();
            mailChimpPhoneField.Add(new MailChimpList("0", "Select recipient field"));
            string listId = string.Empty;
            string campaignId = cmbCampaign.SelectedValue.ToString();
            CampaignFilter options = new CampaignFilter();
            options.CampaignId = campaignId;
            CampaignListResult lists = mc.GetCampaigns(options);

            foreach (var data in lists.Data)
            {
                listId = data.ListId;
                break;
            }
            //--member of list
            List<EmailParameter> emails = new List<EmailParameter>();
            IEnumerable<string> listEnum = new string[] { listId.ToString() };
            MergeVarResult results = mc.GetMergeVars(listEnum);
            foreach (var list in results.Data)
            {
                foreach (var mergeVars in list.MergeVars)
                {
                    mailChimpField.Add(new MailChimpList(mergeVars.Id.ToString(), mergeVars.Tag));

                    if (mergeVars.FieldType == "phone")
                    {
                        mailChimpPhoneField.Add(new MailChimpList(mergeVars.Id.ToString(), mergeVars.Tag));
                    }
                }
            }

            if (mailChimpField != null && mailChimpField.Count > 0)
            {
               
                lstboxFields.Items.Clear();
                lstboxFields.Items.AddRange(mailChimpField.ToArray());
                lstboxFields.DisplayMember = "Name";

                cmbFieldPhone.DisplayMember = "Name";
                cmbFieldPhone.ValueMember = "Id";
                cmbFieldPhone.DataSource = mailChimpPhoneField;
            }

        }
        private void txtFromNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar) || (e.KeyChar == (char)Keys.Back)))
                e.Handled = true;
        }
     

        private void chkEnableSMS_CheckedChanged(object sender, EventArgs e)
        {
            if (chkEnableSMS.Checked)
            {
                btnNextCampaign.Text = "Next";
            }
            else
            {
                btnNextCampaign.Text = "Send Campaign";
            }
        }

        private void MailChimpCampaign_Shown(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;
            bool isOpen = false;
            foreach (Form frm in fc)
            {
                if (frm.Name == "Settings")
                {
                    frm.Hide();
                }
            }

        }

        private void lstboxFields_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            var insertText = "*|" + lstboxFields.Text + "|*";
            var selectionIndex = txtMessage.SelectionStart;
            txtMessage.Text = txtMessage.Text.Insert(selectionIndex, insertText);
            txtMessage.SelectionStart = selectionIndex + insertText.Length;
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            pnlNexmoMessage.Visible = false;
            pnlCampaign.Visible = true;
        }
    }
}
