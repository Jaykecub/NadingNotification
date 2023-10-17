using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
//using Word = Microsoft.Office.Interop.Word;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Net.Mail;
//using Microsoft.Office.Interop.Word;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Linq.Expressions;
//using Microsoft.Office.Core;

namespace NadingNotif
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            /*
             * Settings
             *      Update data
             *          Get data from email
             *              contact email
             *              
             */
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            
            /* On Compile Button Click
             *      Get file type from dropdown menu
             *          Load that file type from computer
             *              Read Input String I guess?
             *                  (Know exactly where they start?)
             *     Get 'patient' name from text box
             *     Get Date from text box
             *     
             *     Insert name into document
             *     Insert Date into document
             *     Save Document as seperate File
             *     
             *     If all goes correctly --> Compile Successful!
             *     
             *     
             * 
             * */


            /* On Send Button click
             *      Send Current File from compile button to correct place
             *              (email to correct person?)
             *      Maybe on button click send also have textbox for email?
             *      Dropdown menu for email?
             *              Settings --> import contact data from email
             * 
             */

        }
        
        private void guna2TextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            // COMPILE button

            // Gathering Data from textboxes
            string patient = guna2TextBox2.Text;
            string app_date = dateTimePicker1.Text;
            string email = guna2TextBox4.Text;
            string app_time = guna2TextBox3.Text;
            string hygienist = guna2TextBox1.Text;


            //queue.Text = name + " with email: " + email + " at " + app_time + " on " + app_date;

            // Updates 'console' 
            richTextBox1.Text = "";
            var box = richTextBox1;
            box.AppendText(patient, Color.Blue);
            box.AppendText(" with email: ");
            box.AppendText(email, Color.Green);
            box.AppendText(" at ");
            box.AppendText(app_time, Color.Red);
            box.AppendText(" on ");
            box.AppendText(app_date, Color.Red);
            box.AppendText(" with ");
            box.AppendText(hygienist, Color.Blue);


            if (string.IsNullOrEmpty(patient))
                return;

            // Create combined list for inputs
            ListViewItem lv = new ListViewItem(patient, 0);
            lv.SubItems.Add(email);
            lv.SubItems.Add(app_date);
            lv.SubItems.Add(app_time);
            lv.SubItems.Add(hygienist);

            listView1.Items.Add(lv);

            /*
            string savefile =  "C:\\Users\\123\\Desktop\\" + name + "testfile.docx";

            //string str = "Hello " + userName + ". Today is " + dateString + ".";


            //string name = guna2TextBox2.Text;
            CreateWordDocument("C:\\Users\\123\\Desktop\\TestWord.docx", savefile) ;
            */
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            // On Button Click for SEND
            //Should load textbox information from listview
            
            //if nothing selected:
            if (listView1.SelectedItems.Count == 0)
            {
                return;
            }

            string selectedname = listView1.SelectedItems[0].SubItems[0].Text;
            string selectedemail = listView1.SelectedItems[0].SubItems[1].Text;
            string selecteddate = listView1.SelectedItems[0].SubItems[2].Text;
            string selectedtime = listView1.SelectedItems[0].SubItems[3].Text;
            string selectedHyg = listView1.SelectedItems[0].SubItems[4].Text;

            // For when there is no entry for a certain selection of information to ensure no human mistakes in sending email
            if (selectedname == "")
             {
                if (MessageBox.Show("Name is missing... \nSend Email Anyways?", "Send Email Anyways?", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                {
                    return;
                }
             }
            if (selectedemail == "")
            {
                MessageBox.Show("Email is missing...", "Send Email Anyways?", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                    return;
                
            }
            if (selecteddate == "")
            {
                if (MessageBox.Show("Date is missing... \nSend Email Anyways?", "Send Email Anyways?", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                {
                    return;
                }
            }
            if (selectedtime == "")
            {
                if (MessageBox.Show("Time is missing... \nSend Email Anyways?", "Send Email Anyways?", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                {
                    return;
                }
            }
            if (selectedHyg == "")
            {
                if (MessageBox.Show("Hygentist is missing... \nSend Email Anyways?", "Send Email Anyways?", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                {
                    return;
                }
            }


            string first_name = selectedname;

            // for last names: 
            bool name_whitespace = selectedname.Contains(" ");
            if (name_whitespace)
            {
                first_name = selectedname.Substring(0, selectedname.IndexOf(' '));
            }


            // Update 'console' 
            richTextBox1.Text = "";
            var box = richTextBox1;
            box.AppendText(selectedname, Color.Blue);
            box.AppendText(" is being sent a recall email...");


            string email_body = "Hello " + first_name + ",\r\n \r\nThis is a reminder of your periodontal maintenance appointment, " + selecteddate + " at " + selectedtime +
                    ".  If you cannot keep this appointment, please call as soon as possible so this time can be reserved for another patient.\r\n \r\n" +
                    selectedHyg + " will be here to see you. Our office is in a transition phase for a while, but we will do our best to help you with your dental needs as we can." +
                    " \r\n \r\nPlease reply YES to let us know you will be here.\r\n \r\nThank you,  \r\nDr. Name's Office \r\n(123)456-7890;";

            string email_body_formatted = email_body.Insert(0, "<font size=16px>");
            email_body_formatted = email_body;


            // Obtains user email and password from 'login page' (first panel)
            string userEmail = UserEmail.Text;
            string password = UserPassword.Text;

            //email jumbos 
            try
            {
                MailMessage mail = new System.Net.Mail.MailMessage();
                SmtpClient smtp = new SmtpClient("smtp.gmail.com");
                mail.From = new MailAddress(userEmail);
                mail.To.Add(selectedemail);
                mail.Subject = "Dr. Nading's Office";
                mail.Body = email_body_formatted;
                // html formatting... i have it written for cs so just gonna disable for the moment... look at email_body_formatted to change context
                mail.IsBodyHtml = false;
               

                smtp.Port = 587;
                smtp.Credentials = new System.Net.NetworkCredential(userEmail, password);
                smtp.EnableSsl = true;
                smtp.Send(mail);
                MessageBox.Show("Mail has been successfully sent", "Email sent", MessageBoxButtons.OK, MessageBoxIcon.Information);

                /*
                MailMessage me = new System.Net.Mail.MailMessage();
                SmtpClient sc = new SmtpClient("smtp.gmail.com");
                me.From = new MailAddresss(from.text);

                client.Port = 587;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                System.Net.NetworkCredential credential = new System.Net.NetworkCredential("jacobnading@gmail.com", "RedPanda1995");
                client.EnableSsl = true;
                client.Credentials = credential;

                MailMessage message = new MailMessage("jacobnading@gmail.com", "jacobnading@gmail.com");
                message.Subject = "pando snackies";
                message.Body = "<h1>This is the mail body</h1>";
                message.IsBodyHtml = true;
                client.Send(message);
                */

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            


            //remove selection from list because the selection has been emailed (don't want to double email)
            listView1.SelectedItems[0].Remove();

        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            //Panel_List.Visible = true;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void guna2TextBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void guna2TextBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void guna2GradientPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            //Panel_List.Visible = false;
        }

        private void queue_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void guna2Button1_Click_1(object sender, EventArgs e)
        {
            // Remove button

            //if nothing selected:
            if (listView1.SelectedItems.Count == 0)
            {
                return;
            }

            //read selected item from listView1

            string selectedname = listView1.SelectedItems[0].SubItems[0].Text;
            string selectedemail = listView1.SelectedItems[0].SubItems[1].Text;
            string selecteddate = listView1.SelectedItems[0].SubItems[2].Text;
            string selectedtime = listView1.SelectedItems[0].SubItems[3].Text;
            string selectedHyg = listView1.SelectedItems[0].SubItems[4].Text;

            richTextBox1.Text = "";
            var box = richTextBox1;
            box.AppendText(selectedname, Color.Blue);
            box.AppendText(" has been removed from the list");

            listView1.SelectedItems[0].Remove();

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*      wasn't sure if these variables could be used elsewhere because private but only making panel invisible so doesn't delete textboxes
            string user_email = UserEmail.Text;
            string user_password = UserPassword.Text;
            */

            PanelStart.Visible = false;

        }

        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
public static class RichTextBoxExtensions
{
    public static void AppendText(this RichTextBox box, string text, Color color)
    {
        box.SelectionStart = box.TextLength;
        box.SelectionLength = 0;

        box.SelectionColor = color;
        box.AppendText(text);
        box.SelectionColor = box.ForeColor;
    }
}