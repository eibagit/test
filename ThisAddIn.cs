using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Drawing;

namespace OutlookAddIn2
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.Inspectors.NewInspector += Inspectors_NewInspector;
        }


        private void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            if (Inspector.CurrentItem is Outlook.MailItem)
            {
                Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;

                

                // show a dialog box to select string
                using (Form form = new Form())
                {
                    form.Text = "Select option";
                    form.Font = new Font("Arial", 10);
                    form.FormBorderStyle = FormBorderStyle.FixedDialog;
                    form.StartPosition = FormStartPosition.CenterScreen;

                    string[] lines = System.IO.File.ReadAllLines(@"C:\temp\list.txt"); // Replace with the path to your file
                    int currentY = 10;
                    Panel[] panels = new Panel[lines.Length];
                    Panel selectedPanel = null;  // This will keep track of the selected panel

                    for (int i = 0; i < lines.Length; i++)
                    {
                        LinkLabel linkLabel = new LinkLabel()
                        {
                            Text = lines[i],
                            Location = new Point(5, 5),  // Adjusted for the offset within the panel
                            Font = new Font("Arial", 10),
                            ForeColor = Color.Black,
                            BackColor = Color.White,
                            AutoSize = true,
                            LinkBehavior = LinkBehavior.NeverUnderline,  // This will remove the underline
                            LinkColor = Color.Black  // This will change the color of the link to black
                        };

                        Panel panel = new Panel()
                        {
                            Location = new Point(10, currentY),
                            Size = new Size(form.ClientSize.Width - 20, 30),  // Adjust the size as needed
                            BorderStyle = BorderStyle.FixedSingle,
                            BackColor = Color.White
                        };
                        panel.Controls.Add(linkLabel);
                        panels[i] = panel;
                        form.Controls.Add(panel);
                        currentY += 40; // Add 40 to Y location for the next option, adjusted for the size of the panel

                        panel.Click += (sender, e) =>
                        {
                            // When a panel is clicked, change its background color
                            // and reset the color of the previously selected panel

                            if (selectedPanel != null)
                            {
                                selectedPanel.BackColor = Color.White;
                            }

                            selectedPanel = panel;
                            selectedPanel.BackColor = Color.Blue;
                        };

                        linkLabel.Click += (sender, e) =>
                        {
                            // Also handle the click event on the label itself

                            if (selectedPanel != null)
                            {
                                selectedPanel.BackColor = Color.White;
                            }

                            selectedPanel = panel;
                            selectedPanel.BackColor = Color.Blue;
                        };
                    }

                    Button buttonOK = new Button() { Text = "OK", Location = new Point(10, currentY), Font = new Font("Arial", 10), BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat };
                    form.Controls.Add(buttonOK);

                    // handle button click event
                    buttonOK.Click += (sender, e) =>
                    {
                        string selectedString = null;

                        if (selectedPanel != null)
                        {
                            selectedString = selectedPanel.Controls.OfType<LinkLabel>().First()?.Text;  // Use the text of the selected label
                        }

                        if (selectedString != null)
                        {
                            // Use a timer to wait for the signature to be loaded
                            System.Timers.Timer timer = new System.Timers.Timer(500); // Set interval to 500ms, you may need to adjust this value
                            timer.Elapsed += (s, elapsedArgs) =>
                            {
                                timer.Stop(); // Stop the timer

                                // Now you can add your string before or after the existing email body (which should contain the signature)
                                string selectedHtmlString = "<b><font color='red' size='6'>" + selectedString + "</font></b>";  // HTML formatted string
                                string existingBody = mailItem.HTMLBody;  // Get the existing body as HTML

                                mailItem.HTMLBody = selectedHtmlString + "<br>" + existingBody; // Add before

                                // mailItem.To = ""; // Clear the To field
                                // var editor = mailItem.GetInspector.WordEditor;
                                //mailItem.To = "";
                                //mailItem.GetInspector.Activate();
                                //mailItem.GetInspector.WordEditor.Range(1, 0).Select();

                                //mailItem.GetInspector.WordEditor.Range(mailItem.GetInspector.WordEditor.Content. - 1, mailItem.GetInspector.WordEditor.Content.End - 1).Select();
                                mailItem.GetInspector.WordEditor.Range(1, 0).Select();


                                // Display the MailItem
                                mailItem.Display(false);  // This should put the focus in the "To" field by default
                            };
                            timer.Start(); // Start the timer

                        }

                        form.Close();
                    };

                    form.ShowDialog();
                }
            }
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }
        
        #endregion
    }
}
