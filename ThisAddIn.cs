using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace OutlookAddIn2
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.Inspectors.NewInspector += Inspectors_NewInspector;
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);

        [DllImport("user32.dll")]
        static extern bool SetFocus(IntPtr hWnd);

        public void SetFocusToField(String windowTitle)
        {
            // Find the window
            IntPtr outlookWindow = FindWindow(null, windowTitle);
            Debug.Write(outlookWindow);
            // Find the child window
            IntPtr toField = FindWindowEx(outlookWindow, IntPtr.Zero, null, "To");

            // Set focus
            SetFocus(toField);
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

                        linkLabel.DoubleClick += (sender, e) =>
                        {
                            string selectedString = null;

                            if (selectedPanel != null)
                            {
                                selectedString = selectedPanel.Controls.OfType<LinkLabel>().First()?.Text;
                            }

                            if (selectedString != null)
                            {
                                // Use a timer to wait for the signature to be loaded
                                System.Timers.Timer timer = new System.Timers.Timer(500);
                                timer.Elapsed += (s, elapsedArgs) =>
                                {
                                    timer.Stop();

                                    //string selectedHtmlString = "<p>" + selectedString + "</p>";
                                    string selectedHtmlString = "<b><font color='red' size='6'>" + selectedString + "</font></b>";  // HTML formatted string
                                    string existingBody = mailItem.HTMLBody;
                                    mailItem.HTMLBody = selectedHtmlString + "<br>" + existingBody; // Add before

                                    // Display the MailItem
                                    //mailItem.Display(false);  // This should put the focus in the "To" field by default
                                    
                                };

                                timer.Start();
                            }

                            form.Close();
                        };

                        Panel panel = new Panel()
                        {
                            Location = new Point(10, currentY),
                            Size = new Size(form.ClientSize.Width - 20, 30),  // Adjust the size as needed
                            BorderStyle = BorderStyle.FixedSingle,
                            BackColor = Color.White
                        };

                        panel.DoubleClick += (sender, e) =>
                        {
                            string selectedString = null;

                            if (selectedPanel != null)
                            {
                                selectedString = selectedPanel.Controls.OfType<LinkLabel>().First()?.Text;
                            }

                            if (selectedString != null)
                            {
                                // Use a timer to wait for the signature to be loaded
                                System.Timers.Timer timer = new System.Timers.Timer(500);
                                timer.Elapsed += (s, elapsedArgs) =>
                                {
                                    timer.Stop();

                                    //string selectedHtmlString = "<p>" + selectedString + "</p>";
                                    string selectedHtmlString = "<b><font color='red' size='6'>" + selectedString + "</font></b>";  // HTML formatted string
                                    string existingBody = mailItem.HTMLBody;
                                    mailItem.HTMLBody = selectedHtmlString + "<br>" + existingBody; // Add before

                                    // Display the MailItem
                                    //mailItem.Display(false);  // This should put the focus in the "To" field by default

                                };

                                timer.Start();
                            }

                            form.Close();
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

                                //mailItem.To = "";
                                //mailItem.Display(false);
                                //mailItem.GetInspector.Activate();
                                //mailItem.GetInspector.WordEditor.Application.ActiveWindow.Selection.MoveRight(1);


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
