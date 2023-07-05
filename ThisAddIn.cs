using System;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Drawing;


namespace OutlookAddIn2
{
    public partial class ThisAddIn
    {
        // This method is called when the add-in starts up
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Add-in startup");
            // Register an event handler for the ItemLoad event of the Outlook application
            // This event is triggered whenever a new item (like a mail, appointment, etc.) is loaded in Outlook
            this.Application.ItemLoad += new Outlook.ApplicationEvents_11_ItemLoadEventHandler(Application_ItemLoad);
        }

        // This method is called when a new item is loaded in Outlook
        private void Application_ItemLoad(object Item)
        {
            // Check if the loaded item is a mail item
            if (Item is Outlook.MailItem)
            {
                // If it's a mail item, cast it to a MailItem object and show the form
                Outlook.MailItem mailItem = Item as Outlook.MailItem;
                ShowForm(mailItem);
            }
        }

        // This method creates a new form and shows it
        private void ShowForm(Outlook.MailItem mailItem)
        {
            // show a dialog box to select string
            using (Form form = new Form())
            {
                try
                {
                    // Set the properties of the form
                    form.Text = "Select option";
                    form.Font = new Font("Arial", 10);
                    form.FormBorderStyle = FormBorderStyle.FixedDialog;
                    form.StartPosition = FormStartPosition.CenterScreen;

                    // Read the lines from a text file
                    // Each line should contain an option and a color, separated by a comma
                    string[] lines = null;
                    string filePath = @"H:\list.txt"; // Replace with the path to your file
                    if (!System.IO.File.Exists(filePath))
                    {
                        filePath = @"C:\temp\list.txt"; // Use this file if the first one does not exist
                    }

                    if (System.IO.File.Exists(filePath))
                    {
                        lines = System.IO.File.ReadAllLines(filePath);
                        // rest of your code...
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("File not found: " + filePath);
                    }
                    int currentY = 10;
                    Panel[] panels = new Panel[lines.Length];
                    Panel selectedPanel = null; // This will keep track of the selected panel
                    LinkLabel selectedLinkLabel = null; // This will keep track of the selected LinkLabel

                    // Loop through the lines in the text file
                    for (int i = 0; i < lines.Length; i++)
                    {
                        // Split the line into parts
                        string[] parts = lines[i].Split(',');
                        string option = parts[0];
                        string color = parts.Length > 1 ? parts[1] : "Black";  // Default to black if no color is specified

                        LinkLabel linkLabel = new LinkLabel()
                        {
                            Text = option,
                            Location = new Point(5, 5),  // Adjusted for the offset within the panel
                            Font = new Font("Arial", 10),
                            ForeColor = Color.Black,
                            BackColor = Color.White,
                            AutoSize = true,
                            LinkBehavior = LinkBehavior.NeverUnderline,  // This will remove the underline
                            LinkColor = Color.Black,  // This will change the color of the link to black
                            ActiveLinkColor = Color.Black,  // This will change the color of the active link to black
                            VisitedLinkColor = Color.Black,  // This will change the color of the visited link to black
                            Cursor = Cursors.Default  // This will change the cursor to the default cursor
                        };                        

                        linkLabel.DoubleClick += (sender, e) =>
                        {
                            // This block of code is executed when the link label is double-clicked
                            // It gets the text and color of the selected option and inserts them into the body of the mail item
                            // Then it closes the form
                            string selectedString = null;
                            string selectedColor = null;

                            if (selectedPanel != null)
                            {
                                selectedString = selectedPanel.Controls.OfType<LinkLabel>().First()?.Text;
                                selectedColor = selectedPanel.Tag as string;  // Use the Tag property to store the color
                                                                              //System.Diagnostics.Debug.WriteLine($"Selected color is '{selectedColor}'");
                            }

                            if (selectedString != null)
                            {
                                // Use a timer to wait for the signature to be loaded
                                System.Timers.Timer timer = new System.Timers.Timer(500);
                                timer.Elapsed += (s, elapsedArgs) =>
                                {
                                    timer.Stop();

                                    //string selectedHtmlString = "<p>" + selectedString + "</p>";
                                    string selectedHtmlString = $"<b><font color='{selectedColor}' size='6'>{selectedString}</font></b>";  // Use the selected color
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

                        panel.Tag = color;  // Store the color in the Tag property
                        panel.DoubleClick += (sender, e) =>
                        {
                            // This block of code is executed when the panel is double-clicked
                            // It gets the text and color of the selected option and inserts them into the body of the mail item
                            // Then it closes the form

                            string selectedString = null;
                            string selectedColor = null;

                            if (selectedPanel != null)
                            {
                                selectedString = selectedPanel.Controls.OfType<LinkLabel>().First()?.Text;
                                selectedColor = selectedPanel.Tag as string;  // Use the Tag property to store the color
                                System.Diagnostics.Debug.WriteLine($"Selected color is '{selectedColor}'");
                            }

                            if (selectedString != null)
                            {
                                // Use a timer to wait for the signature to be loaded
                                System.Timers.Timer timer = new System.Timers.Timer(500);
                                timer.Elapsed += (s, elapsedArgs) =>
                                {
                                    timer.Stop();

                                    //string selectedHtmlString = "<p>" + selectedString + "</p>";
                                    string selectedHtmlString = $"<b><font color='{selectedColor}' size='6'>{selectedString}</font></b>";  // Use the selected color
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
                                selectedLinkLabel.BackColor = Color.White;
                            }

                            //if (selectedLinkLabel != null)
                            //{
                            //    selectedLinkLabel.BackColor = Color.White;
                            //}

                            System.Diagnostics.Debug.WriteLine("Panel clicked");

                            selectedPanel = panel;
                            selectedLinkLabel = panel.Controls.OfType<LinkLabel>().First();
                            string selectedColor = null;
                            selectedColor = selectedPanel.Tag as string;  // Use the Tag property to store the color
                            //selectedPanel.BackColor = Color.FromName("#FFDFD991");
                            selectedPanel.BackColor = ColorTranslator.FromHtml(selectedColor);  // Change the color to match the selected string
                            selectedLinkLabel.BackColor = ColorTranslator.FromHtml(selectedColor);
                            //selectedPanel.BackColor = Color.Blue;
                            System.Diagnostics.Debug.WriteLine($"Selected color is '{selectedColor}'");
                        };

                        linkLabel.Click += (sender, e) =>
                        {

                            // Also handle the click event on the label itself
                            // This block of code is executed when the link label is clicked
                            // It changes the background color of the panel to indicate that it is selected
                            // and resets the color of the previously selected panel

                            if (selectedPanel != null)
                            {
                                selectedPanel.BackColor = Color.White;
                            }

                            if (selectedLinkLabel != null)
                            {
                                selectedLinkLabel.BackColor = Color.White;
                            }

                            System.Diagnostics.Debug.WriteLine("LinkLabel clicked");

                            selectedPanel = panel;
                            selectedLinkLabel = linkLabel;
                            string selectedColor = null;
                            selectedColor = selectedPanel.Tag as string;  // Use the Tag property to store the color
                            //selectedPanel.BackColor = Color.FromName(selectedColor);
                            selectedPanel.BackColor = ColorTranslator.FromHtml(selectedColor);  // Change the color to match the selected string
                            selectedLinkLabel.BackColor = ColorTranslator.FromHtml(selectedColor);
                            //selectedPanel.BackColor = Color.Blue;
                            System.Diagnostics.Debug.WriteLine($"Selected color is '{selectedColor}'");
                        };
                    }

                    Button buttonOK = new Button() { Text = "OK", Location = new Point(10, currentY), Font = new Font("Arial", 10), BackColor = Color.LightGray, FlatStyle = FlatStyle.Flat };
                    form.Controls.Add(buttonOK);

                    // handle button click event
                    buttonOK.Click += (sender, e) =>
                    {
                        // This block of code is executed when the OK button is clicked
                        // It gets the text and color of the selected option and inserts them into the body of the mail item
                        // Then it closes the form
                        string selectedString = null;
                        string selectedColor = null;

                        if (selectedPanel != null)
                        {
                            selectedString = selectedPanel.Controls.OfType<LinkLabel>().First()?.Text;  // Use the text of the selected label
                            selectedColor = selectedPanel.Tag as string;
                        }

                        if (selectedString != null)
                        {
                            // Use a timer to wait for the signature to be loaded
                            System.Timers.Timer timer = new System.Timers.Timer(500); // Set interval to 500ms, you may need to adjust this value
                            timer.Elapsed += (s, elapsedArgs) =>
                            {
                                timer.Stop(); // Stop the timer

                                // Now you can add your string before or after the existing email body (which should contain the signature)
                                string selectedHtmlString = $"<b><font color='{selectedColor}' size='6'>{selectedString}</font></b>";  // Use the selected color
                                string existingBody = mailItem.HTMLBody;  // Get the existing body as HTML

                                mailItem.HTMLBody = selectedHtmlString + "<br>" + existingBody; // Add before

                            };
                            timer.Start(); // Start the timer

                        }

                        form.Close();
                    };

                    form.ShowDialog();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Exception: " + ex.Message);
                }
            }
        }

        private void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            
            if (Inspector.CurrentItem is Outlook.MailItem)
            {
                Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;

                
            }
        }


        #region VSTO generated code
        // This method is called by the VSTO runtime to initialize the add-in
        private void InternalStartup()
        {
            // Register an event handler for the Startup event of the add-in
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }
        #endregion
    }
}