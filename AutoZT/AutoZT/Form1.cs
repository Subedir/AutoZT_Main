using System;
using System.Windows.Forms;

namespace AutoZT
{
    public partial class Form1 : Form
    {
        private TagListFile aFile = null;


        public Form1()
        {
            InitializeComponent();
            //disable all controls
            //tabConTrolSetup.Enabled = false;
     
            //Blackbox Tab

            //site info
            txtAONumberBox.Enabled = false;
            txtSiteBox.Enabled = false;

            //datalogging software
            comboSoftware.Enabled = false;
            comboSoftware.Visible = false;
            
            //IFix 
            txtChannelBox.Enabled = false;
            txtDatabasebox.Enabled = false;
            txtPLCBox.Enabled = false;
            buttonIfixDatabase.Enabled = false;
            buttonIGS.Enabled = false;
            buttonIfixScript.Enabled = false;
            buttonIfixScript.Visible = false;
            comboDriverBox.Enabled = false;

            //OPC
            buttonOPCFile.Enabled = false;
            txtChannelBoxOPC.Enabled = false;
            txtPLCBoxOPC.Enabled = false;

            //hide controls not used

            //IFix
            txtChannelBox.Visible = false;
            txtDatabasebox.Visible = false;
            txtPLCBox.Visible = false;
            buttonIfixDatabase.Visible = false;
            buttonIfixScript.Visible = false;
            comboDriverBox.Visible = false;
            
            //IGS
            buttonIGS.Visible = false;
            
            
            //OPC
            buttonOPCFile.Visible = false;
            txtChannelBoxOPC.Visible = false;
            txtPLCBoxOPC.Visible = false;
            
            //labels
            labelChannelName.Visible = false;
            labelDatabaseName.Visible = false;
            labelPLCName.Visible = false;
            //OPC
            labelCHNameOPC.Visible = false;
            labelPLCNameOPC.Visible = false;

           
            //Database Tab
            comboSoftwareBox.Enabled = false;
            comboSiteAssignedBox.Enabled = false;
            txtAObox.Enabled = false;
            txtSiteNameBox.Enabled = false;
            NoCassPerTrain.Enabled = false;
            NoModulesPerCass.Enabled = false;
            NoAreaPerModule.Enabled = false;
            comboflowrateUnitsBox.Enabled = false;
            comboTemperatureBox.Enabled = false;
            buttonSqlScript.Enabled = false;
            radioSquareFeet.Enabled = false;
            radioSquareMetres.Enabled = false;
      
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Application.Exit();
        }
        /// <summary>
        /// Opens an Excel file and enables controls.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openExcelFileStripMenuItem_Click(object sender, EventArgs e)
        {
            //openExcelFile.Filter = "Excel file (*.xls)|*.xls";
            openExcelFile.Filter = "Excel file (*.xls;*.xlsx)|*.xls;*xlsx|All files (*.*)|*.*";
            System.Windows.Forms.DialogResult result;

            result = openExcelFile.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    aFile = new TagListFile(openExcelFile.FileName);
                    aFile.ReadDataFromExcelToDataTable();
                    //dataGridView1.DataSource = aFile.m_excelData;                                                   

                    //site info
                    txtAONumberBox.Enabled = true;
                    txtAONumberBox.Focus();
                    txtSiteBox.Enabled = true;

                    //datalogging software
                    comboSoftware.Enabled = true;
                    comboSoftware.Visible = true;
                    //comboDriverBox.Enabled = true;
                    //comboDriverBox.Visible = true;
                    //Blackbox Tab
                    txtChannelBox.Enabled = false;
                    txtDatabasebox.Enabled = false;
                    txtPLCBox.Enabled = false;
                    buttonIfixDatabase.Enabled = false;
                    buttonIGS.Enabled = false;
                    buttonIfixScript.Enabled = false;
                    buttonIfixScript.Visible = false;
                    buttonOPCFile.Enabled = false;


                    //hide controls not used
                    txtChannelBox.Visible = false;
                    txtDatabasebox.Visible = false;
                    txtPLCBox.Visible = false;
                    buttonIGS.Visible = false;
                    buttonIfixDatabase.Visible = false;
                    buttonIfixScript.Visible = false;
                    
                    //labels
                    labelChannelName.Visible = false;
                    labelDatabaseName.Visible = false;
                    labelPLCName.Visible = false;
                 
                                      
                    //Activate Database fields
                    //Database Tab
                    comboSoftwareBox.Enabled = true;
                    comboSiteAssignedBox.Enabled = true;
                    txtAObox.Enabled = true;
                    txtSiteNameBox.Enabled = true;
                    NoCassPerTrain.Enabled = true;
                    NoModulesPerCass.Enabled = true;
                    NoAreaPerModule.Enabled = true;
                    comboflowrateUnitsBox.Enabled = true;
                    comboTemperatureBox.Enabled = true;
                    buttonSqlScript.Enabled = true;
                    radioSquareFeet.Enabled = true;
                    radioSquareMetres.Enabled = true;
                    
                }
                

                catch (System.IO.FileNotFoundException fnfe)
                {
                    MessageBox.Show("File does not exist!\r\n\r\n" + fnfe.Message, "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An Error Occured\r\n\r\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
           
        }

       
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox ab = new AboutBox();
            ab.ShowDialog();
        }

        private void buttonIFIX_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtSiteBox.Text))
            {
                MessageBox.Show("Please enter a site name.");
                txtSiteBox.Focus();
                return;

            }
            
            if (string.IsNullOrEmpty((string)comboDriverBox.SelectedItem ))
            {
                MessageBox.Show("Please select the driver.");
                comboDriverBox.Focus();
                return;
            }
            else
            {
                if (comboDriverBox.SelectedItem.ToString() == "IGS")
                {
                    if (string.IsNullOrEmpty(txtChannelBox.Text))
                    {
                        MessageBox.Show("Please enter the Channel name for the IGS driver.");
                        txtChannelBox.Focus();
                        return;
                    }
                    if (string.IsNullOrEmpty(txtPLCBox.Text))
                    {
                        MessageBox.Show("Please enter the PLC name for the IGS driver.");
                        txtPLCBox.Focus();
                        return;
                    }
                    if (string.IsNullOrEmpty(txtDatabasebox.Text))
                    {
                        MessageBox.Show("Please enter the IFIX Database name.");
                        txtDatabasebox.Focus();
                        return;
                    }
                    else if (txtDatabasebox.Text.Length > 8)
                    {
                        MessageBox.Show("The maximum database name length can be only 8 characters long. Please type a new name.");
                        txtDatabasebox.Focus();
                        return;

                    }
                }
                else if (comboDriverBox.SelectedItem.ToString() == "GE9")
                {
                    if (string.IsNullOrEmpty(txtPLCBox.Text))
                    {
                        MessageBox.Show("Please enter the PLC name for the GE9 driver.");
                        txtPLCBox.Focus();
                        return;
                    }
                    if (string.IsNullOrEmpty(txtDatabasebox.Text))
                    {
                        MessageBox.Show("Please enter the IFIX Database name.");
                        txtDatabasebox.Focus();
                        return;
                    }
                    else if (txtDatabasebox.Text.Length > 8)
                    {
                        MessageBox.Show("The maximum database name length can be only 8 characters long. Please type a new name.");
                        txtDatabasebox.Focus();
                        return;

                    }
                }

            }

            saveFile.FileName = txtAONumberBox.Text + txtSiteBox.Text + "IfixDatabase";
            saveFile.Filter = "CSV File (*.csv)|*.csv";
            System.Windows.Forms.DialogResult result;
           
            result = saveFile.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    
                    aFile.SaveIfixDataTableToCsvFile(saveFile.FileName,txtDatabasebox.Text, comboDriverBox.SelectedItem.ToString(),txtChannelBox.Text,txtPLCBox.Text);
                    //enable ifix script button
                    buttonIfixDatabase.Enabled = false;
                    buttonIfixDatabase.Visible = false;
                    buttonIfixScript.Enabled = true;
                    buttonIfixScript.Visible = true;
                    buttonIfixScript.Focus();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An Error Occured\r\n\r\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }              
            }
            

        }

        private void buttonIGS_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtSiteBox.Text))
            {
                MessageBox.Show("Please enter a site name.");
                txtSiteBox.Focus();
                return;

            }
            if (string.IsNullOrEmpty((string)comboSoftware.SelectedItem))
            {
                MessageBox.Show("Please select the software used.");
                comboSoftware.Focus();
                return;
            }

            saveFile.FileName = txtAONumberBox.Text + txtSiteBox.Text + "IGSDriver";
            saveFile.Filter = "CSV file (*.csv)|*.csv";
            System.Windows.Forms.DialogResult result;
           

            result = saveFile.ShowDialog();
            

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    aFile.SaveIGSDataTableToCsvFile(saveFile.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An Error Occured\r\n\r\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            
        }

        private void buttonIfixScript_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtSiteBox.Text))
            {
                MessageBox.Show("Please enter a site name.");
                txtSiteBox.Focus();
                return;

            }
            saveFile.FileName = txtAONumberBox.Text + txtSiteBox.Text + "IfixScript";
            saveFile.Filter = "text file (*.txt)|*.txt";
            System.Windows.Forms.DialogResult result;

            result = saveFile.ShowDialog();
            
            

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    aFile.WriteIfixScriptToTextFile(saveFile.FileName);
               
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An Error Occured\r\n\r\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void buttonSqlScript_Click(object sender, EventArgs e)
        {
            //AO can be blank as there are sites that do not have an AO number
            //if (string.IsNullOrEmpty(txtAObox.Text))
            //{
            //    MessageBox.Show("This site does not have an AO number.");
            //    txtAObox.Focus();
                //return;
                
            //}
            
           
            if (string.IsNullOrEmpty((string)comboSoftwareBox.SelectedItem))
            {
                MessageBox.Show("Please select the software used.");
                comboSoftwareBox.Focus();
                return;
            }
            if (string.IsNullOrEmpty((string)comboSiteAssignedBox.SelectedItem))
            { 
                MessageBox.Show("Please select the person the site is assigned to.");
                comboSiteAssignedBox.Focus();
                return;
            }

            if (string.IsNullOrEmpty(txtSiteNameBox.Text))
            {
                MessageBox.Show("Please enter a site name.");
                txtSiteNameBox.Focus();
                return;

            }

            if (string.IsNullOrEmpty((string)comboTemperatureBox.SelectedItem))
            {
                MessageBox.Show("Please select the unit of temperature for this site.");
                comboTemperatureBox.Focus();
                return;
            }

            if (string.IsNullOrEmpty((string)comboflowrateUnitsBox.SelectedItem))
            {
                MessageBox.Show("Please select the unit of the flow rates for this site.");
                comboflowrateUnitsBox.Focus();
                return;
            }
         
            else if (comboflowrateUnitsBox.SelectedItem.ToString() != "No FlowRates")
            {

                if (NoCassPerTrain.Value == 0 )
                {
                    MessageBox.Show("Please enter the number of cassettes per train.");
                    NoCassPerTrain.Focus();
                    return;
                }
                if (NoModulesPerCass.Value == 0)
                {
                    MessageBox.Show("Please enter the number of modules per cassette.");
                    NoModulesPerCass.Focus();
                    return;
                }
                if (NoAreaPerModule.Value == 0)
                {
                    MessageBox.Show("Please enter the area per module.");
                    NoAreaPerModule.Focus();
                    return;
                }
                if (radioSquareFeet.Checked == false && radioSquareMetres.Checked == false)
                {
                    MessageBox.Show("Please select the units of the area.");
                    radioSquareFeet.Focus();
                    return;
                }
                    
            }
          
              
            saveFile.FileName = txtAObox.Text + txtSiteNameBox.Text + "SQLDatabaseScript";
            saveFile.Filter = "SQL file (*.sql)|*.sql";
            System.Windows.Forms.DialogResult result;

            result = saveFile.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    aFile.WriteSQLDatabaseScriptToTextFile(saveFile.FileName, comboSoftwareBox.SelectedItem.ToString(), txtAObox.Text, txtSiteNameBox.Text, comboTemperatureBox.SelectedItem.ToString(), comboflowrateUnitsBox.SelectedItem.ToString(), Convert.ToInt32(NoCassPerTrain.Value), Convert.ToInt32(NoModulesPerCass.Value), Convert.ToSingle(NoAreaPerModule.Value), radioSquareFeet.Checked, comboSiteAssignedBox.SelectedItem.ToString());
                    

                }
                catch (Exception ex)
                {
                    MessageBox.Show("An Error Occured\r\n\r\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void comboflowrateUnitsBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (comboflowrateUnitsBox.SelectedItem.ToString() == "No FlowRates")
            {
                //disable area controls if no flowrate
                    NoCassPerTrain.Enabled = false;
                    NoModulesPerCass.Enabled = false;
                    NoAreaPerModule.Enabled = false;
                    radioSquareFeet.Enabled = false;
                    radioSquareMetres.Enabled = false;
                          
            }
            else
            {
                NoCassPerTrain.Enabled = true;
                NoModulesPerCass.Enabled = true;
                NoAreaPerModule.Enabled = true;
                radioSquareFeet.Enabled = true;
                radioSquareMetres.Enabled = true;

            }

        }

        private void comboDriverBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (comboDriverBox.SelectedItem.ToString() == "IGS")
            {
               buttonIGS.Enabled = true;
               buttonIGS.Visible = true;

               //enable

               txtChannelBox.Enabled = true;
               txtPLCBox.Enabled = true;
               txtDatabasebox.Enabled = true;
               buttonIfixDatabase.Enabled = true;

               //hide controls not used
               txtChannelBox.Visible = true;
               txtDatabasebox.Visible = true;
               txtPLCBox.Visible = true;
               buttonIfixDatabase.Visible = true;

               //labels
               labelChannelName.Visible = true;
               labelDatabaseName.Visible = true;
               labelPLCName.Visible = true;

                //ifix script button
               buttonIfixScript.Enabled = false;
               buttonIfixScript.Visible = false;

              
        
                              
                
            }
            else if (comboDriverBox.SelectedItem.ToString( )== "GE9")
            {
                buttonIGS.Enabled = false;
                buttonIGS.Visible = false;
                txtDatabasebox.Enabled = true;
                buttonIfixDatabase.Enabled = true;
                //unhide controls not used
                txtDatabasebox.Visible = true;
                buttonIfixDatabase.Visible = true;
              
                //labels
                labelDatabaseName.Visible = true;

                //IFIX
                txtChannelBox.Enabled = false;
                txtPLCBox.Enabled = true;
             
                //hide controls not used
                txtChannelBox.Visible = false;
                txtPLCBox.Visible = true;
             
                //labels
                labelChannelName.Visible = false;
                labelPLCName.Visible = true;

                //ifix script button
                buttonIfixScript.Enabled = false;
                buttonIfixScript.Visible = false;
               
            }
            else
            {
             
                buttonIGS.Enabled = false;
                
            }
        }

        private void comboSoftware_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (comboSoftware.SelectedItem.ToString() == "IFIX")
            {
                comboDriverBox.Enabled = true;
                comboDriverBox.Visible = true;
                //disable igs button
                buttonIGS.Enabled = false;
                buttonIGS.Visible = false;

                //disable opc
                buttonOPCFile.Visible = false;
                buttonOPCFile.Enabled = false;
                txtChannelBoxOPC.Enabled = false;
                txtChannelBoxOPC.Visible = false;
                txtPLCBoxOPC.Enabled = false;
                txtPLCBoxOPC.Visible = false;
                //labels
                labelCHNameOPC.Visible = false;
                labelPLCNameOPC.Visible = false;
                buttonOPCFile.Visible = false;
            }
            else if (comboSoftware.SelectedItem.ToString() == "OPC Trend")
            {
                buttonIGS.Enabled = true;
                buttonIGS.Visible = true;
                comboDriverBox.Enabled = false;
                comboDriverBox.Visible = false;
                
                //opc
                buttonOPCFile.Visible = true;
                buttonOPCFile.Enabled = true;
                txtChannelBoxOPC.Enabled = true;
                txtChannelBoxOPC.Visible = true;
                txtPLCBoxOPC.Enabled = true;
                txtPLCBoxOPC.Visible = true;
                //labels
                labelCHNameOPC.Visible = true;
                labelPLCNameOPC.Visible = true;

                //IFIX
                txtChannelBox.Enabled = false;
                txtPLCBox.Enabled = false;
                txtDatabasebox.Enabled = false;
                buttonIfixDatabase.Enabled = false;

                //hide controls not used
                txtChannelBox.Visible = false;
                txtDatabasebox.Visible = false;
                txtPLCBox.Visible = false;
                buttonIfixDatabase.Visible = false;

                //labels
                labelChannelName.Visible = false;
                labelDatabaseName.Visible = false;
                labelPLCName.Visible = false;
            }
        }

        private void buttonOPCFile_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtSiteBox.Text))
            {
                MessageBox.Show("Please enter a site name.");
                txtSiteBox.Focus();
                return;

            }

            if (string.IsNullOrEmpty(txtChannelBoxOPC.Text))
            {
                MessageBox.Show("Please enter the IGS channel name.");
                txtChannelBoxOPC.Focus();
                return;

            }
            if (string.IsNullOrEmpty(txtPLCBoxOPC.Text))
            {
                MessageBox.Show("Please enter the IGS PLC name.");
                txtPLCBoxOPC.Focus();
                return;

            }


            saveFile.FileName = txtAONumberBox.Text + txtSiteBox.Text + "OPCFile";
            saveFile.Filter = "CSV file (*.csv)|*.csv";
            System.Windows.Forms.DialogResult result;


            result = saveFile.ShowDialog();


            if (result == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    aFile.SaveOPCDataTableToCsvFile(saveFile.FileName, txtChannelBoxOPC.Text, txtPLCBoxOPC.Text);
                       
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An Error Occured\r\n\r\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Application.Exit();
        }

        private void comboflowrateUnitsBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}