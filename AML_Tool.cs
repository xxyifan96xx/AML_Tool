// Author: Yifan Wang

using System.ComponentModel;
using System.Diagnostics;
using Microsoft.WindowsAPICodePack.Dialogs;

using AMLConversion_ClassLibrary;

namespace AML_Tool
{
    public partial class AML_Tool : Form
    {
        ///////////////////////////////////////////////////////////////////////////////////
        //
        // Constructor
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region Constructor

        private delegate void debugDelegate(string text, bool msgBox);
        private delegate void labelDelegate(string text);
        private delegate void progressBarDelegate(int progress);
        private delegate void buttonDelegate(string text);
        private delegate void buttonEnableDelegate(bool enable);

        private bool debug;
        private string developmentDirectory;
        private CommonOpenFileDialog openFolderDialog;

        private FeeAPI Api;
        private Importer_fee Imp;
        private Exporter_physicalPlant Exp;
        private SerialCommunication Ser;
        private Simulator Sim;

        public AML_Tool()  // Constructor
        {
            debug = false;
            openFolderDialog = new CommonOpenFileDialog();
            Api = new FeeAPI();
            Ser = new SerialCommunication();
            Imp = new Importer_fee(Api);
            Exp = new Exporter_physicalPlant(Ser);
            Sim = new Simulator(Api, Ser);
            
            Ser.PropertyChanged += SerOnPropertyChanged;
            Api.PropertyChanged += ApiOnPropertyChanged;
            Imp.PropertyChanged += ImpOnPropertyChanged;
            Exp.PropertyChanged += ExpOnPropertyChanged;
            Sim.PropertyChanged += SimOnPropertyChanged;
            InitializeComponent();
            try
            {
                developmentDirectory = Directory.GetParent(Process.GetCurrentProcess().MainModule.FileName).Parent.Parent.FullName;
                if (debug == true)
                    developmentDirectory = Directory.GetParent(developmentDirectory).Parent.Parent.FullName;
            }
            catch (Exception ex)
            {
                debug_write(ex.ToString());
            }
            debug_write(developmentDirectory);
        }
#endregion

        ///////////////////////////////////////////////////////////////////////////////////
        //
        // PropertyChanged event listener
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region PropertyChanged event listener

        private void ApiOnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if(e.PropertyName == "ConnectionState")
            {
                debug_write("[Fee API Connection] " + Api.ConnectionState);
                label_APIconnectionState_refresh(Api.ConnectionState);
            }
            else if(e.PropertyName == "IsApiConnected")
            {
                if (Api.IsApiConnected)
                {
                    label_APIconnectionStateLamp.ForeColor = Color.Green;
                    label_APIconnectionStateLamp2.ForeColor = Color.Green;
                    button_APIconnect_refresh("Disconnect");
                }
                else
                {
                    label_APIconnectionStateLamp.ForeColor = Color.Red;
                    label_APIconnectionStateLamp2.ForeColor = Color.Red;
                    button_APIconnect_refresh("Connect");
                }
            }
            else if (e.PropertyName == "SimulationState")
            {
                debug_write("[Fee API Simulation] " + Api.ConnectionState);
            }
        }

        private void SerOnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "StatusText")
            {
                label_serialConnectionState_refresh(Ser.Status.Text);
                debug_write("[Serial] " + Ser.Status.Text);
            }
            else if (e.PropertyName == "ARConnectionState")
            {
                label_ARSerialConnectionState_refresh(Ser.ARConnectionState);
            }
            else if (e.PropertyName == "SSConnectionState")
            {
                label_SSSerialConnectionState_refresh(Ser.SSConnectionState);
            }
            else if (e.PropertyName == "IsSSConnected" || e.PropertyName == "IsARConnected")
            {
                if (Ser.IsARConnected)
                {
                    label_ARSerialConnectionStateLamp.ForeColor = Color.Green;
                    label_ARSerialConnectionStateLamp2.ForeColor = Color.Green;
                }
                else
                {
                    label_ARSerialConnectionStateLamp.ForeColor = Color.Red;
                    label_ARSerialConnectionStateLamp2.ForeColor = Color.Red;
                }
                if(Ser.IsSSConnected)
                {
                    label_SSSerialConnectionStateLamp.ForeColor = Color.Green;
                    label_SSSerialConnectionStateLamp2.ForeColor = Color.Green;
                }
                else
                {
                    label_SSSerialConnectionStateLamp.ForeColor = Color.Red;
                    label_SSSerialConnectionStateLamp2.ForeColor = Color.Red;
                }

                if (Ser.IsARConnected && Ser.IsSSConnected)
                {
                    button_serialConnect_refresh("Disconnect");
                    groupBox_scan.Enabled = true;
                }
                else
                {
                    button_serialConnect_refresh("Connect");
                    groupBox_scan.Enabled = false;
                }
            }
        }

        private void ImpOnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "StatusText")
            {
                label_statusText_refresh(Imp.Status.Text);
                debug_write("[Importer] " + Imp.Status.Text);
            }
            else if (e.PropertyName == "StatusProgress")
            {
                label_statusProgress_refresh("[" + Imp.Status.Progress.ToString() + "%]");
            }
            else if (e.PropertyName == "StatusIsWorking");
            {
                if (Imp.Status.IsWorking == true)
                    button_import_enable(false);
                else
                    button_import_enable(true);
            }
        }

        private void ExpOnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "StatusText")
            {
                label_statusText_refresh(Exp.Status.Text);
                debug_write("[Exporter] " + Exp.Status.Text);
            }
            else if (e.PropertyName == "StatusProgress")
            {
                label_statusProgress_refresh("[" + Exp.Status.Progress.ToString() + "%]");
            }
            else if (e.PropertyName == "StatusIsWorking") ;
            {
                if (Exp.Status.IsWorking == true)
                {
                    button_export_enable(false);
                }
                else
                {
                    button_export_enable(true);
                }
            }
        }

        private void SimOnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "StatusText")
            {
                label_statusText_refresh(Sim.Status.Text);
                debug_write("[Simulator] " + Sim.Status.Text);
            }
            else if (e.PropertyName == "StatusProgress")
            {
                label_statusProgress_refresh("[" + Sim.Status.Progress.ToString() + "%]");
            }
            else if (e.PropertyName == "StatusIsWorking")
            {
                if (Sim.Status.IsWorking == true)
                {
                    button_simulate_enable(false);
                }
                else
                {
                    button_simulate_enable(true);
                }
            }
            else if (e.PropertyName == "IsScanned")
            {
                if (Sim.IsScanned)
                {
                    label_scanStatus_refresh("Scanned");
                    label_scanStatusLamp.ForeColor = Color.Green;
                }
                else
                {
                    label_scanStatus_refresh("Not scanned");
                    label_scanStatusLamp.ForeColor = Color.Red;
                }
            }
        }

#endregion

        ///////////////////////////////////////////////////////////////////////////////////
        //
        // General forms functions
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region General forms functions

        private void AMLTool_Load(object sender, EventArgs e)
        {
        }

        private void AMLTool_FormClosing(object sender, FormClosingEventArgs e)
        {
            Api.Disconnect(deleteEventHandler: true);
            Ser.Disconnect();
        }

        private void button_import_enable(bool enable)
        {
            if (button_import.InvokeRequired)
            {
                var d = new buttonEnableDelegate(button_import_enable);
                button_import?.Invoke(d, new object[] { enable });
            }
            else
            {
                button_import.Enabled = enable;
            }
        }

        private void button_simulate_enable(bool enable)
        {
            if (button_simulate.InvokeRequired || button_scan.InvokeRequired)
            {
                var d = new buttonEnableDelegate(button_simulate_enable);
                button_simulate?.Invoke(d, new object[] { enable });
            }
            else
            {
                button_simulate.Enabled = enable;
                button_scan.Enabled = enable;
            }
        }

        private void button_export_enable(bool enable)
        {
            if (button_plantmodelExport.InvokeRequired || button_exportTopology.InvokeRequired)
            {
                var d = new buttonEnableDelegate(button_export_enable);
                button_plantmodelExport?.Invoke(d, new object[] { enable });
            }
            else
            {
                button_plantmodelExport.Enabled = enable;
                button_exportTopology.Enabled = enable;
            }
        }

        private void button_APIconnect_refresh(string text)
        {
            if (button_APIconnect.InvokeRequired)
            {
                var d = new buttonDelegate(button_APIconnect_refresh);
                button_APIconnect?.Invoke(d, new object[] { text });
            }
            else
            {
                button_APIconnect.Text = text;
                button_APIconnect2.Text = text;
            }
        }

        private void button_serialConnect_refresh(string text)
        {
            if (button_serialConnect.InvokeRequired)
            {
                var d = new buttonDelegate(button_serialConnect_refresh);
                button_serialConnect?.Invoke(d, new object[] { text });
            }
            else
            {
                button_serialConnect.Text = text;
                button_serialConnect2.Text = text;
            }
        }

        private void label_APIconnectionState_refresh(string text)
        {
            if (label_APIconnectionState.InvokeRequired)
            {
                var d = new labelDelegate(label_APIconnectionState_refresh);
                label_APIconnectionState?.Invoke(d, new object[] { text });
            }
            else
            {
                label_APIconnectionState.Text = text;
                label_APIconnectionState2.Text = text;
            }
        }

        private void label_ARSerialConnectionState_refresh(string text)
        {
            if (label_ARSerialConnectionState.InvokeRequired)
            {
                var d = new labelDelegate(label_ARSerialConnectionState_refresh);
                label_ARSerialConnectionState?.Invoke(d, new object[] { text });
            }
            else
            {
                label_ARSerialConnectionState.Text = text;
                label_ARSerialConnectionState2.Text = text;
            }
        }

        private void label_SSSerialConnectionState_refresh(string text)
        {
            if (label_SSSerialConnectionState.InvokeRequired)
            {
                var d = new labelDelegate(label_SSSerialConnectionState_refresh);
                label_SSSerialConnectionState?.Invoke(d, new object[] { text });
            }
            else
            {
                label_SSSerialConnectionState.Text = text;
                label_SSSerialConnectionState2.Text = text;
            }
        }

        private void label_serialConnectionState_refresh(string text)
        {
            if (label_SSSerialConnectionState.InvokeRequired)
            {
                var d = new labelDelegate(label_serialConnectionState_refresh);
                label_SSSerialConnectionState?.Invoke(d, new object[] { text });
            }
            else
            {
                label_serialConnectionState.Text = text;
                label_serialConnectionState2.Text = text;
            }
        }

        private void label_scanStatus_refresh(string text)
        {
            if (label_scanStatus.InvokeRequired)
            {
                var d = new labelDelegate(label_scanStatus_refresh);
                label_scanStatus?.Invoke(d, new object[] { text });
            }
            else
            {
                label_scanStatus.Text = text;
            }
        }

        private void label_statusText_refresh(string text)
        {
            if (label_statusText.InvokeRequired)
            {
                var d = new labelDelegate(label_statusText_refresh);
                label_statusText?.Invoke(d, new object[] { text });
            }
            else
            {
                label_statusText.Text = text;
                label_statusText.Refresh();
            }
        }

        private void label_statusProgress_refresh(string text)
        {
            if (label_statusText.InvokeRequired)
            {
                var d = new labelDelegate(label_statusProgress_refresh);
                label_statusProgress?.Invoke(d, new object[] { text });
            }
            else
            {
                label_statusProgress.Text = text;
                label_statusProgress.Refresh();
            }
        }

        private void debug_write(string text, bool msgBox = false)
        {
            if (listbox_debugging.InvokeRequired)
            {
                var d = new debugDelegate(debug_write);
                listbox_debugging?.Invoke(d, new object[] { text, msgBox });
            }
            else
            {
                listbox_debugging.Items.Add(text);
                listbox_debugging.SelectedIndex = listbox_debugging.Items.Count - 1;  // autoscroll the listBox

                if(msgBox)
                {
                    MessageBox.Show(text);
                }
            }
        }

#endregion

        ///////////////////////////////////////////////////////////////////////////////////
        //
        // Topology Export tab functions
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region Topology Export tab form functions

        private void button_blueprintLibraries_Click(object sender, EventArgs e)
        {
            // Open open-file dialog
            var directoryPath = developmentDirectory + @"\AML\Vorlagen";
            
            openFileDialog_blueprint.InitialDirectory = directoryPath;
            openFileDialog_blueprint.FileName = "Demonstrator.aml";
            openFileDialog_blueprint.Filter = ".aml | *.aml";
            openFileDialog_blueprint.ShowDialog();
        }

        private void openFileDialog_blueprint_FileOk(object sender, CancelEventArgs e)
        {
            button_blueprintLibraries.Text = openFileDialog_blueprint.FileName;
        }

        private void radioButton_physicalDemonstrator_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_physicalDemonstrator.Checked)
            {
                groupBox_serialConnection.Enabled = true;
                groupBox_scan.Enabled = true;
            }
            else
            {
                groupBox_serialConnection.Enabled = false;
                groupBox_scan.Enabled = false;
            }
        }

        private void button_serialConnect_Click(object sender, EventArgs e)
        {
            if(button_serialConnect.Text == "Connect")
            {
                Ser.Connect();
            }
            else if (button_serialConnect.Text == "Disconnect")
            {
                Ser.Disconnect();
            }
        }

        private async void button_scan_Click(object sender, EventArgs e)
        {
            Exp.SlotProperties = await Sim.Scan_slotPropertiesAsync();
        }

        private void radioButton_importCSV_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_importCSV.Checked)
            {
                button_slotProperties.Enabled = true;
            }
            else
            {
                button_slotProperties.Enabled = false;
            }
        }

        private void button_slotProperties_Click(object sender, EventArgs e)
        {
            // Open open-file dialog
            var directoryPath = developmentDirectory + @"\CSV";

            openFileDialog_slotProperties.InitialDirectory = directoryPath;
            openFileDialog_slotProperties.FileName = "slotProperties.csv";
            openFileDialog_slotProperties.Filter = ".csv | *.csv";
            openFileDialog_slotProperties.ShowDialog();
        }

        private void openFileDialog_slotProperties_FileOk(object sender, CancelEventArgs e)
        {
            button_slotProperties.Text = openFileDialog_slotProperties.FileName;
        }

        private async void button_exportTopology_Click(object sender, EventArgs e)
        {
            //Check for given blueprint library file
            if (openFileDialog_blueprint.FileName == null || openFileDialog_blueprint.FileName == "")
            {
                debug_write("AML blueprint libraries are missing!");
            }
            else if (radioButton_importCSV.Checked && (openFileDialog_slotProperties.FileName == null || openFileDialog_slotProperties.FileName == ""))
            {
                debug_write("CSV slotProperties are missing!");
            }
            else if (radioButton_physicalDemonstrator.Checked && !Sim.IsScanned)
            {
                debug_write("Scan of slot properties is required!");
            }
            else
            {
                var directoryPath = developmentDirectory + @"\AML";

                // Execute export asnychronously
                if (radioButton_importCSV.Checked)
                {
                    // reset scanned data
                    Sim.IsScanned = false;

                    await Exp.Export_topologyAsync(openFileDialog_blueprint.FileName, CSVPath: openFileDialog_slotProperties.FileName);
                }
                else
                {
                    // reset chosen CSV file
                    openFileDialog_slotProperties.FileName = "";
                    button_slotProperties.Text = "Choose file";

                    await Exp.Export_topologyAsync(openFileDialog_blueprint.FileName);
                }

                // Open save-to-file dialog
                saveFileDialog_topology.InitialDirectory = directoryPath;
                saveFileDialog_topology.FileName = "Topology.aml";
                saveFileDialog_topology.Filter = ".aml | *.aml";
                saveFileDialog_topology.ShowDialog();
            }
        }

        private void saveFileDialog_topology_FileOk(object sender, CancelEventArgs e)
        {
            Exp.Save_file(saveFileDialog_topology.FileName);

            // Intelligent FileName entry on other tabs
            openFileDialog_topology.FileName = saveFileDialog_topology.FileName;
            button_topologyFile2.Text = saveFileDialog_topology.FileName;
            button_topologyFile1.Text = saveFileDialog_topology.FileName;
        }

#endregion

        ///////////////////////////////////////////////////////////////////////////////////
        //
        // Plantmodel Export tab functions
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region Plantmodel Export form functions
        private void button_topologyFile1_Click(object sender, EventArgs e)
        {
            // Open open-file dialog
            var directoryPath = developmentDirectory + @"\AML";

            openFileDialog_topology.InitialDirectory = directoryPath;
            openFileDialog_topology.FileName = "Topology.aml";
            openFileDialog_topology.Filter = ".aml | *.aml";
            openFileDialog_topology.ShowDialog();
        }

        private void openFileDialog_topology_FileOk(object sender, CancelEventArgs e)
        {
            button_topologyFile2.Text = openFileDialog_topology.FileName;
            button_topologyFile1.Text = openFileDialog_topology.FileName;
        }

        private void radioButton_custom_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_custom.Checked)
            {
                groupBox_manualSequence.Enabled = true;
            }
            else
            {
                groupBox_manualSequence.Enabled = false;
            }
        }

        private void button_red_Click(object sender, EventArgs e)
        {
            var counter = listView_sequence.Items.Count + 1;
            ListViewItem item = new ListViewItem(Convert.ToString(counter));
            item.SubItems.Add("red");
            listView_sequence.Items.Add(item);
        }

        private void button_green_Click(object sender, EventArgs e)
        {
            var counter = listView_sequence.Items.Count + 1;
            ListViewItem item = new ListViewItem(Convert.ToString(counter));
            item.SubItems.Add("green");
            listView_sequence.Items.Add(item);
        }

        private void button_blue_Click(object sender, EventArgs e)
        {
            var counter = listView_sequence.Items.Count + 1;
            ListViewItem item = new ListViewItem(Convert.ToString(counter));
            item.SubItems.Add("blue");
            listView_sequence.Items.Add(item);
        }

        private void button_remove_Click(object sender, EventArgs e)
        {
            if (listView_sequence.Items.Count > 0)
            {
                listView_sequence.Items.RemoveAt(listView_sequence.Items.Count - 1);
            }
        }

        private void button_clear_Click(object sender, EventArgs e)
        {
            listView_sequence.Items.Clear();
        }

        private async void button_plantmodelExport_Click(object sender, EventArgs e)
        {
            // Check for given topology file
            if (openFileDialog_topology.FileName == null || openFileDialog_topology.FileName == "")
            {
                debug_write("AML topology file is missing!");
            }
            else
            {
                if (radioButton_automatic.Checked)
                {
                    await Exp.Export_plantfileAsync(openFileDialog_topology.FileName, "automatic");
                }
                if (radioButton_custom.Checked)
                {
                    var listViewItems = new List<string>();
                    foreach (ListViewItem item in listView_sequence.Items)
                    {
                        listViewItems.Add(item.SubItems[1].Text);
                    }

                    await Exp.Export_plantfileAsync(openFileDialog_topology.FileName, "custom", listViewItems);
                }

                // Open save-to-file dialog
                var directoryPath = developmentDirectory + @"\AML";

                saveFileDialog_plantmodel.InitialDirectory = directoryPath;
                saveFileDialog_plantmodel.FileName = "Plantmodel.aml";
                saveFileDialog_plantmodel.Filter = ".aml | *.aml";
                saveFileDialog_plantmodel.ShowDialog();
            }
        }

        private void saveFileDialog_plantmodel_FileOk(object sender, CancelEventArgs e)
        {
            Exp.Save_file(saveFileDialog_plantmodel.FileName);

            // Intelligent FileName entry on other tabs
            openFileDialog_plantmodel.FileName = saveFileDialog_plantmodel.FileName;
            button_plantFile.Text = saveFileDialog_plantmodel.FileName;
        }

#endregion

        ///////////////////////////////////////////////////////////////////////////////////
        //
        // VC Topology Import tab functions
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region VC Topology Import tab form functions

        private void button_topologyFile2_Click(object sender, EventArgs e)
        {
            button_topologyFile1_Click(sender, e);
        }

        private void button_feeLibraryFolder_Click(object sender, EventArgs e)
        {
            // Open folder browser dialog
            var directoryPath = developmentDirectory + @"\FeeAPI\API Directory";

            openFolderDialog.InitialDirectory = directoryPath;
            openFolderDialog.IsFolderPicker = true;
            if (openFolderDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                button_feeLibraryFolder.Text = openFolderDialog.FileName;
            }

        }

        public void button_APIconnect_Click(object sender, EventArgs e)
        {
            if (button_APIconnect.Text == "Connect")
            {
                Api.Connect();
            }
            else if (button_APIconnect.Text == "Disconnect")
            {
                Api.Disconnect();
            }
        }

        private async void button_import_Click(object sender, EventArgs e)
        {
            if (Api.IsApiConnected == false)
            {
                debug_write("Connection to Fee API is missing!");
            }
            // Check for given plant model file
            else if (openFileDialog_topology.FileName == null || openFileDialog_topology.FileName == "")
            {
                debug_write("AML blueprint libraries are missing!");
            }
            else
            {
                await Imp.ImportAsync(openFileDialog_topology.FileName, openFolderDialog.FileName);
            }
        }

#endregion

        ///////////////////////////////////////////////////////////////////////////////////
        //
        // VC Simulate tab functions
        //
        ///////////////////////////////////////////////////////////////////////////////////
#region VC Simulate tab form functions

        private void button_plantFile_Click(object sender, EventArgs e)
        {
            // Open open-file dialog
            var directoryPath = developmentDirectory + @"\AML";

            openFileDialog_plantmodel.InitialDirectory = directoryPath;
            openFileDialog_plantmodel.FileName = "Plantmodel.aml";
            openFileDialog_plantmodel.Filter = ".aml | *.aml";
            openFileDialog_plantmodel.ShowDialog();
        }

        private void openFileDialog_plantmodel_FileOk(object sender, CancelEventArgs e)
        {
            button_plantFile.Text = openFileDialog_plantmodel.FileName;
        }

        private void checkBox_simulateVC_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox_simulateVC.Checked == true)
            {
                groupBox_APIConnection2.Enabled = true;
            }
            else
            {
                groupBox_APIConnection2.Enabled = false;
            }
        }

        private void button_APIconnect2_Click(object sender, EventArgs e)
        {
            button_APIconnect_Click(sender, e);
        }

        private void checkBox_simulatePD_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_simulatePD.Checked == true)
            {
                groupBox_serialConnection2.Enabled = true;
            }
            else
            {
                groupBox_serialConnection2.Enabled = false;
            }
        }

        private void button_serialConnect2_Click(object sender, EventArgs e)
        {
            button_serialConnect_Click(sender, e);
        }

        private async void button_simulate_Click(object sender, EventArgs e)
        {
            if (openFileDialog_plantmodel.FileName == null || openFileDialog_plantmodel.FileName == "")
            {
                debug_write("AML plantfile is missing!");
            }
            else if (checkBox_simulateVC.Checked && Api.IsApiConnected == false)
            {
                debug_write("Connection to Fee API is missing!");
            }
            else if (checkBox_simulatePD.Checked && (Ser.IsARConnected == false || Ser.IsSSConnected == false))
            {
                debug_write("Serial connection is missing!");
            }
            else
            {
                await Sim.SimulateAsync(openFileDialog_plantmodel.FileName, checkBox_simulatePD.Checked, checkBox_simulateVC.Checked);
            }
        }

#endregion
    }
}