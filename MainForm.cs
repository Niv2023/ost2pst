using ost2pst;
using System;
using System.Windows.Forms;

namespace ost2pst
{
    public partial class ost2pst : Form
    {
        private const string version = "1.0 (Jul-2025)"; // Version of the application
        private string folderToExport = string.Empty;
        private bool treeViewEnabled = false; // Flag to check if the tree view is enabled  
        public ost2pst()
        {
            InitializeComponent();
            Text = $"OST to PTS v={version}"; // Set the form title with the version
            openOST.Enabled = true; // Enable the button to open OST/PST files
            exportPST.Enabled = false; // Disable the export button until a folder is selected
        }
        public void statusMSG(string msg, bool newLine = true)
        {
            if (newLine)
            {
                statusList.Items.Add(msg);
            }
            else
            {
                statusList.Items[statusList.Items.Count - 1] = msg; // Update the last item without adding a new line
            }
            statusList.SelectedIndex = statusList.Items.Count - 1; // Select the last item
            statusList.Update();
        }

        private void PopulateTreeView(List<Folder> folders)
        {
            folderView.Nodes.Clear(); // Clear existing nodes before populating
            if (folders == null || folders.Count == 0)
                return;

            // Build a dictionary for quick lookup
            var folderNodes = new Dictionary<Folder, TreeNode>();
            folders[0].parent = null; // Ensure the root folder has no parent
            foreach (var folder in folders)
            {
                var node = new TreeNode(folder.name);
                node.Tag = folder;
                folderNodes[folder] = node;

                if (folder.parent == null)
                {
                    folderView.Nodes.Add(node);
                }
                else if (folderNodes.TryGetValue(folder.parent, out TreeNode parentNode))
                {
                    parentNode.Nodes.Add(node);
                }
            }

            folderView.ExpandAll(); // Optional: expands all nodes
            if (FM.srcFile.type == FileType.OST)
            {   // folder selection is only available for OST files
                treeViewEnabled = true; // Set the flag to indicate the tree view is populated and enabled
                folderView.AfterSelect += treeViewFolders_AfterSelect;
            }
            else
            {
                folderView.AfterSelect += null;
                treeViewEnabled = false; // Set the flag to indicate the tree view is populated and enabled
                exportPST.Text = $"Export ALL to PST"; // Set the export button text for PST files
            }
        }
        private void treeViewFolders_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (!treeViewEnabled) return; // Check if the tree view is enabled
            var selectedFolder = e.Node.Tag as Folder;
            if (selectedFolder != null)
            {
                folderToExport = selectedFolder.name;
                exportPST.Text = $"Export \" {folderToExport} \" to PST";
                exportPST.Enabled = true; // Enable the export button when a folder is selected
            }
            else
            {
                folderToExport = string.Empty; // Reset if no folder is selected
                exportPST.Text = $"Select folder to export";
                exportPST.Enabled = false; // Enable the export button when a folder is selected
            }
        }
        private void openOST_Click(object sender, EventArgs e)
        {
            openOST.Enabled = false; // disable the button to open OST/PST files
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "OST files (*.ost)|*.ost|PST files (*.pst)|*.pst|All files (*.*)|*.*";
                    openFileDialog.Title = "Open OST/PST File";
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string srcFile = openFileDialog.FileName;
                        statusMSG($"reading file data");
                        if (FM.OpenSourceFile(srcFile))
                        {
                            statusMSG($"getting folder list");
                            FM.GetFolderList();
                            PopulateTreeView(FM.folders);
                            exportPST.Enabled = true; // Enable the export button after loading folders
                            fileDetails.Text = FM.SourceFileDetails();
                            statusMSG($"folder list completed");
                        }
                    }
                }
            }
            finally
            {
                xbRemovePassword.Checked = false; // Reset the remove password checkbox
                if (FM.srcFile.type == FileType.OST)
                {
                    pstGroup.Enabled = false; // Disable the export group for OST files
                }
                else
                {
                    pstGroup.Enabled = true; // Enable the export group for PST files
                }
                openOST.Enabled = true; // Re-enable the button after file selection`
            }
        }

        private void exportPST_Click(object sender, EventArgs e)
        {
            openOST.Enabled = false; // Enable the button to open OST/PST files
            exportPST.Enabled = false; // Disable the export button while exporting
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "PST files (*.pst)|*.pst|All files (*.*)|*.*";
                    saveFileDialog.Title = "Save PST File";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string destFile = saveFileDialog.FileName;
                        string tmpPstFile = Path.GetTempPath() + "tempPST.pst";
                        if (FM.srcFile.type == FileType.OST)
                        {
                            // for OST files, we need 2 runs for the conversion.
                            // this is because the message sizes may change due to different block/page sizes of the ost file
                            // please note the message sizes seems not to be validated by outlook,
                            // however SCANPST will flag it an repair it
                            // this 2nd run conversion avoids errors in the SCANPST
                            statusMSG($"Converting OST FILE");
                            statusMSG($"Building temp pst file");
                            statusMSG($"to recalculate the message sizes");
                            statusMSG($"PLEASE WAIT...this may take few minutes");
                            if (FM.CreatPstFile(tmpPstFile))
                            {
                                FM.CopySourceDatablocksToPST(folderToExport, Path.GetFileName(destFile));
                                FM.exportNBTnodes();
                                FM.exportBBTnodes();
                                FM.updateNidHighWaterMarks();
                                FM.CloseOutputFile();
                                FM.CloseSourceFile();
                                FM.OpenSourceFile(tmpPstFile);
                            }
                        }
                        statusMSG($"building final pst file");
                        if (FM.CreatPstFile(destFile))
                        {
                            FM.rebuildPSTfile(xbRemovePassword.Checked);   // this will rebuild the pst file with the correct message sizes
                            FM.exportNBTnodes();
                            FM.exportBBTnodes();
                            FM.updateNidHighWaterMarks();
                            FM.CloseOutputFile();
                            statusMSG($"export completed");
                        }
                        pstDetails.Text = FM.OutputFileDetails();
                    }
                }
            }
            finally
            {
                openOST.Enabled = true; // Enable the button to open OST/PST files
                treeViewEnabled = false; // treeview will be enabled after loading a new pst/ost file
            }
        }
    }
}
