using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Security.Cryptography;
using Microsoft.Win32;

namespace SFP
{
    public partial class SFP_main : Form
    {
        string logging = "off"; //by default

        public SFP_main()
        {
            InitializeComponent();
        }

        public void startlogging()
        {
            // Create the trace listener.
            if (File.Exists("sfp-log.txt"))
            {
                Stream sfplog = File.Open("sfp-log.txt", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                Trace.Listeners.Add(new TextWriterTraceListener(sfplog));
            }
            else
            {
                Stream sfplog = File.Create("sfp-log.txt");
                TextWriterTraceListener listener = new TextWriterTraceListener(sfplog);
                Trace.Listeners.Add(listener);
            }
        }

        private void savecsv_format(DataGridView x)
        {
            SaveFileDialog s = new SaveFileDialog();
            s.Filter = "CSV Files (*.csv)|*.csv";
            if (s.ShowDialog() == DialogResult.OK)
            {
                string buffer_csv_return = "\r\n";
                // Create the output CSV file
                TextWriter T = new StreamWriter(s.FileName);

                DataTable csvoutput = x.DataSource as DataTable;
                // print column names first
                foreach (DataColumn column in csvoutput.Columns)
                {
                    string buffer_csv = string.Concat(column.ColumnName, "|"); T.Write(buffer_csv);
                }
                T.Write(buffer_csv_return);

                // now print the actual data
                foreach (DataRow row in csvoutput.Rows)
                {
                    foreach (DataColumn column in csvoutput.Columns)
                    {
                        string buffer_csv = string.Concat(row[column], "|"); T.Write(buffer_csv);
                    }
                    T.Write(buffer_csv_return);
                }
                T.Close();
                MessageBox.Show(string.Concat(Path.GetFileName(s.FileName), " saved sucessfully"));
            }

        }

        public string GetMD5HashFromStream(Stream file)
        {
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] retVal = md5.ComputeHash(file);
            return bytestostring(retVal, null);
        }

        public string GetMD5HashFromFile(String path)
        {
            MD5 md5 = new MD5CryptoServiceProvider();
            FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            byte[] retVal = md5.ComputeHash(fileStream);
            return bytestostring(retVal, null);
        }

        public string bytestostring(byte[] toconvert, string type)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < toconvert.Length; i++)
            {
                sb.Append(toconvert[i].ToString("X2"));
                if (type == "mac" && i != (toconvert.Length - 1))
                {
                    sb.Append(":");
                }
            }
            return sb.ToString();
        }

        public string bytestostring_littleendian(byte[] toconvert)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = toconvert.Length - 1; i >= 0; i--)
            {
                sb.Append(toconvert[i].ToString("X2"));
            }
            return sb.ToString();
        }

        private void parselnk_setuptable(DataTable x, Panel panel, string type)
        {
            if (type == "lnk only")
            {
                tabControl_main.SelectedTab = tabControl_main.TabPages[0];
                // Make the .lnk panel visible.
                panel.Visible = true;
                dataGridView_lnk.Visible = false;
                statuslabel_lnk_filesparsed.Text = "";
                statuslabel_lnk_timetaken.Text = "";
            }
            else if (type == "embedded lnk")
            {
                // embedded lnk within a jump list
            }

            // Create the columns.
            x.Columns.Add("LNK File Name", typeof(string));
            x.Columns.Add("Linked Path", typeof(string));
            x.Columns.Add("LNK File Creation Time (Local)", typeof(DateTime));
            x.Columns.Add("LNK File Access Time (Local)", typeof(DateTime));
            x.Columns.Add("LNK File Written Time (Local)", typeof(DateTime));
            x.Columns.Add("Embedded Creation Time (Local)", typeof(DateTime));
            x.Columns.Add("Embedded Access Time (Local)", typeof(DateTime));
            x.Columns.Add("Embedded Written Time (Local)", typeof(DateTime));
            x.Columns.Add("File Size (Bytes)", typeof(uint));

            // Columns for the LinkInfo section.
            x.Columns.Add("Server Share Path", typeof(string));
            x.Columns.Add("NetBIOS Name", typeof(string));
            x.Columns.Add("Target Drive Type", typeof(string));
            x.Columns.Add("Volume Serial Number", typeof(string));
            x.Columns.Add("Volume Label", typeof(string));

            // Columns for the StringData section.
            x.Columns.Add("Description", typeof(string));
            x.Columns.Add("Relative Path", typeof(string));
            x.Columns.Add("Working Directory", typeof(string));
            x.Columns.Add("Command Line Args", typeof(string));
            x.Columns.Add("Current VolumeID", typeof(string));
            x.Columns.Add("Current ObjectID", typeof(string));
            x.Columns.Add("Current Mac Addr", typeof(string));
            x.Columns.Add("Birth VolumeID", typeof(string));
            x.Columns.Add("Birth ObjectID", typeof(string));
            x.Columns.Add("Birth Mac Addr", typeof(string));
            x.Columns.Add("MD5 Hash", typeof(string));
        }

        private int parselnk_dowork(string[] selected_folder, DataTable datatable)
        {
            int filecount = 0;
            foreach (string file in selected_folder)
            {
                if (Path.GetExtension(file) == ".lnk")
                {
                    filecount++;
                    statuslabel_lnk_filebeingparsed.Text = String.Concat("Currently parsing: ", Path.GetFileName(file));
                    parselnkstream(File.Open(file, FileMode.Open, FileAccess.Read, FileShare.Read), file, datatable, 1, null);
                }
            }
            return filecount;
        }

        private int parsei30_dowork(string[] selected_folder, DataTable datatable)
        {
            int filecount = 0;
            foreach (string file in selected_folder)
            {
                filecount++;
                statuslabel_i30_filebeingparsed.Text = String.Concat("Currently parsing: ", Path.GetFileName(file));
                parsei30stream(File.Open(file, FileMode.Open, FileAccess.Read, FileShare.Read), file, datatable);
            }
            return filecount;
        }

        private void parselnk_finalisetable(DataTable datatable, DataGridView gridview, Panel panel)
        {
            // Set the gridview datasource.
            gridview.DataSource = datatable;

            // Sets the formatting of the grid.
            gridview.Columns["LNK File Creation Time (Local)"].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm:ss";
            gridview.Columns["LNK File Access Time (Local)"].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm:ss";
            gridview.Columns["LNK File Written Time (Local)"].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm:ss";
            gridview.Columns["LNK File Creation Time (Local)"].Width = 111;
            gridview.Columns["LNK File Access Time (Local)"].Width = 111;
            gridview.Columns["LNK File Written Time (Local)"].Width = 111;
            gridview.Columns["Embedded Creation Time (Local)"].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm:ss";
            gridview.Columns["Embedded Access Time (Local)"].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm:ss";
            gridview.Columns["Embedded Written Time (Local)"].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm:ss";
            gridview.Columns["Embedded Creation Time (Local)"].Width = 111;
            gridview.Columns["Embedded Access Time (Local)"].Width = 111;
            gridview.Columns["Embedded Written Time (Local)"].Width = 111;

            gridview.Refresh();
            dataGridView_lnk.Visible = true;
        }

        private void parsei30_finalisetable(DataTable datatable, DataGridView gridview, Panel panel)
        {
            // Set the gridview datasource.
            gridview.DataSource = datatable;

            // Sets the formatting of the grid.

            gridview.Refresh();
            dataGridView_i30.Visible = true;
        }

        private void parselnkstream(Stream x, string i, DataTable gv, int type, string dir_name)
        {
            Trace.Write(string.Concat(x.Length.ToString(), " Stream length\n"));
            // type = 1 for normal lnk stream.
            // type = 2 for lnk stream within jump-list.

            // Define our global variables.
            int lnkfilefalse_counter = 0, lnkfiletrue_counter = 0, header, temp;
            uint filesize = 0;
            byte[] temp1 = new byte[1];
            byte[] temp2 = new byte[2];
            byte[] temp4 = new byte[4];
            byte[] temp8 = new byte[8];
            byte[] temp6 = new byte[6];
            byte[] temp16 = new byte[16];

            // for the embedded date and time
            DateTime? CreationTime = null;
            DateTime? AccessTime = null;
            DateTime? WriteTime = null;

            // for the LNK file date and time we need to set the variables and extract the information
            DateTime? CreationTime_lnkfile = File.GetCreationTime(i);
            DateTime? AccessTime_lnkfile = File.GetLastAccessTime(i);
            DateTime? WriteTime_lnkfile = File.GetLastWriteTime(i);
            

            // set vars
            int LinkInfoSize = 0;
            int LocalBasePathOffset, VolumeIDOffset, VolumeIDstructure_start, VolumeIDSize, CommonNetworkRelativeLinkOffset, CommonNetworkRelativeLinkSize;
            int NetNameOffset, DeviceNameOffset, VolumeLabelOffset, StringData_start, CommonPathSuffixOffset, Sig_extra;
            string VolumeSerialNumber = "", LocalBasePath = "", DriveType1 = "";
            string LinkInfoFlags = "", CommonNetworkRelativeLinkFlags = "", NetName = "";
            string DeviceName = "", NetworkProviderType = "", LinkedPath = "", VolumeLabel = "", IconLocation = "";
            string birthmacaddr = "", newmacaddr = "";

            // set StringData vars
            string NameString = "", RelativePath = "", WorkingDir = "", CommandLineArgs = "", NetBIOS = "";
            string VolumeID_current = "", ObjectID_current = "";
            string VolumeID_birth = "", ObjectID_birth = "";

            // Get MD5 hash.
            string md5hash = GetMD5HashFromStream(x);

            using (BinaryReader R = new BinaryReader(x))
                try
                {
                    // lets go from the start
                    R.BaseStream.Position = 0;

                    // read the header of a file, should be '76' in value
                    R.Read(temp4, 0, 4);
                    header = BitConverter.ToInt32(temp4, 0); //convert to a 32 bit value

                    // read lnk flags
                    R.BaseStream.Position = 20;
                    R.Read(temp4, 0, 1);
                    temp = BitConverter.ToInt32(temp4, 0); // convert bytes to number
                    string lnkflags = Convert.ToString(temp, 2); // convert number to binary
                    char[] charArray = lnkflags.ToCharArray();
                    Array.Reverse(charArray); // reverse the string
                    lnkflags = new string(charArray);
                    Trace.Write(string.Concat(R.BaseStream.Position.ToString(), " ", lnkflags, "LNKflags read\n"));

                    if (header == 76 && temp > 0) // check to see if this file is a valid lnk file
                    {
                        Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "Valid LNK file\n"));
                        lnkfiletrue_counter++; // increment the good lnk file counter

                        // read all times in UTC
                        string timezone = "utc";
                        CreationTime = ReadTargetLnkTime(28, R, temp8, timezone, "no", 0);
                        AccessTime = ReadTargetLnkTime(36, R, temp8, timezone, "no", 0);
                        WriteTime = ReadTargetLnkTime(44, R, temp8, timezone, "no", 0);

                        // read target file size
                        R.BaseStream.Position = 52;
                        R.Read(temp4, 0, 4);
                        filesize = BitConverter.ToUInt32(temp4, 0);

                        // read HasLinkTargetIDList flag, if set (1) then continue
                        int IDListSize = 0;
                        if (lnkflags[0] == '1')
                        {
                            Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "Has LinkTargetIDList flag\n"));
                            R.BaseStream.Position = 76;
                            R.Read(temp4, 0, 2);
                            IDListSize = BitConverter.ToInt16(temp4, 0);
                        }
                        else if (lnkflags[0] == '0')
                        {
                            IDListSize = -2; // we need -2 here to correct the inkinfostructure_start offset
                        } 

                        // read HasLinkInfo flag, if set (1) then continue
                        if (lnkflags[1] == '1')
                        {
                            int lnkinfostructure_start = 78 + IDListSize; // save the start of the LNKINFO structure
                            // read LinkInfoSize
                            R.BaseStream.Position = lnkinfostructure_start;
                            R.Read(temp4, 0, 4);
                            LinkInfoSize = BitConverter.ToInt32(temp4, 0);
                            Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "LNKINFO position start\n"));

                            // read LinkInfoFlags
                            R.BaseStream.Position = lnkinfostructure_start + 8;
                            R.Read(temp4, 0, 1);
                            temp = BitConverter.ToInt32(temp4, 0); // convert bytes to number
                            LinkInfoFlags = Convert.ToString(temp, 2); // convert number to binary
                            // reverse the string
                            char[] charArray1 = LinkInfoFlags.ToCharArray();
                            Array.Reverse(charArray1);
                            LinkInfoFlags = new string(charArray1);
                            lnkflags = lnkflags + " " + LinkInfoFlags;
                            Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "LinkInfoFlags read\n"));

                            // READ OFFSETS
                            // read Volume ID Offset
                            R.BaseStream.Position = lnkinfostructure_start + 12;
                            R.Read(temp4, 0, 4);
                            VolumeIDOffset = BitConverter.ToInt32(temp4, 0);

                            // read LocalBasePathOffset
                            R.BaseStream.Position = lnkinfostructure_start + 16;
                            R.Read(temp4, 0, 4);
                            LocalBasePathOffset = BitConverter.ToInt32(temp4, 0);

                            // read common network relative link offset
                            R.BaseStream.Position = lnkinfostructure_start + 20;
                            R.Read(temp4, 0, 4);
                            CommonNetworkRelativeLinkOffset = BitConverter.ToInt32(temp4, 0);

                            // read common path suffix offset
                            R.BaseStream.Position = lnkinfostructure_start + 24;
                            R.Read(temp4, 0, 4);
                            CommonPathSuffixOffset = BitConverter.ToInt32(temp4, 0);

                            if (LinkInfoFlags[0] == '1') // we have a volumeid and local base path!
                            {
                                Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "we have a volumeid and local base path!\n"));
                                VolumeIDstructure_start = lnkinfostructure_start + VolumeIDOffset;
                                // lets goto the volumeID structre
                                // Read VolumeIDSize
                                R.BaseStream.Position = VolumeIDstructure_start;
                                R.Read(temp4, 0, 4);
                                VolumeIDSize = BitConverter.ToInt32(temp4, 0);

                                // read drive type
                                R.BaseStream.Position = VolumeIDstructure_start + 4;
                                R.Read(temp4, 0, 4);
                                int DriveType1num = BitConverter.ToInt32(temp4, 0);

                                switch (DriveType1num)
                                {
                                    case 0:
                                        DriveType1 = "Unknown"; break;
                                    case 1:
                                        DriveType1 = "No Root Directory"; break;
                                    case 2:
                                        DriveType1 = "Removable"; break;
                                    case 3:
                                        DriveType1 = "Fixed Media"; break;
                                    case 4:
                                        DriveType1 = "Remove (network)"; break;
                                    case 5:
                                        DriveType1 = "CD-ROM Drive"; break;
                                    case 6:
                                        DriveType1 = "RAM Disk"; break;
                                }
                                Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "Drive type read\n"));

                                // read VolumeSerialNumber
                                R.BaseStream.Position = VolumeIDstructure_start + 8;
                                R.Read(temp4, 0, 4);
                                VolumeSerialNumber = bytestostring_littleendian(temp4);

                                // read VolumeLabelOffset
                                R.BaseStream.Position = VolumeIDstructure_start + 12;
                                R.Read(temp4, 0, 4);
                                VolumeLabelOffset = BitConverter.ToInt32(temp4, 0);

                                // read VolumeLabel
                                int tempor = VolumeIDstructure_start + VolumeLabelOffset;
                                R.BaseStream.Position = tempor;

                                int byte11 = 1;
                                string temp11 = "";
                                R.BaseStream.Position = tempor;
                                // we need to check the first byte as this is sometimes null
                                R.Read(temp4, 0, 1);
                                byte11 = BitConverter.ToInt16(temp4, 0);
                                // if the first byte isn't null then lets read the rest of the string
                                while (byte11 != 0) // save UNC path as a string, stop when you get to null
                                {
                                    R.BaseStream.Position = tempor;
                                    R.Read(temp1, 0, 1);
                                    R.Read(temp4, 0, 1); // we have to read this into a 4 byte array (there is no ToInt8)
                                    byte11 = BitConverter.ToInt16(temp4, 0); // check for null
                                    temp11 = System.Text.Encoding.Default.GetString(temp1);
                                    //MessageBox.Show(byte11.ToString());
                                    VolumeLabel = VolumeLabel + temp11;
                                    tempor++;
                                }
                                Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "VolumeLabel read\n"));

                                // now lets read the local base path
                                tempor = lnkinfostructure_start + LocalBasePathOffset;
                                R.BaseStream.Position = tempor;

                                byte11 = 1;
                                temp11 = "";
                                while (byte11 != 0) // save UNC path as a string, stop when you get to null
                                {
                                    R.BaseStream.Position = tempor;
                                    R.Read(temp1, 0, 1);
                                    R.Read(temp4, 0, 1); // we have to read this into a 4 byte array (there is no ToInt8)
                                    byte11 = BitConverter.ToInt16(temp4, 0); // check for null
                                    temp11 = System.Text.Encoding.Default.GetString(temp1);
                                    LocalBasePath = LocalBasePath + temp11;
                                    tempor++;
                                }
                                LinkedPath = LocalBasePath;
                                Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "LocalBasePath read\n"));
                            }
                            else
                            {
                                LocalBasePath = "";
                            }
                            Trace.Write(string.Concat(R.BaseStream.Position.ToString(), " ", LinkInfoFlags, "LocalBasePath read\n"));

                            if (LinkInfoFlags.Length > 1) // then we have a common network relative link and path suffix
                            {
                                Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "we have a common network relative link and path suffix\n"));
                                string CommonPathSuffix = "";
                                if (LinkInfoFlags[1] == '1')
                                {
                                    // now lets go to this offset for the common network relative link
                                    int CommonNetworkRelativeLink = lnkinfostructure_start + CommonNetworkRelativeLinkOffset;
                                    R.BaseStream.Position = CommonNetworkRelativeLink;

                                    // read CommonNetworkRelativeLinkSize
                                    R.BaseStream.Position = CommonNetworkRelativeLink;
                                    R.Read(temp4, 0, 4);
                                    CommonNetworkRelativeLinkSize = BitConverter.ToInt32(temp4, 0);

                                    // read CommonNetworkRelativeLinkFlags
                                    R.BaseStream.Position = CommonNetworkRelativeLink + 4;
                                    R.Read(temp4, 0, 1);
                                    temp = BitConverter.ToInt32(temp4, 0); // convert bytes to number
                                    CommonNetworkRelativeLinkFlags = Convert.ToString(temp, 2); // convert number to binary
                                    // reverse the string
                                    char[] charArray2 = CommonNetworkRelativeLinkFlags.ToCharArray();
                                    Array.Reverse(charArray2);
                                    CommonNetworkRelativeLinkFlags = new string(charArray2);

                                    // read NetNameOffset
                                    R.BaseStream.Position = CommonNetworkRelativeLink + 8;
                                    R.Read(temp4, 0, 4);
                                    NetNameOffset = BitConverter.ToInt32(temp4, 0);

                                    // read DeviceNameOffset
                                    R.BaseStream.Position = CommonNetworkRelativeLink + 12;
                                    R.Read(temp4, 0, 4);
                                    DeviceNameOffset = BitConverter.ToInt32(temp4, 0);

                                    // read NetworkProviderType (maybe not needed)
                                    R.BaseStream.Position = CommonNetworkRelativeLink + 16;
                                    R.Read(temp4, 0, 4);
                                    NetworkProviderType = BitConverter.ToString(temp4, 0);

                                    // read NetName
                                    int tempor = CommonNetworkRelativeLink + NetNameOffset;
                                    R.BaseStream.Position = tempor;

                                    int byte11 = 1;
                                    string temp11 = "";
                                    while (byte11 != 0) // save UNC path as a string, stop when you get to null
                                    {
                                        R.BaseStream.Position = tempor;
                                        R.Read(temp1, 0, 1);
                                        R.Read(temp4, 0, 1); // we have to read this into a 4 byte array (there is no ToInt8)
                                        byte11 = BitConverter.ToInt16(temp4, 0); // check for null
                                        temp11 = System.Text.Encoding.Default.GetString(temp1);
                                        NetName = NetName + temp11;
                                        tempor++;
                                    }
                                    Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "Netname read\n"));

                                    if (CommonNetworkRelativeLinkFlags[0] == '1') // we have a valid device!
                                    {
                                        Trace.Write(String.Concat(R.BaseStream.Position.ToString() + ";" + i + ";" + "Link file has a valid device name\n"));
                                        // read DeviceName
                                        tempor = CommonNetworkRelativeLink + DeviceNameOffset;
                                        R.BaseStream.Position = tempor;

                                        //we need to check for null straight away
                                        R.Read(temp4, 0, 1); // we have to read this into a 4 byte array (there is no ToInt8)
                                        byte11 = BitConverter.ToInt16(temp4, 0); // check for null

                                        byte11 = 1;
                                        temp11 = "";
                                        while (byte11 != 0) // save UNC path as a string, stop when you get to null
                                        {
                                            R.BaseStream.Position = tempor;
                                            R.Read(temp1, 0, 1);
                                            R.Read(temp4, 0, 1); // we have to read this into a 4 byte array (there is no ToInt8)
                                            byte11 = BitConverter.ToInt16(temp4, 0); // check for null
                                            temp11 = System.Text.Encoding.Default.GetString(temp1);
                                            DeviceName = DeviceName + temp11;
                                            tempor++;
                                        }
                                        DeviceName = String.Concat(DeviceName, @"\");
                                        Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "Device name read\n"));
                                    }
                                    else
                                    {
                                        DeviceName = "";
                                    }

                                    // now we go to the offset for the common path suffix
                                    tempor = lnkinfostructure_start + CommonPathSuffixOffset;
                                    R.BaseStream.Position = tempor;

                                    byte11 = 1;
                                    temp11 = "";
                                    while (byte11 != 0) // save UNC path as a string, stop when you get to null
                                    {
                                        R.BaseStream.Position = tempor;
                                        R.Read(temp1, 0, 1);
                                        R.Read(temp4, 0, 1); // we have to read this into a 4 byte array (there is no ToInt8)
                                        byte11 = BitConverter.ToInt16(temp4, 0); // check for null
                                        temp11 = System.Text.Encoding.Default.GetString(temp1);
                                        CommonPathSuffix = CommonPathSuffix + temp11;
                                        tempor++;
                                    }
                                    LinkedPath = DeviceName + LinkedPath + CommonPathSuffix;
                                    Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "FULL linked path read\n"));
                                }
                            }
                            else
                            {
                                NetName = "";
                                DeviceName = "";
                            }
                        }
                        Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "End of LNKINFO section\n"));

                        // read StringData
                        if (lnkflags[1] == '0') StringData_start = 78 + IDListSize; // as there is no LinkInfo structure
                        else StringData_start = 78 + IDListSize + LinkInfoSize; // as there is a LinkInfo structure, we need to go past this

                        int CurrentPosition;
                        CurrentPosition = StringData_start;

                        if (lnkflags[2] == '1') NameString = parseStringData(R, CurrentPosition, out CurrentPosition);
                        if (lnkflags[3] == '1') RelativePath = parseStringData(R, CurrentPosition, out CurrentPosition);
                        if (lnkflags[4] == '1') WorkingDir = parseStringData(R, CurrentPosition, out CurrentPosition);
                        if (lnkflags[5] == '1') CommandLineArgs = parseStringData(R, CurrentPosition, out CurrentPosition);
                        if (lnkflags[6] == '1') IconLocation = parseStringData(R, CurrentPosition, out CurrentPosition);

                        int BlockSize_extra;
                        // we need to parse the ExtraData section now
                        int endoflnkfile = (int)R.BaseStream.Length - 4; // the end of the file padded with 4 nulls
                        Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "End of LNK file is ", endoflnkfile.ToString(), "\n"));

                        while (R.BaseStream.Position < endoflnkfile)
                        {
                            Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "StringData loop position\n"));
                            //first of all, lets read the ExtraData BlockSize
                            R.BaseStream.Position = CurrentPosition;
                            R.Read(temp4, 0, 4);
                            BlockSize_extra = BitConverter.ToInt32(temp4, 0);
                            // now we need to read the signature
                            // read sig
                            R.BaseStream.Position = CurrentPosition + 4;
                            R.Read(temp2, 0, 2);
                            Sig_extra = BitConverter.ToInt16(temp2, 0);
                            Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "Signature is ", Sig_extra, "\n"));

                            string typeofblock = "";
                            int exit_now = 0;
                            switch (Sig_extra)
                            {
                                case 2:
                                    typeofblock = "ConsoleDataBlock";
                                    break;
                                case 4:
                                    typeofblock = "ConsoleFEDataBlock";
                                    break;
                                case 6:
                                    typeofblock = "DarwinDataBlock";
                                    break;
                                case 1:
                                    typeofblock = "EnvironmentVariableDataBlock";
                                    break;
                                case 7:
                                    typeofblock = "IconEnvironmentDataBlock";
                                    break;
                                case 11:
                                    typeofblock = "KnownFolderDataBlock";
                                    break;
                                case 9:
                                    typeofblock = "PropertyStoreDataBlock";
                                    break;
                                case 8:
                                    typeofblock = "ShimDataBlock";
                                    break;
                                case 5:
                                    typeofblock = "SpecialFolderDataBlock";
                                    break;
                                case 3:
                                    typeofblock = "TrackerDataBlock";
                                    // parse the NetBIOS name
                                    int tempor = CurrentPosition + 16;
                                    R.BaseStream.Position = tempor;

                                    int byte11 = 1;
                                    string temp11 = "";
                                    while (byte11 != 0) // save UNC path as a string, stop when you get to null
                                    {
                                        Trace.Write(string.Concat(R.BaseStream.Position.ToString(), "Inside the tracker block netBIOS name\n"));
                                        R.BaseStream.Position = tempor;
                                        if (R.BaseStream.Position > endoflnkfile + 3) break; // we need this for carved LNKs
                                        R.Read(temp1, 0, 1);
                                        R.Read(temp4, 0, 1); // we have to read this into a 4 byte array (there is no ToInt8)
                                        byte11 = BitConverter.ToInt16(temp4, 0); // check for null
                                        temp11 = System.Text.Encoding.Default.GetString(temp1);
                                        NetBIOS = NetBIOS + temp11;
                                        tempor++;
                                    }
                                    Trace.Write(String.Concat(R.BaseStream.Position.ToString() + ";" + i + ";" + "NetBIOS name parsed\n"));

                                    // we need this to get the current Volume/Object ID
                                    int VolumeID_offset = CurrentPosition + BlockSize_extra - 64;
                                    R.BaseStream.Position = VolumeID_offset;
                                    R.Read(temp16, 0, 16);
                                    VolumeID_current = bytestostring(temp16, null);
                                    Trace.Write(String.Concat(R.BaseStream.Position.ToString() + ";" + i + ";" + "Volume ID parsed\n"));
                                    // we need this to get the current Object ID
                                    R.BaseStream.Position = VolumeID_offset + 16;
                                    R.Read(temp16, 0, 16);
                                    ObjectID_current = bytestostring(temp16, null);
                                    Trace.Write(String.Concat(R.BaseStream.Position.ToString() + ";" + i + ";" + "Object ID parsed\n"));
                                    // extract the MAC address
                                    R.BaseStream.Position = VolumeID_offset + 26;
                                    R.Read(temp6, 0, 6);
                                    newmacaddr = bytestostring(temp6, "mac");

                                    // we need this to get the birth Volume/Object ID
                                    VolumeID_offset = CurrentPosition + BlockSize_extra - 32;
                                    R.BaseStream.Position = VolumeID_offset;
                                    R.Read(temp16, 0, 16);
                                    VolumeID_birth = bytestostring(temp16, null);
                                    Trace.Write(String.Concat(R.BaseStream.Position.ToString() + ";" + i + ";" + "Birth volume ID parsed\n"));
                                    // we need this to get the birth Object ID
                                    R.BaseStream.Position = VolumeID_offset + 16;
                                    R.Read(temp16, 0, 16);
                                    ObjectID_birth = bytestostring(temp16, null);
                                    Trace.Write(String.Concat(R.BaseStream.Position.ToString() + ";" + i + ";" + "Birth object ID parsed\n"));
                                    // extract the birth MAC address
                                    R.BaseStream.Position = VolumeID_offset + 26;
                                    R.Read(temp6, 0, 6);
                                    birthmacaddr = bytestostring(temp6, "mac");

                                    break;
                                case 12:
                                    typeofblock = "VistaAndAboveIDListDataBlock";
                                    Trace.Write(String.Concat(R.BaseStream.Position.ToString() + ";" + i + ";" + "Type of block is " + typeofblock + "\n"));
                                    break;
                                default:
                                    exit_now = 1;
                                    Trace.Write(String.Concat(R.BaseStream.Position.ToString() + ";" + i + ";" + "Not a valid block - end parsing NOW\n"));
                                    break;
                            }

                            if (exit_now == 1) { break; } // check for exit invalid link file
                            // move to the next extra data block
                            CurrentPosition = CurrentPosition + BlockSize_extra;
                            if (CurrentPosition == endoflnkfile) { break; } //for some reason we need this
                        }                 
                    }
                    else
                    {
                        lnkfilefalse_counter++;
                    }
                }
                catch (Exception e)
                {
                    Trace.Write("{0} Exception caught.", e.ToString());
                    MessageBox.Show(String.Concat("Error parsing .lnk file: " + Path.GetFileName(i)), "Error"); //e.ToString(),Path.GetFileName(i));
                    Trace.Flush();
                }
                finally // write data into database
                {
                    if (type == 1) // normal LNK file
                    {
                        //toUTClnkToolStripMenuItem.Visible = true;
                        gv.Rows.Add(Path.GetFileName(i), LinkedPath, CreationTime_lnkfile, AccessTime_lnkfile, WriteTime_lnkfile, CreationTime, AccessTime,
                            WriteTime, filesize, NetName, NetBIOS, DriveType1, VolumeSerialNumber, VolumeLabel, NameString, RelativePath,
                            WorkingDir, CommandLineArgs, VolumeID_current, ObjectID_current, newmacaddr, VolumeID_birth, ObjectID_current, birthmacaddr, md5hash);
                    }
                    else if (type == 2) // jump-list embedded file
                    {
                        //toUTCjllnkToolStripMenuItem.Visible = true;
                        gv.Rows.Add(dir_name, LinkedPath, CreationTime, AccessTime,
                        WriteTime, filesize, NetName, NetBIOS, DriveType1, VolumeSerialNumber, VolumeLabel, NameString, RelativePath,
                        WorkingDir, CommandLineArgs, VolumeID_current, ObjectID_current, newmacaddr, VolumeID_birth, ObjectID_current, birthmacaddr, md5hash);
                    }
                }
        }

        private void parsei30stream(Stream x, string i, DataTable gv)
        {
            Trace.Write(string.Concat(x.Length.ToString(), " Stream length\n"));

            // Define global vars
            byte[] temp1 = new byte[1];
            byte[] temp2 = new byte[2];
            byte[] temp4 = new byte[4];
            byte[] temp8 = new byte[8];
            byte[] temp6 = new byte[6];
            byte[] temp16 = new byte[16];

            // Define our header variables
            int header, index_entry_offset, sizeof_entries, sizeof_entries_alloc, numberof_records;
            int updated_sequence_array_offset, sizeof_updated_sequence_array;
            long logfile_seq_num;

            // we need this for the "fix-up" values to be applied
            int update_seq_number;

            // Define our INDX variables
            int indx_record_start;
            long mftref = 0, parent_mftref;
            int sizeof_indx_entry, sizeof_stream;
            DateTime? Creation_Time = null;
            DateTime? Modified_Time = null;
            DateTime? Access_Time = null;
            DateTime? mft_record_change_time = null;
            long phy_sizeof_file, logical_sizeof_file, start_of_block;
            int filename_length;
            string embedded_filename = "", slack_space = "";

            using (BinaryReader R = new BinaryReader(x))
                try
                {
                    Trace.Write("Starting $I30 parsing\n");
                    // lets go from the start
                    R.BaseStream.Position = 0;
                    start_of_block = 0;

                    // read the header
                    R.Read(temp4, 0, 4);
                    header = BitConverter.ToInt32(temp4, 0); //convert to a 32 bit value

                    if (header == 1480871497) // decimal of string == INDX
                    {
                        Trace.Write("Valid $I30 file found\n");
                        // we have a valid $I30 header so lets go do stuff

                        int exit = 0;
                        while (exit < 1) // 0 = no, 1 = yes
                        {
                            R.BaseStream.Position = start_of_block;
                            start_of_block = R.BaseStream.Position;
                            Trace.Write(String.Concat(R.BaseStream.Position.ToString(), " Start of array\n"));

                            // read the header
                            R.Read(temp4, 0, 4);
                            header = BitConverter.ToInt32(temp4, 0); //convert to a 32 bit value
                            if (header == 1480871497) //valid INDX record
                            {
                                // read updated_sequence_array_offset
                                R.BaseStream.Position = start_of_block + 4;
                                R.Read(temp2, 0, 2);
                                updated_sequence_array_offset = BitConverter.ToInt16(temp2, 0);
                                Trace.Write(String.Concat(updated_sequence_array_offset, " Updated Seq Array Off.\n"));
                                // now we go to this offset and save the sequence number (the last two bytes of every block should be like this)
                                R.BaseStream.Position = start_of_block + updated_sequence_array_offset;
                                R.Read(temp2, 0, 2);
                                update_seq_number = BitConverter.ToInt16(temp2, 0);
                                // read size of updated_sequence_array
                                R.BaseStream.Position = start_of_block + 6;
                                R.Read(temp2, 0, 2);
                                sizeof_updated_sequence_array = BitConverter.ToInt16(temp2, 0);
                                Trace.Write(String.Concat(sizeof_updated_sequence_array, " Size of above\n"));

                                // now we have to read the updated sequence array and save it as an.. array!
                                R.BaseStream.Position = start_of_block + updated_sequence_array_offset + 2;
                                int[] seq_array = new int[sizeof_updated_sequence_array]; 
                                for (int counter = 0; counter < sizeof_updated_sequence_array; counter++)
                                {
                                    R.Read(temp2, 0, 2);
                                    seq_array[counter] = BitConverter.ToInt16(temp2, 0);
                                    //MessageBox.Show(string.Concat(update_seq_number, " should be ", seq_array[counter]));
                                }

                                //// now we need to go through the stream and edit all the fix-up values
                                //// I know its bad to edit a file in forensics but here we are only modifying the stream in memory
                                //long temppor = start_of_block;
                                //for (int counter = 0; counter < sizeof_updated_sequence_array; counter++)
                                //{
                                //    R.BaseStream.Position = temppor + 510; // to read the last two bytes in the sector
                                //    R.Read(temp2, 0, 2);
                                //    int temp = BitConverter.ToInt16(temp2, 0);
                                //    if (temp == update_seq_number)
                                //    {
                                //        BinaryWriter W = new BinaryWriter
                                //        MessageBox.Show(string.Concat("Do something! ", temp, " should be ", seq_array[counter]));
                                //    }
                                //}

                                // read log file seq. number
                                R.BaseStream.Position = start_of_block + 8;
                                R.Read(temp8, 0, 8);
                                logfile_seq_num = BitConverter.ToInt64(temp8, 0);
                                Trace.Write(String.Concat(logfile_seq_num, " log file seq num\n"));

                                // read index entry offset
                                R.BaseStream.Position = start_of_block + 24;
                                R.Read(temp4, 0, 4);
                                index_entry_offset = BitConverter.ToInt32(temp4, 0);
                                Trace.Write(String.Concat(index_entry_offset, " Index entry offset\n"));

                                // size of entries
                                R.BaseStream.Position = start_of_block + 28;
                                R.Read(temp4, 0, 4);
                                sizeof_entries = BitConverter.ToInt32(temp4, 0);
                                Trace.Write(String.Concat(sizeof_entries, " size of entries\n"));

                                // size of entries (allocated)
                                R.BaseStream.Position = start_of_block + 32;
                                R.Read(temp4, 0, 4);
                                sizeof_entries_alloc = BitConverter.ToInt32(temp4, 0);
                                Trace.Write(String.Concat(sizeof_entries_alloc, " sizeof entries allocated\n"));

                                int end_of_block = (int)Math.Ceiling((decimal)sizeof_entries_alloc / 512);
                                end_of_block = end_of_block * 512;

                                // lets go through all the ALLOCATED records
                                indx_record_start = (int)start_of_block + 64; // first record is located here
                                while (R.BaseStream.Position < (start_of_block + sizeof_entries))
                                {
                                    slack_space = "No";
                                    // now lets go parse the INDX records
                                    R.BaseStream.Position = indx_record_start;

                                    // read MFT ref
                                    R.Read(temp8, 0, 8);
                                    mftref = BitConverter.ToInt64(temp8, 0);
                                    Trace.Write(String.Concat(mftref, " MFT ref\n"));

                                    // read size of INDX entry
                                    R.BaseStream.Position = indx_record_start + 8;
                                    R.Read(temp2, 0, 2);
                                    sizeof_indx_entry = BitConverter.ToInt16(temp2, 0);
                                    Trace.Write(String.Concat(sizeof_indx_entry, " Size of index entry\n"));

                                    // read size of stream
                                    R.BaseStream.Position = indx_record_start + 10;
                                    R.Read(temp2, 0, 2);
                                    sizeof_stream = BitConverter.ToInt16(temp2, 0);
                                    Trace.Write(String.Concat(sizeof_stream, " size of stream\n"));

                                    // read parent MFT ref
                                    R.BaseStream.Position = indx_record_start + 16;
                                    R.Read(temp8, 0, 8);
                                    parent_mftref = BitConverter.ToInt64(temp8, 0);
                                    Trace.Write(String.Concat(parent_mftref, " Parent MFT ref\n"));

                                    // read times - making sure to check if they exist on 512 boundaries to that fixup values can be applied
                                    // 
                                    if ((indx_record_start + 24 + 8) % 512 == 0)
                                    {
                                        // exists on boundary, we must edit this hex before passing it to "ReadtargetLNKtime"
                                        // we need to see how far the boundary is...
                                        // first off we need to start at the 4096 boundary...
                                        int temp = ((indx_record_start + 24 + 8) / 4096); // this is to see how far along we are. for example 4608/4096 = 1.125. This is saved as 1 (as it is an int). So 1 * 4096 = 4096.
                                        temp = temp * 4096;
                                        // now we need to see how far we are along the 4096 sector..
                                        int unique_sec_number_inarray = ((indx_record_start - temp + 24 + 8) / 512 - 1);
                                        Creation_Time = ReadTargetLnkTime(indx_record_start + 24, R, temp8, "local", "yes", seq_array[unique_sec_number_inarray]);
                                    }
                                    else
                                    {
                                        Creation_Time = ReadTargetLnkTime(indx_record_start + 24, R, temp8, "local", "no", 0);
                                    }
                                    Trace.Write(String.Concat(Creation_Time, " Creation\n"));

                                    if ((indx_record_start + 32 + 8) % 512 == 0)
                                    {
                                        // exists on boundary, we must edit this hex before passing it to "ReadtargetLNKtime"
                                        // we need to see how far the boundary is...
                                        // first off we need to start at the 4096 boundary...
                                        int temp = ((indx_record_start + 32 + 8) / 4096); // this is to see how far along we are. for example 4608/4096 = 1.125. This is saved as 1 (as it is an int). So 1 * 4096 = 4096.
                                        temp = temp * 4096;
                                        // now we need to see how far we are along the 4096 sector..
                                        int unique_sec_number_inarray = ((indx_record_start - temp + 32 + 8) / 512 - 1);
                                        Modified_Time = ReadTargetLnkTime(indx_record_start + 32, R, temp8, "local", "yes", seq_array[unique_sec_number_inarray]);
                                    }
                                    else
                                    {
                                        Modified_Time = ReadTargetLnkTime(indx_record_start + 32, R, temp8, "local", "no", 0);
                                    }
                                    Trace.Write(String.Concat(Modified_Time, " Modified\n"));

                                    mft_record_change_time = ReadTargetLnkTime(indx_record_start + 40, R, temp8, "local", "no", 0);
                                    Access_Time = ReadTargetLnkTime(indx_record_start + 48, R, temp8, "local", "no", 0);

                                    // read physical size of file
                                    R.BaseStream.Position = indx_record_start + 56;
                                    R.Read(temp8, 0, 8);
                                    phy_sizeof_file = BitConverter.ToInt64(temp8, 0);

                                    // read logical size of file
                                    R.BaseStream.Position = indx_record_start + 64;
                                    R.Read(temp8, 0, 8);
                                    logical_sizeof_file = BitConverter.ToInt64(temp8, 0);
                                    Trace.Write(String.Concat(logical_sizeof_file, " Logical size\n"));

                                    // read filename length
                                    R.BaseStream.Position = indx_record_start + 80;
                                    R.Read(temp2, 0, 1);
                                    filename_length = BitConverter.ToInt16(temp2, 0);
                                    filename_length = filename_length * 2; // as we are reading the values in unicode
                                    Trace.Write(String.Concat(filename_length, " File name length\n"));

                                    // go read filename
                                    // read entry name
                                    embedded_filename = "";
                                    string temp11 = "";
                                    int tempor = indx_record_start + 82;
                                    while (tempor < indx_record_start + 82 + filename_length)
                                    {
                                        R.BaseStream.Position = tempor;
                                        R.Read(temp1, 0, 1);
                                        temp11 = System.Text.Encoding.Default.GetString(temp1);
                                        embedded_filename = embedded_filename + temp11;
                                        tempor = tempor + 2; // skip the null-delimited
                                    }

                                    int end_of_record = (int)Math.Ceiling((decimal)R.BaseStream.Position / 8);
                                    end_of_record = end_of_record * 8;

                                    // put data in the database
                                    gv.Rows.Add(Path.GetFileName(i), embedded_filename, slack_space, mftref, parent_mftref, Creation_Time,
                                        Modified_Time, Access_Time, mft_record_change_time, phy_sizeof_file, logical_sizeof_file);

                                    // move the stream on
                                    if (sizeof_indx_entry > 82)
                                    {
                                        indx_record_start = indx_record_start + sizeof_indx_entry;
                                    }
                                    else if (sizeof_indx_entry <= 0)
                                    {
                                        indx_record_start = end_of_record;
                                    }
                                }

                                int end_of_allocated = (int)Math.Ceiling((decimal)R.BaseStream.Position / 8);
                                end_of_allocated = end_of_allocated * 8;

                                // now lets attempt to parse the records in UNALLOCATED
                                //MessageBox.Show(String.Concat("Position start of unalloc ", end_of_allocated.ToString()));

                                indx_record_start = end_of_allocated; // move to the unallocated space
                                R.BaseStream.Position = indx_record_start;
                                int inc = 0;
                                while (R.BaseStream.Position < (start_of_block + end_of_block))
                                {
                                    Trace.Write(String.Concat("Position: ", R.BaseStream.Position.ToString(), "\n"));

                                    //first, we did to check this record has valid data! As it can be corrupt... best way to check this is with the 4 date/time stamps
                                    // read times, we need to redefine these for some reason
                                    DateTime Creation = ReadTargetLnkTime_noq(indx_record_start + 24, R, temp8, "local");
                                    DateTime Modified = ReadTargetLnkTime_noq(indx_record_start + 32, R, temp8, "local");
                                    DateTime mft_record = ReadTargetLnkTime_noq(indx_record_start + 40, R, temp8, "local");
                                    DateTime Access = ReadTargetLnkTime_noq(indx_record_start + 48, R, temp8, "local");
                                    DateTime comp = new DateTime(1990, 1, 1, 0, 0, 0);

                                    if (DateTime.Compare(Creation, comp) + DateTime.Compare(Modified, comp) + DateTime.Compare(mft_record, comp) + DateTime.Compare(Access, comp) == 4)
                                    {
                                        // now lets go parse the INDX records
                                        R.BaseStream.Position = indx_record_start;
                                        slack_space = String.Concat("Yes: at offset ", R.BaseStream.Position.ToString("X"));

                                        // read MFT ref
                                        R.Read(temp8, 0, 8);
                                        mftref = BitConverter.ToInt64(temp8, 0);
                                        Trace.Write(String.Concat(mftref, " MFT ref\n"));

                                        // read size of INDX entry
                                        R.BaseStream.Position = indx_record_start + 8;
                                        R.Read(temp2, 0, 2);
                                        sizeof_indx_entry = BitConverter.ToInt16(temp2, 0);
                                        Trace.Write(String.Concat(sizeof_indx_entry, " Size of index entry\n"));

                                        // read size of stream
                                        R.BaseStream.Position = indx_record_start + 10;
                                        R.Read(temp2, 0, 2);
                                        sizeof_stream = BitConverter.ToInt16(temp2, 0);
                                        Trace.Write(String.Concat(sizeof_stream, " size of stream\n"));

                                        // read parent MFT ref
                                        R.BaseStream.Position = indx_record_start + 16;
                                        R.Read(temp8, 0, 8);
                                        parent_mftref = BitConverter.ToInt64(temp8, 0);
                                        Trace.Write(String.Concat(parent_mftref, " Parent MFT ref\n"));

                                        // read times
                                        Creation_Time = ReadTargetLnkTime(indx_record_start + 24, R, temp8, "local", "no", 0);
                                        Trace.Write(String.Concat(Creation_Time, " Creation\n"));
                                        Modified_Time = ReadTargetLnkTime(indx_record_start + 32, R, temp8, "local", "no", 0);
                                        Trace.Write(String.Concat(Modified_Time, " Modified\n"));
                                        mft_record_change_time = ReadTargetLnkTime(indx_record_start + 40, R, temp8, "local", "no", 0);
                                        Access_Time = ReadTargetLnkTime(indx_record_start + 48, R, temp8, "local", "no", 0);

                                        // read physical size of file
                                        R.BaseStream.Position = indx_record_start + 56;
                                        R.Read(temp8, 0, 8);
                                        phy_sizeof_file = BitConverter.ToInt64(temp8, 0);

                                        // read logical size of file
                                        R.BaseStream.Position = indx_record_start + 64;
                                        R.Read(temp8, 0, 8);
                                        logical_sizeof_file = BitConverter.ToInt64(temp8, 0);
                                        Trace.Write(String.Concat(logical_sizeof_file, " Logical size\n"));

                                        // read filename length
                                        R.BaseStream.Position = indx_record_start + 80;
                                        R.Read(temp2, 0, 1);
                                        filename_length = BitConverter.ToInt16(temp2, 0);
                                        filename_length = filename_length * 2; // as we are reading the values in unicode
                                        Trace.Write(String.Concat(filename_length, " File name length\n"));

                                        // go read filename
                                        // read entry name
                                        embedded_filename = "";
                                        string temp11 = "";
                                        int tempor = indx_record_start + 82;
                                        while (tempor < indx_record_start + 82 + filename_length)
                                        {
                                            R.BaseStream.Position = tempor;
                                            R.Read(temp1, 0, 1);
                                            temp11 = System.Text.Encoding.Default.GetString(temp1);
                                            embedded_filename = embedded_filename + temp11;
                                            tempor = tempor + 2; // skip the null-delimited
                                        }
                                        Trace.Write(String.Concat(embedded_filename, " File name\n"));

                                        int end_of_record = (int)Math.Ceiling((decimal)R.BaseStream.Position / 8);
                                        end_of_record = end_of_record * 8;

                                        // put data in the database
                                        gv.Rows.Add(Path.GetFileName(i), embedded_filename, slack_space, mftref, parent_mftref, Creation_Time,
                                            Modified_Time, Access_Time, mft_record_change_time, phy_sizeof_file, logical_sizeof_file);

                                        // move the stream on
                                        if (sizeof_indx_entry > 82)
                                        {
                                            indx_record_start = indx_record_start + sizeof_indx_entry;
                                        }
                                        else
                                        {
                                            indx_record_start = end_of_record;
                                        }
                                        inc++;
                                    }
                                    else
                                    {
                                        indx_record_start = indx_record_start + 8;
                                    }
                                }

                                // move onto the next block
                                start_of_block = start_of_block + (long)end_of_block;
                                if (start_of_block >= x.Length)
                                {
                                    exit = 1;
                                }
                            }
                            else
                            {
                                start_of_block = start_of_block + 8;
                                if (start_of_block >= x.Length)
                                {
                                    exit = 1;
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Trace.Write("{0} Exception caught.", e.ToString());
                    //MessageBox.Show(String.Concat("Error parsing $I30 file: " + Path.GetFileName(i)), "Error"); //e.ToString(),Path.GetFileName(i));
                    Trace.Flush();
                }
                finally // write data into database
                {
                    
                }
        }

        private void parsedestliststream(Stream x, DataTable gv, int dircount)
        {
            dircount--; // we don't need to count the destlist node
            using (BinaryReader R = new BinaryReader(x))
                try
                {
                    int temppos = 32; // move past the header
                    for (int count1 = 0; count1 < dircount; count1++)
                    {
                        // declare our file times and vars
                        byte[] temp1 = new byte[1];
                        byte[] temp2 = new byte[2];
                        byte[] temp4 = new byte[4];
                        byte[] temp8 = new byte[8];
                        DateTime? embeddedtime_utc = null, embeddedtime_local = null;
                        string netBIOS = "";
                        string data = "";
                        R.BaseStream.Position = temppos;

                        int byte11 = 1;
                        string temp11 = "";
                        int tempor = temppos + 72;
                        R.BaseStream.Position = temppos + 72;
                        // we need to check the first byte as this is sometimes null
                        R.Read(temp4, 0, 1);
                        byte11 = BitConverter.ToInt16(temp4, 0);
                        // if the first byte isn't null then lets read the rest of the string
                        while (byte11 != 0) // save UNC path as a string, stop when you get to null
                        {
                            R.BaseStream.Position = tempor;
                            R.Read(temp1, 0, 1);
                            R.Read(temp4, 0, 1); // we have to read this into a 4 byte array (there is no ToInt8)
                            byte11 = BitConverter.ToInt16(temp4, 0); // check for null
                            temp11 = System.Text.Encoding.Default.GetString(temp1);
                            //MessageBox.Show(byte11.ToString());
                            netBIOS = netBIOS + temp11;
                            tempor++;
                        }

                        // read stream number
                        R.BaseStream.Position = temppos + 88;
                        R.Read(temp1, 0, 1);
                        string streamnumber = BitConverter.ToString(temp1, 0);

                        // read embedded time
                        embeddedtime_local = ReadTargetLnkTime(temppos + 100, R, temp8, "local", "no", 0);
                        embeddedtime_utc = ReadTargetLnkTime(temppos + 100, R, temp8, "utc", "no", 0);

                        // read number of characters for the data string
                        R.BaseStream.Position = temppos + 112;
                        R.Read(temp2, 0, 2);
                        int char_num = BitConverter.ToInt16(temp2, 0);
                        char_num = char_num * 2; // number of unicode values

                        temp11 = "";
                        tempor = temppos + 114;
                        R.BaseStream.Position = temppos + 114;
                        // we need to check the first byte as this is sometimes null
                        R.Read(temp4, 0, 1);
                        byte11 = BitConverter.ToInt16(temp4, 0);
                        // if the first byte isn't null then lets read the rest of the string
                        for (int count = 0; count < char_num; count = count + 2)
                        {
                            R.BaseStream.Position = tempor;
                            R.Read(temp1, 0, 1);
                            R.Read(temp4, 0, 1); // we have to read this into a 4 byte array (there is no ToInt8)
                            byte11 = BitConverter.ToInt16(temp4, 0); // check for null
                            temp11 = System.Text.Encoding.Default.GetString(temp1);
                            //MessageBox.Show(byte11.ToString());
                            data = data + temp11;
                            tempor = tempor + 2;
                        }

                        gv.Rows.Add(streamnumber, netBIOS, embeddedtime_local, embeddedtime_utc, data);
                        temppos = temppos + 114 + char_num;
                    }
                }
                finally
                {
                }
        }

        private void parsei30_setuptable(DataTable x, Panel panel)
        {
            tabControl_main.SelectedTab = tabControl_main.TabPages[3];
            // Make the $I30 panel visible.
            panel.Visible = true;
            dataGridView_i30.Visible = false;
            statuslabel_i30_filesparsed.Text = "";
            statuslabel_i30_timetaken.Text = "";

            // Create the columns.
            x.Columns.Add("$I30 File Name", typeof(string));
            x.Columns.Add("Embedded File Name", typeof(string));
            x.Columns.Add("Slack Record", typeof(string));
            x.Columns.Add("MFT Ref", typeof(long));
            x.Columns.Add("Parent Dir", typeof(long));
            x.Columns.Add("Creation Time (Local)", typeof(DateTime));
            x.Columns.Add("Last Modified Time (Local)", typeof(DateTime));
            x.Columns.Add("Last Access Time (Local)", typeof(DateTime));
            x.Columns.Add("MFT Record Change Time (Local)", typeof(DateTime));
            x.Columns.Add("Physical File Size (Bytes)", typeof(long));
            x.Columns.Add("Logical File Size (Bytes)", typeof(long));
        }

        private void parsejl_setuptable(DataTable x, Panel panel)
        {
            tabControl_main.SelectedTab = tabControl_main.TabPages[2];
            // Make the .jl panel visible.
            panel.Visible = true;

            // create the columns
            x.Columns.Add("FileName", typeof(string));
            x.Columns.Add("Application", typeof(string));
            x.Columns.Add("MD5 Hash", typeof(string));
            x.Columns.Add("FullPath", typeof(string));
        }

        private void parsepf_setuptable(DataTable datatable, Panel panel)
        {
            tabControl_main.SelectedTab = tabControl_main.TabPages[1];
            panel.Visible = true;
            dataGridview_pfinfo.Visible = false;
            statuslabel_pf_timetaken.Text = "";
            statuslabel_pf_filesparsed.Text = "";

            // create the columns
            datatable.Columns.Add("FileName", typeof(string));
            datatable.Columns.Add("Boot Process Name", typeof(string));
            datatable.Columns.Add("Hash Value", typeof(string));
            datatable.Columns.Add("Last Run Time (Local)", typeof(DateTime));
            datatable.Columns.Add("Run Count", typeof(int));
            datatable.Columns.Add("Version", typeof(string));
            datatable.Columns.Add("Volume Serial", typeof(string));
            datatable.Columns.Add("Volume Created Date (Local)", typeof(DateTime));
            datatable.Columns.Add("MD5 Hash", typeof(string));
            datatable.Columns.Add("FullPath", typeof(string));
        }

        private int parsepf_dowork(string[] selected_folder, DataTable datatable)
        {
            // define our variables
            int pffilefalse_counter = 0, pffiletrue_counter = 0, version, header, last_run_time_pos = 0, last_run_count_pos = 0;
            int pfruncount;
            string pf_version = "";
            byte[] temp4 = new byte[4];
            byte[] temp1 = new byte[1];
            byte[] temp2 = new byte[2];
            byte[] temp8 = new byte[8];
            DateTime LastRunTime_local;

            // lets go through all the files one by one
            foreach (string i in selected_folder)
            {
                if (Path.GetExtension(i) == ".pf")
                {
                    statuslabel_pf_filebeingparsed.Text = String.Concat("Currently parsing: ", Path.GetFileName(i));
                    string md5hash = GetMD5HashFromFile(i);
                    using (BinaryReader R = new BinaryReader(File.Open(i, FileMode.Open, FileAccess.Read, FileShare.Read)))
                    try
                    {
                        //check header string to see if it is a valid pf file
                        R.BaseStream.Position = 4;
                        R.Read(temp4, 0, 4);
                        header = BitConverter.ToInt32(temp4, 0);
                        if (header == 1094927187)
                        {
                            pffiletrue_counter++;
                            R.BaseStream.Position = 0; // lets go from the start
                            //parse information from the header structure
                            //version number
                            R.Read(temp4, 0, 4);
                            version = BitConverter.ToInt32(temp4, 0);

                            // set the positions of the Last Run Time (Local) and count
                            if (version == 17) //this is xp
                            {
                                last_run_time_pos = 120;
                                last_run_count_pos = 144;
                                pf_version = "XP";
                            }
                            else if (version == 23) //this is vista/7
                            {
                                last_run_time_pos = 128;
                                last_run_count_pos = 152;
                                pf_version = "Vista / 7";
                            }

                            // read bootprocess name
                            string bootprocess = "";
                            int tempor = 16;
                            R.BaseStream.Position = tempor;

                            int byte11 = 1;
                            string temp11 = "";
                            R.BaseStream.Position = tempor;

                            // if the first byte isn't null then lets read the rest of the string
                            while (byte11 != 0) // save name as a string, stop when you get to null
                            {
                                R.BaseStream.Position = tempor;
                                R.Read(temp1, 0, 1);
                                R.Read(temp4, 0, 2); // we have to read this into a 4 byte array (there is no ToInt8)
                                byte11 = BitConverter.ToInt16(temp4, 0); // check for null
                                temp11 = System.Text.Encoding.Default.GetString(temp1);
                                bootprocess = bootprocess + temp11;
                                tempor = tempor + 2; // we are reading a unicode value here
                            }

                            //PF hash value
                            string pfhashvalue;
                            R.BaseStream.Position = 76;
                            R.Read(temp4, 0, 4);
                            pfhashvalue = bytestostring_littleendian(temp4);

                            //get Last Run Time (Local)
                            R.BaseStream.Position = last_run_time_pos;
                            R.Read(temp8, 0, 8);
                            LastRunTime_local = DateTime.FromFileTime(BitConverter.ToInt64(temp8, 0));

                            //get run count
                            R.BaseStream.Position = last_run_count_pos;
                            R.Read(temp4, 0, 4);
                            pfruncount = BitConverter.ToInt32(temp4, 0);

                            //get file ref offset
                            R.BaseStream.Position = 100;
                            R.Read(temp4, 0, 4);
                            int file_ref_offset = BitConverter.ToInt32(temp4, 0);

                            //get file ref length
                            R.BaseStream.Position = 104;
                            R.Read(temp4, 0, 4);
                            int file_ref_length = BitConverter.ToInt32(temp4, 0);

                            //get volume information offset
                            R.BaseStream.Position = 108;
                            R.Read(temp4, 0, 4);
                            int vol_info_offset = BitConverter.ToInt32(temp4, 0);
                            //lets go get this information
                            R.BaseStream.Position = vol_info_offset + 8;
                            R.Read(temp8, 0, 8);
                            DateTime vol_create_time_local = DateTime.FromFileTime(BitConverter.ToInt64(temp8, 0));
                            R.BaseStream.Position = vol_info_offset + 16;
                            R.Read(temp4, 0, 4);
                            string vol_serial = bytestostring_littleendian(temp4);

                            // write data into database
                            datatable.Rows.Add(Path.GetFileName(i), bootprocess, pfhashvalue, LastRunTime_local, pfruncount, pf_version, 
                                vol_serial, vol_create_time_local, md5hash, i);
                        }

                        else if (header != 1094927187)
                        {
                            pffilefalse_counter++;
                        }
                    }
                    catch (Exception e)
                    {
                        Trace.Write("{0} Exception caught.", e.ToString());
                        //MessageBox.Show("Error parsing Prefetch file"); //e.ToString(),Path.GetFileName(i));
                        Trace.Flush();
                    }
                    finally
                    {
                        R.Close();
                    }
                }
                else if (Path.GetExtension(i) != ".pf")
                {
                }
            }
            if (pffilefalse_counter > 0)
            {
                MessageBox.Show(String.Concat(pffilefalse_counter, " invalid PF files found"), "Error");
            }
            return pffiletrue_counter;
        }

        private void parsepffilefolder_dowork(string[] selected_items)
        {
            // define our variables
            byte[] temp4 = new byte[4];
            byte[] temp1 = new byte[1];
            byte[] temp2 = new byte[2];
            byte[] temp8 = new byte[8];

            DataTable pfinfosubfile = new DataTable("pfinfosub");
            pfinfosubfile.Columns.Add("File Name", typeof(string));
            pfinfosubfile.Columns.Add("Path", typeof(string));
            DataTable pfinfosubfolder = new DataTable("pfinfosub");
            pfinfosubfolder.Columns.Add("File Name", typeof(string));
            pfinfosubfolder.Columns.Add("Path", typeof(string));

            foreach (string i in selected_items)
            {
                // lets go through all the files one by one
                if (Path.GetExtension(i) == ".pf")
                {
                    using (BinaryReader R = new BinaryReader(File.Open(i, FileMode.Open, FileAccess.Read, FileShare.Read)))
                        try
                        {
                            //get file ref offset
                            R.BaseStream.Position = 100;
                            R.Read(temp4, 0, 4);
                            int file_ref_offset = BitConverter.ToInt32(temp4, 0);

                            //get file ref length
                            R.BaseStream.Position = 104;
                            R.Read(temp4, 0, 4);
                            int file_ref_length = BitConverter.ToInt32(temp4, 0);

                            //get volume information offset
                            R.BaseStream.Position = 108;
                            R.Read(temp4, 0, 4);
                            int vol_info_offset = BitConverter.ToInt32(temp4, 0);

                            //get the folder reference area offset
                            R.BaseStream.Position = vol_info_offset + 28;
                            R.Read(temp4, 0, 4);
                            int folder_ref_offset = BitConverter.ToInt32(temp4, 0);
                            folder_ref_offset = vol_info_offset + folder_ref_offset + 2; // we need the +2 as it is unicode
                            R.BaseStream.Position = vol_info_offset + 28 + 4;
                            R.Read(temp4, 0, 4);
                            int folder_ref_number = BitConverter.ToInt32(temp4, 0);

                            //lets go through all the file references
                            int byte11 = 1;
                            string temp11 = "";
                            string[] filerefs = new string[5000];
                            R.BaseStream.Position = file_ref_offset;
                            int tempor = file_ref_offset;
                            int endoffileref = file_ref_offset + file_ref_length;
                            int temppos = 0;
                            while (tempor < endoffileref)
                            {
                                byte11 = 1;
                                temp11 = "";
                                // we need to check the first byte as this is sometimes null
                                R.Read(temp4, 0, 1);
                                // if the first byte isn't null then lets read the rest of the string
                                while (byte11 != 0) // save name as a string, stop when you get to null
                                {
                                    R.BaseStream.Position = tempor;
                                    R.Read(temp1, 0, 1);
                                    R.Read(temp4, 0, 2); // we have to read this into a 4 byte array (there is no ToInt8)
                                    byte11 = BitConverter.ToInt16(temp4, 0); // check for null
                                    temp11 = System.Text.Encoding.Default.GetString(temp1);
                                    filerefs[temppos] = filerefs[temppos] + temp11;
                                    tempor = tempor + 2; // we are reading a unicode value here
                                }
                                // write data into database
                                pfinfosubfile.Rows.Add(Path.GetFileName(i), filerefs[temppos]);
                                temppos++;
                                tempor = tempor + 2; // move past the null
                            }

                            //lets go through all the folder references
                            string[] folderrefs = new string[folder_ref_number];
                            R.BaseStream.Position = folder_ref_offset;
                            tempor = folder_ref_offset;
                            temppos = 0;
                            int x = 0;
                            while (x < folder_ref_number)
                            {
                                byte11 = 1;
                                temp11 = "";
                                while (byte11 != 0) // save name as a string, stop when you get to null
                                {
                                    R.BaseStream.Position = tempor;
                                    R.Read(temp1, 0, 1);
                                    R.Read(temp4, 0, 2); // we have to read this into a 4 byte array (there is no ToInt8)
                                    byte11 = BitConverter.ToInt16(temp4, 0); // check for null
                                    temp11 = System.Text.Encoding.Default.GetString(temp1);
                                    folderrefs[temppos] = folderrefs[temppos] + temp11;
                                    tempor = tempor + 2; // we are reading a unicode value here
                                }
                                // write data into database
                                pfinfosubfolder.Rows.Add(Path.GetFileName(i), folderrefs[temppos]);
                                temppos++;
                                tempor = tempor + 4; // +4 to go past the 2 nulls in unicode
                                x++;
                            }
                        }
                        catch (Exception e)
                        {
                            Trace.Write("{0} Exception caught.", e.ToString());
                            MessageBox.Show("Error parsing prefetch file/folder structures"); //e.ToString(),Path.GetFileName(i));
                            Trace.Flush();
                        }
                        finally
                        {
                            dataGridpfinfosub_fileref.DataSource = pfinfosubfile;
                            dataGridpfinfosub_folderref.DataSource = pfinfosubfolder;
                        }
                }
            }
        }

        private void parsepf_finalisetable(DataTable datatable, DataGridView gridview, Panel panel)
        {
            // Set the gridview datasource.
            gridview.DataSource = datatable;

            // set formatting of columns
            gridview.Columns["Last Run Time (Local)"].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm:ss";
            gridview.Columns["Volume Created Date (Local)"].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm:ss";
            gridview.Columns["FullPath"].Visible = false;

            //correct column sizes
            for (int i = 0; i < gridview.Columns.Count; i++)
            {
                int colw = gridview.Columns[i].Width;
                gridview.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                gridview.Columns[i].Width = colw;
            }
            gridview.Refresh();
            dataGridview_pfinfo.Visible = true;
            toUTCToolStripMenuItem_pf.Visible = true;
        }

        private DateTime? ReadTargetLnkTime(int i, BinaryReader j, byte[] k, string timezone, string fixup, int fixup_value)
        {
            // declare vars
            long filetime;
            // do work
            j.BaseStream.Position = i;
            if (fixup == "no")
            {
                j.Read(k, 0, 8);
            }
            else if (fixup == "yes")
            {
                j.Read(k, 0, 6);
                byte[] intBytes = BitConverter.GetBytes(fixup_value);
                byte[] newArray = new byte[k.Length];
                k.CopyTo(newArray, 0);
                newArray[6] = intBytes[0];
                newArray[7] = intBytes[1];
                k = newArray;
                string temppor = bytestostring(k, null);
                filetime = BitConverter.ToInt64(k, 0);
            }
            filetime = BitConverter.ToInt64(k, 0);
            if (filetime > 0)
            {
                if (timezone == "utc") { return DateTime.FromFileTimeUtc(filetime); }
                else if (timezone == "local") { return DateTime.FromFileTime(filetime); }
                else { return null; }
            }
            else { return null; }
        }

        private DateTime ReadTargetLnkTime_noq(int i, BinaryReader j, byte[] k, string timezone)
        {
            // declare vars
            long filetime;
            // do work
            j.BaseStream.Position = i;
            j.Read(k, 0, 8);
            filetime = BitConverter.ToInt64(k, 0);
            if (timezone == "utc") { return DateTime.FromFileTimeUtc(filetime); }
            else { return DateTime.FromFileTime(filetime); } //(timezone == "local")
        }

        public string parseStringData(BinaryReader R, int CurrentPosition, out int NewPosition)
        {
            byte[] temp2 = new byte[2];
            int StringData_temp;
            string varname;
            R.BaseStream.Position = CurrentPosition;
            R.Read(temp2, 0, 2);
            StringData_temp = BitConverter.ToInt16(temp2, 0);
            StringData_temp = StringData_temp * 2; // we are reading Unicode values here
            byte[] StringData_String = new byte[StringData_temp];
            R.BaseStream.Position = CurrentPosition + 2; // lets move to that string
            R.Read(StringData_String, 0, StringData_temp);
            varname = System.Text.Encoding.Unicode.GetString(StringData_String);
            NewPosition = CurrentPosition + 2 + StringData_temp;
            return varname;
        }

        private void folderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder_browser = new FolderBrowserDialog();
            if (folder_browser.ShowDialog() == DialogResult.OK)
            {
                // First we must setup the tables.
                DataTable lnkinfo = new DataTable("lnkinfo");
                parselnk_setuptable(lnkinfo, panel_lnk, "lnk only");

                worker_lnk_vars lnk_vars = new worker_lnk_vars
                {
                    path = folder_browser.SelectedPath,
                    datatable = lnkinfo,
                    datagridview = dataGridView_lnk,
                    panel = panel_lnk
                };
                worker_lnk.RunWorkerAsync(lnk_vars);
            }
        }

        private void worker_lnk_DoWork(object sender, DoWorkEventArgs e)
        {
            // Now we must do the work.
            Stopwatch sw = new Stopwatch();
            sw.Start();
            int filecount;

            worker_lnk_vars vars = e.Argument as worker_lnk_vars;
            if (vars.flags == 1) // we have a drag and drop operation!
            {
                filecount = parselnk_dowork(vars.list_of_files.ToArray(), vars.datatable);
            }
            else //continue with normal operations
            {
                filecount = parselnk_dowork(Directory.GetFiles(vars.path), vars.datatable);
            }
            vars.filecount = filecount;

            sw.Stop();
            statuslabel_lnk_timetaken.Text = String.Concat("Take taken: ", sw.Elapsed.TotalSeconds.ToString(), " seconds.");
            statuslabel_lnk_filesparsed.Text = String.Concat(filecount.ToString(), " .lnk files parsed.");

            e.Result = vars; 
        }

        class worker_lnk_vars
        {
            public List<string> list_of_files { get; set; }
            public string path { get; set; }
            public DataTable datatable { get; set; }
            public DataGridView datagridview { get; set; }
            public Panel panel { get; set; }
            public int flags { get; set; }
            public int filecount { get; set; }
        }

        private void worker_lnk_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            worker_lnk_vars vars = e.Result as worker_lnk_vars;
            parselnk_finalisetable(vars.datatable, vars.datagridview, vars.panel);
            statuslabel_lnk_filebeingparsed.Text = "";
            toUTCToolStripMenuItem_lnk.Visible = true;
            toLocalToolStripMenuItem_lnk.Visible = false;
            dataGridView_lnk.Columns["Embedded Creation Time (Local)"].HeaderText = String.Concat("Embedded Creation Time (Local)");
            dataGridView_lnk.Columns["Embedded Access Time (Local)"].HeaderText = String.Concat("Embedded Access Time (Local)");
            dataGridView_lnk.Columns["Embedded Written Time (Local)"].HeaderText = String.Concat("Embedded Written Time (Local)");
            toLocalToolStripMenuItem_lnk.Visible = false;
            toUTCToolStripMenuItem_lnk.Visible = true;
            tabPage_lnk.Text = string.Concat(".lnk Results - ", vars.filecount);
        }

        private void toUTCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toUTCToolStripMenuItem_lnk.Visible = false;
            toLocalToolStripMenuItem_lnk.Visible = true;
            changetimezone_lnk(dataGridView_lnk, "UTC");
        }

        private void changetimezone_lnk(DataGridView gridview, string timezone)
        {
            foreach (DataGridViewRow row in gridview.Rows)
            {
                DataGridViewCell cell1 = row.Cells["Embedded Creation Time (Local)"];
                {
                    if (cell1.Value != DBNull.Value)
                    {
                        DateTime x = (DateTime)cell1.Value;
                        if (timezone == "UTC")
                        {
                            x = x.ToUniversalTime();
                        }
                        else if (timezone == "Local")
                        {
                            x = x.ToLocalTime();
                        }
                        cell1.Value = x;
                    }
                }
                gridview.Columns["Embedded Creation Time (Local)"].HeaderText = String.Concat("Embedded Creation Time (", timezone, ")");
                DataGridViewCell cell2 = row.Cells["Embedded Written Time (Local)"];
                {
                    if (cell2.Value != DBNull.Value)
                    {
                        DateTime x = (DateTime)cell2.Value;
                        if (timezone == "UTC")
                        {
                            x = x.ToUniversalTime();
                        }
                        if (timezone == "Local")
                        {
                            x = x.ToLocalTime();
                        }
                        cell2.Value = x;
                    }
                }
                gridview.Columns["Embedded Written Time (Local)"].HeaderText = String.Concat("Embedded Written Time (", timezone, ")");
                DataGridViewCell cell3 = row.Cells["Embedded Access Time (Local)"];
                {
                    if (cell3.Value != DBNull.Value)
                    {
                        DateTime x = (DateTime)cell3.Value;
                        if (timezone == "UTC")
                        {
                            x = x.ToUniversalTime();
                        }
                        if (timezone == "Local")
                        {
                            x = x.ToLocalTime();
                        }
                        cell3.Value = x;
                    }
                }
                gridview.Columns["Embedded Access Time (Local)"].HeaderText = String.Concat("Embedded Access Time (", timezone, ")");

                DataGridViewCell cell4 = row.Cells["LNK File Creation Time (Local)"];
                {
                    if (cell3.Value != DBNull.Value)
                    {
                        DateTime x = (DateTime)cell4.Value;
                        if (timezone == "UTC")
                        {
                            x = x.ToUniversalTime();
                        }
                        if (timezone == "Local")
                        {
                            x = x.ToLocalTime();
                        }
                        cell4.Value = x;
                    }
                }
                gridview.Columns["LNK File Creation Time (Local)"].HeaderText = String.Concat("LNK File Creation Time (", timezone, ")");

                DataGridViewCell cell5 = row.Cells["LNK File Written Time (Local)"];
                {
                    if (cell3.Value != DBNull.Value)
                    {
                        DateTime x = (DateTime)cell5.Value;
                        if (timezone == "UTC")
                        {
                            x = x.ToUniversalTime();
                        }
                        if (timezone == "Local")
                        {
                            x = x.ToLocalTime();
                        }
                        cell3.Value = x;
                    }
                }
                gridview.Columns["LNK File Written Time (Local)"].HeaderText = String.Concat("LNK File Written Time (", timezone, ")");

                DataGridViewCell cell6 = row.Cells["LNK File Access Time (Local)"];
                {
                    if (cell3.Value != DBNull.Value)
                    {
                        DateTime x = (DateTime)cell6.Value;
                        if (timezone == "UTC")
                        {
                            x = x.ToUniversalTime();
                        }
                        if (timezone == "Local")
                        {
                            x = x.ToLocalTime();
                        }
                        cell3.Value = x;
                    }
                }
                gridview.Columns["LNK File Access Time (Local)"].HeaderText = String.Concat("LNK File Access Time (", timezone, ")");
            }
            gridview.Refresh();
        }

        private void changetimezone_pf(string timezone)
        {
            foreach (DataGridViewRow row in dataGridview_pfinfo.Rows)
            {
                DataGridViewCell cell1 = row.Cells["Last Run Time (Local)"];
                {
                    if (cell1.Value != DBNull.Value)
                    {
                        DateTime x = (DateTime)cell1.Value;
                        if (timezone == "UTC")
                        {
                            x = x.ToUniversalTime();
                        }
                        if (timezone == "Local")
                        {
                            x = x.ToLocalTime();
                        }
                        cell1.Value = x;
                    }
                }
                dataGridview_pfinfo.Columns["Last Run Time (Local)"].HeaderText = String.Concat("Last Run Time (", timezone, ")");
                DataGridViewCell cell2 = row.Cells["Volume Created Date (Local)"];
                {
                    if (cell2.Value != DBNull.Value)
                    {
                        DateTime x = (DateTime)cell2.Value;
                        if (timezone == "UTC")
                        {
                            x = x.ToUniversalTime();
                        }
                        if (timezone == "Local")
                        {
                            x = x.ToLocalTime();
                        }
                        cell2.Value = x;
                    }
                }
                dataGridview_pfinfo.Columns["Volume Created Date (Local)"].HeaderText = String.Concat("Volume Created Date (", timezone, ")");
            }
            dataGridview_pfinfo.Refresh();
        }

        private void changetimezone_i30(string timezone)
        {
            foreach (DataGridViewRow row in dataGridView_i30.Rows)
            {
                DataGridViewCell cell1 = row.Cells["Creation Time (Local)"];
                {
                    if (cell1.Value != DBNull.Value)
                    {
                        DateTime x = (DateTime)cell1.Value;
                        if (timezone == "UTC")
                        {
                            x = x.ToUniversalTime();
                        }
                        if (timezone == "Local")
                        {
                            x = x.ToLocalTime();
                        }
                        cell1.Value = x;
                    }
                }
                dataGridView_i30.Columns["Creation Time (Local)"].HeaderText = String.Concat("Creation Time (", timezone, ")");
                
                DataGridViewCell cell2 = row.Cells["Last Modified Time (Local)"];
                {
                    if (cell2.Value != DBNull.Value)
                    {
                        DateTime x = (DateTime)cell2.Value;
                        if (timezone == "UTC")
                        {
                            x = x.ToUniversalTime();
                        }
                        if (timezone == "Local")
                        {
                            x = x.ToLocalTime();
                        }
                        cell2.Value = x;
                    }
                }
                dataGridView_i30.Columns["Last Modified Time (Local)"].HeaderText = String.Concat("Last Modified Time (", timezone, ")");

                DataGridViewCell cell3 = row.Cells["Last Access Time (Local)"];
                {
                    if (cell3.Value != DBNull.Value)
                    {
                        DateTime x = (DateTime)cell3.Value;
                        if (timezone == "UTC")
                        {
                            x = x.ToUniversalTime();
                        }
                        if (timezone == "Local")
                        {
                            x = x.ToLocalTime();
                        }
                        cell3.Value = x;
                    }
                }
                dataGridView_i30.Columns["Last Access Time (Local)"].HeaderText = String.Concat("Last Access Time (", timezone, ")");

                DataGridViewCell cell4 = row.Cells["MFT Record Change Time (Local)"];
                {
                    if (cell4.Value != DBNull.Value)
                    {
                        DateTime x = (DateTime)cell4.Value;
                        if (timezone == "UTC")
                        {
                            x = x.ToUniversalTime();
                        }
                        if (timezone == "Local")
                        {
                            x = x.ToLocalTime();
                        }
                        cell4.Value = x;
                    }
                }
                dataGridView_i30.Columns["MFT Record Change Time (Local)"].HeaderText = String.Concat("MFT Record Change Time (", timezone, ")");
            }
            dataGridView_i30.Refresh();
        }

        private void toLocalToolStripMenuItem_lnk_Click(object sender, EventArgs e)
        {
            toUTCToolStripMenuItem_lnk.Visible = true;
            toLocalToolStripMenuItem_lnk.Visible = false;
            changetimezone_lnk(dataGridView_lnk, "Local");
        }

        private void recentFolderLocalMachineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // First we must setup the tables.
            DataTable lnkinfo = new DataTable("lnkinfo");
            parselnk_setuptable(lnkinfo, panel_lnk, "lnk only");

            worker_lnk_vars lnk_vars = new worker_lnk_vars
            {
                path = String.Concat(System.Environment.GetEnvironmentVariable("USERPROFILE"), @"\AppData\Roaming\Microsoft\Windows\Recent"),
                datatable = lnkinfo,
                datagridview = dataGridView_lnk,
                panel = panel_lnk
            };
            worker_lnk.RunWorkerAsync(lnk_vars);
        }

        private void toCSVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            savecsv_format(dataGridView_lnk);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //File.Delete("sfp-log.txt");
            Close();
        }

        private void folderToolStripMenuItem_pfparsefolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder_browser = new FolderBrowserDialog();
            if (folder_browser.ShowDialog() == DialogResult.OK)
            {
                // First we must setup the tables.
                DataTable pfinfo = new DataTable("pfinfo");
                parsepf_setuptable(pfinfo, panel_pf);

                worker_lnk_vars pf_vars = new worker_lnk_vars
                {
                    path = folder_browser.SelectedPath,
                    datatable = pfinfo,
                    datagridview = dataGridview_pfinfo,
                    panel = panel_pf
                };
                worker_pf.RunWorkerAsync(pf_vars);
            }
        }

        private void prefetchFolderLocalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            object temp = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters", "EnablePrefetcher", 0);
            string temp1 = temp.ToString(); // save this setting
            Trace.Write(String.Concat(";" + ";" + "Prefetch setting is " + temp1 + "\n"));
            Trace.Write(String.Concat(";" + ";" + "Parsing of the Local Machine's PF dir selected\n"));
            string SelectedPath;
            SelectedPath = String.Concat(System.Environment.GetEnvironmentVariable("SystemRoot"), @"\Prefetch");
            Trace.Write(String.Concat(";" + ";" + "Local Directory is " + SelectedPath + "\n"));

            // First we must setup the tables.
            DataTable pfinfo = new DataTable("pfinfo");
            parsepf_setuptable(pfinfo, panel_pf);

            worker_lnk_vars pf_vars = new worker_lnk_vars
            {
                path = SelectedPath,
                datatable = pfinfo,
                datagridview = dataGridview_pfinfo,
                panel = panel_pf
            };
            worker_pf.RunWorkerAsync(pf_vars);
        }

        private void worker_pf_DoWork(object sender, DoWorkEventArgs e)
        {
            // Now we must do the work.
            Stopwatch sw = new Stopwatch();
            sw.Start();
            int filecount;
            worker_lnk_vars vars = e.Argument as worker_lnk_vars;

            if (vars.flags == 1) // we have a drag and drop operation!
            {
                filecount = parsepf_dowork(vars.list_of_files.ToArray(), vars.datatable);
            }
            else //continue with normal operations
            {
                filecount = parsepf_dowork(Directory.GetFiles(vars.path), vars.datatable);
            }
            vars.filecount = filecount;
            
            sw.Stop();
            statuslabel_pf_timetaken.Text = String.Concat("Take taken: ", sw.Elapsed.TotalSeconds.ToString(), " seconds.");
            statuslabel_pf_filesparsed.Text = String.Concat(filecount.ToString(), " PF files parsed.");
            e.Result = vars;
        }

        private void worker_pf_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            worker_lnk_vars vars = e.Result as worker_lnk_vars;
            parsepf_finalisetable(vars.datatable, vars.datagridview, vars.panel);
            statuslabel_pf_filebeingparsed.Text = "";
            dataGridview_pfinfo.Columns["Volume Created Date (Local)"].HeaderText = String.Concat("Volume Created Date (Local)");
            dataGridview_pfinfo.Columns["Last Run Time (Local)"].HeaderText = String.Concat("Last Run Time (Local)");
            toLocalToolStripMenuItem_pf.Visible = false;
            toUTCToolStripMenuItem_pf.Visible = true;
            tabPage_pf.Text = string.Concat("Prefetch Results - ", vars.filecount);
        }

        private void toUTCToolStripMenuItem_pf_Click(object sender, EventArgs e)
        {
            toUTCToolStripMenuItem_pf.Visible = false;
            toLocalToolStripMenuItem_pf.Visible = true;
            changetimezone_pf("UTC");
        }

        private void toLocalToolStripMenuItem_pf_Click(object sender, EventArgs e)
        {
            toUTCToolStripMenuItem_pf.Visible = true;
            toLocalToolStripMenuItem_pf.Visible = false;
            changetimezone_pf("Local");
        }

        private void toCSVToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            savecsv_format(dataGridview_pfinfo);
        }

        private void selectedRowsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int selectedRowCount = dataGridview_pfinfo.Rows.GetRowCount(DataGridViewElementStates.Selected);
            string[] selectedrows = new string[dataGridview_pfinfo.Rows.GetRowCount(DataGridViewElementStates.Selected)];
            if (selectedRowCount > 0)
            {
                for (int i = 0; i < selectedRowCount; i++)
                {
                    selectedrows[i] = dataGridview_pfinfo["FullPath", dataGridview_pfinfo.SelectedRows[i].Index].Value.ToString();
                }
            }
            parsepffilefolder_dowork(selectedrows);
            splitContainer_pf.Panel2Collapsed = false;
        }

        private void toCSVToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            SaveFileDialog s = new SaveFileDialog();
            s.Filter = "CSV Files (*.csv)|*.csv";
            if (s.ShowDialog() == DialogResult.OK)
            {
                string buffer_csv_return = "\r\n";
                // Create the output CSV file
                TextWriter T = new StreamWriter(s.FileName);
                DataTable csvoutput = dataGridpfinfosub_fileref.DataSource as DataTable;
                // print file info first

                // print column names first
                foreach (DataColumn column in csvoutput.Columns)
                {
                    string buffer_csv = string.Concat(column.ColumnName, ","); T.Write(buffer_csv);
                }
                T.Write(buffer_csv_return);

                // now print the file references
                foreach (DataRow row in csvoutput.Rows)
                {
                    foreach (DataColumn column in csvoutput.Columns)
                    {
                        string buffer_csv = string.Concat(row[column], ","); T.Write(buffer_csv);
                    }
                    T.Write(buffer_csv_return);
                }

                csvoutput = dataGridpfinfosub_folderref.DataSource as DataTable;
                // now print the folder references
                foreach (DataRow row in csvoutput.Rows)
                {
                    foreach (DataColumn column in csvoutput.Columns)
                    {
                        string buffer_csv = string.Concat(row[column], ","); T.Write(buffer_csv);
                    }
                    T.Write(buffer_csv_return);
                }

                // close the file
                T.Close();
                MessageBox.Show(string.Concat(Path.GetFileName(s.FileName), " saved sucessfully"));
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Simple File Parser v1.6 created by Chris Mayhew\n\nctmayhew@gmail.com\n\nhttps://github.com/ctmayhew/simplefileparser", "About");
        }

        private void oNToolStripMenuItem_Click(object sender, EventArgs e)
        {
            logging = "on";
            oNToolStripMenuItem.Visible = false;
            oFFToolStripMenuItem.Visible = true;
        }

        private void oFFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            logging = "off";
            oNToolStripMenuItem.Visible = true;
            oFFToolStripMenuItem.Visible = false;
        }

        private void centralStandardTimeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toUTCToolStripMenuItem_lnk.Visible = true;
            toLocalToolStripMenuItem_lnk.Visible = true;
            centralStandardTimeToolStripMenuItem_LNK.Visible = false;
            changetimezone_lnk(dataGridView_lnk, "Central Standard Time");
        }

        private void folderOfI30FilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder_browser = new FolderBrowserDialog();
            if (folder_browser.ShowDialog() == DialogResult.OK)
            {
                // First we must setup the tables.
                DataTable i30info = new DataTable("i30info");
                parsei30_setuptable(i30info, paneli30);

                worker_lnk_vars i30_vars = new worker_lnk_vars
                {
                    path = folder_browser.SelectedPath,
                    datatable = i30info,
                    datagridview = dataGridView_i30,
                    panel = paneli30
                };
                worker_i30.RunWorkerAsync(i30_vars);
            }
        }

        private void worker_i30_DoWork(object sender, DoWorkEventArgs e)
        {
            // Now we must do the work.
            Stopwatch sw = new Stopwatch();
            int filecount;
            sw.Start();

            worker_lnk_vars vars = e.Argument as worker_lnk_vars;

            if (vars.flags == 1) // we have a drag and drop operation!
            {
                filecount = parsei30_dowork(vars.list_of_files.ToArray(), vars.datatable);
            }
            else //continue with normal operations
            {
                filecount = parsei30_dowork(Directory.GetFiles(vars.path), vars.datatable);
            }
            vars.filecount = filecount;

            sw.Stop();
            statuslabel_i30_timetaken.Text = String.Concat("Take taken: ", sw.Elapsed.TotalSeconds.ToString(), " seconds.");
            statuslabel_i30_filesparsed.Text = String.Concat(filecount.ToString(), " $I30 files parsed.");

            e.Result = vars; 
        }

        private void worker_i30_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            worker_lnk_vars vars = e.Result as worker_lnk_vars;
            parsei30_finalisetable(vars.datatable, vars.datagridview, vars.panel);
            statuslabel_i30_filebeingparsed.Text = "";
            toUTCToolStripMenuItem_i30.Visible = true;
            paneli30.Visible = true;
            tabpage_i30.Text = string.Concat("$I30 Results - ", vars.filecount);
        }

        private void toCSVToolStripMenuItem_i30_Click(object sender, EventArgs e)
        {
            savecsv_format(dataGridView_i30);
        }

        private void toLocalToolStripMenuItem_i30_Click(object sender, EventArgs e)
        {
            toUTCToolStripMenuItem_i30.Visible = true;
            toLocalToolStripMenuItem_i30.Visible = false;
            changetimezone_i30("Local");
        }

        private void toUTCToolStripMenuItem_i30_Click(object sender, EventArgs e)
        {
            toUTCToolStripMenuItem_i30.Visible = false;
            toLocalToolStripMenuItem_i30.Visible = true;
            changetimezone_i30("UTC");
        }

        private void SFP_main_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void SFP_main_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);
                List<string> all_files = new List<string>();

                //create the lists we need for parsing each file
                List<string> lnk_files = new List<string>();
                List<string> pf_files = new List<string>();
                List<string> jl_files = new List<string>();
                List<string> i30_files = new List<string>();

                foreach (string fileName in files)
                {
                    //check to see if it a directory - we need to enumarate these - ONE level only.
                    if (Directory.Exists(fileName))
                    {
                        
                        //Add files from folder
                        foreach (string s in Directory.GetFiles(fileName))
                        {
                            all_files.Add(s);
                            switch (Path.GetExtension(s))
                            {
                                case ".lnk":
                                    lnk_files.Add(s); break;
                                case ".pf":
                                    pf_files.Add(s); break;
                                default: break;
                            }
                        }
                    }
                    else
                    {
                        string s = fileName;
                        all_files.Add(s);
                        switch (Path.GetExtension(s))
                        {
                            case ".lnk":
                                lnk_files.Add(s); break;
                            case ".pf":
                                pf_files.Add(s); break;
                            default: break;
                        }

                    }
                }

                //now lets go do the work!
                if (lnk_files.Any()) //we need to check if the array contains any data
                {
                    // First we must setup the tables.
                    DataTable lnkinfo = new DataTable("lnkinfo");
                    parselnk_setuptable(lnkinfo, panel_lnk, "lnk only");

                    worker_lnk_vars lnk_vars = new worker_lnk_vars
                    {
                        list_of_files = lnk_files,
                        datatable = lnkinfo,
                        datagridview = dataGridView_lnk,
                        panel = panel_lnk,
                        flags = 1
                    };
                    worker_lnk.RunWorkerAsync(lnk_vars);
                }

                if (pf_files.Any())
                {
                    // First we must setup the tables.
                    DataTable pfinfo = new DataTable("pfinfo");
                    parsepf_setuptable(pfinfo, panel_pf);

                    worker_lnk_vars pf_vars = new worker_lnk_vars
                    {
                        list_of_files = pf_files,
                        datatable = pfinfo,
                        datagridview = dataGridview_pfinfo,
                        flags = 1
                    };
                    worker_pf.RunWorkerAsync(pf_vars);
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine("Error in DragDrop function: " + ex.Message);

                // don't show MessageBox here - Explorer is waiting !
            }
        }
    }
}
