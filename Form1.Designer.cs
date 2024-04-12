namespace Simu_Gen
{
    partial class Simu_Gen
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Simu_Gen));
            this.File_Entry = new System.Windows.Forms.TextBox();
            this.Search = new System.Windows.Forms.Button();
            this.MTK = new System.Windows.Forms.CheckBox();
            this.Output = new System.Windows.Forms.Button();
            this.Output_Folder = new System.Windows.Forms.TextBox();
            this.Browse = new System.Windows.Forms.Button();
            this.Sector_Entry = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.Track_Circuit = new System.Windows.Forms.CheckBox();
            this.Traffic_Direction = new System.Windows.Forms.CheckBox();
            this.MBL = new System.Windows.Forms.CheckBox();
            this.Subroute = new System.Windows.Forms.CheckBox();
            this.Route = new System.Windows.Forms.CheckBox();
            this.Point = new System.Windows.Forms.CheckBox();
            this.Three_Aspect_Signal = new System.Windows.Forms.CheckBox();
            this.Cycle = new System.Windows.Forms.CheckBox();
            this.Shunt_Signal = new System.Windows.Forms.CheckBox();
            this.Overlap = new System.Windows.Forms.CheckBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.Select_All = new System.Windows.Forms.Button();
            this.Deselect_All = new System.Windows.Forms.Button();
            this.MTK_Entry = new System.Windows.Forms.TextBox();
            this.Search_MTK = new System.Windows.Forms.Button();
            this.timelabel = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.flowLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // File_Entry
            // 
            this.File_Entry.Location = new System.Drawing.Point(12, 12);
            this.File_Entry.Multiline = true;
            this.File_Entry.Name = "File_Entry";
            this.File_Entry.Size = new System.Drawing.Size(527, 31);
            this.File_Entry.TabIndex = 0;
            this.File_Entry.Text = "Enter Control Table Location\r\n";
            this.File_Entry.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.File_Entry.TextChanged += new System.EventHandler(this.File_Entry_TextChanged);
            // 
            // Search
            // 
            this.Search.Location = new System.Drawing.Point(560, 12);
            this.Search.Name = "Search";
            this.Search.Size = new System.Drawing.Size(75, 31);
            this.Search.TabIndex = 1;
            this.Search.Text = "Search_CT";
            this.Search.UseVisualStyleBackColor = true;
            this.Search.Click += new System.EventHandler(this.Search_Click);
            // 
            // MTK
            // 
            this.MTK.AutoSize = true;
            this.MTK.Location = new System.Drawing.Point(13, 23);
            this.MTK.Name = "MTK";
            this.MTK.Padding = new System.Windows.Forms.Padding(0, 0, 10, 10);
            this.MTK.Size = new System.Drawing.Size(59, 27);
            this.MTK.TabIndex = 2;
            this.MTK.Text = "MTK";
            this.MTK.UseVisualStyleBackColor = true;
            this.MTK.CheckedChanged += new System.EventHandler(this.MTK_CheckedChanged);
            // 
            // Output
            // 
            this.Output.Location = new System.Drawing.Point(439, 407);
            this.Output.Name = "Output";
            this.Output.Size = new System.Drawing.Size(75, 31);
            this.Output.TabIndex = 4;
            this.Output.Text = "Output";
            this.Output.UseVisualStyleBackColor = true;
            this.Output.Click += new System.EventHandler(this.Output_Click);
            // 
            // Output_Folder
            // 
            this.Output_Folder.Location = new System.Drawing.Point(12, 124);
            this.Output_Folder.Multiline = true;
            this.Output_Folder.Name = "Output_Folder";
            this.Output_Folder.Size = new System.Drawing.Size(527, 31);
            this.Output_Folder.TabIndex = 5;
            this.Output_Folder.Text = "Enter Output Folder Name";
            this.Output_Folder.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.Output_Folder.TextChanged += new System.EventHandler(this.Output_Folder_TextChanged);
            // 
            // Browse
            // 
            this.Browse.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Browse.Location = new System.Drawing.Point(560, 124);
            this.Browse.Name = "Browse";
            this.Browse.Size = new System.Drawing.Size(75, 31);
            this.Browse.TabIndex = 6;
            this.Browse.Text = "Browse";
            this.Browse.UseVisualStyleBackColor = true;
            this.Browse.Click += new System.EventHandler(this.Browse_Click);
            // 
            // Sector_Entry
            // 
            this.Sector_Entry.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.Sector_Entry.Location = new System.Drawing.Point(423, 250);
            this.Sector_Entry.Name = "Sector_Entry";
            this.Sector_Entry.Size = new System.Drawing.Size(100, 20);
            this.Sector_Entry.TabIndex = 7;
            this.Sector_Entry.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.Sector_Entry.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(420, 218);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Enter Sector Code";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.flowLayoutPanel1.Controls.Add(this.MTK);
            this.flowLayoutPanel1.Controls.Add(this.Track_Circuit);
            this.flowLayoutPanel1.Controls.Add(this.Traffic_Direction);
            this.flowLayoutPanel1.Controls.Add(this.MBL);
            this.flowLayoutPanel1.Controls.Add(this.Subroute);
            this.flowLayoutPanel1.Controls.Add(this.Route);
            this.flowLayoutPanel1.Controls.Add(this.Point);
            this.flowLayoutPanel1.Controls.Add(this.Three_Aspect_Signal);
            this.flowLayoutPanel1.Controls.Add(this.Cycle);
            this.flowLayoutPanel1.Controls.Add(this.Shunt_Signal);
            this.flowLayoutPanel1.Controls.Add(this.Overlap);
            this.flowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(12, 218);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Padding = new System.Windows.Forms.Padding(10, 20, 20, 20);
            this.flowLayoutPanel1.Size = new System.Drawing.Size(368, 177);
            this.flowLayoutPanel1.TabIndex = 10;
            this.flowLayoutPanel1.Paint += new System.Windows.Forms.PaintEventHandler(this.flowLayoutPanel1_Paint_1);
            // 
            // Track_Circuit
            // 
            this.Track_Circuit.AutoSize = true;
            this.Track_Circuit.Location = new System.Drawing.Point(13, 56);
            this.Track_Circuit.Name = "Track_Circuit";
            this.Track_Circuit.Padding = new System.Windows.Forms.Padding(0, 0, 10, 10);
            this.Track_Circuit.Size = new System.Drawing.Size(96, 27);
            this.Track_Circuit.TabIndex = 4;
            this.Track_Circuit.Text = "Track Circuit";
            this.Track_Circuit.UseVisualStyleBackColor = true;
            this.Track_Circuit.CheckedChanged += new System.EventHandler(this.Track_Circuit_CheckedChanged);
            // 
            // Traffic_Direction
            // 
            this.Traffic_Direction.AutoSize = true;
            this.Traffic_Direction.Location = new System.Drawing.Point(13, 89);
            this.Traffic_Direction.Name = "Traffic_Direction";
            this.Traffic_Direction.Padding = new System.Windows.Forms.Padding(0, 0, 10, 10);
            this.Traffic_Direction.Size = new System.Drawing.Size(111, 27);
            this.Traffic_Direction.TabIndex = 5;
            this.Traffic_Direction.Text = "Traffic Direction";
            this.Traffic_Direction.UseVisualStyleBackColor = true;
            this.Traffic_Direction.CheckedChanged += new System.EventHandler(this.Traffic_Direction_CheckedChanged);
            // 
            // MBL
            // 
            this.MBL.AutoSize = true;
            this.MBL.Location = new System.Drawing.Point(13, 122);
            this.MBL.Name = "MBL";
            this.MBL.Padding = new System.Windows.Forms.Padding(0, 0, 10, 10);
            this.MBL.Size = new System.Drawing.Size(58, 27);
            this.MBL.TabIndex = 6;
            this.MBL.Text = "MBL";
            this.MBL.UseVisualStyleBackColor = true;
            this.MBL.CheckedChanged += new System.EventHandler(this.MBL_CheckedChanged);
            // 
            // Subroute
            // 
            this.Subroute.AutoSize = true;
            this.Subroute.Location = new System.Drawing.Point(130, 23);
            this.Subroute.Name = "Subroute";
            this.Subroute.Padding = new System.Windows.Forms.Padding(0, 0, 10, 10);
            this.Subroute.Size = new System.Drawing.Size(79, 27);
            this.Subroute.TabIndex = 7;
            this.Subroute.Text = "Subroute";
            this.Subroute.UseVisualStyleBackColor = true;
            this.Subroute.CheckedChanged += new System.EventHandler(this.Subroute_CheckedChanged);
            // 
            // Route
            // 
            this.Route.AutoSize = true;
            this.Route.Location = new System.Drawing.Point(130, 56);
            this.Route.Name = "Route";
            this.Route.Padding = new System.Windows.Forms.Padding(0, 0, 10, 10);
            this.Route.Size = new System.Drawing.Size(65, 27);
            this.Route.TabIndex = 8;
            this.Route.Text = "Route";
            this.Route.UseVisualStyleBackColor = true;
            this.Route.CheckedChanged += new System.EventHandler(this.Route_CheckedChanged);
            // 
            // Point
            // 
            this.Point.AutoSize = true;
            this.Point.Location = new System.Drawing.Point(130, 89);
            this.Point.Name = "Point";
            this.Point.Padding = new System.Windows.Forms.Padding(0, 0, 10, 10);
            this.Point.Size = new System.Drawing.Size(60, 27);
            this.Point.TabIndex = 9;
            this.Point.Text = "Point";
            this.Point.UseVisualStyleBackColor = true;
            this.Point.CheckedChanged += new System.EventHandler(this.Point_CheckedChanged);
            // 
            // Three_Aspect_Signal
            // 
            this.Three_Aspect_Signal.AutoSize = true;
            this.Three_Aspect_Signal.Location = new System.Drawing.Point(130, 122);
            this.Three_Aspect_Signal.Name = "Three_Aspect_Signal";
            this.Three_Aspect_Signal.Padding = new System.Windows.Forms.Padding(0, 0, 10, 10);
            this.Three_Aspect_Signal.Size = new System.Drawing.Size(132, 27);
            this.Three_Aspect_Signal.TabIndex = 10;
            this.Three_Aspect_Signal.Text = "Three Aspect Signal";
            this.Three_Aspect_Signal.UseVisualStyleBackColor = true;
            this.Three_Aspect_Signal.CheckedChanged += new System.EventHandler(this.Three_Aspect_Signal_CheckedChanged);
            // 
            // Cycle
            // 
            this.Cycle.AutoSize = true;
            this.Cycle.Location = new System.Drawing.Point(268, 23);
            this.Cycle.Name = "Cycle";
            this.Cycle.Padding = new System.Windows.Forms.Padding(0, 0, 10, 10);
            this.Cycle.Size = new System.Drawing.Size(62, 27);
            this.Cycle.TabIndex = 11;
            this.Cycle.Text = "Cycle";
            this.Cycle.UseVisualStyleBackColor = true;
            this.Cycle.CheckedChanged += new System.EventHandler(this.Cycle_CheckedChanged);
            // 
            // Shunt_Signal
            // 
            this.Shunt_Signal.AutoSize = true;
            this.Shunt_Signal.Location = new System.Drawing.Point(268, 56);
            this.Shunt_Signal.Name = "Shunt_Signal";
            this.Shunt_Signal.Padding = new System.Windows.Forms.Padding(0, 0, 10, 10);
            this.Shunt_Signal.Size = new System.Drawing.Size(96, 27);
            this.Shunt_Signal.TabIndex = 12;
            this.Shunt_Signal.Text = "Shunt Signal";
            this.Shunt_Signal.UseVisualStyleBackColor = true;
            this.Shunt_Signal.CheckedChanged += new System.EventHandler(this.Shunt_Signal_CheckedChanged);
            // 
            // Overlap
            // 
            this.Overlap.AutoSize = true;
            this.Overlap.Location = new System.Drawing.Point(268, 89);
            this.Overlap.Name = "Overlap";
            this.Overlap.Padding = new System.Windows.Forms.Padding(0, 0, 10, 10);
            this.Overlap.Size = new System.Drawing.Size(73, 27);
            this.Overlap.TabIndex = 13;
            this.Overlap.Text = "Overlap";
            this.Overlap.UseVisualStyleBackColor = true;
            this.Overlap.CheckedChanged += new System.EventHandler(this.Overlap_CheckedChanged);
            // 
            // richTextBox1
            // 
            this.richTextBox1.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.richTextBox1.Location = new System.Drawing.Point(560, 177);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(194, 245);
            this.richTextBox1.TabIndex = 11;
            this.richTextBox1.Text = "";
            // 
            // Select_All
            // 
            this.Select_All.Location = new System.Drawing.Point(27, 407);
            this.Select_All.Name = "Select_All";
            this.Select_All.Size = new System.Drawing.Size(75, 31);
            this.Select_All.TabIndex = 12;
            this.Select_All.Text = "Select All";
            this.Select_All.UseVisualStyleBackColor = true;
            this.Select_All.Click += new System.EventHandler(this.Select_All_Button);
            // 
            // Deselect_All
            // 
            this.Deselect_All.Location = new System.Drawing.Point(144, 407);
            this.Deselect_All.Name = "Deselect_All";
            this.Deselect_All.Size = new System.Drawing.Size(75, 31);
            this.Deselect_All.TabIndex = 13;
            this.Deselect_All.Text = "Deselect All";
            this.Deselect_All.UseVisualStyleBackColor = true;
            this.Deselect_All.Click += new System.EventHandler(this.Deselect_All_Button);
            // 
            // MTK_Entry
            // 
            this.MTK_Entry.Location = new System.Drawing.Point(12, 69);
            this.MTK_Entry.Multiline = true;
            this.MTK_Entry.Name = "MTK_Entry";
            this.MTK_Entry.Size = new System.Drawing.Size(527, 31);
            this.MTK_Entry.TabIndex = 14;
            this.MTK_Entry.Text = "Enter MTK File Location\r\n";
            this.MTK_Entry.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // Search_MTK
            // 
            this.Search_MTK.Location = new System.Drawing.Point(560, 69);
            this.Search_MTK.Name = "Search_MTK";
            this.Search_MTK.Size = new System.Drawing.Size(75, 31);
            this.Search_MTK.TabIndex = 15;
            this.Search_MTK.Text = "Search MTK";
            this.Search_MTK.UseVisualStyleBackColor = true;
            this.Search_MTK.Click += new System.EventHandler(this.Search_MTK_Button);
            // 
            // timelabel
            // 
            this.timelabel.AutoSize = true;
            this.timelabel.Location = new System.Drawing.Point(14, 171);
            this.timelabel.Name = "timelabel";
            this.timelabel.Size = new System.Drawing.Size(77, 13);
            this.timelabel.TabIndex = 16;
            this.timelabel.Text = "Elapsed Time: ";
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.notifyIcon1.BalloonTipText = "Completed";
            this.notifyIcon1.BalloonTipTitle = "Finished";
            this.notifyIcon1.Text = "notifyIcon1";
            this.notifyIcon1.Visible = true;
            // 
            // Simu_Gen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.timelabel);
            this.Controls.Add(this.Search_MTK);
            this.Controls.Add(this.MTK_Entry);
            this.Controls.Add(this.Deselect_All);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.flowLayoutPanel1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Sector_Entry);
            this.Controls.Add(this.Browse);
            this.Controls.Add(this.Output_Folder);
            this.Controls.Add(this.Output);
            this.Controls.Add(this.Search);
            this.Controls.Add(this.File_Entry);
            this.Controls.Add(this.Select_All);
            this.DoubleBuffered = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Simu_Gen";
            this.Text = "Simu_Gen";
            this.TransparencyKey = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.Load += new System.EventHandler(this.Form1_Load);
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox File_Entry;
        private System.Windows.Forms.Button Search;
        private System.Windows.Forms.CheckBox MTK;
        private System.Windows.Forms.Button Output;
        private System.Windows.Forms.TextBox Output_Folder;
        private System.Windows.Forms.Button Browse;
        private System.Windows.Forms.TextBox Sector_Entry;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.CheckBox Track_Circuit;
        private System.Windows.Forms.CheckBox Traffic_Direction;
        private System.Windows.Forms.CheckBox MBL;
        private System.Windows.Forms.CheckBox Subroute;
        private System.Windows.Forms.CheckBox Point;
        private System.Windows.Forms.CheckBox Three_Aspect_Signal;
        private System.Windows.Forms.CheckBox Cycle;
        private System.Windows.Forms.CheckBox Shunt_Signal;
        private System.Windows.Forms.CheckBox Overlap;
        private System.Windows.Forms.CheckBox Route;
        private System.Windows.Forms.Button Select_All;
        private System.Windows.Forms.Button Deselect_All;
        private System.Windows.Forms.TextBox MTK_Entry;
        private System.Windows.Forms.Button Search_MTK;
        private System.Windows.Forms.Label timelabel;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
    }
}

