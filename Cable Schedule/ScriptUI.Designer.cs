namespace Script
{
    partial class ScriptUI
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ScriptUI));
            this.ni_script = new System.Windows.Forms.NotifyIcon(this.components);
            this.cms_script = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.Settings = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.Exit = new System.Windows.Forms.ToolStripMenuItem();
            this.btn_Do = new System.Windows.Forms.Button();
            this.ms_script = new System.Windows.Forms.MenuStrip();
            this.настройкиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lb_Status = new System.Windows.Forms.Label();
            this.pb_Progress = new System.Windows.Forms.ProgressBar();
            this.btn_script_exit = new System.Windows.Forms.Button();
            this.lbl_sheet = new System.Windows.Forms.Label();
            this.cms_script.SuspendLayout();
            this.ms_script.SuspendLayout();
            this.SuspendLayout();
            // 
            // ni_script
            // 
            this.ni_script.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.ni_script.ContextMenuStrip = this.cms_script;
            this.ni_script.Icon = ((System.Drawing.Icon)(resources.GetObject("ni_script.Icon")));
            this.ni_script.Text = "Кабельный журнал";
            this.ni_script.Visible = true;
            this.ni_script.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.ni_script_MouseDoubleClick);
            // 
            // cms_script
            // 
            this.cms_script.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.Settings,
            this.toolStripSeparator1,
            this.Exit});
            this.cms_script.Name = "cms_script";
            this.cms_script.Size = new System.Drawing.Size(135, 54);
            // 
            // Settings
            // 
            this.Settings.Name = "Settings";
            this.Settings.Size = new System.Drawing.Size(134, 22);
            this.Settings.Text = "Настройки";
            this.Settings.Click += new System.EventHandler(this.SettingsToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(131, 6);
            // 
            // Exit
            // 
            this.Exit.Name = "Exit";
            this.Exit.Size = new System.Drawing.Size(134, 22);
            this.Exit.Text = "Выход";
            this.Exit.Click += new System.EventHandler(this.Exit_Click);
            // 
            // btn_Do
            // 
            this.btn_Do.Location = new System.Drawing.Point(12, 108);
            this.btn_Do.Name = "btn_Do";
            this.btn_Do.Size = new System.Drawing.Size(79, 24);
            this.btn_Do.TabIndex = 0;
            this.btn_Do.Text = "Расчитать";
            this.btn_Do.UseVisualStyleBackColor = true;
            this.btn_Do.Click += new System.EventHandler(this.btn_Do_Click);
            // 
            // ms_script
            // 
            this.ms_script.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.настройкиToolStripMenuItem});
            this.ms_script.Location = new System.Drawing.Point(0, 0);
            this.ms_script.Name = "ms_script";
            this.ms_script.Size = new System.Drawing.Size(213, 24);
            this.ms_script.TabIndex = 1;
            this.ms_script.Text = "menuStrip1";
            // 
            // настройкиToolStripMenuItem
            // 
            this.настройкиToolStripMenuItem.Name = "настройкиToolStripMenuItem";
            this.настройкиToolStripMenuItem.Size = new System.Drawing.Size(79, 20);
            this.настройкиToolStripMenuItem.Text = "Настройки";
            this.настройкиToolStripMenuItem.Click += new System.EventHandler(this.SettingsToolStripMenuItem_Click);
            // 
            // lb_Status
            // 
            this.lb_Status.AutoSize = true;
            this.lb_Status.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.lb_Status.Location = new System.Drawing.Point(12, 43);
            this.lb_Status.MaximumSize = new System.Drawing.Size(250, 13);
            this.lb_Status.Name = "lb_Status";
            this.lb_Status.Size = new System.Drawing.Size(184, 13);
            this.lb_Status.TabIndex = 2;
            this.lb_Status.Text = "Ожидание запуска                           ";
            // 
            // pb_Progress
            // 
            this.pb_Progress.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.pb_Progress.Location = new System.Drawing.Point(7, 59);
            this.pb_Progress.Name = "pb_Progress";
            this.pb_Progress.Size = new System.Drawing.Size(189, 23);
            this.pb_Progress.Step = 1;
            this.pb_Progress.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.pb_Progress.TabIndex = 3;
            // 
            // btn_script_exit
            // 
            this.btn_script_exit.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btn_script_exit.Location = new System.Drawing.Point(108, 108);
            this.btn_script_exit.Name = "btn_script_exit";
            this.btn_script_exit.Size = new System.Drawing.Size(87, 24);
            this.btn_script_exit.TabIndex = 5;
            this.btn_script_exit.Text = "Выйти";
            this.btn_script_exit.UseVisualStyleBackColor = true;
            this.btn_script_exit.Click += new System.EventHandler(this.btn_script_exit_Click);
            // 
            // lbl_sheet
            // 
            this.lbl_sheet.AutoSize = true;
            this.lbl_sheet.Location = new System.Drawing.Point(12, 85);
            this.lbl_sheet.Name = "lbl_sheet";
            this.lbl_sheet.Size = new System.Drawing.Size(0, 13);
            this.lbl_sheet.TabIndex = 6;
            // 
            // ScriptUI
            // 
            this.AcceptButton = this.btn_Do;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btn_script_exit;
            this.ClientSize = new System.Drawing.Size(213, 144);
            this.ContextMenuStrip = this.cms_script;
            this.Controls.Add(this.lbl_sheet);
            this.Controls.Add(this.btn_script_exit);
            this.Controls.Add(this.lb_Status);
            this.Controls.Add(this.btn_Do);
            this.Controls.Add(this.ms_script);
            this.Controls.Add(this.pb_Progress);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.ms_script;
            this.MaximizeBox = false;
            this.Name = "ScriptUI";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Кабельный журнал";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ScriptUI_FormClosing);
            this.cms_script.ResumeLayout(false);
            this.ms_script.ResumeLayout(false);
            this.ms_script.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Do;
        private System.Windows.Forms.MenuStrip ms_script;
        private System.Windows.Forms.ToolStripMenuItem настройкиToolStripMenuItem;
        private System.Windows.Forms.Label lb_Status;
        private System.Windows.Forms.ProgressBar pb_Progress;
        private System.Windows.Forms.ContextMenuStrip cms_script;
        private System.Windows.Forms.ToolStripMenuItem Settings;
        private System.Windows.Forms.ToolStripMenuItem Exit;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.Button btn_script_exit;
        private System.Windows.Forms.NotifyIcon ni_script;
        private System.Windows.Forms.Label lbl_sheet;
    }
}

