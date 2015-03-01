namespace Exportar
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btt_Products = new System.Windows.Forms.Button();
            this.lbl_show = new System.Windows.Forms.Label();
            this.btt_Category = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.nupLinesImport = new System.Windows.Forms.NumericUpDown();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.nupLinesImport)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btt_Products
            // 
            this.btt_Products.Location = new System.Drawing.Point(57, 30);
            this.btt_Products.Name = "btt_Products";
            this.btt_Products.Size = new System.Drawing.Size(80, 42);
            this.btt_Products.TabIndex = 0;
            this.btt_Products.Text = "Products";
            this.btt_Products.UseVisualStyleBackColor = true;
            this.btt_Products.Click += new System.EventHandler(this.button1_Click);
            // 
            // lbl_show
            // 
            this.lbl_show.AutoSize = true;
            this.lbl_show.Location = new System.Drawing.Point(229, 173);
            this.lbl_show.Name = "lbl_show";
            this.lbl_show.Size = new System.Drawing.Size(70, 13);
            this.lbl_show.TabIndex = 1;
            this.lbl_show.Text = "Final Column:";
            this.lbl_show.Visible = false;
            // 
            // btt_Category
            // 
            this.btt_Category.Location = new System.Drawing.Point(232, 30);
            this.btt_Category.Name = "btt_Category";
            this.btt_Category.Size = new System.Drawing.Size(92, 42);
            this.btt_Category.TabIndex = 2;
            this.btt_Category.Text = "Category";
            this.btt_Category.UseVisualStyleBackColor = true;
            this.btt_Category.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(224, 204);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(100, 20);
            this.textBox1.TabIndex = 3;
            this.textBox1.Text = "HM";
            this.textBox1.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(54, 173);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(108, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Nr Of Lines To Import";
            // 
            // nupLinesImport
            // 
            this.nupLinesImport.Location = new System.Drawing.Point(57, 204);
            this.nupLinesImport.Name = "nupLinesImport";
            this.nupLinesImport.Size = new System.Drawing.Size(80, 20);
            this.nupLinesImport.TabIndex = 5;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(140, 86);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(85, 37);
            this.button1.TabIndex = 6;
            this.button1.Text = "Relate";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_2);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(381, 262);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.nupLinesImport);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btt_Category);
            this.Controls.Add(this.lbl_show);
            this.Controls.Add(this.btt_Products);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Commerciol Automatics - Importer";
            ((System.ComponentModel.ISupportInitialize)(this.nupLinesImport)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btt_Products;
        private System.Windows.Forms.Label lbl_show;
        private System.Windows.Forms.Button btt_Category;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown nupLinesImport;
        private System.Windows.Forms.Button button1;
    }
}

