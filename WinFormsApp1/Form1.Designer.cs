namespace WinFormsApp1
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            button2 = new Button();
            button3 = new Button();
            button4 = new Button();
            SuspendLayout();
            // 
            // button2
            // 
            button2.Location = new Point(519, 357);
            button2.Margin = new Padding(4, 5, 4, 5);
            button2.Name = "button2";
            button2.Size = new Size(107, 38);
            button2.TabIndex = 1;
            button2.Text = "excel";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            //// button3
            //// 
            //button3.Location = new Point(520, 440);
            //button3.Margin = new Padding(4, 5, 4, 5);
            //button3.Name = "button3";
            //button3.Size = new Size(107, 38);
            //button3.TabIndex = 2;
            //button3.Text = "format csv";
            //button3.UseVisualStyleBackColor = true;
            //button3.Click += button3_Click;
            //// 
            // button4
            // 
            //button4.Location = new Point(519, 522);
            //button4.Margin = new Padding(4, 5, 4, 5);
            //button4.Name = "button4";
            //button4.Size = new Size(107, 33);
            //button4.TabIndex = 3;
            //button4.Text = "format pdf";
            //button4.UseVisualStyleBackColor = true;
            //button4.Click += button4_Click;
            //// 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1143, 750);
            Controls.Add(button4);
            Controls.Add(button3);
            Controls.Add(button2);
            Margin = new Padding(4, 5, 4, 5);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
        }

        #endregion
        private Button button2;
        private Button button3;
        private Button button4;
    }
}
