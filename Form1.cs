using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace YAMLConv
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public void SetText(string s)
        {
            textBox1.Text = s;
        }

        public delegate void Event(object sender, EventArgs e);
        public Event TsvCommentCheckBox_CheckedChanged { get; set; }

        private void button1_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(textBox1.Text);
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            TsvCommentCheckBox_CheckedChanged(sender, e);
        }
    }
}
