using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace kedouwenc
{
    public partial class Gif_Help : Form
    {
        public Gif_Help()
        {
            InitializeComponent();      
            switch (Ribbon1.gifid)
            {
                case "ReadMode":
                    this.pictureBox1.Image = Properties.Resources.ReadMode;
                    break;
                case "ForDisplay":
                    this.pictureBox1.Image = Properties.Resources.ForDisplay;                   
                    break;
                case "FromHere":
                    this.pictureBox1.Image = Properties.Resources.FromHere;
                    break;
                case "ShowAndHide":
                    this.pictureBox1.Image = Properties.Resources.ShowAndHide;
                    break;
                case "Help":
                    this.pictureBox1.Image = Properties.Resources.Help;
                    break;
                default:
                    //  Me.PictureBox1.Image = My.Resources.Help
                    break;
            };


        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }      
    }
}
