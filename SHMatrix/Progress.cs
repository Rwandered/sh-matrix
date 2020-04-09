using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SHMatrix
{
    public partial class Progress : Form
    {
        Bitmap animatedImage;
        bool currentlyAnimating = false;
        public Progress()
        {
            InitializeComponent();
            animatedImage = new Bitmap("prog.gif");
            this.Location = new Point(DataR.left, DataR.top);
        }


        public void AnimateImage()
        {
            if (!currentlyAnimating)
            {
                //Begin the animation only once. 
                ImageAnimator.Animate(animatedImage, new EventHandler(this.OnFrameChanged));
                currentlyAnimating = true;
            }
        }
        private void OnFrameChanged(object o, EventArgs e)
        {
            //Force a call to the Paint event handler. 
            this.Invalidate();
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            //Begin the animation. 
            AnimateImage();
            //Get the next frame ready for rendering. 
            ImageAnimator.UpdateFrames();
            //Draw the next frame in the animation. 
            e.Graphics.DrawImage(this.animatedImage, new Point(20, 20));
        }

        private void buttonStopPlay_Click(object sender, EventArgs e)
        {
            if (currentlyAnimating)
            {
                ImageAnimator.StopAnimate(animatedImage, new EventHandler(this.OnFrameChanged));
                currentlyAnimating = false;
            }
            else
            {
                AnimateImage();
            }
        }

        private void Progress_Load(object sender, EventArgs e)
        {
            this.Location = new Point(DataR.left, DataR.top);
        }
    }
}
