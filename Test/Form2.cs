using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Test
{
	public partial class Form2 : Form
	{
		public Form2() {
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e) {
			decimal.TryParse(textBox1.Text, out decimal num);

			textBox2.Text = XlFunctions.GeoMoney(num,"{l} ლ. {t} თ.");
		}
	}
}
