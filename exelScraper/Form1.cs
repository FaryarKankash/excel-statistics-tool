using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using Microsoft.Office.Interop.Excel;
using Bunifu.UI.WinForms;
using ExcelDataReader;
using System.Runtime.InteropServices.ComTypes;
using DataTable = System.Data.DataTable;
using OfficeOpenXml.Drawing.Chart.ChartEx;

namespace exelScraper
{


    public partial class Form1 : Form
    {

        readonly Events eventHandlerList = new Events();
        
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonGetFile_Click(object sender, EventArgs e)
        {
            eventHandlerList.readExcel(dataGridView);
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                columnDropDown.Items.Add(column.HeaderText);
                columnDropDown2.Items.Add(column.HeaderText);
            }
        }

        private void avrageButton_Click(object sender, EventArgs e)
        {
            eventHandlerList.getAvrage(dataGridView, columnDropDown);
        }

        private void rangeButton_Click(object sender, EventArgs e)
        {

            eventHandlerList.getRange(dataGridView, columnDropDown);
        }

        private void quantileButton_Click(object sender, EventArgs e)
        {
            eventHandlerList.getQuantile(dataGridView, columnDropDown);
        }

        private void varianceBtn_Click(object sender, EventArgs e)
        {
            eventHandlerList.variance(dataGridView, columnDropDown);
        }

        private void standardDeviationBtn_Click(object sender, EventArgs e)
        {
            eventHandlerList.standardDeviation(dataGridView, columnDropDown);
        }

        private void covButton_Click(object sender, EventArgs e)
        {
            eventHandlerList.covariance(dataGridView, columnDropDown, columnDropDown2);
        }

        private void correlationCoefficientBtn_Click(object sender, EventArgs e)
        {
            eventHandlerList.correlationCoefficient(dataGridView, columnDropDown, columnDropDown2);
        }
    }


}
