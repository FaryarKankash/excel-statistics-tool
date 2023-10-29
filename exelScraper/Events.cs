using Bunifu.UI.WinForms;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;

namespace exelScraper
{
    internal class Events
    {
        public void readExcel(BunifuDataGridView dataGrid)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel Sheet(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
            dialog.Multiselect = false;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                String path = dialog.FileName;

                FileStream fs = File.Open(path, FileMode.Open, FileAccess.Read);
                var reader = ExcelReaderFactory.CreateReader(fs);
                var result = reader.AsDataSet();
                dataGrid.DataSource = result.Tables[0];
                for (int i = 0; i < result.Tables[0].Columns.Count; i++)
                {
                    dataGrid.Columns[i].HeaderText = result.Tables[0].Rows[0][i].ToString();
                }
                dataGrid.Rows.RemoveAt(0);
            }
        }

        public void getAvrage(BunifuDataGridView dataGrid, BunifuDropdown avrageSelectBox)
        {
            decimal totalPrice = 0;
            string desiredColumnHeaderText = avrageSelectBox.Text;
            var error = false;
            var found = false;
            var legnth = 0;

            foreach (DataGridViewColumn column in dataGrid.Columns)
            {
                if (column.HeaderText == desiredColumnHeaderText)
                {
                    found = true;
                    int columnIndex = column.Index;

                    foreach (DataGridViewRow row in dataGrid.Rows)
                    {
                        if (row.Cells[columnIndex].Value != null)
                        {
                            try
                            {
                                decimal price = Convert.ToDecimal(row.Cells[columnIndex].Value);
                                totalPrice += price;
                                legnth++;
                            }
                            catch
                            {
                                error = true;
                            };
                        }
                    }

                    break;
                }
            }

            if (!found)
            {
                MessageBox.Show("There is not any column with name: " + desiredColumnHeaderText);
            }

            if (error)
            {
                MessageBox.Show("Column " + desiredColumnHeaderText + " has fields that can't convert to decimal!");
            }

            if (legnth > 0)
            {
                MessageBox.Show("Avrage is: " + Math.Round((totalPrice / legnth) * 1000) / 1000);
            }
        }

        public void getQuantile(BunifuDataGridView dataGrid, BunifuDropdown avrageSelectBox)
        {
            DataTable dataTable = (DataTable)dataGrid.DataSource;
            int columnIndex = avrageSelectBox.SelectedIndex;
            List<double> columnValues = dataTable.AsEnumerable()
            .Where(row => row.RowState != DataRowState.Deleted &&
                           row[columnIndex] != DBNull.Value &&
                           double.TryParse(row[columnIndex].ToString(), out _))
            .Select(row => Convert.ToDouble(row[columnIndex]))
            .ToList();

            var q1 = columnValues.Quantile(0.25);
            var q3 = columnValues.Quantile(0.75);

            var iqr = q3 - q1;

            var quartileDeviation = iqr / 2;

            MessageBox.Show("Quartile deviation is: " + quartileDeviation);
        }

        public void getRange(BunifuDataGridView dataGrid, BunifuDropdown avrageSelectBox)
        {
            if (avrageSelectBox.Text.Length > 0)
            {
                DataTable dataTable = (DataTable)dataGrid.DataSource;

                var filteredRows = dataTable.AsEnumerable()
                .Where(row => row.RowState != DataRowState.Deleted &&
                   double.TryParse(row[avrageSelectBox.SelectedIndex].ToString(), out _));

                var minValue = filteredRows.AsEnumerable().Min(row => (double)row[avrageSelectBox.SelectedIndex]);
                var maxValue = filteredRows.AsEnumerable().Max(row => (double)row[avrageSelectBox.SelectedIndex]);

                MessageBox.Show("Range of change is: " + (maxValue - minValue));
            }
        }

        public void variance(BunifuDataGridView dataGridView, BunifuDropdown column)
        {
            int columnIndex = column.SelectedIndex;
            double[] values = dataGridView.Rows.Cast<DataGridViewRow>()
                .Where(row => row.Cells[columnIndex].Value != null && row.Cells[columnIndex].Value != DBNull.Value)
                .Select(row => Convert.ToDouble(row.Cells[columnIndex].Value))
                .ToArray();
            if (values.Length > 1)
            {
                double avg = values.Average();
                double variance = values.Select(val => Math.Pow(val - avg, 2)).Sum();
                MessageBox.Show("Variance is: " + Convert.ToString(variance / (values.Length - 1)));
            }
            else
            {
                MessageBox.Show(Convert.ToString(0.0));
            }
        }

        public void standardDeviation(BunifuDataGridView dataGridView, BunifuDropdown column)
        {
            int columnIndex = column.SelectedIndex;
            double[] values = dataGridView.Rows.Cast<DataGridViewRow>()
                .Where(row => row.Cells[columnIndex].Value != null && row.Cells[columnIndex].Value != DBNull.Value)
                .Select(row => Convert.ToDouble(row.Cells[columnIndex].Value))
                .ToArray();
            if (values.Length > 1)
            {
                double avg = values.Average();
                double variance = values.Select(val => Math.Pow(val - avg, 2)).Sum();
                MessageBox.Show("Standard deviation is: " + Convert.ToString(Math.Round(Math.Sqrt(variance / (values.Length - 1)) * 1000) / 1000));
            }
            else
            {
                MessageBox.Show(Convert.ToString(0.0));
            }
        }

        public void covariance(DataGridView dataGridView, BunifuDropdown columnName1, BunifuDropdown columnName2)
        {
            int columnIndex1 = columnName1.SelectedIndex;
            int columnIndex2 = columnName2.SelectedIndex;
            double[] values1 = dataGridView.Rows.Cast<DataGridViewRow>()
                .Where(row => row.Cells[columnIndex1].Value != null && row.Cells[columnIndex1].Value != DBNull.Value)
                .Select(row => Convert.ToDouble(row.Cells[columnIndex1].Value))
                .ToArray();
            double[] values2 = dataGridView.Rows.Cast<DataGridViewRow>()
                .Where(row => row.Cells[columnIndex2].Value != null && row.Cells[columnIndex2].Value != DBNull.Value)
                .Select(row => Convert.ToDouble(row.Cells[columnIndex2].Value))
                .ToArray();
            double avg1 = values1.Average();
            double avg2 = values2.Average();
            double covariance = values1.Zip(values2, (x, y) => (x - avg1) * (y - avg2)).Sum();
            MessageBox.Show("cov is: " + Convert.ToString(covariance / (values1.Length - 1)));
        }

        public void correlationCoefficient(DataGridView dataGridView, BunifuDropdown columnName1, BunifuDropdown columnName2)
        {
            int columnIndex1 = columnName1.SelectedIndex;
            int columnIndex2 = columnName2.SelectedIndex;
            double[] values1 = dataGridView.Rows.Cast<DataGridViewRow>()
                .Where(row => row.Cells[columnIndex1].Value != null && row.Cells[columnIndex1].Value != DBNull.Value)
                .Select(row => Convert.ToDouble(row.Cells[columnIndex1].Value))
                .ToArray();
            double[] values2 = dataGridView.Rows.Cast<DataGridViewRow>()
                .Where(row => row.Cells[columnIndex2].Value != null && row.Cells[columnIndex2].Value != DBNull.Value)
                .Select(row => Convert.ToDouble(row.Cells[columnIndex2].Value))
                .ToArray();
            double avg1 = values1.Average();
            double avg2 = values2.Average();
            double numerator = values1.Zip(values2, (x, y) => (x - avg1) * (y - avg2)).Sum();
            double denominator = Math.Sqrt(values1.Select(x => Math.Pow(x - avg1, 2)).Sum() * values2.Select(y => Math.Pow(y - avg2, 2)).Sum());
             ;
            MessageBox.Show("Correlation coefficient is: " + Convert.ToString(numerator / denominator));
        }
    }

    public static class EnumerableExtensions
    {
        public static double Quantile(this IEnumerable<double> sequence, double quantile)
        {
            var sortedSequence = sequence.OrderBy(x => x).ToList();
            var position = (sortedSequence.Count - 1) * quantile;
            var lowerIndex = (int)position;
            var upperIndex = lowerIndex + 1;
            var lowerValue = sortedSequence[lowerIndex];
            var upperValue = sortedSequence[upperIndex];
            var interpolation = position - lowerIndex;
            return lowerValue + (upperValue - lowerValue) * interpolation;
        }
    }
}
