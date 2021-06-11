using DevExpress.XtraBars;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KerinoProductsImport
{
    public partial class Form1 : DevExpress.XtraBars.FluentDesignSystem.FluentDesignForm
    {
        DataTable dimension1, dimension2, dimension3, finalTable;

        public Form1()
        {
            InitializeComponent();
            labelCount.Text = null;
            dimension1 = new DataTable();
            dimension1.Columns.Add("Κωδικός", typeof(string));
            dimension1.Columns.Add("Περιγραφή", typeof(string));
            dimension2 = dimension1.Clone();
            dimension3 = dimension1.Clone();
            gridControl2.DataSource = dimension1;
            gridControl3.DataSource = dimension2;
            gridControl4.DataSource = dimension3;
            gridView2.OptionsView.NewItemRowPosition = NewItemRowPosition.Top;
            gridView3.OptionsView.NewItemRowPosition = NewItemRowPosition.Top;
            gridView4.OptionsView.NewItemRowPosition = NewItemRowPosition.Top;
            gridView2.ViewCaption = "Διάσταση 1";
            gridView3.ViewCaption = "Διάσταση 2";
            gridView4.ViewCaption = "Διάσταση 3";

            finalTable = new DataTable();
            finalTable.Columns.Add("Κωδικός", typeof(string));
            finalTable.Columns.Add("ΕΑΝ", typeof(string));
            finalTable.Columns.Add("Περιγραφή", typeof(string));
            finalTable.Columns.Add("ΜΜ", typeof(string));
            finalTable.Columns.Add("Ομάδα", typeof(string));
            finalTable.Columns.Add("ΦΠΑ%", typeof(double));
            finalTable.Columns.Add("Τιμή", typeof(double));

            gridControl1.DataSource = finalTable;
        }

        private void accordionControlElement2_Click(object sender, EventArgs e)
        {
            //Δημιουργία Πίνακα
            string kwdBase = textBoxKwd.Text;
            string nameBase = textBoxDescription.Text;
            string mm = textBoxMM.Text;
            string group = textBoxGroup.Text;
            double fpa = Convert.ToDouble(numericUpDownVat.Value.ToString());
            double price = Convert.ToDouble(numericUpDownPrice.Value.ToString());

            for (int a = 0; a < gridView2.RowCount; a++)
            {
                string codeA = gridView2.GetRowCellDisplayText(a, "Κωδικός");
                string perA = gridView2.GetRowCellDisplayText(a, "Περιγραφή");
                if (gridView3.RowCount == 0)
                {
                    DataRow row = finalTable.NewRow();
                    row["Κωδικός"] = $"{kwdBase}-{codeA}";
                    row["Περιγραφή"] = $"{nameBase} {perA}";
                    row["ΜΜ"] = mm;
                    row["Ομάδα"] = group;
                    row["ΦΠΑ%"] = fpa;
                    row["Τιμή"] = price;
                    finalTable.Rows.Add(row);
                }
                else
                {
                    for (int b = 0; b < gridView3.RowCount; b++)
                    {
                        string codeB = gridView3.GetRowCellDisplayText(b, "Κωδικός");
                        string perB = gridView3.GetRowCellDisplayText(b, "Περιγραφή");
                        if (gridView4.RowCount == 0)
                        {
                            DataRow row = finalTable.NewRow();
                            row["Κωδικός"] = $"{kwdBase}-{codeA}-{codeB}";
                            row["Περιγραφή"] = $"{nameBase} {perA} {perB}";
                            row["ΜΜ"] = mm;
                            row["Ομάδα"] = group;
                            row["ΦΠΑ%"] = fpa;
                            row["Τιμή"] = price;
                            finalTable.Rows.Add(row);
                        }
                        else
                        {
                            for (int c = 0; c < gridView4.RowCount; c++)
                            {
                                string codeC = gridView4.GetRowCellDisplayText(c, "Κωδικός");
                                string perC = gridView4.GetRowCellDisplayText(c, "Περιγραφή");

                                DataRow row = finalTable.NewRow();
                                row["Κωδικός"] = $"{kwdBase}-{codeA}-{codeB}-{codeC}";
                                row["Περιγραφή"] = $"{nameBase} {perA} {perB} {perC}";
                                row["ΜΜ"] = mm;
                                row["Ομάδα"] = group;
                                row["ΦΠΑ%"] = fpa;
                                row["Τιμή"] = price;
                                finalTable.Rows.Add(row);
                            }
                        }
                    }
                }
            }
            gridView1.BestFitColumns();
            labelCount.Text = $"Σύνολο παραλαγών {finalTable.Rows.Count}";
        }

        private void accordionControlElement3_Click(object sender, EventArgs e)
        {
            //Εξαγωγή Excel
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Αρχεία Excel (*.xlsx)|*.xlsx";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.CreatePrompt = true;
            saveFileDialog.Title = "Αποθήκευση αρχείου Excel";

            DialogResult res = saveFileDialog.ShowDialog();
            if (res == DialogResult.OK)
            {
                gridControl1.ExportToXlsx(saveFileDialog.FileName);
            }

            
        }

        private void accordionControlElement4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < gridView2.RowCount;)
                gridView2.DeleteRow(i);
        }

        private void accordionControlElement5_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < gridView3.RowCount;)
                gridView3.DeleteRow(i);
        }

        private void accordionControlElement6_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < gridView4.RowCount;)
                gridView4.DeleteRow(i);
        }

        private void accordionControlElement7_Click(object sender, EventArgs e)
        {
            finalTable.Rows.Clear();
            labelCount.Text = null;
        }
    }
}
