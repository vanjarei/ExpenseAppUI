using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace ExpenseApp
{
    public partial class Expense : Form
    {
        Label lblDate, lblCategory, lblDescription, lblAmount;
        Label lblTotalToday, lblTotalMonth;
        DateTimePicker dtpDate;
        ComboBox cmbCategory;
        TextBox txtDescription, txtAmount;
        Button btnAddExpense, btnDelete, btnEditExpense, btnExportExcel;
        ListView lvExpenses;

        string connectionString = @"Server=IAMVANJARE;Database=ExpenseDB;Trusted_Connection=True;";
        int editingExpenseId = -1;

        public Expense()
        {
            InitializeComponent();

            this.Text = "Expense Tracker";
            this.ClientSize = new Size(650, 600);
            this.BackColor = Color.FromArgb(245, 246, 250);
            this.Font = new Font("Segoe UI", 10);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            Font labelFont = new Font("Segoe UI", 10, FontStyle.Bold);
            Color labelColor = Color.FromArgb(50, 50, 50);

            lblDate = new Label() { Text = "Date", Location = new Point(30, 30), AutoSize = true, Font = labelFont, ForeColor = labelColor };
            dtpDate = new DateTimePicker() { Location = new Point(150, 25), Width = 200, Format = DateTimePickerFormat.Short };

            lblCategory = new Label() { Text = "Category", Location = new Point(30, 75), AutoSize = true, Font = labelFont, ForeColor = labelColor };
            cmbCategory = new ComboBox() { Location = new Point(150, 70), Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };

            lblDescription = new Label() { Text = "Description", Location = new Point(30, 120), AutoSize = true, Font = labelFont, ForeColor = labelColor };
            txtDescription = new TextBox() { Location = new Point(150, 115), Width = 350 };

            lblAmount = new Label() { Text = "Amount", Location = new Point(30, 165), AutoSize = true, Font = labelFont, ForeColor = labelColor };
            txtAmount = new TextBox() { Location = new Point(150, 160), Width = 200 };

            int buttonWidth = 160, buttonHeight = 40, startX = 30, spacing = 20, yPos = 210;

            btnAddExpense = new Button() { Text = "Add Expense", Location = new Point(startX, yPos), Width = buttonWidth, Height = buttonHeight };
            btnEditExpense = new Button() { Text = "Edit Expense", Location = new Point(startX + buttonWidth + spacing, yPos), Width = buttonWidth, Height = buttonHeight, Enabled = false };
            btnDelete = new Button() { Text = "Delete", Location = new Point(startX + 2 * (buttonWidth + spacing), yPos), Width = buttonWidth, Height = buttonHeight };
            btnExportExcel = new Button() { Text = "Export to Excel", Location = new Point(30, 530), Width = 180, Height = 40 };

            lvExpenses = new ListView()
            {
                Location = new Point(30, 270),
                Width = 580,
                Height = 240,
                View = View.Details,
                FullRowSelect = true,
                GridLines = false,
                Font = new Font("Segoe UI", 10),
                BackColor = Color.White
            };
            lvExpenses.Columns.Add("Id", 0);
            lvExpenses.Columns.Add("Date", 100);
            lvExpenses.Columns.Add("Category", 120);
            lvExpenses.Columns.Add("Description", 260);
            lvExpenses.Columns.Add("Amount", 70, HorizontalAlignment.Right);

            lblTotalToday = new Label() { Location = new Point(350, 530), AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            lblTotalMonth = new Label() { Location = new Point(350, 555), AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold) };

            // Add controls
            this.Controls.AddRange(new Control[] {
                lblDate, dtpDate, lblCategory, cmbCategory, lblDescription, txtDescription,
                lblAmount, txtAmount, btnAddExpense, btnEditExpense, btnDelete,
                btnExportExcel, lvExpenses, lblTotalToday, lblTotalMonth
            });

            // Apply Modern Styles
            ApplyModernUI();

            // Events
            btnAddExpense.Click += btnAddExpense_Click;
            btnEditExpense.Click += btnEditExpense_Click;
            btnDelete.Click += btnDelete_Click;
            btnExportExcel.Click += btnExportExcel_Click;
            lvExpenses.SelectedIndexChanged += LvExpenses_SelectedIndexChanged;

            LoadCategories();
            LoadExpenses();
            UpdateTotalsFromDB();
        }

        // ----------------- UI Styling -----------------
        private void ApplyModernUI()
        {
            StyleButton(btnAddExpense, Color.FromArgb(0, 120, 215), Color.White);
            StyleButton(btnEditExpense, Color.FromArgb(255, 193, 7), Color.Black);
            StyleButton(btnDelete, Color.FromArgb(220, 53, 69), Color.White);
            StyleButton(btnExportExcel, Color.FromArgb(40, 167, 69), Color.White);

            lvExpenses.BorderStyle = BorderStyle.None;
            lvExpenses.HeaderStyle = ColumnHeaderStyle.Nonclickable;
        }

        private void StyleButton(Button btn, Color backColor, Color foreColor)
        {
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.BackColor = backColor;
            btn.ForeColor = foreColor;
            btn.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            btn.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, btn.Width, btn.Height, 10, 10));
        }

        [DllImport("gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect,
            int nWidthEllipse, int nHeightEllipse);

        // ----------------- Database Methods -----------------
        private void LoadCategories()
        {
            cmbCategory.Items.Clear();
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    var cmd = new SqlCommand("SELECT CategoryName FROM Categories ORDER BY CategoryName", conn);
                    var reader = cmd.ExecuteReader();
                    while (reader.Read()) cmbCategory.Items.Add(reader["CategoryName"].ToString());
                }
                if (cmbCategory.Items.Count > 0) cmbCategory.SelectedIndex = 0;
            }
            catch (Exception ex) { MessageBox.Show("Error loading categories: " + ex.Message); }
        }

        private void LoadExpenses()
        {
            lvExpenses.Items.Clear();
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    var cmd = new SqlCommand("SELECT Id, ExpenseDate, Category, Description, Amount FROM Expenses ORDER BY ExpenseDate DESC", conn);
                    var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var item = new ListViewItem(reader["Id"].ToString());
                        item.SubItems.Add(((DateTime)reader["ExpenseDate"]).ToShortDateString());
                        item.SubItems.Add(reader["Category"].ToString());
                        item.SubItems.Add(reader["Description"].ToString());
                        item.SubItems.Add(((decimal)reader["Amount"]).ToString("F2"));
                        lvExpenses.Items.Add(item);
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Error loading expenses: " + ex.Message); }
        }

        private void UpdateTotalsFromDB()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    var cmdToday = new SqlCommand("SELECT ISNULL(SUM(Amount), 0) FROM Expenses WHERE ExpenseDate = CAST(GETDATE() AS DATE)", conn);
                    lblTotalToday.Text = $"Today: {(decimal)cmdToday.ExecuteScalar():C2}";

                    var cmdMonth = new SqlCommand("SELECT ISNULL(SUM(Amount), 0) FROM Expenses WHERE YEAR(ExpenseDate) = YEAR(GETDATE()) AND MONTH(ExpenseDate) = MONTH(GETDATE())", conn);
                    lblTotalMonth.Text = $"This Month: {(decimal)cmdMonth.ExecuteScalar():C2}";
                }
            }
            catch (Exception ex) { MessageBox.Show("Error calculating totals: " + ex.Message); }
        }

        // ----------------- Event Handlers -----------------
        private void LvExpenses_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasSelection = lvExpenses.SelectedItems.Count > 0;
            btnEditExpense.Enabled = hasSelection;
            btnDelete.Enabled = hasSelection;
        }

        private bool ValidateInputs(out decimal amount)
        {
            amount = 0;
            if (!decimal.TryParse(txtAmount.Text, out amount) || amount <= 0)
            {
                MessageBox.Show("Enter a valid positive amount."); return false;
            }
            if (string.IsNullOrWhiteSpace(txtDescription.Text))
            {
                MessageBox.Show("Enter a description."); return false;
            }
            if (cmbCategory.SelectedIndex < 0)
            {
                MessageBox.Show("Select a category."); return false;
            }
            return true;
        }

        private void ClearInputs()
        {
            txtDescription.Clear();
            txtAmount.Clear();
            dtpDate.Value = DateTime.Now;
            if (cmbCategory.Items.Count > 0) cmbCategory.SelectedIndex = 0;
            editingExpenseId = -1;
            btnAddExpense.Enabled = true;
            btnEditExpense.Text = "Edit Expense";
        }

        private void btnAddExpense_Click(object sender, EventArgs e)
        {
            if (!ValidateInputs(out decimal amount)) return;
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    var cmd = new SqlCommand("INSERT INTO Expenses (ExpenseDate, Category, Description, Amount) VALUES (@date, @category, @desc, @amount)", conn);
                    cmd.Parameters.AddWithValue("@date", dtpDate.Value.Date);
                    cmd.Parameters.AddWithValue("@category", cmbCategory.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@desc", txtDescription.Text.Trim());
                    cmd.Parameters.AddWithValue("@amount", amount);
                    cmd.ExecuteNonQuery();
                }
                LoadExpenses();
                UpdateTotalsFromDB();
                ClearInputs();
            }
            catch (Exception ex) { MessageBox.Show("Error saving expense: " + ex.Message); }
        }

        private void btnEditExpense_Click(object sender, EventArgs e)
        {
            if (editingExpenseId == -1)
            {
                if (lvExpenses.SelectedItems.Count == 0) return;
                var item = lvExpenses.SelectedItems[0];
                editingExpenseId = int.Parse(item.SubItems[0].Text);
                dtpDate.Value = DateTime.Parse(item.SubItems[1].Text);
                cmbCategory.SelectedItem = item.SubItems[2].Text;
                txtDescription.Text = item.SubItems[3].Text;
                txtAmount.Text = item.SubItems[4].Text;
                btnAddExpense.Enabled = false;
                btnEditExpense.Text = "Save Edit";
            }
            else
            {
                if (!ValidateInputs(out decimal amount)) return;
                try
                {
                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        var cmd = new SqlCommand("UPDATE Expenses SET ExpenseDate=@date, Category=@cat, Description=@desc, Amount=@amt WHERE Id=@id", conn);
                        cmd.Parameters.AddWithValue("@date", dtpDate.Value.Date);
                        cmd.Parameters.AddWithValue("@cat", cmbCategory.SelectedItem.ToString());
                        cmd.Parameters.AddWithValue("@desc", txtDescription.Text.Trim());
                        cmd.Parameters.AddWithValue("@amt", amount);
                        cmd.Parameters.AddWithValue("@id", editingExpenseId);
                        cmd.ExecuteNonQuery();
                    }
                    LoadExpenses();
                    UpdateTotalsFromDB();
                    ClearInputs();
                }
                catch (Exception ex) { MessageBox.Show("Error updating expense: " + ex.Message); }
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (lvExpenses.SelectedItems.Count == 0) return;
            var id = int.Parse(lvExpenses.SelectedItems[0].SubItems[0].Text);
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    var cmd = new SqlCommand("DELETE FROM Expenses WHERE Id = @id", conn);
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }
                LoadExpenses();
                UpdateTotalsFromDB();
            }
            catch (Exception ex) { MessageBox.Show("Error deleting expense: " + ex.Message); }
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            if (lvExpenses.Items.Count == 0) { MessageBox.Show("No expenses to export."); return; }
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", FileName = "Expenses.xlsx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var wb = new XLWorkbook())
                        {
                            var ws = wb.Worksheets.Add("Expenses");
                            ws.Cell(1, 1).Value = "Date";
                            ws.Cell(1, 2).Value = "Category";
                            ws.Cell(1, 3).Value = "Description";
                            ws.Cell(1, 4).Value = "Amount";

                            int row = 2;
                            foreach (ListViewItem item in lvExpenses.Items)
                            {
                                ws.Cell(row, 1).Value = item.SubItems[1].Text;
                                ws.Cell(row, 2).Value = item.SubItems[2].Text;
                                ws.Cell(row, 3).Value = item.SubItems[3].Text;
                                ws.Cell(row, 4).Value = decimal.Parse(item.SubItems[4].Text);
                                row++;
                            }
                            ws.Range(1, 1, 1, 4).Style.Font.Bold = true;
                            ws.Column(4).Style.NumberFormat.Format = "$#,##0.00";
                            ws.Columns().AdjustToContents();
                            wb.SaveAs(sfd.FileName);
                        }
                        MessageBox.Show("Exported successfully!");
                    }
                    catch (Exception ex) { MessageBox.Show("Error exporting: " + ex.Message); }
                }
            }
        }
    }
}
