using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace ExpenseApp
{
    public partial class LoginForm : Form
    {
        string connectionString = @"Server=IAMVANJARE;Database=ExpenseDB;Trusted_Connection=True;";

        TextBox txtUsername, txtPassword;
        Button btnLogin;
        Label lblMessage;

        public LoginForm()
        {
            InitializeComponent();
            this.Text = "Login";
            this.Size = new Size(350, 250);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            Label lblUser = new Label() { Text = "Username", Location = new Point(30, 30), AutoSize = true };
            txtUsername = new TextBox() { Location = new Point(120, 25), Width = 150 };

            Label lblPass = new Label() { Text = "Password", Location = new Point(30, 70), AutoSize = true };
            txtPassword = new TextBox() { Location = new Point(120, 65), Width = 150, PasswordChar = '*' };

            btnLogin = new Button() { Text = "Login", Location = new Point(120, 110), Width = 100, Height = 35 };
            btnLogin.Click += BtnLogin_Click;

            lblMessage = new Label() { ForeColor = Color.Red, Location = new Point(30, 160), AutoSize = true };

            this.Controls.AddRange(new Control[] { lblUser, txtUsername, lblPass, txtPassword, btnLogin, lblMessage });
        }

        private void BtnLogin_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtUsername.Text) || string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                lblMessage.Text = "Enter username and password";
                return;
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    var cmd = new SqlCommand("SELECT COUNT(*) FROM Users WHERE Username=@user AND Password=@pass", conn);
                    cmd.Parameters.AddWithValue("@user", txtUsername.Text.Trim());
                    cmd.Parameters.AddWithValue("@pass", txtPassword.Text.Trim()); // Plain text for now

                    int count = (int)cmd.ExecuteScalar();
                    if (count > 0)
                    {
                        this.Hide();
                        Expense expenseForm = new Expense();
                        expenseForm.ShowDialog();
                        this.Close();
                    }
                    else
                    {
                        lblMessage.Text = "Invalid username or password";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error logging in: " + ex.Message);
            }
        }
    }
}
