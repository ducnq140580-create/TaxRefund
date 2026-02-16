using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TaxRefund
{
    public class PasswordStrengthControl : UserControl
    {
        private TextBox passwordTextBox;
        private ProgressBar strengthBar;
        private Label strengthLabel;
        private ToolTip strengthToolTip;

        public TextBox PasswordTextBox
        {
            get => passwordTextBox;
            set
            {
                if (passwordTextBox != null)
                    passwordTextBox.TextChanged -= PasswordTextChanged;

                passwordTextBox = value;

                if (passwordTextBox != null)
                    passwordTextBox.TextChanged += PasswordTextChanged;
            }
        }

        public PasswordStrengthControl()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.SuspendLayout();

            // Progress Bar
            strengthBar = new ProgressBar
            {
                Dock = DockStyle.Top,
                Height = 10,
                Minimum = 0,
                Maximum = 100,
                Style = ProgressBarStyle.Continuous
            };

            // Strength Label
            strengthLabel = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(5, 0, 0, 0)
            };

            // Tooltip
            strengthToolTip = new ToolTip
            {
                AutomaticDelay = 500,
                AutoPopDelay = 5000,
                InitialDelay = 500,
                ReshowDelay = 100
            };

            // Control Layout
            this.Controls.Add(strengthLabel);
            this.Controls.Add(strengthBar);
            this.Height = 30;

            this.ResumeLayout(false);
        }

        private void PasswordTextChanged(object sender, EventArgs e)
        {
            if (PasswordTextBox == null) return;

            var strength = PasswordStrengthChecker.GetPasswordStrength(PasswordTextBox.Text);
            UpdateStrengthDisplay(strength);
        }

        private void UpdateStrengthDisplay(PasswordStrengthChecker.PasswordStrength strength)
        {
            int strengthValue = (int)strength * 20;
            strengthBar.Value = strengthValue;

            string tooltipText = "Password should contain:\n- At least 8 characters\n- Uppercase letters\n- Numbers\n- Special characters";

            switch (strength)
            {
                case PasswordStrengthChecker.PasswordStrength.Blank:
                    strengthBar.ForeColor = SystemColors.Control;
                    strengthLabel.Text = "";
                    break;
                case PasswordStrengthChecker.PasswordStrength.VeryWeak:
                    strengthBar.ForeColor = Color.Red;
                    strengthLabel.Text = "Very Weak";
                    strengthToolTip.SetToolTip(strengthBar, "Very weak password\n" + tooltipText);
                    break;
                case PasswordStrengthChecker.PasswordStrength.Weak:
                    strengthBar.ForeColor = Color.Orange;
                    strengthLabel.Text = "Weak";
                    strengthToolTip.SetToolTip(strengthBar, "Weak password\n" + tooltipText);
                    break;
                case PasswordStrengthChecker.PasswordStrength.Medium:
                    strengthBar.ForeColor = Color.Yellow;
                    strengthLabel.Text = "Medium";
                    strengthToolTip.SetToolTip(strengthBar, "Medium strength\n" + tooltipText);
                    break;
                case PasswordStrengthChecker.PasswordStrength.Strong:
                    strengthBar.ForeColor = Color.LightGreen;
                    strengthLabel.Text = "Strong";
                    strengthToolTip.SetToolTip(strengthBar, "Strong password");
                    break;
                case PasswordStrengthChecker.PasswordStrength.VeryStrong:
                    strengthBar.ForeColor = Color.Green;
                    strengthLabel.Text = "Very Strong";
                    strengthToolTip.SetToolTip(strengthBar, "Very strong password");
                    break;
            }
        }
    }
}
