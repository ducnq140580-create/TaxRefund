using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TaxRefund
{
    public class PasswordStrengthChecker
    {
        public enum PasswordStrength
        {
            Blank,
            VeryWeak,
            Weak,
            Medium,
            Strong,
            VeryStrong
        }

        public static PasswordStrength GetPasswordStrength(string password)
        {
            int score = 0;

            if (string.IsNullOrEmpty(password))
                return PasswordStrength.Blank;

            // Length check
            if (password.Length >= 8) score++;
            if (password.Length >= 12) score++;

            // Complexity checks
            if (Regex.IsMatch(password, @"[0-9]+")) score++; // Numbers
            if (Regex.IsMatch(password, @"[a-z]")) score++; // Lowercase
            if (Regex.IsMatch(password, @"[A-Z]")) score++; // Uppercase
            if (Regex.IsMatch(password, @"[!@#$%^&*()_+=\[{\]};:<>|./?,-]")) score++; // Special chars

            //return score switch
            //{
            //    0 => PasswordStrength.VeryWeak,
            //    1 => PasswordStrength.Weak,
            //    2 => PasswordStrength.Medium,
            //    3 => PasswordStrength.Strong,
            //    _ => PasswordStrength.VeryStrong,
            //};

            if (score == 0) return PasswordStrength.VeryWeak;
            if (score == 1) return PasswordStrength.Weak;
            if (score == 2) return PasswordStrength.Medium;
            if (score == 3) return PasswordStrength.Strong;        
            return PasswordStrength.VeryStrong;

        }
    }
}
