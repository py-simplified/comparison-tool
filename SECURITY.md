# Security Guide - Excel Comparison Tool

## 🔒 Password Protection Overview

The Excel Comparison Tool includes built-in password protection to prevent unauthorized access to sensitive financial data and comparison results.

## 🔑 Default Configuration

### Default Password
- **Password**: `1234`
- **Format**: 4-digit numeric code
- **Storage**: SHA-256 hashed in source code
- **Attempts**: Maximum 3 failed attempts before lockout

### Security Hash
```
Default Password: 1234
SHA-256 Hash: 03ac674216f3e15c761ee1a5e255f067953623c8b388b4459e13f978d7c846f4
```

## 🛠️ Password Management Options

### Option 1: Using Built-in Password Change
```bash
# Command line
python test.py --change-password
python excel_comparator_advanced.py --change-password

# Windows batch file
change_password.bat

# PowerShell script
change_password.ps1

# Unix/Linux shell script  
change_password.sh
```

### Option 2: Using Password Configuration Utility
```bash
# Command line (updates source code directly)
python password_config.py

# Windows batch file
password_config.bat
```

### Option 3: Manual Configuration
1. Generate password hash:
   ```python
   import hashlib
   password = "your_new_password"
   hash_value = hashlib.sha256(password.encode()).hexdigest()
   print(hash_value)
   ```

2. Update the `password_hash` variable in:
   - `test.py`
   - `excel_comparator_advanced.py`

## 🔐 Security Features

### Password Input Security
- **Hidden Input**: Password characters are not displayed on screen
- **No Echo**: Terminal doesn't show typed characters
- **Memory Protection**: Password is not stored in memory longer than necessary

### Hash Security
- **SHA-256 Algorithm**: Industry-standard cryptographic hash function
- **Salt-free**: Simple implementation for 4-digit codes
- **One-way**: Cannot reverse hash to get original password

### Access Control
- **Three-Strike Rule**: Maximum 3 incorrect attempts
- **Immediate Lockout**: Access denied after failed attempts
- **Session-based**: Password required for each execution

## 🚨 Security Best Practices

### Password Selection
- ✅ **Do**: Use a unique 4-digit code
- ✅ **Do**: Avoid common patterns (1234, 0000, 1111)
- ✅ **Do**: Choose numbers with personal significance but not obvious
- ❌ **Don't**: Use birthdates, addresses, or public information
- ❌ **Don't**: Share password via email or text messages

### Password Management
- 🔄 **Change Regularly**: Update password monthly or quarterly
- 📝 **Document Securely**: Store in password manager or secure location
- 👥 **Limit Access**: Share only with authorized personnel
- 🔍 **Monitor Usage**: Be aware of who has access

### File Security
- 💾 **Backup Regularly**: Keep secure backups of configuration
- 🔒 **Protect Source Code**: Limit access to Python files
- 📁 **Secure Storage**: Store project files in protected directories
- 🛡️ **Version Control**: Use private repositories for sensitive projects

## 🔧 Troubleshooting

### Common Issues

**Issue**: "Password incorrect" despite entering correct password
- **Solution**: Ensure no extra spaces or characters
- **Solution**: Check if Caps Lock is on (though numbers aren't affected)
- **Solution**: Verify the hash in source code matches your password

**Issue**: "Maximum attempts exceeded"
- **Solution**: Restart the application to reset attempt counter
- **Solution**: Use password change utility to verify current password

**Issue**: Password change utility not working
- **Solution**: Ensure virtual environment is activated
- **Solution**: Check file permissions for writing to Python files
- **Solution**: Run as administrator if necessary

### Recovery Options

**If you forget the password:**
1. **Option 1**: Manually reset the hash in source code to default
   ```python
   self.password_hash = "03ac674216f3e15c761ee1a5e255f067953623c8b388b4459e13f978d7c846f4"
   ```
   (This resets password to "1234")

2. **Option 2**: Edit the source code to temporarily bypass password check
   ```python
   def verify_password(self):
       return True  # Temporarily bypass
   ```

3. **Option 3**: Use the password configuration utility with admin privileges

## 📊 Security Audit

### Regular Security Checks
- [ ] Password changed from default
- [ ] Password is not easily guessable
- [ ] Source code files are protected
- [ ] Access logs reviewed (if applicable)
- [ ] Backup passwords are secure
- [ ] Team members trained on security practices

### Compliance Considerations
- **Data Protection**: Ensure password protection meets organizational requirements
- **Audit Trail**: Consider logging access attempts for compliance
- **Encryption**: Consider additional encryption for highly sensitive data
- **Access Control**: Implement role-based access if needed

## 🆘 Emergency Procedures

### In Case of Security Breach
1. **Immediately** change the password using password configuration utility
2. **Review** recent access and file modifications
3. **Check** comparison_results folder for unauthorized access
4. **Update** all team members with new password securely
5. **Consider** additional security measures if breach is severe

### For Critical Systems
- Implement additional encryption layers
- Use hardware security modules (HSM) for enterprise environments
- Consider multi-factor authentication for high-security requirements
- Regular security audits and penetration testing

## 📞 Support and Resources

### Getting Help
- Check this security guide first
- Review README.md for general usage
- Check error messages for specific issues
- Contact system administrator for enterprise deployments

### Additional Resources
- [Python Security Best Practices](https://docs.python.org/3/library/security.html)
- [SHA-256 Information](https://en.wikipedia.org/wiki/SHA-2)
- [Password Security Guidelines](https://www.nist.gov/password-guidelines)

---

**⚠️ Important**: This security implementation is designed for basic protection of business data. For highly sensitive or regulated environments, consider additional security measures and professional security assessment.
