# GitHub Repository Setup Commands

## After creating the repository on GitHub, run these commands:

# 1. Add the remote repository (replace 'yourusername' with your GitHub username)
git remote add origin https://github.com/yourusername/excel-comparison-tool.git

# 2. Rename the default branch to 'main' (GitHub's current default)
git branch -M main

# 3. Push the code to GitHub
git push -u origin main

## Alternative: If you prefer SSH (requires SSH key setup)
# git remote add origin git@github.com:yourusername/excel-comparison-tool.git
# git branch -M main
# git push -u origin main

## Repository Details to Use:
Repository Name: excel-comparison-tool
Description: Automated Excel file comparison tool with password protection and visual difference highlighting
Visibility: Private
Initialize with: None (we already have README, .gitignore, and LICENSE)

## What's included in this repository:
- Complete Excel comparison tool with password protection
- Cross-platform scripts (Windows, Linux, macOS)
- Comprehensive documentation
- Security features with 4-digit password protection
- Automated setup scripts
- Example Excel files for testing
