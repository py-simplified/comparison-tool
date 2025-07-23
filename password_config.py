import hashlib
import getpass
import re
import os

def update_password_hash_in_file(filepath, new_hash):
    """
    Update the password hash in a Python file
    
    Args:
        filepath (str): Path to the Python file
        new_hash (str): New password hash to set
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            content = file.read()
        
        # Pattern to match the password_hash assignment
        pattern = r'self\.password_hash\s*=\s*"[^"]*"'
        replacement = f'self.password_hash = "{new_hash}"'
        
        # Check if pattern exists
        if re.search(pattern, content):
            # Replace the password hash
            new_content = re.sub(pattern, replacement, content)
            
            # Write back to file
            with open(filepath, 'w', encoding='utf-8') as file:
                file.write(new_content)
            
            print(f"✅ Updated password hash in {filepath}")
            return True
        else:
            print(f"❌ Could not find password_hash variable in {filepath}")
            return False
            
    except Exception as e:
        print(f"❌ Error updating {filepath}: {e}")
        return False

def main():
    """
    Password Configuration Utility
    """
    print("🔧 Password Configuration Utility")
    print("=" * 40)
    print("This utility will update the password hash in all Python files.")
    print("⚠️  Warning: This will modify the source code files directly!")
    print()
    
    # Get current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Files to update
    files_to_update = [
        os.path.join(current_dir, "test.py"),
        os.path.join(current_dir, "excel_comparator_advanced.py")
    ]
    
    # Check if files exist
    missing_files = [f for f in files_to_update if not os.path.exists(f)]
    if missing_files:
        print("❌ Missing files:")
        for f in missing_files:
            print(f"   - {f}")
        return
    
    # Get new password
    while True:
        new_password = getpass.getpass("Enter new 4-digit password: ")
        
        if not new_password.isdigit() or len(new_password) != 4:
            print("❌ Invalid format! Password must be exactly 4 digits.")
            continue
        
        confirm_password = getpass.getpass("Confirm new password: ")
        
        if new_password != confirm_password:
            print("❌ Passwords don't match! Please try again.")
            continue
        
        break
    
    # Generate hash
    new_hash = hashlib.sha256(new_password.encode()).hexdigest()
    print(f"\n🔐 Generated hash: {new_hash}")
    
    # Confirm update
    confirm = input("\n⚠️  Are you sure you want to update the password in all files? (y/N): ")
    if confirm.lower() != 'y':
        print("❌ Operation cancelled.")
        return
    
    # Update files
    print("\n📝 Updating files...")
    success_count = 0
    
    for filepath in files_to_update:
        if update_password_hash_in_file(filepath, new_hash):
            success_count += 1
    
    print(f"\n✅ Successfully updated {success_count}/{len(files_to_update)} files.")
    
    if success_count == len(files_to_update):
        print("🎉 Password update completed successfully!")
        print("🔒 New password is now active in all components.")
    else:
        print("⚠️  Some files could not be updated. Please check manually.")
    
    print("\n📋 Next steps:")
    print("1. Test the new password with the comparison tool")
    print("2. Keep the new password secure and confidential")
    print("3. Consider backing up your files before making changes")

if __name__ == "__main__":
    main()
