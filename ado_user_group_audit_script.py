import base64
import requests
import time
from openpyxl import Workbook
import os
from datetime import datetime

# ======================================================
# CONFIGURATION
# ======================================================
ORG = "parmeshwargudadhe"
PAT = ""
API_VERSION = "7.1-preview.1"

# ======================================================
# AUTH HEADER
# ======================================================
def headers():
    token = base64.b64encode(f":{PAT}".encode()).decode()
    return {
        "Authorization": f"Basic {token}",
        "Content-Type": "application/json"
    }

# ======================================================
# GET ALL USERS
# ======================================================
def get_all_users():
    users = []
    continuation = None

    while True:
        url = f"https://vssps.dev.azure.com/{ORG}/_apis/graph/users?api-version={API_VERSION}"
        if continuation:
            url += f"&continuationToken={continuation}"

        r = requests.get(url, headers=headers())
        r.raise_for_status()

        for u in r.json().get("value", []):
            if (
                u.get("subjectKind") == "user"
                and u.get("principalName")
                and "@" in u.get("principalName")
            ):
                users.append(u)

        continuation = r.headers.get("x-ms-continuationtoken")
        if not continuation:
            break

        time.sleep(0.2)

    return users

# ======================================================
# SEARCH USER BY EMAIL
# ======================================================
def search_user_by_email(email):
    users = get_all_users()
    for user in users:
        if user.get("principalName", "").lower() == email.lower():
            return user
    return None

# ======================================================
# GET USER GROUP MEMBERSHIPS
# ======================================================
def get_user_groups(user_descriptor):
    url = f"https://vssps.dev.azure.com/{ORG}/_apis/graph/memberships/{user_descriptor}?direction=up&api-version={API_VERSION}"
    r = requests.get(url, headers=headers())
    r.raise_for_status()
    return [m["containerDescriptor"] for m in r.json().get("value", [])]

# ======================================================
# GET GROUP DETAILS
# ======================================================
def get_group(group_descriptor):
    url = f"https://vssps.dev.azure.com/{ORG}/_apis/graph/groups/{group_descriptor}?api-version={API_VERSION}"
    r = requests.get(url, headers=headers())
    r.raise_for_status()
    return r.json()

# ======================================================
# DETERMINE SCOPE (FIXED VERSION)
# ======================================================
def determine_scope(group_data):
    display_name = group_data.get("displayName", "")
    principal_name = group_data.get("principalName", "")
    
    # Check if it's an organization-level group
    if (display_name == "Project Collection Administrators" or 
        principal_name.startswith("vssgp.") or  # This indicates it's a built-in group
        "[" not in principal_name):  # Project groups always have [ProjectName] format
        
        return "Organization", ORG  # Return actual organization name
    
    # Project-level groups look like: [ProjectName]\GroupName
    if principal_name.startswith("[") and "]" in principal_name:
        # Extract project name from format like: [ProjectName]\GroupName
        project_part = principal_name.split("]")[0]
        project_name = project_part.strip("[")
        
        # If it's something like "parmeshwargudadhe" or empty, it's likely organization level
        if project_name.lower() == "parmeshwargudadhe" or not project_name:
            return "Organization", ORG
        
        return "Project", project_name
    
    return "Organization", ORG  # Default to organization

# ======================================================
# SAVE RESULTS TO EXCEL
# ======================================================
def save_to_excel(user_data, filename=None):
    """Save user group data to Excel file"""
    if not filename:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"ado_user_groups_{timestamp}.xlsx"
    
    wb = Workbook()
    ws = wb.active
    ws.title = "User Group Memberships"
    
    # Write headers
    headers = [
        "User Email",
        "User Display Name",
        "Group Name",
        "Group Principal Name",
        "Scope Type",
        "Scope Name",
        "Group Description"
    ]
    ws.append(headers)
    
    # Write data
    total_groups = 0
    for user_email, user_display_name, groups in user_data:
        if not groups:
            # Write user even if no groups
            ws.append([
                user_email,
                user_display_name,
                "No group memberships",
                "",
                "",
                "",
                ""
            ])
        else:
            for group_info in groups:
                ws.append([
                    user_email,
                    user_display_name,
                    group_info["group_name"],
                    group_info["principal_name"],
                    group_info["scope_type"],
                    group_info["scope_name"],
                    group_info.get("description", "")
                ])
                total_groups += 1
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    wb.save(filename)
    return filename, total_groups

# ======================================================
# GET USER DETAILS WITH EXCEL EXPORT
# ======================================================
def get_user_details_with_export(email=None, save_to_excel_flag=True):
    """Get user details and optionally save to Excel"""
    if email:
        user = search_user_by_email(email)
        if not user:
            print(f"❌ User with email '{email}' not found.")
            return
        users = [user]
        print(f"🔍 Found user: {user['principalName']}")
        user_list_desc = f"single_user_{email.split('@')[0]}"
    else:
        users = get_all_users()
        print(f"📊 Total users found: {len(users)}")
        user_list_desc = f"all_{len(users)}_users"
    
    print("\n" + "="*80)
    print("USER GROUP MEMBERSHIPS")
    print("="*80)
    
    all_user_data = []
    group_count = 0
    
    for user in users:
        email = user["principalName"]
        display_name = user.get("displayName", "N/A")
        print(f"\n👤 User: {email} ({display_name})")
        print("-" * 40)
        
        try:
            groups = get_user_groups(user["descriptor"])
        except Exception as e:
            print(f"⚠️ Failed to get groups: {e}")
            user_groups = []
            all_user_data.append((email, display_name, user_groups))
            continue
        
        if not groups:
            print("No group memberships found.")
            user_groups = []
            all_user_data.append((email, display_name, user_groups))
            continue
        
        user_groups = []
        for i, group_descriptor in enumerate(groups, 1):
            try:
                gd = get_group(group_descriptor)
                
                # Skip Security Service Group
                if gd.get("displayName") == "Security Service Group":
                    continue
                
                scope_type, scope_name = determine_scope(gd)
                
                group_info = {
                    "group_name": gd.get("displayName", "Unknown"),
                    "principal_name": gd.get("principalName", ""),
                    "scope_type": scope_type,
                    "scope_name": scope_name,
                    "description": gd.get("description", "")
                }
                user_groups.append(group_info)
                
                print(f"  {i}. {group_info['group_name']}")
                print(f"     Scope: {scope_type} - {scope_name}")
                print(f"     Principal Name: {group_info['principal_name']}")
                
                group_count += 1
                
            except Exception as e:
                print(f"  ⚠️ Error processing group: {e}")
        
        all_user_data.append((email, display_name, user_groups))
        
        # Add a small delay to avoid rate limiting
        time.sleep(0.1)
    
    # Save to Excel if requested
    if save_to_excel_flag and all_user_data:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if email:
            filename = f"ado_groups_{email.split('@')[0]}_{timestamp}.xlsx"
        else:
            filename = f"ado_groups_all_users_{timestamp}.xlsx"
        
        saved_file, total_saved = save_to_excel(all_user_data, filename)
        print(f"\n💾 Results saved to: {saved_file}")
        print(f"📊 Total groups saved: {total_saved}")
    
    print(f"\n📊 Total users processed: {len(all_user_data)}")
    print(f"📊 Total groups found: {group_count}")
    
    return all_user_data

# ======================================================
# FULL AUDIT (All users) - Optimized version
# ======================================================
def full_audit():
    print("🔍 Starting Azure DevOps user group membership audit...\n")
    
    users = get_all_users()
    print(f"📊 Total users found: {len(users)}\n")
    
    all_user_data = []
    processed_count = 0
    
    for idx, user in enumerate(users, 1):
        email = user["principalName"]
        display_name = user.get("displayName", "N/A")
        
        # Progress indicator
        if idx % 10 == 0 or idx == 1:
            print(f"Processing user {idx}/{len(users)}: {email}")
        
        try:
            groups = get_user_groups(user["descriptor"])
            user_groups = []
            
            for group_descriptor in groups:
                try:
                    gd = get_group(group_descriptor)
                    
                    if gd.get("displayName") == "Security Service Group":
                        continue
                    
                    scope_type, scope_name = determine_scope(gd)
                    
                    group_info = {
                        "group_name": gd.get("displayName", "Unknown"),
                        "principal_name": gd.get("principalName", ""),
                        "scope_type": scope_type,
                        "scope_name": scope_name,
                        "description": gd.get("description", "")
                    }
                    user_groups.append(group_info)
                    
                except Exception as e:
                    print(f"  ⚠️ Error processing group for {email}: {e}")
            
            all_user_data.append((email, display_name, user_groups))
            processed_count += 1
            
        except Exception as e:
            print(f"  ⚠️ Failed to get groups for {email}: {e}")
            all_user_data.append((email, display_name, []))
        
        # Add delay to avoid rate limiting
        time.sleep(0.1)
    
    # Save to Excel
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"ado_full_audit_{timestamp}.xlsx"
    
    saved_file, total_groups = save_to_excel(all_user_data, filename)
    
    # Print summary
    print(f"\n" + "="*60)
    print("AUDIT COMPLETE")
    print("="*60)
    print(f"✅ File saved: {saved_file}")
    print(f"📊 Total users processed: {processed_count}/{len(users)}")
    print(f"📊 Total user-group records: {total_groups}")
    print(f"📊 File size: {os.path.getsize(saved_file) / 1024:.2f} KB")
    
    return saved_file

# ======================================================
# MAIN MENU
# ======================================================
def main():
    print("="*60)
    print("AZURE DEVOPS USER GROUP AUDIT TOOL")
    print("="*60)
    print("\nSelect an option:")
    print("1. 🔍 Search for a specific user's group memberships")
    print("2. 📊 Run full audit for all users")
    print("3. 📋 Get total user count")
    print("4. ❌ Exit")
    
    while True:
        try:
            choice = input("\nEnter your choice (1-4): ").strip()
            
            if choice == "1":
                email = input("Enter user email to search: ").strip()
                if not email:
                    print("Please enter a valid email address.")
                    continue
                
                print("\n" + "="*60)
                print(f"SEARCHING FOR: {email}")
                print("="*60)
                
                # Ask if user wants to save to Excel
                save_option = input("\nSave results to Excel file? (y/n): ").strip().lower()
                save_to_excel_flag = save_option == 'y'
                
                get_user_details_with_export(email, save_to_excel_flag)
                
                input("\nPress Enter to return to main menu...")
                print("\n" + "="*60)
                main()  # Return to main menu
                
            elif choice == "2":
                confirm = input("⚠️  This will audit ALL users (900+). Continue? (y/n): ").strip().lower()
                if confirm == 'y':
                    print("\n" + "="*60)
                    print("STARTING FULL AUDIT")
                    print("="*60)
                    
                    # Start the audit
                    start_time = time.time()
                    saved_file = full_audit()
                    elapsed_time = time.time() - start_time
                    
                    print(f"⏱️  Time taken: {elapsed_time:.2f} seconds")
                    print(f"📁 Output file: {os.path.abspath(saved_file)}")
                    
                    # Offer to open the file
                    open_file = input("\nOpen the Excel file? (y/n): ").strip().lower()
                    if open_file == 'y':
                        try:
                            os.startfile(saved_file)  # Windows
                        except:
                            try:
                                import subprocess
                                subprocess.call(['open', saved_file])  # macOS
                            except:
                                print(f"File saved at: {os.path.abspath(saved_file)}")
                    
                    input("\nPress Enter to return to main menu...")
                    print("\n" + "="*60)
                    main()  # Return to main menu
                else:
                    print("Audit cancelled.")
                    continue
                    
            elif choice == "3":
                print("\n📊 Fetching user count...")
                users = get_all_users()
                print(f"\nTotal users in organization: {len(users)}")
                input("\nPress Enter to return to main menu...")
                print("\n" + "="*60)
                main()  # Return to main menu
                
            elif choice == "4":
                print("👋 Exiting...")
                break
                
            else:
                print("Invalid choice. Please enter 1, 2, 3, or 4.")
                
        except KeyboardInterrupt:
            print("\n\n👋 Operation cancelled by user.")
            break
        except Exception as e:
            print(f"\n❌ An error occurred: {e}")
            import traceback
            traceback.print_exc()
            input("Press Enter to continue...")
            main()

# ======================================================
if __name__ == "__main__":
    main()
