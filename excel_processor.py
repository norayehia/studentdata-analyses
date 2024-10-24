import pandas as pd

def process_excel(file_path):
    # Get user input for the semester, academic year, password, and role preference
    semester = input("Enter the semester (e.g., 'Fall', 'Spring'): ")
    academic_year = input("Enter the academic year (e.g., '2024-2025'): ")
    split_names = input("Do you want to split the name into first and last name? (yes/no): ").strip().lower()
    custom_password = input("Enter the password (or press Enter for default 'ABC@123456'): ").strip()
    custom_role = input("Enter the role (or press Enter for default 'student'): ").strip()

    # Set default password and role if not provided
    if not custom_password:
        custom_password = 'ABC@123456'
    if not custom_role:
        custom_role = 'student'

    # Load the Excel file
    df = pd.read_excel(file_path)

    # Rename columns
    df.columns = df.columns.str.replace('Field6', 'FacultyName')
    df.columns = df.columns.str.replace('Field8', 'Department')

    # Filter for entries from the Faculty of Computer Sci & Eng
    filtered_df = df[df['FacultyName'] == 'Faculty of Computer Sci & Eng']

    # Ensure that 'Name', 'Subject code', and 'Catalog number' columns exist
    if 'Name' not in filtered_df.columns or 'Subject code' not in filtered_df.columns or 'Catalog number' not in filtered_df.columns:
        raise KeyError("'Name', 'Subject code', or 'Catalog number' column missing in the data.")

    # Create a new 'course-code' column by concatenating the required columns
    filtered_df['course-code'] = (
        filtered_df['Subject code'].astype(str) +
        filtered_df['Catalog number'].astype(str) +
        semester + academic_year
    )

    # Add the custom or default password and role to the DataFrame
    filtered_df['password'] = custom_password
    filtered_df['role'] = custom_role

    if split_names == 'yes':
        # Function to adjust the name and split it into 'First Name' and 'Last Name'
        def split_name(name):
            parts = name.split(maxsplit=1)  # Split into two parts: first name and the rest
            first_name = parts[0] if len(parts) > 0 else ''
            last_name = parts[1].replace(' ', '') if len(parts) > 1 else ''  # Remove all remaining spaces in the last name
            return pd.Series([first_name, last_name], index=['First Name', 'Last Name'])

        # Apply the function to create 'First Name' and 'Last Name' columns
        filtered_df[['First Name', 'Last Name']] = filtered_df['Name'].apply(split_name)

        # Select columns for the first output with split names
        output_df1 = filtered_df[['ID', 'First Name', 'Last Name', 'Email', 'password', 'course-code', 'role']]
    else:
        # If not splitting, leave the 'Name' column as is and use it in the output
        output_df1 = filtered_df[['ID', 'Name', 'Email', 'password', 'course-code', 'role']]

    # Save the first output file
    output_df1.to_excel('subjectn.xlsx', index=False, sheet_name='Sheet1')
    print("File 'subjectn.xlsx' created successfully.")

    # Function to determine the course level based on the first digit of the Catalog number
    def determine_level(catalog_number):
        catalog_number_str = str(catalog_number).strip()
        if not catalog_number_str.isdigit():
            return None  # Return None if not a valid number
        level_digit = catalog_number_str[0]  # Get the first character
        if level_digit == '0':
            return 'level1'
        elif level_digit == '1':
            return 'level2'
        elif level_digit == '2':
            return 'level3'
        elif level_digit == '3':
            return 'level4'
        else:
            return None

    # Apply the level determination function
    filtered_df['category-level'] = filtered_df['Catalog number'].apply(determine_level)

    # Select and rename columns for the second output
    output_df2 = filtered_df[['course-code', 'course Title', 'category-level']]
    output_df2.columns = ['course', 'description', 'category-level']
    output_df2.to_excel('Course_info.xlsx', index=False, sheet_name='Sheet1')
    print("File 'Course_info.xlsx' created successfully.")

    return output_df1, output_df2
