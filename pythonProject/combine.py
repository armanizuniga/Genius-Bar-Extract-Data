import pandas as pd
from Technician import Technician
from openpyxl import load_workbook


# Prompts the user to drag and drop the Excel files into the terminal.
# Processes the input paths: removes backslashes (\) and strips trailing spaces.
# Returns cleaned file paths for both Excel files.
def get_excel_file_paths():
    # Path to the Excel file ---------Adjust to file name-------------
    # ********** maybe take as an input by clicking and dragging into terminal ********
    excel_file_path1 = input("Drag and drop WFM excel sheet, then press Enter:\n")
    excel_file_path2 = input("Drag and drop Benchmark excel sheet, then press Enter:\n")
    # Using replace() method
    excel_file_path1 = excel_file_path1.replace("\\", "")
    excel_file_path2 = excel_file_path2.replace("\\", "")
    # Using strip method to get rid of trailing space
    excel_file_path1 = excel_file_path1.strip()
    excel_file_path2 = excel_file_path2.strip()
    return excel_file_path1, excel_file_path2


# Function that returns two Dataframes from the filepaths provided from User
def get_dataframes(excel_file_path1, excel_file_path2):
    try:
        # Read all Excel sheets into a dictionary of DataFrames -------Adjust to Sheet you want (PM, QTD)----
        wfm_scheduled = pd.read_excel(excel_file_path1, sheet_name='Sheet1')
        benchmark_prior_month = pd.read_excel(excel_file_path2, sheet_name='PM', skiprows=8)
        print("Processed both Excel sheets successfully.")

    except FileNotFoundError:
        print("File not found. Please provide a valid file path.")
    except Exception as e:
        print("An error occurred:", e)

    return wfm_scheduled, benchmark_prior_month


# Function that takes two parameters WFH and Benchmarks DataFrames from the Excel files
# It filters both DataFrames to create new DataFrames with relevant Genius Bar information
# Loops through for all Technicians to store data into class objects and return a list of objects
def combine(zoning, benchmarks):
    # Filter the zoning DataFrame to extract Genius Bar technician information only
    # List of the job titles we want which is used to filter the zoning data frame
    job_titles = ['Genius', 'Technical Expert', 'Technical Specialist']
    zoning_info = zoning[zoning['Merlin Job Title'].isin(job_titles)]
    date_df = zoning_info['Date']
    date_string = date_df.iloc[0].strftime('%Y-%m-%d %H:%M:%S')  # Convert to string
    year_value = int(date_string[:4])  # Extract the first four characters for the year
    month_value = int(date_string[5:7])  # Extract characters 6 and 7 for the month

    # Create a list without duplicates from the zoning DataFrame column
    unique_names = zoning_info['Employee Name'].drop_duplicates().tolist()
    # Sort the list of names alphabetically
    unique_names = sorted(unique_names)
    # print(unique_names)

    # Filter the benchmarks DataFrame to extract relevant Genius Bar information only
    # Find the row position of the row containing 'Section : Genius & Forum', all info below is related to GB
    start_row_position = benchmarks[benchmarks['Name'] == 'Section : Genius & Forum'].index[0]
    # Create a new DataFrame containing all rows below the 'Section : Genius & Forum' row
    # and including only the first 10 columns from the original DataFrame
    df_clean_table = benchmarks.iloc[start_row_position + 2:, :10]
    # Updating the column list names used for df_clean_table DataFrame
    df_clean_table.columns = [
        'Not Sure',
        'Name',
        'Job',
        'Customers Helped in Today at Apple Sessions',
        'Today at Apple Sessions Delivered',
        'Customers Helped',
        'Mac Queue Sessions Delivered',
        'Mac Queue Avg Session Duration',
        'Mobile Queue Sessions Delivered',
        'Mobile Queue Avg Session Duration'
    ]
    # This continues to filter the benchmarks -> df_clean_ DataFrame based on 'Job' column
    filtered_df = df_clean_table[df_clean_table['Job'].isin(['Genius', 'Technical Expert', 'Technical Specialist'])]
    # List for selecting specific columns from the benchmarks DataFrame
    selected_columns = ['Name', 'Job', 'Customers Helped',
                        'Mac Queue Sessions Delivered', 'Mac Queue Avg Session Duration',
                        'Mobile Queue Sessions Delivered', 'Mobile Queue Avg Session Duration']
    # Create the final benchmarks DataFrame with filtered rows and selected columns
    new_df = filtered_df[selected_columns]

    # Business intro DataFrame
    df_business_intro = benchmarks.iloc[1:start_row_position, :7]
    df_business_intro.columns = [
        'Not Sure',
        'Name',
        'Job',
        'Total Sales($)',
        'Total Sales % to store',
        'Business sales',
        'Business Introductions'
    ]
    # List of the job titles we want which is used to filter the zoning data frame
    job_titles = ['Genius', 'Technical Expert', 'Technical Specialist']
    df_business_intro = df_business_intro[df_business_intro['Job'].isin(job_titles)]

    # This List is used to store all class object technicians
    list_of_objects = []

    # Loop through the zoning DataFrame for all Genius Bar Technicians from list of unique names
    # Because there are multiple instances of the same Technician, this will calculate the sum
    # For each column and creates a class object for each Technician
    for name in unique_names:
        temp_df = zoning_info[zoning_info['Employee Name'] == name]

        # Get the index of the row where 'Employee Name' matches the given name
        row_index = temp_df.index[temp_df['Employee Name'] == name].tolist()
        # Assuming 'name' is the name of the employee and 'index_val' is the index of the row
        role = temp_df.loc[row_index[0], 'Merlin Job Title']
        # Assuming 'name' is the name of the employee and 'index_val' is the index of the row
        worker_type = temp_df.loc[row_index[0], 'Worker Type']

        # Calculate the sum of values in the specific column
        mobile_support_sum = temp_df['Mobile Support'].sum()
        mac_support_sum = temp_df['Mac Support'].sum()
        daily_download = temp_df['Daily Download'].sum()
        repair_pickup = temp_df['Repair Pick-Up'].sum()
        guided = temp_df['Guided'].sum()
        gb_on_point = temp_df['GB On Point'].sum()
        connection = temp_df['Connection'].sum()
        iphone_repair = temp_df['iPhone Repair'].sum()
        mac_repair = temp_df['Mac Repair'].sum()
        total = temp_df['Total'].sum()

        # Extract first and last name
        fix_name = name.split(',')
        last_name = fix_name[0].strip()
        first_name = fix_name[1].strip()

        # update class object attributes the sum
        tech_instance = Technician(name=first_name + " " + last_name, job=role, job_type=worker_type,
                                   mobile_support_hours=mobile_support_sum, mac_support_hours=mac_support_sum,
                                   iphone_repair_hours=iphone_repair, mac_repair_hours=mac_repair,
                                   repair_pickup_hours=repair_pickup, gb_on_point_hours=gb_on_point,
                                   daily_download_hours=daily_download, guided_hours=guided,
                                   connection_hours=connection, total_hours=total, year=year_value, month=month_value)
        # Append the tech_instance object to the list
        list_of_objects.append(tech_instance)

    # Iterate over each row in the benchmarks DataFrame
    for index, row in new_df.iterrows():
        # Access individual elements by column name
        name = row['Name'].strip()
        customers_helped = row['Customers Helped']
        mac_queue_sessions_delivered = row['Mac Queue Sessions Delivered']
        mac_queue_avg_session_duration = row['Mac Queue Avg Session Duration']
        mobile_queue_sessions_delivered = row['Mobile Queue Sessions Delivered']
        mobile_queue_avg_session_duration = row['Mobile Queue Avg Session Duration']

        # Iterate over the list of tech instances
        for tech_instance in list_of_objects:
            # Check if the name matches the name of the tech instance you want to update
            if tech_instance.name == name:
                # Update the desired attribute of the tech instance
                tech_instance.customers_helped = customers_helped
                # Calculate SPQH (Sessions Per Queued Hour) with a conditional check to avoid divide by zero error
                mobile = tech_instance.mobile_support_hours
                mac = tech_instance.mac_support_hours

                # Avoid divide by zero error
                tech_instance.spqh = round(customers_helped / (mobile + mac) if (mobile + mac) != 0 else 0, 2)
                tech_instance.mac_queue_sessions_delivered = mac_queue_sessions_delivered
                tech_instance.mac_queue_avg_session_duration = mac_queue_avg_session_duration
                tech_instance.mobile_queue_sessions_delivered = mobile_queue_sessions_delivered
                tech_instance.mobile_queue_avg_session_duration = mobile_queue_avg_session_duration

    # Go through df_business_intro DataFrame to update technician object
    for index, row in df_business_intro.iterrows():
        # Find the Technician object with the same name
        for technician in list_of_objects:
            technician_name = row['Name']
            business_intro_value = row['Business Introductions']
            if technician.name == technician_name:
                # Update the business_intro attribute
                technician.business_intro = business_intro_value
                break  # Break the loop once the object is found

    return list_of_objects


# Function takes in a list of class objects and stores into a Dictionary with customer columns
# For a new dataframe. Returns a dictionary that will all Genius Bar information
def class_to_dataframe(tech_instances):
    data = {
        "Month": [tech_instance.month for tech_instance in tech_instances],
        "Year": [tech_instance.year for tech_instance in tech_instances],
        "Job": [tech_instance.job for tech_instance in tech_instances],
        "Type": [tech_instance.job_type for tech_instance in tech_instances],
        "Name": [tech_instance.name for tech_instance in tech_instances],
        "SPQH": [tech_instance.spqh for tech_instance in tech_instances],
        "Customers Helped": [tech_instance.customers_helped for tech_instance in tech_instances],
        "Mac Duration": [tech_instance.mac_queue_avg_session_duration for tech_instance in tech_instances],
        "Mobile Duration": [tech_instance.mobile_queue_avg_session_duration for tech_instance in tech_instances],
        "NPS": [tech_instance.nps for tech_instance in tech_instances],
        "TMS": [tech_instance.tms for tech_instance in tech_instances],
        "SUR": [tech_instance.sur for tech_instance in tech_instances],
        "Business Intros": [tech_instance.business_intro for tech_instance in tech_instances],
        "Discussed AppleCare": [tech_instance.apple_care for tech_instance in tech_instances],
        "Offered Trade In": [tech_instance.trade_in for tech_instance in tech_instances],
        "Survey Qty": [tech_instance.survey for tech_instance in tech_instances],
        "Full Survey Qty": [tech_instance.full_survey for tech_instance in tech_instances],
        "Opportunities": [tech_instance.opportunities for tech_instance in tech_instances],
        "Mac Sessions": [tech_instance.mac_queue_sessions_delivered for tech_instance in tech_instances],
        "Mobile Sessions": [tech_instance.mobile_queue_sessions_delivered for tech_instance in tech_instances],
        "Mobile Support": [tech_instance.mobile_support_hours for tech_instance in tech_instances],
        "Mac Support": [tech_instance.mac_support_hours for tech_instance in tech_instances],
        "iPhone Repair": [tech_instance.iphone_repair_hours for tech_instance in tech_instances],
        "Mac Repair": [tech_instance.mac_repair_hours for tech_instance in tech_instances],
        "Repair Pickup": [tech_instance.repair_pickup_hours for tech_instance in tech_instances],
        "GB On Point": [tech_instance.gb_on_point_hours for tech_instance in tech_instances],
        "Daily Download": [tech_instance.daily_download_hours for tech_instance in tech_instances],
        "Guided": [tech_instance.guided_hours for tech_instance in tech_instances],
        "Connection": [tech_instance.connection_hours for tech_instance in tech_instances],
        "Total Hours": [tech_instance.total_hours for tech_instance in tech_instances]
        # Add more columns as needed
    }
    return data


# Function that takes the main dataframe from the list of class objects and separates them by job title
# Exports them as three different Excel Spreadsheets
def split_df(df):
    # df is your new DataFrame containing all the Genius Bar data
    # Step 1: Filter rows by their Job title and create new DataFrame
    genius_df = df[df['Job'] == 'Genius']
    expert_df = df[df['Job'] == 'Technical Expert']
    specialist_df = df[df['Job'] == 'Technical Specialist']

    # Step 2: Sort the concatenated DataFrames by the "Name" column
    genius_sorted = genius_df.sort_values(by='Name')
    expert_sorted = expert_df.sort_values(by='Name')
    specialist_sorted = specialist_df.sort_values(by='Name')

    # file_path where you want to save the Excel file
    genius_path = "~/Desktop/Genius.xlsx"
    expert_path = "~/Desktop/Technical Expert.xlsx"
    specialist_path = "~/Desktop/Technical Specialist.xlsx"

    # Step 3: Write each DataFrame to a separate Excel file
    genius_sorted.to_excel(genius_path, index=False)  # Set index=False to exclude row numbers in the Excel file
    expert_sorted.to_excel(expert_path, index=False)
    specialist_sorted.to_excel(specialist_path, index=False)

    # Step 4: Print a report to the console about some stats
    print("Total Customers Helped:                   ", df['Customers Helped'].sum())
    print("Total Mac Sessions:                       ", df['Mac Sessions'].sum())
    print("Total Mobile Sessions:                    ", df['Mobile Sessions'].sum())
    print("Amount of Genius Employees:               ", len(genius_df))
    print("Amount of Technical Expert Employees:     ", len(expert_df))
    print("Amount of Technical Specialist Employees: ", len(specialist_df))
    print(f"Excel sheets successfully saved to Desktop")


# Function that opens the 3 newly created Excel sheets and formats cell size for better readability
# ********************** NEEDS TO BE ADJUSTED TO USER *********************************
def adjust_excel():
    file_list = ["/Users/armanizuniga/Desktop/Genius.xlsx", "/Users/armanizuniga/Desktop/Technical Expert.xlsx",
                 "/Users/armanizuniga/Desktop/Technical Specialist.xlsx"]

    for file_path1 in file_list:
        wb = load_workbook(file_path1)
        # Get the active sheet
        ws = wb.active
        # Iterate over all columns and adjust the column widths based on the maximum length of content in each column
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column name
            for cell in col:
                try:  # Avoid error if cell value is None
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2  # Adjusted width with padding
            ws.column_dimensions[column].width = adjusted_width  # Set column width

        # Save the changes to the Excel file
        wb.save(file_path1)

    print(f"DataFrame successfully adjusted")
