"""
Owner: Armani Zuniga
Date Created: March 2024
Description: Python program that grabs Genius Bar metric information from two different Excel spreadsheets and
organizes the data into three new Excel spreadsheets for Monthly Tracker.
"""
import pandas as pd
import combine

if __name__ == "__main__":
    # Prompts the user to drag and drop the Excel files into the terminal window.
    excel_file_path1, excel_file_path2 = combine.get_excel_file_paths()

    # Read the WFH Sheet1 and Benchmark Prior Month into separate DataFrames
    WFM_Scheduled, Benchmark_prior_month = combine.get_dataframes(excel_file_path1, excel_file_path2)

    # Combine function extracts all the Genius Bar data from both files
    # Returns a list of class objects that holds all the Technicians information
    tech_instances = combine.combine(WFM_Scheduled, Benchmark_prior_month)

    # Create a dictionary to store all class object attribute values for new columns
    # This dictionary will be used to create the DataFrame will all class object information
    data = combine.class_to_dataframe(tech_instances)
    df = pd.DataFrame(data)

    # Export the Genius, Technical Expert, and Technical Specialist DataFrames to Excel
    combine.split_df(df)
    # Format each of the newly created Excel Files so that it is easier to read
    combine.adjust_excel()
