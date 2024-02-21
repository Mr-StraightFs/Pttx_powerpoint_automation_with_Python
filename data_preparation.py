from constants import NAME_MAPPING, HR_TEAMS, JOB_FAMILIES,DESIRED_LOCATIONS
import pandas as pd
import streamlit as st
import time


def control_columns_in_input(dataframe, columns_to_control):
    # Check if all specified columns exist in the dataframe

    missing_columns = [col for col in columns_to_control if col not in dataframe.columns]
    empty_columns = [col for col in columns_to_control if dataframe[col].dropna().empty]

    if missing_columns:
        missing_cols_message = f"This input file MUST contain the following columns: {', '.join(missing_columns)}. Please Try again."
        #st.error(missing_cols_message)  # Display the error message to the user
        raise ValueError(missing_cols_message)
        # Check if any of the specified columns are empty (contain only NaN values)

    

    if empty_columns:
        empty_cols_message = f" The Following Columns are empty: {', '.join(empty_columns)}"
        st.error(empty_cols_message)  # Display the error message to the user
        raise ValueError(empty_cols_message)  # Raise the ValueError to stop code execution

    return dataframe

def merge_and_clean_dataframes(df1: pd.DataFrame, df2: pd.DataFrame, merge_column: str) -> pd.DataFrame:
    """
    Merge two dataframes based on a common column and perform data cleaning.

    Args:
        df1 (pd.DataFrame): The first dataframe to be merged.
        df2 (pd.DataFrame): The second dataframe to be merged.
        merge_column (str): The name of the common column used for merging.

    Returns:
        pd.DataFrame: The cleaned and merged dataframe.
    """

    # Clean the input dataframes
    cleaned_df1 = clean_dataframe(df1)
    cleaned_df2 = clean_dataframe(df2)
    # Merge the two dataframes
    merged_df = pd.merge(cleaned_df1, cleaned_df2, on=merge_column, how='left')
    return merged_df

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean a dataframe by removing duplicates, handling missing values, and cleaning text columns.

    Args:
        df (pd.DataFrame): The dataframe to be cleaned.

    Returns:
        pd.DataFrame: The cleaned dataframe.
    """
    # Automatically detect text columns and clean them
    text_columns = [column for column in df.columns if df[column].dtype == 'object']

    # Remove leading and trailing spaces and replace multiple spaces with a single space in text columns
    for column in text_columns:
        df[column] = df[column].str.strip()
        df[column] = df[column].str.replace(r'\s+', ' ', regex=True)

    # Drop duplicate rows
    df = df.drop_duplicates()

    # Drop rows with missing values
    df = df.dropna()

    # Reset the index
    cleaned_df = df.reset_index(drop=True)

    # Add additional cleaning or transformations here if needed

    return cleaned_df


def data_preparation_processor(raw_df):

    # Clean the 'Employee Name' column
    raw_df['Employee Name'] = raw_df['Employee Name'].str.strip()  # Remove leading and trailing spaces
    raw_df['Employee Name'] = raw_df['Employee Name'].str.replace(r'\s+', ' ',
                                                                  regex=True)  # Replace multiple spaces with a single space



    # Rename the columns using the mapping
    raw_df = raw_df.rename(columns=NAME_MAPPING)

    # Fill missing values with an empty string
    raw_df.fillna('', inplace=True)

    # Add the "location_flag" column using apply() and lambda function
    raw_df['location_flag'] = raw_df['location'].apply(lambda x: 'location.png' if x in DESIRED_LOCATIONS else '')

    # Replace NaN values with an empty string
    raw_df['job_position'] = raw_df['job_position'].fillna('').astype(str)
    # Create a new column 'employee_pic' by formatting 'full_name'
    raw_df['employee_pic'] = 'ipp_' + raw_df['full_name'].str.replace(' ', '_').str.lower() + '.png'
    # Create a new column 'employee_missing_pic' and assign 0 to all rows
    raw_df['employee_missing_pic'] = 0
    # Create a new column 'comment' and instantiate it at 'OK'
    raw_df['comment'] = 'OK'

    # Sort by location and full name
    #raw_df = raw_df.sort_values(by=['location', 'full_name'])

    # Convert job_position column to lowercase and remove leading/trailing spaces

    # Add columns for different job position patterns
    raw_df['Is_partner'] = raw_df['job_position'].str.contains('partner', case=False).astype(int)
    raw_df['Is_Head'] = raw_df['job_position'].str.contains('head', case=False).astype(int)
    raw_df['Is_officer'] = raw_df['job_position'].str.contains('officer', case=False).astype(int)
    raw_df['Is_manager'] = raw_df['job_position'].str.contains('manager', case=False).astype(int)
    raw_df['Is_country_manager'] = raw_df['job_position'].str.contains('country manager', case=False).astype(int)
    raw_df['Is_team_lead'] = raw_df['job_position'].str.contains('team lead|team leader', case=False, regex=True).astype(int)
    raw_df['Is_vice_president'] = raw_df['job_position'].str.contains('vice president|vp', case=False, regex=True).astype(int)
    raw_df['Is_acc_exec'] = raw_df['job_position'].str.contains('account executive', case=False).astype(int)
    raw_df['Is_associate'] = raw_df['job_position'].str.contains('associate', case=False).astype(int)
    raw_df['Is_admin'] = raw_df['job_position'].str.contains('admin|administrator', case=False, regex=True).astype(int)
    raw_df['Is_lead'] = raw_df['job_position'].str.contains('lead', case=False).astype(int)
    raw_df['Is_senior'] = raw_df['job_position'].str.contains('senior|Sr.', case=False).astype(int)
    raw_df['Is_director'] = raw_df['job_position'].str.contains('director', case=False).astype(int)
    raw_df['Is_coordinator'] = raw_df['job_position'].str.contains('coordinator', case=False).astype(int)
    raw_df['Is_designer'] = raw_df['job_position'].str.contains('designer', case=False).astype(int)
    raw_df['Is_junior'] = raw_df['job_position'].str.contains('junior', case=False).astype(int)
    raw_df['Is_analyst'] = raw_df['job_position'].str.contains('analyst', case=False).astype(int)
    raw_df['Is_COO'] = raw_df['job_position'].str.contains(r'\bCOO\b|\bChief Operations Officer\b', case=False,regex=True).astype(int)
    raw_df['Is_BR_leader'] = (raw_df['Is_lead'] | raw_df['Is_manager'] | raw_df['Is_Head']).astype(int)
    raw_df['Is_HR_Administration'] = raw_df['HC team breakdown'].str.contains('HR Administration', case=False).astype(int)
    raw_df['Is_Executive_Management'] = raw_df['HC team breakdown'].str.contains('Executive Management', case=False).astype(int)
    raw_df['Is_Recruitment'] = raw_df['HC team breakdown'].str.contains('Recruitment', case=False).astype(int)
    raw_df['Is_PD'] = raw_df['HC team breakdown'].str.contains('PD', case=False).astype(int)

    #raw_df['Is_cleaning'] = raw_df['job_position'].str.contains('cleaning', case=False).astype(int)

    # Drop rows where Is_cleaning is 1 from raw_df
    #raw_df.drop(raw_df[raw_df['Is_cleaning'] == 1].index, inplace=True)

    #raw_df.to_excel(r'.\test_data\processed_empl_data.xlsx')

    # Create separate dataframes for each HR team
    team_dfs = [raw_df[raw_df['hr_team'] == team] for team in HR_TEAMS]

    # Create separate dataframes for each job family
    job_family_dfs = [raw_df[raw_df['job_family'] == family] for family in JOB_FAMILIES]

    return team_dfs, job_family_dfs , raw_df








