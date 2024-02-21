
# constants.py

#User Input Control
COLUMNS_TO_CONTROL_ODOO=['Employee Name', 'Company/Display Name', 'Job Position/Display Name', 'Hr Team/Display Name', 'Work Location/Display Name', 'Nomenclature Hadia','Job Family']
COLUMNS_TO_CONTROL_NOMEMCLATURE=['WhosWho Level', 'HC team breakdown']
#Define the common column Name
COMMON_COLUMN_NAME = 'Nomenclature Hadia'



# File paths
TEST_PICS_PATH = './input_data/new_empl_pics'
OUTPUT_DATA_PATH = './output_data/issues_with_pics_2.xlsx'


INPUT_TEMPLATE_PATH = "./input_data/template_whoswho_2023.pptx"
INPUT_EXCEL_PATH = "./input_data/whoswho_main_input_29112023_5.xlsx"
OUTPUT_PATH_FORMAT = "./output_data/auto_generated_whoswho_{date}.pptx"


# Constants for column name mapping
NAME_MAPPING = {
    'Employee Name': 'full_name',
    'Company/Display Name': 'company_name',
    'Job Position/Display Name': 'old_job_position',
    'Hr Team/Display Name': 'hr_team',
    'Work Location/Display Name': 'location',
    'Job Level':'job_level',
    'Nomenclature Hadia':'job_position',
    'Job Family':'job_family',
    'Image': 'image_path'
}

# Constants for HR teams
HR_TEAMS = ["McK Team", "Office Admin team", "IK Team", "BCG ME Team", "Business Translation Team", "IT Team",
            "Finance Team", "BCG Europe Team", "Graphic Design Team", "BCG Africa Team", "Marketing team",
            "BCG KT support", "Operations Team", "Data Analytics Services", "Business Development Team",
            "HR team", "Research Team", "Executive Management", "IKT/Finance"]

# Constants for job families
JOB_FAMILIES = ['Business Research','Office Management','Business Translation','IT','Finance','Graphic Design','Marketing',
                'Operations','Data Analytics','Business Development','Human Capital','Executive Management','Sales Representative']

# Define the locations for which to store "location.png"
DESIRED_LOCATIONS = ['Morocco', 'Egypt', 'Mexico', 'UAE', 'Spain']
