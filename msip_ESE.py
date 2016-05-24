#!/depot/Python-3.5.0/bin/python

# ------------------Import external library/commands---------------#

from __future__ import print_function
from abc import ABCMeta, abstractmethod
import os
import sys
from subprocess import Popen
from subprocess import PIPE
import tarfile
import shutil
import time
import datetime
from xlrd import open_workbook as read_excel_module
from xlrd import XLRDError

__author__ = 'vlad'

"""
USAGE:  program
        EXAMPLE: msip_ESE.py

DESCRIPTION:
        The script is running ESE GUI for executing its flow.
        ESE - Extraction and Simulation Evaluation flow is for evaluating the CCS/PCS setup updates' impact in simulation results.

FOR SUPPORT(BUG/ENHANCEMENT):
        Please send e-mail to "vlad@synopsys.com"

AUTHOR:
        Vladimir Danielyan

ALL CLASSES
    MsipEse:    The main class of the script's "RUN" execution part
    MsipEseQa:  The class of the script's results "QA" part


ALL FUNCTIONS:
    open_file_for_writing(file_path, writing_file_name)
    open_file_for_reading(file_path, reading_file_name)
    check_for_file_existence(path_to_check, item_to_check)
    check_for_dir_existence(path_to_check, item_to_check)
    print_to_stdout(class_object_name, text_to_print)
    print_to_stderr(class_object_name, text_to_print)
    get_class_name(class_object)
    get_index_of_list(list_name, list_item_name)
    get_list_length(list_name)

"""

# --------------------------------------------------- #
# ---------------- Global Variables ----------------- #
# --------------------------------------------------- #


# The script environment directories list
environment_directories_name_list = ["LOGS",        # Index[0] Logs directory name
                                     "REPORTS",     # Index[1] Reports directory name
                                     "RESULTS",     # Index[2] Results directory name
                                     "RUN_DIR",     # Index[3] Run directory name
                                     "SCRIPTS",     # Index[4] Scripts directory name
                                     "TESTCASES",   # Index[5] Test cases directory name
                                     "DATA"         # Index[6] Internal data directory name. DATA/ [PEX_SAMPLE_RUN_SCRIPTS, SAMPLE_OA_LIBRARIES, SIM_SAMPLE_RUN_SCRIPTS]
                                     ]

# Available Options For the Script
available_script_options = ["-excelFile",                   # Index[0] Excel file
                            "-targetProjectName",           # Index[1] Target Project Name
                            "-targetProjectRelease",        # Index[2] Target Project Release
                            "-referenceProjectName",        # Index[3] Reference Project Name
                            "-referenceProjectRelease",     # Index[4] Reference Project Release
                            "-runDirectory",                # Index[5] Script Run Directory
                            "-executedTestCasePackage",     # Index[6] Executed test case package(s)
                            "-projectsRootDirectory"        # Index[7] Projects root directory path
                            ]

# Available excel parameters. NOTE!!! If the list value changed please make appropriate change in ReadExcel class for get_* functions
# Important do not make any change in list order, as the script recognised the values by exact index. If there is need to do modification please update
# available_excel_options variable in the script file
available_excel_options = ["Test Case Name",
                           "Date",
                           "2 Contact Persons (email)",
                           "CCS or PCS Name",
                           "Release",
                           "Test Case Package Path",
                           "Test bench",
                           "GDS file(s)",
                           "LVS Netlist file(s)",
                           "Extract Type For GDS file(s)",
                           "Simulation options",
                           "Measure file(s)",
                           "Other Include(s)",
                           "Measure results",
                           "Extract Netlist(s) - .spf file(s)",
                           "Target CCS/PCS",
                           "Target Release Version",
                           "Reference CCS/PCS",
                           "Reference Release Version",
                           "Measured variables",
                           "Comments for criteria",
                           "Target LVS Tool Name",
                           "Target LVS Tool Version",
                           "Reference LVS Tool Name",
                           "Reference LVS Tool Version",
                           "Target LVS deck",
                           "Reference LVS deck",
                           "Target LVS options/sourceme",
                           "Reference LVS options/sourceme",
                           "Target RCXT version",
                           "Reference RCXT version",
                           "Target RCXT deck",
                           "Reference RCXT deck",
                           "Target RCXT starcmd",
                           "Reference RCXT starcmd",
                           "Target Simulation Tool Name",
                           "Target Simulation Tool Version",
                           "Reference Simulation Tool Name",
                           "Reference Simulation Tool Version",
                           "Other Comments"]


# --------------------------------------------------- #
# -------------------- Functions -------------------- #
# --------------------------------------------------- #


def print_description(violated_case):
    """
    The script is printing description
    :param violated_case: String why the
    :return:
    """

    # Description
    description = """
        USAGE:  program <OPTION(S)>
                EXAMPLE: msip_ESE.py -excelFile ./x649_bias_check_test_case.xlsx

        OPTIONS:
                {0}

        DESCRIPTION:
                The script is running ESE GUI for executing its flow.
                ESE - Extraction and Simulation Evaluation flow is for evaluating the CCS/PCS setup updates' impact in simulation results.

        FOR SUPPORT(BUG/ENHANCEMENT):
                Please send e-mail to "vlad@synopsys.com"

        AUTHOR:
                Vladimir Danielyan
        """.format(get_script_options_string())

    print("\n\n\t" + str(violated_case) + "\n\n")
    exit(description)


def get_current_time():
    """
    The function is returning time in string format
    :return:
    """

    current_time = time.time()
    current_date_time = datetime.datetime.fromtimestamp(current_time).strftime('%m/%d %H:%M:%S')
    return str(current_date_time)


def get_latest_release_version(releases_list):
    """
    The function is returning latest release version from input list of releases
    :param releases_list:
    :return:
    """

    latest_release = ""

    for release_version in releases_list:
        if release_version > latest_release:
            latest_release = release_version

    if get_string_length(latest_release) > 0:
        return latest_release
    else:
        return None


def get_script_options_string():
    """
    The function is returning script options by string for description section
    :return:
    """

    global available_script_options

    final_string = ""
    for option_name in available_script_options:
        final_string += str(option_name) + "\n\t\t"

    return final_string


# noinspection PyUnboundLocalVariable
def untar_zip_package(zip_file, path_to_extract):
    """
    The function is un taring tar.gz file
    :param path_to_extract:
    :param zip_file:
    :return:
    """

    if check_for_file_existence(get_file_path(zip_file), get_file_name_from_path(zip_file)):
        try:
            tar_file_object = tarfile.open(zip_file)
        except IOError:
            exit("ERROR!: Cannot extract .tar.gz file\n\t" + zip_file)

        try:
            tar_file_members = tar_file_object.getmembers()
        except tarfile.TarError:
            exit("ERROR!: Cannot extract .tar.gz file\n\t" + zip_file)

        for member in tar_file_members:
            member = str(member).split("<TarInfo \'")[1].split("'")[0]
            try:
                tar_file_object.extract(member, path=path_to_extract)
            except tarfile.ExtractError:
                exit("ERROR!: Cannot extract .tar.gz file\n\t" + zip_file)

        tar_file_object.close()
    else:
        exit("ERROR!:\tCannot find zip file:\t" + str(zip_file))


def get_directory_items_list(directory_path):
    """
    The function is returning content
    :param directory_path:
    :return:
    """

    if check_for_dir_existence(get_file_path(directory_path), get_file_name_from_path(directory_path)):
        return os.listdir(directory_path)
    else:
        return []


def clean_directories(class_object_name, directory_path):
    """
    The function is cleaning all unnecessary files from the selected directory
    :param class_object_name:
    :param directory_path:
    :return:
    """

    item_names_to_be_removed = ["UNTAR"]

    print_to_stdout(class_object_name, "Cleaning directory:\n\t" + directory_path)

    for root_path, dirs, files in os.walk(directory_path):
        for directory_name in dirs:
            for item_name_to_remove in item_names_to_be_removed:
                if directory_name.upper() == item_name_to_remove:
                    path_to_remove = os.path.join(root_path, directory_name)
                    print_to_stdout(class_object_name, "\tRemoving Directory\t" + path_to_remove)
                    shutil.rmtree(path_to_remove)

    print_to_stdout(class_object_name, "Cleaning process completed successfully" + directory_path)


def execute_external_command(command):
    """
    The function is executing process through Popen function
    :param command:
    :return:
    """

    return Popen(command, shell=True, stdout=PIPE, stderr=PIPE)


# noinspection PyUnboundLocalVariable
def open_file_for_writing(file_path, writing_file_name):
    """
    The function is generating write+ file object in the mentioned path
    :param writing_file_name: Input file name which need to been created
    :param file_path: Input file_path where to create file
    :return:
    """

    try:
        file_object = open(os.path.join(file_path, writing_file_name), mode="w+")
    except IOError:
        exit("ERROR: Cannot create file:\n\t" + str(os.path.join(file_path, writing_file_name) + "\nScript Finished with error.\n"))

    try:
        os.chmod(os.path.join(file_path, writing_file_name), mode=0o777)
    except OSError:
        return file_object

    return file_object


def open_file_for_reading(file_path, reading_file_name):
    """
    The function is returning read file object of the mentioned path
    :param file_path:
    :param reading_file_name:
    :return:
    """

    try:
        return open(os.path.join(file_path, reading_file_name), mode="r")
    except IOError:
        exit("ERROR: Cannot read file:\n\t" + str(os.path.join(file_path, reading_file_name) + "\nScript Finished with error.\n"))


def check_for_file_existence(path_to_check, item_to_check):
    """
    The function is checking the selected file existence in the selected path
    :param path_to_check: Path where to check item existence
    :param item_to_check: Item name to check for existence
    :return: True if exist or False if not
    """

    if os.path.isfile(os.path.join(path_to_check, item_to_check)):
        return True
    else:
        return False


def check_for_dir_existence(path_to_check, item_to_check):
    """
    The function is checking the selected directory existence in the selected path
    :param path_to_check: Path where to check item existence
    :param item_to_check: Item name to check for existence
    :return: True if exist or False if not
    """

    if os.path.isdir(os.path.join(path_to_check, item_to_check)):
        return True
    else:
        return False


def get_file_path(full_path_to_the_file):
    """
    The function is returning path to the file location
    :param full_path_to_the_file:
    :return:
    """

    return os.path.dirname(os.path.abspath(full_path_to_the_file))


def get_file_name_from_path(full_path_to_the_file):
    """
    The function is returning file name from the full path
    :param full_path_to_the_file:
    :return:
    """

    return str(os.path.basename(full_path_to_the_file))


def print_to_stdout(class_object_name, text_to_print):
    """
    The function is printing report in STDOUT file
    :param class_object_name: The object name
    :param text_to_print: The input text/digital value
    :return:
    """

    print(str(get_current_time() + ":\t\t" + str(text_to_print)), file=class_object_name.object_stdout_file)


def print_to_stderr(object_name, text_to_print):
    """
    The function is printing report in STDOUT file
    :param object_name: The object name
    :param text_to_print:  The input text/digital value
    :return:
    """

    print(str(get_current_time() + ":ERROR!:\t" + str(text_to_print)), file=object_name.object_stderr_file)
    exit("\n\nScript finished with errors 0_o. Please check log files\n\n")


def get_class_name(class_object):
    """
    The function is returning the name of the class
    :param class_object:
    :return:
    """

    return class_object.__class__.__name__


def get_list_length(list_name):
    """
    The function is returning the length of the list
    :param list_name:
    :return:
    """

    return len(list_name)


def get_item_index_in_list(list_name, list_item_name):
    """
    The function is returning index of the list
    :param list_name:
    :param list_item_name:
    :return:
    """

    return list_name.index(list_item_name)


def get_next_value_of_list(list_name, index_value):
    """
    The function is returning the next value of the list index if it is exist, if not empty string
    :param list_name:
    :param index_value:
    :return:
    """

    try:
        return str(list_name[index_value + 1])
    except IndexError:
        return ""


def check_if_string_is_empty(string_name):
    """
    The function is checking if the string is empty or not
    :param string_name:
    :return: True if empty, False if not
    """

    if len(string_name) > 0:
        return False
    else:
        return True


def create_directory(path_to_create, directory_to_create):
    """
    The function is creating directory on the selected path
    :param path_to_create:
    :param directory_to_create:
    :return:
    """

    if not check_for_dir_existence(path_to_create, directory_to_create):
        try:
            os.mkdir(os.path.join(path_to_create, directory_to_create))
        except OSError:
            exit("ERROR: Cannot create directory\n\t" + str(os.path.join(path_to_create, directory_to_create)))


def get_current_path():
    """
    The function is returning current directory path
    :return:
    """

    return os.getcwd()


def get_string_length(string_value):
    """
    The function is returning length of string
    :param string_value:
    :return:
    """

    return len(string_value)


def set_number_of_tabs(string_value, max_tabs_number):
    """
    The function is calculating how much tabs needs. If the string <40 symbols it will calculate how much needs to have read friendly looks
    :param string_value:
    :param max_tabs_number:
    :return:
    """

    # Max tabs used
    one_tab_in_space = 8

    number_of_symbols = get_string_length(string_value)
    string_symbols_in_tabs = int(number_of_symbols) // one_tab_in_space

    return abs(string_symbols_in_tabs - max_tabs_number)


def string_column_decoration(column_one_list, column_two_list, max_tabs_number, begin_tab_number):
    """
    The function is printing into
    :param column_one_list:
    :param column_two_list:
    :param max_tabs_number:
    :param begin_tab_number:
    :return:
    """

    final_string = ""

    if get_list_length(column_one_list) <= get_list_length(column_two_list):
        list_values_count = get_list_length(column_one_list)
    else:
        list_values_count = get_list_length(column_two_list)
    for index_value in range(list_values_count):
        tabs_string = "\t" * set_number_of_tabs(column_one_list[index_value], max_tabs_number)
        final_string += str("\t" * begin_tab_number) + column_one_list[index_value] + tabs_string + column_two_list[index_value] + "\n"

    return final_string


def create_multiple_directories(path_to_create, directory_list):
    """
    The function is creating multiple directory in the current directory
    :param directory_list: directories name list
    :param path_to_create:
    :return:
    """

    for directory_name in directory_list:
        create_directory(path_to_create, directory_name)


# --------------------------------------------------- #
# --------------------- Classes --------------------- #
# --------------------------------------------------- #


class Data(object):
    """
    The metaclass of the MsipEse script
    """

    __metaclass__ = ABCMeta

    def __init__(self):
        self.main()

    @abstractmethod
    def main(self):
        """
        The main function of the class
        :return:
        """

        pass


class SubClass(Data):
    """
    The sub class of the Data class
    """

    def main(self):
        pass


class ScriptArguments:
    """
    The class is grabbing input parameters of the script
    """

    global available_script_options

    def __init__(self):
        """
        Input Class's __init__ function
        """

        self.user_arguments = self.get_all_arguments()

        for input_argument_value in self.user_arguments:
            self.check_if_help_option(input_argument_value)

    @staticmethod
    def get_all_arguments():
        """
        The function is checking and getting the script inputs information
        :return:
        """

        user_arguments = sys.argv

        if get_list_length(user_arguments) < 2:
            print_description("ERROR!:\tNo any option selected")
        else:
            del user_arguments[0]
            return user_arguments

    @staticmethod
    def check_if_help_option(input_argument):
        """
        The function is checking if the user use help option in arguments
        :param input_argument:
        :return:
        """

        available_help_options = ["-h", "--h", "-help", "--help"]

        for help_option in available_help_options:
            if input_argument == help_option:
                print_description("Help Option Selected")

    def get_user_all_inputs(self):
        """
        The main function of the arguments grabbing class
        :return: The class object
        """

        return self.user_arguments


class MsipEse:
    """
    The main class of the script's "RUN" execution part
    """

    global environment_directories_name_list
    global available_excel_options

    # --------------------------------------------------- #
    # ------------ Initialase Default Values ------------ #
    # --------------------------------------------------- #

    def __init__(self):
        """
        Project Main Run Class __init__ function
        """

        # Environment Properties

        # The objects log Name
        self.object_log_name = get_class_name(self)

        self.object_stdout_file = None
        self.object_stderr_file = None

        # The script File Name and Basename
        self.script_file_name = get_file_name_from_path(__file__)
        self.script_file_base_name = get_file_name_from_path(__file__).replace(".py", "")

        # The script execution environment path
        self.script_environment_path = None
        self.set_script_environment_path("/slowfs/us01dwt3p170/msip_ESE")

        # The projects root directory
        self.projects_root_dir = None
        self.set_projects_root_dir("/remote/proj")

        # The script Log directory
        self.script_log_dir = None

        # The internal data directory
        self.script_data_directory = None

        # The reports path directory
        self.script_reports_directory = None

        # The results path directory
        self.script_results_directory = None

        # The run path directory
        self.script_run_directory = None

        # The scripts path directory
        self.scripts_files_directory = None

        # The test cases files path directory
        self.scripts_test_cases_directory = None

        # The user arguments
        self.user_script_inputs = []

        # Setting all environment properties to default
        self.set_script_env_property()

        # Excel Properties

        self.excel_setup = {}

        # Project Properties

        # Test case excel file
        self.excel_file = None

        # Target Project Name
        self.target_project_name = None

        # Target Project Release
        self.target_project_release = None

        # Reference Project Name
        self.reference_project_name = None

        # Reference Project Release
        self.reference_project_release = None

        # Executed Test Case Package
        self.executed_test_case_package = None

    # --------------------------------------------------- #
    # ----------------- Class Functions ----------------- #
    # --------------------------------------------------- #

    def get_excel_setup(self):
        """
        The function is returning excel setup hash variable
        :return:
        """

        return self.excel_setup

    def set_excel_setup(self, excel_object):
        """
        The function is setting excel setup
        :param excel_object:
        :return:
        """

        # The excel Setup Initialization
        for optionName in available_excel_options:
            self.excel_setup[optionName] = None

        if self.get_script_excel_file is not None:
            excel_object.set_excel_setup()
            return True
        else:
            print_to_stdout(self, "No any excel file is selected")
            return False

    def set_user_script_arguments(self, user_arguments):
        """
        The function is setting user arguments
        :return:
        """

        self.user_script_inputs = user_arguments

    @property
    def get_user_script_arguments(self):
        """
        The function is returning scripts user arguments
        :return:
        """

        return self.user_script_inputs

    def create_script_env_directories(self):
        """
        The function is generating
        :return:
        """

        directories_list = self.get_script_env_property
        directories_list.remove(self.get_script_environment_path)

        for directory_path in directories_list:
            create_directory(self.get_script_environment_path, get_file_name_from_path(directory_path))

    def set_script_env_property(self):
        """
        The function is setting all property
        :return:
        """

        self.set_log_directory(os.path.join(self.get_script_environment_path, environment_directories_name_list[0]))
        self.set_reports_directory(os.path.join(self.get_script_environment_path, environment_directories_name_list[1]))
        self.set_results_directory(os.path.join(self.get_script_environment_path, environment_directories_name_list[2]))
        self.set_script_run_directory(os.path.join(self.get_script_environment_path, environment_directories_name_list[3]))
        self.set_scripts_files_directory(os.path.join(self.get_script_environment_path, environment_directories_name_list[4]))
        self.set_test_cases_directory(os.path.join(self.get_script_environment_path, environment_directories_name_list[5]))
        self.set_data_directory(os.path.join(self.get_script_environment_path, environment_directories_name_list[6]))

    @property
    def get_script_env_property(self):
        """
        The function is returning all properties as a list
        :return:
        """

        return [self.get_script_environment_path,
                self.get_log_directory,
                self.get_reports_directory,
                self.get_results_directory,
                self.get_script_run_directory,
                self.get_scripts_files_directory,
                self.get_test_cases_directory,
                self.get_data_directory]

    def create_all_environment_directories(self):
        """
        The function is creating all environment directories
        :return:
        """

        # Creating environment directories
        create_multiple_directories(self.get_script_environment_path, self.environment_directories_name_list)

    def set_script_environment_path(self, directory_path):
        """
        The function is setting script environment directory
        :param directory_path:
        :return:
        """

        self.script_environment_path = directory_path

    @property
    def get_script_environment_path(self):
        """
        The function is returning script environment path
        :return:
        """

        return self.script_environment_path

    def set_log_directory(self, directory_path):
        """
        The function changing log directory path
        :param directory_path:
        :return:
        """

        self.script_log_dir = directory_path

    @property
    def get_log_directory(self):
        """
        The function is returning LOG directory
        :return:
        """

        return self.script_log_dir

    def set_data_directory(self, directory_path):
        """
        The function is changing data directory path
        :param directory_path:
        :return:
        """

        self.script_data_directory = directory_path

    @property
    def get_data_directory(self):
        """
        The function is returning DATA directory
        :return:
        """

        return self.script_data_directory

    def set_reports_directory(self, directory_path):
        """
        The function is changing reports directory path
        :param directory_path:
        :return:
        """

        self.script_reports_directory = directory_path

    @property
    def get_reports_directory(self):
        """
        The function is returning reports directory
        :return:
        """

        return self.script_reports_directory

    def set_results_directory(self, directory_path):
        """
        The function is changing reports directory path
        :param directory_path:
        :return:
        """

        self.script_results_directory = directory_path

    @property
    def get_results_directory(self):
        """
        The function is returning reports directory
        :return:
        """

        return self.script_results_directory

    def set_scripts_files_directory(self, directory_path):
        """
        The function is changing reports directory path
        :param directory_path:
        :return:
        """

        self.scripts_files_directory = directory_path

    @property
    def get_scripts_files_directory(self):
        """
        The function is returning reports directory
        :return:
        """

        return self.scripts_files_directory

    def set_test_cases_directory(self, directory_path):
        """
        The function is changing reports directory path
        :param directory_path:
        :return:
        """

        self.scripts_test_cases_directory = directory_path

    @property
    def get_test_cases_directory(self):
        """
        The function is returning reports directory
        :return:
        """

        return self.scripts_test_cases_directory

    def set_script_run_directory(self, directory_path):
        """
        The function is changing reports directory path
        :param directory_path:
        :return:
        """

        self.script_run_directory = directory_path

    @property
    def get_script_run_directory(self):
        """
        The function is returning reports directory
        :return:
        """

        return self.script_run_directory

    def set_projects_root_dir(self, directory_path):
        """
        The function is defining projects root directory, by default it is /remote/proj
        :param directory_path:
        :return:
        """

        self.projects_root_dir = directory_path

    @property
    def get_projects_root_dir(self):
        """
        The function is returning projects root directory, by default it is /remote/proj
        :return:
        """

        return self.projects_root_dir

    def set_script_excel_file(self, file_location):
        """
        The function is defining projects root directory, by default it is /remote/proj
        :param file_location:
        :return:
        """

        self.excel_file = file_location

    @property
    def get_script_excel_file(self):
        """
        The function is returning reports directory
        :return:
        """

        return self.excel_file

    def set_target_project_name(self, value):
        """
        The function is defining projects root directory, by default it is /remote/proj
        :param value:
        :return:
        """

        self.target_project_name = value

    @property
    def get_target_project_name(self):
        """
        The function is returning reports directory
        :return:
        """

        return self.target_project_name

    def set_target_project_release(self, value):
        """
        The function is defining projects root directory, by default it is /remote/proj
        :param value:
        :return:
        """

        self.target_project_release = value

    @property
    def get_target_project_release(self):
        """
        The function is returning reports directory
        :return:
        """

        return self.target_project_release

    def set_reference_project_name(self, value):
        """
        The function is defining projects root directory, by default it is /remote/proj
        :param value:
        :return:
        """

        self.reference_project_name = value

    @property
    def get_reference_project_name(self):
        """
        The function is returning reports directory
        :return:
        """

        return self.reference_project_name

    def set_reference_project_release(self, value):
        """
        The function is defining projects root directory, by default it is /remote/proj
        :param value:
        :return:
        """

        self.reference_project_release = value

    @property
    def get_reference_project_release(self):
        """
        The function is returning reports directory
        :return:
        """

        return self.reference_project_release

    def set_executed_test_case_package(self, value):
        """
        The function is defining projects root directory, by default it is /remote/proj
        :param value:
        :return:
        """

        self.executed_test_case_package = value

    @property
    def get_executed_test_case_package(self):
        """
        The function is returning reports directory
        :return:
        """

        return self.executed_test_case_package

    # --------------------------------------------------- #
    # ----------------- Internal Class ------------------ #
    # --------------------------------------------------- #

    class ScriptInputs:
        """
        The Script Input class, which will check for script input correctness
        """

        global available_script_options

        # --------------------------------------------------- #
        # ------------ Initialase Default Values ------------ #
        # --------------------------------------------------- #

        def __init__(self, msip_ese_object):
            """
            Project Main Run Class __init__ function
            """

            # The MsipEse object's instance
            self.msip_ese_object = msip_ese_object

            self.all_arguments = self.msip_ese_object.get_user_script_arguments

        @staticmethod
        def get_option_name_and_value(option_name, argument_list):
            """
            The function is returning option name and value. If it is not under available option or no any value exist, returning empty list
            :return:
            """

            global available_script_options

            for available_option_name in available_script_options:
                if option_name == available_option_name:
                    list_next_item = get_next_value_of_list(argument_list, get_item_index_in_list(argument_list, option_name))
                    if list_next_item != "":
                        return [option_name, list_next_item]
                    else:
                        return []
            return []

        def read_arguments(self):
            """
            The function is returning hash, which contains all options of the scrip if they are used (as key), and their values
            :return:
            """

            grab_script_setup = {}

            for script_argument in self.all_arguments:
                if "-" in script_argument:
                    ScriptArguments.check_if_help_option(script_argument)
                    used_option = self.get_option_name_and_value(script_argument, self.all_arguments)
                    if get_list_length(used_option) == 2:
                        if used_option[1][0] != "-":
                            grab_script_setup[used_option[0]] = used_option[1]

            return grab_script_setup

        def get_script_arguments(self):
            """
            The main function of the ScriptInputs class
            :return:
            """

            script_all_inputs_hash = self.read_arguments()
            all_options_name = script_all_inputs_hash.keys()
            enable_script_execution = False
            for option_name in all_options_name:
                if available_script_options[5] == option_name:
                    if check_for_dir_existence(get_file_path(script_all_inputs_hash[option_name]), get_file_name_from_path(script_all_inputs_hash[option_name])):
                        self.msip_ese_object.set_script_environment_path(script_all_inputs_hash[option_name])
                        self.msip_ese_object.set_script_env_property()
                    else:
                        exit("ERROR!:\tThe run directory path is not exist\t'" + str(script_all_inputs_hash[option_name]) + "'\n\tPlease check script arguments")
                elif available_script_options[0] == option_name:
                    enable_script_execution = True
                elif available_script_options[1] == option_name:
                    enable_script_execution = True

            if not enable_script_execution:
                print_description(
                    "ERROR!:\tUser should define at least one of the following options:\n\t\t'" + available_script_options[0] + "'\n\t\t'" + available_script_options[1] + "'\n")
                return None
            else:
                return script_all_inputs_hash

        def set_script_inputs(self, script_inputs_argument_hash):
            """
            The function is printing
            :param script_inputs_argument_hash:
            :return:
            """

            script_inputs_all_options = script_inputs_argument_hash.keys()

            print("USER INPUTS:\n")

            # Setting script inputs
            for script_option_name in script_inputs_all_options:
                script_option_value = script_inputs_argument_hash[script_option_name]
                print("\t" + script_option_name + "\t" + script_option_value)
                if script_option_name == available_script_options[0]:
                    self.msip_ese_object.set_script_excel_file(script_option_value)
                elif script_option_name == available_script_options[1]:
                    self.msip_ese_object.set_target_project_name(script_option_value)
                elif script_option_name == available_script_options[2]:
                    self.msip_ese_object.set_target_project_release(script_option_value)
                elif script_option_name == available_script_options[3]:
                    self.msip_ese_object.set_reference_project_name(script_option_value)
                elif script_option_name == available_script_options[4]:
                    self.msip_ese_object.set_reference_project_release(script_option_value)
                elif script_option_name == available_script_options[5]:
                    self.msip_ese_object.set_script_environment_path(script_option_value)
                elif script_option_name == available_script_options[6]:
                    self.msip_ese_object.set_executed_test_case_package(script_option_value)
                elif script_option_name == available_script_options[7]:
                    self.msip_ese_object.set_projects_root_directory(script_option_value)

    class Excel:
        """
        The class is for read and get appropriate information from excel file
        """

        global available_excel_options

        def __init__(self, msip_ese_object):
            """
            Initial function of the class
            """

            # --------------- Variables --------------- #

            self.msip_ese_object = msip_ese_object

        def check_excel_option_name_and_value(self, excel_option_name, excel_option_value):
            """
            The function is checking the option name and value for correctness,
            and returns None if not correct and added value on appropriate option setup of excel variable self.excelSetup
            :param excel_option_name:
            :param excel_option_value:
            :return: True if found and False if not
            """

            global available_excel_options

            for optionName in available_excel_options:
                if excel_option_name.upper() == str(optionName).upper():
                    if not check_if_string_is_empty(excel_option_value):
                        self.msip_ese_object.excel_setup[optionName] = excel_option_value
                        return True
                    else:
                        return False

            return False

        def read_excel(self, excel_file):
            """
            The function is reading excel file
            :param excel_file:
            :return:
            """

            print_to_stdout(self.msip_ese_object, "Reading excel file\t" + excel_file)
            try:
                excel_workbook_object = read_excel_module(excel_file)
            except XLRDError as xlrdException:
                print_to_stderr(self.msip_ese_object, "File\t" + excel_file + "\n\t\t" + str(xlrdException))

            # noinspection PyUnboundLocalVariable
            all_sheets_name = excel_workbook_object.sheet_names()

            for sheet_name in all_sheets_name:
                work_sheet_object = excel_workbook_object.sheet_by_name(sheet_name)
                row_number = work_sheet_object.nrows
                for current_row in range(1, row_number):
                    excel_option_name = str(work_sheet_object.cell_value(current_row, 1))
                    excel_option_value = str(work_sheet_object.cell_value(current_row, 2))
                    excel_option_comment = str(work_sheet_object.cell_value(current_row, 4))
                    row_contains_information = self.check_excel_option_name_and_value(excel_option_name, excel_option_value)
                    if row_contains_information and (not check_if_string_is_empty(excel_option_comment)):
                        print("IMPORTANT NOTE! USER MAKES COMMENT FOR TEST CASE OPTION IN EXCEL FILE\n\tCOMMENT:\t" + excel_option_comment + "\n\tLINE:\t\t" + str(current_row + 1))

        def main(self, excel_file):
            """
            Main function of the class
            :param excel_file:
            :return:
            """

            if check_if_string_is_empty(excel_file):
                print_to_stdout(self.msip_ese_object, "No any excel file selected.\nSkip the step")
            else:
                if check_for_file_existence(get_file_path(excel_file), get_file_name_from_path(excel_file)):
                    self.read_excel(excel_file)
                else:
                    print_to_stderr(self.msip_ese_object, "Cannot read excel file\t" + os.path.join(get_file_path(excel_file), excel_file))

            print_to_stdout(self.msip_ese_object, "USER IS SET FOLLOWING EXCEL'S OPTION(S)\n")
            excel_options = self.msip_ese_object.excel_setup.keys()
            for option in excel_options:
                if self.msip_ese_object.excel_setup[option] is not None:
                    number_of_tabs = "\t" * set_number_of_tabs(option, 5)
                    print_to_stdout(self.msip_ese_object, str("\t" + option + number_of_tabs + self.msip_ese_object.excel_setup[option]))

            print_to_stdout(self.msip_ese_object, "\n\nFOLLOWING EXCEL'S OPTIONS ARE NOT USED\n")
            for option in excel_options:
                if self.msip_ese_object.excel_setup[option] is None:
                    number_of_tabs = "\t" * set_number_of_tabs(option, 5)
                    print_to_stdout(self.msip_ese_object, str("\t" + option + number_of_tabs + str(self.msip_ese_object.excel_setup[option])))

            return self

    def main(self):
        """
        Main Function of the MsipEse Class
        :return:
        """

        script_inputs_instance = self.ScriptInputs(self)
        script_arguments = script_inputs_instance.get_script_arguments()
        script_inputs_instance.set_script_inputs(script_arguments)

        # Creating environment directories

        self.create_script_env_directories()
        # Opening log files for the script
        # The script stdout file object
        self.object_stdout_file = open_file_for_writing(self.script_log_dir, self.object_log_name + ".stdout")
        # The script stderr file object
        self.object_stderr_file = open_file_for_writing(self.script_log_dir, self.object_log_name + ".stderr")

        print_to_stdout(self, "READING SCRIPT ARGUMENTS")
        print_to_stdout(self, "Script Inputs Is:\n" + string_column_decoration(list(script_arguments.keys()), list(script_arguments.values()), 6, 4))


def main():
    """
    The main function of the script
    :return:
    """

    user_script_inputs = ScriptArguments().get_user_all_inputs()

    evaluation_object = MsipEse()
    evaluation_object.set_user_script_arguments(user_script_inputs)
    evaluation_object.main()


if __name__ == '__main__':

    print("\n\nSTART TIME:\t" + get_current_time() + "\n\n")

    main()

    print("\n\nFINISHED TIME:\t" + get_current_time())
    print("\n\nScript Finished Successfully ^_^\n\n")
