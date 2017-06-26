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
import errno

__author__ = 'Vladimir'

"""
USAGE:  program
        EXAMPLE: msip_ESE.py

DESCRIPTION:
        The script is running ESE flow.
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


# Sample run script execution wait time. Default = 3 minute, You can change it
sample_process_wait_time = 3

# The script environment directories list
environment_directories_name_list = ["LOGS",  # Index[0] Logs directory name
                                     "REPORTS",  # Index[1] Reports directory name
                                     "RESULTS",  # Index[2] Results directory name
                                     "RUN_DIR",  # Index[3] Run directory name
                                     "SCRIPTS",  # Index[4] Scripts directory name
                                     "TESTCASES",  # Index[5] Test cases directory name
                                     "DATA"  # Index[6] Internal data directory name. DATA/ [PEX_SAMPLE_RUN_SCRIPTS, SAMPLE_OA_LIBRARIES, SIM_SAMPLE_RUN_SCRIPTS]
                                     ]

# Available Options For the Script
available_script_options = ["-excelFile",  # Index[0] Excel file
                            "-targetProjectName",  # Index[1] Target Project Name
                            "-targetProjectRelease",  # Index[2] Target Project Release
                            "-referenceProjectName",  # Index[3] Reference Project Name
                            "-referenceProjectRelease",  # Index[4] Reference Project Release
                            "-runDirectory",  # Index[5] Script Run Directory
                            "-executedTestCasePackage",  # Index[6] Executed test case package(s)
                            "-projectsRootDirectory",  # Index[7] Projects root directory path
                            "-forceUpdateTestCase",  # Index[8] Force Updating Test Case Package
                            "-executeFlow"  # Index[9]  Execute only selected step. Available values ENV_UPDATE/TEST_CASE_UPDATE/LVS/PEX/SIM/REPORT/CLEAN/ALL
                            ]

# Available Steps Of The Flow For The Script
available_flows = ["UPDATE_ENV",
                   "UPDATE_TEST_CASE",
                   "PEX",
                   "SIM",
                   "REPORT",
                   "CLEAN",
                   "ALL"]

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
                           "Extraction Criteria For GDS file(s)",
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
                           "Target LVS sourceme",
                           "Target LVS options",
                           "Reference LVS sourceme",
                           "Reference LVS options",
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

# All available data regarding project setup
project_setup_data = ["PEX TOOL NAME",  # Index[0]  The LVS tool name ICV/CALIBRE/HERCULES
                      "PEX TOOL VERSION",  # Index[1]  The LVS tool version
                      "PEX DECK",  # Index[2]  The LVS deck
                      "RCXT TOOL VERSION",  # Index[3]  The STAR RC tool version
                      "RCXT DECK",  # Index[4]  The STAR RC typical.rules tcad grid file
                      "RCXT STARCMD",  # Index[5]  The STAR RC starcmd file
                      ]

# Test Case Directory Structure
project_test_case_directories_list = ["EXCEL", "GDS", "LVS_NETLIST", "TEST_BENCH", "MEASURE_FILES", "OTHER_INCLUDES", "USER_RESULTS"]
untar_directory_name = "UNTAR"

# The project environment file/directories name
project_environment_file_name = "env.tcl"
project_cad_directory_name = "cad"
project_extract_directory_name = "PEX"
project_sim_directory_name = "SIM"
project_SPF_directory_name = "SPF"
project_sample_oa_library_directory_name = "SAMPLE_OA_LIBRARIES"
project_sample_oa_library_names_list = ["SampleLibrary"]
project_sample_oa_cell_name = "SampleExtract"
project_sample_runscript_file_name = "sample_runscript.sh"
project_sample_runscript_location_dir_name = "SAMPLE_RUNSCRIPT_FILES"
available_project_tools_name = ["ICV",  # INDEX 0 Default value
                                "HERCULES",
                                "CALIBRE"]

# The lvs report extensions
project_lvs_report_extensions = {available_project_tools_name[0]: ".LVS_ERRORS",
                                 available_project_tools_name[1]: ".LVS_ERRORS",
                                 available_project_tools_name[2]: ".cell_results"}

project_extract_file_extension = ".spf"
project_extract_ideal_file_prefix = "ideal_"
project_extract_ideal_file_extension = ".raw"
gds_file_extension = ".gds"
netlist_file_extension = [".cdl", ".sp", ".cir"]
gds_config_file_extension = ".config"
tar_file_extension = ".tar.gz"

available_package_directory_tags_list = ["insideTarFile:", "insideTestCasePackagePath:"]

top_cell_subckt_file_name = "topCellSubcktFile"
subckt_start_recognition_word = ".subckt"
subckt_end_recognition_word = ".end"


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
                The script is running ESE flow.
                ESE - Extraction and Simulation Evaluation flow is for evaluating the CCS/PCS setup updates' impact in simulation results.

        FOR SUPPORT(BUG/ENHANCEMENT):
                Please send e-mail to "vlad@synopsys.com"

        AUTHOR:
                Vladimir Danielyan
        """.format(get_script_options_string())

    print("\n\t" + str(violated_case) + "\n")
    exit(description)


def get_file_size(file_item):
    """
    The function is returning file size
    :param file_item:
    :return:
    """

    if file_item is not None:
        if check_for_file_existence(get_file_path(file_item), get_file_name_from_path(file_item)):
            try:
                return int(os.path.getsize(file_item))
            except PermissionError:
                return 0
    else:
        return 0


def get_current_time():
    """
    The function is returning time in string format
    :return:
    """

    current_time = time.time()
    current_date_time = datetime.datetime.fromtimestamp(current_time).strftime('%m/%d %H:%M')
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
        if option_name == available_script_options[8]:
            final_string += string_column_decoration([str(option_name)], ["# Available Value:\t| TRUE | (default is FALSE)"], 5, 2)
        elif option_name == available_script_options[9]:
            all_values = ""
            for value in available_flows:
                all_values += "| " + value + " |"
            all_values += " (default is ALL)"
            final_string += string_column_decoration([str(option_name)], ["# Available Values:\t" + all_values], 5, 2)
        else:
            final_string += string_column_decoration([str(option_name)], ["# Available Value:\t'" + str(option_name).replace("-", "") + "'"], 5, 2)

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
        try:
            return os.listdir(directory_path)
        except PermissionError:
            print("WARNING!:\tPermission denied. Cannot read information from directory:\t'" + directory_path + "'")
            return []
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


def process_timeout(process_object, text_to_display):
    """
    The function is
    :param process_object:
    :param text_to_display:
    :return:
    """

    print("in function")

    if process_object.poll() is None:
        print("in if else")
        try:
            print("in try")
            process_object.kill()
            print(text_to_display)
            return True
        except OSError as e:
            if e.errno != errno.ESRCH:
                raise

    else:
        print("false")
        return False


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

    try:
        return os.path.dirname(os.path.abspath(full_path_to_the_file))
    except IOError:
        exit("ERROR!:\tCannot find file\t'" + full_path_to_the_file + "'")


def get_file_name_from_path(full_path_to_the_file):
    """
    The function is returning file name from the full path
    :param full_path_to_the_file:
    :return:
    """

    try:
        return str(os.path.basename(full_path_to_the_file))
    except IOError:
        exit("ERROR!:\tCannot find file\t'" + full_path_to_the_file + "'")


def print_to_stdout(class_object_name, text_to_print):
    """
    The function is printing report in STDOUT file
    :param class_object_name: The object name
    :param text_to_print: The input text/digital value
    :return:
    """

    if "NEW LINE" != str(text_to_print).upper():
        print(str(get_current_time() + ":\t\t" + str(text_to_print)), file=class_object_name.object_stdout_file)
    else:
        print("\n", file=class_object_name.object_stdout_file)


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


def create_directories_hierarchy(path_to_create, directory_list_to_create):
    """
    The function is creating directory on the selected path
    :param path_to_create:
    :param directory_list_to_create:
    :return:
    """

    for directory_name in directory_list_to_create:
        create_directory(path_to_create, directory_name)
        path_to_create = os.path.join(path_to_create, directory_name)

    return path_to_create


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
        tabs_string = "\t" * set_number_of_tabs(str(column_one_list[index_value]), max_tabs_number)
        final_string += str("\t" * begin_tab_number) + str(column_one_list[index_value]) + tabs_string + str(column_two_list[index_value]) + "\n"

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
    global available_project_tools_name

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
        self.set_projects_root_dir("/remote/cad-rep/projects")

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
        # Initialisation of excel_setup hash
        self.set_excel_setup_none_value()

        # Project Properties

        # Test case excel file
        self.excel_file = None

        # Target Project Type
        self.target_project_type = None

        # Target Project Name
        self.target_project_name = None

        # Target Project Release
        self.target_project_release = None

        # Target Project Metal Stack List
        self.target_project_metal_stack_list = []

        # Target Project Layer Map File Layers
        self.target_project_layers = {}

        # Reference Project Type
        self.reference_project_type = None

        # Reference Project Name
        self.reference_project_name = None

        # Reference Project Release
        self.reference_project_release = None

        # Target Reference Metal Stack List
        self.reference_project_metal_stack_list = []

        # Reference Project Layer Map File Layers
        self.reference_project_layers = {}

        # Executed Test Case Package
        self.executed_test_case_package = None

        # The test cases hash. Key = test case name, Value = test case path
        self.project_test_cases = {}

        # Target PEX Tool name
        self.target_project_pex_tool_name = None
        self.set_target_project_pex_tool_name(None)

        # Reference PEX Tool name
        self.reference_project_pex_tool_name = None
        self.set_reference_project_pex_tool_name(None)

        # Target PEX Tool name
        self.target_project_pex_tool_version = None
        self.set_target_project_pex_tool_version(None)

        # Reference PEX Tool name
        self.reference_project_pex_tool_version = None
        self.set_reference_project_pex_tool_version(None)

        # Target PEX Tool name
        self.target_project_pex_tool_deck = None
        self.set_target_project_pex_tool_deck(None)

        # Reference PEX Tool name
        self.reference_project_pex_tool_deck = None
        self.set_reference_project_pex_tool_deck(None)

        # Target PEX Tool name
        self.target_project_pex_tool_option_file = None
        self.set_target_project_pex_tool_option_file(None)

        # Reference PEX Tool name
        self.reference_project_pex_tool_option_file = None
        self.set_reference_project_pex_tool_option_file(None)

        self.target_project_pex_tool_source_file = None
        self.set_target_project_pex_tool_source_file(None)

        # Reference PEX Tool name
        self.reference_project_pex_tool_source_file = None
        self.set_reference_project_pex_tool_source_file(None)

        # Target EXTRACT Tool
        self.target_project_extract_tool_version = None
        self.set_target_project_extract_tool_version(None)

        # Reference EXTRACT Tool
        self.reference_project_extract_tool_version = None
        self.set_reference_project_extract_tool_version(None)

        # Target EXTRACT Tool
        self.target_project_extract_tool_deck = None
        self.set_target_project_extract_tool_deck(None)

        # Reference EXTRACT Tool
        self.reference_project_extract_tool_deck = None
        self.set_reference_project_extract_tool_deck(None)

        # Target EXTRACT Tool
        self.target_project_extract_tool_starcmd = None
        self.set_target_project_extract_tool_starcmd(None)

        # Reference EXTRACT Tool
        self.reference_project_extract_tool_starcmd = None
        self.set_reference_project_extract_tool_starcmd(None)

        # Force adding test case enable
        self.force_add_test_case = False

        # Script flow values
        self.update_environment = False
        self.update_test_case = False
        self.execute_pex = False
        self.execute_simulation = False
        self.execute_report = False
        self.execute_clean_project = False

    # --------------------------------------------------- #
    # ----------------- Class Functions ----------------- #
    # --------------------------------------------------- #

    def check_for_reference_project_execution(self):
        """
        The function is returning True if there is reference project setup and False if not
        :return:
        """

        if self.get_reference_project_name is not None:
            return True
        else:
            return False

    def check_script_setup_correctness(self):
        """
        The script is checking if the script setup is correct
        :return:
        """

        target_project_path = os.path.join(str(self.get_projects_root_dir),
                                           str(self.get_target_project_type),
                                           str(self.get_target_project_name),
                                           str(self.get_target_project_release))

        if not check_for_dir_existence(target_project_path, project_cad_directory_name):
            print_to_stderr(self, "Cannot find target project. Please check Script inputs. No such directory:\t'" + str(
                os.path.join(target_project_path, project_cad_directory_name)) + "'")

    def enable_update_environment(self):
        """
        The function is enabling Script flow #
        :return:
        """

        self.update_environment = True

    def disable_update_environment(self):
        """
        The function is disabling Script flow #
        :return:
        """

        self.update_environment = False

    def check_if_update_environment(self):
        """
        The function is returning script flow value
        :return:
        """

        return self.update_environment

    def enable_execute_pex(self):
        """
        The function is enabling Script flow #
        :return:
        """

        self.execute_pex = True

    def disable_execute_pex(self):
        """
        The function is disabling Script flow #
        :return:
        """

        self.execute_pex = False

    def check_if_execute_pex(self):
        """
        The function is returning script flow value
        :return:
        """

        return self.execute_pex

    def enable_update_test_case(self):
        """
        The function is enabling Script flow #
        :return:
        """

        self.update_test_case = True

    def disable_update_test_case(self):
        """
        The function is disabling Script flow #
        :return:
        """

        self.update_test_case = False

    def check_if_update_test_case(self):
        """
        The function is returning script flow value
        :return:
        """

        return self.update_test_case

    def enable_execute_simulation(self):
        """
        The function is enabling Script flow #
        :return:
        """

        self.execute_simulation = True

    def disable_execute_simulation(self):
        """
        The function is disabling Script flow #
        :return:
        """

        self.execute_simulation = False

    def check_if_execute_simulation(self):
        """
        The function is returning script flow value
        :return:
        """

        return self.execute_simulation

    def enable_execute_report(self):
        """
        The function is enabling Script flow #
        :return:
        """

        self.execute_report = True

    def disable_execute_report(self):
        """
        The function is disabling Script flow #
        :return:
        """

        self.execute_report = False

    def check_if_execute_report(self):
        """
        The function is returning script flow value
        :return:
        """

        return self.execute_report

    def enable_execute_clean_project(self):
        """
        The function is enabling Script flow #
        :return:
        """

        self.execute_clean_project = True

    def disable_execute_clean_project(self):
        """
        The function is disabling Script flow #
        :return:
        """

        self.execute_clean_project = False

    def check_if_execute_clean_project(self):
        """
        The function is returning script flow value
        :return:
        """

        return self.execute_clean_project

    def set_executed_flow(self, flow_option_value):
        """
        The function is setting the executed flow
        :param flow_option_value:
        :return:
        """

        if str(flow_option_value).upper() == "NONE":
            return
        elif flow_option_value == available_flows[0]:
            self.enable_update_environment()  # 1. Enabled
            self.disable_update_test_case()  # Disabled
            self.disable_execute_pex()  # Disabled
            self.disable_execute_simulation()  # Disabled
            self.disable_execute_report()  # Disabled
            self.disable_execute_clean_project()  # Disabled
        elif flow_option_value == available_flows[1]:
            self.disable_update_environment()  # Disabled
            self.enable_update_test_case()  # 2. Enabled
            self.disable_execute_pex()  # Disabled
            self.disable_execute_simulation()  # Disabled
            self.disable_execute_report()  # Disabled
            self.disable_execute_clean_project()  # Disabled
        elif flow_option_value == available_flows[2]:
            # Executing Flow Step 3
            self.disable_update_environment()  # Disabled
            self.disable_update_test_case()  # Disabled
            self.enable_execute_pex()  # 3. Enabled
            self.disable_execute_simulation()  # Disabled
            self.disable_execute_report()  # Disabled
            self.disable_execute_clean_project()  # Disabled
        elif flow_option_value == available_flows[3]:
            # Executing Flow Step 4
            self.disable_update_environment()  # Disabled
            self.disable_update_test_case()  # Disabled
            self.disable_execute_pex()  # Disabled
            self.enable_execute_simulation()  # 4. Enabled
            self.disable_execute_report()  # Disabled
            self.disable_execute_clean_project()  # Disabled
        elif flow_option_value == available_flows[4]:
            # Executing Flow Step 5
            self.disable_update_environment()  # Disabled
            self.disable_update_test_case()  # Disabled
            self.disable_execute_pex()  # Disabled
            self.disable_execute_simulation()  # Disabled
            self.enable_execute_report()  # 5. Enabled
            self.disable_execute_clean_project()  # Disabled
        elif flow_option_value == available_flows[5]:
            # Executing Flow Step 5
            self.disable_update_environment()  # Disabled
            self.disable_update_test_case()  # Disabled
            self.disable_execute_pex()  # Disabled
            self.disable_execute_simulation()  # Disabled
            self.disable_execute_report()  # Disabled
            self.enable_execute_clean_project()  # 6. Enabled
        else:
            # Executing All Steps of The Flow
            self.enable_update_environment()
            self.enable_update_test_case()
            self.enable_execute_pex()
            self.enable_execute_simulation()
            self.enable_execute_report()
            self.enable_execute_clean_project()

    def enable_force_add_test_case(self):
        """
        The function is enabling force adding test case enable option
        :return:
        """

        self.force_add_test_case = True

    def disable_force_add_test_case(self):
        """
        The function is disabling force adding test case enable option
        :return:
        """

        self.force_add_test_case = False

    @property
    def get_force_add_test_case_option(self):
        """
        The function is returning force adding test case option
        :return:
        """

        return self.force_add_test_case

    def set_target_project_pex_tool_name(self, value):
        """
        The function is setting target project PEX tool name
        :param value:
        :return:
        """

        for tool_name in available_project_tools_name:
            if str(value).upper() == tool_name:
                self.target_project_pex_tool_name = tool_name
                return
            else:
                self.target_project_pex_tool_name = available_project_tools_name[0]

        return

    @property
    def get_target_project_pex_tool_name(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.target_project_pex_tool_name

    def set_reference_project_pex_tool_name(self, value):
        """
        The function is setting target project PEX tool name
        :param value:
        :return:
        """

        for tool_name in available_project_tools_name:
            if str(value).upper() == tool_name:
                self.reference_project_pex_tool_name = tool_name
            else:
                self.reference_project_pex_tool_name = available_project_tools_name[0]

    @property
    def get_reference_project_pex_tool_name(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.reference_project_pex_tool_name

    def set_target_project_pex_tool_version(self, value):
        """
        The function is setting target project PEX tool version
        :param value:
        :return:
        """

        self.target_project_pex_tool_version = value

    @property
    def get_target_project_pex_tool_version(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.target_project_pex_tool_version

    def set_reference_project_pex_tool_version(self, value):
        """
        The function is setting target project PEX tool name
        :param value:
        :return:
        """

        self.reference_project_pex_tool_version = value

    @property
    def get_reference_project_pex_tool_version(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.reference_project_pex_tool_version

    def set_target_project_extract_tool_version(self, value):
        """
        The function is setting target project PEX tool version
        :param value:
        :return:
        """

        self.target_project_extract_tool_version = value

    @property
    def get_target_project_extract_tool_version(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.target_project_extract_tool_version

    def set_reference_project_extract_tool_version(self, value):
        """
        The function is setting target project PEX tool name
        :param value:
        :return:
        """

        self.reference_project_extract_tool_version = value

    @property
    def get_reference_project_extract_tool_version(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.reference_project_extract_tool_version

    def set_target_project_extract_tool_deck(self, value):
        """
        The function is setting target project PEX tool version
        :param value:
        :return:
        """

        self.target_project_extract_tool_deck = value

    @property
    def get_target_project_extract_tool_deck(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.target_project_extract_tool_deck

    def set_reference_project_extract_tool_deck(self, value):
        """
        The function is setting target project PEX tool name
        :param value:
        :return:
        """

        self.reference_project_extract_tool_deck = value

    @property
    def get_reference_project_extract_tool_deck(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.reference_project_extract_tool_deck

    def set_target_project_extract_tool_starcmd(self, value):
        """
        The function is setting target project PEX tool version
        :param value:
        :return:
        """

        self.target_project_extract_tool_starcmd = value

    @property
    def get_target_project_extract_tool_starcmd(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.target_project_extract_tool_starcmd

    def set_reference_project_extract_tool_starcmd(self, value):
        """
        The function is setting target project PEX tool name
        :param value:
        :return:
        """

        self.reference_project_extract_tool_starcmd = value

    @property
    def get_reference_project_extract_tool_starcmd(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.reference_project_extract_tool_starcmd

    def set_target_project_pex_tool_deck(self, value):
        """
        The function is setting target project PEX tool version
        :param value:
        :return:
        """

        self.target_project_pex_tool_deck = value

    @property
    def get_target_project_pex_tool_deck(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.target_project_pex_tool_deck

    def set_reference_project_pex_tool_deck(self, value):
        """
        The function is setting target project PEX tool name
        :param value:
        :return:
        """

        self.reference_project_pex_tool_deck = value

    @property
    def get_reference_project_pex_tool_deck(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.reference_project_pex_tool_deck

    def set_target_project_pex_tool_option_file(self, value):
        """
        The function is setting target project PEX tool version
        :param value:
        :return:
        """

        self.target_project_pex_tool_option_file = value

    @property
    def get_target_project_pex_tool_option_file(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.target_project_pex_tool_option_file

    def set_reference_project_pex_tool_option_file(self, value):
        """
        The function is setting target project PEX tool name
        :param value:
        :return:
        """

        self.reference_project_pex_tool_option_file = value

    @property
    def get_reference_project_pex_tool_option_file(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.reference_project_pex_tool_option_file

    def set_target_project_pex_tool_source_file(self, value):
        """
        The function is setting target project PEX tool version
        :param value:
        :return:
        """

        self.target_project_pex_tool_source_file = value

    @property
    def get_target_project_pex_tool_source_file(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.target_project_pex_tool_source_file

    def set_reference_project_pex_tool_source_file(self, value):
        """
        The function is setting target project PEX tool name
        :param value:
        :return:
        """

        self.reference_project_pex_tool_source_file = value

    @property
    def get_reference_project_pex_tool_source_file(self):
        """
        The function is returning target project pex tool name
        :return:
        """

        return self.reference_project_pex_tool_source_file

    def get_excel_setup(self):
        """
        The function is returning excel setup hash variable
        :return:
        """

        return self.excel_setup

    def set_excel_setup_none_value(self):
        """
        The function is setting excel setup
        :return:
        """

        # The excel Setup Initialization
        for optionName in available_excel_options:
            self.excel_setup[optionName] = None

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
        The function is defining projects root directory, by default it is /remote/cad-rep/projects
        :param directory_path:
        :return:
        """

        self.projects_root_dir = directory_path

    @property
    def get_projects_root_dir(self):
        """
        The function is returning projects root directory, by default it is /remote/cad-rep/projects
        :return:
        """

        return self.projects_root_dir

    def set_script_excel_file(self, file_location):
        """
        The function is defining projects root directory, by default it is /remote/cad-rep/projects
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

    def set_target_project_type(self, value):
        """
        The function is defining target project type
        :param value:
        :return:
        """

        self.target_project_type = value

    @property
    def get_target_project_type(self):
        """
        The function is defining target project type
        :return:
        """

        return self.target_project_type

    def set_target_project_name(self, value):
        """
        The function is defining projects root directory, by default it is /remote/cad-rep/projects
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
        The function is defining projects root directory, by default it is /remote/cad-rep/projects
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

    def set_reference_project_type(self, value):
        """
        The function is defining target project type
        :param value:
        :return:
        """

        self.reference_project_type = value

    @property
    def get_reference_project_type(self):
        """
        The function is defining target project type
        :return:
        """

        return self.reference_project_type

    def set_reference_project_name(self, value):
        """
        The function is defining projects root directory, by default it is /remote/cad-rep/projects
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
        The function is defining projects root directory, by default it is /remote/cad-rep/projects
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

    def set_target_project_metal_stack_list(self, list_value):
        """
        The function is defining projects root directory, by default it is /remote/cad-rep/projects
        :param list_value:
        :return:
        """

        self.target_project_metal_stack_list = []
        for value in list_value:
            self.target_project_metal_stack_list.append(value)

    @property
    def get_target_project_metal_stack_list(self):
        """
        The function is returning reports directory
        :return:
        """

        return self.target_project_metal_stack_list

    def set_reference_project_metal_stack_list(self, list_value):
        """
        The function is defining projects root directory, by default it is /remote/cad-rep/projects
        :param list_value:
        :return:
        """

        self.reference_project_metal_stack_list = []
        for value in list_value:
            self.reference_project_metal_stack_list.append(value)

    @property
    def get_reference_project_metal_stack_list(self):
        """
        The function is returning reports directory
        :return:
        """

        return self.reference_project_metal_stack_list

    def set_project_test_cases(self, hash_value):
        """
        The project is setting project_test_cases hash
        :param hash_value:
        :return:
        """

        self.project_test_cases = hash_value

    @property
    def get_project_test_cases(self):
        """
        The function is returning project_test_cases
        :return:
        """

        return self.project_test_cases

    def set_executed_test_case_package(self, value):
        """
        The function is defining projects root directory, by default it is /remote/cad-rep/projects
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

    # The MsipEse class methods

    # --------------------------------------------------- #
    # ----------------- Internal Class ------------------ #
    # --------------------------------------------------- #

    class ProjectEnvironment:
        """
        The class contains project environment variables and methods to setup environment and do sample extract flow
        """

        global available_excel_options
        global project_environment_file_name
        global project_cad_directory_name
        global project_sample_oa_library_names_list
        global project_sample_oa_library_directory_name
        global project_sample_runscript_file_name
        global project_lvs_report_extensions
        global project_extract_file_extension
        global project_sample_runscript_location_dir_name

        def __init__(self, msip_ese_object):
            """
            Initial function of the class
            :param msip_ese_object:
            """

            self.msip_ese_object = msip_ese_object

        def setup_target_project_name(self):
            """
            The function is returning project name value
            :return:
            """

            if self.msip_ese_object.get_target_project_name is not None:
                print_to_stdout(self.msip_ese_object, "FOUND TARGET PROJECT NAME\t" + str(self.msip_ese_object.get_target_project_name))
                return None
            elif self.msip_ese_object.excel_setup[available_excel_options[15]] is not None:
                self.msip_ese_object.set_target_project_name(self.msip_ese_object.excel_setup[available_excel_options[15]])
                print_to_stdout(self.msip_ese_object, "FOUND TARGET PROJECT NAME\t" + str(self.msip_ese_object.get_target_project_name))
            else:
                print_to_stderr(self.msip_ese_object, "Cannot find target project name. Please check script/excel file inputs")

            return None

        def setup_reference_project_name(self):
            """
            The function is returning project name value
            :return:
            """

            if self.msip_ese_object.get_reference_project_name is not None:
                print_to_stdout(self.msip_ese_object, "FOUND REFERENCE PROJECT NAME\t" + str(self.msip_ese_object.get_reference_project_name))
                return None
            elif self.msip_ese_object.excel_setup[available_excel_options[17]] is not None:
                self.msip_ese_object.set_reference_project_name(self.msip_ese_object.excel_setup[available_excel_options[17]])
                print_to_stdout(self.msip_ese_object, "FOUND REFERENCE PROJECT NAME\t" + str(self.msip_ese_object.get_reference_project_name))
            else:
                print_to_stdout(self.msip_ese_object, "WARNING!:\tCannot find reference project name. Please check script/excel file inputs")

            return None

        def setup_target_project_release(self):
            """
            The function is returning project release value
            :return:
            """

            if self.msip_ese_object.get_target_project_release is not None:
                print_to_stdout(self.msip_ese_object, "FOUND TARGET PROJECT RELEASE\t" + str(self.msip_ese_object.get_target_project_release))
                return None
            elif self.msip_ese_object.excel_setup[available_excel_options[16]] is not None:
                self.msip_ese_object.set_target_project_release(self.msip_ese_object.excel_setup[available_excel_options[16]])
                print_to_stdout(self.msip_ese_object, "FOUND TARGET PROJECT RELEASE\t" + str(self.msip_ese_object.get_target_project_release))
            else:
                print_to_stderr(self.msip_ese_object, "Cannot find target project release. Please check script/excel file inputs")

            return None

        def setup_reference_project_release(self):
            """
            The function is returning project release value
            :return:
            """

            if self.msip_ese_object.get_reference_project_release is not None:
                print_to_stdout(self.msip_ese_object, "FOUND REFERENCE PROJECT RELEASE\t" + str(self.msip_ese_object.get_reference_project_release))
                return None
            elif self.msip_ese_object.excel_setup[available_excel_options[18]] is not None:
                self.msip_ese_object.set_reference_project_release(self.msip_ese_object.excel_setup[available_excel_options[18]])
                print_to_stdout(self.msip_ese_object, "FOUND REFERENCE PROJECT RELEASE\t" + str(self.msip_ese_object.get_reference_project_release))
            else:
                print_to_stderr(self.msip_ese_object, "Cannot find reference project name. Please check script/excel file inputs")

            return None

        def find_project_type(self, project_name):
            """
            The function is returning project type
            :param project_name:
            :return:
            """

            projects_root_dir = self.msip_ese_object.get_projects_root_dir

            project_all_types = get_directory_items_list(projects_root_dir)

            if project_all_types is not None:
                for project_type in project_all_types:
                    available_project_names = get_directory_items_list(os.path.join(projects_root_dir, project_type))
                    for value in available_project_names:
                        if project_name == value:
                            return project_type
                print_to_stderr(self.msip_ese_object, "Cannot find project '" + project_name + "' project type under directory\t'" + str(projects_root_dir) + "'")
            else:
                print_to_stderr(self.msip_ese_object, "Cannot find project '" + project_name + "' project type under directory\t'" + str(projects_root_dir) + "'")

        def setup_target_project_type(self):
            """
            The function setup project type
            :return:
            """

            if self.msip_ese_object.get_target_project_type is None:
                if self.msip_ese_object.get_target_project_name is not None:
                    project_type = str(self.find_project_type(self.msip_ese_object.get_target_project_name))
                    self.msip_ese_object.set_target_project_type(project_type)
                    print_to_stdout(self.msip_ese_object, "FOUND TARGET PROJECT TYPE\t" + str(self.msip_ese_object.get_target_project_type))
                else:
                    print_to_stderr(self.msip_ese_object, "Target project name is not defined")

        def setup_reference_project_type(self):
            """
            The function setup project type
            :return:
            """

            if self.msip_ese_object.get_reference_project_type is None:
                if self.msip_ese_object.get_reference_project_name is not None:
                    project_type = str(self.find_project_type(self.msip_ese_object.get_reference_project_name))
                    self.msip_ese_object.set_reference_project_type(project_type)
                    print_to_stdout(self.msip_ese_object, "FOUND REFERENCE PROJECT TYPE\t" + str(self.msip_ese_object.get_reference_project_type))
                else:
                    print_to_stdout(self.msip_ese_object, "Reference project name is not defined")

        @staticmethod
        def get_metal_stack_dir_list(directory_path):
            """
            The function is returning all directory name which under cad directory and contains env.tcl
            Generally such dirs are metal stacks
            :param directory_path:
            :return:
            """

            all_metal_stacks = []

            all_directories = get_directory_items_list(directory_path)
            for metal_stack in all_directories:
                if check_for_file_existence(os.path.join(directory_path, metal_stack), project_environment_file_name):
                    all_metal_stacks.append(metal_stack)

            return all_metal_stacks

        def setup_target_project_metal_stack_list(self):
            """
            The function setup project metal stack list
            :return:
            """

            if get_list_length(self.msip_ese_object.get_target_project_metal_stack_list) < 1:
                self.msip_ese_object.set_target_project_metal_stack_list(self.get_metal_stack_dir_list(os.path.join(self.msip_ese_object.get_projects_root_dir,
                                                                                                                    self.msip_ese_object.get_target_project_type,
                                                                                                                    self.msip_ese_object.get_target_project_name,
                                                                                                                    self.msip_ese_object.get_target_project_release,
                                                                                                                    project_cad_directory_name)))
                print_to_stdout(self.msip_ese_object, "FOUND TARGET PROJECT METAL STACKS LIST:\t" + str(self.msip_ese_object.get_target_project_metal_stack_list))
            else:
                print_to_stdout(self.msip_ese_object, "FOUND TARGET PROJECT METAL STACKS LIST:\t" + str(self.msip_ese_object.get_target_project_metal_stack_list))

        def setup_reference_project_metal_stack_list(self):
            """
            The function setup project metal stack list
            :return:
            """

            if get_list_length(self.msip_ese_object.get_reference_project_metal_stack_list) < 1:
                self.msip_ese_object.set_reference_project_metal_stack_list(self.get_metal_stack_dir_list(os.path.join(self.msip_ese_object.get_projects_root_dir,
                                                                                                                       self.msip_ese_object.get_reference_project_type,
                                                                                                                       self.msip_ese_object.get_reference_project_name,
                                                                                                                       self.msip_ese_object.get_reference_project_release,
                                                                                                                       project_cad_directory_name)))
                print_to_stdout(self.msip_ese_object, "FOUND REFERENCE PROJECT METAL STACKS LIST:\t" + str(self.msip_ese_object.get_reference_project_metal_stack_list))
            else:
                print_to_stdout(self.msip_ese_object, "FOUND REFERENCE PROJECT METAL STACKS LIST:\t" + str(self.msip_ese_object.get_reference_project_metal_stack_list))

        def setup_environment(self):
            """
            The function setups the Environment setup class
            :return:
            """

            print_to_stdout(self.msip_ese_object, "new line")
            print_to_stdout(self.msip_ese_object, "SEARCHING FOR PROJECT INFO\n")

            self.setup_target_project_name()
            self.setup_target_project_release()
            self.setup_target_project_type()
            self.setup_target_project_metal_stack_list()

            self.setup_reference_project_name()
            if self.msip_ese_object.check_for_reference_project_execution():
                self.setup_reference_project_release()
                self.setup_reference_project_type()
                self.setup_reference_project_metal_stack_list()

        @staticmethod
        def generate_ude_command(project_type, project_name, project_release, project_metal_stack, run_directory):
            """
            The function is generating string, which contains ude command
            :return:
            """

            ude_command = """#!/bin/bash --norc

VERIFICATION_PATH={0}
source /remote/cad-rep/etc/.bashrc

module unload msip_cd_pv
module load msip_cd_pv/{1}

ude \\
    --projectType {2} \\
    --projectName {3} \\
    --releaseName {4} \\
    --metalStack  {5} \\
    --nograph \\
    --sourceShellFile {0}/sourceme \\
    -- log            {0}/cdesigner.log \\
    --command "source {0}/command.tcl"
""".format(run_directory, "2016.02", project_type, project_name, project_release, project_metal_stack)

            return ude_command

        @staticmethod
        def gen_ude_tcl_command(tool_name, run_directory, sample_library_name):
            """
            The function is generating ude tcl command
            :param tool_name:
            :param run_directory:
            :param sample_library_name:
            :return:
            """

            target_library_path = create_directories_hierarchy(run_directory, ["LIB", sample_library_name])

            tcl_command = str("""set newSampleLibrary [dm::createLib SampleLibrary -path $RUN_DIR]

db::attachTech $newSampleLibrary -refLibName devices
db::attachTech $newSampleLibrary -refLibName devices

set newSampleCell [dm::createCell SampleExtract -libName SampleLibrary]
set layoutCell [dm::createCellView layout -cell $newSampleCell -viewType maskLayout]
set schematicCell [dm::createCellView schematic -cell $newSampleCell -viewType schematic]

set layoutDesign [de::open $layoutCell]
le::createInst rpp -libName devices -cellName rpp -viewName layout -design [ed] -origin {0 0}
de::save [de::getContexts]
de::close [de::getContexts]

set schematicDesign [de::open $schematicCell]
se::createInst rpp -libName devices -cellName rpp -viewName symbol -design [ed] -origin {0 0}
de::save [de::getContexts]
de::close [de::getContexts]""").replace("$RUN_DIR", target_library_path)

            # Before used "dm::addToLibDefs {0} -path {2}/{0}" + below
            tcl_command += str("\nMSIP_PV::runBatchList lpe RCXT {0} SampleExtract layout {1} {2}/config").format(sample_library_name, tool_name, run_directory)

            return tcl_command

        @staticmethod
        def gen_config_command(run_directory, output_directory):
            """
            The function is generating ude config command
            :param run_directory:
            :param output_directory:
            :return:
            """

            config_command = """#set rcxtTypes "1 1 1 1 1 1 1 1"
#set cornerVal "SigCmax SigCmin SigRCmax SigRCmin SigCmaxDP_ErPlus SigCminDP_ErMinus SigRCmaxDP_ErPlus SigRCminDP_ErMinus FuncCmax FuncCmin FuncRCmax FuncRCmin'
#FuncCmaxDP_ErPlus FuncCminDP_ErMinus FuncRCmaxDP_
#set extractedNetlistPProcessor "0"
set rundir "{0}"
set outDir "{1}" \n""".format(run_directory, output_directory)

            return config_command

        @staticmethod
        def gen_sourceme_command():
            """
            The function is generating ude sourceme command
            :return:
            """

            # sourceme_command = "module unload msip_cd_pv\nmodule load msip_cd_pv"

            return ""

        def remove_existing_sample_library(self, project_type, project_name, project_release, lib_path):
            """
            The function is checking for sample file present and removing if it is exist
            :return:
            """

            print_to_stdout(self.msip_ese_object, "Removing Sample Library from Run directory if it is exist, and creating new one")

            user_home_directory = os.environ["HOME"]
            project_lib_defs_file = os.path.join(user_home_directory, "cd_lib", project_type, project_name, project_release, "design")
            lib_defs_file_object = open_file_for_reading(project_lib_defs_file, "lib.defs")
            new_lib_defs_content = ""
            for line in lib_defs_file_object.readlines():
                if "DEFINE SampleLibrary " not in line:
                    new_lib_defs_content += line

            lib_defs_file_object.close()
            lib_defs_file_object = open_file_for_writing(project_lib_defs_file, "lib.defs")
            lib_defs_file_object.write(new_lib_defs_content)
            lib_defs_file_object.close()

            print_to_stdout(self.msip_ese_object, "Removing sample library directory\t" + os.path.join(lib_path, "LIB"))

            try:
                shutil.rmtree(os.path.join(lib_path, "LIB"))
            except IOError:
                exit("Cannot remove directory\t" + os.path.join(lib_path, "LIB"))

        def generate_sample_environment(self, pex_tool_name, sample_library_name, project_type, project_name, project_release, project_metal_stack, command_run_directory,
                                        output_directory):
            """
            The function is generating sample environment (generating sample runscript file, based on sample OA library), using PV Batch
            :param project_metal_stack:
            :param sample_library_name:
            :param project_type:
            :param project_name:
            :param project_release:
            :param command_run_directory:
            :param output_directory:
            :param pex_tool_name"
            :return:
            """

            print_to_stdout(self.msip_ese_object, str("Generating Project Sample environment for project:\n"
                                                      "\tPROJECT TYPE:{0}\n\tPROJECT NAME:{1}\n\tPROJECT RELEASE:{2}\n"
                                                      "\tPROJECT METAL STACK:{3}").format(str(set_number_of_tabs("PROJECT TYPE:", 3) * "\t" + project_type),
                                                                                          str(set_number_of_tabs("PROJECT NAME:", 3) * "\t" + project_name),
                                                                                          str(set_number_of_tabs("PROJECT RELEASE:", 3) * "\t" + project_release),
                                                                                          str(set_number_of_tabs("PROJECT METAL STACK:", 3) * "\t" + project_metal_stack)))

            ude_command = self.generate_ude_command(project_type, project_name, project_release, project_metal_stack, command_run_directory)
            ude_command_file_object = open_file_for_writing(command_run_directory,
                                                            "execute_ude_" + project_type + "_" + project_name + "_" + project_release + "_" + project_metal_stack)
            ude_command_file_object.writelines(ude_command)
            ude_command_file_object.close()

            ude_tcl_command_file_object = open_file_for_writing(command_run_directory, "command.tcl")
            tcl_command = self.gen_ude_tcl_command(pex_tool_name, command_run_directory, sample_library_name)
            ude_tcl_command_file_object.writelines(tcl_command)
            ude_tcl_command_file_object.close()

            ude_sourceme_command_file_object = open_file_for_writing(command_run_directory, "sourceme")
            sourceme_command = self.gen_sourceme_command()
            ude_sourceme_command_file_object.writelines(sourceme_command)
            ude_sourceme_command_file_object.close()

            ude_config_command_file_object = open_file_for_writing(command_run_directory, "config")
            config_command = self.gen_config_command(command_run_directory, output_directory)
            ude_config_command_file_object.writelines(config_command)
            ude_config_command_file_object.close()

            self.remove_existing_sample_library(project_type, project_name, project_release, command_run_directory)

            process = execute_external_command(
                os.path.join(command_run_directory, "execute_ude_" + project_type + "_" + project_name + "_" + project_release + "_" + project_metal_stack))

            return process

        def extract_sample_cell(self, pex_tool_name, project_type, project_name, project_release, metal_stack, run_dir):
            """
            The function is extracting sample cell
            :param project_release:
            :param project_name:
            :param project_type:
            :param metal_stack:
            :param run_dir:
            :param pex_tool_name:
            :return:
            """

            for sample_library_name in project_sample_oa_library_names_list:
                create_directory(run_dir, sample_library_name)
                target_dir = os.path.join(run_dir, sample_library_name)
                print_to_stdout(self.msip_ese_object, "GENERATING SAMPLE LIBRARY EXTRACTION FOR METAL STACK:\t" + str(metal_stack))
                # untar_zip_package(os.path.join(self.msip_ese_object.get_data_directory, project_sample_oa_library_directory_name, sample_library_name + tar_file_extension),
                #                   target_dir)
                process = self.generate_sample_environment(str(pex_tool_name).lower(), sample_library_name, project_type, project_name, project_release, metal_stack, target_dir,
                                                           target_dir)
                print_to_stdout(self.msip_ese_object, "Executing sample extraction command\t" + target_dir)
                time.sleep(sample_process_wait_time * 60)  # 3 * 60sec = 3 minutes
                report_text_if_long_run = str("\n\tThe Sample Runscript execution is take more than "
                                              "" + str(sample_process_wait_time) + " min. ESE flow is killed the sample runscript execution. Please check what is caused the issue"
                                                                                   "\n\tPath of the command file is:\t" + target_dir + "\n" + str(process.stdout.read()))
                if not process_timeout(process, report_text_if_long_run):
                    print_to_stdout(self.msip_ese_object, "\nEnvironment executed successfully\n")
                else:
                    print_to_stderr(self.msip_ese_object, report_text_if_long_run)

        def run_all_sample_extracts(self):
            """
            The function is executing sample extract
            :return:
            """

            print_to_stdout(self.msip_ese_object, "new line")
            print_to_stdout(self.msip_ese_object, "RUNNING SAMPLE EXTRACT STEP\n")

            all_target_metal_stack = self.msip_ese_object.get_target_project_metal_stack_list

            # Creating running directory
            for metal_stack in all_target_metal_stack:
                target_run_path = create_directories_hierarchy(self.msip_ese_object.get_script_run_directory, [self.msip_ese_object.get_target_project_type,
                                                                                                               self.msip_ese_object.get_target_project_name,
                                                                                                               self.msip_ese_object.get_target_project_release,
                                                                                                               metal_stack,
                                                                                                               project_extract_directory_name])

                self.extract_sample_cell(self.msip_ese_object.get_target_project_pex_tool_name,
                                         self.msip_ese_object.get_target_project_type,
                                         self.msip_ese_object.get_target_project_name,
                                         self.msip_ese_object.get_target_project_release,
                                         metal_stack,
                                         target_run_path)

            if self.msip_ese_object.check_for_reference_project_execution():
                all_reference_metal_stack = self.msip_ese_object.get_reference_project_metal_stack_list

                for metal_stack in all_reference_metal_stack:
                    reference_run_path = create_directories_hierarchy(self.msip_ese_object.get_script_run_directory, [self.msip_ese_object.get_reference_project_type,
                                                                                                                      self.msip_ese_object.get_reference_project_name,
                                                                                                                      self.msip_ese_object.get_reference_project_release,
                                                                                                                      metal_stack,
                                                                                                                      project_extract_directory_name])

                    self.extract_sample_cell(self.msip_ese_object.get_reference_project_pex_tool_name,
                                             self.msip_ese_object.get_reference_project_type,
                                             self.msip_ese_object.get_reference_project_name,
                                             self.msip_ese_object.get_reference_project_release,
                                             metal_stack,
                                             reference_run_path)

        def check_for_ude_extract_flow_correctness(self, run_directory, tool_name):
            """
            The script is checking for .LVS_ERRORS file correctness and extract file existence
            :param run_directory:
            :param tool_name:
            :return:
            """

            lvs_correctness = False
            extract_correctness = False

            print_to_stdout(self.msip_ese_object, "Checking directory:\t" + str(run_directory))

            for root_path, dirs, files in os.walk(run_directory):
                for lvs_report in files:
                    if lvs_report.endswith(project_lvs_report_extensions[tool_name]):
                        lvs_report_file = os.path.join(root_path, lvs_report)
                        if get_file_size(lvs_report_file) > 0:
                            lvs_correctness = True

                for extract_report in files:
                    if extract_report.endswith(project_extract_file_extension):
                        extract_file = os.path.join(root_path, extract_report)
                        if get_file_size(extract_file) > 0:
                            extract_correctness = True

            if lvs_correctness and extract_correctness:
                return True
            else:
                return False

        def get_sample_runscript(self, path_to_search):
            """
            The function is returning sample runscript file
            :param path_to_search:
            :return:
            """

            print_to_stdout(self.msip_ese_object, "Searching sample_runscript file under directory:\t" + str(path_to_search))

            for root_path, dirs, files in os.walk(path_to_search):
                for sample_file_name in files:
                    if sample_file_name == project_sample_runscript_file_name:
                        sample_file = os.path.join(root_path, sample_file_name)
                        print_to_stdout(self.msip_ese_object, "Found sample file:\t" + str(sample_file))
                        return sample_file

            print_to_stdout(self.msip_ese_object, "Warning!!: No any sample file found")
            return None

        @staticmethod
        def change_module_load_line(line):
            """
            The function is changing module load line to have the version in variable
            :param line:
            :return:
            """

            if "module load " in line:
                line_list = line.split()
                if get_list_length(line_list) > 2:
                    if line_list[0][0] == "#":
                        # Checking if the line is not starting with comment
                        return line
                    else:
                        module_original_name = str(line_list[2]).split("/")[0]
                        module_name = module_original_name + '_VERSION'
                        try:
                            module_version = "/" + str(line_list[2]).split("/")[1]
                        except IndexError:
                            module_version = ""
                        line = "export " + module_name + '="' + module_version + '"\n'
                        line += "module load " + module_original_name + "$" + module_name + "\n"
                        return line
                else:
                    return line
            else:
                return line

        @staticmethod
        def get_gds_file_name_from_line(line):
            """
            The function is returning lane with layer map file
            :return:
            """

            gds_name = "NONE"

            line_l = str(line).split()
            layer_gds_command = "-gds"
            if layer_gds_command in line:
                layer_gds_file_index = line_l.index(layer_gds_command) + 1
                gds_name = line_l[layer_gds_file_index]

            return gds_name

        @staticmethod
        def get_lvs_file_name_from_line(line):
            """
            The function is returning lane with layer map file
            :return:
            """

            lvs_file = "NONE"

            line_l = str(line).split()
            command = "-sp"
            if command in line:
                layer_gds_file_index = line_l.index(command) + 1
                lvs_file = line_l[layer_gds_file_index]

            return lvs_file

        @staticmethod
        def replace_source_me_file_name_from_line(line):
            """
            The function is returning lane with layer map file
            :return:
            """

            if "source" in line:
                if "global" not in line:
                    line_l = str(line).split()
                    return 'export PEX_SOURCE_ME=' + line_l[1] + '\n' + line.replace(line_l[1], "$PEX_SOURCE_ME")
                else:
                    return 'export PEX_SOURCE_ME=""\n' + line
            else:
                return line

        @staticmethod
        def replace_gen_ev_file_name_from_line(line):
            """
            The function is returning lane with layer map file
            :return:
            """

            deck = "NONE"
            option_file = "NONE"
            stream_map = "NONE"
            line_l = str(line).split()

            if "-foundry-rule" in line:
                deck = line_l[line_l.index("-foundry-rule") + 1]
            if "-options-file" in line:
                option_file = line_l[line_l.index("-options-file") + 1]
            if "-stream-map" in line:
                stream_map = line_l[line_l.index("-stream-map") + 1]

            target_line = 'export PEX_DECK="' + deck + '"\n' + 'export PEX_OPTION_FILE="' + option_file + '"\n' + 'export STREAM_FILE="' + stream_map + '"\n'
            return target_line + line.replace(deck, "$PEX_DECK").replace(option_file, "$PEX_OPTION_FILE").replace(stream_map, "$STREAM_FILE") + "\n"

        @staticmethod
        def replace_command_line(line, gds_name, lvs_netlist, layer_map_file, source_me_file, option_file, run_dir, output_dir):
            """
            The function is changing the line
            :param option_file:
            :param source_me_file:
            :param line:
            :param gds_name:
            :param lvs_netlist:
            :param layer_map_file:
            :param run_dir:
            :param output_dir:
            :return:
            """

            line = line.replace(lvs_netlist, "$LVS_NETLIST").replace(gds_name + ".gz", "$GDS_FILE").replace(gds_name, "$GDS_FILE")
            line = line.replace(layer_map_file, "$LAYER_MAP_FILE").replace(source_me_file, "$PEX_SOURCE_ME").replace(option_file, "$OPTION_FILE")
            return line.replace(run_dir, "${RUN_DIR}").replace(output_dir, "$OUTPUT_DIR").replace(project_sample_oa_cell_name, "${TOP_CELL_NAME}")

        @staticmethod
        def replace_output_directory(line, run_dir, lvs_netlist):

            line_l = line.split()
            if "-output" in line:
                output_dir = get_file_path(line_l[line_l.index("-output") + 1])
                return ['export OUTPUT_DIR="";\n' + line.replace(lvs_netlist, "${LVS_NETLIST}").replace(run_dir, "${RUN_DIR}").replace(output_dir, "${OUTPUT_DIR}"), output_dir]
            else:
                return ["", "NONE"]

        @staticmethod
        def parse_star_cmd_line(line):
            """
            The function is replacing star rc deck and starcmd files to global variable
            :param line:
            :return:
            """

            target_line = ""
            line_list = line.split("\n")
            for line_string in line_list:
                if "gen_starcmd" in line_string:
                    line_string_list = line_string.split()
                    starcmd = line_string_list[line_string_list.index("-cf") + 1]
                    target_line += 'export STARCMD="' + starcmd + '"\n'
                    tcad = line_string_list[line_string_list.index("-tcad") + 1]
                    target_line += 'export TCAD="' + tcad + '"\n'
                    target_line += line_string.replace(starcmd, "$STARCMD").replace(tcad, "$TCAD")
                else:
                    target_line += line_string + "\n"

            return target_line

        def update_environment_sample_runscript_files(self, file_item, path_to_place, tool_name):
            """
            The function is updating sample run script file in the environment
            :param file_item:
            :param path_to_place:
            :param tool_name:
            :return:
            """

            gds_recognition_line = "exportStream"
            lvs_recognition_line = "nettran"
            source_me_recognition_line = "source"
            star_cmd_recognition_line = "gen_starcmd"

            gen_ev_recognition_line = "gen_" + str(tool_name).lower()

            unneeded_line_list = ["gzip", "COMPRESS_GDS", "export_stream", "mkdir", "rm -fR"]
            gds_file = "NONE"
            lvs_file = "NONE"
            layer_map_file = "NONE"
            source_me_file = "NONE"
            option_file = "NONE"
            output_dir = "NONE"

            file_name = get_file_name_from_path(file_item)

            if file_name is not None:
                file_object = open_file_for_reading(get_file_path(file_item), file_name)
                target_sample_command_file_object = open_file_for_writing(path_to_place, file_name)
                for line in file_object.readlines():
                    line = self.change_module_load_line(line)

                    if gds_recognition_line in line:
                        gds_file = self.get_gds_file_name_from_line(line)
                        line = "\n"

                    if lvs_recognition_line in line:
                        lvs_file = self.get_lvs_file_name_from_line(line)

                    if source_me_recognition_line in line:
                        line = self.replace_source_me_file_name_from_line(line)

                    if gen_ev_recognition_line in line:
                        line = self.replace_gen_ev_file_name_from_line(line)

                    if star_cmd_recognition_line in line:
                        command_result = self.replace_output_directory(line, get_file_path(file_item), lvs_file)
                        line = command_result[0]
                        line = self.parse_star_cmd_line(line)
                        output_dir = command_result[1]

                    enable_writing = True
                    for unneeded_line in unneeded_line_list:
                        if str(unneeded_line).upper() in line.upper():
                            enable_writing = False

                    if enable_writing:
                        if "export METAL_STACK" in line:
                            target_sample_command_file_object.writelines(
                                "\n\n" + line + 'export RUN_DIR="";\nexport TOP_CELL_NAME="";\nexport GDS_FILE="";\nexport LVS_NETLIST="";\n')
                        else:
                            line_for_writing = self.replace_command_line(line, gds_file, lvs_file, layer_map_file, source_me_file, option_file, get_file_path(file_item),
                                                                         output_dir)
                            target_sample_command_file_object.writelines(line_for_writing)

        def grab_all_sample_run_scripts(self):
            """
            The main function of the ProjectEnvironment Class
            :return:
            """

            print_to_stdout(self.msip_ese_object, "new line")
            print_to_stdout(self.msip_ese_object, "GRABBING SAMPLE RUNSCRIPT FILES\n")

            all_target_sample_runscript_files = {}

            all_target_metal_stack = self.msip_ese_object.get_target_project_metal_stack_list

            for metal_stack in all_target_metal_stack:
                target_path = os.path.join(self.msip_ese_object.get_script_run_directory,
                                           self.msip_ese_object.get_target_project_type,
                                           self.msip_ese_object.get_target_project_name,
                                           self.msip_ese_object.get_target_project_release,
                                           metal_stack,
                                           project_extract_directory_name
                                           )
                if self.check_for_ude_extract_flow_correctness(target_path, self.msip_ese_object.get_target_project_pex_tool_name):
                    sample_runscript_file = self.get_sample_runscript(target_path)
                    if sample_runscript_file is not None:
                        if get_file_size(sample_runscript_file) > 0:
                            all_target_sample_runscript_files[metal_stack] = sample_runscript_file
                        else:
                            all_target_sample_runscript_files[metal_stack] = None
                    else:
                        all_target_sample_runscript_files[metal_stack] = None
                else:
                    print_to_stderr(self.msip_ese_object, "Something wrong with sample cell extraction step\n\t\t"
                                                          "No LVS result or SPF file exist. Please check\n\t\t'" + str(target_path) + "'")

            for metal_stack in all_target_metal_stack:
                destination_path = create_directories_hierarchy(self.msip_ese_object.get_data_directory, [project_sample_runscript_location_dir_name,
                                                                                                          self.msip_ese_object.get_target_project_type,
                                                                                                          self.msip_ese_object.get_target_project_name,
                                                                                                          self.msip_ese_object.get_target_project_release,
                                                                                                          metal_stack,
                                                                                                          project_extract_directory_name])
                self.update_environment_sample_runscript_files(file_item=all_target_sample_runscript_files[metal_stack],
                                                               path_to_place=destination_path,
                                                               tool_name=self.msip_ese_object.get_target_project_pex_tool_name)

            if self.msip_ese_object.check_for_reference_project_execution():
                all_reference_sample_runscript_files = {}
                all_reference_metal_stack = self.msip_ese_object.get_reference_project_metal_stack_list

                for metal_stack in all_reference_metal_stack:
                    reference_path = os.path.join(self.msip_ese_object.get_script_run_directory,
                                                  self.msip_ese_object.get_reference_project_type,
                                                  self.msip_ese_object.get_reference_project_name,
                                                  self.msip_ese_object.get_reference_project_release,
                                                  metal_stack,
                                                  project_extract_directory_name
                                                  )
                    if self.check_for_ude_extract_flow_correctness(reference_path, self.msip_ese_object.get_reference_project_pex_tool_name):
                        sample_runscript_file = self.get_sample_runscript(reference_path)
                        if sample_runscript_file is not None:
                            if get_file_size(sample_runscript_file) > 0:
                                all_reference_sample_runscript_files[metal_stack] = sample_runscript_file
                            else:
                                all_reference_sample_runscript_files[metal_stack] = None
                        else:
                            all_reference_sample_runscript_files[metal_stack] = None
                    else:
                        print_to_stderr(self.msip_ese_object, "Something wrong with sample cell extraction step\n\t\t"
                                                              "No LVS result or SPF file exist. Please check\n\t\t'" + str(reference_path) + "'")

                for metal_stack in all_reference_metal_stack:
                    destination_path = create_directories_hierarchy(self.msip_ese_object.get_data_directory, [project_sample_runscript_location_dir_name,
                                                                                                              self.msip_ese_object.get_reference_project_type,
                                                                                                              self.msip_ese_object.get_reference_project_name,
                                                                                                              self.msip_ese_object.get_reference_project_release,
                                                                                                              metal_stack,
                                                                                                              project_extract_directory_name])
                    self.update_environment_sample_runscript_files(file_item=all_reference_sample_runscript_files[metal_stack],
                                                                   path_to_place=destination_path,
                                                                   tool_name=self.msip_ese_object.get_reference_project_pex_tool_name)

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

            self.msip_ese_object.set_executed_flow("")

            # Setting script inputs
            for script_option_name in script_inputs_all_options:
                script_option_value = script_inputs_argument_hash[script_option_name]
                print("\t" + script_option_name + "\t" * set_number_of_tabs(script_option_name, 4) + script_option_value)
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
                    self.msip_ese_object.set_projects_root_dir(script_option_value)
                elif script_option_name == available_script_options[8]:
                    self.msip_ese_object.enable_force_add_test_case()
                elif script_option_name == available_script_options[9]:
                    self.msip_ese_object.set_executed_flow(script_option_value)

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

            print_to_stdout(self.msip_ese_object, "READING EXCEL FILE:\t'" + excel_file + "'")
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
                        str_to_display = "IMPORTANT NOTE! USER MAKES COMMENT FOR TEST CASE OPTION IN EXCEL FILE\n\tCOMMENT:\t" + excel_option_comment + "\n\tLINE:\t\t" + str(
                            current_row + 1)
                        print(str_to_display)
                        print_to_stdout(self.msip_ese_object, str_to_display)

        def get_information_from_excel_file(self, excel_file):
            """
            Main function of the class
            :param excel_file:
            :return:
            """

            if excel_file is None:
                print_to_stdout(self.msip_ese_object, "No any excel file selected.\nSkip the step")
            else:
                if check_for_file_existence(get_file_path(excel_file), get_file_name_from_path(excel_file)):
                    self.read_excel(excel_file)
                else:
                    print_to_stderr(self.msip_ese_object, "Cannot read excel file\t" + os.path.join(get_file_path(excel_file), excel_file))

            print_to_stdout(self.msip_ese_object, "new line")
            print_to_stdout(self.msip_ese_object, "User is set following excel's option(s)\n")
            excel_options = self.msip_ese_object.excel_setup.keys()
            for option in excel_options:
                if self.msip_ese_object.excel_setup[option] is not None:
                    number_of_tabs = "\t" * set_number_of_tabs(option, 5)
                    print_to_stdout(self.msip_ese_object, str("\t" + option + number_of_tabs + self.msip_ese_object.excel_setup[option]))

            print_to_stdout(self.msip_ese_object, "new line")
            print_to_stdout(self.msip_ese_object, "Following excel options are not used\n")
            for option in excel_options:
                if self.msip_ese_object.excel_setup[option] is None:
                    number_of_tabs = "\t" * set_number_of_tabs(option, 5)
                    print_to_stdout(self.msip_ese_object, str("\t" + option + number_of_tabs + str(self.msip_ese_object.excel_setup[option])))

            # Setting target and reference lvs/pex tool name, by default it is ICV

            self.msip_ese_object.set_target_project_pex_tool_name(self.msip_ese_object.excel_setup[available_excel_options[21]])
            self.msip_ese_object.set_reference_project_pex_tool_name(self.msip_ese_object.excel_setup[available_excel_options[23]])

            # Checking if correct test case path

            test_case_path = self.msip_ese_object.excel_setup[available_excel_options[5]]
            if test_case_path is not None:
                if get_string_length(test_case_path) > 0:
                    if not check_for_file_existence(get_file_path(test_case_path), get_file_name_from_path(test_case_path)):
                        if not check_for_dir_existence(get_file_path(test_case_path), get_file_name_from_path(test_case_path)):
                            print_to_stdout(self.msip_ese_object, "WARNING!!:\tWrong file/directory for '{0}' excel option".format(available_excel_options[5]))
                else:
                    print_to_stdout(self.msip_ese_object, "WARNING!!:\tEmpty value for '{0}' excel option".format(available_excel_options[5]))

            # In This line all spaces was removing to make parsing step of the excel file more easy
            for index_number in range(get_list_length(available_excel_options)):
                if self.msip_ese_object.excel_setup[available_excel_options[index_number]] is not None:
                    self.msip_ese_object.excel_setup[available_excel_options[index_number]] = self.msip_ese_object.excel_setup[available_excel_options[index_number]].replace(" ",
                                                                                                                                                                              "")

            return self

    class TestCases:
        """
        The class of test cases updates
        """

        def __init__(self, msip_ese_object):
            self.msip_ese_object = msip_ese_object

        def check_for_excel_file_required_information(self):
            """
            The function is checking for
            :return:
            """

            for required_options_index in [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 15, 16]:
                if self.msip_ese_object.excel_setup[available_excel_options[required_options_index]] is None:
                    print_to_stderr(self.msip_ese_object,
                                    "Required field in excel file is empty:\t'" + str(available_excel_options[required_options_index]) + "'")

        def check_for_test_case_existence(self, path_to_test_case):
            """
            The function is returning the test case
            :return:
            """

            if not self.msip_ese_object.get_force_add_test_case_option:
                if check_for_dir_existence(path_to_test_case, project_test_case_directories_list[1]):
                    all_files = get_directory_items_list(os.path.join(path_to_test_case, project_test_case_directories_list[1]))
                    for file_name in all_files:
                        if file_name.endswith(gds_file_extension):
                            return True

            return False

        def generate_gds_config_file(self, gds_file, untar_directory_path, target_dir):
            """
            The function is generating gds config file
            :param gds_file:
            :param untar_directory_path:
            :param target_dir:
            :return:
            """

            gds_file_name = get_file_name_from_path(gds_file)

            print_to_stdout(self.msip_ese_object, "Generating gds config file for GDS:\t'" + gds_file + "'")
            icwbev_mac_file_content = """layout open GDS_FILE ??
foreach topLevel [layout root cells] {
cell open $topLevel
}
set gds_info [open "TARGET_DIR/GDS_NAME.config" "w+"]
puts $gds_info "TOP_CELL_NAME:\t\t\t [cell active]"
puts $gds_info "ALL_LAYERS:\t\t\t [cell layers -all]"
close $gds_info
exit""".replace(".config", gds_config_file_extension).replace("GDS_FILE", gds_file).replace("GDS_NAME", gds_file_name).replace("TARGET_DIR", target_dir)

            icwb_mac_file_object = open_file_for_writing(untar_directory_path, gds_file_name + ".mac")
            icwb_mac_file_object.write(icwbev_mac_file_content)
            icwb_mac_file_object.close()

            shell_command = """#!/bin/bash
source /remote/cad-rep/etc/.bashrc

export SNPSLMD_LICENSE_FILE=26585@am04-lic3:26585@am04-lic2:26585@de02_lic5:26585@us01snpslmd5:27000@bear
export LM_LICENSE_FILE=1717@de02-lic5:1717@de02_lic4:26585@am04-lic3:26585@am04-lic2:26585@de02_lic5:26585@us01snpslmd5

cd RUN_DIR

module unload icwbev_plus
module load icwbev_plus/2015.06
icwbev -run GDS_NAME.mac -nodisplay\n
chmod -R 777 *
""".replace("RUN_DIR", untar_directory_path).replace("GDS_NAME", gds_file_name)

            shell_file_object = open_file_for_writing(untar_directory_path, gds_file_name + "_export_gds_layers.sh")
            shell_file_object.write(shell_command)
            shell_file_object.close()

            process = execute_external_command(os.path.join(untar_directory_path, gds_file_name + "_export_gds_layers.sh"))
            process.wait()

            print_to_stdout(self.msip_ese_object, "GDS layers are in file\t" + os.path.join(untar_directory_path, gds_file_name + gds_config_file_extension))

        def move_file(self, excel_information, source, destination):
            """
            The function is moving
            :param source:
            :param excel_information:
            :param destination:
            :return:
            """

            excel_information = str(excel_information).replace(available_package_directory_tags_list[0], "").replace(available_package_directory_tags_list[1], "")
            source = os.path.join(source, excel_information)
            if check_for_file_existence(get_file_path(source), get_file_name_from_path(source)):
                if check_for_dir_existence(get_file_path(destination), get_file_name_from_path(destination)):
                    shutil.copy(source, os.path.join(destination, get_file_name_from_path(source)))
                    try:
                        os.chmod(os.path.join(destination, get_file_name_from_path(source)), mode=0o777)
                    except OSError:
                        print_to_stdout(self.msip_ese_object, "Permission denied\t" + os.path.join(destination, get_file_name_from_path(source)))
                    return os.path.join(destination, get_file_name_from_path(source))

            print_to_stderr(self.msip_ese_object, "Cannot copy file from:\t'" + source + "'\tTo\t'" + os.path.join(destination, get_file_name_from_path(source)) + "'")

        def check_config_file_existence(self, target_gds_file):
            """
            The function is checking for config file correctness
            :param target_gds_file:
            :return:
            """

            if check_for_file_existence(get_file_path(target_gds_file), get_file_name_from_path(target_gds_file) + gds_config_file_extension):
                if get_file_size(os.path.join(get_file_path(target_gds_file), get_file_name_from_path(target_gds_file) + gds_config_file_extension)) > 0:
                    return

            print_to_stderr(self.msip_ese_object, "Cannot find gds config file for:\t" + target_gds_file)

        def get_list_from_excel_line(self, line, index_name):
            """
            The function is simply making list from line, by splitting it with comma
            :param line:
            :param index_name
            :return:
            """

            if line is not None:
                try:
                    return line.split(",")
                except IOError:
                    print_to_stderr(self.msip_ese_object, "Cannot parse the line from excel file\t" + index_name + ":\t" + line)
            else:
                return []

        def move_test_case_files(self, source_directory, destination_directory, untar_directory_path):
            """
            The function is moving all necessary data of the test case from source path to environment
            :param source_directory:
            :param destination_directory:
            :param untar_directory_path:
            :return:
            """

            # Moving test bench files
            create_directory(destination_directory, project_test_case_directories_list[3])
            self.move_file(self.msip_ese_object.excel_setup[available_excel_options[6]], source_directory, os.path.join(destination_directory,
                                                                                                                        project_test_case_directories_list[3]))
            # Moving LVS and GDS Files
            gds_files_list = self.get_list_from_excel_line(self.msip_ese_object.excel_setup[available_excel_options[7]], available_excel_options[7])
            create_directory(destination_directory, project_test_case_directories_list[1])
            for gds_file in gds_files_list:
                gds_target_file = self.move_file(gds_file, source_directory, os.path.join(destination_directory, project_test_case_directories_list[1]))
                self.generate_gds_config_file(gds_target_file, untar_directory_path, get_file_path(gds_target_file))
                self.check_config_file_existence(gds_target_file)

            lvs_files_list = self.get_list_from_excel_line(self.msip_ese_object.excel_setup[available_excel_options[8]], available_excel_options[8])
            create_directory(destination_directory, project_test_case_directories_list[2])
            for lvs_file in lvs_files_list:
                self.move_file(lvs_file, source_directory, os.path.join(destination_directory, project_test_case_directories_list[2]))

            # Moving Measure Files
            measure_files_list = self.get_list_from_excel_line(self.msip_ese_object.excel_setup[available_excel_options[11]], available_excel_options[11])
            create_directory(destination_directory, project_test_case_directories_list[4])
            for measure_file in measure_files_list:
                self.move_file(measure_file, source_directory, os.path.join(destination_directory, project_test_case_directories_list[4]))

            # Moving Other Include Files
            include_files_list = self.get_list_from_excel_line(self.msip_ese_object.excel_setup[available_excel_options[12]], available_excel_options[12])
            create_directory(destination_directory, project_test_case_directories_list[5])
            for include_file in include_files_list:
                self.move_file(include_file, source_directory, os.path.join(destination_directory, project_test_case_directories_list[5]))

            # Moving User result files
            user_result_files_list = self.get_list_from_excel_line(self.msip_ese_object.excel_setup[available_excel_options[13]], available_excel_options[13])
            create_directory(destination_directory, project_test_case_directories_list[6])
            for measure_file in user_result_files_list:
                self.move_file(measure_file, source_directory, os.path.join(destination_directory, project_test_case_directories_list[6]))

            user_result_files_list = self.get_list_from_excel_line(self.msip_ese_object.excel_setup[available_excel_options[14]], available_excel_options[14])
            for spf_file in user_result_files_list:
                self.move_file(spf_file, source_directory, os.path.join(destination_directory, project_test_case_directories_list[6]))

            # Copy the excel file into the test bench directory
            create_directory(destination_directory, project_test_case_directories_list[0])
            excel_file = self.msip_ese_object.get_script_excel_file
            shutil.copy(excel_file, os.path.join(destination_directory, project_test_case_directories_list[0], get_file_name_from_path(excel_file)))

        def update_test_cases(self):
            """
            The main function of TestCase class
            """

            if self.msip_ese_object.get_script_excel_file is not None:
                self.check_for_excel_file_required_information()
                print_to_stdout(self.msip_ese_object, "UPDATING TEST CASES STEP")
                test_case_directory = create_directories_hierarchy(self.msip_ese_object.get_test_cases_directory, [self.msip_ese_object.excel_setup[available_excel_options[0]],
                                                                                                                   self.msip_ese_object.excel_setup[available_excel_options[3]]])

                if not self.check_for_test_case_existence(test_case_directory):
                    self.msip_ese_object.enable_force_add_test_case()
                else:
                    self.msip_ese_object.disable_force_add_test_case()

                if self.msip_ese_object.get_force_add_test_case_option:
                    test_case_untar_directory = create_directories_hierarchy(self.msip_ese_object.get_script_run_directory,
                                                                             [self.msip_ese_object.excel_setup[available_excel_options[0]],
                                                                              self.msip_ese_object.excel_setup[available_excel_options[3]],
                                                                              untar_directory_name])
                    if str(self.msip_ese_object.excel_setup[available_excel_options[5]]).endswith(tar_file_extension):
                        untar_zip_package(self.msip_ese_object.excel_setup[available_excel_options[5]], test_case_untar_directory)
                        source_directory_path = test_case_untar_directory
                    else:
                        source_directory_path = str(self.msip_ese_object.excel_setup[available_excel_options[5]])

                    if check_for_dir_existence(get_file_path(source_directory_path), get_file_name_from_path(source_directory_path)):
                        self.move_test_case_files(source_directory_path, test_case_directory, test_case_untar_directory)

    class Extract:
        """
        The Extract class
        """

        global project_sample_runscript_location_dir_name

        def __init__(self, msip_ese_object):
            """
            Initialisation of the class
            """

            self.msip_ese_object = msip_ese_object

        def grab_layer_numbers_from_layer_map(self, layer_map_file):
            """
            The function is returning list of layer numbers
            :param layer_map_file:
            :return:
            """

            all_layers = []

            print_to_stdout(self.msip_ese_object, "Grabbing all layers from layer map:\t" + layer_map_file)

            layer_map_object = open_file_for_reading(get_file_path(layer_map_file), get_file_name_from_path(layer_map_file))
            for line in layer_map_object.readlines():
                line_list = line.split()
                if "#" != line_list[0][0]:
                    if get_list_length(line_list) > 3:
                        all_layers.append(line_list[2] + ":" + line_list[3])

            return all_layers

        def grab_all_layer_numbers_from_layer_map_file(self, project_type, project_name, project_release, metal_stack_list):
            """
            The function is searching for layermap file of the project and parsing all layers
            :return:
            """

            layer_hash = {}

            print_to_stdout(self.msip_ese_object,
                            "Grabbing all layers for each metal stack's layer map file:\n\t" + project_type + "\n\t" + project_name + "\n\t" + project_release)

            # Delete
            for metal_stack in metal_stack_list:
                layer_hash[metal_stack] = []
                sample_runscript_path = os.path.join(self.msip_ese_object.get_data_directory, project_sample_runscript_location_dir_name, project_type, project_name,
                                                     project_release, metal_stack, project_extract_directory_name)
                runscript_file_object = open_file_for_reading(sample_runscript_path, project_sample_runscript_file_name)
                for line in runscript_file_object.readlines():
                    if "export STREAM_FILE=" in line:
                        layer_hash[metal_stack] = self.grab_layer_numbers_from_layer_map(line.split('STREAM_FILE="')[1].replace('"', ""))

        def get_test_cases(self):
            """
            The function is returning hash with test case packages
            :return:
            """

            test_cases_hash = {}

            project_name = self.msip_ese_object.get_reference_project_name
            if project_name is None:
                project_name = self.msip_ese_object.get_target_project_name

            if self.msip_ese_object.get_executed_test_case_package is None:
                all_available_test_cases = get_directory_items_list(self.msip_ese_object.get_test_cases_directory)
                for test_case_name in all_available_test_cases:
                    test_case_directory = os.path.join(self.msip_ese_object.get_test_cases_directory, test_case_name, project_name)
                    if check_for_dir_existence(test_case_directory, project_test_case_directories_list[1]):
                        all_directory_items = get_directory_items_list(os.path.join(test_case_directory, project_test_case_directories_list[1]))
                        for file_name in all_directory_items:
                            if file_name.endswith(gds_file_extension):
                                test_cases_hash[test_case_name] = test_case_directory
            else:
                test_case_path = self.msip_ese_object.get_executed_test_case_package
                try:
                    test_case_name = test_case_path.split("/")[-2]
                except IndexError:
                    print_to_stderr(self.msip_ese_object, "Cannot find test case name from path\t" + str(test_case_path))

                # noinspection PyUnboundLocalVariable
                test_cases_hash[test_case_name] = test_case_path

            if get_list_length(test_cases_hash.keys()) > 0:
                self.msip_ese_object.set_project_test_cases(test_cases_hash)
            else:
                print_to_stderr(self.msip_ese_object, "Cannot find any test case for project:\t'" + project_name + "' inside directory:\t'" + str(
                    self.msip_ese_object.get_test_cases_directory) +
                                "'")

        def create_sample_runscript(self, extract_run_directory, extract_output_dir, test_case_path, file_name, top_cell_name, sample_file_directory):
            """
            The function is generating environment for execution extract
            :param extract_run_directory:
            :param extract_output_dir:
            :param test_case_path:
            :param top_cell_name:
            :param file_name:
            :param sample_file_directory:
            :return:
            """

            file_base_name = get_file_name_from_path(file_name).replace(gds_file_extension, "")

            try:
                os.symlink(os.path.join(test_case_path, project_test_case_directories_list[1], file_name), os.path.join(extract_run_directory, file_name))
            except FileExistsError:
                print_to_stdout(self.msip_ese_object, "Link files already exist:\t" + os.path.join(extract_run_directory, file_name))

            for extension in netlist_file_extension:
                if check_for_file_existence(os.path.join(test_case_path, project_test_case_directories_list[2]), file_base_name + extension):
                    try:
                        os.symlink(os.path.join(test_case_path, project_test_case_directories_list[2], file_base_name + extension),
                                   os.path.join(extract_run_directory, file_base_name + ".cdl"))
                    except FileExistsError:
                        print_to_stdout(self.msip_ese_object, "Link files already exist:\t" + os.path.join(extract_run_directory, file_base_name + ".cdl"))

            try:
                os.symlink(os.path.join(test_case_path, project_test_case_directories_list[1], file_name + gds_config_file_extension),
                           os.path.join(extract_run_directory, file_name + gds_config_file_extension))
            except FileExistsError:
                print_to_stdout(self.msip_ese_object, "Link files already exist:\t" + os.path.join(extract_run_directory, file_name + gds_config_file_extension))

            sample_file_object = open_file_for_reading(sample_file_directory, project_sample_runscript_file_name)
            block_extract_command_file_object = open_file_for_writing(extract_run_directory, file_base_name + "_" + project_extract_directory_name + ".sh")

            for line in sample_file_object.readlines():
                line = line.replace('export RUN_DIR="";', "export RUN_DIR=" + extract_run_directory + ";")
                line = line.replace("export TOP_CELL_NAME=\"\";", "export TOP_CELL_NAME=" + top_cell_name + ";")
                line = line.replace("export GDS_FILE=\"\";", "export GDS_FILE=" + file_base_name + gds_file_extension + ";")
                line = line.replace("export LVS_NETLIST=\"\";", "export LVS_NETLIST=" + file_base_name + ".cdl")
                line = line.replace("export OUTPUT_DIR=\"\"", "export OUTPUT_DIR=" + extract_output_dir)
                line = line.replace("cd ${RUN_DIR};", "cd ${RUN_DIR};\nmkdir STAR_rcc_typical\nmkdir STAR_rc_typical\nmkdir STAR_srccpcc_typical\nmkdir STAR_cc_typical\n")
                block_extract_command_file_object.writelines(line)

            sample_file_object.close()
            block_extract_command_file_object.close()

            self.create_top_cell_subckt_file(top_cell_name, file_base_name + ".cdl", extract_run_directory)

        def get_top_cell_name_and_metal(self, test_case_path, gds_file_name):
            """
            The function is returning top cell name and metal stack
            :return:
            """

            return_variable = ["", "12M_2X_vh_1Ya_v_4Y_hvhv_2Yy2Z"]

            config_file = os.path.join(test_case_path, project_test_case_directories_list[1], gds_file_name + gds_config_file_extension)
            if not check_for_file_existence(os.path.join(test_case_path, project_test_case_directories_list[1]), gds_file_name + gds_config_file_extension):
                print_to_stderr(self.msip_ese_object, "Config file does not exist:\t" + config_file)
            else:
                config_file_object = open_file_for_reading(os.path.join(test_case_path, project_test_case_directories_list[1]), gds_file_name + gds_config_file_extension)
                for line in config_file_object.readlines():
                    if "TOP_CELL_NAME:" in line:
                        return_variable[0] = line.split()[1]

                config_file_object.close()

            return return_variable

        @staticmethod
        def create_top_cell_subckt_file(top_cell_name, lvs_netlist, target_dir):
            """
            The function is generating top cell subckt file for starcmd execution flow
            :return:
            """

            global top_cell_subckt_file_name
            lvs_file_object = open_file_for_reading(target_dir, get_file_name_from_path(lvs_netlist))
            subckt_file_object = open_file_for_writing(target_dir, top_cell_subckt_file_name)

            enable_writing = False

            for line in lvs_file_object.readlines():
                if line.startswith(subckt_start_recognition_word + " " + top_cell_name + " "):
                    enable_writing = True
                elif line.startswith(subckt_end_recognition_word):
                    if enable_writing:
                        subckt_file_object.writelines(line)
                    enable_writing = False

                if enable_writing:
                    subckt_file_object.writelines(line)

            lvs_file_object.close()
            subckt_file_object.close()

        def create_extract_environment(self, test_case_name, test_case_path):
            """
            The function is creating extraction environments
            :return:
            """

            shell_command_path = os.path.join(self.msip_ese_object.get_script_run_directory, test_case_name)

            test_case_target_root_dir = create_directories_hierarchy(self.msip_ese_object.get_script_run_directory,
                                                                     [test_case_name,
                                                                      self.msip_ese_object.get_target_project_name,
                                                                      self.msip_ese_object.get_target_project_release,
                                                                      project_extract_directory_name])

            test_case_target_result_directory = create_directories_hierarchy(self.msip_ese_object.get_results_directory,
                                                                             [test_case_name,
                                                                              self.msip_ese_object.get_target_project_name,
                                                                              self.msip_ese_object.get_target_project_release,
                                                                              project_SPF_directory_name])

            if self.msip_ese_object.check_for_reference_project_execution():
                test_case_reference_root_dir = create_directories_hierarchy(self.msip_ese_object.get_script_run_directory,
                                                                            [test_case_name,
                                                                             self.msip_ese_object.get_reference_project_name,
                                                                             self.msip_ese_object.get_reference_project_release,
                                                                             project_extract_directory_name])

                test_case_reference_result_directory = create_directories_hierarchy(self.msip_ese_object.get_results_directory,
                                                                                    [test_case_name,
                                                                                     self.msip_ese_object.get_reference_project_name,
                                                                                     self.msip_ese_object.get_reference_project_release,
                                                                                     project_SPF_directory_name])

            test_case_all_gds_files = get_directory_items_list(os.path.join(test_case_path, project_test_case_directories_list[1]))

            for file_name in test_case_all_gds_files:
                if file_name.endswith(gds_file_extension):
                    file_abs_name = get_file_name_from_path(file_name)

                    create_directory(test_case_target_root_dir, file_abs_name.upper())
                    create_directory(test_case_target_result_directory, file_abs_name.upper())

                    test_case_target_dir = os.path.join(test_case_target_root_dir, file_abs_name.upper())
                    test_case_target_output_dir = os.path.join(test_case_target_result_directory, file_abs_name.upper())

                    gds_info = self.get_top_cell_name_and_metal(test_case_path, file_name)
                    metal_stack = gds_info[1]

                    self.create_sample_runscript(test_case_target_dir, test_case_target_output_dir, test_case_path, file_name, gds_info[0],
                                                 os.path.join(self.msip_ese_object.get_data_directory,
                                                              project_sample_runscript_location_dir_name,
                                                              self.msip_ese_object.get_target_project_type,
                                                              self.msip_ese_object.get_target_project_name,
                                                              self.msip_ese_object.get_target_project_release,
                                                              metal_stack,
                                                              project_extract_directory_name))

                    # Reference Project Part

                    if self.msip_ese_object.check_for_reference_project_execution():
                        create_directory(test_case_reference_root_dir, file_abs_name.upper())
                        create_directory(test_case_reference_result_directory, file_abs_name.upper())

                        test_case_reference_dir = os.path.join(test_case_reference_root_dir, file_abs_name.upper())
                        test_case_reference_output_dir = os.path.join(test_case_reference_result_directory, file_abs_name.upper())

                        self.create_sample_runscript(test_case_reference_dir, test_case_reference_output_dir, test_case_path, file_name, gds_info[0],
                                                     os.path.join(self.msip_ese_object.get_data_directory,
                                                                  project_sample_runscript_location_dir_name,
                                                                  self.msip_ese_object.get_reference_project_type,
                                                                  self.msip_ese_object.get_reference_project_name,
                                                                  self.msip_ese_object.get_reference_project_release,
                                                                  metal_stack,
                                                                  project_extract_directory_name))

            return shell_command_path

        def create_all_test_cases_extract_environments(self):
            """
            The function is generating all extract environments
            :return:
            """

            test_cases_hash = self.msip_ese_object.get_project_test_cases
            all_test_cases_name = test_cases_hash.keys()

            all_shell_files = []

            for test_case_name in all_test_cases_name:
                all_shell_files.append(self.create_extract_environment(test_case_name, test_cases_hash[test_case_name]))

            return all_shell_files

        def execute_pex(self, test_case_dirs):
            """
            The function is executing all sh commands found under RUN_DIR
            :param test_case_dirs:
            :return:
            """

            for directory in test_case_dirs:
                for root, directories, files in os.walk(directory):
                    for file_name in files:
                        if (file_name.endswith("_" + project_extract_directory_name + ".sh")) and ("UNTAR" not in root):
                            process = execute_external_command(os.path.join(root, file_name))
                            print_to_stdout(self.msip_ese_object, "EXECUTING EXTERNAL PEX COMMAND:\t" + os.path.join(root, file_name))
                            process.wait()

    class Simulation:
        """
        The Simulation instance class. Creating final deck for sim , executing and storing simulation
        """

        def __init__(self, msip_ese_object):
            """
            The initial function of the class
            """

            self.msip_ese_object = msip_ese_object

        def run_simulation(self):
            """
            The function is executing simulations
            :return:
            """

            print_to_stdout(self.msip_ese_object, "Executing Simulation")

    class Report:
        """
        The Reporting class
        """

        def __init__(self, msip_ese_object):
            """
            The initial function of the class
            """

            self.msip_ese_object = msip_ese_object

        def gen_excel_report(self):
            """
            The function is generating excel report
            :return:
            """

            print_to_stdout(self.msip_ese_object, "Generating excel report")

    def main(self):
        """
        Main Function of the MsipEse Class
        :return:
        """

        script_inputs_instance = self.ScriptInputs(self)
        script_arguments = script_inputs_instance.get_script_arguments()
        script_inputs_instance.set_script_inputs(script_arguments)

        print("\nPROCESSING ...\n")

        # Creating environment directories

        print("\tSTEP1:\tTIME:" + get_current_time() + "\tPROCESSING ...\t\t# Reading Script Inputs")

        self.create_script_env_directories()
        # Opening log files for the script
        # The script stdout file object
        self.object_stdout_file = open_file_for_writing(self.script_log_dir, self.object_log_name + ".stdout")
        # The script stderr file object
        self.object_stderr_file = open_file_for_writing(self.script_log_dir, self.object_log_name + ".stderr")

        print_to_stdout(self, "READING SCRIPT ARGUMENTS")
        print_to_stdout(self, "Script Inputs Is:\n" + string_column_decoration(list(script_arguments.keys()), list(script_arguments.values()), 5, 4))

        # All Instances of the Script

        # The initialisation of excel class and reading it
        script_excel_instance = self.Excel(self)

        # The initialisation of ProjectEnvironment class instance
        project_environment = self.ProjectEnvironment(self)

        # The initialisation of TestCases class instance
        test_cases = self.TestCases(self)

        # The initialisation of Extract class instance
        test_cases_extract = self.Extract(self)

        # The initialisation of Simulation class instance
        simulation = self.Simulation(self)

        # The initialisation of Report class
        report = self.Report(self)

        # Setting Target Project Name/Release from Script Input or Excel File
        script_excel_instance.get_information_from_excel_file(self.get_script_excel_file)
        project_environment.setup_environment()

        # Checking for script input correctness
        self.check_script_setup_correctness()
        print("\t\tTIME:" + get_current_time() + "\tCOMPLETED")

        if self.check_if_update_environment():
            print("\tSTEP2:\tTIME:" + get_current_time() + "\tPROCESSING ...\t\t# Checking For Project Environment Update")
            # The sample library extraction part
            project_environment.run_all_sample_extracts()

            # Grabbing and updating in the script environment the sample runscript files
            project_environment.grab_all_sample_run_scripts()
            print("\t\tTIME:" + get_current_time() + "\tCOMPLETED")
        else:
            print("\tSTEP2:\tSkipping STEP 'Checking For Project Environment Update'\tTIME:" + get_current_time())

        if self.check_if_update_test_case():
            print("\tSTEP3:\tTIME:" + get_current_time() + "\tPROCESSING ...\t\t# Checking For Test Case Update")
            # Updating test cases
            test_cases.update_test_cases()
            print("\t\tTIME:" + get_current_time() + "\tCOMPLETED")
        else:
            print("\tSTEP3:\tSkipping STEP 'Checking For Test Case Update'\tTIME:" + get_current_time())

        if self.check_if_execute_pex():
            print("\tSTEP4:\tTIME:" + get_current_time() + "\tPROCESSING ...\t\t# Running PEX on Test Case(s)")
            # Do extraction
            test_cases_extract.get_test_cases()
            test_case_dirs_list = test_cases_extract.create_all_test_cases_extract_environments()
            test_cases_extract.execute_pex(test_case_dirs_list)
            print("\t\tTIME:" + get_current_time() + "\tCOMPLETED")
        else:
            print("\tSTEP4:\tSkipping STEP 'Running PEX on Test Case(s)'\tTIME:" + get_current_time())

        if self.check_if_execute_simulation():
            print("\tSTEP5:\tTIME:" + get_current_time() + "\tPROCESSING ...\t\t# Running SIM on Test Case(s)")
            simulation.run_simulation()
            print("\t\tTIME:" + get_current_time() + "\tCOMPLETED")
        else:
            print("\tSTEP5:\tSkipping STEP 'Running SIM on Test Case(s)'\tTIME:" + get_current_time())

        if self.check_if_execute_report():
            print("\tSTEP6:\tTIME:" + get_current_time() + "\tPROCESSING ...\t\t# Running Reporting step\tTIME:" + get_current_time())
            report.gen_excel_report()
            print("\t\tTIME:" + get_current_time() + "\tCOMPLETED")
        else:
            print("\tSTEP6:\tSkipping STEP 'Running Reporting step'\tTIME:" + get_current_time())


def main():
    """
    The main function of the script
    :return:
    """

    user_script_inputs = ScriptArguments().get_user_all_inputs()

    evaluation_object = MsipEse()
    evaluation_object.set_user_script_arguments(user_script_inputs)
    evaluation_object.main()

    # print_to_stdout(evaluation_object, string_column_decoration(list(evaluation_object.__dict__.keys()), list(evaluation_object.__dict__.values()), 8, 1))


if __name__ == '__main__':
    print("\n\nSTART TIME:\t" + get_current_time() + "\n\n")

    main()

    print("\n\nFINISHED TIME:\t" + get_current_time())
    print("\n\nScript Finished Successfully ^_^\n\n")
