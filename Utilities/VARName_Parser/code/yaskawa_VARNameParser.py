import logging
import argparse
import os
from os import listdir
from os.path import join
import xlrd
import pandas as pd
from collections import OrderedDict


# Owned
__author__ = "Rebecca Iglesias-Flores"
__copyright__ = "Copyright 2021,  "
__credits__ = ["Rebecca Iglesias-Flores",
               "Daniel Luna",
               "Andrew King"]
__license__ = "MIT"
__version__ = "0.1.0"
__maintainer__ = ["Rebecca Iglesias-Flores",
                  "Daniel Luna",
                  "Andrew King"]
__email__ = ["irebecca@seas.upenn.edu"
             "dluna@noveon.co",
             "aking@noveon.co",
             "acking1187@gmail.com"]
__status__ = "Dev"

"""Description of module: yaskawa_parser module reads input .xlsx tracker file from data/ folder and outputs .DAT file that's in the format the robot controller needs."""

"""Usage:
Source Folder: /Users/riglesias-flores/Documents/Noveon_Development/Betatest/
Select a Source Folder of your own choosing (mine is shown above^) and ensure 
there is a 'code/' folder with the yaskawa_parser.py script in it and 
a 'data/' folder with your excel input file.  Then run command line below from
terminal from your Source Folder.


python -u code/yaskawa_VARNameParser.py \
--var_tracker_fname 792164-1-1_CDS_VAR_Rev.0.xlsx \
--dataDir data/ \
--outDir out/ \
--outSubdir text_parser_project/ \
> yaskawa_parser.log 2>&1
"""

'''global variables'''
sourceDir = os.path.realpath('.') + '/'
print(f'\nSource Directory: {sourceDir}')
header_dict = OrderedDict({
                                'B VAR': '///B',
                                'I VAR':'///I',
                                'D VAR':'///D',
                                'R VAR':'///R',
                                'S VAR':'///S',
                                'P VAR':'///P',
                                'TM VAR':'///BP',
                                'EX VAR':'///EX'
                                })
print(header_dict.keys())

def register_arguments():
    """Registers the arguments in the argparser into a global variable.

    Args:
      N/A

    Returns:
      N/A, sets the global args variable
    """

    global args

    parser = argparse.ArgumentParser()

    # Specify command line arguments.
    parser.add_argument(
        '--var_tracker_fname', type=str,
        required=True,
        help="Filename of the excel variable tracker file. ex. '_792164-1-1_CDS_VAR_Rev.0.xlsx'"
        )
    parser.add_argument(
        '--dataDir', type=str,
        required=True,
        help="Name of data directory, tell script where to get input data files from (the variable tracker file is located here)."
        )
    parser.add_argument(
        '--outDir', type=str,
        required=True,
        help="Name of output directory for the .DAT file."
        )
    parser.add_argument(
        '--outSubdir', type=str,
        required=True,
        help="Name of output sub-directory (Subdir), located from the Source Folder in the out/ directory, this is the subdirectory where results will be written to or read from."
        )

    # Parse command line arguments.
    args = parser.parse_args()

    # print command line arguments for this run
    LOGGER.info("---confirm argparser---")
    for arg in vars(args):
        print(arg, getattr(args, arg))

#####################
#### HELPER FUNCTIONS
#####################

def set_logger():
    """Helper function that formats a logger use programmers can easily debug their scripts.

    Args:
      N/A

    Returns:
      logger object

    Note:
      You can refer to this tutorial for more info on how to use logger: https://towardsdatascience.com/stop-using-print-and-start-using-logging-a3f50bc8ab0
    """
    # create a logger object instance
    logger = logging.getLogger()

    # specifies the lowest severity for logging
    logger.setLevel(logging.NOTSET)

    console_handler = logging.StreamHandler()

    # set the logging format for your handler
    log_format = '\n%(asctime)s | Line %(lineno)d in %(filename)s: %(funcName)s() | %(levelname)s: \n%(message)s'
    console_handler.setFormatter(logging.Formatter(log_format))

    logger.addHandler(console_handler)

    return logger

def confirm_directories():


    # confirm script directory
    if not os.path.exists('code/'):
        LOGGER.info('Directory Created: %s', 'code/')
        os.mkdir('code/')

    # confirm dataDir
    if not os.path.exists(args.dataDir):
        LOGGER.info('Directory Created: %s', args.dataDir)
        os.mkdir(args.dataDir)

    # confirm outDir
    if not os.path.exists(args.outDir):
        LOGGER.info("Directory created: %s: ", args.outDir)
        os.mkdir(args.outDir)

    # confirm this run's output folder
    output_folder = args.outDir + args.outSubdir
    if not os.path.exists(output_folder):
        LOGGER.info('Directory Created: %s', output_folder)
        os.mkdir(output_folder)

    # confirm excel infile exists or immediately exit out
    var_tracker_fpath = args.dataDir + args.var_tracker_fname
    try:
        assert os.path.exists(var_tracker_fpath)
    except AssertionError:
        LOGGER.debug("File does not exist!")
        print(f'check for <<{args.var_tracker_fname}>> in dataDir: <<{args.dataDir}>>')
        print('exiting script...')
        exit()

    return var_tracker_fpath

def pad_data(DAT_dict):

    # pad with carriage returns
    padded_lengths = [3,2004,4005,6006,8007,10008,20009,30010] + [40011]

    for section_no, (section_title, data) in enumerate(DAT_dict.items()):
        curr_len = len(data) + 1 # plus 1 to account for section header taking up a line          
        required_padding = padded_lengths[section_no + 1] - padded_lengths[section_no]
        actual_padding_remaining = required_padding - curr_len
        LOGGER.info("Padding current data for %s", section_title)
        print(f'length of curr_data: {curr_len} padding to {padded_lengths[section_no]} = {actual_padding_remaining}')
        ncarriage_returns = ['\r\n']*(actual_padding_remaining)
        data.extend(ncarriage_returns)


def write_to_DAT_file(DAT_dict, lengths, outfile):


    with open(outfile, 'w') as fout:

        # write static header
        fout.write('//VARNAME\n')
        fout.write('///SHARE 2000,2000,2000,2000,2000,10000,10000,10000\n')
        
        # write data
        for sheet_name, data in DAT_dict.items():

            # write section header
            section_title = header_dict[sheet_name]

            fout.write(section_title + '\n') 

            # section data
            for line_no, item in enumerate(data):

                if item == '\r\n':
                    # write padded carriage returns
                    fout.write('\n')
                elif str(item[1]) == 'nan':
                    fout.write('\n')
                else:
                    # write section data
                    # <class 'str'>, <class 'str'>
                    str_int, name = item[0], item[1]
                    fout.write(str_int + " 1,0," + name + '\n')


#########################
###### MAIN
#########################

def run():
    
    # confirm required directories and infiles
    var_tracker_fpath = confirm_directories()

    # read in file
    # dict_keys(['Cover', 'B VAR', 'I VAR', 'D VAR', 'R VAR', 'S VAR', 'P VAR', 'TM VAR', 'TF VAR', 'FL VAR', 'BP VAR', 'EX VAR', 'INFORM ALARMS', 'CUBES', 'USERFRAMES', 'TCP', 'COLLISION DETECT', 'LAYOUT', 'Notes(1)', 'Notes(2)', 'Notes(3)'])    
    excel_workbook_dict = pd.read_excel(var_tracker_fpath, sheet_name=None)
    print(f'infile dictionary keys: {excel_workbook_dict.keys()}')
    relevant_sheets = header_dict.keys()

    # outfile name
    DAT_outfile = args.outDir + args.outSubdir + "VARNAME.DAT"

    # extract contents from excel file
    DAT_dict, DAT_lengths = OrderedDict(), []
    for i, sheet_name in enumerate(relevant_sheets):

        LOGGER.info('parsing current sheet: %s', sheet_name)

        # <class 'pandas.core.frame.DataFrame'>
        df = excel_workbook_dict[sheet_name]

        # data always starts at index 8
        # <class 'pandas.core.frame.DataFrame'>
        selected_rows = df.iloc[8:,]
        print(f'confirm selected rows for sheet {sheet_name}')
        print(selected_rows.head())

        # select the hashtag column and the Name column
        temp_hashtag_col, name_col = selected_rows.iloc[:,0].map(str).to_list(), selected_rows.iloc[:,1].to_list()
        # pad zeros
        hashtag_col = [num.zfill(4) for num in temp_hashtag_col]

        # zip data for outfile
        curr_data = list(zip(hashtag_col, name_col))
        curr_data_len = len(curr_data)

        # append to data structures
        DAT_dict[sheet_name] = curr_data
        DAT_lengths.append(str(curr_data_len))

    LOGGER.info("confirm DAT data:")
    for section, length in zip(DAT_dict.keys(), DAT_lengths):
        print(f'section: {section} \t calculated length: {length}')

    # dynamically pad compiled data
    pad_data(DAT_dict)

    # write outfile
    write_to_DAT_file(DAT_dict, DAT_lengths, DAT_outfile)


if __name__ == '__main__':


    # set console logger for info and debugging
    LOGGER = set_logger()
    
    # register command line arguments
    register_arguments()

    # run main program
    run()

    print('\nProcess complete.')






