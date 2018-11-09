#!/usr/bin/env python3

# import modules.
import sys; sys.path.append("..")
import glob
import logging
import os
import plac
import tempfile
import unittest
from tomes_pst_extractor import pst_extractor

# enable logging.
logging.basicConfig(level=logging.DEBUG)


def EXTRACT_AND_COUNT(pst_file):
    """ Converts @pst_file to MIME inside a temporary directory. After the number of EML files
    is tallied the directory is deleted.
    
    Returns:
        int: The return value.
        The number of EML files.

    Raises:
        - FileNotFoundError: If @pst_file isn't a file.
    """

    # make sure @pst_file exists.
    if not os.path.isfile(pst_file):
        msg = "Can't find file: {}".format(pst_file)
        raise FileNotFoundError(msg)

    # create temporary account name.
    account_name = "sample"  
    
    # create temporary folder; extract PST and count EML file.
    with tempfile.TemporaryDirectory(dir=os.path.dirname(__file__)) as temp_dir:

        logging.info("Created temporary folder: {}".format(temp_dir))

        # extract PST.
        logging.info("Extracting PST file '{}' to: {}/{}".format(pst_file, temp_dir, 
            account_name))
        pst2mime = pst_extractor.PSTExtractor(account_name, pst_file, temp_dir)
        pst2mime.extract()

        # count EML files.
        eml_files = glob.glob(temp_dir + "/**/*.eml", recursive=True)
        eml_count = len(eml_files)
        logging.info("EML count: {}".format(eml_count))
    
    return eml_count


class Test_PSTConnverter(unittest.TestCase):


    def setUp(self):
        
        # set attributes.
        self.sample_file = "sample_files/sample.pst"


    def test__count_emls(self):
        """ Is the number of EML files extracted from @self.sample_file correct? """ 
        
        eml_count = EXTRACT_AND_COUNT(self.sample_file)
        self.assertTrue(eml_count == 3)


# CLI.
def main(pst_file: ("the PST file from which to tally EML files")):
    
    """Prints the number of EML files extracted from as PST file.\
    \nexample: `python3 test__pst_converter.py sample_files/sample.pst`\
    \n\nWARNING: Only use this with *small* PST data (< 100 messages)."    
    """

    try:
        eml_count = EXTRACT_AND_COUNT(pst_file)
        print("EML count: {}".format(eml_count))
    except Exception as err:
        logging.error(err)


if __name__ == "__main__":    
    plac.call(main)
