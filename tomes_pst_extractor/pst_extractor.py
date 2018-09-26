#!/usr/bin/env python3

""" This module contains a class for converting a PST file to MIME/EML. It is a wrapper around
./lib/pst_to_mime.exe, itself compiled from ./lib/pst_to_mime.cs.

Todo:
    - Complete .md files in docs.
    - Add unit tests.
    - Verify/validate sample data output.
    - Find a safe sample PST.
    - Test with PSTs that have an "@" in the filename.
"""

__NAME__ = "tomes_pst_extractor"
__FULLNAME__ = "TOMES PST Extractor"
__DESRIPTION__ = "Part of the TOMES project: converts a PST file to MIME/EML."
__URL__ = "https://github.com/statearchivesofnorthcarolina/tomes-pst-extractor"
__VERSION__ = "0.0.2"
__AUTHOR__ = ["Jeremy Gibson", "Nitin Arora"]
__AUTHOR_EMAIL__ = "nitin.a.arora@ncdcr.gov"

# import modules.
import sys; sys.path.append("..")
import logging
import logging.config
import os
import plac
import platform
import subprocess
import yaml


class PSTExtractor():
    """ A class for converting a PST file to MIME/EML.  It is a wrapper around 
    ./lib/pst_to_mime.exe, itself compiled from ./lib/pst_to_mime.cs.

    Example:
        >>> sample = "../tests/sample_files/sample.pst"
        >>> pst_extractor = PSTExtractor("sample_account", sample, "../tests/sample_files")
        >>> pst_extractor.extract()
    """
    

    def __init__(self, account_name, pst_file, output_path, charset="utf-8"): 
        """ Sets instance attributes.
    
        Args:
            - account_name (str): An identifier for the email account.
            - pst_file (str): The PST file to convert.
            - output_path (str): The parent folder in which to place the MIME/EML data.
            - charset (str): Optional encoding for the captured output from the 
            ./lib/pst_to_mime.exe PST converter.
        """
    
        # set logging.
        self.logger = logging.getLogger(__name__)        
        self.logger.addHandler(logging.NullHandler())
        self.event_logger = logging.getLogger("event_logger")
        self.event_logger.addHandler(logging.NullHandler())

        # convenience functions to clean up path notation.
        self._normalize_sep = lambda p: p.replace(os.sep, os.altsep) if (
                os.altsep == "/") else p
        self._normalize_path = lambda p: self._normalize_sep(os.path.normpath(p)) if (
                p != "") else ""
        self._join_paths = lambda *p: self._normalize_path(os.path.join(*p))

        # set attributes.
        self.account_name = account_name
        self.pst_file = self._normalize_path(pst_file)
        self.output_path = self._normalize_path(output_path)
        self.charset = charset

        # set EML folder path.
        self.mime_path = self._join_paths(self.output_path, self.account_name)

        # set path to .exe PST extractor.
        self.extractor_app = "pst_to_mime.exe"
        self.extractor_path = self._join_paths(os.path.dirname(__file__), "lib", 
                self.extractor_app)
        if not os.path.isfile(self.extractor_path):
            msg = "Can't find .exe file: {}".format(self.extractor_app)
            raise FileNotFoundError(msg)

        # set logger for output from @self.extractor.
        self.subprocess_logger = logging.getLogger(self.extractor_app)
        self.subprocess_logger.addHandler(logging.NullHandler())

        # create a list of valid logging levels.
        self.valid_log_levels = ["debug", "info", "warning", "error", "critical"]

        # validate instance arguments.
        self.validate()


    def validate(self):
        """ Validates instance arguments passed to the constructor.

        Returns:
            None

        Raises:
            - FileNotFoundError: If @self.pst_file is not a file.
            - NotADirectoryError: If @self.output_path is not a directory.
            - IsADirectoryError: If @self.mime_path already exists.
        """

        self.logger.info("Validating instance arguments.")

        # test if @self.account_name is an identifier.
        if not self.account_name.isidentifier():
            msg = "Account name '{}' is not a valid identifier; problems may occur.".format(
                    self.account_name)
            self.logger.warning(msg)

        # verify @self.pst_file exists.
        if not os.path.isfile(self.pst_file):
            msg = "Can't find PST file: {}".format(self.pst_file)
            raise FileNotFoundError(msg)

        # verify @self.output_path exists.
        if not os.path.isdir(self.output_path):
            msg = "Can't find folder: {}".format(self.output_path)
            raise NotADirectoryError(msg)

        # make sure @self.mime_path doesn't already exist.
        if os.path.isdir(self.mime_path):
            msg = "Can't overwrite existing MIME folder: {}".format(self.mime_path)
            raise IsADirectoryError(msg)

        return


    def _log_subprocess_line(self, line):
        """ Using @self.subprocess_logger, this logs a @line of text outputted by 
        @self.extractor_app with the appropriate logging level.
        
        Args:
            line (str): The text outputted by @self.extractor_app.
        
        Returns:
            None
        """

        # determine logging level based on the line prefix ("ERROR: ", "WARNING: ", etc.).
        try:
            level, message = line.split(":", 1)
            level, message = level.strip().lower(), message.strip()
        except:
            level, message = None, line.strip()

        # assume "info" level if @level is not a valid level.
        if level not in self.valid_log_levels:
            level, message = "info", line.strip()

        # log @message.
        try:
            getattr(self.subprocess_logger, level)(message)
        except Exception as err:
            self.logger.warning("Can't log subprocess message: {}".format(message))
            self.logger.error(err)

        return


    def _run_extractor(self):
        """ Runs @self.extractor_app as a subprocess. If the operating system is not Windows,
        Mono is used to run the .exe file.
        
        Returns:
            None
        """
        
        # create the command to run.
        cli_args = [self.extractor_path, self.account_name, self.pst_file, self.output_path]
        if platform.system() != "Windows":
            cli_args.insert(0, "mono")
        self.logger.debug("Running command: {}".format(" ".join(cli_args)))
    
        # prepare to capture each character outputted from @self.extractor_app.
        line_parts = []

        # run @self.extractor_app.
        # based on: https://stackoverflow.com/a/803396
        process = subprocess.Popen(cli_args, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                universal_newlines=True)
        while process.poll() is None:
             
            # save output to @line_parts as long as the output is not a line break.
            # if the output is a line break, @line_parts is converted to a string and logged
            # and @line_parts is cleared.
            for std_out in process.stdout.read(1):
                
                if std_out != "\n":
                    std_out = std_out.encode(self.charset).decode(self.charset, 
                            errors="replace")
                    line_parts.append(std_out)
                    process.stdout.flush()
                else:
                    line = "".join(line_parts)
                    line_parts[:] = []
                    self._log_subprocess_line(line)
        
        return
 

    def extract(self):
        """ Extracts @self.pst_file to @self.mime_path. 
        
        Returns:
            None
        """

        self.logger.info("Extracting '{}' to: {}".format(self.pst_file, self.mime_path))

        # extract @self.pst_file.
        try:
            self._run_extractor()
            self.logger.info("Created MIME at: {}".format(self.mime_path))
            self.event_logger.info({"entity": "agent", "name": __NAME__,  
                "fullname": __FULLNAME__, "uri": __URL__, "version": __VERSION__})
            self.event_logger.info({"entity": "event", "name": "pst_to_mime", 
                "agent": __NAME__, "object": "mime"})
            self.event_logger.info({"entity": "object", "name": "pst", 
                "category": "representation"})
            self.event_logger.info({"entity": "object", "name": "mime", 
                "category": "representation"})
        except Exception as err:
            self.logger.warning("Can't extract PST.")
            self.logger.error(err)

        return

    
# CLI.
def main(account_name: ("email account identifier"), 
        pst_file: ("PST file to convert"),
        output_path: ("parent folder for MIME/EML data"),
        silent: ("disable console logs", "flag", "s")):

    "Converts a PST file to MIME/EML.\
    \nexample: `python3 pst_extractor.py sample ../tests/sample_files/sample.pst ../tests/sample_files/`"

    # make sure logging directory exists.
    logdir = "log"
    if not os.path.isdir(logdir):
        os.mkdir(logdir)

    # get absolute path to logging config file.
    config_dir = os.path.dirname(os.path.abspath(__file__))
    config_file = os.path.join(config_dir, "logger.yaml")
    
    # load logging config file.
    with open(config_file) as cf:
        config = yaml.safe_load(cf.read())
    if silent:
        config["handlers"]["console"]["level"] = 100
    logging.config.dictConfig(config)
    
    # convert PST file.
    logging.info("Running CLI: " + " ".join(sys.argv))
    try:
        pst_extractor = PSTExtractor(account_name, pst_file, output_path)
        pst_extractor.extract()
        logging.info("Done.")
        sys.exit()
    except Exception as err:
        logging.critical(err)
        sys.exit(err.__repr__())


if __name__ == "__main__":
    plac.call(main)
