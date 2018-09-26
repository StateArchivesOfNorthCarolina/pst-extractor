# Introduction

**TOMES PST Extractor** is part of the [TOMES](https://www.ncdcr.gov/resources/records-management/tomes) project.

It is written in Python.

Its purpose is to create a MIME/EML version of a PST file.

# External Dependencies
TOMES PST Extractor requires the following:

- [Python](https://www.python.org) 3.0+ (using 3.5+)
	- See the `./requirements.txt` file for additional module dependencies.
	- You will also want to install [pip](https://pypi.python.org/pypi/pip) for Python 3.
- [MailBee.NET Objects](https://afterlogic.com/mailbee-net/email-components) 11.2.0.590+ (using 11.2.0.590 for .NET 4.5)
	- The required file, `MailBee.NET.dll`, is included in the `./tomes_pst_extractor/lib/` folder.
- [Mono](https://www.mono-project.com) 5.14.0+ (using 5.14.0.177)
	- Mono is only required for Linux/macOS users.

# Installation
After installing the external dependencies above, you'll need to install some required Python packages.

The required packages are listed in the `./requirements.txt` file and can easily be installed via PIP <sup>[1]</sup>: `pip3 install -r requirements.txt`

You should now be able to use TOMES PST Extractor from the command line or as a locally importable Python module.

If you want to install TOMES PST Extractor as a Python package, do: `pip3 install . -r requirements.txt`

Running `pip3 uninstall tomes_pst_extractor` will uninstall the TOMES PST Extractor package.

# Making the .exe file
If you need to edit the source C# code for `./tomes_pst_extractor/lib/pst_to_mime.cs` and recompile it as `./tomes_pst_extractor/lib/pst_to_mime.exe`, you will need a valid [MailBee License Key](https://afterlogic.com/mailbee-net/docs/keys.html).

The current version of the `./tomes_pst_extractor/lib/pst_to_mime.exe` was compiled on Linux with the **Mono C# Compiler** using the following steps:

1. `cd tomes_pst_extractor/lib`
2. `mkdir bin`
	1. This makes a temporary directory in which to compile a **copy** the source code. Using a copy helps prevent leaving your MailBee License Key in  the source code.
3. `cp pst_to_mime.cs bin/pst_to_mime.cs`
4. `vi bin/pst_to_mime.cs`
	1. After opening the file, insert a valid license key (see the source code for more information) and save the file.
5. `mcs /reference:MailBee.NET.dll bin/pst_to_mime.cs`
6. `mv bin/pst_to_mime.exe .`
7. `rm -r bin`
	1. This is important because it deletes the temporary directory that contains the copy of the source code with your license key.

# Unit Tests
While not true unit tests that test each function or method of a given module or class, basic unit tests help with testing overall module workflows.

Unit tests reside in the `./tests` directory and start with "test__".

## Running the tests
To run all the unit tests do <sup>[1]</sup>: `python3 -m unittest` from within the `./tests` directory. 

## Using the command line
All of the unit tests have command line options.

To see the options and usage examples simply call the scripts with the `-h` option: `python3 test__[rest of filename].py -h` and try the example.

To run all the unit tests do <sup>[1]</sup>: `python3 -m unittest` from within 
Sample files are located in the `./tests/sample_files` directory.

The sample files can be used with the command line options of some of the unit tests.

# Modules
TOMES PST Extractor consists of single-purpose high, level module, `pst_extractor.py`. It can be used as native Python class or as command line script.

## Using pst_extractor.py with Python
To get started, import the module and run help():

	>>> from tomes_pst_extractor import pst_extractor
	>>> help(pst_extractor)

*Note: docstring and command line examples may reference sample and data files that are NOT included in the installed Python package. Please use appropriate paths to sample and data files as needed.*

## Using pst_extractor.py from the command line
1. From the `./tomes_pst_extractor` directory do: `python3 pst_extractor.py -h` to see an example command.
2. Run the example command.

-----
*[1] Depending on your system configuration, you might need to specify "py -3", etc. instead of "python3" from the command line. Similar differences might apply for PIP.*
