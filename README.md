# show_me_the_ole
show_me_the_ole
Introduction

show_me_the_ole is a simple Python script designed to analyze and display the streams and storage objects within an OLE (Object Linking and Embedding) file. This tool is particularly useful for examining the structure of older Microsoft Office documents and other files using the OLE format. It can help you identify and inspect macros for malware analysis.
Features

    Validates if the provided file is a valid OLE file.
    Lists streams and storage objects within the OLE file.
    Displays basic information about the OLE file's header and Sector Allocation Table (SAT).
    Can be used to see macros for malware analysis.

Prerequisites

    Python 3.x

Installation

There are no external dependencies required for this script. Simply download the script to your local machine.
Usage

    Download the Script:

    Save the script as show_me_the_ole.py.

    Run the Script:

    Open a terminal or command prompt and navigate to the directory where show_me_the_ole.py is saved.

Script Overview
Functions

    print_usage(): Prints usage instructions.
    validate_file(file_path): Validates that the input file exists and is a valid OLE file.
    read_sector(f, sector, sector_size): Reads a specific sector from the file.
    list_streams(file_path): Lists streams and storage objects in the OLE file by parsing the directory entries.
    main(): Orchestrates the script, validating input and listing streams.
