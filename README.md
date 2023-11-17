# mms_bizdev_project

REQUIREMENTS (make sure you have these installed before running):
- ffmpeg
- python 3.11.5

to run script:
- create venv with `python -m venv venv`
- activate venv `./venv/Scripts/Activate`
- install req `pip install -r requirements.txt`
- ask me for .env (only for the homies)
- run program

# Command-Line Tool Documentation

This documentation provides details on the available command-line arguments for the tool.

## Arguments

- `-f`, `--files`
  - **Description**: Files to be processed
  - **Destination**: `files`
  - **Type**: Multiple arguments allowed (`nargs='*'`)
  
- `-x`, `--xytech`
  - **Description**: Xytech file to be processed
  - **Destination**: `xytech_file`
  - **Type**: Optional argument (`nargs='?'`)

- `-v`, `--verbose`
  - **Description**: Enable verbose mode
  - **Action**: `store_true` (flag to activate verbose mode)

- `-o`, `--output`
  - **Description**: Format of the output (either as csv file or database (or xls only if --process is flagged))

- `-p`, `--process`
  - **Description**: Specify the video to process (default: project 3)
  - **Destination**: `process`
  - **Type**: Optional argument (`nargs='?'`)

## Usage

```shell
python project.py -f file1 file2 -x xytech_YYYYMMDD.txt -v -o csv
