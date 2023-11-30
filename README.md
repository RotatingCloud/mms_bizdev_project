# mms_bizdev_project

REQUIREMENTS (make sure you have these installed before running):
- ffmpeg
- python 3.11.5

to run script (windows):
- create venv with `python -m venv venv`
- activate venv `./venv/Scripts/Activate`
- install req `pip install -r requirements.txt`
- ask me for .env (only for the homies)
- run program

project_full.py contains the code for the previous two phases along with the new phase three__
project_lite.py contains only the code for phase three__

# Command-Line Tool Documentation

This documentation provides details on the available command-line arguments for the tool.

## Arguments 
##(ONLY USEABLE WITH project_full.py)

- `-f`, `--files`
  - **Description**: Files to be processed
  - **Type**: Multiple arguments allowed (`nargs='*'`)
  
- `-x`, `--xytech`
  - **Description**: Xytech file to be processed
  - **Type**: Optional argument (`nargs='?'`)

- `-v`, `--verbose`
  - **Description**: Enable verbose mode
  - **Action**: `store_true` (flag to activate verbose mode)

- `-o`, `--output`
  - **Description**: Format of the output (either as csv file or database (or xls only if --process is flagged))
    - none if no output needed

##(USEABLE WITH BOTH project_full.py and project_lite.py)

- `-p`, `--process`
  - **Description**: Specify the video to process
  - **Type**: Optional argument (`nargs='?'`)

## Example Usage

```shell
python project.py -f file1.txt file2.txt -x xytech_YYYYMMDD.txt -v -o csv

or

python project.py -p video.mp4 -v -o xls
