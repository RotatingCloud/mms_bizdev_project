import argparse as ap
import re
import pymongo
import os
import subprocess
import xlwt
import xlrd
import concurrent.futures
from PIL import Image
from xlutils.copy import copy
import shutil
import requests
from tqdm import tqdm
from dotenv import load_dotenv
load_dotenv()
import logging
logging.basicConfig(filename='error_log.log', level=logging.DEBUG, format='%(asctime)s %(levelname)s:%(message)s')

class style():

    RESET = '\033[0m'
    BLACK = '\033[30m'
    RED = '\033[31m'
    GREEN = '\033[32m'
    YELLOW = '\033[33m'
    BLUE = '\033[34m'
    MAGENTA = '\033[35m'
    CYAN = '\033[36m'
    WHITE = '\033[37m'
    BRIGHT_BLACK = '\033[90m'
    BRIGHT_RED = '\033[91m'
    BRIGHT_GREEN = '\033[92m'
    BRIGHT_YELLOW = '\033[93m'
    BRIGHT_BLUE = '\033[94m'
    BRIGHT_MAGENTA = '\033[95m'
    BRIGHT_CYAN = '\033[96m'
    BRIGHT_WHITE = '\033[97m'

class Project3:

    def __init__(self, video_to_process, output, verbose):

        self.video_to_process = video_to_process
        self.output = output
        self.verbose = verbose

    def to_int(self, s):
        
        if isinstance(s, int):
            return s
        return int(s.split('-')[0])

    def frame_to_timecode(self, frame, fps):

        hours = frame // (3600 * fps)
        minutes = (frame % (3600 * fps)) // (60 * fps)  
        seconds = (frame % (60 * fps)) // fps
        frames = frame % fps
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}:{frames:02d}"

    def process_frames(self, x):

        if type(x) == str:
            first, last = map(int, x.split('-'))
            middle = first + (last - first) // 2
            return int(first), int(last), int(middle)
        elif type(x) == float:
            return int(x), int(x), int(x)
        return int(x), int(x), int(x)

    def export_to_xls(self, array):

        simplified_array = []
        thumbnail_timecodes = []
        for x in array:
            first, last, middle = self.process_frames(x['frame/range'])
            timecode = f"{self.frame_to_timecode(first, 60)}-{self.frame_to_timecode(last, 60)}"
            entry = [x['path'], x['frame/range'], timecode, None]
            simplified_array.append(entry)
            thumbnail_timecodes.append(middle)
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("Sheet1")
        for i, x in enumerate(simplified_array):
            for j, y in enumerate(x):
                sheet.write(i, j, y)
        workbook.save("output.xls")
        return thumbnail_timecodes

    def generate_thumbnail(self, frame):

        frame_rate = 60  
        hours = int(frame / (frame_rate * 3600))
        minutes = int((frame % (frame_rate * 3600)) / (frame_rate * 60))
        seconds = int((frame % (frame_rate * 60)) / frame_rate)
        milliseconds = int((frame % frame_rate) * (1000 / frame_rate))
        timestamp = f"{hours:02d}:{minutes:02d}:{seconds:02d}.{milliseconds:03d}"
        if not os.path.exists(f'thumbnails/{frame}.bmp'):
            try:
                command = [
                    'ffmpeg', '-y', '-ss', timestamp, '-i', f'{self.video_to_process}',
                    '-vf', 'scale=96:74',  
                    '-frames:v', '1', f'thumbnails/{frame}.bmp',
                    '-hide_banner'
                ]
                subprocess.call(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                #self.verbose and print(style.BRIGHT_GREEN + f"Thumbnail for frame {frame} generated successfully." + style.RESET)
            except Exception as e:
                print(f"An error occurred: {e}")
        else:
            self.verbose and print(f"Thumbnail for frame {frame} already exists.")

    def process_frames_in_parallel(self, frames):

        with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
            future_to_frame = {executor.submit(self.generate_thumbnail, frame): frame for frame in frames}
            for future in tqdm(concurrent.futures.as_completed(future_to_frame), total=len(frames), desc=(style.BRIGHT_GREEN + "Generating thumbnails and inserting into .xls" + style.RESET), ncols=100):
                frame = future_to_frame[future]
                try:
                    future.result()
                except Exception as e:
                    self.verbose and print(f'Frame {frame} generated an exception: {e}')

    def insert_images(self):

        in_book = xlrd.open_workbook('output.xls', formatting_info=True)
        out_book = copy(in_book)  
        sheet = out_book.get_sheet(0)  
        in_sheet = in_book.sheet_by_index(0)
        for row_idx in range(in_sheet.nrows):
            row = [in_sheet.cell(row_idx, col_idx).value for col_idx in range(in_sheet.ncols)]
            _, _, thumbnail_frame = self.process_frames(row[1])
            thumbnail_path = f"thumbnails/{thumbnail_frame}.bmp"
            if os.path.exists(thumbnail_path):
                img = Image.open(thumbnail_path)
                img.thumbnail((96, 74))
                img.save(thumbnail_path)
                sheet.insert_bitmap(thumbnail_path, row_idx, 3)

        out_book.save('output.xls')
        self.verbose and print(style.GREEN + "\t- All thumbnails inserted successfully." + style.RESET)

    def upload_image(self, thumbnail, url, token):

        path = f"thumbnails/{thumbnail}"
        try:
            payload = {
                'filesize': os.path.getsize(path),
                'filetype': 'image/bmp',
                'name': thumbnail,
                'type': 'file',
            }
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {token}"
            }
            response = requests.request("POST", url, json=payload, headers=headers)
            data = response.json()

            try:
                upload_url = data['upload_urls'][0]
            except KeyError:
                # Log detailed error information
                logging.error(f"KeyError: 'upload_urls' not found in response. Response data: {data}")
                return  # Or handle the error as needed

            with open(path, 'rb') as f:
                headers = {
                    'content-type': 'image/bmp',
                    'x-amz-acl': 'private',
                }
                response = requests.put(upload_url, data=f, headers=headers)
            if response.status_code == 200:
                #self.verbose and print(f"{thumbnail} uploaded successfully")
                pass
            else:
                print(f"An error occurred while uploading {thumbnail}: {response.text}")
        
        except Exception as e:

            print(f"An error occurred while uploading {thumbnail}: {e}")
            logging.error(f"An error occurred: {e}")
            logging.debug(f"Error details: {e}, Response: {response.text}")

    def process_upload_in_parallel(self):

        asset_id = os.getenv('FRAMEIO_ASSET_ID')
        url = "https://api.frame.io/v2/assets/" + asset_id + "/children"
        thumbnails = os.listdir('thumbnails')
        token = os.getenv('FRAMEIO_TOKEN')
        with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
            futures = [executor.submit(self.upload_image, thumbnail, url, token) for thumbnail in thumbnails]
            for future in tqdm(concurrent.futures.as_completed(futures), total=len(thumbnails), desc=(style.BRIGHT_CYAN + "Sending to Frame.io" + style.RESET), ncols=100):
                future.result()  

    def delete_thumbnails(self):
            
        shutil.rmtree('thumbnails')

    def process(self):

        myclient = pymongo.MongoClient("mongodb://localhost:27017/")
        mydb = myclient["project2"]
        col = mydb["workorder_info"]
        frames = []
        command = f'ffprobe -v error -select_streams v:0 -count_packets -show_entries stream=nb_read_packets -of csv=p=0 {self.video_to_process}'
        output = subprocess.check_output(command, shell=True)
        num_of_frames = int(re.sub('[^0-9]', '', str(output)))
        for x in col.find():
            frame = self.to_int(x['frame/range'])
            if frame < num_of_frames:
                frames.append(x)
        sorted_frames = sorted(frames, key=lambda x: self.to_int(x['frame/range']))
        if self.output == 'none':
            self.verbose and print("\nNot exporting to anything")
        elif self.output != 'xls':
            self.verbose and print("\nInvalid output type")
        else:
            self.verbose and print(style.YELLOW + "Exporting to .xls file" + style.RESET)
            frames = self.export_to_xls(sorted_frames)
            if not os.path.exists('thumbnails'):
                os.makedirs('thumbnails')
            else:
                self.delete_thumbnails()
                os.makedirs('thumbnails')
            self.process_frames_in_parallel(frames)
            self.insert_images()
            self.process_upload_in_parallel()
            self.delete_thumbnails()
            self.verbose and print('\033[47m' + style.BRIGHT_GREEN + " Done! " + style.RESET)

parser = ap.ArgumentParser()
parser.add_argument('-v', '--verbose', help="Verbose mode", action="store_true")
parser.add_argument('-o','--output', help="Output as csv file or database")
parser.add_argument('-p','--process', dest='process', nargs='?', help="project 3")
args = parser.parse_args()

output = args.output
video_to_process = args.process
verbose = args.verbose

p = Project3(video_to_process, output, verbose)
p.process()