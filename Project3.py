import argparse as ap
import re
import pymongo
import csv
import sys
import os
from datetime import datetime
import subprocess
import xlsxwriter
import shutil
import requests

from dotenv import load_dotenv
load_dotenv()

class Entry:

    def __init__(self, name, path, frames):

        self.name = name
        self.path = path
        self.frames = frames

class Project2:

    def __init__(self, files_to_process, xytech_file, output, verbose, Keyword, User):

        self.files_to_process = files_to_process
        self.xytech_file = xytech_file
        self.output = output
        self.Keyword = Keyword
        self.User = User
        self.verbose = verbose

        x = re.sub(r'\..*$', '', xytech_file).split('_')
        date = x[1]
        date_obj = datetime.strptime(date, '%Y%m%d')
        self.Date = date_obj

    def chunk(self, array):

        chunked = []
        for i in range(len(array)):
            s = array[i]
            if not s or not s.isdigit():
                continue
            num = int(s)
            if i == 0:
                chunked.append([num])
            elif array[i-1].isdigit() and num == int(array[i-1]) + 1:
                chunked[-1].append(num)
            else:
                chunked.append([num])
        return chunked

    def to_int(self, s):

        if isinstance(s, int):
            return s
        return int(s.split('-')[0])

    def parse_line(self, line):

        parts = re.split(r'\s+', line)
        path_components = []
        numbers = []
        for part in parts:
            if part.isdigit():
                numbers.append(part)
            else:
                if part != '<err>' and part != '<null>':
                    path_components.append(part)
        path = " ".join(path_components)
        return path, numbers

    def clip_path(self, path): 

        _, _, result = path.partition(self.Keyword)
        return self.Keyword + result

    def validate_file(self, file):

        if os.path.splitext(file)[1].lower() != '.txt':
            print("This file is not a .txt file :(")
            sys.exit(2) 
        x = re.sub(r'\..*$', '', file).split('_')
        machine = x[0]
        name = x[1]
        date = x[2]
        date_obj = datetime.strptime(date, '%Y%m%d')
        formatted_date = date_obj.strftime('%B %d, %Y')
        if machine.lower() != "baselight" and machine.lower() != "flame":
            print("Theres something wrong with one of the files provided :(")
            sys.exit(2)
        else:
            return machine, name, formatted_date

    def validate_xytech(self, file):

        xytech = re.sub(r'\..*$', '', file).split('_')
        test_date = xytech[1]
        date_obj = datetime.strptime(test_date, '%Y%m%d')
        formatted_date = date_obj.strftime('%B %d, %Y')
        producer = ""
        operator = ""
        job = ""
        notes = ""
        locations = []
        with open(f"import_files/{file}", "r") as f:
            lines = f.readlines()
        for i, line in enumerate(lines):
            if line.startswith("Producer"):
                producer = line.split(":")[1].strip()
            elif line.startswith("Operator"):
                operator = line.split(":")[1].strip()
            elif line.startswith("Job"):
                job = line.split(":")[1].strip()
            elif line.startswith("Location"):
                for j in range(i + 1, len(lines)):
                    if lines[j] != "\n":
                        locations.append(lines[j].strip())
                    else:
                        break
            elif line.startswith("Notes"):

                notes = lines[i + 1].strip()
        return formatted_date, producer, operator, job, notes, locations
    
    def export_to_database(self, results):

        try:
            myclient = pymongo.MongoClient("mongodb://localhost:27017/")
            mydb = myclient["project2"]
            col1 = mydb["workorder_log"]
            col2 = mydb["workorder_info"]
        except Exception as e:
            print(f"Database connection failed: {e}")
            sys.exit(2)  
        user_that_ran_script = self.User
        for file in self.files_to_process:
            machine, name, _ = self.validate_file(file)
            if not machine or not name or not self.Date:
                print(f"Validation failed for file {file}")
                continue  
            x1 = {"user": user_that_ran_script, "machine": machine, "name": name, "date": self.Date.strftime("%Y%m%d"), "submitted": datetime.now().strftime("%Y%m%d")}
            col1.insert_one(x1)
        for r in results:
            x2 = {"name": r[0], "path": r[1], "date": self.Date.strftime("%Y%m%d"), "frame/range": r[2]}
            col2.insert_one(x2)

    def export_to_csv(self, results, Producer, Operator, Job, Notes):

        with open(f'{self.User}_{self.Date.strftime("%Y%m%d")}.csv', 'w', newline='') as file:
            writer = csv.writer(file)
            values = f"Producer: {Producer} / Operator: {Operator} / Job: {Job} / Notes: {Notes}"
            producer, operator, job, notes = values.split("/")
            writer.writerow([producer, operator, job, notes])
            for _ in range(2): writer.writerow([])
            for r in results:
                writer.writerow([r[1], r[2]])

    def process(self):

        entries = []
        self.verbose and print(f"\n\033[1m========================================\n\033[0m")
        for i, x in enumerate(self.files_to_process):
            self.verbose and print(f"\033[1mProcessing file {i+1} of {len(self.files_to_process)}:\033[1m \n    {x}\n")
            m, n, d = self.validate_file(x)
            with open(f"import_files/{x}", "r") as f:
                lines = f.readlines()
                for line in lines:
                    path, frames = self.parse_line(line)
                    entry = Entry(n, self.clip_path(path).strip(), self.chunk(frames))
                    entries.append(entry)
        c1 = 0
        for e in entries:
            c2 = 0
            for y in e.frames:
                c2 += len(y)
            c1 += c2
        self.verbose and print(f"\033[1mTotal frames before processing:\033[0m {c1}")
        date, producer, operator, job, notes, locations = self.validate_xytech(self.xytech_file)
        self.verbose and print(f"\033[1m\nDate: {date} \nProducer: {producer} \nOperator: {operator} \nJob: {job} \nNotes: {notes}\033[0m")
        self.verbose and print(f"\033[1m\nLocations:\033[0m")
        if self.verbose:
            for location in locations:
                print(f"    {location}")
        for e in entries:
            for location in locations:
                if e.path.strip() in location.strip():
                    e.path = location
                else:
                    continue
        final_result = []
        for e in entries:
            for x in e.frames:
                first= x[0]
                last = x[-1]
                if first == last:
                    final_result.append([e.name, e.path, first])
                else:
                    final_result.append([e.name, e.path, f"{first}-{last}"])
        final_result = sorted(final_result, key=lambda x: self.to_int(x[2]))
        if self.output == 'db':
            self.verbose and print("\nExporting to database")
            self.export_to_database(final_result)
        elif self.output == 'csv':
            self.verbose and print("\nExporting to CSV file")
            self.export_to_csv(final_result, producer, operator, job, notes)
        elif self.output == 'none':
            self.verbose and print("\nNot exporting to anything")
        return final_result, producer, operator, job, notes

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

    def export_to_xlsx(self, array):
        
        workbook = xlsxwriter.Workbook('output.xlsx')
        worksheet = workbook.add_worksheet()

        for i, x in enumerate(array):
            first, last, middle = self.process_frames(x['frame/range'])
            timecode = f"{self.frame_to_timecode(first, 60)}-{self.frame_to_timecode(last, 60)}"
            image = f"thumbnails/{middle}.png"
            worksheet.write(f'A{i+1}', x['path'])
            worksheet.write(f'B{i+1}', x['frame/range'])
            worksheet.write(f'C{i+1}', timecode)
            worksheet.insert_image(f'D{i+1}', image)

        workbook.close()

    def generate_thumbnail(self, frame):

        frame_rate = 60  
        hours = int(frame / (frame_rate * 3600))
        minutes = int((frame % (frame_rate * 3600)) / (frame_rate * 60))
        seconds = int((frame % (frame_rate * 60)) / frame_rate)
        milliseconds = int((frame % frame_rate) * (1000 / frame_rate))
        timestamp = f"{hours:02d}:{minutes:02d}:{seconds:02d}.{milliseconds:03d}"
        if not os.path.exists(f'thumbnails/{frame}.png'):
            try:
                command = [
                    'ffmpeg', '-y', '-ss', timestamp, '-i', f'{self.video_to_process}',
                    '-vf', 'scale=96:74',  
                    '-frames:v', '1', f'thumbnails/{frame}.png',
                    '-hide_banner'
                ]
                subprocess.call(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                #self.verbose and print(style.BRIGHT_GREEN + f"Thumbnail for frame {frame} generated successfully." + style.RESET)
            except Exception as e:
                print(f"An error occurred: {e}")
        else:
            self.verbose and print(f"Thumbnail for frame {frame} already exists.")

        self.verbose and print(f"DEBUG: {frame} generated")

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
                return

            with open(path, 'rb') as f:

                headers = {
                    'content-type': 'image/bmp',
                    'x-amz-acl': 'private',
                }

                response = requests.put(upload_url, data=f, headers=headers)

            if response.status_code == 200:
                
                self.verbose and print(f"Thumbnail {thumbnail} uploaded successfully.")
                pass

            else:
                
                print(f"An error occurred while uploading {thumbnail}: {response.text}")
        
        except Exception as e:

            print(f"An error occurred while uploading {thumbnail}: {e}")

    def upload_images(self):

        asset_id = os.getenv('FRAMEIO_ASSET_ID')
        url = "https://api.frame.io/v2/assets/" + asset_id + "/children"
        thumbnails = os.listdir('thumbnails')
        token = os.getenv('FRAMEIO_TOKEN')
        for thumbnail in thumbnails:
            self.upload_image(thumbnail, url, token)

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
            if not os.path.exists('thumbnails'):
                os.makedirs('thumbnails')
            else:
                self.delete_thumbnails()
                os.makedirs('thumbnails')

            for frame in sorted_frames:
                _, _, middle = self.process_frames(frame['frame/range'])
                self.generate_thumbnail(self.to_int(middle))

            self.export_to_xlsx(sorted_frames)
            self.upload_images()
            self.delete_thumbnails()
            self.verbose and print(" Done! ")

parser = ap.ArgumentParser()
parser.add_argument('-f','--files', dest='files', nargs='*', help="Files to be processed")
parser.add_argument('-x','--xytech', dest='xytech_file' ,help="Output file name", nargs='?')
parser.add_argument('-v', '--verbose', help="Verbose mode", action="store_true")
parser.add_argument('-o','--output', help="Output as csv file or database")
parser.add_argument('-p','--process', dest='process', nargs='?', help="project 3")
args = parser.parse_args()

xytech_file = args.xytech_file
verbose = args.verbose
output = args.output
video_to_process = args.process

if args.files == None:

    files_to_process = []

else:

    files_to_process = args.files

project2_mode = (len(files_to_process) != 0) and xytech_file != None

if(project2_mode == True):

    p = Project2(files_to_process, xytech_file, output, verbose, "Avatar", "JKwon")
    processed, Producer, Operator, Job, Notes = p.process()

    if verbose:

        print(f"\n\033[1m--------------------\033[0m")
        print(f"Producer:\033[0m {Producer}")
        print(f"\033[1mOperator:\033[0m {Operator}")
        print(f"\033[1mJob:\033[0m {Job}")
        print(f"\033[1mNotes:\033[0m {Notes}")
        print(f"\033[1mLocations:\033[0m")

        c1 = 0
        c2 = 0

        for i, x in enumerate(processed):

            number = x[2]

            if isinstance(number, int):

                c1 += 1
                print(f"\t{i+1, x, 1, c1}")

            elif isinstance(number, str) and '-' in number:

                first, last = map(int, number.split('-'))
                c1 += last - first + 1
                print(f"\t{i+1, x, last - first + 1, c1}")

            c2 += 1

        print(f"\n\033[1mTotal frames after processing:\033[0m {c1}")
        print(f"\033[1m    Total entries:\033[0m {c2}")
        print(f"\033[1m========================================\033[0m\n")

elif project2_mode == False and video_to_process == None:

    print("Invalid arguments. Please try again.")

elif project2_mode == False and video_to_process != None:

    if len(files_to_process) != 0 or xytech_file != None:

        print("Warning: Ignoring files and xytech file")

    p = Project3(video_to_process, output, verbose)
    p.process()