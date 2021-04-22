'''
packages
'''
import gspread
from oauth2client.service_account import ServiceAccountCredentials

from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import io
from apiclient.http import MediaIoBaseDownload

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

from os import listdir
from os.path import isfile, join
import PIL.ImageOps

from PIL import Image, ImageChops
from pytesseract import *

from pandas import DataFrame
from datetime import date

from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

import io
from apiclient.http import MediaIoBaseDownload

'''
구글 드라이브에서 이미지 가져오기 
'''
def get_credentials_and_file_list(PHOTO_TYPE, FILE_ID):
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secrets.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    service = build('drive', 'v3', credentials=creds)
    # Call the Drive v3 API
    EXPORTED_FILE_NAME = 'result\\' + PHOTO_TYPE + '\\ex_' + FILE_ID + '.png'
    print(EXPORTED_FILE_NAME)
    
    #request = service.files().export_media(fileId=FILE_ID, mimeType=PNG)
    request = service.files().get_media(fileId = FILE_ID)
    fh = io.FileIO(EXPORTED_FILE_NAME, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print('Download %d%%.' % int(status.progress() * 100))

'''
이미지 trim/resize/crop
'''
def trim(image):
    bg = Image.new(image.mode, image.size, image.getpixel((0,0)))
    diff = ImageChops.difference(image, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        return image.crop(bbox)

def resize(image, basewidth):
    basewidth = basewidth
    wpercent = (basewidth/float(image.size[0]))
    hsize = int((float(image.size[1])*float(wpercent)))
    img = image.resize((basewidth,hsize), Image.ANTIALIAS)
    return img

def crop_to_text(save_list, image, area):    
    cropped_img = image.crop(area)
    #inverted_img = PIL.ImageOps.invert(cropped_img)
    txt = pytesseract.image_to_string(cropped_img, lang="kor+eng")
    #save_list.append(txt.split('|')[0].replace('\n\x0c','').replace('\n','').replace(',','').replace('(','').replace('.','').replace(')',''))
    save_list.append(txt.replace('\n\x0c','').replace('\n','').replace(',','').replace('(','').replace('.','').replace(')',''))
    return save_list

'''
setting 정보 
'''

CREDENTIAL_DIR = 'C:\\google-drive\\'
CREDENTIAL_FILENAME = 'kvk-contribution-8432479ff624.json'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'ant-kvk'
# If modifying these scopes, delete the file token.json.
SCOPES = [
'https://spreadsheets.google.com/feeds',
'https://www.googleapis.com/auth/drive',
]
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1a-WvWsMAr40qnU-m3DjTcBZKPyrOcoxeFSyLTVFpOFk/edit#gid=1644684386'

#### tesseract location
pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

'''
설문 시트 가져오기 
설문 시트의 이미지 id를 활용하기 위함 
'''

json_file_name = 'kvk-contribution-16b840c3f466.json'
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, SCOPES)
gc = gspread.authorize(credentials)

doc = gc.open_by_url(spreadsheet_url)
worksheet = doc.worksheet('설문지 응답 시트1' ) # 시트 선택하기

'''
프로필 정보 
'''
checked = worksheet.col_values(4)

PHOTO_TYPE = 'profile'
profile = worksheet.col_values(2) # Profile 열 선택하기
LEN = len(profile) - 1

uid_list = []
id_list = []
alliance_list = []
power_list = []
profile_image_id_list =  [0] * (LEN-len(checked))

for i in range(len(checked), LEN):
    print(i)
    FILE_ID = profile[i+1].split('=')[1]
    print(FILE_ID)
    get_credentials_and_file_list(PHOTO_TYPE, FILE_ID)
    image = Image.open('result\\' + PHOTO_TYPE + '\\ex_' + FILE_ID + '.png').convert("L")
    image = trim(image)
    image = resize(image, 1000)
    #uid_area = (454, 132, 545, 153) # ID
    uid_area = (454, 132, 590, 153) # ID
    uid_list = crop_to_text(uid_list, image, uid_area)
    id_area = (370, 164, 653, 191) # UID
    id_list = crop_to_text(id_list, image, id_area)
    alliance_area = (372, 242, 575, 265) # Alliance
    alliance_list = crop_to_text(alliance_list, image, alliance_area)
    power_area = (604, 242, 728, 265) # power
    power_list = crop_to_text(power_list, image, power_area)
    profile_image_id_list[i-len(checked)] = FILE_ID

'''
전사자수/킬수 정보 
'''
PHOTO_TYPE = 'info'
info = worksheet.col_values(3) # info 열 선택하기

kill_4t_front_list = []
kill_5t_front_list = []
dead_list = []
info_id_list =  [0] * (LEN-len(checked))

for i in range(len(checked), LEN):
    print(i)
    FILE_ID = info[i+1].split('=')[1]
    print(FILE_ID)
    get_credentials_and_file_list(PHOTO_TYPE, FILE_ID)
    image = Image.open('result\\' + PHOTO_TYPE + '\\ex_' + FILE_ID + '.png').convert("L")
    image = trim(image)
    image = resize(image, 1000)
    kill_4t_area = (837, 397, 950, 418) # kill 4t
    kill_4t_front_list = crop_to_text(kill_4t_front_list, image, kill_4t_area)
    kill_5t_area = (867, 430, 950, 453) # kill 5t
    kill_5t_front_list = crop_to_text(kill_5t_front_list, image, kill_5t_area)
    dead_area = (782, 559, 899, 586) # 전사자 수 
    dead_list = crop_to_text(dead_list, image, dead_area)
    info_id_list[i-len(checked)] = FILE_ID

'''
data frame화 
'''
contribution_data = DataFrame(
    { 'date' : date.today().strftime("%b-%d-%Y"),
      'uid' : uid_list,
      'id' : id_list,
      'power' : power_list,
      'alliance' : alliance_list,
      'kill_4t' : kill_4t_front_list,
      'kill_5t' : kill_5t_front_list,
      'death' : dead_list,
      'image_profile_id' : profile_image_id_list,
      'image_info_id' : info_id_list
    })

'''
새 스프레드 시트에 반영하기 
'''
import gspread_dataframe as gd

ws = gc.open("2차 kvk 기여도 설문").worksheet("4.21 데이터 추출")
existing = gd.get_as_dataframe(ws)
updated = existing.append(contribution_data)
gd.set_with_dataframe(ws, updated)
