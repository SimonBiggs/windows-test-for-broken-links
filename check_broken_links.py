# Copyright 2016 Simon Biggs

# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at

#    http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.


import os
import sys
import traceback
import requests

from glob import glob
import win32com.client 

import numpy as np
import pandas as pd

import yaml
import json


broken_link_record = 'last_broken_list.yml'
drive_to_check = 'C:\\'


# https://documentation.mailgun.com/quickstart-sending.html#send-with-smtp-or-api
def send_email(subject, message):
    return requests.post(
        "https://api.mailgun.net/v3/YOUR_DOMAIN_NAME/messages",
        auth=("api", "YOUR_API_KEY"),
        data={"from": "Broken Link Bot <mailgun@YOUR_DOMAIN_NAME>",
              "to": ["bar@example.com", "YOU@YOUR_DOMAIN_NAME"],
              "subject": subject,
              "text": message})
              

def check_links():
    top_level = glob(drive_to_check + r"*.lnk")
    one_deep = glob(drive_to_check + r"*\*.lnk")
    two_deep = glob(drive_to_check + r"*\*\*.lnk")
    three_deep = glob(drive_to_check + r"*\*\*\*.lnk")
    all_paths = np.array(top_level + one_deep + two_deep + three_deep)

    shell = win32com.client.Dispatch("WScript.Shell")

    all_shorcuts = [
        shell.CreateShortCut(path)
        for path in all_paths
    ]

    target_exists = np.array([
        os.path.exists(shortcut.Targetpath)
        for shortcut in all_shorcuts
    ])

    current_broken = all_paths[np.invert(target_exists)].tolist()

    with open(broken_link_record, 'r') as file:
        previous_broken = yaml.load(file)
        
    new_broken = np.setdiff1d(current_broken, previous_broken)

    if len(new_broken) > 0:
        message = '''New broken links:
        ''' + json.dumps(new_broken.tolist()) + '''
        
All broken links:
        ''' + json.dumps(current_broken)
        send_email("New Links Broken", message)
        
    with open(broken_link_record, 'w') as outfile:
        yaml.dump(current_broken, outfile)
        

try:
    check_links()
except Exception:
    send_email("Link Checker Had Error", ''.join(traceback.format_exc()))
    raise