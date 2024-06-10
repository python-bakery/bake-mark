"""
Library based off of the AWS Polly example code, with some modifications for this project.

Basically, this provides functions for generating speech from text using AWS Polly, and saving the resulting audio files to disk.
"""
from contextlib import closing
import os
import sys
import subprocess
import json
from tempfile import gettempdir
import shutil
from textwrap import fill
from datetime import datetime

from boto3 import Session
from botocore.exceptions import BotoCoreError, ClientError
from pydub import AudioSegment

from friendly_hash import hash
from locations import VOICES_DIR, DUBS_FILE_PATH, BACKUP_DUBS_FILE_PATH, USED_DUBS_FILE_PATH

def make_default_files():
    # Make sure voices directory exists
    os.makedirs(VOICES_DIR, exist_ok=True)
    # Make sure dub index and its backup exists
    for path in [DUBS_FILE_PATH, BACKUP_DUBS_FILE_PATH]:
        if not os.path.exists(path):
            with open(path, 'w') as dub_file:
                json.dump({}, dub_file)

make_default_files()

def add_dub_entry(hash_text, text):
    try:
        with open(DUBS_FILE_PATH) as dub_file:
            existing = json.load(dub_file)
    except json.JSONDecodeError as e:
        raise Exception("Error while loading dub index; perhaps corrupted? Check the backup!\nOriginal error was:", str(e))
    # Is the update actually needed?
    if hash_text in existing and existing[hash_text] == text:
        return
    # Make a backup in case things get interrupted and the file is corrupted
    shutil.copy(DUBS_FILE_PATH, BACKUP_DUBS_FILE_PATH)
    # Update the index file then!
    existing[hash_text] = text
    with open(DUBS_FILE_PATH, 'w') as dub_file:
        json.dump(existing, dub_file, indent=4)

def reencode_mp3(path):
    """
    MP3s were getting corrupted, strangely enough. This function re-encodes the file to fix that.
    """
    song = AudioSegment.from_mp3(path)
    song.export(path, format="mp3")
    
    
def remember_used(label, hash_name):
    if not os.path.exists(USED_DUBS_FILE_PATH):
        with open(USED_DUBS_FILE_PATH, 'w') as used_file:
            json.dump({}, used_file)
    with open(USED_DUBS_FILE_PATH) as used_file:
        existing = json.load(used_file)
    if hash_name not in existing:
        existing[hash_name] = []
    existing[hash_name].append({"label": label, "when": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
    with open(USED_DUBS_FILE_PATH, 'w') as used_file:
        json.dump(existing, used_file, indent=4)


def speech(text, voice, use_remote=True, label=""):
    hash_name = "speech"+str(hash(text))
    remember_used(label, hash_name)
    output = os.path.join(VOICES_DIR, voice, hash_name+'.mp3')
    if os.path.exists(output):
        # Might need to update the index file!
        add_dub_entry(hash_name, text)
        return output
    if not use_remote:
        raise Exception(f"Local speech file {output!r} missing for voice {voice!r}. Text of speech was:\n"+
                        fill(text, initial_indent='    ', subsequent_indent='    '))

    session = Session(profile_name="default")
    polly = session.client("polly")
    try:
        # Request speech synthesis
        response = polly.synthesize_speech(Text=text, OutputFormat="mp3",
                                            VoiceId=voice, Engine="neural")
    except (BotoCoreError, ClientError) as error:
        # The service returned an error, exit gracefully
        print(error)
        sys.exit(-1)
    if "AudioStream" in response:
        with closing(response["AudioStream"]) as stream:
            try:
                # Open a file for writing the output as a binary stream
                with open(output, "wb") as file:
                    file.write(stream.read())
                add_dub_entry(hash_name, text)
                return output
            except IOError as error:
                # Could not write to file, exit gracefully
                print(error)
                sys.exit(-1)
    else:
        # The response didn't contain audio data, exit gracefully
        print("Could not stream audio")
        sys.exit(-1)
