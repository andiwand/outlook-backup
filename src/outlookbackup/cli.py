# regedit reference: http://www.robvanderwoude.com/regedit.php

import os
import sys
import ctypes
import tempfile
import shutil
import subprocess
import json
import getpass
import datetime
import socket
import zipfile
import distutils.dir_util
import argparse
import re

VERSION = "1.0.0"
APPDATA_ROAMING = os.getenv("APPDATA")
APPDATA_LOCAL = os.getenv("LOCALAPPDATA")

def ignore_db(path, children):
	return [child for child in children if child.endswith(".pst") or child.endswith(".ost") or child.endswith(".nst")]

def fix_registry(path, meta):
	olduserdir = meta["userdir"].replace("\\", "\\\\")
	newuserdir = get_meta()["userdir"].replace("\\", "\\\\")
	with open(path, "r", encoding="utf-16-le") as infile:
		lines = infile.readlines()
	with open(path, "w", encoding="utf-16-le") as outfile:
		for line in lines:
			oldline = line
			line = line.replace(olduserdir, newuserdir)
			outfile.write(line)

def get_meta():
	return {
		"version": VERSION,
		"username": getpass.getuser(),
		"userdir": os.path.expanduser("~"),
		"date": datetime.datetime.now().isoformat(),
		"hostname": socket.gethostname(),
	}

def backup(path):
	dirpath = tempfile.mkdtemp()
	try:
		# save meta
		meta = get_meta()
		tmpdst = os.path.join(dirpath, "meta.json")
		with open(tmpdst, "w") as outfile:
			json.dump(meta, outfile, indent=4, sort_keys=True)

		# save registry "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook"
		tmpsrc = r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook"
		tmpdst = os.path.join(dirpath, "registry.reg")
		r = subprocess.call(["regedit", r"/E", tmpdst, tmpsrc])
		if r != 0:
			print("!!! regedit error %d !!!" % r)
			return 1

		# save "C:\Users\%username%\AppData\Local\Microsoft\Outlook" ignore pst, ost, nst
		tmpsrc = os.path.join(APPDATA_LOCAL, r"Microsoft\Outlook")
		tmpdst = os.path.join(dirpath, r"files\AppData\Local\Microsoft\Outlook")
		shutil.copytree(tmpsrc, tmpdst, ignore=ignore_db)

		# save "C:\Users\%username%\AppData\Roaming\Microsoft\Outlook"
		tmpsrc = os.path.join(APPDATA_ROAMING, r"Microsoft\Outlook")
		tmpdst = os.path.join(dirpath, r"files\AppData\Roaming\Microsoft\Outlook")
		shutil.copytree(tmpsrc, tmpdst)

		# save "C:\Users\%username%\AppData\Roaming\Microsoft\Signatures"
		tmpsrc = os.path.join(APPDATA_ROAMING, r"Microsoft\Signatures")
		tmpdst = os.path.join(dirpath, r"files\AppData\Roaming\Microsoft\Signatures")
		shutil.copytree(tmpsrc, tmpdst)

		path, extension = os.path.splitext(path)
		shutil.make_archive(path, "zip", dirpath)
		return 0
	finally:
		shutil.rmtree(dirpath)

def restore(path):
	dirpath = tempfile.mkdtemp()
	try:
		zip = zipfile.ZipFile(path, "r")
		zip.extractall(dirpath)
		zip.close()

		tmpsrc = os.path.join(dirpath, "meta.json")
		with open(tmpsrc, "r") as infile:
			meta = json.load(infile)

		# TODO: clear registry tree, clear folders

		# restore registry "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook"
		tmpsrc = os.path.join(dirpath, "registry.reg")
		fix_registry(tmpsrc, meta)
		r = subprocess.call(["regedit", r"/S", tmpsrc])
		if r != 0:
			print("!!! regedit error %d !!!" % r)
			return 1

		# restore "C:\Users\%username%\AppData\Local\Microsoft\Outlook"
		tmpsrc = os.path.join(dirpath, r"files\AppData\Local\Microsoft\Outlook")
		tmpdst = os.path.join(APPDATA_LOCAL, r"Microsoft\Outlook")
		distutils.dir_util.copy_tree(tmpsrc, tmpdst)

		# restore "C:\Users\%username%\AppData\Roaming\Microsoft\Outlook"
		tmpsrc = os.path.join(dirpath, r"files\AppData\Roaming\Microsoft\Outlook")
		tmpdst = os.path.join(APPDATA_ROAMING, r"Microsoft\Outlook")
		distutils.dir_util.copy_tree(tmpsrc, tmpdst)

		# restore "C:\Users\%username%\AppData\Roaming\Microsoft\Signatures"
		tmpsrc = os.path.join(dirpath, r"files\AppData\Roaming\Microsoft\Signatures")
		tmpdst = os.path.join(APPDATA_ROAMING, r"Microsoft\Signatures")
		distutils.dir_util.copy_tree(tmpsrc, tmpdst)
		
		return 0
	finally:
		shutil.rmtree(dirpath)

def parse_args(args=None):
	parser = argparse.ArgumentParser(description="Backup and restore script for Microsoft Outlook settings.")
	parser.add_argument("path", help="path to archive")
	parser.add_argument("-r", "--restore", help="import instead of export", action="store_true")
	return parser.parse_args(args)

def is_admin():
	try:
		return ctypes.windll.shell32.IsUserAnAdmin()
	except:
		return False

def main():
	if not is_admin():
		print("need administrator privileges")
		return 2

	args = parse_args()

	if args.restore:
		r = restore(args.path)
	else:
		r = backup(args.path)

	return r

if __name__ == "__main__":
	exit = main()
	sys.exit(exit)
