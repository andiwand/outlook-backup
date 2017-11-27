import os
import sys
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

appdata_roaming = os.getenv("APPDATA")
appdata_local = os.getenv("LOCALAPPDATA")

def ignore_db(path, children):
	return [child for child in children if child.endswith(".pst") or child.endswith(".ost") or child.endswith(".nst")]

def backup(path):
	dirpath = tempfile.mkdtemp()
	try:
		# save meta
		meta = {
			"username": getpass.getuser(),
			"userdir": os.path.expanduser("~"),
			"date": datetime.datetime.now().isoformat(),
			"hostname": socket.gethostname(),
		}
		tmpdst = os.path.join(dirpath, "meta.json")
		with open(tmpdst, "w") as outfile:
			json.dump(meta, outfile, indent=4, sort_keys=True)

		# save registry "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook"
		tmpsrc = r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook"
		tmpdst = os.path.join(dirpath, "registry.reg")
		r = subprocess.call(["regedit", r"/E", tmpdst, tmpsrc])
		if r != 0:
			print("!!! regedit error %d !!!" % r)
			return False

		# save "C:\Users\%username%\AppData\Local\Microsoft\Outlook" ignore pst, ost, nst
		tmpsrc = os.path.join(appdata_local, r"Microsoft\Outlook")
		tmpdst = os.path.join(dirpath, r"files\AppData\Local\Microsoft\Outlook")
		shutil.copytree(tmpsrc, tmpdst, ignore=ignore_db)

		# save "C:\Users\%username%\AppData\Roaming\Microsoft\Outlook"
		tmpsrc = os.path.join(appdata_roaming, r"Microsoft\Outlook")
		tmpdst = os.path.join(dirpath, r"files\AppData\Roaming\Microsoft\Outlook")
		shutil.copytree(tmpsrc, tmpdst)

		# save "C:\Users\%username%\AppData\Roaming\Microsoft\Signatures"
		tmpsrc = os.path.join(appdata_roaming, r"Microsoft\Signatures")
		tmpdst = os.path.join(dirpath, r"files\AppData\Roaming\Microsoft\Signatures")
		shutil.copytree(tmpsrc, tmpdst)

		path, extension = os.path.splitext(path)
		shutil.make_archive(path, "zip", dirpath)
		return True
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

		# restore registry "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook"
		# TODO: fix paths in registry.reg
		tmpsrc = os.path.join(dirpath, "registry.reg")
		r = subprocess.call(["regedit", r"/S", tmpsrc])
		if r != 0:
			print("!!! regedit error %d !!!" % r)
			return False

		# restore "C:\Users\%username%\AppData\Local\Microsoft\Outlook"
		tmpsrc = os.path.join(dirpath, r"files\AppData\Local\Microsoft\Outlook")
		tmpdst = os.path.join(appdata_local, r"Microsoft\Outlook")
		distutils.dir_util.copy_tree(tmpsrc, tmpdst)

		# restore "C:\Users\%username%\AppData\Roaming\Microsoft\Outlook"
		tmpsrc = os.path.join(dirpath, r"files\AppData\Roaming\Microsoft\Outlook")
		tmpdst = os.path.join(appdata_roaming, r"Microsoft\Outlook")
		distutils.dir_util.copy_tree(tmpsrc, tmpdst)

		# restore "C:\Users\%username%\AppData\Roaming\Microsoft\Signatures"
		tmpsrc = os.path.join(dirpath, r"files\AppData\Roaming\Microsoft\Signatures")
		tmpdst = os.path.join(appdata_roaming, r"Microsoft\Signatures")
		distutils.dir_util.copy_tree(tmpsrc, tmpdst)
		
		return True
	finally:
		shutil.rmtree(dirpath)

def parse_args(args=None):
	parser = argparse.ArgumentParser(description="Backup and restore script for Microsoft Outlook.")
	parser.add_argument("path", help="path to archive")
	parser.add_argument("-r", "--restore", help="don't start services", action="store_true")
	return parser.parse_args(args)

def main():
	args = parse_args()
	
	if args.restore:
		r = restore(args.path)
	else:
		r = backup(args.path)
	
	return 0 if r else 1

if __name__ == "__main__":
	exit = main()
	sys.exit(exit)
