import os
import tempfile
import shutil
import subprocess
import json
import getpass
import datetime

def ignore_db(path, children):
	return [child for child in children if child.endswith(".pst") or child.endswith(".ost") or child.endswith(".nst")]

appdata_roaming = os.getenv("APPDATA")
appdata_local = os.getenv("LOCALAPPDATA")

dirpath = tempfile.mkdtemp()

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

# save registry "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook"
tmpsrc = r"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook"
tmpdst = os.path.join(dirpath, "registry.reg")
r = subprocess.call(["regedit", r"/E", tmpdst, tmpsrc])
if r != 0: print("!!! regedit error %d !!!" % r)

# save meta
meta = {
	"username": getpass.getuser(),
	"userdir": os.path.expanduser("~"),
	"date": datetime.datetime.now().isoformat(),
}
tmpdst = os.path.join(dirpath, "meta.json")
with open(tmpdst, "w") as outfile:
	json.dump(meta, outfile, indent=4, sort_keys=True)

shutil.make_archive("backup", "zip", dirpath)
shutil.rmtree(dirpath)
