		#from/import for portability
from psutil import process_iter
from shutil import copyfile
from os import path, makedirs, system, remove
from datetime import date
from socket import *
from glob import glob
from pathlib import Path
from multiprocessing import Process
from datetime import datetime
from time import sleep


#pyinstaller --onefile --icon=TMF.ico CompanyOutlookBackup.py


	# Each of the users in a dict of their logon(key) 
	# and the name of their backup drive folder(value)
	# Use individual backup folders because most people
	# have named their .pst files Outlook.pst
  
userKeyDict = {
		'zvanmeter': 'Zach'
		}

	# These users are shared computers, or Userless computers
floorKeyDict = {
		'Floor':'weld',
		'Floor':'intg', 
		'Floor':'INTG', 
		'Floor':'mill', 
		'Floor':'inspect', 
		'Floor':'cnc', 
		'Floor':'shipping', 
		'Floor':'lathe',  
		'Floor':'lathe',  
		'Floor':'lwhite', 
		'Floor':'receiving', 
		'Floor':'tool',
		'Floor':'troya',
		'Floor':'maint',
		'Floor':'Maint',
		'Floor':'saw',
		'Floor':'toolcrib'
		}	

def is_up(addr):
		# Simple ping script
	s = socket(AF_INET, SOCK_STREAM)
	s.settimeout(0.01)    ## set a timeout of 0.01 sec
	if not s.connect_ex((addr,135)):    # connect to the remote host on port 135
		s.close()                       ## (port 135 is always open on Windows machines, AFAIK)
		return 1
	s.close()

def GenDeviceDict():
		# These are all of the machines we care to check
	deviceDict = {
		'ext336':'',
		'ext316':'',
		'ext355':'',
		'ext355-1':'',
		'ext306':'',
		'ext317':'',
		'ex368':'',
		'ext367':'',
		'ext325':'',
		'ext307':'',
		'ex319':'',
		'ex329':'',
		'ex325':'',
		'ex326':'',
		'ex327':'',
		'ext326':'',
		'ext326-1':'',
		'ext300':'',
		'EXT330-2':'',
		'ext314':'',
		'ext321':'',
		'ext331':'',
		'maint15':'',
		'shipping14':'',
		'mill18':'',
		'ext309':'',
		'ext312':'',
		'ex304':'',
		'epicor356':'',
		'ext308-1':'',
		'ext318':'',
		'cmm2':'',
		'cmmlarge':'',
		'romerlaptop':'',
		'ext425':'',
		'cnc18':'',
		'intg14':'',
		'saw18':'',
		'lathe17':'',
		'weld14':'',
		'RECEIVING14':'',
		'TOOLCRIB':''
		}
	return deviceDict

def GenDeviceMap():
	print(' ')
	
		# Only look at computers we actually care about
	deviceDict = GenDeviceDict()
	
		# 'ping' addresses 192.168.1.1 to .1.255
	for ip in range(1,256):    
		addr = '192.168.1.'+str(ip)
			
			# If the ping is good, lets format our device map dictionary
		if is_up(addr):
			deviceName = getfqdn(addr).replace('.tmfdomain.local','')
			deviceName = deviceName.replace('.TMFDOMAIN.local','') #it returns in caps on users comp
			for key, value in deviceDict.items():
				if key.upper() == deviceName.upper():
					deviceDict[key] = addr
	return deviceDict	

def CloseOutlook(CompIP,dstUser):	#$ dstUser added for test
	print('Terminating outlook.')
	system("taskkill /s "+CompIP+" /u zvanmeter /FI \"IMAGENAME eq OUTLOOK.EXE\"")

def copyPst(pathDir, dstDir, dstUser, filename, deviceName, CompIP):
		# Kill Outlook on target computer so that we can read the .pst
	if not CompIP == '':
		Process(target=CloseOutlook, args=(CompIP,dstUser)).start()
		sleep(5)
		
		# Format Destination .pst file
	if filename == '':
		newfilepath = "\\\\tmfsvr7\\Users\\%s\\Email Backups\\"%(dstUser)
		head, sep, filename = pathDir.rpartition('\\')
		filename, sep, tail = filename.partition('.pst')
		d = date.today()
		dst=newfilepath+filename+" Backup "+str(d.month)+' '+str(d.day)+".pst"
	else:
		newfilepath = dstDir+'\\%s %s\\'%(deviceName, dstUser)
		dst=newfilepath+filename
		
		# Only launch if we don't already have a backup
	if not path.isfile(dst):
		print('Launching MP')
		Process(target=DoCopyBit, args=(newfilepath,dst,pathDir,dstUser)).start()

def DoCopyBit(newfilepath,dst,pathDir,dstUser):
		# Make a new destination folder and local log file if this 
		# is the first time
	try:
		makedirs(newfilepath)
	except FileExistsError:
		pass
	Path('Log.txt').touch()
	
		# Now lets back up the file, and log success or errors
	try:
		copyfile(pathDir, dst)
		now = datetime.strftime(datetime.now(),'%d/%m, %H:%M')
		msg = now+' Backup Successful: '+pathDir
	except PermissionError as e:		
		now = datetime.strftime(datetime.now(),'%d/%m, %H:%M')
		msg = now+' Backup Failed, I probably couldnt close Outlook: '+pathDir+' '+str(e)
		remove(dst)
	except Exception as e:		
		now = datetime.strftime(datetime.now(),'%d/%m, %H:%M')
		msg = now+' Backup Failed: '+pathDir+' '+str(e)
		remove(dst)
	print(msg)
	with open('Log.txt', 'a') as f:
		f.write(msg+'\n')
	return


def getPstList():
	pathDirDict = {}
	for deviceName, CompIP in GenDeviceMap().items():
		if not CompIP == '':
			print(CompIP, '=', deviceName)
			
			# These three locations are the only locations we've ever found .pst files in
		Locations = [
			'\\\\%s\\c$\\Users\\*\\Documents\\Outlook Files\\*.pst' % (deviceName),
			'\\\\%s\\c$\\Users\\*\\AppData\\Local\\Microsoft\\Outlook\\*.pst' % (deviceName),
			'\\\\%s\\c$\\Documents and Settings\\*\\\Local Settings\Application Data\Microsoft\Outlook\\*.pst' % (deviceName)
			]
			
			# Get all .pst files and link them with the IP of each comp
		for PstPath in Locations:
			globList = glob(PstPath)
			for item in globList:
				pathDirDict[item] = CompIP
				
		'''		# Really really dont use this often
			# Otherwise, use the deepscan
		alphabet = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P',
			'Q','R','S','T','U','V','W','S','Y','Z']
		for letter in alphabet:
			globList = glob('\\\\%s\\%s$\\**/*.pst'%(deviceName,letter), recursive=True)	
			for item in globList:
				pathDirDict[item] = CompIP
		'''
	return pathDirDict
	
def reverseDict(loginUser,pathDir):	#dstDir, dstUser = reverseDict(loginUser)
	if loginUser in userKeyDict:
		#print(pathDir)		#all user .pst files
		dstUser = list(userKeyDict.values())[list(userKeyDict.keys()).index(loginUser)]
		return '\\\\tmfsvr7\\Users\\%s'%(dstUser), dstUser, ''
		
	else:
		if loginUser in floorKeyDict:
			dstDir = '\\\\TMFSVR7\\Users\\Floor Backups'	
			#print(pathDir)		#all floor .pst files
		else:
			dstDir = '\\\\TMFSVR7\\Users\\All Leftover Email Backups'
			#print(pathDir)		#all unallocated .pst files
		head,sep,filename = pathDir.rpartition('\\')
		filename = '%s %s' %(loginUser, filename)
		return dstDir, loginUser, filename

def CountDown():
		# Wait to continue until 12:01 am
	target = '00:01'
	while True:
		sleep(1)
		now = datetime.strftime(datetime.now(),'%H:%M')
		print(now, target,end="\r")
		if now == target:
			Main()
			break

def Main():
		# Find all .pst files in the domain
	for pathDir, CompIP in getPstList().items():
		#print(pathDir)		#all .pst files
		
			# Format variables
		if '\\Users\\' in pathDir:
			head,sep,tail = pathDir.partition('\\Users\\')
		elif '\\Documents and Settings\\' in pathDir:
			head,sep,tail = pathDir.partition('\\Documents and Settings\\')
		deviceName = head.replace('\\','').replace('c$','')
		loginUser,sep,tail = tail.partition('\\')
		dstDir, dstUser, filename = reverseDict(loginUser,pathDir)
		
			# Proceed to copy .pst files
		copyPst(pathDir, dstDir, dstUser, filename, deviceName, CompIP)

		
if __name__ == '__main__':
	#CountDown()
	Main()
				
	while True:
		sleep(1)
		
	
