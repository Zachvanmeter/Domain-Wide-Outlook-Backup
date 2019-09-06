		# from/import for portability
from psutil import process_iter
from shutil import copyfile
from os import path, makedirs, system, remove
from socket import *
from glob import glob
from pathlib import Path
from multiprocessing import Process
from datetime import datetime,date
from time import sleep
from contextlib import suppress

# pyinstaller -F -I=TMF.ico CompanyOutlookBackup.py

	# Each of the users in a dict of their logon(key) 
	# and the name of their backup drive folder(value)
	# Use individual backup folders because most people
	# have named their .pst files Outlook.pst
	
userKeyDict = {
		'jwayne': 'John',
		'drenalds': 'Dennis', 
		'deerenalds': 'Dee', 
		'frenalds': 'Frank', 
		'ckelly': 'Charlie', 
		'rmcdonald': 'Mac', 
		'zvanmeter': 'Zach'
		}

floorKeyList = [	# These users are on shared computers, or Userless computers
		'weld',
		'intg', 
		'mill', 
		'inspect', 
		'cnc', 
		'shipping',  
		'lathe',  
		'receiving', 
		'tool',
		'troya',
		'maint',
		'saw',
		'toolcrib',
		'cps',
		'tblanton',
		]

deviceDict = {
		'ex304':'',
		'ex311':'',
		'ex319':'',
		'ex325':'',
		'ex326':'',
		'ex327':'',
		'ex329':'',
		'ex368':'',
		'ex374':'',
		'ext300':'',
		'ext306':'',
		'ext307':'',
		'ext308-1':'',
		'ext309':'',
		'ext312':'',
		'ext314':'',
		'ext316':'',
		'ext317':'',
		'ext318':'',
		'ext320':'',
		'ext321':'',
		'ext325':'',
		'ext326':'',
		'ext326-1':'',
		'ext330-2':'',
		'ext331':'',
		'ext332':'',
		'ext336':'',
		'ext355':'',
		'ext355-1':'',
		'ext367':'',
		'ext370':'',
		'ext373':'',
		'cmmcnc':'',
		'cmmlarge':'',
		'cnc18':'',
		'cps':'',
		'epicor356':'',
		'intg14':'',
		'lathe17':'',
		'maint15':'',
		'mill18':'',
		'receiving14':'',
		'saw18':'',
		'shipping14':'',
		'toolcrib':'',
		'weld14':'',
		'ext425':'dummy value',	# Cant ping Plant II for some reason
		}
		
def CopyPst(pathDir,DoCopy,FilePathList):
	def reverseDict(pathDir):
		deviceName, loginUser = ParsePath(pathDir)
		#print(pathDir)			# all .pst files
		if loginUser in userKeyDict:
			#print(pathDir)		# all user .pst files
			dst = GenDst('',deviceName,userKeyDict[loginUser],'\\\\tmfsvr7\\Users\\%s'%(userKeyDict[loginUser]),pathDir)
			return userKeyDict[loginUser], dst, deviceName
		else:
			#print(pathDir)		# all leftover .pst files
			dstDir = '\\\\tmfsvr7\\Users\\Floor Backups' if loginUser in floorKeyList else '\\\\tmfsvr7\\Users\\All Leftover Email Backups'
			_,_,filename = pathDir.rpartition('\\')
			dst = GenDst('%s %s'%(loginUser,filename),deviceName,loginUser,dstDir,pathDir)
			return loginUser, dst, deviceName
	def ParsePath(pathDir):
		RplStr = '\\Users\\' if '\\Users\\' in pathDir else '\\Documents and Settings\\'
		deviceName,_,tail = pathDir.partition(RplStr)
		loginUser, _,tail = tail.partition('\\')
		return deviceName.replace('\\c$',''), loginUser.lower()
	def GenDst(filename,deviceName,dstUser,dstDir,pathDir):	# Format Destination .pst file
		if filename == '':
			newfilepath = "\\\\tmfsvr7\\Users\\%s\\Email Backups\\"%(dstUser)
			_,_, filename = pathDir.replace('.pst','').rpartition('\\')
			return ('%s Backup %s %s.pst')%(newfilepath+filename,str(date.today().year),str(date.today().month))
		else: return dstDir+'\\%s %s\\'%(deviceName, dstUser)+filename
	# ############################################################ #
	dstUser, dst, deviceName = reverseDict(pathDir)	
	if DoCopy == 'Floor' and not 'floor' in dst.lower(): return
	if DoCopy == 'Test': print('Copy: '+pathDir+'\nTo: '+dst)
	if not path.isfile(dst):
		Process(target=CloseOutlook, args=(deviceName,)).start()
		Process(target=DoCopyBit, args=(dst,pathDir,dstUser,FilePathList)).start()
		print('Copying', dst)
	else: print('Already Done:', dst)
	
def CloseOutlook(deviceName): system("taskkill /s "+deviceName+" /u zvanmeter /FI \"IMAGENAME eq OUTLOOK.EXE\"")
def DoCopyBit(dst,pathDir,dstUser,FilePathList):
	with suppress(FileExistsError):
		filepath,_,_ = dst.rpartition('\\')
		makedirs(filepath)
	try:
		sleep(5)
		copyfile(pathDir, dst)
		msg = datetime.strftime(datetime.now(),'%d/%m, %H:%M')+' Backup Successful: '+pathDir
	except Exception as e:
		msg = datetime.strftime(datetime.now(),'%d/%m, %H:%M')+' Backup Failed: '+pathDir+' '+str(e)
		if path.isfile(dst): remove(dst)
	print(msg)
	with open(FilePathList[2], 'a') as f:
		f.write(msg+'\n')

def GenPstList(Deepscan):						# Find all .pst files in the domain
	def is_up(addr):							# Simple ping script
		s = socket(AF_INET, SOCK_STREAM)
		s.settimeout(0.01)
		with suppress(gaierror):
			if not s.connect_ex((addr,135)):    # connect to the remote host on port 135
				s.close()                       
				return 1
		s.close()
	def GenDeviceMap(Deepscan):
		for ip in range(1,256 if Deepscan==0 else 501):   				
			addr = '192.168.1.'+str(ip)
			if is_up(addr):						# If the ping is good, lets format our device map dictionary
				deviceName = getfqdn(addr).replace('.tmfdomain.local','').replace('.TMFDOMAIN.local','')
				if Deepscan == 1: 
					deviceDict[deviceName] = addr
				else:
					for key, value in deviceDict.items():
						if key.upper() == deviceName.upper():
							deviceDict[key] = addr
		return deviceDict
	pathDirDict = {}
	for deviceName, CompIP in GenDeviceMap(Deepscan).items():
		if not CompIP == '':
			for PstPath in [	# These three locations are the only locations we've ever found .pst files in
				'\\\\%s\\c$\\Users\\*\\Documents\\Outlook Files\\' % (deviceName),
				'\\\\%s\\c$\\Users\\*\\AppData\\Local\\Microsoft\\Outlook\\' % (deviceName),
				'\\\\%s\\c$\\Documents and Settings\\*\\\Local Settings\Application Data\Microsoft\Outlook\\' % (deviceName),
				]:
				for item in glob(PstPath+'*.ost'):
					pathDirDict[item] = CompIP
				for item in glob(PstPath+'*.pst'):
					pathDirDict[item] = CompIP
	return pathDirDict
	
def Main(FilePathList,DoCountDown=1,DoCopy=1,AllPsts=[]):
	def CountDown():
		if DoCountDown == 1:
			target = '00:01'
			while True:
				sleep(1)
				now = datetime.strftime(datetime.now(),'%H:%M')
				print(now, target,end="\r")
				if now == target: return	
		print('The countdown has been skipped.')
	def LogData(pathDir,FilePathList):			# Generate list of psts to check against
		pathDir = pathDir.lower()
		with open(FilePathList[0],'r') as f:
			lines = f.readlines()
		if not pathDir+'\n' in lines:
			print('Logging New Pst: '+pathDir)
			with open(FilePathList[0],'a') as f:
				f.write(pathDir+'\n')
	def CheckRecordedPsts(AllPsts,FilePathList):	# Now we check to see if we backed up everything we expected to
		with open(FilePathList[0],'r') as f:
			lines = f.readlines()
		print('Found: %s/%s'%(len(AllPsts),len(lines)))
		with open(FilePathList[1],'a') as f:
			for line in lines:
				if not line in AllPsts:
					print('We couldnt find: '+line.replace('\n',''))
					f.write(str(date.today())+','+line)
	# ############################################################ #
	Path(FilePathList[2]).touch()
	CountDown()
	Deepscan = 1 if DoCopy == 'Dpscn' else 0
	print('We\'re doing this for real, hide your outlook, hide your wife. cause we backing up everything out here')
	for pathDir, _ in GenPstList(Deepscan).items():
		AllPsts.append(pathDir.lower()+'\n')
		try:
			if DoCopy == 1: 								CopyPst(pathDir,DoCopy,FilePathList)
			elif DoCopy == 'Floor': 						CopyPst(pathDir,DoCopy,FilePathList)
			elif DoCopy == 'Test' and 'ex326' in pathDir:	CopyPst(pathDir,DoCopy,FilePathList)
			elif DoCopy == 'Dpscn': 						print(pathDir)
			LogData(pathDir,FilePathList)	
		except Exception as e: print(e)
	CheckRecordedPsts(AllPsts,FilePathList)	
	if DoCopy == 1: input('Press Enter to close\n')
		
if __name__ == '__main__':
	LocalDirectory  = '\\\\TMFSVR7\\Users\\Zach\\Script Backups\\Python Scripts\\Outlook Backup\\'
	FilePathList = [
			LocalDirectory+'List of All PSTs.txt',
			LocalDirectory+'Files we Missed.txt',
			LocalDirectory+'Log.txt'
			]
	DoCopyOptions = [
			0,		#0	# Do not Copy
			1,		#1	# Proceed with no special conditions
			'Floor',	#2	# Only Copy Floor Dict
			'Test',		#3	# Target Specific Computers
			'Dpscn'		#4	# Run a single iteration
			]
	Main(FilePathList,DoCountDown=0,DoCopy=DoCopyOptions[1]) 
