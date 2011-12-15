import sys
import os
import string
import re
import ConfigParser
import logging
import urllib
import zipfile
import smtplib
import time
import hashlib
import shutil
import quopri
import getopt
from datetime import datetime


sys.path.append(os.path.dirname(os.path.abspath(__file__)) + '/lib/')
import feedparser

PRG_VERSION = '0.6.1 EatMyShorts'
PRG_NAME, _ = os.path.splitext(os.path.basename(__file__))
WRK_DIR = os.path.dirname(os.path.abspath(__file__))
TMP_DIR = WRK_DIR + '/tmp'
LOG_DIR = WRK_DIR + '/log'
CFG_FILE = WRK_DIR + '/' + PRG_NAME + '.ini'

class Logger:

	def __init__(self, file):

		self.logger = logging.getLogger('logging')
	
		if config['log.level'] == 'DEBUG':
			self.logger.setLevel(logging.DEBUG)
		elif config['log.level'] == 'WARNING':
			self.logger.setLevel(logging.WARN)
		elif config['log.level'] == 'ERROR':
			self.logger.setLevel(logging.ERROR)
		else:
			self.logger.setLevel(logging.INFO)

		formatter = logging.Formatter('%(asctime)s\t%(levelname)s\t%(message)s')
		
		self.fh = logging.FileHandler(file)
		self.fh.setFormatter(formatter)
		self.logger.addHandler(self.fh)

		if config['log.onscreen']:
			self.console = logging.StreamHandler()
			self.console.setFormatter(formatter)
			self.logger.addHandler(self.console)

	def debug(self, msg):
		self.logger.debug(msg)

	def info(self, msg):
		self.logger.info(msg)

	def warn(self, msg):
		self.logger.warn(msg)

	def error(self, msg):
		self.logger.error(msg)

def fixBadZipfile(zipFile):  
	
     f = open(zipFile, 'r+b')  
     data = f.read()  
     pos = data.find('\x50\x4b\x05\x06') # End of central directory signature  
     if (pos > 0):  
         log.debug("Trancating file at location " + str(pos + 22)+ ".")  
         f.seek(pos + 22)   # size of 'ZIP end of central directory record' 
         f.truncate()  
         f.close() 
 
def md5Sum(file):

	fileObj = open(file, 'rb')
	m = hashlib.md5()
	while True:
		d = fileObj.read(8096)
		if not d:
			break
		m.update(d)

	return m.hexdigest()

def loadConfig(file):

	if not os.path.isfile(file):
		print('Unable to find configuration file %s' % file)
		sys.exit()

	config = ConfigParser.RawConfigParser()
	cfg = {}

	try:

		config.read(file)

		cfg['shows.path'] = config.get('shows', 'path')
		
		cfg['rss.url'] = config.get('rss', 'url')
		cfg['rss.backlog'] = config.get('rss', 'backlog')
		cfg['rss.mapping'] = config.get('rss', 'mapping')
		
		cfg['process.file_ext'] = config.get('process', 'file_ext').translate(string.maketrans('', ''),' ').split(',')
		cfg['process.tags'] = config.getboolean('process', 'tags')
		cfg['process.exclude'] = config.get('process', 'exclude').translate(string.maketrans(',', '|'), ' ')
		cfg['process.rename'] = config.getboolean('process', 'rename')
		cfg['process.test'] = config.getboolean('process', 'test')
			
		cfg['email.notify'] = config.getboolean('email', 'notify')
		cfg['email.from'] = config.get('email', 'from').strip()
		cfg['email.to'] = config.get('email', 'to').strip()

		cfg['smtp.server'] = config.get('smtp', 'server')
		cfg['smtp.port'] = config.getint('smtp', 'port')
		cfg['smtp.username'] = config.get('smtp', 'username')
		cfg['smtp.password'] = config.get('smtp', 'password')

		cfg['log.level'] = config.get('log', 'level')
		cfg['log.onscreen'] = config.getboolean('log', 'onscreen')

	except ConfigParser.NoOptionError, error:
		print('Missing mandatory parameter: %s' % error)
		sys.exit()
	except Exception, error:
		print('Error while reading config: %s' % error)
		sys.exit()

	return cfg

def sendNotification(subject, body):

	log.debug('Notification...')

	if config['email.notify']:

		now = datetime.now()
		day = now.strftime('%a')
		date = now.strftime('%d %b %Y %X')

		header  = 'To: ' + config['email.to'] + '\n'
		header += 'From: ' + config['email.from'] + '\n'
		header += 'Date : ' + day + ', ' + date + ' -0000\r\n'
		header += 'Subject: ' + subject + '\n'
		header += 'X-priority: 5\n'
		header += 'X-MS-priority: 5\n'
		header += 'MIME-Version: 1.0\n'
		header += 'Content-Type: text/plain; charset="iso-8859-1"\n'
		header += 'Content-Transfer-Encoding: quoted-printable\n'

		msg = header + '\n' + quopri.encodestring(body) + '\n\n'

		server = config['smtp.server']
		port = config['smtp.port']

		try:
			log.debug('Connecting to SMTP server %s:%s with SSL' % (server, port) )
			smtpserver = smtplib.SMTP_SSL(server, port)
			smtpserver.ehlo()
		except Exception, error:
			log.debug('SSL connection failed: %s ' % error)
			try:
				log.debug('Reconnecting to SMTP server %s:%s without SSL' % (server, port) )
				smtpserver = smtplib.SMTP(server, port)
				smtpserver.ehlo()
			except Exception, error:
				log.error('Failed to connect to SMTP server: %s ' % error)
				return
			else:
				log.debug('Connected to SMTP server %s:%s without SSL' % (server, port) )
		else:
			log.debug('Connected to SMTP server %s:%s with SSL' % (server, port) )

		if smtpserver.ehlo_resp and re.search('STARTTLS', smtpserver.ehlo_resp, re.IGNORECASE):
			try:
				smtpserver.starttls()
				smtpserver.ehlo
			except Exception, error:
				log.error('Unable to start TLS: %s ' % error)
				return
			else:
				log.debug('TLS sarted')
		else:
			log.debug('No TLS needed')

		if config['smtp.username'] != '' and config['smtp.password'] != '':
			try:
				smtpserver.login(config['smtp.username'], config['smtp.password'])
				smtpserver.ehlo
			except smtplib.SMTPAuthenticationError, error:
				log.error('Authentication error: %s' % error)
				return
			except Exception, error:
				log.error('Unexpected error during authentication: %s' % error)
				return
			else:
				log.debug('Authentication successful')
		else:
			log.debug('No authentication configured')

		try:
			smtpserver.sendmail(config['email.from'], config['email.to'], msg)
		except smtplib.SMTPHeloError, error:
			log.error('The server refused our HELO message: %s' % error )
			return
		except smtplib.SMTPRecipientsRefused, error:
			log.error('All recipient addresses refused: %s' % error)
			return
		except smtplib.SMTPSenderRefused, error:
			log.error('Sender address refused: %s' % error )
			return
		except smtplib.SMTPDataError, error:
			log.error('The SMTP server refused to accept the message data: %s' % error )
			return
		except Exception, error:
			log.error('Unexpected error when sending message: %s' % error )
			return
		else:
			log.info('Notification email sent')

		try:
			smtpserver.close()
		except Exception, error:
			log.error('Unexpected error when closing connection:' % error )
		else:
			log.debug('Successfuly disconnected from SMTP server')
	
	else:
		log.debug('No email sent as notification is disabled')

def listShows(path, forBackLog=False):

	if not forBackLog:
		trans = string.maketrans(' ', '.')
	else:
		trans = string.maketrans(' ', '_')

	shows = {}
	try:
		for show_name in os.listdir(path):
			if os.path.isdir(path + show_name) == True:
				sanitized_show_name = show_name.translate(trans, '\'().!').lower()
				shows[sanitized_show_name] = show_name
	except:
		log.error('Unable to find ' + path)
		sys.exit()

	items = config['rss.mapping'].split(',')
	if len(items) > 0:
		for item in items:
			i = item.split('=')
			if len(i) > 0:
				local = i[0].strip()
				dist = i[1].strip().translate(trans, '\'().!').lower()
				shows[dist] = local	
				log.debug('Extra TV shows mapping : %s => %s' % (dist, local) )

	return shows

def parseRSS(url):
	
	subtitles = {}
	d = feedparser.parse(url)
	for i in range(0, len(d['entries'])):
		filename = d['entries'][i].title
		url = d['entries'][i].link
		subtitles[filename] = url
	
	return subtitles

def episodeAttributes(file, show):

	attr = {}
	season = number = alt_number = -1
	release = subrelease = alt_release = 'none'
	double = False

	# Episode season, number & release
	match = re.match('^(.+)\.s(\d+)e(\d+)e(\d+)\..*-([^\. \-]+)[\. \-]', file, re.IGNORECASE) \
			or re.match('^(.+)\.(\d+)(\d\d)(\d\d)\..*-([^\. \-]+)[\. \-]', file, re.IGNORECASE)
	if match:
		season = int(match.group(2))
		number = int(match.group(3))
		alt_number = int(match.group(4))
		release = match.group(5)
		double = True
	else:
		match = re.match('^(.+)\.s(\d+)e(\d+)\..*-([^\. \-]+)[\. \-]', file, re.IGNORECASE) \
				or re.match('^(.+)\.(\d+)(\d\d)\..*-([^\. \-]+)[\. \-]', file, re.IGNORECASE) \
				or re.match('^(.+)\.(\d+)x(\d+)\..*-([^\. \-]+)[\. \-]', file, re.IGNORECASE)
		if match:
			season = int(match.group(2))
			number = int(match.group(3))
			release = match.group(4)
			alt_number = number

	# Episode subrelease
	match = re.match('.+(proper|repack).+', file, re.IGNORECASE)
	if match:
		subrelease = match.group(1)
	else:
		subrelease = 'none'

	attr['file'] = file
	attr['show'] = show
	attr['season'] = season
	attr['double'] = double
	attr['number'] = number
	attr['alt_number'] = alt_number
	attr['release'] = release.lower()
	attr['alt_release'] = release[0:3].lower()
	if attr['alt_release'] == "p0w":
		attr['alt_release2'] = "pow"
	else:
		attr['alt_release2'] = attr['alt_release']
	attr['subrelease'] = subrelease.lower()

 	return attr

def saveSubtitleFiles(files, path, ep, url, backlog):

	saved = list()

	for member in files:

		log.info('Processing subtitle file: %s' % member)
		log.debug('%s will be saved to %s' % (member, path))

		notif_msg  = 'File %s found in RSS feed\n\n' % os.path.basename(url)
		notif_msg += 'Episode: ' + ep['file'] + ' matches season and episode number. Release: ' + ep['release'] + ' SubRelease: ' + ep['subrelease'] + '\n\n'
		notif_msg += 'Subtitle file: %s is suitable\n\n' % os.path.relpath(member, TMP_DIR)

		if config['process.rename'] and len(files) == 1:
			dest_file_name, _ = os.path.splitext(os.path.basename(ep['file']))
			_, dest_file_ext = os.path.splitext(os.path.basename(member))
			dest_file = dest_file_name + dest_file_ext
			log.info('Rename feature is on and only one sub found. Destination name is %s' % dest_file)
			notif_msg += 'Rename feature is on and only one sub found. Destination name is %s\n' % dest_file
		else:
			dest_file = os.path.basename(member)
			log.info('Rename feature is off or more than one sub found. Destination name is %s' % dest_file)
			notif_msg += 'Rename feature is off or more than one sub found. Destination name is %s\n' % dest_file

		if os.path.isfile(path + '/' + dest_file):

			log.info('%s already exists in %s' % (dest_file, path))
			notif_msg += '%s already exists in %s\n' % (dest_file, path)

			if md5Sum(path + '/' + dest_file) != md5Sum(member):
				log.info('File in the archive is different. Overwriting local file')
				notif_msg += 'File in the archive is different. Overwriting local file\n'
				if not config['process.test']:
					shutil.move(member, path + '/' + dest_file)
					if not backlog:
						sendNotification('%s has retrieved updated subtitles for %s S%dE%02d' % (PRG_NAME, ep['show'], ep['season'], ep['number']), notif_msg)
					else:
						saved.append(dest_file)

			else:
				log.info('File in the archive is identical. Keeping local file')

		else:

			log.info('%s does not exist in %s. Saving it' % (dest_file, path))
			notif_msg += '%s does not exist in %s. Saving it\n' % (dest_file, path)
			if not config['process.test']:
				shutil.move(member, path + '/' + dest_file)
				if not backlog:
					sendNotification('%s has retrieved subtitles for %s S%dE%02d' % (PRG_NAME, ep['show'], ep['season'], ep['number']), notif_msg)
				else:
					saved.append(dest_file)

	return saved

def cleanSuitableFilesList(suitables, ep):
	
	release = ep['release']
	alt_release = ep['alt_release']
	subrelease = ep['subrelease']
	tagsWanted = config['process.tags']

	tagsFound = noTagsFound = properFound = repackFound = False

	for suitable in suitables:
		if re.match('.+[\.\- ]Tag[\.\- ]', suitable, re.IGNORECASE):
			tagsFound = True
		if re.match('.+[\.\- ]NoTag[\.\- ]', suitable, re.IGNORECASE):
			noTagsFound = True
		if re.match('.+[\.\- ]proper[\.\- ]', suitable, re.IGNORECASE):
			properFound = True
		if re.match('.+[\.\- ]repack[\.\- ]', suitable, re.IGNORECASE):
			repackFound = True
	log.debug('tagsFound=%s noTagsFound=%s properFound=%s repackFound=%s' % (tagsFound, noTagsFound, properFound, repackFound))

	log.debug('Browsing suitable file(s) for cleanup')
	tmp_suitables = list(suitables)
	for i in range(0, len(tmp_suitables)):
		f = tmp_suitables[i]
		if ( ( subrelease == 'proper' and ( ( properFound and re.match('.+[\.\- ]proper[\.\- ]', f, re.IGNORECASE) ) or ( not properFound and not re.match('.+[\.\- ]repack[\.\- ]', f, re.IGNORECASE) ) ) ) 
		    or ( subrelease == 'repack' and ( ( repackFound and re.match('.+[\.\- ]repack[\.\- ]', f, re.IGNORECASE) ) or ( not repackFound and not re.match('.+[\.\- ]proper[\.\- ]', f, re.IGNORECASE) ) ) )\
			or ( subrelease == 'none' and not re.match('.+[\.\- ](proper|repack)[\.\- ]', f, re.IGNORECASE) ) ) \
		and ( ( tagsWanted and ( ( tagsFound and re.match('.+[\.\- ]Tag[\.\- ]', f, re.IGNORECASE) ) or ( not tagsFound and not re.match('.+[\.\- ]NoTag[\.\- ]', f, re.IGNORECASE) ) ) )  \
			or ( not tagsWanted and ( ( noTagsFound and re.match('.+[\.\- ]NoTag[\.\- ]', f, re.IGNORECASE) ) or ( not noTagsFound and not re.match('.+[\.\- ]Tag[\.\- ]', f, re.IGNORECASE) ) ) ) ):
			log.info('[X] File %s' % f)
		else:
			log.debug('[ ] File %s' % f)
			suitables.remove(tmp_suitables[i])

	log.info('%s suitable file(s) found after cleanup' % len(suitables))

	return suitables

def getSuitableFilesListFromZip(file, ep, file_ext, prefix='', indent=''):

	matches = list()

	release = ep['release']
	alt_release = ep['alt_release']
	alt_release2 = ep['alt_release2']

	if prefix == '':
		prefix, _ = os.path.splitext(file)

	log.info('%sZip file is \'%s\'' % (indent, file))
	log.debug('%sPrefix is \'%s\'' % (indent, prefix))

	if not zipfile.is_zipfile(file):
		log.debug('Trying to repair %s' % file)
		fixBadZipfile(file)

	try:
		zip = zipfile.ZipFile(file)
	except:
		log.error('Error when opening zip file %s' % file)
		return
	
	for info in zip.infolist():
		rep, ext = os.path.splitext(info.filename)
		if ext == ".zip":
			zip.extract(info, prefix)
			matches += getSuitableFilesListFromZip(prefix + '/' + info.filename, ep, file_ext, prefix + '/' + rep, indent+'  ')
			os.remove(prefix + '/' + info.filename)
		else:
			full_name = prefix + '/' + unicode(info.filename, 'iso-8859-1')
			if not re.match('.+[\.\- ](' + config['process.exclude'] + ')[\.\- ]', full_name, re.IGNORECASE) \
			and re.match('.+[\.\-\ ](' + release + '|' + alt_release + '|' + alt_release2 + ')[\.\-\ ].*(' + file_ext + ')$', full_name, re.IGNORECASE):
				log.info(indent + '[X] File ' + os.path.relpath(full_name, TMP_DIR))
				matches.append(full_name)
				zip.extract(info, prefix)
			else:
				log.debug(indent + '[ ] File ' + os.path.relpath(full_name, TMP_DIR))

	return matches

def process(url, backlog):

	snatched = list()

	log.info('Path to TV Shows is %s' % config['shows.path'])

	shows = listShows(config['shows.path'])
	log.debug(shows)

	log.info('Retrieving RSS from %s' % url)
	subs = parseRSS(url)
	log.info('%d items retrieved from RSS' % len(subs))


	log.info('Test mode: %s' % config['process.test'])

	for item in subs.keys():
		log.info('---')
		log.info('Processing ' + item)
		match = re.match('^(.+)\.(\d+)x(\d+)', item)
		if match:
			show = match.group(1).lower()
			season = int(match.group(2))
			ep_number = int(match.group(3))
			log.info('%s => Show:%s Season:%d Episode:%d' % (item, show, season, ep_number))
			if shows.has_key(show):
				log.info(show + ' is in my watch list => ' + shows[show])
				path = '%s%s/Season %02d/' % (config['shows.path'], shows[show], season)
				if os.path.isdir(path):
					log.info('Scan in progress : ' + path)
					for file in os.listdir(path):
						if not file.startswith('.') and ( file.endswith('.avi') or file.endswith('.mkv')):
							ep_attr = episodeAttributes(file, shows[show])

							if ( ep_attr['number'] == ep_number or ep_attr['alt_number'] == ep_number ) and ep_attr['season'] == season:
								log.info('[X] Episode: ' + file + " matches season and episode number. Release: " + ep_attr['release'] + " SubRelease: " + ep_attr['subrelease'])

								url = subs[item]
								log.info('Downloading ' + url)
								local_zip_name = TMP_DIR + '/' + os.path.basename(url)
								urllib.urlretrieve(url, local_zip_name)
								log.debug(url + ' retrieved to ' + local_zip_name)
	
								subsFound = False
								for file_ext in config['process.file_ext']:
									if not subsFound:
										log.debug('Entension is now %s' % file_ext)
										suitableList = getSuitableFilesListFromZip(local_zip_name, ep_attr, file_ext)
										log.info('%s suitable file(s) found before cleanup' % len(suitableList))

										cleanedList = cleanSuitableFilesList(suitableList, ep_attr)
										if cleanedList:
											saved = saveSubtitleFiles(cleanedList, path, ep_attr, url, backlog)
											subsFound = True
											if saved:
												ep_attr['subs'] = saved
												snatched.append(ep_attr)
										else:
											log.info('No Subs returned with extension %s' % file_ext)

								os.remove(local_zip_name)
								unzipped_root, _ = os.path.splitext(local_zip_name)
								if os.path.isdir(unzipped_root):
									shutil.rmtree(unzipped_root)

							else:
								log.debug('[ ] Episode: ' + file + " does not match season and episode number. Release: " + ep_attr['release'] + " SubRelease: " + ep_attr['subrelease'])
				else:
					log.info('Show dir ' + path + ' not found. Skipping')
			else:
				log.info(show + ' is not in my watch list')
		else:
			log.warn('Unable to parse filename: ' + item)

	return snatched

def backLogSearch():

	shows = listShows(config['shows.path'], True)
	log.debug(shows)

	for dist, local in shows.iteritems():
		url = config['rss.backlog'].replace('<show>', dist)
		try:
			log.info('---')
			log.info('--- %s' % local)
			log.info('---')
			savedSubs = process(url, True)

			if savedSubs:

				subj =  "%s: Backlog search for %s" % (PRG_NAME, local)
				notif_msg = 'During backlog search, %s has retrieved subs for : \n\n' % PRG_NAME

				for ep in savedSubs:
					if ep['number'] != ep['alt_number']:
						notif_msg += 'Episode S%dE%02d-%02d Release: %s Subrelease: %s\n' % (ep['season'], ep['number'], ep['alt_number'], ep['release'], ep['subrelease'])
					else:
						notif_msg += 'Episode S%dE%02d Release: %s Subrelease: %s\n' % (ep['season'], ep['number'], ep['release'], ep['subrelease'])
					for sub in ep['subs']:
						notif_msg += ' - %s\n' % ( os.path.basename(sub) )
					notif_msg += '\n'
				
				sendNotification(subj, notif_msg)

		except Exception, error:
			log.debug(error)
	
def main():
	try:
		opts, args = getopt.getopt(sys.argv[1:], "b", ['backlog'])
	except getopt.GetoptError:
		print 'Available options: --backlog'
		sys.exit()

	backlog = False

	for o, a in opts:
		if o in ('-b', '--backlog'):
			backlog = True

	if not os.path.isdir(TMP_DIR):
		os.mkdir(TMP_DIR)

	if not os.path.isdir(LOG_DIR):
		os.mkdir(LOG_DIR)

	log.info('%s version [%s] is starting...' % (PRG_NAME, PRG_VERSION) )
	log.info('---')

	log.debug(config)

	log.info('Back log mode is %s' % backlog)

	if backlog:
		backLogSearch()
	else:
		process(config['rss.url'], False)	

	log.info('---')
	log.info('%s has ended normally' % PRG_NAME)

if __name__ == '__main__':

	config = loadConfig(CFG_FILE)

	log = Logger(LOG_DIR + '/' + PRG_NAME + '_' + time.strftime('%Y%m%d') + '.log')

	main()


