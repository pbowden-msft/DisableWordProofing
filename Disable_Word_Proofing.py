#!/usr/bin/python
#
# Microsoft Word - Disable Proofing Tools
# Script Version 1.0
#
# Copyright (c) 2019 Microsoft Corp. All rights reserved.
# Scripts are not supported under any Microsoft standard support program or service. The scripts are provided AS IS
# without warranty of any kind. Microsoft disclaims all implied warranties including, without limitation, any implied
# warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or
# performance of the scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone
# else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever
# (including, without limitation, damages for loss of business profits, business interruption, loss of business
# information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or
# documentation, even if Microsoft has been advised of the possibility of such damages.
# Feedback: pbowden@microsoft.com

import grp
import os
import pwd
import shutil
import subprocess
import sys
import SystemConfiguration

### Constants

word_path = '/Applications/Microsoft Word.app'
word_proofing_path = os.path.join(word_path, 'Contents/SharedSupport/Proofing Tools')
acl_filename = 'Microsoft Office ACL [English]'

### Functions

# Function to read the bundle version of an app
def get_app_version(bundle):
    path = os.path.join(bundle, 'Contents/Info.plist')
    try:
        result = subprocess.check_output(['defaults', 'read', path, 'CFBundleVersion']).strip()
        return result
    except subprocess.CalledProcessError:
        return None

# Function to get the current logged-on username
def get_current_console_username():
    username = (SystemConfiguration.SCDynamicStoreCopyConsoleUser(None, None, None) or [None])[0]
    username = [username, ''][username in [u'loginwindow', None, u'']]
    return username

# Function to return the full path of the Office Group Container
def get_office_home_folder(username):
	from os.path import expanduser
	home_office = expanduser('~' + str(username)) + str('/Library/Group Containers/UBF8T346G9.Office')
	if os.path.isdir(home_office):
		return home_office
	return None

# Function to create the Office Group Container
def create_office_home_folder(username):
	from os.path import expanduser
	os.makedirs(expanduser('~' + str(username)) + str('/Library/Group Containers/UBF8T346G9.Office'), 0700)
	home_office = expanduser('~' + str(username)) + str('/Library/Group Containers/UBF8T346G9.Office')
	uid = pwd.getpwnam(username).pw_uid
	gid = grp.getgrnam('staff').gr_gid
	os.chown(home_office, uid, gid)
	print 'Created group container for user ' + str(username)
	return None

# Function to write an empty auto-correct list to a given path
def write_empty_acl(path, username):
	empty_acl = b'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'
	with open(os.path.join(path, acl_filename), "wb") as fh:
		fh.write(empty_acl.decode('base64'))
		print 'Created empty auto-correct list at ' + str(path)
		uid = pwd.getpwnam(username).pw_uid
		gid = grp.getgrnam('staff').gr_gid
		filepath = str(os.path.join(path, acl_filename))
		os.chown(filepath, uid, gid)

### MAIN

# Test if Word is installed
word_installed = os.path.isdir(word_path)
if word_installed:
	print 'Found Microsoft Word ' + str(get_app_version(word_path))
else:
	print 'Microsoft Word is not installed'
	sys.exit(1)

# Remove the Proofing Tools folder from the Word app bundle
proofing_tools_present = os.path.isdir(word_proofing_path)
if proofing_tools_present:
	if os.geteuid() == 0:
		print 'Removing Speller and Grammar checker'
		shutil.rmtree(word_proofing_path, ignore_errors=True)
	else:
		print 'This script must be run with elevated permissions'
		sys.exit(1)
else:
	print 'Proofing tools are already disabled'

# Create empty auto-correct list
local_username = get_current_console_username()
local_userpath = get_office_home_folder(local_username)
if local_userpath:
	write_empty_acl(local_userpath, local_username)
else:
	create_office_home_folder(local_username)
	local_userpath = get_office_home_folder(local_username)
	write_empty_acl(local_userpath, local_username)


print ''
sys.exit(0)