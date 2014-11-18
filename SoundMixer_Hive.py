# Sound Mixer Acquisition from Registry Hives
# Retrieves a list of applications with sound stored by
# Windows in the Sound Mixer preferences from a given set
# of registry files.
# Version 1.0
# @bbaskin - brian [at] thebaskins [dot] com
# Released 11 Nov 14

from __future__ import unicode_literals
import datetime
import os
import sys
from argparse import ArgumentParser
try:
    from Registry import Registry
except ImportError:
    print 'This script requires python-registry, obtained from: http://www.williballenthin.com/registry/index.html'
    quit()

    
def file_exists(fname):
    """
    Determine if a file exists
    Arguments:
        fname: path to a file
    Results:
        boolean value if file exists
    """
    return os.path.exists(fname) and os.access(fname, os.F_OK)

    
def GetOutputDevice(hive, guid):
    """
    Acquires the Windows audio output device name for a given GUID
    Arguments:
        guid: Windows GUID to query
    Results:
        Device name or original GUID if no results
    """
    if not hive:
        return guid
    key = '%s\\%s\\%s' % (r'Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render', guid, 'Properties')

    HKLM = Registry.Registry(hive)
    try:
        deviceKey = HKLM.open(key)
    except Registry.RegistryKeyNotFoundException:
        return guid

    try:
        deviceNameKey = deviceKey.value('{a45c254e-df1c-4efd-8020-67d146a850e0},2').value()
        return deviceNameKey
    except Registry.RegistryValueNotFoundException:
        return guid


def GetInputDevice(hive, guid):
    """
    Acquires the Windows audio input device name for a given GUID
    Arguments:
        guid: Windows GUID to query
    Results:
        Device name or original GUID if no results
    """
    if guid == '{00000000-0000-0000-0000-000000000000}':
        return 'N/A'
    if not hive:
        return guid
    key = '%s\\%s\\%s' % (r'Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture', guid, 'Properties')

    HKLM = Registry.Registry(hive)
    try:
        deviceKey = HKLM.open(key)
    except Registry.RegistryKeyNotFoundException:
        return guid

    try:
        deviceNameKey = deviceKey.value('{a45c254e-df1c-4efd-8020-67d146a850e0},2').value()
        return deviceNameKey
    except Registry.RegistryValueNotFoundException:
        return guid

        
def getArgs():
    """
    Parse the command line arguments
    Results:
        Dictionary of arguments
    """
    parser = ArgumentParser()
    parser.add_argument('-s', '--software', help='Registry SOFTWARE file location', required=False)
    parser.add_argument('-n', '--ntuser', help='Registry NTUSER.DAT file location', required=True)
    args = parser.parse_args()
    if (args.software) and (not file_exists(args.software)):
        print '[!] File not found: %s' % args.software
        quit()
    if not file_exists(args.ntuser):
        print '[!] File not found: %s' % args.ntuser
        quit()
    return args

    
def main():
    args = getArgs()
    registryBase = r'SOFTWARE\Microsoft\Internet Explorer\LowRegistry\Audio\PolicyConfig\PropertyStore'
    volumeSuffix = '{219ED5A0-9CBF-4F3A-B927-37C9E5C5F14F}'
    HKCU = Registry.Registry(args.ntuser)
    
    try:
        sndKey = HKCU.open(registryBase)
    except Registry.RegistryKeyNotFoundException:
        print '[!] Could not find Sound Mixer key. Exiting...'
        quit()
    
    for storeKey in sndKey.subkeys():
        try:
            sndAppValue = storeKey.value('').value() # Get (Default) value
            modDate = str(storeKey.timestamp()).replace(' ', ',').split('.')[0]
        except EnvironmentError, WindowsError:
            break 
        except Registry.RegistryValueNotFoundException:
            print '[!] Unable to open key. This is an unexpected error.'
            quit()

        volume = 0
        try:
            volumeKey = '%s\\%s\\%s' % (registryBase, storeKey.name(), volumeSuffix)
            volumeData = HKCU.open(volumeKey).value('3').value()
            volume = ord(volumeData[10])
        except Registry.RegistryKeyNotFoundException:
            volume = 'N/A'
            
        filename = unicode(sndAppValue).split('|')[1].split('%')[0]
        OutputGUID = sndAppValue.split('|')[0].split('}.')[1]
        OutputDevice = GetOutputDevice(args.software, OutputGUID)
        InputGUID = sndAppValue.split('%b')[1]
        InputDevice = GetInputDevice(args.software, InputGUID)
        output = u'%s,%s,%s,%s,%s' % (modDate, OutputDevice, volume, InputDevice, filename)
        print output.encode('utf8', 'replace')

    
if __name__ == '__main__':
    main()