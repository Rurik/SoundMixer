# Sound Mixer Acquisition from Live Registry
# Retrieves a list of applications with sound stored by
# Windows in the Sound Mixer preferences from a live system.
# Version 1.0
# @bbaskin - brian [at] thebaskins [dot] com
# Released 11 Nov 14

from __future__ import unicode_literals
import datetime
import sys
try:
    import _winreg
except ImportError:
    print 'This script must be run from within a Windows environment.'
    quit()

    
def ConvertDate(ns):
    """
    Converts a given number of nanoseconds since 1 Jan 1601 to a standard date.
    Arguments:
        ns: number of nanoseconds
    Results:
        string based on a datetime structure
    """
    startDate = datetime.datetime(1601, 1, 1)
    dateStruct = '%Y-%m-%d,%H:%M:%S'
    dateModified = startDate + datetime.timedelta(seconds = ns*(10**-9)*100)
    return dateModified.strftime(dateStruct)


def GetOutputDevice(guid):
    """
    Acquires the Windows audio output device name for a given GUID
    Arguments:
        guid: Windows GUID to query
    Results:
        Device name or original GUID if no results
    """
    key = '%s\\%s\\%s' % (r'SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render', guid, 'Properties')

    try:
        deviceKey = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, key, 0, _winreg.KEY_READ | _winreg.KEY_WOW64_64KEY)
    except WindowsError:
        return guid
    deviceNameKey = _winreg.QueryValueEx(deviceKey, '{a45c254e-df1c-4efd-8020-67d146a850e0},2')[0]
    if deviceNameKey:
        return deviceNameKey
    else:
        return guid


def GetInputDevice(guid):
    """
    Acquires the Windows audio input device name for a given GUID
    Arguments:
        guid: Windows GUID to query
    Results:
        Device name or original GUID if no results
    """
    if guid == '{00000000-0000-0000-0000-000000000000}':
        return 'N/A'

    key = '%s\\%s\\%s' % (r'SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture', guid, 'Properties')
    try:
        deviceKey = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, key, 0, _winreg.KEY_READ | _winreg.KEY_WOW64_64KEY)
    except WindowsError:
        return guid
    deviceNameKey = _winreg.QueryValueEx(deviceKey, '{a45c254e-df1c-4efd-8020-67d146a850e0},2')[0]
    if deviceNameKey:
        return deviceNameKey
    else:
        return guid


def main():
    registryBase = r'SOFTWARE\Microsoft\Internet Explorer\LowRegistry\Audio\PolicyConfig\PropertyStore'
    volumeSuffix = r'{219ED5A0-9CBF-4F3A-B927-37C9E5C5F14F}'
    sndKey = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER, registryBase)
    sndKeyMeta = _winreg.QueryInfoKey(sndKey)
    for i in range(0, sndKeyMeta[0]):
        try:
            storeKey = _winreg.EnumKey(sndKey,i)
            sndAppKey = _winreg.OpenKey(sndKey,storeKey)
            sndAppMeta = _winreg.QueryInfoKey(sndAppKey)
            sndAppValue = _winreg.QueryValueEx(sndAppKey, "")
        except EnvironmentError, WindowsError:
            break 

        try:
            volumeRegistry = registryBase + '\\' + storeKey + '\\' + volumeSuffix
            volumeKey = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER, volumeRegistry)
            volumeData = _winreg.QueryValueEx(volumeKey, '3')
            volume = ord(volumeData[0][10])
        except EnvironmentError, WindowsError:
            volume = 'N/A'
            
        filename = unicode(sndAppValue[0]).split('|')[1].split('%')[0]
        modDate = ConvertDate(int(sndAppMeta[2]))
        OutputGUID = sndAppValue[0].split('|')[0].split('}.')[1]
        OutputDevice = GetOutputDevice(OutputGUID)
        InputGUID = sndAppValue[0].split("%b")[1]
        InputDevice = GetInputDevice(InputGUID)
        output = u'%s,%s,%s,%s,%s' % (modDate, OutputDevice, volume, InputDevice, filename)
        print output.encode('utf8', 'replace')

    _winreg.CloseKey(sndKey)
    
    
if __name__ == '__main__':
    main()