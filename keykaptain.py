#!/usr/bin/python

license = """
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""

import os
import binascii
import re
import webbrowser

regex = re.compile("^\\\\(.+)$")
hosts = list()

def get_hosts():
        raw_output = os.popen("net view").read()
        host_list = raw_output.split()
        for host in host_list:
                match = regex.match(host)
                if match:
                        hosts.append(match.group(1)[1:])


if __name__ == "__main__":
        get_hosts()
        command = "ProductKeyGrabber.exe"
        for host in hosts:
                command += " " + host
        print command
        os.system(command)
        webbrowser.open("webpage.html")
        """
        file = open("windowskeys.txt", "r+")
        text = file.read()
        text = list(text).sort()
        for line in text:
                file.writelines(line)
        file.close()
        file = open("officekeys.txt", "r+")
        text = file.read()
        text = list(text).sort()
        for line in text:
                file.writelines(line)
        file.close()
        """
