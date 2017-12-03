#!/usr/bin/python3
#
# ProPresenter 6 to plain text converter
#
# Copyright 2017 Alex Henrie
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

import os
import sys
from base64 import b64decode
from subprocess import call, DEVNULL
from sys import argv
from tempfile import mkdtemp
from xml.etree import ElementTree

if len(argv) == 1:
    print('You must specify at least one pro6 file to convert.')
    exit()

temp_dir = mkdtemp()
temp_rtf = temp_dir + '/string.rtf'
temp_txt = temp_dir + '/string.txt'

for in_file in argv[1:]:
    out_file = os.path.splitext(in_file)[0] + '.txt'
    print(in_file + ' => ' + out_file)

    tree = ElementTree.parse(in_file)
    strings = []
    for el in tree.findall('.//NSString'):
        if el.attrib.get('rvXMLIvarName') != 'RTFData':
            continue
        open(temp_rtf, 'w').write(b64decode(el.text).decode('utf-8'))
        call(
            ['soffice', '--headless', '--convert-to', 'txt:Text', temp_rtf, '--outdir', temp_dir],
            stdout=DEVNULL
        )
        strings.append(open(temp_txt).read().replace('\uFEFF', '').strip())

    open(out_file, 'w').write('\n\n'.join(strings))

os.remove(temp_rtf)
os.remove(temp_txt)
os.rmdir(temp_dir)
