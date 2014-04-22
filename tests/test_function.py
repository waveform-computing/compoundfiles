#!/usr/bin/env python
# vim: set et sw=4 sts=4 fileencoding=utf-8:
#
# A library for reading Microsoft's OLE Compound Document format
# Copyright (c) 2014 Dave Hughes <dave@waveform.org.uk>
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

from __future__ import (
    unicode_literals,
    absolute_import,
    print_function,
    division,
    )
str = type('')


import compoundfiles
import pytest
import warnings
from collections import namedtuple

DirEntry = namedtuple('DirEntry', ('name', 'isfile', 'size'))

def setup_module(module):
    warnings.simplefilter('always')

def verify_contents(doc, contents):
    for entry in contents:
        assert entry.name in doc.root
        assert doc.root[entry.name].isfile == entry.isfile
        assert not doc.root[entry.name].isdir == entry.isfile
        if entry.isfile:
            assert doc.root[entry.name].size == entry.size
        else:
            assert len(doc.root[entry.name]) == entry.size


def test_function_sample1_doc():
    with compoundfiles.CompoundFileReader('tests/sample1.doc') as doc:
        contents = (
            DirEntry('1Table', True, 8375),
            DirEntry('\x01CompObj', True, 106),
            DirEntry('ObjectPool', False, 0),
            DirEntry('WordDocument', True, 9280),
            DirEntry('\x05SummaryInformation', True, 4096),
            DirEntry('\x05DocumentSummaryInformation', True, 4096),
            )
        verify_contents(doc, contents)

def test_function_sample1_xls():
    with compoundfiles.CompoundFileReader('tests/sample1.xls') as doc:
        contents = (
            DirEntry('Workbook', True, 11073),
            DirEntry('\x05SummaryInformation', True, 4096),
            DirEntry('\x05DocumentSummaryInformation', True, 4096),
            )
        verify_contents(doc, contents)

def test_function_sample2_doc():
    with compoundfiles.CompoundFileReader('tests/sample2.doc') as doc:
        contents = (
            DirEntry('Data', True, 8420),
            DirEntry('1Table', True, 19168),
            DirEntry('\x01CompObj', True, 113),
            DirEntry('WordDocument', True, 25657),
            DirEntry('\x05SummaryInformation', True, 4096),
            DirEntry('\x05DocumentSummaryInformation', True, 4096),
            )
        verify_contents(doc, contents)

def test_function_sample2_xls():
    with compoundfiles.CompoundFileReader('tests/sample2.xls') as doc:
        contents = (
            DirEntry('\x01Ole', True, 20),
            DirEntry('\x01CompObj', True, 73),
            DirEntry('Workbook', True, 1695),
            DirEntry('\x05SummaryInformation', True, 228),
            DirEntry('\x05DocumentSummaryInformation', True, 116),
            )
        verify_contents(doc, contents)

