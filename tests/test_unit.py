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


import io
import compoundfiles
import pytest
import mock


def test_reader_open_filename():
    with mock.patch('io.open') as mock_open:
        try:
            compoundfiles.CompoundFileReader('foo.doc')
        except ValueError:
            pass
        mock_open.assert_called_once_with('foo.doc', 'rb')

def test_reader_open_fileno():
    fileobj = mock.Mock()
    try:
        compoundfiles.CompoundFileReader(fileobj)
    except TypeError:
        pass
    fileobj.fileno.assert_called_with()

def test_reader_open_stream():
    with mock.patch('tempfile.SpooledTemporaryFile') as tempf, \
            mock.patch('shutil.copyfileobj') as copy:
        stream = io.BytesIO()
        try:
            compoundfiles.CompoundFileReader(stream)
        except ValueError:
            pass
        tempf.assert_called_once_with()
        copy.assert_called_once_with(stream, tempf.return_value)

def test_reader_open_invalid():
    with pytest.raises(IOError):
        compoundfiles.CompoundFileReader(object())

