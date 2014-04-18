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
import warnings
import compoundfiles
import pytest
from mock import MagicMock, patch, call


V3_HEADER = (
    compoundfiles.const.COMPOUND_MAGIC, # magic
    b'\0' * 16,                         # uuid
    0,                                  # minor ver
    3,                                  # major ver
    0xFFFE,                             # BOM
    9,                                  # 512 byte sectors
    6,                                  # 64 byte mini-sectors
    b'\0' * 6,                          # unused
    0,                                  # dir sector count
    0,                                  # normal sector count
    0,                                  # dir first sector
    0,                                  # txn signature
    0,                                  # mini FAT size limit
    0,                                  # mini first sector
    0,                                  # mini sector count
    compoundfiles.const.END_OF_CHAIN,   # master first sector
    0,                                  # master sector count
    )

V4_HEADER = (
    compoundfiles.const.COMPOUND_MAGIC, # magic
    b'\0' * 16,                         # uuid
    0,                                  # minor ver
    4,                                  # major ver
    0xFFFE,                             # BOM
    12,                                 # 4096 byte sectors
    6,                                  # 64 byte mini-sectors
    b'\0' * 6,                          # unused
    1,                                  # dir sector count
    0,                                  # normal sector count
    0,                                  # dir first sector
    0,                                  # txn signature
    0,                                  # mini FAT size limit
    0,                                  # mini first sector
    0,                                  # mini sector count
    compoundfiles.const.END_OF_CHAIN,   # master first sector
    0,                                  # master sector count
    )

def test_reader_open_filename():
    with patch('io.open') as m:
        try:
            compoundfiles.CompoundFileReader('foo.doc')
        except ValueError:
            pass
        assert m.mock_calls[:2] == [call('foo.doc', 'rb'), call().fileno()]

def test_reader_open_stream():
    with patch('tempfile.SpooledTemporaryFile') as m_temp, \
            patch('shutil.copyfileobj') as m_copy:
        stream = io.BytesIO()
        try:
            compoundfiles.CompoundFileReader(stream)
        except ValueError:
            pass
        assert m_temp.mock_calls[:2] == [call(), call().fileno()]
        assert m_copy.mock_calls == [call(stream, m_temp.return_value)]

def test_reader_open_invalid():
    with pytest.raises(IOError):
        compoundfiles.CompoundFileReader(object())

def test_reader_header_magic():
    with patch('mmap.mmap') as mmap, \
            patch('compoundfiles.reader.COMPOUND_HEADER') as header:
        header.unpack.return_value = list(V3_HEADER)
        header.unpack.return_value[0] = 0
        with pytest.raises(compoundfiles.CompoundFileInvalidMagicError):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_header_bom():
    with patch('mmap.mmap') as mmap, \
            patch('compoundfiles.reader.COMPOUND_HEADER') as header:
        header.unpack.return_value = list(V3_HEADER)
        header.unpack.return_value[4] = 0xFEFF
        with pytest.raises(compoundfiles.CompoundFileInvalidBOMError):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_silly_normal_sector_size():
    with patch('mmap.mmap') as mmap, \
            patch('compoundfiles.reader.COMPOUND_HEADER') as header:
        header.unpack.return_value = list(V3_HEADER)
        header.unpack.return_value[5] = 21
        with warnings.catch_warnings(record=True) as w:
            try:
                compoundfiles.CompoundFileReader(MagicMock())
            except:
                pass
            assert len(w) > 0
            assert issubclass(w[0].category, compoundfiles.CompoundFileSectorSizeWarning)

def test_reader_silly_mini_sector_size():
    with patch('mmap.mmap') as mmap, \
            patch('compoundfiles.reader.COMPOUND_HEADER') as header:
        header.unpack.return_value = [compoundfiles.const.COMPOUND_MAGIC, 0, 0, 0, 0xFFFE, 9, 9] + [0] * 10
        with warnings.catch_warnings(record=True) as w:
            try:
                compoundfiles.CompoundFileReader(MagicMock())
            except IOError:
                pass
            assert len(w) > 0
            assert issubclass(w[0].category, compoundfiles.CompoundFileSectorSizeWarning)

def test_reader_strange_v3_settings():
    with patch('mmap.mmap') as mmap, \
            patch('compoundfiles.reader.COMPOUND_HEADER') as header:
        header.unpack.return_value = [compoundfiles.const.COMPOUND_MAGIC, 0, 0, 3, 0xFFFE, 10, 9, 0, 1] + [0] * 8
        with warnings.catch_warnings(record=True) as w:
            try:
                compoundfiles.CompoundFileReader(MagicMock())
            except:
                pass
            assert len(w) > 2
            # XXX Order of warnings shouldn't matter...
            assert issubclass(w[0].category, compoundfiles.CompoundFileSectorSizeWarning)
            assert issubclass(w[1].category, compoundfiles.CompoundFileHeaderWarning)
            assert issubclass(w[2].category, compoundfiles.CompoundFileSectorSizeWarning)

def test_reader_strange_v4_settings():
    with patch('mmap.mmap') as mmap, \
            patch('compoundfiles.reader.COMPOUND_HEADER') as header:
        header.unpack.return_value = [compoundfiles.const.COMPOUND_MAGIC, 0, 0, 4, 0xFFFE, 10, 9, 0, 1] + [0] * 8
        with warnings.catch_warnings(record=True) as w:
            try:
                compoundfiles.CompoundFileReader(MagicMock())
            except:
                pass
            assert len(w) > 0
            assert issubclass(w[0].category, compoundfiles.CompoundFileSectorSizeWarning)

def test_reader_invalid_dll_version():
    with patch('mmap.mmap') as mmap, \
            patch('compoundfiles.reader.COMPOUND_HEADER') as header:
        header.unpack.return_value = [compoundfiles.const.COMPOUND_MAGIC, 0, 0, 5, 0xFFFE, 10, 9, 0, 1] + [0] * 8
        with pytest.raises(compoundfiles.CompoundFileVersionError):
            compoundfiles.CompoundFileReader(MagicMock())
