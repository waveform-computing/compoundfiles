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
native_str = str
str = type('')


import io
import warnings
import compoundfiles
import pytest
import struct
from mock import MagicMock, patch, call

warnings.simplefilter('error')

V3_HEADER = (
    compoundfiles.const.COMPOUND_MAGIC, # 0  magic
    b'\0' * 16,                         # 1  uuid
    0,                                  # 2  minor ver
    3,                                  # 3  major ver
    0xFFFE,                             # 4  BOM
    9,                                  # 5  512 byte sectors
    6,                                  # 6  64 byte mini-sectors
    b'\0' * 6,                          # 7  unused
    0,                                  # 8  dir sector count
    1,                                  # 9  normal sector count
    0,                                  # 10 dir first sector
    0,                                  # 11 txn signature
    0,                                  # 12 mini FAT size limit
    0,                                  # 13 mini first sector
    0,                                  # 14 mini sector count
    compoundfiles.const.END_OF_CHAIN,   # 15 master first sector
    0,                                  # 16 master sector count
    )

V4_HEADER = (
    compoundfiles.const.COMPOUND_MAGIC, # 0  magic
    b'\0' * 16,                         # 1  uuid
    0,                                  # 2  minor ver
    4,                                  # 3  major ver
    0xFFFE,                             # 4  BOM
    12,                                 # 5  4096 byte sectors
    6,                                  # 6  64 byte mini-sectors
    b'\0' * 6,                          # 7  unused
    1,                                  # 8  dir sector count
    1,                                  # 9  normal sector count
    0,                                  # 10 dir first sector
    0,                                  # 11 txn signature
    0,                                  # 12 mini atT size limit
    0,                                  # 13 mini first sector
    0,                                  # 14 mini sector count
    compoundfiles.const.END_OF_CHAIN,   # 15 master first sector
    0,                                  # 16 master sector count
    )


class mmap_mock(bytes):
    def size(self):
        return len(self)


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
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        header[0] = b'\0'
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header)
            )
        with pytest.raises(compoundfiles.CompoundFileInvalidMagicError):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_header_bom():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        header[4] = 0xFEFF
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header)
            )
        with pytest.raises(compoundfiles.CompoundFileInvalidBomError):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_silly_normal_sector_size():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        header[5] = 21
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header)
            )
        with pytest.raises(compoundfiles.CompoundFileSectorSizeWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_silly_mini_sector_size():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        header[6] = 9
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header)
            )
        with pytest.raises(compoundfiles.CompoundFileSectorSizeWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_strange_v3_settings():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        header[5] = 10
        header[6] = 9
        header[8] = 1
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header)
            )
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter('always')
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
    with patch('mmap.mmap') as mmap:
        header = list(V4_HEADER)
        header[5] = 10
        header[6] = 9
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header)
            )
        with pytest.raises(compoundfiles.CompoundFileSectorSizeWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_invalid_dll_version():
    with patch('mmap.mmap') as mmap:
        header = list(V4_HEADER)
        header[3] = 5
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header)
            )
        with pytest.raises(compoundfiles.CompoundFileVersionError):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_invalid_clsid():
    with patch('mmap.mmap') as mmap:
        header = list(V4_HEADER)
        header[1] = b'\1'
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header)
            )
        with pytest.raises(compoundfiles.CompoundFileHeaderWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_invalid_txn_signature():
    with patch('mmap.mmap') as mmap:
        header = list(V4_HEADER)
        header[11] = 1
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header)
            )
        with pytest.raises(compoundfiles.CompoundFileHeaderWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_non_empty_unused():
    with patch('mmap.mmap') as mmap:
        header = list(V4_HEADER)
        header[7] = b'\1'
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header)
            )
        with pytest.raises(compoundfiles.CompoundFileHeaderWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_master_extension_pointer_free():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        header[15] = compoundfiles.const.FREE_SECTOR
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header) +
            struct.pack('<109L', *([0] * 109))
            )
        with pytest.raises(compoundfiles.CompoundFileMasterFatWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_master_extension_pointer_invalid():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        header[15] = 1
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header) +
            struct.pack('<109L', *([0] * 109))
            )
        with pytest.raises(compoundfiles.CompoundFileMasterFatWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_master_extension_chained():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        header[16] = 1
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header) +
            struct.pack('<109L', *([0] * 109))
            )
        with pytest.raises(compoundfiles.CompoundFileMasterFatWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_master_terminated_with_free():
    with patch('mmap.mmap') as mmap:
        master_fat = [0, compoundfiles.const.FREE_SECTOR] + [0] * 107
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*V3_HEADER) +
            struct.pack('<109L', *master_fat)
            )
        with pytest.raises(compoundfiles.CompoundFileMasterFatWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_master_sector_beyond_eof():
    with patch('mmap.mmap') as mmap:
        master_fat = [20] + [0] * 108
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*V3_HEADER) +
            struct.pack('<109L', *master_fat)
            )
        with pytest.raises(compoundfiles.CompoundFileMasterFatWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_master_invalid_special():
    with patch('mmap.mmap') as mmap:
        master_fat = [0, compoundfiles.const.MASTER_FAT_SECTOR] + [0] * 107
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*V3_HEADER) +
            struct.pack('<109L', *master_fat)
            )
        with pytest.raises(compoundfiles.CompoundFileMasterFatWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_master_loop():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        header[15] = 0
        header[16] = 1
        master_fat = list(range(1, 1+109)) + list(range(109, 109+127)) + [0]
        header[9] = len(master_fat) - 1
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header) +
            struct.pack('<109L128L', *master_fat) +
            struct.pack('<%dx' % ((109+128)*512))
            )
        with pytest.raises(compoundfiles.CompoundFileMasterLoopError):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_master_too_large():
    with patch('mmap.mmap') as mmap:
        header = list(V4_HEADER)
        header[15] = 0
        header[16] = 1
        master_fat = [0] * (109 + (30000 * 1024))
        for next_sector, index in enumerate(range(109 + 1024 - 1, len(master_fat), 1024)):
            master_fat[index] = next_sector + 1
        #master_fat = [0] * 109
        #for next_sector in range(30000):
        #    master_fat.extend([0] * 1023)
        #    master_fat.append(next_sector + 1)
        header[9] = len(master_fat) - 30000
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header) +
            struct.pack('<109L3584x%dL' % (1024 * 30000), *master_fat)
            )
        with pytest.raises(compoundfiles.CompoundFileLargeNormalFatError):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_master_ends_early():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        header[15] = 0
        header[16] = 10
        master_fat = [0] * 109
        for next_sector in range(4):
            master_fat.extend([0] * 127)
            master_fat.append(next_sector + 1)
        master_fat[-1] = compoundfiles.const.END_OF_CHAIN
        header[9] = len(master_fat) - 4
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header) +
            struct.pack('<109L', *master_fat[:109]) +
            struct.pack('<%dL' % (128*4), *master_fat[109:])
            )
        with pytest.raises(compoundfiles.CompoundFileMasterFatWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_master_ends_late():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        header[15] = 0
        header[16] = 1
        master_fat = [0] * 109
        for next_sector in range(4):
            master_fat.extend([0] * 127)
            master_fat.append(next_sector + 1)
        master_fat[-1] = compoundfiles.const.END_OF_CHAIN
        header[9] = len(master_fat) - 4
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header) +
            struct.pack('<109L', *master_fat[:109]) +
            struct.pack('<%dL' % (128*4), *master_fat[109:])
            )
        with pytest.raises(compoundfiles.CompoundFileMasterFatWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_normal_count_wrong():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        header[15] = 0
        header[16] = 4
        master_fat = [0] * 109
        for next_sector in range(4):
            master_fat.extend([0] * 127)
            master_fat.append(next_sector + 1)
        master_fat[-1] = compoundfiles.const.END_OF_CHAIN
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header) +
            struct.pack('<109L', *master_fat[:109]) +
            struct.pack('<%dL' % (128*4), *master_fat[109:])
            )
        with pytest.raises(compoundfiles.CompoundFileMasterFatWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_master_sector_marked_wrong():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        header[15] = 0
        header[16] = 1
        master_fat = ([1] * 109) + ([1] * 127) + [compoundfiles.const.END_OF_CHAIN]
        header[9] = len(master_fat) - 1
        normal_fat = [0] * 128
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header) +
            struct.pack('<109L', *master_fat[:109]) +
            struct.pack('<128L', *master_fat[109:]) +
            struct.pack('<128L', *normal_fat)
            )
        with pytest.raises(compoundfiles.CompoundFileMasterSectorWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_normal_sector_marked_wrong():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        header[15] = 0
        header[16] = 1
        master_fat = ([1] * 109) + ([1] * 127) + [compoundfiles.const.END_OF_CHAIN]
        header[9] = len(master_fat) - 1
        normal_fat = [compoundfiles.const.MASTER_FAT_SECTOR] + [0] * 127
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header) +
            struct.pack('<109L', *master_fat[:109]) +
            struct.pack('<128L', *master_fat[109:]) +
            struct.pack('<128L', *normal_fat)
            )
        with pytest.raises(compoundfiles.CompoundFileNormalSectorWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_mini_too_large():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        master_fat = [0, compoundfiles.const.END_OF_CHAIN] + [0] * 107
        normal_fat = [compoundfiles.const.NORMAL_FAT_SECTOR] + [0] * 127
        header[14] = 300000
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header) +
            struct.pack('<109L', *master_fat) +
            struct.pack('<128L', *normal_fat)
            )
        with pytest.raises(compoundfiles.CompoundFileLargeMiniFatError):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_mini_free_sector():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        master_fat = [0, compoundfiles.const.END_OF_CHAIN] + [0] * 107
        normal_fat = [compoundfiles.const.NORMAL_FAT_SECTOR] + [0] * 127
        header[13] = compoundfiles.const.FREE_SECTOR
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header) +
            struct.pack('<109L', *master_fat) +
            struct.pack('<128L', *normal_fat)
            )
        with pytest.raises(compoundfiles.CompoundFileMiniFatWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_mini_pointer_free():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        master_fat = [0, compoundfiles.const.END_OF_CHAIN] + [0] * 107
        normal_fat = [compoundfiles.const.NORMAL_FAT_SECTOR] + [0] * 127
        header[13] = compoundfiles.const.FREE_SECTOR
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header) +
            struct.pack('<109L', *master_fat) +
            struct.pack('<128L', *normal_fat)
            )
        with pytest.raises(compoundfiles.CompoundFileMiniFatWarning):
            compoundfiles.CompoundFileReader(MagicMock())

def test_reader_mini_pointer_invalid():
    with patch('mmap.mmap') as mmap:
        header = list(V3_HEADER)
        master_fat = [0, compoundfiles.const.END_OF_CHAIN] + [0] * 107
        normal_fat = [compoundfiles.const.NORMAL_FAT_SECTOR] + [0] * 127
        header[13] = 10
        mmap.return_value = mmap_mock(
            compoundfiles.const.COMPOUND_HEADER.pack(*header) +
            struct.pack('<109L', *master_fat) +
            struct.pack('<128L', *normal_fat)
            )
        with pytest.raises(compoundfiles.CompoundFileMiniFatWarning):
            compoundfiles.CompoundFileReader(MagicMock())

