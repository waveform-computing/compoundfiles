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
import compoundfiles as cf
import pytest
import warnings
from collections import namedtuple

DirEntry = namedtuple('DirEntry', ('name', 'isfile', 'size'))

def setup_module(module):
    warnings.simplefilter('always')

def verify_contents(doc, contents):
    # Routine for checking the contents of CompoundFileReader instance "doc"
    # against an iterable of DirEntry instances in "contents"
    for entry in contents:
        entity = doc.root
        for part in entry.name.split('/'):
            assert part in entity
            entity = entity[part]
        assert entity.isfile == entry.isfile
        assert not entity.isdir == entry.isfile
        if entry.isfile:
            assert entity.size == entry.size
        else:
            assert len(entity) == entry.size

def verify_example(doc, contents=None):
    # Checks that the specified "doc" matches the specification's example
    # document which contains a single storage containing a single stream. This
    # example (corrupted in various ways) is used numerous times in this test
    # suite
    if contents is None:
        contents = (
            DirEntry('Storage 1', False, 1),
            DirEntry('Storage 1/Stream 1', True, 544),
            )
    verify_contents(doc, contents)


def test_function_sample1_doc():
    with cf.CompoundFileReader('tests/sample1.doc') as doc:
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
    with cf.CompoundFileReader('tests/sample1.xls') as doc:
        contents = (
            DirEntry('Workbook', True, 11073),
            DirEntry('\x05SummaryInformation', True, 4096),
            DirEntry('\x05DocumentSummaryInformation', True, 4096),
            )
        verify_contents(doc, contents)

def test_function_sample2_doc():
    with cf.CompoundFileReader('tests/sample2.doc') as doc:
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
    with cf.CompoundFileReader('tests/sample2.xls') as doc:
        contents = (
            DirEntry('\x01Ole', True, 20),
            DirEntry('\x01CompObj', True, 73),
            DirEntry('Workbook', True, 1695),
            DirEntry('\x05SummaryInformation', True, 228),
            DirEntry('\x05DocumentSummaryInformation', True, 116),
            )
        verify_contents(doc, contents)

def test_entries_iter():
    with cf.CompoundFileReader('tests/example.dat') as doc:
        assert len([e for e in doc.root]) == 1

def test_entries_index():
    with cf.CompoundFileReader('tests/example.dat') as doc:
        assert doc.root[0] == doc.root['Storage 1']

def test_entries_bytes():
    with cf.CompoundFileReader('tests/example.dat') as doc:
        assert b'Storage 1' in doc.root
        assert doc.root['Storage 1'] is doc.root[b'Storage 1']

def test_entries_not_contains():
    with cf.CompoundFileReader('tests/example.dat') as doc:
        assert 'Storage 2' not in doc.root
        assert doc.root['Storage 1']['Stream 1'] not in doc.root

def test_spec_example():
    with cf.CompoundFileReader('tests/example.dat') as doc:
        verify_example(doc)

def test_invalid_name():
    # Same file as example.dat with corrupted names (no NULL terminator in
    # name1, incorrect length in name2); library continues as normal in these
    # cases
    with warnings.catch_warnings(record=True) as w:
        doc = cf.CompoundFileReader('tests/invalid_name1.dat')
        assert issubclass(w[0].category, cf.CompoundFileDirNameWarning)
        assert len(w) == 1
        verify_example(doc)
    with warnings.catch_warnings(record=True) as w:
        doc = cf.CompoundFileReader('tests/invalid_name2.dat')
        assert issubclass(w[0].category, cf.CompoundFileDirNameWarning)
        assert len(w) == 1
        verify_example(doc)

def test_invalid_root_type():
    with warnings.catch_warnings(record=True) as w:
        # Same file as example.dat with root dir-entry type corrupted; in this
        # case the library corrects the dir-entry type and continues
        doc = cf.CompoundFileReader('tests/invalid_root_type.dat')
        assert issubclass(w[0].category, cf.CompoundFileDirTypeWarning)
        assert len(w) == 1
        verify_example(doc)

def test_invalid_stream_type():
    with warnings.catch_warnings(record=True) as w:
        # Same file as example.dat with Stream 1 dir-entry type corrupted; in
        # this case the library ignores the Stream 1 entry
        doc = cf.CompoundFileReader('tests/invalid_stream_type.dat')
        assert issubclass(w[0].category, cf.CompoundFileDirTypeWarning)
        assert issubclass(w[1].category, cf.CompoundFileDirNameWarning)
        assert issubclass(w[2].category, cf.CompoundFileDirNameWarning)
        assert issubclass(w[3].category, cf.CompoundFileDirEntryWarning)
        assert issubclass(w[4].category, cf.CompoundFileDirSizeWarning)
        assert issubclass(w[5].category, cf.CompoundFileDirSizeWarning)
        assert len(w) == 6
        verify_example(doc, (
            DirEntry('Storage 1', False, 1),
            ))

def test_invalid_dir_indexes():
    with warnings.catch_warnings(record=True) as w:
        # Same file as example.dat with the root siblings corrupted, and
        # Stream 1 child corrupted; library continues as normal
        doc = cf.CompoundFileReader('tests/invalid_dir_indexes1.dat')
        assert issubclass(w[0].category, cf.CompoundFileDirIndexWarning)
        assert issubclass(w[1].category, cf.CompoundFileDirIndexWarning)
        assert issubclass(w[2].category, cf.CompoundFileDirIndexWarning)
        assert len(w) == 3
        verify_example(doc)
    with warnings.catch_warnings(record=True) as w:
        # Same file as example.dat with the Stream 1 siblings corrupted;
        # library continues as normal
        doc = cf.CompoundFileReader('tests/invalid_dir_indexes2.dat')
        assert issubclass(w[0].category, cf.CompoundFileDirIndexWarning)
        assert issubclass(w[1].category, cf.CompoundFileDirIndexWarning)
        assert len(w) == 2
        verify_example(doc)
    with warnings.catch_warnings(record=True) as w:
        # Same file as example.dat with the root child index corrupted; library
        # continues but file is effectively empty
        doc = cf.CompoundFileReader('tests/invalid_dir_indexes3.dat')
        assert issubclass(w[0].category, cf.CompoundFileDirIndexWarning)
        assert len(w) == 1
        verify_example(doc, ())

def test_invalid_dir_misc():
    with warnings.catch_warnings(record=True) as w:
        # Same file as example.dat with UUID and timestamps corrupted in
        # Stream 1; library continues as normal
        doc = cf.CompoundFileReader('tests/invalid_dir_misc.dat')
        assert issubclass(w[0].category, cf.CompoundFileDirEntryWarning)
        assert issubclass(w[1].category, cf.CompoundFileDirTimeWarning)
        assert issubclass(w[2].category, cf.CompoundFileDirTimeWarning)
        assert len(w) == 3
        verify_example(doc)

def test_invalid_dir_size():
    with warnings.catch_warnings(record=True) as w:
        # Same file as example.dat with size of Stream 1 corrupted (> 32-bits);
        # library re-writes size to be within 32-bits
        doc = cf.CompoundFileReader('tests/invalid_dir_size1.dat')
        assert issubclass(w[0].category, cf.CompoundFileDirSizeWarning)
        assert issubclass(w[1].category, cf.CompoundFileDirSizeWarning)
        assert len(w) == 2
        verify_example(doc, (
            DirEntry('Storage 1', False, 1),
            DirEntry('Storage 1/Stream 1', True, 0xFFFFFFFF),
            ))

def test_invalid_dir_loop():
    with pytest.raises(cf.CompoundFileDirLoopError):
        # Same as example.dat but with Stream 1's left pointer corrupted to
        # point to the Root Entry
        doc = cf.CompoundFileReader('tests/invalid_dir_loop.dat')

def test_invalid_fat_loop():
    with pytest.raises(cf.CompoundFileNormalLoopError):
        # Sample as example.dat but with Stream 1's FAT entry corrupted to
        # point to itself
        doc = cf.CompoundFileReader('tests/invalid_fat_loop.dat')

def test_stream_attr():
    with cf.CompoundFileReader('tests/example.dat') as doc:
        with doc.open('Storage 1/Stream 1') as f:
            assert f.readable()
            assert not f.writable()
            assert f.seekable()

def test_stream_seek():
    with cf.CompoundFileReader('tests/example.dat') as doc:
        with doc.open('Storage 1/Stream 1') as f:
            assert f.seek(0, io.SEEK_END) == 544
            assert f.seek(0, io.SEEK_CUR) == 544
            with pytest.raises(ValueError):
                f.seek(-1)

def test_stream_read():
    with cf.CompoundFileReader('tests/example2.dat') as doc:
        # Same file as example.dat with an additional Stream 2 which is 4112
        # bytes long (too long for mini FAT)
        with doc.open('Storage 1/Stream 1') as f:
            assert len(f.read()) == 544
            f.seek(0)
            assert len(f.read(1024)) == 544
            f.seek(0)
            assert len(f.read1()) == 64
            f.seek(0, io.SEEK_END)
            assert f.read1() == b''
        with doc.open('Storage 1/Stream 2') as f:
            assert len(f.read()) == 4112
            f.seek(0)
            assert len(f.read1()) == 512
            f.seek(0, io.SEEK_END)
            assert f.read1() == b''

def test_stream_read_broken_size():
    with cf.CompoundFileReader('tests/invalid_dir_size2.dat') as doc:
        # Same file as example.dat with size of Stream 1 corrupted to 3072
        # bytes (small enough to fit in the mini FAT but too large for the
        # actual data which is 544 bytes), and additional Stream 2 which has
        # corrupted size 8192 bytes (actual size 512 bytes)
        with doc.open('Storage 1/Stream 1') as f:
            assert f.seek(0, io.SEEK_END) == 576
        with doc.open('Storage 1/Stream 2') as f:
            assert f.seek(0, io.SEEK_END) == 512
