#!/usr/bin/env python
# vim: set et sw=4 sts=4 fileencoding=utf-8:
#
# A library for reading Microsoft's OLE Compound Document format
# Copyright (c) 2014 Dave Jones <dave@waveform.org.uk>
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

def setup_function(fn):
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

def verify_example(doc, contents=None, test_contents=True):
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
    if test_contents:
        for entry in contents:
            if entry.isfile:
                assert doc.open(entry.name).read() == (b'Data' * ((entry.size + 3) // 4))[:entry.size]

@pytest.fixture(params=(
    ('tests/sample1.doc', (
        DirEntry('1Table', True, 8375),
        DirEntry('\x01CompObj', True, 106),
        DirEntry('ObjectPool', False, 0),
        DirEntry('WordDocument', True, 9280),
        DirEntry('\x05SummaryInformation', True, 4096),
        DirEntry('\x05DocumentSummaryInformation', True, 4096),
        )),
    ('tests/sample1.xls', (
        DirEntry('Workbook', True, 11073),
        DirEntry('\x05SummaryInformation', True, 4096),
        DirEntry('\x05DocumentSummaryInformation', True, 4096),
        )),
    ('tests/sample2.doc', (
        DirEntry('Data', True, 8420),
        DirEntry('1Table', True, 19168),
        DirEntry('\x01CompObj', True, 113),
        DirEntry('WordDocument', True, 25657),
        DirEntry('\x05SummaryInformation', True, 4096),
        DirEntry('\x05DocumentSummaryInformation', True, 4096),
        )),
    ('tests/sample2.xls', (
        DirEntry('\x01Ole', True, 20),
        DirEntry('\x01CompObj', True, 73),
        DirEntry('Workbook', True, 1695),
        DirEntry('\x05SummaryInformation', True, 228),
        DirEntry('\x05DocumentSummaryInformation', True, 116),
        )),
    ('tests/example.dat', (
        DirEntry('Storage 1', False, 1),
        DirEntry('Storage 1/Stream 1', True, 544),
        )),
    ))
def sample(request):
    return request.param


def test_sample_from_filename(sample):
    filename, contents = sample
    with cf.CompoundFileReader(filename) as doc:
        verify_contents(doc, contents)

def test_sample_from_file(sample):
    filename, contents = sample
    f = io.open(filename, 'rb')
    with cf.CompoundFileReader(f) as doc:
        verify_contents(doc, contents)

def test_sample_from_fake_mmap(sample):
    filename, contents = sample
    stream = io.BytesIO()
    with io.open(filename, 'rb') as source:
        stream.write(source.read())
    with cf.CompoundFileReader(stream) as doc:
        verify_contents(doc, contents)

def test_reader_invalid_source():
    with pytest.raises(TypeError):
        cf.CompoundFileReader(object())

def test_reader_bad_sector():
    with cf.CompoundFileReader('tests/example.dat') as doc:
        with pytest.raises(cf.CompoundFileError):
            # Can't do this with a corrupted file (or at least I haven't
            # figured out how yet!); too many safeguards!
            doc._read_sector(6)

def test_entries_iter():
    with cf.CompoundFileReader('tests/example.dat') as doc:
        assert len([e for e in doc.root]) == 1
        assert len([e for e in doc]) == 1
        assert len(doc) == 1

def test_entries_index():
    with cf.CompoundFileReader('tests/example.dat') as doc:
        assert doc.root[0] == doc.root['Storage 1']
        assert doc[0] == doc.root['Storage 1']

def test_entries_bytes():
    with cf.CompoundFileReader('tests/example.dat') as doc:
        assert b'Storage 1' in doc.root
        assert doc.root['Storage 1'] is doc.root[b'Storage 1']
        assert doc.open('Storage 1/Stream 1').read() == doc.open(b'Storage 1/Stream 1').read()

def test_entries_not_contains():
    with cf.CompoundFileReader('tests/example.dat') as doc:
        assert 'Storage 2' not in doc.root
        assert 'Storage 2' not in doc
        assert doc.root['Storage 1']['Stream 1'] not in doc.root

def test_entry_not_found():
    with cf.CompoundFileReader('tests/example.dat') as doc:
        with pytest.raises(cf.CompoundFileNotFoundError):
            doc.open('Storage 1/Stream 2')

def test_entry_not_stream():
    with cf.CompoundFileReader('tests/example.dat') as doc:
        with pytest.raises(cf.CompoundFileNotStreamError):
            doc.open('Storage 1')

def test_invalid_name():
    # Same file as example.dat with corrupted names (no NULL terminator in
    # name1, incorrect length in name2); library continues as normal in these
    # cases
    with warnings.catch_warnings(record=True) as w:
        with cf.CompoundFileReader('tests/invalid_name1.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileDirNameWarning)
            assert len(w) == 1
            verify_example(doc)
    with warnings.catch_warnings(record=True) as w:
        with cf.CompoundFileReader('tests/invalid_name2.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileDirNameWarning)
            assert len(w) == 1
            verify_example(doc)

def test_invalid_root_type():
    with warnings.catch_warnings(record=True) as w:
        # Same file as example.dat with root dir-entry type corrupted; in this
        # case the library corrects the dir-entry type and continues
        with cf.CompoundFileReader('tests/invalid_root_type.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileDirTypeWarning)
            assert len(w) == 1
            verify_example(doc)

def test_invalid_stream_type():
    with warnings.catch_warnings(record=True) as w:
        # Same file as example.dat with Stream 1 dir-entry type corrupted; in
        # this case the library ignores the Stream 1 entry
        with cf.CompoundFileReader('tests/invalid_stream_type.dat') as doc:
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
        with cf.CompoundFileReader('tests/invalid_dir_indexes1.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileDirIndexWarning)
            assert issubclass(w[1].category, cf.CompoundFileDirIndexWarning)
            assert issubclass(w[2].category, cf.CompoundFileDirIndexWarning)
            assert len(w) == 3
            verify_example(doc)
    with warnings.catch_warnings(record=True) as w:
        # Same file as example.dat with the Stream 1 siblings corrupted;
        # library continues as normal
        with cf.CompoundFileReader('tests/invalid_dir_indexes2.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileDirIndexWarning)
            assert issubclass(w[1].category, cf.CompoundFileDirIndexWarning)
            assert len(w) == 2
            verify_example(doc)
    with warnings.catch_warnings(record=True) as w:
        # Same file as example.dat with the root child index corrupted; library
        # continues but file is effectively empty
        with cf.CompoundFileReader('tests/invalid_dir_indexes3.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileDirIndexWarning)
            assert len(w) == 1
            verify_example(doc, ())

def test_invalid_dir_misc():
    with warnings.catch_warnings(record=True) as w:
        # Same file as example.dat with UUID and timestamps corrupted in
        # Stream 1; library continues as normal
        with cf.CompoundFileReader('tests/invalid_dir_misc.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileDirEntryWarning)
            assert issubclass(w[1].category, cf.CompoundFileDirTimeWarning)
            assert issubclass(w[2].category, cf.CompoundFileDirTimeWarning)
            assert len(w) == 3
            verify_example(doc)

def test_invalid_dir_size():
    with warnings.catch_warnings(record=True) as w:
        # Same file as example.dat with size of Stream 1 corrupted (> 32-bits);
        # library re-writes size to be within 32-bits
        with cf.CompoundFileReader('tests/invalid_dir_size1.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileDirSizeWarning)
            assert issubclass(w[1].category, cf.CompoundFileDirSizeWarning)
            assert len(w) == 2
            verify_example(doc, (
                DirEntry('Storage 1', False, 1),
                DirEntry('Storage 1/Stream 1', True, 0xFFFFFFFF),
                ), test_contents=False)

def test_invalid_dir_loop():
    with pytest.raises(cf.CompoundFileDirLoopError):
        # Same as example.dat but with Stream 1's left pointer corrupted to
        # point to the Root Entry
        doc = cf.CompoundFileReader('tests/invalid_dir_loop.dat')

def test_invalid_fat_loop():
    with pytest.raises(cf.CompoundFileNormalLoopError):
        # Same as example.dat but with Stream 1's FAT entry corrupted to
        # point to itself
        doc = cf.CompoundFileReader('tests/invalid_fat_loop.dat')

def test_invalid_magic():
    with pytest.raises(cf.CompoundFileInvalidMagicError):
        # Same as example.dat with corrupted magic block at the start
        cf.CompoundFileReader('tests/invalid_magic.dat')

def test_invalid_bom():
    with pytest.raises(cf.CompoundFileInvalidBomError):
        # Same as example.dat with big endian BOM
        cf.CompoundFileReader('tests/invalid_big_endian_bom.dat')
    with pytest.raises(cf.CompoundFileInvalidBomError):
        # Same as example.dat with invalid BOM
        cf.CompoundFileReader('tests/invalid_bom.dat')

def test_invalid_sector_size():
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with sector size shifts corrupted (sector size
        # 4 bytes, mini sector size 2 bytes); reader assumes 512 and 64 bytes
        # respectively as sector sizes <128 bytes are impossible, and <512
        # bytes are highly unlikely
        with cf.CompoundFileReader('tests/invalid_sector_size.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileSectorSizeWarning)
            assert issubclass(w[1].category, cf.CompoundFileSectorSizeWarning)
            assert len(w) == 2
            verify_example(doc)

def test_invalid_dir_sector_count():
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with dir sector count filled in (not allowed in
        # v3 file); reader ignores the value in both v3 and v4 anyway
        with cf.CompoundFileReader('tests/invalid_dir_sector_count.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileHeaderWarning)
            assert len(w) == 1
            verify_example(doc)

def test_invalid_master_sectors():
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with second master FAT block pointing beyond EOF;
        # reader ignores it and all further blocks
        with cf.CompoundFileReader('tests/invalid_master_eof.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileMasterFatWarning)
            assert len(w) == 1
            verify_example(doc)
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with second master FAT block set to
        # MASTER_FAT_SECTOR; reader ignores it and all further blocks
        with cf.CompoundFileReader('tests/invalid_master_special.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileMasterFatWarning)
            assert len(w) == 1
            verify_example(doc)

def test_invalid_master_ext():
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with master FAT extension pointer set to
        # FREE_SECTOR (should be END_OF_CHAIN for no extension); reader assumes
        # END_OF_CHAIN
        with cf.CompoundFileReader('tests/invalid_master_ext_free.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileMasterFatWarning)
            assert len(w) == 1
            verify_example(doc)
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with master FAT extension pointer set to 6
        # (beyond EOF) and count set to 0
        with cf.CompoundFileReader('tests/invalid_master_ext_eof.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileMasterFatWarning)
            assert len(w) == 1
            verify_example(doc)
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with master FAT extension counter set to 2,
        # but extension pointer set to END_OF_CHAIN
        with cf.CompoundFileReader('tests/invalid_master_ext_count.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileMasterFatWarning)
            assert issubclass(w[1].category, cf.CompoundFileMasterFatWarning)
            assert len(w) == 2
            verify_example(doc)

def test_invalid_fat_len():
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with length of FAT in header tweaked to 2 (should
        # be 1); reader ignores erroneous length
        with cf.CompoundFileReader('tests/invalid_fat_len.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileMasterFatWarning)
            assert len(w) == 1
            verify_example(doc)

def test_invalid_fat_types():
    with warnings.catch_warnings(record=True) as w:
        # Same as strange_master_ext.dat with a couple of sectors mis-marked in
        # the FAT (both marked as FREE_SECTOR when they should be
        # NORMAL_FAT_SECTOR and MASTER_FAT_SECTOR respectively)
        with cf.CompoundFileReader('tests/invalid_fat_types.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileMasterSectorWarning)
            assert issubclass(w[1].category, cf.CompoundFileNormalSectorWarning)
            assert len(w) == 2
            verify_example(doc)

def test_invalid_mini_fat():
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with the mini-FAT start sector corrupted to be
        # FREE_SECTOR; reader assumes no mini-FAT in this case
        with cf.CompoundFileReader('tests/invalid_mini_free.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileMiniFatWarning)
            assert len(w) == 1
            verify_example(doc, test_contents=False)
            with pytest.raises(cf.CompoundFileNoMiniFatError):
                doc.open('Storage 1/Stream 1')
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with the mini-FAT start sector corrupted to be
        # beyond EOF; reader assumes no mini-FAT in this case
        with cf.CompoundFileReader('tests/invalid_mini_eof.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileMiniFatWarning)
            assert len(w) == 1
            verify_example(doc, test_contents=False)

def test_invalid_header_misc():
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with uuid, txn, and reserved fields set to
        # non-zero values; violates spec but reader ignores it after warning
        with cf.CompoundFileReader('tests/invalid_header_misc.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileHeaderWarning)
            assert issubclass(w[1].category, cf.CompoundFileHeaderWarning)
            assert issubclass(w[2].category, cf.CompoundFileHeaderWarning)
            assert len(w) == 3
            verify_example(doc)

def test_strange_sector_size_v3():
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with sector size re-written to 1024 and file
        # content padded out accordingly; reader warns out this but carries
        # on anyway
        with cf.CompoundFileReader('tests/strange_sector_size_v3.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileSectorSizeWarning)
            assert len(w) == 1
            verify_example(doc)

def test_strange_sector_size_v4():
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with DLL version set to 4 and dir sector count
        # filled in (required in v4 files); reader warns about unusual
        # sector size by carries on anyway
        with cf.CompoundFileReader('tests/strange_sector_size_v4.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileSectorSizeWarning)
            assert len(w) == 1
            verify_example(doc)

def test_strange_dll_version():
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with DLL version set to 5; reader warns about
        # this but ignores it
        with cf.CompoundFileReader('tests/strange_dll_version.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileVersionWarning)
            assert len(w) == 1
            verify_example(doc)

def test_strange_mini_sector_size():
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with mini sector size set to 128 and mini FAT
        # re-written accordingly
        with cf.CompoundFileReader('tests/strange_mini_sector_size.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileSectorSizeWarning)
            assert len(w) == 1
            verify_example(doc)

def test_strange_master_full():
    # Same as example.dat with FAT extended to include enough blank sectors to
    # fill the DIFAT without an extension (e.g. a compound document where lots
    # of contents have been deleted and the FAT hasn't been compressed but
    # simply overwritten)
    with warnings.catch_warnings(record=True) as w:
        with cf.CompoundFileReader('tests/strange_master_full.dat') as doc:
            assert len(w) == 0
            verify_example(doc)

def test_strange_master_ext():
    # Same as example.dat with FAT extended to include enough blank sectors to
    # necessitate a DIFAT extension (e.g. a compound document where lots of
    # contents have been deleted and the FAT hasn't been compressed but simply
    # overwritten)
    with warnings.catch_warnings(record=True) as w:
        with cf.CompoundFileReader('tests/strange_master_ext.dat') as doc:
            assert len(w) == 0
            verify_example(doc)

def test_invalid_master_loop():
    with pytest.raises(cf.CompoundFileMasterLoopError):
        # Same as strange_master_ext.dat but with DIFAT extension sector filled
        # and terminated with a self-reference
        doc = cf.CompoundFileReader('tests/invalid_master_loop.dat')

def test_invalid_master_len():
    with warnings.catch_warnings(record=True) as w:
        # Same as strange_master_ext.dat with master extension count set to 2
        # (should be 1); reader ignores DIFAT count
        with cf.CompoundFileReader('tests/invalid_master_underrun.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileMasterFatWarning)
            assert len(w) == 1
            verify_example(doc)
    with warnings.catch_warnings(record=True) as w:
        # Same as strange_master_ext.dat with master extension count set to 0
        # (should be 1); reader ignores DIFAT count
        with cf.CompoundFileReader('tests/invalid_master_overrun.dat') as doc:
            assert issubclass(w[0].category, cf.CompoundFileMasterFatWarning)
            assert issubclass(w[1].category, cf.CompoundFileMasterFatWarning)
            assert len(w) == 2
            verify_example(doc)

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

def test_stream_truncated():
    with warnings.catch_warnings(record=True) as w:
        # Same as example.dat with mini FAT and directory entries re-written
        # to extend Stream 1 beyond the end of the compound document
        with cf.CompoundFileReader('tests/invalid_truncated.dat') as doc:
            with doc.open('Storage 1/Stream 1') as f:
                f.read()
            assert issubclass(w[0].category, cf.CompoundFileTruncatedWarning)
            assert len(w) == 1
            verify_example(doc, (
                DirEntry('Storage 1', False, 1),
                DirEntry('Storage 1/Stream 1', True, 1500),
                ), test_contents=False)
