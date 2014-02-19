#!/usr/bin/env python3
# vim: set et sw=4 sts=4 fileencoding=utf-8:
#
# A library for reading Microsoft's OLE Compound Document format
# Copyright (c) 2014 Dave Hughes <dave@waveform.org.uk>
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
#     * Redistributions of source code must retain the above copyright
#       notice, this list of conditions and the following disclaimer.
#     * Redistributions in binary form must reproduce the above copyright
#       notice, this list of conditions and the following disclaimer in the
#       documentation and/or other materials provided with the distribution.
#     * Neither the name of the copyright holder nor the
#       names of its contributors may be used to endorse or promote products
#       derived from this software without specific prior written permission.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
# ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
# LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
# CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
# SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
# INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
# ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
# POSSIBILITY OF SUCH DAMAGE.

"""
Package root for compound_files

Most of the work in this package was derived from the specification for `OLE
Compound Document`_ files published by OpenOffice, and the specification for
the `Advanced Authoring Format`_ (AAF) published by Microsoft.

.. _OLE Compound Document: http://www.openoffice.org/sc/compdocfileformat.pdf
.. _Advanced Authoring Format: http://www.amwa.tv/downloads/specifications/aafcontainerspec-v1.0.1.pdf
"""

from __future__ import (
    unicode_literals,
    absolute_import,
    print_function,
    division,
    )
str = type('')


import io
import os
import sys
import struct as st
import logging
import warnings
from array import array


__all__ = [
    'CompoundFileError',
    'CompoundFileWarning',
    'CompoundFileReader',
    'CompoundFileStream',
    ]

# A quick personal rant: the AAF or OLE Compound Document format is yet another
# example of bad implementations of a bad specification (thanks Microsoft! See
# the W3C log file format for previous examples of MS' incompetence in this
# area)...
#
# The specification doesn't try and keep the design simple (the DIFAT could be
# fully in the header or partially in the header, and the header itself doesn't
# necessarily match the sector size), whoever wrote the spec didn't quite
# understand what version numbers are used for (several versions exist, but the
# spec doesn't specify exactly which bits of the header became relevant in
# which versions), and the spec has huge amounts of redundancy (always fun as
# it inevitably leads to implementations getting one bit right and another bit
# wrong, leaving readers to guess which is correct).
#
# TL;DR: if you're looking for a nice fast binary format with good random
# access characteristics this may look attractive, but please don't use it.
# Ideally, loop-mounting a proper file-system would be the way to go, although
# it generally involves jumping through several hoops due to mount being a
# privileged operation.

# Magic identifier at the start of the file
COMPOUND_MAGIC = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'

FREE_SECTOR       = 0xFFFFFFFF # denotes an unallocated (free) sector
END_OF_CHAIN      = 0xFFFFFFFE # denotes the end of a stream chain
NORMAL_FAT_SECTOR = 0xFFFFFFFD # denotes a sector used for the regular FAT
MASTER_FAT_SECTOR = 0xFFFFFFFC # denotes a sector used for the master FAT
MAX_NORMAL_SECTOR = 0xFFFFFFFA # the maximum sector in a file

MAX_REG_SID    = 0xFFFFFFFA # maximum directory entry ID
NO_STREAM      = 0xFFFFFFFF # unallocated directory entry

DIR_INVALID    = 0 # unknown/empty(?) storage type
DIR_STORAGE    = 1 # element is a storage (dir) object
DIR_STREAM     = 2 # element is a stream (file) object
DIR_LOCKBYTES  = 3 # element is an ILockBytes object
DIR_PROPERTY   = 4 # element is an IPropertyStorage object
DIR_ROOT       = 5 # element is the root storage object


COMPOUND_HEADER = st.Struct(b''.join((
    b'<',    # little-endian format
    b'8s',   # magic string
    b'16s',  # file UUID (ignored)
    b'H',    # file header major version
    b'H',    # file header minor version
    b'H',    # byte order mark
    b'H',    # sector size (actual size is 2**sector_size)
    b'H',    # mini sector size (actual size is 2**short_sector_size)
    b'6s',   # unused
    b'L',    # sector count of directory chain
    b'L',    # FAT sector count
    b'L',    # ID of first sector of the FAT
    b'L',    # transaction signature
    b'L',    # minimum size of a normal stream
    b'L',    # ID of first sector of the mini-FAT
    b'L',    # mini-FAT sector count
    b'L',    # ID of first sector of the master-FAT
    b'L',    # master-FAT sector count
    )))


class CompoundFileError(IOError):
    pass


class CompoundFileWarning(Warning):
    pass


# XXX RawIOBase?
class CompoundFileStream(io.IOBase):
    """
    Represents a stream within an OLE Compound Document.

    Instances of :class:`CompoundFileStream` are returned by the
    :meth:`CompoundFileReader.open` method. They support all common methods
    associated with read-only streams (:meth:`read`, :meth:`seek`,
    :meth:`tell`, and so forth).

    .. note::

        The implementation attempts to duplicate the parent object's file
        descriptor upon construction which theoretically means multiple threads
        can simultaneously read different files in the compound document.
        However, if duplication of the file descriptor fails for any reason,
        the implementation falls back on sharing the parent object's file
        descriptor. In this case, thread safety is not guaranteed. Check the
        :attr:`thread_safe` attribute to determine if duplication succeeded.
    """

    def __init__(self, parent, start, length=None):
        super(CompoundFileStream, self).__init__()
        self._sector_index = 0
        self._sector_offset = 0
        self._sector_size = parent._normal_sector_size
        self._header_size = parent._header_size
        try:
            fd = os.dup(parent._file.fileno())
        except AttributeError, OSError:
            # Share the parent's _file if we fail to duplicate the descriptor
            self._file = parent._file
            self.thread_safe = False
        else:
            self._file = io.open(fd, 'rb')
            self.thread_safe = True
        try:
            finish = start
            while True:
                sector = parent._normal_fat[finish]
                if parent._max_sector < sector < MAX_NORMAL_SECTOR:
                    warnings.warn(
                            'invalid sector in stream chain (%d)' % value,
                            CompoundFileWarning)
                    break
                elif sector == END_OF_CHAIN:
                    break
                finish += 1
        except IndexError:
            warnings.warn(
                    'missing end of chain for file at sector %d' % start,
                    CompoundFileWarning)
            finish -= 1
        if start == finish:
            raise CompoundFileError(
                    'empty chain for file starting at sector %d' % start)
        # On Python 3 we guard against malicious files which repeat the same
        # enormous FAT chain for multiple files by simply acquiring a memory
        # view of the chain in the parent's FAT. On Python 2, memoryviews
        # don't work with arrays (despite the doc's claim) so we have to hope
        # the protections built into parsing the directory tree are sufficient
        if sys.version_info[0] >= 3:
            self._sectors = memoryview(parent._normal_fat)[start:finish]
        else:
            self._sectors = parent._normal_fat[start:finish]
        if length is None:
            length = len(self._sectors) * self._sector_size
        self._length = length

    def close(self):
        """
        Close the file pointer.
        """
        if self.thread_safe:
            try:
                self._file.close()
            except AttributeError:
                pass
        self._file = None

    def readable(self):
        """
        Returns ``True``, indicating that the stream supports :meth:`read`.
        """
        return True

    def writable(self):
        """
        Returns ``False``, indicating that the stream doesn't support
        :meth:`write`.
        """
        return False

    def seekable(self):
        """
        Returns ``True``, indicating that the stream supports :meth:`seek`.
        """
        return True

    def _set_pos(self, value):
        self._sector_index = value // self._sector_size
        self._sector_offset = value % self._sector_size
        if self._sector_index < len(self._sectors):
            self._file.seek(
                    self._header_size +
                    (self._sectors[self._sector_index] * self._sector_size) +
                    self._sector_offset)

    def tell(self):
        """
        Return the current stream position.
        """
        return (self._sector_index * self._sector_size) + self._sector_offset

    def seek(self, offset, whence=io.SEEK_SET):
        """
        Change the stream position to the given byte *offset*. *offset* is
        interpreted relative to the position indicated by *whence*. Values for
        *whence* are:

        * ``SEEK_SET`` or ``0`` – start of the stream (the default); *offset*
          should be zero or positive

        * ``SEEK_CUR`` or ``1`` – current stream position; *offset* may be
          negative

        * ``SEEK_END`` or ``2`` – end of the stream; *offset* is usually
          negative

        Return the new absolute position.
        """
        if whence == io.SEEK_CUR:
            offset = self.tell() + offset
        elif whence == io.SEEK_END:
            offset = self._length + offset
        if offset < 0:
            raise ValueError(
                    'New position is before the start of the stream')
        self._set_pos(offset)
        return offset

    def read1(self, n=-1):
        """
        Read up to *n* bytes from the stream using only a single call to the
        underlying object.

        In the case of :class:`CompoundFileStream` this roughly corresponds to
        returning the content from the current position up to the end of the
        current sector.
        """
        if not self.thread_safe:
            # If we're sharing a file-pointer with the parent object we can't
            # guarantee the file point is where we left it, so force a seek
            self._set_pos(self.tell())
        if n == -1:
            n = max(0, self._length - self.tell())
        else:
            n = max(0, min(n, self._length - self.tell()))
        n = min(n, self._sector_size - self._sector_offset)
        if n == 0:
            return b''
        try:
            result = self._file.read1(n)
        except AttributeError:
            result = self._file.read(n)
        assert len(result) == n
        self._sector_index += 1
        self._sector_offset = 0
        if self._sector_index < len(self._sectors):
            self._file.seek(
                    self._header_size +
                    self._sectors[self._sector_index])
        return result

    def read(self, n=-1):
        """
        Read up to *n* bytes from the stream and return them. As a convenience,
        if *n* is unspecified or -1, :meth:`readall` is called. Fewer than *n*
        bytes may be returned if there are fewer than *n* bytes from the
        current stream position to the end of the stream.

        If 0 bytes are returned, and *n* was not 0, this indicates end of the
        stream.
        """
        if n == -1:
            n = max(0, self._length - self.tell())
        else:
            n = max(0, min(n, self._length - self.tell()))
        result = b''
        while n > 0:
            buf = self.read1(n)
            if not buf:
                break
            n -= len(buf)
            result += buf
        return result


class CompoundFileReader(object):
    """
    Provides an interface for reading `OLE Compound Document`_ files.

    The :class:`CompoundFileReader` class provides a relatively simple
    interface for interpreting the content of Microsoft's `OLE Compound
    Document`_ files. These files can be thought of as a file-system in a file
    (or a loop-mounted FAT file-system for Unix folk).

    The class can be constructed with a filename or a file-like object. In the
    latter case, the object must support the ``read``, ``seek``, and ``tell``
    methods. For optimal usage, it should also provide a valid fd in response
    to a call to ``fileno``, and provide a ``read1`` method, but these are not
    mandatory.

    Instances of the class act can be enumerated to obtain a sequence of
    :class:`CompoundDirEntry` objects which represent the elements of the root
    directory in the file. An :meth:`open` method is provided which returns a
    stream representing the content of files stored within the compound
    document.

    Finally, context manager protocol is also supported, permitting usage of
    the class like so::

        with CompoundFileReader('foo.doc') as doc:
            # Iterate over items in the root directory of the compound document
            for entry in doc:
                # If any entry is a file, attempt to read the data from it
                if entry.isfile:
                    with doc.open(entry) as f:
                        f.read()

    .. _OLE Compound Document: http://www.openoffice.org/sc/compdocfileformat.pdf
    """

    def __init__(self, filename_or_obj):
        super(CompoundFileReader, self).__init__()
        if isinstance(filename_or_obj, (str, bytes)):
            self._opened = True
            self._file = io.open(filename_or_obj, 'rb')
        else:
            self._opened = False
            self._file = filename_or_obj
        self._master_fat = array(b'L')
        self._normal_fat = array(b'L')
        self._mini_fat = array(b'L')
        (
            magic,
            uuid,
            self._minor_version,
            self._dll_version,
            bom,
            normal_sector_size,
            mini_sector_size,
            unused,
            self._dir_sector_count,
            self._normal_sector_count,
            self._normal_first_sector,
            txn_signature,
            self._mini_size_limit,
            self._mini_first_sector,
            self._mini_sector_count,
            self._master_first_sector,
            self._master_sector_count,
        ) = COMPOUND_HEADER.unpack(self._file.read(COMPOUND_HEADER.size))
        if magic != COMPOUND_MAGIC:
            raise CompoundFileError(
                    '%s does not appear to be an OLE compound '
                    'document' % filename_or_obj)
        if bom != 0xFFFE:
            raise CompoundFileError(
                    '%s uses an unsupported byte ordering (big '
                    'endian)' % filename_or_obj)
        self._normal_sector_size = 2 ** normal_sector_size
        self._normal_sector_format = st.Struct(
                bytes('<%dL' % (self._normal_sector_size // 4)))
        assert self._normal_sector_size == self._normal_sector_format.size
        self._mini_sector_size = 2 ** mini_sector_size
        self._mini_sector_format = st.Struct(
                bytes('<%dL' % (self._mini_sector_size // 4)))
        assert self._mini_sector_size == self._mini_sector_format.size
        if self._dll_version == 3:
            if self._normal_sector_size != 512:
                warnings.warn(
                        'unexpected sector size in v3 file '
                        '(%d)' % self._normal_sector_size, CompoundFileWarning)
            if self._dir_sector_count != 0:
                warnings.warn(
                        'directory chain sector count is non-zero '
                        '(%d)' % self._dir_sector_count, CompoundFileWarning)
        elif self._dll_version == 4:
            if self._normal_sector_size != 4096:
                warnings.warn(
                        'unexpected sector size in v4 file '
                        '(%d)' % self._normal_sector_size, CompoundFileWarning)
        else:
            raise CompoundFileError(
                    'unsupported DLL version (%d)' % self._dll_version)
        if self._mini_sector_size != 64:
            warnings.warn(
                    'unexpected mini sector size '
                    '(%d)' % self._mini_sector_size, CompoundFileWarning)
        if uuid != '\x00' * 16:
            warnings.warn(
                    'CLSID of compound file is non-zero (%r)' % uuid,
                    CompoundFileWarning)
        if txn_signature != 0:
            warnings.warn(
                    'transaction signature is non-zero '
                    '(%d)' % txn_signature, CompoundFileWarning)
        if unused != (b'\x00'*6):
            warnings.warn(
                    'unused header bytes are non-zero '
                    '(%r)' % unused, CompoundFileWarning)
        self._file.seek(0, io.SEEK_END)
        self._file_size = self._file.tell()
        self._header_size = max(self._normal_sector_size, 512)
        self._max_sector = (self._file_size - self._header_size) // self._normal_sector_size
        self._load_normal_fat(self._load_master_fat())

    def close(self):
        if self._opened:
            self._file.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()

    def _seek_sector(self, sector):
        if sector > self._max_sector:
            raise CompoundFileError('seek to invalid sector (%d)' % sector)
        self._file.seek(
                self._header_size + (sector * self._normal_sector_size))

    def _load_master_fat(self):
        # Note: when reading the DIFAT we deliberately disregard the DIFAT
        # sector count read from the header as implementations may set this
        # incorrectly. Instead, we scan for END_OF_CHAIN (or FREE_SECTOR) in
        # the DIFAT after each read and stop when we find it. In order to avoid
        # infinite loops (in the case of a *really* stupid file) we keep track
        # of each sector we seek to and quit in the event of a repeat
        self._master_fat = array(b'L')
        count = self._master_sector_count
        checked = 0
        sectors = set()

        # Special case: the first 109 entries are stored at the end of the file
        # header and the next sector of the DIFAT is stored in the header
        self._file.seek(COMPOUND_HEADER.size)
        self._master_fat.extend(
                st.unpack(b'<109L', self._file.read(109 * 4)))
        sector = self._master_first_sector
        if count == 0 and sector != END_OF_CHAIN:
            warnings.warn(
                    'DIFAT pointer with zero count', CompoundFileWarning)
        elif count != 0 and sector == END_OF_CHAIN:
            warnings.warn(
                    'DIFAT chained from header, '
                    'or incorrect count', CompoundFileWarning)
            sector = self._master_fat.pop()

        while True:
            # Check for an END_OF_CHAIN marker in the existing stream
            for index in range(checked, len(self._master_fat) - 1):
                value = self._master_fat[index]
                if value == END_OF_CHAIN:
                    break
                elif value == FREE_SECTOR:
                    warnings.warn(
                            'DIFAT terminated by FREE_SECTOR',
                            CompoundFileWarning)
                    value = END_OF_CHAIN
                    break
                elif self._max_sector < value < MAX_NORMAL_SECTOR:
                    warnings.warn(
                            'invalid sector in DIFAT chain (%d)' % value,
                            CompoundFileWarning)
                    value = END_OF_CHAIN
                    break
            if value == END_OF_CHAIN:
                del self._master_fat[index:]
                break
            checked = len(self._master_fat)
            # Step case: if we're reading a subsequent block we need to seek to
            # the indicated sector, read it, and find the next sector in the
            # last value
            count -= 1
            sectors.add(sector)
            self._seek_sector(sector)
            self._master_fat.extend(
                    self._normal_sector_format.unpack(
                        self._file.read(self._normal_sector_format.size)))
            # Guard against malicious files which could cause excessive memory
            # allocation when reading the FAT. If the FAT alone would exceed
            # 100Mb of RAM, raise an error
            if len(self._master_fat) * self._normal_sector_size > 100*1024*1024:
                raise CompoundFileError(
                        'excessively large FAT (malicious file?)')
            sector = self._master_fat.pop()
            if sector in sectors:
                raise CompoundFileError(
                        'DIFAT loop encountered (sector %d)' % sector)

        if count > 0:
            warnings.warn(
                    'DIFAT end encountered early (expected %d more '
                    'sectors)' % count, CompoundFileWarning)
        elif count < 0:
            warnings.warn(
                    'DIFAT end encountered late (overran by %d '
                    'sectors)' % -count, CompoundFileWarning)
        if len(self._master_fat) != self._normal_sector_count:
            warnings.warn(
                    'DIFAT length does not match FAT sector count '
                    '(%d != %d)' % (len(self._master_fat), self._normal_sector_count),
                    CompoundFileWarning)
        return sectors

    def _load_normal_fat(self, master_sectors):
        # Again, when reading the FAT we deliberately disregard the FAT sector
        # read from the header as some implementations get it wrong. Instead,
        # we just read the sectors that the DIFAT chain tells us to (no need to
        # check for loops or invalid sectors here though - the _load_master_fat
        # method takes of those). After reading the FAT we check the DIFAT
        # sectors are marked correctly.
        self._normal_fat = array(b'L')
        for sector in self._master_fat:
            self._seek_sector(sector)
            self._normal_fat.extend(
                    self._normal_sector_format.unpack(
                        self._file.read(self._normal_sector_format.size)))

        # The following simply verifies that all FAT and DIFAT sectors are
        # marked appropriately in the FAT
        for master_sector in master_sectors:
            if self._normal_fat[master_sector] != MASTER_FAT_SECTOR:
                warnings.warn(
                        'DIFAT sector %d marked incorrectly in FAT '
                        '(%d != %d)' % (
                            master_sector,
                            self._normal_fat[master_sector],
                            MASTER_FAT_SECTOR,
                            ), CompoundFileWarning)
        for normal_sector in self._master_fat:
            if self._normal_fat[normal_sector] != NORMAL_FAT_SECTOR:
                warnings.warn(
                        'FAT sector %d marked incorrectly in FAT '
                        '(%d != %d)' % (
                            normal_sector,
                            self._normal_fat[normal_sector],
                            NORMAL_FAT_SECTOR,
                            ), CompoundFileWarning)

