#!/usr/bin/env python3
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

"""
Most of the work in this package was derived from the specification for `OLE
Compound Document`_ files published by OpenOffice, and the specification for
the `Advanced Authoring Format`_ (AAF) published by Microsoft.

.. _OLE Compound Document: http://www.openoffice.org/sc/compdocfileformat.pdf
.. _Advanced Authoring Format: http://www.amwa.tv/downloads/specifications/aafcontainerspec-v1.0.1.pdf


CompoundFileReader
==================

.. autoclass:: CompoundFileReader
    :members:


CompoundFileStream
==================

.. autoclass:: CompoundFileStream
    :members:


CompoundFileEntity
==================

.. autoclass:: CompoundFileEntity
    :members:


Exceptions
==========

.. autoexception:: CompoundFileError

.. autoexception:: CompoundFileWarning

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
import datetime as dt
from pprint import pformat
from array import array


__all__ = [
    'CompoundFileError',
    'CompoundFileWarning',
    'CompoundFileReader',
    'CompoundFileNormalStream',
    'CompoundFileMiniStream',
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
# privileged operation.</rant>
#
# In the interests of trying to keep naming vaguely consistent and sensible
# here's a translation list with the names we'll be using first and the names
# other documents use after:
#
#   normal-FAT = FAT = SAT
#   master-FAT = DIFAT = DIF = MSAT
#   mini-FAT = miniFAT = SSAT
#
# And here's a brief description of the compound document structure:
#
# Compound documents consist of a header, followed by a number of equally sized
# sectors numbered incrementally. Within the sectors are stored the master-FAT,
# normal-FAT, and (optional) mini-FAT, directory entries, and file streams. A
# FAT is simply an indexed list of sectors, with each sector pointing to the
# next in the chain, the last holding the END_OF_CHAIN value.
#
# The master-FAT (the location of which is determined by the header) stores
# which sectors are occupied by the normal-FAT. It must be read first in order
# to read sectors that make up the normal-FAT in order.
#
# The normal-FAT stores the locations of directory entries, file streams, and
# tracks which sectors are allocated to the master-FAT and itself.
#
# The mini-FAT (if present) is stored as a file stream, virtually divided into
# smaller sectors for the purposes of efficiently storing files smaller than
# the normal sector size.

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

FILENAME_ENCODING = 'latin-1'


COMPOUND_HEADER = st.Struct(b''.join((
    b'<',    # little-endian format
    b'8s',   # magic string
    b'16s',  # file UUID (unused)
    b'H',    # file header major version
    b'H',    # file header minor version
    b'H',    # byte order mark
    b'H',    # sector size (actual size is 2**sector_size)
    b'H',    # mini sector size (actual size is 2**short_sector_size)
    b'6s',   # unused
    b'L',    # directory chain sector count
    b'L',    # normal-FAT sector count
    b'L',    # ID of first sector of the normal-FAT
    b'L',    # transaction signature (unused)
    b'L',    # minimum size of a normal stream
    b'L',    # ID of first sector of the mini-FAT
    b'L',    # mini-FAT sector count
    b'L',    # ID of first sector of the master-FAT
    b'L',    # master-FAT sector count
    )))

DIR_HEADER = st.Struct(b''.join((
    b'<',    # little-endian format
    b'64s',  # NULL-terminated filename in UTF-16 little-endian encoding
    b'H',    # length of filename (why?!)
    b'B',    # dir-entry type
    b'B',    # red (0) or black (1) entry
    b'L',    # ID of left-sibling node
    b'L',    # ID of right-sibling node
    b'L',    # ID of children's root node
    b'16s',  # dir-entry UUID (unused)
    b'L',    # user flags (unused)
    b'Q',    # creation timestamp
    b'Q',    # modification timestamp
    b'L',    # start sector of stream
    b'L',    # low 32-bits of stream size
    b'L',    # high 32-bits of stream size
    )))


class CompoundFileError(IOError):
    """
    Base class for exceptions arising from reading compound documents.
    """


class CompoundFileWarning(Warning):
    """
    Base class for warnings arising from reading compound documents.
    """


class CompoundFileStream(io.RawIOBase):
    """
    Abstract base class for streams within an OLE Compound Document.

    Instances of :class:`CompoundFileStream` are not constructed
    directly, but are returned by the :meth:`CompoundFileReader.open` method.
    They support all common methods associated with read-only streams
    (:meth:`read`, :meth:`seek`, :meth:`tell`, and so forth).

    .. note::

        The implementation attempts to duplicate the parent object's file
        descriptor upon construction which theoretically means multiple threads
        can simultaneously read different files in the compound document.
        However, if duplication of the file descriptor fails for any reason,
        the implementation falls back on sharing the parent object's file
        descriptor. In this case, thread safety is not guaranteed. Check the
        :attr:`thread_safe` attribute to determine if duplication succeeded.
    """
    def __init__(self):
        super(CompoundFileStream, self).__init__()
        self._sectors = array(b'L')
        self._sector_index = None
        self._sector_offset = None

    def _load_sectors(self, start, fat):
        # To guard against cyclic FAT chains we use the tortoise'n'hare
        # algorithm here. If hare is ever equal to tortoise after a step, then
        # the hare somehow got transported behind the tortoise (via a loop) so
        # we raise an error
        hare = start
        tortoise = start
        while tortoise != END_OF_CHAIN:
            self._sectors.append(tortoise)
            tortoise = fat[tortoise]
            if hare != END_OF_CHAIN:
                hare = fat[hare]
                if hare != END_OF_CHAIN:
                    hare = fat[hare]
                    if hare == tortoise:
                        raise CompoundFileError(
                                'cyclic FAT chain found starting at %d' % start)

    def _set_pos(self, value):
        self._sector_index = value // self._sector_size
        self._sector_offset = value % self._sector_size
        if self._sector_index < len(self._sectors):
            self._file.seek(
                    self._header_size +
                    (self._sectors[self._sector_index] * self._sector_size) +
                    self._sector_offset)

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
        :meth:`write` or :meth:`truncate`.
        """
        return False

    def seekable(self):
        """
        Returns ``True``, indicating that the stream supports :meth:`seek`.
        """
        return True

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

        * ``SEEK_SET`` or ``0`` - start of the stream (the default); *offset*
          should be zero or positive

        * ``SEEK_CUR`` or ``1`` - current stream position; *offset* may be
          negative

        * ``SEEK_END`` or ``2`` - end of the stream; *offset* is usually
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
            # guarantee the file pointer is where we left it, so force a seek
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
        # Only perform a seek to a different sector if we've crossed into one
        if self._sector_offset + n < self._sector_size:
            self._sector_offset += n
        else:
            self._set_pos(self.tell() + n)
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


class CompoundFileNormalStream(CompoundFileStream):
    def __init__(self, parent, start, length=None):
        super(CompoundFileNormalStream, self).__init__()
        self._load_sectors(start, parent._normal_fat)
        self._sector_size = parent._normal_sector_size
        self._header_size = parent._header_size
        try:
            fd = os.dup(parent._file.fileno())
        except (AttributeError, OSError) as e:
            # Share the parent's _file if we fail to duplicate the descriptor
            self._file = parent._file
            self.thread_safe = False
        else:
            self._file = io.open(fd, 'rb')
            self.thread_safe = True
        min_length = (len(self._sectors) - 1) * self._sector_size
        max_length = len(self._sectors) * self._sector_size
        if length is None:
            self._length = max_length
        elif not (min_length <= length <= max_length):
            warnings.warn(
                    'length (%d) of stream at sector %d exceeds bounds '
                    '(%d-%d)' % (length, start, min_length, max_length),
                    CompoundFileWarning)
            self._length = max_length
        else:
            self._length = length
        self._set_pos(0)


class CompoundFileMiniStream(CompoundFileStream):
    def __init__(self, parent, start, length=None):
        super(CompoundFileMiniStream, self).__init__()
        self._load_sectors(start, parent._mini_fat)
        self._sector_size = parent._mini_sector_size
        self._header_size = 0
        self._file = CompoundFileNormalStream(
                parent, parent.root._start_sector, parent.root.size)
        self.thread_safe = self._file.thread_safe
        max_length = len(self._sectors) * self._sector_size
        if length is not None and length > max_length:
            warnings.warn(
                    'length (%d) of stream at sector %d exceeds max' % (
                        length, start, max_length),
                    CompoundFileWarning)
        self._length = min(max_length, length or max_length)
        self._set_pos(0)


class CompoundFileReader(object):
    """
    Provides an interface for reading `OLE Compound Document`_ files.

    The :class:`CompoundFileReader` class provides a relatively simple
    interface for interpreting the content of Microsoft's `OLE Compound
    Document`_ files. These files can be thought of as a file-system in a file
    (or a loop-mounted FAT file-system for Unix folk).

    The class can be constructed with a filename or a file-like object. In the
    latter case, the object must support the ``read``, ``seek``, and ``tell``
    methods. For optimal usage, it should also provide a valid file descriptor
    in response to a call to ``fileno``, and provide a ``read1`` method, but
    these are not mandatory.

    The :attr:`root` attribute represents the root storage entity in the
    compound document. An :meth:`open` method is provided which (given a
    :class:`CompoundFileEntity` instance representing a stream), returns a
    file-like object representing the content of the stream.

    Finally, the context manager protocol is also supported, permitting usage
    of the class like so::

        with CompoundFileReader('foo.doc') as doc:
            # Iterate over items in the root directory of the compound document
            for entry in doc.root:
                # If any entry is a file, attempt to read the data from it
                if entry.isfile:
                    with doc.open(entry) as f:
                        f.read()

    .. attribute:: root

        The root attribute represents the root storage entity in the compound
        document. As a :class:`CompoundFileEntity` instance, it (and child
        storages) can be enumerated, accessed by index, or by name (like a
        dict) to obtain :class:`CompoundFileEntity` instances representing the
        content of the compound document.

        Both :class:`CompoundFileReader` and :class:`CompoundFileEntity`
        support human-readable representations making it relatively simple to
        browse and extract information from compound documents simply by using
        the interactive Python command line.
    """

    def __init__(self, filename_or_obj):
        super(CompoundFileReader, self).__init__()
        if isinstance(filename_or_obj, (str, bytes)):
            self._opened = True
            self._file = io.open(filename_or_obj, 'rb')
        else:
            self._opened = False
            self._file = filename_or_obj

        self._master_fat = None
        self._normal_fat = None
        self._mini_fat = None
        self.root = None
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
            self._dir_first_sector,
            txn_signature,
            self._mini_size_limit,
            self._mini_first_sector,
            self._mini_sector_count,
            self._master_first_sector,
            self._master_sector_count,
        ) = COMPOUND_HEADER.unpack(self._file.read(COMPOUND_HEADER.size))

        # Check the header for basic correctness
        if magic != COMPOUND_MAGIC:
            raise CompoundFileError(
                    '%s does not appear to be an OLE compound '
                    'document' % filename_or_obj)
        if bom != 0xFFFE:
            raise CompoundFileError(
                    '%s uses an unsupported byte ordering (big '
                    'endian)' % filename_or_obj)
        if normal_sector_size > 20:
            warnings.warn(
                    'FAT sector size is excessively large, assuming 512',
                    CompoundFileWarning)
            normal_sector_size = 9
        if mini_sector_size >= normal_sector_size:
            warnings.warn(
                    'mini FAT sector size greater than or equal to FAT '
                    'sector size, assuming 64', CompoundFileWarning)
            mini_sector_size = 6
        self._normal_sector_size = 1 << normal_sector_size
        self._mini_sector_size = 1 << mini_sector_size
        self._normal_sector_format = st.Struct(
                bytes('<%dL' % (self._normal_sector_size // 4)))
        self._mini_sector_format = st.Struct(
                bytes('<%dL' % (self._mini_sector_size // 4)))
        assert self._normal_sector_size == self._normal_sector_format.size
        assert self._mini_sector_size == self._mini_sector_format.size

        # More correctness checks, but mostly warnings at this stage
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
        if uuid != (b'\0' * 16):
            warnings.warn(
                    'CLSID of compound file is non-zero (%r)' % uuid,
                    CompoundFileWarning)
        if txn_signature != 0:
            warnings.warn(
                    'transaction signature is non-zero '
                    '(%d)' % txn_signature, CompoundFileWarning)
        if unused != (b'\0' * 6):
            warnings.warn(
                    'unused header bytes are non-zero '
                    '(%r)' % unused, CompoundFileWarning)
        self._file.seek(0, io.SEEK_END)
        self._file_size = self._file.tell()
        self._header_size = max(self._normal_sector_size, 512)
        self._max_sector = (self._file_size - self._header_size) // self._normal_sector_size
        self._load_normal_fat(self._load_master_fat())
        self._load_mini_fat()
        self._load_directory()

    def open(self, filename_or_entity):
        """
        Return a file-like object with the content of the specified entity.

        Given a :class:`CompoundFileEntity` instance which represents a stream,
        or a string representing the path to one (using ``/`` separators), this
        method returns an instance of :class:`CompoundFileStream` which can be
        used to read the content of the stream.
        """
        if isinstance(filename_or_entity, bytes):
            filename_or_entity = filename_or_entity.decode(FILENAME_ENCODING)
        if isinstance(filename_or_entity, str):
            entity = self.root
            for name in filename_or_entity.split('/'):
                if name:
                    try:
                        entity = entity[name]
                    except KeyError:
                        raise CompoundFileError(
                                'unable to locate %s in compound '
                                'file' % filename_or_entity)
            filename_or_entity = entity
        if not filename_or_entity.isfile:
            raise CompoundFileError(
                    '%s is not a stream' % filename_or_entity.name)
        cls = (
                CompoundFileMiniStream
                if filename_or_entity.size < self._mini_size_limit else
                CompoundFileNormalStream)
        return cls(
                self, filename_or_entity._start_sector,
                filename_or_entity.size)

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
        # Note: when reading the master-FAT we deliberately disregard the
        # master-FAT sector count read from the header as implementations may
        # set this incorrectly. Instead, we scan for END_OF_CHAIN (or
        # FREE_SECTOR) in the DIFAT after each read and stop when we find it.
        # In order to avoid infinite loops (in the case of a stupid or
        # malicious file) we keep track of each sector we seek to and quit in
        # the event of a repeat
        self._master_fat = array(b'L')
        count = self._master_sector_count
        checked = 0
        sectors = set()

        # Special case: the first 109 entries are stored at the end of the file
        # header and the next sector of the master-FAT is stored in the header
        self._file.seek(COMPOUND_HEADER.size)
        self._master_fat.extend(
                st.unpack(b'<109L', self._file.read(109 * 4)))
        sector = self._master_first_sector
        if count == 0 and sector == FREE_SECTOR:
            warnings.warn(
                    'DIFAT extension pointer is FREE_SECTOR, assuming no '
                    'extension', CompoundFileWarning)
            sector = END_OF_CHAIN
        elif count == 0 and sector != END_OF_CHAIN:
            warnings.warn(
                    'DIFAT extension pointer with zero count',
                    CompoundFileWarning)
        elif count != 0 and sector == END_OF_CHAIN:
            warnings.warn(
                    'DIFAT chained from header, or incorrect '
                    'count', CompoundFileWarning)
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
                elif self._max_sector < value <= MAX_NORMAL_SECTOR:
                    warnings.warn(
                            'sector in DIFAT chain beyond file end '
                            '(%d)' % value, CompoundFileWarning)
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
            # allocation when reading the normal-FAT. If the normal-FAT alone
            # would exceed 100Mb of RAM, raise an error
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
        # Again, when reading the FAT we deliberately disregard the normal-FAT
        # sector count from the header as some implementations get it wrong.
        # Instead, we just read the sectors that the master-FAT chain tells us
        # to (no need to check for loops or invalid sectors here though - the
        # _load_master_fat method takes of those). After reading the normal-FAT
        # we check the master-FAT and normal-FAT sectors are marked correctly.
        self._normal_fat = array(b'L')
        # XXX This is the major cost at the moment - reading the fragmented
        # sectors of the FAT into an array. Perhaps look at optimizing reads
        # of contiguous sectors? Or make the array lazy-read whenever a block
        # needs filling?
        for sector in self._master_fat:
            self._seek_sector(sector)
            self._normal_fat.extend(
                    self._normal_sector_format.unpack(
                        self._file.read(self._normal_sector_format.size)))

        # The following simply verifies that all normal-FAT and master-FAT
        # sectors are marked appropriately in the normal-FAT
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

    def _load_mini_fat(self):
        # Guard against malicious files which could cause excessive memory
        # allocation when reading the mini-FAT. If the mini-FAT alone
        # would exceed 100Mb of RAM, raise an error
        if self._mini_sector_count * self._normal_sector_size > 100*1024*1024:
            raise CompoundFileError(
                    'excessively large mini-FAT (malicious file?)')
        self._mini_fat = array(b'L')

        # Construction of the stream below will construct the list of sectors
        # the mini-FAT occupies, and will constrain the length to the declared
        # mini-FAT sector count, or the number of occupied sectors (whichever
        # is shorter)
        if self._mini_first_sector == FREE_SECTOR:
            warnings.warn(
                    'mini FAT first sector set to FREE_SECTOR',
                    CompoundFileWarning)
            self._mini_first_sector = END_OF_CHAIN
        elif self._max_sector < self._mini_first_sector <= MAX_NORMAL_SECTOR:
            warnings.warn(
                    'mini FAT first sector beyond file end '
                    '(%d)' % self._mini_first_sector, CompoundFileWarning)
            self._mini_first_sector = END_OF_CHAIN
        if self._mini_first_sector != END_OF_CHAIN:
            with CompoundFileNormalStream(
                    self, self._mini_first_sector,
                    self._mini_sector_count * self._normal_sector_size) as stream:
                for i in range(stream._length // self._normal_sector_size):
                    self._mini_fat.extend(
                            self._normal_sector_format.unpack(
                                stream.read(self._normal_sector_format.size)))

    def _load_directory(self):
        # When reading the directory we don't attempt to accurately reconstruct
        # the red-black tree, partially because some implementations don't
        # write a correct red-black tree and partially because it doesn't
        # matter for users of the library. Instead we simply read the whole
        # stream of directory entries and construct a hierarchy of
        # CompoundFileEntity objects from this.
        #
        # In older compound files we have no idea how many entries are actually
        # in the directory, so we calculate an upper bound from the directory
        # stream's length
        stream = CompoundFileNormalStream(self, self._dir_first_sector)
        entries = [
                CompoundFileEntity(self, stream, index)
                for index in range(stream._length // DIR_HEADER.size)
                ]
        self.root = entries[0]
        self.root._build_tree(entries)

    def __len__(self):
        return len(self.root)

    def __getitem__(self, key):
        return self.root[key]

    def __contains__(self, key):
        return key in self.root


class CompoundFileEntity(object):
    """
    Represents an entity in an OLE Compound Document.

    An entity in an OLE Compound Document can be a "stream" (analogous to a
    file in a file-system) which has a :attr:`size` and can be opened by a call
    to the parent object's :meth:`~CompoundFileReader.open` method.
    Alternatively, it can be a "storage" (analogous to a directory in a
    file-system), which has no size but has :attr:`created` and
    :attr:`modified` time-stamps, and can contain other streams and storages.

    If the entity is a storage, it will act as an iterable read-only sequence,
    indexable by ordinal or by name, and compatible with the ``in`` operator
    and built-in :func:`len` function.

    .. attribute:: created

        For storage entities (where :attr:`isdir` is ``True``), this returns
        the creation date of the storage. Returns ``None`` for stream entities.

    .. attribute:: isdir

        Returns True if this is a storage entity which can contain other
        entities.

    .. attribute:: isfile

        Returns True if this is a stream entity which can be opened.

    .. attribute:: modified

        For storage entities (where :attr:`isdir` is True), this returns the
        last modification date of the storage. Returns ``None`` for stream
        entities.

    .. attribute:: name

        Returns the name of entity. This can be up to 31 characters long and
        may contain any character representable in UTF-16 except the NULL
        character. Names are considered case-insensitive for comparison
        purposes.

    .. attribute:: size

        For stream entities (where :attr:`isfile` is ``True``), this returns
        the number of bytes occupied by the stream. Returns 0 for storage
        entities.
    """

    def __init__(self, parent, stream, index):
        super(CompoundFileEntity, self).__init__()
        self._index = index
        self._children = None
        (
            name,
            name_len,
            self._entry_type,
            self._entry_color,
            self._left_index,
            self._right_index,
            self._child_index,
            self.uuid,
            user_flags,
            created,
            modified,
            self._start_sector,
            size_low,
            size_high,
        ) = DIR_HEADER.unpack(stream.read(DIR_HEADER.size))
        self.name = name.decode('utf-16le')
        try:
            self.name = self.name[:self.name.index('\0')]
        except ValueError:
            self._check(False, 'missing NULL terminator in name')
            self.name = self.name[:name_len]
        if index == 0:
            self._check(self._entry_type == DIR_ROOT, 'invalid type')
            self._entry_type = DIR_ROOT
        elif not self._entry_type in (DIR_STREAM, DIR_STORAGE, DIR_INVALID):
                self._check(False, 'invalid type')
                self._entry_type = DIR_INVALID
        if self._entry_type == DIR_INVALID:
            self._check(self.name == '', 'non-empty name')
            self._check(name_len == 0, 'invalid name length (%d)' % name_len)
            self._check(user_flags == 0, 'non-zero user flags')
        else:
            # Name length is in bytes, including NULL terminator ... for a
            # unicode encoded name ... *headdesk*
            self._check(
                    (len(self.name) + 1) * 2 == name_len,
                    'invalid name length (%d)' % name_len)
        if self._entry_type in (DIR_INVALID, DIR_ROOT):
            self._check(self._left_index == NO_STREAM, 'invalid left sibling')
            self._check(self._right_index == NO_STREAM, 'invalid right sibling')
            self._left_index = NO_STREAM
            self._right_index = NO_STREAM
        if self._entry_type in (DIR_INVALID, DIR_STREAM):
            self._check(self._child_index == NO_STREAM, 'invalid child index')
            self._check(self.uuid == b'\0' * 16, 'non-zero UUID')
            self._check(created == 0, 'non-zero creation timestamp')
            self._check(modified == 0, 'non-zero modification timestamp')
            self._child_index = NO_STREAM
            self.uuid = b'\0' * 16
            created = 0
            modified = 0
        if self._entry_type in (DIR_INVALID, DIR_STORAGE):
            self._check(self._start_sector == 0,
                    'non-zero start sector (%d)' % self._start_sector)
            self._check(size_low == 0,
                    'non-zero size low-bits (%d)' % size_low)
            self._check(size_high == 0,
                    'non-zero size high-bits (%d)' % size_high)
            self._start_sector = 0
            size_low = 0
            size_high = 0
        if parent._normal_sector_size == 512:
            # Surely this should be checking DLL version instead of sector
            # size?! But the spec does state sector size ...
            self._check(size_high == 0, 'invalid size in small sector file')
            self._check(size_low < 1<<31, 'size too large for small sector file')
            size_high = 0
        self.size = (size_high << 32) | size_low
        epoch = dt.datetime(1601, 1, 1)
        self.created = (
                epoch + dt.timedelta(microseconds=created // 10)
                if created != 0 else None)
        self.modified = (
                epoch + dt.timedelta(microseconds=created // 10)
                if modified != 0 else None)

    @property
    def isfile(self):
        return self._entry_type == DIR_STREAM

    @property
    def isdir(self):
        return self._entry_type in (DIR_STORAGE, DIR_ROOT)

    def _check(self, valid, message):
        if not valid:
            warnings.warn(
                    '%s in dir entry %d' % (message, self._index),
                    CompoundFileWarning)

    def _build_tree(self, entries):

        # XXX Need cycle detection in here - add a visited flag?
        def walk(node):
            if node._left_index != NO_STREAM:
                try:
                    walk(entries[node._left_index])
                except IndexError:
                    node._check(False, 'invalid left index')
            self._children.append(node)
            if node._right_index != NO_STREAM:
                try:
                    walk(entries[node._right_index])
                except IndexError:
                    node._check(False, 'invalid right index')
            if node._child_index != NO_STREAM:
                node._build_tree(entries)

        self._children = []
        try:
            walk(entries[self._child_index])
        except IndexError:
            self._check(False, 'invalid child index')

    def __len__(self):
        return len(self._children)

    def __iter__(self):
        return iter(self._children)

    def __contains__(self, name_or_obj):
        if isinstance(name_or_obj, bytes):
            name_or_obj = name_or_obj.decode(FILENAME_ENCODING)
        if isinstance(name_or_obj, str):
            try:
                self.__getitem__(name_or_obj)
                return True
            except KeyError:
                return False
        else:
            return name_or_obj in self._children

    def __getitem__(self, index_or_name):
        if isinstance(index_or_name, bytes):
            index_or_name = index_or_name.decode(FILENAME_ENCODING)
        if isinstance(index_or_name, str):
            name = index_or_name.lower()
            for item in self._children:
                if item.name.lower() == name:
                    return item
            raise KeyError(index_or_name)
        else:
            return self._children[index_or_name]

    def __repr__(self):
        return (
            "<CompoundFileEntity name='%s'>" % self.name
            if self.isfile else
            pformat([
                "<CompoundFileEntity dir='%s'>" % c.name
                if c.isdir else
                repr(c)
                for c in self._children
                ])
            if self.isdir else
            "<CompoundFileEntry ???>"
            )
