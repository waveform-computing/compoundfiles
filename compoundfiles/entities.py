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


import warnings
import datetime as dt
from pprint import pformat

from compoundfiles.errors import CompoundFileError, CompoundFileWarning
from compoundfiles.const import (
    NO_STREAM,
    DIR_INVALID,
    DIR_STORAGE,
    DIR_STREAM,
    DIR_ROOT,
    DIR_HEADER,
    FILENAME_ENCODING,
    )


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
            node._build_tree(entries)

        if self.isdir:
            self._children = []
            try:
                walk(entries[self._child_index])
            except IndexError:
                if self._child_index != NO_STREAM:
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

