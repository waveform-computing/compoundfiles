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


import sys
import io
import mmap
import compoundfiles.mmap as fake_mmap
import pytest


@pytest.fixture(scope='session')
def source():
    return io.open('tests/mmap.dat', 'rb')

@pytest.fixture(params=(False, True))
def test_map(request, source):
    # Run all tests against a real memory map, and our fake memory map, both
    # covering the same underlying file (which just contains a..z), to ensure
    # that the emulated behaviour matches the real implementation
    if request.param:
        return fake_mmap.FakeMemoryMap(source)
    else:
        return mmap.mmap(source.fileno(), 0, access=mmap.ACCESS_READ)


def test_read_only(test_map):
    with pytest.raises(TypeError):
        test_map.write(b'foo')
    with pytest.raises(TypeError):
        test_map.write_byte(b'f')
    with pytest.raises(TypeError):
        test_map[:6] = b'foobar'
    with pytest.raises(TypeError):
        test_map.move(5, 10, 4)
    with pytest.raises(TypeError):
        test_map.resize(1000)
    # This doesn't raise an error ... even with a read-only mmap!
    test_map.flush()

def test_len(test_map):
    assert len(test_map) == 26

@pytest.mark.skipif(sys.version_info[0] == 3,
        reason="py2 mmap is an array of chars")
def test_indexing_v2(test_map):
    assert test_map[0] == b'a'
    assert test_map[4] == b'e'
    assert test_map[-1] == b'z'
    assert test_map[-3] == b'x'

@pytest.mark.skipif(sys.version_info[0] == 2,
        reason="py3 mmap is an array of ints")
def test_indexing_v3(test_map):
    assert test_map[0] == ord(b'a')
    assert test_map[4] == ord(b'e')
    assert test_map[-1] == ord(b'z')
    assert test_map[-3] == ord(b'x')

def test_indexing(test_map):
    with pytest.raises(IndexError):
        test_map[30]
    with pytest.raises(IndexError):
        test_map[-30]

def test_slicing(test_map):
    assert test_map[:0] == b''
    assert test_map[:1] == b'a'
    assert test_map[:5] == b'abcde'
    assert test_map[1:5] == b'bcde'
    assert test_map[1:5:2] == b'bd'
    assert test_map[2:5:3] == b'c'
    assert test_map[:-3] == b'abcdefghijklmnopqrstuvw'
    assert test_map[0:] == b'abcdefghijklmnopqrstuvwxyz'
    assert test_map[-3:] == b'xyz'
    assert test_map[-3::1] == b'xyz'
    assert test_map[-3::2] == b'xz'
    assert test_map[-5:-1:2] == b'vx'

def test_negative_slicing(test_map):
    assert test_map[::-1] == b'zyxwvutsrqponmlkjihgfedcba'
    assert test_map[::-2] == b'zxvtrpnljhfdb'
    assert test_map[-2::-2] == b'ywusqomkigeca'
    assert test_map[5:10:-1] == b''
    assert test_map[10:5:-1] == b'kjihg'

def test_bad_slice(test_map):
    with pytest.raises(ValueError):
        test_map[5:10:0]

def test_contains(test_map):
    # See comments in FakeMemoryMap.__contains__ to understand this!
    assert b'a' in test_map
    assert b'abc' not in test_map
    assert b'd' in test_map
    assert b'def' not in test_map
    assert b'vwxyz' not in test_map
    assert b'blah' not in test_map

def test_find(test_map):
    assert test_map.find(b'abc') == 0
    assert test_map.find(b'abc', 5) == -1
    assert test_map.find(b'xyz') == 23
    assert test_map.find(b'xyz', 0, -1) == -1
    assert test_map.find(b'foobar') == -1

def test_rfind(test_map):
    assert test_map.rfind(b'abc') == 0
    assert test_map.rfind(b'abc', 5) == -1
    assert test_map.rfind(b'xyz') == 23
    assert test_map.rfind(b'xyz', 0, -1) == -1
    assert test_map.rfind(b'foobar') == -1

@pytest.mark.skipif(sys.version_info[0] == 3,
        reason="py2 read_byte returns bytes")
def test_seek_n_read_v2(test_map):
    test_map.seek(-3, io.SEEK_END)
    assert test_map.read_byte() == b'x'
    assert test_map.read_byte() == b'y'
    assert test_map.read_byte() == b'z'

@pytest.mark.skipif(sys.version_info[0] == 2,
        reason="py3 read_byte returns an int")
def test_seek_n_read_v3(test_map):
    test_map.seek(-3, io.SEEK_END)
    assert test_map.read_byte() == ord(b'x')
    assert test_map.read_byte() == ord(b'y')
    assert test_map.read_byte() == ord(b'z')

def test_seek_n_read(test_map):
    test_map.seek(0)
    assert test_map.read(26) == b'abcdefghijklmnopqrstuvwxyz'
    test_map.seek(0)
    assert test_map.read(5) == b'abcde'
    assert test_map.read(2) == b'fg'
    test_map.seek(-3, io.SEEK_END)
    assert test_map.read(-1) == b'xyz'
    assert test_map.read(-1) == b''
    test_map.seek(0, io.SEEK_SET)
    assert test_map.readline() == b'abcdefghijklmnopqrstuvwxyz'

def test_seek_n_tell(test_map):
    test_map.seek(0)
    assert test_map.tell() == 0
    test_map.seek(0, io.SEEK_END)
    assert test_map.tell() == 26
    test_map.seek(-3, io.SEEK_CUR)
    assert test_map.tell() == 23
    test_map.seek(0, io.SEEK_SET)
    assert test_map.tell() == 0
