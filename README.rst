=============
compoundfiles
=============

|pypi| |rtd| |travis|

This package provides a library for reading Microsoft's `Compound File Binary`_
format (CFB), formerly known as `OLE Compound Documents`_, the `Advanced
Authoring Format`_ (AAF), or just plain old Microsoft Office files (the non-XML
sort). This format is also widely used with certain media systems and a number
of scientific applications (tomography and microscopy).

The code is pure Python and should run on any platform; it is compatible with
Python 2.7 (or above) and Python 3.2 (or above). The library has an emphasis
on rigour and performs numerous validity checks on opened files.  By default,
the library merely warns when it comes across non-fatal errors in source files
but this behaviour is configurable by developers through Python's ``warnings``
mechanisms.

Links
=====

* The code is licensed under the `MIT license`_
* The `source code`_ can be obtained from GitHub, which also hosts the `bug
  tracker`_
* The `documentation`_ (which includes installation instructions and
  quick-start examples) can be read on ReadTheDocs
* The `build status`_ can be observed on Travis CI

.. _documentation: http://compound-files.readthedocs.org/
.. _source code: https://github.com/waveform80/compoundfiles
.. _bug tracker: https://github.com/waveform80/compoundfiles/issues
.. _Compound File Binary: http://msdn.microsoft.com/en-gb/library/dd942138.aspx
.. _OLE Compound Documents: http://www.openoffice.org/sc/compdocfileformat.pdf
.. _Advanced Authoring Format: http://www.amwa.tv/downloads/specifications/aafcontainerspec-v1.0.1.pdf
.. _MIT license: http://opensource.org/licenses/MIT
.. _build status: https://travis-ci.org/waveform80/compoundfiles

.. |pypi| image:: https://pypip.in/version/compoundfiles/badge.svg
    :target: https://pypi.python.org/pypi/compoundfiles
    :alt: Latest release

.. |rtd| image:: https://readthedocs.org/projects/compound-files/badge/?version=latest
    :target: https://compound-files.readthedocs.org/
    :alt: Documentation status

.. |travis| image:: https://travis-ci.org/waveform80/compoundfiles.svg?branch=master
    :target: https://travis-ci.org/waveform80/compoundfiles
    :alt: Build status

