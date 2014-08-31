=============
compoundfiles
=============

This package provides a library for reading Microsoft's `OLE Compound
Document`_ format, which also forms the basis of the `Advanced Authoring
Format`_ (AAF) published by Microsoft Corporation. It is compatible with
Python 2.7 (or above) and Python 3.2 (or above).

The code is pure Python and should run on any platform. The library has an
emphasis on rigour and performs numerous validity checks on opened files.  By
default, the library merely warnings when it comes across non-fatal errors in
source files but this behaviour is configurable by developers through Python's
``warnings`` mechanisms.

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
.. _OLE Compound Document: http://www.openoffice.org/sc/compdocfileformat.pdf
.. _Advanced Authoring Format: http://www.amwa.tv/downloads/specifications/aafcontainerspec-v1.0.1.pdf
.. _MIT license: http://opensource.org/licenses/MIT
.. _build status: https://travis-ci.org/waveform80/compoundfiles

