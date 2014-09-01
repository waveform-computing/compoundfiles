.. _changelog:

==========
Change log
==========


Release 0.3 (2014-09-01)
========================

* Added a comprehensive test suite and fixed several small bugs as a result
  (all to do with invalid file handling) (`#2`_)
* Added an mmap emulation to enable reading of massive files on 32-bit systems;
  the emulation is necessarily slower than "proper" mmap but that's the cost
  of staying on 32-bit! (`#6`_)
* Extended the warning and error hierarchy so that users of the library can
  fine tune exactly what warnings they want to consider errors (`#3`_)

.. _#2: https://github.com/waveform80/compoundfiles/issues/2
.. _#3: https://github.com/waveform80/compoundfiles/issues/3
.. _#6: https://github.com/waveform80/compoundfiles/issues/6


Release 0.2 (2014-04-23)
========================

* Fixed a nasty bug where opening multiple streams in a compound document would
  result in shared file pointer state (`#4`_)
* Fixed Python 3 compatibility - many thanks to Martin Panter for the bug
  report! (`#5`_)

.. _#4: https://github.com/waveform80/compoundfiles/issues/4
.. _#5: https://github.com/waveform80/compoundfiles/issues/5


Release 0.1 (2014-02-22)
========================

Initial release.
