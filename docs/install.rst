.. _install:

============
Installation
============

The library has no dependencies in and of itself, and consists entirely of
Python code, so installation should be trivial on most platforms. A debian
packaged variant is available from the author's PPA for Ubuntu users, otherwise
installation from PyPI is recommended.


.. _ubuntu_install:

Ubuntu installation
===================

To install from the author's `PPA`_::

    $ sudo add-apt-repository ppa://waveform/ppa
    $ sudo apt-get update
    $ sudo apt-get install python-compound-files

To remove the installation::

    $ sudo apt-get remove python-compound-files


.. _pypi_install:

PyPI installation
=================

To install from `PyPI`_::

    $ sudo pip install compound_files

To upgrade the installation::

    $ sudo pip install -U compound_files

To remove the installation::

    $ sudo pip uninstall compound_files


Development installation
========================

If you wish to develop the library yourself, you are best off doing so within
a virtualenv with source checked out from `GitHub`_, like so::

    $ sudo apt-get install make git python-virtualenv exuberant-ctags
    $ virtualenv sandbox
    $ source sandbox/bin/activate
    $ git clone https://github.com/waveform80/compound_files.git
    $ cd compound_files
    $ make develop

The above instructions assume Ubuntu Linux, and use the included Makefile to
perform a development installation into the constructed virtualenv (as well
as constructing a tags file for easier navigation in vim/emacs). Please feel
free to extend this section with instructions for alternate platforms.


.. _PyPI: http://pypi.python.org/pypi/compound_files
.. _GitHub: https://github.com/waveform80/compound_files
.. _PPA: https://launchpad.net/~waveform/+archive/ppa
