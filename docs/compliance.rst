.. _compliance:

=====================
Compliance mechanisms
=====================

As noted in the `CFB`_ specification, the compound document format presents a
number of validation challenges. For example, maliciously constructed files
might include circular references in their FAT table, leading a naive reader
into an infinite loop, or they may allocate a large number of DIFAT sectors
hoping to cause resource exhaustion when the reader goes to allocate memory for
reading the FAT.

The compoundfiles library goes to some lengths to detect erroneous structures
(whether malicious in intent or otherwise) and work around them where possible.
Some issues are considered fatal and will always raise an exception (circular
chains in the FAT are an example of this). Other issues are considered
non-fatal and will raise a warning (unusual sector sizes are an example of
this). Python :mod:`warnings` are a special sort of exception with particularly
flexible handling.

With Python's defaults, a specific warning will print a message to the console
the first time it is encountered and will then do nothing if it's encountered
again (this avoids spamming the console in case a warning is raised in a tight
loop). With some simple code, you can specify alternative behaviours: warnings
can be raised as full-blown exceptions, or suppressed entirely. The
compoundfiles library defines a large hierarchy of errors and warnings to
enable developers to finetune their handling.

For example, consider a developer writing an application for working with
computed tomography (CT) scans. The files produced by the scanner's software
are compound documents, but they use an unusual sector size. Whenever the
developer's Python script opens a file the following warning is emitted::

    /usr/lib/pyshared/python2.7/compoundfiles/compoundfiles/reader.py:275: CompoundFileSectorSizeWarning: unexpected sector size in v3 file (1024)

Other than this, the script runs successfully. The developer decides the
warning is unimportant (after all there's nothing he can do about it given he
can't change the scanner's software) and wishes to suppress it entirely, so he
adds the following line to the top of his script::

    import warnings
    import compoundfiles as cf

    warnings.filterwarnings('ignore', category=cf.CompoundFileSectorSizeWarning)

Another developer is working on a file validation service. She wishes to use
the compoundfiles library to extract and examine the contents of such files.
For safety, she decides to treat any violation of the specification as an
error, so she adds the following line to the top of her script to tell Python
to convert all compound file warnings into exceptions::

    import warnings
    import compoundfiles as cf

    warnings.filterwarnings('error', category=cf.CompoundFileWarning)

The class hierarchies for compoundfiles warnings and errors is illustrated
below:

.. image:: warnings.*
    :align: center

.. image:: errors.*
    :align: center

To set filters on all warnings in the hierarchy, simply use the category
:exc:`~compoundfiles.CompoundFileWarning`. Otherwise, you can use intermediate
or leaf classes in the hierarchy for more specific filters. Likewise, when
catching exceptions you can target the root of the hierarchy
(:exc:`~compoundfiles.CompoundFileError`) to catch any error that the
compoundfiles library might raise, or a more specific class to deal with a
particular error.

.. _CFB: http://msdn.microsoft.com/en-gb/library/dd942138.aspx
