.. _quickstart:

===========
Quick Start
===========

Import the library and open a compound document file::

    >>> import compoundfiles
    >>> doc = compoundfiles.CompoundFileReader('foo.txm')
    compoundfiles/__init__.py:606: CompoundFileWarning: DIFAT terminated by FREE_SECTOR
      CompoundFileWarning)

When opening the file you may see various warnings printed to the console (as
in the example above). The library performs numerous checks for compliance with
the specification, but many implementations don't *quite* conform. By default
the warnings are simply printed and can be ignored, but via Python's
:mod:`warnings` system you can either silence the warnings entirely or convert
them into full blown exceptions.

You can list the contents of the compound file via the
:attr:`~compoundfiles.CompoundFileReader.root` attribute which can be treated
like a dictionary::

    >>> doc.root
    ["<CompoundFileEntity name='Version'>",
     u"<CompoundFileEntity dir='AutoRecon'>",
     u"<CompoundFileEntity dir='ImageInfo'>",
     u"<CompoundFileEntity dir='ImageData1'>",
     u"<CompoundFileEntity dir='ImageData2'>",
     u"<CompoundFileEntity dir='ImageData3'>",
     u"<CompoundFileEntity dir='ImageData4'>",
     u"<CompoundFileEntity dir='ImageData5'>",
     u"<CompoundFileEntity dir='ImageData6'>",
     u"<CompoundFileEntity dir='ImageData7'>",
     u"<CompoundFileEntity dir='ImageData8'>",
     u"<CompoundFileEntity dir='ImageData9'>",
     u"<CompoundFileEntity dir='SampleInfo'>",
     u"<CompoundFileEntity dir='ImageData10'>",
     u"<CompoundFileEntity dir='ImageData11'>",
     u"<CompoundFileEntity dir='ImageData12'>",
     u"<CompoundFileEntity dir='ImageData13'>",
     u"<CompoundFileEntity dir='ImageData14'>",
     u"<CompoundFileEntity dir='ImageData15'>",
     u"<CompoundFileEntity dir='ImageData16'>",
     u"<CompoundFileEntity dir='ImageData17'>",
     u"<CompoundFileEntity dir='ImageData18'>",
     u"<CompoundFileEntity dir='ImageData19'>",
     u"<CompoundFileEntity dir='ImageData20'>"]
    >>> doc.root['ImageInfo']
    ["<CompoundFileEntity name='Date'>",
     "<CompoundFileEntity name='Angles'>",
     "<CompoundFileEntity name='Energy'>",
     "<CompoundFileEntity name='Current'>",
     "<CompoundFileEntity name='Voltage'>",
     "<CompoundFileEntity name='CameraNo'>",
     "<CompoundFileEntity name='DataType'>",
     "<CompoundFileEntity name='ExpTimes'>",
     "<CompoundFileEntity name='PixelSize'>",
     "<CompoundFileEntity name='XPosition'>",
     "<CompoundFileEntity name='YPosition'>",
     "<CompoundFileEntity name='ZPosition'>",
     "<CompoundFileEntity name='ImageWidth'>",
     "<CompoundFileEntity name='MosiacMode'>",
     "<CompoundFileEntity name='MosiacRows'>",
     "<CompoundFileEntity name='NoOfImages'>",
     "<CompoundFileEntity name='FocusTarget'>",
     "<CompoundFileEntity name='ImageHeight'>",
     "<CompoundFileEntity name='ImagesTaken'>",
     "<CompoundFileEntity name='ReadoutFreq'>",
     "<CompoundFileEntity name='ReadOutTime'>",
     "<CompoundFileEntity name='Temperature'>",
     "<CompoundFileEntity name='DtoRADistance'>",
     "<CompoundFileEntity name='HorizontalBin'>",
     "<CompoundFileEntity name='MosiacColumns'>",
     "<CompoundFileEntity name='NanoImageMode'>",
     "<CompoundFileEntity name='ObjectiveName'>",
     "<CompoundFileEntity name='ReferenceFile'>",
     "<CompoundFileEntity name='StoRADistance'>",
     "<CompoundFileEntity name='VerticalalBin'>",
     "<CompoundFileEntity name='BackgroundFile'>",
     "<CompoundFileEntity name='MosaicFastAxis'>",
     "<CompoundFileEntity name='MosaicSlowAxis'>",
     "<CompoundFileEntity name='AcquisitionMode'>",
     "<CompoundFileEntity name='TubelensPosition'>",
     "<CompoundFileEntity name='IonChamberCurrent'>",
     "<CompoundFileEntity name='NoOfImagesAveraged'>",
     "<CompoundFileEntity name='OpticalMagnification'>",
     "<CompoundFileEntity name='AbsorptionScaleFactor'>",
     "<CompoundFileEntity name='AbsorptionScaleOffset'>",
     "<CompoundFileEntity name='TransmissionScaleFactor'>",
     "<CompoundFileEntity name='OriginalDataRefCorrected'>",
     "<CompoundFileEntity name='RefTypeToApplyIfAvailable'>"]

Use the :meth:`~compoundfiles.CompoundFileReader.open` method with a
:class:`~compoundfiles.CompoundFileEntity`, or with a name that leads to one,
to obtain a file-like object which can read the stream's content::

    >>> doc.open('AutoRecon/BeamHardeningFilename').read()
    'Standard Beam Hardening Correction\x00'
    >>> f = doc.open(doc.root['ImageData1']['Image1'])
    >>> f.tell()
    0
    >>> import os
    >>> f.seek(0, os.SEEK_END)
    8103456
    >>> f.seek(0)
    0
    >>> f.read(10)
    '\xb3\x0c\xb3\x0c\xb3\x0c\xb3\x0c\xb3\x0c'
    >>> f.close()

You can also use entities as iterators, and the context manager protocol is
supported for file and stream opening::

    >>> with compoundfiles.CompoundFileReader('foo.txm') as doc:
    ...     for entry in doc.root['AutoRecon']:
    ...         if entry.isfile:
    ...             with doc.open(entry) as stream:
    ...                 print(repr(stream.read()))
    ... 
    '"\x00>C'
    '\x81\x02SG'
    '\x1830\xc5'
    '\x00\x00\x00\x00'
    '\x9a\x99\x99?'
    '\xcf.AD'
    '(\x1c\x1cF'
    ',E\xd6\xc3'
    '\x02\x00\x00\x00'
    '\x01\x00\x00\x00'
    '\x00\x00\x00\x00'
    '\x00\x00\x00\x00'
    '\xd4\xfe\x9fA'
    '\xd1\x07\x00\x00'
    '\x05\x00\x00\x00'
    '\x00\x00\x00\x00'
    'p\xff\x1fB'
    '\x00\x00\x00\x00'
    '\x02\x00\x00\x00'
    '\x01\x00\x00\x00'
    'Standard Beam Hardening Correction\x00'
    '\x00'

