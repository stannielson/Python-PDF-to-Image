"""module that performs simple conversion of PDF files to multiple formats
of image files, along with the capability to retrieve information about PDF
files, in a Windows-based environment
            
Author(s):  Stanton K. Nielson
Date:       January 17, 2023
Version:    1.0


NOTE: This module relies on either the installation of Poppler or reference
to Poppler binaries to function. To download Poppler for Windows use, refer
to:

    https://github.com/oschwartz10612/poppler-windows/releases/


-------------------------------------------------------------------------------
This is free and unencumbered software released into the public domain.

Anyone is free to copy, modify, publish, use, compile, sell, or
distribute this software, either in source code form or as a compiled
binary, for any purpose, commercial or non-commercial, and by any
means.

In jurisdictions that recognize copyright laws, the author or authors
of this software dedicate any and all copyright interest in the
software to the public domain. We make this dedication for the benefit
of the public at large and to the detriment of our heirs and
successors. We intend this dedication to be an overt act of
relinquishment in perpetuity of all present and future rights to this
software under copyright law.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
OTHER DEALINGS IN THE SOFTWARE.

For more information, please refer to <http://unlicense.org/>
-------------------------------------------------------------------------------
"""


import os, platform, subprocess


__all__ = 'convert', 'pdfinfo'

PAGEKEY = 'Pages'
EXT = dict(tiff='tif', png='png', jpeg='jpg', jpegcmyk='jpg')
TIFFCOMPRESSION = 'packbits', 'jpeg', 'lzw', 'deflate'
COMPSWITCH = '-tiffcompression'
ADDLSWITCHES = ['-aa', 'yes', '-aaVector', 'yes', '-singlefile', '-f',
                '{filepage}', '-l', '{filepage}']


def convert(source_path: str, output_path_prefix: str,
            img_format: str = 'tiff', dpi: int = 300,
            user_password: str = None, owner_password: str = None,
            grayscale: bool = False, tiff_compression: str = None,
            page_num: int = False, page_num_offset: int = None,
            page_num_zfill: int = None, poppler_bin_path: str = None) -> list:
    """Converts a PDF file to an image or series of images and returns the
    paths to converted files as a list object

    Parameters:
    
        source_path: str
            The full path to the source PDF file.

        output_path_prefix: str
            The prefix for output files, including the output directory and
            the first part of the desired filename. Adding a '{page}' string
            anywhere in the filename will set where the number will appear for
            page numbering, including for a single page if page_num is set to
            True. If the prefix contains a file extension, it will be stripped
            from the path and replaced.

        img_format (optional): str
            The format of output images. Options include 'png', 'jpeg',
            'jpegcmyk', or 'tiff'. DEFAULT: 'tiff'

        dpi (optional): int
            The DPI of output images. DEFAULT: 300

        user_password (optional): str
            The user password for the PDF file. DEFAULT: None

        owner_password (optional): str
            The owner password for the PDF file. DEFAULT: None

        grayscale (optional): bool
            Specifies if images will be produced in grayscale. DEFAULT: False

        tiff_compression (optional): str
            Specifies the compression scheme for the output TIFF format.
            Options include 'packbits', 'jpeg', 'lzw', 'deflate', or None. The
            compression schemes offer the following:

                LZW:        Lossless compression without data artifacts, but
                            with larger filesize; standard for TIFF
                PACKBITS:   Lossless compression with high application
                            compatibility, but with larger filesize
                JPEG:       Lower filesize, but lossy compression with
                            limitations in use
                DEFLATE:    Lossless compression, but with larger filesize
        
            DEFAULT: None

        page_num (optional): bool
            Specifies if page numbering should be used, where source PDF would
            produce multiple output images or numbering is otherwise desired.
            If multiple output images will occur, numbering will be applied
            regardless. DEFAULT: False

        page_num_offset (optional): int
            Specifies the offset for page numbering, where the offset can be
            any positive integer or -1 (which starts a page or series of
            pages at zero). DEFAULT: None
            
        page_num_zfill (optional): int
            Specifies the zero-character padding for page numbers. If not
            specified, no padding will occur. DEFAULT: None

        poppler_bin_path (optional): str
            The full path to the Poppler binary folder (for use if the path to
            Poppler binaries is not in the PATH environmental variable).
            DEFAULT: None
    """
    converted = list()
    source_path = _getquotepath(source_path)
    source_info = pdfinfo(source_path, user_password, owner_password,
                          poppler_bin_path=poppler_bin_path)
    pages = source_info[PAGEKEY]
    output_path = _getquotepath(_stripextension(output_path_prefix))
    params = [_getcommandpath('pdftoppm', poppler_bin_path)]
    params.extend(['-r', str(dpi)])
    if user_password: params.extend(['-upw', user_password])
    if owner_password: params.extend(['-opw', owner_password])
    if img_format in EXT:
        params.append('-{}'.format(img_format))
        ext = EXT[img_format]
    else: ext = 'ppm'
    if img_format == 'tiff' and tiff_compression in TIFFCOMPRESSION:
        params.extend([COMPSWITCH, tiff_compression])
    elif img_format == 'tiff': params.extend([COMPSWITCH, 'none'])
    if grayscale: params.append('-gray')
    params.extend(ADDLSWITCHES)
    params.extend([source_path, output_path])
    for index in range(pages):
        page = index + 1
        if page_num_offset and page_num_offset >= -1: page += page_num_offset
        page = str(page).zfill(page_num_zfill) if page_num_zfill else str(page)
        fill = dict(filepage=index + 1)
        if pages > 1 or page_num: fill['page'] = page
        else: fill['page'] = ''
        command = ' '.join(params).format(**fill)
        process = _getprocess(command, poppler_bin_path)
        converted.append(
            '.'.join((output_path.format(**fill).strip('"'), ext)))
    return converted


def pdfinfo(source_path: str, user_password: str = None,
            owner_password: str = None, raw_dates: bool = False,
            timeout: int = None, poppler_bin_path: str = None) -> dict:
    """Returns the information related to a PDF file as a dictionary object

    Parameters:
    
        source_path: str
            The full path to the source PDF file.

        user_password (optional): str
            The user password for the PDF file. DEFAULT: None

        owner_password (optional): str
            The owner password for the PDF file. DEFAULT: None

        raw_dates (optional): bool
            Specifies if the undecoded data strings from the PDF are included.
            DEFAULT: False

        timeout (optional): int
            Specifies the timeout limit in seconds. DEFAULT: None

        poppler_bin_path (optional): str
            The full path to the Poppler binary folder (for use if the path to
            Poppler binaries is not in the PATH environmental variable).
            DEFAULT: None
    """
    params = [_getcommandpath('pdfinfo', poppler_bin_path),
              _getquotepath(source_path)]
    switches = '-upw', '-opw', '-rawdates'
    for switch, arg in zip(
        switches, (user_password, owner_password, raw_dates)):
        if arg and arg == str(arg): params.extend([switch, arg])
        elif arg: params.append(switch)
    command = ' '.join(params)
    process = _getprocess(command, poppler_bin_path)
    try: data, errors = process.communicate(timeout=timeout)
    except subprocess.TimeoutExpired: process.kill()
    info = dict((i.split(':')[0].strip(), ':'.join(i.split(':')[1:]).strip())
                for i in data.decode('utf8', 'ignore').split('\n')
                if i.strip())
    if info.get(PAGEKEY): info[PAGEKEY] = int(info[PAGEKEY])
    else: raise Exception('Unable to retrieve PDF pages')
    return info


def _getcommandpath(name: str, poppler_bin_path: str=None) -> str:
    """Internal function to return a command path for an executable within
    the Poppler binary directory path, including the binary directory
    in the command path if specified
    """
    name = '.'.join((name, 'exe'))
    if poppler_bin_path:
        return _getquotepath(os.path.join(poppler_bin_path, name))
    return name


def _getprocess(command, poppler_bin_path: str=None):
    """Internal function that returns an opened process that functions
    quiety (i.e., does not create a command line window in execution)
    """
    environs = os.environ.copy()
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    process = subprocess.Popen(command, env=environs, stdout=subprocess.PIPE,
                               stderr=subprocess.PIPE, startupinfo=startupinfo)
    return process


def _stripextension(path):
    """Internal function to strip the extension from a filename in a path"""
    folder, file = os.path.dirname(path), os.path.basename(path)
    name, ext = os.path.splitext(file)
    if '{page}' in ext: name += ext
    if '{page}' not in name: name += {page}
    return os.path.join(folder, name)


def _getquotepath(path): return '"{}"'.format(path.strip('"'))
