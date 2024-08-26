# LibreOfficeKit

Wrapper around the C++ LibreOfficeKit library.

This is a modified and improved version of https://github.com/undeflife/libreoffice-rs aiming to be more complex and convert more
use cases, to use as a backend for an Office file format conversion server.


> [!IMPORTANT]
>
> LibreOffice has some broken behavior in the newer versions where the cleanup
> functions of Office cause segmentation faults, you must keep around a single 
> instance of Office around for the duration of your program