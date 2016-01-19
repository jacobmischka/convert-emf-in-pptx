I wrote this because for some reason a faculty where I work uses some addon that gives him EMF images in his presentations.

These images are massive for no reason, so converting them to regular images makes the resulting presentation much smaller.

This filthily calls some shell commands to invoke other tools, so really this should be written in bash, but bash is confusing.

Requires `unoconv` and imagemagick's `convert` to be installed and in your path.

This was written for Linux, if you're on Windows I'm pretty sure you can convert to regular images directly with imagemagick. This script converts it to a PDF first using LibreOffice/OpenOffice via `unoconv`, because for some reason that works best on Linux based on my stackoverflow searches.

**Usage**: `python3 convert-emf-in-pptx <PPTX file>`
