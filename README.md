# pdf2pptx
Convert your (Beamer) PDF slides to (Powerpoint) PPTX

So you don't like using Powerpoint and would rather use Latex/Beamer to make your slides,
however, you have a fancy Surface and you want to use the pen during the presentation? Here is my solution.

This script gets a PDF file as input and generates a Powerpoint PPTX file while preserving the format of the original PDF. Theoretically all PDF files, regardless of the original generator, can be converted to PPTX slides with this (not tested though).

Simply explained, I convert all the slides to high-quality image files first, and then push them into a Powerpoint project as a slide.

# How to run
* Execute `./pdf2pptx.sh test.pdf` to generate a `test.pdf.pptx` file  (replace `test.pdf` with your filename).
* By default the output powerpoint project is in the widescreen mode. If your slides are not for widescreen you can alternatively run `./pdf2pptx.sh test.pdf notwide` to generate a 4:3 standard PPTX project.

# Dependencies
* You need `convert` from [ImageMagick](http://www.imagemagick.org/script/binary-releases.php)
* `zip` and `sed`
* (Optional) `perl`, `python`, or `ruby` if you use a symlink to pdf2pptx.sh

If you're using *Linux* you *probably* already have all the above.

If you're using *OSX* you need to install **ImageMagick** and make sure `convert` is accessible from your Terminal.

If you're using *Windows* you can use *Cygwin*, but if you don't have it already, it is not recommended!
If you're using *Windows 10*, you can easily set up Ubuntu bash ([More Info](https://www.howtogeek.com/249966/how-to-install-and-use-the-linux-bash-shell-on-windows-10/)), install ImageMagick and then use the script natively.
Another solution for *Windows* users is to access a linux box (such as your university servers) to take care of the task.

# New issues with ImageMagick
ImageMagick no longer allows PDF to image conversion. If you get the following error on the test example:

```
Doing test.pdf
convert: not authorized `test.pdf' @ error/constitute.c/ReadImage/412.
convert: no images defined `./test.pdf.temp/slide.png' @ error/convert.c/ConvertImageCommand/3210.
Error with extraction
```

in `/etc/ImageMagick-6/policy.xml` or `/etc/ImageMagick/policy.xml`, change:

```XML
<policy domain="coder" rights="none" pattern="PDF" />
```

to

```XML
<policy domain="coder" rights="read" pattern="PDF" />
```

Now it should work. Note that modifying the policy file would require `root` privileges. If you do not have root access on your machine, you can alternatively compile and use an older version of ImageMagick.


# Acknowledgement
Thanks to [Melissa O'Neill](https://www.cs.hmc.edu/~oneill/freesoftware/pdftokeynote.html) for providing a Pdf2Keynote tool for mac which has motivated this small project!
