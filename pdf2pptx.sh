#!/bin/bash
# Alireza Shafaei - shafaei@cs.ubc.ca - Jan 2016

#colorspace="-depth 8"
colorspace="-colorspace sRGB -background white -alpha remove"
makeWide=true

usage="Usage: $(basename $0) [-n] [-w width] [-h height] [-r resolution] [-d dpi] [-o output] file.pdf\n\t\
-n is for 4:3 aspect ratio (not wide format)\n\t\
-w width and -h height should be used together and be in pixels\n\t\
-o output file, e.g. out.pptx (otherwise uses file.pdf name/directory)"

if [ $# -eq 0 ]; then
    echo "No arguments supplied!"
    echo -e "$usage"
    echo -e "\tGenerates file.pptx in widescreen format (by default)"
    echo -e "\t$(basename $0) -n file.pdf"
    echo -e "\tGenerates file.pptx in 4:3 format"
    exit 1
fi

while getopts 'nw:h:r:d:o:' OPTION; do
  case "$OPTION" in
    n) makeWide=false ;;
    w) width="$OPTARG"
    	W=$((width*10000)) ;;
    h) height="$OPTARG"
      H=$((height*10000)) ;;
    r) resolution="$OPTARG" ;;
	 d) density="$OPTARG" ;;
	 o) output="$OPTARG" ;;
    ?) echo -e $usage >&2
      exit 1 ;;
  esac
done

# Catch bad arguments (must have -w and -h, together)
if [[ -z $W ]]; then
	if ! [[ -z $H ]]; then
		echo -e $usage >&2
		exit 1
	fi
elif [[ -z $H ]]; then
	echo -e $usage >&2
	exit 1
fi

# Parse input file.pdf positional argument
INPUT=${@:$OPTIND+0}

# Set width W, height H, resolution, density (dpi), output
if ! [[ $makeWide = true ]]; then
	W=12192000
	H=6858000
elif ! [[ -n $W ]]; then
	W=9144000
	H=6858000
fi
if ! [[ -n $resolution ]]; then
	resolution=1024
fi
if ! [[ -n $density ]]; then
	density=300
fi
if ! [[ -n $output ]]; then
	output=$(basename "$INPUT")".pptx"
fi

# Remind us of the arguments we passed
echo "Input file: $INPUT"
if [[ -n $width ]]; then
	echo "width: $width"
fi
if [[ -n $height ]]; then
	echo "height: $height"
fi
if [[ -n $resolution ]]; then
	echo "resolution: $resolution"
fi
if [[ -n $density ]]; then
	echo "dpi: $density"
fi
echo "output: $output"


tempname="$INPUT.temp"
if [ -d "$tempname" ]; then
	# echo "Removing ${tempname}"
	rm -rf "$tempname"
fi

mkdir "$tempname"

# Set return code of piped command to first nonzero return code
set -o pipefail
n_pages=$(identify "$INPUT" | wc -l)
returncode=$?
if [ $returncode -ne 0 ]; then
   echo "Unable to count number of PDF pages, exiting"
   exit $returncode
fi
if [ $n_pages -eq 0 ]; then
   echo "Empty PDF (0 pages), exiting"
   exit 1
fi

for ((i=0; i<n_pages; i++))
do
    convert -density $density $colorspace -resize "x${resolution}" "$INPUT[$i]" "$tempname"/slide-$i.png
    returncode=$?
    if [ $returncode -ne 0 ]; then break; fi
done

if [ $returncode -eq 0 ]; then
	echo -e "\tExtraction success!"
else
	echo -e "\tError with extraction"
	exit $returncode
fi

if (which perl > /dev/null); then
	# https://stackoverflow.com/questions/1055671/how-can-i-get-the-behavior-of-gnus-readlink-f-on-a-mac#comment47931362_1115074
	mypath=$(perl -MCwd=abs_path -le '$file=shift; print abs_path -l $file? readlink($file): $file;' "$0")
elif (which python > /dev/null); then
	# https://stackoverflow.com/questions/1055671/how-can-i-get-the-behavior-of-gnus-readlink-f-on-a-mac#comment42284854_1115074
	mypath=$(python -c 'import os,sys; print(os.path.realpath(os.path.expanduser(sys.argv[1])))' "$0")
elif (which ruby > /dev/null); then
	mypath=$(ruby -e 'puts File.realpath(ARGV[0])' "$0")
else
	mypath="$0"
fi
mydir=$(dirname "$mypath")

pptname="$output.base"
fout=$(basename "$output")
rm -rf "$pptname"
cp -r "$mydir"/template "$pptname"

mkdir "$pptname"/ppt/media

cp "$tempname"/*.png "$pptname/ppt/media/"

function call_sed {
	if [ "$(uname -s)" == "Darwin" ]; then
		{ # try macOS sed
			sed -i "" "$@" 2>/dev/null # suppress error message
		} || { # catch
			sed -i "$@" # for macOS with GNU sed
		}
	else
		sed -i "$@"
	fi
}

function add_slide {
	pat='slide1\.xml\"\/>'
	id=$1
	id=$((id+8))
	entry='<Relationship Id=\"rId'$id'\" Type=\"http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slide\" Target=\"slides\/slide-'$1'\.xml"\/>'
	rep="${pat}${entry}"
	call_sed "s/${pat}/${rep}/g" ../_rels/presentation.xml.rels

	pat='slide1\.xml\" ContentType=\"application\/vnd\.openxmlformats-officedocument\.presentationml\.slide+xml\"\/>'
	entry='<Override PartName=\"\/ppt\/slides\/slide-'$1'\.xml\" ContentType=\"application\/vnd\.openxmlformats-officedocument\.presentationml\.slide+xml\"\/>'
	rep="${pat}${entry}"
	call_sed "s/${pat}/${rep}/g" ../../\[Content_Types\].xml

	sid="$1"
	sid=$((sid+256))
	pat='<p:sldIdLst>'
	entry='<p:sldId id=\"'$sid'\" r:id=\"rId'$id'\"\/>'
	rep="${pat}${entry}"
	call_sed "s/${pat}/${rep}/g" ../presentation.xml
}

function make_slide {
	cp ../slides/slide1.xml "../slides/slide-$1.xml"
	cat ../slides/_rels/slide1.xml.rels | sed "s/image1\.JPG/slide-${slide}.png/g" > "../slides/_rels/slide-${1}.xml.rels"
	add_slide "$1"
}

pushd "$pptname"/ppt/media/ &> /dev/null
count=`ls -ltr | wc -l`
for (( slide=$count-2; slide>=0; slide-- ))
do
	echo -ne "\tprocessing slide $slide ... "
	make_slide $slide
	echo -e "done."
done

if [[ "$makeWide" = true || (-n $W) ]]; then
	pat='<p:sldSz cx=\"9144000\" cy=\"6858000\" type=\"screen4x3\"\/>'
	wscreen="<p:sldSz cy=\"${H}\" cx=\"${W}\"\/>"
	call_sed "s/${pat}/${wscreen}/g" ../presentation.xml
fi
popd &> /dev/null

pushd "$pptname" &> /dev/null
rm -rf ../"$fout"
zip -q -r ../"$fout" .
popd &> /dev/null

echo "Success: created $output"

rm -rf "$pptname"
rm -rf "$tempname"
