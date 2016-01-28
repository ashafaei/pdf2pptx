#!/bin/bash
# Alireza Shafaei - shafaei@cs.ubc.ca - Jan 2016

resolution=1024
density=300
makeWide=true

if [ $# -eq 0 ]; then
    echo "No arguments supplied!"
    echo "Usage: ./pdf2pptx.sh file.pdf"
    echo "			Generates file.pdf.pptx in widescreen format (by default)"
    echo "       ./pdf2pptx.sh file.pdf notwide"
    echo "			Generates file.pdf.pptx in 4:3 format"
    exit 1
fi

if [ $# -eq 2 ]; then
	if [ "$2" == "notwide" ]; then
		makeWide=false
	fi
fi

echo "Doing $1"
tempname="$1.temp"
if [ -d $tempname ]; then
	echo "Removing ${tempname}"
	rm -rf $tempname
fi

mkdir $tempname
convert -density $density -depth 8 -resize "x${resolution}" $1 ./$tempname/slide.png

if [ $? -eq 0 ]; then
	echo "Extraction succ!"
else
	echo "Error with extraction"
	exit
fi

pptname="$1.pptx.base"
fout="$1.pptx"
rm -rf $pptname
cp -r template $pptname

mkdir $pptname/ppt/media

cp ./$tempname/*.png "$pptname/ppt/media/"

function add_slide {
	pat='slide1\.xml\"\/>'
	id=$1
	id=$((id+8))
	entry='<Relationship Id=\"rId'$id'\" Type=\"http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slide\" Target=\"slides\/slide-'$1'\.xml"\/>'
	rep="${pat}${entry}"
	sed -i "s/${pat}/${rep}/g" ../_rels/presentation.xml.rels 

	pat='slide1\.xml\" ContentType=\"application\/vnd\.openxmlformats-officedocument\.presentationml\.slide+xml\"\/>'
	entry='<Override PartName=\"\/ppt\/slides\/slide-'$1'\.xml\" ContentType=\"application\/vnd\.openxmlformats-officedocument\.presentationml\.slide+xml\"\/>'
	rep="${pat}${entry}"
	sed -i "s/${pat}/${rep}/g" ../../\[Content_Types\].xml

	sid=$1
	sid=$((sid+256))
	pat='<p:sldIdLst>'
	entry='<p:sldId id=\"'$sid'\" r:id=\"rId'$id'\"\/>'
	rep="${pat}${entry}"
	sed -i "s/${pat}/${rep}/g" ../presentation.xml
}

function make_slide {
	cp ../slides/slide1.xml ../slides/slide-$1.xml
	cat ../slides/_rels/slide1.xml.rels | sed "s/image1\.JPG/slide-${slide}.png/g" > ../slides/_rels/slide-$1.xml.rels
	add_slide $1
}

pushd $pptname/ppt/media/
count=`ls -ltr | wc -l`
for (( slide=$count-2; slide>=0; slide-- ))
do
	echo "Processing "$slide
	make_slide $slide
done

if [ "$makeWide" = true ]; then
	pat='<p:sldSz cx=\"9144000\" cy=\"6858000\" type=\"screen4x3\"\/>'
	wscreen='<p:sldSz cy=\"6858000\" cx=\"12192000\"\/>'
	sed -i "s/${pat}/${wscreen}/g" ../presentation.xml
fi
popd

pushd $pptname
rm -rf ../$fout
zip -q -r ../$fout .
popd

rm -rf $pptname
rm -rf $tempname
