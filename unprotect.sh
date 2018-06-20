#!/bin/bash

#
# MIT License
#
# Copyright (c) 2018 Ryan Smith
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

#
# Attributions:
# Temporary directory code inspired from:
#	https://stackoverflow.com/questions/4632028/how-to-create-a-temporary-directory
# Filename and extension extraction inspired from:
#	https://stackoverflow.com/questions/965053/extract-filename-and-extension-in-bash/965072#965072
# Binary sed editing inspired from:
#	http://everydaywithlinux.blogspot.com/2012/11/patch-strings-in-binary-files-with-sed.html
# vbaProject.bin details inspired from:
#	https://www.thegrideon.com/vba-internals.html and http://lbeliarl.blogspot.com/
# workbookProtection and sheetProtection details inspired from:
#	http://datapigtechnologies.com/blog/index.php/hack-into-a-protected-excel-2007-or-2010-workbook/comment-page-3/
#

if [ "$#" -ne 1 ]; then
	echo "Usage: $0 <workbook filename .xlsx or .xlsm>" >&2
	exit 1
fi

if ! [ -f $1 ]; then
	echo "Errror: Input file $1 does not exist." >&2
	exit 1
fi

if ! [ "`file -b $1`" = "Microsoft OOXML" ]; then
	echo "Error: Input file $1 does not appear to be an XML file (.xlsx or .xlsm)." >&2
	exit 1
fi

if ! [ -x "$(command -v mktemp)" ]; then
	echo 'Error: mktemp is not installed.' >&2
	exit 1
fi

if ! [ -x "$(command -v unzip)" ]; then
	echo 'Error: unzip is not installed.' >&2
	exit 1
fi

if ! [ -x "$(command -v zip)" ]; then
	echo 'Error: zip is not installed.' >&2
	exit 1
fi

if ! [ -x "$(command -v hexdump)" ]; then
	echo 'Error: hexdump is not installed.' >&2
	exit 1
fi

if ! [ -x "$(command -v xxd)" ]; then
	echo 'Error: xxd is not installed.' >&2
	exit 1
fi

FILENAME=$(basename -- "$1")
EXTENSION="${FILENAME##*.}"
FILENAME="${FILENAME%.*}"
DIR="$( cd "$( dirname -- "$1" )" && pwd )"
WORK_DIR=`mktemp -d`

echo 'Creating temporary directory ...'

# check if tmp dir was created
if [[ ! "$WORK_DIR" || ! -d "$WORK_DIR" ]]; then
	echo "Could not create temp dir." >&2
	exit 1
fi

# deletes the temp directory
function cleanup {
	echo "Deleting temp working directory $WORK_DIR ..."
	rm -rf "$WORK_DIR"
	echo "Done!"
}

# register the cleanup function to be called on the EXIT signal
trap cleanup EXIT

echo "Extracting workbook files from ${DIR}/${FILENAME}${EXTENSION} ..."
unzip -q $1 -d $WORK_DIR

echo 'Unprotecting workbook ...'
sed -i 's/<workbookProtection[^>]*>//g' $WORK_DIR/xl/workbook.xml

echo 'Making all worksheets visible ...'
sed -i 's/state="hidden" //g' $WORK_DIR/xl/workbook.xml
sed -i 's/state="veryHidden" //g' $WORK_DIR/xl/workbook.xml

if [ -f $WORK_DIR/xl/vbaProject.bin ] ; then
	echo 'Removing VBA password ...'
	hexdump -ve '1/1 "%.2X"' $WORK_DIR/xl/vbaProject.bin | sed 's/49443D227B.*7D220D0A/49443D227B46363035383546332D323538332D343843452D414237342D4346433834434337354338447D220D0A/' | sed 's/434D473D22.*220D0A4450423D22.*220D0A47433D22.*220D0A0D0A/434D473D2230363034313644353136353744333542443335424433354244333542220000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000D0A4450423D2243374335443731343539313431423135314231353142220000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000D0A47433D2238383841393839423939394239393634220000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000D0A0D0A/g' | xxd -r -p > $WORK_DIR/xl/vbaProject.bin.tmp
	chmod --reference $WORK_DIR/xl/vbaProject.bin $WORK_DIR/xl/vbaProject.bin.tmp
	mv $WORK_DIR/xl/vbaProject.bin.tmp $WORK_DIR/xl/vbaProject.bin
fi

for f in $WORK_DIR/xl/worksheets/*.xml
do
	echo "Unprotecting worksheet $f ..."
	sed -i 's/<sheetProtection[^>]*>//g' $f
done

echo "Zipping workbook files to ${DIR}/${FILENAME}-Unprotected.${EXTENSION} ..."
cd $WORK_DIR
zip -q -r "${DIR}/${FILENAME}-Unprotected.${EXTENSION}" ./
