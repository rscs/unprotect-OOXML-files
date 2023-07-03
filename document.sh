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
	echo "Usage: $0 <workbook filename .docx or .docm>" >&2
	exit 1
fi

if ! [ -f "$1" ]; then
	echo "Error: Input file $1 does not exist." >&2
	exit 1
fi

if ! [[ $(file -b "$1") = "Microsoft OOXML" || $(file -b "$1") = "Microsoft Word 2007+" ]]; then
	echo "Error: Input file $1 does not appear to be an XML file (.docx or .docm)." >&2
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

FILENAME=$(basename -- "$1")
EXTENSION="${FILENAME##*.}"
FILENAME="${FILENAME%.*}"
DIR="$( cd "$( dirname -- "$1" )" && pwd )"
WORK_DIR=$(mktemp -d)

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

echo "Extracting document files from ${DIR}/${FILENAME}${EXTENSION} ..."
unzip -q "$1" -d "$WORK_DIR"

echo 'Unprotecting document ...'
sed -i 's/<w:documentProtection[^>]*>//g' "$WORK_DIR/word/settings.xml"

echo "Zipping document files to ${DIR}/${FILENAME}-Unprotected.${EXTENSION} ..."
cd "$WORK_DIR" || exit
zip -q -r "${DIR}/${FILENAME}-Unprotected.${EXTENSION}" ./
