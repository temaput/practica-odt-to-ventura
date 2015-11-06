#!/bin/bash
usage() {
    echo "Usage: $0 odt_file"
}

tpdir=~/text-proc
result=$tpdir/to_ventura.txt

echo "$1"
if [ -e "$1" ] ; then
    unzip -o $1 content.xml -d $tpdir
    /usr/local/bin/xsltproc_new -o $result odt-to-ventura.xsl $tpdir/content.xml
    echo "done!"
    echo "File saved as $result"
else
    usage
fi



