#!/usr/bin/env python3
# vi:fileencoding=utf-8

from subprocess import check_call
import argparse
from os.path import expanduser, join

tpdir = expanduser('~/text-proc')
result = join(tpdir, 'to_ventura.txt')
content = 'content.xml'
xsl = 'odt-to-ventura.xsl'
xsltproc = '/usr/local/bin/xsltproc_new'

if __name__ == '__main__':
    argparser = argparse.ArgumentParser()
    argparser.add_argument('odt_document', help='odt file to process')
    args = argparser.parse_args()
    target = expanduser(args.odt_document)

    check_call(['unzip', '-o', target, content, '-d', tpdir])
    check_call([xsltproc, '-o', result, xsl, join(tpdir, content)])
    print("done!")
    print("File saved as %s" % result)
