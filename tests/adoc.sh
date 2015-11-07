#!/bin/bash

set -e
asciidoc -b docbook -o docbook.xml adoc.adoc
ls docbook.xml
