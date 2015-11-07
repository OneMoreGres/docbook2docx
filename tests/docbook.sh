#!/bin/bash

set -e
python3 ../docbook2docx.py docbook.xml docbook.docx template.docx
ls docbook.docx
