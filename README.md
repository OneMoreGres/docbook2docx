[![Build Status](https://travis-ci.org/OneMoreGres/docbook2docx.svg)]
(https://travis-ci.org/OneMoreGres/docbook2docx.svg)

# Docbook to docx conversion script

## Intro

As name says, this script performs conversion of docbook file into docx file.

It also performs substitutions in docx document with data, defined in `dokinfo` tag,
and automatic figure/table numeration.

It was made as a part of asciidoc-to-docx conversion chain to produce software
documentation in specific format.
So it may process some specific tags and not process some common.

Currently there are a lot of limitations, but main part seems to be working.

## Requirements

Script requires python3 and its PIL module installed.

It also requires docx template file, that contains information about styles to use
and header/footer data.

## Usage

    python3 docbook2docx.py <in_docbook_file> [<out_docx_file>] [<docx_template_file>]

When script finishes work open docx document, select all and press `F9` in order to update
automatically generated fields (ToC, figures/tables, links, etc).

## Template rules

Template file example can be found in `tests` dir.

All data between words `removefromhere` and `removetillhere` will be replaced with
generated text.
Current solution have quiet ugly implementation, so in order to make this words start working
sometimes you should select and apply `Clear format` action on them in MS Word.

Another useful feature is variable substutution.
Script searches text inside `{{}}` and replaces it with variables in docbook's `docinfo` tag.
