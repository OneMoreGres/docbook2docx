#encoding: utf-8

import os
import shutil
import zipfile
import re

import xml.etree.ElementTree as ET
from sys import argv
from PIL import Image

SupportedTags = {
'doc':
u'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" 
 xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" 
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" 
 xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" 
 xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
 xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" 
 xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">
  <w:body>{childTexts}</w:body>
</w:document>''',

'appendix':
u'''<w:p>
  <w:pPr><w:pStyle w:val="{style}"/></w:pPr>{bookmarkStart}
  <w:r><w:t></w:t></w:r>{bookmarkEnd}
</w:p>{childTexts}''',

'heading':
u'''<w:p>
  <w:pPr><w:pStyle w:val="{style}"/></w:pPr>{bookmarkStart}
  <w:r><w:t>{childTexts}</w:t></w:r>{bookmarkEnd}
</w:p>''',

'paragraph':
u'''<w:p>
  <w:pPr><w:pStyle w:val="{style}"/></w:pPr>{bookmarkStart}
  <w:r><w:t xml:space="preserve">{childTexts}</w:t></w:r>{bookmarkEnd}
</w:p>''',

'pageBreak':
u'''<w:p><w:r><w:br w:type="page"/></w:r></w:p>''',

'image':
u'''<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>{bookmarkStart}
      <w:r>
        <w:drawing>
          <wp:inline xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <wp:extent cx="{width}" cy="{height}"/>
            <wp:docPr id="1" name="Picture 1"/>
            <wp:cNvGraphicFramePr>
              <a:graphicFrameLocks noChangeAspect="1"/>
            </wp:cNvGraphicFramePr>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr>
                    <pic:cNvPr id="0" name="fileName"/>
                    <pic:cNvPicPr/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="{imageId}"/>
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="{width}" cy="{height}"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect"/>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>{bookmarkEnd}
    </w:p>''',

'listItem':
u'''<w:p>
    <w:pPr>
        <w:pStyle w:val="{style}"/>
        <w:numPr>
            <w:ilvl w:val="{currentListLevel}"/>
            <w:numId w:val="{currentListId}"/>
        </w:numPr>
    </w:pPr>
    <w:r><w:t xml:space="preserve">{childTexts}</w:t></w:r>
</w:p>''',

'toc':
u'''<w:sdt>
      <w:sdtPr>
        <w:docPartObj>
          <w:docPartGallery w:val="Table of Contents"/>
          <w:docPartUnique/>
        </w:docPartObj>
      </w:sdtPr>
      <w:sdtEndPr/>
      <w:sdtContent>
        <w:p>
          <w:pPr>
            <w:pStyle w:val="{titleStyle}"/>
          </w:pPr>
          <w:r>
            <w:t>{titleText}</w:t>
          </w:r>
        </w:p>
        <w:p>
          <w:r>
            <w:fldChar w:fldCharType="begin"/>
          </w:r>
          <w:r>
            <w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \\u </w:instrText>
          </w:r>
        </w:p>
        <w:p>
          <w:r>
            <w:fldChar w:fldCharType="end"/>
          </w:r>
        </w:p>
      </w:sdtContent>
    </w:sdt>''',

'tableFigureTitle':
u'''<w:p>
  <w:pPr><w:pStyle w:val="{style}"/></w:pPr>{bookmarkStart}
  <w:r><w:rPr><w:rStyle w:val="{numberStyle}"/></w:rPr><w:t xml:space="preserve">{numberTextBefore}</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:rPr><w:rStyle w:val="{numberStyle}"/></w:rPr><w:instrText xml:space="preserve"> REF {bookmarkRefName} \h \\t \w </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
  <w:r><w:rPr><w:rStyle w:val="{numberStyle}"/></w:rPr><w:t xml:space="preserve">{numberTextAfter}</w:t></w:r>
  <w:r><w:t xml:space="preserve">{childTexts}</w:t></w:r>{bookmarkEnd}
</w:p>''',

'table':
'''<w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/>
      </w:tblPr>
      {childTexts}
</w:tbl>''',

'tableRow':
'''<w:tr>
        <w:trPr>
          <w:trHeight w:val="454"/>
        </w:trPr>
{childTexts}
</w:tr>''',

'tableHead':
'''<w:tr>
        <w:trPr>
          <w:trHeight w:val="454"/>
          <w:tblHeader/>
        </w:trPr>
{childTexts}
</w:tr>''',

'tableCell':
'''<w:tc><w:tcPr><w:tcBorders>
              <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
              <w:left w:val="{leftBorder}" w:sz="4" w:space="0" w:color="auto"/>
            </w:tcBorders></w:tcPr>{childTexts}</w:tc>''',

'tableCellHead':
'''<w:tc><w:tcPr><w:tcBorders>
              <w:top w:val="single" w:sz="12" w:space="0" w:color="auto"/>
              <w:bottom w:val="single" w:sz="12" w:space="0" w:color="auto"/>
              <w:left w:val="{leftBorder}" w:sz="4" w:space="0" w:color="auto"/>
            </w:tcBorders>
            <w:vAlign w:val="center"/>
            </w:tcPr><w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr><w:r><w:t>{childTexts}</w:t></w:r></w:p></w:tc>''',

'glossaryTable':
'''<w:p><w:pPr><w:pStyle w:val="Pragraph"/></w:pPr><w:r><w:tbl>
      <w:tblPr>
        <w:tblW w:w="4670" w:type="pct"/>
        <w:tblInd w:w="817" w:type="dxa"/>
        <w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/>
      </w:tblPr>
      {childTexts}
</w:tbl></w:r></w:p>''',

'glossatyTermCell':
'''<w:tc><w:tcPr></w:tcPr><w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr><w:r><w:t>{childTexts}</w:t></w:r></w:p></w:tc>
'<w:tc><w:tcPr></w:tcPr><w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr><w:r><w:t>–</w:t></w:r></w:p></w:tc>''',

'tableCellBorderless':
'''<w:tc><w:tcPr></w:tcPr><w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr><w:r><w:t>{childTexts}</w:t></w:r></w:p></w:tc>''',

'passText':
'''{childTexts}''',

'bookmarkStart':
'''<w:bookmarkStart w:id="{bookmarkId}" w:name="{bookmarkRefName}"/>''',

'bookmarkEnd':
'''<w:bookmarkEnd w:id="{bookmarkId}"/>''',

'xref':
'''</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> REF {bookmarkRefName} \h \\t \w </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
<w:r><w:t xml:space="preserve">''',

'emphasis':
'''</w:t></w:r><w:r><w:rPr>{emphasisTags}</w:rPr><w:t xml:space="preserve">{childTexts}</w:t></w:r><w:r><w:t xml:space="preserve">''',

'sectionPr':
'''<w:sectPr w:rsidR="00CE772B" w:rsidSect="00CE772B">
          <w:headerReference w:type="default" r:id="rId9"/>
          <w:footerReference w:type="default" r:id="rId10"/>
          <w:pgSz w:w="11906" w:h="16838"/>
          <w:pgMar w:top="1417" w:right="567" w:bottom="850" w:left="1134" w:header="454" w:footer="454" w:gutter="0"/>
          <w:cols w:space="708"/>
          <w:docGrid w:linePitch="360"/>
        </w:sectPr>''',

'abstractNum':
'''<w:abstractNum w:abstractNumId="{numId}">
    <w:multiLevelType w:val="{multiLevelType}"/>
    {abstractNumLvls}
  </w:abstractNum>''',

'abstractNumLvl':
'''<w:lvl w:ilvl="{iLvl}">
      <w:start w:val="1"/>
      <w:numFmt w:val="{numFmt}"/>
      <w:lvlText w:val="{lvlText}"/>
      <w:lvlJc w:val="center"/>
    </w:lvl>''',

'num':
'''<w:num w:numId="{numId}">
    <w:abstractNumId w:val="{abstractNumId}"/>
</w:num>''',

'relationship':
'''  <Relationship Id="{id}" Type="{type}" Target="{target}"/>'''
}

def getTagName(node):
    'Get tag name without namespace'
    name = node.tag
    if name[0] == "{":
        return name[1:].split("}")[1]
    else:
        return name

class State:
    def __init__(self, converter):
        #TODO make more values stacked (not just last one)
        self.converter = converter
        
        self.currentHeadingLevel = 0
        self.currentListLevel = -1
        self.currentListId = 100
        self.usedLists = {}
        
        self.lastRsId = 100
        self.lastImageFile = ''
        self.lastImageWidth = 0
        self.lastImageHeight = 0
        self.maxImageWidth = 650
        self.maxImageHeight = 550
        self.usedImages = {}
        
        self.lastTitleBookmarkName = ''
        self.lastBookmarkName = ''
        self.lastXrefName = ''
        self.lastBookmarkId = 1
        self.bookmarks = {}
        
        self.lastEmphasisTagNames = []
        
        self.tableCellCol = -1
        
        self.listTags = ['orderedlist','itemizedlist']
        self.headingTags = ['section','chapter','preface', 'appendix']
        
    def enterTag(self, tagName, node):
        if tagName in self.listTags:
            self.currentListLevel += 1
            if self.currentListLevel == 0:
                self.currentListId += 1
                
        if tagName in self.headingTags:
            self.currentHeadingLevel += 1
              
        if tagName == 'imagedata':
            self.lastImageFile = node.attrib['fileref']
            if not self.lastImageFile in self.usedImages:
                self.lastRsId += 1   
                self.usedImages[self.lastImageFile] = self.currentRsId()
            img = Image.open(self.lastImageFile)
            w, h = img.size
            multiplier = 914400 / 96
            if (w/h) > 1:
                width = min([w, self.maxImageWidth])
                self.lastImageWidth = int(width * multiplier);
                scale = width/w
                self.lastImageHeight = int(h * scale * multiplier);
            else:
                height = min([h, self.maxImageHeight])
                self.lastImageHeight = int(height * multiplier);
                scale = height/h
                self.lastImageWidth = int(w * scale * multiplier);
            
        if 'id' in node.attrib:
            self.lastTitleBookmarkName = self.lastBookmarkName = node.attrib['id']
        elif self.lastBookmarkName and tagName in ['title']: #transfer id from meta tags(section, chapter) to their first title
            node.attrib['id'] = self.lastBookmarkName
            self.lastBookmarkName = ''
            
        if tagName == 'xref':
            self.lastXrefName = self.addBookmark(node.attrib['linkend'])[0]
            
        if tagName == 'emphasis':
            self.lastEmphasisTagNames.append('' if 'role' not in node.attrib else node.attrib['role'])
                    
        if tagName == 'row':
            self.tableCellCol = -1
        if tagName == 'entry':
            self.tableCellCol += 1
            
        if tagName == 'programlisting':
            self.converter.splitToParagraphs(node)
                
        if tagName == 'glossary':
            self.converter.filterGlossary(node)
                
        
    def leaveTag(self, tagName, node, currentText):
        if tagName in self.listTags:
            if not self.currentListId in self.usedLists:
                self.usedLists[self.currentListId] = []
            numeration = (node.attrib['numeration'] if 'numeration' in node.attrib else tagName)
            self.usedLists[self.currentListId].append({'level': self.currentListLevel, 'numeration': numeration})
            self.currentListLevel -= 1
            
        if tagName in self.headingTags:
            self.currentHeadingLevel -= 1
            
        if tagName == 'imageobject':
            self.lastImageFile = ''
            self.lastImageHeight = self.lastImageWidth = 0
            
        if tagName == 'xref':
            self.lastXrefName = ''
            
        if tagName == 'emphasis':
            self.lastEmphasisTagNames.pop()
                
    def currentListStyle(self):
        return 'ListLevel' + str(self.currentListLevel + 1)
    
    def addBookmark(self, name):
        if name in self.bookmarks:
            return self.bookmarks[name]['name'], self.bookmarks[name]['id']
        refName = "_Ref" + str(self.lastBookmarkId)
        bId = self.lastBookmarkId
        self.bookmarks[name] = {'id':bId, 'name':refName}
        self.lastBookmarkId += 1
        return refName, bId
        
    def leftCellBorder(self):
        return 'nil' if self.tableCellCol < 1 else 'single'
        
    def lastTitleBookmark(self):
        return self.addBookmark(self.lastTitleBookmarkName)[0]
        
    def headingStyle(self):
        return 'Heading' + str(self.currentHeadingLevel)
    
    def appendixHeadingStyle(self):
        return 'AppendixHeading' + str(self.currentHeadingLevel)
        
    def currentRsId(self):
        return 'rId' + str(self.lastRsId)
    
    def lastEmphasisTags(self):
        tags = []
        for tag in self.lastEmphasisTagNames:
            wordTag = {'strong': '<w:b/>', '': '<w:i/>'}[tag]
            tags.append(wordTag)
        return ''.join(tags)
    



class Converter:
    def __init__(self):
        self.path = ['']
        self.state = State(self)
        
        # Docx internal structure
        self.workDir = './work_bundle/'
        self.keepWorkDir = False
        self.contentTypesFileName = '[Content_Types].xml'
        self.documentFileName = 'word/document.xml'
        self.numberingFileName = 'word/numbering.xml'
        self.relationshipsFileName = 'word/_rels/document.xml.rels'
        self.mediaFileDir = 'word/media/'
        self.mediaLinkDir = 'media/'
        self.substituteDir = 'word/'
        
        # Conversion rules
        rules = {} 
        rules['//toc'] = 'Tag=toc, Format=titleStyle:TocTitle, Format=titleText:Содержание'
        rules['//simpara'] = 'Tag=paragraph, Format=style:Paragraph'
        rules['//itemizedlist/listitem/simpara'] = 'Tag=listItem, Format=style:state.currentListStyle, Format=currentListLevel:state.currentListLevel, Format=currentListId:state.currentListId'
        rules['//orderedlist/listitem/simpara'] = 'Tag=listItem, Format=style:state.currentListStyle, Format=currentListLevel:state.currentListLevel, Format=currentListId:state.currentListId'
        rules['//preface/title'] = 'Tag=heading, Format=style:Preface'
        rules['//chapter/title'] = 'Tag=heading, Format=style:state.headingStyle'
        rules['//section/title'] = 'Tag=heading, Format=style:state.headingStyle'
        rules['//glossary/title'] = 'Tag=heading, Format=style:GlossaryTitle'
        rules['//appendix'] = 'Tag=appendix, Format=style:Appendix'
        rules['//appendix/title'] = 'Tag=heading, Format=style:AppendixHeading1, Format=bookmarkStart: , Format=bookmarkEnd: '
        rules['//appendix/chapter/title'] = 'Tag=heading, Format=style:state.appendixHeadingStyle'
        rules['//appendix/section/title'] = 'Tag=heading, Format=style:state.appendixHeadingStyle'
        rules['//figure/title'] = 'Tag=tableFigureTitle, Format=style:FigureTitle, Format=numberStyle:FigureTitle, Format=numberTextBefore:Рис. , Format=numberTextAfter:. ,Format=bookmarkRefName:state.lastTitleBookmark'
        rules['//table/title'] = 'Tag=tableFigureTitle, Format=style:TableTitle, Format=numberStyle:TableNumber0, Format=numberTextBefore:Таблица , Format=numberTextAfter: – ,Format=bookmarkRefName:state.lastTitleBookmark'
        rules['//imageobject'] = 'Tag=image, Format=imageId:state.currentRsId, Format=style:Figure, Format=width:state.lastImageWidth, Format=height:state.lastImageHeight'
        rules['//informaltable'] = 'Tag=table'
        rules['//table'] = 'Tag=table'
        rules['//tbody/row'] = 'Tag=tableRow'
        rules['//thead/row'] = 'Tag=tableHead'
        rules['//tbody/row/entry'] = 'Tag=tableCell, Format=leftBorder:state.leftCellBorder'
        rules['//thead/row/entry/simpara'] = 'Tag=passText'
        rules['//thead/row/entry'] = 'Tag=tableCellHead, Format=leftBorder:state.leftCellBorder, Format=style:TableCellHead'
        rules['//row/entry/simpara'] = 'Tag=paragraph, Format=style:TableCell'
        rules['//glossary/variablelist'] = 'Tag=glossaryTable'
        rules['//glossary//varlistentry'] = 'Tag=tableRow'
        rules['//glossary//varlistentry/term'] = 'Tag=glossatyTermCell, Format=style:GlossaryEntry'
        rules['//glossary//varlistentry//simpara'] = 'Tag=tableCellBorderless, Format=style:GlossaryEntry'
        rules['//xref'] = 'Tag=xref, Format=bookmarkRefName:state.lastXrefName'
        rules['//programlisting/simpara'] = 'Tag=paragraph, Format=style:ProgramListing'
        rules['//literallayout'] = 'Tag=paragraph, Format=style:Paragraph'
        rules['//emphasis'] = 'Tag=emphasis, Format=emphasisTags:state.lastEmphasisTags'
        self.rules = self.compileRules(rules)

    def compileRules(self, rules):
        compiled = []
        for path, rule in rules.items():
            newPath = path.replace('//', '///')[1:].split('/')
            
            specificLength = 0
            for tag in reversed(newPath):
                if not tag:
                    break;
                specificLength += 1
                
            newRule = []
            for command in rule.split(','):
                params = []
                for param in command.lstrip().split('='):
                    if ':' in param:
                        params.append(param.lstrip().split(':'))
                    else:
                        params.append(param)
                newRule.append(params)
            compiled.append({'path':newPath, 'specificity': specificLength, 'rule': newRule})
        return compiled
        
    def getCurrentRule(self):
        maxSpecificity = -1
        maxPathLen = 0;
        matchedRule = [] # to return most detailed match
        for rule in self.rules:
            rulePath = rule['path'][:]
            pathLen = len(rulePath)
            match = True
            anyPath = False
            for tag in reversed(self.path):
                if not rulePath:
                    if not anyPath:
                        match = False;
                    break;
                if rulePath and tag != rulePath[-1]:
                    if anyPath:
                        continue;
                    if len(rulePath) > 1 and not rulePath[-1] and not rulePath[-2]:
                        anyPath = True;
                        rulePath.pop()
                        rulePath.pop()
                        continue;
                    match = False
                    break;
                anyPath = False
                rulePath.pop()
                
            if not match or (anyPath and rulePath):
                continue;
            if rule['specificity'] > maxSpecificity or (rule['specificity'] == maxSpecificity and pathLen > maxPathLen):
                maxPathLen = pathLen
                matchedRule = rule['rule']
                maxSpecificity = rule['specificity']
        return matchedRule
    
    def execFunc(self, name):
        complexTrigger = 'state.'
        if name.startswith(complexTrigger):
            attrName = name[len(complexTrigger):]
            if not attrName in dir(self.state):
                return name
            attrValue = getattr(self.state, attrName)
            if callable(attrValue):
                return attrValue()
            else:
                return attrValue
        return name
    
    def processRule(self, rule, childTexts, node):
        if not rule:
            return ''.join(childTexts)
        tag = ''
        formatArgs = {'childTexts':''.join(childTexts),'bookmarkStart': '','bookmarkEnd':''}
        if 'id' in node.attrib:
            refName, bId = self.state.addBookmark(node.attrib['id'])
            formatArgs['bookmarkStart'] = SupportedTags['bookmarkStart'].format(bookmarkId=bId,bookmarkRefName=refName)
            formatArgs['bookmarkEnd'] = SupportedTags['bookmarkEnd'].format(bookmarkId=bId)
        for command in rule:
            name, params = command[0], command[1]
            if name == 'Tag':
                tag = SupportedTags[params]
            elif name == 'Format':
                formatArgs[params[0]] = self.execFunc(params[1])
        text = ''
        if tag:
            text = tag.format(**formatArgs)
        return text
            
    def convertDoc(self, inFileName, outFileName, templateFileName):
        registerXmlNamespaces()
        self.extractTemplate(templateFileName);
        self.docDir = os.path.dirname(inFileName)
        tree = ET.parse(inFileName)
        self.root = tree.getroot()
        self.updateDocument(self.convert(self.root))
        self.updateNumbering()
        self.updateContentTypes()
        self.updateImages()
        self.makeSubstitutions()
        self.save(outFileName)

    def extractTemplate(self, templateFile):
        shutil.rmtree(self.workDir, True)
        with zipfile.ZipFile(templateFile, 'r') as f:
            f.extractall(self.workDir)

    def save(self, fileName):
        with zipfile.ZipFile(fileName, 'w', zipfile.ZIP_DEFLATED) as f:
            for path, _, files in os.walk(self.workDir):
                for fileName in files:
                    fullPath = os.path.normpath(os.path.join(path,fileName))
                    f.write(fullPath, os.path.relpath(fullPath, self.workDir))
                    
        if not self.keepWorkDir:
            shutil.rmtree(self.workDir)

    def readFile(self, fileName):
        with open(os.path.join(self.workDir, fileName), 'r', encoding='utf-8') as f:
            return f.read()

    def writeFile(self, fileName, data):
        with open(os.path.join(self.workDir, fileName), 'wb') as f:
            f.write(data.encode('utf8'))
        return None

    def updateDocument(self, converted):
        document = self.readFile(self.documentFileName)
        start = document.find('<w:t>removefromhere</w:t>')
        end = document.find('<w:t>removetillhere</w:t>')
        if start == -1 or end == -1:
            document = SupportedTags['doc'].format(childTexts=converted)
            return
        start = document.rfind('<w:p ', 0, start)
        end = document.find('</w:p>', end) + len('</w:p>')
        document = document[0:start] + '\n' + converted + '\n' + document[end:]
        self.writeFile(self.documentFileName, document)
        
    def updateNumbering(self):
        numbering = self.readFile(self.numberingFileName)
        lists = []
        usedLists = self.state.usedLists
        numFormats = {'itemizedlist': 'bullet', 'orderedlist': 'decimal', 'arabic': 'decimal', 'loweralpha': 'lowerLetter', 'bullet': 'bullet'}
        for listId, entries in usedLists.items():
            abstractLvls = []
            abstractLvlText = '' # combines with previous lvls
            usedLvls = [] # to avoid same level double definition
            sortedEntries = sorted(entries, key = lambda listItem: listItem['level'])
            for entry in sortedEntries:
                level = entry['level']
                if level in usedLvls:
                    continue
                usedLvls.append(level)
                if '–' in abstractLvlText or abstractLvlText.endswith(')'):
                    abstractLvlText = ''
                numFmt = numFormats[entry['numeration']]
                if numFmt == 'bullet':
                    abstractLvlText = '–'
                elif not abstractLvlText and numFmt in ['decimal', 'loweralpha', 'lowerLetter']:
                    abstractLvlText = '%' + str(level + 1) + ')'
                else:
                    abstractLvlText = abstractLvlText + '%' + str(entry['level'] + 1) + '.'
                abstractLvls.append(SupportedTags['abstractNumLvl'].format(iLvl=level,lvlText=abstractLvlText,numFmt=numFmt))

            multiLevelType = ('hybridMultilevel' if len(abstractLvls) > 1 else 'singleLevel') 
            abstract = SupportedTags['abstractNum'].format(numId=listId,abstractNumLvls=('\n'.join(abstractLvls)),multiLevelType=multiLevelType)
            lists.append(abstract)

        for listId, entries in usedLists.items():
            real = SupportedTags['num'].format(numId=listId,abstractNumId=listId)
            lists.append(real)
        numbering = numbering.replace('<w:num ', '\n' + '\n'.join(lists) + '\n<w:num ', 1)
        self.writeFile(self.numberingFileName, numbering)

    def updateContentTypes(self):
        contentTypes = self.readFile(self.contentTypesFileName)
        addTypes = []
        if contentTypes.find('Extension="png"') == -1:
            addTypes.append('<Default Extension="png" ContentType="image/png"/>')
        if contentTypes.find('Extension="jpg"') == -1:
            addTypes.append('<Default Extension="jpg" ContentType="image/jpeg"/>')
        if contentTypes.find('Extension="jpeg"') == -1:
            addTypes.append('<Default Extension="jpeg" ContentType="image/jpeg"/>')
        contentTypes = contentTypes.replace('</Types>', '\n' + '\n'.join(addTypes) + '\n</Types>')
        self.writeFile(self.contentTypesFileName, contentTypes)
    
    def updateImages(self):
        relationships = self.readFile(self.relationshipsFileName)
        rels = []
        mediaFileDir = os.path.join(self.workDir, self.mediaFileDir)
        if not os.path.exists(mediaFileDir):
            os.mkdir(mediaFileDir)
        for image, rId in self.state.usedImages.items():
            shutil.copyfile(os.path.join(self.docDir, image), os.path.join (mediaFileDir, os.path.basename(image)))
            relType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
            target = os.path.join(self.mediaLinkDir,os.path.basename(image))
            rels.append(SupportedTags['relationship'].format(id=rId,type=relType,target=target));
        relationships = relationships.replace('</Relationships>', '\n' + '\n'.join(rels) + '\n</Relationships>')
        self.writeFile(self.relationshipsFileName, relationships)
        
    def makeSubstitutions(self):
        substituteFiles = os.listdir(os.path.join(self.workDir, self.substituteDir))
        for fileName in substituteFiles:
            fullName = os.path.join(self.workDir, self.substituteDir, fileName)
            if os.path.isdir(fullName):
                continue
            tree = ET.parse(fullName)
            docxRoot = tree.getroot()
            if self.substitute(self.root, docxRoot):
                del docxRoot.attrib["{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"]
                with open(fullName, 'wb') as f:
                    f.write(ET.tostring(docxRoot, 'utf-8'))

    def substitute(self, docbook, docx):
        changed = False
        if getTagName(docx) == 'p':
            texts = self.getParagraphTexts(docx)
            if texts and self.substituteText(docbook, texts):
                index = 0
                for child in docx:
                    if getTagName(child) != 'r':
                        continue
                    for subchild in child:
                        if getTagName(subchild) != 't':
                            continue
                        subchild.text = texts[index]
                        index += 1
                        changed = True
        for child in docx:
            changed = self.substitute(docbook, child) or changed
        return changed

    def getParagraphTexts(self, node):
        texts = []
        text = ''
        for child in node:
            if getTagName(child) != 'r':
                continue
            added = False
            for subchild in child:
                if getTagName(subchild) != 't':
                    continue
                added = True
                if subchild.text:
                    text += subchild.text
                texts.append('')
            if not added and text:
                texts[-1] = text
                text = ''
        if text:
            texts[-1] = text
        return texts
    
    def substituteText(self, docbook, texts):
        modified = False
        i = 0
        for text in texts:
            start = text.find('{{')
            end = text.find('}}')
            while start != -1 and end != -1:
                found = text[start:end + 2]
                path = './bookinfo/' + found[2:-2]
                sub = docbook.find(path)
                modified = True
                text = text.replace(found, path if (sub is None or sub.text is None) else sub.text)
                texts[i] = text
                start = text.find('{{')
                end = text.find('}}')
            i += 1
        return modified
    
    def convert(self, node):
        'Convert docbook node with children'
        tagName = getTagName(node)
        self.path.append(tagName)
        self.state.enterTag(tagName, node)
        childTexts = []
        for child in node:
            childTexts.append(self.convert(child))
        rule = self.getCurrentRule()
        if len(node) > 0 and node.text and len(node.text.strip()): # mixed text and tags
            childTexts = self.replaceConvertedChildren(node, childTexts)
        elif rule and node.text:
            childTexts.append(node.text.replace('<','&lt;').replace('>','&gt;'))
        currentText = self.processRule(rule, childTexts, node)
        self.state.leaveTag(tagName, node, currentText)
        self.path.pop()
        return currentText
    
    def replaceConvertedChildren(self, node, childTexts):
        'Replace converted children texts in node text. For mixed text/children node (bla-bla <tag>tagged</tag> bla-bla)'
        resultTexts = []
        lastFullXmlPos = 0
        fullXml = ET.tostring(node, 'utf-8')
        start, end = fullXml.find(b'>'), fullXml.rfind(b'<') # to extract tag contents
        fullXml = fullXml[start + 1:end]
        childIndex = 0
        for child in node:
            childXml = ET.tostring(child, 'utf-8')
            start, end = childXml.find(b'<'), childXml.rfind(b'>') # extract tag because ET can append tailing text to child
            childXml = childXml[start:end + 1]
            fullXmlChildPos = fullXml.find(childXml, lastFullXmlPos)
            if fullXmlChildPos > lastFullXmlPos:
                resultTexts.append (fullXml[lastFullXmlPos:fullXmlChildPos].decode('utf-8'))
            lastFullXmlPos = fullXmlChildPos + len(childXml)
            resultTexts.append (childTexts[childIndex])
            childIndex += 1
        if lastFullXmlPos < len(fullXml):
            resultTexts.append (fullXml[lastFullXmlPos:len(fullXml)].decode('utf-8'))
        return resultTexts
    
    def filterGlossary(self, glossary):
        document = ET.tostring(self.root, 'utf-8').decode('utf-8')
        for child in glossary:
            if getTagName(child) != 'variablelist':
                continue
            toRemove = []
            for subchild in child: #entry
                for i in subchild:
                    if getTagName(i) == 'term':
                        m = re.search(r'(?!<term>)[\s«\(",\.]' + i.text.strip() + r'[\s"»\),\.](?!</term>)', document)
                        if not m:
                            toRemove.append(subchild)
                            break  
            for remove in toRemove:
                child.remove(remove)
                
    def splitToParagraphs(self, node):
        parts = node.text.split('\n')
        for part in parts:
            el = ET.Element('simpara')
            el.text = part
            node.append(el)

def registerXmlNamespaces():
    ET.register_namespace('wpc', 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas') 
    ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006') 
    ET.register_namespace('o', 'urn:schemas-microsoft-com:office:office') 
    ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') 
    ET.register_namespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
    ET.register_namespace('v', 'urn:schemas-microsoft-com:vml') 
    ET.register_namespace('wp14', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing') 
    ET.register_namespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing') 
    ET.register_namespace('w10', 'urn:schemas-microsoft-com:office:word') 
    ET.register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') 
    ET.register_namespace('w14', 'http://schemas.microsoft.com/office/word/2010/wordml')
    ET.register_namespace('wpg', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup') 
    ET.register_namespace('wpi', 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk') 
    ET.register_namespace('wne', 'http://schemas.microsoft.com/office/word/2006/wordml') 
    ET.register_namespace('wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape')
    ET.register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    ET.register_namespace('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture')

def usage():
    print('Usage: convertor <in_docbook_file> [<out_docx_file>] [<docx_template_file>]')

def mainFunc(argv):
    if len(argv) < 1:
        usage()
        return
    inFileName = argv[0]
    outFileName = inFileName + '.docx' if len(argv) < 2 else argv[1]
    templateFileName = "./template.docx" if len(argv) < 3 else argv[2] 
    converter = Converter ()
    converter.convertDoc(inFileName, outFileName, templateFileName)

if __name__ == "__main__":
    mainFunc(argv[1:])
