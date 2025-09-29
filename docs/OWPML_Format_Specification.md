# HWPX File Format Specification

## Overview

HWPX (Hancom Word Processor X) is an XML-based document format used by Hancom Office Hangul (한글). It is based on the OWPML (Open Word processing ML) specification and follows an OOXML-like structure. HWPX files are essentially ZIP archives containing XML files and associated resources that define the document's content, formatting, and metadata.

## File Structure

A HWPX file is a ZIP archive with the following internal structure:

```
document.hwpx
├── mimetype                          # MIME type declaration
├── version.xml                       # Version information
├── META-INF/
│   ├── container.xml                 # Container description
│   ├── manifest.xml                  # File manifest
│   ├── signatures.xml                # Digital signatures (optional)
│   └── container.rdf                 # RDF metadata (optional)
├── Contents/
│   ├── content.hpf                   # Main content package file
│   ├── header.xml                    # Document header information
│   ├── section0.xml                  # Document sections (section0, section1, ...)
│   ├── section1.xml
│   └── ...
├── Preview/
│   ├── PrvText.txt                   # Preview text
│   └── PrvImage.png                  # Preview image
├── Scripts/                          # JavaScript files (optional)
│   ├── headerScripts.js
│   ├── sourceScripts.js
│   ├── preScripts.js
│   └── postScripts.js
├── XMLTemplate/                      # XML template information (optional)
│   ├── TemplateSchemaName.txt
│   ├── TemplateSchema.xsd
│   └── TemplateInstance.xml
├── DocHistory/                       # Document history (optional)
│   └── historylastdoc.hml
├── Custom/                          # Custom data (optional)
│   └── bibliography.xml
├── BinData/                         # Binary data (images, objects)
│   ├── BIN0001.png
│   ├── BIN0002.jpg
│   └── ...
└── Chart/                           # Chart data (optional)
    ├── chart0.xml
    └── ...
```

## Core Files Description

### mimetype
A plain text file containing the MIME type identifier:
```
application/vnd.hancom.hwpx
```

### version.xml
Contains version information for the HWPX format:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<version xmlns="http://www.hancom.co.kr/hwpml/2011/version">1.31</version>
```

### META-INF/container.xml
Defines the main content file location:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<container xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
    <rootFiles>
        <rootFile media-type="application/vnd.hancom.hwpx"
                  full-path="Contents/content.hpf"/>
    </rootFiles>
</container>
```

### META-INF/manifest.xml
Lists all files in the package with their media types:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<manifest xmlns="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
    <file-entry media-type="application/vnd.hancom.hwpx" full-path="/"/>
    <file-entry media-type="application/xml" full-path="Contents/header.xml"/>
    <file-entry media-type="application/xml" full-path="Contents/section0.xml"/>
    <!-- ... more entries ... -->
</manifest>
```

### Contents/content.hpf
The main content package file that references other content files:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf">
    <metadata>
        <!-- Document metadata -->
    </metadata>
    <manifest>
        <!-- File manifest -->
    </manifest>
    <spine>
        <!-- Reading order -->
    </spine>
</package>
```

### Contents/header.xml
Contains document-level formatting information, styles, and properties. This file uses the HWPML namespace and includes:

- Document properties and settings
- Font definitions
- Style definitions (paragraph styles, character styles)
- Numbering definitions
- Page layout settings
- Header/footer definitions

```xml
<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" version="1.31" secCnt="1">
    <!-- Font faces -->
    <hh:fontfaces>
        <hh:fontface name="맑은 고딕" type="ttf">
            <!-- Font definition -->
        </hh:fontface>
    </hh:fontfaces>
    
    <!-- Border fills -->
    <hh:borderfills>
        <!-- Border and fill definitions -->
    </hh:borderfills>
    
    <!-- Character shapes -->
    <hh:charshapes>
        <!-- Character formatting definitions -->
    </hh:charshapes>
    
    <!-- Paragraph shapes -->
    <hh:parashapes>
        <!-- Paragraph formatting definitions -->
    </hh:parashapes>
    
    <!-- Styles -->
    <hh:styles>
        <!-- Style definitions -->
    </hh:styles>
</hh:head>
```

### Contents/section[n].xml
Contains the actual document content organized by sections. Each section represents a major part of the document with its own page settings and content.

```xml
<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section">
    <hp:p xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
        <hp:run>
            <hp:t>Hello World</hp:t>
        </hp:run>
    </hp:p>
</hs:sec>
```

## OWPML Namespaces

The OWPML format uses several XML namespaces:

- **hh**: `http://www.hancom.co.kr/hwpml/2011/head` - Header elements
- **hs**: `http://www.hancom.co.kr/hwpml/2011/section` - Section elements  
- **hp**: `http://www.hancom.co.kr/hwpml/2011/paragraph` - Paragraph elements
- **hc**: `http://www.hancom.co.kr/hwpml/2011/core` - Core elements
- **he**: `http://www.hancom.co.kr/hwpml/2011/etc` - Miscellaneous elements

## Document Structure Hierarchy

```
Document
├── Head (header.xml)
│   ├── Fontfaces
│   ├── BorderFills  
│   ├── CharShapes
│   ├── ParaShapes
│   ├── Styles
│   └── Numberings
└── Sections (section0.xml, section1.xml, ...)
    └── Paragraphs
        └── Runs
            ├── Text
            ├── Images
            ├── Tables
            ├── Objects
            └── Controls
```

## Content Elements

### Paragraph Structure
Paragraphs (`<hp:p>`) are the fundamental content units containing:
- **Runs** (`<hp:run>`): Text segments with consistent formatting
- **Text** (`<hp:t>`): Actual text content
- **Line segments** (`<hp:lineseg>`): Line break information
- **Fields** (`<hp:fld>`): Dynamic content like page numbers, dates
- **Controls**: Form controls, OLE objects, images, tables

### Text Elements
- `<hp:t>`: Plain text content
- `<hp:char>`: Special characters and symbols
- `<hp:tab>`: Tab characters
- `<hp:nbspace>`: Non-breaking spaces
- `<hp:hypen>`: Hyphenation points

### Object Elements
- `<hp:table>`: Tables with rows and cells
- `<hp:pic>`: Images and pictures
- `<hp:ole>`: OLE objects
- `<hp:equation>`: Mathematical equations
- `<hp:chart>`: Charts and graphs

## Formatting Model

### Character Formatting
Character properties are defined in the header and referenced by ID:
- Font family, size, color
- Bold, italic, underline, strikethrough
- Superscript, subscript
- Character spacing and scaling

### Paragraph Formatting  
Paragraph properties include:
- Alignment (left, center, right, justify)
- Indentation and spacing
- Line spacing
- Borders and shading
- Tab stops
- Numbering and bullets

### Page Layout
Page settings define:
- Page size and orientation
- Margins
- Headers and footers
- Columns
- Section breaks

## Binary Data Handling

Binary resources (images, OLE objects) are stored in the `BinData/` directory and referenced from XML content using:
- Unique identifiers (BIN0001, BIN0002, etc.)
- File paths relative to the package root
- Media type information in the manifest

## Compatibility and Versioning

The HWPX format includes compatibility settings for:
- Different versions of Hancom Office
- Microsoft Word compatibility
- Layout and rendering compatibility options

Current format version: **1.31** (as of 2021)

## Security Features

HWPX supports:
- Digital signatures (stored in META-INF/signatures.xml)
- Document encryption
- Password protection
- Access control and permissions

## Extensibility

The format supports extensions through:
- Custom XML elements in defined namespaces
- Embedded scripts and macros
- Custom document properties
- Template systems
- Bibliography and citation management

## Related Specifications

- **OOXML**: Office Open XML (ISO/IEC 29500)
- **ODF**: Open Document Format (ISO/IEC 26300)  
- **OWPML**: Open Word processing ML (Hancom specification)

For the complete OWPML specification, refer to: https://www.hancom.com/etc/hwpDownload.do