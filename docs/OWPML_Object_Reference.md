# OWPML(HWPX) Object Type Reference

This document provides detailed XML specifications for every object type in the OWPML(HWPX) document model, focusing on XML elements, attributes, namespaces, and schema structure.

## XML Namespaces and Schema Overview

HWPX documents are built using the OWPML (Open Word processing ML) XML schema with multiple namespaces. Understanding these namespaces is critical for working with HWPX objects:

### Primary OWPML Namespaces

| Prefix | Namespace URI | Purpose | Example Elements |
|--------|---------------|---------|------------------|
| `hh` | `http://www.hancom.co.kr/hwpml/2011/head` | Header definitions | `<hh:head>`, `<hh:fontface>`, `<hh:style>` |
| `hs` | `http://www.hancom.co.kr/hwpml/2011/section` | Section structure | `<hs:sec>` |
| `hp` | `http://www.hancom.co.kr/hwpml/2011/paragraph` | Paragraph content | `<hp:p>`, `<hp:run>`, `<hp:t>` |
| `hc` | `http://www.hancom.co.kr/hwpml/2011/core` | Core elements | `<hc:sz>`, `<hc:pos>` |
| `he` | `http://www.hancom.co.kr/hwpml/2011/etc` | Miscellaneous elements | `<he:reference>` |

### Standard XML Namespaces

| Prefix | Namespace URI | Purpose |
|--------|---------------|---------|
| `opf` | `http://www.idpf.org/2007/opf/` | Open Packaging Format |
| `dc` | `http://purl.org/dc/elements/1.1/` | Dublin Core metadata |
| `odf` | `urn:oasis:names:tc:opendocument:xmlns:manifest:1.0` | ODF manifest |
| `ocf` | `urn:oasis:names:tc:opendocument:xmlns:container` | OCF container |

### XML Schema Structure

The HWPX XML schema follows a hierarchical structure:

```xml
<!-- Document root with namespace declarations -->
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" 
         xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
         xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
         xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"
         version="1.31" secCnt="1">
    
    <!-- Document-level definitions -->
    <hh:fontfaces>...</hh:fontfaces>
    <hh:charshapes>...</hh:charshapes>
    <hh:parashapes>...</hh:parashapes>
    <hh:styles>...</hh:styles>
</hh:head>
```

## Head Objects (Document-Level XML Definitions)

### HWPMLHeadType
**XML Element**: `<hh:head>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/head`  
**Purpose**: Root container for all document-level definitions including styles, fonts, and global settings.

**XML Schema Definition**:
```xml
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" 
         version="1.31" 
         secCnt="1">
    <hh:beginNum>...</hh:beginNum>
    <hh:refList>...</hh:refList>
    <hh:fontfaces>...</hh:fontfaces>
    <hh:borderfills>...</hh:borderfills>
    <hh:charshapes>...</hh:charshapes>
    <hh:parashapes>...</hh:parashapes>
    <hh:styles>...</hh:styles>
    <hh:numberings>...</hh:numberings>
    <hh:bullets>...</hh:bullets>
</hh:head>
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `version` | string | Yes | OWPML format version (e.g., "1.31") |
| `secCnt` | integer | Yes | Number of sections in the document |

**XML Child Elements**:
| Element | XML Tag | Occurrence | Description |
|---------|---------|------------|-------------|
| Begin Numbers | `<hh:beginNum>` | 0..1 | Starting numbers for various counters |
| Reference List | `<hh:refList>` | 0..1 | Reference mappings and lookup tables |
| Forbidden Words | `<hh:forbiddenWordList>` | 0..1 | Words that cannot be hyphenated |
| Compatible Document | `<hh:compatibleDocument>` | 0..1 | Compatibility settings |
| Document Options | `<hh:docOption>` | 0..1 | Document behavior options |
| Track Changes | `<hh:trackChangeConfig>` | 0..1 | Track changes configuration |
| Font Faces | `<hh:fontfaces>` | 1 | Font definitions collection |
| Border Fills | `<hh:borderfills>` | 1 | Border and fill style definitions |
| Character Shapes | `<hh:charshapes>` | 1 | Character formatting definitions |
| Paragraph Shapes | `<hh:parashapes>` | 1 | Paragraph formatting definitions |
| Styles | `<hh:styles>` | 1 | Named style definitions |
| Numberings | `<hh:numberings>` | 0..1 | Numbering scheme definitions |
| Bullets | `<hh:bullets>` | 0..1 | Bullet style definitions |
| Memo Properties | `<hh:memoProperties>` | 0..1 | Comment and annotation properties |

### FontfaceType
**XML Element**: `<hh:fontface>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/head`  
**Parent Element**: `<hh:fontfaces>`  
**Purpose**: Defines a font family with language-specific variants and fallback options.

**XML Schema Definition**:
```xml
<hh:fontface name="맑은 고딕" type="ttf">
    <hh:font lang="HANGUL" name="맑은 고딕"/>
    <hh:font lang="LATIN" name="Malgun Gothic"/>
    <hh:font lang="HANJA" name="맑은 고딕"/>
    <hh:font lang="JAPANESE" name="Malgun Gothic"/>
    <hh:font lang="OTHER" name="Malgun Gothic"/>
    <hh:font lang="SYMBOL" name="Malgun Gothic"/>
    <hh:font lang="USER" name="Malgun Gothic"/>
    <hh:substFont type="ttf" name="Arial"/>
    <hh:typeInfo familyType="2" serifType="11" weight="5" 
                 proportion="2" contrast="4" strokeVariation="5" 
                 armStyle="5" letterform="2" midline="2" xHeight="4"/>
</hh:fontface>
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `name` | string | Yes | Font family name |
| `type` | enum | Yes | Font type: "ttf", "otf", "woff", "eot" |

**XML Child Elements**:
| Element | XML Tag | Occurrence | Description |
|---------|---------|------------|-------------|
| Font Definition | `<hh:font>` | 1..* | Individual font per language |
| Substitute Font | `<hh:substFont>` | 0..1 | Fallback font mapping |
| Type Information | `<hh:typeInfo>` | 0..1 | Font metrics and characteristics |

**Language-Specific Font Support**:
The `<hh:font>` element supports different languages via the `lang` attribute:
- `HANGUL`: Korean characters
- `LATIN`: Western/Latin characters  
- `HANJA`: Chinese characters
- `JAPANESE`: Japanese characters
- `SYMBOL`: Symbol characters
- `OTHER`: Other character sets
- `USER`: User-defined characters

### CharShapeType
**XML Element**: `<hh:charshape>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/head`  
**Parent Element**: `<hh:charshapes>`  
**Purpose**: Defines character-level formatting that can be applied to text runs.

**XML Schema Definition**:
```xml
<hh:charshape id="0" height="1000" textColor="0" shadeColor="4294967295" 
              useFontSpace="false" useKerning="false" symMark="0">
    <hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" 
                other="0" symbol="0" user="0"/>
    <hh:ratio hangul="100" latin="100" hanja="100" japanese="100" 
              other="100" symbol="100" user="100"/>
    <hh:spacing hangul="0" latin="0" hanja="0" japanese="0" 
                other="0" symbol="0" user="0"/>
    <hh:relSz hangul="100" latin="100" hanja="100" japanese="100" 
              other="100" symbol="100" user="100"/>
    <hh:offset hangul="0" latin="0" hanja="0" japanese="0" 
               other="0" symbol="0" user="0"/>
    <hh:italic hangul="false" latin="false" hanja="false" japanese="false" 
               other="false" symbol="false" user="false"/>
    <hh:bold hangul="false" latin="false" hanja="false" japanese="false" 
             other="false" symbol="false" user="false"/>
    <hh:underline type="none" shape="straight" color="0"/>
    <hh:strikeout type="none" shape="straight" color="0"/>
    <hh:outline type="none"/>
    <hh:shadow type="none" color="0" offsetX="283" offsetY="283"/>
    <hh:emboss type="none"/>
    <hh:engrave type="none"/>
    <hh:supscript type="none"/>
    <hh:subscript type="none"/>
</hh:charshape>
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `id` | integer | Yes | Unique identifier for referencing |
| `height` | integer | Yes | Base font size in HWPML units (1000 = 10pt) |
| `textColor` | integer | No | Text color as RGB integer |
| `shadeColor` | integer | No | Background color for text |
| `useFontSpace` | boolean | No | Whether to use font-specific spacing |
| `useKerning` | boolean | No | Whether to apply kerning |
| `symMark` | integer | No | Symbol mark type for diacritics |

**XML Child Elements** (all language-specific):
| Element | XML Tag | Attributes | Description |
|---------|---------|------------|-------------|
| Font References | `<hh:fontRef>` | hangul, latin, hanja, japanese, other, symbol, user | References to font faces |
| Size Ratios | `<hh:ratio>` | hangul, latin, hanja, japanese, other, symbol, user | Width/height scaling ratios |
| Character Spacing | `<hh:spacing>` | hangul, latin, hanja, japanese, other, symbol, user | Character spacing adjustments |
| Relative Sizes | `<hh:relSz>` | hangul, latin, hanja, japanese, other, symbol, user | Relative size adjustments |
| Baseline Offsets | `<hh:offset>` | hangul, latin, hanja, japanese, other, symbol, user | Baseline offset adjustments |
| Italic Style | `<hh:italic>` | hangul, latin, hanja, japanese, other, symbol, user | Italic formatting per language |
| Bold Style | `<hh:bold>` | hangul, latin, hanja, japanese, other, symbol, user | Bold formatting per language |
| Underline | `<hh:underline>` | type, shape, color | Underline style definition |
| Strikeout | `<hh:strikeout>` | type, shape, color | Strikethrough formatting |
| Outline | `<hh:outline>` | type | Text outline effects |
| Shadow | `<hh:shadow>` | type, color, offsetX, offsetY | Text shadow effects |
| Emboss | `<hh:emboss>` | type | Embossed text effects |
| Engrave | `<hh:engrave>` | type | Engraved text effects |
| Superscript | `<hh:supscript>` | type | Superscript positioning |
| Subscript | `<hh:subscript>` | type | Subscript positioning |

### ParaShapeType
**XML Element**: `<hh:parashape>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/head`  
**Parent Element**: `<hh:parashapes>`  
**Purpose**: Defines paragraph-level formatting including alignment, spacing, and indentation.

**XML Schema Definition**:
```xml
<hh:parashape id="0" align="LEFT" heading="0" breakSetting="0" 
              condense="0" fontLineHeight="false" snapToGrid="false" 
              suppressLineNumbers="false" checked="false">
    <hh:margin left="0" right="0" first="0" hanging="0"/>
    <hh:lineSpacing type="PERCENT" value="160" unit="PERCENT"/>
    <hh:border borderFillIDRef="1"/>
    <hh:autoSpacing eParaHeadTail="false" eParaTail="false" 
                    eParaHead="false" eParaMiddle="false"/>
</hh:parashape>
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `id` | integer | Yes | Unique identifier |
| `align` | enum | No | Text alignment: "LEFT", "CENTER", "RIGHT", "JUSTIFY", "DISTRIBUTE" |
| `heading` | integer | No | Outline heading level |
| `breakSetting` | integer | No | Page and column break behavior flags |
| `condense` | integer | No | Text condensation percentage |
| `fontLineHeight` | boolean | No | Line height calculation method |
| `snapToGrid` | boolean | No | Whether to snap to document grid |
| `suppressLineNumbers` | boolean | No | Whether to suppress line numbering |
| `checked` | boolean | No | Checkbox state for list items |

**XML Child Elements**:
| Element | XML Tag | Attributes | Description |
|---------|---------|------------|-------------|
| Margins | `<hh:margin>` | left, right, first, hanging | Paragraph margin specifications |
| Line Spacing | `<hh:lineSpacing>` | type, value, unit | Line spacing definitions |
| Border | `<hh:border>` | borderFillIDRef | Reference to border/fill definition |
| Auto Spacing | `<hh:autoSpacing>` | eParaHeadTail, eParaTail, eParaHead, eParaMiddle | Automatic spacing between paragraphs |

### BorderFillType
**XML Element**: `<hh:borderfill>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/head`  
**Parent Element**: `<hh:borderfills>`  
**Purpose**: Defines border and background fill styles for various elements.

**XML Schema Definition**:
```xml
<hh:borderfill id="1" threeD="false" shadow="false">
    <hh:slash type="NONE" crooked="false" isCounter="false"/>
    <hh:leftBorder type="NONE" width="0" color="0"/>
    <hh:rightBorder type="NONE" width="0" color="0"/>
    <hh:topBorder type="NONE" width="0" color="0"/>
    <hh:bottomBorder type="NONE" width="0" color="0"/>
    <hh:diagonal type="NONE" width="0" color="0"/>
    <hh:fillBrush>
        <hh:windowBrush faceColor="4294967295" hatchColor="0" hatchStyle="HS_HORIZONTAL"/>
    </hh:fillBrush>
</hh:borderfill>
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `id` | integer | Yes | Unique identifier |
| `threeD` | boolean | No | 3D border effect settings |
| `shadow` | boolean | No | Shadow effect settings |

**XML Child Elements**:
| Element | XML Tag | Attributes | Description |
|---------|---------|------------|-------------|
| Slash Pattern | `<hh:slash>` | type, crooked, isCounter | Diagonal line patterns |
| Border Sides | `<hh:leftBorder>`, `<hh:rightBorder>`, `<hh:topBorder>`, `<hh:bottomBorder>` | type, width, color | Individual border definitions |
| Diagonal | `<hh:diagonal>` | type, width, color | Diagonal border |
| Fill Brush | `<hh:fillBrush>` | - | Background fill definition |

**Border Line Types**: `NONE`, `SOLID`, `THICK`, `DOUBLE`, `DOTTED`, `DASHED`, `DASHDOTTED`, `DASHDOTDOTTED`, `LONGDASH`, `CIRCLED`, `DOUBLE_SLIM`, `SLIM_THICK`, `THICK_SLIM`, `SLIM_THICK_SLIM`, `WAVE`, `DOUBLE_WAVE`, `THICK_3D`, `THICK_3D_LOWERED`, `3D`, `3D_LOWERED`

### StyleType
**XML Element**: `<hh:style>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/head`  
**Parent Element**: `<hh:styles>`  
**Purpose**: Named style definitions that combine character and paragraph formatting.

**XML Schema Definition**:
```xml
<hh:style id="0" type="PARA" name="본문" engName="Normal" 
          paraPr="0" charPr="0" nextStyle="0" langId="1042" lockForm="false">
    <!-- Style properties are referenced through paraPr and charPr IDs -->
</hh:style>
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `id` | integer | Yes | Unique identifier |
| `type` | enum | Yes | Style type: "PARA" (paragraph), "CHAR" (character) |
| `name` | string | Yes | User-visible style name |
| `engName` | string | No | English name for the style |
| `paraPr` | integer | No | Reference to paragraph shape ID |
| `charPr` | integer | No | Reference to character shape ID |
| `nextStyle` | integer | No | Style ID to apply to following paragraphs |
| `langId` | integer | No | Language identifier (1042 = Korean) |
| `lockForm` | boolean | No | Whether style formatting is locked |

### NumberingType
**XML Element**: `<hh:numbering>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/head`  
**Parent Element**: `<hh:numberings>`  
**Purpose**: Defines numbering schemes for ordered lists and outline numbering.

**XML Schema Definition**:
```xml
<hh:numbering id="0">
    <hh:paraHead level="0" align="LEFT" useInstWidth="false" 
                 autoIndent="true" widthAdjust="false" 
                 textOffsetType="PERCENT" textOffset="50" 
                 numFormat="DIGIT" charShape="0">
        <hh:start value="1"/>
    </hh:paraHead>
    <hh:paraHead level="1" align="LEFT" useInstWidth="false" 
                 autoIndent="true" widthAdjust="false" 
                 textOffsetType="PERCENT" textOffset="50" 
                 numFormat="CIRCLED_DIGIT" charShape="0">
        <hh:start value="1"/>
    </hh:paraHead>
    <!-- Additional levels as needed -->
</hh:numbering>
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `id` | integer | Yes | Unique identifier |

**XML Child Elements**:
| Element | XML Tag | Attributes | Description |
|---------|---------|------------|-------------|
| Paragraph Head | `<hh:paraHead>` | level, align, useInstWidth, autoIndent, widthAdjust, textOffsetType, textOffset, numFormat, charShape | Numbering definition for each level |
| Start Number | `<hh:start>` | value | Starting number for the level |

**Numbering Formats** (numFormat attribute):
- `DIGIT`: Decimal (1, 2, 3...)
- `CIRCLED_DIGIT`: Circled decimal (①, ②, ③...)
- `ROMAN_CAPITAL`: Upper Roman (I, II, III...)
- `ROMAN_SMALL`: Lower Roman (i, ii, iii...)
- `LETTER_CAPITAL`: Upper letters (A, B, C...)
- `LETTER_SMALL`: Lower letters (a, b, c...)
- `HANGUL_SYLLABLE`: Korean syllables (가, 나, 다...)
- `HANGUL_JAMO`: Korean jamo (ㄱ, ㄴ, ㄷ...)
- `HANGUL_PHONETIC`: Korean phonetic
- `IDEOGRAPH_DIGITAL`: Chinese numerals

### BulletType
**XML Element**: `<hh:bullet>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/head`  
**Parent Element**: `<hh:bullets>`  
**Purpose**: Defines bullet styles for unordered lists.

**XML Schema Definition**:
```xml
<hh:bullet id="0" char="●" checkedChar="☑" useImage="false">
    <hh:img src="BIN0001.png" embedded="true" binaryItem="BIN0001"/>
</hh:bullet>
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `id` | integer | Yes | Unique identifier |
| `char` | string | No | Bullet character (Unicode) |
| `checkedChar` | string | No | Character for checked items |
| `useImage` | boolean | No | Whether to use image bullets |

**XML Child Elements**:
| Element | XML Tag | Attributes | Description |
|---------|---------|------------|-------------|
| Image Bullet | `<hh:img>` | src, embedded, binaryItem | Image source for custom bullets |

## Section Objects (Document Structure XML)

### SectionType
**XML Element**: `<hs:sec>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/section`  
**Purpose**: Major document divisions with independent page layout and content flow.

**XML Schema Definition**:
```xml
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
    <hs:secPr>
        <hs:pageSz width="59528" height="84188"/>
        <hs:pageMargin left="8504" right="8504" top="5668" bottom="4536" 
                       header="4536" footer="4536" gutter="0"/>
        <hs:landscape value="false"/>
        <hs:cols type="1" layout="LEFT"/>
    </hs:secPr>
    
    <!-- Section content -->
    <hp:p paraShapeIDRef="0" styleIDRef="0">
        <hp:run charShapeIDRef="0">
            <hp:t>Section content here</hp:t>
        </hp:run>
    </hp:p>
</hs:sec>
```

**XML Child Elements**:
| Element | XML Tag | Namespace | Description |
|---------|---------|-----------|-------------|
| Section Properties | `<hs:secPr>` | hs | Page layout and formatting |
| Page Size | `<hs:pageSz>` | hs | Page dimensions |
| Page Margins | `<hs:pageMargin>` | hs | Page margin settings |
| Landscape | `<hs:landscape>` | hs | Page orientation |
| Columns | `<hs:cols>` | hs | Column layout specification |
| Paragraphs | `<hp:p>` | hp | Content paragraphs |
| Objects | Various | hp | Tables, images, drawings, etc. |

**Properties**:
- Page dimensions in HWPML units (1/7200 inch)
- Margin settings for all sides plus header/footer
- Multi-column layout support
- Section break behavior control

## Paragraph Objects (Content Organization XML)

### PType (Paragraph)
**XML Element**: `<hp:p>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/paragraph`  
**Purpose**: Container for text and inline objects with paragraph-level formatting.

**XML Schema Definition**:
```xml
<hp:p paraShapeIDRef="0" styleIDRef="0" pageBreak="false" columnBreak="false">
    <hp:run charShapeIDRef="0">
        <hp:t>This is paragraph text content.</hp:t>
    </hp:run>
    <hp:run charShapeIDRef="1">
        <hp:t>Text with different formatting.</hp:t>
    </hp:run>
    <hp:lineseg/>
</hp:p>
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `paraShapeIDRef` | integer | No | Reference to paragraph shape definition |
| `styleIDRef` | integer | No | Reference to named style |
| `pageBreak` | boolean | No | Force page break before paragraph |
| `columnBreak` | boolean | No | Force column break before paragraph |

**XML Child Elements**:
| Element | XML Tag | Occurrence | Description |
|---------|---------|------------|-------------|
| Text Run | `<hp:run>` | 0..* | Text segments with character formatting |
| Line Segment | `<hp:lineseg>` | 0..* | Line break indicators |
| Control | `<hp:ctrl>` | 0..* | Control characters and special elements |

### RunType
**XML Element**: `<hp:run>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/paragraph`  
**Parent Element**: `<hp:p>`  
**Purpose**: Sequence of text with consistent character formatting.

**XML Schema Definition**:
```xml
<hp:run charShapeIDRef="0">
    <hp:t>Regular text content</hp:t>
    <hp:char code="9" val="&#9;"/>  <!-- Tab character -->
    <hp:nbspace/>  <!-- Non-breaking space -->
    <hp:lineBreak/>  <!-- Manual line break -->
    <hp:bookmark name="bookmark1"/>
    <hp:markpenBegin/>
    <hp:t>Highlighted text</hp:t>
    <hp:markpenEnd/>
</hp:run>
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `charShapeIDRef` | integer | No | Reference to character shape definition |

**XML Child Elements**:
| Element | XML Tag | Attributes | Description |
|---------|---------|------------|-------------|
| Text Content | `<hp:t>` | - | Plain Unicode text |
| Special Character | `<hp:char>` | code, val | Special/control characters |
| Tab | `<hp:tab>` | - | Tab character |
| Non-breaking Space | `<hp:nbspace>` | - | Non-breaking space |
| Hyphen | `<hp:hypen>` | - | Hyphenation point |
| Fixed-width Space | `<hp:fwspace>` | - | Fixed-width space |
| Line Break | `<hp:lineBreak>` | - | Manual line break |
| Index Mark | `<hp:indexmark>` | keyword1, keyword2 | Index entry marker |
| Bookmark | `<hp:bookmark>` | name | Bookmark definition |
| Highlight Begin | `<hp:markpenBegin>` | - | Start of highlighted text |
| Highlight End | `<hp:markpenEnd>` | - | End of highlighted text |

### T (Text)
**XML Element**: `<hp:t>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/paragraph`  
**Parent Element**: `<hp:run>`  
**Purpose**: Plain text content within a text run.

**XML Schema Definition**:
```xml
<hp:t>This is plain Unicode text content.</hp:t>
```

**Content**: Unicode text string  
**Attributes**: None - contains pure text content  
**Character Encoding**: UTF-8/UTF-16 Unicode

### Char (Special Character)
**XML Element**: `<hp:char>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/paragraph`  
**Parent Element**: `<hp:run>`  
**Purpose**: Special characters requiring specific handling or rendering.

**XML Schema Definition**:
```xml
<hp:char code="9" val="&#9;"/>      <!-- Tab -->
<hp:char code="13" val="&#13;"/>    <!-- Carriage return -->
<hp:char code="8203" val="&#8203;"/> <!-- Zero-width space -->
<hp:char code="8364" val="€"/>      <!-- Euro symbol -->
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `code` | integer | Yes | Unicode character code point |
| `val` | string | No | Character value or replacement text |

**Special Character Categories**:
- **Control Characters**: Tab (9), Line Feed (10), Carriage Return (13)
- **Mathematical Symbols**: Various Unicode math symbols
- **Diacritical Marks**: Combining diacritical marks
- **Typography**: Em dash, en dash, various quotes
- **Currency**: Euro (€), Yen (¥), Pound (£), etc.
- **Zero-width Characters**: Zero-width space, joiner, non-joiner

## Table Objects (Tabular Data XML)

### TableType
**XML Element**: `<hp:tbl>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/paragraph`  
**Parent Element**: `<hp:p>` (within control element)  
**Purpose**: Structured tabular data with complex formatting capabilities.

**XML Schema Definition**:
```xml
<hp:tbl inlineTable="false" pageBreak="false" repeatHeader="false">
    <hp:shapeComponent>
        <hp:offset x="0" y="0"/>
        <hp:orgSz width="42520" height="12700"/>
        <hp:curSz width="42520" height="12700"/>
        <hp:flip horizontal="false" vertical="false"/>
        <hp:rotationInfo angle="0"/>
        <hp:renderingInfo>
            <hp:scaleRotateMatrixRef ref="0"/>
        </hp:renderingInfo>
    </hp:shapeComponent>
    
    <hp:cellzoneList>
        <hp:cellzone startRow="0" startCol="0" endRow="1" endCol="2" 
                     borderFillIDRef="1"/>
    </hp:cellzoneList>
    
    <hp:coldef>
        <hp:col width="14173"/>
        <hp:col width="14173"/>
        <hp:col width="14174"/>
    </hp:coldef>
    
    <hp:tr>
        <hp:tc colspan="1" rowspan="1" width="14173" height="6350" 
               header="false" hasMargin="false" protect="false" 
               editable="true" dirty="false" borderFillIDRef="1">
            <hp:p paraShapeIDRef="0">
                <hp:run charShapeIDRef="0">
                    <hp:t>Cell 1,1</hp:t>
                </hp:run>
            </hp:p>
        </hp:tc>
        <hp:tc colspan="1" rowspan="1" width="14173" height="6350" 
               header="false" hasMargin="false" protect="false" 
               editable="true" dirty="false" borderFillIDRef="1">
            <hp:p paraShapeIDRef="0">
                <hp:run charShapeIDRef="0">
                    <hp:t>Cell 1,2</hp:t>
                </hp:run>
            </hp:p>
        </hp:tc>
        <hp:tc colspan="1" rowspan="1" width="14174" height="6350" 
               header="false" hasMargin="false" protect="false" 
               editable="true" dirty="false" borderFillIDRef="1">
            <hp:p paraShapeIDRef="0">
                <hp:run charShapeIDRef="0">
                    <hp:t>Cell 1,3</hp:t>
                </hp:run>
            </hp:p>
        </hp:tc>
    </hp:tr>
    <!-- Additional rows -->
</hp:tbl>
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `inlineTable` | boolean | No | Whether table is inline or floating |
| `pageBreak` | boolean | No | Allow page breaks within table |
| `repeatHeader` | boolean | No | Repeat header rows on new pages |

**XML Child Elements**:
| Element | XML Tag | Occurrence | Description |
|---------|---------|------------|-------------|
| Shape Component | `<hp:shapeComponent>` | 1 | Table positioning and sizing |
| Cell Zone List | `<hp:cellzoneList>` | 0..1 | Cell grouping and styling |
| Column Definition | `<hp:coldef>` | 1 | Column width specifications |
| Table Row | `<hp:tr>` | 1..* | Individual table rows |

### TrType (Table Row)
**XML Element**: `<hp:tr>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/paragraph`  
**Parent Element**: `<hp:tbl>`  
**Purpose**: Horizontal grouping of table cells.

**XML Schema Definition**:
```xml
<hp:tr>
    <hp:tc colspan="1" rowspan="1" width="14173" height="6350" 
           header="false" hasMargin="false" protect="false" 
           editable="true" dirty="false" borderFillIDRef="1">
        <!-- Cell content -->
    </hp:tc>
    <!-- Additional cells -->
</hp:tr>
```

**XML Child Elements**:
| Element | XML Tag | Occurrence | Description |
|---------|---------|------------|-------------|
| Table Cell | `<hp:tc>` | 1..* | Individual table cells |

### TcType (Table Cell)
**XML Element**: `<hp:tc>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/paragraph`  
**Parent Element**: `<hp:tr>`  
**Purpose**: Individual table cell containing document content.

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `colspan` | integer | No | Number of columns spanned (default: 1) |
| `rowspan` | integer | No | Number of rows spanned (default: 1) |
| `width` | integer | No | Cell width in HWPML units |
| `height` | integer | No | Cell height in HWPML units |
| `header` | boolean | No | Whether cell is a header cell |
| `hasMargin` | boolean | No | Whether cell has custom margins |
| `protect` | boolean | No | Whether cell content is protected |
| `editable` | boolean | No | Whether cell content is editable |
| `dirty` | boolean | No | Whether cell content has been modified |
| `borderFillIDRef` | integer | No | Reference to border/fill definition |

**XML Child Elements**:
| Element | XML Tag | Occurrence | Description |
|---------|---------|------------|-------------|
| Cell Margin | `<hp:cellMargin>` | 0..1 | Custom cell margin specifications |
| Paragraph | `<hp:p>` | 0..* | Cell content paragraphs |
| Sub-table | `<hp:tbl>` | 0..* | Nested tables within cells |

## Media and Object Elements (XML Schema)

### PictureType
**XML Element**: `<hp:pic>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/paragraph`  
**Parent Element**: Control element within `<hp:p>`  
**Purpose**: Embedded or linked image content with positioning and formatting.

**XML Schema Definition**:
```xml
<hp:pic reverse="false" watermark="false" effect="REAL_PIC">
    <hp:shapeComponent>
        <hp:offset x="0" y="0"/>
        <hp:orgSz width="21260" height="14173"/>
        <hp:curSz width="21260" height="14173"/>
        <hp:flip horizontal="false" vertical="false"/>
        <hp:rotationInfo angle="0"/>
        <hp:renderingInfo>
            <hp:scaleRotateMatrixRef ref="0"/>
        </hp:renderingInfo>
    </hp:shapeComponent>
    
    <hp:img src="BIN0001.png" binaryItem="BIN0001" embedded="true">
        <hp:clip left="0" top="0" right="21260" bottom="14173"/>
        <hp:effects>
            <hp:shadow type="DROP_SHADOW" color="0" offsetX="283" offsetY="283" 
                       blurRadius="141" transparency="50"/>
        </hp:effects>
    </hp:img>
</hp:pic>
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `reverse` | boolean | No | Whether image is reversed/mirrored |
| `watermark` | boolean | No | Whether image is used as watermark |
| `effect` | enum | No | Image effect type: "REAL_PIC", "GRAY_SCALE", "BLACK_WHITE" |

**XML Child Elements**:
| Element | XML Tag | Occurrence | Description |
|---------|---------|------------|-------------|
| Shape Component | `<hp:shapeComponent>` | 1 | Positioning and transformation |
| Image | `<hp:img>` | 1 | Image source and specifications |

### ImgType
**XML Element**: `<hp:img>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/paragraph`  
**Parent Element**: `<hp:pic>`  
**Purpose**: Image source and display specifications.

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `src` | string | Yes | Image source file reference |
| `binaryItem` | string | No | Reference to binary data in package |
| `embedded` | boolean | No | Whether image is embedded or linked |

**XML Child Elements**:
| Element | XML Tag | Attributes | Description |
|---------|---------|------------|-------------|
| Clipping | `<hp:clip>` | left, top, right, bottom | Image clipping boundaries |
| Effects | `<hp:effects>` | - | Visual effects container |

### EquationType
**XML Element**: `<hp:equation>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/paragraph`  
**Parent Element**: Control element within `<hp:p>`  
**Purpose**: Mathematical expressions and formulas.

**XML Schema Definition**:
```xml
<hp:equation name="Equation1" font="Cambria Math" fontSize="1100" 
             textColor="0" baseLine="BASELINE">
    <hp:shapeComponent>
        <hp:offset x="0" y="0"/>
        <hp:orgSz width="2835" height="1134"/>
        <hp:curSz width="2835" height="1134"/>
        <hp:flip horizontal="false" vertical="false"/>
        <hp:rotationInfo angle="0"/>
    </hp:shapeComponent>
    
    <hp:script><![CDATA[{x = {-b pm sqrt {b^2 - 4 a c}} over {2 a}}]]></hp:script>
</hp:equation>
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `name` | string | No | Equation name or identifier |
| `font` | string | No | Font used for equation rendering |
| `fontSize` | integer | No | Base font size in HWPML units |
| `textColor` | integer | No | Text color as RGB integer |
| `baseLine` | enum | No | Baseline alignment: "BASELINE", "TOP", "CENTER", "BOTTOM" |

**XML Child Elements**:
| Element | XML Tag | Content | Description |
|---------|---------|---------|-------------|
| Shape Component | `<hp:shapeComponent>` | - | Positioning and sizing |
| Script | `<hp:script>` | CDATA | Mathematical expression markup |

**Mathematical Markup Features**:
- Fraction notation: `{numerator} over {denominator}`
- Radicals: `sqrt {expression}`, `nroot {n} {expression}`
- Superscripts/Subscripts: `base^{super}`, `base_{sub}`
- Summation: `sum from {start} to {end} {expression}`
- Integration: `int from {start} to {end} {expression}`
- Matrix notation: `matrix { a # b ## c # d }`

### OLEType
**XML Element**: `<hp:ole>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/paragraph`  
**Parent Element**: Control element within `<hp:p>`  
**Purpose**: Embedded objects from external applications.

**XML Schema Definition**:
```xml
<hp:ole progId="Excel.Sheet.12" drawAspect="CONTENT">
    <hp:shapeComponent>
        <hp:offset x="0" y="0"/>
        <hp:orgSz width="28346" height="19843"/>
        <hp:curSz width="28346" height="19843"/>
        <hp:flip horizontal="false" vertical="false"/>
        <hp:rotationInfo angle="0"/>
    </hp:shapeComponent>
    
    <hp:img src="BIN0002.png" binaryItem="BIN0002" embedded="true"/>
    
    <!-- Binary OLE data follows -->
    <hp:oleData>
        <![CDATA[... base64 encoded OLE object data ...]]>
    </hp:oleData>
</hp:ole>
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `progId` | string | Yes | Program identifier for the embedded object |
| `drawAspect` | enum | No | Display aspect: "CONTENT", "ICON", "THUMBNAIL" |

**XML Child Elements**:
| Element | XML Tag | Content | Description |
|---------|---------|---------|-------------|
| Shape Component | `<hp:shapeComponent>` | - | Positioning and sizing |
| Preview Image | `<hp:img>` | - | Preview/thumbnail image |
| OLE Data | `<hp:oleData>` | CDATA | Base64-encoded binary object data |

**Common Program IDs**:
- `Excel.Sheet.12`: Microsoft Excel worksheet
- `PowerPoint.Slide.12`: Microsoft PowerPoint slide
- `Word.Document.12`: Microsoft Word document
- `Visio.Drawing.15`: Microsoft Visio drawing
- `AcroExch.Document`: Adobe PDF document

## Form Controls

### EditType
**XML Element**: `<hp:edit>`  
**Purpose**: Text input controls for data entry.

**Attributes**:
- `multiLine`: Whether control accepts multiple lines
- `password`: Whether to mask input (password field)
- `maxLength`: Maximum input length
- `scrollBars`: Scroll bar display options
- `tabStop`: Whether control participates in tab order
- `fieldName`: Field name for form processing
- `value`: Default or current value

**Validation**: Supports input validation rules and formatting constraints.

### ComboBoxType
**XML Element**: `<hp:combobox>`  
**Purpose**: Dropdown selection controls.

**Attributes**:
- `dropDown`: Whether to show dropdown list
- `editable`: Whether custom values are allowed
- `listItemCount`: Number of visible list items
- `selectedIndex`: Currently selected item index

**Child Elements**:
- `listItem`: Individual dropdown options

### ListBoxType
**XML Element**: `<hp:listbox>`  
**Purpose**: Multi-selection list controls.

**Attributes**:
- `multiSelect`: Whether multiple selection is allowed
- `size`: Number of visible items
- `listItemCount`: Total number of items

**Child Elements**:
- `listItem`: Individual list options

### AbstractButtonObjectType
**Purpose**: Base class for button controls including radio buttons, checkboxes, and command buttons.

**Common Attributes**:
- `caption`: Button text or label
- `value`: Button value or state
- `checked`: Whether button is selected/checked
- `groupName`: Radio button group identifier

## Drawing Objects

### LineShapeType
**XML Element**: `<hp:line>`  
**Purpose**: Simple line drawing objects.

**Attributes**:
- `startPos`: Line starting position
- `endPos`: Line ending position
- `lineType`: Line style and weight

### RectangleType
**XML Element**: `<hp:rect>`  
**Purpose**: Rectangle and square drawing objects.

**Attributes**:
- `round`: Corner rounding radius

### EllipseType
**XML Element**: `<hp:ellipse>`  
**Purpose**: Circle and ellipse drawing objects.

**Properties**: Standard shape positioning and styling attributes.

### ArcType
**XML Element**: `<hp:arc>`  
**Purpose**: Arc and curved line segments.

**Attributes**:
- `startAngle`: Arc starting angle
- `endAngle`: Arc ending angle
- `clockwise`: Arc direction

### PolygonType
**XML Element**: `<hp:polygon>`  
**Purpose**: Multi-sided polygon shapes.

**Child Elements**:
- `pt`: Polygon vertex points

### CurveType
**XML Element**: `<hp:curve>`  
**Purpose**: Bezier curves and complex paths.

**Child Elements**:
- `pt`: Curve control points

### ConnectLineType
**XML Element**: `<hp:connectLine>`  
**Purpose**: Connector lines between shapes.

**Attributes**:
- `startShape`: Starting shape reference
- `endShape`: Ending shape reference
- `startPos`: Connection point on start shape
- `endPos`: Connection point on end shape

## Field Objects (Dynamic Content)

### FieldBegin/FieldEnd
**XML Elements**: `<hp:fieldBegin>`, `<hp:fieldEnd>`  
**Purpose**: Delimit field content for dynamic data.

**Field Types**:
- Page numbering
- Date and time
- Document properties
- Cross-references
- Mail merge fields

### PageNumCtrlType
**XML Element**: `<hp:pageNumCtrl>`  
**Purpose**: Page number display and formatting.

**Attributes**:
- `pageNumFormat`: Number format (decimal, roman, etc.)
- `sideName`: Side identifier for different page types

### BookmarkType
**XML Element**: `<hp:bookmark>`  
**Purpose**: Named locations for cross-referencing and navigation.

**Attributes**:
- `name`: Bookmark name/identifier

## Layout and Positioning Objects

### PositionType
**Purpose**: Defines object positioning within document flow.

**Positioning Types**:
- Inline with text
- Floating with text wrapping
- Absolute positioning
- Relative to page, paragraph, or other objects

### InsideMarginType
**XML Element**: `<hp:insideMargin>`  
**Purpose**: Margin specifications for floating objects.

### CellMarginType
**XML Element**: `<hp:cellMargin>`  
**Purpose**: Margin specifications for table cells.

### TextMarginType
**XML Element**: `<hp:textMargin>`  
**Purpose**: Text margins within objects and containers.

## Effects and Visual Enhancement

### EffectsType
**XML Element**: `<hp:effects>`  
**Purpose**: Visual effects for images and objects.

**Effect Types**:
- `shadow`: Drop shadow effects
- `glow`: Glow and halo effects
- `reflection`: Mirror reflection effects
- `softEdge`: Soft edge blurring

### ShadowType
**XML Element**: `<hp:shadow>`  
**Purpose**: Shadow effect specifications.

**Attributes**:
- `offsetX`, `offsetY`: Shadow offset distances
- `blurRadius`: Shadow blur amount
- `color`: Shadow color
- `transparency`: Shadow opacity

### GlowType
**XML Element**: `<hp:glow>`  
**Purpose**: Glow effect around objects.

### ReflectionType
**XML Element**: `<hp:reflection>`  
**Purpose**: Mirror reflection effects below objects.

## Core Data Types (XML Schema Primitives)

### PointType
**XML Usage**: Attribute pairs in various elements  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/core`  
**Purpose**: 2D coordinate specifications in HWPML units.

**XML Schema Definition**:
```xml
<!-- Used as attributes in elements -->
<hp:offset x="1417" y="2835"/>
<hp:orgSz width="21260" height="14173"/>
<hp:curSz width="21260" height="14173"/>
```

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `x` | integer | Yes | Horizontal coordinate in HWPML units |
| `y` | integer | Yes | Vertical coordinate in HWPML units |

**Unit System**: HWPML uses 1/7200 inch as base unit (1 inch = 7200 HWPML units)

### RGBType
**XML Element**: `<hp:rgb>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/paragraph`  
**Purpose**: RGB color specifications for display.

**XML Schema Definition**:
```xml
<hp:rgb r="255" g="128" b="64"/>
```

**XML Attributes**:
| Attribute | Type | Required | Range | Description |
|-----------|------|----------|-------|-------------|
| `r` | integer | Yes | 0-255 | Red component |
| `g` | integer | Yes | 0-255 | Green component |
| `b` | integer | Yes | 0-255 | Blue component |

### CMYKType
**XML Element**: `<hp:cmyk>`  
**XML Namespace**: `http://www.hancom.co.kr/hwpml/2011/paragraph`  
**Purpose**: CMYK color specifications for print.

**XML Schema Definition**:
```xml
<hp:cmyk c="100" m="50" y="0" k="25"/>
```

**XML Attributes**:
| Attribute | Type | Required | Range | Description |
|-----------|------|----------|-------|-------------|
| `c` | integer | Yes | 0-100 | Cyan percentage |
| `m` | integer | Yes | 0-100 | Magenta percentage |
| `y` | integer | Yes | 0-100 | Yellow percentage |
| `k` | integer | Yes | 0-100 | Black (Key) percentage |

### MatrixType
**XML Usage**: Transformation matrix attributes  
**Purpose**: 2D transformation matrices for object positioning and scaling.

**XML Schema Definition**:
```xml
<hp:matrix e1="1.0" e2="0.0" e3="0.0" e4="1.0" e5="0.0" e6="0.0"/>
```

**Matrix Elements**:
- `e1`: X-axis scaling factor
- `e2`: X-axis skewing factor  
- `e3`: Y-axis skewing factor
- `e4`: Y-axis scaling factor
- `e5`: X-axis translation
- `e6`: Y-axis translation

### ImageType
**XML Usage**: Various image-related elements  
**Purpose**: Image data and format specifications.

**Supported Image Formats**:
- **PNG**: Portable Network Graphics (.png)
- **JPEG**: Joint Photographic Experts Group (.jpg, .jpeg)
- **GIF**: Graphics Interchange Format (.gif)
- **BMP**: Windows Bitmap (.bmp)
- **TIFF**: Tagged Image File Format (.tif, .tiff)
- **SVG**: Scalable Vector Graphics (.svg)

## Document Package Objects (Container XML)

### PackageType
**XML Element**: `<package>`  
**XML Namespace**: `http://www.idpf.org/2007/opf/`  
**File Location**: `Contents/content.hpf`  
**Purpose**: OPF package definition for document structure (EPUB-based).

**XML Schema Definition**:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" 
         xmlns:dc="http://purl.org/dc/elements/1.1/" 
         unique-identifier="BookId" version="2.0">
         
    <metadata>
        <dc:title>Document Title</dc:title>
        <dc:creator>Author Name</dc:creator>
        <dc:identifier id="BookId">unique-doc-id</dc:identifier>
        <dc:language>ko</dc:language>
        <meta name="HWPX-version" content="1.31"/>
    </metadata>
    
    <manifest>
        <item id="header" href="header.xml" media-type="application/xml"/>
        <item id="section0" href="section0.xml" media-type="application/xml"/>
        <item id="BIN0001" href="../BinData/BIN0001.png" media-type="image/png"/>
    </manifest>
    
    <spine>
        <itemref idref="section0"/>
    </spine>
</package>
```

**XML Child Elements**:
| Element | XML Tag | Namespace | Description |
|---------|---------|-----------|-------------|
| Metadata | `<metadata>` | - | Document metadata container |
| Manifest | `<manifest>` | - | File listing container |
| Spine | `<spine>` | - | Reading order specification |

### ManifestType
**XML Element**: `<manifest>`  
**XML Namespace**: `urn:oasis:names:tc:opendocument:xmlns:manifest:1.0`  
**File Location**: `META-INF/manifest.xml`  
**Purpose**: Complete file listing for document package (ODF-based).

**XML Schema Definition**:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<manifest xmlns="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
    <file-entry media-type="application/vnd.hancom.hwpx" full-path="/"/>
    <file-entry media-type="application/xml" full-path="Contents/content.hpf"/>
    <file-entry media-type="application/xml" full-path="Contents/header.xml"/>
    <file-entry media-type="application/xml" full-path="Contents/section0.xml"/>
    <file-entry media-type="image/png" full-path="BinData/BIN0001.png"/>
    <file-entry media-type="image/jpeg" full-path="BinData/BIN0002.jpg"/>
</manifest>
```

**XML Child Elements**:
| Element | XML Tag | Attributes | Description |
|---------|---------|------------|-------------|
| File Entry | `<file-entry>` | media-type, full-path | Individual file registry |

### ItemType
**XML Element**: `<item>`  
**XML Namespace**: `http://www.idpf.org/2007/opf/`  
**Parent Element**: `<manifest>` (in content.hpf)  
**Purpose**: Individual file entry in content manifest.

**XML Attributes**:
| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| `id` | string | Yes | Unique item identifier |
| `href` | string | Yes | File path relative to content.hpf |
| `media-type` | string | Yes | MIME type of the file |

### MetadataType
**XML Element**: `<metadata>`  
**XML Namespace**: Mix of Dublin Core and OPF  
**Parent Element**: `<package>`  
**Purpose**: Document metadata in Dublin Core format.

**XML Schema Definition**:
```xml
<metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:title>HWPX Document Title</dc:title>
    <dc:creator>Document Author</dc:creator>
    <dc:subject>Document Subject</dc:subject>
    <dc:description>Document description and summary</dc:description>
    <dc:publisher>Publishing Organization</dc:publisher>
    <dc:date>2023-12-07T10:30:00Z</dc:date>
    <dc:language>ko-KR</dc:language>
    <dc:identifier id="DocId">hwpx-doc-12345</dc:identifier>
    <meta name="HWPX-version" content="1.31"/>
    <meta name="HWPX-application" content="Hancom Office Hangul"/>
</metadata>
```

**Dublin Core Elements**:
| Element | XML Tag | Description |
|---------|---------|-------------|
| Title | `<dc:title>` | Document title |
| Creator | `<dc:creator>` | Document author/creator |
| Subject | `<dc:subject>` | Document subject/topic |
| Description | `<dc:description>` | Document description |
| Publisher | `<dc:publisher>` | Publishing information |
| Date | `<dc:date>` | Creation/modification dates (ISO 8601) |
| Language | `<dc:language>` | Document language (RFC 3066) |
| Identifier | `<dc:identifier>` | Unique document identifier |

## XML Validation and Schema Information

### Schema Validation
HWPX documents follow multiple XML schema standards:

1. **OWPML Schema**: Custom schema for document content
2. **OPF Schema**: EPUB Open Packaging Format for container
3. **ODF Schema**: OpenDocument Format for manifest
4. **Dublin Core**: Metadata standards

### Namespace Validation
All HWPX XML files must declare appropriate namespaces:

```xml
<!-- Header file namespace declarations -->
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"
         xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
         xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
         xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">

<!-- Section file namespace declarations -->  
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">

<!-- Container file namespace declarations -->
<container xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
```

### Character Encoding
All HWPX XML files use UTF-8 encoding:
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
```

This comprehensive XML-focused reference provides detailed specifications for every major object type in the HWPX document model, emphasizing XML structure, namespaces, schemas, and validation requirements.