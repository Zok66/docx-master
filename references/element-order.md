# OOXML 元素顺序表（自动生成，勿手改）

> 由 `pro/scripts/element_order.py --doc` 从 `pro/schemas/*.xsd` 生成。
> XSD 更新后重跑该命令即可刷新，**不要手工编辑本文件**。

## 为什么顺序是硬约束

OOXML 用 `xsd:sequence` 定义子元素：顺序不是风格问题，而是 schema 的一部分。
顺序错的后果不是「Word 给个警告」，而是 `Schemas validity error` /
「文件内容有问题，Word 无法打开」。因此**任何插入操作都要先决定落点**：

1. 在当前子元素序列里找到「应该排在自己前面的那个元素」，插到它后面；
2. 若前面没有更早的元素，插到第一个位置；
3. 插到末尾前先确认没有「必须收尾」的元素（如 `w:sectPr` 必须留在 `w:body` 最后）。

`apply` 引擎与 `extend` 扩展包都已按本表自动排序（`settings.elements`、
`sectPr` 子元素、`docProps` 子元素等），只有走 `edit` 的自由 `xml` 片段
或自写脚本时才需要人工对照本表。

## 记号

| 记号 | 含义 |
|---|---|
| `name` | 必需，恰好一次 |
| `name?` | 可选（`minOccurs="0"`） |
| `name+` | 一次或多次（`maxOccurs="unbounded"`） |
| `name?*` / `name*` | 可选且可重复 |
| `name{0,6}` | 有上限的重复（如 `sectPr` 里页眉页脚引用最多 6 次） |
| `(a \| b)` | `xsd:choice`：这一组里**只能选一个** |
| `<xsd:any>` | 通配符槽位（扩展点，见 `ooxml-extension-points.md`） |
| `[GroupName]` | group 自引用；内容已在前面就地展开 |

代码块里的换行只是排版，**顺序沿列连续**。

## 速查：最容易插错的三处

1. **`w:sectPr`** —— `headerReference` / `footerReference` 必须排在最前；
   改页边距（`pgMar`）时要落在 `pgSz` 之后、`pgBorders` 之前。
2. **`w:body`** —— 段落的 `sectPr` 永远是 body 的**最后一个**子元素。
3. **`c:dLbls`** —— `showXxx` 系列与 `txPr` 同在 `maxOccurs=1` 的 choice 里：
   二者**不能并存**，数据标签的字号请改放 `c:catAx` / `c:valAx` 的 `c:txPr`。

## 怎么读必填项

**没有 `?` 的就是必需**（`minOccurs=1`）。造元素之前先扫一遍必需项，能省掉一轮
「not expected / missing required element」的往返。典型例子：

- `CT_MailMerge` → `dataType` 必需（只写 `mainDocumentType` + `destination` 会失败）。
- `CT_DocPartPr` → `name` 必需。
- `CT_Chart` → `plotArea` 必需（`autoTitleDeleted`、`pivotFmts` 等反而可选）。
- `CT_SectPr` → 一个必需项都没有：整节属性全部可选。

遇到 `(a | b | c)` 时先确认「这一组里到底要选哪个」，再确认它排在哪几个兄弟之间 ——
choice 内部成员的**相对位置**同样受前后 sequence 约束。

## 正文与段落

段落是最常被改写的地方：pPr 与 rPr 里的任何新增子元素都必须落进上面的槽位。

### `CT_Body`

来源：`pro/schemas/wml.xsd`

```
    (((customXml | sdt | p? | tbl? | (proofErr? | permStart? | permEnd? | (bookmarkStart | bookmarkEnd | moveFromRangeStart | moveFromRangeEnd | moveToRangeStart | moveToRangeEnd | commentRangeStart | commentRangeEnd | customXmlInsRangeStart | customXmlInsRangeEnd | customXmlDelRangeStart | customXmlDelRangeEnd | customXmlMoveFromRangeStart | customXmlMoveFromRangeEnd | customXmlMoveToRangeStart | customXmlMoveToRangeEnd)? | ins? | del? | moveFrom | moveTo | (oMathPara | oMath)?)?)?)? | altChunk?)?
    sectPr?
```

可出现的子元素（去重）：`customXml` `sdt` `p` `tbl` `proofErr` `permStart` `permEnd` `bookmarkStart` `bookmarkEnd` `moveFromRangeStart` `moveFromRangeEnd` `moveToRangeStart` `moveToRangeEnd` `commentRangeStart` `commentRangeEnd` `customXmlInsRangeStart` `customXmlInsRangeEnd` `customXmlDelRangeStart` `customXmlDelRangeEnd` `customXmlMoveFromRangeStart` `customXmlMoveFromRangeEnd` `customXmlMoveToRangeStart` `customXmlMoveToRangeEnd` `ins` `del` `moveFrom` `moveTo` `oMathPara` `oMath` `altChunk` `sectPr`

### `CT_P`

来源：`pro/schemas/wml.xsd`

```
    pPr?
    ((customXml | smartTag | sdt | dir | bdo | r | (proofErr? | permStart? | permEnd? | (bookmarkStart | bookmarkEnd | moveFromRangeStart | moveFromRangeEnd | moveToRangeStart | moveToRangeEnd | commentRangeStart | commentRangeEnd | customXmlInsRangeStart | customXmlInsRangeEnd | customXmlDelRangeStart | customXmlDelRangeEnd | customXmlMoveFromRangeStart | customXmlMoveFromRangeEnd | customXmlMoveToRangeStart | customXmlMoveToRangeEnd)? | ins? | del? | moveFrom | moveTo | (oMathPara | oMath)?)?)? | fldSimple? | hyperlink | subDoc)?
```

可出现的子元素（去重）：`pPr` `customXml` `smartTag` `sdt` `dir` `bdo` `r` `proofErr` `permStart` `permEnd` `bookmarkStart` `bookmarkEnd` `moveFromRangeStart` `moveFromRangeEnd` `moveToRangeStart` `moveToRangeEnd` `commentRangeStart` `commentRangeEnd` `customXmlInsRangeStart` `customXmlInsRangeEnd` `customXmlDelRangeStart` `customXmlDelRangeEnd` `customXmlMoveFromRangeStart` `customXmlMoveFromRangeEnd` `customXmlMoveToRangeStart` `customXmlMoveToRangeEnd` `ins` `del` `moveFrom` `moveTo` `oMathPara` `oMath` `fldSimple` `hyperlink` `subDoc`

### `CT_PPr`

来源：`pro/schemas/wml.xsd`

```
    pStyle? keepNext? keepLines? pageBreakBefore? framePr? widowControl? numPr?
    suppressLineNumbers? pBdr? shd? tabs? suppressAutoHyphens? kinsoku? wordWrap? overflowPunct?
    topLinePunct? autoSpaceDE? autoSpaceDN? bidi? adjustRightInd? snapToGrid? spacing? ind?
    contextualSpacing? mirrorIndents? suppressOverlap? jc? textDirection? textAlignment?
    textboxTightWrap? outlineLvl? divId? cnfStyle? rPr? sectPr? pPrChange?
```

### `CT_PPrBase`

来源：`pro/schemas/wml.xsd`

```
    pStyle? keepNext? keepLines? pageBreakBefore? framePr? widowControl? numPr?
    suppressLineNumbers? pBdr? shd? tabs? suppressAutoHyphens? kinsoku? wordWrap? overflowPunct?
    topLinePunct? autoSpaceDE? autoSpaceDN? bidi? adjustRightInd? snapToGrid? spacing? ind?
    contextualSpacing? mirrorIndents? suppressOverlap? jc? textDirection? textAlignment?
    textboxTightWrap? outlineLvl? divId? cnfStyle?
```

### `CT_RPr`

来源：`pro/schemas/wml.xsd`

```
    (rStyle | rFonts | b | bCs | i | iCs | caps | smallCaps | strike | dstrike | outline | shadow | emboss | imprint | noProof | snapToGrid | vanish | webHidden | color | spacing | w | kern | position | sz | szCs | highlight | u | effect | bdr | shd | fitText | vertAlign | rtl | cs | em | lang | eastAsianLayout | specVanish | oMath)?
    rPrChange?
```

可出现的子元素（去重）：`rStyle` `rFonts` `b` `bCs` `i` `iCs` `caps` `smallCaps` `strike` `dstrike` `outline` `shadow` `emboss` `imprint` `noProof` `snapToGrid` `vanish` `webHidden` `color` `spacing` `w` `kern` `position` `sz` `szCs` `highlight` `u` `effect` `bdr` `shd` `fitText` `vertAlign` `rtl` `cs` `em` `lang` `eastAsianLayout` `specVanish` `oMath` `rPrChange`

### `CT_ParaRPr`

来源：`pro/schemas/wml.xsd`

```
    ins? del? moveFrom? moveTo?
    (rStyle | rFonts | b | bCs | i | iCs | caps | smallCaps | strike | dstrike | outline | shadow | emboss | imprint | noProof | snapToGrid | vanish | webHidden | color | spacing | w | kern | position | sz | szCs | highlight | u | effect | bdr | shd | fitText | vertAlign | rtl | cs | em | lang | eastAsianLayout | specVanish | oMath)?
    rPrChange?
```

可出现的子元素（去重）：`ins` `del` `moveFrom` `moveTo` `rStyle` `rFonts` `b` `bCs` `i` `iCs` `caps` `smallCaps` `strike` `dstrike` `outline` `shadow` `emboss` `imprint` `noProof` `snapToGrid` `vanish` `webHidden` `color` `spacing` `w` `kern` `position` `sz` `szCs` `highlight` `u` `effect` `bdr` `shd` `fitText` `vertAlign` `rtl` `cs` `em` `lang` `eastAsianLayout` `specVanish` `oMath` `rPrChange`

### `CT_R`

来源：`pro/schemas/shared-math.xsd`

```
    rPr? rPr?
    ((br | t | contentPart | delText | instrText | delInstrText | noBreakHyphen | softHyphen? | dayShort? | monthShort? | yearShort? | dayLong? | monthLong? | yearLong? | annotationRef? | footnoteRef? | endnoteRef? | separator? | continuationSeparator? | sym? | pgNum? | cr? | tab? | object | pict | fldChar | ruby | footnoteReference | endnoteReference | commentReference | drawing | ptab? | lastRenderedPageBreak?) | t?)?
```

可出现的子元素（去重）：`rPr` `br` `t` `contentPart` `delText` `instrText` `delInstrText` `noBreakHyphen` `softHyphen` `dayShort` `monthShort` `yearShort` `dayLong` `monthLong` `yearLong` `annotationRef` `footnoteRef` `endnoteRef` `separator` `continuationSeparator` `sym` `pgNum` `cr` `tab` `object` `pict` `fldChar` `ruby` `footnoteReference` `endnoteReference` `commentReference` `drawing` `ptab` `lastRenderedPageBreak`

### `CT_TrackChange`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_PPrChange`

来源：`pro/schemas/wml.xsd`

```
    pPr
```

### `CT_RPrChange`

来源：`pro/schemas/wml.xsd`

```
    rPr
```

## 节与页面

sectPr 的 headerReference/footerReference 必须排在最前，type 决定分节方式。

### `CT_SectPr`

来源：`pro/schemas/wml.xsd`

```
    (headerReference? | footerReference?){0,6} footnotePr? endnotePr? type? pgSz? pgMar?
    paperSrc? pgBorders? lnNumType? pgNumType? cols? formProt? vAlign? noEndnote? titlePg?
    textDirection? bidi? rtlGutter? docGrid? printerSettings? sectPrChange?
```

可出现的子元素（去重）：`headerReference` `footerReference` `footnotePr` `endnotePr` `type` `pgSz` `pgMar` `paperSrc` `pgBorders` `lnNumType` `pgNumType` `cols` `formProt` `vAlign` `noEndnote` `titlePg` `textDirection` `bidi` `rtlGutter` `docGrid` `printerSettings` `sectPrChange`

### `CT_PageMar`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_PageSz`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_Columns`

来源：`pro/schemas/wml.xsd`

```
    col{1,45}
```

### `CT_PageBorders`

来源：`pro/schemas/wml.xsd`

```
    top? left? bottom? right?
```

### `CT_PageNumber`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_LineNumber`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_DocGrid`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_HdrFtr`

来源：`pro/schemas/wml.xsd`

```
    (((customXml | sdt | p? | tbl? | (proofErr? | permStart? | permEnd? | (bookmarkStart | bookmarkEnd | moveFromRangeStart | moveFromRangeEnd | moveToRangeStart | moveToRangeEnd | commentRangeStart | commentRangeEnd | customXmlInsRangeStart | customXmlInsRangeEnd | customXmlDelRangeStart | customXmlDelRangeEnd | customXmlMoveFromRangeStart | customXmlMoveFromRangeEnd | customXmlMoveToRangeStart | customXmlMoveToRangeEnd)? | ins? | del? | moveFrom | moveTo | (oMathPara | oMath)?)?)?)? | altChunk?)+
```

可出现的子元素（去重）：`customXml` `sdt` `p` `tbl` `proofErr` `permStart` `permEnd` `bookmarkStart` `bookmarkEnd` `moveFromRangeStart` `moveFromRangeEnd` `moveToRangeStart` `moveToRangeEnd` `commentRangeStart` `commentRangeEnd` `customXmlInsRangeStart` `customXmlInsRangeEnd` `customXmlDelRangeStart` `customXmlDelRangeEnd` `customXmlMoveFromRangeStart` `customXmlMoveFromRangeEnd` `customXmlMoveToRangeStart` `customXmlMoveToRangeEnd` `ins` `del` `moveFrom` `moveTo` `oMathPara` `oMath` `altChunk`

## 表格

tcPr 里边框/底纹/宽度有固定次序；tblPr 的 tblStyle 必须第一个出现。

### `CT_Tbl`

来源：`pro/schemas/wml.xsd`

```
    (bookmarkStart | bookmarkEnd | moveFromRangeStart | moveFromRangeEnd | moveToRangeStart | moveToRangeEnd | commentRangeStart | commentRangeEnd | customXmlInsRangeStart | customXmlInsRangeEnd | customXmlDelRangeStart | customXmlDelRangeEnd | customXmlMoveFromRangeStart | customXmlMoveFromRangeEnd | customXmlMoveToRangeStart | customXmlMoveToRangeEnd)?
    tblPr tblGrid
    (tr? | customXml | sdt | (proofErr? | permStart? | permEnd? | (bookmarkStart | bookmarkEnd | moveFromRangeStart | moveFromRangeEnd | moveToRangeStart | moveToRangeEnd | commentRangeStart | commentRangeEnd | customXmlInsRangeStart | customXmlInsRangeEnd | customXmlDelRangeStart | customXmlDelRangeEnd | customXmlMoveFromRangeStart | customXmlMoveFromRangeEnd | customXmlMoveToRangeStart | customXmlMoveToRangeEnd)? | ins? | del? | moveFrom | moveTo | (oMathPara | oMath)?)?)?
```

可出现的子元素（去重）：`bookmarkStart` `bookmarkEnd` `moveFromRangeStart` `moveFromRangeEnd` `moveToRangeStart` `moveToRangeEnd` `commentRangeStart` `commentRangeEnd` `customXmlInsRangeStart` `customXmlInsRangeEnd` `customXmlDelRangeStart` `customXmlDelRangeEnd` `customXmlMoveFromRangeStart` `customXmlMoveFromRangeEnd` `customXmlMoveToRangeStart` `customXmlMoveToRangeEnd` `tblPr` `tblGrid` `tr` `customXml` `sdt` `proofErr` `permStart` `permEnd` `ins` `del` `moveFrom` `moveTo` `oMathPara` `oMath`

### `CT_TblPr`

来源：`pro/schemas/wml.xsd`

```
    tblStyle? tblpPr? tblOverlap? bidiVisual? tblStyleRowBandSize? tblStyleColBandSize? tblW?
    jc? tblCellSpacing? tblInd? tblBorders? shd? tblLayout? tblCellMar? tblLook? tblCaption?
    tblDescription? tblPrChange?
```

### `CT_TblPrBase`

来源：`pro/schemas/wml.xsd`

```
    tblStyle? tblpPr? tblOverlap? bidiVisual? tblStyleRowBandSize? tblStyleColBandSize? tblW?
    jc? tblCellSpacing? tblInd? tblBorders? shd? tblLayout? tblCellMar? tblLook? tblCaption?
    tblDescription?
```

### `CT_TblPrEx`

来源：`pro/schemas/wml.xsd`

```
    tblW? jc? tblCellSpacing? tblInd? tblBorders? shd? tblLayout? tblCellMar? tblLook?
    tblPrExChange?
```

### `CT_TrPr`

来源：`pro/schemas/wml.xsd`

```
    (cnfStyle? | divId? | gridBefore? | gridAfter? | wBefore? | wAfter? | cantSplit? | trHeight? | tblHeader? | tblCellSpacing? | jc? | hidden?)+
    ins? del? trPrChange?
```

可出现的子元素（去重）：`cnfStyle` `divId` `gridBefore` `gridAfter` `wBefore` `wAfter` `cantSplit` `trHeight` `tblHeader` `tblCellSpacing` `jc` `hidden` `ins` `del` `trPrChange`

### `CT_Tc`

来源：`pro/schemas/wml.xsd`

```
    tcPr?
    (((customXml | sdt | p? | tbl? | (proofErr? | permStart? | permEnd? | (bookmarkStart | bookmarkEnd | moveFromRangeStart | moveFromRangeEnd | moveToRangeStart | moveToRangeEnd | commentRangeStart | commentRangeEnd | customXmlInsRangeStart | customXmlInsRangeEnd | customXmlDelRangeStart | customXmlDelRangeEnd | customXmlMoveFromRangeStart | customXmlMoveFromRangeEnd | customXmlMoveToRangeStart | customXmlMoveToRangeEnd)? | ins? | del? | moveFrom | moveTo | (oMathPara | oMath)?)?)?)? | altChunk?)+
```

可出现的子元素（去重）：`tcPr` `customXml` `sdt` `p` `tbl` `proofErr` `permStart` `permEnd` `bookmarkStart` `bookmarkEnd` `moveFromRangeStart` `moveFromRangeEnd` `moveToRangeStart` `moveToRangeEnd` `commentRangeStart` `commentRangeEnd` `customXmlInsRangeStart` `customXmlInsRangeEnd` `customXmlDelRangeStart` `customXmlDelRangeEnd` `customXmlMoveFromRangeStart` `customXmlMoveFromRangeEnd` `customXmlMoveToRangeStart` `customXmlMoveToRangeEnd` `ins` `del` `moveFrom` `moveTo` `oMathPara` `oMath` `altChunk`

### `CT_TcPr`

来源：`pro/schemas/wml.xsd`

```
    cnfStyle? tcW? gridSpan? hMerge? vMerge? tcBorders? shd? noWrap? tcMar? textDirection?
    tcFitText? vAlign? hideMark? headers? (cellIns? | cellDel? | cellMerge?) tcPrChange?
```

可出现的子元素（去重）：`cnfStyle` `tcW` `gridSpan` `hMerge` `vMerge` `tcBorders` `shd` `noWrap` `tcMar` `textDirection` `tcFitText` `vAlign` `hideMark` `headers` `cellIns` `cellDel` `cellMerge` `tcPrChange`

### `CT_TblGrid`

来源：`pro/schemas/wml.xsd`

```
    gridCol? tblGridChange?
```

### `CT_TblWidth`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_TblBorders`

来源：`pro/schemas/wml.xsd`

```
    top? start? left? bottom? end? right? insideH? insideV?
```

### `CT_TblCellMar`

来源：`pro/schemas/wml.xsd`

```
    top? start? left? bottom? end? right?
```

### `CT_TblLayoutType`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_TblLook`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_TcBorders`

来源：`pro/schemas/wml.xsd`

```
    top? start? left? bottom? end? right? insideH? insideV? tl2br? tr2bl?
```

### `CT_Shd`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_VMerge`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_HMerge`

来源：`pro/schemas/wml.xsd`

```

```

## 修订与批注

修订类元素（ins/del/moveFrom/moveTo）包住的是内容，不是兄弟节点，插错会破坏树。

### `CT_Markup`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_TrackChangesView`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_Comment`

来源：`pro/schemas/pml.xsd`

```
    pos text extLst?
```

### `CT_Comments`

来源：`pro/schemas/sml.xsd`

```
    authors commentList extLst?
```

## 设置部件 w:settings

全库最长的 sequence：98 个子元素固定顺序，`apply` 的 `settings.elements` 会自动按此排序插入。CT_MailMerge 里的 `dataType` 是**必需项**——只写 mainDocumentType + destination 会直接校验失败。

### `CT_Settings`

来源：`pro/schemas/wml.xsd`

```
    writeProtection? view? zoom? removePersonalInformation? removeDateAndTime?
    doNotDisplayPageBoundaries? displayBackgroundShape? printPostScriptOverText?
    printFractionalCharacterWidth? printFormsData? embedTrueTypeFonts? embedSystemFonts?
    saveSubsetFonts? saveFormsData? mirrorMargins? alignBordersAndEdges?
    bordersDoNotSurroundHeader? bordersDoNotSurroundFooter? gutterAtTop? hideSpellingErrors?
    hideGrammaticalErrors? activeWritingStyle? proofState? formsDesign? attachedTemplate?
    linkStyles? stylePaneFormatFilter? stylePaneSortMethod? documentType? mailMerge?
    revisionView? trackRevisions? doNotTrackMoves? doNotTrackFormatting? documentProtection?
    autoFormatOverride? styleLockTheme? styleLockQFSet? defaultTabStop? autoHyphenation?
    consecutiveHyphenLimit? hyphenationZone? doNotHyphenateCaps? showEnvelope? summaryLength?
    clickAndTypeStyle? defaultTableStyle? evenAndOddHeaders? bookFoldRevPrinting?
    bookFoldPrinting? bookFoldPrintingSheets? drawingGridHorizontalSpacing?
    drawingGridVerticalSpacing? displayHorizontalDrawingGridEvery?
    displayVerticalDrawingGridEvery? doNotUseMarginsForDrawingGridOrigin?
    drawingGridHorizontalOrigin? drawingGridVerticalOrigin? doNotShadeFormData?
    noPunctuationKerning? characterSpacingControl? printTwoOnOne? strictFirstAndLastChars?
    noLineBreaksAfter? noLineBreaksBefore? savePreviewPicture? doNotValidateAgainstSchema?
    saveInvalidXml? ignoreMixedContent? alwaysShowPlaceholderText? doNotDemarcateInvalidXml?
    saveXmlDataOnly? useXSLTWhenSaving? saveThroughXslt? showXMLTags? alwaysMergeEmptyNamespace?
    updateFields? hdrShapeDefaults? footnotePr? endnotePr? compat? docVars? rsids? mathPr?
    attachedSchema? themeFontLang? clrSchemeMapping? doNotIncludeSubdocsInStats?
    doNotAutoCompressPictures? forceUpgrade? captions? readModeInkLockDown? smartTagType?
    schemaLibrary? shapeDefaults? doNotEmbedSmartTags? decimalSymbol? listSeparator?
```

### `CT_MailMerge`

来源：`pro/schemas/wml.xsd`

```
    mainDocumentType linkToQuery? dataType connectString? query? dataSource? headerSource?
    doNotSuppressBlankLines? destination? addressFieldName? mailSubject? mailAsAttachment?
    viewMergedData? activeRecord? checkErrors? odso?
```

### `CT_DocVars`

来源：`pro/schemas/wml.xsd`

```
    docVar?
```

### `CT_Compat`

来源：`pro/schemas/wml.xsd`

```
    useSingleBorderforContiguousCells? wpJustification? noTabHangInd? noLeading? spaceForUL?
    noColumnBalance? balanceSingleByteDoubleByteWidth? noExtraLineSpacing?
    doNotLeaveBackslashAlone? ulTrailSpace? doNotExpandShiftReturn? spacingInWholePoints?
    lineWrapLikeWord6? printBodyTextBeforeHeader? printColBlack? wpSpaceWidth?
    showBreaksInFrames? subFontBySize? suppressBottomSpacing? suppressTopSpacing?
    suppressSpacingAtTopOfPage? suppressTopSpacingWP? suppressSpBfAfterPgBrk?
    swapBordersFacingPages? convMailMergeEsc? truncateFontHeightsLikeWP6? mwSmallCaps?
    usePrinterMetrics? doNotSuppressParagraphBorders? wrapTrailSpaces? footnoteLayoutLikeWW8?
    shapeLayoutLikeWW8? alignTablesRowByRow? forgetLastTabAlignment? adjustLineHeightInTable?
    autoSpaceLikeWord95? noSpaceRaiseLower? doNotUseHTMLParagraphAutoSpacing?
    layoutRawTableWidth? layoutTableRowsApart? useWord97LineBreakRules? doNotBreakWrappedTables?
    doNotSnapToGridInCell? selectFldWithFirstOrLastChar? applyBreakingRules?
    doNotWrapTextWithPunct? doNotUseEastAsianBreakRules? useWord2002TableStyleRules?
    growAutofit? useFELayout? useNormalStyleForList? doNotUseIndentAsNumberingTabStop?
    useAltKinsokuLineBreakRules? allowSpaceOfSameStyleInTable? doNotSuppressIndentation?
    doNotAutofitConstrainedTables? autofitToFirstFixedWidthCell? underlineTabInNumList?
    displayHangulFixedWidth? splitPgBreakAndParaMark? doNotVertAlignCellWithSp?
    doNotBreakConstrainedForcedTable? doNotVertAlignInTxbx? useAnsiKerningPairs?
    cachedColBalance? compatSetting?
```

### `CT_WriteProtection`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_DocProtect`

来源：`pro/schemas/wml.xsd`

```

```

## 样式与编号

styles.xml / numbering.xml 的部件级顺序；lvl 内的 pPr/rPr 各自再走上面的段落表。

### `CT_Styles`

来源：`pro/schemas/wml.xsd`

```
    docDefaults? latentStyles? style?
```

### `CT_Style`

来源：`pro/schemas/dml-chart.xsd`

```

```

### `CT_LatentStyles`

来源：`pro/schemas/wml.xsd`

```
    lsdException?
```

### `CT_LsdException`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_Numbering`

来源：`pro/schemas/wml.xsd`

```
    numPicBullet? abstractNum? num? numIdMacAtCleanup?
```

### `CT_Num`

来源：`pro/schemas/wml.xsd`

```
    abstractNumId lvlOverride{0,9}
```

### `CT_AbstractNum`

来源：`pro/schemas/wml.xsd`

```
    nsid? multiLevelType? tmpl? name? styleLink? numStyleLink? lvl{0,9}
```

### `CT_Lvl`

来源：`pro/schemas/dml-chart.xsd`

```
    pt?
```

### `CT_NumPr`

来源：`pro/schemas/wml.xsd`

```
    ilvl? numId? numberingChange? ins?
```

### `CT_DecimalNumber`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_MultiLevelType`

来源：`pro/schemas/wml.xsd`

```

```

## 构建基块与内容控件

gallery 只能挂在 category 里，且取值受 CT_DocPartGallery 枚举约束。

### `CT_GlossaryDocument`

来源：`pro/schemas/wml.xsd`

```
    background? docParts?
```

### `CT_DocParts`

来源：`pro/schemas/wml.xsd`

```
    (docPart+)
```

可出现的子元素（去重）：`docPart`

### `CT_DocPart`

来源：`pro/schemas/wml.xsd`

```
    docPartPr? docPartBody?
```

### `CT_DocPartPr`

来源：`pro/schemas/wml.xsd`

```
    name style? category? types? behaviors? description? guid?
```

### `CT_DocPartName`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_DocPartCategory`

来源：`pro/schemas/wml.xsd`

```
    name gallery
```

### `CT_DocPartTypes`

来源：`pro/schemas/wml.xsd`

```
    (type+)
```

可出现的子元素（去重）：`type`

### `CT_DocPartBehavior`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_DocPartGallery`

来源：`pro/schemas/wml.xsd`

```

```

### `CT_SdtBlock`

来源：`pro/schemas/wml.xsd`

```
    sdtPr? sdtEndPr? sdtContent?
```

### `CT_SdtRun`

来源：`pro/schemas/wml.xsd`

```
    sdtPr? sdtEndPr? sdtContent?
```

### `CT_SdtPr`

来源：`pro/schemas/wml.xsd`

```
    rPr? alias? tag? id? lock? placeholder? temporary? showingPlcHdr? dataBinding? label?
    tabIndex?
    (equation | comboBox | date | docPartObj | docPartList | dropDownList | picture | richText | text | citation | group | bibliography)?
```

可出现的子元素（去重）：`rPr` `alias` `tag` `id` `lock` `placeholder` `temporary` `showingPlcHdr` `dataBinding` `label` `tabIndex` `equation` `comboBox` `date` `docPartObj` `docPartList` `dropDownList` `picture` `richText` `text` `citation` `group` `bibliography`

### `CT_SdtContentRun`

来源：`pro/schemas/wml.xsd`

```
    ((customXml | smartTag | sdt | dir | bdo | r | (proofErr? | permStart? | permEnd? | (bookmarkStart | bookmarkEnd | moveFromRangeStart | moveFromRangeEnd | moveToRangeStart | moveToRangeEnd | commentRangeStart | commentRangeEnd | customXmlInsRangeStart | customXmlInsRangeEnd | customXmlDelRangeStart | customXmlDelRangeEnd | customXmlMoveFromRangeStart | customXmlMoveFromRangeEnd | customXmlMoveToRangeStart | customXmlMoveToRangeEnd)? | ins? | del? | moveFrom | moveTo | (oMathPara | oMath)?)?)? | fldSimple? | hyperlink | subDoc)?
```

可出现的子元素（去重）：`customXml` `smartTag` `sdt` `dir` `bdo` `r` `proofErr` `permStart` `permEnd` `bookmarkStart` `bookmarkEnd` `moveFromRangeStart` `moveFromRangeEnd` `moveToRangeStart` `moveToRangeEnd` `commentRangeStart` `commentRangeEnd` `customXmlInsRangeStart` `customXmlInsRangeEnd` `customXmlDelRangeStart` `customXmlDelRangeEnd` `customXmlMoveFromRangeStart` `customXmlMoveFromRangeEnd` `customXmlMoveToRangeStart` `customXmlMoveToRangeEnd` `ins` `del` `moveFrom` `moveTo` `oMathPara` `oMath` `fldSimple` `hyperlink` `subDoc`

## 图表 chart

dLbls 的 Group_DLblShared 是 maxOccurs=1 的 choice —— 标了 showVal 就不能再跟 txPr。

### `CT_ChartSpace`

来源：`pro/schemas/dml-chart.xsd`

```
    date1904? lang? roundedCorners? style? clrMapOvr? pivotSource? protection? chart spPr? txPr?
    externalData? printSettings? userShapes? extLst?
```

### `CT_Chart`

来源：`pro/schemas/dml-chart.xsd`

```
    title? autoTitleDeleted? pivotFmts? view3D? floor? sideWall? backWall? plotArea legend?
    plotVisOnly? dispBlanksAs? showDLblsOverMax? extLst?
```

### `CT_PlotArea`

来源：`pro/schemas/dml-chart.xsd`

```
    layout?
    (areaChart | area3DChart | lineChart | line3DChart | stockChart | radarChart | scatterChart | pieChart | pie3DChart | doughnutChart | barChart | bar3DChart | ofPieChart | surfaceChart | surface3DChart | bubbleChart)+
    (valAx | catAx | dateAx | serAx)? dTable? spPr? extLst?
```

可出现的子元素（去重）：`layout` `areaChart` `area3DChart` `lineChart` `line3DChart` `stockChart` `radarChart` `scatterChart` `pieChart` `pie3DChart` `doughnutChart` `barChart` `bar3DChart` `ofPieChart` `surfaceChart` `surface3DChart` `bubbleChart` `valAx` `catAx` `dateAx` `serAx` `dTable` `spPr` `extLst`

### `CT_BarChart`

来源：`pro/schemas/dml-chart.xsd`

```
    barDir grouping? varyColors? ser? dLbls? gapWidth? overlap? serLines? axId{2,2} extLst?
```

### `CT_BarSer`

来源：`pro/schemas/dml-chart.xsd`

```
    idx order tx? spPr? invertIfNegative? pictureOptions? dPt? dLbls? trendline? errBars? cat?
    val? shape? extLst?
```

### `CT_LineChart`

来源：`pro/schemas/dml-chart.xsd`

```
    grouping varyColors? ser? dLbls? dropLines? hiLowLines? upDownBars? marker? smooth?
    axId{2,2} extLst?
```

### `CT_LineSer`

来源：`pro/schemas/dml-chart.xsd`

```
    idx order tx? spPr? marker? dPt? dLbls? trendline? errBars? cat? val? smooth? extLst?
```

### `CT_PieChart`

来源：`pro/schemas/dml-chart.xsd`

```
    varyColors? ser? dLbls? firstSliceAng? extLst?
```

### `CT_PieSer`

来源：`pro/schemas/dml-chart.xsd`

```
    idx order tx? spPr? explosion? dPt? dLbls? cat? val? extLst?
```

### `CT_CatAx`

来源：`pro/schemas/dml-chart.xsd`

```
    axId scaling delete? axPos majorGridlines? minorGridlines? title? numFmt? majorTickMark?
    minorTickMark? tickLblPos? spPr? txPr? crossAx (crosses | crossesAt)? auto? lblAlgn?
    lblOffset? tickLblSkip? tickMarkSkip? noMultiLvlLbl? extLst?
```

可出现的子元素（去重）：`axId` `scaling` `delete` `axPos` `majorGridlines` `minorGridlines` `title` `numFmt` `majorTickMark` `minorTickMark` `tickLblPos` `spPr` `txPr` `crossAx` `crosses` `crossesAt` `auto` `lblAlgn` `lblOffset` `tickLblSkip` `tickMarkSkip` `noMultiLvlLbl` `extLst`

### `CT_ValAx`

来源：`pro/schemas/dml-chart.xsd`

```
    axId scaling delete? axPos majorGridlines? minorGridlines? title? numFmt? majorTickMark?
    minorTickMark? tickLblPos? spPr? txPr? crossAx (crosses | crossesAt)? crossBetween?
    majorUnit? minorUnit? dispUnits? extLst?
```

可出现的子元素（去重）：`axId` `scaling` `delete` `axPos` `majorGridlines` `minorGridlines` `title` `numFmt` `majorTickMark` `minorTickMark` `tickLblPos` `spPr` `txPr` `crossAx` `crosses` `crossesAt` `crossBetween` `majorUnit` `minorUnit` `dispUnits` `extLst`

### `CT_DLbls`

来源：`pro/schemas/dml-chart.xsd`

```
    dLbl?
    (delete | numFmt? | spPr? | txPr? | dLblPos? | showLegendKey? | showVal? | showCatName? | showSerName? | showPercent? | showBubbleSize? | separator? | showLeaderLines? | leaderLines?)
    extLst?
```

可出现的子元素（去重）：`dLbl` `delete` `numFmt` `spPr` `txPr` `dLblPos` `showLegendKey` `showVal` `showCatName` `showSerName` `showPercent` `showBubbleSize` `separator` `showLeaderLines` `leaderLines` `extLst`

### `CT_DLbl`

来源：`pro/schemas/dml-chart.xsd`

```
    idx
    (delete | layout? | tx? | numFmt? | spPr? | txPr? | dLblPos? | showLegendKey? | showVal? | showCatName? | showSerName? | showPercent? | showBubbleSize? | separator?)
    extLst?
```

可出现的子元素（去重）：`idx` `delete` `layout` `tx` `numFmt` `spPr` `txPr` `dLblPos` `showLegendKey` `showVal` `showCatName` `showSerName` `showPercent` `showBubbleSize` `separator` `extLst`

### `CT_Legend`

来源：`pro/schemas/dml-chart.xsd`

```
    legendPos? legendEntry? layout? overlay? spPr? txPr? extLst?
```

### `CT_Title`

来源：`pro/schemas/dml-chart.xsd`

```
    tx? layout? overlay? spPr? txPr? extLst?
```

### `CT_Scaling`

来源：`pro/schemas/dml-chart.xsd`

```
    logBase? orientation? max? min? extLst?
```

### `CT_NumFmt`

来源：`pro/schemas/dml-chart.xsd`

```

```

### `CT_Boolean`

来源：`pro/schemas/dml-chart.xsd`

```

```

## 绘图与图形

wp:inline / wp:anchor 内部的顺序错一位，Word 会整张图丢掉。

### `CT_Inline`

来源：`pro/schemas/dml-wordprocessingDrawing.xsd`

```
    extent effectExtent? docPr cNvGraphicFramePr? graphic
```

### `CT_Anchor`

来源：`pro/schemas/dml-wordprocessingDrawing.xsd`

```
    simplePos positionH positionV extent effectExtent?
    (wrapNone | wrapSquare | wrapTight | wrapThrough | wrapTopAndBottom) docPr
    cNvGraphicFramePr? graphic
```

可出现的子元素（去重）：`simplePos` `positionH` `positionV` `extent` `effectExtent` `wrapNone` `wrapSquare` `wrapTight` `wrapThrough` `wrapTopAndBottom` `docPr` `cNvGraphicFramePr` `graphic`

### `CT_GraphicalObjectFrame`

来源：`pro/schemas/dml-spreadsheetDrawing.xsd`

```
    nvGraphicFramePr xfrm graphic
```

### `CT_GraphicalObject`

来源：`pro/schemas/dml-main.xsd`

```
    graphicData
```

### `CT_ShapeProperties`

来源：`pro/schemas/dml-main.xsd`

```
    xfrm? (custGeom | prstGeom) (noFill | solidFill | gradFill | blipFill | pattFill | grpFill)
    ln? (effectLst | effectDag) scene3d? sp3d? extLst?
```

可出现的子元素（去重）：`xfrm` `custGeom` `prstGeom` `noFill` `solidFill` `gradFill` `blipFill` `pattFill` `grpFill` `ln` `effectLst` `effectDag` `scene3d` `sp3d` `extLst`

### `CT_TextBody`

来源：`pro/schemas/dml-main.xsd`

```
    bodyPr lstStyle? p+
```

### `CT_TextParagraph`

来源：`pro/schemas/dml-main.xsd`

```
    pPr? (r | br | fld)? endParaRPr?
```

可出现的子元素（去重）：`pPr` `r` `br` `fld` `endParaRPr`

### `CT_TextBodyProperties`

来源：`pro/schemas/dml-main.xsd`

```
    prstTxWarp? (noAutofit | normAutofit | spAutoFit) scene3d? (sp3d | flatTx) extLst?
```

可出现的子元素（去重）：`prstTxWarp` `noAutofit` `normAutofit` `spAutoFit` `scene3d` `sp3d` `flatTx` `extLst`

### `CT_Transform2D`

来源：`pro/schemas/dml-main.xsd`

```
    off? ext?
```

### `CT_Picture`

来源：`pro/schemas/dml-chartDrawing.xsd`

```
    nvPicPr blipFill spPr style?
```

### `CT_BlipFillProperties`

来源：`pro/schemas/dml-main.xsd`

```
    blip? srcRect? (tile | stretch)
```

可出现的子元素（去重）：`blip` `srcRect` `tile` `stretch`

### `CT_Blip`

来源：`pro/schemas/dml-main.xsd`

```
    (alphaBiLevel | alphaCeiling | alphaFloor | alphaInv | alphaMod | alphaModFix | alphaRepl | biLevel | blur | clrChange | clrRepl | duotone | fillOverlay | grayscl | hsl | lum | tint)?
    extLst?
```

可出现的子元素（去重）：`alphaBiLevel` `alphaCeiling` `alphaFloor` `alphaInv` `alphaMod` `alphaModFix` `alphaRepl` `biLevel` `blur` `clrChange` `clrRepl` `duotone` `fillOverlay` `grayscl` `hsl` `lum` `tint` `extLst`

### `CT_GroupShape`

来源：`pro/schemas/dml-chartDrawing.xsd`

```
    nvGrpSpPr grpSpPr (sp | grpSp | graphicFrame | cxnSp | pic)?
```

可出现的子元素（去重）：`nvGrpSpPr` `grpSpPr` `sp` `grpSp` `graphicFrame` `cxnSp` `pic`

### `CT_Connector`

来源：`pro/schemas/dml-chartDrawing.xsd`

```
    nvCxnSpPr spPr style?
```

## WordprocessingML 扩展（2010/2012/2018）

扩展命名空间的元素同样受各自 XSD 约束；它们不能被丢进 mc:Ignorable。

### `CT_SdtContentBlock`

来源：`pro/schemas/wml.xsd`

```
    (customXml | sdt | p? | tbl? | (proofErr? | permStart? | permEnd? | (bookmarkStart | bookmarkEnd | moveFromRangeStart | moveFromRangeEnd | moveToRangeStart | moveToRangeEnd | commentRangeStart | commentRangeEnd | customXmlInsRangeStart | customXmlInsRangeEnd | customXmlDelRangeStart | customXmlDelRangeEnd | customXmlMoveFromRangeStart | customXmlMoveFromRangeEnd | customXmlMoveToRangeStart | customXmlMoveToRangeEnd)? | ins? | del? | moveFrom | moveTo | (oMathPara | oMath)?)?)?
```

可出现的子元素（去重）：`customXml` `sdt` `p` `tbl` `proofErr` `permStart` `permEnd` `bookmarkStart` `bookmarkEnd` `moveFromRangeStart` `moveFromRangeEnd` `moveToRangeStart` `moveToRangeEnd` `commentRangeStart` `commentRangeEnd` `customXmlInsRangeStart` `customXmlInsRangeEnd` `customXmlDelRangeStart` `customXmlDelRangeEnd` `customXmlMoveFromRangeStart` `customXmlMoveFromRangeEnd` `customXmlMoveToRangeStart` `customXmlMoveToRangeEnd` `ins` `del` `moveFrom` `moveTo` `oMathPara` `oMath`

### `CT_SdtContentCell`

来源：`pro/schemas/wml.xsd`

```
    (tc? | customXml | sdt | (proofErr? | permStart? | permEnd? | (bookmarkStart | bookmarkEnd | moveFromRangeStart | moveFromRangeEnd | moveToRangeStart | moveToRangeEnd | commentRangeStart | commentRangeEnd | customXmlInsRangeStart | customXmlInsRangeEnd | customXmlDelRangeStart | customXmlDelRangeEnd | customXmlMoveFromRangeStart | customXmlMoveFromRangeEnd | customXmlMoveToRangeStart | customXmlMoveToRangeEnd)? | ins? | del? | moveFrom | moveTo | (oMathPara | oMath)?)?)?
```

可出现的子元素（去重）：`tc` `customXml` `sdt` `proofErr` `permStart` `permEnd` `bookmarkStart` `bookmarkEnd` `moveFromRangeStart` `moveFromRangeEnd` `moveToRangeStart` `moveToRangeEnd` `commentRangeStart` `commentRangeEnd` `customXmlInsRangeStart` `customXmlInsRangeEnd` `customXmlDelRangeStart` `customXmlDelRangeEnd` `customXmlMoveFromRangeStart` `customXmlMoveFromRangeEnd` `customXmlMoveToRangeStart` `customXmlMoveToRangeEnd` `ins` `del` `moveFrom` `moveTo` `oMathPara` `oMath`
