# Third-party notices

`@filen/office` is licensed under the GNU Affero General Public License v3.0. It contains code derived from the projects below; each derived file carries a header naming its source.

## SuperDoc 1.x

- Source: SuperDoc v1.46.3 (`e096323e7d5e54f95b4817d1d3aefde8c37d52cf`), https://github.com/superdoc/docx-editor
- Copyright: Harbour and the SuperDoc contributors
- Licence: GNU Affero General Public License v3.0
- Derived files, modified for `@filen/office` (AGPL-3.0 §5a), first ported and last changed on 2026-10-11:
  - `src/core/docx/numbering/formats.ts` (from `shared/common/list-numbering`)
  - `src/core/docx/numbering/counter.ts` (from `packages/word-layout/src/numbering-manager.ts`)
  - `src/core/docx/styles/merge.ts` (from `packages/layout-engine/style-engine/src/cascade.ts`)
  - `src/core/docx/styles/resolve.ts` and `src/core/docx/styles/tableStyle.ts` (from `packages/layout-engine/style-engine/src/ooxml/index.ts`)
  - `src/core/math/omml.ts` (from `packages/layout-engine/painters/dom/src/features/math`)
  - `src/core/math/tokens.ts` (from `packages/layout-engine/painters/dom/src/features/math/converters/math-run.ts`)

Only the AGPL-3.0 1.x line is used. No code from `superdoc` 2.3 or later, or from `@superdoc/docx-engine`, is included.

## docx-preview (docxjs)

- Source: docx-preview 0.4.1 (docxjs `c533f383`), https://github.com/VolodymyrBaydalka/docxjs
- Copyright: Volodymyr Baydalka
- Licence: Apache License 2.0 (text below)
- Derived files, modified for `@filen/office`, first ported and last changed on 2026-10-11:
  - `src/core/math/omml.ts` (from `src/html-renderer.ts`, the `renderMml*` functions)

## ECMA-376 preset shape definitions

- Source: ECMA-376 Part 1, 5th edition (December 2016), annex `OfficeOpenXML-DrawingMLGeometries.zip`, `presetShapeDefinitions.xml` (SHA-256 `2f7c868d857c1e3c4b5a6068759fe0e07d77ad58377a6618d1b02ba3507b6939`), https://ecma-international.org/publications-and-standards/standards/ecma-376/
- Copyright: Ecma International
- `src/core/drawingml/geometry/presets.generated.ts` is generated from it by `src/scripts/presetGeometry.ts` (re-encoded, not copied), implementing the specification's functionality in a conformant product (permission (iv) below).
- Also: ECMA-376 Part 4, 5th edition (December 2016), `OfficeOpenXML-XMLSchema-Transitional.zip`, `wml.xsd` (SHA-256 `0f7784b36ddc42950e5ff354cd608b83b60e4de39325d5568913f1225365029b`) and `sml.xsd` (SHA-256 `495debc8fa967b77ed37799747b049832f2c95b2ecdb9f19ddcd68f8c9f96ab9`). `src/core/docx/write/order.generated.ts` and `src/core/xlsx/write/order.generated.ts` list child element orders and enumerated values read from them by `src/scripts/wmlOrder.ts` and `src/scripts/smlOrder.ts`, under the same permission. The schema validation tests (`src/tests/docxWriteSchema.test.ts`, `src/tests/oracles/sheet/sheetSchema.test.ts`) read the whole schema set from the pinned archive at test time; none of it is committed.
- Licence: the Ecma default copyright notice ("Ecma text copyright policy", version 3, https://ecma-international.org/wp-content/uploads/Ecma_text_copyright_policy_version_3.pdf, SHA-256 `2dcbec958059c8b512b6798ee1f0b8416c33d507a21f5bd9a20c19e68218776f`), reproduced verbatim (line breaks of the PDF joined into paragraphs):

```
COPYRIGHT NOTICE
© Ecma International
By obtaining and/or copying this work, you (the licensee) agree that you have read, understood, and will comply with the following terms and conditions.
This document may be copied, published and distributed to others, and certain derivative works of it may be prepared, copied, published, and distributed, in whole or in part, provided that the above copyright notice and this Copyright License and Disclaimer are included on all such copies and derivative works. The only derivative works that are permissible under this Copyright License and Disclaimer are:
(i) works which incorporate all or portion of this document for the purpose of providing commentary or explanation (such as an annotated version of the document),
(ii) works which incorporate all or portion of this document for the purpose of incorporating features that provide accessibility,
(iii) translations of this document into languages other than English and into different formats and
(iv) works by making use of this specification in standard conformant products by implementing (e.g. by copy and paste wholly or partly) the functionality therein.
However, the content of this document itself may not be modified in any way, including by removing the copyright notice or references to Ecma International, except as required to translate it into languages other than English or into a different format.
The official version of an Ecma International document is the English language version on the Ecma International website. In the event of discrepancies between a translated version and the official version, the official version shall govern.
The limited permissions granted above are perpetual and will not be revoked by Ecma International or its successors or assigns.
This document and the information contained herein is provided on an “AS IS” basis and ECMA INTERNATIONAL DISCLAIMS ALL WARRANTIES, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTY THAT THE USE OF THE INFORMATION HEREIN WILL NOT INFRINGE ANY OWNERSHIP RIGHTS OR ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR FITNESS FOR A PARTICULAR PURPOSE.
```

## Unicode symbol-font mappings

`src/core/fonts/symbols.generated.ts` is generated by `src/scripts/symbolTables.ts` from two Unicode, Inc. publications (re-encoded as code tables, not copied):

- Wingdings, Wingdings 2, Wingdings 3 and Webdings: ISO/IEC JTC1/SC2/WG2 N4384 (Unicode L2/12-368, "Status of encoding of Wingdings and Webdings Symbols", 2012-11-06), https://www.unicode.org/L2/L2012/12368-n4384.pdf (SHA-256 `15839e133b3ab10b84d68f41f25138d9dedbc94a8a6c2dbb155439d597084a0b`), used under the Unicode Terms of Use, https://www.unicode.org/copyright.html
- Symbol: "Adobe Symbol Encoding to Unicode", table version 1.0 (2011 July 12), https://www.unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt (SHA-256 `deb78ca840a429311939b9d165890873f71fb23ef223ceeb144a6c6d641a7e52`), whose notice is reproduced verbatim as it requires:

```
Copyright (c) 1991-2011 Unicode, Inc. All Rights reserved.

This file is provided as-is by Unicode, Inc. (The Unicode Consortium). No
claims are made as to fitness for any particular purpose. No warranties of
any kind are expressed or implied. The recipient agrees to determine
applicability of information provided. If this file has been provided on
magnetic media by Unicode, Inc., the sole remedy for any claim will be
exchange of defective media within 90 days of receipt.

Unicode, Inc. hereby grants the right to freely use the information
supplied in this file in the creation of products supporting the
Unicode Standard, and to make copies of this file in any form for
internal or external distribution as long as this notice remains
attached.
```

## Metafile symbol-font supplements

`src/core/metafile/text.ts` and `src/core/metafile/mtextra.ts` map further symbol-font tables for metafile text, re-encoded as code tables (no file is copied):

- The Symbol font's bracket, brace and integral pieces and arrow extensions: 22 factual code-to-character mappings (0xBD → U+23D0, 0xBE → U+23AF, 0xE6–0xEF → U+239B–U+239D, U+23A1–U+23A3, U+23A7–U+23AA, 0xF4 → U+23AE, 0xF6–0xFE → U+239E–U+23A0, U+23A4–U+23A6, U+23AB–U+23AD) taken from Apple's "Map (external version) from Mac OS Symbol character set to Unicode 4.0 and later", version c02 (2005-04-05), https://www.unicode.org/Public/MAPPINGS/VENDORS/APPLE/SYMBOL.TXT (SHA-256 `b8c529daba09a45872ec3b493e944680528206b829193594dcdd774a02d49d12`). Its copyright line and disclaimer, reproduced verbatim:

```
Copyright:  (c) 1994-2002, 2005 by Apple Computer, Inc., all rights
            reserved.

Apple Computer, Inc. ("Apple") makes no warranty or representation,
either express or implied, with respect to this document and the
included data, its quality, accuracy, or fitness for a particular
purpose. In no event will Apple be liable for direct, indirect,
special, incidental, or consequential damages resulting from any
defect or inaccuracy in this document or the included data.
```

- The advance widths of those pieces in Symbol (five facts: 384, 494, 686, 603 and 1000 units per em), from Adobe's Symbol font metrics, `Symbol.afm` version 001.008 of the Core 14 AFM set ("Copyright (c) 1985, 1987, 1989, 1990, 1997 Adobe Systems Incorporated. All rights reserved."), used to centre a piece in its Symbol cell.
- MT Extra (MathType's companion font), codes 0x20–0x7F, in `src/core/metafile/mtextra.ts`:
  - Source: the Mozilla MathML project's "MT Extra" encoding page, https://www-archive.mozilla.org/projects/mathml/fonts/encoding/mtextra (SHA-256 `b36a96233316626a49c26e9e30d42ae422283798c3b87664510474c7da03e06a`). No MT Extra font file was read.
  - Attribution: "Portions of this content are © 1998–2009 by individual mozilla.org contributors".
  - Licence to us: Creative Commons Attribution-ShareAlike 4.0 International (https://creativecommons.org/licenses/by-sa/4.0/). The page's licence (https://www.mozilla.org/foundation/licensing/website-content/) is CC BY-SA 3.0 Unported "or any later version".
  - Modified: the combining embellishments are mapped to their spacing forms, the arrow embellishments to spacing arrows, and the stretchy brace pieces are left out.
  - Our adaptation is licensed under the GNU General Public License version 3 (GPL-3.0-only), under Creative Commons' declared one-way compatibility of BY-SA 4.0 with GPLv3 (https://wiki.creativecommons.org/wiki/ShareAlike_compatibility:_GPLv3), and is combined with the rest of `@filen/office` under section 13 of the GNU Affero General Public License v3.0.

## pdf.js

- Source: pdfjs-dist 6.3.289, https://github.com/mozilla/pdf.js
- Copyright: 2014 Mozilla Foundation
- Licence: Apache License 2.0 (text below)
- Derived files, modified for `@filen/office`:
  - `src/css/pdf.css` (the `.textLayer` rules of `web/pdf_viewer.css`, selector renamed and colours made variables)

## CodeMirror highlight styles

- Source: @codemirror/language 6.12.4 (`defaultHighlightStyle`) and @codemirror/theme-one-dark 6.1.3 (`oneDarkHighlightStyle`), https://github.com/codemirror
- Copyright: 2018-2021 Marijn Haverbeke and others
- Licence: MIT (text below)
- Derived files, modified for `@filen/office`:
  - `src/dom/code/theme.ts` (the two styles' colours, weights and decorations per highlight tag, as CSS variables)

```
MIT License

Copyright (C) 2018-2021 by Marijn Haverbeke <marijn@haverbeke.berlin> and others

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
```

## Apache License 2.0

```
                                 Apache License
                           Version 2.0, January 2004
                        http://www.apache.org/licenses/

   TERMS AND CONDITIONS FOR USE, REPRODUCTION, AND DISTRIBUTION

   1. Definitions.

      "License" shall mean the terms and conditions for use, reproduction,
      and distribution as defined by Sections 1 through 9 of this document.

      "Licensor" shall mean the copyright owner or entity authorized by
      the copyright owner that is granting the License.

      "Legal Entity" shall mean the union of the acting entity and all
      other entities that control, are controlled by, or are under common
      control with that entity. For the purposes of this definition,
      "control" means (i) the power, direct or indirect, to cause the
      direction or management of such entity, whether by contract or
      otherwise, or (ii) ownership of fifty percent (50%) or more of the
      outstanding shares, or (iii) beneficial ownership of such entity.

      "You" (or "Your") shall mean an individual or Legal Entity
      exercising permissions granted by this License.

      "Source" form shall mean the preferred form for making modifications,
      including but not limited to software source code, documentation
      source, and configuration files.

      "Object" form shall mean any form resulting from mechanical
      transformation or translation of a Source form, including but
      not limited to compiled object code, generated documentation,
      and conversions to other media types.

      "Work" shall mean the work of authorship, whether in Source or
      Object form, made available under the License, as indicated by a
      copyright notice that is included in or attached to the work
      (an example is provided in the Appendix below).

      "Derivative Works" shall mean any work, whether in Source or Object
      form, that is based on (or derived from) the Work and for which the
      editorial revisions, annotations, elaborations, or other modifications
      represent, as a whole, an original work of authorship. For the purposes
      of this License, Derivative Works shall not include works that remain
      separable from, or merely link (or bind by name) to the interfaces of,
      the Work and Derivative Works thereof.

      "Contribution" shall mean any work of authorship, including
      the original version of the Work and any modifications or additions
      to that Work or Derivative Works thereof, that is intentionally
      submitted to Licensor for inclusion in the Work by the copyright owner
      or by an individual or Legal Entity authorized to submit on behalf of
      the copyright owner. For the purposes of this definition, "submitted"
      means any form of electronic, verbal, or written communication sent
      to the Licensor or its representatives, including but not limited to
      communication on electronic mailing lists, source code control systems,
      and issue tracking systems that are managed by, or on behalf of, the
      Licensor for the purpose of discussing and improving the Work, but
      excluding communication that is conspicuously marked or otherwise
      designated in writing by the copyright owner as "Not a Contribution."

      "Contributor" shall mean Licensor and any individual or Legal Entity
      on behalf of whom a Contribution has been received by Licensor and
      subsequently incorporated within the Work.

   2. Grant of Copyright License. Subject to the terms and conditions of
      this License, each Contributor hereby grants to You a perpetual,
      worldwide, non-exclusive, no-charge, royalty-free, irrevocable
      copyright license to reproduce, prepare Derivative Works of,
      publicly display, publicly perform, sublicense, and distribute the
      Work and such Derivative Works in Source or Object form.

   3. Grant of Patent License. Subject to the terms and conditions of
      this License, each Contributor hereby grants to You a perpetual,
      worldwide, non-exclusive, no-charge, royalty-free, irrevocable
      (except as stated in this section) patent license to make, have made,
      use, offer to sell, sell, import, and otherwise transfer the Work,
      where such license applies only to those patent claims licensable
      by such Contributor that are necessarily infringed by their
      Contribution(s) alone or by combination of their Contribution(s)
      with the Work to which such Contribution(s) was submitted. If You
      institute patent litigation against any entity (including a
      cross-claim or counterclaim in a lawsuit) alleging that the Work
      or a Contribution incorporated within the Work constitutes direct
      or contributory patent infringement, then any patent licenses
      granted to You under this License for that Work shall terminate
      as of the date such litigation is filed.

   4. Redistribution. You may reproduce and distribute copies of the
      Work or Derivative Works thereof in any medium, with or without
      modifications, and in Source or Object form, provided that You
      meet the following conditions:

      (a) You must give any other recipients of the Work or
          Derivative Works a copy of this License; and

      (b) You must cause any modified files to carry prominent notices
          stating that You changed the files; and

      (c) You must retain, in the Source form of any Derivative Works
          that You distribute, all copyright, patent, trademark, and
          attribution notices from the Source form of the Work,
          excluding those notices that do not pertain to any part of
          the Derivative Works; and

      (d) If the Work includes a "NOTICE" text file as part of its
          distribution, then any Derivative Works that You distribute must
          include a readable copy of the attribution notices contained
          within such NOTICE file, excluding those notices that do not
          pertain to any part of the Derivative Works, in at least one
          of the following places: within a NOTICE text file distributed
          as part of the Derivative Works; within the Source form or
          documentation, if provided along with the Derivative Works; or,
          within a display generated by the Derivative Works, if and
          wherever such third-party notices normally appear. The contents
          of the NOTICE file are for informational purposes only and
          do not modify the License. You may add Your own attribution
          notices within Derivative Works that You distribute, alongside
          or as an addendum to the NOTICE text from the Work, provided
          that such additional attribution notices cannot be construed
          as modifying the License.

      You may add Your own copyright statement to Your modifications and
      may provide additional or different license terms and conditions
      for use, reproduction, or distribution of Your modifications, or
      for any such Derivative Works as a whole, provided Your use,
      reproduction, and distribution of the Work otherwise complies with
      the conditions stated in this License.

   5. Submission of Contributions. Unless You explicitly state otherwise,
      any Contribution intentionally submitted for inclusion in the Work
      by You to the Licensor shall be under the terms and conditions of
      this License, without any additional terms or conditions.
      Notwithstanding the above, nothing herein shall supersede or modify
      the terms of any separate license agreement you may have executed
      with Licensor regarding such Contributions.

   6. Trademarks. This License does not grant permission to use the trade
      names, trademarks, service marks, or product names of the Licensor,
      except as required for reasonable and customary use in describing the
      origin of the Work and reproducing the content of the NOTICE file.

   7. Disclaimer of Warranty. Unless required by applicable law or
      agreed to in writing, Licensor provides the Work (and each
      Contributor provides its Contributions) on an "AS IS" BASIS,
      WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or
      implied, including, without limitation, any warranties or conditions
      of TITLE, NON-INFRINGEMENT, MERCHANTABILITY, or FITNESS FOR A
      PARTICULAR PURPOSE. You are solely responsible for determining the
      appropriateness of using or redistributing the Work and assume any
      risks associated with Your exercise of permissions under this License.

   8. Limitation of Liability. In no event and under no legal theory,
      whether in tort (including negligence), contract, or otherwise,
      unless required by applicable law (such as deliberate and grossly
      negligent acts) or agreed to in writing, shall any Contributor be
      liable to You for damages, including any direct, indirect, special,
      incidental, or consequential damages of any character arising as a
      result of this License or out of the use or inability to use the
      Work (including but not limited to damages for loss of goodwill,
      work stoppage, computer failure or malfunction, or any and all
      other commercial damages or losses), even if such Contributor
      has been advised of the possibility of such damages.

   9. Accepting Warranty or Additional Liability. While redistributing
      the Work or Derivative Works thereof, You may choose to offer,
      and charge a fee for, acceptance of support, warranty, indemnity,
      or other liability obligations and/or rights consistent with this
      License. However, in accepting such obligations, You may act only
      on Your own behalf and on Your sole responsibility, not on behalf
      of any other Contributor, and only if You agree to indemnify,
      defend, and hold each Contributor harmless for any liability
      incurred by, or claims asserted against, such Contributor by reason
      of your accepting any such warranty or additional liability.

   END OF TERMS AND CONDITIONS

   APPENDIX: How to apply the Apache License to your work.

      To apply the Apache License to your work, attach the following
      boilerplate notice, with the fields enclosed by brackets "{}"
      replaced with your own identifying information. (Don't include
      the brackets!)  The text should be enclosed in the appropriate
      comment syntax for the file format. We also recommend that a
      file or class name and description of purpose be included on the
      same "printed page" as the copyright notice for easier
      identification within third-party archives.

   Copyright (c) 2016-2023 Volodymyr Baydalka

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
```
