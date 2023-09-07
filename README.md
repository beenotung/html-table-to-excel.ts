# html-table-to-excel.ts

Convert a html table element into excel file

[![npm Package Version](https://img.shields.io/npm/v/html-table-to-excel.ts)](https://www.npmjs.com/package/html-table-to-excel.ts)
[![Minified Package Size](https://img.shields.io/bundlephobia/min/html-table-to-excel.ts)](https://bundlephobia.com/package/html-table-to-excel.ts)
[![Minified and Gzipped Package Size](https://img.shields.io/bundlephobia/minzip/html-table-to-excel.ts)](https://bundlephobia.com/package/html-table-to-excel.ts)
[![npm Package Downloads](https://img.shields.io/npm/dm/html-table-to-excel.ts)](https://www.npmtrends.com/html-table-to-excel.ts)

## Installation

```bash
npm install html-table-to-excel.ts
```

You can also install with pnpm, yarn, or slnpm

Then import from typescript:

```typescript
import { downloadTableAsExcel, tableToFile } from 'html-table-to-excel.ts'
```

Or require from javascript:

```javascript
let { downloadTableAsExcel, tableToFile } = require('html-table-to-excel.ts')
```

You can also load the browser.js or esm.js from CDN

Load as traditional javascript library

```html
<script src="https://cdn.jsdelivr.net/npm/html-table-to-excel.ts@1/browser.js"></script>
```

Or load as ESM module

```javascript
import {
  downloadTableAsExcel,
  tableToFile,
} from 'https://cdn.jsdelivr.net/npm/html-table-to-excel.ts@1/esm.js'
```

## Usage Example

```typescript
import { downloadTableAsExcel, tableToFile } from 'html-table-to-excel.ts'

let table = document.querySelector('table')!

document.querySelector('#download')!.addEventListener('click', () => {
  downloadTableAsExcel(table, 'data.xls')
})

document.querySelector('#upload')!.addEventListener('click', () => {
  let file = tableToFile(table, 'data.xls')
  let formData = new FormData()
  formData.append('file', file)
  fetch('/upload', { method: 'POST', body: formData })
})
```

## Typescript Types

```typescript
/* core functions */

// download from browser
export function downloadTableAsExcel(
  table: HTMLTableElement,
  filename?: string,
): void

// to be uploaded to server
export function tableToFile(
  table: HTMLTableElement,
  filename?: string,
  options?: Pick<FilePropertyBag, 'lastModified'>,
): File

/* helper functions */

export function tableHTMLToExcelXML(tableHTML: string): string

export function excelXMLToDataUrl(xml: string): string

export function downloadDataUrlAsFile(dataUrl: string, filename: string): void

export function textToBlob(text: string, mimeType: string): Blob

export function excelXMLToFile(
  xml: string,
  filename?: string,
  options?: Pick<FilePropertyBag, 'lastModified'>,
): File
```

## License

This project is licensed with [BSD-2-Clause](./LICENSE)

This is free, libre, and open-source software. It comes down to four essential freedoms [[ref]](https://seirdy.one/2021/01/27/whatsapp-and-the-domestication-of-users.html#fnref:2):

- The freedom to run the program as you wish, for any purpose
- The freedom to study how the program works, and change it so it does your computing as you wish
- The freedom to redistribute copies so you can help others
- The freedom to distribute copies of your modified versions to others
