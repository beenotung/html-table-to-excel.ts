export function tableHTMLToExcelXML(tableHTML: string) {
  return /* html */ `
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<!--[if gte mso 9]>
<xml>
  <x:ExcelWorkbook>
    <x:ExcelWorksheets>
      <x:ExcelWorksheet>
        <x:Name>name</x:Name>
        <x:WorksheetOptions>
          <x:DisplayGridlines/>
        </x:WorksheetOptions>
      </x:ExcelWorksheet>
    </x:ExcelWorksheets>
  </x:ExcelWorkbook>
</xml>
<![endif]-->
</head>
<body>
${tableHTML}
</body>
</html>
`.trim()
}

export function excelXMLToDataUrl(xml: string): string {
  return `data:${mimeType},${xml}`
}

export function downloadDataUrlAsFile(dataUrl: string, filename: string) {
  let a = document.createElement('a')
  document.body.appendChild(a)
  a.href = dataUrl
  a.download = filename
  a.click()
  a.remove()
}

let mimeType = 'application/vnd.ms-excel'
let defaultFilename = 'data.xls'

export function downloadTableAsExcel(
  table: HTMLTableElement,
  filename: string = defaultFilename,
) {
  let xml = tableHTMLToExcelXML(table.outerHTML)
  let dataUrl = excelXMLToDataUrl(xml)
  downloadDataUrlAsFile(dataUrl, filename)
}

// TODO handle unicode
export function textToBlob(text: string, mimeType: string): Blob {
  let n = text.length
  let buffer = new Uint8Array(n)
  for (let i = 0; i < n; i++) {
    buffer[i] = text.charCodeAt(i)
  }
  return new Blob([buffer], { type: mimeType })
}

export function excelXMLToFile(
  xml: string,
  filename: string = defaultFilename,
  options?: Pick<FilePropertyBag, 'lastModified'>,
) {
  let blob = textToBlob(xml, mimeType)
  return new File([blob], filename, {
    type: blob.type,
    lastModified: options?.lastModified,
  })
}

export function tableToFile(
  table: HTMLTableElement,
  filename: string = defaultFilename,
  options?: Pick<FilePropertyBag, 'lastModified'>,
): File {
  let xml = tableHTMLToExcelXML(table.outerHTML)
  return excelXMLToFile(xml, filename, options)
}
