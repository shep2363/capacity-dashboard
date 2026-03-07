import html2canvas from 'html2canvas'
import { jsPDF } from 'jspdf'

interface ExportPdfOptions {
  fileName: string
}

export async function exportReportElementToPdf(element: HTMLElement, options: ExportPdfOptions): Promise<void> {
  element.classList.add('pdf-exporting')

  try {
    const canvas = await html2canvas(element, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff',
      logging: false,
    })

    const pdf = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' })
    const pageWidth = pdf.internal.pageSize.getWidth()
    const pageHeight = pdf.internal.pageSize.getHeight()
    const margin = 24
    const printableWidth = pageWidth - margin * 2
    const printableHeight = pageHeight - margin * 2

    const pageSliceHeightPx = Math.floor((printableHeight * canvas.width) / printableWidth)
    let renderedHeightPx = 0
    let pageIndex = 0

    while (renderedHeightPx < canvas.height) {
      const sliceHeightPx = Math.min(pageSliceHeightPx, canvas.height - renderedHeightPx)
      const pageCanvas = document.createElement('canvas')
      pageCanvas.width = canvas.width
      pageCanvas.height = sliceHeightPx

      const pageCtx = pageCanvas.getContext('2d')
      if (!pageCtx) {
        throw new Error('Unable to prepare PDF canvas context.')
      }

      pageCtx.drawImage(
        canvas,
        0,
        renderedHeightPx,
        canvas.width,
        sliceHeightPx,
        0,
        0,
        canvas.width,
        sliceHeightPx,
      )

      const imageData = pageCanvas.toDataURL('image/png')
      const renderHeightPt = (sliceHeightPx * printableWidth) / canvas.width
      if (pageIndex > 0) {
        pdf.addPage()
      }
      pdf.addImage(imageData, 'PNG', margin, margin, printableWidth, renderHeightPt, undefined, 'FAST')

      renderedHeightPx += sliceHeightPx
      pageIndex += 1
    }

    pdf.save(options.fileName)
  } finally {
    element.classList.remove('pdf-exporting')
  }
}
