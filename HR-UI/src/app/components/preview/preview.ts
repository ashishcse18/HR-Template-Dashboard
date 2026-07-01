import { Component, inject, Output, EventEmitter, Input } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DocumentService } from '../../services/document.service';
import { DownloadService } from '../../services/download.service';

@Component({
  selector: 'app-preview',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './preview.html',
  styleUrls: ['./preview.css'],
})
export class PreviewComponent {
  private readonly docService = inject(DocumentService);
  private readonly downloadService = inject(DownloadService);

  @Input() isOpen = false;
  @Input() htmlContent = '';

  @Output() close = new EventEmitter<void>();

  zoom = 75;

  get cleanHtml(): string {
    let clean = this.htmlContent;

    clean = clean.replace(/contenteditable="true"/g, 'contenteditable="false"');

    clean = clean.replace(/class="doc-var"/g, 'style="font-weight:600;border-bottom:none;"');

    return clean;
  }

  adjustZoom(amount: number): void {
    const newZoom = this.zoom + amount;

    if (newZoom >= 30 && newZoom <= 150) {
      this.zoom = newZoom;
    }
  }

  async downloadPDF() {
    const name = this.docService.activeTemplateName();
    await this.downloadService.downloadAsPDF('preview-print-container', name);
  }

  printDocument(): void {
    const printContent = document.getElementById('preview-print-container')?.innerHTML;

    const printWindow = window.open(
      'about:blank',
      new Date().getTime().toString(),
      'left=50000,top=50000,width=0,height=0',
    );

    if (printWindow) {
      printWindow.document.write(`
      <html>
      <head>
        <title>${this.docService.activeTemplateName()}</title>

        <style>
          body{
            font-family:Arial,sans-serif;
            padding:20px;
          }

          table{
            width:100%;
            border-collapse:collapse;
          }

          th,td{
            border:1px solid #e1e5eb;
            padding:8px;
          }
        </style>

      </head>

      <body>
        ${printContent}
      </body>

      </html>
      `);

      printWindow.document.close();
      printWindow.focus();
      printWindow.print();
      printWindow.close();
    }
  }
}