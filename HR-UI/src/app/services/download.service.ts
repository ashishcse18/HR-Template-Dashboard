import { Injectable } from '@angular/core';
import { saveAs } from 'file-saver';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle } from 'docx';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';

@Injectable({
  providedIn: 'root'
})
export class DownloadService {

  public async downloadAsPDF(elementId: string, filename: string): Promise<void> {
    const element = document.getElementById(elementId);
    if (!element) {
      console.error(`Element with id ${elementId} not found.`);
      return;
    }

    try {
      const canvas = await html2canvas(element, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff',
        logging: false
      });

      const imgData = canvas.toDataURL('image/png');
      const pdf = new jsPDF('p', 'mm', 'a4');
      const imgWidth = 210;
      const pageHeight = 297;
      const imgHeight = (canvas.height * imgWidth) / canvas.width;
      let heightLeft = imgHeight;
      let position = 0;


      pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
      heightLeft -= pageHeight;

      while (heightLeft > 1.0) {
        position = heightLeft - imgHeight;
        pdf.addPage();
        pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
        heightLeft -= pageHeight;
      }

      pdf.save(`${filename.replace(/\s+/g, '_')}.pdf`);
    } catch (error) {
      console.error('Error generating PDF:', error);
    }
  }


  public downloadAsWordHTML(htmlContent: string, filename: string): void {
    const header = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office'
            xmlns:w='urn:schemas-microsoft-com:office:office:word'
            xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset='utf-8'>
        <title>${filename}</title>
        <!--[if gte mso 9]>
        <xml>
          <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>100</w:Zoom>
            <w:DoNotOptimizeForBrowser/>
          </w:WordDocument>
        </xml>
        <![endif]-->
        <style>
          body {
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 11pt;
            color: #2e353e;
            line-height: 1.5;
            margin: 1in;
          }
          h2 { color: #5974a3; font-size: 18pt; font-weight: bold; margin-bottom: 5px; }
          h3 { color: #2e353e; font-size: 14pt; font-weight: bold; text-align: center; margin-bottom: 20px; }
          p { margin-bottom: 10px; }
          ol, ul { margin-left: 20px; margin-bottom: 15px; }
          li { margin-bottom: 5px; }
          table { width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; font-size: 10pt; }
          th { background-color: #5974a3; color: #ffffff; font-weight: bold; padding: 8px; border: 1px solid #e1e5eb; text-align: left; }
          td { padding: 8px; border: 1px solid #e1e5eb; }
          .net-pay-box { background-color: #eff3f8; border: 1px solid #5974a3; padding: 12px; margin-top: 15px; border-radius: 4px; }
        </style>
      </head>
      <body>
    `;
    const footer = '</body></html>';

    const fullContent = header + htmlContent + footer;

    const blob = new Blob(['\ufeff' + fullContent], {
      type: 'application/msword;charset=utf-8'
    });

    saveAs(blob, `${filename.replace(/\s+/g, '_')}.doc`);
  }

  public async downloadAsNativeDocx(variables: Record<string, string>, type: string, filename: string): Promise<void> {
    try {
      let doc: Document;

      if (type === 'payslip') {
        doc = this.generatePayslipDocx(variables);
      } else {
        doc = this.generateLetterDocx(variables, type);
      }

      const blob = await Packer.toBlob(doc);
      saveAs(blob, `${filename.replace(/\s+/g, '_')}.docx`);
    } catch (error) {
      console.error('Error generating native DOCX:', error);
    }
  }

  private generateLetterDocx(vars: Record<string, string>, type: string): Document {
    const title = type === 'offer' ? 'LETTER OF INTENT AND OFFER' :
                  type === 'relieving' ? 'RELIEVING ORDER & SERVICE CERTIFICATE' : 'EXPERIENCE CERTIFICATE';

    const paragraphs: Paragraph[] = [
      new Paragraph({
        children: [
          new TextRun({
            text: vars['companyName'] || 'Company Name',
            bold: true,
            size: 32,
            color: '5974a3',
          }),
        ],
        spacing: { after: 120 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: vars['companyAddress'] || 'Company Address',
            size: 18,
            color: '718096',
          }),
        ],
        spacing: { after: 400 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: title,
            bold: true,
            size: 26,
            color: '2e353e',
            underline: {},
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: `Date: ${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}`,
            bold: true,
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: 'To,',
            bold: true,
            size: 22,
          }),
        ],
        spacing: { after: 60 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: vars['employeeName'] || 'Employee Name',
            bold: true,
            size: 22,
            color: '5974a3',
          }),
        ],
        spacing: { after: 300 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: `Dear ${vars['employeeName'] || 'Employee'},`,
            size: 22,
          }),
        ],
        spacing: { after: 200 },
      }),
    ];

    if (type === 'offer') {
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `With reference to your interview and subsequent discussions we had with you, we are pleased to offer you employment with `,
            }),
            new TextRun({
              text: vars['companyName'] || 'our company',
              bold: true,
            }),
            new TextRun({
              text: ` as a `,
            }),
            new TextRun({
              text: vars['designation'] || 'Software Engineer',
              bold: true,
            }),
            new TextRun({
              text: `.`,
            }),
          ],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: `Your joining date is scheduled to be `,
            }),
            new TextRun({
              text: vars['joiningDate'] || 'Joining Date',
              bold: true,
            }),
            new TextRun({
              text: `. You will report to `,
            }),
            new TextRun({
              text: vars['reportingManager'] || 'Reporting Manager',
              bold: true,
            }),
            new TextRun({
              text: `. Your annual Gross Compensation (CTC) will be `,
            }),
            new TextRun({
              text: `INR ${vars['salary'] || 'Salary'}/- per annum`,
              bold: true,
            }),
            new TextRun({
              text: `.`,
            }),
          ],
          spacing: { after: 200 },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: `Terms & Conditions:`,
              bold: true,
              underline: {},
            }),
          ],
          spacing: { after: 120 },
        }),
        new Paragraph({
          text: `1. Probation: You will be on a probation period of ${vars['probationPeriod'] || '6 months'} from the date of joining.`,
          spacing: { after: 100 },
        }),
        new Paragraph({
          text: `2. Notice Period: Either party may terminate this employment by giving 60 days of written notice.`,
          spacing: { after: 200 },
        })
      );
    } else if (type === 'relieving') {
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `This is to certify and confirm that `,
            }),
            new TextRun({
              text: vars['employeeName'] || 'Employee Name',
              bold: true,
            }),
            new TextRun({
              text: ` (Employee ID: `,
            }),
            new TextRun({
              text: vars['empId'] || 'Emp ID',
              bold: true,
            }),
            new TextRun({
              text: `) was employed with `,
            }),
            new TextRun({
              text: vars['companyName'] || 'Company Name',
              bold: true,
            }),
            new TextRun({
              text: ` as a `,
            }),
            new TextRun({
              text: vars['designation'] || 'Designation',
              bold: true,
            }),
            new TextRun({
              text: ` from `,
            }),
            new TextRun({
              text: vars['joiningDate'] || 'Joining Date',
              bold: true,
            }),
            new TextRun({
              text: ` to `,
            }),
            new TextRun({
              text: vars['relievingDate'] || 'Relieving Date',
              bold: true,
            }),
            new TextRun({
              text: `.`,
            }),
          ],
          spacing: { after: 200 },
        }),
        new Paragraph({
          text: `He has resigned from the services of the Company of his own accord, and his resignation was accepted by the management. Consequently, he has been relieved of all responsibilities and duties in the company with effect from the close of business hours on ${vars['relievingDate'] || 'Relieving Date'}.`,
          spacing: { after: 200 },
        }),
        new Paragraph({
          text: `During his tenure with us, we found him to be dedicated, hard-working, and highly professional. We wish him all success in his future endeavors.`,
          spacing: { after: 400 },
        })
      );
    } else {
      // Service
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `This is to certify that `,
            }),
            new TextRun({
              text: vars['employeeName'] || 'Employee Name',
              bold: true,
            }),
            new TextRun({
              text: ` has worked in `,
            }),
            new TextRun({
              text: vars['companyName'] || 'Company Name',
              bold: true,
            }),
            new TextRun({
              text: ` in the position of `,
            }),
            new TextRun({
              text: vars['designation'] || 'Designation',
              bold: true,
            }),
            new TextRun({
              text: ` from `,
            }),
            new TextRun({
              text: vars['joiningDate'] || 'Joining Date',
              bold: true,
            }),
            new TextRun({
              text: ` to `,
            }),
            new TextRun({
              text: vars['relievingDate'] || 'Relieving Date',
              bold: true,
            }),
            new TextRun({
              text: `.`,
            }),
          ],
          spacing: { after: 200 },
        }),
        new Paragraph({
          text: `During this period, his duties included developing robust scalable web applications in Angular, implementing premium UI mockups, and collaborating with cross-functional product development teams. He has shown deep technical expertise and strong analytical capabilities.`,
          spacing: { after: 200 },
        }),
        new Paragraph({
          text: `We wish him the absolute best for all his future career milestones.`,
          spacing: { after: 400 },
        })
      );
    }
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `For ${vars['companyName'] || 'Company Name'}`,
            bold: true,
            color: '5974a3',
          }),
        ],
        spacing: { before: 400, after: 800 },
      }),
      new Paragraph({
        children: [
          new TextRun({
            text: `Human Resources Department`,
            bold: true,
          }),
        ],
      })
    );

    return new Document({
      sections: [{
        properties: {},
        children: paragraphs,
      }],
    });
  }

  private generatePayslipDocx(vars: Record<string, string>): Document {
    const basicVal = parseFloat(vars['basic'] || '0') || 0;
    const hraVal = parseFloat(vars['hra'] || '0') || 0;
    const convVal = parseFloat(vars['conveyance'] || '0') || 0;
    const specVal = parseFloat(vars['special'] || '0') || 0;
    const pfVal = parseFloat(vars['pf'] || '0') || 0;
    const ptVal = parseFloat(vars['pt'] || '0') || 0;
    const taxVal = parseFloat(vars['tax'] || '0') || 0;

    const totalEarnings = basicVal + hraVal + convVal + specVal;
    const totalDeductions = pfVal + ptVal + taxVal;
    const netSalary = totalEarnings - totalDeductions;

    const table = new Table({
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: "EARNINGS", bold: true, color: "ffffff" })] })],
              shading: { fill: "5974a3" },
            }),
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: "AMOUNT (INR)", bold: true, color: "ffffff" })], alignment: AlignmentType.RIGHT })],
              shading: { fill: "5974a3" },
            }),
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: "DEDUCTIONS", bold: true, color: "ffffff" })] })],
              shading: { fill: "5974a3" },
            }),
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: "AMOUNT (INR)", bold: true, color: "ffffff" })], alignment: AlignmentType.RIGHT })],
              shading: { fill: "5974a3" },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: "Basic Pay" })] }),
            new TableCell({ children: [new Paragraph({ text: basicVal.toLocaleString(), alignment: AlignmentType.RIGHT })] }),
            new TableCell({ children: [new Paragraph({ text: "Provident Fund (PF)" })] }),
            new TableCell({ children: [new Paragraph({ text: pfVal.toLocaleString(), alignment: AlignmentType.RIGHT })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: "House Rent Allowance" })] }),
            new TableCell({ children: [new Paragraph({ text: hraVal.toLocaleString(), alignment: AlignmentType.RIGHT })] }),
            new TableCell({ children: [new Paragraph({ text: "Professional Tax (PT)" })] }),
            new TableCell({ children: [new Paragraph({ text: ptVal.toLocaleString(), alignment: AlignmentType.RIGHT })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: "Conveyance Allowance" })] }),
            new TableCell({ children: [new Paragraph({ text: convVal.toLocaleString(), alignment: AlignmentType.RIGHT })] }),
            new TableCell({ children: [new Paragraph({ text: "TDS / Income Tax" })] }),
            new TableCell({ children: [new Paragraph({ text: taxVal.toLocaleString(), alignment: AlignmentType.RIGHT })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: "Special Allowance" })] }),
            new TableCell({ children: [new Paragraph({ text: specVal.toLocaleString(), alignment: AlignmentType.RIGHT })] }),
            new TableCell({ children: [new Paragraph({ text: "-" })] }),
            new TableCell({ children: [new Paragraph({ text: "0", alignment: AlignmentType.RIGHT })] }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Total Earnings (A)", bold: true })] })], shading: { fill: "f7fafc" } }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: totalEarnings.toLocaleString(), bold: true, color: "38a169" })], alignment: AlignmentType.RIGHT })], shading: { fill: "f7fafc" } }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Total Deductions (B)", bold: true })] })], shading: { fill: "f7fafc" } }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: totalDeductions.toLocaleString(), bold: true, color: "e53e3e" })], alignment: AlignmentType.RIGHT })], shading: { fill: "f7fafc" } }),
          ],
        }),
      ],
    });

    return new Document({
      sections: [{
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: vars['companyName'] || 'Company Name',
                bold: true,
                size: 28,
                color: '5974a3',
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 60 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: vars['companyAddress'] || 'Company Address',
                size: 16,
                color: '718096',
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `PAYSLIP FOR THE MONTH OF ${vars['monthYear'] || 'MONTH'}`,
                bold: true,
                size: 20,
                color: '57a6b0',
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 300 },
          }),
          // Metadata Paragraph
          new Paragraph({
            children: [
              new TextRun({ text: `Employee Name: `, bold: true }),
              new TextRun({ text: vars['employeeName'] || 'Name' }),
              new TextRun({ text: `\tEmployee ID: `, bold: true }),
              new TextRun({ text: vars['empId'] || 'ID' }),
            ],
            spacing: { after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: `Designation: `, bold: true }),
              new TextRun({ text: vars['designation'] || 'Designation' }),
              new TextRun({ text: `\tBank Account: `, bold: true }),
              new TextRun({ text: `${vars['bankAcc'] || ''} (${vars['bankName'] || ''})` }),
            ],
            spacing: { after: 300 },
          }),
          table,
          new Paragraph({
            children: [
              new TextRun({ text: `\nNet Take-Home Salary: `, bold: true }),
              new TextRun({ text: `INR ${netSalary.toLocaleString()}/-`, bold: true, color: "5974a3" }),
            ],
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: `Amount in Words: `, italics: true }),
              new TextRun({ text: this.numberToWords(netSalary) + " Only", italics: true, bold: true }),
            ],
            spacing: { after: 400 },
          }),
          new Paragraph({
            children: [new TextRun({ text: "This is a computer-generated payslip and does not require a physical signature.", size: 16, color: "a0aec0" })],
            alignment: AlignmentType.CENTER,
          }),
        ],
      }],
    });
  }

  private numberToWords(num: number): string {
    const a = ['', 'One ', 'Two ', 'Three ', 'Four ', 'Five ', 'Six ', 'Seven ', 'Eight ', 'Nine ', 'Ten ', 'Eleven ', 'Twelve ', 'Thirteen ', 'Fourteen ', 'Fifteen ', 'Sixteen ', 'Seventeen ', 'Eighteen ', 'Nineteen '];
    const b = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

    if ((num = Math.floor(num)) === 0) return 'Zero';

    const n = ('000000000' + num).substr(-9).match(/^(\d{2})(\d{2})(\d{2})(\d{1})(\d{2})$/);
    if (!n) return '';

    let str = '';
    str += Number(n[1]) != 0 ? (a[Number(n[1])] || b[Number(n[1].substr(0, 1))] + ' ' + a[Number(n[1].substr(1))]) + 'Crore ' : '';
    str += Number(n[2]) != 0 ? (a[Number(n[2])] || b[Number(n[2].substr(0, 1))] + ' ' + a[Number(n[2].substr(1))]) + 'Lakh ' : '';
    str += Number(n[3]) != 0 ? (a[Number(n[3])] || b[Number(n[3].substr(0, 1))] + ' ' + a[Number(n[3].substr(1))]) + 'Thousand ' : '';
    str += Number(n[4]) != 0 ? (a[Number(n[4])] || b[Number(n[4].substr(0, 1))] + ' ' + a[Number(n[4].substr(1))]) + 'Hundred ' : '';
    str += Number(n[5]) != 0 ? ((str != '') ? 'and ' : '') + (a[Number(n[5])] || b[Number(n[5].substr(0, 1))] + ' ' + a[Number(n[5].substr(1))]) : '';
    return str.trim();
  }
}