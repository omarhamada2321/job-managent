import { Document, Packer, Paragraph, HeadingLevel, AlignmentType, Table, TableRow, TableCell, BorderStyle, WidthType, TextRun } from 'docx';
import { jsPDF } from 'jspdf';

export interface Task {
  id: string;
  title: string;
  completed: boolean;
  date: string;
}

export interface TasksByDate {
  [date: string]: Task[];
}

export interface ReportData {
  startDate: string;
  endDate: string;
  tasks: TasksByDate;
  totalTasks: number;
  totalCompleted: number;
  completionRate: number;
}

const formatDate = (dateStr: string) => {
  const date = new Date(dateStr + 'T00:00:00');
  return date.toLocaleDateString('en-US', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric',
  });
};

const formatShortDate = (dateStr: string) => {
  const date = new Date(dateStr + 'T00:00:00');
  return date.toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
  });
};

export const generateTextReport = (data: ReportData): string => {
  let report = `WEEKLY TASK REPORT\n`;
  report += `${'='.repeat(60)}\n\n`;

  report += `Report Period: ${formatDate(data.startDate)} - ${formatDate(data.endDate)}\n`;
  report += `Generated: ${new Date().toLocaleString()}\n\n`;

  report += `${'='.repeat(60)}\n`;
  report += `SUMMARY\n`;
  report += `${'='.repeat(60)}\n`;
  report += `Total Tasks: ${data.totalTasks}\n`;
  report += `Completed: ${data.totalCompleted}\n`;
  report += `Pending: ${data.totalTasks - data.totalCompleted}\n`;
  report += `Completion Rate: ${data.completionRate}%\n\n`;

  const dates = Object.keys(data.tasks).sort();
  report += `${'='.repeat(60)}\n`;
  report += `DETAILED BREAKDOWN\n`;
  report += `${'='.repeat(60)}\n\n`;

  dates.forEach((date) => {
    const dayTasks = data.tasks[date];
    const completed = dayTasks.filter((t) => t.completed).length;

    report += `${formatDate(date)}\n`;
    report += `${'-'.repeat(60)}\n`;

    dayTasks.forEach((task) => {
      const status = task.completed ? '[COMPLETED]' : '[PENDING]';
      report += `${status} ${task.title}\n`;
    });

    report += `\nDay Summary: ${completed}/${dayTasks.length} tasks completed\n\n`;
  });

  report += `${'='.repeat(60)}\n`;
  report += `END OF REPORT\n`;
  report += `${'='.repeat(60)}\n`;

  return report;
};

export const generateWordReport = async (data: ReportData, fileName: string) => {
  const dates = Object.keys(data.tasks).sort();

  const sections: Paragraph[] = [
    new Paragraph({
      text: 'WEEKLY TASK REPORT',
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: `Report Period: ${formatDate(data.startDate)} - ${formatDate(data.endDate)}`,
      spacing: { after: 100 },
    }),
    new Paragraph({
      text: `Generated: ${new Date().toLocaleString()}`,
      spacing: { after: 400 },
    }),
    new Paragraph({
      text: 'Summary',
      heading: HeadingLevel.HEADING_2,
      spacing: { after: 200 },
    }),
  ];

  const summaryTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        cells: [
          new TableCell({
            children: [new Paragraph({ text: 'Total Tasks', bold: true })],
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            },
          }),
          new TableCell({
            children: [new Paragraph(data.totalTasks.toString())],
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            },
          }),
        ],
      }),
      new TableRow({
        cells: [
          new TableCell({
            children: [new Paragraph({ text: 'Completed', bold: true })],
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            },
          }),
          new TableCell({
            children: [new Paragraph(data.totalCompleted.toString())],
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            },
          }),
        ],
      }),
      new TableRow({
        cells: [
          new TableCell({
            children: [new Paragraph({ text: 'Pending', bold: true })],
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            },
          }),
          new TableCell({
            children: [new Paragraph((data.totalTasks - data.totalCompleted).toString())],
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            },
          }),
        ],
      }),
      new TableRow({
        cells: [
          new TableCell({
            children: [new Paragraph({ text: 'Completion Rate', bold: true })],
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            },
          }),
          new TableCell({
            children: [new Paragraph(`${data.completionRate}%`)],
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
              right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
            },
          }),
        ],
      }),
    ],
  });

  sections.push(summaryTable);
  sections.push(new Paragraph({ text: '', spacing: { after: 400 } }));
  sections.push(
    new Paragraph({
      text: 'Daily Breakdown',
      heading: HeadingLevel.HEADING_2,
      spacing: { after: 200 },
    })
  );

  dates.forEach((date) => {
    const dayTasks = data.tasks[date];
    const completed = dayTasks.filter((t) => t.completed).length;

    sections.push(
      new Paragraph({
        text: `${formatDate(date)}`,
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 200, after: 100 },
      })
    );

    dayTasks.forEach((task) => {
      const status = task.completed ? '✓' : '○';
      sections.push(
        new Paragraph({
          children: [
            new TextRun({ text: `${status} `, bold: true }),
            new TextRun({ text: task.title, strike: task.completed }),
          ],
          spacing: { after: 50 },
        })
      );
    });

    sections.push(
      new Paragraph({
        text: `Day Summary: ${completed} of ${dayTasks.length} tasks completed`,
        spacing: { after: 200 },
        italics: true,
      })
    );
  });

  const doc = new Document({
    sections: [
      {
        children: sections,
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};

export const generatePdfReport = (data: ReportData, fileName: string) => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4',
  });

  let yPos = 20;
  const pageHeight = pdf.internal.pageSize.getHeight();
  const margin = 15;

  pdf.setFontSize(16);
  pdf.text('WEEKLY TASK REPORT', margin, yPos);
  yPos += 12;

  pdf.setFontSize(10);
  pdf.text(`Period: ${formatDate(data.startDate)} - ${formatDate(data.endDate)}`, margin, yPos);
  yPos += 6;
  pdf.text(`Generated: ${new Date().toLocaleString()}`, margin, yPos);
  yPos += 12;

  pdf.setFontSize(12);
  pdf.text('Summary', margin, yPos);
  yPos += 8;

  pdf.setFontSize(10);
  pdf.text(`Total Tasks: ${data.totalTasks}`, margin, yPos);
  yPos += 5;
  pdf.text(`Completed: ${data.totalCompleted}`, margin, yPos);
  yPos += 5;
  pdf.text(`Pending: ${data.totalTasks - data.totalCompleted}`, margin, yPos);
  yPos += 5;
  pdf.text(`Completion Rate: ${data.completionRate}%`, margin, yPos);
  yPos += 12;

  const dates = Object.keys(data.tasks).sort();

  pdf.setFontSize(12);
  pdf.text('Daily Breakdown', margin, yPos);
  yPos += 8;

  dates.forEach((date) => {
    if (yPos > pageHeight - 30) {
      pdf.addPage();
      yPos = 20;
    }

    const dayTasks = data.tasks[date];
    const completed = dayTasks.filter((t) => t.completed).length;

    pdf.setFontSize(11);
    pdf.setTextColor(0, 51, 102);
    pdf.text(formatDate(date), margin, yPos);
    pdf.setTextColor(0, 0, 0);
    yPos += 6;

    pdf.setFontSize(9);
    dayTasks.forEach((task) => {
      if (yPos > pageHeight - 10) {
        pdf.addPage();
        yPos = 20;
      }
      const status = task.completed ? '[✓]' : '[ ]';
      pdf.text(`  ${status} ${task.title}`, margin + 2, yPos);
      yPos += 4;
    });

    yPos += 2;
    pdf.setFontSize(8);
    pdf.setTextColor(100, 100, 100);
    pdf.text(`Summary: ${completed}/${dayTasks.length} tasks completed`, margin + 2, yPos);
    pdf.setTextColor(0, 0, 0);
    yPos += 8;
  });

  pdf.save(fileName);
};
