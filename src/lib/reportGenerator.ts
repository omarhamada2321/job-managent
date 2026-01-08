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
      const status = task.completed ? 'DONE' : 'PENDING';
      sections.push(
        new Paragraph({
          children: [
            new TextRun({ text: `[${status}] `, bold: true, color: task.completed ? '52BE80' : 'FF9800' }),
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
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  link.style.display = 'none';
  document.body.appendChild(link);
  link.click();
  await new Promise(resolve => setTimeout(resolve, 500));
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};

export const generatePdfReport = (data: ReportData, fileName: string) => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4',
  });

  let yPos = 12;
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const margin = 10;
  const contentWidth = pageWidth - 2 * margin;

  const checkPageBreak = (needed: number) => {
    if (yPos + needed > pageHeight - 10) {
      pdf.addPage();
      yPos = 12;
    }
  };

  pdf.setFillColor(30, 70, 150);
  pdf.rect(margin - 1, yPos - 3, contentWidth + 2, 7, 'F');
  pdf.setTextColor(255, 255, 255);
  pdf.setFontSize(14);
  pdf.setFont('helvetica', 'bold');
  pdf.text('WEEKLY TASK REPORT', margin + 2, yPos);
  pdf.setTextColor(0, 0, 0);
  yPos += 10;

  pdf.setFontSize(8);
  pdf.setTextColor(100, 100, 100);
  const startFormatted = formatDate(data.startDate);
  const endFormatted = formatDate(data.endDate);
  pdf.text(`Period: ${startFormatted} to ${endFormatted}`, margin, yPos);
  yPos += 3;
  pdf.text(`Generated: ${new Date().toLocaleString()}`, margin, yPos);
  yPos += 7;

  pdf.setTextColor(25, 110, 200);
  pdf.setFontSize(10);
  pdf.setFont('helvetica', 'bold');
  pdf.text('SUMMARY', margin, yPos);
  pdf.setTextColor(0, 0, 0);
  yPos += 4;

  pdf.setFontSize(9);
  pdf.setFont('helvetica', 'normal');
  pdf.text(`Total Tasks: ${data.totalTasks}`, margin + 2, yPos);
  yPos += 4;
  pdf.text(`Completed: ${data.totalCompleted}`, margin + 2, yPos);
  yPos += 4;
  pdf.text(`Pending: ${data.totalTasks - data.totalCompleted}`, margin + 2, yPos);
  yPos += 4;
  pdf.text(`Completion Rate: ${data.completionRate}%`, margin + 2, yPos);
  yPos += 7;

  const dates = Object.keys(data.tasks).sort();

  pdf.setTextColor(25, 110, 200);
  pdf.setFontSize(10);
  pdf.setFont('helvetica', 'bold');
  pdf.text('DAILY BREAKDOWN', margin, yPos);
  pdf.setTextColor(0, 0, 0);
  yPos += 5;

  dates.forEach((date) => {
    checkPageBreak(8);

    const dayTasks = data.tasks[date];
    const completed = dayTasks.filter((t) => t.completed).length;
    const dayRate = dayTasks.length > 0 ? Math.round((completed / dayTasks.length) * 100) : 0;

    pdf.setFontSize(9);
    pdf.setTextColor(30, 100, 200);
    pdf.setFont('helvetica', 'bold');
    pdf.text(`${formatShortDate(date)} (${completed}/${dayTasks.length} - ${dayRate}%)`, margin + 1, yPos);
    pdf.setTextColor(0, 0, 0);
    pdf.setFont('helvetica', 'normal');
    yPos += 4;

    dayTasks.forEach((task) => {
      checkPageBreak(4);

      const status = task.completed ? 'DONE' : 'TODO';
      const statusColor = task.completed ? [82, 190, 128] : [255, 152, 0];

      pdf.setFontSize(8);
      pdf.setTextColor(statusColor[0], statusColor[1], statusColor[2]);
      pdf.setFont('helvetica', 'bold');
      pdf.text(`[${status}]`, margin + 2, yPos);

      pdf.setTextColor(0, 0, 0);
      pdf.setFont('helvetica', 'normal');
      const taskX = margin + 17;
      const maxWidth = pageWidth - taskX - margin - 2;
      const textLines = pdf.splitTextToSize(task.title, maxWidth);

      pdf.setFontSize(8);
      textLines.forEach((line, idx) => {
        pdf.text(line, taskX, yPos + (idx * 3.5));
      });

      yPos += Math.max(3.5, textLines.length * 3.5);
    });

    yPos += 2;
  });

  checkPageBreak(12);

  pdf.setTextColor(25, 110, 200);
  pdf.setFontSize(10);
  pdf.setFont('helvetica', 'bold');
  pdf.text('STATISTICS', margin, yPos);
  pdf.setTextColor(0, 0, 0);
  yPos += 5;

  pdf.setFontSize(9);
  pdf.setFont('helvetica', 'normal');

  const totalCompletePercentage = data.completionRate;
  const avgDailyCompletion = dates.length > 0
    ? Math.round(
        (Object.values(data.tasks).reduce((sum, dayTasks) => {
          const completed = dayTasks.filter((t) => t.completed).length;
          return sum + (dayTasks.length > 0 ? (completed / dayTasks.length) * 100 : 0);
        }, 0) /
          dates.length) as any
      )
    : 0;

  pdf.text(`Overall Completion: ${totalCompletePercentage}%`, margin + 2, yPos);
  yPos += 4;
  pdf.text(`Average Daily Completion: ${avgDailyCompletion}%`, margin + 2, yPos);
  yPos += 4;
  pdf.text(`Days Tracked: ${dates.length}`, margin + 2, yPos);
  yPos += 4;

  const pendingTasks = data.totalTasks - data.totalCompleted;
  let insight = '';
  if (data.completionRate === 100) {
    insight = 'Excellent! All tasks completed.';
  } else if (data.completionRate >= 80) {
    insight = `Good progress. ${pendingTasks} remaining.`;
  } else if (data.completionRate >= 50) {
    insight = `Fair progress. ${pendingTasks} need attention.`;
  } else {
    insight = `Review priorities. ${pendingTasks} pending.`;
  }

  pdf.setTextColor(50, 50, 50);
  pdf.setFont('helvetica', 'italic');
  pdf.text(insight, margin + 2, yPos);

  pdf.save(fileName);
};
