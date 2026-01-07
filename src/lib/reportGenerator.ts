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

  let yPos = 15;
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const margin = 12;
  const contentWidth = pageWidth - 2 * margin;

  const addHeaderBox = (title: string) => {
    pdf.setFillColor(30, 70, 150);
    pdf.rect(margin, yPos - 4, contentWidth, 8, 'F');
    pdf.setTextColor(255, 255, 255);
    pdf.setFontSize(12);
    pdf.setFont('helvetica', 'bold');
    pdf.text(title, margin + 3, yPos);
    pdf.setTextColor(0, 0, 0);
    yPos += 12;
  };

  const addSectionTitle = (title: string, color = [50, 120, 200]) => {
    pdf.setTextColor(color[0], color[1], color[2]);
    pdf.setFontSize(11);
    pdf.setFont('helvetica', 'bold');
    pdf.text(title, margin, yPos);
    pdf.setTextColor(0, 0, 0);
    pdf.setFont('helvetica', 'normal');
    yPos += 6;
  };

  const addDivider = () => {
    pdf.setDrawColor(200, 200, 200);
    pdf.line(margin, yPos, pageWidth - margin, yPos);
    yPos += 3;
  };

  const addText = (text: string, size = 9, bold = false, color = [0, 0, 0]) => {
    pdf.setFontSize(size);
    pdf.setTextColor(color[0], color[1], color[2]);
    pdf.setFont('helvetica', bold ? 'bold' : 'normal');
    pdf.text(text, margin + 2, yPos);
    yPos += 4.5;
  };

  const addMetricBox = (label: string, value: string | number, color: number[]) => {
    const boxWidth = (contentWidth - 4) / 2;
    const boxHeight = 14;

    pdf.setFillColor(color[0], color[1], color[2]);
    pdf.rect(margin + (boxWidth + 2) * (label === 'Total Tasks' || label === 'Completed' ? 0 : 1), yPos, boxWidth, boxHeight, 'F');

    pdf.setTextColor(255, 255, 255);
    pdf.setFontSize(9);
    pdf.setFont('helvetica', 'normal');
    pdf.text(label, margin + 3 + (boxWidth + 2) * (label === 'Total Tasks' || label === 'Completed' ? 0 : 1), yPos + 4);

    pdf.setFontSize(14);
    pdf.setFont('helvetica', 'bold');
    pdf.text(String(value), margin + 3 + (boxWidth + 2) * (label === 'Total Tasks' || label === 'Completed' ? 0 : 1), yPos + 11);

    if (label === 'Completed' || label === 'Pending') {
      yPos += boxHeight + 2;
    }
  };

  const checkPageBreak = (space: number) => {
    if (yPos + space > pageHeight - 15) {
      pdf.addPage();
      yPos = 15;
    }
  };

  addHeaderBox('WEEKLY TASK REPORT');

  pdf.setFontSize(9);
  pdf.setTextColor(100, 100, 100);
  const startFormatted = formatDate(data.startDate);
  const endFormatted = formatDate(data.endDate);
  pdf.text(`Period: ${startFormatted} to ${endFormatted}`, margin, yPos);
  yPos += 4;
  pdf.text(`Generated: ${new Date().toLocaleString()}`, margin, yPos);
  yPos += 8;

  addSectionTitle('KEY METRICS', [25, 110, 200]);
  addDivider();

  const summaryStartY = yPos;
  addMetricBox('Total Tasks', data.totalTasks, [66, 153, 225]);
  addMetricBox('Completed', data.totalCompleted, [82, 190, 128]);
  yPos = summaryStartY + 16;
  addMetricBox('Pending', data.totalTasks - data.totalCompleted, [255, 152, 0]);
  addMetricBox('Completion Rate', `${data.completionRate}%`, [156, 39, 176]);

  yPos += 4;
  checkPageBreak(10);

  addDivider();
  addSectionTitle('PERFORMANCE INSIGHT', [25, 110, 200]);

  const pendingTasks = data.totalTasks - data.totalCompleted;
  let insight = '';
  if (data.completionRate === 100) {
    insight = 'Excellent! All tasks completed on schedule.';
  } else if (data.completionRate >= 80) {
    insight = `Good progress. ${pendingTasks} task${pendingTasks !== 1 ? 's' : ''} remaining.`;
  } else if (data.completionRate >= 50) {
    insight = `Fair progress. ${pendingTasks} task${pendingTasks !== 1 ? 's' : ''} need attention.`;
  } else {
    insight = `Review priority. ${pendingTasks} task${pendingTasks !== 1 ? 's' : ''} pending.`;
  }
  addText(insight, 9, false, [50, 50, 50]);
  yPos += 2;

  checkPageBreak(15);

  addDivider();
  addSectionTitle('DAILY BREAKDOWN', [25, 110, 200]);
  yPos += 2;

  const dates = Object.keys(data.tasks).sort();

  dates.forEach((date) => {
    checkPageBreak(12);

    const dayTasks = data.tasks[date];
    const completed = dayTasks.filter((t) => t.completed).length;
    const dayRate = dayTasks.length > 0 ? Math.round((completed / dayTasks.length) * 100) : 0;

    pdf.setFontSize(10);
    pdf.setTextColor(30, 100, 200);
    pdf.setFont('helvetica', 'bold');
    pdf.text(`${formatShortDate(date)}  -  ${completed}/${dayTasks.length} completed (${dayRate}%)`, margin, yPos);
    pdf.setTextColor(0, 0, 0);
    pdf.setFont('helvetica', 'normal');
    yPos += 5;

    pdf.setDrawColor(220, 220, 220);
    pdf.line(margin + 1, yPos, pageWidth - margin - 1, yPos);
    yPos += 3;

    dayTasks.forEach((task) => {
      checkPageBreak(5);

      const status = task.completed ? 'DONE' : 'PENDING';
      const statusColor = task.completed ? [82, 190, 128] : [255, 152, 0];

      pdf.setFontSize(8);
      pdf.setTextColor(statusColor[0], statusColor[1], statusColor[2]);
      pdf.setFont('helvetica', 'bold');
      pdf.text(`[${status}]`, margin + 2, yPos);

      pdf.setTextColor(0, 0, 0);
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(9);
      const taskX = margin + 18;
      pdf.text(task.title, taskX, yPos);

      yPos += 4.5;
    });

    yPos += 2;
  });

  checkPageBreak(20);

  addDivider();
  addSectionTitle('CONCLUSION', [25, 110, 200]);

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

  addText(`Overall Completion: ${totalCompletePercentage}%`, 9, true, [25, 110, 200]);
  addText(`Average Daily Completion: ${avgDailyCompletion}%`, 9);
  addText(`Report Duration: ${dates.length} day${dates.length !== 1 ? 's' : ''} tracked`, 9);

  yPos += 3;
  pdf.setFontSize(8);
  pdf.setTextColor(120, 120, 120);
  pdf.text('---', margin, yPos);
  yPos += 4;
  pdf.text('This report provides a comprehensive overview of task completion metrics.', margin, yPos);
  yPos += 4;
  pdf.text('Review daily breakdowns to identify patterns and optimize future planning.', margin, yPos);

  pdf.save(fileName);
};
