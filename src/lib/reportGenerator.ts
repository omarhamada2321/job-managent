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
  try {
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
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 100);
  } catch (error) {
    console.error('Error generating Word report:', error);
    throw error;
  }
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

  const checkPageBreak = (space: number) => {
    if (yPos + space > pageHeight - 12) {
      pdf.addPage();
      yPos = 12;
    }
  };

  const addText = (text: string, size = 9, bold = false, color = [0, 0, 0]) => {
    pdf.setFontSize(size);
    pdf.setTextColor(color[0], color[1], color[2]);
    pdf.setFont('helvetica', bold ? 'bold' : 'normal');
    pdf.text(text, margin, yPos);
    yPos += size === 8 ? 3.5 : 4.5;
  };

  const addDivider = () => {
    pdf.setDrawColor(180, 180, 180);
    pdf.line(margin, yPos, pageWidth - margin, yPos);
    yPos += 2;
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
  yPos += 6;

  pdf.setTextColor(25, 110, 200);
  pdf.setFontSize(10);
  pdf.setFont('helvetica', 'bold');
  pdf.text('KEY METRICS', margin, yPos);
  pdf.setTextColor(0, 0, 0);
  yPos += 4;

  addDivider();

  const boxWidth = (contentWidth - 2) / 2;
  const metrics = [
    { label: 'Total Tasks', value: data.totalTasks, color: [66, 153, 225] },
    { label: 'Completed', value: data.totalCompleted, color: [82, 190, 128] },
    { label: 'Pending', value: data.totalTasks - data.totalCompleted, color: [255, 152, 0] },
    { label: 'Completion Rate', value: `${data.completionRate}%`, color: [156, 39, 176] },
  ];

  let col = 0;
  let rowYPos = yPos;
  metrics.forEach((metric) => {
    const xPos = margin + col * (boxWidth + 1);
    pdf.setFillColor(metric.color[0], metric.color[1], metric.color[2]);
    pdf.rect(xPos, rowYPos, boxWidth, 10, 'F');
    pdf.setTextColor(255, 255, 255);
    pdf.setFontSize(7);
    pdf.setFont('helvetica', 'normal');
    pdf.text(metric.label, xPos + 2, rowYPos + 3);
    pdf.setFontSize(11);
    pdf.setFont('helvetica', 'bold');
    pdf.text(String(metric.value), xPos + 2, rowYPos + 8);
    col++;
    if (col === 2) {
      col = 0;
      rowYPos += 11;
    }
  });

  yPos = rowYPos + 2;
  checkPageBreak(10);

  pdf.setTextColor(25, 110, 200);
  pdf.setFontSize(10);
  pdf.setFont('helvetica', 'bold');
  pdf.text('PERFORMANCE INSIGHT', margin, yPos);
  pdf.setTextColor(0, 0, 0);
  yPos += 4;
  addDivider();

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
  addText(insight, 8, false, [50, 50, 50]);
  yPos += 1;

  checkPageBreak(12);

  pdf.setTextColor(25, 110, 200);
  pdf.setFontSize(10);
  pdf.setFont('helvetica', 'bold');
  pdf.text('DAILY BREAKDOWN', margin, yPos);
  pdf.setTextColor(0, 0, 0);
  yPos += 4;
  addDivider();

  const dates = Object.keys(data.tasks).sort();

  dates.forEach((date) => {
    checkPageBreak(8);

    const dayTasks = data.tasks[date];
    const completed = dayTasks.filter((t) => t.completed).length;
    const dayRate = dayTasks.length > 0 ? Math.round((completed / dayTasks.length) * 100) : 0;

    pdf.setFontSize(9);
    pdf.setTextColor(30, 100, 200);
    pdf.setFont('helvetica', 'bold');
    pdf.text(`${formatShortDate(date)} | ${completed}/${dayTasks.length} (${dayRate}%)`, margin, yPos);
    pdf.setTextColor(0, 0, 0);
    pdf.setFont('helvetica', 'normal');
    yPos += 3.5;

    dayTasks.forEach((task) => {
      checkPageBreak(3);

      const status = task.completed ? 'DONE' : 'PENDING';
      const statusColor = task.completed ? [82, 190, 128] : [255, 152, 0];

      pdf.setFontSize(7.5);
      pdf.setTextColor(statusColor[0], statusColor[1], statusColor[2]);
      pdf.setFont('helvetica', 'bold');
      pdf.text(`[${status}]`, margin + 1, yPos);

      pdf.setTextColor(0, 0, 0);
      pdf.setFont('helvetica', 'normal');
      const taskX = margin + 15;
      const maxWidth = pageWidth - taskX - margin - 2;
      const textLines = pdf.splitTextToSize(task.title, maxWidth);
      pdf.text(textLines, taskX, yPos);

      yPos += textLines.length * 3.2;
    });

    yPos += 2;
  });

  checkPageBreak(15);

  pdf.setTextColor(25, 110, 200);
  pdf.setFontSize(10);
  pdf.setFont('helvetica', 'bold');
  pdf.text('CONCLUSION', margin, yPos);
  pdf.setTextColor(0, 0, 0);
  yPos += 4;
  addDivider();

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

  addText(`Overall Completion: ${totalCompletePercentage}%`, 8, true, [25, 110, 200]);
  addText(`Average Daily Completion: ${avgDailyCompletion}%`, 8);
  addText(`Report Duration: ${dates.length} day${dates.length !== 1 ? 's' : ''} tracked`, 8);

  pdf.setFontSize(7);
  pdf.setTextColor(150, 150, 150);
  yPos += 2;
  pdf.text('This report provides a comprehensive overview of task completion metrics.', margin, yPos);
  yPos += 3;
  pdf.text('Review daily breakdowns to identify patterns and optimize future planning.', margin, yPos);

  pdf.save(fileName);
};
